"""
Document generators for NY Surrogate's Court Probate Assistant
Generates filled Word docs and PDFs from case data
"""

import io
import os
import re
from datetime import datetime
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
import fitz

TEMPLATES_DIR       = os.path.join(os.path.dirname(__file__), "templates")
ADMIN_TEMPLATES_DIR = os.path.join(TEMPLATES_DIR, "Admin")
PROBATE_TEMPLATES_DIR = os.path.join(TEMPLATES_DIR, "Probate")
WORD_TEMPLATES_DIR  = os.path.join(TEMPLATES_DIR, "Not Using Word Docs")
PDFS_DIR            = os.path.dirname(__file__)

COUNTY_INFO = {
    "Bronx": {
        "address": "851 Grand Concourse, 3rd Floor",
        "city_state_zip": "Bronx, NY 10451",
        "dept_probate": "Probate Department",
        "dept_admin": "Administration Department",
    },
    "Kings": {
        "address": "2 Johnson Street",
        "city_state_zip": "Brooklyn, NY 11201",
        "dept_probate": "Probate Department",
        "dept_admin": "Administration Department",
    },
    "Nassau": {
        "address": "262 Old Country Road",
        "city_state_zip": "Mineola, NY 11501",
        "dept_probate": "Probate Department",
        "dept_admin": "Administration Department",
    },
    "New York": {
        "address": "31 Chambers Street",
        "city_state_zip": "New York, NY 10007",
        "dept_probate": "Probate Department",
        "dept_admin": "Administration Department",
    },
    "Queens": {
        "address": "88-11 Sutphin Blvd",
        "city_state_zip": "Jamaica, NY 11435",
        "dept_probate": "Probate Department",
        "dept_admin": "Administration Department",
    },
    "Richmond": {
        "address": "18 Richmond Terrace",
        "city_state_zip": "Staten Island, NY 10301",
        "dept_probate": "Probate Department",
        "dept_admin": "Administration Department",
    },
    "Suffolk": {
        "address": "320 Center Drive",
        "city_state_zip": "Riverhead, NY 11901",
        "dept_probate": "Probate Department",
        "dept_admin": "Administration Department",
    },
}

SIGNERS = {
    "Jessica Wilson": "Jessica Wilson, Esq.",
    "Robyn Foresta": "Robyn Foresta, Legal Assistant",
}

# Relationship keywords that trigger Rule 207.16(c) on their own
# (grandparents, aunts/uncles, first cousins, first cousins once removed)
_DISTANT_REL_KEYWORDS = [
    "grandparent", "grandfather", "grandmother",
    "aunt", "uncle",
    "cousin",
]


def needs_family_tree_affidavit(data):
    """Return True if Rule 207.16(c) requires an Affidavit of Family Tree.

    Triggers when:
      - 0 or 1 distributee survives, OR
      - any distributee's relationship is grandparents, aunts/uncles,
        first cousins, or first cousins once removed.
    """
    dists = [d for d in data.get("distributees", []) if d.get("name")]
    if len(dists) <= 1:
        return True
    for d in dists:
        rel = (d.get("relationship") or "").lower()
        if any(k in rel for k in _DISTANT_REL_KEYWORDS):
            return True
    return False


def needs_family_tree_diagram(data):
    """Return True if the FT-1 diagram is also required (Rule 207.16(c)).

    The diagram is NOT required when the sole distributee is the spouse
    or only child of the decedent.
    """
    if not needs_family_tree_affidavit(data):
        return False
    dists = [d for d in data.get("distributees", []) if d.get("name")]
    if len(dists) == 1:
        rel = (dists[0].get("relationship") or "").lower()
        if any(k in rel for k in ["spouse", "child", "son", "daughter"]):
            return False
    return True


def family_tree_trigger_reason(data):
    """Return a short human-readable string explaining why 207.16(c) fired
    (used in the case summary).  Returns empty string if not triggered."""
    dists = [d for d in data.get("distributees", []) if d.get("name")]
    if len(dists) == 0:
        return "no distributees"
    if len(dists) == 1:
        return f"only one distributee ({dists[0].get('name', '')})"
    for d in dists:
        rel = (d.get("relationship") or "").lower()
        if any(k in rel for k in _DISTANT_REL_KEYWORDS):
            return f"distributee relationship: {d.get('relationship', '')}"
    return ""


def today():
    return datetime.now().strftime("%B %d, %Y")


def format_date_long(date_str):
    """Convert MM/DD/YYYY to 'Month DD, YYYY' (e.g. '03/15/1945' → 'March 15, 1945')."""
    try:
        dt = datetime.strptime(date_str, "%m/%d/%Y")
        return dt.strftime("%B %d, %Y")
    except Exception:
        return date_str


def nonzero(v):
    """Return v only if it's a non-empty, non-zero value."""
    s = str(v or "").strip()
    return s if s and s not in ("0", "0.0", "0.00") else ""


def replace_in_doc(doc, replacements):
    """Replace placeholder text throughout a Word document.
    Handles placeholders split across multiple runs by normalizing para text first."""
    def replace_in_para(para):
        for key, value in replacements.items():
            if key not in para.text:
                continue
            # If key is in a single run, do fast replacement
            for run in para.runs:
                if key in run.text:
                    run.text = run.text.replace(key, value or "")
                    break
            else:
                # Key is split across runs — consolidate into first run
                full_text = para.text.replace(key, value or "")
                if para.runs:
                    para.runs[0].text = full_text
                    for run in para.runs[1:]:
                        run.text = ""
                else:
                    para.add_run(full_text)

    for para in doc.paragraphs:
        replace_in_para(para)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    replace_in_para(para)


def replace_para(para, old_text, new_text):
    """Replace text within a paragraph's runs, preserving formatting."""
    full = para.text
    if old_text not in full and old_text != full:
        return
    for run in para.runs:
        if old_text in run.text:
            run.text = run.text.replace(old_text, new_text)
            return
    new_full = full.replace(old_text, new_text)
    for run in para.runs:
        run.text = ""
    if para.runs:
        para.runs[0].text = new_full
    else:
        para.add_run(new_full)


def make_docx_bytes(doc):
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


# ─── COVER LETTER ─────────────────────────────────────────────────────────────

def generate_cover_letter(data):
    county = data.get("county", "")
    proceeding = data.get("proceedingType", "Probate")
    signer_key = data.get("signer", "Jessica Wilson")
    signer = SIGNERS.get(signer_key, signer_key)
    decedent = decedent_full(data)
    efile_date = data.get("efileDate", today())
    enclosures = data.get("enclosures", [])

    county_info = COUNTY_INFO.get(county, {})
    address = county_info.get("address", "")
    city_state_zip = county_info.get("city_state_zip", "")
    dept = county_info.get("dept_probate" if proceeding == "Probate" else "dept_admin", "")

    doc = Document()

    # Set default style to single-spaced
    style = doc.styles['Normal']
    style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    style.paragraph_format.space_after = Pt(0)
    style.paragraph_format.space_before = Pt(0)

    def _para(text="", space_after=0):
        p = doc.add_paragraph(text)
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.space_before = Pt(0)
        return p

    # Date
    _para(today(), space_after=12)

    # Addressee
    _para(f"Surrogate's Court, {county} County")
    _para(f"Attn: {dept}")
    _para(address)
    _para(city_state_zip, space_after=12)

    # RE line
    _para(f"RE: Estate of {decedent}", space_after=12)

    _para("Greetings,", space_after=6)

    proc_word = proceeding.lower()
    _para(
        f"Our office efiled the above referenced petition for {proc_word} on {efile_date}. "
        f"Please find enclosed the following original documents required by the Court:",
        space_after=6
    )

    # Enclosures as bullet list
    for enc in enclosures:
        p = doc.add_paragraph(style="List Bullet")
        p.text = enc
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        p.paragraph_format.space_after = Pt(0)

    _para("", space_after=6)
    _para("Please do not hesitate to call our office if you have concerns and questions.")
    _para("", space_after=6)
    _para("Sincerely,")
    _para("")
    _para("")
    _para(signer)
    _para("Enc.")

    return make_docx_bytes(doc)


# ─── 805 AFFIDAVIT ────────────────────────────────────────────────────────────

def generate_805(data):
    """Build the 805 Affidavit of Assets & Liabilities from scratch for
    consistent formatting (Times New Roman 12pt, 1-inch margins)."""
    doc = Document()

    # ── Page margins: 1" all sides ────────────────────────────────────────────
    for section in doc.sections:
        section.top_margin    = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin   = Inches(1)
        section.right_margin  = Inches(1)

    county    = data.get("county", "")
    decedent  = decedent_full(data)
    petitioner = petitioner_full(data)
    file_no   = data.get("fileNo", "")
    aka       = (data.get("decedentAKA") or "").strip()
    year      = datetime.now().strftime("%Y")

    # ── Helpers ───────────────────────────────────────────────────────────────
    FONT = "Times New Roman"
    SIZE = Pt(12)

    def _run(para, text, bold=False, italic=False):
        r = para.add_run(text)
        r.font.name  = FONT
        r.font.size  = SIZE
        r.bold       = bold
        r.italic     = italic
        return r

    def line(text="", bold=False, italic=False, center=False,
             space_before=0, space_after=0, left_indent=None):
        p = doc.add_paragraph()
        p.paragraph_format.alignment    = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after  = Pt(space_after)
        if left_indent is not None:
            p.paragraph_format.left_indent = Inches(left_indent)
        if text:
            _run(p, text, bold=bold, italic=italic)
        return p

    def blank(n=1):
        for _ in range(n):
            line()

    # ── Caption ──────────────────────────────────────────────────────────────
    proceeding = data.get("proceedingType", "Administration")
    letters_type = data.get("lettersType", "Letters of Administration")

    line("SURROGATE\u2019S COURT OF THE STATE OF NEW YORK", bold=True)
    line(f"COUNTY OF {county.upper()}", bold=True, space_after=2)

    divider = "\u2500" * 43 + "x"
    line(divider, space_after=0)

    # Two-column caption using a borderless table
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    def _no_border(cell):
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        tcBorders = OxmlElement("w:tcBorders")
        for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
            el = OxmlElement(f"w:{side}")
            el.set(qn("w:val"), "none")
            el.set(qn("w:sz"), "0")
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), "auto")
            tcBorders.append(el)
        tcPr.append(tcBorders)

    # Build caption rows — left column has matter text, right has doc title
    aka_line = f"    a/k/a {aka}," if aka else ""
    left_lines = [
        "In the Matter of the Application for",
        "",
        f"{letters_type} of the Estate of",
        "",
        f"    {decedent.upper()},",
    ]
    if aka_line:
        left_lines.append(aka_line)
    left_lines += [
        "",
        "                Deceased.",
    ]

    right_lines = [
        ("AFFIDAVIT OF ASSETS", True),
        ("& LIABILITIES", True),
        ("(SCPA 805)", True),
        ("", False),
        (f"File No. {file_no}" if file_no else "", False),
    ]
    # Pad right column to match left
    while len(right_lines) < len(left_lines):
        right_lines.append(("", False))

    cap_tbl = doc.add_table(rows=len(left_lines), cols=2)
    cap_tbl.style = "Table Grid"

    # Set table width and column widths
    tbl_el = cap_tbl._tbl
    tblPr = tbl_el.tblPr if tbl_el.tblPr is not None else OxmlElement("w:tblPr")
    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:w"), "0")
    tblW.set(qn("w:type"), "auto")
    tblPr.append(tblW)
    # Remove table borders at the table level too
    tblBorders = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"), "none")
        el.set(qn("w:sz"), "0")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "auto")
        tblBorders.append(el)
    tblPr.append(tblBorders)

    for row_i, row in enumerate(cap_tbl.rows):
        for col_i, cell in enumerate(row.cells):
            _no_border(cell)
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            if col_i == 0:
                _run(p, left_lines[row_i])
            else:
                # Add vertical bar separator before right-column text
                txt, bld = right_lines[row_i]
                cell_text = f"\u2502  {txt}" if txt else "\u2502"
                _run(p, cell_text, bold=bld)

    line(divider, space_after=4)

    # ── Venue block ───────────────────────────────────────────────────────────
    line("STATE OF NEW YORK\t\t\t\t)")
    line("\t\t\t\t\t\t) ss:")
    line(f"COUNTY OF {county.upper()}\t\t\t\t)")
    blank()

    # ── Oath paragraph ────────────────────────────────────────────────────────
    line(
        "I, the undersigned being duly sworn, depose and say:  I have personal knowledge "
        "as to the assets, debts and/or liabilities of the estate of the decedent. "
        "The assets of the estate, including real and/or personal property held solely "
        "by the decedent consist of:",
        space_after=4
    )

    # Assets — skip empty and zero values
    pp = nonzero(data.get("personalPropertyValue"))
    ir = nonzero(data.get("improvedRealProperty"))
    ur = nonzero(data.get("unimprovedRealProperty"))
    rd = (data.get("realPropertyDescription") or "").strip()
    gr = nonzero(data.get("grossRents18mo"))

    asset_lines = []
    if pp: asset_lines.append(f"Personal Property:  ${pp}")
    if ir: asset_lines.append(f"Improved Real Property (NY):  ${ir}")
    if ur: asset_lines.append(f"Unimproved Real Property (NY):  ${ur}")
    if rd: asset_lines.append(f"Description:  {rd}")
    if gr: asset_lines.append(f"Gross Rents (18 months):  ${gr}")
    if not asset_lines:
        asset_lines = ["NONE"]

    for asset in asset_lines:
        line(asset, left_indent=0.5)

    blank()

    # ── Liabilities ───────────────────────────────────────────────────────────
    line(
        "All the liabilities of the decedent known to me are as follows "
        "(Indicate AMOUNT DUE or answer \u201cNONE\u201d):",
        space_before=4, space_after=4
    )

    mort = (data.get("mortgageAmount") or "").strip()
    fp   = (data.get("funeralPaid") or "").strip()
    fo   = (data.get("funeralOutstanding") or "").strip()
    misc = (data.get("miscDebts") or "").strip()

    line(f"Amount of outstanding mortgages:  {mort or 'NONE'}", left_indent=0.5)
    line(
        f"Amount of funeral expenses paid (attach copy of paid funeral bill):  {fp or 'NONE'}",
        left_indent=0.5
    )
    line(f"Amount of funeral expenses still outstanding:  {fo or 'NONE'}", left_indent=0.5)
    blank()
    line(
        "Itemize and specify amount of any miscellaneous expenses payable "
        "(i.e. credit card, utility bills, insurance premiums, etc.  "
        "Use attachments if more space is required.)",
        italic=True, space_after=2
    )
    line("NOTE: ANY UNSECURED DEBT MAY BE BONDED", bold=True, space_after=4)

    if misc:
        for ln in misc.splitlines():
            if ln.strip():
                line(ln.strip(), left_indent=0.5)
    else:
        line("NONE", left_indent=0.5)

    blank()

    # ── WHEREFORE clause ──────────────────────────────────────────────────────
    line(
        f"WHEREFORE, your deponent prays, that the filing of a bond by {petitioner} "
        f"as administrator and sole distributee be dispensed with.",
        space_before=6, space_after=18
    )

    # ── Signature block ───────────────────────────────────────────────────────
    line("__________________________________", space_after=2)
    line(petitioner, space_after=14)

    line(f"Sworn to before me this _________")
    line(f"day of __________________, {year}")
    blank()
    line("__________________________________", space_after=2)
    line("Notary Public")

    return make_docx_bytes(doc)


# ─── AFFIDAVIT OF HEIRSHIP ────────────────────────────────────────────────────

def generate_heirship(data):
    proceeding = data.get("proceedingType", "Administration")
    if proceeding == "Probate":
        template = "Affidavit_of_Heirship_Full_Probate.docx"
        letters_phrase = "Letters Testamentary"
    else:
        template = "Affidavit_of_Heirship_Full_Admin.docx"
        letters_phrase = "Letters of Administration"
    doc = Document(os.path.join(WORD_TEMPLATES_DIR, template))
    decedent = decedent_full(data)
    county = data.get("county", "")
    petitioner = petitioner_full(data)
    deponent = data.get("deponentName", petitioner)
    deponent_address = data.get("deponentAddress", data.get("petitionerStreet", ""))
    deponent_rel = data.get("deponentRelationship", "")
    years_known = data.get("yearsKnown", "")
    dob = data.get("decedentDOB", "")
    dod = data.get("decedentDOD", "")
    marital_status    = (data.get("maritalStatus") or "").strip()      # never_married / married / divorced / widowed
    spouse_name       = (data.get("spouseName") or "").strip()
    divorce_year      = (data.get("divorceYear") or "").strip()
    prior_spouse_death = (data.get("priorSpouseDeathDate") or "").strip()
    children_note = data.get("childrenNote", "").strip()
    mother_name = data.get("motherName", "")
    mother_dod = data.get("motherDOD", "")
    father_name = data.get("fatherName", "")
    father_dod = data.get("fatherDOD", "")
    sole_distributee = data.get("soleDistributee", petitioner)

    was_married = marital_status in ("married", "divorced", "widowed")
    has_children = bool(children_note and "never had" not in children_note.lower())

    # Build the marriage sentence for para 21
    if marital_status == "married":
        marriage_sentence = (
            f"Decedent was married to {spouse_name} at the time of death "
            f"and was never divorced."
        )
    elif marital_status == "divorced":
        yr = f"in {divorce_year}" if divorce_year else "prior to death"
        marriage_sentence = (
            f"Decedent was married to {spouse_name}, which said marriage ended in "
            f"divorce {yr}. The decedent never remarried after said divorce."
        )
    elif marital_status == "widowed":
        when = f"on {prior_spouse_death}" if prior_spouse_death else "prior to the decedent's death"
        marriage_sentence = (
            f"Decedent was married to {spouse_name}, who predeceased the decedent "
            f"{when}. The decedent never remarried after the death of said spouse."
        )
    else:
        marriage_sentence = None  # never married — use para 23 instead

    paras_to_delete = []

    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()

        if "which said marriage ended in divorce" in text:
            if marriage_sentence:
                replace_para(para, para.text, marriage_sentence)
            else:
                paras_to_delete.append(i)

        elif "The decedent was/never married" in text:
            if was_married:
                paras_to_delete.append(i)
            else:
                replace_para(para, para.text,
                             "The decedent was never married.")

        elif "never had any children" in text or "did have children" in text:
            if has_children:
                replace_para(para, "The decedent never had any children, adopted, out of wedlock nor marital. Or did have children.", children_note)
            else:
                replace_para(para, "The decedent never had any children, adopted, out of wedlock nor marital. Or did have children.",
                    "The decedent never had any children, adopted, out of wedlock nor marital.")

        elif "The marriage of" in text and "bore no children" in text:
            if was_married:
                replace_para(para, "___________ and ____________",
                             f"{decedent} and {spouse_name}")
            else:
                paras_to_delete.append(i)

        elif "There were no children of the decedent" in text:
            if was_married or has_children:
                paras_to_delete.append(i)

    for i in sorted(paras_to_delete, reverse=True):
        p = doc.paragraphs[i]._element
        p.getparent().remove(p)

    # Format dates as written-out (e.g. "March 15, 1945")
    dob_long = format_date_long(dob)
    dod_long = format_date_long(dod)

    replace_in_doc(doc, {
        "COUNTY OF _____________": f"COUNTY OF {county.upper()}",
        "___________________\t\t\t\t\tAFFIDAVIT OF HEIRSHIP": f"{decedent}\t\t\t\t\tAFFIDAVIT OF HEIRSHIP",
        "A/K/A ___________________": f"A/K/A {data.get('decedentAKA', '')}",
        "COUNTY OF \t\t\t)": f"COUNTY OF {county.upper()}\t\t\t)",
        "\tI, ______________, being duly sworn, deposes and says:": f"\tI, {deponent}, being duly sworn, deposes and says:",
        "I reside at _________________________.  I am over the age of eighteen (18) years and I am fully familiar with the facts and circumstances herein, the decedent\u2019s family tree, as I am the ______________of the Decedent and have known the Decedent for over _____ years.":
            f"I reside at {deponent_address}.  I am over the age of eighteen (18) years and I am fully familiar with the facts and circumstances herein, the decedent\u2019s family tree, as I am the {deponent_rel} of the Decedent and have known the Decedent for over {years_known} years.",
        "The Decedent was born on ___________ and died on __________________.": f"The Decedent was born on {dob_long} and died on {dod_long}.",
        "Mother: ": f"Mother: {mother_name}",
        "Father: ": f"Father: {father_name}",
        f"Therefore, ______________ is the sole distributee of the Estate of ______________":
            f"Therefore, {sole_distributee} is the sole distributee of the Estate of {decedent}",
        f"This affidavit is made with my personal knowledge knowing the ______________ County Surrogate\u2019s Court will rely thereon in issuing Letters Testamentary to _________________, the petitioner." if proceeding == "Probate" else
        f"This affidavit is made with my personal knowledge knowing the ______________ County Surrogate\u2019s Court will rely thereon in issuing Letters of Administration to _________________, the petitioner.":
            f"This affidavit is made with my personal knowledge knowing the {county} County Surrogate\u2019s Court will rely thereon in issuing {letters_phrase} to {petitioner}, the petitioner.",
    })

    mother_dod_filled = False
    for para in doc.paragraphs:
        t = para.text.strip()
        if t.startswith("Date of Death:"):
            if not mother_dod_filled:
                replace_para(para, "Date of Death:", f"Date of Death: {mother_dod}")
                mother_dod_filled = True
            else:
                replace_para(para, "Date of Death:", f"Date of Death: {father_dod}")

    return make_docx_bytes(doc)


# ─── WAIVER COVER LETTER ──────────────────────────────────────────────────────

def generate_waiver_cover(data, distributee):
    doc = Document(os.path.join(WORD_TEMPLATES_DIR, "Waiver_cover_letter.docx"))
    decedent = decedent_full(data)
    petitioner = petitioner_full(data)
    dist_name = distributee.get("name", "")
    dist_rel = distributee.get("relationship", "")
    dist_addr = distributee.get("address", "")

    replace_in_doc(doc, {
        "September 27, 2022": today(),
        "(Distributee)": dist_name,
        "(Distributee Address)": dist_addr,
        "(Deceased)": decedent,
        "(Petitioner)": petitioner,
    })

    return make_docx_bytes(doc)


# ─── ATTORNEY CERTIFICATION ───────────────────────────────────────────────────

def generate_attorney_cert(data):
    doc = Document(os.path.join(WORD_TEMPLATES_DIR, "newcertform_6_59_19_PM.docx"))
    replace_in_doc(doc, {
        "Dated:": f"Dated: {today()}",
    })
    return make_docx_bytes(doc)


# ─── PDF FILLING (pymupdf/fitz) ──────────────────────────────────────────────

def fill_pdf(template_path, fields):
    """Universal PDF form filler using pymupdf/fitz.

    Handles text, checkboxes (True/False), radio buttons (export value like '/0'),
    and combo/dropdown fields. Calls widget.update() on every filled field to bake
    in appearance streams so fields render in any viewer including macOS Preview.
    """
    doc = fitz.open(template_path)
    for page in doc:
        for widget in page.widgets():
            name = widget.field_name
            if name not in fields:
                continue
            value = fields[name]
            if widget.field_type == fitz.PDF_WIDGET_TYPE_CHECKBOX:
                widget.field_value = bool(value)
            elif widget.field_type == fitz.PDF_WIDGET_TYPE_RADIOBUTTON:
                widget.field_value = str(value).lstrip("/")
            else:
                s = str(value) if value is not None else ""
                widget.field_value = s
                if s == "X":
                    widget.text_fontsize = 10
            widget.update()
    buf = io.BytesIO()
    doc.save(buf)
    doc.close()
    buf.seek(0)
    return buf.read()


def _extract_pages(pdf_bytes, page_indices):
    """Extract specific pages from PDF bytes, preserving form widgets."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    doc.select(page_indices)
    buf = io.BytesIO()
    doc.save(buf, garbage=4, deflate=True)
    doc.close()
    buf.seek(0)
    return buf.read()


def extract_pdf_pages(template_path, fields, page_indices):
    """Fill a PDF template and extract specific pages."""
    filled = fill_pdf(template_path, fields)
    return _extract_pages(filled, page_indices)


# ─── PROBATE PDF (P-1 + OATH + WITNESS) ──────────────────────────────────────


def _build_probate_fields(data):
    """Build field name→value dict for Probate Petition + Oath.pdf."""
    county   = data.get("county", "")
    dec      = decedent_full(data)
    pet      = petitioner_full(data)
    lt       = data.get("lettersType", "")
    letters_to = data.get("lettersTo", "") or pet
    witnesses = ", ".join(filter(None, [data.get("witness1", ""), data.get("witness2", "")]))
    pet_addr  = ", ".join(filter(None, [
        data.get("petitionerStreet", ""), data.get("petitionerCity", ""),
        data.get("petitionerState", ""), data.get("petitionerZip", ""),
    ]))

    # Surviving relatives → Dropdown 5a–5g (EPTL 4-1.1 order, 7 classes)
    # Logic: "No" for prior classes, number/Yes for first surviving class, "X" for all after
    surv_keys = [
        "survivingSpouse", "survivingChildren", "survivingParents",
        "survivingSiblings", "survivingGrandparents", "survivingAuntsUncles",
        "survivingFirstCousinsOnceRemoved",
    ]
    # Find first surviving class
    first_surviving = None
    for idx, key in enumerate(surv_keys):
        raw = data.get(key)
        if raw and str(raw).strip().lower() not in ("false", "0", "no", ""):
            first_surviving = idx
            break
    dropdown_vals = []
    for idx, key in enumerate(surv_keys):
        raw = data.get(key)
        if first_surviving is None:
            dropdown_vals.append("No")
        elif idx < first_surviving:
            dropdown_vals.append("No")
        elif idx == first_surviving:
            s = str(raw).strip()
            dropdown_vals.append(s if s.lower() not in ("true", "yes") else "Yes")
        else:
            dropdown_vals.append("X")

    fields = {
        # ── Petition (pages 1-4) ────────────────────────────────────────────────
        "COUNTY OF": county,
        "To the Surrogates Court County of": county,
        "decedent": dec,
        "a Name": dec,
        "aka": data.get("decedentAKA", ""),
        "Name_petitioner": pet,
        "1": pet,
        "Domicile or Principal Office": data.get("petitionerStreet", ""),
        "City Village or Town": data.get("petitionerCity", ""),
        "State": data.get("petitionerState", ""),
        "Zip Code": data.get("petitionerZip", ""),
        "Citizen of": data.get("petitionerCitizenship", "U.S.A."),
        "b Date of death": data.get("decedentDOD", ""),
        "c Place of death": data.get("decedentPlaceOfDeath", ""),
        "d Domicile Street": data.get("decedentStreet", ""),
        "City Town Village": data.get("decedentCity", ""),
        "County": data.get("decedentCounty", ""),
        "State_2": data.get("decedentState", ""),
        "e Citizen of": data.get("decedentCitizenship", "U.S.A."),
        "Date of Will": data.get("willDate", ""),
        "Names of All Witnesses to Will": witnesses,
        "Date of Codicil": data.get("codicilDate", ""),
        "follows Enter NONE or specify 1": data.get("noOtherWill", "NONE"),
        "the nature of the confidential relationship 1": "NONE",
        "Improved real property in New York State": data.get("improvedRealProperty", ""),
        "Unimproved real property in New York State": data.get("unimprovedRealProperty", ""),
        "Estimated gross rents for a period of 18 months": data.get("grossRents18mo", ""),
        "the estate except as follows Enter NONE or specify": data.get("otherAssets", "NONE"),
        "but less than": data.get("personalPropertyValue", ""),
        # Surviving relatives dropdowns
        "Dropdown 5a": dropdown_vals[0],
        "Dropdown 5b": dropdown_vals[1],
        "Dropdown 5c": dropdown_vals[2],
        "Dropdown 5d": dropdown_vals[3],
        "Dropdown 5e": dropdown_vals[4],
        "Dropdown 5f": dropdown_vals[5],
        "Dropdown 5g": dropdown_vals[6],
        # Prayer / letters (page 4) — only fill the matching "to" field
        "Petitioner_1": letters_to if "Testamentary" in lt else "",
        "Petitioner_2": letters_to if "Testamentary" in lt else "",
        "Letters of Trusteeship to 1": letters_to if "Trusteeship" in lt else "",
        "Letters of Administration cta to": letters_to if "c.t.a" in lt else "",
        "Dated": "",
        "Print Name": pet,

        # ── Oath and Designation (page 5) ───────────────────────────────────────
        "STATE OF NEW YORK": "New York",
        "COUNTY OF_2": county,
        "ss": county,
        "OATH OF": pet,
        "Surrogates Court of": county,
        "My domicile is": pet_addr,
        "Street Address": data.get("petitionerStreet", ""),
        "Print Name_3": data.get("attorneyName") or "Jessica Wilson, Esq.",
        "Signature of Attorney": data.get("attorneyName") or "Jessica Wilson, Esq.",
        "Print Name_4": data.get("attorneyName") or "Jessica Wilson, Esq.",
        "Firm Name": data.get("firmName") or "Law Office of Jessica Wilson",
        "Tel No": data.get("attorneyPhone") or "(212) 739-1736",
        "Email": data.get("attorneyEmail", ""),
        "Address of Attorney": data.get("firmAddress") or "221 Columbia Street, Brooklyn NY 11231",

        # ── Attesting Witness (page 10) ─────────────────────────────────────────
        "COUNTY OF_7": county,
        "WILL OF 1": data.get("decedentFirstName", ""),
        "WILL OF 2": data.get("decedentLastName", ""),
        "aka 1": data.get("decedentAKA", ""),
        "File_2": data.get("fileNo", ""),
        "STATE OF NEW YORK_5": "New York",
        "COUNTY OF_8": county,
        "I have been shown check one": "X",
        "the original instrument dated": data.get("willDate", ""),
        "purporting to be the last Will and TestamentCodicil of the abovenamed decedent": "X",
        "and I saw the other witness es": witnesses,
        "I am making this affidavit at the request of 1": pet,
    }

    # Letters type checkboxes (text fields — fill with "X")
    if "Testamentary" in lt:
        fields["Letters Testamentary"] = "X"
        fields["EXECUTOR"] = "X"
    elif "Trusteeship" in lt:
        fields["Letters of Trusteeship"] = "X"
    elif "c.t.a" in lt:
        fields["Letters of Administration cta"] = "X"
        fields["ADMINISTRATOR cta"] = "X"
    elif "Temporary" in lt:
        fields["Temporary Administration"] = "X"

    # Petitioner interest
    if "Executor" in data.get("petitionerInterest", ""):
        fields["Executor s named in decedents Will"] = "X"
    if data.get("petitionerIsAttorney") == "Yes":
        fields["is"] = "X"
    else:
        fields["is not an attorney"] = "X"

    # Distributees (page 2, section 6a — 3 columns: name / address / interest)
    name_f = ["1_2", "2_2", "3", "4", "5", "6", "7"]
    addr_f = ["1_3", "2_3", "3_2", "4_2", "5_2", "6_2", "7_2"]
    int_f  = [f"Interest or Nature of Fiduciary Status {i}" for i in range(1, 8)]
    for i, dist in enumerate(data.get("distributees", [])[:7]):
        if dist.get("name"):
            fields[name_f[i]] = dist["name"]
            fields[addr_f[i]] = f"{dist.get('address', '')} | {dist.get('citizenship', '')}"
            fields[int_f[i]]  = dist.get("relationship", "Distributee")

    return fields


def generate_probate_docs(data):
    """
    Returns list of (filename, bytes) for the full probate packet:
      - P-1 Petition (pages 1-4)
      - Combined Verification, Oath and Designation (page 5)
      - Affidavit of Attesting Witness (page 10) — omitted if self-proving will
    Fills the source PDF only once for efficiency.
    """
    template = os.path.join(PROBATE_TEMPLATES_DIR, "Probate Petition + Oath.pdf")
    fields = _build_probate_fields(data)
    filled = fill_pdf(template, fields)
    last = data.get("decedentLastName", "estate").replace(" ", "_")
    docs = [
        (f"02_Petition_P1_{last}.pdf",        _extract_pages(filled, [0, 1, 2, 3])),
        (f"03_Oath_Designation_{last}.pdf",   _extract_pages(filled, [4])),
    ]
    if not data.get("selfProvingAffidavit"):
        docs.append(
            (f"04_Affidavit_Attesting_Witness_{last}.pdf", _extract_pages(filled, [9]))
        )
    return docs


def fill_probate_pdf(data):
    template = os.path.join(PROBATE_TEMPLATES_DIR, "Probate Petition + Oath.pdf")
    return extract_pdf_pages(template, _build_probate_fields(data), [0, 1, 2, 3])


# ─── ANCILLARY ADMIN PDF (AA-1) ───────────────────────────────────────────────

def fill_ancillary_pdf(data):
    """Fill the AA-1 Ancillary Administration Petition PDF form.

    Field mappings verified against admin_ancil.pdf template:
    - Text Field 19 = Mailing Address (NOT citizenship)
    - Text Field 20 = Citizen of (petitioner 1)
    - Radio Button 2 = Interest of petitioner (/0=Admin, /1=Distributee, /2=Creditor, /3=Other)
    - Text Field 28 = Distributee relationship text
    - Text Field 29 = Other/specify text for interest
    - Text Field 76 = WHEREFORE "Letters to" name (parent-child field)
    - Radio Button 3 = WHEREFORE prayer type (/0=Ancillary Letters, /1=d.b.n.)
    - Text Field 75 = "No other persons interested" paragraph (NOT WHEREFORE)
    """
    dec = decedent_full(data)
    pet = petitioner_full(data)
    letters_to = data.get("lettersTo", "") or pet
    county = data.get("county", "")
    foreign_state = data.get("foreignState", "")

    def v(key, default=""):
        val = str(data.get(key, "") or "").strip()
        return val if val else default

    # Compute total NY property value
    try:
        total = sum(float(data.get(k) or 0) for k in [
            "personalPropertyValue", "improvedRealProperty",
            "unimprovedRealProperty", "grossRents18mo"
        ])
        total_str = f"{total:,.2f}" if total > 0 else "0.00"
    except Exception:
        total_str = ""

    petitioner_address = ", ".join(filter(None, [
        data.get("petitionerStreet", ""),
        data.get("petitionerCity", ""),
        data.get("petitionerState", ""),
        data.get("petitionerZip", "")
    ]))

    # Petitioner interest logic
    pet_interest = v("petitionerInterest", "Distributee")
    is_distributee = pet_interest.lower() == "distributee"

    # Radio button values
    radio_interest_val = "/1" if is_distributee else "/3"
    if pet_interest.lower() == "administrator":
        radio_interest_val = "/0"
    elif pet_interest.lower() == "creditor":
        radio_interest_val = "/2"

    fields = {
        # ── PAGE 1 ────────────────────────────────────────────────
        "Text Field 8":  county,
        "Text Field 9":  dec,
        "Text Field 10": v("decedentAKA"),
        "Text Field 11": foreign_state,
        "Text Field 12": v("fileNo"),
        "Text Field 13": county,

        "Text Field 14": pet,
        "Text Field 15": v("petitionerStreet"),
        "Text Field 16": v("petitionerCity"),
        "Text Field 17": v("petitionerState"),
        "Text Field 18": v("petitionerZip"),
        "Text Field 19": petitioner_address,
        "Text Field 20": v("petitionerCitizenship", "U.S.A."),

        # Interest of petitioner (radio + text)
        "Radio Button 2": radio_interest_val,
        "Text Field 28": v("petitionerRelationship") if is_distributee else "",
        "Text Field 29": "" if is_distributee else pet_interest,

        # Para 2 — Decedent
        "Text Field 30": v("decedentDOD"),
        "Text Field 31": v("decedentPlaceOfDeath"),
        "Text Field 32": v("decedentStreet"),
        "Text Field 33": v("decedentCity"),
        "Text Field 34": v("decedentCounty"),
        "Text Field 35": foreign_state,
        "Text Field 36": v("decedentZip"),
        "Text Field 37": v("decedentCitizenship", "U.S.A."),

        # ── PAGE 2 ────────────────────────────────────────────────
        "Text Field 38": v("foreignLettersDate"),
        "Text Field 39": v("foreignLettersIssuedTo", letters_to),
        "Text Field 40": v("foreignCourtName"),
        "Text Field 41": foreign_state,
        "Text Field 42": v("foreignBondAmount", "0"),

        "Text Field 43": v("personalPropertyValue", "0.00"),
        "Text Field 44": v("improvedRealProperty", "0.00"),
        "Text Field 45": v("unimprovedRealProperty", "0.00"),
        "Text Field 46": v("grossRents18mo", "0.00"),
        "Text Field 47": total_str,

        "Text Field 48": v("otherAssets", "NONE"),
        "Text Field 49": "",

        "Text Field 50": "N/A",

        # ── PAGE 3 ────────────────────────────────────────────────
        # WHEREFORE clause
        "Text Field 76": letters_to,
        "Radio Button 3": "/0",
        "Text Field 1065": "",
        "Text Field 77":   "",
        "Text Field 79":   "NONE",
        "Text Field 80":   "",

        # ── PAGE 4 — Combined Verification, Oath and Designation ──────────────
        "Text Field 85": v("petitionerState", "New York"),
        "Text Field 87": county,
        "Text Field 89": county,
        "Text Field 91": petitioner_address,
        "Text Field 97": pet,
    }

    # Para 6(a) distributees — 3 rows (name / address / interest)
    dist_rows = [
        ("Text Field 57", "Text Field 58", "Text Field 59"),
        ("Text Field 60", "Text Field 61", "Text Field 62"),
        ("Text Field 63", "Text Field 64", "Text Field 65"),
    ]
    for i, dist in enumerate(data.get("distributees", [])[:3]):
        if dist.get("name"):
            nf, af, rf = dist_rows[i]
            fields[nf] = dist["name"]
            fields[af] = dist.get("address", "")
            fields[rf] = dist.get("relationship", "")

    template = os.path.join(PDFS_DIR, "admin_ancil.pdf")
    return fill_pdf(template, fields)


# ─── HELPERS ──────────────────────────────────────────────────────────────────


def decedent_full(data):
    return " ".join(filter(None, [
        data.get("decedentFirstName", ""),
        data.get("decedentMiddleName", ""),
        data.get("decedentLastName", "")
    ]))

def petitioner_full(data):
    return " ".join(filter(None, [
        data.get("petitionerFirstName", ""),
        data.get("petitionerMiddleName", ""),
        data.get("petitionerLastName", "")
    ]))


# ─── ADMINISTRATION PETITION (A-1) ────────────────────────────────────────────

def fill_administration_pdf(data):
    """Fill the A-1 Administration Petition + Oath PDF form.

    NOTE: The notary block on page 5 has mixed fonts in the PDF template.
    This must be fixed in the PDF template itself (Adobe Acrobat), not in code.
    """
    county    = data.get("county", "")
    dec       = decedent_full(data)
    pet       = petitioner_full(data)
    lt        = data.get("lettersType", "Letters of Administration")
    lt_lower  = lt.lower()
    letters_to = data.get("lettersTo", "") or pet

    def v(key, default=""):
        return str(data.get(key, "") or "").strip() or default

    # Letters type flags
    is_limited    = "limited" in lt_lower and "limitation" not in lt_lower
    is_limitation = "limitation" in lt_lower
    is_temporary  = "temporary" in lt_lower
    is_standard   = not any([is_limited, is_limitation, is_temporary])

    # Citizenship flags
    pet_cit = v("petitionerCitizenship", "U.S.A.")
    dec_cit = v("decedentCitizenship",   "U.S.A.")
    pet_us  = "U.S.A" in pet_cit or "usa" in pet_cit.lower()
    dec_us  = "U.S.A" in dec_cit or "usa" in dec_cit.lower()

    is_attorney = data.get("petitionerIsAttorney") == "Yes"

    # Surviving relatives → dropdowns 6a–6h (EPTL 4-1.1 order)
    # Logic: "No" for prior classes, number/Yes for first surviving class, "X" for all after
    surv_keys = [
        "survivingSpouse", "survivingChildren", "survivingIssue",
        "survivingParents", "survivingSiblings", "survivingGrandparents",
        "survivingAuntsUncles", "survivingFirstCousinsOnceRemoved",
    ]
    first_surviving = None
    for idx, key in enumerate(surv_keys):
        raw = data.get(key)
        if raw and str(raw).strip().lower() not in ("false", "0", "no", ""):
            first_surviving = idx
            break
    dropdown_vals = []
    for idx, key in enumerate(surv_keys):
        raw = data.get(key)
        if first_surviving is None:
            dropdown_vals.append("No")
        elif idx < first_surviving:
            dropdown_vals.append("No")
        elif idx == first_surviving:
            s = str(raw).strip()
            dropdown_vals.append(s if s.lower() not in ("true", "yes") else "Yes")
        else:
            dropdown_vals.append("X")

    # Debts
    debt_lines = []
    for key, label in [("mortgageAmount",    "Outstanding Mortgage: ${}"),
                       ("funeralPaid",        "Funeral Expenses Paid: ${}"),
                       ("funeralOutstanding", "Funeral Expenses Outstanding: ${}"),
                       ("miscDebts",          "Misc Debts: {}")]:
        val = (data.get(key, "") or "").strip()
        if val:
            debt_lines.append(label.format(val))
    if not debt_lines:
        debt_lines = ["NONE"]

    pet_addr = ", ".join(filter(None, [
        v("petitionerStreet"), v("petitionerCity"),
        v("petitionerState"), v("petitionerZip"),
    ]))

    fields = {
        # ── PAGE 1: Caption ──────────────────────────────────────────
        "COUNTY OF":                        county.upper(),
        "Estate of 1":                      dec,
        "aka":                              v("decedentAKA"),
        "File No":                          v("fileNo"),
        "TO THE SURROGATES COURT COUNTY OF": county.upper(),

        # Caption checkboxes (letters type)
        "petition for letters of admin":    is_standard,
        "limited admin":                    is_limited,
        "limited admin with lim":           is_limitation,
        "temp admin":                       is_temporary,

        # ── PAGE 1: Petitioner ───────────────────────────────────────
        "Name":                             pet,
        "Domicile":                         v("petitionerStreet"),
        "County":                           v("petitionerCity"),
        "State":                            v("petitionerState"),
        "Zip":                              v("petitionerZip"),
        "yes us citizen":                   pet_us,
        "NO us citizen":                    not pet_us,
        "Distributee of decedent state relationship":
            v("petitionerRelationship") if v("petitionerInterest", "").lower() in ("", "distributee") else "",
        "Otherspecify":
            "" if v("petitionerInterest", "").lower() in ("", "distributee") else v("petitionerInterest"),
        "Mark if Distributee":
            v("petitionerInterest", "").lower() in ("", "distributee"),
        "Mark if other and then specifiy":
            bool(v("petitionerInterest")) and v("petitionerInterest", "").lower() != "distributee",
        "yes attorney":                     is_attorney,
        "NO not an attorney":               not is_attorney,
        "not a convicted felon":            True,

        # ── PAGE 1: Decedent ─────────────────────────────────────────
        "Name_2":                           dec,
        "Domicile_2":                       v("decedentStreet"),
        "City/Town/Village":                v("decedentCity"),
        "State_2":                          v("decedentState"),
        "Zip Code":                         v("decedentZip"),
        "Township of":                      v("decedentCounty", v("decedentCity")),
        "Date of Death":                    v("decedentDOD"),
        "Place of Death":                   v("decedentPlaceOfDeath"),
        "yes us citizen 1":                 dec_us,
        "NO not US Citizen 2":              not dec_us,

        # ── PAGE 2: Property values ──────────────────────────────────
        "gross value personal":             v("personalPropertyValue", "0"),
        "gross value real property":        v("realPropertyValue", "0"),
        "improved":                         bool(nonzero(data.get("improvedRealProperty"))),
        "unimproved":                       bool(nonzero(data.get("unimprovedRealProperty"))),
        "A brief description of each parcel is as follows":
                                            v("realPropertyDescription"),
        "c The estimated gross rent for a period of eighteen 18 months is the sum of":
                                            v("grossRents18mo"),

        # Surviving relatives dropdowns
        "Dropdown 6a": dropdown_vals[0],
        "Dropdown 6b": dropdown_vals[1],
        "Dropdown 6c": dropdown_vals[2],
        "Dropdown 6d": dropdown_vals[3],
        "Dropdown 6e": dropdown_vals[4],
        "Dropdown 6f": dropdown_vals[5],
        "Dropdown 6g": dropdown_vals[6],
        "Dropdown 6h": dropdown_vals[7],

        # ── PAGE 4: Prayer for relief ────────────────────────────────
        "a-process issue letters":          True,
        "c a decree award letters of":      True,
        "9c1":                              is_standard,
        "9c2":                              is_limited,
        "9c3":                              is_limitation,
        "9c4":                              is_temporary,
        "Administration to":                letters_to if is_standard   else "",
        "Limited Administration to":        letters_to if is_limited    else "",
        "Administration with Limitation to": letters_to if is_limitation else "",
        "Temporary Administration to":      letters_to if is_temporary  else "",
        "Dated":                            "",
        "Print Name":                       pet,

        # ── PAGE 1: Petitioner phone (also used on page 5) ───────────
        "Telephone Number":                 v("petitionerPhone", "(212) 739-1736"),

        # ── PAGE 5: Combined Verification, Oath & Designation ────────
        "ss":                               v("petitionerState", "New York"),
        "County of":                        county.upper(),
        "My domicile is":                   pet_addr,
        "before me personally came":        pet,
        "Print Name_3":                     v("attorneyName", "Jessica Wilson, Esq."),
        "Firm Name":                        v("attorneyFirm", "Law Office of Jessica Wilson"),
        "TelNo":                            v("attorneyPhone", "(212) 739-1736"),
        "Address of Attorney":              v("attorneyAddress", "221 Columbia Street, Brooklyn NY 11231"),

        # Wrongful death (always No for standard admin)
        "yes wrongful death":               False,
    }

    # ── PAGE 3: Distributees (full age / sound mind) — up to 8 rows ─
    for i, dist in enumerate(data.get("distributees", [])[:8]):
        if dist.get("name"):
            n = str(i + 1)
            fields[f"Name {n}"]                        = dist["name"]
            fields[f"Relationship {n}"]                = dist.get("relationship", "")
            fields[f"Domicile and Mailing Address {n}"] = dist.get("address", "")
            fields[f"Citizenship {n}"]                 = dist.get("citizenship", "U.S.A.")

    # ── PAGE 3: Debts ────────────────────────────────────────────────
    debt_key = "8 There are no outstanding debts or funeral expenses except Write NONE or state same {}"
    for i, line in enumerate(debt_lines[:9]):
        fields[debt_key.format(i + 1)] = line

    template = os.path.join(ADMIN_TEMPLATES_DIR, "Admin Petition + Oath.pdf")
    return fill_pdf(template, fields)


def fill_nondom_pdf(data):
    """Fill the Non-Domiciliary Administration Petition + Oath PDF form.

    Uses the same field mapping as fill_administration_pdf but with the
    Non Dom template which has additional non-domiciliary specific fields.
    """
    county    = data.get("county", "")
    dec       = decedent_full(data)
    pet       = petitioner_full(data)
    lt        = data.get("lettersType", "Letters of Administration")
    lt_lower  = lt.lower()
    letters_to = data.get("lettersTo", "") or pet

    def v(key, default=""):
        return str(data.get(key, "") or "").strip() or default

    is_limited    = "limited" in lt_lower and "limitation" not in lt_lower
    is_limitation = "limitation" in lt_lower
    is_temporary  = "temporary" in lt_lower
    is_standard   = not any([is_limited, is_limitation, is_temporary])

    pet_cit = v("petitionerCitizenship", "U.S.A.")
    dec_cit = v("decedentCitizenship",   "U.S.A.")
    pet_us  = "U.S.A" in pet_cit or "usa" in pet_cit.lower()
    dec_us  = "U.S.A" in dec_cit or "usa" in dec_cit.lower()

    is_attorney = data.get("petitionerIsAttorney") == "Yes"

    surv_keys = [
        "survivingSpouse", "survivingChildren", "survivingIssue",
        "survivingParents", "survivingSiblings", "survivingGrandparents",
        "survivingAuntsUncles", "survivingFirstCousinsOnceRemoved",
    ]
    first_surviving = None
    for idx, key in enumerate(surv_keys):
        raw = data.get(key)
        if raw and str(raw).strip().lower() not in ("false", "0", "no", ""):
            first_surviving = idx
            break
    dropdown_vals = []
    for idx, key in enumerate(surv_keys):
        raw = data.get(key)
        if first_surviving is None:
            dropdown_vals.append("No")
        elif idx < first_surviving:
            dropdown_vals.append("No")
        elif idx == first_surviving:
            s = str(raw).strip()
            dropdown_vals.append(s if s.lower() not in ("true", "yes") else "Yes")
        else:
            dropdown_vals.append("X")

    debt_lines = []
    for key, label in [("mortgageAmount",    "Outstanding Mortgage: ${}"),
                       ("funeralPaid",        "Funeral Expenses Paid: ${}"),
                       ("funeralOutstanding", "Funeral Expenses Outstanding: ${}"),
                       ("miscDebts",          "Misc Debts: {}")]:
        val = (data.get(key, "") or "").strip()
        if val:
            debt_lines.append(label.format(val))
    if not debt_lines:
        debt_lines = ["NONE"]

    pet_addr = ", ".join(filter(None, [
        v("petitionerStreet"), v("petitionerCity"),
        v("petitionerState"), v("petitionerZip"),
    ]))

    # Foreign letters info for non-domiciliary
    foreign_state = v("foreignState", v("decedentState"))

    fields = {
        "COUNTY OF":                        county.upper(),
        "Estate of 1":                      dec,
        "aka":                              v("decedentAKA"),
        "File No":                          v("fileNo"),
        "TO THE SURROGATES COURT COUNTY OF": county.upper(),

        "Name":                             pet,
        "Domicile":                         v("petitionerStreet"),
        "County":                           v("petitionerCity"),
        "State":                            v("petitionerState"),
        "Zip":                              v("petitionerZip"),
        "Mailing address is":               pet_addr,
        "yes us citizen":                   pet_us,
        "NO us citizen":                    not pet_us,
        "Distributee of decedent state relationship":
            v("petitionerRelationship") if v("petitionerInterest", "").lower() in ("", "distributee") else "",
        "Otherspecify":
            "" if v("petitionerInterest", "").lower() in ("", "distributee") else v("petitionerInterest"),
        "Mark if Distributee":
            v("petitionerInterest", "").lower() in ("", "distributee"),
        "Mark if other and then specifiy":
            bool(v("petitionerInterest")) and v("petitionerInterest", "").lower() != "distributee",
        "yes attorney":                     is_attorney,
        "NO not an attorney":               not is_attorney,
        "not a convicted felon":            True,

        "Name_2":                           dec,
        "Domicile_2":                       v("decedentStreet"),
        "City/Town/Village":                v("decedentCity"),
        "State_2":                          v("decedentState"),
        "Zip Code":                         v("decedentZip"),
        "Township of":                      v("decedentCounty", v("decedentCity")),
        "Date of Death":                    v("decedentDOD"),
        "Place of Death":                   v("decedentPlaceOfDeath"),
        "yes us citizen 1":                 dec_us,
        "NO not US Citizen 2":              not dec_us,

        "gross value personal":             v("personalPropertyValue", "0"),
        "gross value real property":        v("realPropertyValue", "0"),
        "improved":                         bool(nonzero(data.get("improvedRealProperty"))),
        "unimproved":                       bool(nonzero(data.get("unimprovedRealProperty"))),
        "A brief description of each parcel is as follows":
                                            v("realPropertyDescription"),
        "c The estimated gross rent for a period of eighteen 18 months is the sum of":
                                            v("grossRents18mo"),

        "Dropdown 6a": dropdown_vals[0],
        "Dropdown 6b": dropdown_vals[1],
        "Dropdown 6c": dropdown_vals[2],
        "Dropdown 6d": dropdown_vals[3],
        "Dropdown 6e": dropdown_vals[4],
        "Dropdown 6f": dropdown_vals[5],
        "Dropdown 6g": dropdown_vals[6],
        "Dropdown 6h": dropdown_vals[7],

        "a-process issue letters":          True,
        "c a decree award letters of":      True,
        "9c1":                              is_standard,
        "9c2":                              is_limited,
        "9c3":                              is_limitation,
        "9c4":                              is_temporary,
        "Administration to":                letters_to if is_standard   else "",
        "Limited Administration to":        letters_to if is_limited    else "",
        "Administration with Limitation to": letters_to if is_limitation else "",
        "Temporary Administration to":      letters_to if is_temporary  else "",
        "Dated":                            "",
        "Print Name":                       pet,

        "Telephone Number":                 v("petitionerPhone", "(212) 739-1736"),

        "ss":                               v("petitionerState", "New York"),
        "My domicile is":                   pet_addr,
        "before me personally came":        pet,
        "Print Name_3":                     v("attorneyName", "Jessica Wilson, Esq."),
        "Firm Name":                        v("attorneyFirm", "Law Office of Jessica Wilson"),
        "TelNo":                            v("attorneyPhone", "(212) 739-1736"),
        "Address of Attorney":              v("attorneyAddress", "221 Columbia Street, Brooklyn NY 11231"),

        "yes wrongful death":               False,
    }

    # Distributees — full age / sound mind (rows 1-8)
    for i, dist in enumerate(data.get("distributees", [])[:8]):
        if dist.get("name"):
            n = str(i + 1)
            fields[f"Name {n}"]                        = dist["name"]
            fields[f"Relationship {n}"]                = dist.get("relationship", "")
            fields[f"Domicile and Mailing Address {n}"] = dist.get("address", "")
            fields[f"Citizenship {n}"]                 = dist.get("citizenship", "U.S.A.")

    # Debts
    debt_key = "8 There are no outstanding debts or funeral expenses except Write NONE or state same {}"
    for i, line in enumerate(debt_lines[:9]):
        fields[debt_key.format(i + 1)] = line

    template = os.path.join(ADMIN_TEMPLATES_DIR, "Non Dom Petition + Oath.pdf")
    return fill_pdf(template, fields)


# ─── FAMILY TREE WORKSHEET (FT-1) ─────────────────────────────────────────────

def fill_ft1_pdf(data):
    """Fill the actual FT-1 Family Tree Affidavit court form PDF."""
    dec_name    = decedent_full(data)
    aka         = data.get("decedentAKA", "")
    file_no     = data.get("fileNo", "")
    pet_name    = petitioner_full(data)
    pet_addr    = ", ".join(filter(None, [
        data.get("petitionerStreet", ""),
        data.get("petitionerCity", ""),
        data.get("petitionerState", "NY"),
        data.get("petitionerZip", ""),
    ]))
    pet_rel     = data.get("petitionerRelationship", "")
    marital     = (data.get("maritalStatus") or "").strip()
    spouse_name = (data.get("spouseName") or "").strip()
    divorce_yr  = (data.get("divorceYear") or "").strip()

    # Distribute distributees into sections by relationship keyword
    all_dists = data.get("distributees", [])

    def _match(d, *keywords):
        return any(k in (d.get("relationship") or "").lower() for k in keywords)

    children  = [d for d in all_dists if _match(d, "child", "son", "daughter")]
    siblings  = [d for d in all_dists if _match(d, "brother", "sister", "sibling")]
    nieces    = [d for d in all_dists if _match(d, "niece", "nephew")]
    mat_aunts = [d for d in all_dists if _match(d, "maternal aunt", "maternal uncle")]
    pat_aunts = [d for d in all_dists if _match(d, "paternal aunt", "paternal uncle")]
    cousins   = [d for d in all_dists if _match(d, "cousin")]

    fields = {}

    # ── Header ──────────────────────────────────────────────────────────────────
    fields["128"]         = dec_name
    fields["230"]         = aka
    fields["412"]         = file_no
    letters_type = (data.get("lettersType") or "Letters of Administration").upper()
    fields["Combo Box00"] = letters_type

    # ── Deponent (petitioner) ───────────────────────────────────────────────────
    fields["5a5"] = pet_name
    fields["5b6"] = pet_rel
    fields["5c7"] = pet_addr

    # ── Section 1a: Marriages ───────────────────────────────────────────────────
    if marital == "never_married":
        fields["Check Box01h"] = True
    elif marital == "married" and spouse_name:
        fields["6a9"] = spouse_name
    elif marital == "divorced" and spouse_name:
        fields["6b10"] = spouse_name
        fields["Check Box01a"] = True
        if divorce_yr:
            fields["6a9"] = f"divorced {divorce_yr}"
    elif marital == "widowed" and spouse_name:
        fields["6b10"] = spouse_name
        fields["Check Box01b"] = True

    # ── Section 1b: Children (6 slots) ─────────────────────────────────────────
    child_name_f = ["816",  "917",  "1018",  "1119",  "1220",  "1321"]
    for i, c in enumerate(children[:6]):
        if c.get("name"):
            fields[child_name_f[i]] = c["name"]

    # ── Section 3a: Siblings (6 slots, page 2) ─────────────────────────────────
    sib_name_f = ["27", "28", "29", "30", "31", "32"]
    for i, s in enumerate(siblings[:6]):
        if s.get("name"):
            fields[sib_name_f[i]] = s["name"]

    # ── Section 3b: Nieces/Nephews (7 slots, page 2) ───────────────────────────
    nie_name_f = ["33","34","35","36","37","38","39"]
    for i, n in enumerate(nieces[:7]):
        if n.get("name"):
            fields[nie_name_f[i]] = n["name"]

    # ── Section 4b: Maternal Aunts/Uncles (7 slots, page 3) ────────────────────
    mat_name_f = ["49","50","51","52","53","54","55"]
    for i, a in enumerate(mat_aunts[:7]):
        if a.get("name"):
            fields[mat_name_f[i]] = a["name"]

    # ── Section 5b: Paternal Aunts/Uncles (7 slots, page 4) ────────────────────
    pat_name_f = ["71","72","73","74","75","76","77"]
    for i, a in enumerate(pat_aunts[:7]):
        if a.get("name"):
            fields[pat_name_f[i]] = a["name"]

    template = os.path.join(ADMIN_TEMPLATES_DIR, "Family_Tree_Affidavit_Fill-In.pdf")
    return fill_pdf(template, fields)


def generate_ft1(data):
    return fill_ft1_pdf(data)


# ─── ACCOUNTING EXCEL ─────────────────────────────────────────────────────────

def _calc_commission(total):
    t1 = min(total, 100000)
    t2 = min(max(total - 100000, 0), 200000)
    t3 = min(max(total - 300000, 0), 700000)
    t4 = max(total - 1000000, 0)
    return t1 * 0.05 + t2 * 0.04 + t3 * 0.03 + t4 * 0.025


def generate_accounting_excel(form_data, assets_data):
    """Generate a full Schedules A–H accounting workbook from asset list data."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    decedent = decedent_full(form_data)

    wb = Workbook()
    ws = wb.active
    ws.title = "Accounting"

    # ── Styles ────────────────────────────────────────────────────────────────
    GOLD_FILL   = PatternFill("solid", fgColor="7A5C1E")
    LIGHT_FILL  = PatternFill("solid", fgColor="FDF8EE")
    TOTAL_FILL  = PatternFill("solid", fgColor="F4F1EB")
    HDR_FONT    = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    BOLD        = Font(name="Arial", bold=True, size=11)
    NORMAL      = Font(name="Arial", size=11)
    MONEY       = Font(name="Courier New", size=11)
    LABEL_FONT  = Font(name="Arial", bold=True, size=11)
    thin = Side(style="thin", color="DDDDDD")
    BORDER = Border(bottom=Side(style="thin", color="DDDDDD"))

    def money_fmt(cell):
        cell.number_format = '#,##0.00'
        cell.font = MONEY

    def section_header(row, title):
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
        c = ws.cell(row=row, column=1, value=title)
        c.font = HDR_FONT
        c.fill = GOLD_FILL
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[row].height = 22

    def col_headers(row, *headers):
        for i, h in enumerate(headers, 1):
            c = ws.cell(row=row, column=i, value=h)
            c.font = BOLD
            c.fill = LIGHT_FILL
            c.alignment = Alignment(horizontal="left" if i < len(headers) else "right")

    def total_row(row, label, value):
        c1 = ws.cell(row=row, column=1, value=label)
        c1.font = BOLD
        c1.fill = TOTAL_FILL
        c3 = ws.cell(row=row, column=3, value=value)
        c3.font = Font(name="Courier New", bold=True, size=11)
        c3.fill = TOTAL_FILL
        c3.number_format = '#,##0.00'
        c3.alignment = Alignment(horizontal="right")

    def blank_rows(start_row, count):
        for r in range(start_row, start_row + count):
            ws.cell(row=r, column=1).border = BORDER
            ws.cell(row=r, column=3).border = BORDER
            ws.cell(row=r, column=3).number_format = '#,##0.00'
            ws.cell(row=r, column=3).alignment = Alignment(horizontal="right")

    # ── Column widths ─────────────────────────────────────────────────────────
    ws.column_dimensions['A'].width = 40
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 16

    # ── Title ─────────────────────────────────────────────────────────────────
    r = 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    title_cell = ws.cell(row=r, column=1,
        value=f"Estate of {decedent} — Informal Accounting")
    title_cell.font = Font(name="Arial", bold=True, size=13)
    title_cell.alignment = Alignment(horizontal="center")
    ws.row_dimensions[r].height = 24
    r += 1

    date_cell = ws.cell(row=r, column=1, value=f"Generated: {today()}")
    date_cell.font = NORMAL
    r += 2

    # ── Schedule A — Estate Assets ────────────────────────────────────────────
    section_header(r, "Schedule A — Estate Assets"); r += 1
    col_headers(r, "Institution / Description", "Category", "Value ($)"); r += 1

    sched_a_total = 0.0
    for a in assets_data:
        val = 0.0
        try:
            val = float(str(a.get("value", "0")).replace(",", "").replace("$", "").strip() or 0)
        except Exception:
            pass
        sched_a_total += val
        c1 = ws.cell(row=r, column=1, value=a.get("institution") or a.get("category", ""))
        c1.font = NORMAL
        c2 = ws.cell(row=r, column=2, value=a.get("category", ""))
        c2.font = NORMAL
        c3 = ws.cell(row=r, column=3, value=val)
        money_fmt(c3)
        c3.alignment = Alignment(horizontal="right")
        r += 1
    total_row(r, "Schedule A Total", sched_a_total); r += 2

    # ── Schedule B — Income / Receipts ────────────────────────────────────────
    section_header(r, "Schedule B — Income / Receipts"); r += 1
    col_headers(r, "Description", "", "Amount ($)"); r += 1
    b_start = r
    blank_rows(r, 10); r += 10
    total_row(r, "Schedule B Subtotal", 0); r += 2

    # ── Schedule C — Disbursements ────────────────────────────────────────────
    section_header(r, "Schedule C — Disbursements"); r += 1
    col_headers(r, "Description", "", "Amount ($)"); r += 1
    blank_rows(r, 10); r += 10
    total_row(r, "Schedule C Subtotal", 0); r += 2

    # ── Schedule D — Prior Distributions ─────────────────────────────────────
    section_header(r, "Schedule D — Prior Distributions"); r += 1
    ws.cell(row=r, column=1, value="Prior Distributions").font = NORMAL
    d_cell = ws.cell(row=r, column=3)
    d_cell.number_format = '#,##0.00'
    d_cell.alignment = Alignment(horizontal="right")
    d_cell.border = BORDER
    r += 1
    total_row(r, "Schedule D Total", 0); r += 2

    # ── Schedule E — Commission ───────────────────────────────────────────────
    section_header(r, "Schedule E — Executor/Administrator Commission (NY SCPA)"); r += 1
    commission = _calc_commission(sched_a_total)
    tiers = [
        ("First $100,000 × 5%",       min(sched_a_total, 100000),         0.05),
        ("Next $200,000 × 4%",         min(max(sched_a_total - 100000, 0), 200000), 0.04),
        ("Next $700,000 × 3%",         min(max(sched_a_total - 300000, 0), 700000), 0.03),
        ("Balance over $1,000,000 × 2.5%", max(sched_a_total - 1000000, 0), 0.025),
    ]
    note = ws.cell(row=r, column=1,
        value=f"Commission base (Schedule A total): ${sched_a_total:,.2f}")
    note.font = Font(name="Arial", italic=True, size=10, color="888888")
    r += 1
    for label, base, rate in tiers:
        if base > 0:
            c1 = ws.cell(row=r, column=1, value=label)
            c1.font = NORMAL
            c3 = ws.cell(row=r, column=3, value=base * rate)
            money_fmt(c3)
            c3.alignment = Alignment(horizontal="right")
            r += 1
    total_row(r, "Total Commission", commission); r += 2

    # ── Schedule F — Estate Account Balance ───────────────────────────────────
    section_header(r, "Schedule F — Balance in Estate Account"); r += 1
    ws.cell(row=r, column=1, value="Current balance in estate account").font = NORMAL
    f_cell = ws.cell(row=r, column=3)
    f_cell.number_format = '#,##0.00'
    f_cell.alignment = Alignment(horizontal="right")
    f_cell.border = BORDER
    r += 1
    total_row(r, "Schedule F Balance", 0); r += 2

    # ── Schedule G — Reconciliation ───────────────────────────────────────────
    section_header(r, "Schedule G — Reconciliation"); r += 1
    rows_g = [
        ("Schedule A + B (Total Receipts)", sched_a_total),
        ("Less: Schedule C (Disbursements)", 0),
        ("Less: Schedule D (Prior Distributions)", 0),
        ("Net Balance", sched_a_total),
        ("Schedule F (Estate Account Balance)", 0),
        ("Difference (should be zero)", sched_a_total),
    ]
    for label, val in rows_g:
        c1 = ws.cell(row=r, column=1, value=label)
        c1.font = BOLD if "Net Balance" in label or "Difference" in label else NORMAL
        c3 = ws.cell(row=r, column=3, value=val)
        money_fmt(c3)
        c3.alignment = Alignment(horizontal="right")
        r += 1
    r += 1

    # ── Schedule H — Distribution Plan ────────────────────────────────────────
    section_header(r, "Schedule H — Distribution Plan"); r += 1
    col_headers(r, "Beneficiary / Purpose", "", "Amount ($)"); r += 1
    c1 = ws.cell(row=r, column=1, value="Executor/Administrator Commission (from Sched E)")
    c1.font = NORMAL
    c3 = ws.cell(row=r, column=3, value=commission)
    money_fmt(c3)
    c3.alignment = Alignment(horizontal="right")
    r += 1
    blank_rows(r, 10); r += 10
    total_row(r, "Schedule H Total", 0); r += 1

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ─── LETTER OF AUTHORIZATION ──────────────────────────────────────────────────

def generate_auth_letter(data, asset):
    """Pre-letters letter from nominated executor/administrator authorizing the
    law office to speak with the institution."""
    lt = data.get("lettersType", "")
    role = "executor" if "Testamentary" in lt else "administrator"
    decedent = decedent_full(data)
    petitioner = petitioner_full(data)
    institution = asset.get("institution", "").strip() or "Financial Institution"
    account_no = asset.get("accountNumber", "").strip() or "N/A"

    doc = Document()

    FONT = "Times New Roman"
    SIZE = Pt(12)

    def _run(para, text, bold=False):
        r = para.add_run(text)
        r.font.name = FONT
        r.font.size = SIZE
        r.bold = bold
        return r

    def line(text="", bold=False, space_after=6):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(space_after)
        if text:
            _run(p, text, bold=bold)
        return p

    line(today(), space_after=12)
    line("")
    line(institution, bold=True)
    line(f"Re: Estate of {decedent}")
    line(f"    Account No.: {account_no}", space_after=12)
    line("")
    line("To Whom It May Concern:", space_after=12)
    line("")
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(12)
    _run(p, f"I, {petitioner}, am the nominated {role} of the Estate of "
         f"{decedent}, deceased. I hereby authorize the Law Office of Jessica Wilson "
         f"to discuss, obtain information about, and act on my behalf with respect to "
         f"the above-referenced account and any other accounts held by the above-named estate.")
    line("Please extend your full cooperation to our office upon request.", space_after=24)
    line("")
    line("Sincerely,", space_after=48)
    line("")
    line("")
    line(petitioner)

    return make_docx_bytes(doc)


# ─── LETTER OF INSTRUCTION ────────────────────────────────────────────────────

def generate_instruction_letter(data, asset, marshal_action="check"):
    """Post-letters letter requesting the institution marshal assets."""
    lt = data.get("lettersType", "")
    role = "executor" if "Testamentary" in lt else "administrator"
    letters_label = "Letters Testamentary" if "Testamentary" in lt else "Letters of Administration"
    decedent = decedent_full(data)
    petitioner = petitioner_full(data)
    county = data.get("county", "")
    dod = data.get("decedentDOD", "")
    institution = asset.get("institution", "").strip() or "Financial Institution"
    account_no = asset.get("accountNumber", "").strip() or "N/A"
    signer_key = data.get("signer", "Jessica Wilson")
    signer = SIGNERS.get(signer_key, signer_key)

    if marshal_action == "transfer":
        marshal_text = "transfer all funds to the estate account"
    else:
        marshal_text = f"remit payment by check payable to 'Estate of {decedent}'"

    doc = Document()

    FONT = "Times New Roman"
    SIZE = Pt(12)

    def _run(para, text, bold=False):
        r = para.add_run(text)
        r.font.name = FONT
        r.font.size = SIZE
        r.bold = bold
        return r

    def line(text="", bold=False, space_after=6):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(space_after)
        if text:
            _run(p, text, bold=bold)
        return p

    line(today(), space_after=12)
    line("")
    line(institution, bold=True)
    line(f"Re: Estate of {decedent}")
    line(f"    Account No.: {account_no}", space_after=12)
    line("")
    line("Dear Sir or Madam:", space_after=12)
    line("")
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(12)
    _run(p, f"Our office represents {petitioner}, the duly appointed {role} of the "
         f"Estate of {decedent}, who died on {dod}. "
         f"{letters_label} were issued by the Surrogate's Court, {county} County.")
    p2 = doc.add_paragraph()
    p2.paragraph_format.space_after = Pt(12)
    _run(p2, f"Please marshal all assets held in the above-referenced account and "
         f"{marshal_text} at your earliest convenience. "
         f"Please find enclosed a certified copy of the Letters.")
    line("Please do not hesitate to contact our office should you require any additional "
         "information or documentation.", space_after=24)
    line("")
    line("Very truly yours,", space_after=48)
    line("")
    line("")
    line(signer)

    return make_docx_bytes(doc)
