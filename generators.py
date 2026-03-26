"""
Document generators for NY Surrogate's Court Probate HQ
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
    if proceeding == "Probate":
        left_lines = [
            "PROBATE PROCEEDING, WILL OF",
            "",
            f"    {decedent.upper()},",
        ]
    else:
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

    # Assets — use individual asset tracker entries if available, else summary fields
    tracked_assets = [a for a in data.get("assets", []) if a.get("institution")]
    asset_lines = []

    if tracked_assets:
        for a in tracked_assets:
            val = nonzero(a.get("value"))
            inst = a.get("institution", "")
            cat = a.get("category", "")
            acct = a.get("accountNumber", "")
            desc = f"{cat} – {inst}" if cat and inst else (inst or cat)
            if acct:
                desc += f" (acct ...{acct[-4:]})" if len(acct) >= 4 else f" (acct {acct})"
            if val:
                asset_lines.append(f"{desc}:  ${val}")
            else:
                asset_lines.append(desc)

        # Also include real property from summary fields (not tracked in asset cards)
        ir = nonzero(data.get("improvedRealProperty"))
        ur = nonzero(data.get("unimprovedRealProperty"))
        rd = (data.get("realPropertyDescription") or "").strip()
        gr = nonzero(data.get("grossRents18mo"))
        if ir: asset_lines.append(f"Improved Real Property (NY):  ${ir}")
        if ur: asset_lines.append(f"Unimproved Real Property (NY):  ${ur}")
        if rd: asset_lines.append(f"Description:  {rd}")
        if gr: asset_lines.append(f"Gross Rents (18 months):  ${gr}")
    else:
        pp = nonzero(data.get("personalPropertyValue"))
        ir = nonzero(data.get("improvedRealProperty"))
        ur = nonzero(data.get("unimprovedRealProperty"))
        rd = (data.get("realPropertyDescription") or "").strip()
        gr = nonzero(data.get("grossRents18mo"))
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
                    # Size the X to fit the field box
                    h = widget.rect.height
                    widget.text_fontsize = max(5, min(h - 1, 10))
                elif len(s) > 0:
                    # Auto-shrink font for long text in narrow fields
                    w = widget.rect.width
                    current_size = widget.text_fontsize or 12
                    # Rough estimate: each char ~0.6x font size in width
                    est_width = len(s) * current_size * 0.5
                    if est_width > w and w > 0:
                        fitted = w / (len(s) * 0.5)
                        widget.text_fontsize = max(5, min(fitted, current_size))
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
    #
    # Auto-derive from distributees if manual fields are empty.
    # Map relationship keywords → EPTL class index
    surv_keys = [
        "survivingSpouse", "survivingChildren", "survivingParents",
        "survivingSiblings", "survivingGrandparents", "survivingAuntsUncles",
        "survivingFirstCousinsOnceRemoved",
    ]

    # Check if any manual surviving fields are filled
    has_manual = any(
        data.get(k) and str(data.get(k)).strip().lower() not in ("false", "0", "no", "")
        for k in surv_keys
    )

    if not has_manual:
        # Auto-derive from distributees' relationships
        rel_class_map = {
            "spouse": 0, "husband": 0, "wife": 0,
            "son": 1, "daughter": 1, "child": 1, "children": 1, "issue": 1,
            "grandchild": 1, "grandson": 1, "granddaughter": 1,
            "mother": 2, "father": 2, "parent": 2,
            "sister": 3, "brother": 3, "sibling": 3, "half-sister": 3, "half-brother": 3,
            "niece": 3, "nephew": 3,
            "grandmother": 4, "grandfather": 4, "grandparent": 4,
            "aunt": 5, "uncle": 5, "cousin": 5,
        }
        class_counts = [0] * 7
        for dist in data.get("distributees", []):
            rel = (dist.get("relationship") or "").strip().lower()
            for keyword, cls_idx in rel_class_map.items():
                if keyword in rel:
                    class_counts[cls_idx] += 1
                    break

        first_surviving = None
        for idx, count in enumerate(class_counts):
            if count > 0:
                first_surviving = idx
                break

        dropdown_vals = []
        for idx, count in enumerate(class_counts):
            if first_surviving is None:
                dropdown_vals.append("No")
            elif idx < first_surviving:
                dropdown_vals.append("No")
            elif idx == first_surviving:
                dropdown_vals.append(str(count) if count > 0 else "Yes")
            else:
                dropdown_vals.append("X")
    else:
        # Use manual surviving fields
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
        "aka2": "",  # second AKA line in caption (renamed from field "1")
        "Name_petitioner": pet,
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

    # Petitioner interest — default to Executor for probate
    pet_interest = data.get("petitionerInterest", "")
    if "Executor" in pet_interest or not pet_interest:
        fields["Executor s named in decedents Will"] = "X"
    else:
        fields["Other Specify Check"] = "X"
        fields["Other Specify"] = pet_interest
    if data.get("petitionerIsAttorney") == "Yes":
        fields["is"] = "X"
    else:
        fields["is not an attorney"] = "X"

    # ── Distributees — route to correct petition section ─────────────────────
    # For probate: "interest" = description of legacy/devise under the will
    # For administration: fall back to relationship
    all_dists = [d for d in data.get("distributees", []) if d.get("name")]

    # Split into 4 groups
    primary_adults    = [d for d in all_dists if (d.get("beneficiaryType") or "primary") == "primary" and not d.get("isMinor")]
    primary_minors    = [d for d in all_dists if (d.get("beneficiaryType") or "primary") == "primary" and d.get("isMinor")]
    successor_adults  = [d for d in all_dists if d.get("beneficiaryType") == "successor" and not d.get("isMinor")]
    successor_minors  = [d for d in all_dists if d.get("beneficiaryType") == "successor" and d.get("isMinor")]

    def _interest(dist):
        interest = (dist.get("interest") or "").strip()
        return interest if interest else dist.get("relationship", "Distributee")

    def _minor_desc(dist):
        """Build the 7b description: name, DOB, relationship, domicile, guardian."""
        parts = [dist.get("name", "")]
        if dist.get("dob"):
            parts.append(f"DOB: {dist['dob']}")
        if dist.get("relationship"):
            parts.append(dist["relationship"])
        if dist.get("address"):
            parts.append(dist["address"])
        if dist.get("guardianInfo"):
            parts.append(f"Guardian: {dist['guardianInfo']}")
        return "; ".join(parts)

    # Page 2, section 6a — Primary beneficiaries (8 rows)
    p2_6a_name = ["1_2", "2_2", "3", "4", "5", "6", "7", "8"]
    p2_6a_addr = ["1_3", "2_3", "3_2", "4_2", "5_2", "6_2", "7_2", "8"]
    p2_6a_int  = [f"Interest or Nature of Fiduciary Status {i}" for i in range(1, 9)]
    for i, dist in enumerate(primary_adults[:8]):
        fields[p2_6a_name[i]] = dist["name"]
        fields[p2_6a_addr[i]] = f"{dist.get('address', '')} | {dist.get('citizenship', '')}"
        fields[p2_6a_int[i]]  = _interest(dist)

    # Page 2, section 6b — Primary beneficiaries under disability (6 rows)
    p2_7b_name = ["1_4", "2_4", "3_3", "4_3", "5_3", "6_3"]
    p2_7b_addr = ["1_5", "2_5", "3_4", "4_4", "5_4", "6_4"]
    p2_7b_int  = [f"Interest or Nature of Fiduciary Status {i}_2" for i in range(1, 7)]
    for i, dist in enumerate(primary_minors[:6]):
        fields[p2_7b_name[i]] = _minor_desc(dist)
        fields[p2_7b_addr[i]] = dist.get("address", "")
        fields[p2_7b_int[i]]  = _interest(dist)

    # Page 3, section 7a — Substitute executors, trustees, guardians, other beneficiaries (8 rows)
    p3_6a_name = ["1_9", "2_9", "3_5", "4_5", "5_5", "6_5", "7_3", "8_2"]
    p3_6a_addr = ["1_10", "2_10", "3_6", "4_6", "5_6", "6_6", "7_4", "8_2"]
    p3_6a_int  = [f"Interest or Nature of Fiduciary Status {i}_3" for i in range(1, 9)]
    for i, dist in enumerate(successor_adults[:8]):
        fields[p3_6a_name[i]] = dist["name"]
        fields[p3_6a_addr[i]] = f"{dist.get('address', '')} | {dist.get('citizenship', '')}"
        fields[p3_6a_int[i]]  = _interest(dist)

    # Page 3, section 7b — Persons under disability from section 7a (7 rows)
    p3_7b_name = ["1_11", "2_11", "3_7", "4_7", "5_7", "6_7", "7_5"]
    p3_7b_addr = ["1_12", "2_12", "3_8", "4_8", "5_8", "6_8", "7_6"]
    p3_7b_int  = [f"Interest or Nature of Fiduciary Status {i}_4" for i in range(1, 8)]
    for i, dist in enumerate(successor_minors[:7]):
        fields[p3_7b_name[i]] = _minor_desc(dist)
        fields[p3_7b_addr[i]] = dist.get("address", "")
        fields[p3_7b_int[i]]  = _interest(dist)

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


# ─── SCHEDULE D(a) — POST-DECEASED DISTRIBUTEE ──────────────────────────────

def fill_schedule_da_pdf(data, dist):
    """Fill the Schedule D(a) form for a distributee who post-deceased the decedent.

    Field mapping (by rect position on page):
    - Combo Box0:     County
    - Text Field167:  File #
    - Text Field164:  Estate of (decedent name)
    - Text Field165:  a/k/a
    - Text Field168:  1. Name of post-deceased distributee
    - Text Field169:  Date of Death of post-deceased
    - Text Field170:  Relationship to decedent
    - Text Field171:  Last permanent address (domicile)
    - Check Box01:    Yes/No fiduciary appointed
    - Text Field174:  3(a) Fiduciary row 1 (name / address / citizenship / court)
    - Text Field176:  3(a) Fiduciary row 2
    - Text Field175:  3(b) Distributee row 1
    - Text Field177:  3(b) Distributee row 2
    - Text Field178:  3(b) Distributee row 3
    - Text Field179:  3(b) Distributee row 4
    - Text Field180:  3(b) Distributee row 5
    - Text Field181:  3(b) Distributee row 6
    """
    dec = decedent_full(data)
    aka = data.get("decedentAKA", "")
    county = data.get("county", "")
    file_no = data.get("fileNo", "")

    fields = {}

    # Header
    fields["Combo Box0"] = county
    fields["Text Field167"] = file_no
    fields["Text Field164"] = dec
    fields["Text Field165"] = aka

    # Section 1 — post-deceased distributee info
    fields["Text Field168"] = dist.get("name", "")
    fields["Text Field169"] = dist.get("postDeceasedDOD", "")
    fields["Text Field170"] = dist.get("relationship", "")
    fields["Text Field171"] = dist.get("address", "")

    # Section 2 — fiduciary yes/no
    has_fiduciary = dist.get("hasFiduciary", False)
    if has_fiduciary:
        fields["Check Box01"] = True

    # Section 3(a) — fiduciary details (2 rows)
    fid = dist.get("fiduciary", {})
    if has_fiduciary and fid:
        row1_parts = [fid.get("name", ""), fid.get("address", ""),
                      fid.get("citizenship", ""), fid.get("court", "")]
        fields["Text Field174"] = "     ".join(p for p in row1_parts if p)
        fields["Text Field176"] = fid.get("row2", "")

    # Section 3(b) — post-deceased person's distributees (up to 6 rows)
    pd_dists = dist.get("postDeceasedDistributees", [])
    row_fields = ["Text Field175", "Text Field177", "Text Field178",
                  "Text Field179", "Text Field180", "Text Field181"]
    for idx, pd in enumerate(pd_dists[:6]):
        parts = [pd.get("name", ""), pd.get("address", ""),
                 pd.get("citizenship", ""), pd.get("relationship", "")]
        fields[row_fields[idx]] = "     ".join(p for p in parts if p)

    template = os.path.join(PROBATE_TEMPLATES_DIR,
                            "Schedule D(a)- Distributee Who Post-Deceased Decedent.pdf")
    return fill_pdf(template, fields)


# ─── ADMIN CTA (ADMINISTRATION C.T.A.) ──────────────────────────────────────

def fill_cta_pdf(data):
    """Fill the Administration C.T.A. petition PDF (SCPA 1418/1419).

    Used when a will exists but the named executor cannot serve
    (died, resigned, or was removed).

    Template: templates/Probate/probcta.pdf  (7 pages)
    Page 0: CTA-1 Petition (sections 1-2)
    Page 1: CTA-1 Petition cont (sections 3-7, WHEREFORE, signatures)
    Page 2: Combined Verification, Oath & Designation
    Page 3: Corporate Verification, Consent & Designation
    Page 4: CTA Citation
    Page 5: CTA-3 Waiver/Renunciation
    Page 6: P-12 Affidavit of No Debt
    """
    county   = data.get("county", "")
    dec      = decedent_full(data)
    aka      = data.get("decedentAKA", "")
    pet      = petitioner_full(data)
    file_no  = data.get("fileNo", "")
    letters_to = data.get("lettersTo", "") or pet

    def v(key, default=""):
        return str(data.get(key, "") or "").strip() or default

    pet_street = v("petitionerStreet")
    pet_city   = v("petitionerCity")
    pet_state  = v("petitionerState", "NY")
    pet_zip    = v("petitionerZip")
    pet_cit    = v("petitionerCitizenship", "US Citizen")
    dec_street = v("decedentStreet")
    dec_city   = v("decedentCity")
    dec_county = v("decedentCounty", county)
    dec_state  = v("decedentState", "NY")
    dec_zip    = v("decedentZip")

    # CTA-specific fields
    orig_county   = v("ctaOriginalCounty")
    orig_date     = v("ctaOriginalDate")
    orig_executor = v("ctaOriginalExecutor")
    exec_reason   = v("ctaExecutorReason")  # died / resigned / removed
    pet_interest  = v("ctaPetitionerInterest", "Residuary Beneficiary")
    is_attorney   = v("ctaAdminIsAttorney", "no")

    # Estate values
    personal = v("personalPropertyValue", "0")
    real_imp = v("improvedRealProperty", "0")
    real_unimp = v("unimprovedRealProperty", "0")
    gross_rents = v("grossRents18mo", "0")

    # Distributees
    dists = data.get("distributees", [])

    fields = {}

    # ═══ PAGE 0: CTA-1 Petition ═══════════════════════════════════════════
    fields["Decedent_Name"]   = dec
    fields["Decedent_AKA"]    = aka
    fields["File_No"]         = file_no

    # Section 1(a) — Petitioner info
    fields["TextField13[0]"]  = pet                  # Petitioner name
    fields["TextField14[0]"]  = pet_street            # Street address
    fields["TextField15[0]"]  = pet_city              # City/Village/Town
    fields["TextField16[0]"]  = pet_state             # State... but wrong field?
    fields["TextField17[0]"]  = dec_county            # County
    fields["TextField18[0]"]  = pet_state             # State
    fields["TextField19[0]"]  = pet_zip               # Zip
    fields["TextField20[0]"]  = ""                    # Telephone
    fields["TextField21[0]"]  = ""                    # Mailing address if different

    # Citizenship checkboxes
    if "us" in pet_cit.lower() or "citizen" in pet_cit.lower():
        fields["CheckBox1[0]"] = True   # USA
    else:
        fields["CheckBox2[0]"] = True   # Other
        fields["TextField22[0]"] = pet_cit

    # Second petitioner (leave blank)
    fields["TextField23[0]"]  = ""

    # Interest checkboxes
    if "sole" in pet_interest.lower():
        fields["CheckBox5[0]"] = True   # Sole Beneficiary
    elif "residuary" in pet_interest.lower():
        fields["CheckBox6[0]"] = True   # Residuary Beneficiary
    else:
        fields["CheckBox7[0]"] = True   # Other
        fields["TextField32[0]"] = pet_interest

    # 1(b) — Is admin CTA an attorney?
    if is_attorney.lower() == "yes":
        fields["CheckBox8[0]"] = True   # is an attorney
    else:
        fields["CheckBox9[0]"] = True   # is not an attorney

    # Section 2 — Original probate info
    fields["TextField33[0]"]  = orig_county           # County where probated
    fields["TextField34[0]"]  = orig_date             # Date probated
    fields["TextField35[0]"]  = orig_executor         # Original executor name
    fields["TextField36[0]"]  = ""                    # "who on [date]..."

    # Reason checkboxes
    if exec_reason == "died":
        fields["CheckBox10[0]"] = True
    elif exec_reason == "resigned":
        fields["CheckBox11[0]"] = True
    elif exec_reason == "removed":
        fields["CheckBox12[0]"] = True

    # ═══ PAGE 1: Petition continued ═══════════════════════════════════════
    # Section 3 — Persons with prior/equal right (SCPA 1418)
    if len(dists) > 0:
        d = dists[0]
        fields["TextField37[0]"]  = d.get("name", "")
        fields["TextField38[0]"]  = ""                 # Description of legacy
        fields["TextField39[0]"]  = d.get("relationship", "")
        fields["TextField40[0]"]  = ""                 # Mailing address
        fields["TextField41[0]"]  = ""                 # Fiduciary status
        fields["TextField42[0]"]  = d.get("address", "")
        fields["TextField43[0]"]  = ""                 # Additional line

    # Section 4 — Other beneficiaries
    if len(dists) > 1:
        d = dists[1]
        fields["TextField44[0]"]  = d.get("name", "")
        fields["TextField45[0]"]  = ""
        fields["TextField46[0]"]  = d.get("relationship", "")
        fields["TextField47[0]"]  = ""
        fields["TextField67[0]"]  = ""
        fields["TextField49[0]"]  = d.get("address", "")
        fields["TextField50[0]"]  = ""

    # Section 6 — Debts
    fields["TextField51[0]"]  = ""    # Debts/funeral expenses (leave for manual)

    # Section 7 — Estate values
    fields["TextField52[0]"]  = personal    # Personal property
    fields["TextField53[0]"]  = real_imp    # Improved real property
    fields["TextField54[0]"]  = real_unimp  # Unimproved real property
    fields["TextField55[0]"]  = gross_rents # Estimated gross rents 18 months
    fields["TextField56[0]"]  = ""          # Other assets / cause of action

    # WHEREFORE
    fields["Petitioner"]      = letters_to   # Letters of Admin CTA to
    fields["TextField58[0]"]  = ""           # Other relief
    fields["TextField59[0]"]  = today()      # Dated

    # Petitioner signatures (print names)
    fields["TextField60[0]"]  = pet          # Signature line 1 (print name)
    fields["TextField61[0]"]  = ""           # Signature line 2
    fields["TextField62[0]"]  = pet          # Print name 1
    fields["TextField63[0]"]  = ""           # Print name 2

    # ═══ PAGE 2: Verification, Oath & Designation ═════════════════════════
    fields["STATE OF_F02"]             = "NEW YORK"
    fields["COUNTY OF_F13"]            = county.upper()
    fields["of_F24"]                   = county  # Designation county
    fields["(Street Address)_F35"]     = pet_street
    fields["(City/Town/Village)_F46"]  = pet_city
    fields["(State)_F57"]              = pet_state
    fields["(Print Name)_F68"]         = pet
    fields["came_F79"]                 = pet     # "came [name]"
    fields["Date0"]                    = today()
    fields["Year1"]                    = ""

    # ═══ PAGE 3: Corporate Verification ═══════════════════════════════════
    fields["TextField86[0]"]  = ""    # State
    fields["TextField87[0]"]  = ""    # County
    # Corporate fields left blank (filled when corporate petitioner)

    # ═══ PAGE 4: Citation ═════════════════════════════════════════════════
    fields["TextField108[0]"] = file_no               # File No
    fields["TextField109[0]"] = county                 # County
    # TO lines (cite parties)
    cite_names = [d.get("name", "") for d in dists if d.get("disposition") == "citation"]
    to_fields = ["TextField110[0]", "TextField111[0]", "TextField112[0]",
                 "TextField113[0]", "TextField114[0]"]
    for i, name in enumerate(cite_names[:5]):
        fields[to_fields[i]] = name

    fields["TextField115[0]"] = pet                    # Petitioner name
    fields["TextField116[0]"] = pet_street             # Petitioner domicile
    fields["TextField117[0]"] = f"{pet_city}, {pet_state}"
    fields["TextField118[0]"] = county                 # County
    fields["TextField120[0]"] = dec                    # Estate of
    fields["TextField123[0]"] = f"{dec_street}, {dec_city}, {dec_state}"  # Domicile
    fields["TextField124[0]"] = dec_county             # County
    fields["TextField125[0]"] = ""                     # Surrogate name
    fields["TextField126[0]"] = letters_to             # Letters to
    fields["TextField135[0]"] = ""                     # Attorney for petitioner
    fields["TextField136[0]"] = ""                     # Telephone
    fields["TextField137[0]"] = ""                     # Address of attorney

    # ═══ PAGE 5: CTA-3 Waiver/Renunciation ═══════════════════════════════
    fields["TextField138[0]"] = county                 # County
    fields["TextField139[0]"] = dec                    # Will of
    fields["TextField140[0]"] = aka                    # a/k/a
    fields["TextField141[0]"] = file_no                # File No
    fields["TextField142[0]"] = ""                     # Undersigned name (filled by signer)

    # Interest checkboxes (page 5)
    # CheckBox13 = beneficiary with equal/prior right
    # CheckBox14 = beneficiary of estate
    # CheckBox15 = creditor
    # CheckBox16 = other

    fields["TextField144[0]"] = letters_to             # Letters CTA to
    fields["TextField145[0]"] = county                 # County for notary

    # ═══ PAGE 6: P-12 Affidavit of No Debt ═══════════════════════════════
    fields["TextField161[0]"] = county                 # County
    fields["TextField162[0]"] = dec                    # Will of
    fields["TextField163[0]"] = aka                    # a/k/a
    fields["TextField164[0]"] = file_no                # File No
    fields["TextField165[0]"] = county                 # County (SS:)
    fields["TextField166[0]"] = pet                    # Deponent name
    fields["TextField170[0]"] = pet_street             # Resides at
    fields["TextField168[0]"] = dec_county             # County of residence
    fields["TextField169[0]"] = pet_state              # State
    fields["TextField171[0]"] = personal               # Estate value

    template = os.path.join(PROBATE_TEMPLATES_DIR, "probcta.pdf")
    return fill_pdf(template, fields)


# ─── WAIVER OF CONSENT AND RENUNCIATION (A-8 Individual) ────────────────────

def fill_waiver_individual_pdf(data, dist):
    """Fill the A-8 Waiver, Consent and Renunciation form for an individual distributee.

    Field mapping (Waiver of Consent and Renunciation.pdf):
    - county of 111:        County
    - Estate of 111:        Estate name (decedent)
    - aka of 111:           a/k/a
    - File No_7:            File number
    - Print Name_5:         Distributee print name
    - Street Address:       Distributee street address
    - TownStateZip:         Distributee city/state/zip
    - Relationship_2:       Relationship to decedent
    - be issued to:         Letters to (administrator name)
    - COUNTY OF_5:          County (notary section)
    - Name of Attorney:     Attorney name
    - 1_4:                  Attorney address line 1
    - 2_4:                  Attorney address line 2
    - Telephone Number_2:   Attorney phone
    """
    dec = decedent_full(data)
    aka = data.get("decedentAKA", "")
    county = data.get("county", "")
    file_no = data.get("fileNo", "")
    letters_to = data.get("lettersTo", "") or petitioner_full(data)

    dist_name = dist.get("name", "")
    dist_addr = dist.get("address", "")
    dist_rel = dist.get("relationship", "")

    fields = {
        "county of 111":     county,
        "Estate of 111":     dec,
        "aka of 111":        aka,
        "File No_7":         file_no,
        "Print Name_5":      dist_name,
        "Street Address":    dist_addr,
        "TownStateZip":      "",
        "Relationship_2":    dist_rel,
        "be issued to":      letters_to,
        "COUNTY OF_5":       county,
        "Name of Attorney":  data.get("attorneyName", "Jessica Wilson, Esq."),
        "1_4":               data.get("firmAddress", "221 Columbia Street"),
        "2_4":               data.get("firmAddress2", "Brooklyn NY 11231"),
        "Telephone Number_2": data.get("attorneyPhone", "(212) 739-1736"),
    }

    # Split address into street + city/state/zip if comma-separated
    if dist_addr and ", " in dist_addr:
        parts = dist_addr.split(", ", 1)
        fields["Street Address"] = parts[0]
        fields["TownStateZip"] = parts[1] if len(parts) > 1 else ""
    else:
        fields["Street Address"] = dist_addr

    template = os.path.join(ADMIN_TEMPLATES_DIR, "Waiver of Consent and Renunciation.pdf")
    return fill_pdf(template, fields)


# ─── WAIVER & CONSENT CORPORATE (A-9) ───────────────────────────────────────

def fill_waiver_corporate_pdf(data, dist):
    """Fill the A-9 Waiver & Consent form for a corporate distributee.

    Field mapping (Waiver & Consent Corp.pdf):
    - county of 112:                     County
    - Estate of 112:                     Estate name (decedent)
    - aka of 112:                        a/k/a
    - File No_8:                         File number
    - Name of Corporation:               Corporation name
    - a citation ... be issued to:       Letters to (administrator name)
    - COUNTY OF_6:                       County (notary section)
    - Name of Attorney_2:               Attorney name
    - 1_5:                               Attorney address line 1
    - 2_5:                               Attorney address line 2
    - Telephone Number_3:               Attorney phone
    """
    dec = decedent_full(data)
    aka = data.get("decedentAKA", "")
    county = data.get("county", "")
    file_no = data.get("fileNo", "")
    letters_to = data.get("lettersTo", "") or petitioner_full(data)

    corp_name = dist.get("name", "")

    fields = {
        "county of 112":     county,
        "Estate of 112":     dec,
        "aka of 112":        aka,
        "File No_8":         file_no,
        "Name of Corporation": corp_name,
        "a citation in this matter and consents that Letters of Administration be issued to": letters_to,
        "COUNTY OF_6":       county,
        "Name of Attorney_2": data.get("attorneyName", "Jessica Wilson, Esq."),
        "1_5":               data.get("firmAddress", "221 Columbia Street"),
        "2_5":               data.get("firmAddress2", "Brooklyn NY 11231"),
        "Telephone Number_3": data.get("attorneyPhone", "(212) 739-1736"),
    }

    template = os.path.join(ADMIN_TEMPLATES_DIR, "Waiver & Consent Corp.pdf")
    return fill_pdf(template, fields)


# ─── CITATION (Admin) ───────────────────────────────────────────────────────

def fill_citation_pdf(data):
    """Fill the Citation PDF form (post-filing document).

    Field mapping (Citation.pdf):
    - SURROGATES COURT:                  County
    - File No_2:                         File number
    - A petition having been duly filed by: Petitioner name
    - who is domicilied at:              Petitioner address
    - county 444:                        County
    - decree should not be made in the estate of: Decedent name
    - decree should not be made in the estate of 222: a/k/a
    - lately domiciled at:               Decedent address
    - lately domiciled at the county of: Decedent county
    - estate of decentt to 222345:       Letters to name
    - Attorney for Petitioner:           Attorney name
    - TelNo_2:                           Attorney phone
    - Address of Attorney_2:             Attorney address
    """
    dec = decedent_full(data)
    pet = petitioner_full(data)
    county = data.get("county", "")
    file_no = data.get("fileNo", "")
    letters_to = data.get("lettersTo", "") or pet

    pet_addr = ", ".join(filter(None, [
        data.get("petitionerStreet", ""),
        data.get("petitionerCity", ""),
        data.get("petitionerState", ""),
        data.get("petitionerZip", ""),
    ]))
    dec_addr = ", ".join(filter(None, [
        data.get("decedentStreet", ""),
        data.get("decedentCity", ""),
        data.get("decedentState", ""),
        data.get("decedentZip", ""),
    ]))

    fields = {
        "SURROGATES COURT":                 county,
        "File No_2":                        file_no,
        "A petition having been duly filed by": pet,
        "who is domicilied at":             pet_addr,
        "county 444":                       county,
        "decree should not be made in the estate of": dec,
        "decree should not be made in the estate of 222": data.get("decedentAKA", ""),
        "lately domiciled at":              dec_addr,
        "lately domiciled at the county of": data.get("decedentCounty", county),
        "estate of decentt to 222345":      letters_to,
        "Attorney for Petitioner":          data.get("attorneyName", "Jessica Wilson, Esq."),
        "TelNo_2":                          data.get("attorneyPhone", "(212) 739-1736"),
        "Address of Attorney_2":            data.get("attorneyAddress", "221 Columbia Street, Brooklyn NY 11231"),
    }

    template = os.path.join(ADMIN_TEMPLATES_DIR, "Citation.pdf")
    return fill_pdf(template, fields)


# ─── AFFIDAVIT OF SERVICE ───────────────────────────────────────────────────

def fill_affidavit_of_service_pdf(data):
    """Fill the Affidavit of Service PDF header fields (post-filing document).

    Only fills county, estate, and file number. Person-served details are
    completed manually after service is actually made.

    Field mapping (Affid of Service.pdf):
    - county of 113:    County
    - Estate of 113:    Estate name (decedent)
    - File No_9:        File number
    """
    dec = decedent_full(data)
    county = data.get("county", "")
    file_no = data.get("fileNo", "")

    fields = {
        "county of 113":  county,
        "Estate of 113":  dec,
        "File No_9":      file_no,
    }

    template = os.path.join(ADMIN_TEMPLATES_DIR, "Affid of Service.pdf")
    return fill_pdf(template, fields)


# ─── NOTICE OF APPLICATION (SCPA 1005) ──────────────────────────────────────

def fill_notice_of_application_pdf(data):
    """Fill the Notice of Application (SCPA 1005) PDF form.

    Field mapping (Notice of App SCPA 1005.pdf):
    - County of 56:      County
    - Estate of 56:      Estate name (decedent)
    - aka of 56:         a/k/a
    - File No_3:         File number
    - petitioner:        Petitioner name
    - 3 petitioner prays...: Letters to name
    - Name of Distributee 1/2/3:              Distributee names (section 4a)
    - Domicile and Post Office Address 1/2/3: Distributee addresses (section 4a)
    - Name of Distributee 1_2/2_2/3_2:       Distributee names (section 4b)
    - Domicile and Post Office Address 1_2/2_2/3_2: Distributee addresses (section 4b)
    - Attorney for Petitioner_2:             Attorney name
    - Print Name_4:                          Attorney print name
    - Address Office:                        Attorney address
    """
    dec = decedent_full(data)
    pet = petitioner_full(data)
    county = data.get("county", "")
    file_no = data.get("fileNo", "")
    aka = data.get("decedentAKA", "")
    letters_to = data.get("lettersTo", "") or pet

    fields = {
        "County of 56":    county,
        "Estate of 56":    dec,
        "aka of 56":       aka,
        "File No_3":       file_no,
        "petitioner":      pet,
        "1 an application for Letters of Administration upon the estate of the abovenamed decedent has been made": "",
        "3 petitioner prays that a decree be made directing the issuance of Letters of Administration to": letters_to,
        "Attorney for Petitioner_2": data.get("attorneyName", "Jessica Wilson, Esq."),
        "Print Name_4":    data.get("attorneyName", "Jessica Wilson, Esq."),
        "Address Office":  data.get("attorneyAddress", "221 Columbia Street, Brooklyn NY 11231"),
    }

    # Section 4a — distributees with full age and sound mind (up to 3)
    dists = data.get("distributees", [])
    name_fields_a = ["Name of Distributee 1", "Name of Distributee 2", "Name of Distributee 3"]
    addr_fields_a = ["Domicile and Post Office Address 1", "Domicile and Post Office Address 2",
                     "Domicile and Post Office Address 3"]
    name_fields_b = ["Name of Distributee 1_2", "Name of Distributee 2_2", "Name of Distributee 3_2"]
    addr_fields_b = ["Domicile and Post Office Address 1_2", "Domicile and Post Office Address 2_2",
                     "Domicile and Post Office Address 3_2"]

    for i, dist in enumerate(dists[:3]):
        if dist.get("name"):
            fields[name_fields_a[i]] = dist["name"]
            fields[addr_fields_a[i]] = dist.get("address", "")

    template = os.path.join(ADMIN_TEMPLATES_DIR, "Notice of App SCPA 1005.pdf")
    return fill_pdf(template, fields)


# ─── AFFIDAVIT OF MAILING ───────────────────────────────────────────────────

def fill_affidavit_of_mailing_pdf(data):
    """Fill the Affidavit of Mailing PDF form header and distributee addresses.

    Field mapping (Affid of Mailing.pdf):
    - County of 57:      County
    - Estate of 57:      Estate name (decedent)
    - aka of 57:         a/k/a
    - File No_4:         File number
    - COUNTY OF_2:       County (venue)
    - whose post office address is / _2 / _3 / _4 / _5 / _6 / _7 / _8:
                         Distributee addresses (up to 8)
    """
    dec = decedent_full(data)
    county = data.get("county", "")
    file_no = data.get("fileNo", "")
    aka = data.get("decedentAKA", "")

    fields = {
        "County of 57":   county,
        "Estate of 57":   dec,
        "aka of 57":      aka,
        "File No_4":      file_no,
        "COUNTY OF_2":    county,
    }

    # Fill distributee addresses (up to 8 slots)
    addr_fields = [
        "whose post office address is",
        "whose post office address is_2",
        "whose post office address is_3",
        "whose post office address is_4",
        "whose post office address is_5",
        "whose post office address is_6",
        "whose post office address is_7",
        "whose post office address is_8",
    ]
    dists = data.get("distributees", [])
    for i, dist in enumerate(dists[:8]):
        if dist.get("name"):
            addr = dist.get("address", "")
            fields[addr_fields[i]] = f"{dist['name']}, {addr}" if addr else dist["name"]

    template = os.path.join(ADMIN_TEMPLATES_DIR, "Affid of Mailing.pdf")
    return fill_pdf(template, fields)


# ─── AFFIDAVIT OF REGULARITY ────────────────────────────────────────────────

def fill_affidavit_of_regularity_pdf(data):
    """Fill the Affidavit of Regularity PDF form (post-filing document).

    Field mapping (Affid of Regularity.pdf):
    - county of 10:      County
    - Estate of 10:      Estate name (decedent)
    - aka of10:          a/k/a
    - File No_6:         File number
    - COUNTY OF_4:       County (venue)
    - being duly sworn deposes and says: Attorney name (deponent)
    - 1 That heshe is the attorney for: Petitioner name
    - Name 1_5 / Name 2_5:              Waiver distributee names (section c)
    - Address 1_3 / Address 2_3:         Waiver distributee addresses (section c)
    - Name 1_3 / Name 2_3:              Citation distributee names (section a)
    - Address 1 / Address 2:            Citation distributee addresses (section a)
    """
    dec = decedent_full(data)
    pet = petitioner_full(data)
    county = data.get("county", "")
    file_no = data.get("fileNo", "")
    aka = data.get("decedentAKA", "")
    attorney = data.get("attorneyName", "Jessica Wilson, Esq.")

    fields = {
        "county of 10":    county,
        "Estate of 10":    dec,
        "aka of10":        aka,
        "File No_6":       file_no,
        "COUNTY OF_4":     county,
        "being duly sworn deposes and says": attorney,
        "1 That heshe is the attorney for": pet,
    }

    # Separate distributees by disposition
    dists = data.get("distributees", [])
    waiver_dists = [d for d in dists if d.get("disposition") == "waiver" and d.get("name")]
    citation_dists = [d for d in dists if d.get("disposition") == "citation" and d.get("name")]

    # Section (c) — waivers (up to 2)
    waiver_name_fields = ["Name 1_5", "Name 2_5"]
    waiver_addr_fields = ["Address 1_3", "Address 2_3"]
    for i, d in enumerate(waiver_dists[:2]):
        fields[waiver_name_fields[i]] = d["name"]
        fields[waiver_addr_fields[i]] = d.get("address", "")

    # Section (a) — citations (up to 2)
    cite_name_fields = ["Name 1_3", "Name 2_3"]
    cite_addr_fields = ["Address 1", "Address 2"]
    for i, d in enumerate(citation_dists[:2]):
        fields[cite_name_fields[i]] = d["name"]
        fields[cite_addr_fields[i]] = d.get("address", "")

    template = os.path.join(ADMIN_TEMPLATES_DIR, "Affid of Regularity.pdf")
    return fill_pdf(template, fields)


# ─── PROPOSED DECREE ─────────────────────────────────────────────────────────

def fill_proposed_decree_pdf(data):
    """Fill the Proposed Decree PDF form (post-filing document).

    Field mapping (Proposed Decree.pdf):
    - in and for the County of:          County
    - Estate of 9:                       Estate name (decedent)
    - aka of 9:                          a/k/a
    - FileNo:                            File number
    - A petition having been filed by:   Petitioner name
    - of the goods chattels...:          Letters to name
    - that:                              Petitioner name (competency statement)
    - ORDERED AND DECREED...:            Letters to name
    - ORDERED AND DECREED... 22:         Letters to name (bond dispensed)
    - bond having been filed and approved...: Bond amount
    - bond having been filed:            Checkbox — bond filed
    - bond having been dispensed:        Checkbox — bond dispensed
    """
    dec = decedent_full(data)
    pet = petitioner_full(data)
    county = data.get("county", "")
    file_no = data.get("fileNo", "")
    aka = data.get("decedentAKA", "")
    letters_to = data.get("lettersTo", "") or pet
    bond_amount = data.get("bondAmount", "")

    fields = {
        "in and for the County of":       county,
        "Estate of 9":                    dec,
        "aka of 9":                       aka,
        "FileNo":                         file_no,
        "A petition having been filed by": pet,
        "of the goods chattels and credits of the abovenamed decedent be granted to": letters_to,
        "that":                           pet,
        "is in all respects competent to act as administrat": "",
        "ORDERED AND DECREED that Letters of Administration issue to": letters_to,
        "ORDERED AND DECREED that Letters of Administration issue to 22": letters_to,
    }

    # Bond: filed vs dispensed
    if bond_amount and bond_amount.strip() not in ("0", ""):
        fields["bond having been filed"] = True
        fields["bond having been filed and approved in the amount of"] = bond_amount
    else:
        fields["bond having been dispensed"] = True

    template = os.path.join(ADMIN_TEMPLATES_DIR, "Proposed Decree.pdf")
    return fill_pdf(template, fields)


# ─── SCHEDULE A — NONMARITAL PERSONS ────────────────────────────────────────

def fill_schedule_a_pdf(data, dist):
    """Fill Schedule A (Nonmarital Persons) for a per-distributee schedule.

    Field mapping (Schedule A Nonmarital Persons.pdf):
    - County of 2:             County
    - Estate of 2:             Estate name (decedent)
    - aka of 2:                a/k/a
    - File:                    File number
    - Name of alleged distributee: Distributee name
    - Date of birth:           Distributee DOB
    - Relationship to decedent: Relationship
    - Name of father:          Father name
    - Name of mother:          Mother name
    """
    dec = decedent_full(data)
    county = data.get("county", "")
    file_no = data.get("fileNo", "")
    aka = data.get("decedentAKA", "")

    fields = {
        "County of 2":              county,
        "Estate of 2":              dec,
        "aka of 2":                 aka,
        "File":                     file_no,
        "Name of alleged distributee": dist.get("name", ""),
        "Date of birth":            dist.get("dob", ""),
        "Relationship to decedent": dist.get("relationship", ""),
        "Name of father":           dist.get("fatherName", ""),
        "Name of mother":           dist.get("motherName", ""),
    }

    template = os.path.join(ADMIN_TEMPLATES_DIR, "Schedule A Nonmarital Persons.pdf")
    return fill_pdf(template, fields)


# ─── SCHEDULE B — ADOPTION ──────────────────────────────────────────────────

def fill_schedule_b_pdf(data, dist):
    """Fill Schedule B (Adoption) for a per-distributee schedule.

    Field mapping (Sched B Adoption.pdf):
    - County of 3:             County
    - Estate of 3:             Estate name (decedent)
    - aka of 3:                a/k/a
    - File_2:                  File number
    - Name of child:           Adopted child name
    - Relationship to decedent prior to adoption: Prior relationship
    - Date of adoption:        Adoption date
    - If yesname of adoptive father or mother: Adoptive parent name
    - Name of the adoptive parent: Adoptive parent name
    """
    dec = decedent_full(data)
    county = data.get("county", "")
    file_no = data.get("fileNo", "")
    aka = data.get("decedentAKA", "")

    fields = {
        "County of 3":     county,
        "Estate of 3":     dec,
        "aka of 3":        aka,
        "File_2":          file_no,
        "Name of child":   dist.get("name", ""),
        "Relationship to decedent prior to adoption": dist.get("priorRelationship", ""),
        "Date of adoption": dist.get("adoptionDate", ""),
        "If yesname of adoptive father or mother": dist.get("adoptiveParent", ""),
        "Name of the adoptive parent": dist.get("adoptiveParent", ""),
    }

    template = os.path.join(ADMIN_TEMPLATES_DIR, "Sched B Adoption.pdf")
    return fill_pdf(template, fields)


# ─── SCHEDULE C — INFANTS ───────────────────────────────────────────────────

def fill_schedule_c_pdf(data, dist):
    """Fill Schedule C (Infants) for a per-distributee schedule.

    Field mapping (Sched C Infants.pdf):
    - County of 4:             County
    - Estate of 4:             Estate name (decedent)
    - aka of 4:                a/k/a
    - File_3:                  File number
    - Name_3:                  Infant name
    - Date of birth 1:         DOB line 1
    - Date of birth 2:         DOB line 2
    - Relationship to the decedent: Relationship
    - With whom does the infant reside: Residence info
    - Name of mother_2:        Mother name
    - Is she alive:            Mother alive
    - Name of Father:          Father name
    - Is he alive:             Father alive
    - If yes name and address of guardian: Guardian info
    """
    dec = decedent_full(data)
    county = data.get("county", "")
    file_no = data.get("fileNo", "")
    aka = data.get("decedentAKA", "")

    fields = {
        "County of 4":     county,
        "Estate of 4":     dec,
        "aka of 4":        aka,
        "File_3":          file_no,
        "Name_3":          dist.get("name", ""),
        "Date of birth 1": dist.get("dob", ""),
        "Relationship to the decedent": dist.get("relationship", ""),
        "With whom does the infant reside": dist.get("residesWithWhom", ""),
        "Name of mother_2": dist.get("motherName", ""),
        "Is she alive":    dist.get("motherAlive", ""),
        "Name of Father":  dist.get("fatherName", ""),
        "Is he alive":     dist.get("fatherAlive", ""),
        "If yes name and address of guardian": dist.get("guardianInfo", ""),
    }

    template = os.path.join(ADMIN_TEMPLATES_DIR, "Sched C Infants.pdf")
    return fill_pdf(template, fields)


# ─── SCHEDULE D — DISABILITY ────────────────────────────────────────────────

def fill_schedule_d_pdf(data, dist):
    """Fill Schedule D (Disability) for a per-distributee schedule.

    Field mapping (Sched D Disability.pdf):
    - County of 5:             County
    - Estate of 5:             Estate name (decedent)
    - aka of 5:                a/k/a
    - File_4:                  File number
    - 1 Name:                  Person's name
    - Relationship:            Relationship to decedent
    - Residence:               Residence address
    - With whom does this person reside: Caretaker info
    - If this person is in prison name of prison: Prison name
    - If yesgive nametitle and address 1: Court-appointed attorney line 1
    - If yesgive nametitle and address 2: Court-appointed attorney line 2
    - If nodescribe nature of disability 1: Disability description line 1
    - If nodescribe nature of disability 2: Disability description line 2
    - If nogive name and address of relative or friend... 1: Interested person line 1
    - If nogive name and address of relative or friend... 2: Interested person line 2
    """
    dec = decedent_full(data)
    county = data.get("county", "")
    file_no = data.get("fileNo", "")
    aka = data.get("decedentAKA", "")

    fields = {
        "County of 5":     county,
        "Estate of 5":     dec,
        "aka of 5":        aka,
        "File_4":          file_no,
        "1 Name":          dist.get("name", ""),
        "Relationship":    dist.get("relationship", ""),
        "Residence":       dist.get("address", ""),
        "With whom does this person reside": dist.get("residesWithWhom", ""),
        "If this person is in prison name of prison": dist.get("prisonName", ""),
        "If yesgive nametitle and address 1": dist.get("courtAttorneyInfo", ""),
        "If yesgive nametitle and address 2": "",
        "If nodescribe nature of disability 1": dist.get("disabilityDescription", ""),
        "If nodescribe nature of disability 2": "",
        "If nogive name and address of relative or friend interested in his or her welfare 1":
            dist.get("interestedPerson", ""),
        "If nogive name and address of relative or friend interested in his or her welfare 2": "",
    }

    template = os.path.join(ADMIN_TEMPLATES_DIR, "Sched D Disability.pdf")
    return fill_pdf(template, fields)


# ─── WORD TEMPLATE GENERATORS ───────────────────────────────────────────────

def generate_waiver_probate(data, dist):
    """Generate the P-4 Waiver of Process, Consent to Probate Word document
    for a Probate proceeding distributee.

    Template placeholders (Waiver_Probate.docx):
    - _________________  (county, in 'County of ___')
    - No actual bracket-style placeholders; uses blanks for manual fill.
    We replace the county blank and leave signature blanks for manual completion.
    """
    doc = Document(os.path.join(WORD_TEMPLATES_DIR, "Waiver_Probate.docx"))
    county = data.get("county", "")

    replace_in_doc(doc, {
        "County of _________________": f"County of {county}",
    })

    return make_docx_bytes(doc)


def generate_notice_of_probate(data):
    """Generate the Notice of Probate + Affidavit of Mailing Word document.

    Template placeholders (Notice_of_Probate.docx):
    - [COUNTY]:        County name
    - [DECEDENT]:      Decedent full name
    - [DECEDENT AKA]:  a/k/a
    - [county]:        County (lowercase placeholder)
    - [Petitioner]:    Petitioner name
    """
    doc = Document(os.path.join(WORD_TEMPLATES_DIR, "Notice_of_Probate.docx"))
    dec = decedent_full(data)
    pet = petitioner_full(data)
    county = data.get("county", "")
    aka = data.get("decedentAKA", "")

    replace_in_doc(doc, {
        "[COUNTY]":       county.upper(),
        "[DECEDENT]":     dec,
        "[DECEDENT AKA]": aka,
        "[county]":       county,
        "[Petitioner]":   pet,
    })

    return make_docx_bytes(doc)


def generate_bond_affidavit(data):
    """Generate the Bond Affidavit Word document.

    Template uses hardcoded sample data (PHILLIP WILSON-CAMHI / KINGS).
    We replace sample names/values with actual case data.

    Template placeholders (Bond_Affidavit.docx):
    - COUNTY OF KINGS:                     County
    - PHILLIP WILSON-CAMHI:                Petitioner name (appears 3x)
    - 9 Mills Road, Stony Brook, New York 11790: Petitioner address
    - 31 years of age:                     Decedent age at death
    - February 2018:                       Month/Year for notary
    """
    doc = Document(os.path.join(WORD_TEMPLATES_DIR, "Bond_Affidavit.docx"))
    pet = petitioner_full(data)
    county = data.get("county", "")

    # Calculate decedent age at death
    age_str = ""
    try:
        dob = data.get("decedentDOB", "")
        dod = data.get("decedentDOD", "")
        if dob and dod:
            from datetime import datetime as _dt
            dt_dob = _dt.strptime(dob, "%m/%d/%Y")
            dt_dod = _dt.strptime(dod, "%m/%d/%Y")
            age = dt_dod.year - dt_dob.year - (
                (dt_dod.month, dt_dod.day) < (dt_dob.month, dt_dob.day))
            age_str = str(age)
    except Exception:
        pass

    # Build relationship string
    pet_rel = data.get("petitionerRelationship", "Distributee")

    replace_in_doc(doc, {
        "COUNTY OF KINGS":         f"COUNTY OF {county.upper()}",
        "PHILLIP WILSON-CAMHI":    pet,
        "9 Mills Road, Stony Brook, New York 11790": ", ".join(filter(None, [
            data.get("petitionerStreet", ""),
            data.get("petitionerCity", ""),
            data.get("petitionerState", ""),
            data.get("petitionerZip", ""),
        ])),
        "31 years of age":         f"{age_str} years of age" if age_str else "__ years of age",
        "Distributee of said deceased": f"{pet_rel} of said deceased",
        "February 2018":           datetime.now().strftime("%B %Y"),
    })

    return make_docx_bytes(doc)


def generate_petition_scpa_2203(data):
    """Generate the Petition SCPA 2203 (Voluntary Accounting) Word document.

    This template uses hardcoded sample data. We replace the county header.
    Most fields require manual completion as the template is a filled sample.

    Template: Petition_SCPA_2203.docx
    """
    doc = Document(os.path.join(WORD_TEMPLATES_DIR, "Petition_SCPA_2203.docx"))
    county = data.get("county", "")

    replace_in_doc(doc, {
        "COUNTY OF BRONX": f"COUNTY OF {county.upper()}",
    })

    return make_docx_bytes(doc)


# ─── REFUNDING AGREEMENT (Receipt, Release, Indemnification & Refunding) ──────


def generate_refunding_agreement(data):
    """Generate the Receipt, Release, Indemnification & Refunding Agreement.

    Template: RRI_Refunding_Agreement.docx (converted from legacy .doc)

    Auto-fills case header info (county, decedent, executor, date of death).
    Bracketed optional clauses (e.g. [WHEREAS...], [his/her]) are left as-is
    for the attorney to select/edit manually in Word.

    Placeholders replaced:
    - COUNTY OF SUFFOLK          → actual county
    - DECEDENT (in header/body)  → decedent full name
    - EXECUTOR (in header/body)  → petitioner/executor name
    - "died on DATE"             → date of death (long format)
    - "County of COUNTY"         → county name
    - EXEC (commission para)     → executor name
    - BENE1 / BENE 1            → first distributee name (if available)
    """
    doc = Document(os.path.join(WORD_TEMPLATES_DIR, "RRI_Refunding_Agreement.docx"))

    dec = decedent_full(data)
    pet = petitioner_full(data)
    county = data.get("county", "")
    dod = data.get("decedentDOD", "")
    dod_long = format_date_long(dod) if dod else "________"

    # Get first distributee name if available
    dists = data.get("distributees", [])
    bene1_name = ""
    if dists:
        bene1_name = " ".join(filter(None, [
            dists[0].get("firstName", ""),
            dists[0].get("middleName", ""),
            dists[0].get("lastName", ""),
        ]))

    replacements = {
        "COUNTY OF SUFFOLK":  f"COUNTY OF {county.upper()}" if county else "COUNTY OF __________",
        "died on DATE":       f"died on {dod_long}",
        "County of COUNTY":   f"County of {county}" if county else "County of __________",
    }

    # Replace DECEDENT — but only the standalone placeholder, not inside
    # "Decedent" (which appears as a defined term in the body)
    if dec:
        replacements["DECEDENT,"] = f"{dec.upper()},"
        replacements["DECEDENT, (the"] = f"{dec.upper()}, (the"
        replacements["of EXECUTOR, as Executor"] = f"of {pet.upper()}, as Executor"
        replacements["EXECUTOR was appointed"] = f"{pet.upper()} was appointed"

    # Executor signature block and notary
    if pet:
        replacements["EXEC individually"] = f"{pet.upper()} individually"

    # Beneficiary name fill (first bene only — others need manual entry)
    if bene1_name:
        replacements["BENE1  hereby"] = f"{bene1_name.upper()}  hereby"
        replacements["BENE 1"] = bene1_name.upper()

    replace_in_doc(doc, replacements)

    return make_docx_bytes(doc)


# ─── FORMAL ACCOUNTING (Judicial Settlement) ─────────────────────────────────


def generate_formal_accounting(form_data, entries):
    """Generate a formal accounting document (Word) matching Surrogate's Court format.

    Produces cover page, summary statement, and Schedules A through K.
    """
    from docx.enum.section import WD_ORIENT

    doc = Document()

    # ── Page setup ────────────────────────────────────────────────────────────
    section = doc.sections[0]
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1.25)
    section.right_margin = Inches(1.25)

    dec = decedent_full(form_data)
    aka = form_data.get("decedentAKA", "")
    pet = petitioner_full(form_data)
    county = form_data.get("county", "")
    dod = form_data.get("decedentDOD", "")
    dod_long = format_date_long(dod) if dod else "________"
    file_no = form_data.get("fileNo", "")
    proc = form_data.get("proceedingType", "Administration")
    role = "Executor" if proc == "Probate" else "Administrator"

    # Group entries by schedule
    by_sched = {}
    for e in entries:
        s = e.get("schedule", "")
        by_sched.setdefault(s, []).append(e)

    def sched_total(s):
        return sum(float(e.get("amount", 0) or 0) for e in by_sched.get(s, []))

    def add_para(text, bold=False, size=12, alignment=None, space_after=6):
        p = doc.add_paragraph()
        run = p.add_run(text)
        run.bold = bold
        run.font.size = Pt(size)
        run.font.name = "Times New Roman"
        if alignment is not None:
            p.alignment = alignment
        p.paragraph_format.space_after = Pt(space_after)
        return p

    def money(val):
        try:
            v = float(val or 0)
            return f"${v:,.2f}"
        except (ValueError, TypeError):
            return "$0.00"

    # ── COVER PAGE ────────────────────────────────────────────────────────────
    add_para("SURROGATE'S COURT OF THE STATE OF NEW YORK",
             bold=True, size=12, alignment=WD_ALIGN_PARAGRAPH.LEFT)
    add_para(f"COUNTY OF {county.upper()}" if county else "COUNTY OF __________",
             bold=True, size=12, alignment=WD_ALIGN_PARAGRAPH.LEFT)

    # Caption box
    dec_display = dec.upper()
    if aka:
        dec_display += f", a/k/a\n{aka.upper()}"

    caption_left = (
        f"In the Matter of the Judicial Settlement of the Final Account of\n\n"
        f"{pet.upper()}, as {role}\n\n"
        f"of the Estate of\n\n"
        f"{dec_display},\n\n"
        f"{'':>40}Deceased."
    )
    add_para(caption_left, size=12)

    file_line = f"File No:    {file_no}" if file_no else "File No:    __________"
    add_para(file_line, size=12, alignment=WD_ALIGN_PARAGRAPH.RIGHT)

    # Accounting type
    add_para(f"ACCOUNTING BY:\n  {role}", size=11, space_after=12)

    # Court address and period
    add_para(f"TO THE SURROGATE'S COURT OF THE COUNTY OF {county.upper() if county else '__________'}:",
             bold=True, size=11, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
    add_para(
        f"The undersigned does hereby render the account of the proceedings as follows:\n"
        f"Period of account from {dod_long} to {today()}\n"
        f"This is a first and final account containing the following schedules.",
        size=11, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=12
    )

    # ── TABLE OF CONTENTS ─────────────────────────────────────────────────────
    toc_items = [
        ("A", "Principal Received"),
        ("AA", "Subsequent Receipts of Principal"),
        ("A-1", "Realized Increases"),
        ("A-2", "Income Collected"),
        ("B", "Realized Decreases"),
        ("C", "Funeral and Administration Expenses and Taxes"),
        ("C-1", "Unpaid Administration Expenses"),
        ("D", "Creditors' Claims"),
        ("E", "Distributions of Principal"),
        ("F", "New Investments, Exchanges and Stock Distributions"),
        ("G", "Principal Remaining on Hand"),
        ("H", "Interested Parties"),
        ("I", "Computation of Commissions"),
        ("J", "Other Pertinent Facts and Cash Reconciliation"),
        ("K", "Estate Taxes Paid and Allocation of Estate Taxes"),
    ]

    add_para("PRINCIPAL", bold=True, size=11, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
    for sched, title in toc_items[:11]:
        add_para(f"Schedule {sched}        {title}", size=11, space_after=2)
    add_para("OTHER", bold=True, size=11, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
    for sched, title in toc_items[11:]:
        add_para(f"Schedule {sched}        {title}", size=11, space_after=2)

    doc.add_page_break()

    # ── SUMMARY STATEMENT ─────────────────────────────────────────────────────
    add_para("SUMMARY STATEMENT", bold=True, size=12,
             alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
    add_para("COMBINED ACCOUNT", bold=True, size=12,
             alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=12)

    tot_a = sched_total("A")
    tot_aa = sched_total("AA")
    tot_a1 = sched_total("A-1")
    tot_a2 = sched_total("A-2")
    tot_b = sched_total("B")
    tot_c = sched_total("C")
    tot_c1 = sched_total("C-1")
    tot_d = sched_total("D")
    tot_e = sched_total("E")
    tot_g = sched_total("G")

    # Unrealized: G inventory vs market
    unreal_inc = 0
    unreal_dec = 0
    for e in by_sched.get("G", []):
        inv = float(e.get("inventory_value", 0) or 0)
        mkt = float(e.get("market_value", 0) or float(e.get("amount", 0) or 0))
        diff = mkt - inv
        if diff > 0:
            unreal_inc += diff
        else:
            unreal_dec += abs(diff)

    charges = tot_a + tot_aa + tot_a1 + tot_a2 + unreal_inc
    credits = tot_b + tot_c + tot_d + tot_e + unreal_dec
    balance = charges - credits

    # Charges table
    add_para("CHARGES:", bold=True, size=11, space_after=4)
    charge_items = [
        ('Schedule "A"', "(Principal Received)", tot_a),
        ('Schedule "AA"', "(Subsequent Receipts)", tot_aa),
        ('Schedule "A-1"', "(Realized Increases)", tot_a1),
        ('Schedule "A-2"', "(Income Collected)", tot_a2),
        ('Schedule "G"', "(Unrealized Increases)", unreal_inc),
    ]
    for label, desc, val in charge_items:
        add_para(f"{label:20s} {desc:40s} {money(val):>15s}", size=11, space_after=1)
    add_para(f"{'Total Charges':20s} {'':40s} {money(charges):>15s}", bold=True, size=11, space_after=8)

    add_para("CREDITS:", bold=True, size=11, space_after=4)
    credit_items = [
        ('Schedule "B"', "(Realized Decreases)", tot_b),
        ('Schedule "C"', "(Funeral and Administration Expenses)", tot_c),
        ('Schedule "D"', "(Creditors' Claims Actually Paid)", tot_d),
        ('Schedule "E"', "(Distributions)", tot_e),
        ('Schedule "G"', "(Unrealized Decreases)", unreal_dec),
    ]
    for label, desc, val in credit_items:
        add_para(f"{label:20s} {desc:40s} {money(val):>15s}", size=11, space_after=1)
    add_para(f"{'Total Credits':20s} {'':40s} {money(credits):>15s}", bold=True, size=11, space_after=4)
    bal_label = 'Balance on Hand Shown by Schedule "G"'
    add_para(f"{bal_label:40s} {money(balance):>15s}",
             bold=True, size=11, space_after=8)

    doc.add_page_break()

    # ── SUMMARY NARRATIVE ─────────────────────────────────────────────────────
    add_para("SUMMARY STATEMENT", bold=True, size=12,
             alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=12)
    add_para(
        f"The foregoing balance of {money(balance)} consists of "
        f"cash and other property on hand as of {today()}. "
        f"It is subject to deductions of estimated principal commissions "
        f"amounting to {money(_calc_commission(charges))}, "
        f"shown in Schedule I and to the proper charge to principal of expenses of this "
        f"accounting.",
        size=11, space_after=12
    )
    add_para("The attached schedules are part of this account.",
             size=11, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=24)
    add_para(f"{'_' * 35}", size=11, space_after=2)
    add_para(pet, size=11, space_after=2)
    add_para(role, size=11, space_after=0)

    doc.add_page_break()

    # ── SCHEDULE GENERATION HELPER ────────────────────────────────────────────
    def add_schedule(sched_id, title, subtitle, cols, amt_col="amount"):
        """Add a schedule section with header and entry table."""
        estate_header = f"Estate of {dec}"
        if aka:
            estate_header += f", aka {aka}"
        add_para(estate_header, bold=True, size=11,
                 alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
        add_para(f"Schedule {sched_id}", bold=True, size=11,
                 alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
        add_para(subtitle, bold=True, size=11,
                 alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=12)

        sched_entries = by_sched.get(sched_id, [])

        if not sched_entries:
            add_para("None", size=11, space_after=6)
            total = 0
        else:
            # Build table
            table = doc.add_table(rows=1, cols=len(cols))
            table.style = "Table Grid"

            # Header row
            for i, (hdr, _) in enumerate(cols):
                cell = table.rows[0].cells[i]
                cell.text = hdr
                for p in cell.paragraphs:
                    for r in p.runs:
                        r.bold = True
                        r.font.size = Pt(10)
                        r.font.name = "Times New Roman"

            total = 0
            for e in sched_entries:
                row = table.add_row()
                for i, (_, field) in enumerate(cols):
                    val = e.get(field, "") or ""
                    if field == amt_col or field in ("inventory_value", "market_value", "amount"):
                        try:
                            val = money(float(val or 0))
                        except (ValueError, TypeError):
                            val = ""
                    cell = row.cells[i]
                    cell.text = str(val)
                    for p in cell.paragraphs:
                        for r in p.runs:
                            r.font.size = Pt(10)
                            r.font.name = "Times New Roman"
                amt = float(e.get(amt_col, 0) or 0)
                total += amt

        add_para(f"\nTotal Schedule {sched_id}:  {money(total)}",
                 bold=True, size=11, space_after=6)

        doc.add_page_break()
        return total

    # ── GENERATE ALL SCHEDULES ────────────────────────────────────────────────
    add_schedule("A", "Schedule A", "Receipts",
                 [("Description", "description"), ("Institution", "institution"),
                  ("Inventory Value", "amount")])

    add_schedule("AA", "Schedule AA", "Statement of Subsequent Receipts of Principal",
                 [("Date Received", "date"), ("Description", "description"),
                  ("Inventory Value", "amount")])

    add_schedule("A-1", "Schedule A-1",
                 "Statement of Increases on Sales, Liquidation or Distribution",
                 [("Description", "description"),
                  ("Proceeds", "amount"), ("Inventory Value", "inventory_value")])

    add_schedule("A-2", "Schedule A-2", "Statement of All Income Collected",
                 [("Date", "date"), ("Description", "description"),
                  ("Institution", "institution"), ("Amount", "amount")])

    add_schedule("B", "Schedule B",
                 "Statement of Decreases Due to Sales, Liquidation, Collection, Distribution, or Uncollectibility",
                 [("Date", "date"), ("Description", "description"),
                  ("Proceeds", "amount"), ("Inventory Value", "inventory_value")])

    add_schedule("C", "Schedule C",
                 "Statement of Funeral and Administration Expenses and Taxes",
                 [("Date", "date"), ("Description", "description"), ("Amount", "amount")])

    add_schedule("C-1", "Schedule C-1", "Statement of Unpaid Administration Expenses",
                 [("Description", "description"), ("Amount", "amount")])

    add_schedule("D", "Schedule D", "Statement of All Creditors' Claims",
                 [("Description", "description"), ("Amount", "amount")])

    add_schedule("E", "Schedule E", "Distributions",
                 [("Description", "description"), ("Distribution Value", "amount")])

    add_schedule("F", "Schedule F",
                 "Statement of New Investments, Exchanges and Stock Distributions",
                 [("Date", "date"), ("Description", "description"),
                  ("Shares", "shares"), ("Inventory Value", "amount")])

    add_schedule("G", "Schedule G", "Balance On Hand",
                 [("Description", "description"), ("Shares", "shares"),
                  ("Market Value", "market_value"), ("Inventory Value", "inventory_value")])

    # ── SCHEDULE H — Interested Parties ───────────────────────────────────────
    estate_header = f"Estate of {dec}"
    if aka:
        estate_header += f", aka {aka}"
    add_para(estate_header, bold=True, size=11,
             alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    add_para("Schedule H", bold=True, size=11,
             alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    add_para("Statement of Interested Parties", bold=True, size=11,
             alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=12)

    h_entries = by_sched.get("H", [])
    if h_entries:
        table = doc.add_table(rows=1, cols=3)
        table.style = "Table Grid"
        for i, hdr in enumerate(["Name and Post Office Address", "Relationship", "Nature of Interest"]):
            cell = table.rows[0].cells[i]
            cell.text = hdr
            for p in cell.paragraphs:
                for r in p.runs:
                    r.bold = True
                    r.font.size = Pt(10)
                    r.font.name = "Times New Roman"
        for e in h_entries:
            row = table.add_row()
            row.cells[0].text = f"{e.get('description', '')}\n{e.get('institution', '')}"
            row.cells[1].text = e.get("category", "")
            row.cells[2].text = "Distributee"
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        r.font.size = Pt(10)
                        r.font.name = "Times New Roman"
    else:
        add_para("None", size=11)

    add_para(
        "\nThe records of this Court have been searched for powers of attorney and "
        "assignments and encumbrances made and executed by any of the persons interested "
        "in or entitled to share in the estate. No such powers of attorney, assignments "
        "or encumbrances were found to have been filed or recorded in this Court, and the "
        "accounting party has no knowledge of the execution of any such power of attorney, "
        "assignment or encumbrance that is not so filed and recorded.",
        size=10, space_after=6
    )

    doc.add_page_break()

    # ── SCHEDULE I — Commission Computation ───────────────────────────────────
    add_para(estate_header, bold=True, size=11,
             alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    add_para("Schedule I", bold=True, size=11,
             alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    add_para("Statement of Computation of Commissions", bold=True, size=11,
             alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=12)

    # Receiving commission — use net equity (gross minus liens/mortgages)
    total_liens = sum(float(e.get("lien_amount", 0) or 0)
                      for e in by_sched.get("A", []) + by_sched.get("AA", []))
    net_principal = tot_a + tot_aa - total_liens

    add_para("For Receiving Principal", bold=True, size=11, space_after=4)
    recv_base = net_principal + tot_a1 + tot_a2 + unreal_inc
    add_para(f"Principal Received (Schedule A + AA)     {money(tot_a + tot_aa)}", size=11, space_after=1)
    if total_liens > 0:
        add_para(f"Less: Liens/Mortgages on Real Property  ({money(total_liens)})", size=11, space_after=1)
        add_para(f"Net Equity                              {money(net_principal)}", size=11, space_after=1)
    add_para(f"Increases on Principal (Schedule A-1)    {money(tot_a1)}", size=11, space_after=1)
    add_para(f"Income Collected (Schedule A-2)          {money(tot_a2)}", size=11, space_after=1)
    add_para(f"Unrealized Increases (Schedule G)        {money(unreal_inc)}", size=11, space_after=4)
    add_para(f"Commission Base                          {money(recv_base)}", bold=True, size=11, space_after=6)

    tiers = [
        (0.05, 100000), (0.04, 200000), (0.03, 700000), (0.025, float('inf')),
    ]
    remaining = recv_base
    recv_comm = 0
    for rate, bracket in tiers:
        if remaining <= 0:
            break
        base = min(remaining, bracket)
        comm = base * rate
        recv_comm += comm
        pct = int(rate * 100) if rate * 100 == int(rate * 100) else rate * 100
        add_para(f"  {pct}% on {money(base):>20s} = {money(comm):>15s}", size=11, space_after=1)
        remaining -= base

    recv_half = recv_comm / 2
    add_para(f"\n1/2 Thereof for Receiving               {money(recv_half)}",
             bold=True, size=11, space_after=8)

    # Paying commission
    paying_base = tot_c + tot_d + tot_e + tot_g
    add_para("For Paying Principal", bold=True, size=11, space_after=4)
    add_para(f"Funeral and Administration Expenses (Schedule C)  {money(tot_c)}", size=11, space_after=1)
    add_para(f"Payment of Debts (Schedule D)                     {money(tot_d)}", size=11, space_after=1)
    add_para(f"Distributions of Principal (Schedule E)           {money(tot_e)}", size=11, space_after=1)
    add_para(f"Principal on Hand (Schedule G)                    {money(tot_g)}", size=11, space_after=4)
    add_para(f"Total Principal                                   {money(paying_base)}",
             bold=True, size=11, space_after=6)

    remaining = paying_base
    pay_comm = 0
    for rate, bracket in tiers:
        if remaining <= 0:
            break
        base = min(remaining, bracket)
        comm = base * rate
        pay_comm += comm
        pct = int(rate * 100) if rate * 100 == int(rate * 100) else rate * 100
        add_para(f"  {pct}% on {money(base):>20s} = {money(comm):>15s}", size=11, space_after=1)
        remaining -= base

    pay_half = pay_comm / 2
    add_para(f"\n1/2 Thereof for Paying                  {money(pay_half)}",
             bold=True, size=11, space_after=12)

    total_comm = recv_half + pay_half
    add_para(f"Total Commissions Due Each {role}", bold=True, size=11, space_after=4)
    add_para(f"  Receiving     {money(recv_half)}", size=11, space_after=1)
    add_para(f"  Paying        {money(pay_half)}", size=11, space_after=4)
    add_para(f"  Total         {money(total_comm)}", bold=True, size=11, space_after=6)
    add_para(f"\nTotal commissions available for allocation:   {money(total_comm)}",
             bold=True, size=11, space_after=6)

    doc.add_page_break()

    # ── SCHEDULE J — Cash Reconciliation ──────────────────────────────────────
    add_para(estate_header, bold=True, size=11,
             alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    add_para("Schedule J", bold=True, size=11,
             alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    add_para("Statement of Other Pertinent Facts and Cash Reconciliation", bold=True,
             size=11, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=12)

    add_para("Other Pertinent Facts", bold=True, size=11, space_after=4)
    add_para("None", size=11, space_after=8)

    add_para("Reconciliation of Cash and Other Assets", bold=True, size=11, space_after=8)

    recon_items = [
        ("Schedule A", "Receipts", tot_a, "CREDITS"),
        ("Schedule AA", "Subsequent Receipts", tot_aa, "CREDITS"),
        ("Schedule A-2", "Income Collected", tot_a2, "CREDITS"),
        ("Schedule B", "Proceeds on Sales, Etc.", tot_b, "DEBITS"),
        ("Schedule C", "Admin/Funeral Expenses", tot_c, "DEBITS"),
        ("Schedule F", "Purchases, Etc.", sched_total("F"), "DEBITS"),
        ("Schedule G", "On Hand", tot_g, "DEBITS"),
    ]

    add_para(f"{'':30s} {'DEBITS':>15s} {'CREDITS':>15s}", bold=True, size=11, space_after=4)
    cash_debits = 0
    cash_credits = 0
    for label, desc, val, side in recon_items:
        debit = money(val) if side == "DEBITS" else ""
        credit = money(val) if side == "CREDITS" else ""
        if side == "DEBITS":
            cash_debits += val
        else:
            cash_credits += val
        add_para(f"{label:12s} {desc:18s} {debit:>15s} {credit:>15s}", size=11, space_after=1)

    add_para(f"\n{'Total':30s} {money(cash_debits):>15s} {money(cash_credits):>15s}",
             bold=True, size=11, space_after=6)

    doc.add_page_break()

    # ── SCHEDULE K — Estate Taxes ─────────────────────────────────────────────
    add_schedule("K", "Schedule K",
                 "Statement of Estate Taxes Paid and Allocation Thereof",
                 [("Description", "description"), ("Amount", "amount")])

    return make_docx_bytes(doc)
