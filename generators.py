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
from docx.enum.text import WD_ALIGN_PARAGRAPH
from pypdf import PdfReader, PdfWriter

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


def today():
    return datetime.now().strftime("%B %d, %Y")


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

    # Date
    p = doc.add_paragraph(today())
    p.paragraph_format.space_after = Pt(12)

    # Addressee
    doc.add_paragraph(f"Surrogate's Court, {county} County")
    doc.add_paragraph(f"Attn: {dept}")
    doc.add_paragraph(address)
    doc.add_paragraph(city_state_zip)
    doc.add_paragraph("")

    # RE line
    re_p = doc.add_paragraph(f"RE: Estate of {decedent}")
    re_p.paragraph_format.space_after = Pt(12)

    doc.add_paragraph("Greetings,")
    doc.add_paragraph("")

    proc_word = proceeding.lower()
    body = doc.add_paragraph(
        f"Our office efiled the above referenced petition for {proc_word} on {efile_date}. "
        f"Please find enclosed the following documents:"
    )
    body.paragraph_format.space_after = Pt(6)

    # Enclosures as bullet list
    for enc in enclosures:
        p = doc.add_paragraph(style="List Bullet")
        p.text = enc

    doc.add_paragraph("")
    doc.add_paragraph("Please do not hesitate to call our office if you have concerns and questions.")
    doc.add_paragraph("")
    doc.add_paragraph("Sincerely,")
    doc.add_paragraph("")
    doc.add_paragraph("")
    doc.add_paragraph(signer)
    doc.add_paragraph("Enc.")

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

    def nonzero(v):
        """Return v only if it's a non-empty, non-zero value."""
        s = str(v or "").strip()
        return s if s and s not in ("0", "0.0", "0.00") else ""

    # ── Header ────────────────────────────────────────────────────────────────
    line(f"SURROGATE\u2019S COURT \u2014 {county.upper()} COUNTY",
         bold=True, center=True, space_after=2)

    # Two-column row using a borderless table
    hdr_tbl = doc.add_table(rows=2, cols=2)
    hdr_tbl.style = "Table Grid"
    # Remove all borders
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    def _no_border(cell):
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        tcBorders = OxmlElement("w:tcBorders")
        for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
            el = OxmlElement(f"w:{side}")
            el.set(qn("w:val"), "none")
            tcBorders.append(el)
        tcPr.append(tcBorders)

    for row in hdr_tbl.rows:
        for cell in row.cells:
            _no_border(cell)

    def _cell(row_i, col_i, text, bold=False):
        cell = hdr_tbl.rows[row_i].cells[col_i]
        p = cell.paragraphs[0]
        _run(p, text, bold=bold)

    _cell(0, 0, "Administration Proceeding", bold=True)
    _cell(0, 1, "AFFIDAVIT OF ASSETS & LIABILITIES", bold=True)
    _cell(1, 0, f"Estate of {decedent}", bold=True)
    _cell(1, 1, "(To dispense with filing of bond / SCPA 805)")

    if aka:
        line(f"a/k/a {aka}")
    line("")  # blank

    # Deceased / File No
    p = line()
    _run(p, "Deceased", bold=True)
    if file_no:
        _run(p, f"\t\t\t\t\tFile No. {file_no}")

    blank()

    # ── Venue block ───────────────────────────────────────────────────────────
    line("STATE OF NEW YORK\t\t\t\t)")
    line("\t\t\t\t\t\t) ss:")
    line(f"COUNTY OF {county}\t\t\t\t)")
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
    doc = Document(os.path.join(WORD_TEMPLATES_DIR, "Affidavit_of_Heirship_Full_Admin.docx"))
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

    replace_in_doc(doc, {
        "COUNTY OF _____________": f"COUNTY OF {county}",
        "___________________\t\t\t\t\tAFFIDAVIT OF HEIRSHIP": f"{decedent}\t\t\t\t\tAFFIDAVIT OF HEIRSHIP",
        "A/K/A ___________________": f"A/K/A {data.get('decedentAKA', '')}",
        "COUNTY OF \t\t\t)": f"COUNTY OF {county}\t\t\t)",
        "\tI, ______________, being duly sworn, deposes and says:": f"\tI, {deponent}, being duly sworn, deposes and says:",
        "I reside at _________________________.  I am over the age of eighteen (18) years and I am fully familiar with the facts and circumstances herein, the decedent's family tree, as I am the ______________of the Decedent and have known the Decedent for over _____ years.":
            f"I reside at {deponent_address}.  I am over the age of eighteen (18) years and I am fully familiar with the facts and circumstances herein, the decedent's family tree, as I am the {deponent_rel} of the Decedent and have known the Decedent for over {years_known} years.",
        "The Decedent was born on ___________ and died on __________________.": f"The Decedent was born on {dob} and died on {dod}.",
        "Mother: ": f"Mother: {mother_name}",
        "Father: ": f"Father: {father_name}",
        f"Therefore, ______________ is the sole distributee of the Estate of ______________":
            f"Therefore, {sole_distributee} is the sole distributee of the Estate of {decedent}",
        f"This affidavit is made with my personal knowledge knowing the ______________ County Surrogate's Court will rely thereon in issuing Letters of Administration to _________________, the petitioner.":
            f"This affidavit is made with my personal knowledge knowing the {county} County Surrogate's Court will rely thereon in issuing Letters of Administration to {petitioner}, the petitioner.",
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


# ─── PROBATE PDF (P-1 + OATH + WITNESS) ──────────────────────────────────────

def _extract_pdf_pages(writer, page_indices):
    """Extract specific page indices from a PdfWriter into new PDF bytes."""
    out = PdfWriter()
    for idx in page_indices:
        out.add_page(writer.pages[idx])
    buf = io.BytesIO()
    out.write(buf)
    buf.seek(0)
    return buf.read()


def _build_probate_writer(data):
    """Fill all fields in Probate-_NY_Court_Forms.pdf; return (writer, reader)."""
    reader = PdfReader(os.path.join(PDFS_DIR, "Probate-_NY_Court_Forms.pdf"))
    writer = PdfWriter()
    writer.clone_reader_document_root(reader)

    county   = data.get("county", "")
    dec      = decedent_full(data)
    pet      = petitioner_full(data)
    lt       = data.get("lettersType", "")
    witnesses = ", ".join(filter(None, [data.get("witness1", ""), data.get("witness2", "")]))
    pet_addr  = ", ".join(filter(None, [
        data.get("petitionerStreet", ""), data.get("petitionerCity", ""),
        data.get("petitionerState", ""), data.get("petitionerZip", ""),
    ]))

    fields = {
        # ── Petition (pages 1-4) ────────────────────────────────────────────────
        "COUNTY OF": county,
        "To the Surrogates Court County of": county,
        "WILL OF": dec,
        "a Name": dec,
        "aka": data.get("decedentAKA", ""),
        "Name": pet,
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
        "Letters Testamentary to 1": data.get("lettersTo", ""),
        "Letters Testamentary to 2": data.get("lettersTo", ""),
        "Letters of Administration cta to": data.get("lettersTo", ""),
        "Print Name": pet,

        # ── Oath and Designation (page 5) ───────────────────────────────────────
        "COUNTY OF_2": county,
        "OATH OF": pet,
        "Surrogates Court of": county,
        "My domicile is": pet_addr,
        "Street Address": data.get("petitionerStreet", ""),
        "Print Name_3": pet,
        "Signature of Attorney": data.get("attorneyName", ""),
        "Print Name_4": data.get("attorneyName", ""),
        "Firm Name": data.get("firmName", ""),
        "Tel No": data.get("attorneyPhone", ""),
        "Email": data.get("attorneyEmail", ""),
        "Address of Attorney": data.get("firmAddress", ""),

        # ── Attesting Witness (page 10) ─────────────────────────────────────────
        "COUNTY OF_7": county,
        "WILL OF 1": data.get("decedentFirstName", ""),
        "WILL OF 2": data.get("decedentLastName", ""),
        "aka 1": data.get("decedentAKA", ""),
        "STATE OF NEW YORK_5": "New York",
        "COUNTY OF_8": county,
    }

    # Letters type checkboxes
    if "Testamentary" in lt:
        fields["Letters Testamentary"] = "X"
        fields["EXECUTOR"] = "X"          # oath page 5
    elif "Trusteeship" in lt:
        fields["Letters of Trusteeship"] = "X"
    elif "c.t.a" in lt:
        fields["Letters of Administration cta"] = "X"
        fields["ADMINISTRATOR cta"] = "X"  # oath page 5
    elif "Temporary" in lt:
        fields["Temporary Administration"] = "X"

    # Petitioner interest
    if "Executor" in data.get("petitionerInterest", ""):
        fields["Executor s named in decedents Will"] = "X"
    if data.get("petitionerIsAttorney") == "Yes":
        fields["is"] = "X"
    else:
        fields["is not an attorney"] = "X"

    # Surviving relatives
    surv_map = {
        "survivingSpouse": "Spouse husbandwife",
        "survivingChildren": "Child or children andor issue of predeceased child or children",
        "survivingParents": "MotherFather",
        "survivingSiblings": "Sisters andor brothers either of the whole or half blood and issue of predeceased sisters",
        "survivingGrandparents": "Grandparents Include maternal and paternal",
        "survivingAuntsUncles": "Aunts andor uncles and children of predeceased aunts andor uncles first cousins",
        "survivingFirstCousinsOnceRemoved": "First cousins once removed children of predeceased first cousins Include maternal and",
    }
    for key, field in surv_map.items():
        if data.get(key):
            fields[field] = "X"

    # Distributees
    name_f = ["1_2", "3", "4", "5", "6", "7"]
    addr_f = ["2_2", "3_2", "4_2", "5_2", "6_2", "7_2"]
    int_f = [f"Interest or Nature of Fiduciary Status {i}" for i in range(1, 8)]
    for i, dist in enumerate(data.get("distributees", [])[:7]):
        if dist.get("name"):
            fields[name_f[i]] = f"{dist['name']} ({dist.get('relationship','')})"
            fields[addr_f[i]] = f"{dist.get('address','')} | {dist.get('citizenship','')}"
            fields[int_f[i]] = dist.get("relationship", "Distributee")

    # Apply all fields across all pages
    all_fields = reader.get_fields() or {}
    for fid, val in fields.items():
        if fid in all_fields and val:
            for page in writer.pages:
                try:
                    writer.update_page_form_field_values(page, {fid: val})
                except Exception:
                    pass

    return writer, reader


def generate_probate_docs(data):
    """
    Returns list of (filename, bytes) for the full probate packet:
      - P-1 Petition (pages 1-4)
      - Combined Verification, Oath and Designation (page 5)
      - Affidavit of Attesting Witness (page 10) — omitted if self-proving will
    Fills the source PDF only once for efficiency.
    """
    writer, _ = _build_probate_writer(data)
    last = data.get("decedentLastName", "estate").replace(" ", "_")
    docs = [
        (f"02_Petition_P1_{last}.pdf",        _extract_pdf_pages(writer, [0, 1, 2, 3])),
        (f"03_Oath_Designation_{last}.pdf",   _extract_pdf_pages(writer, [4])),
    ]
    if not data.get("selfProvingAffidavit"):
        docs.append(
            (f"04_Affidavit_Attesting_Witness_{last}.pdf", _extract_pdf_pages(writer, [9]))
        )
    return docs


# Keep for backward compatibility (ancillary PDF still uses its own function)
def fill_probate_pdf(data):
    writer, _ = _build_probate_writer(data)
    return _extract_pdf_pages(writer, [0, 1, 2, 3])


# ─── ANCILLARY ADMIN PDF (AA-1) ───────────────────────────────────────────────

def fill_ancillary_pdf(data):
    reader = PdfReader(os.path.join(PDFS_DIR, "admin_ancil.pdf"))
    writer = PdfWriter()
    writer.clone_reader_document_root(reader)

    dec = decedent_full(data)
    pet = petitioner_full(data)
    letters_to = data.get("lettersTo", "") or pet
    county = data.get("county", "")
    foreign_state = data.get("foreignState", "")

    def v(key, default=""):
        """Get value or default."""
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

    fields = {
        # ── PAGE 1 ────────────────────────────────────────────────
        # Header
        "Text Field 8":  county,                        # COUNTY OF
        "Text Field 9":  dec,                           # ESTATE OF
        "Text Field 10": v("decedentAKA"),              # a/k/a
        "Text Field 11": foreign_state,                 # domiciliary of State of
        "Text Field 12": v("fileNo"),                   # File No.
        "Text Field 13": county,                        # To the Surrogate's Court, County of

        # Para 1 — Petitioner
        "Text Field 14": pet,                           # Name
        "Text Field 15": v("petitionerStreet"),         # Street
        "Text Field 16": v("petitionerCity"),           # City/Village/Town
        "Text Field 17": v("petitionerState"),          # State
        "Text Field 18": v("petitionerZip"),            # Zip
        "Text Field 19": v("petitionerCitizenship", "U.S.A."),  # Citizen of

        # Para 1 — Interest of petitioner (other field for "Other/Designee" text)
        "Text Field 27": v("petitionerRelationship"),   # relationship if distributee
        "Text Field 28": v("petitionerInterest"),       # Other/specify

        # Para 2 — Decedent
        "Text Field 29": dec,                           # (a) Name
        "Text Field 30": v("decedentDOD"),              # (b) Date of Death
        "Text Field 31": v("decedentPlaceOfDeath"),     # (c) Place of death
        "Text Field 32": v("decedentStreet"),           # (d) Street
        "Text Field 33": v("decedentCity"),             # City/Town/Village
        "Text Field 34": v("decedentCounty"),           # County
        "Text Field 35": foreign_state,                 # State (foreign domicile)
        "Text Field 36": v("decedentZip"),              # Zip
        "Text Field 37": v("decedentCitizenship", "U.S.A."),  # (e) Citizen of

        # ── PAGE 2 ────────────────────────────────────────────────
        # Para 3 — Foreign letters info
        "Text Field 38": v("foreignLettersDate"),       # date letters issued
        "Text Field 39": v("foreignLettersIssuedTo", letters_to),  # issued to
        "Text Field 40": v("foreignCourtName"),         # by [Court name]
        "Text Field 41": foreign_state,                 # State of
        "Text Field 42": v("foreignBondAmount", "0"),   # security/bond amount $

        # Para 4(a) — NY property values
        "Text Field 43": v("personalPropertyValue", "0.00"),      # Personal Property $
        "Text Field 44": v("improvedRealProperty", "0.00"),       # Improved real property $
        "Text Field 45": v("unimprovedRealProperty", "0.00"),     # Unimproved real property $
        "Text Field 46": v("grossRents18mo", "0.00"),             # Gross rents 18mo $
        "Text Field 47": total_str,                               # Total $

        # Para 4(b) — Other assets
        "Text Field 48": v("otherAssets", "NONE"),      # [NONE or specify] line 1
        "Text Field 49": "",                            # line 2

        # Para 5 — NY Dept of Tax always required, others if applicable
        "Text Field 50": "N/A",                         # Amount of claim for Dept of Tax

        # ── PAGE 3 ────────────────────────────────────────────────
        # Para 6(a) — Distributees of full age (3 rows: name / address / interest)
        # Rows at y≈715, 699, 684 → field sets (57,58,59), (60,61,62), (63,64,65)

        # Para 7 — no other persons / no previous application (boilerplate, no fill)

        # WHEREFORE clause
        # "Ancillary Letters of Administration to:" → Text Field 75 (y=503, wide)
        # "Ancillary Letters of Administration d.b.n. to:" → Text Field 1065 (y=443)
        # Para (d) limitation → Text Field 77 (y=441) — leave blank
        # Para (e) limitation → Text Field 79 (y=410) — NONE
        # Para (f) other relief → Text Field 1067 (y=412) — leave blank
        "Text Field 75":   letters_to,                  # Ancillary Letters of Admin to: [NAME]
        "Text Field 1065": "",                          # d.b.n. to (leave blank unless needed)
        "Text Field 77":   "",                          # para (d) no limitation
        "Text Field 79":   "NONE",                      # para (e) -> NONE
        "Text Field 80":   today(),                     # Dated

        # ── PAGE 4 — Combined Verification, Oath and Designation ──────────────
        "Text Field 85": v("petitionerState", "New York"),  # STATE OF
        "Text Field 87": county,                        # COUNTY OF
        "Text Field 89": county,                        # designate Clerk of ... County
        "Text Field 91": petitioner_address,            # My domicile is
        "Text Field 97": pet,                           # Print Name (petitioner)
    }

    # Para 6(a) distributees — 3 rows
    dist_rows = [
        ("Text Field 57", "Text Field 58", "Text Field 59"),   # name, address, interest
        ("Text Field 60", "Text Field 61", "Text Field 62"),
        ("Text Field 63", "Text Field 64", "Text Field 65"),
    ]
    for i, dist in enumerate(data.get("distributees", [])[:3]):
        if dist.get("name"):
            nf, af, rf = dist_rows[i]
            fields[nf] = dist["name"]
            fields[af] = dist.get("address", "")
            fields[rf] = dist.get("relationship", "")

    # Build page->field map so we update each field on the correct page
    field_page_map = {}
    for page_num, page in enumerate(reader.pages):
        annots = page.get('/Annots', [])
        for annot in annots:
            try:
                obj = annot.get_object()
                name = obj.get('/T', '')
                if name:
                    field_page_map[name] = page_num
            except Exception:
                pass

    # Group fields by page and fill each page once
    from collections import defaultdict
    by_page = defaultdict(dict)
    all_field_names = set(reader.get_fields().keys()) if reader.get_fields() else set()
    for fid, val in fields.items():
        if fid in all_field_names and val and fid in field_page_map:
            by_page[field_page_map[fid]][fid] = val

    for page_num, page_fields in by_page.items():
        try:
            writer.update_page_form_field_values(writer.pages[page_num], page_fields)
        except Exception:
            pass

    buf = io.BytesIO()
    writer.write(buf)
    buf.seek(0)
    return buf.read()


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
    """Fill the A-1 Administration Petition + Oath PDF form."""
    from collections import defaultdict

    reader = PdfReader(os.path.join(ADMIN_TEMPLATES_DIR, "Admin Petition + Oath.pdf"))
    writer = PdfWriter()
    writer.clone_reader_document_root(reader)

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
    # Instructions: "No" for all prior classes, number/Yes for surviving classes, "X" for all subsequent
    surv_keys = [
        "survivingSpouse", "survivingChildren", "survivingIssue",
        "survivingParents", "survivingSiblings", "survivingGrandparents",
        "survivingAuntsUncles", "survivingFirstCousinsOnceRemoved",
    ]
    surv_vals = [bool(data.get(k)) for k in surv_keys]
    first_surv = next((i for i, s in enumerate(surv_vals) if s), None)
    last_surv  = (len(surv_vals) - 1 - next(
                     (i for i, s in enumerate(reversed(surv_vals)) if s), -1)
                 ) if first_surv is not None else None

    dropdown_vals = []
    for i, surv in enumerate(surv_vals):
        if first_surv is None:
            dropdown_vals.append("-")
        elif i < first_surv:
            dropdown_vals.append("No")
        elif last_surv is not None and i > last_surv:
            dropdown_vals.append("X")
        else:
            dropdown_vals.append("Yes" if surv else "No")

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
        "COUNTY OF":                        county,
        "Estate of 1":                      dec,
        "aka":                              v("decedentAKA"),
        "File No":                          v("fileNo"),
        "TO THE SURROGATES COURT COUNTY OF": county,

        # Caption checkboxes (letters type)
        "petition for letters of admin":    "/Yes" if is_standard   else "/Off",
        "limited admin":                    "/Yes" if is_limited    else "/Off",
        "limited admin with lim":           "/Yes" if is_limitation else "/Off",
        "temp admin":                       "/Yes" if is_temporary  else "/Off",

        # ── PAGE 1: Petitioner ───────────────────────────────────────
        "Name":                             pet,
        "Domicile":                         v("petitionerStreet"),
        "County":                           v("petitionerCity"),
        "State":                            v("petitionerState"),
        "Zip":                              v("petitionerZip"),
        "Telephone Number":                 v("petitionerPhone"),
        "yes us citizen":                   "/Yes" if pet_us  else "/Off",
        "NO us citizen":                    "/Yes" if not pet_us else "/Off",
        "Distributee of decedent state relationship": v("petitionerRelationship"),
        "Otherspecify":                     v("petitionerInterest"),
        "yes attorney":                     "/Yes" if is_attorney     else "/Off",
        "NO not an attorney":               "/Yes" if not is_attorney else "/Off",
        "not a convicted felon":            "/Yes",

        # ── PAGE 1: Decedent ─────────────────────────────────────────
        "Name_2":                           dec,
        "Domicile_2":                       v("decedentStreet"),
        "State_2":                          v("decedentState"),
        "Zip Code":                         v("decedentZip"),
        "Township of":                      v("decedentCounty", v("decedentCity")),
        "Date of Death":                    v("decedentDOD"),
        "Place of Death":                   v("decedentPlaceOfDeath"),
        "yes us citizen 1":                 "/Yes" if dec_us  else "/Off",
        "NO not US Citizen 2":              "/Yes" if not dec_us else "/Off",

        # ── PAGE 2: Property values ──────────────────────────────────
        "undefined_4":                      v("personalPropertyValue", "0"),
        "undefined_5":                      v("realPropertyValue", "0"),
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
        "a-process issue letters":          "/Yes",
        "c a decree award letters of":      "/Yes",
        "9c1":                              "/Yes" if is_standard   else "/Off",
        "9c2":                              "/Yes" if is_limited    else "/Off",
        "9c3":                              "/Yes" if is_limitation else "/Off",
        "9c4":                              "/Yes" if is_temporary  else "/Off",
        "Administration to":                letters_to if is_standard   else "",
        "Limited Administration to":        letters_to if is_limited    else "",
        "Administration with Limitation to": letters_to if is_limitation else "",
        "Temporary Administration to":      letters_to if is_temporary  else "",
        "Dated":                            today(),
        "Print Name":                       pet,

        # ── PAGE 5: Combined Verification, Oath & Designation ────────
        "ss":                               v("petitionerState", "New York"),
        "My domicile is":                   pet_addr,
        "before me personally came":        pet,
        "Print Name_3":                     v("attorneyName"),
        "Firm Name":                        v("attorneyFirm"),
        "TelNo":                            v("attorneyPhone"),
        "Address of Attorney":              v("attorneyAddress"),
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

    # Fill page-by-page using the same pattern as fill_ancillary_pdf
    field_page_map = {}
    for page_num, page in enumerate(reader.pages):
        for annot in (page.get('/Annots') or []):
            try:
                obj = annot.get_object()
                name = obj.get('/T', '')
                if name:
                    field_page_map[str(name)] = page_num
            except Exception:
                pass

    all_field_names = set(reader.get_fields().keys()) if reader.get_fields() else set()
    by_page = defaultdict(dict)
    for fid, val in fields.items():
        if fid in all_field_names and fid in field_page_map and val not in ("", "/Off"):
            by_page[field_page_map[fid]][fid] = val

    for page_num, page_fields in by_page.items():
        try:
            writer.update_page_form_field_values(writer.pages[page_num], page_fields)
        except Exception:
            pass

    buf = io.BytesIO()
    writer.write(buf)
    buf.seek(0)
    return buf.read()


# ─── FAMILY TREE WORKSHEET (FT-1) ─────────────────────────────────────────────

def fill_ft1_pdf(data):
    """Fill the actual FT-1 Family Tree Affidavit court form PDF."""
    reader = PdfReader(os.path.join(ADMIN_TEMPLATES_DIR, "Family_Tree_Affidavit_Fill-In.pdf"))
    writer = PdfWriter()
    writer.clone_reader_document_root(reader)

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
    fields["128"]         = dec_name          # Estate of
    fields["230"]         = aka               # a/k/a
    fields["331"]         = dec_name          # repeated on "Deceased" line
    fields["412"]         = file_no           # File No.
    fields["Combo Box00"] = "LETTERS OF ADMINISTRATION"

    # ── Deponent (petitioner) ───────────────────────────────────────────────────
    fields["5a5"] = pet_name   # I, ___
    fields["5b6"] = pet_addr   # reside at
    fields["5c7"] = pet_rel    # relationship to decedent

    # ── Section 1a: Marriages ───────────────────────────────────────────────────
    if marital == "never_married":
        fields["Check Box01h"] = "/Yes"
    elif marital == "married" and spouse_name:
        fields["6a9"] = spouse_name            # surviving spouse
    elif marital == "divorced" and spouse_name:
        fields["6b10"] = spouse_name           # ex-spouse name
        fields["Check Box01a"] = "/Yes"        # divorced
        if divorce_yr:
            fields["6a9"] = f"divorced {divorce_yr}"
    elif marital == "widowed" and spouse_name:
        # Spouse predeceased — list as ex-spouse who died while married
        fields["6b10"] = spouse_name
        fields["Check Box01b"] = "/Yes"        # died while married to decedent

    # ── Section 1b: Children (6 slots) ─────────────────────────────────────────
    child_name_f = ["816",  "917",  "1018",  "1119",  "1220",  "1321"]
    child_dod_f  = ["8a22", "9a23", "10a24", "11a25", "12a26", "13a27"]
    for i, c in enumerate(children[:6]):
        if c.get("name"):
            fields[child_name_f[i]] = c["name"]

    # ── Section 2: Parents (page 2, fields 25/26) ───────────────────────────────
    # Not in our data model — left blank for manual completion

    # ── Section 3a: Siblings (6 slots, page 2) ─────────────────────────────────
    sib_name_f = ["27", "28", "29", "30", "31", "32"]
    sib_dod_f  = ["27a","28a","29a","30a","31a","32a"]
    for i, s in enumerate(siblings[:6]):
        if s.get("name"):
            fields[sib_name_f[i]] = s["name"]

    # ── Section 3b: Nieces/Nephews (7 slots, page 2) ───────────────────────────
    # fields: name / child-of / DOD
    nie_name_f = ["33","34","35","36","37","38","39"]
    nie_cof_f  = ["33a","34a","35a","36a","37a","38a","39a"]
    nie_dod_f  = ["33b","34b","35b","36b","37b","38b","39b"]
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

    # ── Apply fields across all pages ───────────────────────────────────────────
    all_form_fields = reader.get_fields() or {}
    for fid, val in fields.items():
        if fid in all_form_fields and val:
            for page in writer.pages:
                try:
                    writer.update_page_form_field_values(page, {fid: val})
                except Exception:
                    pass

    buf = io.BytesIO()
    writer.write(buf)
    buf.seek(0)
    return buf.read()


def generate_ft1(data):
    return fill_ft1_pdf(data)
