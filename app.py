"""
NY Surrogate's Court Probate Assistant v2
Full document packet generator

Run: python3 app.py
Open: http://localhost:8080
"""

import io
import os
import traceback
import zipfile
from flask import Flask, request, jsonify, send_file, send_from_directory
from generators import (
    generate_cover_letter, generate_805, generate_heirship,
    generate_waiver_cover, generate_attorney_cert,
    fill_probate_pdf, fill_ancillary_pdf,
    fill_administration_doc, generate_ft1,
    COUNTY_INFO, today, decedent_full, petitioner_full
)

app = Flask(__name__, static_folder="static")


def make_zip(files):
    """files = list of (filename, bytes)"""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in files:
            zf.writestr(name, data)
    buf.seek(0)
    return buf.read()


@app.route("/")
def index():
    return send_from_directory("static", "index.html")


@app.route("/counties")
def counties():
    return jsonify(list(COUNTY_INFO.keys()))


@app.route("/generate-packet", methods=["POST"])
def generate_packet():
    data = request.get_json()
    if not data:
        return jsonify({"error": "No data"}), 400

    proceeding = data.get("proceedingType", "Probate")
    last_name = data.get("decedentLastName", "estate").replace(" ", "_")
    non_probate = proceeding in ("Administration", "Ancillary", "NonDomiciliary")
    files = []
    errors = []

    print(f"\n[PACKET] proceeding={proceeding!r}  decedent={last_name!r}")

    # 1. Court cover letter (always)
    try:
        files.append((f"01_Cover_Letter_{last_name}.docx", generate_cover_letter(data)))
        print("[OK] 01 Cover letter")
    except Exception as e:
        print(f"[ERR] 01 Cover letter: {e}")
        traceback.print_exc()
        errors.append(f"Cover letter: {e}")
        files.append((f"01_MISSING_Cover_Letter.txt",
                       f"FAILED TO GENERATE\n\nError: {e}".encode()))

    # 2. Main petition
    try:
        if proceeding == "Probate":
            files.append((f"02_Petition_Probate_P1_{last_name}.pdf", fill_probate_pdf(data)))
            print("[OK] 02 Probate petition (P-1)")
        elif proceeding == "Ancillary":
            files.append((f"02_Petition_Ancillary_AA1_{last_name}.pdf", fill_ancillary_pdf(data)))
            print("[OK] 02 Ancillary petition (AA-1)")
        elif proceeding == "Administration":
            files.append((f"02_Petition_Administration_A1_{last_name}.docx",
                           fill_administration_doc(data)))
            print("[OK] 02 Administration petition (A-1)")
        elif proceeding == "NonDomiciliary":
            files.append((f"02_NOTE_NonDomiciliary_A1.txt",
                f"PETITION FOR NON-DOMICILIARY LETTERS OF ADMINISTRATION\n\n"
                f"Use the A-1 form. Title: 'Petition for Non-Domiciliary Letters of Administration'\n\n"
                f"Decedent domicile (out of state):\n"
                f"  {data.get('decedentStreet','')}\n"
                f"  {data.get('decedentCity','')}, {data.get('foreignState','')} {data.get('decedentZip','')}\n\n"
                f"Petitioner: {data.get('petitionerFirstName','')} {data.get('petitionerLastName','')}\n"
                f"  {data.get('petitionerStreet','')}, {data.get('petitionerCity','')}, NY {data.get('petitionerZip','')}\n\n"
                f"NY Property:\n  {data.get('realPropertyDescription','(see estate details)')}".encode()))
            print("[OK] 02 NonDomiciliary note")
        else:
            print(f"[WARN] 02 No petition generated — unrecognised proceedingType={proceeding!r}")
    except Exception as e:
        print(f"[ERR] 02 Main petition: {e}")
        traceback.print_exc()
        errors.append(f"Main petition: {e}")
        files.append((f"02_MISSING_Main_Petition.txt",
                       f"FAILED TO GENERATE\n\nError: {e}".encode()))

    # 3. 805 Affidavit (always)
    try:
        files.append((f"03_805_Affidavit_{last_name}.docx", generate_805(data)))
        print("[OK] 03 805 Affidavit")
    except Exception as e:
        print(f"[ERR] 03 805 Affidavit: {e}")
        traceback.print_exc()
        errors.append(f"805 Affidavit: {e}")
        files.append((f"03_MISSING_805_Affidavit.txt",
                       f"FAILED TO GENERATE\n\nError: {e}".encode()))

    # 4. Affidavit of Heirship (non-probate only)
    if non_probate:
        try:
            files.append((f"04_Affidavit_of_Heirship_{last_name}.docx", generate_heirship(data)))
            print("[OK] 04 Heirship affidavit")
        except Exception as e:
            print(f"[ERR] 04 Heirship affidavit: {e}")
            traceback.print_exc()
            errors.append(f"Heirship affidavit: {e}")
            files.append((f"04_MISSING_Heirship_Affidavit.txt",
                           f"FAILED TO GENERATE\n\nError: {e}".encode()))

    # 5. FT-1 Family Tree (non-probate only)
    if non_probate:
        try:
            files.append((f"05_FT1_Family_Tree_{last_name}.pdf", generate_ft1(data)))
            print("[OK] 05 FT-1 Family Tree")
        except Exception as e:
            print(f"[ERR] 05 FT-1: {e}")
            traceback.print_exc()
            errors.append(f"FT-1: {e}")
            files.append((f"05_MISSING_FT1_Family_Tree.txt",
                           f"FAILED TO GENERATE\n\nError: {e}".encode()))

    # 6/5. Attorney certification (06 for non-probate, 05 for probate)
    atty_num = "06" if non_probate else "05"
    try:
        files.append((f"{atty_num}_Attorney_Certification_{last_name}.docx",
                       generate_attorney_cert(data)))
        print(f"[OK] {atty_num} Attorney cert")
    except Exception as e:
        print(f"[ERR] {atty_num} Attorney cert: {e}")
        traceback.print_exc()
        errors.append(f"Attorney cert: {e}")
        files.append((f"{atty_num}_MISSING_Attorney_Cert.txt",
                       f"FAILED TO GENERATE\n\nError: {e}".encode()))

    # 7/6. Waiver cover letters for distributees who agreed
    waiver_num = "07" if non_probate else "06"
    distributees = data.get("distributees", [])
    for dist in distributees:
        if dist.get("disposition") == "waiver" and dist.get("name"):
            try:
                fname = f"{waiver_num}_Waiver_Cover_{dist['name'].replace(' ','_')}.docx"
                files.append((fname, generate_waiver_cover(data, dist)))
                print(f"[OK] {waiver_num} Waiver cover: {dist['name']}")
            except Exception as e:
                print(f"[ERR] {waiver_num} Waiver cover {dist.get('name')}: {e}")
                traceback.print_exc()
                errors.append(f"Waiver cover for {dist.get('name')}: {e}")
                safe = dist['name'].replace(' ', '_')
                files.append((f"{waiver_num}_MISSING_Waiver_{safe}.txt",
                               f"FAILED TO GENERATE\n\nError: {e}".encode()))

    # 00. Summary sheet (prepended)
    summary = build_summary(data, proceeding, len(files), errors)
    files.insert(0, (f"00_CASE_SUMMARY_{last_name}.txt", summary.encode("utf-8")))

    if errors:
        files.append(("ERRORS.txt", "\n".join(errors).encode("utf-8")))

    print(f"[PACKET] {len(files)} files in ZIP  errors={errors or 'none'}\n")

    zip_bytes = make_zip(files)
    zip_name = f"{last_name}_packet.zip"

    response = send_file(
        io.BytesIO(zip_bytes),
        as_attachment=True,
        download_name=zip_name,
        mimetype="application/zip"
    )
    if errors:
        response.headers["X-Generation-Errors"] = str(len(errors))
    return response


def build_summary(data, proceeding, doc_count, errors):
    proc_display = {
        "Probate": "PROBATE",
        "Administration": "ADMINISTRATION", 
        "NonDomiciliary": "NON-DOMICILIARY ADMINISTRATION",
        "Ancillary": "ANCILLARY ADMINISTRATION"
    }.get(proceeding, proceeding.upper())
    lines = [
        "=" * 60,
        f"  CASE SUMMARY — {proc_display} PROCEEDING",
        "=" * 60,
        f"  Generated: {today()}",
        "",
        f"  DECEDENT:   {decedent_full(data)}",
        f"  A/K/A:      {data.get('decedentAKA', 'N/A')}",
        f"  DATE OF DEATH: {data.get('decedentDOD', '')}",
        f"  PLACE OF DEATH: {data.get('decedentPlaceOfDeath', '')}",
        f"  DOMICILE:   {data.get('decedentStreet', '')}, {data.get('decedentCity', '')}, {data.get('decedentState', '')} {data.get('decedentZip', '')}",
        "",
        f"  PETITIONER: {petitioner_full(data)}",
        f"  ADDRESS:    {data.get('petitionerStreet', '')}, {data.get('petitionerCity', '')}, {data.get('petitionerState', '')}",
        "",
        f"  COURT:      {data.get('county', '')} County Surrogate's Court",
        f"  LETTERS:    {data.get('lettersType', '')} to {data.get('lettersTo', '')}",
        "",
        "  ESTATE VALUE:",
        f"    Personal Property:  ${data.get('personalPropertyValue', '0')}",
        f"    Real Property:      ${data.get('realPropertyValue', '0')}",
        "",
        "  DISTRIBUTEES:",
    ]
    for d in data.get("distributees", []):
        if d.get("name"):
            disp = d.get("disposition", "tbd").upper()
            lines.append(f"    {d['name']} ({d.get('relationship', '')}) — {disp}")
    lines += [
        "",
        f"  DOCUMENTS GENERATED: {doc_count}",
    ]
    if data.get("proceedingType") == "Probate":
        lines.append(f"  SELF-PROVING AFFIDAVIT: {'Yes — witness affidavit not needed' if data.get('selfProvingAffidavit') else 'No — include witness affidavit'}")
    if errors:
        lines += ["", "  ERRORS:", *[f"    - {e}" for e in errors]]
    lines += ["", "=" * 60]
    return "\n".join(lines)


@app.route("/check")
def check():
    pdfs = {
        "Probate-_NY_Court_Forms.pdf": os.path.exists(
            os.path.join(os.path.dirname(__file__), "Probate-_NY_Court_Forms.pdf")),
        "admin_ancil.pdf": os.path.exists(
            os.path.join(os.path.dirname(__file__), "admin_ancil.pdf")),
    }
    templates = {f: os.path.exists(os.path.join(os.path.dirname(__file__), "templates", f))
                 for f in ["805_Affidavit_of_Assets_and_Liabilities_template.docx",
                           "Affidavit_of_Heirship_Full_Admin.docx",
                           "Waiver_cover_letter.docx",
                           "newcertform_6_59_19_PM.docx"]}
    return jsonify({"pdfs": pdfs, "templates": templates, "status": "ok"})


@app.route("/parse-pdf", methods=["POST"])
def parse_pdf():
    """Extract text from uploaded intake PDF and return it."""
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    f = request.files["file"]
    if not f.filename.lower().endswith(".pdf"):
        return jsonify({"error": "Please upload a PDF file"}), 400

    try:
        import pdfplumber
        import io as _io
        pdf_bytes = f.read()
        text_pages = []
        with pdfplumber.open(_io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    text_pages.append(t)
        full_text = "\n\n".join(text_pages)
        return jsonify({"text": full_text, "pages": len(text_pages)})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    print("=" * 60)
    print("  NY Surrogate's Court Probate Assistant v2")
    print("=" * 60)
    print(f"  Counties: {', '.join(COUNTY_INFO.keys())}")
    print()
    print("  Open: http://localhost:8080")
    print("=" * 60)
    app.run(debug=True, port=8080)
