"""
NY Surrogate's Court Probate Assistant v2
Full document packet generator

Run: python3 app.py
Open: http://localhost:8080
"""

import io
import ipaddress
import json
import os
import traceback
import zipfile
try:
    from dotenv import load_dotenv
    # Load .env from same directory as app.py (works both in dev and packaged app)
    load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"))
except ImportError:
    pass
from flask import Flask, request, jsonify, send_file, send_from_directory

# ── Output folder settings ────────────────────────────────────────────────────
SETTINGS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "settings.json")

def _load_settings():
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, "r") as f:
                return json.load(f)
        except (json.JSONDecodeError, IOError):
            return {}
    return {}

def _save_settings(settings):
    with open(SETTINGS_FILE, "w") as f:
        json.dump(settings, f, indent=2)

import glob as _glob


def get_drive_roots():
    """Return all connected drive roots to search for estate folders."""
    home = os.path.expanduser("~")
    roots = []
    # Local Documents
    docs = os.path.join(home, "Documents")
    if os.path.isdir(docs):
        roots.append(docs)
    # Dropbox
    dropbox = os.path.join(home, "Dropbox")
    if os.path.isdir(dropbox):
        roots.append(dropbox)
        clio = os.path.join(dropbox, "Clio")
        if os.path.isdir(clio):
            roots.append(clio)
    # Google Drive (multiple possible locations)
    for pattern in [
        os.path.join(home, "Google Drive"),
        os.path.join(home, "Google Drive", "My Drive"),
        os.path.join(home, "Library", "CloudStorage", "GoogleDrive-*"),
        os.path.join(home, "Library", "CloudStorage", "GoogleDrive-*", "My Drive"),
    ]:
        for p in _glob.glob(pattern):
            if os.path.isdir(p):
                roots.append(p)
    # OneDrive
    for pattern in [
        os.path.join(home, "OneDrive"),
        os.path.join(home, "OneDrive - *"),
        os.path.join(home, "Library", "CloudStorage", "OneDrive-*"),
    ]:
        for p in _glob.glob(pattern):
            if os.path.isdir(p):
                roots.append(p)
    # iCloud Drive
    icloud = os.path.join(home, "Library", "Mobile Documents", "com~apple~CloudDocs")
    if os.path.isdir(icloud):
        roots.append(icloud)
    # Desktop
    desktop = os.path.join(home, "Desktop")
    if os.path.isdir(desktop):
        roots.append(desktop)
    return roots


def find_estate_folder(decedent_name):
    """Search all connected drives for an existing 'Estate of [Name]' folder.

    Searches up to 2 levels deep in each root. Returns list of matches.
    """
    target = f"Estate of {decedent_name}".lower()
    matches = []
    seen = set()
    for root in get_drive_roots():
        try:
            for entry in os.scandir(root):
                if not entry.is_dir():
                    continue
                if entry.name.lower() == target:
                    real = os.path.realpath(entry.path)
                    if real not in seen:
                        seen.add(real)
                        matches.append(entry.path)
                # Also check one level deeper
                try:
                    for sub in os.scandir(entry.path):
                        if sub.is_dir() and sub.name.lower() == target:
                            real = os.path.realpath(sub.path)
                            if real not in seen:
                                seen.add(real)
                                matches.append(sub.path)
                except (PermissionError, OSError):
                    pass
        except (PermissionError, OSError):
            pass
    return matches


def get_output_folder():
    """Return the configured output folder, or best default."""
    settings = _load_settings()
    folder = settings.get("output_folder", "")
    if folder and os.path.isdir(folder):
        return folder
    dropbox = os.path.expanduser("~/Dropbox/Clio")
    if os.path.isdir(dropbox):
        return dropbox
    return os.path.expanduser("~/Documents")


def save_to_output(data, files):
    """Save generated files to the estate folder.

    First searches all drives for an existing 'Estate of [Name]' folder.
    If found, saves to [found folder]/Drafts/.
    If not found, creates it in the configured output folder.

    Returns the folder path on success, or raises on failure.
    files = list of (filename, bytes)
    """
    first = (data.get("decedentFirstName") or "").strip()
    last  = (data.get("decedentLastName")  or "").strip()
    name  = f"{first} {last}".strip() or "Unknown"

    # Search for existing estate folder across all drives
    existing = find_estate_folder(name)
    if existing:
        # Use the first match — save to Drafts subfolder
        estate_dir = os.path.join(existing[0], "Drafts")
        print(f"[OUTPUT] Found existing estate folder: {existing[0]}")
    else:
        # Create new in configured output folder
        base = get_output_folder()
        estate_dir = os.path.join(base, f"Estate of {name}", "Drafts")
        print(f"[OUTPUT] No existing folder found. Creating in: {base}")

    os.makedirs(estate_dir, exist_ok=True)
    for fname, fbytes in files:
        dest = os.path.join(estate_dir, fname)
        with open(dest, "wb") as fh:
            fh.write(fbytes)
    return estate_dir
from generators import (
    generate_cover_letter, generate_805, generate_heirship,
    generate_waiver_cover, generate_attorney_cert,
    generate_probate_docs, fill_ancillary_pdf,
    fill_administration_pdf, fill_nondom_pdf, fill_cta_pdf, generate_ft1,
    generate_auth_letter, generate_instruction_letter,
    generate_accounting_excel, fill_schedule_da_pdf,
    needs_family_tree_affidavit, needs_family_tree_diagram,
    family_tree_trigger_reason,
    # New PDF fill functions (Admin forms)
    fill_waiver_individual_pdf, fill_waiver_corporate_pdf,
    fill_citation_pdf, fill_affidavit_of_service_pdf,
    fill_notice_of_application_pdf, fill_affidavit_of_mailing_pdf,
    fill_affidavit_of_regularity_pdf, fill_proposed_decree_pdf,
    fill_schedule_a_pdf, fill_schedule_b_pdf,
    fill_schedule_c_pdf, fill_schedule_d_pdf,
    # New Word template generators
    generate_waiver_probate, generate_bond_affidavit,
    generate_notice_of_probate, generate_petition_scpa_2203,
    COUNTY_INFO, today, decedent_full, petitioner_full
)

app = Flask(__name__, static_folder="static")

# ── IP Allowlist ───────────────────────────────────────────────────────────────
# Allow localhost and the local network subnet.
# To restrict to specific machines only, replace the subnet with individual IPs:
#   ALLOWED = ["127.0.0.1", "::1", "192.168.1.251", "192.168.1.100"]
ALLOWED_NETWORKS = [
    ipaddress.ip_network("127.0.0.0/8"),    # localhost
    ipaddress.ip_network("::1/128"),         # IPv6 localhost
    ipaddress.ip_network("192.168.1.0/24"), # local Wi-Fi network
]

@app.before_request
def check_ip():
    try:
        remote = ipaddress.ip_address(request.remote_addr)
    except ValueError:
        return "Forbidden", 403
    if not any(remote in net for net in ALLOWED_NETWORKS):
        return "Forbidden", 403


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
    non_probate = proceeding in ("Administration", "Ancillary", "NonDomiciliary", "AdminCTA")
    files = []
    errors = []

    print(f"\n[PACKET] proceeding={proceeding!r}  decedent={last_name!r}")

    # ── 01. Court cover letter (always) ─────────────────────────────────────────
    try:
        print("[TRYING] 01 generate_cover_letter()")
        files.append((f"01_Cover_Letter_{last_name}.docx", generate_cover_letter(data)))
        print("[OK] 01 Cover letter")
    except Exception as e:
        print(f"[ERR] 01 Cover letter: {e}")
        traceback.print_exc()
        errors.append(f"Cover letter: {e}")
        files.append((f"01_MISSING_Cover_Letter.txt",
                       f"FAILED TO GENERATE\n\nError: {e}".encode()))

    # ── 02-04. Main petition (and oath + witness for Probate) ────────────────────
    # Probate:         02=P-1, 03=Oath & Designation, 04=Attesting Witness
    # Administration:  02=A-1 petition
    # Ancillary:       02=AA-1 petition
    # NonDomiciliary:  02=instructions note
    if proceeding == "Probate":
        try:
            print("[TRYING] generate_probate_docs() → P-1 + Oath & Designation + Attesting Witness")
            for fname, fbytes in generate_probate_docs(data):
                files.append((fname, fbytes))
                print(f"[OK] {fname}")
        except Exception as e:
            print(f"[ERR] Probate docs: {e}")
            traceback.print_exc()
            errors.append(f"Probate docs: {e}")
            files.append((f"02_MISSING_Probate_Docs.txt",
                           f"FAILED TO GENERATE\n\nError: {e}".encode()))
    else:
        try:
            print(f"[TRYING] 02 petition for proceeding={proceeding!r}")
            if proceeding == "Ancillary":
                files.append((f"02_Petition_Ancillary_AA1_{last_name}.pdf", fill_ancillary_pdf(data)))
                print("[OK] 02 Ancillary petition (AA-1)")
            elif proceeding == "Administration":
                files.append((f"02_Petition_Administration_A1_{last_name}.pdf",
                               fill_administration_pdf(data)))
                print("[OK] 02 Administration petition (A-1)")
            elif proceeding == "NonDomiciliary":
                files.append((f"02_Petition_NonDomiciliary_{last_name}.pdf",
                               fill_nondom_pdf(data)))
                print("[OK] 02 NonDomiciliary petition (Non-Dom)")
            elif proceeding == "AdminCTA":
                files.append((f"02_Petition_AdminCTA_{last_name}.pdf",
                               fill_cta_pdf(data)))
                print("[OK] 02 Admin CTA petition (CTA-1)")
            else:
                print(f"[WARN] 02 No petition — unrecognised proceedingType={proceeding!r}")
        except Exception as e:
            print(f"[ERR] 02 petition: {e}")
            traceback.print_exc()
            errors.append(f"Main petition: {e}")
            files.append((f"02_MISSING_Main_Petition.txt",
                           f"FAILED TO GENERATE\n\nError: {e}".encode()))

    # ── 805 Affidavit (always)
    # Probate:    slot 05  (after petition 02, oath 03, witness 04)
    # Non-probate: slot 03  (petition is only 02)
    afft_num = "05" if proceeding == "Probate" else "03"
    try:
        print(f"[TRYING] {afft_num} generate_805()")
        files.append((f"{afft_num}_805_Affidavit_{last_name}.docx", generate_805(data)))
        print(f"[OK] {afft_num} 805 Affidavit")
    except Exception as e:
        print(f"[ERR] {afft_num} 805 Affidavit: {e}")
        traceback.print_exc()
        errors.append(f"805 Affidavit: {e}")
        files.append((f"{afft_num}_MISSING_805_Affidavit.txt",
                       f"FAILED TO GENERATE\n\nError: {e}".encode()))

    # ── Affidavit of Heirship ────────────────────────────────────────────────────
    # Non-probate: always (slot 04)
    # Probate:     only when Rule 207.16(c) triggers (slot 06)
    if non_probate:
        try:
            print("[TRYING] 04 generate_heirship()")
            files.append((f"04_Affidavit_of_Heirship_{last_name}.docx", generate_heirship(data)))
            print("[OK] 04 Heirship affidavit")
        except Exception as e:
            print(f"[ERR] 04 Heirship affidavit: {e}")
            traceback.print_exc()
            errors.append(f"Heirship affidavit: {e}")
            files.append((f"04_MISSING_Heirship_Affidavit.txt",
                           f"FAILED TO GENERATE\n\nError: {e}".encode()))

    # Rule 207.16(c) check for Probate
    probate_needs_ft_aff  = (proceeding == "Probate") and needs_family_tree_affidavit(data)
    probate_needs_ft_diag = (proceeding == "Probate") and needs_family_tree_diagram(data)

    if probate_needs_ft_aff:
        reason = family_tree_trigger_reason(data)
        print(f"[207.16(c)] Family tree affidavit required for Probate — {reason}")
        try:
            print("[TRYING] 06 generate_heirship() [Probate 207.16(c)]")
            files.append((f"06_Affidavit_of_Heirship_{last_name}.docx", generate_heirship(data)))
            print("[OK] 06 Heirship affidavit (Probate)")
        except Exception as e:
            print(f"[ERR] 06 Heirship affidavit (Probate): {e}")
            traceback.print_exc()
            errors.append(f"Heirship affidavit (Probate): {e}")
            files.append((f"06_MISSING_Heirship_Affidavit.txt",
                           f"FAILED TO GENERATE\n\nError: {e}".encode()))

    # ── FT-1 Family Tree ─────────────────────────────────────────────────────────
    # Non-probate: always (slot 05)
    # Probate:     only when Rule 207.16(c) requires the diagram (slot 07)
    #              (diagram not required for sole spouse or sole child)
    if non_probate:
        try:
            print("[TRYING] 05 generate_ft1()")
            files.append((f"05_FT1_Family_Tree_{last_name}.pdf", generate_ft1(data)))
            print("[OK] 05 FT-1 Family Tree")
        except Exception as e:
            print(f"[ERR] 05 FT-1: {e}")
            traceback.print_exc()
            errors.append(f"FT-1: {e}")
            files.append((f"05_MISSING_FT1_Family_Tree.txt",
                           f"FAILED TO GENERATE\n\nError: {e}".encode()))

    if probate_needs_ft_diag:
        try:
            print("[TRYING] 07 generate_ft1() [Probate 207.16(c)]")
            files.append((f"07_FT1_Family_Tree_{last_name}.pdf", generate_ft1(data)))
            print("[OK] 07 FT-1 Family Tree (Probate)")
        except Exception as e:
            print(f"[ERR] 07 FT-1 (Probate): {e}")
            traceback.print_exc()
            errors.append(f"FT-1 (Probate): {e}")
            files.append((f"07_MISSING_FT1_Family_Tree.txt",
                           f"FAILED TO GENERATE\n\nError: {e}".encode()))

    # ── Attorney certification ────────────────────────────────────────────────────
    # Slot 06 normally; shifts to 08 when Probate 207.16(c) docs are inserted
    cert_slot = "08" if probate_needs_ft_aff else "06"
    try:
        print(f"[TRYING] {cert_slot} generate_attorney_cert()")
        files.append((f"{cert_slot}_Attorney_Certification_{last_name}.docx",
                       generate_attorney_cert(data)))
        print(f"[OK] {cert_slot} Attorney cert")
    except Exception as e:
        print(f"[ERR] {cert_slot} Attorney cert: {e}")
        traceback.print_exc()
        errors.append(f"Attorney cert: {e}")
        files.append((f"{cert_slot}_MISSING_Attorney_Cert.txt",
                       f"FAILED TO GENERATE\n\nError: {e}".encode()))

    # ── Waiver cover letters (all proceedings)
    # Slot 07 normally; shifts to 09 when Probate 207.16(c) inserts any docs (cert shifts to 08)
    waiver_slot = "09" if probate_needs_ft_aff else "07"
    distributees = data.get("distributees", [])
    for dist in distributees:
        if dist.get("disposition") == "waiver" and dist.get("name"):
            try:
                print(f"[TRYING] {waiver_slot} generate_waiver_cover() for {dist['name']!r}")
                fname = f"{waiver_slot}_Waiver_Cover_{dist['name'].replace(' ','_')}.docx"
                files.append((fname, generate_waiver_cover(data, dist)))
                print(f"[OK] {waiver_slot} Waiver cover: {dist['name']}")
            except Exception as e:
                print(f"[ERR] {waiver_slot} Waiver cover {dist.get('name')}: {e}")
                traceback.print_exc()
                errors.append(f"Waiver cover for {dist.get('name')}: {e}")
                safe = dist['name'].replace(' ', '_')
                files.append((f"{waiver_slot}_MISSING_Waiver_{safe}.txt",
                               f"FAILED TO GENERATE\n\nError: {e}".encode()))

    # ── Waiver form PDFs (alongside cover letters)
    # Probate → P-4 (Word), Admin/NonDom/Ancillary/CTA → A-8 individual PDF
    for dist in distributees:
        if dist.get("disposition") == "waiver" and dist.get("name"):
            safe = dist['name'].replace(' ', '_')
            if proceeding == "Probate":
                try:
                    print(f"[TRYING] {waiver_slot} generate_waiver_probate() for {dist['name']!r}")
                    fname = f"{waiver_slot}_Waiver_P4_{safe}.docx"
                    files.append((fname, generate_waiver_probate(data, dist)))
                    print(f"[OK] {waiver_slot} Waiver P-4: {dist['name']}")
                except Exception as e:
                    print(f"[ERR] {waiver_slot} Waiver P-4 {dist.get('name')}: {e}")
                    traceback.print_exc()
                    errors.append(f"Waiver P-4 for {dist.get('name')}: {e}")
                    files.append((f"{waiver_slot}_MISSING_Waiver_P4_{safe}.txt",
                                   f"FAILED TO GENERATE\n\nError: {e}".encode()))
            else:
                # Admin proceeding — A-8 (individual) or A-9 (corporate)
                is_corp = dist.get("isCorporate", False)
                try:
                    if is_corp:
                        print(f"[TRYING] {waiver_slot} fill_waiver_corporate_pdf() for {dist['name']!r}")
                        fname = f"{waiver_slot}_Waiver_A9_Corp_{safe}.pdf"
                        files.append((fname, fill_waiver_corporate_pdf(data, dist)))
                        print(f"[OK] {waiver_slot} Waiver A-9 (Corp): {dist['name']}")
                    else:
                        print(f"[TRYING] {waiver_slot} fill_waiver_individual_pdf() for {dist['name']!r}")
                        fname = f"{waiver_slot}_Waiver_A8_{safe}.pdf"
                        files.append((fname, fill_waiver_individual_pdf(data, dist)))
                        print(f"[OK] {waiver_slot} Waiver A-8: {dist['name']}")
                except Exception as e:
                    form_type = "A-9 Corp" if is_corp else "A-8"
                    print(f"[ERR] {waiver_slot} Waiver {form_type} {dist.get('name')}: {e}")
                    traceback.print_exc()
                    errors.append(f"Waiver {form_type} for {dist.get('name')}: {e}")
                    files.append((f"{waiver_slot}_MISSING_Waiver_{safe}.txt",
                                   f"FAILED TO GENERATE\n\nError: {e}".encode()))

    # ── Schedule D(a) — post-deceased distributees (same slot as waivers)
    for dist in distributees:
        if dist.get("disposition") == "postDeceased" and dist.get("name"):
            try:
                safe = dist['name'].replace(' ', '_')
                print(f"[TRYING] {waiver_slot} fill_schedule_da_pdf() for {dist['name']!r}")
                fname = f"{waiver_slot}_Schedule_Da_{safe}.pdf"
                files.append((fname, fill_schedule_da_pdf(data, dist)))
                print(f"[OK] {waiver_slot} Schedule D(a): {dist['name']}")
            except Exception as e:
                print(f"[ERR] {waiver_slot} Schedule D(a) {dist.get('name')}: {e}")
                traceback.print_exc()
                errors.append(f"Schedule D(a) for {dist.get('name')}: {e}")
                safe = dist['name'].replace(' ', '_')
                files.append((f"{waiver_slot}_MISSING_Schedule_Da_{safe}.txt",
                               f"FAILED TO GENERATE\n\nError: {e}".encode()))

    # ── Bond Affidavit (all proceeding types)
    bond_slot = waiver_slot  # same slot grouping as waivers
    try:
        print(f"[TRYING] {bond_slot} generate_bond_affidavit()")
        files.append((f"{bond_slot}_Bond_Affidavit_{last_name}.docx",
                       generate_bond_affidavit(data)))
        print(f"[OK] {bond_slot} Bond Affidavit")
    except Exception as e:
        print(f"[ERR] {bond_slot} Bond Affidavit: {e}")
        traceback.print_exc()
        errors.append(f"Bond Affidavit: {e}")
        files.append((f"{bond_slot}_MISSING_Bond_Affidavit.txt",
                       f"FAILED TO GENERATE\n\nError: {e}".encode()))

    # TODO: Schedules A-D (Nonmarital, Adoption, Infants, Disability)
    # These per-distributee schedules require additional UI fields not yet collected:
    #   - fill_schedule_a_pdf(data, dist) — for nonmarital distributees
    #   - fill_schedule_b_pdf(data, dist) — for adopted distributees
    #   - fill_schedule_c_pdf(data, dist) — for infant distributees
    #   - fill_schedule_d_pdf(data, dist) — for disabled distributees
    # Wire these in once the UI collects distributee sub-type attributes.

    # NOTE: The following post-filing forms are available as standalone functions
    # but are NOT auto-generated in the initial packet:
    #   - fill_citation_pdf(data)
    #   - fill_affidavit_of_service_pdf(data)
    #   - fill_notice_of_application_pdf(data)
    #   - fill_affidavit_of_mailing_pdf(data)
    #   - fill_affidavit_of_regularity_pdf(data)
    #   - fill_proposed_decree_pdf(data)
    #   - generate_notice_of_probate(data)
    #   - generate_petition_scpa_2203(data)
    # These can be wired up later as standalone API endpoints.

    # 00. Summary sheet (prepended)
    if proceeding == "Probate":
        if probate_needs_ft_aff:
            reason = family_tree_trigger_reason(data)
            diag_note = "" if probate_needs_ft_diag else " (diagram waived — sole spouse/child)"
            ft_info = f"Required{diag_note} — {reason}"
        else:
            ft_info = "Not required (spouse/children as distributees)"
    else:
        ft_info = "N/A — included in all non-probate proceedings"
    summary = build_summary(data, proceeding, len(files), errors, ft_info=ft_info)
    files.insert(0, (f"00_CASE_SUMMARY_{last_name}.txt", summary.encode("utf-8")))

    if errors:
        files.append(("ERRORS.txt", "\n".join(errors).encode("utf-8")))

    print(f"[PACKET] {len(files)} files in ZIP  errors={errors or 'none'}\n")

    # ── Auto-save to Dropbox/Clio/Estate of X/Drafts/ ────────────────────────
    saved_to = None
    try:
        saved_to = save_to_output(data, files)
        print(f"[DROPBOX] Saved {len(files)} files → {saved_to}")
    except Exception as e:
        print(f"[DROPBOX] Save failed: {e}")
        traceback.print_exc()

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
    if saved_to:
        response.headers["X-Saved-To"] = saved_to
    return response


def build_summary(data, proceeding, doc_count, errors, ft_info=None):
    proc_display = {
        "Probate": "PROBATE",
        "Administration": "ADMINISTRATION",
        "NonDomiciliary": "NON-DOMICILIARY ADMINISTRATION",
        "Ancillary": "ANCILLARY ADMINISTRATION",
        "AdminCTA": "ADMINISTRATION C.T.A."
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
    if ft_info:
        lines.append(f"  RULE 207.16(c): {ft_info}")
    if errors:
        lines += ["", "  ERRORS:", *[f"    - {e}" for e in errors]]
    lines += ["", "=" * 60]
    return "\n".join(lines)


@app.route("/generate-auth-letter", methods=["POST"])
def gen_auth_letter():
    body = request.get_json()
    data = body.get("data", {})
    asset = body.get("asset", {})
    doc_bytes = generate_auth_letter(data, asset)
    last = data.get("decedentLastName", "Estate").replace(" ", "_")
    inst = asset.get("institution", "Institution").replace(" ", "_")[:30]
    filename = f"Auth_Letter_{inst}_{last}.docx"
    buf = io.BytesIO(doc_bytes)
    resp = send_file(buf, as_attachment=True, download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    resp.headers["X-Filename"] = filename
    return resp


@app.route("/generate-instruction-letter", methods=["POST"])
def gen_instruction_letter():
    body = request.get_json()
    data = body.get("data", {})
    asset = body.get("asset", {})
    action = body.get("marshalAction", "check")
    doc_bytes = generate_instruction_letter(data, asset, action)
    last = data.get("decedentLastName", "Estate").replace(" ", "_")
    inst = asset.get("institution", "Institution").replace(" ", "_")[:30]
    filename = f"Instruction_Letter_{inst}_{last}.docx"
    buf = io.BytesIO(doc_bytes)
    resp = send_file(buf, as_attachment=True, download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    resp.headers["X-Filename"] = filename
    return resp


@app.route("/generate-accounting-excel", methods=["POST"])
def gen_accounting_excel():
    body = request.get_json()
    data = body.get("data", {})
    assets_data = body.get("assets", [])
    xls_bytes = generate_accounting_excel(data, assets_data)
    last = data.get("decedentLastName", "Estate").replace(" ", "_")
    filename = f"Accounting_{last}.xlsx"
    buf = io.BytesIO(xls_bytes)
    resp = send_file(buf, as_attachment=True, download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    resp.headers["X-Filename"] = filename
    return resp


@app.route("/smart-intake", methods=["POST"])
def smart_intake():
    """Accept one or more PDFs, extract text, send to Claude, return probate field JSON."""
    if not request.files:
        return jsonify({"error": "No files uploaded"}), 400

    api_key = os.environ.get("ANTHROPIC_API_KEY", "")

    import pdfplumber
    import anthropic as _anthropic

    # ── Extract text from all uploaded PDFs ───────────────────────────────────
    doc_texts = []
    for key in request.files:
        f = request.files[key]
        if not f.filename.lower().endswith(".pdf"):
            continue
        try:
            pdf_bytes = f.read()
            pages = []
            # Try text extraction first
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                for page in pdf.pages:
                    t = page.extract_text()
                    if t and t.strip():
                        pages.append(t)
            # If no text found, try OCR for scanned documents
            if not pages:
                try:
                    import pytesseract
                    from pdf2image import convert_from_bytes
                    print(f"[SMART-INTAKE] No text in {f.filename}, trying OCR...")
                    images = convert_from_bytes(pdf_bytes, dpi=300)
                    for img in images:
                        t = pytesseract.image_to_string(img)
                        if t and t.strip():
                            pages.append(t)
                    if pages:
                        print(f"[SMART-INTAKE] OCR extracted {len(pages)} pages from {f.filename}")
                except ImportError:
                    print(f"[SMART-INTAKE] OCR not available, trying image-based Claude extraction")
                except Exception as ocr_err:
                    print(f"[SMART-INTAKE] OCR failed: {ocr_err}, trying image-based Claude extraction")

            # If still no text, send pages as images to Claude directly
            if not pages:
                try:
                    import fitz as _fitz
                    print(f"[SMART-INTAKE] Converting {f.filename} to images for Claude vision...")
                    pdf_doc = _fitz.open(stream=pdf_bytes, filetype="pdf")
                    import base64 as _b64
                    image_contents = []
                    for page_num in range(min(len(pdf_doc), 10)):  # max 10 pages
                        pix = pdf_doc[page_num].get_pixmap(dpi=200)
                        img_bytes = pix.tobytes("png")
                        img_b64 = _b64.b64encode(img_bytes).decode("utf-8")
                        image_contents.append({
                            "type": "image",
                            "source": {"type": "base64", "media_type": "image/png", "data": img_b64}
                        })
                    pdf_doc.close()
                    if image_contents:
                        # Send images directly to Claude for extraction
                        _api_key = os.environ.get("ANTHROPIC_API_KEY", "")
                        client = _anthropic.Anthropic(api_key=_api_key)
                        vision_prompt = image_contents + [{"type": "text", "text": "Extract ALL text from these scanned document pages. Return the full text content."}]
                        msg = client.messages.create(
                            model="claude-sonnet-4-6",
                            max_tokens=4096,
                            messages=[{"role": "user", "content": vision_prompt}],
                        )
                        extracted_text = msg.content[0].text.strip()
                        if extracted_text:
                            pages.append(extracted_text)
                            print(f"[SMART-INTAKE] Claude vision extracted text from {f.filename}")
                except Exception as vision_err:
                    print(f"[SMART-INTAKE] Vision extraction failed: {vision_err}")

            text = "\n\n".join(pages).strip()
            if text:
                doc_texts.append(f"=== {f.filename} ===\n{text}")
        except Exception as e:
            print(f"[SMART-INTAKE] Error reading {f.filename}: {e}")

    if not doc_texts:
        return jsonify({"error": "Could not extract any text from the uploaded PDFs."}), 400

    combined = "\n\n".join(doc_texts)

    # ── Claude prompt ──────────────────────────────────────────────────────────
    prompt = f"""You are a New York probate attorney's assistant. Extract information from the uploaded legal documents and return it as JSON to pre-fill a Surrogate's Court petition.

RULES:
- Dates: MM/DD/YYYY format
- Money: numbers only, no $ or commas (e.g. "150000")
- Missing fields: use null
- maritalStatus must be one of: never_married, married, divorced, widowed
- proceedingType must be one of: Probate, Administration (Probate if a Will exists, Administration if no Will)
- survivingX fields: "Yes" if that class survives, "No" if they existed but predeceased, null if unknown
- For distributees: include all known heirs with name, relationship, address (if available), citizenship (default "US Citizen" if unknown)

Return ONLY valid JSON with this exact structure (use null for unknown fields):

{{
  "proceedingType": "Probate or Administration",
  "decedentFirstName": null,
  "decedentMiddleName": null,
  "decedentLastName": null,
  "decedentAKA": null,
  "decedentDOB": null,
  "decedentDOD": null,
  "decedentPlaceOfDeath": null,
  "decedentStreet": null,
  "decedentCity": null,
  "decedentState": null,
  "decedentZip": null,
  "decedentCitizenship": null,
  "ssn": null,
  "maritalStatus": null,
  "spouseName": null,
  "divorceYear": null,
  "priorSpouseDeathDate": null,
  "motherName": null,
  "motherDOD": null,
  "fatherName": null,
  "fatherDOD": null,
  "childrenNote": null,
  "petitionerFirstName": null,
  "petitionerMiddleName": null,
  "petitionerLastName": null,
  "petitionerStreet": null,
  "petitionerCity": null,
  "petitionerState": null,
  "petitionerZip": null,
  "petitionerRelationship": null,
  "petitionerCitizenship": null,
  "personalPropertyValue": null,
  "realPropertyValue": null,
  "willDate": null,
  "codicilDate": null,
  "witness1": null,
  "witness2": null,
  "lettersTo": null,
  "survivingSpouse": null,
  "survivingChildren": null,
  "survivingParents": null,
  "survivingSiblings": null,
  "survivingGrandparents": null,
  "survivingAuntsUncles": null,
  "survivingFirstCousinsOnceRemoved": null,
  "distributees": []
}}

Each distributee in the array should be:
{{"name": "Full Name", "relationship": "Son/Daughter/Spouse/etc", "address": "full address or null", "citizenship": "US Citizen"}}

=== DOCUMENTS ===
{combined}"""

    # ── Call Claude ────────────────────────────────────────────────────────────
    try:
        client = _anthropic.Anthropic(api_key=api_key)
        message = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=4096,
            messages=[{"role": "user", "content": prompt}],
        )
        response_text = message.content[0].text.strip()

        # Strip markdown code fences if present
        if response_text.startswith("```"):
            response_text = response_text.split("```")[1]
            if response_text.startswith("json"):
                response_text = response_text[4:]
            response_text = response_text.strip()

        # Extract JSON if Claude added surrounding text
        first = response_text.find("{")
        last = response_text.rfind("}")
        if first != -1 and last != -1:
            response_text = response_text[first:last+1]

        extracted = json.loads(response_text)
        return jsonify({"ok": True, "data": extracted})

    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": f"Claude error: {e}"}), 500


# ── Server-side case storage ──────────────────────────────────────────────────
CASES_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cases.json")

def _load_cases():
    if os.path.exists(CASES_FILE):
        try:
            with open(CASES_FILE, "r") as f:
                return json.load(f)
        except (json.JSONDecodeError, IOError):
            return {}
    return {}

def _save_cases(cases):
    with open(CASES_FILE, "w") as f:
        json.dump(cases, f, indent=2)


@app.route("/cases", methods=["GET"])
def get_cases():
    return jsonify(_load_cases())


@app.route("/cases", methods=["POST"])
def save_case():
    body = request.get_json()
    name = body.get("name", "").strip()
    data = body.get("data")
    if not name or data is None:
        return jsonify({"error": "name and data required"}), 400
    cases = _load_cases()
    cases[name] = data
    _save_cases(cases)
    return jsonify({"ok": True})


@app.route("/cases/<name>", methods=["DELETE"])
def delete_case(name):
    cases = _load_cases()
    if name in cases:
        del cases[name]
        _save_cases(cases)
    return jsonify({"ok": True})


@app.route("/settings", methods=["GET"])
def get_settings():
    settings = _load_settings()
    settings["output_folder"] = get_output_folder()
    # Detect available folder options
    options = []
    docs = os.path.expanduser("~/Documents")
    if os.path.isdir(docs):
        options.append({"label": "Documents", "path": docs})
    dropbox = os.path.expanduser("~/Dropbox")
    if os.path.isdir(dropbox):
        options.append({"label": "Dropbox", "path": dropbox})
        clio = os.path.join(dropbox, "Clio")
        if os.path.isdir(clio):
            options.append({"label": "Dropbox/Clio", "path": clio})
    # Check for Google Drive
    for gd in [os.path.expanduser("~/Google Drive"),
               os.path.expanduser("~/Google Drive/My Drive"),
               os.path.expanduser("~/Library/CloudStorage/GoogleDrive-*/My Drive")]:
        import glob as _glob
        for p in _glob.glob(gd):
            if os.path.isdir(p):
                label = "Google Drive" if "My Drive" not in p else "Google Drive/My Drive"
                options.append({"label": label, "path": p})
    settings["folder_options"] = options
    return jsonify(settings)


@app.route("/settings", methods=["POST"])
def update_settings():
    body = request.get_json()
    settings = _load_settings()
    if "output_folder" in body:
        folder = body["output_folder"]
        if folder and os.path.isdir(folder):
            settings["output_folder"] = folder
            _save_settings(settings)
            return jsonify({"ok": True, "output_folder": folder})
        return jsonify({"error": "Folder does not exist"}), 400
    return jsonify({"error": "No settings to update"}), 400


@app.route("/browse-folders", methods=["GET"])
def browse_folders():
    """List subdirectories of a given path for folder browsing."""
    folder = request.args.get("path", os.path.expanduser("~"))
    if not os.path.isdir(folder):
        return jsonify({"error": "Not a directory"}), 400
    try:
        entries = []
        for name in sorted(os.listdir(folder)):
            full = os.path.join(folder, name)
            if os.path.isdir(full) and not name.startswith("."):
                entries.append({"name": name, "path": full})
        return jsonify({"current": folder, "parent": os.path.dirname(folder), "folders": entries})
    except PermissionError:
        return jsonify({"error": "Permission denied"}), 403


@app.route("/find-estate-folder", methods=["GET"])
def find_estate():
    """Search all drives for an existing estate folder matching the decedent name."""
    name = request.args.get("name", "").strip()
    if not name:
        return jsonify({"matches": [], "drives": [r for r in get_drive_roots()]})
    matches = find_estate_folder(name)
    return jsonify({"matches": matches, "name": name})


APP_VERSION = "1.2.0"
GITHUB_REPO = "Ranguana/Surrogate-Court-Forms"


@app.route("/check-update")
def check_update():
    """Check GitHub Releases for a newer version."""
    import urllib.request
    try:
        url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
        req = urllib.request.Request(url, headers={"User-Agent": "ProbateAssistant"})
        with urllib.request.urlopen(req, timeout=5) as resp:
            release = json.loads(resp.read().decode())
        latest = release.get("tag_name", "").lstrip("v")
        if not latest:
            return jsonify({"update": False, "current": APP_VERSION})
        if latest != APP_VERSION:
            # Find DMG download URL
            download_url = ""
            for asset in release.get("assets", []):
                if asset["name"].endswith(".dmg"):
                    download_url = asset["browser_download_url"]
                    break
            return jsonify({
                "update": True,
                "current": APP_VERSION,
                "latest": latest,
                "download_url": download_url,
                "release_notes": release.get("body", ""),
                "html_url": release.get("html_url", ""),
            })
        return jsonify({"update": False, "current": APP_VERSION})
    except Exception as e:
        print(f"[UPDATE] Check failed: {e}")
        return jsonify({"update": False, "current": APP_VERSION, "error": str(e)})


@app.route("/app-version")
def app_version():
    return jsonify({"version": APP_VERSION})


@app.route("/check")
def check():
    pdfs = {
        "Probate-_NY_Court_Forms.pdf": os.path.exists(
            os.path.join(os.path.dirname(__file__), "Probate-_NY_Court_Forms.pdf")),
        "admin_ancil.pdf": os.path.exists(
            os.path.join(os.path.dirname(__file__), "admin_ancil.pdf")),
    }
    templates = {f: os.path.exists(os.path.join(os.path.dirname(__file__), "templates", "Not Using Word Docs", f))
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
    import socket

    port = 52845  # uncommon port to avoid conflicts

    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        local_ip = s.getsockname()[0]
        s.close()
    except Exception:
        local_ip = "unknown"
    print("=" * 60)
    print("  NY Surrogate's Court Probate Assistant v2")
    print("=" * 60)
    print(f"  Counties: {', '.join(COUNTY_INFO.keys())}")
    print()
    print(f"  Local:   http://localhost:{port}")
    print(f"  Network: http://{local_ip}:{port}")
    print()
    print("  Share the Network URL with others on your Wi-Fi.")
    print("=" * 60)
    app.run(debug=False, host="0.0.0.0", port=port)
