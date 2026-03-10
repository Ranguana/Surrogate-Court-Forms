#!/usr/bin/env python3
"""
Test script — generate ALL documents with sample data into test_output/.
Run: python3 test_generate.py

Tests every generator in generators.py with both Administration and Probate data.
"""
import os
import sys
import traceback

sys.path.insert(0, os.path.dirname(__file__))

from generators import (
    generate_cover_letter,
    generate_805,
    generate_heirship,
    fill_administration_pdf,
    fill_ft1_pdf,
    generate_probate_docs,
    fill_probate_pdf,
    fill_ancillary_pdf,
    generate_waiver_cover,
    generate_attorney_cert,
    generate_accounting_excel,
    generate_auth_letter,
    generate_instruction_letter,
)

OUT_DIR = os.path.join(os.path.dirname(__file__), "test_output")
os.makedirs(OUT_DIR, exist_ok=True)

# ─── ADMINISTRATION SAMPLE DATA ────────────────────────────────────────────────

ADMIN_DATA = {
    # Decedent
    "decedentFirstName": "John",
    "decedentMiddleName": "Michael",
    "decedentLastName": "Smith",
    "decedentAKA": "Johnny Smith",
    "decedentDOB": "03/15/1945",
    "decedentDOD": "01/10/2026",
    "decedentPlaceOfDeath": "Brooklyn, NY",
    "decedentStreet": "123 Main St",
    "decedentCity": "Brooklyn",
    "decedentCounty": "Kings",
    "decedentState": "NY",
    "decedentZip": "11201",
    "decedentCitizenship": "U.S.A.",

    # Petitioner
    "petitionerFirstName": "Mary",
    "petitionerMiddleName": "Jane",
    "petitionerLastName": "Smith",
    "petitionerStreet": "123 Main St",
    "petitionerCity": "Brooklyn",
    "petitionerState": "NY",
    "petitionerZip": "11201",
    "petitionerCitizenship": "U.S.A.",
    "petitionerRelationship": "Spouse",
    "petitionerInterest": "Distributee",
    "petitionerPhone": "",
    "petitionerIsAttorney": "No",

    # Case info
    "county": "Kings",
    "fileNo": "2026-1234",
    "proceedingType": "Administration",
    "lettersType": "Letters of Administration",
    "lettersTo": "Mary Jane Smith",

    # Property
    "personalPropertyValue": "50000",
    "improvedRealProperty": "250000",
    "unimprovedRealProperty": "",
    "realPropertyValue": "250000",
    "realPropertyDescription": "123 Main St, Brooklyn, NY 11201",
    "grossRents18mo": "",

    # Surviving relatives
    "survivingSpouse": True,
    "survivingChildren": False,
    "survivingIssue": False,
    "survivingParents": False,
    "survivingSiblings": False,
    "survivingGrandparents": False,
    "survivingAuntsUncles": False,
    "survivingFirstCousinsOnceRemoved": False,

    # Marital info
    "maritalStatus": "married",
    "spouseName": "Mary Jane Smith",

    # Distributees
    "distributees": [
        {
            "name": "Mary Jane Smith",
            "relationship": "Spouse",
            "address": "123 Main St, Brooklyn, NY 11201",
            "citizenship": "U.S.A.",
            "disposition": "waiver",
        },
    ],
    "soleDistributee": "Mary Jane Smith",

    # Heirship/deponent
    "deponentName": "Mary Jane Smith",
    "deponentAddress": "123 Main St, Brooklyn, NY 11201",
    "deponentRelationship": "Spouse",
    "yearsKnown": "30",
    "motherName": "Jane Doe",
    "motherDOD": "05/20/2010",
    "fatherName": "Robert Smith",
    "fatherDOD": "11/03/2015",

    # Cover letter
    "signer": "Jessica Wilson",
    "efileDate": "February 25, 2026",
    "enclosures": [
        "Petition for Administration (A-1)",
        "Affidavit of Assets & Liabilities (SCPA 805)",
        "Affidavit of Heirship",
        "Family Tree (FT-1)",
        "Death Certificate",
    ],

    # Debts
    "mortgageAmount": "150000",
    "funeralPaid": "8500",
    "funeralOutstanding": "",
    "miscDebts": "",
}

# ─── PROBATE SAMPLE DATA ───────────────────────────────────────────────────────

PROBATE_DATA = {
    # Decedent
    "decedentFirstName": "Eleanor",
    "decedentMiddleName": "Rose",
    "decedentLastName": "Williams",
    "decedentAKA": "Ellie Williams",
    "decedentDOB": "07/22/1938",
    "decedentDOD": "12/01/2025",
    "decedentPlaceOfDeath": "Manhattan, NY",
    "decedentStreet": "456 Park Ave",
    "decedentCity": "New York",
    "decedentCounty": "New York",
    "decedentState": "NY",
    "decedentZip": "10022",
    "decedentCitizenship": "U.S.A.",

    # Petitioner
    "petitionerFirstName": "David",
    "petitionerMiddleName": "A",
    "petitionerLastName": "Williams",
    "petitionerStreet": "789 Broadway",
    "petitionerCity": "New York",
    "petitionerState": "NY",
    "petitionerZip": "10003",
    "petitionerCitizenship": "U.S.A.",
    "petitionerRelationship": "Son",
    "petitionerInterest": "Executor",
    "petitionerPhone": "(917) 555-0123",
    "petitionerIsAttorney": "No",

    # Case info
    "county": "New York",
    "fileNo": "2025-5678",
    "proceedingType": "Probate",
    "lettersType": "Letters Testamentary",
    "lettersTo": "David A Williams",

    # Will info
    "willDate": "06/15/2020",
    "witness1": "Alice Brown",
    "witness2": "Robert Green",
    "codicilDate": "",
    "noOtherWill": "NONE",
    "selfProvingAffidavit": False,

    # Property
    "personalPropertyValue": "375000",
    "improvedRealProperty": "1200000",
    "unimprovedRealProperty": "50000",
    "realPropertyValue": "1250000",
    "realPropertyDescription": "456 Park Ave, Apt 12B, New York, NY 10022",
    "grossRents18mo": "45000",
    "otherAssets": "NONE",

    # Surviving relatives
    "survivingSpouse": False,
    "survivingChildren": True,
    "survivingIssue": False,
    "survivingParents": False,
    "survivingSiblings": True,
    "survivingGrandparents": False,
    "survivingAuntsUncles": False,
    "survivingFirstCousinsOnceRemoved": False,

    # Marital info
    "maritalStatus": "widowed",
    "spouseName": "Harold Williams",
    "priorSpouseDeathDate": "03/10/2022",

    # Distributees
    "distributees": [
        {
            "name": "David A Williams",
            "relationship": "Son",
            "address": "789 Broadway, New York, NY 10003",
            "citizenship": "U.S.A.",
            "disposition": "waiver",
        },
        {
            "name": "Sarah Williams-Chen",
            "relationship": "Daughter",
            "address": "55 Water St, Brooklyn, NY 11201",
            "citizenship": "U.S.A.",
            "disposition": "waiver",
        },
        {
            "name": "Thomas Williams",
            "relationship": "Son",
            "address": "100 Elm St, Queens, NY 11375",
            "citizenship": "U.S.A.",
        },
    ],
    "soleDistributee": "",
    "childrenNote": "The decedent had three children: David A Williams, Sarah Williams-Chen, and Thomas Williams, all of whom survive.",

    # Heirship/deponent
    "deponentName": "David A Williams",
    "deponentAddress": "789 Broadway, New York, NY 10003",
    "deponentRelationship": "Son",
    "yearsKnown": "55",
    "motherName": "Margaret Rose Davis",
    "motherDOD": "09/14/1998",
    "fatherName": "Harold James Williams",
    "fatherDOD": "03/10/2022",

    # Cover letter
    "signer": "Jessica Wilson",
    "efileDate": "February 20, 2026",
    "enclosures": [
        "Petition for Probate (P-1)",
        "Combined Verification, Oath and Designation",
        "Affidavit of Attesting Witness",
        "Death Certificate",
        "Original Will dated June 15, 2020",
    ],

    # Debts
    "mortgageAmount": "450000",
    "funeralPaid": "12000",
    "funeralOutstanding": "3500",
    "miscDebts": "Credit card (Chase): $4,200\nMedical bills: $8,750",
}

# ─── ANCILLARY SAMPLE DATA ─────────────────────────────────────────────────────

ANCILLARY_DATA = {
    # Decedent
    "decedentFirstName": "Margaret",
    "decedentMiddleName": "",
    "decedentLastName": "Thompson",
    "decedentAKA": "",
    "decedentDOB": "11/08/1950",
    "decedentDOD": "06/20/2025",
    "decedentPlaceOfDeath": "Hartford, CT",
    "decedentStreet": "200 Elm St",
    "decedentCity": "Hartford",
    "decedentCounty": "",
    "decedentState": "CT",
    "decedentZip": "06103",
    "decedentCitizenship": "U.S.A.",

    # Petitioner
    "petitionerFirstName": "James",
    "petitionerMiddleName": "",
    "petitionerLastName": "Thompson",
    "petitionerStreet": "200 Elm St",
    "petitionerCity": "Hartford",
    "petitionerState": "CT",
    "petitionerZip": "06103",
    "petitionerCitizenship": "U.S.A.",
    "petitionerRelationship": "Son",
    "petitionerInterest": "Distributee",
    "petitionerPhone": "",
    "petitionerIsAttorney": "No",

    # Case info
    "county": "Queens",
    "fileNo": "",
    "proceedingType": "Ancillary",
    "lettersType": "Letters of Administration",
    "lettersTo": "James Thompson",
    "foreignState": "Connecticut",
    "foreignLettersDate": "08/15/2025",
    "foreignLettersIssuedTo": "James Thompson",
    "foreignCourtName": "Probate Court, District of Hartford",
    "foreignBondAmount": "0",

    # Property (NY only)
    "personalPropertyValue": "85000",
    "improvedRealProperty": "",
    "unimprovedRealProperty": "",
    "realPropertyValue": "",
    "realPropertyDescription": "",
    "grossRents18mo": "",

    # Distributees
    "distributees": [
        {
            "name": "James Thompson",
            "relationship": "Son",
            "address": "200 Elm St, Hartford, CT 06103",
            "citizenship": "U.S.A.",
        },
    ],

    # Surviving relatives (for other forms if needed)
    "survivingSpouse": False,
    "survivingChildren": True,
    "survivingIssue": False,
    "survivingParents": False,
    "survivingSiblings": False,
    "survivingGrandparents": False,
    "survivingAuntsUncles": False,
    "survivingFirstCousinsOnceRemoved": False,

    "maritalStatus": "widowed",
    "spouseName": "Robert Thompson",
    "priorSpouseDeathDate": "01/05/2020",

    # Cover letter
    "signer": "Jessica Wilson",
    "efileDate": "February 25, 2026",
    "enclosures": [
        "Petition for Ancillary Administration (AA-1)",
        "Certified Copy of Foreign Letters",
        "Death Certificate",
    ],

    # Debts
    "mortgageAmount": "",
    "funeralPaid": "",
    "funeralOutstanding": "",
    "miscDebts": "",
}

# ─── ASSETS FOR ACCOUNTING / AUTH / INSTRUCTION LETTERS ─────────────────────────

SAMPLE_ASSETS = [
    {"institution": "Chase Bank", "category": "Checking", "value": "12500", "accountNumber": "****4567"},
    {"institution": "Chase Bank", "category": "Savings", "value": "45000", "accountNumber": "****4568"},
    {"institution": "Fidelity", "category": "IRA", "value": "185000", "accountNumber": "****9012"},
    {"institution": "MetLife", "category": "Life Insurance", "value": "100000", "accountNumber": "POL-555-123"},
    {"institution": "NY State Pension", "category": "Pension", "value": "32500", "accountNumber": "N/A"},
]


def save(filename, data_bytes):
    path = os.path.join(OUT_DIR, filename)
    with open(path, "wb") as f:
        f.write(data_bytes)
    size = len(data_bytes)
    print(f"  OK  {filename} ({size:,} bytes)")


errors = []
count = 0


def run(label, fn, *args, **kwargs):
    global count
    try:
        result = fn(*args, **kwargs)
        if isinstance(result, list):
            # generate_probate_docs returns list of (filename, bytes)
            for fname, data_bytes in result:
                save(fname, data_bytes)
                count += 1
        else:
            save(label, result)
            count += 1
    except Exception as e:
        errors.append((label, e))
        print(f"  FAIL  {label}: {e}")
        traceback.print_exc()


def main():
    global count
    print(f"Generating ALL test documents to {OUT_DIR}/\n")

    # ════════════════════════════════════════════════════════════════════════════
    print("── ADMINISTRATION PROCEEDING ──────────────────────────────────────")
    # ════════════════════════════════════════════════════════════════════════════

    run("Admin_01_Cover_Letter.docx",
        generate_cover_letter, ADMIN_DATA)

    run("Admin_02_Petition_A1.pdf",
        fill_administration_pdf, ADMIN_DATA)

    run("Admin_03_805_Affidavit.docx",
        generate_805, ADMIN_DATA)

    run("Admin_04_Heirship.docx",
        generate_heirship, ADMIN_DATA)

    run("Admin_05_FT1_Family_Tree.pdf",
        fill_ft1_pdf, ADMIN_DATA)

    run("Admin_06_Attorney_Cert.docx",
        generate_attorney_cert, ADMIN_DATA)

    # Waiver cover letter for the distributee
    for dist in ADMIN_DATA["distributees"]:
        if dist.get("disposition") == "waiver":
            safe_name = dist["name"].replace(" ", "_")
            run(f"Admin_07_Waiver_Cover_{safe_name}.docx",
                generate_waiver_cover, ADMIN_DATA, dist)

    # ════════════════════════════════════════════════════════════════════════════
    print("\n── PROBATE PROCEEDING ─────────────────────────────────────────────")
    # ════════════════════════════════════════════════════════════════════════════

    run("Probate_01_Cover_Letter.docx",
        generate_cover_letter, PROBATE_DATA)

    # Full probate packet (P-1 + Oath + Attesting Witness)
    run("probate_docs",
        generate_probate_docs, PROBATE_DATA)

    # Also test standalone fill_probate_pdf
    run("Probate_02b_Petition_P1_standalone.pdf",
        fill_probate_pdf, PROBATE_DATA)

    run("Probate_03_805_Affidavit.docx",
        generate_805, PROBATE_DATA)

    run("Probate_04_Heirship.docx",
        generate_heirship, PROBATE_DATA)

    run("Probate_05_FT1_Family_Tree.pdf",
        fill_ft1_pdf, PROBATE_DATA)

    run("Probate_06_Attorney_Cert.docx",
        generate_attorney_cert, PROBATE_DATA)

    # Waiver cover letters
    for dist in PROBATE_DATA["distributees"]:
        if dist.get("disposition") == "waiver":
            safe_name = dist["name"].replace(" ", "_")
            run(f"Probate_07_Waiver_Cover_{safe_name}.docx",
                generate_waiver_cover, PROBATE_DATA, dist)

    # ════════════════════════════════════════════════════════════════════════════
    print("\n── ANCILLARY PROCEEDING ───────────────────────────────────────────")
    # ════════════════════════════════════════════════════════════════════════════

    run("Ancillary_01_Cover_Letter.docx",
        generate_cover_letter, ANCILLARY_DATA)

    run("Ancillary_02_Petition_AA1.pdf",
        fill_ancillary_pdf, ANCILLARY_DATA)

    run("Ancillary_03_805_Affidavit.docx",
        generate_805, ANCILLARY_DATA)

    # ════════════════════════════════════════════════════════════════════════════
    print("\n── ACCOUNTING & LETTERS ───────────────────────────────────────────")
    # ════════════════════════════════════════════════════════════════════════════

    run("Accounting_Smith.xlsx",
        generate_accounting_excel, ADMIN_DATA, SAMPLE_ASSETS)

    # Auth letter (one per asset)
    for asset in SAMPLE_ASSETS[:2]:  # test 2 institutions
        safe_inst = asset["institution"].replace(" ", "_")
        run(f"Auth_Letter_{safe_inst}.docx",
            generate_auth_letter, ADMIN_DATA, asset)

    # Instruction letter — check mode
    run("Instruction_Letter_Chase_check.docx",
        generate_instruction_letter, ADMIN_DATA, SAMPLE_ASSETS[0], "check")

    # Instruction letter — transfer mode
    run("Instruction_Letter_Fidelity_transfer.docx",
        generate_instruction_letter, ADMIN_DATA, SAMPLE_ASSETS[2], "transfer")

    # ════════════════════════════════════════════════════════════════════════════
    print(f"\n{'=' * 60}")
    if errors:
        print(f"DONE with {len(errors)} ERROR(S) out of {count + len(errors)} documents:")
        for label, e in errors:
            print(f"  FAIL  {label}: {e}")
    else:
        print(f"SUCCESS! All {count} documents generated to {OUT_DIR}/")
    print(f"{'=' * 60}")


if __name__ == "__main__":
    main()
