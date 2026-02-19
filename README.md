# NY Surrogate's Court — Probate Assistant v2

Full document packet generator for probate, administration, and ancillary proceedings.

## What's in this folder

```
probate_v2/
├── app.py                          ← Flask server
├── generators.py                   ← Document generation logic
├── requirements.txt                ← Python dependencies
├── Probate-_NY_Court_Forms.pdf     ← Blank probate petition (P-1)
├── admin_ancil.pdf                 ← Blank ancillary admin petition (AA-1)
├── ADM_doc.docx                    ← Administration petition (A-1) — reference copy
├── static/
│   └── index.html                  ← Web app UI
└── templates/
    ├── 805_Affidavit_of_Assets_and_Liabilities_template.docx
    ├── Affidavit_of_Heirship_Full_Admin.docx
    ├── Waiver_cover_letter.docx
    └── newcertform_6_59_19_PM.docx ← Attorney certification
```

## Setup (one time only)

```bash
cd probate_v2
pip3 install -r requirements.txt
```

## Running

```bash
python3 app.py
```

Open **http://localhost:8080** in your browser.

## What the app generates

### For Probate:
- Court cover letter (auto-addressed by county, signed by selected staff)
- Filled Petition for Probate (P-1) PDF
- 805 Affidavit of Assets & Liabilities (Word)
- Attorney Certification (Word)
- Waiver cover letter per distributee who agreed to sign
- (Optional) Witness Affidavit if no self-proving affidavit

### For Administration:
- Court cover letter
- 805 Affidavit (Word)
- Affidavit of Heirship (Word)
- Attorney Certification (Word)
- Waiver cover letters per distributee
- Note: A-1 petition (Word) — fill manually from ADM_doc.docx

### For Ancillary Administration:
- Court cover letter
- Filled Ancillary Petition (AA-1) PDF
- 805 Affidavit (Word)
- Affidavit of Heirship (Word)
- Attorney Certification (Word)

## Counties supported
Bronx, Kings (Brooklyn), Nassau, New York, Queens, Richmond (Staten Island), Suffolk

## Signers
- Jessica Wilson, Esq.
- Robyn Foresta, Legal Assistant

## Stopping the server
Press `Ctrl+C` in the terminal.

## Next time you want to use it
```bash
cd ~/Downloads/probate_v2
python3 app.py
```
