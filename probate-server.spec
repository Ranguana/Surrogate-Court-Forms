# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=[('static', 'static'), ('templates', 'templates'), ('Accounting', 'Accounting'), ('generators.py', '.'), ('cases.json', '.'), ('contacts.json', '.'), ('favicon.svg', '.'), ('.env', '.'), ('Probate-_NY_Court_Forms.pdf', '.'), ('admin_ancil.pdf', '.'), ('Petition_for_Non-Domciliary_Letters_of_Admin.pdf', '.'), ('smart_intake_prompt.md', '.')],
    hiddenimports=['flask', 'dotenv', 'docx', 'pypdf', 'pdfplumber', 'fitz', 'openpyxl', 'anthropic', 'pytesseract', 'pdf2image', 'requests', 'pdfplumber.utils', 'pdfplumber.page'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='probate-server',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
