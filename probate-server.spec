# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['runner.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['flask', 'werkzeug', 'jinja2', 'markupsafe', 'click', 'itsdangerous', 'blinker', 'docx', 'pypdf', 'pdfplumber', 'pdfminer', 'pdfminer.high_level', 'fitz', 'openpyxl', 'dotenv', 'anthropic', 'httpx', 'pytesseract', 'pdf2image', 'requests'],
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
