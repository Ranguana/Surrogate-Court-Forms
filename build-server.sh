#!/bin/bash
# ─────────────────────────────────────────────────────────────────
# build-server.sh — Freeze the Python runtime + deps into a binary
# The binary does NOT contain app.py — it loads it from disk at runtime.
# Run this before `npm run build`.
# ─────────────────────────────────────────────────────────────────
set -euo pipefail

echo "==> Installing PyInstaller (if needed)..."
pip3 install pyinstaller --quiet

echo "==> Installing app dependencies (so PyInstaller can bundle them)..."
pip3 install flask python-docx pypdf pdfplumber pymupdf openpyxl \
  python-dotenv anthropic pytesseract pdf2image requests --quiet

echo "==> Freezing runner into standalone binary..."
pyinstaller \
  --onefile \
  --name probate-server \
  --hidden-import=flask \
  --hidden-import=werkzeug \
  --hidden-import=jinja2 \
  --hidden-import=markupsafe \
  --hidden-import=click \
  --hidden-import=itsdangerous \
  --hidden-import=blinker \
  --hidden-import=docx \
  --hidden-import=pypdf \
  --hidden-import=pdfplumber \
  --hidden-import=pdfminer \
  --hidden-import=pdfminer.high_level \
  --hidden-import=fitz \
  --hidden-import=openpyxl \
  --hidden-import=dotenv \
  --hidden-import=anthropic \
  --hidden-import=httpx \
  --hidden-import=pytesseract \
  --hidden-import=pdf2image \
  --hidden-import=requests \
  runner.py

echo ""
echo "==> Done! Binary at: dist/probate-server"
echo "==> Size: $(du -sh dist/probate-server | awk '{print $1}')"
echo ""
echo "Now run: npm run build"
