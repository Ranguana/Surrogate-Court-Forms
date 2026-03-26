#!/bin/bash
# ─────────────────────────────────────────────────────────────────────
# bundle-python.sh — Create a self-contained Python bundle for Probate HQ
#
# Run this ONCE on your Mac before `npm run build`.
# It creates a python_bundle/ directory that electron-builder will
# embed in the DMG via extraResources.
#
# Requirements: Python 3.11+ installed on YOUR machine (brew or python.org)
# ─────────────────────────────────────────────────────────────────────
set -euo pipefail

BUNDLE_DIR="$(cd "$(dirname "$0")" && pwd)/python_bundle"
DEPS="flask python-docx pypdf pdfplumber pymupdf openpyxl python-dotenv anthropic pytesseract pdf2image requests"

echo "==> Creating python_bundle at: $BUNDLE_DIR"

# Clean previous bundle
rm -rf "$BUNDLE_DIR"

# Find system Python 3
PYTHON=""
for p in /opt/homebrew/bin/python3 /usr/local/bin/python3 /usr/bin/python3; do
  if [ -x "$p" ]; then
    PYTHON="$p"
    break
  fi
done

if [ -z "$PYTHON" ]; then
  echo "ERROR: No python3 found. Install Python 3 first."
  exit 1
fi

echo "==> Using system Python: $PYTHON"
echo "==> Python version: $($PYTHON --version)"

# Create venv
echo "==> Creating virtual environment..."
"$PYTHON" -m venv "$BUNDLE_DIR"

# Activate and install deps
echo "==> Installing dependencies..."
"$BUNDLE_DIR/bin/python3" -m pip install --upgrade pip --quiet
"$BUNDLE_DIR/bin/python3" -m pip install $DEPS --quiet

echo "==> Installed packages:"
"$BUNDLE_DIR/bin/python3" -m pip list --format=columns

# ── Make the venv relocatable ──────────────────────────────────────
# The venv hardcodes absolute paths in shebangs and pyvenv.cfg.
# We need to fix these so they work when the bundle lands in
# /Applications/Probate HQ.app/Contents/Resources/python_bundle/
echo "==> Making bundle relocatable..."

# Fix shebangs in bin/ scripts to use relative python
# (Electron spawns python3 by absolute path, so the shebang doesn't
#  matter for app.py, but pip/flask console scripts need it if you
#  ever run them manually from inside the bundle)
find "$BUNDLE_DIR/bin" -type f -exec grep -l "^#!.*python" {} \; | while read f; do
  sed -i '' "1s|^#!.*python.*|#!/usr/bin/env python3|" "$f" 2>/dev/null || true
done

# Remove pyvenv.cfg home reference (not needed at runtime)
if [ -f "$BUNDLE_DIR/pyvenv.cfg" ]; then
  sed -i '' '/^home = /d' "$BUNDLE_DIR/pyvenv.cfg"
fi

# ── Optional: bundle tesseract binary ──────────────────────────────
# If tesseract is installed via Homebrew, copy the binary + tessdata
# so OCR works without Homebrew on the target machine.
TESS=$(which tesseract 2>/dev/null || true)
if [ -n "$TESS" ]; then
  echo "==> Bundling Tesseract OCR..."
  TESS_DEST="$BUNDLE_DIR/tesseract"
  mkdir -p "$TESS_DEST/bin"

  # Copy binary
  cp "$TESS" "$TESS_DEST/bin/"

  # Copy tessdata (language files)
  TESSDATA=""
  if [ -d "/opt/homebrew/share/tessdata" ]; then
    TESSDATA="/opt/homebrew/share/tessdata"
  elif [ -d "/usr/local/share/tessdata" ]; then
    TESSDATA="/usr/local/share/tessdata"
  fi

  if [ -n "$TESSDATA" ]; then
    mkdir -p "$TESS_DEST/share/tessdata"
    # Only copy English to save space
    cp "$TESSDATA/eng.traineddata" "$TESS_DEST/share/tessdata/" 2>/dev/null || true
    cp "$TESSDATA/osd.traineddata" "$TESS_DEST/share/tessdata/" 2>/dev/null || true
  fi

  # Copy dylib dependencies
  mkdir -p "$TESS_DEST/lib"
  otool -L "$TESS" | grep "/opt/homebrew\|/usr/local" | awk '{print $1}' | while read dylib; do
    if [ -f "$dylib" ]; then
      cp "$dylib" "$TESS_DEST/lib/" 2>/dev/null || true
    fi
  done

  echo "   Tesseract bundled at: $TESS_DEST"
else
  echo "==> Tesseract not found — OCR will use Claude vision fallback"
fi

# ── Optional: bundle poppler (pdftoppm for pdf2image) ──────────────
PDFTOPPM=$(which pdftoppm 2>/dev/null || true)
if [ -n "$PDFTOPPM" ]; then
  echo "==> Bundling Poppler (pdftoppm)..."
  POPPLER_DEST="$BUNDLE_DIR/poppler/bin"
  mkdir -p "$POPPLER_DEST"
  cp "$PDFTOPPM" "$POPPLER_DEST/"

  # Copy dylib deps
  mkdir -p "$BUNDLE_DIR/poppler/lib"
  otool -L "$PDFTOPPM" | grep "/opt/homebrew\|/usr/local" | awk '{print $1}' | while read dylib; do
    if [ -f "$dylib" ]; then
      cp "$dylib" "$BUNDLE_DIR/poppler/lib/" 2>/dev/null || true
    fi
  done
  echo "   Poppler bundled"
else
  echo "==> Poppler not found — pdf2image may not work (Claude vision fallback)"
fi

# ── Clean up unnecessary files to reduce bundle size ───────────────
echo "==> Cleaning up to reduce size..."
# Remove pip cache
rm -rf "$BUNDLE_DIR/lib/python*/site-packages/pip" 2>/dev/null || true
rm -rf "$BUNDLE_DIR/lib/python*/site-packages/setuptools" 2>/dev/null || true
# Remove __pycache__
find "$BUNDLE_DIR" -type d -name "__pycache__" -exec rm -rf {} + 2>/dev/null || true
# Remove .dist-info (saves ~5-10MB)
# find "$BUNDLE_DIR" -type d -name "*.dist-info" -exec rm -rf {} + 2>/dev/null || true

BUNDLE_SIZE=$(du -sh "$BUNDLE_DIR" | awk '{print $1}')
echo ""
echo "==> Done! Bundle size: $BUNDLE_SIZE"
echo "==> Now run: npm run build"
