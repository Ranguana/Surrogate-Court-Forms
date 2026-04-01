# Probate HQ - Development Notes

## What This Is
Probate HQ is a **packaged Electron desktop app** (.dmg) for NYS Surrogate's Court filings.
It is NOT a dev-mode web app. It gets installed on user machines from a built DMG.

## Architecture
- **Frontend:** Electron app (main.js) serving static/index.html
- **Backend:** Bundled Flask/Python server (app.py) runs as a subprocess inside the packaged app
- **The .env file (with ANTHROPIC_API_KEY) is bundled inside the DMG at build time**
- **Supabase** is used for case storage/sync across machines
- PDF generation uses pdftk and custom field mappings

## Build Process
- Built with electron-builder, output goes to `dist/`
- Python backend is compiled with PyInstaller into a standalone binary (`probate-server`)
- The binary and .env are included as extraResources in the package

## Known Issues & Fixes
- **Smart Intake spinning forever:** Added API key validation (returns error immediately if missing), 120s timeout on Claude API calls, and 2.5min frontend fetch timeout. If it still spins, the bundled Python server likely isn't starting — check macOS Gatekeeper, permissions, or Console.app for errors.
- **First launch on a new Mac:** User must right-click > Open to bypass Gatekeeper. Double-clicking may silently block the embedded executables.

## Key Files
- `main.js` — Electron main process
- `app.py` — Flask backend (PDF generation, Smart Intake, Supabase sync)
- `static/index.html` — All frontend UI and JS
- `templates/` — PDF templates for court forms
- `field_mappings.py` — Maps form fields to PDF template fields
- `package.json` — Electron + build config
- `.env` — Anthropic API key (bundled in builds)
