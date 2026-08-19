# tools/

One-off template build scripts (not used at runtime).

- `build_a8_template.py` — builds the fillable A-8 waiver template from the
  official NYSBA form (`knowledge/Administration Forms/Waiver and Consent.pdf`).
  NOTE: the live template `templates/Admin/Waiver of Consent and Renunciation.pdf`
  has since been hand-edited in Acrobat (caption). Do NOT re-run this over it
  without re-applying those edits.
- `rework_a9_template.py` — moves the A-9 corporate waiver's attorney block to
  two rows across the bottom (name/firm, address/phone/email).
