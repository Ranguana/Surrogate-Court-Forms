"""Build a fillable A-8 template from the official NYSBA form (no fields)."""
import fitz, sys

SRC, DST = sys.argv[1], sys.argv[2]
doc = fitz.open(SRC)

def text(pg, name, x0, y0, x1, y1, size=9):
    w = fitz.Widget()
    w.field_name = name
    w.field_type = fitz.PDF_WIDGET_TYPE_TEXT
    w.rect = fitz.Rect(x0, y0, x1, y1)
    w.text_fontsize = size
    w.text_font = "Helv"
    w.border_width = 0
    w.fill_color = None
    pg.add_widget(w); w.update()

def check(pg, name, x0, y0, x1, y1):
    w = fitz.Widget()
    w.field_name = name
    w.field_type = fitz.PDF_WIDGET_TYPE_CHECKBOX
    w.rect = fitz.Rect(x0, y0, x1, y1)
    w.border_width = 0
    w.fill_color = None
    pg.add_widget(w); w.update()

p1, p2 = doc[0], doc[1]

# Strip the NYSBA / Matthew Bender footer text on both pages (page number stays)
for pg in (p1, p2):
    pg.add_redact_annot(fitz.Rect(30, 745, 200, 760))   # "NYSBA's Surrogate's Court Form A-8 (3/05)"
    pg.add_redact_annot(fitz.Rect(330, 745, 580, 760))  # "© 2015 Matthew Bender & Co., ..."
    pg.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)

# Strip the NYSBA logo + "New York State Surrogate's Court / NYS Bar Association
# Official OCA Forms" header on page 1 (keeps "Form A-8" + title on the right)
p1.add_redact_annot(fitz.Rect(30, 30, 582, 73))   # whole header banner incl. "Form A-8" + rules
p1.apply_redactions(images=fitz.PDF_REDACT_IMAGE_REMOVE)

# ── Page 1 ──────────────────────────────────────────────────────
text(p1, "county",        107, 90,  250, 103)          # COUNTY OF ____
text(p1, "estate_of",     40,  134, 295, 150, 10)      # decedent name (line under "ESTATE OF")
text(p1, "aka",           64,  152, 295, 166)          # a/k/a ____
text(p1, "file_no",       383, 219, 575, 232)          # File No.
text(p1, "court_county",  363, 281, 537, 294)          # Surrogate's Court of ____ County
check(p1, "cb_letters",        72.4, 327.0, 83.2, 337.8)
check(p1, "cb_letters_limits", 72.4, 345.5, 83.2, 356.3)
check(p1, "cb_limited",        72.4, 364.0, 83.2, 374.8)
text(p1, "be_issued_to",  100, 391, 576, 404)
check(p1, "cb_bond_dispensed", 72.4, 440.4, 83.2, 451.2)
check(p1, "cb_bond_amount",    72.4, 478.9, 83.2, 489.7)
text(p1, "bond_amount",   234, 477, 353, 490)
text(p1, "dated",         72,  590, 244, 603)
text(p1, "signature",     324, 588, 576, 601)
text(p1, "print_name",    324, 617, 576, 630)
text(p1, "street",        36,  646, 576, 659)
text(p1, "city",          36,  673, 250, 686)
text(p1, "state",         253, 673, 350, 686)
text(p1, "zip",           353, 673, 430, 686)
text(p1, "country",       434, 673, 576, 686)
text(p1, "relationship",  36,  701, 576, 714)

# ── Page 2 ──────────────────────────────────────────────────────
text(p2, "notary_state",  93,  46,  226, 59)
text(p2, "notary_county", 104, 59,  226, 72)
text(p2, "notary_date",   90,  117, 233, 130)
text(p2, "notary_appeared", 399, 117, 577, 130)
text(p2, "atty_name",     326, 275, 577, 288)
text(p2, "atty_firm",     38,  303, 314, 316)
text(p2, "atty_phone",    326, 303, 577, 316)
# Address line: official form runs full width; split it so Email fits at right
p2.add_redact_annot(fitz.Rect(314, 340, 580, 350))   # trim right part of the Address underline
p2.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)
p2.draw_line(fitz.Point(37.4, 345.2), fitz.Point(314.6, 345.2), width=0.5)
p2.draw_line(fitz.Point(325.4, 345.2), fitz.Point(577.4, 345.2), width=0.5)
p2.insert_text(fitz.Point(325.8, 354.5), "Email", fontsize=8, fontname="helv")
text(p2, "atty_address",  38,  332, 314, 345)
text(p2, "atty_email",    326, 332, 577, 345)

doc.save(DST, garbage=3, deflate=True)
print("fields:", sum(1 for pg in doc for _ in pg.widgets()))
