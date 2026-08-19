"""A-9 corporate waiver: move the attorney block from the lower-right stack
to two rows across the bottom under the notary block:
  Name of Attorney | Firm Name
  Address          | Telephone | Email
"""
import fitz, sys, shutil
SRC = sys.argv[1]
doc = fitz.open(SRC)
pg = doc[0]

OLD = ["Name of Attorney_2", "1_6", "2_6", "Telephone Number_3"]
for w in list(pg.widgets()):
    if w.field_name in OLD:
        pg.delete_widget(w)

# white-out old right-side labels/underlines and the "A-9" label
pg.add_redact_annot(fitz.Rect(330, 500, 560, 642))
pg.add_redact_annot(fitz.Rect(34, 641, 56, 656))
pg.add_redact_annot(fitz.Rect(500, 760, 600, 785))  # "Page 17 of 18" footer
pg.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)

H = 13
ROW1 = 672   # underline y for row 1
ROW2 = 702   # underline y for row 2
row1 = [("Name of Attorney", "Name of Attorney_2", 36, 296),
        ("Firm Name",        "Firm Name_2",        306, 576)]
row2 = [("Address",          "Address_2",          36, 296),
        ("Telephone",        "Telephone Number_3", 306, 426),
        ("Email",            "Email_2",            436, 576)]

def add_field(name, rect):
    w = fitz.Widget()
    w.field_name = name
    w.field_type = fitz.PDF_WIDGET_TYPE_TEXT
    w.rect = rect
    w.text_fontsize = 9
    w.text_font = "Helv"
    w.border_width = 0
    w.fill_color = None
    pg.add_widget(w); w.update()

for base, row in ((ROW1, row1), (ROW2, row2)):
    for label, fname, x0, x1 in row:
        pg.draw_line(fitz.Point(x0, base), fitz.Point(x1, base), width=0.6)
        pg.insert_text(fitz.Point(x0, base + 9), label, fontsize=7.5, fontname="helv")
        add_field(fname, fitz.Rect(x0, base - H - 1, x1, base - 1))

pg.insert_text(fitz.Point(36, ROW2 + 28), "A-9", fontsize=9, fontname="helv")

doc.save(SRC + ".new", garbage=3, deflate=True)
doc.close()
shutil.move(SRC + ".new", SRC)
print("done")
