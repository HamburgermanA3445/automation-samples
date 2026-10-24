"""Generate the sample invoice PDFs used to demo extract_invoices.py.

Three fictional vendors, each with its own layout, all billing the same
property management company. One extra file is an image-only "scan" so
the extractor has something it can't read.
"""

from datetime import date, timedelta
from pathlib import Path

import pypdfium2 as pdfium
from PIL import Image, ImageFilter
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

OUT_DIR = Path(__file__).parent / "sample_invoices"
PAGE_W, PAGE_H = letter

BILL_TO = [
    "Maple Row Property Management",
    "Attn: Accounts Payable",
    "41 Maple Row, Suite 200",
    "Dayton, OH 45402",
]


def money(value):
    return f"${value:,.2f}"


def totals(items, tax_rate):
    subtotal = round(sum(qty * price for _, qty, price in items), 2)
    tax = round(subtotal * tax_rate, 2)
    return subtotal, tax, round(subtotal + tax, 2)


# --- Northfield Plumbing: boxed table, dark header row --------------------

def northfield_invoice(path, number, inv_date, items):
    tax_rate = 0.0725
    subtotal, tax, total = totals(items, tax_rate)
    c = canvas.Canvas(str(path), pagesize=letter)
    c.setTitle(f"Invoice {number}")

    c.setFont("Helvetica-Bold", 18)
    c.drawString(54, 730, "Northfield Plumbing & Heating")
    c.setFont("Helvetica", 9.5)
    for i, line in enumerate(["2290 Valley Pike", "Dayton, OH 45404",
                              "(937) 555-0148  |  office@northfieldplumbing.com"]):
        c.drawString(54, 714 - i * 13, line)

    c.setFont("Helvetica-Bold", 22)
    c.drawRightString(558, 728, "INVOICE")
    c.setFont("Helvetica", 10)
    due = inv_date + timedelta(days=30)
    meta = [("Invoice No:", number),
            ("Invoice Date:", inv_date.strftime("%m/%d/%Y")),
            ("Due Date:", due.strftime("%m/%d/%Y")),
            ("Terms:", "Net 30")]
    for i, (label, value) in enumerate(meta):
        y = 690 - i * 14
        c.drawRightString(480, y, label)
        c.drawRightString(558, y, value)

    c.setFont("Helvetica-Bold", 10)
    c.drawString(54, 620, "Bill To")
    c.setFont("Helvetica", 10)
    for i, line in enumerate(BILL_TO):
        c.drawString(54, 605 - i * 13, line)

    cols = [54, 340, 400, 480, 558]
    y = 530
    c.setFillColor(colors.HexColor("#27384a"))
    c.rect(cols[0], y - 6, cols[-1] - cols[0], 20, stroke=0, fill=1)
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(cols[0] + 6, y, "Description")
    c.drawRightString(cols[2] - 6, y, "Qty")
    c.drawRightString(cols[3] - 6, y, "Unit Price")
    c.drawRightString(cols[4] - 6, y, "Amount")
    c.setFillColor(colors.black)
    c.setFont("Helvetica", 10)
    c.setStrokeColor(colors.HexColor("#9aa5b1"))
    top = y - 6
    for desc, qty, price in items:
        y -= 22
        c.drawString(cols[0] + 6, y, desc)
        c.drawRightString(cols[2] - 6, y, f"{qty:g}")
        c.drawRightString(cols[3] - 6, y, money(price))
        c.drawRightString(cols[4] - 6, y, money(qty * price))
        c.line(cols[0], y - 7, cols[-1], y - 7)
    for x in cols:
        c.line(x, top, x, y - 7)

    y -= 34
    for label, value, bold in [("Subtotal", subtotal, False),
                               ("Sales Tax (7.25%)", tax, False),
                               ("Total Due", total, True)]:
        c.setFont("Helvetica-Bold" if bold else "Helvetica", 10.5)
        c.drawRightString(480, y, label)
        c.drawRightString(552, y, money(value))
        y -= 16

    c.setFont("Helvetica", 9)
    c.drawString(54, 90, "Please make checks payable to Northfield Plumbing & Heating "
                         "and include the invoice number.")
    c.drawString(54, 77, "Questions about this invoice? Call Denise at (937) 555-0148.")
    c.save()


# --- Bright Line Office Supply: item codes, striped rows, long date -------

def brightline_invoice(path, number, inv_date, po, items):
    tax_rate = 0.075
    subtotal, tax, total = totals(items, tax_rate)
    c = canvas.Canvas(str(path), pagesize=letter)
    c.setTitle(f"Invoice {number}")

    accent = colors.HexColor("#c8541c")
    c.setFillColor(accent)
    c.rect(0, 752, PAGE_W, 40, stroke=0, fill=1)
    c.setFillColor(colors.black)

    c.setFont("Helvetica-Bold", 16)
    c.drawString(50, 718, "Bright Line Office Supply")
    c.setFont("Helvetica", 9)
    c.drawString(50, 704, "880 Commerce Dr, Columbus, OH 43219  -  orders@brightlineoffice.com")

    c.setFont("Helvetica-Bold", 13)
    c.drawString(50, 668, "Invoice")
    c.setFont("Helvetica", 10)
    c.drawString(50, 652, f"Invoice # {number}")
    c.drawString(50, 638, f"Date: {inv_date:%B} {inv_date.day}, {inv_date.year}")
    c.drawString(50, 624, f"PO Number: {po}")

    c.setFont("Helvetica-Bold", 10)
    c.drawString(330, 668, "Sold To:")
    c.setFont("Helvetica", 10)
    for i, line in enumerate(BILL_TO[:1] + BILL_TO[2:]):
        c.drawString(330, 652 - i * 14, line)

    y = 570
    c.setFont("Helvetica-Bold", 9.5)
    c.drawString(50, y, "Item")
    c.drawString(120, y, "Description")
    c.drawRightString(400, y, "Qty")
    c.drawRightString(475, y, "Price")
    c.drawRightString(560, y, "Line Total")
    c.setStrokeColor(accent)
    c.setLineWidth(1.5)
    c.line(50, y - 6, 560, y - 6)
    c.setLineWidth(1)

    c.setFont("Helvetica", 9.5)
    for i, (code_desc, qty, price) in enumerate(items):
        code, desc = code_desc.split(" ", 1)
        y -= 20
        if i % 2 == 0:
            c.setFillColor(colors.HexColor("#f4eee9"))
            c.rect(50, y - 6, 510, 20, stroke=0, fill=1)
            c.setFillColor(colors.black)
        c.drawString(54, y, code)
        c.drawString(120, y, desc)
        c.drawRightString(400, y, f"{qty:g}")
        c.drawRightString(475, y, f"{price:,.2f}")
        c.drawRightString(556, y, f"{qty * price:,.2f}")

    y -= 40
    c.setFont("Helvetica", 10)
    c.drawRightString(475, y, "Sub-total")
    c.drawRightString(556, y, f"{subtotal:,.2f}")
    c.drawRightString(475, y - 15, "Tax 7.5%")
    c.drawRightString(556, y - 15, f"{tax:,.2f}")
    c.setFont("Helvetica-Bold", 11)
    c.drawRightString(475, y - 33, "TOTAL")
    c.drawRightString(556, y - 33, money(total))

    c.setFont("Helvetica-Oblique", 8.5)
    c.drawString(50, 60, "Returns accepted within 30 days with receipt. Payment due 30 days from invoice date.")
    c.save()


# --- Greenway Lawn & Landscape: plain text, qty first, no table lines -----

def greenway_invoice(path, number, inv_date, items):
    tax_rate = 0.0
    subtotal, tax, total = totals(items, tax_rate)
    c = canvas.Canvas(str(path), pagesize=letter)
    c.setTitle(f"Invoice {number}")

    c.setFillColor(colors.HexColor("#2f6b3a"))
    c.setFont("Times-Bold", 20)
    c.drawString(60, 725, "Greenway Lawn & Landscape")
    c.setFillColor(colors.black)
    c.setFont("Times-Roman", 10.5)
    c.drawString(60, 708, "PO Box 1187, Kettering, OH 45429   (937) 555-0162")

    c.setFont("Times-Roman", 11)
    c.drawString(60, 670, f"Inv. Number: {number}")
    c.drawString(60, 655, f"Date: {inv_date.strftime('%m/%d/%Y')}")
    c.drawString(60, 640, "Customer: Maple Row Property Management")
    c.drawString(60, 625, "Service address: 41 Maple Row, Dayton OH")

    y = 585
    c.setFont("Times-Bold", 11)
    c.drawString(60, y, "Qty")
    c.drawString(100, y, "Service")
    c.drawRightString(470, y, "Rate")
    c.drawRightString(550, y, "Amount")
    c.setFont("Times-Roman", 11)
    for desc, qty, price in items:
        y -= 18
        c.drawString(60, y, f"{qty:g}")
        c.drawString(100, y, desc)
        c.drawRightString(470, y, f"{price:,.2f}")
        c.drawRightString(550, y, f"{qty * price:,.2f}")

    y -= 30
    c.drawRightString(470, y, "Subtotal")
    c.drawRightString(550, y, f"{subtotal:,.2f}")
    c.drawRightString(470, y - 16, "Tax")
    c.drawRightString(550, y - 16, f"{tax:,.2f}")
    c.setFont("Times-Bold", 11)
    c.drawRightString(470, y - 34, "Amount Due")
    c.drawRightString(550, y - 34, money(total))

    c.setFont("Times-Italic", 10)
    c.drawString(60, 90, "Landscape services are not subject to Ohio sales tax. Due on receipt.")
    c.save()


def make_scanned_copy(src_pdf, dst_pdf):
    """Turn a PDF into an image-only PDF, like a phone or copier scan."""
    pdf = pdfium.PdfDocument(str(src_pdf))
    img = pdf[0].render(scale=1.4).to_pil().convert("L")
    pdf.close()
    img = img.rotate(0.8, expand=True, fillcolor=255).filter(ImageFilter.GaussianBlur(0.6))
    img = Image.eval(img, lambda p: int(p * 0.92 + 12))
    img.save(dst_pdf, "PDF", resolution=100)


def main():
    OUT_DIR.mkdir(exist_ok=True)

    northfield_invoice(OUT_DIR / "Northfield_NP-10482.pdf", "NP-10482", date(2026, 8, 14), [
        ("Service call - 112 Elm St, Unit 3B", 1, 95.00),
        ("Kitchen faucet cartridge (Moen 1225)", 1, 38.50),
        ("Labor (hours)", 1.5, 85.00),
    ])
    northfield_invoice(OUT_DIR / "Northfield_NP-10517.pdf", "NP-10517", date(2026, 9, 2), [
        ("Emergency call - 41 Maple Row basement", 1, 175.00),
        ("Sump pump, 1/3 HP submersible", 1, 219.00),
        ("Check valve and PVC fittings", 1, 42.75),
        ("Labor (hours)", 3, 85.00),
    ])
    northfield_invoice(OUT_DIR / "Northfield_NP-10561.pdf", "NP-10561", date(2026, 9, 29), [
        ("Water heater flush - 18 Linden Ct", 2, 89.00),
        ("Anode rod replacement", 1, 64.00),
    ])

    brightline_invoice(OUT_DIR / "BrightLine_Invoice_0883.pdf", "BL-2026-0883", date(2026, 9, 8), "MR-4471", [
        ("PAP-5001 Copy paper, letter, 10-ream case", 3, 46.99),
        ("TON-2211 Toner cartridge, HP 58A", 2, 112.40),
        ("ENV-1040 #10 envelopes, box of 500", 1, 24.85),
        ("BND-0300 3-ring binders, 1.5 in (12 pk)", 1, 31.20),
    ])
    brightline_invoice(OUT_DIR / "BrightLine_Invoice_0921.pdf", "BL-2026-0921", date(2026, 10, 6), "MR-4502", [
        ("PAP-5001 Copy paper, letter, 10-ream case", 2, 46.99),
        ("LBL-7760 Shipping labels, 2x4 in (1000)", 1, 38.65),
        ("CHR-9105 Mesh task chair, black", 1, 189.00),
    ])

    greenway_invoice(OUT_DIR / "greenway inv 2026-117.pdf", "2026-117", date(2026, 9, 30), [
        ("Weekly mowing and trimming (Sept)", 4, 65.00),
        ("Shrub pruning, front beds", 1, 140.00),
        ("Fall cleanup discount (returning customer)", 1, -40.00),
    ])
    greenway_invoice(OUT_DIR / "greenway inv 2026-124.pdf", "2026-124", date(2026, 10, 15), [
        ("Weekly mowing and trimming (Oct 1-15)", 2, 65.00),
        ("Leaf removal and haul-away", 1, 225.00),
        ("Hardwood mulch, installed (cu yd)", 3, 58.00),
    ])

    # A scanned copy of an older Greenway invoice. No text layer, so the
    # extractor should log it instead of guessing.
    tmp = OUT_DIR / "_tmp_scan_source.pdf"
    greenway_invoice(tmp, "2026-109", date(2026, 8, 31), [
        ("Weekly mowing and trimming (Aug)", 4, 65.00),
        ("Weed control, parking strip", 1, 55.00),
    ])
    make_scanned_copy(tmp, OUT_DIR / "scan_20260904_0915.pdf")
    tmp.unlink()

    print(f"Wrote {len(list(OUT_DIR.glob('*.pdf')))} PDFs to {OUT_DIR.name}/")


if __name__ == "__main__":
    main()
