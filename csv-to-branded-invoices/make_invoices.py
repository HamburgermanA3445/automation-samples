"""Turn a CSV or Excel file of orders into one PDF invoice per customer.

Usage:
    python make_invoices.py orders.csv --date 2026-10-19
    python make_invoices.py orders.xlsx --business business.json --out output

Business details (name, colors, tax rate, payment terms, next invoice
number) come from business.json so nothing is hard-coded in the script.
"""

import argparse
import csv
import json
import re
from collections import Counter
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.utils import simpleSplit
from reportlab.pdfgen import canvas

HERE = Path(__file__).parent
REQUIRED_COLUMNS = ["order_date", "customer", "email", "address", "item", "qty", "unit_price"]

PAGE_W, PAGE_H = letter
LEFT, RIGHT = 54, PAGE_W - 54


# --- Reading and cleaning the orders ---------------------------------------

def clean_text(value):
    return re.sub(r"\s+", " ", str(value)).strip() if pd.notna(value) else ""


def to_number(value):
    text = clean_text(value).replace("$", "").replace(",", "")
    return float(text) if text else None


def load_orders(path):
    if path.suffix.lower() in (".xlsx", ".xls"):
        df = pd.read_excel(path, dtype=str)
    else:
        df = pd.read_csv(path, dtype=str)
    df.columns = [c.strip().lower().replace(" ", "_") for c in df.columns]
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise SystemExit(f"{path.name} is missing columns: {', '.join(missing)}")

    orders, skipped = [], []
    for row_no, row in enumerate(df.to_dict("records"), start=2):
        qty, price = to_number(row["qty"]), to_number(row["unit_price"])
        customer = clean_text(row["customer"])
        if not customer or qty is None or price is None:
            skipped.append(row_no)
            continue
        orders.append({
            "date": pd.to_datetime(clean_text(row["order_date"])).date(),
            "customer": customer,
            "email": clean_text(row["email"]).lower(),
            "address": clean_text(row["address"]),
            "item": clean_text(row["item"]),
            "qty": qty,
            "unit_price": price,
        })
    if skipped:
        print(f"Skipped rows with no customer, qty or price: {skipped}")
    return orders


def group_by_customer(orders):
    """Group orders by customer, ignoring case and stray spaces in the name."""
    groups = {}
    for order in orders:
        groups.setdefault(order["customer"].casefold(), []).append(order)

    customers = []
    for rows in groups.values():
        # Use the spelling that appears most often, e.g. "Copper Kettle Cafe"
        # rather than "copper kettle cafe".
        name = Counter(r["customer"] for r in rows).most_common(1)[0][0]
        first = rows[0]
        customers.append({"name": name, "email": first["email"], "address": first["address"],
                          "lines": sorted(rows, key=lambda r: r["date"])})
    return sorted(customers, key=lambda c: c["name"].casefold())


# --- Drawing the PDF --------------------------------------------------------

def money(value):
    return f"${value:,.2f}"


def draw_logo(c, x, y, initials, color):
    radius = 26
    c.setFillColor(color)
    c.circle(x + radius, y, radius, stroke=0, fill=1)
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 19)
    c.drawCentredString(x + radius, y - 7, initials)
    c.setFillColor(colors.black)


def draw_header(c, biz, brand, invoice_no, inv_date, due_date, customer):
    draw_logo(c, LEFT, PAGE_H - 84, biz["initials"], brand)
    c.setFont("Helvetica-Bold", 16)
    c.drawString(LEFT + 66, PAGE_H - 78, biz["name"])
    c.setFont("Helvetica", 9)
    c.setFillColor(colors.HexColor("#555555"))
    contact = biz["address"] + [f"{biz['phone']}  |  {biz['email']}"]
    for i, line in enumerate(contact):
        c.drawString(LEFT + 66, PAGE_H - 93 - i * 11.5, line)
    c.setFillColor(colors.black)

    c.setFillColor(brand)
    c.setFont("Helvetica-Bold", 24)
    c.drawRightString(RIGHT, PAGE_H - 80, "INVOICE")
    c.setFillColor(colors.black)
    meta = [("Invoice no.", invoice_no),
            ("Invoice date", inv_date.strftime("%b %d, %Y")),
            ("Due date", due_date.strftime("%b %d, %Y"))]
    for i, (label, value) in enumerate(meta):
        y = PAGE_H - 102 - i * 14
        c.setFont("Helvetica", 9.5)
        c.drawRightString(RIGHT - 90, y, label)
        c.setFont("Helvetica-Bold", 9.5)
        c.drawRightString(RIGHT, y, value)

    y = PAGE_H - 178
    c.setStrokeColor(colors.HexColor("#dddddd"))
    c.line(LEFT, y + 18, RIGHT, y + 18)
    c.setFont("Helvetica-Bold", 9)
    c.setFillColor(brand)
    c.drawString(LEFT, y, "BILL TO")
    c.setFillColor(colors.black)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(LEFT, y - 16, customer["name"])
    c.setFont("Helvetica", 9.5)
    street, _, city = customer["address"].partition(", ")
    for i, line in enumerate([street, city, customer["email"]]):
        c.drawString(LEFT, y - 30 - i * 12, line)
    return y - 84


COLS = [(LEFT, "Date", "l"), (LEFT + 80, "Item", "l"), (RIGHT - 170, "Qty", "r"),
        (RIGHT - 85, "Unit price", "r"), (RIGHT, "Amount", "r")]


def draw_table_header(c, y, brand):
    c.setFillColor(brand)
    c.rect(LEFT, y - 6, RIGHT - LEFT, 20, stroke=0, fill=1)
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 9.5)
    for x, label, align in COLS:
        if align == "l":
            c.drawString(x + 6, y, label)
        else:
            c.drawRightString(x - 6, y, label)
    c.setFillColor(colors.black)
    return y - 22


def build_invoice(path, biz, customer, invoice_no, inv_date):
    brand = colors.HexColor(biz["brand_color"])
    tint = colors.Color(brand.red, brand.green, brand.blue, alpha=0.08)
    due_date = inv_date + timedelta(days=biz["payment_days"])

    subtotal = round(sum(l["qty"] * l["unit_price"] for l in customer["lines"]), 2)
    tax = round(subtotal * biz["tax_rate"], 2)
    total = round(subtotal + tax, 2)

    c = canvas.Canvas(str(path), pagesize=letter)
    c.setTitle(f"{biz['name']} invoice {invoice_no}")
    c.setAuthor(biz["name"])

    y = draw_header(c, biz, brand, invoice_no, inv_date, due_date, customer)
    y = draw_table_header(c, y, brand)
    c.setFont("Helvetica", 9.5)
    for i, line in enumerate(customer["lines"]):
        if y < 150:  # leave room for totals; continue on a new page
            c.showPage()
            y = draw_table_header(c, PAGE_H - 72, brand)
            c.setFont("Helvetica", 9.5)
        if i % 2 == 1:
            c.setFillColor(tint)
            c.rect(LEFT, y - 6, RIGHT - LEFT, 19, stroke=0, fill=1)
            c.setFillColor(colors.black)
        amount = line["qty"] * line["unit_price"]
        c.drawString(COLS[0][0] + 6, y, line["date"].strftime("%b %d"))
        c.drawString(COLS[1][0] + 6, y, line["item"])
        c.drawRightString(COLS[2][0] - 6, y, f"{line['qty']:g}")
        c.drawRightString(COLS[3][0] - 6, y, money(line["unit_price"]))
        c.drawRightString(COLS[4][0] - 6, y, money(amount))
        y -= 19

    c.setStrokeColor(colors.HexColor("#cccccc"))
    c.line(LEFT, y + 6, RIGHT, y + 6)
    y -= 14
    c.setFont("Helvetica", 10)
    for label, value in [("Subtotal", subtotal), (biz["tax_label"], tax)]:
        c.drawRightString(RIGHT - 90, y, label)
        c.drawRightString(RIGHT - 6, y, money(value))
        y -= 16
    y -= 10
    c.setFillColor(brand)
    c.rect(RIGHT - 230, y - 9, 230, 26, stroke=0, fill=1)
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 11)
    c.drawRightString(RIGHT - 90, y, "Total due")
    c.drawRightString(RIGHT - 6, y, money(total))
    c.setFillColor(colors.black)

    y -= 50
    c.setFont("Helvetica-Bold", 9.5)
    c.drawString(LEFT, y, f"Payment due by {due_date.strftime('%B %d, %Y')}")
    c.setFont("Helvetica", 9)
    for line in simpleSplit(biz["payment_terms"], "Helvetica", 9, RIGHT - LEFT):
        y -= 12
        c.drawString(LEFT, y, line)

    c.setFont("Helvetica", 8.5)
    c.setFillColor(colors.HexColor("#777777"))
    c.drawCentredString(PAGE_W / 2, 40, f"Thank you for your order. Questions about this invoice: "
                                        f"{biz['email']} or {biz['phone']}")
    c.save()
    return total, due_date


def slugify(text):
    text = text.lower().replace("'", "").replace("&", "and")
    return re.sub(r"[^a-z0-9]+", "-", text).strip("-")


def main():
    parser = argparse.ArgumentParser(description="Make one PDF invoice per customer.")
    parser.add_argument("orders", type=Path, help="CSV or Excel file of order lines")
    parser.add_argument("--business", type=Path, default=HERE / "business.json")
    parser.add_argument("--out", type=Path, default=HERE / "output")
    parser.add_argument("--date", help="invoice date as YYYY-MM-DD (default: today)")
    args = parser.parse_args()

    biz = json.loads(args.business.read_text(encoding="utf-8"))
    inv_date = datetime.strptime(args.date, "%Y-%m-%d").date() if args.date else date.today()
    customers = group_by_customer(load_orders(args.orders))
    args.out.mkdir(parents=True, exist_ok=True)

    summary = []
    number = biz["next_invoice_number"]
    for customer in customers:
        invoice_no = f"{biz['invoice_prefix']}{number}"
        pdf_name = f"{invoice_no}_{slugify(customer['name'])}.pdf"
        total, due_date = build_invoice(args.out / pdf_name, biz, customer, invoice_no, inv_date)
        summary.append([invoice_no, customer["name"], customer["email"],
                        inv_date.isoformat(), due_date.isoformat(), f"{total:.2f}", pdf_name])
        print(f"{invoice_no}  {customer['name']:<28} {len(customer['lines']):>2} lines  {money(total):>10}")
        number += 1

    with open(args.out / "invoice_summary.csv", "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["invoice_no", "customer", "email", "invoice_date", "due_date", "total", "file"])
        writer.writerows(summary)
    print(f"{len(summary)} invoices written to {args.out.name}/. "
          f"Next invoice number: {number} (update business.json)")


if __name__ == "__main__":
    main()
