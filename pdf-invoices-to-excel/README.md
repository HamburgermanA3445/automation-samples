# PDF invoices to Excel

Reads every PDF in a folder, pulls out the invoice details and writes them
to one Excel workbook. Files it can't read are listed instead of skipped
silently.

The workbook has three sheets:

- **Invoices**: one row per invoice. File, vendor, invoice number, invoice
  date, subtotal, tax, total, number of line items, and a Check column.
- **Line Items**: one row per line on each invoice (description, qty, unit
  price, amount).
- **Problems**: files that couldn't be parsed or whose numbers don't add up.

The same messages are written to `output/extract_log.txt` with a timestamp.

## Run it

Requires Python 3.11 or newer.

```
pip install -r ../requirements.txt
python extract_invoices.py sample_invoices output/invoices.xlsx
```

Both arguments are optional; those are the defaults.

## What it expects

Text-based PDFs, the kind accounting software and invoicing tools export.
It looks for common labels rather than fixed positions, so it handles
different layouts:

- invoice number: `Invoice No`, `Invoice #`, `Invoice Number`, `Inv. Number`
- date: `Date:` or `Invoice Date:` in formats like `09/02/2026`,
  `2026-09-02` or `September 2, 2026` (a `Due Date` is ignored)
- totals: lines starting with `Subtotal`/`Sub-total`, `Tax`/`Sales Tax`,
  and `Total`/`Total Due`/`Amount Due`
- line items: the rows between the table header (a line containing "Qty")
  and the subtotal, with qty, unit price and amount at the end of the line,
  or qty at the start

The vendor name is taken from the first line of the page.

## Checks

For each invoice it confirms that the line items add up to the subtotal and
that subtotal + tax equals the total. Anything off shows in the Check column
and on the Problems sheet so it can be looked at by hand.

Scanned invoices (images with no text layer) are logged as needing OCR. In
the sample folder, `scan_20260904_0915.pdf` is one of these on purpose.

## Sample data

`sample_invoices/` has 7 invoices from three fictional vendors, each with a
different layout (boxed table, striped rows with item codes, plain text with
qty first), plus the scanned one. They were made with
`make_sample_invoices.py`. The result of running the script on them is in
`output/`.

## Adapting it to your invoices

Each vendor's layout is a little different. If a vendor uses a label not
listed above, it is usually a one-line change to the matching pattern near
the top of `extract_invoices.py`.
