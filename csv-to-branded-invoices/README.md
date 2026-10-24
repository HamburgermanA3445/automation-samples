# CSV to branded invoices

Takes a CSV or Excel file of order lines and makes one PDF invoice per
customer, with the business's name, initials logo, color, tax, due date and
payment terms. Also writes `invoice_summary.csv` (invoice number, customer,
email, total, due date, file name), which is handy for emailing them out or
importing into bookkeeping.

## Run it

Requires Python 3.11 or newer.

```
pip install -r ../requirements.txt
python make_invoices.py orders.csv --date 2026-10-19
```

Options:

- `--date` invoice date, `YYYY-MM-DD`. Defaults to today.
- `--business` path to the business details file. Defaults to `business.json`.
- `--out` output folder. Defaults to `output/`.

An `.xlsx` file works the same way as a CSV.

## Input columns

`order_date, customer, email, address, item, qty, unit_price`

One row per item ordered. A customer can appear on many rows and dates;
all of their rows go on one invoice. The address is written as
`street, city, state zip` and split onto two lines on the invoice.

Small mistakes in the data are handled: customer names that differ only in
capitalization or extra spaces are treated as the same customer, and prices
can be written with or without a `$`. Rows missing a customer, qty or price
are skipped and their row numbers printed.

## Business details

Everything specific to the business lives in `business.json`: name,
initials for the logo, address, contact info, brand color, tax rate and
label, payment terms and the next invoice number. After a run, the script
prints the next free invoice number; update `next_invoice_number` so the
next batch carries on from there.

## Sample

`orders.csv` holds two weeks of wholesale orders for Harbor Street Bakery
(fictional), including a couple of messy rows ("copper kettle cafe " in
lowercase with a trailing space, a price typed as `$6.50`). `output/` has
the four invoices it produced.
