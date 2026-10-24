# Automation samples

Small, working examples of spreadsheet and file automation jobs, built with
made-up businesses and data. Each folder has its own README, sample input
and the output it produced.

- `pdf-invoices-to-excel/`: PDF invoices from different vendors into one Excel workbook
- `csv-to-branded-invoices/`: a spreadsheet of orders into one PDF invoice per customer
- `sheet-cleanup-dedupe/`: cleaning and de-duplicating a messy contacts list
- `google-sheets-auto-mailer/`: Apps Script that emails rows marked "Ready" in a Google Sheet

## Setup

```
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

`make_gallery_images.py` rebuilds the PNGs in each `screenshots/` folder
from the sample outputs. Run the three Python samples first.
