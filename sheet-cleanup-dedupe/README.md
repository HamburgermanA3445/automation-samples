# Spreadsheet cleanup and de-duplication

Cleans a contacts or customer list and removes duplicates. Writes a new
workbook and leaves the original file untouched.

## Run it

Requires Python 3.11 or newer.

```
pip install -r ../requirements.txt
python clean_contacts.py messy_contacts.csv
```

Output goes to `output/contacts_clean.xlsx` (change it with `--out`).
Input can be `.csv` or `.xlsx`.

## Expected columns

`Name, Company, Email, Phone, City, State, Last Contact`

Header names are matched loosely, so `email address`, `Phone Number` or
`last_contact` also work.

## What it changes

- Trims spaces at the ends of every cell and collapses double spaces.
- Names and cities typed in all lowercase or all caps are put in normal
  case (`o'connell` becomes `O'Connell`, `MIKE DELGADO` becomes
  `Mike Delgado`). Names already in mixed case, like `MacLeod`, are left as
  they are.
- Emails are lowercased. Emails that aren't valid (missing domain, a comma
  instead of a dot, spaces, double dots) are kept but flagged.
- US phone numbers are written as `(614) 555-0177`, keeping extensions
  (`x204`) and dropping a leading `+1`. Numbers without 10 digits are kept
  as entered and flagged.
- State names become two-letter codes (`Ohio` becomes `OH`).
- Dates in mixed formats (`9/22/2026`, `Oct 2, 2026`, `2026-10-02`) become
  real dates. `10-05-2026` is read as month-day-year.

## Duplicates

A row is treated as a duplicate when:

- its email matches an earlier row, ignoring case and spaces, or
- it has no matching email but the same name and phone number as an
  earlier row.

Completely empty rows are removed too.

The first row is kept. If the duplicate has information the kept row is
missing, it is copied over, and the more recent Last Contact date is kept.

## Output

- **Contacts**: the cleaned list, sorted by last name. `Source Row` is the
  row number in the original file. Flagged cells are highlighted yellow and
  explained in the `Issues` column.
- **Removed Rows**: every removed row as it was originally entered, why it
  was removed, and which row it was merged into.

## Sample

`messy_contacts.csv` has 26 rows with the usual problems: stray spaces,
all-caps names, people entered twice with different capitalization or
spacing, a row copied twice, an empty row, and a few bad emails and phone
numbers. The script keeps 20 contacts (5 flagged for review) and
removes 6 rows.
