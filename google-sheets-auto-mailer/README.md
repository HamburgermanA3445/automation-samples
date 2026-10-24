# Google Sheets auto mailer

Apps Script for a Google Sheet. It sends one email per row where Status is
"Ready", fills the email from that row's columns, then sets Status to "Sent"
and writes the time into "Sent At". Rows with any other status are skipped,
so you can run it as often as you like without emailing anyone twice.

## Sheet setup

1. Make a sheet named `Contacts`. Row 1 is the header. It needs at least
   `Email`, `Status` and `Sent At`; add any other columns you want to use
   in the email. `example_contacts.csv` can be imported as a starting point
   (File > Import > Upload).
2. Make a second sheet named `Template`:
   - A1: `Subject`, B1: the subject line
   - A2: `Body`, B2: the email body (Ctrl+Enter makes a new line inside the cell)

   Put column names in double braces to fill them in, for example:

   ```
   B1: Your renewal with {{Company}} is coming up
   B2: Hi {{First Name}},

       Just a reminder that your service plan renews on {{Renewal Date}}.
       Reply to this email if anything has changed.

       Thanks,
       Sam
   ```

## Install the script

1. In the sheet, open Extensions > Apps Script.
2. Delete what is in `Code.gs` and paste in the contents of `Code.gs` from this folder. Save.
3. Reload the sheet. A **Mailer** menu appears after a few seconds.
4. The first time you use the menu, Google asks you to authorize the script
   to send email as you. Accept it.

## Use it

- **Mailer > Send me a test of the first "Ready" row** emails you a copy
  (subject starts with `[TEST]`) and changes nothing in the sheet. Do this first.
- **Mailer > Send to "Ready" rows** sends the real emails.

If a row fails (empty email, bad address), its Status becomes `Error` and the
reason is attached as a note on that cell. Fix it, set the status back to
`Ready` and run again.

## Run it on a schedule (optional)

In the Apps Script editor, open Triggers (clock icon) > Add Trigger, choose
`sendReadyRows`, event source "Time-driven", and pick an interval such as
every hour.

## Settings and limits

The constants at the top of `Code.gs` hold the sheet names, column names,
an optional sender name and `MAX_PER_RUN` (50). Free Gmail accounts can send
about 100 script emails a day and Google Workspace accounts about 1,500; the
script checks the remaining quota and stops before it runs out.

Before sending anything, the script checks that every `{{placeholder}}` in
the template matches a column header, so a typo stops the run instead of
sending emails with blanks in them.
