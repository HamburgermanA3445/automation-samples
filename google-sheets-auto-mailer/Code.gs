/**
 * Sends a templated email to every row whose Status is "Ready",
 * then marks the row "Sent" with a timestamp.
 *
 * Sheet layout (see README.md):
 *   "Contacts" sheet: header row with at least Email, Status and Sent At.
 *                     Any other column can be used in the template.
 *   "Template" sheet: B1 = subject, B2 = body. Use {{Column Name}} placeholders.
 */

const CONTACTS_SHEET = 'Contacts';
const TEMPLATE_SHEET = 'Template';
const EMAIL_COLUMN = 'Email';
const STATUS_COLUMN = 'Status';
const SENT_AT_COLUMN = 'Sent At';
const READY = 'ready';            // compared case-insensitively
const SENDER_NAME = '';           // leave blank to use your account name
const MAX_PER_RUN = 50;           // stays well under the daily quota

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Mailer')
    .addItem('Send to "Ready" rows', 'sendReadyRows')
    .addItem('Send me a test of the first "Ready" row', 'sendTestToMe')
    .addToUi();
}

function sendReadyRows() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) {
    notify_('Another send is already running. Try again in a minute.');
    return;
  }
  try {
    const ctx = loadContext_();
    const quota = MailApp.getRemainingDailyQuota();
    const limit = Math.min(MAX_PER_RUN, quota);
    let sent = 0;
    let failed = 0;

    for (let i = 0; i < ctx.rows.length && sent < limit; i++) {
      const row = ctx.rows[i];
      if (String(row[ctx.col.status]).trim().toLowerCase() !== READY) continue;

      const sheetRow = i + 2; // +1 for header, +1 because sheets are 1-based
      const statusCell = ctx.sheet.getRange(sheetRow, ctx.col.status + 1);
      const to = String(row[ctx.col.email]).trim();

      try {
        if (!to) throw new Error('Email is empty');
        const subject = fillTemplate_(ctx.subject, ctx.headers, row);
        const body = fillTemplate_(ctx.body, ctx.headers, row);
        sendEmail_(to, subject, body);

        statusCell.setValue('Sent').clearNote();
        ctx.sheet.getRange(sheetRow, ctx.col.sentAt + 1).setValue(new Date());
        sent++;
      } catch (err) {
        statusCell.setValue('Error').setNote(err.message);
        failed++;
      }
      // Write each result right away so a timeout never causes a double send.
      SpreadsheetApp.flush();
    }

    let msg = `Sent ${sent} email(s).`;
    if (failed) msg += ` ${failed} row(s) marked "Error" (hover the cell to see why).`;
    if (sent === limit && limit < MAX_PER_RUN) {
      msg += ' Daily email quota reached.';
    } else if (sent === MAX_PER_RUN) {
      msg += ` Stopped at ${MAX_PER_RUN} per run. Run it again for the rest.`;
    }
    notify_(msg);
  } finally {
    lock.releaseLock();
  }
}

/** Sends the first "Ready" row to your own address without changing the sheet. */
function sendTestToMe() {
  const ctx = loadContext_();
  const row = ctx.rows.find(r => String(r[ctx.col.status]).trim().toLowerCase() === READY);
  if (!row) {
    notify_('No rows with Status "Ready".');
    return;
  }
  const me = Session.getActiveUser().getEmail();
  sendEmail_(me,
             '[TEST] ' + fillTemplate_(ctx.subject, ctx.headers, row),
             fillTemplate_(ctx.body, ctx.headers, row));
  notify_(`Test sent to ${me}.`);
}

function loadContext_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(CONTACTS_SHEET);
  const template = ss.getSheetByName(TEMPLATE_SHEET);
  if (!sheet) throw new Error(`No sheet named "${CONTACTS_SHEET}".`);
  if (!template) throw new Error(`No sheet named "${TEMPLATE_SHEET}".`);

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  const required = { email: EMAIL_COLUMN, status: STATUS_COLUMN, sentAt: SENT_AT_COLUMN };
  const col = {};
  const missing = [];
  for (const [key, name] of Object.entries(required)) {
    col[key] = headers.indexOf(name);
    if (col[key] === -1) missing.push(name);
  }
  if (missing.length) {
    throw new Error(`Missing column(s) in "${CONTACTS_SHEET}": ${missing.join(', ')}`);
  }

  const subject = String(template.getRange('B1').getValue());
  const body = String(template.getRange('B2').getValue());
  if (!subject || !body) throw new Error('Template sheet needs a subject in B1 and a body in B2.');

  // Catch typos like {{Frist Name}} before anything is sent.
  const unknown = placeholders_(subject + body).filter(p => !headers.includes(p));
  if (unknown.length) {
    throw new Error('Template uses columns that do not exist: ' + unknown.join(', '));
  }

  return { sheet, headers, col, rows: values.slice(1), subject, body };
}

function placeholders_(text) {
  const names = [];
  text.replace(/\{\{\s*([^}]+?)\s*\}\}/g, (_, name) => names.push(name));
  return names;
}

function fillTemplate_(text, headers, row) {
  const tz = Session.getScriptTimeZone();
  return text.replace(/\{\{\s*([^}]+?)\s*\}\}/g, (_, name) => {
    const value = row[headers.indexOf(name)];
    if (value instanceof Date) return Utilities.formatDate(value, tz, 'MMMM d, yyyy');
    return value === null || value === undefined ? '' : String(value);
  });
}

function sendEmail_(to, subject, body) {
  const html = body
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/\n/g, '<br>');
  const options = { htmlBody: html };
  if (SENDER_NAME) options.name = SENDER_NAME;
  MailApp.sendEmail(to, subject, body, options);
}

function notify_(message) {
  // The toast shows in the open sheet; the log keeps a record for trigger runs.
  console.log(message);
  SpreadsheetApp.getActive().toast(message, 'Mailer', 8);
}
