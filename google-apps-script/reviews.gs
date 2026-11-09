// Google Apps Script backend for the website's client reviews.
// 1) Create a Google Sheet, copy its ID, and paste it below.
// 2) Deploy this project as a Web app: Execute as "Me", access "Anyone".
const SPREADSHEET_ID = 'PASTE_YOUR_GOOGLE_SHEET_ID_HERE';
const SHEET_NAME = 'Reviews';
const HEADERS = ['Created At', 'Name', 'Rating', 'Review', 'Status'];

function getReviewsSheet_() {
  if (!SPREADSHEET_ID || SPREADSHEET_ID.indexOf('PASTE_') === 0) {
    throw new Error('Set SPREADSHEET_ID in the Apps Script first.');
  }
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
  return sheet;
}

function safeCell_(value, maxLength) {
  let text = String(value || '').trim().slice(0, maxLength);
  // Prevent user-submitted text from being interpreted as a spreadsheet formula.
  if (/^[=+\-@]/.test(text)) text = "'" + text;
  return text;
}

function jsonp_(callback, data) {
  const body = JSON.stringify(data).replace(/</g, '\\u003c').replace(/>/g, '\\u003e').replace(/&/g, '\\u0026');
  const safeCallback = /^[A-Za-z_$][0-9A-Za-z_$\.]*$/.test(callback || '') ? callback : '';
  const output = safeCallback ? safeCallback + '(' + body + ');' : JSON.stringify(data);
  return ContentService.createTextOutput(output)
    .setMimeType(safeCallback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}

function doGet(e) {
  const callback = e && e.parameter ? e.parameter.callback : '';
  try {
    const sheet = getReviewsSheet_();
    const rows = sheet.getDataRange().getValues();
    const reviews = [];
    let total = 0;
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (String(row[4]).toUpperCase() !== 'APPROVED') continue;
      const rating = Math.max(1, Math.min(5, Number(row[2]) || 0));
      if (!rating) continue;
      total += rating;
      reviews.push({
        name: String(row[1] || ''),
        rating: rating,
        review: String(row[3] || ''),
        date: row[0] instanceof Date
          ? Utilities.formatDate(row[0], Session.getScriptTimeZone(), 'yyyy-MM-dd')
          : ''
      });
    }
    reviews.reverse();
    const average10 = reviews.length ? Math.round((total / reviews.length) * 2 * 10) / 10 : 0;
    return jsonp_(callback, { ok: true, count: reviews.length, average10: average10, reviews: reviews });
  } catch (err) {
    return jsonp_(callback, { ok: false, count: 0, average10: 0, reviews: [], error: String(err.message || err) });
  }
}

function doPost(e) {
  let requestId = '';
  try {
    const p = e && e.parameter ? e.parameter : {};
    requestId = safeCell_(p.requestId, 80);

    if (p.website) return submissionPage_('Spam check failed.', false, requestId);
    if (p.action !== 'submit') return submissionPage_('Invalid request.', false, requestId);
    if (p.consent !== 'yes') return submissionPage_('Consent is required.', false, requestId);

    const rating = Number(p.rating);
    const review = safeCell_(p.review, 700);
    const name = safeCell_(p.name, 60);

    if (!Number.isInteger(rating) || rating < 1 || rating > 5 || review.length < 3) {
      return submissionPage_('Please provide a valid rating and review.', false, requestId);
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const sheet = getReviewsSheet_();
      sheet.appendRow([new Date(), name, rating, review, 'APPROVED']);
    } finally {
      lock.releaseLock();
    }

    return submissionPage_('Review published successfully.', true, requestId);
  } catch (err) {
    return submissionPage_('Submission failed: ' + String(err.message || err), false, requestId);
  }
}

function submissionPage_(message, ok, requestId) {
  const payload = JSON.stringify({
    type: 'reviews-submit-result',
    ok: Boolean(ok),
    requestId: String(requestId || ''),
    message: String(message || '')
  }).replace(/</g, '\\u003c');
  const safeMessage = String(message || '').replace(/[<>&"]/g, '');
  return ContentService.createTextOutput(
    '<!doctype html><html><head><meta charset="utf-8"></head><body>' +
    safeMessage + '<script>window.top.postMessage(' + payload + ', "*");</script></body></html>'
  ).setMimeType(ContentService.MimeType.HTML);
}

// Run this once from the Apps Script editor after setting SPREADSHEET_ID
// to authorize access and create the Reviews tab with its header row.
function setupReviewsSheet() {
  getReviewsSheet_();
}
