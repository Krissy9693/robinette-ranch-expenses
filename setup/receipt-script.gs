/**
 * Receipt Tracker -> Google Sheet
 *
 * This script lives inside the Google Sheet (Extensions > Apps Script).
 * The receipt page sends each receipt here, and this script adds a row to
 * the "Receipts" tab (created automatically the first time).
 *
 * Receipt photos are saved to a Google Drive folder named below, and the
 * row links to the photo. They stay private to the owner of this sheet.
 */

const TAB_NAME = 'Receipts';
const PHOTO_FOLDER_NAME = 'Receipt Tracker - Photos';
const MAX_TEXT = 5000;
const HEADERS = ['Saved', 'Receipt date', 'Vendor', 'Amount', 'Payment method', 'Tax category', 'Other (what it was for)', 'Notes', 'Receipt photo'];

// Where the page's random Dallas quote comes from. The list is re-read every 6 hours.
const QUOTES_PAGE = 'https://dallastvshow.com/dallas-quotes/';
const QUOTES_CACHE_SECONDS = 6 * 60 * 60;

// The receipt page asks for "?action=quote" to get a random Dallas quote.
// Opening the web app address in a browser shows a quick "is it on?" check.
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'quote') {
    try {
      const quotes = dallasQuotes();
      if (!quotes.length) throw new Error('No quotes found');
      return reply(Object.assign({ ok: true }, quotes[Math.floor(Math.random() * quotes.length)]));
    } catch (err) {
      return reply({ ok: false, error: String(err && err.message || err) });
    }
  }
  return ContentService.createTextOutput('The receipt tracker is connected and running.');
}

function dallasQuotes() {
  const cache = CacheService.getScriptCache();
  const saved = cache.get('dallas-quotes');
  if (saved) return JSON.parse(saved);
  const response = UrlFetchApp.fetch(QUOTES_PAGE, { muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) throw new Error('Quotes page unavailable');
  const quotes = parseQuotes(response.getContentText());
  if (quotes.length) cache.put('dallas-quotes', JSON.stringify(quotes), QUOTES_CACHE_SECONDS);
  return quotes;
}

// The quotes page has one section per character: a heading with their name,
// then their lines as list items (and sometimes a highlighted quote).
function parseQuotes(html) {
  const quotes = [];
  html.split('<section class="bio-section"').slice(1).forEach(function (part) {
    part = part.split('</section>')[0];
    const heading = part.match(/<h2[^>]*>([\s\S]*?)<\/h2>/);
    if (!heading) return;
    const who = plainText(heading[1]);
    const lines = /<(li|blockquote)\b[^>]*>([\s\S]*?)<\/\1>/g;
    let match;
    while ((match = lines.exec(part))) {
      const text = plainText(match[2].replace(/<span[\s\S]*?<\/span>/g, ''));
      const quoted = text.match(/^[“"]([\s\S]+)[”"]/); // leaves out notes after the quote
      if (!quoted) continue;
      const quote = quoted[1].trim();
      if (quote.length >= 5 && quote.length <= 220) quotes.push({ quote: quote, who: who });
    }
  });
  return quotes;
}

function plainText(html) {
  return html
    .replace(/<[^>]+>/g, '')
    .replace(/&#0?39;|&rsquo;|&lsquo;|&#8217;|&#8216;/g, "'")
    .replace(/&quot;|&ldquo;|&rdquo;|&#8220;|&#8221;/g, '"')
    .replace(/&amp;/g, '&')
    .replace(/&nbsp;/g, ' ')
    .replace(/&[a-z]+;|&#\d+;/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  let locked = false;
  try {
    lock.waitLock(20000); // one receipt at a time, so rows never collide
    locked = true;

    const data = JSON.parse(e.postData.contents);
    if (data.hp_field) return reply({ ok: true }); // filled only by spam bots

    const amount = Number(data.amount);
    if (!data.date || !data.vendor || !data.category || isNaN(amount)) throw new Error('Missing details');

    const photoLink = data.photo ? savePhoto(data.photo, data.date, data.vendor) : '';
    const sheet = receiptsTab();
    sheet.appendRow([
      new Date(),
      clean(data.date),
      clean(data.vendor),
      amount,
      clean(data.payment),
      clean(data.category),
      clean(data.other),
      clean(data.notes),
      photoLink
    ]);
    sheet.getRange(sheet.getLastRow(), 4).setNumberFormat('$#,##0.00');
    return reply({ ok: true });
  } catch (err) {
    return reply({ ok: false, error: String(err && err.message || err) });
  } finally {
    if (locked) lock.releaseLock();
  }
}

function reply(body) {
  return ContentService.createTextOutput(JSON.stringify(body)).setMimeType(ContentService.MimeType.JSON);
}

function receiptsTab() {
  const book = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = book.getSheetByName(TAB_NAME);
  if (!sheet) {
    sheet = book.insertSheet(TAB_NAME);
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Keep what was typed as plain text. Without this, text that starts with
// = + - or @ would be treated by Sheets as a formula.
function clean(value) {
  if (value === undefined || value === null) return '';
  let text = String(value).slice(0, MAX_TEXT);
  if (/^[=+\-@]/.test(text)) text = "'" + text;
  return text;
}

function savePhoto(photo, date, vendor) {
  if (!photo || !/^image\//.test(photo.type || '')) return '';
  const name = String(date || '').replace(/[^\d-]/g, '') + ' ' +
    String(vendor || 'Receipt').replace(/[^\w .'&-]/g, '').slice(0, 60) + '.jpg';
  const blob = Utilities.newBlob(Utilities.base64Decode(photo.data), photo.type, name);
  return photoFolder().createFile(blob).getUrl();
}

function photoFolder() {
  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty('PHOTO_FOLDER_ID');
  if (savedId) {
    try {
      return DriveApp.getFolderById(savedId);
    } catch (err) {
      // folder was deleted; make a new one below
    }
  }
  const folder = DriveApp.createFolder(PHOTO_FOLDER_NAME);
  props.setProperty('PHOTO_FOLDER_ID', folder.getId());
  return folder;
}
