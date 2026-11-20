/**
 * RSVP backend for the 40th anniversary evening.
 *
 * Setup: create an Apps Script project, paste this file as Code.gs and add an HTML file
 * named "Index" with the contents of Index.html. Run setup() once: it creates the
 * response spreadsheet in your Drive and logs its URL. Then
 * Deploy -> New deployment -> Web app (Execute as: Me, Who has access: Anyone).
 *
 * Responses go to the SHEET_NAME sheet. If the same first + last name submits again,
 * the existing row is updated instead of adding a duplicate.
 * The SUMMARY_NAME sheet shows live totals.
 */

// Optional: ID of an existing Google Sheet (from its URL). When empty, setup() creates one
// and remembers its ID in Script Properties. A sheet-bound script uses its own sheet.
const SPREADSHEET_ID = '';
const SPREADSHEET_TITLE = 'אישורי הגעה – 40 שנות נישואין';
const SHEET_NAME = 'אישורי הגעה';
const SUMMARY_NAME = 'סיכום';
const HEADERS = ['עודכן', 'שם פרטי', 'שם משפחה', 'מגיעים?', 'כמה', 'הערות'];
const ALLOWED = ['כן', 'לא', 'אולי'];

/** Serves the landing page. */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('אישור הגעה – 40 שנות נישואין')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** JSON endpoint, used only when the page is hosted outside Apps Script. */
function doPost(e) {
  let result;
  try {
    result = submitRsvp(JSON.parse(e.postData.contents));
  } catch (err) {
    result = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Validates one response and writes it to the sheet. Called from the page. */
function submitRsvp(data) {
  data = data || {};
  if (data.website) return { ok: true }; // honeypot: silently drop bots

  const firstName = clean_(data.firstName, 60);
  const lastName = clean_(data.lastName, 60);
  const attending = clean_(data.attending, 10);
  const notes = clean_(data.notes, 1000);
  let guests = '';

  if (!firstName || !lastName) throw new Error('Missing name');
  if (ALLOWED.indexOf(attending) === -1) throw new Error('Invalid attending value');
  if (attending !== 'לא') {
    guests = Math.min(30, Math.max(1, parseInt(data.guests, 10) || 1));
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet_();
    const row = [new Date(), firstName, lastName, attending, guests, notes];

    const key = normalize_(firstName) + '|' + normalize_(lastName);
    const last = sheet.getLastRow();
    let target = -1;
    if (last > 1) {
      const names = sheet.getRange(2, 2, last - 1, 2).getValues();
      for (let i = 0; i < names.length; i++) {
        if (normalize_(names[i][0]) + '|' + normalize_(names[i][1]) === key) {
          target = i + 2;
          break;
        }
      }
    }

    if (target > 0) {
      sheet.getRange(target, 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }
  } finally {
    lock.releaseLock();
  }
  return { ok: true };
}

/** Returns the responses sheet, creating it (and the summary sheet) on first use. */
function getSheet_() {
  const ss = spreadsheet_();
  if (!ss) throw new Error('Run setup() first');
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME, 0);
    sheet.setRightToLeft(true);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#f3e3e6');
    sheet.getRange('A:A').setNumberFormat('dd/MM/yyyy HH:mm');
    sheet.setColumnWidth(6, 320);
    createSummary_(ss);
  }
  return sheet;
}

function createSummary_(ss) {
  if (ss.getSheetByName(SUMMARY_NAME)) return;
  const s = ss.insertSheet(SUMMARY_NAME, 1);
  s.setRightToLeft(true);
  const ref = "'" + SHEET_NAME + "'";
  s.getRange('A1:C1').setValues([['', 'משפחות / רשומות', 'אנשים']]).setFontWeight('bold');
  s.getRange('A2:C4').setFormulas([
    ['מגיעים', '=COUNTIF(' + ref + '!D:D,"כן")', '=SUMIF(' + ref + '!D:D,"כן",' + ref + '!E:E)'],
    ['אולי', '=COUNTIF(' + ref + '!D:D,"אולי")', '=SUMIF(' + ref + '!D:D,"אולי",' + ref + '!E:E)'],
    ['לא מגיעים', '=COUNTIF(' + ref + '!D:D,"לא")', ''],
  ]);
  s.getRange('A2:A4').setFontWeight('bold');
  s.setColumnWidths(1, 3, 140);
}

/** Run once manually from the editor to create the sheets before sharing the link. */
function setup() {
  let ss = spreadsheet_();
  if (!ss) {
    ss = SpreadsheetApp.create(SPREADSHEET_TITLE);
    PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
  } else if (SPREADSHEET_TITLE) {
    ss.rename(SPREADSHEET_TITLE);
  }
  getSheet_();
  // Remove the empty default tab if it is still there.
  const def = ss.getSheets().filter(function (sh) {
    return sh.getName() !== SHEET_NAME && sh.getName() !== SUMMARY_NAME && sh.getLastRow() === 0;
  });
  def.forEach(function (sh) { ss.deleteSheet(sh); });
  Logger.log('Responses sheet: ' + ss.getUrl());
}

function spreadsheet_() {
  const id = SPREADSHEET_ID || PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  return SpreadsheetApp.getActiveSpreadsheet(); // null for a standalone script before setup()
}

function clean_(value, max) {
  const s = String(value == null ? '' : value).trim().slice(0, max);
  // Prevent guest input from being interpreted as a spreadsheet formula.
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function normalize_(value) {
  return String(value || '').trim().replace(/\s+/g, ' ').toLowerCase();
}
