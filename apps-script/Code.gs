/** Edutel Year-End Function 2026: RSVP back end (Apps Script bound to the RSVP Sheet). */

const SHEET_NAME = 'RSVPs';
const SUMMARY_NAME = 'Summary';

// Must match DIET_OPTIONS in index.html.
const DIET_OPTIONS = [
  'No dietary requirements',
  'Vegetarian',
  'Vegan',
  'Halaal',
  'Food allergies or intolerances',
];
const DIET_ALLERGY = 'Food allergies or intolerances';

// The RSVPs tab's columns: [field name the page sends, column heading].
// Columns are found by heading, so organisers can reorder them or add their own.
const COLUMNS = [
  ['submitted', 'Submitted'],
  ['updated', 'Last updated'],
  ['name', 'Full name'],
  ['email', 'Work email'],
  ['phone', 'Cellphone'],
  ['attending', 'Attending'],
  ['dietary', 'Dietary requirement'],
  ['dietaryDetails', 'Allergy / intolerance details'],
];

// After this many wrong passcodes, the organiser view locks for 15 minutes.
const MAX_WRONG_PASSCODES = 20;

/** Creates the RSVPs and Summary tabs and remembers this Sheet. Safe to re-run: existing responses are kept. */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  props.setProperty('SHEET_ID', ss.getId());
  ss.setSpreadsheetTimeZone('Africa/Johannesburg');

  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    const sheets = ss.getSheets();
    // A new Sheet starts with an empty "Sheet1": reuse it rather than adding a tab.
    sheet = sheets.length === 1 && sheets[0].getLastRow() === 0
      ? sheets[0].setName(SHEET_NAME)
      : ss.insertSheet(SHEET_NAME, 0);
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, COLUMNS.length).setValues([COLUMNS.map(c => c[1])]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.getRange('A:B').setNumberFormat('ddd d mmm yyyy HH:mm');
  }
  buildSummary_(ss, columnIndex_(sheet));

  if (!props.getProperty('PASSCODE')) {
    console.warn('PASSCODE script property is not set.');
  }
  console.log('Setup done for ' + ss.getUrl());
}

/** Opening the web app link in a browser shows this: a quick "is it running?" check. */
function doGet() {
  return json_({ ok: true, service: 'Edutel RSVP' });
}

/** The page sends everything here as JSON: RSVPs, and the organiser view's requests. */
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    return json_(body.action === 'list' ? listRsvps_(body.passcode) : saveRsvp_(body));
  } catch (err) {
    if (err.userFacing) return json_({ ok: false, error: err.message });
    console.error(err);
    return json_({ ok: false, error: 'Something went wrong on our side. Please try again in a moment.' });
  }
}

function saveRsvp_(data) {
  if (data.hp) return { ok: true }; // the hidden spam-trap field was filled in: a bot, so ignore it
  const entry = clean_(data);

  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // one save at a time, so RSVPs sent at the same moment can't collide
  try {
    const sheet = rsvpSheet_();
    const col = columnIndex_(sheet);
    const now = new Date();
    const row = findRow_(sheet, col.email, entry.email);

    if (row) {
      // Same email again: update their answer in place.
      Object.keys(entry).forEach(key => sheet.getRange(row, col[key] + 1).setValue(asText_(entry[key])));
      sheet.getRange(row, col.updated + 1).setValue(now);
      return { ok: true, updated: true };
    }

    const values = new Array(sheet.getLastColumn()).fill('');
    Object.keys(entry).forEach(key => { values[col[key]] = asText_(entry[key]); });
    values[col.submitted] = now;
    values[col.updated] = now;
    sheet.appendRow(values);
    return { ok: true, updated: false };
  } finally {
    lock.releaseLock();
  }
}

/** Checks and tidies what the page sent. Anything wrong stops here with a message for the guest. */
function clean_(data) {
  const text = (value, max) => String(value == null ? '' : value).trim().slice(0, max);
  const entry = {
    name: text(data.name, 100),
    email: text(data.email, 254).toLowerCase(),
    phone: text(data.phone, 30),
    attending: data.attending === 'Yes' || data.attending === 'No' ? data.attending : '',
    dietary: text(data.dietary, 60),
    dietaryDetails: text(data.dietaryDetails, 300),
  };

  if (!entry.name) fail_('Please tell us your full name.');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(entry.email)) fail_('Please enter a valid work email.');
  if (entry.phone && !isPhone_(entry.phone)) fail_('Please check your cellphone number, or leave it blank.');
  if (!entry.attending) fail_('Please let us know if you can make it.');

  if (entry.attending === 'No') {
    entry.dietary = '';
    entry.dietaryDetails = '';
  } else {
    if (DIET_OPTIONS.indexOf(entry.dietary) === -1) fail_('Please choose a dietary requirement.');
    if (entry.dietary !== DIET_ALLERGY) entry.dietaryDetails = '';
    else if (!entry.dietaryDetails) fail_('Please tell us about your allergy or intolerance.');
  }
  return entry;
}

function isPhone_(phone) {
  const digits = phone.replace(/\D/g, '').length;
  return /^[0-9+()\-\s]+$/.test(phone) && digits >= 9 && digits <= 15;
}

function listRsvps_(passcode) {
  checkPasscode_(passcode);
  const sheet = rsvpSheet_();
  const col = columnIndex_(sheet);
  const lastRow = sheet.getLastRow();
  const rows = lastRow < 2 ? [] : sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  const entries = rows
    .filter(row => String(row[col.email]).trim() !== '')
    .map(row => {
      const entry = {};
      COLUMNS.forEach(([key]) => {
        const value = row[col[key]];
        entry[key] = isDate_(value) ? value.toISOString() : String(value);
      });
      return entry;
    });
  return { ok: true, entries: entries, sheetUrl: sheet.getParent().getUrl() };
}

function checkPasscode_(given) {
  const expected = PropertiesService.getScriptProperties().getProperty('PASSCODE');
  if (!expected) fail_('The organiser passcode hasn’t been set up yet.');
  const cache = CacheService.getScriptCache();
  const wrong = Number(cache.get('wrongPasscodes') || 0);
  if (wrong >= MAX_WRONG_PASSCODES) fail_('Too many wrong passcodes. Please try again in 15 minutes.');
  if (String(given || '') !== expected) {
    cache.put('wrongPasscodes', String(wrong + 1), 15 * 60);
    fail_('That passcode isn’t right.');
  }
}

/** Builds the Summary tab: live counts for the caterer, as formulas that read the RSVPs tab. */
function buildSummary_(ss, col) {
  const sheet = ss.getSheetByName(SUMMARY_NAME) || ss.insertSheet(SUMMARY_NAME);
  sheet.clear();
  const range = key => `${SHEET_NAME}!${colLetter_(col[key])}2:${colLetter_(col[key])}`;
  const attending = `${range('attending')},"Yes"`;

  const rows = [
    ['RSVP summary', ''],
    ['Responses', `=COUNTA(${range('email')})`],
    ['Attending', `=COUNTIF(${attending})`],
    ['Not attending', `=COUNTIF(${range('attending')},"No")`],
    ['', ''],
    ['Dietary requirements (attending)', ''],
    ...DIET_OPTIONS.map(option => [option, `=COUNTIFS(${attending},${range('dietary')},"${option}")`]),
    ['', ''],
    ['Allergies & intolerances (attending)', ''],
  ];
  rows.forEach(([label, formula], i) => {
    const cell = sheet.getRange(i + 1, 1).setValue(label);
    if (formula) sheet.getRange(i + 1, 2).setFormula(formula);
    else if (label) cell.setFontWeight('bold');
  });
  // Lists each attending guest's name and allergy details under the last heading.
  sheet.getRange(rows.length + 1, 1).setFormula(
    `=IFERROR(FILTER({${range('name')},${range('dietaryDetails')}},` +
    `${range('attending')}="Yes",${range('dietary')}="${DIET_ALLERGY}"),"None so far")`);
  sheet.setColumnWidth(1, 320);
  sheet.setColumnWidth(2, 260);
}

function spreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}

function rsvpSheet_() {
  const sheet = spreadsheet_().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`There's no "${SHEET_NAME}" tab. Run setup() first.`);
  return sheet;
}

/** Maps each field to its column position (0-based), found by heading. */
function columnIndex_(sheet) {
  const headings = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const col = {};
  COLUMNS.forEach(([key, heading]) => {
    col[key] = headings.indexOf(heading);
    if (col[key] === -1) throw new Error(`The "${heading}" column is missing from the ${SHEET_NAME} tab.`);
  });
  return col;
}

/** Row number (1-based) holding this email, or 0 if it's a new one. */
function findRow_(sheet, emailCol, email) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  const emails = sheet.getRange(2, emailCol + 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < emails.length; i++) {
    if (String(emails[i][0]).trim().toLowerCase() === email) return i + 2;
  }
  return 0;
}

// A leading apostrophe makes Sheets store the value as plain text: "=..." can't
// run as a formula, and the 0 at the start of a cellphone number is kept.
function asText_(value) {
  return value === '' ? '' : "'" + value;
}

function isDate_(value) {
  return Object.prototype.toString.call(value) === '[object Date]';
}

function colLetter_(index) { // 0 → A, 25 → Z, 26 → AA
  let letters = '';
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) {
    letters = String.fromCharCode(65 + ((n - 1) % 26)) + letters;
  }
  return letters;
}

function fail_(message) {
  const err = new Error(message);
  err.userFacing = true;
  throw err;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
