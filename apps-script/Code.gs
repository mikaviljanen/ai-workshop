/**
 * Backend for the workshop guide page.
 * Bind this script to a Google Sheet (Extensions > Apps Script), paste this file,
 * add a script property FACILITATOR_KEY (Project Settings > Script properties),
 * run setup() once, then Deploy > New deployment > Web app
 * (Execute as: Me, Who has access: Anyone).
 * The key lives in script properties so this file can sit in a public repo.
 */

function facKey_() {
  const k = PropertiesService.getScriptProperties().getProperty('FACILITATOR_KEY');
  if (!k) throw new Error('Set the FACILITATOR_KEY script property first.');
  return k;
}
const SHEET_NAME = 'entries';
const HEADERS = [
  'id', 'created', 'updated', 'group', 'author', 'title', 'description', 'pain',
  'frequency', 'inputs', 'outputs', 'done', 'votes', 'status', 'appLink', 'lessons'
];
const OPEN_FIELDS = ['group', 'author', 'title', 'description', 'pain', 'frequency', 'inputs', 'outputs', 'done'];
const SHOWCASE_FIELDS = ['appLink', 'lessons'];
const STATUSES = ['open', 'picked', 'built', 'parked'];
const MAX_LEN = 3000;

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  sh.setFrozenRows(1);
}

function doGet() {
  return json_({ ok: true, entries: readAll_() });
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'Request was not valid JSON.' });
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    switch (body.action) {
      case 'add':
        return json_({ ok: true, entry: add_(body.entry || {}) });
      case 'vote':
        return json_({ ok: true, entry: patch_(body.id, row => { row.votes = Number(row.votes || 0) + 1; }) });
      case 'showcase':
        return json_({ ok: true, entry: patch_(body.id, row => copy_(row, body.fields || {}, SHOWCASE_FIELDS)) });
      case 'status':
        if (body.key !== facKey_()) return json_({ ok: false, error: 'Wrong facilitator key.' });
        if (STATUSES.indexOf(body.status) < 0) return json_({ ok: false, error: 'Unknown status.' });
        return json_({ ok: true, entry: patch_(body.id, row => { row.status = body.status; }) });
      case 'checkKey':
        return json_({ ok: body.key === facKey_() });
      default:
        return json_({ ok: false, error: 'Unknown action.' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function sheet_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('Run setup() first.');
  return sh;
}

function readAll_() {
  const values = sheet_().getDataRange().getValues();
  return values.slice(1).filter(r => r[0]).map(toObj_);
}

function toObj_(r) {
  const o = {};
  HEADERS.forEach((h, i) => { o[h] = r[i] instanceof Date ? r[i].toISOString() : r[i]; });
  o.votes = Number(o.votes || 0);
  return o;
}

function clean_(v) {
  let s = String(v == null ? '' : v).slice(0, MAX_LEN);
  // Stop spreadsheet formula injection.
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function copy_(row, src, fields) {
  fields.forEach(f => { if (f in src) row[f] = clean_(src[f]); });
}

function add_(entry) {
  if (!entry.title || !String(entry.title).trim()) throw new Error('A title is required.');
  if (entry.noStudentData !== 'yes') throw new Error('Confirm that the workflow does not handle students\' personal data.');
  const now = new Date().toISOString();
  const row = { id: Utilities.getUuid().slice(0, 8), created: now, updated: now, votes: 0, status: 'open', appLink: '', lessons: '' };
  copy_(row, entry, OPEN_FIELDS);
  sheet_().appendRow(HEADERS.map(h => row[h] == null ? '' : row[h]));
  return row;
}

function patch_(id, fn) {
  const sh = sheet_();
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      const row = toObj_(values[i]);
      fn(row);
      row.updated = new Date().toISOString();
      sh.getRange(i + 1, 1, 1, HEADERS.length).setValues([HEADERS.map(h => row[h] == null ? '' : row[h])]);
      return row;
    }
  }
  throw new Error('No entry with that id.');
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
