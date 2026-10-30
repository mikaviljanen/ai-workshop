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
const SPRINT_SHEET = 'sprints';
const SPRINT_HEADERS = ['id', 'created', 'entryId', 'group', 'sprint', 'did', 'tested', 'wrong', 'next'];
const CHECK_SHEET = 'checks';
const CHECK_HEADERS = ['group', 'item', 'done', 'updated'];
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
  sprintSheet_();
  checkSheet_();
}

function sprintSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SPRINT_SHEET);
  if (!sh) {
    sh = ss.insertSheet(SPRINT_SHEET);
    sh.getRange(1, 1, 1, SPRINT_HEADERS.length).setValues([SPRINT_HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function checkSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CHECK_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CHECK_SHEET);
    sh.getRange(1, 1, 1, CHECK_HEADERS.length).setValues([CHECK_HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function readChecks_() {
  return checkSheet_().getDataRange().getValues().slice(1).filter(r => r[0] !== '')
    .map(r => ({ group: String(r[0]), item: String(r[1]), done: r[2] === true || r[2] === 'TRUE' }));
}

function setCheck_(c) {
  const group = clean_(c.group).slice(0, 60), item = clean_(c.item).slice(0, 60);
  if (!group || !item) throw new Error('Group and item are required.');
  const sh = checkSheet_(), values = sh.getDataRange().getValues(), now = new Date().toISOString();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === group && String(values[i][1]) === item) {
      sh.getRange(i + 1, 3, 1, 2).setValues([[!!c.done, now]]);
      return { group, item, done: !!c.done };
    }
  }
  sh.appendRow([group, item, !!c.done, now]);
  return { group, item, done: !!c.done };
}

function readSprints_() {
  const values = sprintSheet_().getDataRange().getValues();
  return values.slice(1).filter(r => r[0]).map(r => {
    const o = {};
    SPRINT_HEADERS.forEach((h, i) => { o[h] = r[i] instanceof Date ? r[i].toISOString() : r[i]; });
    o.sprint = Number(o.sprint || 0);
    return o;
  });
}

function logSprint_(s) {
  if (!s.entryId) throw new Error('Choose the task first.');
  if (!String(s.did || '').trim()) throw new Error('Say what you did.');
  const sh = sprintSheet_();
  const n = readSprints_().filter(x => String(x.entryId) === String(s.entryId)).length + 1;
  const row = { id: Utilities.getUuid().slice(0, 8), created: new Date().toISOString(), sprint: n };
  ['entryId', 'group', 'did', 'tested', 'wrong', 'next'].forEach(f => { row[f] = clean_(s[f]); });
  sh.appendRow(SPRINT_HEADERS.map(h => row[h] == null ? '' : row[h]));
  return row;
}

function doGet() {
  return json_({ ok: true, entries: readAll_(), sprints: readSprints_(), checks: readChecks_() });
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
      case 'log':
        return json_({ ok: true, sprint: logSprint_(body.sprint || {}) });
      case 'check':
        return json_({ ok: true, check: setCheck_(body.check || {}) });
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
