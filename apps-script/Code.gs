// Nook Tours — invoices storage (Google Apps Script web app)
// Access: Deploy > Execute as: Me, Who has access: Anyone.
// Users live in the "users" sheet (username, salt, hash). Only salted SHA-256 hashes are stored, never passwords.
// Every user has the same permissions: all invoices.

var SHEET_NAME = 'invoices';
var USERS_SHEET = 'users';
var HEAD = ['id', 'updated', 'customer', 'phone', 'total', 'status', 'json'];
var USERS_HEAD = ['username', 'salt', 'hash'];

// Run setupUsers() once from the editor to create the users sheet with these rows.
var USERS = [
  ['amanub', 'da685a1d156d09420325211dd0e13b90', '40bf6a7c7b405a7694855fd321a0a824787a791b76301c5e4b00c1823426d767'],
  ['beshoy', 'd381eb0093eba3732644a4c865c8be26', 'c22d8a4b2519660b3a2a8d1ba897934c59e5cf331222c2786a2103093d07022a'],
  ['john', '76bf9b91653dd61c33de5965dbedeb09', '9fa6250c339ba423a57f00b1e7012bea0e8cb7216dd5212c56ac3593c669a811'],
  ['joseph', '6dd8c88ef9ff31f393f1cfa297c04193', 'eb5e788e22e0e368155abc71bc0f036b31d416ef103da40e8d26e8bbbe247189']
];

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function sha_(s) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8);
  return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

// Returns the username if user+password match a row in the users sheet, otherwise ''.
function login_(user, pw) {
  var sh = usersSheet_();
  var values = sh.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]).toLowerCase() === String(user || '').toLowerCase()) {
      var h = sha_(String(values[i][1]) + ':' + String(pw || ''));
      return h === String(values[i][2]) ? String(values[i][0]) : '';
    }
  }
  return '';
}

function usersSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(USERS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(USERS_SHEET);
    sh.appendRow(USERS_HEAD);
    sh.setFrozenRows(1);
  }
  return sh;
}

// Run once from the editor: writes the users above into the "users" sheet (replaces existing rows).
function setupUsers() {
  var sh = usersSheet_();
  sh.clearContents();
  sh.appendRow(USERS_HEAD);
  USERS.forEach(function (u) { sh.appendRow(u); });
  sh.setFrozenRows(1);
}

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEAD);
    sh.setFrozenRows(1);
  }
  return sh;
}

function findRow_(values, id) {
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) return i + 1;
  }
  return 0;
}

// Health check only (opening the URL in a browser). No data here.
function doGet() {
  return out_({ ok: true, info: 'Nook Tours invoices API' });
}

// POST, body = JSON as text: {user, pw, action, ...}
//   action 'list'                  -> {ok, items:[{id,updated,customer,phone,total,status}]}  newest first
//   action 'get', id               -> {ok, data:{...}}
//   action 'save', data:{id,...}   -> {ok, id, by}   (creates or overwrites the row with that id)
function doPost(e) {
  var body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) {}

  var who = login_(body.user, body.pw);
  if (!who) return out_({ ok: false, error: 'bad password' });

  var sh = sheet_();
  // keep the phone column as text so numbers like 01005377933 keep their leading zero
  sh.getRange(1, 4, sh.getMaxRows(), 1).setNumberFormat('@');
  var values = sh.getDataRange().getValues();

  if (body.action === 'list') {
    var items = [];
    for (var i = 1; i < values.length; i++) {
      var saved = {};
      try { saved = JSON.parse(values[i][6] || '{}'); } catch (err) {}
      items.push({
        id: values[i][0], updated: values[i][1], customer: values[i][2],
        phone: saved.phone || String(values[i][3]), total: values[i][4], status: values[i][5],
        by: saved.by || ''
      });
    }
    items.reverse();
    return out_({ ok: true, items: items, user: who });
  }

  if (body.action === 'get') {
    var row = findRow_(values, body.id);
    if (!row) return out_({ ok: false, error: 'not found' });
    return out_({ ok: true, data: JSON.parse(values[row - 1][6] || '{}') });
  }

  if (body.action === 'save') {
    var d = body.data || {};
    var id = String(d.id || '');
    if (!id) return out_({ ok: false, error: 'no id' });
    d.by = who;
    var rec = [id, new Date(), d.customer || '', '', d.total || '', d.status || '', JSON.stringify(d)];
    var row2 = findRow_(values, id);
    if (row2) {
      sh.getRange(row2, 1, 1, rec.length).setValues([rec]);
    } else {
      sh.appendRow(rec);
      row2 = sh.getLastRow();
    }
    // phone: set text format first, then the value, so a leading 0 is kept
    sh.getRange(row2, 4).setNumberFormat('@').setValue(String(d.phone || ''));
    return out_({ ok: true, id: id, by: who });
  }

  return out_({ ok: false, error: 'unknown action' });
}
