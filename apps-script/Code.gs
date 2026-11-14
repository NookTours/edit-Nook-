// Nook Tours — invoices storage (Google Apps Script web app)
// Put this in Apps Script (Extensions > Apps Script inside the "invoices" spreadsheet).
// Access: Deploy > Execute as: Me, Who has access: Anyone.
// Every request must include the Studio password (pw). The password is never stored here, only its salted SHA-256 hash.

var SALT = 'f93e791b4e301396c32f035371bdef68';
var HASH = '92c880970cafb356f08969e21500bb4bcab21fc28ca8a3dc06a0ac2d8f58f7f3';
var SHEET_NAME = 'invoices';
var HEAD = ['id', 'updated', 'customer', 'phone', 'total', 'status', 'json'];

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function sha_(s) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8);
  return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

function authed_(pw) {
  return sha_(SALT + ':' + String(pw || '')) === HASH;
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

// POST, body = JSON as text: {pw, action, ...}
//   action 'list'                  -> {ok, items:[{id,updated,customer,phone,total,status}]}  newest first
//   action 'get', id               -> {ok, data:{...}}
//   action 'save', data:{id,...}   -> {ok, id}   (creates or overwrites the row with that id)
function doPost(e) {
  var body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) {}

  if (!authed_(body.pw)) return out_({ ok: false, error: 'bad password' });

  var sh = sheet_();
  var values = sh.getDataRange().getValues();

  if (body.action === 'list') {
    var items = [];
    for (var i = 1; i < values.length; i++) {
      items.push({
        id: values[i][0], updated: values[i][1], customer: values[i][2],
        phone: values[i][3], total: values[i][4], status: values[i][5]
      });
    }
    items.reverse();
    return out_({ ok: true, items: items });
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
    var rec = [id, new Date(), d.customer || '', d.phone || '', d.total || '', d.status || '', JSON.stringify(d)];
    var row2 = findRow_(values, id);
    if (row2) sh.getRange(row2, 1, 1, rec.length).setValues([rec]);
    else sh.appendRow(rec);
    return out_({ ok: true, id: id });
  }

  return out_({ ok: false, error: 'unknown action' });
}
