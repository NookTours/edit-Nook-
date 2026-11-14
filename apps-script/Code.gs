// Nook Tours — invoices storage (Google Apps Script web app)
// Access: Deploy > Execute as: Me, Who has access: Anyone.
// No login: anyone who has the web app URL can read and save invoices.

var SHEET_NAME = 'invoices';
var HEAD = ['id', 'updated', 'customer', 'phone', 'total', 'status', 'json'];

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
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

// Health check only (opening the URL in a browser).
function doGet() {
  return out_({ ok: true, info: 'Nook Tours invoices API' });
}

// POST, body = JSON as text: {action, ...}
//   action 'list'                  -> {ok, items:[{id,updated,customer,phone,total,status}]}  newest first
//   action 'get', id               -> {ok, data:{...}}
//   action 'save', data:{id,...}   -> {ok, id}   (creates or overwrites the row with that id)
function doPost(e) {
  var body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) {}

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
        phone: saved.phone || String(values[i][3]), total: values[i][4], status: values[i][5]
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
    return out_({ ok: true, id: id });
  }

  return out_({ ok: false, error: 'unknown action' });
}
