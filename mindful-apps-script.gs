// ═══════════════════════════════════════════════════════════════
// Still — Daily Practice · Google Apps Script
// ═══════════════════════════════════════════════════════════════
//
// SETUP
// ─────
// 1. script.google.com → New project → paste this → save.
// 2. Deploy → New deployment → Web app
//    Execute as: Me  |  Access: Anyone
// 3. Copy the /exec URL → paste into Still → Settings.
// 4. A "entries" sheet is auto-created on first sync.
// ═══════════════════════════════════════════════════════════════

var SHEET_ID = ''; // Leave blank for bound spreadsheet.

function getSheet(name) {
  var ss = SHEET_ID
    ? SpreadsheetApp.openById(SHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function doPost(e) {
  try {
    var raw = e.parameter.data || (e.postData && e.postData.contents) || '{}';
    var payload = JSON.parse(decodeURIComponent(raw));
    if (payload.entries) writeEntries(payload.entries);
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    var data = JSON.stringify({ entries: readEntries() });
    var cb = e && e.parameter && e.parameter.callback;
    if (cb) {
      return ContentService
        .createTextOutput(cb + '(' + data + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService
      .createTextOutput(data)
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

var HEADERS = ['id', 'date', 'type', 'data_json', 'preview'];

function writeEntries(entries) {
  var sheet = getSheet('entries');
  sheet.clearContents();
  sheet.appendRow(HEADERS);

  // Sort by date desc
  var sorted = entries.slice().sort(function(a, b) {
    return (b.date || '').localeCompare(a.date || '');
  });

  sorted.forEach(function(e) {
    sheet.appendRow([
      e.id   || '',
      e.date || '',
      e.type || '',
      JSON.stringify(e.data || {}),
      entryPreview(e)
    ]);
  });

  var hdr = sheet.getRange(1, 1, 1, HEADERS.length);
  hdr.setFontWeight('bold');
  hdr.setBackground('#0d0d1a');
  hdr.setFontColor('#b4a0ff');
  sheet.autoResizeColumns(1, HEADERS.length);
}

function readEntries() {
  var sheet = getSheet('entries');
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0];
  return data.slice(1).map(function(row) {
    var id   = row[0], date = row[1], type = row[2], dataJson = row[3];
    var d = {};
    try { d = JSON.parse(dataJson); } catch(e) {}
    return { id: id, date: date, type: type, data: d };
  });
}

function entryPreview(e) {
  var d = e.data || {};
  switch (e.type) {
    case 'meditation': return (d.duration || '') + 'min' + (d.type ? ' — ' + d.type : '') + (d.notes ? ' — ' + d.notes : '');
    case 'morning':    return [d.q1, d.q2, d.q3].filter(Boolean).join(' / ');
    case 'evening':    return [d.q1, d.q2, d.q3].filter(Boolean).join(' / ');
    case 'gratitude':  return [d.g1, d.g2, d.g3].filter(Boolean).join(' · ');
    case 'mood':       return ['','Low','Meh','Okay','Good','Great'][d.mood || 0] + (d.note ? ' — ' + d.note : '');
    case 'scripture':  return (d.ref || '') + (d.notes ? ' — ' + d.notes : '');
    case 'prayer':     return (d.type || '') + (d.notes ? ' — ' + d.notes : '');
    default:           return '';
  }
}
