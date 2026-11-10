/**
 * Joe's Bookings: Sheet connection (Google Apps Script bound to the "Joe's Bookings" Sheet).
 * Deploy: Deploy > Manage deployments > (pencil) > Version: New version > Deploy. Web app, Execute as: Me, Who has access: Anyone.
 *
 * GET  /exec            -> {"updated","busy":["YYYY-MM-DD HH",...]}  hours only, no names or contact details. Pending and Confirmed rows hold a slot.
 * GET  /exec?debug=1    -> how the script read the Bookings tab (for checking).
 * Columns on "Bookings": Booking ID, Date, Start time, End time, Client, Contact, Gym, Status, Notes, Changed from.
 */
var SHEET = 'Bookings';

function pad2(n) { return ('0' + n).slice(-2); }

function parseDate(s) {
  s = String(s).trim();
  var m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return m[1] + '-' + pad2(m[2]) + '-' + pad2(m[3]);
  m = s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})/);
  if (m) { var y = m[3].length === 2 ? '20' + m[3] : m[3]; return y + '-' + pad2(m[2]) + '-' + pad2(m[1]); }
  var d = new Date(s);
  return isNaN(d) ? null : Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function parseHour(s) {
  var m = String(s).trim().match(/^(\d{1,2})(?::(\d{2}))?\s*([ap]m)?/i);
  if (!m) return null;
  var h = parseInt(m[1], 10);
  if (m[3]) { var pm = m[3].toLowerCase() === 'pm'; if (pm && h < 12) h += 12; if (!pm && h === 12) h = 0; }
  return h > 24 ? null : h;
}

function readBusy_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET) || SpreadsheetApp.getActive().getSheets()[0];
  var rows = sh.getDataRange().getDisplayValues();   // text exactly as shown: avoids date/time cell types
  var busy = {}, used = 0, skipped = [];
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (!r[1] && !r[2]) continue;
    var status = String(r[7]).trim().toLowerCase();
    if (status === 'cancelled' || status === 'canceled') continue;
    var date = parseDate(r[1]), a = parseHour(r[2]), b = parseHour(r[3]);
    if (!date || a === null) { skipped.push(i + 1); continue; }
    if (b === null || b <= a) b = a + 1;
    for (var h = a; h < b; h++) busy[date + ' ' + pad2(h)] = true;
    used++;
  }
  return { busy: Object.keys(busy).sort(), used: used, skipped: skipped, rows: rows };
}

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function doGet(e) {
  var tz = Session.getScriptTimeZone(), res = readBusy_();
  var out = { updated: Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm'), busy: res.busy };
  if (e && e.parameter && e.parameter.debug) { out.rowsRead = res.rows.length - 1; out.rowsUsed = res.used; out.skippedRows = res.skipped; out.headers = res.rows[0]; out.firstRow = res.rows[1] || null; out.tz = tz; }
  return json_(out);
}
