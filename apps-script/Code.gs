/**
 * Joe's Bookings: Sheet connection (Google Apps Script bound to the "Joe's Bookings" Sheet).
 * Deploy: Deploy > Manage deployments > (pencil) > Version: New version > Deploy. Web app, Execute as: Me, Who has access: Anyone.
 *
 * GET  /exec            -> {"updated","busy":["YYYY-MM-DD HH",...]}  hours only, no names or contact details. Pending and Confirmed rows hold a slot.
 * GET  /exec?debug=1    -> how the script read the Bookings tab (for checking).
 * POST /exec  (text/plain JSON {name, contact, slot:"YYYY-MM-DD HH", hours, note, change, website})
 *                       -> adds a row with Status "Pending" for Joe to confirm in the Sheet.
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

// Joe's weekly hours live in the "Schedule" tab: Type (Open/Shift/Away), Day, From hour, To hour, Away from, Away to, Label.
// Returns null if the tab is missing or has no usable rows, so the site falls back to its saved schedule.
function readSchedule_() {
  var sh = SpreadsheetApp.getActive().getSheetByName('Schedule');
  if (!sh) return null;
  var rows = sh.getDataRange().getDisplayValues(), open = {}, blocked = [], away = [], n = 0;
  var days = ['monday','tuesday','wednesday','thursday','friday','saturday','sunday'];
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i], t = String(r[0]).trim().toLowerCase();
    var di = days.indexOf(String(r[1]).trim().toLowerCase()), day = di < 0 ? null : days[di].charAt(0).toUpperCase() + days[di].slice(1);
    var a = parseHour(r[2]), b = parseHour(r[3]);
    if (t === 'open' && day && a !== null && b !== null && b > a) { (open[day] = open[day] || []).push([a, b]); n++; }
    else if ((t === 'shift' || t === 'blocked') && day && a !== null && b !== null && b > a) { blocked.push({ day: day, from: a, to: b, label: r[6] || '' }); n++; }
    else if (t === 'away') { var f = parseDate(r[4]), u = parseDate(r[5]) || f; if (f) { away.push({ from: f, to: u, note: r[6] || '' }); n++; } }
  }
  return n ? { open: open, blocked: blocked, away: away } : null;
}

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function doGet(e) {
  var tz = Session.getScriptTimeZone(), res = readBusy_();
  var out = { updated: Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm'), busy: res.busy };
  var sch = readSchedule_(); if (sch) out.schedule = sch;
  if (e && e.parameter && e.parameter.debug) { out.rowsRead = res.rows.length - 1; out.rowsUsed = res.used; out.skippedRows = res.skipped; out.headers = res.rows[0]; out.firstRow = res.rows[1] || null; out.tz = tz; }
  return json_(out);
}

function doPost(e) {
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var d = JSON.parse(e.postData.contents);
    if (d.website) return json_({ ok: true });                     // honeypot
    var name = String(d.name || '').trim().slice(0, 80), contact = String(d.contact || '').trim().slice(0, 100);
    var note = String(d.note || '').trim().slice(0, 300), hours = Math.min(Math.max(parseInt(d.hours, 10) || 1, 1), 3);
    var m = String(d.slot || '').match(/^(\d{4}-\d{2}-\d{2}) (\d{2})$/);
    if (!name || !contact || !m) return json_({ ok: false, error: 'Please give your name, contact and a time.' });
    var start = parseInt(m[2], 10), busy = readBusy_().busy;
    for (var h = start; h < start + hours; h++) if (busy.indexOf(m[1] + ' ' + pad2(h)) >= 0) return json_({ ok: false, error: 'Sorry, that time has just been taken.' });
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET);
    if (sh.getLastRow() > 500) return json_({ ok: false, error: 'Please contact Joe directly.' });
    sh.appendRow(['R' + new Date().getTime().toString(36), m[1], pad2(start) + ':00', pad2(start + hours) + ':00', name, contact, '', 'Pending',
                  (d.change ? 'CHANGE OF TIME REQUEST. ' : '') + note, '']);
    return json_({ ok: true });
  } catch (err) { return json_({ ok: false, error: 'Could not save: ' + err }); }
  finally { lock.releaseLock(); }
}
