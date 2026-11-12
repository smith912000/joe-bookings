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
var EARLIEST = 7;   // earliest PT start hour (7am)

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

// Joe's weekly hours live in the "Schedule" tab: Type (Open/Shift/Away), Day, From hour, To hour, Away from, Away to, Label, Week of (optional).
// Rows 'PT before class' and 'PT after class' (hours in From hour) open that many hours of PT either side of every Boxing Academy class.
// A row with a Week of date (the Monday of that week) applies to that week only and replaces the usual rows for that day.
// Returns null if the tab is missing or has no usable rows, so the site falls back to its saved schedule.
function readSchedule_() {
  var sh = SpreadsheetApp.getActive().getSheetByName('Schedule');
  if (!sh) return null;
  var rows = sh.getDataRange().getDisplayValues(), open = {}, blocked = [], away = [], weeks = [], around = { before: 0, after: 0 }, n = 0;
  var days = ['monday','tuesday','wednesday','thursday','friday','saturday','sunday'];
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i], t = String(r[0]).trim().toLowerCase();
    var di = days.indexOf(String(r[1]).trim().toLowerCase()), day = di < 0 ? null : days[di].charAt(0).toUpperCase() + days[di].slice(1);
    var a = parseHour(r[2]), b = parseHour(r[3]);
    var wk = r[7] ? parseDate(r[7]) : null, kind = t === 'open' ? 'open' : (t === 'shift' || t === 'blocked') ? 'shift' : null;
    if (kind && day && a !== null && b !== null && b > a) {
      if (wk) { weeks.push({ week: wk, day: day, type: kind, from: a, to: b, label: r[6] || '' }); }
      else if (kind === 'open') { (open[day] = open[day] || []).push([a, b, r[6] || '']); }
      else { blocked.push({ day: day, from: a, to: b, label: r[6] || '' }); }
      n++;
    }
    else if (t === 'pt before class' || t === 'pt after class') { var hrs = parseInt(r[2], 10); if (hrs >= 0 && hrs <= 4) { around[t === 'pt before class' ? 'before' : 'after'] = hrs; n++; } }
    else if (t === 'away') { var f = parseDate(r[4]), u = parseDate(r[5]) || f; if (f) { away.push({ from: f, to: u, note: r[6] || '' }); n++; } }
  }
  return n ? { open: open, blocked: blocked, away: away, weeks: weeks, around: around } : null;
}

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function doGet(e) {
  var tz = Session.getScriptTimeZone(), res = readBusy_();
  var out = { updated: Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm'), busy: res.busy };
  var sch = readSchedule_(); if (sch) out.schedule = sch;
  if (e && e.parameter && e.parameter.debug) { out.rowsRead = res.rows.length - 1; out.rowsUsed = res.used; out.skippedRows = res.skipped; out.headers = res.rows[0]; out.firstRow = res.rows[1] || null; out.tz = tz; }
  return json_(out);
}

// Coach Joe's page saves his usual week here. The key lives in the "Admin link" tab of his sheet (cell B1), which only
// he and Brandon can open; customers never see it. Rows with a Week of date and the notes row are left as they are.
function saveSchedule_(d) {
  var ss = SpreadsheetApp.getActive(), ks = ss.getSheetByName('Admin link');
  var key = ks ? String(ks.getRange('B1').getValue()).trim() : '';
  // Checked before any lock is taken, so a run of wrong keys can never hold up real bookings.
  if (key.length < 20 || String(d.key || '') !== key) { Utilities.sleep(1500); return json_({ ok: false, error: 'That link is not recognised. Open the link from your bookings sheet again.' }); }
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try { return writeSchedule_(ss, d); } finally { lock.releaseLock(); }
}
function writeSchedule_(ss, d) {
  var days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
  var usual = Array.isArray(d.usual) ? d.usual : [], away = Array.isArray(d.away) ? d.away : [], ar = d.around || {};
  if (usual.length > 60 || away.length > 40) return json_({ ok: false, error: 'Too many rows.' });
  var out = [];
  for (var i = 0; i < usual.length; i++) {
    var u = usual[i], type = u.type === 'Open' ? 'Open' : u.type === 'Shift' ? 'Shift' : null, di = days.indexOf(u.day);
    var f = parseInt(u.from, 10), t = parseInt(u.to, 10);
    if (!type || di < 0 || !(f >= 0 && f < 24) || !(t > f && t <= 24)) return json_({ ok: false, error: 'Row ' + (i + 1) + ' needs a day, a type and a start hour before the end hour.' });
    out.push({ di: di, f: f, row: [type, u.day, f, t, '', '', String(u.label || '').trim().slice(0, 80), ''] });
  }
  out.sort(function (a, b) { return a.di - b.di || a.f - b.f; });
  var rowsOut = out.map(function (o) { return o.row; });
  var nb = Math.min(Math.max(parseInt(ar.before, 10) || 0, 0), 4), na = Math.min(Math.max(parseInt(ar.after, 10) || 0, 0), 4);
  rowsOut.push(['PT before class', '', nb, '', '', '', 'Hours of PT just before every Boxing Academy class (0 switches it off)', '']);
  rowsOut.push(['PT after class', '', na, '', '', '', 'Hours of PT just after every Boxing Academy class (0 switches it off)', '']);
  for (var j = 0; j < away.length; j++) {
    var a = parseDate(away[j].from || ''), b = parseDate(away[j].to || '') || a;
    if (!a || b < a) return json_({ ok: false, error: 'Away row ' + (j + 1) + ' needs a first and last date.' });
    rowsOut.push(['Away', '', '', '', a, b, String(away[j].note || '').trim().slice(0, 80), '']);
  }
  var sh = ss.getSheetByName('Schedule');
  if (!sh) return json_({ ok: false, error: 'Schedule tab not found.' });
  var old = sh.getDataRange().getDisplayValues(), keep = [], notes = [], managed = { 'open': 1, 'shift': 1, 'blocked': 1, 'away': 1, 'pt before class': 1, 'pt after class': 1 };
  for (var k = 1; k < old.length; k++) {
    var r = old[k].slice(0, 8); while (r.length < 8) r.push('');
    var ty = String(r[0]).trim().toLowerCase();
    if (!ty && !r[1] && !r[7]) continue;
    if (managed[ty] && String(r[7]).trim()) keep.push(r);        // week-only rows stay
    else if (!managed[ty]) notes.push(r);                          // the notes row stays, last
  }
  var all = rowsOut.concat(keep, notes);
  if (old.length > 1) sh.getRange(2, 1, old.length - 1, 8).clearContent();
  sh.getRange(2, 5, all.length, 3).setNumberFormat('@');                      // dates and labels as plain text
  sh.getRange(2, 8, all.length, 1).setNumberFormat('@');
  sh.getRange(2, 1, all.length, 8).setValues(all);
  return json_({ ok: true, saved: rowsOut.length, updated: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') });
}

function doPost(e) {
  var d0;
  try { d0 = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'Could not read that.' }); }
  if (d0 && d0.action === 'saveSchedule') return saveSchedule_(d0);   // Coach Joe's page, needs the key from his sheet; takes its own lock
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var d = d0;
    if (d.website) return json_({ ok: true });                     // honeypot
    var name = String(d.name || '').trim().slice(0, 80), contact = String(d.contact || '').trim().slice(0, 100);
    var note = String(d.note || '').trim().slice(0, 300), hours = Math.min(Math.max(parseInt(d.hours, 10) || 1, 1), 3);
    var m = String(d.slot || '').match(/^(\d{4})-(\d{2})-(\d{2}) (\d{2})$/);
    if (!name || !contact || !m) return json_({ ok: false, error: 'Please give your name, contact and a time.' });
    var date = m[1] + '-' + m[2] + '-' + m[3], start = parseInt(m[4], 10);
    if (start + hours > 24 || start < EARLIEST) return json_({ ok: false, error: 'Sorry, that time is not available. The earliest PT is 7am.' });

    // The slot must be one Joe offers: open window, not a shift or class, not an away date, in the future, within 8 weeks.
    var sch = readSchedule_();
    if (!sch) return json_({ ok: false, error: 'Please contact Coach Joe directly.' });
    var days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    var day = days[new Date(+m[1], +m[2] - 1, +m[3]).getDay()], now = new Date();
    if (new Date(+m[1], +m[2] - 1, +m[3], start) < now || new Date(+m[1], +m[2] - 1, +m[3], start) > new Date(now.getTime() + 56 * 86400000))
      return json_({ ok: false, error: 'Sorry, that time is not available.' });
    // Rules for that date: the weekly pattern, unless the sheet has rows for that exact week and day.
    var dt = new Date(+m[1], +m[2] - 1, +m[3]), mon = new Date(dt.getTime()); mon.setDate(mon.getDate() - ((dt.getDay() + 6) % 7));
    var wkKey = mon.getFullYear() + '-' + pad2(mon.getMonth() + 1) + '-' + pad2(mon.getDate());
    var ov = (sch.weeks || []).filter(function (w) { return w.week === wkKey && w.day === day; });
    var rOpen = ov.length ? ov.filter(function (w) { return w.type === 'open'; }).map(function (w) { return [w.from, w.to]; }) : (sch.open[day] || []);
    var rBlocked = ov.length ? ov.filter(function (w) { return w.type === 'shift'; }) : sch.blocked.filter(function (b) { return b.day === day; });
    (sch.around && (sch.around.before || sch.around.after) ? rBlocked : []).forEach(function (b) {
      if (!/boxing/i.test(b.label || '')) return;
      if (sch.around.before > 0) rOpen = rOpen.concat([[Math.max(0, b.from - sch.around.before), b.from]]);
      if (sch.around.after > 0) rOpen = rOpen.concat([[b.to, Math.min(24, b.to + sch.around.after)]]);
    });
    for (var k = start; k < start + hours; k++) {
      var inOpen = rOpen.some(function (w) { return k >= w[0] && k < w[1]; });
      var inShift = rBlocked.some(function (b) { return k >= b.from && k < b.to; });
      var inAway = sch.away.some(function (a) { return date >= a.from && date <= (a.to || a.from); });
      if (!inOpen || inShift || inAway) return json_({ ok: false, error: 'Sorry, that time is not available.' });
    }

    var res = readBusy_(), busy = res.busy;
    for (var h = start; h < start + hours; h++) if (busy.indexOf(date + ' ' + pad2(h)) >= 0) return json_({ ok: false, error: 'Sorry, that time has just been taken.' });

    // Spam limits: only Pending rows count, so junk can't fill the sheet for good; a hard cap stops runaway growth.
    var pending = 0, mine = 0, lc = contact.toLowerCase();
    for (var i = 1; i < res.rows.length; i++) {
      if (String(res.rows[i][7]).trim().toLowerCase() !== 'pending') continue;
      pending++;
      if (String(res.rows[i][5]).trim().toLowerCase() === lc) mine++;
    }
    if (mine >= 3) return json_({ ok: false, error: 'You already have requests waiting for Coach Joe to confirm. Please wait for him to reply.' });
    if (pending >= 60 || res.rows.length > 3000) return json_({ ok: false, error: 'Please contact Coach Joe directly.' });

    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET), row = sh.getLastRow() + 1;
    // Name, contact and notes are stored as plain text, so a value starting with = + - or @ can never run as a formula.
    sh.getRange(row, 5, 1, 2).setNumberFormat('@');
    sh.getRange(row, 9, 1, 1).setNumberFormat('@');
    sh.getRange(row, 1, 1, 10).setValues([['R' + new Date().getTime().toString(36), date, pad2(start) + ':00', pad2(start + hours) + ':00', name, contact, '', 'Pending',
                  (d.change ? 'CHANGE OF TIME REQUEST. ' : '') + note, '']]);
    return json_({ ok: true });
  } catch (err) { return json_({ ok: false, error: 'Could not save: ' + err }); }
  finally { lock.releaseLock(); }
}
