/**
 * Joe's Bookings: Sheet connection (Google Apps Script, bound to the "Joe's Bookings" Sheet).
 * Deploy once: Deploy > New deployment > Web app > Execute as: Me, Who has access: Anyone.
 * Put the resulting /exec URL in data/config.json as "busy_url".
 *
 * GET  -> {"updated": "...", "busy": ["YYYY-MM-DD HH", ...]}   (hours only, no names or contact details)
 * The customer page reads this live, so a change made in the Sheet shows up on the calendar at once.
 * Sheet tab "Bookings" columns: Booking ID, Date, Start time, End time, Client, Contact, Gym, Status, Notes, Changed from.
 */
function doGet() {
  var sh = SpreadsheetApp.getActive().getSheetByName('Bookings');
  var tz = Session.getScriptTimeZone();
  var rows = sh.getDataRange().getValues();
  var busy = {};
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (!r[1] || r[7] === 'Cancelled') continue;
    var date = r[1] instanceof Date ? Utilities.formatDate(r[1], tz, 'yyyy-MM-dd') : String(r[1]);
    var a = hourOf(r[2], tz), b = hourOf(r[3], tz);
    if (a === null) continue;
    if (b === null || b <= a) b = a + 1;
    for (var h = a; h < b; h++) busy[date + ' ' + ('0' + h).slice(-2)] = true;
  }
  var out = { updated: Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm'), busy: Object.keys(busy).sort() };
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
function hourOf(v, tz) {
  if (v === '' || v === null) return null;
  if (v instanceof Date) return parseInt(Utilities.formatDate(v, tz, 'H'), 10);
  var n = parseInt(String(v).split(':')[0], 10);
  return isNaN(n) ? null : n;
}
