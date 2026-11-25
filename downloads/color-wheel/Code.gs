/**
 * The Color Wheel: saves student results to the Google Sheet this script is attached to.
 * One row per student per activity. Retrying an activity updates that row and adds 1 to Attempts.
 * Setup: Sheet > Extensions > Apps Script, paste this as Code.gs, add an HTML file named "Index",
 * then Deploy > New deployment > Web app.
 */
const SHEET_NAME = 'Results';
const HEADERS = ['Last updated', 'Student email', 'Name', 'Period', 'Activity', 'Correct', 'Total', 'Placed', 'Attempts', 'Answers'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('The Color Wheel')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Stops student-typed text from being read as a formula, and caps length.
function clean_(v, max) {
  let s = String(v == null ? '' : v).slice(0, max || 200);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}
function num_(v) { const n = Number(v); return isFinite(n) ? n : 0; }
function key_(name, period, activity) {
  return [name, period, activity].map(x => String(x).replace(/^'/, '').trim().toLowerCase()).join('|');
}

function saveResults(payload) {
  if (!payload || !payload.rows || !payload.rows.length || payload.rows.length > 20) throw new Error('Nothing to save.');
  const name = clean_(payload.name, 80);
  if (!name) throw new Error('Please type your name first.');
  const period = clean_(payload.period, 20);
  const email = Session.getActiveUser().getEmail() || '';   // filled in for district sign-ins
  const now = new Date();

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) sh = ss.insertSheet(SHEET_NAME);
    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    }
    const last = sh.getLastRow();
    const existing = last > 1 ? sh.getRange(2, 1, last - 1, 9).getValues() : [];
    const index = {};
    existing.forEach((r, i) => { index[key_(r[2], r[3], r[4])] = { row: i + 2, attempts: Number(r[8]) || 1 }; });

    const toAppend = [];
    payload.rows.forEach(r => {
      const activity = clean_(r.activity, 100);
      const vals = [now, email, name, period, activity, num_(r.right), num_(r.total), num_(r.placed), 1, clean_(r.answers, 3000)];
      const hit = index[key_(name, period, activity)];
      if (hit) {
        vals[8] = hit.attempts + 1;
        sh.getRange(hit.row, 1, 1, HEADERS.length).setValues([vals]);
      } else {
        toAppend.push(vals);
      }
    });
    if (toAppend.length) sh.getRange(sh.getLastRow() + 1, 1, toAppend.length, HEADERS.length).setValues(toAppend);
  } finally {
    lock.releaseLock();
  }
  return { time: Utilities.formatDate(now, Session.getScriptTimeZone(), 'h:mm a') };
}
