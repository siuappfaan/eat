// Meal Planner sync. In your Google Sheet: Extensions > Apps Script, paste this,
// then Deploy > New deployment > Web app > Execute as: Me, Who has access: Anyone.
const DAYS = ['Mon','Tue','Wed','Thu','Fri'];
function doGet() { return out(readState()); }
function doPost(e) {
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const inc = JSON.parse(e.postData.contents), cur = readState();
    if (!cur || inc.updatedAt >= cur.updatedAt) { writeState(inc); return out(inc); }
    return out(cur);
  } finally { lock.releaseLock(); }
}
function tab(n) { const ss = SpreadsheetApp.getActive(); return ss.getSheetByName(n) || ss.insertSheet(n); }
function readState() { const v = tab('MealState').getRange('A1').getValue(); return v ? JSON.parse(v) : null; }
function writeState(s) {
  tab('MealState').getRange('A1').setValue(JSON.stringify(s));
  const rows = [['', s.names[0], s.names[1]]];
  for (let i = 0; i < 10; i++) {
    const f = c => s.cells[c][i].t + (s.cells[c][i].w ? ' (WFH)' : '');
    rows.push([DAYS[Math.floor(i/2)] + (i % 2 ? ' Dinner' : ' Lunch'), f(0), f(1)]);
  }
  const p = tab('Meal Plan'); p.clear(); p.getRange(1, 1, 11, 3).setValues(rows);
}
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
