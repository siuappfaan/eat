// Meal Planner sync + archive. In your Google Sheet: Extensions > Apps Script, paste this,
// then Deploy > Manage deployments > edit (pencil) > Version: New version > Deploy.
// (First time: Deploy > New deployment > Web app > Execute as: Me, Who has access: Anyone.)
const DAYS = ['Mon','Tue','Wed','Thu','Fri','Sat','Sun'];

function doGet(e) {
  if (e && e.parameter && e.parameter.a === 'archive') return out({ weeks: readArchive() });
  return out(readState());
}
function doPost(e) {
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const inc = JSON.parse(e.postData.contents);
    if (inc.type === 'archive') { upsertArchive(inc.week); return out({ ok: true }); }
    const cur = readState();
    if (!cur || inc.updatedAt >= cur.updatedAt) { writeState(inc); return out(inc); }
    return out(cur);
  } finally { lock.releaseLock(); }
}
function tab(n) { const ss = SpreadsheetApp.getActive(); return ss.getSheetByName(n) || ss.insertSheet(n); }
function tags(i) { return [i.w ? 'WFH' : '', i.o ? 'Out' : ''].filter(String).join(', '); }
function readState() { const v = tab('MealState').getRange('A1').getValue(); return v ? JSON.parse(v) : null; }

// Current week, shown as a readable table
function writeState(s) {
  tab('MealState').getRange('A1').setValue(JSON.stringify(s));
  const n = s.weekend ? 14 : 10, rows = [['Week of ' + s.ws, s.names[0], s.names[1]]];
  for (let i = 0; i < n; i++) {
    const f = c => { const x = s.cells[c][i]; return x.t + (tags(x) ? ' (' + tags(x) + ')' : ''); };
    rows.push([DAYS[i >> 1] + (i % 2 ? ' Dinner' : ' Lunch'), f(0), f(1)]);
  }
  const p = tab('Meal Plan'); p.clear(); p.getRange(1, 1, rows.length, 3).setValues(rows);
}

// Archive of all past weeks: one row per meal slot, never deleted
function archiveTab(names) {
  const sh = tab('Archive');
  if (sh.getLastRow() === 0) {
    sh.appendRow(['Week starting', 'Day', 'Meal', names[0], names[0] + ' tags', names[1], names[1] + ' tags']);
    sh.setFrozenRows(1); sh.getRange('A:A').setNumberFormat('@');
  }
  return sh;
}
function upsertArchive(w) {
  const sh = archiveTab(w.names), last = sh.getLastRow();
  if (last > 1) {
    const v = sh.getRange(2, 1, last - 1, 1).getValues();
    for (let r = v.length - 1; r >= 0; r--) if (String(v[r][0]) === w.ws) sh.deleteRow(r + 2);
  }
  const n = w.weekend ? 14 : 10, rows = [];
  for (let i = 0; i < n; i++) {
    const a = w.cells[0][i], b = w.cells[1][i];
    rows.push([w.ws, DAYS[i >> 1], i % 2 ? 'Dinner' : 'Lunch', a.t, tags(a), b.t, tags(b)]);
  }
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, 7).setValues(rows);
}
function readArchive() {
  const sh = archiveTab(['Tiffany', 'Eric']), last = sh.getLastRow();
  if (last < 2) return [];
  const blank = () => Array.from({ length: 14 }, () => ({ t: '', w: false, o: false }));
  const item = (t, g) => ({ t: String(t), w: String(g).indexOf('WFH') >= 0, o: String(g).indexOf('Out') >= 0 });
  const m = {};
  sh.getRange(2, 1, last - 1, 7).getValues().forEach(r => {
    const ws = String(r[0]), d = DAYS.indexOf(r[1]); if (d < 0) return;
    const w = m[ws] || (m[ws] = { ws: ws, weekend: false, cells: [blank(), blank()] });
    const i = d * 2 + (r[2] === 'Dinner' ? 1 : 0); if (d > 4) w.weekend = true;
    w.cells[0][i] = item(r[3], r[4]); w.cells[1][i] = item(r[5], r[6]);
  });
  return Object.keys(m).map(k => m[k]);
}
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
