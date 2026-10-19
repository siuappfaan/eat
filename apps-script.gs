// Meal Planner sync + archive (v1.2). Replace the old script with this, then
// Deploy > Manage deployments > edit (pencil) > Version: New version > Deploy.
// Slots: (day offset from Monday + 2) * 2 + (0 lunch, 1 dinner). Days run Sat, Sun, Mon..Fri, Sat, Sun.
const DAYS = ['Sat','Sun','Mon','Tue','Wed','Thu','Fri','Sat','Sun'];

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
function readState() { const v = tab('MealState').getRange('A1').getValue(); return v ? JSON.parse(v) : null; }
function kind(k) { return k === 'home' ? 'Home' : k === 'out' ? 'Eating out' : ''; }

// Current week as a readable table (meals that span several slots repeat in each)
function writeState(s) {
  tab('MealState').getRange('A1').setValue(JSON.stringify(s));
  if (!s.blocks) return; // old v1.1 state, replaced as soon as the new app syncs
  const lo = 4 - 2 * (s.pre || 0), hi = 13 + 2 * (s.post || 0), rows = [['Week of ' + s.ws, s.names[0], s.names[1]]];
  for (let i = lo; i <= hi; i++) {
    const f = c => {
      const b = s.blocks.filter(x => x.s <= i && i < x.s + x.n && (x.l === 'b' || x.l === (c ? 'e' : 't')))[0];
      return b ? b.t + (b.k ? ' (' + kind(b.k) + ')' : '') : '';
    };
    rows.push([DAYS[i >> 1] + (i % 2 ? ' Dinner' : ' Lunch'), f(0), f(1)]);
  }
  const p = tab('Meal Plan'); p.clear(); p.getRange(1, 1, rows.length, 3).setValues(rows);
}

// Archive of all past weeks: one row per meal, never deleted
function archiveTab(names) {
  const sh = tab('Meal Archive');
  if (sh.getLastRow() === 0) {
    sh.appendRow(['Week starting', 'Date', 'Day', 'Meal', 'Meals covered', 'Who', 'Meal name', 'Where', 'Lane']);
    sh.setFrozenRows(1); sh.getRange('A:B').setNumberFormat('@');
  }
  return sh;
}
function upsertArchive(w) {
  const sh = archiveTab(), last = sh.getLastRow(), tz = Session.getScriptTimeZone();
  if (last > 1) {
    const v = sh.getRange(2, 1, last - 1, 1).getValues();
    for (let r = v.length - 1; r >= 0; r--) if (String(v[r][0]) === w.ws) sh.deleteRow(r + 2);
  }
  const rows = w.blocks.map(b => [w.ws, b.d, Utilities.formatDate(new Date(b.d + 'T12:00:00'), tz, 'EEE'),
    b.m ? 'Dinner' : 'Lunch', b.n, b.l === 'b' ? 'Both' : w.names[b.l === 't' ? 0 : 1], b.t, kind(b.k), b.l]);
  if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, 9).setValues(rows);
}
function readArchive() {
  const sh = archiveTab(), last = sh.getLastRow(), m = {};
  if (last < 2) return [];
  sh.getRange(2, 1, last - 1, 9).getValues().forEach(r => {
    const ws = String(r[0]); (m[ws] = m[ws] || { ws: ws, blocks: [] }).blocks.push({
      d: String(r[1]), m: r[3] === 'Dinner' ? 1 : 0, n: Number(r[4]) || 1, l: String(r[8]) || 'b',
      t: String(r[6]), k: r[7] === 'Home' ? 'home' : r[7] === 'Eating out' ? 'out' : '' });
  });
  return Object.keys(m).map(k => m[k]);
}
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
