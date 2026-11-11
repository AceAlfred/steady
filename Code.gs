/**
 * Steady → Google Sheets sync
 * Paste this whole file into Extensions → Apps Script (replace everything in Code.gs),
 * run setup() once, then Deploy → New deployment → Web app.
 */

const TABS = {
  'Daily':    ['Date', 'Day', 'Weight kg', 'Calories', 'Protein g', 'Carbs g', 'Water (glasses)', 'Active min', 'Burned kcal', 'Habits done', 'Habits', 'Updated'],
  'Meals':    ['Date', 'Meal', 'Food', 'Calories', 'Protein g', 'Carbs g', 'ID'],
  'Workouts': ['Date', 'Type', 'Minutes', 'Intensity', 'Burned kcal', 'ID'],
  'Fasts':    ['Start', 'End', 'Hours', 'Goal h', 'Completed'],
  'Settings': ['Setting', 'Value'],
};
const STATE = '_state';          // raw app data (lets the app restore on a new phone)
const CHUNK = 45000;             // a cell holds max 50k characters
const HABIT_NAMES = {weigh:'Weigh-in', protein:'Protein every meal', plate:'Plate rule', nocarb:'No bread/rice/pasta/fries', move:'Moved', stretch:'Stretched', water:'Water goal', closed:'Kitchen closed'};

/* ---------- one-time setup: run this from the editor ---------- */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(Session.getScriptTimeZone() || 'Europe/Stockholm');
  Object.keys(TABS).forEach(name => {
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    sh.getRange(1, 1, 1, TABS[name].length).setValues([TABS[name]])
      .setFontWeight('bold').setBackground('#1e2723').setFontColor('#c6f432');
    sh.setFrozenRows(1);
  });
  ss.getSheetByName('Daily').getRange('A2:A').setNumberFormat('yyyy-mm-dd');
  ss.getSheetByName('Meals').getRange('A2:A').setNumberFormat('yyyy-mm-dd');
  ss.getSheetByName('Workouts').getRange('A2:A').setNumberFormat('yyyy-mm-dd');
  ss.getSheetByName('Fasts').getRange('A2:B').setNumberFormat('yyyy-mm-dd hh:mm');
  const st = ss.getSheetByName(STATE) || ss.insertSheet(STATE);
  st.hideSheet();

  // remove the empty default tab
  ['Sheet1', 'Blad1', 'Ark1'].forEach(n => { const s = ss.getSheetByName(n); if (s && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s); });

  // weight chart
  const daily = ss.getSheetByName('Daily');
  if (daily.getCharts().length === 0) {
    daily.insertChart(daily.newChart().asLineChart()
      .addRange(daily.getRange('A1:A1000')).addRange(daily.getRange('C1:C1000'))
      .setOption('title', 'Weight (kg)').setOption('interpolateNulls', true)
      .setOption('legend', {position: 'none'}).setOption('colors', ['#5c8a1f'])
      .setPosition(2, 14, 0, 0).build());
  }

  // secret sync key
  const props = PropertiesService.getScriptProperties();
  let key = props.getProperty('SYNC_KEY');
  if (!key) { key = Utilities.getUuid().replace(/-/g, '').slice(0, 20); props.setProperty('SYNC_KEY', key); }

  const start = ss.getSheetByName('Start here') || ss.insertSheet('Start here', 0);
  start.clear();
  start.getRange('A1:B6').setValues([
    ['Steady sync', ''],
    ['Your sync key', key],
    ['', ''],
    ['Next step', 'Deploy → New deployment → type: Web app → Execute as: Me → Who has access: Anyone → Deploy. Copy the Web app URL.'],
    ['Then', 'In the Steady app: Coach tab → Google Sheets sync → paste the URL and the sync key → Connect.'],
    ['Keep private', 'Anyone with both the URL and the key can read/write your log. Don\'t share them.'],
  ]);
  start.getRange('A1').setFontSize(16).setFontWeight('bold');
  start.getRange('A2:A6').setFontWeight('bold');
  start.getRange('B2').setFontFamily('Roboto Mono').setFontSize(13).setBackground('#eef9d0');
  start.setColumnWidth(1, 140); start.setColumnWidth(2, 640);
  start.getRange('B4:B6').setWrap(true);
  Logger.log('Setup done. Your sync key: ' + key);
}

/* ---------- web endpoints ---------- */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!p.key) return json_({ok: true, message: 'Steady sync is running ✓ Connect from the app (Coach → Google Sheets sync).'});
  if (!authorized_(p.key)) return json_({ok: false, error: 'Wrong sync key'});
  if (p.action === 'pull') return json_({ok: true, items: readState_()});
  return json_({ok: true, app: 'steady'});
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ok: false, error: 'Bad JSON'}); }
  if (!authorized_(body.key)) return json_({ok: false, error: 'Wrong sync key'});
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const saved = writeState_(body.items || {});
    saved.forEach(k => project_(k, body.items[k].v));
    return json_({ok: true, saved: saved});
  } catch (err) {
    return json_({ok: false, error: String(err)});
  } finally {
    lock.releaseLock();
  }
}

function authorized_(key) {
  const real = PropertiesService.getScriptProperties().getProperty('SYNC_KEY');
  return real && key && String(key) === real;
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* ---------- raw state (key | updatedAt | json chunks…) ---------- */
function stateSheet_() { const ss = SpreadsheetApp.getActiveSpreadsheet(); return ss.getSheetByName(STATE) || ss.insertSheet(STATE); }

function readState_() {
  const sh = stateSheet_(), out = {};
  if (sh.getLastRow() < 1) return out;
  sh.getDataRange().getValues().forEach(r => {
    if (!r[0]) return;
    try { out[r[0]] = {t: Number(r[1]) || 0, v: JSON.parse(r.slice(2).join(''))}; } catch (e) {}
  });
  return out;
}

function writeState_(items) {
  const sh = stateSheet_();
  const rows = sh.getLastRow() ? sh.getRange(1, 1, sh.getLastRow(), 2).getValues() : [];
  const index = {}; rows.forEach((r, i) => { if (r[0]) index[r[0]] = {row: i + 1, t: Number(r[1]) || 0}; });
  const saved = [];
  Object.keys(items).forEach(k => {
    const it = items[k]; if (!it || typeof it.t !== 'number') return;
    if (index[k] && index[k].t > it.t) return;              // sheet already has newer data
    const s = JSON.stringify(it.v), parts = [];
    for (let i = 0; i < s.length; i += CHUNK) parts.push(s.slice(i, i + CHUNK));
    const row = [k, it.t].concat(parts);
    let r = index[k] ? index[k].row : sh.getLastRow() + 1;
    if (index[k]) sh.getRange(r, 1, 1, sh.getMaxColumns()).clearContent();
    if (sh.getMaxColumns() < row.length) sh.insertColumnsAfter(sh.getMaxColumns(), row.length - sh.getMaxColumns());
    sh.getRange(r, 1, 1, row.length).setNumberFormat('@').setValues([row]);
    index[k] = {row: r, t: it.t};
    saved.push(k);
  });
  return saved;
}

/* ---------- readable tabs ---------- */
function project_(key, v) {
  if (!v) return;
  if (key.indexOf('day-') === 0) return projectDay_(key.slice(4), v);
  if (key === 'fast') return projectFasts_(v);
  if (key === 'profile') return projectProfile_(v);
  if (key === 'weights') return projectWeights_(v);
}

let TZ_ = null;
function tz_() { return TZ_ || (TZ_ = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone()); }
function toDate_(iso) { const p = iso.split('-').map(Number); return new Date(p[0], p[1] - 1, p[2]); }
function isoOf_(v) { return v instanceof Date ? Utilities.formatDate(v, tz_(), 'yyyy-MM-dd') : String(v).slice(0, 10); }
function sheet_(name) { const ss = SpreadsheetApp.getActiveSpreadsheet(); let sh = ss.getSheetByName(name); if (!sh) { setup(); sh = ss.getSheetByName(name); } return sh; }

// replace every row for one date in a tab
function replaceRowsForDate_(name, iso, newRows) {
  const sh = sheet_(name), width = TABS[name].length, last = sh.getLastRow();
  let keep = last > 1 ? sh.getRange(2, 1, last - 1, width).getValues().filter(r => r[0] !== '' && isoOf_(r[0]) !== iso) : [];
  keep = keep.concat(newRows).sort((a, b) => isoOf_(a[0]) < isoOf_(b[0]) ? -1 : isoOf_(a[0]) > isoOf_(b[0]) ? 1 : 0);
  if (last > 1) sh.getRange(2, 1, last - 1, width).clearContent();
  if (keep.length) sh.getRange(2, 1, keep.length, width).setValues(keep);
}

function upsertDaily_(iso, patch) {
  const sh = sheet_('Daily'), width = TABS.Daily.length, last = sh.getLastRow();
  const vals = last > 1 ? sh.getRange(2, 1, last - 1, width).getValues() : [];
  let i = vals.findIndex(r => r[0] !== '' && isoOf_(r[0]) === iso);
  const d = toDate_(iso);
  let row = i >= 0 ? vals[i] : [d, Utilities.formatDate(d, tz_(), 'EEE'), '', '', '', '', '', '', '', '', '', ''];
  Object.keys(patch).forEach(col => { row[TABS.Daily.indexOf(col)] = patch[col]; });
  row[TABS.Daily.indexOf('Updated')] = new Date();
  if (i >= 0) { sh.getRange(i + 2, 1, 1, width).setValues([row]); return; }
  // insert keeping date order
  let pos = vals.findIndex(r => r[0] !== '' && isoOf_(r[0]) > iso);
  if (pos < 0) sh.getRange(last + 1, 1, 1, width).setValues([row]);
  else { sh.insertRowBefore(pos + 2); sh.getRange(pos + 2, 1, 1, width).setValues([row]); }
}

function projectDay_(iso, day) {
  const d = toDate_(iso);
  const meals = (day.meals || []).map(m => [d, m.slot, m.n, Math.round(m.k), round1_(m.p), round1_(m.c), m.id]);
  const moves = (day.moves || []).map(m => [d, m.type === 'NTC' ? 'Nike Training Club' : m.type, m.min, ({easy: 'Easy', mod: 'Moderate', hard: 'Hard'})[m.int] || m.int, Math.round(m.k), m.id]);
  replaceRowsForDate_('Meals', iso, meals);
  replaceRowsForDate_('Workouts', iso, moves);
  const sum = (a, f) => a.reduce((s, x) => s + (Number(x[f]) || 0), 0);
  const checks = Object.keys(day.checks || {}).filter(k => day.checks[k]);
  if (day.weight) checks.push('weigh');
  if ((day.moves || []).some(m => m.type === 'Bike' || m.type === 'NTC')) checks.push('move');
  if ((day.moves || []).some(m => m.type === 'Stretch')) checks.push('stretch');
  const habits = Array.from(new Set(checks));
  const patch = {
    'Calories': Math.round(sum(day.meals || [], 'k')),
    'Protein g': Math.round(sum(day.meals || [], 'p')),
    'Carbs g': Math.round(sum(day.meals || [], 'c')),
    'Water (glasses)': day.water || 0,
    'Active min': sum(day.moves || [], 'min'),
    'Burned kcal': Math.round(sum(day.moves || [], 'k')),
    'Habits done': habits.length,
    'Habits': habits.map(h => HABIT_NAMES[h] || h).join(', '),
  };
  if (day.weight) patch['Weight kg'] = day.weight;   // never blank out a weight already logged
  upsertDaily_(iso, patch);
}

function projectWeights_(w) {
  const entries = (w && w.entries) || {};
  const sh = sheet_('Daily'), last = sh.getLastRow();
  const have = {};
  if (last > 1) sh.getRange(2, 1, last - 1, 3).getValues().forEach(r => { if (r[0] !== '') have[isoOf_(r[0])] = r[2]; });
  Object.keys(entries).forEach(iso => { if (String(have[iso]) !== String(entries[iso])) upsertDaily_(iso, {'Weight kg': Number(entries[iso])}); });
}

function projectFasts_(f) {
  const sh = sheet_('Fasts'), width = TABS.Fasts.length, last = sh.getLastRow();
  const rows = (f.log || []).map(x => [new Date(x.start), new Date(x.end), round1_((x.end - x.start) / 36e5), x.target, (x.end - x.start) >= x.target * 36e5 ? 'Yes' : 'No']);
  if (f.active) rows.push([new Date(f.active.start), '', 'running…', f.active.target, '']);
  if (last > 1) sh.getRange(2, 1, last - 1, width).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, width).setValues(rows);
}

function projectProfile_(p) {
  const days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
  const rows = [
    ['Name', p.name || ''], ['Age', p.age || ''], ['Height cm', p.height || ''],
    ['Start weight kg', p.weight || ''], ['Goal weight kg', p.goal || ''],
    ['Calories / day', p.kcal], ['Protein g / day', p.protein], ['Carb limit g / day', p.carbs],
    ['Water glasses / day', p.water], ['Active min / week', p.weekMin], ['Kitchen closes', p.close],
    ['Fast days', (p.fastDays || []).map(i => days[i]).join(', ')], ['Fast length h', p.fastH],
    ['My foods', (p.myFoods || []).map(f => f.n + ' (' + f.k + ' kcal)').join('; ')],
  ];
  const sh = sheet_('Settings'), last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, 2).clearContent();
  sh.getRange(2, 1, rows.length, 2).setValues(rows);
}

function round1_(n) { return Math.round((Number(n) || 0) * 10) / 10; }

/* Optional: rebuild every readable tab from the raw data (run from the editor if a tab gets messed up) */
function rebuildAll() {
  const items = readState_();
  ['Meals', 'Workouts', 'Daily'].forEach(n => { const sh = sheet_(n); if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, TABS[n].length).clearContent(); });
  Object.keys(items).sort().forEach(k => project_(k, items[k].v));
}
