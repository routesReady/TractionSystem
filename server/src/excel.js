const XLSX = require('xlsx');

const clean = (value) => value === undefined || value === null ? '' : String(value).trim();
const key = (value) => clean(value)
  .replace(/\r?\n/g, ' ')
  .replace(/[_./\\-]+/g, ' ')
  .replace(/\s+/g, ' ')
  .toUpperCase()
  .trim();

function matrix(buf) {
  const w = XLSX.read(buf, { type: 'buffer', cellDates: true, raw: false });
  return XLSX.utils.sheet_to_json(w.Sheets[w.SheetNames[0]], { header: 1, defval: '', raw: false });
}

function headerMap(headerRow) {
  const map = new Map();
  headerRow.forEach((h, i) => { const k = key(h); if (k) map.set(k, i); });
  return map;
}

function findIndex(map, aliases) {
  for (const alias of aliases) {
    const idx = map.get(key(alias));
    if (idx !== undefined) return idx;
  }
  // Fuzzy fallback for headers such as "Loco No." / "LOCO_NO".
  for (const [h, idx] of map.entries()) {
    if (aliases.some(a => h.includes(key(a)) || key(a).includes(h))) return idx;
  }
  return -1;
}

const TRAIN = {
  date: ['DATE'], day: ['DAY'], trainNo: ['TRAIN_No', 'TRAIN NO', 'TRAIN_NO', 'TRAIN NUMBER'],
  dir: ['DIR', 'DIRECTION'], trType: ['TR_TYPE', 'TR TYPE', 'TRAIN TYPE'], time: ['T_O_Time', 'T O TIME', 'TO TIME'],
  attachedDiv: ['LOCO_ATTACHED_DIV.', 'LOCO ATTACHED DIV', 'LOCO_ATTACHED_DIV'], fromTo: ['From_To', 'FROM TO', 'FROM_TO'],
  locoLink: ['LOCO_link', 'LOCO LINK', 'LOCO_LINK'], locoLinkDiv: ['LOCO_LINK_DIV', 'LOCO LINK DIV'],
  dayWork: ['Day_of_wkg_in_territory', 'DAY OF WKG IN TERRITORY'], locoNo: ['Loco_No', 'LOCO NO', 'LOCO_NO'],
  shed: ['SHED'], locoType: ['Loco_Type', 'LOCO TYPE'], rly: ['RLY'], kavachMake: ['Loco_Kavach_Make', 'LOCO KAVACH MAKE'],
  brakeSystem: ['Brake_system', 'BRAKE SYSTEM'], kavachSection: ['KAVACH_WKG_SECTION', 'KAVACH WKG SECTION'],
  worked: ['Train_worked_with_Kavach_YES_NO', 'TRAIN WORKED WITH KAVACH YES NO'],
  noReason: ['If_No_the_Reason', 'IF NO THE REASON'], remarks: ['Remarks']
};

function trains(buf) {
  const m = matrix(buf), out = [];
  if (!m.length) return out;
  const hm = headerMap(m[0]);
  const ix = Object.fromEntries(Object.entries(TRAIN).map(([name, aliases]) => [name, findIndex(hm, aliases)]));
  for (let i = 1; i < m.length; i++) {
    const r = m[i];
    if (r.every(x => clean(x) === '')) continue;
    const get = name => ix[name] >= 0 ? clean(r[ix[name]]) : '';
    out.push({
      sourceRow: i + 1,
      date: get('date'), day: get('day'), trainNo: get('trainNo'), dir: get('dir'), trType: get('trType'), time: get('time'),
      attachedDiv: get('attachedDiv'), fromTo: get('fromTo'), locoLink: get('locoLink'), locoLinkDiv: get('locoLinkDiv'),
      dayWork: get('dayWork'), locoNo: get('locoNo'), shed: get('shed'), locoType: get('locoType'), rly: get('rly'),
      kavachMake: get('kavachMake'), brakeSystem: get('brakeSystem'), kavachSection: get('kavachSection'), worked: get('worked'),
      noReason: get('noReason'), remarks: get('remarks')
    });
  }
  return out;
}

const KAVACH = {
  locoNo: ['Loco_No', 'LOCO NO', 'LOCO_NO', 'LOCO NUMBER', 'LOCO NUMBER.'],
  shed: ['SHED'], locoType: ['Loco_Type', 'LOCO TYPE'], rly: ['RLY', 'RAILWAY'],
  kavachMake: ['Loco_Kavach_Make', 'LOCO KAVACH MAKE', 'KAVACH MAKE'],
  brakeSystem: ['Brake_system', 'BRAKE SYSTEM', 'BRAKE_SYSTEM']
};

function kavach(buf) {
  const m = matrix(buf), out = [];
  if (!m.length) return out;
  const hm = headerMap(m[0]);
  const ix = Object.fromEntries(Object.entries(KAVACH).map(([name, aliases]) => [name, findIndex(hm, aliases)]));
  if (ix.locoNo < 0) throw new Error('Kavach_Loco_Details.xlsx must contain a Loco_No column.');
  for (let i = 1; i < m.length; i++) {
    const r = m[i];
    if (r.every(x => clean(x) === '')) continue;
    const get = name => ix[name] >= 0 ? clean(r[ix[name]]) : '';
    out.push({ sourceRow: i + 1, locoNo: get('locoNo'), shed: get('shed'), locoType: get('locoType'), rly: get('rly'), kavachMake: get('kavachMake'), brakeSystem: get('brakeSystem') });
  }
  return out;
}

function norm(x) {
  let s = clean(x);
  if (!s) return '';
  if (/^\d+(?:\.0+)?$/.test(s)) s = s.replace(/\.0+$/, '');
  return s.toUpperCase().replace(/\s+/g, '');
}
function tval(x) {
  const m = clean(x).match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?/);
  return m ? Number(m[1]) * 3600 + Number(m[2]) * 60 + Number(m[3] || 0) : Infinity;
}
function validDate(s) {
  if (!/^\d{2}-\d{2}-\d{4}$/.test(s)) return false;
  const [d, m, y] = s.split('-').map(Number), x = new Date(y, m - 1, d);
  return x.getFullYear() === y && x.getMonth() === m - 1 && x.getDate() === d;
}
function weekday(s) {
  const [d, m, y] = s.split('-').map(Number);
  return new Date(y, m - 1, d).toLocaleDateString('en-US', { weekday: 'long' }).toUpperCase();
}

const TEMPLATE_HEADERS = [
  'S.N','DATE','DAY','TRAIN_No','DIR','TR_TYPE','T_O_Time','LOCO_ATTACHED_DIV.','From_To','LOCO_link','LOCO_LINK_DIV',
  'Day_of_wkg_in_territory','Loco_No','SHED','Loco_Type','RLY','Loco_Kavach_Make','Brake_system','KAVACH_WKG_SECTION',
  'Train_worked_with_Kavach_YES_NO','If_No_the_Reason \n( Non Kavach Loco\n/ Kavach defective\n/ KAVACH Fitness cirtification N.Avl.\n/ Crew incompetency)','Remarks'
];

module.exports = { trains, kavach, norm, tval, validDate, weekday, TEMPLATE_HEADERS };
