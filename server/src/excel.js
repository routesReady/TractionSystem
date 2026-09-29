const XLSX = require('xlsx');
const zlib = require('zlib');

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
  dayWork: ['Day_of_wkg_in_territory', 'DAY OF WKG IN TERRITORY'], locoNo: ['Loco_No', 'LOCO NO', 'LOCO_NO', 'LOCO NUMBER'],
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
      sourceRow: i + 1, date: get('date'), day: get('day'), trainNo: get('trainNo'), dir: get('dir'), trType: get('trType'), time: get('time'),
      attachedDiv: get('attachedDiv'), fromTo: get('fromTo'), locoLink: get('locoLink'), locoLinkDiv: get('locoLinkDiv'),
      dayWork: get('dayWork'), locoNo: get('locoNo'), shed: get('shed'), locoType: get('locoType'), rly: get('rly'),
      kavachMake: get('kavachMake'), brakeSystem: get('brakeSystem'), kavachSection: get('kavachSection'), worked: get('worked'),
      noReason: get('noReason'), remarks: get('remarks')
    });
  }
  return out;
}

// Excel list validation values are stored in the workbook XML, not in the
// cells. This reads direct list validations such as "YES, NO" and the reason
// list used by the TRAINS_PER_DAY workbook.
function zipEntries(buf) {
  const eocdSig = 0x06054b50;
  const cdSig = 0x02014b50;
  const eocd = buf.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06]));
  if (eocd < 0) return [];
  const count = buf.readUInt16LE(eocd + 10);
  const offset = buf.readUInt32LE(eocd + 16);
  const out = [];
  let p = offset;
  for (let i = 0; i < count; i++) {
    if (buf.readUInt32LE(p) !== cdSig) break;
    const method = buf.readUInt16LE(p + 10);
    const compressedSize = buf.readUInt32LE(p + 20);
    const uncompressedSize = buf.readUInt32LE(p + 24);
    const nameLen = buf.readUInt16LE(p + 28);
    const extraLen = buf.readUInt16LE(p + 30);
    const commentLen = buf.readUInt16LE(p + 32);
    const localOffset = buf.readUInt32LE(p + 42);
    const name = buf.slice(p + 46, p + 46 + nameLen).toString('utf8');
    out.push({ name, method, compressedSize, uncompressedSize, localOffset });
    p += 46 + nameLen + extraLen + commentLen;
  }
  return out;
}

function zipRead(buf, name) {
  const entry = zipEntries(buf).find(x => x.name === name);
  if (!entry) return null;
  const p = entry.localOffset;
  if (buf.readUInt32LE(p) !== 0x04034b50) return null;
  const nameLen = buf.readUInt16LE(p + 26);
  const extraLen = buf.readUInt16LE(p + 28);
  const start = p + 30 + nameLen + extraLen;
  const raw = buf.slice(start, start + entry.compressedSize);
  if (entry.method === 0) return raw;
  if (entry.method === 8) return zlib.inflateRawSync(raw);
  return null;
}

function xmlUnescape(s) {
  return String(s || '')
    .replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>').replace(/&amp;/g, '&');
}

function validationColumn(sqref) {
  const m = String(sqref || '').match(/(?:^|\s)([A-Z]{1,3})\d+(?::[A-Z]{1,3}\d+)?/i);
  if (!m) return -1;
  let n = 0;
  for (const ch of m[1].toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64;
  return n - 1;
}

function directValidationList(formula) {
  let s = xmlUnescape(formula).trim();
  if (!/^"[\s\S]*"$/.test(s)) return [];
  s = s.slice(1, -1);
  return s.split(',').map(x => x.trim()).filter(Boolean);
}

function trainValidationOptions(buf) {
  const out = { worked: [], noReason: [] };
  try {
    const entries = zipEntries(buf);
    const sheetNames = entries.filter(x => /^xl\/worksheets\/sheet\d+\.xml$/i.test(x.name)).map(x => x.name);
    for (const sheetName of sheetNames) {
      const xmlBuf = zipRead(buf, sheetName);
      if (!xmlBuf) continue;
      const xml = xmlBuf.toString('utf8');
      const matches = xml.match(/<dataValidation\b[\s\S]*?<\/dataValidation>/gi) || [];
      for (const block of matches) {
        if (!/\btype=["']list["']/i.test(block)) continue;
        const sq = block.match(/\bsqref=["']([^"']+)["']/i);
        const f = block.match(/<formula1\b[^>]*>([\s\S]*?)<\/formula1>/i);
        if (!sq || !f) continue;
        const idx = validationColumn(sq[1]);
        const values = directValidationList(f[1]);
        if (!values.length) continue;
        if (idx === 18) out.worked.push(...values); // S
        if (idx === 19) out.noReason.push(...values); // T
      }
    }
  } catch {}
  const uniq = xs => Array.from(new Set(xs.map(clean).filter(Boolean)));
  return { worked: uniq(out.worked), noReason: uniq(out.noReason) };
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
  'Train_worked_with_Kavach_YES_NO','If_No_the_Reason \n( Non Kavach Loco\n/ Kavach defective\n/ KAVACH Fitness cirtification N.Avl.\n/ Crew incompetency)','Remarks','Shed_Remark','OEM'
];

module.exports = { trains, kavach, norm, tval, validDate, weekday, TEMPLATE_HEADERS, trainValidationOptions };
