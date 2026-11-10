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
  if (!w.SheetNames.length) throw new Error('Excel workbook contains no worksheets.');
  return XLSX.utils.sheet_to_json(w.Sheets[w.SheetNames[0]], { header: 1, defval: '', raw: false });
}

// Retain every source column (including future columns) alongside the mapped
// operational fields. Duplicate/blank headers are made safe without changing
// the displayed source label or shifting row values.
function sourceColumns(buf) {
  const m = matrix(buf);
  if (!m.length) throw new Error('Excel workbook is empty or has no header row.');
  const seen = new Map();
  return m[0].map((h, i) => {
    const label = clean(h);
    // Ignore blank Excel header cells instead of exposing generated
    // "Unnamed column N" fields in tables and exports. Keep the original
    // index so values in neighboring columns never shift.
    if (!label) return null;
    const count = (seen.get(label.toUpperCase()) || 0) + 1;
    seen.set(label.toUpperCase(), count);
    return { key: count === 1 ? label : `${label} (${count})`, label, index: i };
  }).filter(Boolean);
}
function sourceValues(row, columns) {
  const out = {};
  columns.forEach(c => { out[c.key] = clean(row[c.index]); });
  return out;
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
  const columns = sourceColumns(buf);
  const hm = headerMap(m[0]);
  const ix = Object.fromEntries(Object.entries(TRAIN).map(([name, aliases]) => [name, findIndex(hm, aliases)]));
  for (let i = 1; i < m.length; i++) {
    const r = m[i];
    if (r.every(x => clean(x) === '')) continue;
    const get = name => ix[name] >= 0 ? clean(r[ix[name]]) : '';
    out.push({
      sourceRow: i + 1, sourceData: sourceValues(r, columns), date: get('date'), day: get('day'), trainNo: get('trainNo'), dir: get('dir'), trType: get('trType'), time: get('time'),
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

function columnIndexFromCell(ref) {
  const match = String(ref || '').match(/^\$?([A-Z]{1,3})\$?\d+/i);
  if (!match) return -1;
  let n = 0;
  for (const ch of match[1].toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64;
  return n - 1;
}

function columnHeaderIndex(buf, target) {
  const m = matrix(buf);
  const wanted = key(target);
  return m.length ? m[0].findIndex(h => key(h) === wanted) : -1;
}

function workbookRangeValues(wb, formula, currentSheetName) {
  let expr = xmlUnescape(String(formula || '')).trim().replace(/^=/, '');
  if (!expr || /^"[\s\S]*"$/.test(expr)) return [];
  const names = wb.Workbook?.Names || [];
  const named = names.find(n => key(n.Name) === key(expr) || key(n.Name) === key(expr.replace(/^.*!/, '')));
  if (named?.Ref) expr = String(named.Ref).replace(/^=/, '');
  const match = expr.match(/^(?:'((?:[^']|'')+)'|([^!]+))?!?\s*(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)$/i);
  if (!match) return [];
  const sheetName = (match[1] || match[2] || currentSheetName || '').replace(/''/g, "'").trim();
  const ref = match[3];
  const sheet = wb.Sheets[sheetName] || (wb.SheetNames.includes(currentSheetName) ? wb.Sheets[currentSheetName] : null);
  if (!sheet) return [];
  const range = XLSX.utils.decode_range(ref.replace(/\$/g, ''));
  const out = [];
  for (let r = range.s.r; r <= range.e.r; r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = sheet[XLSX.utils.encode_cell({ r, c })];
      const value = clean(cell?.w ?? cell?.v);
      if (value) out.push(value);
    }
  }
  return out;
}

// Read list validations from the uploaded workbook itself. Associate each
// validation with the column's actual header, not a fixed S/T position, and
// resolve both inline lists and worksheet/named-range references.
function trainValidationOptions(buf) {
  const result = { worked: [], noReason: [] };
  try {
    const wb = XLSX.read(buf, { type: 'buffer', cellDates: true, raw: false, bookFiles: true });
    const currentSheetName = wb.SheetNames[0];
    const headerRow = currentSheetName ? XLSX.utils.sheet_to_json(wb.Sheets[currentSheetName], { header: 1, defval: '', range: 0 })[0] || [] : [];
    const wantedHeaders = new Map();
    headerRow.forEach((h, i) => {
      const hkey = key(h);
      if (!hkey) return;
      if (hkey.includes('TRAIN WORKED WITH KAVACH')) wantedHeaders.set(i, 'worked');
      else if (hkey.includes('IF NO THE REASON') || hkey.includes('IF NO REASON')) wantedHeaders.set(i, 'noReason');
    });
    const entries = zipEntries(buf);
    const sheetNames = entries.filter(x => /^xl\/worksheets\/sheet\d+\.xml$/i.test(x.name)).map(x => x.name);
    for (let si = 0; si < sheetNames.length; si++) {
      const xmlBuf = zipRead(buf, sheetNames[si]);
      if (!xmlBuf) continue;
      const xml = xmlBuf.toString('utf8');
      const matches = xml.match(/<dataValidation\b[\s\S]*?<\/dataValidation>/gi) || [];
      for (const block of matches) {
        if (!/\btype=["']list["']/i.test(block)) continue;
        const sq = block.match(/\bsqref=["']([^"']+)["']/i);
        const f = block.match(/<formula1\b[^>]*>([\s\S]*?)<\/formula1>/i);
        if (!sq || !f) continue;
        const ref = sq[1].trim().split(/\s+/)[0];
        const idx = columnIndexFromCell(ref);
        const kind = wantedHeaders.get(idx);
        if (!kind) continue;
        const formula = xmlUnescape(f[1]).trim();
        let values = directValidationList(formula);
        if (!values.length) values = workbookRangeValues(wb, formula, wb.SheetNames[si] || currentSheetName);
        if (values.length) result[kind].push(...values);
      }
    }
  } catch (error) {
    // Invalid/unsupported validation metadata is handled by the caller's
    // per-column fallback; never let a parser issue break workbook upload.
  }
  const uniqueInOrder = values => {
    const seen = new Set();
    return values.map(clean).filter(value => {
      if (!value || seen.has(value)) return false;
      seen.add(value);
      return true;
    });
  };
  return { worked: uniqueInOrder(result.worked), noReason: uniqueInOrder(result.noReason) };
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
  const columns = sourceColumns(buf);
  const hm = headerMap(m[0]);
  const ix = Object.fromEntries(Object.entries(KAVACH).map(([name, aliases]) => [name, findIndex(hm, aliases)]));
  if (ix.locoNo < 0) throw new Error('Kavach_Loco_Details.xlsx must contain a Loco_No column.');
  for (let i = 1; i < m.length; i++) {
    const r = m[i];
    if (r.every(x => clean(x) === '')) continue;
    const get = name => ix[name] >= 0 ? clean(r[ix[name]]) : '';
    out.push({ sourceRow: i + 1, sourceData: sourceValues(r, columns), locoNo: get('locoNo'), shed: get('shed'), locoType: get('locoType'), rly: get('rly'), kavachMake: get('kavachMake'), brakeSystem: get('brakeSystem') });
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
  'Train_worked_with_Kavach_YES_NO','If_No_the_Reason \n( Non Kavach Loco\n/ Kavach defective\n/ KAVACH Fitness cirtification N.Avl.\n/ Crew incompetency)','Remarks','Shed Remarks','OEM Remarks'
];

module.exports = { trains, kavach, norm, tval, validDate, weekday, TEMPLATE_HEADERS, trainValidationOptions, sourceColumns };
