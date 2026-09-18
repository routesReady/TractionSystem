require('dotenv').config();
const dns = require('dns');
dns.setServers(['8.8.8.8', '8.8.4.4']);
const express = require('express');
const cors = require('cors');
const mongoose = require('mongoose');
const multer = require('multer');
const jwt = require('jsonwebtoken');
const XLSX = require('xlsx');
const { SourceDataset, SavedRow } = require('./models');
const { trains, kavach, norm, tval, validDate, weekday, TEMPLATE_HEADERS, trainValidationOptions } = require('./excel');

const app = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 15 * 1024 * 1024 } });

const allowedOrigins = String(process.env.CLIENT_ORIGIN || 'http://localhost:5173')
  .split(',').map(v => v.trim().replace(/\/+$/, '')).filter(Boolean);
const vercelOrigin = /^https:\/\/(?:traction-operation|traction-system)-[a-z0-9-]+\.vercel\.app$/i;

app.use(cors({
  origin(origin, callback) {
    if (!origin || allowedOrigins.includes(origin) || vercelOrigin.test(origin) || origin === 'https://traction-operation.vercel.app') return callback(null, true);
    return callback(new Error('CORS origin not allowed'));
  }
}));
app.use(express.json({ limit: '2mb' }));

const auth = (req, res, next) => {
  try {
    const header = req.headers.authorization || '';
    if (!header.startsWith('Bearer ')) throw new Error('Missing token');
    req.user = jwt.verify(header.slice(7), process.env.JWT_SECRET);
    next();
  } catch { res.status(401).json({ message: 'Authentication required.' }); }
};

const isExcel = name => /\.(xlsx|xls|xlsm)$/i.test(name || '');
const UPLOAD_SECRET = process.env.UPLOAD_PASSWORD || '';
const authenticatePassword = password => Boolean(UPLOAD_SECRET && password && password === UPLOAD_SECRET);

const saveSource = async (key, fileName, rows, buffer, contentType, dropdowns = { worked: [], noReason: [] }) => SourceDataset.findOneAndUpdate(
  { key },
  { key, fileName, rows, fileData: buffer, contentType: contentType || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', dropdowns, uploadedAt: new Date() },
  { upsert: true, new: true }
);

function toBuffer(value) {
  if (!value) return null;
  if (Buffer.isBuffer(value)) return value;
  if (value._bsontype === 'Binary' && value.buffer) return Buffer.from(value.buffer);
  if (value.type === 'Buffer' && Array.isArray(value.data)) return Buffer.from(value.data);
  if (value.buffer && Buffer.isBuffer(value.buffer)) return Buffer.from(value.buffer);
  try { return Buffer.from(value); } catch { return null; }
}

app.get('/', (req, res) => res.json({ ok: true, service: 'Traction Operation API' }));
app.get('/api/health', (req, res) => res.json({ ok: true }));

app.get('/api/upload/status', async (req, res) => {
  try {
    const train = await SourceDataset.findOne({ key: 'TRAINS_PER_DAY' }).lean();
    const kav = await SourceDataset.findOne({ key: 'KAVACH_LOCO_DETAILS' }).lean();
    res.json({
      initialized: Boolean(train && kav),
      train: train && { fileName: train.fileName, uploadedAt: train.uploadedAt, count: train.rows.length, hasFile: Boolean(train.fileData), dropdowns: train.dropdowns || { worked: [], noReason: [] } },
      kavach: kav && { fileName: kav.fileName, uploadedAt: kav.uploadedAt, count: kav.rows.length, hasFile: Boolean(kav.fileData) }
    });
  } catch (e) { res.status(500).json({ message: e.message }); }
});

app.post('/api/upload/authenticate', async (req, res) => {
  try {
    if (!authenticatePassword(req.body?.password)) return res.status(401).json({ message: 'Incorrect upload password.' });
    res.json({ token: jwt.sign({ role: 'uploader' }, process.env.JWT_SECRET, { expiresIn: '30m' }) });
  } catch (e) { res.status(500).json({ message: e.message }); }
});

app.post('/api/upload/initial', auth, upload.fields([{ name: 'trainsFile', maxCount: 1 }, { name: 'kavachFile', maxCount: 1 }]), async (req, res) => {
  try {
    const trainFile = req.files?.trainsFile?.[0];
    const kavachFile = req.files?.kavachFile?.[0];
    if (!trainFile || !kavachFile || !isExcel(trainFile.originalname) || !isExcel(kavachFile.originalname)) return res.status(400).json({ message: 'Both Excel files are required.' });
    const existingTrain = await SourceDataset.exists({ key: 'TRAINS_PER_DAY' });
    const existingKavach = await SourceDataset.exists({ key: 'KAVACH_LOCO_DETAILS' });
    if (existingTrain && existingKavach) return res.status(409).json({ message: 'Both source files already exist. Use Update Source Files.' });
    const trainRows = trains(trainFile.buffer);
    const kavachRows = kavach(kavachFile.buffer);
    const dropdowns = trainValidationOptions(trainFile.buffer);
    await saveSource('TRAINS_PER_DAY', trainFile.originalname, trainRows, trainFile.buffer, trainFile.mimetype, dropdowns);
    await saveSource('KAVACH_LOCO_DETAILS', kavachFile.originalname, kavachRows, kavachFile.buffer, kavachFile.mimetype);
    res.status(201).json({ message: 'Source files uploaded successfully.', trainCount: trainRows.length, kavachCount: kavachRows.length, dropdowns });
  } catch (e) { res.status(500).json({ message: e.message }); }
});

app.post('/api/upload/update', auth, upload.fields([{ name: 'trainsFile', maxCount: 1 }, { name: 'kavachFile', maxCount: 1 }]), async (req, res) => {
  try {
    const trainFile = req.files?.trainsFile?.[0];
    const kavachFile = req.files?.kavachFile?.[0];
    if (!trainFile && !kavachFile) return res.status(400).json({ message: 'Select at least one updated file.' });
    if (trainFile) {
      if (!isExcel(trainFile.originalname)) return res.status(400).json({ message: 'Invalid train Excel file.' });
      const buffer = trainFile.buffer;
      await saveSource('TRAINS_PER_DAY', trainFile.originalname, trains(buffer), buffer, trainFile.mimetype, trainValidationOptions(buffer));
    }
    if (kavachFile) {
      if (!isExcel(kavachFile.originalname)) return res.status(400).json({ message: 'Invalid Kavach Excel file.' });
      await saveSource('KAVACH_LOCO_DETAILS', kavachFile.originalname, kavach(kavachFile.buffer), kavachFile.buffer, kavachFile.mimetype);
    }
    res.json({ message: 'Source data updated successfully.' });
  } catch (e) { res.status(500).json({ message: e.message }); }
});

function generatedSourceWorkbook(source) {
  const data = source.key === 'TRAINS_PER_DAY'
    ? source.rows.map((r, i) => ({
      'S.N': r.sourceRow || i + 1, DATE: r.date, DAY: r.day, TRAIN_No: r.trainNo, DIR: r.dir, TR_TYPE: r.trType, T_O_Time: r.time,
      'LOCO_ATTACHED_DIV.': r.attachedDiv, From_To: r.fromTo, LOCO_link: r.locoLink, LOCO_LINK_DIV: r.locoLinkDiv,
      Day_of_wkg_in_territory: r.dayWork, Loco_No: r.locoNo, SHED: r.shed, Loco_Type: r.locoType, RLY: r.rly,
      Loco_Kavach_Make: r.kavachMake, Brake_system: r.brakeSystem, KAVACH_WKG_SECTION: r.kavachSection,
      Train_worked_with_Kavach_YES_NO: r.worked,
      'If_No_the_Reason \n( Non Kavach Loco\n/ Kavach defective\n/ KAVACH Fitness cirtification N.Avl.\n/ Crew incompetency)': r.noReason,
      Remarks: r.remarks, Shed_Remark: '', OEM: ''
    }))
    : source.rows.map(r => ({ Loco_No: r.locoNo, SHED: r.shed, Loco_Type: r.locoType, RLY: r.rly, Loco_Kavach_Make: r.kavachMake, Brake_system: r.brakeSystem }));
  const headers = source.key === 'TRAINS_PER_DAY' ? TEMPLATE_HEADERS : ['Loco_No', 'SHED', 'Loco_Type', 'RLY', 'Loco_Kavach_Make', 'Brake_system'];
  const ws = XLSX.utils.json_to_sheet(data, { header: headers });
  if (source.key === 'TRAINS_PER_DAY') {
    ws['!cols'] = headers.map((h, i) => ({ wch: i === 20 ? 42 : i === 21 ? 34 : Math.max(12, Math.min(28, String(h).length + 2)) }));
    const worked = source.dropdowns?.worked || ['YES', 'NO'];
    const reasons = source.dropdowns?.noReason || [];
    ws['!dataValidation'] = [];
    // SheetJS community builds do not reliably serialize validation metadata.
    // Stored original workbooks are returned whenever available, so this path is
    // only a compatibility fallback for old database records without fileData.
    void worked; void reasons;
  }
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, source.key === 'TRAINS_PER_DAY' ? 'TRAINS_PER_DAY' : 'Kavach_Loco_Details');
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}

app.get('/api/upload/download/:key', async (req, res) => {
  try {
    const key = req.params.key === 'trains' ? 'TRAINS_PER_DAY' : req.params.key === 'kavach' ? 'KAVACH_LOCO_DETAILS' : null;
    if (!key) return res.status(400).json({ message: 'Unknown source file.' });
    const source = await SourceDataset.findOne({ key }).lean();
    if (!source) return res.status(404).json({ message: 'Source file is not uploaded yet.' });
    let filename = source.fileName || (key === 'TRAINS_PER_DAY' ? 'TRAINS_PER_DAY.xlsx' : 'Kavach_Loco_Details.xlsx');
    let data = toBuffer(source.fileData);
    let generated = false;
    if (!data) { data = generatedSourceWorkbook(source); generated = true; }
    else { try { XLSX.read(data, { type: 'buffer' }); } catch { data = generatedSourceWorkbook(source); generated = true; } }
    if (generated && !/\.xlsx$/i.test(filename)) filename = filename.replace(/\.(xlsx|xls|xlsm)$/i, '') + '.xlsx';
    const extension = /\.xlsm$/i.test(filename) ? 'application/vnd.ms-excel.sheet.macroEnabled.12' : /\.xls$/i.test(filename) ? 'application/vnd.ms-excel' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    const safeName = filename.replace(/[^a-zA-Z0-9._ -]/g, '_');
    res.setHeader('Content-Type', extension); res.setHeader('Content-Disposition', `attachment; filename="${safeName}"`); res.send(data);
  } catch (e) { res.status(500).json({ message: e.message }); }
});

app.get('/api/data/options', async (req, res) => {
  try {
    const source = await SourceDataset.findOne({ key: 'TRAINS_PER_DAY' }).lean();
    if (!source) return res.status(409).json({ message: 'Upload TRAINS_PER_DAY.xlsx first.' });
    let workedOptions = source.dropdowns?.worked || [];
    let reasonOptions = source.dropdowns?.noReason || [];
    if (!workedOptions.length || !reasonOptions.length) {
      const sourceBuffer = toBuffer(source.fileData);
      if (sourceBuffer) {
        const parsed = trainValidationOptions(sourceBuffer);
        if (parsed.worked.length) workedOptions = parsed.worked;
        if (parsed.noReason.length) reasonOptions = parsed.noReason;
      }
    }
    const unique = field => Array.from(new Set(source.rows.map(r => String(r[field] ?? '').trim()).filter(Boolean)));
    if (!workedOptions.length) workedOptions = unique('worked');
    if (!reasonOptions.length) reasonOptions = unique('noReason');
    res.json({ workedOptions, reasonOptions });
  } catch (e) { res.status(500).json({ message: e.message }); }
});

app.get('/api/kavach/lookup', async (req, res) => {
  try {
    const locoNo = norm(req.query.locoNo);
    if (locoNo.length > 6) return res.status(400).json({ message: 'Loco_No cannot exceed 6 characters.' });
    if (!locoNo) return res.status(400).json({ message: 'Enter a Loco_No.' });
    const source = await SourceDataset.findOne({ key: 'KAVACH_LOCO_DETAILS' }).lean();
    if (!source) return res.status(409).json({ message: 'Upload Kavach_Loco_Details.xlsx first.' });
    const match = source.rows.find(x => norm(x.locoNo) === locoNo);
    if (!match) return res.status(404).json({ message: `Loco_No ${req.query.locoNo} was not found in Kavach_Loco_Details.xlsx.` });
    res.json({ found: true, locoNo: match.locoNo, shed: match.shed || '', locoType: match.locoType || '', rly: match.rly || '', kavachMake: match.kavachMake || '', brakeSystem: match.brakeSystem || '' });
  } catch (e) { res.status(500).json({ message: e.message }); }
});

function buildRow(x, searchDate, day, k) {
  return {
    sourceRow: x.sourceRow, sn: x.sourceRow, date: searchDate, day,
    trainNo: x.trainNo, dir: x.dir, trType: x.trType, time: x.time, attachedDiv: x.attachedDiv, fromTo: x.fromTo,
    locoLink: x.locoLink, locoLinkDiv: x.locoLinkDiv, dayWork: x.dayWork, locoNo: k ? k.locoNo : '', shed: k?.shed || '',
    locoType: k?.locoType || '', rly: k?.rly || '', kavachMake: k?.kavachMake || '', brakeSystem: k?.brakeSystem || '',
    kavachSection: x.kavachSection || '', worked: x.worked || '', noReason: x.noReason || '', remarks: x.remarks || '',
    shedRemark: '', oem: '', kavachFound: Boolean(k)
  };
}

app.get('/api/data/search', async (req, res) => {
  try {
    const { date } = req.query;
    if (!validDate(date || '')) return res.status(400).json({ message: 'Use DD-MM-YYYY.' });
    const trainSource = await SourceDataset.findOne({ key: 'TRAINS_PER_DAY' }).lean();
    const kavachSource = await SourceDataset.findOne({ key: 'KAVACH_LOCO_DETAILS' }).lean();
    if (!trainSource || !kavachSource) return res.status(409).json({ message: 'Upload both source files first.' });
    const kavachMap = new Map(kavachSource.rows.map(x => [norm(x.locoNo), x]));
    const day = weekday(date);
    const rows = trainSource.rows.filter(x => String(x.day).trim().toUpperCase() === day).sort((a, b) => tval(a.time) - tval(b.time));
    const sourceRows = rows.map(x => x.sourceRow);
    const saved = await SavedRow.find({ searchDate: date, sourceRow: { $in: sourceRows } }).select('sourceRow row').lean();
    const savedMap = new Map(saved.map(x => [String(x.sourceRow), x]));
    const enriched = rows.map((x, i) => {
      const savedRecord = savedMap.get(String(x.sourceRow));
      const savedRow = savedRecord?.row;
      const k = savedRow?.locoNo ? kavachMap.get(norm(savedRow.locoNo)) : null;
      return { ...buildRow(x, date, day, k), ...(savedRow || {}), sourceRow: x.sourceRow, sn: i + 1, date, day, added: Boolean(savedRecord), editing: false, kavachFound: Boolean(k), kavachError: '' };
    });
    res.json({ date, weekday: day, total: enriched.length, rows: enriched });
  } catch (e) { res.status(500).json({ message: e.message }); }
});

const EDITABLE_SAVED_FIELDS = ['locoNo', 'worked', 'noReason', 'remarks', 'shedRemark', 'oem'];

app.post('/api/data/rows', async (req, res) => {
  try {
    const { searchDate, weekday: day, row } = req.body || {};
    const locoNo = norm(row?.locoNo);
    if (locoNo.length > 6) return res.status(400).json({ message: 'Loco_No cannot exceed 6 characters.' });
    if (!searchDate || !validDate(searchDate) || !day || !row?.sourceRow || !locoNo) return res.status(400).json({ message: 'Valid search date and Loco_No are required.' });
    if (String(row?.remarks || '').length > 500) return res.status(400).json({ message: 'Remarks cannot exceed 500 characters.' });
    const source = await SourceDataset.findOne({ key: 'KAVACH_LOCO_DETAILS' }).lean();
    if (!source) return res.status(409).json({ message: 'Kavach source is not uploaded.' });
    const match = source.rows.find(x => norm(x.locoNo) === locoNo);
    if (!match) return res.status(400).json({ message: `Loco_No ${row.locoNo} was not found in Kavach_Loco_Details.xlsx.` });

    const existing = await SavedRow.findOne({ searchDate, sourceRow: row.sourceRow }).lean();
    let base = existing?.row ? { ...existing.row } : { ...row };
    if (existing?.row) {
      for (const field of EDITABLE_SAVED_FIELDS) base[field] = row[field] ?? '';
    }
    const savedRow = {
      ...base, sourceRow: row.sourceRow, locoNo: match.locoNo, shed: match.shed || '', locoType: match.locoType || '', rly: match.rly || '',
      kavachMake: match.kavachMake || '', brakeSystem: match.brakeSystem || '', kavachFound: true,
      remarks: String(base.remarks || '').slice(0, 500), shedRemark: String(base.shedRemark || '').slice(0, 300), oem: String(base.oem || '').slice(0, 200), editing: false
    };
    if (String(savedRow.worked || '').trim().toUpperCase() === 'YES') savedRow.noReason = '';

    const record = await SavedRow.findOneAndUpdate(
      { searchDate, sourceRow: row.sourceRow },
      { searchDate, weekday: day, sourceRow: row.sourceRow, row: savedRow, shedRemark: savedRow.shedRemark || '', oem: savedRow.oem || '', addedAt: new Date() },
      { upsert: true, new: true, setDefaultsOnInsert: true }
    );
    res.json({ message: existing ? 'Record updated successfully.' : 'Record saved successfully.', id: record._id, added: true, row: savedRow });
  } catch (e) {
    if (e.code === 11000) return res.status(409).json({ message: 'This row is already saved.' });
    res.status(500).json({ message: e.message });
  }
});

function parseDMY(s) {
  if (!validDate(s || '')) return null;
  const [d, m, y] = s.split('-').map(Number);
  return new Date(y, m - 1, d);
}
function formatDMY(d) { return `${String(d.getDate()).padStart(2, '0')}-${String(d.getMonth() + 1).padStart(2, '0')}-${d.getFullYear()}`; }
const SUMMARY_METRICS = new Set(['totalTrains','fitted','working','defective','kernex','doNotStart','isolation','other','total']);
const SUMMARY_METRIC_LABELS = {
  totalTrains: 'TOTAL TRAINS LOCO IN KAVACH SECTION', fitted: 'NO. OF LOCO KAVACH FITTED', working: 'NO. OF LOCO KAVACH WORKING',
  defective: 'KAVACH DEFECTIVE', kernex: 'KERNEX MAKE APPROVAL PENDING', doNotStart: 'DO NOT START REMARK/ STICKER',
  isolation: 'KAVACH ISOLATION PREVIOUS REMARKS', other: 'OTHER REASON', total: 'TOTAL'
};
function filterSaved(all, startDate, endDate, dir = '', metric = '') {
  const from = startDate ? parseDMY(startDate) : null, to = endDate ? parseDMY(endDate) : null;
  const rawDir = String(dir || '').trim().toUpperCase();
  const targetDir = rawDir === 'ALL' ? '' : rawDir;
  const targetMetric = String(metric || '').trim();
  const reasonMap = {
    defective: 'KAVACH DEFECTIVE', kernex: 'KERNEX APP. PENDING', doNotStart: 'KAVACH DO NOT START',
    isolation: 'KAVACH ISOLATION PREV.', other: 'OTHER'
  };
  const fittedReasons = new Set(Object.values(reasonMap));
  return all.filter(x => {
    const d = parseDMY(x.searchDate); const rowDir = String(x.row?.dir || '').trim().toUpperCase();
    if (!(d && (!from || d >= from) && (!to || d <= to) && (!targetDir || rowDir === targetDir))) return false;
    if (!targetMetric || targetMetric === 'totalTrains') return true;
    const worked = String(x.row?.worked || '').trim().toUpperCase() === 'YES';
    const reason = String(x.row?.noReason || '').trim().toUpperCase();
    if (targetMetric === 'working') return worked;
    if (targetMetric === 'fitted') return worked || fittedReasons.has(reason);
    if (targetMetric === 'total') return fittedReasons.has(reason);
    if (reasonMap[targetMetric]) return reason === reasonMap[targetMetric];
    return true;
  });
}

app.get('/api/data/saved', async (req, res) => {
  try {
    const { startDate, endDate, dir, metric } = req.query;
    if (startDate && !validDate(startDate)) return res.status(400).json({ message: 'Start date must be DD-MM-YYYY.' });
    if (endDate && !validDate(endDate)) return res.status(400).json({ message: 'End date must be DD-MM-YYYY.' });
    if (startDate && endDate && parseDMY(startDate) > parseDMY(endDate)) return res.status(400).json({ message: 'Start date cannot be after end date.' });
    if (dir && !['ALL', 'UP', 'DN'].includes(String(dir).toUpperCase())) return res.status(400).json({ message: 'DIR must be ALL, UP or DN.' });
    if (metric && !SUMMARY_METRICS.has(String(metric))) return res.status(400).json({ message: 'Unknown summary metric.' });
    const rows = filterSaved(await SavedRow.find({}).sort({ createdAt: -1 }).lean(), startDate, endDate, dir, metric);
    res.json({ total: rows.length, metric: metric || '', metricLabel: SUMMARY_METRIC_LABELS[metric] || '', rows: rows.map((x, i) => ({ ...x, row: { ...(x.row || {}), sn: i + 1 } })) });
  } catch (e) { res.status(500).json({ message: e.message }); }
});

function exportRows(rows) {
  return rows.map((x, i) => ({
    'S.N': i + 1, DATE: x.searchDate, DAY: x.weekday, KAVACH_WKG_SECTION: x.row?.kavachSection || '', DIR: x.row?.dir || '', TR_TYPE: x.row?.trType || '', T_O_Time: x.row?.time || '',
    'LOCO_ATTACHED_DIV.': x.row?.attachedDiv || '', From_To: x.row?.fromTo || '', LOCO_link: x.row?.locoLink || '', LOCO_LINK_DIV: x.row?.locoLinkDiv || '',
    Day_of_wkg_in_territory: x.row?.dayWork || '', Loco_No: x.row?.locoNo || '', SHED: x.row?.shed || '', Loco_Type: x.row?.locoType || '', RLY: x.row?.rly || '',
    Loco_Kavach_Make: x.row?.kavachMake || '', Brake_system: x.row?.brakeSystem || '', TRAIN_No: x.row?.trainNo || '',
    Train_worked_with_Kavach_YES_NO: x.row?.worked || '',
    'If_No_the_Reason \n( Non Kavach Loco\n/ Kavach defective\n/ KAVACH Fitness cirtification N.Avl.\n/ Crew incompetency)': x.row?.noReason || '',
    Remarks: x.row?.remarks || '', Shed_Remark: x.row?.shedRemark || '', OEM: x.row?.oem || ''
  }));
}

app.get('/api/data/export', async (req, res) => {
  try {
    const { startDate, endDate, dir } = req.query;
    if (startDate && !validDate(startDate)) return res.status(400).json({ message: 'Start date must be DD-MM-YYYY.' });
    if (endDate && !validDate(endDate)) return res.status(400).json({ message: 'End date must be DD-MM-YYYY.' });
    if (startDate && endDate && parseDMY(startDate) > parseDMY(endDate)) return res.status(400).json({ message: 'Start date cannot be after end date.' });
    if (dir && !['ALL', 'UP', 'DN'].includes(String(dir).toUpperCase())) return res.status(400).json({ message: 'DIR must be ALL, UP or DN.' });
    if (req.query.metric && !SUMMARY_METRICS.has(String(req.query.metric))) return res.status(400).json({ message: 'Unknown summary metric.' });
    const rows = filterSaved(await SavedRow.find({}).sort({ createdAt: 1 }).lean(), startDate, endDate, dir, req.query.metric || '');
    const headers = TEMPLATE_HEADERS;
    const ws = XLSX.utils.json_to_sheet(exportRows(rows), { header: headers });
    ws['!cols'] = headers.map((h, i) => ({ wch: i === 20 ? 44 : i === 21 ? 38 : i === 22 ? 28 : i === 23 ? 24 : Math.max(12, Math.min(28, String(h).length + 3)) }));
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Traction Operation');
    const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    const suffix = startDate || endDate ? `_${startDate || 'all'}_to_${endDate || 'all'}${dir ? '_' + String(dir).toUpperCase() : ''}` : '_all';
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="Traction_Operation_Data${suffix}.xlsx"`); res.send(buffer);
  } catch (e) { res.status(500).json({ message: e.message }); }
});

function summaryForRecords(records, date, dir) {
  const day = weekday(date).slice(0, 3);
  const working = records.filter(x => String(x.row?.worked || '').trim().toUpperCase() === 'YES').length;
  const reason = value => records.filter(x => String(x.row?.noReason || '').trim().toUpperCase() === value).length;
  const defective = reason('KAVACH DEFECTIVE');
  const kernex = reason('KERNEX APP. PENDING');
  const doNotStart = reason('KAVACH DO NOT START');
  const isolation = reason('KAVACH ISOLATION PREV.');
  const other = reason('OTHER');
  const total = defective + kernex + doNotStart + isolation + other;
  const fitted = working + total;
  return {
    day, date, dir, totalTrains: records.length, fitted, working, defective, kernex, doNotStart, isolation, other, total,
    percentWorking: fitted > 0 ? Number(((working / fitted) * 100).toFixed(2)) : 0
  };
}

// Adds the final TOTAL row used by the reference Summary workbook.
// All totals are calculated from the currently filtered date range + DIR only.
function summaryTotalRow(rows, dir) {
  const totalTrains = rows.reduce((sum, r) => sum + Number(r.totalTrains || 0), 0);
  const fitted = rows.reduce((sum, r) => sum + Number(r.fitted || 0), 0);
  const working = rows.reduce((sum, r) => sum + Number(r.working || 0), 0);
  const defective = rows.reduce((sum, r) => sum + Number(r.defective || 0), 0);
  const kernex = rows.reduce((sum, r) => sum + Number(r.kernex || 0), 0);
  const doNotStart = rows.reduce((sum, r) => sum + Number(r.doNotStart || 0), 0);
  const isolation = rows.reduce((sum, r) => sum + Number(r.isolation || 0), 0);
  const other = rows.reduce((sum, r) => sum + Number(r.other || 0), 0);
  const total = defective + kernex + doNotStart + isolation + other;
  return {
    day: 'TOTAL', date: '', dir, totalTrains, fitted, working, defective, kernex,
    doNotStart, isolation, other, total,
    percentWorking: fitted > 0 ? Number(((working / fitted) * 100).toFixed(2)) : 0,
    isTotal: true
  };
}

app.get('/api/summary', async (req, res) => {
  try {
    const { fromDate, toDate, dir } = req.query;
    if (!validDate(fromDate || '') || !validDate(toDate || '')) return res.status(400).json({ message: 'From Date and To Date must be DD-MM-YYYY.' });
    if (parseDMY(fromDate) > parseDMY(toDate)) return res.status(400).json({ message: 'From Date cannot be after To Date.' });
    const targetDir = String(dir || '').trim().toUpperCase();
    if (!['ALL', 'UP', 'DN'].includes(targetDir)) return res.status(400).json({ message: 'DIR must be ALL, UP or DN.' });

    const all = await SavedRow.find({}).sort({ searchDate: 1, sourceRow: 1 }).lean();
    const filtered = filterSaved(all, fromDate, toDate, targetDir);
    const byDate = new Map();
    filtered.forEach(x => { if (!byDate.has(x.searchDate)) byDate.set(x.searchDate, []); byDate.get(x.searchDate).push(x); });
    const rows = [];
    const cursor = parseDMY(fromDate), end = parseDMY(toDate);
    while (cursor <= end) {
      const d = formatDMY(cursor);
      rows.push(summaryForRecords(byDate.get(d) || [], d, targetDir));
      cursor.setDate(cursor.getDate() + 1);
    }
    const totalRow = summaryTotalRow(rows, targetDir);
    res.json({ fromDate, toDate, dir: targetDir, totalRecords: filtered.length, rows, totalRow });
  } catch (e) { res.status(500).json({ message: e.message }); }
});

app.get('/api/summary/export', async (req, res) => {
  try {
    const { fromDate, toDate, dir } = req.query;
    if (!validDate(fromDate || '') || !validDate(toDate || '')) return res.status(400).json({ message: 'From Date and To Date must be DD-MM-YYYY.' });
    if (parseDMY(fromDate) > parseDMY(toDate)) return res.status(400).json({ message: 'From Date cannot be after To Date.' });
    const targetDir = String(dir || '').trim().toUpperCase();
    if (!['ALL', 'UP', 'DN'].includes(targetDir)) return res.status(400).json({ message: 'DIR must be ALL, UP or DN.' });
    const all = await SavedRow.find({}).sort({ searchDate: 1, sourceRow: 1 }).lean();
    const filtered = filterSaved(all, fromDate, toDate, targetDir);
    const byDate = new Map(); filtered.forEach(x => { if (!byDate.has(x.searchDate)) byDate.set(x.searchDate, []); byDate.get(x.searchDate).push(x); });
    const rows = []; const cursor = parseDMY(fromDate), end = parseDMY(toDate);
    while (cursor <= end) { const d = formatDMY(cursor); rows.push(summaryForRecords(byDate.get(d) || [], d, targetDir)); cursor.setDate(cursor.getDate() + 1); }
    const totalRow = summaryTotalRow(rows, targetDir);
    const exportRows = [...rows, totalRow];
    const flat = exportRows.map(x => ({
      DAY: x.day, DATE: x.date, DIR: x.dir, 'TOTAL TRAINS LOCO IN KAVACH SECTION': x.totalTrains, 'NO. OF LOCO KAVACH FITTED': x.fitted,
      'NO. OF LOCO KAVACH WORKING': x.working, 'KAVACH DEFECTIVE': x.defective, 'KERNEX MAKE APPROVAL PENDING': x.kernex,
      'DO NOT START REMARK/ STICKER': x.doNotStart, 'KAVACH ISOLATION PREVIOUS REMARKS': x.isolation, 'OTHER REASON': x.other,
      TOTAL: x.total, '% WORKING': x.percentWorking
    }));
    const headers = ['DAY','DATE','DIR','TOTAL TRAINS LOCO IN KAVACH SECTION','NO. OF LOCO KAVACH FITTED','NO. OF LOCO KAVACH WORKING','KAVACH DEFECTIVE','KERNEX MAKE APPROVAL PENDING','DO NOT START REMARK/ STICKER','KAVACH ISOLATION PREVIOUS REMARKS','OTHER REASON','TOTAL','% WORKING'];
    const ws = XLSX.utils.json_to_sheet(flat, { header: headers });
    ws['!cols'] = headers.map(h => ({ wch: Math.max(12, Math.min(34, h.length + 3)) }));
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Summary');
    const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="Traction_Operation_Summary_${fromDate}_to_${toDate}_${targetDir}.xlsx"`); res.send(buffer);
  } catch (e) { res.status(500).json({ message: e.message }); }
});

mongoose.connect(process.env.MONGODB_URI)
  .then(() => app.listen(Number(process.env.PORT || 5000), () => console.log('API running on port ' + (process.env.PORT || 5000))))
  .catch(e => { console.error(e); process.exit(1); });
