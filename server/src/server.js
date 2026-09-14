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
  .split(',')
  .map(v => v.trim().replace(/\/+$/, ''))
  .filter(Boolean);

app.use(cors({
  origin(origin, callback) {
    if (!origin || allowedOrigins.includes(origin)) return callback(null, true);
    if (/^https:\/\/traction-system-[a-z0-9-]+\.vercel\.app$/i.test(origin)) return callback(null, true);
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
  } catch {
    res.status(401).json({ message: 'Authentication required.' });
  }
};

const isExcel = name => /\.(xlsx|xls|xlsm)$/i.test(name || '');
const UPLOAD_SECRET = process.env.UPLOAD_PASSWORD || '';
const authenticatePassword = password => Boolean(UPLOAD_SECRET && password && password === UPLOAD_SECRET);

const saveSource = async (key, fileName, rows, buffer, contentType) => SourceDataset.findOneAndUpdate(
  { key },
  {
    key,
    fileName,
    rows,
    fileData: buffer,
    contentType: contentType || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    uploadedAt: new Date()
  },
  { upsert: true, new: true }
);

// Mongoose .lean() may expose a MongoDB Binary rather than a Node Buffer.
// Always normalize it before sending a stored workbook to the browser.
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
      train: train && { fileName: train.fileName, uploadedAt: train.uploadedAt, count: train.rows.length, hasFile: Boolean(train.fileData) },
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
    if (!trainFile || !kavachFile || !isExcel(trainFile.originalname) || !isExcel(kavachFile.originalname)) {
      return res.status(400).json({ message: 'Both Excel files are required.' });
    }
    const existingTrain = await SourceDataset.exists({ key: 'TRAINS_PER_DAY' });
    const existingKavach = await SourceDataset.exists({ key: 'KAVACH_LOCO_DETAILS' });
    if (existingTrain && existingKavach) return res.status(409).json({ message: 'Both source files already exist. Use Update Source Files.' });

    const trainRows = trains(trainFile.buffer);
    const kavachRows = kavach(kavachFile.buffer);
    await saveSource('TRAINS_PER_DAY', trainFile.originalname, trainRows, trainFile.buffer, trainFile.mimetype);
    await saveSource('KAVACH_LOCO_DETAILS', kavachFile.originalname, kavachRows, kavachFile.buffer, kavachFile.mimetype);
    res.status(201).json({ message: 'Source files uploaded successfully.', trainCount: trainRows.length, kavachCount: kavachRows.length });
  } catch (e) { res.status(500).json({ message: e.message }); }
});

app.post('/api/upload/update', auth, upload.fields([{ name: 'trainsFile', maxCount: 1 }, { name: 'kavachFile', maxCount: 1 }]), async (req, res) => {
  try {
    const trainFile = req.files?.trainsFile?.[0];
    const kavachFile = req.files?.kavachFile?.[0];
    if (!trainFile && !kavachFile) return res.status(400).json({ message: 'Select at least one updated file.' });
    if (trainFile) {
      if (!isExcel(trainFile.originalname)) return res.status(400).json({ message: 'Invalid train Excel file.' });
      await saveSource('TRAINS_PER_DAY', trainFile.originalname, trains(trainFile.buffer), trainFile.buffer, trainFile.mimetype);
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
      Remarks: r.remarks
    }))
    : source.rows.map(r => ({ Loco_No: r.locoNo, SHED: r.shed, Loco_Type: r.locoType, RLY: r.rly, Loco_Kavach_Make: r.kavachMake, Brake_system: r.brakeSystem }));
  const headers = source.key === 'TRAINS_PER_DAY' ? TEMPLATE_HEADERS : ['Loco_No', 'SHED', 'Loco_Type', 'RLY', 'Loco_Kavach_Make', 'Brake_system'];
  const ws = XLSX.utils.json_to_sheet(data, { header: headers });
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

    // Validate the stored bytes. If an older database record contains an invalid
    // BSON/Buffer representation, generate a clean XLSX from the parsed rows.
    if (!data) { data = generatedSourceWorkbook(source); generated = true; }
    else {
      try { XLSX.read(data, { type: 'buffer' }); }
      catch { data = generatedSourceWorkbook(source); generated = true; }
    }
    if (generated && !/\.xlsx$/i.test(filename)) filename = filename.replace(/\.(xlsx|xls|xlsm)$/i, '') + '.xlsx';

    const extension = /\.xlsm$/i.test(filename) ? 'application/vnd.ms-excel.sheet.macroEnabled.12' :
      /\.xls$/i.test(filename) ? 'application/vnd.ms-excel' :
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    const safeName = filename.replace(/[^a-zA-Z0-9._ -]/g, '_');
    res.setHeader('Content-Type', extension);
    res.setHeader('Content-Disposition', `attachment; filename="${safeName}"`);
    res.send(data);
  } catch (e) { res.status(500).json({ message: e.message }); }
});

app.get('/api/data/options', async (req, res) => {
  try {
    const source = await SourceDataset.findOne({ key: 'TRAINS_PER_DAY' }).lean();
    if (!source) return res.status(409).json({ message: 'Upload TRAINS_PER_DAY.xlsx first.' });

    // Always derive these lists from the latest uploaded workbook. This means
    // changing the source Excel values and re-uploading the file automatically
    // changes the dropdowns without any code/configuration change.
    let rows = source.rows || [];
    const sourceBuffer = toBuffer(source.fileData);
    let workbookOptions = { worked: [], noReason: [] };
    if (sourceBuffer) {
      try {
        const parsed = trains(sourceBuffer);
        if (parsed.length) rows = parsed;
      } catch {
        // Keep the already-parsed database rows as a safe fallback.
      }
      // IMPORTANT: Excel dropdowns are Data Validation lists. In the supplied
      // TRAINS_PER_DAY workbook, columns S and T contain no row values; the
      // actual choices live in the worksheet's <dataValidation> XML. Read those
      // lists directly so the portal mirrors Excel exactly after every upload.
      workbookOptions = trainValidationOptions(sourceBuffer);
    }

    const unique = field => Array.from(new Set(
      rows.map(r => String(r?.[field] ?? '').trim()).filter(Boolean)
    ));

    res.json({
      workedOptions: workbookOptions.worked.length ? workbookOptions.worked : unique('worked'),
      reasonOptions: workbookOptions.noReason.length ? workbookOptions.noReason : unique('noReason')
    });
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
    trainNo: x.trainNo, dir: x.dir, trType: x.trType, time: x.time,
    attachedDiv: x.attachedDiv, fromTo: x.fromTo, locoLink: x.locoLink, locoLinkDiv: x.locoLinkDiv,
    dayWork: x.dayWork, locoNo: k ? k.locoNo : '', shed: k?.shed || '', locoType: k?.locoType || '', rly: k?.rly || '',
    kavachMake: k?.kavachMake || '', brakeSystem: k?.brakeSystem || '', kavachSection: x.kavachSection || '',
    worked: x.worked || '', noReason: x.noReason || '', remarks: x.remarks || '', kavachFound: Boolean(k)
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
    const rows = trainSource.rows
      .filter(x => String(x.day).trim().toUpperCase() === day)
      .sort((a, b) => tval(a.time) - tval(b.time));
    const sourceRows = rows.map(x => x.sourceRow);
    const saved = await SavedRow.find({ searchDate: date, sourceRow: { $in: sourceRows } }).select('sourceRow row').lean();
    const savedMap = new Map(saved.map(x => [String(x.sourceRow), x]));

    const enriched = rows.map(x => {
      const savedRecord = savedMap.get(String(x.sourceRow));
      const savedRow = savedRecord?.row;
      const k = savedRow?.locoNo ? kavachMap.get(norm(savedRow.locoNo)) : null;
      return {
        ...buildRow(x, date, day, k),
        ...(savedRow || {}),
        sourceRow: x.sourceRow,
        sn: x.sourceRow,
        date,
        day,
        added: Boolean(savedRecord),
        editing: false,
        kavachFound: Boolean(k),
        kavachError: ''
      };
    });

    res.json({ date, weekday: day, total: enriched.length, rows: enriched });
  } catch (e) { res.status(500).json({ message: e.message }); }
});

app.post('/api/data/rows', async (req, res) => {
  try {
    const { searchDate, weekday: day, row } = req.body || {};
    const sourceRowNumber = Number(row?.sourceRow);
    const locoNo = norm(row?.locoNo);
    const worked = String(row?.worked ?? '').trim();
    const noReason = String(row?.noReason ?? '').trim();
    const remarks = String(row?.remarks ?? '').slice(0, 500);

    if (locoNo.length > 6) return res.status(400).json({ message: 'Loco_No cannot exceed 6 characters.' });
    if (!searchDate || !validDate(searchDate) || !day || !Number.isInteger(sourceRowNumber) || sourceRowNumber < 2 || !locoNo) {
      return res.status(400).json({ message: 'Valid search date, source row and Loco_No are required.' });
    }

    const [trainSource, kavachSource] = await Promise.all([
      SourceDataset.findOne({ key: 'TRAINS_PER_DAY' }).lean(),
      SourceDataset.findOne({ key: 'KAVACH_LOCO_DETAILS' }).lean()
    ]);
    if (!trainSource || !kavachSource) return res.status(409).json({ message: 'Both source files must be uploaded.' });

    const trainSourceRow = trainSource.rows.find(x => Number(x.sourceRow) === sourceRowNumber);
    if (!trainSourceRow) return res.status(404).json({ message: `Source train row ${sourceRowNumber} was not found in TRAINS_PER_DAY.xlsx.` });

    const match = kavachSource.rows.find(x => norm(x.locoNo) === locoNo);
    if (!match) return res.status(400).json({ message: `Loco_No ${row.locoNo} was not found in Kavach_Loco_Details.xlsx.` });

    // The source workbook owns all roster columns. The user can only change
    // Loco_No, the two dropdown fields, and Remarks. Kavach-derived fields are
    // always refreshed from the currently uploaded Kavach workbook.
    const sourceBase = buildRow(trainSourceRow, searchDate, day, match);
    const immutableSourceFields = {
      sourceRow: sourceRowNumber,
      sn: sourceRowNumber,
      date: searchDate,
      day,
      trainNo: sourceBase.trainNo,
      dir: sourceBase.dir,
      trType: sourceBase.trType,
      time: sourceBase.time,
      attachedDiv: sourceBase.attachedDiv,
      fromTo: sourceBase.fromTo,
      locoLink: sourceBase.locoLink,
      locoLinkDiv: sourceBase.locoLinkDiv,
      dayWork: sourceBase.dayWork,
      kavachSection: sourceBase.kavachSection
    };

    const savedRow = {
      ...immutableSourceFields,
      locoNo: match.locoNo,
      shed: match.shed || '',
      locoType: match.locoType || '',
      rly: match.rly || '',
      kavachMake: match.kavachMake || '',
      brakeSystem: match.brakeSystem || '',
      worked,
      noReason,
      remarks,
      kavachFound: true,
      editing: false
    };

    const record = await SavedRow.findOneAndUpdate(
      { searchDate, sourceRow: sourceRowNumber },
      { $set: { searchDate, weekday: day, sourceRow: sourceRowNumber, row: savedRow, addedAt: new Date() } },
      { upsert: true, new: true, setDefaultsOnInsert: true }
    );

    res.json({ message: 'Record saved successfully.', id: record._id, added: true, row: savedRow });
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
function filterSaved(all, startDate, endDate) {
  const from = startDate ? parseDMY(startDate) : null;
  const to = endDate ? parseDMY(endDate) : null;
  return all.filter(x => {
    const d = parseDMY(x.searchDate);
    return (!from || d >= from) && (!to || d <= to);
  });
}

app.get('/api/data/saved', async (req, res) => {
  try {
    const { startDate, endDate } = req.query;
    if (startDate && !validDate(startDate)) return res.status(400).json({ message: 'Start date must be DD-MM-YYYY.' });
    if (endDate && !validDate(endDate)) return res.status(400).json({ message: 'End date must be DD-MM-YYYY.' });
    if (startDate && endDate && parseDMY(startDate) > parseDMY(endDate)) return res.status(400).json({ message: 'Start date cannot be after end date.' });
    const rows = filterSaved(await SavedRow.find({}).sort({ createdAt: -1 }).lean(), startDate, endDate);
    res.json({ total: rows.length, rows });
  } catch (e) { res.status(500).json({ message: e.message }); }
});

app.get('/api/data/export', async (req, res) => {
  try {
    const { startDate, endDate } = req.query;
    if (startDate && !validDate(startDate)) return res.status(400).json({ message: 'Start date must be DD-MM-YYYY.' });
    if (endDate && !validDate(endDate)) return res.status(400).json({ message: 'End date must be DD-MM-YYYY.' });
    if (startDate && endDate && parseDMY(startDate) > parseDMY(endDate)) return res.status(400).json({ message: 'Start date cannot be after end date.' });

    const rows = filterSaved(await SavedRow.find({}).sort({ createdAt: 1 }).lean(), startDate, endDate);
    const flat = rows.map((x, i) => ({
      'S.N': i + 1, DATE: x.searchDate, DAY: x.weekday, TRAIN_No: x.row?.trainNo || '', DIR: x.row?.dir || '', TR_TYPE: x.row?.trType || '', T_O_Time: x.row?.time || '',
      'LOCO_ATTACHED_DIV.': x.row?.attachedDiv || '', From_To: x.row?.fromTo || '', LOCO_link: x.row?.locoLink || '', LOCO_LINK_DIV: x.row?.locoLinkDiv || '',
      Day_of_wkg_in_territory: x.row?.dayWork || '', Loco_No: x.row?.locoNo || '', SHED: x.row?.shed || '', Loco_Type: x.row?.locoType || '', RLY: x.row?.rly || '',
      Loco_Kavach_Make: x.row?.kavachMake || '', Brake_system: x.row?.brakeSystem || '', KAVACH_WKG_SECTION: x.row?.kavachSection || '',
      Train_worked_with_Kavach_YES_NO: x.row?.worked || '',
      'If_No_the_Reason \n( Non Kavach Loco\n/ Kavach defective\n/ KAVACH Fitness cirtification N.Avl.\n/ Crew incompetency)': x.row?.noReason || '',
      Remarks: x.row?.remarks || ''
    }));
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(flat, { header: TEMPLATE_HEADERS });
    ws['!cols'] = TEMPLATE_HEADERS.map((h, i) => ({ wch: i === 20 ? 46 : i === 21 ? 42 : Math.max(14, Math.min(28, String(h).length + 3)) }));
    XLSX.utils.book_append_sheet(wb, ws, 'Traction Operation');
    const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    const suffix = startDate || endDate ? `_${startDate || 'all'}_to_${endDate || 'all'}` : '_all';
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="Traction_Operation_Data${suffix}.xlsx"`);
    res.send(buffer);
  } catch (e) { res.status(500).json({ message: e.message }); }
});

mongoose.connect(process.env.MONGODB_URI)
  .then(() => app.listen(Number(process.env.PORT || 5000), () => console.log('API running on port ' + (process.env.PORT || 5000))))
  .catch(e => { console.error(e); process.exit(1); });
