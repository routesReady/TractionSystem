require('dotenv').config();
const dns = require('dns');
dns.setServers(['8.8.8.8', '8.8.4.4']);
const express = require('express');
const cors = require('cors');
const mongoose = require('mongoose');
const bcrypt = require('bcryptjs');
const crypto = require('crypto');
const multer = require('multer');
const jwt = require('jsonwebtoken');
const XLSX = require('xlsx');
const XLSXStyle = require('xlsx-js-style');
const { SourceDataset, SavedRow, User, PasswordResetRequest } = require('./models');
const { trains, kavach, norm, tval, validDate, weekday, TEMPLATE_HEADERS, trainValidationOptions } = require('./excel');

const app = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 15 * 1024 * 1024 } });

const allowedOrigins = String(process.env.CLIENT_ORIGIN || 'http://localhost:5173')
  .split(',').map(v => v.trim().replace(/\/+$/, '')).filter(Boolean);
const vercelOrigin = /^https:\/\/(?:traction-operation|traction-system)-[a-z0-9-]+\.vercel\.app$/i;

const corsOptions = {
  origin(origin, callback) {
    // Credentialed browser requests must receive an explicit allowed origin.
    // Never use '*' here because authentication uses an HttpOnly cross-origin cookie.
    if (!origin || allowedOrigins.includes(origin) || vercelOrigin.test(origin) || origin === 'https://traction-operation.vercel.app') {
      return callback(null, true);
    }
    return callback(new Error('CORS origin not allowed'));
  },
  credentials: true,
  methods: ['GET', 'HEAD', 'POST', 'PUT', 'PATCH', 'DELETE', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization', 'X-Requested-With'],
  optionsSuccessStatus: 204
};

app.use(cors(corsOptions));
app.use(express.json({ limit: '2mb' }));

app.use((req, res, next) => {
  // Do not cache authenticated/API responses. This also reduces the chance of
  // showing stale protected content after logout.
  if (req.path.startsWith('/api/')) res.setHeader('Cache-Control', 'no-store');
  next();
});

const AUTH_COOKIE = 'traction_auth';
const DEFAULT_USER_PASSWORD = '123';

function normalizeUsername(value) {
  return String(value || '').trim().replace(/\s+/g, ' ').toUpperCase();
}
function normalizeRole(value) {
  const r = String(value || '').trim().toLowerCase();
  return r === 'admin' || r === 'shed' || r === 'oem' ? r : '';
}
function cookieOptions() {
  const production = process.env.NODE_ENV === 'production';
  return {
    httpOnly: true,
    secure: production,
    sameSite: production ? 'none' : 'lax',
    maxAge: 8 * 60 * 60 * 1000,
    path: '/'
  };
}
function issueAuthCookie(res, user) {
  const token = jwt.sign(
    { sub: String(user._id), username: user.username, role: user.role, version: user.authVersion || 0 },
    process.env.JWT_SECRET,
    { expiresIn: '8h', issuer: 'traction-operation' }
  );
  res.cookie(AUTH_COOKIE, token, cookieOptions());
}
function clearAuthCookie(res) {
  res.clearCookie(AUTH_COOKIE, { httpOnly: true, secure: process.env.NODE_ENV === 'production', sameSite: process.env.NODE_ENV === 'production' ? 'none' : 'lax', path: '/' });
}
async function getAuthenticatedUser(req) {
  const cookies = Object.fromEntries(String(req.headers.cookie || '').split(';').map(v => v.trim()).filter(Boolean).map(v => { const i=v.indexOf('='); return i>0 ? [decodeURIComponent(v.slice(0,i)), decodeURIComponent(v.slice(i+1))] : [v,'']; }));
  const token = cookies[AUTH_COOKIE];
  if (!token) return null;
  const payload = jwt.verify(token, process.env.JWT_SECRET, { issuer: 'traction-operation' });
  const user = await User.findById(payload.sub);
  if (!user || !user.isActive || Number(user.authVersion || 0) !== Number(payload.version || 0)) return null;
  return user;
}
async function userAuth(req, res, next) {
  try {
    const user = await getAuthenticatedUser(req);
    if (!user) return res.status(401).json({ message: 'Authentication required.' });
    req.authUser = user;
    next();
  } catch {
    return res.status(401).json({ message: 'Authentication required.' });
  }
}
function requireRole(...roles) {
  return (req, res, next) => {
    if (!req.authUser || !roles.includes(req.authUser.role)) return res.status(403).json({ message: 'You are not authorized to access this resource.' });
    next();
  };
}
function publicUser(user) {
  return { id: String(user._id), username: user.username, role: user.role, mobileNumber: user.mobileNumber || '', mustChangePassword: Boolean(user.mustChangePassword) };
}

async function ensureAdminUser() {
  const username = normalizeUsername(process.env.ADMIN_USERNAME);
  const password = String(process.env.ADMIN_PASSWORD || '');
  if (!username || !password) {
    console.warn('ADMIN_USERNAME / ADMIN_PASSWORD are not configured. Admin login is disabled until they are set.');
    return;
  }
  const existing = await User.findOne({ username, role: 'admin' });
  if (existing) return;
  const passwordHash = await bcrypt.hash(password, 12);
  await User.create({ username, role: 'admin', passwordHash, isActive: true, authVersion: 0 });
}
async function syncExcelUsers() {
  // Shed names can exist in the daily train workbook even when a particular
  // shed currently has no matching row in the Kavach-loco workbook. OEM names
  // normally come from the Kavach workbook. Build both lists from every
  // available source so an Excel-created account can always be provisioned.
  const [trainSource, kavachSource] = await Promise.all([
    SourceDataset.findOne({ key: 'TRAINS_PER_DAY' }).lean(),
    SourceDataset.findOne({ key: 'KAVACH_LOCO_DETAILS' }).lean()
  ]);

  const sheds = new Set(), oems = new Set();
  for (const row of trainSource?.rows || []) {
    const shed = normalizeUsername(row.shed);
    if (shed) sheds.add(shed);
  }
  for (const row of kavachSource?.rows || []) {
    const shed = normalizeUsername(row.shed);
    const oem = normalizeUsername(row.kavachMake);
    if (shed) sheds.add(shed);
    if (oem) oems.add(oem);
  }

  const syncRole = async (names, role) => {
    if (!names.size) return;
    const hash = await bcrypt.hash(DEFAULT_USER_PASSWORD, 12);
    for (const username of names) {
      const existing = await User.findOne({ username, role });
      if (!existing) {
        await User.create({ username, role, passwordHash: hash, isActive: true, mustChangePassword: true, authVersion: 0 });
      } else if (existing.mustChangePassword) {
        // A temporary/reset account must continue to use the temporary
        // password until the user completes Change Password. Do not touch
        // accounts that have already selected their own password.
        await User.updateOne(
          { _id: existing._id },
          { $set: { passwordHash: hash, isActive: true } }
        );
      }
    }
  };
  await syncRole(sheds, 'shed');
  await syncRole(oems, 'oem');
}
async function sendResetSms(mobileNumber) {
  const sid = String(process.env.TWILIO_ACCOUNT_SID || '');
  const token = String(process.env.TWILIO_AUTH_TOKEN || '');
  const from = String(process.env.TWILIO_FROM_NUMBER || '');
  if (!sid || !token || !from) return { ok: false, status: 'not_configured', error: 'SMS provider is not configured.' };
  const body = 'Kavach Monitoring System, Kota Division (WCR) has reset your password. Your temporary password is 123.';
  try {
    const authHeader = Buffer.from(`${sid}:${token}`).toString('base64');
    const form = new URLSearchParams({ To: String(mobileNumber), From: from, Body: body });
    const response = await fetch(`https://api.twilio.com/2010-04-01/Accounts/${encodeURIComponent(sid)}/Messages.json`, {
      method: 'POST',
      headers: { Authorization: `Basic ${authHeader}`, 'Content-Type': 'application/x-www-form-urlencoded' },
      body: form.toString()
    });
    if (!response.ok) return { ok: false, status: 'failed', error: `SMS provider returned HTTP ${response.status}.` };
    return { ok: true, status: 'sent', error: '' };
  } catch (e) {
    return { ok: false, status: 'failed', error: e.message };
  }
}

app.post('/api/auth/login', async (req, res) => {
  try {
    const username = normalizeUsername(req.body?.username);
    const password = String(req.body?.password || '');
    if (!username || !password) return res.status(400).json({ message: 'Username and password are required.' });
    // Re-run the lightweight account bootstrap at login so newly uploaded
    // Excel users are immediately available without waiting for a server
    // restart. Admin is created from ADMIN_USERNAME/ADMIN_PASSWORD when set.
    await ensureAdminUser();
    await syncExcelUsers();
    const candidates = await User.find({ username, isActive: true }).limit(5);
    let matched = null;
    for (const candidate of candidates) {
      if (await bcrypt.compare(password, candidate.passwordHash)) { matched = candidate; break; }
    }
    if (!matched) return res.status(401).json({ message: 'Invalid username or password. Please try again.' });
    issueAuthCookie(res, matched);
    res.json({ user: publicUser(matched) });
  } catch (e) {
    res.status(500).json({ message: 'Unable to sign in right now.' });
  }
});

app.get('/api/auth/me', userAuth, (req, res) => res.json({ user: publicUser(req.authUser) }));

app.post('/api/auth/logout', userAuth, async (req, res) => {
  try {
    await User.updateOne({ _id: req.authUser._id }, { $inc: { authVersion: 1 } });
  } finally {
    clearAuthCookie(res);
  }
  res.json({ message: 'Logged out successfully.' });
});

app.post('/api/auth/change-password', userAuth, async (req, res) => {
  try {
    const oldPassword = String(req.body?.oldPassword || '');
    const newPassword = String(req.body?.newPassword || '');
    const confirmPassword = String(req.body?.confirmPassword || '');
    if (!oldPassword || !newPassword || newPassword !== confirmPassword) return res.status(400).json({ message: 'Please enter a valid old password and matching new password.' });
    if (newPassword.length < 6) return res.status(400).json({ message: 'New password must be at least 6 characters.' });
    if (!await bcrypt.compare(oldPassword, req.authUser.passwordHash)) return res.status(401).json({ message: 'Old password is incorrect.' });
    const hash = await bcrypt.hash(newPassword, 12);
    await User.updateOne({ _id: req.authUser._id }, { $set: { passwordHash: hash, mustChangePassword: false }, $inc: { authVersion: 1 } });
    clearAuthCookie(res);
    res.json({ message: 'Password changed successfully. Please log in again.' });
  } catch (e) { res.status(500).json({ message: e.message }); }
});

app.post('/api/auth/reset-request', async (req, res) => {
  try {
    const username = normalizeUsername(req.body?.username);
    const role = normalizeRole(req.body?.role);
    const mobileNumber = String(req.body?.mobileNumber || '').trim();
    if (!username || !['shed','oem'].includes(role) || !mobileNumber) return res.status(400).json({ message: 'Username, role and mobile number are required.' });
    await syncExcelUsers();
    const user = await User.findOne({ username, role, isActive: true });
    if (!user) return res.status(400).json({ message: 'Unable to create the reset request with the supplied details.' });
    const pending = await PasswordResetRequest.findOne({ username, role, status: 'Pending' });
    if (pending) return res.json({ message: 'Your password reset request has already been submitted to the Admin.' });
    const requestId = `PR-${Date.now().toString(36).toUpperCase()}-${crypto.randomBytes(3).toString('hex').toUpperCase()}`;
    await PasswordResetRequest.create({ requestId, username, role, mobileNumber, status: 'Pending' });
    await User.updateOne({ _id: user._id }, { $set: { mobileNumber } });
    res.status(201).json({ message: 'Your password reset request has been submitted to the Admin.' });
  } catch (e) { res.status(500).json({ message: e.message }); }
});

app.get('/api/auth/reset-requests', userAuth, requireRole('admin'), async (req, res) => {
  const rows = await PasswordResetRequest.find({}).sort({ requestedAt: -1 }).lean();
  res.json({ requests: rows });
});

app.post('/api/auth/reset-requests/:id/approve', userAuth, requireRole('admin'), async (req, res) => {
  try {
    const request = await PasswordResetRequest.findOne({ requestId: req.params.id, status: 'Pending' });
    if (!request) return res.status(404).json({ message: 'Pending reset request not found.' });
    const user = await User.findOne({ username: request.username, role: request.role, isActive: true });
    if (!user) return res.status(404).json({ message: 'The requested user no longer exists.' });
    const passwordHash = await bcrypt.hash(DEFAULT_USER_PASSWORD, 12);
    await User.updateOne({ _id: user._id }, { $set: { passwordHash, mobileNumber: request.mobileNumber, mustChangePassword: true }, $inc: { authVersion: 1 } });
    const sms = await sendResetSms(request.mobileNumber);
    await PasswordResetRequest.updateOne(
      { _id: request._id },
      { $set: { status: 'Approved', approvedAt: new Date(), approvedBy: req.authUser.username, smsStatus: sms.status, smsError: sms.error || '' } }
    );
    res.json({ message: sms.ok ? 'Password reset successfully. SMS notification sent.' : 'Password reset successfully, but the SMS notification could not be delivered.', smsStatus: sms.status });
  } catch (e) { res.status(500).json({ message: e.message }); }
});

app.post('/api/auth/reset-requests/:id/reject', userAuth, requireRole('admin'), async (req, res) => {
  try {
    const request = await PasswordResetRequest.findOneAndUpdate(
      { requestId: req.params.id, status: 'Pending' },
      { $set: { status: 'Rejected', approvedAt: new Date(), approvedBy: req.authUser.username } },
      { new: true }
    );
    if (!request) return res.status(404).json({ message: 'Pending reset request not found.' });
    res.json({ message: 'Password reset request rejected.' });
  } catch (e) { res.status(500).json({ message: e.message }); }
});

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

app.get('/api/upload/status', userAuth, requireRole('admin'), async (req, res) => {
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

app.post('/api/upload/authenticate', userAuth, requireRole('admin'), async (req, res) => {
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
    await syncExcelUsers();
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
      await syncExcelUsers();
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

app.get('/api/upload/download/:key', userAuth, requireRole('admin'), async (req, res) => {
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

app.get('/api/data/options', userAuth, requireRole('admin'), async (req, res) => {
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

app.get('/api/kavach/lookup', userAuth, requireRole('admin'), async (req, res) => {
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

app.get('/api/data/search', userAuth, requireRole('admin'), async (req, res) => {
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

app.post('/api/data/rows', userAuth, requireRole('admin'), async (req, res) => {
  try {
    const { searchDate, weekday: day, row } = req.body || {};
    const locoNo = norm(row?.locoNo);
    if (locoNo.length > 6) return res.status(400).json({ message: 'Loco_No cannot exceed 6 characters.' });
    if (!searchDate || !validDate(searchDate) || !day || !row?.sourceRow || !locoNo) return res.status(400).json({ message: 'Valid search date and Loco_No are required.' });
    const workedValue = String(row?.worked || '').trim().toUpperCase();
    const noReasonValue = String(row?.noReason || '').trim();
    if (!workedValue) return res.status(400).json({ message: 'Please select Train Worked with Kavach before updating the record.' });
    if (workedValue === 'NO' && !noReasonValue) return res.status(400).json({ message: 'Please select If No – Reason before updating the record.' });
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
      { searchDate, weekday: day, sourceRow: row.sourceRow, row: savedRow, shedRemark: savedRow.shedRemark || '', oem: savedRow.oem || '', shedRemarkClosed: Boolean(existing?.shedRemarkClosed), oemRemarkClosed: Boolean(existing?.oemRemarkClosed), addedAt: new Date() },
      { upsert: true, new: true, setDefaultsOnInsert: true }
    );
    res.json({ message: existing ? 'Record updated successfully.' : 'Record saved successfully.', id: record._id, added: true, row: savedRow });
  } catch (e) {
    if (e.code === 11000) return res.status(409).json({ message: 'This row is already saved.' });
    res.status(500).json({ message: e.message });
  }
});


app.get('/api/remark-records', userAuth, requireRole('shed','oem'), async (req, res) => {
  try {
    const user = req.authUser;
    const all = await SavedRow.find({}).sort({ searchDate: -1, sourceRow: 1 }).lean();
    const isShed = user.role === 'shed';
    const allowedReasons = isShed
      ? new Set(['KAVACH DEFECTIVE','KAVACH DO NOT START','KAVACH ISOLATION PREV.'])
      : new Set(['KAVACH DEFECTIVE','KAVACH DO NOT START','KAVACH ISOLATION PREV.','KERNEX APP PENDING','KERNEX APP. PENDING']);
    const rows = all.filter(x => {
      const r=x.row||{};
      const identity=isShed ? String(r.shed||'').trim().toUpperCase()===user.username : String(r.kavachMake||'').trim().toUpperCase()===user.username;
      const worked=String(r.worked||'').trim().toUpperCase();
      const reason=String(r.noReason||'').trim().toUpperCase();
      const closed = isShed ? Boolean(x.shedRemarkClosed) : Boolean(x.oemRemarkClosed);
      return identity && !closed && (worked==='NO' || worked==='PARTIALLY WORKING') && allowedReasons.has(reason);
    });
    res.json({ user: publicUser(user), rows });
  } catch(e){ res.status(500).json({message:e.message}); }
});

app.patch('/api/remark-records/:id', userAuth, requireRole('shed','oem'), async (req,res) => {
  try {
    const record = await SavedRow.findById(req.params.id);
    if (!record) return res.status(404).json({message:'Record not found.'});
    const row=record.row||{}, user=req.authUser;
    const isShed=user.role==='shed';
    const identity=isShed ? String(row.shed||'').trim().toUpperCase()===user.username : String(row.kavachMake||'').trim().toUpperCase()===user.username;
    if(!identity) return res.status(403).json({message:'You are not authorized to modify this record.'});
    const value=String(req.body?.remark??'');
    if(value.length>(isShed?300:200)) return res.status(400).json({message:`Remark cannot exceed ${isShed?300:200} characters.`});
    if(isShed){
      const updated = await SavedRow.findByIdAndUpdate(
        record._id,
        { $set: { 'row.shedRemark': value, shedRemark: value } },
        { new:true, runValidators:true }
      ).lean();
      return res.json({message:'Shed remark updated successfully.', row:updated});
    }
    const updated = await SavedRow.findByIdAndUpdate(
      record._id,
      { $set: { 'row.oem': value, oem: value } },
      { new:true, runValidators:true }
    ).lean();
    res.json({message:'OEM remark updated successfully.', row:updated});
  }catch(e){res.status(500).json({message:e.message});}
});

app.patch('/api/remark-records/:id/close', userAuth, requireRole('shed','oem'), async (req,res) => {
  try {
    const record = await SavedRow.findById(req.params.id).lean();
    if (!record) return res.status(404).json({message:'Record not found.'});
    const user=req.authUser, row=record.row||{}, isShed=user.role==='shed';
    const identity=isShed ? String(row.shed||'').trim().toUpperCase()===user.username : String(row.kavachMake||'').trim().toUpperCase()===user.username;
    if(!identity) return res.status(403).json({message:'You are not authorized to close this record.'});
    const hasRemark=isShed ? String(row.shedRemark||record.shedRemark||'').trim() : String(row.oem||record.oem||'').trim();
    if(!hasRemark) return res.status(400).json({message:`Save the ${isShed?'Shed':'OEM'} remark before closing this record.`});
    const field=isShed?'shedRemarkClosed':'oemRemarkClosed';
    const updated=await SavedRow.findByIdAndUpdate(record._id,{ $set:{ [field]:true } },{new:true,runValidators:true}).lean();
    res.json({message:'Record closed successfully.', row:updated});
  } catch(e){res.status(500).json({message:e.message});}
});

function parseDMY(s) {
  if (!validDate(s || '')) return null;
  const [d, m, y] = s.split('-').map(Number);
  return new Date(y, m - 1, d);
}
function formatDMY(d) { return `${String(d.getDate()).padStart(2, '0')}-${String(d.getMonth() + 1).padStart(2, '0')}-${d.getFullYear()}`; }
function savedReason(x) {
  return String(x?.row?.noReason || '').trim().toUpperCase();
}
function savedDir(x) {
  return String(x?.row?.dir || '').trim().toUpperCase();
}
function savedValue(x, keys) {
  for (const key of keys) {
    const value = x?.row?.[key];
    if (value !== undefined && value !== null && String(value).trim() !== '') return String(value).trim();
  }
  return '';
}
const NON_KAVACH_CATEGORIES = Object.freeze({
  LINK: 'MISLINK LOCO BY LOCO LINK DIV',
  ATTACH: 'MISLINK LOCO BY LOCO ATTACH DIV',
  PLAIN: 'NON KAVACH LOCO'
});
const DEFECTIVE_REASONS = new Set(['KAVACH DEFECTIVE','KERNEX APP. PENDING','KAVACH DO NOT START','KAVACH ISOLATION PREV.']);
const DETAIL_KAVACH_DEFECTIVE_REASONS = DEFECTIVE_REASONS;
const RUNNING_WITH_KAVACH_REASONS = new Set(['KAVACH DEFECTIVE','OTHER']);
const SUMMARY_METRICS = new Set(['totalTrains','fitted','runningWithKavach','working','defective','kernex','doNotStart','isolation','other','total','nonKavach','nonKavachLink','nonKavachAttach','nonKavachPlain','otherMake']);
const SUMMARY_METRIC_LABELS = {
  totalTrains: 'TOTAL TRAINS LOCO IN KAVACH SECTION', fitted: 'NO. OF LOCO KAVACH FITTED', runningWithKavach: 'TOTAL NO. OF TRAIN RUNNING WITH KAVACH', working: 'NO. OF LOCO KAVACH WORKING',
  defective: 'KAVACH DEFECTIVE', kernex: 'KERNEX MAKE APPROVAL PENDING', doNotStart: 'DO NOT START REMARK/ STICKER',
  isolation: 'KAVACH ISOLATION PREVIOUS REMARKS', other: 'OTHER REASON', total: 'TOTAL', nonKavach: 'NON KAVACH LOCO',
  nonKavachLink: 'NON KAVACH LOCO (MISLINK BY LOCO LINK DIVISION)', nonKavachAttach: 'NON KAVACH LOCO (MISLINK BY LOCO ATTACH DIVISION)',
  defectiveByShed: 'KAVACH DEFECTIVE LOCO',
  nonKavachPlain: 'NON KAVACH LOCO', otherMake: 'OTHER MAKE'
};
function canonicalMetric(metric) {
  const key=String(metric||'').trim();
  return ({targetTrains:'totalTrains',workingCondition:'working',otherReason:'other'})[key] || key;
}
function metricRecordMatches(x, metric, groupValue='') {
  metric=canonicalMetric(metric);
  const reason = savedReason(x);
  const worked = String(x?.row?.worked || '').trim().toUpperCase() === 'YES';
  const otherMake = /OTHER MAKE/.test(reason) || /OTHER MAKE/.test(String(x?.row?.kavachMake || '').toUpperCase());
  const group = String(groupValue || '').trim();
  switch (String(metric || '').trim()) {
    case '':
    case 'targetTrains':
    case 'totalTrains':
      return true;
    case 'runningWithKavach':
    case 'fitted':
      return worked || reason === 'KAVACH DEFECTIVE' || reason === 'OTHER' || otherMake;
    case 'working':
      return worked;
    case 'defective':
    case 'defectiveByShed':
      return DETAIL_KAVACH_DEFECTIVE_REASONS.has(reason) &&
        (metric !== 'defectiveByShed' || String(x?.row?.shed || '').trim() === group);
    case 'kernex':
      return reason === 'KERNEX APP. PENDING';
    case 'doNotStart':
      return reason === 'KAVACH DO NOT START';
    case 'isolation':
      return reason === 'KAVACH ISOLATION PREV.';
    case 'other':
    case 'otherReason':
      return reason === 'OTHER';
    case 'otherMake':
      return otherMake;
    case 'nonKavach':
      return Object.values(NON_KAVACH_CATEGORIES).includes(reason);
    case 'nonKavachLink':
      return reason === NON_KAVACH_CATEGORIES.LINK &&
        (!group || savedValue(x, ['locoLinkDiv','LOCO_LINK_DIV','LOCO_LINK_DIVISION','Loco Link Division']) === group);
    case 'nonKavachAttach':
      return reason === NON_KAVACH_CATEGORIES.ATTACH &&
        (!group || savedValue(x, ['attachedDiv','LOCO_ATTACHED_DIV.','LOCO_ATTACHED_DIV','LOCO_ATTACHED_DIVISION','Loco Attached Division']) === group);
    case 'nonKavachPlain':
      return reason === NON_KAVACH_CATEGORIES.PLAIN &&
        (!group || savedValue(x, ['attachedDiv','LOCO_ATTACHED_DIV.','LOCO_ATTACHED_DIV','LOCO_ATTACHED_DIVISION','Loco Attached Division']) === group);
    case 'total':
      return DEFECTIVE_REASONS.has(reason) || reason === 'OTHER';
    default:
      return false;
  }
}

// Single source of truth for every Summary / Summary & Detail clickable metric.
// Both displayed counts and popup/export records are derived from this exact
// predicate after applying the same date + DIR base filter.
function metricRecords(records, metric, groupValue='') {
  return records.filter(x => metricRecordMatches(x, metric, groupValue));
}

function filterSaved(all, startDate, endDate, dir = '', metric = '', groupValue='') {
  const from = startDate ? parseDMY(startDate) : null, to = endDate ? parseDMY(endDate) : null;
  const rawDir = String(dir || '').trim().toUpperCase();
  const targetDir = rawDir === 'ALL' ? '' : rawDir;
  const base = all.filter(x => {
    const d = parseDMY(x.searchDate), rowDir = savedDir(x);
    return Boolean(d && (!from || d >= from) && (!to || d <= to) && (!targetDir || rowDir === targetDir));
  });
  return metric ? metricRecords(base, canonicalMetric(metric), groupValue) : base;
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

function styleGeneratedWorksheet(ws, headers, rowCount) {
  const endCol = XLSXStyle.utils.encode_col(Math.max(0, headers.length - 1));
  const endRow = Math.max(1, rowCount + 1);
  ws['!autofilter'] = { ref: `A1:${endCol}${endRow}` };
  ws['!freeze'] = { xSplit: 0, ySplit: 1, topLeftCell: 'A2', activePane: 'bottomRight', state: 'frozen' };
  ws['!rows'] = [{ hpt: 28 }];
  for (let c = 0; c < headers.length; c++) {
    const addr = XLSXStyle.utils.encode_cell({ r: 0, c });
    if (!ws[addr]) ws[addr] = { t: 's', v: headers[c] };
    ws[addr].s = {
      font: { bold: true, color: { rgb: 'FFFFFF' }, sz: 10 },
      fill: { patternType: 'solid', fgColor: { rgb: '173B6B' } },
      alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
      border: { top:{style:'thin',color:{rgb:'9AA9B8'}}, bottom:{style:'thin',color:{rgb:'9AA9B8'}}, left:{style:'thin',color:{rgb:'9AA9B8'}}, right:{style:'thin',color:{rgb:'9AA9B8'}} }
    };
  }
  for (let r = 1; r < endRow; r++) {
    ws['!rows'][r] = { hpt: 20 };
    for (let c = 0; c < headers.length; c++) {
      const addr = XLSXStyle.utils.encode_cell({ r, c });
      if (!ws[addr]) continue;
      ws[addr].s = {
        font: { color: { rgb: '1F2937' }, sz: 10 },
        alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
        border: { top:{style:'thin',color:{rgb:'D4DDE5'}}, bottom:{style:'thin',color:{rgb:'D4DDE5'}}, left:{style:'thin',color:{rgb:'D4DDE5'}}, right:{style:'thin',color:{rgb:'D4DDE5'}} }
      };
    }
  }
}
function makeStyledWorkbook(sheetName, headers, rows, widths) {
  const ws = XLSXStyle.utils.json_to_sheet(rows, { header: headers });
  ws['!cols'] = headers.map((h, i) => ({ wch: widths?.[i] || Math.max(12, Math.min(36, String(h).length + 3)) }));
  styleGeneratedWorksheet(ws, headers, rows.length);
  const wb = XLSXStyle.utils.book_new(); XLSXStyle.utils.book_append_sheet(wb, ws, sheetName);
  return XLSXStyle.write(wb, { type: 'buffer', bookType: 'xlsx', cellStyles: true });
}

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
    const buffer = makeStyledWorkbook('Traction Operation', headers, exportRows(rows), headers.map((h, i) => i === 20 ? 44 : i === 21 ? 38 : i === 22 ? 28 : i === 23 ? 24 : Math.max(12, Math.min(28, String(h).length + 3))));
    const suffix = startDate || endDate ? `_${startDate || 'all'}_to_${endDate || 'all'}${dir ? '_' + String(dir).toUpperCase() : ''}` : '_all';
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="Traction_Operation_Data${suffix}.xlsx"`); res.send(buffer);
  } catch (e) { res.status(500).json({ message: e.message }); }
});

function summaryForRecords(records, date, dir) {
  const day = weekday(date).slice(0, 3);
  const count = metric => metricRecords(records, metric).length;
  const working = count('working');
  const defective = count('defective');
  const kernex = count('kernex');
  const doNotStart = count('doNotStart');
  const isolation = count('isolation');
  const other = count('other');
  const otherMake = count('otherMake');
  const total = count('total');
  // NO. OF LOCO KAVACH FITTED / RUNNING WITH KAVACH uses the exact same
  // record predicate as Summary & Detail and its popup.
  const fitted = count('fitted');
  const runningWithKavach = fitted;
  return {
    day, date, dir, totalTrains: records.length, fitted, runningWithKavach, working,
    defective, kernex, doNotStart, isolation, other, total, otherMake,
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


function detailSummaryForRecords(records) {
  const up = records.filter(x => savedDir(x) === 'UP');
  const dn = records.filter(x => savedDir(x) === 'DN');
  const section = arr => {
    const count = metric => metricRecords(arr, metric).length;
    return {
      targetTrains: arr.length,
      runningWithKavach: count('runningWithKavach'),
      workingCondition: count('working'),
      nonKavach: count('nonKavach'),
      defective: count('defective'),
      otherMake: count('otherMake'),
      otherReason: count('other')
    };
  };
  const defectiveRows = metricRecords(records, 'defective').map((x,i)=>({
    sn:i+1,date:x.searchDate||'',trainNo:x.row?.trainNo||'',locoNo:x.row?.locoNo||'',shed:x.row?.shed||'',kavachMake:x.row?.kavachMake||'',
    tlcRemark:x.row?.remarks||'',shedRemark:x.row?.shedRemark||'',oemRemark:x.row?.oem||''
  }));
  const nonKavachPlainRows = metricRecords(records,'nonKavachPlain').map((x,i)=>({
    sn:i+1,date:x.searchDate||'',trainNo:x.row?.trainNo||'',locoNo:x.row?.locoNo||'',shed:x.row?.shed||'',
    locoAttachDiv:savedValue(x,['attachedDiv','LOCO_ATTACHED_DIV.','LOCO_ATTACHED_DIV','LOCO_ATTACHED_DIVISION','Loco Attached Division']),
    tlcRemark:x.row?.remarks||''
  }));
  const otherRows = metricRecords(records, 'other').map((x,i)=>({
    sn:i+1,date:x.searchDate||'',trainNo:x.row?.trainNo||'',locoNo:x.row?.locoNo||'',shed:x.row?.shed||'',kavachMake:x.row?.kavachMake||'',
    tlcRemark:x.row?.remarks||'',shedRemark:x.row?.shedRemark||'',oemRemark:x.row?.oem||''
  }));
  const shedMap = new Map();
  metricRecords(records, 'defective').forEach(x=>{
    const name=String(x.row?.shed||'—').trim()||'—';
    if(!shedMap.has(name))shedMap.set(name,{name,total:0,up:0,dn:0});
    const item=shedMap.get(name); item.total++; if(savedDir(x)==='UP')item.up++; if(savedDir(x)==='DN')item.dn++;
  });
  return {
    total:section(records),up:section(up),dn:section(dn),defectiveRows,
    defectiveByShed:[...shedMap.values()].sort((a,b)=>b.total-a.total || a.name.localeCompare(b.name)),
    nonKavachPlainRows,otherRows
  };
}

function detailMetricRecords(records, metric, dir='', groupValue='') {
  let rows = records;
  const targetDir = String(dir||'').trim().toUpperCase();
  if (targetDir && targetDir !== 'ALL') rows = rows.filter(x => savedDir(x) === targetDir);
  return metricRecords(rows, metric, groupValue);
}



function pdfEscape(value) {
  return String(value ?? '').replace(/\\/g,'\\\\').replace(/\(/g,'\\(').replace(/\)/g,'\\)').replace(/\r?\n/g,' ');
}
function pdfAscii(value) {
  return String(value ?? '').normalize('NFKD').replace(/[^\x20-\x7E]/g,' ');
}
function wrapPdfText(value, maxChars) {
  const text = pdfAscii(value).trim();
  if (!text) return [''];
  const words = text.split(/\s+/); const lines=[]; let line='';
  for (const word of words) {
    if (word.length > maxChars) {
      if (line) { lines.push(line); line=''; }
      for (let i=0;i<word.length;i+=maxChars) lines.push(word.slice(i,i+maxChars));
      continue;
    }
    const next=line?`${line} ${word}`:word;
    if (next.length<=maxChars) line=next; else { lines.push(line); line=word; }
  }
  if (line) lines.push(line); return lines.length?lines:[''];
}
function buildPdfDocument(pages) {
  const objects=[]; const add=o=>{objects.push(o); return objects.length;};
  const fontRegular=add('<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>');
  const fontBold=add('<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >>');
  const pageIds=[]; const contentIds=[];
  // A4 portrait report. Dynamic wrapping and page breaks keep all content visible.
  const W=595.28,H=841.89;
  for(const page of pages){
    const stream=page.join('\n');
    const contentId=add(`<< /Length ${Buffer.byteLength(stream,'binary')} >>\nstream\n${stream}\nendstream`); contentIds.push(contentId);
    const pageId=add(`<< /Type /Page /Parent PAGES /MediaBox [0 0 ${W} ${H}] /Resources << /Font << /F1 ${fontRegular} 0 R /F2 ${fontBold} 0 R >> >> /Contents ${contentId} 0 R >>`); pageIds.push(pageId);
  }
  const kids=pageIds.map(id=>`${id} 0 R`).join(' '); const pagesId=add(`<< /Type /Pages /Kids [${kids}] /Count ${pageIds.length} >>`);
  for(const id of pageIds) objects[id-1]=objects[id-1].replace('PAGES',`${pagesId} 0 R`);
  const catalogId=add(`<< /Type /Catalog /Pages ${pagesId} 0 R >>`);
  let pdf='%PDF-1.4\n'; const offsets=[0];
  for(let i=0;i<objects.length;i++){ offsets.push(Buffer.byteLength(pdf,'binary')); pdf+=`${i+1} 0 obj\n${objects[i]}\nendobj\n`; }
  const xref=Buffer.byteLength(pdf,'binary'); pdf+=`xref\n0 ${objects.length+1}\n0000000000 65535 f \n`; for(let i=1;i<offsets.length;i++) pdf+=`${String(offsets[i]).padStart(10,'0')} 00000 n \n`;
  pdf+=`trailer\n<< /Size ${objects.length+1} /Root ${catalogId} 0 R >>\nstartxref\n${xref}\n%%EOF`;
  return Buffer.from(pdf,'binary');
}

function createSummaryDetailPdf(report){
  const W=595.28,H=841.89,M=22,contentW=W-M*2,pages=[],ops=[];
  const headerH=54,footerH=22,bottomLimit=M+footerH+6;
  let y=H-M-headerH,pageNo=1;
  const text=(x,yy,str,size=7,bold=false)=>{const font=bold?'/F2':'/F1';ops.push(`BT ${font} ${size} Tf ${x.toFixed(2)} ${yy.toFixed(2)} Td (${pdfEscape(pdfAscii(str))}) Tj ET`);};
  const line=(x1,y1,x2,y2)=>ops.push(`${x1.toFixed(2)} ${y1.toFixed(2)} m ${x2.toFixed(2)} ${y2.toFixed(2)} l S`);
  const rectFill=(x,yy,w,h,gray)=>{ops.push(`${gray} g`);ops.push(`${x.toFixed(2)} ${yy.toFixed(2)} ${w.toFixed(2)} ${h.toFixed(2)} re f`);ops.push('0 g');};
  const rectStroke=(x,yy,w,h)=>ops.push(`${x.toFixed(2)} ${yy.toFixed(2)} ${w.toFixed(2)} ${h.toFixed(2)} re S`);
  const drawPageHeader=()=>{rectFill(M,H-M-38,contentW,38,0.93);rectStroke(M,H-M-38,contentW,38);text(M+8,H-M-14,'TRACTION OPERATION - KAVACH MONITORING SYSTEM',10,true);text(M+8,H-M-27,'West Central Railway, Kota Division',7.2,false);const dateLabel=`Summary & Detail | ${report.fromDate} to ${report.toDate}`;text(W-M-8-(dateLabel.length*3.0),H-M-14,dateLabel,6.7,true);line(M,H-M-42,W-M,H-M-42);};
  const drawPageFooter=()=>{line(M,M+14,W-M,M+14);text(M,M+5,'Traction Operation - Kavach Monitoring System | Kota Division',6,false);text(W-M-42,M+5,`Page ${pageNo}`,6,false);};
  const finishPage=()=>{drawPageFooter();pages.push(ops.splice(0,ops.length));};
  const startPage=()=>{drawPageHeader();y=H-M-headerH;};
  const newPage=()=>{finishPage();pageNo++;startPage();};
  startPage();
  const wrapFor=(value,width,fs,pad=3)=>wrapPdfText(value,Math.max(2,Math.floor((width-pad*2)/(fs*0.82))));
  const rowHeight=(lines,lineH,pad)=>Math.max(1,lines)*lineH+pad*2;
  const available=(lineH,pad)=>Math.max(1,Math.floor((y-bottomLimit-pad*2)/lineH));
  const table=(headers,rows,widths,opts={})=>{
    const fs=opts.fontSize||7,pad=opts.pad||3,lineH=opts.lineH||8.8,totalW=widths.reduce((a,b)=>a+b,0);let rowIndex=0;
    const drawChunk=(wraps,start,count,isHeader=false)=>{const rh=rowHeight(count,lineH,pad);if(y-rh<bottomLimit)return false;const bg=isHeader?0.88:(rowIndex%2===0?0.97:1);if(bg<1)rectFill(M,y-rh,totalW,rh,bg);line(M,y,M+totalW,y);let x=M;for(let i=0;i<widths.length;i++){line(x,y,x,y-rh);line(x+widths[i],y,x+widths[i],y-rh);const vals=wraps[i].slice(start,start+count);for(let j=0;j<vals.length;j++){const value=vals[j];const left=opts.leftBodyColumns?.includes(i)&&!isHeader;const centered=!left&&Boolean(opts.alignCenter);const approx=value.length*fs*0.23;const tx=centered?x+(widths[i]-approx)/2:x+pad;text(Math.max(x+1,Math.min(tx,x+widths[i]-pad-1)),y-pad-fs-j*lineH,value,fs,isHeader);}x+=widths[i];}line(M,y-rh,M+totalW,y-rh);y-=rh;return true;};
    const drawHeader=()=>{const wraps=headers.map((c,i)=>wrapFor(c,widths[i],fs,pad));const lines=Math.max(1,...wraps.map(a=>a.length));const rh=rowHeight(lines,lineH,pad);if(y-rh<bottomLimit)newPage();drawChunk(wraps,0,lines,true);};
    drawHeader();
    for(const row of rows){
      const wraps=row.map((c,i)=>wrapFor(c,widths[i],fs,pad));
      const totalLines=Math.max(1,...wraps.map(a=>a.length));
      const fullHeight=rowHeight(totalLines,lineH,pad);
      const remaining=y-bottomLimit;
      const pageBodyHeight=(H-M-headerH)-bottomLimit;
      if(fullHeight>remaining && fullHeight<=pageBodyHeight && y!==H-M-headerH){newPage();drawHeader();}
      let start=0;
      while(start<totalLines){
        const avail=available(lineH,pad);
        const count=Math.min(avail,totalLines-start);
        if(!drawChunk(wraps,start,count,false)){newPage();drawHeader();continue;}
        start+=count;
        if(start<totalLines){newPage();drawHeader();}
      }
      rowIndex++;
    }
    y-=7;
  };
  const section=(title,headers,rows,widths,opts={})=>{const fs=opts.fontSize||7,pad=opts.pad||3,lineH=opts.lineH||8.8;const headingH=24;const headerLines=Math.max(1,...headers.map((c,i)=>wrapFor(c,widths[i],fs,pad).length));const first=rows.length?rows[0]:headers.map(()=> '');const firstLines=Math.max(1,...first.map((c,i)=>wrapFor(c,widths[i],fs,pad).length));const needed=headingH+rowHeight(headerLines,lineH,pad)+Math.min(firstLines,3)*lineH+8;if(y-needed<bottomLimit)newPage();rectFill(M,y-headingH,contentW,headingH,0.90);rectStroke(M,y-headingH,contentW,headingH);text(M+8,y-15,title,8.2,true);y-=headingH+4;table(headers,rows,widths,opts);};

  text(M,y,'Summary & Detail',13,true);y-=16;text(M,y,`Date From: ${report.fromDate}    To: ${report.toDate}`,7.5,false);y-=11;
  const mrows=[['1','No of Target Trains',report.total.targetTrains,report.up.targetTrains,report.dn.targetTrains],['2','TOTAL NO. OF TRAIN RUNNING WITH KAVACH',report.total.runningWithKavach,report.up.runningWithKavach,report.dn.runningWithKavach],['3','TOTAL NO. OF TRAIN RUNNING WITH KAVACH IN WORKING CONDITION',report.total.workingCondition,report.up.workingCondition,report.dn.workingCondition],['4','NON KAVACH LOCO',report.total.nonKavach,report.up.nonKavach,report.dn.nonKavach],['5','KAVACH DEFECTIVE LOCO',report.total.defective,report.up.defective,report.dn.defective],['6','OTHER MAKE',report.total.otherMake,report.up.otherMake,report.dn.otherMake],['7','OTHER REASON',report.total.otherReason,report.up.otherReason,report.dn.otherReason]];
  table(['','DETAILS','TOTAL','UP','DN'],mrows,[24,345,68,68,68],{alignCenter:true,leftBodyColumns:[1],fontSize:7.1,lineH:8.8,pad:3});
  const divisionRows=(rows)=>rows.length?rows.map((r,i)=>[i+1,r.name,r.total,r.up,r.dn]):[[1,'No data',0,0,0]];
  section('KAVACH DEFECTIVE LOCO',['','SHED','TOTAL','UP','DN'],divisionRows(report.defectiveByShed),[24,345,68,68,68],{alignCenter:true,leftBodyColumns:[1],fontSize:7,lineH:8.6,pad:3});
  const nonRows=report.nonKavachPlainRows?.length?report.nonKavachPlainRows.map(r=>[r.sn,r.date,r.trainNo,r.locoNo,r.shed,r.locoAttachDiv,r.tlcRemark]):[[1,'No data','','','','','']];
  section('NON KAVACH LOCO ATTACH BY ATTACH DIVISON',['S.No.','Date','Train No','Loco No','SHED','Loco Attach Div','TLC Remark'],nonRows,[25,65,56,54,45,70,236],{alignCenter:true,leftBodyColumns:[1,2,3,4,5,6],fontSize:6.6,lineH:8.1,pad:3});
  const detailTable=(title,rows)=>{const heads=['S.N','Date','Train No','Loco No','SHED','KAVACH MAKE','TLC Remark','Shed Remark','OEM Remark'];const rr=rows.length?rows.map(r=>[r.sn,r.date,r.trainNo,r.locoNo,r.shed,r.kavachMake,r.tlcRemark,r.shedRemark,r.oemRemark]):[[1,'No data','','','','','','','']];section(title,heads,rr,[18,65,40,38,32,42,105,105,106],{alignCenter:true,leftBodyColumns:[1,2,3,4,5,6,7,8],fontSize:5.1,lineH:6.7,pad:2.5});};
  detailTable('KAVACH DEFECTIVE LOCO DETAILS',report.defectiveRows);detailTable('OTHER REASON',report.otherRows);
  if(ops.length)finishPage();return buildPdfDocument(pages);
}

async function getSummaryMetricRecords({fromDate,toDate,dir='ALL',metric='totalTrains',groupValue=''}) {
  if(!validDate(fromDate||'') || !validDate(toDate||'')) {
    const err = new Error('From Date and To Date must be DD-MM-YYYY.'); err.statusCode = 400; throw err;
  }
  if(parseDMY(fromDate)>parseDMY(toDate)) {
    const err = new Error('From Date cannot be after To Date.'); err.statusCode = 400; throw err;
  }
  const targetDir=String(dir||'ALL').trim().toUpperCase();
  if(!['ALL','UP','DN'].includes(targetDir)) {
    const err = new Error('DIR must be ALL, UP or DN.'); err.statusCode = 400; throw err;
  }
  const metricName=canonicalMetric(metric||'totalTrains');
  if(!SUMMARY_METRICS.has(metricName) && !['defectiveByShed'].includes(metricName)) {
    const err = new Error('Unknown summary metric.'); err.statusCode = 400; throw err;
  }
  const all=await SavedRow.find({}).sort({searchDate:1,sourceRow:1}).lean();
  // IMPORTANT: start with the exact same date + DIR dataset used by Summary,
  // then apply the exact same metric predicate. This is the single source
  // for both the displayed count and the popup/export records.
  const base=filterSaved(all,fromDate,toDate,targetDir);
  const rows=detailMetricRecords(base,metricName,targetDir,groupValue||'');
  return {fromDate,toDate,dir:targetDir,metric:metricName,metricLabel:SUMMARY_METRIC_LABELS[metricName]||metricName,total:rows.length,rows};
}

// Common metric-record endpoint used by Summary popups and Summary & Detail.
// In particular, clicking the WORKING count uses the exact `working` predicate
// that generated the displayed count, rather than fetching all records for the date.
app.get('/api/summary/metric-records', async (req,res)=>{
  try {
    const payload=await getSummaryMetricRecords(req.query);
    res.json(payload);
  } catch(e){res.status(e.statusCode||500).json({message:e.message});}
});

// Backward-compatible alias for existing Summary & Detail clients.
app.get('/api/summary-detail/records', async (req,res)=>{
  try {
    const payload=await getSummaryMetricRecords(req.query);
    res.json(payload);
  } catch(e){res.status(e.statusCode||500).json({message:e.message});}
});

app.get('/api/summary/metric-export', async (req,res)=>{
  try {
    const {fromDate,toDate,dir,groupValue} = req.query;
    const metric=canonicalMetric(req.query.metric||'');
    if(!validDate(fromDate||'') || !validDate(toDate||'')) return res.status(400).json({message:'From Date and To Date must be DD-MM-YYYY.'});
    if(parseDMY(fromDate)>parseDMY(toDate)) return res.status(400).json({message:'From Date cannot be after To Date.'});
    if(!SUMMARY_METRICS.has(metric) && !['defectiveByShed'].includes(metric)) return res.status(400).json({message:'Unknown summary metric.'});
    const targetDir=String(dir||'ALL').trim().toUpperCase();
    if(!['ALL','UP','DN'].includes(targetDir)) return res.status(400).json({message:'DIR must be ALL, UP or DN.'});
    const all=await SavedRow.find({}).sort({searchDate:1,sourceRow:1}).lean();
    const base=filterSaved(all,fromDate,toDate,targetDir);
    const rows=metricRecords(base,String(metric),String(groupValue||''));
    const headers=TEMPLATE_HEADERS;
    const buffer=makeStyledWorkbook('Summary Records',headers,exportRows(rows),headers.map((h,i)=>i===20?44:i===21?38:i===22?28:i===23?24:Math.max(12,Math.min(28,String(h).length+3))));
    const label=String(SUMMARY_METRIC_LABELS[metric]||metric).replace(/[^A-Za-z0-9]+/g,'_').replace(/^_|_$/g,'');
    const suffix=`_${fromDate}_to_${toDate}_${targetDir}_${label||'records'}`;
    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition',`attachment; filename="Summary_${suffix}.xlsx"`);
    res.send(buffer);
  } catch(e){res.status(500).json({message:e.message});}
});

app.get('/api/summary-detail/pdf', async (req,res)=>{
  try {
    const {fromDate,toDate}=req.query;
    if(!validDate(fromDate||'') || !validDate(toDate||'')) return res.status(400).json({message:'From Date and To Date must be DD-MM-YYYY.'});
    if(parseDMY(fromDate)>parseDMY(toDate)) return res.status(400).json({message:'From Date cannot be after To Date.'});
    const all=await SavedRow.find({}).sort({searchDate:1,sourceRow:1}).lean();
    const records=filterSaved(all,fromDate,toDate,'ALL');
    const report={fromDate,toDate,...detailSummaryForRecords(records)};
    const buffer=createSummaryDetailPdf(report);
    res.setHeader('Content-Type','application/pdf');
    res.setHeader('Content-Disposition',`attachment; filename="Summary_Detail_${fromDate}_to_${toDate}.pdf"`);
    res.setHeader('Content-Length',buffer.length); res.send(buffer);
  } catch(e){res.status(500).json({message:e.message});}
});

app.get('/api/summary-detail', async (req,res)=>{
  try {
    const {fromDate,toDate}=req.query;
    if(!validDate(fromDate||'') || !validDate(toDate||'')) return res.status(400).json({message:'From Date and To Date must be DD-MM-YYYY.'});
    if(parseDMY(fromDate)>parseDMY(toDate)) return res.status(400).json({message:'From Date cannot be after To Date.'});
    const all=await SavedRow.find({}).sort({searchDate:1,sourceRow:1}).lean();
    const records=filterSaved(all,fromDate,toDate,'ALL');
    res.json({fromDate,toDate,totalRecords:records.length,...detailSummaryForRecords(records)});
  } catch(e){res.status(500).json({message:e.message});}
});

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
    const buffer = makeStyledWorkbook('Summary', headers, flat, headers.map(h => Math.max(14, Math.min(36, h.length + 3))));
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="Traction_Operation_Summary_${fromDate}_to_${toDate}_${targetDir}.xlsx"`); res.send(buffer);
  } catch (e) { res.status(500).json({ message: e.message }); }
});

mongoose.connect(process.env.MONGODB_URI)
  .then(async () => { await ensureAdminUser(); await syncExcelUsers(); return app.listen(Number(process.env.PORT || 5000), () => console.log('API running on port ' + (process.env.PORT || 5000))); })
  .catch(e => { console.error(e); process.exit(1); });
