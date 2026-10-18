require('dotenv').config();
const dns = require('dns');
dns.setServers(['8.8.8.8', '8.8.4.4']);
const express = require('express');
const cors = require('cors');
const mongoose = require('mongoose');
const bcrypt = require('bcryptjs');
const crypto = require('crypto');
const nodemailer = require('nodemailer');
const multer = require('multer');
const jwt = require('jsonwebtoken');
const XLSX = require('xlsx');
const XLSXStyle = require('xlsx-js-style');
const { SourceDataset, SavedRow, User, PasswordResetRequest, ActiveSession } = require('./models');
const { trains, kavach, norm, tval, validDate, weekday, TEMPLATE_HEADERS, trainValidationOptions, sourceColumns } = require('./excel');

const app = express();
// API responses (including protected record data) must never be restored from browser/proxy caches after logout.
app.use('/api', (req, res, next) => { res.set('Cache-Control', 'no-store, no-cache, must-revalidate, private'); res.set('Pragma', 'no-cache'); res.set('Expires', '0'); next(); });
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
async function issueAuthCookie(res, user) {
  const sessionId = crypto.randomUUID();
  const now = new Date();
  const inactivityMs = 5 * 60 * 1000;
  const expiresAt = user.role === 'admin' ? new Date(now.getTime() + 8 * 60 * 60 * 1000) : new Date(now.getTime() + inactivityMs);
  await ActiveSession.create({ sessionId, userId: user._id, username: user.username, role: user.role, loginAt: now, lastActivityAt: now, expiresAt, active: true });
  const token = jwt.sign(
    { sub: String(user._id), username: user.username, role: user.role, version: user.authVersion || 0, sid: sessionId, loginAt: now.toISOString() },
    process.env.JWT_SECRET,
    { expiresIn: '8h', issuer: 'traction-operation' }
  );
  res.cookie(AUTH_COOKIE, token, cookieOptions());
  return { sessionId, loginAt: now.toISOString(), inactivityTimeoutMs: user.role === 'admin' ? null : inactivityMs };
}
function clearAuthCookie(res) {
  res.clearCookie(AUTH_COOKIE, { httpOnly: true, secure: process.env.NODE_ENV === 'production', sameSite: process.env.NODE_ENV === 'production' ? 'none' : 'lax', path: '/' });
}
async function readAuthPayload(req) {
  const cookies = Object.fromEntries(String(req.headers.cookie || '').split(';').map(v => v.trim()).filter(Boolean).map(v => { const i=v.indexOf('='); return i>0 ? [decodeURIComponent(v.slice(0,i)), decodeURIComponent(v.slice(i+1))] : [v,'']; }));
  const token = cookies[AUTH_COOKIE];
  if (!token) return null;
  const payload = jwt.verify(token, process.env.JWT_SECRET, { issuer: 'traction-operation' });
  if (!payload.sid) return null;
  return payload;
}
async function getAuthenticatedUser(req, { touch = true } = {}) {
  const payload = await readAuthPayload(req);
  if (!payload) return null;
  const [user, session] = await Promise.all([User.findById(payload.sub), ActiveSession.findOne({ sessionId: payload.sid, active: true })]);
  if (!user || !user.isActive || Number(user.authVersion || 0) !== Number(payload.version || 0) || !session || String(session.userId) !== String(user._id)) return null;
  const now = Date.now();
  if (user.role !== 'admin' && now - new Date(session.lastActivityAt).getTime() >= 5 * 60 * 1000) {
    await ActiveSession.updateOne({ _id: session._id }, { $set: { active: false } });
    return null;
  }
  req.authSession = session;
  if (touch && user.role !== 'admin' && !['/api/auth/me','/api/auth/session-status','/api/admin/sessions'].includes(req.path)) {
    const lastActivityAt = new Date();
    await ActiveSession.updateOne({ _id: session._id, active: true }, { $set: { lastActivityAt, expiresAt: new Date(now + 5 * 60 * 1000) } });
    session.lastActivityAt = lastActivityAt;
  }
  return user;
}
async function userAuth(req, res, next) {
  try {
    const user = await getAuthenticatedUser(req);
    if (!user) {
      clearAuthCookie(res);
      return res.status(401).json({ message: 'Session expired or authentication required.', code: 'SESSION_EXPIRED' });
    }
    req.authUser = user;
    next();
  } catch {
    clearAuthCookie(res);
    return res.status(401).json({ message: 'Session expired or authentication required.', code: 'SESSION_EXPIRED' });
  }
}
function requireRole(...roles) {
  return (req, res, next) => {
    if (!req.authUser || !roles.includes(req.authUser.role)) return res.status(403).json({ message: 'You are not authorized to access this resource.' });
    next();
  };
}
function publicUser(user) {
  return { id: String(user._id), username: user.username, role: user.role, shedIdentifier: user.shedIdentifier || '', oemIdentifier: user.oemIdentifier || '', mobileNumber: user.mobileNumber || '', mustChangePassword: Boolean(user.mustChangePassword) };
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
  // Build authorized Shed and OEM accounts from the currently uploaded source
  // datasets. OEM identities are explicitly bound to a valid SHED + make pair.
  const [trainSource, kavachSource, masterSource] = await Promise.all([
    SourceDataset.findOne({ key: 'TRAINS_PER_DAY' }).lean(),
    SourceDataset.findOne({ key: 'KAVACH_LOCO_DETAILS' }).lean(),
    SourceDataset.findOne({ key: MASTER_DATA_KEY }).lean()
  ]);
  const sheds = new Set(), oemPairs = new Map();
  const masterSheds = new Set((masterSource?.rows || []).map(row => normalizeUsername(row.SHED || row.shed)).filter(Boolean));
  for (const row of trainSource?.rows || []) {
    const shed = normalizeUsername(row.shed);
    if (shed) sheds.add(shed);
  }
  for (const row of kavachSource?.rows || []) {
    const shed = normalizeUsername(row.shed);
    const oem = normalizeUsername(row.kavachMake);
    if (shed) sheds.add(shed);
    if (shed && oem && (!masterSheds.size || masterSheds.has(shed))) oemPairs.set(`${shed}-${oem}`, { shed, oem });
  }

  const hash = await bcrypt.hash(DEFAULT_USER_PASSWORD, 12);
  for (const username of sheds) {
    const existing = await User.findOne({ username, role: 'shed' });
    if (!existing) {
      await User.create({ username, role: 'shed', passwordHash: hash, isActive: true, mustChangePassword: true, authVersion: 0 });
    } else if (existing.mustChangePassword) {
      await User.updateOne({ _id: existing._id }, { $set: { passwordHash: hash, isActive: true } });
    }
  }

  const validPairUsernames = new Set(oemPairs.keys());
  for (const [username, pair] of oemPairs) {
    const existing = await User.findOne({ username, role: 'oem' });
    if (!existing) {
      await User.create({ username, role: 'oem', shedIdentifier: pair.shed, oemIdentifier: pair.oem, passwordHash: hash, isActive: true, mustChangePassword: true, authVersion: 0 });
    } else {
      const changes = { shedIdentifier: pair.shed, oemIdentifier: pair.oem, isActive: true };
      if (existing.mustChangePassword) changes.passwordHash = hash;
      await User.updateOne({ _id: existing._id }, { $set: changes });
    }
  }
  // Retire legacy make-only OEM logins so OEM access always uses SHED-OEM.
  await User.updateMany({ role: 'oem', username: { $nin: [...validPairUsernames] }, isActive: true }, { $set: { isActive: false }, $inc: { authVersion: 1 } });
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
    const session = await issueAuthCookie(res, matched);
    res.json({ user: publicUser(matched), session });
  } catch (e) {
    res.status(500).json({ message: 'Unable to sign in right now.' });
  }
});

app.get('/api/auth/me', userAuth, (req, res) => res.json({ user: publicUser(req.authUser), session: { loginAt: req.authSession?.loginAt, lastActivityAt: req.authSession?.lastActivityAt, inactivityTimeoutMs: req.authUser.role === 'admin' ? null : 300000 } }));

app.get('/api/auth/session-status', userAuth, (req,res)=>res.json({active:true,role:req.authUser.role,loginAt:req.authSession?.loginAt,lastActivityAt:req.authSession?.lastActivityAt,expiresAt:req.authSession?.expiresAt,inactivityTimeoutMs:req.authUser.role==='admin'?null:300000}));
app.post('/api/auth/activity', userAuth, async (req,res)=>{
  if(req.authUser.role==='admin') return res.json({active:true,role:'admin',inactivityTimeoutMs:null});
  const now=new Date(); await ActiveSession.updateOne({sessionId:req.authSession.sessionId,active:true},{$set:{lastActivityAt:now,expiresAt:new Date(now.getTime()+300000)}});
  res.json({active:true,lastActivityAt:now.toISOString(),expiresAt:new Date(now.getTime()+300000).toISOString(),inactivityTimeoutMs:300000});
});

app.post('/api/auth/logout', userAuth, async (req, res) => {
  try {
    if (req.authSession?.sessionId) await ActiveSession.updateOne({ sessionId: req.authSession.sessionId }, { $set: { active: false } });
  } finally {
    clearAuthCookie(res);
  }
  res.json({ message: 'Logged out successfully.' });
});

app.get('/api/admin/sessions', userAuth, requireRole('admin'), async (req,res)=>{
  try {
    const cutoff=new Date(Date.now()-5*60*1000);
    await ActiveSession.updateMany({active:true,role:{$in:['shed','oem']},lastActivityAt:{$lte:cutoff}},{$set:{active:false}});
    const sessions=await ActiveSession.find({active:true,role:{$in:['shed','oem']}}).sort({lastActivityAt:-1}).lean();
    res.json({sessions:sessions.map(s=>({id:s.sessionId,username:s.username,role:s.role,loginAt:s.loginAt,lastActivityAt:s.lastActivityAt,expiresAt:new Date(new Date(s.lastActivityAt).getTime()+300000)}))});
  } catch(e){res.status(500).json({message:'Unable to load active sessions.'});}
});
app.post('/api/admin/sessions/:id/logout', userAuth, requireRole('admin'), async (req,res)=>{
  try { const result=await ActiveSession.updateOne({sessionId:String(req.params.id),role:{$in:['shed','oem']},active:true},{$set:{active:false}}); if(!result.modifiedCount) return res.status(404).json({message:'This session is no longer active.'}); res.json({message:'User session terminated successfully.'}); }
  catch(e){res.status(500).json({message:'Unable to terminate this session.'});}
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
    await ActiveSession.updateMany({ userId: req.authUser._id, active: true }, { $set: { active: false } });
    clearAuthCookie(res);
    res.json({ message: 'Password changed successfully. Please log in again.' });
  } catch (e) { res.status(500).json({ message: e.message }); }
});

app.post('/api/auth/reset-request', async (req, res) => {
  try {
    const username = normalizeUsername(req.body?.username);
    const role = normalizeRole(req.body?.role);
    const mobileNumber = String(req.body?.mobileNumber || '').trim();
    const name=String(req.body?.name||'').trim(), designation=String(req.body?.designation||'').trim(), hqStation=String(req.body?.hqStation||'').trim(), division=String(req.body?.division||'').trim(), emailAddress=String(req.body?.emailAddress||'').trim();
    const userType=String(req.body?.userType||'').trim();
    if (!username || !['shed','oem'].includes(role) || !/^\d{10}$/.test(mobileNumber) || !name || !designation || !hqStation || !division || !['SHED User','OEM User'].includes(userType) || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(emailAddress)) return res.status(400).json({ message: 'Complete all required fields. Mobile must be 10 digits and email must be valid.' });
    if ((userType==='SHED User'?'shed':'oem')!==role) return res.status(400).json({message:'User Type must match the selected account role.'});
    await syncExcelUsers();
    const user = await User.findOne({ username, role, isActive: true });
    if (!user) return res.status(400).json({ message: 'Unable to create the reset request with the supplied details.' });
    const pending = await PasswordResetRequest.findOne({ username, role, status: 'Pending' });
    if (pending) return res.json({ message: 'Your password reset request has already been submitted to the Admin.' });
    const requestId = `PR-${Date.now().toString(36).toUpperCase()}-${crypto.randomBytes(3).toString('hex').toUpperCase()}`;
    await PasswordResetRequest.create({ requestId, username, role, mobileNumber, name, designation, userType, hqStation, division, emailAddress, status: 'Pending' });
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
    await ActiveSession.updateMany({ userId: user._id, active: true }, { $set: { active: false } });
    const sms = await sendResetSms(request.mobileNumber);
    let emailStatus='not_configured', emailError='';
    try {
      if (process.env.SMTP_HOST && process.env.SMTP_USER && process.env.SMTP_PASS && request.emailAddress) {
        const transporter=nodemailer.createTransport({host:process.env.SMTP_HOST,port:Number(process.env.SMTP_PORT||587),secure:String(process.env.SMTP_SECURE||'false')==='true',auth:{user:process.env.SMTP_USER,pass:process.env.SMTP_PASS}});
        await transporter.sendMail({from:`Traction Operation <${process.env.SMTP_FROM||'mujahatalikhan@gmail.com'}>`,to:request.emailAddress,subject:'Your Password Has Been Reset Successfully',text:`Hi ${request.name||request.username},\n\nYour password has been reset successfully.\n\nYour temporary password is: 123\n\nPlease log in using your account credentials and change this temporary password after logging in.\n\nRegards,\nTraction Operation – Kavach Monitoring System`});
        emailStatus='sent';
      } else emailError='SMTP settings or request email address are missing.';
    } catch(mailErr) { emailStatus='failed'; emailError=String(mailErr.message||mailErr).slice(0,500); }
    await PasswordResetRequest.updateOne({ _id: request._id }, { $set: { status:'Approved', approvedAt:new Date(), approvedBy:req.authUser.username, smsStatus:sms.status, smsError:sms.error||'', emailStatus, emailError, ...(emailStatus==='sent'?{emailSentAt:new Date()}: {}) } });
    res.json({ message: emailStatus==='sent' ? 'Password reset successfully and confirmation email sent.' : `Password reset successfully, but confirmation email was not sent (${emailStatus}). ${emailError}`, smsStatus:sms.status, emailStatus });
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
const MASTER_DATA_KEY = 'MASTER_DATA_SHED_RLY_DIVISION';
function parseMasterData(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer' });
  const sheetName = wb.SheetNames[0];
  if (!sheetName) throw new Error('Master Data workbook has no worksheet.');
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: '', raw: false });
  const required = ['SHED','RLY','SHED_NAME','DIVISION'];
  const headers = Object.keys(rows[0] || {}).map(v => String(v).trim().toUpperCase());
  if (!required.every(h => headers.includes(h))) {
    throw new Error('Master Data file must contain SHED, RLY, SHED_NAME and DIVISION columns.');
  }
  const normalized = rows.map(r => {
    const by = {};
    Object.entries(r).forEach(([k,v]) => { by[String(k).trim().toUpperCase()] = String(v ?? '').trim(); });
    return { SHED: by.SHED || '', RLY: by.RLY || '', SHED_NAME: by.SHED_NAME || '', DIVISION: by.DIVISION || '' };
  }).filter(r => r.SHED);
  if (!normalized.length) throw new Error('Master Data file contains no SHED mappings.');
  const seen = new Set(), deduped = [];
  normalized.forEach(r => {
    const key = r.SHED.toUpperCase();
    if (!seen.has(key)) { seen.add(key); deduped.push(r); }
  });
  return deduped;
}
const UPLOAD_SECRET = process.env.UPLOAD_PASSWORD || '';
const authenticatePassword = password => Boolean(UPLOAD_SECRET && password && password === UPLOAD_SECRET);

const saveSource = async (key, fileName, rows, buffer, contentType, dropdowns = { worked: [], noReason: [] }, columns = []) => SourceDataset.findOneAndUpdate(
  { key },
  { key, fileName, rows, columns, fileData: buffer, contentType: contentType || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', dropdowns, uploadedAt: new Date() },
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
    const [train, kav, master] = await Promise.all([
      SourceDataset.findOne({ key: 'TRAINS_PER_DAY' }).lean(),
      SourceDataset.findOne({ key: 'KAVACH_LOCO_DETAILS' }).lean(),
      SourceDataset.findOne({ key: MASTER_DATA_KEY }).lean()
    ]);
    res.json({
      initialized: Boolean(train && kav),
      train: train && { fileName: train.fileName, uploadedAt: train.uploadedAt, count: train.rows.length, hasFile: Boolean(train.fileData), columns: train.columns || [], dropdowns: train.dropdowns || { worked: [], noReason: [] } },
      kavach: kav && { fileName: kav.fileName, uploadedAt: kav.uploadedAt, count: kav.rows.length, hasFile: Boolean(kav.fileData), columns: kav.columns || [] },
      master: master && { fileName: master.fileName, uploadedAt: master.uploadedAt, count: master.rows.length, hasFile: Boolean(master.fileData), columns: master.columns || [] }
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
    // Parse and validate both files fully before replacing either working source.
    const trainRows = trains(trainFile.buffer);
    const kavachRows = kavach(kavachFile.buffer);
    const trainColumns = sourceColumns(trainFile.buffer), kavachColumns = sourceColumns(kavachFile.buffer);
    const dropdowns = trainValidationOptions(trainFile.buffer);
    if (!trainRows.length) throw new Error('TRAINS_PER_DAY workbook contains no data rows.');
    if (!kavachRows.length) throw new Error('Kavach_Loco_Details workbook contains no data rows.');
    await saveSource('TRAINS_PER_DAY', trainFile.originalname, trainRows, trainFile.buffer, trainFile.mimetype, dropdowns, trainColumns);
    await saveSource('KAVACH_LOCO_DETAILS', kavachFile.originalname, kavachRows, kavachFile.buffer, kavachFile.mimetype, { worked: [], noReason: [] }, kavachColumns);
    await syncExcelUsers();
    res.status(201).json({ message: 'Source files uploaded successfully.', trainCount: trainRows.length, kavachCount: kavachRows.length, trainColumns, kavachColumns, dropdowns });
  } catch (e) { res.status(500).json({ message: e.message }); }
});

app.post('/api/upload/update', auth, upload.fields([{ name: 'trainsFile', maxCount: 1 }, { name: 'kavachFile', maxCount: 1 }]), async (req, res) => {
  try {
    const trainFile = req.files?.trainsFile?.[0];
    const kavachFile = req.files?.kavachFile?.[0];
    if (!trainFile && !kavachFile) return res.status(400).json({ message: 'Select at least one updated file.' });
    // Parse all selected replacements before writing any of them so invalid files
    // never replace the last valid dataset.
    let parsedTrain, parsedKavach;
    if (trainFile) {
      if (!isExcel(trainFile.originalname)) return res.status(400).json({ message: 'Invalid train Excel file.' });
      parsedTrain = { rows: trains(trainFile.buffer), columns: sourceColumns(trainFile.buffer), dropdowns: trainValidationOptions(trainFile.buffer) };
      if (!parsedTrain.rows.length) return res.status(400).json({ message: 'Train workbook contains no data rows; the previous source was retained.' });
    }
    if (kavachFile) {
      if (!isExcel(kavachFile.originalname)) return res.status(400).json({ message: 'Invalid Kavach Excel file.' });
      parsedKavach = { rows: kavach(kavachFile.buffer), columns: sourceColumns(kavachFile.buffer) };
      if (!parsedKavach.rows.length) return res.status(400).json({ message: 'Kavach workbook contains no data rows; the previous source was retained.' });
    }
    if (trainFile) await saveSource('TRAINS_PER_DAY', trainFile.originalname, parsedTrain.rows, trainFile.buffer, trainFile.mimetype, parsedTrain.dropdowns, parsedTrain.columns);
    if (kavachFile) {
      await saveSource('KAVACH_LOCO_DETAILS', kavachFile.originalname, parsedKavach.rows, kavachFile.buffer, kavachFile.mimetype, { worked: [], noReason: [] }, parsedKavach.columns);
      await syncExcelUsers();
    }
    res.json({ message: 'Source data updated successfully.' });
  } catch (e) { res.status(500).json({ message: e.message }); }
});


app.post('/api/master-data/upload', userAuth, requireRole('admin'), upload.single('masterFile'), async (req, res) => {
  try {
    const file = req.file;
    if (!file || !isExcel(file.originalname)) return res.status(400).json({ message: 'Select a valid Master_Data_Shed_Rly_Division Excel file.' });
    const existing = await SourceDataset.exists({ key: MASTER_DATA_KEY });
    if (existing && !authenticatePassword(req.body?.password)) {
      return res.status(401).json({ message: 'Password required to replace the existing Master Data file.' });
    }
    const rows = parseMasterData(file.buffer);
    await saveSource(MASTER_DATA_KEY, file.originalname, rows, file.buffer, file.mimetype, { worked: [], noReason: [] }, sourceColumns(file.buffer));
    res.status(existing ? 200 : 201).json({ message: existing ? 'Master Data updated successfully.' : 'Master Data uploaded successfully.', count: rows.length });
  } catch (e) { res.status(e.statusCode || 500).json({ message: e.message }); }
});

app.get('/api/master-data/download', userAuth, requireRole('admin'), async (req, res) => {
  try {
    const source = await SourceDataset.findOne({ key: MASTER_DATA_KEY }).lean();
    if (!source) return res.status(404).json({ message: 'Master Data file is not uploaded yet.' });
    let data = toBuffer(source.fileData);
    if (!data) {
      const rows = source.rows || [];
      const headers = ['SHED','RLY','SHED_NAME','DIVISION'];
      const ws = XLSX.utils.json_to_sheet(rows, { header: headers });
      const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Sheet3');
      data = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    }
    const filename = String(source.fileName || 'Master_Data_Shed_Rly_Division.xlsx').replace(/[^a-zA-Z0-9._ -]/g, '_');
    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition',`attachment; filename="${filename}"`);
    res.send(data);
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
    // Re-read the currently uploaded workbook so stale dropdown metadata from
    // earlier uploads can never override the latest Excel validation rules.
    const sourceBuffer = toBuffer(source.fileData);
    const parsed = sourceBuffer ? trainValidationOptions(sourceBuffer) : { worked: [], noReason: [] };
    const unique = field => Array.from(new Set(source.rows.map(r => String(r[field] ?? '').trim()).filter(Boolean)));
    // Fallbacks are strictly per-field and are used only when that field has
    // no recoverable validation definition in the current workbook.
    const workedOptions = parsed.worked.length ? parsed.worked : unique('worked');
    const reasonOptions = parsed.noReason.length ? parsed.noReason : unique('noReason');
    res.setHeader('Cache-Control', 'no-store');
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
    sourceRow: x.sourceRow, sn: x.sourceRow, date: searchDate, day, sourceData: x.sourceData || {}, sourceColumns: x.sourceColumns || [], kavachSourceData: k?.sourceData || {}, kavachSourceColumns: k?.sourceColumns || [],
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
      return { ...buildRow(x, date, day, k), ...(savedRow || {}), sourceData: x.sourceData || {}, sourceColumns: trainSource.columns || x.sourceColumns || [], kavachSourceData: k?.sourceData || {}, kavachSourceColumns: kavachSource.columns || k?.sourceColumns || [], sourceRow: x.sourceRow, sn: i + 1, date, day, added: Boolean(savedRecord), editing: false, kavachFound: Boolean(k), kavachError: '' };
    });
    res.json({ date, weekday: day, total: enriched.length, columns: trainSource.columns || (rows[0]?.sourceColumns || []), kavachColumns: kavachSource.columns || [], rows: enriched });
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
    const [source, trainSource] = await Promise.all([SourceDataset.findOne({ key: 'KAVACH_LOCO_DETAILS' }).lean(), SourceDataset.findOne({ key: 'TRAINS_PER_DAY' }).lean()]);
    if (!source) return res.status(409).json({ message: 'Kavach source is not uploaded.' });
    const match = source.rows.find(x => norm(x.locoNo) === locoNo);
    if (!match) return res.status(400).json({ message: `Loco_No ${row.locoNo} was not found in Kavach_Loco_Details.xlsx.` });

    const existing = await SavedRow.findOne({ searchDate, sourceRow: row.sourceRow }).lean();
    let base = existing?.row ? { ...existing.row } : { ...row };
    if (existing?.row) {
      for (const field of EDITABLE_SAVED_FIELDS) base[field] = row[field] ?? '';
    }
    const savedRow = {
      ...base, sourceRow: row.sourceRow, sourceData: row.sourceData || trainSource?.rows?.find(x => Number(x.sourceRow) === Number(row.sourceRow))?.sourceData || base.sourceData || {}, sourceColumns: trainSource?.columns || row.sourceColumns || base.sourceColumns || [], kavachSourceData: match.sourceData || {}, kavachSourceColumns: source.columns || match.sourceColumns || [], locoNo: match.locoNo, shed: match.shed || '', locoType: match.locoType || '', rly: match.rly || '',
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
    const [trainSource, kavachSource] = await Promise.all([SourceDataset.findOne({ key: 'TRAINS_PER_DAY' }).lean(), SourceDataset.findOne({ key: 'KAVACH_LOCO_DETAILS' }).lean()]);
    const sourceByRow = new Map((trainSource?.rows || []).map(r => [Number(r.sourceRow), r]));
    const kavachByLoco = new Map((kavachSource?.rows || []).map(r => [norm(r.locoNo), r]));
    const isShed = user.role === 'shed';
    const allowedReasons = isShed
      ? new Set(['KAVACH DEFECTIVE','KAVACH DO NOT START','KAVACH ISOLATION PREV.'])
      : new Set(['KAVACH DEFECTIVE','KAVACH DO NOT START','KAVACH ISOLATION PREV.','KERNEX APP PENDING','KERNEX APP. PENDING']);
    const rows = all.filter(x => {
      const r=x.row||{};
      const identity=isShed ? String(r.shed||'').trim().toUpperCase()===user.username : (String(r.shed||'').trim().toUpperCase()===(user.shedIdentifier||'') && String(r.kavachMake||'').trim().toUpperCase()===(user.oemIdentifier||''));
      const worked=String(r.worked||'').trim().toUpperCase();
      const reason=String(r.noReason||'').trim().toUpperCase();
      const closed = isShed ? Boolean(x.shedRemarkClosed) : Boolean(x.oemRemarkClosed);
      return identity && !closed && (worked==='NO' || worked==='PARTIALLY WORKING') && allowedReasons.has(reason);
    });
    const enrichedRows = rows.map(x => { const current = sourceByRow.get(Number(x.sourceRow)); return { ...x, row: { ...(x.row || {}), sourceData: current?.sourceData || x.row?.sourceData || {}, sourceColumns: trainSource?.columns || current?.sourceColumns || x.row?.sourceColumns || [], kavachSourceData: kavachByLoco.get(norm(x.row?.locoNo))?.sourceData || x.row?.kavachSourceData || {}, kavachSourceColumns: kavachSource?.columns || x.row?.kavachSourceColumns || [] } }; });
    res.json({ user: publicUser(user), rows: enrichedRows });
  } catch(e){ res.status(500).json({message:e.message}); }
});

app.patch('/api/remark-records/:id', userAuth, requireRole('shed','oem'), async (req,res) => {
  try {
    const record = await SavedRow.findById(req.params.id);
    if (!record) return res.status(404).json({message:'Record not found.'});
    const row=record.row||{}, user=req.authUser;
    const isShed=user.role==='shed';
    const identity=isShed ? String(row.shed||'').trim().toUpperCase()===user.username : (String(row.shed||'').trim().toUpperCase()===(user.shedIdentifier||'') && String(row.kavachMake||'').trim().toUpperCase()===(user.oemIdentifier||''));
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
    const identity=isShed ? String(row.shed||'').trim().toUpperCase()===user.username : (String(row.shed||'').trim().toUpperCase()===(user.shedIdentifier||'') && String(row.kavachMake||'').trim().toUpperCase()===(user.oemIdentifier||''));
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
function validKavachMakeValue(x) {
  const value = String(x?.row?.kavachMake ?? '').trim().toUpperCase();
  return Boolean(value && value !== 'N' && value !== '-' && value !== '—');
}
// Shared definition for a train counted as running with Kavach in Summary and
// Summary & Details. Keep this predicate as the single source of truth.
function isRunningWithKavach(x) {
  return validKavachMakeValue(x);
}
const REMARK_PENDING_REASONS = new Set(['KAVACH DEFECTIVE','KAVACH DO NOT START','KAVACH ISOLATION PREV.','KERNEX APP PENDING','KERNEX APP. PENDING']);
function remarkPendingMatch(x, role, groupValue='') {
  const r=x?.row||{};
  const worked=String(r.worked||'').trim().toUpperCase();
  const reason=savedReason(x);
  if (!['NO','PARTIALLY WORKING'].includes(worked) || !REMARK_PENDING_REASONS.has(reason)) return false;
  const group=String(groupValue||'').trim().toUpperCase();
  if(role==='shed') {
    const shed=String(r.shed||'').trim();
    return (!group || shed.toUpperCase()===group) && !String(r.shedRemark||x.shedRemark||'').trim();
  }
  if(role==='oem') {
    const oem=String(r.kavachMake||'').trim();
    return (!group || oem.toUpperCase()===group) && !String(r.oem||x.oem||'').trim();
  }
  return false;
}
const NON_KAVACH_CATEGORIES = Object.freeze({
  LINK: 'MISLINK LOCO BY LOCO LINK DIV',
  ATTACH: 'MISLINK LOCO BY LOCO ATTACH DIV',
  PLAIN: 'NON KAVACH LOCO'
});
const DEFECTIVE_REASONS = new Set(['KAVACH DEFECTIVE','KERNEX APP. PENDING','KAVACH DO NOT START','KAVACH ISOLATION PREV.']);
const DETAIL_KAVACH_DEFECTIVE_REASONS = DEFECTIVE_REASONS;
const RUNNING_WITH_KAVACH_REASONS = new Set(['KAVACH DEFECTIVE','OTHER']);
const SUMMARY_METRICS = new Set(['totalTrains','fitted','runningWithKavach','detailKavachMake','working','defective','kernex','doNotStart','isolation','other','total','nonKavach','nonKavachLink','nonKavachAttach','nonKavachPlain','otherMake','remarkPendingShed','remarkPendingOem']);
const SUMMARY_METRIC_LABELS = {
  totalTrains: 'TOTAL TRAINS LOCO IN KAVACH SECTION', fitted: 'NO. OF LOCO KAVACH FITTED', runningWithKavach: 'TOTAL NO. OF TRAIN RUNNING WITH KAVACH', working: 'NO. OF LOCO KAVACH WORKING',
  defective: 'KAVACH DEFECTIVE', kernex: 'KERNEX MAKE APPROVAL PENDING', doNotStart: 'DO NOT START REMARK/ STICKER',
  isolation: 'KAVACH ISOLATION PREVIOUS REMARKS', other: 'OTHER REASON', total: 'TOTAL', nonKavach: 'NON KAVACH LOCO',
  nonKavachLink: 'NON KAVACH LOCO ATTACH BY LAST ATTACH DIVISON IN LINK', nonKavachAttach: 'NON KAVACH LOCO (MISLINK BY LOCO ATTACH DIVISION)',
  defectiveByShed: 'KAVACH DEFECTIVE LOCO',
  nonKavachPlain: 'NON KAVACH LOCO', otherMake: 'OTHER MAKE',
  detailKavachMake: 'TOTAL NO. OF TRAIN RUNNING WITH KAVACH',
  remarkPendingShed: 'SHED REMARKS PENDING',
  remarkPendingOem: 'OEM REMARKS PENDING'
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
      return isRunningWithKavach(x);
    case 'detailKavachMake':
      return isRunningWithKavach(x) && (!group || String(x?.row?.kavachMake || '').trim() === group);
    case 'remarkPendingShed':
      return remarkPendingMatch(x, 'shed', group);
    case 'remarkPendingOem':
      return remarkPendingMatch(x, 'oem', group);
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
function normalizedMake(value) { return String(value ?? '').trim().toUpperCase(); }
function metricRecords(records, metric, groupValue='', makeValue='') {
  const requestedMake = normalizedMake(makeValue);
  return records.filter(x => (!requestedMake || normalizedMake(x?.row?.kavachMake) === requestedMake) && metricRecordMatches(x, metric, groupValue));
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


// Public About Us summary: one fixed rolling 30-day window anchored to the
// supplied login/page-entry timestamp, with reason labels read from the live
// uploaded source workbook. Counts always come from saved operational records.
app.get('/api/public/about-summary', async (req,res)=>{
  try {
    const anchorRaw=String(req.query.anchor||'');
    const parsedAnchor=anchorRaw ? new Date(anchorRaw) : new Date();
    const anchor=Number.isNaN(parsedAnchor.getTime()) ? new Date() : parsedAnchor;
    const start=new Date(anchor.getTime()-30*24*60*60*1000);
    const fromDate=formatDMY(start), toDate=formatDMY(anchor);
    const [all,source]=await Promise.all([
      SavedRow.find({}).sort({searchDate:1,sourceRow:1}).lean(),
      SourceDataset.findOne({key:'TRAINS_PER_DAY'}).lean()
    ]);
    const records=filterSaved(all,fromDate,toDate,'ALL');
    // Always derive the reason list from the currently uploaded workbook's
    // validation definition first. Older SourceDataset metadata may have been
    // populated from the Train Worked with Kavach list, which caused YES/NO/
    // PARTIALLY WORKING to appear as reason cards.
    const workbookBuffer=toBuffer(source?.fileData);
    const parsedOptions=workbookBuffer ? trainValidationOptions(workbookBuffer) : { worked: [], noReason: [] };
    const savedReasonOptions=Array.isArray(source?.dropdowns?.noReason)?source.dropdowns.noReason:[];
    const options=parsedOptions.noReason.length ? parsedOptions.noReason : savedReasonOptions;
    const normalized=v=>String(v??'').trim().replace(/\s+/g,' ').toUpperCase();
    const workingValues=new Set(['NA']);
    const statusValues=new Set(['YES','NO','PARTIALLY WORKING','PARTIALLY WKG','PARTIAL WORKING']);
    const reasons=[]; const seen=new Set();
    for(const raw of options){
      const label=String(raw??'').trim().replace(/\s+/g,' '), key=normalized(label);
      if(!label||workingValues.has(key)||statusValues.has(key)||seen.has(key)) continue;
      seen.add(key); reasons.push({label,key,count:0});
    }
    // If validation metadata cannot be read (for example a legacy workbook),
    // derive only reason-field values from the uploaded source rows as a safe
    // fallback. Never fall back to the separate working-status field.
    if(!reasons.length){
      for(const raw of (source?.rows||[]).map(r=>r.noReason)){
        const label=String(raw??'').trim().replace(/\s+/g,' '), key=normalized(label);
        if(!label||workingValues.has(key)||statusValues.has(key)||seen.has(key)) continue;
        seen.add(key); reasons.push({label,key,count:0});
      }
    }
    let working=0;
    for(const record of records){
      const reason=normalized(record?.row?.noReason);
      if(workingValues.has(reason)||!reason){working++;continue;}
      const card=reasons.find(r=>r.key===reason);
      if(card) card.count++;
    }
    res.json({anchor:anchor.toISOString(),fromDate,toDate,totalTrains:records.length,kavachWorking:working,reasons:reasons.map(({label,count})=>({label,count})),updatedAt:new Date().toISOString()});
  }catch(e){res.status(500).json({message:'Unable to load About Us summary cards.'});}
});


app.get('/api/data/saved', async (req, res) => {
  try {
    const { startDate, endDate, dir, metric } = req.query;
    if (startDate && !validDate(startDate)) return res.status(400).json({ message: 'Start date must be DD-MM-YYYY.' });
    if (endDate && !validDate(endDate)) return res.status(400).json({ message: 'End date must be DD-MM-YYYY.' });
    if (startDate && endDate && parseDMY(startDate) > parseDMY(endDate)) return res.status(400).json({ message: 'Start date cannot be after end date.' });
    if (dir && !['ALL', 'UP', 'DN'].includes(String(dir).toUpperCase())) return res.status(400).json({ message: 'DIR must be ALL, UP or DN.' });
    if (metric && !SUMMARY_METRICS.has(String(metric))) return res.status(400).json({ message: 'Unknown summary metric.' });
    const rows = filterSaved(await SavedRow.find({}).sort({ createdAt: -1 }).lean(), startDate, endDate, dir, metric);
    const [source, kavachSource] = await Promise.all([SourceDataset.findOne({ key: 'TRAINS_PER_DAY' }).lean(), SourceDataset.findOne({ key: 'KAVACH_LOCO_DETAILS' }).lean()]);
    const sourceByRow = new Map((source?.rows || []).map(r => [Number(r.sourceRow), r]));
    const kavachByLoco = new Map((kavachSource?.rows || []).map(r => [norm(r.locoNo), r]));
    res.json({ total: rows.length, metric: metric || '', metricLabel: SUMMARY_METRIC_LABELS[metric] || '', rows: rows.map((x, i) => { const current = sourceByRow.get(Number(x.sourceRow)); const k = kavachByLoco.get(norm(x.row?.locoNo)); return { ...x, row: { ...(current?.sourceData ? { sourceData: current.sourceData, sourceColumns: source?.columns || current.sourceColumns || [] } : {}), ...(x.row || {}), sourceData: current?.sourceData || x.row?.sourceData || {}, sourceColumns: source?.columns || current?.sourceColumns || x.row?.sourceColumns || [], kavachSourceData: k?.sourceData || x.row?.kavachSourceData || {}, kavachSourceColumns: kavachSource?.columns || k?.sourceColumns || x.row?.kavachSourceColumns || [], sn: i + 1 } }; }) });
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


function makeShedRemarkPendingWorkbook(reportRows, fromDate, toDate) {
  const headers=['S.No.','SHED','RLY','SHED_NAME','DIVISION','Remark Updated','Remark Pending','Total'];
  const ws=XLSXStyle.utils.aoa_to_sheet([
    [`SHED REMARK PENDING IN KAVACH POSITION From ${fromDate} To ${toDate}`],
    headers,
    ...reportRows.map(r=>headers.map(h=>r[h] ?? ''))
  ]);
  ws['!merges']=[{s:{r:0,c:0},e:{r:0,c:headers.length-1}}];
  ws['!cols']=[8,14,12,34,22,18,18,12].map(wch=>({wch}));
  ws['!freeze']={xSplit:0,ySplit:2,topLeftCell:'A3',activePane:'bottomRight',state:'frozen'};
  const totalUpdated=reportRows.reduce((sum,r)=>sum+Number(r['Remark Updated']||0),0);
  const totalPending=reportRows.reduce((sum,r)=>sum+Number(r['Remark Pending']||0),0);
  const totalAll=reportRows.reduce((sum,r)=>sum+Number(r.Total||0),0);
  const totalRow=['','TOTAL','','','',''+totalUpdated,''+totalPending,''+totalAll];
  XLSXStyle.utils.sheet_add_aoa(ws,[totalRow],{origin:{r:reportRows.length+2,c:0}});
  const lastDataRow=reportRows.length+2;
  ws['!autofilter']={ref:`A2:H${Math.max(2,lastDataRow)}`};
  ws['!rows']=[{hpt:30},{hpt:28}];
  const headerStyle={font:{bold:true,color:{rgb:'FFFFFF'},sz:10},fill:{patternType:'solid',fgColor:{rgb:'173B6B'}},alignment:{horizontal:'center',vertical:'center',wrapText:true},border:{top:{style:'thin',color:{rgb:'9AA9B8'}},bottom:{style:'thin',color:{rgb:'9AA9B8'}},left:{style:'thin',color:{rgb:'9AA9B8'}},right:{style:'thin',color:{rgb:'9AA9B8'}}}};
  const titleStyle={font:{bold:true,color:{rgb:'173B6B'},sz:12},fill:{patternType:'solid',fgColor:{rgb:'EAF0F5'}},alignment:{horizontal:'center',vertical:'center',wrapText:true},border:{top:{style:'thin',color:{rgb:'9AA9B8'}},bottom:{style:'thin',color:{rgb:'9AA9B8'}},left:{style:'thin',color:{rgb:'9AA9B8'}},right:{style:'thin',color:{rgb:'9AA9B8'}}}};
  ws['A1'].s=titleStyle;
  for(let c=0;c<headers.length;c++){
    const a=XLSXStyle.utils.encode_cell({r:1,c});
    ws[a].s=headerStyle;
  }
  for(let r=2;r<reportRows.length+2;r++){
    ws['!rows'][r]={hpt:24};
    for(let c=0;c<headers.length;c++){
      const a=XLSXStyle.utils.encode_cell({r,c});
      if(!ws[a])continue;
      ws[a].s={font:{color:{rgb:'1F2937'},sz:10},alignment:{horizontal:c===0||c>=5?'center':'left',vertical:'center',wrapText:true},border:{top:{style:'thin',color:{rgb:'D4DDE5'}},bottom:{style:'thin',color:{rgb:'D4DDE5'}},left:{style:'thin',color:{rgb:'D4DDE5'}},right:{style:'thin',color:{rgb:'D4DDE5'}}},fill:{patternType:'solid',fgColor:{rgb:r%2===0?'F8FBFD':'FFFFFF'}}};
    }
  }
  const totalExcelRow=reportRows.length+2;
  ws['!rows'][totalExcelRow]={hpt:26};
  for(let c=0;c<headers.length;c++){
    const a=XLSXStyle.utils.encode_cell({r:totalExcelRow,c});
    if(!ws[a])continue;
    ws[a].s={font:{bold:true,color:{rgb:'173B6B'},sz:10},alignment:{horizontal:c===0||c>=5?'center':'left',vertical:'center',wrapText:true},border:{top:{style:'medium',color:{rgb:'173B6B'}},bottom:{style:'medium',color:{rgb:'173B6B'}},left:{style:'thin',color:{rgb:'D4DDE5'}},right:{style:'thin',color:{rgb:'D4DDE5'}}},fill:{patternType:'solid',fgColor:{rgb:'EAF0F5'}}};
  }
  const wb=XLSXStyle.utils.book_new();
  XLSXStyle.utils.book_append_sheet(wb,ws,'Sheet1');
  return XLSXStyle.write(wb,{type:'buffer',bookType:'xlsx',cellStyles:true});
}

function exportHeaderToken(value) { return String(value ?? '').trim().replace(/[\r\n]+/g, ' ').replace(/[_./\\-]+/g, ' ').replace(/\s+/g, ' ').toUpperCase(); }
const MAPPED_EXPORT_HEADERS = new Set([
  ...TEMPLATE_HEADERS.map(exportHeaderToken),
  'SHED REMARK', 'SHED REMARKS', 'OEM', 'OEM REMARK', 'OEM REMARKS',
  'IF NO THE REASON', 'IF NO REASON', 'TRAIN WORKED WITH KAVACH YES NO'
]);
function isMappedExportHeader(header) { return MAPPED_EXPORT_HEADERS.has(exportHeaderToken(header)); }
function uniqueExportHeaders(headers) {
  const seen = new Set();
  return headers.filter(h => { const token = exportHeaderToken(h); if (!token || seen.has(token)) return false; seen.add(token); return true; });
}

function exportRows(rows) {
  return rows.map((x, i) => ({
    'S.N': i + 1, DATE: x.searchDate, DAY: x.weekday, KAVACH_WKG_SECTION: x.row?.kavachSection || '', DIR: x.row?.dir || '', TR_TYPE: x.row?.trType || '', T_O_Time: x.row?.time || '',
    'LOCO_ATTACHED_DIV.': x.row?.attachedDiv || '', From_To: x.row?.fromTo || '', LOCO_link: x.row?.locoLink || '', LOCO_LINK_DIV: x.row?.locoLinkDiv || '',
    Day_of_wkg_in_territory: x.row?.dayWork || '', Loco_No: x.row?.locoNo || '', SHED: x.row?.shed || '', Loco_Type: x.row?.locoType || '', RLY: x.row?.rly || '',
    Loco_Kavach_Make: x.row?.kavachMake || '', Brake_system: x.row?.brakeSystem || '', TRAIN_No: x.row?.trainNo || '',
    Train_worked_with_Kavach_YES_NO: x.row?.worked || '',
    'If_No_the_Reason \n( Non Kavach Loco\n/ Kavach defective\n/ KAVACH Fitness cirtification N.Avl.\n/ Crew incompetency)': x.row?.noReason || '',
    Remarks: x.row?.remarks || '', 'Shed Remarks': x.row?.shedRemark || x.shedRemark || '', 'OEM Remarks': x.row?.oem || x.oem || '',
    ...Object.fromEntries(Object.entries(x.row?.sourceData || {}).filter(([key]) => !isMappedExportHeader(key))),
    ...Object.fromEntries(Object.entries(x.row?.kavachSourceData || {}).filter(([key]) => !isMappedExportHeader(key)).map(([key, value]) => [`Kavach — ${key}`, value]))
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
    const [trainSource, kavachSource] = await Promise.all([SourceDataset.findOne({ key: 'TRAINS_PER_DAY' }).lean(), SourceDataset.findOne({ key: 'KAVACH_LOCO_DETAILS' }).lean()]);
    const trainByRow = new Map((trainSource?.rows || []).map(r => [Number(r.sourceRow), r]));
    const kavachByLoco = new Map((kavachSource?.rows || []).map(r => [norm(r.locoNo), r]));
    const exportReadyRows = rows.map(x => { const t=trainByRow.get(Number(x.sourceRow)); const k=kavachByLoco.get(norm(x.row?.locoNo)); return { ...x, row: { ...(x.row||{}), sourceData:t?.sourceData||x.row?.sourceData||{}, sourceColumns:trainSource?.columns||x.row?.sourceColumns||[], kavachSourceData:k?.sourceData||x.row?.kavachSourceData||{}, kavachSourceColumns:kavachSource?.columns||x.row?.kavachSourceColumns||[] } }; });
    const exportedRows = exportRows(exportReadyRows);
    const headers = uniqueExportHeaders([...TEMPLATE_HEADERS, ...Array.from(new Set(exportedRows.flatMap(r => Object.keys(r)))).filter(h => !isMappedExportHeader(h))]);
    const buffer = makeStyledWorkbook('Traction Operation', headers, exportedRows, headers.map((h, i) => i === 20 ? 44 : i === 21 ? 38 : i === 22 ? 28 : i === 23 ? 24 : Math.max(12, Math.min(36, String(h).length + 3))));
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
  const fitted = count('detailKavachMake');
  const runningWithKavach = count('detailKavachMake');
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
      runningWithKavach: count('detailKavachMake'),
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
  // Build make-specific Summary metrics using the SAME predicates as the public
  // Summary page and its popup. No source workbook or hardcoded make list.
  const kavachMakeMap = new Map();
  records.filter(validKavachMakeValue).forEach(x => {
    const name = normalizedMake(x.row?.kavachMake);
    if (!kavachMakeMap.has(name)) kavachMakeMap.set(name, []);
    kavachMakeMap.get(name).push(x);
  });
  const kavachMakeBreakdown = [...kavachMakeMap].map(([name, makeRecords]) => {
    const summary = summaryForRecords(makeRecords, '', 'ALL');
    return {
      name, ...summary, up:makeRecords.filter(x=>savedDir(x)==='UP').length,
      dn:makeRecords.filter(x=>savedDir(x)==='DN').length
    };
  }).sort((a,b)=>b.totalTrains-a.totalTrains||a.name.localeCompare(b.name));
  const divisionRows = (metric, fieldKeys) => {
    const map = new Map();
    metricRecords(records, metric).forEach(x => {
      const name = savedValue(x, fieldKeys) || '—';
      if (!map.has(name)) map.set(name, { name, total: 0, up: 0, dn: 0 });
      const item = map.get(name); item.total++;
      if (savedDir(x) === 'UP') item.up++;
      if (savedDir(x) === 'DN') item.dn++;
    });
    return [...map.values()].sort((a,b)=>b.total-a.total || a.name.localeCompare(b.name));
  };
  const nonKavachLinkByDivision = divisionRows('nonKavachLink', ['locoLinkDiv','LOCO_LINK_DIV','LOCO_LINK_DIVISION','Loco Link Division']);
  const nonKavachPlainByStation = divisionRows('nonKavachPlain', ['attachedDiv','LOCO_ATTACHED_DIV.','LOCO_ATTACHED_DIV','LOCO_ATTACHED_DIVISION','Loco Attached Division']);
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
    total:section(records),up:section(up),dn:section(dn),defectiveRows,kavachMakeBreakdown,
    defectiveByShed:[...shedMap.values()].sort((a,b)=>b.total-a.total || a.name.localeCompare(b.name)),
    nonKavachLinkByDivision, nonKavachPlainByStation, nonKavachPlainRows,otherRows
  };
}

function detailMetricRecords(records, metric, dir='', groupValue='', makeValue='') {
  let rows = records;
  const targetDir = String(dir||'').trim().toUpperCase();
  if (targetDir && targetDir !== 'ALL') rows = rows.filter(x => savedDir(x) === targetDir);
  return metricRecords(rows, metric, groupValue, makeValue);
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
  const drawPageHeader=()=>{rectFill(M,H-M-38,contentW,38,0.93);rectStroke(M,H-M-38,contentW,38);text(M+8,H-M-14,'TRACTION OPERATION - KAVACH MONITORING SYSTEM',9.2,true);text(M+8,H-M-27,'West Central Railway, Kota Division',7.0,false);const dateLabel=`Summary & Detail | ${report.fromDate} to ${report.toDate}`;const dateWidth=Math.min(190, dateLabel.length*3.0);text(W-M-8-dateWidth,H-M-14,dateLabel,6.3,true);line(M,H-M-42,W-M,H-M-42);};
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

  y-=6;
  const mrows=[['1','No of Target Trains',report.total.targetTrains,report.up.targetTrains,report.dn.targetTrains],['2','TOTAL NO. OF TRAIN RUNNING WITH KAVACH',report.total.runningWithKavach,report.up.runningWithKavach,report.dn.runningWithKavach],['3','TOTAL NO. OF TRAIN RUNNING WITH KAVACH IN WORKING CONDITION',report.total.workingCondition,report.up.workingCondition,report.dn.workingCondition],['4','NON KAVACH LOCO',report.total.nonKavach,report.up.nonKavach,report.dn.nonKavach],['5','KAVACH DEFECTIVE LOCO',report.total.defective,report.up.defective,report.dn.defective],['6','OTHER MAKE',report.total.otherMake,report.up.otherMake,report.dn.otherMake],['7','OTHER REASON',report.total.otherReason,report.up.otherReason,report.dn.otherReason]];
  table(['','DETAILS','TOTAL','UP','DN'],mrows,[22,300,76,76,77],{alignCenter:true,leftBodyColumns:[1],fontSize:7.1,lineH:8.8,pad:3});
  const divisionRows=(rows)=>rows.length?rows.map((r,i)=>[i+1,r.name,r.total,r.up,r.dn]):[[1,'No data',0,0,0]];
  const makeHeaders=['MAKE','TOTAL TRAINS','KAVACH WORKING','KAVACH DEFECTIVE','DO NOT START','ISOLATION','OTHER REASON','TOTAL','% WORKING'];
  const makeRows=report.kavachMakeBreakdown.length?report.kavachMakeBreakdown.map(r=>[r.name,r.totalTrains,r.working,r.defective,r.doNotStart,r.isolation,r.other,r.total,`${r.percentWorking}%`]):[['No data',0,0,0,0,0,0,0,'0%']];
  section('KAVACH MAKE SUMMARY BREAK UP',makeHeaders,makeRows,[82,58,60,63,60,60,60,58,55],{alignCenter:true,leftBodyColumns:[0],fontSize:5.4,lineH:7.2,pad:2.5});
  section('KAVACH DEFECTIVE LOCO',['','SHED','TOTAL','UP','DN'],divisionRows(report.defectiveByShed),[22,300,76,76,77],{alignCenter:true,leftBodyColumns:[1],fontSize:7,lineH:8.6,pad:3});
  const divisionRowsPdf=(rows)=>rows.length?rows.map((r,i)=>[i+1,r.name,r.total,r.up,r.dn]):[[1,'No data',0,0,0]];
  section('NON KAVACH LOCO',['','Station','TOTAL','UP','DN'],divisionRowsPdf(report.nonKavachPlainByStation),[22,300,76,76,77],{alignCenter:true,leftBodyColumns:[1],fontSize:7,lineH:8.6,pad:3});
  const nonRows=report.nonKavachPlainRows?.length?report.nonKavachPlainRows.map(r=>[r.sn,r.date,r.trainNo,r.locoNo,r.shed,r.locoAttachDiv,r.tlcRemark]):[[1,'No data','','','','','']];
  section('NON KAVACH LOCO ATTACH BY LAST ATTACH DIVISON IN LINK',['S.No.','Date','Train No','Loco No','SHED','Loco Attach Div','TLC Remark'],nonRows,[22,65,52,52,45,63,252],{alignCenter:true,leftBodyColumns:[1,2,3,4,5,6],fontSize:6.6,lineH:8.1,pad:3});
  const detailTable=(title,rows)=>{const heads=['S.N','Date','Train No','Loco No','SHED','KAVACH MAKE','TLC Remark','Shed Remark','OEM Remark'];const rr=rows.length?rows.map(r=>[r.sn,r.date,r.trainNo,r.locoNo,r.shed,r.kavachMake,r.tlcRemark,r.shedRemark,r.oemRemark]):[[1,'No data','','','','','','','']];section(title,heads,rr,[18,55,40,38,32,42,105,105,116],{alignCenter:true,leftBodyColumns:[1,2,3,4,5,6,7,8],fontSize:5.1,lineH:6.7,pad:2.5});};
  detailTable('KAVACH DEFECTIVE LOCO DETAILS',report.defectiveRows);detailTable('OTHER REASON',report.otherRows);
  if(ops.length)finishPage();return buildPdfDocument(pages);
}

async function getSummaryMetricRecords({fromDate,toDate,dir='ALL',metric='totalTrains',groupValue='',makeValue=''}) {
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
  const rows=detailMetricRecords(base,metricName,targetDir,groupValue||'',makeValue||'');
  return {fromDate,toDate,dir:targetDir,metric:metricName,makeValue:makeValue||'',metricLabel:(makeValue?`${makeValue} — `:'')+(SUMMARY_METRIC_LABELS[metricName]||metricName),total:rows.length,rows};
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
    const {fromDate,toDate,dir,groupValue,makeValue} = req.query;
    const metric=canonicalMetric(req.query.metric||'');
    if(!validDate(fromDate||'') || !validDate(toDate||'')) return res.status(400).json({message:'From Date and To Date must be DD-MM-YYYY.'});
    if(parseDMY(fromDate)>parseDMY(toDate)) return res.status(400).json({message:'From Date cannot be after To Date.'});
    if(!SUMMARY_METRICS.has(metric) && !['defectiveByShed'].includes(metric)) return res.status(400).json({message:'Unknown summary metric.'});
    const targetDir=String(dir||'ALL').trim().toUpperCase();
    if(!['ALL','UP','DN'].includes(targetDir)) return res.status(400).json({message:'DIR must be ALL, UP or DN.'});
    const all=await SavedRow.find({}).sort({searchDate:1,sourceRow:1}).lean();
    const base=filterSaved(all,fromDate,toDate,targetDir);
    const rows=metricRecords(base,String(metric),String(groupValue||''),String(makeValue||''));
    const headers=TEMPLATE_HEADERS;
    const buffer=makeStyledWorkbook('Summary Records',headers,exportRows(rows),headers.map((h,i)=>i===20?44:i===21?38:i===22?28:i===23?24:Math.max(12,Math.min(28,String(h).length+3))));
    const label=String(SUMMARY_METRIC_LABELS[metric]||metric).replace(/[^A-Za-z0-9]+/g,'_').replace(/^_|_$/g,'');
    const suffix=`_${fromDate}_to_${toDate}_${targetDir}_${label||'records'}`;
    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition',`attachment; filename="Summary_${suffix}.xlsx"`);
    res.send(buffer);
  } catch(e){res.status(500).json({message:e.message});}
});


function remarkStatusMatch(x, role, status, groupValue='') {
  const r=x?.row||{};
  const worked=String(r.worked||'').trim().toUpperCase();
  const reason=savedReason(x);
  if (!['NO','PARTIALLY WORKING'].includes(worked) || !REMARK_PENDING_REASONS.has(reason)) return false;
  const group=String(groupValue||'').trim().toUpperCase();
  if(role==='shed') {
    const shed=String(r.shed||'').trim();
    if(group && shed.toUpperCase()!==group) return false;
    const has=Boolean(String(r.shedRemark||x.shedRemark||'').trim());
    return status==='updated' ? has : status==='pending' ? !has : true;
  }
  if(role==='oem') {
    const oem=String(r.kavachMake||'').trim();
    if(group && oem.toUpperCase()!==group) return false;
    const has=Boolean(String(r.oem||x.oem||'').trim());
    return status==='updated' ? has : status==='pending' ? !has : true;
  }
  return false;
}

async function getRemarkPendingRecords({role,fromDate,toDate,groupValue='',status='pending',authUser=null}) {
  if(!['shed','oem'].includes(role)) { const e=new Error('Select Shed or OEM.'); e.statusCode=400; throw e; }
  if(!validDate(fromDate||'') || !validDate(toDate||'')) { const e=new Error('From Date and To Date must be DD-MM-YYYY.'); e.statusCode=400; throw e; }
  if(parseDMY(fromDate)>parseDMY(toDate)) { const e=new Error('From Date cannot be after To Date.'); e.statusCode=400; throw e; }
  if(authUser?.role==='shed' && role!=='shed') { const e=new Error('Shed users can only view the SHED report.'); e.statusCode=403; throw e; }
  if(authUser?.role==='oem' && role!=='oem') { const e=new Error('OEM users can only view the OEM report.'); e.statusCode=403; throw e; }
  if(!['pending','updated','all'].includes(status)) { const e=new Error('Invalid remark status.'); e.statusCode=400; throw e; }

  const all=await SavedRow.find({}).sort({searchDate:1,sourceRow:1}).lean();
  let base=filterSaved(all,fromDate,toDate,'ALL');
  if(authUser?.role==='shed') base=base.filter(x=>String(x.row?.shed||'').trim().toUpperCase()===String(authUser.username||'').trim().toUpperCase());
  if(authUser?.role==='oem') base=base.filter(x=>String(x.row?.kavachMake||'').trim().toUpperCase()===String(authUser.username||'').trim().toUpperCase());
  const rows=base.filter(x=>remarkStatusMatch(x,role,status,groupValue));
  return {role,fromDate,toDate,groupValue:groupValue||'',status,total:rows.length,rows};
}

async function getMasterMap() {
  const source=await SourceDataset.findOne({key:MASTER_DATA_KEY}).lean();
  const map=new Map();
  (source?.rows||[]).forEach(r=>{
    const shed=String(r.SHED||r.shed||'').trim().toUpperCase();
    if(shed) map.set(shed,{SHED:String(r.SHED||'').trim(),RLY:String(r.RLY||'').trim()||'N/A',SHED_NAME:String(r.SHED_NAME||'').trim()||'N/A',DIVISION:String(r.DIVISION||'').trim()||'N/A'});
  });
  return {map,loaded:Boolean(source)};
}

app.get('/api/remarks-pending', async (req,res,next)=>{ try { req.authUser = await getAuthenticatedUser(req,{touch:false}); next(); } catch { req.authUser=null; next(); } }, async (req,res)=>{
  try {
    const role=String(req.query.role||'').trim().toLowerCase();
    const {fromDate,toDate}=req.query;
    const authUser=req.authUser;
    if(!['shed','oem'].includes(role)) return res.status(400).json({message:'Select Shed or OEM.'});
    if(authUser?.role==='shed' && role!=='shed') return res.status(403).json({message:'Shed users can only view the SHED report.'});
    if(authUser?.role==='oem' && role!=='oem') return res.status(403).json({message:'OEM users can only view the OEM report.'});

    if(role==='oem') {
      const payload=await getRemarkPendingRecords({role,fromDate,toDate,status:'pending',authUser});
      const groups=new Map();
      payload.rows.forEach(x=>{
        const name=String(x.row?.kavachMake||'').trim();
        if(name) groups.set(name,(groups.get(name)||0)+1);
      });
      const rows=[...groups.entries()].map(([name,count])=>({name,count})).sort((a,b)=>b.count-a.count||a.name.localeCompare(b.name));
      return res.json({role,fromDate,toDate,metric:'remarkPendingOem',rows,total:rows.reduce((s,r)=>s+r.count,0),legacyOem:true});
    }

    const basePayload=await getRemarkPendingRecords({role,fromDate,toDate,status:'all',authUser});
    const {map,loaded}=await getMasterMap();
    const groups=new Map();
    basePayload.rows.forEach(x=>{
      const shed=String(x.row?.shed||'').trim();
      if(!shed) return;
      const key=shed.toUpperCase();
      if(!groups.has(key)) groups.set(key,{shed,rly:'N/A',shedName:'N/A',division:'N/A',updated:0,pending:0,total:0});
      const g=groups.get(key);
      const master=map.get(key);
      if(master){g.rly=master.RLY;g.shedName=master.SHED_NAME;g.division=master.DIVISION;}
      else console.warn(`[Remarks Pending] Missing Master Data mapping for SHED: ${shed}`);
      if(remarkStatusMatch(x,'shed','updated',shed)) g.updated++;
      if(remarkStatusMatch(x,'shed','pending',shed)) g.pending++;
      g.total++;
    });
    const rows=[...groups.values()].sort((a,b)=>b.total-a.total||a.shed.localeCompare(b.shed)).map((r,i)=>({sno:i+1,...r}));
    const totals={updated:rows.reduce((s,r)=>s+r.updated,0),pending:rows.reduce((s,r)=>s+r.pending,0),total:rows.reduce((s,r)=>s+r.total,0)};
    res.json({role,fromDate,toDate,rows,totals,total:totals.total,masterDataUploaded:loaded});
  } catch(e){res.status(e.statusCode||500).json({message:e.message});}
});

app.get('/api/remarks-pending/records', async (req,res,next)=>{ try { req.authUser = await getAuthenticatedUser(req,{touch:false}); next(); } catch { req.authUser=null; next(); } }, async (req,res)=>{
  try {
    const role=String(req.query.role||'').trim().toLowerCase();
    const payload=await getRemarkPendingRecords({role,fromDate:req.query.fromDate,toDate:req.query.toDate,groupValue:req.query.groupValue||'',status:req.query.status||'pending',authUser:req.authUser});
    if(!req.authUser) payload.rows=payload.rows.map(x=>({_id:String(x._id),searchDate:x.searchDate,sourceRow:x.sourceRow,row:{date:x.row?.date,trainNo:x.row?.trainNo,locoNo:x.row?.locoNo,shed:x.row?.shed,kavachMake:x.row?.kavachMake,worked:x.row?.worked,noReason:x.row?.noReason,remarks:x.row?.remarks,shedRemark:x.row?.shedRemark||x.shedRemark||'',oem:x.row?.oem||x.oem||''},shedRemark:x.shedRemark||'',oem:x.oem||''}));
    res.json({...payload,metricLabel:`${role.toUpperCase()} REMARK ${String(payload.status).toUpperCase()}`});
  } catch(e){res.status(e.statusCode||500).json({message:e.message});}
});

app.get('/api/remarks-pending/export', async (req,res,next)=>{ try { req.authUser = await getAuthenticatedUser(req,{touch:false}); next(); } catch { req.authUser=null; next(); } }, async (req,res)=>{
  try {
    const role=String(req.query.role||'').trim().toLowerCase();
    const status=String(req.query.status||'pending').trim().toLowerCase();
    const payload=await getRemarkPendingRecords({role,fromDate:req.query.fromDate,toDate:req.query.toDate,groupValue:req.query.groupValue||'',status,authUser:req.authUser});
    let rows=payload.rows;
    if(!req.authUser) rows=rows.map(x=>({_id:x._id,searchDate:x.searchDate,weekday:x.weekday,sourceRow:x.sourceRow,row:{date:x.row?.date,day:x.row?.day,trainNo:x.row?.trainNo,locoNo:x.row?.locoNo,shed:x.row?.shed,kavachMake:x.row?.kavachMake,worked:x.row?.worked,noReason:x.row?.noReason,remarks:x.row?.remarks,shedRemark:x.row?.shedRemark||x.shedRemark||'',oem:x.row?.oem||x.oem||''},shedRemark:x.shedRemark||'',oem:x.oem||''}));
    if(req.query.kind==='report' && role==='shed') {
      const {map}=await getMasterMap();
      const groups=new Map();
      rows.forEach(x=>{
        const shed=String(x.row?.shed||'').trim(); if(!shed)return;
        const key=shed.toUpperCase(); if(!groups.has(key)) groups.set(key,{shed,rly:'N/A',shedName:'N/A',division:'N/A',updated:0,pending:0,total:0});
        const g=groups.get(key), master=map.get(key);
        if(master){g.rly=master.RLY;g.shedName=master.SHED_NAME;g.division=master.DIVISION;}
        if(remarkStatusMatch(x,'shed','updated',shed))g.updated++;
        if(remarkStatusMatch(x,'shed','pending',shed))g.pending++;
        g.total++;
      });
      const reportRows=[...groups.values()].sort((a,b)=>b.total-a.total||a.shed.localeCompare(b.shed)).map((r,i)=>({'S.No.':i+1,SHED:r.shed,RLY:r.rly,SHED_NAME:r.shedName,DIVISION:r.division,'Remark Updated':r.updated,'Remark Pending':r.pending,Total:r.total}));
      const headers=['S.No.','SHED','RLY','SHED_NAME','DIVISION','Remark Updated','Remark Pending','Total'];
      const buffer=makeShedRemarkPendingWorkbook(reportRows,payload.fromDate,payload.toDate);
      res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition',`attachment; filename="Report_Kavach_Shed_Remark_Pending_${payload.fromDate}_to_${payload.toDate}.xlsx"`);
      return res.send(buffer);
    }
    const buffer=makeStyledWorkbook('Remark Records',TEMPLATE_HEADERS,exportRows(rows),TEMPLATE_HEADERS.map((h,i)=>i===20?44:i===21?38:i===22?28:i===23?24:Math.max(12,Math.min(28,String(h).length+3))));
    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition',`attachment; filename="Remarks_${role}_${status}_${payload.fromDate}_to_${payload.toDate}.xlsx"`);
    res.send(buffer);
  } catch(e){res.status(e.statusCode||500).json({message:e.message});}
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
