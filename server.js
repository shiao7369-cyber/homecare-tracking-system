const express = require('express');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');
const multer = require('multer');
const XLSX = require('xlsx');
const admin = require('firebase-admin');
const helmet = require('helmet');
const rateLimit = require('express-rate-limit');
const bcrypt = require('bcryptjs');
const cors = require('cors');

const app = express();
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 10 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    const allowed = ['.xls', '.xlsx'];
    const ext = path.extname(file.originalname).toLowerCase();
    if (allowed.includes(ext)) cb(null, true);
    else cb(new Error('僅接受 .xls 或 .xlsx 檔案'));
  }
});
const PORT = process.env.PORT || 3000;

// ===== 安全性中介軟體 =====
app.use(helmet({
  contentSecurityPolicy: {
    directives: {
      defaultSrc: ["'self'"],
      scriptSrc: ["'self'", "'unsafe-inline'", "https://www.gstatic.com", "https://cdn.jsdelivr.net"],
      scriptSrcAttr: ["'unsafe-inline'"],
      styleSrc: ["'self'", "'unsafe-inline'", "https://fonts.googleapis.com", "https://cdn.jsdelivr.net"],
      fontSrc: ["'self'", "https://fonts.gstatic.com", "https://cdn.jsdelivr.net"],
      connectSrc: ["'self'", "https://*.firebaseio.com", "https://*.googleapis.com", "https://firestore.googleapis.com"],
      imgSrc: ["'self'", "data:"],
      frameSrc: ["'none'"]
    }
  },
  crossOriginEmbedderPolicy: false
}));

app.use(cors({
  origin: process.env.ALLOWED_ORIGINS ? process.env.ALLOWED_ORIGINS.split(',') : true,
  credentials: true
}));

// HTTPS 重導向（Railway 提供 TLS，透過 x-forwarded-proto 判斷）
app.use((req, res, next) => {
  if (req.headers['x-forwarded-proto'] === 'http') {
    return res.redirect(301, 'https://' + req.headers.host + req.url);
  }
  next();
});

app.use(express.json({ limit: '10mb' }));

// ===== 靜態檔案限制：只提供安全的檔案 =====
const ALLOWED_EXTENSIONS = ['.html', '.css', '.js', '.png', '.jpg', '.ico', '.svg', '.woff', '.woff2', '.ttf'];
const BLOCKED_FILES = ['serviceAccountKey.json', 'users.json', 'package.json', 'package-lock.json',
  'firebase.json', '.firebaserc', 'firestore.rules', 'railway.toml', '.gitignore'];

app.use((req, res, next) => {
  const reqPath = decodeURIComponent(req.path);
  const filename = path.basename(reqPath);
  if (BLOCKED_FILES.includes(filename)) {
    return res.status(404).send('Not found');
  }
  if (reqPath.includes('..')) {
    return res.status(400).send('Bad request');
  }
  next();
});
app.use(express.static(path.join(__dirname), {
  dotfiles: 'deny',
  index: false
}));

// ===== Rate Limiting =====
const loginLimiter = rateLimit({
  windowMs: 15 * 60 * 1000,
  max: 10,
  message: { error: '登入嘗試次數過多，請 15 分鐘後再試' },
  standardHeaders: true,
  legacyHeaders: false
});

const apiLimiter = rateLimit({
  windowMs: 1 * 60 * 1000,
  max: 100,
  message: { error: '請求過於頻繁，請稍後再試' }
});

app.use('/api/', apiLimiter);

// ===== Firebase Admin 初始化 =====
let firestoreDb = null;

function initFirebase() {
  try {
    if (process.env.FIREBASE_SERVICE_ACCOUNT) {
      const serviceAccount = JSON.parse(process.env.FIREBASE_SERVICE_ACCOUNT);
      admin.initializeApp({ credential: admin.credential.cert(serviceAccount) });
    } else if (fs.existsSync(path.join(__dirname, 'serviceAccountKey.json'))) {
      const serviceAccount = require('./serviceAccountKey.json');
      admin.initializeApp({ credential: admin.credential.cert(serviceAccount) });
    } else {
      admin.initializeApp({ projectId: 'homecare-system-f37ea' });
    }
    firestoreDb = admin.firestore();
    console.log('Firebase Admin 初始化成功');
  } catch (err) {
    console.error('Firebase Admin 初始化失敗:', err.message);
  }
}

initFirebase();

// ===== 稽核日誌 =====
const AUDIT_COLLECTION = 'audit_logs';

async function auditLog(action, userId, details = {}) {
  const entry = {
    action,
    userId: userId || 'anonymous',
    timestamp: new Date().toISOString(),
    ip: details.ip || '',
    details: details.msg || ''
  };
  console.log(`[AUDIT] ${entry.action} by ${entry.userId} - ${entry.details}`);
  if (firestoreDb) {
    try {
      await firestoreDb.collection(AUDIT_COLLECTION).add(entry);
    } catch (e) { /* 稽核寫入失敗不影響主流程 */ }
  }
}

// ===== 使用者資料（Firestore 持久化）=====
const USERS_COLLECTION = 'users';
let _usersCache = null;
const BCRYPT_ROUNDS = 10;

function getDefaultUsers() {
  return [
    {
      id: crypto.randomUUID(),
      username: '蕭輝哲',
      password: bcrypt.hashSync('Hc@2025!Admin', BCRYPT_ROUNDS),
      displayName: '蕭輝哲',
      role: 'admin',
      status: 'active',
      createdAt: new Date().toISOString(),
      mustChangePassword: true
    }
  ];
}

// 密碼複雜度驗證
function validatePassword(pw) {
  if (pw.length < 8) return '密碼長度至少 8 個字元';
  if (!/[A-Za-z]/.test(pw) || !/[0-9]/.test(pw)) return '密碼須包含英文字母和數字';
  return null;
}

// 兼容舊 SHA-256 密碼的驗證函數
function verifyPassword(inputPw, storedHash) {
  // 新格式：bcrypt hash ($2a$ 或 $2b$ 開頭)
  if (storedHash.startsWith('$2')) {
    return bcrypt.compareSync(inputPw, storedHash);
  }
  // 舊格式：SHA-256 hex (64 字元)
  if (storedHash.length === 64) {
    const sha256 = crypto.createHash('sha256').update(inputPw).digest('hex');
    return sha256 === storedHash;
  }
  return false;
}

// 遷移舊密碼為 bcrypt
async function migratePasswordIfNeeded(user, inputPw) {
  if (!user.password.startsWith('$2')) {
    user.password = bcrypt.hashSync(inputPw, BCRYPT_ROUNDS);
    const users = await loadUsers();
    await saveUsers(users);
  }
}

async function loadUsers() {
  if (_usersCache) return _usersCache;

  if (firestoreDb) {
    try {
      const snapshot = await firestoreDb.collection(USERS_COLLECTION).get();
      if (!snapshot.empty) {
        _usersCache = snapshot.docs.map(doc => doc.data());
        return _usersCache;
      }
    } catch (err) {
      console.error('Firestore 讀取使用者失敗:', err.message);
    }
  }

  const localPath = path.join(__dirname, 'users.json');
  if (fs.existsSync(localPath)) {
    try {
      _usersCache = JSON.parse(fs.readFileSync(localPath, 'utf8'));
      await saveUsers(_usersCache);
      return _usersCache;
    } catch (e) { }
  }

  _usersCache = getDefaultUsers();
  await saveUsers(_usersCache);
  return _usersCache;
}

async function saveUsers(users) {
  _usersCache = users;
  if (firestoreDb) {
    try {
      const batch = firestoreDb.batch();
      const snapshot = await firestoreDb.collection(USERS_COLLECTION).get();
      snapshot.docs.forEach(doc => batch.delete(doc.ref));
      users.forEach(user => {
        const ref = firestoreDb.collection(USERS_COLLECTION).doc(user.id);
        batch.set(ref, user);
      });
      await batch.commit();
      return;
    } catch (err) {
      console.error('Firestore 寫入使用者失敗:', err.message);
    }
  }
  try {
    fs.writeFileSync(path.join(__dirname, 'users.json'), JSON.stringify(users, null, 2), 'utf8');
  } catch (e) {
    console.log('警告：使用者資料僅存於記憶體');
  }
}

// ===== Session 管理（含過期機制）=====
const sessions = {};
const SESSION_MAX_AGE = 8 * 60 * 60 * 1000; // 8 小時
const SESSION_IDLE_TIMEOUT = 30 * 60 * 1000; // 30 分鐘閒置

// 定期清理過期 session
setInterval(() => {
  const now = Date.now();
  for (const [token, session] of Object.entries(sessions)) {
    if (now - session.createdAt > SESSION_MAX_AGE || now - session.lastActivity > SESSION_IDLE_TIMEOUT) {
      delete sessions[token];
    }
  }
}, 5 * 60 * 1000);

function createSession(user) {
  const token = crypto.randomBytes(32).toString('hex');
  sessions[token] = {
    userId: user.id,
    username: user.username,
    displayName: user.displayName,
    role: user.role,
    createdAt: Date.now(),
    lastActivity: Date.now()
  };
  return token;
}

function getSession(req) {
  const token = req.headers['authorization']?.replace('Bearer ', '');
  if (!token || !sessions[token]) return null;

  const session = sessions[token];
  const now = Date.now();

  // 檢查 session 是否過期
  if (now - session.createdAt > SESSION_MAX_AGE || now - session.lastActivity > SESSION_IDLE_TIMEOUT) {
    delete sessions[token];
    return null;
  }

  session.lastActivity = now;
  return session;
}

function requireAuth(req, res, next) {
  const session = getSession(req);
  if (!session) return res.status(401).json({ error: '請先登入' });
  req.session = session;
  next();
}

function requireAdmin(req, res, next) {
  requireAuth(req, res, () => {
    if (req.session.role !== 'admin') return res.status(403).json({ error: '需要管理員權限' });
    next();
  });
}

// ===== API 路由 =====

// 登入（含 rate limit）
app.post('/api/login', loginLimiter, async (req, res) => {
  const { username, password } = req.body;
  if (!username || !password) return res.status(400).json({ error: '請輸入帳號和密碼' });

  const users = await loadUsers();
  const user = users.find(u => u.username === username && u.status === 'active');
  if (!user || !verifyPassword(password, user.password)) {
    await auditLog('LOGIN_FAILED', username, { ip: req.ip, msg: '帳號或密碼錯誤' });
    return res.status(401).json({ error: '帳號或密碼錯誤' });
  }

  // 自動遷移舊 SHA-256 密碼為 bcrypt
  await migratePasswordIfNeeded(user, password);

  const token = createSession(user);
  await auditLog('LOGIN_SUCCESS', user.id, { ip: req.ip, msg: user.username });
  res.json({
    token,
    user: { id: user.id, username: user.username, displayName: user.displayName, role: user.role },
    mustChangePassword: !!user.mustChangePassword,
    sessionMaxAge: SESSION_MAX_AGE,
    sessionIdleTimeout: SESSION_IDLE_TIMEOUT
  });
});

// 登出
app.post('/api/logout', (req, res) => {
  const token = req.headers['authorization']?.replace('Bearer ', '');
  if (token && sessions[token]) {
    auditLog('LOGOUT', sessions[token].userId, { ip: req.ip });
    delete sessions[token];
  }
  res.json({ ok: true });
});

// 取得目前使用者
app.get('/api/me', requireAuth, (req, res) => {
  res.json(req.session);
});

// ===== 使用者管理 (管理員) =====

app.get('/api/users', requireAdmin, async (req, res) => {
  const users = (await loadUsers()).map(u => ({
    id: u.id, username: u.username, displayName: u.displayName,
    role: u.role, status: u.status, createdAt: u.createdAt
  }));
  res.json(users);
});

app.post('/api/users', requireAdmin, async (req, res) => {
  const { username, password, displayName, role } = req.body;
  if (!username || !password) return res.status(400).json({ error: '帳號和密碼為必填' });

  const pwErr = validatePassword(password);
  if (pwErr) return res.status(400).json({ error: pwErr });

  const users = await loadUsers();
  if (users.find(u => u.username === username)) {
    return res.status(400).json({ error: '此帳號已存在' });
  }

  const newUser = {
    id: crypto.randomUUID(),
    username,
    password: bcrypt.hashSync(password, BCRYPT_ROUNDS),
    displayName: displayName || username,
    role: role || 'user',
    status: 'active',
    createdAt: new Date().toISOString()
  };
  users.push(newUser);
  await saveUsers(users);
  await auditLog('USER_CREATED', req.session.userId, { ip: req.ip, msg: `新增使用者: ${username}` });
  res.json({ id: newUser.id, username: newUser.username, displayName: newUser.displayName, role: newUser.role });
});

app.put('/api/users/:id', requireAdmin, async (req, res) => {
  const users = await loadUsers();
  const user = users.find(u => u.id === req.params.id);
  if (!user) return res.status(404).json({ error: '使用者不存在' });

  const { username, password, displayName, role, status } = req.body;
  if (username) user.username = username;
  if (password) {
    const pwErr = validatePassword(password);
    if (pwErr) return res.status(400).json({ error: pwErr });
    user.password = bcrypt.hashSync(password, BCRYPT_ROUNDS);
  }
  if (displayName) user.displayName = displayName;
  if (role) user.role = role;
  if (status) user.status = status;

  await saveUsers(users);
  await auditLog('USER_UPDATED', req.session.userId, { ip: req.ip, msg: `修改使用者: ${user.username}` });
  res.json({ id: user.id, username: user.username, displayName: user.displayName, role: user.role, status: user.status });
});

app.delete('/api/users/:id', requireAdmin, async (req, res) => {
  let users = await loadUsers();
  const idx = users.findIndex(u => u.id === req.params.id);
  if (idx === -1) return res.status(404).json({ error: '使用者不存在' });
  if (users[idx].username === '蕭輝哲') return res.status(400).json({ error: '不能刪除系統管理員' });

  const deletedName = users[idx].username;
  users.splice(idx, 1);
  await saveUsers(users);
  await auditLog('USER_DELETED', req.session.userId, { ip: req.ip, msg: `刪除使用者: ${deletedName}` });
  res.json({ ok: true });
});

app.put('/api/change-password', requireAuth, async (req, res) => {
  const { oldPassword, newPassword } = req.body;
  if (!oldPassword || !newPassword) return res.status(400).json({ error: '請輸入舊密碼和新密碼' });

  const pwErr = validatePassword(newPassword);
  if (pwErr) return res.status(400).json({ error: pwErr });

  const users = await loadUsers();
  const user = users.find(u => u.id === req.session.userId);
  if (!user || !verifyPassword(oldPassword, user.password)) {
    return res.status(400).json({ error: '舊密碼錯誤' });
  }

  user.password = bcrypt.hashSync(newPassword, BCRYPT_ROUNDS);
  user.mustChangePassword = false;
  await saveUsers(users);
  await auditLog('PASSWORD_CHANGED', req.session.userId, { ip: req.ip });
  res.json({ ok: true });
});

// ===== 資料同步 API（取代前端直接寫 Firestore）=====
const DATA_COLLECTIONS = ['members', 'cases', 'opinions', 'services', 'billings'];

// 讀取所有資料
app.get('/api/data', requireAuth, async (req, res) => {
  try {
    if (!firestoreDb) return res.status(503).json({ error: '資料庫未就緒' });
    const result = {};
    for (const col of DATA_COLLECTIONS) {
      const snapshot = await firestoreDb.collection(col).get();
      result[col] = snapshot.docs.map(doc => doc.data());
    }
    const metaDoc = await firestoreDb.collection('system').doc('meta').get();
    result.dataVersion = metaDoc.exists ? (metaDoc.data().dataVersion || null) : null;
    await auditLog('DATA_READ', req.session.userId, { ip: req.ip, msg: `讀取資料: ${result.cases?.length || 0} 個案` });
    res.json(result);
  } catch (err) {
    console.error('讀取資料失敗:', err);
    res.status(500).json({ error: '讀取資料失敗' });
  }
});

// 儲存單一集合的項目
app.put('/api/data/:collection/:id', requireAuth, async (req, res) => {
  const { collection, id } = req.params;
  if (!DATA_COLLECTIONS.includes(collection)) return res.status(400).json({ error: '無效的集合' });
  if (!id) return res.status(400).json({ error: '缺少 ID' });
  try {
    if (!firestoreDb) return res.status(503).json({ error: '資料庫未就緒' });
    await firestoreDb.collection(collection).doc(id).set(req.body, { merge: true });
    res.json({ ok: true });
  } catch (err) {
    res.status(500).json({ error: '儲存失敗' });
  }
});

// 批次上傳所有資料
app.post('/api/data/sync', requireAuth, async (req, res) => {
  try {
    if (!firestoreDb) return res.status(503).json({ error: '資料庫未就緒' });
    const data = req.body;
    let totalItems = 0;
    for (const col of DATA_COLLECTIONS) {
      const items = data[col];
      if (!items || !Array.isArray(items)) continue;
      const snapshot = await firestoreDb.collection(col).get();
      for (let i = 0; i < snapshot.docs.length; i += 400) {
        const batch = firestoreDb.batch();
        snapshot.docs.slice(i, i + 400).forEach(doc => batch.delete(doc.ref));
        await batch.commit();
      }
      for (let i = 0; i < items.length; i += 400) {
        const batch = firestoreDb.batch();
        items.slice(i, i + 400).forEach(item => {
          if (item.id) batch.set(firestoreDb.collection(col).doc(item.id), item);
        });
        await batch.commit();
      }
      totalItems += items.length;
    }
    if (data.dataVersion) {
      await firestoreDb.collection('system').doc('meta').set({
        initialized: true,
        dataVersion: data.dataVersion,
        updatedAt: admin.firestore.FieldValue.serverTimestamp()
      });
    }
    await auditLog('DATA_SYNC', req.session.userId, { ip: req.ip, msg: `同步 ${totalItems} 筆資料` });
    res.json({ ok: true, totalItems });
  } catch (err) {
    console.error('資料同步失敗:', err);
    res.status(500).json({ error: '資料同步失敗' });
  }
});

// ===== LCMS 同步 API =====
app.post('/api/lcms-sync', requireAuth, upload.single('file'), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: '請上傳 .xls 檔案' });
    }

    const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });

    if (rows.length === 0) {
      return res.status(400).json({ error: '檔案中無資料' });
    }

    const lcmsCases = rows.map(row => {
      const cmsRaw = String(row['CMS'] || '');
      const cmsMatch = cmsRaw.match(/(\d+)/);
      const cmsLevel = cmsMatch ? parseInt(cmsMatch[1]) : null;

      const ageRaw = String(row['年齡'] || '');
      const ageMatch = ageRaw.match(/(\d+)/);
      const age = ageMatch ? parseInt(ageMatch[1]) : null;

      const welfareRaw = String(row['福利身分'] || '');
      let category = '';
      if (welfareRaw.includes('第一類')) category = '第一類';
      else if (welfareRaw.includes('第二類')) category = '第二類';
      else if (welfareRaw.includes('第三類')) category = '第三類';

      function rocToAD(rocDate) {
        if (!rocDate) return '';
        const str = String(rocDate).trim();
        const m = str.match(/^(\d{2,3})\/(\d{1,2})\/(\d{1,2})$/);
        if (!m) return '';
        const year = parseInt(m[1]) + 1911;
        const month = m[2].padStart(2, '0');
        const day = m[3].padStart(2, '0');
        return `${year}-${month}-${day}`;
      }

      const statusRaw = String(row['案件狀態'] || '');
      let status = 'active';
      if (statusRaw.includes('結案') || statusRaw.includes('終止')) {
        status = 'closed';
      }

      return {
        caseNo: String(row['案號'] || '').trim(),
        lcmsStatus: statusRaw.trim(),
        name: String(row['姓名'] || '').trim(),
        idNumber: String(row['身分證號'] || '').trim().toUpperCase(),
        birthday: rocToAD(row['出生日期']),
        age: age,
        cmsLevel: cmsLevel,
        category: category,
        welfareRaw: welfareRaw.trim(),
        district: String(row['居住地(行政區)'] || '').trim(),
        village: String(row['居住地(村里)'] || '').trim(),
        registeredCity: String(row['戶籍地(縣市)'] || '').trim(),
        registeredDistrict: String(row['戶籍地(行政區)'] || '').trim(),
        livingCity: String(row['居住地(縣市)'] || '').trim(),
        careCenter: String(row['照管中心'] || '').trim(),
        careManager: String(row['照管專員'] || '').trim(),
        unitName: String(row['A單位名稱'] || '').trim(),
        lcmsOpinionCount: parseInt(row['意見書數量(當年度)']) || 0,
        lcmsBillingCount: parseInt(row['申報紀錄數量(當年度)']) || 0,
        hospital: String(row['主責居家醫師院所'] || '').trim(),
        enrollDate: rocToAD(row['派案日期']),
        homeVisitDates: String(row['家訪日期'] || '').trim(),
        status: status
      };
    }).filter(c => c.idNumber);

    auditLog('LCMS_SYNC', req.session.userId, { ip: req.ip, msg: `同步 ${lcmsCases.length} 筆個案` });

    res.json({
      total: lcmsCases.length,
      cases: lcmsCases,
      columns: Object.keys(rows[0] || {})
    });
  } catch (err) {
    console.error('LCMS 解析失敗:', err);
    res.status(500).json({ error: 'LCMS 檔案解析失敗，請確認檔案格式正確' });
  }
});

// cases-data.js 不在 git 中（含個資），提供空白 fallback
app.get('/cases-data.js', (req, res) => {
  res.type('application/javascript').send('const RAW_CASES_DATA = [];');
});

// SPA fallback
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
