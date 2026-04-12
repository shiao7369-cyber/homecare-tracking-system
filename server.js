const express = require('express');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');
const multer = require('multer');
const XLSX = require('xlsx');
const { Pool } = require('pg');
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
      scriptSrc: ["'self'", "'unsafe-inline'", "https://www.gstatic.com", "https://cdn.jsdelivr.net", "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/", "https://unpkg.com"],
      scriptSrcAttr: ["'unsafe-inline'"],
      styleSrc: ["'self'", "'unsafe-inline'", "https://fonts.googleapis.com", "https://cdn.jsdelivr.net", "https://unpkg.com"],
      fontSrc: ["'self'", "https://fonts.gstatic.com", "https://cdn.jsdelivr.net"],
      connectSrc: ["'self'", "https://nominatim.openstreetmap.org"],
      imgSrc: ["'self'", "data:", "https://*.tile.openstreetmap.org", "https://*.basemaps.cartocdn.com"],
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
  index: false,
  setHeaders: (res, filePath) => {
    if (filePath.endsWith('.js') || filePath.endsWith('.css') || filePath.endsWith('.html')) {
      res.setHeader('Cache-Control', 'no-cache, must-revalidate');
    }
  }
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

// ===== PostgreSQL 初始化 =====
const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: process.env.DATABASE_URL?.includes('railway') ? { rejectUnauthorized: false } : false
});

async function initDatabase() {
  const client = await pool.connect();
  try {
    await client.query(`
      CREATE TABLE IF NOT EXISTS users (
        id TEXT PRIMARY KEY,
        username TEXT UNIQUE NOT NULL,
        password TEXT NOT NULL,
        "displayName" TEXT,
        role TEXT DEFAULT 'user',
        status TEXT DEFAULT 'active',
        "createdAt" TEXT,
        "mustChangePassword" BOOLEAN DEFAULT false
      );

      CREATE TABLE IF NOT EXISTS collections (
        collection TEXT NOT NULL,
        doc_id TEXT NOT NULL,
        data JSONB NOT NULL,
        PRIMARY KEY (collection, doc_id)
      );

      CREATE TABLE IF NOT EXISTS audit_logs (
        id SERIAL PRIMARY KEY,
        action TEXT,
        "userId" TEXT,
        timestamp TEXT,
        ip TEXT,
        details TEXT
      );

      CREATE TABLE IF NOT EXISTS system_meta (
        key TEXT PRIMARY KEY,
        value TEXT
      );
    `);
    console.log('PostgreSQL 資料庫初始化成功');

    // 初始化預設管理員
    const { rows } = await client.query('SELECT COUNT(*) as cnt FROM users');
    if (parseInt(rows[0].cnt) === 0) {
      // 嘗試從 users.json 遷移
      const localPath = path.join(__dirname, 'users.json');
      if (fs.existsSync(localPath)) {
        try {
          const jsonUsers = JSON.parse(fs.readFileSync(localPath, 'utf8'));
          for (const u of jsonUsers) {
            await client.query(
              `INSERT INTO users (id, username, password, "displayName", role, status, "createdAt", "mustChangePassword")
               VALUES ($1, $2, $3, $4, $5, $6, $7, $8) ON CONFLICT (id) DO NOTHING`,
              [u.id, u.username, u.password, u.displayName || u.username,
               u.role || 'user', u.status || 'active', u.createdAt || new Date().toISOString(),
               !!u.mustChangePassword]
            );
          }
          console.log(`從 users.json 遷移了 ${jsonUsers.length} 位使用者`);
          return;
        } catch (e) { console.error('users.json 遷移失敗:', e.message); }
      }
      // 建立預設管理員
      await client.query(
        `INSERT INTO users (id, username, password, "displayName", role, status, "createdAt", "mustChangePassword")
         VALUES ($1, $2, $3, $4, $5, $6, $7, $8)`,
        [crypto.randomUUID(), '蕭輝哲', bcrypt.hashSync('Hc@2025!Admin', 10),
         '蕭輝哲', 'admin', 'active', new Date().toISOString(), true]
      );
      console.log('已建立預設管理員帳號');
    }
  } finally {
    client.release();
  }
}

// ===== 稽核日誌 =====
async function auditLog(action, userId, details = {}) {
  const timestamp = new Date().toISOString();
  const ip = details.ip || '';
  const msg = details.msg || '';
  console.log(`[AUDIT] ${action} by ${userId || 'anonymous'} - ${msg}`);
  try {
    await pool.query(
      'INSERT INTO audit_logs (action, "userId", timestamp, ip, details) VALUES ($1, $2, $3, $4, $5)',
      [action, userId || 'anonymous', timestamp, ip, msg]
    );
  } catch (e) { /* 稽核寫入失敗不影響主流程 */ }
}

// ===== 使用者資料（PostgreSQL 持久化）=====
const BCRYPT_ROUNDS = 10;

function validatePassword(pw) {
  if (pw.length < 8) return '密碼長度至少 8 個字元';
  if (!/[A-Za-z]/.test(pw) || !/[0-9]/.test(pw)) return '密碼須包含英文字母和數字';
  return null;
}

function verifyPassword(inputPw, storedHash) {
  if (storedHash.startsWith('$2')) {
    return bcrypt.compareSync(inputPw, storedHash);
  }
  if (storedHash.length === 64) {
    const sha256 = crypto.createHash('sha256').update(inputPw).digest('hex');
    return sha256 === storedHash;
  }
  return false;
}

async function migratePasswordIfNeeded(user, inputPw) {
  if (!user.password.startsWith('$2')) {
    const newHash = bcrypt.hashSync(inputPw, BCRYPT_ROUNDS);
    await pool.query('UPDATE users SET password = $1 WHERE id = $2', [newHash, user.id]);
  }
}

async function loadUsers() {
  const { rows } = await pool.query('SELECT * FROM users');
  return rows;
}

async function saveUser(user) {
  await pool.query(
    `INSERT INTO users (id, username, password, "displayName", role, status, "createdAt", "mustChangePassword")
     VALUES ($1, $2, $3, $4, $5, $6, $7, $8)
     ON CONFLICT (id) DO UPDATE SET
       username = EXCLUDED.username, password = EXCLUDED.password,
       "displayName" = EXCLUDED."displayName", role = EXCLUDED.role,
       status = EXCLUDED.status, "mustChangePassword" = EXCLUDED."mustChangePassword"`,
    [user.id, user.username, user.password, user.displayName || user.username,
     user.role || 'user', user.status || 'active', user.createdAt || new Date().toISOString(),
     !!user.mustChangePassword]
  );
}

// ===== Session 管理（含過期機制）=====
const sessions = {};
const SESSION_MAX_AGE = 8 * 60 * 60 * 1000; // 8 小時
const SESSION_IDLE_TIMEOUT = 30 * 60 * 1000; // 30 分鐘閒置

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

// 登入
app.post('/api/login', loginLimiter, async (req, res) => {
  const { username, password } = req.body;
  if (!username || !password) return res.status(400).json({ error: '請輸入帳號和密碼' });

  const users = await loadUsers();
  const user = users.find(u => u.username === username && u.status === 'active');
  if (!user || !verifyPassword(password, user.password)) {
    auditLog('LOGIN_FAILED', username, { ip: req.ip, msg: '帳號或密碼錯誤' });
    return res.status(401).json({ error: '帳號或密碼錯誤' });
  }

  await migratePasswordIfNeeded(user, password);

  const token = createSession(user);
  auditLog('LOGIN_SUCCESS', user.id, { ip: req.ip, msg: user.username });
  res.json({
    token,
    user: { id: user.id, username: user.username, displayName: user.displayName, role: user.role },
    mustChangePassword: !!user.mustChangePassword,
    sessionMaxAge: SESSION_MAX_AGE,
    sessionIdleTimeout: SESSION_IDLE_TIMEOUT
  });
});

// SSO 單一登入（從 community-med 跳轉）
app.post('/api/sso', async (req, res) => {
  const { token } = req.body;
  if (!token) return res.status(400).json({ error: '缺少 SSO token' });

  const ssoSecret = process.env.SSO_SECRET?.trim();
  if (!ssoSecret) return res.status(500).json({ error: 'SSO 未設定' });

  try {
    // Token format: base64url(JSON payload).signature
    const [payloadB64, signature] = token.split('.');
    if (!payloadB64 || !signature) return res.status(401).json({ error: 'Token 格式錯誤' });

    // Verify HMAC signature
    const expectedSig = crypto.createHmac('sha256', ssoSecret).update(payloadB64).digest('base64url');
    if (signature !== expectedSig) return res.status(401).json({ error: 'Token 簽名無效' });

    // Decode payload
    const payload = JSON.parse(Buffer.from(payloadB64, 'base64url').toString());

    // Check expiry (30 seconds)
    if (Date.now() - payload.iat > 30000) return res.status(401).json({ error: 'Token 已過期' });

    // Find or auto-create user by username
    const users = await loadUsers();
    let user = users.find(u => u.username === payload.username && u.status === 'active');
    if (!user) {
      // Auto-create user from SSO payload
      user = {
        id: crypto.randomUUID(),
        username: payload.username,
        password: bcrypt.hashSync(crypto.randomUUID(), BCRYPT_ROUNDS),
        displayName: payload.displayName || payload.username,
        role: payload.role || 'user',
        status: 'active',
        createdAt: new Date().toISOString()
      };
      await saveUser(user);
      auditLog('SSO_USER_CREATED', user.id, { ip: req.ip, msg: `SSO auto-created: ${user.username}` });
    }

    // Create session (same as login)
    const sessionToken = createSession(user);
    auditLog('SSO_LOGIN', user.id, { ip: req.ip, msg: `SSO from community-med: ${user.username}` });

    res.json({
      token: sessionToken,
      user: { id: user.id, username: user.username, displayName: user.displayName, role: user.role },
      sessionMaxAge: SESSION_MAX_AGE,
      sessionIdleTimeout: SESSION_IDLE_TIMEOUT
    });
  } catch (err) {
    res.status(401).json({ error: 'Token 驗證失敗' });
  }
});

// 免登入自動取得 token（開發/內部用途）
app.post('/api/auto-login', async (req, res) => {
  const users = await loadUsers();
  const admin = users.find(u => u.role === 'admin' && u.status === 'active') || users.find(u => u.status === 'active');
  if (!admin) return res.status(500).json({ error: '無可用帳號' });
  const token = createSession(admin);
  auditLog('AUTO_LOGIN', admin.id, { ip: req.ip, msg: '免登入自動登入' });
  res.json({
    token,
    user: { id: admin.id, username: admin.username, displayName: admin.displayName, role: admin.role },
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
  await saveUser(newUser);
  auditLog('USER_CREATED', req.session.userId, { ip: req.ip, msg: `新增使用者: ${username}` });
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

  await saveUser(user);
  auditLog('USER_UPDATED', req.session.userId, { ip: req.ip, msg: `修改使用者: ${user.username}` });
  res.json({ id: user.id, username: user.username, displayName: user.displayName, role: user.role, status: user.status });
});

app.delete('/api/users/:id', requireAdmin, async (req, res) => {
  const users = await loadUsers();
  const idx = users.findIndex(u => u.id === req.params.id);
  if (idx === -1) return res.status(404).json({ error: '使用者不存在' });
  if (users[idx].username === '蕭輝哲') return res.status(400).json({ error: '不能刪除系統管理員' });

  const deletedName = users[idx].username;
  await pool.query('DELETE FROM users WHERE id = $1', [users[idx].id]);
  auditLog('USER_DELETED', req.session.userId, { ip: req.ip, msg: `刪除使用者: ${deletedName}` });
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
  await saveUser(user);
  auditLog('PASSWORD_CHANGED', req.session.userId, { ip: req.ip });
  res.json({ ok: true });
});

// ===== 資料同步 API（PostgreSQL）=====
const DATA_COLLECTIONS = ['members', 'cases', 'opinions', 'services', 'billings'];

// 讀取所有資料
app.get('/api/data', requireAuth, async (req, res) => {
  try {
    const result = {};
    for (const col of DATA_COLLECTIONS) {
      const { rows } = await pool.query('SELECT data FROM collections WHERE collection = $1', [col]);
      result[col] = rows.map(r => r.data);
    }
    const { rows: metaRows } = await pool.query('SELECT value FROM system_meta WHERE key = $1', ['dataVersion']);
    result.dataVersion = metaRows.length > 0 ? metaRows[0].value : null;
    auditLog('DATA_READ', req.session.userId, { ip: req.ip, msg: `讀取資料: ${result.cases?.length || 0} 個案` });
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
    // merge: 使用 JSONB || 合併
    await pool.query(
      `INSERT INTO collections (collection, doc_id, data) VALUES ($1, $2, $3::jsonb)
       ON CONFLICT (collection, doc_id) DO UPDATE SET data = collections.data || $3::jsonb`,
      [collection, id, JSON.stringify(req.body)]
    );
    res.json({ ok: true });
  } catch (err) {
    console.error('儲存失敗:', err);
    res.status(500).json({ error: '儲存失敗' });
  }
});

// 批次上傳所有資料
app.post('/api/data/sync', requireAuth, async (req, res) => {
  const client = await pool.connect();
  try {
    const data = req.body;
    let totalItems = 0;

    await client.query('BEGIN');
    for (const col of DATA_COLLECTIONS) {
      const items = data[col];
      if (!items || !Array.isArray(items)) continue;
      await client.query('DELETE FROM collections WHERE collection = $1', [col]);
      for (const item of items) {
        if (item.id) {
          await client.query(
            'INSERT INTO collections (collection, doc_id, data) VALUES ($1, $2, $3::jsonb)',
            [col, item.id, JSON.stringify(item)]
          );
        }
      }
      totalItems += items.length;
    }
    if (data.dataVersion) {
      await client.query(
        'INSERT INTO system_meta (key, value) VALUES ($1, $2) ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value',
        ['dataVersion', data.dataVersion]
      );
      await client.query(
        'INSERT INTO system_meta (key, value) VALUES ($1, $2) ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value',
        ['updatedAt', new Date().toISOString()]
      );
    }
    await client.query('COMMIT');

    auditLog('DATA_SYNC', req.session.userId, { ip: req.ip, msg: `同步 ${totalItems} 筆資料` });
    res.json({ ok: true, totalItems });
  } catch (err) {
    await client.query('ROLLBACK');
    console.error('資料同步失敗:', err);
    res.status(500).json({ error: '資料同步失敗' });
  } finally {
    client.release();
  }
});

// ===== LCMS 同步 API =====
app.post('/api/lcms-sync', requireAuth, upload.single('file'), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: '請上傳 .xls 檔案' });
    }

    console.log('[LCMS] file received:', req.file.originalname, 'size:', req.file.size);

    const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });

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

    function excelDateToISO(serial) {
      if (!serial || typeof serial !== 'number') return '';
      const d = new Date((serial - 25569) * 86400000);
      return d.toISOString().slice(0, 10);
    }

    // --- 自動偵測格式 ---
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    const headerRow = (aoa[0] || []).map(h => String(h).replace(/[\r\n\s]/g, ''));

    // 清除換行的 header 用於比對
    const hasLCMSColumns = headerRow.includes('身分證號') && headerRow.includes('案號');
    const hasLocalColumns = headerRow.some(h => h.includes('身份證字號')) || headerRow.includes('照管案號');

    console.log('[LCMS] detected format:', hasLCMSColumns ? 'LCMS' : hasLocalColumns ? 'LOCAL' : 'UNKNOWN');

    let lcmsCases = [];

    if (hasLCMSColumns) {
      // ===== LCMS 長照平台匯出格式 =====
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
      lcmsCases = rows.map(row => {
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

        const statusRaw = String(row['案件狀態'] || '');
        let status = 'active';
        if (statusRaw.includes('結案') || statusRaw.includes('終止')) status = 'closed';

        return {
          caseNo: String(row['案號'] || '').trim(),
          lcmsStatus: statusRaw.trim(),
          name: String(row['姓名'] || '').trim(),
          idNumber: String(row['身分證號'] || '').trim().toUpperCase(),
          birthday: rocToAD(row['出生日期']),
          age, cmsLevel, category,
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
          status
        };
      }).filter(c => c.idNumber);

    } else if (hasLocalColumns) {
      // ===== 本地個案清冊格式（雙行 header，欄位有換行符） =====
      // 用 column index 讀取，跳過前 2 行 header
      // col: 0序號 1照管案號 2敏盛案號 3姓名 4身份證字號 5性別 6身分別 7CMS 8地址 9里別 10主要聯絡人 11連絡電話 12照會日期 13主責醫師 14個管師家訪 15醫師家訪 16時效

      // 收案個案
      for (let i = 2; i < aoa.length; i++) {
        const r = aoa[i];
        const name = String(r[3] || '').trim();
        const idNumber = String(r[4] || '').trim().toUpperCase();
        if (!name || !idNumber || idNumber.length < 8) continue;

        const cmsRaw = String(r[7] || '');
        const cmsMatch = cmsRaw.toString().match(/(\d+)/);
        const cmsLevel = cmsMatch ? parseInt(cmsMatch[1]) : null;

        const catRaw = String(r[6] || '').trim();
        let category = '';
        if (catRaw.includes('第一類') || catRaw.includes('一般戶')) category = '第一類';
        else if (catRaw.includes('第二類') || catRaw.includes('中低')) category = '第二類';
        else if (catRaw.includes('第三類') || catRaw.includes('低收')) category = '第三類';
        else category = catRaw;

        // 照會日期可能是 Excel serial number
        let enrollDate = '';
        if (typeof r[12] === 'number') enrollDate = excelDateToISO(r[12]);
        else enrollDate = rocToAD(r[12]);

        // 醫師家訪日期
        let doctorVisitDate = '';
        if (typeof r[15] === 'number') doctorVisitDate = excelDateToISO(r[15]);
        else doctorVisitDate = rocToAD(r[15]);

        // 地址拆出行政區
        const addr = String(r[8] || '').trim();
        let district = '';
        const distMatch = addr.match(/([\u4e00-\u9fff]+[區鄉鎮市])/);
        if (distMatch) district = distMatch[1];

        lcmsCases.push({
          caseNo: String(r[1] || '').trim(),
          name, idNumber,
          gender: String(r[5] || '').trim(),
          cmsLevel, category,
          address: addr,
          district,
          village: String(r[9] || '').trim(),
          contactPerson: String(r[10] || '').replace(/[\r\n]/g, ' ').trim(),
          phone: String(r[11] || '').replace(/[\r\n]/g, ' ').trim(),
          doctorName: String(r[13] || '').trim(),
          enrollDate,
          doctorVisitDate,
          status: 'active'
        });
      }

      // 如果有「結案總表」sheet，也讀入標記為 closed
      if (workbook.SheetNames.includes('結案總表')) {
        const closedSheet = workbook.Sheets['結案總表'];
        const closedAoa = XLSX.utils.sheet_to_json(closedSheet, { header: 1, defval: '' });
        for (let i = 2; i < closedAoa.length; i++) {
          const r = closedAoa[i];
          const name = String(r[3] || '').trim();
          const idNumber = String(r[4] || '').trim().toUpperCase();
          if (!name || !idNumber || idNumber.length < 8) continue;

          // 檢查是否已在收案清冊中
          const existing = lcmsCases.find(c => c.idNumber === idNumber);
          if (existing) {
            existing.status = 'closed';
          } else {
            const cmsRaw = String(r[7] || '');
            const cmsMatch = cmsRaw.toString().match(/(\d+)/);
            const cmsLevel = cmsMatch ? parseInt(cmsMatch[1]) : null;

            lcmsCases.push({
              caseNo: String(r[1] || '').trim(),
              name, idNumber,
              gender: String(r[5] || '').trim(),
              cmsLevel,
              category: String(r[6] || '').trim(),
              address: String(r[8] || '').trim(),
              village: String(r[9] || '').trim(),
              doctorName: String(r[13] || '').trim(),
              status: 'closed'
            });
          }
        }
      }

    } else {
      // 未知格式，嘗試用 sheet_to_json 的 key 搜尋
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
      if (rows.length === 0) return res.status(400).json({ error: '檔案中無資料' });
      console.log('[LCMS] unknown format, columns:', Object.keys(rows[0]));
      return res.status(400).json({
        error: '無法辨識檔案格式，請使用 LCMS 匯出的個案清冊或本地管理清冊',
        columns: Object.keys(rows[0])
      });
    }

    console.log('[LCMS] parsed cases:', lcmsCases.length);
    auditLog('LCMS_SYNC', req.session?.userId, { ip: req.ip, msg: `同步 ${lcmsCases.length} 筆個案` });

    res.json({
      total: lcmsCases.length,
      cases: lcmsCases,
      format: hasLCMSColumns ? 'lcms' : 'local',
      columns: headerRow.filter(h => h)
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

// ===== 地理編碼代理（多 API Fallback + 行政區驗證） =====
const geoHeaders = { 'User-Agent': 'HomecareTrackingSystem/1.0', 'Accept-Language': 'zh-TW' };

// 桃園各行政區中心座標（用於驗證定位結果）
const DISTRICT_CENTERS = {
  '桃園區':[24.9936,121.3010],'中壢區':[24.9656,121.2249],'平鎮區':[24.9457,121.2183],
  '八德區':[24.9527,121.2855],'楊梅區':[24.9077,121.1455],'蘆竹區':[25.0457,121.2920],
  '龜山區':[25.0335,121.3457],'龍潭區':[24.8635,121.2165],'大溪區':[24.8833,121.2873],
  '大園區':[25.0629,121.1975],'觀音區':[25.0335,121.0835],'新屋區':[24.9721,121.1062],
  '復興區':[24.8208,121.3530]
};

// 驗證定位結果是否在正確行政區（容許半徑 ~5km）
function validateGeoResult(data, expectedDist) {
  if (!data || data.length === 0 || !expectedDist) return data;
  const center = DISTRICT_CENTERS[expectedDist];
  if (!center) return data; // 無法驗證的行政區，放行

  const lat = parseFloat(data[0].lat);
  const lon = parseFloat(data[0].lon);
  if (isNaN(lat) || isNaN(lon)) return [];

  // 計算與預期行政區中心的距離（度數，1度≈111km）
  const dLat = lat - center[0];
  const dLon = lon - center[1];
  const distDeg = Math.sqrt(dLat * dLat + dLon * dLon);
  const distKm = distDeg * 111;

  // 行政區半徑約 5km，給寬容值到 8km
  if (distKm > 8) {
    console.log(`[Geocode] 驗證失敗: "${expectedDist}" 預期 [${center}]，結果 [${lat},${lon}]，距離 ${distKm.toFixed(1)}km — 拒絕`);
    return []; // 回傳空，讓下一策略接手
  }
  return data;
}

// 清理地址（共用）
function cleanGeoAddr(raw) {
  let addr = raw.replace(/[０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0));
  addr = addr.replace(/[Ａ-Ｚａ-ｚ]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0));
  addr = addr.replace(/[一二三四五六七八九十\d]+樓.*$/, '').replace(/\d+[fF].*$/, '');
  addr = addr.replace(/之\d+/g, '').replace(/(號).*$/, '$1');
  addr = addr.replace(/([區鄉鎮市])[\u4e00-\u9fff]{1,5}里(\d{1,3}鄰)?/g, '$1');
  addr = addr.replace(/[\u4e00-\u9fff]{1,5}里(\d{1,3}鄰)?$/, '');
  return addr.trim();
}

// API 1: TGOS 國土測繪中心（台灣地址最精準）
async function tgosGeocode(addr) {
  try {
    const qs = new URLSearchParams({ fuzzyType: '2', queryAddr: addr, isLockCounty: 'N', isFullMatch: 'N' });
    const resp = await fetch(`https://addr.tgos.tw/addrws/v40/QueryAddr.asmx/QueryAddr?${qs}`, {
      headers: { ...geoHeaders, 'Accept': 'application/json' }, signal: AbortSignal.timeout(8000)
    });
    const json = await resp.json();
    if (json && json.AddressList && json.AddressList.length > 0) {
      const item = json.AddressList[0];
      if (item.X && item.Y) {
        return [{ lat: String(item.Y), lon: String(item.X), display_name: item.FULL_ADDR || addr, source: 'tgos' }];
      }
    }
  } catch (e) { /* TGOS 失敗，繼續下一個 */ }
  return [];
}

// API 2: NLSC 內政部地籍門牌（備用台灣定位）
async function nlscGeocode(addr) {
  try {
    const qs = new URLSearchParams({ word: addr });
    const resp = await fetch(`https://api.nlsc.gov.tw/other/TownVillagePointQuery/${encodeURIComponent(addr)}`, {
      headers: geoHeaders, signal: AbortSignal.timeout(8000)
    });
    const text = await resp.text();
    // NLSC 回傳 XML，解析 <coorX> <coorY>
    const xMatch = text.match(/<coorX>([\d.]+)<\/coorX>/);
    const yMatch = text.match(/<coorY>([\d.]+)<\/coorY>/);
    if (xMatch && yMatch) {
      return [{ lat: yMatch[1], lon: xMatch[1], display_name: addr, source: 'nlsc' }];
    }
  } catch (e) { /* NLSC 失敗 */ }
  return [];
}

// API 3: Nominatim（OpenStreetMap）
async function nominatimSearch(params) {
  const qs = new URLSearchParams({ ...params, format: 'json', limit: '3', countrycodes: 'tw' });
  const resp = await fetch(`https://nominatim.openstreetmap.org/search?${qs}`, { headers: geoHeaders });
  return resp.json();
}

// 從 Nominatim 多筆結果中挑選最佳的（最靠近預期行政區）
function pickBestNominatim(results, expectedDist) {
  if (!results || results.length === 0) return [];
  if (!expectedDist || !DISTRICT_CENTERS[expectedDist]) return [results[0]];

  const center = DISTRICT_CENTERS[expectedDist];
  let best = null, bestDist = Infinity;
  for (const r of results) {
    const lat = parseFloat(r.lat), lon = parseFloat(r.lon);
    if (isNaN(lat) || isNaN(lon)) continue;
    const d = Math.sqrt((lat - center[0]) ** 2 + (lon - center[1]) ** 2);
    if (d < bestDist) { bestDist = d; best = r; }
  }
  return best ? [best] : [results[0]];
}

app.get('/api/geocode', requireAuth, async (req, res) => {
  const raw = req.query.q;
  if (!raw) return res.status(400).json({ error: '缺少地址參數' });
  try {
    const addr = cleanGeoAddr(raw);
    const cityMatch = addr.match(/([\u4e00-\u9fff]+[市縣])/);
    const distMatch = addr.match(/[市縣]([\u4e00-\u9fff]+[區鄉鎮市])/);
    const streetMatch = addr.match(/[區鄉鎮市]([\u4e00-\u9fff]+[路街道大][\u4e00-\u9fff]*段?\d*[\u4e00-\u9fff]*)/);
    const city = cityMatch ? cityMatch[1] : '';
    const dist = distMatch ? distMatch[1] : '';

    let data = [];

    // === 策略1: TGOS 門牌定位（最精準）===
    data = validateGeoResult(await tgosGeocode(addr), dist);
    if (data.length > 0) return res.json(data);

    // 簡化地址再試 TGOS（去巷弄號）
    if (streetMatch) {
      const simpleAddr = (city || '桃園市') + (dist || '') + streetMatch[1];
      data = validateGeoResult(await tgosGeocode(simpleAddr), dist);
      if (data.length > 0) return res.json(data);
    }

    // === 策略2: Nominatim 結構化查詢 ===
    if (streetMatch && city) {
      const results = await nominatimSearch({ street: streetMatch[1], city: city + (dist || '') });
      data = validateGeoResult(pickBestNominatim(results, dist), dist);
      if (data.length > 0) return res.json(data);
    }

    // === 策略3: Nominatim 全文查詢（取多筆選最佳）===
    {
      const results = await nominatimSearch({ q: addr });
      data = validateGeoResult(pickBestNominatim(results, dist), dist);
      if (data.length > 0) return res.json(data);
    }

    // === 策略4: 去掉巷弄號只用路街 ===
    {
      const roadOnly = addr.replace(/\d+巷.*/, '').replace(/\d+弄.*/, '').replace(/\d+號.*/, '').trim();
      if (roadOnly.length > 4) {
        const results = await nominatimSearch({ q: roadOnly });
        data = validateGeoResult(pickBestNominatim(results, dist), dist);
        if (data.length > 0) return res.json(data);
      }
    }

    // === 策略5: 只用路名 + 行政區 ===
    if (streetMatch && dist) {
      const roadAddr = (city || '桃園市') + dist + streetMatch[1].replace(/\d+.*/, '');
      const results = await nominatimSearch({ q: roadAddr });
      data = validateGeoResult(pickBestNominatim(results, dist), dist);
      if (data.length > 0) return res.json(data);
    }

    // === 策略6: 只用行政區（粗略定位，保底）===
    if (city && dist) {
      data = await nominatimSearch({ q: city + dist });
      // 行政區查詢不驗證（本身就是粗略定位）
      if (data.length > 0) return res.json(pickBestNominatim(data, dist));
    }

    res.json(data);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// OSRM 路線規劃代理
app.get('/api/route', requireAuth, async (req, res) => {
  const { coords } = req.query; // "lng1,lat1;lng2,lat2;..."
  if (!coords) return res.status(400).json({ error: '缺少座標參數' });
  try {
    const url = `https://router.project-osrm.org/route/v1/driving/${coords}?overview=full&geometries=geojson&steps=true`;
    const resp = await fetch(url, { headers: geoHeaders });
    const data = await resp.json();
    res.json(data);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// SPA fallback
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

// ===== 啟動伺服器 =====
initDatabase()
  .then(() => {
    app.listen(PORT, () => {
      console.log(`Server running on port ${PORT}`);
    });
  })
  .catch(err => {
    console.error('資料庫初始化失敗:', err);
    process.exit(1);
  });
