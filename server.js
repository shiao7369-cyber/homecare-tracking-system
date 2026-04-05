const express = require('express');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');
const multer = require('multer');
const XLSX = require('xlsx');
const admin = require('firebase-admin');
const app = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 10 * 1024 * 1024 } });
const PORT = process.env.PORT || 3000;

app.use(express.json());
app.use(express.static(path.join(__dirname)));

// ===== Firebase Admin 初始化 =====
let firestoreDb = null;

function initFirebase() {
  try {
    if (process.env.FIREBASE_SERVICE_ACCOUNT) {
      const serviceAccount = JSON.parse(process.env.FIREBASE_SERVICE_ACCOUNT);
      admin.initializeApp({
        credential: admin.credential.cert(serviceAccount)
      });
    } else if (fs.existsSync(path.join(__dirname, 'serviceAccountKey.json'))) {
      const serviceAccount = require('./serviceAccountKey.json');
      admin.initializeApp({
        credential: admin.credential.cert(serviceAccount)
      });
    } else {
      admin.initializeApp({
        projectId: 'homecare-system-f37ea'
      });
    }
    firestoreDb = admin.firestore();
    console.log('Firebase Admin 初始化成功');
  } catch (err) {
    console.error('Firebase Admin 初始化失敗:', err.message);
  }
}

initFirebase();

// ===== 使用者資料（Firestore 持久化）=====
const USERS_COLLECTION = 'users';
let _usersCache = null;

function getDefaultUsers() {
  return [
    {
      id: 'U001',
      username: '蕭輝哲',
      password: hashPassword('蕭輝哲'),
      displayName: '蕭輝哲',
      role: 'admin',
      status: 'active',
      createdAt: new Date().toISOString()
    }
  ];
}

async function loadUsers() {
  if (_usersCache) return _usersCache;

  if (firestoreDb) {
    try {
      const snapshot = await firestoreDb.collection(USERS_COLLECTION).get();
      if (!snapshot.empty) {
        _usersCache = snapshot.docs.map(doc => doc.data());
        console.log(`從 Firestore 載入 ${_usersCache.length} 位使用者`);
        return _usersCache;
      }
    } catch (err) {
      console.error('Firestore 讀取使用者失敗:', err.message);
    }
  }

  // Fallback: 本地檔案
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
      // 先刪除所有舊的
      const snapshot = await firestoreDb.collection(USERS_COLLECTION).get();
      snapshot.docs.forEach(doc => batch.delete(doc.ref));
      // 寫入所有使用者
      users.forEach(user => {
        const ref = firestoreDb.collection(USERS_COLLECTION).doc(user.id);
        batch.set(ref, user);
      });
      await batch.commit();
      console.log(`使用者資料已存到 Firestore (${users.length} 位)`);
      return;
    } catch (err) {
      console.error('Firestore 寫入使用者失敗:', err.message);
    }
  }

  // Fallback: 本地檔案
  try {
    fs.writeFileSync(path.join(__dirname, 'users.json'), JSON.stringify(users, null, 2), 'utf8');
  } catch (e) {
    console.log('警告：使用者資料僅存於記憶體');
  }
}

function hashPassword(pw) {
  return crypto.createHash('sha256').update(pw).digest('hex');
}

// 簡易 session 管理
const sessions = {};

function createSession(user) {
  const token = crypto.randomBytes(32).toString('hex');
  sessions[token] = {
    userId: user.id,
    username: user.username,
    displayName: user.displayName,
    role: user.role,
    createdAt: Date.now()
  };
  return token;
}

function getSession(req) {
  const token = req.headers['authorization']?.replace('Bearer ', '');
  return token ? sessions[token] : null;
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
app.post('/api/login', async (req, res) => {
  const { username, password } = req.body;
  if (!username || !password) return res.status(400).json({ error: '請輸入帳號和密碼' });

  const users = await loadUsers();
  console.log(`登入嘗試: "${username}", 使用者數: ${users.length}, 帳號列表: ${users.map(u=>u.username).join(',')}`);
  const user = users.find(u => u.username === username && u.status === 'active');
  if (!user) {
    console.log(`找不到使用者: "${username}"`);
    return res.status(401).json({ error: '帳號或密碼錯誤' });
  }
  if (user.password !== hashPassword(password)) {
    console.log(`密碼不符: stored=${user.password.substring(0,8)}..., input=${hashPassword(password).substring(0,8)}...`);
    return res.status(401).json({ error: '帳號或密碼錯誤' });
  }

  const token = createSession(user);
  res.json({
    token,
    user: { id: user.id, username: user.username, displayName: user.displayName, role: user.role }
  });
});

// 登出
app.post('/api/logout', (req, res) => {
  const token = req.headers['authorization']?.replace('Bearer ', '');
  if (token) delete sessions[token];
  res.json({ ok: true });
});

// 取得目前使用者
app.get('/api/me', requireAuth, (req, res) => {
  res.json(req.session);
});

// ===== 使用者管理 (管理員) =====

// 列出所有使用者
app.get('/api/users', requireAdmin, async (req, res) => {
  const users = (await loadUsers()).map(u => ({
    id: u.id, username: u.username, displayName: u.displayName,
    role: u.role, status: u.status, createdAt: u.createdAt
  }));
  res.json(users);
});

// 新增使用者
app.post('/api/users', requireAdmin, async (req, res) => {
  const { username, password, displayName, role } = req.body;
  if (!username || !password) return res.status(400).json({ error: '帳號和密碼為必填' });

  const users = await loadUsers();
  if (users.find(u => u.username === username)) {
    return res.status(400).json({ error: '此帳號已存在' });
  }

  const newUser = {
    id: 'U' + String(users.length + 1).padStart(3, '0'),
    username,
    password: hashPassword(password),
    displayName: displayName || username,
    role: role || 'user',
    status: 'active',
    createdAt: new Date().toISOString()
  };
  users.push(newUser);
  await saveUsers(users);
  res.json({ id: newUser.id, username: newUser.username, displayName: newUser.displayName, role: newUser.role });
});

// 修改使用者
app.put('/api/users/:id', requireAdmin, async (req, res) => {
  const users = await loadUsers();
  const user = users.find(u => u.id === req.params.id);
  if (!user) return res.status(404).json({ error: '使用者不存在' });

  const { username, password, displayName, role, status } = req.body;
  if (username) user.username = username;
  if (password) user.password = hashPassword(password);
  if (displayName) user.displayName = displayName;
  if (role) user.role = role;
  if (status) user.status = status;

  await saveUsers(users);
  res.json({ id: user.id, username: user.username, displayName: user.displayName, role: user.role, status: user.status });
});

// 刪除使用者
app.delete('/api/users/:id', requireAdmin, async (req, res) => {
  let users = await loadUsers();
  const idx = users.findIndex(u => u.id === req.params.id);
  if (idx === -1) return res.status(404).json({ error: '使用者不存在' });
  if (users[idx].username === '蕭輝哲') return res.status(400).json({ error: '不能刪除系統管理員' });

  users.splice(idx, 1);
  await saveUsers(users);
  res.json({ ok: true });
});

// 修改密碼（自己改自己的）
app.put('/api/change-password', requireAuth, async (req, res) => {
  const { oldPassword, newPassword } = req.body;
  if (!oldPassword || !newPassword) return res.status(400).json({ error: '請輸入舊密碼和新密碼' });

  const users = await loadUsers();
  const user = users.find(u => u.id === req.session.userId);
  if (!user || user.password !== hashPassword(oldPassword)) {
    return res.status(400).json({ error: '舊密碼錯誤' });
  }

  user.password = hashPassword(newPassword);
  await saveUsers(users);
  res.json({ ok: true });
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

    // Map LCMS columns to system fields
    const lcmsCases = rows.map(row => {
      // Extract CMS level number from "6級" format
      const cmsRaw = String(row['CMS'] || '');
      const cmsMatch = cmsRaw.match(/(\d+)/);
      const cmsLevel = cmsMatch ? parseInt(cmsMatch[1]) : null;

      // Extract age number from "74歲" format
      const ageRaw = String(row['年齡'] || '');
      const ageMatch = ageRaw.match(/(\d+)/);
      const age = ageMatch ? parseInt(ageMatch[1]) : null;

      // Map welfare category
      const welfareRaw = String(row['福利身分'] || '');
      let category = '';
      if (welfareRaw.includes('第一類')) category = '第一類';
      else if (welfareRaw.includes('第二類')) category = '第二類';
      else if (welfareRaw.includes('第三類')) category = '第三類';

      // Convert ROC date to AD date
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

      // Case status mapping
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
    }).filter(c => c.idNumber); // Filter out rows without ID number

    res.json({
      total: lcmsCases.length,
      cases: lcmsCases,
      columns: Object.keys(rows[0] || {})
    });
  } catch (err) {
    console.error('LCMS 解析失敗:', err);
    res.status(500).json({ error: 'LCMS 檔案解析失敗: ' + err.message });
  }
});

// SPA fallback
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
