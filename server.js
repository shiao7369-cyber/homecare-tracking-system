const express = require('express');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');
const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json());
app.use(express.static(path.join(__dirname)));

// ===== 使用者資料檔 =====
const USERS_FILE = path.join(__dirname, 'users.json');

function loadUsers() {
  if (fs.existsSync(USERS_FILE)) {
    try { return JSON.parse(fs.readFileSync(USERS_FILE, 'utf8')); } catch (e) { }
  }
  // 預設管理員帳號
  const defaultUsers = [
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
  saveUsers(defaultUsers);
  return defaultUsers;
}

function saveUsers(users) {
  fs.writeFileSync(USERS_FILE, JSON.stringify(users, null, 2), 'utf8');
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
app.post('/api/login', (req, res) => {
  const { username, password } = req.body;
  if (!username || !password) return res.status(400).json({ error: '請輸入帳號和密碼' });

  const users = loadUsers();
  const user = users.find(u => u.username === username && u.status === 'active');
  if (!user || user.password !== hashPassword(password)) {
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
app.get('/api/users', requireAdmin, (req, res) => {
  const users = loadUsers().map(u => ({
    id: u.id, username: u.username, displayName: u.displayName,
    role: u.role, status: u.status, createdAt: u.createdAt
  }));
  res.json(users);
});

// 新增使用者
app.post('/api/users', requireAdmin, (req, res) => {
  const { username, password, displayName, role } = req.body;
  if (!username || !password) return res.status(400).json({ error: '帳號和密碼為必填' });

  const users = loadUsers();
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
  saveUsers(users);
  res.json({ id: newUser.id, username: newUser.username, displayName: newUser.displayName, role: newUser.role });
});

// 修改使用者
app.put('/api/users/:id', requireAdmin, (req, res) => {
  const users = loadUsers();
  const user = users.find(u => u.id === req.params.id);
  if (!user) return res.status(404).json({ error: '使用者不存在' });

  const { username, password, displayName, role, status } = req.body;
  if (username) user.username = username;
  if (password) user.password = hashPassword(password);
  if (displayName) user.displayName = displayName;
  if (role) user.role = role;
  if (status) user.status = status;

  saveUsers(users);
  res.json({ id: user.id, username: user.username, displayName: user.displayName, role: user.role, status: user.status });
});

// 刪除使用者
app.delete('/api/users/:id', requireAdmin, (req, res) => {
  let users = loadUsers();
  const idx = users.findIndex(u => u.id === req.params.id);
  if (idx === -1) return res.status(404).json({ error: '使用者不存在' });
  if (users[idx].username === '蕭輝哲') return res.status(400).json({ error: '不能刪除系統管理員' });

  users.splice(idx, 1);
  saveUsers(users);
  res.json({ ok: true });
});

// 修改密碼（自己改自己的）
app.put('/api/change-password', requireAuth, (req, res) => {
  const { oldPassword, newPassword } = req.body;
  if (!oldPassword || !newPassword) return res.status(400).json({ error: '請輸入舊密碼和新密碼' });

  const users = loadUsers();
  const user = users.find(u => u.id === req.session.userId);
  if (!user || user.password !== hashPassword(oldPassword)) {
    return res.status(400).json({ error: '舊密碼錯誤' });
  }

  user.password = hashPassword(newPassword);
  saveUsers(users);
  res.json({ ok: true });
});

// SPA fallback
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
