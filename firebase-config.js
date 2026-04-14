/* ========================================
   認證與雲端資料層
   - 登入/登出：自訂後台 API
   - 資料同步：Firebase Firestore
   ======================================== */


// ===== Session 管理 =====
let currentUser = null;
let authToken = null;

// 安全讀取 token（檢查儲存時間，超過 8 小時自動清除）
(function loadToken() {
  const stored = sessionStorage.getItem('auth_token');
  const storedAt = parseInt(sessionStorage.getItem('auth_token_time') || '0');
  if (stored && (Date.now() - storedAt < 8 * 60 * 60 * 1000)) {
    authToken = stored;
  } else {
    sessionStorage.removeItem('auth_token');
    sessionStorage.removeItem('auth_token_time');
    // 遷移舊 localStorage token
    const old = localStorage.getItem('auth_token');
    if (old) {
      sessionStorage.removeItem('auth_token'); sessionStorage.removeItem('auth_token_time');
    }
  }
})();

async function apiCall(method, url, body, _retried) {
  const opts = { method, headers: { 'Content-Type': 'application/json' } };
  if (authToken) opts.headers['Authorization'] = 'Bearer ' + authToken;
  if (body) opts.body = JSON.stringify(body);
  const res = await fetch(url, opts);
  const data = await res.json();
  if (res.status === 401 && !_retried && url !== '/api/auto-login') {
    // Session 過期，嘗試自動重新登入後重試
    try {
      const loginData = await apiCall('POST', '/api/auto-login', null, true);
      authToken = loginData.token;
      currentUser = loginData.user;
      sessionStorage.setItem('auth_token', authToken);
      sessionStorage.setItem('auth_token_time', String(Date.now()));
      return apiCall(method, url, body, true);
    } catch (e) {
      throw new Error(data.error || '請求失敗');
    }
  }
  if (!res.ok) throw new Error(data.error || '請求失敗');
  return data;
}

// ===== 登入/登出 =====
async function doLogin() {
  const username = document.getElementById('login-username').value.trim();
  const pw = document.getElementById('login-password').value;
  if (!username || !pw) { showLoginError('請輸入帳號和密碼'); return; }

  try {
    const data = await apiCall('POST', '/api/login', { username, password: pw });
    authToken = data.token;
    currentUser = data.user;
    sessionStorage.setItem('auth_token', authToken);
    sessionStorage.setItem('auth_token_time', String(Date.now()));
    onLoginSuccess(data);
  } catch (err) {
    showLoginError(err.message);
  }
}

function doLogout() {
  stopIdleDetection();
  apiCall('POST', '/api/logout').catch(() => {});
  authToken = null;
  currentUser = null;
  sessionStorage.removeItem('auth_token'); sessionStorage.removeItem('auth_token_time');
  document.getElementById('login-page').style.display = 'flex';
  document.getElementById('sidebar').style.display = 'none';
  document.getElementById('main-content').style.display = 'none';
  document.getElementById('login-password').value = '';
}

function showLoginError(msg) {
  const el = document.getElementById('login-error');
  el.textContent = msg;
  el.style.display = 'block';
}

// ===== 閒置偵測與 Session 過期警告 =====
let _idleTimer = null;
let _idleWarningTimer = null;
let _sessionMaxAge = 8 * 60 * 60 * 1000;
let _sessionIdleTimeout = 30 * 60 * 1000;

function resetIdleTimer() {
  clearTimeout(_idleTimer);
  clearTimeout(_idleWarningTimer);
  const warningEl = document.getElementById('session-warning');
  if (warningEl) warningEl.style.display = 'none';
  _idleWarningTimer = setTimeout(() => {
    const w = document.getElementById('session-warning');
    if (w) w.style.display = 'block';
  }, _sessionIdleTimeout - 5 * 60 * 1000);
  _idleTimer = setTimeout(() => {
    if (authToken) { alert('已閒置超過 30 分鐘，系統將自動登出。'); doLogout(); }
  }, _sessionIdleTimeout);
}

function startIdleDetection() {
  ['click', 'keydown', 'mousemove', 'scroll', 'touchstart'].forEach(evt => {
    document.addEventListener(evt, resetIdleTimer, { passive: true });
  });
  resetIdleTimer();
}

function stopIdleDetection() {
  clearTimeout(_idleTimer);
  clearTimeout(_idleWarningTimer);
  ['click', 'keydown', 'mousemove', 'scroll', 'touchstart'].forEach(evt => {
    document.removeEventListener(evt, resetIdleTimer);
  });
}

function onLoginSuccess(loginResponse) {
  if (loginResponse) {
    if (loginResponse.sessionMaxAge) _sessionMaxAge = loginResponse.sessionMaxAge;
    if (loginResponse.sessionIdleTimeout) _sessionIdleTimeout = loginResponse.sessionIdleTimeout;
  }
  document.getElementById('login-page').style.display = 'none';
  document.getElementById('sidebar').style.display = 'flex';
  document.getElementById('main-content').style.display = 'flex';
  document.getElementById('login-error').style.display = 'none';
  const nameEl = document.querySelector('.user-name');
  if (nameEl) nameEl.textContent = currentUser.displayName;
  const roleEl = document.querySelector('.user-role');
  if (roleEl) roleEl.textContent = currentUser.role === 'admin' ? '系統管理員' : '一般使用者';
  const adminNav = document.getElementById('nav-users');
  if (adminNav) adminNav.style.display = currentUser.role === 'admin' ? '' : 'none';
  if (loginResponse && loginResponse.mustChangePassword) {
    showForceChangePasswordModal();
    return;
  }
  startIdleDetection();
  loadCloudData().then(() => {
    initNav(); initFilters(); renderDashboard(); updateAlertBadge();
    document.getElementById('today-date').textContent = formatDateCN(new Date());
  });
}

// ===== 強制更改密碼 =====
function showForceChangePasswordModal() {
  openModal('首次登入 — 請更改預設密碼', `
    <p style="margin-bottom:1rem;color:var(--danger)">為確保帳號安全，首次登入必須更改密碼。</p>
    <div class="form-group"><label class="form-label">目前密碼</label><input class="form-input" id="f-force-old-pw" type="password"></div>
    <div class="form-group"><label class="form-label">新密碼（至少 8 字元，含英文和數字）</label><input class="form-input" id="f-force-new-pw" type="password"></div>
    <div class="form-group"><label class="form-label">確認新密碼</label><input class="form-input" id="f-force-confirm-pw" type="password"></div>
  `, `<button class="btn btn-primary" onclick="doForceChangePassword()">更改密碼</button>`);
}

async function doForceChangePassword() {
  const oldPw = document.getElementById('f-force-old-pw').value;
  const newPw = document.getElementById('f-force-new-pw').value;
  const confirmPw = document.getElementById('f-force-confirm-pw').value;
  if (!oldPw || !newPw) { alert('請填寫所有欄位'); return; }
  if (newPw !== confirmPw) { alert('新密碼與確認密碼不一致'); return; }
  if (newPw.length < 8 || !/[A-Za-z]/.test(newPw) || !/[0-9]/.test(newPw)) { alert('密碼須至少 8 字元且包含英文和數字'); return; }
  try {
    await apiCall('PUT', '/api/change-password', { oldPassword: oldPw, newPassword: newPw });
    closeModal(); showToast('密碼已更改'); startIdleDetection();
    loadCloudData().then(() => { initNav(); initFilters(); renderDashboard(); updateAlertBadge(); document.getElementById('today-date').textContent = formatDateCN(new Date()); });
  } catch (err) { alert(err.message); }
}

// ===== 使用者自行修改密碼 =====
function showChangePasswordModal() {
  openModal('修改密碼', `
    <div class="form-group"><label class="form-label">目前密碼</label><input class="form-input" id="f-chg-old-pw" type="password"></div>
    <div class="form-group"><label class="form-label">新密碼（至少 8 字元，含英文和數字）</label><input class="form-input" id="f-chg-new-pw" type="password"></div>
    <div class="form-group"><label class="form-label">確認新密碼</label><input class="form-input" id="f-chg-confirm-pw" type="password"></div>
  `, `<button class="btn btn-outline" onclick="closeModal()">取消</button><button class="btn btn-primary" onclick="doChangeMyPassword()">儲存</button>`);
}

async function doChangeMyPassword() {
  const oldPw = document.getElementById('f-chg-old-pw').value;
  const newPw = document.getElementById('f-chg-new-pw').value;
  const confirmPw = document.getElementById('f-chg-confirm-pw').value;
  if (!oldPw || !newPw) { alert('請填寫所有欄位'); return; }
  if (newPw !== confirmPw) { alert('新密碼與確認密碼不一致'); return; }
  try {
    await apiCall('PUT', '/api/change-password', { oldPassword: oldPw, newPassword: newPw });
    closeModal(); showToast('密碼已更改成功');
  } catch (err) { alert(err.message); }
}

// ===== SSO 單一登入（從 community-med 跳轉）=====
async function handleSSO() {
  const params = new URLSearchParams(window.location.search);
  const ssoToken = params.get('sso');
  if (!ssoToken) return false;

  // 清除 URL 中的 sso 參數
  const cleanUrl = window.location.pathname;
  window.history.replaceState({}, '', cleanUrl);

  // 清除舊 session，確保用 SSO 使用者登入
  sessionStorage.removeItem('auth_token');
  sessionStorage.removeItem('auth_token_time');
  authToken = null;

  try {
    const data = await apiCall('POST', '/api/sso', { token: ssoToken });
    authToken = data.token;
    currentUser = data.user;
    sessionStorage.setItem('auth_token', authToken);
    sessionStorage.setItem('auth_token_time', String(Date.now()));
    onLoginSuccess(data);
    return true;
  } catch (e) {
    console.error('SSO 登入失敗:', e.message);
    alert('SSO 登入失敗: ' + e.message);
    return false;
  }
}

// ===== 頁面載入時直接進入（跳過登入） =====
document.addEventListener('DOMContentLoaded', async () => {
  // 先嘗試 SSO
  if (await handleSSO()) return;

  // 再嘗試已存在的 session
  if (authToken) {
    try {
      currentUser = await apiCall('GET', '/api/me');
      onLoginSuccess();
      return;
    } catch (e) {
      sessionStorage.removeItem('auth_token'); sessionStorage.removeItem('auth_token_time');
      authToken = null;
    }
  }

  // 無 token 時自動登入（跳過登入畫面）
  try {
    const data = await apiCall('POST', '/api/auto-login');
    authToken = data.token;
    currentUser = data.user;
    sessionStorage.setItem('auth_token', authToken);
    sessionStorage.setItem('auth_token_time', String(Date.now()));
    onLoginSuccess(data);
  } catch (err) {
    console.error('自動登入失敗', err);
  }
});

// ===== 雲端資料讀寫 =====
const COLLECTIONS = ['members', 'cases', 'opinions', 'services', 'billings', 'schedules'];

async function loadCloudData() {
  try {
    const cloudData = await apiCall('GET', '/api/data');
    if (cloudData.cases && cloudData.cases.length > 0) {
      // 伺服器有資料，直接使用
      const _members = cloudData.members || [];
      // 自動修正醫師專科
      if (typeof DOCTOR_SPECIALTY_MAP !== 'undefined') {
        _members.forEach(m => {
          if (DOCTOR_SPECIALTY_MAP[m.name] && m.specialty !== DOCTOR_SPECIALTY_MAP[m.name]) {
            m.specialty = DOCTOR_SPECIALTY_MAP[m.name];
          }
        });
      }
      db = { members: _members, cases: cloudData.cases || [],
             opinions: cloudData.opinions || [], services: cloudData.services || [],
             billings: cloudData.billings || [], schedules: cloudData.schedules || [], diseases: [] };
      // 行程資料雙向同步
      if (db.schedules && db.schedules.length > 0) {
        // 雲端有行程 → 寫入 localStorage 供 getScheduleData() 讀取
        localStorage.setItem('homecare_schedule', JSON.stringify(db.schedules));
      } else {
        // 雲端無行程 → 從 localStorage 補回，下次 saveDB 時會推到雲端
        try {
          const localSchedules = JSON.parse(localStorage.getItem('homecare_schedule') || '[]');
          if (localSchedules.length > 0) db.schedules = localSchedules;
        } catch(e) {}
      }
      // 合併本地 geocache 座標：如果雲端個案沒有座標但本地快取有，回寫
      try {
        const geoCache = JSON.parse(localStorage.getItem('geocache') || '{}');
        if (db.cases && Object.keys(geoCache).length > 0) {
          db.cases.forEach(c => {
            if (c.lat && c.lng) return; // 已有座標不覆蓋
            // 嘗試用地址找快取
            let addr = '';
            if (c.address) {
              addr = c.address.replace(/[０-９]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0))
                .replace(/[一二三四五六七八九十\d]+樓.*$/, '').replace(/\d+[fF].*$/, '')
                .replace(/([區鄉鎮市])[^\s路街道巷弄號]{1,5}里(\d{1,3}鄰)?/g, '$1')
                .replace(/[\u4e00-\u9fff]{1,5}里(\d{1,3}鄰)?$/, '')
                .replace(/之\d+/, '').replace(/(號).*$/, '$1').trim();
            } else if (c.district) {
              addr = '桃園市' + c.district;
            }
            if (addr && geoCache[addr]) {
              c.lat = geoCache[addr][0];
              c.lng = geoCache[addr][1];
            }
          });
        }
      } catch(e) { /* geocache 合併失敗不影響主流程 */ }
    } else {
      // 伺服器無資料，嘗試從 localStorage 快取上傳
      const cached = localStorage.getItem(DB_KEY);
      if (cached) {
        try {
          db = JSON.parse(cached);
          if (db.cases && db.cases.length > 0) {
            console.log('伺服器無資料，從本地快取上傳...');
            syncDataToBackend(db).catch(e => console.error('自動上傳失敗:', e));
            return;
          }
        } catch (e) { /* 快取損壞 */ }
      }
      // 都沒有，產生初始資料
      if (typeof RAW_CASES_DATA !== 'undefined' && RAW_CASES_DATA.length > 0) {
        db = generateDemoData();
        syncDataToBackend(db).catch(e => console.error('雲端同步失敗:', e));
      } else {
        db = generateDemoData();
      }
    }
    saveDB_local(db);
  } catch (err) {
    console.error('雲端載入失敗，使用本地快取:', err);
    const cached = localStorage.getItem(DB_KEY);
    if (cached) {
      try { db = JSON.parse(cached); return; } catch (e) { /* 快取損壞 */ }
    }
    db = generateDemoData(); saveDB_local(db);
  }
}

async function syncDataToBackend(data) {
  const payload = { dataVersion: DB_VERSION };
  COLLECTIONS.forEach(col => { if (data[col]) payload[col] = data[col]; });
  await apiCall('POST', '/api/data/sync', payload);
}

function saveDB_local(data) {
  localStorage.setItem(DB_KEY, JSON.stringify(data));
  localStorage.setItem(DB_VERSION_KEY, DB_VERSION);
  if (data.members) saveMembers(data.members);
}

// saveDB - via backend API
let _saveDebounceTimer = null;
function saveDB(data) {
  localStorage.setItem(DB_KEY, JSON.stringify(data));
  if (data.members) saveMembers(data.members);
  clearTimeout(_saveDebounceTimer);
  _saveDebounceTimer = setTimeout(() => {
    COLLECTIONS.forEach(col => {
      if (data[col]) {
        data[col].forEach(item => {
          if (item.id) {
            apiCall('PUT', `/api/data/${col}/${item.id}`, item)
              .catch(err => console.error(`同步 ${col}/${item.id} 失敗:`, err));
          }
        });
      }
    });
  }, 500);
}

// ===== 手動同步到雲端 =====
async function syncToCloud() {
  const btn = document.getElementById('btn-sync-cloud');
  if (!currentUser) { alert('請先登入再同步'); return; }
  btn.disabled = true; btn.textContent = '☁ 同步中...';
  try {
    if (!db || !db.cases || db.cases.length === 0) { db = generateDemoData(); saveDB_local(db); }
    await syncDataToBackend(db);
    btn.textContent = '☁ 同步完成 ✓'; btn.style.background = '#2196F3';
    alert(`雲端同步完成！\n${db.cases.length} 個案、${db.members.length} 成員、${db.services.length} 服務紀錄已上傳`);
  } catch (err) {
    console.error('同步失敗:', err);
    btn.textContent = '☁ 同步失敗 ✗'; btn.style.background = '#f44336';
    alert('同步失敗: ' + err.message);
  } finally {
    setTimeout(() => { btn.disabled = false; btn.textContent = '☁ 同步到雲端'; btn.style.background = '#4CAF50'; }, 3000);
  }
}

// ===== 使用者管理 (管理員) =====
async function renderUserManagement() {
  const container = document.getElementById('page-users');
  if (!container) return;
  if (!currentUser || currentUser.role !== 'admin') {
    container.innerHTML = '<p>無權限存取</p>';
    return;
  }

  try {
    const users = await apiCall('GET', '/api/users');
    container.innerHTML = `
      <div class="card">
        <div class="card-header">
          <h3>使用者管理</h3>
          <button class="btn btn-sm btn-primary" onclick="openAddUserModal()">+ 新增使用者</button>
        </div>
        <table class="data-table">
          <thead>
            <tr>
              <th>帳號</th><th>顯示名稱</th><th>角色</th><th>狀態</th><th>建立時間</th><th>操作</th>
            </tr>
          </thead>
          <tbody>
            ${users.map(u => `
              <tr>
                <td><strong>${esc(u.username)}</strong></td>
                <td>${esc(u.displayName)}</td>
                <td><span class="badge ${u.role === 'admin' ? 'badge-danger' : 'badge-info'}">${u.role === 'admin' ? '管理員' : '使用者'}</span></td>
                <td><span class="badge ${u.status === 'active' ? 'badge-success' : 'badge-warning'}">${u.status === 'active' ? '啟用' : '停用'}</span></td>
                <td>${u.createdAt ? u.createdAt.substring(0, 10) : '-'}</td>
                <td>
                  <button class="btn btn-sm btn-outline" onclick="openEditUserModal('${esc(u.id)}','${esc(u.username)}','${esc(u.displayName)}','${esc(u.role)}','${esc(u.status)}')">編輯</button>
                  ${u.username !== '蕭輝哲' ? `<button class="btn btn-sm btn-outline" style="color:red;border-color:red" onclick="deleteUser('${esc(u.id)}','${esc(u.username)}')">刪除</button>` : ''}
                </td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
    `;
  } catch (err) {
    container.innerHTML = `<p style="color:red">載入失敗: ${esc(err.message)}</p>`;
  }
}

function openAddUserModal() {
  document.getElementById('modal-title').textContent = '新增使用者';
  document.getElementById('modal-body').innerHTML = `
    <div class="form-group">
      <label class="form-label">帳號 *</label>
      <input class="form-input" id="f-new-username" placeholder="登入帳號">
    </div>
    <div class="form-group">
      <label class="form-label">密碼 *</label>
      <input class="form-input" id="f-new-password" type="password" placeholder="登入密碼">
    </div>
    <div class="form-group">
      <label class="form-label">顯示名稱</label>
      <input class="form-input" id="f-new-displayname" placeholder="側欄顯示的名稱">
    </div>
    <div class="form-group">
      <label class="form-label">角色</label>
      <select class="form-input" id="f-new-role">
        <option value="user">一般使用者</option>
        <option value="admin">管理員</option>
      </select>
    </div>
  `;
  document.getElementById('modal-footer').innerHTML = `
    <button class="btn btn-outline" onclick="closeModal()">取消</button>
    <button class="btn btn-primary" onclick="addUser()">新增</button>
  `;
  document.getElementById('modal-overlay').classList.add('active');
}

async function addUser() {
  const username = document.getElementById('f-new-username').value.trim();
  const password = document.getElementById('f-new-password').value;
  const displayName = document.getElementById('f-new-displayname').value.trim() || username;
  const role = document.getElementById('f-new-role').value;

  if (!username || !password) { alert('帳號和密碼為必填'); return; }

  try {
    await apiCall('POST', '/api/users', { username, password, displayName, role });
    closeModal();
    renderUserManagement();
    showToast('使用者已新增');
  } catch (err) {
    alert(err.message);
  }
}

function openEditUserModal(id, username, displayName, role, status) {
  document.getElementById('modal-title').textContent = '編輯使用者 - ' + username;
  document.getElementById('modal-body').innerHTML = `
    <input type="hidden" id="f-edit-id" value="${esc(id)}">
    <div class="form-group">
      <label class="form-label">帳號</label>
      <input class="form-input" id="f-edit-username" value="${esc(username)}">
    </div>
    <div class="form-group">
      <label class="form-label">新密碼 (留空不修改，至少8字元含英數)</label>
      <input class="form-input" id="f-edit-password" type="password" placeholder="不修改請留空">
    </div>
    <div class="form-group">
      <label class="form-label">顯示名稱</label>
      <input class="form-input" id="f-edit-displayname" value="${esc(displayName)}">
    </div>
    <div class="form-group">
      <label class="form-label">角色</label>
      <select class="form-input" id="f-edit-role">
        <option value="user" ${role === 'user' ? 'selected' : ''}>一般使用者</option>
        <option value="admin" ${role === 'admin' ? 'selected' : ''}>管理員</option>
      </select>
    </div>
    <div class="form-group">
      <label class="form-label">狀態</label>
      <select class="form-input" id="f-edit-status">
        <option value="active" ${status === 'active' ? 'selected' : ''}>啟用</option>
        <option value="disabled" ${status !== 'active' ? 'selected' : ''}>停用</option>
      </select>
    </div>
  `;
  document.getElementById('modal-footer').innerHTML = `
    <button class="btn btn-outline" onclick="closeModal()">取消</button>
    <button class="btn btn-primary" onclick="updateUser()">儲存</button>
  `;
  document.getElementById('modal-overlay').classList.add('active');
}

async function updateUser() {
  const id = document.getElementById('f-edit-id').value;
  const body = {
    username: document.getElementById('f-edit-username').value.trim(),
    displayName: document.getElementById('f-edit-displayname').value.trim(),
    role: document.getElementById('f-edit-role').value,
    status: document.getElementById('f-edit-status').value,
  };
  const pw = document.getElementById('f-edit-password').value;
  if (pw) body.password = pw;

  try {
    await apiCall('PUT', '/api/users/' + id, body);
    closeModal();
    renderUserManagement();
    showToast('使用者已更新');
  } catch (err) {
    alert(err.message);
  }
}

async function deleteUser(id, username) {
  if (!confirm(`確定刪除使用者「${username}」？`)) return;
  try {
    await apiCall('DELETE', '/api/users/' + id);
    renderUserManagement();
    showToast('使用者已刪除');
  } catch (err) {
    alert(err.message);
  }
}

/* ===== LCMS 功能已移至 app.js 的統一上傳 =====
  // Create a hidden file input and trigger it
  const fileInput = document.createElement('input');
  fileInput.type = 'file';
  fileInput.accept = '.xls,.xlsx';
  fileInput.style.display = 'none';
  document.body.appendChild(fileInput);

  fileInput.addEventListener('change', async (e) => {
    const file = e.target.files[0];
    document.body.removeChild(fileInput);
    if (!file) return;

    // Upload and parse
    const formData = new FormData();
    formData.append('file', file);

    try {
      const res = await fetch('/api/lcms-sync', {
        method: 'POST',
        headers: authToken ? { 'Authorization': 'Bearer ' + authToken } : {},
        body: formData
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || '上傳失敗');

      // Compare with current db.cases
      const lcmsCases = data.cases;
      const diff = computeLCMSDiff(lcmsCases);
      showLCMSDiffModal(diff, lcmsCases);
    } catch (err) {
      alert('LCMS 同步失敗: ' + err.message);
    }
  });

  fileInput.click();
}

function computeLCMSDiff(lcmsCases) {
  const systemCases = (db && db.cases) ? db.cases : [];

  // Build lookup maps by idNumber (case-insensitive)
  const systemMap = {};
  systemCases.forEach(c => {
    if (c.idNumber) systemMap[c.idNumber.toUpperCase()] = c;
  });
  const lcmsMap = {};
  lcmsCases.forEach(c => {
    if (c.idNumber) lcmsMap[c.idNumber.toUpperCase()] = c;
  });

  const newCases = [];      // In LCMS but not in system
  const changed = [];       // In both but fields differ
  const possiblyClosed = []; // In system (active) but not in LCMS

  // Check LCMS cases against system
  lcmsCases.forEach(lc => {
    const key = lc.idNumber.toUpperCase();
    const sc = systemMap[key];
    if (!sc) {
      newCases.push(lc);
    } else {
      // Compare fields
      const diffs = [];
      if (lc.cmsLevel != null && sc.cmsLevel != null && String(lc.cmsLevel) !== String(sc.cmsLevel)) {
        diffs.push({ field: 'CMS等級', oldVal: sc.cmsLevel, newVal: lc.cmsLevel });
      }
      if (lc.category && sc.category && lc.category !== sc.category) {
        diffs.push({ field: '福利身分', oldVal: sc.category, newVal: lc.category });
      }
      if (lc.district && sc.district && lc.district !== sc.district) {
        diffs.push({ field: '行政區', oldVal: sc.district, newVal: lc.district });
      }
      if (lc.age != null && sc.age != null && String(lc.age) !== String(sc.age)) {
        diffs.push({ field: '年齡', oldVal: sc.age, newVal: lc.age });
      }
      if (lc.caseNo && sc.caseNo && lc.caseNo !== sc.caseNo) {
        diffs.push({ field: '案號', oldVal: sc.caseNo, newVal: lc.caseNo });
      }
      if (lc.enrollDate && sc.enrollDate && lc.enrollDate !== sc.enrollDate) {
        diffs.push({ field: '派案日期', oldVal: sc.enrollDate, newVal: lc.enrollDate });
      }
      if (lc.careManager && lc.careManager !== (sc.careManager || '')) {
        diffs.push({ field: '照管專員', oldVal: sc.careManager || '(空)', newVal: lc.careManager });
      }
      if (lc.village && lc.village !== (sc.village || '')) {
        diffs.push({ field: '村里', oldVal: sc.village || '(空)', newVal: lc.village });
      }
      if (lc.unitName && lc.unitName !== (sc.unitName || '')) {
        diffs.push({ field: 'A單位名稱', oldVal: sc.unitName || '(空)', newVal: lc.unitName });
      }
      // Always update LCMS-only fields if they have values
      if (lc.lcmsOpinionCount != null && lc.lcmsOpinionCount !== (sc.lcmsOpinionCount || 0)) {
        diffs.push({ field: '意見書數量(年度)', oldVal: sc.lcmsOpinionCount || 0, newVal: lc.lcmsOpinionCount });
      }
      if (lc.lcmsBillingCount != null && lc.lcmsBillingCount !== (sc.lcmsBillingCount || 0)) {
        diffs.push({ field: '申報紀錄數量(年度)', oldVal: sc.lcmsBillingCount || 0, newVal: lc.lcmsBillingCount });
      }
      if (lc.homeVisitDates && lc.homeVisitDates !== (sc.homeVisitDates || '')) {
        diffs.push({ field: '家訪日期', oldVal: sc.homeVisitDates || '(空)', newVal: lc.homeVisitDates });
      }
      if (diffs.length > 0) {
        changed.push({ lcms: lc, system: sc, diffs: diffs });
      }
    }
  });

  // Check system cases not in LCMS (possibly closed)
  systemCases.forEach(sc => {
    if (sc.status !== 'active') return;
    const key = (sc.idNumber || '').toUpperCase();
    if (key && !lcmsMap[key]) {
      possiblyClosed.push(sc);
    }
  });

  return {
    lcmsTotal: lcmsCases.length,
    systemTotal: systemCases.filter(c => c.status === 'active').length,
    newCases: newCases,
    changed: changed,
    possiblyClosed: possiblyClosed
  };
}

function showLCMSDiffModal(diff, lcmsCases) {
  const modalTitle = document.getElementById('modal-title');
  const modalBody = document.getElementById('modal-body');
  const modalFooter = document.getElementById('modal-footer');

  modalTitle.textContent = 'LCMS 同步報告';

  // Build summary
  let html = `
    <div style="max-height:70vh;overflow-y:auto;">
      <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:0.75rem;margin-bottom:1rem;">
        <div style="background:#e3f2fd;padding:0.75rem;border-radius:8px;text-align:center;">
          <div style="font-size:1.5rem;font-weight:bold;color:#1565c0;">${diff.lcmsTotal}</div>
          <div style="font-size:0.85rem;color:#555;">LCMS 個案數</div>
        </div>
        <div style="background:#f3e5f5;padding:0.75rem;border-radius:8px;text-align:center;">
          <div style="font-size:1.5rem;font-weight:bold;color:#7b1fa2;">${diff.systemTotal}</div>
          <div style="font-size:0.85rem;color:#555;">系統收案數</div>
        </div>
        <div style="background:#e8f5e9;padding:0.75rem;border-radius:8px;text-align:center;">
          <div style="font-size:1.5rem;font-weight:bold;color:#2e7d32;">${diff.newCases.length}</div>
          <div style="font-size:0.85rem;color:#555;">新個案</div>
        </div>
        <div style="background:#fff3e0;padding:0.75rem;border-radius:8px;text-align:center;">
          <div style="font-size:1.5rem;font-weight:bold;color:#e65100;">${diff.changed.length}</div>
          <div style="font-size:0.85rem;color:#555;">異動</div>
        </div>
        <div style="background:#fce4ec;padding:0.75rem;border-radius:8px;text-align:center;">
          <div style="font-size:1.5rem;font-weight:bold;color:#c62828;">${diff.possiblyClosed.length}</div>
          <div style="font-size:0.85rem;color:#555;">可能結案</div>
        </div>
      </div>
  `;

  // New cases section
  if (diff.newCases.length > 0) {
    html += `
      <details open style="margin-bottom:1rem;">
        <summary style="font-weight:bold;font-size:1rem;cursor:pointer;color:#2e7d32;margin-bottom:0.5rem;">
          新個案 (${diff.newCases.length})
        </summary>
        <div style="overflow-x:auto;">
          <table class="data-table" style="font-size:0.85rem;">
            <thead>
              <tr>
                <th><input type="checkbox" id="lcms-new-all" onchange="toggleLCMSCheckAll(this, 'lcms-new-chk')" checked></th>
                <th>案號</th><th>姓名</th><th>身分證號</th><th>CMS</th><th>身分別</th><th>行政區</th><th>派案日期</th>
              </tr>
            </thead>
            <tbody>
              ${diff.newCases.map((c, i) => `
                <tr>
                  <td><input type="checkbox" class="lcms-new-chk" data-idx="${i}" checked></td>
                  <td>${esc(c.caseNo)}</td>
                  <td>${esc(c.name)}</td>
                  <td>${maskId(c.idNumber)}</td>
                  <td>${c.cmsLevel || '-'}</td>
                  <td>${esc(c.category || '-')}</td>
                  <td>${esc(c.district || '-')}</td>
                  <td>${esc(c.enrollDate || '-')}</td>
                </tr>
              `).join('')}
            </tbody>
          </table>
        </div>
      </details>
    `;
  }

  // Changed cases section
  if (diff.changed.length > 0) {
    html += `
      <details open style="margin-bottom:1rem;">
        <summary style="font-weight:bold;font-size:1rem;cursor:pointer;color:#e65100;margin-bottom:0.5rem;">
          異動 (${diff.changed.length})
        </summary>
        <div style="overflow-x:auto;">
          <table class="data-table" style="font-size:0.85rem;">
            <thead>
              <tr>
                <th><input type="checkbox" id="lcms-chg-all" onchange="toggleLCMSCheckAll(this, 'lcms-chg-chk')" checked></th>
                <th>姓名</th><th>身分證號</th><th>異動欄位</th><th>原值</th><th>新值</th>
              </tr>
            </thead>
            <tbody>
              ${diff.changed.map((item, i) => item.diffs.map((d, j) => `
                <tr>
                  ${j === 0 ? `
                    <td rowspan="${item.diffs.length}"><input type="checkbox" class="lcms-chg-chk" data-idx="${i}" checked></td>
                    <td rowspan="${item.diffs.length}">${esc(item.lcms.name)}</td>
                    <td rowspan="${item.diffs.length}">${maskId(item.lcms.idNumber)}</td>
                  ` : ''}
                  <td><span style="color:#e65100;font-weight:500;">${esc(d.field)}</span></td>
                  <td style="color:#999;text-decoration:line-through;">${esc(d.oldVal)}</td>
                  <td style="color:#2e7d32;font-weight:500;">${esc(d.newVal)}</td>
                </tr>
              `).join('')).join('')}
            </tbody>
          </table>
        </div>
      </details>
    `;
  }

  // Possibly closed section
  if (diff.possiblyClosed.length > 0) {
    html += `
      <details style="margin-bottom:1rem;">
        <summary style="font-weight:bold;font-size:1rem;cursor:pointer;color:#c62828;margin-bottom:0.5rem;">
          可能結案 (${diff.possiblyClosed.length}) — 系統中收案但 LCMS 無資料
        </summary>
        <div style="overflow-x:auto;">
          <table class="data-table" style="font-size:0.85rem;">
            <thead>
              <tr>
                <th><input type="checkbox" id="lcms-close-all" onchange="toggleLCMSCheckAll(this, 'lcms-close-chk')"></th>
                <th>姓名</th><th>身分證號</th><th>CMS</th><th>負責醫師</th><th>收案日期</th>
              </tr>
            </thead>
            <tbody>
              ${diff.possiblyClosed.map((c, i) => `
                <tr>
                  <td><input type="checkbox" class="lcms-close-chk" data-idx="${i}"></td>
                  <td>${esc(c.name || '-')}</td>
                  <td>${maskId(c.idNumber)}</td>
                  <td>${c.cmsLevel || '-'}</td>
                  <td>${esc(c.doctorName || '-')}</td>
                  <td>${esc(c.enrollDate || '-')}</td>
                </tr>
              `).join('')}
            </tbody>
          </table>
        </div>
      </details>
    `;
  }

  if (diff.newCases.length === 0 && diff.changed.length === 0 && diff.possiblyClosed.length === 0) {
    html += '<p style="text-align:center;color:#666;padding:2rem;">資料完全一致，無需同步。</p>';
  }

  html += '</div>';
  modalBody.innerHTML = html;

  // Store diff data for apply functions
  window._lcmsDiff = diff;

  const hasChanges = diff.newCases.length > 0 || diff.changed.length > 0 || diff.possiblyClosed.length > 0;
  modalFooter.innerHTML = `
    <button class="btn btn-outline" onclick="closeModal()">關閉</button>
    ${hasChanges ? '<button class="btn btn-primary" onclick="applyLCMSSync()">套用勾選項目</button>' : ''}
  `;

  document.getElementById('modal-overlay').classList.add('active');

  // Widen modal for diff view
  const modal = document.getElementById('modal');
  modal.style.maxWidth = '900px';
  modal.style.width = '90vw';
}

function toggleLCMSCheckAll(masterCheckbox, className) {
  const checkboxes = document.querySelectorAll('.' + className);
  checkboxes.forEach(cb => { cb.checked = masterCheckbox.checked; });
}

function applyLCMSSync() {
  const diff = window._lcmsDiff;
  if (!diff) return;

  let addedCount = 0;
  let updatedCount = 0;
  let closedCount = 0;

  // Apply new cases
  const newCheckboxes = document.querySelectorAll('.lcms-new-chk:checked');
  newCheckboxes.forEach(cb => {
    const idx = parseInt(cb.dataset.idx);
    const lc = diff.newCases[idx];
    if (!lc) return;

    // Generate new case ID
    const maxId = db.cases.reduce((max, c) => {
      const num = parseInt((c.id || '').replace('C', ''));
      return num > max ? num : max;
    }, 0);
    const newId = 'C' + String(maxId + 1 + addedCount).padStart(3, '0');

    const newCase = {
      id: newId,
      caseNo: lc.caseNo || '',
      name: lc.name,
      idNumber: lc.idNumber,
      gender: '',
      cmsLevel: lc.cmsLevel,
      category: lc.category || '',
      district: lc.district || '',
      village: lc.village || '',
      address: '',
      status: 'active',
      doctorName: '',
      enrollDate: lc.enrollDate || '',
      age: lc.age,
      careManager: lc.careManager || '',
      unitName: lc.unitName || '',
      lcmsOpinionCount: lc.lcmsOpinionCount || 0,
      lcmsBillingCount: lc.lcmsBillingCount || 0,
      homeVisitDates: lc.homeVisitDates || '',
      // Default empty values for other standard fields
      phone: '',
      diagnosis: '',
      notes: '',
      opinionDate: '',
      opinionExpiry: '',
      acpStatus: 'not_started',
      adStatus: 'not_started',
      acpExplainDate: '',
      adExplainDate: '',
      acpSignDate: '',
      nhiCardDate: '',
      familyExplainDate: ''
    };

    db.cases.push(newCase);
    addedCount++;
  });

  // Apply changes
  const chgCheckboxes = document.querySelectorAll('.lcms-chg-chk:checked');
  chgCheckboxes.forEach(cb => {
    const idx = parseInt(cb.dataset.idx);
    const item = diff.changed[idx];
    if (!item) return;

    const sc = db.cases.find(c => c.idNumber && c.idNumber.toUpperCase() === item.lcms.idNumber.toUpperCase());
    if (!sc) return;

    // Apply each field diff
    item.diffs.forEach(d => {
      switch (d.field) {
        case 'CMS等級': sc.cmsLevel = item.lcms.cmsLevel; break;
        case '福利身分': sc.category = item.lcms.category; break;
        case '行政區': sc.district = item.lcms.district; break;
        case '年齡': sc.age = item.lcms.age; break;
        case '案號': sc.caseNo = item.lcms.caseNo; break;
        case '派案日期': sc.enrollDate = item.lcms.enrollDate; break;
        case '照管專員': sc.careManager = item.lcms.careManager; break;
        case '村里': sc.village = item.lcms.village; break;
        case 'A單位名稱': sc.unitName = item.lcms.unitName; break;
        case '意見書數量(年度)': sc.lcmsOpinionCount = item.lcms.lcmsOpinionCount; break;
        case '申報紀錄數量(年度)': sc.lcmsBillingCount = item.lcms.lcmsBillingCount; break;
        case '家訪日期': sc.homeVisitDates = item.lcms.homeVisitDates; break;
      }
    });
    updatedCount++;
  });

  // Apply closures
  const closeCheckboxes = document.querySelectorAll('.lcms-close-chk:checked');
  closeCheckboxes.forEach(cb => {
    const idx = parseInt(cb.dataset.idx);
    const sc = diff.possiblyClosed[idx];
    if (!sc) return;

    const target = db.cases.find(c => c.id === sc.id);
    if (target) {
      target.status = 'closed';
      closedCount++;
    }
  });

  // Save
  if (addedCount > 0 || updatedCount > 0 || closedCount > 0) {
    saveDB(db);
  }

  closeModal();

  // Reset modal width
  const modal = document.getElementById('modal');
  modal.style.maxWidth = '';
  modal.style.width = '';

  // Show result
  const msgs = [];
  if (addedCount > 0) msgs.push(`新增 ${addedCount} 個案`);
  if (updatedCount > 0) msgs.push(`更新 ${updatedCount} 個案`);
  if (closedCount > 0) msgs.push(`結案 ${closedCount} 個案`);

  if (msgs.length > 0) {
    alert('LCMS 同步完成：\n' + msgs.join('\n'));
    // Refresh UI
    if (typeof renderDashboard === 'function') renderDashboard();
    if (typeof renderCases === 'function') renderCases();
    if (typeof updateAlertBadge === 'function') updateAlertBadge();
  } else {
    alert('未選取任何項目，未做變更。');
  }
}
===== */
