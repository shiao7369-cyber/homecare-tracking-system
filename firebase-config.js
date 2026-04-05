/* ========================================
   認證與雲端資料層
   - 登入/登出：自訂後台 API
   - 資料同步：Firebase Firestore
   ======================================== */

const firebaseConfig = {
  apiKey: "AIzaSyA8-p3jb110_eML7tTwOhq09UouSDlR92o",
  authDomain: "homecare-system-f37ea.firebaseapp.com",
  projectId: "homecare-system-f37ea",
  storageBucket: "homecare-system-f37ea.firebasestorage.app",
  messagingSenderId: "501052920202",
  appId: "1:501052920202:web:c838641f2bed35f81e7873"
};

firebase.initializeApp(firebaseConfig);
const fsdb = firebase.firestore();

// ===== Session 管理 =====
let currentUser = null;
let authToken = localStorage.getItem('auth_token');

async function apiCall(method, url, body) {
  const opts = { method, headers: { 'Content-Type': 'application/json' } };
  if (authToken) opts.headers['Authorization'] = 'Bearer ' + authToken;
  if (body) opts.body = JSON.stringify(body);
  const res = await fetch(url, opts);
  const data = await res.json();
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
    localStorage.setItem('auth_token', authToken);
    onLoginSuccess();
  } catch (err) {
    showLoginError(err.message);
  }
}

function doLogout() {
  apiCall('POST', '/api/logout').catch(() => {});
  authToken = null;
  currentUser = null;
  localStorage.removeItem('auth_token');
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

function onLoginSuccess() {
  document.getElementById('login-page').style.display = 'none';
  document.getElementById('sidebar').style.display = 'flex';
  document.getElementById('main-content').style.display = 'flex';
  document.getElementById('login-error').style.display = 'none';

  // 更新使用者資訊
  const nameEl = document.querySelector('.user-name');
  if (nameEl) nameEl.textContent = currentUser.displayName;
  const roleEl = document.querySelector('.user-role');
  if (roleEl) roleEl.textContent = currentUser.role === 'admin' ? '系統管理員' : '一般使用者';

  // 管理員才顯示「使用者管理」
  const adminNav = document.getElementById('nav-users');
  if (adminNav) adminNav.style.display = currentUser.role === 'admin' ? '' : 'none';

  // 載入資料
  loadCloudData().then(() => {
    initNav();
    initFilters();
    renderDashboard();
    updateAlertBadge();
    document.getElementById('today-date').textContent = formatDateCN(new Date());
  });
}

// ===== 頁面載入時嘗試自動登入 =====
document.addEventListener('DOMContentLoaded', async () => {
  if (authToken) {
    try {
      currentUser = await apiCall('GET', '/api/me');
      onLoginSuccess();
    } catch (e) {
      localStorage.removeItem('auth_token');
      authToken = null;
    }
  }
});

// ===== Firestore 雲端資料讀寫 =====
const COLLECTIONS = ['members', 'cases', 'opinions', 'services', 'billings'];

async function loadCloudData() {
  try {
    const metaDoc = await fsdb.collection('system').doc('meta').get();
    const cloudVersion = metaDoc.exists ? (metaDoc.data().dataVersion || null) : null;
    const needsInit = !metaDoc.exists || cloudVersion !== DB_VERSION;

    if (!needsInit) {
      console.log('從雲端載入資料...');
      db = { members: [], cases: [], opinions: [], services: [], billings: [], diseases: [] };
      for (const col of COLLECTIONS) {
        const snapshot = await fsdb.collection(col).get();
        db[col] = snapshot.docs.map(doc => ({ ...doc.data(), _docId: doc.id }));
      }
      console.log(`雲端資料載入: ${db.cases.length} 個案, ${db.members.length} 成員`);

      if (db.cases.length === 0 && typeof RAW_CASES_DATA !== 'undefined' && RAW_CASES_DATA.length > 0) {
        console.warn('雲端資料為空，重新從本地清冊初始化...');
        db = generateDemoData();
        uploadAllData(db).catch(e => console.error('背景上傳失敗:', e));
        saveDB_local(db);
        return;
      }
    }

    if (needsInit) {
      console.log(`雲端資料版本不符 (${cloudVersion} → ${DB_VERSION})，重新初始化...`);
      db = generateDemoData();
      saveDB_local(db);
      console.log(`本地資料已就緒: ${db.cases.length} 個案, ${db.members.length} 成員`);

      (async () => {
        try {
          for (const col of COLLECTIONS) {
            const snapshot = await fsdb.collection(col).get();
            const docs = snapshot.docs;
            for (let i = 0; i < docs.length; i += 400) {
              const batch = fsdb.batch();
              docs.slice(i, i + 400).forEach(doc => batch.delete(doc.ref));
              await batch.commit();
            }
          }
          await uploadAllData(db);
          await fsdb.collection('system').doc('meta').set({
            initialized: true,
            dataVersion: DB_VERSION,
            updatedAt: firebase.firestore.FieldValue.serverTimestamp()
          });
          console.log('雲端資料同步完成');
        } catch (e) {
          console.error('雲端同步失敗（本地資料仍可用）:', e);
        }
      })();
    }

    saveDB_local(db);
  } catch (err) {
    console.error('雲端載入失敗，使用本地資料:', err);
    db = generateDemoData();
    saveDB_local(db);
  }
}

function saveDB_local(data) {
  localStorage.setItem(DB_KEY, JSON.stringify(data));
  localStorage.setItem(DB_VERSION_KEY, DB_VERSION);
}

async function uploadAllData(data) {
  for (const col of COLLECTIONS) {
    const items = data[col] || [];
    for (let i = 0; i < items.length; i += 400) {
      const batch = fsdb.batch();
      const chunk = items.slice(i, i + 400);
      chunk.forEach(item => {
        const docRef = fsdb.collection(col).doc(item.id);
        batch.set(docRef, item);
      });
      await batch.commit();
    }
  }
}

// ===== 覆寫 saveDB =====
function saveDB(data) {
  localStorage.setItem(DB_KEY, JSON.stringify(data));
  COLLECTIONS.forEach(col => {
    if (data[col]) {
      data[col].forEach(item => {
        if (item.id) {
          fsdb.collection(col).doc(item.id).set(item, { merge: true })
            .catch(err => console.error(`同步 ${col}/${item.id} 失敗:`, err));
        }
      });
    }
  });
}

// ===== 手動同步到雲端 =====
async function syncToCloud() {
  const btn = document.getElementById('btn-sync-cloud');
  if (!currentUser) { alert('請先登入再同步'); return; }

  btn.disabled = true;
  btn.textContent = '☁ 同步中...';

  try {
    if (!db || !db.cases || db.cases.length === 0) {
      db = generateDemoData();
      saveDB_local(db);
    }

    btn.textContent = '☁ 清除舊資料...';
    for (const col of COLLECTIONS) {
      const snapshot = await fsdb.collection(col).get();
      const docs = snapshot.docs;
      for (let i = 0; i < docs.length; i += 400) {
        const batch = fsdb.batch();
        docs.slice(i, i + 400).forEach(doc => batch.delete(doc.ref));
        await batch.commit();
      }
    }

    let uploaded = 0;
    const totalItems = COLLECTIONS.reduce((sum, col) => sum + (db[col] || []).length, 0);
    for (const col of COLLECTIONS) {
      const items = db[col] || [];
      for (let i = 0; i < items.length; i += 400) {
        const batch = fsdb.batch();
        const chunk = items.slice(i, i + 400);
        chunk.forEach(item => {
          const docRef = fsdb.collection(col).doc(item.id);
          batch.set(docRef, item);
        });
        await batch.commit();
        uploaded += chunk.length;
        btn.textContent = `☁ 上傳中 ${Math.round(uploaded / totalItems * 100)}%`;
      }
    }

    await fsdb.collection('system').doc('meta').set({
      initialized: true,
      dataVersion: DB_VERSION,
      updatedAt: firebase.firestore.FieldValue.serverTimestamp()
    });

    btn.textContent = '☁ 同步完成 ✓';
    btn.style.background = '#2196F3';
    alert(`雲端同步完成！\n${db.cases.length} 個案、${db.members.length} 成員、${db.services.length} 服務紀錄已上傳`);
  } catch (err) {
    console.error('同步失敗:', err);
    btn.textContent = '☁ 同步失敗 ✗';
    btn.style.background = '#f44336';
    alert('同步失敗: ' + err.message);
  } finally {
    setTimeout(() => {
      btn.disabled = false;
      btn.textContent = '☁ 同步到雲端';
      btn.style.background = '#4CAF50';
    }, 3000);
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
                <td><strong>${u.username}</strong></td>
                <td>${u.displayName}</td>
                <td><span class="badge ${u.role === 'admin' ? 'badge-danger' : 'badge-info'}">${u.role === 'admin' ? '管理員' : '使用者'}</span></td>
                <td><span class="badge ${u.status === 'active' ? 'badge-success' : 'badge-warning'}">${u.status === 'active' ? '啟用' : '停用'}</span></td>
                <td>${u.createdAt ? u.createdAt.substring(0, 10) : '-'}</td>
                <td>
                  <button class="btn btn-sm btn-outline" onclick="openEditUserModal('${u.id}','${u.username}','${u.displayName}','${u.role}','${u.status}')">編輯</button>
                  ${u.username !== '蕭輝哲' ? `<button class="btn btn-sm btn-outline" style="color:red;border-color:red" onclick="deleteUser('${u.id}','${u.username}')">刪除</button>` : ''}
                </td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
    `;
  } catch (err) {
    container.innerHTML = `<p style="color:red">載入失敗: ${err.message}</p>`;
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
    <input type="hidden" id="f-edit-id" value="${id}">
    <div class="form-group">
      <label class="form-label">帳號</label>
      <input class="form-input" id="f-edit-username" value="${username}">
    </div>
    <div class="form-group">
      <label class="form-label">新密碼 (留空不修改)</label>
      <input class="form-input" id="f-edit-password" type="password" placeholder="不修改請留空">
    </div>
    <div class="form-group">
      <label class="form-label">顯示名稱</label>
      <input class="form-input" id="f-edit-displayname" value="${displayName}">
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

// ===== LCMS 批次同步 =====
function syncLCMS() {
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
                  <td>${c.caseNo}</td>
                  <td>${c.name}</td>
                  <td>${c.idNumber}</td>
                  <td>${c.cmsLevel || '-'}</td>
                  <td>${c.category || '-'}</td>
                  <td>${c.district || '-'}</td>
                  <td>${c.enrollDate || '-'}</td>
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
                    <td rowspan="${item.diffs.length}">${item.lcms.name}</td>
                    <td rowspan="${item.diffs.length}">${item.lcms.idNumber}</td>
                  ` : ''}
                  <td><span style="color:#e65100;font-weight:500;">${d.field}</span></td>
                  <td style="color:#999;text-decoration:line-through;">${d.oldVal}</td>
                  <td style="color:#2e7d32;font-weight:500;">${d.newVal}</td>
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
                  <td>${c.name || '-'}</td>
                  <td>${c.idNumber || '-'}</td>
                  <td>${c.cmsLevel || '-'}</td>
                  <td>${c.doctorName || '-'}</td>
                  <td>${c.enrollDate || '-'}</td>
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
