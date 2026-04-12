/* ========================================
   居家失能個案追蹤管理系統 — 主應用邏輯
   ======================================== */

let db;
let currentMemberTab = 'doctors';

// ===== 安全工具函數 =====
function esc(str) {
  if (str == null) return '';
  return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}

function maskId(idNumber) {
  if (!idNumber || idNumber.length < 4) return '***';
  return idNumber.substring(0, 3) + '****' + idNumber.substring(idNumber.length - 3);
}

function escCSV(val) {
  if (val == null) return '';
  const s = String(val);
  if (/^[=+\-@\t\r]/.test(s)) return "'" + s;
  if (s.includes(',') || s.includes('"') || s.includes('\n')) return '"' + s.replace(/"/g, '""') + '"';
  return s;
}

function debounce(fn, delay = 300) {
  let timer;
  return function(...args) { clearTimeout(timer); timer = setTimeout(() => fn.apply(this, args), delay); };
}

// ===== 初始化 =====
// 由 firebase-config.js 的 onAuthStateChanged 觸發
document.addEventListener('DOMContentLoaded', () => {
  document.getElementById('sidebar-toggle').addEventListener('click', () => {
    document.getElementById('sidebar').classList.toggle('collapsed');
    document.getElementById('sidebar').classList.toggle('open');
  });
});

// ===== 導航 =====
function initNav() {
  document.querySelectorAll('.nav-item, .card-link').forEach(el => {
    el.addEventListener('click', e => {
      e.preventDefault();
      const page = el.dataset.page;
      if (page) switchPage(page);
    });
  });
}

const pageTitles = {
  dashboard: '主控儀表板', cases: '個案管理', members: '成員管理',
  services: '服務紀錄', opinion: '醫師意見書', billing: '費用申報',
  casemap: '個案地圖', alerts: '警示與待辦', reports: '報表匯出', users: '使用者管理'
};

function switchPage(page) {
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
  const el = document.getElementById('page-' + page);
  if (el) el.classList.add('active');
  const nav = document.querySelector(`.nav-item[data-page="${page}"]`);
  if (nav) nav.classList.add('active');
  document.getElementById('page-title').textContent = pageTitles[page] || '';

  const renderers = {
    dashboard: renderDashboard, cases: renderCases, members: renderMembers,
    services: renderServices, opinion: renderOpinions, billing: renderBilling,
    casemap: renderCaseMap, alerts: renderAlerts, users: renderUserManagement
  };
  if (renderers[page]) renderers[page]();
}

// ===== 篩選器初始化 =====
function initFilters() {
  const debouncedCases = debounce(renderCases);
  const debouncedServices = debounce(renderServices);
  const debouncedOpinions = debounce(renderOpinions);
  const debouncedMembers = debounce(renderMembers);
  ['case-search'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.addEventListener('input', debouncedCases);
  });
  ['case-filter-status','case-filter-level','case-filter-doctor','case-filter-category','case-filter-nurse'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.addEventListener('change', renderCases);
  });
  ['service-search'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.addEventListener('input', debouncedServices);
  });
  ['service-filter-type','service-filter-month'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.addEventListener('change', renderServices);
  });
  ['opinion-search'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.addEventListener('input', debouncedOpinions);
  });
  ['opinion-filter-status'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.addEventListener('change', renderOpinions);
  });
  ['billing-month','billing-filter-code','billing-filter-status'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.addEventListener('change', renderBilling);
  });
  ['member-search'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.addEventListener('input', debouncedMembers);
  });
  ['casemap-filter-doctor','casemap-filter-district','casemap-filter-cms'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.addEventListener('change', renderCaseMap);
  });
}

// ===== 格式化工具 =====
function formatDateCN(d) {
  if (typeof d === 'string') d = new Date(d);
  return `${d.getFullYear()}/${String(d.getMonth()+1).padStart(2,'0')}/${String(d.getDate()).padStart(2,'0')}`;
}

function serviceTypeLabel(t) {
  return { home:'家訪', phone:'電訪', video:'遠距視訊' }[t] || t;
}
function serviceTypeIcon(t) {
  return { home:'🏠', phone:'📞', video:'💻' }[t] || '📋';
}
function respondentLabel(r) {
  const map = { patient:'個案本人', spouse:'配偶', son:'兒子', daughter:'女兒',
    caregiver_foreign:'外籍看護', caregiver_local:'本國看護' };
  return map[r] || r;
}

function billingStatusBadge(s) {
  const map = {
    pending: '<span class="badge badge-warning">待申報</span>',
    submitted: '<span class="badge badge-info">已申報</span>',
    approved: '<span class="badge badge-success">已核付</span>',
    rejected: '<span class="badge badge-danger">退件</span>',
  };
  return map[s] || s;
}

function billingCodeLabel(code) {
  const map = {
    AA12:'醫師意見書', YA01:'電訪/視訊管理費', YA02:'家訪訪視費',
    YA03:'原民離島電訪/視訊', YA04:'原民離島家訪'
  };
  return code + ' ' + (map[code] || '');
}

function opinionStatusBadge(s) {
  const map = {
    valid: '<span class="badge badge-success">有效</span>',
    expiring: '<span class="badge badge-warning">即將到期</span>',
    expired: '<span class="badge badge-danger">已過期</span>',
    pending: '<span class="badge badge-gray">待開立</span>',
  };
  return map[s] || s;
}

// ===== Toast =====
function showToast(msg, type='success') {
  const container = document.getElementById('toast-container');
  const toast = document.createElement('div');
  toast.className = 'toast toast-' + type;
  toast.textContent = msg;
  container.appendChild(toast);
  setTimeout(() => { toast.style.opacity = '0'; setTimeout(() => toast.remove(), 300); }, 3000);
}

// ===== Modal =====
function openModal(title, bodyHTML, footerHTML) {
  document.getElementById('modal-title').textContent = title;
  document.getElementById('modal-body').innerHTML = bodyHTML;
  document.getElementById('modal-footer').innerHTML = footerHTML || '';
  document.getElementById('modal-overlay').classList.add('active');
}

function closeModal() {
  document.getElementById('modal-overlay').classList.remove('active');
}

// ==========================================
//  Dashboard
// ==========================================
function renderDashboard() {
  const active = getActiveCases(db);
  const docs = getDoctors(db).filter(d => d.status === 'active');
  const now = new Date();
  const ym = fmt(now).substring(0,7);

  // Stats
  document.getElementById('stat-total-cases').textContent = active.length;
  const thisMonthCases = active.filter(c => c.enrollDate.startsWith(ym)).length;
  document.getElementById('trend-cases').textContent = `+${thisMonthCases} 本月`;
  document.getElementById('stat-total-doctors').textContent = docs.length;
  document.getElementById('trend-doctors').textContent = `平均 ${docs.length ? Math.round(active.length/docs.length) : 0} 案/醫師`;

  // 意見書狀態（基於醫師家訪日期，6個月有效期）
  let opExpired = 0, opExpiring = 0, opNone = 0;
  const todayStr = fmt(new Date());
  active.forEach(c => {
    const lastVisit = c.doctorVisitDate || '';
    const op = getLatestOpinion(db, c.id);
    // 優先用 doctorVisitDate，其次用意見書記錄
    const lastDate = lastVisit || (op ? op.issueDate : '');
    if (!lastDate) { opNone++; return; }
    const daysSince = daysBetween(lastDate, todayStr);
    if (daysSince > 180) opExpired++;        // 超過6個月 → 逾期
    else if (daysSince >= 150) opExpiring++;  // 5~6個月 → 待更新
  });
  const opDue = opExpired + opExpiring + opNone;
  document.getElementById('stat-opinion-due').textContent = opDue;
  const opDetailEl = document.getElementById('trend-opinion');
  if (opDetailEl) opDetailEl.innerHTML = opExpired ? `<span style="color:var(--danger)">${opExpired}逾期</span> ${opExpiring}待更新 ${opNone}未開` : `${opExpiring}待更新 ${opNone}未開`;

  // 警示
  const alerts = generateAlerts();
  const dangerCount = alerts.filter(a => a.level === 'danger').length;
  document.getElementById('stat-alerts').textContent = dangerCount;
  updateAlertBadge();

  // KPI Bars
  const kpiData = computeKPI();
  const kpiContainer = document.getElementById('dashboard-kpi-bars');
  kpiContainer.innerHTML = kpiData.map(k => {
    const pct = Math.min(100, k.value);
    const color = k.value >= k.target ? 'var(--success)' : (k.value >= k.target * 0.7 ? 'var(--warning)' : 'var(--danger)');
    return `<div class="kpi-bar-item">
      <div class="kpi-bar-header">
        <span class="kpi-bar-label">${k.label}</span>
        <span class="kpi-bar-values">${k.value.toFixed(1)}% / 目標 ${k.target}%</span>
      </div>
      <div class="kpi-bar-track">
        <div class="kpi-bar-fill" style="width:${pct}%;background:${color}">${k.value.toFixed(0)}%</div>
        <div class="kpi-bar-target" style="left:${k.target}%" data-label="目標${k.target}%"></div>
      </div>
    </div>`;
  }).join('');

  // Doctor Load
  const loadContainer = document.getElementById('doctor-load-chart');
  loadContainer.innerHTML = docs.map(d => {
    const count = getCaseCountByMember(db, d.id);
    const pct = (count / 200 * 100).toFixed(0);
    const color = count > 180 ? 'var(--danger)' : (count > 140 ? 'var(--warning)' : 'var(--primary)');
    return `<div class="load-bar-item">
      <div class="load-bar-name" title="${esc(d.name)}">${esc(d.name)}</div>
      <div class="load-bar-track">
        <div class="load-bar-fill" style="width:${pct}%;background:${color}">${count}</div>
      </div>
      <div class="load-bar-count">${count}/200</div>
    </div>`;
  }).join('');

  // Recent Services
  const recentSvc = db.services.sort((a,b) => b.date.localeCompare(a.date)).slice(0, 7);
  const recentContainer = document.getElementById('recent-services-list');
  recentContainer.innerHTML = recentSvc.map(s => {
    const c = findCase(db, s.caseId);
    const n = findMember(db, s.nurseId);
    return `<div class="recent-item">
      <div class="recent-type ${s.type}">${serviceTypeIcon(s.type)}</div>
      <div class="recent-info"><strong>${c ? esc(c.name) : esc(s.caseId)}</strong> — ${serviceTypeLabel(s.type)}${n ? ' / ' + esc(n.name) : ''}</div>
      <div class="recent-date">${s.date}</div>
    </div>`;
  }).join('');

  // Alerts
  const dashAlerts = alerts.slice(0, 8);
  const alertContainer = document.getElementById('dashboard-alerts-list');
  alertContainer.innerHTML = dashAlerts.length === 0
    ? '<div style="text-align:center;color:var(--gray-400);padding:2rem">目前無待辦事項</div>'
    : dashAlerts.map(a => `<div class="alert-item level-${a.level}">
      <div class="alert-icon">${a.level === 'danger' ? '🔴' : (a.level === 'warning' ? '🟡' : '🔵')}</div>
      <div class="alert-body">
        <div class="alert-title">${esc(a.title)}</div>
        <div class="alert-desc">${esc(a.desc)}</div>
      </div>
      <div class="alert-action"><button class="btn btn-xs btn-outline" onclick="switchPage('${a.page || 'alerts'}')">處理</button></div>
    </div>`).join('');

  // 更新醫師下拉選單
  populateDoctorSelects();
}

function populateDoctorSelects() {
  const docs = getDoctors(db).filter(d => d.status === 'active');
  ['case-filter-doctor','kpi-filter-doctor'].forEach(id => {
    const el = document.getElementById(id);
    if (!el) return;
    const firstOpt = el.options[0].outerHTML;
    el.innerHTML = firstOpt + docs.map(d => `<option value="${d.name}">${esc(d.name)}</option>`).join('');
  });
  const nurses = getNurses(db).filter(n => n.status === 'active');
  const nurseEl = document.getElementById('case-filter-nurse');
  if (nurseEl) {
    const firstOpt = nurseEl.options[0].outerHTML;
    nurseEl.innerHTML = firstOpt + nurses.map(n => `<option value="${n.id}">${esc(n.name)}</option>`).join('');
  }
}

// ==========================================
//  個案管理
// ==========================================
function renderCases() {
  const search = (document.getElementById('case-search')?.value || '').toLowerCase();
  const status = document.getElementById('case-filter-status')?.value || '';
  const level = document.getElementById('case-filter-level')?.value || '';
  const doctor = document.getElementById('case-filter-doctor')?.value || '';
  const category = document.getElementById('case-filter-category')?.value || '';
  const nurseFilter = document.getElementById('case-filter-nurse')?.value || '';

  let filtered = db.cases.filter(c => {
    if (search && !c.name.toLowerCase().includes(search) && !c.idNumber.toLowerCase().includes(search) && !(c.id && c.id.toLowerCase().includes(search)) && !(c.district && c.district.includes(search)) && !(c.contactPerson && c.contactPerson.includes(search))) return false;
    if (status && c.status !== status) return false;
    if (level && c.cmsLevel !== parseInt(level)) return false;
    if (doctor) {
      const docMember = db.members.find(m => m.name === doctor);
      const matchById = docMember && c.doctorId === docMember.id;
      const matchByName = c.doctorName === doctor;
      if (!matchById && !matchByName) return false;
    }
    if (category && c.category !== category) return false;
    if (nurseFilter && c.nurseId !== nurseFilter) return false;
    return true;
  });

  const tbody = document.getElementById('cases-tbody');
  tbody.innerHTML = filtered.map(c => {
    const doc = findMember(db, c.doctorId);
    const nurse = findMember(db, c.nurseId);
    const op = getLatestOpinion(db, c.id);
    const svcThisMonth = getServiceThisMonth(db, c.id);
    // 意見書狀態基於醫師家訪日期
    const _lastVD = c.doctorVisitDate || (op ? op.issueDate : '');
    let opBadge;
    if (!_lastVD) {
      opBadge = '<span class="badge badge-gray">待開立</span>';
    } else {
      const _ds = daysBetween(_lastVD, fmt(new Date()));
      if (_ds > 180) opBadge = '<span class="badge badge-danger">逾期</span>';
      else if (_ds >= 150) opBadge = '<span class="badge badge-warning">待更新</span>';
      else opBadge = '<span class="badge badge-success">有效</span>';
    }
    const statusBadge = c.status === 'active'
      ? '<span class="badge badge-success">收案中</span>'
      : '<span class="badge badge-gray">已結案</span>';
    const svcBadge = svcThisMonth.length > 0
      ? `<span class="badge badge-success">${svcThisMonth.length}次</span>`
      : '<span class="badge badge-danger">未服務</span>';

    // 本月追蹤狀態：從 monthlyTracking 取當月資料
    const curMonth = String(new Date().getMonth() + 1);
    const trackingText = c.monthlyTracking ? c.monthlyTracking[curMonth] || '' : '';
    const trackBadge = trackingText
      ? `<span class="badge badge-success" title="${esc(trackingText)}">${esc(trackingText)}</span>`
      : (c.status === 'active' ? '<span class="badge badge-danger">未追蹤</span>' : '-');

    // 本月電訪狀態
    const phoneThisMonth = getPhoneVisitsThisMonth(db, c.id);
    const phoneBadge = phoneThisMonth.length > 0
      ? `<span class="badge badge-success">電${phoneThisMonth.length}</span>`
      : (c.status === 'active' ? '<span class="badge badge-warning">未電訪</span>' : '-');

    // 上次家訪狀態
    const lastHomeVisit = db.services.filter(s => s.caseId === c.id && s.type === 'home')
      .sort((a,b) => b.date.localeCompare(a.date))[0];
    let homeBadge = '-';
    if (lastHomeVisit) {
      const homeAge = daysBetween(lastHomeVisit.date, fmt(new Date()));
      const homeColor = homeAge > 120 ? 'danger' : homeAge > 90 ? 'warning' : 'success';
      homeBadge = `<span class="badge badge-${homeColor}" title="${homeAge}天前">${lastHomeVisit.date.substring(5)}</span>`;
    } else if (c.status === 'active') {
      homeBadge = '<span class="badge badge-danger">未家訪</span>';
    }

    return `<tr class="${c.status === 'active' && !trackingText ? 'row-warning' : ''}">
      <td><small>${c.caseNo || c.id}</small></td>
      <td><strong>${esc(c.name)}</strong></td>
      <td><small>${esc(c.category || '-')}</small></td>
      <td><span class="badge badge-primary">${c.cmsLevel || '-'}</span></td>
      <td><small>${esc(c.village || c.district || '-')}</small></td>
      <td>${doc ? esc(doc.name) : esc(c.doctorName || '-')}</td>
      <td>${opBadge}</td>
      <td>${phoneBadge}</td>
      <td>${homeBadge}</td>
      <td>${trackBadge}</td>
      <td>${statusBadge}</td>
      <td>
        <button class="btn btn-xs btn-outline" onclick="viewCase('${c.id}')">檢視</button>
        ${c.status === 'active' ? `<button class="btn btn-xs btn-warning" onclick="closeCase('${c.id}')">結案</button>` : ''}
      </td>
    </tr>`;
  }).join('');
}

function openCaseModal(caseId) {
  const c = caseId ? findCase(db, caseId) : null;
  const docs = getDoctors(db).filter(d => d.status === 'active');
  const nurs = getNurses(db).filter(n => n.status === 'active');

  const body = `
    <div class="form-row">
      <div class="form-group">
        <label class="form-label">照管案號 *</label>
        <input class="form-input" id="f-case-no" value="${c ? (c.caseNo || c.id) : ''}" style="width:100%" ${c ? 'readonly' : ''}>
      </div>
      <div class="form-group">
        <label class="form-label">姓名 *</label>
        <input class="form-input" id="f-case-name" value="${c ? esc(c.name) : ''}" style="width:100%">
      </div>
      <div class="form-group">
        <label class="form-label">身分證字號 *</label>
        <input class="form-input" id="f-case-id-number" value="${c ? esc(c.idNumber) : ''}" style="width:100%">
      </div>
    </div>
    <div class="form-row">
      <div class="form-group">
        <label class="form-label">性別</label>
        <select class="form-select" id="f-case-gender" style="width:100%">
          <option value="M" ${c?.gender==='M'?'selected':''}>男</option>
          <option value="F" ${c?.gender==='F'?'selected':''}>女</option>
        </select>
      </div>
      <div class="form-group">
        <label class="form-label">身分別</label>
        <select class="form-select" id="f-case-category" style="width:100%">
          <option value="第一類" ${c?.category==='第一類'?'selected':''}>第一類</option>
          <option value="第二類" ${c?.category==='第二類'?'selected':''}>第二類</option>
          <option value="第三類" ${c?.category==='第三類'?'selected':''}>第三類</option>
          <option value="一般戶" ${c?.category==='一般戶'?'selected':''}>一般戶</option>
        </select>
      </div>
      <div class="form-group">
        <label class="form-label">CMS等級 *</label>
        <select class="form-select" id="f-case-level" style="width:100%">
          ${[2,3,4,5,6,7,8].map(l => `<option value="${l}" ${c?.cmsLevel===l?'selected':''}>${l}</option>`).join('')}
        </select>
      </div>
    </div>
    <div class="form-row">
      <div class="form-group">
        <label class="form-label">聯絡電話</label>
        <input class="form-input" id="f-case-phone" value="${c ? esc(c.phone) : ''}" style="width:100%">
      </div>
      <div class="form-group">
        <label class="form-label">主要聯絡人</label>
        <input class="form-input" id="f-case-contact" value="${c ? esc(c.contactPerson||'') : ''}" style="width:100%">
      </div>
      <div class="form-group">
        <label class="form-label">收案日期 *</label>
        <input type="date" class="form-input" id="f-case-enroll" value="${c ? c.enrollDate : fmt(new Date())}" style="width:100%">
      </div>
    </div>
    <div class="form-row">
      <div class="form-group" style="flex:2">
        <label class="form-label">住址</label>
        <input class="form-input" id="f-case-address" value="${c ? esc(c.address) : ''}" style="width:100%">
      </div>
      <div class="form-group">
        <label class="form-label">居住地里</label>
        <input class="form-input" id="f-case-district" value="${c ? esc(c.district||'') : ''}" style="width:100%">
      </div>
    </div>
    <div class="form-row">
      <div class="form-group">
        <label class="form-label">負責醫師 *</label>
        <select class="form-select" id="f-case-doctor" style="width:100%">
          ${docs.map(d => `<option value="${d.id}" ${c?.doctorId===d.id?'selected':''}>${esc(d.name)} (${esc(d.specialty)}) [${getCaseCountByMember(db,d.id)}/200]</option>`).join('')}
        </select>
      </div>
      <div class="form-group">
        <label class="form-label">個案管理師 *</label>
        <select class="form-select" id="f-case-nurse" style="width:100%">
          ${nurs.map(n => `<option value="${n.id}" ${c?.nurseId===n.id?'selected':''}>${esc(n.name)} [${getCaseCountByMember(db,n.id)}/200]</option>`).join('')}
        </select>
      </div>
    </div>
    <div class="form-group">
      <label class="form-label">備註</label>
      <input class="form-input" id="f-case-notes" value="${c ? esc(c.notes||'') : ''}" style="width:100%">
    </div>
    <div class="form-group">
      <label class="form-label">是否為原住民/離島地區</label>
      <label class="form-checkbox"><input type="checkbox" id="f-case-remote" ${c?.isRemoteArea?'checked':''}> 是</label>
    </div>
  `;
  const footer = `
    <button class="btn btn-outline" onclick="closeModal()">取消</button>
    <button class="btn btn-primary" onclick="saveCase('${caseId || ''}')">${c ? '更新' : '新增'}個案</button>
  `;
  openModal(c ? '編輯個案' : '新增個案', body, footer);
}

function saveCase(editId) {
  const name = document.getElementById('f-case-name').value.trim();
  const idNumber = document.getElementById('f-case-id-number').value.trim();
  if (!name || !idNumber) { showToast('請填寫必要欄位', 'warning'); return; }

  const doctorId = document.getElementById('f-case-doctor').value;
  const docMember = findMember(db, doctorId);

  const data = {
    name,
    idNumber,
    gender: document.getElementById('f-case-gender').value,
    category: document.getElementById('f-case-category').value,
    cmsLevel: parseInt(document.getElementById('f-case-level').value),
    phone: document.getElementById('f-case-phone').value,
    contactPerson: document.getElementById('f-case-contact').value,
    address: document.getElementById('f-case-address').value,
    district: document.getElementById('f-case-district').value,
    enrollDate: document.getElementById('f-case-enroll').value,
    doctorId,
    doctorName: docMember ? docMember.name : '',
    nurseId: document.getElementById('f-case-nurse').value,
    notes: document.getElementById('f-case-notes').value,
    isRemoteArea: document.getElementById('f-case-remote').checked,
  };

  // 驗證醫師案量上限
  const docCount = getCaseCountByMember(db, data.doctorId);
  if (!editId && docCount >= 200) {
    showToast('該醫師已達收案上限 200 案', 'danger');
    return;
  }

  if (editId) {
    const c = findCase(db, editId);
    Object.assign(c, data);
    showToast('個案已更新');
  } else {
    const caseNo = document.getElementById('f-case-no').value.trim() || ('C' + String(db.cases.length + 1).padStart(4, '0'));
    db.cases.push({
      id: caseNo, ...data, status: 'active', closeReason: null, closeDate: null,
      firstReferralDate: data.enrollDate,
      nurseVisitDate: null, doctorVisitDate: null,
      serviceDays: 0, monthlyTracking: {},
      scheduledVisit: '', closeInfo: '',
      diagnoses: [], hasHypertension: false, hasDiabetes: false, hasHyperlipidemia: false,
      diseaseStatus: '穩定',
      acpExplained: false, acpExplainedDate: null,
      adExplained: false, adExplainedDate: null,
      acpSigned: false, acpSignedDate: null,
      nhiRegistered: false, familyAcpExplained: false,
    });
    showToast('個案新增成功');
  }
  saveDB(db);
  closeModal();
  renderCases();
}

function viewCase(caseId) {
  const c = findCase(db, caseId);
  if (!c) return;
  const doc = findMember(db, c.doctorId);
  const nurse = findMember(db, c.nurseId);
  const op = getLatestOpinion(db, c.id);
  const svcs = getServicesByCase(db, c.id);
  const svcThisMonth = getServiceThisMonth(db, c.id);
  const homeVisits = getHomeVisitsThisYear(db, c.id);

  // 月度追蹤彙整
  const monthNames = ['1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月'];
  const trackingHtml = monthNames.map((mName, idx) => {
    const val = c.monthlyTracking ? c.monthlyTracking[String(idx + 1)] || '' : '';
    const bgColor = val.includes('家') ? '#dcfce7' : val.includes('電') ? '#dbeafe' : val.includes('視') ? '#fef3c7' : val === '結' ? '#fee2e2' : '';
    return `<td style="padding:4px 6px;text-align:center;font-size:.78rem;background:${bgColor}">${val || '-'}</td>`;
  }).join('');

  const body = `
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:1rem">
      <div>
        <h4 style="margin-bottom:.5rem;color:var(--gray-600)">基本資料</h4>
        <table style="font-size:.85rem;width:100%">
          <tr><td style="padding:2px 8px;color:var(--gray-500)">照管案號</td><td><strong>${c.caseNo || c.id}</strong></td></tr>
          <tr><td style="padding:2px 8px;color:var(--gray-500)">姓名</td><td><strong>${esc(c.name)}</strong></td></tr>
          <tr><td style="padding:2px 8px;color:var(--gray-500)">身分證</td><td>${maskId(c.idNumber)}</td></tr>
          <tr><td style="padding:2px 8px;color:var(--gray-500)">性別</td><td>${c.gender==='M'?'男':'女'}</td></tr>
          <tr><td style="padding:2px 8px;color:var(--gray-500)">身分別</td><td>${esc(c.category || '-')}</td></tr>
          <tr><td style="padding:2px 8px;color:var(--gray-500)">CMS等級</td><td>第${c.cmsLevel}級</td></tr>
          <tr><td style="padding:2px 8px;color:var(--gray-500)">收案日期</td><td>${c.enrollDate || '-'}</td></tr>
          <tr><td style="padding:2px 8px;color:var(--gray-500)">服務天數</td><td>${c.serviceDays || '-'} 天</td></tr>
          <tr><td style="padding:2px 8px;color:var(--gray-500)">地址</td><td>${esc(c.address)}</td></tr>
          <tr><td style="padding:2px 8px;color:var(--gray-500)">居住地里</td><td>${esc(c.village || c.district || '-')}</td></tr>
        </table>
      </div>
      <div>
        <h4 style="margin-bottom:.5rem;color:var(--gray-600)">聯絡與照護資訊</h4>
        <table style="font-size:.85rem;width:100%">
          <tr><td style="padding:2px 8px;color:var(--gray-500)">聯絡電話</td><td>${esc(c.phone || '-')}</td></tr>
          <tr><td style="padding:2px 8px;color:var(--gray-500)">主要聯絡人</td><td>${esc(c.contactPerson || '-')}</td></tr>
          <tr><td style="padding:2px 8px;color:var(--gray-500)">A單位名稱</td><td>${esc(c.unitName || '-')}</td></tr>
          <tr><td style="padding:2px 8px;color:var(--gray-500)">負責醫師</td><td>${doc ? esc(doc.name) : esc(c.doctorName || '-')}</td></tr>
          <tr><td style="padding:2px 8px;color:var(--gray-500)">個管師</td><td>${nurse ? esc(nurse.name) : '-'}</td></tr>
          <tr><td style="padding:2px 8px;color:var(--gray-500)">醫師家訪日</td><td>${c.doctorVisitDate || '-'}</td></tr>
          <tr><td style="padding:2px 8px;color:var(--gray-500)">個管師家訪日</td><td>${c.nurseVisitDate || '-'}</td></tr>
          <tr><td style="padding:2px 8px;color:var(--gray-500)">意見書</td><td>${(() => {
            const vd = c.doctorVisitDate || (op ? op.issueDate : '');
            if (!vd) return '尚未開立';
            const ds = daysBetween(vd, fmt(new Date()));
            const expD = new Date(vd); expD.setMonth(expD.getMonth() + 6);
            const badge = ds > 180 ? '<span class="badge badge-danger">逾期</span>' : ds >= 150 ? '<span class="badge badge-warning">待更新</span>' : '<span class="badge badge-success">有效</span>';
            return badge + ' 家訪 ' + vd + '（到期 ' + expD.toISOString().slice(0,10) + '）';
          })()}</td></tr>
          <tr><td style="padding:2px 8px;color:var(--gray-500)">預約家訪</td><td>${esc(c.scheduledVisit || '-')}</td></tr>
          <tr><td style="padding:2px 8px;color:var(--gray-500)">備註</td><td>${esc(c.notes || '-')}</td></tr>
          ${c.status === 'closed' ? `<tr><td style="padding:2px 8px;color:var(--gray-500)">結案資訊</td><td style="color:#dc2626">${esc(c.closeInfo || c.closeReason || '-')}</td></tr>` : ''}
        </table>
      </div>
    </div>
    <h4 style="margin:1rem 0 .5rem;color:var(--gray-600)">追蹤狀態摘要</h4>
    <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:.5rem;margin-bottom:1rem">
      ${(() => {
        const yy = String(new Date().getFullYear());
        const phoneAll = svcs.filter(s => s.type === 'phone' && s.date.startsWith(yy));
        const homeAll = svcs.filter(s => s.type === 'home' && s.date.startsWith(yy));
        const lastPhone = svcs.find(s => s.type === 'phone');
        const lastHomeV = svcs.find(s => s.type === 'home');
        const opInfo = getOpinionExpiryInfo(db, c.id);
        const opText = opInfo.opinion
          ? (opInfo.daysLeft > 30 ? `有效 (剩${opInfo.daysLeft}天)` : opInfo.daysLeft > 0 ? `即將到期 (剩${opInfo.daysLeft}天)` : '已過期')
          : '未開立';
        const opColor = opInfo.opinion ? (opInfo.daysLeft > 30 ? '#16a34a' : opInfo.daysLeft > 0 ? '#d97706' : '#dc2626') : '#dc2626';
        return `
          <div style="background:#f0f9ff;border-radius:8px;padding:.5rem;text-align:center">
            <div style="font-size:1.2rem;font-weight:700;color:#2563eb">${phoneAll.length}</div>
            <div style="font-size:.75rem;color:var(--gray-500)">本年電訪</div>
            <div style="font-size:.7rem;color:var(--gray-400)">${lastPhone ? '上次: ' + lastPhone.date : '無紀錄'}</div>
          </div>
          <div style="background:#f0fdf4;border-radius:8px;padding:.5rem;text-align:center">
            <div style="font-size:1.2rem;font-weight:700;color:#16a34a">${homeAll.length}</div>
            <div style="font-size:.75rem;color:var(--gray-500)">本年家訪</div>
            <div style="font-size:.7rem;color:var(--gray-400)">${lastHomeV ? '上次: ' + lastHomeV.date : '無紀錄'}</div>
          </div>
          <div style="background:#fffbeb;border-radius:8px;padding:.5rem;text-align:center">
            <div style="font-size:1.2rem;font-weight:700;color:${opColor}">${esc(opText)}</div>
            <div style="font-size:.75rem;color:var(--gray-500)">意見書</div>
            <div style="font-size:.7rem;color:var(--gray-400)">${opInfo.opinion ? '到期: ' + opInfo.opinion.expiryDate : '-'}</div>
          </div>
          <div style="background:#fdf2f8;border-radius:8px;padding:.5rem;text-align:center">
            <div style="font-size:1.2rem;font-weight:700;color:#7c3aed">${svcThisMonth.length}</div>
            <div style="font-size:.75rem;color:var(--gray-500)">本月服務</div>
            <div style="font-size:.7rem;color:var(--gray-400)">電${getPhoneVisitsThisMonth(db, c.id).length} 家${svcThisMonth.filter(s=>s.type==='home').length}</div>
          </div>`;
      })()}
    </div>
    <h4 style="margin:1rem 0 .5rem;color:var(--gray-600)">115年度月追蹤紀錄 <small style="color:var(--gray-400)">(家=家訪 電=電訪 視=視訊 結=結案)</small></h4>
    <table class="data-table" style="font-size:.82rem">
      <thead><tr>${monthNames.map(m => `<th style="text-align:center">${m}</th>`).join('')}</tr></thead>
      <tbody><tr>${trackingHtml}</tr></tbody>
    </table>
    <h4 style="margin:1rem 0 .5rem;color:var(--gray-600)">疾病診斷</h4>
    <table class="data-table" style="font-size:.82rem">
      <thead><tr><th>ICD-10</th><th>疾病名稱</th><th>發病時間</th></tr></thead>
      <tbody>${(c.diagnoses||[]).map(d => `<tr><td>${esc(d.icd)}</td><td>${esc(d.name)}</td><td>${esc(d.onset)}</td></tr>`).join('') || '<tr><td colspan="3" style="text-align:center;color:var(--gray-400)">尚未記錄</td></tr>'}</tbody>
    </table>
    <h4 style="margin:1rem 0 .5rem;color:var(--gray-600)">近期服務紀錄 (最近5筆)</h4>
    <table class="data-table" style="font-size:.82rem">
      <thead><tr><th>日期</th><th>形式</th><th>血壓</th><th>血糖</th><th>衛教</th></tr></thead>
      <tbody>${svcs.slice(0,5).map(s => `<tr>
        <td>${s.date}</td><td>${serviceTypeLabel(s.type)}</td>
        <td>${s.bpMeasured ? s.bpSystolic+'/'+s.bpDiastolic : '-'}</td>
        <td>${s.hba1cMonitored ? 'HbA1c '+s.hba1cValue+'%' : '-'}</td>
        <td>${s.educationProvided ? '✅' : '-'}</td>
      </tr>`).join('') || '<tr><td colspan="5" style="text-align:center;color:var(--gray-400)">尚無紀錄</td></tr>'}</tbody>
    </table>
  `;
  const footer = `
    <button class="btn btn-outline" onclick="closeModal()">關閉</button>
    <button class="btn btn-primary" onclick="closeModal();openCaseModal('${c.id}')">編輯</button>
  `;
  openModal('個案詳情 — ' + esc(c.name), body, footer);
}

function closeCase(caseId) {
  const body = `
    <div class="form-group">
      <label class="form-label">結案原因 *</label>
      <select class="form-select" id="f-close-reason" style="width:100%">
        <option value="死亡">死亡</option>
        <option value="遷居">遷居</option>
        <option value="入住機構">入住機構</option>
        <option value="拒絕訪視">拒絕訪視</option>
      </select>
    </div>
    <div class="form-group">
      <label class="form-label">結案日期</label>
      <input type="date" class="form-input" id="f-close-date" value="${fmt(new Date())}" style="width:100%">
    </div>
  `;
  const footer = `
    <button class="btn btn-outline" onclick="closeModal()">取消</button>
    <button class="btn btn-danger" onclick="confirmCloseCase('${caseId}')">確認結案</button>
  `;
  openModal('結案確認', body, footer);
}

function confirmCloseCase(caseId) {
  const c = findCase(db, caseId);
  c.status = 'closed';
  c.closeReason = document.getElementById('f-close-reason').value;
  c.closeDate = document.getElementById('f-close-date').value;
  saveDB(db);
  closeModal();
  showToast('個案已結案');
  renderCases();
}

// ==========================================
//  成員管理
// ==========================================
function switchMemberTab(tab) {
  currentMemberTab = tab;
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.toggle('active', b.dataset.tab === tab));
  renderMembers();
}

function renderMembers() {
  const search = (document.getElementById('member-search')?.value || '').toLowerCase();
  const list = currentMemberTab === 'doctors' ? getDoctors(db) : getNurses(db);
  const filtered = list.filter(m => !search || m.name.toLowerCase().includes(search));

  document.getElementById('members-tbody').innerHTML = filtered.map(m => {
    const count = getCaseCountByMember(db, m.id);
    const pct = (count / 200 * 100).toFixed(0);
    const loadColor = count > 180 ? 'danger' : (count > 140 ? 'warning' : 'success');
    return `<tr>
      <td><strong>${esc(m.name)}</strong></td>
      <td>${m.role === 'doctor' ? '醫師' : '個管師'}</td>
      <td>${esc(m.specialty)}</td>
      <td>${count}</td>
      <td>200</td>
      <td><span class="badge badge-${loadColor}">${pct}%</span></td>
      <td>${m.acpTrained ? '<span class="badge badge-success">已完成</span>' : '<span class="badge badge-danger">未完成</span>'}</td>
      <td>${m.role === 'doctor' ? (m.opinionTrained ? '<span class="badge badge-success">已完成</span>' : '<span class="badge badge-danger">未完成</span>') : '-'}</td>
      <td><span class="badge badge-${m.status==='active'?'success':'gray'}">${m.status==='active'?'在職':'離職'}</span></td>
      <td><button class="btn btn-xs btn-outline" onclick="openMemberModal('${m.id}')">編輯</button></td>
    </tr>`;
  }).join('');
}

function openMemberModal(memberId) {
  const m = memberId ? findMember(db, memberId) : null;
  const body = `
    <div class="form-row">
      <div class="form-group">
        <label class="form-label">姓名 *</label>
        <input class="form-input" id="f-member-name" value="${m?esc(m.name):''}" style="width:100%">
      </div>
      <div class="form-group">
        <label class="form-label">職稱 *</label>
        <select class="form-select" id="f-member-role" style="width:100%">
          <option value="doctor" ${m?.role==='doctor'?'selected':''}>醫師</option>
          <option value="nurse" ${m?.role==='nurse'?'selected':''}>個案管理師(護理師)</option>
        </select>
      </div>
    </div>
    <div class="form-row">
      <div class="form-group">
        <label class="form-label">專科</label>
        <input class="form-input" id="f-member-spec" value="${m?esc(m.specialty):''}" style="width:100%">
      </div>
      <div class="form-group">
        <label class="form-label">聯絡電話</label>
        <input class="form-input" id="f-member-phone" value="${m?esc(m.phone):''}" style="width:100%">
      </div>
    </div>
    <div class="form-group">
      <label class="form-label">加入日期</label>
      <input type="date" class="form-input" id="f-member-join" value="${m?m.joinDate:fmt(new Date())}" style="width:100%">
    </div>
    <div class="form-row">
      <div class="form-group">
        <label class="form-checkbox"><input type="checkbox" id="f-member-acp" ${m?.acpTrained?'checked':''}> ACP訓練已完成</label>
      </div>
      <div class="form-group">
        <label class="form-checkbox"><input type="checkbox" id="f-member-opinion" ${m?.opinionTrained?'checked':''}> 醫師意見書訓練已完成</label>
      </div>
    </div>
  `;
  const footer = `
    <button class="btn btn-outline" onclick="closeModal()">取消</button>
    <button class="btn btn-primary" onclick="saveMember('${memberId||''}')">${m?'更新':'新增'}</button>
  `;
  openModal(m ? '編輯成員' : '新增成員', body, footer);
}

function saveMember(editId) {
  const name = document.getElementById('f-member-name').value.trim();
  if (!name) { showToast('請填寫姓名', 'warning'); return; }
  const data = {
    name,
    role: document.getElementById('f-member-role').value,
    specialty: document.getElementById('f-member-spec').value,
    phone: document.getElementById('f-member-phone').value,
    joinDate: document.getElementById('f-member-join').value,
    acpTrained: document.getElementById('f-member-acp').checked,
    acpTrainedDate: document.getElementById('f-member-acp').checked ? fmt(new Date()) : null,
    opinionTrained: document.getElementById('f-member-opinion').checked,
    opinionTrainedDate: document.getElementById('f-member-opinion').checked ? fmt(new Date()) : null,
    status: 'active',
  };
  if (editId) {
    const m = findMember(db, editId);
    Object.assign(m, data);
    showToast('成員已更新');
  } else {
    const prefix = data.role === 'doctor' ? 'D' : 'N';
    const count = db.members.filter(m => m.role === data.role).length;
    data.id = prefix + String(count + 1).padStart(3, '0');
    db.members.push(data);
    showToast('成員新增成功');
  }
  saveDB(db);
  closeModal();
  renderMembers();
}

// ==========================================
//  服務紀錄
// ==========================================
function renderServices() {
  const search = (document.getElementById('service-search')?.value || '').toLowerCase();
  const type = document.getElementById('service-filter-type')?.value || '';
  const month = document.getElementById('service-filter-month')?.value || '';

  let filtered = db.services.filter(s => {
    if (type && s.type !== type) return false;
    if (month && !s.date.startsWith(month)) return false;
    if (search) {
      const c = findCase(db, s.caseId);
      if (!c || !c.name.toLowerCase().includes(search)) return false;
    }
    return true;
  }).sort((a, b) => b.date.localeCompare(a.date));

  document.getElementById('services-tbody').innerHTML = filtered.slice(0, 100).map(s => {
    const c = findCase(db, s.caseId);
    const n = findMember(db, s.nurseId);
    return `<tr>
      <td>${s.date}</td>
      <td>${c ? esc(c.name) : esc(s.caseId)}</td>
      <td><span class="badge badge-${s.type==='home'?'primary':(s.type==='phone'?'success':'info')}">${serviceTypeLabel(s.type)}</span></td>
      <td>${respondentLabel(s.respondent)}</td>
      <td>${n ? esc(n.name) : '-'}</td>
      <td>${s.bpMeasured ? s.bpSystolic + '/' + s.bpDiastolic : '-'}</td>
      <td>${s.hba1cMonitored ? s.hba1cValue + '%' : '-'}</td>
      <td>${s.lipidMonitored ? '✅' : '-'}</td>
      <td>${s.educationProvided ? '✅' : '-'}</td>
      <td>${s.acpExplained ? '✅' : '-'}</td>
      <td>${billingStatusBadge(s.billingStatus)}</td>
      <td><button class="btn btn-xs btn-outline" onclick="viewService('${s.id}')">檢視</button></td>
    </tr>`;
  }).join('');
}

function openServiceModal() {
  const activeCases = getActiveCases(db);
  const nurses = getNurses(db).filter(n => n.status === 'active');
  const body = `
    <div class="form-row">
      <div class="form-group">
        <label class="form-label">個案 *</label>
        <select class="form-select" id="f-svc-case" style="width:100%">
          ${activeCases.map(c => `<option value="${c.id}">${esc(c.name)} (${esc(c.caseNo || c.id)})</option>`).join('')}
        </select>
      </div>
      <div class="form-group">
        <label class="form-label">服務日期 *</label>
        <input type="date" class="form-input" id="f-svc-date" value="${fmt(new Date())}" style="width:100%">
      </div>
    </div>
    <div class="form-row">
      <div class="form-group">
        <label class="form-label">服務形式 *</label>
        <select class="form-select" id="f-svc-type" style="width:100%">
          <option value="home">家訪</option>
          <option value="phone">電訪</option>
          <option value="video">遠距視訊</option>
        </select>
      </div>
      <div class="form-group">
        <label class="form-label">個管師</label>
        <select class="form-select" id="f-svc-nurse" style="width:100%">
          ${nurses.map(n => `<option value="${n.id}">${esc(n.name)}</option>`).join('')}
        </select>
      </div>
    </div>
    <div class="form-row">
      <div class="form-group">
        <label class="form-label">受訪者</label>
        <select class="form-select" id="f-svc-respondent" style="width:100%">
          <option value="patient">個案本人</option>
          <option value="spouse">配偶</option>
          <option value="son">兒子</option><option value="daughter">女兒</option>
          <option value="caregiver_foreign">外籍看護</option>
          <option value="caregiver_local">本國看護</option>
        </select>
      </div>
      <div class="form-group">
        <label class="form-label">是否為高血脂患者</label>
        <select class="form-select" id="f-svc-hyperlipid" style="width:100%">
          <option value="0">否</option><option value="1">是</option>
        </select>
      </div>
    </div>
    <h4 style="margin:1rem 0 .5rem;color:var(--gray-600);font-size:.9rem">服務內容</h4>
    <div class="form-row">
      <div class="form-group">
        <label class="form-checkbox"><input type="checkbox" id="f-svc-bp" checked> 測量血壓</label>
      </div>
      <div class="form-group">
        <label class="form-label">收縮壓/舒張壓</label>
        <div style="display:flex;gap:.25rem">
          <input type="number" class="form-input" id="f-svc-sys" placeholder="SYS" style="width:50%">
          <input type="number" class="form-input" id="f-svc-dia" placeholder="DIA" style="width:50%">
        </div>
      </div>
    </div>
    <div class="form-row">
      <div class="form-group">
        <label class="form-checkbox"><input type="checkbox" id="f-svc-hba1c"> 監測糖化血紅素</label>
      </div>
      <div class="form-group">
        <label class="form-checkbox"><input type="checkbox" id="f-svc-lipid"> 監測血脂</label>
      </div>
    </div>
    <div class="form-row">
      <div class="form-group">
        <label class="form-checkbox"><input type="checkbox" id="f-svc-edu" checked> 提供衛教指導</label>
      </div>
      <div class="form-group">
        <label class="form-checkbox"><input type="checkbox" id="f-svc-chronic"> 評估慢性病控制</label>
      </div>
    </div>
    <div class="form-row">
      <div class="form-group">
        <label class="form-checkbox"><input type="checkbox" id="f-svc-acp"> 完成ACP及AD說明</label>
      </div>
      <div class="form-group">
        <label class="form-checkbox"><input type="checkbox" id="f-svc-referral"> 轉介長照/醫療</label>
      </div>
    </div>
    <div class="form-group">
      <label class="form-label">備註</label>
      <textarea class="form-input" id="f-svc-notes" rows="2" style="width:100%"></textarea>
    </div>
  `;
  const footer = `
    <button class="btn btn-outline" onclick="closeModal()">取消</button>
    <button class="btn btn-primary" onclick="saveService()">儲存紀錄</button>
  `;
  openModal('新增服務紀錄', body, footer);
}

function saveService() {
  const caseId = document.getElementById('f-svc-case').value;
  const date = document.getElementById('f-svc-date').value;
  const type = document.getElementById('f-svc-type').value;
  if (!caseId || !date) { showToast('請填寫必要欄位','warning'); return; }

  const c = findCase(db, caseId);

  // 驗證: 收案後第一次需家訪
  const existingSvc = getServicesByCase(db, caseId);
  if (existingSvc.length === 0 && type !== 'home') {
    showToast('收案後第一次服務需為家訪','danger');
    return;
  }

  const bpChecked = document.getElementById('f-svc-bp').checked;
  const newId = 'S' + String(db.services.length + 1).padStart(5, '0');
  const svc = {
    id: newId, caseId,
    nurseId: document.getElementById('f-svc-nurse').value,
    doctorId: c ? c.doctorId : '',
    date, type,
    respondent: document.getElementById('f-svc-respondent').value,
    bpMeasured: bpChecked,
    bpSystolic: bpChecked ? parseInt(document.getElementById('f-svc-sys').value) || null : null,
    bpDiastolic: bpChecked ? parseInt(document.getElementById('f-svc-dia').value) || null : null,
    hba1cMonitored: document.getElementById('f-svc-hba1c').checked,
    hba1cValue: null,
    lipidMonitored: document.getElementById('f-svc-lipid').checked,
    educationProvided: document.getElementById('f-svc-edu').checked,
    acpExplained: document.getElementById('f-svc-acp').checked,
    chronicDiseaseEval: document.getElementById('f-svc-chronic').checked,
    referralLtc: document.getElementById('f-svc-referral').checked,
    referralMedical: false,
    notes: document.getElementById('f-svc-notes').value,
    billingStatus: 'pending',
  };
  db.services.push(svc);

  // 自動產生費用申報
  const isRemote = c && c.isRemoteArea;
  let code, amount;
  if (type === 'home') {
    code = isRemote ? 'YA04' : 'YA02';
    amount = isRemote ? 1200 : 1000;
  } else {
    code = isRemote ? 'YA03' : 'YA01';
    amount = isRemote ? 300 : 250;
  }
  db.billings.push({
    id: 'B' + String(db.billings.length + 1).padStart(5, '0'),
    caseId, code, serviceDate: date, billingMonth: date.substring(0, 7),
    memberId: svc.nurseId, amount, status: 'pending',
  });

  // 更新 ACP 狀態
  if (svc.acpExplained && c) {
    c.acpExplained = true;
    c.acpExplainedDate = c.acpExplainedDate || date;
  }

  saveDB(db);
  closeModal();
  showToast('服務紀錄已儲存，費用已自動產生');
  renderServices();
}

function viewService(svcId) {
  const s = db.services.find(x => x.id === svcId);
  if (!s) return;
  const c = findCase(db, s.caseId);
  const n = findMember(db, s.nurseId);
  const body = `
    <table style="font-size:.85rem;width:100%">
      <tr><td style="padding:4px 8px;color:var(--gray-500);width:120px">個案</td><td>${c?esc(c.name):esc(s.caseId)}</td></tr>
      <tr><td style="padding:4px 8px;color:var(--gray-500)">日期</td><td>${s.date}</td></tr>
      <tr><td style="padding:4px 8px;color:var(--gray-500)">形式</td><td>${serviceTypeLabel(s.type)}</td></tr>
      <tr><td style="padding:4px 8px;color:var(--gray-500)">個管師</td><td>${n?esc(n.name):'-'}</td></tr>
      <tr><td style="padding:4px 8px;color:var(--gray-500)">受訪者</td><td>${respondentLabel(s.respondent)}</td></tr>
      <tr><td style="padding:4px 8px;color:var(--gray-500)">血壓</td><td>${s.bpMeasured?s.bpSystolic+'/'+s.bpDiastolic+' mmHg':'未測量'}</td></tr>
      <tr><td style="padding:4px 8px;color:var(--gray-500)">糖化血紅素</td><td>${s.hba1cMonitored?(s.hba1cValue||'已測'):'未測'}</td></tr>
      <tr><td style="padding:4px 8px;color:var(--gray-500)">血脂</td><td>${s.lipidMonitored?'已監測':'未測'}</td></tr>
      <tr><td style="padding:4px 8px;color:var(--gray-500)">衛教指導</td><td>${s.educationProvided?'✅':'❌'}</td></tr>
      <tr><td style="padding:4px 8px;color:var(--gray-500)">ACP說明</td><td>${s.acpExplained?'✅':'❌'}</td></tr>
      <tr><td style="padding:4px 8px;color:var(--gray-500)">慢性病評估</td><td>${s.chronicDiseaseEval?'✅':'❌'}</td></tr>
      <tr><td style="padding:4px 8px;color:var(--gray-500)">轉介</td><td>${s.referralLtc?'長照 ':''} ${s.referralMedical?'醫療':''} ${!s.referralLtc&&!s.referralMedical?'無':''}</td></tr>
      <tr><td style="padding:4px 8px;color:var(--gray-500)">申報狀態</td><td>${billingStatusBadge(s.billingStatus)}</td></tr>
      <tr><td style="padding:4px 8px;color:var(--gray-500)">備註</td><td>${esc(s.notes||'-')}</td></tr>
    </table>
  `;
  openModal('服務紀錄詳情', body, '<button class="btn btn-outline" onclick="closeModal()">關閉</button>');
}

// ==========================================
//  醫師意見書
// ==========================================
function renderOpinions() {
  const search = (document.getElementById('opinion-search')?.value || '').toLowerCase();
  const status = document.getElementById('opinion-filter-status')?.value || '';
  const now = new Date();
  const todayStr = fmt(now);

  // 基於 doctorVisitDate 計算每個個案的意見書狀態
  const activeCases = getActiveCases(db);
  const caseOpStatus = activeCases.map(c => {
    const op = getLatestOpinion(db, c.id);
    const lastDate = c.doctorVisitDate || (op ? op.issueDate : '');
    if (!lastDate) return { c, op, status: 'pending', lastDate: '', expiryDate: '', daysLeft: -1 };
    const daysSince = daysBetween(lastDate, todayStr);
    const expD = new Date(lastDate); expD.setMonth(expD.getMonth() + 6);
    const expiryDate = expD.toISOString().slice(0, 10);
    const daysLeft = daysBetween(todayStr, expiryDate);
    let st = 'valid';
    if (daysSince > 180) st = 'expired';
    else if (daysSince >= 150) st = 'expiring';
    return { c, op, status: st, lastDate, expiryDate, daysLeft };
  });

  const valid = caseOpStatus.filter(x => x.status === 'valid').length;
  const expiring = caseOpStatus.filter(x => x.status === 'expiring').length;
  const expired = caseOpStatus.filter(x => x.status === 'expired').length;
  const pending = caseOpStatus.filter(x => x.status === 'pending').length;

  const _cur = status || 'all';
  const btnStyle = (key) => `cursor:pointer;transition:all .15s;${_cur === key ? 'outline:3px solid var(--primary);outline-offset:2px;transform:scale(1.03);' : ''}`;
  document.getElementById('opinion-summary').innerHTML = `
    <div class="summary-item" style="${btnStyle('all')}" onclick="document.getElementById('opinion-filter-status').value='';renderOpinions()"><div class="summary-value">${activeCases.length}</div><div class="summary-label">總意見書數</div></div>
    <div class="summary-item" style="${btnStyle('valid')}" onclick="document.getElementById('opinion-filter-status').value='valid';renderOpinions()"><div class="summary-value" style="color:var(--success)">${valid}</div><div class="summary-label">有效</div></div>
    <div class="summary-item" style="${btnStyle('expiring')}" onclick="document.getElementById('opinion-filter-status').value='expiring';renderOpinions()"><div class="summary-value" style="color:var(--warning)">${expiring}</div><div class="summary-label">即將到期(30天內)</div></div>
    <div class="summary-item" style="${btnStyle('expired')}" onclick="document.getElementById('opinion-filter-status').value='expired';renderOpinions()"><div class="summary-value" style="color:var(--danger)">${expired}</div><div class="summary-label">已過期</div></div>
    <div class="summary-item" style="${btnStyle('pending')}" onclick="document.getElementById('opinion-filter-status').value='pending';renderOpinions()"><div class="summary-value" style="color:var(--gray-500)">${pending}</div><div class="summary-label">待開立</div></div>
  `;

  let filtered = caseOpStatus.filter(x => {
    if (status && x.status !== status) return false;
    if (search && !x.c.name.toLowerCase().includes(search)) return false;
    return true;
  }).sort((a, b) => {
    // 逾期和待開立排前面
    const order = { expired: 0, expiring: 1, pending: 2, valid: 3 };
    if (order[a.status] !== order[b.status]) return order[a.status] - order[b.status];
    return (a.expiryDate || '').localeCompare(b.expiryDate || '');
  });

  document.getElementById('opinion-tbody').innerHTML = filtered.map(x => {
    const c = x.c;
    const rowClass = x.status === 'expired' ? 'row-danger' : (x.status === 'expiring' ? 'row-warning' : (x.status === 'pending' ? 'row-gray' : ''));
    const statusBadge = x.status === 'expired' ? '<span class="badge badge-danger">逾期</span>'
      : x.status === 'expiring' ? '<span class="badge badge-warning">待更新</span>'
      : x.status === 'pending' ? '<span class="badge badge-gray">待開立</span>'
      : '<span class="badge badge-success">有效</span>';
    const daysText = x.status === 'pending' ? '-' : (x.daysLeft > 0 ? x.daysLeft + '天' : '已過期 ' + Math.abs(x.daysLeft) + '天');
    return `<tr class="${rowClass}">
      <td>${esc(c.name)}</td>
      <td>${esc(c.doctorName || '-')}</td>
      <td>${x.lastDate || '-'}</td>
      <td>${x.op ? '第' + x.op.sequence + '次' : '-'}</td>
      <td>${x.lastDate || '-'}</td>
      <td>${x.expiryDate || '-'}</td>
      <td><strong>${daysText}</strong></td>
      <td>${x.op ? x.op.diseaseStatus : '-'}</td>
      <td>-</td>
      <td>${statusBadge}</td>
      <td><button class="btn btn-xs btn-primary" onclick="renewOpinion('${c.id}')">更新</button></td>
    </tr>`;
  }).join('');
}

function openOpinionModal() {
  const activeCases = getActiveCases(db);
  const docs = getDoctors(db).filter(d => d.status === 'active');
  const body = `
    <div class="form-row">
      <div class="form-group">
        <label class="form-label">個案 *</label>
        <select class="form-select" id="f-op-case" style="width:100%">
          ${activeCases.map(c => `<option value="${c.id}">${esc(c.name)} (${esc(c.caseNo || c.id)})</option>`).join('')}
        </select>
      </div>
      <div class="form-group">
        <label class="form-label">開立醫師 *</label>
        <select class="form-select" id="f-op-doctor" style="width:100%">
          ${docs.map(d => `<option value="${d.id}">${esc(d.name)}</option>`).join('')}
        </select>
      </div>
    </div>
    <div class="form-row">
      <div class="form-group">
        <label class="form-label">家訪日期 *</label>
        <input type="date" class="form-input" id="f-op-visit" value="${fmt(new Date())}" style="width:100%">
      </div>
      <div class="form-group">
        <label class="form-label">開立日期 *</label>
        <input type="date" class="form-input" id="f-op-date" value="${fmt(new Date())}" style="width:100%">
      </div>
    </div>
    <div class="form-row">
      <div class="form-group">
        <label class="form-label">病情狀態</label>
        <select class="form-select" id="f-op-status" style="width:100%">
          <option value="穩定">穩定</option>
          <option value="不穩定">不穩定</option>
          <option value="不明">不明</option>
        </select>
      </div>
      <div class="form-group">
        <label class="form-label">功能預後</label>
        <select class="form-select" id="f-op-prognosis" style="width:100%">
          <option value="穩定">穩定</option>
          <option value="進步">進步</option>
          <option value="退步">退步</option>
          <option value="無法確定">無法確定</option>
        </select>
      </div>
    </div>
    <div class="form-hint">※ 醫師意見書需於收案後14天(日曆天)內完成，每6個月更新一次，每年上限2次</div>
  `;
  const footer = `
    <button class="btn btn-outline" onclick="closeModal()">取消</button>
    <button class="btn btn-primary" onclick="saveOpinion()">開立意見書</button>
  `;
  openModal('開立醫師意見書', body, footer);
}

function saveOpinion() {
  const caseId = document.getElementById('f-op-case').value;
  const doctorId = document.getElementById('f-op-doctor').value;
  const issueDate = document.getElementById('f-op-date').value;
  const visitDate = document.getElementById('f-op-visit').value;

  // 驗證醫師訓練
  const doc = findMember(db, doctorId);
  if (doc && !doc.opinionTrained) {
    showToast('該醫師尚未完成意見書訓練課程','danger');
    return;
  }

  // 年度次數檢查
  const yy = issueDate.substring(0, 4);
  const yearOps = db.opinions.filter(o => o.caseId === caseId && o.issueDate.startsWith(yy));
  if (yearOps.length >= 2) {
    showToast('該個案本年度已達意見書上限(2次)','danger');
    return;
  }

  const expDate = new Date(issueDate);
  expDate.setMonth(expDate.getMonth() + 6);

  const newId = 'OP' + String(db.opinions.length + 1).padStart(4, '0');
  db.opinions.push({
    id: newId, caseId, doctorId, issueDate, homeVisitDate: visitDate,
    expiryDate: fmt(expDate), sequence: yearOps.length + 1,
    yearCount: yearOps.length + 1,
    diseaseStatus: document.getElementById('f-op-status').value,
    functionalPrognosis: document.getElementById('f-op-prognosis').value,
    status: 'valid',
  });

  // 費用
  const c = findCase(db, caseId);
  const isRemote = c && c.isRemoteArea;
  db.billings.push({
    id: 'B' + String(db.billings.length + 1).padStart(5, '0'),
    caseId, code: 'AA12', serviceDate: issueDate, billingMonth: issueDate.substring(0, 7),
    memberId: doctorId, amount: isRemote ? 1800 : 1500, status: 'pending',
  });

  saveDB(db);
  closeModal();
  showToast('醫師意見書已開立，費用已自動產生');
  renderOpinions();
}

function renewOpinion(caseId) {
  openOpinionModal();
  setTimeout(() => {
    const sel = document.getElementById('f-op-case');
    if (sel) sel.value = caseId;
  }, 100);
}

// ==========================================
//  費用申報
// ==========================================
function renderBilling() {
  const month = document.getElementById('billing-month')?.value || '';
  const code = document.getElementById('billing-filter-code')?.value || '';
  const status = document.getElementById('billing-filter-status')?.value || '';

  let filtered = db.billings.filter(b => {
    if (month && b.billingMonth !== month) return false;
    if (code && b.code !== code) return false;
    if (status && b.status !== status) return false;
    return true;
  }).sort((a, b) => b.serviceDate.localeCompare(a.serviceDate));

  // Summary
  const totalAmt = filtered.reduce((s, b) => s + b.amount, 0);
  const pending = filtered.filter(b => b.status === 'pending');
  const submitted = filtered.filter(b => b.status === 'submitted');
  const approved = filtered.filter(b => b.status === 'approved');
  const rejected = filtered.filter(b => b.status === 'rejected');

  document.getElementById('billing-summary').innerHTML = `
    <div class="summary-item"><div class="summary-value">${filtered.length}</div><div class="summary-label">總筆數</div></div>
    <div class="summary-item"><div class="summary-value" style="color:var(--primary)">$${totalAmt.toLocaleString()}</div><div class="summary-label">總金額</div></div>
    <div class="summary-item"><div class="summary-value" style="color:var(--warning)">${pending.length}</div><div class="summary-label">待申報</div></div>
    <div class="summary-item"><div class="summary-value" style="color:var(--info)">${submitted.length}</div><div class="summary-label">已申報</div></div>
    <div class="summary-item"><div class="summary-value" style="color:var(--success)">$${approved.reduce((s,b)=>s+b.amount,0).toLocaleString()}</div><div class="summary-label">已核付金額</div></div>
    <div class="summary-item"><div class="summary-value" style="color:var(--danger)">${rejected.length}</div><div class="summary-label">退件</div></div>
  `;

  document.getElementById('billing-tbody').innerHTML = filtered.slice(0, 100).map(b => {
    const c = findCase(db, b.caseId);
    const m = findMember(db, b.memberId);
    const displayName = c ? esc(c.name) : (b.caseName ? esc(b.caseName) : esc(b.caseId));
    const displayMember = m ? esc(m.name) : (b.doctorName ? esc(b.doctorName) : '-');
    return `<tr>
      <td><input type="checkbox" class="billing-cb" data-id="${b.id}" ${b.status==='pending'?'':'disabled'}></td>
      <td>${b.billingMonth}</td>
      <td>${displayName}</td>
      <td><span class="badge badge-info">${b.code}</span></td>
      <td>${b.serviceDate}</td>
      <td>${displayMember}</td>
      <td>$${b.amount.toLocaleString()}</td>
      <td>${billingStatusBadge(b.status)}</td>
      <td>${b.status === 'pending' ? `<button class="btn btn-xs btn-primary" onclick="submitSingleBilling('${b.id}')">申報</button>` : ''}</td>
    </tr>`;
  }).join('');

  document.getElementById('billing-total').innerHTML = `
    <span>篩選金額合計：<strong>$${totalAmt.toLocaleString()}</strong></span>
    <span>待申報金額：<strong style="color:var(--warning)">$${pending.reduce((s,b)=>s+b.amount,0).toLocaleString()}</strong></span>
  `;
}

function toggleBillingAll(el) {
  document.querySelectorAll('.billing-cb:not(:disabled)').forEach(cb => { cb.checked = el.checked; });
}

function submitSingleBilling(id) {
  const b = db.billings.find(x => x.id === id);
  if (b) { b.status = 'submitted'; saveDB(db); renderBilling(); showToast('已申報'); }
}

function submitBilling() {
  const checked = document.querySelectorAll('.billing-cb:checked');
  if (checked.length === 0) { showToast('請勾選要申報的項目', 'warning'); return; }
  checked.forEach(cb => {
    const b = db.billings.find(x => x.id === cb.dataset.id);
    if (b && b.status === 'pending') b.status = 'submitted';
  });
  saveDB(db);
  renderBilling();
  showToast(`已批次申報 ${checked.length} 筆`);
}

function generateBilling() {
  showToast('申報清冊已產生，可下載 CSV 檔案');
  // 實際系統中會產生 CSV/Excel
  const month = document.getElementById('billing-month')?.value || fmt(new Date()).substring(0, 7);
  const items = db.billings.filter(b => b.billingMonth === month);
  let csv = '\uFEFF申報月份,個案編號,個案姓名,組合代碼,服務日期,服務人員,金額,狀態\n';
  items.forEach(b => {
    const c = findCase(db, b.caseId);
    const m = findMember(db, b.memberId);
    csv += `${escCSV(b.billingMonth)},${escCSV(b.caseId)},${c?escCSV(c.name):''},${escCSV(b.code)},${escCSV(b.serviceDate)},${m?escCSV(m.name):''},${b.amount},${escCSV(b.status)}\n`;
  });
  downloadCSV(csv, `申報清冊_${month}.csv`);
}

// ==========================================
//  匯入核銷清冊 / 產生核銷總表
// ==========================================
function importBillingFile(input) {
  const file = input.files[0];
  if (!file) return;
  input.value = '';
  const reader = new FileReader();
  reader.onload = function(ev) {
    try {
      const workbook = XLSX.read(ev.target.result, { type: 'array' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
      // 找到標題列
      let headerIdx = -1;
      for (let i = 0; i < Math.min(aoa.length, 5); i++) {
        const row = aoa[i].map(h => String(h).replace(/[\r\n\s]/g, ''));
        if (row.includes('服務代碼') && row.includes('身分證號')) { headerIdx = i; break; }
      }
      if (headerIdx < 0) { showToast('無法辨識核銷清冊格式，需包含「服務代碼」和「身分證號」欄位', 'danger'); return; }
      const headers = aoa[headerIdx].map(h => String(h).replace(/[\r\n\s]/g, ''));
      const col = (name) => headers.indexOf(name);

      function rocToAD(rocStr) {
        if (!rocStr) return '';
        const s = String(rocStr).trim();
        const m = s.match(/^(\d{2,3})\/(\d{1,2})\/(\d{1,2})/);
        if (!m) return '';
        return `${parseInt(m[1])+1911}-${m[2].padStart(2,'0')}-${m[3].padStart(2,'0')}`;
      }
      function extractDate(plan) {
        const s = String(plan).trim();
        const m = s.match(/^(\d{2,3})\/(\d{1,2})\/(\d{1,2})/);
        if (m) return `${parseInt(m[1])+1911}-${m[2].padStart(2,'0')}-${m[3].padStart(2,'0')}`;
        return '';
      }

      let imported = 0, skipped = 0, duplicated = 0;
      const maxId = db.billings.reduce((mx, b) => {
        const n = parseInt((b.id || '').replace('B', ''));
        return n > mx ? n : mx;
      }, 0);

      for (let i = headerIdx + 1; i < aoa.length; i++) {
        const row = aoa[i];
        const code = String(row[col('服務代碼')] || '').trim();
        if (!code || code === '總計') continue;
        const idNumber = String(row[col('身分證號')] || '').trim().toUpperCase();
        if (!idNumber) { skipped++; continue; }
        const caseName = String(row[col('個案姓名')] || '').trim();
        const doctorName = String(row[col('醫師/個管員')] || '').trim();
        const plan = String(row[col('採用計畫')] || '');
        const serviceDate = extractDate(plan);
        const cmsLevel = String(row[col('CMS等級')] || '').trim();
        const serviceType = String(row[col('服務項目類別')] || '').trim();
        const amount = parseFloat(row[col('小計')]) || 0;
        const qty = parseFloat(row[col('數量')]) || 1;
        const unitPrice = parseFloat(row[col('給付價格')]) || 0;
        const district = String(row[col('居住行政區')] || '').trim();
        const city = String(row[col('居住縣市')] || '').trim();
        const careManager = String(row[col('照管專員')] || '').trim();

        // 找對應個案
        const caseMatch = db.cases.find(c => c.idNumber && c.idNumber.toUpperCase() === idNumber);
        const caseId = caseMatch ? caseMatch.id : '';
        // 找醫師/個管
        const memberMatch = db.members.find(m => m.name === doctorName);
        const memberId = memberMatch ? memberMatch.id : '';

        // 計算申報月份 (從計畫日期推算，或用服務日期)
        const billingMonth = serviceDate ? serviceDate.substring(0, 7) : '';

        // 重複檢查
        const isDup = db.billings.some(b =>
          b.code === code && b.serviceDate === serviceDate &&
          ((caseId && b.caseId === caseId) || b.idNumber === idNumber)
        );
        if (isDup) { duplicated++; continue; }

        imported++;
        db.billings.push({
          id: 'B' + String(maxId + imported).padStart(5, '0'),
          caseId, code, serviceDate, billingMonth,
          memberId, amount, status: 'pending',
          // 核銷清冊額外欄位
          idNumber, caseName, doctorName, cmsLevel, serviceType,
          unitPrice, qty, district, city, careManager, plan,
          source: 'lcms_billing'
        });
      }

      saveDB(db);
      renderBilling();
      showToast(`核銷清冊匯入完成：新增 ${imported} 筆${duplicated ? `，重複跳過 ${duplicated} 筆` : ''}${skipped ? `，無效跳過 ${skipped} 筆` : ''}`);
    } catch (err) {
      console.error('核銷清冊匯入失敗:', err);
      showToast('核銷清冊匯入失敗: ' + err.message, 'danger');
    }
  };
  reader.readAsArrayBuffer(file);
}

function exportBillingSummary() {
  const month = document.getElementById('billing-month')?.value || '';
  if (!month) { showToast('請先選擇申報月份', 'warning'); return; }
  const items = db.billings.filter(b => b.billingMonth === month);
  if (items.length === 0) { showToast('該月份無申報資料', 'warning'); return; }

  // 依服務代碼分類統計
  const codeMap = {};
  items.forEach(b => {
    if (!codeMap[b.code]) codeMap[b.code] = { count: 0, total: 0 };
    codeMap[b.code].count += 1;
    codeMap[b.code].total += b.amount;
  });

  // 總表欄位對照
  const t06 = 0; // AA01~AA02 照顧組合
  const t07 = (codeMap['AA03']?.total||0)+(codeMap['AA04']?.total||0)+(codeMap['AA05']?.total||0)
             +(codeMap['AA06']?.total||0)+(codeMap['AA07']?.total||0)+(codeMap['AA08']?.total||0)
             +(codeMap['AA09']?.total||0)+(codeMap['AA10']?.total||0)+(codeMap['AA11']?.total||0)
             +(codeMap['AA12']?.total||0); // 政策鼓勵 (含AA12醫師意見書)
  const t08 = 0; // BA 居家照顧
  const t09 = 0; // BB 日間照顧
  const t10 = 0; // BC 家庭托顧
  const t11 = 0; // BD 社區式
  const t12 = 0; // C 專業服務
  const t13 = 0; // D 交通接送
  const t14 = 0; // E 輔具
  const t15 = 0; // F 無障礙
  const t16 = 0; // G 喘息
  const t17 = t06+t07+t08+t09+t10+t11+t12+t13+t14+t15+t16; // 申報費用(含部分負擔)
  const t18 = 0; // 部分負擔
  const t19 = t17 - t18; // 申請費用
  const t20 = 0; // 膳費
  const t21 = 0; // 縣市補助
  const t22 = (codeMap['YA01']?.total||0)+(codeMap['YA02']?.total||0)
             +(codeMap['YA03']?.total||0)+(codeMap['YA04']?.total||0); // 其他服務
  const t23 = t20+t21+t22; // 非照顧組合小計
  const t24 = t19+t23; // 總計

  const [yyyy, mm] = month.split('-');
  const rocYear = parseInt(yyyy) - 1911;
  const monthStr = `${rocYear}年${mm}月`;

  // 產生 Excel 格式核銷總表
  const wb = XLSX.utils.book_new();
  const data = [
    ['特約長照服務提供者服務費用申報總表'],
    [],
    ['服務提供者', '敏盛綜合醫院', '', '費用年月', `${rocYear}年 ${mm}月`, '', '申報方式', '□書面 □網路 ■媒體'],
    ['申報類別', '■送核 □補報', '', '發文日期', '', '', '收文日期', ''],
    [],
    ['服務項目類別', '', '申報費用(元)', '筆數', '備註'],
    ['照顧組合', 'A碼 照顧管理', t06, '', 'AA01~AA02'],
    ['', '政策鼓勵', t07, codeMap['AA12']?.count||0, 'AA03~AA12 (含醫師意見書)'],
    ['', 'B碼 居家照顧服務', t08, '', 'BA01~BA22'],
    ['', '日間照顧服務', t09, '', 'BB01~BB14'],
    ['', '家庭托顧服務', t10, '', 'BC01~BC14'],
    ['', '社區式照顧', t11, '', 'BD01~BD03'],
    ['', 'C碼 專業服務', t12, '', ''],
    ['', 'D碼 交通接送服務', t13, '', ''],
    ['', 'E碼 輔具服務', t14, '', ''],
    ['', 'F碼 居家無障礙', t15, '', ''],
    ['', 'G碼 喘息服務', t16, '', ''],
    ['申報費用(含部分負擔)', '', t17, '', 't06~t16加總'],
    ['僅部分負擔費用', '', t18, '', ''],
    ['申請(補助)費用 (t17-t18)', '', t19, '', ''],
    [],
    ['非照顧組合', '營養餐飲(膳費)', t20, '', ''],
    ['', '縣市政府補助', t21, '', ''],
    ['', '其他服務', t22, (codeMap['YA01']?.count||0)+(codeMap['YA02']?.count||0)+(codeMap['YA03']?.count||0)+(codeMap['YA04']?.count||0), 'YA01~YA04'],
    ['非照顧組合小計', '', t23, '', 't20~t22加總'],
    [],
    ['總計(系統計算)', '', t24, items.length, 't19+t23'],
    [],
    ['本次申報起迄日期', `${rocYear}年${mm}月01日`, '~', `${rocYear}年${mm}月${new Date(parseInt(yyyy), parseInt(mm), 0).getDate()}日`],
    [],
    ['=== 各代碼明細 ==='],
  ];

  // 加入各代碼明細統計
  Object.keys(codeMap).sort().forEach(code => {
    const info = codeMap[code];
    const label = billingCodeLabel(code);
    data.push([code, label, info.total, info.count, `均價 $${Math.round(info.total/info.count)}`]);
  });

  const ws = XLSX.utils.aoa_to_sheet(data);
  // 設定欄寬
  ws['!cols'] = [{wch:20},{wch:22},{wch:15},{wch:8},{wch:25}];
  // 合併標題列
  ws['!merges'] = [{s:{r:0,c:0},e:{r:0,c:4}}];
  XLSX.utils.book_append_sheet(wb, ws, '核銷總表');

  // 加入核銷清冊 sheet
  const listData = [['服務代碼','醫師/個管員','採用計畫','CMS等級','服務項目類別','身分證號','個案姓名','給付價格','數量','小計','居住縣市','居住行政區','照管專員']];
  items.forEach(b => {
    const c = findCase(db, b.caseId);
    listData.push([
      b.code, b.doctorName || (db.members.find(m=>m.id===b.memberId)?.name||''),
      b.plan || '', b.cmsLevel || (c?.cmsLevel||''),
      b.serviceType || billingCodeLabel(b.code),
      b.idNumber || (c?.idNumber||''), b.caseName || (c?.name||''),
      b.unitPrice || b.amount, b.qty || 1, b.amount,
      b.city || (c?.city||''), b.district || (c?.district||''),
      b.careManager || (c?.careManager||'')
    ]);
  });
  listData.push(['總計','','','','','','','','', items.reduce((s,b)=>s+b.amount,0),'','','']);
  const ws2 = XLSX.utils.aoa_to_sheet(listData);
  ws2['!cols'] = [{wch:10},{wch:10},{wch:30},{wch:8},{wch:25},{wch:14},{wch:10},{wch:10},{wch:6},{wch:10},{wch:8},{wch:8},{wch:8}];
  XLSX.utils.book_append_sheet(wb, ws2, '核銷清冊');

  XLSX.writeFile(wb, `核銷總表_${month}.xlsx`);
  showToast(`已產生 ${monthStr} 核銷總表 (含清冊)，共 ${items.length} 筆，總金額 $${t24.toLocaleString()}`);
}

// ==========================================
//  個案地圖
// ==========================================
let caseMapInstance = null;
let caseMapMarkers = [];
const DISTRICT_COORDS = {
  '桃園區':[24.9936,121.3010],'中壢區':[24.9656,121.2249],'平鎮區':[24.9457,121.2183],
  '八德區':[24.9527,121.2855],'楊梅區':[24.9077,121.1455],'蘆竹區':[25.0457,121.2920],
  '龜山區':[25.0335,121.3457],'龍潭區':[24.8635,121.2165],'大溪區':[24.8833,121.2873],
  '大園區':[25.0629,121.1975],'觀音區':[25.0335,121.0835],'新屋區':[24.9721,121.1062],
  '復興區':[24.8208,121.3530]
};
const CMS_COLORS = {
  '第1級':'#4CAF50','第2級':'#8BC34A','第3級':'#CDDC39','第4級':'#FFC107',
  '第5級':'#FF9800','第6級':'#FF5722','第7級':'#E91E63','第8級':'#9C27B0'
};

// 地理編碼快取 (存 localStorage)
function getGeoCache() {
  try { return JSON.parse(localStorage.getItem('geocache') || '{}'); } catch { return {}; }
}
function setGeoCache(cache) {
  localStorage.setItem('geocache', JSON.stringify(cache));
}

function initCaseMap() {
  const container = document.getElementById('casemap-container');
  if (!container || caseMapInstance) return;
  caseMapInstance = L.map(container, { zoomControl: true }).setView([24.99, 121.25], 12);
  L.tileLayer('https://{s}.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}{r}.png', {
    attribution: '&copy; <a href="https://openstreetmap.org">OSM</a> &copy; <a href="https://carto.com/">CARTO</a>',
    subdomains: 'abcd', maxZoom: 19
  }).addTo(caseMapInstance);

  // 圖例
  const legend = L.control({ position: 'bottomright' });
  legend.onAdd = function() {
    const div = L.DomUtil.create('div', 'leaflet-control');
    div.style.cssText = 'background:#fff;padding:8px 12px;border-radius:8px;box-shadow:0 2px 6px rgba(0,0,0,0.2);font-size:0.8rem;line-height:1.8';
    div.innerHTML = '<strong>CMS 等級</strong><br>' +
      Object.entries(CMS_COLORS).map(([k,v]) =>
        `<span style="display:inline-block;width:12px;height:12px;border-radius:50%;background:${v};margin-right:4px;vertical-align:middle"></span>${k}`
      ).join('<br>');
    return div;
  };
  legend.addTo(caseMapInstance);
}

function renderCaseMap() {
  const active = getActiveCases(db);
  // 填充篩選器
  const doctorSel = document.getElementById('casemap-filter-doctor');
  const districtSel = document.getElementById('casemap-filter-district');
  const cmsSel = document.getElementById('casemap-filter-cms');
  if (doctorSel && doctorSel.options.length <= 1) {
    const doctors = getDoctors(db).filter(d => d.status === 'active');
    doctors.forEach(d => { doctorSel.add(new Option(d.name, d.id)); });
  }
  if (districtSel && districtSel.options.length <= 1) {
    const districts = [...new Set(active.map(c => c.district).filter(Boolean))].sort();
    districts.forEach(d => { districtSel.add(new Option(d, d)); });
  }
  if (cmsSel && cmsSel.options.length <= 1) {
    for (let i = 1; i <= 8; i++) cmsSel.add(new Option(`第${i}級`, `第${i}級`));
  }

  const filterDoctor = doctorSel?.value || '';
  const filterDistrict = districtSel?.value || '';
  const filterCms = cmsSel?.value || '';

  let filtered = active.filter(c => {
    if (filterDoctor && c.doctorId !== filterDoctor) return false;
    if (filterDistrict && c.district !== filterDistrict) return false;
    if (filterCms && c.cmsLevel !== filterCms) return false;
    return true;
  });

  document.getElementById('casemap-count').textContent = `顯示 ${filtered.length} / ${active.length} 個案`;

  // 初始化地圖
  initCaseMap();
  if (!caseMapInstance) return;
  setTimeout(() => caseMapInstance.invalidateSize(), 100);

  // 清除舊 markers
  caseMapMarkers.forEach(m => caseMapInstance.removeLayer(m));
  caseMapMarkers = [];

  const geoCache = getGeoCache();
  let geocoded = 0, fallback = 0;

  filtered.forEach((c, idx) => {
    let lat, lng;
    const cacheKey = getCaseGeoAddress(c);

    // 優先使用個案已存的座標
    if (c.lat && c.lng) {
      lat = c.lat; lng = c.lng;
      geocoded++;
    }
    // 其次使用地理編碼快取
    else if (cacheKey && geoCache[cacheKey]) {
      lat = geoCache[cacheKey][0]; lng = geoCache[cacheKey][1];
      geocoded++;
    }
    // 最後退回行政區中心 + 散布
    else {
      const dist = c.district || '';
      const baseCoord = DISTRICT_COORDS[dist];
      if (!baseCoord) return;
      const angle = (idx * 2.399) + (c.id ? c.id.charCodeAt(c.id.length-1)*0.1 : 0); // golden angle spread
      const spread = 0.006;
      const r = spread * Math.sqrt((idx % 50) / 50);
      lat = baseCoord[0] + r * Math.cos(angle);
      lng = baseCoord[1] + r * Math.sin(angle);
      fallback++;
    }

    const color = CMS_COLORS[c.cmsLevel] || '#E53935';
    const doctor = findMember(db, c.doctorId);
    const isGeocoded = (c.lat && c.lng) || (cacheKey && geoCache[cacheKey]);
    const label = `${c.caseNo || c.id} ${c.name}`;

    // 圖釘 SVG icon
    const pinSvg = `<svg xmlns="http://www.w3.org/2000/svg" width="28" height="40" viewBox="0 0 28 40">
      <path d="M14 0C6.27 0 0 6.27 0 14c0 10.5 14 26 14 26s14-15.5 14-26C28 6.27 21.73 0 14 0z" fill="${color}" stroke="#fff" stroke-width="1.5"/>
      <circle cx="14" cy="13" r="5.5" fill="#fff" opacity="0.9"/>
    </svg>`;
    const pinIcon = L.divIcon({
      html: pinSvg,
      className: 'casemap-pin',
      iconSize: [28, 40],
      iconAnchor: [14, 40],
      popupAnchor: [0, -36]
    });

    const marker = L.marker([lat, lng], { icon: pinIcon, draggable: true }).addTo(caseMapInstance);

    // 永久顯示姓名標籤
    marker.bindTooltip(label, {
      permanent: true, direction: 'right', offset: [12, -20],
      className: 'casemap-label'
    });

    // 拖拉結束 → 儲存新座標
    marker.on('dragend', function() {
      const pos = marker.getLatLng();
      c.lat = pos.lat; c.lng = pos.lng;
      const gc = getGeoCache();
      const ck = getCaseGeoAddress(c);
      if (ck) gc[ck] = [pos.lat, pos.lng];
      setGeoCache(gc);
      saveDB(db);
      showToast(`${c.name} 位置已更新`);
    });

    marker.bindPopup(`
      <div style="min-width:220px;line-height:1.7;font-size:0.9rem">
        <div style="font-size:1.1rem;font-weight:700;margin-bottom:4px">${esc(c.name)}
          <span style="background:${color};color:#fff;padding:1px 8px;border-radius:4px;font-size:0.75rem;margin-left:6px">${c.cmsLevel || '-'}</span>
        </div>
        <div>📋 ${esc(c.caseNo || c.id)}</div>
        <div>🏠 ${esc(c.address || (c.district ? '桃園市' + c.district + (c.village||'') : '-'))}</div>
        <div>👨‍⚕️ ${doctor ? esc(doctor.name) : (c.doctorName || '-')}</div>
        <div>📅 收案 ${c.enrollDate || '-'}</div>
        ${c.phone ? `<div>📞 ${esc(c.phone)}</div>` : ''}
        ${c.contactPerson ? `<div>👤 ${esc(c.contactPerson)}</div>` : ''}
        <div style="margin-top:6px;color:#999;font-size:0.75rem">💡 可拖拉圖釘修正位置</div>
      </div>
    `);
    marker.on('mouseover', function() { this.openPopup(); });
    marker.on('mouseout', function() { this.closePopup(); });
    caseMapMarkers.push(marker);
  });

  // 更新地理編碼狀態
  const bar = document.getElementById('casemap-geocode-bar');
  if (bar) {
    bar.style.display = 'flex';
    document.getElementById('casemap-geocode-status').innerHTML =
      `<span class="badge badge-success">已定位 ${geocoded}</span> ` +
      `<span class="badge badge-gray">約略位置 ${fallback}</span> ` +
      (fallback > 0 ? '<span style="color:#999">← 點「📍 地址定位」可精確定位</span>' : '<span style="color:#4CAF50">✓ 全部已精確定位</span>');
  }

  // fit bounds
  if (caseMapMarkers.length > 0) {
    const group = L.featureGroup(caseMapMarkers);
    caseMapInstance.fitBounds(group.getBounds().pad(0.1));
  }
}

// 在地圖上新增單一個案圖釘
function addCasePin(c) {
  if (!caseMapInstance || !c.lat || !c.lng) return;
  const color = CMS_COLORS[c.cmsLevel] || '#E53935';
  const doctor = findMember(db, c.doctorId);
  const label = `${c.caseNo || c.id} ${c.name}`;
  const pinSvg = `<svg xmlns="http://www.w3.org/2000/svg" width="28" height="40" viewBox="0 0 28 40">
    <path d="M14 0C6.27 0 0 6.27 0 14c0 10.5 14 26 14 26s14-15.5 14-26C28 6.27 21.73 0 14 0z" fill="${color}" stroke="#fff" stroke-width="1.5"/>
    <circle cx="14" cy="13" r="5.5" fill="#fff" opacity="0.9"/>
  </svg>`;
  const pinIcon = L.divIcon({
    html: pinSvg, className: 'casemap-pin',
    iconSize: [28, 40], iconAnchor: [14, 40], popupAnchor: [0, -36]
  });
  const marker = L.marker([c.lat, c.lng], { icon: pinIcon, draggable: true }).addTo(caseMapInstance);
  marker.bindTooltip(label, {
    permanent: true, direction: 'right', offset: [12, -20], className: 'casemap-label'
  });
  marker.on('dragend', function() {
    const pos = marker.getLatLng();
    c.lat = pos.lat; c.lng = pos.lng;
    const gc = getGeoCache();
    const ck = getCaseGeoAddress(c);
    if (ck) gc[ck] = [pos.lat, pos.lng];
    setGeoCache(gc);
    saveDB(db);
    showToast(`${c.name} 位置已更新`);
  });
  marker.bindPopup(`
    <div style="min-width:220px;line-height:1.7;font-size:0.9rem">
      <div style="font-size:1.1rem;font-weight:700;margin-bottom:4px">${esc(c.name)}
        <span style="background:${color};color:#fff;padding:1px 8px;border-radius:4px;font-size:0.75rem;margin-left:6px">${c.cmsLevel || '-'}</span>
      </div>
      <div>📋 ${esc(c.caseNo || c.id)}</div>
      <div>🏠 ${esc(c.address || (c.district ? '桃園市' + c.district + (c.village||'') : '-'))}</div>
      <div>👨‍⚕️ ${doctor ? esc(doctor.name) : (c.doctorName || '-')}</div>
      <div>📅 收案 ${c.enrollDate || '-'}</div>
      ${c.phone ? `<div>📞 ${esc(c.phone)}</div>` : ''}
      <div style="margin-top:6px;color:#999;font-size:0.75rem">💡 可拖拉圖釘修正位置</div>
    </div>
  `);
  marker.on('mouseover', function() { this.openPopup(); });
  marker.on('mouseout', function() { this.closePopup(); });
  caseMapMarkers.push(marker);
}

// 清理地址供地理編碼用
function cleanAddressForGeo(addr) {
  if (!addr) return '';
  let s = addr;
  // 全形數字轉半形
  s = s.replace(/[０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0));
  // 全形英文轉半形
  s = s.replace(/[Ａ-Ｚａ-ｚ]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0));
  // 移除樓層資訊 (三樓、3樓、3F、之1 等)
  s = s.replace(/[一二三四五六七八九十\d]+樓.*$/, '');
  s = s.replace(/\d+[fF].*$/, '');
  // 移除「里」名稱（如「莊敬里」），因為會干擾定位
  s = s.replace(/([區鄉鎮市])[^\s路街道巷弄號]+里/, '$1');
  // 移除「之N」
  s = s.replace(/之\d+/, '');
  // 移除「號」後面的內容（如「號三樓」「號之1」）
  s = s.replace(/(號).*$/, '$1');
  return s.trim();
}

// 地理編碼單一地址（透過 server proxy）
async function geocodeAddress(addr) {
  try {
    const resp = await apiCall('GET', `/api/geocode?q=${encodeURIComponent(addr)}`);
    if (Array.isArray(resp) && resp.length > 0) {
      return { lat: parseFloat(resp[0].lat), lng: parseFloat(resp[0].lon) };
    }
  } catch (e) { /* ignore */ }
  return null;
}

// 取得個案的最佳地址字串（原始，用於快取 key）
function getCaseGeoAddress(c) {
  if (c.address) return c.address;
  let addr = '桃園市';
  if (c.district) addr += c.district;
  if (c.village) addr += c.village;
  return addr.length > 3 ? addr : '';
}

// 清除定位快取，重新定位
function resetGeocode() {
  if (!confirm('確定要清除所有定位快取並重新定位？')) return;
  localStorage.removeItem('geocache');
  getActiveCases(db).forEach(c => { delete c.lat; delete c.lng; });
  saveDB(db);
  renderCaseMap();
  showToast('已清除定位快取，請點「📍 地址定位」重新定位');
}

// 批次地理編碼
async function geocodeAllCases() {
  const btn = document.getElementById('btn-geocode');
  const bar = document.getElementById('casemap-geocode-bar');
  const statusEl = document.getElementById('casemap-geocode-status');
  if (!bar || !statusEl) return;
  bar.style.display = 'flex';
  btn.disabled = true;
  btn.textContent = '⏳ 定位中...';

  const active = getActiveCases(db);
  const geoCache = getGeoCache();
  const toGeocode = active.filter(c => {
    if (c.lat && c.lng) return false;
    const addr = getCaseGeoAddress(c);
    if (!addr) return false;
    if (geoCache[addr]) return false;
    return true;
  });

  if (toGeocode.length === 0) {
    statusEl.innerHTML = '<span style="color:#4CAF50">✓ 所有個案都已定位</span>';
    btn.disabled = false; btn.textContent = '📍 地址定位';
    renderCaseMap();
    return;
  }

  let done = 0, success = 0, fail = 0;
  statusEl.textContent = `正在定位 0 / ${toGeocode.length} ...`;

  for (const c of toGeocode) {
    const cacheKey = getCaseGeoAddress(c);

    // 伺服器端自動嘗試多種策略
    let result = await geocodeAddress(cacheKey);

    if (result) {
      c.lat = result.lat; c.lng = result.lng;
      geoCache[cacheKey] = [result.lat, result.lng];
      success++;
      // 即時在地圖上加入圖釘
      if (caseMapInstance) addCasePin(c);
    } else {
      fail++;
    }

    done++;
    statusEl.textContent = `正在定位 ${done} / ${toGeocode.length} (成功 ${success}, 失敗 ${fail})...`;

    // 每 20 筆存檔一次
    if (done % 20 === 0) { setGeoCache(geoCache); saveDB(db); }

    // Nominatim rate limit: 1 req/sec
    await new Promise(r => setTimeout(r, 1100));
  }

  setGeoCache(geoCache);
  saveDB(db);
  statusEl.innerHTML = `<span style="color:#4CAF50">✓ 定位完成：成功 ${success}，失敗 ${fail}</span>`;
  btn.disabled = false; btn.textContent = '📍 地址定位';
  renderCaseMap();
}

function computeKPI(doctorFilter) {
  const active = getActiveCases(db).filter(c => {
    if (!doctorFilter) return true;
    const docMember = db.members.find(m => m.name === doctorFilter);
    return (docMember && c.doctorId === docMember.id) || c.doctorName === doctorFilter;
  });
  const now = new Date();
  const yy = String(now.getFullYear());

  if (active.length === 0) return [
    { label:'高血壓測量率', value:0, target:95 },
    { label:'高血糖監測率', value:0, target:70 },
    { label:'高血脂監測率', value:0, target:70 },
    { label:'ACP訓練完成率', value:0, target:100 },
    { label:'ACP/AD完成率', value:0, target:30 },
  ];

  // 血壓: 家訪時量血壓
  const bpCases = active.filter(c => c.hasHypertension || true); // 所有個案家訪都應量血壓
  const bpMeasured = bpCases.filter(c => {
    const homeVisits = db.services.filter(s => s.caseId === c.id && s.type === 'home' && s.date.startsWith(yy));
    return homeVisits.some(s => s.bpMeasured);
  });
  const bpRate = bpCases.length ? (bpMeasured.length / bpCases.length * 100) : 0;

  // 血糖: 糖尿病穩定者一年至少二次 HbA1c
  const dmCases = active.filter(c => c.hasDiabetes);
  const dmMonitored = dmCases.filter(c => {
    const count = db.services.filter(s => s.caseId === c.id && s.hba1cMonitored && s.date.startsWith(yy)).length;
    return count >= 2;
  });
  const dmRate = dmCases.length ? (dmMonitored.length / dmCases.length * 100) : 100;

  // 血脂
  const lipidCases = active.filter(c => c.hasHyperlipidemia);
  const lipidMonitored = lipidCases.filter(c => {
    const count = db.services.filter(s => s.caseId === c.id && s.lipidMonitored && s.date.startsWith(yy)).length;
    return count >= 2;
  });
  const lipidRate = lipidCases.length ? (lipidMonitored.length / lipidCases.length * 100) : 100;

  // ACP 訓練
  const relevantMembers = db.members.filter(m => m.status === 'active' && (!doctorFilter || m.name === doctorFilter));
  const acpTrained = relevantMembers.filter(m => m.acpTrained);
  const acpTrainRate = relevantMembers.length ? (acpTrained.length / relevantMembers.length * 100) : 0;

  // ACP/AD 完成率
  const sixMonthAgo = new Date(now);
  sixMonthAgo.setMonth(sixMonthAgo.getMonth() - 6);
  const eligibleCases = active.filter(c => new Date(c.enrollDate) <= sixMonthAgo);
  const acpCompleted = eligibleCases.filter(c => c.acpExplained && c.adExplained);
  const acpRate = eligibleCases.length ? (acpCompleted.length / eligibleCases.length * 100) : 0;

  return [
    { label: '高血壓測量率', value: bpRate, target: 95, detail: `${bpMeasured.length}/${bpCases.length}` },
    { label: '高血糖監測率', value: dmRate, target: 70, detail: `${dmMonitored.length}/${dmCases.length}` },
    { label: '高血脂監測率', value: lipidRate, target: 70, detail: `${lipidMonitored.length}/${lipidCases.length}` },
    { label: 'ACP訓練完成率', value: acpTrainRate, target: 100, detail: `${acpTrained.length}/${relevantMembers.length}` },
    { label: 'ACP/AD完成率', value: acpRate, target: 30, detail: `${acpCompleted.length}/${eligibleCases.length}` },
  ];
}

function renderKPI() {
  const doctorFilter = document.getElementById('kpi-filter-doctor')?.value || '';
  const kpiData = computeKPI(doctorFilter);

  document.getElementById('kpi-detail-grid').innerHTML = kpiData.map(k => {
    const pct = Math.min(100, k.value);
    const color = k.value >= k.target ? 'var(--success)' : (k.value >= k.target * 0.7 ? 'var(--warning)' : 'var(--danger)');
    const statusText = k.value >= k.target ? '✅ 已達標' : '⚠️ 未達標';
    return `<div class="kpi-card">
      <div class="kpi-card-title">${k.label}</div>
      <div class="kpi-card-value" style="color:${color}">${k.value.toFixed(1)}%</div>
      <div class="kpi-card-target">目標: ${k.target}% ${k.detail ? '(' + k.detail + ')' : ''} — ${statusText}</div>
      <div class="kpi-card-bar"><div class="kpi-card-bar-fill" style="width:${pct}%;background:${color}"></div></div>
    </div>`;
  }).join('');

  // 各醫師明細
  const docs = getDoctors(db).filter(d => d.status === 'active');
  document.getElementById('kpi-doctor-tbody').innerHTML = docs.map(d => {
    const dk = computeKPI(d.id);
    const allMet = dk.every(k => k.value >= k.target);
    return `<tr>
      <td><strong>${esc(d.name)}</strong></td>
      <td>${getCaseCountByMember(db, d.id)}</td>
      ${dk.map(k => {
        const color = k.value >= k.target ? 'success' : (k.value >= k.target*0.7 ? 'warning' : 'danger');
        return `<td><span class="badge badge-${color}">${k.value.toFixed(0)}%</span></td>`;
      }).join('')}
      <td>${allMet ? '<span class="badge badge-success">全達標</span>' : '<span class="badge badge-danger">未達標</span>'}</td>
    </tr>`;
  }).join('');
}

// ==========================================
//  ACP/AD 追蹤
// ==========================================
function renderACP() {
  const search = (document.getElementById('acp-search')?.value || '').toLowerCase();
  const filter = document.getElementById('acp-filter')?.value || '';
  const now = new Date();
  const sixMonthAgo = new Date(now); sixMonthAgo.setMonth(sixMonthAgo.getMonth() - 6);

  const active = getActiveCases(db);

  // Summary
  const total = active.length;
  const explained = active.filter(c => c.acpExplained).length;
  const signed = active.filter(c => c.acpSigned).length;
  const registered = active.filter(c => c.nhiRegistered).length;
  const eligible = active.filter(c => new Date(c.enrollDate) <= sixMonthAgo).length;

  document.getElementById('acp-summary').innerHTML = `
    <div class="summary-item"><div class="summary-value">${total}</div><div class="summary-label">收案中個案</div></div>
    <div class="summary-item"><div class="summary-value" style="color:var(--info)">${explained}</div><div class="summary-label">已ACP說明</div></div>
    <div class="summary-item"><div class="summary-value" style="color:var(--success)">${signed}</div><div class="summary-label">已簽署</div></div>
    <div class="summary-item"><div class="summary-value" style="color:var(--primary)">${registered}</div><div class="summary-label">已健保卡註記</div></div>
    <div class="summary-item"><div class="summary-value">${eligible}</div><div class="summary-label">收案滿6月(需完成)</div></div>
    <div class="summary-item"><div class="summary-value" style="color:var(--warning)">${eligible>0?((explained/eligible*100).toFixed(0)+'%'):'N/A'}</div><div class="summary-label">完成率(目標30%)</div></div>
  `;

  let filtered = active.filter(c => {
    if (search && !c.name.toLowerCase().includes(search)) return false;
    if (filter === 'not_started' && c.acpExplained) return false;
    if (filter === 'explained' && !c.acpExplained) return false;
    if (filter === 'signed' && !c.acpSigned) return false;
    if (filter === 'registered' && !c.nhiRegistered) return false;
    return true;
  });

  document.getElementById('acp-tbody').innerHTML = filtered.map(c => {
    const d = findMember(db, c.doctorId);
    const enrollDate = new Date(c.enrollDate);
    const over6m = enrollDate <= sixMonthAgo;
    const statusBadge = c.nhiRegistered ? '<span class="badge badge-success">已註記</span>'
      : c.acpSigned ? '<span class="badge badge-primary">已簽署</span>'
      : c.acpExplained ? '<span class="badge badge-info">已說明</span>'
      : '<span class="badge badge-gray">未開始</span>';

    return `<tr class="${over6m && !c.acpExplained ? 'row-warning' : ''}">
      <td><strong>${esc(c.name)}</strong></td>
      <td>${d ? esc(d.name) : '-'}</td>
      <td>${c.enrollDate}</td>
      <td>${over6m ? '<span class="badge badge-warning">已滿6月</span>' : '<span class="badge badge-gray">未滿</span>'}</td>
      <td>${c.acpExplained ? '✅ ' + (c.acpExplainedDate || '') : '❌'}</td>
      <td>${c.adExplained ? '✅ ' + (c.adExplainedDate || '') : '❌'}</td>
      <td>${c.acpSigned ? '✅ ' + (c.acpSignedDate || '') : '❌'}</td>
      <td>${c.nhiRegistered ? '✅' : '❌'}</td>
      <td>${c.familyAcpExplained ? '✅' : '❌'}</td>
      <td>${statusBadge}</td>
      <td><button class="btn btn-xs btn-primary" onclick="updateACP('${c.id}')">更新</button></td>
    </tr>`;
  }).join('');
}

function updateACP(caseId) {
  const c = findCase(db, caseId);
  if (!c) return;
  const body = `
    <p style="margin-bottom:1rem;font-weight:600">個案：${esc(c.name)} (${esc(c.caseNo || c.id)})</p>
    <div class="form-group"><label class="form-checkbox"><input type="checkbox" id="f-acp-explained" ${c.acpExplained?'checked':''}> 已完成 ACP 說明</label></div>
    <div class="form-group"><label class="form-checkbox"><input type="checkbox" id="f-ad-explained" ${c.adExplained?'checked':''}> 已完成 AD 說明</label></div>
    <div class="form-group"><label class="form-checkbox"><input type="checkbox" id="f-acp-signed" ${c.acpSigned?'checked':''}> 已簽署 ACP</label></div>
    <div class="form-group"><label class="form-checkbox"><input type="checkbox" id="f-nhi-reg" ${c.nhiRegistered?'checked':''}> 已健保卡註記</label></div>
    <div class="form-group"><label class="form-checkbox"><input type="checkbox" id="f-family-acp" ${c.familyAcpExplained?'checked':''}> 已向家屬說明</label></div>
    <div class="form-hint">※ 每完成1名個案之預立醫療決定簽署(健保卡註記)，補助1,500元</div>
  `;
  const footer = `
    <button class="btn btn-outline" onclick="closeModal()">取消</button>
    <button class="btn btn-primary" onclick="saveACP('${caseId}')">儲存</button>
  `;
  openModal('更新 ACP/AD 狀態', body, footer);
}

function saveACP(caseId) {
  const c = findCase(db, caseId);
  const wasRegistered = c.nhiRegistered;
  c.acpExplained = document.getElementById('f-acp-explained').checked;
  c.acpExplainedDate = c.acpExplained ? (c.acpExplainedDate || fmt(new Date())) : null;
  c.adExplained = document.getElementById('f-ad-explained').checked;
  c.adExplainedDate = c.adExplained ? (c.adExplainedDate || fmt(new Date())) : null;
  c.acpSigned = document.getElementById('f-acp-signed').checked;
  c.acpSignedDate = c.acpSigned ? (c.acpSignedDate || fmt(new Date())) : null;
  c.nhiRegistered = document.getElementById('f-nhi-reg').checked;
  c.familyAcpExplained = document.getElementById('f-family-acp').checked;

  // 完成健保卡註記獎勵
  if (!wasRegistered && c.nhiRegistered) {
    db.billings.push({
      id: 'B' + String(db.billings.length + 1).padStart(5, '0'),
      caseId, code: 'ACP獎勵', serviceDate: fmt(new Date()),
      billingMonth: fmt(new Date()).substring(0, 7),
      memberId: c.doctorId, amount: 1500, status: 'pending',
    });
    showToast('已產生 ACP 獎勵 $1,500');
  }

  saveDB(db);
  closeModal();
  showToast('ACP/AD 狀態已更新');
  renderACP();
}

// ==========================================
//  警示與待辦
// ==========================================
function generateAlerts() {
  const alerts = [];
  const now = new Date();
  const active = getActiveCases(db);
  const ym = fmt(now).substring(0, 7);

  active.forEach(c => {
    // 本月未服務
    const svcThisMonth = getServiceThisMonth(db, c.id);
    if (svcThisMonth.length === 0) {
      const lastSvc = db.services.filter(s => s.caseId === c.id).sort((a,b) => b.date.localeCompare(a.date))[0];
      alerts.push({
        level: 'danger', title: `${esc(c.name)} 本月尚未服務`,
        desc: `上次服務: ${lastSvc ? lastSvc.date : '無紀錄'}`,
        page: 'services', caseId: c.id,
      });
    }

    // 意見書狀態（基於醫師家訪日期，6個月有效期）
    const op = getLatestOpinion(db, c.id);
    const lastVisitDate = c.doctorVisitDate || (op ? op.issueDate : '');
    if (!lastVisitDate) {
      // 從未開立意見書
      alerts.push({
        level: 'danger', title: `${esc(c.name)} 意見書未開立`,
        desc: `收案日 ${c.enrollDate}，尚無醫師家訪紀錄`,
        page: 'opinion', caseId: c.id,
      });
    } else {
      const daysSince = daysBetween(lastVisitDate, fmt(now));
      const expiryDate = new Date(lastVisitDate);
      expiryDate.setMonth(expiryDate.getMonth() + 6);
      const expiryStr = expiryDate.toISOString().slice(0, 10);
      if (daysSince > 180) {
        // 超過6個月 → 逾期
        alerts.push({
          level: 'danger', title: `${esc(c.name)} 意見書已逾期`,
          desc: `上次醫師家訪 ${lastVisitDate}，已逾期 ${daysSince - 180} 天，到期日 ${expiryStr}`,
          page: 'opinion', caseId: c.id,
        });
      } else if (daysSince >= 150) {
        // 5~6個月 → 待更新（到期前1個月內）
        const daysLeft = 180 - daysSince;
        alerts.push({
          level: 'warning', title: `${esc(c.name)} 意見書即將到期`,
          desc: `上次醫師家訪 ${lastVisitDate}，剩餘 ${daysLeft} 天到期，請安排醫師家訪更新`,
          page: 'opinion', caseId: c.id,
        });
      }
    }

    // 4個月未家訪
    const lastHome = db.services.filter(s => s.caseId === c.id && s.type === 'home')
      .sort((a,b) => b.date.localeCompare(a.date))[0];
    if (lastHome) {
      const daysSince = daysBetween(lastHome.date, fmt(now));
      if (daysSince > 120) {
        alerts.push({
          level: 'warning', title: `${esc(c.name)} 超過4個月未家訪`,
          desc: `上次家訪 ${lastHome.date}，已 ${daysSince} 天`,
          page: 'services', caseId: c.id,
        });
      }
    }

    // 本月未電訪
    const enrollDaysTotal = daysBetween(c.enrollDate, fmt(now));
    if (enrollDaysTotal > 30) {
      const phoneThisMonth = getPhoneVisitsThisMonth(db, c.id);
      if (phoneThisMonth.length === 0) {
        const lastPhone = db.services.filter(s => s.caseId === c.id && s.type === 'phone')
          .sort((a,b) => b.date.localeCompare(a.date))[0];
        alerts.push({
          level: 'warning', title: `${esc(c.name)} 本月尚未電訪`,
          desc: `上次電訪: ${lastPhone ? lastPhone.date : '無紀錄'}`,
          page: 'services', caseId: c.id,
        });
      }
    }

    // 個管師家訪逾期 (>90天)
    const lastNurseHome = getLastNurseHomeVisit(db, c.id);
    const nurseHomeDays = lastNurseHome ? daysBetween(lastNurseHome.date, fmt(now)) : (enrollDaysTotal > 90 ? 999 : 0);
    if (nurseHomeDays > 90) {
      const nurseName = c.nurseId ? (findMember(db, c.nurseId)?.name || '') : '';
      alerts.push({
        level: 'warning', title: `${esc(c.name)} 個管師超過3個月未家訪`,
        desc: `${nurseName ? '個管師: ' + esc(nurseName) + '，' : ''}上次家訪: ${lastNurseHome ? lastNurseHome.date : '無紀錄'}`,
        page: 'services', caseId: c.id,
      });
    }

    // ACP 收案滿6月未說明
    const sixMonthAgo = new Date(now); sixMonthAgo.setMonth(sixMonthAgo.getMonth() - 6);
    if (new Date(c.enrollDate) <= sixMonthAgo && !c.acpExplained) {
      alerts.push({
        level: 'info', title: `${esc(c.name)} 收案滿6月尚未完成ACP說明`,
        desc: `收案日 ${c.enrollDate}`,
        page: 'acp', caseId: c.id,
      });
    }
  });

  // 成員 ACP 訓練未完成
  db.members.filter(m => m.status === 'active').forEach(m => {
    if (!m.acpTrained) {
      const monthsSinceJoin = monthsBetween(m.joinDate, fmt(now));
      if (monthsSinceJoin >= 6) {
        alerts.push({
          level: 'danger', title: `${esc(m.name)} ACP訓練逾期`,
          desc: `加入日 ${m.joinDate}，已超過6個月仍未完成`,
          page: 'members',
        });
      } else {
        alerts.push({
          level: 'warning', title: `${esc(m.name)} ACP訓練待完成`,
          desc: `需於加入後6個月內完成，剩餘 ${6 - monthsSinceJoin} 個月`,
          page: 'members',
        });
      }
    }
  });

  // 費用待申報 (次月10日前)
  const pendingBillings = db.billings.filter(b => b.status === 'pending');
  if (pendingBillings.length > 0) {
    alerts.push({
      level: 'warning', title: `${pendingBillings.length} 筆費用待申報`,
      desc: `應於次月10日前完成申報`,
      page: 'billing',
    });
  }

  // 排序: danger > warning > info
  const levelOrder = { danger: 0, warning: 1, info: 2 };
  alerts.sort((a, b) => levelOrder[a.level] - levelOrder[b.level]);
  return alerts;
}

function renderAlerts() {
  const alerts = generateAlerts();
  const danger = alerts.filter(a => a.level === 'danger');
  const warning = alerts.filter(a => a.level === 'warning');
  const info = alerts.filter(a => a.level === 'info');

  document.getElementById('alert-categories').innerHTML = `
    <div class="alert-cat-card cat-danger" onclick="filterAlerts('danger')">
      <div class="alert-cat-count" style="color:var(--danger)">${danger.length}</div>
      <div class="alert-cat-label">緊急 (逾期)</div>
    </div>
    <div class="alert-cat-card cat-warning" onclick="filterAlerts('warning')">
      <div class="alert-cat-count" style="color:var(--warning)">${warning.length}</div>
      <div class="alert-cat-label">警告 (即將到期)</div>
    </div>
    <div class="alert-cat-card cat-info" onclick="filterAlerts('info')">
      <div class="alert-cat-count" style="color:var(--info)">${info.length}</div>
      <div class="alert-cat-label">提醒</div>
    </div>
  `;

  renderAlertList(alerts);
}

function renderAlertList(alerts) {
  document.getElementById('alerts-full-list').innerHTML = alerts.length === 0
    ? '<div style="text-align:center;padding:3rem;color:var(--gray-400)">🎉 目前無待辦事項</div>'
    : alerts.map(a => `<div class="alert-item level-${a.level}">
      <div class="alert-icon">${a.level === 'danger' ? '🔴' : (a.level === 'warning' ? '🟡' : '🔵')}</div>
      <div class="alert-body">
        <div class="alert-title">${esc(a.title)}</div>
        <div class="alert-desc">${esc(a.desc)}</div>
      </div>
      <div class="alert-action">
        <button class="btn btn-xs btn-outline" onclick="switchPage('${a.page || 'dashboard'}')">前往處理</button>
      </div>
    </div>`).join('');
}

function filterAlerts(level) {
  const alerts = generateAlerts().filter(a => a.level === level);
  renderAlertList(alerts);
}

function updateAlertBadge() {
  const alerts = generateAlerts();
  const count = alerts.filter(a => a.level === 'danger').length;
  const badge = document.getElementById('alert-badge');
  badge.textContent = count;
  badge.style.display = count > 0 ? 'block' : 'none';
}

// ==========================================
//  報表匯出
// ==========================================
function generateReport(type) {
  const now = new Date();
  const ym = fmt(now).substring(0, 7);
  let csv = '\uFEFF';
  let filename = '';

  switch (type) {
    case 'monthly': {
      filename = `月報表_${ym}.csv`;
      const active = getActiveCases(db);
      const svcMonth = db.services.filter(s => s.date.startsWith(ym));
      const billingMonth = db.billings.filter(b => b.billingMonth === ym);
      csv += '項目,數值\n';
      csv += `報表月份,${ym}\n`;
      csv += `收案中個案數,${active.length}\n`;
      csv += `本月服務次數,${svcMonth.length}\n`;
      csv += `家訪次數,${svcMonth.filter(s=>s.type==='home').length}\n`;
      csv += `電訪次數,${svcMonth.filter(s=>s.type==='phone').length}\n`;
      csv += `視訊次數,${svcMonth.filter(s=>s.type==='video').length}\n`;
      csv += `本月申報總額,$${billingMonth.reduce((s,b)=>s+b.amount,0).toLocaleString()}\n`;
      const phoneDoneCount = active.filter(c => getPhoneVisitsThisMonth(db, c.id).length > 0).length;
      csv += `個管師電訪完成案數,${phoneDoneCount}/${active.length}\n`;
      const homeDoneCount = active.filter(c => { const lh = getLastNurseHomeVisit(db, c.id); return lh && daysBetween(lh.date, fmt(now)) <= 90; }).length;
      csv += `個管師家訪3月內完成案數,${homeDoneCount}/${active.length}\n`;
      const kpi = computeKPI();
      kpi.forEach(k => csv += `${k.label},${k.value.toFixed(1)}%\n`);
      break;
    }
    case 'case_list': {
      filename = `個案清冊_${ym}.csv`;
      csv += '個案編號,姓名,性別,年齡,CMS等級,負責醫師,個管師,收案日期,狀態,本月電訪次數,上次家訪日,意見書狀態,意見書到期日,疾病診斷,ACP狀態\n';
      db.cases.forEach(c => {
        const d = findMember(db, c.doctorId);
        const n = findMember(db, c.nurseId);
        const phoneCount = getPhoneVisitsThisMonth(db, c.id).length;
        const lastH = db.services.filter(s => s.caseId === c.id && s.type === 'home').sort((a,b) => b.date.localeCompare(a.date))[0];
        const opInfo = getOpinionExpiryInfo(db, c.id);
        const opStatus = opInfo.opinion ? (opInfo.opinion.status === 'valid' ? '有效' : opInfo.opinion.status === 'expiring' ? '即將到期' : '已過期') : '未開立';
        csv += `${escCSV(c.id)},${escCSV(c.name)},${c.gender==='M'?'男':'女'},${c.age},${c.cmsLevel},${d?escCSV(d.name):''},${n?escCSV(n.name):''},${escCSV(c.enrollDate)},${c.status==='active'?'收案中':'已結案'},${phoneCount},${lastH?lastH.date:''},${opStatus},${opInfo.opinion?opInfo.opinion.expiryDate:''},${escCSV(c.diagnoses.map(d=>d.name).join(';'))},${c.acpExplained?'已說明':'未說明'}\n`;
      });
      break;
    }
    case 'billing_summary': {
      filename = `費用彙總_${ym}.csv`;
      csv += '月份,AA12意見書,YA01電訪管理,YA02家訪訪視,YA03原民電訪,YA04原民家訪,合計\n';
      const months = [...new Set(db.billings.map(b => b.billingMonth))].sort();
      months.forEach(m => {
        const items = db.billings.filter(b => b.billingMonth === m);
        const aa12 = items.filter(b=>b.code==='AA12').reduce((s,b)=>s+b.amount,0);
        const ya01 = items.filter(b=>b.code==='YA01').reduce((s,b)=>s+b.amount,0);
        const ya02 = items.filter(b=>b.code==='YA02').reduce((s,b)=>s+b.amount,0);
        const ya03 = items.filter(b=>b.code==='YA03').reduce((s,b)=>s+b.amount,0);
        const ya04 = items.filter(b=>b.code==='YA04').reduce((s,b)=>s+b.amount,0);
        csv += `${m},${aa12},${ya01},${ya02},${ya03},${ya04},${aa12+ya01+ya02+ya03+ya04}\n`;
      });
      break;
    }
    case 'member_performance': {
      filename = `成員績效_${ym}.csv`;
      csv += '姓名,職稱,收案數,血壓測量率,血糖監測率,血脂監測率,ACP訓練,ACP/AD完成率\n';
      getDoctors(db).filter(d => d.status === 'active').forEach(d => {
        const kpi = computeKPI(d.id);
        csv += `${escCSV(d.name)},醫師,${getCaseCountByMember(db,d.id)},${kpi[0].value.toFixed(0)}%,${kpi[1].value.toFixed(0)}%,${kpi[2].value.toFixed(0)}%,${d.acpTrained?'是':'否'},${kpi[4].value.toFixed(0)}%\n`;
      });
      break;
    }
    case 'quarterly': {
      filename = `季報表_${ym}.csv`;
      csv += '季度報表\n';
      csv += '項目,數值\n';
      const kpi = computeKPI();
      kpi.forEach(k => csv += `${k.label},${k.value.toFixed(1)}% (目標${k.target}%)\n`);
      csv += `\n跨專業討論會議\n`;
      csv += `應召開次數,1次/季\n`;
      break;
    }
    case 'annual': {
      filename = `年度評核報告_${now.getFullYear()}.csv`;
      csv += '年度評核指標報告\n\n';
      csv += '策略目標,績效指標,衡量標準,目標值,實際值,達標\n';
      const kpi = computeKPI();
      const rows = [
        ['定期監測個案健康及慢性病情形', '高血壓測量率', '家訪時量血壓之個案數', '95%', kpi[0]],
        ['定期監測個案健康及慢性病情形', '高血糖監測率', '糖尿病穩定者一年至少二次HbA1c', '70%', kpi[1]],
        ['定期監測個案健康及慢性病情形', '高血脂監測率', '一年至少二次完成三酸甘油脂等檢測', '70%', kpi[2]],
        ['推動尊嚴善終', 'ACP訓練完成率', '加入6個月內完成ACP訓練', '100%', kpi[3]],
        ['推動尊嚴善終', 'ACP/AD完成率', '收案滿6月推動ACP/AD', '30%', kpi[4]],
      ];
      rows.forEach(r => {
        csv += `${r[0]},${r[1]},${r[2]},${r[3]},${r[4].value.toFixed(1)}%,${r[4].value >= r[4].target ? '是' : '否'}\n`;
      });
      break;
    }
  }

  if (csv && filename) downloadCSV(csv, filename);
}

function downloadCSV(csv, filename) {
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
  showToast(`已下載 ${filename}`);
}

// ==========================================
//  匯出/匯入
// ==========================================
function exportData() {
  const json = JSON.stringify(db, null, 2);
  const blob = new Blob([json], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `居家失能系統備份_${fmt(new Date())}.json`;
  a.click();
  URL.revokeObjectURL(url);
  showToast('資料已匯出');
}

// ==========================================
//  統一拖拉上傳
// ==========================================
function openUploadZone() {
  document.getElementById('upload-overlay').style.display = 'flex';
}
function closeUploadZone() {
  document.getElementById('upload-overlay').style.display = 'none';
}

// 拖拉事件
document.addEventListener('DOMContentLoaded', () => {
  const overlay = document.getElementById('upload-overlay');
  const zone = document.getElementById('upload-zone');
  if (!zone) return;

  ['dragenter','dragover'].forEach(evt => {
    zone.addEventListener(evt, e => { e.preventDefault(); zone.classList.add('dragover'); });
  });
  ['dragleave','drop'].forEach(evt => {
    zone.addEventListener(evt, e => { e.preventDefault(); zone.classList.remove('dragover'); });
  });
  zone.addEventListener('drop', e => {
    const file = e.dataTransfer.files[0];
    if (file) processUploadFile(file);
  });
  // 點擊背景關閉
  overlay.addEventListener('click', e => { if (e.target === overlay) closeUploadZone(); });
});

function handleUnifiedUpload(e) {
  const file = e.target.files[0];
  if (file) processUploadFile(file);
  e.target.value = '';
}

function processUploadFile(file) {
  const ext = file.name.toLowerCase().split('.').pop();

  if (ext === 'json') {
    // JSON 系統備份匯入
    const reader = new FileReader();
    reader.onload = function(ev) {
      try {
        const data = JSON.parse(ev.target.result);
        if (data.cases && data.members && Array.isArray(data.cases) && Array.isArray(data.members)) {
          db = data;
          saveDB(db);
          closeUploadZone();
          showToast('系統備份匯入成功');
          switchPage('dashboard');
        } else {
          showToast('JSON 格式不正確，需包含 cases 和 members', 'danger');
        }
      } catch(err) {
        showToast('JSON 檔案解析失敗', 'danger');
      }
    };
    reader.readAsText(file);
  } else if (ext === 'xls' || ext === 'xlsx') {
    // LCMS Excel 上傳
    closeUploadZone();
    uploadLCMSFile(file);
  } else {
    showToast('不支援的檔案格式，請上傳 .xls / .xlsx / .json', 'danger');
  }
}

async function uploadLCMSFile(file) {
  showToast('正在解析檔案...');
  try {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    const headerRow = (aoa[0] || []).map(h => String(h).replace(/[\r\n\s]/g, ''));

    const hasLCMS = headerRow.includes('身分證號') && headerRow.includes('案號');
    const hasLocal = headerRow.some(h => h.includes('身份證字號')) || headerRow.includes('照管案號');

    function rocToAD(rocDate) {
      if (!rocDate) return '';
      const str = String(rocDate).trim();
      const m = str.match(/^(\d{2,3})\/(\d{1,2})\/(\d{1,2})$/);
      if (!m) return '';
      const year = parseInt(m[1]) + 1911;
      return `${year}-${m[2].padStart(2,'0')}-${m[3].padStart(2,'0')}`;
    }
    function excelDateToISO(serial) {
      if (!serial || typeof serial !== 'number') return '';
      const d = new Date((serial - 25569) * 86400000);
      return d.toISOString().slice(0, 10);
    }

    let lcmsCases = [];

    if (hasLCMS) {
      // LCMS 長照平台匯出格式
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
      lcmsCases = rows.map(row => {
        const cmsMatch = String(row['CMS'] || '').match(/(\d+)/);
        const ageMatch = String(row['年齡'] || '').match(/(\d+)/);
        const welfareRaw = String(row['福利身分'] || '');
        let category = '';
        if (welfareRaw.includes('第一類')) category = '第一類';
        else if (welfareRaw.includes('第二類')) category = '第二類';
        else if (welfareRaw.includes('第三類')) category = '第三類';
        const statusRaw = String(row['案件狀態'] || '');
        let status = (statusRaw.includes('結案') || statusRaw.includes('終止')) ? 'closed' : 'active';
        return {
          caseNo: String(row['案號'] || '').trim(),
          name: String(row['姓名'] || '').trim(),
          idNumber: String(row['身分證號'] || '').trim().toUpperCase(),
          birthday: rocToAD(row['出生日期']),
          age: ageMatch ? parseInt(ageMatch[1]) : null,
          cmsLevel: cmsMatch ? parseInt(cmsMatch[1]) : null,
          category, district: String(row['居住地(行政區)'] || '').trim(),
          village: String(row['居住地(村里)'] || '').trim(),
          careManager: String(row['照管專員'] || '').trim(),
          unitName: String(row['A單位名稱'] || '').trim(),
          lcmsOpinionCount: parseInt(row['意見書數量(當年度)']) || 0,
          lcmsBillingCount: parseInt(row['申報紀錄數量(當年度)']) || 0,
          hospital: String(row['主責居家醫師院所'] || '').trim(),
          enrollDate: rocToAD(row['派案日期']),
          homeVisitDates: (() => {
            const raw = String(row['家訪日期'] || '').trim();
            if (!raw) return '';
            return raw.split(',').map(d => rocToAD(d.trim())).filter(d => d).join(',');
          })(),
          doctorVisitDate: (() => {
            const raw = String(row['家訪日期'] || '').trim();
            if (!raw) return '';
            const dates = raw.split(',').map(d => rocToAD(d.trim())).filter(d => d);
            if (dates.length === 0) return '';
            dates.sort();
            return dates[dates.length - 1]; // 最後一次家訪日期
          })(),
          status
        };
      }).filter(c => c.idNumber);

    } else if (hasLocal) {
      // 本地個案清冊格式（雙行 header）
      for (let i = 2; i < aoa.length; i++) {
        const r = aoa[i];
        const name = String(r[3] || '').trim();
        const idNumber = String(r[4] || '').trim().toUpperCase();
        if (!name || !idNumber || idNumber.length < 8) continue;
        const cmsMatch = String(r[7] || '').match(/(\d+)/);
        const catRaw = String(r[6] || '').trim();
        let category = '';
        if (catRaw.includes('第一類') || catRaw.includes('一般戶')) category = '第一類';
        else if (catRaw.includes('第二類') || catRaw.includes('中低')) category = '第二類';
        else if (catRaw.includes('第三類') || catRaw.includes('低收')) category = '第三類';
        else category = catRaw;
        let enrollDate = typeof r[12] === 'number' ? excelDateToISO(r[12]) : rocToAD(r[12]);
        const addr = String(r[8] || '').trim();
        const distMatch = addr.match(/([\u4e00-\u9fff]+[區鄉鎮市])/);
        let doctorVisitDate = '';
        if (typeof r[15] === 'number') doctorVisitDate = excelDateToISO(r[15]);
        else if (r[15]) doctorVisitDate = rocToAD(r[15]);

        lcmsCases.push({
          caseNo: String(r[1] || '').trim(), name, idNumber,
          gender: String(r[5] || '').trim(),
          cmsLevel: cmsMatch ? parseInt(cmsMatch[1]) : null,
          category, address: addr, district: distMatch ? distMatch[1] : '',
          village: String(r[9] || '').trim(),
          contactPerson: String(r[10] || '').replace(/[\r\n]/g, ' ').trim(),
          phone: String(r[11] || '').replace(/[\r\n]/g, ' ').trim(),
          doctorName: String(r[13] || '').trim(),
          enrollDate, doctorVisitDate, status: 'active'
        });
      }
      // 讀取結案總表
      if (workbook.SheetNames.includes('結案總表')) {
        const cs = workbook.Sheets['結案總表'];
        const ca = XLSX.utils.sheet_to_json(cs, { header: 1, defval: '' });
        for (let i = 2; i < ca.length; i++) {
          const r = ca[i];
          const name = String(r[3] || '').trim();
          const idNumber = String(r[4] || '').trim().toUpperCase();
          if (!name || !idNumber || idNumber.length < 8) continue;
          const existing = lcmsCases.find(c => c.idNumber === idNumber);
          if (existing) { existing.status = 'closed'; }
          else {
            const cmsMatch = String(r[7] || '').match(/(\d+)/);
            lcmsCases.push({
              caseNo: String(r[1] || '').trim(), name, idNumber,
              gender: String(r[5] || '').trim(),
              cmsLevel: cmsMatch ? parseInt(cmsMatch[1]) : null,
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
      showToast('無法辨識檔案格式，請使用 LCMS 或本地個案清冊', 'danger');
      return;
    }

    console.log('[Upload] format:', hasLCMS ? 'LCMS' : 'LOCAL', 'cases:', lcmsCases.length);
    showToast(`解析完成：${lcmsCases.length} 筆個案`);
    const diff = computeSmartDiff(lcmsCases);
    showSmartDiffModal(diff, lcmsCases);
  } catch (err) {
    console.error('Excel 解析失敗:', err);
    showToast('檔案解析失敗: ' + err.message, 'danger');
  }
}

function computeSmartDiff(lcmsCases) {
  const systemCases = (db && db.cases) ? db.cases : [];

  const systemMap = {};
  systemCases.forEach(c => {
    if (c.idNumber) systemMap[c.idNumber.toUpperCase()] = c;
  });
  const lcmsMap = {};
  lcmsCases.forEach(c => {
    if (c.idNumber) lcmsMap[c.idNumber.toUpperCase()] = c;
  });

  const newCases = [];
  const changed = [];
  const possiblyClosed = [];

  lcmsCases.forEach(lc => {
    const key = lc.idNumber.toUpperCase();
    const sc = systemMap[key];
    if (!sc) {
      newCases.push(lc);
    } else {
      const diffs = [];
      if (lc.cmsLevel != null && sc.cmsLevel != null && String(lc.cmsLevel) !== String(sc.cmsLevel))
        diffs.push({ field: 'CMS等級', oldVal: sc.cmsLevel, newVal: lc.cmsLevel });
      if (lc.category && sc.category && lc.category !== sc.category)
        diffs.push({ field: '福利身分', oldVal: sc.category, newVal: lc.category });
      if (lc.district && sc.district && lc.district !== sc.district)
        diffs.push({ field: '行政區', oldVal: sc.district, newVal: lc.district });
      if (lc.age != null && sc.age != null && String(lc.age) !== String(sc.age))
        diffs.push({ field: '年齡', oldVal: sc.age, newVal: lc.age });
      if (lc.caseNo && sc.caseNo && lc.caseNo !== sc.caseNo)
        diffs.push({ field: '案號', oldVal: sc.caseNo, newVal: lc.caseNo });
      if (lc.enrollDate && sc.enrollDate && lc.enrollDate !== sc.enrollDate)
        diffs.push({ field: '派案日期', oldVal: sc.enrollDate, newVal: lc.enrollDate });
      if (lc.careManager && lc.careManager !== (sc.careManager || ''))
        diffs.push({ field: '照管專員', oldVal: sc.careManager || '(空)', newVal: lc.careManager });
      if (lc.village && lc.village !== (sc.village || ''))
        diffs.push({ field: '村里', oldVal: sc.village || '(空)', newVal: lc.village });
      if (lc.unitName && lc.unitName !== (sc.unitName || ''))
        diffs.push({ field: 'A單位名稱', oldVal: sc.unitName || '(空)', newVal: lc.unitName });
      if (lc.lcmsOpinionCount != null && lc.lcmsOpinionCount !== (sc.lcmsOpinionCount || 0))
        diffs.push({ field: '意見書數量(年度)', oldVal: sc.lcmsOpinionCount || 0, newVal: lc.lcmsOpinionCount });
      if (lc.lcmsBillingCount != null && lc.lcmsBillingCount !== (sc.lcmsBillingCount || 0))
        diffs.push({ field: '申報紀錄數量(年度)', oldVal: sc.lcmsBillingCount || 0, newVal: lc.lcmsBillingCount });
      if (lc.homeVisitDates && lc.homeVisitDates !== (sc.homeVisitDates || ''))
        diffs.push({ field: '家訪日期', oldVal: sc.homeVisitDates || '(空)', newVal: lc.homeVisitDates });
      // 本地清冊格式的欄位
      if (lc.doctorVisitDate && lc.doctorVisitDate !== (sc.doctorVisitDate || ''))
        diffs.push({ field: '醫師家訪日期', oldVal: sc.doctorVisitDate || '(空)', newVal: lc.doctorVisitDate });
      if (lc.doctorName && lc.doctorName !== (sc.doctorName || ''))
        diffs.push({ field: '主責醫師', oldVal: sc.doctorName || '(空)', newVal: lc.doctorName });
      if (lc.address && lc.address !== (sc.address || ''))
        diffs.push({ field: '地址', oldVal: sc.address || '(空)', newVal: lc.address });
      if (lc.phone && lc.phone !== (sc.phone || ''))
        diffs.push({ field: '電話', oldVal: sc.phone || '(空)', newVal: lc.phone });
      if (lc.gender && lc.gender !== (sc.gender || ''))
        diffs.push({ field: '性別', oldVal: sc.gender || '(空)', newVal: lc.gender });
      // 檔案中標記結案的個案
      if (lc.status === 'closed' && sc.status === 'active')
        diffs.push({ field: '狀態', oldVal: '收案中', newVal: '結案' });
      if (diffs.length > 0) {
        changed.push({ lcms: lc, system: sc, diffs });
      }
    }
  });

  // 系統中收案但 LCMS 無資料 → 可能結案
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
    newCases, changed, possiblyClosed
  };
}

function showSmartDiffModal(diff) {
  const modalTitle = document.getElementById('modal-title');
  const modalBody = document.getElementById('modal-body');
  const modalFooter = document.getElementById('modal-footer');

  modalTitle.textContent = '個案資料比對結果';

  let html = `<div style="max-height:70vh;overflow-y:auto;">
    <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(130px,1fr));gap:0.75rem;margin-bottom:1rem;">
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
        <div style="font-size:0.85rem;color:#555;">新增個案</div>
      </div>
      <div style="background:#fff3e0;padding:0.75rem;border-radius:8px;text-align:center;">
        <div style="font-size:1.5rem;font-weight:bold;color:#e65100;">${diff.changed.length}</div>
        <div style="font-size:0.85rem;color:#555;">資料異動</div>
      </div>
      <div style="background:#fce4ec;padding:0.75rem;border-radius:8px;text-align:center;">
        <div style="font-size:1.5rem;font-weight:bold;color:#c62828;">${diff.possiblyClosed.length}</div>
        <div style="font-size:0.85rem;color:#555;">可能結案</div>
      </div>
    </div>`;

  if (diff.newCases.length > 0) {
    html += `<details open style="margin-bottom:1rem;">
      <summary style="font-weight:bold;font-size:1rem;cursor:pointer;color:#2e7d32;margin-bottom:0.5rem;">
        新增個案 (${diff.newCases.length})
      </summary>
      <div style="overflow-x:auto;">
        <table class="data-table" style="font-size:0.85rem;"><thead><tr>
          <th><input type="checkbox" id="diff-new-all" onchange="toggleDiffCheckAll(this,'diff-new-chk')" checked></th>
          <th>案號</th><th>姓名</th><th>身分證號</th><th>CMS</th><th>身分別</th><th>行政區</th><th>派案日期</th>
        </tr></thead><tbody>
          ${diff.newCases.map((c, i) => `<tr>
            <td><input type="checkbox" class="diff-new-chk" data-idx="${i}" checked></td>
            <td>${esc(c.caseNo)}</td><td>${esc(c.name)}</td><td>${maskId(c.idNumber)}</td>
            <td>${c.cmsLevel || '-'}</td><td>${esc(c.category || '-')}</td>
            <td>${esc(c.district || '-')}</td><td>${esc(c.enrollDate || '-')}</td>
          </tr>`).join('')}
        </tbody></table>
      </div>
    </details>`;
  }

  if (diff.changed.length > 0) {
    html += `<details open style="margin-bottom:1rem;">
      <summary style="font-weight:bold;font-size:1rem;cursor:pointer;color:#e65100;margin-bottom:0.5rem;">
        資料異動 (${diff.changed.length})
      </summary>
      <div style="overflow-x:auto;">
        <table class="data-table" style="font-size:0.85rem;"><thead><tr>
          <th><input type="checkbox" id="diff-chg-all" onchange="toggleDiffCheckAll(this,'diff-chg-chk')" checked></th>
          <th>姓名</th><th>身分證號</th><th>異動欄位</th><th>原值</th><th>新值</th>
        </tr></thead><tbody>
          ${diff.changed.map((item, i) => item.diffs.map((d, j) => `<tr>
            ${j === 0 ? `<td rowspan="${item.diffs.length}"><input type="checkbox" class="diff-chg-chk" data-idx="${i}" checked></td>
              <td rowspan="${item.diffs.length}">${esc(item.lcms.name)}</td>
              <td rowspan="${item.diffs.length}">${maskId(item.lcms.idNumber)}</td>` : ''}
            <td><span style="color:#e65100;font-weight:500;">${esc(d.field)}</span></td>
            <td style="color:#999;text-decoration:line-through;">${esc(String(d.oldVal))}</td>
            <td style="color:#2e7d32;font-weight:500;">${esc(String(d.newVal))}</td>
          </tr>`).join('')).join('')}
        </tbody></table>
      </div>
    </details>`;
  }

  if (diff.possiblyClosed.length > 0) {
    html += `<details open style="margin-bottom:1rem;">
      <summary style="font-weight:bold;font-size:1rem;cursor:pointer;color:#c62828;margin-bottom:0.5rem;">
        可能結案 (${diff.possiblyClosed.length}) — 系統中收案但 LCMS 無資料
      </summary>
      <div style="overflow-x:auto;">
        <table class="data-table" style="font-size:0.85rem;"><thead><tr>
          <th><input type="checkbox" id="diff-close-all" onchange="toggleDiffCheckAll(this,'diff-close-chk')" checked></th>
          <th>姓名</th><th>身分證號</th><th>CMS</th><th>負責醫師</th><th>收案日期</th>
        </tr></thead><tbody>
          ${diff.possiblyClosed.map((c, i) => `<tr>
            <td><input type="checkbox" class="diff-close-chk" data-idx="${i}" checked></td>
            <td>${esc(c.name || '-')}</td><td>${maskId(c.idNumber)}</td>
            <td>${c.cmsLevel || '-'}</td><td>${esc(c.doctorName || '-')}</td>
            <td>${esc(c.enrollDate || '-')}</td>
          </tr>`).join('')}
        </tbody></table>
      </div>
    </details>`;
  }

  if (diff.newCases.length === 0 && diff.changed.length === 0 && diff.possiblyClosed.length === 0) {
    html += '<p style="text-align:center;color:#666;padding:2rem;">資料完全一致，無需同步。</p>';
  }

  html += '</div>';
  modalBody.innerHTML = html;
  window._smartDiff = diff;

  const hasChanges = diff.newCases.length > 0 || diff.changed.length > 0 || diff.possiblyClosed.length > 0;
  modalFooter.innerHTML = `
    <button class="btn btn-outline" onclick="closeModal()">關閉</button>
    ${hasChanges ? '<button class="btn btn-primary" onclick="applySmartDiff()">套用勾選項目</button>' : ''}
  `;

  const modal = document.getElementById('modal');
  modal.style.maxWidth = '900px';
  modal.style.width = '90vw';
  document.getElementById('modal-overlay').classList.add('active');
}

function toggleDiffCheckAll(master, className) {
  document.querySelectorAll('.' + className).forEach(cb => { cb.checked = master.checked; });
}

async function applySmartDiff() {
  const diff = window._smartDiff;
  if (!diff) return;

  let addedCount = 0, updatedCount = 0, closedCount = 0;

  // 新增個案
  document.querySelectorAll('.diff-new-chk:checked').forEach(cb => {
    const lc = diff.newCases[parseInt(cb.dataset.idx)];
    if (!lc) return;
    const maxId = db.cases.reduce((max, c) => {
      const num = parseInt((c.id || '').replace('C', ''));
      return num > max ? num : max;
    }, 0);
    db.cases.push({
      id: 'C' + String(maxId + 1 + addedCount).padStart(3, '0'),
      caseNo: lc.caseNo || '', name: lc.name, idNumber: lc.idNumber,
      gender: lc.gender || '', cmsLevel: lc.cmsLevel, category: lc.category || '',
      district: lc.district || '', village: lc.village || '', address: lc.address || '',
      status: lc.status === 'closed' ? 'closed' : 'active',
      doctorName: lc.doctorName || '', enrollDate: lc.enrollDate || '', age: lc.age,
      doctorVisitDate: lc.doctorVisitDate || '',
      careManager: lc.careManager || '', unitName: lc.unitName || '',
      lcmsOpinionCount: lc.lcmsOpinionCount || 0, lcmsBillingCount: lc.lcmsBillingCount || 0,
      homeVisitDates: lc.homeVisitDates || '',
      phone: lc.phone || '', contactPerson: lc.contactPerson || '',
      diagnosis: '', notes: '',
      opinionDate: '', opinionExpiry: '',
      acpStatus: 'not_started', adStatus: 'not_started',
      acpExplainDate: '', adExplainDate: '', acpSignDate: '',
      nhiCardDate: '', familyExplainDate: ''
    });
    addedCount++;
  });

  // 異動更新
  document.querySelectorAll('.diff-chg-chk:checked').forEach(cb => {
    const item = diff.changed[parseInt(cb.dataset.idx)];
    if (!item) return;
    const sc = db.cases.find(c => c.idNumber && c.idNumber.toUpperCase() === item.lcms.idNumber.toUpperCase());
    if (!sc) return;
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
        case '醫師家訪日期': sc.doctorVisitDate = item.lcms.doctorVisitDate; break;
        case '主責醫師': sc.doctorName = item.lcms.doctorName; break;
        case '地址': sc.address = item.lcms.address; break;
        case '電話': sc.phone = item.lcms.phone; break;
        case '性別': sc.gender = item.lcms.gender; break;
        case '狀態': sc.status = 'closed'; break;
      }
    });
    updatedCount++;
  });

  // 結案
  document.querySelectorAll('.diff-close-chk:checked').forEach(cb => {
    const sc = diff.possiblyClosed[parseInt(cb.dataset.idx)];
    if (!sc) return;
    const target = db.cases.find(c => c.id === sc.id);
    if (target) { target.status = 'closed'; closedCount++; }
  });

  // 根據 doctorName 自動配對 doctorId
  const doctors = getDoctors(db).filter(d => d.status === 'active');
  db.cases.forEach(c => {
    if (c.doctorId || !c.doctorName) return;
    const match = doctors.find(d => d.name === c.doctorName);
    if (match) c.doctorId = match.id;
  });

  // 根據醫師家訪日期自動建立意見書
  let opinionCount = 0;
  if (!db.opinions) db.opinions = [];
  db.cases.forEach(c => {
    if (!c.doctorVisitDate || c.status !== 'active') return;
    // 檢查是否已有同日期的意見書
    const exists = db.opinions.some(o => o.caseId === c.id && o.issueDate === c.doctorVisitDate);
    if (exists) return;
    const expDate = new Date(c.doctorVisitDate);
    expDate.setMonth(expDate.getMonth() + 6);
    const expiryStr = expDate.toISOString().slice(0, 10);
    const now = new Date();
    let status = 'active';
    if (now > expDate) status = 'expired';
    else if ((expDate - now) / (1000*60*60*24) <= 30) status = 'expiring';
    const maxOpId = db.opinions.reduce((max, o) => {
      const num = parseInt((o.id || '').replace('OP', ''));
      return num > max ? num : max;
    }, 0);
    db.opinions.push({
      id: 'OP' + String(maxOpId + 1 + opinionCount).padStart(4, '0'),
      caseId: c.id, doctorId: '', doctorName: c.doctorName || '',
      issueDate: c.doctorVisitDate, homeVisitDate: c.doctorVisitDate,
      expiryDate: expiryStr, sequence: 1, yearCount: 1,
      diseaseStatus: '穩定', functionalPrognosis: '穩定', status
    });
    opinionCount++;
  });

  if (addedCount > 0 || updatedCount > 0 || closedCount > 0) {
    saveDB_local(db);
    closeModal();
    const modal = document.getElementById('modal');
    modal.style.maxWidth = ''; modal.style.width = '';

    const msgs = [];
    if (addedCount > 0) msgs.push(`新增 ${addedCount} 個案`);
    if (updatedCount > 0) msgs.push(`更新 ${updatedCount} 個案`);
    if (closedCount > 0) msgs.push(`結案 ${closedCount} 個案`);
    if (opinionCount > 0) msgs.push(`建立 ${opinionCount} 筆意見書`);
    showToast(msgs.join('、') + ' — 正在同步到雲端...');

    if (typeof renderDashboard === 'function') renderDashboard();
    if (typeof renderCases === 'function') renderCases();
    if (typeof updateAlertBadge === 'function') updateAlertBadge();

    try {
      await syncDataToBackend(db);
      showToast(msgs.join('、') + ' — 雲端同步完成 ✓');
    } catch (err) {
      console.error('雲端同步失敗:', err);
      showToast('資料已存本地，但雲端同步失敗: ' + err.message, 'danger');
    }
  } else {
    closeModal();
    const modal = document.getElementById('modal');
    modal.style.maxWidth = ''; modal.style.width = '';
    showToast('未選取任何項目');
  }
}
