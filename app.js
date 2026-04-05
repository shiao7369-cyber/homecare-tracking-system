/* ========================================
   居家失能個案追蹤管理系統 — 主應用邏輯
   ======================================== */

let db;
let currentMemberTab = 'doctors';

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
  kpi: 'KPI 績效監控', acp: 'ACP/AD 追蹤', alerts: '警示與待辦', reports: '報表匯出', users: '使用者管理'
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
    kpi: renderKPI, acp: renderACP, alerts: renderAlerts, users: renderUserManagement
  };
  if (renderers[page]) renderers[page]();
}

// ===== 篩選器初始化 =====
function initFilters() {
  ['case-search','case-filter-status','case-filter-level','case-filter-doctor','case-filter-category'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.addEventListener('input', renderCases);
    if (el) el.addEventListener('change', renderCases);
  });
  ['service-search','service-filter-type','service-filter-month'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.addEventListener('input', renderServices);
    if (el) el.addEventListener('change', renderServices);
  });
  ['opinion-search','opinion-filter-status'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.addEventListener('input', renderOpinions);
    if (el) el.addEventListener('change', renderOpinions);
  });
  ['billing-month','billing-filter-code','billing-filter-status'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.addEventListener('change', renderBilling);
  });
  ['acp-search','acp-filter'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.addEventListener('input', renderACP);
    if (el) el.addEventListener('change', renderACP);
  });
  ['member-search'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.addEventListener('input', renderMembers);
  });
  ['kpi-year','kpi-filter-doctor'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.addEventListener('change', renderKPI);
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
  toast.innerHTML = msg;
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

  // 意見書待更新
  let opDue = 0;
  active.forEach(c => {
    const op = getLatestOpinion(db, c.id);
    if (!op || op.status === 'expired' || op.status === 'expiring') opDue++;
  });
  document.getElementById('stat-opinion-due').textContent = opDue;

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
      <div class="load-bar-name" title="${d.name}">${d.name}</div>
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
      <div class="recent-info"><strong>${c ? c.name : s.caseId}</strong> — ${serviceTypeLabel(s.type)}${n ? ' / ' + n.name : ''}</div>
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
        <div class="alert-title">${a.title}</div>
        <div class="alert-desc">${a.desc}</div>
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
    el.innerHTML = firstOpt + docs.map(d => `<option value="${d.id}">${d.name}</option>`).join('');
  });
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

  let filtered = db.cases.filter(c => {
    if (search && !c.name.toLowerCase().includes(search) && !c.idNumber.toLowerCase().includes(search) && !(c.id && c.id.toLowerCase().includes(search)) && !(c.district && c.district.includes(search)) && !(c.contactPerson && c.contactPerson.includes(search))) return false;
    if (status && c.status !== status) return false;
    if (level && c.cmsLevel !== parseInt(level)) return false;
    if (doctor && c.doctorId !== doctor) return false;
    if (category && c.category !== category) return false;
    return true;
  });

  const tbody = document.getElementById('cases-tbody');
  tbody.innerHTML = filtered.map(c => {
    const doc = findMember(db, c.doctorId);
    const nurse = findMember(db, c.nurseId);
    const op = getLatestOpinion(db, c.id);
    const svcThisMonth = getServiceThisMonth(db, c.id);
    const opBadge = op ? opinionStatusBadge(op.status) : '<span class="badge badge-gray">待開立</span>';
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
      ? `<span class="badge badge-success" title="${trackingText}">${trackingText}</span>`
      : (c.status === 'active' ? '<span class="badge badge-danger">未追蹤</span>' : '-');

    return `<tr class="${c.status === 'active' && !trackingText ? 'row-warning' : ''}">
      <td><small>${c.id}</small></td>
      <td><strong>${c.name}</strong></td>
      <td><small>${c.category || '-'}</small></td>
      <td><span class="badge badge-primary">${c.cmsLevel || '-'}</span></td>
      <td><small>${c.district || '-'}</small></td>
      <td>${doc ? doc.name : (c.doctorName || '-')}</td>
      <td>${opBadge}</td>
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
        <input class="form-input" id="f-case-no" value="${c ? c.id : ''}" style="width:100%" ${c ? 'readonly' : ''}>
      </div>
      <div class="form-group">
        <label class="form-label">姓名 *</label>
        <input class="form-input" id="f-case-name" value="${c ? c.name : ''}" style="width:100%">
      </div>
      <div class="form-group">
        <label class="form-label">身分證字號 *</label>
        <input class="form-input" id="f-case-id-number" value="${c ? c.idNumber : ''}" style="width:100%">
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
        <input class="form-input" id="f-case-phone" value="${c ? c.phone : ''}" style="width:100%">
      </div>
      <div class="form-group">
        <label class="form-label">主要聯絡人</label>
        <input class="form-input" id="f-case-contact" value="${c ? (c.contactPerson||'') : ''}" style="width:100%">
      </div>
      <div class="form-group">
        <label class="form-label">收案日期 *</label>
        <input type="date" class="form-input" id="f-case-enroll" value="${c ? c.enrollDate : fmt(new Date())}" style="width:100%">
      </div>
    </div>
    <div class="form-row">
      <div class="form-group" style="flex:2">
        <label class="form-label">住址</label>
        <input class="form-input" id="f-case-address" value="${c ? c.address : ''}" style="width:100%">
      </div>
      <div class="form-group">
        <label class="form-label">里別</label>
        <input class="form-input" id="f-case-district" value="${c ? (c.district||'') : ''}" style="width:100%">
      </div>
    </div>
    <div class="form-row">
      <div class="form-group">
        <label class="form-label">負責醫師 *</label>
        <select class="form-select" id="f-case-doctor" style="width:100%">
          ${docs.map(d => `<option value="${d.id}" ${c?.doctorId===d.id?'selected':''}>${d.name} (${d.specialty}) [${getCaseCountByMember(db,d.id)}/200]</option>`).join('')}
        </select>
      </div>
      <div class="form-group">
        <label class="form-label">個案管理師 *</label>
        <select class="form-select" id="f-case-nurse" style="width:100%">
          ${nurs.map(n => `<option value="${n.id}" ${c?.nurseId===n.id?'selected':''}>${n.name} [${getCaseCountByMember(db,n.id)}/200]</option>`).join('')}
        </select>
      </div>
    </div>
    <div class="form-group">
      <label class="form-label">備註</label>
      <input class="form-input" id="f-case-notes" value="${c ? (c.notes||'') : ''}" style="width:100%">
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
          <tr><td style="padding:2px 8px;color:var(--gray-500)">照管案號</td><td><strong>${c.id}</strong></td></tr>
          <tr><td style="padding:2px 8px;color:var(--gray-500)">姓名</td><td><strong>${c.name}</strong></td></tr>
          <tr><td style="padding:2px 8px;color:var(--gray-500)">身分證</td><td>${c.idNumber}</td></tr>
          <tr><td style="padding:2px 8px;color:var(--gray-500)">性別</td><td>${c.gender==='M'?'男':'女'}</td></tr>
          <tr><td style="padding:2px 8px;color:var(--gray-500)">身分別</td><td>${c.category || '-'}</td></tr>
          <tr><td style="padding:2px 8px;color:var(--gray-500)">CMS等級</td><td>第${c.cmsLevel}級</td></tr>
          <tr><td style="padding:2px 8px;color:var(--gray-500)">收案日期</td><td>${c.enrollDate || '-'}</td></tr>
          <tr><td style="padding:2px 8px;color:var(--gray-500)">服務天數</td><td>${c.serviceDays || '-'} 天</td></tr>
          <tr><td style="padding:2px 8px;color:var(--gray-500)">地址</td><td>${c.address}</td></tr>
          <tr><td style="padding:2px 8px;color:var(--gray-500)">里別</td><td>${c.district || '-'}</td></tr>
        </table>
      </div>
      <div>
        <h4 style="margin-bottom:.5rem;color:var(--gray-600)">聯絡與照護資訊</h4>
        <table style="font-size:.85rem;width:100%">
          <tr><td style="padding:2px 8px;color:var(--gray-500)">聯絡電話</td><td>${c.phone || '-'}</td></tr>
          <tr><td style="padding:2px 8px;color:var(--gray-500)">主要聯絡人</td><td>${c.contactPerson || '-'}</td></tr>
          <tr><td style="padding:2px 8px;color:var(--gray-500)">負責醫師</td><td>${doc ? doc.name : (c.doctorName || '-')}</td></tr>
          <tr><td style="padding:2px 8px;color:var(--gray-500)">個管師</td><td>${nurse ? nurse.name : '-'}</td></tr>
          <tr><td style="padding:2px 8px;color:var(--gray-500)">醫師家訪日</td><td>${c.doctorVisitDate || '-'}</td></tr>
          <tr><td style="padding:2px 8px;color:var(--gray-500)">個管師家訪日</td><td>${c.nurseVisitDate || '-'}</td></tr>
          <tr><td style="padding:2px 8px;color:var(--gray-500)">意見書</td><td>${op ? opinionStatusBadge(op.status) + ' ' + op.issueDate : '尚未開立'}</td></tr>
          <tr><td style="padding:2px 8px;color:var(--gray-500)">預約家訪</td><td>${c.scheduledVisit || '-'}</td></tr>
          <tr><td style="padding:2px 8px;color:var(--gray-500)">備註</td><td>${c.notes || '-'}</td></tr>
          ${c.status === 'closed' ? `<tr><td style="padding:2px 8px;color:var(--gray-500)">結案資訊</td><td style="color:#dc2626">${c.closeInfo || c.closeReason || '-'}</td></tr>` : ''}
        </table>
      </div>
    </div>
    <h4 style="margin:1rem 0 .5rem;color:var(--gray-600)">115年度月追蹤紀錄 <small style="color:var(--gray-400)">(家=家訪 電=電訪 視=視訊 結=結案)</small></h4>
    <table class="data-table" style="font-size:.82rem">
      <thead><tr>${monthNames.map(m => `<th style="text-align:center">${m}</th>`).join('')}</tr></thead>
      <tbody><tr>${trackingHtml}</tr></tbody>
    </table>
    <h4 style="margin:1rem 0 .5rem;color:var(--gray-600)">疾病診斷</h4>
    <table class="data-table" style="font-size:.82rem">
      <thead><tr><th>ICD-10</th><th>疾病名稱</th><th>發病時間</th></tr></thead>
      <tbody>${(c.diagnoses||[]).map(d => `<tr><td>${d.icd}</td><td>${d.name}</td><td>${d.onset}</td></tr>`).join('') || '<tr><td colspan="3" style="text-align:center;color:var(--gray-400)">尚未記錄</td></tr>'}</tbody>
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
  openModal('個案詳情 — ' + c.name, body, footer);
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
      <td><strong>${m.name}</strong></td>
      <td>${m.role === 'doctor' ? '醫師' : '個管師'}</td>
      <td>${m.specialty}</td>
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
        <input class="form-input" id="f-member-name" value="${m?m.name:''}" style="width:100%">
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
        <input class="form-input" id="f-member-spec" value="${m?m.specialty:''}" style="width:100%">
      </div>
      <div class="form-group">
        <label class="form-label">聯絡電話</label>
        <input class="form-input" id="f-member-phone" value="${m?m.phone:''}" style="width:100%">
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
      <td>${c ? c.name : s.caseId}</td>
      <td><span class="badge badge-${s.type==='home'?'primary':(s.type==='phone'?'success':'info')}">${serviceTypeLabel(s.type)}</span></td>
      <td>${respondentLabel(s.respondent)}</td>
      <td>${n ? n.name : '-'}</td>
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
          ${activeCases.map(c => `<option value="${c.id}">${c.name} (${c.id})</option>`).join('')}
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
          ${nurses.map(n => `<option value="${n.id}">${n.name}</option>`).join('')}
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
      <tr><td style="padding:4px 8px;color:var(--gray-500);width:120px">個案</td><td>${c?c.name:s.caseId}</td></tr>
      <tr><td style="padding:4px 8px;color:var(--gray-500)">日期</td><td>${s.date}</td></tr>
      <tr><td style="padding:4px 8px;color:var(--gray-500)">形式</td><td>${serviceTypeLabel(s.type)}</td></tr>
      <tr><td style="padding:4px 8px;color:var(--gray-500)">個管師</td><td>${n?n.name:'-'}</td></tr>
      <tr><td style="padding:4px 8px;color:var(--gray-500)">受訪者</td><td>${respondentLabel(s.respondent)}</td></tr>
      <tr><td style="padding:4px 8px;color:var(--gray-500)">血壓</td><td>${s.bpMeasured?s.bpSystolic+'/'+s.bpDiastolic+' mmHg':'未測量'}</td></tr>
      <tr><td style="padding:4px 8px;color:var(--gray-500)">糖化血紅素</td><td>${s.hba1cMonitored?(s.hba1cValue||'已測'):'未測'}</td></tr>
      <tr><td style="padding:4px 8px;color:var(--gray-500)">血脂</td><td>${s.lipidMonitored?'已監測':'未測'}</td></tr>
      <tr><td style="padding:4px 8px;color:var(--gray-500)">衛教指導</td><td>${s.educationProvided?'✅':'❌'}</td></tr>
      <tr><td style="padding:4px 8px;color:var(--gray-500)">ACP說明</td><td>${s.acpExplained?'✅':'❌'}</td></tr>
      <tr><td style="padding:4px 8px;color:var(--gray-500)">慢性病評估</td><td>${s.chronicDiseaseEval?'✅':'❌'}</td></tr>
      <tr><td style="padding:4px 8px;color:var(--gray-500)">轉介</td><td>${s.referralLtc?'長照 ':''} ${s.referralMedical?'醫療':''} ${!s.referralLtc&&!s.referralMedical?'無':''}</td></tr>
      <tr><td style="padding:4px 8px;color:var(--gray-500)">申報狀態</td><td>${billingStatusBadge(s.billingStatus)}</td></tr>
      <tr><td style="padding:4px 8px;color:var(--gray-500)">備註</td><td>${s.notes||'-'}</td></tr>
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

  // 更新狀態
  db.opinions.forEach(op => {
    const exp = new Date(op.expiryDate);
    if (exp < now) op.status = 'expired';
    else if (exp - now < 30 * 86400000) op.status = 'expiring';
    else op.status = 'valid';
  });

  // Summary
  const total = db.opinions.length;
  const valid = db.opinions.filter(o => o.status === 'valid').length;
  const expiring = db.opinions.filter(o => o.status === 'expiring').length;
  const expired = db.opinions.filter(o => o.status === 'expired').length;
  const pending = getActiveCases(db).filter(c => !getLatestOpinion(db, c.id)).length;

  document.getElementById('opinion-summary').innerHTML = `
    <div class="summary-item"><div class="summary-value">${total}</div><div class="summary-label">總意見書數</div></div>
    <div class="summary-item"><div class="summary-value" style="color:var(--success)">${valid}</div><div class="summary-label">有效</div></div>
    <div class="summary-item"><div class="summary-value" style="color:var(--warning)">${expiring}</div><div class="summary-label">即將到期(30天內)</div></div>
    <div class="summary-item"><div class="summary-value" style="color:var(--danger)">${expired}</div><div class="summary-label">已過期</div></div>
    <div class="summary-item"><div class="summary-value" style="color:var(--gray-500)">${pending}</div><div class="summary-label">待開立</div></div>
  `;

  let filtered = db.opinions.filter(op => {
    if (status && op.status !== status) return false;
    if (search) {
      const c = findCase(db, op.caseId);
      if (!c || !c.name.toLowerCase().includes(search)) return false;
    }
    return true;
  }).sort((a, b) => a.expiryDate.localeCompare(b.expiryDate));

  document.getElementById('opinion-tbody').innerHTML = filtered.map(op => {
    const c = findCase(db, op.caseId);
    const d = findMember(db, op.doctorId);
    const daysLeft = daysBetween(fmt(now), op.expiryDate);
    const rowClass = op.status === 'expired' ? 'row-danger' : (op.status === 'expiring' ? 'row-warning' : '');
    return `<tr class="${rowClass}">
      <td>${c ? c.name : op.caseId}</td>
      <td>${d ? d.name : '-'}</td>
      <td>${op.issueDate}</td>
      <td>第${op.sequence}次</td>
      <td>${op.homeVisitDate}</td>
      <td>${op.expiryDate}</td>
      <td><strong>${daysLeft > 0 ? daysLeft + '天' : '已過期'}</strong></td>
      <td>${op.diseaseStatus}</td>
      <td>${op.yearCount}/2</td>
      <td>${opinionStatusBadge(op.status)}</td>
      <td><button class="btn btn-xs btn-primary" onclick="renewOpinion('${op.caseId}')">更新</button></td>
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
          ${activeCases.map(c => `<option value="${c.id}">${c.name} (${c.id})</option>`).join('')}
        </select>
      </div>
      <div class="form-group">
        <label class="form-label">開立醫師 *</label>
        <select class="form-select" id="f-op-doctor" style="width:100%">
          ${docs.map(d => `<option value="${d.id}">${d.name}</option>`).join('')}
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
    return `<tr>
      <td><input type="checkbox" class="billing-cb" data-id="${b.id}" ${b.status==='pending'?'':'disabled'}></td>
      <td>${b.billingMonth}</td>
      <td>${c ? c.name : b.caseId}</td>
      <td><span class="badge badge-info">${b.code}</span></td>
      <td>${b.serviceDate}</td>
      <td>${m ? m.name : '-'}</td>
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
    csv += `${b.billingMonth},${b.caseId},${c?c.name:''},${b.code},${b.serviceDate},${m?m.name:''},${b.amount},${b.status}\n`;
  });
  downloadCSV(csv, `申報清冊_${month}.csv`);
}

// ==========================================
//  KPI 績效
// ==========================================
function computeKPI(doctorId) {
  const active = getActiveCases(db).filter(c => !doctorId || c.doctorId === doctorId);
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
  const relevantMembers = db.members.filter(m => m.status === 'active' && (!doctorId || m.id === doctorId));
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
  const doctorId = document.getElementById('kpi-filter-doctor')?.value || '';
  const kpiData = computeKPI(doctorId);

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
      <td><strong>${d.name}</strong></td>
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
      <td><strong>${c.name}</strong></td>
      <td>${d ? d.name : '-'}</td>
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
    <p style="margin-bottom:1rem;font-weight:600">個案：${c.name} (${c.id})</p>
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
        level: 'danger', title: `${c.name} 本月尚未服務`,
        desc: `上次服務: ${lastSvc ? lastSvc.date : '無紀錄'}`,
        page: 'services', caseId: c.id,
      });
    }

    // 意見書過期或即將到期
    const op = getLatestOpinion(db, c.id);
    if (!op) {
      const enrollDays = daysBetween(c.enrollDate, fmt(now));
      if (enrollDays > 14) {
        alerts.push({
          level: 'danger', title: `${c.name} 意見書逾期未開立`,
          desc: `收案日 ${c.enrollDate}，已超過14天`,
          page: 'opinion', caseId: c.id,
        });
      } else {
        alerts.push({
          level: 'warning', title: `${c.name} 意見書待開立`,
          desc: `收案日 ${c.enrollDate}，剩餘 ${14 - enrollDays} 天`,
          page: 'opinion', caseId: c.id,
        });
      }
    } else if (op.status === 'expired') {
      alerts.push({
        level: 'danger', title: `${c.name} 意見書已過期`,
        desc: `到期日 ${op.expiryDate}，請盡速更新`,
        page: 'opinion', caseId: c.id,
      });
    } else if (op.status === 'expiring') {
      alerts.push({
        level: 'warning', title: `${c.name} 意見書即將到期`,
        desc: `到期日 ${op.expiryDate}`,
        page: 'opinion', caseId: c.id,
      });
    }

    // 4個月未家訪
    const lastHome = db.services.filter(s => s.caseId === c.id && s.type === 'home')
      .sort((a,b) => b.date.localeCompare(a.date))[0];
    if (lastHome) {
      const daysSince = daysBetween(lastHome.date, fmt(now));
      if (daysSince > 120) {
        alerts.push({
          level: 'warning', title: `${c.name} 超過4個月未家訪`,
          desc: `上次家訪 ${lastHome.date}，已 ${daysSince} 天`,
          page: 'services', caseId: c.id,
        });
      }
    }

    // ACP 收案滿6月未說明
    const sixMonthAgo = new Date(now); sixMonthAgo.setMonth(sixMonthAgo.getMonth() - 6);
    if (new Date(c.enrollDate) <= sixMonthAgo && !c.acpExplained) {
      alerts.push({
        level: 'info', title: `${c.name} 收案滿6月尚未完成ACP說明`,
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
          level: 'danger', title: `${m.name} ACP訓練逾期`,
          desc: `加入日 ${m.joinDate}，已超過6個月仍未完成`,
          page: 'members',
        });
      } else {
        alerts.push({
          level: 'warning', title: `${m.name} ACP訓練待完成`,
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
        <div class="alert-title">${a.title}</div>
        <div class="alert-desc">${a.desc}</div>
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
      const kpi = computeKPI();
      kpi.forEach(k => csv += `${k.label},${k.value.toFixed(1)}%\n`);
      break;
    }
    case 'case_list': {
      filename = `個案清冊_${ym}.csv`;
      csv += '個案編號,姓名,性別,年齡,CMS等級,負責醫師,個管師,收案日期,狀態,疾病診斷,ACP狀態\n';
      db.cases.forEach(c => {
        const d = findMember(db, c.doctorId);
        const n = findMember(db, c.nurseId);
        csv += `${c.id},${c.name},${c.gender==='M'?'男':'女'},${c.age},${c.cmsLevel},${d?d.name:''},${n?n.name:''},${c.enrollDate},${c.status==='active'?'收案中':'已結案'},"${c.diagnoses.map(d=>d.name).join(';')}",${c.acpExplained?'已說明':'未說明'}\n`;
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
        csv += `${d.name},醫師,${getCaseCountByMember(db,d.id)},${kpi[0].value.toFixed(0)}%,${kpi[1].value.toFixed(0)}%,${kpi[2].value.toFixed(0)}%,${d.acpTrained?'是':'否'},${kpi[4].value.toFixed(0)}%\n`;
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

function importData() {
  document.getElementById('import-file-input').click();
}

function handleImport(e) {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function(ev) {
    try {
      const data = JSON.parse(ev.target.result);
      if (data.cases && data.members) {
        db = data;
        saveDB(db);
        showToast('資料匯入成功');
        switchPage('dashboard');
      } else {
        showToast('檔案格式不正確', 'danger');
      }
    } catch(err) {
      showToast('檔案解析失敗', 'danger');
    }
  };
  reader.readAsText(file);
  e.target.value = '';
}
