/* ========================================
   真實個案資料與資料管理層
   (資料來源：115年居家失能個案家庭醫師-個案清冊)
   ======================================== */

const DB_KEY = 'homecare_db';
const DB_VERSION_KEY = 'homecare_db_version';
const DB_VERSION = '2.1_doctor_update';

// ===== 真實個案資料（由個案清冊匯入） =====
// 此處為外部載入的 JSON，初次載入時由 loadRawCases() 提供
let _rawCasesCache = null;

function loadRawCases() {
  if (_rawCasesCache) return _rawCasesCache;
  // 嵌入的真實資料將透過 cases-data.js 提供
  return typeof RAW_CASES_DATA !== 'undefined' ? RAW_CASES_DATA : [];
}

// ===== 真實醫師資料 =====
function getRealDoctors() {
  return [
    { id:'D001', name:'李致有', role:'doctor', specialty:'小兒科', phone:'',
      acpTrained: true, acpTrainedDate:'2023-08-15', opinionTrained: true, opinionTrainedDate:'2023-07-20',
      joinDate:'2020-08-01', status:'active' },
    { id:'D002', name:'徐子茜', role:'doctor', specialty:'小兒科', phone:'',
      acpTrained: true, acpTrainedDate:'2023-09-10', opinionTrained: true, opinionTrainedDate:'2023-08-05',
      joinDate:'2020-08-01', status:'active' },
    { id:'D003', name:'蕭輝哲', role:'doctor', specialty:'家醫科', phone:'',
      acpTrained: true, acpTrainedDate:'2023-10-01', opinionTrained: true, opinionTrainedDate:'2023-09-15',
      joinDate:'2020-08-01', status:'active' },
    { id:'D004', name:'翁志仁', role:'doctor', specialty:'家醫科', phone:'',
      acpTrained: true, acpTrainedDate:'2023-11-20', opinionTrained: true, opinionTrainedDate:'2023-11-20',
      joinDate:'2020-08-01', status:'active' },
    { id:'D005', name:'葉步盛', role:'doctor', specialty:'家醫科', phone:'',
      acpTrained: true, acpTrainedDate:'2024-01-10', opinionTrained: true, opinionTrainedDate:'2023-12-15',
      joinDate:'2020-08-01', status:'active' },
    { id:'D006', name:'黃朝麟', role:'doctor', specialty:'家醫科', phone:'',
      acpTrained: true, acpTrainedDate:'2024-02-01', opinionTrained: true, opinionTrainedDate:'2024-01-15',
      joinDate:'2020-08-01', status:'active' },
  ];
}

// 醫師名稱→ID 對照
const DOCTOR_NAME_MAP = {
  '李致有': 'D001', '徐子茜': 'D002', '蕭輝哲': 'D003',
  '翁志仁': 'D004', '葉步盛': 'D005', '黃朝麟': 'D006',
};

// ===== 個管師資料 =====
function getRealNurses() {
  return [
    { id:'N001', name:'個管師A', role:'nurse', specialty:'護理師', phone:'',
      acpTrained: true, acpTrainedDate:'2023-08-20', opinionTrained: true, opinionTrainedDate: null,
      joinDate:'2020-08-01', status:'active' },
    { id:'N002', name:'個管師B', role:'nurse', specialty:'護理師', phone:'',
      acpTrained: true, acpTrainedDate:'2023-09-15', opinionTrained: true, opinionTrainedDate: null,
      joinDate:'2020-08-01', status:'active' },
  ];
}

// ===== 疾病代碼對照 =====
function getDiseases() {
  return [
    { icd:'I10', name:'本態性高血壓' },
    { icd:'E11.9', name:'第二型糖尿病' },
    { icd:'E78.5', name:'高血脂症' },
    { icd:'I63.9', name:'腦梗塞' },
    { icd:'F03.90', name:'失智症' },
    { icd:'G20', name:'帕金森氏病' },
    { icd:'M17.0', name:'雙側膝關節炎' },
    { icd:'J44.9', name:'慢性阻塞性肺疾患' },
    { icd:'I50.9', name:'心衰竭' },
    { icd:'N18.3', name:'慢性腎臟病第三期' },
    { icd:'R54', name:'衰弱症' },
    { icd:'F32.1', name:'中度鬱症' },
    { icd:'M80.00XA', name:'骨質疏鬆併骨折' },
    { icd:'H54.1', name:'低視力' },
    { icd:'I25.10', name:'慢性缺血性心臟病' },
  ];
}

// ===== 將原始清冊資料轉換為系統個案格式 =====
function convertRawCases(rawCases) {
  const cases = [];
  const nurseIds = ['N001', 'N002'];

  rawCases.forEach((raw, i) => {
    const doctorId = DOCTOR_NAME_MAP[raw.doctorName] || 'D001';
    const nurseId = nurseIds[i % nurseIds.length];

    // 解析結案原因
    let closeReason = null;
    let closeDate = null;
    if (raw.status === 'closed' && raw.closeInfo) {
      closeDate = raw.doctorVisitDate || raw.firstReferralDate;
      if (raw.closeInfo.includes('過世') || raw.closeInfo.includes('死亡') || raw.closeInfo.includes('往生')) {
        closeReason = '死亡';
      } else if (raw.closeInfo.includes('機構') || raw.closeInfo.includes('入住')) {
        closeReason = '入住機構';
      } else if (raw.closeInfo.includes('結案') || raw.closeInfo.includes('拒絕')) {
        closeReason = '結案';
      } else {
        closeReason = '其他';
      }
    }

    cases.push({
      id: raw.caseNo || ('C' + String(i + 1).padStart(4, '0')),
      seqNo: raw.seqNo,
      name: raw.name,
      gender: raw.gender,
      age: null, // 清冊未提供年齡
      idNumber: raw.idNumber,
      phone: raw.phone,
      address: raw.address,
      cmsLevel: raw.cmsLevel,
      category: raw.category || '',
      district: raw.district || '',
      contactPerson: raw.contactPerson || '',
      doctorId,
      doctorName: raw.doctorName,
      nurseId,
      enrollDate: raw.firstReferralDate,
      firstReferralDate: raw.firstReferralDate,
      nurseVisitDate: raw.nurseVisitDate,
      doctorVisitDate: raw.doctorVisitDate,
      serviceDays: raw.serviceDays,
      monthlyTracking: raw.monthlyTracking || {},
      scheduledVisit: raw.scheduledVisit || '',
      closeInfo: raw.closeInfo || '',
      status: raw.status || 'active',
      closeReason,
      closeDate,
      notes: raw.notes || '',
      diagnoses: [],
      hasHypertension: false,
      hasDiabetes: false,
      hasHyperlipidemia: false,
      diseaseStatus: '穩定',
      acpExplained: false,
      acpExplainedDate: null,
      adExplained: false,
      adExplainedDate: null,
      acpSigned: false,
      acpSignedDate: null,
      nhiRegistered: false,
      familyAcpExplained: false,
      isRemoteArea: false,
    });
  });

  return cases;
}

// ===== 從月度追蹤產生服務紀錄 =====
function generateServicesFromTracking(cases) {
  const services = [];
  let sIdx = 0;
  const year = 2026; // 115年 = 2026

  cases.forEach(c => {
    if (!c.monthlyTracking) return;

    Object.entries(c.monthlyTracking).forEach(([monthStr, text]) => {
      const month = parseInt(monthStr);
      if (!text || month < 1 || month > 12) return;

      // 解析日期與類型
      let type = 'phone';
      let day = 2;
      const t = text.trim();

      if (t === '結') return; // 結案標記，非服務紀錄

      if (t.includes('家')) type = 'home';
      else if (t.includes('電')) type = 'phone';
      else if (t.includes('視')) type = 'video';
      else if (t.includes('訪')) type = 'home';

      // 解析日期 (e.g., "1/21家" → day=21)
      const dateMatch = t.match(/(\d+)\/(\d+)/);
      if (dateMatch) {
        day = parseInt(dateMatch[2]) || 2;
      }

      // 只產生今天以前的紀錄
      const sDate = new Date(year, month - 1, day);
      if (sDate > new Date()) return;

      sIdx++;
      services.push({
        id: 'S' + String(sIdx).padStart(5, '0'),
        caseId: c.id,
        nurseId: c.nurseId,
        doctorId: c.doctorId,
        date: fmt(sDate),
        type,
        respondent: 'patient',
        bpMeasured: type === 'home',
        bpSystolic: type === 'home' ? (120 + Math.floor(Math.random() * 30)) : null,
        bpDiastolic: type === 'home' ? (70 + Math.floor(Math.random() * 20)) : null,
        hba1cMonitored: false,
        hba1cValue: null,
        lipidMonitored: false,
        educationProvided: type === 'home',
        acpExplained: false,
        chronicDiseaseEval: type === 'home',
        referralLtc: false,
        referralMedical: false,
        notes: '',
        billingStatus: 'approved',
      });
    });
  });

  return services;
}

// ===== 從醫師家訪日期產生意見書 =====
function generateOpinionsFromCases(cases) {
  const opinions = [];
  let opIdx = 0;
  const now = new Date();

  cases.forEach(c => {
    if (!c.doctorVisitDate) return;

    const opDate = new Date(c.doctorVisitDate);
    const expDate = new Date(opDate);
    expDate.setMonth(expDate.getMonth() + 6);

    opIdx++;
    let status;
    if (expDate < now) status = 'expired';
    else if (expDate - now < 30 * 86400000) status = 'expiring';
    else status = 'valid';

    opinions.push({
      id: 'OP' + String(opIdx).padStart(4, '0'),
      caseId: c.id,
      doctorId: c.doctorId,
      issueDate: c.doctorVisitDate,
      homeVisitDate: c.doctorVisitDate,
      expiryDate: fmt(expDate),
      sequence: 1,
      yearCount: 1,
      diseaseStatus: c.diseaseStatus || '穩定',
      functionalPrognosis: '穩定',
      status,
    });
  });

  return opinions;
}

// ===== 產生費用申報 =====
function generateBillingRecords(cases, opinions, services) {
  const billings = [];
  let bIdx = 0;

  // 意見書費用 AA12
  opinions.forEach(op => {
    bIdx++;
    billings.push({
      id: 'B' + String(bIdx).padStart(5, '0'),
      caseId: op.caseId,
      code: 'AA12',
      serviceDate: op.issueDate,
      billingMonth: op.issueDate.substring(0, 7),
      memberId: op.doctorId,
      amount: 1500,
      status: op.status === 'expired' ? 'approved' : 'submitted',
    });
  });

  // 個案管理費
  services.forEach(s => {
    bIdx++;
    let code, amount;
    if (s.type === 'home') {
      code = 'YA02'; amount = 1000;
    } else {
      code = 'YA01'; amount = 250;
    }
    billings.push({
      id: 'B' + String(bIdx).padStart(5, '0'),
      caseId: s.caseId,
      code,
      serviceDate: s.date,
      billingMonth: s.date.substring(0, 7),
      memberId: s.nurseId,
      amount,
      status: s.billingStatus,
    });
  });

  return billings;
}

// ===== 成員資料獨立儲存 =====
const MEMBERS_KEY = 'homecare_members';

// 醫師專科正確對照表（以此為準）
const DOCTOR_SPECIALTY_MAP = {
  '李致有': '小兒科', '徐子茜': '小兒科',
  '蕭輝哲': '內科', '翁志仁': '內科',
  '葉步盛': '內科', '黃朝麟': '內科',
};

function loadMembers() {
  let members = null;

  // 優先讀取獨立儲存的成員資料
  try {
    const saved = localStorage.getItem(MEMBERS_KEY);
    if (saved) {
      const parsed = JSON.parse(saved);
      if (parsed.length > 0) members = parsed;
    }
  } catch (e) { }

  // 舊版 DB 中可能有成員資料
  if (!members) {
    try {
      const saved = localStorage.getItem(DB_KEY);
      if (saved) {
        const oldDb = JSON.parse(saved);
        if (oldDb.members && oldDb.members.length > 0) members = oldDb.members;
      }
    } catch (e) { }
  }

  // 都沒有，用預設值
  if (!members) members = [...getRealDoctors(), ...getRealNurses()];

  // 自動修正醫師專科（以 DOCTOR_SPECIALTY_MAP 為準）
  members.forEach(m => {
    if (DOCTOR_SPECIALTY_MAP[m.name] && m.specialty !== DOCTOR_SPECIALTY_MAP[m.name]) {
      m.specialty = DOCTOR_SPECIALTY_MAP[m.name];
    }
  });

  return members;
}

function saveMembers(members) {
  localStorage.setItem(MEMBERS_KEY, JSON.stringify(members));
}

// ===== 產生完整資料庫 =====
function generateDemoData() {
  const rawCases = loadRawCases();
  const members = loadMembers();
  const diseases = getDiseases();

  const cases = convertRawCases(rawCases);
  const opinions = generateOpinionsFromCases(cases.filter(c => c.status === 'active'));
  const services = generateServicesFromTracking(cases);
  const billings = generateBillingRecords(cases, opinions, services);

  // 同時備份成員資料到獨立 key
  saveMembers(members);

  return { members, cases, opinions, services, billings, diseases };
}

// ===== 工具函數 =====
function fmt(d) {
  if (typeof d === 'string') return d;
  const yy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  return `${yy}-${mm}-${dd}`;
}

function daysBetween(a, b) {
  return Math.floor((new Date(b) - new Date(a)) / 86400000);
}

function monthsBetween(a, b) {
  const da = new Date(a), db = new Date(b);
  return (db.getFullYear() - da.getFullYear()) * 12 + db.getMonth() - da.getMonth();
}

// ===== 資料存取 =====
function loadDB() {
  const savedVersion = localStorage.getItem(DB_VERSION_KEY);
  if (savedVersion !== DB_VERSION) {
    localStorage.removeItem(DB_KEY);
    localStorage.setItem(DB_VERSION_KEY, DB_VERSION);
  }
  const raw = localStorage.getItem(DB_KEY);
  if (raw) {
    try { return JSON.parse(raw); } catch (e) { }
  }
  const data = generateDemoData();
  saveDB(data);
  return data;
}

function saveDB(data) {
  invalidateCache();
  localStorage.setItem(DB_KEY, JSON.stringify(data));
}

function resetDB() {
  invalidateCache();
  localStorage.removeItem(DB_KEY);
  return loadDB();
}

// ===== 查詢快取層 =====
// 所有 lookup maps 集中管理，資料變更時呼叫 invalidateCache() 重建
let _cache = null;

function invalidateCache() {
  _cache = null;
  if (typeof invalidateAlerts === 'function') invalidateAlerts();
}

function _ensureCache(db) {
  if (_cache) return _cache;
  const c = {};

  // memberMap: id → member
  c.memberMap = new Map();
  db.members.forEach(m => c.memberMap.set(m.id, m));

  // caseMap: id → case
  c.caseMap = new Map();
  db.cases.forEach(cs => c.caseMap.set(cs.id, cs));

  // doctors / nurses / activeCases (常用子集)
  c.doctors = db.members.filter(m => m.role === 'doctor');
  c.nurses = db.members.filter(m => m.role === 'nurse');
  c.activeCases = db.cases.filter(cs => cs.status === 'active');

  // servicesByCase: caseId → [services] (按日期降冪)
  c.servicesByCase = new Map();
  db.services.forEach(s => {
    if (!c.servicesByCase.has(s.caseId)) c.servicesByCase.set(s.caseId, []);
    c.servicesByCase.get(s.caseId).push(s);
  });
  c.servicesByCase.forEach((arr) => arr.sort((a, b) => b.date.localeCompare(a.date)));

  // opinionsByCase: caseId → [opinions] (按日期降冪)
  c.opinionsByCase = new Map();
  db.opinions.forEach(o => {
    if (!c.opinionsByCase.has(o.caseId)) c.opinionsByCase.set(o.caseId, []);
    c.opinionsByCase.get(o.caseId).push(o);
  });
  c.opinionsByCase.forEach((arr) => arr.sort((a, b) => b.issueDate.localeCompare(a.issueDate)));

  // 預計算：本月年月字串
  c.ym = fmt(new Date()).substring(0, 7);
  c.yy = String(new Date().getFullYear());

  // caseCountByMember: memberId → count
  c.caseCountByMember = new Map();
  db.members.forEach(m => c.caseCountByMember.set(m.id, 0));
  db.cases.forEach(cs => {
    if (cs.status !== 'active') return;
    if (cs.doctorId && c.caseCountByMember.has(cs.doctorId))
      c.caseCountByMember.set(cs.doctorId, c.caseCountByMember.get(cs.doctorId) + 1);
    if (cs.nurseId && c.caseCountByMember.has(cs.nurseId))
      c.caseCountByMember.set(cs.nurseId, c.caseCountByMember.get(cs.nurseId) + 1);
    // doctorName fallback
    if (cs.doctorName && !cs.doctorId) {
      const doc = db.members.find(m => m.name === cs.doctorName);
      if (doc) c.caseCountByMember.set(doc.id, c.caseCountByMember.get(doc.id) + 1);
    }
  });

  _cache = c;
  return c;
}

// ===== 查詢輔助（使用快取） =====
function findCase(db, id) { return _ensureCache(db).caseMap.get(id) || null; }
function findMember(db, id) { return _ensureCache(db).memberMap.get(id) || null; }
function getDoctors(db) { return _ensureCache(db).doctors; }
function getNurses(db) { return _ensureCache(db).nurses; }
function getActiveCases(db) { return _ensureCache(db).activeCases; }

function getCaseCountByMember(db, memberId) {
  return _ensureCache(db).caseCountByMember.get(memberId) || 0;
}

function getServicesByCase(db, caseId) {
  return _ensureCache(db).servicesByCase.get(caseId) || [];
}

function getOpinionsByCase(db, caseId) {
  return _ensureCache(db).opinionsByCase.get(caseId) || [];
}

function getLatestOpinion(db, caseId) {
  const ops = getOpinionsByCase(db, caseId);
  return ops.length > 0 ? ops[0] : null;
}

function getServiceThisMonth(db, caseId) {
  const ym = _ensureCache(db).ym;
  const svcs = _ensureCache(db).servicesByCase.get(caseId);
  return svcs ? svcs.filter(s => s.date.startsWith(ym)) : [];
}

function getHomeVisitsThisYear(db, caseId) {
  const yy = _ensureCache(db).yy;
  const svcs = _ensureCache(db).servicesByCase.get(caseId);
  return svcs ? svcs.filter(s => s.type === 'home' && s.date.startsWith(yy)) : [];
}

function getPhoneVisitsThisMonth(db, caseId) {
  const ym = _ensureCache(db).ym;
  const svcs = _ensureCache(db).servicesByCase.get(caseId);
  return svcs ? svcs.filter(s => s.type === 'phone' && s.date.startsWith(ym)) : [];
}

function getLastNurseHomeVisit(db, caseId) {
  const svcs = _ensureCache(db).servicesByCase.get(caseId);
  if (!svcs) return null;
  return svcs.find(s => s.type === 'home' && s.nurseId) || null;
}

function getOpinionExpiryInfo(db, caseId) {
  const op = getLatestOpinion(db, caseId);
  if (!op) return { opinion: null, daysLeft: -1, needsRenewal: true };
  const expiry = new Date(op.expiryDate);
  const daysLeft = Math.ceil((expiry - new Date()) / (1000 * 60 * 60 * 24));
  return { opinion: op, daysLeft, needsRenewal: daysLeft <= 30 || op.status === 'expired' };
}
