/* ========================================
   Demo 資料與資料管理層
   ======================================== */

const DB_KEY = 'homecare_db';

// ===== 初始 Demo 資料 =====
function generateDemoData() {
  const now = new Date();
  const y = now.getFullYear();
  const m = now.getMonth();

  // --- 醫師 ---
  const doctors = [
    { id:'D001', name:'王大明', role:'doctor', specialty:'家醫科', phone:'0912-345-678',
      acpTrained: true, acpTrainedDate:'2023-08-15', opinionTrained: true, opinionTrainedDate:'2023-07-20',
      joinDate:'2023-07-01', status:'active' },
    { id:'D002', name:'李美玲', role:'doctor', specialty:'內科', phone:'0923-456-789',
      acpTrained: true, acpTrainedDate:'2023-09-10', opinionTrained: true, opinionTrainedDate:'2023-08-05',
      joinDate:'2023-07-15', status:'active' },
    { id:'D003', name:'陳志偉', role:'doctor', specialty:'神經科', phone:'0934-567-890',
      acpTrained: true, acpTrainedDate:'2023-10-01', opinionTrained: true, opinionTrainedDate:'2023-09-15',
      joinDate:'2023-08-01', status:'active' },
    { id:'D004', name:'張淑芬', role:'doctor', specialty:'家醫科', phone:'0945-678-901',
      acpTrained: false, acpTrainedDate: null, opinionTrained: true, opinionTrainedDate:'2023-11-20',
      joinDate:'2023-11-01', status:'active' },
    { id:'D005', name:'林建宏', role:'doctor', specialty:'復健科', phone:'0956-789-012',
      acpTrained: true, acpTrainedDate:'2024-01-10', opinionTrained: true, opinionTrainedDate:'2023-12-15',
      joinDate:'2023-12-01', status:'active' },
  ];

  // --- 個管師 ---
  const nurses = [
    { id:'N001', name:'黃雅琪', role:'nurse', specialty:'護理師', phone:'0911-111-222',
      acpTrained: true, acpTrainedDate:'2023-08-20', opinionTrained: true, opinionTrainedDate: null,
      joinDate:'2023-07-01', status:'active' },
    { id:'N002', name:'吳佩珊', role:'nurse', specialty:'護理師', phone:'0922-222-333',
      acpTrained: true, acpTrainedDate:'2023-09-15', opinionTrained: true, opinionTrainedDate: null,
      joinDate:'2023-07-15', status:'active' },
    { id:'N003', name:'周美惠', role:'nurse', specialty:'護理師', phone:'0933-333-444',
      acpTrained: false, acpTrainedDate: null, opinionTrained: true, opinionTrainedDate: null,
      joinDate:'2023-10-01', status:'active' },
    { id:'N004', name:'鄭雅文', role:'nurse', specialty:'護理師', phone:'0944-444-555',
      acpTrained: true, acpTrainedDate:'2024-01-05', opinionTrained: true, opinionTrainedDate: null,
      joinDate:'2023-12-01', status:'active' },
  ];

  const members = [...doctors, ...nurses];

  // --- 疾病代碼對照 ---
  const diseases = [
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

  // --- 個案 ---
  const lastNames = ['陳','林','黃','張','劉','王','蔡','楊','許','郭','吳','謝','鄭','曾','賴'];
  const firstNamesM = ['進財','金發','文雄','正義','國興','清水','阿土','福來','萬福','天賜'];
  const firstNamesF = ['秀英','美玉','麗華','月娥','桂花','金枝','玉蘭','阿珠','素梅','春嬌'];
  const addresses = [
    '台北市大安區和平東路一段100號','新北市板橋區中山路200號','台北市萬華區西園路三段50號',
    '新北市中和區景安路88號','台北市信義區基隆路二段33號','新北市永和區永和路66號',
    '台北市中正區南昌路一段25號','新北市新莊區中正路150號','台北市松山區南京東路五段77號',
    '新北市三重區重新路四段22號','台北市文山區木柵路三段44號','新北市土城區金城路55號',
    '台北市北投區中和街180號','新北市樹林區中華路120號','新北市蘆洲區長安街36號',
    '台北市大同區民生西路200號','新北市汐止區大同路一段100號','台北市內湖區成功路四段88號',
    '新北市淡水區中正東路65號','台北市士林區中山北路六段40號',
  ];

  const cases = [];
  for (let i = 0; i < 48; i++) {
    const isFemale = Math.random() > 0.45;
    const ln = lastNames[i % lastNames.length];
    const fn = isFemale ? firstNamesF[i % firstNamesF.length] : firstNamesM[i % firstNamesM.length];
    const age = 65 + Math.floor(Math.random() * 30);
    const level = 2 + Math.floor(Math.random() * 7);
    const dIdx = i % doctors.length;
    const nIdx = i % nurses.length;
    const enroll = new Date(y, m - Math.floor(Math.random() * 12) - 1, 1 + Math.floor(Math.random() * 28));
    const isClosed = i >= 44;

    // 疾病
    const numDx = 1 + Math.floor(Math.random() * 3);
    const caseDx = [];
    const used = new Set();
    for (let d = 0; d < numDx; d++) {
      let di;
      do { di = Math.floor(Math.random() * diseases.length); } while (used.has(di));
      used.add(di);
      caseDx.push({ ...diseases[di], onset: ['<6m','6-12m','>1y'][Math.floor(Math.random()*3)] });
    }

    const hasHypertension = caseDx.some(d => d.icd === 'I10') || Math.random() > 0.3;
    const hasDiabetes = caseDx.some(d => d.icd.startsWith('E11'));
    const hasHyperlipidemia = caseDx.some(d => d.icd.startsWith('E78'));

    cases.push({
      id: 'C' + String(i+1).padStart(4,'0'),
      name: ln + fn,
      gender: isFemale ? 'F' : 'M',
      age,
      idNumber: (isFemale ? 'A2' : 'A1') + String(10000000 + Math.floor(Math.random()*89999999)),
      phone: '02-' + String(20000000 + Math.floor(Math.random()*9999999)),
      address: addresses[i % addresses.length],
      cmsLevel: level,
      doctorId: doctors[dIdx].id,
      nurseId: nurses[nIdx].id,
      enrollDate: fmt(enroll),
      status: isClosed ? 'closed' : 'active',
      closeReason: isClosed ? ['死亡','遷居','入住機構','拒絕訪視'][i%4] : null,
      closeDate: isClosed ? fmt(new Date(y, m - 1, 15)) : null,
      diagnoses: caseDx,
      hasHypertension,
      hasDiabetes,
      hasHyperlipidemia,
      diseaseStatus: ['穩定','不穩定','不明'][Math.floor(Math.random()*3)],
      acpExplained: Math.random() > 0.4,
      acpExplainedDate: Math.random() > 0.4 ? fmt(new Date(y, m - Math.floor(Math.random()*6), 10)) : null,
      adExplained: Math.random() > 0.5,
      adExplainedDate: Math.random() > 0.5 ? fmt(new Date(y, m - Math.floor(Math.random()*6), 15)) : null,
      acpSigned: Math.random() > 0.7,
      acpSignedDate: Math.random() > 0.7 ? fmt(new Date(y, m - Math.floor(Math.random()*4), 20)) : null,
      nhiRegistered: Math.random() > 0.8,
      familyAcpExplained: Math.random() > 0.6,
      isRemoteArea: i >= 40 && i < 44,
    });
  }

  // --- 醫師意見書 ---
  const opinions = [];
  const activeCases = cases.filter(c => c.status === 'active');
  activeCases.forEach((c, idx) => {
    const opDate = new Date(c.enrollDate);
    opDate.setDate(opDate.getDate() + Math.floor(Math.random() * 12));
    const expDate = new Date(opDate);
    expDate.setMonth(expDate.getMonth() + 6);

    opinions.push({
      id: 'OP' + String(idx+1).padStart(4,'0'),
      caseId: c.id,
      doctorId: c.doctorId,
      issueDate: fmt(opDate),
      homeVisitDate: fmt(opDate),
      expiryDate: fmt(expDate),
      sequence: 1 + Math.floor(idx / activeCases.length * 2),
      yearCount: 1,
      diseaseStatus: c.diseaseStatus,
      functionalPrognosis: ['退步','穩定','進步','無法確定'][Math.floor(Math.random()*4)],
      status: expDate < now ? 'expired' : (expDate - now < 30*86400000 ? 'expiring' : 'valid'),
    });
  });

  // --- 服務紀錄 ---
  const services = [];
  let sIdx = 0;
  activeCases.forEach(c => {
    const monthsActive = Math.max(1, Math.floor((now - new Date(c.enrollDate)) / (30*86400000)));
    const totalRecords = Math.min(monthsActive, 6);
    for (let mi = 0; mi < totalRecords; mi++) {
      const sDate = new Date(y, m - totalRecords + mi + 1, 5 + Math.floor(Math.random()*20));
      if (sDate > now) continue;
      const types = ['home','phone','video'];
      const type = mi === 0 ? 'home' : (mi % 4 === 0 ? 'home' : types[Math.floor(Math.random()*3)]);
      const bpMeasured = type === 'home' || Math.random() > 0.3;
      const hba1cMonitored = c.hasDiabetes && Math.random() > 0.3;
      const lipidMonitored = c.hasHyperlipidemia && Math.random() > 0.4;

      sIdx++;
      services.push({
        id: 'S' + String(sIdx).padStart(5,'0'),
        caseId: c.id,
        nurseId: c.nurseId,
        doctorId: c.doctorId,
        date: fmt(sDate),
        type,
        respondent: Math.random() > 0.3 ? 'patient' : ['spouse','son','daughter','caregiver_foreign'][Math.floor(Math.random()*4)],
        bpMeasured,
        bpSystolic: bpMeasured ? 110 + Math.floor(Math.random()*50) : null,
        bpDiastolic: bpMeasured ? 60 + Math.floor(Math.random()*30) : null,
        hba1cMonitored,
        hba1cValue: hba1cMonitored ? (5.5 + Math.random()*4).toFixed(1) : null,
        lipidMonitored,
        educationProvided: Math.random() > 0.2,
        acpExplained: Math.random() > 0.7,
        chronicDiseaseEval: Math.random() > 0.3,
        referralLtc: Math.random() > 0.85,
        referralMedical: Math.random() > 0.9,
        notes: '',
        billingStatus: mi < totalRecords - 1 ? 'approved' : (Math.random() > 0.5 ? 'submitted' : 'pending'),
      });
    }
  });

  // --- 費用申報 ---
  const billings = [];
  let bIdx = 0;

  // 意見書費用 AA12
  opinions.forEach(op => {
    bIdx++;
    const c = cases.find(x => x.id === op.caseId);
    const isRemote = c && c.isRemoteArea;
    billings.push({
      id: 'B' + String(bIdx).padStart(5,'0'),
      caseId: op.caseId,
      code: 'AA12',
      serviceDate: op.issueDate,
      billingMonth: op.issueDate.substring(0,7),
      memberId: op.doctorId,
      amount: isRemote ? 1800 : 1500,
      status: op.status === 'expired' ? 'approved' : 'submitted',
    });
  });

  // 個案管理費 YA01-YA04
  services.forEach(s => {
    bIdx++;
    const c = cases.find(x => x.id === s.caseId);
    const isRemote = c && c.isRemoteArea;
    let code, amount;
    if (s.type === 'home') {
      code = isRemote ? 'YA04' : 'YA02';
      amount = isRemote ? 1200 : 1000;
    } else {
      code = isRemote ? 'YA03' : 'YA01';
      amount = isRemote ? 300 : 250;
    }
    billings.push({
      id: 'B' + String(bIdx).padStart(5,'0'),
      caseId: s.caseId,
      code,
      serviceDate: s.date,
      billingMonth: s.date.substring(0,7),
      memberId: s.nurseId,
      amount,
      status: s.billingStatus,
    });
  });

  return { members, cases, opinions, services, billings, diseases };
}

// ===== 工具函數 =====
function fmt(d) {
  if (typeof d === 'string') return d;
  const yy = d.getFullYear();
  const mm = String(d.getMonth()+1).padStart(2,'0');
  const dd = String(d.getDate()).padStart(2,'0');
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
  const raw = localStorage.getItem(DB_KEY);
  if (raw) {
    try { return JSON.parse(raw); } catch(e) {}
  }
  const data = generateDemoData();
  saveDB(data);
  return data;
}

function saveDB(data) {
  localStorage.setItem(DB_KEY, JSON.stringify(data));
}

function resetDB() {
  localStorage.removeItem(DB_KEY);
  return loadDB();
}

// ===== 查詢輔助 =====
function findCase(db, id) { return db.cases.find(c => c.id === id); }
function findMember(db, id) { return db.members.find(m => m.id === id); }
function getDoctors(db) { return db.members.filter(m => m.role === 'doctor'); }
function getNurses(db) { return db.members.filter(m => m.role === 'nurse'); }
function getActiveCases(db) { return db.cases.filter(c => c.status === 'active'); }

function getCaseCountByMember(db, memberId) {
  return db.cases.filter(c => c.status === 'active' && (c.doctorId === memberId || c.nurseId === memberId)).length;
}

function getServicesByCase(db, caseId) {
  return db.services.filter(s => s.caseId === caseId).sort((a,b) => b.date.localeCompare(a.date));
}

function getOpinionsByCase(db, caseId) {
  return db.opinions.filter(o => o.caseId === caseId).sort((a,b) => b.issueDate.localeCompare(a.issueDate));
}

function getLatestOpinion(db, caseId) {
  const ops = getOpinionsByCase(db, caseId);
  return ops.length > 0 ? ops[0] : null;
}

function getServiceThisMonth(db, caseId) {
  const ym = fmt(new Date()).substring(0,7);
  return db.services.filter(s => s.caseId === caseId && s.date.startsWith(ym));
}

function getHomeVisitsThisYear(db, caseId) {
  const yy = String(new Date().getFullYear());
  return db.services.filter(s => s.caseId === caseId && s.type === 'home' && s.date.startsWith(yy));
}
