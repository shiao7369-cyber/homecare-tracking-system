/* ========================================
   Firebase 初始化與雲端資料層
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

const auth = firebase.auth();
const fsdb = firebase.firestore();

// ===== 登入/註冊/登出 =====
function doLogin() {
  const email = document.getElementById('login-email').value.trim();
  const pw = document.getElementById('login-password').value;
  if (!email || !pw) { showLoginError('請輸入 Email 和密碼'); return; }

  auth.signInWithEmailAndPassword(email, pw)
    .then(() => { /* onAuthStateChanged 會處理 */ })
    .catch(err => {
      const msgs = {
        'auth/user-not-found': '此帳號不存在，請先註冊',
        'auth/wrong-password': '密碼錯誤',
        'auth/invalid-email': 'Email 格式不正確',
        'auth/invalid-credential': '帳號或密碼錯誤',
      };
      showLoginError(msgs[err.code] || err.message);
    });
}

function doRegister() {
  const email = document.getElementById('login-email').value.trim();
  const pw = document.getElementById('login-password').value;
  if (!email || !pw) { showLoginError('請輸入 Email 和密碼'); return; }
  if (pw.length < 6) { showLoginError('密碼至少需要 6 個字元'); return; }

  auth.createUserWithEmailAndPassword(email, pw)
    .then(() => { /* onAuthStateChanged 會處理 */ })
    .catch(err => {
      const msgs = {
        'auth/email-already-in-use': '此 Email 已註冊，請直接登入',
        'auth/weak-password': '密碼強度不足，至少 6 個字元',
        'auth/invalid-email': 'Email 格式不正確',
      };
      showLoginError(msgs[err.code] || err.message);
    });
}

function doLogout() {
  auth.signOut();
}

function showLoginError(msg) {
  const el = document.getElementById('login-error');
  el.textContent = msg;
  el.style.display = 'block';
}

// ===== 監聽登入狀態 =====
auth.onAuthStateChanged(user => {
  if (user) {
    // 已登入 → 隱藏登入頁，顯示主系統
    document.getElementById('login-page').style.display = 'none';
    document.getElementById('sidebar').style.display = 'flex';
    document.getElementById('main-content').style.display = 'flex';

    // 更新使用者資訊
    const nameEl = document.querySelector('.user-name');
    if (nameEl) nameEl.textContent = user.email.split('@')[0];

    // 載入雲端資料
    loadCloudData().then(() => {
      initNav();
      initFilters();
      renderDashboard();
      updateAlertBadge();
      document.getElementById('today-date').textContent = formatDateCN(new Date());
    });
  } else {
    // 未登入 → 顯示登入頁
    document.getElementById('login-page').style.display = 'flex';
    document.getElementById('sidebar').style.display = 'none';
    document.getElementById('main-content').style.display = 'none';
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
      // 版本一致，嘗試從雲端載入
      console.log('從雲端載入資料...');
      db = { members: [], cases: [], opinions: [], services: [], billings: [], diseases: [] };
      for (const col of COLLECTIONS) {
        const snapshot = await fsdb.collection(col).get();
        db[col] = snapshot.docs.map(doc => ({ ...doc.data(), _docId: doc.id }));
      }
      console.log(`雲端資料載入: ${db.cases.length} 個案, ${db.members.length} 成員`);

      // 安全檢查：雲端資料為空但本地有真實資料 → 重新初始化
      if (db.cases.length === 0 && typeof RAW_CASES_DATA !== 'undefined' && RAW_CASES_DATA.length > 0) {
        console.warn('雲端資料為空，重新從本地清冊初始化...');
        db = generateDemoData();
        uploadAllData(db).catch(e => console.error('背景上傳失敗:', e));
        saveDB_local(db);
        return;
      }
    }

    if (needsInit) {
      // 版本不符或首次使用：用本地真實資料初始化
      console.log(`雲端資料版本不符 (${cloudVersion} → ${DB_VERSION})，重新初始化...`);
      db = generateDemoData();
      saveDB_local(db);
      console.log(`本地資料已就緒: ${db.cases.length} 個案, ${db.members.length} 成員`);

      // 背景上傳到雲端（不阻塞頁面載入）
      (async () => {
        try {
          // 分批刪除舊資料
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

    // 同步到 localStorage
    saveDB_local(db);
  } catch (err) {
    console.error('雲端載入失敗，使用本地資料:', err);
    db = generateDemoData();
    saveDB_local(db);
  }
}

// 只存 localStorage，不觸發雲端同步
function saveDB_local(data) {
  localStorage.setItem(DB_KEY, JSON.stringify(data));
  localStorage.setItem(DB_VERSION_KEY, DB_VERSION);
}

async function uploadAllData(data) {
  for (const col of COLLECTIONS) {
    const items = data[col] || [];
    const batch_size = 400; // Firestore batch limit is 500
    for (let i = 0; i < items.length; i += batch_size) {
      const batch = fsdb.batch();
      const chunk = items.slice(i, i + batch_size);
      chunk.forEach(item => {
        const docRef = fsdb.collection(col).doc(item.id);
        batch.set(docRef, item);
      });
      await batch.commit();
    }
  }
}

// ===== 覆寫 saveDB：同時存雲端 =====
const _originalSaveDB = typeof saveDB === 'function' ? saveDB : null;

function saveToCloud(collection, item) {
  if (!item || !item.id) return;
  fsdb.collection(collection).doc(item.id).set(item, { merge: true })
    .catch(err => console.error('雲端儲存失敗:', err));
}

// 包裝 saveDB，存本地同時存雲端
function saveDB(data) {
  localStorage.setItem(DB_KEY, JSON.stringify(data));

  // 差異同步到雲端（只更新有改動的資料）
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
