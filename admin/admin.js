// ── Excel Parser (replicates import_warz_excel.py) ───
const _SKIP_SHEETS    = new Set(['Data', 'Template', 'รายชื่อกิจกรรม']);
const _EXCL_EVENTS    = new Set(['ภารกิจค้นหาเสบียงลับ']);
const _HDR_LITERALS   = new Set(['item id','item en','item type','own period','amount','image url','image','bundle']);
const _ICONS          = ['broadcast','share','calendar','clock'];

function _cellText(v) {
  if (v == null) return '';
  if (v instanceof Date) return v.toLocaleDateString('th-TH');
  const s = String(v).trim();
  if (s.endsWith('.0') && /^\d+$/.test(s.slice(0, -2))) return s.slice(0, -2);
  return s;
}

function _cleanName(name) {
  return name.replace(/[\u{1F49B}\u{1F49C}⚙️\u{1F381}\u{1F579}\u{1F310}✅\u{1F4CC}\u{1FA75}\u{1F9E1}]/gu, '').trim() || name.trim();
}

function _isHdrVal(v) { return _HDR_LITERALS.has(_cellText(v).toLowerCase()); }

function _findHeaderRow(rows) {
  for (let i = 0; i < rows.length; i++) {
    const vals = rows[i].map(_cellText);
    if (vals.includes('NO.') && vals.includes('Reward') && vals.includes('Facebook') && vals.includes('UID')) return i;
  }
  return -1;
}

// itemMap: optional Map<id, warzItem> for cross-reference validation
function _buildRewardInfo(row, itemMap) {
  const reward = _cellText(row[1]);
  let itemId   = _cellText(row[10]);
  let itemEn   = _cellText(row[11]);
  let ownPer   = _cellText(row[12]);
  let itemType = _cellText(row[13]);
  let amount   = _cellText(row[14]);
  let imgUrl   = _cellText(row[16]);

  if (_isHdrVal(itemId) || _isHdrVal(itemEn) || _isHdrVal(amount) || _isHdrVal(imgUrl)) {
    itemId = itemEn = itemType = ownPer = amount = imgUrl = '';
  }

  if (!reward && !itemId && !itemEn && !imgUrl) return null;

  // Cross-reference with warz_data if itemMap provided and itemId is present
  if (itemMap && itemMap.size > 0 && itemId) {
    const warzItem = itemMap.get(itemId);
    if (warzItem) {
      return {
        name: warzItem.name || reward || 'รางวัลกิจกรรม',
        forumReward: reward || warzItem.name || 'รางวัลกิจกรรม',
        itemEn: warzItem.name || itemEn,
        itemId,
        itemType: warzItem.type || itemType,
        ownPeriod: ownPer,
        amount: amount || '1',
        imageUrl: warzItem.image || imgUrl,
        hasItem: true,
        warzVerified: true,
      };
    }
    // Item ID present but not in warz_data — flag as unverified
    return {
      name: reward || itemEn || 'รางวัลกิจกรรม',
      forumReward: reward || 'รางวัลกิจกรรม',
      itemEn, itemId, itemType, ownPeriod: ownPer,
      amount: amount || '1', imageUrl: imgUrl,
      hasItem: !!(itemEn || itemId || imgUrl),
      warzVerified: false,
    };
  }

  return {
    name: reward || 'รางวัลกิจกรรม',
    forumReward: reward || 'รางวัลกิจกรรม',
    itemEn, itemId, itemType, ownPeriod: ownPer,
    amount: amount || '1', imageUrl: imgUrl,
    hasItem: !!(itemEn || itemId || imgUrl),
  };
}

function _findMasterCodes(rows, title, status, expiresAt) {
  const codes = [];
  const re = /(?:master\s*code|item\s*code)\s*:\s*([A-Z0-9-]+)/i;
  for (const row of rows) {
    for (const cell of row) {
      const m = _cellText(cell).match(re);
      if (m) codes.push({ code: m[1].trim(), eventName: title, status: status || 'พร้อมใช้', expiresAt: expiresAt || '-', items: [] });
    }
  }
  return codes;
}

function parseExcelWorkbook(wb, itemMap) {
  const events = [], masterCodes = [];
  let unverifiedCount = 0;

  for (const sheetName of wb.SheetNames) {
    if (_SKIP_SHEETS.has(sheetName)) continue;
    const title = _cleanName(sheetName);
    if (_EXCL_EVENTS.has(title)) continue;

    const sheet = wb.Sheets[sheetName];
    const rows  = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false });

    const deliveryDate = _cellText(rows[1]?.[1]);
    const workStatus   = _cellText(rows[2]?.[1]);
    const claimStart   = _cellText(rows[4]?.[2]);
    const claimEnd     = _cellText(rows[5]?.[2]);
    const cutoff       = _cellText(rows[2]?.[4]);
    const owner        = _cellText(rows[2]?.[3]);

    masterCodes.push(..._findMasterCodes(rows, title, workStatus, claimEnd));

    const headerIdx = _findHeaderRow(rows);
    if (headerIdx === -1) continue;

    const participants = {};

    for (let i = headerIdx + 1; i < rows.length; i++) {
      const row      = rows[i];
      const no       = _cellText(row[0]);
      const reward   = _cellText(row[1]);
      const facebook = _cellText(row[2]);
      const uid      = _cellText(row[3]);
      const method   = _cellText(row[4]);
      const status   = _cellText(row[5]);
      const note     = _cellText(row[7]);

      if (!no && !reward && !facebook && !uid) continue;
      if (reward === 'Reward' || no === 'NO.') continue;

      const rInfo = _buildRewardInfo(row, itemMap) || {
        name: reward || 'รางวัลกิจกรรม', forumReward: reward || 'รางวัลกิจกรรม',
        itemEn: '', itemId: '', itemType: '', ownPeriod: '', amount: '1', imageUrl: '', hasItem: false,
      };

      if (rInfo.warzVerified === false) unverifiedCount++;

      if (facebook || uid) {
        const key = (uid || facebook).trim().toLowerCase();
        if (!participants[key]) {
          participants[key] = { uid, facebook: facebook || '-', character: uid || '-', claimMethod: method, claimStatus: status || workStatus || 'กำลังดำเนินการ', updatedAt: deliveryDate || claimStart || '', note, rewards: [] };
        } else {
          const e = participants[key];
          if (facebook && e.facebook === '-') e.facebook = facebook;
          if (uid && !e.uid) e.uid = uid;
          if (method && !e.claimMethod) e.claimMethod = method;
          if (status) e.claimStatus = status;
          if (note) e.note = note;
        }
        const entry  = participants[key];
        const rKey   = rInfo.itemId || rInfo.imageUrl || rInfo.itemEn || rInfo.forumReward;
        const exists = new Set(entry.rewards.map(r => r.itemId || r.imageUrl || r.itemEn || r.forumReward));
        if (rKey && !exists.has(rKey)) entry.rewards.push(rInfo);
      }
    }

    const num     = events.length + 1;
    const winners = Object.values(participants).sort((a, b) => (a.note ? 1 : 0) - (b.note ? 1 : 0));
    events.push({
      id: `event-${Date.now()}-${num}`,
      name: title, shortName: title.slice(0, 18),
      icon: _ICONS[(num - 1) % _ICONS.length],
      cycle: `กิจกรรมที่ ${num}`,
      period: `${claimStart || '-'} - ${claimEnd || '-'}`,
      resetDate: `ตัดรอบ/จัดส่ง: ${cutoff || 'ทุกวันพุธ'}`,
      latest: deliveryDate || 'รออัปเดต',
      status: workStatus || 'กำลังดำเนินการ',
      owner, reward: 'ดูรางวัลในรายชื่อผู้ได้รับรางวัล',
      winners, pendingRewards: [],
    });
  }

  return { events, codes: masterCodes, unverifiedCount };
}

async function importExcelFile(file) {
  const buf = await file.arrayBuffer();
  const wb  = XLSX.read(buf, { type: 'array', cellDates: true });
  return parseExcelWorkbook(wb, state.itemMap);
}

// ── Constants ─────────────────────────────────────────
const AUTH_KEY      = 'warz_admin_hash';
const SESSION_KEY   = 'warz_admin_session';
const DATA_KEY      = 'warz_admin_data';
const SAVED_KEY     = 'warz_admin_saved_at';
const CRED_VER_KEY  = 'warz_cred_ver';
const CRED_VER      = 'mondkub888-v1';
const ITEMS_DB_KEY  = 'warz_items_db';

// ── State ─────────────────────────────────────────────
const state = {
  events: [],
  codes: [],
  items: [],
  itemMap: new Map(),
  currentEventId: null,
};

// ── Crypto ────────────────────────────────────────────
async function sha256(text) {
  const buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(text));
  return Array.from(new Uint8Array(buf)).map(b => b.toString(16).padStart(2, '0')).join('');
}

// ── Auth ──────────────────────────────────────────────
function isLoggedIn() {
  return sessionStorage.getItem(SESSION_KEY) === 'active';
}

async function ensureDefaultCredentials() {
  if (localStorage.getItem(CRED_VER_KEY) !== CRED_VER) {
    const hash = await sha256('mondkub888:676767');
    localStorage.setItem(AUTH_KEY, hash);
    localStorage.setItem(CRED_VER_KEY, CRED_VER);
    sessionStorage.removeItem(SESSION_KEY);
  }
}

async function tryLogin(user, pass) {
  const stored = localStorage.getItem(AUTH_KEY);
  const hash = await sha256(user + ':' + pass);
  if (hash === stored) {
    sessionStorage.setItem(SESSION_KEY, 'active');
    return { ok: true };
  }
  return { error: 'Username หรือ Password ไม่ถูกต้อง' };
}

async function saveCredentials(user, pass) {
  const hash = await sha256(user + ':' + pass);
  localStorage.setItem(AUTH_KEY, hash);
  localStorage.setItem(CRED_VER_KEY, CRED_VER);
  sessionStorage.setItem(SESSION_KEY, 'active');
}

function logout() {
  sessionStorage.removeItem(SESSION_KEY);
  location.reload();
}

// ── Data ──────────────────────────────────────────────
function cleanLoadedStatuses() {
  state.events.forEach(ev => (ev.winners || []).forEach(w => {
    if (w.claimStatus) w.claimStatus = normalizeStatus(w.claimStatus);
  }));
}

function loadData() {
  const raw = localStorage.getItem(DATA_KEY);
  if (raw) {
    try {
      const parsed = JSON.parse(raw);
      state.events = parsed.events ?? [];
      state.codes  = parsed.codes  ?? [];
      cleanLoadedStatuses();
      return;
    } catch (_) {}
  }
  state.events = JSON.parse(JSON.stringify(window.WARZ_EVENTS        ?? []));
  state.codes  = JSON.parse(JSON.stringify(window.WARZ_MASTER_CODES  ?? []));
  cleanLoadedStatuses();
}

let _undoStack = [];

function persistData(silent = false) {
  // Snapshot the previous version (still in localStorage) for undo
  const prev = localStorage.getItem(DATA_KEY);
  if (prev) {
    _undoStack.push(prev);
    if (_undoStack.length > 40) _undoStack.shift();
  }
  localStorage.setItem(DATA_KEY, JSON.stringify({ events: state.events, codes: state.codes }));
  const now = new Date().toLocaleString('th-TH');
  localStorage.setItem(SAVED_KEY, now);
  updateSavedLabel();
  updateUndoButton();
  if (!silent) toast('บันทึกแล้ว ✓');
}

function updateUndoButton() {
  const btn = document.getElementById('undo-btn');
  if (!btn) return;
  btn.disabled = _undoStack.length === 0;
  btn.textContent = _undoStack.length ? `↶ ย้อนกลับ (${_undoStack.length})` : '↶ ย้อนกลับ';
}

function undoLast() {
  if (!_undoStack.length) return;
  const prev = _undoStack.pop();
  localStorage.setItem(DATA_KEY, prev);
  try {
    const parsed = JSON.parse(prev);
    state.events = parsed.events ?? [];
    state.codes  = parsed.codes  ?? [];
  } catch (_) {}
  updateUndoButton();
  refreshActiveView();
  toast('ย้อนกลับแล้ว ↶');
}

function refreshActiveView() {
  const active = document.querySelector('.admin-view.is-active');
  const id = active?.id || '';
  if (id === 'view-winners') renderWinners();
  else if (id === 'view-events') renderEvents();
  else if (id === 'view-codes') renderCodes();
  else if (id === 'view-items') renderItems();
  else if (id === 'view-dashboard') renderDashboard();
  else if (id === 'view-export') renderExport();
}

// ── Backup / Restore ──────────────────────────────────
function doBackup() {
  const payload = {
    _type: 'warz-admin-backup',
    savedAt: new Date().toISOString(),
    events: state.events,
    codes: state.codes,
    items: state.items,
  };
  const stamp = new Date().toISOString().slice(0, 19).replace(/[:T]/g, '-');
  const a = Object.assign(document.createElement('a'), {
    href: URL.createObjectURL(new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' })),
    download: `warz-backup-${stamp}.json`,
  });
  a.click();
  toast('ดาวน์โหลด Backup แล้ว ✓');
}

function restoreBackup(file) {
  const reader = new FileReader();
  reader.onload = () => {
    let data;
    try { data = JSON.parse(reader.result); }
    catch { toast('ไฟล์ไม่ถูกต้อง'); return; }
    if (!data || (!Array.isArray(data.events) && !Array.isArray(data.items))) {
      toast('ไฟล์ Backup ไม่ถูกต้อง'); return;
    }
    if (!confirm('กู้คืนข้อมูลจาก Backup นี้?\nข้อมูลปัจจุบันจะถูกแทนที่ (กด ↶ ย้อนกลับได้)')) return;
    if (Array.isArray(data.events)) state.events = data.events;
    if (Array.isArray(data.codes))  state.codes  = data.codes;
    if (Array.isArray(data.items))  { state.items = data.items; rebuildItemMap(); persistItemDb(); }
    persistData();
    refreshActiveView();
    toast('กู้คืนข้อมูลแล้ว ✓');
  };
  reader.readAsText(file);
}

function updateSavedLabel() {
  const el = document.getElementById('last-saved-label');
  if (!el) return;
  const t = localStorage.getItem(SAVED_KEY);
  el.textContent = t ? `บันทึกล่าสุด: ${t}` : '';
}

// ── Item DB ───────────────────────────────────────────
async function loadItemDb() {
  // Try localStorage first (user may have edited items)
  const cached = localStorage.getItem(ITEMS_DB_KEY);
  if (cached) {
    try {
      const parsed = JSON.parse(cached);
      state.items = parsed.items ?? [];
      rebuildItemMap();
      return;
    } catch (_) {}
  }
  // Fall back to warz_data.json file
  try {
    const res = await fetch('../warz_data.json');
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const data = await res.json();
    state.items = data.items ?? [];
    rebuildItemMap();
  } catch (err) {
    console.warn('loadItemDb: could not load warz_data.json', err);
    state.items = [];
    state.itemMap = new Map();
  }
}

function rebuildItemMap() {
  state.itemMap = new Map(state.items.map(item => [String(item.id), item]));
}

async function reloadItemDbFromFile() {
  if (!confirm('โหลด Items ใหม่จากไฟล์ warz_data.json?\nการแก้ไข Items ที่ยังไม่ได้ Export จะถูกเขียนทับ')) return;
  localStorage.removeItem(ITEMS_DB_KEY);
  try {
    const res = await fetch('../warz_data.json?t=' + Date.now());
    const data = await res.json();
    state.items = data.items ?? [];
    rebuildItemMap();
    renderItems();
    toast(`โหลด ${state.items.length} items จากไฟล์แล้ว ✓`);
  } catch (err) {
    toast('โหลดไฟล์ไม่ได้');
  }
}

function persistItemDb() {
  localStorage.setItem(ITEMS_DB_KEY, JSON.stringify({ items: state.items }));
  toast('บันทึก Items DB แล้ว ✓');
}

// ── Views ─────────────────────────────────────────────
function showView(name) {
  document.querySelectorAll('.admin-view').forEach(v => v.classList.remove('is-active'));
  document.querySelectorAll('.admin-nav-btn').forEach(b => b.classList.remove('is-active'));
  const view = document.getElementById(`view-${name}`);
  const btn  = document.querySelector(`.admin-nav-btn[data-view="${name}"]`);
  if (view) view.classList.add('is-active');
  if (btn)  btn.classList.add('is-active');

  if (name === 'dashboard') renderDashboard();
  if (name === 'events')    renderEvents();
  if (name === 'codes')     renderCodes();
  if (name === 'items')     renderItems();
  if (name === 'export')    renderExport();
  if (name === 'settings')  renderSettings();
  if (name === 'help')      renderHelp();
}

// ── Revert to last published (reload shared data from cloud) ──
async function revertToPublished() {
  if (!confirm('ดึงข้อมูลที่เผยแพร่ล่าสุดจากเว็บกลับมา?\nการแก้ไขที่ค้างไว้ (ยังไม่ได้เผยแพร่) จะหายไป')) return;
  toast('กำลังดึงข้อมูลที่เผยแพร่ล่าสุด...');
  const ok = await loadSharedData();
  if (ok) {
    localStorage.setItem(DATA_KEY, JSON.stringify({ events: state.events, codes: state.codes }));
    refreshActiveView();
    toast('ดึงข้อมูลที่เผยแพร่ล่าสุดแล้ว ✓');
  } else {
    toast('ดึงข้อมูลไม่ได้ (ยังไม่มีข้อมูลที่เผยแพร่ หรือเชื่อมต่อไม่ได้)');
  }
}

// ── Help / วิธีใช้ ─────────────────────────────────────
function renderHelp() {
  const el = document.getElementById('help-content');
  if (!el) return;
  el.innerHTML = `
    <div class="help-section">
      <h3>🚀 ขั้นตอนใช้งานหลัก</h3>
      <ol>
        <li>เข้าหลังบ้าน → ระบบโหลด<strong>ข้อมูลล่าสุดจาก cloud</strong>ให้อัตโนมัติ</li>
        <li>แก้ไขข้อมูล (รายชื่อ/สถานะ/รางวัล/โค้ด) — ขั้นนี้เก็บในเครื่องก่อน <em>ยังไม่ขึ้นเว็บจริง</em></li>
        <li>กด <strong>"🚀 เผยแพร่ขึ้นเว็บเลย"</strong> (เมนู Export) → ทุกคน + ผู้เล่นเห็นทันที</li>
      </ol>
      <p class="help-note">⚠️ ถ้าปิดหน้าก่อนกดเผยแพร่ การแก้ไขจะหาย — กดเผยแพร่ทุกครั้งที่แก้เสร็จ</p>
    </div>

    <div class="help-section">
      <h3>📋 จัดการรายชื่อ (เมนู กิจกรรม → รายชื่อ)</h3>
      <ul>
        <li><strong>เปลี่ยนสถานะ/รางวัล:</strong> คลิก dropdown ในตารางได้เลย (เปลี่ยนทันที)</li>
        <li><strong>เลือกหลายคน:</strong> ติ๊ก checkbox หน้าแถว → ใช้แถบด้านบนเปลี่ยนสถานะ/รางวัล/ลบ ทีเดียวหลายคน</li>
        <li><strong>กรอง:</strong> ใช้ dropdown "ทุกสถานะ/ทุกรางวัล" เพื่อดูเฉพาะกลุ่ม</li>
        <li><strong>วางทีละเยอะ:</strong> ปุ่ม "📋 วางรายชื่อ" → ก๊อปจาก Excel มาวางได้เลย</li>
        <li><strong>⚠ CS:</strong> ปุ่มในคอลัมน์หมายเหตุ — คลิกเพื่อแจ้งให้ผู้เล่นติดต่อ CS ผ่าน Ticket (แถวจะแดง + บนเว็บกดเป็นลิงก์ได้)</li>
        <li><strong>ตรวจซ้ำ:</strong> ปุ่ม "🔍 ตรวจหาซ้ำ" หา UID/กิลด์ซ้ำในกิจกรรม</li>
      </ul>
    </div>

    <div class="help-section">
      <h3>🎁 จัดการรางวัล + ไอเทม</h3>
      <ul>
        <li>แก้กิจกรรม → ส่วน "รางวัลแยกตามหมวด" → เลือกหมวด (Luckydraw/ได้ทุกคน/อันดับ) แล้วเพิ่มไอเทม</li>
        <li><strong>วาง Item ID หลายตัว:</strong> ก๊อป ID จาก Excel ทั้งคอลัมน์ มาวางในช่องค้นหา → เพิ่มทีเดียว</li>
        <li>ถ้าหา item ไม่เจอ → เมนู Items DB → กด "↻ โหลดใหม่จากไฟล์"</li>
      </ul>
    </div>

    <div class="help-section">
      <h3>🖼️ รูป Cover + วันเวลา + ลิงก์</h3>
      <ul>
        <li>แก้กิจกรรม → ใส่รูป Cover (ระบบย่อให้พอดี), ตั้ง "หมดเขตกดรับ" (มี countdown บนเว็บ), ใส่ลิงก์โพสต์ Facebook</li>
        <li>กิจกรรมแบบ <strong>กิลด์</strong>: เปลี่ยน "ประเภทกิจกรรม" เป็นกิลด์ → ใส่ชื่อกิลด์แทน UID</li>
      </ul>
    </div>

    <div class="help-section">
      <h3>↺ พลาด/อยากย้อน</h3>
      <ul>
        <li>ปุ่ม <strong>"↶ ย้อนกลับ"</strong> (ซ้ายล่าง) — ย้อนการแก้ทีละขั้น</li>
        <li>เมนู Export → <strong>"↺ ดึงข้อมูลที่เผยแพร่ล่าสุด"</strong> — โหลดข้อมูลบนเว็บจริงกลับมา (ทิ้งที่แก้ค้างไว้)</li>
        <li>เมนู Export → <strong>Backup</strong> — ดาวน์โหลด/กู้คืนไฟล์สำรอง</li>
      </ul>
    </div>

    <div class="help-section">
      <h3>👥 ทำงานหลายคน</h3>
      <p>ทุกคนใช้ login + Publish Token เดียวกัน เปิดหลังบ้านจะเห็นข้อมูลชุดเดียวกัน — <strong>เลี่ยงแก้พร้อมกัน 2 คนเวลาเดียว</strong> (กันทับกัน) ใครเผยแพร่ทีหลังจะทับของก่อนหน้า</p>
    </div>
  `;
}

// ── Dashboard ─────────────────────────────────────────
function isDeliveredStatus(status) {
  const s = String(status || '').toLowerCase();
  return s.includes('จัดส่งแล้ว') || s.includes('รับรางวัลแล้ว');
}

function eventStats(ev) {
  const total = ev.winners.length;
  const done = ev.winners.filter(w => isDeliveredStatus(w.claimStatus)).length;
  const problem = ev.winners.filter(w => w.note).length;
  return { total, done, pending: total - done, problem, pct: total ? Math.round(done / total * 100) : 0 };
}

function renderDashboard() {
  updateSavedLabel();
  const totalWinners   = state.events.reduce((s, e) => s + e.winners.length, 0);
  const totalDone      = state.events.reduce((s, e) => s + e.winners.filter(w => isDeliveredStatus(w.claimStatus)).length, 0);
  const problemWinners = state.events.reduce((s, e) => s + e.winners.filter(w => w.note).length, 0);

  document.getElementById('stats-grid').innerHTML = `
    <div class="stat-card"><span class="stat-num">${state.events.length}</span><span class="stat-label">กิจกรรม</span></div>
    <div class="stat-card"><span class="stat-num">${totalWinners}</span><span class="stat-label">รายชื่อทั้งหมด</span></div>
    <div class="stat-card done"><span class="stat-num">${totalDone}</span><span class="stat-label">จัดส่งแล้ว</span></div>
    <div class="stat-card warn"><span class="stat-num">${problemWinners}</span><span class="stat-label">มีปัญหา</span></div>
    <div class="stat-card"><span class="stat-num">${state.codes.length}</span><span class="stat-label">Master Code</span></div>
  `;

  document.getElementById('dash-events').innerHTML = `
    <h3 class="dash-section-title">สรุปรายกิจกรรม</h3>
    <div class="dash-event-grid">
      ${state.events.map(ev => {
        const st = eventStats(ev);
        return `
        <div class="dash-event-card">
          <strong>${esc(ev.name)}</strong>
          <div class="dash-mini-bar"><div class="dash-mini-fill" style="width:${st.pct}%"></div></div>
          <div class="dash-counts">
            <span class="dc-done">✓ ${st.done}</span>
            <span class="dc-pending">รอ ${st.pending}</span>
            ${st.problem ? `<span class="dc-problem">⚠ ${st.problem}</span>` : ''}
            <span class="dc-total">/ ${st.total}</span>
          </div>
          <button class="btn-sm btn-ghost" onclick="openWinnersView('${ev.id}')">จัดการ →</button>
        </div>`;
      }).join('')}
    </div>
  `;
}

// ── Cross-event winner search (dashboard) ─────────────
function renderGlobalSearch() {
  const resultsEl = document.getElementById('global-search-results');
  if (!resultsEl) return;
  const q = (document.getElementById('global-search')?.value || '').toLowerCase().trim();
  if (!q) { resultsEl.innerHTML = ''; return; }

  const matches = [];
  for (const ev of state.events) {
    for (const w of ev.winners) {
      const hay = `${w.uid || ''} ${w.facebook || ''} ${w.guild || ''}`.toLowerCase();
      if (hay.includes(q)) matches.push({ ev, w });
      if (matches.length >= 40) break;
    }
    if (matches.length >= 40) break;
  }

  if (!matches.length) { resultsEl.innerHTML = '<p class="muted-label" style="padding:8px 0">ไม่พบ</p>'; return; }

  resultsEl.innerHTML = `
    <table class="admin-table">
      <thead><tr><th>กิจกรรม</th><th>${'UID / กิลด์'}</th><th>Facebook</th><th>รางวัล</th><th>สถานะ</th><th></th></tr></thead>
      <tbody>
        ${matches.map(({ ev, w }) => {
          const idx = ev.winners.indexOf(w);
          return `<tr class="${w.note ? 'row-problem' : ''}">
            <td>${esc(ev.shortName || ev.name)}</td>
            <td><code>${esc(w.guild || w.uid || '-')}</code></td>
            <td>${esc(w.facebook || '-')}</td>
            <td><span class="status-badge">${esc(winnerCategory(w))}</span></td>
            <td><span class="status-badge ${statusBadgeClass(w.claimStatus)}">${esc(normalizeStatus(w.claimStatus))}</span></td>
            <td><button class="btn-xs btn-secondary" onclick="jumpToWinner('${ev.id}',${idx})">แก้ไข →</button></td>
          </tr>`;
        }).join('')}
      </tbody>
    </table>
  `;
}

function jumpToWinner(eventId, idx) {
  openWinnersView(eventId);
  setTimeout(() => openEditWinnerModal(eventId, idx), 100);
}

// ── Events ────────────────────────────────────────────
function renderEvents() {
  const el = document.getElementById('events-list');
  if (!state.events.length) { el.innerHTML = '<p class="empty-msg">ยังไม่มีกิจกรรม</p>'; return; }

  el.innerHTML = state.events.map((ev, i) => {
    const guild = isGuildEvent(ev);
    const thumb = ev.coverImage
      ? `<img src="${esc(ev.coverImage)}" class="event-row-thumb" alt="" />`
      : `<div class="event-row-thumb event-row-noimg">${esc((ev.shortName || ev.name || '?')[0])}</div>`;
    const deadline = ev.claimDeadline ? formatDeadlineShort(ev.claimDeadline) : '';
    return `
    <div class="list-row">
      ${thumb}
      <div class="list-row-info">
        <strong>${esc(ev.name)} <span class="type-tag ${guild ? 'tag-guild' : ''}">${guild ? 'กิลด์' : 'ผู้เล่น'}</span></strong>
        <span>${ev.winners.length} ${guild ? 'กิลด์' : 'รายชื่อ'}${deadline ? ` · ⏰ ${deadline}` : ''}${ev.fbPostUrl ? ' · 🔗 มีลิงก์ FB' : ''}</span>
        <span class="status-badge">${esc(ev.status)}</span>
      </div>
      <div class="list-row-actions">
        <button class="btn-sm btn-ghost" onclick="openWinnersView('${ev.id}')">รายชื่อ</button>
        <button class="btn-sm btn-secondary" onclick="openEditEventModal(${i})">แก้ไข</button>
        <button class="btn-sm btn-danger"    onclick="deleteEvent(${i})">ลบ</button>
      </div>
    </div>`;
  }).join('');
}

// Combine the date + time inputs into a "YYYY-MM-DDTHH:mm" deadline string
function buildDeadline(f) {
  const date = (f.get('claimDate') || '').trim();
  if (!date) return '';
  const time = (f.get('claimTime') || '').trim() || '23:59';
  return `${date}T${time}`;
}

function formatDeadlineShort(dl) {
  const d = new Date(dl);
  if (isNaN(d)) return esc(dl);
  return d.toLocaleString('th-TH', { day: '2-digit', month: 'short', hour: '2-digit', minute: '2-digit' });
}

let _editingCover = '';

function eventFormHtml(ev) {
  const dl = ev?.claimDeadline || '';
  const dlDate = dl.includes('T') ? dl.split('T')[0] : '';
  const dlTime = dl.includes('T') ? dl.split('T')[1].slice(0, 5) : '';
  return `
    <form id="ef" class="admin-form">
      <div class="form-section-title">ข้อมูลกิจกรรม</div>
      <label class="field-label">ชื่อกิจกรรม <input type="text" name="name" value="${esc(ev?.name || '')}" required /></label>
      <div class="form-row-2">
        <label class="field-label">ประเภทกิจกรรม
          <select name="eventType" class="status-select">
            <option value="player" ${(ev?.eventType || 'player') === 'player' ? 'selected' : ''}>ผู้เล่น (UID / Facebook)</option>
            <option value="guild" ${ev?.eventType === 'guild' ? 'selected' : ''}>กิลด์ (ประกาศชื่อกิลด์)</option>
          </select>
        </label>
        <label class="field-label">สถานะ <input type="text" name="status" value="${esc(ev?.status || 'กำลังดำเนินการ')}" /></label>
      </div>
      <label class="field-label">ช่วงเวลากิจกรรม <input type="text" name="period" value="${esc(ev?.period || '')}" placeholder="DD/MM/YYYY - DD/MM/YYYY" /></label>
      <label class="field-label">⏰ หมดเขตกดรับรางวัล
        <div class="deadline-inputs">
          <input type="date" name="claimDate" value="${esc(dlDate)}" />
          <input type="time" name="claimTime" value="${esc(dlTime)}" />
        </div>
      </label>
      <label class="field-label">วันส่งของรางวัลล่าสุด <input type="text" name="latest" value="${esc(ev?.latest || '')}" /></label>
      <label class="field-label">🔗 ลิงก์โพสต์ Facebook กิจกรรม <input type="url" name="fbPostUrl" value="${esc(ev?.fbPostUrl || '')}" placeholder="https://www.facebook.com/..." /></label>
      <label class="field-label">ข้อมูลตัดรอบ <input type="text" name="resetDate" value="${esc(ev?.resetDate || 'ตัดรอบ/จัดส่ง: ทุกวันพุธ')}" /></label>

      <div class="form-section-title">รูป Cover (สำหรับหน้าเว็บ + โพสต์ Facebook)</div>
      <div class="cover-field">
        <div id="cover-preview" class="cover-preview"></div>
        <div class="cover-actions">
          <button type="button" class="btn-secondary btn-sm" id="cover-pick-btn">เลือกรูป (1200×1200)</button>
          <button type="button" class="btn-ghost btn-sm" id="cover-remove-btn">ลบรูป</button>
        </div>
        <input type="file" id="event-cover-input" accept="image/*" class="hidden" />
      </div>

      <div class="form-section-title">รางวัล</div>
      ${eventRewardSetsHtml()}
      <div class="form-actions">
        <button type="submit" class="btn-primary">บันทึก</button>
        <button type="button" class="btn-ghost" onclick="closeModal()">ยกเลิก</button>
      </div>
    </form>
  `;
}

// ── Event reward sets (grouped by category) ───────────
// _editingSets: [{ category, items: [{id,name,image,type,amount}] }]
let _editingSets = [];

function eventRewardSetsHtml() {
  return `
    <div class="reward-sets-editor">
      <label class="field-label" style="margin-bottom:6px">รางวัลแยกตามหมวด</label>
      <div id="reward-sets-list"></div>
      <div class="reward-set-add">
        <select id="reward-set-cat" class="status-select">
          ${REWARD_CATEGORIES.map(c => `<option value="${esc(c)}">${esc(c)}</option>`).join('')}
        </select>
        <input type="text" id="reward-set-search" class="search-input" placeholder="ค้นหา / วาง Item ID (ก๊อปจาก Excel หลายตัวได้)..." autocomplete="off" oninput="renderRewardSetResults()" onpaste="handleRewardSetPaste(event)" onkeydown="handleRewardSetKey(event)" />
      </div>
      <div id="reward-set-results" class="item-picker-results"></div>
    </div>
  `;
}

function deriveEventRewardSets(ev) {
  if (Array.isArray(ev.rewardSets) && ev.rewardSets.length) {
    return ev.rewardSets.map(s => ({
      category: s.category,
      items: (s.items || []).map(rewardToEditItem),
    }));
  }
  // Derive from winners grouped by each winner's reward category
  const byCat = new Map();
  for (const w of ev.winners) {
    const cat = winnerCategory(w);
    const items = (w.rewards || []).filter(r => r && r.hasItem && r.itemEn);
    if (!items.length) continue;
    if (!byCat.has(cat)) byCat.set(cat, new Map());
    const m = byCat.get(cat);
    for (const r of items) {
      const key = r.itemId || r.itemEn;
      if (!m.has(key)) m.set(key, rewardToEditItem(r));
    }
  }
  // Legacy flat pendingRewards with no category → put under "อื่นๆ"
  const flat = (ev.pendingRewards || []).filter(r => r && r.hasItem && r.itemEn);
  if (flat.length && !byCat.size) {
    const m = new Map();
    for (const r of flat) { const k = r.itemId || r.itemEn; if (!m.has(k)) m.set(k, rewardToEditItem(r)); }
    byCat.set('อื่นๆ', m);
  }
  return orderSets([...byCat.entries()].map(([category, m]) => ({ category, items: [...m.values()] })));
}

function orderSets(sets) {
  return sets.slice().sort((a, b) =>
    REWARD_CATEGORIES.indexOf(a.category) - REWARD_CATEGORIES.indexOf(b.category));
}

function renderRewardSets() {
  const el = document.getElementById('reward-sets-list');
  if (!el) return;
  _editingSets = orderSets(_editingSets);
  if (!_editingSets.length) {
    el.innerHTML = '<p class="muted-label" style="font-size:0.78rem;padding:4px 0 8px">ยังไม่มีรางวัล — เลือกหมวดแล้วค้นหา item ด้านล่าง</p>';
    return;
  }
  el.innerHTML = _editingSets.map((set, si) => `
    <div class="reward-set">
      <div class="reward-set-head">
        <span class="reward-set-cat">${esc(set.category)}</span>
        <span class="reward-set-count">${set.items.length} ไอเทม</span>
      </div>
      ${set.items.length
        ? set.items.map((item, ii) => `
          <div class="editing-item-row">
            ${item.image ? `<img src="${esc(item.image)}" class="editing-item-thumb" onerror="this.style.display='none'" />` : '<div class="editing-item-thumb editing-item-no-img">?</div>'}
            <div class="editing-item-info"><strong>${esc(item.name)}</strong><code>${esc(item.id)}</code></div>
            <label class="editing-item-amount">x<input type="number" min="1" value="${esc(item.amount || '1')}" oninput="setRewardSetAmount(${si},${ii},this.value)" /></label>
            <button type="button" class="btn-xs btn-danger" onclick="removeRewardSetItem(${si},${ii})">✕</button>
          </div>`).join('')
        : '<p class="muted-label" style="font-size:0.74rem;padding:2px 0">หมวดนี้ยังไม่มี item</p>'}
    </div>
  `).join('');
}

function renderRewardSetResults() {
  const input = document.getElementById('reward-set-search');
  const resultsEl = document.getElementById('reward-set-results');
  const cat = document.getElementById('reward-set-cat')?.value;
  if (!input || !resultsEl) return;
  const q = input.value.toLowerCase().trim();
  if (!q) { resultsEl.innerHTML = ''; return; }
  const set = _editingSets.find(s => s.category === cat);
  const existing = new Set(set ? set.items.map(i => String(i.id)) : []);
  const matches = state.items
    .filter(it => !existing.has(String(it.id)) &&
      ((it.name || '').toLowerCase().includes(q) || String(it.id).includes(q) || (it.type || '').toLowerCase().includes(q)))
    .slice(0, 8);
  if (!matches.length) { resultsEl.innerHTML = '<p class="muted-label" style="font-size:0.78rem;padding:6px 0">ไม่พบ Item</p>'; return; }
  resultsEl.innerHTML = matches.map(item => `
    <button type="button" class="item-picker-result" onclick="addRewardSetItem('${esc(String(item.id))}')">
      ${item.image ? `<img src="${esc(item.image)}" class="editing-item-thumb" onerror="this.style.display='none'" />` : '<div class="editing-item-thumb editing-item-no-img">?</div>'}
      <span>${esc(item.name)}</span>
      <code class="item-picker-id">${esc(item.id)}</code>
    </button>
  `).join('');
}

function addRewardSetItem(itemId) {
  const cat = document.getElementById('reward-set-cat')?.value || 'อื่นๆ';
  const item = state.itemMap.get(String(itemId));
  const editItem = item
    ? { id: String(item.id), name: item.name, image: item.image || '', type: item.type || '', amount: '1' }
    : { id: itemId, name: itemId, image: '', type: '', amount: '1' };
  let set = _editingSets.find(s => s.category === cat);
  if (!set) { set = { category: cat, items: [] }; _editingSets.push(set); }
  if (set.items.find(i => String(i.id) === String(itemId))) return;
  set.items.push(editItem);
  renderRewardSets();
  const search = document.getElementById('reward-set-search');
  const results = document.getElementById('reward-set-results');
  if (search) search.value = '';
  if (results) results.innerHTML = '';
}

// Add many Item IDs at once (e.g. pasted from an Excel column) into the selected category
function addRewardSetIds(ids) {
  const cat = document.getElementById('reward-set-cat')?.value || 'อื่นๆ';
  let set = _editingSets.find(s => s.category === cat);
  if (!set) { set = { category: cat, items: [] }; _editingSets.push(set); }
  let added = 0, dup = 0, notfound = 0;
  for (const raw of ids) {
    const key = String(raw).trim();
    if (!key) continue;
    if (set.items.find(i => String(i.id) === key)) { dup++; continue; }
    const item = state.itemMap.get(key);
    if (!item) { notfound++; continue; }
    set.items.push({ id: String(item.id), name: item.name, image: item.image || '', type: item.type || '', amount: '1' });
    added++;
  }
  renderRewardSets();
  const search = document.getElementById('reward-set-search');
  const results = document.getElementById('reward-set-results');
  if (search) search.value = '';
  if (results) results.innerHTML = '';
  toast(`เพิ่ม ${added} ไอเทมเข้า "${cat}"${dup ? ` · ซ้ำ ${dup}` : ''}${notfound ? ` · ไม่พบใน DB ${notfound}` : ''}`);
}

function handleRewardSetPaste(e) {
  const text = (e.clipboardData || window.clipboardData)?.getData('text') || '';
  const ids = text.split(/[\s,;\t\r\n]+/).map(s => s.trim()).filter(Boolean);
  if (!ids.length) return;
  // Intercept if it's a list, or any token matches a known Item ID
  const anyKnown = ids.some(id => state.itemMap.has(String(id)));
  if (ids.length > 1 || anyKnown) {
    e.preventDefault();
    addRewardSetIds(ids);
  }
}

function handleRewardSetKey(e) {
  if (e.key !== 'Enter') return;
  e.preventDefault();
  const val = (e.target.value || '').trim();
  if (!val) return;
  // Exact Item ID → add it; otherwise add the first search result if any
  if (state.itemMap.has(val)) {
    addRewardSetIds([val]);
  } else {
    const firstResult = document.querySelector('#reward-set-results .item-picker-result');
    if (firstResult) firstResult.click();
  }
}

function setRewardSetAmount(si, ii, value) {
  if (_editingSets[si] && _editingSets[si].items[ii]) _editingSets[si].items[ii].amount = String(value || '1');
}

function removeRewardSetItem(si, ii) {
  if (!_editingSets[si]) return;
  _editingSets[si].items.splice(ii, 1);
  if (!_editingSets[si].items.length) _editingSets.splice(si, 1); // drop empty set
  renderRewardSets();
}

// Convert editing sets → stored rewardSets + a flat pendingRewards (compat)
function editingSetsToStored() {
  const rewardSets = _editingSets
    .filter(s => s.items.length)
    .map(s => ({ category: s.category, items: s.items.map(editItemToReward) }));
  const pendingRewards = rewardSets.flatMap(s => s.items);
  return { rewardSets, pendingRewards };
}

// Resize an image file to fit within max×max (preserving aspect), return JPEG data URL
function resizeImage(file, max = 1200) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = reject;
    reader.onload = () => {
      const img = new Image();
      img.onerror = reject;
      img.onload = () => {
        let { width, height } = img;
        if (width > max || height > max) {
          const scale = Math.min(max / width, max / height);
          width = Math.round(width * scale);
          height = Math.round(height * scale);
        }
        const canvas = document.createElement('canvas');
        canvas.width = width;
        canvas.height = height;
        canvas.getContext('2d').drawImage(img, 0, 0, width, height);
        resolve(canvas.toDataURL('image/jpeg', 0.85));
      };
      img.src = reader.result;
    };
    reader.readAsDataURL(file);
  });
}

function renderCoverPreview() {
  const el = document.getElementById('cover-preview');
  if (!el) return;
  el.innerHTML = _editingCover
    ? `<img src="${_editingCover}" alt="cover" />`
    : `<span class="cover-empty">ยังไม่มีรูป Cover</span>`;
}

function wireCoverInput() {
  renderCoverPreview();
  const input = document.getElementById('event-cover-input');
  document.getElementById('cover-pick-btn')?.addEventListener('click', () => input?.click());
  document.getElementById('cover-remove-btn')?.addEventListener('click', () => { _editingCover = ''; renderCoverPreview(); });
  input?.addEventListener('change', async e => {
    const file = e.target.files[0];
    if (!file) return;
    toast('กำลังย่อรูป...');
    try {
      _editingCover = await resizeImage(file, 1200);
      renderCoverPreview();
      toast('ใส่รูป Cover แล้ว ✓');
    } catch { toast('อ่านรูปไม่ได้'); }
    e.target.value = '';
  });
}

function openAddEventModal() {
  _editingSets = [];
  _editingCover = '';
  showModal('เพิ่มกิจกรรมใหม่', eventFormHtml(null));
  renderRewardSets();
  wireCoverInput();
  document.getElementById('ef').onsubmit = e => {
    e.preventDefault();
    const f = new FormData(e.target);
    const num = state.events.length + 1;
    const { rewardSets, pendingRewards } = editingSetsToStored();
    state.events.push({
      id: `event-${Date.now()}`,
      name: f.get('name'),
      shortName: f.get('name').slice(0, 18),
      icon: 'broadcast',
      cycle: `กิจกรรมที่ ${num}`,
      period: f.get('period') || '-',
      resetDate: f.get('resetDate') || 'ตัดรอบ/จัดส่ง: ทุกวันพุธ',
      latest: f.get('latest') || 'รออัปเดต',
      status: f.get('status') || 'กำลังดำเนินการ',
      eventType: f.get('eventType') || 'player',
      claimDeadline: buildDeadline(f),
      fbPostUrl: (f.get('fbPostUrl') || '').trim(),
      owner: '', reward: 'ดูรางวัลในรายชื่อผู้ได้รับรางวัล',
      coverImage: _editingCover,
      winners: [], rewardSets, pendingRewards,
    });
    persistData(); closeModal(); renderEvents();
  };
}

// Aggregate the item rewards that the public site shows for an event:
// the event's own pendingRewards if set, else unique items from winners.
function eventDisplayItems(ev) {
  const source = (ev.pendingRewards && ev.pendingRewards.length)
    ? ev.pendingRewards
    : ev.winners.flatMap(w => w.rewards || []);
  const map = new Map();
  for (const r of source) {
    if (!r || typeof r === 'string') continue;
    if (!r.hasItem || !r.itemEn) continue;
    const key = r.itemId || r.itemEn;
    if (!map.has(key)) map.set(key, rewardToEditItem(r));
  }
  return [...map.values()];
}

function openEditEventModal(i) {
  const ev = state.events[i];
  _editingSets = deriveEventRewardSets(ev);
  _editingCover = ev.coverImage || '';
  showModal('แก้ไขกิจกรรม', eventFormHtml(ev));
  renderRewardSets();
  wireCoverInput();
  document.getElementById('ef').onsubmit = e => {
    e.preventDefault();
    const f = new FormData(e.target);
    const { rewardSets, pendingRewards } = editingSetsToStored();
    state.events[i] = {
      ...ev,
      name: f.get('name'), shortName: f.get('name').slice(0, 18),
      period: f.get('period'), latest: f.get('latest'),
      status: f.get('status'), resetDate: f.get('resetDate'),
      eventType: f.get('eventType') || 'player',
      claimDeadline: buildDeadline(f),
      fbPostUrl: (f.get('fbPostUrl') || '').trim(),
      coverImage: _editingCover,
      rewardSets, pendingRewards,
    };
    persistData(); closeModal(); renderEvents();
  };
}

function deleteEvent(i) {
  const ev = state.events[i];
  if (!confirm(`ลบกิจกรรม "${ev.name}" และรายชื่อทั้งหมด ${ev.winners.length} รายการ?\nไม่สามารถย้อนกลับได้`)) return;
  state.events.splice(i, 1);
  persistData(); renderEvents();
}

// ── Winners ───────────────────────────────────────────
let _selectedWinners = new Set();   // holds realIdx of selected rows

function openWinnersView(eventId) {
  state.currentEventId = eventId;
  _selectedWinners.clear();
  const ev = state.events.find(e => e.id === eventId);
  document.getElementById('winners-event-name').textContent = ev.name;
  document.querySelectorAll('.admin-view').forEach(v => v.classList.remove('is-active'));
  document.querySelectorAll('.admin-nav-btn').forEach(b => b.classList.remove('is-active'));
  document.getElementById('view-winners').classList.add('is-active');
  document.getElementById('winners-search').value = '';
  const guild = isGuildEvent(ev);
  document.getElementById('add-winner-btn').textContent = guild ? '+ เพิ่มกิลด์' : '+ เพิ่มรายชื่อ';
  document.getElementById('paste-winners-btn').textContent = guild ? '📋 วางกิลด์' : '📋 วางรายชื่อ';
  document.getElementById('winners-search').placeholder = guild ? 'ค้นหาชื่อกิลด์...' : 'ค้นหา UID / Facebook...';
  populateWinnerFilters(ev);
  renderWinners();
}

function isGuildEvent(ev) {
  return ev && ev.eventType === 'guild';
}

function getFilteredWinnerList() {
  const ev = state.events.find(e => e.id === state.currentEventId);
  if (!ev) return { ev: null, list: [] };
  const q = (document.getElementById('winners-search')?.value || '').toLowerCase().trim();
  const rewardF = document.getElementById('winners-reward-filter')?.value || 'all';
  const statusF = document.getElementById('winners-status-filter')?.value || 'all';

  const list = ev.winners.filter(w => {
    if (rewardF !== 'all' && winnerCategory(w) !== rewardF) return false;
    if (statusF !== 'all' && (w.claimStatus || 'กำลังดำเนินการ') !== statusF) return false;
    if (q) {
      const hay = `${w.uid || ''} ${w.facebook || ''} ${w.guild || ''}`.toLowerCase();
      if (!hay.includes(q)) return false;
    }
    return true;
  });
  return { ev, list };
}

// Populate the reward/status filter dropdowns from the event's winners
function populateWinnerFilters(ev) {
  const rewardSel = document.getElementById('winners-reward-filter');
  const statusSel = document.getElementById('winners-status-filter');
  if (rewardSel) {
    const cats = [...new Set(ev.winners.map(winnerCategory))]
      .sort((a, b) => REWARD_CATEGORIES.indexOf(a) - REWARD_CATEGORIES.indexOf(b));
    rewardSel.innerHTML = '<option value="all">ทุกรางวัล</option>' +
      cats.map(c => `<option value="${esc(c)}">${esc(c)}</option>`).join('');
    rewardSel.value = 'all';
  }
  if (statusSel) {
    const sts = [...new Set(ev.winners.map(w => w.claimStatus || 'กำลังดำเนินการ'))];
    statusSel.innerHTML = '<option value="all">ทุกสถานะ</option>' +
      sts.map(s => `<option value="${esc(s)}">${esc(s)}</option>`).join('');
    statusSel.value = 'all';
  }
}

function renderWinners() {
  const { ev, list } = getFilteredWinnerList();
  if (!ev) return;

  // Drop selections no longer in the current filtered view
  const visibleIdx = new Set(list.map(w => ev.winners.indexOf(w)));
  _selectedWinners.forEach(i => { if (!visibleIdx.has(i)) _selectedWinners.delete(i); });

  document.getElementById('winners-count-label').textContent =
    list.length !== ev.winners.length ? `${list.length} / ${ev.winners.length} รายการ` : `${ev.winners.length} รายการ`;

  const wrap = document.getElementById('winners-table-wrap');
  if (!list.length) { wrap.innerHTML = '<p class="empty-msg">ไม่พบรายชื่อ</p>'; renderBulkBar(); return; }

  const allChecked = list.every(w => _selectedWinners.has(ev.winners.indexOf(w)));
  const guild = isGuildEvent(ev);

  wrap.innerHTML = `
    <table class="admin-table">
      <thead>
        <tr>
          <th class="check-cell"><input type="checkbox" id="winners-check-all" ${allChecked ? 'checked' : ''} onchange="toggleSelectAllWinners(this.checked)" /></th>
          <th>#</th>${guild ? '<th>กิลด์</th>' : '<th>UID</th><th>Facebook</th>'}<th>รางวัล</th><th>สถานะ</th><th>หมายเหตุ</th><th>จัดการ</th>
        </tr>
      </thead>
      <tbody>
        ${list.map((w, displayIdx) => {
          const realIdx = ev.winners.indexOf(w);
          const checked = _selectedWinners.has(realIdx);
          const idCols = guild
            ? `<td><strong>${esc(w.guild || '-')}</strong></td>`
            : `<td><code>${esc(w.uid || '-')}</code></td><td>${esc(w.facebook)}</td>`;
          return `
            <tr class="${w.note ? 'row-problem' : ''} ${checked ? 'row-selected' : ''}">
              <td class="check-cell"><input type="checkbox" ${checked ? 'checked' : ''} onchange="toggleWinnerSelect(${realIdx}, this.checked)" /></td>
              <td class="num-cell">${displayIdx + 1}</td>
              ${idCols}
              <td>${inlineRewardHtml(realIdx, winnerCategory(w))}</td>
              <td>${inlineStatusHtml(realIdx, w.claimStatus)}</td>
              <td class="note-cell">
                ${w.note ? `<span class="problem-note">${esc(w.note)}</span>` : ''}
                <button type="button" class="cs-toggle ${w.note === CS_NOTE ? 'is-on' : ''}" onclick="toggleWinnerCsNote(${realIdx})" title="ติ๊กเพื่อแจ้งให้ติดต่อ CS ผ่าน Ticket">⚠ CS</button>
              </td>
              <td class="action-cell">
                <button class="btn-xs btn-secondary" onclick="openEditWinnerModal('${ev.id}',${realIdx})">แก้ไข</button>
                <button class="btn-xs btn-danger"    onclick="deleteWinner('${ev.id}',${realIdx})">ลบ</button>
              </td>
            </tr>
          `;
        }).join('')}
      </tbody>
    </table>
  `;
  renderBulkBar();
}

// ── Bulk actions ──────────────────────────────────────
function renderBulkBar() {
  const bar = document.getElementById('winners-bulk-bar');
  if (!bar) return;
  const n = _selectedWinners.size;
  if (!n) { bar.classList.add('hidden'); bar.innerHTML = ''; return; }

  bar.classList.remove('hidden');
  bar.innerHTML = `
    <span class="bulk-count">เลือก <strong>${n}</strong> รายการ</span>
    <div class="bulk-controls">
      <select id="bulk-status-select" class="status-select">
        ${WINNER_STATUSES.map(s => `<option value="${esc(s)}">${esc(s)}</option>`).join('')}
      </select>
      <button class="btn-sm btn-primary" onclick="applyBulkStatus()">เปลี่ยนสถานะ</button>
      <select id="bulk-reward-select" class="status-select">
        ${REWARD_CATEGORIES.map(c => `<option value="${esc(c)}">${esc(c)}</option>`).join('')}
      </select>
      <button class="btn-sm btn-primary" onclick="applyBulkReward()">เปลี่ยนรางวัล</button>
      <button class="btn-sm btn-danger" onclick="bulkDeleteWinners()">ลบที่เลือก</button>
      <button class="btn-sm btn-ghost" onclick="clearWinnerSelection()">ยกเลิกการเลือก</button>
    </div>
  `;
}

function toggleWinnerSelect(realIdx, checked) {
  if (checked) _selectedWinners.add(realIdx);
  else _selectedWinners.delete(realIdx);
  // Update the "select all" checkbox + row highlight without a full re-render
  const { ev, list } = getFilteredWinnerList();
  const allEl = document.getElementById('winners-check-all');
  if (allEl) allEl.checked = list.length > 0 && list.every(w => _selectedWinners.has(ev.winners.indexOf(w)));
  renderBulkBar();
}

function toggleSelectAllWinners(checked) {
  const { ev, list } = getFilteredWinnerList();
  list.forEach(w => {
    const idx = ev.winners.indexOf(w);
    if (checked) _selectedWinners.add(idx);
    else _selectedWinners.delete(idx);
  });
  renderWinners();
}

function clearWinnerSelection() {
  _selectedWinners.clear();
  renderWinners();
}

function applyBulkStatus() {
  const { ev } = getFilteredWinnerList();
  if (!ev || !_selectedWinners.size) return;
  const status = document.getElementById('bulk-status-select')?.value;
  if (!status) return;
  if (!confirm(`เปลี่ยนสถานะของ ${_selectedWinners.size} รายการเป็น "${status}"?`)) return;
  _selectedWinners.forEach(i => { if (ev.winners[i]) ev.winners[i].claimStatus = status; });
  persistData();
  _selectedWinners.clear();
  renderWinners();
  toast(`เปลี่ยนสถานะแล้ว ✓`);
}

function bulkDeleteWinners() {
  const { ev } = getFilteredWinnerList();
  if (!ev || !_selectedWinners.size) return;
  if (!confirm(`ลบ ${_selectedWinners.size} รายการที่เลือก?\nไม่สามารถย้อนกลับได้`)) return;
  // Remove from highest index down so indices stay valid
  [..._selectedWinners].sort((a, b) => b - a).forEach(i => ev.winners.splice(i, 1));
  _selectedWinners.clear();
  persistData();
  renderWinners();
  toast('ลบที่เลือกแล้ว ✓');
}

// ── Paste import (winners) ────────────────────────────
function openPasteWinnersModal() {
  const ev = state.events.find(e => e.id === state.currentEventId);
  if (!ev) return;
  const guild = isGuildEvent(ev);
  const cols = guild
    ? '<strong>ชื่อกิลด์ · รางวัล · สถานะ · หมายเหตุ</strong>'
    : '<strong>UID · Facebook · รางวัล · สถานะ · หมายเหตุ</strong>';
  const placeholder = guild
    ? "Guild Alpha&#9;อันดับ 1&#9;จัดส่งแล้ว\nGuild Bravo&#9;อันดับ 2&#9;กำลังดำเนินการ"
    : "024VHD&#9;Sarada Ken&#9;Luckydraw&#9;จัดส่งแล้ว\nHFL131&#9;Jui Na Ja&#9;ได้ทุกคน&#9;กำลังดำเนินการ";
  showModal(`${guild ? 'วางรายชื่อกิลด์' : 'วางรายชื่อ'} — ${ev.name}`, `
    <p class="muted-label" style="margin-bottom:8px">
      ก๊อปจาก Excel / Google Sheets แล้ววางด้านล่าง — แต่ละแถว 1 ${guild ? 'กิลด์' : 'คน'} เรียงคอลัมน์:
      ${cols} (คั่นด้วย Tab หรือ , ก็ได้ / เว้นว่างได้)
    </p>
    <textarea id="paste-box" class="paste-box" rows="10" placeholder="${placeholder}"></textarea>
    <label class="field-label" style="flex-direction:row;align-items:center;gap:8px;margin-top:10px">
      <input type="checkbox" id="paste-replace" style="width:auto" /> แทนที่รายการเดิมทั้งหมด (ไม่ติ๊ก = เพิ่มต่อท้าย)
    </label>
    <div class="form-actions">
      <button class="btn-primary" onclick="applyPasteWinners()">เพิ่มรายการ</button>
      <button class="btn-ghost" onclick="closeModal()">ยกเลิก</button>
    </div>
  `);
  setTimeout(() => document.getElementById('paste-box')?.focus(), 50);
}

function parsePastedWinners(text, workStatus, guild = false) {
  const lines = text.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
  const winners = [];
  for (const line of lines) {
    const cols = (line.includes('\t') ? line.split('\t') : line.split(',')).map(c => c.trim());
    if (guild) {
      const [name, reward, status, note] = cols;
      if (!name) continue;
      if (name.toLowerCase() === 'guild' || name === 'ชื่อกิลด์') continue;
      winners.push({
        guild: name, uid: '', facebook: '-', character: name,
        claimMethod: '', claimStatus: status || workStatus || 'กำลังดำเนินการ',
        updatedAt: '', note: note || '',
        rewardCategory: normalizeRewardCategory(reward),
        rewards: [],
      });
    } else {
      const [uid, facebook, reward, status, note] = cols;
      if (!uid && !facebook) continue;
      if ((uid || '').toLowerCase() === 'uid') continue;
      winners.push({
        uid: uid || '', facebook: facebook || '-', character: uid || '-',
        claimMethod: '', claimStatus: status || workStatus || 'กำลังดำเนินการ',
        updatedAt: '', note: note || '',
        rewardCategory: normalizeRewardCategory(reward),
        rewards: [],
      });
    }
  }
  return winners;
}

function applyPasteWinners() {
  const ev = state.events.find(e => e.id === state.currentEventId);
  if (!ev) return;
  const text = document.getElementById('paste-box')?.value || '';
  const replace = document.getElementById('paste-replace')?.checked;
  const parsed = parsePastedWinners(text, ev.status, isGuildEvent(ev));
  if (!parsed.length) { toast('ไม่พบรายการในข้อความ'); return; }
  ev.winners = replace ? parsed : ev.winners.concat(parsed);
  ev.winners.sort((a, b) => (a.note ? 1 : 0) - (b.note ? 1 : 0));
  const dupes = findDuplicates(true); // auto-check, silent (we toast below)
  persistData();
  closeModal();
  renderWinners();
  toast(dupes
    ? `เพิ่ม ${parsed.length} รายการ — พบซ้ำ ${dupes} รายการ (ใส่หมายเหตุแล้ว) ⚠`
    : `เพิ่ม ${parsed.length} รายการแล้ว ✓`);
}

function applyBulkReward() {
  const { ev } = getFilteredWinnerList();
  if (!ev || !_selectedWinners.size) return;
  const category = document.getElementById('bulk-reward-select')?.value;
  if (!category) return;
  if (!confirm(`เปลี่ยนรางวัลของ ${_selectedWinners.size} รายการเป็น "${category}"?`)) return;
  _selectedWinners.forEach(i => { if (ev.winners[i]) ev.winners[i].rewardCategory = category; });
  persistData();
  _selectedWinners.clear();
  renderWinners();
  toast('เปลี่ยนรางวัลที่เลือกแล้ว ✓');
}

function statusBadgeClass(status) {
  const s = String(status || '').toLowerCase();
  if (s.includes('จัดส่งแล้ว') || s.includes('รับรางวัลแล้ว')) return 'badge-done';
  if (s.includes('ดำเนินการ') || s.includes('รอกดรับ') || s.includes('ติดต่อ')) return 'badge-pending';
  if (s.includes('หมดเขต') || s.includes('ใช้แล้ว')) return 'badge-expired';
  return '';
}

function inlineStatusHtml(realIdx, current) {
  const cur = normalizeStatus(current);
  const known = WINNER_STATUSES.includes(cur);
  return `
    <select class="inline-status ${statusBadgeClass(cur)}" onchange="setWinnerStatus(${realIdx}, this.value, this)" onclick="event.stopPropagation()">
      ${WINNER_STATUSES.map(s => `<option value="${esc(s)}" ${s === cur ? 'selected' : ''}>${esc(s)}</option>`).join('')}
      ${known ? '' : `<option value="${esc(cur)}" selected>${esc(cur)}</option>`}
    </select>
  `;
}

function setWinnerStatus(realIdx, value, el) {
  const ev = state.events.find(e => e.id === state.currentEventId);
  if (!ev || !ev.winners[realIdx]) return;
  ev.winners[realIdx].claimStatus = value;
  // Update the badge colour in place (avoid full re-render so editing many rows stays smooth)
  if (el) el.className = `inline-status ${statusBadgeClass(value)}`;
  persistData(true);
}

// ── Reward categories ─────────────────────────────────
const REWARD_CATEGORIES = [
  'Luckydraw', 'ได้ทุกคน', 'ถูกใจทีมงาน',
  'อันดับ 1', 'อันดับ 2', 'อันดับ 3', 'อันดับที่ 4 - 10', 'อื่นๆ',
];

function normalizeRewardCategory(str) {
  const s = String(str || '').toLowerCase();
  if (!s) return 'อื่นๆ';
  if (REWARD_CATEGORIES.includes(str)) return str;
  if (s.includes('lucky')) return 'Luckydraw';
  if (s.includes('ทุกคน')) return 'ได้ทุกคน';
  if (s.includes('ถูกใจ')) return 'ถูกใจทีมงาน';
  if (s.includes('อันดับ')) {
    if (/(4|5|6|7|8|9|10|4\s*-\s*10)/.test(s)) return 'อันดับที่ 4 - 10';
    if (s.includes('1')) return 'อันดับ 1';
    if (s.includes('2')) return 'อันดับ 2';
    if (s.includes('3')) return 'อันดับ 3';
  }
  return 'อื่นๆ';
}

function winnerCategory(w) {
  if (w.rewardCategory) return w.rewardCategory;
  const fr = (w.rewards || []).map(r => (typeof r === 'string' ? r : r.forumReward)).find(Boolean);
  return normalizeRewardCategory(fr);
}

function inlineRewardHtml(realIdx, current) {
  const cur = REWARD_CATEGORIES.includes(current) ? current : 'อื่นๆ';
  return `
    <select class="inline-reward" onchange="setWinnerReward(${realIdx}, this.value)" onclick="event.stopPropagation()">
      ${REWARD_CATEGORIES.map(c => `<option value="${esc(c)}" ${c === cur ? 'selected' : ''}>${esc(c)}</option>`).join('')}
    </select>
  `;
}

function setWinnerReward(realIdx, value) {
  const ev = state.events.find(e => e.id === state.currentEventId);
  if (!ev || !ev.winners[realIdx]) return;
  ev.winners[realIdx].rewardCategory = value;
  persistData(true);
}

// One-click toggle: flag a winner as "contact CS via Ticket" (sets/clears the note)
function toggleWinnerCsNote(realIdx) {
  const ev = state.events.find(e => e.id === state.currentEventId);
  if (!ev || !ev.winners[realIdx]) return;
  const w = ev.winners[realIdx];
  w.note = (w.note === CS_NOTE) ? '' : CS_NOTE;
  persistData(true);
  renderWinners();
}

const WINNER_STATUSES = ['กำลังดำเนินการ', 'จัดส่งแล้ว', 'รับรางวัลแล้ว', 'รอกดรับ', 'ติดต่อแก้ไขข้อมูล', 'หมดเขต'];

// Map any status value (with emoji prefixes / spelling variants) to a clean standard status
function normalizeStatus(status) {
  let s = String(status || '').replace(/[\u{1F000}-\u{1FAFF}\u{2600}-\u{27BF}\u{2B00}-\u{2BFF}\u{FE0F}]/gu, '').trim();
  if (!s) return 'กำลังดำเนินการ';
  if (WINNER_STATUSES.includes(s)) return s;
  if (s.includes('ติดต่อ'))     return 'ติดต่อแก้ไขข้อมูล';
  if (s.includes('จัดส่ง'))     return 'จัดส่งแล้ว';
  if (s.includes('รับรางวัล'))  return 'รับรางวัลแล้ว';
  if (s.includes('รอกด'))       return 'รอกดรับ';
  if (s.includes('หมดเขต') || s.includes('ใช้แล้ว')) return 'หมดเขต';
  if (s.includes('ดำเนินการ'))  return 'กำลังดำเนินการ';
  return s; // truly custom value → keep as-is
}

function statusSelectHtml(name, current, options) {
  let cur = current || options[0];
  if (options === WINNER_STATUSES) cur = normalizeStatus(cur);
  else if (typeof REWARD_CATEGORIES !== 'undefined' && options === REWARD_CATEGORIES) cur = normalizeRewardCategory(cur);
  const known = options.includes(cur);
  return `
    <select name="${name}" class="status-select">
      ${options.map(o => `<option value="${esc(o)}" ${o === cur ? 'selected' : ''}>${esc(o)}</option>`).join('')}
      <option value="${esc(cur)}" ${!known ? 'selected' : ''}>${known ? 'อื่นๆ (กำหนดเอง)' : esc(cur) + ' (กำหนดเอง)'}</option>
    </select>
  `;
}

const CS_NOTE = 'ข้อมูลไม่ถูกต้อง กรุณาติดต่อ CS ผ่าน Ticket';
const NOTE_PRESETS = [CS_NOTE, 'ที่อยู่ไม่ครบ', 'รอติดต่อกลับ', 'ข้อมูลซ้ำ', 'กรอกข้อมูลผิด'];

function toggleCsNote(checked) {
  const noteInput = document.getElementById('wf-note');
  if (!noteInput) return;
  if (checked) {
    noteInput.value = CS_NOTE;
  } else if (noteInput.value === CS_NOTE) {
    noteInput.value = '';
  }
}

// Quick-fill the note field from a preset chip
function setNoteValue(val) {
  const noteInput = document.getElementById('wf-note');
  if (noteInput) noteInput.value = val;
  const cs = document.getElementById('wf-cs-flag');
  if (cs) cs.checked = (val === CS_NOTE);
}

function winnerFormHtml(w, guild) {
  const idFields = guild
    ? `<label class="field-label">ชื่อกิลด์ <input type="text" name="guild" value="${esc(w?.guild || '')}" required /></label>`
    : `<div class="form-row-2">
        <label class="field-label">UID <input type="text" name="uid" value="${esc(w?.uid || '')}" /></label>
        <label class="field-label">Facebook <input type="text" name="facebook" value="${esc(w?.facebook || '')}" /></label>
      </div>`;
  return `
    <form id="wf" class="admin-form">
      ${idFields}
      <div class="form-row-2">
        <label class="field-label">รางวัล ${statusSelectHtml('reward', winnerCategory(w || {}), REWARD_CATEGORIES)}</label>
        <label class="field-label">สถานะ ${statusSelectHtml('status', w?.claimStatus || 'กำลังดำเนินการ', WINNER_STATUSES)}</label>
      </div>
      ${guild ? '' : `<label class="field-label">วิธีรับรางวัล <input type="text" name="method" value="${esc(w?.claimMethod || '')}" /></label>`}
      <label class="field-label">วันที่อัปเดต <input type="text" name="updatedAt" value="${esc(w?.updatedAt || '')}" /></label>
      <label class="field-label">
        หมายเหตุ / ปัญหา
        <input type="text" name="note" id="wf-note" value="${esc(w?.note || '')}" placeholder="ถ้ามีปัญหาใส่ที่นี่ (แถวจะแดง)" />
      </label>
      <label class="field-label" style="flex-direction:row;align-items:center;gap:8px;margin-top:-4px">
        <input type="checkbox" id="wf-cs-flag" style="width:auto" ${(w?.note || '') === CS_NOTE ? 'checked' : ''} onchange="toggleCsNote(this.checked)" />
        ⚠️ ข้อมูลไม่ถูกต้อง — แจ้งให้ติดต่อ CS ผ่าน Ticket
      </label>
      <div class="note-presets">
        <span class="note-presets-label">หมายเหตุด่วน:</span>
        ${NOTE_PRESETS.map(n => `<button type="button" class="note-preset-chip" onclick="setNoteValue('${esc(n)}')">${esc(n)}</button>`).join('')}
        <button type="button" class="note-preset-chip note-clear" onclick="setNoteValue('')">ล้าง</button>
      </div>
      ${guild ? '' : itemPickerHtml('winner-items', 'ไอเทมที่ได้รับ (รูป + จำนวน)')}
      <div class="form-actions">
        <button type="submit" class="btn-primary">บันทึก</button>
        <button type="button" class="btn-ghost" onclick="closeModal()">ยกเลิก</button>
      </div>
    </form>
  `;
}

function buildWinner(fd, existing = {}, guild = false) {
  const category = fd.get('reward') || 'อื่นๆ';
  const rewards = guild ? (existing.rewards || []) : _editingItems.map(editItemToReward);
  if (guild) {
    return {
      guild:         fd.get('guild')     || existing.guild        || '',
      uid: '', facebook: '-', character: fd.get('guild') || existing.guild || '-',
      claimMethod:   existing.claimMethod || '',
      claimStatus:   fd.get('status')    || existing.claimStatus  || 'กำลังดำเนินการ',
      updatedAt:     fd.get('updatedAt') || existing.updatedAt    || '',
      note:          fd.get('note')      || '',
      rewardCategory: category,
      rewards,
    };
  }
  return {
    uid:           fd.get('uid')       || existing.uid          || '',
    facebook:      fd.get('facebook')  || existing.facebook     || '-',
    character:     fd.get('uid')       || existing.character    || '-',
    claimMethod:   fd.get('method')    || existing.claimMethod  || '',
    claimStatus:   fd.get('status')    || existing.claimStatus  || 'กำลังดำเนินการ',
    updatedAt:     fd.get('updatedAt') || existing.updatedAt    || '',
    note:          fd.get('note')      || '',
    rewardCategory: category,
    rewards,
  };
}

function openAddWinnerModal() {
  const ev = state.events.find(e => e.id === state.currentEventId);
  const guild = isGuildEvent(ev);
  _editingItems = [];
  showModal(`${guild ? 'เพิ่มกิลด์' : 'เพิ่มรายชื่อ'} — ${ev.name}`, winnerFormHtml(null, guild));
  if (!guild) renderEditingItemsList('winner-items');
  document.getElementById('wf').onsubmit = e => {
    e.preventDefault();
    ev.winners.push(buildWinner(new FormData(e.target), {}, guild));
    ev.winners.sort((a, b) => (a.note ? 1 : 0) - (b.note ? 1 : 0));
    persistData(); closeModal(); renderWinners();
  };
}

function openEditWinnerModal(eventId, idx) {
  const ev = state.events.find(e => e.id === eventId);
  const w  = ev.winners[idx];
  const guild = isGuildEvent(ev);
  _editingItems = guild ? [] : (w.rewards || []).filter(r => r && r.hasItem).map(rewardToEditItem);
  showModal(guild ? 'แก้ไขกิลด์' : 'แก้ไขรายชื่อ', winnerFormHtml(w, guild));
  if (!guild) renderEditingItemsList('winner-items');
  document.getElementById('wf').onsubmit = e => {
    e.preventDefault();
    ev.winners[idx] = buildWinner(new FormData(e.target), w, guild);
    ev.winners.sort((a, b) => (a.note ? 1 : 0) - (b.note ? 1 : 0));
    persistData(); closeModal(); renderWinners();
  };
}

function deleteWinner(eventId, idx) {
  const ev = state.events.find(e => e.id === eventId);
  const w  = ev.winners[idx];
  if (!confirm(`ลบ "${w.guild || w.uid || w.facebook}" ออกจากรายชื่อ?`)) return;
  ev.winners.splice(idx, 1);
  persistData(); renderWinners();
}

// ── Duplicate detection ───────────────────────────────
function stripNoteToken(note, token) {
  if (!note) return '';
  return note.split('/').map(s => s.trim()).filter(s => s && s !== token).join(' / ');
}

// Scan the current event for duplicate UID (or guild name) and flag them with a note.
// Returns the number of flagged rows. If `silent`, no toast/persist (caller handles it).
function findDuplicates(silent = false) {
  const ev = state.events.find(e => e.id === state.currentEventId);
  if (!ev) return 0;
  const guild = isGuildEvent(ev);
  const token = guild ? 'ชื่อกิลด์ซ้ำ' : 'UID ซ้ำ';
  const keyOf = w => (guild ? (w.guild || '') : (w.uid || '')).trim().toLowerCase();

  // Count occurrences of each non-blank key
  const counts = {};
  ev.winners.forEach(w => { const k = keyOf(w); if (k) counts[k] = (counts[k] || 0) + 1; });

  // Re-apply: clean old token everywhere, then mark current duplicates
  let flagged = 0;
  ev.winners.forEach(w => {
    w.note = stripNoteToken(w.note, token);
    const k = keyOf(w);
    if (k && counts[k] > 1) {
      w.note = w.note ? `${w.note} / ${token}` : token;
      flagged++;
    }
  });

  ev.winners.sort((a, b) => (a.note ? 1 : 0) - (b.note ? 1 : 0));

  if (!silent) {
    persistData();
    renderWinners();
    toast(flagged ? `พบซ้ำ ${flagged} รายการ — ใส่หมายเหตุแล้ว ⚠` : 'ไม่พบรายการซ้ำ ✓');
  }
  return flagged;
}

// ── Item Picker (shared for events, codes & winners) ──
// Internal editing shape: { id, name, image, type, amount }
let _editingItems = [];

// ---- Converters between editing shape and stored shapes ----

// Public site reward object (used by events.pendingRewards & winners.rewards)
function editItemToReward(it) {
  const display = it.name || it.id || 'รางวัลกิจกรรม';
  return {
    name:        display,
    forumReward: it.forumReward || display,
    itemEn:      display,
    itemId:      String(it.id ?? ''),
    itemType:    it.type || '',
    ownPeriod:   it.ownPeriod || '',
    amount:      String(it.amount || '1'),
    imageUrl:    it.image || '',
    hasItem:     true,
  };
}
function rewardToEditItem(r) {
  if (typeof r === 'string') return { id: r, name: r, image: '', type: '', amount: '1' };
  return {
    id:     String(r.itemId || r.itemEn || r.name || ''),
    name:   r.itemEn || r.name || r.forumReward || '',
    image:  r.imageUrl || '',
    type:   r.itemType || '',
    amount: String(r.amount || '1'),
    forumReward: r.forumReward || '',
  };
}

// Master-code item object (used by codes.items — renderCodes reads itemEn/imageUrl/amount)
function editItemToCodeItem(it) {
  const display = it.name || it.id || '';
  return {
    itemEn:   display,
    itemId:   String(it.id ?? ''),
    itemType: it.type || '',
    amount:   String(it.amount || '1'),
    imageUrl: it.image || '',
  };
}
function codeItemToEditItem(ci) {
  if (typeof ci === 'string') return { id: ci, name: ci, image: '', type: '', amount: '1' };
  return {
    id:     String(ci.itemId || ci.itemEn || ''),
    name:   ci.itemEn || ci.name || '',
    image:  ci.imageUrl || '',
    type:   ci.itemType || '',
    amount: String(ci.amount || '1'),
  };
}

function itemPickerHtml(listId, label) {
  return `
    <div class="item-picker-wrap">
      <label class="field-label" style="margin-bottom:6px">${esc(label || 'Items / รางวัล')}</label>
      <div id="${listId}-list" class="editing-items-list"></div>
      <div class="item-picker-row">
        <input type="text" id="${listId}-search" class="search-input" placeholder="ค้นหา / วาง Item ID (ก๊อปจาก Excel หลายตัวได้)..." autocomplete="off" oninput="renderItemPickerResults('${listId}')" onpaste="handleItemPickerPaste(event,'${listId}')" onkeydown="handleItemPickerKey(event,'${listId}')" />
      </div>
      <div id="${listId}-results" class="item-picker-results"></div>
    </div>
  `;
}

function renderEditingItemsList(listId) {
  const el = document.getElementById(`${listId}-list`);
  if (!el) return;
  if (!_editingItems.length) {
    el.innerHTML = '<p class="muted-label" style="font-size:0.78rem;padding:4px 0 8px">ยังไม่มี Item</p>';
    return;
  }
  el.innerHTML = _editingItems.map((item, i) => `
    <div class="editing-item-row">
      ${item.image ? `<img src="${esc(item.image)}" class="editing-item-thumb" onerror="this.style.display='none'" />` : '<div class="editing-item-thumb editing-item-no-img">?</div>'}
      <div class="editing-item-info">
        <strong>${esc(item.name)}</strong>
        <code>${esc(item.id)}</code>
      </div>
      <label class="editing-item-amount">x<input type="number" min="1" value="${esc(item.amount || '1')}" oninput="setEditingItemAmount(${i}, this.value)" /></label>
      <button type="button" class="btn-xs btn-danger" onclick="removeEditingItem(${i},'${listId}')">✕</button>
    </div>
  `).join('');
}

function renderItemPickerResults(listId) {
  const input = document.getElementById(`${listId}-search`);
  const resultsEl = document.getElementById(`${listId}-results`);
  if (!input || !resultsEl) return;
  const q = input.value.toLowerCase().trim();
  if (!q) { resultsEl.innerHTML = ''; return; }

  const existing = new Set(_editingItems.map(i => String(i.id)));
  const matches  = state.items
    .filter(it => !existing.has(String(it.id)) &&
      ((it.name||'').toLowerCase().includes(q) || String(it.id).includes(q) || (it.type||'').toLowerCase().includes(q)))
    .slice(0, 8);

  if (!matches.length) {
    resultsEl.innerHTML = '<p class="muted-label" style="font-size:0.78rem;padding:6px 0">ไม่พบ Item ใน Items DB</p>';
    return;
  }
  resultsEl.innerHTML = matches.map(item => `
    <button type="button" class="item-picker-result" onclick="addEditingItem('${esc(String(item.id))}','${listId}')">
      ${item.image ? `<img src="${esc(item.image)}" class="editing-item-thumb" onerror="this.style.display='none'" />` : '<div class="editing-item-thumb editing-item-no-img">?</div>'}
      <span>${esc(item.name)}</span>
      <code class="item-picker-id">${esc(item.id)}</code>
    </button>
  `).join('');
}

function addEditingItem(itemId, listId) {
  if (_editingItems.find(i => String(i.id) === String(itemId))) return;
  const item = state.itemMap.get(String(itemId));
  if (!item) {
    _editingItems.push({ id: itemId, name: itemId, image: '', type: '', amount: '1' });
  } else {
    _editingItems.push({ id: String(item.id), name: item.name, image: item.image || '', type: item.type || '', amount: '1' });
  }
  renderEditingItemsList(listId);
  const search = document.getElementById(`${listId}-search`);
  const results = document.getElementById(`${listId}-results`);
  if (search)  search.value = '';
  if (results) results.innerHTML = '';
}

// Add many Item IDs at once (pasted from an Excel column) to the shared picker
function addEditingItemIds(ids, listId) {
  let added = 0, dup = 0, notfound = 0;
  for (const raw of ids) {
    const key = String(raw).trim();
    if (!key) continue;
    if (_editingItems.find(i => String(i.id) === key)) { dup++; continue; }
    const item = state.itemMap.get(key);
    if (!item) { notfound++; continue; }
    _editingItems.push({ id: String(item.id), name: item.name, image: item.image || '', type: item.type || '', amount: '1' });
    added++;
  }
  renderEditingItemsList(listId);
  const search = document.getElementById(`${listId}-search`);
  const results = document.getElementById(`${listId}-results`);
  if (search)  search.value = '';
  if (results) results.innerHTML = '';
  toast(`เพิ่ม ${added} ไอเทม${dup ? ` · ซ้ำ ${dup}` : ''}${notfound ? ` · ไม่พบใน DB ${notfound}` : ''}`);
}

function handleItemPickerPaste(e, listId) {
  const text = (e.clipboardData || window.clipboardData)?.getData('text') || '';
  const ids = text.split(/[\s,;\t\r\n]+/).map(s => s.trim()).filter(Boolean);
  if (!ids.length) return;
  const anyKnown = ids.some(id => state.itemMap.has(String(id)));
  if (ids.length > 1 || anyKnown) {
    e.preventDefault();
    addEditingItemIds(ids, listId);
  }
}

function handleItemPickerKey(e, listId) {
  if (e.key !== 'Enter') return;
  e.preventDefault();
  const val = (e.target.value || '').trim();
  if (!val) return;
  if (state.itemMap.has(val)) {
    addEditingItemIds([val], listId);
  } else {
    const first = document.querySelector(`#${listId}-results .item-picker-result`);
    if (first) first.click();
  }
}

function setEditingItemAmount(idx, value) {
  if (_editingItems[idx]) _editingItems[idx].amount = String(value || '1');
}

function removeEditingItem(idx, listId) {
  _editingItems.splice(idx, 1);
  renderEditingItemsList(listId);
}

// ── Master Codes ──────────────────────────────────────
function renderCodes() {
  const el = document.getElementById('codes-list');
  if (!state.codes.length) { el.innerHTML = '<p class="empty-msg">ยังไม่มี Master Code</p>'; return; }

  el.innerHTML = state.codes.map((c, i) => `
    <div class="list-row">
      <div class="list-row-info">
        <strong class="code-mono">${esc(c.code)}</strong>
        <span>${esc(c.eventName)} · หมดอายุ: ${esc(c.expiresAt)}</span>
        <span class="status-badge">${esc(c.status)}</span>
      </div>
      <div class="list-row-actions">
        <button class="btn-sm btn-secondary" onclick="openEditCodeModal(${i})">แก้ไข</button>
        <button class="btn-sm btn-danger"    onclick="deleteCode(${i})">ลบ</button>
      </div>
    </div>
  `).join('');
}

function codeFormHtml(c) {
  return `
    <form id="cf" class="admin-form">
      <label class="field-label">Code <input type="text" name="code" value="${esc(c?.code || '')}" required class="mono-input" /></label>
      <label class="field-label">ชื่อกิจกรรม (พิมพ์เองได้ หรือเลือกจากรายการ)
        <input type="text" name="eventName" list="code-event-names" value="${esc(c?.eventName || '')}" placeholder="พิมพ์ชื่อกิจกรรม..." autocomplete="off" />
        <datalist id="code-event-names">
          ${state.events.map(ev => `<option value="${esc(ev.name)}"></option>`).join('')}
        </datalist>
      </label>
      <div class="form-row-2">
        <label class="field-label">สถานะ <input type="text" name="status" value="${esc(c?.status || 'พร้อมใช้')}" /></label>
        <label class="field-label">หมดอายุ <input type="text" name="expiresAt" value="${esc(c?.expiresAt || '-')}" /></label>
      </div>
      ${itemPickerHtml('code-items')}
      <div class="form-actions">
        <button type="submit" class="btn-primary">บันทึก</button>
        <button type="button" class="btn-ghost" onclick="closeModal()">ยกเลิก</button>
      </div>
    </form>
  `;
}

function openAddCodeModal() {
  _editingItems = [];
  showModal('เพิ่ม Master Code', codeFormHtml(null));
  renderEditingItemsList('code-items');
  document.getElementById('cf').onsubmit = e => {
    e.preventDefault();
    const f = new FormData(e.target);
    state.codes.push({ code: f.get('code').toUpperCase().trim(), eventName: f.get('eventName'), status: f.get('status'), expiresAt: f.get('expiresAt'), items: _editingItems.map(editItemToCodeItem) });
    persistData(); closeModal(); renderCodes();
  };
}

function openEditCodeModal(i) {
  const c = state.codes[i];
  _editingItems = Array.isArray(c.items) ? c.items.map(codeItemToEditItem) : [];
  showModal('แก้ไข Master Code', codeFormHtml(c));
  renderEditingItemsList('code-items');
  document.getElementById('cf').onsubmit = e => {
    e.preventDefault();
    const f = new FormData(e.target);
    state.codes[i] = { ...c, code: f.get('code').toUpperCase().trim(), eventName: f.get('eventName'), status: f.get('status'), expiresAt: f.get('expiresAt'), items: _editingItems.map(editItemToCodeItem) };
    persistData(); closeModal(); renderCodes();
  };
}

function deleteCode(i) {
  if (!confirm(`ลบ Master Code "${state.codes[i].code}"?`)) return;
  state.codes.splice(i, 1);
  persistData(); renderCodes();
}

// ── Items DB ──────────────────────────────────────────
const ITEM_TYPES = ['Ammo','Weapon','Attachment','Consumable','Equipment','Character','Components','Blueprint','Other'];

function renderItems() {
  const q       = (document.getElementById('items-search')?.value || '').toLowerCase().trim();
  const typeF   = document.getElementById('items-type-filter')?.value || '';
  const countEl = document.getElementById('items-count-label');

  // Populate type filter (once)
  const filterSel = document.getElementById('items-type-filter');
  if (filterSel && filterSel.options.length <= 1) {
    const types = [...new Set(state.items.map(it => it.type).filter(Boolean))].sort();
    types.forEach(t => {
      const opt = document.createElement('option');
      opt.value = t; opt.textContent = t;
      filterSel.appendChild(opt);
    });
  }

  let list = state.items;
  if (typeF) list = list.filter(it => it.type === typeF);
  if (q)     list = list.filter(it =>
    (it.name || '').toLowerCase().includes(q) ||
    String(it.id).toLowerCase().includes(q) ||
    (it.type || '').toLowerCase().includes(q)
  );

  if (countEl) countEl.textContent = `${list.length} / ${state.items.length} รายการ`;

  const grid = document.getElementById('items-grid');
  if (!list.length) {
    grid.innerHTML = '<p class="empty-msg">ไม่พบ Item</p>';
    return;
  }

  grid.innerHTML = list.map((item, _) => {
    const realIdx = state.items.indexOf(item);
    const img = item.image
      ? `<img src="${esc(item.image)}" alt="${esc(item.name)}" class="item-card-img" loading="lazy" onerror="this.style.display='none'" />`
      : `<div class="item-card-img item-card-no-img">?</div>`;
    return `
      <div class="item-card ${item.confirmed ? '' : 'item-card-unconfirmed'}">
        ${img}
        <div class="item-card-body">
          <span class="item-type-badge">${esc(item.type || '-')}</span>
          <strong class="item-card-name" title="${esc(item.name)}">${esc(item.name)}</strong>
          <code class="item-card-id">${esc(item.id)}</code>
          ${item.confirmed ? '' : '<span class="item-unconfirmed-badge">ยังไม่ยืนยัน</span>'}
        </div>
        <div class="item-card-actions">
          <button class="btn-xs btn-secondary" onclick="openEditItemModal(${realIdx})">แก้ไข</button>
          <button class="btn-xs btn-danger" onclick="deleteItem(${realIdx})">ลบ</button>
        </div>
      </div>
    `;
  }).join('');
}

function itemFormHtml(item) {
  const typeOptions = ITEM_TYPES.map(t =>
    `<option value="${t}" ${item?.type === t ? 'selected' : ''}>${t}</option>`
  ).join('');
  return `
    <form id="itf" class="admin-form">
      <div class="form-row-2">
        <label class="field-label">Item ID <input type="text" name="id" value="${esc(item?.id || '')}" required placeholder="400136" /></label>
        <label class="field-label">ประเภท <select name="type">${typeOptions}</select></label>
      </div>
      <label class="field-label">ชื่อ (EN) <input type="text" name="name" value="${esc(item?.name || '')}" required /></label>
      <label class="field-label">Image URL <input type="url" name="image" value="${esc(item?.image || '')}" placeholder="https://..." /></label>
      <div class="form-row-2">
        <label class="field-label">ราคา (Baht) <input type="text" name="price" value="${esc(item?.price || '')}" /></label>
        <label class="field-label">ราคา (GC) <input type="text" name="gcPrice" value="${esc(item?.gcPrice || '')}" /></label>
      </div>
      <label class="field-label">ราคา (Dollar) <input type="text" name="dollarPrice" value="${esc(item?.dollarPrice || '')}" /></label>
      <label class="field-label">Stackable
        <select name="stackable">
          <option value="Yes" ${item?.stackable === 'Yes' ? 'selected' : ''}>Yes</option>
          <option value="No"  ${item?.stackable === 'No'  ? 'selected' : ''}>No</option>
        </select>
      </label>
      <label class="field-label">คำอธิบาย <textarea name="description" rows="2">${esc(item?.description || '')}</textarea></label>
      <label class="field-label" style="flex-direction:row;align-items:center;gap:8px">
        <input type="checkbox" name="confirmed" ${item?.confirmed ? 'checked' : ''} style="width:auto" />
        ยืนยันแล้ว (confirmed)
      </label>
      <div class="form-actions">
        <button type="submit" class="btn-primary">บันทึก</button>
        <button type="button" class="btn-ghost" onclick="closeModal()">ยกเลิก</button>
      </div>
    </form>
  `;
}

function openAddItemModal() {
  showModal('เพิ่ม Item ใหม่', itemFormHtml(null));
  document.getElementById('itf').onsubmit = e => {
    e.preventDefault();
    const f = new FormData(e.target);
    const newItem = {
      id:           f.get('id').trim(),
      name:         f.get('name').trim(),
      price:        f.get('price').trim() || '-',
      gcPrice:      f.get('gcPrice').trim() || '-',
      dollarPrice:  f.get('dollarPrice').trim() || '-',
      type:         f.get('type'),
      stackable:    f.get('stackable'),
      image:        f.get('image').trim(),
      description:  f.get('description').trim(),
      confirmed:    f.has('confirmed'),
    };
    if (state.itemMap.has(newItem.id)) {
      showModal('เพิ่ม Item ใหม่', itemFormHtml(null));
      alert(`Item ID ${newItem.id} มีอยู่แล้วในฐานข้อมูล`);
      return;
    }
    state.items.push(newItem);
    rebuildItemMap();
    persistItemDb();
    closeModal();
    renderItems();
  };
}

function openEditItemModal(i) {
  const item = state.items[i];
  showModal(`แก้ไข Item — ${item.name}`, itemFormHtml(item));
  document.getElementById('itf').onsubmit = e => {
    e.preventDefault();
    const f = new FormData(e.target);
    state.items[i] = {
      id:           f.get('id').trim(),
      name:         f.get('name').trim(),
      price:        f.get('price').trim() || '-',
      gcPrice:      f.get('gcPrice').trim() || '-',
      dollarPrice:  f.get('dollarPrice').trim() || '-',
      type:         f.get('type'),
      stackable:    f.get('stackable'),
      image:        f.get('image').trim(),
      description:  f.get('description').trim(),
      confirmed:    f.has('confirmed'),
    };
    rebuildItemMap();
    persistItemDb();
    closeModal();
    renderItems();
  };
}

function deleteItem(i) {
  const item = state.items[i];
  if (!confirm(`ลบ "${item.name}" (ID: ${item.id}) ออกจาก Items DB?`)) return;
  state.items.splice(i, 1);
  rebuildItemMap();
  persistItemDb();
  renderItems();
}

function doExportItems() {
  const content = JSON.stringify({ items: state.items }, null, 2);
  const a = Object.assign(document.createElement('a'), {
    href: URL.createObjectURL(new Blob([content], { type: 'application/json' })),
    download: 'warz_data.json',
  });
  a.click();
  toast('ดาวน์โหลด warz_data.json แล้ว ✓');
}

// ── Publish (live, via Netlify function) ──────────────
const PUBLISH_TOKEN_KEY = 'warz_publish_token';

async function publishLive() {
  let token = localStorage.getItem(PUBLISH_TOKEN_KEY);
  if (!token) {
    token = prompt('ใส่ Publish Token (รหัสเผยแพร่) — ใส่ครั้งเดียว ระบบจะจำไว้:');
    if (!token) return;
    token = token.trim();
    localStorage.setItem(PUBLISH_TOKEN_KEY, token);
  }

  const btn = document.getElementById('publish-btn');
  if (btn) { btn.disabled = true; btn.textContent = '⏳ กำลังเผยแพร่...'; }
  toast('กำลังเผยแพร่ขึ้นเว็บ...');

  try {
    const res = await fetch('/api/save', {
      method: 'POST',
      headers: { 'content-type': 'application/json', 'x-publish-token': token },
      body: JSON.stringify({ events: state.events, codes: state.codes, items: state.items }),
    });

    if (res.status === 401) {
      localStorage.removeItem(PUBLISH_TOKEN_KEY);
      toast('Token ไม่ถูกต้อง — ลองใหม่อีกครั้ง');
      return;
    }
    if (!res.ok) { toast(`เผยแพร่ไม่สำเร็จ (${res.status})`); return; }

    const out = await res.json();
    localStorage.setItem('warz_last_published', out.savedAt || new Date().toISOString());
    renderExport();
    toast('เผยแพร่แล้ว ✓ ทุกคน + ผู้เล่นเห็นข้อมูลชุดเดียวกันทันที');
  } catch (err) {
    toast('เชื่อมต่อไม่ได้ — ลองใหม่อีกครั้ง');
  } finally {
    if (btn) { btn.disabled = false; btn.textContent = '🚀 เผยแพร่ขึ้นเว็บเลย'; }
  }
}

// ── Export ────────────────────────────────────────────
function renderExport() {
  const total = state.events.reduce((s, e) => s + e.winners.length, 0);
  const infoEl = document.getElementById('export-info');
  if (infoEl) {
    infoEl.innerHTML = `
      <div class="export-stat"><span>${state.events.length}</span> กิจกรรม</div>
      <div class="export-stat"><span>${total}</span> รายชื่อ</div>
      <div class="export-stat"><span>${state.codes.length}</span> Master Code</div>
    `;
  }
  const pubEl = document.getElementById('publish-info');
  if (pubEl) {
    const last = localStorage.getItem('warz_last_published');
    pubEl.textContent = last
      ? `เผยแพร่ล่าสุด: ${new Date(last).toLocaleString('th-TH')} · ${state.events.length} กิจกรรม · ${total} รายชื่อ`
      : `${state.events.length} กิจกรรม · ${total} รายชื่อ · ${state.codes.length} โค้ด`;
  }
}

function doExport() {
  const content =
    'window.WARZ_EVENTS = ' + JSON.stringify(state.events, null, 2) +
    ';\n\nwindow.WARZ_MASTER_CODES = ' + JSON.stringify(state.codes, null, 2) + ';\n';
  const a = Object.assign(document.createElement('a'), {
    href: URL.createObjectURL(new Blob([content], { type: 'text/javascript' })),
    download: 'events-data.js',
  });
  a.click();
  toast('ดาวน์โหลด events-data.js แล้ว ✓');
}

// ── Settings ──────────────────────────────────────────
function renderSettings() {}

// ── Modal ─────────────────────────────────────────────
function showModal(title, bodyHtml) {
  document.getElementById('modal-content').innerHTML = `<h3 class="modal-title">${title}</h3>${bodyHtml}`;
  document.getElementById('modal-overlay').classList.remove('hidden');
  document.querySelector('#modal-content input, #modal-content select')?.focus();
}
function closeModal() {
  document.getElementById('modal-overlay').classList.add('hidden');
}

// ── Toast ─────────────────────────────────────────────
let _toastTimer;
function toast(msg) {
  const el = document.getElementById('toast');
  el.textContent = msg;
  el.classList.add('is-visible');
  clearTimeout(_toastTimer);
  _toastTimer = setTimeout(() => el.classList.remove('is-visible'), 2400);
}

// ── Util ──────────────────────────────────────────────
function esc(s) {
  return String(s ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

// ── Import Preview ────────────────────────────────────
function showImportPreview(result, fileName) {
  const unverified = result.unverifiedCount || 0;
  const existingNames = new Set(state.events.map(e => e.name));

  const rows = result.events.map((ev, i) => `
    <tr>
      <td class="check-cell"><input type="checkbox" class="import-sheet-cb" data-idx="${i}" checked /></td>
      <td>${esc(ev.name)} ${existingNames.has(ev.name) ? '<span class="import-exists">มีอยู่แล้ว</span>' : ''}</td>
      <td style="text-align:right;font-family:var(--font-mono)">${ev.winners.length}</td>
      <td><span class="status-badge">${esc(ev.status)}</span></td>
    </tr>
  `).join('');

  const unverifiedWarning = unverified > 0
    ? `<div class="import-warning">⚠️ พบ <strong>${unverified}</strong> รายการที่ Item ID ไม่อยู่ใน Items DB — import ได้แต่ไม่มีรูปจาก warz_data</div>`
    : (state.itemMap.size > 0 ? `<div class="import-ok">✓ Item ทุกรายการตรวจสอบแล้วกับ Items DB</div>` : '');

  showModal(`📥 Import จาก ${esc(fileName)}`, `
    <p class="muted-label" style="margin-bottom:10px">เลือกกิจกรรม (sheet) ที่ต้องการนำเข้า — ติ๊กเฉพาะที่ต้องการ</p>
    ${unverifiedWarning}
    <div class="import-select-actions">
      <button type="button" class="btn-xs btn-ghost" onclick="toggleAllImportSheets(true)">เลือกทั้งหมด</button>
      <button type="button" class="btn-xs btn-ghost" onclick="toggleAllImportSheets(false)">ไม่เลือกเลย</button>
    </div>
    <div class="table-wrap" style="max-height:300px;overflow-y:auto;margin:10px 0">
      <table class="admin-table">
        <thead><tr><th class="check-cell"></th><th>กิจกรรม</th><th style="text-align:right">รายชื่อ</th><th>สถานะ</th></tr></thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
    <label class="field-label" style="margin-bottom:6px">โหมดนำเข้า</label>
    <label class="import-mode-opt"><input type="radio" name="import-mode" value="merge" checked /> <span><strong>เพิ่ม / อัปเดตทับของเดิม</strong> — กิจกรรมชื่อซ้ำจะถูกอัปเดต, กิจกรรมอื่นที่มีอยู่ยังอยู่ครบ</span></label>
    <label class="import-mode-opt"><input type="radio" name="import-mode" value="replace" /> <span><strong>แทนที่ทั้งหมด</strong> — ลบกิจกรรมเดิมทั้งหมด เหลือเฉพาะที่เลือก</span></label>
    <div class="form-actions">
      <button class="btn-primary" onclick="confirmImport(importPending)">✓ นำเข้าที่เลือก</button>
      <button class="btn-ghost" onclick="closeModal()">ยกเลิก</button>
    </div>
  `);
  window.importPending = result;
}

function toggleAllImportSheets(checked) {
  document.querySelectorAll('.import-sheet-cb').forEach(cb => { cb.checked = checked; });
}

function confirmImport(result) {
  const picked = [...document.querySelectorAll('.import-sheet-cb:checked')].map(cb => Number(cb.dataset.idx));
  if (!picked.length) { toast('ยังไม่ได้เลือกกิจกรรม'); return; }
  const mode = document.querySelector('input[name="import-mode"]:checked')?.value || 'merge';

  const events = picked.map(i => result.events[i]);
  const names  = new Set(events.map(e => e.name));
  const codes  = result.codes.filter(c => names.has(c.eventName));

  if (mode === 'replace') {
    state.events = events;
    state.codes  = codes;
  } else {
    // merge: update events with same name, add new ones
    events.forEach(ne => {
      const idx = state.events.findIndex(e => e.name === ne.name);
      if (idx >= 0) state.events[idx] = { ...ne, id: state.events[idx].id };
      else state.events.push(ne);
    });
    codes.forEach(nc => {
      const idx = state.codes.findIndex(c => c.code === nc.code);
      if (idx >= 0) state.codes[idx] = nc;
      else state.codes.push(nc);
    });
  }

  persistData();
  closeModal();
  showView('events');
  const w = events.reduce((s, e) => s + e.winners.length, 0);
  toast(`นำเข้า ${events.length} กิจกรรม, ${w} รายชื่อแล้ว ✓`);
}

// Load the shared data from the cloud store (so everyone edits the same data).
// Falls back to localStorage / baked-in events-data.js if unavailable.
async function loadSharedData() {
  try {
    const res = await fetch('/api/data', { cache: 'no-store' });
    if (!res.ok) return false;
    const data = await res.json();
    if (data && Array.isArray(data.events) && data.events.length) {
      state.events = data.events;
      state.codes  = Array.isArray(data.codes) ? data.codes : [];
      if (Array.isArray(data.items) && data.items.length) {
        state.items = data.items;
        rebuildItemMap();
      }
      cleanLoadedStatuses();
      localStorage.setItem('warz_shared_loaded_at', data.savedAt || new Date().toISOString());
      return true;
    }
  } catch (_) {}
  return false;
}

async function startApp() {
  document.getElementById('login-screen').classList.add('hidden');
  document.getElementById('admin-app').classList.remove('hidden');
  const shared = await loadSharedData();
  if (!shared) loadData();                 // fallback: localStorage / baked-in
  if (!state.items || !state.items.length) await loadItemDb();
  showView('dashboard');
}

// ── Init ──────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', async () => {
  await ensureDefaultCredentials();

  // Always wire up all event listeners (needed whether already logged in or not)

  // Login form submit
  document.getElementById('login-form').onsubmit = async e => {
    e.preventDefault();
    const user = document.getElementById('login-user').value.trim();
    const pass = document.getElementById('login-pass').value;
    const res  = await tryLogin(user, pass);
    if (res.ok) { startApp(); return; }
    document.getElementById('login-error').textContent = res.error;
  };

  // Nav buttons
  document.querySelectorAll('.admin-nav-btn').forEach(btn => {
    btn.addEventListener('click', () => showView(btn.dataset.view));
  });

  // Back from winners
  document.getElementById('back-to-events').addEventListener('click', () => showView('events'));

  // Import Excel
  document.getElementById('import-excel-btn').addEventListener('click', () => {
    document.getElementById('excel-file-input').click();
  });
  document.getElementById('excel-file-input').addEventListener('change', async e => {
    const file = e.target.files[0];
    if (!file) return;
    e.target.value = '';
    toast('กำลังอ่านไฟล์ Excel...');
    try {
      const result = await importExcelFile(file);
      showImportPreview(result, file.name);
    } catch (err) {
      showModal('เกิดข้อผิดพลาด', `<p style="color:#e47b6f">อ่านไฟล์ไม่ได้: ${esc(String(err))}</p><div class="form-actions"><button class="btn-ghost" onclick="closeModal()">ปิด</button></div>`);
    }
  });

  // CRUD triggers
  document.getElementById('add-event-btn').addEventListener('click',   openAddEventModal);
  document.getElementById('add-winner-btn').addEventListener('click',  openAddWinnerModal);
  document.getElementById('add-code-btn').addEventListener('click',    openAddCodeModal);
  document.getElementById('add-item-btn').addEventListener('click',    openAddItemModal);
  document.getElementById('export-btn').addEventListener('click',      doExport);
  document.getElementById('publish-btn').addEventListener('click',     publishLive);
  document.getElementById('export-items-btn').addEventListener('click',doExportItems);
  document.getElementById('reload-items-btn').addEventListener('click', reloadItemDbFromFile);
  document.getElementById('paste-winners-btn').addEventListener('click', openPasteWinnersModal);
  document.getElementById('check-dupes-btn').addEventListener('click', () => findDuplicates(false));
  document.getElementById('undo-btn').addEventListener('click', undoLast);
  document.getElementById('backup-btn').addEventListener('click', doBackup);
  document.getElementById('restore-btn').addEventListener('click', () => document.getElementById('restore-file-input').click());
  document.getElementById('restore-file-input').addEventListener('change', e => {
    const file = e.target.files[0];
    if (file) restoreBackup(file);
    e.target.value = '';
  });

  // Items live search + filter
  document.getElementById('items-search').addEventListener('input', renderItems);
  document.getElementById('items-type-filter').addEventListener('change', renderItems);

  // Winners live search
  document.getElementById('winners-search').addEventListener('input', renderWinners);
  document.getElementById('global-search').addEventListener('input', renderGlobalSearch);
  document.getElementById('winners-reward-filter').addEventListener('change', renderWinners);
  document.getElementById('winners-status-filter').addEventListener('change', renderWinners);

  // Modal close
  document.getElementById('modal-close').addEventListener('click', closeModal);
  document.getElementById('modal-overlay').addEventListener('click', e => {
    if (e.target.id === 'modal-overlay') closeModal();
  });
  document.addEventListener('keydown', e => { if (e.key === 'Escape') closeModal(); });

  // Logout
  document.getElementById('logout-btn').addEventListener('click', logout);

  // Change password
  document.getElementById('change-pass-form').onsubmit = async e => {
    e.preventDefault();
    const u  = document.getElementById('settings-user').value.trim();
    const p  = document.getElementById('settings-pass').value;
    const p2 = document.getElementById('settings-pass2').value;
    const errEl = document.getElementById('settings-error');
    if (!u)        { errEl.textContent = 'กรุณาใส่ Username'; return; }
    if (p.length < 6) { errEl.textContent = 'Password ต้องมีอย่างน้อย 6 ตัวอักษร'; return; }
    if (p !== p2)  { errEl.textContent = 'Password ไม่ตรงกัน'; return; }
    await saveCredentials(u, p);
    errEl.textContent = '';
    toast('เปลี่ยนรหัสผ่านแล้ว ✓');
  };

  // Revert to last published (re-load shared data from cloud)
  document.getElementById('revert-published-btn').addEventListener('click', revertToPublished);

  // Auto-start if session is active
  if (isLoggedIn()) {
    await startApp();
  }
});
