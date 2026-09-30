/**
 * Stock Count - Code.gs (revised)
 * - ยอดคงเหลือเก็บแยก 3 หน่วย (ลัง / แพ็ค / ชิ้น) ตามที่บันทึกใน Movements จริง
 *   IN = บวกเข้าแต่ละหน่วย, OUT = หักออกจากหน่วยเดียวกัน, ADJUST = ตั้งค่าแต่ละหน่วยตามที่กรอก
 * - ไม่มีการแปลง ลัง <-> แพ็ค <-> ชิ้น ในการแสดงผล / การตรวจยอด / การเปรียบเทียบยอดนับ
 * - อัตราแปลงหน่วย (PacksPerCarton, UnitsPerPack) ใช้เฉพาะ "เกณฑ์แจ้งเตือนใกล้หมด" เท่านั้น
 * - ทุกฟังก์ชันที่ client เรียกต้องส่ง token (ออกตอน login, เก็บใน CacheService)
 * - ฟังก์ชันที่ลงท้ายด้วย "_" เป็น private เรียกจากเบราว์เซอร์ไม่ได้
 * - [แก้ไข] ผู้ที่ไม่ใช่ Warehouse นับได้แม้ยังไม่มียอดตั้งต้นจาก WH โดยจะถูกนับเป็น "นับครั้งที่ 2"
 * - [แก้ไข] เปลี่ยนชื่อ EXP เป็น "EXP Date / Lot No." ในหัวตารางชีทที่ระบบสร้างใหม่
 */

const DEFAULT_UNITS_ = ['CARTON', 'PACK', 'PIECE'];
const SESSION_TTL_ = 21600;          // 6 ชั่วโมง (สูงสุดของ CacheService)
const HISTORY_LIMIT_ = 1500;
const MAX_LOGIN_FAILS_ = 5;
const REQUIRED_SHEETS_ = ['Products', 'Movements', 'Audits', 'StockSummary', 'Users', 'Rows'];
const EXP_HEADER_ = 'EXP Date / Lot No.';

// คอลัมน์ที่อนุญาตให้แก้ไข (index เริ่มที่ 0)
const EDIT_RULES_ = {
  Audits:    { cols: [6, 7, 8, 9, 10, 13],          nums: [8, 9, 10],    dates: [6, 7], userCol: 11 },
  Movements: { cols: [6, 7, 8, 9, 10, 11, 12, 15],  nums: [10, 11, 12],  dates: [6, 7], userCol: 13, typeCol: 8, destCol: 9, whOnly: true }
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Stock Count - ระบบจัดการคลังสินค้า')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

// ===========================================================================
// เมนูใน Google Sheets
// ===========================================================================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Stock Count')
    .addItem('ติดตั้ง / ตรวจสอบชีท (Setup)', 'menuSetup')
    .addItem('สร้าง StockSummary + LatestStock ใหม่', 'menuRebuild')
    .addItem('แปลงรหัสผ่านใน Users เป็นแฮช', 'menuHashPasswords')
    .addToUi();
}
function assertEditor_() {
  try { SpreadsheetApp.getUi(); }
  catch (e) { throw new Error('ฟังก์ชันนี้เรียกได้จากเมนูใน Google Sheets เท่านั้น'); }
}
function menuSetup() { assertEditor_(); SpreadsheetApp.getUi().alert(setupSheet_()); }
function menuRebuild() {
  assertEditor_();
  refreshStockSummary_();
  generateLatestStockSummary_();
  SpreadsheetApp.getUi().alert('สร้างใหม่เรียบร้อย');
}
function menuHashPasswords() {
  assertEditor_();
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  if (!sh || sh.getLastRow() < 2) return;
  const range = sh.getRange(2, 6, sh.getLastRow() - 1, 1);
  const vals = range.getValues();
  let n = 0;
  const out = vals.map(r => {
    const s = String(r[0]).trim();
    if (s && s.indexOf('sha256$') !== 0) { n++; return [makeHash_(s)]; }
    return [r[0]];
  });
  range.setNumberFormat('@').setValues(out);
  SpreadsheetApp.getUi().alert('แปลงรหัสผ่านเป็นแฮชแล้ว ' + n + ' บัญชี');
}

// ===========================================================================
// Helpers
// ===========================================================================
function tz_() { return Session.getScriptTimeZone(); }

function cleanCell_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, tz_(), 'dd/MM/yyyy');
  return String(v == null ? '' : v).replace(/^'/, '').trim();
}
function num_(v) {
  const n = Number(String(v == null ? '' : v).replace(/,/g, ''));
  return isFinite(n) ? n : 0;
}
function round2_(n) { return Math.round(n * 100) / 100; }
function fmtTs_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, tz_(), 'dd/MM/yyyy HH:mm');
  return String(v == null ? '' : v);
}
function safeText_(s) {
  s = String(s == null ? '' : s).trim();
  return /^[=+\-@]/.test(s) ? ' ' + s : s;
}
function newId_(prefix) {
  return prefix + '-' + Utilities.formatDate(new Date(), tz_(), 'yyMMddHHmmss') + '-' + Utilities.getUuid().slice(0, 4).toUpperCase();
}

function normalizeRowNameServer_(str) {
  const s = cleanCell_(str);
  if (!s || s === '-') return '-';
  const m = s.match(/^row\s*(\d+)$/i);
  if (m) return 'Row ' + String(m[1]).padStart(2, '0');
  return s;
}

function normalizeDate_(s) {
  s = cleanCell_(s);
  if (!s || s === '-') return '-';
  const m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/);
  if (!m) throw new Error('รูปแบบวันที่ไม่ถูกต้อง: ' + s + ' (ใช้ วว/ดด/ปปปป)');
  const d = +m[1], mo = +m[2];
  let y = +m[3];
  if (y < 100) y += 2000;
  if (y > 2400) y -= 543;
  const dt = new Date(y, mo - 1, d);
  if (dt.getFullYear() !== y || dt.getMonth() !== mo - 1 || dt.getDate() !== d) throw new Error('วันที่ไม่ถูกต้อง: ' + s);
  return ('0' + d).slice(-2) + '/' + ('0' + mo).slice(-2) + '/' + y;
}
function dateCell_(v) {
  const s = cleanCell_(v);
  if (!s || s === '-') return '-';
  try { return normalizeDate_(s); } catch (e) { return s; }
}
function parseDate_(s) {
  if (!s || s === '-') return null;
  const m = String(s).match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!m) return null;
  const d = new Date(+m[3], +m[2] - 1, +m[1]);
  return isNaN(d.getTime()) ? null : d;
}
function checkDateOrder_(mfd, exp) {
  const a = parseDate_(mfd), b = parseDate_(exp);
  if (a && b && b < a) throw new Error('วันหมดอายุต้องไม่ก่อนวันผลิต');
}

// ---------------------------------------------------------------------------
// ปริมาณแบบแยกหน่วย {c: ลัง, p: แพ็ค, u: ชิ้น}  (ไม่แปลงหน่วยระหว่างกัน)
// ---------------------------------------------------------------------------
function zeroQ_() { return { c: 0, p: 0, u: 0 }; }
function mkQ_(c, p, u) { return { c: c, p: p, u: u }; }
function addQ_(t, q, sign) {
  t.c = round2_(t.c + sign * q.c);
  t.p = round2_(t.p + sign * q.p);
  t.u = round2_(t.u + sign * q.u);
  return t;
}
function isZeroQ_(q) { return round2_(q.c) === 0 && round2_(q.p) === 0 && round2_(q.u) === 0; }
function qtyText_(q) {
  const parts = [];
  if (round2_(q.c) !== 0) parts.push(round2_(q.c) + ' ลัง');
  if (round2_(q.p) !== 0) parts.push(round2_(q.p) + ' แพ็ค');
  if (round2_(q.u) !== 0) parts.push(round2_(q.u) + ' ชิ้น');
  return parts.length ? parts.join(' ') : '0';
}

function unitInfo_(pCtn, uPack) {
  const cartonOnly = (pCtn === 0 && uPack === 0);
  const upp = uPack === 0 ? 1 : uPack;
  const multiplier = cartonOnly ? 1 : (pCtn === 0 ? 1 : pCtn) * upp;
  return { cartonOnly: cartonOnly, upp: upp, multiplier: multiplier };
}

// ค่าที่ใช้เทียบ "เกณฑ์แจ้งเตือน" เท่านั้น (ไม่ได้ใช้แสดงยอด)
// - ไม่ได้ตั้งอัตราแปลง (0/0): รวมตัวเลขตรงๆ
// - หน่วยเดียว: เทียบเป็นหน่วยนั้น | หลายหน่วย: เทียบเป็นลัง
function statusQty_(p, q) {
  const u = unitInfo_(p.packsPerCarton, p.unitsPerPack);
  if (u.cartonOnly) return q.c + q.p + q.u;
  const base = q.c * u.multiplier + q.p * u.upp + q.u;
  const size = p.singleUnit === 'PACK' ? u.upp : (p.singleUnit === 'PIECE' ? 1 : u.multiplier);
  return round2_(base / size);
}

function parseAllowedUnits_(raw, pCtn, uPack) {
  let units = String(raw == null ? '' : raw).split(',').map(s => s.trim().toUpperCase())
    .filter(u => DEFAULT_UNITS_.indexOf(u) !== -1);
  if (units.length === 0) units = (pCtn === 0 && uPack === 0) ? ['CARTON'] : DEFAULT_UNITS_.slice();
  return units;
}

// ===========================================================================
// Session / Auth
// ===========================================================================
function makeHash_(pw) {
  const salt = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
  return 'sha256$' + salt + '$' + hashPassword_(salt, pw);
}
function hashPassword_(salt, pw) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + pw, Utilities.Charset.UTF_8);
  return bytes.map(b => ('0' + (b < 0 ? b + 256 : b).toString(16)).slice(-2)).join('');
}
function verifyPassword_(stored, pw) {
  if (!stored || !pw) return false;
  if (stored.indexOf('sha256$') === 0) {
    const parts = stored.split('$');
    return parts.length === 3 && hashPassword_(parts[1], pw) === parts[2];
  }
  return stored === pw;
}

function loginUser(email, password) {
  const cache = CacheService.getScriptCache();
  const em = String(email || '').trim().toLowerCase();
  const pw = String(password || '').trim();
  const failKey = 'fail_' + em;
  if (Number(cache.get(failKey) || 0) >= MAX_LOGIN_FAILS_) {
    return JSON.stringify({ success: false, message: 'เข้าสู่ระบบผิดหลายครั้ง กรุณารอ 10 นาทีแล้วลองใหม่' });
  }
  const uSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  if (!uSheet) return JSON.stringify({ success: false, message: 'ไม่พบฐานข้อมูลผู้ใช้ (กรุณาติดตั้งระบบจากเมนู Stock Count)' });

  const data = uSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][3]).trim() !== 'Active') continue;
    if (String(data[i][4]).trim().toLowerCase() !== em) continue;
    const stored = String(data[i][5]).trim();
    if (verifyPassword_(stored, pw)) {
      if (stored.indexOf('sha256$') !== 0) {
        uSheet.getRange(i + 1, 6).setNumberFormat('@').setValue(makeHash_(pw));
      }
      cache.remove(failKey);
      const user = { name: String(data[i][1]), role: String(data[i][2]).trim() };
      const token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, '');
      cache.put('sess_' + token, JSON.stringify(user), SESSION_TTL_);
      return JSON.stringify({ success: true, token: token, user: user });
    }
  }
  cache.put(failKey, String(Number(cache.get(failKey) || 0) + 1), 600);
  return JSON.stringify({ success: false, message: 'อีเมลหรือรหัสผ่านไม่ถูกต้อง!' });
}

function logoutUser(token) {
  if (token) CacheService.getScriptCache().remove('sess_' + token);
  return JSON.stringify({ success: true });
}

function requireUser_(token, roles) {
  if (!token) throw new Error('SESSION_EXPIRED');
  const cache = CacheService.getScriptCache();
  const raw = cache.get('sess_' + token);
  if (!raw) throw new Error('SESSION_EXPIRED');
  cache.put('sess_' + token, raw, SESSION_TTL_);
  const u = JSON.parse(raw);
  if (roles && roles.indexOf(u.role) === -1) throw new Error('ไม่มีสิทธิ์ดำเนินการ (เฉพาะ ' + roles.join(', ') + ')');
  return u;
}

// ===========================================================================
// Setup
// ===========================================================================
function setupSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let pSheet = ss.getSheetByName('Products');
  if (!pSheet) {
    pSheet = ss.insertSheet('Products');
    pSheet.appendRow(['Barcode', 'Name', 'Brand', 'PacksPerCarton', 'UnitsPerPack', 'MinStock', 'Status', 'ImageURL', 'AllowedUnits']);
  } else {
    const lastCol = pSheet.getLastColumn();
    const headerVals = lastCol > 0 ? pSheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
    if (headerVals.indexOf('AllowedUnits') === -1) pSheet.getRange(1, 9).setValue('AllowedUnits');
  }
  pSheet.getRange('A:A').setNumberFormat('@');

  let mSheet = ss.getSheetByName('Movements');
  if (!mSheet) {
    mSheet = ss.insertSheet('Movements');
    mSheet.appendRow(['Timestamp', 'ID', 'LocationRow', 'Barcode', 'Brand', 'Name', 'MFD', EXP_HEADER_,
      'Type', 'Destination', 'Cartons', 'Packs', 'Pieces', 'User', 'UserRole', 'Note']);
  }
  mSheet.getRange('B:D').setNumberFormat('@');
  mSheet.getRange('G:H').setNumberFormat('@');

  let aSheet = ss.getSheetByName('Audits');
  if (!aSheet) {
    aSheet = ss.insertSheet('Audits');
    aSheet.appendRow(['Timestamp', 'AuditID', 'LocationRow', 'Barcode', 'Brand', 'Name', 'MFD', EXP_HEADER_,
      'Cartons', 'Packs', 'Pieces', 'User', 'UserRole', 'Note']);
  }
  aSheet.getRange('B:D').setNumberFormat('@');
  aSheet.getRange('G:H').setNumberFormat('@');

  let sSheet = ss.getSheetByName('StockSummary');
  if (!sSheet) {
    sSheet = ss.insertSheet('StockSummary');
    sSheet.appendRow(['Barcode', 'Name', 'Brand', 'LocationRow', 'MFD', EXP_HEADER_, 'ShelfLife', 'RemainDays', 'PercentDay',
      'RemainCartons', 'RemainPacks', 'RemainPieces', 'TotalUnits', 'MinStock', 'Status', 'LastUpdated']);
  }
  sSheet.getRange('A:A').setNumberFormat('@');
  sSheet.getRange('D:F').setNumberFormat('@');

  let uSheet = ss.getSheetByName('Users');
  if (!uSheet) {
    uSheet = ss.insertSheet('Users');
    uSheet.appendRow(['UserID', 'Name', 'Role', 'Status', 'Email', 'Password']);
    uSheet.appendRow(['USR-001', 'โกดังสมชาย', 'Warehouse', 'Active', 'wh@test.com', '1234']);
    uSheet.appendRow(['USR-002', 'พนักงานตรวจ A', 'Staff', 'Active', 'staff@test.com', '1234']);
  }
  uSheet.getRange('F:F').setNumberFormat('@');

  let rSheet = ss.getSheetByName('Rows');
  if (!rSheet) {
    rSheet = ss.insertSheet('Rows');
    rSheet.appendRow(['RowName', 'Type', 'Status']);
    rSheet.appendRow(['Row 01', 'Online', 'Active']);
  }

  refreshStockSummary_();
  return 'ติดตั้งระบบเสร็จสิ้น!';
}
function ensureSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (REQUIRED_SHEETS_.some(n => !ss.getSheetByName(n))) setupSheet_();
}

// ===========================================================================
// Stock calculation (เก็บยอดแยกตามหน่วยจริง)
// ===========================================================================
function readProductMeta_(sheet) {
  const data = sheet.getDataRange().getValues();
  const map = {}, list = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const barcode = cleanCell_(r[0]);
    if (!barcode) continue;
    const pCtn = num_(r[3]), uPack = num_(r[4]);
    const allowed = parseAllowedUnits_(r[8], pCtn, uPack);
    const meta = {
      barcode: barcode, name: String(r[1] == null ? '' : r[1]), brand: String(r[2] == null ? '' : r[2]),
      packsPerCarton: pCtn, unitsPerPack: uPack, multiplier: unitInfo_(pCtn, uPack).multiplier,
      minStock: num_(r[5]), status: String(r[6] || 'Active').trim() || 'Active',
      imageUrl: String(r[7] || '').trim(),
      allowedUnits: allowed,
      singleUnit: allowed.length === 1 ? allowed[0] : ''
    };
    if (map[barcode]) { const idx = list.indexOf(map[barcode]); if (idx !== -1) list.splice(idx, 1); }
    map[barcode] = meta;
    list.push(meta);
  }
  return { map: map, list: list };
}

function computeStockState_(mData) {
  const goodMap = {}, totalMap = {}, byBarcode = {};
  for (let i = 1; i < mData.length; i++) {
    const r = mData[i];
    const barcode = cleanCell_(r[3]);
    if (!barcode) continue;
    const loc = normalizeRowNameServer_(r[2]);
    const mfd = dateCell_(r[6]), exp = dateCell_(r[7]);
    const type = String(r[8]).trim();
    const q = mkQ_(num_(r[10]), num_(r[11]), num_(r[12]));

    const key = barcode + '|' + loc + '|' + mfd + '|' + exp;
    if (!(key in goodMap)) {
      goodMap[key] = zeroQ_();
      (byBarcode[barcode] = byBarcode[barcode] || []).push(key);
    }
    if (!(barcode in totalMap)) totalMap[barcode] = zeroQ_();

    if (type === 'IN') { addQ_(goodMap[key], q, 1); addQ_(totalMap[barcode], q, 1); }
    else if (type === 'OUT') { addQ_(goodMap[key], q, -1); addQ_(totalMap[barcode], q, -1); }
    else if (type === 'ADJUST') {
      // ตั้งยอดของล็อตนี้เท่ากับที่กรอก (แต่ละหน่วย) แล้วปรับยอดรวมตามส่วนต่าง
      const diff = mkQ_(q.c - goodMap[key].c, q.p - goodMap[key].p, q.u - goodMap[key].u);
      goodMap[key] = mkQ_(round2_(q.c), round2_(q.p), round2_(q.u));
      addQ_(totalMap[barcode], diff, 1);
    }
  }
  return { goodMap: goodMap, totalMap: totalMap, byBarcode: byBarcode };
}

function buildStockRows_(productList, state) {
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const DAY = 86400000;
  const rows = [];
  productList.forEach(p => {
    const totalQty = state.totalMap[p.barcode] ? mkQ_(state.totalMap[p.barcode].c, state.totalMap[p.barcode].p, state.totalMap[p.barcode].u) : zeroQ_();
    const hasPositive = totalQty.c > 0 || totalQty.p > 0 || totalQty.u > 0;
    let status = 'ปกติ';
    if (!hasPositive) status = 'หมดสต็อก';
    else if (statusQty_(p, totalQty) <= p.minStock) status = 'ใกล้หมด';

    const keys = (state.byBarcode[p.barcode] || []).filter(k => !isZeroQ_(state.goodMap[k])).sort();
    if (keys.length === 0) {
      rows.push({ p: p, loc: '-', mfd: '-', exp: '-', shelfLife: '-', remainDays: '-', percentDay: '-', qty: zeroQ_(), totalQty: totalQty, status: status });
      return;
    }
    keys.forEach(k => {
      const parts = k.split('|');
      const cur = state.goodMap[k];
      let shelfLife = '-', remainDays = '-', percentDay = '-';
      const mfdD = parseDate_(parts[2]), expD = parseDate_(parts[3]);
      if (mfdD && expD) {
        shelfLife = Math.round((expD - mfdD) / DAY);
        remainDays = Math.ceil((expD - today) / DAY);
        percentDay = shelfLife > 0 ? Math.round((remainDays / shelfLife) * 100) : 0;
      } else if (expD) {
        remainDays = Math.ceil((expD - today) / DAY);
      }
      rows.push({ p: p, loc: parts[1], mfd: parts[2], exp: parts[3], shelfLife: shelfLife, remainDays: remainDays, percentDay: percentDay,
        qty: mkQ_(cur.c, cur.p, cur.u), totalQty: totalQty, status: status });
    });
  });
  return rows;
}

function computeInventory_(ss) {
  const pm = readProductMeta_(ss.getSheetByName('Products'));
  const mData = ss.getSheetByName('Movements').getDataRange().getValues();
  const state = computeStockState_(mData);
  const rows = buildStockRows_(pm.list, state);
  return { pm: pm, mData: mData, state: state, rows: rows };
}

function refreshStockSummary_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sSheet = ss.getSheetByName('StockSummary');
  if (!sSheet) sSheet = ss.insertSheet('StockSummary');
  const inv = computeInventory_(ss);
  const now = Utilities.formatDate(new Date(), tz_(), 'dd/MM/yyyy HH:mm');

  const header = ['Barcode', 'Name', 'Brand', 'LocationRow', 'MFD', EXP_HEADER_, 'ShelfLife', 'RemainDays', 'PercentDay',
    'RemainCartons', 'RemainPacks', 'RemainPieces', 'TotalUnits', 'MinStock', 'Status', 'LastUpdated'];
  sSheet.getRange('A:A').setNumberFormat('@');
  sSheet.getRange('D:F').setNumberFormat('@');
  sSheet.getRange('M:M').setNumberFormat('@');
  sSheet.getRange(1, 1, 1, 16).setValues([header]).setFontWeight('bold').setBackground('#e2e8f0');
  if (sSheet.getLastRow() > 1) sSheet.getRange(2, 1, sSheet.getLastRow() - 1, 16).clearContent();

  // TotalUnits = ข้อความยอดตามหน่วยจริง (เช่น "2 ลัง 3 แพ็ค")
  const out = inv.rows.map(r => [r.p.barcode, r.p.name, r.p.brand, r.loc, r.mfd, r.exp, r.shelfLife, r.remainDays, r.percentDay,
    r.qty.c, r.qty.p, r.qty.u, qtyText_(r.qty), r.p.minStock, r.status, now]);
  if (out.length) sSheet.getRange(2, 1, out.length, 16).setValues(out);
}

// ===========================================================================
// อ่านข้อมูลทั้งหมดให้หน้าเว็บ
// ===========================================================================
function getInventoryData(token) {
  requireUser_(token);
  ensureSheets_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inv = computeInventory_(ss);
  const aData = ss.getSheetByName('Audits').getDataRange().getValues();
  const rData = ss.getSheetByName('Rows').getDataRange().getValues();
  const tzName = tz_();

  const masterRows = [];
  for (let i = 1; i < rData.length; i++) {
    const rowName = String(rData[i][0] || '').trim();
    const type = String(rData[i][1] || '').trim();
    const status = String(rData[i][2] || '').trim();
    if (rowName && status !== 'Inactive') masterRows.push({ row: rowName, type: type || 'General' });
  }

  const hasRecords = {};
  for (let i = 1; i < inv.mData.length; i++) { const b = cleanCell_(inv.mData[i][3]); if (b) hasRecords[b] = true; }
  for (let i = 1; i < aData.length; i++) { const b = cleanCell_(aData[i][3]); if (b) hasRecords[b] = true; }

  // ---- products (ต่อ Row/ล็อต) ----
  const products = inv.rows.map(r => ({
    barcode: r.p.barcode, name: r.p.name, brand: r.p.brand || '-',
    locationRow: r.loc, mfdDate: r.mfd, expDate: r.exp,
    shelfLife: r.shelfLife, remainDays: r.remainDays, percentDay: r.percentDay,
    packsPerCarton: r.p.packsPerCarton, unitsPerPack: r.p.unitsPerPack,
    qty: r.qty, totalQty: r.totalQty,
    minStock: r.p.minStock, status: r.p.status,
    isLow: r.status === 'ใกล้หมด', isOut: r.status === 'หมดสต็อก',
    imageUrl: r.p.imageUrl, allowedUnits: r.p.allowedUnits, singleUnit: r.p.singleUnit
  }));

  let lowStockCount = 0, outOfStockCount = 0, totalCartons = 0, totalSku = 0;
  const counted = {};
  const statusByBarcode = {};
  inv.rows.forEach(r => {
    statusByBarcode[r.p.barcode] = { totalQty: r.totalQty, isLow: r.status === 'ใกล้หมด', isOut: r.status === 'หมดสต็อก' };
    if (r.p.status !== 'Active') return;
    if (!counted[r.p.barcode]) {
      counted[r.p.barcode] = true;
      totalSku++;
      if (r.status === 'หมดสต็อก') outOfStockCount++;
      else if (r.status === 'ใกล้หมด') lowStockCount++;
      totalCartons += r.totalQty.c;   // นับเฉพาะ "ลัง" ที่บันทึกจริง ไม่แปลงจากหน่วยอื่น
    }
  });

  const uniqueProducts = inv.pm.list.map(p => {
    const st = statusByBarcode[p.barcode] || { totalQty: zeroQ_(), isLow: false, isOut: true };
    return Object.assign({}, p, { hasRecords: !!hasRecords[p.barcode], totalQty: st.totalQty, isLow: st.isLow, isOut: st.isOut });
  });

  // ---- history (รับ-จ่าย) ----
  const history = [];
  let truncated = false;
  const todayStr = Utilities.formatDate(new Date(), tzName, 'dd/MM/yyyy');
  let todayMoves = 0;
  for (let i = 1; i < inv.mData.length; i++) {
    const t = inv.mData[i][0];
    if (t instanceof Date && Utilities.formatDate(t, tzName, 'dd/MM/yyyy') === todayStr) todayMoves++;
  }
  for (let i = inv.mData.length - 1; i >= 1; i--) {
    const row = inv.mData[i];
    const barcode = cleanCell_(row[3]);
    if (!barcode) continue;
    if (history.length >= HISTORY_LIMIT_) { truncated = true; break; }
    history.push({
      sheetRowIndex: i + 1, id: String(row[1] || ''), time: fmtTs_(row[0]),
      locationRow: normalizeRowNameServer_(row[2]), barcode: barcode,
      brand: String(row[4] || '-'), name: String(row[5] || '-'),
      mfdDate: dateCell_(row[6]), expDate: dateCell_(row[7]),
      type: String(row[8]).trim(), destination: String(row[9] || '-'),
      qty: mkQ_(num_(row[10]), num_(row[11]), num_(row[12])),
      user: String(row[13] || ''), userRole: String(row[14] || 'Staff').trim(), note: String(row[15] || '')
    });
  }

  // ---- audit pivot ----
  const pivot = {};
  for (let i = 1; i < aData.length; i++) {
    const row = aData[i];
    const barcode = cleanCell_(row[3]);
    if (!barcode) continue;
    const locRow = normalizeRowNameServer_(row[2]);
    const brand = String(row[4] || '-'), name = String(row[5] || '-');
    const mfd = dateCell_(row[6]), exp = dateCell_(row[7]);
    const goodQty = mkQ_(num_(row[8]), num_(row[9]), num_(row[10]));
    const user = String(row[11] || '').trim();
    const userRole = String(row[12] || 'Staff').trim();
    const note = String(row[13] || '');

    const key = locRow + '|' + barcode + '|' + mfd + '|' + exp;
    const prev = pivot[key];
    const hasOtherRole = prev && prev.roles.some(r => r !== 'Warehouse');
    // [แก้ไข] WH นับเริ่มรอบใหม่ (ล้างของเก่า) เฉพาะเมื่อ WH เคยนับตั้งต้นไปแล้ว
    // ถ้าก่อนหน้านี้มีแต่ Staff นับ (ยังไม่มียอด WH) ให้ WH เติมเป็นนับครั้งที่ 1 โดยไม่ล้างยอดของ Staff
    if (!prev || (userRole === 'Warehouse' && hasOtherRole && prev.startGood !== '-')) {
      pivot[key] = {
        row: locRow, barcode: barcode, brand: brand, name: name, mfdDate: mfd, expDate: exp,
        startGood: '-', startUser: '-', startSheetRowIndex: null,
        count2Good: '-', count2User: '-', count2SheetRowIndex: null,
        count3Good: '-', count3User: '-', count3SheetRowIndex: null,
        notes: [], roles: []
      };
    }
    const g = pivot[key];
    g.roles.push(userRole);
    const cleanNote = note.replace(/\((เริ่มต้น|นับครั้งที่\s*\d+)\)/g, '').trim();
    if (cleanNote) g.notes.push(cleanNote);

    if (userRole === 'Warehouse' && g.startGood === '-') {
      g.startGood = goodQty; g.startUser = user; g.startSheetRowIndex = i + 1;
    } else if (g.count2Good === '-') {
      g.count2Good = goodQty; g.count2User = user; g.count2SheetRowIndex = i + 1;
    } else if (g.count3Good === '-') {
      g.count3Good = goodQty; g.count3User = user; g.count3SheetRowIndex = i + 1;
    } else {
      g.notes.push(user + ': ยอด ' + qtyText_(goodQty));
    }
  }
  const auditReportData = Object.keys(pivot).map(k => {
    const g = pivot[k];
    g.noteSummary = g.notes.length ? Array.from(new Set(g.notes)).join(' • ') : '-';
    return g;
  });

  return JSON.stringify({
    products: products, uniqueProducts: uniqueProducts, history: history, historyTruncated: truncated,
    auditReportData: auditReportData, masterRows: masterRows,
    summary: {
      totalSku: totalSku, totalCartons: round2_(totalCartons),
      lowStockCount: lowStockCount, outOfStockCount: outOfStockCount, todayMovements: todayMoves
    }
  });
}

// ===========================================================================
// บันทึกรับเข้า / จ่ายออก / ปรับยอด  (Warehouse เท่านั้น)
// ===========================================================================
function readQty_(form, p) {
  const allowed = p.allowedUnits;
  const pick = (unit, val) => {
    if (allowed.indexOf(unit) === -1) return 0;
    const n = (val === '' || val == null) ? 0 : Number(val);
    if (!isFinite(n) || n < 0) throw new Error('จำนวนต้องเป็นตัวเลขที่ไม่ติดลบ');
    return n;
  };
  const ctn = pick('CARTON', form.cartons), pck = pick('PACK', form.packs), pcs = pick('PIECE', form.pieces);
  // total ใช้ตรวจว่า "กรอกอย่างน้อย 1 หน่วย" เท่านั้น (ไม่ใช่ยอดแปลงหน่วย)
  return { ctn: ctn, pck: pck, pcs: pcs, total: ctn + pck + pcs };
}

function recordMovement(token, form) {
  const user = requireUser_(token, ['Warehouse']);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    form = form || {};
    const type = String(form.type || '').toUpperCase();
    if (['IN', 'OUT', 'ADJUST'].indexOf(type) === -1) throw new Error('ประเภทรายการไม่ถูกต้อง');

    const inv = computeInventory_(ss);
    const barcode = cleanCell_(form.barcode);
    const p = inv.pm.map[barcode];
    if (!p) throw new Error('ไม่พบสินค้าบาร์โค้ดนี้ในระบบ');
    if (p.status !== 'Active') throw new Error('สินค้านี้ถูกตั้งเป็น Inactive');

    const q = readQty_(form, p);
    if (type !== 'ADJUST' && q.total <= 0) throw new Error('กรุณาระบุจำนวนอย่างน้อย 1 หน่วย');

    const mfd = normalizeDate_(form.mfdDate), exp = normalizeDate_(form.expDate);
    checkDateOrder_(mfd, exp);
    const locRow = normalizeRowNameServer_(form.locationRow);

    let destination = '-';
    if (type === 'OUT') {
      destination = safeText_(form.destination);
      if (!destination) throw new Error('กรุณาระบุปลายทาง');
      const avail = inv.state.goodMap[barcode + '|' + locRow + '|' + mfd + '|' + exp] || zeroQ_();
      const eps = 0.0001;
      if (q.ctn > avail.c + eps || q.pck > avail.p + eps || q.pcs > avail.u + eps) {
        throw new Error('ยอดคงเหลือใน ' + locRow + ' (MFD ' + mfd + ' / ' + EXP_HEADER_ + ' ' + exp + ') มีไม่พอ: คงเหลือ ' + qtyText_(avail) +
          ' (ระบบไม่แปลงหน่วย ต้องจ่ายตามหน่วยเดียวกับที่รับเข้า)');
      }
    }

    ss.getSheetByName('Movements').appendRow([
      new Date(), newId_('MOV'), locRow, barcode, p.brand, p.name, mfd, exp,
      type, destination, q.ctn, q.pck, q.pcs,
      user.name, user.role, safeText_(form.note)
    ]);

    refreshStockSummary_();
    return JSON.stringify({ success: true, message: 'บันทึกรายการ ' + type + ' สำเร็จเรียบร้อย' });
  } finally {
    lock.releaseLock();
  }
}

// ===========================================================================
// บันทึกการนับสินค้า (ทุก Role)
// - Warehouse = นับครั้งที่ 1
// - ไม่ใช่ Warehouse = นับครั้งที่ 2, 3, ... (นับได้เลยแม้ยังไม่มียอดตั้งต้นจาก WH)
// ===========================================================================
function recordAudit(token, form) {
  const user = requireUser_(token);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    form = form || {};
    const barcode = cleanCell_(form.barcode);
    const pm = readProductMeta_(ss.getSheetByName('Products'));
    const p = pm.map[barcode];
    if (!p) throw new Error('ไม่พบสินค้าบาร์โค้ดนี้ในระบบ');
    if (p.status !== 'Active') throw new Error('สินค้านี้ถูกตั้งเป็น Inactive');

    const q = readQty_(form, p);
    if (q.total <= 0) throw new Error('กรุณาระบุจำนวนอย่างน้อย 1 หน่วย');

    const mfd = normalizeDate_(form.mfdDate), exp = normalizeDate_(form.expDate);
    checkDateOrder_(mfd, exp);
    const locRow = normalizeRowNameServer_(form.locationRow);
    if (locRow === '-') throw new Error('กรุณาระบุแถว (Row)');

    const aSheet = ss.getSheetByName('Audits');
    const aData = aSheet.getDataRange().getValues();
    const same = r => normalizeRowNameServer_(r[2]).toLowerCase() === locRow.toLowerCase() &&
      cleanCell_(r[3]) === barcode && dateCell_(r[6]) === mfd && dateCell_(r[7]) === exp;
    const roleOf = r => String(r[12] || '').trim();

    // [แก้ไข] ตัดเงื่อนไขบังคับให้ต้องมียอดตั้งต้นจาก Warehouse ออก
    // ผู้ที่ไม่ใช่ Warehouse นับได้เสมอ และถ้ายังไม่มียอด WH จะเริ่มเป็น "นับครั้งที่ 2"

    let lastIdx = -1;
    for (let i = aData.length - 1; i >= 1; i--) { if (same(aData[i])) { lastIdx = i; break; } }
    if (lastIdx > 0 && String(aData[lastIdx][11] || '').trim().toLowerCase() === user.name.trim().toLowerCase()) {
      const r = aData[lastIdx];
      const lastNote = String(r[13] || '');
      const addText = '(บวกทบ +' + qtyText_(mkQ_(q.ctn, q.pck, q.pcs)) + ')';
      aSheet.getRange(lastIdx + 1, 9, 1, 3).setValues([[num_(r[8]) + q.ctn, num_(r[9]) + q.pck, num_(r[10]) + q.pcs]]);
      aSheet.getRange(lastIdx + 1, 14).setValue(lastNote ? lastNote + ' ' + addText : addText);
      generateLatestStockSummary_();
      return JSON.stringify({ success: true, message: 'ทำการบวกทบยอดเดิมเรียบร้อย!' });
    }

    let statusTag;
    if (user.role === 'Warehouse') {
      statusTag = '(นับครั้งที่ 1)';
    } else {
      const existingNonWh = aData.slice(1).filter(r => same(r) && roleOf(r) !== 'Warehouse').length;
      statusTag = '(นับครั้งที่ ' + (existingNonWh + 2) + ')';
    }
    const note = safeText_(form.note);

    aSheet.appendRow([
      new Date(), newId_('AUD'), locRow, barcode, p.brand, p.name, mfd, exp,
      q.ctn, q.pck, q.pcs, user.name, user.role, note ? note + ' ' + statusTag : statusTag
    ]);

    generateLatestStockSummary_();
    return JSON.stringify({ success: true, message: 'บันทึกการนับ [' + locRow + '] สำเร็จ ' + statusTag });
  } finally {
    lock.releaseLock();
  }
}

// ===========================================================================
// แก้ไข / ลบ / อ่านแถว (เฉพาะ Audits และ Movements)
// ===========================================================================
function authorizeRecord_(user, sheetName, rowValues) {
  const rule = EDIT_RULES_[sheetName];
  if (!rule) throw new Error('ไม่อนุญาตให้แก้ไขชีทนี้');
  if (user.role === 'Warehouse') return rule;
  if (rule.whOnly) throw new Error('ไม่มีสิทธิ์ดำเนินการ (เฉพาะ Warehouse)');
  if (rowValues && String(rowValues[rule.userCol] || '').trim().toLowerCase() !== user.name.trim().toLowerCase()) {
    throw new Error('แก้ไขได้เฉพาะรายการของตัวเอง');
  }
  return rule;
}
function loadRecordRow_(sheet, rowIndex, expectedId) {
  rowIndex = Number(rowIndex);
  if (!rowIndex || rowIndex < 2 || rowIndex > sheet.getLastRow()) throw new Error('ไม่พบแถวข้อมูล กรุณารีเฟรชแล้วลองใหม่');
  const values = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (String(values[1]) !== String(expectedId)) throw new Error('ข้อมูลถูกเปลี่ยนแปลงแล้ว กรุณารีเฟรชแล้วลองใหม่');
  return values;
}

function getRowData(token, sheetName, rowIndex) {
  const user = requireUser_(token);
  try {
    const rule = EDIT_RULES_[sheetName];
    if (!rule) throw new Error('ไม่อนุญาตให้เข้าถึงชีทนี้');
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    rowIndex = Number(rowIndex);
    if (!rowIndex || rowIndex < 2 || rowIndex > sheet.getLastRow()) throw new Error('ไม่พบแถวข้อมูล กรุณารีเฟรชแล้วลองใหม่');
    const values = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    authorizeRecord_(user, sheetName, values);
    return JSON.stringify({
      success: true,
      values: values.map((v, i) => i === 0 ? fmtTs_(v) : (v instanceof Date ? cleanCell_(v) : String(v == null ? '' : v).replace(/^'/, '')))
    });
  } catch (err) {
    return JSON.stringify({ success: false, message: err.message });
  }
}

function updateRecord(token, sheetName, rowIndex, expectedId, rowData) {
  const user = requireUser_(token);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!EDIT_RULES_[sheetName] || !sheet) throw new Error('ไม่อนุญาตให้แก้ไขชีทนี้');
    const values = loadRecordRow_(sheet, rowIndex, expectedId);
    const rule = authorizeRecord_(user, sheetName, values);
    rowData = rowData || [];

    const next = values.slice();
    rule.cols.forEach(c => {
      let v = rowData[c];
      v = v == null ? '' : String(v).trim();
      if (rule.dates.indexOf(c) !== -1) v = normalizeDate_(v);
      else if (rule.nums.indexOf(c) !== -1) {
        const n = v === '' ? 0 : Number(v);
        if (!isFinite(n) || n < 0) throw new Error('จำนวนต้องเป็นตัวเลขที่ไม่ติดลบ');
        v = n;
      } else if (c === rule.typeCol) {
        v = v.toUpperCase();
        if (['IN', 'OUT', 'ADJUST'].indexOf(v) === -1) throw new Error('ประเภทรายการไม่ถูกต้อง');
      } else v = safeText_(v);
      next[c] = v;
    });
    checkDateOrder_(next[6], next[7]);

    sheet.getRange(Number(rowIndex), 1, 1, next.length).setValues([next]);
    if (sheetName === 'Movements') refreshStockSummary_(); else generateLatestStockSummary_();
    return JSON.stringify({ success: true });
  } finally {
    lock.releaseLock();
  }
}

function deleteRecord(token, sheetName, rowIndex, expectedId) {
  const user = requireUser_(token);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!EDIT_RULES_[sheetName] || !sheet) throw new Error('ไม่อนุญาตให้ลบข้อมูลชีทนี้');
    const values = loadRecordRow_(sheet, rowIndex, expectedId);
    authorizeRecord_(user, sheetName, values);
    sheet.deleteRow(Number(rowIndex));
    if (sheetName === 'Movements') refreshStockSummary_(); else generateLatestStockSummary_();
    return JSON.stringify({ success: true });
  } finally {
    lock.releaseLock();
  }
}

// ===========================================================================
// จัดการสินค้า (Warehouse เท่านั้น)
// ===========================================================================
function saveProduct(token, data) {
  requireUser_(token, ['Warehouse']);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    data = data || {};
    const sheet = ss.getSheetByName('Products');
    const barcode = cleanCell_(data.barcode);
    if (!barcode) throw new Error('กรุณาระบุ Barcode');
    const name = safeText_(data.name), brand = safeText_(data.brand);
    if (!name) throw new Error('กรุณาระบุชื่อสินค้า');

    const pCtn = num_(data.packsPerCarton), uPack = num_(data.unitsPerPack), minStock = num_(data.minStock);
    if (pCtn < 0 || uPack < 0 || minStock < 0) throw new Error('ค่าตัวเลขต้องไม่ติดลบ');
    const status = String(data.status) === 'Inactive' ? 'Inactive' : 'Active';
    let imageUrl = String(data.imageUrl || '').trim();
    if (imageUrl && !/^https?:\/\//i.test(imageUrl)) throw new Error('ลิงก์รูปภาพต้องขึ้นต้นด้วย http:// หรือ https://');
    const allowed = parseAllowedUnits_(data.allowedUnits, pCtn, uPack).join(',');

    const values = sheet.getDataRange().getValues();
    let foundRow = -1;
    for (let i = 1; i < values.length; i++) {
      if (cleanCell_(values[i][0]) === barcode) { foundRow = i + 1; break; }
    }
    sheet.getRange('A:A').setNumberFormat('@');
    if (foundRow > 0) {
      sheet.getRange(foundRow, 2, 1, 8).setValues([[name, brand, pCtn, uPack, minStock, status, imageUrl, allowed]]);
    } else {
      sheet.appendRow([barcode, name, brand, pCtn, uPack, minStock, status, imageUrl, allowed]);
    }
    refreshStockSummary_();
    return JSON.stringify({ success: true });
  } finally {
    lock.releaseLock();
  }
}

// ===========================================================================
// LatestStock (สรุปยอดนับล่าสุดต่อ Row + สินค้า + MFD + EXP)
// ===========================================================================
function generateLatestStockSummary_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aSheet = ss.getSheetByName('Audits');
  if (!aSheet) return;
  const aData = aSheet.getDataRange().getValues();

  const latest = {};
  for (let i = 1; i < aData.length; i++) {
    const row = aData[i];
    const barcode = cleanCell_(row[3]);
    if (!barcode) continue;
    const loc = normalizeRowNameServer_(row[2]);
    const mfd = dateCell_(row[6]), exp = dateCell_(row[7]);
    const note = String(row[13] || '');
    let cycle = 'รอบแรก';
    const m = note.match(/นับครั้งที่\s*\d+/);
    if (m) cycle = m[0];
    else if (note.indexOf('เริ่มต้น') !== -1) cycle = 'นับครั้งที่ 1 (WH)';
    latest[loc + '|' + barcode + '|' + mfd + '|' + exp] = {
      barcode: barcode, brand: String(row[4] || '-'), name: String(row[5] || '-'), mfd: mfd, exp: exp,
      ctn: num_(row[8]), pck: num_(row[9]), pcs: num_(row[10]), loc: loc, cycle: cycle
    };
  }

  let sheet = ss.getSheetByName('LatestStock');
  if (!sheet) sheet = ss.insertSheet('LatestStock'); else sheet.clear();
  sheet.getRange('A:A').setNumberFormat('@');
  sheet.getRange('D:E').setNumberFormat('@');
  sheet.getRange(1, 1, 1, 9).setValues([['Barcode', 'Brand', 'รายการสินค้า', 'MFD', EXP_HEADER_, 'ลัง', 'แพ็ค', 'ชิ้น', 'แหล่งที่มา (Row & รอบนับ)']])
    .setFontWeight('bold').setBackground('#e2e8f0');
  const out = Object.keys(latest).map(k => {
    const it = latest[k];
    return [it.barcode, it.brand, it.name, it.mfd, it.exp, it.ctn, it.pck, it.pcs, 'Row: ' + it.loc + ' (' + it.cycle + ')'];
  });
  if (out.length) sheet.getRange(2, 1, out.length, 9).setValues(out);
}

// ===========================================================================
// ล้างข้อมูลการนับทั้งหมด (Warehouse) - สำรองไว้ใน AuditsArchive ก่อนลบเสมอ
// ===========================================================================
function clearAllAuditData(token) {
  requireUser_(token, ['Warehouse']);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const aSheet = ss.getSheetByName('Audits');
    const lastRow = aSheet.getLastRow();
    if (lastRow > 1) {
      const cols = aSheet.getLastColumn();
      const rows = aSheet.getRange(2, 1, lastRow - 1, cols).getValues();
      let arc = ss.getSheetByName('AuditsArchive');
      if (!arc) {
        arc = ss.insertSheet('AuditsArchive');
        arc.appendRow(['ArchivedAt'].concat(aSheet.getRange(1, 1, 1, cols).getValues()[0]));
        arc.getRange('C:E').setNumberFormat('@');
        arc.getRange('H:I').setNumberFormat('@');
      }
      const stamp = new Date();
      arc.getRange(arc.getLastRow() + 1, 1, rows.length, cols + 1).setValues(rows.map(r => [stamp].concat(r)));
      aSheet.deleteRows(2, lastRow - 1);
    }
    generateLatestStockSummary_();
    return JSON.stringify({ success: true });
  } finally {
    lock.releaseLock();
  }
}
