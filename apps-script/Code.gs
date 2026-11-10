/******************************************************************
 * NHÀ NGÔ – BACKEND ĐẶT BÁNH THEO ĐỢT (Google Apps Script)
 *
 * Cách cài:
 *  1. Mở file Google Sheet "Đơn đặt hàng" → Tiện ích mở rộng → Apps Script
 *  2. Sao lưu code cũ (copy ra chỗ khác), rồi dán đè toàn bộ file này
 *  3. Sửa dòng ADMIN_PASSWORD bên dưới thành mật khẩu quản lý của bạn
 *  4. Triển khai → Quản lý bản triển khai → ✏️ → Phiên bản: "Phiên bản mới" → Triển khai
 *     (giữ nguyên bản triển khai cũ để link web không đổi)
 *
 * Sheet được dùng / tự tạo:
 *  - SAN_PHAM : danh sách bánh (cột TỒN KHO = số lượng mặc định khi mở đợt mới)
 *  - DOT_DAT  : danh sách đợt đặt (admin tạo trên trang quản lý)
 *  - ĐỢT dd-mm-yyyy : mỗi đợt một sheet đơn hàng riêng, tự tạo khi tạo đợt
 *  - DON_HANG : đơn cũ trước khi có đợt (giữ nguyên, không xoá)
 ******************************************************************/

const SHEET_ID = '1tbwliFL7Q48u0OaXCAC_ZJRhdlM4b-i_2PlnL0OEEl8';
const ADMIN_PASSWORD = 'NHAP_MAT_KHAU_CUA_BAN';
const TZ = 'Asia/Ho_Chi_Minh';

const PRODUCT_SHEET = 'SAN_PHAM';
const BATCH_SHEET = 'DOT_DAT';
const LEGACY_ORDER_SHEET = 'DON_HANG';
const PRODUCT_HEADERS = ['ID', 'TÊN BÁNH', 'GIÁ', 'TỒN KHO', 'ĐANG BÁN', 'ẢNH URL', 'MÔ TẢ'];
const BATCH_HEADERS = ['MÃ ĐỢT', 'TÊN ĐỢT', 'NGÀY NHẬN BÁNH', 'HẠN CHỐT ĐƠN', 'TRẠNG THÁI', 'SỐ LƯỢNG MỞ BÁN', 'GHI CHÚ', 'TÊN SHEET', 'NGÀY TẠO'];
const ORDER_HEADERS = ['MÃ ĐƠN', 'THỜI GIAN', 'TÊN KHÁCH', 'SĐT', 'SẢN PHẨM', 'TỔNG SL', 'TỔNG TIỀN', 'THANH TOÁN', 'TRẠNG THÁI', 'GHI CHÚ', 'CHI TIẾT JSON', 'ẢNH THANH TOÁN'];
const ORDER_STATUSES = ['CHỜ XÁC NHẬN', 'ĐÃ NHẬN TIỀN', 'ĐANG LÀM', 'HOÀN THÀNH', 'HỦY'];
const BATCH_STATUSES = ['ĐANG MỞ', 'ĐÃ CHỐT', 'KẾT THÚC'];
const TOKEN_DAYS = 7;

/* ======================= ROUTING ======================= */

function doGet(e) {
  const p = (e && e.parameter) || {};
  let out;
  try { out = routeGet_(p); } catch (err) { out = { success: false, message: String(err && err.message || err) }; }
  const cb = p.callback;
  if (cb && /^[\w.$]+$/.test(cb)) {
    return ContentService.createTextOutput(cb + '(' + JSON.stringify(out) + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const p = (e && e.parameter) || {};
  let out;
  try { out = p.payload ? placeOrder_(JSON.parse(p.payload)) : routePost_(p); }
  catch (err) { out = { success: false, message: String(err && err.message || err) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function routeGet_(p) {
  switch (p.api) {
    case 'products': return legacyProducts_();             // trang cũ
    case 'shop': return shop_(p.batch);
    case 'lookup': return lookup_(p.batch, p.phone);
    case 'admin-challenge': return challenge_();
    case 'admin-login': return login_(p.nonce, p.proof);
    case 'admin-data': auth_(p.token); return adminData_(p.batch);
    default: return { success: false, message: 'API không hợp lệ' };
  }
}

function routePost_(p) {
  auth_(p.token);
  switch (p.api) {
    case 'admin-status': return setOrderStatus_(p.orderId, p.status, p.batch);
    case 'admin-save-product': return saveProduct_(JSON.parse(p.product || '{}'));
    case 'admin-delete-product': return deleteProduct_(p.productId);
    case 'admin-create-batch': return createBatch_(JSON.parse(p.batchData || '{}'));
    case 'admin-update-batch': return updateBatch_(JSON.parse(p.batchData || '{}'));
    default: return { success: false, message: 'API không hợp lệ' };
  }
}

/* ======================= HELPERS ======================= */

function ss_() { return SpreadsheetApp.openById(SHEET_ID); }

function sheet_(name, headers, index) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = index === undefined ? ss.insertSheet(name) : ss.insertSheet(name, index);
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#513725').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

function dataRows_(sh) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
}

function num_(v) {
  if (typeof v === 'number') return v;
  const s = String(v == null ? '' : v).replace(/[^\d-]/g, '');
  return s ? Number(s) : 0;
}

function fmt_(d, pattern) {
  if (d instanceof Date && !isNaN(d)) return Utilities.formatDate(d, TZ, pattern);
  return d == null ? '' : String(d);
}

function digits_(s) { return String(s == null ? '' : s).replace(/\D/g, '').replace(/^0+/, ''); }

function toDate_(v) {
  if (v instanceof Date) return v;
  if (!v) return null;
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2}))?$/);
  if (m) return Utilities.parseDate(m[1] + '-' + m[2] + '-' + m[3] + ' ' + (m[4] || '00') + ':' + (m[5] || '00'), TZ, 'yyyy-MM-dd HH:mm');
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?: (\d{1,2}):(\d{2}))?/);
  if (m) return Utilities.parseDate(m[3] + '-' + m[2] + '-' + m[1] + ' ' + (m[4] || '00') + ':' + (m[5] || '00'), TZ, 'yyyy-M-d HH:mm');
  return null;
}

function sha256Hex_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)
    .map(function (b) { return ((b < 0 ? b + 256 : b).toString(16)).padStart(2, '0'); }).join('');
}

function saveImage_(dataUrl, fileName, propKey, folderName, share) {
  const m = String(dataUrl || '').match(/^data:(image\/[\w+.-]+);base64,(.+)$/);
  if (!m) return '';
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], fileName);
  const props = PropertiesService.getScriptProperties();
  let folder = null;
  const id = props.getProperty(propKey);
  if (id) { try { folder = DriveApp.getFolderById(id); } catch (e) { folder = null; } }
  if (!folder) { folder = DriveApp.createFolder(folderName); props.setProperty(propKey, folder.getId()); }
  const file = folder.createFile(blob);
  if (share) file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file;
}

/* ======================= DATA ======================= */

function getProducts_() {
  const sh = sheet_(PRODUCT_SHEET, PRODUCT_HEADERS);
  return dataRows_(sh).map(function (r, i) {
    return {
      row: i + 2, id: String(r[0]).trim(), name: String(r[1]).trim(), price: num_(r[2]), stock: num_(r[3]),
      active: r[4] === true || String(r[4]).toUpperCase() === 'TRUE',
      imageUrl: String(r[5] || ''), desc: String(r[6] || '')
    };
  }).filter(function (p) { return p.id; });
}

function parseLimits_(s) {
  const o = {};
  String(s || '').split(/[;\n]/).forEach(function (part) {
    const kv = part.split('=');
    if (kv.length >= 2 && kv[0].trim()) o[kv[0].trim()] = num_(kv[1]);
  });
  return o;
}

function limitsText_(o) {
  return Object.keys(o || {}).filter(function (k) { return k; }).map(function (k) { return k + '=' + Math.max(0, num_(o[k])); }).join('; ');
}

function getBatches_() {
  const sh = sheet_(BATCH_SHEET, BATCH_HEADERS);
  return dataRows_(sh).map(function (r, i) {
    return {
      row: i + 2, code: String(r[0]).trim(), name: String(r[1] || ''), pickupDate: toDate_(r[2]), deadline: toDate_(r[3]),
      status: BATCH_STATUSES.indexOf(String(r[4]).trim()) >= 0 ? String(r[4]).trim() : 'ĐANG MỞ',
      limits: parseLimits_(r[5]), note: String(r[6] || ''), sheetName: String(r[7] || r[0]).trim(), created: r[8]
    };
  }).filter(function (b) { return b.code; });
}

function pickupKey_(b) { return b.pickupDate ? b.pickupDate.getTime() : 0; }

function isOpen_(b) {
  if (!b || b.status !== 'ĐANG MỞ') return false;
  if (b.deadline && new Date() > b.deadline) return false;
  return true;
}

function defaultBatch_(batches) {
  const asc = batches.slice().sort(function (a, b) { return pickupKey_(a) - pickupKey_(b); });
  return asc.filter(isOpen_)[0] || asc.filter(function (b) { return b.status !== 'KẾT THÚC'; }).pop() || asc.pop() || null;
}

function batchPublic_(b) {
  return {
    code: b.code, name: b.name, status: b.status, open: isOpen_(b), note: b.note,
    pickupDate: fmt_(b.pickupDate, 'yyyy-MM-dd'), pickupText: fmt_(b.pickupDate, 'dd/MM/yyyy'),
    deadline: b.deadline ? b.deadline.getTime() : null, deadlineText: b.deadline ? fmt_(b.deadline, 'HH:mm dd/MM/yyyy') : '',
    deadlineInput: b.deadline ? fmt_(b.deadline, "yyyy-MM-dd'T'HH:mm") : ''
  };
}

function orderSheet_(b) {
  const ss = ss_();
  let sh = ss.getSheetByName(b.sheetName);
  if (!sh) {
    sh = sheet_(b.sheetName, ORDER_HEADERS, ss.getSheets().length);
    sh.getRange('B:B').setNumberFormat('dd/MM/yyyy HH:mm:ss');
    sh.getRange('D:D').setNumberFormat('@');
    sh.getRange('G:G').setNumberFormat('#,##0" đ"');
  }
  return sh;
}

function orderObj_(r, i) {
  let detail = [];
  try { detail = JSON.parse(r[10] || '[]'); } catch (e) { detail = []; }
  return {
    row: i + 2, id: String(r[0]).trim(), time: r[1], timeText: fmt_(r[1], 'HH:mm dd/MM/yyyy'),
    name: String(r[2] || ''), phone: String(r[3] || ''), items: String(r[4] || ''), qty: num_(r[5]), total: num_(r[6]),
    payment: String(r[7] || ''), status: String(r[8] || 'CHỜ XÁC NHẬN').trim(), note: String(r[9] || ''),
    detail: detail, paymentImage: String(r[11] || '')
  };
}

function getOrders_(b) {
  const sh = ss_().getSheetByName(b ? b.sheetName : LEGACY_ORDER_SHEET);
  if (!sh) return [];
  return dataRows_(sh).map(orderObj_).filter(function (o) { return o.id; });
}

function remaining_(b, products, orders) {
  const used = {};
  orders.forEach(function (o) {
    if (o.status === 'HỦY') return;
    o.detail.forEach(function (it) { used[it.id] = (used[it.id] || 0) + num_(it.qty); });
  });
  return products.filter(function (p) { return b.limits[p.id] !== undefined; }).map(function (p) {
    const limit = num_(b.limits[p.id]), sold = used[p.id] || 0;
    return { id: p.id, name: p.name, price: p.price, imageUrl: p.imageUrl, desc: p.desc, active: p.active, limit: limit, sold: sold, stock: Math.max(0, limit - sold) };
  });
}

function maskName_(n) {
  const w = String(n || '').trim().split(/\s+/);
  return w[w.length - 1] || 'Khách';
}

function orderPublic_(o) {
  return {
    id: o.id, name: maskName_(o.name), phoneTail: digits_(o.phone).slice(-3),
    items: o.detail.map(function (it) { return { id: it.id, name: it.name, qty: it.qty }; }),
    itemsText: o.items, qty: o.qty, total: o.total, status: o.status, timeText: o.timeText
  };
}

/* ======================= KHÁCH ======================= */

function legacyProducts_() {
  const products = getProducts_();
  const b = getBatches_().filter(isOpen_).sort(function (a, c) { return pickupKey_(a) - pickupKey_(c); })[0];
  if (!b) return products.filter(function (p) { return p.active; });
  return remaining_(b, products, getOrders_(b)).filter(function (p) { return p.active; });
}

function shop_(code) {
  const all = getBatches_();
  let visible = all.filter(function (b) { return b.status !== 'KẾT THÚC'; });
  const b = all.filter(function (x) { return x.code === code; })[0] || defaultBatch_(visible.length ? visible : all);
  if (!b) return { success: true, batches: [], batch: null, products: [], orders: [] };
  if (visible.indexOf(b) < 0) visible.push(b);
  visible.sort(function (a, c) { return pickupKey_(a) - pickupKey_(c); });
  const orders = getOrders_(b);
  return {
    success: true,
    batches: visible.map(batchPublic_),
    batch: batchPublic_(b),
    products: remaining_(b, getProducts_(), orders).filter(function (p) { return p.active || p.sold; }),
    orders: orders.map(orderPublic_)
  };
}

function lookup_(code, phone) {
  const ph = digits_(phone);
  if (ph.length < 8) return { success: false, message: 'Số điện thoại chưa đúng' };
  const b = getBatches_().filter(function (x) { return x.code === code; })[0];
  if (!b) return { success: false, message: 'Không tìm thấy đợt' };
  const ids = getOrders_(b).filter(function (o) { return digits_(o.phone) === ph; }).map(function (o) { return o.id; });
  return { success: true, ids: ids };
}

function placeOrder_(d) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const batches = getBatches_();
    const b = d.batchCode
      ? batches.filter(function (x) { return x.code === d.batchCode; })[0]
      : batches.filter(isOpen_).sort(function (a, c) { return pickupKey_(a) - pickupKey_(c); })[0];
    const products = getProducts_();
    const orders = getOrders_(b || null);
    const code = String(d.orderCode || ('NG' + Date.now())).trim();
    if (orders.some(function (o) { return o.id === code; })) return { success: true, duplicate: true, orderCode: code };

    const avail = {};
    if (b) remaining_(b, products, orders).forEach(function (x) { avail[x.id] = x.stock; });
    const items = [], warn = [];
    (d.items || []).forEach(function (it) {
      const p = products.filter(function (x) { return x.id === it.id; })[0];
      const q = Math.max(0, Math.floor(num_(it.qty)));
      if (!p || !q) return;
      if (b ? (avail[p.id] === undefined || avail[p.id] < q) : p.stock < q) warn.push(p.name);
      items.push({ id: p.id, name: p.name, qty: q, price: p.price, subtotal: q * p.price });
    });
    if (!items.length) throw new Error('Đơn chưa có bánh');
    if (b && !isOpen_(b)) warn.push('đợt đã chốt');

    const file = d.paymentImage ? saveImage_(d.paymentImage, code + '-thanh-toan.jpg', 'PAYMENT_FOLDER_ID', 'NHÀ NGÔ - ẢNH THANH TOÁN', false) : null;
    const total = items.reduce(function (s, x) { return s + x.subtotal; }, 0);
    const qty = items.reduce(function (s, x) { return s + x.qty; }, 0);
    let note = String(d.note || '').trim();
    if (warn.length) note = '[CẦN KIỂM TRA – vượt số lượng: ' + warn.join(', ') + '] ' + note;

    const sh = b ? orderSheet_(b) : sheet_(LEGACY_ORDER_SHEET, ORDER_HEADERS);
    sh.appendRow([
      code, new Date(), String(d.name || '').trim(), (b ? '' : "'") + String(d.phone || '').trim(),
      items.map(function (x) { return x.name + ' x' + x.qty; }).join(', '), qty, total,
      file ? 'Khách báo đã chuyển' : 'Chưa có ảnh', 'CHỜ XÁC NHẬN', note, JSON.stringify(items), file ? file.getUrl() : ''
    ]);

    if (!b) { // chế độ cũ (chưa có đợt): trừ tồn kho trực tiếp
      const psh = sheet_(PRODUCT_SHEET, PRODUCT_HEADERS);
      items.forEach(function (x) {
        const p = products.filter(function (y) { return y.id === x.id; })[0];
        psh.getRange(p.row, 4).setValue(Math.max(0, p.stock - x.qty));
      });
    }
    return { success: true, orderCode: code, batch: b ? b.code : '' };
  } finally {
    lock.releaseLock();
  }
}

/* ======================= ADMIN ======================= */

function challenge_() {
  const nonce = Utilities.getUuid();
  CacheService.getScriptCache().put('nonce_' + nonce, '1', 300);
  return { success: true, nonce: nonce };
}

function adminPassword_() {
  return PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD') || ADMIN_PASSWORD;
}

function login_(nonce, proof) {
  const cache = CacheService.getScriptCache();
  if (!nonce || !cache.get('nonce_' + nonce)) return { success: false, message: 'Phiên đăng nhập hết hạn, thử lại.' };
  cache.remove('nonce_' + nonce);
  if (sha256Hex_(adminPassword_() + ':' + nonce) !== String(proof || '')) return { success: false, message: 'Sai mật khẩu' };
  const props = PropertiesService.getScriptProperties(), now = Date.now(), all = props.getProperties();
  Object.keys(all).forEach(function (k) { if (k.indexOf('tok_') === 0 && Number(all[k]) < now) props.deleteProperty(k); });
  const token = Utilities.getUuid();
  props.setProperty('tok_' + token, String(now + TOKEN_DAYS * 864e5));
  return { success: true, token: token, data: adminData_() };
}

function auth_(token) {
  const exp = token && PropertiesService.getScriptProperties().getProperty('tok_' + token);
  if (!exp || Number(exp) < Date.now()) throw new Error('Phiên đăng nhập đã hết.');
}

function adminData_(code) {
  const batches = getBatches_(), products = getProducts_();
  const b = batches.filter(function (x) { return x.code === code; })[0]
    || defaultBatch_(batches.filter(function (x) { return x.status !== 'KẾT THÚC'; }))
    || defaultBatch_(batches);
  const orders = getOrders_(b || null);
  const sorted = batches.slice().sort(function (a, c) { return pickupKey_(c) - pickupKey_(a); });
  return {
    success: true,
    batches: sorted.map(function (x) { const o = batchPublic_(x); o.limits = x.limits; o.sheetName = x.sheetName; return o; }),
    batch: b ? (function () { const o = batchPublic_(b); o.limits = b.limits; o.sheetName = b.sheetName; return o; })() : null,
    products: products.map(function (p) { return { id: p.id, name: p.name, price: p.price, stock: p.stock, active: p.active, imageUrl: p.imageUrl, desc: p.desc }; }),
    orders: orders.map(function (o) {
      return { id: o.id, timeText: o.timeText, name: o.name, phone: o.phone, items: o.items, qty: o.qty, total: o.total, payment: o.payment, status: o.status, note: o.note, paymentImage: o.paymentImage, detail: o.detail };
    }).reverse(),
    summary: b ? remaining_(b, products, orders) : []
  };
}

function setOrderStatus_(orderId, status, code) {
  if (ORDER_STATUSES.indexOf(status) < 0) throw new Error('Trạng thái không hợp lệ');
  const batches = getBatches_();
  const sheets = [];
  const first = batches.filter(function (x) { return x.code === code; })[0];
  if (first) sheets.push(first.sheetName);
  batches.forEach(function (x) { if (sheets.indexOf(x.sheetName) < 0) sheets.push(x.sheetName); });
  sheets.push(LEGACY_ORDER_SHEET);
  const ss = ss_();
  for (let i = 0; i < sheets.length; i++) {
    const sh = ss.getSheetByName(sheets[i]);
    if (!sh) continue;
    const rows = dataRows_(sh);
    for (let r = 0; r < rows.length; r++) {
      if (String(rows[r][0]).trim() === String(orderId).trim()) {
        sh.getRange(r + 2, 9).setValue(status);
        if (status === 'ĐÃ NHẬN TIỀN') sh.getRange(r + 2, 8).setValue('Đã xác nhận tiền');
        return { success: true };
      }
    }
  }
  throw new Error('Không tìm thấy đơn ' + orderId);
}

function saveProduct_(p) {
  const id = String(p.id || '').trim();
  if (!id || !String(p.name || '').trim()) throw new Error('Thiếu mã hoặc tên bánh');
  const sh = sheet_(PRODUCT_SHEET, PRODUCT_HEADERS);
  const existing = getProducts_().filter(function (x) { return x.id === id; })[0];
  let imageUrl = existing ? existing.imageUrl : '';
  if (p.imageData) {
    const f = saveImage_(p.imageData, id + '.jpg', 'PRODUCT_FOLDER_ID', 'NHÀ NGÔ - ẢNH SẢN PHẨM', true);
    if (f) imageUrl = 'https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w1200';
  }
  const row = [id, String(p.name).trim(), num_(p.price), num_(p.stock), p.active !== false, imageUrl, p.desc !== undefined ? String(p.desc) : (existing ? existing.desc : '')];
  if (existing) sh.getRange(existing.row, 1, 1, row.length).setValues([row]);
  else sh.appendRow(row);
  return { success: true };
}

function deleteProduct_(id) {
  const p = getProducts_().filter(function (x) { return x.id === String(id).trim(); })[0];
  if (!p) throw new Error('Không tìm thấy sản phẩm');
  sheet_(PRODUCT_SHEET, PRODUCT_HEADERS).deleteRow(p.row);
  return { success: true };
}

function createBatch_(d) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const pickup = toDate_(d.pickupDate);
    if (!pickup) throw new Error('Chưa chọn ngày nhận bánh');
    const batches = getBatches_();
    const base = 'D' + fmt_(pickup, 'ddMMyy');
    let code = /^[A-Z0-9-]{3,20}$/.test(String(d.code || '')) ? String(d.code) : base, n = 2;
    while (batches.some(function (b) { return b.code === code; })) code = base + '-' + (n++);
    const ss = ss_();
    let sheetName = 'ĐỢT ' + fmt_(pickup, 'dd-MM-yyyy') + (code === base ? '' : ' (' + code.split('-').pop() + ')');
    while (ss.getSheetByName(sheetName)) sheetName += '+';

    const limits = {};
    if (d.limits && typeof d.limits === 'object') Object.keys(d.limits).forEach(function (k) { if (num_(d.limits[k]) > 0) limits[k] = num_(d.limits[k]); });
    else getProducts_().forEach(function (p) { if (p.active && p.stock > 0) limits[p.id] = p.stock; });

    const name = String(d.name || '').trim() || ('Đợt ' + fmt_(pickup, 'dd/MM'));
    const sh = sheet_(BATCH_SHEET, BATCH_HEADERS);
    sh.appendRow([code, name, pickup, toDate_(d.deadline) || '', 'ĐANG MỞ', limitsText_(limits), String(d.note || ''), sheetName, new Date()]);
    const last = sh.getLastRow();
    sh.getRange(last, 3).setNumberFormat('dd/MM/yyyy');
    sh.getRange(last, 4).setNumberFormat('HH:mm dd/MM/yyyy');
    sh.getRange(last, 9).setNumberFormat('dd/MM/yyyy HH:mm');
    orderSheet_({ sheetName: sheetName });

    if (d.closeOthers) {
      batches.forEach(function (b) { if (b.status === 'ĐANG MỞ') sh.getRange(b.row, 5).setValue('ĐÃ CHỐT'); });
    }
    return { success: true, code: code, sheetName: sheetName };
  } finally {
    lock.releaseLock();
  }
}

function updateBatch_(d) {
  const b = getBatches_().filter(function (x) { return x.code === d.code; })[0];
  if (!b) throw new Error('Không tìm thấy đợt');
  const sh = sheet_(BATCH_SHEET, BATCH_HEADERS);
  if (d.name !== undefined) sh.getRange(b.row, 2).setValue(String(d.name));
  if (d.pickupDate) { const p = toDate_(d.pickupDate); if (p) sh.getRange(b.row, 3).setValue(p).setNumberFormat('dd/MM/yyyy'); }
  if (d.deadline !== undefined) sh.getRange(b.row, 4).setValue(toDate_(d.deadline) || '').setNumberFormat('HH:mm dd/MM/yyyy');
  if (d.status) {
    if (BATCH_STATUSES.indexOf(d.status) < 0) throw new Error('Trạng thái đợt không hợp lệ');
    sh.getRange(b.row, 5).setValue(d.status);
  }
  if (d.limits && typeof d.limits === 'object') sh.getRange(b.row, 6).setValue(limitsText_(d.limits));
  if (d.note !== undefined) sh.getRange(b.row, 7).setValue(String(d.note));
  return { success: true };
}

/* Chạy hàm này 1 lần trong trình soạn thảo để cấp quyền & tạo sheet DOT_DAT */
function caiDat() {
  sheet_(BATCH_SHEET, BATCH_HEADERS);
  Logger.log('OK – đã có sheet ' + BATCH_SHEET);
}
