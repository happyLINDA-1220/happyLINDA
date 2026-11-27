/**
 * 金灶鍋貼｜線上點餐 × Google Sheets（Google Apps Script）
 *
 * 功能
 *  - doPost：接收網站送來的訂單 → 驗證 → 以「菜單」工作表的價格重新計算金額 → 寫入「訂單」「訂單明細」→ Email 通知店家
 *  - doGet ：健康檢查；?action=status&orderId=...&phone=... 可查詢訂單狀態
 *  - setup ：第一次使用時執行，自動建立工作表、標題列、菜單價格、狀態下拉選單與每日營收統計
 *
 * 安裝步驟請看同資料夾的 README.md
 */

// ====================== 設定 ======================
const CONFIG = {
  STORE_NAME: '金灶鍋貼',
  NOTIFY_EMAIL: '',             // 新訂單通知信要寄到哪裡；留空則寄給部署這支程式的 Google 帳號
  SEND_EMAIL: true,             // 不想收通知信改成 false
  ORDER_PREFIX: 'JZ',           // 訂單編號開頭，例如 JZ-20261127-001
  RATE_LIMIT_SECONDS: 30,       // 同一支電話多久內只能送一張單
  MAX_QTY_PER_ITEM: 99,
  TIMEZONE: 'Asia/Taipei',
  SHEETS: { ORDERS: '訂單', ITEMS: '訂單明細', MENU: '菜單', DAILY: '每日營收' },
  STATUSES: ['新訂單', '製作中', '可取餐', '已完成', '已取消'],
  ORDER_TYPES: ['內用', '外帶', '宅配']
};

const ORDER_HEADERS = ['訂單編號', '建立時間', '狀態', '取餐方式', '姓名', '電話', '宅配地址', '備註', '品項摘要', '數量合計', '總金額'];
const ITEM_HEADERS  = ['訂單編號', '建立時間', '品項ID', '品名', '單價', '數量', '小計'];
const MENU_HEADERS  = ['品項ID', '品名', '分類', '單價', '上架'];

// 與網站目前菜單一致的初始資料（之後請直接在「菜單」工作表改價格或下架）
const DEFAULT_MENU = [
  ['m1', '黃金招牌鍋貼', '招牌鍋貼', 120, true],
  ['m2', '爆汁韭菜鍋貼', '招牌鍋貼', 130, true],
  ['m3', '麻辣牛肉鍋貼', '招牌鍋貼', 140, true],
  ['m4', '鮮蝦高麗菜鍋貼', '招牌鍋貼', 150, true],
  ['m5', '經典豬肉水餃', '精選水餃', 110, true],
  ['m6', '高麗菜鮮蝦水餃', '精選水餃', 135, true],
  ['m7', '翡翠野菜水餃', '精選水餃', 120, true],
  ['m8', '酸辣湯', '湯品/飲料', 60, true],
  ['m9', '玉米濃湯', '湯品/飲料', 55, true],
  ['m10', '手炒酸梅湯', '湯品/飲料', 45, true],
  ['m11', '涼拌小黃瓜', '經典小菜', 50, true],
  ['m12', '古早味皮蛋豆腐', '經典小菜', 65, true],
  ['g1', '招財鍋貼禮盒 50入', '冷凍禮盒', 680, true],
  ['g2', '三味水餃綜合箱 90入', '冷凍禮盒', 890, true],
  ['g3', '節慶尊榮禮盒 120入', '冷凍禮盒', 1580, true]
];

// ====================== 第一次使用：建立工作表 ======================
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const orders = ensureSheet_(ss, CONFIG.SHEETS.ORDERS, ORDER_HEADERS);
  orders.getRange('B:B').setNumberFormat('yyyy/mm/dd hh:mm');
  orders.getRange('K:K').setNumberFormat('"NT$"#,##0');
  const statusRule = SpreadsheetApp.newDataValidation().requireValueInList(CONFIG.STATUSES, true).setAllowInvalid(false).build();
  orders.getRange(2, 3, orders.getMaxRows() - 1, 1).setDataValidation(statusRule);
  // 依狀態上色
  const colors = { '新訂單': '#FDE7D3', '製作中': '#FFF4C2', '可取餐': '#DFF2D8', '已完成': '#E8E8E8', '已取消': '#F6D5D5' };
  orders.setConditionalFormatRules(Object.keys(colors).map(function (s) {
    return SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$C2="' + s + '"').setBackground(colors[s])
      .setRanges([orders.getRange('A2:K')]).build();
  }));

  const items = ensureSheet_(ss, CONFIG.SHEETS.ITEMS, ITEM_HEADERS);
  items.getRange('B:B').setNumberFormat('yyyy/mm/dd hh:mm');
  items.getRange('E:E').setNumberFormat('"NT$"#,##0');
  items.getRange('G:G').setNumberFormat('"NT$"#,##0');

  const menu = ensureSheet_(ss, CONFIG.SHEETS.MENU, MENU_HEADERS);
  if (menu.getLastRow() < 2) {
    menu.getRange(2, 1, DEFAULT_MENU.length, MENU_HEADERS.length).setValues(DEFAULT_MENU);
  }
  menu.getRange('D:D').setNumberFormat('"NT$"#,##0');
  menu.getRange(2, 5, menu.getMaxRows() - 1, 1).insertCheckboxes();

  const daily = ensureSheet_(ss, CONFIG.SHEETS.DAILY, []);
  daily.clear();
  daily.getRange('A1').setFormula(
    '=QUERY(' + CONFIG.SHEETS.ORDERS + '!A:K,' +
    '"select toDate(B), count(A), sum(J), sum(K) where A is not null and C <> \'已取消\' ' +
    'group by toDate(B) order by toDate(B) desc ' +
    'label toDate(B) \'日期\', count(A) \'訂單數\', sum(J) \'份數\', sum(K) \'營收\'",1)'
  );
  daily.getRange('A:A').setNumberFormat('yyyy/mm/dd');
  daily.getRange('D:D').setNumberFormat('"NT$"#,##0');
  daily.setFrozenRows(1);

  SpreadsheetApp.flush();
  return '設定完成';
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (headers.length) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#4A2F23').setFontColor('#F7F1E6');
    sh.setFrozenRows(1);
  }
  return sh;
}

// ====================== 網站呼叫入口 ======================
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === 'status') return json_(getOrderStatus_(p.orderId, p.phone));
  return json_({ ok: true, service: CONFIG.STORE_NAME + ' 訂單 API' });
}

function doPost(e) {
  let data;
  try {
    data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json_({ ok: false, message: '資料格式錯誤' });
  }
  try {
    if (data.action !== 'createOrder') return json_({ ok: false, message: '不支援的操作' });
    return json_(createOrder_(data));
  } catch (err) {
    console.error(err);
    return json_({ ok: false, message: '系統忙碌中，請稍後再試或直接來電訂購' });
  }
}

// ====================== 建立訂單 ======================
function createOrder_(data) {
  // 機器人陷阱：真人看不到這個欄位，有填就默默丟掉
  if (data.website) return { ok: true, orderId: 'IGNORED', total: 0 };

  const name = clean_(data.name, 30);
  const phone = String(data.phone || '').replace(/[^\d+]/g, '');
  const orderType = String(data.orderType || '');
  const address = clean_(data.address, 120);
  const note = clean_(data.note, 200);

  if (!name) return { ok: false, message: '請填寫姓名' };
  if (!/^(\+?886|0)\d{8,9}$/.test(phone)) return { ok: false, message: '電話格式不正確' };
  if (CONFIG.ORDER_TYPES.indexOf(orderType) < 0) return { ok: false, message: '請選擇取餐方式' };
  if (orderType === '宅配' && !address) return { ok: false, message: '宅配請填寫收件地址' };
  if (!Array.isArray(data.items) || !data.items.length) return { ok: false, message: '購物車是空的' };

  // 防止重複送出
  const cache = CacheService.getScriptCache();
  const rateKey = 'rate_' + phone;
  if (cache.get(rateKey)) return { ok: false, message: '訂單已送出，請稍候 ' + CONFIG.RATE_LIMIT_SECONDS + ' 秒再下新訂單' };

  // 以「菜單」工作表為準重新計價，不相信網頁傳來的價格
  const menu = loadMenu_();
  const lines = [];
  for (let i = 0; i < data.items.length; i++) {
    const it = data.items[i];
    const m = menu[String(it.id)];
    const qty = Math.floor(Number(it.qty));
    if (!m) return { ok: false, message: '有品項已下架或不存在，請重新整理頁面' };
    if (!(qty >= 1 && qty <= CONFIG.MAX_QTY_PER_ITEM)) return { ok: false, message: '「' + m.name + '」數量不正確' };
    lines.push({ id: m.id, name: m.name, price: m.price, qty: qty, subtotal: m.price * qty });
  }
  const total = lines.reduce(function (s, l) { return s + l.subtotal; }, 0);
  const totalQty = lines.reduce(function (s, l) { return s + l.qty; }, 0);

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let orderId;
  const now = new Date();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const orders = ss.getSheetByName(CONFIG.SHEETS.ORDERS);
    const items = ss.getSheetByName(CONFIG.SHEETS.ITEMS);
    if (!orders || !items) throw new Error('請先執行 setup()');

    orderId = nextOrderId_(now);
    const summary = lines.map(function (l) { return l.name + ' x' + l.qty; }).join('、');
    orders.appendRow([orderId, now, '新訂單', orderType, name, "'" + phone, address, note, summary, totalQty, total]);
    items.getRange(items.getLastRow() + 1, 1, lines.length, ITEM_HEADERS.length).setValues(
      lines.map(function (l) { return [orderId, now, l.id, l.name, l.price, l.qty, l.subtotal]; })
    );
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  cache.put(rateKey, '1', CONFIG.RATE_LIMIT_SECONDS);

  try { notify_(orderId, now, orderType, name, phone, address, note, lines, total); }
  catch (err) { console.error('通知信寄送失敗', err); }

  return { ok: true, orderId: orderId, total: total };
}

function loadMenu_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.MENU);
  const rows = sh.getRange(2, 1, Math.max(0, sh.getLastRow() - 1), MENU_HEADERS.length).getValues();
  const menu = {};
  rows.forEach(function (r) {
    const id = String(r[0]).trim();
    const price = Number(r[3]);
    if (id && r[4] === true && price > 0) menu[id] = { id: id, name: String(r[1]), price: price };
  });
  return menu;
}

// 訂單編號：JZ-日期-當日流水號，例如 JZ-20261127-001
function nextOrderId_(now) {
  const props = PropertiesService.getScriptProperties();
  const day = Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyyyMMdd');
  const key = 'seq_' + day;
  const seq = Number(props.getProperty(key) || 0) + 1;
  props.setProperty(key, String(seq));
  return CONFIG.ORDER_PREFIX + '-' + day + '-' + ('00' + seq).slice(-3);
}

function notify_(orderId, now, orderType, name, phone, address, note, lines, total) {
  if (!CONFIG.SEND_EMAIL) return;
  const to = CONFIG.NOTIFY_EMAIL || Session.getEffectiveUser().getEmail();
  if (!to) return;
  const when = Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyyy/MM/dd HH:mm');
  const rows = lines.map(function (l) {
    return '<tr><td>' + esc_(l.name) + '</td><td align="right">' + l.qty + '</td><td align="right">NT$ ' + l.subtotal + '</td></tr>';
  }).join('');
  const html =
    '<h2 style="color:#4A2F23">新訂單 ' + orderId + '</h2>' +
    '<p>時間：' + when + '<br>取餐方式：' + esc_(orderType) + '<br>姓名：' + esc_(name) + '<br>電話：' + esc_(phone) +
    (address ? '<br>地址：' + esc_(address) : '') + (note ? '<br>備註：' + esc_(note) : '') + '</p>' +
    '<table cellpadding="6" style="border-collapse:collapse" border="1"><tr><th>品項</th><th>數量</th><th>小計</th></tr>' + rows +
    '<tr><td colspan="2" align="right"><b>總金額</b></td><td align="right"><b>NT$ ' + total + '</b></td></tr></table>' +
    '<p><a href="' + SpreadsheetApp.getActiveSpreadsheet().getUrl() + '">開啟訂單試算表</a></p>';
  MailApp.sendEmail({ to: to, subject: '【' + CONFIG.STORE_NAME + '】新訂單 ' + orderId + '｜' + orderType + '｜NT$ ' + total, htmlBody: html });
}

// ====================== 查詢訂單狀態 ======================
function getOrderStatus_(orderId, phone) {
  orderId = String(orderId || '').trim();
  phone = String(phone || '').replace(/[^\d+]/g, '');
  if (!orderId || !phone) return { ok: false, message: '請提供訂單編號與電話' };
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.ORDERS);
  const found = sh.getRange('A:A').createTextFinder(orderId).matchEntireCell(true).findNext();
  if (!found) return { ok: false, message: '查無此訂單' };
  const row = sh.getRange(found.getRow(), 1, 1, ORDER_HEADERS.length).getValues()[0];
  if (String(row[5]).replace(/[^\d+]/g, '') !== phone) return { ok: false, message: '查無此訂單' };
  return { ok: true, orderId: orderId, status: row[2], orderType: row[3], total: row[10] };
}

// ====================== 小工具 ======================
function clean_(v, max) {
  let s = String(v == null ? '' : v).replace(/[\u0000-\u001F]/g, ' ').trim().slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;   // 避免被試算表當成公式執行
  return s;
}

function esc_(s) {
  return String(s).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
