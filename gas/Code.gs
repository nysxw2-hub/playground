/**
 * 金月軒 線上訂單 — Google Apps Script（綁定在 Google 試算表上）
 *
 * 功能
 *  - doPost：網站送出訂單 → 驗證、以「商品」工作表的價格重新計價 → 寫入「訂單」「訂單明細」→ 寄通知信
 *  - doGet：健康檢查（網站可確認是否接單中）
 *  - onOpen：試算表上方的「金月軒」選單
 *  - onEdit：修改訂單狀態時自動記錄「狀態更新時間」
 *  - 每日統計：手動或排程產生「每日統計」工作表
 *
 * 部署步驟見 gas/README.md
 */

const TZ = 'Asia/Taipei';
const SHEETS = { orders: '訂單', items: '訂單明細', products: '商品', settings: '設定', daily: '每日統計' };
const STATUSES = ['新訂單', '製作中', '可取餐', '已完成', '已取消'];
const STATUS_COLORS = { 新訂單: '#fde7c8', 製作中: '#fff3b0', 可取餐: '#d7ecd2', 已完成: '#e5e5e5', 已取消: '#f4cccc' };
const FULFILLMENTS = ['外帶', '內用', '宅配'];
const ORDER_HEADERS = ['訂單編號', '建立時間', '狀態', '取餐方式', '姓名', '電話', '地址', '備註', '品項摘要', '數量', '金額', '狀態更新時間'];
const ITEM_HEADERS = ['訂單編號', '建立時間', '商品ID', '商品名稱', '單價', '數量', '小計'];
const PRODUCT_HEADERS = ['商品ID', '分類', '名稱', '單價', '單位', '上架', '可宅配'];
const DAILY_HEADERS = ['日期', '有效訂單', '營收', '外帶', '內用', '宅配', '取消', '熱銷品項'];
const COL = Object.fromEntries(ORDER_HEADERS.map((h, i) => [h, i + 1]));

// 與網站 PRODUCTS 相同的初始商品；之後以試算表「商品」工作表為準（可直接改價格、下架）
const SEED_PRODUCTS = [
  ['p1', '鍋貼', '經典鮮肉鍋貼', 95, '10 顆', true, false],
  ['p2', '鍋貼', '韭香鮮肉鍋貼', 105, '10 顆', true, false],
  ['p3', '鍋貼', '香辣椒麻鍋貼', 115, '10 顆', true, false],
  ['p4', '鍋貼', '玉米鮮肉鍋貼', 105, '10 顆', true, false],
  ['p5', '水餃', '手作鮮肉水餃', 90, '10 顆', true, false],
  ['p6', '水餃', '鮮蝦高麗菜水餃', 135, '10 顆', true, false],
  ['p7', '湯飲', '暖心酸辣湯', 55, '1 碗', true, false],
  ['p8', '湯飲', '清香冷泡茶', 45, '1 杯', true, false],
  ['p9', '小菜', '涼拌小黃瓜', 50, '1 份', true, false],
  ['p10', '小菜', '香滷豆干', 50, '1 份', true, false],
  ['gift', '禮盒', '招牌招財鍋貼禮盒', 599, '50 顆', true, true],
];

const MAX_QTY = 99;
const MAX_LINES = 30;
const RATE_LIMIT = { perPhone: 5, perAll: 120, windowSec: 600 };

/* ───────────── 網頁應用程式入口 ───────────── */

function doPost(e) {
  try {
    const raw = (e && e.postData && e.postData.contents) || '';
    if (raw.length > 20000) throw userError_('BAD_REQUEST', '資料過大，請重新整理後再試。');
    let body;
    try { body = JSON.parse(raw || '{}'); } catch (err) { throw userError_('BAD_REQUEST', '資料格式錯誤，請重新整理後再試。'); }
    return json_(createOrder_(body));
  } catch (err) {
    if (err && err.userMessage) return json_({ ok: false, code: err.code, error: err.userMessage });
    console.error(err && err.stack || err);
    return json_({ ok: false, code: 'SERVER_ERROR', error: '系統忙碌中，請稍後再試或來電訂購。' });
  }
}

function doGet() {
  const s = getSettings_();
  return json_({ ok: true, service: '金月軒訂單', accepting: s.accepting, message: s.accepting ? '' : s.pausedMessage });
}

/* ───────────── 建立訂單 ───────────── */

function createOrder_(body) {
  // 機器人陷阱欄位：真人看不到，有填就當作垃圾訂單
  if (str_(body.website)) throw userError_('REJECTED', '無法送出訂單。');

  const settings = getSettings_();
  if (!settings.accepting) throw userError_('PAUSED', settings.pausedMessage);

  const clientRef = str_(body.clientRef).slice(0, 64);
  const cache = CacheService.getScriptCache();
  if (clientRef) {
    const seen = cache.get('ref:' + clientRef);
    if (seen) return JSON.parse(seen); // 同一筆重送（網路重試）→ 回傳同一張訂單
  }

  const order = validate_(body);
  const priced = price_(order);

  rateLimit_(cache, order.phone);

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  let result;
  try {
    if (clientRef) {
      const seen = cache.get('ref:' + clientRef);
      if (seen) return JSON.parse(seen);
    }
    const now = new Date();
    const orderId = nextOrderId_(now);
    const summary = priced.lines.map(l => l.name + '×' + l.qty).join('、');
    const qty = priced.lines.reduce((n, l) => n + l.qty, 0);

    const ss = SpreadsheetApp.getActive();
    const orders = ss.getSheetByName(SHEETS.orders);
    orders.getRange(orders.getLastRow() + 1, 1, 1, ORDER_HEADERS.length).setValues([[
      orderId, now, '新訂單', order.fulfillment, safe_(order.name), safe_(order.phone), safe_(order.address),
      safe_(order.note), safe_(summary), qty, priced.total, now,
    ]]);
    const items = ss.getSheetByName(SHEETS.items);
    items.getRange(items.getLastRow() + 1, 1, priced.lines.length, ITEM_HEADERS.length).setValues(
      priced.lines.map(l => [orderId, now, l.id, safe_(l.name), l.price, l.qty, l.price * l.qty]));
    SpreadsheetApp.flush();

    result = {
      ok: true, orderId, fulfillment: order.fulfillment, total: priced.total,
      items: priced.lines.map(l => ({ id: l.id, name: l.name, qty: l.qty, price: l.price })),
      priceChanged: Number(body.total) !== priced.total,
    };
    if (clientRef) cache.put('ref:' + clientRef, JSON.stringify(result), 21600);
  } finally {
    lock.releaseLock();
  }

  notifyStore_(settings, result, order);
  return result;
}

function validate_(b) {
  const name = str_(b.customerName);
  const phone = str_(b.phone);
  const fulfillment = str_(b.fulfillment);
  const address = str_(b.address);
  const note = str_(b.note);
  if (!name || name.length > 50) throw userError_('INVALID', '請填寫聯絡姓名（50 字以內）。');
  if (!/^[0-9+() -]{8,20}$/.test(phone)) throw userError_('INVALID', '請填寫正確的聯絡電話。');
  if (FULFILLMENTS.indexOf(fulfillment) < 0) throw userError_('INVALID', '請選擇取餐方式。');
  if (fulfillment === '宅配' && !address) throw userError_('INVALID', '宅配請填寫地址。');
  if (address.length > 200) throw userError_('INVALID', '地址請在 200 字以內。');
  if (note.length > 300) throw userError_('INVALID', '備註請在 300 字以內。');

  if (!Array.isArray(b.items) || !b.items.length) throw userError_('EMPTY', '購物車是空的。');
  if (b.items.length > MAX_LINES) throw userError_('INVALID', '品項太多，請分開下單。');
  const merged = {};
  b.items.forEach(it => {
    const id = str_(it && it.id);
    const qty = Number(it && it.qty);
    if (!id || !Number.isInteger(qty) || qty < 1 || qty > MAX_QTY) throw userError_('INVALID', '商品數量有誤，請重新整理後再試。');
    merged[id] = (merged[id] || 0) + qty;
    if (merged[id] > MAX_QTY) throw userError_('INVALID', '單一商品最多 ' + MAX_QTY + ' 份。');
  });
  return { name, phone, fulfillment, address: fulfillment === '宅配' ? address : '', note, items: merged };
}

function price_(order) {
  const catalog = getProducts_();
  const lines = [];
  let total = 0;
  Object.keys(order.items).forEach(id => {
    const p = catalog[id];
    if (!p || !p.active) throw userError_('UNAVAILABLE', '部分商品已售完或下架，請移除後再送出。');
    if (order.fulfillment === '宅配' && !p.delivery) throw userError_('INVALID', '「' + p.name + '」不提供宅配，請改選外帶或內用。');
    const qty = order.items[id];
    lines.push({ id, name: p.name, price: p.price, qty });
    total += p.price * qty;
  });
  return { lines, total };
}

function nextOrderId_(now) {
  const day = Utilities.formatDate(now, TZ, 'yyMMdd');
  const props = PropertiesService.getScriptProperties();
  const n = Number(props.getProperty('seq:' + day) || 0) + 1;
  props.setProperty('seq:' + day, String(n));
  return 'GM' + day + '-' + ('00' + n).slice(-3);
}

function rateLimit_(cache, phone) {
  const bump = (key, max) => {
    const n = Number(cache.get(key) || 0) + 1;
    if (n > max) throw userError_('RATE_LIMIT', '短時間內送出太多筆訂單，請稍後再試或來電訂購。');
    cache.put(key, String(n), RATE_LIMIT.windowSec);
  };
  bump('rl:phone:' + phone.replace(/\D/g, ''), RATE_LIMIT.perPhone);
  bump('rl:all', RATE_LIMIT.perAll);
}

function notifyStore_(settings, r, order) {
  if (!settings.notifyEmail) return;
  try {
    const rows = r.items.map(i => '<tr><td>' + esc_(i.name) + '</td><td align="right">' + i.qty + '</td><td align="right">' + (i.price * i.qty) + '</td></tr>').join('');
    MailApp.sendEmail({
      to: settings.notifyEmail,
      subject: '【金月軒新訂單】' + r.orderId + '｜' + r.fulfillment + '｜NT$ ' + r.total,
      htmlBody:
        '<h2 style="margin:0 0 8px">' + esc_(r.orderId) + '</h2>' +
        '<p>取餐方式：' + esc_(r.fulfillment) + '<br>姓名：' + esc_(order.name) + '<br>電話：' + esc_(order.phone) +
        (order.address ? '<br>地址：' + esc_(order.address) : '') + (order.note ? '<br>備註：' + esc_(order.note) : '') + '</p>' +
        '<table border="1" cellpadding="6" style="border-collapse:collapse"><tr><th>品項</th><th>數量</th><th>小計</th></tr>' + rows +
        '<tr><td colspan="2" align="right"><b>合計</b></td><td align="right"><b>NT$ ' + r.total + '</b></td></tr></table>' +
        '<p><a href="' + SpreadsheetApp.getActive().getUrl() + '">開啟訂單試算表</a></p>',
    });
  } catch (err) {
    console.error('通知信寄送失敗', err);
  }
}

/* ───────────── 試算表：選單、狀態、初始化 ───────────── */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('金月軒')
    .addItem('初始化／修復工作表', 'setup')
    .addSeparator()
    .addItem('產生今日統計', 'reportToday')
    .addItem('產生昨日統計', 'reportYesterday')
    .addItem('安裝每晚自動統計（22:00）', 'installDailyTrigger')
    .addSeparator()
    .addItem('寄送測試通知信', 'sendTestEmail')
    .addItem('模擬一筆測試訂單', 'createTestOrder')
    .addToUi();
}

function onEdit(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== SHEETS.orders || e.range.getRow() < 2) return;
  if (e.range.getColumn() <= COL['狀態'] && e.range.getLastColumn() >= COL['狀態']) {
    sh.getRange(e.range.getRow(), COL['狀態更新時間'], e.range.getNumRows(), 1).setValue(new Date());
  }
}

function setup() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(TZ);

  const orders = sheet_(ss, SHEETS.orders, ORDER_HEADERS);
  orders.getRange('B:B').setNumberFormat('yyyy/mm/dd hh:mm');
  orders.getRange('L:L').setNumberFormat('yyyy/mm/dd hh:mm');
  orders.getRange('F:F').setNumberFormat('@');
  orders.getRange('K:K').setNumberFormat('"NT$ "#,##0');
  [110, 130, 80, 80, 100, 120, 220, 200, 320, 60, 90, 130].forEach((w, i) => orders.setColumnWidth(i + 1, w));
  const statusCol = orders.getRange(2, COL['狀態'], orders.getMaxRows() - 1, 1);
  statusCol.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).setAllowInvalid(false).build());
  orders.setConditionalFormatRules(STATUSES.map(s => SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$C2="' + s + '"').setBackground(STATUS_COLORS[s])
    .setRanges([orders.getRange(2, 1, orders.getMaxRows() - 1, ORDER_HEADERS.length)]).build()));

  const items = sheet_(ss, SHEETS.items, ITEM_HEADERS);
  items.getRange('B:B').setNumberFormat('yyyy/mm/dd hh:mm');
  items.getRange('E:E').setNumberFormat('#,##0');
  items.getRange('G:G').setNumberFormat('#,##0');

  const products = sheet_(ss, SHEETS.products, PRODUCT_HEADERS);
  if (products.getLastRow() < 2) products.getRange(2, 1, SEED_PRODUCTS.length, PRODUCT_HEADERS.length).setValues(SEED_PRODUCTS);
  products.getRange(2, 6, products.getMaxRows() - 1, 2).insertCheckboxes();

  const settings = sheet_(ss, SHEETS.settings, ['項目', '值', '說明']);
  if (settings.getLastRow() < 2) {
    settings.getRange(2, 1, 3, 3).setValues([
      ['接單中', true, '取消勾選即暫停網站接單'],
      ['通知信箱', Session.getEffectiveUser().getEmail(), '新訂單通知寄到這裡；多個信箱用逗號分隔，留空則不寄'],
      ['暫停接單訊息', '目前暫停線上接單，請稍後再試或來電訂購。', '暫停接單時顯示給顧客'],
    ]);
    settings.getRange('B2').insertCheckboxes();
  }
  settings.setColumnWidth(1, 120); settings.setColumnWidth(2, 320); settings.setColumnWidth(3, 360);

  const daily = sheet_(ss, SHEETS.daily, DAILY_HEADERS);
  daily.getRange('A:A').setNumberFormat('yyyy/mm/dd');
  daily.getRange('C:C').setNumberFormat('"NT$ "#,##0');
  daily.setColumnWidth(8, 360);

  const first = ss.getSheets()[0];
  if (first.getName() !== SHEETS.orders && first.getLastRow() === 0) ss.deleteSheet(first);
  ss.setActiveSheet(orders);
  ss.toast('工作表已就緒', '金月軒', 5);
}

/* ───────────── 每日統計 ───────────── */

function reportToday() { buildDailyReport_(new Date()); }
function reportYesterday() { buildDailyReport_(new Date(Date.now() - 864e5)); }
function dailyReportTrigger() { const r = buildDailyReport_(new Date()); emailDailyReport_(r); }

function installDailyTrigger() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'dailyReportTrigger').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('dailyReportTrigger').timeBased().atHour(22).everyDays(1).inTimezone(TZ).create();
  SpreadsheetApp.getActive().toast('每晚 22:00 會自動產生統計並寄到通知信箱', '金月軒', 6);
}

function buildDailyReport_(date) {
  const ss = SpreadsheetApp.getActive();
  const key = Utilities.formatDate(date, TZ, 'yyyy/MM/dd');
  const sameDay = d => d instanceof Date && Utilities.formatDate(d, TZ, 'yyyy/MM/dd') === key;

  const orders = ss.getSheetByName(SHEETS.orders).getDataRange().getValues().slice(1).filter(r => sameDay(r[1]));
  const valid = orders.filter(r => r[2] !== '已取消');
  const validIds = new Set(valid.map(r => r[0]));
  const by = f => valid.filter(r => r[3] === f).length;
  const revenue = valid.reduce((n, r) => n + Number(r[10] || 0), 0);

  const counts = {};
  ss.getSheetByName(SHEETS.items).getDataRange().getValues().slice(1)
    .filter(r => validIds.has(r[0]))
    .forEach(r => { counts[r[3]] = (counts[r[3]] || 0) + Number(r[5] || 0); });
  const top = Object.keys(counts).sort((a, b) => counts[b] - counts[a]).slice(0, 5).map(n => n + '×' + counts[n]).join('、');

  const row = [key, valid.length, revenue, by('外帶'), by('內用'), by('宅配'), orders.length - valid.length, top];
  const daily = ss.getSheetByName(SHEETS.daily);
  const dates = daily.getRange(1, 1, Math.max(daily.getLastRow(), 1), 1).getDisplayValues().map(r => r[0]);
  const at = dates.indexOf(key);
  daily.getRange(at > 0 ? at + 1 : daily.getLastRow() + 1, 1, 1, row.length).setValues([row]);
  return { date: key, orders: valid.length, revenue, cancelled: orders.length - valid.length, top };
}

function emailDailyReport_(r) {
  const to = getSettings_().notifyEmail;
  if (!to) return;
  MailApp.sendEmail(to, '【金月軒每日統計】' + r.date + '｜' + r.orders + ' 筆｜NT$ ' + r.revenue,
    r.date + '\n有效訂單：' + r.orders + ' 筆\n營收：NT$ ' + r.revenue + '\n取消：' + r.cancelled + ' 筆\n熱銷：' + (r.top || '—') +
    '\n\n' + SpreadsheetApp.getActive().getUrl());
}

/* ───────────── 測試工具 ───────────── */

function sendTestEmail() {
  const s = getSettings_();
  if (!s.notifyEmail) { SpreadsheetApp.getUi().alert('「設定」工作表的通知信箱是空的'); return; }
  MailApp.sendEmail(s.notifyEmail, '【金月軒】通知信測試', '收到這封信代表新訂單通知設定正確。');
  SpreadsheetApp.getActive().toast('已寄到 ' + s.notifyEmail, '金月軒', 5);
}

function createTestOrder() {
  const r = createOrder_({
    clientRef: 'test-' + Date.now(), fulfillment: '外帶', customerName: '測試顧客', phone: '0900000000',
    note: '這是測試訂單，可刪除', items: [{ id: 'p1', qty: 2 }, { id: 'p7', qty: 1 }],
  });
  SpreadsheetApp.getActive().toast('已建立 ' + r.orderId + '（NT$ ' + r.total + '）', '金月軒', 6);
}

/* ───────────── 小工具 ───────────── */

function getSettings_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.settings);
  const map = {};
  if (sh && sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(r => { map[String(r[0]).trim()] = r[1]; });
  return {
    accepting: map['接單中'] !== false && String(map['接單中']).toUpperCase() !== 'FALSE',
    notifyEmail: String(map['通知信箱'] || '').trim(),
    pausedMessage: String(map['暫停接單訊息'] || '目前暫停線上接單，請稍後再試。'),
  };
}

function getProducts_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.products);
  if (!sh || sh.getLastRow() < 2) throw new Error('商品工作表是空的，請先執行 setup');
  const out = {};
  sh.getRange(2, 1, sh.getLastRow() - 1, PRODUCT_HEADERS.length).getValues().forEach(r => {
    const id = String(r[0]).trim();
    const price = Number(r[3]);
    if (!id || !(price >= 0)) return;
    out[id] = { name: String(r[2]), price, active: r[5] === true, delivery: r[6] === true };
  });
  return out;
}

function sheet_(ss, name, headers) {
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#30231c').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  return sh;
}

function userError_(code, message) { const e = new Error(message); e.code = code; e.userMessage = message; return e; }
function str_(v) { return v == null ? '' : String(v).trim(); }
// 防止「=」「+」「-」「@」開頭的內容在試算表被當成公式執行
function safe_(v) { const s = str_(v); return /^[=+\-@]/.test(s) ? "'" + s : s; }
function esc_(s) { return String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
function json_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
