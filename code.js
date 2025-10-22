/** ================== 基本設定 ================== */
var SHEET_CFG = {
  spreadsheetId: '1qgECWIRQvcYCIpzhpHc2RJDfgfj-PT0Ba8oxQYkEreg', // 你的 Google Sheet ID
  sheetName: 'collected',                                       // 分頁名稱
  headerRow: 1,                                                 // 標題列（資料從下一列開始）
  COL: { DATE:1, TITLE:2, CAT:3, AMT:4, TYPE:5, PAYER:6 }       // A..F
};

/** ================== 路由 ================== */
function doGet() {
  var t = HtmlService.createTemplateFromFile('form'); // 這個檔名要和你的主頁 form.html 對上
  t.DEPLOY_TAG = new Date().toISOString();           // 給你頁面裡的 <?= DEPLOY_TAG ?> 用
  return t.evaluate()
    .setTitle('碧柳記帳冊 v10')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** ================== 打開 Sheet ================== */
function _sheet(){
  var ss = SpreadsheetApp.openById(SHEET_CFG.spreadsheetId);
  var sh = ss.getSheetByName(SHEET_CFG.sheetName);
  if (!sh) throw new Error('找不到工作表：' + SHEET_CFG.sheetName);
  return sh;
}

/** ================== 服務 ================== */
function svcInfo() {
  var sh = _sheet();
  return {
    ok: true,
    version: 'v10',
    spreadsheetUrl: sh.getParent().getUrl(),
    sheetName: sh.getName(),
    lastRow: sh.getLastRow(),
    tz: Session.getScriptTimeZone() || 'Asia/Taipei'
  };
}

/** 下拉選單來源（類別 / 付款人） */
function readOptions(){
  var sh = _sheet();
  var rng = sh.getDataRange().getValues();
  var head = rng.shift() || [];
  var idx = {
    date: head.indexOf('日期'), title: head.indexOf('項目名稱'), category: head.indexOf('項目類別'),
    amount: head.indexOf('金額'), type: head.indexOf('收入/支出'), payer: head.indexOf('付款人')
  };
  var derivedCats = {};
  var payers = {};
  rng.forEach(function(row){
    var cat = (row[idx.category]||'').toString().trim();
    var py  = (row[idx.payer]||'').toString().trim();
    if (cat) derivedCats[cat] = true;
    if (py)  payers[py] = true;
  });

  // Script Properties 中的類別，與資料中出現過的類別合併
  var stored = _getCategories(); // [{name, ioType}]
  var catSet = {};
  stored.forEach(function(it){ if (it && it.name) catSet[it.name] = true; });
  Object.keys(derivedCats).forEach(function(k){ catSet[k] = true; });

  var categories = Object.keys(catSet).sort();
  var payList    = Object.keys(payers).sort();
  return { categories: categories, payers: payList };
}

/** 新增一筆資料（表單） */
function writeEntry(payload) {
  var row = _normalize(payload);
  var sh  = _sheet();
  sh.appendRow(row);
  SpreadsheetApp.flush(); // 確保立即寫入
  return { ok: true, lastRow: sh.getLastRow() };
}

/** KPI：總收入 / 總支出 / 目前餘額 */
function readTotals(){
  var sh = _sheet();
  var rng = sh.getDataRange().getValues();
  var head = rng.shift() || [];
  var idx = {
    category: head.indexOf('項目類別'), amount: head.indexOf('金額'), type: head.indexOf('收入/支出')
  };
  var inc=0, exp=0;
  rng.forEach(function(row){
    var n = Number(row[idx.amount]||0) || 0;
    var t = (row[idx.type]||'').toString().trim();
    if (t === '收入') inc += n;
    else if (t === '支出') exp += n;
    // t === '中性' → 不計損益
  });
  return { totalIncome: inc, totalExpense: exp, balance: inc - exp };
}
/** ================== 程式頁：自訂指令（儲存於試算表 code 分頁） ================== */
function _codeSheet_(){
  var ss = SpreadsheetApp.openById(SHEET_CFG.spreadsheetId);
  var sh = ss.getSheetByName('code');
  if (!sh){
    sh = ss.insertSheet('code');
  }
  // 確保表頭
  if (sh.getLastRow() === 0){
    sh.getRange(1,1,1,4).setValues([["時間","標題","說明","指令"]]);
  } else {
    var headerRow = sh.getRange(1,1,1,4).getValues();
    var emptyHead = true;
    for (var i=0;i<4;i++){ if (String(headerRow[0][i]||'').trim() !== '') { emptyHead=false; break; } }
    if (emptyHead) sh.getRange(1,1,1,4).setValues([["時間","標題","說明","指令"]]);
  }
  return sh;
}

function listCommands(){
  var sh = _codeSheet_();
  var lr = sh.getLastRow();
  if (lr <= 1) return [];
  var values = sh.getRange(2,1,lr-1,4).getValues();
  var out = [];
  for (var i=0;i<values.length;i++){
    out.push({
      title: String(values[i][1] || '').trim(),
      note:  String(values[i][2] || '').trim(),
      body:  String(values[i][3] || '').trim()
    });
  }
  return out;
}

function addCommand(obj){
  if (!obj) throw new Error('缺少參數');
  var title = String(obj.title || '').trim();
  var note  = String(obj.note  || '').trim();
  var body  = String(obj.body  || '').trim();
  if (!title) throw new Error('標題不可空白');
  if (!body)  throw new Error('指令不可空白');
  if (title.length > 100) title = title.slice(0,100);
  if (note.length  > 200) note  = note.slice(0,200);
  if (body.length  > 4000) body = body.slice(0,4000);

  var sh = _codeSheet_();
  sh.appendRow([new Date(), title, note, body]);
  SpreadsheetApp.flush();
  return { ok:true };
}

/** 分析頁：詳細清單（最新 N 筆，預設 1000）— 以字串為主，避免型別踩雷 */
function readRowsLatest(limit) {
  limit = Math.max(1, Math.min(1000, Number(limit || 1000)));

  var sh = _sheet();
  var lr = sh.getLastRow();
  if (lr <= SHEET_CFG.headerRow) return { total: 0, rows: [] };

  var start = SHEET_CFG.headerRow + 1;
  var rowsN = lr - SHEET_CFG.headerRow;

  // 用 displayValues：全部以「字串」讀出（含日期、金額）
  var rng = sh.getRange(start, 1, rowsN, 6).getDisplayValues(); // A..F 皆為 string
  var tz = Session.getScriptTimeZone() || 'Asia/Taipei';

  // 轉成統一的物件陣列；dateStr 就用表內顯示字串，amount 先保留字串，給前端轉數字
  var all = rng.map(function(a){
    return {
      dateStr: String(a[0] || '').trim(),   // 例如 2025/7/18 或 2025-07-18
      title:   String(a[1] || '').trim(),
      category:String(a[2] || '').trim(),
      amount:  String(a[3] || '').trim(),   // 可能含千分位
      type:    String(a[4] || '').trim(),
      payer:   String(a[5] || '').trim()
    };
  });

  // 依日期字串嘗試轉為可比較的時間戳，排 DESC；無法解析放最下面
  function toEpoch(s){
    // 支援 yyyy/MM/dd 或 yyyy-MM-dd
    var m = s.match(/^(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})$/);
    if (m) return new Date(Number(m[1]), Number(m[2])-1, Number(m[3])).getTime();
    var d = new Date(s); return isNaN(d) ? -1 : d.getTime();
  }
  all.sort(function(a,b){ return toEpoch(b.dateStr) - toEpoch(a.dateStr); });

  // 取前 N 筆
  var slice = all.slice(0, limit);

  return { total: all.length, rows: slice };
}

/** （可選）極小除錯：最後列數 + 樣本 */
function readRowsCount(){
  var sh = _sheet();
  var lr = sh.getLastRow();
  var n  = Math.max(0, lr - SHEET_CFG.headerRow);
  var sample = n > 0 ? sh.getRange(SHEET_CFG.headerRow + 1, 1, Math.min(3, n), 6).getValues() : [];
  return { ok:true, lastRow: lr, dataRows: n, sample: sample };
}

/** ================== 私有工具 ================== */
function _num(v){
  var s = (v === undefined || v === null) ? '' : String(v);
  s = s.replace(/,/g,'').trim();
  var n = Number(s);
  return isNaN(n) ? 0 : n;
}

function _parseDate_(raw){
  if (raw instanceof Date) return raw;
  var s = String(raw || '').trim();
  var m = s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})$/);
  if (m) return new Date(Number(m[1]), Number(m[2])-1, Number(m[3]));
  var d = new Date(s);
  return isNaN(d) ? new Date() : d;
}

function _normalize(p) {
  if (!p) throw new Error('沒有資料');

  // 日期
  var date = (function(raw){
    if (!raw) throw new Error('請選擇日期');
    if (raw instanceof Date) return raw;
    var s = String(raw).trim();
    var m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/) || s.match(/^(\d{4})[\/.](\d{1,2})[\/.](\d{1,2})$/);
    if (m) return new Date(Number(m[1]), Number(m[2])-1, Number(m[3]));
    var d = new Date(s);
    if (isNaN(d)) throw new Error('日期格式不正確');
    return d;
  })(p.date);

  // 金額
  var amount = (function(raw){
    var n = _num(raw);
    if (!isFinite(n) || n <= 0) throw new Error('金額需為大於 0 的數字');
    return Math.round(n);
  })(p.amount);

  // 收入/支出
  var type = String(p.type || '').trim();
  if (type !== '收入' && type !== '支出') throw new Error('請選擇「收入/支出」');

  // 其他欄位
  var title = String(p.title || '').trim();
  if (!title) throw new Error('請輸入「項目名稱」');

  var category = String(p.category || '').trim();
  if (!category) throw new Error('請選擇「項目類別」');

  var payer = String(p.payer || '').trim();
  if (!payer) throw new Error('請選擇「付款人」');

  // 依欄位順序回傳（A..F）
  return [ date, title, category, amount, type, payer ];
}

function readLogoUrl() {
  // 讀取 collected!J2：可填入直接的圖片網址，或 dataURL
  // 範例假設是網址
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('collected');
  var v = sh.getRange('J2').getValue();
  // 回傳 {url:"..."} 或 {dataUrl:"..."} 皆可
  return { url: String(v) };
}

function getLogoUrl() {
  var sh = SpreadsheetApp.openById(SHEET_CFG.spreadsheetId).getSheetByName("collected");
  var url = sh.getRange("J2").getValue();
  return { url: url };
}

/** ================== 資訊頁：房貸進度 ==================
 * total 可傳入覆蓋，預設 12000000
 * 條件：類別 = '房貸'（type 是否為支出，皆會納入）
 */
function readLoanProgress(total) {
  var TOTAL = Number(total || 12000000); // 預設 1200 萬
  var sh = _sheet();
  var lr = sh.getLastRow();
  if (lr <= SHEET_CFG.headerRow) {
    return { paid: 0, total: TOTAL, percent: 0 };
  }

  var start = SHEET_CFG.headerRow + 1;
  var rows  = lr - SHEET_CFG.headerRow;

  // 直接用顯示值（避免日期/格式踩雷）
  var cat = sh.getRange(start, SHEET_CFG.COL.CAT,  rows, 1).getDisplayValues();
  var amt = sh.getRange(start, SHEET_CFG.COL.AMT,  rows, 1).getDisplayValues();

  var paid = 0;
  for (var i=0;i<rows;i++) {
    var c = String(cat[i][0] || '').trim();
    if (c === '房貸') {
      var n = Number(String(amt[i][0] || '').replace(/,/g,'').trim());
      if (!isNaN(n)) paid += n;
    }
  }
  var pct = TOTAL > 0 ? Math.max(0, Math.min(100, (paid / TOTAL) * 100)) : 0;
  return {
    paid: Math.round(paid),
    total: Math.round(TOTAL),
    percent: Math.round(pct * 10) / 10  // 1 位小數
  };
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// --- Script Properties 儲存類別 ---
function _prop(){ return PropertiesService.getScriptProperties(); }
function _getCategories(){
  try{ var txt = _prop().getProperty('categories'); return txt ? JSON.parse(txt) : []; }catch(e){ return []; }
}
function _setCategories(list){ _prop().setProperty('categories', JSON.stringify(list||[])); }

// Public API: 新增類別（不建新分頁）
function addCategory(obj){
  obj = obj || {};
  var name = (obj.name||'').toString().trim();
  var io   = (obj.ioType||'支出').toString().trim(); // 收入/支出/中性
  if (!name) throw new Error('缺少類別名稱');
  var list = _getCategories();
  var exists = list.some(function(it){ return it && it.name === name; });
  if (!exists){ list.push({ name:name, ioType:io }); _setCategories(list); }
  return { ok:true, count:list.length };
}