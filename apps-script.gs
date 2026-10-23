/**
 * ACL 術後地圖：匿名費用回報
 *
 * 這段程式放在你的 Google 試算表裡（擴充功能 → Apps Script）。
 * - 網站送出的回報，會存到「回報」分頁的新的一列。
 * - 網站讀取統計時，只會拿到中位數和你勾選「公開」的留言，原始回報不會外流。
 *
 * 你在試算表裡可以做的事：
 * - 「公開」欄打勾：這筆的「想對傷友說的話」會出現在網站時間軸上（對方也要同意公開才會顯示）。
 * - 「排除」欄打勾：這筆不計入統計（例如亂填、重複、明顯錯誤）。
 * 改完之後，網站最多 5 分鐘內會更新。
 */

const SHEET_NAME = '回報';
const MIN_N = 3;            // 同一間醫院、同一個類別，累積幾筆以上才公開數字
const CACHE_SECONDS = 300;  // 統計結果暫存 5 分鐘

const HEADERS = [
  '送出時間', '手術年份', '縣市', '醫院', '醫院層級', '移植物', '其他處置',
  '自費醫材或注射', '自費醫材其他', '自費醫材注射總額', '病房差額', '住院天數', '支架護具花費',
  '麻醉止痛項目', '麻醉止痛其他', '麻醉止痛總額', '住院照顧', '看護費用',
  '術後物理治療', '自費物治每次', '商業保險理賠',
  '想對傷友說的話', '送給哪個階段', '同意公開', '公開', '排除'
];

const CITIES = ['臺北市', '新北市', '基隆市', '桃園市', '新竹市', '新竹縣', '苗栗縣', '臺中市', '彰化縣', '南投縣', '雲林縣', '嘉義市', '嘉義縣', '臺南市', '高雄市', '屏東縣', '宜蘭縣', '花蓮縣', '臺東縣', '澎湖縣', '金門縣', '連江縣',
  '台北市', '台中市', '台南市', '台東縣'];
const LEVELS = ['醫學中心', '區域醫院', '地區醫院', '診所', '不確定'];
const GRAFTS = ['自體膕旁肌腱', '自體髕骨肌腱', '自體股四頭肌腱', '異體肌腱', '人工韌帶', '不確定'];
const COMBOS = ['半月板修補', '半月板部分切除', '其他韌帶', '軟骨處置'];
const DEVICES = ['可吸收骨釘', '懸吊式固定', '縫合錨釘', '異體肌腱', '人工韌帶', '半月板縫合器', '軟骨修復', '關節鏡耗材', '防沾黏', '自費敷料', 'PRP', '羊膜', '玻尿酸', '無術中注射', '其他', '不確定'];
const ANES = ['神經阻斷', '自控式止痛', '自費止痛針', '止暈止吐', '保暖毯', '新型麻醉解藥', '其他', '沒有'];
const CARERS = ['家人朋友', '醫院配合的看護', '健保住院整合照護', '仲介或平台找的看護', '自己'];
const PT_TYPES = ['健保復健', '自費物理治療', '兩者都有', '還沒開始'];
const INSURANCE = ['有', '沒有', '還在申請'];
const STAGES = ['pre', 'w0', 'w2', 'w6', 'm3', 'm9'];

/** 第一次使用時，在 Apps Script 編輯器選這個函式並按「執行」，會建立分頁和欄位。 */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  const rows = Math.max(sheet.getMaxRows() - 1, 1);
  sheet.getRange(2, HEADERS.indexOf('公開') + 1, rows, 2).insertCheckboxes();
}

// ---------------------------------------------------------------- 收回報

function doPost(e) {
  try {
    if (tooMany_()) return json_({ ok: false, error: 'busy' });
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (d.website) return json_({ ok: true }); // 防機器人的隱藏欄位
    const row = toRow_(d);
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const sheet = sheet_();
      sheet.appendRow(row);
      const r = sheet.getLastRow();
      sheet.getRange(r, HEADERS.indexOf('公開') + 1, 1, 2).insertCheckboxes();
      CacheService.getScriptCache().remove('stats');
    } finally {
      lock.releaseLock();
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: 'invalid' });
  }
}

function toRow_(d) {
  const year = Number(d.year);
  const thisYear = new Date().getFullYear();
  if (!(year >= thisYear - 30 && year <= thisYear + 1)) throw new Error('year');
  const city = pick_(d.city, CITIES, true);
  const level = pick_(d.level, LEVELS, true);
  const graft = pick_(d.graft, GRAFTS, true);
  return [
    new Date(),
    year,
    city.replace(/^台/, '臺'),
    text_(d.hospital, 40),
    level,
    graft,
    list_(d.combined, COMBOS),
    list_(d.devices, DEVICES),
    text_(d.devOther, 60),
    num_(d.deviceCost, 1000000),
    num_(d.wardCost, 1000000),
    num_(d.days, 60),
    num_(d.braceCost, 200000),
    list_(d.anes, ANES),
    text_(d.anesOther, 60),
    num_(d.anesCost, 200000),
    pick_(d.carer, CARERS, false),
    num_(d.carerCost, 300000),
    pick_(d.ptType, PT_TYPES, false),
    num_(d.ptCost, 20000),
    pick_(d.insurance, INSURANCE, false),
    text_(d.note, 300),
    STAGES.indexOf(d.noteStage) >= 0 ? d.noteStage : '',
    d.noteConsent === true,
    false,
    false
  ];
}

function pick_(v, allowed, required) {
  if (allowed.indexOf(v) >= 0) return v;
  if (required) throw new Error('missing');
  return '';
}
function list_(v, allowed) {
  return Array.isArray(v) ? v.filter(function (x) { return allowed.indexOf(x) >= 0; }).join('、') : '';
}
function num_(v, max) {
  if (v === null || v === undefined || v === '') return '';
  const n = Math.round(Number(v));
  if (!isFinite(n) || n < 0 || n > max) throw new Error('number');
  return n;
}
function text_(v, max) {
  let s = String(v || '').replace(/\s+/g, ' ').trim().slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s; // 避免被試算表當成公式
  return s;
}
function tooMany_() {
  const cache = CacheService.getScriptCache();
  const key = 'posts-' + Math.floor(Date.now() / 60000);
  const n = Number(cache.get(key) || 0) + 1;
  cache.put(key, String(n), 120);
  return n > 30; // 每分鐘最多 30 筆
}

// ---------------------------------------------------------------- 提供統計

function doGet() {
  const cache = CacheService.getScriptCache();
  let out = cache.get('stats');
  if (!out) {
    out = JSON.stringify(buildStats_(readRows_()));
    cache.put('stats', out, CACHE_SECONDS);
  }
  return ContentService.createTextOutput(out).setMimeType(ContentService.MimeType.JSON);
}

function readRows_() {
  const sheet = sheet_();
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const values = sheet.getRange(2, 1, last - 1, HEADERS.length).getValues();
  return values.map(function (v, i) {
    const o = { row: i + 2 };
    HEADERS.forEach(function (h, j) { o[h] = v[j]; });
    return o;
  });
}

function buildStats_(rows) {
  const ok = rows.filter(function (r) { return r['手術年份'] && r['排除'] !== true; });
  const med = function (list, key) {
    const a = list.map(function (r) { return r[key]; })
      .filter(function (v) { return typeof v === 'number' && isFinite(v); })
      .sort(function (x, y) { return x - y; });
    if (!a.length) return { n: 0, median: null };
    const m = Math.floor(a.length / 2);
    return { n: a.length, median: a.length % 2 ? a[m] : Math.round((a[m - 1] + a[m]) / 2) };
  };
  const gate = function (s) { return s.n >= MIN_N ? s.median : null; };
  const group = function (list, keyFn) {
    const g = {};
    list.forEach(function (r) { const k = keyFn(r); if (k) (g[k] = g[k] || []).push(r); });
    return g;
  };

  const enough = ok.length >= MIN_N;
  const byGraft = group(ok, function (r) { return r['移植物'] || '不確定'; });
  const byHosp = group(ok, function (r) {
    const h = String(r['醫院'] || '').replace(/^'/, '').replace(/台/g, '臺').trim();
    return h ? h + '｜' + r['縣市'] : '';
  });
  const hospKeys = Object.keys(byHosp);

  return {
    updated: new Date().toISOString(),
    minN: MIN_N,
    n: ok.length,
    overall: {
      deviceCost: enough ? med(ok, '自費醫材注射總額').median : null,
      wardCost: enough ? med(ok, '病房差額').median : null
    },
    byGraft: Object.keys(byGraft).map(function (g) {
      return { graft: g, n: byGraft[g].length, deviceCost: byGraft[g].length >= MIN_N ? med(byGraft[g], '自費醫材注射總額').median : null };
    }).sort(function (a, b) { return b.n - a.n; }),
    other: [
      ['anesCost', '麻醉止痛總額'], ['carerCost', '看護費用'], ['braceCost', '支架護具花費'], ['ptCost', '自費物治每次']
    ].map(function (p) { const s = med(ok, p[1]); return { key: p[0], n: s.n, median: gate(s) }; }),
    byHosp: hospKeys.filter(function (k) { return byHosp[k].length >= MIN_N; }).map(function (k) {
      const list = byHosp[k], parts = k.split('｜');
      return {
        hospital: parts[0], city: parts[1], n: list.length,
        deviceCost: med(list, '自費醫材注射總額').median,
        wardCost: med(list, '病房差額').median,
        days: med(list, '住院天數').median
      };
    }).sort(function (a, b) { return b.n - a.n; }),
    hospitalsWaiting: hospKeys.filter(function (k) { return byHosp[k].length < MIN_N; }).length,
    wall: ok.filter(function (r) { return r['公開'] === true && r['同意公開'] === true && String(r['想對傷友說的話'] || '').trim(); })
      .map(function (r) { return { id: r.row, text: String(r['想對傷友說的話']).replace(/^'/, ''), stage: r['送給哪個階段'] || '' }; })
  };
}

// ---------------------------------------------------------------- 小工具

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('請先執行 setup()');
  return sheet;
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
