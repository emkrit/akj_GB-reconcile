/**
 * PGDRCVNT — simple refno → ตู้ matching
 *
 * Menu "PGDRCVNT":
 *   📥 Import & Clean      reads the raw receipts from Google Sheet "PGDRCVNT_Copy" (tab Sheet1),
 *                          cleans them into the PGDRCVNT tab, then refreshes Refno + Containers.
 *   🔄 Refresh             re-calculates totals and suggestions (run after typing some ตู้).
 *   ⏰ ตั้งรันอัตโนมัติ       installs the weekly run: every Monday 08:00 (Asia/Bangkok) = Import & Clean + Refresh.
 *   ⛔ ยกเลิกรันอัตโนมัติ     removes it. The two buttons above keep working either way.
 *
 * Every run (button or weekly) adds one row to the log sheet "Claude project log update":
 *   Finished (TH) | Task | Trigger | Run days | Status | Message | Datebased | Remark (# + list of Pending refno)
 *
 * Tabs:
 *   Refno       one row per refno. You only type in column G "ตู้":
 *                 C22          -> Match
 *                 ZS4/ZS5      -> Match, lines are split between the two ตู้ by product code
 *                 N/A          -> freebie / sample, no container
 *                 (empty)      -> Pending, with up to 3 suggestions
 *               Column "หมายเหตุ" is yours; it is kept on every refresh.
 *   Containers  one row per ตู้ from the payment tabs: paid vs received (from refnos you matched).
 *   PGDRCVNT    cleaned import. Column A = Status, B = ตู้ of each line, C = amount (price × qty).
 *
 * Suggestions: same supplier, ETD within WINDOW_DAYS of the refno's receive date,
 * ranked by how close the refno amount is to the ตู้'s amount not yet received.
 *
 * The source sheet, the payment tabs (CNS/HU/ZS/FX/NT) and Summary are only read, never edited.
 */

var CFG = {
  SOURCE_SHEET_ID: '1ZROpt5D4Pwwh1reKBIJQyukx1b7zgpB2iJVfP55bzBQ',   // PGDRCVNT_Copy
  SOURCE_TAB: 'Sheet1',
  RCV_SHEET: 'PGDRCVNT',
  REF_SHEET: 'Refno',
  CONT_SHEET: 'Containers',
  SEED_SHEET: 'Container_Map',     // used once, only if the Refno tab doesn't exist yet
  DATA_START_ROW: 3,               // PGDRCVNT rows 1-2 are headers
  RCV_HEADER_ROW: 2,               // row 2 = the column keys the script reads (no, run, refno, ...)
  PAY_TABS: ['CNS', 'HU', 'ZS', 'FX', 'NT'],
  SUPPLIER_BY_COOR: {
    'จ.006': 'CNS', 'ค.022': 'CNS',
    'อ.015': 'FX', 'อ.085': 'FX',
    'อ.042': 'HU', 'ฮ.005': 'ZS', 'น.008': 'NT'
  },
  WINDOW_DAYS: 21,                 // suggestion only if |receive date − ETD| ≤ this
  BEFORE_FIRST_ETD_DAYS: 7,        // refnos received earlier than (first ETD of that supplier − 7 days)
                                   // are older than the payment tabs and are left out of Refno
  SUGGEST_N: 3,
  NA: 'N/A',
  OK_ABS: 1, OK_PCT: 0.001,        // ✅ when |diff| ≤ max(1, 0.1%)
  NEAR_PCT: 0.02,                  // 🟡 when |diff| ≤ 2%

  // weekly run + log
  LOG_SHEET_ID: '1Va-g3b_m0oc6Zwwe9sP2FJOyafJmj__3Su1Wdiy7QCI',   // "Claude project log update"
  LOG_TAB: 'Sheet1',
  LOG_TASK: 'GB auto reconcile',
  RUN_DAYS: 'Every Monday',
  WEEKLY_FN: 'weeklyRun',
  WEEKLY_HOUR: 8,                  // Google starts it between 08:00 and 09:00
  TZ: 'Asia/Bangkok'
};

/* ============================== MENU ============================== */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('PGDRCVNT')
    .addItem('📥 Import & Clean', 'importAndClean')
    .addItem('🔄 Refresh', 'refreshMenu')
    .addSeparator()
    .addItem('⏰ ตั้งรันอัตโนมัติ (ทุกจันทร์ 8:00)', 'installWeekly')
    .addItem('⛔ ยกเลิกรันอัตโนมัติ', 'removeWeekly')
    .addToUi();
}

function refreshMenu() {
  var s;
  try { s = refresh(); } catch (e) { s = { error: e.message }; }
  writeRunLog('Manual: Refresh', s.error ? 'Error' : 'Success', s.error ? s.error : summaryLine(s), s);
  SpreadsheetApp.getUi().alert(summaryText(s));
}

/* ============================== WEEKLY RUN ============================== */

/** Started by the time trigger. No pop-ups (nobody is there to click them). */
function weeklyRun() {
  var out;
  try {
    out = importCore({ confirmDrop: function () { return false; } }); // big drop in rows -> stop, don't guess
  } catch (e) {
    out = { error: e.message };
  }
  writeRunLog('Script Auto run', out.error ? 'Error' : 'Success', out.error ? out.error : out.line, out.s, out.res);
}

function installWeekly() {
  var ui = SpreadsheetApp.getUi();
  removeTriggers_();
  ScriptApp.newTrigger(CFG.WEEKLY_FN).timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(CFG.WEEKLY_HOUR).nearMinute(0).inTimezone(CFG.TZ).create();
  ui.alert('⏰ ตั้งแล้ว: Import & Clean + Refresh อัตโนมัติ ทุกวันจันทร์ ' + CFG.WEEKLY_HOUR + ':00–' +
    (CFG.WEEKLY_HOUR + 1) + ':00 (เวลาไทย)\nผลแต่ละรอบบันทึกในไฟล์ "Claude project log update"' +
    '\nปุ่ม 📥 Import & Clean และ 🔄 Refresh ยังกดเองได้ตามปกติ');
}

function removeWeekly() {
  var n = removeTriggers_();
  SpreadsheetApp.getUi().alert(n ? '⛔ ยกเลิกรันอัตโนมัติแล้ว' : 'ยังไม่มีการตั้งรันอัตโนมัติ');
}

function removeTriggers_() {
  var n = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === CFG.WEEKLY_FN) { ScriptApp.deleteTrigger(t); n++; }
  });
  return n;
}

/* ============================== RUN LOG ============================== */

/** Appends one row to the log sheet. Never stops the run if the log can't be written. */
function writeRunLog(trigger, status, message, s, res) {
  try {
    var sh = SpreadsheetApp.openById(CFG.LOG_SHEET_ID).getSheetByName(CFG.LOG_TAB);
    if (!sh) return;
    var row = logRow(new Date(), trigger, status, message, s, res);
    sh.appendRow(row);
    sh.getRange(sh.getLastRow(), 2).setNumberFormat('dd/MM/yyyy HH:mm');
  } catch (e) {
    console.log('Log not written: ' + e.message);
  }
}

/** One log row, columns A..I. Pure, so it can be tested. */
function logRow(when, trigger, status, message, s, res) {
  var pend = s && s.pendingList ? s.pendingList : [];
  var remark = s && !s.error
    ? '# Refno Pending = ' + pend.length + (pend.length ? ': ' + pend.join(', ') : '')
    : '';
  var based = 'PGDRCVNT_Copy' + (res && res.from ? ' (' + res.from + ' – ' + res.to + ')' : '');
  return ['', when, CFG.LOG_TASK, trigger, CFG.RUN_DAYS, status, String(message).slice(0, 5000), based, remark.slice(0, 45000)];
}

function summaryLine(s) {
  if (s.error) return s.error;
  return 'refno ' + s.refs + ' | Match ' + s.Match + ' | N/A ' + s.NA + ' | Pending ' + s.Pending +
    (s.Pending ? ' (ยอด ' + fmtN(r2(s.pendAmt)).replace('+', '') + ')' : '') +
    ' | Containers: ' + Object.keys(s.boxes).map(function (k) { return k + ' ' + s.boxes[k]; }).join(', ');
}

/* ============================== IMPORT ============================== */

/**
 * Source columns (PGDRCVNT_Copy, row 1) -> key used in the PGDRCVNT tab (row 2).
 * Names that appear twice (ส่วนลด, หน่วย) are matched by their order.
 * The ERP document number is "code" (book is always 0002); it is stored as key "no".
 */
var SOURCE_COLS = [
  ['ab', 'ab'], ['วันที่', 'date'], ['เลขที่อ้างอิง', 'refno'], ['รหัส ผู้จำหน่าย', 'coorcode'],
  ['ชื่อ ผู้จำหน่าย', 'coorname'], ['แผนก', 'sect'], ['มูลค่าก่อนส่วนลด', 'amtbfdisc'], ['ส่วนลด', 'disc'],
  ['มูลค่า Vat', 'vat'], ['มูลค่าสุทธิ', 'totamt'], ['ใบสั่งซื้อ', 'refpo'], ['run', 'run'],
  ['รหัสสินค้า', 'prodcode'], ['ชื่อสินค้า', 'prodname'], ['ราคา/หน่วย', 'prodprice'], ['ส่วนลด', 'itemdisc'],
  ['จำนวน', 'qty'], ['หน่วย', 'unit'], ['จำนวนมาตรฐาน', 'stdqty'], ['หน่วย', 'stdunit'],
  ['จำนวนหน่วยคุมสต๊อค', 'stqty'], ['หน่วย', 'stunit'], ['ค้างทำ', 'bakqty'],
  ['invdate', 'invdate'], ['invcode', 'invcode'], ['invdisc', 'invdisc'], ['invprice', 'invprice'],
  ['invqty', 'invqty'], ['invunit', 'invunit'], ['invqtystd', 'invqtystd'], ['invstdunit', 'invstdunit'],
  ['book', 'book'], ['code', 'no'], ['refpo_line', 'refpo_line'], ['invtype', 'invtype']
];
var NEEDED_KEYS = ['ab', 'date', 'refno', 'coorcode', 'run', 'prodcode', 'prodprice', 'qty', 'no'];
var TEXT_KEYS = ['refno', 'coorcode', 'refpo', 'prodcode', 'invcode', 'book', 'no', 'refpo_line', 'ab'];
var DATE_KEYS = ['date', 'invdate'];
var NUM_KEYS = ['amtbfdisc', 'disc', 'vat', 'totamt', 'run', 'prodprice', 'itemdisc', 'qty', 'stdqty', 'stqty',
  'bakqty', 'invdisc', 'invprice', 'invqty', 'invqtystd'];

function importAndClean() {
  var ui = SpreadsheetApp.getUi();
  var out;
  try {
    out = importCore({
      interactive: true,
      confirmDrop: function (n, before) {
        return ui.alert('ข้อมูลใหม่มี ' + n + ' บรรทัด แต่ข้อมูลเดิมมี ' + before + ' บรรทัด\n' +
          'ช่วงวันที่อาจสั้นลง ต้องการ import ต่อไหม?', ui.ButtonSet.YES_NO) === ui.Button.YES;
      }
    });
  } catch (e) {
    out = { error: e.message };
  }
  if (out.cancelled) return;
  writeRunLog('Manual: Import & Clean', out.error ? 'Error' : 'Success', out.error ? out.error : out.line, out.s, out.res);
  ui.alert(out.error ? '❌ ' + out.error : out.text);
}

/**
 * Import + clean + refresh, shared by the button and the weekly run.
 * opts.confirmDrop(newRows, oldRows) -> true to go on when the new data has under half the old rows.
 * Returns {error} | {cancelled} | {res, s, text, line}.
 */
function importCore(opts) {
  var src;
  try {
    src = SpreadsheetApp.openById(CFG.SOURCE_SHEET_ID).getSheetByName(CFG.SOURCE_TAB);
  } catch (e) {
    return { error: 'เปิดไฟล์ PGDRCVNT_Copy ไม่ได้: ' + e.message };
  }
  if (!src) return { error: 'ไม่พบ tab "' + CFG.SOURCE_TAB + '" ในไฟล์ PGDRCVNT_Copy' };
  var res = cleanRows(src.getDataRange().getValues());
  if (res.error) return { error: 'Import stopped, nothing was changed. ' + res.error };
  var clean = res.rows, keys = res.keys, titles = res.titles;

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.RCV_SHEET);
  if (!sheet) return { error: 'Sheet "' + CFG.RCV_SHEET + '" not found.' };
  var before = Math.max(sheet.getLastRow() - CFG.DATA_START_ROW + 1, 0);
  if (before > 50 && clean.length < before * 0.5 && !opts.confirmDrop(clean.length, before)) {
    if (opts.interactive) return { cancelled: true };   // you clicked No
    return { error: 'Import stopped, nothing was changed. ข้อมูลใหม่มี ' + clean.length + ' บรรทัด แต่ข้อมูลเดิมมี ' +
      before + ' บรรทัด (น้อยกว่าครึ่ง) — ตรวจไฟล์ PGDRCVNT_Copy แล้วกด 📥 Import & Clean เอง' };
  }
  if (sheet.getFilter()) sheet.getFilter().remove(); // a filter during paste skips hidden rows
  var W = 3 + keys.length;
  if (sheet.getMaxColumns() < W) sheet.insertColumnsAfter(sheet.getMaxColumns(), W - sheet.getMaxColumns());
  var need = CFG.DATA_START_ROW + clean.length;
  if (sheet.getMaxRows() < need) sheet.insertRowsAfter(sheet.getMaxRows(), need - sheet.getMaxRows());
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearContent();

  // headers: row 1 = names as in the source, row 2 = keys the script reads
  sheet.getRange(1, 1, 2, W).setValues([['Status', 'ตู้', 'Amount'].concat(titles), ['', '', ''].concat(keys)])
    .setFontWeight('bold');
  // text columns stay text (09739 keeps its zero; 3297-1 is not a date)
  keys.forEach(function (k, j) {
    if (TEXT_KEYS.indexOf(k) >= 0) sheet.getRange(CFG.DATA_START_ROW, 4 + j, clean.length, 1).setNumberFormat('@');
    if (DATE_KEYS.indexOf(k) >= 0) sheet.getRange(CFG.DATA_START_ROW, 4 + j, clean.length, 1).setNumberFormat('dd/MM/yyyy');
  });
  sheet.getRange(CFG.DATA_START_ROW, 4, clean.length, keys.length).setValues(clean);

  // column C = price × qty
  var col = function (k) { return colLetter(4 + keys.indexOf(k)); };
  var r = CFG.DATA_START_ROW, lastRow = r - 1 + clean.length;
  var c = sheet.getRange(r, 3);
  c.setFormula('=' + col('prodprice') + r + '*' + col('qty') + r);
  if (lastRow > r) c.copyTo(sheet.getRange(r + 1, 3, lastRow - r, 1));
  sheet.getRange(r, 3, clean.length, 1).setNumberFormat('#,##0.00');
  SpreadsheetApp.flush();

  var s = refresh();
  var text = importText(res, clean.length) + '\n\n' + summaryText(s);
  var line = (s.error ? '' : 'Imported ' + clean.length + ' lines | ') + summaryLine(s) +
    (res.broken.length ? ' | ⚠️ ' + res.broken.length + ' rows unreadable' : '');
  return s.error ? { error: s.error, res: res } : { res: res, s: s, text: text, line: line };
}

function importText(res, n) {
  return 'Imported ' + n + ' lines from PGDRCVNT_Copy (' + res.from + ' – ' + res.to + ').' +
    '\nRead ' + res.read + ' rows. Removed: ' + res.blank + ' empty rows, ' + res.dupes + ' duplicate lines.' +
    '\nบรรทัดต่อ supplier: ' + Object.keys(res.bySup).map(function (k) { return k + ' ' + res.bySup[k]; }).join(', ') +
    (res.other ? '\nข้าม ' + res.other + ' บรรทัดของ supplier อื่น' : '') +
    (res.broken.length ? '\n⚠️ ' + res.broken.length + ' rows could not be read and were skipped (code/run): ' +
      res.broken.slice(0, 15).join(', ') + (res.broken.length > 15 ? ' …' : '') : '');
}

function colLetter(n) {
  var s = '';
  while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

/**
 * Cleans the raw rows (2-D array incl. header). Pure, so it can be tested.
 * Returns {rows, keys, titles, read, blank, dupes, broken, bySup, other, from, to} or {error}.
 */
function cleanRows(values) {
  if (!values || values.length < 2) return { error: 'ไฟล์ PGDRCVNT_Copy ไม่มีข้อมูล' };
  var head = values[0].map(function (h) { return String(h).replace(/\s+/g, ' ').trim(); });
  if (head.some(function (h) { return /^#(REF|N\/A|ERROR)|^Loading/i.test(h); }) || !head[0]) {
    return { error: 'หัวคอลัมน์ใน PGDRCVNT_Copy ยังไม่โหลด / IMPORTRANGE error: ' + head.slice(0, 5).join(', ') +
      '\nเปิดไฟล์ PGDRCVNT_Copy ให้ข้อมูลขึ้นก่อน แล้ว import ใหม่' };
  }
  // map source columns by name (2nd ส่วนลด / หน่วย by order)
  var keys = [], titles = [], idx = [], used = {};
  SOURCE_COLS.forEach(function (sc) {
    var from = used[sc[0]] == null ? 0 : used[sc[0]] + 1, i = head.indexOf(sc[0], from);
    if (i < 0) return;
    used[sc[0]] = i; keys.push(sc[1]); titles.push(sc[0]); idx.push(i);
  });
  var missing = NEEDED_KEYS.filter(function (k) { return keys.indexOf(k) < 0; });
  if (missing.length) return { error: 'ไม่พบคอลัมน์ที่ต้องใช้: ' + missing.join(', ') + '\nหัวคอลัมน์ที่พบ: ' + head.join(', ') };
  var K = {}; keys.forEach(function (k, j) { K[k] = j; });
  var data = values.slice(1).map(function (r) { return idx.map(function (i) { return i < r.length ? r[i] : ''; }); });

  // 1) drop empty rows
  var filled = data.filter(function (r) { return String(r[K.no]).trim() !== '' && String(r[K.run]).trim() !== ''; });
  // 2) dedupe on (ab, document no, run), preferring a row with invoice date
  var lineKey = function (r) { return String(r[K.ab]).trim() + '|' + String(r[K.no]).trim() + '|' + String(r[K.run]).trim(); };
  var hasInv = function (r) { return K.invdate != null && String(r[K.invdate]).trim() !== ''; };
  var ix = filled.map(function (r, i) { return { r: r, i: i, inv: hasInv(r) }; });
  ix.sort(function (a, b) { return a.inv === b.inv ? a.i - b.i : (b.inv ? 1 : 0) - (a.inv ? 1 : 0); });
  var seen = {}, keep = [];
  ix.forEach(function (x) { var k = lineKey(x.r); if (!seen[k]) { seen[k] = 1; keep.push(x); } });
  keep.sort(function (a, b) { return a.i - b.i; });
  // 3) readable rows only: run whole number, price and qty numbers
  var isNum = function (v) { return typeof v === 'number' || /^\s*\(?-?[\d,]*\.?\d+\)?\s*$/.test(String(v == null ? '' : v)); };
  var broken = [], good = [];
  keep.forEach(function (x) {
    var r = x.r;
    if (/^\s*\d+\s*$/.test(String(r[K.run])) && isNum(r[K.prodprice]) && isNum(r[K.qty])) good.push(r);
    else broken.push(String(r[K.no]).trim() + '/' + String(r[K.run]).trim());
  });
  // 4) import suppliers only
  var bySup = {}, other = 0, rows = [];
  good.forEach(function (r) {
    var sup = CFG.SUPPLIER_BY_COOR[String(r[K.coorcode]).trim()];
    if (!sup) { other++; return; }
    bySup[sup] = (bySup[sup] || 0) + 1; rows.push(r);
  });
  if (!rows.length) return { error: 'PGDRCVNT_Copy มี ' + good.length + ' บรรทัด แต่ไม่มี supplier นำเข้า (' +
    Object.keys(CFG.SUPPLIER_BY_COOR).join(', ') + ')' };
  // 5) types: real dates, numbers as numbers, ids as text
  var from = null, to = null;
  rows = rows.map(function (r) {
    return r.map(function (v, j) {
      var k = keys[j];
      if (DATE_KEYS.indexOf(k) >= 0) {
        var d = toDate(v);
        if (d && k === 'date') { if (!from || d < from) from = d; if (!to || d > to) to = d; }
        return d || (v === '' ? '' : String(v));
      }
      if (TEXT_KEYS.indexOf(k) >= 0) return v === '' || v == null ? '' : String(v).trim();
      if (NUM_KEYS.indexOf(k) >= 0) return v === '' || v == null ? '' : num(v);
      return v == null ? '' : v;
    });
  });
  var f = function (d) { return d ? d.getDate() + '/' + (d.getMonth() + 1) + '/' + d.getFullYear() : '?'; };
  return { rows: rows, keys: keys, titles: titles, read: data.length, blank: data.length - filled.length,
    dupes: filled.length - keep.length, broken: broken, bySup: bySup, other: other, from: f(from), to: f(to) };
}

/* ============================== HELPERS ============================== */

function num(v) {
  if (typeof v === 'number') return v;
  var s = String(v == null ? '' : v).trim();
  if (s === '' || s === '-') return 0;
  var neg = /^\(.*\)$/.test(s) || /^-/.test(s);
  var n = parseFloat(s.replace(/[(),\s-]/g, ''));
  if (isNaN(n)) return 0;
  return neg ? -n : n;
}

function toDate(v) {
  if (v instanceof Date) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  if (typeof v === 'number' && v > 20000) {
    var d = new Date(Math.round((v - 25569) * 864e5));
    return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
  }
  var iso = String(v || '').trim().match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (iso) return new Date(+iso[1], +iso[2] - 1, +iso[3]);
  var m = String(v || '').trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (!m) return null;
  var y = +m[3]; if (y < 100) y += 2000;
  return new Date(y, +m[2] - 1, +m[1]);
}

function normRef(v) { return String(v == null ? '' : v).toLowerCase().replace(/\s+/g, ' ').trim().replace(/^0+/, ''); }
function normBox(v) { return String(v == null ? '' : v).replace(/\s+/g, ' ').trim(); }
function normHead(h) { return String(h).toLowerCase().replace(/\s+/g, ' ').trim(); }
function r2(n) { return Math.round(n * 100) / 100; }
function days(a, b) { return Math.round((a - b) / 864e5); }
function fmtN(n) {
  var s = Math.abs(n).toFixed(0).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return (n < 0 ? '-' : '+') + s;
}

/* ============================== CORE (pure, no Sheets calls) ============================== */

/** lines -> one object per refno (by supplier + normalized refno). */
function buildRefnos(lines) {
  var by = {}, order = [];
  lines.forEach(function (l) {
    var x = by[l.refKey];
    if (!x) {
      x = by[l.refKey] = { key: l.refKey, sup: l.sup, refno: l.refno, date: null, n: 0, qty: 0, amt: 0 };
      order.push(x);
    }
    x.n++; x.qty += l.qty; x.amt += l.amt;
    if (l.date && (!x.date || l.date < x.date)) x.date = l.date;
  });
  order.forEach(function (x) { x.qty = r2(x.qty); x.amt = r2(x.amt); });
  return order;
}

/** payment lines -> one object per ตู้. */
function buildContainers(pay) {
  var by = {}, order = [];
  pay.forEach(function (p) {
    var c = by[p.box];
    if (!c) { c = by[p.box] = { box: p.box, sup: p.sup, etd: null, qty: 0, paid: 0, recv: 0, recvQty: 0, refs: {} }; order.push(c); }
    c.qty += p.qty; c.paid += p.amt;
    if (p.etd && (!c.etd || p.etd < c.etd)) c.etd = p.etd;
  });
  order.forEach(function (c) { c.qty = r2(c.qty); c.paid = r2(c.paid); });
  return { by: by, list: order };
}

/**
 * Gives every line its ตู้ from what you typed in Refno!G (input: {refKey: text}).
 * "A/B": each line goes to the ตู้ whose payment line has the same product code
 * (exact amount first, then closest), otherwise to the first ตู้.
 * Returns {refKey: {status, boxes, warn}}.
 */
function assignBoxes(lines, input, pay, conts) {
  var canon = {};
  conts.list.forEach(function (c) { canon[c.box.toLowerCase()] = c.box; });
  var info = {};
  Object.keys(input).forEach(function (k) {
    var t = normBox(input[k]);
    if (!t) { info[k] = { status: 'Pending', boxes: [], warn: '' }; return; }
    if (t.toUpperCase() === CFG.NA) { info[k] = { status: CFG.NA, boxes: [], warn: '' }; return; }
    var boxes = [], bad = [];
    t.split('/').forEach(function (b) {
      b = normBox(b); if (!b) return;
      var c = canon[b.toLowerCase()];
      if (c) { if (boxes.indexOf(c) < 0) boxes.push(c); } else bad.push(b);
    });
    info[k] = { status: 'Match', boxes: boxes, warn: bad.length ? '⚠️ ไม่พบตู้ ' + bad.join(', ') + ' ใน payment tab' : '' };
  });

  // Lines of a refno typed as "A/B" are split over the containers in passes, best evidence first,
  // across ALL lines of the refno (not line by line), so one line can't take another line's match:
  //   1. same product code + same amount       2. same amount (payment line may lack a product code)
  //   3. same receipt document (ab+no) as lines already placed
  //   4. same product code, closest amount      5. container whose ETD is closest to the receive date
  pay.forEach(function (p) { p.used = false; });
  var split = {};
  lines.forEach(function (l) {
    l.box = '';
    var x = info[l.refKey] || { status: 'Pending', boxes: [] };
    l.status = x.status;
    if (x.boxes.length === 1) l.box = x.boxes[0];
    else if (x.boxes.length > 1) (split[l.refKey] = split[l.refKey] || []).push(l);
  });
  var near = function (a, b) { return Math.abs(a - b) <= 0.02; };
  Object.keys(split).forEach(function (k) {
    var ls = split[k], boxes = info[k].boxes;
    var pl = pay.filter(function (p) { return boxes.indexOf(p.box) >= 0; });
    var open = function () { return ls.filter(function (l) { return !l.box; }); };
    var take = function (l, test) {
      var c = pl.filter(function (p) { return !p.used && test(p); });
      if (!c.length) return false;
      c.sort(function (a, b) { return Math.abs(a.amt - l.amt) - Math.abs(b.amt - l.amt); });
      c[0].used = true; l.box = c[0].box; return true;
    };
    open().forEach(function (l) { take(l, function (p) { return p.sku === l.sku && near(p.amt, l.amt); }); });
    open().forEach(function (l) { take(l, function (p) { return near(p.amt, l.amt); }); });
    open().forEach(function (l) {
      var cnt = {};
      ls.forEach(function (o) { if (o.box && o.doc && o.doc === l.doc) cnt[o.box] = (cnt[o.box] || 0) + 1; });
      var best = Object.keys(cnt).sort(function (a, b) { return cnt[b] - cnt[a]; })[0];
      if (best) { l.box = best; take({ amt: l.amt }, function (p) { return p.box === best && p.sku === l.sku; }); l.box = best; }
    });
    open().forEach(function (l) { take(l, function (p) { return p.sku === l.sku; }); });
    open().forEach(function (l) {
      var bx = boxes.slice().sort(function (a, b) {
        var ea = conts.by[a] && conts.by[a].etd, eb = conts.by[b] && conts.by[b].etd;
        return (l.date && ea ? Math.abs(l.date - ea) : 9e15) - (l.date && eb ? Math.abs(l.date - eb) : 9e15);
      });
      l.box = bx[0];
    });
  });

  lines.forEach(function (l) {
    if (!l.box) return;
    var c = conts.by[l.box]; if (!c) return;
    c.recv += l.amt; c.recvQty += l.qty; c.refs[l.refno] = 1;
  });
  conts.list.forEach(function (c) { c.recv = r2(c.recv); c.recvQty = r2(c.recvQty); });
  return info;
}

/** Up to SUGGEST_N suggestions for a Pending refno. */
function suggest(ref, conts) {
  if (Math.abs(ref.amt) < 0.005) return ['ยอด 0 → N/A?'];
  if (!ref.date) return ['ไม่มีวันที่รับ'];
  var c = conts.list.filter(function (c) {
    return c.sup === ref.sup && c.etd && Math.abs(days(ref.date, c.etd)) <= CFG.WINDOW_DAYS;
  }).map(function (c) {
    var left = r2(c.paid - c.recv);
    return { box: c.box, left: left, diff: r2(ref.amt - left), gap: days(ref.date, c.etd) };
  }).sort(function (a, b) { return Math.abs(a.diff) - Math.abs(b.diff); });
  if (!c.length) return ['ไม่มีตู้ ' + ref.sup + ' ภายใน ' + CFG.WINDOW_DAYS + ' วันจาก ETD'];
  return c.slice(0, CFG.SUGGEST_N).map(function (s) {
    var d = Math.abs(s.diff) <= Math.max(1, Math.abs(ref.amt) * 0.001) ? 'พอดี' : 'ต่าง ' + fmtN(s.diff);
    return s.box + ' · ' + d + ' · ' + (s.gap >= 0 ? s.gap + ' วันหลัง ETD' : -s.gap + ' วันก่อน ETD');
  });
}

function containerResult(c) {
  if (Math.abs(c.recv) < 0.005) return '⏳ ยังไม่รับ';
  var d = Math.abs(r2(c.recv - c.paid));
  if (d <= Math.max(CFG.OK_ABS, Math.abs(c.paid) * CFG.OK_PCT)) return '✅ ตรง';
  if (d <= Math.abs(c.paid) * CFG.NEAR_PCT) return '🟡 ใกล้เคียง';
  return '❌ ต่าง';
}

/**
 * Whole calculation. input: {refKey: typed ตู้}. Returns rows ready to write.
 */
function compute(lines, pay, input, carry) {
  // leave out refnos received before the payment tabs start (per supplier), unless a ตู้ was typed
  var first = {};
  pay.forEach(function (p) { if (p.etd && (!first[p.sup] || p.etd < first[p.sup])) first[p.sup] = p.etd; });
  var cut = {};
  Object.keys(first).forEach(function (s) { cut[s] = new Date(first[s].getTime() - CFG.BEFORE_FIRST_ETD_DAYS * 864e5); });
  var refDate = {};
  lines.forEach(function (l) { if (l.date && (!refDate[l.refKey] || l.date < refDate[l.refKey])) refDate[l.refKey] = l.date; });
  var skipped = {};
  lines = lines.filter(function (l) {
    var d = refDate[l.refKey];
    if (normBox(input[l.refKey]) || !d || !cut[l.sup] || d >= cut[l.sup]) return true;
    skipped[l.refKey] = 1; return false;
  });
  // refnos already in Refno but not in this CSV (export date range changed): keep them with their last totals
  var have = {}, old = {};
  lines.forEach(function (l) { have[l.refKey] = 1; });
  var all = lines.slice();
  Object.keys(carry || {}).forEach(function (k) {
    if (have[k]) return;
    var c = carry[k]; old[k] = c;
    all.push({ row: null, sup: c.sup, refno: c.refno, refKey: k, sku: '', qty: c.qty, amt: c.amt, date: c.date });
  });
  lines = all;
  var refs = buildRefnos(lines);
  refs.forEach(function (r) { if (old[r.key]) { r.n = old[r.key].n; r.old = true; } });
  refs.forEach(function (r) { if (!(r.key in input)) input[r.key] = ''; });
  var conts = buildContainers(pay);
  var info = assignBoxes(lines, input, pay, conts);
  var rank = { 'Pending': 0, 'Match': 1 }; rank[CFG.NA] = 2;
  refs.forEach(function (r) {
    var x = info[r.key];
    r.status = x.status; r.typed = normBox(input[r.key]);
    r.hints = x.status === 'Pending' ? suggest(r, conts) : (x.warn ? [x.warn] : []);
    if (r.old) r.hints = ['ไม่อยู่ใน CSV ล่าสุด (ใช้ยอดเดิม)'].concat(r.hints).slice(0, 3);
    // payment vs received, for the ตู้ this refno is in (all refnos of that ตู้ together)
    r.pay = null;
    if (x.status === 'Match' && x.boxes.length) {
      var pq = 0, pa = 0, rq = 0, ra = 0, others = {};
      x.boxes.forEach(function (b) {
        var c = conts.by[b]; if (!c) return;
        pq += c.qty; pa += c.paid; rq += c.recvQty; ra += c.recv;
        Object.keys(c.refs).forEach(function (o) { if (o !== r.refno) others[o] = 1; });
      });
      r.pay = { qty: r2(pq), amt: r2(pa), dQty: r2(rq - pq), dAmt: r2(ra - pa) };
      var o = Object.keys(others);
      if (o.length) r.hints = r.hints.concat(['รวมทั้งตู้ กับ refno ' + o.join(', ')]).slice(0, 3);
    }
  });
  refs.sort(function (a, b) {
    return rank[a.status] - rank[b.status] || CFG.PAY_TABS.indexOf(a.sup) - CFG.PAY_TABS.indexOf(b.sup) ||
      (a.date || 0) - (b.date || 0) || a.refno.localeCompare(b.refno);
  });
  conts.list.sort(function (a, b) {
    return CFG.PAY_TABS.indexOf(a.sup) - CFG.PAY_TABS.indexOf(b.sup) || (a.etd || 0) - (b.etd || 0) || a.box.localeCompare(b.box);
  });
  conts.list.forEach(function (c) { c.diff = r2(c.recv - c.paid); c.result = containerResult(c); });
  return { refs: refs, conts: conts.list, lines: lines, skipped: Object.keys(skipped).length, cut: cut };
}

/* ============================== SHEETS I/O ============================== */

function readLines(ss) {
  var sh = ss.getSheetByName(CFG.RCV_SHEET), last = sh.getLastRow();
  if (last < CFG.DATA_START_ROW) return [];
  var head = sh.getRange(CFG.RCV_HEADER_ROW, 1, 1, sh.getLastColumn()).getValues()[0].map(normHead);
  var c = function (n) { var i = head.indexOf(n); if (i < 0) throw new Error('PGDRCVNT: column "' + n + '" not found in row ' + CFG.RCV_HEADER_ROW); return i; };
  var iNo = c('no'), iRun = c('run'), iDate = c('date'), iRef = c('refno'), iCoor = c('coorcode'),
    iSku = c('prodcode'), iPrice = c('prodprice'), iQty = c('qty');
  var iAb = head.indexOf('ab');
  var vals = sh.getRange(CFG.DATA_START_ROW, 1, last - CFG.DATA_START_ROW + 1, sh.getLastColumn()).getValues();
  var out = [];
  vals.forEach(function (v, i) {
    if (v[iNo] === '' || v[iNo] == null) return;
    var sup = CFG.SUPPLIER_BY_COOR[String(v[iCoor]).trim()];
    if (!sup) return;
    out.push({
      row: CFG.DATA_START_ROW + i, doc: (iAb >= 0 ? String(v[iAb]).trim() : '') + '|' + v[iNo],
      sup: sup, refno: String(v[iRef]).trim(), refKey: sup + '|' + normRef(v[iRef]),
      sku: String(v[iSku]).trim(), qty: num(v[iQty]), amt: r2(num(v[iPrice]) * num(v[iQty])), date: toDate(v[iDate])
    });
  });
  return out;
}

function readPay(ss) {
  var out = [];
  CFG.PAY_TABS.forEach(function (tab) {
    var sh = ss.getSheetByName(tab);
    if (!sh || sh.getLastRow() < 2) return;
    var vals = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
    var raw = vals[0].map(function (h) { return String(h).trim(); }), head = raw.map(normHead);
    var iBox = head.indexOf('ตู้'), iEtd = head.indexOf('etd'), iSku = head.indexOf('product code'),
      iQty = head.indexOf('etd xx'), iAmt = raw.lastIndexOf('amount'); // lowercase "amount" = paid
    if (iBox < 0 || iSku < 0 || iAmt < 0) throw new Error(tab + ': needs "ตู้", "Product Code" and lowercase "amount" in row 1');
    for (var r = 1; r < vals.length; r++) {
      var v = vals[r], b = normBox(v[iBox]), amt = num(v[iAmt]);
      if (!b || amt === 0) continue;
      out.push({ sup: tab, box: b, etd: iEtd >= 0 ? toDate(v[iEtd]) : null, sku: String(v[iSku]).trim(),
        qty: iQty >= 0 ? num(v[iQty]) : 0, amt: r2(amt) });
    }
  });
  return out;
}

var REF_HEAD = ['Supplier', 'refno', 'วันที่รับ', 'บรรทัด', 'Qty รับ', 'Amount รับ', 'ตู้ (พิมพ์)', 'Status',
  'Qty จ่าย (ตู้)', 'Amount จ่าย (ตู้)', 'Diff Qty (รับ−จ่าย)', 'Diff Amount (รับ−จ่าย)',
  'แนะนำ 1', 'แนะนำ 2', 'แนะนำ 3', 'หมายเหตุ'];
var REF_STATUS_FORMULA = '={"Status";ARRAYFORMULA(IF(B2:B="","",IF(TRIM(G2:G)="","Pending",IF(UPPER(TRIM(G2:G))="N/A","N/A","Match"))))}';

/** Typed ตู้ + notes from Refno, or the one-time seed from Container_Map. */
function readInput(ss) {
  var input = {}, notes = {}, carry = {}, sh = ss.getSheetByName(CFG.REF_SHEET);
  if (sh && sh.getLastRow() >= 2) {
    var w = sh.getLastColumn();
    var hd = sh.getRange(1, 1, 1, w).getValues()[0].map(function (h) { return String(h).trim(); });
    var iNote = hd.indexOf('หมายเหตุ');
    sh.getRange(2, 1, sh.getLastRow() - 1, w).getValues().forEach(function (v) {
      if (!v[0] || v[1] === '') return;
      var k = String(v[0]).trim() + '|' + normRef(v[1]);
      input[k] = normBox(v[6]);
      if (iNote >= 0 && v[iNote] !== '') notes[k] = v[iNote];
      if (normBox(v[6])) carry[k] = { sup: String(v[0]).trim(), refno: String(v[1]).trim(), date: toDate(v[2]),
        n: num(v[3]), qty: num(v[4]), amt: r2(num(v[5])) };
    });
    return { input: input, notes: notes, carry: carry, sheet: sh };
  }
  var seed = ss.getSheetByName(CFG.SEED_SHEET);
  if (seed && seed.getLastRow() >= 2) {
    seed.getRange(2, 1, seed.getLastRow() - 1, 4).getValues().forEach(function (v) {
      if (String(v[3]).indexOf('✅') !== 0 || !normBox(v[2])) return;
      var k = String(v[0]).trim() + '|' + normRef(v[1]);
      input[k] = input[k] ? input[k] + '/' + normBox(v[2]) : normBox(v[2]);
    });
  }
  return { input: input, notes: notes, carry: carry, sheet: null };
}

function setupRefSheet(ss) {
  var sh = ss.insertSheet(CFG.REF_SHEET, 0);
  sh.setFrozenRows(1);
  return sh;
}

/** Header, Status formula, colours and widths of Refno (applied on every refresh). */
function layoutRefSheet(sh, n) {
  var W = REF_HEAD.length;
  if (sh.getMaxColumns() < W) sh.insertColumnsAfter(sh.getMaxColumns(), W - sh.getMaxColumns());
  var head = REF_HEAD.slice(); head[7] = REF_STATUS_FORMULA;
  sh.getRange(1, 1, 1, W).setValues([head]).setFontWeight('bold').setBackground('#e8eef7').setWrap(true);
  sh.setFrozenRows(1);
  sh.getRange(1, 7, sh.getMaxRows(), 1).setBackground('#fff8dc');
  sh.getRange(1, 7).setBackground('#e8eef7');
  sh.getRange(1, 9, 1, 4).setBackground('#dbeafe');   // payment side
  var R = function (a1) { return sh.getRange(a1); };
  var txt = function (t, bg) {
    return SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(t).setBackground(bg).setRanges([R('H2:H')]).build();
  };
  var f = function (formula, bg, a1) {
    return SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(formula).setBackground(bg).setRanges([R(a1)]).build();
  };
  // Diff Amount (L): green ≤ max(1, 0.1% of paid), yellow ≤ 2%, red above. Diff Qty (K): green ≤ 0.5, yellow ≤ 2%, red above.
  sh.setConditionalFormatRules([
    txt('Pending', '#fde68a'), txt('Match', '#bbf7d0'), txt('N/A', '#e5e7eb'),
    f('=AND(ISNUMBER($L2),ABS($L2)<=MAX(1,ABS($J2)*' + CFG.OK_PCT + '))', '#bbf7d0', 'L2:L'),
    f('=AND(ISNUMBER($L2),ABS($L2)<=ABS($J2)*' + CFG.NEAR_PCT + ')', '#fef9c3', 'L2:L'),
    f('=ISNUMBER($L2)', '#fecaca', 'L2:L'),
    f('=AND(ISNUMBER($K2),ABS($K2)<=0.5)', '#bbf7d0', 'K2:K'),
    f('=AND(ISNUMBER($K2),ABS($K2)<=ABS($I2)*' + CFG.NEAR_PCT + ')', '#fef9c3', 'K2:K'),
    f('=ISNUMBER($K2)', '#fecaca', 'K2:K')
  ]);
  [70, 120, 85, 55, 90, 105, 120, 75, 95, 110, 95, 115, 230, 230, 230, 200].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.setRowHeight(1, 42);
}

function refresh() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lines = readLines(ss), pay = readPay(ss), inp = readInput(ss);
  if (!lines.length) return { error: 'PGDRCVNT ไม่มีบรรทัดของ supplier นำเข้า (CNS/HU/ZS/FX/NT) จึงไม่ได้แก้ tab Refno / Containers.\n' +
    'ตรวจว่า CSV ที่ import เป็นรายงานรับของจาก supplier นำเข้า แล้วกด 📥 Import & Clean CSV ใหม่' };
  var R = compute(lines, pay, inp.input, inp.carry);

  // Refno
  var sh = inp.sheet || setupRefSheet(ss);
  layoutRefSheet(sh, R.refs.length);
  var max = Math.max(sh.getMaxRows() - 1, 1);
  sh.getRange(2, 1, max, 7).clearContent();
  sh.getRange(2, 9, max, REF_HEAD.length - 8).clearContent();
  if (R.refs.length) {
    sh.getRange(2, 2, R.refs.length, 1).setNumberFormat('@'); // keep refno as text (3297-1 is not a date)
    sh.getRange(2, 7, R.refs.length, 1).setNumberFormat('@');
    sh.getRange(2, 1, R.refs.length, 7).setValues(R.refs.map(function (r) {
      return [r.sup, r.refno, r.date || '', r.n, r.qty, r.amt, r.typed];
    }));
    sh.getRange(2, 9, R.refs.length, 8).setValues(R.refs.map(function (r) {
      var p = r.pay ? [r.pay.qty, r.pay.amt, r.pay.dQty, r.pay.dAmt] : ['', '', '', ''];
      var h = r.hints.concat(['', '', '']).slice(0, 3);
      return p.concat(h, [inp.notes[r.key] || '']);
    }));
    sh.getRange(2, 3, R.refs.length, 1).setNumberFormat('dd/MM/yyyy');
    sh.getRange(2, 5, R.refs.length, 2).setNumberFormat('#,##0.00');
    sh.getRange(2, 9, R.refs.length, 4).setNumberFormat('#,##0.00;[Red]-#,##0.00');
    var list = R.conts.map(function (c) { return c.box; }).concat([CFG.NA]).slice(0, 499);
    sh.getRange(2, 7, R.refs.length, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(true)
        .setHelpText('เลือกตู้, พิมพ์ ZS4/ZS5 ถ้าแยก 2 ตู้, หรือ N/A').build());
  }

  // Containers
  var ch = ss.getSheetByName(CFG.CONT_SHEET) || ss.insertSheet(CFG.CONT_SHEET, 1);
  ch.clear();
  var head = ['Supplier', 'ตู้', 'ETD', 'Qty (payment)', 'Paid amount', 'Received (refno ที่ match)', 'Diff', 'Result', 'refno'];
  ch.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#e8eef7');
  ch.setFrozenRows(1);
  if (R.conts.length) {
    ch.getRange(2, 9, R.conts.length, 1).setNumberFormat('@');
    ch.getRange(2, 1, R.conts.length, head.length).setValues(R.conts.map(function (c) {
      return [c.sup, c.box, c.etd || '', c.qty, c.paid, c.recv, c.diff, c.result, Object.keys(c.refs).join(', ')];
    }));
    ch.getRange(2, 3, R.conts.length, 1).setNumberFormat('dd/MM/yyyy');
    ch.getRange(2, 4, R.conts.length, 4).setNumberFormat('#,##0.00;[Red]-#,##0.00');
  }
  [70, 110, 85, 100, 110, 150, 100, 100, 260].forEach(function (w, i) { ch.setColumnWidth(i + 1, w); });

  // PGDRCVNT A/B
  var rs = ss.getSheetByName(CFG.RCV_SHEET), last = rs.getLastRow();
  if (last >= CFG.DATA_START_ROW) {
    var ab = [];
    for (var i = CFG.DATA_START_ROW; i <= last; i++) ab.push(['', '']);
    R.lines.forEach(function (l) { if (l.row) ab[l.row - CFG.DATA_START_ROW] = [l.status, l.box]; });
    rs.getRange(CFG.DATA_START_ROW, 1, ab.length, 2).setValues(ab);
  }

  var s = summarize(R);
  return s;
}

/** Counts for the pop-up and the log. Pure. pendingList = "CNS 11244" … (Pending refnos, in Refno order). */
function summarize(R) {
  var s = { refs: R.refs.length, Pending: 0, Match: 0, NA: 0, pendAmt: 0, boxes: {}, skipped: R.skipped, pendingList: [] };
  R.refs.forEach(function (r) {
    if (r.status === 'Pending') { s.Pending++; s.pendAmt += r.amt; s.pendingList.push(r.sup + ' ' + r.refno); }
    else if (r.status === 'Match') s.Match++; else s.NA++;
  });
  R.conts.forEach(function (c) { s.boxes[c.result] = (s.boxes[c.result] || 0) + 1; });
  return s;
}

function summaryText(s) {
  if (s.error) return '⚠️ ' + s.error;
  return 'refno ทั้งหมด ' + s.refs + (s.skipped ? ' (ไม่นับ ' + s.skipped + ' refno ที่รับของก่อนตู้แรกใน payment tab)' : '') +
    '\n  Match: ' + s.Match + '\n  N/A: ' + s.NA +
    '\n  Pending: ' + s.Pending + (s.Pending ? ' (ยอด ' + fmtN(r2(s.pendAmt)).replace('+', '') + ')' : '') +
    '\n\nContainers:\n' + Object.keys(s.boxes).map(function (k) { return '  ' + k + ': ' + s.boxes[k]; }).join('\n') +
    (s.Pending ? '\n\n→ ไปที่ tab Refno พิมพ์ตู้ในคอลัมน์ G สำหรับแถว Pending' : '\n\nไม่มี Pending 🎉');
}

/* for local testing in Node */
if (typeof module !== 'undefined') module.exports = { summarize: summarize, logRow: logRow, summaryLine: summaryLine, cleanRows: cleanRows, colLetter: colLetter, compute: compute, buildContainers: buildContainers, toDate: toDate, normRef: normRef, CFG: CFG };
