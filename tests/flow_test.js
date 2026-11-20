// node tests/flow_test.js
// Runs the script with stand-ins for SpreadsheetApp / ScriptApp and checks the weekly run, the log row
// and the schedule buttons. No Google account needed.
const fs = require('fs'), path = require('path'), vm = require('vm'), assert = require('assert');
const code = fs.readFileSync(path.join(__dirname, '..', 'PGDRCVNT_Simple.gs'), 'utf8');

const logged = [], alerts = [];
let triggers = [];
const sheet = (vals) => ({
  getDataRange: () => ({ getValues: () => vals }),
  getLastRow: () => 2000,
  appendRow: (r) => logged.push(r),
  getRange: () => ({ setNumberFormat() {} })
});
const ctx = {
  console,
  SpreadsheetApp: {
    // log sheet ok; source sheet returns a broken IMPORTRANGE header
    openById: (id) => ({ getSheetByName: () => (id.startsWith('1Va') ? sheet() : sheet([['#REF!']])) }),
    getActiveSpreadsheet: () => ({ getSheetByName: () => sheet() }),
    getUi: () => ({ alert: (m) => alerts.push(m) })
  },
  ScriptApp: {
    WeekDay: { MONDAY: 'MONDAY' },
    getProjectTriggers: () => triggers,
    deleteTrigger: (t) => { triggers = triggers.filter((x) => x !== t); },
    newTrigger: (fn) => {
      const o = {};
      Object.assign(o, {
        timeBased: () => o, onWeekDay: (d) => ((o.day = d), o), atHour: (h) => ((o.hour = h), o),
        nearMinute: () => o, inTimezone: (z) => ((o.tz = z), o),
        create: () => triggers.push({ getHandlerFunction: () => fn, ...o })
      });
      return o;
    }
  }
};
vm.createContext(ctx);
vm.runInContext(code, ctx);

// 1) weekly run with a broken source: nothing imported, one Error row logged, no pop-up
ctx.weeklyRun();
assert.strictEqual(logged.length, 1);
assert.strictEqual(logged[0][2], 'GB auto reconcile');
assert.strictEqual(logged[0][3], 'Script Auto run');
assert.strictEqual(logged[0][5], 'Error');
assert.strictEqual(alerts.length, 0);

// 2) schedule buttons: clicking twice keeps one trigger, Monday 8, Bangkok; remove clears it
ctx.installWeekly(); ctx.installWeekly();
assert.strictEqual(triggers.length, 1);
assert.deepStrictEqual([triggers[0].day, triggers[0].hour, triggers[0].tz], ['MONDAY', 8, 'Asia/Bangkok']);
ctx.removeWeekly();
assert.strictEqual(triggers.length, 0);

// 3) summary + log row: Remark lists the Pending refnos
const s = ctx.summarize({
  refs: [
    { status: 'Pending', sup: 'CNS', refno: '11244', amt: 1000 },
    { status: 'Match', sup: 'ZS', refno: '8262', amt: 5 },
    { status: 'N/A', sup: 'HU', refno: 'gift', amt: 0 },
    { status: 'Pending', sup: 'FX', refno: '16060-1', amt: 250 }
  ],
  conts: [{ result: '✅ ตรง' }, { result: '⏳ ยังไม่รับ' }],
  skipped: 0
});
const row = ctx.logRow(new Date(2026, 10, 23, 8, 3), 'Script Auto run', 'Success', ctx.summaryLine(s), s,
  { from: '3/1/2025', to: '18/11/2026' });
assert.strictEqual(row.length, 9);
assert.strictEqual(row[7], 'PGDRCVNT_Copy (3/1/2025 – 18/11/2026)');
assert.strictEqual(row[8], '# Refno Pending = 2: CNS 11244, FX 16060-1');
assert.ok(row[6].includes('Pending 2'));

console.log('All checks passed');
