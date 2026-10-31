// Checks Data.js: reads a copy of the real data tabs (test/fixtures/schedule.json) through the fake
// spreadsheet, then a few broken rows, and confirms what comes back. Run with `npm test`.
'use strict';

const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { makeEnvironment } = require('./fake-sheets');

const FIXTURE = require('./fixtures/schedule.json');
const SRC = path.join(__dirname, '..', 'src');

/** A fresh fake spreadsheet holding the given tabs, with the script loaded into the same global scope. */
function setUp(tabs) {
  const env = makeEnvironment(new Date(2026, 9, 31, 12, 0));
  Object.entries(tabs).forEach(([name, rows]) => {
    const sh = env.book.insertSheet(name);
    sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  });
  const ctx = vm.createContext({ ...env.globals });
  fs.readdirSync(SRC).filter((f) => f.endsWith('.js')).sort()
    .forEach((f) => vm.runInContext(fs.readFileSync(path.join(SRC, f), 'utf8'), ctx, { filename: f }));
  return ctx;
}

const tests = [];
const test = (name, fn) => tests.push([name, fn]);

// ---------------------------------------------------------------- parsing

test('parseDays_ reads ranges, lists and full names', () => {
  const app = setUp(FIXTURE);
  assert.deepStrictEqual([...app.parseDays_('Mon-Fri')], ['Mon', 'Tue', 'Wed', 'Thu', 'Fri']);
  assert.deepStrictEqual([...app.parseDays_('Tue, Thu, Fri')], ['Tue', 'Thu', 'Fri']);
  assert.deepStrictEqual([...app.parseDays_('Fri, Mon')], ['Mon', 'Fri']);
  assert.deepStrictEqual([...app.parseDays_('monday, Wed–Fri')], ['Mon', 'Wed', 'Thu', 'Fri']);
  assert.strictEqual(app.parseDays_('Fri-Mon'), null);
  assert.strictEqual(app.parseDays_('Funday'), null);
  assert.strictEqual(app.parseDays_(''), null);
});

test('parseTime_ reads 12- and 24-hour times', () => {
  const app = setUp(FIXTURE);
  assert.strictEqual(app.parseTime_('7:30 AM'), 450);
  assert.strictEqual(app.parseTime_('12:00 PM'), 720);
  assert.strictEqual(app.parseTime_('10:00 PM'), 1320);
  assert.strictEqual(app.parseTime_('19:00'), 1140);
  assert.strictEqual(app.parseTime_('12:00 AM'), 1440);
  assert.strictEqual(app.parseTime_('13:00 PM'), null);
  assert.strictEqual(app.parseTime_('soon'), null);
  assert.strictEqual(app.formatTime_(450), '7:30 AM');
  assert.strictEqual(app.formatTime_(1320), '10:00 PM');
});

// ---------------------------------------------------------------- the real data

test('the real data reads cleanly', () => {
  const app = setUp(FIXTURE);
  const { categories, catalog, plan, problems } = app.readSchedule_();
  assert.deepStrictEqual([...problems], []);
  assert.strictEqual(categories.length, 6);
  assert.strictEqual(catalog.size, 9);
  assert.strictEqual(catalog.get('Lunch').category, null);
  assert.deepStrictEqual([...catalog.values()].filter((i) => i.breakdown).map((i) => i.name), ['Wake-Up', 'Daily Prep']);
  assert.strictEqual(plan.length, 14);

  const dnd = plan.find((p) => p.item === 'D&D');
  assert.deepStrictEqual({ start: dnd.start, end: dnd.end, days: [...dnd.days] }, { start: 1140, end: 1320, days: ['Mon'] });
  assert.strictEqual(plan.find((p) => p.item === 'Wake-Up').label, 'Wake-Up');   // blank label → item name
});

test('weekday focus time adds up to 6 hours a day', () => {
  const app = setUp(FIXTURE);
  const { plan } = app.readSchedule_();
  const perDay = {};
  plan.filter((p) => p.item === 'Block').forEach((p) => p.days.forEach((d) => { perDay[d] = (perDay[d] || 0) + (p.end - p.start); }));
  assert.deepStrictEqual(perDay, { Mon: 360, Tue: 360, Wed: 360, Thu: 360, Fri: 360 });
});

// ---------------------------------------------------------------- broken rows

test('broken rows are reported, not fatal', () => {
  const tabs = JSON.parse(JSON.stringify(FIXTURE));
  tabs.Catalog.push(['Yoga', 'Middle', 'Fitness', 'green', '', false]);
  tabs['Week Plan'].push(
    ['Nap', '', 'Mon', '2:00 PM', '3:00 PM', 'Weekday', ''],             // not on the Catalog
    ['Lunch', '', 'Sat', '12:00 PM', '1:00 PM', 'Weekday', ''],          // Sat isn't a weekday
    ['Lunch', '', 'Mon', '1:00 PM', '12:00 PM', 'Weekday', ''],          // ends before it starts
    ['Lunch', '', 'Someday', '12:00 PM', '1:00 PM', 'Weekday', ''],      // unreadable days
    ['Coffee Walk', '', 'Sat', '9:15 AM', '9:45 AM', 'Weekend - Home', ''],  // fine, but off the half hour
  );
  const app = setUp(tabs);
  const { plan, problems } = app.readSchedule_();
  const messages = [...problems].map((p) => `${p.tab}:${p.row} ${p.message}`);
  assert.deepStrictEqual(messages, [
    'Catalog:11 Yoga: Side should be Schedule or Activities',
    'Catalog:11 Yoga: category "Fitness" isn\'t on Categories',
    'Catalog:11 Yoga: Color should look like #9FC5E8',
    "Week Plan:16 Nap: isn't on the Catalog",
    "Week Plan:17 Lunch: Sat isn't part of the Weekday grid",
    'Week Plan:18 Lunch: End is not after Start',
    'Week Plan:19 Lunch: Days should look like "Mon-Fri" or "Tue, Thu"',
    'Week Plan:20 Coffee Walk: times not on a half hour; grids round them to the nearest half hour',
  ]);
  assert.strictEqual(plan.length, 15);   // the 14 real rows plus the off-the-half-hour one
});

// ---------------------------------------------------------------- run

let failed = 0;
tests.forEach(([name, fn]) => {
  try { fn(); console.log(`✓ ${name}`); } catch (err) { failed++; console.log(`✗ ${name}\n  ${err.message.split('\n').join('\n  ')}`); }
});
console.log(failed ? `\n${failed} of ${tests.length} failed` : `\nAll ${tests.length} passed`);
process.exitCode = failed ? 1 : 0;
