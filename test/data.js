// Checks Data.js: reads a copy of the real data tabs (test/fixtures/schedule.json) through the fake
// spreadsheet, then a few broken rows, and confirms what comes back. Run with `npm test`.
'use strict';

const assert = require('assert');
const helpers = require('./helpers');

const FIXTURE = helpers.fixture();

/** The script loaded into a fake spreadsheet holding `tabs` (the real data unless given). */
const setUp = (tabs) => helpers.setUp(tabs).app;

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
    'Catalog:11 Yoga: Side should be Schedule, Activities or Touchstone',
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

test('View min on the Catalog: blank means 5, bad values are reported', () => {
  const tabs = JSON.parse(JSON.stringify(FIXTURE));
  tabs.Catalog[2][6] = 0;                                              // Block: out of range
  const { catalog, problems } = setUp(tabs).readSchedule_();
  assert.strictEqual(catalog.get('Wake-Up').viewMinutes, 1);
  assert.strictEqual(catalog.get('Daily Prep').viewMinutes, 5);
  assert.strictEqual(catalog.get('Block').viewMinutes, 5);
  assert.deepStrictEqual([...problems].map((p) => p.message), ['Block: View min should be a whole number from 1 to 30']);
});

test('items starting before 6 AM are reported, and still drawn from 6 AM', () => {
  const tabs = JSON.parse(JSON.stringify(FIXTURE));
  tabs['Week Plan'].push(['Coffee Walk', 'Early walk', 'Sat', '5:00 AM', '6:30 AM', 'Weekend - Home', '']);
  const env = helpers.setUp(tabs);
  const { problems } = env.app.readSchedule_();
  assert.deepStrictEqual([...problems].map((p) => `${p.row} ${p.message}`), ["16 Coffee Walk: runs outside the grids' 6:00 AM – midnight; the part outside isn't shown"]);
  env.app.drawAll_(false);
  assert.strictEqual(env.book.dump().sheets.Weekends.cells.H5.v, 'Early walk', 'clipped to the 6:00 AM chunk');
});

test('Sort Week Plan: grid, then first day, then start; unreadable rows end their grid, blanks last', () => {
  const tabs = JSON.parse(JSON.stringify(FIXTURE));
  tabs['Week Plan'].splice(2, 0, ['Coffee Walk', '', 'Sat', '9:00 AM', '9:30 AM', 'Weekend - Home', 'home']);
  tabs['Week Plan'].splice(1, 0, ['', '', '', '', '', '', '']);
  tabs['Week Plan'].push(['Lunch', '', 'Someday', '12:00 PM', '1:00 PM', 'Weekday', 'broken']);
  tabs['Week Plan'].push(['Coffee Walk', '', 'Fri', '7:00 PM', '9:00 PM', 'Weekend - Travel', 'travel']);
  const env = helpers.setUp(tabs);
  env.app.sortWeekPlan();
  const rows = env.book.getSheetByName('Week Plan').getRange(2, 1, 18, 7).getValues();
  const brief = rows.map((r) => (r[0] ? `${r[0]}|${r[2]}|${r[3]}` : ''));
  assert.deepStrictEqual(brief, [
    'Wake-Up|Mon-Fri|6:30 AM', 'Daily Prep|Mon-Fri|7:00 AM', 'Block|Mon-Fri|7:30 AM', 'Coffee Walk|Mon-Fri|9:30 AM',
    'Block|Mon-Fri|10:00 AM', 'EMDR|Mon|12:00 PM', 'Block|Mon, Tue, Thu, Fri|1:00 PM', 'D&D|Mon|7:00 PM',
    'Lunch|Tue, Thu, Fri|12:00 PM', 'Privates Block|Tue|4:00 PM', 'Reverie|Tue|6:00 PM', 'EMDR|Wed|11:00 AM',
    'Block|Wed|3:00 PM', 'Privates Block|Thu|6:00 PM', 'Lunch|Someday|12:00 PM',
    'Coffee Walk|Sat|9:00 AM', 'Coffee Walk|Fri|7:00 PM', '',
  ]);
  assert.strictEqual(rows[15][6], 'home', 'whole rows move together, notes included');
});

// ---------------------------------------------------------------- run

helpers.run(tests);
