// Checks that problems show up where they'd be fixed: the ⚠ Check column on each input tab, cleared when fixed,
// plus the Check data dialog.
'use strict';

const assert = require('assert');
const { fixture, setUp, run } = require('./helpers');

/** Fixture (optionally changed), drawn once. */
function drawn(change = () => {}) {
  const tabs = fixture();
  tabs.Steps = [['Item', 'Step', 'Min'], ['Daily Prep'], ['', 'Shower', 10]];
  change(tabs);
  const env = setUp(tabs);
  env.app.drawAll_(false);
  return env;
}

/** Type values into cells the way a person would, then fire onEdit. */
function type(app, sh, row, col, values) {
  const range = sh.getRange(row, col, 1, values.length);
  range.setValues([values]);
  app.onEdit({ range });
}

/** Non-empty ⚠ Check messages on a tab, by row: { 16: '…' }. */
function checks(book, tab, col) {
  const sh = book.getSheetByName(tab);
  const out = {};
  sh.getRange(2, col, Math.max(sh.getLastRow() - 1, 1), 1).getValues().forEach((r, i) => { if (r[0]) out[i + 2] = r[0]; });
  return out;
}

const tests = [];
const test = (name, fn) => tests.push([name, fn]);

test('clean data: every input tab gets an empty ⚠ Check column', () => {
  const { book } = drawn();
  [['Categories', 'E'], ['Catalog', 'I'], ['Week Plan', 'H'], ['Steps', 'D']].forEach(([tab, col]) => {
    assert.strictEqual(book.dump().sheets[tab].cells[`${col}1`].v, '⚠ Check', tab);
  });
  assert.deepStrictEqual(checks(book, 'Week Plan', 8), {});
});

test('a broken row gets its reason in ⚠ Check, and fixing it clears the message', () => {
  const { app, book } = drawn((tabs) => tabs['Week Plan'].push(['Lunch', '', 'Someday', '12:00 PM', '1:00 PM', 'Weekday', '']));
  assert.deepStrictEqual(checks(book, 'Week Plan', 8), { 16: 'Lunch: Days should look like "Mon-Fri" or "Tue, Thu"' });
  type(app, book.getSheetByName('Week Plan'), 16, 3, ['Sat']);
  assert.deepStrictEqual(checks(book, 'Week Plan', 8), { 16: "Lunch: Sat isn't part of the Weekday grid" });
  type(app, book.getSheetByName('Week Plan'), 16, 6, ['Weekend - Home']);
  assert.deepStrictEqual(checks(book, 'Week Plan', 8), {});
});

test('several problems on one row are joined; overlaps land on the later row', () => {
  const { book } = drawn((tabs) => {
    tabs.Catalog.push(['Yoga', 'Middle', 'Fitness', 'green', '', false, '']);
    tabs['Week Plan'].push(['Coffee Walk', 'Late Coffee', 'Tue', '12:30 PM', '1:30 PM', 'Weekday', '']);
  });
  assert.deepStrictEqual(checks(book, 'Catalog', 9), {
    11: 'Yoga: Side should be Schedule, Activities or Touchstone · Yoga: category "Fitness" isn\'t on Categories · Yoga: Color should look like #9FC5E8',
  });
  assert.deepStrictEqual(checks(book, 'Week Plan', 8), { 16: 'Late Coffee overlaps Lunch on Tue (Activities)' });
});

test('upgrade: an old Catalog with ⚠ Check in H gets Split in H and ⚠ Check moves to I, nothing typed lost', () => {
  const { book } = drawn((tabs) => {
    tabs.Catalog[0].push('⚠ Check');
    tabs.Catalog[1].push('old message');
  });
  const row1 = book.getSheetByName('Catalog').getRange(1, 1, 1, 9).getValues()[0];
  assert.deepStrictEqual(row1.slice(6), ['View min', 'Split', '⚠ Check']);
  assert.strictEqual(book.getSheetByName('Catalog').getRange(2, 7).getValue(), 1, "Wake-Up's View min untouched");
  assert.deepStrictEqual(checks(book, 'Catalog', 9), {}, 'the stale message is cleared from its new spot');
  assert.deepStrictEqual(checks(book, 'Catalog', 8), {}, 'and nothing is left in Split');
});

test('Steps problems are marked on Steps', () => {
  const { book } = drawn((tabs) => { tabs.Steps.push(['Nap'], ['', 'Lie down', 20]); });
  assert.deepStrictEqual(checks(book, 'Steps', 4), { 4: "Nap: isn't on the Catalog" });
});

test('the column-A highlight rule is added once, and Steps keeps its heading shading', () => {
  const { app, book } = drawn((tabs) => { delete tabs.Steps; });          // let the script create Steps
  app.drawAll_(false);
  app.refreshAll();
  const rules = book.dump().sheets.Steps.cfRules;
  assert.strictEqual(rules.length, 2, 'heading shading + problem highlight, no duplicates');
  assert.strictEqual(book.dump().sheets['Week Plan'].cfRules.length, 1);
});

test('Check data lists every problem, or says there are none', () => {
  const clean = drawn();
  clean.app.checkData();
  assert.deepStrictEqual(clean.alerts.pop(), ['🗓️ Schedule', 'No problems found ✓', 'OK']);

  const broken = drawn((tabs) => {
    tabs['Week Plan'].push(['Nap', '', 'Mon', '2:00 PM', '3:00 PM', 'Weekday', '']);
    tabs.Steps.push(['', '', ''], ['Nap']);
  });
  broken.app.checkData();
  const [title, body] = broken.alerts.pop();
  assert.strictEqual(title, '🗓️ Schedule: check your data');
  assert.deepStrictEqual(body.split('\n'), ["Week Plan, row 16: Nap: isn't on the Catalog", "Steps, row 5: Nap: isn't on the Catalog"]);
});

run(tests);
