// Checks Breakdowns.js against a copy of the real data: what gets drawn, that typed steps survive every
// redraw, overflow flagging, unchecking/rechecking, and that a redraw of an unchanged tab changes nothing.
'use strict';

const assert = require('assert');
const { fixture, setUp, run } = require('./helpers');

/** The Breakdowns tab as it was left after the hand-built grouped layout (Daily Prep with steps A and B). */
const LEGACY = [
  ['Item', 'Step', 'Min', 'Time', 'End (min)'],
  ['Daily Prep', '', '', '', ''],
  ['', 'A', '', '0:00–0:00', 0],
  ['', 'B', '', '0:00–0:00', 0],
];

/** Fixture tabs plus a Breakdowns tab, drawn once. */
function drawn(breakdowns = LEGACY, change = () => {}) {
  const tabs = fixture();
  tabs.Breakdowns = breakdowns;
  change(tabs);
  const { app, book } = setUp(tabs);
  app.drawBreakdowns_();
  return { app, book, sh: book.getSheetByName('Breakdowns') };
}

/** All values on a sheet as plain rows, trailing blanks trimmed. */
function rows(sh) {
  const last = sh.getLastRow();
  return sh.getRange(1, 1, last, sh.getMaxColumns()).getValues().map((r) => {
    const out = [...r];
    while (out.length && out[out.length - 1] === '') out.pop();
    return out;
  });
}

/** Row number of the first row whose column A (the key) and column B match. */
function findRow(sh, key, step) {
  return rows(sh).findIndex((r) => (r[0] || '') === key && (r[1] || '') === step) + 1;
}

/** Type values into cells the way a person would, then fire onEdit. */
function type(app, sh, row, col, values) {
  const range = sh.getRange(row, col, 1, values.length);
  range.setValues([values]);
  app.onEdit({ range });
}

const steps = (app, sh) => Object.fromEntries([...app.readBreakdownSteps_(sh)].map(([k, v]) => [k, [...v].map((s) => `${s.step}:${s.min}`)]));

const tests = [];
const test = (name, fn) => tests.push([name, fn]);

test('the hand-built layout is read and its steps carried over', () => {
  const { app, sh } = drawn();
  assert.deepStrictEqual(steps(app, sh), { 'Daily Prep': ['A:0', 'B:0'] });
});

test('groups are ordered by start time, each with a clock-time view of its slot', () => {
  const { sh } = drawn();
  const r = rows(sh);
  const titles = r.filter((row) => String(row[0]).startsWith('§') && row[0] !== '§§').map((row) => row[1]);
  assert.deepStrictEqual(titles, ['Wake-Up', 'Daily Prep']);
  const prep = r.find((row) => row[0] === '§Daily Prep');
  assert.strictEqual(prep[6], '7:00 AM – 7:30 AM · Mon–Fri · 0 of 30 min');
  const wake = r.find((row) => row[0] === '§Wake-Up');
  assert.strictEqual(wake[6], '6:30 AM – 7:00 AM · Mon–Fri · 0 of 30 min');
});

test('typing a step into the empty row adds it and redraws the view', () => {
  const { app, sh } = drawn();
  const empty = findRow(sh, 'Daily Prep', '');                          // the always-there empty step row
  assert.ok(empty > 0);
  // Fill in minutes for A and B, then add a third step
  type(app, sh, findRow(sh, 'Daily Prep', 'A'), 3, [10]);
  type(app, sh, findRow(sh, 'Daily Prep', 'B'), 3, [15]);
  type(app, sh, findRow(sh, 'Daily Prep', ''), 2, ['Plan the day', 5]);
  assert.deepStrictEqual(steps(app, sh), { 'Daily Prep': ['A:10', 'B:15', 'Plan the day:5'] });

  const r = rows(sh);
  const title = r.findIndex((row) => row[0] === '§Daily Prep');
  assert.strictEqual(r[title][2], '30 min');
  assert.strictEqual(r[title][6], '7:00 AM – 7:30 AM · Mon–Fri · 30 of 30 min');
  // View rows are 5 minutes each: A at 7:00, B at 7:10 (row 3), Plan the day at 7:25 (row 6)
  assert.deepStrictEqual([r[title + 1][4], r[title + 1][6]], ['7:00 AM', 'A']);
  assert.deepStrictEqual([r[title + 3][4], r[title + 3][6]], ['7:10 AM', 'B']);
  assert.deepStrictEqual([r[title + 6][4], r[title + 6][6]], ['7:25 AM', 'Plan the day']);
  assert.ok(findRow(sh, 'Daily Prep', '') > 0, 'a new empty row is always kept');
});

test('unused time in the slot is shown as free', () => {
  const { app, sh } = drawn();
  type(app, sh, findRow(sh, 'Daily Prep', 'A'), 3, [20]);
  const r = rows(sh);
  const title = r.findIndex((row) => row[0] === '§Daily Prep');
  assert.deepStrictEqual([r[title + 5][4], r[title + 5][6]], ['7:20 AM', 'free · 10 min']);
});

test('steps past the end of the slot are flagged', () => {
  const { app, book, sh } = drawn();
  type(app, sh, findRow(sh, 'Daily Prep', 'A'), 3, [20]);
  type(app, sh, findRow(sh, 'Daily Prep', 'B'), 3, [15]);
  const r = rows(sh);
  const title = r.findIndex((row) => row[0] === '§Daily Prep');
  assert.strictEqual(r[title][6], '7:00 AM – 7:30 AM · Mon–Fri · 35 of 30 min');
  const cells = book.dump().sheets.Breakdowns.cells;
  assert.strictEqual(cells[`G${title + 1}`].fmt.color, '#CC0000', 'title turns red');
  assert.strictEqual(cells[`F${title + 1 + 7}`].fmt.bg, '#E06666', 'the 7th view row (7:30) is past the slot');
  assert.strictEqual(cells[`F${title + 1 + 6}`].fmt.bg, '#F9CB9C', 'the 6th view row is still inside it');
});

test('a row inserted without a key joins the group above it', () => {
  const { app, sh } = drawn();
  const b = findRow(sh, 'Daily Prep', 'B');
  sh.getRange(b, 1, 1, 3).setValues([['', 'Inserted', 5]]);           // as if typed into a blank inserted row
  app.drawBreakdowns_();
  assert.deepStrictEqual(steps(app, sh), { 'Daily Prep': ['A:0', 'Inserted:5'] });
});

test('unchecking keeps the steps hidden; checking again brings them back', () => {
  const { app, book, sh } = drawn();
  type(app, sh, findRow(sh, 'Daily Prep', 'A'), 3, [10]);
  const catalog = book.getSheetByName('Catalog');
  type(app, catalog, 5, 6, [false]);                                    // Daily Prep's Breakdown box

  const r = rows(sh);
  const marker = r.findIndex((row) => row[0] === '§§') + 1;
  const hiddenTitle = r.findIndex((row) => row[0] === '§Daily Prep') + 1;
  assert.ok(hiddenTitle > marker, 'Daily Prep now sits below the hidden marker');
  const hiddenRows = book.dump().sheets.Breakdowns.hiddenRows;
  assert.ok(hiddenRows.includes(marker) && hiddenRows.includes(hiddenTitle + 2), 'the holding area is hidden');

  type(app, catalog, 5, 6, [true]);
  assert.deepStrictEqual(steps(app, sh), { 'Daily Prep': ['A:10', 'B:0'] });
  const markerNow = rows(sh).findIndex((row) => row[0] === '§§') + 1;
  assert.ok(book.dump().sheets.Breakdowns.hiddenRows.every((n) => n >= markerNow), 'only the holding area is hidden');
  const back = rows(sh).findIndex((row) => row[0] === '§Daily Prep') + 1;
  assert.ok(back < rows(sh).findIndex((row) => row[0] === '§§') + 1, 'Daily Prep is back above the marker');
});

test('moving an item on the Week Plan moves its view', () => {
  const { app, book, sh } = drawn();
  const plan = book.getSheetByName('Week Plan');
  type(app, plan, 3, 4, ['7:15 AM', '7:45 AM']);                       // Daily Prep row
  const title = rows(sh).find((row) => row[0] === '§Daily Prep');
  assert.strictEqual(title[6], '7:15 AM – 7:45 AM · Mon–Fri · 0 of 30 min');
});

test('an item with two different slots gets two views side by side', () => {
  const { sh } = drawn(LEGACY, (tabs) => tabs['Week Plan'].push(['Wake-Up', '', 'Sat-Sun', '8:00 AM', '8:30 AM', 'Weekend - Home', '']));
  const wake = rows(sh).find((row) => row[0] === '§Wake-Up');
  assert.strictEqual(wake[6], '6:30 AM – 7:00 AM · Mon–Fri · 0 of 30 min');
  assert.strictEqual(wake[11], '8:00 AM – 8:30 AM · Sat–Sun · 0 of 30 min');
});

test('redrawing an unchanged tab changes nothing', () => {
  const { app, book, sh } = drawn();
  type(app, sh, findRow(sh, 'Daily Prep', 'A'), 3, [10]);
  const before = JSON.stringify(book.dump().sheets.Breakdowns);
  app.drawBreakdowns_();
  assert.strictEqual(JSON.stringify(book.dump().sheets.Breakdowns), before);
});

test('edits elsewhere on Breakdowns or on other tabs do not redraw', () => {
  const { app, book } = drawn();
  const cats = book.getSheetByName('Categories');
  let calls = 0;
  app.drawBreakdowns_ = () => { calls++; };
  app.onEdit({ range: cats.getRange(2, 2) });
  app.onEdit({ range: book.getSheetByName('Breakdowns').getRange(6, 7) });   // a view cell
  app.onEdit({ range: book.getSheetByName('Breakdowns').getRange(1, 2) });   // the title
  assert.strictEqual(calls, 0);
});

run(tests);
