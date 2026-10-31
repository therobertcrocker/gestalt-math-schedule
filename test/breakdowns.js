// Checks Steps.js and Breakdowns.js against a copy of the real data: moving steps over from the first Breakdowns
// layout, headings added for checked items without ever touching typed rows, and the drawn view: clock times,
// free time, overflow, markers, per-item row sizes, multiple slots.
'use strict';

const assert = require('assert');
const { fixture, setUp, run } = require('./helpers');

/** The live Breakdowns tab before the split: Daily Prep shown, Wake-Up parked in the hidden holding area. */
const FIRST_LAYOUT = [
  ['', 'Breakdowns'],
  ['', 'Check Breakdown on the Catalog to add an item. Type steps and minutes on the left; the view redraws itself.'],
  ['', 'Step', 'Min'],
  ['§Daily Prep', 'Daily Prep', '15 min'],
  ['Daily Prep', 'A', 5],
  ['Daily Prep', 'B', 10],
  ['Daily Prep'],
  [],
  ['§§', 'Hidden: items with Breakdown unchecked (steps kept)'],
  ['§Wake-Up', 'Wake-Up'],
  ['Wake-Up', 'Take Meds', 2],
  ['Wake-Up', 'Brush Teeth', 5],
];

/** Fixture tabs plus extra tabs (e.g. { Steps: rows }), with drawAll_ run once. */
function drawn(extra = {}, change = () => {}) {
  const tabs = Object.assign(fixture(), extra);
  change(tabs);
  const { app, book } = setUp(tabs);
  app.drawAll_();
  return { app, book, steps: book.getSheetByName('Steps'), view: book.getSheetByName('Breakdowns') };
}

/** All values on a sheet as plain rows, trailing blanks trimmed. */
function rows(sh) {
  const last = sh.getLastRow();
  if (!last) return [];
  return sh.getRange(1, 1, last, sh.getMaxColumns()).getValues().map((r) => {
    const out = [...r];
    while (out.length && out[out.length - 1] === '') out.pop();
    return out;
  });
}

/**
 * Index (0-based) of an item's slot-title row on Breakdowns: the row right under its name row. View columns:
 * B time, C bar, D label, E bar, F gap; a second view starts at G (label in I).
 */
const titleAt = (r, name) => {
  const nameRow = r.findIndex((row, i) => row[1] === name && row.length === 2 && i >= 2);
  return nameRow < 0 ? -1 : nameRow + 1;
};
const TIME = 1;
const LABEL = 3;

/** Type values into cells the way a person would, then fire onEdit. */
function type(app, sh, row, col, values) {
  const range = sh.getRange(row, col, 1, values.length);
  range.setValues([values]);
  app.onEdit({ range });
}

/** The Steps tab as written: headings as "# Item", steps as "Step:min". */
const stepsTab = (sh) => rows(sh).slice(1).map((r) => (r[0] && !r[1] ? `# ${r[0]}` : r.length ? `${r[1]}:${r[2]}` : ''));

const tests = [];
const test = (name, fn) => tests.push([name, fn]);

// ---------------------------------------------------------------- Steps tab

test('first run: steps move from the first Breakdowns layout to Steps, including the hidden ones', () => {
  const { steps } = drawn({ Breakdowns: FIRST_LAYOUT });
  assert.deepStrictEqual(stepsTab(steps), ['# Daily Prep', 'A:5', 'B:10', '', '# Wake-Up', 'Take Meds:2', 'Brush Teeth:5']);
});

test('a checked item with no heading gets one appended; typed rows are never moved', () => {
  const { app, book, steps } = drawn({ Steps: [['Item', 'Step', 'Min'], ['Daily Prep'], ['', 'Shower', 10]] });
  // Wake-Up is checked in the fixture but had no heading: appended below Daily Prep
  assert.deepStrictEqual(stepsTab(steps), ['# Daily Prep', 'Shower:10', '', '# Wake-Up']);
  // Check Coffee Walk: its heading goes at the bottom, and nothing above changes
  type(app, book.getSheetByName('Catalog'), 6, 6, [true]);
  assert.deepStrictEqual(stepsTab(steps), ['# Daily Prep', 'Shower:10', '', '# Wake-Up', '', '# Coffee Walk']);
});

test('unchecking an item leaves its steps on Steps and only removes its group from Breakdowns', () => {
  const { app, book, steps, view } = drawn({ Steps: [['Item', 'Step', 'Min'], ['Daily Prep'], ['', 'Shower', 10]] });
  type(app, book.getSheetByName('Catalog'), 5, 6, [false]);           // Daily Prep's Breakdown box
  assert.deepStrictEqual(stepsTab(steps).slice(0, 2), ['# Daily Prep', 'Shower:10']);
  assert.strictEqual(titleAt(rows(view), 'Daily Prep'), -1);
  assert.ok(titleAt(rows(view), 'Wake-Up') > 0);
});

test('steps typed one-name-per-row read the same as grouped ones', () => {
  const { app } = drawn({ Steps: [['Item', 'Step', 'Min'], ['Daily Prep', 'Shower', 10], ['Daily Prep', 'Dress', 5]] });
  const { steps } = app.readSchedule_();
  assert.deepStrictEqual([...steps.get('Daily Prep')].map((s) => `${s.step}:${s.min}`), ['Shower:10', 'Dress:5']);
});

test('Steps problems: a heading not on the Catalog, and a step with no heading', () => {
  const { app } = drawn({ Steps: [['Item', 'Step', 'Min'], ['', 'Orphan', 5], ['Nap'], ['', 'Lie down', 20]] });
  const messages = [...app.readSchedule_().problems].map((p) => `${p.tab}:${p.row} ${p.message}`);
  assert.deepStrictEqual(messages, ['Steps:2 Orphan: no item heading above this step', "Steps:3 Nap: isn't on the Catalog"]);
});

test('editing Steps redraws Breakdowns; editing Breakdowns or Categories does nothing', () => {
  const { app, book, steps, view } = drawn({ Steps: [['Item', 'Step', 'Min'], ['Daily Prep'], ['', 'Shower', 10]] });
  type(app, steps, 3, 3, [20]);
  const r = rows(view);
  assert.strictEqual(r[titleAt(r, 'Daily Prep')][LABEL], '7:00 AM – 7:30 AM · Mon–Fri · 20 of 30 min');
  let calls = 0;
  app.drawAll_ = () => { calls++; };
  app.onEdit({ range: view.getRange(6, 4) });
  app.onEdit({ range: book.getSheetByName('Categories').getRange(2, 2) });
  assert.strictEqual(calls, 0);
});

// ---------------------------------------------------------------- the view

/** Draw with Daily Prep's steps set to `list` ([step, min] pairs). */
const prep = (list) => drawn({ Steps: [['Item', 'Step', 'Min'], ['Daily Prep'], ...list.map(([s, m]) => ['', s, m])] });

test('groups are ordered by start time, each titled with its slot', () => {
  const r = rows(prep([]).view);
  assert.ok(titleAt(r, 'Wake-Up') < titleAt(r, 'Daily Prep'));
  assert.strictEqual(r[titleAt(r, 'Wake-Up')][LABEL], '6:30 AM – 7:00 AM · Mon–Fri · 0 of 30 min');
});

test('each group is a name row, a slot title, then steps at clock time', () => {
  const r = rows(prep([['A', 10], ['B', 15], ['Plan the day', 5]]).view);
  const t = titleAt(r, 'Daily Prep');
  assert.deepStrictEqual(r[t - 1], ['', 'Daily Prep'], 'name row, nothing else on it');
  assert.strictEqual(r[t][LABEL], '7:00 AM – 7:30 AM · Mon–Fri · 30 of 30 min');
  assert.deepStrictEqual([r[t + 1][TIME], r[t + 1][LABEL]], ['7:00 AM', 'A']);
  assert.deepStrictEqual([r[t + 3][TIME], r[t + 3][LABEL]], ['7:10 AM', 'B']);
  assert.deepStrictEqual([r[t + 6][TIME], r[t + 6][LABEL]], ['7:25 AM', 'Plan the day']);
});

test('unused time in the slot is shown as free', () => {
  const r = rows(prep([['A', 15], ['B', 5]]).view);
  const t = titleAt(r, 'Daily Prep');
  assert.deepStrictEqual([r[t + 5][TIME], r[t + 5][LABEL]], ['7:20 AM', 'free · 10 min']);
});

test('a step with no minutes is a marker at its moment, sharing a row if needed', () => {
  const r = rows(prep([['A', 20], ['B', ''], ['Meds', '']]).view);
  const t = titleAt(r, 'Daily Prep');
  assert.deepStrictEqual([r[t + 5][TIME], r[t + 5][LABEL]], ['7:20 AM', '◆ B · ◆ Meds · free · 10 min']);
});

test('steps past the end of the slot are flagged', () => {
  const { book, view } = prep([['A', 20], ['B', 15]]);
  const r = rows(view);
  const t = titleAt(r, 'Daily Prep');
  assert.strictEqual(r[t][LABEL], '7:00 AM – 7:30 AM · Mon–Fri · 35 of 30 min');
  const cells = book.dump().sheets.Breakdowns.cells;
  assert.strictEqual(cells[`D${t + 1}`].fmt.color, '#CC0000', 'title turns red');
  assert.strictEqual(cells[`C${t + 1 + 7}`].fmt.bg, '#E06666', 'the 7th view row (7:30) is past the slot');
  assert.strictEqual(cells[`C${t + 1 + 6}`].fmt.bg, '#F9CB9C', 'the 6th view row is still inside it');
});

test('View min 1 gives one row per minute, so short steps each get their own label', () => {
  const r = rows(drawn({ Steps: [['Item', 'Step', 'Min'], ['Wake-Up'], ['', 'Brush teeth', 2], ['', 'Floss', 1], ['', 'Wash face', 3]] }).view);
  const t = titleAt(r, 'Wake-Up');
  const d = titleAt(r, 'Daily Prep');
  assert.strictEqual(d - t - 3, 30, '30 one-minute rows, the blank row between groups, then the next name row');
  assert.deepStrictEqual([r[t + 1][TIME], r[t + 1][LABEL]], ['6:30 AM', 'Brush teeth']);
  assert.deepStrictEqual([r[t + 3][TIME], r[t + 3][LABEL]], ['6:32 AM', 'Floss']);
  assert.deepStrictEqual([r[t + 4][TIME], r[t + 4][LABEL]], ['6:33 AM', 'Wash face']);
  assert.deepStrictEqual([r[t + 7][TIME], r[t + 7][LABEL]], ['6:36 AM', 'free · 24 min']);
});

test('an item with two different slots gets two views side by side', () => {
  const { view } = drawn({}, (tabs) => tabs['Week Plan'].push(['Wake-Up', '', 'Sat-Sun', '8:00 AM', '8:30 AM', 'Weekend - Home', '']));
  const r = rows(view);
  const wake = r[titleAt(r, 'Wake-Up')];
  assert.strictEqual(wake[LABEL], '6:30 AM – 7:00 AM · Mon–Fri · 0 of 30 min');
  assert.strictEqual(wake[LABEL + 5], '8:00 AM – 8:30 AM · Sat–Sun · 0 of 30 min');
});

test('moving an item on the Week Plan moves its view', () => {
  const { app, book, view } = prep([]);
  type(app, book.getSheetByName('Week Plan'), 3, 4, ['7:15 AM', '7:45 AM']);
  const r = rows(view);
  assert.strictEqual(r[titleAt(r, 'Daily Prep')][LABEL], '7:15 AM – 7:45 AM · Mon–Fri · 0 of 30 min');
});

test('Breakdowns is sized exactly to its content, and redrawing changes nothing', () => {
  const { app, book } = prep([['A', 10]]);
  const before = JSON.stringify(book.dump().sheets.Breakdowns);
  const page = app.layoutBreakdowns_(app.shownGroups_(app.readSchedule_()));
  assert.strictEqual(book.dump().sheets.Breakdowns.size[0], page.rows.length, 'no spare rows piling up');
  app.drawAll_();
  assert.strictEqual(JSON.stringify(book.dump().sheets.Breakdowns), before);
});

run(tests);
