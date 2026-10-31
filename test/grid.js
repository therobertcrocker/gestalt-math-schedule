// Checks Grid.js: the Weekdays tab drawn from a copy of the real data should match the hand-built
// Schedule Planning tab it replaces — same labels in the same cells, same bars, same row and column sizes.
'use strict';

const assert = require('assert');
const { fixture, setUp, run } = require('./helpers');

/** Every item label on the hand-built grid (with D&D at its corrected 7 PM), by cell. */
const HAND_BUILT = {
  'Wake-Up': ['D7', 'N7', 'X7', 'AH7', 'AR7'],
  'Daily Prep': ['H9', 'R9', 'AB9', 'AL9', 'AV9'],
  'Early Morning Block': ['D11', 'N11', 'X11', 'AH11', 'AR11'],
  'Coffee Walk': ['H19', 'R19', 'AB19', 'AL19', 'AV19'],
  'Mid-Morning Block': ['D21', 'N21', 'X21', 'AH21', 'AR21'],
  'EMDR': ['AB25', 'H29'],
  'Lunch': ['R29', 'AL29', 'AV29'],
  'Mid Day Block': ['D33', 'N33', 'AH33', 'AR33'],
  'Afternoon Block': ['X41'],
  'Privates Block': ['N45', 'AH53'],
  'Reverie': ['R53'],
  'D&D': ['H57'],
};

/** Draw Weekdays from the fixture (optionally changed first). Returns { app, book, cells, dump, problems }. */
function drawn(change = () => {}) {
  const tabs = fixture();
  change(tabs);
  const { app, book } = setUp(tabs);
  const problems = [...app.drawWeekdays_(app.readSchedule_())];
  const dump = book.dump().sheets.Weekdays;
  return { app, book, dump, cells: dump.cells, problems };
}

/** Label cells: every non-empty value below the headers, outside the hour column. */
function labels(cells) {
  const out = {};
  Object.entries(cells).forEach(([a1, c]) => {
    const [, col, row] = a1.match(/^([A-Z]+)(\d+)$/);
    if (c.v === undefined || +row < 5 || col === 'A') return;
    (out[c.v] = out[c.v] || []).push(a1);
  });
  Object.values(out).forEach((list) => list.sort());
  return out;
}

const tests = [];
const test = (name, fn) => tests.push([name, fn]);

test('every label lands in the same cell as on the hand-built grid', () => {
  const { cells } = drawn();
  const sorted = Object.fromEntries(Object.entries(HAND_BUILT).map(([k, v]) => [k, [...v].sort()]));
  assert.deepStrictEqual(labels(cells), sorted);
});

test('labels carry a comment line; bars run the full length on both sides', () => {
  const { cells } = drawn();
  assert.strictEqual(cells.D11.fmt.border[2], true, 'Early Morning Block label has a bottom line');
  // Early Morning Block on Monday: 7:30–9:30 → chunk rows 11, 13, 15, 17 plus the dividers between
  for (let r = 11; r <= 17; r++) {
    assert.strictEqual(cells[`C${r}`].fmt.bg, '#9FC5E8', `left bar, row ${r}`);
    assert.strictEqual(cells[`E${r}`].fmt.bg, '#9FC5E8', `right bar, row ${r}`);
  }
  assert.strictEqual(cells.C18.fmt.bg, '#FFFFFF', 'the divider after the block stays white');
  // D&D, Monday 7–10 PM: rows 57 through 67
  assert.strictEqual(cells.G57.fmt.bg, '#B4A7D6');
  assert.strictEqual(cells.G67.fmt.bg, '#B4A7D6');
  assert.strictEqual(cells.G68.fmt.bg, '#FFFFFF');
});

test('headers, hour labels and separators are where they were', () => {
  const { cells, dump } = drawn();
  assert.deepStrictEqual(['C1', 'M1', 'W1', 'AG1', 'AQ1'].map((a) => cells[a].v), ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']);
  assert.deepStrictEqual([cells.C3.v, cells.G3.v], ['Schedule', 'Activities']);
  assert.deepStrictEqual([cells.A5.v, cells.A9.v, cells.A57.v, cells.A77.v], ['6:00 AM', '7:00 AM', '7:00 PM', '12:00 AM']);
  assert.ok(dump.merges.includes('5,1,7,1'), 'hour label spans chunk · divider · chunk');
  assert.strictEqual(cells.K40.fmt.bg, '#434343', 'dark separator after Monday');
});

test('row heights alternate 21/5 and column widths match the hand-built grid', () => {
  const { dump } = drawn();
  assert.strictEqual(Object.keys(dump.rowHeights).length >= 80, true);
  for (let r = 1; r <= 80; r++) assert.strictEqual(dump.rowHeights[r], r % 2 ? 21 : 5, `row ${r}`);
  const widths = [100, 5, 25, 100, 25, 5, 25, 100, 25, 5, 5, 5, 25, 100, 25, 5, 25, 100, 25, 5, 5, 5, 25];
  widths.forEach((px, i) => assert.strictEqual(dump.colWidths[i + 1], px, `column ${i + 1}`));
});

test('overlaps on the same side of a day are drawn and reported', () => {
  const { cells, problems } = drawn((tabs) => tabs['Week Plan'].push(['Coffee Walk', 'Late Coffee', 'Tue', '12:30 PM', '1:30 PM', 'Weekday', '']));
  assert.strictEqual(cells.R31.v, 'Late Coffee');
  assert.deepStrictEqual(problems.map((p) => `${p.tab}:${p.row} ${p.message}`), ['Week Plan:16 Late Coffee overlaps Lunch on Tue (Activities)']);
});

test('items on the other side of the same time do not count as overlapping', () => {
  const { problems } = drawn();   // e.g. Wednesday: Mid-Morning Block (Schedule) and EMDR (Activities) at 11
  assert.deepStrictEqual(problems, []);
});

test('weekend rows stay off the weekday grid', () => {
  const { cells } = drawn((tabs) => tabs['Week Plan'].push(['Coffee Walk', 'Sat coffee', 'Sat', '9:00 AM', '9:30 AM', 'Weekend - Home', '']));
  assert.ok(!Object.values(cells).some((c) => c.v === 'Sat coffee'));
});

test('a Week Plan edit redraws both drawn tabs', () => {
  const tabs = fixture();
  tabs.Breakdowns = [['Item', 'Step', 'Min']];
  const { app, book } = setUp(tabs);
  const plan = book.getSheetByName('Week Plan');
  const range = plan.getRange(15, 4, 1, 2);
  range.setValues([['8:00 PM', '11:00 PM']]);                   // move D&D later
  app.onEdit({ range });
  const cells = book.dump().sheets.Weekdays.cells;
  assert.strictEqual(cells.H61.v, 'D&D');
  assert.ok(book.getSheetByName('Breakdowns'), 'Breakdowns drawn too');
});

run(tests);
