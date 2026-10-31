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

test('labels are boxes merged down the item, outlined in a darker shade, as on the hand-built grid', () => {
  const { cells, dump } = drawn();
  assert.ok(dump.merges.includes('11,4,17,4'), 'Early Morning Block label spans 7:30–9:30 (rows 11–17)');
  assert.ok(dump.merges.includes('57,8,67,8'), 'D&D label spans 7–10 PM');
  assert.ok(!dump.merges.some((m) => m.startsWith('7,4,')), 'one-chunk Wake-Up is a single cell, not merged');
  assert.deepStrictEqual(cells.D11.fmt.border.slice(0, 4), [true, true, true, true], 'outlined on all four sides');
  assert.strictEqual(cells.D11.fmt.border[6], '#3183CE', 'a darker blue than the #9FC5E8 bars');
  assert.strictEqual(cells.D7.fmt.border[6], '#ECB513', 'Wake-Up: darker yellow');
  assert.strictEqual(cells.D11.fmt.wrap, true);
});

test('bars run the full length on both sides', () => {
  const { cells } = drawn();
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

test('the tab is exactly the grid: 80 rows, 52 columns; heights and widths match the hand-built grid', () => {
  const { dump } = drawn();
  assert.deepStrictEqual(dump.size, [80, 52]);
  for (let r = 1; r <= 80; r++) assert.strictEqual(dump.rowHeights[r], r % 2 ? 21 : 5, `row ${r}`);
  const widths = [100, 5, 25, 100, 25, 5, 25, 100, 25, 5, 5, 5, 25, 100, 25, 5, 25, 100, 25, 5, 5, 5, 25];
  widths.forEach((px, i) => assert.strictEqual(dump.colWidths[i + 1], px, `column ${i + 1}`));
});

test('overlaps on the same side of a day are drawn and reported', () => {
  const { cells, problems } = drawn((tabs) => tabs['Week Plan'].push(['Coffee Walk', 'Late Coffee', 'Tue', '12:30 PM', '1:30 PM', 'Weekday', '']));
  assert.strictEqual(cells.R29.v, 'Lunch · Late Coffee', 'the overlapping label joins the box it starts in');
  assert.strictEqual(cells.Q33.fmt.bg, '#F9CB9C', 'its bars are still drawn through 1:30');
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

test('a Week Plan edit redraws every drawn tab', () => {
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
  assert.ok(book.getSheetByName('Weekends'), 'Weekends drawn too');
});

// ---------------------------------------------------------------- Touchstones

/** The six touchstones from the hand-built sheet, as Catalog items (one per event) and Week Plan rows. */
const TOUCHSTONES = (tabs) => {
  tabs.Catalog.push(
    ['Sam Wakes Up', 'Touchstone', '', '#6AA84F', '', false, ''],
    ['Kate Wakes Up', 'Touchstone', '', '#A64D79', '', false, ''],
    ['Kate Leaves', 'Touchstone', '', '#A64D79', '', false, ''],
    ['Sam Leaves', 'Touchstone', '', '#B6D7A8', '', false, ''],
    ['Kate Home', 'Touchstone', '', '#A64D79', '', false, ''],
  );
  tabs['Week Plan'].push(
    ['Sam Wakes Up', '', 'Mon-Fri', '9:00 AM', '', 'Weekday', ''],
    ['Kate Wakes Up', '', 'Mon-Fri', '11:00 AM', '', 'Weekday', ''],
    ['Kate Leaves', '', 'Mon-Fri', '2:00 PM', '', 'Weekday', ''],
    ['Sam Leaves', '', 'Thu', '5:30 PM', '', 'Weekday', ''],
    ['Kate Home', '', 'Mon, Thu', '8:30 PM', '', 'Weekday', ''],
    ['Kate Home', '', 'Tue', '9:30 PM', '', 'Weekday', ''],
  );
};

test('touchstones: notes on the same cells as the hand-built sheet, with no problems', () => {
  const { cells, problems, app } = drawn(TOUCHSTONES);
  const notes = Object.fromEntries(Object.entries(cells).filter(([, c]) => c.note).map(([a1, c]) => [a1, c.note]));
  assert.deepStrictEqual(notes, {
    A17: 'Sam Wakes Up', A25: 'Kate Wakes Up', A37: 'Kate Leaves', AG51: 'Sam Leaves',
    C63: 'Kate Home', AG63: 'Kate Home', M67: 'Kate Home',
  });
  assert.deepStrictEqual(problems, []);
  assert.deepStrictEqual([...app.readSchedule_().problems], [], 'End can be left blank');
});

test('touchstones: an every-day line spans the whole week from the hour labels; a one-day line spans its day', () => {
  const { cells } = drawn(TOUCHSTONES);
  const top = (a1) => cells[a1] && cells[a1].fmt.border && cells[a1].fmt.border[0] === true ? cells[a1].fmt.border[6] : null;
  assert.deepStrictEqual(['A17', 'H17', 'AW17'].map(top), ['#6AA84F', '#6AA84F', '#6AA84F'], 'Sam Wakes Up: A through Friday');
  assert.deepStrictEqual(['AG51', 'AM51', 'AN51', 'AF51'].map(top), ['#B6D7A8', '#B6D7A8', null, null], 'Sam Leaves: Thursday only');
});

test('touchstones are not blocks: no bars, no overlaps, no Breakdowns group', () => {
  const { cells, problems } = drawn((tabs) => {
    TOUCHSTONES(tabs);
    tabs.Catalog[tabs.Catalog.length - 1][5] = true;                        // Kate Home with Breakdown checked
  });
  assert.strictEqual(cells.C63.fmt.bg, '#FFFFFF', 'no bar drawn for Kate Home');
  assert.deepStrictEqual(problems, []);
});

// ---------------------------------------------------------------- Weekends

/** Draw Weekends with a few weekend rows added to the real data. */
function weekends() {
  const tabs = fixture();
  tabs['Week Plan'].push(
    ['Wake-Up', '', 'Sat-Sun', '8:00 AM', '8:30 AM', 'Weekend - Home', ''],
    ['Block', 'Event', 'Fri', '6:00 PM', '11:00 PM', 'Weekend - Travel', ''],
    ['Coffee Walk', '', 'Sat', '10:00 AM', '10:30 AM', 'Weekend - Travel', ''],
  );
  const { app, book } = setUp(tabs);
  const problems = [...app.drawWeekends_(app.readSchedule_())];
  return { problems, dump: book.dump().sheets.Weekends, schedule: app.readSchedule_() };
}

test('Weekends: home Sat–Sun then travel Fri–Sun, side by side', () => {
  const { dump, problems, schedule } = weekends();
  assert.deepStrictEqual([...schedule.problems], [], 'travel Friday is allowed');
  assert.deepStrictEqual(problems, []);
  const c = dump.cells;
  // Travel days start 3 columns later than they would without the section break (W → Z)
  assert.deepStrictEqual(['C1', 'M1', 'Z1', 'AJ1', 'AT1'].map((a) => c[a].v),
    ['Saturday · Home', 'Sunday · Home', 'Friday · Travel', 'Saturday · Travel', 'Sunday · Travel']);
  assert.deepStrictEqual(labels(c), {
    'Wake-Up': ['D13', 'N13'],              // 8:00 AM on both home days
    'Event': ['AA53'],                       // travel Friday 6 PM
    'Coffee Walk': ['AO21'],                 // travel Saturday 10 AM
  });
  assert.deepStrictEqual(dump.size, [80, 55]);
});

test('Weekends: thin line, big gap, thin line between Home and Travel', () => {
  const { dump } = weekends();
  // Sunday · Home ends with its own thin line (U) and gap (V); then W white 40px, X dark 5px, Y white 5px; Travel at Z
  assert.strictEqual(dump.cells.U40.fmt.bg, '#434343', "Sunday · Home's own thin line stays");
  assert.deepStrictEqual([23, 24, 25].map((col) => dump.colWidths[col]), [40, 5, 5]);
  assert.strictEqual(dump.cells.X40.fmt.bg, '#434343');
  assert.strictEqual(dump.cells.W40.fmt.bg, '#FFFFFF');
  assert.strictEqual(dump.cells.Y40.fmt.bg, '#FFFFFF');
});

test('weekday rows stay off Weekends, and weekend rows stay off Weekdays', () => {
  const { dump } = weekends();
  assert.ok(!Object.values(dump.cells).some((cell) => cell.v === 'Early Morning Block'));
});

run(tests);
