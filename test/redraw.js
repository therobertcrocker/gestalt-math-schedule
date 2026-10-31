// Checks that edits only redraw what they change: unchanged tabs are skipped, sizes are only set when a tab's
// shape changes, a skipped-resize repaint ends up identical to drawing from scratch, and Refresh everything
// always redraws.
'use strict';

const assert = require('assert');
const { fixture, setUp, run } = require('./helpers');

const DRAWN = ['Breakdowns', 'Weekdays', 'Weekends'];

/** Fixture drawn once, with every drawn tab's clear() and setRowHeight() counted from then on. */
function drawn(change = () => {}) {
  const tabs = fixture();
  tabs.Steps = [['Item', 'Step', 'Min'], ['Daily Prep'], ['', 'Shower', 10]];
  change(tabs);
  const { app, book } = setUp(tabs);
  app.drawAll_(false);
  const counts = {};
  DRAWN.forEach((name) => {
    const sh = book.getSheetByName(name);
    counts[name] = { clear: 0, heights: 0 };
    const clear = sh.clear.bind(sh);
    const setRowHeight = sh.setRowHeight.bind(sh);
    sh.clear = () => { counts[name].clear++; return clear(); };
    sh.setRowHeight = (...a) => { counts[name].heights++; return setRowHeight(...a); };
  });
  return { app, book, counts };
}

/** Type values into cells the way a person would, then fire onEdit. */
function type(app, sh, row, col, values) {
  const range = sh.getRange(row, col, 1, values.length);
  range.setValues([values]);
  app.onEdit({ range });
}

/** Which drawn tabs were cleared (i.e. redrawn). */
const redrawn = (counts) => DRAWN.filter((n) => counts[n].clear > 0);

const tests = [];
const test = (name, fn) => tests.push([name, fn]);

test('editing a Description or a Notes cell redraws nothing', () => {
  const { app, book, counts } = drawn();
  type(app, book.getSheetByName('Catalog'), 3, 5, ['Deep work']);
  type(app, book.getSheetByName('Week Plan'), 4, 7, ['a note']);
  assert.deepStrictEqual(redrawn(counts), []);
});

test('a weekend-only change redraws Weekends and nothing else', () => {
  const { app, book, counts } = drawn();
  const plan = book.getSheetByName('Week Plan');
  type(app, plan, 16, 1, ['Coffee Walk', '', 'Sat', '9:00 AM', '9:30 AM', 'Weekend - Home', '']);
  assert.deepStrictEqual(redrawn(counts), ['Weekends']);
});

test('moving an item repaints its grid without re-setting row heights', () => {
  const { app, book, counts } = drawn();
  type(app, book.getSheetByName('Week Plan'), 15, 4, ['8:00 PM', '11:00 PM']);      // D&D later
  assert.deepStrictEqual(redrawn(counts), ['Weekdays']);
  assert.strictEqual(counts.Weekdays.heights, 0, 'same shape: heights left alone');
});

test('a repaint that skips resizing ends up identical to drawing the same data from scratch', () => {
  const { app, book } = drawn();
  type(app, book.getSheetByName('Week Plan'), 15, 4, ['8:00 PM', '11:00 PM']);
  type(app, book.getSheetByName('Steps'), 3, 3, [20]);
  const fresh = drawn((tabs) => {
    tabs['Week Plan'][14][3] = '8:00 PM';
    tabs['Week Plan'][14][4] = '11:00 PM';
    tabs.Steps[2][2] = 20;
  });
  DRAWN.forEach((name) => assert.deepStrictEqual(book.dump().sheets[name], fresh.book.dump().sheets[name], name));
});

test('a step edit redraws Breakdowns only', () => {
  const { app, book, counts } = drawn();
  type(app, book.getSheetByName('Steps'), 3, 3, [15]);
  assert.deepStrictEqual(redrawn(counts), ['Breakdowns']);
});

test('Refresh everything redraws every tab, sizes included', () => {
  const { app, counts } = drawn();
  app.refreshAll();
  assert.deepStrictEqual(redrawn(counts), DRAWN);
  assert.ok(counts.Weekdays.heights > 0);
});

test('a drawn tab that was deleted comes back on the next edit', () => {
  const { app, book } = drawn();
  book.deleteSheet(book.getSheetByName('Weekdays'));
  type(app, book.getSheetByName('Catalog'), 3, 5, ['Deep work']);              // a change that alone redraws nothing
  assert.ok(book.getSheetByName('Weekdays'));
  assert.strictEqual(book.dump().sheets.Weekdays.cells.D11.v, 'Early Morning Block');
});

run(tests);
