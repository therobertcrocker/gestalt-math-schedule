// Checks that renaming an item on the Catalog carries the new name to the Week Plan and Steps, and that it
// stays put when carrying it would merge two items or Google didn't report the old name.
'use strict';

const assert = require('assert');
const { fixture, setUp, run } = require('./helpers');

/** Fixture drawn once, with Daily Prep and Block steps on Steps. */
function drawn() {
  const tabs = fixture();
  tabs.Steps = [['Item', 'Step', 'Min'], ['Daily Prep'], ['', 'Shower', 10], ['', ''], ['Block', 'Inbox', 15]];
  const env = setUp(tabs);
  env.app.drawAll_(false);
  return env;
}

/** Type one value into a Catalog Item cell the way a person would: Google passes the old and new value. */
function renameRow(app, book, row, value) {
  const range = book.getSheetByName('Catalog').getRange(row, 1);
  const oldValue = range.getValue();
  range.setValue(value);
  app.onEdit({ range, value, oldValue });
}

const column = (book, tab) => book.getSheetByName(tab).getRange(2, 1, book.getSheetByName(tab).getLastRow() - 1, 1).getValues().map((r) => r[0]);

const tests = [];
const test = (name, fn) => tests.push([name, fn]);

test('renaming Block carries to every Week Plan row and to Steps; labels stay', () => {
  const { app, book } = drawn();
  renameRow(app, book, 3, 'Focus Block');
  assert.strictEqual(column(book, 'Week Plan').filter((n) => n === 'Focus Block').length, 4);
  assert.ok(!column(book, 'Week Plan').includes('Block'));
  assert.ok(column(book, 'Steps').includes('Focus Block'));
  assert.deepStrictEqual([...app.readSchedule_().problems], []);
  assert.strictEqual(book.dump().sheets.Weekdays.cells.D11.v, 'Early Morning Block', 'labels are names of their own');
});

test('renaming a Breakdown item keeps its steps and its Breakdowns group', () => {
  const { app, book } = drawn();
  renameRow(app, book, 5, 'Morning Prep');
  assert.ok(column(book, 'Steps').includes('Morning Prep'));
  const { steps } = app.readSchedule_();
  assert.deepStrictEqual([...steps.get('Morning Prep')].map((s) => s.step), ['Shower']);
  const view = book.getSheetByName('Breakdowns');
  assert.ok(view.getRange(1, 2, view.getLastRow(), 1).getValues().some((r) => r[0] === 'Morning Prep'));
});

test('no carry when the new name already belongs to another item', () => {
  const { app, book } = drawn();
  renameRow(app, book, 3, 'Lunch');                               // Block → Lunch would merge the two
  assert.ok(column(book, 'Week Plan').includes('Block'));
  assert.ok(app.readSchedule_().problems.some((p) => /listed twice/.test(p.message)));
});

test('no carry for a paste over several cells (Google gives no old values)', () => {
  const { app, book } = drawn();
  const range = book.getSheetByName('Catalog').getRange(3, 1, 2, 1);
  range.setValues([['Focus Block'], ['Private Lessons']]);
  app.onEdit({ range });
  assert.ok(column(book, 'Week Plan').includes('Block'));
});

test('clearing a name is not a rename', () => {
  const { app, book } = drawn();
  renameRow(app, book, 3, '');
  assert.ok(column(book, 'Week Plan').includes('Block'));
});

run(tests);
