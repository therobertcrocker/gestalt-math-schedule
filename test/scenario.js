// A realistic run through the sheet, for `npm run compare`: starts from the real data with steps still on the first
// Breakdowns layout, then edits the way a person would (steps, times, a weekend, a mistake and its fix, a rename,
// unchecking a breakdown, sorting) and snapshots the whole spreadsheet after each step.
'use strict';

const { fixture, setUp } = require('./helpers');

const FIRST_LAYOUT = [
  ['', 'Breakdowns'], ['', 'hint'], ['', 'Step', 'Min'],
  ['§Daily Prep', 'Daily Prep', '15 min'], ['Daily Prep', 'A', 5], ['Daily Prep', 'B', 10], ['Daily Prep'], [],
  ['§§', 'Hidden'], ['§Wake-Up', 'Wake-Up'], ['Wake-Up', 'Take Meds', 2], ['Wake-Up', 'Brush Teeth', 5],
];

/** Run the scenario with the script in srcDir. Returns { snaps: [[name, dump]], menus }. */
function run(srcDir) {
  const tabs = fixture();
  tabs.Breakdowns = FIRST_LAYOUT;
  const { app, book, menus } = setUp(tabs, srcDir);
  const snaps = [];
  const snap = (name) => snaps.push([name, JSON.parse(JSON.stringify(book.dump()))]);
  const sheet = (n) => book.getSheetByName(n);
  const type = (tab, row, col, values, oldValue) => {
    const range = sheet(tab).getRange(row, col, 1, values.length);
    range.setValues([values]);
    app.onEdit({ range, value: values.length === 1 ? values[0] : undefined, oldValue });
  };

  app.onOpen();
  app.refreshAll();
  snap('first refresh (steps move to Steps)');
  type('Steps', 3, 3, [20]);
  type('Steps', 5, 2, ['Plan the day', 5]);
  snap('steps edited');
  type('Week Plan', 15, 4, ['8:00 PM', '11:00 PM']);
  snap('D&D moved');
  type('Week Plan', 16, 1, ['Wake-Up', '', 'Sat-Sun', '8:00 AM', '8:30 AM', 'Weekend - Home', '']);
  type('Week Plan', 17, 1, ['Block', 'Event', 'Fri', '6:00 PM', '11:00 PM', 'Weekend - Travel', '']);
  snap('weekend rows');
  type('Week Plan', 18, 1, ['Lunch', '', 'Someday', '12:00 PM', '1:00 PM', 'Weekday', '']);
  snap('a broken row');
  type('Week Plan', 18, 3, ['Wed']);
  snap('broken row fixed');
  type('Catalog', 11, 1, ['Kate Home', 'Touchstone', '', '#A64D79', '', false, '']);
  type('Week Plan', 19, 1, ['Kate Home', '', 'Mon, Thu', '8:30 PM', '', 'Weekday', '']);
  snap('a touchstone');
  type('Catalog', 3, 1, ['Focus Block'], 'Block');
  snap('Block renamed');
  type('Catalog', 5, 6, [false]);
  snap('Daily Prep unchecked');
  if (app.sortWeekPlan) app.sortWeekPlan();             // older versions don't have it
  snap('Week Plan sorted');
  app.refreshAll();
  snap('refresh everything');
  return { snaps, menus };
}

module.exports = { run };
