// Checks Stats.js against the real Week Plan: hours per category (Activities taking overlapped block time), open
// time per day, target statuses, and that touchstones, weekend rows and Categories edits behave.
'use strict';

const assert = require('assert');
const { fixture, setUp, run } = require('./helpers');

/** Fixture (optionally changed) drawn once; returns the Stats tab's rows as { label: [values…] } plus the app. */
function drawn(change = () => {}) {
  const tabs = fixture();
  change(tabs);
  const env = setUp(tabs);
  env.app.drawAll_(false);
  return Object.assign(env, { table: table(env.book) });
}

/** The Stats tab as two lookups: totals (label → [hours, target, status]) and perDay (label → [Mon…Fri, Week]). */
function table(book) {
  const sh = book.getSheetByName('Stats');
  const rows = sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).getValues();
  const split = rows.findIndex((r) => r[0] === 'Per day');
  const pick = (from, to, n) => Object.fromEntries(rows.slice(from, to).filter((r) => r[0]).map((r) => [r[0], r.slice(1, 1 + n)]));
  return { totals: pick(4, split, 3), perDay: pick(split + 1, rows.length, 6), head: rows[split] };
}

const tests = [];
const test = (name, fn) => tests.push([name, fn]);

test('week totals: every minute counts once, Wednesday EMDR takes an hour of the Mid-Morning Block', () => {
  const { table: t } = drawn();
  const hours = Object.fromEntries(Object.entries(t.totals).map(([k, v]) => [k, v[0]]));
  assert.deepStrictEqual(hours, {
    'Work': '29h',                                // 3 blocks × 2h × 5 days, minus Wed 11–12
    'Dance': '5h',                                // Privates Tue 2h + Thu 2h, Reverie 1h
    'Health & Self-Care': '9.5h',                 // Wake-Up, Daily Prep, Coffee Walk ½h × 5 each, EMDR 2h
    'Partner Time': '–',
    'Hobbies': '3h',                              // D&D
    'House Projects': '–',
    'Not counted': '3h',                          // Lunch Tue, Thu, Fri
    'Open time': '40.5h',
  });
});

test('per day: columns Mon–Fri plus the week, and open time each day', () => {
  const { table: t } = drawn();
  assert.deepStrictEqual(t.head.slice(0, 7), ['Per day', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Week']);
  assert.deepStrictEqual(t.perDay['Work'], ['6h', '6h', '5h', '6h', '6h', '29h']);
  assert.deepStrictEqual(t.perDay['Open time'], ['6.5h', '6.5h', '10.5h', '7.5h', '9.5h', '40.5h']);
  // 18 hours a day, 90 a week, all accounted for
  const week = Object.values(t.perDay).map((r) => parseFloat(r[5]) || 0).reduce((a, b) => a + b, 0);
  assert.strictEqual(week, 90);
});

test('targets: on target, under and over, with the gap', () => {
  const { table: t } = drawn((tabs) => {
    tabs.Categories[1][1] = 25; tabs.Categories[1][2] = 30;            // Work 25–30h
    tabs.Categories[2][1] = 8;                                        // Dance ≥ 8h
    tabs.Categories[5][2] = 2;                                        // Hobbies ≤ 2h
  });
  assert.deepStrictEqual(t.totals['Work'], ['29h', '25–30h', '✓ on target']);
  assert.deepStrictEqual(t.totals['Dance'], ['5h', '≥ 8h', '▼ 3h under']);
  assert.deepStrictEqual(t.totals['Hobbies'], ['3h', '≤ 2h', '▲ 1h over']);
  assert.deepStrictEqual(t.totals['Partner Time'].slice(0, 3), ['–', '–', '']);
});

test('a Split shares focus block hours; the rest (personal work) is Not counted', () => {
  const { table: t } = drawn((tabs) => { tabs.Catalog[2][7] = 'Work 70%, House Projects 20%'; });   // Block
  const hours = Object.fromEntries(Object.entries(t.totals).map(([k, v]) => [k, v[0]]));
  assert.strictEqual(hours['Work'], '20.3h');                 // 29h × 70%
  assert.strictEqual(hours['House Projects'], '5.8h');        // 29h × 20%
  assert.strictEqual(hours['Not counted'], '5.9h');           // Lunch 3h + 29h × 10%
  assert.deepStrictEqual(t.perDay['Work'], ['4.2h', '4.2h', '3.5h', '4.2h', '4.2h', '20.3h']);
  const week = Object.values(t.perDay).map((r) => parseFloat(r[5]) || 0).reduce((a, b) => a + b, 0);
  assert.strictEqual(Math.round(week * 10) / 10, 90, 'still 90 hours in all');
});

test('a Split that can\'t be used is reported, and the item counts toward its Category', () => {
  const cases = [
    ['Work seventy', 'Block: Split should look like "Work 70%, House Projects 20%"'],
    ['Work 70%, Chores 20%', 'Block: Split: "Chores" isn\'t on Categories'],
    ['Work 80%, House Projects 30%', 'Block: Split adds up to more than 100%'],
  ];
  cases.forEach(([split, message]) => {
    const { app, table: t } = drawn((tabs) => { tabs.Catalog[2][7] = split; });
    assert.deepStrictEqual([...app.readSchedule_().problems].map((p) => p.message), [message]);
    assert.strictEqual(t.totals['Work'][0], '29h');
  });
});

test('touchstones and weekend rows are not counted', () => {
  const { table: t } = drawn((tabs) => {
    tabs.Catalog.push(['Kate Home', 'Touchstone', 'Partner Time', '#A64D79', '', false, '']);
    tabs['Week Plan'].push(['Kate Home', '', 'Mon-Fri', '8:30 PM', '', 'Weekday', '']);
    tabs['Week Plan'].push(['D&D', '', 'Sat', '1:00 PM', '5:00 PM', 'Weekend - Home', '']);
  });
  assert.strictEqual(t.totals['Partner Time'][0], '–');
  assert.strictEqual(t.totals['Hobbies'][0], '3h');
});

test('editing a target on Categories redraws Stats (and only Stats)', () => {
  const { app, book } = drawn();
  const cleared = [];
  ['Breakdowns', 'Weekdays', 'Weekends', 'Stats'].forEach((name) => {
    const sh = book.getSheetByName(name);
    const clear = sh.clear.bind(sh);
    sh.clear = () => { cleared.push(name); return clear(); };
  });
  const range = book.getSheetByName('Categories').getRange(2, 2, 1, 2);
  range.setValues([[25, 30]]);
  app.onEdit({ range });
  assert.deepStrictEqual(cleared, ['Stats']);
  assert.deepStrictEqual(table(book).totals['Work'], ['29h', '25–30h', '✓ on target']);
});

run(tests);
