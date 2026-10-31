// Shared test setup: a fake spreadsheet holding the given tabs, with every src/ file loaded into one global
// scope (as Apps Script does), plus a tiny test runner.
'use strict';

const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { makeEnvironment } = require('./fake-sheets');

const SRC = path.join(__dirname, '..', 'src');
const FIXTURE = require('./fixtures/schedule.json');

/** A fresh copy of the real data tabs, safe to change in one test. */
function fixture() {
  return JSON.parse(JSON.stringify(FIXTURE));
}

/** Load the script into a fake spreadsheet containing `tabs` ({ name: rows }). Returns { app, book }. */
function setUp(tabs, srcDir = SRC) {
  const env = makeEnvironment(new Date(2026, 9, 31, 12, 0));
  Object.entries(tabs).forEach(([name, rows]) => {
    const sh = env.book.insertSheet(name);
    const width = Math.max(...rows.map((r) => r.length));
    sh.getRange(1, 1, rows.length, width).setValues(rows.map((r) => [...r, ...Array(width - r.length).fill('')]));
  });
  const app = vm.createContext({ ...env.globals });
  fs.readdirSync(srcDir).filter((f) => f.endsWith('.js')).sort()
    .forEach((f) => vm.runInContext(fs.readFileSync(path.join(srcDir, f), 'utf8'), app, { filename: f }));
  return { app, book: env.book, alerts: env.alerts };
}

/** Run a list of [name, fn] tests, print ✓/✗, and set the exit code. */
function run(tests) {
  let failed = 0;
  tests.forEach(([name, fn]) => {
    try { fn(); console.log(`✓ ${name}`); } catch (err) { failed++; console.log(`✗ ${name}\n  ${String(err.stack || err.message).split('\n').slice(0, 8).join('\n  ')}`); }
  });
  console.log(failed ? `\n${failed} of ${tests.length} failed\n` : `\nAll ${tests.length} passed\n`);
  if (failed) process.exitCode = 1;
}

module.exports = { fixture, setUp, run };
