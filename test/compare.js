#!/usr/bin/env node
// Behavior check: run the scenario on the working copy's src/ and on another git ref's src/ (default origin/main),
// then report every difference in the resulting spreadsheets. No differences = the change didn't alter behavior.
//
//   npm run compare               # vs origin/main
//   npm run compare -- HEAD~1     # vs any git ref
'use strict';

const { execSync } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { run } = require('./scenario');

const ref = process.argv[2] || 'origin/main';
const root = path.resolve(__dirname, '..');
const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'gestalt-schedule-'));
for (const f of execSync(`git ls-tree --name-only ${ref} src/`, { cwd: root }).toString().trim().split('\n')) {
  if (f.endsWith('.js')) fs.writeFileSync(path.join(tmp, path.basename(f)), execSync(`git show ${ref}:${f}`, { cwd: root }));
}

const before = run(tmp), after = run(path.join(root, 'src'));
const diffs = [], LIMIT = 200;   // stop collecting after this many (a layout change can differ in every cell)
const walk = (a, b, where) => {
  if (diffs.length >= LIMIT || JSON.stringify(a) === JSON.stringify(b)) return;
  if (a && b && typeof a === 'object' && typeof b === 'object' && !Array.isArray(a)) {
    for (const k of new Set([...Object.keys(a), ...Object.keys(b)])) walk(a[k], b[k], where + ' › ' + k);
  } else {
    diffs.push(`${where}\n    ${ref}: ${JSON.stringify(a)}\n    working copy: ${JSON.stringify(b)}`);
  }
};
before.snaps.forEach(([name, s], i) => walk(s, after.snaps[i] && after.snaps[i][1], name));
walk(before.menus, after.menus, 'menus');

const cells = before.snaps.reduce((n, [, s]) => n + Object.values(s.sheets).reduce((m, sh) => m + Object.keys(sh.cells).length, 0), 0);
console.log(`Compared ${before.snaps.length} checkpoints (${cells} cell states) against ${ref}.`);
if (!diffs.length) { console.log('✓ No differences.'); process.exit(0); }
console.log(`✗ ${diffs.length >= LIMIT ? LIMIT + '+' : diffs.length} difference(s):\n`);
diffs.slice(0, 40).forEach((d) => console.log('  ' + d));
if (diffs.length > 40) console.log(`  …and ${diffs.length - 40} more`);
process.exit(1);
