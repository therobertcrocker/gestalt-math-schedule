// The Steps tab: the one place breakdown steps are typed. It's a plain grouped list. A heading row holds an item's
// name in column A; the rows under it hold that item's steps (Step, Min) until the next heading.
//
// The script never moves or clears anything here. The only writes are:
//  - creating the tab the first time (and moving steps over from the first Breakdowns layout), and
//  - appending a heading at the bottom when an item gets its Breakdown box checked and has no heading yet.

// ---------------------------------------------------------------- reading

/**
 * Steps tab → { steps: Map item name → [{ step, min, row }], headings: Set of item names with a heading }.
 * A step row with an item name in column A (typed the "one name per row" way) also works. Problems: a heading
 * for an item that isn't on the Catalog, and steps with no heading above them.
 */
function readSteps_(catalog, problems) {
  const steps = new Map();
  const headings = new Set();
  const sh = SpreadsheetApp.getActive().getSheetByName(TABS.steps);
  if (!sh) return { steps, headings };

  const c = STEPS.col;
  let current = null;
  readRows_(TABS.steps, STEPS.firstRow, 3).forEach(r => {
    const v = r.values;
    const name = text_(v[c.item - 1]);
    const step = text_(v[c.step - 1]);
    const min = minutes_(v[c.min - 1]);
    if (name && !step && min === null) {                       // a heading
      current = name;
      if (!headings.has(name) && !catalog.has(name)) problems.push({ tab: TABS.steps, row: r.row, message: `${name}: isn't on the Catalog` });
      headings.add(name);
      return;
    }
    if (!step && min === null) return;                          // a blank row
    const item = name || current;
    if (!item) return problems.push({ tab: TABS.steps, row: r.row, message: `${step}: no item heading above this step` });
    if (!steps.has(item)) steps.set(item, []);
    steps.get(item).push({ step, min: min || 0, row: r.row });
  });
  return { steps, headings };
}

/** A Min cell → whole minutes (≥ 0), or null when blank. "10", 10 and "10 min" all read as 10. */
function minutes_(value) {
  if (value === '' || value === null || value === undefined) return null;
  const n = typeof value === 'number' ? value : parseFloat(String(value));
  return Number.isFinite(n) && n >= 0 ? Math.round(n) : 0;
}

// ---------------------------------------------------------------- the two writes

/**
 * Make sure the Steps tab exists. The first time, create it and move over any steps typed on the first
 * Breakdowns layout (including ones parked in its hidden holding area), so nothing typed is lost.
 */
function ensureStepsTab_() {
  const ss = SpreadsheetApp.getActive();
  if (ss.getSheetByName(TABS.steps)) return;
  const sh = ss.insertSheet(TABS.steps);
  setUpStepsTab_(sh);
  const old = ss.getSheetByName(TABS.breakdowns);
  const legacy = old ? readLegacyBreakdownSteps_(old) : new Map();
  const rows = [];
  legacy.forEach((steps, name) => {
    if (rows.length) rows.push(['', '', '']);
    rows.push([name, '', '']);
    steps.forEach(s => rows.push(['', s.step, s.min]));
  });
  if (rows.length) sh.getRange(STEPS.firstRow, 1, rows.length, 3).setValues(rows);
}

/**
 * Append a heading for every item with Breakdown checked that has none yet, earliest start first, so there's
 * somewhere to type its steps. Updates `schedule.stepHeadings` to match.
 */
function addMissingStepHeadings_(schedule) {
  const missing = [...schedule.catalog.values()]
    .filter(item => item.breakdown && item.side !== TOUCHSTONE && !schedule.stepHeadings.has(item.name))
    .sort((a, b) => firstStart_(a.name, schedule.plan) - firstStart_(b.name, schedule.plan));
  if (!missing.length) return;

  const sh = SpreadsheetApp.getActive().getSheetByName(TABS.steps);
  const last = Math.max(sh.getLastRow(), STEPS.firstRow - 1);
  const rows = [];
  missing.forEach(item => {
    if (last >= STEPS.firstRow || rows.length) rows.push(['', '', '']);
    rows.push([item.name, '', '']);
    schedule.stepHeadings.add(item.name);
  });
  ensureSheetSize_(sh, last + rows.length, 3);
  sh.getRange(last + 1, 1, rows.length, 3).setValues(rows);
}

/** Earliest start of an item on the Week Plan (Infinity when it has none). */
function firstStart_(name, plan) {
  return Math.min(Infinity, ...plan.filter(p => p.item === name).map(p => p.start));
}

// ---------------------------------------------------------------- setup

/** Headers, widths, the item dropdown and the heading-row shading of a new Steps tab. */
function setUpStepsTab_(sh) {
  const c = STEPS.col;
  sh.getRange(1, 1, 1, 3).setValues([['Item', 'Step', 'Min']]).setFontWeight('bold').setBackground(COLOR.titleBand);
  sh.setFrozenRows(1);
  sh.setColumnWidth(c.item, STEPS.widths.item);
  sh.setColumnWidth(c.step, STEPS.widths.step);
  sh.setColumnWidth(c.min, STEPS.widths.min);
  const body = sh.getMaxRows() - 1;
  sh.getRange(STEPS.firstRow, c.min, body, 1).setHorizontalAlignment('center');
  sh.getRange(STEPS.firstRow, c.item, body, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInRange(SpreadsheetApp.getActive().getRange(`'${TABS.catalog}'!A2:A`), true).setAllowInvalid(true).build());
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND($A${STEPS.firstRow}<>"",$B${STEPS.firstRow}="")`)
      .setBold(true).setBackground(COLOR.titleBand)
      .setRanges([sh.getRange(STEPS.firstRow, 1, body, 3)])
      .build(),
  ]);
}

// ---------------------------------------------------------------- the first Breakdowns layout

/**
 * Steps typed on the first Breakdowns layout → Map item name → [{ step, min }], shown groups first, then the hidden
 * holding area. Column A held a key: "§Item" on titles, "§§" on the holding-area marker, the item name on step
 * rows. Hand-built tabs (item name in A with nothing beside it as a heading) read the same way.
 */
function readLegacyBreakdownSteps_(sh) {
  const legacy = new Map();
  const first = LEGACY_BREAKDOWNS.readFrom;
  if (sh.getLastRow() < first) return legacy;
  let current = null;
  sh.getRange(first, 1, sh.getLastRow() - first + 1, 3).getValues().forEach(v => {
    const key = text_(v[0]);
    const step = text_(v[1]);
    const min = minutes_(v[2]);
    if (key === LEGACY_BREAKDOWNS.hiddenMark) { current = null; return; }
    if (key.startsWith(LEGACY_BREAKDOWNS.titleMark)) { current = key.slice(LEGACY_BREAKDOWNS.titleMark.length); return; }
    if (key && !step && min === null) { current = key; return; }
    if (!step && min === null) return;
    const item = key || current;
    if (!item) return;
    if (!legacy.has(item)) legacy.set(item, []);
    legacy.get(item).push({ step, min: min || 0 });
  });
  return legacy;
}
