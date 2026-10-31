// The Breakdowns tab: one group per item with Breakdown checked on the Catalog, ordered by start time.
// On the left the user types steps and minutes; on the right the script draws a view of those steps at clock
// time, one view per distinct time slot the item has on the Week Plan.
//
// The tab is redrawn from scratch every time, after first reading the steps already on it, so nothing typed
// is lost. Items that get unchecked keep their steps in a hidden holding area at the bottom.

// ---------------------------------------------------------------- entry point

/** Read the steps on the tab, then redraw it from the Catalog, the Week Plan and those steps. */
function drawBreakdowns_() {
  const ss = SpreadsheetApp.getActive();
  const schedule = readSchedule_();
  const sh = ss.getSheetByName(TABS.breakdowns) || ss.insertSheet(TABS.breakdowns);
  const saved = readBreakdownSteps_(sh);
  const groups = breakdownGroups_(schedule, saved);
  paintPage_(sh, layoutBreakdowns_(groups));
  reportProblems_(schedule.problems);
}

// ---------------------------------------------------------------- reading what's on the tab

/**
 * The steps currently on the tab → Map of item name → [{ step, min }], in order.
 * A row belongs to the item named in the hidden key column. A row with a blank key (one the user inserted)
 * belongs to the group above it. A row with an item name in column A but no step is treated as a title,
 * which also reads tabs laid out by hand.
 */
function readBreakdownSteps_(sh) {
  const c = BREAKDOWNS.col;
  const saved = new Map();
  const first = BREAKDOWNS.readFrom;
  if (sh.getLastRow() < first) return saved;

  let current = null;
  sh.getRange(first, 1, sh.getLastRow() - first + 1, c.min).getValues().forEach(v => {
    const key = text_(v[c.key - 1]);
    const step = text_(v[c.step - 1]);
    const min = minutes_(v[c.min - 1]);
    if (key === BREAKDOWNS.hiddenMark) { current = null; return; }
    if (key.startsWith(BREAKDOWNS.titleMark)) { current = key.slice(BREAKDOWNS.titleMark.length); return; }
    if (key && !step && min === null) { current = key; return; }   // title written by hand, or an empty step row
    if (!step && min === null) return;
    const item = key || current;
    if (!item) return;
    if (!saved.has(item)) saved.set(item, []);
    saved.get(item).push({ step, min: min || 0 });
  });
  return saved;
}

/** A Min cell → whole minutes (≥ 0), or null when blank. "10", 10 and "10 min" all read as 10. */
function minutes_(value) {
  if (value === '' || value === null || value === undefined) return null;
  const n = typeof value === 'number' ? value : parseFloat(String(value));
  return Number.isFinite(n) && n >= 0 ? Math.round(n) : 0;
}

// ---------------------------------------------------------------- deciding what to draw

/**
 * Split items into groups to show (Breakdown checked, ordered by earliest start) and groups to keep hidden
 * (unchecked, or no longer on the Catalog, but with steps worth keeping).
 */
function breakdownGroups_(schedule, saved) {
  const shown = [];
  const hidden = [];
  schedule.catalog.forEach(item => {
    const steps = saved.get(item.name) || [];
    if (item.breakdown) shown.push({ item, steps, slots: slotsFor_(item.name, schedule.plan) });
    else if (steps.length) hidden.push({ item, steps });
  });
  saved.forEach((steps, name) => {
    if (!schedule.catalog.has(name) && steps.length) hidden.push({ item: { name, color: COLOR.free }, steps });
  });
  const firstStart = g => (g.slots.length ? g.slots[0].start : Infinity);
  shown.sort((a, b) => firstStart(a) - firstStart(b));
  return { shown, hidden };
}

/**
 * The distinct time slots an item has on the Week Plan, earliest first: [{ start, end, days, grids }].
 * An item with no slots gets one view in relative time (start null).
 */
function slotsFor_(name, plan) {
  const slots = new Map();
  plan.filter(p => p.item === name).forEach(p => {
    const key = p.start + '-' + p.end;
    if (!slots.has(key)) slots.set(key, { start: p.start, end: p.end, days: new Set(), grids: new Set() });
    p.days.forEach(d => slots.get(key).days.add(d));
    slots.get(key).grids.add(p.grid);
  });
  return [...slots.values()]
    .map(s => ({ ...s, days: DAYS.filter(d => s.days.has(d)), grids: [...s.grids] }))
    .sort((a, b) => a.start - b.start);
}

// ---------------------------------------------------------------- layout

/**
 * Turn the groups into a page: a list of rows, each a Map of column → cell, where a cell is
 * { value, bg, color, bold, line } (line = a comment line under the cell). Also returns which rows to hide
 * and how many view columns the widest group needs.
 */
function layoutBreakdowns_(groups) {
  const rows = [];
  const add = () => { const row = new Map(); rows.push(row); return row; };
  const viewCount = Math.max(1, ...groups.shown.map(g => Math.max(1, g.slots.length)));

  // Title, hint, headers
  add().set(BREAKDOWNS.col.step, { value: 'Breakdowns', bold: true, size: 14 });
  add().set(BREAKDOWNS.col.step, { value: 'Check Breakdown on the Catalog to add an item. Type steps and minutes on the left; the view redraws itself.', color: COLOR.muted });
  const headers = add();
  headers.set(BREAKDOWNS.col.step, { value: 'Step', bold: true });
  headers.set(BREAKDOWNS.col.min, { value: 'Min', bold: true, halign: 'center' });

  groups.shown.forEach(g => {
    layoutGroup_(g, add);
    add();                                              // blank row between groups
  });

  // Hidden holding area: unchecked items keep their steps here
  const hiddenFrom = rows.length + 1;
  const marker = add();
  marker.set(BREAKDOWNS.col.key, { value: BREAKDOWNS.hiddenMark });
  marker.set(BREAKDOWNS.col.step, { value: 'Hidden: items with Breakdown unchecked (steps kept)', color: COLOR.muted });
  groups.hidden.forEach(g => {
    const title = add();
    title.set(BREAKDOWNS.col.key, { value: BREAKDOWNS.titleMark + g.item.name });
    title.set(BREAKDOWNS.col.step, { value: g.item.name, bold: true });
    g.steps.forEach(s => {
      const row = add();
      row.set(BREAKDOWNS.col.key, { value: g.item.name });
      row.set(BREAKDOWNS.col.step, { value: s.step });
      row.set(BREAKDOWNS.col.min, { value: s.min, halign: 'center' });
    });
  });

  return { rows, hiddenFrom, viewCount };
}

/** One shown group: a title row, then steps on the left and one view per slot on the right, side by side. */
function layoutGroup_(g, add) {
  const c = BREAKDOWNS.col;
  const total = g.steps.reduce((sum, s) => sum + s.min, 0);
  const slots = g.slots.length ? g.slots : [{ start: null, end: null, days: [], grids: [] }];
  const views = slots.map(slot => viewCells_(g, slot, total));
  const bodyRows = Math.max(g.steps.length + 1, ...views.map(v => v.body.length));   // +1: the empty step row

  const title = add();
  title.set(c.key, { value: BREAKDOWNS.titleMark + g.item.name });
  title.set(c.step, { value: g.item.name, bold: true, bg: COLOR.titleBand });
  title.set(c.min, { value: `${total} min`, bg: COLOR.titleBand, halign: 'center', color: COLOR.muted });
  views.forEach((v, k) => v.title.forEach((cell, offset) => title.set(viewColumn_(k, offset), cell)));

  for (let i = 0; i < bodyRows; i++) {
    const row = add();
    row.set(c.key, { value: g.item.name });                  // so a step typed into any of these rows lands here
    const s = g.steps[i];
    if (s) {
      row.set(c.step, { value: s.step });
      row.set(c.min, { value: s.min, halign: 'center' });
    }
    views.forEach((v, k) => (v.body[i] || new Map()).forEach((cell, offset) => row.set(viewColumn_(k, offset), cell)));
  }
}

/**
 * The cells of one view: a title (slot and how full it is) and one body row per 5 minutes, each a Map of
 * column offset within the view → cell. Rows past the end of the slot are flagged; unused rows in the slot are grey.
 */
function viewCells_(g, slot, total) {
  const v = BREAKDOWNS.view;
  const step = BREAKDOWNS.minutesPerRow;
  const relative = slot.start === null;
  const length = relative ? total : slot.end - slot.start;
  const clock = offset => (relative ? relativeTime_(offset) : formatTime_(slot.start + offset));

  const title = new Map();
  const when = relative ? 'Not on the Week Plan' : `${formatTime_(slot.start)} – ${formatTime_(slot.end)} · ${compressDays_(slot.days)}`;
  const fill = relative ? `${total} min` : `${total} of ${length} min`;
  title.set(v.time, { value: '', bg: COLOR.titleBand });
  title.set(v.barLeft, { value: '', bg: g.item.color });
  title.set(v.label, { value: `${when} · ${fill}`, bold: true, bg: COLOR.titleBand, color: total > length ? COLOR.overText : COLOR.text });
  title.set(v.barRight, { value: '', bg: g.item.color });

  // Where each step starts, in minutes from the start of the slot
  const starts = [];
  let at = 0;
  g.steps.forEach(s => { if (s.min > 0) { starts.push({ at, step: s }); at += s.min; } });

  const rows = Math.ceil(Math.max(length, total) / step);
  const body = [];
  for (let i = 0; i < rows; i++) {
    const offset = i * step;
    const over = offset >= length;
    const used = offset < total;
    const bar = over ? COLOR.over : used ? g.item.color : COLOR.free;
    const row = new Map([[v.barLeft, { value: '', bg: bar }], [v.barRight, { value: '', bg: bar }]]);
    const begins = starts.find(s => s.at >= offset && s.at < offset + step);
    if (begins) {
      row.set(v.time, { value: clock(begins.at), color: over ? COLOR.overText : COLOR.muted, halign: 'right' });
      row.set(v.label, { value: begins.step.step, color: over ? COLOR.overText : COLOR.text, line: true });
    } else if (!used && offset === Math.ceil(total / step) * step && !over) {
      row.set(v.time, { value: clock(total), color: COLOR.muted, halign: 'right' });
      row.set(v.label, { value: `free · ${length - total} min`, color: COLOR.muted, line: true });
    }
    body.push(row);
  }
  return { title, body };
}

/** Sheet column of a cell inside the k-th view (0-based). */
function viewColumn_(k, offset) {
  return BREAKDOWNS.col.firstView + k * BREAKDOWNS.view.width + offset;
}

// ---------------------------------------------------------------- painting

/** Write a laid-out page onto the tab in a few batched calls. */
function paintPage_(sh, page) {
  const c = BREAKDOWNS.col;
  const v = BREAKDOWNS.view;
  const w = BREAKDOWNS.widths;
  const nRows = Math.max(page.rows.length, 1);
  const nCols = c.firstView + page.viewCount * v.width - 1;

  clearSheet_(sh);
  ensureSheetSize_(sh, nRows + 20, nCols);

  const grid = (fn) => page.rows.map(row => Array.from({ length: nCols }, (_, j) => fn(row.get(j + 1))));
  const range = sh.getRange(1, 1, page.rows.length, nCols);
  range.setValues(grid(cell => (cell && cell.value !== undefined ? cell.value : '')));
  range.setBackgrounds(grid(cell => (cell && cell.bg) || COLOR.white));
  range.setFontColors(grid(cell => (cell && cell.color) || COLOR.text));
  range.setFontWeights(grid(cell => (cell && cell.bold ? 'bold' : 'normal')));
  range.setHorizontalAlignments(grid(cell => (cell && cell.halign) || 'left'));
  range.setVerticalAlignment('middle');

  // Comment lines under step labels, and the big title
  page.rows.forEach((row, i) => row.forEach((cell, col) => {
    if (cell.line) sh.getRange(i + 1, col).setBorder(null, null, true, null, null, null, COLOR.line, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    if (cell.size) sh.getRange(i + 1, col).setFontSize(cell.size);
  }));

  // Column widths: key, step, min, gap, then each view
  sh.setColumnWidth(c.key, w.key);
  sh.setColumnWidth(c.step, w.step);
  sh.setColumnWidth(c.min, w.min);
  sh.setColumnWidth(c.min + 1, w.gap);
  for (let k = 0; k < page.viewCount; k++) {
    const widths = [w.time, w.bar, w.label, w.bar, w.gap];
    widths.forEach((px, offset) => sh.setColumnWidth(viewColumn_(k, offset), px));
  }

  sh.hideColumns(c.key);
  sh.setFrozenRows(BREAKDOWNS.firstRow - 1);
  if (page.rows.length >= page.hiddenFrom) sh.hideRows(page.hiddenFrom, page.rows.length - page.hiddenFrom + 1);
}

// ---------------------------------------------------------------- small helpers

/** ['Mon', 'Tue', 'Wed', 'Thu', 'Fri'] → "Mon–Fri"; ['Mon', 'Wed', 'Thu'] → "Mon, Wed–Thu". */
function compressDays_(days) {
  const idx = days.map(d => DAYS.indexOf(d)).sort((a, b) => a - b);
  const runs = [];
  idx.forEach(i => {
    const last = runs[runs.length - 1];
    if (last && i === last[1] + 1) last[1] = i; else runs.push([i, i]);
  });
  return runs.map(([a, b]) => (a === b ? DAYS[a] : `${DAYS[a]}–${DAYS[b]}`)).join(', ');
}

/** Minutes from the start of an item → "0:05". */
function relativeTime_(minutes) {
  return `${Math.floor(minutes / 60)}:${String(minutes % 60).padStart(2, '0')}`;
}
