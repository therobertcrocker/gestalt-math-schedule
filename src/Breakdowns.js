// The Breakdowns tab (drawn, view only): one group per item with Breakdown checked on the Catalog, ordered by
// start time. Each group shows the item's steps (typed on the Steps tab) at clock time, one view per distinct time
// slot the item has on the Week Plan, flagging steps that run past the slot.
//
// Nothing here is typed by hand, so the tab is simply redrawn whole from the schedule.

// ---------------------------------------------------------------- entry point

/** Redraw the Breakdowns tab from `schedule` (see readSchedule_), skipping it if unchanged unless `force`. */
function drawBreakdowns_(schedule, force) {
  const page = layoutBreakdowns_(shownGroups_(schedule));
  const plan = redrawPlan_(TABS.breakdowns, page, { viewCount: page.viewCount, BREAKDOWNS }, force);
  if (!plan.draw) return;
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(TABS.breakdowns) || ss.insertSheet(TABS.breakdowns);
  paintPage_(sh, page, plan.resize);
  plan.remember();
}

// ---------------------------------------------------------------- deciding what to draw

/** Items with Breakdown checked, earliest start first: [{ item, steps, slots }]. */
function shownGroups_(schedule) {
  return [...schedule.catalog.values()]
    .filter(item => item.breakdown)
    .map(item => ({ item, steps: schedule.steps.get(item.name) || [], slots: slotsFor_(item.name, schedule.plan) }))
    .sort((a, b) => firstStart_(a.item.name, schedule.plan) - firstStart_(b.item.name, schedule.plan));
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
 * Turn the groups into a page: a list of rows, each a Map of column → cell (see paintRows_), and how many view
 * columns the widest group needs.
 */
function layoutBreakdowns_(groups) {
  const rows = [];
  const add = () => { const row = new Map(); rows.push(row); return row; };
  const viewCount = Math.max(1, ...groups.map(g => Math.max(1, g.slots.length)));
  const left = viewColumn_(0, 0);

  // Title and hint
  add().set(left, { value: 'Breakdowns', bold: true, size: 14 });
  add().set(left, { value: `Drawn from the ${TABS.steps} tab: type steps there. Check Breakdown on the ${TABS.catalog} to add an item.`, color: COLOR.muted });

  groups.forEach(g => {
    layoutGroup_(g, add);
    add();                                              // blank row between groups
  });
  return { rows, viewCount };
}

/** One group: the item's name, then one view per slot side by side (a slot title row, then the view's rows). */
function layoutGroup_(g, add) {
  const total = g.steps.reduce((sum, s) => sum + s.min, 0);
  const slots = g.slots.length ? g.slots : [{ start: null, end: null, days: [], grids: [] }];
  const views = slots.map(slot => viewCells_(g, slot, total));

  add().set(viewColumn_(0, BREAKDOWNS.view.time), { value: g.item.name, bold: true, size: 12 });
  const title = add();
  views.forEach((v, k) => v.title.forEach((cell, offset) => title.set(viewColumn_(k, offset), cell)));

  const bodyRows = Math.max(...views.map(v => v.body.length));
  for (let i = 0; i < bodyRows; i++) {
    const row = add();
    views.forEach((v, k) => (v.body[i] || new Map()).forEach((cell, offset) => row.set(viewColumn_(k, offset), cell)));
  }
}

/**
 * The cells of one view: a title (slot and how full it is) and one body row per `viewMinutes` (default 5), each a
 * Map of column offset within the view → cell. Rows past the end of the slot are flagged; unused rows in the slot
 * are grey.
 */
function viewCells_(g, slot, total) {
  const v = BREAKDOWNS.view;
  const step = g.item.viewMinutes || BREAKDOWNS.minutesPerRow;
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

  // Where each step starts, in minutes from the start of the slot. A step with no minutes is a marker:
  // it's labeled at the moment it happens but takes up no time.
  const starts = [];
  let at = 0;
  g.steps.forEach(s => { starts.push({ at, step: s }); at += s.min; });

  const lastMarker = Math.max(-1, ...starts.map(s => s.at));
  const rows = Math.max(Math.ceil(Math.max(length, total) / step), Math.floor(lastMarker / step) + 1);
  const body = [];
  for (let i = 0; i < rows; i++) {
    const offset = i * step;
    const over = offset >= length;
    const used = offset < total;
    const bar = over ? COLOR.over : used ? g.item.color : COLOR.free;
    const row = new Map([[v.barLeft, { value: '', bg: bar }], [v.barRight, { value: '', bg: bar }]]);
    const begins = starts.filter(s => s.at >= offset && s.at < offset + step);
    const freeStarts = !used && !over && length > total && offset === Math.ceil(total / step) * step;
    // Steps starting inside the same row share one label ("Meds · Shower"); free time joins it too
    const parts = begins.map(s => (s.step.min > 0 ? s.step.step : BREAKDOWNS.marker + s.step.step));
    if (freeStarts) parts.push(`free · ${length - total} min`);
    if (parts.length) {
      const when = begins.length ? begins[0].at : total;
      const color = over ? COLOR.overText : begins.length ? COLOR.text : COLOR.muted;
      row.set(v.time, { value: clock(when), color: over ? COLOR.overText : COLOR.muted, halign: 'right' });
      row.set(v.label, { value: parts.join(' · '), color, line: true });
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

/**
 * Write a laid-out page onto a blank tab sized exactly to it. Column widths and frozen rows are only set when
 * `resize` (the number of views changed, or the tab is new). Rows are always the default height here.
 */
function paintPage_(sh, page, resize) {
  const c = BREAKDOWNS.col;
  const w = BREAKDOWNS.widths;
  const nCols = c.firstView + page.viewCount * BREAKDOWNS.view.width - 1;

  clearSheet_(sh, true);
  fitSheetSize_(sh, page.rows.length, nCols);
  paintRows_(sh, 1, page.rows, nCols);
  if (!resize) return;

  // Column widths: margin, then each view
  sh.setColumnWidth(c.margin, w.margin);
  for (let k = 0; k < page.viewCount; k++) {
    [w.time, w.bar, w.label, w.bar, w.gap].forEach((px, offset) => sh.setColumnWidth(viewColumn_(k, offset), px));
  }
  sh.setFrozenRows(BREAKDOWNS.firstRow - 1);
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
