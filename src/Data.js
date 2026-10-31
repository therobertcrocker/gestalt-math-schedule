// Reading the input tabs into plain objects (Steps is read in Steps.js). Every view (grids, Breakdowns, Stats) works from these, never
// from cell positions, so a layout change on a data tab only touches this file and Config.js.
//
// Anything wrong in the data (an item missing from the Catalog, End before Start, a day the grid doesn't have)
// is collected into a `problems` list with its tab and row, instead of stopping the script.

// ---------------------------------------------------------------- whole schedule

/**
 * Read every input tab. Returns { categories, catalog, plan, steps, stepHeadings, problems }:
 *  - categories:   [{ name, min, max, notes }]           (min/max are hours per week, or null)
 *  - catalog:      Map item name → { name, side, category, color, description, breakdown, viewMinutes, split }
 *                  (split: [{ category, share }] with shares 0–1, or null)
 *  - plan:         [{ row, item, label, days, start, end, grid, notes, touchstone }]   (start/end in minutes;
 *                  touchstones have no end)
 *  - steps:        Map item name → [{ step, min, row }]   (from the Steps tab; see Steps.js)
 *  - stepHeadings: Set of item names that have a heading on Steps
 *  - problems:     [{ tab, row, message }]
 */
function readSchedule_() {
  const problems = [];
  const categories = readCategories_(problems);
  const catalog = readCatalog_(categories, problems);
  const plan = readPlan_(catalog, problems);
  const { steps, headings } = readSteps_(catalog, problems);
  return { categories, catalog, plan, steps, stepHeadings: headings, problems };
}

// ---------------------------------------------------------------- tabs

/** Categories tab → [{ name, min, max, notes }], skipping blank rows. */
function readCategories_(problems) {
  const c = CATEGORIES.col;
  return readRows_(TABS.categories, CATEGORIES.firstRow, 4)
    .filter(r => text_(r.values[c.name - 1]))
    .map(r => {
      const v = r.values;
      const category = {
        name: text_(v[c.name - 1]),
        min: hours_(v[c.min - 1]),
        max: hours_(v[c.max - 1]),
        notes: text_(v[c.notes - 1]),
      };
      if (category.min !== null && category.max !== null && category.min > category.max) {
        problems.push({ tab: TABS.categories, row: r.row, message: `${category.name}: Min is more than Max` });
      }
      return category;
    });
}

/** Catalog tab → Map of item name → item. Later duplicates of a name are reported and ignored. */
function readCatalog_(categories, problems) {
  const c = CATALOG.col;
  const known = new Set(categories.map(cat => cat.name));
  const catalog = new Map();
  readRows_(TABS.catalog, CATALOG.firstRow, CATALOG.col.split).forEach(r => {
    const v = r.values;
    const name = text_(v[c.item - 1]);
    if (!name) return;
    const report = message => problems.push({ tab: TABS.catalog, row: r.row, message: `${name}: ${message}` });
    if (catalog.has(name)) return report('listed twice; using the first row');

    const item = {
      name,
      side: text_(v[c.side - 1]),
      category: text_(v[c.category - 1]) || null,     // blank = shown on grids, not counted in Stats
      color: text_(v[c.color - 1]).toUpperCase(),
      description: text_(v[c.description - 1]),
      breakdown: v[c.breakdown - 1] === true,
      viewMinutes: BREAKDOWNS.minutesPerRow,
    };
    const viewMinutes = v[c.viewMinutes - 1];
    if (viewMinutes !== '' && viewMinutes !== null && viewMinutes !== undefined) {
      const n = Number(viewMinutes);
      if (Number.isInteger(n) && n >= 1 && n <= 30) item.viewMinutes = n;
      else report('View min should be a whole number from 1 to 30');
    }
    item.split = parseSplit_(text_(v[c.split - 1]), known, report);
    if (!SIDES.includes(item.side)) report(`Side should be ${SIDES.slice(0, -1).join(', ')} or ${SIDES[SIDES.length - 1]}`);
    if (item.category && !known.has(item.category)) report(`category "${item.category}" isn't on Categories`);
    if (!/^#[0-9A-F]{6}$/.test(item.color)) report('Color should look like #9FC5E8');
    catalog.set(name, item);
  });
  return catalog;
}

/** Week Plan tab → placements, skipping blank rows and reporting rows that can't be drawn. */
function readPlan_(catalog, problems) {
  const c = PLAN.col;
  const plan = [];
  readRows_(TABS.plan, PLAN.firstRow, 7).forEach(r => {
    const v = r.values;
    const shown = r.display;
    const item = text_(v[c.item - 1]);
    if (!item) return;
    const report = message => problems.push({ tab: TABS.plan, row: r.row, message: `${item}: ${message}` });

    const grid = text_(v[c.grid - 1]);
    const days = parseDays_(text_(v[c.days - 1]));
    const start = parseTime_(shown[c.start - 1]);
    const end = parseTime_(shown[c.end - 1]);

    if (!catalog.has(item)) return report("isn't on the Catalog");
    if (!GRIDS[grid]) return report(`Grid should be one of: ${Object.keys(GRIDS).join(', ')}`);
    if (!days) return report('Days should look like "Mon-Fri" or "Tue, Thu"');
    const outside = days.filter(d => !GRIDS[grid].includes(d));
    if (outside.length) return report(`${outside.join(', ')} isn't part of the ${grid} grid`);
    const label = text_(v[c.label - 1]) || item;
    const notes = text_(v[c.notes - 1]);

    if (catalog.get(item).side === TOUCHSTONE) {                // a moment: Start only, End ignored
      if (start === null) return report('Start needs a time, like 7:30 AM');
      if (start < TIME.dayStart || start >= TIME.dayEnd) return report(`is outside the grids' ${formatTime_(TIME.dayStart)} – midnight`);
      if (start % TIME.chunk) report('time not on a half hour; grids draw it at the nearest half hour');
      plan.push({ row: r.row, item, label, days, start, end: null, grid, notes, touchstone: true });
      return;
    }

    if (start === null || end === null) return report('Start and End need a time, like 7:30 AM');
    if (end <= start) return report('End is not after Start');
    if (start < TIME.dayStart || end > TIME.dayEnd) report(`runs outside the grids' ${formatTime_(TIME.dayStart)} – midnight; the part outside isn't shown`);
    if (start % TIME.chunk || end % TIME.chunk) report('times not on a half hour; grids round them to the nearest half hour');

    plan.push({ row: r.row, item, label, days, start, end, grid, notes, touchstone: false });
  });
  return plan;
}

/**
 * Reorder the Week Plan's rows: by grid (in GRIDS order), then first day, then start time, then end time. A row
 * whose days or times can't be read goes at the end of its grid (one with no readable grid, after all grids);
 * blank rows go last. Ties keep their order. Only the typed columns move (⚠ Check is rewritten afterwards).
 */
function sortPlanRows_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(TABS.plan);
  const count = sh.getLastRow() - PLAN.firstRow + 1;
  if (count < 2) return;
  const c = PLAN.col;
  const range = sh.getRange(PLAN.firstRow, 1, count, c.notes);
  const values = range.getValues();
  const shown = range.getDisplayValues();
  const grids = Object.keys(GRIDS);
  const key = i => {
    const v = values[i];
    const days = parseDays_(text_(v[c.days - 1]));
    const grid = grids.indexOf(text_(v[c.grid - 1]));
    return [
      v.every(x => text_(x) === '') ? 1 : 0,
      grid < 0 ? grids.length : grid,
      days ? DAYS.indexOf(days[0]) : DAYS.length,
      parseTime_(shown[i][c.start - 1]) ?? Infinity,
      parseTime_(shown[i][c.end - 1]) ?? Infinity,
      i,
    ];
  };
  const order = values.map((_, i) => key(i)).sort((a, b) => {
    for (let k = 0; k < a.length; k++) if (a[k] !== b[k]) return a[k] - b[k];
    return 0;
  });
  const sorted = order.map(k => values[k[k.length - 1]]);
  if (order.some((k, i) => k[k.length - 1] !== i)) range.setValues(sorted);
}

// ---------------------------------------------------------------- parsing

/**
 * A Split cell → [{ category, share }] (share 0–1), or null when blank. "Work 70%, House Projects 20%" (the % is
 * optional). Shares can't add up to more than 100%; whatever's left over counts as Not counted in Stats.
 * Anything unreadable is reported and the split ignored (the item then counts toward its Category).
 */
function parseSplit_(textValue, knownCategories, report) {
  if (!textValue) return null;
  const parts = textValue.split(',').map(part => part.trim().match(/^(.+?)\s+(\d+(?:\.\d+)?)\s*%?$/));
  if (parts.some(m => !m)) { report('Split should look like "Work 70%, House Projects 20%"'); return null; }
  const split = parts.map(m => ({ category: m[1].trim(), share: Number(m[2]) / 100 }));
  const unknown = split.filter(s => !knownCategories.has(s.category)).map(s => s.category);
  if (unknown.length) { report(`Split: "${unknown.join('", "')}" isn't on Categories`); return null; }
  if (split.reduce((sum, s) => sum + s.share, 0) > 1.0001) { report('Split adds up to more than 100%'); return null; }
  return split;
}

/**
 * "Mon-Fri", "Tue, Thu, Fri", "Sat" or a mix ("Mon, Wed-Fri") → ['Mon', 'Wed', 'Thu', 'Fri'], in week order.
 * Full names and any capitalization work ("monday", "Tuesday"). Returns null if anything can't be read.
 */
function parseDays_(textValue) {
  if (!textValue) return null;
  const dayIndex = word => DAYS.findIndex(d => word.trim().slice(0, 3).toLowerCase() === d.toLowerCase());
  const picked = new Set();
  for (const part of textValue.split(',')) {
    const ends = part.split(/[-–]/).map(dayIndex);
    if (ends.length > 2 || ends.some(i => i < 0)) return null;
    const [from, to] = [ends[0], ends[ends.length - 1]];
    if (to < from) return null;                                   // ranges don't wrap around the week
    for (let i = from; i <= to; i++) picked.add(i);
  }
  return [...picked].sort((a, b) => a - b).map(i => DAYS[i]);
}

/**
 * A time as the sheet shows it ("7:30 AM", "19:00", "12:00 PM") → minutes after midnight, or null.
 * Times are read from what the sheet displays rather than as Dates: Sheets stores a bare time on
 * 30 Dec 1899, and old timezone rules for that date can shift it by several minutes.
 * "12:00 AM" at the end of a day means midnight, so it comes back as 24 * 60.
 */
function parseTime_(shown) {
  const m = String(shown || '').trim().match(/^(\d{1,2}):(\d{2})(?::\d{2})?\s*([AaPp][Mm])?$/);
  if (!m) return null;
  let hour = Number(m[1]);
  const minute = Number(m[2]);
  const half = m[3] ? m[3].toUpperCase() : null;
  if (minute > 59 || hour > 24 || (half && (hour < 1 || hour > 12))) return null;
  if (half === 'AM' && hour === 12) hour = 0;
  if (half === 'PM' && hour !== 12) hour += 12;
  const minutes = hour * 60 + minute;
  return minutes === 0 ? 24 * 60 : minutes;
}

/** Minutes after midnight → "7:30 AM". */
function formatTime_(minutes) {
  const hour = Math.floor(minutes / 60) % 24;
  const minute = String(minutes % 60).padStart(2, '0');
  return `${hour % 12 || 12}:${minute} ${hour < 12 ? 'AM' : 'PM'}`;
}

// ---------------------------------------------------------------- helpers

/** Every row from firstRow down to the last filled one: [{ row, values, display }] for the first nCols columns. */
function readRows_(tabName, firstRow, nCols) {
  const sh = SpreadsheetApp.getActive().getSheetByName(tabName);
  if (!sh) throw new Error(`Can't find the "${tabName}" tab`);
  const count = sh.getLastRow() - firstRow + 1;
  if (count < 1) return [];
  const range = sh.getRange(firstRow, 1, count, nCols);
  const values = range.getValues();
  const display = range.getDisplayValues();
  return values.map((v, i) => ({ row: firstRow + i, values: v, display: display[i] }));
}

/** A cell value as trimmed text ('' for blank). */
function text_(value) {
  return value === null || value === undefined ? '' : String(value).trim();
}

/** A target cell → hours as a number, or null when blank or not a number. */
function hours_(value) {
  if (value === '' || value === null || value === undefined) return null;
  const n = Number(value);
  return Number.isFinite(n) && n >= 0 ? n : null;
}
