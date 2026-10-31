/**
 * 🗓️ Gestalt Math – Schedule: the Apps Script behind the schedule sheet.
 *
 * What it does
 *  - The data tabs are the only place anything is typed: Categories (with optional weekly targets),
 *    Catalog (what each item is), Week Plan (when each item happens) and the steps on Breakdowns.
 *  - Everything else is drawn from them: the weekday and weekend grids, the Breakdowns views and Stats.
 *
 * Conventions (same as house-tasks)
 *  - Every file in src/ shares one global scope (Apps Script has no imports), so names must be unique.
 *  - A trailing underscore (readPlan_) marks an internal helper. Apps Script hides those from the Run menu.
 *  - All settings live in this file, grouped into frozen (read-only) objects. Nothing else runs at load time.
 *  - Sheet rows and columns are 1-based, as in the Sheets API. Times of day are minutes after midnight.
 *
 * Deploying: `clasp push` from the repo (see README), then reload the sheet.
 */

// ---------------------------------------------------------------- sheets

/** Names of the data tabs the script reads. */
const TABS = Object.freeze({
  categories: 'Categories',
  catalog: 'Catalog',
  plan: 'Week Plan',
  breakdowns: 'Breakdowns',
});

// ---------------------------------------------------------------- days and grids

/** Days as written in the Week Plan "Days" column, Monday first. */
const DAYS = Object.freeze(['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']);

/** Week Plan "Grid" values, and the days each one covers. */
const GRIDS = Object.freeze({
  'Weekday': Object.freeze(['Mon', 'Tue', 'Wed', 'Thu', 'Fri']),
  'Weekend - Home': Object.freeze(['Sat', 'Sun']),
  'Weekend - Travel': Object.freeze(['Sat', 'Sun']),
});

/** Catalog "Side" values: which column of a day an item is drawn in. */
const SIDES = Object.freeze(['Schedule', 'Activities']);

/** The day every view covers, in minutes after midnight, and the size of one grid chunk. */
const TIME = Object.freeze({ dayStart: 6 * 60, dayEnd: 24 * 60, chunk: 30 });

// ---------------------------------------------------------------- data tab layouts

/** Categories: one row per category; Min/Max are optional weekly targets in hours. */
const CATEGORIES = Object.freeze({
  firstRow: 2,
  col: Object.freeze({ name: 1, min: 2, max: 3, notes: 4 }),
});

/** Catalog: one row per item. Color is a "#RRGGBB" string; Breakdown is a checkbox. */
const CATALOG = Object.freeze({
  firstRow: 2,
  col: Object.freeze({ item: 1, side: 2, category: 3, color: 4, description: 5, breakdown: 6 }),
});

/** Week Plan: one row per placement. Days like "Mon-Fri" or "Tue, Thu"; Start/End are times of day. */
const PLAN = Object.freeze({
  firstRow: 2,
  col: Object.freeze({ item: 1, label: 2, days: 3, start: 4, end: 5, grid: 6, notes: 7 }),
});

// ---------------------------------------------------------------- Breakdowns layout

/**
 * Breakdowns tab. Rows 1–3 are the title, a hint and column headers; groups start below.
 * Left: a hidden key column (which item each row belongs to), then Step and Min, which the user types.
 * Right: one view per distinct time slot of the item, each `view.width` columns wide, drawn by the script.
 */
const BREAKDOWNS = Object.freeze({
  firstRow: 4,                            // first group row of a drawn tab
  readFrom: 2,                            // where reading steps starts: also picks up older hand-built layouts
  minutesPerRow: 5,                       // each view row is 5 minutes
  titleMark: '§',                         // key-column prefix on a group's title row
  hiddenMark: '§§',                       // key-column value on the row that starts the hidden holding area
  col: Object.freeze({ key: 1, step: 2, min: 3, firstView: 5 }),
  // Inside one view, columns counted from its first: time · bar · label · bar · gap
  view: Object.freeze({ time: 0, barLeft: 1, label: 2, barRight: 3, width: 5 }),
  widths: Object.freeze({ key: 20, step: 220, min: 50, gap: 24, time: 70, bar: 14, label: 190 }),
});

// ---------------------------------------------------------------- look (plain for now; styling comes later)

const COLOR = Object.freeze({
  text: '#000000',
  muted: '#666666',
  titleBand: '#F3F3F3',      // group title rows
  free: '#EFEFEF',           // view rows inside the slot with no step
  over: '#E06666',           // view rows past the end of the slot
  overText: '#CC0000',
  line: '#999999',           // the comment line under each step label
  white: '#FFFFFF',
});
