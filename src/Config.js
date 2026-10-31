/**
 * 🗓️ Gestalt Math – Schedule: the Apps Script behind the schedule sheet.
 *
 * What it does
 *  - The input tabs are the only place anything is typed: Categories (with optional weekly targets),
 *    Catalog (what each item is), Week Plan (when each item happens) and Steps (routine breakdowns).
 *  - Everything else is drawn from them: the weekday and weekend grids, Breakdowns and Stats.
 *  - The script never redraws a cell you type in. The only thing it ever writes on an input tab is a new heading
 *    on Steps when an item gets its Breakdown box checked.
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
  steps: 'Steps',
  breakdowns: 'Breakdowns',
  weekdays: 'Weekdays',
  weekends: 'Weekends',
});

/** Full day names for grid headers. */
const DAY_NAMES = Object.freeze({ Mon: 'Monday', Tue: 'Tuesday', Wed: 'Wednesday', Thu: 'Thursday', Fri: 'Friday', Sat: 'Saturday', Sun: 'Sunday' });

// ---------------------------------------------------------------- days and grids

/** Days as written in the Week Plan "Days" column, Monday first. */
const DAYS = Object.freeze(['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']);

/** Week Plan "Grid" values, and the days each one covers. */
const GRIDS = Object.freeze({
  'Weekday': Object.freeze(['Mon', 'Tue', 'Wed', 'Thu', 'Fri']),
  'Weekend - Home': Object.freeze(['Sat', 'Sun']),
  'Weekend - Travel': Object.freeze(['Fri', 'Sat', 'Sun']),   // travel weekends start Friday
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

/**
 * Catalog: one row per item. Color is a "#RRGGBB" string; Breakdown is a checkbox; View min is how many minutes
 * each row of the item's Breakdowns view covers (blank = BREAKDOWNS.minutesPerRow).
 */
const CATALOG = Object.freeze({
  firstRow: 2,
  col: Object.freeze({ item: 1, side: 2, category: 3, color: 4, description: 5, breakdown: 6, viewMinutes: 7 }),
});

/** Week Plan: one row per placement. Days like "Mon-Fri" or "Tue, Thu"; Start/End are times of day. */
const PLAN = Object.freeze({
  firstRow: 2,
  col: Object.freeze({ item: 1, label: 2, days: 3, start: 4, end: 5, grid: 6, notes: 7 }),
});

/**
 * Steps: grouped breakdown steps. A heading row has the item in column A and nothing else; the rows under it are
 * its steps (Step and Min). The script only ever appends headings for newly checked items; it never moves or
 * clears a row here.
 */
const STEPS = Object.freeze({
  firstRow: 2,
  col: Object.freeze({ item: 1, step: 2, min: 3 }),
  widths: Object.freeze({ item: 160, step: 260, min: 60 }),
});

/** The first Breakdowns layout (steps typed on Breakdowns itself), read once to move its steps onto Steps. */
const LEGACY_BREAKDOWNS = Object.freeze({ readFrom: 2, titleMark: '§', hiddenMark: '§§' });

// ---------------------------------------------------------------- Breakdowns layout (drawn, view only)

/**
 * Breakdowns tab. Rows 1–3 are the title, a hint and column headers; groups start below, one per item with
 * Breakdown checked, ordered by start time. Left: the item's steps (read-only copy of Steps). Right: one view per
 * distinct time slot of the item, each `view.width` columns wide.
 */
const BREAKDOWNS = Object.freeze({
  firstRow: 4,
  minutesPerRow: 5,                       // default minutes per view row (Catalog "View min" overrides per item)
  marker: '◆ ',                           // prefix for a step with no minutes (a moment, not a span)
  col: Object.freeze({ margin: 1, step: 2, min: 3, firstView: 5 }),
  // Inside one view, columns counted from its first: time · bar · label · bar · gap
  view: Object.freeze({ time: 0, barLeft: 1, label: 2, barRight: 3, width: 5 }),
  widths: Object.freeze({ margin: 20, step: 220, min: 50, gap: 24, time: 70, bar: 14, label: 190 }),
});

// ---------------------------------------------------------------- grid layout (Weekdays / Weekends)

/**
 * The week grids, matching the original hand-built Schedule Planning tab:
 *  - Rows: header (21px), divider (5px), Schedule/Activities sub-headers (21px), divider, then one 21px row per
 *    30-minute chunk, each followed by a 5px divider row, from 6:00 AM to midnight. Hour labels in column A span
 *    an hour's chunk, divider and chunk.
 *  - Columns: hour labels (A), a 5px gutter (B), then 10 columns per day:
 *    bar · label · bar · gap  (Schedule)   bar · label · bar · gap  (Activities)   dark separator · gap
 */
const LAYOUT = Object.freeze({
  headerRow: 1,
  subHeaderRow: 3,
  firstChunkRow: 5,
  hourCol: 1,
  firstDayCol: 3,
  dayWidth: 10,
  side: Object.freeze({
    Schedule: Object.freeze({ barLeft: 0, label: 1, barRight: 2 }),
    Activities: Object.freeze({ barLeft: 4, label: 5, barRight: 6 }),
  }),
  separator: 8,                                 // offset of the dark column inside a day
  // Between sections (Home → Travel on Weekends), after the last day's own thin separator: a big white gap,
  // then a second thin dark line and its gap, so the sections read as separate panels
  sectionBreak: Object.freeze([Object.freeze({ width: 40 }), Object.freeze({ width: 5, dark: true }), Object.freeze({ width: 5 })]),
  widths: Object.freeze([25, 100, 25, 5, 25, 100, 25, 5, 5, 5]),   // one day's 10 columns
  hourWidth: 100,
  gutterWidth: 5,
  heights: Object.freeze({ chunk: 21, line: 5 }),
});

/** Days drawn on the Weekdays tab, left to right. */
const WEEKDAY_COLUMNS = Object.freeze(['Mon', 'Tue', 'Wed', 'Thu', 'Fri'].map(day => Object.freeze({ day, grid: 'Weekday', title: DAY_NAMES[day] })));

/** Days drawn on the Weekends tab, left to right: the home weekend, then the travel weekend (which starts Friday). */
const WEEKEND_COLUMNS = Object.freeze([
  ...GRIDS['Weekend - Home'].map(day => Object.freeze({ day, grid: 'Weekend - Home', title: `${DAY_NAMES[day]} · Home` })),
  ...GRIDS['Weekend - Travel'].map(day => Object.freeze({ day, grid: 'Weekend - Travel', title: `${DAY_NAMES[day]} · Travel` })),
]);

// ---------------------------------------------------------------- look (plain for now; styling comes later)

const COLOR = Object.freeze({
  text: '#000000',
  muted: '#666666',
  titleBand: '#F3F3F3',      // group title rows; Steps heading rows
  free: '#EFEFEF',           // view rows inside the slot with no step
  over: '#E06666',           // view rows past the end of the slot
  overText: '#CC0000',
  line: '#999999',           // the comment line under each step label
  white: '#FFFFFF',
  separator: '#434343',      // dark column between days on the grids
});
