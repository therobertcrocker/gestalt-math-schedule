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
