# Gestalt Math – Schedule

Apps Script behind the **Gestalt Math - Schedule v2** Google Sheet: a big-picture view of the ideal week,
drawn from a few data tabs, with routine breakdowns and time stats.

The script lives in `src/` (that's clasp's rootDir, so only `src/` is ever pushed to Apps Script).

## How the sheet works

You only type into the data tabs. Everything else is drawn from them.

| Tab | You edit it? | What it holds |
|---|---|---|
| Categories | yes | Category names, with optional Min/Max hours per week |
| Catalog | yes | Every item: which side of the grid (Schedule / Activities), category, color, description, Breakdown checkbox, View min (minutes per Breakdowns view row; blank = 5) |
| Week Plan | yes | When each item happens: item, optional label, days (`Mon-Fri`, `Tue, Thu`), start, end, which grid |
| Breakdowns | steps only | One group per checked item, by start time. You type Step and Min; the clock-time view beside it is drawn. A step with no minutes is a ◆ marker |
| Weekdays | no | Mon–Fri, drawn from the Week Plan in the original grid's layout (30-minute chunks, 10 columns per day) |
| Weekends | no | Home Sat–Sun and Travel Fri–Sun side by side, same layout as Weekdays. Week Plan Grid = `Weekend - Home` / `Weekend - Travel` |
| Stats | no | Hours per category against targets *(coming)* |

## Where things are

Apps Script has no imports: every file in `src/` shares one global scope, so a function in one file can call
any other directly, and names must be unique across all files.

| File | What's in it |
|---|---|
| `Config.js` | Every setting: tab names, column layouts, days, grids, the time range |
| `Data.js` | Reading the data tabs into plain objects, and reporting rows that can't be used |
| `Breakdowns.js` | The Breakdowns tab: reading typed steps, grouping by checked item, the clock-time views, single-group repaints |
| `Grid.js` | The week grids (Weekdays, Weekends): placing Week Plan rows into 30-minute chunks, overlap checks, label boxes |
| `Triggers.js` | Entry points: the 🗓️ Schedule menu, `onEdit` auto-redraw, the document lock |
| `Drawing.js` | Helpers every drawn tab uses: clearing a sheet, the batched row painter, reporting data problems |

### How Breakdowns keeps your steps

A full redraw starts from a blank tab, after first reading the steps already on it. Typing a step only repaints
that group, as long as it stays the same height. A hidden key
column (A) records which item each row belongs to; title rows are marked `§Item`. A row you insert without a key
joins the group above it. Unchecked items move below a hidden `§§` marker row with their steps, and come back
when re-checked.

## Conventions

Same as house-tasks:

- **Settings live in `Config.js`**, in frozen objects. If a number or color describes the layout, it goes there.
- **`const` by default**, `let` only for a value that really changes.
- **Trailing underscore = internal helper** (`readPlan_`). Functions without one are called by name from the
  menu or triggers, so don't rename them.
- **One job per function**, named for what it does, with a `/** … */` comment on top.
- **Data is read as named objects** (`readSchedule_()`), so code says `placement.start`, not `row[3]`.
- **Times of day are minutes after midnight** (7:30 AM = 450). Times are read from what the sheet displays,
  not as Dates, because Sheets stores a bare time on 30 Dec 1899 and old timezone rules can shift it.

## Checking a change: `npm test`

`test/` holds a small stand-in for Google Sheets (copied from house-tasks) that runs the script on your
computer. `test/fixtures/schedule.json` is a copy of the real data tabs.

- `test/data.js`: the data reads cleanly, and broken rows are reported instead of crashing.
- `test/breakdowns.js`: typed steps survive every redraw, views follow the Week Plan, overflow is flagged,
  markers and per-item row sizes, unchecking/rechecking keeps steps, a single-group repaint matches a full redraw.
- `test/grid.js`: every label lands in the same cell as on the hand-built Schedule Planning tab, with the same
  bars, row heights and column widths; label boxes; overlaps are reported; Weekends puts home and travel days
  side by side.

## Workflow

```sh
clasp pull          # only if you edited in the online Apps Script editor
# edit files in src/
npm test
clasp push          # then reload the sheet
git commit -am "…"  # one commit per working change
```

## After a container rebuild

Run `clasp login` again (the token in `~/.clasprc.json` isn't kept, and is never committed).
