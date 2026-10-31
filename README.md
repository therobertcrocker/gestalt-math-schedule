# Gestalt Math – Schedule

Apps Script behind the **Gestalt Math - Schedule v2** Google Sheet: a big-picture view of the ideal week,
drawn from a few data tabs, with routine breakdowns and time stats.

The script lives in `src/` (that's clasp's rootDir, so only `src/` is ever pushed to Apps Script).

## How the sheet works

You only type into the input tabs. Everything else is drawn from them, and the script never redraws a cell you
type in (its only write to an input tab is adding a heading on Steps when you check an item's Breakdown box).

| Tab | You edit it? | What it holds |
|---|---|---|
| Categories | yes | Category names, with optional Min/Max hours per week |
| Catalog | yes | Every item: which side of the grid (Schedule / Activities), category, color, description, Breakdown checkbox, View min (minutes per Breakdowns view row; blank = 5) |
| Week Plan | yes | When each item happens: item, optional label, days (`Mon-Fri`, `Tue, Thu`), start, end, which grid |
| Steps | yes | Routine steps, grouped: a heading row with the item's name, then its steps (Step, Min). A step with no minutes is a ◆ marker in the view |
| Breakdowns | no | One group per checked item, by start time: its steps at clock time in its Week Plan slot, flagging steps that run over |
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
| `Steps.js` | The Steps tab: reading grouped steps, adding headings for newly checked items, the one-time move from the first Breakdowns layout |
| `Breakdowns.js` | The Breakdowns tab (view only): one group per checked item, the clock-time views |
| `Grid.js` | The week grids (Weekdays, Weekends): placing Week Plan rows into 30-minute chunks, overlap checks, label boxes |
| `Triggers.js` | Entry points: the 🗓️ Schedule menu, `onEdit` auto-redraw, the document lock |
| `Drawing.js` | Helpers every drawn tab uses: clearing a sheet, the batched row painter, fingerprints for skipping unchanged tabs, reporting data problems |

### Input tabs vs drawn tabs

Inputs: Categories, Catalog, Week Plan, Steps. Drawn: Breakdowns, Weekdays, Weekends (and Stats, coming). Editing an
input tab redraws only the drawn tabs whose content would change: each drawn tab keeps a fingerprint of the page it
last drew (in the document's properties) and is skipped when the new page matches, and row heights and column widths
are only set when a tab's shape changes. **🗓️ Schedule → 🔄 Refresh everything** redraws every tab from scratch.
The first refresh after this version creates the Steps tab and moves over any steps typed on the old Breakdowns
layout, including ones parked in its hidden holding area.

### Problems

A row that can't be used (an unknown item, End before Start, a day its grid doesn't have, overlapping items…) gets
its reason in the **⚠ Check** column at the right edge of its tab, and its column-A cell is highlighted. The message
clears itself once the row is fixed. **🗓️ Schedule → ⚠ Check data** lists every problem at once.

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
- `test/breakdowns.js`: steps move over from the old layout; headings are added without touching typed rows;
  the view's clock times, free time, overflow, markers, per-item row sizes and multiple slots.
- `test/grid.js`: every label lands in the same cell as on the hand-built Schedule Planning tab, with the same
  bars, row heights and column widths; label boxes; overlaps are reported; Weekends puts home and travel days
  side by side.
- `test/redraw.js`: edits only redraw what they change, sizes are only set when needed, a repaint matches drawing
  from scratch, and Refresh everything always redraws.
- `test/problems.js`: problems land in each tab's ⚠ Check column and clear when fixed; Check data lists them all.

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
