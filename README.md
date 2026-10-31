# Gestalt Math – Schedule

Apps Script behind the **Gestalt Math - Schedule v2** Google Sheet: a big-picture view of the ideal week,
drawn from a few data tabs, with routine breakdowns and time stats.

The script lives in `src/` (that's clasp's rootDir, so only `src/` is ever pushed to Apps Script).

## How the sheet works

You only type into the data tabs. Everything else is drawn from them.

| Tab | You edit it? | What it holds |
|---|---|---|
| Categories | yes | Category names, with optional Min/Max hours per week |
| Catalog | yes | Every item: which side of the grid (Schedule / Activities), category, color, description, Breakdown checkbox |
| Week Plan | yes | When each item happens: item, optional label, days (`Mon-Fri`, `Tue, Thu`), start, end, which grid |
| Breakdowns | steps only | One group per item with Breakdown checked; you fill in steps and minutes, the view beside it is drawn |
| Weekdays / Weekends | no | The grids, drawn in 30-minute chunks *(coming)* |
| Stats | no | Hours per category against targets *(coming)* |

## Where things are

Apps Script has no imports: every file in `src/` shares one global scope, so a function in one file can call
any other directly, and names must be unique across all files.

| File | What's in it |
|---|---|
| `Config.js` | Every setting: tab names, column layouts, days, grids, the time range |
| `Data.js` | Reading the data tabs into plain objects, and reporting rows that can't be used |

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
computer. `test/fixtures/schedule.json` is a copy of the real data tabs; `test/data.js` checks that it reads
cleanly and that broken rows are reported instead of crashing.

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
