// The Stats tab (drawn): where a Weekday week's hours go, per category, against the targets on Categories. Three
// tables: week totals with a target status, hours per day, and the open time left each day.
//
// Counting: each weekday is laid out minute by minute from 6 AM to midnight. Schedule blocks go down first, then
// Activities on top, so an Activity takes the time of any block it overlaps and every minute counts once.

// ---------------------------------------------------------------- entry point

/** Redraw the Stats tab from `schedule`, skipping it if unchanged unless `force`. */
function drawStats_(schedule, force) {
  const page = layoutStats_(countWeek_(schedule), schedule.categories);
  const plan = redrawPlan_(TABS.stats, page, { nCols: page.nCols, STATS }, force);
  if (!plan.draw) return;
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(TABS.stats) || ss.insertSheet(TABS.stats);
  clearSheet_(sh, !plan.resize);
  fitSheetSize_(sh, page.rows.length, page.nCols);
  paintRows_(sh, 1, page.rows, page.nCols);
  if (plan.resize) {
    sh.setColumnWidth(STATS.col.label, STATS.widths.label);
    for (let col = STATS.col.first; col <= page.nCols; col++) sh.setColumnWidth(col, STATS.widths.number);
    sh.setFrozenColumns(1);
  }
  plan.remember();
}

// ---------------------------------------------------------------- counting

/**
 * Minutes per day per bucket for the Weekday grid: { days, byDay: Map day → Map bucket → minutes, totals: Map
 * bucket → minutes }. Buckets are category names, STATS.uncounted (items with no category) and STATS.open.
 */
function countWeek_(schedule) {
  const days = GRIDS[STATS.grid];
  const length = TIME.dayEnd - TIME.dayStart;
  const byDay = new Map();
  const totals = new Map();
  days.forEach(day => {
    const minutes = new Array(length).fill(STATS.open);
    ['Schedule', 'Activities'].forEach(side => {                // Activities last, so they win overlaps
      schedule.plan
        .filter(p => !p.touchstone && p.grid === STATS.grid && p.days.includes(day))
        .filter(p => schedule.catalog.get(p.item).side === side)
        .forEach(p => {
          const bucket = schedule.catalog.get(p.item).category || STATS.uncounted;
          for (let m = Math.max(p.start, TIME.dayStart); m < Math.min(p.end, TIME.dayEnd); m++) minutes[m - TIME.dayStart] = bucket;
        });
    });
    const counts = new Map();
    minutes.forEach(bucket => counts.set(bucket, (counts.get(bucket) || 0) + 1));
    byDay.set(day, counts);
    counts.forEach((n, bucket) => totals.set(bucket, (totals.get(bucket) || 0) + n));
  });
  return { days, byDay, totals };
}

/** Where a total sits against a category's target: { text, color }. */
function targetStatus_(hours, category) {
  const { min, max } = category;
  if (min === null && max === null) return { text: '', color: COLOR.muted };
  if (min !== null && hours < min) return { text: `▼ ${formatHours_(min - hours)} under`, color: COLOR.statusUnder };
  if (max !== null && hours > max) return { text: `▲ ${formatHours_(hours - max)} over`, color: COLOR.statusOver };
  return { text: '✓ on target', color: COLOR.statusOn };
}

// ---------------------------------------------------------------- layout

/** The page: week totals with targets, hours per day, and open time per day. Rows are Maps of column → cell. */
function layoutStats_(week, categories) {
  const c = STATS.col;
  const rows = [];
  const add = () => { const row = new Map(); rows.push(row); return row; };
  const hours = (bucket, day) => (day ? week.byDay.get(day).get(bucket) || 0 : week.totals.get(bucket) || 0) / 60;
  const buckets = [...categories.map(cat => cat.name), STATS.uncounted];
  const nCols = c.first + week.days.length;                    // label, one column per day, total

  add().set(c.label, { value: 'Stats', bold: true, size: 14 });
  add().set(c.label, { value: `Weekdays (${compressDays_(week.days)}). Activities take the time of any block they overlap. Touchstones aren't counted.`, color: COLOR.muted });
  add();

  // Week totals vs targets
  const head = add();
  [['Week totals', c.label], ['Hours', c.first], ['Target', c.first + 1], ['Status', c.first + 2]]
    .forEach(([value, col]) => head.set(col, { value, bold: true, bg: COLOR.titleBand, halign: col === c.label ? 'left' : 'center' }));
  categories.forEach(cat => {
    const total = hours(cat.name);
    const status = targetStatus_(total, cat);
    const row = add();
    row.set(c.label, { value: cat.name });
    row.set(c.first, { value: formatHours_(total), halign: 'center' });
    row.set(c.first + 1, { value: formatTarget_(cat), halign: 'center', color: COLOR.muted });
    row.set(c.first + 2, { value: status.text, color: status.color });
  });
  [STATS.uncounted, STATS.open].forEach(bucket => {
    const row = add();
    row.set(c.label, { value: bucket, color: COLOR.muted });
    row.set(c.first, { value: formatHours_(hours(bucket)), halign: 'center', color: COLOR.muted });
  });
  add();

  // Hours per day
  const dayHead = add();
  dayHead.set(c.label, { value: 'Per day', bold: true, bg: COLOR.titleBand });
  week.days.forEach((day, i) => dayHead.set(c.first + i, { value: day, bold: true, bg: COLOR.titleBand, halign: 'center' }));
  dayHead.set(nCols, { value: 'Week', bold: true, bg: COLOR.titleBand, halign: 'center' });
  [...buckets, STATS.open].forEach(bucket => {
    const row = add();
    const muted = bucket === STATS.uncounted || bucket === STATS.open;
    row.set(c.label, { value: bucket, color: muted ? COLOR.muted : COLOR.text, bold: bucket === STATS.open });
    week.days.forEach((day, i) => row.set(c.first + i, { value: formatHours_(hours(bucket, day)), halign: 'center', color: muted ? COLOR.muted : COLOR.text, bold: bucket === STATS.open }));
    row.set(nCols, { value: formatHours_(hours(bucket)), halign: 'center', bold: true });
  });

  return { rows, nCols };
}

// ---------------------------------------------------------------- formatting

/** 29 → "29h", 2.5 → "2.5h", 0.25 → "0.25h", 0 → "–". */
function formatHours_(hours) {
  if (!hours) return '–';
  return `${Math.round(hours * 100) / 100}h`;
}

/** A category's target as text: "25–30h", "≥ 8h", "≤ 2h", or "–". */
function formatTarget_(cat) {
  if (cat.min !== null && cat.max !== null) return `${cat.min}–${cat.max}h`;
  if (cat.min !== null) return `≥ ${cat.min}h`;
  if (cat.max !== null) return `≤ ${cat.max}h`;
  return '–';
}
