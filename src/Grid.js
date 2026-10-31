// The week grids: Weekdays (Mon–Fri) and Weekends (home Sat–Sun, travel Fri–Sun). Each day is drawn from the Week Plan in the layout of the original
// hand-built Schedule Planning tab: Schedule and Activities side by side, 30-minute chunks separated by thin
// divider rows, colored bars on both sides of each item, and the item's label with a comment line at its start.
//
// The tab is redrawn from scratch every time; nothing on it is typed by hand.

// ---------------------------------------------------------------- entry points

/** Draw the Weekdays tab. Returns any overlap problems found, for the caller to report. */
function drawWeekdays_(schedule) {
  return drawGrid_(TABS.weekdays, WEEKDAY_COLUMNS, schedule);
}

/** Draw the Weekends tab: home Sat–Sun, then travel Fri–Sun. Returns any overlap problems found. */
function drawWeekends_(schedule) {
  return drawGrid_(TABS.weekends, WEEKEND_COLUMNS, schedule);
}

/**
 * Draw one grid tab with the given day columns ([{ day, grid, title }]) from the schedule.
 * Returns problems for items on the same side of the same day that overlap.
 */
function drawGrid_(tabName, columns, schedule) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(tabName) || ss.insertSheet(tabName);
  const placed = placeItems_(columns, schedule);
  const page = layoutGrid_(columns, placed.blocks);

  clearSheet_(sh);
  fitSheetSize_(sh, page.rows.length, page.nCols);
  paintRows_(sh, 1, page.rows, page.nCols);
  page.merges.forEach(([row, col, nRows, nCols]) => sh.getRange(row, col, nRows, nCols).merge());
  drawLabelBoxes_(sh, page.boxes);

  // Sizes: alternating chunk/divider rows, and the fixed column pattern
  page.rows.forEach((_, i) => sh.setRowHeight(i + 1, (i % 2 === 0) ? LAYOUT.heights.chunk : LAYOUT.heights.line));
  sh.setColumnWidth(LAYOUT.hourCol, LAYOUT.hourWidth);
  sh.setColumnWidth(LAYOUT.hourCol + 1, LAYOUT.gutterWidth);
  columns.forEach((_, d) => LAYOUT.widths.forEach((px, offset) => sh.setColumnWidth(dayColumn_(columns, d, offset), px)));
  sectionBreakColumns_(columns).forEach(start => LAYOUT.sectionBreak.forEach((part, offset) => sh.setColumnWidth(start + offset, part.width)));
  sh.setFrozenRows(LAYOUT.subHeaderRow + 1);
  sh.setFrozenColumns(LAYOUT.hourCol + 1);
  return placed.problems;
}

// ---------------------------------------------------------------- placing items

/**
 * Turn Week Plan rows into blocks to draw: [{ col (day index), side, first, last, label, color, row }], where
 * first/last are chunk indexes (0 = 6:00 AM) and last is inclusive. Times are rounded to the nearest chunk.
 * Blocks that overlap another on the same side of the same day are drawn anyway and reported.
 */
function placeItems_(columns, schedule) {
  const chunks = (TIME.dayEnd - TIME.dayStart) / TIME.chunk;
  const blocks = [];
  const problems = [];
  schedule.plan.forEach(p => {
    const item = schedule.catalog.get(p.item);
    const first = Math.max(0, Math.round((p.start - TIME.dayStart) / TIME.chunk));
    const end = Math.min(chunks, Math.round((p.end - TIME.dayStart) / TIME.chunk));
    columns.forEach((c, col) => {
      if (c.grid !== p.grid || !p.days.includes(c.day)) return;
      if (end <= first) {
        problems.push({ tab: TABS.plan, row: p.row, message: `${p.item}: outside the grid's ${formatTime_(TIME.dayStart)} – midnight` });
        return;
      }
      blocks.push({ col, side: item.side, first, last: end - 1, label: p.label, color: item.color, row: p.row, item: p.item, day: c.day });
    });
  });

  // Overlaps: same day, same side, sharing at least one chunk
  blocks.forEach((b, i) => blocks.slice(0, i).forEach(a => {
    if (a.col === b.col && a.side === b.side && a.first <= b.last && b.first <= a.last) {
      problems.push({ tab: TABS.plan, row: b.row, message: `${b.label} overlaps ${a.label} on ${b.day} (${b.side})` });
    }
  }));
  return { blocks, problems };
}

// ---------------------------------------------------------------- layout

/**
 * Build the page: rows (Maps of column → cell), merges ([row, col, nRows, nCols]), label boxes
 * ([{ row, col, nRows, color }]) and the width in columns.
 */
function layoutGrid_(columns, blocks) {
  const chunks = (TIME.dayEnd - TIME.dayStart) / TIME.chunk;
  const nRows = LAYOUT.firstChunkRow - 1 + chunks * 2 + 4;          // headers, chunks + dividers, the midnight label
  const nCols = dayColumn_(columns, columns.length - 1, LAYOUT.dayWidth - 1);
  const breaks = sectionBreakColumns_(columns);
  const rows = Array.from({ length: nRows }, () => new Map());
  const merges = [];
  const boxes = [];
  const set = (row, col, cell) => rows[row - 1].set(col, Object.assign({}, rows[row - 1].get(col), cell));

  // Headers
  columns.forEach((c, d) => {
    const s = LAYOUT.side;
    set(LAYOUT.headerRow, dayColumn_(columns, d, 0), { value: c.title, bold: true, size: 14, halign: 'center' });
    merges.push([LAYOUT.headerRow, dayColumn_(columns, d, 0), 1, s.Activities.barRight + 1]);
    Object.entries(s).forEach(([name, cols]) => {
      set(LAYOUT.subHeaderRow, dayColumn_(columns, d, cols.barLeft), { value: name, bold: true, halign: 'center', line: true });
      merges.push([LAYOUT.subHeaderRow, dayColumn_(columns, d, cols.barLeft), 1, 3]);
    });
    for (let r = 1; r <= nRows; r++) set(r, dayColumn_(columns, d, LAYOUT.separator), { bg: COLOR.separator });
  });

  // Section breaks: a big white gap, then a thin dark line (see LAYOUT.sectionBreak)
  breaks.forEach(start => LAYOUT.sectionBreak.forEach((part, offset) => {
    if (part.dark) for (let r = 1; r <= nRows; r++) set(r, start + offset, { bg: COLOR.separator });
  }));

  // Hour labels down the left, each spanning chunk · divider · chunk; the last one marks midnight
  for (let minutes = TIME.dayStart, row = LAYOUT.firstChunkRow; minutes <= TIME.dayEnd; minutes += 60, row += 4) {
    set(row, LAYOUT.hourCol, { value: formatTime_(minutes), bold: true, halign: 'center' });
    merges.push([row, LAYOUT.hourCol, 3, 1]);
  }

  // Items: bars on both sides from the first chunk through the last (including the dividers between),
  // and the label with its comment line on the first chunk
  blocks.forEach(b => {
    const cols = LAYOUT.side[b.side];
    const top = chunkRow_(b.first);
    const bottom = chunkRow_(b.last);
    for (let r = top; r <= bottom; r++) {
      set(r, dayColumn_(columns, b.col, cols.barLeft), { bg: b.color });
      set(r, dayColumn_(columns, b.col, cols.barRight), { bg: b.color });
    }
    boxes.push({ row: top, col: dayColumn_(columns, b.col, cols.label), nRows: bottom - top + 1, color: outlineColor_(b.color), label: b.label });
  });

  const kept = keepSeparateBoxes_(boxes);
  kept.forEach(box => set(box.row, box.col, { value: box.label, halign: 'center' }));
  return { rows, merges, boxes: kept, nCols };
}

/**
 * Overlapping items (already reported as problems) would need overlapping merged label boxes, which Sheets
 * refuses. A box that starts inside an earlier box in the same column is folded into it: its label joins the
 * earlier one ("Lunch · Late Coffee") and it gets no box of its own.
 */
function keepSeparateBoxes_(boxes) {
  const kept = [];
  [...boxes].sort((a, b) => a.col - b.col || a.row - b.row).forEach(box => {
    const host = kept.find(k => k.col === box.col && box.row >= k.row && box.row < k.row + k.nRows);
    if (host) host.label += ' · ' + box.label;
    else kept.push(Object.assign({}, box));
  });
  return kept;
}

/**
 * Label boxes, as on the hand-built grid: each label cell is merged down the item's full length, centered and
 * wrapped, and outlined on all four sides in a darker shade of the item's color. Outlines are set one call per
 * color, wrapping in one call for all boxes.
 */
function drawLabelBoxes_(sh, boxes) {
  if (!boxes.length) return;
  const byColor = new Map();
  boxes.forEach(b => {
    if (b.nRows > 1) sh.getRange(b.row, b.col, b.nRows, 1).merge();
    const a1 = b.nRows > 1 ? `${a1_(b.row, b.col)}:${a1_(b.row + b.nRows - 1, b.col)}` : a1_(b.row, b.col);
    if (!byColor.has(b.color)) byColor.set(b.color, []);
    byColor.get(b.color).push(a1);
  });
  byColor.forEach((ranges, color) => sh.getRangeList(ranges).setBorder(true, true, true, true, null, null, color, SpreadsheetApp.BorderStyle.SOLID_MEDIUM));
  sh.getRangeList([].concat(...byColor.values())).setWrap(true);
}

/**
 * Sheet column of a cell inside the d-th day (0-based), counted from the day's first column. Where the grid
 * switches sections (Home → Travel on Weekends), a section break sits between the two days.
 */
function dayColumn_(columns, d, offset) {
  return LAYOUT.firstDayCol + d * LAYOUT.dayWidth + sectionBreaksBefore_(columns, d) * LAYOUT.sectionBreak.length + offset;
}

/** How many section breaks come before the d-th day: one each time the Week Plan grid changes. */
function sectionBreaksBefore_(columns, d) {
  let breaks = 0;
  for (let i = 1; i <= d; i++) if (columns[i].grid !== columns[i - 1].grid) breaks++;
  return breaks;
}

/** First sheet column of each section break (the white gap before the dark band). */
function sectionBreakColumns_(columns) {
  const starts = [];
  columns.forEach((c, d) => {
    if (d > 0 && c.grid !== columns[d - 1].grid) starts.push(dayColumn_(columns, d, 0) - LAYOUT.sectionBreak.length);
  });
  return starts;
}

/** Sheet row of a 30-minute chunk (0 = 6:00 AM). */
function chunkRow_(chunk) {
  return LAYOUT.firstChunkRow + chunk * 2;
}
