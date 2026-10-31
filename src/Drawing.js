// Small drawing helpers shared by every drawn tab.

/** Grow a sheet to at least `rows` × `cols` (never shrinks it). */
function ensureSheetSize_(sh, rows, cols) {
  if (sh.getMaxColumns() < cols) sh.insertColumnsAfter(sh.getMaxColumns(), cols - sh.getMaxColumns());
  if (sh.getMaxRows() < rows) sh.insertRowsAfter(sh.getMaxRows(), rows - sh.getMaxRows());
}

/** Make a sheet exactly `rows` × `cols`, adding or deleting rows and columns at the end. */
function fitSheetSize_(sh, rows, cols) {
  ensureSheetSize_(sh, rows, cols);
  if (sh.getMaxRows() > rows) sh.deleteRows(rows + 1, sh.getMaxRows() - rows);
  if (sh.getMaxColumns() > cols) sh.deleteColumns(cols + 1, sh.getMaxColumns() - cols);
}

/**
 * A darker shade of a "#RRGGBB" color, for outlines: same hue, lightness brought down to about the middle,
 * saturation kept in a range that reads as color rather than grey or neon. Close to Google's own "dark" palette
 * (e.g. #9FC5E8 → about #3D85C6).
 */
function outlineColor_(hex) {
  const [r, g, b] = [1, 3, 5].map(i => parseInt(hex.slice(i, i + 2), 16) / 255);
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const light = (max + min) / 2;
  const delta = max - min;
  let hue = 0;
  if (delta) {
    if (max === r) hue = ((g - b) / delta) % 6;
    else if (max === g) hue = (b - r) / delta + 2;
    else hue = (r - g) / delta + 4;
  }
  hue = (hue * 60 + 360) % 360;
  const sat = delta ? delta / (1 - Math.abs(2 * light - 1)) : 0;
  return hslToHex_(hue, Math.min(0.85, Math.max(0.35, sat)), Math.min(light, 0.5));
}

/** Hue (degrees), saturation and lightness (0–1) → "#RRGGBB". */
function hslToHex_(h, s, l) {
  const c = (1 - Math.abs(2 * l - 1)) * s;
  const x = c * (1 - Math.abs(((h / 60) % 2) - 1));
  const m = l - c / 2;
  const [r, g, b] = h < 60 ? [c, x, 0] : h < 120 ? [x, c, 0] : h < 180 ? [0, c, x] : h < 240 ? [0, x, c] : h < 300 ? [x, 0, c] : [c, 0, x];
  return '#' + [r, g, b].map(v => Math.round((v + m) * 255).toString(16).padStart(2, '0')).join('').toUpperCase();
}

/**
 * Wipe a sheet back to blank: values, formats, notes, merges, dropdowns, highlight rules, hidden rows and columns.
 * Row heights are reset to the default too, unless `keepSizes` (the caller knows they're already right).
 */
function clearSheet_(sh, keepSizes) {
  const all = sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns());
  sh.showRows(1, sh.getMaxRows());
  sh.showColumns(1, sh.getMaxColumns());
  all.breakApart();
  sh.clear();
  all.clearNote();
  all.clearDataValidations();
  sh.clearConditionalFormatRules();
  if (!keepSizes) sh.setRowHeights(1, sh.getMaxRows(), 21);
}

// ---------------------------------------------------------------- skipping work that's already done

/**
 * How much of a drawn tab needs redrawing, judged against what was last drawn there:
 *  - draw:   the laid-out page differs (or `force`, or the tab doesn't exist yet). If false, skip the tab entirely.
 *  - resize: the page's shape (row count, columns, widths) differs, so row heights and column widths need setting.
 * Call remember() once the tab has been drawn. Fingerprints are kept in the document's properties.
 */
function redrawPlan_(tabName, page, shape, force) {
  const props = PropertiesService.getDocumentProperties();
  const isNew = !SpreadsheetApp.getActive().getSheetByName(tabName);
  const pageKey = `drawn:${tabName}:page`;
  const shapeKey = `drawn:${tabName}:shape`;
  const pagePrint = fingerprint_(page);
  const shapePrint = fingerprint_(shape);
  return {
    draw: force || isNew || props.getProperty(pageKey) !== pagePrint,
    resize: force || isNew || props.getProperty(shapeKey) !== shapePrint,
    remember: () => props.setProperties({ [pageKey]: pagePrint, [shapeKey]: shapePrint }),
  };
}

/** A short fingerprint of any value (Maps and Sets included): equal values give equal fingerprints. */
function fingerprint_(value) {
  const json = JSON.stringify(value, (_, v) => (v instanceof Map ? [...v.entries()] : v instanceof Set ? [...v] : v));
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, json));
}

/**
 * Write laid-out rows starting at `firstRow`, nCols wide. Each row is a Map of column → cell, where a cell is
 * { value, bg, color, bold, halign, line, size } (line = a comment line under the cell). Everything goes out in a
 * handful of batched calls: values and formats as whole grids, then all comment lines as one range list.
 */
function paintRows_(sh, firstRow, rows, nCols) {
  if (!rows.length) return;
  const grid = fn => rows.map(row => Array.from({ length: nCols }, (_, j) => fn(row.get(j + 1))));
  const range = sh.getRange(firstRow, 1, rows.length, nCols);
  range.setValues(grid(cell => (cell && cell.value !== undefined ? cell.value : '')));
  range.setBackgrounds(grid(cell => (cell && cell.bg) || COLOR.white));
  range.setFontColors(grid(cell => (cell && cell.color) || COLOR.text));
  range.setFontWeights(grid(cell => (cell && cell.bold ? 'bold' : 'normal')));
  range.setHorizontalAlignments(grid(cell => (cell && cell.halign) || 'left'));
  range.setVerticalAlignment('middle');

  const lines = [];
  const sized = new Map();
  rows.forEach((row, i) => row.forEach((cell, col) => {
    if (cell.line) lines.push(a1_(firstRow + i, col));
    if (cell.size) sized.set(a1_(firstRow + i, col), cell.size);
  }));
  if (lines.length) sh.getRangeList(lines).setBorder(null, null, true, null, null, null, COLOR.line, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sized.forEach((size, a1) => sh.getRange(a1).setFontSize(size));
}

/** Row and column (1-based) → "C12". */
function a1_(row, col) {
  let letters = '';
  for (let n = col; n > 0; n = Math.floor((n - 1) / 26)) letters = String.fromCharCode(65 + ((n - 1) % 26)) + letters;
  return letters + row;
}

// ---------------------------------------------------------------- problems

/**
 * Show problems where they'd be fixed: each one's message in its tab's ⚠ Check column (several on one row are
 * joined with " · "), stale messages cleared, and a short toast pointing there.
 */
function reportProblems_(problems) {
  markProblems_(problems);
  if (!problems.length) return;
  const tabs = [...new Set(problems.map(p => p.tab))].join(', ');
  const count = problems.length === 1 ? '1 row needs' : `${problems.length} rows need`;
  SpreadsheetApp.getActive().toast(`${count} attention. See ${CHECK.header} on ${tabs}.`, '🗓️ Schedule', 8);
}

/** Write every input tab's ⚠ Check column from `problems`, touching only tabs whose messages changed. */
function markProblems_(problems) {
  const ss = SpreadsheetApp.getActive();
  Object.entries(CHECK.col).forEach(([tabName, col]) => {
    const sh = ss.getSheetByName(tabName);
    if (!sh) return;
    if (text_(sh.getRange(1, col).getValue()) !== CHECK.header) setUpCheckColumn_(sh, col);

    const byRow = new Map();
    problems.filter(p => p.tab === tabName).forEach(p => byRow.set(p.row, (byRow.has(p.row) ? byRow.get(p.row) + ' · ' : '') + p.message));
    const last = Math.max(sh.getLastRow(), ...byRow.keys(), 2);
    const range = sh.getRange(2, col, last - 1, 1);
    const wanted = Array.from({ length: last - 1 }, (_, i) => [byRow.get(i + 2) || '']);
    const current = range.getValues().map(r => [text_(r[0])]);
    if (JSON.stringify(current) !== JSON.stringify(wanted)) range.setValues(wanted);
  });
}

/** First time on a tab: the ⚠ Check header, a wide red-text column, and the column-A highlight for problem rows. */
function setUpCheckColumn_(sh, col) {
  ensureSheetSize_(sh, sh.getMaxRows(), col);
  sh.getRange(1, col).setValue(CHECK.header).setFontWeight('bold').setBackground(COLOR.titleBand);
  sh.getRange(2, col, sh.getMaxRows() - 1, 1).setFontColor(COLOR.problemText).setWrap(true);
  sh.setColumnWidth(col, CHECK.width);
  const letter = a1_(1, col).replace(/\d+/, '');
  const rules = sh.getConditionalFormatRules();
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=$${letter}2<>""`)
    .setBackground(COLOR.problem)
    .setRanges([sh.getRange(2, 1, sh.getMaxRows() - 1, 1)])
    .build());
  sh.setConditionalFormatRules(rules);
}
