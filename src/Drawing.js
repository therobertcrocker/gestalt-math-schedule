// Small drawing helpers shared by every drawn tab.

/** Grow a sheet to at least `rows` × `cols` (never shrinks it). */
function ensureSheetSize_(sh, rows, cols) {
  if (sh.getMaxColumns() < cols) sh.insertColumnsAfter(sh.getMaxColumns(), cols - sh.getMaxColumns());
  if (sh.getMaxRows() < rows) sh.insertRowsAfter(sh.getMaxRows(), rows - sh.getMaxRows());
}

/** Wipe a sheet back to blank: values, formats, merges, dropdowns, highlight rules, hidden rows and columns. */
function clearSheet_(sh) {
  const all = sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns());
  sh.showRows(1, sh.getMaxRows());
  sh.showColumns(1, sh.getMaxColumns());
  all.breakApart();
  sh.clear();
  all.clearDataValidations();
  sh.clearConditionalFormatRules();
  sh.setRowHeights(1, sh.getMaxRows(), 21);
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

/** Tell the user about data rows that couldn't be used, without stopping anything. */
function reportProblems_(problems) {
  if (!problems.length) return;
  const first = problems[0];
  const more = problems.length > 1 ? ` (and ${problems.length - 1} more)` : '';
  SpreadsheetApp.getActive().toast(`${first.tab} row ${first.row}: ${first.message}${more}`, '🗓️ Schedule: check your data', 10);
}
