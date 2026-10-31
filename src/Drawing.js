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

/** Tell the user about data rows that couldn't be used, without stopping anything. */
function reportProblems_(problems) {
  if (!problems.length) return;
  const first = problems[0];
  const more = problems.length > 1 ? ` (and ${problems.length - 1} more)` : '';
  SpreadsheetApp.getActive().toast(`${first.tab} row ${first.row}: ${first.message}${more}`, '🗓️ Schedule: check your data', 10);
}
