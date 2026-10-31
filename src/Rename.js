// Renaming an item on the Catalog carries the new name everywhere the old one is used: the Item column of the
// Week Plan and of Steps (headings and one-name-per-row steps). Labels on the Week Plan are left alone, since
// they're names of their own ("Early Morning Block").
//
// It only happens for a plain one-cell edit of a Catalog Item cell (Google reports the old value only then), and
// not when it would merge two items: the new name already belongs to another Catalog row, or another row still
// uses the old name.

/** onEdit, Catalog tab: if this edit renamed an item, rename it on the Week Plan and Steps too. */
function carryRename_(e) {
  const r = e.range;
  if (r.getNumRows() !== 1 || r.getNumColumns() !== 1) return;
  if (r.getColumn() !== CATALOG.col.item || r.getRow() < CATALOG.firstRow) return;
  const from = text_(e.oldValue);
  const to = text_(e.value);
  if (!from || !to || from === to) return;

  const names = readRows_(TABS.catalog, CATALOG.firstRow, 1).map(row => text_(row.values[0]));
  if (names.includes(from) || names.filter(n => n === to).length > 1) return;

  const changed = [TABS.plan, TABS.steps].map(tab => [tab, renameInColumn_(tab, 1, from, to)]).filter(([, n]) => n);
  if (changed.length) {
    const where = changed.map(([tab, n]) => `${tab} (${n} row${n === 1 ? '' : 's'})`).join(' and ');
    SpreadsheetApp.getActive().toast(`Renamed "${from}" to "${to}" on ${where}.`, '🗓️ Schedule', 6);
  }
}

/** Replace `from` with `to` in one column of a tab (from row 2 down). Returns how many cells changed. */
function renameInColumn_(tabName, col, from, to) {
  const sh = SpreadsheetApp.getActive().getSheetByName(tabName);
  if (!sh || sh.getLastRow() < 2) return 0;
  const range = sh.getRange(2, col, sh.getLastRow() - 1, 1);
  const values = range.getValues();
  let count = 0;
  const renamed = values.map(row => (text_(row[0]) === from ? (count++, [to]) : row));
  if (count) range.setValues(renamed);
  return count;
}
