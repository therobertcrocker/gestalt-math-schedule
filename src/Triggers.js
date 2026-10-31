// Entry points: everything Google calls directly. The 🗓️ Schedule menu and the edit handler. These keep their
// names (no trailing underscore) because triggers and menu items find them by name.

/** Simple trigger: adds the 🗓️ Schedule menu. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🗓️ Schedule')
    .addItem('🔄 Refresh breakdowns', 'refreshBreakdowns')
    .addToUi();
}

/** Menu: redraw the Breakdowns tab. */
function refreshBreakdowns() {
  withLock_(drawBreakdowns_);
}

/**
 * Simple trigger: redraws Breakdowns when something it shows changes. A step or minute typed on Breakdowns
 * repaints just that group; any Catalog edit (checkbox, name, color) or Week Plan edit redraws the whole tab.
 */
function onEdit(e) {
  if (!e || !e.range) return;
  const r = e.range;
  const tab = r.getSheet().getName();
  const onSteps = tab === TABS.breakdowns && r.getColumn() <= BREAKDOWNS.col.min && r.getLastRow() >= BREAKDOWNS.firstRow;
  if (onSteps) withLock_(() => redrawStepEdit_(r.getRow(), r.getLastRow()));
  else if (tab === TABS.catalog || tab === TABS.plan) withLock_(drawBreakdowns_);
}

/** Run fn while holding the document lock, so two redraws never overlap. */
function withLock_(fn) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}
