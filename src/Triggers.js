// Entry points: everything Google calls directly. The 🗓️ Schedule menu and the edit handler. These keep their
// names (no trailing underscore) because triggers and menu items find them by name.

/** Simple trigger: adds the 🗓️ Schedule menu. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🗓️ Schedule')
    .addItem('🔄 Refresh everything', 'refreshAll')
    .addItem('🔄 Refresh breakdowns', 'refreshBreakdowns')
    .addToUi();
}

/** Menu: redraw every drawn tab. */
function refreshAll() {
  withLock_(drawAll_);
}

/** Menu: redraw the Breakdowns tab. */
function refreshBreakdowns() {
  withLock_(() => drawBreakdowns_());
}

/** Read the data once, draw every drawn tab from it, then report all problems together. */
function drawAll_() {
  const schedule = readSchedule_();
  drawBreakdowns_(schedule);
  const gridProblems = [...drawWeekdays_(schedule), ...drawWeekends_(schedule)];
  reportProblems_([...schedule.problems, ...gridProblems]);
}

/**
 * Simple trigger: redraws what an edit affects. A step or minute typed on Breakdowns repaints just that group;
 * any Catalog or Week Plan edit redraws every drawn tab.
 */
function onEdit(e) {
  if (!e || !e.range) return;
  const r = e.range;
  const tab = r.getSheet().getName();
  const onSteps = tab === TABS.breakdowns && r.getColumn() <= BREAKDOWNS.col.min && r.getLastRow() >= BREAKDOWNS.firstRow;
  if (onSteps) withLock_(() => redrawStepEdit_(r.getRow(), r.getLastRow()));
  else if (tab === TABS.catalog || tab === TABS.plan) withLock_(drawAll_);
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
