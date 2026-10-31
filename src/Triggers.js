// Entry points: everything Google calls directly. The 🗓️ Schedule menu and the edit handler. These keep their
// names (no trailing underscore) because triggers and menu items find them by name.

/** Simple trigger: adds the 🗓️ Schedule menu. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🗓️ Schedule')
    .addItem('🔄 Refresh everything', 'refreshAll')
    .addToUi();
}

/** Menu: redraw every drawn tab. */
function refreshAll() {
  withLock_(drawAll_);
}

/**
 * Simple trigger: an edit on any input tab (Catalog, Week Plan, Steps) redraws the drawn tabs. Edits on the drawn
 * tabs themselves, or on Categories (nothing drawn uses it yet), do nothing.
 */
function onEdit(e) {
  if (!e || !e.range) return;
  const tab = e.range.getSheet().getName();
  if ([TABS.catalog, TABS.plan, TABS.steps].includes(tab)) withLock_(drawAll_);
}

/**
 * Read the input tabs once, add Steps headings for newly checked items, draw every drawn tab from that one read,
 * then report all problems together. Creates the Steps tab the first time (moving over steps from the first
 * Breakdowns layout).
 */
function drawAll_() {
  ensureStepsTab_();
  const schedule = readSchedule_();
  addMissingStepHeadings_(schedule);
  drawBreakdowns_(schedule);
  const gridProblems = [...drawWeekdays_(schedule), ...drawWeekends_(schedule)];
  reportProblems_([...schedule.problems, ...gridProblems]);
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
