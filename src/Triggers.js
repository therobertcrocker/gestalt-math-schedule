// Entry points: everything Google calls directly. The 🗓️ Schedule menu and the edit handler. These keep their
// names (no trailing underscore) because triggers and menu items find them by name.

/** Simple trigger: adds the 🗓️ Schedule menu. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🗓️ Schedule')
    .addItem('🔄 Refresh everything', 'refreshAll')
    .addItem('⚠ Check data', 'checkData')
    .addToUi();
}

/** Menu: list every problem on the input tabs in one dialog (they're also in each tab's ⚠ Check column). */
function checkData() {
  const problems = withLock_(() => drawAll_(false));
  const ui = SpreadsheetApp.getUi();
  if (!problems.length) return ui.alert('🗓️ Schedule', 'No problems found ✓', ui.ButtonSet.OK);
  const lines = problems.map(p => `${p.tab}, row ${p.row}: ${p.message}`);
  ui.alert('🗓️ Schedule: check your data', lines.join('\n'), ui.ButtonSet.OK);
}

/** Menu: redraw every drawn tab from scratch, even ones that look unchanged (a reset button). */
function refreshAll() {
  withLock_(() => drawAll_(true));
}

/**
 * Simple trigger: an edit on any input tab (Catalog, Week Plan, Steps) redraws the drawn tabs. Edits on the drawn
 * tabs themselves, or on Categories (nothing drawn uses it yet), do nothing.
 */
function onEdit(e) {
  if (!e || !e.range) return;
  const tab = e.range.getSheet().getName();
  if ([TABS.catalog, TABS.plan, TABS.steps].includes(tab)) withLock_(() => drawAll_(false));
}

/**
 * Read the input tabs once, add Steps headings for newly checked items, draw every drawn tab from that one read,
 * then mark all problems together (and return them). Tabs whose content wouldn't change are skipped unless `force`. Creates the
 * Steps tab the first time (moving over steps from the first Breakdowns layout).
 */
function drawAll_(force) {
  ensureStepsTab_();
  const schedule = readSchedule_();
  addMissingStepHeadings_(schedule);
  drawBreakdowns_(schedule, force);
  const gridProblems = [...drawWeekdays_(schedule, force), ...drawWeekends_(schedule, force)];
  const problems = [...schedule.problems, ...gridProblems];
  reportProblems_(problems);
  return problems;
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
