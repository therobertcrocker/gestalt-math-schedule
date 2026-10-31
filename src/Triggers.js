// Entry points: everything Google calls directly. The 🗓️ Schedule menu and the edit handler. These keep their
// names (no trailing underscore) because triggers and menu items find them by name.

/** Simple trigger: adds the 🗓️ Schedule menu. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🗓️ Schedule')
    .addItem('🔄 Refresh everything', 'refreshAll')
    .addItem('⚠ Check data', 'checkData')
    .addSeparator()
    .addItem('↕️ Sort Week Plan', 'sortWeekPlan')
    .addToUi();
}

/** Menu: order the Week Plan by grid (Weekday, Home, Travel), then first day, then start time. */
function sortWeekPlan() {
  withLock_(() => {
    sortPlanRows_();
    drawAll_(false);
  });
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
 * Simple trigger: an edit on any input tab (Categories, Catalog, Week Plan, Steps) redraws the drawn tabs whose
 * content changes. Renaming an item on the Catalog first carries the new name to the Week Plan and Steps. Edits on
 * the drawn tabs themselves do nothing.
 */
function onEdit(e) {
  if (!e || !e.range) return;
  const tab = e.range.getSheet().getName();
  if (![TABS.categories, TABS.catalog, TABS.plan, TABS.steps].includes(tab)) return;
  withLock_(() => {
    if (tab === TABS.catalog) carryRename_(e);
    drawAll_(false);
  });
}

/**
 * Read the input tabs once, add Steps headings for newly checked items, draw every drawn tab from that one read,
 * then mark all problems together (and return them). Tabs whose content wouldn't change are skipped unless `force`. Creates the
 * Steps tab the first time (moving over steps from the first Breakdowns layout).
 */
function drawAll_(force) {
  upgradeInputTabs_();
  ensureStepsTab_();
  const schedule = readSchedule_();
  addMissingStepHeadings_(schedule);
  drawBreakdowns_(schedule, force);
  const gridProblems = [...drawWeekdays_(schedule, force), ...drawWeekends_(schedule, force)];
  drawStats_(schedule, force);
  const problems = [...schedule.problems, ...gridProblems];
  reportProblems_(problems);
  return problems;
}

/**
 * Bring input tabs laid out by an earlier version up to date, without touching anything typed:
 *  - Catalog: a Split column before ⚠ Check (which used to sit where Split now goes).
 */
function upgradeInputTabs_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(TABS.catalog);
  if (!sh) return;
  const col = CATALOG.col.split;
  ensureSheetSize_(sh, sh.getMaxRows(), col);
  const header = text_(sh.getRange(1, col).getValue());
  if (header === CATALOG.splitHeader) return;
  if (header === CHECK.header) sh.insertColumnBefore(col);       // the old ⚠ Check column moves one to the right
  else if (header) return;                                       // something typed there: leave it alone
  sh.getRange(1, col).setValue(CATALOG.splitHeader).setFontWeight('bold').setBackground(COLOR.titleBand);
  sh.setColumnWidth(col, 220);
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
