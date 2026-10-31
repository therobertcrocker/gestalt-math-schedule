// A small in-memory stand-in for the parts of SpreadsheetApp (and friends) that the script uses.
// It records the *final state* of every sheet — values, formulas, formatting, merges, sizes, validation and
// conditional-format rules — so two versions of the code can be compared by what they leave behind.
'use strict';

const colLetter = (n) => {
  let s = '';
  for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
};
const colNumber = (s) => [...s].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);

/** Serialize anything a cell or builder might hold into plain, comparable JSON. */
function plain(v) {
  if (v instanceof Date || (v && typeof v.getTime === 'function' && typeof v.getFullYear === 'function')) {
    const p = (n) => String(n).padStart(2, '0');
    return `Date(${v.getFullYear()}-${p(v.getMonth() + 1)}-${p(v.getDate())} ${p(v.getHours())}:${p(v.getMinutes())}:${p(v.getSeconds())})`;
  }
  if (Array.isArray(v)) return v.map(plain);
  if (v && typeof v === 'object') {
    if (v.__range) return v.__range;
    return Object.fromEntries(Object.entries(v).map(([k, x]) => [k, plain(x)]));
  }
  return v;
}

/** Records a fluent builder chain (rule/validation/text-style builders) as a list of calls. */
function recorder(kind) {
  const calls = [];
  const proxy = new Proxy({}, {
    get(_, name) {
      if (name === 'build') return () => ({ kind, calls: calls.slice() });
      if (name === 'copy') return () => proxy;
      return (...args) => { calls.push([name, ...args.map(plain)]); return proxy; };
    },
  });
  return proxy;
}

class FakeSheet {
  constructor(book, name, rows = 1000, cols = 26) {
    this.book = book;
    this.name = name;
    this.maxRows = rows;
    this.maxCols = cols;
    this.cells = new Map();          // "r,c" -> { v, f, fmt: {}, dv, rich }
    this.merges = new Set();         // "r1,c1,r2,c2"
    this.rowHeights = new Map();
    this.colWidths = new Map();
    this.hiddenCols = new Set();
    this.hiddenRows = new Set();
    this.frozenRows = 0;
    this.hidden = false;
    this.gridlinesHidden = false;
    this.tabColor = null;
    this.cfRules = [];
  }
  cell(r, c, create = true) {
    const k = r + ',' + c;
    if (!this.cells.has(k) && create) this.cells.set(k, { v: '', f: '', fmt: {} });
    return this.cells.get(k);
  }
  getName() { return this.name; }
  getMaxRows() { return this.maxRows; }
  getMaxColumns() { return this.maxCols; }
  insertRowsAfter(_, n) { this.maxRows += n; }
  insertColumnsAfter(_, n) { this.maxCols += n; }
  deleteRows(r, n) { this.maxRows = r - 1 + Math.max(0, this.maxRows - (r - 1) - n); this.dropCells((row) => row >= r && row < r + n); }
  deleteColumns(c, n) { this.maxCols = c - 1 + Math.max(0, this.maxCols - (c - 1) - n); this.dropCells((_, col) => col >= c && col < c + n); }
  dropCells(test) { [...this.cells.keys()].forEach((k) => { const [r, c] = k.split(',').map(Number); if (test(r, c)) this.cells.delete(k); }); }
  getLastRow() {
    let last = 0;
    for (const [k, c] of this.cells) if ((c.v !== '' && c.v !== null && c.v !== undefined) || c.f) last = Math.max(last, +k.split(',')[0]);
    return last;
  }
  getRange(a, b, nr, nc) {
    if (typeof a === 'string') {
      const [p, q] = a.split(':');
      const parse = (s) => { const m = s.match(/^([A-Z]+)(\d*)$/); return [m[2] ? +m[2] : null, colNumber(m[1])]; };
      const [r1, c1] = parse(p), [r2, c2] = q ? parse(q) : [r1, c1];
      if (r1 === null) return new FakeRange(this, 1, c1, this.maxRows, c2 - c1 + 1);   // whole columns, e.g. "A:B"
      return new FakeRange(this, r1, c1, r2 - r1 + 1, c2 - c1 + 1);
    }
    return new FakeRange(this, a, b, nr || 1, nc || 1);
  }
  appendRow(vals) { this.getRange(this.getLastRow() + 1, 1, 1, vals.length).setValues([vals]); }
  deleteRow(r) {
    const moved = new Map();
    for (const [k, c] of this.cells) {
      const [rr, cc] = k.split(',').map(Number);
      if (rr === r) continue;
      moved.set((rr > r ? rr - 1 : rr) + ',' + cc, c);
    }
    this.cells = moved;
  }
  clear() { this.cells = new Map(); }
  clearConditionalFormatRules() { this.cfRules = []; }
  setConditionalFormatRules(rules) { this.cfRules = rules.map(plain); }
  getConditionalFormatRules() { return this.cfRules.slice(); }
  setRowHeight(r, h) { this.rowHeights.set(r, h); }
  setRowHeights(r, n, h) { for (let i = 0; i < n; i++) this.rowHeights.set(r + i, h); }
  setColumnWidth(c, w) { this.colWidths.set(c, w); }
  setColumnWidths(c, n, w) { for (let i = 0; i < n; i++) this.colWidths.set(c + i, w); }
  hideColumns(c, n = 1) { for (let i = 0; i < n; i++) this.hiddenCols.add(c + i); }
  showColumns(c, n = 1) { for (let i = 0; i < n; i++) this.hiddenCols.delete(c + i); }
  getRangeList(a1s) {
    const ranges = a1s.map((a) => this.getRange(a));
    const list = {
      setBorder: (...args) => { ranges.forEach((r) => r.setBorder(...args)); return list; },
      setWrap: (v) => { ranges.forEach((r) => r.setWrap(v)); return list; },
    };
    return list;
  }
  hideRows(r, n = 1) { for (let i = 0; i < n; i++) this.hiddenRows.add(r + i); }
  showRows(r, n = 1) { for (let i = 0; i < n; i++) this.hiddenRows.delete(r + i); }
  setFrozenRows(n) { this.frozenRows = n; }
  setFrozenColumns(n) { this.frozenCols = n; }
  setHiddenGridlines(b) { this.gridlinesHidden = b; }
  setTabColor(c) { this.tabColor = c; }
  hideSheet() { this.hidden = true; }

  /** Everything about this sheet, as sorted plain JSON. */
  dump() {
    const cells = {};
    [...this.cells.keys()].sort((x, y) => { const [a, b] = x.split(',').map(Number), [c, d] = y.split(',').map(Number); return a - c || b - d; })
      .forEach((k) => {
        const c = this.cells.get(k);
        const [r, col] = k.split(',').map(Number);
        const out = {};
        if (c.f) out.f = c.f; else if (c.v !== '' && c.v !== undefined) out.v = plain(c.v);
        if (Object.keys(c.fmt).length) out.fmt = Object.fromEntries(Object.entries(c.fmt).sort());
        if (c.dv) out.dv = c.dv;
        if (c.rich) out.rich = c.rich;
        if (Object.keys(out).length) cells[colLetter(col) + r] = out;
      });
    const sortMap = (m) => Object.fromEntries([...m.entries()].sort((a, b) => a[0] - b[0]));
    return {
      size: [this.maxRows, this.maxCols], hidden: this.hidden, tabColor: this.tabColor, frozenRows: this.frozenRows,
      gridlinesHidden: this.gridlinesHidden, rowHeights: sortMap(this.rowHeights), colWidths: sortMap(this.colWidths),
      hiddenCols: [...this.hiddenCols].sort((a, b) => a - b), hiddenRows: [...this.hiddenRows].sort((a, b) => a - b), merges: [...this.merges].sort(), cfRules: this.cfRules, cells,
    };
  }
}

/** Write one typed-in value: "=…" is a formula, and a leading apostrophe means "keep as text" (Sheets drops it). */
function store(c, v) {
  if (typeof v === 'string' && v.startsWith('=')) { c.f = v; c.v = ''; return; }
  c.v = typeof v === 'string' && v.startsWith("'") ? v.slice(1) : v;
  c.f = '';
}

class FakeRange {
  constructor(sheet, r, c, nr, nc) { Object.assign(this, { sheet, r, c, nr, nc }); this.__range = this.getA1Notation(); }
  each(fn) { for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) fn(this.sheet.cell(this.r + i, this.c + j), i, j); return this; }
  fmt(key, v) { return this.each((c) => { c.fmt[key] = plain(v); }); }
  getSheet() { return this.sheet; }
  getRow() { return this.r; }
  getColumn() { return this.c; }
  getNumRows() { return this.nr; }
  getNumColumns() { return this.nc; }
  getLastRow() { return this.r + this.nr - 1; }
  getLastColumn() { return this.c + this.nc - 1; }
  getCell(i, j) { return new FakeRange(this.sheet, this.r + i - 1, this.c + j - 1, 1, 1); }
  getA1Notation() {
    const a = colLetter(this.c) + this.r;
    return this.nr === 1 && this.nc === 1 ? a : a + ':' + colLetter(this.c + this.nc - 1) + (this.r + this.nr - 1);
  }
  getValue() { const c = this.sheet.cell(this.r, this.c, false); return c ? (c.f ? '' : c.v) : ''; }
  getValues() {
    const out = [];
    for (let i = 0; i < this.nr; i++) {
      const row = [];
      for (let j = 0; j < this.nc; j++) { const c = this.sheet.cell(this.r + i, this.c + j, false); row.push(c ? (c.f ? '' : c.v) : ''); }
      out.push(row);
    }
    return out;
  }
  /** What the sheet would show: tests store times as display text ("7:30 AM"), so plain String() is enough. */
  getDisplayValues() { return this.getValues().map((row) => row.map((v) => (typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v)))); }
  setValue(v) { return this.each((c) => store(c, v)); }
  setValues(vals) {
    if (vals.length !== this.nr || vals.some((r) => r.length !== this.nc)) throw new Error(`setValues shape mismatch at ${this.getA1Notation()}`);
    return this.each((c, i, j) => { if (vals[i][j] !== null) store(c, vals[i][j]); });
  }
  setFormula(f) { return this.each((c) => { c.f = f; c.v = ''; }); }
  setFormulas(fs) { return this.each((c, i, j) => { const f = fs[i][j]; if (f) { c.f = f; c.v = ''; } else { c.f = ''; c.v = ''; } }); }
  setRichTextValue(rt) { return this.each((c) => { c.rich = plain(rt); c.v = ''; c.f = ''; }); }
  clearContent() { return this.each((c) => { c.v = ''; c.f = ''; c.rich = undefined; }); }
  clearDataValidations() { return this.each((c) => { delete c.dv; }); }
  setDataValidation(dv) { return this.each((c) => { if (dv) c.dv = plain(dv); else delete c.dv; }); }
  insertCheckboxes() { return this.each((c) => { c.dv = { kind: 'checkbox' }; c.v = false; c.f = ''; }); }
  merge() {
    // Like Sheets: merging across part of an existing merged range is an error
    const [r1, c1, r2, c2] = [this.r, this.c, this.getLastRow(), this.getLastColumn()];
    this.sheet.merges.forEach((m) => {
      const [a1, b1, a2, b2] = m.split(',').map(Number);
      if (m !== [r1, c1, r2, c2].join(',') && a1 <= r2 && r1 <= a2 && b1 <= c2 && c1 <= b2) throw new Error(`You must select all cells in a merged range to merge or unmerge them (${this.getA1Notation()})`);
    });
    this.sheet.merges.add([r1, c1, r2, c2].join(','));
    return this.each((c, i, j) => { if (i || j) { c.v = ''; c.f = ''; } });   // merging keeps only the top-left value
  }
  breakApart() {
    for (const m of [...this.sheet.merges]) {
      const [r1, c1, r2, c2] = m.split(',').map(Number);
      if (r1 >= this.r && r2 <= this.getLastRow() && c1 >= this.c && c2 <= this.getLastColumn()) this.sheet.merges.delete(m);
    }
    return this;
  }
  setBackground(v) { return this.fmt('bg', v); }
  /** Per-cell versions of the setters above, taking a 2D array the size of the range. */
  fmtEach(key, vals) { return this.each((c, i, j) => { c.fmt[key] = plain(vals[i][j]); }); }
  setBackgrounds(vals) { return this.fmtEach('bg', vals); }
  setFontColors(vals) { return this.fmtEach('color', vals); }
  setFontWeights(vals) { return this.fmtEach('weight', vals); }
  setHorizontalAlignments(vals) { return this.fmtEach('halign', vals); }
  setFontFamily(v) { return this.fmt('font', v); }
  setFontColor(v) { return this.fmt('color', v); }
  setFontSize(v) { return this.fmt('size', v); }
  setFontWeight(v) { return this.fmt('weight', v); }
  setFontStyle(v) { return this.fmt('style', v); }
  setVerticalAlignment(v) { return this.fmt('valign', v); }
  setHorizontalAlignment(v) { return this.fmt('halign', v); }
  setWrap(v) { return this.fmt('wrap', v); }
  setNumberFormat(v) { return this.fmt('numfmt', v); }
  setTextRotation(v) { return this.fmt('rotation', v); }
  setBorder(...args) {
    if (args.slice(0, 6).every((a) => a === false)) return this.each((c) => { delete c.fmt.border; });   // clears borders
    return this.fmt('border', args.map(plain));
  }
}

class FakeBook {
  constructor() { this.sheets = []; this.toasts = []; this.active = null; this.timeZone = null; }
  getSheetByName(n) { return this.sheets.find((s) => s.name === n) || null; }
  /** "'Tab'!A2:A" → that range on that tab. */
  getRange(a1) {
    const m = a1.match(/^'?([^'!]+)'?!(.+)$/);
    return this.getSheetByName(m[1]).getRange(m[2]);
  }
  deleteSheet(sh) { this.sheets = this.sheets.filter((s) => s !== sh); }
  insertSheet(n) { const s = new FakeSheet(this, n); this.sheets.push(s); this.active = s; return s; }
  getActiveSheet() { return this.active; }
  setActiveSheet(s) { this.active = s; return s; }
  moveActiveSheet(pos) { const i = this.sheets.indexOf(this.active); this.sheets.splice(i, 1); this.sheets.splice(pos - 1, 0, this.active); }
  toast(msg, title) { this.toasts.push([title || '', msg]); }
  setSpreadsheetTimeZone(tz) { this.timeZone = tz; }
  dump() { return { order: this.sheets.map((s) => s.name), toasts: this.toasts, sheets: Object.fromEntries(this.sheets.map((s) => [s.name, s.dump()])) }; }
}

/** Fresh globals for one run of the script: a spreadsheet, a controllable clock, and the Apps Script services. */
function makeEnvironment(startTime) {
  const book = new FakeBook();
  const RealDate = Date;
  const clock = { now: startTime.getTime() };
  class FakeDate extends RealDate {
    constructor(...a) { if (a.length) super(...a); else super(clock.now); }
    static now() { return clock.now; }
  }
  const triggers = [];
  const menus = [];
  const menu = (title) => {
    const m = { title, items: [] };
    menus.push(m);
    const api = {
      addItem: (label, fn) => { m.items.push([label, fn]); return api; },
      addSeparator: () => { m.items.push(['---']); return api; },
      addSubMenu: (sub) => { m.items.push(['submenu', sub.__title]); return api; },
      addToUi: () => api,
      __title: title,
    };
    return api;
  };
  const triggerBuilder = (fn) => {
    const t = { fn, spec: [] };
    const api = new Proxy({}, { get: (_, k) => (k === 'create' ? () => { triggers.push(t); return { getHandlerFunction: () => fn }; } : (...a) => { t.spec.push([k, ...a]); return api; }) });
    return api;
  };
  const pad = (n) => String(n).padStart(2, '0');
  const DAY = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  const MON = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  const props = new Map();
  const alerts = [];
  const documentProperties = {
    getProperty: (k) => (props.has(k) ? props.get(k) : null),
    setProperty: (k, v) => { props.set(k, String(v)); return documentProperties; },
    setProperties: (o) => { Object.entries(o).forEach(([k, v]) => props.set(k, String(v))); return documentProperties; },
    deleteProperty: (k) => { props.delete(k); return documentProperties; },
  };
  const globals = {
    Date: FakeDate,
    SpreadsheetApp: {
      getActive: () => book,
      getUi: () => ({
        createMenu: menu,
        alert: (...args) => { alerts.push(args.filter((a) => typeof a === 'string')); return 'OK'; },
        Button: { OK: 'OK', CANCEL: 'CANCEL' },
        ButtonSet: { OK: 'OK', OK_CANCEL: 'OK_CANCEL' },
      }),
      newDataValidation: () => recorder('validation'),
      newConditionalFormatRule: () => recorder('cfRule'),
      newRichTextValue: () => recorder('richText'),
      newTextStyle: () => recorder('textStyle'),
      BorderStyle: { SOLID: 'SOLID', SOLID_MEDIUM: 'SOLID_MEDIUM', SOLID_THICK: 'SOLID_THICK' },
      InterpolationType: { NUMBER: 'NUMBER' },
      flush: () => {},
    },
    ScriptApp: {
      getProjectTriggers: () => triggers.map((t) => ({ getHandlerFunction: () => t.fn, __t: t })),
      deleteTrigger: (t) => { const i = triggers.indexOf(t.__t); if (i >= 0) triggers.splice(i, 1); },
      newTrigger: triggerBuilder,
    },
    LockService: { getDocumentLock: () => ({ waitLock() {}, releaseLock() {} }) },
    Session: { getScriptTimeZone: () => 'America/Los_Angeles' },
    PropertiesService: { getDocumentProperties: () => documentProperties },
    Utilities: {
      DigestAlgorithm: { MD5: 'md5' },
      computeDigest: (alg, text) => [...require('crypto').createHash(alg).update(text).digest()],
      base64Encode: (bytes) => Buffer.from(bytes).toString('base64'),
      formatDate: (d, tz, pattern) => pattern.replace(/EEEE|MMMM|MMM|yyyy|d/g, (t) => ({
        EEEE: DAY[d.getDay()], MMMM: MON[d.getMonth()], MMM: MON[d.getMonth()].slice(0, 3), yyyy: String(d.getFullYear()), d: String(d.getDate()),
      })[t]),
    },
    console,
  };
  return { book, clock, globals, triggers, menus, alerts, FakeDate, pad };
}

module.exports = { makeEnvironment, colLetter, colNumber };
