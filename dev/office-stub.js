// A small stand-in for Office.js and the Word and Excel APIs, for running the pane in headless Chrome
// (dev/e2e.mjs). It covers only what the pane uses; window.__host picks the app ('Word' unless set).
// Word: the document is one string with \r at the end of each paragraph, like Word's range.text, plus a
// style per paragraph. Ranges are live: every edit shifts them, the way Word keeps tracked ranges in
// place. It also enforces Word's rules the pane must follow: a property is readable only after load()
// and sync(), and an object can be reused in a later Word.run only if it was tracked. Excel's model is
// further down.
(() => {
  const HOST = globalThis.__host ?? 'Word';
  const handlers = []; // DocumentSelectionChanged, in either app
  const fireSelection = () => { for (const h of handlers) h({ type: 'documentSelectionChanged' }); };
  const doc = { text: '', styles: [], flags: [], selection: [0, 0], ranges: new Set(), edits: 0 };
  const marks = () => [...doc.text.matchAll(/\r/g)].map((m) => m.index);
  const paraSpans = () => {
    let s = 0;
    return marks().map((e, i) => { const span = { i, start: s, end: e }; s = e + 1; return span; });
  };

  const err = (code, message) => Object.assign(new Error(message || code), { code, name: 'RichApi.Error' });

  // ---- edits -----------------------------------------------------------------------------------
  function edit(s, e, inserted) {
    inserted = inserted.replace(/\r\n?|\n/g, '\r'); // Word turns a line feed into a paragraph mark
    const before = marks();
    const gone = before.filter((m) => m >= s && m < e).length;
    const firstGone = before.findIndex((m) => m >= s);
    // A deleted mark takes its paragraph's style with it; the surviving paragraph keeps the style of
    // the mark at the end. New marks copy the style of the paragraph they split.
    if (gone) {
      doc.styles.splice(firstGone, gone);
      doc.flags.splice(firstGone, gone);
    }
    doc.text = doc.text.slice(0, s) + inserted + doc.text.slice(e);
    const added = (inserted.match(/\r/g) ?? []).length;
    if (added) {
      const at = before.filter((m) => m < s).length;
      doc.styles.splice(at, 0, ...Array(added).fill(doc.styles[at] ?? 'Normal'));
      doc.flags.splice(at, 0, ...Array(added).fill({}));
    }
    // Points before the edit stay, points after it move by the change in length, and a range that
    // ended inside the replaced text now ends after the new text.
    const delta = inserted.length - (e - s);
    const shift = (p, isEnd) => (p <= s ? p : p >= e ? p + delta : isEnd ? s + inserted.length : s);
    for (const r of doc.ranges) {
      r._s = shift(r._s, false);
      r._e = Math.max(r._s, shift(r._e, true));
    }
    doc.edits++;
    return [s, s + inserted.length];
  }

  // ---- proxies ---------------------------------------------------------------------------------
  class Context {
    constructor() { this.objects = new Set(); this.tracked = new Set(); this.pending = new Set(); this.alive = true; }
    get document() { return new DocumentProxy(this); }
    get trackedObjects() { return { add: (o) => this.tracked.add(o), remove: (o) => this.tracked.delete(o) }; }
    async sync() {
      await new Promise((r) => setTimeout(r, 1));
      for (const o of this.pending) o._snapshot();
      this.pending.clear();
    }
  }

  class ClientObject {
    constructor(ctx) { this.context = ctx; this._loaded = new Set(); this._values = {}; ctx.objects.add(this); }
    _check() { if (!this.context.alive && !this.context.tracked.has(this)) throw err('InvalidObjectPath', 'This object was not tracked, so it cannot be used in a later Word.run.'); }
    load(props) { this._check(); String(props || '').split(',').map((p) => p.trim()).filter(Boolean).forEach((p) => this._loaded.add(p)); this.context.pending.add(this); return this; }
    _get(name) { if (!(name in this._values)) throw err('PropertyNotLoaded', `The property '${name}' is not available. Call load() and sync() first.`); return this._values[name]; }
    track() { this.context.tracked.add(this); return this; }
    untrack() { this.context.tracked.delete(this); return this; }
  }

  class RangeProxy extends ClientObject {
    constructor(ctx, s, e, isNull = false) {
      super(ctx);
      this._s = s; this._e = e; this._null = isNull;
      if (!isNull) doc.ranges.add(this);
    }
    _snapshot() {
      for (const p of this._loaded) {
        if (p === 'text') this._values.text = this._null ? '' : doc.text.slice(this._s, this._e);
        if (p === 'isNullObject') this._values.isNullObject = this._null;
        if (p === 'style') this._values.style = doc.styles[paraSpans().find((sp) => sp.end >= this._s)?.i ?? 0];
        if (p === 'items') { /* collections */ }
      }
    }
    get text() { return this._get('text'); }
    get isNullObject() { return this._get('isNullObject'); }
    get style() { return this._get('style'); }
    get paragraphs() {
      this._check();
      const spans = paraSpans().filter((sp) => (this._s === this._e ? sp.start <= this._s && this._s <= sp.end : sp.start < this._e && sp.end >= this._s));
      return new Collection(this.context, spans.map((sp) => new ParagraphProxy(this.context, sp)));
    }
    get tables() { return this._flagged('table'); }
    get inlinePictures() { return this._flagged('picture'); }
    _flagged(flag) {
      this._check();
      const hit = paraSpans().filter((sp) => sp.start < this._e && sp.end >= this._s && doc.flags[sp.i]?.[flag]);
      return new Collection(this.context, hit.map((sp) => ({ flag, i: sp.i })));
    }
    getRange(where = 'Whole') {
      this._check();
      if (where === 'Start') return new RangeProxy(this.context, this._s, this._s);
      if (where === 'End' || where === 'After') return new RangeProxy(this.context, this._e, this._e);
      return new RangeProxy(this.context, this._s, this._e);
    }
    intersectWithOrNullObject(other) {
      this._check();
      const s = Math.max(this._s, other._s);
      const e = Math.min(this._e, other._e);
      return e > s || (e === s && this._s === this._e) ? new RangeProxy(this.context, s, e) : new RangeProxy(this.context, 0, 0, true);
    }
    expandTo(other) {
      this._check();
      return new RangeProxy(this.context, Math.min(this._s, other._s), Math.max(this._e, other._e));
    }
    insertText(text, where) {
      this._check();
      if (where !== 'Replace') throw err('NotImplemented', `stub: insertText(${where})`);
      if (doc.locked) throw err('AccessDenied', 'The document is protected.');
      // Word never deletes the document's last paragraph mark.
      const e = this._e === doc.text.length ? this._e - 1 : this._e;
      const [s2, e2] = edit(this._s, Math.max(this._s, e), text);
      return new RangeProxy(this.context, s2, e2);
    }
    select() {
      this._check();
      doc.selection = [this._s, this._e];
      setTimeout(fireSelection, 5);
    }
  }

  class ParagraphProxy extends RangeProxy {
    constructor(ctx, sp) { super(ctx, sp.start, sp.end + 1); this._i = sp.i; }
    getRange(where = 'Whole') {
      this._check();
      if (where === 'Content') return new RangeProxy(this.context, this._s, this._e - 1);
      if (where === 'End') return new RangeProxy(this.context, this._e - 1, this._e - 1);
      return super.getRange(where);
    }
  }

  class Collection extends ClientObject {
    constructor(ctx, list) { super(ctx); this._list = list; }
    _snapshot() { if (this._loaded.has('items')) this._values.items = this._list; }
    get items() { return this._get('items'); }
  }

  class DocumentProxy {
    constructor(ctx) { this.context = ctx; }
    getSelection() { return new RangeProxy(this.context, ...doc.selection); }
    get body() { const ctx = this.context; return { getRange: () => new RangeProxy(ctx, 0, doc.text.length) }; }
  }

  // ---- Excel -----------------------------------------------------------------------------------
  // A workbook of sheets, each a map of cells "row,col" -> { v, f }. A value keeps its type (string,
  // number, boolean); a formula has its source in f and its result in v. Commands queue up and run in
  // order at sync(), as in Excel, so a read queued after a write in the same batch sees the write. A
  // property is readable only after load() and sync(), and nothing works while a cell is being edited.
  const book = { sheets: [], sel: null, editing: false, locked: false, edits: 0, ids: 0 };
  const colName = (c) => { let s = ''; for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };
  const parseCell = (a) => {
    const m = /^\$?([A-Z]+)\$?(\d+)$/.exec(a);
    if (!m) throw err('InvalidArgument', `stub: not a cell address: ${a}`);
    return { r: Number(m[2]) - 1, c: [...m[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1 };
  };
  const parseBox = (a) => {
    const [x, y = x] = a.split(':');
    const p = parseCell(x);
    const q = parseCell(y);
    return { r0: Math.min(p.r, q.r), c0: Math.min(p.c, q.c), r1: Math.max(p.r, q.r), c1: Math.max(p.c, q.c) };
  };
  const sheetRef = (n) => (/^[A-Za-z0-9_]+$/.test(n) ? n : `'${n.replace(/'/g, "''")}'`);
  const typeOf = (cell) => (!cell || cell.v === '' ? 'Empty' : typeof cell.v === 'number' ? 'Double' : typeof cell.v === 'boolean' ? 'Boolean' : cell.f && String(cell.v).startsWith('#') ? 'Error' : 'String');
  // What Excel makes of a string written to a cell, as if typed: a number stays a number, = + - @ start a
  // formula (here always a #NAME? error), and a leading apostrophe keeps the rest as text.
  function enter(val) {
    if (typeof val !== 'string') return { v: val };
    if (val.startsWith("'")) return { v: val.slice(1) };
    if (val.trim() !== '' && !Number.isNaN(Number(val))) return { v: Number(val) };
    if (/^[=+\-@]/.test(val)) return { v: '#NAME?', f: val.startsWith('=') ? val : `=${val}` };
    return { v: val };
  }

  class XContext {
    constructor() { this.queue = []; this.alive = true; }
    get workbook() { return new XWorkbook(this); }
    async sync() {
      await new Promise((r) => setTimeout(r, 1));
      const queue = this.queue;
      this.queue = [];
      if (book.editing) throw err('InvalidOperationInCellEditMode', 'Excel is in cell-editing mode. Exit the edit mode by pressing ENTER or TAB.');
      for (const op of queue) op();
    }
  }

  class XObject {
    constructor(ctx) { this.context = ctx; this._values = {}; }
    _check() { if (!this.context.alive) throw err('InvalidObjectPath', 'This object belongs to an Excel.run that has finished.'); }
    load(props) {
      this._check();
      const list = String(props || '').split(',').map((p) => p.trim()).filter(Boolean);
      this.context.queue.push(() => { for (const p of list) this._values[p] = this._read(p); });
      return this;
    }
    _get(name) { if (!(name in this._values)) throw err('PropertyNotLoaded', `The property '${name}' is not available. Call load() and sync() first.`); return this._values[name]; }
    get isNullObject() { return this._get('isNullObject'); }
  }

  class XRange extends XObject {
    constructor(ctx, sheet, box) { super(ctx); this._sheet = sheet; this._box = box; }
    _grid(fn) {
      const { r0, c0, r1, c1 } = this._box;
      // Excel refuses to send this many cells in one answer; the pane must not ask for them.
      if ((r1 - r0 + 1) * (c1 - c0 + 1) > 1e6) throw err('ResponsePayloadSizeLimitExceeded', 'The response payload size has exceeded the limit.');
      return Array.from({ length: r1 - r0 + 1 }, (_, i) => Array.from({ length: c1 - c0 + 1 }, (_, j) => fn(this._sheet.cells.get(`${r0 + i},${c0 + j}`))));
    }
    _read(p) {
      if (p === 'isNullObject') return !this._box;
      if (!this._box) return null;
      const { r0, c0, r1, c1 } = this._box;
      const n = (r1 - r0 + 1) * (c1 - c0 + 1);
      if (p === 'address') return `${sheetRef(this._sheet.name)}!${colName(c0)}${r0 + 1}${n === 1 ? '' : `:${colName(c1)}${r1 + 1}`}`;
      if (p === 'cellCount') return n > 2147483647 ? -1 : n;
      if (p === 'values') return this._grid((c) => c?.v ?? '');
      if (p === 'valueTypes') return this._grid(typeOf);
      if (p === 'formulas') return this._grid((c) => c?.f ?? c?.v ?? '');
      throw err('NotImplemented', `stub: Range.${p}`);
    }
    get address() { return this._get('address'); }
    get cellCount() { return this._get('cellCount'); }
    get values() { return this._get('values'); }
    get valueTypes() { return this._get('valueTypes'); }
    get formulas() { return this._get('formulas'); }
    set values(rows) {
      this._check();
      const { r0, c0 } = this._box;
      this.context.queue.push(() => {
        if (book.locked) throw err('AccessDenied', 'The cell or chart you\'re trying to change is on a protected sheet.');
        rows.forEach((row, i) => row.forEach((val, j) => {
          const cell = enter(val);
          if (cell.v === '' && !cell.f) this._sheet.cells.delete(`${r0 + i},${c0 + j}`);
          else this._sheet.cells.set(`${r0 + i},${c0 + j}`, cell);
        }));
        book.edits++;
      });
    }
    get worksheet() { this._check(); return new XSheet(this.context, this._sheet); }
    getIntersectionOrNullObject(other) {
      this._check();
      const a = this._box;
      const b = other._box;
      if (!a || !b || other._sheet !== this._sheet) return new XRange(this.context, this._sheet, null);
      const box = { r0: Math.max(a.r0, b.r0), c0: Math.max(a.c0, b.c0), r1: Math.min(a.r1, b.r1), c1: Math.min(a.c1, b.c1) };
      return new XRange(this.context, this._sheet, box.r0 <= box.r1 && box.c0 <= box.c1 ? box : null);
    }
    select() {
      this._check();
      this.context.queue.push(() => { book.sel = { sheet: this._sheet, ...this._box }; setTimeout(fireSelection, 5); });
    }
  }

  class XSheet extends XObject {
    constructor(ctx, sheet) { super(ctx); this._sheet = sheet; }
    _read(p) {
      if (p === 'isNullObject') return !this._sheet;
      if (!this._sheet) return null;
      if (p === 'id' || p === 'name') return this._sheet[p];
      throw err('NotImplemented', `stub: Worksheet.${p}`);
    }
    get id() { return this._get('id'); }
    get name() { return this._get('name'); }
    getRange(address) { this._check(); return new XRange(this.context, this._sheet, parseBox(address)); }
    getUsedRangeOrNullObject() {
      this._check();
      const at = [...this._sheet.cells.keys()].map((k) => k.split(',').map(Number));
      if (!at.length) return new XRange(this.context, this._sheet, null);
      const rs = at.map(([r]) => r);
      const cs = at.map(([, c]) => c);
      return new XRange(this.context, this._sheet, { r0: Math.min(...rs), c0: Math.min(...cs), r1: Math.max(...rs), c1: Math.max(...cs) });
    }
  }

  class XWorkbook {
    constructor(ctx) { this.context = ctx; }
    getSelectedRange() { const { sheet, ...box } = book.sel; return new XRange(this.context, sheet, box); }
    get worksheets() {
      const ctx = this.context;
      return { getItemOrNullObject: (key) => new XSheet(ctx, book.sheets.find((s) => s.id === key || s.name === key) ?? null) };
    }
  }

  globalThis.Excel = {
    async run(batch) {
      const ctx = new XContext();
      try {
        const out = await batch(ctx);
        await ctx.sync();
        return out;
      } finally {
        ctx.alive = false;
      }
    },
  };

  // ---- globals ---------------------------------------------------------------------------------
  globalThis.Word = {
    async run(a, b) {
      const batch = typeof a === 'function' ? a : b;
      const ctx = typeof a === 'function' ? new Context() : a.context;
      a?._check?.();
      ctx.alive = true;
      try {
        return await batch(ctx);
      } finally {
        ctx.alive = false;
        for (const o of ctx.objects) if (!ctx.tracked.has(o) && o instanceof RangeProxy) doc.ranges.delete(o);
      }
    },
  };

  // Office's theme: Colorful unless the test sets window.__theme (read live, so a test can change it).
  const COLORFUL = { themeId: 'Colorful', isDarkTheme: false, bodyBackgroundColor: '#ffffff', bodyForegroundColor: '#000000', controlBackgroundColor: '#ffffff', controlForegroundColor: '#000000' };
  // Links the pane asked Office to open in the browser (OpenBrowserWindowApi), for the test to read back as
  // window.__opened. window.__noBrowserWindow makes an Office without that API, so the pane falls back to window.open.
  const opened = [];
  globalThis.__opened = opened;
  globalThis.Office = {
    HostType: { Word: 'Word', Excel: 'Excel' },
    EventType: { DocumentSelectionChanged: 'documentSelectionChanged', OfficeThemeChanged: 'officeThemeChanged' },
    onReady: async () => ({ host: HOST, platform: 'PC' }),
    context: {
      requirements: { isSetSupported: (name) => name === (HOST === 'Excel' ? 'ExcelApi' : 'WordApi') || (name === 'OpenBrowserWindowApi' && !globalThis.__noBrowserWindow) },
      get officeTheme() { return globalThis.__theme ?? COLORFUL; },
      document: {
        // The add-in's settings stored in the document; window.__autoshow tags it to open the pane with it.
        settings: { get: (name) => (name === 'Office.AutoShowTaskpaneWithDocument' ? globalThis.__autoshow ?? null : null) },
        // Word and Excel raise only the selection event here; OfficeThemeChanged is Outlook's.
        addHandlerAsync: (type, handler, cb) => {
          if (type !== 'documentSelectionChanged') return cb?.({ status: 'failed', error: { code: 5001, message: 'Unsupported event' } });
          handlers.push(handler);
          cb?.({ status: 'succeeded' });
        },
      },
      ui: { openBrowserWindow: (url) => { opened.push(url); } },
    },
  };

  // ---- the test's handle on the document --------------------------------------------------------
  globalThis.__doc = {
    // paragraphs: strings or { text, style, table, picture }
    set(paragraphs) {
      const ps = paragraphs.map((p) => (typeof p === 'string' ? { text: p } : p));
      doc.text = ps.map((p) => `${p.text}\r`).join('');
      doc.styles = ps.map((p) => p.style ?? 'Normal');
      doc.flags = ps.map((p) => ({ table: !!p.table, picture: !!p.picture }));
      doc.selection = [0, 0];
      doc.edits = 0;
      doc.locked = false;
    },
    // Selects `needle` (the nth match), or a whole paragraph with its mark (a triple-click) when `paragraph` is set.
    select(needle, { paragraph = null, nth = 0 } = {}) {
      if (paragraph !== null) {
        const sp = paraSpans()[paragraph];
        doc.selection = [sp.start, sp.end + 1];
      } else if (needle === null) {
        doc.selection = [0, 0];
      } else {
        let at = -1;
        for (let k = 0; k <= nth; k++) at = doc.text.indexOf(needle, at + 1);
        if (at < 0) throw new Error(`not in the document: ${needle}`);
        doc.selection = [at, at + needle.length];
      }
      fireSelection();
      return doc.text.slice(...doc.selection);
    },
    // The user typing in the document: replaces `needle` with `by`.
    type(needle, by) { const at = doc.text.indexOf(needle); edit(at, at + needle.length, by); },
    lock(on = true) { doc.locked = on; },
    paragraphs: () => doc.text.split('\r').slice(0, -1),
    styles: () => [...doc.styles],
    selected: () => doc.text.slice(...doc.selection),
    edits: () => doc.edits,
    liveRanges: () => doc.ranges.size,
  };

  // The test's handle on the workbook.
  globalThis.__book = {
    // { 'Sheet name': { A1: 'text' | 42 | { v, f } } }
    set(sheets) {
      book.sheets = Object.entries(sheets).map(([name, cells]) => ({
        id: `{0000000${++book.ids}-sheet}`,
        name,
        cells: new Map(Object.entries(cells).map(([a, c]) => { const { r, c: col } = parseCell(a); return [`${r},${col}`, typeof c === 'object' ? { ...c } : { v: c }]; })),
      }));
      book.sel = { sheet: book.sheets[0], ...parseBox('A1') };
      Object.assign(book, { editing: false, locked: false, edits: 0 });
    },
    // Selects `address` on `sheet` (default: the sheet the selection is on).
    select(address, sheet) {
      book.sel = { sheet: sheet ? book.sheets.find((s) => s.name === sheet) : book.sel.sheet, ...parseBox(address) };
      fireSelection();
    },
    // The user typing into a cell of the selection's sheet.
    type(address, value) { const { r, c } = parseCell(address); book.sel.sheet.cells.set(`${r},${c}`, enter(value)); },
    cell(address, sheet) {
      const s = sheet ? book.sheets.find((x) => x.name === sheet) : book.sel.sheet;
      const { r, c } = parseCell(address);
      const cell = s.cells.get(`${r},${c}`);
      return { value: cell?.v ?? '', type: typeOf(cell), formula: cell?.f ?? null };
    },
    rename(from, to) { book.sheets.find((s) => s.name === from).name = to; },
    editing(on = true) { book.editing = on; },
    lock(on = true) { book.locked = on; },
    selected() { const { sheet, r0, c0, r1, c1 } = book.sel; return `${sheet.name}!${colName(c0)}${r0 + 1}${r0 === r1 && c0 === c1 ? '' : `:${colName(c1)}${r1 + 1}`}`; },
    edits: () => book.edits,
  };

  // dev/e2e.mjs sets these before the page loads: the document or workbook, and what the user has selected.
  if (HOST === 'Excel') {
    if (globalThis.__setup) globalThis.__book.set(globalThis.__setup);
    if (globalThis.__select) globalThis.__book.select(...globalThis.__select);
  } else {
    if (globalThis.__setup) globalThis.__doc.set(globalThis.__setup);
    if (globalThis.__select) globalThis.__doc.select(...globalThis.__select);
  }
})();
