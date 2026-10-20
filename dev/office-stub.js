// A small stand-in for Office.js and the Word API, for running the pane in headless Chrome (dev/e2e.mjs).
// It covers only what the pane uses. The document is one string with \r at the end of each paragraph,
// like Word's range.text, plus a style per paragraph. Ranges are live: every edit shifts them, the way
// Word keeps tracked ranges in place. It also enforces Word's rules the pane must follow: a property
// is readable only after load() and sync(), and an object can be reused in a later Word.run only if
// it was tracked.
(() => {
  const doc = { text: '', styles: [], flags: [], selection: [0, 0], ranges: new Set(), handlers: [], edits: 0 };
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

  function fireSelection() {
    for (const h of doc.handlers) h({ type: 'documentSelectionChanged' });
  }

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

  globalThis.Office = {
    HostType: { Word: 'Word' },
    EventType: { DocumentSelectionChanged: 'documentSelectionChanged' },
    onReady: async () => ({ host: 'Word', platform: 'PC' }),
    context: {
      requirements: { isSetSupported: (name) => name === 'WordApi' },
      document: { addHandlerAsync: (type, handler, cb) => { doc.handlers.push(handler); cb?.({ status: 'succeeded' }); } },
      ui: {},
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
  // dev/e2e.mjs sets these before the page loads: the document, and what the user has selected.
  if (globalThis.__setup) globalThis.__doc.set(globalThis.__setup);
  if (globalThis.__select) globalThis.__doc.select(...globalThis.__select);
})();
