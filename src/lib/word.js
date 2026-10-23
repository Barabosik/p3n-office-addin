/* global Word */
// Word glue: read the selection (or the whole document), put a rewrite back where it came from, and mark the
// sentences a Check lists. Word's text ends each paragraph with \r and writes a line break (Shift+Enter) as \v. The
// API reads a blank line as a paragraph break, so each Word paragraph goes over as a paragraph of its own.
import { locate } from './match.js';

const CONTROL = /[\u0000-\u0008\u000c\u000e-\u001f]/g; // field marks, cell marks, picture anchors

export function paragraphsOf(raw) {
  return String(raw ?? '').replace(/\r\n?/g, '\r').split('\r')
    .map((p) => p.replace(/\v/g, '\n').replace(CONTROL, '').trim())
    .filter(Boolean);
}

// The text as the pane shows it and the API receives it.
export const toPlain = (raw) => paragraphsOf(raw).join('\n\n');

// The rewrite back as paragraphs. The API ends it with "\n"; that never becomes an empty paragraph.
export function blocksOf(text) {
  return String(text ?? '').split(/\n[ \t]*\n/).map((b) => b.trim()).filter(Boolean);
}

const words = (t) => (t.match(/\S+/g) ?? []).length;
const same = (a, b) => String(a ?? '').replace(/\r\n?/g, '\r').trimEnd() === String(b ?? '').replace(/\r\n?/g, '\r').trimEnd();

// How much is selected right now, for the "Use selection" label. Reads locally; nothing is sent.
export async function peekSelection() {
  return Word.run(async (ctx) => {
    const sel = ctx.document.getSelection();
    sel.load('text');
    await ctx.sync();
    const text = toPlain(sel.text);
    return { words: words(text), chars: text.length };
  });
}

// Reads the selection, or with `whole` the document when nothing is selected. `ref` keeps a tracked
// reference to that range so Replace can find it again after other edits move it.
export async function capture({ whole = false } = {}) {
  return Word.run(async (ctx) => {
    let range = ctx.document.getSelection();
    range.load('text');
    await ctx.sync();
    let scope = 'selection';
    if (!toPlain(range.text)) {
      if (!whole) return { text: '', scope: 'none' };
      range = ctx.document.body.getRange('Whole');
      range.load('text');
      await ctx.sync();
      scope = 'document';
      if (!toPlain(range.text)) return { text: '', scope: 'empty' };
    }
    range.track();
    await ctx.sync();
    return { text: toPlain(range.text), scope, ref: { range, raw: range.text } };
  });
}

export function release(ref) {
  if (!ref?.range) return;
  Word.run(ref.range, async (ctx) => { ref.range.untrack(); await ctx.sync(); }).catch(() => {});
}

// Puts `rewrite` where `ref` was. Refuses if that text changed since it was read, so it never writes
// over something the user didn't send. Paragraph by paragraph when the counts match, so each keeps its
// style (heading, list item, quote); otherwise one insert that keeps the first paragraph's style. The
// paragraph mark at the end of the range is never touched, so no blank paragraph is added or lost.
// It is all one batch of ordinary Word edits, so Ctrl+Z in the document undoes it.
export async function replace(ref, rewrite) {
  try {
    return await Word.run(ref.range, async (ctx) => {
      const range = ref.range;
      range.load('text');
      const paras = range.paragraphs;
      paras.load('items');
      const tables = range.tables;
      tables.load('items');
      const pictures = range.inlinePictures;
      pictures.load('items');
      await ctx.sync();
      if (!same(range.text, ref.raw)) return { ok: false, reason: 'changed' };
      if (tables.items.length) return { ok: false, reason: 'table' };
      if (pictures.items.length) return { ok: false, reason: 'picture' };

      const parts = paras.items.map((p) => p.getRange('Content').intersectWithOrNullObject(range));
      parts.forEach((p) => p.load('text,isNullObject'));
      await ctx.sync();
      const targets = parts.filter((p) => !p.isNullObject && p.text.replace(CONTROL, '').trim());
      const blocks = blocksOf(rewrite);
      if (!targets.length || !blocks.length) return { ok: false, reason: 'changed' };

      // A line break inside a paragraph goes back as a line break, and the spaces around the selected
      // words stay (a double-click selects the word and the space after it).
      const soft = (b) => b.replace(/\n/g, '\v');
      const edge = (t, s) => t.text.match(/^\s*/)[0] + s + t.text.match(/\s*$/)[0];
      let first;
      let last;
      if (blocks.length === targets.length) {
        const done = targets.map((t, i) => t.insertText(edge(t, soft(blocks[i])), 'Replace'));
        [first, last] = [done[0], done[done.length - 1]];
      } else {
        const span = targets[0].expandTo(targets[targets.length - 1]);
        span.load('text');
        await ctx.sync();
        first = last = span.insertText(edge(span, blocks.map(soft).join('\n')), 'Replace');
      }
      const now = first.expandTo(last);
      now.load('text');
      now.track();
      range.untrack();
      now.select();
      await ctx.sync();
      return { ok: true, ref: { range: now, raw: now.text } };
    });
  } catch (e) {
    // The range was deleted, the document is protected, or the text sits in a locked content control.
    return { ok: false, reason: e?.code === 'ItemNotFound' ? 'gone' : 'locked', detail: e?.message };
  }
}

// ---- marks: Word's annotations ---------------------------------------------------------------
// Word draws an annotation (WordApi 1.7) as a coloured underline over the text; it isn't saved in the document, and
// nothing is written into the text or its formatting. Each one is placed by paragraph and offset into that
// paragraph, so the sentences are found in the checked range's paragraphs (lib/match.js), never with search(), which
// stops at 255 characters. The listed sentences take Red and the pattern hits Berry; Word sets the underline itself.
// Word puts up no popup without strings named in the manifest, so the reason stays in the pane.
const SCHEME = { sentence: 'Red', hit: 'Berry' };

// Marks `items` ({ text, kind }) where they are in the text `ref` was taken from, if that text is unchanged.
// Returns the annotations' ids and a tracked range of its own, for unchanged() and unannotate().
export async function annotate(ref, items) {
  return Word.run(ref.range, async (ctx) => {
    const range = ref.range;
    range.load('text');
    const paras = range.paragraphs;
    paras.load('items');
    await ctx.sync();
    if (!same(range.text, ref.raw)) return { ids: [], ref: null };
    // The part of each paragraph inside the range, and where that part starts in its paragraph (a selection can
    // start partway into one).
    const parts = paras.items.map((p) => ({ p, part: p.getRange('Content').intersectWithOrNullObject(range) }));
    parts.forEach(({ part }) => part.load('text,isNullObject'));
    await ctx.sync();
    const inside = parts.filter(({ part }) => !part.isNullObject);
    const leads = inside.map(({ p, part }) => p.getRange('Start').expandTo(part.getRange('Start')));
    leads.forEach((l) => l.load('text'));
    await ctx.sync();

    let at = 0;
    const segs = inside.map(({ p, part }, i) => {
      const seg = { p, from: at, to: at + part.text.length, offset: leads[i].text.length, critiques: [] };
      at = seg.to + 1;
      return seg;
    });
    const spans = locate(inside.map(({ part }) => part.text).join('\r'), items.map((x) => x.text));
    spans.forEach((span, i) => {
      for (const seg of span ? segs : []) {
        const s = Math.max(span.start, seg.from);
        const e = Math.min(span.end, seg.to);
        if (e > s) seg.critiques.push({ colorScheme: SCHEME[items[i].kind], start: seg.offset + s - seg.from, length: e - s });
      }
    });
    // A paragraph at a time, so one Word refuses costs only its own marks and every id that's made is known.
    const ids = [];
    for (const seg of segs.filter((x) => x.critiques.length)) {
      try {
        const made = seg.p.insertAnnotations({ critiques: seg.critiques });
        await ctx.sync();
        ids.push(...(made.value ?? []));
      } catch { /* left to the list */ }
    }
    const own = range.getRange('Whole');
    own.track();
    await ctx.sync();
    return { ids, ref: { range: own, raw: ref.raw } };
  });
}

// Removes the annotations and lets go of the range. One that's gone already (its text was deleted) is skipped.
export async function unannotate(marks) {
  release(marks?.ref);
  if (!marks?.ids?.length) return;
  await Word.run(async (ctx) => {
    for (const id of marks.ids) {
      ctx.document.getAnnotationById(id).delete();
      try { await ctx.sync(); } catch { /* not there any more */ }
    }
  }).catch(() => {});
}

// Whether the marked text still reads as it did when it was checked.
export async function unchanged(ref) {
  try {
    return await Word.run(ref.range, async (ctx) => {
      ref.range.load('text');
      await ctx.sync();
      return same(ref.range.text, ref.raw);
    });
  } catch {
    return false;
  }
}

// Calls `fn` when paragraphs change, are added or are deleted, by the user or a co-author (WordApi 1.6).
export async function onEdit(fn) {
  await Word.run(async (ctx) => {
    ctx.document.onParagraphChanged.add(fn);
    ctx.document.onParagraphAdded.add(fn);
    ctx.document.onParagraphDeleted.add(fn);
    await ctx.sync();
  });
}
