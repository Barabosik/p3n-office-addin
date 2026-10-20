/* global Excel */
// Excel glue: read the selected cells' text, and write a rewrite back into the one cell it came from.
// Each cell with text goes over as a paragraph of its own (the API reads a blank line as a paragraph
// break). Numbers, dates and empty cells stay out. A rewrite goes back only when the text came from a
// single cell that holds plain text, so it never lands over a formula or spreads across cells.

const words = (t) => (t.match(/\S+/g) ?? []).length;

// Far more text than the API takes. A bigger selection (a whole column of a long sheet) isn't read.
const MAX_CELLS = 10000;

// "Sheet1!B3:C4" or "'Q3 plan'!B3" -> the first cell's column and row (0-based).
function origin(address) {
  const m = /([A-Z]+)(\d+)(?::|$)/.exec(String(address).split('!').pop().replace(/\$/g, ''));
  if (!m) return { col: 0, row: 0 };
  const col = [...m[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  return { col, row: Number(m[2]) - 1 };
}
function cellName(col, row) {
  let s = '';
  for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return `${s}${row + 1}`;
}

// A formula shows its result in `values` and its source in `formulas`; plain text has the same in both.
const isFormula = (f, v) => typeof f === 'string' && f.startsWith('=') && f !== String(v);

// The cells with text in the selection, read without loading the empty rest of a big selection.
async function readSelection(ctx) {
  const sel = ctx.workbook.getSelectedRange();
  const sheet = sel.worksheet;
  const used = sheet.getUsedRangeOrNullObject(true);
  sel.load('cellCount');
  sheet.load('id,name');
  used.load('isNullObject');
  await ctx.sync();
  if (used.isNullObject) return { cells: [], sheet };
  const part = sel.getIntersectionOrNullObject(used);
  part.load('isNullObject,cellCount');
  await ctx.sync();
  if (part.isNullObject) return { cells: [], sheet };
  if (part.cellCount < 0 || part.cellCount > MAX_CELLS) return { cells: [], sheet, tooBig: true };
  part.load('address,values,valueTypes,formulas');
  await ctx.sync();
  const at = origin(part.address);
  const cells = [];
  part.values.forEach((row, r) => row.forEach((v, c) => {
    if (part.valueTypes[r][c] !== 'String' || !String(v).trim()) return;
    cells.push({ value: v, text: String(v).replace(/\r\n?/g, '\n').trim(), formula: isFormula(part.formulas[r][c], v), cell: cellName(at.col + c, at.row + r) });
  }));
  return { cells, sheet };
}

// How much is selected right now, for the "Use selection" label. Reads locally; nothing is sent.
export async function peekSelection() {
  return Excel.run(async (ctx) => {
    const { cells, tooBig } = await readSelection(ctx);
    const text = cells.map((c) => c.text).join('\n\n');
    return { words: words(text), chars: text.length, tooBig: !!tooBig };
  });
}

// Reads the selected cells. `ref` is where a rewrite can go back: set only when the text is one cell's.
export async function capture() {
  return Excel.run(async (ctx) => {
    const { cells, sheet, tooBig } = await readSelection(ctx);
    if (tooBig) return { text: '', scope: 'too_big' };
    if (!cells.length) return { text: '', scope: 'none' };
    const text = cells.map((c) => c.text).join('\n\n');
    const one = cells.length === 1 ? cells[0] : null;
    const why = !one ? 'cells' : one.formula ? 'formula' : null;
    return {
      text,
      scope: 'selection',
      where: one?.cell ?? null,
      ref: why ? null : { sheetId: sheet.id, cell: one.cell, value: one.value },
      why,
    };
  });
}

export function release() {}

// Excel would read a rewrite that starts with = + - or @ as a formula, and one like "2029" as a number.
// The apostrophe in front keeps it text, as when you type it, and isn't part of the value.
const asText = (t) => (/^[=+\-@]/.test(t) || (t.trim() !== '' && !Number.isNaN(Number(t))) ? `'${t}` : t);

// Writes `rewrite` into the cell `ref` points at, if that cell still holds what was read from it.
// It's one value change in the cell, which keeps the cell's formatting.
export async function replace(ref, rewrite) {
  try {
    return await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItemOrNullObject(ref.sheetId);
      sheet.load('isNullObject');
      await ctx.sync();
      if (sheet.isNullObject) return { ok: false, reason: 'gone' };
      const cell = sheet.getRange(ref.cell);
      cell.load('values,formulas');
      await ctx.sync();
      if (isFormula(cell.formulas[0][0], cell.values[0][0])) return { ok: false, reason: 'formula' };
      if (cell.values[0][0] !== ref.value) return { ok: false, reason: 'changed' };
      cell.values = [[asText(rewrite)]];
      cell.select();
      cell.load('values');
      await ctx.sync();
      return { ok: true, ref: { ...ref, value: cell.values[0][0] }, where: ref.cell };
    });
  } catch (e) {
    // Excel refuses API edits while a cell is being typed in, and on a protected sheet.
    return { ok: false, reason: e?.code === 'InvalidOperationInCellEditMode' ? 'editing' : e?.code === 'ItemNotFound' ? 'gone' : 'locked', detail: e?.message };
  }
}
