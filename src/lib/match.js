// Where the sentences a Check lists are in the text that was checked, compared as the engine compares them:
// curly quotes as straight, any run of whitespace (line and paragraph breaks too) as one space, any case, and the
// punctuation at either end of a sentence left out. A sentence matches whole words only. One listed twice takes its
// next occurrence; one that isn't there is skipped (the pane still lists it).

const QUOTES = { '‘': "'", '’': "'", '‚': "'", '‛': "'", '′': "'", '“': '"', '”': '"', '„': '"', '‟': '"', '″': '"' };
const CONTROL = /[\u0000-\u0008\u000c\u000e-\u001f]/; // Word's field, cell and picture marks: not text
const WORDISH = /[\p{L}\p{N}]/u;

// The text folded for comparison, and for each folded character the index it came from.
function fold(text) {
  let out = '';
  const at = [];
  for (let i = 0; i < text.length; i++) {
    let c = text[i];
    if (CONTROL.test(c)) continue;
    if (/\s/.test(c)) {
      if (out && !out.endsWith(' ')) { out += ' '; at.push(i); }
      continue;
    }
    c = (QUOTES[c] ?? c).toLowerCase();
    out += c;
    for (let k = 0; k < c.length; k++) at.push(i);
  }
  return { out, at };
}

const key = (sentence) => fold(String(sentence ?? '')).out.replace(/^[^\p{L}\p{N}]+|[^\p{L}\p{N}]+$/gu, '');

// For each sentence, { start, end } in `text` (end exclusive), or null. Spans never overlap.
export function locate(text, sentences) {
  const hay = fold(String(text ?? ''));
  const next = new Map(); // where to look next for a sentence already found
  const taken = [];
  return sentences.map((sentence) => {
    const needle = key(sentence);
    if (!needle) return null;
    for (let i = hay.out.indexOf(needle, next.get(needle) ?? 0); i >= 0; i = hay.out.indexOf(needle, i + 1)) {
      const end = i + needle.length;
      if (WORDISH.test(hay.out[i - 1] ?? '') || WORDISH.test(hay.out[end] ?? '')) continue;
      const span = { start: hay.at[i], end: hay.at[end - 1] + 1 };
      if (taken.some((t) => t.start < span.end && span.start < t.end)) continue;
      next.set(needle, end);
      taken.push(span);
      return span;
    }
    return null;
  });
}
