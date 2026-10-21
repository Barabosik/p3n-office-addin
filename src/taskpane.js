/* global Office */
import { runSkill, me, setClient, ApiError, LINKS, MAX_CHARS } from './lib/api.js';
import { apiKey, freeLeft } from './lib/store.js';
import * as word from './lib/word.js';
import * as excel from './lib/excel.js';
import { followOfficeTheme } from './lib/theme.js';

const $ = (s) => document.querySelector(s);
const els = {
  text: $('#text'), count: $('#count'), check: $('#check'), rewrite: $('#rewrite'), grab: $('#grab'), hint: $('#hint'),
  busy: $('#busy'), busyText: $('#busy-text'), cancel: $('#cancel'), result: $('#result'), quota: $('#quota'), account: $('#account'),
  main: $('#main'), settings: $('#settings'), back: $('#back'), key: $('#key'), keyStatus: $('#key-status'), remove: $('#remove'),
};
// host: the Word or Excel glue and words below, once Office says where the pane runs.
// source: where the text came from ({ sent, ref, where, why }); ref is the glue's handle on it for Replace,
// and null when it can't go back there (why says why).
const state = { key: apiKey.get(), host: null, source: null, controller: null, timer: null, free: freeLeft.get(), selected: null };

// What differs between Word and Excel: the glue, the API version it needs, and the words for where text lives.
const HOSTS = {
  Word: {
    glue: word, api: ['WordApi', '1.3'],
    placeholder: 'Select text in your document and click Use selection. Or paste it here.',
    nothing: 'Use whole document',
    hints: {
      none: 'Nothing is selected. Select text in the document first.',
      empty: 'The document is empty. Type or paste text in the box.',
      document: 'Took the whole document, since nothing was selected.',
      failed: 'p3n can\'t read the document right now. Copy the text and paste it in the box.',
      old: 'This version of Word is too old for p3n to read the document. Paste the text in the box.',
    },
    replace: () => 'Replace in document',
    replaced: () => `Replaced in the document and selected there. ${undoKey()} in the document undoes it.`,
    refused: {
      changed: 'That text changed in the document since you took it, so p3n left it alone.',
      gone: 'That text is no longer in the document.',
      table: 'That part of the document has a table, so p3n won\'t paste over it.',
      picture: 'That part of the document has a picture, so p3n won\'t paste over it.',
      locked: 'Word didn\'t allow the edit there (the document or that part of it may be protected).',
    },
    retry: 'Copy the rewrite and paste it instead, or select the text again and click Use selection.',
  },
  Excel: {
    glue: excel, api: ['ExcelApi', '1.4'],
    placeholder: 'Select cells and click Use selection. Or paste text here.',
    nothing: 'Use selection',
    hints: {
      none: 'The selected cells have no text. Select cells with text first.',
      too_big: 'That selection is too big to read. Select fewer cells.',
      failed: 'p3n can\'t read the sheet right now. If you\'re typing in a cell, press Enter, then try again.',
      old: 'This version of Excel is too old for p3n to read the sheet. Paste the text in the box.',
    },
    replace: (where) => `Replace in ${where}`,
    replaced: (where) => `Replaced the text in ${where} and selected it. ${undoKey()} in the sheet undoes it.`,
    // Why a rewrite of these cells has no Replace button.
    kept: {
      cells: 'p3n writes a rewrite back only into a single cell. Copy it, or select one cell and click Use selection.',
      formula: 'That cell holds a formula, so p3n won\'t write over it. Copy the rewrite instead.',
    },
    refused: {
      changed: 'That cell changed since you took it, so p3n left it alone.',
      gone: 'That sheet is no longer in the workbook.',
      formula: 'That cell holds a formula now, so p3n won\'t write over it.',
      editing: 'Excel is busy editing a cell. Press Enter or Esc there, then try again.',
      locked: 'Excel didn\'t allow the edit (the sheet may be protected).',
    },
    retry: 'Copy the rewrite and paste it instead, or select the cell again and click Use selection.',
  },
};

// Tiny DOM builder. Strings become text nodes, so API text is never parsed as HTML. Styles go through
// the CSSOM rather than style attributes.
function h(tag, attrs = {}, ...kids) {
  const el = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (v === false || v == null) continue;
    if (k.startsWith('on')) el.addEventListener(k.slice(2), v);
    else if (k === 'class') el.className = v;
    else if (k === 'style') el.style.cssText = v;
    else el.setAttribute(k, v === true ? '' : v);
  }
  el.append(...kids.flat().filter((k) => k !== null && k !== undefined && k !== false));
  return el;
}

const pct = (x) => Math.round((x || 0) * 100);
const tone = (x) => (x >= 0.7 ? 'ai' : x >= 0.4 ? 'mid' : 'human');
const LABELS = { human: 'Likely human-written', ai: 'Likely AI-written', ai_edited: 'AI-written, then edited', humanized: 'AI-written, then humanized', mixed: 'Mixed: part human, part AI' };
const words = (t) => (t || '').trim().split(/\s+/).filter(Boolean).length;
const clip = (t, n = 220) => (t && t.length > n ? t.slice(0, n - 1) + '…' : t);
const limit = () => (state.key ? MAX_CHARS.key : MAX_CHARS.free);
const calm = matchMedia('(prefers-reduced-motion: reduce)');
const undoKey = () => (/Mac/.test(navigator.platform) ? 'Cmd+Z' : 'Ctrl+Z');

// Links open in the browser. In Word on Windows and Mac a plain link would try to load in the pane.
function openLink(url) {
  const office = globalThis.Office;
  if (office?.context?.ui?.openBrowserWindow && office.context.requirements?.isSetSupported('OpenBrowserWindowApi', '1.1')) {
    office.context.ui.openBrowserWindow(url);
  } else {
    window.open(url, '_blank', 'noopener');
  }
}
document.addEventListener('click', (e) => {
  const a = e.target.closest?.('a[href]');
  if (!a || !/^https:/.test(a.href)) return;
  e.preventDefault();
  openLink(a.href);
});

// ---- motion ------------------------------------------------------------------------------

// The one moment that moves: the score runs from where the text was to where it is now (0 for a check, the
// score before for a rewrite), and the number and the bar take the colour of each band they pass through.
// The markup already holds the end state, so reduced motion or a failed frame shows the right result.
function runScore(root) {
  const score = root.querySelector('[data-to]');
  const num = score?.querySelector('.num');
  if (!num) return;
  const pctEl = score.querySelector('.pct');
  const bar = root.querySelector('.meter i');
  const from = Number(score.dataset.from);
  const to = Number(score.dataset.to);
  if (calm.matches || from === to) return;
  const draw = (v, t = tone(v / 100)) => {
    num.textContent = `${Math.round(v)}%`;
    pctEl.classList.remove('t-ai', 't-mid', 't-human');
    pctEl.classList.add(`t-${t}`);
    if (bar) bar.className = `b-${t}`;
  };
  const scale = (v) => { if (bar) bar.style.transform = `scaleX(${Math.max(3, v) / Math.max(3, to)})`; };
  const t0 = performance.now() + 120;
  const ms = 800;
  draw(from);
  scale(from);
  const step = (now) => {
    if (!num.isConnected) return;
    const p = Math.min(1, Math.max(0, (now - t0) / ms));
    const v = from + (to - from) * (1 - (1 - p) ** 4);
    if (p < 1) {
      draw(v);
      scale(v);
      requestAnimationFrame(step);
    } else {
      draw(to, score.dataset.tone);
      if (bar) bar.style.transform = '';
    }
  };
  requestAnimationFrame(step);
}

// The pane is short: a result that lands below the fold scrolls up into view, so the score and
// Replace are seen without hunting for them.
function reveal(el) {
  if (el.getBoundingClientRect().top > innerHeight * 0.55) el.scrollIntoView({ block: 'start', behavior: calm.matches ? 'auto' : 'smooth' });
}

// Result parts arrive in reading order.
function stagger(box) {
  [...box.children].forEach((c, i) => c.style.setProperty('--i', i));
  return box;
}

// Text arriving from the document lands in the box with a brief highlight.
function arrived() {
  els.text.classList.remove('arrived');
  void els.text.offsetWidth; // restart the animation
  els.text.classList.add('arrived');
}

// A check mark that draws itself, for Copied and Replaced.
function checkMark() {
  const ns = 'http://www.w3.org/2000/svg';
  const svg = document.createElementNS(ns, 'svg');
  const path = document.createElementNS(ns, 'path');
  for (const [k, v] of Object.entries({ viewBox: '0 0 16 16', class: 'tick', 'aria-hidden': 'true' })) svg.setAttribute(k, v);
  for (const [k, v] of Object.entries({ d: 'M3 8.5l3.2 3L13 4.5', fill: 'none', stroke: 'currentColor', 'stroke-width': '2', 'stroke-linecap': 'round', 'stroke-linejoin': 'round' })) path.setAttribute(k, v);
  svg.append(path);
  return svg;
}

// ---- input -------------------------------------------------------------------------------

function refreshInput() {
  const n = els.text.value.trim().length;
  els.count.textContent = n ? `${n.toLocaleString()} / ${limit().toLocaleString()} characters` : '';
  els.count.classList.toggle('t-ai', n > limit());
  const off = (!n && !state.selected?.words) || !!state.controller;
  els.check.disabled = off;
  els.rewrite.disabled = off;
}

// The selection link says what it will take: the selection and its size, or (in Word) the whole document.
function refreshGrab() {
  els.grab.hidden = !state.host;
  const w = state.selected?.words ?? 0;
  els.grab.textContent = w ? `Use selection (${w.toLocaleString()} ${w === 1 ? 'word' : 'words'})` : state.host?.nothing ?? '';
}

function hint(msg) {
  els.hint.textContent = msg || '';
  els.hint.hidden = !msg;
}

// Reads the selection (or the document) into the box and remembers where it came from, for Replace.
async function grab({ whole = true } = {}) {
  const { glue, hints } = state.host;
  hint('');
  try {
    const got = await glue.capture({ whole });
    if (!got.text) {
      hint(hints[got.scope] ?? hints.none);
      return null;
    }
    glue.release(state.source?.ref);
    state.source = { sent: got.text, ref: got.ref, where: got.where, why: got.why };
    els.text.value = got.text;
    arrived();
    refreshInput();
    if (got.scope === 'document') hint(hints.document);
    return got.text;
  } catch {
    hint(hints.failed);
    return null;
  }
}

els.text.addEventListener('input', refreshInput);
els.check.addEventListener('click', () => go('detect'));
els.rewrite.addEventListener('click', () => go('humanize'));
els.cancel.addEventListener('click', () => state.controller?.abort());
els.grab.addEventListener('click', () => grab());
els.text.addEventListener('keydown', (e) => {
  if (e.key === 'Enter' && (e.ctrlKey || e.metaKey) && !els.check.disabled) { e.preventDefault(); go('detect'); }
});

// With the box empty, Check and Rewrite take the selection first.
async function go(skill) {
  let text = els.text.value;
  if (!text.trim() && state.host) text = await grab({ whole: false });
  if (text) start(skill, text);
}

// Selection changes in the document update the link's word count (read locally, nothing is sent).
let peekTimer;
function watchSelection() {
  const peek = async () => {
    try { state.selected = await state.host.glue.peekSelection(); } catch { state.selected = null; }
    refreshGrab();
    refreshInput();
  };
  Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, () => {
    clearTimeout(peekTimer);
    peekTimer = setTimeout(peek, 200);
  });
  return peek();
}

// ---- running a skill ---------------------------------------------------------------------

async function start(skill, text, options) {
  text = (text || '').trim();
  if (!text) return;
  state.controller?.abort();
  if (text.length > limit()) {
    return showError(new ApiError(422, 'too_long', state.key
      ? `That's ${text.length.toLocaleString()} characters; the API takes up to ${MAX_CHARS.key.toLocaleString()}. Select less.`
      : `That's ${text.length.toLocaleString()} characters. Without an API key p3n takes up to ${MAX_CHARS.free.toLocaleString()}. Select less, or add a key.`));
  }
  const controller = new AbortController();
  state.controller = controller;
  els.result.replaceChildren();
  setBusy(skill);
  try {
    const { run, free } = await runSkill(skill, text, { key: state.key, options, signal: controller.signal });
    if (free !== null) { state.free = free; freeLeft.set(free); }
    // The progress bar goes before the result scrolls into view; gone after, it would pull the score up
    // out of sight.
    setBusy(null);
    render(run, text, skill);
  } catch (e) {
    setBusy(null);
    if (e.name !== 'AbortError') showError(e);
    if (e.free != null) { state.free = e.free; freeLeft.set(e.free); }
  } finally {
    if (state.controller === controller) state.controller = null;
    setBusy(null);
    refreshInput();
    refreshAccount();
  }
}

function setBusy(skill) {
  clearInterval(state.timer);
  els.busy.hidden = !skill;
  if (!skill) return;
  const t0 = Date.now();
  const what = skill === 'detect' ? 'Checking' : 'Rewriting';
  const tick = () => {
    const s = Math.round((Date.now() - t0) / 1000);
    els.busyText.textContent = `${what}… ${s} s` + (skill === 'humanize' && s > 8 ? ' (a rewrite takes up to a minute)' : '');
  };
  tick();
  state.timer = setInterval(tick, 1000);
}

// ---- results -----------------------------------------------------------------------------

function render(run, text, skill) {
  const o = run.output;
  // An answer without what the pane shows is a failed run to the user, not a 0% score or a script error.
  if ((skill === 'detect' && typeof o?.ai !== 'number') || (skill === 'humanize' && typeof o?.text !== 'string')) {
    return showError(new ApiError(502, 'bad_output', 'p3n sent an answer this version of the add-in can\'t read. Try again in a moment.'));
  }
  const body = skill === 'detect' ? renderDetect(o) : skill === 'humanize' ? renderHumanize(o, text) : [h('div', { class: 'rewrite' }, o?.text ?? '')];
  const meta = h('p', { class: 'small muted' }, `${run.output?.version ?? run.skill}${run.cached ? ' · from cache' : ''}${state.key && run.charged_usd ? ` · $${Number(run.charged_usd).toFixed(4)}` : ''}`);
  const box = stagger(h('div', { class: 'result-in stack' }, body, meta));
  els.result.replaceChildren(box);
  runScore(box);
  reveal(els.result);
}

// The big number: animated for the eye, stated once for screen readers (the result is a live region).
const scoreNum = (p) => [h('span', { class: 'num', 'aria-hidden': 'true' }, `${p}%`), h('span', { class: 'sr' }, `${p}%`)];
const meter = (x, t) => h('div', { class: 'meter' }, h('i', { class: `b-${t}`, style: `width:${Math.max(3, pct(x))}%` }));

function renderDetect(o) {
  const t = tone(o.ai);
  const parts = [
    h('div', { class: 'score', 'data-from': 0, 'data-to': pct(o.ai), 'data-tone': t },
      h('p', { class: 'label' }, LABELS[o.label] ?? 'Unknown'),
      h('p', { class: `pct t-${t}` }, scoreNum(pct(o.ai)), h('small', {}, 'likely AI'))),
    meter(o.ai, t),
  ];
  if (o.verdict) parts.push(h('p', {}, o.verdict));
  if (o.confidence?.level) parts.push(h('p', { class: 'small muted' }, `Confidence ${o.confidence.level}: ${o.confidence.reason ?? ''}`));
  // The pane is narrow: the sentences come first, and a pattern hit is listed only if its sentence isn't already there.
  const norm = (t) => (t || '').replace(/\s+/g, ' ').trim();
  const shown = new Set((o.sentences ?? []).slice(0, 5).map((s) => norm(s.text)));
  if (shown.size) {
    parts.push(h('div', {}, h('h3', {}, 'Sentences that read as AI'),
      h('ul', { class: 'quotes' }, o.sentences.slice(0, 5).map((s, j) => h('li', { class: 'hit', style: `--j:${j}` }, s.why ? h('span', { class: 'why' }, s.why) : null, h('q', {}, clip(s.text)))))));
  }
  const hits = (o.examples ?? []).filter((e) => !shown.has(norm(e.sentence))).slice(0, 6);
  if (hits.length) {
    parts.push(h('div', {}, h('h3', {}, 'Pattern hits'),
      h('ul', { class: 'quotes' }, hits.map((e) => h('li', {},
        h('span', { class: 'why' }, (e.rules ?? []).map((r) => String(r.name ?? '').replace(/\s*\(LLM\)$/, '')).join(' · ')),
        h('q', {}, clip(e.sentence)))))));
  }
  parts.push(h('p', { class: 'small muted' }, 'A score is a probability, not proof. Read the sentences yourself before deciding anything about a person.'));
  return parts;
}

async function copyText(text) {
  try {
    await navigator.clipboard.writeText(text);
    return true;
  } catch {
    // Word on the web runs the pane in a frame that may not allow the clipboard API.
    const ta = h('textarea', { class: 'sr', 'aria-hidden': 'true' });
    ta.value = text;
    document.body.append(ta);
    ta.select();
    const done = document.execCommand('copy');
    ta.remove();
    return done;
  }
}

function renderHumanize(o, sent) {
  const t = tone(o.ai_after);
  // The rewrite ends with a line break; the pane shows it without, and Replace never adds a paragraph for it.
  const rewrite = o.text.trim();
  // Only offer Replace when what was rewritten is still what came from the document, not text edited here,
  // and the glue can put it back there.
  const fromDoc = !!state.source && state.source.sent.trim() === sent;
  const replaceable = fromDoc && !!state.source.ref;
  const status = h('p', { class: 'small', hidden: true, 'data-status': true });
  const copyBtn = h('button', { type: 'button', class: replaceable ? 'btn ghost' : 'btn accent' }, 'Copy');
  copyBtn.addEventListener('click', async () => {
    if (!(await copyText(rewrite))) return;
    copyBtn.replaceChildren(checkMark(), 'Copied');
    clearTimeout(copyBtn.reset);
    copyBtn.reset = setTimeout(() => copyBtn.replaceChildren('Copy'), 1500);
  });
  const buttons = [];
  if (replaceable) {
    const rep = h('button', { type: 'button', class: 'btn accent' }, state.host.replace(state.source.where));
    rep.addEventListener('click', () => replaceInDocument(rewrite, sent, rep, status));
    buttons.push(rep);
  }
  buttons.push(copyBtn, h('button', { type: 'button', class: 'link small push', onclick: () => start('humanize', sent, { fresh: true }) }, 'Try another rewrite'));

  // "No longer" and "Still" only when the detector model put the text sent in the AI class: it can score
  // text at 0% that Check called likely AI, and "No longer reads as AI, 0% → 2%" reads wrong.
  const before = o.ai_before != null ? pct(o.ai_before) : pct(o.ai_after);
  const wasAI = o.ai_before != null && o.ai_before >= 0.5;
  const headline = o.passes ? (wasAI ? 'No longer reads as AI' : "Doesn't read as AI") : (wasAI ? 'Still reads as AI' : 'Reads as AI');
  const parts = [
    h('div', { class: 'score', 'data-from': before, 'data-to': pct(o.ai_after), 'data-tone': t },
      h('div', {}, h('p', { class: 'label' }, headline),
        h('p', { class: 'small muted' }, `Detector model: ${LABELS[o.label_after] ?? 'unknown'}`)),
      h('p', { class: `pct t-${t}`, title: 'Detector model, AI class' },
        o.ai_before != null ? h('span', { class: 'from' }, `${before}% → `) : null,
        scoreNum(pct(o.ai_after)), h('small', {}, 'AI class'))),
    meter(o.ai_after, t),
    h('div', { class: 'rewrite' }, rewrite),
    h('div', { class: 'row' }, buttons),
    status,
  ];
  const kept = fromDoc && !replaceable && state.host?.kept?.[state.source.why];
  if (kept) parts.push(h('p', { class: 'small muted' }, kept));
  if (words(o.text) / Math.max(1, words(sent)) < 0.6) {
    parts.push(h('p', { class: 'note warn small' }, `Kept ${words(o.text)} of ${words(sent)} words. Most of this text was general statements with no facts in them, and those are what reads as AI. Add the specifics you meant (who, what, numbers) and run it again.`));
  }
  if (o.invented_blocked?.length) {
    parts.push(h('p', { class: 'small muted' }, `Kept as sent where a rewrite added ${o.invented_blocked.join(', ')}, which the text does not have.`));
  }
  parts.push(h('p', { class: 'small muted' }, 'Humanize drops filler and general statements that carry no fact. Check that every claim you need is still there.'));
  return parts;
}

async function replaceInDocument(text, sent, button, status) {
  const s = state.source;
  const host = state.host;
  const say = (msg, kind) => { status.textContent = msg; status.className = `note small ${kind}`; status.hidden = false; };
  // Use selection took in other text since this rewrite: its ref points there, not at the text rewritten.
  if (s?.sent.trim() !== sent) return say('You took other text from the document since this rewrite. Rewrite that text, or copy this one.', 'warn');
  button.disabled = true;
  const r = s?.ref ? await host.glue.replace(s.ref, text) : { ok: false, reason: 'gone' };
  if (r.ok) {
    // The document now holds this rewrite; another rewrite of the same text can replace it in turn.
    state.source = { ...s, ref: r.ref };
    button.replaceChildren(checkMark(), 'Replaced');
    button.className = 'btn ok';
    return say(host.replaced(s.where), 'ok');
  }
  button.disabled = false;
  // Excel busy in a cell is a wait, not a dead end: there's nothing to copy.
  if (r.reason === 'editing') return say(host.refused.editing, 'warn');
  say(`${host.refused[r.reason] ?? 'p3n couldn\'t put the text back.'} ${host.retry}`, 'warn');
}

function showError(e) {
  const kids = [];
  const keyBtn = (label) => h('button', { type: 'button', class: 'btn ghost', onclick: () => openSettings() }, label);
  const linkBtn = (label, href) => h('a', { class: 'btn accent', href, style: 'text-decoration:none' }, label);
  switch (e.type) {
    case 'free_limit': {
      const mins = e.retryAfter ? Math.max(1, Math.ceil(e.retryAfter / 60)) : null;
      kids.push(h('p', {}, `The free checks for this hour are used up${mins ? `; more in about ${mins} min` : ''}.`),
        h('p', { class: 'small' }, 'Sign up for free credit and an API key, then paste the key here to keep going.'),
        h('div', { class: 'row' }, linkBtn('Sign up', e.signupUrl || LINKS.signup), keyBtn('Add an API key')));
      break;
    }
    case 'insufficient_credit':
      kids.push(h('p', {}, 'Your p3n balance is too low for this.'), h('div', { class: 'row' }, linkBtn('Top up', LINKS.billing)));
      break;
    case 'unauthorized':
      kids.push(h('p', {}, 'p3n didn\'t accept your API key. It may have been revoked.'), h('div', { class: 'row' }, keyBtn('Change the key')));
      break;
    case 'rate_limited':
      kids.push(h('p', {}, 'Too many requests at once. Wait a minute and try again.'));
      break;
    default:
      kids.push(h('p', {}, e.message || 'Something went wrong.'));
  }
  els.result.replaceChildren(stagger(h('div', { class: 'note error result-in stack tight' }, kids)));
  reveal(els.result);
}

// ---- account / quota ---------------------------------------------------------------------

async function refreshAccount() {
  els.account.textContent = state.key ? 'API key' : 'Add API key';
  if (!state.key) {
    els.quota.textContent = state.free != null
      ? `${state.free} free ${state.free === 1 ? 'call' : 'calls'} left this hour without a key.`
      : 'Free without a key: 10 calls an hour, up to 5,000 characters each.';
    return;
  }
  try {
    const m = await me(state.key);
    els.quota.textContent = `Balance $${Number(m.balance_usd).toFixed(2)} · key ${m.key?.name ?? m.key?.prefix ?? ''}`;
  } catch (e) {
    els.quota.textContent = e.type === 'unauthorized' ? 'Your API key wasn\'t accepted. Change it under API key.' : '';
  }
}

// ---- settings: the API key, checked with p3n before it's saved -----------------------------

// The settings view slides over the main one; with View Transitions where the WebView has them.
function swap(toSettings) {
  const apply = () => {
    els.main.hidden = toSettings;
    els.settings.hidden = !toSettings;
    els.quota.hidden = toSettings;
    els.account.hidden = toSettings;
    document.body.dataset.view = toSettings ? 'settings' : 'main';
  };
  if (document.startViewTransition && !calm.matches) return document.startViewTransition(apply).updateCallbackDone;
  apply();
  return Promise.resolve();
}

function openSettings() {
  els.key.value = state.key;
  say('');
  swap(true).then(() => els.key.focus());
  if (state.key) checkAndSave(state.key);
}

function closeSettings() {
  swap(false).then(() => els.account.focus());
  refreshInput();
  refreshAccount();
}

els.account.addEventListener('click', openSettings);
els.back.addEventListener('click', closeSettings);
els.settings.addEventListener('keydown', (e) => { if (e.key === 'Escape') closeSettings(); });

function say(msg, kind = '') {
  els.keyStatus.textContent = msg;
  els.keyStatus.className = `note small ${kind}`;
  els.keyStatus.hidden = !msg;
  if (msg && !calm.matches) {
    els.keyStatus.animate([{ opacity: 0, transform: 'translateY(4px)' }, {}], { duration: 200, easing: 'cubic-bezier(0.22, 1, 0.36, 1)' });
  }
}

// Saves a key only if p3n accepts it, so a typo never replaces a good key.
let pending = 0;
async function checkAndSave(k) {
  const mine = ++pending;
  say('Checking the key…');
  try {
    const m = await me(k);
    if (mine !== pending) return;
    if (k !== state.key && !apiKey.set(k)) return say('The key works, but this device won\'t let p3n store it. Paste it again next time.', 'warn');
    state.key = k;
    say(`Saved. Signed in as ${m.email}, balance $${Number(m.balance_usd).toFixed(2)}.`, 'ok');
  } catch (e) {
    if (mine !== pending) return;
    say(e.type === 'unauthorized' ? 'p3n doesn\'t accept this key, so it wasn\'t saved. Copy it again from the API keys page.' : e.message, 'error');
  }
}

let debounce;
els.key.addEventListener('input', () => {
  clearTimeout(debounce);
  const k = els.key.value.trim();
  if (!k) return say('');
  if (!/^p3n_\S+$/.test(k)) return say('Keys start with p3n_. Copy the whole key from the API keys page.', 'warn');
  debounce = setTimeout(() => checkAndSave(k), 400);
});

els.remove.addEventListener('click', () => {
  apiKey.remove();
  state.key = '';
  els.key.value = '';
  say('Key removed. p3n still works without one, with the free hourly limit.');
});

// ---- boot --------------------------------------------------------------------------------

(async () => {
  refreshInput();
  refreshAccount();
  // Office.js comes from Microsoft's CDN; without it (offline, or the page opened in a browser) the box still works.
  const info = globalThis.Office ? await Office.onReady() : null;
  if (info) followOfficeTheme();
  const host = HOSTS[info?.host];
  if (host) setClient(info.host);
  if (!host || !Office.context.requirements?.isSetSupported(...host.api)) {
    hint(host ? host.hints.old : 'Open p3n from Word or Excel to use your selection. You can paste text here meanwhile.');
    return;
  }
  state.host = host;
  els.text.placeholder = host.placeholder;
  refreshGrab();
  await watchSelection();
  // Opened with text selected (the ribbon button or the right-click menu): take it in, send nothing yet.
  if (state.selected?.words && !els.text.value.trim()) await grab({ whole: false });
})();
