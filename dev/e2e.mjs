// End-to-end check of the task pane in headless Chrome, with dev/office-stub.js standing in for Office.js
// and a Word document or an Excel workbook. By default the API answers with canned runs (no calls to p3n).
//   npm run e2e                 canned answers, every scenario
//   LIVE=1 npm run e2e          the live API without a key: one Check and one Rewrite (2 of the 10 free calls an hour)
//   ONLY=replace,moved          some scenarios (names below); CHROME=path/to/chrome; FRAMES=1 saves motion frames
// The run in real Word or Excel (sideloading, Ctrl+Z) is dev/live.mjs; see CONTRIBUTING.md.
import puppeteer from 'puppeteer-core';
import http from 'node:http';
import path from 'node:path';
import { existsSync } from 'node:fs';
import { mkdir, readFile } from 'node:fs/promises';
import { devManifest, PUBLISHED } from './manifest-dev.mjs';

const ROOT = path.resolve(import.meta.dirname, '..', 'src');
const OUT = path.join(import.meta.dirname, 'out');
const CHROME = process.env.CHROME || { win32: 'C:/Program Files/Google/Chrome/Application/chrome.exe', darwin: '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome' }[process.platform] || '/usr/bin/google-chrome';
const LIVE = !!process.env.LIVE;
const STUB = await readFile(path.join(import.meta.dirname, 'office-stub.js'), 'utf8');
const PKG = JSON.parse(await readFile(path.join(ROOT, '..', 'package.json'), 'utf8'));

const TITLE = 'Quarterly update';
const P1 = "In today's fast-paced world, it's worth noting that effective communication is more important than ever. Moreover, leveraging the right tools can help teams unlock their full potential.";
const P2 = 'We moved the Ljubljana office to Tivolska 50 in March, and the new lease runs to 2029.';
const P3 = 'Ultimately, this seamless transition unlocks a rich tapestry of possibilities for teams of all sizes.';
const DOC = [{ text: TITLE, style: 'Heading 1' }, P1, P2, { text: P3, style: 'List Paragraph' }];
// The workbook: notes in column A, numbers in B, a formula that makes text in B4, on a sheet whose name
// needs quoting in an address.
const SHEET = 'Q3 plan';
const P7 = 'In summary, the numbers speak for themselves.';
const BOOK = { [SHEET]: { A1: 'Notes', B1: 'Hours', A2: P1, B2: 12, A3: P2, B3: 30, A4: P3, B4: { v: 'Owner: Maja', f: '="Owner: "&"Maja"' }, A7: P7 } };

// ---- canned API answers ------------------------------------------------------------------------
const mock = { calls: [], auth: [], pages: [], clients: [], humanizeBlocks: null, humanizeScores: null, detectOutput: null, status: null, long: false };
// P7's rewrite starts with a minus sign, which Excel would read as the start of a formula.
const REWRITES = { [P1]: 'Good tools help a team talk. Which ones? That depends on the team.', [P3]: 'It works for teams of any size.', 'seamless transition': 'move', [P7]: '-30% hours, as planned.' };
const rewriteOf = (p) => REWRITES[p] ?? p.split(' ').slice(0, 8).join(' ') + '.';
function answer(url, body) {
  const free = { 'X-Free-Remaining': '7' };
  if (mock.status) return { status: mock.status.code, headers: free, body: { error: mock.status.error } };
  if (url.endsWith('/v1/me')) {
    return body?.auth === 'Bearer p3n_good_key'
      ? { status: 200, body: { email: 'dev@example.com', balance_usd: '4.20', key: { name: 'word', prefix: 'p3n_good' } } }
      : { status: 401, body: { error: { type: 'unauthorized', message: 'Invalid API key.' } } };
  }
  const text = body?.text ?? '';
  if (url.endsWith('/v1/detect')) {
    const a = { status: 200, headers: free, body: { id: 'mock-detect', skill: 'detect', status: 'succeeded', cached: false, error: null, charged_usd: '0',
      output: { version: 'detect-v1.2', ai: 0.91, label: 'ai', verdict: 'Generic claims with no fact behind them.', confidence: { level: 'medium', reason: 'Short text.' },
        sentences: [{ text: text.split(/(?<=\.)\s/)[0], why: 'A stock opening with a vague claim.' }], examples: [], ...mock.detectOutput } } };
    // A long result, so a short pane has room to scroll past the score.
    if (mock.long) a.body.output.sentences = text.split(/(?<=\.)\s/).map((t) => ({ text: t, why: 'Generic, with no fact behind it. '.repeat(4) }));
    return a;
  }
  if (url.endsWith('/v1/humanize')) {
    let paras = text.split(/\n\s*\n/).map(rewriteOf);
    if (mock.humanizeBlocks) paras = [paras.slice(0, -1).join(' '), paras.at(-1)].slice(0, mock.humanizeBlocks);
    // The live API ends the rewrite with a line break; the pane must not turn it into a paragraph.
    return { status: 200, headers: free, body: { id: 'mock-humanize', skill: 'humanize', status: 'succeeded', cached: false, error: null, charged_usd: '0',
      output: { version: 'humanize-v1.1', model: 'haiku', text: paras.join('\n\n') + '\n', ai_before: 0.88, ai_after: 0.06, label_after: 'human', passes: true, invented_blocked: [], ...mock.humanizeScores } } };
  }
  return { status: 404, body: { error: { type: 'not_found', message: 'No such endpoint.' } } };
}

// ---- results -----------------------------------------------------------------------------------
const results = [];
const ok = (name, pass, detail = '') => { results.push({ name, pass, detail }); console.log(`${pass ? 'PASS' : 'FAIL'}  ${name}${detail ? `  (${detail})` : ''}`); };

// ---- server and browser ------------------------------------------------------------------------
const TYPES = { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css', '.png': 'image/png' };
const server = http.createServer(async (req, res) => {
  const file = path.join(ROOT, new URL(req.url, 'http://x').pathname);
  try { res.writeHead(200, { 'Content-Type': TYPES[path.extname(file)] ?? 'application/octet-stream' }); res.end(await readFile(file)); } catch { res.writeHead(404).end(); }
});
await new Promise((r) => server.listen(0, '127.0.0.1', r));
const base = `http://127.0.0.1:${server.address().port}`;

await mkdir(OUT, { recursive: true });
// CI runners have no user namespaces for Chrome's sandbox.
const browser = await puppeteer.launch({ executablePath: CHROME, headless: true, pipe: true, args: ['--lang=en-US', ...(process.env.CI ? ['--no-sandbox'] : [])] });
const errors = [];

async function openPane({ name, width = 320, height = 900, calm = false, key = null, host = 'Word', scheme = 'light', theme = null, pick = null }) {
  const page = await browser.newPage();
  page.on('pageerror', (e) => errors.push(`${name}: ${e.message}`));
  page.on('console', (m) => {
    // Expected: the settings check sends a wrong key on purpose (401), and the limits check answers 429.
    if (m.type() === 'error' && !(name === 'settings' && /401/.test(m.text())) && !(name === 'limits' && /429/.test(m.text()))) errors.push(`${name}: ${m.text()}`);
  });
  await page.setViewport({ width, height });
  // The system's light or dark is set either way, so the captures don't follow the machine's own.
  await page.emulateMediaFeatures([{ name: 'prefers-color-scheme', value: scheme }, ...(calm ? [{ name: 'prefers-reduced-motion', value: 'reduce' }] : [])]);
  await page.setRequestInterception(true);
  page.on('request', (r) => {
    const url = r.url();
    if (url === 'https://appsforoffice.microsoft.com/lib/1/hosted/office.js') return r.respond({ status: 200, contentType: 'text/javascript', body: STUB });
    if (url.startsWith('https://api.p3n.si/') && !LIVE) {
      if (r.method() === 'OPTIONS') return r.respond({ status: 204, headers: { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Headers': 'Authorization, Content-Type, Accept, X-P3N-Client', 'Access-Control-Allow-Methods': 'GET, POST' } });
      const body = r.postData() ? JSON.parse(r.postData()) : {};
      body.auth = r.headers().authorization;
      mock.calls.push(new URL(url).pathname);
      mock.auth.push(body.auth ?? null);
      mock.pages.push(name);
      mock.clients.push(r.headers()['x-p3n-client'] ?? null);
      const a = answer(url, body);
      return r.respond({ status: a.status, contentType: 'application/json', headers: { 'Access-Control-Allow-Origin': '*', 'Access-Control-Expose-Headers': 'X-Free-Remaining', ...(a.headers ?? {}) }, body: JSON.stringify(a.body) });
    }
    if (url.startsWith(base) || url.startsWith('https://api.p3n.si/')) return r.continue();
    errors.push(`${name}: unexpected request to ${url}`);
    return r.abort();
  });
  // The document or workbook and the user's selection exist before the pane opens.
  await page.evaluateOnNewDocument((app, d, k, t, p) => {
    window.__host = app;
    window.__setup = d;
    if (t) window.__theme = t;
    // Pages share the test origin's storage: each starts with the key and the light or dark it asks for, or none.
    if (k) localStorage.setItem('p3n.apiKey', k);
    else localStorage.removeItem('p3n.apiKey');
    if (p) localStorage.setItem('p3n.theme', p);
    else localStorage.removeItem('p3n.theme');
    // The theme when <body> appears, before the pane's script has run.
    new MutationObserver((_, o) => { if (document.body) { window.__firstTheme = document.documentElement.dataset.theme; o.disconnect(); } }).observe(document, { childList: true, subtree: true });
  }, host, host === 'Excel' ? BOOK : DOC, key, theme, pick);
  return page;
}

async function boot(page, select) {
  if (select) await page.evaluateOnNewDocument((s) => { window.__select = s; }, select);
  await page.goto(`${base}/taskpane.html`);
  await page.waitForFunction(() => !document.querySelector('#grab').hidden, { timeout: 5000 });
}

const settled = () => document.querySelector('#result')?.textContent.trim().length > 0 && document.querySelector('#busy').hidden;
const scoreSettled = () => {
  const n = document.querySelector('.score .num');
  const sr = document.querySelector('.score .sr');
  return !!n && n.textContent === sr.textContent && !document.querySelector('.meter i')?.style.transform;
};
// Office's dark themes as the pane gets them. Microsoft doesn't publish the pane colours, and the pane takes whatever
// Office reports, so these are a dark ground and a lighter one where the palette's muted tone would miss AA unadjusted.
const officeDark = (ground) => ({ themeId: 'Black', isDarkTheme: true, bodyBackgroundColor: ground, bodyForegroundColor: '#ffffff', controlBackgroundColor: ground, controlForegroundColor: '#ffffff' });
// Every visible text in the pane against the background it's drawn on: WCAG AA (3:1 for large text). After the
// colour transitions a theme change starts.
const lowContrast = async () => {
  await Promise.all(document.getAnimations().filter((a) => a instanceof CSSTransition).map((a) => a.finished.catch(() => {})));
  const parse = (s) => s.match(/[\d.]+/g).map(Number);
  const lum = (c) => c.slice(0, 3).map((v) => { v /= 255; return v <= 0.04045 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4; }).reduce((s, v, i) => s + v * [0.2126, 0.7152, 0.0722][i], 0);
  const bad = [];
  for (const el of document.querySelectorAll('body *')) {
    if (!el.checkVisibility() || el.matches(':disabled') || ![...el.childNodes].some((n) => n.nodeType === 3 && n.textContent.trim())) continue;
    let bg = null;
    for (let a = el; a && !bg; a = a.parentElement) { const c = parse(getComputedStyle(a).backgroundColor); if ((c[3] ?? 1) > 0) bg = c; }
    const cs = getComputedStyle(el);
    const [x, y] = [lum(parse(cs.color)), lum(bg)].sort((m, n) => n - m);
    const size = parseFloat(cs.fontSize);
    const need = size >= 24 || (size >= 18.66 && Number(cs.fontWeight) >= 700) ? 3 : 4.5;
    if ((x + 0.05) / (y + 0.05) < need) bad.push(`${el.tagName.toLowerCase()}${el.className ? `.${el.className}` : ''} ${((x + 0.05) / (y + 0.05)).toFixed(2)}`);
  }
  return bad;
};
const paneTheme = (page) => page.evaluate(() => ({ theme: document.documentElement.dataset.theme, bg: getComputedStyle(document.body).backgroundColor,
  logo: [...document.querySelectorAll('.brand img')].find((i) => i.checkVisibility())?.getAttribute('src') }));
// The X-P3N-Client header of each canned call one pane made.
const clientsOf = (name) => mock.clients.filter((_, i) => mock.pages[i] === name);
const click = (page, sel) => page.evaluate((s) => document.querySelector(s).click(), sel);
const clickText = (page, text) => page.evaluate((t) => [...document.querySelectorAll('button, a')].find((b) => b.textContent.startsWith(t)).click(), text);
const hasButton = (page, text) => page.evaluate((t) => [...document.querySelectorAll('#result button')].some((b) => b.textContent.startsWith(t)), text);
const note = (page) => page.evaluate(() => document.querySelector('#result [data-status]')?.textContent ?? '');
const docParas = (page) => page.evaluate(() => window.__doc.paragraphs());
const cell = (page, a) => page.evaluate((x) => window.__book.cell(x), a);
const docStyles = (page) => page.evaluate(() => window.__doc.styles());
// Use selection, then wait for the text to land in the box.
async function grab(page) {
  await page.$eval('#text', (t) => { t.value = ''; });
  await click(page, '#grab');
  await page.waitForFunction(() => document.querySelector('#text').value || !document.querySelector('#hint').hidden, { timeout: 5000 });
}
async function run(page, button) {
  await click(page, button);
  await page.waitForFunction(settled, { timeout: 120000, polling: 100 });
}
async function replaceNow(page) {
  await page.evaluate(() => { const s = document.querySelector('#result [data-status]'); if (s) s.hidden = true; });
  await clickText(page, 'Replace in');
  await page.waitForFunction(() => { const s = document.querySelector('#result [data-status]'); return s && !s.hidden; }, { timeout: 10000 });
}

// ---- scenarios ---------------------------------------------------------------------------------
const SCENARIOS = {
  // Opened with a paragraph selected: the text comes in, nothing is sent until Check.
  async check() {
    const page = await openPane({ name: 'check' });
    const calls = mock.calls.length;
    await boot(page, [P1]);
    await page.waitForFunction(() => document.querySelector('#text').value, { timeout: 5000 }).catch(() => {});
    const box = await page.$eval('#text', (t) => t.value);
    ok('check: the selection is in the box when the pane opens', box === P1, JSON.stringify(box.slice(0, 40)));
    ok('check: nothing is sent before a click', !mock.pages.slice(calls).includes('check'), mock.calls.slice(calls).join(','));
    const grab = await page.$eval('#grab', (b) => b.textContent);
    ok('check: the link counts the selected words', /Use selection \(\d+ words\)/.test(grab), grab);
    const t0 = Date.now();
    await run(page, '#check');
    const res = await page.evaluate(() => ({ text: document.querySelector('#result').innerText, error: !!document.querySelector('#result .note.error'), quota: document.querySelector('#quota').textContent }));
    ok('check: the result shows', !res.error && /likely AI/.test(res.text), res.error ? res.text.slice(0, 160) : `${((Date.now() - t0) / 1000).toFixed(1)} s, ${res.quota}`);
    if (!LIVE) ok('check: the call names the client (X-P3N-Client)', clientsOf('check').join() === `p3n-word/${PKG.version}`, clientsOf('check').join(', '));
    if (process.env.FRAMES) for (const [i, ms] of [0, 150, 300, 500, 800].entries()) { await new Promise((r) => setTimeout(r, ms ? 150 : 0)); await page.screenshot({ path: path.join(OUT, `frame-check-${i}.png`) }); }
    await new Promise((r) => setTimeout(r, 1300));
    ok('check: the score ends on the stated value', await page.evaluate(scoreSettled));
    await page.screenshot({ path: path.join(OUT, 'pane-check.png'), fullPage: true });
    await page.close();
  },

  // Rewrite one paragraph and put it back: the rest of the document, the styles and the paragraph count stay.
  async replace() {
    const page = await openPane({ name: 'replace' });
    await boot(page, [P3]);
    await run(page, '#rewrite');
    const rewrite = await page.$eval('#result .rewrite', (r) => r.textContent).catch(() => null);
    ok('replace: the rewrite shows, without the trailing line break', !!rewrite && rewrite === rewrite.trim(), JSON.stringify(rewrite?.slice(-20)));
    ok('replace: Replace is offered', await hasButton(page, 'Replace in document'));
    ok('replace: "No longer reads as AI" when the text sent scored as AI', (await page.$eval('.score .label', (n) => n.textContent)) === 'No longer reads as AI');
    await new Promise((r) => setTimeout(r, 1300));
    await page.screenshot({ path: path.join(OUT, 'pane-rewrite.png'), fullPage: true });
    await replaceNow(page);
    const paras = await docParas(page);
    ok('replace: the rewrite is in the document', paras[3] === rewrite, JSON.stringify(paras[3]));
    ok('replace: the other paragraphs are untouched', paras[0] === TITLE && paras[1] === P1 && paras[2] === P2);
    ok('replace: no paragraph added or lost', paras.length === DOC.length, `${paras.length} paragraphs`);
    ok('replace: the paragraph keeps its style', (await docStyles(page))[3] === 'List Paragraph');
    ok('replace: the new text is selected in the document', (await page.evaluate(() => window.__doc.selected())) === rewrite);
    const msg = await note(page);
    ok('replace: says how to undo', /Ctrl\+Z/.test(msg), msg);
    await new Promise((r) => setTimeout(r, 500));
    await page.screenshot({ path: path.join(OUT, 'pane-replaced.png'), fullPage: true });

    // Another rewrite of the same text replaces the first one in turn.
    if (!LIVE) {
      await clickText(page, 'Try another rewrite');
      await page.waitForFunction(settled, { timeout: 10000 });
      ok('replace: a second rewrite can replace the first', await hasButton(page, 'Replace in document'));
      await replaceNow(page);
      ok('replace: still one paragraph there', (await docParas(page)).length === DOC.length);
    }
    await page.close();
  },

  // The headline says "No longer" or "Still" only when the detector model put the text sent in the AI class
  // (ai_before >= 0.5). It can score text at 0% that Check called likely AI.
  async headline() {
    if (LIVE) return;
    const page = await openPane({ name: 'headline' });
    await boot(page);
    const cases = [
      [{ ai_before: 0.88, ai_after: 0.06, passes: true }, 'No longer reads as AI'],
      [{ ai_before: 0.02, ai_after: 0.03, passes: true }, "Doesn't read as AI"],
      [{ ai_before: 0.88, ai_after: 0.71, passes: false, label_after: 'ai' }, 'Still reads as AI'],
      [{ ai_before: 0.3, ai_after: 0.62, passes: false, label_after: 'ai' }, 'Reads as AI'],
      [{ ai_before: null, ai_after: 0.04, passes: true }, "Doesn't read as AI"],
    ];
    for (const [i, [scores, want]] of cases.entries()) {
      mock.humanizeScores = scores;
      // A different text each time, so no answer comes from the pane's cache.
      await page.evaluate((t) => { const x = document.querySelector('#text'); x.value = t; x.dispatchEvent(new Event('input')); }, `${P2} Case ${i + 1}.`);
      await run(page, '#rewrite');
      const label = await page.$eval('.score .label', (n) => n.textContent);
      ok(`headline: "${want}" for ai_before ${scores.ai_before}, passes ${scores.passes}`, label === want, label);
    }
    mock.humanizeScores = null;
    await page.close();
  },

  // The text was edited in the document after it was taken: Replace refuses and leaves it alone.
  async changed() {
    const page = await openPane({ name: 'changed' });
    await boot(page, [P1]);
    await run(page, '#rewrite');
    await page.evaluate(() => window.__doc.type('effective communication', 'clear writing'));
    const before = await docParas(page);
    await replaceNow(page);
    const msg = await note(page);
    ok('changed: Replace refuses when the text changed', /changed in the document/.test(msg), msg);
    ok('changed: the document is left as the user edited it', JSON.stringify(await docParas(page)) === JSON.stringify(before));
    await page.close();
  },

  // An answer missing the fields the pane shows: an error the user can act on, not a script error or a 0% score.
  async malformed() {
    if (LIVE) return;
    const page = await openPane({ name: 'malformed' });
    await boot(page);
    const type = (t) => page.evaluate((v) => { const x = document.querySelector('#text'); x.value = v; x.dispatchEvent(new Event('input')); }, t);
    mock.detectOutput = { ai: undefined };
    await type(`${P1} Malformed check.`);
    await run(page, '#check');
    const d = await page.$eval('#result', (r) => ({ text: r.innerText, error: !!r.querySelector('.note.error') }));
    ok('malformed: a check with no score says so instead of showing 0%', d.error && !/likely AI/.test(d.text), d.text.slice(0, 120));
    mock.detectOutput = null;
    mock.humanizeScores = { text: undefined };
    await type(`${P1} Malformed rewrite.`);
    await run(page, '#rewrite');
    const h = await page.$eval('#result', (r) => ({ text: r.innerText, error: !!r.querySelector('.note.error') }));
    ok('malformed: a rewrite with no text says so, with no script error', h.error && !/Cannot read|undefined/.test(h.text), h.text.slice(0, 120));
    mock.humanizeScores = null;
    await page.close();
  },

  // A rewrite still on screen after other text was taken in: its Replace refuses, and the new text stays.
  async stale() {
    const page = await openPane({ name: 'stale' });
    await boot(page, [P1]);
    await run(page, '#rewrite');
    await page.evaluate((p) => window.__doc.select(p), P2);
    await grab(page);
    await replaceNow(page);
    const msg = await note(page);
    ok('stale: Replace refuses after other text was taken in', /other text/.test(msg), msg);
    ok('stale: the document is left alone', JSON.stringify(await docParas(page)) === JSON.stringify([TITLE, P1, P2, P3]), JSON.stringify(await docParas(page)));
    await page.close();
  },

  // Edits elsewhere move the text: the tracked range follows it, and Replace still hits the right words.
  async moved() {
    const page = await openPane({ name: 'moved' });
    await boot(page, [P3]);
    await run(page, '#rewrite');
    await page.evaluate(() => window.__doc.type('Quarterly update', 'Quarterly update for the board, October'));
    await replaceNow(page);
    const paras = await docParas(page);
    ok('moved: Replace lands on the moved text', paras[3] === 'It works for teams of any size.' && paras[0].endsWith('October'), JSON.stringify(paras));
    await page.close();
  },

  // Three paragraphs with three styles; the rewrite has three paragraphs: each keeps its style.
  async paragraphs() {
    const page = await openPane({ name: 'paragraphs' });
    await boot(page);
    await page.evaluate((t) => window.__doc.select(t), [P1, P2, P3].join('\r'));
    await page.waitForFunction(() => /Use selection \(\d+ words\)/.test(document.querySelector('#grab').textContent));
    await grab(page);
    const box = await page.$eval('#text', (t) => t.value);
    ok('paragraphs: Word paragraphs arrive as paragraphs (blank line between)', box === [P1, P2, P3].join('\n\n'), JSON.stringify(box.slice(0, 80)));
    await run(page, '#rewrite');
    await replaceNow(page);
    const paras = await docParas(page);
    const styles = await docStyles(page);
    ok('paragraphs: each paragraph is replaced in place', paras.length === 4 && paras[2] === rewriteOf(P2) && paras[3] === rewriteOf(P3), JSON.stringify(paras));
    ok('paragraphs: each keeps its own style', JSON.stringify(styles) === JSON.stringify(['Heading 1', 'Normal', 'Normal', 'List Paragraph']), JSON.stringify(styles));
    await page.close();
  },

  // The rewrite merged two paragraphs: one insert, still no empty paragraph at the end.
  async merged() {
    const page = await openPane({ name: 'merged' });
    mock.humanizeBlocks = 2;
    await boot(page);
    // Down to and including the last paragraph mark, as a triple-click or Shift+Down selects.
    await page.evaluate((t) => window.__doc.select(t), [P1, P2, P3].join('\r') + '\r');
    await page.waitForFunction(() => /Use selection/.test(document.querySelector('#grab').textContent));
    await grab(page);
    await run(page, '#rewrite');
    await replaceNow(page);
    mock.humanizeBlocks = null;
    const paras = await docParas(page);
    ok('merged: the rewrite replaces the three paragraphs with two', paras.length === 3 && paras[0] === TITLE, JSON.stringify(paras));
    ok('merged: no empty paragraph left behind', paras.every((p) => p.trim()), JSON.stringify(paras));
    await page.close();
  },

  // Part of a sentence, with the space after it (a double-click selection): the space stays.
  async partial() {
    const page = await openPane({ name: 'partial' });
    await boot(page, ['seamless transition ']);
    await run(page, '#rewrite');
    await replaceNow(page);
    const p = (await docParas(page))[3];
    ok('partial: only the selected words change, the space after them stays', p === P3.replace('seamless transition', 'move'), JSON.stringify(p));
    await page.close();
  },

  // Nothing selected: the link offers the whole document.
  async whole() {
    const page = await openPane({ name: 'whole' });
    await boot(page);
    const label = await page.$eval('#grab', (b) => b.textContent);
    ok('whole: with nothing selected the link offers the whole document', label === 'Use whole document', label);
    await grab(page);
    const box = await page.$eval('#text', (t) => t.value);
    ok('whole: the whole document comes in', box === [TITLE, P1, P2, P3].join('\n\n'));
    ok('whole: says so', /whole document/.test(await page.$eval('#hint', (h) => h.textContent)));
    await page.close();
  },

  // Text edited in the pane is no longer the document's text: no Replace for it.
  async edited() {
    const page = await openPane({ name: 'edited' });
    await boot(page, [P1]);
    await page.evaluate(() => { const t = document.querySelector('#text'); t.value += ' We ship it in two weeks.'; t.dispatchEvent(new Event('input')); });
    await run(page, '#rewrite');
    ok('edited: no Replace for text edited in the pane', !(await hasButton(page, 'Replace in document')) && (await hasButton(page, 'Copy')));
    await page.close();
  },

  // Tables and pictures aren't pasted over.
  async table() {
    const page = await openPane({ name: 'table' });
    await page.evaluateOnNewDocument(() => { window.__setup = [...window.__setup.slice(0, 2), { text: 'Q3\tQ4', table: true }, ...window.__setup.slice(2)]; });
    await boot(page);
    await page.evaluate((t) => window.__doc.select(t), `${P1}\rQ3\tQ4`);
    await page.waitForFunction(() => /Use selection/.test(document.querySelector('#grab').textContent));
    await grab(page);
    await run(page, '#rewrite');
    const before = await docParas(page);
    await replaceNow(page);
    ok('table: Replace refuses over a table', /table/.test(await note(page)) && JSON.stringify(await docParas(page)) === JSON.stringify(before));
    await page.close();
  },

  // Reduced motion: the final state shows at once. In a short pane (Word's, at 125% scaling) the result
  // scrolls into view.
  async calm() {
    const page = await openPane({ name: 'calm', calm: true, height: 560 });
    await boot(page, [P1]);
    await run(page, '#check');
    ok('calm: reduced motion shows the final score at once', await page.evaluate(scoreSettled));
    ok('calm: a result below the fold scrolls into view', await page.evaluate(() => { const r = document.querySelector('.score').getBoundingClientRect(); return scrollY > 0 && r.top >= 0 && r.bottom <= innerHeight; }));
    await page.close();
  },

  // With motion, in a short pane: the result scrolls up smoothly and the score ends in view, whole. (Word's
  // and Excel's WebView doesn't hold the scroll position when the progress bar above it goes away.)
  async reveal() {
    const page = await openPane({ name: 'reveal', height: 560 });
    await boot(page, [P1]);
    await page.addStyleTag({ content: 'html, body { overflow-anchor: none; }' });
    mock.long = true;
    await run(page, '#check');
    mock.long = false;
    await new Promise((r) => setTimeout(r, 1500));
    const at = await page.evaluate(() => { const r = document.querySelector('.score').getBoundingClientRect(); return { top: Math.round(r.top), bottom: Math.round(r.bottom), scrolled: Math.round(scrollY) }; });
    ok('reveal: the score scrolls into view, not past it', at.scrolled > 0 && at.top >= 0 && at.bottom <= 560, JSON.stringify(at));
    await page.close();
  },

  // Errors and limits read as the extension's do.
  async limits() {
    const page = await openPane({ name: 'limits' });
    await boot(page);
    const calls = mock.calls.length;
    await page.evaluate(() => { const t = document.querySelector('#text'); t.value = 'word '.repeat(1200); t.dispatchEvent(new Event('input')); });
    await click(page, '#check');
    const longMsg = await page.$eval('#result', (r) => r.innerText);
    ok('limits: over 5,000 characters without a key is explained, no call', /5,000/.test(longMsg) && mock.calls.length === calls, longMsg.slice(0, 100));
    if (!LIVE) {
      mock.status = { code: 429, error: { type: 'free_limit', message: 'Free limit reached.', signup_url: 'https://p3n.si/register', retry_after_seconds: 1500 } };
      await page.evaluate(() => { const t = document.querySelector('#text'); t.value = 'A short text to check.'; t.dispatchEvent(new Event('input')); });
      await run(page, '#check');
      mock.status = null;
      const msg = await page.$eval('#result', (r) => r.innerText);
      ok('limits: the free limit says when more come and offers sign-up and a key', /used up; more in about 25 min/.test(msg) && /Sign up/.test(msg) && /Add an API key/.test(msg), msg.replace(/\n/g, ' ').slice(0, 120));
      await new Promise((r) => setTimeout(r, 600)); // the message finishes arriving
      await page.screenshot({ path: path.join(OUT, 'pane-free-limit.png'), fullPage: true });
    }
    await page.close();
  },

  // The API key: a wrong one is refused and not saved; a good one is checked, saved and used.
  async settings() {
    if (LIVE) return;
    const page = await openPane({ name: 'settings', height: 700 });
    await boot(page);
    await click(page, '#account');
    await page.waitForFunction(() => !document.querySelector('#settings').hidden);
    ok('settings: focus goes to the key field', await page.evaluate(() => document.activeElement?.id === 'key'));
    await page.type('#key', 'p3n_wrong_key');
    await page.waitForFunction(() => /accept/.test(document.querySelector('#key-status').textContent), { timeout: 5000 }).catch(() => {});
    ok('settings: a wrong key is refused and not saved', /doesn't accept/.test(await page.$eval('#key-status', (s) => s.textContent)) && (await page.evaluate(() => localStorage.getItem('p3n.apiKey'))) === null);
    await page.$eval('#key', (k) => { k.value = ''; });
    await page.type('#key', 'p3n_good_key');
    await page.waitForFunction(() => /Saved/.test(document.querySelector('#key-status').textContent), { timeout: 5000 }).catch(() => {});
    ok('settings: a good key is saved after p3n accepts it', (await page.evaluate(() => localStorage.getItem('p3n.apiKey'))) === 'p3n_good_key', await page.$eval('#key-status', (s) => s.textContent));
    await new Promise((r) => setTimeout(r, 300));
    await page.screenshot({ path: path.join(OUT, 'pane-settings.png'), fullPage: true });
    await page.keyboard.press('Escape');
    await page.waitForFunction(() => document.querySelector('#settings').hidden && /Balance \$4\.20/.test(document.querySelector('#quota').textContent), { timeout: 5000 }).catch(() => {});
    ok('settings: Escape goes back, the footer shows the balance', await page.evaluate(() => document.querySelector('#settings').hidden && /Balance/.test(document.querySelector('#quota').textContent)));
    ok('settings: focus returns to the API key link', await page.evaluate(() => document.activeElement?.id === 'account'));
    // The key goes out as a bearer token.
    const n = mock.calls.length;
    await page.evaluate(() => { const t = document.querySelector('#text'); t.value = 'A short text to check.'; t.dispatchEvent(new Event('input')); });
    await run(page, '#check');
    ok('settings: checks send the key as a bearer token', mock.calls[n] === '/v1/detect' && mock.auth[n] === 'Bearer p3n_good_key', `${mock.calls[n]} ${mock.auth[n]}`);
    await page.close();
  },

  // ---- Excel ----------------------------------------------------------------------------------

  // Several cells: each with text comes in as a paragraph, numbers stay out. A rewrite of them can be
  // copied but not put back, and the pane says why. A whole column is read only where the sheet has data.
  async excelCheck() {
    const page = await openPane({ name: 'excelCheck', host: 'Excel' });
    const calls = mock.calls.length;
    await boot(page, ['A1:B3']);
    await page.waitForFunction(() => document.querySelector('#text').value, { timeout: 5000 }).catch(() => {});
    const box = await page.$eval('#text', (t) => t.value);
    ok('excel: the selected cells are in the box when the pane opens', box === ['Notes', 'Hours', P1, P2].join('\n\n'), JSON.stringify(box.slice(0, 50)));
    ok('excel: nothing is sent before a click', !mock.pages.slice(calls).includes('excelCheck'), mock.calls.slice(calls).join(','));
    ok('excel: the link counts the words in the cells', /Use selection \(\d+ words\)/.test(await page.$eval('#grab', (b) => b.textContent)));
    ok('excel: the box suggests selecting cells', /Select cells/.test(await page.$eval('#text', (t) => t.placeholder)));
    await run(page, '#check');
    ok('excel: Check shows a result', /likely AI/.test(await page.$eval('#result', (r) => r.innerText)));
    ok('excel: the call names the client (X-P3N-Client)', clientsOf('excelCheck').join() === `p3n-excel/${PKG.version}`, clientsOf('excelCheck').join(', '));
    await run(page, '#rewrite');
    const kept = await page.$eval('#result', (r) => r.innerText);
    ok('excel: a rewrite of several cells has Copy but no Replace, and says why', !(await hasButton(page, 'Replace')) && (await hasButton(page, 'Copy')) && /single cell/.test(kept), kept.split('\n').slice(-3).join(' '));
    await page.evaluate(() => window.__book.select('A1:A1048576'));
    const colWords = ['Notes', P1, P2, P3, P7].join(' ').split(/\s+/).length;
    await page.waitForFunction((n) => document.querySelector('#grab').textContent === `Use selection (${n} words)`, { timeout: 5000 }, colWords).catch(() => {});
    const col = await page.$eval('#grab', (b) => b.textContent);
    ok('excel: a whole column is read only where it has text', col === `Use selection (${colWords} words)`, col);
    await page.close();
  },

  // One cell: Replace writes the rewrite into it, leaves the rest of the sheet alone and selects it.
  async excelReplace() {
    const page = await openPane({ name: 'excelReplace', host: 'Excel' });
    await boot(page, ['A4']);
    await page.waitForFunction(() => document.querySelector('#text').value, { timeout: 5000 });
    await run(page, '#rewrite');
    ok('excel: Replace names the cell', await hasButton(page, 'Replace in A4'));
    await new Promise((r) => setTimeout(r, 1300));
    await page.screenshot({ path: path.join(OUT, 'pane-excel-rewrite.png'), fullPage: true });
    await replaceNow(page);
    const a4 = await cell(page, 'A4');
    ok('excel: the rewrite is in the cell, as text', a4.value === rewriteOf(P3) && a4.type === 'String', JSON.stringify(a4));
    const rest = await Promise.all(['A2', 'A3', 'B3', 'B4', 'A7'].map((a) => cell(page, a)));
    ok('excel: the other cells are untouched, the formula too', rest[0].value === P1 && rest[1].value === P2 && rest[2].value === 30 && rest[3].formula === '="Owner: "&"Maja"' && rest[4].value === P7);
    ok('excel: the cell is selected', (await page.evaluate(() => window.__book.selected())) === `${SHEET}!A4`);
    const msg = await note(page);
    ok('excel: says where and how to undo', /A4/.test(msg) && /Ctrl\+Z/.test(msg), msg);
    await new Promise((r) => setTimeout(r, 500));
    await page.screenshot({ path: path.join(OUT, 'pane-excel-replaced.png'), fullPage: true });
    await clickText(page, 'Try another rewrite');
    await page.waitForFunction(settled, { timeout: 10000 });
    await replaceNow(page);
    ok('excel: a second rewrite replaces the first', /Replaced/.test(await note(page)) && (await cell(page, 'A4')).value === rewriteOf(P3));
    await page.close();
  },

  // A rewrite that starts with a minus sign stays text, not a broken formula.
  async excelText() {
    const page = await openPane({ name: 'excelText', host: 'Excel' });
    await boot(page, ['A7']);
    await page.waitForFunction(() => document.querySelector('#text').value, { timeout: 5000 });
    await run(page, '#rewrite');
    await replaceNow(page);
    const a7 = await cell(page, 'A7');
    ok('excel: a rewrite starting with "-" goes in as text', a7.value === REWRITES[P7] && a7.type === 'String' && !a7.formula, JSON.stringify(a7));
    await page.close();
  },

  // The cell was edited after it was taken: Replace refuses and keeps the edit.
  async excelChanged() {
    const page = await openPane({ name: 'excelChanged', host: 'Excel' });
    await boot(page, ['A2']);
    await page.waitForFunction(() => document.querySelector('#text').value, { timeout: 5000 });
    await run(page, '#rewrite');
    await page.evaluate(() => window.__book.type('A2', 'Edited by hand.'));
    await replaceNow(page);
    const msg = await note(page);
    ok('excel: Replace refuses when the cell changed', /cell changed/.test(msg) && (await cell(page, 'A2')).value === 'Edited by hand.', msg);
    await page.close();
  },

  // A formula's text can be checked and rewritten, but the formula is never written over.
  async excelFormula() {
    const page = await openPane({ name: 'excelFormula', host: 'Excel' });
    await boot(page, ['B4']);
    await page.waitForFunction(() => document.querySelector('#text').value, { timeout: 5000 });
    ok('excel: a formula\'s text comes in', (await page.$eval('#text', (t) => t.value)) === 'Owner: Maja');
    await run(page, '#rewrite');
    ok('excel: no Replace over a formula, and it says so', !(await hasButton(page, 'Replace')) && /formula/.test(await page.$eval('#result', (r) => r.innerText)));
    await page.close();
  },

  // Excel is in the middle of editing a cell: Replace waits for it, then works.
  async excelEditing() {
    const page = await openPane({ name: 'excelEditing', host: 'Excel' });
    await boot(page, ['A3']);
    await page.waitForFunction(() => document.querySelector('#text').value, { timeout: 5000 });
    await run(page, '#rewrite');
    await page.evaluate(() => window.__book.editing(true));
    await replaceNow(page);
    const msg = await note(page);
    ok('excel: while a cell is being edited, Replace says to finish it first', /busy editing a cell/.test(msg) && (await cell(page, 'A3')).value === P2, msg);
    await page.evaluate(() => window.__book.editing(false));
    await replaceNow(page);
    ok('excel: then Replace works', /Replaced/.test(await note(page)) && (await cell(page, 'A3')).value === rewriteOf(P2), await note(page));
    await page.close();
  },

  // The sheet is renamed after the text was taken: Replace still finds the cell. A protected sheet refuses.
  async excelSheet() {
    const page = await openPane({ name: 'excelSheet', host: 'Excel' });
    await boot(page, ['A3']);
    await page.waitForFunction(() => document.querySelector('#text').value, { timeout: 5000 });
    await run(page, '#rewrite');
    await page.evaluate((n) => { window.__book.rename(n, `${n} (final)`); window.__book.lock(); }, SHEET);
    await replaceNow(page);
    ok('excel: a protected sheet is left alone, and it says so', /protected/.test(await note(page)) && (await cell(page, 'A3')).value === P2, await note(page));
    await page.evaluate(() => window.__book.lock(false));
    await replaceNow(page);
    ok('excel: Replace finds the cell after the sheet was renamed', /Replaced/.test(await note(page)) && (await cell(page, 'A3')).value === rewriteOf(P2), await note(page));
    await page.close();
  },

  // An empty cell, and a selection too big to read.
  async excelEmpty() {
    const page = await openPane({ name: 'excelEmpty', host: 'Excel' });
    await page.evaluateOnNewDocument((n) => { window.__setup[n].Z1000 = 'far away'; }, SHEET);
    await boot(page, ['D10']);
    ok('excel: an empty cell offers Use selection, and Check waits for text', (await page.$eval('#grab', (b) => b.textContent)) === 'Use selection' && (await page.$eval('#check', (b) => b.disabled)));
    await grab(page);
    ok('excel: Use selection on an empty cell says to select text', /no text/.test(await page.$eval('#hint', (h) => h.textContent)));
    await page.evaluate(() => window.__book.select('A1:Z1000'));
    await new Promise((r) => setTimeout(r, 400));
    await grab(page);
    ok('excel: a selection too big to read says so', /too big/.test(await page.$eval('#hint', (h) => h.textContent)), await page.$eval('#hint', (h) => h.textContent));
    await page.close();
  },

  // ---- manifests --------------------------------------------------------------------------------

  // The published manifest and its localhost twin say the same thing; every file they name exists.
  async manifests() {
    if (LIVE) return;
    const xml = await readFile(path.join(ROOT, '..', 'manifest.xml'), 'utf8');
    const dev = await readFile(path.join(ROOT, '..', 'manifest.dev.xml'), 'utf8');
    ok('manifests: manifest.dev.xml is manifest.xml served from localhost (npm run manifest)', devManifest(xml) === dev);
    ok('manifests: Word and Excel', /<Host Name="Document"\/>/.test(xml) && /<Host Name="Workbook"\/>/.test(xml) && /ContextMenuText/.test(xml) && /ContextMenuCell/.test(xml));
    const urls = [...xml.matchAll(/DefaultValue="(https:[^"]+)"/g)].map((m) => m[1]).filter((u) => u.startsWith(PUBLISHED));
    const missing = urls.filter((u) => !existsSync(path.join(ROOT, u.slice(PUBLISHED.length))));
    ok('manifests: every pane file the manifest names is in src/', urls.length >= 6 && !missing.length, missing.join(' '));
    const pkg = JSON.parse(await readFile(path.join(ROOT, '..', 'package.json'), 'utf8'));
    ok('manifests: the version matches package.json', xml.includes(`<Version>${pkg.version}.0</Version>`), pkg.version);
  },

  // The pane follows Office's theme, not the system's: a dark Office theme gets the dark palette on Office's own pane
  // colour with every text passing AA, a light one stays light on a dark system, and a change is picked up when the
  // pane comes back into view.
  async theme() {
    if (LIVE) return;
    for (const [ground, rgb] of [['#262626', 'rgb(38, 38, 38)'], ['#4a4a4a', 'rgb(74, 74, 74)']]) {
      const page = await openPane({ name: 'theme', theme: officeDark(ground) });
      await boot(page, [P1]);
      let t = await paneTheme(page);
      ok(`theme: a dark Office theme on ${ground} darkens the pane on Office's colour`, t.theme === 'dark' && t.bg === rgb && t.logo === 'assets/logo-dark.png', JSON.stringify(t));
      const before = await page.evaluate(lowContrast);
      ok(`theme: on ${ground}, every text passes AA before a run`, !before.length, before.join(', '));
      await run(page, '#check');
      await new Promise((r) => setTimeout(r, 1300));
      const checked = await page.evaluate(lowContrast);
      ok(`theme: on ${ground}, every text in a Check result passes AA`, !checked.length, checked.join(', '));
      await page.screenshot({ path: path.join(OUT, `pane-dark-${ground.slice(1)}-check.png`), fullPage: true });
      await grab(page);
      await run(page, '#rewrite');
      await replaceNow(page);
      await new Promise((r) => setTimeout(r, 500));
      const replaced = await page.evaluate(lowContrast);
      ok(`theme: on ${ground}, every text in a replaced rewrite passes AA`, !replaced.length, replaced.join(', '));
      await page.screenshot({ path: path.join(OUT, `pane-dark-${ground.slice(1)}-replaced.png`), fullPage: true });
      if (ground === '#262626') {
        // Back to Colorful while the pane was out of view.
        await page.evaluate(() => { window.__theme = null; window.dispatchEvent(new Event('focus')); });
        t = await paneTheme(page);
        ok('theme: a change to a light Office theme is picked up when the pane is back in view', t.theme === 'light' && t.bg === 'rgb(251, 250, 247)' && t.logo === 'assets/logo.png', JSON.stringify(t));
      }
      await page.close();
    }
    const page = await openPane({ name: 'theme', scheme: 'dark' });
    await boot(page, [P1]);
    const t = await paneTheme(page);
    ok('theme: a light Office theme keeps the pane light on a dark system', t.theme === 'light' && t.bg === 'rgb(251, 250, 247)', JSON.stringify(t));
    await page.close();
  },

  // The sun/moon button: the pane follows Office until a click keeps the other theme, a second click keeps the first,
  // a kept theme shows from the first paint when the pane opens again, and the settings have no theme control.
  async switch() {
    if (LIVE) return;
    const state = (page) => page.evaluate(() => ({ theme: document.documentElement.dataset.theme, bg: getComputedStyle(document.body).backgroundColor,
      kept: localStorage.getItem('p3n.theme'), label: document.querySelector('header [data-theme-toggle]')?.getAttribute('aria-label') }));
    const toggle = (page) => click(page, 'header [data-theme-toggle]');
    const still = async (page, file) => { await new Promise((r) => setTimeout(r, 300)); await page.screenshot({ path: path.join(OUT, file) }); };

    // A light Office theme.
    let page = await openPane({ name: 'switch', height: 560 });
    await boot(page, [P1]);
    let t = await state(page);
    ok('switch: the sun/moon button is in the header, following Office', t.label === 'Switch to dark theme' && t.kept === null && t.theme === 'light', JSON.stringify(t));
    await still(page, 'pane-switch-light.png');
    await toggle(page);
    t = await state(page);
    ok("switch: a click turns a light Office theme's pane dark (p3n's dark) and keeps it", t.theme === 'dark' && t.bg === 'rgb(22, 19, 15)' && t.kept === 'dark' && t.label === 'Switch to light theme', JSON.stringify(t));
    const dark = await page.evaluate(lowContrast);
    ok('switch: every text passes AA in the kept dark', !dark.length, dark.join(', '));
    await still(page, 'pane-switch-dark.png');
    await toggle(page);
    t = await state(page);
    ok('switch: a second click keeps light', t.theme === 'light' && t.kept === 'light', JSON.stringify(t));
    await page.close();

    // Dark kept, the pane opened again: dark from the first paint.
    page = await openPane({ name: 'switch', pick: 'dark' });
    await boot(page, [P1]);
    t = await state(page);
    ok('switch: a kept theme shows when the pane opens again, from the first paint', await page.evaluate(() => window.__firstTheme) === 'dark' && t.theme === 'dark', JSON.stringify(t));
    await page.close();

    // A dark Office theme: light kept, then dark on Office's own colour; no theme control in the settings.
    page = await openPane({ name: 'switch', height: 760, theme: officeDark('#262626') });
    await boot(page, [P1]);
    t = await state(page);
    ok("switch: a dark Office theme's pane starts on Office's colour", t.theme === 'dark' && t.bg === 'rgb(38, 38, 38)' && t.kept === null, JSON.stringify(t));
    await toggle(page);
    t = await state(page);
    ok("switch: a click turns a dark Office theme's pane light", t.theme === 'light' && t.bg === 'rgb(251, 250, 247)' && t.kept === 'light', JSON.stringify(t));
    await toggle(page);
    t = await state(page);
    ok("switch: a second click is dark on Office's own colour", t.theme === 'dark' && t.bg === 'rgb(38, 38, 38)' && t.kept === 'dark', JSON.stringify(t));
    await click(page, '#account');
    await page.waitForFunction(() => !document.querySelector('#settings').hidden, { timeout: 5000 });
    ok('switch: the settings have no theme control', await page.evaluate(() => !document.querySelector('[data-theme-set], [role="radiogroup"]') && !/Appearance/.test(document.querySelector('#settings').innerText)));
    const settings = await page.evaluate(lowContrast);
    ok("switch: every text in the settings passes AA on Office's dark", !settings.length, settings.join(', '));
    await still(page, 'pane-switch-settings-dark.png');
    await page.evaluate(() => localStorage.removeItem('p3n.theme'));
    await page.close();
  },

  // Outside Office (the pane opened in a plain browser): paste still works, no selection link.
  async browser() {
    if (LIVE) return;
    const page = await browser.newPage();
    page.on('pageerror', (e) => errors.push(`browser: ${e.message}`));
    await page.setRequestInterception(true);
    page.on('request', (r) => (r.url().startsWith(base) ? r.continue() : r.abort()));
    await page.goto(`${base}/taskpane.html`);
    await page.waitForFunction(() => !document.querySelector('#hint').hidden, { timeout: 5000 }).catch(() => {});
    ok('browser: without Office.js the pane explains and still takes pasted text', await page.evaluate(() => document.querySelector('#grab').hidden && /Open p3n from Word/.test(document.querySelector('#hint').textContent)));
    // Without Office's theme the system's light or dark decides, live.
    const scheme = async (value) => {
      await page.emulateMediaFeatures([{ name: 'prefers-color-scheme', value }]);
      await page.waitForFunction((v) => document.documentElement.dataset.theme === v, { timeout: 2000, polling: 50 }, value).catch(() => {});
      return paneTheme(page);
    };
    const light = await scheme('light');
    const dark = await scheme('dark');
    ok("browser: without Office the system's light or dark decides", light.theme === 'light' && dark.theme === 'dark' && dark.bg === 'rgb(22, 19, 15)' && dark.logo === 'assets/logo-dark.png', JSON.stringify([light, dark]));
    const bad = await page.evaluate(lowContrast);
    ok('browser: every text passes AA in the dark palette', !bad.length, bad.join(', '));
    await page.close();
  },
};

// LIVE runs only what needs the API, to spend two free calls.
const only = process.env.ONLY?.split(',') ?? (LIVE ? ['check', 'replace'] : Object.keys(SCENARIOS));
try {
  for (const name of only) await SCENARIOS[name]();
  ok('no console errors or unexpected requests', errors.length === 0, errors.slice(0, 5).join(' || '));
} finally {
  await Promise.race([browser.close(), new Promise((r) => setTimeout(r, 5000))]);
  server.close();
}
const failed = results.filter((r) => !r.pass);
console.log(`\n${results.length - failed.length}/${results.length} passed${LIVE ? ' (live API)' : ' (canned API answers)'}. Screenshots in dev/out/.`);
process.exit(failed.length ? 1 : 0);
