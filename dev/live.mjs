// Live run in Word or Excel for Windows: sideloads the add-in with Microsoft's office-addin-debugging, writes
// a test document or sheet through the app's COM interface, and drives the real task pane over WebView2's
// debugging port (opened on 9229 through WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS). It never takes focus from
// what you're doing: the window is placed without activating it, screenshots are of the app's own window
// (so nothing else on the screen gets in), and undo runs the app's Undo command.
//   node dev/live.mjs word           one Check and one Rewrite against the live API (2 of the 10 free calls an hour)
//   node dev/live.mjs excel          the same in Excel
//   MOCK=1 node dev/live.mjs <app>   canned API answers; adds the multi-paragraph, multi-cell and changed-text cases
//   SHOTS=1                          only the store screenshots: a window of 1366x768, one Rewrite and Replace
//                                    (and in Word a Check first)
//   RECORD=1                         save each live answer to dev/out/run-<skill>.json
//   REPLAY=1                         answer from those files instead of calling the API (the same text only)
//   MANIFEST=manifest.xml            the published add-in, its pane from GitHub Pages (default: manifest.dev.xml)
//   KEYS=1                           undo with a real Ctrl+Z keystroke; brings the app to the front first
//   KEEP=1                           leave the app and the dev server running afterwards
// Needs: the app installed and closed before the run; for manifest.dev.xml, the dev certificate
// (npx office-addin-dev-certs install). The store images are made from the raw shots by dev/store-shots.py.
import puppeteer from 'puppeteer-core';
import { spawn, spawnSync } from 'node:child_process';
import path from 'node:path';
import { mkdir, readFile, writeFile } from 'node:fs/promises';

const ROOT = path.resolve(import.meta.dirname, '..');
const OUT = path.join(import.meta.dirname, 'out');
const MOCK = !!process.env.MOCK;
const { SHOTS, RECORD, REPLAY, KEYS } = process.env;
const MANIFEST = process.env.MANIFEST ?? 'manifest.dev.xml';
const APP = (process.argv[2] ?? 'word').toLowerCase();
const results = [];
const ok = (name, pass, detail = '') => { results.push({ name, pass, detail }); console.log(`${pass ? 'PASS' : 'FAIL'}  ${name}${detail ? `  (${detail})` : ''}`); };
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// The manifest says where the pane is served from. office-addin-debugging opens Excel on a workbook named
// after the add-in's Id. Excel joins the list of running COM objects only once its window has lost focus,
// so until then it's reached through that file.
const XML = await readFile(path.join(ROOT, MANIFEST), 'utf8');
const ADDIN_ID = /<Id>([^<]+)<\/Id>/.exec(XML)[1];
const PANE = /<SourceLocation DefaultValue="([^"]+)"/.exec(XML)[1];
const LOCAL = PANE.startsWith('https://localhost:');
const BOOK = `Join-Path $env:TEMP 'Excel add-in ${ADDIN_ID}.xlsx'`;

// What differs between the two apps for COM: the process, how to reach the document, its window.
const APPS = {
  word: {
    name: 'Word', exe: 'WINWORD',
    prelude: `$app=[Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application'); $d=$app.ActiveDocument;`,
    hwnd: '$app.ActiveWindow.Hwnd', normal: '$app.WindowState = 0', front: '$app.Activate()', focus: '$app.ActiveWindow.SetFocus()',
    quit: '$d.Saved = $true; $app.Quit(0)',
  },
  excel: {
    name: 'Excel', exe: 'EXCEL',
    prelude: `try { $app=[Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') } catch { if (-not (Get-Process EXCEL -ErrorAction SilentlyContinue)) { throw }; $app=([Runtime.InteropServices.Marshal]::BindToMoniker((${BOOK}))).Application }; $b=$app.ActiveWorkbook; $s=$b.ActiveSheet;`,
    hwnd: '$app.Hwnd', normal: '$app.WindowState = -4143', front: '$b.Activate()', focus: '$app.ActiveWindow.Activate()',
    quit: '$app.DisplayAlerts = $false; $b.Saved = $true; $app.Quit()',
  },
};
const app = APPS[APP];
if (!app) { console.error('Usage: node dev/live.mjs word|excel'); process.exit(2); }

// ---- the app through COM ------------------------------------------------------------------------
// One PowerShell call per step; each prints JSON. The app is found with GetActiveObject.
function ps(script) {
  // A modal dialog in the app blocks COM; give up after a minute rather than hang.
  const r = spawnSync('powershell', ['-NoProfile', '-NonInteractive', '-Command', `$ErrorActionPreference='Stop'; ${app.prelude} ${script}`], { encoding: 'utf8', timeout: 60000 });
  if (r.error) throw new Error(`PowerShell: ${r.error.message} (is a dialog open in ${app.name}?)`);
  if (r.status !== 0) throw new Error(`PowerShell: ${r.stderr.trim().split('\n')[0]}`);
  const out = r.stdout.trim();
  return out ? JSON.parse(out) : null;
}
const q = (s) => s.replace(/'/g, "''");

// Undo runs the app's Undo command, the one Ctrl+Z is bound to. With KEYS=1 it tries a real Ctrl+Z keystroke
// instead, which needs the app in front: it asks Windows to bring it there, and sends the keys only if the app
// is then the foreground window, so keys never go to another window. Otherwise it falls back to the command.
const WIN32 = `Add-Type -Namespace P3n -Name U -MemberDefinition '[DllImport("user32.dll")] public static extern bool SetProcessDPIAware(); [DllImport("user32.dll")] public static extern System.IntPtr GetForegroundWindow(); [DllImport("user32.dll")] public static extern bool SetForegroundWindow(System.IntPtr h); [DllImport("user32.dll")] public static extern bool PrintWindow(System.IntPtr h, System.IntPtr dc, uint f); [DllImport("user32.dll")] public static extern bool GetWindowRect(System.IntPtr h, out RECT r); [DllImport("user32.dll")] public static extern bool ShowWindow(System.IntPtr h, int c); [DllImport("user32.dll")] public static extern bool SetWindowPos(System.IntPtr h, System.IntPtr a, int x, int y, int cx, int cy, uint f); [DllImport("user32.dll")] public static extern uint GetDpiForWindow(System.IntPtr h); [DllImport("dwmapi.dll")] public static extern int DwmGetWindowAttribute(System.IntPtr h, int a, out RECT r, int n); public struct RECT { public int L, T, R, B; }'; $h=[System.IntPtr]${app.hwnd};`;
function undo() {
  if (!KEYS) return ps(`$app.CommandBars.ExecuteMso('Undo'); Start-Sleep -Milliseconds 400; '"${app.name} Undo command"'`);
  return ps(`${WIN32} ${app.front}; [void][P3n.U]::SetForegroundWindow($h); Start-Sleep -Milliseconds 400;
    if ([P3n.U]::GetForegroundWindow() -eq $h) {
      ${app.focus}; Start-Sleep -Milliseconds 200;
      if ([P3n.U]::GetForegroundWindow() -eq $h) { Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.SendKeys]::SendWait('^z'); Start-Sleep -Milliseconds 600; '"Ctrl+Z keystroke"'; return }
    }
    $app.CommandBars.ExecuteMso('Undo'); Start-Sleep -Milliseconds 400; '"${app.name} Undo command (${app.name} was not in front, so no keystroke)"'`);
}

// The app at a fixed size, placed without taking focus from what you're using (SW_SHOWNOACTIVATE, and
// SWP_NOACTIVATE | SWP_NOZORDER). The size is of the visible window, not counting the invisible resize border,
// at 100% scaling: 1366x768 for the store screenshots, 1600x1000 otherwise. At 125% the window is 1.25 times
// that, and dev/store-shots.py scales the shots back down.
function place() {
  const [w, h] = SHOTS ? [1366, 768] : [1600, 1000];
  ps(`${WIN32} [void][P3n.U]::SetProcessDPIAware(); ${app.normal}; [void][P3n.U]::ShowWindow($h, 4);
      $k = [P3n.U]::GetDpiForWindow($h) / 96; $w = [int](${w} * $k); $hh = [int](${h} * $k);
      [void][P3n.U]::SetWindowPos($h, [System.IntPtr]::Zero, 40, 40, $w, $hh, 0x0014); Start-Sleep -Milliseconds 300;
      $r = New-Object P3n.U+RECT; $f = New-Object P3n.U+RECT; [void][P3n.U]::GetWindowRect($h, [ref]$r); [void][P3n.U]::DwmGetWindowAttribute($h, 9, [ref]$f, 16);
      [void][P3n.U]::SetWindowPos($h, [System.IntPtr]::Zero, 40, 40, $w + ($r.R-$r.L) - ($f.R-$f.L), $hh + ($r.B-$r.T) - ($f.B-$f.T), 0x0014); 'true'`);
}
// A PNG of the app's own window (PrintWindow), whatever is on top of it: other windows never get in. It's
// cut to the visible frame, without the invisible border around it.
function shoot(file) {
  ps(`${WIN32} Add-Type -AssemblyName System.Drawing; [void][P3n.U]::SetProcessDPIAware(); Start-Sleep -Milliseconds 400;
      $r = New-Object P3n.U+RECT; $f = New-Object P3n.U+RECT; [void][P3n.U]::GetWindowRect($h, [ref]$r); [void][P3n.U]::DwmGetWindowAttribute($h, 9, [ref]$f, 16);
      $bmp = New-Object System.Drawing.Bitmap ($r.R-$r.L), ($r.B-$r.T); $g=[System.Drawing.Graphics]::FromImage($bmp); $dc=$g.GetHdc(); [void][P3n.U]::PrintWindow($h, $dc, 2); $g.ReleaseHdc($dc);
      $cut = $bmp.Clone((New-Object System.Drawing.Rectangle ($f.L-$r.L), ($f.T-$r.T), ($f.R-$f.L), ($f.B-$f.T)), $bmp.PixelFormat);
      $cut.Save('${q(path.join(OUT, file))}', [System.Drawing.Imaging.ImageFormat]::Png); 'true'`);
}

// ---- test text ----------------------------------------------------------------------------------
const TITLE = 'Quarterly update';
const P1 = "In today's fast-paced world, it's worth noting that effective communication is more important than ever. Moreover, leveraging the right tools can help teams unlock their full potential and navigate the complexities of modern work.";
const P2 = 'We moved the Ljubljana office to Tivolska 50 in March, and the new lease runs to 2029.';
const P3 = 'Ultimately, this seamless transition unlocks a rich tapestry of possibilities for teams of all sizes, empowering everyone to thrive in an ever-evolving landscape.';
const P7 = 'In summary, the numbers clearly speak for themselves.';
// In MOCK runs: each paragraph becomes "Rewritten paragraph N.", and P7's rewrite starts with a minus sign,
// which Excel would read as a formula.
const mockRewrite = (text) => (text === P7 ? '-30% hours, as planned.' : text.split(/\n\s*\n/).map((p, i) => `Rewritten paragraph ${i + 1}.`).join('\n\n'));

// ---- the run ------------------------------------------------------------------------------------
const running = spawnSync('powershell', ['-NoProfile', '-Command', `@(Get-Process ${app.exe} -ErrorAction SilentlyContinue).Count`], { encoding: 'utf8' }).stdout.trim();
const started = new Date();
if (running !== '0') { console.error(`Close ${app.name} first: the run starts its own ${app.name} with the add-in sideloaded.`); process.exit(2); }
await mkdir(OUT, { recursive: true });

// WebView2 reads extra browser arguments from the environment of the app that hosts it, so the app, started
// from here, opens its panes with a debugging port. (--debug-method direct does the same but stops each
// pane at a "WebView Stop On Load" dialog until you click OK.)
const env = { ...process.env, WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS: '--remote-debugging-port=9229' };
const npx = (args) => spawn('npx', args, { cwd: ROOT, shell: true, stdio: 'inherit', env });
const exited = (p) => new Promise((r) => p.on('exit', r));
console.log('Sideloading with office-addin-debugging…');
// The local pane needs the dev server; the published one is on GitHub Pages already.
const server = LOCAL ? ['--dev-server', '"node dev/server.mjs"', '--dev-server-port', '3000'] : [];
await exited(npx(['office-addin-debugging', 'start', MANIFEST, 'desktop', '--app', APP, '--no-debug', ...server]));

let browser;
const errors = [];
try {
  // The app starts, opens a document with the pane, and the pane's WebView2 listens on 9229.
  let target;
  for (let i = 0; i < 90 && !target; i++) {
    await sleep(1000);
    try {
      browser ??= await puppeteer.connect({ browserURL: 'http://127.0.0.1:9229', defaultViewport: null });
      target = browser.targets().find((t) => t.url().startsWith(PANE));
    } catch { /* not up yet */ }
  }
  ok(`${app.name} starts with the add-in sideloaded and the pane open`, !!target);
  if (!target) throw new Error('no task pane on 9229');
  // The app sometimes replaces the pane once more while the document opens, or the pane's first load
  // stalls. So: wait for it to find the app, reload once, and if that fails too take the pane's target
  // afresh, up to three times.
  const booted = (p) => p.waitForFunction(() => !document.querySelector('#grab')?.hidden, { timeout: 30000 }).then(() => true, () => false);
  let page;
  for (let i = 0; i < 3 && !page; i++) {
    const p = await (browser.targets().find((t) => t.url().startsWith(PANE)) ?? target).page();
    if (await booted(p) || (await p.reload().catch(() => {}), await booted(p))) page = p;
  }
  ok(`the pane loads Office.js and finds ${app.name}`, !!page, page ? await page.evaluate(() => `${Office.context.diagnostics.host} ${Office.context.diagnostics.platform} ${Office.context.diagnostics.version}`) : 'the pane never got past loading');
  if (!page) throw new Error('the pane did not start');
  page.on('pageerror', (e) => errors.push(e.message));
  page.on('console', (m) => { if (m.type() === 'error') errors.push(m.text()); });

  if (RECORD) {
    // Each live answer, kept with the text it answers, for REPLAY. A long run is polled; its final state wins.
    page.on('response', async (res) => {
      const m = /^https:\/\/api\.p3n\.si\/v1\/(detect|humanize|runs\/)/.exec(res.url());
      if (!m || res.request().method() === 'OPTIONS') return;
      const body = await res.json().catch(() => null);
      if (!body?.skill || body.status !== 'succeeded') return;
      const sent = JSON.parse(res.request().postData() || 'null')?.text;
      const file = path.join(OUT, `run-${body.skill}.json`);
      const text = sent ?? JSON.parse(await readFile(file, 'utf8').catch(() => '{}')).text;
      await writeFile(file, JSON.stringify({ text, free: res.headers()['x-free-remaining'] ?? null, body }, null, 2));
    });
  }
  if (REPLAY) {
    // Answers recorded by RECORD=1, for the same text only; anything else fails rather than make something up.
    await page.setRequestInterception(true);
    page.on('request', async (r) => {
      if (!r.url().startsWith('https://api.p3n.si/')) return r.continue();
      const cors = { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Headers': 'Authorization, Content-Type, Accept', 'Access-Control-Expose-Headers': 'X-Free-Remaining' };
      if (r.method() === 'OPTIONS') return r.respond({ status: 204, headers: cors });
      const skill = /\/v1\/(\w+)$/.exec(r.url())?.[1];
      const rec = JSON.parse(await readFile(path.join(OUT, `run-${skill}.json`), 'utf8').catch(() => 'null'));
      const text = JSON.parse(r.postData() || '{}').text;
      if (!rec || rec.text !== text) {
        errors.push(`REPLAY: no recorded ${skill} answer for this text`);
        return r.respond({ status: 500, contentType: 'application/json', headers: cors, body: JSON.stringify({ error: { type: 'error', message: 'Nothing recorded for this text.' } }) });
      }
      return r.respond({ status: 200, contentType: 'application/json', headers: { ...cors, ...(rec.free ? { 'X-Free-Remaining': rec.free } : {}) }, body: JSON.stringify(rec.body) });
    });
  } else if (MOCK) {
    // Canned answers, so no free calls are spent.
    await page.setRequestInterception(true);
    page.on('request', (r) => {
      if (!r.url().startsWith('https://api.p3n.si/')) return r.continue();
      if (r.method() === 'OPTIONS') return r.respond({ status: 204, headers: { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Headers': 'Authorization, Content-Type, Accept' } });
      const text = JSON.parse(r.postData() || '{}').text ?? '';
      const out = r.url().endsWith('/v1/detect')
        ? { id: 'm', skill: 'detect', status: 'succeeded', output: { version: 'detect-mock', ai: 0.93, label: 'ai', verdict: 'Generic claims.', sentences: [{ text: text.split(/(?<=\.)\s/)[0], why: 'A stock opening.' }], examples: [] } }
        : { id: 'm', skill: 'humanize', status: 'succeeded', output: { version: 'humanize-mock', text: `${mockRewrite(text)}\n`, ai_before: 0.9, ai_after: 0.05, label_after: 'human', passes: true, invented_blocked: [] } };
      return r.respond({ status: 200, contentType: 'application/json', headers: { 'Access-Control-Allow-Origin': '*' }, body: JSON.stringify(out) });
    });
  }

  const pane = {
    settled: () => document.querySelector('#result')?.textContent.trim().length > 0 && document.querySelector('#busy').hidden,
    async useSelection() {
      await page.waitForFunction(() => /Use selection \(\d+/.test(document.querySelector('#grab').textContent), { timeout: 10000 });
      await page.evaluate(() => { document.querySelector('#text').value = ''; document.querySelector('#grab').click(); });
      await page.waitForFunction(() => document.querySelector('#text').value, { timeout: 10000 });
      return page.$eval('#text', (t) => t.value);
    },
    async run(sel) { await page.click(sel); await page.waitForFunction(pane.settled, { timeout: 150000, polling: 250 }); await sleep(1500); },
    status: () => page.evaluate(() => document.querySelector('#result [data-status]')?.textContent ?? ''),
    buttons: () => page.evaluate(() => [...document.querySelectorAll('#result button')].map((b) => b.textContent)),
    text: () => page.$eval('#result', (r) => r.innerText),
    rewrite: () => page.$eval('#result .rewrite', (r) => r.textContent).catch(() => null),
    async replace() {
      await page.evaluate(() => { const s = document.querySelector('#result [data-status]'); if (s) s.hidden = true; [...document.querySelectorAll('#result button')].find((b) => b.textContent.startsWith('Replace in')).click(); });
      await page.waitForFunction(() => !document.querySelector('#result [data-status]')?.hidden, { timeout: 15000 });
      await sleep(600);
    },
  };
  await (APP === 'word' ? inWord : inExcel)(page, pane);
  ok('no errors in the pane', errors.length === 0, errors.slice(0, 3).join(' || '));
} finally {
  await browser?.disconnect();
  if (!process.env.KEEP) {
    try { ps(`${app.quit}; 'true'`); } catch { /* already closed */ }
    // Excel can stay behind with no workbook after Quit. Only the instance this run started is stopped:
    // the run refuses to start while the app is open.
    spawnSync('powershell', ['-NoProfile', '-Command', `$t=[datetime]::Parse('${started.toISOString()}').ToLocalTime(); for ($i=0; $i -lt 20 -and (Get-Process ${app.exe} -ErrorAction SilentlyContinue | Where-Object StartTime -gt $t); $i++) { Start-Sleep -Milliseconds 500 }; Get-Process ${app.exe} -ErrorAction SilentlyContinue | Where-Object StartTime -gt $t | Stop-Process -Force`], { encoding: 'utf8', timeout: 30000 });
    await exited(npx(['office-addin-debugging', 'stop', MANIFEST]));
    await exited(npx(['office-addin-dev-settings', 'runtime-log', '--disable'])); // start turns it on; stop leaves it
  }
}
const failed = results.filter((r) => !r.pass);
console.log(`\n${results.length - failed.length}/${results.length} passed in ${app.name}${MOCK ? ' (canned API answers)' : ' (live API)'}. Screenshots in dev/out/.`);
process.exit(failed.length ? 1 : 0);

// ---- Word ---------------------------------------------------------------------------------------
async function inWord(page, pane) {
  const paras = () => ps(`@($d.Paragraphs | ForEach-Object { $_.Range.Text.TrimEnd([char]13) }) | ConvertTo-Json -Compress`);
  const styleNames = () => ps(`@($d.Paragraphs | ForEach-Object { $_.Style.NameLocal }) | ConvertTo-Json -Compress`);
  // Selects paragraph n (1-based) without its paragraph mark, or `from`..`to` with the last mark.
  const selectPara = (n) => ps(`$r=$d.Paragraphs(${n}).Range; [void]$r.MoveEnd(1,-1); $r.Select(); 'true'`);
  const selectParas = (from, to) => ps(`$r=$d.Range($d.Paragraphs(${from}).Range.Start, $d.Paragraphs(${to}).Range.End); $r.Select(); 'true'`);
  const write = () => ps(`$d.Content.Text = '${q(TITLE)}' + [char]13 + '${q(P1)}' + [char]13 + '${q(P2)}' + [char]13 + '${q(P3)}';
      $d.Paragraphs(1).Style = -2; $d.Paragraphs(2).Style = -1; $d.Paragraphs(3).Style = -1; $d.Paragraphs(4).Style = -1;
      $d.Paragraphs(4).Range.ListFormat.ApplyBulletDefault(); $d.UndoClear(); 'true'`);

  write();
  place();
  const original = paras();
  const styles = styleNames();
  ok('test document written', original.length === 4, JSON.stringify(styles));
  // The page from its top, title included, in the shots; the view only, not the selection.
  if (SHOTS) return storeShots(page, pane, () => selectPara(4), () => paras()[3], () => ps(`$app.ActiveWindow.VerticalPercentScrolled = 0; 'true'`), () => selectPara(2));

  // 1. Check: paragraph 2, selected in Word.
  selectPara(2);
  ok('Use selection reads the Word selection', (await pane.useSelection()) === P1);
  await pane.run('#check');
  const check = await pane.text();
  ok('Check for AI shows a result', /likely AI/.test(check), check.split('\n').slice(0, 2).join(' '));
  shoot('word-check.png');
  await page.screenshot({ path: path.join(OUT, 'word-pane-check.png') });

  // 2. Rewrite and Replace: paragraph 4 (a bulleted list item).
  selectPara(4);
  ok('a new selection replaces the old one', (await pane.useSelection()) === P3);
  await pane.run('#rewrite');
  const rewrite = await pane.rewrite();
  ok('Rewrite shows a rewrite', !!rewrite, rewrite ?? (await pane.text()));
  shoot('word-rewrite.png');
  await page.screenshot({ path: path.join(OUT, 'word-pane-rewrite.png') });
  if (rewrite) {
    await pane.replace();
    const after = paras();
    ok('Replace puts the rewrite in the document', after[3] === rewrite, JSON.stringify(after[3]));
    ok('the rest of the document is untouched, no paragraph added', after.length === 4 && after[0] === TITLE && after[1] === P1 && after[2] === P2, JSON.stringify(after.map((p) => p.slice(0, 20))));
    ok('the list paragraph keeps its style and bullet', JSON.stringify(styleNames()) === JSON.stringify(styles) && ps(`$d.Paragraphs(4).Range.ListFormat.ListType`) !== 0, JSON.stringify(styleNames()));
    ok('Replace reports success', /Replaced/.test(await pane.status()), await pane.status());
    shoot('word-replaced.png');
    const how = undo();
    const undone = paras();
    ok('one undo in the document reverts the Replace', JSON.stringify(undone) === JSON.stringify(original), `${how}; ${JSON.stringify(undone[3]?.slice(0, 40))}`);
  }
  if (!MOCK) return;

  // 3. Three paragraphs with different styles, selected down to the last paragraph mark.
  write();
  selectParas(2, 4);
  await pane.useSelection();
  await pane.run('#rewrite');
  await pane.replace();
  const after = paras();
  ok('three paragraphs are replaced one by one', JSON.stringify(after) === JSON.stringify([TITLE, 'Rewritten paragraph 1.', 'Rewritten paragraph 2.', 'Rewritten paragraph 3.']), JSON.stringify(after));
  ok('each paragraph keeps its style', JSON.stringify(styleNames()) === JSON.stringify(styles), JSON.stringify(styleNames()));
  const how = undo();
  const undone = paras();
  ok('one undo reverts all three', JSON.stringify(undone) === JSON.stringify(original), `${how}; ${JSON.stringify(undone.map((p) => p.slice(0, 20)))}`);

  // 4. The text is edited in Word after it was taken: Replace refuses.
  write();
  selectPara(3);
  await pane.useSelection();
  await pane.run('#rewrite');
  ps(`$f = $d.Paragraphs(3).Range.Find; [void]$f.Execute('Ljubljana', $false, $false, $false, $false, $false, $true, 0, $false, 'Maribor', 1); 'true'`);
  await pane.replace();
  ok('Replace refuses when the text changed in Word', /changed in the document/.test(await pane.status()) && paras()[2].includes('Maribor'), await pane.status());

  // 5. Text typed above the selection moves it: Replace still lands on it.
  write();
  selectPara(4);
  await pane.useSelection();
  await pane.run('#rewrite');
  ps(`$d.Paragraphs(1).Range.InsertBefore('Board: '); 'true'`);
  await pane.replace();
  const moved = paras();
  ok('Replace follows text that moved', moved[0] === `Board: ${TITLE}` && moved[3] === 'Rewritten paragraph 1.', JSON.stringify(moved.map((p) => p.slice(0, 24))));
}

// ---- Excel --------------------------------------------------------------------------------------
async function inExcel(page, pane) {
  // Notes in column A, hours in B, and a formula that makes text in B4.
  const CELLS = { A1: 'Notes', B1: 'Hours', A2: P1, B2: 12, A3: P2, B3: 30, A4: P3, A7: P7 };
  const FORMULA = '="Owner: "&"Maja, who signs off the plan by Friday."';
  const write = () => ps(`[void]$s.Cells.Clear(); $s.Name = 'Q3 plan';
      ${Object.entries(CELLS).map(([a, v]) => `$s.Range('${a}').Value2 = ${typeof v === 'number' ? v : `'${q(v)}'`};`).join(' ')}
      $s.Range('B4').Formula = '${q(FORMULA)}'; $s.Range('A1:B1').Font.Bold = $true; $s.Columns('A').ColumnWidth = 60; $s.Range('A:A').WrapText = $true; 'true'`);
  const read = (a) => ps(`@{ value = $s.Range('${a}').Value2; formula = $s.Range('${a}').Formula; hasFormula = $s.Range('${a}').HasFormula; bold = $s.Range('${a}').Font.Bold } | ConvertTo-Json -Compress`);
  const select = (a) => ps(`[void]$s.Range('${a}').Select(); 'true'`);

  write();
  place();
  ok('test sheet written', read('A4').value === P3 && read('B4').hasFormula, read('B4').value);
  if (SHOTS) return storeShots(page, pane, () => select('A4'), () => read('A4').value);

  // 1. Check: two cells, A2:A3. Each comes in as a paragraph of its own.
  select('A2:A3');
  ok('Use selection reads the selected cells', (await pane.useSelection()) === `${P1}\n\n${P2}`);
  await pane.run('#check');
  const check = await pane.text();
  ok('Check for AI shows a result', /likely AI/.test(check), check.split('\n').slice(0, 2).join(' '));
  shoot('excel-check.png');
  await page.screenshot({ path: path.join(OUT, 'excel-pane-check.png') });

  // 2. Rewrite and Replace: one cell, A4.
  select('A4');
  ok('a new selection replaces the old one', (await pane.useSelection()) === P3);
  await pane.run('#rewrite');
  const rewrite = await pane.rewrite();
  ok('Rewrite shows a rewrite', !!rewrite, rewrite ?? (await pane.text()));
  ok('Replace names the cell', (await pane.buttons()).includes('Replace in A4'), (await pane.buttons()).join(', '));
  shoot('excel-rewrite.png');
  await page.screenshot({ path: path.join(OUT, 'excel-pane-rewrite.png') });
  if (rewrite) {
    await pane.replace();
    ok('Replace puts the rewrite in the cell', read('A4').value === rewrite, JSON.stringify(read('A4').value));
    const rest = { A2: read('A2'), A3: read('A3'), B3: read('B3'), B4: read('B4'), A1: read('A1') };
    ok('the other cells are untouched, the formula and the bold header too', rest.A2.value === P1 && rest.A3.value === P2 && rest.B3.value === 30 && rest.B4.formula === FORMULA && rest.A1.bold === true, JSON.stringify(rest.B4));
    ok('Replace reports success', /Replaced the text in A4/.test(await pane.status()), await pane.status());
    shoot('excel-replaced.png');
    const how = undo();
    const back = read('A4').value;
    ok('one undo in the sheet reverts the Replace', back === P3, `${how}; ${JSON.stringify(String(back).slice(0, 40))}`);
  }
  if (!MOCK) return;

  // 3. Several cells: a rewrite can be copied, not put back.
  write();
  select('A2:A3');
  await pane.useSelection();
  await pane.run('#rewrite');
  ok('a rewrite of several cells has no Replace, and says why', !(await pane.buttons()).some((b) => b.startsWith('Replace')) && /single cell/.test(await pane.text()));

  // 4. A formula's text: checked and rewritten, never written over.
  select('B4');
  ok('a formula cell\'s text comes in', (await pane.useSelection()) === 'Owner: Maja, who signs off the plan by Friday.');
  await pane.run('#rewrite');
  ok('no Replace over a formula, and it says so', !(await pane.buttons()).some((b) => b.startsWith('Replace')) && /formula/.test(await pane.text()));

  // 5. The cell is edited after it was taken: Replace refuses.
  select('A3');
  await pane.useSelection();
  await pane.run('#rewrite');
  ps(`$s.Range('A3').Value2 = 'Edited by hand.'; 'true'`);
  await pane.replace();
  ok('Replace refuses when the cell changed', /cell changed/.test(await pane.status()) && read('A3').value === 'Edited by hand.', await pane.status());

  // 6. A rewrite that starts with a minus sign goes in as text, not as a broken formula.
  select('A7');
  await pane.useSelection();
  await pane.run('#rewrite');
  await pane.replace();
  const a7 = read('A7');
  ok('a rewrite starting with "-" goes in as text', a7.value === '-30% hours, as planned.' && !a7.hasFormula, JSON.stringify(a7));
}

// ---- store screenshots --------------------------------------------------------------------------
// In Word, first a Check of the second paragraph. Then the bulleted paragraph in Word, cell A4 in Excel (the
// same sentence): Rewrite, then Replace.
async function storeShots(page, pane, selectIt, readIt, frame = () => {}, selectCheck = null) {
  if (selectCheck) {
    selectCheck();
    ok('Use selection reads the paragraph to check', (await pane.useSelection()) === P1);
    await pane.run('#check');
    const check = await pane.text();
    ok('Check for AI shows a result', /likely AI/.test(check), check.split(/\n/).slice(0, 2).join(' '));
    frame();
    shoot(`raw-${APP}-check.png`);
  }
  selectIt();
  ok('Use selection reads the sentence', (await pane.useSelection()) === P3);
  await pane.run('#rewrite');
  const rewrite = await pane.rewrite();
  ok('Rewrite shows a rewrite', !!rewrite, rewrite ?? (await pane.text()));
  frame();
  shoot(`raw-${APP}-rewrite.png`);
  if (!rewrite) return;
  await pane.replace();
  ok('Replace puts it in place', readIt() === rewrite, await pane.status());
  // The pane as it settles, with its score in view.
  ok('the score is in view', await page.evaluate(() => { const r = document.querySelector('.score').getBoundingClientRect(); return r.top >= 0 && r.bottom <= innerHeight; }));
  frame();
  shoot(`raw-${APP}-replaced.png`);
}
