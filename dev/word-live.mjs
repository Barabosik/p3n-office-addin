// Live run in Word for Windows: sideloads the add-in with Microsoft's office-addin-debugging, writes a test
// document through Word's COM interface, and drives the real task pane over WebView2's debugging port
// (--debug-method direct opens it on 9229). Undo is a real Ctrl+Z keystroke in the document.
//   node dev/word-live.mjs         one Check and one Rewrite against the live API (2 of the 10 free calls an hour)
//   MOCK=1 node dev/word-live.mjs  canned API answers; adds the multi-paragraph and changed-text cases
//   KEEP=1                         leave Word and the dev server running afterwards
// Needs: Word desktop, closed before the run; the dev certificate (npx office-addin-dev-certs install).
import puppeteer from 'puppeteer-core';
import { spawn, spawnSync } from 'node:child_process';
import path from 'node:path';
import { mkdir } from 'node:fs/promises';

const ROOT = path.resolve(import.meta.dirname, '..');
const OUT = path.join(import.meta.dirname, 'out');
const MOCK = !!process.env.MOCK;
const results = [];
const ok = (name, pass, detail = '') => { results.push({ name, pass, detail }); console.log(`${pass ? 'PASS' : 'FAIL'}  ${name}${detail ? `  (${detail})` : ''}`); };
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// ---- Word through COM ---------------------------------------------------------------------------
// One PowerShell call per step; each prints JSON. Word is found with GetActiveObject.
function ps(script) {
  const prelude = `$ErrorActionPreference='Stop'; $w=[Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application'); $d=$w.ActiveDocument;`;
  const r = spawnSync('powershell', ['-NoProfile', '-NonInteractive', '-Command', `${prelude} ${script}`], { encoding: 'utf8' });
  if (r.status !== 0) throw new Error(`PowerShell: ${r.stderr.trim().split('\n')[0]}`);
  const out = r.stdout.trim();
  return out ? JSON.parse(out) : null;
}
const wordParas = () => ps(`@($d.Paragraphs | ForEach-Object { $_.Range.Text.TrimEnd([char]13) }) | ConvertTo-Json -Compress`);
const wordStyles = () => ps(`@($d.Paragraphs | ForEach-Object { $_.Style.NameLocal }) | ConvertTo-Json -Compress`);
// Selects paragraph n (1-based) without its paragraph mark, or `from`..`to` with the last mark.
const selectPara = (n) => ps(`$r=$d.Paragraphs(${n}).Range; [void]$r.MoveEnd(1,-1); $r.Select(); 'true'`);
const selectParas = (from, to) => ps(`$r=$d.Range($d.Paragraphs(${from}).Range.Start, $d.Paragraphs(${to}).Range.End); $r.Select(); 'true'`);

const TITLE = 'Quarterly update';
const P1 = "In today's fast-paced world, it's worth noting that effective communication is more important than ever. Moreover, leveraging the right tools can help teams unlock their full potential and navigate the complexities of modern work.";
const P2 = 'We moved the Ljubljana office to Tivolska 50 in March, and the new lease runs to 2029.';
const P3 = 'Ultimately, this seamless transition unlocks a rich tapestry of possibilities for teams of all sizes, empowering everyone to thrive in an ever-evolving landscape.';
function writeDocument() {
  const q = (s) => s.replace(/'/g, "''");
  ps(`$d.Content.Text = '${q(TITLE)}' + [char]13 + '${q(P1)}' + [char]13 + '${q(P2)}' + [char]13 + '${q(P3)}';
      $d.Paragraphs(1).Style = -2; $d.Paragraphs(2).Style = -1; $d.Paragraphs(3).Style = -1; $d.Paragraphs(4).Style = -1;
      $d.Paragraphs(4).Range.ListFormat.ApplyBulletDefault(); $d.UndoClear(); 'true'`);
}

// A real Ctrl+Z: focus the document body (not the pane), then send the keystroke to Word.
function ctrlZ() {
  return ps(`$w.Activate(); $w.ActiveWindow.SetFocus(); Start-Sleep -Milliseconds 300;
    $sh = New-Object -ComObject WScript.Shell; [void]$sh.AppActivate($w.ActiveWindow.Caption + ' - Word'); Start-Sleep -Milliseconds 300;
    Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.SendKeys]::SendWait('^z'); Start-Sleep -Milliseconds 500; 'true'`);
}

// Places Word at a fixed size on the primary screen and saves a PNG of the window (physical pixels).
function placeWord() {
  ps(`Add-Type -Namespace P3n -Name U -MemberDefinition '[DllImport("user32.dll")] public static extern bool SetProcessDPIAware(); [DllImport("user32.dll")] public static extern bool SetWindowPos(System.IntPtr h, System.IntPtr a, int x, int y, int cx, int cy, uint f); [DllImport("user32.dll")] public static extern bool ShowWindow(System.IntPtr h, int c);';
      [void][P3n.U]::SetProcessDPIAware(); $w.WindowState = 0; $h=[System.IntPtr]$w.ActiveWindow.Hwnd;
      [void][P3n.U]::ShowWindow($h, 9); [void][P3n.U]::SetWindowPos($h, [System.IntPtr]::Zero, 40, 40, 1600, 1000, 0x0040); $w.Activate(); 'true'`);
}
function shootWord(file) {
  ps(`Add-Type -Namespace P3n -Name S -MemberDefinition '[DllImport("user32.dll")] public static extern bool SetProcessDPIAware(); [DllImport("user32.dll")] public static extern bool GetWindowRect(System.IntPtr h, out RECT r); public struct RECT { public int L, T, R, B; }';
      Add-Type -AssemblyName System.Drawing; [void][P3n.S]::SetProcessDPIAware(); $h=[System.IntPtr]$w.ActiveWindow.Hwnd; $r = New-Object P3n.S+RECT; [void][P3n.S]::GetWindowRect($h, [ref]$r);
      $b = New-Object System.Drawing.Bitmap ($r.R-$r.L), ($r.B-$r.T); $g=[System.Drawing.Graphics]::FromImage($b); $g.CopyFromScreen($r.L, $r.T, 0, 0, $b.Size);
      $b.Save('${path.join(OUT, file).replace(/'/g, "''")}', [System.Drawing.Imaging.ImageFormat]::Png); 'true'`);
}

// ---- the run ------------------------------------------------------------------------------------
const running = spawnSync('powershell', ['-NoProfile', '-Command', '@(Get-Process WINWORD -ErrorAction SilentlyContinue).Count'], { encoding: 'utf8' }).stdout.trim();
if (running !== '0') { console.error('Close Word first: the run starts its own Word with the add-in sideloaded.'); process.exit(2); }
await mkdir(OUT, { recursive: true });

const npx = (args) => spawn('npx', args, { cwd: ROOT, shell: true, stdio: 'inherit' });
const exited = (p) => new Promise((r) => p.on('exit', r));
console.log('Sideloading with office-addin-debugging…');
await exited(npx(['office-addin-debugging', 'start', 'manifest.xml', 'desktop', '--app', 'word', '--debug-method', 'direct', '--dev-server', '"node dev/server.mjs"', '--dev-server-port', '3000']));

let browser;
const errors = [];
try {
  // Word starts, opens a document with the pane, and the pane's WebView2 listens on 9229.
  let target;
  for (let i = 0; i < 90 && !target; i++) {
    await sleep(1000);
    try {
      browser ??= await puppeteer.connect({ browserURL: 'http://127.0.0.1:9229', defaultViewport: null });
      target = browser.targets().find((t) => t.url().startsWith('https://localhost:3000/taskpane.html'));
    } catch { /* not up yet */ }
  }
  ok('Word starts with the add-in sideloaded and the pane open', !!target);
  if (!target) throw new Error('no task pane on 9229');
  const page = await target.page();
  page.on('pageerror', (e) => errors.push(e.message));
  page.on('console', (m) => { if (m.type() === 'error') errors.push(m.text()); });
  await page.waitForFunction(() => !document.querySelector('#grab').hidden, { timeout: 30000 });
  ok('the pane loads Office.js and finds Word (WordApi 1.3)', true, await page.evaluate(() => `${Office.context.diagnostics.host} ${Office.context.diagnostics.platform} ${Office.context.diagnostics.version}`));

  if (MOCK) {
    // Canned answers, from the same place e2e.mjs takes them, so no free calls are spent.
    await page.setRequestInterception(true);
    page.on('request', (r) => {
      if (!r.url().startsWith('https://api.p3n.si/')) return r.continue();
      if (r.method() === 'OPTIONS') return r.respond({ status: 204, headers: { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Headers': 'Authorization, Content-Type, Accept' } });
      const text = JSON.parse(r.postData() || '{}').text ?? '';
      const out = r.url().endsWith('/v1/detect')
        ? { id: 'm', skill: 'detect', status: 'succeeded', output: { version: 'detect-mock', ai: 0.93, label: 'ai', verdict: 'Generic claims.', sentences: [{ text: text.split(/(?<=\.)\s/)[0], why: 'A stock opening.' }], examples: [] } }
        : { id: 'm', skill: 'humanize', status: 'succeeded', output: { version: 'humanize-mock', text: text.split(/\n\s*\n/).map((p, i) => `Rewritten paragraph ${i + 1}.`).join('\n\n') + '\n', ai_before: 0.9, ai_after: 0.05, label_after: 'human', passes: true, invented_blocked: [] } };
      return r.respond({ status: 200, contentType: 'application/json', headers: { 'Access-Control-Allow-Origin': '*' }, body: JSON.stringify(out) });
    });
  }

  writeDocument();
  placeWord();
  const original = wordParas();
  const styles = wordStyles();
  ok('test document written', original.length === 4, JSON.stringify(styles));

  const settled = () => document.querySelector('#result')?.textContent.trim().length > 0 && document.querySelector('#busy').hidden;
  const useSelection = async () => {
    await page.waitForFunction(() => /Use selection \(\d+/.test(document.querySelector('#grab').textContent), { timeout: 10000 });
    await page.evaluate(() => { document.querySelector('#text').value = ''; document.querySelector('#grab').click(); });
    await page.waitForFunction(() => document.querySelector('#text').value, { timeout: 10000 });
    return page.$eval('#text', (t) => t.value);
  };
  const run = async (sel) => { await page.click(sel); await page.waitForFunction(settled, { timeout: 150000, polling: 250 }); await sleep(1500); };
  const status = () => page.evaluate(() => document.querySelector('#result [data-status]')?.textContent ?? '');
  const replace = async () => {
    await page.evaluate(() => [...document.querySelectorAll('#result button')].find((b) => b.textContent.startsWith('Replace')).click());
    await page.waitForFunction(() => !document.querySelector('#result [data-status]')?.hidden, { timeout: 15000 });
    await sleep(600);
  };

  // 1. Check: paragraph 2, selected in Word.
  selectPara(2);
  ok('Use selection reads the Word selection', (await useSelection()) === P1);
  await run('#check');
  const check = await page.evaluate(() => ({ text: document.querySelector('#result').innerText, error: !!document.querySelector('#result .note.error') }));
  ok('Check for AI shows a result', !check.error && /likely AI/.test(check.text), check.text.split('\n').slice(0, 2).join(' '));
  shootWord('word-check.png');
  await page.screenshot({ path: path.join(OUT, 'word-pane-check.png') });

  // 2. Rewrite and Replace: paragraph 4 (a bulleted list item).
  selectPara(4);
  ok('a new selection replaces the old one', (await useSelection()) === P3);
  await run('#rewrite');
  const rewrite = await page.$eval('#result .rewrite', (r) => r.textContent).catch(() => null);
  ok('Rewrite shows a rewrite', !!rewrite, rewrite ?? (await page.$eval('#result', (r) => r.innerText)));
  shootWord('word-rewrite.png');
  await page.screenshot({ path: path.join(OUT, 'word-pane-rewrite.png') });
  if (rewrite) {
    await replace();
    const after = wordParas();
    const afterStyles = wordStyles();
    ok('Replace puts the rewrite in the document', after[3] === rewrite, JSON.stringify(after[3]));
    ok('the rest of the document is untouched, no paragraph added', after.length === 4 && after[0] === TITLE && after[1] === P1 && after[2] === P2, JSON.stringify(after.map((p) => p.slice(0, 20))));
    ok('the list paragraph keeps its style and bullet', JSON.stringify(afterStyles) === JSON.stringify(styles) && ps(`$d.Paragraphs(4).Range.ListFormat.ListType`) !== 0, JSON.stringify(afterStyles));
    ok('Replace reports success', /Replaced/.test(await status()), await status());
    shootWord('word-replaced.png');
    ctrlZ();
    const undone = wordParas();
    ok('one Ctrl+Z in the document undoes the Replace', JSON.stringify(undone) === JSON.stringify(original), JSON.stringify(undone[3]?.slice(0, 40)));
  }

  if (MOCK) {
    // 3. Three paragraphs with different styles, selected down to the last paragraph mark.
    writeDocument();
    selectParas(2, 4);
    await useSelection();
    await run('#rewrite');
    await replace();
    const after = wordParas();
    ok('three paragraphs are replaced one by one', JSON.stringify(after) === JSON.stringify([TITLE, 'Rewritten paragraph 1.', 'Rewritten paragraph 2.', 'Rewritten paragraph 3.']), JSON.stringify(after));
    ok('each paragraph keeps its style', JSON.stringify(wordStyles()) === JSON.stringify(styles), JSON.stringify(wordStyles()));
    ctrlZ();
    const undone = wordParas();
    ok('one Ctrl+Z undoes all three', JSON.stringify(undone) === JSON.stringify(original), JSON.stringify(undone.map((p) => p.slice(0, 20))));

    // 4. The text is edited in Word after it was taken: Replace refuses.
    writeDocument();
    selectPara(3);
    await useSelection();
    await run('#rewrite');
    ps(`$f = $d.Paragraphs(3).Range.Find; [void]$f.Execute('Ljubljana', $false, $false, $false, $false, $false, $true, 0, $false, 'Maribor', 1); 'true'`);
    await replace();
    ok('Replace refuses when the text changed in Word', /changed in the document/.test(await status()) && wordParas()[2].includes('Maribor'), await status());

    // 5. Text typed above the selection moves it: Replace still lands on it.
    writeDocument();
    selectPara(4);
    await useSelection();
    await run('#rewrite');
    ps(`$d.Paragraphs(1).Range.InsertBefore('Board: '); 'true'`);
    await replace();
    const moved = wordParas();
    ok('Replace follows text that moved', moved[0] === `Board: ${TITLE}` && moved[3] === 'Rewritten paragraph 1.', JSON.stringify(moved.map((p) => p.slice(0, 24))));
  }
  ok('no errors in the pane', errors.length === 0, errors.slice(0, 3).join(' || '));
} finally {
  await browser?.disconnect();
  if (!process.env.KEEP) {
    try { ps(`$d.Saved = $true; $w.Quit(0); 'true'`); } catch { /* already closed */ }
    await exited(npx(['office-addin-debugging', 'stop', 'manifest.xml']));
  }
}
const failed = results.filter((r) => !r.pass);
console.log(`\n${results.length - failed.length}/${results.length} passed in Word${MOCK ? ' (canned API answers)' : ' (live API)'}. Screenshots in dev/out/.`);
process.exit(failed.length ? 1 : 0);
