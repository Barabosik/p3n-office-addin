/* global Office */
// The pane follows Office's theme (File > Account > Office Theme), not the system's, until the user clicks the header's
// sun/moon button, which keeps the other theme on this device, like the key. Dark Gray
// and Black get the dark palette (ui/base.css) on Office's own pane colour, so the pane matches the ribbon around it:
// the boxes sit a step lighter, the notes take a tint of their colour, and any text tone that would miss WCAG AA
// (4.5:1) on what it sits on is lightened until it passes. Dark picked in a light Office theme is p3n's own dark.
// Office reports a theme change only to Outlook (OfficeThemeChanged), so the theme is also read again whenever the
// pane comes back into view. Without Office's theme (outside Office, or an Office that doesn't share it) the system's
// setting stands in for it.

const KEY = 'p3n.theme';
const WHITE = [255, 255, 255];
// Each text tone and the lightest ground it's drawn on.
const TEXT = [['--ink', '--highlight'], ['--muted', '--tint'], ['--accent-ink', '--tint'], ['--accent-ink', '--error-bg'], ['--accent-text', '--tint'], ['--amber', '--tint'], ['--green', '--tint'],
  ['--warn-ink', '--warn-bg'], ['--error-ink', '--error-bg'], ['--ok-ink', '--ok-bg']];
const GROUNDS = ['--paper', '--tint', '--rule', '--highlight', '--warn-bg', '--error-bg', '--ok-bg'];

const rgb = (hex) => hex.trim().replace('#', '').match(/../g).map((h) => parseInt(h, 16));
const hex = (c) => `#${c.map((v) => Math.round(v).toString(16).padStart(2, '0')).join('')}`;
const mix = (a, b, t) => a.map((v, i) => v + (b[i] - v) * t);
const lum = (c) => {
  const [r, g, b] = c.map((v) => { v /= 255; return v <= 0.04045 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4; });
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
};
export const contrast = (a, b) => { const [x, y] = [lum(rgb(a)), lum(rgb(b))].sort((m, n) => n - m); return (x + 0.05) / (y + 0.05); };

const root = document.documentElement;
const media = matchMedia('(prefers-color-scheme: dark)');
let office = null; // Office's theme, once it shares it

// Light or Dark once the button was clicked, else null: Office's. Storage can be missing or throw (lib/store.js); the
// pane then follows Office.
const kept = () => { try { const t = localStorage.getItem(KEY); return t === 'light' || t === 'dark' ? t : null; } catch { return null; } };

// What Office shows (or the system, without Office's theme): dark or not, and the pane colour.
function followed() {
  if (!office) return { dark: media.matches, ground: null };
  const ground = /^#[0-9a-f]{6}$/i.test(office.bodyBackgroundColor ?? '') ? rgb(office.bodyBackgroundColor) : null;
  return { dark: office.isDarkTheme ?? (ground ? lum(ground) < 0.2 : false), ground };
}

function apply() {
  const theme = kept();
  const base = followed();
  const dark = theme ? theme === 'dark' : base.dark;
  root.dataset.themeFrom = office ? 'office' : 'system';
  root.dataset.theme = dark ? 'dark' : 'light';
  const flip = dark ? 'Switch to light theme' : 'Switch to dark theme';
  document.querySelectorAll('[data-theme-toggle]').forEach((b) => { b.setAttribute('aria-label', flip); b.title = flip; });
  for (const [p] of TEXT) root.style.removeProperty(p);
  for (const p of GROUNDS) root.style.removeProperty(p);
  if (!dark || !base.dark || !base.ground) return;

  const ground = base.ground;
  const css = getComputedStyle(root);
  const token = (p) => rgb(css.getPropertyValue(p));
  const set = (p, c) => root.style.setProperty(p, hex(c));
  set('--paper', ground);
  set('--tint', mix(ground, WHITE, 0.07));
  set('--rule', mix(ground, WHITE, 0.14));
  set('--highlight', mix(mix(ground, WHITE, 0.07), token('--accent'), 0.3));
  set('--warn-bg', mix(ground, token('--warn-ink'), 0.1));
  set('--error-bg', mix(ground, token('--error-ink'), 0.1));
  set('--ok-bg', mix(ground, token('--ok-ink'), 0.1));
  for (const [p, on] of TEXT) {
    let c = token(p);
    for (let i = 0; i < 20 && contrast(hex(c), css.getPropertyValue(on)) < 4.5; i++) c = mix(c, WHITE, 0.15);
    set(p, c);
  }
}

// A change of theme lands everywhere at once, so no button fades on its own behind the rest.
function repaint() {
  const still = document.createElement('style');
  still.textContent = '*, *::before, *::after { transition: none !important; }';
  document.head.append(still);
  apply();
  getComputedStyle(root).color; // the new colours land before the transitions come back
  setTimeout(() => still.remove(), 1);
}

// At boot, before Office is ready: the header button, and the system's setting live until Office's theme is known.
export function themeSwitch() {
  apply();
  media.addEventListener('change', () => { if (!office) repaint(); });
  // Another p3n pane on this device (Word and Excel open side by side) picked one.
  window.addEventListener('storage', (e) => { if (e.key === KEY || e.key === null) repaint(); });
  document.addEventListener('click', (e) => {
    if (!e.target.closest('[data-theme-toggle]')) return;
    try { localStorage.setItem(KEY, root.dataset.theme === 'dark' ? 'light' : 'dark'); } catch {}
    repaint();
  });
}

// After Office.onReady. False when this Office doesn't share its theme (the system's setting stays).
export function followOfficeTheme() {
  const read = () => Office.context.officeTheme;
  if (!read()?.bodyBackgroundColor) return false;
  const take = (t) => { office = t; repaint(); };
  take(read());
  try {
    if (Office.EventType.OfficeThemeChanged) Office.context.document.addHandlerAsync(Office.EventType.OfficeThemeChanged, (e) => take(e?.officeTheme ?? read()), () => {});
  } catch { /* Word and Excel don't raise it */ }
  const again = () => { if (document.visibilityState === 'visible' && read()) take(read()); };
  document.addEventListener('visibilitychange', again);
  window.addEventListener('focus', again);
  return true;
}
