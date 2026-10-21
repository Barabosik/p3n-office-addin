/* global Office */
// The pane follows Office's theme (File > Account > Office Theme), not the system's. Dark Gray and Black get the
// dark palette (ui/base.css) on Office's own pane colour, so the pane matches the ribbon around it: the boxes sit a
// step lighter, the notes take a tint of their colour, and any text tone that would miss WCAG AA (4.5:1) on what it
// sits on is lightened until it passes. Office reports a theme change only to Outlook (OfficeThemeChanged), so the
// theme is also read again whenever the pane comes back into view.

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

function apply(theme) {
  const root = document.documentElement;
  const ground = /^#[0-9a-f]{6}$/i.test(theme.bodyBackgroundColor ?? '') ? rgb(theme.bodyBackgroundColor) : null;
  const dark = theme.isDarkTheme ?? (ground ? lum(ground) < 0.2 : false);
  root.dataset.themeFrom = 'office';
  root.dataset.theme = dark ? 'dark' : 'light';
  for (const [p] of TEXT) root.style.removeProperty(p);
  for (const p of GROUNDS) root.style.removeProperty(p);
  if (!dark || !ground) return;

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

// After Office.onReady. False when this Office doesn't share its theme (the system's setting stays).
export function followOfficeTheme() {
  const read = () => Office.context.officeTheme;
  if (!read()?.bodyBackgroundColor) return false;
  apply(read());
  try {
    if (Office.EventType.OfficeThemeChanged) Office.context.document.addHandlerAsync(Office.EventType.OfficeThemeChanged, (e) => apply(e?.officeTheme ?? read()), () => {});
  } catch { /* Word and Excel don't raise it */ }
  const again = () => { if (document.visibilityState === 'visible' && read()) apply(read()); };
  document.addEventListener('visibilitychange', again);
  window.addEventListener('focus', again);
  return true;
}
