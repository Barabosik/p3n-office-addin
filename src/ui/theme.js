// Light or dark before the pane paints: Light or Dark if the user picked one (lib/theme.js), else the system's
// setting until Office's own theme is known. lib/theme.js takes over once the pane's script runs.
(() => {
  const root = document.documentElement;
  const media = matchMedia('(prefers-color-scheme: dark)');
  let picked = null;
  try { picked = localStorage.getItem('p3n.theme'); } catch {}
  const follow = () => { if (!root.dataset.themeFrom) root.dataset.theme = picked === 'dark' || (picked !== 'light' && media.matches) ? 'dark' : 'light'; };
  follow();
  media.addEventListener('change', follow);
})();
