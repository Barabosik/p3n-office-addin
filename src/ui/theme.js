// Light or dark before the pane paints: the system's setting, until Office's own theme is known (lib/theme.js,
// which then takes over). Outside Office, the page keeps following the system.
(() => {
  const root = document.documentElement;
  const media = matchMedia('(prefers-color-scheme: dark)');
  const follow = () => { if (root.dataset.themeFrom !== 'office') root.dataset.theme = media.matches ? 'dark' : 'light'; };
  follow();
  media.addEventListener('change', follow);
})();
