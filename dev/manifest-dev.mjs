// Writes manifest.dev.xml from manifest.xml: the same add-in, served from https://localhost:3000, with its
// own Id and name so it can sit next to the published one. dev/e2e.mjs fails if the two drift apart.
//   node dev/manifest-dev.mjs
import path from 'node:path';
import { readFile, writeFile } from 'node:fs/promises';
import { pathToFileURL } from 'node:url';

const ROOT = path.resolve(import.meta.dirname, '..');
export const PUBLISHED = 'https://nikita-andreiev.com/p3n-office-addin/';
export const LOCAL = 'https://localhost:3000/';

export function devManifest(xml) {
  return xml
    .replace(/<!--[\s\S]*?-->/, '<!-- Development manifest: the pane is served from https://localhost:3000 (npm run serve).\n     Written from manifest.xml by node dev/manifest-dev.mjs; edit that one. -->')
    .replace('<Id>ebcadc5f-371a-4c6d-b00c-8a6155a11338</Id>', '<Id>863016f5-13dc-42ac-a79d-3d9025a54dd4</Id>')
    .replace('<DisplayName DefaultValue="p3n"/>', '<DisplayName DefaultValue="p3n (dev)"/>')
    .replaceAll(PUBLISHED, LOCAL);
}

if (import.meta.url === pathToFileURL(process.argv[1]).href) {
  const xml = await readFile(path.join(ROOT, 'manifest.xml'), 'utf8');
  await writeFile(path.join(ROOT, 'manifest.dev.xml'), devManifest(xml));
  console.log('manifest.dev.xml written');
}
