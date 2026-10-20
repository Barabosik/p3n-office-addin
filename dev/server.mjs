// Serves src/ at https://localhost:3000 for Word, with the certificate from Microsoft's office-addin-dev-certs.
//   npx office-addin-dev-certs install    (once: makes a local CA, and Windows asks you to trust it)
//   npm run serve                          (PORT=3001 to change the port)
// This never installs the certificate itself: trusting a new root CA is your call, made in that prompt.
import https from 'node:https';
import os from 'node:os';
import path from 'node:path';
import { readFile, stat } from 'node:fs/promises';

const ROOT = path.resolve(import.meta.dirname, '..', 'src');
const CERTS = path.join(os.homedir(), '.office-addin-dev-certs');
const PORT = Number(process.env.PORT) || 3000;
const TYPES = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8', '.png': 'image/png', '.svg': 'image/svg+xml', '.json': 'application/json', '.xml': 'application/xml' };

let tls;
try {
  tls = { key: await readFile(path.join(CERTS, 'localhost.key')), cert: await readFile(path.join(CERTS, 'localhost.crt')) };
} catch {
  console.error(`No development certificate in ${CERTS}.\nRun "npx office-addin-dev-certs install" once, accept the Windows prompt, then start this again.`);
  process.exit(1);
}

const server = https.createServer(tls, async (req, res) => {
  const url = new URL(req.url, 'https://localhost');
  const file = path.join(ROOT, decodeURIComponent(url.pathname === '/' ? '/taskpane.html' : url.pathname));
  if (!file.startsWith(ROOT + path.sep)) { res.writeHead(403).end(); return; }
  try {
    if (!(await stat(file)).isFile()) throw new Error('not a file');
    const body = await readFile(file);
    res.writeHead(200, {
      'Content-Type': TYPES[path.extname(file)] ?? 'application/octet-stream',
      // Development: always the file on disk. Icons may be cached (Office asks for that in production).
      'Cache-Control': /\.png$/.test(file) ? 'public, max-age=3600' : 'no-cache',
      'X-Content-Type-Options': 'nosniff',
      'Referrer-Policy': 'no-referrer',
    });
    res.end(body);
  } catch {
    res.writeHead(404, { 'Content-Type': 'text/plain' }).end('not found');
  }
  if (process.env.LOG) console.log(req.method, url.pathname, res.statusCode);
});
server.listen(PORT, () => console.log(`p3n for Word: https://localhost:${PORT}/taskpane.html`));
