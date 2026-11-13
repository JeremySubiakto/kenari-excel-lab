#!/usr/bin/env node
/*
 * Builds dist/kicau/, the folder published as a private claude.ai artifact so Kicau can be played from a link.
 *   node scripts/dev/artifact.mjs
 * What changes from index.html (the public page):
 *   - index.html is the page body only, because the artifact host adds its own <html> and <head>.
 *   - preview.html is the same page as a full document, for testing on localhost.
 *   - The formula engine's two files ship inside the folder, because an artifact cannot fetch a .wasm from a CDN.
 *     They are downloaded once from the pinned jsDelivr URL into dist/.vendor and reused.
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');
const OUT = path.join(ROOT, 'dist/kicau');
const CACHE = path.join(ROOT, 'assets/engine');
const CDN = 'https://cdn.jsdelivr.net/npm/formualizer@0.13.0/pkg/';
const ENGINE = ['formualizer_wasm_bg.js', 'formualizer_wasm_bg.wasm'];
const FILES = ['assets/css/kicau.css', 'assets/css/intro.css', 'assets/js/kicau/intro.js', 'assets/js/core/data.js', 'assets/js/core/check.js', 'assets/js/core/vba.js', 'assets/js/core/engine.js', 'assets/js/kicau/people.js', 'assets/js/kicau/missions.js', 'assets/js/kicau/app.js'];

fs.rmSync(OUT, { recursive: true, force: true });
fs.mkdirSync(path.join(OUT, 'vendor'), { recursive: true });
fs.mkdirSync(CACHE, { recursive: true });
for (const f of ENGINE) {
  const cached = path.join(CACHE, f);
  if (!fs.existsSync(cached)) {
    const r = await fetch(CDN + f);
    if (!r.ok) throw new Error('Could not download ' + f + ': ' + r.status);
    fs.writeFileSync(cached, Buffer.from(await r.arrayBuffer()));
  }
  fs.copyFileSync(cached, path.join(OUT, 'vendor', f));
}
const wasm = fs.readFileSync(path.join(OUT, 'vendor', ENGINE[1]));
if (wasm.readUInt32BE(0) !== 0x0061736d) throw new Error('The engine file is not WebAssembly');
for (const f of FILES) { fs.mkdirSync(path.dirname(path.join(OUT, f)), { recursive: true }); fs.copyFileSync(path.join(ROOT, f), path.join(OUT, f)); }
/* The practice problems load on demand from assets/js/kicau/practice/, so the whole folder ships. */
const PDIR = 'assets/js/kicau/practice';
fs.mkdirSync(path.join(OUT, PDIR), { recursive: true });
for (const f of fs.readdirSync(path.join(ROOT, PDIR))) if (f.endsWith('.js') && !f.startsWith('.')) fs.copyFileSync(path.join(ROOT, PDIR, f), path.join(OUT, PDIR, f));

const src = fs.readFileSync(path.join(ROOT, 'index.html'), 'utf8');
const head = src.slice(src.indexOf('<head>') + 6, src.indexOf('</head>'));
const body = src.slice(src.indexOf('<body>') + 6, src.indexOf('</body>'));
/* Keep the title, the mood script, fonts, styles and scripts. The host supplies charset and viewport. */
const keep = head.split('\n').filter((l) => !/<meta charset|<meta name="viewport"|<link rel="icon"|<meta name="theme-color"|<meta property=|<meta name="twitter|<meta name="author"|<link rel="canonical"|KENARI_ENGINE_BASE/.test(l)).join('\n').replace(/<title>[^<]*<\/title>/, '<title>Kicau</title>');
/* KENARI_ARTIFACT tells the app it runs inside the artifact frame, where downloads go through the host. */
const engineBase = `<script>window.KENARI_ENGINE_BASE = new URL('vendor/', document.baseURI).href; window.KENARI_ARTIFACT = true;</script>`;
const page = `${keep.trim()}\n${engineBase}\n<style>html,body{height:100%}</style>\n${body.trim()}\n`;
fs.writeFileSync(path.join(OUT, 'index.html'), page);
fs.writeFileSync(path.join(OUT, 'preview.html'), `<!doctype html>\n<html lang="en" data-v="b">\n<head>\n<meta charset="utf-8">\n<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n</head>\n<body>\n${page}</body>\n</html>\n`);
const list = [];
const walk = (d) => { for (const e of fs.readdirSync(d, { withFileTypes: true })) { const p = path.join(d, e.name); if (e.isDirectory()) walk(p); else list.push(path.relative(OUT, p)); } };
walk(OUT);
console.log(list.filter((f) => f !== 'index.html' && f !== 'preview.html').join('\n'));
console.log('Built', OUT, Math.round(list.reduce((a, f) => a + fs.statSync(path.join(OUT, f)).size, 0) / 1024) + ' KB');
