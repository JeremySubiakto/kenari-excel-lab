#!/usr/bin/env node
/* Makes assets/preview.png, the 1200 x 630 link preview: request 1.6 on a laptop, in the default mood, with SUMIFS half typed.
     node --no-warnings scripts/dev/preview-image.mjs
   Serves on port 8766 while it runs. */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { launch, BASE, OUT, sleep } from './cdp.mjs';
import { startServer } from './serve.mjs';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');
const stop = await startServer(8766);
const b = await launch();
const p = await b.page();
await p.viewport(1200, 630, { mobile: false, dpr: 1 });
const until = async (expr, ms = 15000) => { for (let i = 0; i < ms / 50; i++) { if (await p.ev(expr)) return true; await sleep(50); } return false; };
await p.go(BASE + '/index.html?fast&reset', 400);
await until(`!!document.getElementById('wname')`);
await p.ev(`(() => { const i = document.getElementById('wname'); i.value = 'Rani'; i.dispatchEvent(new Event('input')); document.getElementById('go').click(); })()`);
await until(`document.documentElement.dataset.engine === 'ready'`, 45000);
for (let k = 0; k < 5; k++) {
  const m = await p.ev(`(() => { const c = window.__kicau.current(); return { from: c.from, solution: c.solution }; })()`);
  await p.ev(`location.hash = '#chat-${m.from}'`);
  await until(`!!document.getElementById('fx') && !document.querySelector('.typingb')`);
  await p.ev(`window.__kicau.setFormula(${JSON.stringify(m.solution)})`);
  await p.ev(`document.getElementById('send').click()`);
  await until(`window.__kicau.ST.closed.length === ${k + 1}`);
  await sleep(300);
}
await p.ev(`location.hash = '#chat-tika'`);
await until(`!!document.getElementById('fx') && !document.querySelector('.typingb')`);
await sleep(1500);
await p.ev(`window.__kicau.setFormula('=SUMIFS(E5:E24,B5:B24,B1')`);
await sleep(300);
await p.shot('preview');
const shot = path.join(OUT, 'preview.png');
fs.copyFileSync(shot, path.join(ROOT, 'assets/preview.png'));
console.log('Wrote assets/preview.png');
await b.close();
await stop();
