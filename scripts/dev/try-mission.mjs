#!/usr/bin/env node
/*
 * Try a Chapter 11 request before it goes into missions.js.
 *   node --no-warnings scripts/dev/try-mission.mjs path/to/11.4.js [--show]
 * The file holds one JavaScript expression, a request such as real({ id: '11.4', ... }) or
 * Object.assign(real({...}), { kind: 'excel' }), written with the helpers in scope inside missions.js
 * (real, range11, table11, sheetRef, SHEET11, PROD, D, DB, LIST ...). It is spliced in at the
 * CH11-REQUESTS marker of a temporary copy, and the answer battery runs on its id.
 * --show prints the solution's values on both sheets, and every trap's values, so you can see the data.
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';
import { api } from './node-engine.mjs';
import { evalOn, evalFill, evalSpill } from '../../assets/js/core/engine.js';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');
const require = createRequire(import.meta.url);
const Check = require(path.join(ROOT, 'assets/js/core/check.js'));
const file = process.argv[2];
if (!file) { console.log('Usage: try-mission.mjs fragment.js [--show]'); process.exit(2); }
const frag = fs.readFileSync(file, 'utf8').trim().replace(/[;,]\s*$/, '');
const src = fs.readFileSync(path.join(ROOT, 'assets/js/kicau/missions.js'), 'utf8');
if (!src.includes('/* CH11-REQUESTS */')) throw new Error('marker missing in missions.js');
const tmp = path.join(ROOT, 'assets/js/kicau', '.try-' + process.pid + '.js');
fs.writeFileSync(tmp, src.replace('  /* CH11-REQUESTS */', '  LIST.push(' + frag + ');\n  /* CH11-REQUESTS */'));
let K;
try { K = require(tmp); } catch (e) { fs.rmSync(tmp, { force: true }); console.log('The fragment does not load:', e.message); process.exit(1); }
fs.rmSync(tmp, { force: true });
const m = K.LIST[K.LIST.length - 1];
let passes = 0, fails = 0;
const ok = (c, name, extra) => { if (c) passes++; else { fails++; console.log('FAIL', name, extra === undefined ? '' : JSON.stringify(extra).slice(0, 600)); } };
const fillValues = (f, v) => (m.spill ? evalSpill(api, m.variantCells[v], m.spill.anchor, f, m.spill.cols, m.spill.rows) : evalFill(api, m.variantCells[v], m.fill.col, m.fill.first, m.fill.last, f, Check.shiftRows));
const run = (f, v) => (m.fill || m.spill ? Check.columnKey(fillValues(f, v), m.exactCase) : evalOn(api, m.variantCells[v], m.answerCell, f));
const J = (f) => Check.judge(m, f, run);
for (const k of ['id', 'from', 'title', 'intro', 'ask', 'solution', 'replies', 'note', 'hints']) ok(m[k], 'has ' + k);
ok(K.LIST.filter((x) => x.id === m.id).length === 1, 'the id is new: ' + m.id);
const solved = run(m.solution, 'base');
if (m.kind !== 'excel' || m.solution) {
  const v = J(m.solution);
  ok(v.state === 'pass', 'solution passes', v);
  ok(typeof solved === 'number' || (typeof solved === 'string' && solved !== '' && !Check.isError(solved)), 'solution has a value', solved);
  ok(!Check.same(solved, run(m.solution, 'alt'), m.precision, m.exactCase), 'the next-day sheet changes the answer', solved);
  for (const a of m.alts || []) { const r = J(a); ok(r.state === 'pass', 'alternative passes: ' + a, r); }
  for (const t of m.traps || []) { const r = J(t.formula); ok(r.state === 'wrong' && r.code === t.code, 'trap ' + t.code + ': ' + t.formula, r); }
  const ty = J(String(solved)); ok(ty.state === 'wrong' && ty.code === 'typed', 'a typed answer is caught', ty);
  if (!m.fill && !m.spill) { const fr = J(typeof solved === 'number' ? '=' + solved : '="' + solved + '"'); ok(fr.code === 'frozen', 'a frozen answer is caught', fr); }
  if (m.frozenTry) { const fr = J(m.frozenTry); ok(fr.code === 'frozen', 'frozen try: ' + m.frozenTry, fr); }
  if (m.lesson && m.lesson.exFormula) { const ex = evalOn(api, m.cells, 'Z1', m.lesson.exFormula); ok(!Check.isError(ex) && ex !== '', 'the lesson example works', ex); }
}
const text = JSON.stringify([m.intro, m.ask, m.lesson, m.replies, m.traps, m.hints, m.note, m.title]);
ok(!/—/.test(text), 'no em dash');
ok(!/[a-z]{3,}[^.!?]*,\s*\{name\}[.!?]/i.test(text), 'the learner\'s name only in a greeting, never at the end of a line');
if (process.argv.includes('--show')) {
  const show = (f) => (m.fill || m.spill ? fillValues(f, 'base').join(' | ') + '\n    next day: ' + fillValues(f, 'alt').join(' | ') : run(f, 'base') + '   next day: ' + run(f, 'alt'));
  console.log('solution', m.solution, '\n   ', show(m.solution));
  for (const t of m.traps || []) console.log('trap', t.code, t.formula, '\n   ', show(t.formula));
}
console.log(`${m.id}: ${passes} passed, ${fails} failed`);
process.exit(fails ? 1 : 0);
