#!/usr/bin/env node
/*
 * Checks every practice problem on the real engine, in Node.
 *   node --no-warnings scripts/dev/check-practice.mjs              all groups, with the 500 total and quotas
 *   node --no-warnings scripts/dev/check-practice.mjs t-logic.js   one group file (its quota is checked too)
 *   add --show to print each problem's solution values on the visible and hidden sheet.
 * For each problem: the solution is Accepted and gives a real value, the hidden sheet changes the answer, every
 * alternative is Accepted, a typed or frozen answer is refused, the lesson fields exist, building it twice gives the
 * same sheets, and titles are unique across all 500.
 */
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { api } from './node-engine.mjs';
import { evalOn, evalFill, evalSpill } from '../../assets/js/core/engine.js';
const require = createRequire(import.meta.url);
const Check = require('../../assets/js/core/check.js');
const DIR = path.resolve(path.dirname(new URL(import.meta.url).pathname), '../../assets/js/kicau/practice');
const P = require(path.join(DIR, 'core.js'));
const only = process.argv.slice(2).find((a) => a.endsWith('.js'));
const show = process.argv.includes('--show');
const loaded = [];
/* A path outside the list (a model or a draft) is loaded on its own, without quotas. */
if (only && !P.FILES.some((f) => f.file === only)) { require(path.resolve(only)); for (const t of P.TEMPLATES) t.__file = only; }
for (const f of P.FILES) {
  if (only && f.file !== only) continue;
  const fp = path.join(DIR, f.file);
  if (!fs.existsSync(fp)) continue;
  const before = P.TEMPLATES.length;
  require(fp);
  for (const t of P.TEMPLATES.slice(before)) t.__file = f.file;
  loaded.push(f);
}
let passes = 0, fails = 0;
const ok = (c, name, extra) => { if (c) passes++; else { fails++; console.log('FAIL', name, extra === undefined ? '' : JSON.stringify(extra).slice(0, 400)); } };
const BANNED = /\b(simply|just|obviously|easily|basically|trivial|of course|clearly)\b/i;
const fillValues = (m, f, v) => (m.spill ? evalSpill(api, m.variantCells[v], m.spill.anchor, f, m.spill.cols, m.spill.rows) : evalFill(api, m.variantCells[v], m.fill.col, m.fill.first, m.fill.last, f, Check.shiftRows));
const run = (m) => (f, v) => (m.fill || m.spill ? Check.columnKey(fillValues(m, f, v), m.exactCase) : evalOn(api, m.variantCells[v], m.answerCell, f));

for (const t of P.TEMPLATES) {
  ok(t.key && /^[a-z0-9-]+$/.test(t.key), 'template key is lowercase words: ' + t.key);
  ok(['Easy', 'Medium', 'Hard'].includes(t.difficulty), t.key + ' has a difficulty');
  ok(t.level >= 1 && t.level <= 10, t.key + ' has a level from 1 to 10', t.level);
  ok(Array.isArray(t.topics) && t.topics.length > 0, t.key + ' has topics');
  ok(P.TEMPLATES.filter((x) => x.key === t.key).length === 1, 'template key is unique: ' + t.key);
}
const L = P.list();
const seen = new Map();
for (const meta of L) {
  const t = P.TEMPLATES.find((x) => x.key === meta.tkey);
  let m;
  try { m = P.get(meta.id); } catch (e) { ok(false, meta.key + ' builds', String(e)); continue; }
  const name = meta.id + ' ' + meta.key;
  ok(typeof m.title === 'string' && m.title.length > 4 && m.title.length <= 70, name + ' has a title up to 70 characters', m.title);
  ok(!seen.has(m.title), name + ' title is unique: ' + m.title, seen.get(m.title));
  seen.set(m.title, meta.key);
  ok(typeof m.statement === 'string' && m.statement.length > 30, name + ' has a statement');
  ok(Array.isArray(m.hints) && m.hints.length >= 2, name + ' has at least two hints');
  ok(typeof m.explain === 'string' && m.explain.length > 20, name + ' explains the solution');
  ok(typeof m.solution === 'string' && m.solution[0] === '=', name + ' has a solution');
  const text = [m.title, m.statement, m.explain].concat(m.hints).join(' ');
  ok(!/—/.test(text), name + ' has no em dash');
  ok(!BANNED.test(text), name + ' avoids discouraging words', (text.match(BANNED) || [])[0]);
  /* layout: the answer cell is empty on both sheets, and inputs stay in rows 1 to 3 */
  const ac = m.answerCell;
  ok(m.variantCells.base[ac] === undefined && m.variantCells.alt[ac] === undefined, name + ' answer cell starts empty: ' + ac);
  if (!m.fill && !m.spill) ok(/^[A-Z][1-3]$/.test(ac), name + ' a one-cell answer sits in rows 1 to 3', ac);
  /* the answer battery */
  const J = (f) => Check.judge(m, f, run(m));
  const solved = run(m)(m.solution, 'base');
  const v = J(m.solution);
  ok(v.state === 'pass', name + ' solution is Accepted', v);
  ok(typeof solved === 'number' || (typeof solved === 'string' && solved !== '' && !Check.isError(solved)), name + ' solution has a value', solved);
  const hidden = run(m)(m.solution, 'alt');
  ok(!Check.isError(hidden) && hidden !== '', name + ' solution works on the hidden sheet', hidden);
  ok(!Check.same(solved, hidden, m.precision, m.exactCase), name + ' the hidden sheet changes the answer', solved);
  for (const a of m.alts || []) { const r = J(a); ok(r.state === 'pass', name + ' alternative is Accepted: ' + a, r); }
  const ty = J(String(solved)); ok(ty.state === 'wrong' && ty.code === 'typed', name + ' a typed answer is refused', ty);
  if (!m.fill && !m.spill) { const fr = J(typeof solved === 'number' ? '=' + solved : '="' + solved + '"'); ok(fr.code === 'frozen', name + ' a frozen answer is refused', fr); }
  /* the same problem every time, other workbooks included */
  const again = t.make(P.H.rng(meta.key), meta.v, P.H);
  const ser = (vc) => JSON.stringify(['base', 'alt'].map((v) => [vc[v], vc[v][P.H.MAIN] || '', (vc[v][P.H.SHEETS] || []).map((x) => [x.name, x.cells])]));
  ok(ser(again.variantCells) === ser(m.variantCells) && again.solution === m.solution, name + ' builds the same every time');
  /* several workbooks: every open sheet exists, file names survive filling down, and multi problems really reach out */
  if (m.books) {
    const eng = new Set((m.variantCells.base[P.H.SHEETS] || []).map((x) => x.name));
    for (const b of m.books.slice(1)) for (const sh of b.sheets) ok(eng.has(b.file + '|' + sh.name), name + ' opens ' + b.file + ' ' + sh.name);
    ok(m.books.every((b, i) => i === 0 || /^[a-z0-9_]+\.xlsx$/.test(b.file)), name + ' workbook files are lowercase .xlsx names', m.books.map((b) => b.file));
    if (m.fill) { const shifted = Check.shiftRows(m.solution, 3); for (const b of m.books) for (const sh of b.sheets) if (sh.prefix) ok(shifted.split(sh.prefix).length === m.solution.split(sh.prefix).length, name + ' filling down keeps ' + sh.prefix); }
  }
  if (t.multi) ok(m.books && m.books.length >= 2 && /\[[a-z0-9_]+\.xlsx\]/.test(m.solution), name + ' reaches into another workbook', m.solution);
  if (show) console.log(meta.id, meta.difficulty, '|', m.title, '|', m.solution, '\n    visible:', String(solved).slice(0, 120), '\n    hidden: ', String(hidden).slice(0, 120));
}
/* quotas per file and in all */
for (const f of loaded) {
  const got = { Easy: 0, Medium: 0, Hard: 0 };
  for (const meta of L) if (P.TEMPLATES.find((x) => x.key === meta.tkey).__file === f.file) got[meta.difficulty]++;
  for (const d of ['Easy', 'Medium', 'Hard']) ok((f.quota[d] || 0) === got[d], `${f.file} has ${f.quota[d] || 0} ${d} problems`, got);
  if (f.multi) { const multi = L.filter((meta) => { const t = P.TEMPLATES.find((x) => x.key === meta.tkey); return t.__file === f.file && t.multi; }).length; ok(multi >= f.multi, `${f.file} has at least ${f.multi} problems across workbooks`, multi); }
}
if (!only) ok(L.length === 500 && loaded.length === P.FILES.length, 'there are 500 problems from all nine files', [L.length, loaded.length]);
/* The list page's title index must match the problems (run scripts/dev/practice-index.mjs after changes). */
if (!only) {
  const idx = path.join(DIR, 'index.js');
  let map = null;
  if (fs.existsSync(idx)) { const g = globalThis.KicauPractice; globalThis.KicauPractice = { setTitles: (t) => { map = t; } }; require(idx); globalThis.KicauPractice = g; }
  const stale = map ? L.filter((m) => map[m.key] !== P.get(m.id).title).map((m) => m.id) : ['no index.js'];
  ok(stale.length === 0, 'practice/index.js lists every title (run scripts/dev/practice-index.mjs)', stale.slice(0, 5));
}
const byD = { Easy: 0, Medium: 0, Hard: 0 }; L.forEach((x) => byD[x.difficulty]++);
console.log(`${L.length} problems (${byD.Easy} Easy, ${byD.Medium} Medium, ${byD.Hard} Hard) from ${loaded.length} files: ${passes} passed, ${fails} failed`);
process.exit(fails ? 1 : 0);
