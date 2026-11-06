#!/usr/bin/env node
/*
 * The answer battery in Node: every formula request's solution and alternatives pass, every trap gets its own
 * message, a typed or frozen answer is caught, and the next-day sheet really changes the answer.
 *   node --no-warnings scripts/dev/check-missions.mjs          every request
 *   node --no-warnings scripts/dev/check-missions.mjs 11.      only ids starting with 11.
 * Same rules as the battery in kicau-e2e.mjs, with the engine run directly.
 */
import { createRequire } from 'node:module';
import { api } from './node-engine.mjs';
import { evalOn, evalFill, evalSpill } from '../../assets/js/core/engine.js';
const require = createRequire(import.meta.url);
const Check = require('../../assets/js/core/check.js');
const { LIST } = require('../../assets/js/kicau/missions.js');

const only = process.argv[2] || '';
let passes = 0, fails = 0;
const ok = (c, name, extra) => { if (c) passes++; else { fails++; console.log('FAIL', name, extra === undefined ? '' : JSON.stringify(extra).slice(0, 500)); } };
const fillValues = (m, f, v) => (m.spill ? evalSpill(api, m.variantCells[v], m.spill.anchor, f, m.spill.cols, m.spill.rows) : evalFill(api, m.variantCells[v], m.fill.col, m.fill.first, m.fill.last, f, Check.shiftRows));
const run = (m) => (f, v) => (m.fill || m.spill ? Check.columnKey(fillValues(m, f, v), m.exactCase) : evalOn(api, m.variantCells[v], m.answerCell, f));

for (const m of LIST.filter((x) => x.kind === 'formula' && x.id.startsWith(only))) {
  const J = (f) => Check.judge(m, f, run(m));
  const solved = run(m)(m.solution, 'base');
  const v = J(m.solution);
  ok(v.state === 'pass', `${m.id} solution passes`, v);
  ok(typeof solved === 'number' || (typeof solved === 'string' && solved !== '' && !Check.isError(solved)), `${m.id} solution has a value`, solved);
  if (m.chapter >= 11) ok(!Check.same(solved, run(m)(m.solution, 'alt'), m.precision, m.exactCase), `${m.id} the next-day sheet changes the answer`, solved);
  for (const a of m.alts || []) { const r = J(a); ok(r.state === 'pass', `${m.id} alternative passes: ${a}`, r); }
  for (const t of m.traps || []) { const r = J(t.formula); ok(r.state === 'wrong' && r.code === t.code, `${m.id} trap ${t.code}: ${t.formula}`, r); }
  const ty = J(String(solved)); ok(ty.state === 'wrong' && ty.code === 'typed', `${m.id} a typed answer is caught`, ty);
  if (!m.fill && !m.spill) { const fr = J(typeof solved === 'number' ? '=' + solved : '="' + solved + '"'); ok(fr.code === 'frozen', `${m.id} a frozen answer is caught`, fr); }
  if (m.frozenTry) { const fr = J(m.frozenTry); ok(fr.code === 'frozen', `${m.id} frozen try: ${m.frozenTry}`, fr); }
  if (m.lesson && m.lesson.exFormula) { const ex = evalOn(api, m.cells, 'Z1', m.lesson.exFormula); ok(!Check.isError(ex) && ex !== '', `${m.id} the lesson example works`, ex); }
}
console.log(`${passes} passed, ${fails} failed`);
process.exit(fails ? 1 : 0);
