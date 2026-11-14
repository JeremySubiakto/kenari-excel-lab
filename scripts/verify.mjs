#!/usr/bin/env node
/*
 * Fast checks with no browser: the data story, the macro stepper, the answer checker and the copy.
 *   node --no-warnings scripts/verify.mjs
 * The real engine runs in scripts/dev/e2e.mjs, in a real browser.
 */
import { createRequire } from 'node:module';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const require = createRequire(import.meta.url);
const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const Data = require(path.join(ROOT, 'assets/js/core/data.js'));
const Check = require(path.join(ROOT, 'assets/js/core/check.js'));
const VBA = require(path.join(ROOT, 'assets/js/core/vba.js'));

let fails = 0, passes = 0;
const ok = (cond, name, extra) => { if (cond) passes++; else { fails++; console.log('FAIL', name, extra === undefined ? '' : JSON.stringify(extra)); } };

/* ---------- Data ---------- */
const d = Data.build();
const hash = (o) => { let h = 0; const s = JSON.stringify(o); for (let i = 0; i < s.length; i++) h = (Math.imul(h, 31) + s.charCodeAt(i)) | 0; return h; };
const fresh = require('node:child_process').execFileSync(process.execPath, ['-e', `const D=require(${JSON.stringify(path.join(ROOT, 'assets/js/core/data.js'))});const o=D.build();let h=0;const s=JSON.stringify(o);for(let i=0;i<s.length;i++)h=(Math.imul(h,31)+s.charCodeAt(i))|0;console.log(h)`]).toString().trim();
ok(String(hash(d)) === fresh, 'the generator gives the same rows in a fresh process', [hash(d), fresh]);
ok(d.products.length === 12, '12 products');
ok(d.distributors.length === 14, '14 distributors');
ok(d.sales.length > 300, 'enough sales rows', d.sales.length);
ok(d.facilities.length >= 14, 'every distributor has a facility', d.facilities.length);
ok(d.snapshots.length === d.facilities.length * Data.MONTHS.length, 'one snapshot per facility per month');
ok(d.sales.every((s) => s.net > 0 && s.cartons > 0 && s.date <= Data.AS_OF), 'sales rows are sane');
ok(d.sales.every((s) => !s.paid_date || s.paid_date <= Data.AS_OF), 'nothing is paid after the data date');
ok(d.snapshots.every((s) => s.kol === Data.kol(s.dpd)), 'Kol follows days past due');
ok(d.facilities.every((f) => f.outstanding <= f.limit), 'outstanding stays within the limit');

/* Names that are real brands must never appear. Most of these are the big Indonesian FMCG names. */
const REAL = ['indomie', 'sedaap', 'sedap', 'sarimi', 'supermi', 'lemonilo', 'gaga', 'kapal api', 'good day', 'indocafe', 'torabika', 'luwak', 'top coffee', 'abc', 'nescafe', 'kenangan', 'gadjah', 'indofood', 'wings', 'unilever', 'mayora', 'belibis', 'bango', 'uleg', 'lifebuoy', 'giv', 'nuvo', 'dettol', 'shinzui', 'lux', 'harmony', 'medicare', 'jelita', 'bu rudy', 'kecap kenari', 'rasa group'];
const allNames = [...d.products.map((p) => p.name), ...d.distributors.map((x) => x.name), ...d.outlets.map((o) => o.outlet)];
for (const n of allNames) for (const r of REAL) {
  ok(!new RegExp('\\b' + r + '\\b', 'i').test(n), 'not a real brand: ' + n + ' / ' + r);
}

/* Nothing from the owner's real job may appear anywhere in this repository, including inside a made-up name.
   The words live outside the repository, in ~/.config/kenari/private-words.txt (one per line), so this public file
   does not hold them in any form. Without that file the scan is skipped and says so. */
const WORDS_FILE = path.join(process.env.HOME || '', '.config/kenari/private-words.txt');
const PRIVATE_WORDS = fs.existsSync(WORDS_FILE) ? fs.readFileSync(WORDS_FILE, 'utf8').split('\n').map((w) => w.trim()).filter((w) => w && !w.startsWith('#')) : [];
if (!PRIVATE_WORDS.length) console.log('Skipped the private-word scan: no ' + WORDS_FILE);
const siteFiles = [];
const walk = (dir) => { for (const f of fs.readdirSync(dir)) { const p = path.join(dir, f); if (fs.statSync(p).isDirectory()) { if (f !== 'node_modules' && f !== 'dist' && f !== '.git') walk(p); } else if (/\.(html|js|mjs|py|css|md|json)$/.test(f)) siteFiles.push(p); } };
walk(ROOT);
ok(siteFiles.length > 15, 'the private-word scan covers the repository', siteFiles.length);
for (const f of siteFiles) {
  const text = fs.readFileSync(f, 'utf8');
  for (const w of PRIVATE_WORDS) ok(!new RegExp('\\b' + w + '\\b', 'i').test(text), 'a private word in ' + path.relative(ROOT, f));
}

/* The early-warning story the course ends on. */
const ew = Data.earlyWarning();
const trouble = ew.filter((x) => x.trouble);
ok(trouble.length === 3, 'three planted distributors');
for (const t of trouble) { ok(t.net_oct < t.net_may, 'sales fall for ' + t.distributor, t); ok(t.dpd_oct > 60, 'DPD rises for ' + t.distributor, t); }
ok(ew.filter((x) => !x.trouble).every((x) => x.dpd_oct < 30), 'the healthy ones are not overdue');

const ms = Data.missionSales();
const cnt = (n) => ms.rows.filter((r) => r.distributor === n).length;
ok(ms.rows.length === 20, 'mission sheet has 20 rows');
ok(cnt(ms.target) >= 3, 'the target distributor has at least 3 rows', cnt(ms.target));
ok(cnt(ms.other) >= 2 && ms.target !== ms.other, 'the second distributor differs', cnt(ms.other));
ok(ms.example && ms.example !== ms.target, 'the example uses another distributor');

/* ---------- VBA stepper ---------- */
const MACRO = fs.readFileSync(path.join(ROOT, 'assets/js/kicau/app.js'), 'utf8').match(/const MACRO = `([\s\S]*?)`;/)[1];
const rows = Data.macroSheet();
const sheet = { A1: 'Distributor', B1: 'Outstanding', C1: 'DPD', D1: 'Status', E1: 'To review' };
rows.forEach((r, i) => { const n = i + 2; sheet['A' + n] = r.distributor; sheet['B' + n] = r.outstanding; sheet['C' + n] = r.dpd; });
const m = VBA.create(MACRO, sheet);
const st = m.runToEnd();
ok(!st.error, 'the macro runs to the end', st.error);
const expectReview = rows.filter((r) => r.dpd > 30).length;
ok(expectReview > 0 && expectReview < rows.length, 'the demo data has both kinds of row', expectReview);
ok(st.sheet.F1 === expectReview, 'F1 counts the flagged rows', [st.sheet.F1, expectReview]);
rows.forEach((r, i) => {
  const n = i + 2;
  ok(st.sheet['D' + n] === (r.dpd > 30 ? 'REVIEW' : 'OK'), 'status on row ' + n);
  ok((st.fill['D' + n] === 'yellow') === (r.dpd > 30), 'color on row ' + n);
});
ok(st.msgs[0] === expectReview + ' distributors need review', 'the message box text', st.msgs);
m.reset();
ok(m.state.steps === 0 && m.state.sheet.D2 === undefined, 'reset goes back to the start');
m.step(); m.step(); m.step();
const before = m.state.pc;
m.step(); m.back();
ok(m.state.pc === before, 'back undoes one line');
/* The current line never lands on a blank or comment line. */
m.reset();
for (let i = 0; i < 100 && !m.state.done; i++) { const line = m.lines[m.state.pc]; ok(line.trim() && !line.trim().startsWith("'"), 'step lands on code', line); m.step(); }
/* It says so plainly when it meets something outside its subset. */
let threw = '';
try { VBA.create('Sub x()\nActiveSheet.PivotTables(1).RefreshTable\nEnd Sub', {}); } catch (e) { threw = e.message; }
ok(/Line 2/.test(threw), 'an unknown statement names its line', threw);
const noDim = VBA.create('Sub x()\ny = 1\nEnd Sub', {}); const r2 = noDim.runToEnd();
ok(/has not been declared/.test(r2.error || ''), 'an undeclared variable is explained', r2.error);
const stuck = VBA.create('Sub x()\nDim i As Long\nDo While i < 1\ni = 0\nLoop\nEnd Sub', {}); const r3 = stuck.runToEnd();
ok(/ran more than/.test(r3.error || ''), 'an endless loop stops', r3.error);
const elseIf = VBA.create('Sub x()\nDim a As Long\na = 2\nIf a = 1 Then\nRange("A1").Value = "one"\nElseIf a = 2 Then\nRange("A1").Value = "two"\nElse\nRange("A1").Value = "other"\nEnd If\nRange("B1").Value = "after"\nEnd Sub', {}).runToEnd();
ok(elseIf.sheet.A1 === 'two' && elseIf.sheet.B1 === 'after', 'ElseIf takes the right branch', elseIf.sheet);

/* ---------- Answer checker, with a stand-in engine ---------- */
const mission = { solution: '=S', precision: 2, msgs: { typed: { title: 't', plain: 'p' }, value: { title: 'v', plain: 'p' }, frozen: { title: 'f', plain: 'p' } }, variants: [{ id: 'other' }], requires: [{ re: /SUMIFS/, title: 'r', plain: 'p' }], traps: [{ code: 'cartons', formula: '=C', title: 'c', plain: 'p' }] };
const values = { base: { '=S': 100, '=SUMIFS(E,B,B1)': 100, '=SUMIF(B,B1,E)': 100, '=SUMIFS(E,B,"x")': 100, '=C': 7, '=SUMIFS(D,B,B1)': 7, '=SUMIFS(A,B,B1)': 55, '=SUMIFS(BAD)': '#NAME?' }, other: { '=S': 40, '=SUMIFS(E,B,B1)': 40, '=SUMIF(B,B1,E)': 40, '=SUMIFS(E,B,"x")': 100, '=C': 7, '=SUMIFS(D,B,B1)': 7, '=SUMIFS(A,B,B1)': 55, '=SUMIFS(BAD)': '#NAME?' } };
const run = (f, v) => values[v][f];
ok(Check.judge(mission, '=SUMIFS(E,B,B1)', run).state === 'pass', 'a good formula passes');
ok(Check.judge(mission, '=SUMIFS(E;B;B1)', run).state === 'pass', 'semicolons pass, as on Indonesian settings');
ok(Check.judge(mission, '100', run).code === 'typed', 'a typed number is refused');
ok(Check.judge(mission, '=SUMIFS(E,B,"x")', run).code === 'frozen', 'a typed name fails the second sheet');
ok(Check.judge(mission, '=SUMIF(B,B1,E)', run).code === 'technique', 'the right number with the wrong tool is nudged');
ok(Check.judge(mission, '=SUMIFS(D,B,B1)', run).code === 'cartons', 'a known wrong column gets its own message');
ok(Check.judge(mission, '=SUMIFS(A,B,B1)', run).code === 'value', 'any other wrong number');
ok(Check.judge(mission, '=SUMIFS(E,B,B1', run).state === 'error', 'an open bracket stays on Task');
ok(Check.judge(mission, '=SUMIFS(E,B,B1', run).fix.formula === '=SUMIFS(E,B,B1)', 'an open bracket offers the fix');
ok(Check.judge(mission, '=SUMIFS(BAD)', run).state === 'error', 'an error value stays on Task');
ok(Check.judge(mission, '=', run).state === 'empty', 'an empty bar is noticed');
ok(Check.explain('=SUMIFZ(E,B,B1)', '#NAME?').fix.formula === '=SUMIFS(E,B,B1)', 'a typo gets a one-tap fix');
ok(Check.normalize('=IF(A1="a;b";1;2)') === '=IF(A1="a;b",1,2)', 'semicolons inside text stay');
ok(Check.explain('=IF(OR(F5>30,E5/D5>0.9),"Review","OK")', '#VALUE!').title === 'Something here is not a number', 'a #VALUE! in IF is explained as text where a number belongs');
ok(Check.explain('=SUMIFS(E5:E20,B5:B24,B1)', '#VALUE!').title === 'The ranges do not line up', 'a #VALUE! in SUMIFS is explained as ranges of different sizes');

/* ---------- Copy rules from the SQL lab's docs/AUTHORING.md ---------- */
const copyFiles = ['assets/js/core/check.js', 'index.html', 'README.md'];
const BANNED = /\b(simply|just|obviously|easy|easily|basically|trivial|of course|clearly)\b/i;
for (const f of copyFiles) {
  const text = fs.readFileSync(path.join(ROOT, f), 'utf8');
  ok(!/—/.test(text), 'no em dash in ' + f);
  const strings = f.endsWith('.js') ? (text.match(/'(?:[^'\\]|\\.)*'/g) || []).join(' ') : text.replace(/<script[\s\S]*?<\/script>/g, '').replace(/<[^>]+>/g, ' ');
  const hit = strings.match(BANNED);
  ok(!hit, 'no discouraging word in ' + f, hit && hit[0]);
}

/* ---------- Several workbooks: Excel's [file]Sheet! form, rewritten for the engine ---------- */
const { toEngine } = await import(path.join(ROOT, 'assets/js/core/engine.js'));
ok(toEngine('=VLOOKUP(C5,[master_produk.xlsx]Produk!$A$5:$D$16,4,FALSE)') === "=VLOOKUP(C5,'master_produk.xlsx|Produk'!$A$5:$D$16,4,FALSE)", 'an external reference is rewritten for the engine');
ok(toEngine("='[rate card.xlsx]Tarif'!B5") === "='rate card.xlsx|Tarif'!B5", 'a quoted external reference is rewritten');
ok(toEngine('=Diskon!B5+SUM(A1:A3)') === '=Diskon!B5+SUM(A1:A3)', 'a same-workbook sheet reference is left alone');
ok(Check.shiftRows('=VLOOKUP(C5,[master_produk.xlsx]Produk!$A$5:$D$16,4,FALSE)+Diskon!B5', 2) === '=VLOOKUP(C7,[master_produk.xlsx]Produk!$A$5:$D$16,4,FALSE)+Diskon!B7', 'filling down moves relative references on other sheets and keeps pinned ones');

/* ---------- The answer battery, run on the real engine in Node ---------- */
const battery = require('node:child_process').spawnSync(process.execPath, ['--no-warnings', path.join(ROOT, 'scripts/dev/check-missions.mjs')], { encoding: 'utf8' });
const last = (battery.stdout || '').trim().split('\n').pop();
ok(battery.status === 0, 'every formula request passes the answer battery: ' + last, (battery.stdout || '').split('\n').filter((l) => l.startsWith('FAIL')).slice(0, 5));

/* ---------- The 500 practice problems, on the real engine ---------- */
const practice = require('node:child_process').spawnSync(process.execPath, ['--no-warnings', path.join(ROOT, 'scripts/dev/check-practice.mjs')], { encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 });
const plast = (practice.stdout || '').trim().split('\n').pop();
ok(practice.status === 0, 'every practice problem passes its checks: ' + plast, (practice.stdout || '').split('\n').filter((l) => l.startsWith('FAIL')).slice(0, 5));

console.log(`${passes} passed, ${fails} failed`);
process.exit(fails ? 1 : 0);
