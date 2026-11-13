#!/usr/bin/env node
/*
 * Kicau, end to end, in a private headless Chrome with the real engine from the CDN:
 *   1. every mission's answer battery (solutions and alternatives pass, traps and typed answers fail, each with its own code)
 *   2. a full week one on a phone: first run, a wrong answer, hints, every request, the macro, the promotion
 *   3. the laptop layout of the desk, a chat, notes and career
 * Screenshots go to KENARI_SHOTS (default: the system temp folder).
 *   node --no-warnings scripts/dev/kicau-e2e.mjs
 * Serves on port 8766 for the length of the run, and closes only the browser it started.
 */
import { launch, BASE, OUT, sleep } from './cdp.mjs';
import { startServer } from './serve.mjs';

const stop = await startServer(8766);
const b = await launch();
let fails = 0, passes = 0;
const ok = (c, name, extra) => { if (c) passes++; else { fails++; console.log('FAIL', name, extra === undefined ? '' : JSON.stringify(extra).slice(0, 400)); } };
const ready = async (p) => { for (let i = 0; i < 450; i++) { if (await p.ev(`document.documentElement.dataset.engine || ''`)) return p.ev(`document.documentElement.dataset.engine`); await sleep(100); } return 'timeout'; };
const until = async (p, expr, ms = 8000) => { for (let i = 0; i < ms / 50; i++) { if (await p.ev(expr)) return true; await sleep(50); } return false; };
const tap = async (p, sel) => {
  const pos = await p.ev(`(() => { const e = document.querySelector(${JSON.stringify(sel)}); if (!e) return null; e.scrollIntoView({ block: 'nearest', inline: 'center' }); const r = e.getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2 }; })()`);
  if (!pos) return false;
  await sleep(40);
  await p.clickAt(pos.x, pos.y);
  await sleep(80);
  return true;
};
const openK = async (p, path) => { await p.go(BASE + path, 300); await until(p, `!!document.querySelector('#app') && document.querySelector('#app').children.length > 0`, 15000); await sleep(150); };
/* After a request closes, the finished sheet waits for Continue. checkStill also proves nothing moves on by itself. */
const pressContinue = async (p, name, checkStill) => {
  const held = await p.ev(`window.__kicau.ST.hold`);
  if (!held) return;
  const from = await p.ev(`window.__kicau.LIST.find((x) => x.id === window.__kicau.ST.hold).from`);
  if ((await p.ev(`location.hash`)) !== '#chat-' + from) await p.ev(`location.hash = '#chat-${from}'`);
  ok(await until(p, `!!document.querySelector('.dock.review [data-continue]')`), name + ': the finished sheet waits for Continue', held);
  if (checkStill) {
    await sleep(2500);
    ok(await p.ev(`window.__kicau.ST.hold === ${JSON.stringify(held)} && !!document.querySelector('.dock.review')`), name + ': nothing moves on without Continue');
  }
  await tap(p, '.dock.review [data-continue]');
  ok(await until(p, `!window.__kicau.ST.hold`), name + ': Continue moves on');
};
/* The Request closed card keeps its full height and its bottom can be scrolled into view. */
const doneWhole = `(() => { const d = [...document.querySelectorAll('.done')].pop(); const w = document.getElementById('wall'); if (!d || !w) return false; w.scrollTop = w.scrollHeight; const a = d.getBoundingClientRect(), b = w.getBoundingClientRect(); return d.clientHeight >= d.scrollHeight - 1 && a.bottom <= b.bottom + 1; })()`;
const locked = async (p, name) => { const s = await p.scrollInfo(); ok(s.doc <= s.vh + 1 && s.sw <= s.cw + 1, name + ': the page itself does not scroll', s); };
const inView = async (p, sel, name) => {
  const r = await p.ev(`(() => { const e = document.querySelector(${JSON.stringify(sel)}); if (!e) return null; const r = e.getBoundingClientRect(); return { top: r.top, bottom: r.bottom, left: r.left, right: r.right, vh: innerHeight, vw: innerWidth }; })()`);
  ok(r && r.top >= -1 && r.bottom <= r.vh + 1 && r.left >= -1 && r.right <= r.vw + 1, name + ': ' + sel + ' is on screen', r);
};

/* ---------- 1. The answer battery ---------- */
{
  const p = await b.page();
  await p.viewport(1366, 768, { mobile: false, dpr: 1 });
  await openK(p, '/index.html?fast&reset');
  ok(await ready(p) === 'ready', 'the engine loads');
  const report = await p.ev(`(async () => {
    const K = window.__kicau; await K.engine;
    const out = [];
    for (const m of K.LIST.filter((x) => x.kind === 'formula')) {
      const run = K.run(m), J = (f) => K.Check.judge(m, f, run);
      const solved = run(m.solution, 'base');
      out.push({ id: m.id, what: 'solution', f: m.solution, v: J(m.solution), solved });
      for (const a of m.alts || []) out.push({ id: m.id, what: 'alt', f: a, v: J(a) });
      for (const t of m.traps || []) out.push({ id: m.id, what: 'trap', want: t.code, f: t.formula, v: J(t.formula) });
      out.push({ id: m.id, what: 'typed', want: 'typed', f: String(solved), v: J(String(solved)) });
      if (!m.fill && !m.spill) { const frozen = typeof solved === 'number' ? '=' + solved : '="' + solved + '"'; out.push({ id: m.id, what: 'frozen', want: 'frozen', f: frozen, v: J(frozen) }); }
      if (m.frozenTry) out.push({ id: m.id, what: 'frozen try', want: 'frozen', f: m.frozenTry, v: J(m.frozenTry) });
      if (m.cells.B1 && m.solution.includes(',B1)')) { const named = m.solution.replace(/,B1\\)/, ',"' + m.cells.B1 + '")'); out.push({ id: m.id, what: 'frozen name', want: 'frozen', f: named, v: J(named) }); }
    }
    return out;
  })()`);
  for (const r of report) {
    if (r.what === 'solution') { ok(r.v.state === 'pass', `${r.id} solution passes`, r); ok(typeof r.solved === 'number' || typeof r.solved === 'string', `${r.id} solution has a value`, r); }
    else if (r.what === 'alt') ok(r.v.state === 'pass', `${r.id} alternative passes: ${r.f}`, r);
    else ok(r.v.state === 'wrong' && r.v.code === r.want, `${r.id} ${r.what} gives ${r.want}: ${r.f}`, r.v);
  }
  const extra = await p.ev(`(() => {
    const K = window.__kicau, m = K.LIST[0], J = (f) => K.Check.judge(m, f, K.run(m));
    const long = '=' + Array.from({ length: 20 }, (_, i) => 'D' + (5 + i)).join('+');
    return { long: J(long), open: J('=SUM(D5:D24'), typo: J('=SUMM(D5:D24)'), empty: J('=') };
  })()`);
  ok(extra.long.code === 'technique', '1.1 adding cells one by one asks for SUM', extra.long);
  ok(extra.open.state === 'error' && extra.open.fix, 'an open bracket comes with a fix', extra.open);
  ok(extra.typo.state === 'error' && /SUM/.test(extra.typo.fix && extra.typo.fix.formula), 'a typo comes with a fix', extra.typo);
  ok(extra.empty.state === 'empty', 'an empty formula is caught', extra.empty);
  ok(p.errors.length === 0, 'no console errors in the battery', p.errors);
}

/* ---------- 2. Week one on a phone ---------- */
{
  const N = 'phone';
  const p = await b.page();
  await p.viewport(390, 720, { mobile: true, dpr: 2 });
  await p.android();
  await openK(p, '/index.html?fast&reset');
  ok(await p.ev(`!!document.querySelector('.welcome')`), 'first run opens the welcome');
  ok(await p.ev(`document.getElementById('go').disabled`), 'the desk waits for a name');
  ok(await p.ev(`document.documentElement.dataset.v`) === 'b' && await p.ev(`document.querySelector('[data-mood="b"]').getAttribute('aria-pressed')`) === 'true', 'Kenari Yellow is the default mood');
  await locked(p, N + ' welcome');
  await p.shot('kicau-phone-welcome');
  await p.ev(`(() => { const i = document.getElementById('wname'); i.value = 'Rani'; i.dispatchEvent(new Event('input')); })()`);
  await tap(p, '[data-mood="a"]');
  ok(!(await p.ev(`document.getElementById('go').disabled`)), 'a name and a mood unlock the desk');
  await tap(p, '#go');
  await sleep(150);
  ok(await p.ev(`location.hash`) === '#desk', 'the desk opens');
  ok(/Rani/.test(await p.ev(`document.querySelector('.hello h1').textContent`)), 'the desk greets the learner by name');
  ok(await p.ev(`!!document.querySelector('.toast')`), 'Tika’s first message arrives as a notification');
  await ready(p);
  await locked(p, N + ' desk');
  await p.shot('kicau-phone-home');

  await p.ev(`location.hash = '#chat-tika'`);
  ok(await until(p, `!window.__kicau.ST.missions['1.1'] || !document.querySelector('.typingb') && document.querySelectorAll('.mentor.lesson').length > 0`), 'Tika’s request and Sari’s lesson play in');
  await sleep(150);
  await locked(p, N + ' chat');
  await inView(p, '#send', N + ' chat');
  await inView(p, '#fx', N + ' chat');
  const small = await p.ev(`[...document.querySelectorAll('#keys button, #chips button, .send, .coffee, thead button, .chathd a')].filter((e) => { const r = e.getBoundingClientRect(); return r.width > 0 && (r.height < 38 || r.width < 38); }).map((e) => e.textContent.trim() + ':' + Math.round(e.getBoundingClientRect().height))`);
  ok(small.length === 0, 'tap targets are large enough', small);
  ok(await p.ev(`document.getElementById('fx').getAttribute('inputmode')`) === 'none', 'tapping the grid does not raise the phone keyboard');
  ok(await p.ev(`[...document.querySelectorAll('#keys button:not(.wd)')].every((b) => getComputedStyle(b).fontVariantLigatures === 'none')`), 'formula chips show <= as typed, not as one sign');
  ok(await p.ev(`document.getElementById('fx').value`) === '=SUM(', 'the formula starts with SUM(');
  /* Build the wrong answer by tapping, as a learner would. */
  await tap(p, 'button[data-col="E"]');
  await tap(p, '#keys button[data-key=")"]');
  ok(await p.ev(`document.getElementById('fx').value`) === '=SUM(E5:E24)', 'tapping a column letter inserts its range', await p.ev(`document.getElementById('fx').value`));
  ok(await p.ev(`document.querySelector('#gw td[data-ref="E10"]').className.includes('k0')`), 'the range lights up in the grid');
  ok(await p.ev(`!!document.querySelector('#paint .p0')`), 'the reference is colored in the formula bar');
  await p.shot('kicau-phone-task');
  /* Full screen sheet: the small chat floats over it. */
  await tap(p, '#dockbig');
  ok(await until(p, `document.getElementById('dock').classList.contains('full') && getComputedStyle(document.getElementById('mini')).display !== 'none'`), 'phone: the full screen sheet has a small chat button');
  await tap(p, '#minipill');
  ok(await until(p, `/Request 1\.1/.test(document.getElementById('miniask').textContent) && document.querySelectorAll('#minilog .mm').length > 0`), 'phone: the small chat shows the request and Tika’s messages');
  await inView(p, '#miniclose', 'phone small chat');
  await sleep(400);
  await p.shot('kicau-phone-mini');
  await tap(p, '#miniclose');
  await tap(p, '#dockbig');
  ok(await until(p, `!document.getElementById('dock').classList.contains('full')`), 'phone: the sheet goes back to its usual size');
  await tap(p, '#send');
  ok(await until(p, `[...document.querySelectorAll('.b.them')].some((b) => /Net \\(Rp\\)/.test(b.textContent) && /cartons\\?/i.test(b.textContent))`), 'Tika replies in character to the rupiah mistake');
  ok(await until(p, `!!document.querySelector('.mentor.warn')`), 'Sari explains the mistake');
  ok(await p.ev(`document.querySelector('#gw td[data-ref="B2"]').className`) === 'miss', 'the answer cell shows the wrong value');
  await p.shot('kicau-phone-notyet');
  /* A coffee. */
  await tap(p, '#coffee');
  ok(await until(p, `[...document.querySelectorAll('.mentor')].some((m) => /Coffee with Mbak Sari/.test(m.textContent))`), 'a coffee brings a hint');
  ok(await p.ev(`document.querySelector('#coffee i').textContent`) === '2', 'a coffee is used');
  /* An open bracket is caught before it is sent. */
  await p.ev(`window.__kicau.setFormula('=SUM(D5:D24')`);
  await tap(p, '#send');
  ok(await until(p, `[...document.querySelectorAll('.mentor.warn')].some((m) => /before you send/.test(m.textContent))`), 'an open bracket is caught before sending');
  await tap(p, '[data-fix]');
  ok(await p.ev(`document.getElementById('fx').value`) === '=SUM(D5:D24)', 'the fix closes the bracket');
  await tap(p, '#send');
  ok(await until(p, `!!document.querySelector('.done')`), 'the right answer closes the request');
  ok(await p.ev(`window.__kicau.ST.closed.includes('1.1')`), '1.1 is saved as closed');
  await sleep(300);
  ok(await p.ev(doneWhole), 'phone: the Request closed card is shown whole');
  await p.shot('kicau-phone-pass');
  ok(await until(p, `!!document.querySelector('.dock.review .fxro')`), 'phone: the finished sheet stays up after 1.1 closes');
  ok(/SUM\(D5:D24\)/.test(await p.ev(`document.querySelector('.dock.review .fxro').textContent`)), 'phone: the review shows the formula that solved it');
  await p.shot('kicau-phone-review');
  await pressContinue(p, 'phone 1.1', true);

  /* The rest of the week. */
  for (let guard = 0; guard < 200; guard++) {
    const m = await p.ev(`(() => { const c = window.__kicau.current(); return c ? { id: c.id, from: c.from, kind: c.kind, solution: c.solution, want: c.want, steps: c.solutionSteps, line: c.solutionLine, choices: c.choices, valuesOnly: !!c.valuesOnly, pivotCheck: !!c.pivotCheck } : null; })()`);
    if (!m) break;
    await pressContinue(p, 'phone ' + m.id, false);
    if ((await p.ev(`location.hash`)) !== '#chat-' + m.from) await p.ev(`location.hash = '#chat-${m.from}'`);
    ok(await until(p, `!document.querySelector('.typingb') && [...document.querySelectorAll('.mentor.lesson')].length > 0 && !!document.querySelector('#dock')`), `${m.id}: the request plays in`);
    await sleep(120);
    if (m.kind === 'vbafix') {
      const wrong = m.choices.find((c) => c !== m.line);
      await p.ev(`window.__kicau.setLine(${JSON.stringify(wrong)})`);
      await p.ev(`window.__kicau.sendVba()`);
      ok(await until(p, `!document.querySelector('.typingb') && [...document.querySelectorAll('.mentor')].some((x) => /Read the lines around the blank|could not run/.test(x.textContent))`), `${m.id}: a wrong line gets a nudge`);
      ok(await p.ev(`!window.__kicau.ST.closed.includes('${m.id}')`), `${m.id}: a wrong line does not pass`);
      if (m.id === '9.2') await p.shot('kicau-phone-vbafix');
      await p.ev(`window.__kicau.setLine(${JSON.stringify(m.line)})`);
      await p.ev(`window.__kicau.sendVba()`);
      ok(await until(p, `window.__kicau.ST.closed.includes('${m.id}')`), `${m.id}: the right line closes it`);
    } else if (m.kind === 'query') {
      if (m.id === '7.1') {
        await tap(p, '[data-g="qop"][data-v="promote"]');
        ok(await p.ev(`/Date/.test(document.querySelector('#bprev table.pt tr').textContent)`), 'a step changes the preview at once');
        await tap(p, '#bsend');
        ok(await until(p, `[...document.querySelectorAll('.b.them')].some((b) => /empty lines/.test(b.textContent))`), 'Bayu still sees blank lines');
        await until(p, `!document.querySelector('.typingb')`);
        await p.shot('kicau-phone-query');
      }
      await p.ev(`window.__kicau.setSteps(${JSON.stringify(m.steps)})`);
      await p.ev(`window.__kicau.sendBuild()`);
      ok(await until(p, `window.__kicau.ST.closed.includes('${m.id}')`), `${m.id}: the right steps close it`);
      if (m.id === '7.6') { await sleep(300); await p.shot('kicau-phone-pass-7-6'); }
    } else if (['pivot', 'chart', 'rule', 'goal'].includes(m.kind)) {
      if (m.id === '6.1') {
        await tap(p, '[data-g="field"][data-v="Distributor"]');
        await tap(p, '[data-g="put"][data-v="rows"]');
        await tap(p, '[data-g="field"][data-v="Cartons"]');
        await tap(p, '[data-g="put"][data-v="val"]');
        ok(await p.ev(`document.querySelectorAll('#bprev table.pt tbody tr').length`) >= 5, 'tapping fields builds a live PivotTable');
        await p.shot('kicau-phone-pivot-build');
        await tap(p, '#bsend');
        ok(await until(p, `[...document.querySelectorAll('.b.them')].some((b) => /not rupiah/.test(b.textContent))`), 'Tika says the values are not rupiah');
        await until(p, `!document.querySelector('.typingb')`);
      }
      await p.ev(`window.__kicau.setCfg(${JSON.stringify(m.want)})`);
      await p.ev(`window.__kicau.sendBuild()`);
      ok(await until(p, `window.__kicau.ST.closed.includes('${m.id}')`), `${m.id}: the right ${m.kind} closes it`);
      if (m.id === '6.3' || m.id === '6.8' || m.id === '6.9') { await sleep(300); await p.shot('kicau-phone-pass-' + m.id.replace('.', '-')); }
    } else if (m.pivotCheck) {
      const passed = await p.ev(`(async () => {
        const K = window.__kicau, m = K.current();
        const wb = new ExcelJS.Workbook(); await wb.xlsx.load(await (await K.buildXlsx(m)).arrayBuffer());
        const ws = wb.addWorksheet('Pivot'); ws.getCell('A3').value = 'Row Labels'; ws.getCell('B3').value = 'Sum of Net (Rp)';
        Object.entries(K.MS.DATA_TOTALS).forEach(([k, v], i) => { ws.getCell('A' + (4 + i)).value = k; ws.getCell('B' + (4 + i)).value = v; });
        await K.uploadBuffer(await wb.xlsx.writeBuffer(), 'laporan.xlsx');
        return K.ST.closed.includes(m.id);
      })()`);
      ok(passed, 'a workbook with a real PivotTable layout passes');
      await sleep(300);
      await p.shot('kicau-phone-pivot-excel');
    } else if (m.kind === 'excel') {
      await p.shot('kicau-phone-excel');
      const blobOk = await p.ev(`(async () => { const b = await window.__kicau.buildXlsx(); const wb = new ExcelJS.Workbook(); await wb.xlsx.load(await b.arrayBuffer()); const m = window.__kicau.current(); const ws = wb.getWorksheet(m.sheetName || 'Sheet1'); const h = m.spill ? 'G4' : m.fill ? m.fill.col + (m.fill.first - 1) : 'A1'; return { size: b.size, head: ws.getCell(h).value, want: m.cells[h], req: !!wb.getWorksheet('Request') }; })()`);
      ok(blobOk.size > 4000 && blobOk.head === blobOk.want && blobOk.req, 'the request downloads as a real .xlsx', blobOk);
      /* A file with typed grades first, then one with formulas. */
      const upload = (withFormulas) => p.ev(`(async () => {
        const K = window.__kicau, m = K.current();
        const want = m.wantValues || (m.fill || m.spill ? K.fillValues(m, m.solution, 'base') : [K.run(m)(m.solution, 'base')]);
        const wb = new ExcelJS.Workbook(); await wb.xlsx.load(await (await K.buildXlsx(m)).arrayBuffer());
        const ws = wb.getWorksheet(m.sheetName || 'Sheet1');
        if (m.spill) {
          want.forEach((v, i) => { const ref = String.fromCharCode(71 + (i % m.spill.cols)) + (5 + Math.floor(i / m.spill.cols)); ws.getCell(ref).value = i === 0 && ${withFormulas} ? { formula: m.solution.slice(1), result: v } : (v === '' ? null : v); });
        } else if (!m.fill) { ws.getCell(m.answerCell).value = ${withFormulas} ? { formula: m.solution.slice(1), result: want[0] } : want[0];
        } else want.forEach((v, i) => { const r = m.fill.first + i; ws.getCell(m.fill.col + r).value = ${withFormulas} ? { formula: K.Check.shiftRows(m.solution, i).slice(1), result: v } : v; });
        await K.uploadBuffer(await wb.xlsx.writeBuffer(), 'koleksi_oktober.xlsx');
        return K.ST.closed.includes(m.id);
      })()`);
      if (m.valuesOnly) { ok(await upload(false) === true, 'a macro\'s results pass as plain values'); continue; }
      ok(await upload(false) === false, 'typed grades in the upload are refused');
      ok(await until(p, `[...document.querySelectorAll('.b.them')].some((b) => /typed in, not formulas/.test(b.textContent))`), 'the colleague says the answers are typed');
      await until(p, `!document.querySelector('.typingb') && [...document.querySelectorAll('.mentor')].some((x) => /upload again/.test(x.textContent))`);
      await sleep(100);
      ok(await upload(true) === true, 'a file with formulas passes');
      await sleep(300);
      await p.shot('kicau-phone-excel-pass');
    } else if (m.kind === 'formula') {
      if (m.id === '1.6') await p.shot('kicau-phone-task-1-6');
      if (m.id === '2.2') { await p.ev(`window.__kicau.setFormula('=E5*B1')`); await tap(p, '#send'); ok(await until(p, `[...document.querySelectorAll('.b.them')].some((b) => /Pin it with/.test(b.textContent))`), 'Citra explains the missing $'); await p.shot('kicau-phone-nodollar'); }
      if (m.id === '2.5') await p.shot('kicau-phone-task-2-5');
      if (m.id === '4.3') { await p.ev(`window.__kicau.setFormula('=UPPER(B5)')`); await tap(p, '#send'); ok(await until(p, `[...document.querySelectorAll('.b.them')].some((b) => /extra spaces are still there/.test(b.textContent))`), 'Bayu spots the untrimmed spaces'); await until(p, `!document.querySelector('.typingb')`); await p.shot('kicau-phone-notrim'); }
      if (m.id === '4.8') { await sleep(100); await p.shot('kicau-phone-task-4-8'); }
      if (m.id === '5.3') { await sleep(100); await p.shot('kicau-phone-task-5-3'); }
      if (m.id === '3.1') { await p.ev(`window.__kicau.setFormula('=VLOOKUP(C5,$G$5:$K$16,2,FALSE')`); await sleep(100); await p.shot('kicau-phone-task-3-1'); }
      if (m.id === '3.4') { await p.ev(`window.__kicau.setFormula('=XLOOKUP(C5,$G$5:$G$16,$H$5:$H$16)')`); await tap(p, '#send'); ok(await until(p, `[...document.querySelectorAll('.b.them')].some((b) => /Not in list/.test(b.textContent) && /N\\/A/.test(b.textContent))`), 'Bayu asks for Not in list'); await until(p, `!document.querySelector('.typingb')`); await p.shot('kicau-phone-notfound'); }
      await p.ev(`window.__kicau.setFormula(${JSON.stringify(m.solution)})`);
      await tap(p, '#send');
      ok(await until(p, `window.__kicau.ST.closed.includes('${m.id}')`), `${m.id}: the solution closes it`);
      if (m.id === '1.6') { await sleep(300); await p.shot('kicau-phone-pass-1-6'); }
      if (m.id === '2.5') { await sleep(300); await p.shot('kicau-phone-pass-2-5'); }
      if (m.id === '5.3') { await sleep(300); await p.shot('kicau-phone-pass-5-3'); }
    } else {
      await p.shot('kicau-phone-macro');
      for (let i = 0; i < 6; i++) await tap(p, '#mstep');
      ok(await p.ev(`window.__vba && window.__vba.steps`) >= 6, 'the macro steps line by line');
      await p.shot('kicau-phone-macro-mid');
      await locked(p, N + ' macro');
      ok(await p.ev(`Math.round(document.querySelector('.chathd').getBoundingClientRect().top)`) >= 0, 'the macro keeps the chat header on screen', await p.ev(`document.querySelector('.chathd').getBoundingClientRect().top`));
      await inView(p, '#mstep', N + ' macro');
      await tap(p, '#mend');
      ok(await until(p, `!document.getElementById('mmsg').hidden`), 'the macro ends with its message box');
      await tap(p, '#mok');
      ok(await until(p, `!!document.querySelector('[data-pick]')`), 'Bu Lestari asks how many need review');
      const answer = await p.ev(`window.__kicau.LIST.find((x) => x.id === '1.8').answer`);
      const wrong = await p.ev(`[...document.querySelectorAll('[data-pick]')].map((b) => +b.dataset.pick).find((n) => n !== ${answer})`);
      await tap(p, `[data-pick="${wrong}"]`);
      ok(await until(p, `[...document.querySelectorAll('.mentor')].some((x) => /F1/.test(x.textContent))`), 'a wrong count gets a nudge toward F1');
      await tap(p, `[data-pick="${answer}"]`);
      ok(await until(p, `window.__kicau.ST.closed.includes('1.8')`, 12000), '1.8 closes');
      ok(!(await p.ev(`!!document.querySelector('.promo')`)), 'the promotion waits for Continue');
      await pressContinue(p, 'phone 1.8', false);
      ok(await until(p, `!!document.querySelector('.promo')`, 12000), 'the promotion appears');
      await sleep(600);
      await p.shot('kicau-phone-promo');
      await tap(p, '#promook');
      ok(await until(p, `location.hash === '#chat-citra'`), 'the promotion opens Bank Kenari’s chat');
      await sleep(200);
      await p.shot('kicau-phone-bank');
      ok(await p.ev(`[...document.querySelectorAll('.row b')].filter((b) => /^Citra/.test(b.textContent)).length <= 1`), 'Citra is listed once');
    }
  }
  ok(!(await p.ev(`!!document.querySelector('.promo')`)), 'the graduation waits for Continue');
  await pressContinue(p, 'phone last request', false);
  ok(await until(p, `!!document.querySelector('.promo')`), 'the graduation screen appears');
  await sleep(500); await p.shot('kicau-phone-finale');
  await p.ev(`document.querySelector('.promo') && document.querySelector('.promo').remove()`);
  ok(await p.ev(`window.__kicau.ST.closed.length === window.__kicau.LIST.length`), 'every request closed', await p.ev(`window.__kicau.ST.closed`));
  ok(await p.ev(`window.__kicau.ST.promoted`) === true, 'promoted to Analyst');
  await p.ev(`location.hash = '#notes'`); await sleep(200);
  ok(await p.ev(`document.querySelectorAll('.note:not(.locked)').length === window.__kicau.LIST.length`), 'a note for every request');
  await locked(p, N + ' notes');
  await p.shot('kicau-phone-notes');
  await p.ev(`location.hash = '#career'`); await sleep(200);
  ok(await p.ev(`document.querySelectorAll('.badge:not(.off)').length`) >= 8, 'badges earned', await p.ev(`document.querySelectorAll('.badge:not(.off)').length`));
  await locked(p, N + ' career');
  await p.shot('kicau-phone-career');
  await p.ev(`location.hash = '#desk'`); await sleep(200);
  await p.shot('kicau-phone-home-after');
  /* Progress survives a reload. */
  await openK(p, '/index.html?fast#notes');
  ok(await p.ev(`document.querySelectorAll('.note:not(.locked)').length === window.__kicau.LIST.length`), 'progress survives a reload');
  /* The other moods render. */
  for (const v of ['b', 'c']) {
    await p.ev(`(() => { const s = JSON.parse(localStorage.getItem('kicau.v1')); s.mood = '${v}'; localStorage.setItem('kicau.v1', JSON.stringify(s)); })()`);
    await openK(p, `/index.html?fast&mood=${v}#desk`);
    ok(await p.ev(`document.documentElement.dataset.v`) === v, `mood ${v} applies`);
    await p.shot(`kicau-phone-home-${v}`);
  }
  ok(p.errors.length === 0, 'no console errors on the phone run', p.errors);
}

/* ---------- 3. The laptop ---------- */
{
  const N = 'laptop';
  const p = await b.page();
  await p.viewport(1366, 768, { mobile: false, dpr: 1 });
  await openK(p, '/index.html?fast&reset');
  await p.ev(`(() => { const i = document.getElementById('wname'); i.value = 'Rani'; i.dispatchEvent(new Event('input')); })()`);
  await p.shot('kicau-laptop-welcome');
  await tap(p, '[data-mood="a"]');
  await tap(p, '#go');
  await ready(p);
  await sleep(200);
  await locked(p, N + ' desk');
  await p.shot('kicau-laptop-home');
  for (const id of ['1.1', '1.2', '1.3', '1.4', '1.5']) {
    await p.ev(`window.__kicau.ST.hold && window.__kicau.cont()`);
    const m = await p.ev(`(() => { const c = window.__kicau.current(); return { from: c.from, solution: c.solution }; })()`);
    await p.ev(`location.hash = '#chat-${m.from}'`);
    await until(p, `!document.querySelector('.typingb') && !!document.querySelector('#fx')`);
    await p.ev(`window.__kicau.setFormula(${JSON.stringify(m.solution)})`);
    await p.ev(`document.getElementById('send').click()`);
    ok(await until(p, `window.__kicau.ST.closed.includes('${id}')`), `${N} ${id} closes`);
  }
  await p.ev(`window.__kicau.ST.hold && window.__kicau.cont()`);
  await p.ev(`location.hash = '#chat-tika'`);
  await until(p, `!document.querySelector('.typingb') && !!document.querySelector('#fx')`);
  /* Drag across cells to make a range, as with a mouse in Excel. */
  await p.ev(`window.__kicau.setFormula('=SUMIFS(')`);
  const fit = await p.ev(`(() => { const g = document.getElementById('gw'); return { table: g.querySelector('table').offsetWidth, room: g.clientWidth }; })()`);
  ok(fit.table <= fit.room + 1, 'all five columns fit the laptop workbook', fit);
  const a = await p.ev(`(() => { const r = document.querySelector('#gw td[data-ref="E5"]').getBoundingClientRect(); return { x: r.left + 10, y: r.top + 10 }; })()`);
  const z = await p.ev(`(() => { const e = document.querySelector('#gw td[data-ref="E12"]'); const r = e.getBoundingClientRect(); return { x: r.left + 10, y: r.top + 10 }; })()`);
  await p.s('Input.dispatchMouseEvent', { type: 'mouseMoved', x: a.x, y: a.y });
  await p.s('Input.dispatchMouseEvent', { type: 'mousePressed', x: a.x, y: a.y, button: 'left', buttons: 1, clickCount: 1 });
  await p.s('Input.dispatchMouseEvent', { type: 'mouseMoved', x: z.x, y: z.y, button: 'left', buttons: 1 });
  await p.s('Input.dispatchMouseEvent', { type: 'mouseReleased', x: z.x, y: z.y, button: 'left', buttons: 0, clickCount: 1 });
  await sleep(100);
  ok(await p.ev(`document.getElementById('fx').value`) === '=SUMIFS(E5:E12', 'dragging across cells inserts a range', await p.ev(`document.getElementById('fx').value`));
  await p.ev(`window.__kicau.setFormula('=SUMIFS(E5:E24,B5:B24,B1')`);
  await locked(p, N + ' chat');
  await inView(p, '#send', N);
  await p.shot('kicau-laptop-task');
  await p.ev(`window.__kicau.setFormula('=SUMIFS(D5:D24,B5:B24,B1)')`);
  await p.ev(`document.getElementById('send').click()`);
  ok(await until(p, `!!document.querySelector('.mentor.warn')`), `${N}: Tika answers the cartons mistake`);
  await p.shot('kicau-laptop-notyet');
  await p.ev(`window.__kicau.setFormula('=SUMIFS(E5:E24,B5:B24,B1)')`);
  await p.ev(`document.getElementById('send').click()`);
  ok(await until(p, `window.__kicau.ST.closed.includes('1.6')`), `${N} 1.6 closes`);
  await sleep(400);
  ok(await p.ev(doneWhole), `${N}: the Request closed card is shown whole in a long chat`);
  await p.shot('kicau-laptop-pass');
  await p.ev(`location.hash = '#notes'`); await sleep(200);
  await p.shot('kicau-laptop-notes');
  await p.ev(`location.hash = '#career'`); await sleep(200);
  await p.shot('kicau-laptop-career');
  await p.ev(`location.hash = '#desk'`); await sleep(200);
  await p.shot('kicau-laptop-home-after');
  ok(p.errors.length === 0, 'no console errors on the laptop run', p.errors);
}

/* ---------- 3b. Chapter 11 on a laptop: several workbooks, review before Continue, the maximized sheet ---------- */
{
  const N = 'workbooks';
  const p = await b.page();
  await p.viewport(1366, 768, { mobile: false, dpr: 1 });
  await openK(p, '/index.html?fast&reset');
  await p.ev(`(() => { const i = document.getElementById('wname'); i.value = 'Rani'; i.dispatchEvent(new Event('input')); })()`);
  await tap(p, '#go');
  await ready(p);
  /* Jump to Chapter 11: everything before it closed. */
  await p.ev(`(() => { const K = window.__kicau; K.ST.closed = K.LIST.filter((m) => m.chapter !== 11).map((m) => m.id); K.ST.promoted = true; K.ST.hold = null; })()`);
  const first = await p.ev(`window.__kicau.current().id`);
  ok(first === '11.1', 'Chapter 11 starts at 11.1', first);
  await p.ev(`location.hash = '#chat-bayu'`);
  ok(await until(p, `!document.querySelector('.typingb') && !!document.querySelector('#fx')`, 15000), `${N}: 11.1 plays in`);
  ok(await p.ev(`document.querySelectorAll('#bktabs .bk').length`) === 2, `${N}: one tab per open workbook`, await p.ev(`document.querySelectorAll('#bktabs .bk').length`));
  await p.ev(`window.__kicau.setFormula('=D5*VLOOKUP(C5,')`);
  await tap(p, '#bktabs [data-book="1"]');
  ok(await until(p, `/SKU/.test(document.querySelector('#gw td[data-ref="A4"]').textContent)`), `${N}: the product master shows when its tab is tapped`);
  ok(await p.ev(`document.getElementById('fx').value`) === '=D5*VLOOKUP(C5,', `${N}: switching workbooks keeps the formula`);
  await p.shot('workbooks-other-book');
  /* Drag across the price table, as with a mouse. */
  const a = await p.ev(`(() => { const r = document.querySelector('#gw td[data-ref="A5"]').getBoundingClientRect(); return { x: r.left + 8, y: r.top + 8 }; })()`);
  const z = await p.ev(`(() => { const r = document.querySelector('#gw td[data-ref="D16"]').getBoundingClientRect(); return { x: r.left + 8, y: r.top + 8 }; })()`);
  await p.s('Input.dispatchMouseEvent', { type: 'mouseMoved', x: a.x, y: a.y });
  await p.s('Input.dispatchMouseEvent', { type: 'mousePressed', x: a.x, y: a.y, button: 'left', buttons: 1, clickCount: 1 });
  await p.s('Input.dispatchMouseEvent', { type: 'mouseMoved', x: z.x, y: z.y, button: 'left', buttons: 1 });
  await p.s('Input.dispatchMouseEvent', { type: 'mouseReleased', x: z.x, y: z.y, button: 'left', buttons: 0, clickCount: 1 });
  await sleep(120);
  ok(await p.ev(`document.getElementById('fx').value`) === '=D5*VLOOKUP(C5,[master_produk.xlsx]Produk!$A$5:$D$16', `${N}: a range in another workbook is written as Excel writes it`, await p.ev(`document.getElementById('fx').value`));
  ok(await p.ev(`document.querySelector('#gw td[data-ref="B9"]').className.includes('k')`), `${N}: the range lights up in the other workbook`);
  await tap(p, '#bktabs [data-book="0"]');
  ok(await until(p, `/Date/.test(document.querySelector('#gw td[data-ref="A4"]').textContent)`), `${N}: back to the sales lines`);
  ok(!(await p.ev(`document.querySelector('#gw td[data-ref="B9"]').className.includes('k1')`)), `${N}: the other workbook's range is not drawn on this sheet`);
  await tap(p, '#gw td[data-ref="D5"]');
  /* Column letters in another workbook insert its whole column, pinned. */
  await p.ev(`window.__kicau.setFormula('=')`);
  await tap(p, '#bktabs [data-book="1"]');
  await tap(p, '#gw button[data-col="D"]');
  ok(await p.ev(`document.getElementById('fx').value`) === '=[master_produk.xlsx]Produk!$D$5:$D$16', `${N}: a column in another workbook`, await p.ev(`document.getElementById('fx').value`));
  await tap(p, '#bktabs [data-book="0"]');
  const files = await p.ev(`(async () => (await window.__kicau.buildBooks()).map((f) => f.filename + ':' + f.blob.size))()`);
  ok(files.length === 2 && /^penjualan_q4\.xlsx:/.test(files[0]) && /^master_produk\.xlsx:/.test(files[1]), `${N}: Download for Excel gives both workbooks, with their real names`, files);
  /* Maximize the sheet, then bring the chat back. */
  const w0 = await p.ev(`document.getElementById('gw').getBoundingClientRect().width`);
  await tap(p, '#dockmax');
  ok(await p.ev(`document.querySelector('.lap').classList.contains('max') && getComputedStyle(document.querySelector('.chatcol')).display === 'none'`), `${N}: maximize hides the chat`);
  const w1 = await p.ev(`document.getElementById('gw').getBoundingClientRect().width`);
  ok(w1 > w0 + 400, `${N}: the maximized sheet is much wider`, [w0, w1]);
  await p.ev(`window.__kicau.setFormula('=VLOOKUP(C5,[master_produk.xlsx]Produk!$A$5:$D$16,4,FALSE)')`);
  await p.ev(`document.getElementById('send').click()`);
  ok(await until(p, `document.getElementById('mini').classList.contains('unread') && /^(Bayu|Mbak Sari): /.test(document.getElementById('minitext').textContent)`), `${N}: while maximized, the newest reply shows on the small chat button`, await p.ev(`document.getElementById('minitext').textContent`));
  await p.shot('workbooks-max');
  await tap(p, '#minipill');
  ok(await until(p, `document.getElementById('mini').classList.contains('open')`), `${N}: the small chat opens over the sheet`);
  ok(await p.ev(`/Request 11\.1/.test(document.getElementById('miniask').textContent)`), `${N}: the small chat pins the request`);
  ok(await p.ev(`[...document.querySelectorAll('#minilog .mm.me code')].some((c) => /VLOOKUP/.test(c.textContent)) && document.querySelectorAll('#minilog .mm:not(.me)').length > 0`), `${N}: the small chat shows the sent formula and the replies`);
  await p.shot('workbooks-mini');
  await tap(p, '#miniclose');
  ok(await p.ev(`!document.getElementById('mini').classList.contains('open')`), `${N}: the small chat closes`);
  await tap(p, '#dockmax');
  ok(await p.ev(`!document.querySelector('.lap').classList.contains('max')`), `${N}: the chat comes back`);
  await p.ev(`window.__kicau.setFormula(window.__kicau.current().solution)`);
  await p.ev(`document.getElementById('send').click()`);
  ok(await until(p, `window.__kicau.ST.closed.includes('11.1')`), `${N}: 11.1 closes`);
  ok(await until(p, `!!document.querySelector('.dock.review .fxro')`), `${N}: the finished sheet stays up`);
  await sleep(2200);
  ok(await p.ev(`window.__kicau.ST.hold === '11.1' && !!document.querySelector('.dock.review')`), `${N}: nothing moves on without Continue`);
  await p.shot('workbooks-review');
  ok(p.errors.length === 0, `${N}: no console errors`, p.errors);
}

/* ---------- 3c. The opening: once per visit, skippable, calm with reduced motion ---------- */
for (const [N, w, h, mobile] of [['opening phone', 390, 844, true], ['opening laptop', 1366, 768, false]]) {
  const p = await b.page();
  await p.viewport(w, h, { mobile, dpr: 1 });
  if (mobile) await p.android();
  const t0 = Date.now();
  await p.go(BASE + '/index.html?reset', 0);
  ok(await until(p, `!!document.getElementById('kc-intro')`, 3000), `${N}: the opening plays when the app opens`);
  await sleep(1200);
  await p.shot(N.replace(' ', '-') + '-mid');
  ok(await until(p, `!document.getElementById('kc-intro')`, 7000), `${N}: the opening ends by itself`);
  const took = Date.now() - t0;
  ok(took < 6500, `${N}: the opening is short`, took);
  ok(await until(p, `!!document.querySelector('.welcome')`), `${N}: it hands over to the welcome screen`);
  await locked(p, N + ' after');
  await p.go(BASE + '/index.html', 600);
  ok(await p.ev(`!document.getElementById('kc-intro')`), `${N}: a reload in the same visit does not play it again`);
  await p.go(BASE + '/index.html?intro', 0);
  ok(await until(p, `!!document.getElementById('kc-intro')`, 3000), `${N}: ?intro plays it again`);
  await sleep(700);
  if (mobile) await tap(p, '#kc-intro'); else await p.s('Input.dispatchKeyEvent', { type: 'keyDown', key: 'Escape', code: 'Escape', windowsVirtualKeyCode: 27 });
  ok(await until(p, `!document.getElementById('kc-intro')`, 900), `${N}: ${mobile ? 'a tap' : 'a key'} skips it`);
  await p.s('Emulation.setEmulatedMedia', { features: [{ name: 'prefers-reduced-motion', value: 'reduce' }] });
  const r0 = Date.now();
  await p.go(BASE + '/index.html?intro', 0);
  ok(await until(p, `!!document.getElementById('kc-intro')`, 3000), `${N}: reduced motion still shows the logo`);
  ok(await until(p, `!document.getElementById('kc-intro')`, 4000) && Date.now() - r0 < 3500, `${N}: reduced motion is a short fade`, Date.now() - r0);
  ok(p.errors.length === 0, `${N}: no console errors`, p.errors);
}

/* ---------- 3d. Practice: 500 problems, easy to hard, judged on a hidden test too ---------- */
for (const [N, w, h, mobile] of [['practice laptop', 1366, 768, false], ['practice phone', 390, 844, true]]) {
  const p = await b.page();
  await p.viewport(w, h, { mobile, dpr: mobile ? 2 : 1 });
  if (mobile) await p.android();
  await openK(p, '/index.html?fast&reset');
  await p.ev(`(() => { const i = document.getElementById('wname'); i.value = 'Rani'; i.dispatchEvent(new Event('input')); })()`);
  await tap(p, '#go');
  await ready(p);
  await p.ev(`location.hash = '#practice'`);
  ok(await until(p, `document.querySelectorAll('#plist .pr-row').length === 500`, 15000), `${N}: the list shows 500 problems`, await p.ev(`document.querySelectorAll('#plist .pr-row').length`));
  ok(await p.ev(`[...document.querySelectorAll('#plist .pr-row .pt')].every((t) => !/^Problem \d+$/.test(t.textContent))`), `${N}: every problem has its title`);
  const order = await p.ev(`[...document.querySelectorAll('#plist .pr-row .pd')].map((d) => d.textContent)`);
  const rank = { Easy: 0, Medium: 1, Hard: 2 };
  ok(order.every((d, i) => i === 0 || rank[d] >= rank[order[i - 1]]), `${N}: the list runs from easy to hard`);
  await locked(p, N + ' list');
  await p.shot(N.replace(' ', '-') + '-list');
  await tap(p, '[data-f="d"][data-v="Hard"]');
  ok(await p.ev(`document.querySelectorAll('#plist .pr-row').length`) === 120, `${N}: the Hard filter shows 120`, await p.ev(`document.querySelectorAll('#plist .pr-row').length`));
  await tap(p, '[data-f="d"][data-v="All"]');
  await p.ev(`(() => { const q = document.getElementById('pq'); q.value = 'VLOOKUP'; q.dispatchEvent(new Event('input')); })()`);
  ok(await p.ev(`document.querySelectorAll('#plist .pr-row').length > 0 && document.querySelectorAll('#plist .pr-row').length < 500`), `${N}: search narrows the list`);
  for (const n of [1, 210, 500]) {
    await p.ev(`location.hash = '#p-${n}'`);
    ok(await until(p, `!!document.getElementById('fx') && window.__kicau.UI.prac && window.__kicau.UI.prac.n === ${n}`), `${N}: problem ${n} opens`);
    await p.ev(`window.__kicau.setFormula('=1')`);
    await p.ev(`document.getElementById('send').click()`);
    ok(await until(p, `!!document.querySelector('#pres .verdict.bad')`), `${N}: problem ${n} refuses a wrong answer`);
    await p.ev(`window.__kicau.setFormula(window.__kicau.UI.prac.solution)`);
    await p.ev(`document.getElementById('send').click()`);
    ok(await until(p, `!!document.querySelector('#pres .verdict.ok') && !!document.querySelector('#psol pre.code')`), `${N}: problem ${n} accepts its solution and shows the explanation`);
    if (n === 210) { await locked(p, N + ' problem'); await inView(p, '#send', N + ' problem'); await p.shot(N.replace(' ', '-') + '-problem'); }
  }
  ok(await p.ev(`Object.keys(window.__kicau.PR().solved).length`) === 3, `${N}: three problems saved as solved`);
  await p.ev(`location.hash = '#practice'`);
  await until(p, `!!document.getElementById('pq')`);
  ok(await p.ev(`document.getElementById('pq').value`) === 'VLOOKUP', `${N}: the list keeps your search when you come back`);
  await p.ev(`(() => { const q = document.getElementById('pq'); q.value = ''; q.dispatchEvent(new Event('input')); })()`);
  ok(await until(p, `document.querySelectorAll('#plist .pr-row.done').length === 3`), `${N}: the list ticks the solved ones`);
  ok(p.errors.length === 0, `${N}: no console errors`, p.errors);
}

/* ---------- 4. The published artifact build: the engine from its own folder ---------- */
{
  const p = await b.page();
  await p.viewport(390, 720, { mobile: true, dpr: 2 });
  await p.android();
  await openK(p, '/dist/kicau/preview.html?fast&reset');
  ok(await p.ev(`!!document.querySelector('.welcome')`), 'the artifact build opens the welcome');
  ok(await ready(p) === 'ready', 'the artifact build loads the engine');
  ok(await p.ev(`window.KENARI_ENGINE_BASE.endsWith('/dist/kicau/vendor/')`), 'it points the engine at its own folder', await p.ev(`window.KENARI_ENGINE_BASE`));
  ok(await p.ev(`performance.getEntriesByType('resource').some((r) => r.name.endsWith('/dist/kicau/vendor/formualizer_wasm_bg.wasm')) && !performance.getEntriesByType('resource').some((r) => /jsdelivr/.test(r.name))`), 'the engine comes from the bundle, not the CDN');
  await p.ev(`(() => { const i = document.getElementById('wname'); i.value = 'Rani'; i.dispatchEvent(new Event('input')); })()`);
  await tap(p, '[data-mood="b"]');
  await tap(p, '#go');
  await p.ev(`location.hash = '#chat-tika'`);
  await until(p, `!document.querySelector('.typingb') && !!document.querySelector('#fx')`);
  await p.ev(`window.__kicau.setFormula('=SUM(D5:D24)')`);
  await tap(p, '#send');
  ok(await until(p, `window.__kicau.ST.closed.includes('1.1')`), 'the artifact build passes a request');
  await p.ev(`location.hash = '#you'`); await sleep(150);
  await tap(p, '#wipe');
  ok(await p.ev(`window.__kicau.ST.closed.length`) === 1, 'one tap on start again does not clear');
  await tap(p, '#wipe');
  ok(await p.ev(`window.__kicau.ST.closed.length`) === 0, 'a second tap clears progress');
  ok(p.errors.length === 0, 'no console errors in the artifact build', p.errors);
}

console.log(`${passes} passed, ${fails} failed. Screenshots in ${OUT}`);
await b.close();
await stop();
process.exit(fails ? 1 : 0);
