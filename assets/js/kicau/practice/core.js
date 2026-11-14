/*
 * Kicau Practice: 500 problems from easy to hard, like a coding-practice site, on Rasa Kenari and Bank Kenari data.
 *
 * Problems are not stored as data. Each comes from a template (a small function) and a variant number, with a seeded
 * random generator, so problem 137 is always the same problem and the site stays small. A template file registers
 * templates with KicauPractice.add(); the app numbers every variant of every template from easiest to hardest.
 *
 * Every problem has two sheets: the one the learner sees ("base") and a hidden one ("alt", the same file after a
 * working day). An answer is Accepted only when it is right on both, like a hidden test case.
 *
 * Layout rules every template follows, because the app's grid relies on them:
 *   - inputs in rows 1 to 3: a label in column A and its value in column B;
 *   - the header row is row 4, data starts on row 5;
 *   - a one-cell answer goes in B2 or B3 (answerCell); a column answer is typed in the first data row of the first
 *     empty column and filled down (fill); a spilled answer starts at an anchor (spill).
 * Classic script in the browser (window.KicauPractice), CommonJS in Node.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('../../core/data.js'));
  else root.KicauPractice = factory(root.KenariData);
})(typeof globalThis !== 'undefined' ? globalThis : this, function (D) {
  'use strict';

  /* ---------- seeded randomness ---------- */
  function hash(str) { let h = 2166136261; for (let i = 0; i < str.length; i++) { h ^= str.charCodeAt(i); h = Math.imul(h, 16777619); } return h >>> 0; }
  function rng(seed) {
    let a = hash(String(seed));
    const next = () => { a |= 0; a = (a + 0x6D2B79F5) | 0; let t = Math.imul(a ^ (a >>> 15), 1 | a); t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t; return ((t ^ (t >>> 14)) >>> 0) / 4294967296; };
    const r = next;
    r.int = (lo, hi) => lo + Math.floor(next() * (hi - lo + 1));
    r.pick = (arr) => arr[Math.floor(next() * arr.length)];
    r.chance = (p) => next() < p;
    r.shuffle = (arr) => { const a2 = arr.slice(); for (let i = a2.length - 1; i > 0; i--) { const j = Math.floor(next() * (i + 1)); [a2[i], a2[j]] = [a2[j], a2[i]]; } return a2; };
    r.sample = (arr, n) => r.shuffle(arr).slice(0, n);
    r.round = (v, step) => Math.round(v / step) * step;
    return r;
  }

  /* ---------- the shared world ---------- */
  let DB = null;
  const db = () => (DB || (DB = D.build()));
  const LETTERS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('');
  const REGIONS = ['Jabodetabek', 'Jawa Barat', 'Jawa Tengah', 'Jawa Timur', 'Bali Nusra', 'Sumatera', 'Sulawesi'];
  const REPS = ['Rudi', 'Wulan', 'Agus', 'Sinta', 'Bima', 'Dewi', 'Fajar', 'Lina', 'Yoga', 'Tari', 'Hendra', 'Maya'];
  const WAREHOUSES = ['Gudang Bekasi', 'Gudang Bandung', 'Gudang Semarang', 'Gudang Surabaya', 'Gudang Medan'];
  const CATEGORIES = ['Noodles', 'Coffee', 'Sambal', 'Soap'];
  const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const toSerial = (iso) => D.toSerial(iso);
  const iso = (serial) => { const d = new Date(Date.UTC(1899, 11, 30) + serial * 864e5); return d.toISOString().slice(0, 10); };

  /*
   * A sheet from rows of objects. columns: [{ h: 'Header', k: 'key' or v: (row, i) => value, num: true, date: true }].
   * Returns the cells plus helpers: col('Header') the letter, R('Header') a relative range like C5:C24,
   * A('Header') the pinned range $C$5:$C$24, T(first, last) a pinned table like $A$5:$D$24, cell('Header', i) like C7.
   */
  function table(columns, rows, opt) {
    const o = opt || {};
    const first = o.first || 5, last = first + rows.length - 1;
    const startCol = o.startCol || 0;
    const cells = {}, cols = [], num = [], dateCols = [], map = {};
    columns.forEach((c, j) => {
      const L = LETTERS[startCol + j];
      cols.push(L); map[c.h] = L;
      if (c.num) num.push(L);
      if (c.date) dateCols.push(L);
      cells[L + (first - 1)] = c.h;
      rows.forEach((row, i) => {
        const v = c.v ? c.v(row, i) : row[c.k];
        if (v !== undefined && v !== null && v !== '') cells[L + (first + i)] = c.date && typeof v === 'string' ? toSerial(v) : v;
      });
    });
    const S = { cells, cols, num, dateCols, first, last, map };
    S.col = (h) => { if (!map[h]) throw new Error('No column ' + h); return map[h]; };
    S.R = (h) => S.col(h) + first + ':' + S.col(h) + last;
    S.A = (h) => '$' + S.col(h) + '$' + first + ':$' + S.col(h) + '$' + last;
    S.T = (h1, h2) => '$' + S.col(h1) + '$' + first + ':$' + S.col(h2) + '$' + last;
    S.cell = (h, i) => S.col(h) + (first + i);
    S.next = LETTERS[startCol + columns.length];
    return S;
  }
  /* Copy cells, then change some: edits is { 'C7': value or (old) => value }. */
  function edit(cells, edits) {
    const c = Object.assign({}, cells);
    for (const [ref, v] of Object.entries(edits || {})) c[ref] = typeof v === 'function' ? v(c[ref]) : v;
    return c;
  }

  /* ---------- row makers, all drawn from the made-up companies ---------- */
  const Q4 = (dateIso) => '2026-' + String(10 + ((+dateIso.slice(5, 7) - 5) % 3)) + '-' + dateIso.slice(8, 10);
  /* Sales lines: date, invoice, distributor, region, sku, product, category, cartons, price, gross, discount_pct, net. */
  function sales(r, n, opt) {
    const o = opt || {};
    const pool = db().sales.filter((s) => !o.region || s.region === o.region);
    const rows = r.sample(pool, n).map((s) => {
      const p = D.PRODUCTS.find((x) => x.sku === s.sku) || {};
      return Object.assign({}, s, { date: o.q4 === false ? s.date : Q4(s.date), price: p.carton_price, packs: p.packs_per_carton });
    });
    return o.sort === false ? rows : rows.sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : 0));
  }
  /* Loan facilities: facility, distributor, region, type, limit, outstanding, rate, start, dpd, kol, restructured. */
  function loans(r, n) {
    const base = db().facilities;
    const out = [];
    for (let i = 0; i < n; i++) {
      const f = base[i % base.length], d = db().distributors.find((x) => x.name === f.distributor) || {};
      const limit = r.pick([50, 75, 100, 150, 200]) * 1000000;
      const use = r.pick([.2, .35, .5, .62, .75, .85, .92, .98]);
      const dpd = r.pick([0, 0, 0, 5, 15, 30, 45, 60, 75, 95, 120, 150, 200]);
      out.push({ facility: 'F' + String(201 + i), distributor: i < base.length ? f.distributor : r.pick(db().distributors).name, region: d.region || r.pick(REGIONS), type: f.type || r.pick(['Working Capital', 'Investment', 'Trade Finance']),
        limit, outstanding: Math.round(limit * use / 100000) * 100000, rate: r.pick([8.5, 9, 9.25, 9.75, 10.5, 11, 12]), start: '2025-' + String(r.int(1, 12)).padStart(2, '0') + '-' + String(r.int(1, 28)).padStart(2, '0'),
        dpd, restructured: r.chance(.2) ? 'Y' : 'N' });
    }
    return out;
  }
  /* Invoices: ref, distributor, date, terms, amount, paid (a date or blank). */
  function invoices(r, n) {
    return r.sample(db().sales, n).map((s, i) => {
      const d = db().distributors.find((x) => x.name === s.distributor) || {};
      const date = Q4(s.date);
      return { ref: 'INV-' + (2100 + i * 7 + r.int(0, 6)) + '-' + (d.city || 'JKT').slice(0, 3).toUpperCase(), distributor: s.distributor, city: d.city, date, terms: d.terms_days || 30, amount: s.net, paid: r.chance(.45) ? '' : date };
    });
  }
  /* Stock: sku, product, category, warehouse, on hand, reorder point, lead time days, cost. */
  function stock(r, n) {
    const out = [];
    for (let i = 0; i < n; i++) {
      const p = D.PRODUCTS[i % D.PRODUCTS.length];
      out.push({ sku: p.sku, product: p.name, category: p.category, warehouse: r.pick(WAREHOUSES), onhand: r.int(0, 60) * 10, reorder: r.pick([80, 100, 120, 150, 200]), lead: r.pick([3, 5, 7, 10, 14]), cost: Math.round(p.carton_price * .72 / 100) * 100 });
    }
    return out;
  }
  /* Sales reps: name, area, target, actual, visits, start date. */
  function reps(r, n) {
    return r.sample(REPS, Math.min(n, REPS.length)).map((name) => {
      const target = r.pick([80, 100, 120, 150, 180, 200]) * 1000000;
      return { name, area: r.pick(REGIONS), target, actual: Math.round(target * r.pick([.6, .74, .8, .88, .95, 1, 1.04, 1.12, 1.2, 1.31]) / 100000) * 100000, visits: r.int(18, 64), start: '20' + r.int(19, 25) + '-' + String(r.int(1, 12)).padStart(2, '0') + '-' + String(r.int(1, 28)).padStart(2, '0') };
    });
  }
  /* Monthly figures: month name, sales, returns, cost, for a 12-month view. */
  function monthly(r, opt) {
    const o = opt || {};
    let base = (o.base || 900) * 1000000;
    return MONTHS.map((m, i) => { base = Math.round(base * (1 + (r() - .42) * .12) / 1000000) * 1000000; return { month: m, i: i + 1, sales: base, returns: Math.round(base * r.pick([.01, .015, .02, .03]) / 100000) * 100000, cost: Math.round(base * r.pick([.58, .61, .64, .67]) / 1000000) * 1000000 }; });
  }

  /* ---------- several workbooks ----------
   Real reports pull from more than one file. A problem can open other sheets of its own workbook (sheets) and other
   workbooks (books). Formulas name them as Excel does: Rates!$B$5 for a sheet of the same file,
   [master_produk.xlsx]Produk!$A$5:$D$16 for another file. File names stay lowercase and sheet names have no letter
   followed by digits (Q4, A1), so filling a formula down never changes them. */
  const SHEETS = Symbol.for('kenari.sheets'), MAIN = Symbol.for('kenari.main');
  const quoteName = (n) => (/^[A-Za-z_][A-Za-z0-9_]*$/.test(n) ? n : "'" + n.replace(/'/g, "''") + "'");
  /* A range helper for a sheet seen from the problem's own workbook: file null means the same workbook. */
  function ext(file, name, S) {
    const prefix = file ? (/[^A-Za-z0-9_.]/.test(file + name) ? "'[" + file + ']' + name + "'!" : '[' + file + ']' + name + '!') : quoteName(name) + '!';
    return { prefix, file, name, S, A: (h) => prefix + S.A(h), T: (h1, h2) => prefix + S.T(h1, h2), R: (h) => prefix + S.R(h), cell: (h, i) => prefix + S.cell(h, i), abs: (h, i) => prefix + '$' + S.col(h) + '$' + (S.first + i) };
  }

  /* ---------- the problem shape the app knows ---------- */
  const GENERIC = {
    typed: { title: 'That is a typed value', plain: 'A typed answer stays the same when the sheet changes. Start with = so Excel works it out.' },
    value: { title: 'Wrong answer', plain: 'Your result does not match. Check which columns and rows you used, and the exact rule in the task.' },
    frozen: { title: 'Wrong answer on the hidden test', plain: 'Right on this sheet, wrong on the hidden one (the same file a day later). Point at cells and ranges instead of typing numbers or names.' }
  };
  /*
   * Build a problem. o: { title, statement (HTML), file, sheet (from table()), base, alt (cells), answerCell, fill: true
   * (answer column = sheet.next, typed in the first data row and filled down), fillCol, spill: {cols, rows}, ansLabel,
   * solution, alts, hints, topics, exactCase, ansDate, precision, top (label/value cells already in base/alt) }.
   */
  function problem(o) {
    const S = o.sheet;
    const fillCol = o.fill ? (o.fillCol || S.next) : null;
    const cols = S.cols.slice();
    const extra = fillCol || (o.spill ? (o.spillAnchor || S.next + S.first).replace(/\d+/, '') : null);
    if (extra) { const k = LETTERS.indexOf(extra); for (let j = LETTERS.indexOf(cols[cols.length - 1]) + 1; j <= k + (o.spill ? o.spill.cols - 1 : 0); j++) cols.push(LETTERS[j]); }
    const base = Object.assign({}, o.base || S.cells), alt = Object.assign({}, o.alt || S.cells);
    if (o.ansLabel && fillCol) { base[fillCol + (S.first - 1)] = o.ansLabel; alt[fillCol + (S.first - 1)] = o.ansLabel; }
    const spill = o.spill ? { anchor: o.spillAnchor || S.next + S.first, cols: o.spill.cols, rows: o.spill.rows } : null;
    /* Other sheets and workbooks: o.sheets = [{ name, sheet, base, alt }] in the same file, o.books = [{ file, sheets: [...] }]. */
    let books;
    if ((o.sheets && o.sheets.length) || (o.books && o.books.length)) {
      const file = o.file || 'practice.xlsx', mainName = o.sheetName || 'Data';
      const meta = (sh, f) => ({ name: sh.name, cols: sh.sheet.cols, num: sh.sheet.num, dateCols: sh.sheet.dateCols, first: sh.sheet.first, last: sh.sheet.last, prefix: f ? ext(f, sh.name, sh.sheet).prefix : quoteName(sh.name) + '!' });
      const attach = (cells, v) => {
        cells[MAIN] = mainName;
        cells[SHEETS] = (o.sheets || []).map((sh) => ({ name: sh.name, cells: Object.assign({}, sh[v] || sh.sheet.cells) }))
          .concat(...(o.books || []).map((b) => b.sheets.map((sh) => ({ name: b.file + '|' + sh.name, cells: Object.assign({}, sh[v] || sh.sheet.cells) }))));
      };
      attach(base, 'base'); attach(alt, 'alt');
      books = [{ file, sheets: [{ name: mainName, cols, num: S.num, dateCols: S.dateCols, first: S.first, last: S.last, main: true, prefix: '' }].concat((o.sheets || []).map((sh) => meta(sh, null))) }]
        .concat((o.books || []).map((b) => ({ file: b.file, sheets: b.sheets.map((sh) => meta(sh, b.file)) })));
      o = Object.assign({}, o, { sheetName: mainName });
    }
    if (spill && o.ansLabel) { const a = /^([A-Z])(\d+)$/.exec(spill.anchor); base[a[1] + (+a[2] - 1)] = o.ansLabel; alt[a[1] + (+a[2] - 1)] = o.ansLabel; }
    return Object.assign({
      kind: 'formula', practice: true, file: o.file || 'practice.xlsx', cols, num: S.num, dateCols: S.dateCols, first: S.first, last: S.last,
      answerCell: fillCol ? fillCol + S.first : spill ? spill.anchor : (o.answerCell || 'B2'),
      fill: fillCol ? { col: fillCol, first: S.first, last: S.last } : null, spill,
      colRange: Object.fromEntries(cols.map((c) => [c, c + S.first + ':' + c + S.last])),
      cells: base, variantCells: { base, alt }, variants: [{ id: 'alt', label: 'the hidden test' }],
      precision: 2, traps: [], requires: [], alts: [], hints: [], msgs: GENERIC, starter: '=', chips: []
    }, o, { sheet: undefined, base: undefined, alt: undefined, sheets: undefined, msgs: GENERIC, fill: fillCol ? { col: fillCol, first: S.first, last: S.last } : null, spill, books });
  }

  /* ---------- the registry ---------- */
  const TEMPLATES = [];
  const LEVEL = (d) => (d === 'Easy' ? 0 : d === 'Medium' ? 1 : 2);
  /* add({ key, difficulty: 'Easy'|'Medium'|'Hard', level: 1 to 10 within the whole set, topics: [...], count, make(r, v, H) }) */
  function add() { for (const t of arguments) TEMPLATES.push(t); }
  let LIST = null;
  /* Titles for the list page, from practice/index.js (made by scripts/dev/practice-index.mjs), so the list shows
     500 titles without building 500 sheets. Missing titles are filled in when a problem is first built. */
  let TITLES = {};
  function setTitles(t) { TITLES = t || {}; if (LIST) LIST.forEach((m) => { if (!m.title && TITLES[m.key]) m.title = TITLES[m.key]; }); }
  /* Every problem's number, title, difficulty and topics, in order from easiest to hardest. Cheap: no sheets built. */
  function list() {
    if (LIST) return LIST;
    const all = [];
    for (const t of TEMPLATES) for (let v = 0; v < (t.count || 4); v++) all.push({ t, v, key: t.key + ':' + v });
    all.sort((a, b) => (LEVEL(a.t.difficulty) - LEVEL(b.t.difficulty)) || (a.t.level - b.t.level) || (a.t.key < b.t.key ? -1 : a.t.key > b.t.key ? 1 : 0) || (a.v - b.v));
    LIST = all.map((x, i) => {
      const meta = TITLES[x.key] || (x.t.titles ? x.t.titles[x.v] : null);
      return { n: i + 1, id: 'P' + String(i + 1).padStart(3, '0'), key: x.key, tkey: x.t.key, v: x.v, difficulty: x.t.difficulty, level: x.t.level, topics: x.t.topics || [], title: meta || null };
    });
    return LIST;
  }
  const cache = {};
  /* The full problem, sheets and all, built the first time it is opened. */
  function get(id) {
    if (cache[id]) return cache[id];
    const meta = list().find((x) => x.id === id || x.n === id);
    if (!meta) return null;
    const t = TEMPLATES.find((x) => x.key === meta.tkey);
    const p = t.make(rng(meta.key), meta.v, H);
    const out = Object.assign(p, { id: meta.id, n: meta.n, difficulty: meta.difficulty, level: meta.level, topics: t.topics || [], tkey: t.key });
    if (!meta.title) meta.title = out.title;
    cache[id] = out;
    return out;
  }
  /* Titles come from building each problem once (in Node at test time, and lazily in the browser). */
  function titles() { return list().map((m) => m.title || get(m.id).title); }

  /* The template files, in load order, and how many problems of each difficulty each one owns (500 in all). */
  const FILES = [
    { file: 't-basics.js', topic: 'Basics: arithmetic, SUM, AVERAGE, MIN, MAX, COUNT, ROUND, percentages, $ references', quota: { Easy: 80 } },
    { file: 't-logic.js', topic: 'Logic: IF, AND, OR, NOT, nested IF, IFS, SWITCH, IFERROR', quota: { Easy: 30, Medium: 40 } },
    { file: 't-agg.js', topic: 'Conditional totals: SUMIF(S), COUNTIF(S), AVERAGEIF(S), MAXIFS, MINIFS, wildcards, date ranges', quota: { Easy: 20, Medium: 40, Hard: 10 }, multi: 10 },
    { file: 't-lookup.js', topic: 'Lookups: VLOOKUP, HLOOKUP, XLOOKUP, INDEX and MATCH, two-way, approximate tiers, IFERROR', quota: { Easy: 10, Medium: 40, Hard: 20 }, multi: 32 },
    { file: 't-text.js', topic: 'Text: LEFT, RIGHT, MID, FIND, SEARCH, LEN, TRIM, UPPER, PROPER, SUBSTITUTE, TEXT, TEXTJOIN, VALUE', quota: { Easy: 20, Medium: 25, Hard: 10 } },
    { file: 't-dates.js', topic: 'Dates: DATE, YEAR, MONTH, DAY, EOMONTH, EDATE, WEEKDAY, NETWORKDAYS, WORKDAY, aging', quota: { Medium: 30, Hard: 15 } },
    { file: 't-arrays.js', topic: 'Arrays: FILTER, SORT, SORTBY, UNIQUE, SEQUENCE, LET, SUMPRODUCT, XMATCH', quota: { Medium: 20, Hard: 30 } },
    { file: 't-finance.js', topic: 'Finance: PMT, IPMT, PPMT, FV, PV, NPER, RATE, NPV, IRR, SLN, DDB, growth, CAGR', quota: { Medium: 25, Hard: 15 } },
    { file: 't-cases.js', topic: 'Real cases: several functions together, bank and FMCG', quota: { Hard: 20 }, multi: 20 }
  ];

  const H = { ext, SHEETS, MAIN, rng, table, edit, sales, loans, invoices, stock, reps, monthly, problem, toSerial, iso, Q4, LETTERS, REGIONS, REPS, WAREHOUSES, CATEGORIES, MONTHS, PRODUCTS: D.PRODUCTS, db };
  return { add, list, get, titles, setTitles, H, TEMPLATES, GENERIC, FILES };
});
