/*
 * Practice, Lookups (10 Easy, 40 Medium, 20 Hard): VLOOKUP, HLOOKUP, XLOOKUP, INDEX and MATCH, two-way lookups,
 * approximate-match tiers and IFERROR. Each template makes `count` problems; v (0 to count-1) picks the variant and
 * r is its seeded random. A side table (a price list, a tier table) sits to the right of the main list, with one
 * empty column between them, so a column answer can fill down beside the main list.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) factory(require('./core.js'));
  else factory(root.KicauPractice);
})(typeof globalThis !== 'undefined' ? globalThis : this, function (P) {
  'use strict';
  const H = P.H;
  const L = H.LETTERS;
  const DIST = () => H.db().distributors;
  const WC = () => H.db().facilities.filter((f) => f.type === 'Working Capital');
  const nums = (n) => Array.from({ length: n }, (_, i) => i);

  /*
   * Several tables on one sheet, side by side. The first is the main one: a column answer fills its rows. The grid
   * shows every table down to the longest one, and tapping a column header inserts that table's own range.
   */
  function board(tables) {
    const main = tables[0];
    const cells = Object.assign({}, ...tables.map((t) => t.cells));
    const top = Math.max(...tables.map((t) => L.indexOf(t.cols[t.cols.length - 1])));
    const cols = L.slice(0, top + 1);
    const owner = {};
    tables.forEach((t) => t.cols.forEach((c) => { owner[c] = t; }));
    const sheet = Object.assign({}, main, { cells, cols, num: [].concat(...tables.map((t) => t.num)), dateCols: [].concat(...tables.map((t) => t.dateCols)) });
    const view = { last: Math.max(...tables.map((t) => t.last)), colRange: Object.fromEntries(cols.map((c) => { const t = owner[c] || main; return [c, c + t.first + ':' + c + t.last]; })) };
    return { sheet, view, cells };
  }
  const build = (B, o) => H.problem(Object.assign({ sheet: B.sheet }, B.view, o));

  /* A table that runs across: headings in row 4 from column B, one labelled line per row from row 5. */
  function across(corner, heads, lines) {
    const cells = { A4: corner };
    heads.forEach((h, j) => { cells[L[j + 1] + '4'] = h; });
    lines.forEach((ln, i) => { cells['A' + (5 + i)] = ln.h; ln.vals.forEach((x, j) => { cells[L[j + 1] + (5 + i)] = x; }); });
    const cols = L.slice(0, heads.length + 1);
    return { cells, cols, num: cols.slice(1), dateCols: [], first: 5, last: 4 + lines.length, map: {}, next: L[heads.length + 1] };
  }

  /* Make the largest (dir 1) or smallest (dir -1) value of rows[k] unique, by moving the ties away from it. */
  function uniqueEdge(rows, k, dir, step) {
    const vals = rows.map((x) => x[k]);
    const edge = dir > 0 ? Math.max(...vals) : Math.min(...vals);
    let seen = false;
    rows.forEach((x) => { if (x[k] === edge) { if (seen) x[k] -= dir * step; seen = true; } });
    return rows.findIndex((x) => x[k] === edge);
  }

  /* Net sales per distributor from a sample of sales lines, sorted by name. */
  function totals(r, n) {
    const tot = {}, cnt = {};
    H.sales(r, n).forEach((s) => { tot[s.distributor] = (tot[s.distributor] || 0) + s.net; cnt[s.distributor] = (cnt[s.distributor] || 0) + 1; });
    return Object.keys(tot).sort().map((d) => ({ d, region: DIST().find((x) => x.name === d).region, lines: cnt[d], net: tot[d] }));
  }

  /* ======================= Easy: one exact lookup ======================= */

  P.add({
    key: 'vlookup-one-value', difficulty: 'Easy', level: 2, topics: ['VLOOKUP', 'XLOOKUP'], count: 5,
    make(r, v) {
      /* Five different lists, each asking for a number from a different column. B1 holds the code to find. */
      const V = [
        { title: 'Look up the carton price of a SKU', file: 'daftar_harga.xlsx', label: 'SKU', what: 'a SKU', key: 'SKU', want: 'Carton price (Rp)', step: 9000,
          cols: [{ h: 'SKU', k: 'sku' }, { h: 'Product', k: 'name' }, { h: 'Category', k: 'category' }, { h: 'Packs per carton', k: 'packs_per_carton', num: true }, { h: 'Carton price (Rp)', k: 'carton_price', num: true }],
          rows: () => r.shuffle(H.PRODUCTS) },
        { title: 'Look up the payment terms of a distributor', file: 'distributor.xlsx', label: 'Distributor ID', what: 'a distributor ID', key: 'ID', want: 'Terms (days)', set: 60,
          cols: [{ h: 'ID', k: 'id' }, { h: 'Distributor', k: 'name' }, { h: 'City', k: 'city' }, { h: 'Region', k: 'region' }, { h: 'Terms (days)', k: 'terms_days', num: true }],
          rows: () => r.shuffle(DIST()) },
        { title: 'Look up the outstanding balance of a facility', file: 'buku_pinjaman.xlsx', label: 'Facility', what: 'a facility code', key: 'Facility', want: 'Outstanding (Rp)', step: -1500000,
          cols: [{ h: 'Facility', k: 'facility' }, { h: 'Distributor', k: 'distributor' }, { h: 'Type', k: 'type' }, { h: 'Limit (Rp)', k: 'limit', num: true }, { h: 'Outstanding (Rp)', k: 'outstanding', num: true }],
          rows: () => H.loans(r, 14) },
        { title: 'Look up the actual sales of a sales rep', file: 'target_sales.xlsx', label: 'Sales rep', what: 'the name of a sales rep', its: 'their', key: 'Rep', want: 'Actual (Rp)', step: 2300000,
          cols: [{ h: 'Rep', k: 'name' }, { h: 'Area', k: 'area' }, { h: 'Target (Rp)', k: 'target', num: true }, { h: 'Actual (Rp)', k: 'actual', num: true }, { h: 'Visits', k: 'visits', num: true }],
          rows: () => H.reps(r, 10) },
        { title: 'Look up the cartons on hand for a SKU', file: 'stok_gudang.xlsx', label: 'SKU', what: 'a SKU', key: 'SKU', want: 'On hand', step: 30,
          cols: [{ h: 'SKU', k: 'sku' }, { h: 'Product', k: 'product' }, { h: 'Warehouse', k: 'warehouse' }, { h: 'On hand', k: 'onhand', num: true }, { h: 'Reorder point', k: 'reorder', num: true }],
          rows: () => r.shuffle(H.stock(r, 12)) }
      ][v];
      const S = H.table(V.cols, V.rows());
      const n = S.last - S.first + 1;
      const [i, j] = r.sample(nums(n), 2);
      const kc = S.col(V.key), wc = S.col(V.want);
      const at = (c, k) => S.cells[c + (S.first + k)];
      const lastCol = S.cols[S.cols.length - 1];
      const T = 'A' + S.first + ':' + lastCol + S.last, TP = '$A$' + S.first + ':$' + lastCol + '$' + S.last;
      const idx = S.cols.indexOf(wc) + 1;
      const base = Object.assign({}, S.cells, { A1: V.label, B1: at(kc, i), A2: V.want });
      /* The hidden sheet: B1 names another row, and that row's figure was corrected. */
      let nv = V.set !== undefined ? V.set : at(wc, j) + V.step;
      if (nv === at(wc, i)) nv += V.step;
      const alt = H.edit(base, { B1: at(kc, j), [wc + (S.first + j)]: nv });
      return H.problem({
        title: V.title, file: V.file, sheet: S, base, alt, answerCell: 'B2',
        statement: `<p><b>B1</b> holds ${V.what}. In <b>B2</b>, look up ${V.its || 'its'} <b>${V.want}</b> in the table <b>${T}</b>. Use an exact match.</p>`,
        solution: `=VLOOKUP(B1,${T},${idx},FALSE)`,
        alts: [`=XLOOKUP(B1,${S.R(V.key)},${S.R(V.want)})`, `=INDEX(${S.R(V.want)},MATCH(B1,${S.R(V.key)},0))`, `=VLOOKUP(B1,${TP},${idx},0)`],
        hints: [`Find the value from B1 in the first column of the table, then read across to the ${V.want} column.`,
          `VLOOKUP takes four parts: what to find, the table, the column number (the table's first column is 1), and FALSE for an exact match. ${V.want} is column ${idx}.`,
          `Shape: =VLOOKUP(B1, ${T}, ${idx}, FALSE)`],
        explain: `VLOOKUP searches the first column of ${T} for B1 and returns column ${idx} of that row. FALSE asks for an exact match, so a code that is not listed shows #N/A instead of a neighbour's value.`,
        chips: [T, S.R(V.key), S.R(V.want), String(idx), 'FALSE)'], fnKey: 'VLOOKUP('
      });
    }
  });

  P.add({
    key: 'xlookup-one-label', difficulty: 'Easy', level: 3, topics: ['XLOOKUP', 'VLOOKUP'], count: 5,
    make(r, v) {
      /* Five lists, each returning a word or a name rather than a number. */
      const V = [
        { title: 'Look up the distributor named on an invoice', file: 'faktur_q4.xlsx', label: 'Invoice', what: 'an invoice number', key: 'Invoice', want: 'Distributor',
          cols: [{ h: 'Invoice', k: 'ref' }, { h: 'Date', k: 'date', date: true }, { h: 'Distributor', k: 'distributor' }, { h: 'City', k: 'city' }, { h: 'Amount (Rp)', k: 'amount', num: true }],
          rows: () => H.invoices(r, 14) },
        { title: 'Look up the product name behind a SKU', file: 'master_produk.xlsx', label: 'SKU', what: 'a SKU code', key: 'SKU', want: 'Product',
          cols: [{ h: 'SKU', k: 'sku' }, { h: 'Product', k: 'name' }, { h: 'Category', k: 'category' }, { h: 'Carton price (Rp)', k: 'carton_price', num: true }],
          rows: () => r.shuffle(H.PRODUCTS) },
        { title: "Look up a distributor's city from its ID", file: 'distributor.xlsx', label: 'Distributor ID', what: 'a distributor ID', key: 'ID', want: 'City',
          cols: [{ h: 'ID', k: 'id' }, { h: 'Distributor', k: 'name' }, { h: 'City', k: 'city' }, { h: 'Region', k: 'region' }],
          rows: () => r.shuffle(DIST()) },
        { title: "Look up an outlet's name from its outlet ID", file: 'daftar_outlet.xlsx', label: 'Outlet ID', what: 'an outlet ID', key: 'Outlet ID', want: 'Outlet',
          cols: [{ h: 'Outlet ID', k: 'outlet_id' }, { h: 'Outlet', k: 'outlet' }, { h: 'City', k: 'city' }, { h: 'Visits per month', k: 'visits_per_month', num: true }],
          rows: () => r.sample(H.db().outlets, 14).sort((a, b) => (a.outlet_id < b.outlet_id ? -1 : 1)) },
        { title: 'Look up the sales area of a sales rep', file: 'target_sales.xlsx', label: 'Sales rep', what: 'the name of a sales rep', its: 'their', key: 'Rep', want: 'Area',
          cols: [{ h: 'Rep', k: 'name' }, { h: 'Area', k: 'area' }, { h: 'Target (Rp)', k: 'target', num: true }, { h: 'Visits', k: 'visits', num: true }],
          rows: () => H.reps(r, 10) }
      ][v];
      const S = H.table(V.cols, V.rows());
      const n = S.last - S.first + 1;
      const kc = S.col(V.key), wc = S.col(V.want);
      const at = (c, k) => S.cells[c + (S.first + k)];
      const i = r.int(1, n - 1);
      const others = nums(n).filter((k) => k !== i && at(wc, k) !== at(wc, i));
      const j = r.pick(others);
      const lastCol = S.cols[S.cols.length - 1];
      const T = 'A' + S.first + ':' + lastCol + S.last, idx = S.cols.indexOf(wc) + 1;
      const base = Object.assign({}, S.cells, { A1: V.label, B1: at(kc, i), A2: V.want });
      const alt = H.edit(base, { B1: at(kc, j) });
      return H.problem({
        title: V.title, file: V.file, sheet: S, base, alt, answerCell: 'B2',
        statement: `<p><b>B1</b> holds ${V.what}. In <b>B2</b>, use XLOOKUP to return ${V.its || 'its'} <b>${V.want}</b> from the table in rows ${S.first} to ${S.last}.</p>`,
        solution: `=XLOOKUP(B1,${S.R(V.key)},${S.R(V.want)})`,
        alts: [`=VLOOKUP(B1,${T},${idx},FALSE)`, `=INDEX(${S.R(V.want)},MATCH(B1,${S.R(V.key)},0))`, `=XLOOKUP(B1,${S.A(V.key)},${S.A(V.want)})`],
        hints: [`XLOOKUP needs three parts: what to find (B1), where to find it (the ${V.key} column) and what to bring back (the ${V.want} column).`,
          `The ${V.key} column is ${S.R(V.key)} and the ${V.want} column is ${S.R(V.want)}. Both cover the same rows.`,
          `Shape: =XLOOKUP(B1, ${S.R(V.key)}, ${S.R(V.want)})`],
        explain: `XLOOKUP finds B1 in ${S.R(V.key)} and returns the cell on the same row of ${S.R(V.want)}. It matches exactly by default, and the two ranges line up row for row.`,
        chips: [S.R(V.key), S.R(V.want), 'B1'], fnKey: 'XLOOKUP('
      });
    }
  });

  /* ======================= Medium ======================= */

  P.add({
    key: 'price-times-quantity', difficulty: 'Medium', level: 4, topics: ['XLOOKUP', 'VLOOKUP'], count: 4,
    make(r, v) {
      /* The main list in A to D (the code in C, the quantity in D), a rate list in G and H; the answer fills column E. */
      const V = [
        { title: 'Line value from a carton price list', file: 'pesanan_q4.xlsx', who: 'Tika', ans: 'Value (Rp)',
          main: [{ h: 'Date', k: 'date', date: true }, { h: 'Distributor', k: 'distributor' }, { h: 'SKU', k: 'sku' }, { h: 'Cartons', k: 'cartons', num: true }], rows: () => H.sales(r, 14),
          side: [{ h: 'SKU', k: 'sku' }, { h: 'Carton price (Rp)', k: 'carton_price', num: true }], list: () => r.shuffle(H.PRODUCTS), step: 6000, qstep: 5,
          task: 'multiply the line\'s <b>Cartons</b> by the carton price of its <b>SKU</b>' },
        { title: 'Packs shipped from a packs-per-carton list', file: 'pengiriman_q4.xlsx', who: 'Pak Anton', ans: 'Packs',
          main: [{ h: 'Date', k: 'date', date: true }, { h: 'Invoice', k: 'invoice' }, { h: 'SKU', k: 'sku' }, { h: 'Cartons', k: 'cartons', num: true }], rows: () => H.sales(r, 14),
          side: [{ h: 'SKU', k: 'sku' }, { h: 'Packs per carton', k: 'packs_per_carton', num: true }], list: () => r.shuffle(H.PRODUCTS), step: 8, qstep: 4,
          task: 'turn the line\'s <b>Cartons</b> into packs with the packs per carton of its <b>SKU</b>' },
        { title: 'Stock value from a cost-per-carton list', file: 'stok_gudang.xlsx', who: 'Bu Lestari', ans: 'Stock value (Rp)',
          main: [{ h: 'Warehouse', k: 'warehouse' }, { h: 'Product', k: 'product' }, { h: 'SKU', k: 'sku' }, { h: 'On hand', k: 'onhand', num: true }], rows: () => r.shuffle(H.stock(r, 12)),
          side: [{ h: 'SKU', k: 'sku' }, { h: 'Cost per carton (Rp)', k: 'cost', num: true }], list: () => r.shuffle(H.PRODUCTS).map((p) => ({ sku: p.sku, cost: Math.round(p.carton_price * 0.72 / 100) * 100 })), step: 2500, qstep: 20,
          task: 'multiply the cartons <b>On hand</b> by the cost per carton of its <b>SKU</b>' },
        { title: 'Monthly interest from a loan-type rate table', file: 'buku_pinjaman.xlsx', who: 'Dimas', ans: 'Interest (Rp)', round: true,
          main: [{ h: 'Facility', k: 'facility' }, { h: 'Distributor', k: 'distributor' }, { h: 'Type', k: 'type' }, { h: 'Outstanding (Rp)', k: 'outstanding', num: true }], rows: () => H.loans(r, 14),
          side: [{ h: 'Type', k: 'type' }, { h: 'Rate (% a year)', k: 'rate', num: true }],
          list: () => [{ type: 'Working Capital', rate: r.pick([10, 10.5, 11]) }, { type: 'Term Loan', rate: r.pick([11.5, 12, 12.5]) }, { type: 'Trade Finance', rate: r.pick([9, 9.5]) }], step: 0.5, qstep: 2000000,
          task: 'work out one month of interest: the <b>Outstanding</b> balance times the yearly rate of its loan <b>Type</b>, divided by 100 and by 12, rounded to whole rupiah with ROUND' }
      ][v];
      const rows = V.rows(), list = V.list();
      const S = H.table(V.main, rows), M = H.table(V.side, list, { startCol: 6 });
      const B = board([S, M]);
      const kH = V.main[2].h, qH = V.main[3].h, sk = V.side[0], sv = V.side[1];
      const key0 = rows[0][V.main[2].k];
      const base = B.cells;
      /* The hidden sheet: one rate in the list is updated and one quantity is corrected. */
      const alt = H.edit(base, {
        [M.cell(sv.h, list.findIndex((x) => x[sk.k] === key0))]: (x) => Math.round((x + V.step) * 100) / 100,
        [S.cell(qH, 5)]: (x) => x + V.qstep
      });
      const KA = M.A(sk.h), VA = M.A(sv.h), TA = M.T(sk.h, sv.h);
      const xl = 'XLOOKUP(C5,' + KA + ',' + VA + ')', vl = 'VLOOKUP(C5,' + TA + ',2,FALSE)', im = 'INDEX(' + VA + ',MATCH(C5,' + KA + ',0))';
      const wrap = (x) => (V.round ? '=ROUND(D5*' + x + '/100/12,0)' : '=D5*' + x);
      return build(B, {
        title: V.title, file: V.file, base, alt, fill: true, fillCol: 'E', ansLabel: V.ans,
        statement: `<p>${V.who} needs the <b>${V.ans}</b> of every row. Using the list in <b>G5:H${M.last}</b>, in <b>E5</b> ${V.task}. Then fill it down to <b>E${S.last}</b>.</p>`,
        solution: wrap(xl),
        alts: [wrap(vl), wrap(im), V.round ? '=ROUND(D5*' + vl + '/1200,0)' : '=' + xl + '*D5'],
        hints: [`Look up the ${kH} in C5 in the list in G and H, then work with ${qH} in D5.`,
          `Pin the list with $ so it stays put as you fill down: ${KA} and ${VA}.`,
          `Shape: ${wrap('XLOOKUP(C5, ' + KA + ', ' + VA + ')')}`],
        explain: `XLOOKUP finds the ${kH} from C5 in ${KA} and returns the ${sv.h} beside it. ${V.round ? 'Times D5, divided by 100 and by 12 and rounded, that is one month of interest.' : 'Times D5 that gives the row\'s ' + V.ans + '.'} The $ signs keep the list fixed while the formula fills down.`,
        chips: ['D5*', 'C5', KA, VA, TA, ',FALSE)'], fnKey: 'XLOOKUP('
      });
    }
  });

  P.add({
    key: 'bring-in-a-column', difficulty: 'Medium', level: 4, topics: ['XLOOKUP', 'VLOOKUP'], count: 4,
    make(r, v) {
      /* Fill a column from a side list: the key is in C, the list is in G and H, the answer fills E. */
      const V = [
        { title: 'Region for each sales line from a distributor list', file: 'penjualan_q4.xlsx', ans: 'Region',
          main: [{ h: 'Date', k: 'date', date: true }, { h: 'Invoice', k: 'invoice' }, { h: 'Distributor', k: 'distributor' }, { h: 'Net (Rp)', k: 'net', num: true }], rows: () => H.sales(r, 14),
          side: [{ h: 'Distributor', k: 'name' }, { h: 'Region', k: 'region' }], list: () => r.shuffle(DIST()), sideEdit: (x) => H.REGIONS[(H.REGIONS.indexOf(x) + 3) % H.REGIONS.length] },
        { title: 'Category for each order line from the product list', file: 'pesanan_q4.xlsx', ans: 'Category',
          main: [{ h: 'Date', k: 'date', date: true }, { h: 'Invoice', k: 'invoice' }, { h: 'SKU', k: 'sku' }, { h: 'Cartons', k: 'cartons', num: true }], rows: () => H.sales(r, 14),
          side: [{ h: 'SKU', k: 'sku' }, { h: 'Category', k: 'category' }], list: () => r.shuffle(H.PRODUCTS) },
        { title: 'Payment terms for each invoice from a distributor list', file: 'faktur_q4.xlsx', ans: 'Terms (days)',
          main: [{ h: 'Invoice', k: 'ref' }, { h: 'Date', k: 'date', date: true }, { h: 'Distributor', k: 'distributor' }, { h: 'Amount (Rp)', k: 'amount', num: true }], rows: () => H.invoices(r, 14),
          side: [{ h: 'Distributor', k: 'name' }, { h: 'Terms (days)', k: 'terms_days', num: true }], list: () => r.shuffle(DIST()), sideEdit: () => 60 },
        { title: 'Distributor name for each outlet from its ID', file: 'daftar_outlet.xlsx', ans: 'Distributor',
          main: [{ h: 'Outlet ID', k: 'outlet_id' }, { h: 'Outlet', k: 'outlet' }, { h: 'Distributor ID', k: 'distributor_id' }, { h: 'Visits per month', k: 'visits_per_month', num: true }],
          rows: () => r.sample(H.db().outlets, 14).sort((a, b) => (a.outlet_id < b.outlet_id ? -1 : 1)),
          side: [{ h: 'ID', k: 'id' }, { h: 'Distributor', k: 'name' }], list: () => r.shuffle(DIST()) }
      ][v];
      const rows = V.rows(), list = V.list();
      const S = H.table(V.main, rows), M = H.table(V.side, list, { startCol: 6 });
      const B = board([S, M]);
      const kH = V.main[2].h, kk = V.main[2].k, sk = V.side[0], sv = V.side[1];
      const valOf = (key) => list.find((x) => x[sk.k] === key)[sv.k];
      /* The hidden sheet: two lines are corrected to another key, and one entry in the list changes. */
      const e = {};
      [2, 6].forEach((i) => { const cur = valOf(rows[i][kk]); e[S.cell(kH, i)] = r.pick(list.filter((x) => x[sv.k] !== cur))[sk.k]; });
      if (V.sideEdit) { const w = list.findIndex((x) => x[sk.k] === rows[0][kk]); e[M.cell(sv.h, w)] = V.sideEdit(list[w][sv.k]); }
      const alt = H.edit(B.cells, e);
      const KA = M.A(sk.h), VA = M.A(sv.h), TA = M.T(sk.h, sv.h);
      return build(B, {
        title: V.title, file: V.file, base: B.cells, alt, fill: true, fillCol: 'E', ansLabel: V.ans,
        statement: `<p>In <b>E5</b>, bring in the <b>${sv.h}</b> for the row's <b>${kH}</b> from the list in <b>G5:H${M.last}</b>. Then fill it down to <b>E${S.last}</b>.</p>`,
        solution: `=XLOOKUP(C5,${KA},${VA})`,
        alts: [`=VLOOKUP(C5,${TA},2,FALSE)`, `=INDEX(${VA},MATCH(C5,${KA},0))`, `=VLOOKUP(C5,${TA},2,0)`],
        hints: [`Look up C5 in the ${sk.h} column of the list and return the ${sv.h} beside it.`,
          `Without $, the list slides down one row each time you fill, and lower rows miss the top of it. Pin it: ${KA} and ${VA}.`,
          `Shape: =XLOOKUP(C5, ${KA}, ${VA})`],
        explain: `XLOOKUP finds each row's ${kH} in ${KA} and returns the ${sv.h} on the same row. The key C5 is relative, so it moves down with the fill, while the $ signs keep the list where it is.`,
        chips: ['C5', KA, VA, TA, ',FALSE)'], fnKey: 'XLOOKUP('
      });
    }
  });

  P.add({
    key: 'position-with-match', difficulty: 'Medium', level: 4, topics: ['MATCH', 'XMATCH'], count: 3,
    make(r, v) {
      if (v === 0) {
        /* A ranking: the position in the list is the place. */
        const rows = totals(r, 90).sort((a, b) => b.net - a.net || (a.d < b.d ? -1 : 1));
        const S = H.table([{ h: 'Distributor', k: 'd' }, { h: 'Region', k: 'region' }, { h: 'Net sales (Rp)', k: 'net', num: true }], rows);
        const [i, j, k, l] = r.sample(nums(rows.length), 4);
        const base = Object.assign({}, S.cells, { A1: 'Distributor', B1: rows[i].d, A2: 'Place' });
        /* The hidden sheet: B1 names another distributor, and two others swapped places. */
        const alt = H.edit(base, { B1: rows[j].d, [S.cell('Distributor', k)]: rows[l].d, [S.cell('Distributor', l)]: rows[k].d, [S.cell('Region', k)]: rows[l].region, [S.cell('Region', l)]: rows[k].region });
        const R = S.R('Distributor');
        return H.problem({
          title: 'Place of a distributor in the net sales ranking', file: 'peringkat_distributor.xlsx', sheet: S, base, alt, answerCell: 'B2',
          statement: `<p>The table ranks distributors by net sales, highest first. In <b>B2</b>, return the place of the distributor named in <b>B1</b> (1 for the top row).</p>`,
          solution: `=MATCH(B1,${R},0)`, alts: [`=XMATCH(B1,${R})`, `=XMATCH(B1,${R},0)`, `=MATCH(B1,${S.A('Distributor')},0)`],
          hints: ['The list is already in rank order, so the place is how far down the list the name sits.', 'MATCH returns a position, not a value. The last part, 0, asks for an exact match.', `Shape: =MATCH(B1, ${R}, 0)`],
          explain: `MATCH(B1, ${R}, 0) counts down the list until it finds the name, and returns that position. On a ranked list, the position is the place.`,
          chips: [R, 'B1', ',0)'], fnKey: 'MATCH('
        });
      }
      if (v === 1) {
        /* A delivery route in driving order: stops after = number of stops minus this one's position. */
        const seen = new Set();
        const pool = r.shuffle(H.db().outlets).filter((o) => (seen.has(o.outlet) ? false : (seen.add(o.outlet), true))).slice(0, 12);
        const rows = pool.map((o) => ({ outlet: o.outlet, city: o.city, drop: r.int(3, 25) }));
        const S = H.table([{ h: 'Outlet', k: 'outlet' }, { h: 'City', k: 'city' }, { h: 'Drop (cartons)', k: 'drop', num: true }], rows);
        const i = r.int(1, 9);
        const j = r.pick(nums(11).filter((k) => k !== i));
        const base = Object.assign({}, S.cells, { A1: 'Outlet', B1: rows[i].outlet, A2: 'Stops after' });
        /* The hidden sheet: B1 names another stop, and the last stop was dropped, leaving its row empty. A typed 12 or a
           ROWS count no longer fits; COUNTA does. */
        const alt = H.edit(base, { B1: rows[j].outlet, [S.cell('Drop (cartons)', i)]: (x) => x + 4 });
        S.cols.forEach((c) => { delete alt[c + S.last]; });
        const R = S.R('Outlet');
        return H.problem({
          title: 'Stops left after an outlet on a delivery route', file: 'rute_kirim.xlsx', sheet: S, base, alt, answerCell: 'B2',
          statement: `<p>The truck visits the outlets in the order listed, top to bottom. In <b>B2</b>, work out how many stops come <b>after</b> the outlet in <b>B1</b>. A stop dropped from the end of the route leaves its row empty, so count the outlets listed with COUNTA rather than typing 12.</p>`,
          solution: `=COUNTA(${R})-MATCH(B1,${R},0)`, alts: [`=COUNTA(${R})-XMATCH(B1,${R})`, `=COUNTA(${S.A('Outlet')})-MATCH(B1,${S.A('Outlet')},0)`],
          hints: ['MATCH tells you which stop number the outlet is.', `The stops after it are the outlets listed minus its stop number. COUNTA(${R}) counts the outlets and skips empty rows.`, `Shape: =COUNTA(${R}) - MATCH(B1, ${R}, 0)`],
          explain: `MATCH gives the outlet's stop number. COUNTA(${R}) counts the outlets still on the route, so the difference is the stops still to come. ROWS would also count a row left empty by a dropped stop.`,
          chips: [R, 'COUNTA(', 'B1', ',0)'], fnKey: 'MATCH('
        });
      }
      /* A queue in order of arrival: applications ahead = position minus 1. */
      const ids = r.sample(nums(90), 12).map((x) => 'APP-' + (310 + x));
      const ds = r.sample(DIST(), 12);
      let day = H.toSerial('2026-10-01');
      const rows = ids.map((id, k) => ({ id, d: ds[k].name, rec: (day += r.int(0, 2)), amt: r.pick([25, 40, 50, 75, 100, 150]) * 1000000 }));
      const S = H.table([{ h: 'Application', k: 'id' }, { h: 'Distributor', k: 'd' }, { h: 'Received', k: 'rec', date: true }, { h: 'Amount (Rp)', k: 'amt', num: true }], rows);
      const i = r.int(2, 11);
      const j = r.pick(nums(12).filter((k) => k !== i));
      const base = Object.assign({}, S.cells, { A1: 'Application', B1: rows[i].id, A2: 'Ahead of it' });
      const alt = H.edit(base, { B1: rows[j].id });
      const R = S.R('Application');
      return H.problem({
        title: 'Loan applications ahead of one in the queue', file: 'antrian_kredit.xlsx', sheet: S, base, alt, answerCell: 'B2',
        statement: `<p>Bank Kenari reviews loan applications in the order received, top first. In <b>B2</b>, return how many applications are ahead of the one in <b>B1</b> (0 for the first one).</p>`,
        solution: `=MATCH(B1,${R},0)-1`, alts: [`=XMATCH(B1,${R})-1`, `=MATCH(B1,${S.A('Application')},0)-1`],
        hints: ['MATCH returns the position of B1 in the Application column.', 'Every application above it is ahead of it, so take one off its position.', `Shape: =MATCH(B1, ${R}, 0) - 1`],
        explain: `MATCH gives the application's place in the queue; the ones ahead of it are every place before that, so the answer is the place minus 1.`,
        chips: [R, 'B1', ',0)', '-1'], fnKey: 'MATCH('
      });
    }
  });

  P.add({
    key: 'lookup-to-the-left', difficulty: 'Medium', level: 5, topics: ['XLOOKUP', 'INDEX', 'MATCH'], count: 4,
    make(r, v) {
      /* The value to find sits in the last column and the answer in the first: VLOOKUP cannot look left. */
      const V = [
        { title: "Facility code from a borrower's name (lookup to the left)", file: 'buku_pinjaman.xlsx', label: 'Borrower', what: 'the name of a borrower', key: 'Distributor', want: 'Facility',
          cols: [{ h: 'Facility', k: 'facility_id' }, { h: 'Limit (Rp)', k: 'limit', num: true }, { h: 'Outstanding (Rp)', k: 'outstanding', num: true }, { h: 'Rate (%)', k: 'rate_pct', num: true }, { h: 'Distributor', k: 'distributor' }],
          rows: () => r.shuffle(WC()) },
        { title: 'SKU code from a product name (lookup to the left)', file: 'master_produk.xlsx', label: 'Product', what: 'a product name', key: 'Product', want: 'SKU',
          cols: [{ h: 'SKU', k: 'sku' }, { h: 'Category', k: 'category' }, { h: 'Packs per carton', k: 'packs_per_carton', num: true }, { h: 'Carton price (Rp)', k: 'carton_price', num: true }, { h: 'Product', k: 'name' }],
          rows: () => r.shuffle(H.PRODUCTS) },
        { title: 'Distributor ID from its home city (lookup to the left)', file: 'distributor.xlsx', label: 'City', what: 'a city', key: 'City', want: 'ID',
          cols: [{ h: 'ID', k: 'id' }, { h: 'Distributor', k: 'name' }, { h: 'Region', k: 'region' }, { h: 'Terms (days)', k: 'terms_days', num: true }, { h: 'City', k: 'city' }],
          rows: () => r.shuffle(DIST()) },
        { title: "Staff code from a sales rep's name (lookup to the left)", file: 'tim_sales.xlsx', label: 'Sales rep', what: 'the name of a sales rep', its: 'their', key: 'Rep', want: 'Code',
          cols: [{ h: 'Code', k: 'code' }, { h: 'Area', k: 'area' }, { h: 'Target (Rp)', k: 'target', num: true }, { h: 'Actual (Rp)', k: 'actual', num: true }, { h: 'Rep', k: 'name' }],
          rows: () => { const codes = r.sample(nums(80), 10).map((x) => 'SR-' + (x + 11)); return H.reps(r, 10).map((x, k) => Object.assign(x, { code: codes[k] })); } }
      ][v];
      const S = H.table(V.cols, V.rows());
      const n = S.last - S.first + 1;
      const [i, j] = r.sample(nums(n), 2);
      const kc = S.col(V.key), wc = S.col(V.want);
      const at = (c, k) => S.cells[c + (S.first + k)];
      const base = Object.assign({}, S.cells, { A1: V.label, B1: at(kc, i), A2: V.want });
      const alt = H.edit(base, { B1: at(kc, j) });
      const K = S.R(V.key), W = S.R(V.want);
      return H.problem({
        title: V.title, file: V.file, sheet: S, base, alt, answerCell: 'B2',
        statement: `<p><b>B1</b> holds ${V.what}. In <b>B2</b>, return ${V.its || 'its'} <b>${V.want}</b>. The ${V.key} column is on the right (${K}) and the ${V.want} column is on the left (${W}), so VLOOKUP cannot do it.</p>`,
        solution: `=XLOOKUP(B1,${K},${W})`,
        alts: [`=INDEX(${W},MATCH(B1,${K},0))`, `=INDEX(${W},XMATCH(B1,${K}))`, `=XLOOKUP(B1,${S.A(V.key)},${S.A(V.want)})`],
        hints: ['VLOOKUP only returns columns to the right of the one it searches. XLOOKUP, or INDEX with MATCH, can return any column.',
          `Search ${K}, return from ${W}.`,
          `Shape: =XLOOKUP(B1, ${K}, ${W}), or =INDEX(${W}, MATCH(B1, ${K}, 0))`],
        explain: `XLOOKUP takes the search column and the return column separately, so the return column can sit to the left. INDEX with MATCH does the same: MATCH finds the row, INDEX reads it from ${W}.`,
        chips: [K, W, 'B1', ',0)'], fnKey: 'XLOOKUP('
      });
    }
  });

  P.add({
    key: 'hlookup-across-months', difficulty: 'Medium', level: 5, topics: ['HLOOKUP', 'XLOOKUP'], count: 4,
    make(r, v) {
      /* Tables that run across: the headings are in row 4 and each figure is a row below. */
      if (v <= 1) {
        const m = H.monthly(r, { base: r.pick([700, 850, 1000]) });
        const S = across('Line', H.MONTHS, [{ h: 'Sales (Rp)', vals: m.map((x) => x.sales) }, { h: 'Returns (Rp)', vals: m.map((x) => x.returns) }, { h: 'Cost (Rp)', vals: m.map((x) => x.cost) }]);
        const [i, j] = r.sample(nums(12), 2);
        const T = 'B4:M7', c = (k) => L[k + 1];
        const ans = (k, d) => (v === 0 ? m[k].sales + d : m[k].sales - m[k].cost + d);
        let d = v === 0 ? 7000000 : 4000000;
        if (ans(j, d) === ans(i, 0)) d += 1000000;
        const base = Object.assign({}, S.cells, { A1: 'Month', B1: H.MONTHS[i], A2: v === 0 ? 'Sales (Rp)' : 'Gross margin (Rp)' });
        const alt = H.edit(base, { B1: H.MONTHS[j], [c(j) + (v === 0 ? 5 : 7)]: (x) => (v === 0 ? x + d : x - d) });
        if (v === 0) {
          return H.problem({
            title: 'Sales for a chosen month from a months-across table', file: 'laba_rugi_2026.xlsx', sheet: S, base, alt, answerCell: 'B2',
            statement: '<p>The months run across row 4. In <b>B2</b>, use HLOOKUP to return the <b>Sales (Rp)</b> for the month in <b>B1</b>.</p>',
            solution: `=HLOOKUP(B1,${T},2,FALSE)`, alts: ['=XLOOKUP(B1,B4:M4,B5:M5)', '=INDEX(B5:M5,MATCH(B1,B4:M4,0))', '=HLOOKUP(B1,$B$4:$M$7,2,0)'],
            hints: ['HLOOKUP is VLOOKUP turned on its side: it searches the top row and reads down.', 'In B4:M7 the month headings are row 1 of the range and Sales is row 2.', `Shape: =HLOOKUP(B1, ${T}, 2, FALSE)`],
            explain: 'HLOOKUP finds B1 in the top row of B4:M7 and returns row 2 of the range in that column, the Sales line. FALSE makes it an exact match.',
            chips: [T, 'B4:M4', 'B5:M5', ',2,FALSE)'], fnKey: 'HLOOKUP('
          });
        }
        return H.problem({
          title: 'Gross margin for a chosen month from a months-across table', file: 'laba_rugi_2026.xlsx', sheet: S, base, alt, answerCell: 'B2',
          statement: '<p>The months run across row 4. In <b>B2</b>, return the gross margin for the month in <b>B1</b>: its <b>Sales (Rp)</b> minus its <b>Cost (Rp)</b>. Use a lookup for each, so the answer follows B1.</p>',
          solution: `=HLOOKUP(B1,${T},2,FALSE)-HLOOKUP(B1,${T},4,FALSE)`,
          alts: ['=XLOOKUP(B1,B4:M4,B5:M5)-XLOOKUP(B1,B4:M4,B7:M7)', '=INDEX(B5:M5,MATCH(B1,B4:M4,0))-INDEX(B7:M7,MATCH(B1,B4:M4,0))', '=XLOOKUP(B1,B4:M4,B5:M5-B7:M7)'],
          hints: ['Two lookups on the same month: one for Sales, one for Cost.', 'In B4:M7, Sales is row 2 and Cost is row 4.', `Shape: =HLOOKUP(B1, ${T}, 2, FALSE) - HLOOKUP(B1, ${T}, 4, FALSE)`],
          explain: 'Each HLOOKUP finds the month in the top row; one reads the Sales line, the other the Cost line, and the difference is the gross margin for that month.',
          chips: [T, 'B4:M4', 'B5:M5', 'B7:M7', ',FALSE)'], fnKey: 'HLOOKUP('
        });
      }
      if (v === 2) {
        const ps = r.shuffle(H.PRODUCTS);
        const S = across('SKU', ps.map((p) => p.sku), [{ h: 'Product', vals: ps.map((p) => p.name) }, { h: 'Packs per carton', vals: ps.map((p) => p.packs_per_carton) }, { h: 'Carton price (Rp)', vals: ps.map((p) => p.carton_price) }]);
        const [i, j] = r.sample(nums(12), 2);
        const per = (k, d) => Math.round((ps[k].carton_price + d) / ps[k].packs_per_carton);
        let d = 6000;
        while (per(j, d) === per(i, 0)) d += 6000;
        const base = Object.assign({}, S.cells, { A1: 'SKU', B1: ps[i].sku, A2: 'Price per pack (Rp)' });
        const alt = H.edit(base, { B1: ps[j].sku, [L[j + 1] + '7']: (x) => x + d });
        const T = 'B4:M7';
        return H.problem({
          title: 'Price per pack from a SKUs-across price sheet', file: 'daftar_harga.xlsx', sheet: S, base, alt, answerCell: 'B2',
          statement: '<p>The SKUs run across row 4. In <b>B2</b>, work out the price of one pack of the SKU in <b>B1</b>: its <b>Carton price (Rp)</b> divided by its <b>Packs per carton</b>, rounded to whole rupiah with ROUND.</p>',
          solution: `=ROUND(HLOOKUP(B1,${T},4,FALSE)/HLOOKUP(B1,${T},3,FALSE),0)`,
          alts: ['=ROUND(XLOOKUP(B1,B4:M4,B7:M7)/XLOOKUP(B1,B4:M4,B6:M6),0)', '=ROUND(INDEX(B7:M7,MATCH(B1,B4:M4,0))/INDEX(B6:M6,MATCH(B1,B4:M4,0)),0)'],
          hints: ['Look up the carton price and the packs per carton for the same SKU, then divide.', 'In B4:M7, Packs per carton is row 3 and Carton price is row 4.', `Shape: =ROUND(HLOOKUP(B1, ${T}, 4, FALSE) / HLOOKUP(B1, ${T}, 3, FALSE), 0)`],
          explain: 'Both HLOOKUPs find the SKU among the headings in row 4 and read down its column: row 4 of B4:M7 is the carton price and row 3 the packs per carton. Dividing and rounding gives the pack price.',
          chips: [T, 'B4:M4', 'B6:M6', 'B7:M7', ',FALSE)'], fnKey: 'HLOOKUP('
        });
      }
      const regs = r.shuffle(H.REGIONS), mgr = r.sample(H.REPS, 7);
      const S = across('Region', regs, [{ h: 'Area manager', vals: mgr }, { h: 'Target (Rp)', vals: regs.map(() => r.int(30, 90) * 10000000) }, { h: 'Warehouse', vals: regs.map(() => r.pick(H.WAREHOUSES)) }]);
      const [i, j] = r.sample(nums(7), 2);
      const base = Object.assign({}, S.cells, { A1: 'Region', B1: regs[i], A2: 'Area manager' });
      const alt = H.edit(base, { B1: regs[j], [L[j + 1] + '6']: (x) => x + 50000000 });
      const T = 'B4:H7';
      return H.problem({
        title: 'Area manager of a region from a regions-across table', file: 'struktur_area.xlsx', sheet: S, base, alt, answerCell: 'B2',
        statement: '<p>The regions run across row 4. In <b>B2</b>, use HLOOKUP to return the <b>Area manager</b> of the region in <b>B1</b>.</p>',
        solution: `=HLOOKUP(B1,${T},2,FALSE)`, alts: ['=XLOOKUP(B1,B4:H4,B5:H5)', '=INDEX(B5:H5,MATCH(B1,B4:H4,0))', '=HLOOKUP(B1,$B$4:$H$7,2,0)'],
        hints: ['HLOOKUP searches the top row of a range and reads down the column it finds.', 'The Area manager line is the second row of B4:H7.', `Shape: =HLOOKUP(B1, ${T}, 2, FALSE)`],
        explain: 'HLOOKUP finds the region in row 4 and returns the value one row down in the same column, the Area manager. FALSE makes it an exact match.',
        chips: [T, 'B4:H4', 'B5:H5', ',2,FALSE)'], fnKey: 'HLOOKUP('
      });
    }
  });

  P.add({
    key: 'lookup-with-a-fallback', difficulty: 'Medium', level: 5, topics: ['XLOOKUP', 'IFERROR', 'IFNA'], count: 4,
    make(r, v) {
      /* Some keys are missing from the side list; the answer shows a fallback instead of #N/A. */
      let S, M, rows, list, e = {}, spec;
      if (v === 0) {
        const prods = r.shuffle(H.PRODUCTS), missing = prods.slice(10);
        list = prods.slice(0, 10);
        rows = H.sales(r, 14); rows[2].sku = missing[0].sku; rows[8].sku = missing[1].sku;
        S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'Distributor', k: 'distributor' }, { h: 'SKU', k: 'sku' }, { h: 'Cartons', k: 'cartons', num: true }], rows);
        M = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Carton price (Rp)', k: 'carton_price', num: true }], list, { startCol: 6 });
        const listed = (k) => list.findIndex((p) => p.sku === rows[k].sku);
        const k0 = nums(14).find((k) => listed(k) >= 0);
        e[M.cell('Carton price (Rp)', listed(k0))] = (x) => x + 5000;
        const k4 = [4, 5, 6].find((k) => listed(k) >= 0);
        e[S.cell('SKU', k4)] = missing[1].sku;
        spec = { title: 'Carton price per line, or 0 when the SKU is not listed', file: 'pesanan_q4.xlsx', key: 'C5', keyName: 'SKU', ans: 'List price (Rp)', fb: '0',
          what: 'the carton price of the line\'s <b>SKU</b>', why: 'Two new SKUs are not on the price list yet.', fbText: '<b>0</b>' };
      } else if (v === 1) {
        const ds = DIST().slice(), fac = WC();
        const drop = r.sample(nums(14), 4);
        list = r.shuffle(fac.filter((f, k) => !drop.includes(k)));
        rows = ds;
        S = H.table([{ h: 'ID', k: 'id' }, { h: 'Distributor', k: 'name' }, { h: 'City', k: 'city' }, { h: 'Region', k: 'region' }], rows);
        M = H.table([{ h: 'Distributor', k: 'distributor' }, { h: 'Limit (Rp)', k: 'limit', num: true }], list, { startCol: 6 });
        e[M.cell('Distributor', 0)] = ds[drop[0]].name;
        e[M.cell('Limit (Rp)', 1)] = (x) => x + 50000000;
        spec = { title: 'Credit limit per distributor, or No facility', file: 'limit_kredit.xlsx', key: 'B5', keyName: 'Distributor', ans: 'Limit (Rp)', fb: '"No facility"',
          what: 'the distributor\'s <b>Limit (Rp)</b> at Bank Kenari', why: 'Not every distributor has a facility.', fbText: 'the words <b>No facility</b>' };
      } else if (v === 2) {
        const outs = r.sample(H.db().outlets, 15);
        list = outs.slice(0, 12).sort((a, b) => (a.outlet_id < b.outlet_id ? -1 : 1));
        const strangers = outs.slice(12);
        rows = H.sales(r, 14).map((s, k) => ({ date: s.date, id: [3, 7, 11].includes(k) ? strangers[[3, 7, 11].indexOf(k)].outlet_id : r.pick(list).outlet_id, rep: r.pick(H.REPS), cartons: r.int(2, 20) }));
        S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'Outlet ID', k: 'id' }, { h: 'Rep', k: 'rep' }, { h: 'Cartons', k: 'cartons', num: true }], rows);
        M = H.table([{ h: 'Outlet ID', k: 'outlet_id' }, { h: 'Outlet', k: 'outlet' }], list, { startCol: 6 });
        e[S.cell('Outlet ID', 3)] = list[0].outlet_id;
        e[S.cell('Outlet ID', 0)] = strangers[1].outlet_id;
        spec = { title: 'Outlet name for each visit, or New outlet', file: 'kunjungan_outlet.xlsx', key: 'B5', keyName: 'Outlet ID', ans: 'Outlet', fb: '"New outlet"',
          what: 'the name of the visited <b>Outlet ID</b>', why: 'Reps also visit outlets that are not in the master list yet.', fbText: 'the words <b>New outlet</b>' };
      } else {
        const prods = r.shuffle(H.PRODUCTS);
        list = prods.slice(0, 4).map((p) => ({ sku: p.sku, pct: r.pick([5, 7.5, 10, 12.5, 15]) }));
        const priceOf = (sku) => H.PRODUCTS.find((p) => p.sku === sku).carton_price;
        rows = H.sales(r, 14);
        rows[1].sku = list[0].sku; rows[5].sku = list[1].sku; rows[9].sku = list[2].sku;
        rows.forEach((x) => { x.gross = x.cartons * priceOf(x.sku); });
        S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'SKU', k: 'sku' }, { h: 'Cartons', k: 'cartons', num: true }, { h: 'Gross (Rp)', k: 'gross', num: true }], rows);
        M = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Promo (%)', k: 'pct', num: true }], list, { startCol: 6 });
        e[M.cell('Promo (%)', 0)] = (x) => x + 2.5;
        const k3 = list.some((p) => p.sku === rows[3].sku) ? prods[6].sku : list[3].sku;
        e[S.cell('SKU', 3)] = k3;
        e[S.cell('Gross (Rp)', 3)] = rows[3].cartons * priceOf(k3);
        spec = { title: 'Promo discount per line, or 0 without a promo', file: 'promo_q4.xlsx', key: 'B5', keyName: 'SKU', ans: 'Discount (Rp)', fb: '0', round: true,
          what: 'the promo discount in rupiah: the line\'s <b>Gross (Rp)</b> times the <b>Promo (%)</b> of its SKU, divided by 100 and rounded to whole rupiah', why: 'Only four SKUs are on promo this month.', fbText: '<b>0</b>' };
      }
      const B = board([S, M]);
      const alt = H.edit(B.cells, e);
      const KA = M.A(M.cells[M.cols[0] + '4']), VA = M.A(M.cells[M.cols[1] + '4']), TA = '$G$5:$H$' + M.last;
      const k = spec.key, fb = spec.fb;
      const wrap = (x) => (spec.round ? '=ROUND(D5*' + x + '/100,0)' : '=' + x);
      const solution = wrap(`XLOOKUP(${k},${KA},${VA},${fb})`);
      const alts = spec.round
        ? [`=ROUND(D5*IFERROR(VLOOKUP(${k},${TA},2,FALSE),0)/100,0)`, `=IFERROR(ROUND(D5*VLOOKUP(${k},${TA},2,FALSE)/100,0),0)`, `=ROUND(D5*IFNA(INDEX(${VA},MATCH(${k},${KA},0)),0)/100,0)`]
        : [`=IFERROR(VLOOKUP(${k},${TA},2,FALSE),${fb})`, `=IFNA(INDEX(${VA},MATCH(${k},${KA},0)),${fb})`, `=IFNA(VLOOKUP(${k},${TA},2,FALSE),${fb})`];
      return build(B, {
        title: spec.title, file: spec.file, base: B.cells, alt, fill: true, fillCol: 'E', ansLabel: spec.ans,
        statement: `<p>${spec.why} Using the list in <b>G5:H${M.last}</b>, in <b>E5</b> return ${spec.what}. When the row's ${spec.keyName} (column ${k.replace('5', '')}) is not on the list, show ${spec.fbText} instead of #N/A. Fill it down to <b>E${S.last}</b>.</p>`,
        solution, alts,
        hints: ['A plain lookup gives #N/A for a key that is not on the list. XLOOKUP has a fourth part for that case: what to show when nothing is found.',
          `Pin the list with $ (${KA} and ${VA}). IFERROR(VLOOKUP(...), ${fb}) works too.`,
          `Shape: ${wrap('XLOOKUP(' + k + ', ' + KA + ', ' + VA + ', ' + fb + ')')}`],
        explain: `XLOOKUP's fourth part, if_not_found, is returned when ${k} is missing from ${KA}, so missing keys show ${fb.replace(/"/g, '')} instead of #N/A. IFERROR or IFNA around VLOOKUP gives the same result.`,
        chips: [k, KA, VA, ',' + fb + ')'], fnKey: 'XLOOKUP('
      });
    }
  });

  P.add({
    key: 'is-it-on-the-list', difficulty: 'Medium', level: 5, topics: ['MATCH', 'ISNUMBER', 'COUNTIF'], count: 3,
    make(r, v) {
      /* A one-column list in G; each row of the main list says whether its key is on it. */
      let rows, S, listed, head, yes, no, title, file, why, kH;
      if (v === 0) {
        rows = H.sales(r, 14);
        listed = r.sample(DIST(), 4).map((d) => d.name);
        rows[1].distributor = listed[0]; rows[6].distributor = listed[1]; rows[10].distributor = listed[2];
        S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'Invoice', k: 'invoice' }, { h: 'Distributor', k: 'distributor' }, { h: 'Net (Rp)', k: 'net', num: true }], rows);
        head = 'Watch list'; yes = 'Watch'; no = 'OK'; kH = 'Distributor'; title = 'Flag sales lines from watch-listed distributors'; file = 'penjualan_q4.xlsx';
        why = 'Bank Kenari keeps a watch list of distributors that pay late.';
      } else if (v === 1) {
        rows = r.shuffle(H.stock(r, 12));
        listed = r.sample(rows, 3).map((x) => x.sku);
        S = H.table([{ h: 'Warehouse', k: 'warehouse' }, { h: 'Product', k: 'product' }, { h: 'SKU', k: 'sku' }, { h: 'On hand', k: 'onhand', num: true }], rows);
        head = 'Discontinued'; yes = 'Discontinued'; no = 'Active'; kH = 'SKU'; title = 'Mark discontinued SKUs in a stock list'; file = 'stok_gudang.xlsx';
        why = 'Rasa Kenari is dropping the SKUs listed in column G.';
      } else {
        rows = H.invoices(r, 14);
        listed = r.sample(rows, 5).map((x) => x.ref).concat(['INV-' + r.int(2300, 2399) + '-MDN']);
        S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'Distributor', k: 'distributor' }, { h: 'Invoice', k: 'ref' }, { h: 'Amount (Rp)', k: 'amount', num: true }], rows);
        head = 'Paid (bank statement)'; yes = 'Paid'; no = 'Open'; kH = 'Invoice'; title = 'Mark invoices Paid or Open from a bank statement list'; file = 'faktur_q4.xlsx';
        why = 'Column G lists the invoice numbers found on this week\'s bank statement.';
      }
      const M = H.table([{ h: head, v: (x) => x }], listed, { startCol: 6 });
      const B = board([S, M]);
      const keys = rows.map((x) => S.cells[S.cell(kH, rows.indexOf(x))]);
      const fresh = keys.find((k) => !listed.includes(k));
      const alt = H.edit(B.cells, { [M.cell(head, 0)]: fresh });
      const G = M.A(head);
      return build(B, {
        title, file, base: B.cells, alt, fill: true, fillCol: 'E', ansLabel: 'Status',
        statement: `<p>${why} In <b>E5</b>, show <b>${yes}</b> when the row's <b>${kH}</b> appears in <b>G5:G${M.last}</b>, and <b>${no}</b> when it does not. Fill it down to <b>E${S.last}</b>.</p>`,
        solution: `=IF(ISNUMBER(MATCH(C5,${G},0)),"${yes}","${no}")`,
        alts: [`=IF(COUNTIF(${G},C5)>0,"${yes}","${no}")`, `=IF(ISNA(MATCH(C5,${G},0)),"${no}","${yes}")`, `=IF(ISNUMBER(XMATCH(C5,${G})),"${yes}","${no}")`],
        hints: ['MATCH returns a position when it finds C5 on the list, and #N/A when it does not.', 'ISNUMBER turns that into TRUE or FALSE, which IF can use. COUNTIF(list, C5)>0 works too.', `Shape: =IF(ISNUMBER(MATCH(C5, ${G}, 0)), "${yes}", "${no}")`],
        explain: `MATCH finds C5 in the pinned list ${G} or returns #N/A. ISNUMBER turns the result into TRUE or FALSE, and IF writes ${yes} or ${no}.`,
        chips: ['C5', G, ',0)', '"' + yes + '"', '"' + no + '"'], fnKey: 'MATCH('
      });
    }
  });

  P.add({
    key: 'two-lookups-in-one-formula', difficulty: 'Medium', level: 6, topics: ['XLOOKUP', 'VLOOKUP'], count: 4,
    make(r, v) {
      if (v === 0) {
        const rows = H.loans(r, 12);
        const S = H.table([{ h: 'Facility', k: 'facility' }, { h: 'Distributor', k: 'distributor' }, { h: 'Limit (Rp)', k: 'limit', num: true }, { h: 'Outstanding (Rp)', k: 'outstanding', num: true }], rows);
        const [i, j] = r.sample(nums(12), 2);
        const head = (k) => rows[k].limit - rows[k].outstanding;
        let d = 1200000;
        while (head(j) + d === head(i)) d += 300000;
        const base = Object.assign({}, S.cells, { A1: 'Facility', B1: rows[i].facility, A2: 'Headroom (Rp)' });
        const alt = H.edit(base, { B1: rows[j].facility, [S.cell('Outstanding (Rp)', j)]: (x) => x - d });
        return H.problem({
          title: 'Undrawn headroom on one facility (two lookups)', file: 'buku_pinjaman.xlsx', sheet: S, base, alt, answerCell: 'B2',
          statement: '<p>In <b>B2</b>, work out the headroom of the facility in <b>B1</b>: its <b>Limit (Rp)</b> minus its <b>Outstanding (Rp)</b>. Look both up from the table.</p>',
          solution: '=XLOOKUP(B1,A5:A16,C5:C16)-XLOOKUP(B1,A5:A16,D5:D16)',
          alts: ['=VLOOKUP(B1,A5:D16,3,FALSE)-VLOOKUP(B1,A5:D16,4,FALSE)', '=INDEX(C5:C16,MATCH(B1,A5:A16,0))-INDEX(D5:D16,MATCH(B1,A5:A16,0))', '=XLOOKUP(B1,A5:A16,C5:C16-D5:D16)'],
          hints: ['One lookup returns the limit and another the outstanding balance for the same facility.', 'Both search A5:A16; one returns from C5:C16, the other from D5:D16.', 'Shape: =XLOOKUP(B1, A5:A16, C5:C16) - XLOOKUP(B1, A5:A16, D5:D16)'],
          explain: 'Each XLOOKUP finds the facility in column A and returns one figure from its row. Limit minus outstanding is what the borrower can still draw.',
          chips: ['A5:A16', 'C5:C16', 'D5:D16', 'B1'], fnKey: 'XLOOKUP('
        });
      }
      if (v === 1) {
        const ps = r.shuffle(H.PRODUCTS);
        const S = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Product', k: 'name' }, { h: 'Category', k: 'category' }, { h: 'Carton price (Rp)', k: 'carton_price', num: true }], ps);
        const pr = (k) => ps[k].carton_price;
        const pair = () => { let a, b; do { [a, b] = r.sample(nums(12), 2); } while (pr(a) === pr(b)); return [a, b]; };
        const [i, j] = pair();
        let k, l;
        do { [k, l] = pair(); } while (k === i && l === j);
        let d = 4000;
        while (pr(k) + d - pr(l) === pr(i) - pr(j)) d += 2000;
        const base = Object.assign({}, S.cells, { A1: 'SKU', B1: ps[i].sku, A2: 'Compared with', B2: ps[j].sku, A3: 'Price gap (Rp)' });
        const alt = H.edit(base, { B1: ps[k].sku, B2: ps[l].sku, [S.cell('Carton price (Rp)', k)]: (x) => x + d });
        return H.problem({
          title: 'Price gap between two SKUs (two lookups)', file: 'daftar_harga.xlsx', sheet: S, base, alt, answerCell: 'B3',
          statement: '<p>In <b>B3</b>, work out the carton price of the SKU in <b>B1</b> minus the carton price of the SKU in <b>B2</b>. The answer can be negative.</p>',
          solution: '=XLOOKUP(B1,A5:A16,D5:D16)-XLOOKUP(B2,A5:A16,D5:D16)',
          alts: ['=VLOOKUP(B1,A5:D16,4,FALSE)-VLOOKUP(B2,A5:D16,4,FALSE)', '=INDEX(D5:D16,MATCH(B1,A5:A16,0))-INDEX(D5:D16,MATCH(B2,A5:A16,0))'],
          hints: ['Look up each SKU separately, then subtract.', 'Both lookups search A5:A16 and return from D5:D16; only the value to find changes.', 'Shape: =XLOOKUP(B1, A5:A16, D5:D16) - XLOOKUP(B2, A5:A16, D5:D16)'],
          explain: 'The first XLOOKUP returns the price for B1 and the second the price for B2, from the same columns. Their difference is the gap, negative when B1 is the cheaper one.',
          chips: ['A5:A16', 'D5:D16', 'B1', 'B2'], fnKey: 'XLOOKUP('
        });
      }
      if (v === 2) {
        const rows = H.reps(r, 10);
        const S = H.table([{ h: 'Rep', k: 'name' }, { h: 'Area', k: 'area' }, { h: 'Target (Rp)', k: 'target', num: true }, { h: 'Actual (Rp)', k: 'actual', num: true }], rows);
        const [i, j] = r.sample(nums(10), 2);
        const ach = (k, d) => (rows[k].actual + d) / rows[k].target;
        let d = 3100000;
        while (Math.abs(ach(j, d) - ach(i, 0)) < 1e-4) d += 1000000;
        const base = Object.assign({}, S.cells, { A1: 'Sales rep', B1: rows[i].name, A2: 'Achievement' });
        const alt = H.edit(base, { B1: rows[j].name, [S.cell('Actual (Rp)', j)]: (x) => x + d });
        return H.problem({
          title: 'Target achievement of one rep (two lookups)', file: 'target_sales.xlsx', sheet: S, base, alt, answerCell: 'B2', precision: 4,
          statement: '<p>In <b>B2</b>, work out the target achievement of the rep in <b>B1</b>: <b>Actual (Rp)</b> divided by <b>Target (Rp)</b>, as a decimal (0.95 means 95%).</p>',
          solution: '=XLOOKUP(B1,A5:A14,D5:D14)/XLOOKUP(B1,A5:A14,C5:C14)',
          alts: ['=VLOOKUP(B1,A5:D14,4,FALSE)/VLOOKUP(B1,A5:D14,3,FALSE)', '=INDEX(D5:D14,MATCH(B1,A5:A14,0))/INDEX(C5:C14,MATCH(B1,A5:A14,0))', '=XLOOKUP(B1,A5:A14,D5:D14/C5:C14)'],
          hints: ['Look up the actual and the target for the same rep, then divide.', 'Actual is D5:D14 and Target is C5:C14. Actual goes on top.', 'Shape: =XLOOKUP(B1, A5:A14, D5:D14) / XLOOKUP(B1, A5:A14, C5:C14)'],
          explain: 'Both lookups find the rep in column A; one returns the actual, the other the target. Actual over target is the share of the target reached.',
          chips: ['A5:A14', 'C5:C14', 'D5:D14', 'B1'], fnKey: 'XLOOKUP('
        });
      }
      const rows = totals(r, 70);
      const S = H.table([{ h: 'Distributor', k: 'd' }, { h: 'Region', k: 'region' }, { h: 'Net sales (Rp)', k: 'net', num: true }], rows);
      const n = rows.length;
      const [i, j, k] = r.sample(nums(n), 3);
      const base = Object.assign({}, S.cells, { A1: 'Distributor', B1: rows[i].d, A2: 'Share of total' });
      const alt = H.edit(base, { B1: rows[j].d, [S.cell('Net sales (Rp)', k)]: (x) => x + 9000000 });
      const R = S.R('Net sales (Rp)'), K = S.R('Distributor');
      return H.problem({
        title: "One distributor's share of total net sales", file: 'penjualan_q4.xlsx', sheet: S, base, alt, answerCell: 'B2', precision: 4,
        statement: `<p>In <b>B2</b>, work out the share of all net sales that belongs to the distributor in <b>B1</b>: its <b>Net sales (Rp)</b> divided by the total of ${R}, as a decimal.</p>`,
        solution: `=XLOOKUP(B1,${K},${R})/SUM(${R})`,
        alts: [`=VLOOKUP(B1,A5:C${S.last},3,FALSE)/SUM(${R})`, `=INDEX(${R},MATCH(B1,${K},0))/SUM(${R})`, `=SUMIF(${K},B1,${R})/SUM(${R})`],
        hints: ['The top of the fraction is a lookup; the bottom is a total.', `Look up B1 in ${K} and return from ${R}, then divide by SUM(${R}).`, `Shape: =XLOOKUP(B1, ${K}, ${R}) / SUM(${R})`],
        explain: `XLOOKUP returns the distributor's own net sales and SUM(${R}) the total for everyone; the ratio is its share. Both follow the sheet when a figure changes.`,
        chips: [K, R, 'SUM(', 'B1'], fnKey: 'XLOOKUP('
      });
    }
  });

  P.add({
    key: 'name-of-the-top-row', difficulty: 'Medium', level: 6, topics: ['INDEX', 'MATCH', 'MAX', 'MIN'], count: 3,
    make(r, v) {
      /* Find the extreme with MAX or MIN, then return the name on that row. The extreme is unique on both sheets. */
      let rows, S, keyH, valH, dir, label, title, file, task, alt;
      if (v === 0) {
        rows = totals(r, 70);
        S = null; keyH = 'Distributor'; valH = 'Net sales (Rp)'; dir = 1; label = 'Top distributor'; file = 'penjualan_q4.xlsx';
        title = 'Name the distributor with the top net sales (INDEX and MATCH)'; task = 'the distributor with the <b>highest</b> Net sales (Rp)';
        const t = uniqueEdge(rows, 'net', 1, 100000);
        S = H.table([{ h: 'Distributor', k: 'd' }, { h: 'Region', k: 'region' }, { h: 'Lines', k: 'lines', num: true }, { h: 'Net sales (Rp)', k: 'net', num: true }], rows);
        const k = r.pick(nums(rows.length).filter((x) => x !== t));
        alt = { [S.cell(valH, k)]: rows[t].net + r.int(2, 9) * 100000 };
      } else if (v === 1) {
        rows = r.shuffle(H.stock(r, 12)).map((x) => Object.assign(x, { onhand: r.int(2, 60) * 10 }));
        const t = uniqueEdge(rows, 'onhand', -1, 10);
        keyH = 'Product'; valH = 'On hand'; dir = -1; label = 'Lowest stock'; file = 'stok_gudang.xlsx';
        title = 'Name the product with the fewest cartons on hand (INDEX and MATCH)'; task = 'the product with the <b>lowest</b> On hand';
        S = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Product', k: 'product' }, { h: 'Warehouse', k: 'warehouse' }, { h: 'On hand', k: 'onhand', num: true }], rows);
        const k = r.pick(nums(rows.length).filter((x) => x !== t));
        alt = { [S.cell(valH, k)]: rows[t].onhand - 10 };
      } else {
        rows = H.monthly(r, { base: r.pick([700, 900]) });
        const t = uniqueEdge(rows, 'returns', 1, 100000);
        keyH = 'Month'; valH = 'Returns (Rp)'; dir = 1; label = 'Worst month'; file = 'retur_2026.xlsx';
        title = 'Name the month with the highest returns (INDEX and MATCH)'; task = 'the month with the <b>highest</b> Returns (Rp)';
        S = H.table([{ h: 'Month', k: 'month' }, { h: 'Sales (Rp)', k: 'sales', num: true }, { h: 'Returns (Rp)', k: 'returns', num: true }], rows);
        const k = r.pick(nums(12).filter((x) => x !== t));
        alt = { [S.cell(valH, k)]: rows[t].returns + 300000 };
      }
      const base = Object.assign({}, S.cells, { A1: label });
      const K = S.R(keyH), V = S.R(valH), F = dir > 0 ? 'MAX' : 'MIN';
      return H.problem({
        title, file, sheet: S, base, alt: H.edit(base, alt), answerCell: 'B1',
        statement: `<p>In <b>B1</b>, return the name of ${task}. Find the ${dir > 0 ? 'largest' : 'smallest'} value with ${F}, then the name on its row.</p>`,
        solution: `=INDEX(${K},MATCH(${F}(${V}),${V},0))`,
        alts: [`=XLOOKUP(${F}(${V}),${V},${K})`, `=INDEX(${K},XMATCH(${F}(${V}),${V}))`],
        hints: [`${F}(${V}) gives the value, but you need the name beside it.`, `MATCH(${F}(${V}), ${V}, 0) gives the row number of that value inside the range; INDEX reads the name from the same row of ${K}.`, `Shape: =INDEX(${K}, MATCH(${F}(${V}), ${V}, 0))`],
        explain: `${F} finds the ${dir > 0 ? 'top' : 'bottom'} value, MATCH finds where it sits in ${V}, and INDEX returns the name at that position in ${K}. When a figure changes, the answer follows.`,
        chips: [K, V, F + '(', ',0)'], fnKey: 'INDEX('
      });
    }
  });

  P.add({
    key: 'check-against-the-master', difficulty: 'Medium', level: 7, topics: ['IF', 'XLOOKUP', 'VLOOKUP'], count: 4,
    make(r, v) {
      /* Compare each row with a figure looked up from a master list, and write one of two words. */
      if (v === 0) {
        const rows = H.sales(r, 14);
        const list = r.shuffle(H.PRODUCTS).map((p) => ({ sku: p.sku, oh: r.int(2, 7) * 10 }));
        const S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'Distributor', k: 'distributor' }, { h: 'SKU', k: 'sku' }, { h: 'Cartons', k: 'cartons', num: true }], rows);
        const M = H.table([{ h: 'SKU', k: 'sku' }, { h: 'On hand', k: 'oh', num: true }], list, { startCol: 6 });
        const B = board([S, M]);
        const w = list.findIndex((x) => x.sku === rows[0].sku), c0 = rows[0].cartons;
        const alt = H.edit(B.cells, { [M.cell('On hand', w)]: c0 > list[w].oh ? c0 : c0 - 1 });
        const KA = M.A('SKU'), VA = M.A('On hand');
        return build(B, {
          title: 'Mark order lines that exceed stock on hand', file: 'pesanan_q4.xlsx', base: B.cells, alt, fill: true, fillCol: 'E', ansLabel: 'Stock check',
          statement: `<p>In <b>E5</b>, show <b>Short</b> when the line's <b>Cartons</b> are more than the <b>On hand</b> for its SKU in <b>G5:H16</b>, and <b>OK</b> otherwise (equal is OK). Fill it down to <b>E${S.last}</b>.</p>`,
          solution: `=IF(D5>XLOOKUP(C5,${KA},${VA}),"Short","OK")`,
          alts: [`=IF(D5<=VLOOKUP(C5,$G$5:$H$16,2,FALSE),"OK","Short")`, `=IF(INDEX(${VA},MATCH(C5,${KA},0))<D5,"Short","OK")`],
          hints: ['First look up the On hand for the SKU in C5, then compare it with D5 inside IF.', `Pin the stock list: ${KA} and ${VA}. "More than" is >.`, `Shape: =IF(D5 > XLOOKUP(C5, ${KA}, ${VA}), "Short", "OK")`],
          explain: 'XLOOKUP returns the stock for the line\'s SKU and IF compares the order with it. A line equal to the stock is still OK, because only more than the stock is short.',
          chips: ['D5', 'C5', KA, VA, '"Short"', '"OK"'], fnKey: 'XLOOKUP('
        });
      }
      if (v === 1) {
        const rows = H.sales(r, 14).map((s) => ({ invoice: s.invoice, sku: s.sku, cartons: s.cartons, charged: s.price }));
        r.sample(nums(14), 3).forEach((k) => { rows[k].charged += r.pick([-5000, -3000, 2000, 4000]); });
        const list = r.shuffle(H.PRODUCTS);
        const S = H.table([{ h: 'Invoice', k: 'invoice' }, { h: 'SKU', k: 'sku' }, { h: 'Cartons', k: 'cartons', num: true }, { h: 'Price charged (Rp)', k: 'charged', num: true }], rows);
        const M = H.table([{ h: 'SKU', k: 'sku' }, { h: 'List price (Rp)', k: 'carton_price', num: true }], list, { startCol: 6 });
        const B = board([S, M]);
        const lp = H.PRODUCTS.find((p) => p.sku === rows[0].sku).carton_price;
        const alt = H.edit(B.cells, { [S.cell('Price charged (Rp)', 0)]: rows[0].charged === lp ? lp + 1000 : lp });
        const KA = M.A('SKU'), VA = M.A('List price (Rp)');
        return build(B, {
          title: 'Flag lines charged off the list price', file: 'audit_harga.xlsx', base: B.cells, alt, fill: true, fillCol: 'E', ansLabel: 'Price check',
          statement: `<p>In <b>E5</b>, show <b>Check</b> when the <b>Price charged (Rp)</b> is different from the <b>List price (Rp)</b> of the line's SKU in <b>G5:H16</b>, and <b>OK</b> when they are the same. Fill it down to <b>E${S.last}</b>.</p>`,
          solution: `=IF(D5<>XLOOKUP(B5,${KA},${VA}),"Check","OK")`,
          alts: [`=IF(D5=VLOOKUP(B5,$G$5:$H$16,2,FALSE),"OK","Check")`, `=IF(D5=INDEX(${VA},MATCH(B5,${KA},0)),"OK","Check")`],
          hints: ['Look up the list price for the SKU in B5, then compare it with the price in D5.', '<> means "is not equal to". Pin the price list with $.', `Shape: =IF(D5 <> XLOOKUP(B5, ${KA}, ${VA}), "Check", "OK")`],
          explain: 'XLOOKUP brings the list price for the line\'s SKU and IF compares it with what was charged. Any difference, up or down, is flagged.',
          chips: ['D5', 'B5', KA, VA, '"Check"', '"OK"'], fnKey: 'XLOOKUP('
        });
      }
      if (v === 2) {
        const rows = H.reps(r, 10);
        const list = r.shuffle(rows).map((x) => ({ name: x.name, target: x.target }));
        const S = H.table([{ h: 'Rep', k: 'name' }, { h: 'Area', k: 'area' }, { h: 'Actual (Rp)', k: 'actual', num: true }], rows);
        const M = H.table([{ h: 'Rep', k: 'name' }, { h: 'Target (Rp)', k: 'target', num: true }], list, { startCol: 5 });
        const B = board([S, M]);
        const w = list.findIndex((x) => x.name === rows[0].name);
        const alt = H.edit(B.cells, { [M.cell('Target (Rp)', w)]: rows[0].actual >= rows[0].target ? rows[0].actual + 100000 : rows[0].actual });
        const KA = M.A('Rep'), VA = M.A('Target (Rp)');
        return build(B, {
          title: 'Hit or miss against a separate target table', file: 'target_sales.xlsx', base: B.cells, alt, fill: true, fillCol: 'D', ansLabel: 'Result',
          statement: `<p>The targets sit in their own table, <b>F5:G14</b>, in a different order. In <b>D5</b>, show <b>Hit</b> when the rep's <b>Actual (Rp)</b> is at least their target, and <b>Miss</b> when it is below. Fill it down to <b>D${S.last}</b>.</p>`,
          solution: `=IF(C5>=XLOOKUP(A5,${KA},${VA}),"Hit","Miss")`,
          alts: [`=IF(C5<VLOOKUP(A5,$F$5:$G$14,2,FALSE),"Miss","Hit")`, `=IF(C5>=INDEX(${VA},MATCH(A5,${KA},0)),"Hit","Miss")`],
          hints: ['The two tables are in different orders, so look the target up by name rather than reading the same row.', `"At least" is >=. Pin the target table: ${KA} and ${VA}.`, `Shape: =IF(C5 >= XLOOKUP(A5, ${KA}, ${VA}), "Hit", "Miss")`],
          explain: 'XLOOKUP finds each rep\'s own target by name, wherever it sits in the target table. IF then compares it with the actual; reaching the target exactly counts as a hit.',
          chips: ['A5', 'C5', KA, VA, '"Hit"', '"Miss"'], fnKey: 'XLOOKUP('
        });
      }
      const fac = r.shuffle(WC());
      const ds = r.sample(fac, 12);
      const rows = ds.map((f, k) => ({ req: 'REQ-' + (41 + k * 3 + r.int(0, 2)), d: f.distributor, amt: Math.max(1000000, f.limit - f.outstanding + r.pick([-6, -3, -1, 1, 2, 5]) * 1000000) }));
      /* One request is exactly the headroom, so "at most" against "less than" shows on the visible sheet. */
      rows[3].amt = ds[3].limit - ds[3].outstanding;
      const S = H.table([{ h: 'Request', k: 'req' }, { h: 'Distributor', k: 'd' }, { h: 'Amount (Rp)', k: 'amt', num: true }], rows);
      const M = H.table([{ h: 'Distributor', k: 'distributor' }, { h: 'Limit (Rp)', k: 'limit', num: true }, { h: 'Outstanding (Rp)', k: 'outstanding', num: true }], fac, { startCol: 5 });
      const B = board([S, M]);
      const f0 = fac.findIndex((x) => x.distributor === rows[0].d), head0 = fac[f0].limit - fac[f0].outstanding;
      /* The hidden sheet flips row 1: a request within its headroom gets 100,000 less headroom than it asks for, and one
         over it is cut to exactly the headroom. */
      const alt = H.edit(B.cells, rows[0].amt <= head0 ? { [M.cell('Outstanding (Rp)', f0)]: fac[f0].limit - rows[0].amt + 100000 } : { [S.cell('Amount (Rp)', 0)]: head0 });
      const KA = M.A('Distributor'), LA = M.A('Limit (Rp)'), OA = M.A('Outstanding (Rp)');
      return build(B, {
        title: 'Check drawdown requests against facility headroom', file: 'permintaan_pencairan.xlsx', base: B.cells, alt, fill: true, fillCol: 'D', ansLabel: 'Decision',
        statement: `<p>Each request asks to draw more money from a facility. In <b>D5</b>, show <b>Within</b> when the request's <b>Amount (Rp)</b> is at most the distributor's headroom (Limit minus Outstanding, from <b>F5:H18</b>), and <b>Over</b> when it is more. Fill it down to <b>D${S.last}</b>.</p>`,
        solution: `=IF(C5<=XLOOKUP(B5,${KA},${LA})-XLOOKUP(B5,${KA},${OA}),"Within","Over")`,
        alts: [`=IF(C5>VLOOKUP(B5,$F$5:$H$18,2,FALSE)-VLOOKUP(B5,$F$5:$H$18,3,FALSE),"Over","Within")`, `=IF(C5<=INDEX(${LA},MATCH(B5,${KA},0))-INDEX(${OA},MATCH(B5,${KA},0)),"Within","Over")`, `=IF(C5<=XLOOKUP(B5,${KA},${LA}-${OA}),"Within","Over")`],
        hints: ['Headroom needs two lookups for the same distributor: the limit and the outstanding balance.', `"At most" is <=. Pin the facility table: ${KA}, ${LA} and ${OA}.`, `Shape: =IF(C5 <= XLOOKUP(B5, ${KA}, ${LA}) - XLOOKUP(B5, ${KA}, ${OA}), "Within", "Over")`],
        explain: 'The two XLOOKUPs give the distributor\'s limit and outstanding balance; their difference is the headroom. IF compares the request with it, and a request equal to the headroom still fits.',
        chips: ['B5', 'C5', KA, LA, OA, '"Within"', '"Over"'], fnKey: 'XLOOKUP('
      });
    }
  });

  P.add({
    key: 'partial-name-lookup', difficulty: 'Medium', level: 7, topics: ['XLOOKUP', 'VLOOKUP', 'wildcards'], count: 3,
    make(r, v) {
      /* Find a row from part of its text with the * wildcard: contains, starts with, ends with. */
      if (v === 0) {
        const rows = r.shuffle(WC());
        const S = H.table([{ h: 'Facility', k: 'facility_id' }, { h: 'Distributor', k: 'distributor' }, { h: 'Limit (Rp)', k: 'limit', num: true }, { h: 'Outstanding (Rp)', k: 'outstanding', num: true }], rows);
        const [i, j] = r.sample(nums(rows.length), 2);
        const word = (k) => rows[k].distributor.split(' ')[1];
        let d = 700000;
        while (rows[j].outstanding - d === rows[i].outstanding) d += 100000;
        const base = Object.assign({}, S.cells, { A1: 'Name contains', B1: word(i), A2: 'Outstanding (Rp)' });
        const alt = H.edit(base, { B1: word(j), [S.cell('Outstanding (Rp)', j)]: (x) => x - d });
        return H.problem({
          title: 'Outstanding for a borrower found by part of its name', file: 'buku_pinjaman.xlsx', sheet: S, base, alt, answerCell: 'B2',
          statement: '<p>Citra only remembers one word of the borrower\'s name, in <b>B1</b>. In <b>B2</b>, return the <b>Outstanding (Rp)</b> of the distributor whose name <b>contains</b> that word.</p>',
          solution: '=XLOOKUP("*"&B1&"*",B5:B18,D5:D18,,2)',
          alts: ['=VLOOKUP("*"&B1&"*",B5:D18,3,FALSE)', '=INDEX(D5:D18,MATCH("*"&B1&"*",B5:B18,0))'],
          hints: ['A * stands for any text, so "*"&B1&"*" matches any name with B1 somewhere inside it.', 'XLOOKUP only reads * as a wildcard when its fifth part, match_mode, is 2. VLOOKUP with FALSE and MATCH with 0 read it by default.', 'Shape: =XLOOKUP("*"&B1&"*", B5:B18, D5:D18, , 2)'],
          explain: 'Joining * on both sides of B1 builds a pattern that matches any name containing the word. With match_mode 2, XLOOKUP treats it as a wildcard pattern and returns the outstanding from that row.',
          chips: ['"*"&B1&"*"', 'B5:B18', 'D5:D18', ',,2)'], fnKey: 'XLOOKUP('
        });
      }
      if (v === 1) {
        const rows = r.shuffle(H.stock(r, 12));
        const S = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Product', k: 'product' }, { h: 'Warehouse', k: 'warehouse' }, { h: 'On hand', k: 'onhand', num: true }], rows);
        const brands = r.shuffle(['MLK', 'KSJ', 'SNY', 'SRW']);
        const first = (b) => rows.findIndex((x) => x.sku.startsWith(b));
        const p = brands[0], q = brands[1];
        let d = 40;
        while (rows[first(q)].onhand + d === rows[first(p)].onhand) d += 10;
        const base = Object.assign({}, S.cells, { A1: 'Brand code', B1: p, A2: 'On hand' });
        const alt = H.edit(base, { B1: q, [S.cell('On hand', first(q))]: (x) => x + d });
        return H.problem({
          title: 'On hand for the first SKU starting with a brand code', file: 'stok_gudang.xlsx', sheet: S, base, alt, answerCell: 'B2',
          statement: '<p>Every SKU starts with a brand code, such as MLK for Mi Lekas. In <b>B2</b>, return the <b>On hand</b> of the first line, from the top, whose SKU <b>starts with</b> the code in <b>B1</b>.</p>',
          solution: '=XLOOKUP(B1&"*",A5:A16,D5:D16,,2)',
          alts: ['=VLOOKUP(B1&"*",A5:D16,4,FALSE)', '=INDEX(D5:D16,MATCH(B1&"*",A5:A16,0))'],
          hints: ['B1&"*" means "B1 followed by anything": a starts-with pattern.', 'Lookups search from the top, so they return the first SKU that fits. XLOOKUP needs match_mode 2 to read the *.', 'Shape: =XLOOKUP(B1&"*", A5:A16, D5:D16, , 2)'],
          explain: 'The pattern B1&"*" matches any SKU that begins with the brand code. A lookup stops at the first match from the top, so it returns that line\'s On hand.',
          chips: ['B1&"*"', 'A5:A16', 'D5:D16', ',,2)'], fnKey: 'XLOOKUP('
        });
      }
      const rows = r.shuffle(DIST());
      const S = H.table([{ h: 'Distributor', k: 'name' }, { h: 'City', k: 'city' }, { h: 'Region', k: 'region' }, { h: 'Terms (days)', k: 'terms_days', num: true }], rows);
      const [i, j] = r.sample(nums(14), 2);
      const short = (k) => rows[k].name.replace(/^(CV|UD) /, '');
      const base = Object.assign({}, S.cells, { A1: 'Name without CV or UD', B1: short(i), A2: 'City' });
      const alt = H.edit(base, { B1: short(j) });
      return H.problem({
        title: 'City of a distributor named without CV or UD', file: 'distributor.xlsx', sheet: S, base, alt, answerCell: 'B2',
        statement: '<p>The order form in <b>B1</b> names the distributor without its CV or UD in front. In <b>B2</b>, return the <b>City</b> of the distributor whose full name <b>ends with</b> B1.</p>',
        solution: '=XLOOKUP("*"&B1,A5:A18,B5:B18,,2)',
        alts: ['=VLOOKUP("*"&B1,A5:B18,2,FALSE)', '=INDEX(B5:B18,MATCH("*"&B1,A5:A18,0))'],
        hints: ['Put a * in front of B1 so anything, such as CV or UD, can come before it.', 'XLOOKUP needs match_mode 2 (its fifth part) to read the * as a wildcard.', 'Shape: =XLOOKUP("*"&B1, A5:A18, B5:B18, , 2)'],
        explain: '"*"&B1 matches any name that ends with B1, whatever comes before it. XLOOKUP with match_mode 2 reads the pattern and returns the city from that row.',
        chips: ['"*"&B1', 'A5:A18', 'B5:B18', ',,2)'], fnKey: 'XLOOKUP('
      });
    }
  });

  /* ======================= Hard ======================= */

  P.add({
    key: 'two-way-grid-lookup', difficulty: 'Hard', level: 8, topics: ['INDEX', 'MATCH', 'XLOOKUP'], count: 4,
    make(r, v) {
      /* A grid: one key down column A, the other across row 4. B1 picks the row, B2 the column, the answer goes in B3. */
      if (v <= 2) {
        const V = [
          { title: 'Freight per carton where a region meets a warehouse', file: 'tarif_kirim.xlsx', corner: 'Region', rows: H.REGIONS.slice(), heads: ['Bekasi', 'Bandung', 'Semarang', 'Surabaya', 'Medan'],
            val: () => r.int(6, 40) * 500, rl: 'Region', cl: 'Warehouse', al: 'Freight (Rp per carton)', step: 1500, what: 'the freight cost per carton from the warehouse in <b>B2</b> to the region in <b>B1</b>' },
          { title: 'Interest rate by loan type and tenor (two-way lookup)', file: 'suku_bunga.xlsx', corner: 'Loan type', rows: ['Working Capital', 'Term Loan', 'Trade Finance', 'Investment'], heads: [1, 2, 3, 5],
            val: null, rl: 'Loan type', cl: 'Tenor (years)', al: 'Rate (% a year)', step: 0.25, what: 'the yearly interest rate for the loan type in <b>B1</b> and the tenor in years in <b>B2</b>' },
          { title: "One rep's sales in a chosen month (two-way lookup)", file: 'sales_per_rep.xlsx', corner: 'Rep', rows: r.sample(H.REPS, 8), heads: ['Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
            val: () => r.int(40, 160) * 1000000, rl: 'Sales rep', cl: 'Month', al: 'Sales (Rp)', step: 2500000, what: 'the sales of the rep in <b>B1</b> in the month in <b>B2</b>' }
        ][v];
        const b0 = r.pick([8.5, 9]), add = [0, 1, -0.5, 1.5], ten = [0, 0.25, 0.5, 1];
        const vals = V.rows.map((x, i) => V.heads.map((h, j) => (V.val ? V.val() : b0 + add[i] + ten[j] + r.pick([0, 0.25]))));
        const S = across(V.corner, V.heads, V.rows.map((h, i) => ({ h, vals: vals[i] })));
        const nr = V.rows.length, nc = V.heads.length;
        const i = r.int(1, nr - 1), j = r.int(1, nc - 1);
        let i2, j2;
        do { i2 = r.int(0, nr - 1); j2 = r.int(0, nc - 1); } while (i2 === i || j2 === j);
        let d = V.step;
        while (vals[i2][j2] + d === vals[i][j]) d += V.step;
        const lc = L[nc], lr = 4 + nr;
        const G = 'B5:' + lc + lr, RK = 'A5:A' + lr, CK = 'B4:' + lc + '4';
        const base = Object.assign({}, S.cells, { A1: V.rl, B1: V.rows[i], A2: V.cl, B2: V.heads[j], A3: V.al });
        const alt = H.edit(base, { B1: V.rows[i2], B2: V.heads[j2], [L[j2 + 1] + (5 + i2)]: (x) => x + d });
        return H.problem({
          title: V.title, file: V.file, sheet: S, base, alt, answerCell: 'B3',
          statement: `<p>In <b>B3</b>, return ${V.what}. Find the row with MATCH on ${RK} and the column with MATCH on ${CK}.</p>`,
          solution: `=INDEX(${G},MATCH(B1,${RK},0),MATCH(B2,${CK},0))`,
          alts: [`=XLOOKUP(B1,${RK},XLOOKUP(B2,${CK},${G}))`, `=VLOOKUP(B1,A5:${lc}${lr},MATCH(B2,A4:${lc}4,0),FALSE)`, `=XLOOKUP(B2,${CK},XLOOKUP(B1,${RK},${G}))`, `=INDEX(${G},XMATCH(B1,${RK}),XMATCH(B2,${CK}))`],
          hints: ['INDEX can take a row number and a column number inside a block of cells.', `MATCH(B1, ${RK}, 0) gives the row; MATCH(B2, ${CK}, 0) gives the column.`, `Shape: =INDEX(${G}, MATCH(B1, ${RK}, 0), MATCH(B2, ${CK}, 0))`],
          explain: `The first MATCH finds which row of the grid holds B1 and the second which column holds B2. INDEX returns the cell where that row and that column cross in ${G}.`,
          chips: [G, RK, CK, 'MATCH(B1,', 'MATCH(B2,', ',0)'], fnKey: 'INDEX('
        });
      }
      const ps = r.shuffle(H.PRODUCTS);
      const rows = ps.map((p) => ({ sku: p.sku, packs: p.packs_per_carton, w: p.carton_price, m: Math.round(p.carton_price * 1.08 / 1000) * 1000, k: Math.round(p.carton_price * 1.15 / 1000) * 1000 }));
      const S = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Packs per carton', k: 'packs', num: true }, { h: 'Wholesale', k: 'w', num: true }, { h: 'Modern trade', k: 'm', num: true }, { h: 'Warung', k: 'k', num: true }], rows);
      const ch = ['Wholesale', 'Modern trade', 'Warung'], ck = ['w', 'm', 'k'];
      const i = r.int(1, 11), j = r.int(0, 2);
      let i2, j2;
      do { i2 = r.int(0, 11); j2 = r.int(0, 2); } while (i2 === i || j2 === j);
      const per = (a, b, d) => Math.round((rows[a][ck[b]] + d) / rows[a].packs);
      let d = 12000;
      while (per(i2, j2, d) === per(i, j, 0)) d += 6000;
      const base = Object.assign({}, S.cells, { A1: 'SKU', B1: rows[i].sku, A2: 'Channel', B2: ch[j], A3: 'Price per pack (Rp)' });
      const alt = H.edit(base, { B1: rows[i2].sku, B2: ch[j2], [L[j2 + 2] + (5 + i2)]: (x) => x + d });
      return H.problem({
        title: 'Price per pack by SKU and sales channel', file: 'harga_per_kanal.xlsx', sheet: S, base, alt, answerCell: 'B3',
        statement: '<p>Columns C to E hold the carton price (Rp) for each sales channel. In <b>B3</b>, return the price of one pack of the SKU in <b>B1</b> in the channel in <b>B2</b>: its carton price for that channel divided by its <b>Packs per carton</b>, rounded to whole rupiah with ROUND.</p>',
        solution: '=ROUND(INDEX(C5:E16,MATCH(B1,A5:A16,0),MATCH(B2,C4:E4,0))/XLOOKUP(B1,A5:A16,B5:B16),0)',
        alts: ['=ROUND(XLOOKUP(B1,A5:A16,XLOOKUP(B2,C4:E4,C5:E16))/XLOOKUP(B1,A5:A16,B5:B16),0)', '=ROUND(VLOOKUP(B1,A5:E16,MATCH(B2,A4:E4,0),FALSE)/VLOOKUP(B1,A5:E16,2,FALSE),0)', '=ROUND(INDEX(C5:E16,MATCH(B1,A5:A16,0),MATCH(B2,C4:E4,0))/INDEX(B5:B16,MATCH(B1,A5:A16,0)),0)'],
        hints: ['Two lookups: a two-way one for the carton price (SKU down, channel across) and a plain one for the packs per carton.', 'INDEX(C5:E16, MATCH(B1, A5:A16, 0), MATCH(B2, C4:E4, 0)) is the carton price; XLOOKUP(B1, A5:A16, B5:B16) is the packs.', 'Shape: =ROUND(INDEX(C5:E16, row, column) / XLOOKUP(B1, A5:A16, B5:B16), 0)'],
        explain: 'INDEX with two MATCHes picks the carton price where the SKU row meets the channel column. Dividing by the SKU\'s packs per carton, found with XLOOKUP, and rounding gives the price of one pack.',
        chips: ['C5:E16', 'A5:A16', 'C4:E4', 'B5:B16', ',0)'], fnKey: 'INDEX('
      });
    }
  });

  P.add({
    key: 'tier-from-thresholds', difficulty: 'Hard', level: 8, topics: ['VLOOKUP', 'XLOOKUP', 'approximate match'], count: 4,
    make(r, v) {
      /* A tier table sorted by its lower limits: an approximate match returns the tier the value falls in. */
      const V = [
        { title: 'Volume discount from a carton tier table', file: 'diskon_volume.xlsx', ans: 'Discount (%)', q: 'Cartons',
          main: [{ h: 'Date', k: 'date', date: true }, { h: 'Distributor', k: 'distributor' }, { h: 'SKU', k: 'sku' }, { h: 'Cartons', k: 'cartons', num: true }], rows: () => H.sales(r, 14),
          side: ['From cartons', 'Discount (%)'], tiers: () => { const t = r.pick([[0, 20, 40, 60], [0, 15, 30, 50], [0, 25, 45, 70]]), dd = r.pick([[0, 2, 3.5, 5], [0, 1.5, 3, 4.5], [0, 2.5, 4, 6]]); return t.map((f, k) => [f, dd[k]]); }, bump: 0.5,
          rule: 'the <b>Discount (%)</b> for the line\'s <b>Cartons</b>, as the table shows it (2.5 means 2.5%): a line gets the rate of the highest <b>From cartons</b> it reaches (at least that many cartons)' },
        { title: 'Kol grade from a days-past-due tier table', file: 'kolektibilitas.xlsx', ans: 'Kol', q: 'DPD',
          main: [{ h: 'Facility', k: 'facility' }, { h: 'Distributor', k: 'distributor' }, { h: 'Outstanding (Rp)', k: 'outstanding', num: true }, { h: 'DPD', k: 'dpd', num: true }], rows: () => H.loans(r, 14),
          side: ['From DPD', 'Kol'], tiers: () => [[0, 1], [1, 2], [91, 3], [121, 4], [181, 5]],
          rule: 'the collectibility grade (Kol) for the facility\'s <b>DPD</b> (days past due): 0 days is Kol 1, 1 to 90 is Kol 2, 91 to 120 Kol 3, 121 to 180 Kol 4 and more than 180 Kol 5, as the table lists' },
        { title: 'Rep commission from an achievement tier table', file: 'komisi_sales.xlsx', ans: 'Commission (Rp)', ratio: true,
          main: [{ h: 'Rep', k: 'name' }, { h: 'Area', k: 'area' }, { h: 'Target (Rp)', k: 'target', num: true }, { h: 'Actual (Rp)', k: 'actual', num: true }], rows: () => H.reps(r, 10),
          side: ['From achievement', 'Commission (%)'], tiers: () => { const c = r.pick([[0, 0.5, 1, 1.5], [0, 1, 1.5, 2.5], [0, 0.75, 1.25, 2]]); return [0, 0.8, 1, 1.1].map((f, k) => [f, c[k]]); }, bump: 0.25,
          rule: 'the commission in rupiah: <b>Actual (Rp)</b> times the commission rate for the rep\'s achievement (Actual divided by Target), divided by 100 and rounded to whole rupiah. A rep gets the rate of the highest <b>From achievement</b> reached' },
        { title: 'Aging bucket from a days-overdue tier table', file: 'umur_piutang.xlsx', ans: 'Bucket', q: 'Days overdue',
          main: [{ h: 'Invoice', k: 'ref' }, { h: 'Distributor', k: 'distributor' }, { h: 'Amount (Rp)', k: 'amount', num: true }, { h: 'Days overdue', k: 'over', num: true }],
          rows: () => H.invoices(r, 14).map((x) => Object.assign(x, { over: r.pick([0, 0, 4, 12, 25, 30, 31, 44, 60, 61, 77, 90, 91, 130]) })),
          side: ['From days', 'Bucket'], tiers: () => [[0, 'Current'], [1, '1 to 30'], [31, '31 to 60'], [61, '61 to 90'], [91, 'Over 90']],
          rule: 'the aging bucket for the invoice\'s <b>Days overdue</b>, from the tier table: an invoice falls in the bucket of the highest <b>From days</b> it reaches' }
      ][v];
      const rows = V.rows(), tiers = V.tiers();
      const S = H.table(V.main, rows);
      const M = H.table([{ h: V.side[0], v: (t) => t[0], num: true }, { h: V.side[1], v: (t) => t[1], num: typeof tiers[0][1] === 'number' }], tiers, { startCol: 6 });
      const B = board([S, M]);
      const tierOf = (x) => { let k = 0; tiers.forEach((t, m) => { if (x >= t[0]) k = m; }); return k; };
      const qOf = (k) => (V.ratio ? rows[k].actual / rows[k].target : rows[k][V.main[3].k]);
      /* The hidden sheet: row 1 sits exactly on a tier's lower limit, in a different tier than before. */
      const t0 = tierOf(qOf(0));
      const t = r.pick(nums(tiers.length).filter((m) => m > 0 && m !== t0));
      const e = {};
      if (V.ratio) e[S.cell('Actual (Rp)', 0)] = Math.round(rows[0].target * tiers[t][0] * 100) / 100;
      else e[S.cell(V.q, 0)] = tiers[t][0];
      if (V.bump) e[M.cell(V.side[1], tiers.length - 1)] = (x) => x + V.bump;
      const alt = H.edit(B.cells, e);
      const KA = M.A(V.side[0]), VA = M.A(V.side[1]), TA = '$G$5:$H$' + M.last;
      const x = V.ratio ? 'D5/C5' : 'D5';
      const wrap = (f) => (V.ratio ? '=ROUND(D5*' + f + '/100,0)' : '=' + f);
      return build(B, {
        title: V.title, file: V.file, base: B.cells, alt, fill: true, fillCol: 'E', ansLabel: V.ans,
        statement: `<p>In <b>E5</b>, return ${V.rule}. The tier table in <b>G5:H${M.last}</b> is sorted by its lower limits, smallest first. Fill it down to <b>E${S.last}</b>.</p>`,
        solution: wrap(`VLOOKUP(${x},${TA},2,TRUE)`),
        alts: [wrap(`XLOOKUP(${x},${KA},${VA},,-1)`), wrap(`LOOKUP(${x},${KA},${VA})`), wrap(`INDEX(${VA},MATCH(${x},${KA},1))`), wrap(`VLOOKUP(${x},${TA},2)`)],
        hints: ['An exact match fails here: most values are not in the table. You want the last lower limit that the value reaches.',
          `VLOOKUP with TRUE as its last part does that on a table sorted smallest first; so does XLOOKUP with match_mode -1 (exact or next smaller). Pin the tier table: ${TA}.`,
          `Shape: ${wrap('VLOOKUP(' + x + ', ' + TA + ', 2, TRUE)')}`],
        explain: `With TRUE, VLOOKUP walks down the sorted lower limits and stops at the largest one that is not above ${x}, so a value exactly on a limit gets that tier. XLOOKUP with match_mode -1 gives the same tier without needing the sort.`,
        chips: [x, TA, KA, VA, ',2,TRUE)', ',,-1)'], fnKey: 'VLOOKUP('
      });
    }
  });

  P.add({
    key: 'lookup-on-two-conditions', difficulty: 'Hard', level: 10, topics: ['XLOOKUP', 'INDEX', 'MATCH'], count: 3,
    make(r, v) {
      /* No single column is unique: the row is found by two keys at once, multiplied as 1s and 0s. */
      let rows, S, c1, c2, out, l1, l2, la, title, file, what, text = false, ansDate = false;
      if (v === 0) {
        const skus = r.sample(H.PRODUCTS, 5), whs = r.sample(H.WAREHOUSES, 3);
        rows = [];
        skus.forEach((p) => whs.forEach((w) => rows.push({ sku: p.sku, wh: w, oh: r.int(1, 40) * 10, rp: r.pick([80, 100, 120, 150]) })));
        rows = r.shuffle(rows);
        S = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Warehouse', k: 'wh' }, { h: 'On hand', k: 'oh', num: true }, { h: 'Reorder point', k: 'rp', num: true }], rows);
        c1 = 'SKU'; c2 = 'Warehouse'; out = 'On hand'; l1 = 'SKU'; l2 = 'Warehouse'; la = 'On hand'; file = 'stok_per_gudang.xlsx';
        title = 'On hand for one SKU in one warehouse (two conditions)'; what = 'the <b>On hand</b> for the SKU in <b>B1</b> in the warehouse in <b>B2</b>';
      } else if (v === 1) {
        rows = r.shuffle(H.db().facilities);
        S = H.table([{ h: 'Facility', k: 'facility_id' }, { h: 'Distributor', k: 'distributor' }, { h: 'Type', k: 'type' }, { h: 'Limit (Rp)', k: 'limit', num: true }, { h: 'Outstanding (Rp)', k: 'outstanding', num: true }], rows);
        c1 = 'Distributor'; c2 = 'Type'; out = 'Facility'; l1 = 'Borrower'; l2 = 'Loan type'; la = 'Facility'; text = true; file = 'buku_pinjaman.xlsx';
        title = 'Facility code for a borrower and loan type'; what = 'the <b>Facility</b> code of the borrower in <b>B1</b> for the loan type in <b>B2</b>';
      } else {
        const reg = r.pick(H.REGIONS);
        const seen = new Set();
        rows = H.sales(r, 40, { region: reg }).filter((s) => { const k = s.distributor + '|' + s.sku; if (seen.has(k)) return false; seen.add(k); return true; }).slice(0, 14);
        S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'Invoice', k: 'invoice' }, { h: 'Distributor', k: 'distributor' }, { h: 'SKU', k: 'sku' }, { h: 'Net (Rp)', k: 'net', num: true }], rows);
        c1 = 'Distributor'; c2 = 'SKU'; out = 'Invoice'; l1 = 'Distributor'; l2 = 'SKU'; la = 'Invoice'; text = true; file = 'penjualan_q4.xlsx';
        title = 'Invoice number for a distributor and SKU pair'; what = 'the <b>Invoice</b> on which the distributor in <b>B1</b> bought the SKU in <b>B2</b>';
      }
      const n = rows.length;
      const val = (h, k) => S.cells[S.cell(h, k)];
      /* Pick rows where neither key alone finds the right row, so a one-condition lookup is caught. */
      const firstOf = (h, k) => nums(n).findIndex((m) => val(h, m) === val(h, k));
      const tricky = nums(n).filter((k) => firstOf(c1, k) !== k && firstOf(c2, k) !== k);
      const pool = tricky.length >= 2 ? tricky : nums(n).filter((k) => firstOf(c1, k) !== k);
      const i = r.pick(pool);
      const j = r.pick(pool.filter((k) => k !== i && (val(c1, k) !== val(c1, i) || v !== 1)));
      const base = Object.assign({}, S.cells, { A1: l1, B1: val(c1, i), A2: l2, B2: val(c2, i), A3: la });
      const e = { B1: val(c1, j), B2: val(c2, j) };
      if (!text && val(out, j) === val(out, i)) e[S.cell(out, j)] = val(out, j) + 10;
      const alt = H.edit(base, e);
      const A1r = S.R(c1), A2r = S.R(c2), O = S.R(out);
      const cond = `(${A1r}=B1)*(${A2r}=B2)`;
      const alts = [`=INDEX(${O},MATCH(1,${cond},0))`, `=FILTER(${O},${cond})`, `=XLOOKUP(1,(${A2r}=B2)*(${A1r}=B1),${O})`];
      if (!text) alts.push(`=SUMIFS(${O},${A1r},B1,${A2r},B2)`);
      return H.problem({
        title, file, sheet: S, base, alt, answerCell: 'B3', ansDate,
        statement: `<p>In <b>B3</b>, return ${what}. Neither column alone is unique, so the row must match <b>both</b> B1 and B2.</p>`,
        solution: `=XLOOKUP(1,${cond},${O})`,
        alts,
        hints: [`(${A1r}=B1) is a column of TRUE and FALSE. Multiplying two such columns gives 1 only on rows where both are TRUE.`,
          'Look up the number 1 in that product, and return from the column you want.',
          `Shape: =XLOOKUP(1, (${A1r}=B1)*(${A2r}=B2), ${O})`],
        explain: `Each comparison marks the rows that match one key; multiplying them leaves 1 only where both match. XLOOKUP finds that 1 and returns ${out} from the same row. INDEX with MATCH(1, ..., 0) works the same way.`,
        chips: [A1r, A2r, O, '(' + A1r + '=B1)', '(' + A2r + '=B2)', 'XLOOKUP(1,'], fnKey: 'XLOOKUP('
      });
    }
  });

  P.add({
    key: 'last-match-in-a-log', difficulty: 'Hard', level: 9, topics: ['XLOOKUP', 'XMATCH', 'search mode'], count: 3,
    make(r, v) {
      /* A log in date order: the latest entry for a key is its last match, found by searching from the bottom. */
      if (v === 0) {
        const ids = r.sample(DIST(), 3).map((d) => d.id);
        const rows = r.sample(H.db().sales.filter((s) => ids.includes(s.distributor_id)), 18).map((s) => Object.assign({}, s, { date: H.Q4(s.date) }))
          .sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : a.invoice < b.invoice ? -1 : 1));
        const S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'Invoice', k: 'invoice' }, { h: 'Distributor', k: 'distributor' }, { h: 'Net (Rp)', k: 'net', num: true }], rows);
        /* The asked-for distributor has no tie on its latest date, and its latest order is not its largest, so MAXIFS on
           Net is caught. */
        const rowsOf = (d) => rows.filter((x) => x.distributor === d);
        const lastRow = (d) => rowsOf(d).pop();
        const names = [...new Set(rows.map((x) => x.distributor))].filter((d) => rowsOf(d).length >= 2);
        const uniq = names.filter((d) => rowsOf(d).filter((x) => x.date === lastRow(d).date).length === 1);
        const pool = uniq.length >= 2 ? uniq : names;
        const good = pool.filter((d) => lastRow(d).net < Math.max(...rowsOf(d).map((x) => x.net)));
        const a = r.pick(good.length ? good : pool);
        const b = r.pick(pool.filter((d) => d !== a));
        let bump = 500000;
        while (lastRow(b).net + bump === lastRow(a).net) bump += 100000;
        const base = Object.assign({}, S.cells, { A1: 'Distributor', B1: a, A2: 'Latest Net (Rp)' });
        const alt = H.edit(base, { B1: b, [S.cell('Net (Rp)', rows.indexOf(lastRow(b)))]: (x) => x + bump });
        const D = S.R('Distributor'), T = S.R('Date'), N = S.R('Net (Rp)');
        const alts = [`=INDEX(${N},XMATCH(B1,${D},0,-1))`, `=XLOOKUP(B1,${D},${N},,,-1)`];
        if (pool === uniq) alts.push(`=SUMIFS(${N},${D},B1,${T},MAXIFS(${T},${D},B1))`);
        return H.problem({
          title: "Net value of a distributor's most recent order", file: 'penjualan_q4.xlsx', sheet: S, base, alt, answerCell: 'B2',
          statement: `<p>The sales log is in date order, oldest first. In <b>B2</b>, return the <b>Net (Rp)</b> of the most recent order from the distributor in <b>B1</b>: its last row in the log.</p>`,
          solution: `=XLOOKUP(B1,${D},${N},,0,-1)`,
          alts,
          hints: ['A normal lookup returns the first match from the top, which here is the oldest order.', 'XLOOKUP\'s sixth part, search_mode, set to -1 searches from the bottom up, so it finds the last match.', `Shape: =XLOOKUP(B1, ${D}, ${N}, , 0, -1)`],
          explain: 'With search_mode -1, XLOOKUP starts at the bottom of the log and returns the first match it meets, which is the latest order. The largest Net is not always the latest one, so MAXIFS on Net does not give it.',
          chips: [D, N, 'B1', ',,0,-1)'], fnKey: 'XLOOKUP('
        });
      }
      if (v === 1) {
        const skus = r.sample(H.PRODUCTS, 5);
        const seq = r.shuffle([0, 0, 0, 1, 1, 1, 2, 2, 2, 3, 3, 3, 4, 4, 4, 0]);
        const price = skus.map((p) => p.carton_price);
        let day = H.toSerial('2026-01-05');
        const rows = seq.map((k) => { day += r.int(6, 20); price[k] += r.pick([-6000, -4000, 3000, 5000, 7000, 9000]); return { date: day, sku: skus[k].sku, price: price[k] }; });
        const S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'SKU', k: 'sku' }, { h: 'New price (Rp)', k: 'price', num: true }], rows);
        const lastIdx = (sku) => rows.map((x) => x.sku).lastIndexOf(sku);
        const [a, b] = r.sample(skus.map((p) => p.sku), 2);
        let d = 0;
        if (rows[lastIdx(b)].price === rows[lastIdx(a)].price) d = 2000;
        const base = Object.assign({}, S.cells, { A1: 'SKU', B1: a, A2: 'Current price (Rp)' });
        const alt = H.edit(base, { B1: b, [S.cell('New price (Rp)', lastIdx(b))]: (x) => x + d + 1000 });
        const K = S.R('SKU'), Pr = S.R('New price (Rp)');
        return H.problem({
          title: 'Current price from a price change log', file: 'riwayat_harga.xlsx', sheet: S, base, alt, answerCell: 'B2',
          statement: '<p>Each row records a price change, oldest first; the newest change for a SKU is its current price. In <b>B2</b>, return the current price of the SKU in <b>B1</b>.</p>',
          solution: `=XLOOKUP(B1,${K},${Pr},,0,-1)`,
          alts: [`=INDEX(${Pr},XMATCH(B1,${K},0,-1))`],
          hints: ['The SKU appears several times; you want its last row, not its first.', 'XLOOKUP searches from the bottom when its sixth part, search_mode, is -1.', `Shape: =XLOOKUP(B1, ${K}, ${Pr}, , 0, -1)`],
          explain: 'Searching from the bottom, XLOOKUP meets the SKU\'s newest change first and returns that price. The highest price is not always the current one, since some changes were cuts.',
          chips: [K, Pr, 'B1', ',,0,-1)'], fnKey: 'XLOOKUP('
        });
      }
      const fac = r.sample(H.db().facilities, 4).map((f) => f.facility_id);
      const ends = ['2026-07-31', '2026-08-31', '2026-09-30', '2026-10-31'];
      const dpd = fac.map(() => r.pick([0, 0, 5, 15, 30]));
      const rows = [];
      ends.forEach((m, mi) => r.shuffle(nums(4)).forEach((k) => { if (mi) dpd[k] = Math.max(0, dpd[k] + r.pick([-15, 0, 0, 10, 20, 30])); rows.push({ m, f: fac[k], dpd: dpd[k] }); }));
      const S = H.table([{ h: 'Month end', k: 'm', date: true }, { h: 'Facility', k: 'f' }, { h: 'DPD', k: 'dpd', num: true }], rows);
      const chg = (f) => { const x = rows.filter((y) => y.f === f); return x[x.length - 1].dpd - x[0].dpd; };
      const [a, b] = r.sample(fac, 2);
      const lb = rows.map((y) => y.f).lastIndexOf(b);
      const base = Object.assign({}, S.cells, { A1: 'Facility', B1: a, A2: 'Change in DPD' });
      const alt = H.edit(base, { B1: b, [S.cell('DPD', lb)]: (x) => x + (chg(b) + 5 === chg(a) ? 10 : 5) });
      const K = S.R('Facility'), D = S.R('DPD');
      return H.problem({
        title: 'Change in days past due from first to last check', file: 'pantauan_dpd.xlsx', sheet: S, base, alt, answerCell: 'B2',
        statement: '<p>Dimas records each facility\'s days past due (DPD) at every month end, oldest first. In <b>B2</b>, return the change for the facility in <b>B1</b>: its latest DPD minus its first DPD in the log (negative when it improved).</p>',
        solution: `=XLOOKUP(B1,${K},${D},,0,-1)-XLOOKUP(B1,${K},${D})`,
        alts: [`=INDEX(${D},XMATCH(B1,${K},0,-1))-INDEX(${D},MATCH(B1,${K},0))`],
        hints: ['You need two lookups on the same facility: the first match from the top and the first match from the bottom.', 'A plain XLOOKUP gives the first record; with search_mode -1 (its sixth part) it gives the last.', `Shape: =XLOOKUP(B1, ${K}, ${D}, , 0, -1) - XLOOKUP(B1, ${K}, ${D})`],
        explain: 'The XLOOKUP that searches from the bottom returns the latest DPD, the normal one the earliest. Their difference shows whether the facility got worse (positive) or better (negative).',
        chips: [K, D, 'B1', ',,0,-1)'], fnKey: 'XLOOKUP('
      });
    }
  });

  P.add({
    key: 'price-on-the-order-date', difficulty: 'Hard', level: 9, topics: ['XLOOKUP', 'VLOOKUP', 'approximate match', 'dates'], count: 3,
    make(r, v) {
      /* A history of changes sorted by start date: the one in force on a date is the last start on or before it. */
      const rows = H.sales(r, 14);
      let side, sv, ans, title, file, what, bumpRow, wrap, main, num = true, extra = {};
      const rs = rows.map((x) => H.toSerial(x.date));
      const starts = (at) => { let prev = H.toSerial('2026-09-01'); return [prev].concat(at.map((k) => (prev = Math.max(rs[k], prev + 1)))); };
      if (v === 0) {
        const p = r.pick(H.PRODUCTS);
        let pr = p.carton_price;
        side = starts([3, 7, 11]).map((d, k) => ({ d, x: k ? (pr += r.pick([4000, 6000, 8000, -3000])) : pr }));
        sv = 'Carton price (Rp)'; ans = 'Value (Rp)'; file = 'pesanan_' + p.sku.toLowerCase() + '.xlsx';
        title = 'Order value at the price in force on its date';
        main = [{ h: 'Date', k: 'date', date: true }, { h: 'Invoice', k: 'invoice' }, { h: 'Distributor', k: 'distributor' }, { h: 'Cartons', k: 'cartons', num: true }];
        what = `the value of each order of ${p.name}: its <b>Cartons</b> times the carton price in force on the order <b>Date</b>`;
        wrap = (f) => '=D5*' + f;
        extra = { A1: 'Product', B1: p.name };
      } else if (v === 1) {
        side = starts([r.int(2, 4), r.int(6, 8), r.int(10, 12)]).map((d, k) => ({ d, x: [9.5, 9.75, 10.25, 10][k] + r.pick([-0.25, 0, 0.25]) }));
        rows.forEach((x) => { x.amount = r.pick([20, 25, 40, 50, 75, 100]) * 1000000; x.days = r.pick([30, 45, 60, 90]); });
        sv = 'Base rate (% a year)'; ans = 'Interest (Rp)'; file = 'pencairan_kredit.xlsx';
        title = 'Drawdown interest at the base rate on its date';
        main = [{ h: 'Date', k: 'date', date: true }, { h: 'Distributor', k: 'distributor' }, { h: 'Amount (Rp)', k: 'amount', num: true }, { h: 'Days', k: 'days', num: true }];
        what = 'the interest on each drawdown: <b>Amount (Rp)</b> times the base rate in force on its <b>Date</b>, divided by 100, times <b>Days</b> divided by 365, rounded to whole rupiah with ROUND';
        wrap = (f) => '=ROUND(C5*' + f + '/100*D5/365,0)';
      } else {
        side = starts([r.int(2, 3), r.int(5, 6), r.int(8, 9), r.int(11, 12)]).map((d, k) => ({ d, x: ['No promo', 'Payday Deal', 'No promo', 'Payday Deal', 'Year-end Sale'][k] }));
        sv = 'Promo'; ans = 'Promo'; file = 'kalender_promo.xlsx'; num = false;
        title = 'Promo running on each order date';
        main = [{ h: 'Date', k: 'date', date: true }, { h: 'Invoice', k: 'invoice' }, { h: 'Distributor', k: 'distributor' }, { h: 'Net (Rp)', k: 'net', num: true }];
        what = 'the <b>Promo</b> running on the order <b>Date</b>: the one whose <b>From</b> date is the latest on or before it';
        wrap = (f) => '=' + f;
      }
      const S = H.table(main, rows);
      const M = H.table([{ h: 'From', k: 'd', date: true }, { h: sv, k: 'x', num }], side, { startCol: 6 });
      const B = board([S, M]);
      const ser = side.map((s) => s.d);
      const tierOf = (x) => { let k = 0; ser.forEach((s, m) => { if (x >= s) k = m; }); return k; };
      /* The hidden sheet: one order moves to exactly the start date of a different period. */
      const k = nums(14).find((m) => rs[m] !== ser[2] && side[tierOf(rs[m])].x !== side[2].x);
      bumpRow = k === undefined ? 0 : k;
      const e = { [S.cell('Date', bumpRow)]: ser[2] };
      if (num) e[M.cell(sv, 1)] = (x) => x + (v === 1 ? 0.25 : 2000);
      const alt = H.edit(Object.assign({}, B.cells, extra), e);
      const KA = M.A('From'), VA = M.A(sv), TA = '$G$5:$H$' + M.last;
      return build(B, {
        title, file, base: Object.assign({}, B.cells, extra), alt, fill: true, fillCol: 'E', ansLabel: ans,
        statement: `<p>The table in <b>G5:H${M.last}</b> lists each change with the date it starts, oldest first; it stays in force until the next one starts. In <b>E5</b>, return ${what}. Fill it down to <b>E${S.last}</b>.</p>`,
        solution: wrap(`XLOOKUP(A5,${KA},${VA},,-1)`),
        alts: [wrap(`VLOOKUP(A5,${TA},2,TRUE)`), wrap(`LOOKUP(A5,${KA},${VA})`), wrap(`INDEX(${VA},MATCH(A5,${KA},1))`)],
        hints: ['Dates are numbers, so an approximate match works on them too: you want the last start date that is on or before the order date.',
          `XLOOKUP with match_mode -1 returns the exact date or the next smaller one. Pin the table: ${KA} and ${VA}.`,
          `Shape: ${wrap('XLOOKUP(A5, ' + KA + ', ' + VA + ', , -1)')}`],
        explain: `match_mode -1 makes XLOOKUP fall back to the latest start date before the order date when there is no exact match, which is the period in force. An order on a start date gets the new value. VLOOKUP with TRUE does the same on the sorted table.`,
        chips: ['A5', KA, VA, TA, ',,-1)', ',2,TRUE)'], fnKey: 'XLOOKUP('
      });
    }
  });

  P.add({
    key: 'nested-lookup-chain', difficulty: 'Hard', level: 9, topics: ['XLOOKUP', 'VLOOKUP', 'nested lookup'], count: 3,
    make(r, v) {
      /* The answer is two or three lookups deep: the result of one lookup is what the next one searches for. */
      if (v === 0) {
        const inv = H.invoices(r, 14), dl = r.shuffle(DIST());
        const S = H.table([{ h: 'Invoice', k: 'ref' }, { h: 'Date', k: 'date', date: true }, { h: 'Distributor', k: 'distributor' }, { h: 'Amount (Rp)', k: 'amount', num: true }], inv);
        const M = H.table([{ h: 'Distributor', k: 'name' }, { h: 'City', k: 'city' }, { h: 'Region', k: 'region' }], dl, { startCol: 5 });
        const B = board([S, M]);
        const reg = (k) => DIST().find((d) => d.name === inv[k].distributor).region;
        const i = r.int(1, 13);
        const j = r.pick(nums(14).filter((k) => reg(k) !== reg(i)));
        const base = Object.assign({}, B.cells, { A1: 'Invoice', B1: inv[i].ref, A2: 'Region' });
        const alt = H.edit(base, { B1: inv[j].ref });
        return build(B, {
          title: 'Region of the distributor on an invoice (nested lookup)', file: 'faktur_q4.xlsx', base, alt, answerCell: 'B2',
          statement: '<p>The invoice list does not show regions; the distributor list in <b>F5:H18</b> does. In <b>B2</b>, return the <b>Region</b> of the distributor on the invoice in <b>B1</b>, in one formula.</p>',
          solution: '=XLOOKUP(XLOOKUP(B1,A5:A18,C5:C18),F5:F18,H5:H18)',
          alts: ['=VLOOKUP(VLOOKUP(B1,A5:D18,3,FALSE),F5:H18,3,FALSE)', '=INDEX(H5:H18,MATCH(INDEX(C5:C18,MATCH(B1,A5:A18,0)),F5:F18,0))', '=XLOOKUP(XLOOKUP(B1,A5:A18,C5:C18),F5:F18,H5:H18,"",0)'],
          hints: ['Step one: which distributor is on the invoice? Step two: which region is that distributor in?', 'The first lookup, XLOOKUP(B1, A5:A18, C5:C18), returns a name. Use it as the value to find in the second lookup.', 'Shape: =XLOOKUP(XLOOKUP(B1, A5:A18, C5:C18), F5:F18, H5:H18)'],
          explain: 'The inner XLOOKUP turns the invoice number into a distributor name. The outer one searches that name in the distributor list and returns its region.',
          chips: ['A5:A18', 'C5:C18', 'F5:F18', 'H5:H18', 'B1'], fnKey: 'XLOOKUP('
        });
      }
      if (v === 1) {
        const rows = H.sales(r, 14);
        const prods = r.shuffle(H.PRODUCTS);
        const base4 = [r.pick([16, 18, 20]), r.pick([24, 26, 28]), r.pick([30, 32, 35]), r.pick([25, 27, 29])];
        const cats = r.shuffle(H.CATEGORIES.map((c, k) => ({ c, m: base4[k] })));
        const S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'SKU', k: 'sku' }, { h: 'Cartons', k: 'cartons', num: true }, { h: 'Gross (Rp)', k: 'gross', num: true }], rows);
        const M = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Category', k: 'category' }], prods, { startCol: 6 });
        const C = H.table([{ h: 'Category', k: 'c' }, { h: 'Margin (%)', k: 'm', num: true }], cats, { startCol: 9 });
        const B = board([S, M, C]);
        const catOf = (sku) => H.PRODUCTS.find((p) => p.sku === sku).category;
        const ci = cats.findIndex((x) => x.c === catOf(rows[0].sku));
        const swap = r.pick(H.PRODUCTS.filter((p) => p.category !== catOf(rows[3].sku)));
        const alt = H.edit(B.cells, { [C.cell('Margin (%)', ci)]: (x) => x + 2, [S.cell('SKU', 3)]: swap.sku, [S.cell('Gross (Rp)', 3)]: rows[3].cartons * swap.carton_price });
        const inner = 'XLOOKUP(B5,$G$5:$G$16,$H$5:$H$16)';
        return build(B, {
          title: 'Line margin through a category margin table', file: 'margin_q4.xlsx', base: B.cells, alt, fill: true, fillCol: 'E', ansLabel: 'Margin (Rp)',
          statement: `<p>Margins are set per category, not per SKU. In <b>E5</b>, return the margin in rupiah of each line: its <b>Gross (Rp)</b> times the <b>Margin (%)</b> of its SKU's category, divided by 100 and rounded to whole rupiah with ROUND. The SKU's category is in <b>G5:H16</b> and the category margins in <b>J5:K8</b>. Fill it down to <b>E${S.last}</b>.</p>`,
          solution: `=ROUND(D5*XLOOKUP(${inner},$J$5:$J$8,$K$5:$K$8)/100,0)`,
          alts: ['=ROUND(D5*VLOOKUP(VLOOKUP(B5,$G$5:$H$16,2,FALSE),$J$5:$K$8,2,FALSE)/100,0)', '=ROUND(D5*INDEX($K$5:$K$8,MATCH(INDEX($H$5:$H$16,MATCH(B5,$G$5:$G$16,0)),$J$5:$J$8,0))/100,0)'],
          hints: ['Two steps: the SKU gives a category (G:H), and the category gives a margin (J:K).', `The inner lookup ${inner} returns the category; search for it in $J$5:$J$8. Pin both tables with $.`, `Shape: =ROUND(D5 * XLOOKUP(${inner}, $J$5:$J$8, $K$5:$K$8) / 100, 0)`],
          explain: 'The inner XLOOKUP turns the SKU into its category; the outer one finds that category\'s margin. Times the gross, divided by 100 and rounded, that is the line\'s margin in rupiah.',
          chips: ['D5*', 'B5', '$G$5:$G$16', '$H$5:$H$16', '$J$5:$J$8', '$K$5:$K$8'], fnKey: 'XLOOKUP('
        });
      }
      const outs = r.sample(H.db().outlets, 14).sort((a, b) => (a.outlet_id < b.outlet_id ? -1 : 1));
      const dl = r.shuffle(DIST());
      const map = { Jabodetabek: 'Gudang Bekasi', 'Jawa Barat': 'Gudang Bandung', 'Jawa Tengah': 'Gudang Semarang', 'Jawa Timur': 'Gudang Surabaya', 'Bali Nusra': 'Gudang Surabaya', Sumatera: 'Gudang Medan', Sulawesi: r.pick(['Gudang Surabaya', 'Gudang Semarang']) };
      const regs = r.shuffle(H.REGIONS).map((g) => ({ g, w: map[g] }));
      const S = H.table([{ h: 'Outlet ID', k: 'outlet_id' }, { h: 'Outlet', k: 'outlet' }, { h: 'Distributor ID', k: 'distributor_id' }], outs);
      const M = H.table([{ h: 'ID', k: 'id' }, { h: 'Distributor', k: 'name' }, { h: 'Region', k: 'region' }], dl, { startCol: 4 });
      const C = H.table([{ h: 'Region', k: 'g' }, { h: 'Warehouse', k: 'w' }], regs, { startCol: 8 });
      const B = board([S, M, C]);
      const whOf = (k) => map[DIST().find((d) => d.id === outs[k].distributor_id).region];
      const i = r.int(1, 13);
      const j = r.pick(nums(14).filter((k) => whOf(k) !== whOf(i)));
      const base = Object.assign({}, B.cells, { A1: 'Outlet ID', B1: outs[i].outlet_id, A2: 'Warehouse' });
      const alt = H.edit(base, { B1: outs[j].outlet_id });
      return build(B, {
        title: 'Warehouse serving an outlet, through three lookups', file: 'jalur_distribusi.xlsx', base, alt, answerCell: 'B2',
        statement: '<p>An outlet buys from a distributor (A:C), the distributor sits in a region (E:G), and each region is served by one warehouse (I:J). In <b>B2</b>, return the <b>Warehouse</b> that serves the outlet in <b>B1</b>, in one formula.</p>',
        solution: '=XLOOKUP(XLOOKUP(XLOOKUP(B1,A5:A18,C5:C18),E5:E18,G5:G18),I5:I11,J5:J11)',
        alts: ['=VLOOKUP(VLOOKUP(VLOOKUP(B1,A5:C18,3,FALSE),E5:G18,3,FALSE),I5:J11,2,FALSE)', '=INDEX(J5:J11,MATCH(INDEX(G5:G18,MATCH(INDEX(C5:C18,MATCH(B1,A5:A18,0)),E5:E18,0)),I5:I11,0))'],
        hints: ['Work outwards: outlet to distributor ID, distributor ID to region, region to warehouse.', 'Build it in steps: XLOOKUP(B1, A5:A18, C5:C18) gives the distributor ID; wrap it in a lookup on E5:E18 that returns G5:G18, the region.', 'Shape: =XLOOKUP(XLOOKUP(XLOOKUP(B1, A5:A18, C5:C18), E5:E18, G5:G18), I5:I11, J5:J11)'],
        explain: 'The innermost XLOOKUP turns the outlet into its distributor ID, the middle one turns that ID into a region, and the outer one turns the region into its warehouse. Each answer is the value the next lookup searches for.',
        chips: ['A5:A18', 'C5:C18', 'E5:E18', 'G5:G18', 'I5:I11', 'J5:J11', 'B1'], fnKey: 'XLOOKUP('
      });
    }
  });
});
