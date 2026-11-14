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

  /* ---------- other workbooks ----------
     Real reports pull from several files: the learner's own lines sit in the main workbook, and the masters, rates,
     targets, limits and schedules sit in others. A formula looks things up there first, then does the math. wb() is
     one other workbook with one sheet, as shown and as changed on the hidden test. */
  const X = H.ext;
  const wb = (file, name, S, alt) => ({ file, sheets: [{ name, sheet: S, base: S.cells, alt: alt || S.cells }] });
  const COUNT = ['', 'One', 'Two', 'Three', 'Four'];
  /* "Three workbooks are open: a (x.xlsx), b (y.xlsx) and c (z.xlsx)." from [[what, file], ...]. */
  const opens = (list) => { const p = list.map(([d, f]) => d + ' (<b>' + f + '</b>)'); return COUNT[list.length] + ' workbooks are open: ' + p.slice(0, -1).join(', ') + ' and ' + p[p.length - 1] + '.'; };
  /* The ranges of a grid in another workbook (made by across(), or laid out the same way): the values, the row keys
     down column A, the column keys across row 4, and the same with column A and row 4 included, for VLOOKUP. */
  function gridRefs(file, name, G) {
    const p = X(file, name, G).prefix, lc = G.cols[G.cols.length - 1], lr = G.last;
    return { p, G: p + '$B$5:$' + lc + '$' + lr, RK: p + '$A$5:$A$' + lr, CK: p + '$B$4:$' + lc + '$4', ALL: p + '$A$5:$' + lc + '$' + lr, HEAD: p + '$A$4:$' + lc + '$4' };
  }
  /* Columns of the distributor master, a workbook many of these reports look things up in. */
  const DC = { id: { h: 'ID', k: 'id' }, name: { h: 'Distributor', k: 'name' }, city: { h: 'City', k: 'city' }, region: { h: 'Region', k: 'region' },
    terms: { h: 'Terms (days)', k: 'terms_days', num: true }, disc: { h: 'Discount (%)', k: 'tier_discount', num: true } };
  const regionOf = (name) => DIST().find((d) => d.name === name).region;
  /* The index of the tier a value falls in, for lower limits sorted smallest first. */
  const tierAt = (lows, x) => { let k = 0; lows.forEach((lo, m) => { if (x >= lo) k = m; }); return k; };
  /* Start dates for a schedule: 1 September, then the dates of the given order rows, always rising. */
  const startDates = (rs, at) => { let prev = H.toSerial('2026-09-01'); return [prev].concat(at.map((k) => (prev = Math.max(rs[k], prev + 1)))); };

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
    key: 'price-times-quantity', difficulty: 'Medium', level: 4, topics: ['XLOOKUP', 'VLOOKUP', 'Several workbooks'], count: 4, multi: true,
    make(r, v) {
      /* The lines are in the learner's own file (the code in C, the quantity in D); the rate list is another workbook.
         One lookup into it, times the quantity, fills column E. */
      const V = [
        { title: 'Line value with carton prices from another workbook', file: 'pesanan_q4.xlsx', sheetName: 'Pesanan', mine: 'the order lines', who: 'Tika', ans: 'Value (Rp)',
          main: [{ h: 'Date', k: 'date', date: true }, { h: 'Distributor', k: 'distributor' }, { h: 'SKU', k: 'sku' }, { h: 'Cartons', k: 'cartons', num: true }], rows: () => H.sales(r, 14),
          book: 'daftar_harga.xlsx', bsheet: 'Harga', bdesc: 'the price list',
          side: [{ h: 'SKU', k: 'sku' }, { h: 'Product', k: 'name' }, { h: 'Carton price (Rp)', k: 'carton_price', num: true }], list: () => r.shuffle(H.PRODUCTS), step: 6000, qstep: 5,
          task: 'multiply the line\'s <b>Cartons</b> by the <b>Carton price (Rp)</b> of its <b>SKU</b> in the price list' },
        { title: 'Packs shipped, with pack sizes from the product master', file: 'pengiriman_q4.xlsx', sheetName: 'Kirim', mine: 'the shipments', who: 'Pak Anton', ans: 'Packs',
          main: [{ h: 'Date', k: 'date', date: true }, { h: 'Invoice', k: 'invoice' }, { h: 'SKU', k: 'sku' }, { h: 'Cartons', k: 'cartons', num: true }], rows: () => H.sales(r, 14),
          book: 'master_produk.xlsx', bsheet: 'Produk', bdesc: 'the product master',
          side: [{ h: 'SKU', k: 'sku' }, { h: 'Product', k: 'name' }, { h: 'Category', k: 'category' }, { h: 'Packs per carton', k: 'packs_per_carton', num: true }], list: () => r.shuffle(H.PRODUCTS), step: 8, qstep: 4,
          task: 'turn the line\'s <b>Cartons</b> into packs with the <b>Packs per carton</b> of its <b>SKU</b> in the product master' },
        { title: 'Stock value at cost, with costs kept in another workbook', file: 'stok_gudang.xlsx', sheetName: 'Stok', mine: 'the stock count', who: 'Bu Lestari', ans: 'Stock value (Rp)',
          main: [{ h: 'Warehouse', k: 'warehouse' }, { h: 'Product', k: 'product' }, { h: 'SKU', k: 'sku' }, { h: 'On hand', k: 'onhand', num: true }], rows: () => r.shuffle(H.stock(r, 12)),
          book: 'biaya_produk.xlsx', bsheet: 'Biaya', bdesc: 'the cost file',
          side: [{ h: 'SKU', k: 'sku' }, { h: 'Product', k: 'name' }, { h: 'Cost per carton (Rp)', k: 'cost', num: true }],
          list: () => r.shuffle(H.PRODUCTS).map((p) => ({ sku: p.sku, name: p.name, cost: Math.round(p.carton_price * 0.72 / 100) * 100 })), step: 2500, qstep: 20,
          task: 'multiply the cartons <b>On hand</b> by the <b>Cost per carton (Rp)</b> of its <b>SKU</b> in the cost file' },
        { title: 'Monthly interest with rates from the bank rate card', file: 'buku_pinjaman.xlsx', sheetName: 'Pinjaman', mine: 'the loan book', who: 'Dimas', ans: 'Interest (Rp)', round: true,
          main: [{ h: 'Facility', k: 'facility' }, { h: 'Distributor', k: 'distributor' }, { h: 'Type', k: 'type' }, { h: 'Outstanding (Rp)', k: 'outstanding', num: true }], rows: () => H.loans(r, 14),
          book: 'suku_bunga.xlsx', bsheet: 'Bunga', bdesc: 'the rate card',
          side: [{ h: 'Type', k: 'type' }, { h: 'Rate (% a year)', k: 'rate', num: true }],
          list: () => [{ type: 'Working Capital', rate: r.pick([10, 10.5, 11]) }, { type: 'Term Loan', rate: r.pick([11.5, 12, 12.5]) }, { type: 'Trade Finance', rate: r.pick([9, 9.5]) }], step: 0.5, qstep: 2000000,
          task: 'work out one month of interest: the <b>Outstanding (Rp)</b> balance times the yearly rate of its loan <b>Type</b> in the rate card, divided by 100 and by 12, rounded to whole rupiah with ROUND' }
      ][v];
      const rows = V.rows(), list = V.list();
      const S = H.table(V.main, rows), M = H.table(V.side, list);
      const ex = X(V.book, V.bsheet, M);
      const kH = V.main[2].h, qH = V.main[3].h, sk = V.side[0], sv = V.side[V.side.length - 1];
      const key0 = rows[0][V.main[2].k];
      /* The hidden test: the other workbook updates the rate of row 5's code, and one quantity is corrected. */
      const mAlt = H.edit(M.cells, { [M.cell(sv.h, list.findIndex((x) => x[sk.k] === key0))]: (x) => Math.round((x + V.step) * 100) / 100 });
      const alt = H.edit(S.cells, { [S.cell(qH, 5)]: (x) => x + V.qstep });
      const KA = ex.A(sk.h), VA = ex.A(sv.h), TA = ex.T(sk.h, sv.h);
      const xl = 'XLOOKUP(C5,' + KA + ',' + VA + ')', vl = 'VLOOKUP(C5,' + TA + ',' + V.side.length + ',FALSE)', im = 'INDEX(' + VA + ',MATCH(C5,' + KA + ',0))';
      const wrap = (x) => (V.round ? '=ROUND(D5*' + x + '/100/12,0)' : '=D5*' + x);
      return H.problem({
        title: V.title, file: V.file, sheetName: V.sheetName, sheet: S, base: S.cells, alt, fill: true, fillCol: 'E', ansLabel: V.ans,
        books: [wb(V.book, V.bsheet, M, mAlt)],
        statement: `<p>${opens([[V.mine, V.file], [V.bdesc, V.book]])} ${V.who} needs the <b>${V.ans}</b> of every row. In <b>E5</b>, ${V.task}. Then fill it down to <b>E${S.last}</b>.</p>`,
        solution: wrap(xl),
        alts: [wrap(vl), wrap(im), V.round ? '=ROUND(D5*' + vl + '/1200,0)' : '=' + xl + '*D5'],
        hints: [`Look up the ${kH} in C5 in ${V.book}, then work with the ${qH} in D5.`,
          `Open the ${V.book} tab and tap a column: Kicau writes it as Excel does, pinned with $, such as ${KA}.`,
          `Shape: ${wrap('XLOOKUP(C5, ' + KA + ', ' + VA + ')')}`],
        explain: `XLOOKUP finds the ${kH} from C5 in ${KA}, in the other workbook, and returns the ${sv.h} on the same row. ${V.round ? 'Times D5, divided by 100 and by 12 and rounded, that is one month of interest.' : 'Times D5 that gives the row\'s ' + V.ans + '.'} The ranges in ${V.book} are pinned with $, as Excel writes them, so filling down keeps them in place.`,
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
    key: 'lookup-with-a-fallback', difficulty: 'Medium', level: 5, topics: ['XLOOKUP', 'IFERROR', 'IFNA', 'Several workbooks'], count: 4, multi: true,
    make(r, v) {
      /* Some keys are missing from the list in the other workbook. The lookup falls back to a value instead of #N/A,
         and the math runs on whatever comes back. */
      let p;
      if (v === 0) {
        const prods = r.shuffle(H.PRODUCTS), missing = prods.slice(10), list = prods.slice(0, 10);
        const rows = H.sales(r, 14); rows[2].sku = missing[0].sku; rows[8].sku = missing[1].sku;
        const S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'Distributor', k: 'distributor' }, { h: 'SKU', k: 'sku' }, { h: 'Cartons', k: 'cartons', num: true }], rows);
        const M = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Product', k: 'name' }, { h: 'Carton price (Rp)', k: 'carton_price', num: true }], list);
        const ex = X('daftar_harga.xlsx', 'Harga', M);
        const listed = (k) => list.findIndex((x) => x.sku === rows[k].sku);
        const k0 = nums(14).find((k) => listed(k) >= 0), k4 = [4, 5, 6].find((k) => listed(k) >= 0);
        /* The hidden test: the price list updates one price, and one more line turns out to be a new SKU. */
        const mAlt = H.edit(M.cells, { [M.cell('Carton price (Rp)', listed(k0))]: (x) => x + 5000 });
        const alt = H.edit(S.cells, k4 === undefined ? {} : { [S.cell('SKU', k4)]: missing[1].sku });
        const KA = ex.A('SKU'), VA = ex.A('Carton price (Rp)'), TA = ex.T('SKU', 'Carton price (Rp)');
        p = { title: 'Order value from another price list, 0 for unlisted SKUs', file: 'pesanan_q4.xlsx', sheetName: 'Pesanan', S, alt, books: [wb('daftar_harga.xlsx', 'Harga', M, mAlt)], ans: 'Value (Rp)', KA, VA,
          statement: `<p>${opens([['the order lines', 'pesanan_q4.xlsx'], ['the price list', 'daftar_harga.xlsx']])} Two of the SKUs ordered are not on this price list yet.</p><p>In <b>E5</b>, work out each line's <b>Value (Rp)</b>: its <b>Cartons</b> times the <b>Carton price (Rp)</b> of its SKU in the price list. When the SKU is not on the list, the value is <b>0</b> instead of #N/A. Fill it down to <b>E${S.last}</b>.</p>`,
          solution: `=D5*XLOOKUP(C5,${KA},${VA},0)`,
          alts: [`=D5*IFERROR(VLOOKUP(C5,${TA},3,FALSE),0)`, `=IFNA(D5*INDEX(${VA},MATCH(C5,${KA},0)),0)`, `=D5*IFNA(VLOOKUP(C5,${TA},3,FALSE),0)`],
          fb: '0', shape: `=D5*XLOOKUP(C5, ${KA}, ${VA}, 0)`, chips: ['D5*', 'C5', KA, VA, TA, ',0)'],
          explain: `XLOOKUP's fourth part, if_not_found, is returned when C5 is not in ${KA}, so an unlisted SKU gives 0 and the line's value is 0. IFERROR or IFNA around VLOOKUP gives the same result.` };
      } else if (v === 1) {
        const fac = WC();
        const drop = r.sample(nums(fac.length), 4);
        const list = r.shuffle(fac.filter((f, k) => !drop.includes(k))).map((f) => ({ d: f.distributor, fac: f.facility_id, avail: Math.max(5000000, f.limit - f.outstanding) }));
        const rows = r.shuffle(DIST()).map((d) => {
          const l = list.find((x) => x.d === d.name);
          return { id: d.id, name: d.name, city: d.city, order: l ? Math.max(5, Math.round(l.avail * r.pick([0.5, 0.8, 1.25, 1.6]) / 1000000)) * 1000000 : r.int(10, 80) * 1000000 };
        });
        const S = H.table([{ h: 'ID', k: 'id' }, { h: 'Distributor', k: 'name' }, { h: 'City', k: 'city' }, { h: 'Order (Rp)', k: 'order', num: true }], rows);
        const M = H.table([{ h: 'Distributor', k: 'd' }, { h: 'Facility', k: 'fac' }, { h: 'Available (Rp)', k: 'avail', num: true }], list);
        const ex = X('limit_kredit.xlsx', 'Limit', M);
        const k0 = rows.findIndex((x) => list.some((l) => l.d === x.name)), w = list.findIndex((l) => l.d === rows[k0].name);
        const a0 = list[w].avail, o0 = rows[k0].order;
        /* The hidden test: the bank revises one available limit so that order's cover changes, and one order is corrected. */
        const mAlt = H.edit(M.cells, { [M.cell('Available (Rp)', w)]: a0 >= o0 ? Math.round(o0 / 2000000) * 1000000 : o0 + 5000000 });
        const alt = H.edit(S.cells, { [S.cell('Order (Rp)', (k0 + 3) % rows.length)]: (x) => x + 7000000 });
        const KA = ex.A('Distributor'), VA = ex.A('Available (Rp)'), TA = ex.T('Distributor', 'Available (Rp)');
        p = { title: 'Order amount covered by the available limit in the bank file', file: 'pesanan_minggu.xlsx', sheetName: 'Pesanan', S, alt, books: [wb('limit_kredit.xlsx', 'Limit', M, mAlt)], ans: 'Covered (Rp)', KA, VA,
          statement: `<p>${opens([['this week\'s orders', 'pesanan_minggu.xlsx'], ['Bank Kenari\'s limit file', 'limit_kredit.xlsx']])} The bank finances each order up to the distributor's <b>Available (Rp)</b> limit. A distributor without a facility is not in the limit file.</p><p>In <b>E5</b>, work out the <b>Covered (Rp)</b> part of each order: the smaller of its <b>Order (Rp)</b> and the available limit, with <b>0</b> as the limit when the distributor is not listed. Fill it down to <b>E${S.last}</b>.</p>`,
          solution: `=MIN(D5,XLOOKUP(B5,${KA},${VA},0))`,
          alts: [`=MIN(D5,IFERROR(VLOOKUP(B5,${TA},3,FALSE),0))`, `=IFNA(MIN(D5,INDEX(${VA},MATCH(B5,${KA},0))),0)`, `=MIN(XLOOKUP(B5,${KA},${VA},0),D5)`],
          fb: '0', shape: `=MIN(D5, XLOOKUP(B5, ${KA}, ${VA}, 0))`, chips: ['MIN(D5,', 'B5', KA, VA, TA, ',0)'],
          explain: `XLOOKUP brings the available limit from limit_kredit.xlsx, or 0 when the distributor has no facility. MIN then caps the order at that limit: an order below the limit is covered in full, one above it only up to the limit.` };
      } else if (v === 2) {
        const outs = r.sample(H.db().outlets, 15);
        const list = outs.slice(0, 12).sort((a, b) => (a.outlet_id < b.outlet_id ? -1 : 1)).map((o) => ({ id: o.outlet_id, name: o.outlet, rate: r.int(6, 20) * 100 }));
        const strangers = outs.slice(12);
        const rows = H.sales(r, 14).map((s, k) => ({ date: s.date, id: [3, 7, 11].includes(k) ? strangers[[3, 7, 11].indexOf(k)].outlet_id : r.pick(list).id, rep: r.pick(H.REPS), cartons: r.int(2, 20) }));
        const S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'Outlet ID', k: 'id' }, { h: 'Rep', k: 'rep' }, { h: 'Cartons', k: 'cartons', num: true }], rows);
        const M = H.table([{ h: 'Outlet ID', k: 'id' }, { h: 'Outlet', k: 'name' }, { h: 'Rate (Rp per carton)', k: 'rate', num: true }], list);
        const std = r.pick([2200, 2500, 2800]);
        const ST = H.table([{ h: 'Setting', k: 's' }, { h: 'Value', k: 'v', num: true }], [{ s: 'Standard rate (Rp per carton)', v: std }]);
        const ex = X('tarif_kirim.xlsx', 'Outlet', M), SR = X(null, 'Setup', ST).abs('Value', 0);
        /* The hidden test: the rate card raises one outlet's rate, the standard rate on the Setup sheet goes up, and two
           visits are corrected to other outlets. */
        const mAlt = H.edit(M.cells, { [M.cell('Rate (Rp per carton)', 0)]: (x) => x + 200 });
        const stAlt = H.edit(ST.cells, { [ST.cell('Value', 0)]: std + 300 });
        const alt = H.edit(S.cells, { [S.cell('Outlet ID', 3)]: list[0].id, [S.cell('Outlet ID', 0)]: strangers[1].outlet_id });
        const KA = ex.A('Outlet ID'), VA = ex.A('Rate (Rp per carton)'), TA = ex.T('Outlet ID', 'Rate (Rp per carton)');
        p = { title: 'Delivery cost per visit, standard rate for outlets not listed', file: 'kunjungan_outlet.xlsx', sheetName: 'Kunjungan', S, alt, books: [wb('tarif_kirim.xlsx', 'Outlet', M, mAlt)],
          sheets: [{ name: 'Setup', sheet: ST, base: ST.cells, alt: stAlt }], ans: 'Delivery (Rp)', KA, VA,
          statement: `<p>Two workbooks are open: the visit log (<b>kunjungan_outlet.xlsx</b>, with the sheets Kunjungan and Setup) and the delivery rate card (<b>tarif_kirim.xlsx</b>). Reps also visit new outlets that are not on the rate card yet.</p><p>In <b>E5</b>, work out the <b>Delivery (Rp)</b> of each visit: its <b>Cartons</b> times the outlet's <b>Rate (Rp per carton)</b> from the rate card. For an outlet that is not on the card, use the standard rate on the Setup sheet, cell <b>B5</b>. Fill it down to <b>E${S.last}</b>.</p>`,
          solution: `=D5*XLOOKUP(B5,${KA},${VA},${SR})`,
          alts: [`=D5*IFERROR(VLOOKUP(B5,${TA},3,FALSE),${SR})`, `=D5*IFNA(INDEX(${VA},MATCH(B5,${KA},0)),${SR})`, `=IFNA(D5*XLOOKUP(B5,${KA},${VA}),D5*${SR})`],
          fb: SR, shape: `=D5*XLOOKUP(B5, ${KA}, ${VA}, ${SR})`, chips: ['D5*', 'B5', KA, VA, TA, SR],
          explain: `XLOOKUP's fourth part can be a cell, not only a fixed value. When B5 is not on the rate card, it returns the standard rate from ${SR}, a sheet of the same workbook, written without a file name. The $ signs keep every range in place as the formula fills down.` };
      } else {
        const prods = r.shuffle(H.PRODUCTS);
        const list = prods.slice(0, 4).map((x) => ({ sku: x.sku, name: x.name, pct: r.pick([5, 7.5, 10, 12.5, 15]) }));
        const priceOf = (sku) => H.PRODUCTS.find((x) => x.sku === sku).carton_price;
        const rows = H.sales(r, 14);
        rows[1].sku = list[0].sku; rows[5].sku = list[1].sku; rows[9].sku = list[2].sku;
        rows.forEach((x) => { x.gross = x.cartons * priceOf(x.sku); });
        const S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'SKU', k: 'sku' }, { h: 'Cartons', k: 'cartons', num: true }, { h: 'Gross (Rp)', k: 'gross', num: true }], rows);
        const M = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Product', k: 'name' }, { h: 'Promo (%)', k: 'pct', num: true }], list);
        const ex = X('promo_q4.xlsx', 'Promo', M);
        /* The hidden test: the promo file deepens one promo, and line 4 turns out to be another SKU. */
        const mAlt = H.edit(M.cells, { [M.cell('Promo (%)', 0)]: (x) => x + 2.5 });
        const k3 = list.some((x) => x.sku === rows[3].sku) ? prods[6].sku : list[3].sku;
        const alt = H.edit(S.cells, { [S.cell('SKU', 3)]: k3, [S.cell('Gross (Rp)', 3)]: rows[3].cartons * priceOf(k3) });
        const KA = ex.A('SKU'), VA = ex.A('Promo (%)'), TA = ex.T('SKU', 'Promo (%)');
        p = { title: 'Promo discount from the promo workbook, 0 without a promo', file: 'penjualan_q4.xlsx', sheetName: 'Penjualan', S, alt, books: [wb('promo_q4.xlsx', 'Promo', M, mAlt)], ans: 'Discount (Rp)', KA, VA,
          statement: `<p>${opens([['the sales lines', 'penjualan_q4.xlsx'], ['the Q4 promos', 'promo_q4.xlsx']])} Only four SKUs are on promo.</p><p>In <b>E5</b>, work out the promo discount in rupiah: the line's <b>Gross (Rp)</b> times the <b>Promo (%)</b> of its SKU, divided by 100 and rounded to whole rupiah with ROUND. A SKU that is not in the promo file has a promo of <b>0</b>. Fill it down to <b>E${S.last}</b>.</p>`,
          solution: `=ROUND(D5*XLOOKUP(B5,${KA},${VA},0)/100,0)`,
          alts: [`=ROUND(D5*IFERROR(VLOOKUP(B5,${TA},3,FALSE),0)/100,0)`, `=IFERROR(ROUND(D5*VLOOKUP(B5,${TA},3,FALSE)/100,0),0)`, `=ROUND(D5*IFNA(INDEX(${VA},MATCH(B5,${KA},0)),0)/100,0)`],
          fb: '0', shape: `=ROUND(D5*XLOOKUP(B5, ${KA}, ${VA}, 0)/100, 0)`, chips: ['D5*', 'B5', KA, VA, TA, ',0)'],
          explain: `XLOOKUP returns the SKU's promo from promo_q4.xlsx, or 0 when the SKU has no promo. Gross times that percentage, divided by 100 and rounded, is the discount in rupiah, so a line without a promo gets 0.` };
      }
      return H.problem({
        title: p.title, file: p.file, sheetName: p.sheetName, sheet: p.S, base: p.S.cells, alt: p.alt, fill: true, fillCol: 'E', ansLabel: p.ans, books: p.books, sheets: p.sheets,
        statement: p.statement, solution: p.solution, alts: p.alts,
        hints: ['A plain lookup gives #N/A for a key that is not on the list. XLOOKUP has a fourth part for that case: what to return when nothing is found.',
          `The list sits in another workbook: pin it with $ (${p.KA} and ${p.VA}). IFERROR(VLOOKUP(...), ${p.fb}) works too.`,
          `Shape: ${p.shape}`],
        explain: p.explain, chips: p.chips, fnKey: 'XLOOKUP('
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
    key: 'check-against-the-master', difficulty: 'Medium', level: 7, topics: ['IF', 'XLOOKUP', 'VLOOKUP', 'Several workbooks'], count: 4, multi: true,
    make(r, v) {
      /* Compare each row with a figure looked up in a master kept in another workbook, and write one of two words. */
      if (v === 0) {
        const rows = H.sales(r, 14);
        const list = r.shuffle(H.PRODUCTS).map((x) => ({ sku: x.sku, name: x.name, oh: r.int(2, 7) * 10 }));
        const S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'Distributor', k: 'distributor' }, { h: 'SKU', k: 'sku' }, { h: 'Cartons', k: 'cartons', num: true }], rows);
        const M = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Product', k: 'name' }, { h: 'On hand', k: 'oh', num: true }], list);
        const ex = X('stok_gudang.xlsx', 'Stok', M);
        const w = list.findIndex((x) => x.sku === rows[0].sku), c0 = rows[0].cartons;
        /* The hidden test: the stock count for row 5's SKU is recounted, so that line flips. */
        const mAlt = H.edit(M.cells, { [M.cell('On hand', w)]: c0 > list[w].oh ? c0 : c0 - 1 });
        const KA = ex.A('SKU'), VA = ex.A('On hand'), TA = ex.T('SKU', 'On hand');
        return H.problem({
          title: 'Mark order lines short of stock in the warehouse workbook', file: 'pesanan_q4.xlsx', sheetName: 'Pesanan', sheet: S, base: S.cells, alt: S.cells, fill: true, fillCol: 'E', ansLabel: 'Stock check',
          books: [wb('stok_gudang.xlsx', 'Stok', M, mAlt)],
          statement: `<p>${opens([['the order lines', 'pesanan_q4.xlsx'], ['the warehouse stock count', 'stok_gudang.xlsx']])} In <b>E5</b>, show <b>Short</b> when the line's <b>Cartons</b> are more than the <b>On hand</b> for its SKU in the stock count, and <b>OK</b> otherwise (equal is OK). Fill it down to <b>E${S.last}</b>.</p>`,
          solution: `=IF(D5>XLOOKUP(C5,${KA},${VA}),"Short","OK")`,
          alts: [`=IF(D5<=VLOOKUP(C5,${TA},3,FALSE),"OK","Short")`, `=IF(INDEX(${VA},MATCH(C5,${KA},0))<D5,"Short","OK")`],
          hints: ['First look up the On hand for the SKU in C5 in stok_gudang.xlsx, then compare it with D5 inside IF.', `Pin the stock list: ${KA} and ${VA}. "More than" is >.`, `Shape: =IF(D5 > XLOOKUP(C5, ${KA}, ${VA}), "Short", "OK")`],
          explain: 'XLOOKUP returns the stock for the line\'s SKU from the other workbook, and IF compares the order with it. A line equal to the stock is still OK, because only more than the stock is short.',
          chips: ['D5', 'C5', KA, VA, TA, '"Short"', '"OK"'], fnKey: 'XLOOKUP('
        });
      }
      if (v === 1) {
        const rows = H.sales(r, 14).map((s) => ({ invoice: s.invoice, sku: s.sku, cartons: s.cartons, charged: s.price }));
        r.sample(nums(14), 3).forEach((k) => { rows[k].charged += r.pick([-5000, -3000, 2000, 4000]); });
        const list = r.shuffle(H.PRODUCTS);
        const S = H.table([{ h: 'Invoice', k: 'invoice' }, { h: 'SKU', k: 'sku' }, { h: 'Cartons', k: 'cartons', num: true }, { h: 'Price charged (Rp)', k: 'charged', num: true }], rows);
        const M = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Product', k: 'name' }, { h: 'List price (Rp)', k: 'carton_price', num: true }], list);
        const ex = X('daftar_harga.xlsx', 'Harga', M);
        const lp = H.PRODUCTS.find((x) => x.sku === rows[0].sku).carton_price, w = list.findIndex((x) => x.sku === rows[0].sku);
        /* The hidden test: the price list is revised for row 5's SKU, so lines that matched the old price now differ,
           or an off-price line now matches. */
        const mAlt = H.edit(M.cells, { [M.cell('List price (Rp)', w)]: rows[0].charged === lp ? lp + 1000 : rows[0].charged });
        const KA = ex.A('SKU'), VA = ex.A('List price (Rp)'), TA = ex.T('SKU', 'List price (Rp)');
        return H.problem({
          title: 'Flag lines charged off the list in the price workbook', file: 'audit_harga.xlsx', sheetName: 'Audit', sheet: S, base: S.cells, alt: S.cells, fill: true, fillCol: 'E', ansLabel: 'Price check',
          books: [wb('daftar_harga.xlsx', 'Harga', M, mAlt)],
          statement: `<p>${opens([['the invoiced lines', 'audit_harga.xlsx'], ['the price list', 'daftar_harga.xlsx']])} In <b>E5</b>, show <b>Check</b> when the <b>Price charged (Rp)</b> is different from the <b>List price (Rp)</b> of the line's SKU in the price list, and <b>OK</b> when they are the same. Fill it down to <b>E${S.last}</b>.</p>`,
          solution: `=IF(D5<>XLOOKUP(B5,${KA},${VA}),"Check","OK")`,
          alts: [`=IF(D5=VLOOKUP(B5,${TA},3,FALSE),"OK","Check")`, `=IF(D5=INDEX(${VA},MATCH(B5,${KA},0)),"OK","Check")`],
          hints: ['Look up the list price for the SKU in B5 in daftar_harga.xlsx, then compare it with the price in D5.', '<> means "is not equal to". Pin the price list with $.', `Shape: =IF(D5 <> XLOOKUP(B5, ${KA}, ${VA}), "Check", "OK")`],
          explain: 'XLOOKUP brings the list price for the line\'s SKU from the price workbook, and IF compares it with what was charged. Any difference, up or down, is flagged. When the list price is revised, the flags follow.',
          chips: ['D5', 'B5', KA, VA, TA, '"Check"', '"OK"'], fnKey: 'XLOOKUP('
        });
      }
      if (v === 2) {
        const rows = H.reps(r, 10);
        const list = r.shuffle(rows).map((x) => ({ name: x.name, target: x.target }));
        const S = H.table([{ h: 'Rep', k: 'name' }, { h: 'Area', k: 'area' }, { h: 'Actual (Rp)', k: 'actual', num: true }], rows);
        const M = H.table([{ h: 'Rep', k: 'name' }, { h: 'Target (Rp)', k: 'target', num: true }], list);
        const ex = X('target_sales.xlsx', 'Target', M);
        const w = list.findIndex((x) => x.name === rows[0].name);
        /* The hidden test: the first rep's target is revised in the targets file, so their result flips. */
        const mAlt = H.edit(M.cells, { [M.cell('Target (Rp)', w)]: rows[0].actual >= rows[0].target ? rows[0].actual + 100000 : rows[0].actual });
        const KA = ex.A('Rep'), VA = ex.A('Target (Rp)'), TA = ex.T('Rep', 'Target (Rp)');
        return H.problem({
          title: 'Hit or miss against targets kept in another workbook', file: 'realisasi_sales.xlsx', sheetName: 'Realisasi', sheet: S, base: S.cells, alt: S.cells, fill: true, fillCol: 'D', ansLabel: 'Result',
          books: [wb('target_sales.xlsx', 'Target', M, mAlt)],
          statement: `<p>${opens([['the actual sales', 'realisasi_sales.xlsx'], ['the targets', 'target_sales.xlsx']])} The targets file lists the reps in a different order. In <b>D5</b>, show <b>Hit</b> when the rep's <b>Actual (Rp)</b> is at least their target, and <b>Miss</b> when it is below. Fill it down to <b>D${S.last}</b>.</p>`,
          solution: `=IF(C5>=XLOOKUP(A5,${KA},${VA}),"Hit","Miss")`,
          alts: [`=IF(C5<VLOOKUP(A5,${TA},2,FALSE),"Miss","Hit")`, `=IF(C5>=INDEX(${VA},MATCH(A5,${KA},0)),"Hit","Miss")`],
          hints: ['The two files list the reps in different orders, so look the target up by name rather than reading the same row.', `"At least" is >=. Pin the target range in the other workbook: ${KA} and ${VA}.`, `Shape: =IF(C5 >= XLOOKUP(A5, ${KA}, ${VA}), "Hit", "Miss")`],
          explain: 'XLOOKUP finds each rep\'s own target by name in target_sales.xlsx, wherever it sits. IF then compares it with the actual; reaching the target exactly counts as a hit.',
          chips: ['A5', 'C5', KA, VA, TA, '"Hit"', '"Miss"'], fnKey: 'XLOOKUP('
        });
      }
      const fac = r.shuffle(WC());
      const ds = r.sample(fac, 12);
      const rows = ds.map((f, k) => ({ req: 'REQ-' + (41 + k * 3 + r.int(0, 2)), d: f.distributor, amt: Math.max(1000000, f.limit - f.outstanding + r.pick([-6, -3, -1, 1, 2, 5]) * 1000000) }));
      /* One request is exactly the headroom, so "at most" against "less than" shows on the visible sheet. */
      rows[3].amt = ds[3].limit - ds[3].outstanding;
      const S = H.table([{ h: 'Request', k: 'req' }, { h: 'Distributor', k: 'd' }, { h: 'Amount (Rp)', k: 'amt', num: true }], rows);
      const M = H.table([{ h: 'Distributor', k: 'distributor' }, { h: 'Limit (Rp)', k: 'limit', num: true }, { h: 'Outstanding (Rp)', k: 'outstanding', num: true }], fac);
      const ex = X('buku_pinjaman.xlsx', 'Fasilitas', M);
      const f0 = fac.findIndex((x) => x.distributor === rows[0].d), head0 = fac[f0].limit - fac[f0].outstanding;
      /* The hidden test flips row 5 in the loan book: a request within its headroom meets an outstanding balance that
         leaves 100,000 less than it asks for, and one over it meets a balance that leaves exactly enough. */
      const mAlt = H.edit(M.cells, { [M.cell('Outstanding (Rp)', f0)]: Math.max(0, fac[f0].limit - rows[0].amt + (rows[0].amt <= head0 ? 100000 : 0)) });
      const KA = ex.A('Distributor'), LA = ex.A('Limit (Rp)'), OA = ex.A('Outstanding (Rp)'), TA = ex.T('Distributor', 'Outstanding (Rp)');
      return H.problem({
        title: 'Drawdown requests checked against headroom in the loan book', file: 'permintaan_pencairan.xlsx', sheetName: 'Permintaan', sheet: S, base: S.cells, alt: S.cells, fill: true, fillCol: 'D', ansLabel: 'Decision',
        books: [wb('buku_pinjaman.xlsx', 'Fasilitas', M, mAlt)],
        statement: `<p>${opens([['the drawdown requests', 'permintaan_pencairan.xlsx'], ['the loan book', 'buku_pinjaman.xlsx']])} Each request asks to draw more money from a facility. In <b>D5</b>, show <b>Within</b> when the request's <b>Amount (Rp)</b> is at most the distributor's headroom (its <b>Limit (Rp)</b> minus its <b>Outstanding (Rp)</b> in the loan book), and <b>Over</b> when it is more. Fill it down to <b>D${S.last}</b>.</p>`,
        solution: `=IF(C5<=XLOOKUP(B5,${KA},${LA})-XLOOKUP(B5,${KA},${OA}),"Within","Over")`,
        alts: [`=IF(C5>VLOOKUP(B5,${TA},2,FALSE)-VLOOKUP(B5,${TA},3,FALSE),"Over","Within")`, `=IF(C5<=INDEX(${LA},MATCH(B5,${KA},0))-INDEX(${OA},MATCH(B5,${KA},0)),"Within","Over")`, `=IF(C5<=XLOOKUP(B5,${KA},${LA}-${OA}),"Within","Over")`],
        hints: ['Headroom needs two lookups in buku_pinjaman.xlsx for the same distributor: the limit and the outstanding balance.', `"At most" is <=. Pin the ranges in the loan book: ${KA}, ${LA} and ${OA}.`, `Shape: =IF(C5 <= XLOOKUP(B5, ${KA}, ${LA}) - XLOOKUP(B5, ${KA}, ${OA}), "Within", "Over")`],
        explain: 'The two XLOOKUPs bring the distributor\'s limit and outstanding balance from the loan book; their difference is the headroom. IF compares the request with it, and a request equal to the headroom still fits.',
        chips: ['B5', 'C5', KA, LA, OA, TA, '"Within"', '"Over"'], fnKey: 'XLOOKUP('
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

  /* ======================= Hard: two or more lookups across three or four workbooks ======================= */
  P.add({
    key: 'two-way-grid-lookup', difficulty: 'Hard', level: 8, topics: ['INDEX', 'MATCH', 'XLOOKUP', 'Several workbooks'], count: 4, multi: true,
    make(r, v) {
      /* A rate card in another workbook is a grid: one key down column A, the other across row 4. Each line first looks
         up what the grid needs (a region, a loan type) or what the grid value is divided by (a pack size), then reads
         the grid with INDEX and two MATCHes. */
      if (v === 0) {
        const WH = H.WAREHOUSES, regs = H.REGIONS.slice();
        const rows = H.sales(r, 14).map((s) => ({ date: s.date, d: s.distributor, wh: r.pick(WH), cartons: s.cartons }));
        const S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'Distributor', k: 'd' }, { h: 'Warehouse', k: 'wh' }, { h: 'Cartons', k: 'cartons', num: true }], rows);
        const MD = H.table([DC.id, DC.name, DC.city, DC.region], r.shuffle(DIST()));
        /* Rp per carton grows with the distance from the warehouse to the region, placed roughly west to east. */
        const WP = { 'Gudang Medan': 0, 'Gudang Bekasi': 3, 'Gudang Bandung': 4, 'Gudang Semarang': 6, 'Gudang Surabaya': 8 };
        const RP = { Sumatera: 1, Jabodetabek: 3, 'Jawa Barat': 4, 'Jawa Tengah': 6, 'Jawa Timur': 8, 'Bali Nusra': 10, Sulawesi: 12 };
        const TG = across('Region', WH, regs.map((h) => ({ h, vals: WH.map((w) => 2000 + 1000 * Math.abs(WP[w] - RP[h]) + r.int(0, 2) * 500) })));
        const md = X('master_distributor.xlsx', 'Distributor', MD), g = gridRefs('tarif_kirim.xlsx', 'Tarif', TG);
        /* The hidden test: the rate card raises the rate on row 5's route, and row 11 ships from another warehouse. */
        const tgAlt = H.edit(TG.cells, { [L[WH.indexOf(rows[0].wh) + 1] + (5 + regs.indexOf(regionOf(rows[0].d)))]: (x) => x + 1500 });
        const alt = H.edit(S.cells, { [S.cell('Warehouse', 6)]: WH[(WH.indexOf(rows[6].wh) + 2) % WH.length] });
        const reg = `XLOOKUP(B5,${md.A('Distributor')},${md.A('Region')})`;
        return H.problem({
          title: 'Freight per shipment from a region by warehouse rate card', file: 'pengiriman_q4.xlsx', sheetName: 'Kirim', sheet: S, base: S.cells, alt, fill: true, fillCol: 'E', ansLabel: 'Freight (Rp)',
          books: [wb('master_distributor.xlsx', 'Distributor', MD), wb('tarif_kirim.xlsx', 'Tarif', TG, tgAlt)],
          statement: `<p>${opens([['the shipments', 'pengiriman_q4.xlsx'], ['the distributor master', 'master_distributor.xlsx'], ['the freight rate card', 'tarif_kirim.xlsx']])} The rate card gives Rp per carton, with regions down column A and warehouses across row 4. The shipments do not show the region.</p><p>In <b>E5</b>, work out the freight of each shipment: its <b>Cartons</b> times the rate where the distributor's <b>Region</b> (from the master) meets the shipping <b>Warehouse</b>. Fill it down to <b>E${S.last}</b>.</p>`,
          solution: `=D5*INDEX(${g.G},MATCH(${reg},${g.RK},0),MATCH(C5,${g.CK},0))`,
          alts: [`=D5*XLOOKUP(${reg},${g.RK},XLOOKUP(C5,${g.CK},${g.G}))`, `=D5*INDEX(${g.G},MATCH(VLOOKUP(B5,${md.T('Distributor', 'Region')},3,FALSE),${g.RK},0),MATCH(C5,${g.CK},0))`,
            `=D5*VLOOKUP(${reg},${g.ALL},MATCH(C5,${g.HEAD},0),FALSE)`, `=INDEX(${g.G},XMATCH(${reg},${g.RK}),XMATCH(C5,${g.CK}))*D5`],
          hints: ['Look up the region first: the distributor in B5 gives its Region in master_distributor.xlsx.',
            `Then read the grid: MATCH that region in ${g.RK} for the row, and MATCH the warehouse in C5 in ${g.CK} for the column. INDEX returns the cell where they cross.`,
            `Shape: =D5*INDEX(${g.G}, MATCH(XLOOKUP(B5, ...), ${g.RK}, 0), MATCH(C5, ${g.CK}, 0))`],
          explain: 'XLOOKUP turns the distributor into its region from master_distributor.xlsx. The first MATCH finds that region down the rate card, the second finds the warehouse across it, and INDEX returns the rate where they cross. Times the cartons, that is the freight. Every range in the other files is pinned with $, so filling down keeps them in place.',
          chips: ['D5*', 'B5', 'C5', md.A('Distributor'), md.A('Region'), g.G, g.RK, g.CK, ',0)'], fnKey: 'INDEX('
        });
      }
      if (v === 1) {
        const facs = H.db().facilities;
        const rows = r.sample(facs, 12).map((f) => ({ f: f.facility_id, d: f.distributor, tenor: r.pick([1, 2, 3, 5]), amt: r.pick([25, 40, 50, 75, 100, 150]) * 1000000 }));
        const S = H.table([{ h: 'Facility', k: 'f' }, { h: 'Distributor', k: 'd' }, { h: 'Tenor (years)', k: 'tenor', num: true }, { h: 'Amount (Rp)', k: 'amt', num: true }], rows);
        const BP = H.table([{ h: 'Facility', k: 'facility_id' }, { h: 'Distributor', k: 'distributor' }, { h: 'Type', k: 'type' }, { h: 'Limit (Rp)', k: 'limit', num: true }], r.shuffle(facs));
        const types = ['Working Capital', 'Term Loan', 'Trade Finance', 'Investment'], ten = [1, 2, 3, 5];
        const b0 = r.pick([8.5, 9]), add = [0, 1, -0.5, 1.5], up = [0, 0.25, 0.5, 1];
        const RC = across('Type', ten, types.map((h, i) => ({ h, vals: ten.map((t, j) => b0 + add[i] + up[j] + r.pick([0, 0.25])) })));
        const bp = X('buku_pinjaman.xlsx', 'Fasilitas', BP), g = gridRefs('suku_bunga.xlsx', 'Bunga', RC);
        const typeOf = (id) => facs.find((x) => x.facility_id === id).type;
        /* The hidden test: the rate card moves the rate for row 5's type and tenor, and row 10 gets a longer tenor. */
        const rcAlt = H.edit(RC.cells, { [L[ten.indexOf(rows[0].tenor) + 1] + (5 + types.indexOf(typeOf(rows[0].f)))]: (x) => x + 0.25 });
        const alt = H.edit(S.cells, { [S.cell('Tenor (years)', 5)]: ten[(ten.indexOf(rows[5].tenor) + 1) % 4] });
        const ty = `XLOOKUP(A5,${bp.A('Facility')},${bp.A('Type')})`;
        const wrap = (x) => '=ROUND(D5*' + x + '/100,0)';
        return H.problem({
          title: 'First-year interest from a loan type by tenor rate card', file: 'pencairan_kredit.xlsx', sheetName: 'Pencairan', sheet: S, base: S.cells, alt, fill: true, fillCol: 'E', ansLabel: 'Interest year 1 (Rp)',
          books: [wb('buku_pinjaman.xlsx', 'Fasilitas', BP), wb('suku_bunga.xlsx', 'Bunga', RC, rcAlt)],
          statement: `<p>${opens([['the new drawdowns', 'pencairan_kredit.xlsx'], ['the loan book', 'buku_pinjaman.xlsx'], ['the rate card', 'suku_bunga.xlsx']])} The rate card gives the yearly rate (%) with loan types down column A and tenors in years across row 4. The drawdowns show the facility, not its type.</p><p>In <b>E5</b>, work out the first year of interest on each drawdown: its <b>Amount (Rp)</b> times the rate for its facility's <b>Type</b> (from the loan book) and its <b>Tenor (years)</b>, divided by 100 and rounded to whole rupiah with ROUND. Fill it down to <b>E${S.last}</b>.</p>`,
          solution: wrap(`INDEX(${g.G},MATCH(${ty},${g.RK},0),MATCH(C5,${g.CK},0))`),
          alts: [wrap(`XLOOKUP(${ty},${g.RK},XLOOKUP(C5,${g.CK},${g.G}))`), wrap(`VLOOKUP(${ty},${g.ALL},MATCH(C5,${g.HEAD},0),FALSE)`),
            wrap(`INDEX(${g.G},MATCH(VLOOKUP(A5,${bp.T('Facility', 'Type')},3,FALSE),${g.RK},0),MATCH(C5,${g.CK},0))`)],
          hints: ['Two steps: the facility in A5 gives its loan Type in buku_pinjaman.xlsx, and the type with the tenor gives the rate in suku_bunga.xlsx.',
            `The row of the rate card is MATCH(type, ${g.RK}, 0) and the column is MATCH(C5, ${g.CK}, 0). INDEX on ${g.G} reads the rate.`,
            `Shape: =ROUND(D5*INDEX(${g.G}, MATCH(XLOOKUP(A5, ...), ${g.RK}, 0), MATCH(C5, ${g.CK}, 0))/100, 0)`],
          explain: 'XLOOKUP brings the facility\'s loan type from the loan book. That type and the tenor are the two keys of the rate card: one MATCH finds the row, the other the column, and INDEX returns the yearly rate. Amount times rate divided by 100 is one year of interest.',
          chips: ['D5*', 'A5', 'C5', bp.A('Facility'), bp.A('Type'), g.G, g.RK, g.CK, ',0)'], fnKey: 'INDEX('
        });
      }
      if (v === 2) {
        const months = ['May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct'];
        const reps = r.sample(H.REPS, 8), order = r.shuffle(reps);
        const sales = {};
        reps.forEach((n) => { sales[n] = months.map(() => r.int(40, 160) * 1000000); });
        const SG = across('Rep', months, order.map((h) => ({ h, vals: sales[h] })));
        const tl = r.shuffle(H.REPS).map((n) => ({ n, t: r.pick([80, 90, 100, 120, 150]) * 1000000 }));
        const TG = H.table([{ h: 'Rep', k: 'n' }, { h: 'Monthly target (Rp)', k: 't', num: true }], tl);
        const S = H.table([{ h: 'Rep', k: 'n' }, { h: 'Area', k: 'area' }], reps.map((n) => ({ n, area: r.pick(H.REGIONS) })));
        const i = r.int(0, 5), j = r.pick(nums(6).filter((k) => k !== i));
        const g = gridRefs('penjualan_rep.xlsx', 'Sales', SG), tg = X('target_sales.xlsx', 'Target', TG);
        const base = Object.assign({}, S.cells, { A1: 'Month', B1: months[i] });
        /* The hidden test: B1 asks for another month, one sale in that month is corrected, and one target is raised. */
        const alt = H.edit(base, { B1: months[j] });
        const sgAlt = H.edit(SG.cells, { [L[j + 1] + (5 + order.indexOf(reps[0]))]: (x) => x + 2500000 });
        const tgAlt = H.edit(TG.cells, { [TG.cell('Monthly target (Rp)', tl.findIndex((x) => x.n === reps[1]))]: (x) => x + 20000000 });
        const TR = tg.A('Rep'), TT = tg.A('Monthly target (Rp)'), targ = `XLOOKUP(A5,${TR},${TT})`;
        return H.problem({
          title: "Each rep's achievement in a chosen month, sales and targets apart", file: 'capaian_rep.xlsx', sheetName: 'Capaian', sheet: S, base, alt, fill: true, fillCol: 'C', ansLabel: 'Achievement', precision: 4,
          books: [wb('penjualan_rep.xlsx', 'Sales', SG, sgAlt), wb('target_sales.xlsx', 'Target', TG, tgAlt)],
          statement: `<p>${opens([['the rep list', 'capaian_rep.xlsx'], ['the sales by rep and month', 'penjualan_rep.xlsx'], ['the targets', 'target_sales.xlsx']])} In the sales file the reps run down column A and the months across row 4. <b>B1</b> holds the month to report.</p><p>In <b>C5</b>, work out each rep's achievement for that month: their sales in that month divided by their <b>Monthly target (Rp)</b> from the targets file, as a decimal (0.95 means 95%). Fill it down to <b>C${S.last}</b>.</p>`,
          solution: `=INDEX(${g.G},MATCH(A5,${g.RK},0),MATCH($B$1,${g.CK},0))/${targ}`,
          alts: [`=XLOOKUP(A5,${g.RK},XLOOKUP($B$1,${g.CK},${g.G}))/${targ}`, `=VLOOKUP(A5,${g.ALL},MATCH($B$1,${g.HEAD},0),FALSE)/VLOOKUP(A5,${tg.T('Rep', 'Monthly target (Rp)')},2,FALSE)`,
            `=INDEX(${g.G},XMATCH(A5,${g.RK}),XMATCH($B$1,${g.CK}))/INDEX(${TT},MATCH(A5,${TR},0))`],
          hints: ['The sale is a two-way lookup: the rep in A5 picks the row of the sales grid, the month in B1 picks the column. The target is a plain lookup in the other file.',
            'B1 must stay put as the formula fills down, so write it as $B$1. The ranges in the other workbooks are pinned already.',
            `Shape: =INDEX(${g.G}, MATCH(A5, ${g.RK}, 0), MATCH($B$1, ${g.CK}, 0)) / XLOOKUP(A5, ${TR}, ${TT})`],
          explain: 'INDEX with two MATCHes reads the rep\'s sales for the month in B1 from penjualan_rep.xlsx; XLOOKUP brings the same rep\'s target from target_sales.xlsx. Sales over target is the achievement. Change B1 and every row follows.',
          chips: ['A5', '$B$1', g.G, g.RK, g.CK, TR, TT, ',0)'], fnKey: 'INDEX('
        });
      }
      const CH = ['Wholesale', 'Modern trade', 'Warung'];
      const ps = r.shuffle(H.PRODUCTS);
      const HK = across('SKU', CH, ps.map((x) => ({ h: x.sku, vals: [x.carton_price, Math.round(x.carton_price * 1.08 / 1000) * 1000, Math.round(x.carton_price * 1.15 / 1000) * 1000] })));
      const pl = r.shuffle(H.PRODUCTS);
      const PM = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Product', k: 'name' }, { h: 'Category', k: 'category' }, { h: 'Packs per carton', k: 'packs_per_carton', num: true }], pl);
      const rows = H.sales(r, 14).map((s) => ({ inv: s.invoice, ch: r.pick(CH), sku: s.sku, packs: r.int(2, 20) * 6 }));
      const S = H.table([{ h: 'Invoice', k: 'inv' }, { h: 'Channel', k: 'ch' }, { h: 'SKU', k: 'sku' }, { h: 'Packs', k: 'packs', num: true }], rows);
      const g = gridRefs('harga_kanal.xlsx', 'Harga', HK), pm = X('master_produk.xlsx', 'Produk', PM);
      /* The hidden test: one channel price goes up, one SKU gets a bigger carton in the product master, and row 9 sells
         through another channel. */
      const hkAlt = H.edit(HK.cells, { [L[CH.indexOf(rows[0].ch) + 1] + (5 + ps.findIndex((x) => x.sku === rows[0].sku))]: (x) => x + 12000 });
      const pmAlt = H.edit(PM.cells, { [PM.cell('Packs per carton', pl.findIndex((x) => x.sku === rows[2].sku))]: (x) => x + 8 });
      const alt = H.edit(S.cells, { [S.cell('Channel', 4)]: CH[(CH.indexOf(rows[4].ch) + 1) % 3] });
      const price = `INDEX(${g.G},MATCH(C5,${g.RK},0),MATCH(B5,${g.CK},0))`, PS = pm.A('SKU'), PP = pm.A('Packs per carton'), packs = `XLOOKUP(C5,${PS},${PP})`;
      return H.problem({
        title: 'Pack order value from a channel price grid and the carton size', file: 'pesanan_kanal.xlsx', sheetName: 'Pesanan', sheet: S, base: S.cells, alt, fill: true, fillCol: 'E', ansLabel: 'Value (Rp)',
        books: [wb('harga_kanal.xlsx', 'Harga', HK, hkAlt), wb('master_produk.xlsx', 'Produk', PM, pmAlt)],
        statement: `<p>${opens([['the orders in packs', 'pesanan_kanal.xlsx'], ['the carton prices by channel', 'harga_kanal.xlsx'], ['the product master', 'master_produk.xlsx']])} The price grid is per carton, with SKUs down column A and sales channels across row 4. These orders are counted in packs, not cartons, whatever the channel.</p><p>In <b>E5</b>, work out the value of each order: its <b>Packs</b> times the carton price for its <b>SKU</b> and <b>Channel</b>, divided by the SKU's <b>Packs per carton</b> from the product master, rounded to whole rupiah with ROUND. Fill it down to <b>E${S.last}</b>.</p>`,
        solution: `=ROUND(D5*${price}/${packs},0)`,
        alts: [`=ROUND(D5*XLOOKUP(C5,${g.RK},XLOOKUP(B5,${g.CK},${g.G}))/${packs},0)`, `=ROUND(D5*VLOOKUP(C5,${g.ALL},MATCH(B5,${g.HEAD},0),FALSE)/VLOOKUP(C5,${pm.T('SKU', 'Packs per carton')},4,FALSE),0)`,
          `=ROUND(${price}*D5/${packs},0)`],
        hints: ['Two lookups in two files: a two-way one for the carton price (SKU down, channel across) and a plain one for the packs per carton.',
          `${price.replace(/,/g, ', ')} is the carton price; XLOOKUP(C5, ${PS}, ${PP}) is the packs per carton.`,
          'Shape: =ROUND(D5 * carton price / packs per carton, 0)'],
        explain: 'INDEX with two MATCHes picks the carton price where the SKU row meets the channel column in harga_kanal.xlsx. Divided by the packs per carton from master_produk.xlsx, that is the price of one pack; times the packs ordered and rounded, it is the order value.',
        chips: ['D5*', 'B5', 'C5', g.G, g.RK, g.CK, PS, PP, ',0)'], fnKey: 'INDEX('
      });
    }
  });

  P.add({
    key: 'tier-from-thresholds', difficulty: 'Hard', level: 8, topics: ['VLOOKUP', 'XLOOKUP', 'approximate match', 'Several workbooks'], count: 4, multi: true,
    make(r, v) {
      /* A tier table in another workbook, sorted by its lower limits: an approximate match returns the tier a value
         falls in. A second lookup, in a third workbook, brings what the tier is applied to or measured against. */
      if (v === 0) {
        const rows = H.sales(r, 14);
        const S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'Distributor', k: 'distributor' }, { h: 'SKU', k: 'sku' }, { h: 'Cartons', k: 'cartons', num: true }], rows);
        const pl = r.shuffle(H.PRODUCTS);
        const PM = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Product', k: 'name' }, { h: 'Carton price (Rp)', k: 'carton_price', num: true }], pl);
        const lo = r.pick([[0, 20, 40, 60], [0, 15, 30, 50], [0, 25, 45, 70]]), dd = r.pick([[0, 2, 3.5, 5], [0, 1.5, 3, 4.5], [0, 2.5, 4, 6]]);
        const DV = H.table([{ h: 'From cartons', k: 'f', num: true }, { h: 'Discount (%)', k: 'd', num: true }], lo.map((f, k) => ({ f, d: dd[k] })));
        /* The hidden test: row 5 sits exactly on another tier's lower limit, the top tier gets deeper, and one price rises. */
        const t0 = tierAt(lo, rows[0].cartons), tt = r.pick(nums(4).filter((m) => m > 0 && m !== t0));
        const alt = H.edit(S.cells, { [S.cell('Cartons', 0)]: lo[tt] });
        const dvAlt = H.edit(DV.cells, { [DV.cell('Discount (%)', 3)]: (x) => x + 0.5 });
        const pmAlt = H.edit(PM.cells, { [PM.cell('Carton price (Rp)', pl.findIndex((x) => x.sku === rows[5].sku))]: (x) => x + 4000 });
        const pm = X('master_produk.xlsx', 'Produk', PM), dv = X('diskon_volume.xlsx', 'Diskon', DV);
        const PS = pm.A('SKU'), PP = pm.A('Carton price (Rp)'), DF = dv.A('From cartons'), DD = dv.A('Discount (%)'), DT = dv.T('From cartons', 'Discount (%)');
        const price = `XLOOKUP(C5,${PS},${PP})`;
        const wrap = (p, d) => `=ROUND(D5*${p}*(1-${d}/100),0)`;
        return H.problem({
          title: 'Net order value after a volume discount tier from another file', file: 'pesanan_q4.xlsx', sheetName: 'Pesanan', sheet: S, base: S.cells, alt, fill: true, fillCol: 'E', ansLabel: 'Net (Rp)',
          books: [wb('master_produk.xlsx', 'Produk', PM, pmAlt), wb('diskon_volume.xlsx', 'Diskon', DV, dvAlt)],
          statement: `<p>${opens([['the order lines', 'pesanan_q4.xlsx'], ['the product master', 'master_produk.xlsx'], ['the volume discount table', 'diskon_volume.xlsx']])} The discount table lists the lowest cartons of each tier, smallest first; a line gets the <b>Discount (%)</b> of the highest <b>From cartons</b> it reaches.</p><p>In <b>E5</b>, work out the net value of each line: its <b>Cartons</b> times the SKU's <b>Carton price (Rp)</b> from the master, times (1 minus its discount divided by 100), rounded to whole rupiah with ROUND. Fill it down to <b>E${S.last}</b>.</p>`,
          solution: wrap(price, `VLOOKUP(D5,${DT},2,TRUE)`),
          alts: [wrap(price, `XLOOKUP(D5,${DF},${DD},,-1)`), wrap(`VLOOKUP(C5,${pm.T('SKU', 'Carton price (Rp)')},3,FALSE)`, `VLOOKUP(D5,${DT},2,TRUE)`),
            wrap(`INDEX(${PP},MATCH(C5,${PS},0))`, `INDEX(${DD},MATCH(D5,${DF},1))`), wrap(price, `LOOKUP(D5,${DF},${DD})`)],
          hints: ['Two lookups: an exact one for the price (by SKU, in master_produk.xlsx) and an approximate one for the discount (by cartons, in diskon_volume.xlsx).',
            `Most carton counts are not in the tier table. VLOOKUP with TRUE, or XLOOKUP with match_mode -1, returns the tier of the last lower limit reached: VLOOKUP(D5, ${DT}, 2, TRUE).`,
            `Shape: =ROUND(D5 * XLOOKUP(C5, ${PS}, ${PP}) * (1 - VLOOKUP(D5, ${DT}, 2, TRUE)/100), 0)`],
          explain: 'XLOOKUP brings the carton price from the product master. VLOOKUP with TRUE walks down the sorted lower limits in the discount file and stops at the largest one not above the cartons, so a line exactly on a limit gets that tier. Cartons times price, less the discount, is the net value.',
          chips: ['D5*', 'C5', 'D5', PS, PP, DT, DF, DD, ',2,TRUE)', ',,-1)'], fnKey: 'VLOOKUP('
        });
      }
      if (v === 1) {
        const rows = H.loans(r, 14);
        const S = H.table([{ h: 'Facility', k: 'facility' }, { h: 'Distributor', k: 'distributor' }, { h: 'Outstanding (Rp)', k: 'outstanding', num: true }, { h: 'DPD', k: 'dpd', num: true }], rows);
        const lo = [0, 1, 91, 121, 181];
        const KT = H.table([{ h: 'From DPD', k: 'f', num: true }, { h: 'Kol', k: 'k', num: true }], lo.map((f, k) => ({ f, k: k + 1 })));
        const CD = H.table([{ h: 'Kol', k: 'k', num: true }, { h: 'Provision (%)', k: 'p', num: true }], [1, 5, 15, 50, 100].map((p, k) => ({ k: k + 1, p })));
        /* The hidden test: row 5's DPD moves to exactly another grade's lower limit, and the provision for Kol 2 doubles. */
        const t0 = tierAt(lo, rows[0].dpd), tt = r.pick(nums(5).filter((m) => m > 0 && m !== t0));
        const alt = H.edit(S.cells, { [S.cell('DPD', 0)]: lo[tt] });
        const cdAlt = H.edit(CD.cells, { [CD.cell('Provision (%)', 1)]: 10 });
        const kt = X('kebijakan_kol.xlsx', 'Kol', KT), cd = X('cadangan.xlsx', 'Cadangan', CD);
        const KF = kt.A('From DPD'), KK = kt.A('Kol'), KTT = kt.T('From DPD', 'Kol'), CK = cd.A('Kol'), CP = cd.A('Provision (%)');
        const kol = `VLOOKUP(D5,${KTT},2,TRUE)`;
        const wrap = (x) => `=ROUND(C5*${x}/100,0)`;
        return H.problem({
          title: 'Loan loss provision from Kol tiers and a provision table', file: 'buku_pinjaman.xlsx', sheetName: 'Pinjaman', sheet: S, base: S.cells, alt, fill: true, fillCol: 'E', ansLabel: 'Provision (Rp)',
          books: [wb('kebijakan_kol.xlsx', 'Kol', KT), wb('cadangan.xlsx', 'Cadangan', CD, cdAlt)],
          statement: `<p>${opens([['the loan book', 'buku_pinjaman.xlsx'], ['the Kol policy', 'kebijakan_kol.xlsx'], ['the provision table', 'cadangan.xlsx']])} The Kol policy lists the lowest DPD (days past due) of each collectibility grade, smallest first; a facility gets the grade of the highest <b>From DPD</b> it reaches. The provision table gives the <b>Provision (%)</b> for each grade.</p><p>In <b>E5</b>, work out each facility's provision: its <b>Outstanding (Rp)</b> times the provision % of its grade, divided by 100 and rounded to whole rupiah with ROUND. Fill it down to <b>E${S.last}</b>.</p>`,
          solution: wrap(`XLOOKUP(${kol},${CK},${CP})`),
          alts: [wrap(`INDEX(${CP},${kol})`), wrap(`VLOOKUP(XLOOKUP(D5,${KF},${KK},,-1),${cd.T('Kol', 'Provision (%)')},2,FALSE)`), wrap(`XLOOKUP(LOOKUP(D5,${KF},${KK}),${CK},${CP})`)],
          hints: ['Work from the inside out: the DPD in D5 gives a Kol grade (approximate, from the policy file), and the grade gives a provision % (exact, from the provision table).',
            `${kol.replace(/,/g, ', ')} is the grade. Use it as the value to find in ${CK}.`,
            `Shape: =ROUND(C5 * XLOOKUP(VLOOKUP(D5, ${KTT}, 2, TRUE), ${CK}, ${CP}) / 100, 0)`],
          explain: 'VLOOKUP with TRUE turns the DPD into a Kol grade: the last lower limit it reaches in kebijakan_kol.xlsx. XLOOKUP then finds that grade in cadangan.xlsx and returns its provision %. Outstanding times that %, divided by 100, is the provision the bank sets aside.',
          chips: ['C5*', 'D5', KTT, KF, KK, CK, CP, ',2,TRUE)'], fnKey: 'XLOOKUP('
        });
      }
      if (v === 2) {
        const rows = H.reps(r, 10);
        const S = H.table([{ h: 'Rep', k: 'name' }, { h: 'Area', k: 'area' }, { h: 'Actual (Rp)', k: 'actual', num: true }], rows);
        const tl = r.shuffle(rows).map((x) => ({ n: x.name, t: x.target }));
        const TG = H.table([{ h: 'Rep', k: 'n' }, { h: 'Target (Rp)', k: 't', num: true }], tl);
        const c = r.pick([[0, 0.5, 1, 1.5], [0, 1, 1.5, 2.5], [0, 0.75, 1.25, 2]]), lo = [0, 0.8, 1, 1.1];
        const SK = H.table([{ h: 'From achievement', k: 'f', num: true }, { h: 'Commission (%)', k: 'c', num: true }], lo.map((f, k) => ({ f, c: c[k] })));
        /* The hidden test: row 5's actual lands exactly on another tier's lower limit, the top rate rises, and one target
           is revised. */
        const t0 = tierAt(lo, rows[0].actual / rows[0].target), tt = r.pick(nums(4).filter((m) => m > 0 && m !== t0));
        const alt = H.edit(S.cells, { [S.cell('Actual (Rp)', 0)]: Math.round(rows[0].target * lo[tt] * 100) / 100 });
        const skAlt = H.edit(SK.cells, { [SK.cell('Commission (%)', 3)]: (x) => x + 0.25 });
        const tgAlt = H.edit(TG.cells, { [TG.cell('Target (Rp)', tl.findIndex((x) => x.n === rows[4].name))]: (x) => x - 20000000 });
        const tg = X('target_sales.xlsx', 'Target', TG), sk = X('skema_komisi.xlsx', 'Komisi', SK);
        const TR = tg.A('Rep'), TT = tg.A('Target (Rp)'), SF = sk.A('From achievement'), SC = sk.A('Commission (%)'), ST = sk.T('From achievement', 'Commission (%)');
        const ach = `C5/XLOOKUP(A5,${TR},${TT})`;
        const wrap = (x) => `=ROUND(C5*${x}/100,0)`;
        return H.problem({
          title: 'Rep commission from achievement tiers, targets in another file', file: 'komisi_q4.xlsx', sheetName: 'Realisasi', sheet: S, base: S.cells, alt, fill: true, fillCol: 'D', ansLabel: 'Commission (Rp)',
          books: [wb('target_sales.xlsx', 'Target', TG, tgAlt), wb('skema_komisi.xlsx', 'Komisi', SK, skAlt)],
          statement: `<p>${opens([['the actual sales', 'komisi_q4.xlsx'], ['the targets', 'target_sales.xlsx'], ['the commission scheme', 'skema_komisi.xlsx']])} The scheme lists the lowest achievement (Actual divided by Target) of each tier, smallest first; a rep gets the <b>Commission (%)</b> of the highest <b>From achievement</b> reached.</p><p>In <b>D5</b>, work out each rep's commission: <b>Actual (Rp)</b> times the commission % for their achievement, divided by 100 and rounded to whole rupiah with ROUND. The target is in the targets file. Fill it down to <b>D${S.last}</b>.</p>`,
          solution: wrap(`VLOOKUP(${ach},${ST},2,TRUE)`),
          alts: [wrap(`XLOOKUP(${ach},${SF},${SC},,-1)`), wrap(`INDEX(${SC},MATCH(C5/VLOOKUP(A5,${tg.T('Rep', 'Target (Rp)')},2,FALSE),${SF},1))`), wrap(`LOOKUP(${ach},${SF},${SC})`)],
          hints: ['The achievement is not on the sheet: Actual in C5 divided by the target that XLOOKUP brings from target_sales.xlsx.',
            `That ratio is the value to find in the scheme, with an approximate match: VLOOKUP(ratio, ${ST}, 2, TRUE).`,
            `Shape: =ROUND(C5 * VLOOKUP(C5/XLOOKUP(A5, ${TR}, ${TT}), ${ST}, 2, TRUE) / 100, 0)`],
          explain: 'XLOOKUP brings the rep\'s target, and C5 divided by it is the achievement. VLOOKUP with TRUE finds the last lower limit that achievement reaches in skema_komisi.xlsx, so exactly 100% earns the 100% tier. Actual times that rate, divided by 100, is the commission.',
          chips: ['C5*', 'C5/', 'A5', TR, TT, ST, SF, SC, ',2,TRUE)'], fnKey: 'VLOOKUP('
        });
      }
      const asOf = H.toSerial('2026-12-31');
      const over = r.shuffle([-20, -5, 0, 6, 18, 27, 34, 47, 58, 66, 79, 88, 102, 135]);
      const inv = H.invoices(r, 14).map((x, k) => ({ ref: x.ref, d: x.distributor, date: asOf - DIST().find((y) => y.name === x.distributor).terms_days - over[k], amount: x.amount }));
      const S = H.table([{ h: 'Invoice', k: 'ref' }, { h: 'Distributor', k: 'd' }, { h: 'Invoice date', k: 'date', date: true }, { h: 'Amount (Rp)', k: 'amount', num: true }], inv);
      const dl = r.shuffle(DIST());
      const MD = H.table([DC.id, DC.name, DC.city, DC.terms], dl);
      const lo = [0, 1, 31, 61, 91];
      const AG = H.table([{ h: 'From days', k: 'f', num: true }, { h: 'Bucket', k: 'b' }, { h: 'Allowance (%)', k: 'a', num: true }],
        lo.map((f, k) => ({ f, b: ['Current', '1 to 30', '31 to 60', '61 to 90', 'Over 90'][k], a: [0, 1, 5, 20, 50][k] })));
      const base = Object.assign({}, S.cells, { A1: 'As of', B1: asOf });
      /* The hidden test: the policy raises two allowances, and the distributor of an invoice 1 to 90 days overdue gets
         30 more days of terms in the master, so that invoice drops to an earlier bucket. */
      const agAlt = H.edit(AG.cells, { [AG.cell('Allowance (%)', 1)]: 2, [AG.cell('Allowance (%)', 3)]: 25 });
      const late = Math.max(0, over.findIndex((o) => o >= 1 && o <= 90));
      const mdAlt = H.edit(MD.cells, { [MD.cell('Terms (days)', dl.findIndex((d) => d.name === inv[late].d))]: (x) => x + 30 });
      const md = X('master_distributor.xlsx', 'Distributor', MD), ag = X('kebijakan_piutang.xlsx', 'Aging', AG);
      const MN = md.A('Distributor'), MT = md.A('Terms (days)'), AF = ag.A('From days'), AA = ag.A('Allowance (%)'), AT = ag.T('From days', 'Allowance (%)');
      const days = `MAX(0,$B$1-C5-XLOOKUP(B5,${MN},${MT}))`;
      const wrap = (x) => `=ROUND(D5*${x}/100,0)`;
      return H.problem({
        title: 'Bad-debt allowance from aging tiers and distributor terms', file: 'piutang_q4.xlsx', sheetName: 'Piutang', sheet: S, base, alt: base, fill: true, fillCol: 'E', ansLabel: 'Allowance (Rp)', dateCells: ['B1'],
        books: [wb('master_distributor.xlsx', 'Distributor', MD, mdAlt), wb('kebijakan_piutang.xlsx', 'Aging', AG, agAlt)],
        statement: `<p>${opens([['the open invoices', 'piutang_q4.xlsx'], ['the distributor master', 'master_distributor.xlsx'], ['the credit policy', 'kebijakan_piutang.xlsx']])} <b>B1</b> holds the as-of date. The policy's aging table lists the lowest days overdue of each bucket, smallest first.</p><p>In <b>E5</b>, work out the allowance for each invoice. Its days overdue are B1 minus the <b>Invoice date</b> minus the distributor's <b>Terms (days)</b> from the master; below 0 the invoice is not due yet, so count it as 0 with MAX. The invoice gets the <b>Allowance (%)</b> of the highest <b>From days</b> it reaches. Allowance = <b>Amount (Rp)</b> times that % divided by 100, rounded to whole rupiah with ROUND. Fill it down to <b>E${S.last}</b>.</p>`,
        solution: wrap(`VLOOKUP(${days},${AT},3,TRUE)`),
        alts: [wrap(`XLOOKUP(${days},${AF},${AA},,-1)`), wrap(`INDEX(${AA},MATCH(${days},${AF},1))`), wrap(`LOOKUP(${days},${AF},${AA})`),
          wrap(`VLOOKUP(MAX(0,$B$1-(C5+VLOOKUP(B5,${md.T('Distributor', 'Terms (days)')},3,FALSE))),${AT},3,TRUE)`)],
        hints: ['First the days overdue: $B$1 - C5 - the terms that XLOOKUP brings from master_distributor.xlsx. MAX(0, ...) turns a not-yet-due invoice into 0.',
          `Then an approximate match on those days in the aging table: VLOOKUP(days, ${AT}, 3, TRUE) returns the Allowance (%), the third column.`,
          `Shape: =ROUND(D5 * VLOOKUP(MAX(0, $B$1 - C5 - XLOOKUP(B5, ${MN}, ${MT})), ${AT}, 3, TRUE) / 100, 0)`],
        explain: 'XLOOKUP brings each distributor\'s payment terms from the master, so the invoice date plus the terms is its due date and B1 minus that is the days overdue, never below 0. VLOOKUP with TRUE puts those days in an aging bucket from the policy file and returns its allowance %. Amount times that %, divided by 100, is the allowance.',
        chips: ['D5*', '$B$1', 'C5', 'B5', MN, MT, AT, AF, AA, ',3,TRUE)', 'MAX(0,'], fnKey: 'VLOOKUP('
      });
    }
  });

  P.add({
    key: 'lookup-on-two-conditions', difficulty: 'Hard', level: 10, topics: ['XLOOKUP', 'INDEX', 'MATCH', 'Several workbooks'], count: 3, multi: true,
    make(r, v) {
      /* In the other workbook no single column is unique: the row is found by two keys at once, multiplied as 1s and 0s.
         A second lookup in a third workbook brings the figure the result is multiplied by, or one of the keys. */
      if (v === 0) {
        const skus = r.sample(H.PRODUCTS, 5), whs = r.sample(H.WAREHOUSES, 3);
        let st = [];
        skus.forEach((x) => whs.forEach((w) => st.push({ sku: x.sku, wh: w, oh: r.int(1, 30) * 10, rp: r.pick([80, 100, 120, 150]) })));
        st = r.shuffle(st);
        const ST = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Warehouse', k: 'wh' }, { h: 'On hand', k: 'oh', num: true }, { h: 'Reorder point', k: 'rp', num: true }], st);
        const pl = r.shuffle(H.PRODUCTS);
        const PM = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Product', k: 'name' }, { h: 'Carton price (Rp)', k: 'carton_price', num: true }], pl);
        const rows = r.sample(st, 14).map((x, k) => ({ req: 'REQ-' + (501 + k * 3 + r.int(0, 2)), wh: x.wh, sku: x.sku, cartons: Math.max(10, x.oh + r.pick([-60, -30, -10, 0, 20, 40, 70])) }));
        const S = H.table([{ h: 'Request', k: 'req' }, { h: 'Warehouse', k: 'wh' }, { h: 'SKU', k: 'sku' }, { h: 'Cartons', k: 'cartons', num: true }], rows);
        /* The hidden test: the stock count for row 5's SKU and warehouse is recounted, and one price rises. */
        const s0 = st.findIndex((x) => x.sku === rows[0].sku && x.wh === rows[0].wh);
        const stAlt = H.edit(ST.cells, { [ST.cell('On hand', s0)]: rows[0].cartons > st[s0].oh ? rows[0].cartons + 20 : Math.max(0, rows[0].cartons - 30) });
        const pmAlt = H.edit(PM.cells, { [PM.cell('Carton price (Rp)', pl.findIndex((x) => x.sku === rows[1].sku))]: (x) => x + 6000 });
        const sg = X('stok_gudang.xlsx', 'Stok', ST), pm = X('master_produk.xlsx', 'Produk', PM);
        const SS = sg.A('SKU'), SW = sg.A('Warehouse'), SO = sg.A('On hand'), PS = pm.A('SKU'), PP = pm.A('Carton price (Rp)');
        const cond = `(${SS}=C5)*(${SW}=B5)`, price = `XLOOKUP(C5,${PS},${PP})`;
        return H.problem({
          title: 'Value of each stock shortfall, by SKU and warehouse', file: 'permintaan_stok.xlsx', sheetName: 'Permintaan', sheet: S, base: S.cells, alt: S.cells, fill: true, fillCol: 'E', ansLabel: 'Shortfall (Rp)',
          books: [wb('stok_gudang.xlsx', 'Stok', ST, stAlt), wb('master_produk.xlsx', 'Produk', PM, pmAlt)],
          statement: `<p>${opens([['the stock requests', 'permintaan_stok.xlsx'], ['the stock count', 'stok_gudang.xlsx'], ['the product master', 'master_produk.xlsx']])} The stock count has one row for each SKU in each warehouse, so neither column alone finds the right row.</p><p>In <b>E5</b>, work out the value of the shortfall on each request: the <b>Cartons</b> asked for minus the <b>On hand</b> for that SKU <b>and</b> warehouse, but never below 0 (use MAX), times the SKU's <b>Carton price (Rp)</b> from the product master. Fill it down to <b>E${S.last}</b>.</p>`,
          solution: `=MAX(0,D5-XLOOKUP(1,${cond},${SO}))*${price}`,
          alts: [`=MAX(0,D5-SUMIFS(${SO},${SS},C5,${SW},B5))*${price}`, `=MAX(0,D5-INDEX(${SO},MATCH(1,${cond},0)))*VLOOKUP(C5,${pm.T('SKU', 'Carton price (Rp)')},3,FALSE)`,
            `=MAX(D5-XLOOKUP(1,(${SW}=B5)*(${SS}=C5),${SO}),0)*${price}`],
          hints: [`(${SS}=C5) is a column of TRUE and FALSE. Multiplied by (${SW}=B5), it is 1 only on the row where both match; XLOOKUP(1, ...) finds it.`,
            'D5 minus that stock is the shortfall; MAX(0, ...) turns a covered request into 0. The price is a plain lookup in master_produk.xlsx.',
            `Shape: =MAX(0, D5 - XLOOKUP(1, (${SS}=C5)*(${SW}=B5), ${SO})) * XLOOKUP(C5, ${PS}, ${PP})`],
          explain: 'The two comparisons mark the rows of the stock count that match the SKU and the warehouse; their product is 1 only where both match, and XLOOKUP returns that row\'s On hand. MAX keeps the shortfall at 0 or more, and the carton price from the product master turns it into rupiah. SUMIFS on both conditions gives the same stock, since each pair appears once.',
          chips: ['D5-', 'B5', 'C5', '(' + SS + '=C5)', '(' + SW + '=B5)', SO, PS, PP, 'MAX(0,'], fnKey: 'XLOOKUP('
        });
      }
      if (v === 1) {
        const facs = H.db().facilities;
        const two = [...new Set(facs.filter((f) => f.type === 'Term Loan').map((f) => f.distributor))];
        const both = facs.filter((f) => two.includes(f.distributor));
        const pick = r.shuffle(both.concat(r.shuffle(facs.filter((f) => !two.includes(f.distributor))).slice(0, Math.max(0, 14 - both.length)))).slice(0, 14);
        const rows = pick.map((f) => ({ d: f.distributor, type: f.type, days: r.pick([31, 31, 31, 24, 17, 10]) }));
        const S = H.table([{ h: 'Distributor', k: 'd' }, { h: 'Loan type', k: 'type' }, { h: 'Days', k: 'days', num: true }], rows);
        const bl = r.shuffle(facs);
        const BP = H.table([{ h: 'Facility', k: 'facility_id' }, { h: 'Distributor', k: 'distributor' }, { h: 'Type', k: 'type' }, { h: 'Outstanding (Rp)', k: 'outstanding', num: true }], bl);
        const RC = H.table([{ h: 'Loan type', k: 't' }, { h: 'Rate (% a year)', k: 'x', num: true }],
          [{ t: 'Working Capital', x: r.pick([10, 10.5, 11]) }, { t: 'Term Loan', x: r.pick([11.5, 12, 12.5]) }, { t: 'Trade Finance', x: 9.5 }, { t: 'Investment', x: 11.25 }]);
        /* The hidden test: a term loan's balance falls after a repayment, the Working Capital rate rises, and row 7 was
           drawn for a different number of days. */
        const tl = bl.findIndex((f) => f.type === 'Term Loan' && rows.some((x) => x.d === f.distributor && x.type === 'Term Loan'));
        const bpAlt = H.edit(BP.cells, tl >= 0 ? { [BP.cell('Outstanding (Rp)', tl)]: (x) => x - 3000000 } : {});
        const rcAlt = H.edit(RC.cells, { [RC.cell('Rate (% a year)', 0)]: (x) => x + 0.5 });
        const alt = H.edit(S.cells, { [S.cell('Days', 2)]: (x) => (x === 31 ? 30 : 31) });
        const bp = X('buku_pinjaman.xlsx', 'Fasilitas', BP), rc = X('suku_bunga.xlsx', 'Bunga', RC);
        const BD = bp.A('Distributor'), BT = bp.A('Type'), BO = bp.A('Outstanding (Rp)'), RT = rc.A('Loan type'), RR = rc.A('Rate (% a year)');
        const cond = `(${BD}=A5)*(${BT}=B5)`, rate = `XLOOKUP(B5,${RT},${RR})`;
        const wrap = (o, x) => `=ROUND(${o}*${x}/100*C5/365,0)`;
        return H.problem({
          title: 'Monthly interest bill by borrower and loan type', file: 'tagihan_bunga.xlsx', sheetName: 'Tagihan', sheet: S, base: S.cells, alt, fill: true, fillCol: 'D', ansLabel: 'Interest (Rp)',
          books: [wb('buku_pinjaman.xlsx', 'Fasilitas', BP, bpAlt), wb('suku_bunga.xlsx', 'Bunga', RC, rcAlt)],
          statement: `<p>${opens([['the interest bills', 'tagihan_bunga.xlsx'], ['the loan book', 'buku_pinjaman.xlsx'], ['the rate card', 'suku_bunga.xlsx']])} Some distributors have two facilities, a Working Capital line and a Term Loan, so the loan book must be searched on the distributor <b>and</b> the type.</p><p>In <b>D5</b>, work out each bill: the facility's <b>Outstanding (Rp)</b> from the loan book, times the <b>Rate (% a year)</b> of its loan type from the rate card, divided by 100, times <b>Days</b> divided by 365, rounded to whole rupiah with ROUND. Fill it down to <b>D${S.last}</b>.</p>`,
          solution: wrap(`XLOOKUP(1,${cond},${BO})`, rate),
          alts: [wrap(`SUMIFS(${BO},${BD},A5,${BT},B5)`, rate), wrap(`INDEX(${BO},MATCH(1,${cond},0))`, `VLOOKUP(B5,${rc.T('Loan type', 'Rate (% a year)')},2,FALSE)`),
            wrap(`XLOOKUP(1,(${BT}=B5)*(${BD}=A5),${BO})`, rate)],
          hints: [`The balance needs two conditions in buku_pinjaman.xlsx: (${BD}=A5)*(${BT}=B5) is 1 only on the facility that matches both. XLOOKUP(1, that, ${BO}) returns its balance.`,
            'The rate is a plain lookup of the loan type in suku_bunga.xlsx. Then divide by 100, multiply by C5 and divide by 365.',
            `Shape: =ROUND(XLOOKUP(1, (${BD}=A5)*(${BT}=B5), ${BO}) * XLOOKUP(B5, ${RT}, ${RR}) / 100 * C5 / 365, 0)`],
          explain: 'Multiplying the two comparisons leaves 1 only on the loan book row of that distributor and that type, so XLOOKUP(1, ...) returns the right balance even when the distributor has two facilities. The second XLOOKUP brings the yearly rate for the type. Balance times rate, for the days drawn out of 365, is the bill.',
          chips: ['A5', 'B5', 'C5', '(' + BD + '=A5)', '(' + BT + '=B5)', BO, RT, RR, '/100*C5/365'], fnKey: 'XLOOKUP('
        });
      }
      const skus = r.sample(H.PRODUCTS, 4);
      const fct = {};
      H.REGIONS.forEach((g) => { fct[g] = r.pick([1, 1.02, 1.04, 1.06, 1.08]); });
      const hl = [];
      H.REGIONS.forEach((g) => skus.forEach((x) => hl.push({ g, sku: x.sku, price: Math.round(x.carton_price * fct[g] / 1000) * 1000 })));
      const HR = H.table([{ h: 'Region', k: 'g' }, { h: 'SKU', k: 'sku' }, { h: 'Carton price (Rp)', k: 'price', num: true }], hl);
      const rows = H.sales(r, 14).map((s) => ({ date: s.date, d: s.distributor, sku: r.pick(skus).sku, cartons: s.cartons }));
      const S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'Distributor', k: 'd' }, { h: 'SKU', k: 'sku' }, { h: 'Cartons', k: 'cartons', num: true }], rows);
      const MD = H.table([DC.id, DC.name, DC.city, DC.region], r.shuffle(DIST()));
      /* The hidden test: the regional price for row 5 rises, and row 8 turns out to be another SKU. */
      const h0 = hl.findIndex((x) => x.g === regionOf(rows[0].d) && x.sku === rows[0].sku);
      const hrAlt = H.edit(HR.cells, { [HR.cell('Carton price (Rp)', h0)]: (x) => x + 4000 });
      const alt = H.edit(S.cells, { [S.cell('SKU', 3)]: skus.find((x) => x.sku !== rows[3].sku).sku });
      const md = X('master_distributor.xlsx', 'Distributor', MD), hr = X('harga_regional.xlsx', 'Harga', HR);
      const MN = md.A('Distributor'), MR = md.A('Region'), HG = hr.A('Region'), HS = hr.A('SKU'), HP = hr.A('Carton price (Rp)');
      const reg = `XLOOKUP(B5,${MN},${MR})`;
      return H.problem({
        title: 'Order value from a price list by region and SKU', file: 'pesanan_q4.xlsx', sheetName: 'Pesanan', sheet: S, base: S.cells, alt, fill: true, fillCol: 'E', ansLabel: 'Value (Rp)',
        books: [wb('master_distributor.xlsx', 'Distributor', MD), wb('harga_regional.xlsx', 'Harga', HR, hrAlt)],
        statement: `<p>${opens([['the order lines', 'pesanan_q4.xlsx'], ['the distributor master', 'master_distributor.xlsx'], ['the regional price list', 'harga_regional.xlsx']])} Prices differ by region: the price list has one row for each region and SKU. The orders do not show the region.</p><p>In <b>E5</b>, work out each line's value: its <b>Cartons</b> times the <b>Carton price (Rp)</b> for its SKU in its distributor's region, which you look up in the master first. Fill it down to <b>E${S.last}</b>.</p>`,
        solution: `=D5*XLOOKUP(1,(${HG}=${reg})*(${HS}=C5),${HP})`,
        alts: [`=D5*SUMIFS(${HP},${HG},${reg},${HS},C5)`, `=D5*INDEX(${HP},MATCH(1,(${HG}=VLOOKUP(B5,${md.T('Distributor', 'Region')},3,FALSE))*(${HS}=C5),0))`,
          `=D5*XLOOKUP(1,(${HS}=C5)*(${HG}=${reg}),${HP})`],
        hints: [`The region comes first: XLOOKUP(B5, ${MN}, ${MR}) from master_distributor.xlsx.`,
          `Then find the price row that matches both that region and the SKU in C5: (${HG}=region)*(${HS}=C5) is 1 only there.`,
          `Shape: =D5 * XLOOKUP(1, (${HG}=XLOOKUP(B5, ${MN}, ${MR}))*(${HS}=C5), ${HP})`],
        explain: 'The inner XLOOKUP turns the distributor into its region. Compared with the Region column of the price list and multiplied by the SKU comparison, it leaves 1 on the one row for that region and SKU, and the outer XLOOKUP returns its price. Times the cartons, that is the line\'s value. SUMIFS on the two conditions gives the same price.',
        chips: ['D5*', 'B5', 'C5', MN, MR, HG, HS, HP, 'XLOOKUP(1,'], fnKey: 'XLOOKUP('
      });
    }
  });

  P.add({
    key: 'last-match-in-a-log', difficulty: 'Hard', level: 9, topics: ['XLOOKUP', 'XMATCH', 'search mode', 'Several workbooks'], count: 3, multi: true,
    make(r, v) {
      /* A log in date order, kept in another workbook: the latest entry for a key is its last match, found by searching
         from the bottom. A third workbook holds what it is compared with or combined with. */
      if (v === 0) {
        const ds = r.sample(DIST(), 8);
        const pool = H.db().sales.filter((s) => ds.some((d) => d.id === s.distributor_id));
        const one = ds.map((d) => r.pick(pool.filter((s) => s.distributor_id === d.id)));
        const log = one.concat(r.sample(pool.filter((s) => !one.includes(s)), 16)).map((s) => ({ date: H.Q4(s.date), inv: s.invoice, d: s.distributor, net: s.net }))
          .sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : a.inv < b.inv ? -1 : 1));
        const LG = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'Invoice', k: 'inv' }, { h: 'Distributor', k: 'd' }, { h: 'Net (Rp)', k: 'net', num: true }], log);
        const NICE = [7, 14, 21, 30, 45, 60, 90, 120];
        const cyc = r.shuffle(DIST()).map((d) => ({ d: d.name, region: d.region, c: r.pick([14, 21, 30, 45]) }));
        const MD = H.table([{ h: 'Distributor', k: 'd' }, { h: 'Region', k: 'region' }, { h: 'Order cycle (days)', k: 'c', num: true }], cyc);
        const S = H.table([{ h: 'Distributor', k: 'd' }, { h: 'City', k: 'city' }, { h: 'Rep', k: 'rep' }], ds.map((d) => ({ d: d.name, city: d.city, rep: r.pick(H.REPS) })));
        const asOf = H.toSerial('2026-12-31');
        const since = (name) => asOf - Math.max(...log.filter((x) => x.d === name).map((x) => H.toSerial(x.date)));
        /* The hidden test: the master revises one distributor's order cycle so its status flips. */
        const k = Math.max(0, ds.findIndex((d) => since(d.name) >= 10));
        const w = cyc.findIndex((x) => x.d === ds[k].name), s = since(ds[k].name);
        const nc = cyc[w].c >= s ? Math.max(0, ...NICE.filter((x) => x < s)) : Math.min(...NICE.filter((x) => x >= s));
        const mdAlt = H.edit(MD.cells, { [MD.cell('Order cycle (days)', w)]: nc });
        const base = Object.assign({}, S.cells, { A1: 'As of', B1: asOf });
        const lg = X('penjualan_q4.xlsx', 'Penjualan', LG), md = X('master_distributor.xlsx', 'Distributor', MD);
        const LD = lg.A('Distributor'), LT = lg.A('Date'), MN = md.A('Distributor'), MC = md.A('Order cycle (days)');
        const last = `XLOOKUP(A5,${LD},${LT},,0,-1)`, cycle = `XLOOKUP(A5,${MN},${MC})`;
        return H.problem({
          title: 'Distributors due a follow-up call, from the last order in the log', file: 'tindak_lanjut.xlsx', sheetName: 'Review', sheet: S, base, alt: base, fill: true, fillCol: 'D', ansLabel: 'Status', dateCells: ['B1'],
          books: [wb('penjualan_q4.xlsx', 'Penjualan', LG), wb('master_distributor.xlsx', 'Distributor', MD, mdAlt)],
          statement: `<p>${opens([['the call list', 'tindak_lanjut.xlsx'], ['the sales log', 'penjualan_q4.xlsx'], ['the distributor master', 'master_distributor.xlsx']])} <b>B1</b> holds the review date. The sales log is in date order, oldest first, so a distributor's last row is its latest order.</p><p>In <b>D5</b>, show <b>Follow up</b> when the days since the distributor's latest order (B1 minus that order's <b>Date</b>) are more than its <b>Order cycle (days)</b> in the master, and <b>OK</b> otherwise. Fill it down to <b>D${S.last}</b>.</p>`,
          solution: `=IF($B$1-${last}>${cycle},"Follow up","OK")`,
          alts: [`=IF($B$1-MAXIFS(${LT},${LD},A5)>${cycle},"Follow up","OK")`, `=IF($B$1-INDEX(${LT},XMATCH(A5,${LD},0,-1))>VLOOKUP(A5,${md.T('Distributor', 'Order cycle (days)')},3,FALSE),"Follow up","OK")`,
            `=IF(${last}+${cycle}<$B$1,"Follow up","OK")`],
          hints: ['A normal lookup returns the first match from the top, which in this log is the oldest order. XLOOKUP\'s sixth part, search_mode, set to -1 searches from the bottom, so it finds the latest.',
            `$B$1 minus that date is the days since the last order. Compare it with the cycle that XLOOKUP brings from master_distributor.xlsx.`,
            `Shape: =IF($B$1 - XLOOKUP(A5, ${LD}, ${LT}, , 0, -1) > XLOOKUP(A5, ${MN}, ${MC}), "Follow up", "OK")`],
          explain: 'Searching from the bottom of the sales log, XLOOKUP meets the distributor\'s latest order first and returns its date. B1 minus that date is how long it has been quiet, and the second XLOOKUP brings its usual order cycle from the master. More days than the cycle means a call is due. MAXIFS on the dates gives the same latest date here.',
          chips: ['$B$1-', 'A5', LD, LT, MN, MC, ',,0,-1)', '"Follow up"', '"OK"'], fnKey: 'XLOOKUP('
        });
      }
      if (v === 1) {
        const skus = r.sample(H.PRODUCTS, 5);
        const seq = r.shuffle([0, 0, 0, 1, 1, 1, 2, 2, 2, 3, 3, 3, 4, 4, 4, 0]);
        const price = skus.map((x) => x.carton_price);
        let day = H.toSerial('2026-01-05');
        const log = seq.map((k) => { day += r.int(6, 20); price[k] += r.pick([-6000, -4000, 3000, 5000, 7000, 9000]); return { date: day, sku: skus[k].sku, price: price[k] }; });
        const RH = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'SKU', k: 'sku' }, { h: 'New price (Rp)', k: 'price', num: true }], log);
        const rows = H.sales(r, 14).map((s) => ({ date: s.date, d: s.distributor, sku: r.pick(skus).sku, cartons: s.cartons }));
        const S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'Distributor', k: 'd' }, { h: 'SKU', k: 'sku' }, { h: 'Cartons', k: 'cartons', num: true }], rows);
        const dl = r.shuffle(DIST());
        const MD = H.table([DC.id, DC.name, DC.region, DC.disc], dl);
        /* The hidden test: the latest change for row 5's SKU is corrected in the log, and row 7's distributor gets a
           bigger discount. */
        const lastIdx = (sku) => log.map((x) => x.sku).lastIndexOf(sku);
        const rhAlt = H.edit(RH.cells, { [RH.cell('New price (Rp)', lastIdx(rows[0].sku))]: (x) => x + 3000 });
        const mdAlt = H.edit(MD.cells, { [MD.cell('Discount (%)', dl.findIndex((d) => d.name === rows[2].d))]: (x) => x + 1.5 });
        const rh = X('riwayat_harga.xlsx', 'Riwayat', RH), md = X('master_distributor.xlsx', 'Distributor', MD);
        const RS = rh.A('SKU'), RP = rh.A('New price (Rp)'), MN = md.A('Distributor'), MDc = md.A('Discount (%)');
        const disc = `XLOOKUP(B5,${MN},${MDc})`;
        const wrap = (p, d) => `=ROUND(D5*${p}*(1-${d}/100),0)`;
        return H.problem({
          title: 'Order value at the current price from the price log, after discount', file: 'pesanan_q4.xlsx', sheetName: 'Pesanan', sheet: S, base: S.cells, alt: S.cells, fill: true, fillCol: 'E', ansLabel: 'Value (Rp)',
          books: [wb('riwayat_harga.xlsx', 'Riwayat', RH, rhAlt), wb('master_distributor.xlsx', 'Distributor', MD, mdAlt)],
          statement: `<p>${opens([['the order lines', 'pesanan_q4.xlsx'], ['the price change log', 'riwayat_harga.xlsx'], ['the distributor master', 'master_distributor.xlsx']])} The log records each price change, oldest first; a SKU's last change is its current price, and these orders are billed at the current price.</p><p>In <b>E5</b>, work out each line's value: its <b>Cartons</b> times the SKU's current price, times (1 minus the distributor's <b>Discount (%)</b> from the master divided by 100), rounded to whole rupiah with ROUND. Fill it down to <b>E${S.last}</b>.</p>`,
          solution: wrap(`XLOOKUP(C5,${RS},${RP},,0,-1)`, disc),
          alts: [wrap(`INDEX(${RP},XMATCH(C5,${RS},0,-1))`, disc), wrap(`XLOOKUP(C5,${RS},${RP},,,-1)`, `VLOOKUP(B5,${md.T('Distributor', 'Discount (%)')},3,FALSE)`),
            `=ROUND(D5*XLOOKUP(C5,${RS},${RP},,0,-1)*(100-${disc})/100,0)`],
          hints: ['The SKU appears several times in the log; you want its last row. XLOOKUP searches from the bottom when its sixth part, search_mode, is -1.',
            'The discount is a plain lookup of the distributor in master_distributor.xlsx. (1 - discount/100) is the share the distributor pays.',
            `Shape: =ROUND(D5 * XLOOKUP(C5, ${RS}, ${RP}, , 0, -1) * (1 - XLOOKUP(B5, ${MN}, ${MDc})/100), 0)`],
          explain: 'Searching from the bottom, XLOOKUP meets the SKU\'s newest change first and returns that price; the highest price is not always the current one, since some changes were cuts. The second XLOOKUP brings the distributor\'s discount from the master. Cartons times price, less the discount, is the value.',
          chips: ['D5*', 'B5', 'C5', RS, RP, MN, MDc, ',,0,-1)', '(1-'], fnKey: 'XLOOKUP('
        });
      }
      const fac = r.sample(H.db().facilities, 6);
      const ends = ['2026-07-31', '2026-08-31', '2026-09-30', '2026-10-31'];
      const dpd = fac.map(() => r.pick([0, 0, 5, 30, 60, 95]));
      const log = [];
      ends.forEach((m, mi) => r.shuffle(nums(6)).forEach((k) => { if (mi) dpd[k] = Math.max(0, dpd[k] + r.pick([-30, 0, 0, 15, 30, 45])); log.push({ m, f: fac[k].facility_id, dpd: dpd[k] }); }));
      const DL = H.table([{ h: 'Month end', k: 'm', date: true }, { h: 'Facility', k: 'f' }, { h: 'DPD', k: 'dpd', num: true }], log);
      const lo = [0, 1, 91, 121, 181];
      const KT = H.table([{ h: 'From DPD', k: 'f', num: true }, { h: 'Kol', k: 'k', num: true }], lo.map((f, k) => ({ f, k: k + 1 })));
      const S = H.table([{ h: 'Facility', k: 'f' }, { h: 'Distributor', k: 'd' }, { h: 'Outstanding (Rp)', k: 'out', num: true }], fac.map((f) => ({ f: f.facility_id, d: f.distributor, out: f.outstanding })));
      /* The hidden test: the latest DPD of row 5's facility is revised into another grade. */
      const li = log.map((x) => x.f).lastIndexOf(fac[0].facility_id);
      const dlAlt = H.edit(DL.cells, { [DL.cell('DPD', li)]: tierAt(lo, log[li].dpd) >= 2 ? 30 : 95 });
      const dl = X('pantauan_dpd.xlsx', 'DPD', DL), kt = X('kebijakan_kol.xlsx', 'Kol', KT);
      const DF = dl.A('Facility'), DD = dl.A('DPD'), KF = kt.A('From DPD'), KK = kt.A('Kol'), KTT = kt.T('From DPD', 'Kol');
      const latest = `XLOOKUP(A5,${DF},${DD},,0,-1)`, first = `XLOOKUP(A5,${DF},${DD})`;
      return H.problem({
        title: 'Kol grade change from the first to the latest DPD check', file: 'buku_pinjaman.xlsx', sheetName: 'Pinjaman', sheet: S, base: S.cells, alt: S.cells, fill: true, fillCol: 'D', ansLabel: 'Kol change',
        books: [wb('pantauan_dpd.xlsx', 'DPD', DL, dlAlt), wb('kebijakan_kol.xlsx', 'Kol', KT)],
        statement: `<p>${opens([['the loan book', 'buku_pinjaman.xlsx'], ['the DPD monitoring log', 'pantauan_dpd.xlsx'], ['the Kol policy', 'kebijakan_kol.xlsx']])} Dimas records each facility's DPD (days past due) at every month end, oldest first. The Kol policy lists the lowest DPD of each grade, smallest first.</p><p>In <b>D5</b>, work out how many grades the facility's Kol moved: the Kol of its <b>latest</b> DPD in the log minus the Kol of its <b>first</b> DPD. A positive number means it got worse. Fill it down to <b>D${S.last}</b>.</p>`,
        solution: `=VLOOKUP(${latest},${KTT},2,TRUE)-VLOOKUP(${first},${KTT},2,TRUE)`,
        alts: [`=XLOOKUP(${latest},${KF},${KK},,-1)-XLOOKUP(${first},${KF},${KK},,-1)`, `=VLOOKUP(INDEX(${DD},XMATCH(A5,${DF},0,-1)),${KTT},2,TRUE)-VLOOKUP(INDEX(${DD},MATCH(A5,${DF},0)),${KTT},2,TRUE)`,
          `=LOOKUP(XLOOKUP(A5,${DF},${DD},,,-1),${KF},${KK})-LOOKUP(${first},${KF},${KK})`],
        hints: ['Four lookups in two files: the first and the latest DPD from the log, and the Kol grade of each from the policy.',
          `XLOOKUP(A5, ${DF}, ${DD}) is the first DPD; add , , 0, -1 to search from the bottom for the latest. VLOOKUP(dpd, ${KTT}, 2, TRUE) turns a DPD into its grade.`,
          `Shape: =VLOOKUP(XLOOKUP(A5, ${DF}, ${DD}, , 0, -1), ${KTT}, 2, TRUE) - VLOOKUP(XLOOKUP(A5, ${DF}, ${DD}), ${KTT}, 2, TRUE)`],
        explain: 'The XLOOKUP with search_mode -1 returns the facility\'s latest DPD in the log, the plain one its first. VLOOKUP with TRUE puts each in its Kol grade from the policy file, and the difference is how far the grade moved: 0 for no change, positive for a downgrade.',
        chips: ['A5', DF, DD, KTT, KF, KK, ',,0,-1)', ',2,TRUE)'], fnKey: 'VLOOKUP('
      });
    }
  });

  P.add({
    key: 'price-on-the-order-date', difficulty: 'Hard', level: 9, topics: ['XLOOKUP', 'VLOOKUP', 'approximate match', 'dates', 'Several workbooks'], count: 3, multi: true,
    make(r, v) {
      /* A schedule in another workbook, sorted by start date: the entry in force on a date is the last start on or before
         it. A third workbook brings the figure it is combined with. */
      const rows = H.sales(r, 14);
      const rs = rows.map((x) => H.toSerial(x.date));
      if (v === 0) {
        const per = startDates(rs, [3, 7, 11]);
        const skus = H.PRODUCTS.map((x) => x.sku);
        const pr = H.PRODUCTS.map((x) => { let c = x.carton_price; return per.map((d, i) => (i ? (c += r.pick([0, 2000, 4000, 6000])) : c)); });
        const cells = { A4: 'From' };
        skus.forEach((s, j) => { cells[L[j + 1] + '4'] = s; });
        per.forEach((d, i) => { cells['A' + (5 + i)] = d; skus.forEach((s, j) => { cells[L[j + 1] + (5 + i)] = pr[j][i]; }); });
        const HP = { cells, cols: L.slice(0, 13), num: L.slice(1, 13), dateCols: ['A'], first: 5, last: 4 + per.length };
        const S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'Distributor', k: 'distributor' }, { h: 'SKU', k: 'sku' }, { h: 'Cartons', k: 'cartons', num: true }], rows);
        const dl = r.shuffle(DIST());
        const MD = H.table([DC.id, DC.name, DC.region, DC.disc], dl);
        const at = (m) => tierAt(per, rs[m]), price = (m, i) => pr[skus.indexOf(rows[m].sku)][i];
        /* The hidden test: one order moves to exactly the start of the third period, at a different price; the price
           list corrects row 5's price, and row 6's distributor gets a bigger discount. */
        const km = Math.max(0, nums(14).findIndex((m) => rs[m] !== per[2] && at(m) !== 2 && price(m, at(m)) !== price(m, 2)));
        const alt = H.edit(S.cells, { [S.cell('Date', km)]: per[2] });
        const hpAlt = H.edit(cells, { [L[skus.indexOf(rows[0].sku) + 1] + (5 + at(0))]: (x) => x + 3000 });
        const mdAlt = H.edit(MD.cells, { [MD.cell('Discount (%)', dl.findIndex((d) => d.name === rows[1].distributor))]: (x) => x + 1 });
        const g = gridRefs('harga_periode.xlsx', 'Harga', HP), md = X('master_distributor.xlsx', 'Distributor', MD);
        const MN = md.A('Distributor'), MDc = md.A('Discount (%)'), disc = `XLOOKUP(B5,${MN},${MDc})`;
        const wrap = (p, d) => `=ROUND(D5*${p}*(1-${d}/100),0)`;
        return H.problem({
          title: 'Order value from a price-by-period sheet, after discount', file: 'pesanan_q4.xlsx', sheetName: 'Pesanan', sheet: S, base: S.cells, alt, fill: true, fillCol: 'E', ansLabel: 'Value (Rp)',
          books: [wb('harga_periode.xlsx', 'Harga', HP, hpAlt), wb('master_distributor.xlsx', 'Distributor', MD, mdAlt)],
          statement: `<p>${opens([['the order lines', 'pesanan_q4.xlsx'], ['the price list by period', 'harga_periode.xlsx'], ['the distributor master', 'master_distributor.xlsx']])} The price list has one row per period, with the date it starts in column A, oldest first, and one column per SKU across row 4. A period runs until the next one starts.</p><p>In <b>E5</b>, work out each line's value: its <b>Cartons</b> times the carton price of its SKU in the period that covers its <b>Date</b>, times (1 minus the distributor's <b>Discount (%)</b> divided by 100), rounded to whole rupiah with ROUND. Fill it down to <b>E${S.last}</b>.</p>`,
          solution: wrap(`INDEX(${g.G},MATCH(A5,${g.RK},1),MATCH(C5,${g.CK},0))`, disc),
          alts: [wrap(`XLOOKUP(A5,${g.RK},XLOOKUP(C5,${g.CK},${g.G}),,-1)`, disc), wrap(`VLOOKUP(A5,${g.ALL},MATCH(C5,${g.HEAD},0),TRUE)`, `VLOOKUP(B5,${md.T('Distributor', 'Discount (%)')},3,FALSE)`),
            wrap(`INDEX(${g.G},XMATCH(A5,${g.RK},-1),XMATCH(C5,${g.CK}))`, disc)],
          hints: ['The price is a two-way lookup with a twist: the SKU column is an exact match, but the period row is the last start date on or before the order date.',
            `MATCH(A5, ${g.RK}, 1) gives that row on the sorted dates; MATCH(C5, ${g.CK}, 0) gives the SKU column. The discount is a plain lookup in master_distributor.xlsx.`,
            `Shape: =ROUND(D5 * INDEX(${g.G}, MATCH(A5, ${g.RK}, 1), MATCH(C5, ${g.CK}, 0)) * (1 - XLOOKUP(B5, ${MN}, ${MDc})/100), 0)`],
          explain: 'Dates are numbers, so MATCH with 1 finds the last start date that is not after the order date: the period in force. An order on a start date gets the new price. The other MATCH finds the SKU across row 4, and INDEX returns the price where they cross. The distributor\'s discount from the master takes its share off the value.',
          chips: ['D5*', 'A5', 'B5', 'C5', g.G, g.RK, g.CK, MN, MDc, ',1)', ',0)'], fnKey: 'INDEX('
        });
      }
      if (v === 1) {
        rows.forEach((x) => { x.amount = r.pick([20, 25, 40, 50, 75, 100]) * 1000000; x.days = r.pick([30, 45, 60, 90]); });
        const per = startDates(rs, [r.int(2, 4), r.int(6, 8), r.int(10, 12)]);
        const side = per.map((d, k) => ({ d, x: [9.5, 9.75, 10.25, 10][k] + r.pick([-0.25, 0, 0.25]) }));
        const BR = H.table([{ h: 'From', k: 'd', date: true }, { h: 'Base rate (% a year)', k: 'x', num: true }], side);
        const ll = r.shuffle(DIST()).map((d) => ({ d: d.name, lim: r.pick([100, 150, 200, 250]) * 1000000, sp: r.pick([1.5, 2, 2.5, 3]) }));
        const LK = H.table([{ h: 'Distributor', k: 'd' }, { h: 'Limit (Rp)', k: 'lim', num: true }, { h: 'Spread (%)', k: 'sp', num: true }], ll);
        const S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'Distributor', k: 'distributor' }, { h: 'Amount (Rp)', k: 'amount', num: true }, { h: 'Days', k: 'days', num: true }], rows);
        /* The hidden test: one drawdown moves to exactly the start of the third period, the second base rate is revised,
           and row 6's distributor gets a wider spread. */
        const km = Math.max(0, nums(14).findIndex((m) => rs[m] !== per[2] && side[tierAt(per, rs[m])].x !== side[2].x));
        const alt = H.edit(S.cells, { [S.cell('Date', km)]: per[2] });
        const brAlt = H.edit(BR.cells, { [BR.cell('Base rate (% a year)', 1)]: (x) => x + 0.25 });
        const lkAlt = H.edit(LK.cells, { [LK.cell('Spread (%)', ll.findIndex((x) => x.d === rows[1].distributor))]: (x) => x + 0.5 });
        const br = X('suku_bunga_dasar.xlsx', 'Dasar', BR), lk = X('limit_kredit.xlsx', 'Limit', LK);
        const BF = br.A('From'), BB = br.A('Base rate (% a year)'), LD = lk.A('Distributor'), LS = lk.A('Spread (%)');
        const wrap = (b, s) => `=ROUND(C5*(${b}+${s})/100*D5/365,0)`;
        const spread = `XLOOKUP(B5,${LD},${LS})`;
        return H.problem({
          title: 'Drawdown interest at the base rate on its date plus the spread', file: 'pencairan_kredit.xlsx', sheetName: 'Pencairan', sheet: S, base: S.cells, alt, fill: true, fillCol: 'E', ansLabel: 'Interest (Rp)',
          books: [wb('suku_bunga_dasar.xlsx', 'Dasar', BR, brAlt), wb('limit_kredit.xlsx', 'Limit', LK, lkAlt)],
          statement: `<p>${opens([['the drawdowns', 'pencairan_kredit.xlsx'], ['the base rate history', 'suku_bunga_dasar.xlsx'], ['the limit file', 'limit_kredit.xlsx']])} The base rate history lists each change with the date it starts, oldest first; a rate stays in force until the next one starts. Each distributor pays its own <b>Spread (%)</b> on top, from the limit file.</p><p>In <b>E5</b>, work out the interest on each drawdown: its <b>Amount (Rp)</b> times (the base rate in force on its <b>Date</b> plus the distributor's spread), divided by 100, times <b>Days</b> divided by 365, rounded to whole rupiah with ROUND. Fill it down to <b>E${S.last}</b>.</p>`,
          solution: wrap(`XLOOKUP(A5,${BF},${BB},,-1)`, spread),
          alts: [wrap(`VLOOKUP(A5,${br.T('From', 'Base rate (% a year)')},2,TRUE)`, `VLOOKUP(B5,${lk.T('Distributor', 'Spread (%)')},3,FALSE)`), wrap(`LOOKUP(A5,${BF},${BB})`, spread),
            wrap(`INDEX(${BB},MATCH(A5,${BF},1))`, `INDEX(${LS},MATCH(B5,${LD},0))`)],
          hints: ['Two lookups in two files: the base rate by date (approximate: the last start on or before A5) and the spread by distributor (exact).',
            `XLOOKUP(A5, ${BF}, ${BB}, , -1) returns the exact date or the next earlier one. Add the spread inside brackets before multiplying.`,
            `Shape: =ROUND(C5 * (XLOOKUP(A5, ${BF}, ${BB}, , -1) + XLOOKUP(B5, ${LD}, ${LS})) / 100 * D5 / 365, 0)`],
          explain: 'With match_mode -1, XLOOKUP falls back to the latest start date before the drawdown when there is no exact match, which is the base rate in force; a drawdown on a start date gets the new rate. The second XLOOKUP brings the distributor\'s spread. Their sum is the yearly rate, applied to the amount for the days drawn out of 365.',
          chips: ['C5*', 'A5', 'B5', 'D5', BF, BB, LD, LS, ',,-1)', '/100*D5/365'], fnKey: 'XLOOKUP('
        });
      }
      const per = startDates(rs, [r.int(2, 3), r.int(5, 6), r.int(8, 9), r.int(11, 12)]);
      const pd = r.pick([5, 7.5]);
      const side = per.map((d, k) => ({ d, p: ['No promo', 'Payday Deal', 'No promo', 'Payday Deal', 'Year-end Sale'][k], x: [0, pd, 0, pd, r.pick([8, 10])][k] }));
      const KP = H.table([{ h: 'From', k: 'd', date: true }, { h: 'Promo', k: 'p' }, { h: 'Discount (%)', k: 'x', num: true }], side);
      const pl = r.shuffle(H.PRODUCTS);
      const PM = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Product', k: 'name' }, { h: 'Carton price (Rp)', k: 'carton_price', num: true }], pl);
      const S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'Distributor', k: 'distributor' }, { h: 'SKU', k: 'sku' }, { h: 'Cartons', k: 'cartons', num: true }], rows);
      /* The hidden test: one order moves to exactly the start of a Payday Deal, the first Payday Deal gets deeper, and
         the master raises the price of a SKU sold during a promo. */
      const km = Math.max(0, nums(14).findIndex((m) => rs[m] !== per[3] && side[tierAt(per, rs[m])].x !== side[3].x));
      const alt = H.edit(S.cells, { [S.cell('Date', km)]: per[3] });
      const kpAlt = H.edit(KP.cells, { [KP.cell('Discount (%)', 1)]: (x) => x + 2.5 });
      const kd = nums(14).find((m) => m !== km && side[tierAt(per, rs[m])].x > 0);
      const pmAlt = H.edit(PM.cells, { [PM.cell('Carton price (Rp)', pl.findIndex((x) => x.sku === rows[kd === undefined ? 5 : kd].sku))]: (x) => x + 4000 });
      const kp = X('kalender_promo.xlsx', 'Promo', KP), pm = X('master_produk.xlsx', 'Produk', PM);
      const KF = kp.A('From'), KD = kp.A('Discount (%)'), PS = pm.A('SKU'), PP = pm.A('Carton price (Rp)');
      const price = `XLOOKUP(C5,${PS},${PP})`;
      const wrap = (p, d) => `=ROUND(D5*${p}*${d}/100,0)`;
      return H.problem({
        title: 'Promo discount on the order date, prices from the master', file: 'pesanan_q4.xlsx', sheetName: 'Pesanan', sheet: S, base: S.cells, alt, fill: true, fillCol: 'E', ansLabel: 'Promo discount (Rp)',
        books: [wb('kalender_promo.xlsx', 'Promo', KP, kpAlt), wb('master_produk.xlsx', 'Produk', PM, pmAlt)],
        statement: `<p>${opens([['the order lines', 'pesanan_q4.xlsx'], ['the promo calendar', 'kalender_promo.xlsx'], ['the product master', 'master_produk.xlsx']])} The calendar lists each promo period with the date it starts, oldest first; it runs until the next one starts.</p><p>In <b>E5</b>, work out the promo discount on each line: its <b>Cartons</b> times the SKU's <b>Carton price (Rp)</b> from the master, times the <b>Discount (%)</b> of the promo running on its <b>Date</b>, divided by 100 and rounded to whole rupiah with ROUND. Fill it down to <b>E${S.last}</b>.</p>`,
        solution: wrap(price, `XLOOKUP(A5,${KF},${KD},,-1)`),
        alts: [wrap(`VLOOKUP(C5,${pm.T('SKU', 'Carton price (Rp)')},3,FALSE)`, `VLOOKUP(A5,${kp.T('From', 'Discount (%)')},3,TRUE)`), wrap(price, `LOOKUP(A5,${KF},${KD})`),
          wrap(`INDEX(${PP},MATCH(C5,${PS},0))`, `INDEX(${KD},MATCH(A5,${KF},1))`)],
        hints: ['Two lookups in two files: the carton price by SKU (exact) and the promo discount by date (approximate: the last start on or before A5).',
          `XLOOKUP(A5, ${KF}, ${KD}, , -1) returns the discount of the period in force; a date with no promo returns 0.`,
          `Shape: =ROUND(D5 * XLOOKUP(C5, ${PS}, ${PP}) * XLOOKUP(A5, ${KF}, ${KD}, , -1) / 100, 0)`],
        explain: 'XLOOKUP brings the carton price from the product master. With match_mode -1, the second XLOOKUP falls back to the latest promo start before the order date, so it returns the discount of the promo running then, and an order on a start date gets the new promo. Cartons times price times that %, divided by 100, is the discount.',
        chips: ['D5*', 'A5', 'C5', PS, PP, KF, KD, ',,-1)'], fnKey: 'XLOOKUP('
      });
    }
  });

  P.add({
    key: 'nested-lookup-chain', difficulty: 'Hard', level: 9, topics: ['XLOOKUP', 'VLOOKUP', 'nested lookup', 'Several workbooks'], count: 3, multi: true,
    make(r, v) {
      /* The answer is two or three lookups deep, each in a different workbook: the result of one lookup is what the next
         one searches for. Then the math. */
      if (v === 0) {
        const inv = H.invoices(r, 14);
        const S = H.table([{ h: 'Invoice', k: 'ref' }, { h: 'Date', k: 'date', date: true }, { h: 'Distributor', k: 'distributor' }, { h: 'Amount (Rp)', k: 'amount', num: true }], inv);
        const MD = H.table([DC.id, DC.name, DC.city, DC.region], r.shuffle(DIST()));
        const regs = r.shuffle(H.REGIONS).map((g) => ({ g, p: r.pick([0.5, 0.75, 1, 1.5, 2]) }));
        const IR = H.table([{ h: 'Region', k: 'g' }, { h: 'Incentive (%)', k: 'p', num: true }], regs);
        /* The hidden test: the incentive for row 5's region rises, and one invoice amount is corrected. */
        const irAlt = H.edit(IR.cells, { [IR.cell('Incentive (%)', regs.findIndex((x) => x.g === regionOf(inv[0].distributor)))]: (x) => x + 0.5 });
        const alt = H.edit(S.cells, { [S.cell('Amount (Rp)', 4)]: (x) => x + 2500000 });
        const md = X('master_distributor.xlsx', 'Distributor', MD), ir = X('insentif_regional.xlsx', 'Insentif', IR);
        const MN = md.A('Distributor'), MR = md.A('Region'), IG = ir.A('Region'), IP = ir.A('Incentive (%)');
        const inner = `XLOOKUP(C5,${MN},${MR})`;
        return H.problem({
          title: 'Regional incentive on each invoice, through the distributor master', file: 'faktur_q4.xlsx', sheetName: 'Faktur', sheet: S, base: S.cells, alt, fill: true, fillCol: 'E', ansLabel: 'Incentive (Rp)',
          books: [wb('master_distributor.xlsx', 'Distributor', MD), wb('insentif_regional.xlsx', 'Insentif', IR, irAlt)],
          statement: `<p>${opens([['the invoices', 'faktur_q4.xlsx'], ['the distributor master', 'master_distributor.xlsx'], ['the regional incentives', 'insentif_regional.xlsx']])} Rasa Kenari pays distributors an incentive set per region. The invoices do not show the region.</p><p>In <b>E5</b>, work out the incentive on each invoice: its <b>Amount (Rp)</b> times the <b>Incentive (%)</b> of the distributor's region, divided by 100 and rounded to whole rupiah with ROUND. Fill it down to <b>E${S.last}</b>.</p>`,
          solution: `=ROUND(D5*XLOOKUP(${inner},${IG},${IP})/100,0)`,
          alts: [`=ROUND(D5*VLOOKUP(VLOOKUP(C5,${md.T('Distributor', 'Region')},3,FALSE),${ir.T('Region', 'Incentive (%)')},2,FALSE)/100,0)`,
            `=ROUND(D5*INDEX(${IP},MATCH(INDEX(${MR},MATCH(C5,${MN},0)),${IG},0))/100,0)`, `=ROUND(XLOOKUP(${inner},${IG},${IP})*D5/100,0)`],
          hints: ['Step one: which region is the distributor in? That is in master_distributor.xlsx. Step two: what is that region\'s incentive? That is in insentif_regional.xlsx.',
            `The inner lookup, ${inner.replace(/,/g, ', ')}, returns a region. Use it as the value to find in ${IG}.`,
            `Shape: =ROUND(D5 * XLOOKUP(XLOOKUP(C5, ${MN}, ${MR}), ${IG}, ${IP}) / 100, 0)`],
          explain: 'The inner XLOOKUP turns the distributor into its region from the master; the outer one finds that region in the incentive file and returns its percentage. Amount times that %, divided by 100 and rounded, is the incentive.',
          chips: ['D5*', 'C5', MN, MR, IG, IP, md.T('Distributor', 'Region'), ir.T('Region', 'Incentive (%)')], fnKey: 'XLOOKUP('
        });
      }
      if (v === 1) {
        const rows = H.sales(r, 14);
        const pl = r.shuffle(H.PRODUCTS);
        const base4 = [r.pick([16, 18, 20]), r.pick([24, 26, 28]), r.pick([30, 32, 35]), r.pick([25, 27, 29])];
        const cats = r.shuffle(H.CATEGORIES.map((c, k) => ({ c, m: base4[k] })));
        const S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'SKU', k: 'sku' }, { h: 'Cartons', k: 'cartons', num: true }, { h: 'Gross (Rp)', k: 'gross', num: true }], rows);
        const PM = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Product', k: 'name' }, { h: 'Category', k: 'category' }], pl);
        const MK = H.table([{ h: 'Category', k: 'c' }, { h: 'Margin (%)', k: 'm', num: true }], cats);
        const catOf = (sku) => H.PRODUCTS.find((x) => x.sku === sku).category;
        /* The hidden test: finance raises the margin of row 5's category, and row 8 turns out to be a SKU of another
           category. */
        const mkAlt = H.edit(MK.cells, { [MK.cell('Margin (%)', cats.findIndex((x) => x.c === catOf(rows[0].sku)))]: (x) => x + 2 });
        const swap = r.pick(H.PRODUCTS.filter((x) => x.category !== catOf(rows[3].sku)));
        const alt = H.edit(S.cells, { [S.cell('SKU', 3)]: swap.sku, [S.cell('Gross (Rp)', 3)]: rows[3].cartons * swap.carton_price });
        const pm = X('master_produk.xlsx', 'Produk', PM), mk = X('margin_kategori.xlsx', 'Margin', MK);
        const PS = pm.A('SKU'), PC = pm.A('Category'), KC = mk.A('Category'), KM = mk.A('Margin (%)');
        const inner = `XLOOKUP(B5,${PS},${PC})`;
        return H.problem({
          title: 'Line margin through the product master and category margins', file: 'margin_q4.xlsx', sheetName: 'Penjualan', sheet: S, base: S.cells, alt, fill: true, fillCol: 'E', ansLabel: 'Margin (Rp)',
          books: [wb('master_produk.xlsx', 'Produk', PM), wb('margin_kategori.xlsx', 'Margin', MK, mkAlt)],
          statement: `<p>${opens([['the sales lines', 'margin_q4.xlsx'], ['the product master', 'master_produk.xlsx'], ['the category margins', 'margin_kategori.xlsx']])} Margins are set per category, not per SKU.</p><p>In <b>E5</b>, work out the margin in rupiah of each line: its <b>Gross (Rp)</b> times the <b>Margin (%)</b> of its SKU's category, divided by 100 and rounded to whole rupiah with ROUND. The SKU's category is in the product master. Fill it down to <b>E${S.last}</b>.</p>`,
          solution: `=ROUND(D5*XLOOKUP(${inner},${KC},${KM})/100,0)`,
          alts: [`=ROUND(D5*VLOOKUP(VLOOKUP(B5,${pm.T('SKU', 'Category')},3,FALSE),${mk.T('Category', 'Margin (%)')},2,FALSE)/100,0)`,
            `=ROUND(D5*INDEX(${KM},MATCH(INDEX(${PC},MATCH(B5,${PS},0)),${KC},0))/100,0)`],
          hints: ['Two steps in two files: the SKU gives a category (master_produk.xlsx), and the category gives a margin (margin_kategori.xlsx).',
            `The inner lookup ${inner.replace(/,/g, ', ')} returns the category; search for it in ${KC}.`,
            `Shape: =ROUND(D5 * XLOOKUP(XLOOKUP(B5, ${PS}, ${PC}), ${KC}, ${KM}) / 100, 0)`],
          explain: 'The inner XLOOKUP turns the SKU into its category from the product master; the outer one finds that category\'s margin in the margin file. Times the gross, divided by 100 and rounded, that is the line\'s margin in rupiah.',
          chips: ['D5*', 'B5', PS, PC, KC, KM, pm.T('SKU', 'Category'), mk.T('Category', 'Margin (%)')], fnKey: 'XLOOKUP('
        });
      }
      const outs = r.sample(H.db().outlets, 16).sort((a, b) => (a.outlet_id < b.outlet_id ? -1 : 1));
      const OL = H.table([{ h: 'Outlet ID', k: 'outlet_id' }, { h: 'Outlet', k: 'outlet' }, { h: 'Distributor ID', k: 'distributor_id' }], outs);
      const MD = H.table([DC.id, DC.name, DC.region], r.shuffle(DIST()));
      const rates = r.shuffle(H.REGIONS).map((g) => ({ g, x: r.int(8, 30) * 100 }));
      const TK = H.table([{ h: 'Region', k: 'g' }, { h: 'Rate (Rp per carton)', k: 'x', num: true }], rates);
      const rows = H.sales(r, 14).map((s) => ({ date: s.date, o: r.pick(outs).outlet_id, cartons: r.int(2, 30) }));
      const S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'Outlet ID', k: 'o' }, { h: 'Cartons', k: 'cartons', num: true }], rows);
      const regOfOutlet = (id) => DIST().find((d) => d.id === outs.find((o) => o.outlet_id === id).distributor_id).region;
      /* The hidden test: the rate card raises row 5's region, and the outlet list moves row 7's outlet to a distributor
         in another region. */
      const tkAlt = H.edit(TK.cells, { [TK.cell('Rate (Rp per carton)', rates.findIndex((x) => x.g === regOfOutlet(rows[0].o)))]: (x) => x + 300 });
      const nd = r.pick(DIST().filter((d) => d.region !== regOfOutlet(rows[2].o)));
      const olAlt = H.edit(OL.cells, { [OL.cell('Distributor ID', outs.findIndex((o) => o.outlet_id === rows[2].o))]: nd.id });
      const ol = X('daftar_outlet.xlsx', 'Outlet', OL), md = X('master_distributor.xlsx', 'Distributor', MD), tk = X('tarif_kirim.xlsx', 'Tarif', TK);
      const OI = ol.A('Outlet ID'), OD = ol.A('Distributor ID'), MI = md.A('ID'), MR = md.A('Region'), TR = tk.A('Region'), TV = tk.A('Rate (Rp per carton)');
      const did = `XLOOKUP(B5,${OI},${OD})`, reg = `XLOOKUP(${did},${MI},${MR})`;
      return H.problem({
        title: 'Delivery cost per outlet order, through three other workbooks', file: 'pesanan_outlet.xlsx', sheetName: 'Pesanan', sheet: S, base: S.cells, alt: S.cells, fill: true, fillCol: 'D', ansLabel: 'Delivery (Rp)',
        books: [wb('daftar_outlet.xlsx', 'Outlet', OL, olAlt), wb('master_distributor.xlsx', 'Distributor', MD), wb('tarif_kirim.xlsx', 'Tarif', TK, tkAlt)],
        statement: `<p>${opens([['the outlet orders', 'pesanan_outlet.xlsx'], ['the outlet list', 'daftar_outlet.xlsx'], ['the distributor master', 'master_distributor.xlsx'], ['the delivery rate card', 'tarif_kirim.xlsx']])} An outlet buys from a distributor (outlet list), the distributor sits in a region (master), and delivery is priced per carton by region (rate card).</p><p>In <b>D5</b>, work out the delivery cost of each order: its <b>Cartons</b> times the <b>Rate (Rp per carton)</b> of the region the outlet's distributor sits in. Fill it down to <b>D${S.last}</b>.</p>`,
        solution: `=C5*XLOOKUP(${reg},${TR},${TV})`,
        alts: [`=C5*VLOOKUP(VLOOKUP(VLOOKUP(B5,${ol.T('Outlet ID', 'Distributor ID')},3,FALSE),${md.T('ID', 'Region')},3,FALSE),${tk.T('Region', 'Rate (Rp per carton)')},2,FALSE)`,
          `=C5*INDEX(${TV},MATCH(INDEX(${MR},MATCH(INDEX(${OD},MATCH(B5,${OI},0)),${MI},0)),${TR},0))`, `=XLOOKUP(${reg},${TR},${TV})*C5`],
        hints: ['Work outwards, one workbook at a time: outlet to distributor ID (daftar_outlet.xlsx), distributor ID to region (master_distributor.xlsx), region to rate (tarif_kirim.xlsx).',
          `Build it in steps: ${did.replace(/,/g, ', ')} gives the distributor ID; wrap it in a lookup on ${MI} that returns ${MR}, the region.`,
          `Shape: =C5 * XLOOKUP(XLOOKUP(XLOOKUP(B5, ${OI}, ${OD}), ${MI}, ${MR}), ${TR}, ${TV})`],
        explain: 'The innermost XLOOKUP turns the outlet into its distributor ID, the middle one turns that ID into a region, and the outer one turns the region into its delivery rate. Each answer is the value the next lookup searches for, in the next workbook. Times the cartons, that is the delivery cost.',
        chips: ['C5*', 'B5', OI, OD, MI, MR, TR, TV, ol.T('Outlet ID', 'Distributor ID'), md.T('ID', 'Region'), tk.T('Region', 'Rate (Rp per carton)')], fnKey: 'XLOOKUP('
      });
    }
  });
});
