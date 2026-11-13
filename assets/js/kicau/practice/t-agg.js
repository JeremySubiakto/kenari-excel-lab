/*
 * Practice, Conditional totals (20 Easy, 40 Medium, 10 Hard): SUMIF(S), COUNTIF(S), AVERAGEIF(S), MAXIFS, MINIFS,
 * criteria from a cell, not equal, wildcards, date windows and blanks; the Hard ones add two-way summary tables,
 * either-or totals and month-against-month windows.
 * Each template makes `count` problems; v (0 to count-1) picks the variant and r is its seeded random.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) factory(require('./core.js'));
  else factory(root.KicauPractice);
})(typeof globalThis !== 'undefined' ? globalThis : this, function (P) {
  'use strict';
  const H = P.H;

  /* ---------- small helpers for this file ---------- */
  const byDate = (a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : 0);
  /* Sales lines from the shared world, dated in Q4 2026, optionally only the rows keep() allows. */
  function lines(r, n, keep) {
    const pool = H.db().sales.filter(keep || (() => true));
    return r.sample(pool, n).map((s) => Object.assign({}, s, { date: H.Q4(s.date) })).sort(byDate);
  }
  const names = (r, n) => r.sample(H.db().distributors, n).map((d) => d.name);
  /* The values of a column, most frequent first (ties alphabetical). */
  function ranked(rows, k) {
    const c = {};
    rows.forEach((x) => { const key = typeof k === 'function' ? k(x) : x[k]; c[key] = (c[key] || 0) + 1; });
    return Object.keys(c).sort((a, b) => c[b] - c[a] || (a < b ? -1 : 1));
  }
  const where = (rows, f) => rows.map((x, i) => (f(x, i) ? i : -1)).filter((i) => i >= 0);
  const addDays = (d, n) => H.iso(H.toSerial(d) + n);
  const step = (h) => (/\(Rp\)/.test(h) ? 1500000 : h === 'Rate (%)' ? 0.5 : 6);
  const sorted = (rows, k) => rows.map((x) => x[k]).sort((a, b) => (a < b ? -1 : a > b ? 1 : 0));
  const word = (product) => product.split(' ')[1];
  const family = (sku) => sku.slice(0, 3);
  const month = (d) => +d.slice(5, 7);
  const sumOf = (rows, f, k) => rows.filter(f).reduce((a, x) => a + x[k], 0);
  /* Invoices for some distributors: ref ending in a city code, date, due date from the terms, amount, paid on or blank. */
  function bills(r, n, who) {
    const ds = H.db().distributors;
    return lines(r, n, (s) => who.includes(s.distributor)).map((s, i) => {
      const d = ds.find((x) => x.name === s.distributor);
      return { ref: 'INV-' + (2101 + i * 3 + r.int(0, 2)) + '-' + d.city.slice(0, 3).toUpperCase(), distributor: s.distributor, city: d.city, code: d.city.slice(0, 3).toUpperCase(),
        date: s.date, due: addDays(s.date, d.terms_days), amount: s.net, paid: r.chance(0.42) ? '' : addDays(s.date, r.int(6, d.terms_days + 8)) };
    });
  }
  const COL = {
    date: { h: 'Date', k: 'date', date: true }, dist: { h: 'Distributor', k: 'distributor' }, region: { h: 'Region', k: 'region' },
    product: { h: 'Product', k: 'product' }, category: { h: 'Category', k: 'category' }, sku: { h: 'SKU', k: 'sku' },
    cartons: { h: 'Cartons', k: 'cartons', num: true }, net: { h: 'Net (Rp)', k: 'net', num: true },
    facility: { h: 'Facility', k: 'facility' }, type: { h: 'Type', k: 'type' }, limit: { h: 'Limit (Rp)', k: 'limit', num: true },
    outstanding: { h: 'Outstanding (Rp)', k: 'outstanding', num: true }, rate: { h: 'Rate (%)', k: 'rate', num: true }, dpd: { h: 'DPD', k: 'dpd', num: true },
    restructured: { h: 'Restructured', k: 'restructured' }, start: { h: 'Start', k: 'start', date: true },
    warehouse: { h: 'Warehouse', k: 'warehouse' }, onhand: { h: 'On hand', k: 'onhand', num: true },
    rep: { h: 'Rep', k: 'name' }, area: { h: 'Area', k: 'area' }, visits: { h: 'Visits', k: 'visits', num: true },
    ref: { h: 'Ref', k: 'ref' }, due: { h: 'Due', k: 'due', date: true }, amount: { h: 'Amount (Rp)', k: 'amount', num: true }, paid: { h: 'Paid on', k: 'paid', date: true }
  };
  const repsIn = (r, n) => { const g = r.sample(H.REGIONS, 3); return H.reps(r, n).map((x) => Object.assign({}, x, { area: r.pick(g) })); };
  const total = (xs) => xs.reduce((a, x) => a + x, 0);
  const mean = (xs) => (xs.length ? total(xs) / xs.length : NaN);
  /*
   * The pair the hidden sheet asks for, when B1 and B2 drive the answer. Both parts differ from the visible pair, one
   * figure of the new pair is raised by `bump`, and the new answer differs from the visible one and from what a formula
   * with B1 or B2 typed in would give, so neither input can be hard-coded. Returns the pair and the row to raise.
   */
  function hiddenPair(rows, ka, kb, combos, vk, bump, agg) {
    const [a0, b0] = combos[0];
    const rowOf = (c) => rows.findIndex((x) => x[ka] === c[0] && x[kb] === c[1]);
    for (const c of combos.slice(1)) {
      if (c[0] === a0 || c[1] === b0) continue;
      const j = rowOf(c);
      const alt = rows.map((x, i) => (i === j ? Object.assign({}, x, { [vk]: x[vk] + bump }) : x));
      const f = (a, b) => agg(alt.filter((x) => x[ka] === a && x[kb] === b).map((x) => x[vk]));
      const want = f(c[0], c[1]);
      if (want !== f(a0, b0) && want !== f(a0, c[1]) && want !== f(c[0], b0)) return { pair: c, j };
    }
    return { pair: combos[1], j: rowOf(combos[1]) };
  }

  /* ================= Easy ================= */

  P.add({
    key: 'count-rows-with-one-label', difficulty: 'Easy', level: 2, topics: ['COUNTIF'], count: 4,
    make(r, v) {
      const V = [
        { title: 'Count the sales lines from one region', file: 'penjualan_q4.xlsx', h: 'Region', k: 'region', what: 'sales lines',
          rows: () => { const g = r.sample(H.REGIONS, 4); return lines(r, 14, (s) => g.includes(s.region)); }, cols: [COL.date, COL.dist, COL.region, COL.net] },
        { title: 'Count the facilities of one loan type', file: 'buku_pinjaman.xlsx', h: 'Type', k: 'type', what: 'facilities',
          rows: () => H.loans(r, 16), cols: [COL.facility, COL.dist, COL.type, COL.outstanding] },
        { title: 'Count the stock lines kept in one warehouse', file: 'stok_gudang.xlsx', h: 'Warehouse', k: 'warehouse', what: 'stock lines',
          rows: () => H.stock(r, 14), cols: [COL.sku, COL.product, COL.warehouse, COL.onhand] },
        { title: 'Count the sales lines in one product category', file: 'penjualan_q4.xlsx', h: 'Category', k: 'category', what: 'sales lines',
          rows: () => lines(r, 14), cols: [COL.date, COL.product, COL.category, COL.cartons] }
      ][v];
      const rows = V.rows(), S = H.table(V.cols, rows), keys = ranked(rows, V.k);
      const want = keys[1] || keys[0], R = S.R(V.h);
      const base = Object.assign({}, S.cells, { A1: 'Count: ' + want });
      /* The hidden sheet: two rows are corrected to the asked-for label, so the count rises by two. */
      const others = where(rows, (x) => x[V.k] !== want);
      const alt = H.edit(base, { [S.cell(V.h, others[0])]: want, [S.cell(V.h, others[others.length - 1])]: want });
      return H.problem({
        title: V.title, file: V.file, sheet: S, base, alt, answerCell: 'B1',
        statement: '<p>In <b>B1</b>, count how many ' + V.what + ' have <b>' + want + '</b> in the <b>' + V.h + '</b> column (' + R + ').</p>',
        solution: '=COUNTIF(' + R + ',"' + want + '")', alts: ['=COUNTIF(' + S.A(V.h) + ',"' + want + '")', '=COUNTIFS(' + R + ',"' + want + '")'],
        hints: ['COUNTIF counts the cells in a range that meet one condition.', 'The range is ' + R + '. The condition is the text ' + want + ', written inside double quotes.', 'Shape: =COUNTIF(' + R + ',"' + want + '").'],
        explain: 'COUNTIF checks every cell in ' + R + ' against "' + want + '" and counts the matches. Text in a condition goes inside double quotes, and capitals do not matter.',
        chips: [R, '"' + want + '"'], fnKey: 'COUNTIF('
      });
    }
  });

  P.add({
    key: 'total-for-the-name-in-b1', difficulty: 'Easy', level: 2, topics: ['SUMIF'], count: 4,
    make(r, v) {
      const V = [
        { title: 'Net sales of the distributor named in B1', file: 'penjualan_q4.xlsx', label: 'Distributor', h: 'Distributor', k: 'distributor', sum: 'Net (Rp)', ans: 'Net sales (Rp)',
          rows: () => { const g = names(r, 4); return lines(r, 14, (s) => g.includes(s.distributor)); }, cols: [COL.date, COL.dist, COL.product, COL.net] },
        { title: 'Cartons sold of the product named in B1', file: 'penjualan_q4.xlsx', label: 'Product', h: 'Product', k: 'product', sum: 'Cartons', ans: 'Cartons',
          rows: () => { const g = r.sample(H.PRODUCTS, 4).map((p) => p.name); return lines(r, 14, (s) => g.includes(s.product)); }, cols: [COL.date, COL.dist, COL.product, COL.cartons] },
        { title: 'Outstanding loans of the region named in B1', file: 'buku_pinjaman.xlsx', label: 'Region', h: 'Region', k: 'region', sum: 'Outstanding (Rp)', ans: 'Outstanding (Rp)',
          rows: () => H.loans(r, 16), cols: [COL.facility, COL.dist, COL.region, COL.outstanding] },
        { title: 'Cartons on hand in the warehouse named in B1', file: 'stok_gudang.xlsx', label: 'Warehouse', h: 'Warehouse', k: 'warehouse', sum: 'On hand', ans: 'On hand (cartons)',
          rows: () => H.stock(r, 14), cols: [COL.sku, COL.product, COL.warehouse, COL.onhand] }
      ][v];
      const rows = V.rows(), S = H.table(V.cols, rows), keys = ranked(rows, V.k);
      const c = S.R(V.h), s = S.R(V.sum);
      const base = Object.assign({}, S.cells, { A1: V.label, B1: keys[0], A2: V.ans });
      /* The hidden sheet: B1 names someone else, and one of their figures is corrected. */
      const j = rows.findIndex((x) => x[V.k] === keys[1]);
      const alt = H.edit(base, { B1: keys[1], [S.cell(V.sum, j)]: (x) => x + step(V.sum) });
      return H.problem({
        title: V.title, file: V.file, sheet: S, base, alt, answerCell: 'B2',
        statement: '<p>B1 holds a ' + V.label.toLowerCase() + '. In <b>B2</b>, add up the <b>' + V.sum + '</b> column (' + s + ') for the rows whose <b>' + V.h + '</b> (' + c + ') matches B1.</p>',
        solution: '=SUMIF(' + c + ',B1,' + s + ')', alts: ['=SUMIFS(' + s + ',' + c + ',B1)', '=SUMIF(' + S.A(V.h) + ',$B$1,' + S.A(V.sum) + ')'],
        hints: ['SUMIF adds the numbers in one column for the rows that meet a condition in another column.', 'Point at B1 for the condition instead of typing the name, so the total follows B1 when it changes.', 'Shape: =SUMIF(' + c + ',B1,' + s + ').'],
        explain: 'SUMIF looks down ' + c + ' for the value in B1 and adds the matching cells of ' + s + '. Because the condition is a cell, the same formula works for any name you put in B1.',
        chips: [c, s, 'B1'], fnKey: 'SUMIF('
      });
    }
  });

  P.add({
    key: 'count-past-the-limit-in-b1', difficulty: 'Easy', level: 3, topics: ['COUNTIF'], count: 4,
    make(r, v) {
      const V = [
        { title: 'Count orders bigger than the carton limit in B1', file: 'penjualan_q4.xlsx', h: 'Cartons', op: '>', words: 'more than', label: 'Limit (cartons)', ans: 'Orders above it',
          rows: () => lines(r, 14), cols: [COL.date, COL.dist, COL.product, COL.cartons], at: [7, 4] },
        { title: 'Count facilities with at least the days past due in B1', file: 'buku_pinjaman.xlsx', h: 'DPD', op: '>=', words: 'at least', label: 'Days past due', ans: 'Facilities',
          rows: () => H.loans(r, 16), cols: [COL.facility, COL.dist, COL.outstanding, COL.dpd], fixed: [30, 60] },
        { title: 'Count stock lines below the on-hand level in B1', file: 'stok_gudang.xlsx', h: 'On hand', op: '<', words: 'less than', label: 'Level (cartons)', ans: 'Lines below it',
          rows: () => H.stock(r, 14), cols: [COL.sku, COL.product, COL.warehouse, COL.onhand], fixed: [150, 260] },
        { title: 'Count reps with at least the visits in B1', file: 'tim_sales.xlsx', h: 'Visits', op: '>=', words: 'at least', label: 'Visits', ans: 'Reps',
          rows: () => H.reps(r, 12), cols: [COL.rep, COL.area, COL.visits], at: [6, 3] }
      ][v];
      const rows = V.rows(), S = H.table(V.cols, rows), R = S.R(V.h), k = V.cols.find((c) => c.h === V.h).k;
      const vals = sorted(rows, k);
      const lim = V.fixed || [vals[V.at[0]], vals[V.at[1]]];
      const base = Object.assign({}, S.cells, { A1: V.label, B1: lim[0], A2: V.ans });
      /* The hidden sheet: a new limit in B1, and one row's figure is corrected. For "less than" the corrected row lands
         exactly on the new limit, so <= in place of < gives a different count. */
      const j = rows.findIndex((x) => x[k] === vals[0]);
      const alt = H.edit(base, { B1: lim[1], [S.cell(V.h, j)]: V.op === '<' ? lim[1] : (x) => x + 120 });
      return H.problem({
        title: V.title, file: V.file, sheet: S, base, alt, answerCell: 'B2',
        statement: '<p>In <b>B2</b>, count the rows whose <b>' + V.h + '</b> (' + R + ') is <b>' + V.words + '</b> the number in <b>B1</b>. Use B1 in the formula, so the count follows it.</p>',
        solution: '=COUNTIF(' + R + ',"' + V.op + '"&B1)', alts: ['=COUNTIFS(' + R + ',"' + V.op + '"&B1)', '=COUNTIF(' + S.A(V.h) + ',"' + V.op + '"&$B$1)'],
        hints: ['COUNTIF can compare numbers: "' + V.op + '100" counts the cells that are ' + V.words + ' 100.', 'The limit is in B1, so join the sign to the cell with &: "' + V.op + '"&B1. Writing "' + V.op + 'B1" would compare against the letters B1.', 'Shape: =COUNTIF(' + R + ',"' + V.op + '"&B1).'],
        explain: '"' + V.op + '"&B1 builds a condition such as "' + V.op + lim[0] + '" from the sign and the cell, and COUNTIF counts the cells of ' + R + ' that pass it. "' + V.words[0].toUpperCase() + V.words.slice(1) + '" is ' + V.op + ', so a row equal to B1 is ' + (V.op.includes('=') ? 'counted' : 'left out') + '.',
        chips: [R, '"' + V.op + '"&B1'], fnKey: 'COUNTIF('
      });
    }
  });

  P.add({
    key: 'average-for-the-group-in-b1', difficulty: 'Easy', level: 3, topics: ['AVERAGEIF'], count: 4,
    make(r, v) {
      const V = [
        { title: 'Average cartons per line for the category in B1', file: 'penjualan_q4.xlsx', label: 'Category', h: 'Category', k: 'category', avg: 'Cartons', ans: 'Average cartons',
          rows: () => lines(r, 14), cols: [COL.date, COL.product, COL.category, COL.cartons] },
        { title: 'Average interest rate for the loan type in B1', file: 'buku_pinjaman.xlsx', label: 'Type', h: 'Type', k: 'type', avg: 'Rate (%)', ans: 'Average rate (%)',
          rows: () => H.loans(r, 16), cols: [COL.facility, COL.dist, COL.type, COL.rate] },
        { title: 'Average on-hand cartons for the warehouse in B1', file: 'stok_gudang.xlsx', label: 'Warehouse', h: 'Warehouse', k: 'warehouse', avg: 'On hand', ans: 'Average on hand',
          rows: () => H.stock(r, 14), cols: [COL.sku, COL.product, COL.warehouse, COL.onhand] },
        { title: 'Average outlet visits for the sales area in B1', file: 'tim_sales.xlsx', label: 'Area', h: 'Area', k: 'area', avg: 'Visits', ans: 'Average visits',
          rows: () => repsIn(r, 12), cols: [COL.rep, COL.area, COL.visits] }
      ][v];
      const rows = V.rows(), S = H.table(V.cols, rows), keys = ranked(rows, V.k);
      const c = S.R(V.h), a = S.R(V.avg);
      const base = Object.assign({}, S.cells, { A1: V.label, B1: keys[0], A2: V.ans });
      const j = rows.findIndex((x) => x[V.k] === keys[1]);
      const alt = H.edit(base, { B1: keys[1], [S.cell(V.avg, j)]: (x) => x + step(V.avg) });
      return H.problem({
        title: V.title, file: V.file, sheet: S, base, alt, answerCell: 'B2',
        statement: '<p>In <b>B2</b>, work out the average <b>' + V.avg + '</b> (' + a + ') of the rows whose <b>' + V.h + '</b> (' + c + ') matches B1. Keep the decimals; do not round.</p>',
        solution: '=AVERAGEIF(' + c + ',B1,' + a + ')', alts: ['=AVERAGEIFS(' + a + ',' + c + ',B1)', '=SUMIF(' + c + ',B1,' + a + ')/COUNTIF(' + c + ',B1)'],
        hints: ['AVERAGEIF works like SUMIF, but returns the average of the matching rows.', 'The condition is B1, the range to test is ' + c + ' and the numbers to average are in ' + a + '.', 'Shape: =AVERAGEIF(' + c + ',B1,' + a + ').'],
        explain: 'AVERAGEIF keeps the rows of ' + c + ' that match B1, then averages their ' + V.avg + '. It is the same as SUMIF divided by COUNTIF with the same condition.',
        chips: [c, a, 'B1'], fnKey: 'AVERAGEIF('
      });
    }
  });

  P.add({
    key: 'total-past-the-limit-in-b1', difficulty: 'Easy', level: 3, topics: ['SUMIF'], count: 4,
    make(r, v) {
      const V = [
        { title: 'Net sales from lines worth at least the amount in B1', file: 'penjualan_q4.xlsx', h: 'Net (Rp)', op: '>=', words: 'at least', label: 'Floor (Rp)', ans: 'Total (Rp)',
          rows: () => lines(r, 14), cols: [COL.date, COL.dist, COL.product, COL.net], at: [8, 5], d: 2000000 },
        { title: 'Outstanding of facilities above the balance in B1', file: 'buku_pinjaman.xlsx', h: 'Outstanding (Rp)', op: '>', words: 'more than', label: 'Balance (Rp)', ans: 'Total (Rp)',
          rows: () => H.loans(r, 16), cols: [COL.facility, COL.dist, COL.limit, COL.outstanding], at: [9, 6], d: 5000000 },
        { title: 'Cartons in small orders under the size in B1', file: 'penjualan_q4.xlsx', h: 'Cartons', op: '<', words: 'fewer than', label: 'Size (cartons)', ans: 'Total cartons',
          rows: () => lines(r, 14), cols: [COL.date, COL.dist, COL.product, COL.cartons], at: [6, 9], d: -1 },
        { title: 'On-hand cartons of stock lines at most the level in B1', file: 'stok_gudang.xlsx', h: 'On hand', op: '<=', words: 'at most', label: 'Level (cartons)', ans: 'Total cartons',
          rows: () => H.stock(r, 14), cols: [COL.sku, COL.product, COL.warehouse, COL.onhand], at: [7, 4], d: -10 }
      ][v];
      const rows = V.rows(), S = H.table(V.cols, rows), R = S.R(V.h), k = V.cols.find((c) => c.h === V.h).k;
      const vals = sorted(rows, k), lim = [vals[V.at[0]], vals[V.at[1]]];
      const base = Object.assign({}, S.cells, { A1: V.label, B1: lim[0], A2: V.ans });
      /* The hidden sheet: a new limit, and the row that sits at the old limit is corrected. */
      const j = rows.findIndex((x) => x[k] === lim[0]);
      const alt = H.edit(base, { B1: lim[1], [S.cell(V.h, j)]: (x) => Math.max(1, x + V.d) });
      return H.problem({
        title: V.title, file: V.file, sheet: S, base, alt, answerCell: 'B2',
        statement: '<p>In <b>B2</b>, add up the <b>' + V.h + '</b> values (' + R + ') that are <b>' + V.words + '</b> the number in <b>B1</b>.</p>',
        solution: '=SUMIF(' + R + ',"' + V.op + '"&B1)', alts: ['=SUMIF(' + R + ',"' + V.op + '"&B1,' + R + ')', '=SUMIFS(' + R + ',' + R + ',"' + V.op + '"&B1)'],
        hints: ['When SUMIF gets only two arguments, it adds the same cells it tests.', '"' + V.words[0].toUpperCase() + V.words.slice(1) + '" is ' + V.op + '. Join it to the cell: "' + V.op + '"&B1.', 'Shape: =SUMIF(' + R + ',"' + V.op + '"&B1).'],
        explain: 'SUMIF tests each cell of ' + R + ' against "' + V.op + '"&B1 and, with no third range, adds the cells that pass. A value equal to B1 is ' + (V.op.includes('=') ? 'included' : 'left out') + ' because the sign is ' + V.op + '.',
        chips: [R, '"' + V.op + '"&B1'], fnKey: 'SUMIF('
      });
    }
  });

  /* ================= Medium ================= */

  P.add({
    key: 'everything-except-one', difficulty: 'Medium', level: 4, topics: ['SUMIF', 'COUNTIF', 'AVERAGEIF'], count: 4,
    make(r, v) {
      if (v === 3) {
        /* Not equal to a number: facilities with any days past due. */
        const rows = H.loans(r, 16), S = H.table([COL.facility, COL.dist, COL.outstanding, COL.dpd], rows), R = S.R('DPD');
        const base = Object.assign({}, S.cells, { A1: 'Facilities past due' });
        const zero = where(rows, (x) => x.dpd === 0);
        const fix = zero.length ? { [S.cell('DPD', zero[0])]: 15 } : {};
        if (zero.length > 1) fix[S.cell('DPD', zero[zero.length - 1])] = 45;
        if (!zero.length) fix[S.cell('DPD', 0)] = 0;
        const alt = H.edit(base, fix);
        return H.problem({
          title: 'Count facilities whose DPD is not zero', file: 'buku_pinjaman.xlsx', sheet: S, base, alt, answerCell: 'B1',
          statement: '<p>Citra at Bank Kenari wants the number of facilities with any days past due. In <b>B1</b>, count the rows whose <b>DPD</b> (' + R + ') is <b>not 0</b>.</p>',
          solution: '=COUNTIF(' + R + ',"<>0")', alts: ['=COUNTIF(' + R + ',">0")', '=COUNT(' + R + ')-COUNTIF(' + R + ',0)', '=COUNTIFS(' + R + ',"<>0")'],
          hints: ['In a condition, not equal is written as <>.', 'The condition is the text "<>0", in quotes like any other condition.', 'Shape: =COUNTIF(' + R + ',"<>0").'],
          explain: '"<>0" keeps every cell that is not zero, so COUNTIF counts the facilities with some days past due. Since DPD is never negative, ">0" gives the same count here.',
          chips: [R, '"<>0"'], fnKey: 'COUNTIF('
        });
      }
      const V = [
        { title: 'Net sales of every distributor except the one in B1', file: 'penjualan_q4.xlsx', label: 'Leave out', h: 'Distributor', k: 'distributor', val: 'Net (Rp)', fn: 'SUMIF', ans: 'Net of the rest (Rp)',
          rows: () => { const g = names(r, 4); return lines(r, 14, (s) => g.includes(s.distributor)); }, cols: [COL.date, COL.dist, COL.product, COL.net] },
        { title: 'Count facilities outside the region in B1', file: 'buku_pinjaman.xlsx', label: 'Leave out', h: 'Region', k: 'region', fn: 'COUNTIF', ans: 'Facilities elsewhere',
          rows: () => H.loans(r, 16), cols: [COL.facility, COL.dist, COL.region, COL.outstanding] },
        { title: 'Average cartons of lines outside the category in B1', file: 'penjualan_q4.xlsx', label: 'Leave out', h: 'Category', k: 'category', val: 'Cartons', fn: 'AVERAGEIF', ans: 'Average cartons',
          rows: () => lines(r, 14), cols: [COL.date, COL.product, COL.category, COL.cartons] }
      ][v];
      const rows = V.rows(), S = H.table(V.cols, rows), keys = ranked(rows, V.k), c = S.R(V.h);
      const base = Object.assign({}, S.cells, { A1: V.label, B1: keys[0], A2: V.ans });
      const edits = { B1: keys[1] };
      /* For the count, one row of B1's new group moves to another group, so the hidden answer always differs. */
      if (!V.val) edits[S.cell(V.h, rows.findIndex((x) => x[V.k] === keys[1]))] = keys[2] || keys[0];
      if (V.val) { const j = rows.findIndex((x) => x[V.k] === keys[2] || (keys.length < 3 && x[V.k] === keys[0])); edits[S.cell(V.val, j)] = (x) => x + step(V.val); }
      const alt = H.edit(base, edits);
      const s = V.val ? S.R(V.val) : null;
      const sol = V.fn === 'COUNTIF' ? '=COUNTIF(' + c + ',"<>"&B1)' : '=' + V.fn + '(' + c + ',"<>"&B1,' + s + ')';
      const alts = V.fn === 'COUNTIF' ? ['=ROWS(' + c + ')-COUNTIF(' + c + ',B1)', '=COUNTA(' + c + ')-COUNTIF(' + c + ',B1)', '=COUNTIFS(' + c + ',"<>"&B1)']
        : V.fn === 'SUMIF' ? ['=SUM(' + s + ')-SUMIF(' + c + ',B1,' + s + ')', '=SUMIFS(' + s + ',' + c + ',"<>"&B1)']
          : ['=AVERAGEIFS(' + s + ',' + c + ',"<>"&B1)', '=(SUM(' + s + ')-SUMIF(' + c + ',B1,' + s + '))/(COUNT(' + s + ')-COUNTIF(' + c + ',B1))'];
      const job = V.fn === 'COUNTIF' ? 'count the rows whose <b>' + V.h + '</b> (' + c + ') is anything but B1' : V.fn === 'SUMIF' ? 'add up <b>' + V.val + '</b> (' + s + ') for every row whose <b>' + V.h + '</b> (' + c + ') is not B1' : 'average <b>' + V.val + '</b> (' + s + ') over the rows whose <b>' + V.h + '</b> (' + c + ') is not B1, keeping the decimals';
      return H.problem({
        title: V.title, file: V.file, sheet: S, base, alt, answerCell: 'B2',
        statement: '<p>In <b>B2</b>, ' + job + '.</p>',
        solution: sol, alts,
        hints: ['Not equal is <> in a condition.', 'Join it to the cell: "<>"&B1 means anything except what is in B1.', 'Shape: ' + sol + '.'],
        explain: '"<>"&B1 builds a condition such as "<>' + keys[0] + '", which every other row passes. ' + (V.fn === 'COUNTIF' ? 'Counting every row and taking away the B1 rows gives the same answer.' : V.fn === 'SUMIF' ? 'Taking the whole total and subtracting the B1 part gives the same answer.' : 'AVERAGEIF leaves the B1 rows out of both the total and the count it divides by.'),
        chips: V.val ? [c, s, '"<>"&B1'] : [c, '"<>"&B1'], fnKey: V.fn + '('
      });
    }
  });

  P.add({
    key: 'names-that-start-with', difficulty: 'Medium', level: 4, topics: ['SUMIF', 'COUNTIF', 'wildcards'], count: 4,
    make(r, v) {
      if (v === 3) {
        /* The prefix comes from a cell: B1&"*". */
        const rows = r.shuffle(H.stock(r, 14)), S = H.table([COL.sku, COL.product, COL.warehouse, COL.onhand], rows);
        const fams = ranked(rows, (x) => family(x.sku)), a = S.R('SKU'), s = S.R('On hand');
        const base = Object.assign({}, S.cells, { A1: 'SKU family', B1: fams[0], A2: 'On hand (cartons)' });
        const j = rows.findIndex((x) => family(x.sku) === fams[1]);
        const alt = H.edit(base, { B1: fams[1], [S.cell('On hand', j)]: (x) => x + 30 });
        return H.problem({
          title: 'Cartons on hand for the SKU family code in B1', file: 'stok_gudang.xlsx', sheet: S, base, alt, answerCell: 'B2',
          statement: '<p>Every SKU starts with a family code: MLK, KSJ, SNY or SRW. In <b>B2</b>, add up <b>On hand</b> (' + s + ') for the SKUs (' + a + ') that start with the code in <b>B1</b>.</p>',
          solution: '=SUMIF(' + a + ',B1&"*",' + s + ')', alts: ['=SUMIF(' + a + ',B1&"-*",' + s + ')', '=SUMIFS(' + s + ',' + a + ',B1&"*")'],
          hints: ['A * in a condition stands for any run of characters, so "KSJ*" means starts with KSJ.', 'The code is in B1, so join the star to the cell: B1&"*".', 'Shape: =SUMIF(' + a + ',B1&"*",' + s + ').'],
          explain: 'B1&"*" turns ' + fams[0] + ' into the pattern "' + fams[0] + '*", which matches every SKU of that family. SUMIF then adds their on-hand cartons.',
          chips: [a, s, 'B1&"*"'], fnKey: 'SUMIF('
        });
      }
      const V = [
        { title: 'Net sales from CV distributors only', pat: 'CV*', fn: 'SUMIF', h: 'Distributor', val: 'Net (Rp)', who: 'starts with <b>CV</b>',
          rows: () => lines(r, 14), cols: [COL.date, COL.dist, COL.product, COL.net], test: (x) => x.distributor.startsWith('CV') },
        { title: 'Count the sales lines of UD distributors', pat: 'UD*', fn: 'COUNTIF', h: 'Distributor', who: 'starts with <b>UD</b>',
          rows: () => lines(r, 14), cols: [COL.date, COL.dist, COL.region, COL.net], test: (x) => x.distributor.startsWith('UD') },
        { title: 'Cartons of every Mi Lekas noodle product', pat: 'Mi Lekas*', fn: 'SUMIF', h: 'Product', val: 'Cartons', who: 'starts with <b>Mi Lekas</b>',
          rows: () => lines(r, 5, (s) => s.category === 'Noodles').concat(lines(r, 9, (s) => s.category !== 'Noodles')).sort(byDate), cols: [COL.date, COL.dist, COL.product, COL.cartons], test: (x) => x.product.startsWith('Mi Lekas') }
      ][v];
      const rows = V.rows(), S = H.table(V.cols, rows), c = S.R(V.h), s = V.val ? S.R(V.val) : null;
      const base = Object.assign({}, S.cells, { A1: V.fn === 'COUNTIF' ? 'Lines' : 'Total' });
      const hit = where(rows, V.test), miss = where(rows, (x) => !V.test(x));
      /* The hidden sheet: for a count, one more row belongs to the group; for a total, two matching figures are corrected. */
      const alt = V.fn === 'COUNTIF' ? H.edit(base, { [S.cell(V.h, miss[0])]: rows[hit[0]].distributor })
        : H.edit(base, { [S.cell(V.val, hit[0])]: (x) => x + step(V.val), [S.cell(V.val, hit[hit.length - 1])]: (x) => x + step(V.val) });
      const sol = V.fn === 'COUNTIF' ? '=COUNTIF(' + c + ',"' + V.pat + '")' : '=SUMIF(' + c + ',"' + V.pat + '",' + s + ')';
      return H.problem({
        title: V.title, file: 'penjualan_q4.xlsx', sheet: S, base, alt, answerCell: 'B1',
        statement: '<p>' + (V.h === 'Distributor' ? 'A distributor name begins with its legal form, CV or UD. ' : 'The Mi Lekas noodles come in several flavours. ') + 'In <b>B1</b>, ' + (V.fn === 'COUNTIF' ? 'count the rows' : 'add up <b>' + V.val + '</b> (' + s + ') for the rows') + ' whose <b>' + V.h + '</b> (' + c + ') ' + V.who + '.</p>',
        solution: sol, alts: V.fn === 'COUNTIF' ? ['=COUNTIFS(' + c + ',"' + V.pat + '")'] : ['=SUMIFS(' + s + ',' + c + ',"' + V.pat + '")'],
        hints: ['A * in a condition matches any run of characters.', 'Put the star at the end to mean starts with: "' + V.pat + '".', 'Shape: ' + sol + '.'],
        explain: '"' + V.pat + '" matches every name that begins with ' + V.pat.slice(0, -1) + ', whatever follows. The star only works inside the condition of functions like SUMIF and COUNTIF.',
        chips: V.val ? [c, s, '"' + V.pat + '"'] : [c, '"' + V.pat + '"'], fnKey: V.fn + '('
      });
    }
  });

  P.add({
    key: 'text-that-contains', difficulty: 'Medium', level: 5, topics: ['SUMIF', 'COUNTIF', 'SUMIFS', 'wildcards'], count: 4,
    make(r, v) {
      if (v === 0) {
        const rows = lines(r, 14), S = H.table([COL.date, COL.dist, COL.product, COL.cartons], rows), words = ranked(rows, (x) => word(x.product));
        const p = S.R('Product'), c = S.R('Cartons');
        const base = Object.assign({}, S.cells, { A1: 'Brand word', B1: words[0], A2: 'Cartons' });
        const j = rows.findIndex((x) => word(x.product) === words[1]);
        const alt = H.edit(base, { B1: words[1], [S.cell('Cartons', j)]: (x) => x + 6 });
        return H.problem({
          title: 'Cartons of products whose name contains the word in B1', file: 'penjualan_q4.xlsx', sheet: S, base, alt, answerCell: 'B2',
          statement: '<p>B1 holds a brand word such as Sekejap or Nyala. It sits inside the product name, not at the start. In <b>B2</b>, add up <b>Cartons</b> (' + c + ') for every product (' + p + ') whose name contains that word.</p>',
          solution: '=SUMIF(' + p + ',"*"&B1&"*",' + c + ')', alts: ['=SUMIFS(' + c + ',' + p + ',"*"&B1&"*")'],
          hints: ['A star on both sides means contains: "*Sekejap*".', 'The word is in B1, so build the pattern with &: "*"&B1&"*".', 'Shape: =SUMIF(' + p + ',"*"&B1&"*",' + c + ').'],
          explain: '"*"&B1&"*" becomes "*' + words[0] + '*", which matches the word anywhere in the name. A pattern without the leading star would only match names that begin with it.',
          chips: [p, c, '"*"&B1&"*"'], fnKey: 'SUMIF('
        });
      }
      if (v === 1) {
        const NOTES = ['Promo bundle', 'Year-end promo', 'Repeat order', 'New outlet', 'PROMO display', 'Price check', 'Late delivery', 'Repeat order', 'Weekend Promo', 'Short shipped', 'New outlet', 'Paid in advance', 'Repeat order', 'Price check'];
        const notes = r.shuffle(NOTES);
        const rows = lines(r, 14).map((x, i) => Object.assign({}, x, { note: notes[i] }));
        const S = H.table([COL.date, COL.dist, COL.net, { h: 'Remark', k: 'note' }], rows), R = S.R('Remark');
        const base = Object.assign({}, S.cells, { A1: 'Promo lines' });
        const plain = where(rows, (x) => !/promo/i.test(x.note));
        const alt = H.edit(base, { [S.cell('Remark', plain[0])]: 'Promo end of month', [S.cell('Remark', plain[plain.length - 1])]: 'Extra promo stock' });
        return H.problem({
          title: 'Count the lines whose remark mentions promo', file: 'penjualan_q4.xlsx', sheet: S, base, alt, answerCell: 'B1',
          statement: '<p>Pak Anton wants to know how many lines were part of a promotion. In <b>B1</b>, count the rows whose <b>Remark</b> (' + R + ') has the word <b>promo</b> anywhere in it, in any capitals.</p>',
          solution: '=COUNTIF(' + R + ',"*promo*")', alts: ['=COUNTIFS(' + R + ',"*promo*")', '=COUNTIF(' + R + ',"*PROMO*")'],
          hints: ['Remarks put promo at the start, the end or in the middle, so you need a star on both sides.', 'COUNTIF ignores capitals, so "*promo*" also finds PROMO and Promo.', 'Shape: =COUNTIF(' + R + ',"*promo*").'],
          explain: '"*promo*" matches any remark with promo somewhere inside, and COUNTIF does not care about capitals. "promo*" alone would miss Year-end promo.',
          chips: [R, '"*promo*"'], fnKey: 'COUNTIF('
        });
      }
      if (v === 2) {
        const rows = bills(r, 14, names(r, 4)), S = H.table([COL.ref, COL.date, COL.amount], rows), codes = ranked(rows, 'code');
        const a = S.R('Ref'), s = S.R('Amount (Rp)');
        const base = Object.assign({}, S.cells, { A1: 'City code', B1: codes[0], A2: 'Amount (Rp)' });
        const j = rows.findIndex((x) => x.code === codes[1]);
        const alt = H.edit(base, { B1: codes[1], [S.cell('Amount (Rp)', j)]: (x) => x + 2500000 });
        return H.problem({
          title: 'Invoice total for the city code at the end of the ref', file: 'tagihan_q4.xlsx', sheet: S, base, alt, answerCell: 'B2',
          statement: '<p>Each invoice ref ends with a dash and the city code of the distributor: INV-2107-BAN is Bandung. In <b>B2</b>, add up <b>Amount</b> (' + s + ') for the refs (' + a + ') that end with the code in <b>B1</b>.</p>',
          solution: '=SUMIF(' + a + ',"*-"&B1,' + s + ')', alts: ['=SUMIF(' + a + ',"*"&B1,' + s + ')', '=SUMIFS(' + s + ',' + a + ',"*-"&B1)'],
          hints: ['Ends with means the star goes in front: "*-BAN".', 'The code is in B1: "*-"&B1.', 'Shape: =SUMIF(' + a + ',"*-"&B1,' + s + ').'],
          explain: '"*-"&B1 becomes "*-' + codes[0] + '", which matches refs that end with that code whatever number comes before it. SUMIF adds the amounts of those refs.',
          chips: [a, s, '"*-"&B1'], fnKey: 'SUMIF('
        });
      }
      const g = r.sample(H.REGIONS, 3);
      const rows = lines(r, 16, (s) => g.includes(s.region)), S = H.table([COL.date, COL.region, COL.product, COL.net], rows);
      const combos = ranked(rows, (x) => word(x.product) + '|' + x.region).map((k) => k.split('|'));
      const p = S.R('Product'), q = S.R('Region'), s = S.R('Net (Rp)');
      const base = Object.assign({}, S.cells, { A1: 'Brand word', B1: combos[0][0], A2: 'Region', B2: combos[0][1], A3: 'Net (Rp)' });
      const j = rows.findIndex((x) => word(x.product) === combos[1][0] && x.region === combos[1][1]);
      const alt = H.edit(base, { B1: combos[1][0], B2: combos[1][1], [S.cell('Net (Rp)', j)]: (x) => x + 1800000 });
      return H.problem({
        title: 'Net for a brand word in a region, both from cells', file: 'penjualan_q4.xlsx', sheet: S, base, alt, answerCell: 'B3',
        statement: '<p>B1 holds a brand word that sits inside product names (such as Sekejap) and B2 a region. In <b>B3</b>, add up <b>Net</b> (' + s + ') for the rows whose product (' + p + ') contains the word in B1 and whose region (' + q + ') is B2.</p>',
        solution: '=SUMIFS(' + s + ',' + p + ',"*"&B1&"*",' + q + ',B2)', alts: ['=SUMIFS(' + s + ',' + q + ',B2,' + p + ',"*"&B1&"*")'],
        hints: ['Two conditions call for SUMIFS: the range to add comes first, then pairs of range and condition.', 'Contains is "*"&B1&"*"; the region condition is B2 as it is.', 'Shape: =SUMIFS(' + s + ',' + p + ',"*"&B1&"*",' + q + ',B2).'],
        explain: 'SUMIFS adds a row only when both pairs pass: the product contains the word, and the region equals B2. Order of the pairs does not matter.',
        chips: [s, p, q, '"*"&B1&"*"'], fnKey: 'SUMIFS('
      });
    }
  });

  P.add({
    key: 'total-on-two-conditions', difficulty: 'Medium', level: 5, topics: ['SUMIFS'], count: 4,
    make(r, v) {
      const V = [
        { title: 'Net for one distributor in one category', file: 'penjualan_q4.xlsx', a: 'Distributor', b: 'Category', sum: 'Net (Rp)',
          rows: () => { const g = names(r, 3); return lines(r, 16, (s) => g.includes(s.distributor)); }, cols: [COL.date, COL.dist, COL.category, COL.net] },
        { title: 'Cartons of one product in one region', file: 'penjualan_q4.xlsx', a: 'Region', b: 'Product', sum: 'Cartons',
          rows: () => { const g = r.sample(H.REGIONS, 3), ps = r.sample(H.PRODUCTS, 4).map((x) => x.name); return lines(r, 16, (s) => g.includes(s.region) && ps.includes(s.product)); }, cols: [COL.date, COL.region, COL.product, COL.cartons] },
        { title: 'Outstanding of one loan type in one region', file: 'buku_pinjaman.xlsx', a: 'Region', b: 'Type', sum: 'Outstanding (Rp)',
          rows: () => H.loans(r, 18), cols: [COL.facility, COL.region, COL.type, COL.outstanding] },
        { title: 'On-hand cartons of one category in one warehouse', file: 'stok_gudang.xlsx', a: 'Warehouse', b: 'Category', sum: 'On hand',
          rows: () => H.stock(r, 16), cols: [COL.product, COL.category, COL.warehouse, COL.onhand] }
      ][v];
      const rows = V.rows(), S = H.table(V.cols, rows);
      const ka = V.cols.find((c) => c.h === V.a).k, kb = V.cols.find((c) => c.h === V.b).k;
      const combos = ranked(rows, (x) => x[ka] + '|' + x[kb]).map((k) => k.split('|'));
      const A = S.R(V.a), B = S.R(V.b), s = S.R(V.sum);
      const base = Object.assign({}, S.cells, { A1: V.a, B1: combos[0][0], A2: V.b, B2: combos[0][1], A3: 'Total' });
      /* The hidden sheet: a new pair in B1 and B2, both parts changed, and one of its figures is corrected. */
      const sk = V.cols.find((c) => c.h === V.sum).k;
      const hp = hiddenPair(rows, ka, kb, combos, sk, step(V.sum), total), j = hp.j;
      const alt = H.edit(base, { B1: hp.pair[0], B2: hp.pair[1], [S.cell(V.sum, j)]: (x) => x + step(V.sum) });
      const sol = '=SUMIFS(' + s + ',' + A + ',B1,' + B + ',B2)';
      return H.problem({
        title: V.title, file: V.file, sheet: S, base, alt, answerCell: 'B3',
        statement: '<p>In <b>B3</b>, add up <b>' + V.sum + '</b> (' + s + ') for the rows where <b>' + V.a + '</b> (' + A + ') is B1 and <b>' + V.b + '</b> (' + B + ') is B2.</p>',
        solution: sol, alts: ['=SUMIFS(' + s + ',' + B + ',B2,' + A + ',B1)', '=SUMPRODUCT((' + A + '=B1)*(' + B + '=B2)*' + s + ')'],
        hints: ['SUMIFS takes the range to add first, then a range and a condition for each test.', 'There are two tests: ' + A + ' against B1 and ' + B + ' against B2.', 'Shape: ' + sol + '.'],
        explain: 'SUMIFS adds a row only when every pair passes, so only the rows with both ' + V.a + ' B1 and ' + V.b + ' B2 count. Unlike SUMIF, the range to add comes first.',
        chips: [s, A, B], fnKey: 'SUMIFS('
      });
    }
  });

  P.add({
    key: 'count-on-a-group-and-a-limit', difficulty: 'Medium', level: 5, topics: ['COUNTIFS'], count: 4,
    make(r, v) {
      const V = [
        { title: 'Count orders in one region of at least a carton size', file: 'penjualan_q4.xlsx', a: 'Region', n: 'Cartons', op: '>=', words: 'at least',
          rows: () => { const g = r.sample(H.REGIONS, 3); return lines(r, 16, (s) => g.includes(s.region)); }, cols: [COL.date, COL.dist, COL.region, COL.cartons] },
        { title: 'Count facilities of one type past a DPD limit', file: 'buku_pinjaman.xlsx', a: 'Type', n: 'DPD', op: '>', words: 'more than', fixed: [0, 30],
          rows: () => H.loans(r, 18), cols: [COL.facility, COL.dist, COL.type, COL.dpd] },
        { title: 'Count stock lines in one warehouse below a level', file: 'stok_gudang.xlsx', a: 'Warehouse', n: 'On hand', op: '<', words: 'less than', fixed: [200, 320],
          rows: () => H.stock(r, 16), cols: [COL.sku, COL.product, COL.warehouse, COL.onhand] },
        { title: 'Count reps in one area with at least some visits', file: 'tim_sales.xlsx', a: 'Area', n: 'Visits', op: '>=', words: 'at least',
          rows: () => repsIn(r, 12), cols: [COL.rep, COL.area, COL.visits] }
      ][v];
      const rows = V.rows();
      const ka = V.cols.find((c) => c.h === V.a).k, kn = V.cols.find((c) => c.h === V.n).k, keys = ranked(rows, ka);
      const mid = (key) => { const vals = sorted(rows.filter((x) => x[ka] === key), kn); return vals[Math.floor(vals.length / 2)]; };
      const lim = (V.fixed || [mid(keys[0]), mid(keys[1])]).slice();
      if (lim[1] === lim[0]) lim[1] = lim[0] + 3;
      /* One row of B1's group sits exactly on the visible limit, without changing the count, so the wrong sign
         (> for >=, or <= for <) gives a different answer. */
      if (!rows.some((x) => x[ka] === keys[0] && x[kn] === lim[0])) {
        const i = rows.findIndex((x) => x[ka] === keys[0] && (V.op === '>' ? x[kn] < lim[0] : x[kn] > lim[0]));
        if (i >= 0) rows[i] = Object.assign({}, rows[i], { [kn]: lim[0] });
      }
      const S = H.table(V.cols, rows);
      const A = S.R(V.a), N = S.R(V.n);
      const base = Object.assign({}, S.cells, { A1: V.a, B1: keys[0], A2: V.n + ' limit', B2: lim[0], A3: 'Count' });
      const pass = (x, key, l) => x[ka] === key && (V.op === '>=' ? x[kn] >= l : V.op === '>' ? x[kn] > l : x[kn] < l);
      /*
       * The hidden sheet asks for B1's second group with a new limit in B2. Its rows are adjusted so the hidden count
       * differs from the visible one, and from the counts a formula with B1 or B2 typed in would give.
       */
      const edits = { B1: keys[1], B2: lim[1] }, alt0 = rows.slice();
      const put = (i, h, k, val) => { alt0[i] = Object.assign({}, alt0[i], { [k]: val }); edits[S.cell(h, i)] = val; };
      const cnt = (rs, key, l) => rs.filter((x) => pass(x, key, l)).length;
      const lo = Math.min(lim[0], lim[1]), hi = Math.max(lim[0], lim[1]);
      const own = where(rows, (x) => x[ka] === keys[1]);
      /* One row of the new group passes one limit and fails the other, so a typed B2 gives a different count. */
      let split = own.find((i) => pass(rows[i], keys[1], lim[0]) !== pass(rows[i], keys[1], lim[1]));
      if (split === undefined) { split = own[0]; put(split, V.n, kn, V.op === '>' ? hi : lo); }
      for (let guard = 0; guard < 4; guard++) {
        const n = cnt(alt0, keys[1], lim[1]);
        if (n !== cnt(rows, keys[0], lim[0]) && n !== cnt(alt0, keys[0], lim[1])) break;
        /* Otherwise one more row of the new group passes both limits. */
        const i = own.find((x) => x !== split && !pass(alt0[x], keys[1], lim[0]) && !pass(alt0[x], keys[1], lim[1]));
        if (i !== undefined) { put(i, V.n, kn, V.op === '<' ? Math.max(0, lo - 20) : hi + 5); continue; }
        const o = where(alt0, (x) => x[ka] !== keys[0] && x[ka] !== keys[1])[0];
        if (o === undefined) break;
        put(o, V.a, ka, keys[1]); put(o, V.n, kn, V.op === '<' ? Math.max(0, lo - 20) : hi + 5);
      }
      const alt = H.edit(base, edits);
      const sol = '=COUNTIFS(' + A + ',B1,' + N + ',"' + V.op + '"&B2)';
      return H.problem({
        title: V.title, file: V.file, sheet: S, base, alt, answerCell: 'B3',
        statement: '<p>In <b>B3</b>, count the rows where <b>' + V.a + '</b> (' + A + ') is B1 and <b>' + V.n + '</b> (' + N + ') is <b>' + V.words + '</b> the number in B2.</p>',
        solution: sol, alts: ['=COUNTIFS(' + N + ',"' + V.op + '"&B2,' + A + ',B1)', '=SUMPRODUCT((' + A + '=B1)*(' + N + V.op + 'B2))'],
        hints: ['COUNTIFS takes pairs of range and condition and counts the rows that pass all of them.', 'The first pair is ' + A + ' with B1. The second is ' + N + ' with "' + V.op + '"&B2.', 'Shape: ' + sol + '.'],
        explain: 'COUNTIFS counts a row only when both conditions hold. "' + V.op + '"&B2 joins the sign to the limit, so a row equal to B2 is ' + (V.op.includes('=') ? 'counted' : 'left out') + '.',
        chips: [A, N, '"' + V.op + '"&B2'], fnKey: 'COUNTIFS('
      });
    }
  });

  P.add({
    key: 'largest-or-smallest-in-a-group', difficulty: 'Medium', level: 5, topics: ['MAXIFS', 'MINIFS'], count: 4,
    make(r, v) {
      if (v === 3) {
        /* Earliest date a product shipped in a big order: MINIFS on dates with two conditions. */
        const ps = r.sample(H.PRODUCTS, 4).map((x) => x.name);
        const rows = lines(r, 16, (s) => ps.includes(s.product)), S = H.table([COL.date, COL.dist, COL.product, COL.cartons], rows);
        const pick = [];
        for (const name of ranked(rows, 'product')) {
          const own = rows.filter((x) => x.product === name);
          if (own.length < 2) continue;
          const later = Math.max.apply(null, own.slice(1).map((x) => x.cartons));
          if (later > own[0].cartons) pick.push({ name, lim: later });
          if (pick.length === 2) break;
        }
        while (pick.length < 2) { const name = ranked(rows, 'product')[pick.length]; pick.push({ name, lim: rows.find((x) => x.product === name).cartons }); }
        const p = S.R('Product'), c = S.R('Cartons'), d = S.R('Date');
        const base = Object.assign({}, S.cells, { A1: 'Product', B1: pick[0].name, A2: 'Cartons at least', B2: pick[0].lim, A3: 'First date' });
        const alt = H.edit(base, { B1: pick[1].name, B2: pick[1].lim });
        return H.problem({
          title: 'First date the product in B1 shipped in a big order', file: 'penjualan_q4.xlsx', sheet: S, base, alt, answerCell: 'B3', ansDate: true,
          statement: '<p>In <b>B3</b>, find the earliest <b>Date</b> (' + d + ') on which the product in B1 shipped in an order of <b>at least</b> the cartons in B2. The answer is a date.</p>',
          solution: '=MINIFS(' + d + ',' + p + ',B1,' + c + ',">="&B2)', alts: ['=MINIFS(' + d + ',' + c + ',">="&B2,' + p + ',B1)'],
          hints: ['Dates are numbers in Excel, so the earliest date is the smallest one.', 'MINIFS takes the range to search first, then pairs of range and condition, like SUMIFS.', 'Shape: =MINIFS(' + d + ',' + p + ',B1,' + c + ',">="&B2).'],
          explain: 'MINIFS keeps the rows for the product with at least B2 cartons and returns the smallest date among them, which is the earliest. Smaller orders of the same product are skipped.',
          chips: [d, p, c, '">="&B2'], fnKey: 'MINIFS('
        });
      }
      const V = [
        { title: 'Largest single order of the distributor in B1', file: 'penjualan_q4.xlsx', fn: 'MAXIFS', label: 'Distributor', h: 'Distributor', k: 'distributor', val: 'Net (Rp)', what: 'the largest <b>Net</b>',
          rows: () => { const g = names(r, 4); return lines(r, 14, (s) => g.includes(s.distributor)); }, cols: [COL.date, COL.dist, COL.product, COL.net] },
        { title: 'Smallest outstanding balance in the region in B1', file: 'buku_pinjaman.xlsx', fn: 'MINIFS', label: 'Region', h: 'Region', k: 'region', val: 'Outstanding (Rp)', what: 'the smallest <b>Outstanding</b>',
          rows: () => H.loans(r, 18), cols: [COL.facility, COL.dist, COL.region, COL.outstanding] },
        { title: 'Latest order date of the distributor in B1', file: 'penjualan_q4.xlsx', fn: 'MAXIFS', label: 'Distributor', h: 'Distributor', k: 'distributor', val: 'Date', what: 'the latest <b>Date</b>', date: true,
          rows: () => { const g = names(r, 4); return lines(r, 14, (s) => g.includes(s.distributor)); }, cols: [COL.date, COL.dist, COL.product, COL.cartons] }
      ][v];
      const rows = V.rows(), S = H.table(V.cols, rows), keys = ranked(rows, V.k);
      const c = S.R(V.h), s = S.R(V.val);
      const base = Object.assign({}, S.cells, { A1: V.label, B1: keys[0], A2: V.date ? 'Latest date' : 'Answer (Rp)' });
      /* The hidden sheet: B1 names another group, and that group's extreme figure is corrected outward. */
      const vk = V.cols.find((c) => c.h === V.val).k, mine = where(rows, (x) => x[V.k] === keys[1]);
      const ext = mine.reduce((b, i) => ((V.fn === 'MAXIFS' ? rows[i][vk] > rows[b][vk] : rows[i][vk] < rows[b][vk]) ? i : b), mine[0]);
      const alt = H.edit(base, { B1: keys[1], [S.cell(V.val, ext)]: (x) => x + (V.date ? 1 : V.fn === 'MAXIFS' ? step(V.val) : -step(V.val)) });
      const sol = '=' + V.fn + '(' + s + ',' + c + ',B1)';
      return H.problem({
        title: V.title, file: V.file, sheet: S, base, alt, answerCell: 'B2', ansDate: !!V.date,
        statement: '<p>In <b>B2</b>, find ' + V.what + ' (' + s + ') among the rows whose <b>' + V.h + '</b> (' + c + ') is B1.' + (V.date ? ' The answer is a date.' : '') + '</p>',
        solution: sol, alts: ['=' + (V.fn === 'MAXIFS' ? 'MAX' : 'MIN') + '(FILTER(' + s + ',' + c + '=B1))', '=' + V.fn + '(' + S.A(V.val) + ',' + S.A(V.h) + ',$B$1)'],
        hints: [V.fn + ' returns the ' + (V.fn === 'MAXIFS' ? 'largest' : 'smallest') + ' value among the rows that meet the conditions.' + (V.date ? ' Dates are numbers, so the latest date is the largest.' : ''), 'The range to search comes first, then the range to test and the condition: ' + c + ' and B1.', 'Shape: ' + sol + '.'],
        explain: V.fn + ' looks only at the rows where ' + c + ' equals B1 and returns the ' + (V.fn === 'MAXIFS' ? 'largest' : 'smallest') + ' value of ' + s + '. ' + (V.fn === 'MAXIFS' ? 'MAX' : 'MIN') + '(FILTER(...)) gives the same answer in Excel 365.',
        chips: [s, c, 'B1'], fnKey: V.fn + '('
      });
    }
  });

  P.add({
    key: 'totals-between-two-dates', difficulty: 'Medium', level: 6, topics: ['SUMIFS', 'COUNTIFS', 'dates'], count: 4,
    make(r, v) {
      if (v === 2) {
        /* One whole month from its first day: EOMONTH closes the window. */
        const rows = bills(r, 16, names(r, 5)), S = H.table([COL.ref, COL.date, COL.amount], rows);
        const d = S.R('Date'), s = S.R('Amount (Rp)');
        const mons = ranked(rows, (x) => x.date.slice(0, 7));
        const base = Object.assign({}, S.cells, { A1: 'Month (first day)', B1: H.toSerial(mons[0] + '-01'), A2: 'Invoiced (Rp)' });
        /* The hidden sheet asks for another month, whose first and last invoices are dated on its first and last day,
           so a window that leaves out either end gives a different total. */
        const edits = { B1: H.toSerial(mons[1] + '-01') }, inAlt = where(rows, (x) => x.date.slice(0, 7) === mons[1]);
        const y = +mons[1].slice(0, 4), mo = +mons[1].slice(5, 7);
        const nextFirst = H.toSerial(mo === 12 ? (y + 1) + '-01-01' : y + '-' + String(mo + 1).padStart(2, '0') + '-01');
        if (inAlt.length) edits[S.cell('Date', inAlt[0])] = H.toSerial(mons[1] + '-01');
        if (inAlt.length > 1) edits[S.cell('Date', inAlt[inAlt.length - 1])] = nextFirst - 1;
        const alt = H.edit(base, edits);
        return H.problem({
          title: 'Invoice total for the month that starts on the date in B1', file: 'tagihan_q4.xlsx', sheet: S, base, alt, answerCell: 'B2', dateCells: ['B1'],
          statement: '<p>B1 holds the first day of a month. In <b>B2</b>, add up <b>Amount</b> (' + s + ') for the invoices dated (' + d + ') anywhere in that month, from its first day to its last.</p>',
          solution: '=SUMIFS(' + s + ',' + d + ',">="&B1,' + d + ',"<="&EOMONTH(B1,0))',
          alts: ['=SUMIFS(' + s + ',' + d + ',">="&B1,' + d + ',"<"&EDATE(B1,1))', '=SUMPRODUCT((' + d + '>=B1)*(' + d + '<=EOMONTH(B1,0))*' + s + ')'],
          hints: ['A month is a window: on or after its first day, and on or before its last day.', 'EOMONTH(B1,0) returns the last day of the month B1 is in.', 'Shape: =SUMIFS(' + s + ',' + d + ',">="&B1,' + d + ',"<="&EOMONTH(B1,0)).'],
          explain: 'The two conditions on the same date column make a window from B1 to EOMONTH(B1,0). Using "<"&EDATE(B1,1), before the first day of the next month, gives the same window.',
          chips: [s, d, '">="&B1', '"<="&EOMONTH(B1,0)'], fnKey: 'SUMIFS('
        });
      }
      const V = [
        { title: 'Net sales between the two dates in B1 and B2', file: 'penjualan_q4.xlsx', fn: 'SUMIFS', dh: 'Date', val: 'Net (Rp)', ans: 'Net (Rp)', what: 'add up <b>Net</b>',
          rows: () => lines(r, 14), cols: [COL.date, COL.dist, COL.product, COL.net], at: [3, 10, 1, 8] },
        { title: 'Count the sales lines dated from B1 to B2', file: 'penjualan_q4.xlsx', fn: 'COUNTIFS', dh: 'Date', ans: 'Lines', what: 'count the lines',
          rows: () => lines(r, 16), cols: [COL.date, COL.dist, COL.region, COL.cartons], at: [2, 9, 5, 13] },
        null,
        { title: 'Limits of facilities that started between B1 and B2', file: 'buku_pinjaman.xlsx', fn: 'SUMIFS', dh: 'Start', val: 'Limit (Rp)', ans: 'Limits (Rp)', what: 'add up <b>Limit</b>',
          rows: () => H.loans(r, 16), cols: [COL.facility, COL.dist, COL.start, COL.limit], at: [3, 11, 1, 8] }
      ][v];
      const rows = V.rows(), S = H.table(V.cols, rows);
      const dk = V.cols.find((c) => c.h === V.dh).k, ds = sorted(rows, dk), d = S.R(V.dh);
      const base = Object.assign({}, S.cells, { A1: 'From', B1: H.toSerial(ds[V.at[0]]), A2: 'To', B2: H.toSerial(ds[V.at[1]]), A3: V.ans });
      const alt = H.edit(base, { B1: H.toSerial(ds[V.at[2]]), B2: H.toSerial(ds[V.at[3]]) });
      const s = V.val ? S.R(V.val) : null;
      const sol = V.fn === 'SUMIFS' ? '=SUMIFS(' + s + ',' + d + ',">="&B1,' + d + ',"<="&B2)' : '=COUNTIFS(' + d + ',">="&B1,' + d + ',"<="&B2)';
      const alts = V.fn === 'SUMIFS' ? ['=SUMIFS(' + s + ',' + d + ',"<="&B2,' + d + ',">="&B1)', '=SUMPRODUCT((' + d + '>=B1)*(' + d + '<=B2)*' + s + ')', '=SUMIF(' + d + ',"<="&B2,' + s + ')-SUMIF(' + d + ',"<"&B1,' + s + ')']
        : ['=COUNTIF(' + d + ',"<="&B2)-COUNTIF(' + d + ',"<"&B1)', '=SUMPRODUCT((' + d + '>=B1)*(' + d + '<=B2))'];
      return H.problem({
        title: V.title, file: V.file, sheet: S, base, alt, answerCell: 'B3', dateCells: ['B1', 'B2'],
        statement: '<p>In <b>B3</b>, ' + V.what + (s ? ' (' + s + ') for the rows' : '') + ' whose <b>' + V.dh + '</b> (' + d + ') falls from the date in B1 to the date in B2. Both end dates count.</p>',
        solution: sol, alts,
        hints: ['A date window is two conditions on the same column: on or after the start, and on or before the end.', 'Join the signs to the cells: ">="&B1 and "<="&B2.', 'Shape: ' + sol + '.'],
        explain: 'Dates are numbers, so ">="&B1 and "<="&B2 keep the rows inside the window, ends included. Both conditions test the same date column.',
        chips: V.val ? [s, d, '">="&B1', '"<="&B2'] : [d, '">="&B1', '"<="&B2'], fnKey: V.fn + '('
      });
    }
  });

  P.add({
    key: 'average-on-two-conditions', difficulty: 'Medium', level: 6, topics: ['AVERAGEIFS'], count: 3,
    make(r, v) {
      if (v === 1) {
        /* Average rate of the facilities of one type that are current (DPD 0). */
        const rows = H.loans(r, 18), S = H.table([COL.facility, COL.type, COL.dpd, COL.rate], rows);
        const t = S.R('Type'), q = S.R('DPD'), s = S.R('Rate (%)');
        const types = ranked(rows.filter((x) => x.dpd === 0), 'type');
        const one = types[0] || 'Working Capital', two = one === 'Working Capital' ? 'Term Loan' : 'Working Capital';
        const base = Object.assign({}, S.cells, { A1: 'Type', B1: one, A2: 'Average rate (%)' });
        /* The hidden sheet asks for the other type, and makes sure two of its facilities are current. */
        const other = where(rows, (x) => x.type === two);
        const edits = { B1: two };
        other.slice(0, 2).forEach((i) => { edits[S.cell('DPD', i)] = 0; });
        if (other.length) edits[S.cell('Rate (%)', other[0])] = (x) => x + 0.75;
        const alt = H.edit(base, edits);
        return H.problem({
          title: 'Average rate of current facilities of the type in B1', file: 'buku_pinjaman.xlsx', sheet: S, base, alt, answerCell: 'B2',
          statement: '<p>Dimas wants the average interest rate of the facilities that are current. In <b>B2</b>, average <b>Rate</b> (' + s + ') over the rows whose <b>Type</b> (' + t + ') is B1 and whose <b>DPD</b> (' + q + ') is 0. Keep the decimals.</p>',
          solution: '=AVERAGEIFS(' + s + ',' + t + ',B1,' + q + ',0)', alts: ['=AVERAGEIFS(' + s + ',' + q + ',0,' + t + ',B1)', '=AVERAGEIFS(' + s + ',' + t + ',B1,' + q + ',"=0")', '=SUMIFS(' + s + ',' + t + ',B1,' + q + ',0)/COUNTIFS(' + t + ',B1,' + q + ',0)'],
          hints: ['AVERAGEIFS takes the range to average first, then pairs of range and condition.', 'One pair is ' + t + ' with B1, the other is ' + q + ' with 0.', 'Shape: =AVERAGEIFS(' + s + ',' + t + ',B1,' + q + ',0).'],
          explain: 'AVERAGEIFS keeps the rows of type B1 with no days past due and averages their rates. It equals SUMIFS divided by COUNTIFS with the same two conditions.',
          chips: [s, t, q], fnKey: 'AVERAGEIFS('
        });
      }
      const V = [
        { title: 'Average net per line for a category in a region', file: 'penjualan_q4.xlsx', a: 'Category', b: 'Region', val: 'Net (Rp)', money: true,
          rows: () => { const g = r.sample(H.REGIONS, 3); return lines(r, 16, (s) => g.includes(s.region)); }, cols: [COL.date, COL.region, COL.category, COL.net] },
        null,
        { title: 'Average on-hand cartons for a warehouse and a category', file: 'stok_gudang.xlsx', a: 'Warehouse', b: 'Category', val: 'On hand',
          rows: () => H.stock(r, 16), cols: [COL.product, COL.category, COL.warehouse, COL.onhand] }
      ][v];
      const rows = V.rows(), S = H.table(V.cols, rows);
      const ka = V.cols.find((c) => c.h === V.a).k, kb = V.cols.find((c) => c.h === V.b).k;
      const combos = ranked(rows, (x) => x[ka] + '|' + x[kb]).map((k) => k.split('|'));
      const A = S.R(V.a), B = S.R(V.b), s = S.R(V.val);
      const base = Object.assign({}, S.cells, { A1: V.a, B1: combos[0][0], A2: V.b, B2: combos[0][1], A3: 'Average' });
      const vk = V.cols.find((c) => c.h === V.val).k;
      const hp = hiddenPair(rows, ka, kb, combos, vk, step(V.val), V.money ? (xs) => Math.round(mean(xs)) : mean), j = hp.j;
      const alt = H.edit(base, { B1: hp.pair[0], B2: hp.pair[1], [S.cell(V.val, j)]: (x) => x + step(V.val) });
      const core = 'AVERAGEIFS(' + s + ',' + A + ',B1,' + B + ',B2)';
      const sol = V.money ? '=ROUND(' + core + ',0)' : '=' + core;
      return H.problem({
        title: V.title, file: V.file, sheet: S, base, alt, answerCell: 'B3',
        statement: '<p>In <b>B3</b>, average <b>' + V.val + '</b> (' + s + ') over the rows where <b>' + V.a + '</b> (' + A + ') is B1 and <b>' + V.b + '</b> (' + B + ') is B2.' + (V.money ? ' Round it to whole rupiah with ROUND(...,0).' : ' Keep the decimals.') + '</p>',
        solution: sol, alts: V.money ? ['=ROUND(AVERAGEIFS(' + s + ',' + B + ',B2,' + A + ',B1),0)', '=ROUND(SUMIFS(' + s + ',' + A + ',B1,' + B + ',B2)/COUNTIFS(' + A + ',B1,' + B + ',B2),0)'] : ['=AVERAGEIFS(' + s + ',' + B + ',B2,' + A + ',B1)', '=SUMIFS(' + s + ',' + A + ',B1,' + B + ',B2)/COUNTIFS(' + A + ',B1,' + B + ',B2)'],
        hints: ['AVERAGEIFS takes the range to average first, then pairs of range and condition.', 'The pairs are ' + A + ' with B1 and ' + B + ' with B2.' + (V.money ? ' Wrap the whole thing in ROUND(...,0).' : ''), 'Shape: ' + sol + '.'],
        explain: 'AVERAGEIFS averages ' + V.val + ' over the rows that pass both conditions.' + (V.money ? ' ROUND(...,0) then gives whole rupiah.' : ' It equals SUMIFS divided by COUNTIFS with the same pairs.'),
        chips: [s, A, B], fnKey: 'AVERAGEIFS('
      });
    }
  });

  P.add({
    key: 'group-figure-on-every-line', difficulty: 'Medium', level: 6, topics: ['SUMIF', 'COUNTIF', 'MAXIFS'], count: 4,
    make(r, v) {
      const V = [
        { title: 'Distributor total beside every sales line', label: 'Distributor total (Rp)', key: 'Distributor', val: 'Net (Rp)',
          rows: () => { const g = names(r, 4); return lines(r, 14, (s) => g.includes(s.distributor)); }, cols: [COL.date, COL.dist, COL.product, COL.net],
          f: (K, k, X) => 'SUMIF(' + K + ',' + k + ',' + X + ')', alt2: (K, k, X) => '=SUMIFS(' + X + ',' + K + ',' + k + ')', fn: 'SUMIF(',
          job: 'the total <b>Net</b> of that line\'s distributor', why: 'SUMIF adds Net for every row with the same distributor as this line' },
        { title: 'Number of lines for the product on every line', label: 'Lines of product', key: 'Product',
          rows: () => { const ps = r.sample(H.PRODUCTS, 5).map((x) => x.name); return lines(r, 14, (s) => ps.includes(s.product)); }, cols: [COL.date, COL.dist, COL.product, COL.cartons],
          f: (K, k) => 'COUNTIF(' + K + ',' + k + ')', alt2: (K, k) => '=COUNTIFS(' + K + ',' + k + ')', fn: 'COUNTIF(',
          job: 'how many lines in the whole table have the same <b>Product</b> as that line', why: 'COUNTIF counts the rows whose product equals this line\'s product' },
        { title: 'Each line\'s share of its distributor\'s net', label: 'Share of distributor', key: 'Distributor', val: 'Net (Rp)', share: true,
          rows: () => { const g = names(r, 3); return lines(r, 14, (s) => g.includes(s.distributor)); }, cols: [COL.date, COL.dist, COL.product, COL.net],
          f: (K, k, X) => 'SUMIF(' + K + ',' + k + ',' + X + ')', fn: 'SUMIF(',
          job: 'that line\'s <b>Net</b> as a share (a decimal such as 0.25, not rounded) of its distributor\'s total Net', why: 'the line\'s Net is divided by the SUMIF total of its distributor' },
        { title: 'Biggest order in the region beside every line', label: 'Region max (cartons)', key: 'Region', val: 'Cartons',
          rows: () => { const g = r.sample(H.REGIONS, 3); return lines(r, 14, (s) => g.includes(s.region)); }, cols: [COL.date, COL.dist, COL.region, COL.cartons],
          f: (K, k, X) => 'MAXIFS(' + X + ',' + K + ',' + k + ')', alt2: (K, k, X) => '=MAX(FILTER(' + X + ',' + K + '=' + k + '))', fn: 'MAXIFS(',
          job: 'the largest <b>Cartons</b> of any line in the same <b>Region</b> as that line', why: 'MAXIFS returns the largest cartons among the rows of the same region' }
      ][v];
      const rows = V.rows(), S = H.table(V.cols, rows);
      const K = S.A(V.key), k = S.cell(V.key, 0), X = V.val ? S.A(V.val) : null, x0 = V.val ? S.cell(V.val, 0) : null;
      const Kr = '$' + S.col(V.key) + '$' + S.first + ':$' + S.col(V.key) + '$' + S.last;
      const f = V.f(K, k, X);
      const sol = V.share ? '=' + x0 + '/' + f : '=' + f;
      const alts = [];
      if (V.alt2) alts.push(V.alt2(K, k, X));
      const kk = V.cols.find((c) => c.h === V.key).k;
      /* Rows pinned only (B$5:B$18) fill down as well. */
      const pinRow = (A) => A.replace(/\$([A-Z])\$/g, '$1$');
      alts.push(V.share ? '=' + x0 + '/' + V.f(pinRow(K), k, pinRow(X)) : '=' + V.f(pinRow(K), k, X ? pinRow(X) : X));
      /* The hidden sheet: two figures are corrected and one line is booked to another group. */
      const keys = ranked(rows, kk);
      const edits = {};
      const moved = rows.findIndex((x) => x[kk] === keys[keys.length - 1]);
      edits[S.cell(V.key, moved)] = keys[0];
      if (V.val) { edits[S.cell(V.val, 0)] = (x) => x + step(V.val) * 2; edits[S.cell(V.val, rows.length - 1)] = (x) => x + step(V.val); }
      const alt = H.edit(S.cells, edits);
      return H.problem({
        title: V.title, file: 'penjualan_q4.xlsx', sheet: S, base: S.cells, alt, fill: true, ansLabel: V.label, precision: V.share ? 4 : 2,
        statement: '<p>In <b>' + S.next + S.first + '</b>, show ' + V.job + ', then fill it down to <b>' + S.next + S.last + '</b>. Pin the ranges with $ so they stay put while the formula fills down.</p>',
        solution: sol, alts,
        hints: ['The condition is this line\'s own ' + V.key + ', ' + k + ', which should move as you fill down.', 'The ranges must not move: write them with $, like ' + Kr + '.', 'Shape: ' + sol + '.'],
        explain: 'On every row ' + V.why + '. The ranges are pinned with $ and the condition ' + k + ' is relative, so each row asks about its own ' + V.key + '.',
        chips: V.val ? [K, X, k] : [K, k], fnKey: V.fn
      });
    }
  });

  P.add({
    key: 'running-figure-by-group', difficulty: 'Medium', level: 7, topics: ['SUMIF', 'COUNTIF', 'IF'], count: 3,
    make(r, v) {
      const g = names(r, 3);
      const rows = lines(r, 14, (s) => g.includes(s.distributor)), S = H.table([COL.date, COL.dist, COL.product, COL.net], rows);
      const b = S.col('Distributor'), n = S.col('Net (Rp)'), f = S.first, L = S.next;
      const V = [
        { title: 'Running net per distributor down the sales list', label: 'Running net (Rp)', sol: '=SUMIF($' + b + '$' + f + ':' + b + f + ',' + b + f + ',$' + n + '$' + f + ':' + n + f + ')',
          alts: ['=SUMIFS($' + n + '$' + f + ':' + n + f + ',$' + b + '$' + f + ':' + b + f + ',' + b + f + ')'], fn: 'SUMIF(',
          job: 'the distributor\'s Net so far: the total of that distributor\'s lines from row ' + f + ' down to this row', why: 'SUMIF adds the Net of this distributor\'s lines in the growing range' },
        { title: 'Order number of each line for its distributor', label: 'Order no.', sol: '=COUNTIF($' + b + '$' + f + ':' + b + f + ',' + b + f + ')',
          alts: ['=COUNTIFS($' + b + '$' + f + ':' + b + f + ',' + b + f + ')', '=COUNTIF($' + b + '$4:' + b + (f - 1) + ',' + b + f + ')+1'], fn: 'COUNTIF(',
          job: 'which order this is for its distributor: 1 on the distributor\'s first line, 2 on the second, and so on', why: 'COUNTIF counts this distributor\'s lines in the growing range' },
        { title: 'Mark each distributor\'s first order as First or Repeat', label: 'First or repeat', sol: '=IF(COUNTIF($' + b + '$' + f + ':' + b + f + ',' + b + f + ')=1,"First","Repeat")',
          alts: ['=IF(COUNTIF($' + b + '$' + f + ':' + b + f + ',' + b + f + ')>1,"Repeat","First")', '=IF(COUNTIF($' + b + '$4:' + b + (f - 1) + ',' + b + f + ')=0,"First","Repeat")'], fn: 'COUNTIF(',
          job: 'the word <b>First</b> on a distributor\'s first line in the list and <b>Repeat</b> on every later one', why: 'COUNTIF counts this distributor in the growing range, and IF turns a count of 1 into First' }
      ][v];
      /* The hidden sheet: an early line is booked to another distributor and two amounts are corrected. */
      const keys = ranked(rows, 'distributor');
      const i1 = rows.findIndex((x, i) => i > 0 && x.distributor !== keys[0]);
      const alt = H.edit(S.cells, { [S.cell('Distributor', i1)]: keys[0], [S.cell('Net (Rp)', 1)]: (x) => x + 2000000, [S.cell('Net (Rp)', 6)]: (x) => x + 1000000 });
      return H.problem({
        title: V.title, file: 'penjualan_q4.xlsx', sheet: S, base: S.cells, alt, fill: true, ansLabel: V.label,
        statement: '<p>The lines are in date order. In <b>' + L + f + '</b>, show ' + V.job + '. Fill it down to <b>' + L + S.last + '</b>.</p>',
        solution: V.sol, alts: V.alts,
        hints: ['Use a range that grows as you fill down: its start is pinned ($' + b + '$' + f + ') and its end is this row (' + b + f + ').', 'In the first row the range is $' + b + '$' + f + ':' + b + f + '; on the next row it becomes $' + b + '$' + f + ':' + b + (f + 1) + '.', 'Shape: ' + V.sol + '.'],
        explain: 'The range $' + b + '$' + f + ':' + b + f + ' has a pinned start and a moving end, so it grows by one row each time it is filled down. On every row ' + V.why + '.',
        chips: ['$' + b + '$' + f + ':' + b + f, b + f].concat(v === 0 ? ['$' + n + '$' + f + ':' + n + f] : v === 2 ? ['"First"', '"Repeat"'] : []), fnKey: V.fn
      });
    }
  });

  P.add({
    key: 'blank-cells-as-a-condition', difficulty: 'Medium', level: 7, topics: ['SUMIFS', 'COUNTIFS', 'blanks'], count: 2,
    make(r, v) {
      const who = names(r, 4);
      const rows = bills(r, 16, who), S = H.table([COL.ref, COL.dist, COL.due, COL.amount, COL.paid], rows);
      const p = S.R('Paid on'), d = S.R('Due'), a = S.R('Amount (Rp)'), c = S.R('Distributor');
      if (v === 0) {
        const dues = sorted(rows, 'due');
        const asOf = [dues[9], dues[5]];
        const base = Object.assign({}, S.cells, { A1: 'As of', B1: H.toSerial(asOf[0]), A2: 'Overdue (Rp)' });
        /* The hidden sheet: an earlier as-of date, and one paid invoice turns out unpaid. */
        const paidEarly = rows.findIndex((x) => x.paid && x.due < asOf[1]);
        const alt = H.edit(base, { B1: H.toSerial(asOf[1]) });
        if (paidEarly >= 0) delete alt[S.cell('Paid on', paidEarly)];
        return H.problem({
          title: 'Unpaid invoices already past due on the date in B1', file: 'piutang_q4.xlsx', sheet: S, base, alt, answerCell: 'B2', dateCells: ['B1'],
          statement: '<p>Bu Lestari in credit control needs the overdue amount. An invoice is overdue when <b>Paid on</b> (' + p + ') is blank and its <b>Due</b> date (' + d + ') is <b>before</b> the date in B1. In <b>B2</b>, add up <b>Amount</b> (' + a + ') for those invoices.</p>',
          solution: '=SUMIFS(' + a + ',' + p + ',"",' + d + ',"<"&B1)', alts: ['=SUMIFS(' + a + ',' + d + ',"<"&B1,' + p + ',"")', '=SUMIFS(' + a + ',' + p + ',"=",' + d + ',"<"&B1)', '=SUMPRODUCT((' + p + '="")*(' + d + '<B1)*' + a + ')'],
          hints: ['The condition "" (two double quotes with nothing between) matches blank cells.', 'Add a second condition on the due date: "<"&B1 for before B1.', 'Shape: =SUMIFS(' + a + ',' + p + ',"",' + d + ',"<"&B1).'],
          explain: '"" keeps the invoices with no payment date, and "<"&B1 keeps those due before B1, so SUMIFS adds only unpaid amounts that are already late. An invoice due on B1 itself is not late yet.',
          chips: [a, p, d, '""', '"<"&B1'], fnKey: 'SUMIFS('
        });
      }
      const open = ranked(rows.filter((x) => !x.paid), 'distributor');
      const one = open[0], two = open[1] || who.find((x) => x !== one);
      const base = Object.assign({}, S.cells, { A1: 'Distributor', B1: one, A2: 'Open invoices' });
      const alt = H.edit(base, { B1: two });
      /* The hidden sheet: one of the new distributor's paid invoices turns out unpaid, or two when the counts would tie. */
      const paidTwo = where(rows, (x) => x.distributor === two && x.paid);
      const openOf = (n) => rows.filter((x) => x.distributor === n && !x.paid).length;
      const undo = paidTwo.slice(0, openOf(two) + 1 === openOf(one) ? 2 : 1);
      undo.forEach((i) => { delete alt[S.cell('Paid on', i)]; });
      if (openOf(two) + undo.length === openOf(one)) alt[S.cell('Distributor', rows.findIndex((x) => x.distributor !== two && !x.paid))] = two;
      return H.problem({
        title: 'Count the open invoices of the distributor in B1', file: 'piutang_q4.xlsx', sheet: S, base, alt, answerCell: 'B2',
        statement: '<p>An invoice is open while its <b>Paid on</b> cell (' + p + ') is blank. In <b>B2</b>, count the open invoices of the distributor in B1 (' + c + ').</p>',
        solution: '=COUNTIFS(' + c + ',B1,' + p + ',"")', alts: ['=COUNTIFS(' + p + ',"",' + c + ',B1)', '=COUNTIF(' + c + ',B1)-COUNTIFS(' + c + ',B1,' + p + ',"<>")', '=SUMPRODUCT((' + c + '=B1)*(' + p + '=""))'],
        hints: ['A blank cell is matched by the condition "" with nothing between the quotes.', 'You need two conditions: the distributor is B1, and Paid on is blank.', 'Shape: =COUNTIFS(' + c + ',B1,' + p + ',"").'],
        explain: 'COUNTIFS counts the rows where the distributor is B1 and the payment date is blank. "<>" would do the opposite and count the paid ones.',
        chips: [c, p, '""'], fnKey: 'COUNTIFS('
      });
    }
  });

  /* ================= Hard ================= */

  P.add({
    key: 'two-way-summary-table', difficulty: 'Hard', level: 8, topics: ['SUMIFS', 'INDEX', 'MATCH'], count: 4,
    make(r, v) {
      const QM = ['Oct', 'Nov', 'Dec'];
      const qm = (d) => QM[month(H.Q4(d)) - 10];
      if (v === 0 || v === 3) {
        /* A summary of net sales by month, built from the full sales book. */
        const pool = H.db().sales;
        let rows;
        if (v === 0) {
          const g = names(r, 4);
          rows = [];
          g.forEach((dn) => H.CATEGORIES.forEach((cat) => {
            const own = pool.filter((s) => s.distributor === dn && s.category === cat);
            const o = { distributor: dn, category: cat };
            QM.forEach((m) => { o[m] = own.filter((s) => qm(s.date) === m).reduce((a, s) => a + s.net, 0); });
            rows.push(o);
          }));
        } else {
          rows = r.sample(H.db().distributors, 10).map((d) => {
            const own = pool.filter((s) => s.distributor === d.name);
            const o = { distributor: d.name, region: d.region };
            QM.forEach((m) => { o[m] = own.filter((s) => qm(s.date) === m).reduce((a, s) => a + s.net, 0); });
            return o;
          });
          if (!rows.some((x) => x.distributor.startsWith('CV'))) rows[0] = Object.assign({}, rows[0], { distributor: 'CV Nuri Makmur' });
        }
        const cols = (v === 0 ? [COL.dist, COL.category] : [COL.dist, COL.region]).concat(QM.map((m) => ({ h: m, k: m, num: true })));
        const S = H.table(cols, rows);
        const M = '$C$' + S.first + ':$E$' + S.last, Hd = '$C$4:$E$4';
        if (v === 0) {
          const B = S.A('Category'), pickCat = r.shuffle(H.CATEGORIES);
          const base = Object.assign({}, S.cells, { A1: 'Category', B1: pickCat[0], A2: 'Month', B2: 'Nov', A3: 'Net (Rp)' });
          const j = rows.findIndex((x) => x.category === pickCat[1]);
          const alt = H.edit(base, { B1: pickCat[1], B2: 'Dec', [S.cell('Dec', j)]: (x) => x + 3000000 });
          return H.problem({
            title: 'Category net for the month in B2 from a summary grid', file: 'ringkasan_q4.xlsx', sheet: S, base, alt, answerCell: 'B3',
            statement: '<p>The sheet is a summary: one row per distributor and category, one column per month (C4:E4 hold Oct, Nov, Dec). In <b>B3</b>, add up the net of the category in B1 for the month named in B2.</p>',
            solution: '=SUMIFS(INDEX(' + M + ',0,MATCH(B2,' + Hd + ',0)),' + B + ',B1)',
            alts: ['=SUMPRODUCT((' + B + '=B1)*(' + Hd + '=B2)*' + M + ')', '=SUMIFS(XLOOKUP(B2,' + Hd + ',' + M + '),' + B + ',B1)', '=SUMIF(' + B + ',B1,INDEX(' + M + ',0,MATCH(B2,' + Hd + ',0)))'],
            hints: ['SUMIFS needs a single column to add, so first pick the month column out of C5:E' + S.last + '.', 'MATCH(B2,' + Hd + ',0) says which column B2 is; INDEX(' + M + ',0,that number) returns that whole column.', 'Shape: =SUMIFS(INDEX(' + M + ',0,MATCH(B2,' + Hd + ',0)),' + B + ',B1).'],
            explain: 'INDEX with row 0 returns a whole column, and MATCH finds which one from the month name, so SUMIFS adds the right month for the category in B1. SUMPRODUCT can test the rows and the header at once instead.',
            chips: [M, Hd, B, 'MATCH(B2,' + Hd + ',0)'], fnKey: 'SUMIFS('
          });
        }
        const A = S.A('Distributor');
        const base = Object.assign({}, S.cells, { A1: 'Month', B1: 'Oct', A2: 'CV total (Rp)' });
        const j = rows.findIndex((x) => x.distributor.startsWith('CV'));
        const alt = H.edit(base, { B1: 'Dec', [S.cell('Dec', j)]: (x) => x + 2500000 });
        return H.problem({
          title: 'Net of all CV distributors for the month in B1', file: 'ringkasan_q4.xlsx', sheet: S, base, alt, answerCell: 'B2',
          statement: '<p>The sheet sums net sales per distributor, with one column per month (C4:E4 hold Oct, Nov, Dec). In <b>B2</b>, add up the month named in B1 for the distributors whose name starts with CV.</p>',
          solution: '=SUMIFS(INDEX(' + M + ',0,MATCH(B1,' + Hd + ',0)),' + A + ',"CV*")',
          alts: ['=SUMIFS(XLOOKUP(B1,' + Hd + ',' + M + '),' + A + ',"CV*")', '=SUMIF(' + A + ',"CV*",INDEX(' + M + ',0,MATCH(B1,' + Hd + ',0)))'],
          hints: ['First choose the month column: MATCH(B1,' + Hd + ',0) gives its position.', 'INDEX(' + M + ',0,position) returns that column for SUMIFS to add, and "CV*" keeps the CV names.', 'Shape: =SUMIFS(INDEX(' + M + ',0,MATCH(B1,' + Hd + ',0)),' + A + ',"CV*").'],
          explain: 'INDEX and MATCH hand SUMIFS the one month column named in B1, and the wildcard "CV*" keeps only the CV distributors. XLOOKUP on the header row returns the same column.',
          chips: [M, Hd, A, '"CV*"'], fnKey: 'SUMIFS('
        });
      }
      if (v === 1) {
        /* Facilities with three money columns; pick the column by its header and add only the rows that are not restructured. */
        const rows = H.loans(r, 16).map((x) => Object.assign({}, x, { undrawn: x.limit - x.outstanding }));
        const S = H.table([COL.facility, COL.region, COL.restructured, COL.limit, COL.outstanding, { h: 'Undrawn (Rp)', k: 'undrawn', num: true }], rows);
        const M = '$D$' + S.first + ':$F$' + S.last, Hd = '$D$4:$F$4', B = S.A('Region'), C = S.A('Restructured');
        const regs = ranked(rows.filter((x) => x.restructured === 'N'), 'region');
        const base = Object.assign({}, S.cells, { A1: 'Region', B1: regs[0], A2: 'Measure', B2: 'Undrawn (Rp)', A3: 'Total (Rp)' });
        const j = rows.findIndex((x) => x.region === regs[1] && x.restructured === 'N');
        const alt = H.edit(base, { B1: regs[1], B2: 'Outstanding (Rp)', [S.cell('Outstanding (Rp)', j)]: (x) => x + 4000000 });
        return H.problem({
          title: 'Region total of the loan measure named in B2, not restructured', file: 'buku_pinjaman.xlsx', sheet: S, base, alt, answerCell: 'B3',
          statement: '<p>B2 holds one of the headers in D4:F4 (Limit, Outstanding or Undrawn). In <b>B3</b>, add up that column for the facilities in the region in B1 whose <b>Restructured</b> flag is <b>N</b>.</p>',
          solution: '=SUMIFS(INDEX(' + M + ',0,MATCH(B2,' + Hd + ',0)),' + B + ',B1,' + C + ',"N")',
          alts: ['=SUMIFS(XLOOKUP(B2,' + Hd + ',' + M + '),' + B + ',B1,' + C + ',"N")', '=SUMPRODUCT((' + B + '=B1)*(' + C + '="N")*(' + Hd + '=B2)*' + M + ')', '=SUMIFS(INDEX(' + M + ',0,MATCH(B2,' + Hd + ',0)),' + B + ',B1,' + C + ',"<>Y")'],
          hints: ['The column to add depends on B2, so find it with MATCH(B2,' + Hd + ',0).', 'INDEX(' + M + ',0,MATCH(...)) is the column to add; then give SUMIFS two pairs: region and the N flag.', 'Shape: =SUMIFS(INDEX(' + M + ',0,MATCH(B2,' + Hd + ',0)),' + B + ',B1,' + C + ',"N").'],
          explain: 'MATCH turns the header in B2 into a column number and INDEX returns that column, so one formula works for any of the three measures. SUMIFS then keeps the region in B1 and the facilities that were not restructured.',
          chips: [M, Hd, B, C, '"N"'], fnKey: 'SUMIFS('
        });
      }
      /* v 2: a stock grid, products down and warehouses across; a SKU family and a warehouse from cells. */
      const rows = H.PRODUCTS.map((p) => { const o = { sku: p.sku, product: p.name }; H.WAREHOUSES.forEach((w) => { o[w] = r.int(0, 40) * 10; }); return o; });
      const S = H.table([COL.sku, COL.product].concat(H.WAREHOUSES.map((w) => ({ h: w, k: w, num: true }))), rows);
      const M = '$C$' + S.first + ':$G$' + S.last, Hd = '$C$4:$G$4', A = S.A('SKU');
      const fams = r.shuffle(['MLK', 'KSJ', 'SNY', 'SRW']), whs = r.sample(H.WAREHOUSES, 2);
      const base = Object.assign({}, S.cells, { A1: 'SKU family', B1: fams[0], A2: 'Warehouse', B2: whs[0], A3: 'On hand' });
      const j = rows.findIndex((x) => family(x.sku) === fams[1]);
      const alt = H.edit(base, { B1: fams[1], B2: whs[1], [S.cell(whs[1], j)]: (x) => x + 50 });
      return H.problem({
        title: 'Stock grid: one SKU family in the warehouse named in B2', file: 'stok_gudang.xlsx', sheet: S, base, alt, answerCell: 'B3',
        statement: '<p>The grid has one row per SKU and one column per warehouse (C4:G4). In <b>B3</b>, add up the on-hand cartons of the SKUs that start with the family code in B1 (such as KSJ) in the warehouse named in B2.</p>',
        solution: '=SUMIFS(INDEX(' + M + ',0,MATCH(B2,' + Hd + ',0)),' + A + ',B1&"*")',
        alts: ['=SUMIFS(XLOOKUP(B2,' + Hd + ',' + M + '),' + A + ',B1&"*")', '=SUMIF(' + A + ',B1&"-*",INDEX(' + M + ',0,MATCH(B2,' + Hd + ',0)))'],
        hints: ['Pick the warehouse column first: INDEX(' + M + ',0,MATCH(B2,' + Hd + ',0)).', 'The rows are chosen by a starts-with pattern built from B1: B1&"*".', 'Shape: =SUMIFS(INDEX(' + M + ',0,MATCH(B2,' + Hd + ',0)),' + A + ',B1&"*").'],
        explain: 'INDEX and MATCH return the column of the warehouse in B2, and B1&"*" keeps the SKUs of that family. Two inputs drive one SUMIFS, so the grid answers any family and warehouse pair.',
        chips: [M, Hd, A, 'B1&"*"'], fnKey: 'SUMIFS('
      });
    }
  });

  P.add({
    key: 'either-or-totals', difficulty: 'Hard', level: 9, topics: ['SUMIFS', 'COUNTIFS', 'OR logic'], count: 3,
    make(r, v) {
      if (v === 0) {
        /* OR on two columns at once: either of two regions, and either of two categories. */
        const g = r.sample(H.REGIONS, 4);
        const rows = lines(r, 18, (s) => g.includes(s.region)), S = H.table([COL.date, COL.region, COL.category, COL.net], rows);
        const q = S.R('Region'), c = S.R('Category'), s = S.R('Net (Rp)');
        const cats = ranked(rows, 'category').slice(0, 2), inCat = (x) => cats.includes(x.category);
        const regs = ranked(rows.filter(inCat), 'region').concat(g).filter((x, i, a) => a.indexOf(x) === i);
        const pair = [regs[0], regs[2]], other = [regs[1], regs[3]];
        const base = Object.assign({}, S.cells, { A1: 'Region 1', B1: pair[0], A2: 'Region 2', B2: pair[1], A3: cats[0] + ' or ' + cats[1] + ' (Rp)' });
        /* The hidden sheet: two other regions in B1 and B2, and one of their lines is corrected. */
        const j = rows.findIndex((x) => x.region === other[0] && inCat(x));
        const edits = { B1: other[0], B2: other[1] };
        if (j >= 0) edits[S.cell('Net (Rp)', j)] = (x) => x + 2200000;
        const alt = H.edit(base, edits);
        const one = (reg, cat) => 'SUMIFS(' + s + ',' + q + ',' + reg + ',' + c + ',"' + cat + '")';
        const sol = '=' + one('B1', cats[0]) + '+' + one('B1', cats[1]) + '+' + one('B2', cats[0]) + '+' + one('B2', cats[1]);
        const sp = '=SUMPRODUCT(((' + q + '=B1)+(' + q + '=B2))*((' + c + '="' + cats[0] + '")+(' + c + '="' + cats[1] + '"))*' + s + ')';
        return H.problem({
          title: cats[0] + ' or ' + cats[1] + ' net for the two regions in B1 and B2', file: 'penjualan_q4.xlsx', sheet: S, base, alt, answerCell: 'B3',
          statement: '<p>In <b>B3</b>, add up <b>Net</b> (' + s + ') for the lines whose <b>Region</b> (' + q + ') is the one in B1 <b>or</b> the one in B2, and whose <b>Category</b> (' + c + ') is <b>' + cats[0] + '</b> <b>or</b> <b>' + cats[1] + '</b>. A line counts when both tests pass.</p>',
          solution: sol,
          alts: [sp, '=' + one('B1', cats[0]) + '+' + one('B2', cats[0]) + '+' + one('B1', cats[1]) + '+' + one('B2', cats[1]),
            '=SUMPRODUCT(((' + q + '=B1)+(' + q + '=B2))*((' + c + '="' + cats[1] + '")+(' + c + '="' + cats[0] + '")),' + s + ')'],
          hints: ['SUMIFS joins its conditions with AND, so each choice needs its own SUMIFS: two regions times two categories makes four.', 'A line has one region and one category, so the four totals never share a line and can be added. In SUMPRODUCT, + between two tests works as OR and * as AND.', 'Shape: =SUMIFS(net,region,B1,category,"' + cats[0] + '")+SUMIFS(net,region,B1,category,"' + cats[1] + '")+ the same two for B2, or ' + sp + '.'],
          explain: 'Every line belongs to one region and one category, so the four region and category totals do not overlap and their sum is the OR on both columns. SUMPRODUCT gets there in one step: each bracket of + is 1 when either test passes, and * keeps the lines where both brackets are 1.',
          chips: [s, q, c, '"' + cats[0] + '"', '"' + cats[1] + '"'], fnKey: 'SUMIFS('
        });
      }
      if (v === 1) {
        /* Overlapping OR: restructured, or more than B1 days past due. Count the overlap once. */
        const rows = H.loans(r, 16);
        rows[2] = Object.assign({}, rows[2], { restructured: 'Y', dpd: 120 });
        rows[7] = Object.assign({}, rows[7], { restructured: 'N', dpd: 95 });
        rows[11] = Object.assign({}, rows[11], { restructured: 'Y', dpd: 0 });
        /* One facility sits on the visible limit (90, so "more than" leaves it out) and one between the two limits. */
        rows[4] = Object.assign({}, rows[4], { restructured: 'N', dpd: 90 });
        rows[13] = Object.assign({}, rows[13], { restructured: 'N', dpd: 75 });
        const S = H.table([COL.facility, COL.dist, COL.outstanding, COL.restructured, COL.dpd], rows);
        const o = S.R('Outstanding (Rp)'), y = S.R('Restructured'), d = S.R('DPD');
        const base = Object.assign({}, S.cells, { A1: 'DPD above', B1: 90, A2: 'Watch list (Rp)' });
        const alt = H.edit(base, { B1: 60, [S.cell('Outstanding (Rp)', 2)]: (x) => x + 5000000 });
        return H.problem({
          title: 'Watch list balance: restructured or past the DPD in B1', file: 'buku_pinjaman.xlsx', sheet: S, base, alt, answerCell: 'B2',
          statement: '<p>Citra\'s watch list holds every facility that is restructured (<b>Y</b> in ' + y + ') <b>or</b> whose <b>DPD</b> (' + d + ') is <b>more than</b> the number of days in B1. In <b>B2</b>, add up their <b>Outstanding</b> (' + o + '), counting a facility that meets both tests only once.</p>',
          solution: '=SUMIF(' + y + ',"Y",' + o + ')+SUMIF(' + d + ',">"&B1,' + o + ')-SUMIFS(' + o + ',' + y + ',"Y",' + d + ',">"&B1)',
          alts: ['=SUMIF(' + y + ',"Y",' + o + ')+SUMIFS(' + o + ',' + y + ',"N",' + d + ',">"&B1)', '=SUMPRODUCT(((' + y + '="Y")+(' + d + '>B1)>0)*' + o + ')', '=SUMIFS(' + o + ',' + y + ',"Y")+SUMIFS(' + o + ',' + y + ',"<>Y",' + d + ',">"&B1)'],
          hints: ['Adding the restructured total and the past-due total counts a facility that is both twice.', 'Take the overlap away once: a SUMIFS with both conditions. Or add the restructured ones plus the past-due ones that are not restructured.', 'Shape: =SUMIF(' + y + ',"Y",' + o + ')+SUMIF(' + d + ',">"&B1,' + o + ')-SUMIFS(' + o + ',' + y + ',"Y",' + d + ',">"&B1).'],
          explain: 'A or B equals A plus B minus the rows that are both, so subtracting the two-condition SUMIFS removes the double count. Splitting into restructured plus past due and not restructured gives the same total.',
          chips: [o, y, d, '">"&B1', '"Y"'], fnKey: 'SUMIFS('
        });
      }
      const rows = lines(r, 16), S = H.table([COL.date, COL.dist, COL.product, COL.cartons], rows);
      const b = S.R('Distributor'), c = S.R('Cartons');
      /* The limit is low enough that at least one CV line is also big, so the overlap matters. */
      const cvMax = Math.max.apply(null, rows.filter((x) => x.distributor.startsWith('CV')).map((x) => x.cartons).concat([0]));
      const all = sorted(rows, 'cartons');
      const lim = [Math.min(cvMax || all[9], all[9]), all[6]];
      const base = Object.assign({}, S.cells, { A1: 'Cartons at least', B1: lim[0], A2: 'Lines' });
      const ud = where(rows, (x) => x.distributor.startsWith('UD'));
      const alt = H.edit(base, { B1: lim[1], [S.cell('Cartons', ud[0] === undefined ? 0 : ud[0])]: (x) => x + 40 });
      return H.problem({
        title: 'Count lines from CV distributors or of at least B1 cartons', file: 'penjualan_q4.xlsx', sheet: S, base, alt, answerCell: 'B2',
        statement: '<p>In <b>B2</b>, count the lines that come from a distributor whose name starts with <b>CV</b> (' + b + ') <b>or</b> that have <b>at least</b> the cartons in B1 (' + c + '). A line that meets both counts once.</p>',
        solution: '=COUNTIF(' + b + ',"CV*")+COUNTIF(' + c + ',">="&B1)-COUNTIFS(' + b + ',"CV*",' + c + ',">="&B1)',
        alts: ['=COUNTIF(' + b + ',"CV*")+COUNTIFS(' + b + ',"<>CV*",' + c + ',">="&B1)', '=COUNTIFS(' + b + ',"CV*")+COUNTIFS(' + c + ',">="&B1,' + b + ',"<>CV*")'],
        hints: ['Count each group, then take away the lines that are in both.', 'The overlap is a COUNTIFS with both conditions: "CV*" and ">="&B1.', 'Shape: =COUNTIF(' + b + ',"CV*")+COUNTIF(' + c + ',">="&B1)-COUNTIFS(' + b + ',"CV*",' + c + ',">="&B1).'],
        explain: 'Adding the two counts counts the big CV lines twice, so the COUNTIFS of both conditions is taken away once. Another way: every CV line, plus the big lines that are not CV ("<>CV*").',
        chips: [b, c, '"CV*"', '">="&B1'], fnKey: 'COUNTIFS('
      });
    }
  });

  P.add({
    key: 'one-month-windows', difficulty: 'Hard', level: 8, topics: ['SUMIFS', 'EOMONTH', 'EDATE'], count: 3,
    make(r, v) {
      const g = r.sample(H.REGIONS, 3);
      const rows = lines(r, 18, (s) => g.includes(s.region));
      const V = [
        { title: 'Net growth of the region in B1 for the month in B2', k: 'region', h: 'Region', val: 'Net (Rp)', vk: 'net', cols: [COL.date, COL.region, COL.category, COL.net], label: 'Region' },
        { title: 'Change in cartons of the distributor in B1 for the month in B2', k: 'distributor', h: 'Distributor', val: 'Cartons', vk: 'cartons', cols: [COL.date, COL.dist, COL.product, COL.cartons], label: 'Distributor' },
        { title: 'Share of the category in B1 in the net of the month in B2', k: 'category', h: 'Category', val: 'Net (Rp)', vk: 'net', cols: [COL.date, COL.region, COL.category, COL.net], label: 'Category' }
      ][v];
      const S = H.table(V.cols, rows);
      const d = S.R('Date'), K = S.R(V.h), s = S.R(V.val);
      /* Pick key and month pairs where both the month and the month before have sales. */
      const ok = [];
      for (const key of ranked(rows, V.k)) for (const m of [11, 12]) {
        const cur = sumOf(rows, (x) => x[V.k] === key && month(x.date) === m, V.vk);
        const prev = sumOf(rows, (x) => x[V.k] === key && month(x.date) === m - 1, V.vk);
        if (cur > 0 && (v === 2 || prev > 0) && (v !== 1 || cur !== prev)) ok.push({ key, m });
      }
      /* Prefer the pairs with the most lines in both months, so the windows hold real totals. */
      const lineCount = (o) => Math.min(rows.filter((x) => x[V.k] === o.key && month(x.date) === o.m).length, rows.filter((x) => x[V.k] === o.key && month(x.date) === o.m - 1).length);
      ok.sort((a, b) => lineCount(b) - lineCount(a));
      while (ok.length < 2) ok.push({ key: ranked(rows, V.k)[ok.length], m: 12 - ok.length });
      const two = ok.find((x) => x.key !== ok[0].key && x.m !== ok[0].m) || ok[1];
      const first = (m) => H.toSerial('2026-' + m + '-01');
      const base = Object.assign({}, S.cells, { A1: V.label, B1: ok[0].key, A2: 'Month (first day)', B2: first(ok[0].m), A3: v === 0 ? 'Growth' : v === 1 ? 'Change (cartons)' : 'Share' });
      const j = rows.findIndex((x) => x[V.k] === two.key && month(x.date) === two.m);
      const edits = { B1: two.key, B2: first(two.m), [S.cell(V.val, j)]: (x) => x + (V.vk === 'net' ? 1500000 : 5) };
      /*
       * The hidden sheet also moves some lines to the first and last day of their month (they stay in the same month),
       * so a window that leaves out its first or last day, or lets the two months overlap, gives a different answer.
       */
      const lastDay = (m) => (m === 12 ? H.toSerial('2027-01-01') : first(m + 1)) - 1;
      const toEdges = (m, keep) => {
        const own = where(rows, (x) => keep(x) && month(x.date) === m);
        if (own.length) edits[S.cell('Date', own[0])] = first(m);
        if (own.length > 1) edits[S.cell('Date', own[own.length - 1])] = lastDay(m);
        return own;
      };
      const mine = toEdges(two.m, (x) => x[V.k] === two.key);
      if (v !== 2) toEdges(two.m - 1, (x) => x[V.k] === two.key);
      if (mine.length < 2) {
        /* With one line in the month, the month's last line of another group is booked to B1 and dated on the last day. */
        const o = where(rows, (x) => month(x.date) === two.m && x[V.k] !== two.key);
        if (o.length) { edits[S.cell('Date', o[o.length - 1])] = lastDay(two.m); edits[S.cell(V.h, o[o.length - 1])] = two.key; }
      }
      const alt = H.edit(base, edits);
      const cur = 'SUMIFS(' + s + ',' + K + ',B1,' + d + ',">="&B2,' + d + ',"<="&EOMONTH(B2,0))';
      const prev = 'SUMIFS(' + s + ',' + K + ',B1,' + d + ',">="&EDATE(B2,-1),' + d + ',"<"&B2)';
      const cur2 = 'SUMIFS(' + s + ',' + K + ',B1,' + d + ',">="&B2,' + d + ',"<"&EDATE(B2,1))';
      const prev2 = 'SUMIFS(' + s + ',' + K + ',B1,' + d + ',">"&EOMONTH(B2,-2),' + d + ',"<="&EOMONTH(B2,-1))';
      const month0 = 'SUMIFS(' + s + ',' + d + ',">="&B2,' + d + ',"<="&EOMONTH(B2,0))';
      const common = { file: 'penjualan_q4.xlsx', sheet: S, base, alt, answerCell: 'B3', dateCells: ['B2'], title: V.title };
      if (v === 0) {
        return H.problem(Object.assign(common, {
          precision: 4,
          statement: '<p>B2 holds the first day of a month. In <b>B3</b>, work out how much the <b>Net</b> of the region in B1 grew from the month before to the month in B2: this month\'s net divided by last month\'s net, minus 1, as a decimal (0.25 is 25% growth; the answer is negative when net fell). Keep the decimals.</p>',
          solution: '=' + cur + '/' + prev + '-1',
          alts: ['=' + cur2 + '/' + prev2 + '-1', '=(' + cur + '-' + prev + ')/' + prev],
          hints: ['Each month is a date window: from B2 to EOMONTH(B2,0) for this month, and from EDATE(B2,-1) to the day before B2 for last month.', 'Write one SUMIFS per month, each with the region pair and two date pairs.', 'Shape: =SUMIFS(net,region,B1,date,">="&B2,date,"<="&EOMONTH(B2,0))/SUMIFS(net,region,B1,date,">="&EDATE(B2,-1),date,"<"&B2)-1.'],
          explain: 'EDATE(B2,-1) is the first day of the month before and EOMONTH(B2,0) the last day of this month, so the two SUMIFS give the region\'s net in each window. New over old minus 1 is the growth rate.',
          chips: [s, K, d, '">="&B2', '"<="&EOMONTH(B2,0)', '">="&EDATE(B2,-1)', '"<"&B2'], fnKey: 'SUMIFS('
        }));
      }
      if (v === 1) {
        return H.problem(Object.assign(common, {
          statement: '<p>B2 holds the first day of a month. In <b>B3</b>, work out how many more <b>Cartons</b> the distributor in B1 bought in the month of B2 than in the month before it. The answer is negative when it bought fewer.</p>',
          solution: '=' + cur + '-' + prev,
          alts: ['=' + cur2 + '-' + prev2, '=-(' + prev + '-' + cur + ')'],
          hints: ['Each month is a window of dates: this month runs from B2 to EOMONTH(B2,0).', 'Last month runs from EDATE(B2,-1) up to, but not including, B2.', 'Shape: =SUMIFS(cartons,distributor,B1,date,">="&B2,date,"<="&EOMONTH(B2,0))-SUMIFS(cartons,distributor,B1,date,">="&EDATE(B2,-1),date,"<"&B2).'],
          explain: 'Two SUMIFS with the same distributor and different date windows give the cartons of each month, and their difference is the change. EDATE and EOMONTH keep the windows right for months of any length.',
          chips: [s, K, d, '">="&B2', '"<="&EOMONTH(B2,0)', '">="&EDATE(B2,-1)', '"<"&B2'], fnKey: 'SUMIFS('
        }));
      }
      return H.problem(Object.assign(common, {
        precision: 4,
        statement: '<p>B2 holds the first day of a month. In <b>B3</b>, work out what share of that month\'s <b>Net</b> comes from the category in B1: the category\'s net in the month divided by the net of every line in the same month, as a decimal.</p>',
        solution: '=' + cur + '/' + month0,
        alts: ['=' + cur2 + '/SUMIFS(' + s + ',' + d + ',">="&B2,' + d + ',"<"&EDATE(B2,1))', '=SUMPRODUCT((' + K + '=B1)*(' + d + '>=B2)*(' + d + '<=EOMONTH(B2,0))*' + s + ')/' + month0],
        hints: ['The month window runs from B2 to EOMONTH(B2,0).', 'The top of the fraction has the category condition; the bottom has only the date window.', 'Shape: =SUMIFS(net,category,B1,date,">="&B2,date,"<="&EOMONTH(B2,0))/SUMIFS(net,date,">="&B2,date,"<="&EOMONTH(B2,0)).'],
        explain: 'Both SUMIFS use the same month window, and only the top one also keeps the category in B1, so the ratio is the category\'s share of that month. A share always sits between 0 and 1.',
        chips: [s, K, d, '">="&B2', '"<="&EOMONTH(B2,0)'], fnKey: 'SUMIFS('
      }));
    }
  });
});
