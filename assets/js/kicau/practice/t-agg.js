/*
 * Practice, Conditional totals (20 Easy, 40 Medium, 10 Hard): SUMIF(S), COUNTIF(S), AVERAGEIF(S), MAXIFS, MINIFS,
 * criteria from a cell, not equal, wildcards, date windows and blanks. The Hard ones work across several workbooks, as
 * real reports do: two-way grids, either-or totals and month or period windows whose names, rates, targets, limits
 * and dates are looked up in other open files before the total and the math.
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

  /*
   * Every Hard problem works across several workbooks, as a real report does: the learner's own file holds the main
   * data, other open workbooks hold the master, the rates, the targets, the limits or the calendar. The formula looks
   * them up first, then the conditional total and the math give the report figure. The hidden sheet changes the
   * other workbooks too (a revised target, a new rate, a moved cut-off), so a typed or copied number fails.
   */
  const MASTER = 'master_distributor.xlsx';
  /* The distributor master, the same in every problem that opens it: ID, name, city, region. */
  const masterDist = () => H.table([{ h: 'ID', k: 'id' }, { h: 'Distributor', k: 'name' }, { h: 'City', k: 'city' }, { h: 'Region', k: 'region' }], H.db().distributors);
  const round5m = (x) => Math.max(1, Math.round(x / 5000000)) * 5000000;
  const kolOf = (d) => (d === 0 ? 1 : d <= 90 ? 2 : d <= 120 ? 3 : d <= 180 ? 4 : 5);
  const book = (file, name, S, alt) => ({ file, sheets: [{ name, sheet: S, base: S.cells, alt: alt || S.cells }] });
  const distOf = (name) => H.db().distributors.find((d) => d.name === name);
  const uniq = (xs) => xs.filter((x, i) => xs.indexOf(x) === i);

  P.add({
    key: 'two-way-summary-table', difficulty: 'Hard', level: 8, topics: ['SUMIFS', 'INDEX', 'MATCH', 'XLOOKUP', 'Several workbooks'], count: 4, multi: true,
    make(r, v) {
      const MD = masterDist(), md = H.ext(MASTER, 'Distributor', MD);
      if (v === 0) {
        /* A summary grid (distributor and category down, months across), the master for names, a target grid by ID. */
        const QM = ['Oct', 'Nov', 'Dec'];
        const qm = (d) => QM[month(H.Q4(d)) - 10];
        const pool = H.db().sales;
        const ds = r.sample(H.db().distributors, 4);
        const rows = [];
        ds.forEach((d) => H.CATEGORIES.forEach((cat) => {
          const own = pool.filter((s) => s.distributor === d.name && s.category === cat);
          const o = { distributor: d.name, category: cat };
          QM.forEach((m) => { o[m] = own.filter((s) => qm(s.date) === m).reduce((a, s) => a + s.net, 0); });
          rows.push(o);
        }));
        const S = H.table([COL.dist, COL.category].concat(QM.map((m) => ({ h: m, k: m, num: true }))), rows);
        const M = '$C$' + S.first + ':$E$' + S.last, Hd = '$C$4:$E$4', A = S.A('Distributor');
        const actual = (name, m) => pool.filter((s) => s.distributor === name && qm(s.date) === m).reduce((a, s) => a + s.net, 0);
        const TG = H.table([{ h: 'ID', k: 'id' }].concat(QM.map((m) => ({ h: m, k: m, num: true }))), H.db().distributors.map((d) => {
          const o = { id: d.id };
          QM.forEach((m) => { o[m] = round5m(actual(d.name, m) * r.pick([0.8, 0.9, 0.95, 1.05, 1.15, 1.25])); });
          return o;
        }));
        const tg = H.ext('target_bulanan.xlsx', 'Target', TG), grid = tg.T('Oct', 'Dec'), tgHd = tg.prefix + '$B$4:$D$4', tgHdAll = tg.prefix + '$A$4:$D$4';
        const ms = r.shuffle(QM);
        const base = Object.assign({}, S.cells, { A1: 'Distributor ID', B1: ds[0].id, A2: 'Month', B2: ms[0], A3: 'Achievement' });
        /* The hidden sheet: another ID and month, one figure of the grid corrected, and that target revised. */
        const j = rows.findIndex((x) => x.distributor === ds[1].name && x[ms[1]] > 0);
        const alt = H.edit(base, { B1: ds[1].id, B2: ms[1], [S.cell(ms[1], j < 0 ? 4 : j)]: (x) => x + 3000000 });
        const ti = H.db().distributors.findIndex((d) => d.id === ds[1].id);
        const tgAlt = H.edit(TG.cells, { [TG.cell(ms[1], ti)]: (x) => x + 10000000 });
        const name = 'XLOOKUP(B1,' + md.A('ID') + ',' + md.A('Distributor') + ')';
        return H.problem({
          title: 'Month net of the distributor ID in B1 against its monthly target', file: 'ringkasan_q4.xlsx', sheetName: 'Ringkasan', sheet: S, base, alt, answerCell: 'B3', precision: 4,
          books: [{ file: MASTER, sheets: [{ name: 'Distributor', sheet: MD }] }, book('target_bulanan.xlsx', 'Target', TG, tgAlt)],
          statement: '<p>Three workbooks are open: your Q4 summary (<b>ringkasan_q4.xlsx</b>, one row per distributor and category, one column per month in C4:E4), the distributor master (<b>master_distributor.xlsx</b>, ID and name) and the monthly targets (<b>target_bulanan.xlsx</b>, one row per distributor ID, one column per month).</p><p>B1 holds a distributor ID and B2 a month. In <b>B3</b>, work out the distributor\'s achievement for that month: its net in the month, all categories together, divided by its target for the same month. The summary lists names, not IDs, so find the name in the master first. Leave it as a decimal.</p>',
          solution: '=SUMIFS(XLOOKUP(B2,' + Hd + ',' + M + '),' + A + ',' + name + ')/INDEX(' + grid + ',MATCH(B1,' + tg.A('ID') + ',0),MATCH(B2,' + tgHd + ',0))',
          alts: ['=SUMIFS(INDEX(' + M + ',0,MATCH(B2,' + Hd + ',0)),' + A + ',INDEX(' + md.A('Distributor') + ',MATCH(B1,' + md.A('ID') + ',0)))/XLOOKUP(B1,' + tg.A('ID') + ',XLOOKUP(B2,' + tgHd + ',' + grid + '))',
            '=SUMIF(' + A + ',VLOOKUP(B1,' + md.T('ID', 'Distributor') + ',2,FALSE),INDEX(' + M + ',0,MATCH(B2,' + Hd + ',0)))/VLOOKUP(B1,' + tg.T('ID', 'Dec') + ',MATCH(B2,' + tgHdAll + ',0),FALSE)'],
          hints: ['Three lookups and one total. The name: ' + name + '. The month column of your grid: XLOOKUP(B2,' + Hd + ',' + M + '). The target: where the ID\'s row meets the month\'s column in target_bulanan.xlsx.',
            'SUMIFS adds the month column for the rows whose Distributor is that name. For the target, INDEX(target grid, MATCH(B1, ID column, 0), MATCH(B2, month headers, 0)).',
            'Shape: =SUMIFS(XLOOKUP(B2,' + Hd + ',' + M + '),' + A + ',name from the master)/INDEX(' + grid + ',MATCH(B1,...),MATCH(B2,...)).'],
          explain: 'The master turns the ID into the name the summary uses, and XLOOKUP on the month headers hands SUMIFS the one month column to add. In the targets file, two MATCHes find the ID\'s row and the month\'s column, and INDEX returns the target where they meet. Net over target is the achievement.',
          chips: [M, Hd, A, md.A('ID'), md.A('Distributor'), grid, tg.A('ID'), tgHd], fnKey: 'SUMIFS('
        });
      }
      if (v === 1) {
        /* The loan book, and a provision grid in another workbook: kol down, loan type across. */
        const rows = H.loans(r, 16).map((x, i) => Object.assign({}, x, { type: i % 3 === 1 ? 'Term Loan' : 'Working Capital', kol: kolOf(x.dpd) }));
        const S = H.table([COL.facility, COL.dist, COL.type, COL.outstanding, { h: 'Kol', k: 'kol', num: true }], rows);
        const RATE = { 'Working Capital': [0.01, 0.05, 0.15, 0.5, 1], 'Term Loan': [0.015, 0.06, 0.2, 0.5, 1] };
        const CK = H.table([{ h: 'Kol', k: 'kol', num: true }, { h: 'Working Capital', k: 'wc', num: true }, { h: 'Term Loan', k: 'tl', num: true }],
          [1, 2, 3, 4, 5].map((k) => ({ kol: k, wc: RATE['Working Capital'][k - 1], tl: RATE['Term Loan'][k - 1] })));
        const ck = H.ext('tarif_ckpn.xlsx', 'Ckpn', CK), grid = ck.T('Working Capital', 'Term Loan'), ckHd = ck.prefix + '$B$4:$C$4', ckHdAll = ck.prefix + '$A$4:$C$4';
        const o = S.R('Outstanding (Rp)'), t = S.R('Type'), k = S.R('Kol');
        const combos = ranked(rows.filter((x) => x.kol < 5), (x) => x.kol + '|' + x.type).map((s) => { const p = s.split('|'); return [+p[0], p[1]]; });
        const hp = hiddenPair(rows, 'kol', 'type', combos, 'outstanding', 5000000, total);
        const base = Object.assign({}, S.cells, { A1: 'Kol', B1: combos[0][0], A2: 'Loan type', B2: combos[0][1], A3: 'Provision (Rp)' });
        const alt = H.edit(base, { B1: hp.pair[0], B2: hp.pair[1], [S.cell('Outstanding (Rp)', hp.j)]: (x) => x + 5000000 });
        /* The hidden sheet: the bank revises the rate of the kol and type it asks about. */
        const ckAlt = H.edit(CK.cells, { [CK.cell(hp.pair[1], hp.pair[0] - 1)]: (x) => Math.round((x + 0.02) * 1000) / 1000 });
        const rate = 'INDEX(' + grid + ',MATCH(B1,' + ck.A('Kol') + ',0),MATCH(B2,' + ckHd + ',0))';
        return H.problem({
          title: 'Provision for the kol in B1 and loan type in B2 from the rate grid', file: 'buku_pinjaman.xlsx', sheetName: 'Fasilitas', sheet: S, base, alt, answerCell: 'B3',
          books: [book('tarif_ckpn.xlsx', 'Ckpn', CK, ckAlt)],
          statement: '<p>Two workbooks are open: the loan book (<b>buku_pinjaman.xlsx</b>, your file) and the provision rates (<b>tarif_ckpn.xlsx</b>: one row per kol from 1 to 5, one column per loan type, each rate a decimal, so 0.05 is 5%).</p><p>In <b>B3</b>, work out the provision for the facilities whose <b>Kol</b> is the one in B1 and whose <b>Type</b> is the one in B2: their total Outstanding times the rate where that kol\'s row meets that type\'s column in the rate grid.</p>',
          solution: '=SUMIFS(' + o + ',' + k + ',B1,' + t + ',B2)*' + rate,
          alts: ['=SUMIFS(' + o + ',' + t + ',B2,' + k + ',B1)*XLOOKUP(B1,' + ck.A('Kol') + ',XLOOKUP(B2,' + ckHd + ',' + grid + '))',
            '=SUMIFS(' + o + ',' + k + ',B1,' + t + ',B2)*VLOOKUP(B1,' + ck.T('Kol', 'Term Loan') + ',MATCH(B2,' + ckHdAll + ',0),FALSE)',
            '=SUMPRODUCT((' + k + '=B1)*(' + t + '=B2)*' + o + ')*' + rate],
          hints: ['Two parts: the balance (a SUMIFS on your sheet with two conditions) and the rate (a lookup in tarif_ckpn.xlsx that needs both the row and the column).',
            'MATCH(B1,' + ck.A('Kol') + ',0) finds the kol\'s row and MATCH(B2,' + ckHd + ',0) the type\'s column; INDEX(' + grid + ', row, column) returns the rate.',
            'Shape: =SUMIFS(' + o + ',' + k + ',B1,' + t + ',B2)*' + rate + '.'],
          explain: 'SUMIFS adds the outstanding of the facilities with that kol and type. INDEX with two MATCHes reads the rate at the crossing of the kol\'s row and the type\'s column in the other workbook. The balance times the rate is the provision, and a revised rate in tarif_ckpn.xlsx flows straight into the answer.',
          chips: [o, k, t, grid, ck.A('Kol'), ckHd], fnKey: 'SUMIFS('
        });
      }
      if (v === 2) {
        /* A stock grid (SKU down, warehouse across), the map of which warehouse serves each region, weekly demand. */
        const rows = H.PRODUCTS.map((p) => { const o = { sku: p.sku, product: p.name, category: p.category }; H.WAREHOUSES.forEach((w) => { o[w] = r.int(2, 40) * 10; }); return o; });
        const S = H.table([COL.sku, COL.product, COL.category].concat(H.WAREHOUSES.map((w) => ({ h: w, k: w, num: true }))), rows);
        const M = '$D$' + S.first + ':$H$' + S.last, Hd = '$D$4:$H$4', C = S.A('Category');
        const SERVE = { Jabodetabek: 'Gudang Bekasi', 'Jawa Barat': 'Gudang Bandung', 'Jawa Tengah': 'Gudang Semarang', 'Jawa Timur': 'Gudang Surabaya', 'Bali Nusra': 'Gudang Surabaya', Sumatera: 'Gudang Medan', Sulawesi: 'Gudang Surabaya' };
        const GD = H.table([{ h: 'Region', k: 'region' }, { h: 'Warehouse', k: 'wh' }], H.REGIONS.map((g) => ({ region: g, wh: SERVE[g] })));
        const DM = H.table([{ h: 'Warehouse', k: 'wh' }].concat(H.CATEGORIES.map((c) => ({ h: c, k: c, num: true }))),
          H.WAREHOUSES.map((w) => { const o = { wh: w }; H.CATEGORIES.forEach((c) => { o[c] = r.int(3, 24) * 5; }); return o; }));
        const gd = H.ext('peta_gudang.xlsx', 'Gudang', GD), dm = H.ext('permintaan.xlsx', 'Mingguan', DM);
        const grid = dm.T('Noodles', 'Soap'), dmHd = dm.prefix + '$B$4:$E$4', dmHdAll = dm.prefix + '$A$4:$E$4';
        const reg0 = r.pick(H.REGIONS), w0 = SERVE[reg0];
        const reg1 = r.pick(H.REGIONS.filter((g) => SERVE[g] !== w0));
        const w1 = r.pick(H.WAREHOUSES.filter((w) => w !== w0 && w !== SERVE[reg1]));
        const cats = r.shuffle(H.CATEGORIES);
        const base = Object.assign({}, S.cells, { A1: 'Region', B1: reg0, A2: 'Category', B2: cats[0], A3: 'Weeks of cover' });
        /* The hidden sheet: another region and category, that region now served from another warehouse, demand revised. */
        const sj = rows.findIndex((x) => x.category === cats[1]);
        const alt = H.edit(base, { B1: reg1, B2: cats[1], [S.cell(w1, sj)]: (x) => x + 30 });
        const gdAlt = H.edit(GD.cells, { [GD.cell('Warehouse', H.REGIONS.indexOf(reg1))]: w1 });
        const dmAlt = H.edit(DM.cells, { [DM.cell(cats[1], H.WAREHOUSES.indexOf(w1))]: (x) => x + 10 });
        const W = 'XLOOKUP(B1,' + gd.A('Region') + ',' + gd.A('Warehouse') + ')';
        const Wv = 'VLOOKUP(B1,' + gd.T('Region', 'Warehouse') + ',2,FALSE)';
        return H.problem({
          title: 'Weeks of stock cover for a region, warehouse map and demand apart', file: 'stok_gudang.xlsx', sheetName: 'Stok', sheet: S, base, alt, answerCell: 'B3', precision: 4,
          books: [book('peta_gudang.xlsx', 'Gudang', GD, gdAlt), book('permintaan.xlsx', 'Mingguan', DM, dmAlt)],
          statement: '<p>Three workbooks are open: your stock grid (<b>stok_gudang.xlsx</b>, one row per SKU, one column of on-hand cartons per warehouse in D4:H4), the warehouse map (<b>peta_gudang.xlsx</b>, which warehouse serves each region) and the weekly demand (<b>permintaan.xlsx</b>, one row per warehouse, one column per category, in cartons a week).</p><p>In <b>B3</b>, work out how many weeks the stock lasts: the on-hand cartons of the category in B2 in the warehouse that serves the region in B1, divided by that warehouse\'s weekly demand for the category. Leave it as a decimal.</p>',
          solution: '=SUMIFS(XLOOKUP(' + W + ',' + Hd + ',' + M + '),' + C + ',B2)/INDEX(' + grid + ',MATCH(' + W + ',' + dm.A('Warehouse') + ',0),MATCH(B2,' + dmHd + ',0))',
          alts: ['=LET(w,' + W + ',SUMIFS(XLOOKUP(w,' + Hd + ',' + M + '),' + C + ',B2)/XLOOKUP(w,' + dm.A('Warehouse') + ',XLOOKUP(B2,' + dmHd + ',' + grid + ')))',
            '=SUMIFS(INDEX(' + M + ',0,MATCH(' + Wv + ',' + Hd + ',0)),' + C + ',B2)/VLOOKUP(' + Wv + ',' + dm.T('Warehouse', 'Soap') + ',MATCH(B2,' + dmHdAll + ',0),FALSE)',
            '=SUMIF(' + C + ',B2,XLOOKUP(' + W + ',' + Hd + ',' + M + '))/INDEX(' + grid + ',MATCH(' + W + ',' + dm.A('Warehouse') + ',0),MATCH(B2,' + dmHd + ',0))'],
          hints: ['Start with the warehouse: ' + W + ' reads it from peta_gudang.xlsx. Everything else uses that name.',
            'The stock: XLOOKUP(warehouse,' + Hd + ',' + M + ') returns that warehouse\'s column, and SUMIFS adds the rows of the category in B2. The demand: INDEX(' + grid + ', MATCH(warehouse, ...), MATCH(B2, ...)) in permintaan.xlsx. LET(w, ..., ...) lets you write the warehouse lookup once.',
            'Shape: =SUMIFS(XLOOKUP(warehouse,' + Hd + ',' + M + '),' + C + ',B2)/INDEX(' + grid + ',MATCH(warehouse,' + dm.A('Warehouse') + ',0),MATCH(B2,' + dmHd + ',0)).'],
          explain: 'The warehouse lookup feeds two places: it picks the stock column in your grid and the demand row in permintaan.xlsx. SUMIFS adds the category\'s cartons in that column, and stock divided by weekly demand is the weeks of cover. When the map moves a region to another warehouse, the same formula follows it.',
          chips: [M, Hd, C, gd.A('Region'), gd.A('Warehouse'), grid, dm.A('Warehouse'), dmHd], fnKey: 'SUMIFS('
        });
      }
      /* v 3: sales lines without prices, the master for each distributor's region, a price list by SKU and region. */
      const regs = r.sample(H.REGIONS, 4);
      const dn = regs.map((g) => r.pick(H.db().distributors.filter((d) => d.region === g)).name);
      const rows = lines(r, 18, (s) => dn.includes(s.distributor));
      const S = H.table([COL.date, COL.dist, COL.sku, COL.cartons], rows);
      const b = S.R('Distributor'), k = S.R('SKU'), c = S.R('Cartons');
      const FACTOR = { Jabodetabek: 1, 'Jawa Barat': 0.98, 'Jawa Tengah': 0.96, 'Jawa Timur': 0.97, 'Bali Nusra': 1.04, Sumatera: 1.06, Sulawesi: 1.08 };
      const HG = H.table([COL.sku].concat(H.REGIONS.map((g) => ({ h: g, k: g, num: true }))),
        H.PRODUCTS.map((p) => { const o = { sku: p.sku }; H.REGIONS.forEach((g) => { o[g] = Math.round(p.carton_price * FACTOR[g] / 1000) * 1000; }); return o; }));
      const hg = H.ext('harga_regional.xlsx', 'Harga', HG), grid = hg.T(H.REGIONS[0], H.REGIONS[6]), hgHd = hg.prefix + '$B$4:$H$4', hgHdAll = hg.prefix + '$A$4:$H$4';
      const combos = ranked(rows, (x) => x.distributor + '|' + x.sku).map((s) => s.split('|'));
      const hp = hiddenPair(rows, 'distributor', 'sku', combos, 'cartons', 6, total);
      const base = Object.assign({}, S.cells, { A1: 'Distributor', B1: combos[0][0], A2: 'SKU', B2: combos[0][1], A3: 'List value (Rp)' });
      /* The hidden sheet: another distributor (in another region) and SKU, and that regional price goes up. */
      const alt = H.edit(base, { B1: hp.pair[0], B2: hp.pair[1], [S.cell('Cartons', hp.j)]: (x) => x + 6 });
      const hgAlt = H.edit(HG.cells, { [HG.cell(distOf(hp.pair[0]).region, H.PRODUCTS.findIndex((p) => p.sku === hp.pair[1]))]: (x) => x + 4000 });
      const RG = 'XLOOKUP(B1,' + md.A('Distributor') + ',' + md.A('Region') + ')';
      const total0 = 'SUMIFS(' + c + ',' + b + ',B1,' + k + ',B2)';
      return H.problem({
        title: 'Cartons of one SKU at the regional list price, from three files', file: 'penjualan_q4.xlsx', sheetName: 'Penjualan', sheet: S, base, alt, answerCell: 'B3',
        books: [{ file: MASTER, sheets: [{ name: 'Distributor', sheet: MD }] }, book('harga_regional.xlsx', 'Harga', HG, hgAlt)],
        statement: '<p>Three workbooks are open: your Q4 sales lines (<b>penjualan_q4.xlsx</b>, no prices), the distributor master (<b>master_distributor.xlsx</b>, with each distributor\'s region) and the regional price list (<b>harga_regional.xlsx</b>: one row per SKU, one column per region, the carton price in Rp).</p><p>In <b>B3</b>, work out what the distributor in B1 bought of the SKU in B2 at its list price: its total Cartons of that SKU times the carton price of that SKU in the distributor\'s own region.</p>',
        solution: '=' + total0 + '*INDEX(' + grid + ',MATCH(B2,' + hg.A('SKU') + ',0),MATCH(' + RG + ',' + hgHd + ',0))',
        alts: ['=' + total0 + '*XLOOKUP(B2,' + hg.A('SKU') + ',XLOOKUP(' + RG + ',' + hgHd + ',' + grid + '))',
          '=' + total0 + '*VLOOKUP(B2,' + hg.T('SKU', H.REGIONS[6]) + ',MATCH(VLOOKUP(B1,' + md.T('Distributor', 'Region') + ',3,FALSE),' + hgHdAll + ',0),FALSE)',
          '=SUMPRODUCT((' + b + '=B1)*(' + k + '=B2)*' + c + ')*INDEX(' + grid + ',MATCH(B2,' + hg.A('SKU') + ',0),MATCH(' + RG + ',' + hgHd + ',0))'],
        hints: ['The price depends on the SKU and on the region, and your sheet has no region: ' + RG + ' brings it from the master.',
          'In harga_regional.xlsx, MATCH(B2,' + hg.A('SKU') + ',0) gives the row and MATCH(region,' + hgHd + ',0) the column; INDEX(' + grid + ', row, column) is the price. The cartons come from a SUMIFS with two conditions.',
          'Shape: =' + total0 + '*INDEX(' + grid + ',MATCH(B2,...),MATCH(' + RG + ',' + hgHd + ',0)).'],
        explain: 'SUMIFS adds the cartons of that distributor and SKU on your sheet. The master gives the distributor\'s region, which picks the column of the price list, while the SKU picks the row. Cartons times that regional price is the list value, and a price change in harga_regional.xlsx moves the answer with it.',
        chips: [c, b, k, md.A('Distributor'), md.A('Region'), grid, hg.A('SKU'), hgHd], fnKey: 'SUMIFS('
      });
    }
  });

  P.add({
    key: 'either-or-totals', difficulty: 'Hard', level: 9, topics: ['SUMIFS', 'SUMPRODUCT', 'OR logic', 'XLOOKUP', 'Several workbooks'], count: 3, multi: true,
    make(r, v) {
      if (v === 0) {
        /* A rep covers two regions: net of either area, over the rep's target. Roster and targets are separate files. */
        const g = r.sample(H.REGIONS, 5);
        const rows = lines(r, 20, (s) => g.includes(s.region)), S = H.table([COL.date, COL.dist, COL.region, COL.category, COL.net], rows);
        const q = S.R('Region'), n = S.R('Net (Rp)');
        const pr = uniq(ranked(rows, 'region').concat(g));
        const has = (reg) => rows.some((x) => x.region === reg);
        const who = r.sample(H.REPS, 6);
        const roster = who.map((name, i) => {
          if (i === 0) return { name, a1: pr[0], a2: pr[2] };
          if (i === 1) return { name, a1: pr[1], a2: pr[3] };
          const two = r.sample(H.REGIONS, 2);
          return { name, a1: two[0], a2: two[1] };
        });
        const net = (rep) => sumOf(rows, (x) => x.region === rep.a1 || x.region === rep.a2, 'net');
        const tRows = r.shuffle(roster).map((rep) => ({ name: rep.name, target: round5m(net(rep) * r.pick([0.85, 0.95, 1.05, 1.15, 1.3])) }));
        const order = r.shuffle(roster);
        const RP = H.table([COL.rep, { h: 'Area 1', k: 'a1' }, { h: 'Area 2', k: 'a2' }], order);
        const TG = H.table([COL.rep, { h: 'Q4 target (Rp)', k: 'target', num: true }], tRows);
        const rp = H.ext('daftar_rep.xlsx', 'Rep', RP), tg = H.ext('target_rep.xlsx', 'Target', TG);
        const base = Object.assign({}, S.cells, { A1: 'Rep', B1: who[0], A2: 'Achievement' });
        /* The hidden sheet: another rep, who has been given a new second area, and a revised target. */
        const hi = order.findIndex((x) => x.name === who[1]);
        const newArea = has(pr[4]) ? pr[4] : pr[0];
        const rpAlt = H.edit(RP.cells, { [RP.cell('Area 2', hi)]: newArea });
        const tgAlt = H.edit(TG.cells, { [TG.cell('Q4 target (Rp)', tRows.findIndex((x) => x.name === who[1]))]: (x) => x + 10000000 });
        const j = rows.findIndex((x) => x.region === pr[1]);
        const alt = H.edit(base, { B1: who[1], [S.cell('Net (Rp)', j < 0 ? 0 : j)]: (x) => x + 2000000 });
        const a1 = 'XLOOKUP(B1,' + rp.A('Rep') + ',' + rp.A('Area 1') + ')', a2 = 'XLOOKUP(B1,' + rp.A('Rep') + ',' + rp.A('Area 2') + ')';
        const T = 'XLOOKUP(B1,' + tg.A('Rep') + ',' + tg.A('Q4 target (Rp)') + ')';
        return H.problem({
          title: 'Achievement of a rep who covers two areas, roster and targets apart', file: 'penjualan_q4.xlsx', sheetName: 'Penjualan', sheet: S, base, alt, answerCell: 'B2', precision: 4,
          books: [book('daftar_rep.xlsx', 'Rep', RP, rpAlt), book('target_rep.xlsx', 'Target', TG, tgAlt)],
          statement: '<p>Three workbooks are open: your Q4 sales lines (<b>penjualan_q4.xlsx</b>), the rep roster (<b>daftar_rep.xlsx</b>: every rep covers two regions, Area 1 and Area 2) and the Q4 targets (<b>target_rep.xlsx</b>, one target per rep).</p><p>In <b>B2</b>, work out the achievement of the rep in B1: the Net of the lines whose <b>Region</b> is <b>either</b> of the rep\'s two areas, divided by the rep\'s Q4 target. Leave it as a decimal.</p>',
          solution: '=(SUMIFS(' + n + ',' + q + ',' + a1 + ')+SUMIFS(' + n + ',' + q + ',' + a2 + '))/' + T,
          alts: ['=(SUMIFS(' + n + ',' + q + ',VLOOKUP(B1,' + rp.T('Rep', 'Area 2') + ',2,FALSE))+SUMIFS(' + n + ',' + q + ',VLOOKUP(B1,' + rp.T('Rep', 'Area 2') + ',3,FALSE)))/VLOOKUP(B1,' + tg.T('Rep', 'Q4 target (Rp)') + ',2,FALSE)',
            '=SUMPRODUCT(((' + q + '=' + a1 + ')+(' + q + '=' + a2 + '))*' + n + ')/' + T,
            '=(SUMIF(' + q + ',INDEX(' + rp.A('Area 1') + ',MATCH(B1,' + rp.A('Rep') + ',0)),' + n + ')+SUMIF(' + q + ',INDEX(' + rp.A('Area 2') + ',MATCH(B1,' + rp.A('Rep') + ',0)),' + n + '))/INDEX(' + tg.A('Q4 target (Rp)') + ',MATCH(B1,' + tg.A('Rep') + ',0))'],
          hints: ['Look up the rep\'s two areas in daftar_rep.xlsx and the target in target_rep.xlsx. The two files list the reps in different orders, so look each one up by name.',
            'SUMIFS joins conditions with AND, so either area needs its own SUMIFS; a line has one region, so the two totals never overlap and can be added.',
            'Shape: =(SUMIFS(' + n + ',' + q + ',' + a1 + ')+SUMIFS(' + n + ',' + q + ',' + a2 + '))/' + T + '.'],
          explain: 'Each XLOOKUP into the roster returns one area, which becomes the condition of its own SUMIFS. A line belongs to one region, so adding the two totals is the OR. The target comes from a third file, and net over target is the achievement. When the roster gives the rep a new area, the same formula follows it.',
          chips: [n, q, rp.A('Rep'), rp.A('Area 1'), rp.A('Area 2'), tg.A('Rep'), tg.A('Q4 target (Rp)')], fnKey: 'SUMIFS('
        });
      }
      if (v === 1) {
        /* Facilities for review: past the policy's DPD or past its use cap. The overlap counts once. */
        const TL = [1, 4, 6, 9, 12, 14];
        const rows = H.loans(r, 16).map((x, i) => ({ facility: x.facility, distributor: x.distributor, type: TL.includes(i) ? 'Term Loan' : 'Working Capital', limit: x.limit, outstanding: x.outstanding, dpd: x.dpd }));
        const set = (i, o) => { rows[i] = Object.assign({}, rows[i], o); };
        /* Working Capital, policy 60 days and 0.85 of the limit. */
        set(0, { limit: 100000000, outstanding: 92000000, dpd: 75 });
        set(2, { limit: 150000000, outstanding: 75000000, dpd: 95 });
        set(3, { limit: 200000000, outstanding: 196000000, dpd: 0 });
        set(5, { limit: 75000000, outstanding: 36000000, dpd: 60 });
        /* Term Loan, policy 30 days and 0.9; the hidden sheet's revised policy is 45 days and 0.8. */
        set(1, { limit: 150000000, outstanding: 127500000, dpd: 0 });
        set(4, { limit: 100000000, outstanding: 50000000, dpd: 45 });
        set(6, { limit: 200000000, outstanding: 190000000, dpd: 60 });
        set(9, { limit: 50000000, outstanding: 20000000, dpd: 95 });
        /* Only the revised 45 days catches this one, so the old Max DPD typed as a number fails. */
        set(14, { limit: 75000000, outstanding: 37500000, dpd: 50 });
        const S = H.table([COL.facility, COL.dist, COL.type, COL.limit, COL.outstanding, COL.dpd], rows);
        const t = S.R('Type'), l = S.R('Limit (Rp)'), o = S.R('Outstanding (Rp)'), d = S.R('DPD');
        const KB = H.table([{ h: 'Type', k: 'type' }, { h: 'Max DPD', k: 'dpd', num: true }, { h: 'Max use', k: 'use', num: true }],
          [{ type: 'Working Capital', dpd: 60, use: 0.85 }, { type: 'Term Loan', dpd: 30, use: 0.9 }, { type: 'Trade Finance', dpd: 45, use: 0.8 }]);
        const kb = H.ext('kebijakan_kredit.xlsx', 'Kebijakan', KB);
        const kbAlt = H.edit(KB.cells, { [KB.cell('Max DPD', 1)]: 45, [KB.cell('Max use', 1)]: 0.8 });
        const base = Object.assign({}, S.cells, { A1: 'Loan type', B1: 'Working Capital', A2: 'For review (Rp)' });
        const alt = H.edit(base, { B1: 'Term Loan', [S.cell('Outstanding (Rp)', 6)]: (x) => x + 5000000 });
        const X = 'XLOOKUP(B1,' + kb.A('Type') + ',' + kb.A('Max DPD') + ')', Y = 'XLOOKUP(B1,' + kb.A('Type') + ',' + kb.A('Max use') + ')';
        const Xv = 'VLOOKUP(B1,' + kb.T('Type', 'Max use') + ',2,FALSE)', Yv = 'VLOOKUP(B1,' + kb.T('Type', 'Max use') + ',3,FALSE)';
        const sol = '=SUMPRODUCT((' + t + '=B1)*(((' + d + '>' + X + ')+(' + o + '>' + l + '*' + Y + '))>0)*' + o + ')';
        return H.problem({
          title: 'Loans for review: past the DPD or the use cap in the policy file', file: 'buku_pinjaman.xlsx', sheetName: 'Fasilitas', sheet: S, base, alt, answerCell: 'B2',
          books: [book('kebijakan_kredit.xlsx', 'Kebijakan', KB, kbAlt)],
          statement: '<p>Two workbooks are open: the loan book (<b>buku_pinjaman.xlsx</b>, your file) and the credit policy (<b>kebijakan_kredit.xlsx</b>: for each loan type a <b>Max DPD</b> and a <b>Max use</b>, the highest outstanding as a share of the limit, as a decimal).</p><p>A facility of the type in B1 goes on Citra\'s review list when its <b>DPD</b> is <b>more than</b> the policy\'s Max DPD <b>or</b> its <b>Outstanding</b> is <b>more than</b> Max use times its <b>Limit</b>. In <b>B2</b>, add up the Outstanding of the facilities on the list, counting a facility that breaks both rules once.</p>',
          solution: sol,
          alts: ['=SUMIFS(' + o + ',' + t + ',B1,' + d + ',">"&' + X + ')+SUMPRODUCT((' + t + '=B1)*(' + d + '<=' + X + ')*(' + o + '>' + l + '*' + Y + ')*' + o + ')',
            '=SUMPRODUCT((' + t + '=B1)*(((' + d + '>' + Xv + ')+(' + o + '>' + l + '*' + Yv + '))>0)*' + o + ')',
            '=SUMPRODUCT((' + t + '=B1)*(' + d + '>' + X + ')*' + o + ')+SUMPRODUCT((' + t + '=B1)*(' + o + '>' + l + '*' + Y + ')*' + o + ')-SUMPRODUCT((' + t + '=B1)*(' + d + '>' + X + ')*(' + o + '>' + l + '*' + Y + ')*' + o + ')'],
          hints: ['Look up the two policy values for the type in B1 first: ' + X + ' and ' + Y + '.',
            'The use rule compares two columns row by row (Outstanding against Limit times Max use), which SUMIFS cannot do, so use SUMPRODUCT. In it, + between two tests is OR, and (test1+test2)>0 is 1 once, even when both pass.',
            'Shape: =SUMPRODUCT((' + t + '=B1)*(((' + d + '>Max DPD)+(' + o + '>' + l + '*Max use))>0)*' + o + ').'],
          explain: 'The two XLOOKUPs bring the type\'s Max DPD and Max use from the policy file. Inside SUMPRODUCT, each rule gives 1 or 0 per row; adding them and testing >0 keeps a row that breaks either rule, and counts it once when it breaks both. Multiplying by the type test and Outstanding gives the review total. When the policy is revised, the answer follows it.',
          chips: [o, l, d, t, kb.A('Type'), kb.A('Max DPD'), kb.A('Max use')], fnKey: 'SUMPRODUCT('
        });
      }
      /* v 2: a region's bonus pool: lines of the focus SKU or of at least the minimum cartons, rules in another file. */
      const g = r.sample(H.REGIONS, 3);
      const rows = lines(r, 21, (s) => g.includes(s.region)), S = H.table([COL.date, COL.region, COL.sku, COL.cartons, COL.net], rows);
      const q = S.R('Region'), k = S.R('SKU'), c = S.R('Cartons'), n = S.R('Net (Rp)');
      const regs = uniq(ranked(rows, 'region').concat(g));
      /* The rule for a region: its n-th most common SKU as the focus, and a minimum that one of its lines meets exactly. */
      const rule = (reg, nth) => {
        const own = rows.filter((x) => x.region === reg);
        const skus = ranked(own, 'sku');
        const focus = skus[nth] || skus[0] || H.PRODUCTS[0].sku;
        const cs = own.map((x) => x.cartons).sort((a, b) => a - b);
        const fmax = Math.max.apply(null, own.filter((x) => x.sku === focus).map((x) => x.cartons).concat([0]));
        const min = cs.length ? Math.min(cs[Math.floor(cs.length * 0.6)], fmax || cs[cs.length - 1]) : 20;
        return { focus, min };
      };
      const PB = H.table([{ h: 'Region', k: 'region' }, { h: 'Focus SKU', k: 'focus' }, { h: 'Min cartons', k: 'min', num: true }, { h: 'Bonus rate', k: 'rate', num: true }],
        H.REGIONS.map((reg) => {
          const ru = reg === regs[0] || reg === regs[1] ? rule(reg, 0) : { focus: r.pick(H.PRODUCTS).sku, min: r.int(4, 9) * 5 };
          return { region: reg, focus: ru.focus, min: ru.min, rate: r.pick([0.02, 0.025, 0.03, 0.04]) };
        }));
      const pb = H.ext('program_bonus.xlsx', 'Program', PB);
      /* The hidden sheet: another region, whose focus SKU, minimum and bonus rate are revised, and one of its lines corrected.
         The new rate differs from the visible region's, so a rate typed as a number fails. */
      const hi = H.REGIONS.indexOf(regs[1]), nr = rule(regs[1], 1);
      const rate0 = PB.cells[PB.cell('Bonus rate', H.REGIONS.indexOf(regs[0]))], rate1 = PB.cells[PB.cell('Bonus rate', hi)];
      const newRate = [0.005, 0.01, -0.005].map((dx) => Math.round((rate1 + dx) * 1000) / 1000).find((x) => x !== rate0 && x !== rate1);
      const pbAlt = H.edit(PB.cells, { [PB.cell('Focus SKU', hi)]: nr.focus, [PB.cell('Min cartons', hi)]: nr.min === rule(regs[1], 0).min ? nr.min + 1 : nr.min, [PB.cell('Bonus rate', hi)]: newRate });
      const j = rows.findIndex((x) => x.region === regs[1] && x.sku === nr.focus);
      const base = Object.assign({}, S.cells, { A1: 'Region', B1: regs[0], A2: 'Bonus pool (Rp)' });
      const alt = H.edit(base, { B1: regs[1], [S.cell('Net (Rp)', j < 0 ? 0 : j)]: (x) => x + 1500000 });
      const F = 'XLOOKUP(B1,' + pb.A('Region') + ',' + pb.A('Focus SKU') + ')', Mn = 'XLOOKUP(B1,' + pb.A('Region') + ',' + pb.A('Min cartons') + ')', R = 'XLOOKUP(B1,' + pb.A('Region') + ',' + pb.A('Bonus rate') + ')';
      const vl = (i) => 'VLOOKUP(B1,' + pb.T('Region', 'Bonus rate') + ',' + i + ',FALSE)';
      return H.problem({
        title: 'Region bonus pool: focus SKU or big lines, rules in another file', file: 'penjualan_q4.xlsx', sheetName: 'Penjualan', sheet: S, base, alt, answerCell: 'B2',
        books: [book('program_bonus.xlsx', 'Program', PB, pbAlt)],
        statement: '<p>Two workbooks are open: your Q4 sales lines (<b>penjualan_q4.xlsx</b>) and the Q4 bonus program (<b>program_bonus.xlsx</b>: for each region a <b>Focus SKU</b>, a <b>Min cartons</b> and a <b>Bonus rate</b> as a decimal).</p><p>A line of the region in B1 earns the bonus when its <b>SKU</b> is the region\'s focus SKU <b>or</b> its <b>Cartons</b> are <b>at least</b> the region\'s Min cartons. In <b>B2</b>, work out the region\'s bonus pool: the Net of the lines that earn it, each counted once, times the region\'s bonus rate.</p>',
        solution: '=(SUMIFS(' + n + ',' + q + ',B1,' + k + ',' + F + ')+SUMIFS(' + n + ',' + q + ',B1,' + k + ',"<>"&' + F + ',' + c + ',">="&' + Mn + '))*' + R,
        alts: ['=(SUMIFS(' + n + ',' + q + ',B1,' + k + ',' + F + ')+SUMIFS(' + n + ',' + q + ',B1,' + c + ',">="&' + Mn + ')-SUMIFS(' + n + ',' + q + ',B1,' + k + ',' + F + ',' + c + ',">="&' + Mn + '))*' + R,
          '=SUMPRODUCT((' + q + '=B1)*(((' + k + '=' + F + ')+(' + c + '>=' + Mn + '))>0)*' + n + ')*' + R,
          '=LET(f,' + F + ',m,' + Mn + ',(SUMIFS(' + n + ',' + q + ',B1,' + k + ',f)+SUMIFS(' + n + ',' + q + ',B1,' + k + ',"<>"&f,' + c + ',">="&m))*' + R + ')',
          '=(SUMIFS(' + n + ',' + q + ',B1,' + k + ',' + vl(2) + ')+SUMIFS(' + n + ',' + q + ',B1,' + k + ',"<>"&' + vl(2) + ',' + c + ',">="&' + vl(3) + '))*' + vl(4)],
        hints: ['Three lookups in program_bonus.xlsx, all by the region in B1: the focus SKU, the minimum cartons and the bonus rate.',
          'Split the OR so no line counts twice: the focus SKU lines, plus the lines that are not the focus SKU ("<>"&focus) but have at least the minimum (">="&min). Or add both groups and take away the lines in both.',
          'Shape: =(SUMIFS(' + n + ',' + q + ',B1,' + k + ',focus)+SUMIFS(' + n + ',' + q + ',B1,' + k + ',"<>"&focus,' + c + ',">="&min))*rate.'],
        explain: 'The lookups turn the region into its own rules, which become SUMIFS conditions. The focus SKU lines plus the big lines of other SKUs cover the OR without double counting, and the bonus rate turns that net into the pool. When the program changes the focus SKU, the minimum or the rate, the same formula follows it.',
        chips: [n, q, k, c, pb.A('Region'), pb.A('Focus SKU'), pb.A('Min cartons'), pb.A('Bonus rate')], fnKey: 'SUMIFS('
      });
    }
  });

  P.add({
    key: 'one-month-windows', difficulty: 'Hard', level: 8, topics: ['SUMIFS', 'EOMONTH', 'EDATE', 'XLOOKUP', 'Several workbooks'], count: 3, multi: true,
    make(r, v) {
      const MD = masterDist(), md = H.ext(MASTER, 'Distributor', MD);
      const first = (m) => H.toSerial('2026-' + m + '-01');
      const lastDay = (m) => (m === 12 ? H.toSerial('2027-01-01') : first(m + 1)) - 1;
      if (v === 0) {
        /* The closing calendar: periods run from the 26th to the 25th. Net in the period over the region's target. */
        const g = r.sample(H.REGIONS, 3);
        const rows = lines(r, 22, (s) => g.includes(s.region));
        const PER = ['Oct', 'Nov', 'Dec'], ps = r.shuffle(PER);
        const regs = uniq(ranked(rows, 'region').concat(g));
        const pm = (p) => 10 + PER.indexOf(p);
        const day = (m, dd) => '2026-' + String(m).padStart(2, '0') + '-' + String(dd).padStart(2, '0');
        /* Two lines of the visible region sit where the period and the calendar month disagree. */
        const a = rows.findIndex((x) => x.region === regs[0] && month(x.date) === pm(ps[0]));
        if (a >= 0) rows[a] = Object.assign({}, rows[a], { date: day(pm(ps[0]), 27) });
        const b = rows.findIndex((x, i) => i !== a && x.region === regs[0] && month(x.date) === pm(ps[0]) - 1);
        if (b >= 0) rows[b] = Object.assign({}, rows[b], { date: day(pm(ps[0]) - 1, 26) });
        rows.sort(byDate);
        const S = H.table([COL.date, COL.dist, COL.region, COL.net], rows);
        const d = S.R('Date'), q = S.R('Region'), n = S.R('Net (Rp)');
        const CAL = [['Oct', '2026-09-26', '2026-10-25'], ['Nov', '2026-10-26', '2026-11-25'], ['Dec', '2026-11-26', '2026-12-25']].map((x) => ({ p: x[0], s: x[1], e: x[2] }));
        const KT = H.table([{ h: 'Period', k: 'p' }, { h: 'Start', k: 's', date: true }, { h: 'End', k: 'e', date: true }], CAL);
        const inPer = (x, p) => { const c = CAL.find((y) => y.p === p); return x.date >= c.s && x.date <= c.e; };
        const TG = H.table([{ h: 'Region', k: 'region' }].concat(PER.map((p) => ({ h: p, k: p, num: true }))), H.REGIONS.map((reg) => {
          const o = { region: reg };
          PER.forEach((p) => { const act = sumOf(rows, (x) => x.region === reg && inPer(x, p), 'net'); o[p] = act ? round5m(act * r.pick([0.85, 0.95, 1.05, 1.15, 1.25])) : r.int(2, 8) * 5000000; });
          return o;
        }));
        const kt = H.ext('kalender_tutup.xlsx', 'Kalender', KT), tg = H.ext('target_region.xlsx', 'Target', TG);
        const grid = tg.T('Oct', 'Dec'), tgHd = tg.prefix + '$B$4:$D$4', tgHdAll = tg.prefix + '$A$4:$D$4';
        const base = Object.assign({}, S.cells, { A1: 'Region', B1: regs[0], A2: 'Period', B2: ps[0], A3: 'Achievement' });
        /* The hidden sheet: another region and period, whose closing moves two days later; a line is booked on the new last day. */
        const pi = PER.indexOf(ps[1]), m1 = pm(ps[1]);
        const ktEdits = { [KT.cell('End', pi)]: H.toSerial(day(m1, 27)) };
        if (pi < 2) ktEdits[KT.cell('Start', pi + 1)] = H.toSerial(day(m1, 28));
        const ktAlt = H.edit(KT.cells, ktEdits);
        const own = where(rows, (x) => x.region === regs[1] && month(x.date) === m1);
        const edits = { B1: regs[1], B2: ps[1] };
        if (own.length) edits[S.cell('Date', own[own.length - 1])] = H.toSerial(day(m1, 27));
        if (own.length > 1) edits[S.cell('Net (Rp)', own[0])] = (x) => x + 1500000;
        const alt = H.edit(base, edits);
        /* The revised target never lands on the visible one, so a target typed as a number fails. */
        const t0 = TG.cells[TG.cell(ps[0], H.REGIONS.indexOf(regs[0]))];
        const tgAlt = H.edit(TG.cells, { [TG.cell(ps[1], H.REGIONS.indexOf(regs[1]))]: (x) => x + (x + 10000000 === t0 ? 15000000 : 10000000) });
        const St ='XLOOKUP(B2,' + kt.A('Period') + ',' + kt.A('Start') + ')', En = 'XLOOKUP(B2,' + kt.A('Period') + ',' + kt.A('End') + ')';
        const target = 'INDEX(' + grid + ',MATCH(B1,' + tg.A('Region') + ',0),MATCH(B2,' + tgHd + ',0))';
        const net = 'SUMIFS(' + n + ',' + q + ',B1,' + d + ',">="&' + St + ',' + d + ',"<="&' + En + ')';
        return H.problem({
          title: 'Region net in a closing period against target, from three files', file: 'penjualan_q4.xlsx', sheetName: 'Penjualan', sheet: S, base, alt, answerCell: 'B3', precision: 4,
          books: [book('kalender_tutup.xlsx', 'Kalender', KT, ktAlt), book('target_region.xlsx', 'Target', TG, tgAlt)],
          statement: '<p>Three workbooks are open: your Q4 sales lines (<b>penjualan_q4.xlsx</b>), the closing calendar (<b>kalender_tutup.xlsx</b>: each period\'s Start and End date; a period usually runs from the 26th of one month to the 25th of the next, but the dates can move) and the period targets (<b>target_region.xlsx</b>, one row per region, one column per period).</p><p>In <b>B3</b>, work out the achievement of the region in B1 in the period named in B2: its Net from the period\'s Start to its End, both days included, divided by its target for that period. Leave it as a decimal.</p>',
          solution: '=' + net + '/' + target,
          alts: ['=SUMIFS(' + n + ',' + q + ',B1,' + d + ',">="&VLOOKUP(B2,' + kt.T('Period', 'End') + ',2,FALSE),' + d + ',"<="&VLOOKUP(B2,' + kt.T('Period', 'End') + ',3,FALSE))/VLOOKUP(B1,' + tg.T('Region', 'Dec') + ',MATCH(B2,' + tgHdAll + ',0),FALSE)',
            '=' + net + '/XLOOKUP(B1,' + tg.A('Region') + ',XLOOKUP(B2,' + tgHd + ',' + grid + '))',
            '=SUMPRODUCT((' + q + '=B1)*(' + d + '>=' + St + ')*(' + d + '<=' + En + ')*' + n + ')/' + target],
          hints: ['The period is not the calendar month: look up its first and last day in kalender_tutup.xlsx with ' + St + ' and ' + En + '.',
            'Those two dates become the date window of a SUMIFS (">="&start and "<="&end), next to the region condition. The target sits where the region\'s row meets the period\'s column: INDEX(' + grid + ', MATCH(B1, ...), MATCH(B2, ...)).',
            'Shape: =SUMIFS(' + n + ',' + q + ',B1,' + d + ',">="&start,' + d + ',"<="&end)/' + target + '.'],
          explain: 'The calendar lookups give the period\'s own Start and End, so SUMIFS adds the region\'s net between them. Lines dated after the 25th usually belong to the next period, which a calendar month window would get wrong. The target comes from the grid with INDEX and two MATCHes, and when the closing date moves, the same formula follows it.',
          chips: [n, q, d, kt.A('Period'), kt.A('Start'), kt.A('End'), grid, tg.A('Region'), tgHd], fnKey: 'SUMIFS('
        });
      }
      if (v === 1) {
        /* A report filled down: each distributor's net in the month on the sales sheet, over its credit limit (kept by ID). */
        const ds = r.sample(H.db().distributors, 8).sort((x, y) => (x.id < y.id ? -1 : 1));
        const dn = ds.map((x) => x.name);
        const sales = lines(r, 40, (s) => dn.includes(s.distributor));
        const PJ = H.table([COL.date, COL.dist, COL.net], sales);
        const S = H.table([COL.dist, COL.region], ds.map((x) => ({ distributor: x.name, region: x.region })));
        const lkRows = r.shuffle(H.db().distributors).map((x) => ({ id: x.id, limit: r.pick([25, 30, 40, 50, 60, 75]) * 1000000 }));
        const LK = H.table([{ h: 'ID', k: 'id' }, { h: 'Limit (Rp)', k: 'limit', num: true }], lkRows);
        const pj = H.ext(null, 'Penjualan', PJ), lk = H.ext('limit_kredit.xlsx', 'Limit', LK);
        const ms = r.shuffle([10, 11, 12]);
        const base = Object.assign({}, S.cells, { A1: 'Month (first day)', B1: first(ms[0]) });
        /* The hidden sheet: another month, two of its lines on its first and last day, one corrected, two limits revised. */
        const m1 = ms[1], own = where(sales, (x) => month(x.date) === m1);
        const pjEdits = {};
        if (own.length) pjEdits[PJ.cell('Date', own[0])] = first(m1);
        if (own.length > 1) pjEdits[PJ.cell('Date', own[own.length - 1])] = lastDay(m1);
        if (own.length > 2) pjEdits[PJ.cell('Net (Rp)', own[1])] = (x) => x + 2000000;
        const pjAlt = H.edit(PJ.cells, pjEdits);
        const lkEdits = {};
        uniq(own.map((i) => distOf(sales[i].distributor).id)).slice(0, 2).forEach((id, i) => { lkEdits[LK.cell('Limit (Rp)', lkRows.findIndex((x) => x.id === id))] = (x) => x + (i ? -5000000 : 10000000); });
        const lkAlt = H.edit(LK.cells, lkEdits);
        const alt = H.edit(base, { B1: first(m1) });
        const idOf = 'XLOOKUP(A5,' + md.A('Distributor') + ',' + md.A('ID') + ')';
        const lim = 'XLOOKUP(' + idOf + ',' + lk.A('ID') + ',' + lk.A('Limit (Rp)') + ')';
        const N = pj.A('Net (Rp)'), B = pj.A('Distributor'), D = pj.A('Date');
        const net = 'SUMIFS(' + N + ',' + B + ',A5,' + D + ',">="&$B$1,' + D + ',"<="&EOMONTH($B$1,0))';
        return H.problem({
          title: 'Credit use per distributor for the month in B1, limits kept by ID', file: 'laporan_kredit.xlsx', sheetName: 'Laporan', sheet: S, base, alt, fill: true, ansLabel: 'Credit use', precision: 4, dateCells: ['B1'],
          sheets: [{ name: 'Penjualan', sheet: PJ, base: PJ.cells, alt: pjAlt }],
          books: [{ file: MASTER, sheets: [{ name: 'Distributor', sheet: MD }] }, book('limit_kredit.xlsx', 'Limit', LK, lkAlt)],
          statement: '<p>Three workbooks are open. Your report, <b>laporan_kredit.xlsx</b>, has two sheets: <b>Laporan</b>, the list you fill in, and <b>Penjualan</b>, the Q4 sales lines. Bu Lestari\'s credit limits are in <b>limit_kredit.xlsx</b>, listed by distributor ID, and <b>master_distributor.xlsx</b> gives each distributor\'s ID.</p><p>B1 holds the first day of a month. In <b>C5</b>, work out the distributor\'s credit use for that month: its Net on the Penjualan sheet from B1 to the last day of the month, divided by its credit limit. Fill it down to <b>C' + S.last + '</b>. Leave it as a decimal.</p>',
          solution: '=' + net + '/' + lim,
          alts: ['=SUMIFS(' + N + ',' + B + ',A5,' + D + ',">="&$B$1,' + D + ',"<"&EDATE($B$1,1))/' + lim,
            '=' + net + '/VLOOKUP(INDEX(' + md.A('ID') + ',MATCH(A5,' + md.A('Distributor') + ',0)),' + lk.T('ID', 'Limit (Rp)') + ',2,FALSE)',
            '=' + net + '/INDEX(' + lk.A('Limit (Rp)') + ',MATCH(' + idOf + ',' + lk.A('ID') + ',0))'],
          hints: ['The net comes from the Penjualan sheet of this workbook: a SUMIFS on ' + N + ' for the distributor in A5 and the month that starts in B1. Pin $B$1 so it stays put while A5 moves down.',
            'The limit file knows IDs, not names. First turn the name into an ID with ' + idOf + ' (VLOOKUP cannot return a column to the left of the one it searches), then look that ID up in limit_kredit.xlsx.',
            'Shape: =SUMIFS(' + N + ',' + B + ',A5,' + D + ',">="&$B$1,' + D + ',"<="&EOMONTH($B$1,0))/XLOOKUP(ID of A5,' + lk.A('ID') + ',' + lk.A('Limit (Rp)') + ').'],
          explain: 'On every row SUMIFS adds that distributor\'s net between B1 and EOMONTH($B$1,0) on the Penjualan sheet. The master turns the name into its ID and a second XLOOKUP brings the limit for that ID, so the report joins three sources. The ranges in other sheets and files are pinned, A5 is relative, so each row works out its own distributor.',
          chips: [N, B, D, '">="&$B$1', '"<="&EOMONTH($B$1,0)', md.A('Distributor'), md.A('ID'), lk.A('ID'), lk.A('Limit (Rp)')], fnKey: 'SUMIFS('
        });
      }
      /* v 2: the change in a distributor's freight bill from last month: cartons by month, the master, the rate card. */
      const regs = r.sample(H.REGIONS, 4);
      const dn = regs.map((g) => r.pick(H.db().distributors.filter((x) => x.region === g)).name);
      const rows = lines(r, 20, (s) => dn.includes(s.distributor));
      const S = H.table([COL.date, COL.dist, COL.product, COL.cartons], rows);
      const d = S.R('Date'), b = S.R('Distributor'), c = S.R('Cartons');
      /* Distributor and month pairs where both the month and the month before have cartons, and they differ. */
      const ok = [];
      for (const key of ranked(rows, 'distributor')) for (const m of [11, 12]) {
        const cur = sumOf(rows, (x) => x.distributor === key && month(x.date) === m, 'cartons');
        const prev = sumOf(rows, (x) => x.distributor === key && month(x.date) === m - 1, 'cartons');
        if (cur > 0 && prev > 0 && cur !== prev) ok.push({ key, m });
      }
      const lineCount = (o) => Math.min(rows.filter((x) => x.distributor === o.key && month(x.date) === o.m).length, rows.filter((x) => x.distributor === o.key && month(x.date) === o.m - 1).length);
      ok.sort((x, y) => lineCount(y) - lineCount(x));
      const keys = ranked(rows, 'distributor');
      if (!ok.length) ok.push({ key: keys[0], m: 12 });
      const two = ok.find((x) => x.key !== ok[0].key && x.m !== ok[0].m) || ok.find((x) => x.key !== ok[0].key) || { key: keys.find((x) => x !== ok[0].key), m: ok[0].m === 12 ? 11 : 12 };
      const FREIGHT = { Jabodetabek: 8000, 'Jawa Barat': 9500, 'Jawa Tengah': 11000, 'Jawa Timur': 12000, 'Bali Nusra': 16000, Sumatera: 18500, Sulawesi: 21000 };
      const TK = H.table([{ h: 'Region', k: 'region' }, { h: 'Rate per carton (Rp)', k: 'rate', num: true }], H.REGIONS.map((g) => ({ region: g, rate: FREIGHT[g] + r.int(0, 4) * 250 })));
      const tk = H.ext('tarif_kirim.xlsx', 'Tarif', TK);
      const base = Object.assign({}, S.cells, { A1: 'Distributor ID', B1: distOf(ok[0].key).id, A2: 'Month (first day)', B2: first(ok[0].m), A3: 'Freight change (Rp)' });
      /* The hidden sheet: another distributor and month, some of its lines on the first and last day, a new freight rate. */
      const edits = { B1: distOf(two.key).id, B2: first(two.m) };
      const toEdges = (m) => {
        const own = where(rows, (x) => x.distributor === two.key && month(x.date) === m);
        if (own.length) edits[S.cell('Date', own[0])] = first(m);
        if (own.length > 1) edits[S.cell('Date', own[own.length - 1])] = lastDay(m);
        return own;
      };
      const mine = toEdges(two.m);
      toEdges(two.m - 1);
      if (mine.length) edits[S.cell('Cartons', mine[0])] = (x) => x + 5;
      const alt = H.edit(base, edits);
      const tkAlt = H.edit(TK.cells, { [TK.cell('Rate per carton (Rp)', H.REGIONS.indexOf(distOf(two.key).region))]: (x) => x + 1500 });
      const N = 'XLOOKUP(B1,' + md.A('ID') + ',' + md.A('Distributor') + ')';
      const RT = 'XLOOKUP(XLOOKUP(B1,' + md.A('ID') + ',' + md.A('Region') + '),' + tk.A('Region') + ',' + tk.A('Rate per carton (Rp)') + ')';
      const win = (who) => '(SUMIFS(' + c + ',' + b + ',' + who + ',' + d + ',">="&B2,' + d + ',"<="&EOMONTH(B2,0))-SUMIFS(' + c + ',' + b + ',' + who + ',' + d + ',">="&EDATE(B2,-1),' + d + ',"<"&B2))';
      const mv = 'VLOOKUP(B1,' + md.T('ID', 'Region') + ',';
      return H.problem({
        title: 'Change in freight cost from last month for the distributor ID in B1', file: 'penjualan_q4.xlsx', sheetName: 'Penjualan', sheet: S, base, alt, answerCell: 'B3', dateCells: ['B2'],
        books: [{ file: MASTER, sheets: [{ name: 'Distributor', sheet: MD }] }, book('tarif_kirim.xlsx', 'Tarif', TK, tkAlt)],
        statement: '<p>Three workbooks are open: your Q4 sales lines (<b>penjualan_q4.xlsx</b>), the distributor master (<b>master_distributor.xlsx</b>: ID, name, city and region) and the freight rate card (<b>tarif_kirim.xlsx</b>: the cost of shipping one carton to each region).</p><p>B1 holds a distributor ID and B2 the first day of a month. In <b>B3</b>, work out how much that distributor\'s freight bill changed from the month before: its Cartons in the month of B2 minus its Cartons in the month before, times the rate per carton for the distributor\'s region. The answer is negative when the bill fell.</p>',
        solution: '=' + win(N) + '*' + RT,
        alts: ['=(SUMIFS(' + c + ',' + b + ',' + N + ',' + d + ',">="&B2,' + d + ',"<"&EDATE(B2,1))-SUMIFS(' + c + ',' + b + ',' + N + ',' + d + ',">"&EOMONTH(B2,-2),' + d + ',"<="&EOMONTH(B2,-1)))*' + RT,
          '=LET(who,' + N + ',' + win('who') + '*' + RT + ')',
          '=' + win(mv + '2,FALSE)') + '*VLOOKUP(' + mv + '4,FALSE),' + tk.T('Region', 'Rate per carton (Rp)') + ',2,FALSE)'],
        hints: ['Your sheet has names and no regions. From the master take the name (' + N + ') for SUMIFS and the region for the rate card.',
          'Each month is a date window: this month runs from B2 to EOMONTH(B2,0), last month from EDATE(B2,-1) up to the day before B2. The rate is a lookup of a lookup: XLOOKUP(region from the master,' + tk.A('Region') + ',' + tk.A('Rate per carton (Rp)') + ').',
          'Shape: =(SUMIFS(cartons,distributor,name,date,">="&B2,date,"<="&EOMONTH(B2,0))-SUMIFS(cartons,distributor,name,date,">="&EDATE(B2,-1),date,"<"&B2))*rate.'],
        explain: 'The master turns the ID into the name for both SUMIFS and into the region for the rate card. The two SUMIFS give the cartons of each month, their difference is the change, and the rate per carton turns it into rupiah. When the rate card changes, the answer follows it.',
        chips: [c, b, d, md.A('ID'), md.A('Distributor'), md.A('Region'), tk.A('Region'), tk.A('Rate per carton (Rp)'), '">="&EDATE(B2,-1)', '"<="&EOMONTH(B2,0)'], fnKey: 'SUMIFS('
      });
    }
  });
});
