/*
 * Practice, Basics (Easy, 80 problems): arithmetic, SUM, AVERAGE, MIN, MAX, COUNT, COUNTA, ROUND, percentages and
 * $ references. Each template makes `count` problems; v (0 to count-1) picks the variant and r is its seeded random.
 * Levels: 1 is one function or one operation, 2 adds a second idea (COUNT against COUNTA, ROUND, a pinned input,
 * two functions together), 3 needs a $ reference that behaves when filled, or a ratio built from two totals.
 *
 * The app's grid shows a blank cell in a number or date column as text, so a column with gaps is never flagged num
 * or date: its numbers still show as numbers, and its dates are listed in dateCells.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) factory(require('./core.js'));
  else factory(root.KicauPractice);
})(typeof globalThis !== 'undefined' ? globalThis : this, function (P) {
  'use strict';
  const H = P.H;

  /* ---------- shared columns and helpers ---------- */
  const C = {
    date: { h: 'Date', k: 'date', date: true },
    dist: { h: 'Distributor', k: 'distributor' },
    region: { h: 'Region', k: 'region' },
    product: { h: 'Product', k: 'product' },
    category: { h: 'Category', k: 'category' },
    sku: { h: 'SKU', k: 'sku' },
    cartons: { h: 'Cartons', k: 'cartons', num: true },
    price: { h: 'Carton price (Rp)', k: 'price', num: true },
    packs: { h: 'Packs per carton', k: 'packs', num: true },
    gross: { h: 'Gross (Rp)', k: 'gross', num: true },
    net: { h: 'Net (Rp)', k: 'net', num: true },
    facility: { h: 'Facility', k: 'facility' },
    type: { h: 'Type', k: 'type' },
    limit: { h: 'Limit (Rp)', k: 'limit', num: true },
    outstanding: { h: 'Outstanding (Rp)', k: 'outstanding', num: true },
    dpd: { h: 'Days past due', k: 'dpd', num: true },
    rate: { h: 'Rate (% a year)', k: 'rate', num: true },
    rep: { h: 'Rep', k: 'name' },
    area: { h: 'Area', k: 'area' },
    target: { h: 'Target (Rp)', k: 'target', num: true },
    actual: { h: 'Actual (Rp)', k: 'actual', num: true },
    visits: { h: 'Visits', k: 'visits', num: true },
    month: { h: 'Month', k: 'month' },
    sales: { h: 'Sales (Rp)', k: 'sales', num: true },
    returns: { h: 'Returns (Rp)', k: 'returns', num: true },
    cost: { h: 'Cost (Rp)', k: 'cost', num: true },
    warehouse: { h: 'Warehouse', k: 'warehouse' },
    onhand: { h: 'On hand', k: 'onhand', num: true },
    reorder: { h: 'Reorder point', k: 'reorder', num: true },
    unitCost: { h: 'Cost per carton (Rp)', k: 'cost', num: true },
    invoice: { h: 'Invoice', k: 'ref' },
    amount: { h: 'Amount (Rp)', k: 'amount', num: true },
    /* Paid on has gaps (unpaid invoices), so it is not flagged date: see the note at the top. */
    paid: { h: 'Paid on', v: (x) => (x.paid ? H.toSerial(x.paid) : undefined) }
  };
  /* A cell by header and row index; a negative index counts back from the last row. */
  const at = (S, h, i) => S.col(h) + (i < 0 ? S.last + 1 + i : S.first + i);
  /* The hidden sheet, the same file a day later: [header, row index, change], the change a number to add or a function. */
  function later(base, S, changes) {
    const e = {};
    for (const [h, i, d] of changes) e[at(S, h, i)] = typeof d === 'function' ? d : (x) => Math.max(0, (+x || 0) + d);
    return H.edit(base, e);
  }
  /* One column's values, top to bottom, and the row indexes where a test holds. */
  const colVals = (S, cells, h) => { const out = []; for (let n = S.first; n <= S.last; n++) out.push(cells[S.col(h) + n]); return out; };
  const where = (S, cells, h, test) => colVals(S, cells, h).map((x, i) => (test(x) ? i : -1)).filter((i) => i >= 0);
  /* The refs in one column that hold a number, for dateCells. */
  const filledRefs = (S, cells, h) => where(S, cells, h, (x) => typeof x === 'number').map((i) => at(S, h, i));
  /* The made-up world's sales, added up: a total of `key` over the lines that pass `test`. */
  const dbSum = (test, key) => H.db().sales.reduce((a, s) => a + (test(s) ? s[key] : 0), 0);
  const priceList = (r, n) => r.sample(H.PRODUCTS, n).map((p) => ({ sku: p.sku, product: p.name, category: p.category, price: p.carton_price, packs: p.packs_per_carton }));
  /* Invoices with at least `blanks` unpaid and two paid, so a count is never trivially all or nothing. */
  function invoiceBatch(r, n, blanks) {
    const rows = H.invoices(r, n);
    let open = rows.filter((x) => !x.paid).length;
    for (let i = 1; open < blanks; i += 3) if (rows[i % n].paid) { rows[i % n].paid = ''; open++; }
    for (let i = 0; rows.filter((x) => x.paid).length < 2; i += 2) rows[i % n].paid = rows[i % n].date;
    return rows;
  }
  /* A payment date for an invoice that gets paid on the hidden sheet: its invoice date plus its terms. */
  const payDate = (row) => H.toSerial(row.date) + row.terms;
  /* Known $ mistakes, so the verdict names the mistake rather than the error value it causes. */
  const unpinB1 = (solution) => ({ code: 'unpinned-input', title: 'B1 slid down as you filled', plain: 'Without $ signs, B1 turns into B2, B3 and so on as the formula fills down. Write $B$1 so every row reads the same cell.', formula: solution.replace(/\$B\$1/g, 'B1') });
  const unpinStart = (solution) => ({ code: 'unpinned-start', title: 'The start of the range moved', plain: 'A running total needs its first cell pinned, as in $C$5:C5, so the range grows from the first row. Without the $ signs it slides down and covers one row only.', formula: solution.replace(/\$([A-Z])\$(\d+):/g, '$1$2:') });

  /* ======================= Level 1: one function on one column, one operation on each row ======================= */

  P.add({
    key: 'sum-a-column', difficulty: 'Easy', level: 1, topics: ['SUM'], count: 4,
    make(r, v) {
      /* Four different columns to total, each on its own sheet, so no two variants are the same problem. */
      const V = [
        { what: 'cartons shipped', h: 'Cartons', file: 'penjualan_q4.xlsx', rows: () => H.sales(r, 16), cols: [{ h: 'Date', k: 'date', date: true }, { h: 'Distributor', k: 'distributor' }, { h: 'Product', k: 'product' }, { h: 'Cartons', k: 'cartons', num: true }] },
        { what: 'net sales in rupiah', h: 'Net (Rp)', file: 'penjualan_q4.xlsx', rows: () => H.sales(r, 14), cols: [{ h: 'Date', k: 'date', date: true }, { h: 'Distributor', k: 'distributor' }, { h: 'Region', k: 'region' }, { h: 'Net (Rp)', k: 'net', num: true }] },
        { what: 'outstanding loans in rupiah', h: 'Outstanding (Rp)', file: 'buku_pinjaman.xlsx', rows: () => H.loans(r, 12), cols: [{ h: 'Facility', k: 'facility' }, { h: 'Distributor', k: 'distributor' }, { h: 'Limit (Rp)', k: 'limit', num: true }, { h: 'Outstanding (Rp)', k: 'outstanding', num: true }] },
        { what: 'cartons on hand', h: 'On hand', file: 'stok_gudang.xlsx', rows: () => H.stock(r, 12), cols: [{ h: 'SKU', k: 'sku' }, { h: 'Product', k: 'product' }, { h: 'Warehouse', k: 'warehouse' }, { h: 'On hand', k: 'onhand', num: true }] }
      ][v];
      const S = H.table(V.cols, V.rows());
      const top = { A1: 'Total ' + V.what.replace(/ in rupiah$/, '') };
      const base = Object.assign({}, S.cells, top);
      /* The hidden sheet: the same file a day later, with three figures corrected. */
      const c = S.col(V.h);
      const alt = H.edit(base, { [c + S.first]: (x) => x + (V.h === 'Cartons' || V.h === 'On hand' ? 7 : 2500000), [c + (S.first + 3)]: (x) => x + (V.h === 'Cartons' || V.h === 'On hand' ? 12 : 1200000), [c + S.last]: (x) => Math.max(0, x - (V.h === 'Cartons' || V.h === 'On hand' ? 5 : 900000)) });
      return H.problem({
        title: 'Total ' + V.what, file: V.file, sheet: S, base, alt, answerCell: 'B1',
        statement: '<p>In <b>B1</b>, add up every value in the <b>' + V.h + '</b> column (' + S.R(V.h) + ').</p>',
        solution: '=SUM(' + S.R(V.h) + ')', alts: ['=SUM(' + S.A(V.h) + ')'],
        hints: ['SUM adds every number in a range.', 'The range starts at the first data row and ends at the last: ' + S.R(V.h) + '.'],
        explain: 'SUM(' + S.R(V.h) + ') adds the whole column. Pointing at the range, not typing the numbers, keeps the total right when a figure changes.',
        chips: [S.R(V.h)], fnKey: 'SUM('
      });
    }
  });

  P.add({
    key: 'average-a-column', difficulty: 'Easy', level: 1, topics: ['AVERAGE'], count: 4,
    make(r, v) {
      const V = [
        () => ({ title: 'Average cartons per order line', label: 'Average cartons per line', h: 'Cartons', file: 'penjualan_q4.xlsx', ask: 'Pak Anton asks how many cartons one order line carries on average.', S: H.table([C.date, C.dist, C.product, C.cartons], H.sales(r, 15)), d: [[0, 6], [5, -2], [-1, 9]] }),
        () => ({ title: 'Average outstanding per loan facility', label: 'Average outstanding (Rp)', h: 'Outstanding (Rp)', file: 'buku_pinjaman.xlsx', ask: 'Dimas asks what one facility owes the bank on average.', S: H.table([C.facility, C.dist, C.type, C.outstanding], H.loans(r, 10)), d: [[1, 4000000], [4, -2500000], [-1, 3500000]] }),
        () => ({ title: 'Average days past due across the loan book', label: 'Average days past due', h: 'Days past due', file: 'buku_pinjaman.xlsx', ask: 'Citra asks how many days past due a facility is on average. Facilities at 0 days count too.', S: H.table([C.facility, C.dist, C.outstanding, C.dpd], H.loans(r, 12)), d: [[0, 30], [3, 15], [-2, 10]] }),
        () => ({ title: 'Average customer visits per sales rep', label: 'Average visits per rep', h: 'Visits', file: 'tim_sales.xlsx', ask: 'Pak Anton asks how many customer visits a rep made on average this month.', S: H.table([C.rep, C.area, C.target, C.visits], H.reps(r, 10)), d: [[0, 5], [2, -4], [-1, 8]] })
      ][v]();
      const S = V.S, R = S.R(V.h);
      const base = Object.assign({}, S.cells, { A1: V.label });
      const alt = later(base, S, V.d.map(([i, d]) => [V.h, i, d]));
      return H.problem({
        title: V.title, file: V.file, sheet: S, base, alt, answerCell: 'B1',
        statement: `<p>${V.ask}</p><p>In <b>B1</b>, work out the average of the <b>${V.h}</b> column (${R}).</p>`,
        solution: `=AVERAGE(${R})`, alts: [`=SUM(${R})/COUNT(${R})`, `=AVERAGE(${S.A(V.h)})`],
        hints: ['AVERAGE adds the numbers in a range and divides by how many numbers there are.', `Use the whole column of data, from the first row to the last: ${R}.`],
        explain: `AVERAGE(${R}) is the total of the column divided by the count of numbers in it, in one step. It points at the range, so it updates when a figure changes.`,
        chips: [R], fnKey: 'AVERAGE('
      });
    }
  });

  P.add({
    key: 'largest-value-in-a-column', difficulty: 'Easy', level: 1, topics: ['MAX'], count: 4,
    make(r, v) {
      const V = [
        () => ({ title: 'Largest single order in cartons', label: 'Largest order (cartons)', h: 'Cartons', file: 'penjualan_q4.xlsx', ask: 'What is the largest number of cartons on one order line?', S: H.table([C.date, C.dist, C.product, C.cartons], H.sales(r, 16)), up: 14, up2: 5 }),
        () => ({ title: 'Most days past due on one loan facility', label: 'Most days past due', h: 'Days past due', file: 'buku_pinjaman.xlsx', ask: 'Citra wants the facility furthest behind. Find the highest figure in days past due.', S: H.table([C.facility, C.dist, C.type, C.dpd], H.loans(r, 12)), up: 15, up2: 5 }),
        () => ({ title: 'Best month for sales in the year', label: 'Best month (Rp)', h: 'Sales (Rp)', file: 'penjualan_bulanan.xlsx', ask: 'What were the sales in the best month of the year?', S: H.table([C.month, C.sales, C.cost], H.monthly(r)), up: 25000000, up2: 8000000 }),
        () => ({ title: 'Biggest invoice amount in the batch', label: 'Biggest invoice (Rp)', h: 'Amount (Rp)', file: 'tagihan_q4.xlsx', ask: 'Bu Lestari wants the biggest invoice in this batch.', S: H.table([C.invoice, C.dist, C.date, C.amount], H.invoices(r, 14)), up: 1800000, up2: 600000 })
      ][v]();
      const S = V.S, R = S.R(V.h);
      const base = Object.assign({}, S.cells, { A1: V.label });
      const vs = colVals(S, base, V.h), top = vs.indexOf(Math.max(...vs)), other = top === 0 ? 1 : 0;
      /* The hidden sheet: the top figure is corrected upward, and another row changes as well (in round figures). */
      const alt = later(base, S, [[V.h, top, V.up], [V.h, other, V.up2]]);
      return H.problem({
        title: V.title, file: V.file, sheet: S, base, alt, answerCell: 'B1',
        statement: `<p>${V.ask}</p><p>In <b>B1</b>, return the largest value in <b>${V.h}</b> (${R}).</p>`,
        solution: `=MAX(${R})`, alts: [`=MAX(${S.A(V.h)})`, `=LARGE(${R},1)`],
        hints: ['MAX returns the largest number in a range.', `Point it at ${R}, the whole column.`],
        explain: `MAX(${R}) looks through every number in the column and returns the biggest. If a figure changes, the answer follows it.`,
        chips: [R], fnKey: 'MAX('
      });
    }
  });

  P.add({
    key: 'smallest-value-in-a-column', difficulty: 'Easy', level: 1, topics: ['MIN'], count: 4,
    make(r, v) {
      const V = [
        () => ({ title: 'Smallest single order in cartons', label: 'Smallest order (cartons)', h: 'Cartons', file: 'penjualan_q4.xlsx', ask: 'Pak Anton is looking at small drops. What is the smallest number of cartons on one line?', S: H.table([C.date, C.dist, C.product, C.cartons], H.sales(r, 16)), down: 1 }),
        () => ({ title: 'Weakest month for sales in the year', label: 'Weakest month (Rp)', h: 'Sales (Rp)', file: 'penjualan_bulanan.xlsx', ask: 'What were the sales in the weakest month of the year?', S: H.table([C.month, C.sales, C.returns], H.monthly(r)), down: 30000000 }),
        () => ({ title: 'Fewest customer visits by one sales rep', label: 'Fewest visits', h: 'Visits', file: 'tim_sales.xlsx', ask: 'Which rep made the fewest customer visits? Return that number of visits.', S: H.table([C.rep, C.area, C.actual, C.visits], H.reps(r, 10)), down: 3 }),
        () => ({ title: 'Cheapest carton in the price list', label: 'Cheapest carton (Rp)', h: 'Carton price (Rp)', file: 'daftar_harga.xlsx', ask: 'What does the cheapest carton on this price list cost?', S: H.table([C.sku, C.product, C.category, C.price], priceList(r, 9)), down: 6000 })
      ][v]();
      const S = V.S, R = S.R(V.h);
      const base = Object.assign({}, S.cells, { A1: V.label });
      const vs = colVals(S, base, V.h), low = vs.indexOf(Math.min(...vs)), other = low === 0 ? 1 : 0;
      /* The hidden sheet: the lowest figure is corrected downward (a promo price, a late return), and another row moves up. */
      const alt = later(base, S, [[V.h, low, -V.down], [V.h, other, V.down]]);
      return H.problem({
        title: V.title, file: V.file, sheet: S, base, alt, answerCell: 'B1',
        statement: `<p>${V.ask}</p><p>In <b>B1</b>, return the smallest value in <b>${V.h}</b> (${R}).</p>`,
        solution: `=MIN(${R})`, alts: [`=MIN(${S.A(V.h)})`, `=SMALL(${R},1)`],
        hints: ['MIN returns the smallest number in a range.', `Point it at ${R}, the whole column.`],
        explain: `MIN(${R}) returns the lowest number in the column. Typing the number you see would break as soon as a figure changes; MIN keeps up.`,
        chips: [R], fnKey: 'MIN('
      });
    }
  });

  P.add({
    key: 'multiply-two-columns', difficulty: 'Easy', level: 1, topics: ['Arithmetic'], count: 4,
    make(r, v) {
      const FREIGHT = { Jabodetabek: 1500, 'Jawa Barat': 2000, 'Jawa Tengah': 2500, 'Jawa Timur': 2500, 'Bali Nusra': 4000, Sumatera: 4500, Sulawesi: 5500 };
      const V = [
        () => ({ title: 'Gross value of each order line', ans: 'Gross (Rp)', file: 'penjualan_q4.xlsx', ask: 'Tika needs the gross value of every order line, before any discount', S: H.table([C.date, C.dist, C.cartons, C.price], H.sales(r, 14)), a: 'Cartons', b: 'Carton price (Rp)', d: [['Cartons', 1, 6], ['Carton price (Rp)', 4, 4000], ['Cartons', -1, 3]] }),
        () => ({ title: 'Packs shipped on each order line', ans: 'Packs', file: 'penjualan_q4.xlsx', ask: 'The warehouse counts in packs, not cartons. Work out the packs on every line', S: H.table([C.date, C.dist, C.product, C.cartons, C.packs], H.sales(r, 13)), a: 'Cartons', b: 'Packs per carton', d: [['Cartons', 0, 5], ['Cartons', 6, 2], ['Cartons', -1, 4]] }),
        () => ({ title: 'Value of the stock on hand on each line', ans: 'Stock value (Rp)', file: 'stok_gudang.xlsx', ask: 'Finance values the stock at cost. Work out the value of every stock line', S: H.table([C.sku, C.unitCost, C.warehouse, C.onhand], H.stock(r, 12)), a: 'On hand', b: 'Cost per carton (Rp)', d: [['On hand', 2, 30], ['On hand', 7, 20], ['Cost per carton (Rp)', 0, 3000]] }),
        () => ({ title: 'Freight cost of each delivery by region', ans: 'Freight (Rp)', file: 'pengiriman_q4.xlsx', ask: 'Each region has its own freight rate per carton. Work out the freight on every delivery', S: H.table([C.date, C.dist, C.region, { h: 'Freight per carton (Rp)', v: (x) => FREIGHT[x.region] || 3000, num: true }, C.cartons], H.sales(r, 14)), a: 'Cartons', b: 'Freight per carton (Rp)', d: [['Cartons', 2, 4], ['Cartons', 8, 7], ['Cartons', -1, 3]] })
      ][v]();
      const S = V.S, a = S.cell(V.a, 0), b = S.cell(V.b, 0), first = S.next + S.first, last = S.next + S.last;
      const alt = later(S.cells, S, V.d);
      return H.problem({
        title: V.title, file: V.file, sheet: S, base: S.cells, alt, fill: true, ansLabel: V.ans,
        statement: `<p>${V.ask}.</p><p>In <b>${first}</b>, multiply <b>${V.a}</b> by <b>${V.b}</b>, then fill it down to <b>${last}</b>.</p>`,
        solution: `=${a}*${b}`, alts: [`=${b}*${a}`, `=PRODUCT(${a},${b})`],
        hints: ['The * sign multiplies two cells.', `Start from the first data row: ${a} and ${b}. Leave out the $ signs, because both cells should move down a row as you fill.`],
        explain: `${a}*${b} works out row ${S.first}. Filling down moves both references one row at a time, so every line uses its own figures.`,
        chips: [a, b]
      });
    }
  });

  P.add({
    key: 'subtract-two-columns', difficulty: 'Easy', level: 1, topics: ['Arithmetic'], count: 4,
    make(r, v) {
      const V = [
        () => ({ title: 'Undrawn limit left on each facility', ans: 'Undrawn (Rp)', file: 'buku_pinjaman.xlsx', ask: 'Dimas wants to know how much can still be drawn on each facility (one distributor can have two): <b>Limit (Rp)</b> minus <b>Outstanding (Rp)</b>', S: H.table([C.facility, C.dist, C.limit, C.outstanding], H.loans(r, 12)), a: 'Limit (Rp)', b: 'Outstanding (Rp)', d: [['Outstanding (Rp)', 1, 3000000], ['Outstanding (Rp)', 5, -2000000], ['Limit (Rp)', -1, 25000000]] }),
        () => ({ title: 'Actual minus target for each sales rep', ans: 'Over or under (Rp)', file: 'tim_sales.xlsx', ask: 'Pak Anton wants each rep\'s result against target: <b>Actual (Rp)</b> minus <b>Target (Rp)</b>. A negative number means the rep is short', S: H.table([C.rep, C.area, C.target, C.actual], H.reps(r, 10)), a: 'Actual (Rp)', b: 'Target (Rp)', d: [['Actual (Rp)', 0, 4500000], ['Actual (Rp)', 4, -3000000], ['Target (Rp)', -1, 10000000]] }),
        () => ({ title: 'Sales after returns for each month', ans: 'After returns (Rp)', file: 'penjualan_bulanan.xlsx', ask: 'Take the returned goods off each month\'s sales: <b>Sales (Rp)</b> minus <b>Returns (Rp)</b>', S: H.table([C.month, C.sales, C.returns], H.monthly(r)), a: 'Sales (Rp)', b: 'Returns (Rp)', d: [['Returns (Rp)', 2, 2400000], ['Returns (Rp)', 7, 1100000], ['Sales (Rp)', -1, 18000000]] }),
        () => ({ title: 'Stock above the reorder point for each SKU', ans: 'Above reorder', file: 'stok_gudang.xlsx', ask: 'How far is each SKU above its reorder point? <b>On hand</b> minus <b>Reorder point</b>. A negative number means it is time to reorder', S: H.table([C.sku, C.product, C.onhand, C.reorder], H.stock(r, 12)), a: 'On hand', b: 'Reorder point', d: [['On hand', 1, 40], ['On hand', 6, 70], ['Reorder point', -1, 30]] })
      ][v]();
      const S = V.S, a = S.cell(V.a, 0), b = S.cell(V.b, 0), first = S.next + S.first, last = S.next + S.last;
      const alt = later(S.cells, S, V.d);
      return H.problem({
        title: V.title, file: V.file, sheet: S, base: S.cells, alt, fill: true, ansLabel: V.ans,
        statement: `<p>${V.ask}.</p><p>Type the formula in <b>${first}</b> and fill it down to <b>${last}</b>.</p>`,
        solution: `=${a}-${b}`, alts: [],
        hints: ['The - sign subtracts the second cell from the first.', `Order matters: start with ${V.a} (${a}), then take away ${V.b} (${b}).`],
        explain: `${a}-${b} gives the difference on row ${S.first}; filling down repeats it on every row with that row's own cells.`,
        chips: [a, b]
      });
    }
  });

  P.add({
    key: 'total-across-a-row', difficulty: 'Easy', level: 1, topics: ['SUM', 'Arithmetic'], count: 4,
    make(r, v) {
      const mo = (id, ym) => dbSum((s) => s.distributor_id === id && s.date.startsWith(ym), 'net');
      const V = [
        () => ({ title: 'Net sales per distributor over three months', ans: 'Three months (Rp)', file: 'penjualan_distributor.xlsx', ask: 'Tika wants each distributor\'s net sales for August, September and October together',
          S: H.table([{ h: 'Distributor', k: 'name' }, { h: 'Aug (Rp)', k: 'a', num: true }, { h: 'Sep (Rp)', k: 'b', num: true }, { h: 'Oct (Rp)', k: 'c', num: true }], r.sample(H.db().distributors, 10).map((d) => ({ name: d.name, a: mo(d.id, '2026-08'), b: mo(d.id, '2026-09'), c: mo(d.id, '2026-10') }))),
          d: [['Oct (Rp)', 0, 4200000], ['Sep (Rp)', 5, -1500000], ['Aug (Rp)', -1, 2600000]] }),
        () => {
          const W = r.sample(H.WAREHOUSES, 3);
          return { title: 'Stock across three warehouses for each product', ans: 'All three', file: 'stok_gudang.xlsx', ask: 'Find the total cartons of each product held across the three warehouses',
            S: H.table([{ h: 'Product', k: 'name' }].concat(W.map((w, j) => ({ h: w, k: 'w' + j, num: true }))), H.PRODUCTS.map((p) => ({ name: p.name, w0: r.int(0, 40) * 10, w1: r.int(0, 40) * 10, w2: r.int(0, 40) * 10 }))),
            d: [[W[0], 2, 60], [W[2], 6, 40], [W[1], -1, 30]] };
        },
        () => ({ title: 'Customer visits per rep over four weeks', ans: 'Month total', file: 'kunjungan_sales.xlsx', ask: 'Each rep logs customer visits by week. Add up the four weeks for every rep',
          S: H.table([{ h: 'Rep', k: 'name' }, { h: 'Week 1', k: 'w1', num: true }, { h: 'Week 2', k: 'w2', num: true }, { h: 'Week 3', k: 'w3', num: true }, { h: 'Week 4', k: 'w4', num: true }], r.sample(H.REPS, 10).map((name) => ({ name, w1: r.int(3, 16), w2: r.int(3, 16), w3: r.int(3, 16), w4: r.int(3, 16) }))),
          d: [['Week 4', 0, 3], ['Week 2', 4, 2], ['Week 1', -1, 4]] }),
        () => ({ title: 'Monthly cost from three cost lines', ans: 'Total cost (Rp)', file: 'biaya_bulanan.xlsx', ask: 'Finance keeps three cost lines. Add them up for the total cost of each month',
          S: H.table([C.month, { h: 'Materials (Rp)', k: 'm', num: true }, { h: 'Packaging (Rp)', k: 'p', num: true }, { h: 'Freight (Rp)', k: 'f', num: true }], H.MONTHS.map((month) => ({ month, m: r.int(380, 520) * 1000000, p: r.int(40, 75) * 1000000, f: r.int(25, 60) * 1000000 }))),
          d: [['Freight (Rp)', 3, 4000000], ['Materials (Rp)', 8, 12000000], ['Packaging (Rp)', -1, 2000000]] })
      ][v]();
      const S = V.S, from = S.cols[1] + S.first, to = S.cols[S.cols.length - 1] + S.first, first = S.next + S.first, last = S.next + S.last;
      const plus = S.cols.slice(1).map((c) => c + S.first).join('+');
      const alt = later(S.cells, S, V.d);
      return H.problem({
        title: V.title, file: V.file, sheet: S, base: S.cells, alt, fill: true, ansLabel: V.ans,
        statement: `<p>${V.ask}.</p><p>In <b>${first}</b>, add up the figures across the row, from <b>${from}</b> to <b>${to}</b>, then fill it down to <b>${last}</b>.</p>`,
        solution: `=SUM(${from}:${to})`, alts: [`=${plus}`],
        hints: ['SUM works across a row as well as down a column.', `The range for the first row is ${from}:${to}.`],
        explain: `SUM(${from}:${to}) adds the row. Filling down moves the range to the next row each time, so every line gets its own total.`,
        chips: [`${from}:${to}`], fnKey: 'SUM('
      });
    }
  });

  /* ======================= Level 2: COUNT against COUNTA, ROUND, a pinned input, two functions ======================= */

  P.add({
    key: 'difference-of-two-totals', difficulty: 'Easy', level: 2, topics: ['SUM', 'Arithmetic'], count: 4,
    make(r, v) {
      const V = [
        () => ({ title: 'Gross profit for the whole year', label: 'Gross profit (Rp)', file: 'penjualan_bulanan.xlsx', ask: 'What is the gross profit for the year: total <b>Sales (Rp)</b> minus total <b>Cost (Rp)</b>?', S: H.table([C.month, C.sales, C.cost], H.monthly(r)), a: 'Sales (Rp)', b: 'Cost (Rp)', d: [['Sales (Rp)', 3, 21000000], ['Cost (Rp)', 9, 6000000]] }),
        () => ({ title: 'Undrawn limit across the whole loan book', label: 'Undrawn limit (Rp)', file: 'buku_pinjaman.xlsx', ask: 'How much can the distributors still draw in all: total <b>Limit (Rp)</b> minus total <b>Outstanding (Rp)</b>?', S: H.table([C.facility, C.dist, C.limit, C.outstanding], H.loans(r, 12)), a: 'Limit (Rp)', b: 'Outstanding (Rp)', d: [['Outstanding (Rp)', 2, 5000000], ['Limit (Rp)', -1, 25000000]] }),
        () => ({ title: 'Team gap between total target and total actual', label: 'Team gap (Rp)', file: 'tim_sales.xlsx', ask: 'How far is the team from its target: total <b>Target (Rp)</b> minus total <b>Actual (Rp)</b>? A positive answer means the team is short.', S: H.table([C.rep, C.area, C.target, C.actual], H.reps(r, 10)), a: 'Target (Rp)', b: 'Actual (Rp)', d: [['Actual (Rp)', 1, 6500000], ['Actual (Rp)', 6, 2300000]] }),
        () => ({ title: 'Total discount given on the order lines', label: 'Discount given (Rp)', file: 'penjualan_q4.xlsx', ask: 'How much discount did the distributors get on these lines: total <b>Gross (Rp)</b> minus total <b>Net (Rp)</b>?', S: H.table([C.date, C.dist, C.gross, C.net], H.sales(r, 14)), a: 'Gross (Rp)', b: 'Net (Rp)', d: [['Net (Rp)', 2, -1800000], ['Net (Rp)', -2, -650000]] })
      ][v]();
      const S = V.S, Ra = S.R(V.a), Rb = S.R(V.b);
      const base = Object.assign({}, S.cells, { A1: V.label });
      const alt = later(base, S, V.d);
      return H.problem({
        title: V.title, file: V.file, sheet: S, base, alt, answerCell: 'B1',
        statement: `<p>${V.ask}</p><p>Type it in <b>B1</b>.</p>`,
        solution: `=SUM(${Ra})-SUM(${Rb})`, alts: [`=SUM(${Ra}-${Rb})`, `=SUMPRODUCT(${Ra}-${Rb})`],
        hints: ['Total each column with SUM, then subtract one total from the other.', `The two ranges are ${Ra} and ${Rb}.`, 'The shape: =SUM(first column)-SUM(second column).'],
        explain: `SUM(${Ra}) and SUM(${Rb}) are the two column totals; the - between them gives the difference in one cell.`,
        chips: [Ra, Rb], fnKey: 'SUM('
      });
    }
  });

  P.add({
    key: 'count-numbers-only', difficulty: 'Easy', level: 2, topics: ['COUNT'], count: 4,
    make(r, v) {
      const V = [
        () => {
          const rows = invoiceBatch(r, 14, 3);
          const S = H.table([C.invoice, C.dist, C.amount, C.paid], rows);
          const open = where(S, S.cells, 'Paid on', (x) => x === undefined);
          return { title: 'Count paid invoices from their payment dates', label: 'Paid invoices', file: 'tagihan_q4.xlsx', S, h: 'Paid on', skip: 'blanks',
            ask: 'Bu Lestari asks how many invoices in this batch are paid. A paid invoice has a date in <b>Paid on</b>; an unpaid one is blank.',
            hint: 'Excel stores a date as a number, so COUNT counts dates and skips the blanks.', d: [['Paid on', open[0], () => payDate(rows[open[0]])]],
            other: 'COUNTA gives the same count on this sheet, because the other cells are blank, but it would also count a note such as "unpaid" typed into the column.' };
        },
        () => {
          const rows = H.stock(r, 12).map((x) => Object.assign(x, { counted: r.chance(0.3) ? 'not counted' : x.onhand }));
          [2, 7].forEach((i) => { rows[i].counted = 'not counted'; });
          const S = H.table([C.sku, C.product, C.warehouse, { h: 'Counted', k: 'counted' }], rows);
          const todo = where(S, S.cells, 'Counted', (x) => typeof x !== 'number');
          return { title: 'Count stock lines that have a physical count', label: 'Lines counted', file: 'stock_opname.xlsx', S, h: 'Counted', skip: 'words',
            ask: 'The stock count is half done. How many lines already have a counted figure? Lines still waiting say <b>not counted</b>.',
            hint: 'COUNT skips words such as "not counted". COUNTA would count them, which is the wrong answer here.', d: [['Counted', todo[0], () => 140]] };
        },
        () => {
          const k = r.int(2, 4);
          const S = H.table([C.month, C.sales, { h: 'Returns (Rp)', v: (x, i) => (i >= 12 - k ? 'pending' : x.returns) }], H.monthly(r));
          return { title: 'Count months with returns already booked', label: 'Months booked', file: 'penjualan_bulanan.xlsx', S, h: 'Returns (Rp)', skip: 'words',
            ask: 'Finance books returns a few weeks late. How many months already have a returns figure? The rest say <b>pending</b>.',
            hint: 'A month that says "pending" holds a word, not a number, so COUNT leaves it out.', d: [['Returns (Rp)', 12 - k, () => 9500000]] };
        },
        () => {
          const rows = H.reps(r, 10).map((x) => Object.assign(x, { logged: r.chance(0.25) ? '-' : x.visits }));
          rows[1].logged = '-'; rows[6].logged = undefined;
          const S = H.table([C.rep, C.area, { h: 'Visits logged', k: 'logged' }], rows);
          const dash = where(S, S.cells, 'Visits logged', (x) => x === '-');
          return { title: 'Count reps who logged a visit figure', label: 'Reps with visits logged', file: 'tim_sales.xlsx', S, h: 'Visits logged', skip: 'dashes and blanks',
            ask: 'How many reps logged a number of visits? A rep who has not logged yet shows <b>-</b> or a blank.',
            hint: 'A dash is text to Excel, so COUNT skips it, and it skips blank cells too.', d: [['Visits logged', dash[0], () => 31]] };
        }
      ][v]();
      const S = V.S, R = S.R(V.h);
      const base = Object.assign({}, S.cells, { A1: V.label });
      const alt = later(base, S, V.d);
      return H.problem({
        title: V.title, file: V.file, sheet: S, base, alt, answerCell: 'B1', dateCells: V.h === 'Paid on' ? filledRefs(S, base, V.h) : undefined,
        statement: `<p>${V.ask}</p><p>In <b>B1</b>, count the cells in <b>${V.h}</b> (${R}) that hold ${V.h === 'Paid on' ? 'a date, which Excel stores as a number' : 'a number'}.</p>`,
        solution: `=COUNT(${R})`, alts: [`=COUNT(${S.A(V.h)})`].concat(V.other ? [`=COUNTA(${R})`] : []),
        hints: ['COUNT counts the cells that hold a number.', V.hint],
        explain: `COUNT(${R}) counts only the numbers in the column, so the ${V.skip} drop out. ` + (V.other || 'COUNTA counts every cell that is not empty, which is a different question.'),
        chips: [R], fnKey: 'COUNT('
      });
    }
  });

  P.add({
    key: 'count-filled-cells', difficulty: 'Easy', level: 2, topics: ['COUNTA', 'COUNT'], count: 4,
    make(r, v) {
      const REMARKS = ['Late delivery', 'Damaged carton', 'Price check', 'Short shipped', 'Paid in cash', 'Promo stock'];
      const V = [
        () => {
          const rows = H.sales(r, 14).map((x) => Object.assign(x, { remark: r.chance(0.35) ? r.pick(REMARKS) : undefined }));
          rows[3].remark = 'Damaged carton'; rows[9].remark = 'Late delivery'; rows[0].remark = undefined; rows[12].remark = undefined;
          const S = H.table([C.date, C.dist, C.cartons, { h: 'Remark', k: 'remark' }], rows);
          return { title: 'Count order lines that carry a remark', label: 'Lines with a remark', file: 'penjualan_q4.xlsx', S, h: 'Remark',
            ask: 'How many order lines have a remark written next to them?', hint: 'The remarks are words, and COUNT ignores words. COUNTA counts any cell that is not empty.',
            d: (S2) => [['Remark', where(S2, S2.cells, 'Remark', (x) => x === undefined)[0], () => 'Short shipped']] };
        },
        () => {
          const rows = H.sales(r, 14).map((x) => Object.assign(x, { cx: r.chance(0.2) ? r.pick(['on hold', 'cancelled']) : r.chance(0.15) ? undefined : x.cartons }));
          rows[2].cx = 'on hold'; rows[8].cx = undefined; rows[11].cx = 'cancelled';
          const S = H.table([C.date, C.product, { h: 'Cartons', k: 'cx' }], rows);
          return { title: 'Count lines with any entry in Cartons, words included', label: 'Lines with an entry', file: 'penjualan_q4.xlsx', S, h: 'Cartons',
            ask: 'How many lines have anything at all in <b>Cartons</b>, a number or a word such as <b>on hold</b>? Blank lines are not counted.', hint: 'COUNT would miss "on hold" and "cancelled". COUNTA counts numbers and words alike, and skips blanks.',
            d: (S2) => [['Cartons', where(S2, S2.cells, 'Cartons', (x) => x === undefined)[0], () => 'on hold']] };
        },
        () => {
          const rows = r.sample(H.db().outlets, 14).map((o) => ({ outlet: o.outlet, city: o.city, rep: r.chance(0.6) ? r.pick(H.REPS) : undefined }));
          rows[1].rep = undefined; rows[5].rep = undefined; rows[4].rep = rows[4].rep || 'Wulan'; rows[10].rep = rows[10].rep || 'Bima';
          const S = H.table([{ h: 'Outlet', k: 'outlet' }, { h: 'City', k: 'city' }, { h: 'Sales rep', k: 'rep' }], rows);
          return { title: 'Count outlets that have a sales rep assigned', label: 'Outlets with a rep', file: 'daftar_outlet.xlsx', S, h: 'Sales rep',
            ask: 'Pak Anton wants to know how many outlets already have a sales rep. An outlet with no rep has a blank in <b>Sales rep</b>.', hint: 'The reps are names, so this is COUNTA, not COUNT.',
            d: (S2) => [['Sales rep', where(S2, S2.cells, 'Sales rep', (x) => x === undefined)[0], () => 'Dewi']] };
        },
        () => {
          const rows = invoiceBatch(r, 14, 4);
          const S = H.table([C.invoice, C.dist, C.amount, C.paid], rows);
          return { title: 'Count unpaid invoices with COUNTA and COUNT', label: 'Unpaid invoices', file: 'tagihan_q4.xlsx', S, h: 'Paid on', unpaid: true,
            ask: `Bu Lestari asks how many invoices are still unpaid. Every invoice has a reference in <b>Invoice</b>, text such as ${rows[0].ref}; a paid one also has a date in <b>Paid on</b>.`, hint: 'COUNTA of the Invoice column gives every invoice, and COUNT of Paid on gives the paid ones. Take one from the other.',
            d: (S2) => where(S2, S2.cells, 'Paid on', (x) => x === undefined).slice(0, 2).map((i) => ['Paid on', i, () => payDate(rows[i])]) };
        }
      ][v]();
      const S = V.S, R = S.R(V.h);
      const base = Object.assign({}, S.cells, { A1: V.label });
      const alt = later(base, S, V.d(S));
      const Ri = V.unpaid ? S.R('Invoice') : null;
      return H.problem({
        title: V.title, file: V.file, sheet: S, base, alt, answerCell: 'B1', dateCells: V.unpaid ? filledRefs(S, base, 'Paid on') : undefined,
        statement: V.unpaid ? `<p>${V.ask}</p><p>In <b>B1</b>, count the unpaid invoices: the entries in <b>Invoice</b> (${Ri}) less the dates in <b>Paid on</b> (${R}).</p>`
          : `<p>${V.ask}</p><p>In <b>B1</b>, count the cells in <b>${V.h}</b> (${R}) that are not empty.</p>`,
        solution: V.unpaid ? `=COUNTA(${Ri})-COUNT(${R})` : `=COUNTA(${R})`,
        alts: V.unpaid ? [`=ROWS(${R})-COUNT(${R})`, `=COUNTBLANK(${R})`, `=COUNTA(${S.A('Invoice')})-COUNT(${S.A('Paid on')})`] : [`=COUNTA(${S.A(V.h)})`, `=ROWS(${R})-COUNTBLANK(${R})`],
        hints: ['COUNTA counts every cell that is not empty: numbers, words and dates.', V.hint].concat(V.unpaid ? ['The shape: =COUNTA(Invoice column)-COUNT(Paid on column).'] : []),
        explain: V.unpaid ? `COUNTA(${Ri}) counts all the invoices and COUNT(${R}) counts the payment dates, which are numbers to Excel. The difference is the invoices with no date yet.`
          : `COUNTA(${R}) counts every filled cell in the column, words included, and leaves out the blanks. COUNT would see only the numbers.`,
        chips: V.unpaid ? [Ri, R] : [R], fnKey: 'COUNTA('
      });
    }
  });

  P.add({
    key: 'average-across-a-row', difficulty: 'Easy', level: 2, topics: ['AVERAGE'], count: 4,
    make(r, v) {
      const cart = (sku, ym) => dbSum((s) => s.sku === sku && s.date.startsWith(ym), 'cartons');
      const WARUNG = ['Warung Bu Ani', 'Warung Pojok', 'Toko Sumber', 'Kios Pasar'];
      const V = [
        () => ({ title: 'Average monthly cartons per product, May to July', ans: 'Monthly average', file: 'penjualan_produk.xlsx', ask: 'Tika wants the average cartons a month for every product over May, June and July',
          S: H.table([{ h: 'Product', k: 'name' }, { h: 'May', k: 'a', num: true }, { h: 'Jun', k: 'b', num: true }, { h: 'Jul', k: 'c', num: true }], H.PRODUCTS.map((p) => ({ name: p.name, a: cart(p.sku, '2026-05'), b: cart(p.sku, '2026-06'), c: cart(p.sku, '2026-07') }))),
          d: [['Jul', 1, 12], ['May', 6, -8], ['Jun', -1, 15]], n: 3 }),
        () => ({ title: 'Average weekly visits per rep over four weeks', ans: 'Weekly average', file: 'kunjungan_sales.xlsx', ask: 'Pak Anton wants each rep\'s average visits a week over the last four weeks',
          S: H.table([{ h: 'Rep', k: 'name' }, { h: 'Week 1', k: 'w1', num: true }, { h: 'Week 2', k: 'w2', num: true }, { h: 'Week 3', k: 'w3', num: true }, { h: 'Week 4', k: 'w4', num: true }], r.sample(H.REPS, 10).map((name) => ({ name, w1: r.int(3, 16), w2: r.int(3, 16), w3: r.int(3, 16), w4: r.int(3, 16) }))),
          d: [['Week 3', 0, 4], ['Week 1', 5, 2], ['Week 4', -1, 3]], n: 4 }),
        () => {
          const rows = H.loans(r, 12).map((x) => { const a = r.pick([0, 0, 0, 5, 15, 30, 45, 60]); const b = Math.max(0, a + r.pick([-15, 0, 0, 15, 30])); const c = Math.max(0, b + r.pick([-15, 0, 15, 30])); return Object.assign(x, { a, b, c }); });
          return { title: 'Average month-end days past due per facility', ans: 'Average DPD', file: 'buku_pinjaman.xlsx', ask: 'Citra wants each facility\'s average days past due over the last three month-ends',
            S: H.table([C.facility, C.dist, { h: 'Aug DPD', k: 'a', num: true }, { h: 'Sep DPD', k: 'b', num: true }, { h: 'Oct DPD', k: 'c', num: true }], rows),
            d: [['Oct DPD', 0, 30], ['Oct DPD', 4, 15], ['Sep DPD', -1, 15]], n: 3 };
        },
        () => {
          /* Prices per pack in steps of 300, so every average lands on whole rupiah. A blank means the warung does not stock it. */
          const rows = H.PRODUCTS.map((p, i) => {
            const base = Math.ceil((p.carton_price / p.packs_per_carton) * 1.2 / 300) * 300;
            const x = { name: p.name };
            WARUNG.forEach((w, j) => { x['w' + j] = r.chance(0.22) ? undefined : base + r.pick([0, 0, 300, 600]); });
            if (i === 2) x.w1 = undefined;
            if (i === 7) x.w3 = undefined;
            if (WARUNG.every((w, j) => x['w' + j] === undefined)) x.w0 = base;
            return x;
          });
          return { title: 'Average shelf price across four warungs, skipping blanks', ans: 'Average price (Rp)', file: 'harga_warung.xlsx', ask: 'Pak Anton surveyed the price per pack in four warungs. A blank means that warung does not stock the product, so leave it out of the average',
            S: H.table([{ h: 'Product', k: 'name' }].concat(WARUNG.map((w, j) => ({ h: w, k: 'w' + j }))), rows), gaps: true,
            /* The hidden sheet: Kios Pasar starts stocking a product it lacked (at a price near the row's others), and one price rises. */
            d: (S2) => {
              const blank = where(S2, S2.cells, 'Kios Pasar', (x) => x === undefined)[0], full = where(S2, S2.cells, 'Warung Pojok', (x) => typeof x === 'number');
              const i = blank === undefined ? 0 : blank, near = WARUNG.map((w) => S2.cells[at(S2, w, i)]).find((x) => typeof x === 'number');
              return [['Kios Pasar', i, (x) => (typeof x === 'number' ? x + 300 : near + 600)], ['Warung Pojok', full[1], 300]];
            } };
        }
      ][v]();
      const S = V.S, from = (V.gaps ? S.cols[1] : S.num[0]) + S.first;
      const to = S.cols[S.cols.length - 1] + S.first, rng = `${from}:${to}`, first = S.next + S.first, last = S.next + S.last;
      const alt = later(S.cells, S, typeof V.d === 'function' ? V.d(S) : V.d);
      return H.problem({
        title: V.title, file: V.file, sheet: S, base: S.cells, alt, fill: true, ansLabel: V.ans,
        statement: `<p>${V.ask}.</p><p>In <b>${first}</b>, average the figures from <b>${from}</b> to <b>${to}</b>, then fill it down to <b>${last}</b>.</p>`,
        solution: `=AVERAGE(${rng})`, alts: V.gaps ? [`=SUM(${rng})/COUNT(${rng})`] : [`=SUM(${rng})/${V.n}`],
        hints: ['AVERAGE works across a row as well as down a column.', V.gaps ? 'AVERAGE skips blank cells by itself, so a warung with no price does not pull the average down.' : `The first row's range is ${rng}.`],
        explain: V.gaps ? `AVERAGE(${rng}) adds the prices that are there and divides by how many there are, so blanks are left out. Dividing by 4 every time would count a missing price as 0.`
          : `AVERAGE(${rng}) adds the ${V.n} figures on the row and divides by ${V.n}. Filling down gives every row its own average.`,
        chips: [rng], fnKey: 'AVERAGE('
      });
    }
  });

  P.add({
    key: 'ratio-on-each-row', difficulty: 'Easy', level: 2, topics: ['Arithmetic', 'Percentages'], count: 4,
    make(r, v) {
      const V = [
        () => ({ title: 'Utilization of each loan facility as a decimal', ans: 'Utilization', file: 'buku_pinjaman.xlsx', ask: 'Dimas wants the utilization of every facility: <b>Outstanding (Rp)</b> divided by <b>Limit (Rp)</b>', S: H.table([C.facility, C.dist, C.limit, C.outstanding], H.loans(r, 12)), top: 'Outstanding (Rp)', bottom: 'Limit (Rp)', d: [['Outstanding (Rp)', 0, 4000000], ['Outstanding (Rp)', 6, -3000000], ['Limit (Rp)', -1, 25000000]] }),
        () => ({ title: 'Target achievement of each sales rep', ans: 'Achievement', file: 'tim_sales.xlsx', ask: 'Pak Anton wants each rep\'s achievement: <b>Actual (Rp)</b> divided by <b>Target (Rp)</b>. 1.05 means 105% of target', S: H.table([C.rep, C.area, C.target, C.actual], H.reps(r, 10)), top: 'Actual (Rp)', bottom: 'Target (Rp)', d: [['Actual (Rp)', 1, 5000000], ['Actual (Rp)', 5, -2400000], ['Target (Rp)', -1, 20000000]] }),
        () => ({ title: 'Return rate for each month of the year', ans: 'Return rate', file: 'penjualan_bulanan.xlsx', ask: 'What share of each month\'s sales came back as returns? <b>Returns (Rp)</b> divided by <b>Sales (Rp)</b>', S: H.table([C.month, C.sales, C.returns], H.monthly(r)), top: 'Returns (Rp)', bottom: 'Sales (Rp)', d: [['Returns (Rp)', 2, 3500000], ['Returns (Rp)', 8, 1200000], ['Sales (Rp)', -1, 40000000]] }),
        () => ({ title: 'Gross margin for each month of the year', ans: 'Gross margin', file: 'penjualan_bulanan.xlsx', ask: 'Tika wants the gross margin of every month: <b>Sales (Rp)</b> minus <b>Cost (Rp)</b>, divided by <b>Sales (Rp)</b>', S: H.table([C.month, C.sales, C.cost], H.monthly(r)), top: 'Sales (Rp)', bottom: 'Cost (Rp)', margin: true, d: [['Cost (Rp)', 1, 15000000], ['Cost (Rp)', 7, -9000000], ['Sales (Rp)', -1, 30000000]] })
      ][v]();
      const S = V.S, a = S.cell(V.top, 0), b = S.cell(V.bottom, 0), first = S.next + S.first, last = S.next + S.last;
      const alt = later(S.cells, S, V.d);
      return H.problem({
        title: V.title, file: V.file, sheet: S, base: S.cells, alt, fill: true, ansLabel: V.ans, precision: 4,
        statement: `<p>${V.ask}.</p><p>Type it in <b>${first}</b> as a decimal (0.85 means 85%) and fill it down to <b>${last}</b>.</p>`,
        solution: V.margin ? `=(${a}-${b})/${a}` : `=${a}/${b}`,
        alts: V.margin ? [`=1-${b}/${a}`] : [],
        hints: V.margin ? ['Work out the profit first, then divide it by sales.', `Brackets matter: (${a}-${b}) has to happen before the division.`, `The shape: =(${a}-${b})/${a}.`]
          : ['The / sign divides the first cell by the second.', `The part goes on top: ${a}/${b}.`],
        explain: V.margin ? `(${a}-${b}) is the month's profit; dividing by ${a} turns it into a share of sales. Without the brackets Excel would divide first and subtract after.`
          : `${a}/${b} is the part over the whole on row ${S.first}. Filling down repeats it with each row's own figures.`,
        chips: [a, b]
      });
    }
  });

  P.add({
    key: 'round-an-average', difficulty: 'Easy', level: 2, topics: ['ROUND', 'AVERAGE'], count: 4,
    make(r, v) {
      const V = [
        () => ({ title: 'Average net sale rounded to the nearest thousand rupiah', label: 'Average net (Rp)', h: 'Net (Rp)', file: 'penjualan_q4.xlsx', S: H.table([C.date, C.dist, C.product, C.net], H.sales(r, 14)), digits: -3, words: 'to the nearest thousand rupiah', unit: 100 }),
        () => ({ title: 'Average cartons per line to one decimal place', label: 'Average cartons', h: 'Cartons', file: 'penjualan_q4.xlsx', S: H.table([C.date, C.dist, C.product, C.cartons], H.sales(r, 15)), digits: 1, words: 'to one decimal place', unit: 1 }),
        () => ({ title: 'Average visits per sales rep as a whole number', label: 'Average visits', h: 'Visits', file: 'tim_sales.xlsx', S: H.table([C.rep, C.area, C.actual, C.visits], H.reps(r, 11)), digits: 0, words: 'to a whole number', unit: 1 }),
        () => ({ title: 'Average monthly sales to the nearest million rupiah', label: 'Average month (Rp)', h: 'Sales (Rp)', file: 'penjualan_bulanan.xlsx', S: H.table([C.month, C.sales, C.cost], H.monthly(r)), digits: -6, words: 'to the nearest million rupiah', unit: 1000000 })
      ][v]();
      const S = V.S, R = S.R(V.h), step = Math.pow(10, -V.digits);
      const base = Object.assign({}, S.cells, { A1: V.label });
      /* Make sure the plain average is not already round, so ROUND has work to do. */
      for (let k = 0; k < 6; k++) {
        const vs = colVals(S, base, V.h), avg = vs.reduce((a, x) => a + x, 0) / vs.length, q = avg / step;
        if (Math.abs(q - Math.round(q)) > 1e-6) break;
        base[at(S, V.h, 1)] += V.unit;
      }
      /*
       * The hidden sheet adds to the first and last rows so that the rounded answer moves, and the average lands on the
       * other side of the half-way mark from the visible sheet. ROUNDUP or ROUNDDOWN in place of ROUND is then wrong on
       * one of the two sheets. q is the average counted in rounding steps, worked out without a fractional step.
       */
      const vs = colVals(S, base, V.h), n = vs.length, sum = vs.reduce((a, x) => a + x, 0);
      const q = (s) => (V.digits >= 0 ? s * Math.pow(10, V.digits) / n : s / (n * Math.pow(10, -V.digits)));
      const frac = (x) => x - Math.floor(x), lowSide = frac(q(sum)) < 0.5;
      let add = 0;
      for (let k = 1; k <= 5000 && !add; k++) {
        const q2 = q(sum + k * V.unit), f = frac(q2);
        if (Math.round(q2) !== Math.round(q(sum)) && (lowSide ? f > 0.55 && f < 0.95 : f > 0.05 && f < 0.45)) add = k * V.unit;
      }
      const first = Math.ceil(add / 2 / V.unit) * V.unit;
      const alt = later(base, S, [[V.h, 0, first], [V.h, -1, add - first]]);
      return H.problem({
        title: V.title, file: V.file, sheet: S, base, alt, answerCell: 'B1',
        statement: `<p>In <b>B1</b>, work out the average of <b>${V.h}</b> (${R}), rounded ${V.words}.</p>`,
        solution: `=ROUND(AVERAGE(${R}),${V.digits})`, alts: [`=ROUND(SUM(${R})/COUNT(${R}),${V.digits})`, `=ROUND(AVERAGE(${S.A(V.h)}),${V.digits})`],
        hints: ['ROUND(number, digits) rounds a result. Put the AVERAGE inside it.', 'Digits: 1 keeps one decimal, 0 gives a whole number, -3 rounds to thousands and -6 to millions.', `The shape: =ROUND(AVERAGE(${R}), digits).`],
        explain: `AVERAGE(${R}) works out the average and ROUND(…,${V.digits}) rounds it ${V.words}. Rounding inside the formula keeps the rounded figure for later sums, not only on screen.`,
        chips: [R, `,${V.digits})`], fnKey: 'ROUND('
      });
    }
  });

  P.add({
    key: 'scale-by-an-input-cell', difficulty: 'Easy', level: 2, topics: ['$ references', 'Percentages'], count: 4,
    make(r, v) {
      const V = [
        () => ({ title: 'New carton prices after the price rise in B1', ans: 'New price (Rp)', file: 'daftar_harga.xlsx', label: 'Price rise', b1: 0.05, b1alt: 0.07, ask: 'Rasa Kenari raises its carton prices by the rate in <b>B1</b> (0.05 means 5%). Work out each new price', S: H.table([C.sku, C.product, C.price], priceList(r, 10)), h: 'Carton price (Rp)', d: [['Carton price (Rp)', 3, 4000]], rule: (c) => `${c}*(1+$B$1)`, alts: (c) => [`=${c}+${c}*$B$1`, `=${c}*(1+B$1)`, `=(1+$B$1)*${c}`] }),
        () => ({ title: 'Commission on each rep\'s sales at the rate in B1', ans: 'Commission (Rp)', file: 'tim_sales.xlsx', label: 'Commission rate', b1: 0.02, b1alt: 0.03, ask: 'Reps earn a commission on actual sales at the rate in <b>B1</b> (0.02 means 2%). Work out each rep\'s commission', S: H.table([C.rep, C.area, C.actual], H.reps(r, 10)), h: 'Actual (Rp)', d: [['Actual (Rp)', 2, 3000000]], rule: (c) => `${c}*$B$1`, alts: (c) => [`=$B$1*${c}`, `=${c}*B$1`] }),
        () => ({ title: 'One month of interest at the yearly rate in B1', ans: 'Interest (Rp)', file: 'buku_pinjaman.xlsx', label: 'Interest rate (a year)', b1: 0.09, b1alt: 0.12, ask: 'Citra needs one month of interest on every facility. The yearly rate is in <b>B1</b> (0.09 means 9%), so a month is the outstanding times the rate, divided by 12', S: H.table([C.facility, C.dist, C.outstanding], H.loans(r, 12)), h: 'Outstanding (Rp)', d: [['Outstanding (Rp)', 4, 6000000]], rule: (c) => `${c}*$B$1/12`, alts: (c) => [`=$B$1/12*${c}`, `=${c}*(B$1/12)`] }),
        () => ({ title: 'Net of each line after the promo discount in B1', ans: 'After promo (Rp)', file: 'penjualan_q4.xlsx', label: 'Promo discount', b1: 0.03, b1alt: 0.04, ask: 'A promo takes the discount in <b>B1</b> (0.03 means 3%) off every line\'s gross value. Work out what is left', S: H.table([C.date, C.dist, C.gross], H.sales(r, 14)), h: 'Gross (Rp)', d: [['Gross (Rp)', 5, 2240000]], rule: (c) => `${c}*(1-$B$1)`, alts: (c) => [`=${c}-${c}*$B$1`, `=${c}*(1-B$1)`] })
      ][v]();
      const S = V.S, c = S.cell(V.h, 0), first = S.next + S.first, last = S.next + S.last;
      const base = Object.assign({}, S.cells, { A1: V.label, B1: V.b1 });
      /* The hidden sheet changes the rate in B1, so a rate typed into the formula fails, and one figure moves. */
      const alt = later(Object.assign({}, base, { B1: V.b1alt }), S, V.d);
      return H.problem({
        title: V.title, file: V.file, sheet: S, base, alt, fill: true, ansLabel: V.ans,
        statement: `<p>${V.ask}.</p><p>In <b>${first}</b>, use <b>${V.h}</b> and the rate in <b>B1</b>, then fill it down to <b>${last}</b>. Point at B1 rather than typing the rate, because the rate can change.</p>`,
        solution: '=' + V.rule(c), alts: V.alts(c), traps: [unpinB1('=' + V.rule(c))],
        hints: ['Every row uses the same rate, the one in B1.', 'When you fill down, B1 would slide to B2, B3 and so on. $B$1 stays put.', 'The shape: =' + V.rule(c) + '.'],
        explain: `${V.rule(c)} uses row ${S.first}'s figure and the rate in B1. The $ signs pin B1, so every filled row reads the same rate, and a new rate in B1 updates the whole column.`,
        chips: [c, '$B$1'], fnKey: '$B$1'
      });
    }
  });

  P.add({
    key: 'spread-between-max-and-min', difficulty: 'Easy', level: 2, topics: ['MAX', 'MIN'], count: 4,
    make(r, v) {
      const V = [
        () => ({ title: 'Gap between the dearest and cheapest carton', label: 'Price gap (Rp)', h: 'Carton price (Rp)', file: 'daftar_harga.xlsx', ask: 'How far apart are the dearest and the cheapest carton on this price list?', S: H.table([C.sku, C.product, C.category, C.price], priceList(r, 10)), up: 12000 }),
        () => ({ title: 'Gap between the best and worst sales month', label: 'Best minus worst (Rp)', h: 'Sales (Rp)', file: 'penjualan_bulanan.xlsx', ask: 'How much higher were sales in the best month than in the worst?', S: H.table([C.month, C.sales, C.cost], H.monthly(r)), up: 25000000 }),
        () => ({ title: 'Gap between the most and fewest rep visits', label: 'Visit gap', h: 'Visits', file: 'tim_sales.xlsx', ask: 'How many more visits did the busiest rep make than the quietest?', S: H.table([C.rep, C.area, C.visits], H.reps(r, 10)), up: 6 }),
        () => ({ title: 'Spread between the highest and lowest loan rate', label: 'Rate spread (points)', h: 'Rate (% a year)', file: 'buku_pinjaman.xlsx', ask: 'Dimas wants the spread between the highest and the lowest interest rate in the book, in percentage points.', S: H.table([C.facility, C.dist, C.type, C.rate], H.loans(r, 12)), up: 0.5 })
      ][v]();
      const S = V.S, R = S.R(V.h);
      const base = Object.assign({}, S.cells, { A1: V.label });
      const vs = colVals(S, base, V.h);
      if (Math.max(...vs) === Math.min(...vs)) base[at(S, V.h, 0)] += V.up;
      const top = colVals(S, base, V.h).indexOf(Math.max(...colVals(S, base, V.h)));
      const alt = later(base, S, [[V.h, top, (x) => x + V.up]]);
      return H.problem({
        title: V.title, file: V.file, sheet: S, base, alt, answerCell: 'B1',
        statement: `<p>${V.ask}</p><p>In <b>B1</b>, subtract the smallest value in <b>${V.h}</b> (${R}) from the largest.</p>`,
        solution: `=MAX(${R})-MIN(${R})`, alts: [`=LARGE(${R},1)-SMALL(${R},1)`, `=MAX(${S.A(V.h)})-MIN(${S.A(V.h)})`],
        hints: ['MAX finds the largest value and MIN the smallest.', 'You can put two functions in one formula and subtract one from the other.', `The shape: =MAX(${R})-MIN(${R}).`],
        explain: `MAX(${R})-MIN(${R}) is the range of the column: top value minus bottom value. Both parts follow the data, so a corrected figure updates the gap.`,
        chips: [R], fnKey: 'MAX('
      });
    }
  });

  /* ======================= Level 3: a $ that has to behave when filled, and ratios of totals ======================= */

  P.add({
    key: 'percent-of-total', difficulty: 'Easy', level: 3, topics: ['$ references', 'Percentages', 'SUM'], count: 4,
    make(r, v) {
      const V = [
        () => ({ title: 'Share of total net sales for each order line', ans: 'Share of total', file: 'penjualan_q4.xlsx', ask: 'What share of the total net sales does each line bring in?', S: H.table([C.date, C.dist, C.net], H.sales(r, 12)), h: 'Net (Rp)', d: [[0, 2400000], [5, -800000]] }),
        () => ({ title: 'Share of total outstanding for each facility', ans: 'Share of book', file: 'buku_pinjaman.xlsx', ask: 'Dimas wants each facility\'s share of the total outstanding in the book.', S: H.table([C.facility, C.dist, C.type, C.outstanding], H.loans(r, 12)), h: 'Outstanding (Rp)', d: [[1, 7000000], [-1, -3000000]] }),
        () => ({ title: 'Share of annual sales for each month', ans: 'Share of year', file: 'penjualan_bulanan.xlsx', ask: 'What share of the year\'s sales came from each month?', S: H.table([C.month, C.sales], H.monthly(r)), h: 'Sales (Rp)', d: [[10, 45000000], [3, -12000000]] }),
        () => ({ title: 'Share of team sales for each rep', ans: 'Share of team', file: 'tim_sales.xlsx', ask: 'Pak Anton wants each rep\'s share of the team\'s actual sales.', S: H.table([C.rep, C.area, C.actual], H.reps(r, 10)), h: 'Actual (Rp)', d: [[2, 9000000], [7, -4000000]] })
      ][v]();
      const S = V.S, c = S.cell(V.h, 0), R = S.R(V.h), A = S.A(V.h), first = S.next + S.first, last = S.next + S.last;
      const half = `${S.col(V.h)}$${S.first}:${S.col(V.h)}$${S.last}`;
      const alt = later(S.cells, S, V.d.map(([i, d]) => [V.h, i, d]));
      return H.problem({
        title: V.title, file: V.file, sheet: S, base: S.cells, alt, fill: true, ansLabel: V.ans, precision: 4,
        statement: `<p>${V.ask}</p><p>In <b>${first}</b>, divide the row's <b>${V.h}</b> by the total of ${R}, as a decimal (0.12 means 12%). Fill it down to <b>${last}</b>.</p>`,
        solution: `=${c}/SUM(${A})`, alts: [`=${c}/SUM(${half})`],
        traps: [{ code: 'unpinned-total', title: 'The total slid down as you filled', plain: `Without $ signs the SUM range moves down a row each time, so later rows divide by a smaller total. Pin it: SUM(${A}).`, formula: `=${c}/SUM(${R})` }],
        hints: ['A share is the part divided by the whole: the row\'s figure over the column total.', `The total must not move when you fill down, so pin it: SUM(${A}).`, `The shape: =${c}/SUM(${A}).`],
        explain: `${c} moves down with each row, while SUM(${A}) stays on the same total because of the $ signs. Without them the range would slide down and the shares would come out too big.`,
        chips: [c, A], fnKey: 'SUM('
      });
    }
  });

  P.add({
    key: 'running-total', difficulty: 'Easy', level: 3, topics: ['$ references', 'SUM'], count: 4,
    make(r, v) {
      const V = [
        () => ({ title: 'Year-to-date sales at the end of each month', ans: 'Year to date (Rp)', file: 'penjualan_bulanan.xlsx', ask: 'Tika wants the year-to-date sales at the end of every month: January alone, then January plus February, and so on', S: H.table([C.month, C.sales], H.monthly(r)), h: 'Sales (Rp)', d: [['Sales (Rp)', 2, 16000000], ['Sales (Rp)', 8, -7000000]] }),
        () => ({ title: 'Cartons shipped so far, line by line', ans: 'Cartons so far', file: 'penjualan_q4.xlsx', ask: 'Pak Anton wants a running total of cartons shipped, line by line in date order', S: H.table([C.date, C.dist, C.cartons], H.sales(r, 14)), h: 'Cartons', d: [['Cartons', 1, 9], ['Cartons', 9, 4]] }),
        () => {
          let bal = r.pick([180, 220, 240, 260]);
          const open = bal, start = H.toSerial('2026-10-01'), rows = [];
          let day = start;
          for (let i = 0; i < 14; i++) {
            day += r.int(1, 3);
            if (bal < 120 || (i > 2 && r.chance(0.12))) { const q = r.pick([150, 200, 240]); bal += q; rows.push({ date: day, ref: 'GRN-' + (3100 + i * 7), inn: q }); }
            else { const q = r.int(3, 9) * 10; bal -= q; rows.push({ date: day, ref: 'DO-' + (5200 + i * 3), out: q }); }
          }
          const p = r.pick(H.PRODUCTS), w = r.pick(H.WAREHOUSES);
          const S = H.table([C.date, { h: 'Reference', k: 'ref' }, { h: 'In (cartons)', k: 'inn' }, { h: 'Out (cartons)', k: 'out' }], rows);
          const outs = where(S, S.cells, 'Out (cartons)', (x) => typeof x === 'number');
          return { title: 'Running stock balance from an opening count in B1', ans: 'Balance (cartons)', file: 'mutasi_stok.xlsx', S, balance: true,
            top: { A1: 'Opening stock (cartons)', B1: open, A2: 'Product', B2: p.name, A3: 'Warehouse', B3: w }, topAlt: { B1: open + 20 },
            ask: 'The stock card starts from the opening count in <b>B1</b>. Receipts come <b>In</b> and deliveries go <b>Out</b>. Work out the balance after every movement',
            d: [['Out (cartons)', outs[1], 10]] };
        },
        () => {
          const rows = H.sales(r, 12), total = rows.reduce((a, x) => a + x.cartons, 0), goal = Math.round(total * 0.85 / 10) * 10;
          return { title: 'Cartons still needed to reach the target in B1', ans: 'Still needed', file: 'penjualan_q4.xlsx', S: H.table([C.date, C.dist, C.cartons], rows), h: 'Cartons', left: true,
            top: { A1: 'Quarter target (cartons)', B1: goal }, topAlt: { B1: goal + 40 },
            ask: 'The quarter target in cartons is in <b>B1</b>. After each line, how many cartons are still needed? A negative number means the target is passed',
            d: [['Cartons', 3, 6]] };
        }
      ][v]();
      const S = V.S, first = S.next + S.first, last = S.next + S.last;
      const base = Object.assign({}, S.cells, V.top || {});
      const alt = later(Object.assign({}, base, V.topAlt || {}), S, V.d);
      const run = (h) => { const c = S.col(h); return `$${c}$${S.first}:${c}${S.first}`; };
      const half = (h) => { const c = S.col(h); return `${c}$${S.first}:${c}${S.first}`; };
      let solution, alts, hints, explain, chips;
      if (V.balance) {
        const i = run('In (cartons)'), o = run('Out (cartons)');
        solution = `=$B$1+SUM(${i})-SUM(${o})`;
        alts = [`=SUM(${i})-SUM(${o})+$B$1`, `=$B$1+SUM(${half('In (cartons)')})-SUM(${half('Out (cartons)')})`];
        hints = ['The balance on any row is the opening stock, plus everything in so far, minus everything out so far.', `"So far" is a range that starts on row ${S.first} and grows: pin the start with $ and leave the end free, as in ${i}.`, `The shape: =$B$1+SUM(${i})-SUM(${o}).`];
        explain = `SUM(${i}) grows by one row each time you fill down, because only its start is pinned. Adding the opening stock in $B$1 and taking away the outs so far gives the balance on every row.`;
        chips = ['$B$1', i, o];
      } else {
        const g = run(V.h);
        solution = V.left ? `=$B$1-SUM(${g})` : `=SUM(${g})`;
        alts = V.left ? [`=$B$1-SUM(${half(V.h)})`] : [`=SUM(${half(V.h)})`];
        hints = ['A running total adds everything from the first row down to the current row.', `Pin the start of the range and leave the end free: ${g}. When you fill down the end moves and the start stays.`].concat(V.left ? [`Take that running total away from the target in $B$1: =$B$1-SUM(${g}).`] : []);
        explain = `${g} starts as one cell and grows by a row each time it is filled down, because only the start is pinned.` + (V.left ? ' Taking it away from the target in $B$1 leaves what is still needed.' : ' So each row shows the total up to that point.');
        chips = V.left ? ['$B$1', g] : [g];
      }
      return H.problem({
        title: V.title, file: V.file, sheet: S, base, alt, fill: true, ansLabel: V.ans,
        statement: `<p>${V.ask}.</p><p>Type one formula in <b>${first}</b> and fill it down to <b>${last}</b>.</p>`,
        solution, alts, hints, explain, chips, fnKey: 'SUM(',
        traps: [unpinStart(solution)].concat(/\$B\$1/.test(solution) ? [unpinB1(solution)] : [])
      });
    }
  });

  P.add({
    key: 'growth-between-two-columns', difficulty: 'Easy', level: 3, topics: ['Percentages', 'Arithmetic'], count: 4,
    make(r, v) {
      const net = (id, ym) => dbSum((s) => s.distributor_id === id && s.date.startsWith(ym), 'net');
      const cart = (sku, ym) => dbSum((s) => s.sku === sku && s.date.startsWith(ym), 'cartons');
      const V = [
        () => ({ title: 'Price change for each product, old to new', ans: 'Change', file: 'daftar_harga.xlsx', ask: 'The new price list is out. By what share did each carton price change?',
          S: H.table([C.product, { h: 'Old price (Rp)', k: 'old', num: true }, { h: 'New price (Rp)', k: 'nw', num: true }], priceList(r, 10).map((p) => ({ product: p.product, old: p.price, nw: p.price + r.pick([0, 2000, 4000, 6000, 8000, 10000, -4000]) }))),
          from: 'Old price (Rp)', to: 'New price (Rp)', d: [['New price (Rp)', 2, 2000], ['New price (Rp)', 7, -2000]] }),
        () => ({ title: 'Net sales growth per distributor, Sep to Oct', ans: 'Growth', file: 'penjualan_distributor.xlsx', ask: 'Citra watches distributors whose sales are falling. What is each one\'s net sales growth from September to October?',
          S: H.table([{ h: 'Distributor', k: 'name' }, { h: 'Sep (Rp)', k: 'a', num: true }, { h: 'Oct (Rp)', k: 'b', num: true }], r.sample(H.db().distributors, 10).map((d) => ({ name: d.name, a: net(d.id, '2026-09'), b: net(d.id, '2026-10') }))),
          from: 'Sep (Rp)', to: 'Oct (Rp)', d: [['Oct (Rp)', 1, 3600000], ['Oct (Rp)', 6, -1400000]] }),
        () => ({ title: 'Carton growth per product, Oct against Aug', ans: 'Growth', file: 'penjualan_produk.xlsx', ask: 'How did each product\'s cartons in October compare with August? Careful: October comes first on this sheet.',
          S: H.table([{ h: 'Product', k: 'name' }, { h: 'Oct (cartons)', k: 'b', num: true }, { h: 'Aug (cartons)', k: 'a', num: true }], H.PRODUCTS.map((p) => ({ name: p.name, a: cart(p.sku, '2026-08'), b: cart(p.sku, '2026-10') }))),
          from: 'Aug (cartons)', to: 'Oct (cartons)', d: [['Oct (cartons)', 0, 14], ['Aug (cartons)', 9, 10]] }),
        () => {
          const rows = H.loans(r, 12).map((x) => Object.assign(x, { sep: Math.max(1000000, Math.round(x.outstanding * r.pick([0.9, 0.95, 1, 1, 1.04, 1.1]) / 100000) * 100000) }));
          return { title: 'Change in outstanding per facility over October', ans: 'Change', file: 'buku_pinjaman.xlsx', ask: 'Dimas compares the outstanding at the end of September with the end of October. By what share did each facility move?',
            S: H.table([C.facility, C.dist, { h: 'Sep month-end (Rp)', k: 'sep', num: true }, { h: 'Oct month-end (Rp)', k: 'outstanding', num: true }], rows),
            from: 'Sep month-end (Rp)', to: 'Oct month-end (Rp)', d: [['Oct month-end (Rp)', 3, 5000000], ['Sep month-end (Rp)', 8, 2000000]] };
        }
      ][v]();
      const S = V.S, o = S.cell(V.from, 0), n = S.cell(V.to, 0), first = S.next + S.first, last = S.next + S.last;
      const alt = later(S.cells, S, V.d);
      return H.problem({
        title: V.title, file: V.file, sheet: S, base: S.cells, alt, fill: true, ansLabel: V.ans, precision: 4,
        statement: `<p>${V.ask}</p><p>In <b>${first}</b>, work out the change from <b>${V.from}</b> to <b>${V.to}</b> as a share of <b>${V.from}</b>, as a decimal (0.05 means up 5%, -0.1 means down 10%). Fill it down to <b>${last}</b>.</p>`,
        solution: `=(${n}-${o})/${o}`, alts: [`=${n}/${o}-1`],
        hints: ['Growth is the change divided by where you started: (new minus old) divided by old.', `Here the old figure is ${o} and the new one is ${n}.`, `The shape: =(${n}-${o})/${o}.`],
        explain: `${n}-${o} is the change, and dividing by ${o}, the starting figure, turns it into a share. The brackets make Excel subtract before it divides.`,
        chips: [n, o]
      });
    }
  });

  P.add({
    key: 'round-each-line', difficulty: 'Easy', level: 3, topics: ['ROUND', 'ROUNDUP', 'ROUNDDOWN'], count: 4,
    make(r, v) {
      const V = [
        () => {
          const S = H.table([C.date, C.dist, C.net], H.sales(r, 14));
          const c = S.cell('Net (Rp)', 0);
          return { title: 'VAT on each line, rounded to the nearest hundred rupiah', ans: 'VAT (Rp)', file: 'penjualan_q4.xlsx', S, top: { A1: 'VAT rate', B1: 0.11 }, topAlt: { B1: 0.12 }, d: [['Net (Rp)', 4, 1300000]],
            ask: 'Tika needs the VAT on every line: <b>Net (Rp)</b> times the rate in <b>B1</b> (0.11 means 11%), rounded to the nearest hundred rupiah',
            solution: `=ROUND(${c}*$B$1,-2)`, alts: [`=ROUND($B$1*${c},-2)`, `=ROUND(${c}*B$1,-2)`], fnKey: 'ROUND(',
            hints: ['ROUND(number, -2) rounds to the nearest hundred.', 'The rate is the same for every row, so pin it: $B$1.', `The shape: =ROUND(${c}*$B$1,-2).`],
            explain: `${c}*$B$1 is the VAT on the line, and ROUND(…,-2) takes it to the nearest hundred rupiah. The $ keeps every row on the rate in B1.`, chips: [c, '$B$1', ',-2)'] };
        },
        () => {
          const rows = r.sample(H.db().outlets, 12).map((o) => { const p = r.pick(H.PRODUCTS); return { outlet: o.outlet, product: p.name, packs: p.packs_per_carton, ordered: r.int(3, 40) * 10 + r.pick([0, 4, 6, 12]) }; });
          rows[0].ordered = rows[0].packs * 3 + 1;
          const S = H.table([{ h: 'Outlet', k: 'outlet' }, C.product, C.packs, { h: 'Packs ordered', k: 'ordered', num: true }], rows);
          const a = S.cell('Packs ordered', 0), b = S.cell('Packs per carton', 0);
          /* The hidden sheet: two outlets order more, by at least one whole carton, so the rounded answer has to move. */
          return { title: 'Cartons needed for each pack order, rounded up', ans: 'Cartons to send', file: 'pesanan_outlet.xlsx', S, d: [['Packs ordered', 2, rows[2].packs], ['Packs ordered', 8, rows[8].packs * 2 + 5]],
            ask: 'Outlets order in packs, but the warehouse ships whole cartons. Work out the cartons to send: <b>Packs ordered</b> divided by <b>Packs per carton</b>, always rounded up so no outlet is short',
            solution: `=ROUNDUP(${a}/${b},0)`, alts: [`=CEILING(${a}/${b},1)`, `=-INT(-${a}/${b})`], fnKey: 'ROUNDUP(',
            hints: ['ROUND would sometimes round down and leave an outlet short. ROUNDUP always goes up.', 'ROUNDUP(number, 0) rounds up to a whole number.', `The shape: =ROUNDUP(${a}/${b},0).`],
            explain: `${a}/${b} is the exact number of cartons, often with a fraction left. ROUNDUP(…,0) takes any fraction up to the next whole carton.`, chips: [a, b, ',0)'] };
        },
        () => {
          const S = H.table([C.facility, C.dist, C.outstanding, C.rate], H.loans(r, 12));
          const a = S.cell('Outstanding (Rp)', 0), b = S.cell('Rate (% a year)', 0);
          return { title: 'Monthly interest per facility, rounded down to whole rupiah', ans: 'Interest (Rp)', file: 'buku_pinjaman.xlsx', S, d: [['Outstanding (Rp)', 1, 3700000], ['Rate (% a year)', 6, (x) => x + 0.5]],
            ask: 'Citra needs one month of interest on every facility, rounded down to whole rupiah: <b>Outstanding (Rp)</b> times <b>Rate (% a year)</b>, divided by 100 to turn the rate into a fraction and by 12 for one month',
            solution: `=ROUNDDOWN(${a}*${b}/100/12,0)`, alts: [`=ROUNDDOWN(${a}*${b}/1200,0)`, `=INT(${a}*${b}/1200)`, `=TRUNC(${a}*${b}/100/12)`], fnKey: 'ROUNDDOWN(',
            hints: ['The rate is written as a percent, such as 9.25, so divide by 100. One month is a twelfth of a year, so divide by 12 as well.', 'ROUNDDOWN(number, 0) drops the fraction of a rupiah.', `The shape: =ROUNDDOWN(${a}*${b}/100/12,0).`],
            explain: `${a}*${b}/100/12 is a month of interest; ROUNDDOWN(…,0) cuts off the fraction so the bank never charges part of a rupiah too much.`, chips: [a, b, '/100/12', ',0)'] };
        },
        () => {
          const S = H.table([C.sku, C.product, C.price, C.packs], H.PRODUCTS.map((p) => ({ sku: p.sku, product: p.name, price: p.carton_price, packs: p.packs_per_carton })));
          const a = S.cell('Carton price (Rp)', 0), b = S.cell('Packs per carton', 0);
          return { title: 'Shelf price per pack after the markup in B1, rounded up', ans: 'Shelf price (Rp)', file: 'harga_rak.xlsx', S, top: { A1: 'Warung markup', B1: 0.25 }, topAlt: { B1: 0.3 }, d: [['Carton price (Rp)', 5, 6000]],
            ask: 'A suggested shelf price per pack starts from the price of one pack (<b>Carton price (Rp)</b> divided by <b>Packs per carton</b>), raises it by the markup in <b>B1</b> (0.25 means 25% more), then rounds up to the next hundred rupiah',
            solution: `=ROUNDUP(${a}/${b}*(1+$B$1),-2)`, alts: [`=ROUNDUP(${a}*(1+$B$1)/${b},-2)`, `=ROUNDUP(${a}/${b}*(1+B$1),-2)`], fnKey: 'ROUNDUP(',
            hints: ['First the price of one pack: carton price divided by packs per carton. Then add the markup with *(1+$B$1).', 'ROUNDUP(number, -2) rounds up to the next hundred.', `The shape: =ROUNDUP(${a}/${b}*(1+$B$1),-2).`],
            explain: `${a}/${b} is the cost of one pack, *(1+$B$1) adds the markup, and ROUNDUP(…,-2) lifts it to a price a warung can charge. $B$1 stays pinned as you fill.`, chips: [a, b, '$B$1', ',-2)'] };
        }
      ][v]();
      const S = V.S, first = S.next + S.first, last = S.next + S.last;
      const base = Object.assign({}, S.cells, V.top || {});
      const alt = later(Object.assign({}, base, V.topAlt || {}), S, V.d);
      return H.problem({
        title: V.title, file: V.file, sheet: S, base, alt, fill: true, ansLabel: V.ans,
        statement: `<p>${V.ask}.</p><p>Type it in <b>${first}</b> and fill it down to <b>${last}</b>.</p>`,
        solution: V.solution, alts: V.alts, hints: V.hints, explain: V.explain, chips: V.chips, fnKey: V.fnKey,
        traps: /\$B\$1/.test(V.solution) ? [unpinB1(V.solution)] : []
      });
    }
  });

  P.add({
    key: 'ratio-of-two-totals', difficulty: 'Easy', level: 3, topics: ['SUM', 'Percentages'], count: 4,
    make(r, v) {
      const V = [
        () => ({ title: 'Utilization of the whole loan book', label: 'Book utilization', file: 'buku_pinjaman.xlsx', ask: 'Dimas wants the utilization of the whole book: total <b>Outstanding (Rp)</b> divided by total <b>Limit (Rp)</b>.', S: H.table([C.facility, C.dist, C.limit, C.outstanding], H.loans(r, 12)), top: 'Outstanding (Rp)', bottom: 'Limit (Rp)', d: [['Outstanding (Rp)', 2, 9000000], ['Limit (Rp)', -1, 50000000]] }),
        () => ({ title: 'Achievement of the whole sales team', label: 'Team achievement', file: 'tim_sales.xlsx', ask: 'Pak Anton wants the team\'s achievement: total <b>Actual (Rp)</b> divided by total <b>Target (Rp)</b>.', S: H.table([C.rep, C.area, C.target, C.actual], H.reps(r, 10)), top: 'Actual (Rp)', bottom: 'Target (Rp)', d: [['Actual (Rp)', 3, 8000000], ['Target (Rp)', 0, 20000000]] }),
        () => ({ title: 'Return rate for the whole year', label: 'Return rate', file: 'penjualan_bulanan.xlsx', ask: 'What share of the year\'s sales came back as returns? Total <b>Returns (Rp)</b> divided by total <b>Sales (Rp)</b>.', S: H.table([C.month, C.sales, C.returns], H.monthly(r)), top: 'Returns (Rp)', bottom: 'Sales (Rp)', d: [['Returns (Rp)', 5, 6000000], ['Returns (Rp)', 10, 2500000]] }),
        () => ({ title: 'Overall discount rate across the order lines', label: 'Discount rate', file: 'penjualan_q4.xlsx', ask: 'Tika wants the overall discount rate: total <b>Gross (Rp)</b> minus total <b>Net (Rp)</b>, divided by total <b>Gross (Rp)</b>.', S: H.table([C.date, C.dist, C.gross, C.net], H.sales(r, 14)), top: 'Net (Rp)', bottom: 'Gross (Rp)', disc: true, d: [['Net (Rp)', 1, -1500000], ['Net (Rp)', 9, -900000]] })
      ][v]();
      const S = V.S, T = S.R(V.top), B = S.R(V.bottom);
      const base = Object.assign({}, S.cells, { A1: V.label });
      const alt = later(base, S, V.d);
      return H.problem({
        title: V.title, file: V.file, sheet: S, base, alt, answerCell: 'B1', precision: 4,
        statement: `<p>${V.ask}</p><p>Type it in <b>B1</b> as a decimal (0.85 means 85%). Work from the column totals, not from an average of each row's ratio.</p>`,
        solution: V.disc ? `=(SUM(${B})-SUM(${T}))/SUM(${B})` : `=SUM(${T})/SUM(${B})`,
        alts: V.disc ? [`=1-SUM(${T})/SUM(${B})`, `=SUM(${B}-${T})/SUM(${B})`] : [`=SUM(${S.A(V.top)})/SUM(${S.A(V.bottom)})`],
        hints: [V.disc ? 'Add up each column first, then work with the two totals.' : 'Add up each column first, then divide one total by the other.', V.disc ? 'The discount is gross minus net. Put that difference in brackets before dividing by the gross total.' : `The top of the fraction is SUM(${T}), the bottom SUM(${B}).`, V.disc ? `The shape: =(SUM(${B})-SUM(${T}))/SUM(${B}).` : `The shape: =SUM(${T})/SUM(${B}).`],
        explain: V.disc ? `(SUM(${B})-SUM(${T})) is the discount in rupiah, and dividing by SUM(${B}) makes it a share of the gross. Big lines weigh more, as they should.`
          : `SUM(${T})/SUM(${B}) weighs every row by its size. An average of each row's ratio would give a small row the same weight as a big one.`,
        chips: [T, B], fnKey: 'SUM('
      });
    }
  });
});
