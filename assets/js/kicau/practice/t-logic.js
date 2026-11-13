/*
 * Practice, Logic (30 Easy, 40 Medium): IF, AND, OR, NOT, nested IF, IFS, SWITCH and IFERROR. Each template makes
 * `count` problems; v (0 to count-1) picks the variant and r is its seeded random.
 * Levels: 1 is one IF with one test, 2 returns a number or reads its limit from B1, 3 joins two tests with AND or OR,
 * or turns a column total into a verdict. 4 is three tiers, NOT and IFERROR, 5 is IFS and SWITCH, joins a number
 * into a label, or checks three rules at once with AND, 6 nests AND with OR or multiplies by a tiered rate, and 7
 * applies the first of several overlapping rules in a set order.
 *
 * Most answers are a column filled down. Each variant states its rule as a small function of the row (rule), and the
 * build checks that every outcome shows up on the visible sheet and that the hidden sheet changes at least one row.
 * Rows are planted with values that sit on a limit, so "at least" and "more than" give different answers.
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
    gross: { h: 'Gross (Rp)', k: 'gross', num: true },
    net: { h: 'Net (Rp)', k: 'net', num: true },
    facility: { h: 'Facility', k: 'facility' },
    type: { h: 'Type', k: 'type' },
    limit: { h: 'Limit (Rp)', k: 'limit', num: true },
    outstanding: { h: 'Outstanding (Rp)', k: 'outstanding', num: true },
    dpd: { h: 'Days past due', k: 'dpd', num: true },
    restr: { h: 'Restructured', k: 'restructured' },
    rep: { h: 'Rep', k: 'name' },
    area: { h: 'Area', k: 'area' },
    target: { h: 'Target (Rp)', k: 'target', num: true },
    actual: { h: 'Actual (Rp)', k: 'actual', num: true },
    visits: { h: 'Visits', k: 'visits', num: true },
    start: { h: 'Start date', k: 'start', date: true },
    month: { h: 'Month', k: 'month' },
    sales: { h: 'Sales (Rp)', k: 'sales', num: true },
    returns: { h: 'Returns (Rp)', k: 'returns', num: true },
    cost: { h: 'Cost (Rp)', k: 'cost', num: true },
    warehouse: { h: 'Warehouse', k: 'warehouse' },
    onhand: { h: 'On hand', k: 'onhand', num: true },
    reorder: { h: 'Reorder point', k: 'reorder', num: true },
    lead: { h: 'Lead time (days)', k: 'lead', num: true },
    invoice: { h: 'Invoice', k: 'ref' },
    amount: { h: 'Amount (Rp)', k: 'amount', num: true },
    terms: { h: 'Terms (days)', k: 'terms', num: true },
    /* Paid on has gaps (open invoices), so it is not flagged date: its filled cells go in dateCells instead. */
    paid: { h: 'Paid on', k: 'paidOn' }
  };
  const DIST = () => H.db().distributors;
  /* The k-th distributor of a region, as the two cells of a sales row. */
  const inRegion = (region, k) => ({ Distributor: DIST().filter((d) => d.region === region)[k || 0].name, Region: region });
  /* The k-th product of a category, as the cells of a sales or stock row. */
  const inCategory = (cat, k) => { const p = H.PRODUCTS.filter((x) => x.category === cat)[k || 0]; return { Product: p.name, SKU: p.sku, Category: cat }; };
  const kol = (d) => (d === 0 ? 1 : d <= 90 ? 2 : d <= 120 ? 3 : d <= 180 ? 4 : 5);
  const commas = (x) => String(x).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  const rp = (x) => 'Rp ' + commas(x);
  const r5 = (x) => Math.round(x / 100000) * 100000;
  const r6 = (x) => Math.round(x / 1000000) * 1000000;

  const at = (S, h, i) => S.col(h) + (S.first + i);
  const size = (S) => S.last - S.first + 1;
  /* Row i of a sheet as { header: value }. */
  function row(S, cells, i) { const o = {}; for (const h of Object.keys(S.map)) o[h] = cells[S.map[h] + (S.first + i)]; return o; }
  /* Change some cells of row i: vals is { header: value, or (row) => value }. null empties the cell; unknown headers are skipped. */
  function put(S, cells, i, vals) {
    const cur = row(S, cells, i);
    for (const [h, x] of Object.entries(vals)) {
      if (!S.map[h]) continue;
      const val = typeof x === 'function' ? x(cur) : x;
      if (val === null || val === undefined) delete cells[at(S, h, i)]; else cells[at(S, h, i)] = val;
    }
    return cells;
  }
  /* Copy the cells, then apply [row index, vals] pairs. */
  function plant(S, cells, list) { const c = Object.assign({}, cells); for (const [i, vals] of list) put(S, c, i, vals); return c; }
  const results = (S, cells, rule) => Array.from({ length: size(S) }, (_, i) => rule(row(S, cells, i), cells));
  const sumCol = (S, cells, h) => results(S, cells, (x) => x[h] || 0).reduce((a, x) => a + x, 0);
  const key = (xs) => xs.map((x) => (typeof x === 'number' ? (Math.round(x * 1e4) / 1e4).toString() : String(x).toLowerCase())).join('|');
  /* The refs in one column that hold a number, for dateCells. */
  const filled = (S, cells, h) => Array.from({ length: size(S) }, (_, i) => at(S, h, i)).filter((ref) => typeof cells[ref] === 'number');

  /*
   * A column answer: the formula goes in the first empty column and is filled down. o.rule(row) mirrors the formula
   * in JavaScript; the build stops if an outcome in o.want never shows up, or if the hidden sheet changes no row.
   */
  function column(o) {
    const { S, rule, want, ask, ans, ...rest } = o;
    const a = results(S, o.base, rule), b = results(S, o.alt, rule);
    for (const w of want || []) if (!a.some((x) => key([x]) === key([w]))) throw new Error(o.title + ': no row gives ' + w);
    if (key(a) === key(b)) throw new Error(o.title + ': the hidden sheet changes no row');
    for (const k of Object.keys(rest)) if (rest[k] === undefined) delete rest[k];
    const first = S.next + S.first, last = S.next + S.last;
    return H.problem(Object.assign({ sheet: S, fill: true, ansLabel: ans,
      statement: ask + `<p>Type the formula in <b>${first}</b> and fill it down to <b>${last}</b>.</p>` }, rest));
  }

  /* Comparison words, and the operators that go with them. */
  const WORD = { '>=': 'at least', '>': 'more than', '<': 'below', '<=': 'at most' };
  const INV = { '>=': '<', '>': '<=', '<': '>=', '<=': '>' };
  const SWAP = { '>=': '<=', '>': '<', '<': '>', '<=': '>=' };
  const WRONG = { '>=': '>', '>': '>=', '<': '<=', '<=': '<' };
  const test = (x, op, k) => (op === '>=' ? x >= k : op === '>' ? x > k : op === '<' ? x < k : x <= k);
  /* The mistake on a limit: > where the rule says "at least", and so on. */
  const edgeTrap = (formula, op, what) => ({ code: 'boundary', title: 'Check the row that sits exactly on the limit',
    plain: `One row is exactly ${what}. "${WORD[op]}" ${op.length === 2 ? 'includes' : 'leaves out'} that row, so the test needs ${op}, not ${WRONG[op]}.`, formula });
  const unpinB1 = (formula) => ({ code: 'unpinned-input', title: 'B1 slid down as you filled', plain: 'Without $ signs, B1 turns into B2, B3 and so on as the formula fills down. Write $B$1 so every row reads the same cell.', formula: formula.replace(/\$B\$1/g, 'B1') });

  /* Invoices with a payment date (a serial) on some of them. */
  function invoiceRows(r, n) {
    return H.invoices(r, n).map((x, i) => Object.assign({}, x, { paidOn: i % 3 === 1 || r.chance(0.3) ? H.toSerial(x.date) + r.int(6, 40) : undefined }));
  }
  /* Facilities with the loan types given, picked at random. */
  const typed = (r, n, types) => H.loans(r, n).map((x) => Object.assign(x, { type: r.pick(types) }));

  /* ======================= Easy, level 1: one IF, one test ======================= */

  P.add({
    key: 'flag-against-a-fixed-limit', difficulty: 'Easy', level: 1, topics: ['IF'], count: 4,
    make(r, v) {
      /* Four lists, each with its own limit and its own words. yes and no hold values that pass and fail the test. */
      const V = [
        () => ({ title: 'Mark order lines Bulk from 30 cartons up', file: 'penjualan_q4.xlsx', ans: 'Order size',
          S: H.table([C.date, C.dist, C.product, C.cartons], H.sales(r, 14)), h: 'Cartons', op: '>=', k: 30, kTxt: '30 cartons', T: 'Bulk', F: 'Regular',
          ask: 'Pak Anton calls an order line <b>Bulk</b> when it has at least 30 cartons, and <b>Regular</b> otherwise.', yes: [36, 44], no: [12, 24] }),
        () => ({ title: 'Mark facilities NPL above 90 days past due', file: 'buku_pinjaman.xlsx', ans: 'Status',
          S: H.table([C.facility, C.dist, C.outstanding, C.dpd], H.loans(r, 14)), h: 'Days past due', op: '>', k: 90, kTxt: '90 days', T: 'NPL', F: 'Performing',
          ask: 'Citra reports a facility as <b>NPL</b> (non-performing) when it is more than 90 days past due, and <b>Performing</b> otherwise.', yes: [120, 150], no: [0, 45] }),
        () => ({ title: 'Call or email reminder from the invoice amount', file: 'tagihan_q4.xlsx', ans: 'Reminder',
          S: H.table([C.invoice, C.dist, C.date, C.amount], H.invoices(r, 14)), h: 'Amount (Rp)', op: '>=', k: 5000000, kTxt: rp(5000000), T: 'Call', F: 'Email',
          ask: 'Bu Lestari phones the distributor about every invoice of at least Rp 5,000,000 and emails about the rest. Mark each invoice <b>Call</b> or <b>Email</b>.', yes: [6240000, 8415000], no: [1980000, 3150000] }),
        () => ({ title: 'Mark stock lines Low below 100 cartons on hand', file: 'stok_gudang.xlsx', ans: 'Stock level',
          S: H.table([C.sku, C.product, C.warehouse, C.onhand], r.shuffle(H.stock(r, 12))), h: 'On hand', op: '<', k: 100, kTxt: '100 cartons', T: 'Low', F: 'OK',
          ask: 'The warehouse marks a stock line <b>Low</b> when its <b>On hand</b> is below 100 cartons, and <b>OK</b> otherwise.', yes: [40, 70], no: [260, 180] })
      ][v]();
      const S = V.S, c = S.cell(V.h, 0), T = `"${V.T}"`, F = `"${V.F}"`;
      /* Row 3 sits exactly on the limit; rows 6 and 9 pass and fail. The hidden sheet swaps rows 6 and 9 over. */
      const base = plant(S, S.cells, [[2, { [V.h]: V.k }], [5, { [V.h]: V.yes[0] }], [8, { [V.h]: V.no[0] }]]);
      const alt = plant(S, base, [[5, { [V.h]: V.no[1] }], [8, { [V.h]: V.yes[1] }]]);
      const solution = `=IF(${c}${V.op}${V.k},${T},${F})`;
      return column({
        title: V.title, file: V.file, S, base, alt, ans: V.ans, want: [V.T, V.F], rule: (x) => (test(x[V.h], V.op, V.k) ? V.T : V.F),
        ask: `<p>${V.ask}</p><p>Test the row's <b>${V.h}</b> and return the word <b>${V.T}</b> or <b>${V.F}</b>.</p>`,
        solution, alts: [`=IF(${c}${INV[V.op]}${V.k},${F},${T})`, `=IF(${V.k}${SWAP[V.op]}${c},${T},${F})`],
        traps: [edgeTrap(`=IF(${c}${WRONG[V.op]}${V.k},${T},${F})`, V.op, V.kTxt)],
        hints: ['IF takes three parts: a test, what to return when the test is true, and what to return when it is false.',
          `"${WORD[V.op]}" is ${V.op} in Excel. Words go inside double quotes: ${T} and ${F}.`,
          `Shape: =IF(${c}${V.op}${V.k}, ${T}, ${F})`],
        explain: `IF tests ${c}${V.op}${V.k} and returns ${T} when it holds, ${F} when it does not. Filling down moves ${c} to each row's own figure.`,
        chips: [c, V.op + V.k, T, F], fnKey: 'IF('
      });
    }
  });

  P.add({
    key: 'compare-two-columns-with-if', difficulty: 'Easy', level: 1, topics: ['IF'], count: 4,
    make(r, v) {
      /* The test compares two cells on the same row. eq makes a row sit exactly on the line; up and down pass and fail. */
      const V = [
        () => ({ title: 'Mark each rep Hit or Miss against the target on the row', file: 'tim_sales.xlsx', ans: 'Result',
          S: H.table([C.rep, C.area, C.target, C.actual], H.reps(r, 10)), a: 'Actual (Rp)', b: 'Target (Rp)', op: '>=', T: 'Hit', F: 'Miss', step: 1000000,
          ask: 'Pak Anton marks a rep <b>Hit</b> when <b>Actual (Rp)</b> is at least <b>Target (Rp)</b>, and <b>Miss</b> otherwise.', plant: [6, -9], flip: [-2, 1.5] }),
        () => ({ title: 'Mark SKUs Reorder when on hand is below the reorder point', file: 'stok_gudang.xlsx', ans: 'Action',
          S: H.table([C.sku, C.product, C.onhand, C.reorder], r.shuffle(H.stock(r, 12))), a: 'On hand', b: 'Reorder point', op: '<', T: 'Reorder', F: 'OK', step: 10,
          ask: 'A SKU needs a new order when its <b>On hand</b> is below its <b>Reorder point</b>. Mark those <b>Reorder</b> and the rest <b>OK</b>.', plant: [-3, 12], flip: [2, -4] }),
        () => {
          const rows = H.monthly(r).map((m) => Object.assign(m, { budget: r6(m.cost * r.pick([0.97, 0.99, 1.02, 1.04])) }));
          return { title: 'Mark months Over or Within the cost budget', file: 'biaya_bulanan.xlsx', ans: 'Budget check',
            S: H.table([C.month, { h: 'Budget (Rp)', k: 'budget', num: true }, C.cost], rows), a: 'Cost (Rp)', b: 'Budget (Rp)', op: '>', T: 'Over', F: 'Within', step: 1000000,
            ask: 'Finance marks a month <b>Over</b> when its <b>Cost (Rp)</b> is more than its <b>Budget (Rp)</b>, and <b>Within</b> otherwise.', plant: [14, -9], flip: [-3, 6] };
        },
        () => {
          const mo = (id, ym) => H.db().sales.reduce((s, x) => s + (x.distributor_id === id && x.date.startsWith(ym) ? x.net : 0), 0);
          const rows = r.sample(DIST(), 10).map((d) => ({ name: d.name, sep: mo(d.id, '2026-09'), oct: mo(d.id, '2026-10') }));
          return { title: 'Mark distributors whose October net grew on September', file: 'penjualan_distributor.xlsx', ans: 'Trend',
            S: H.table([{ h: 'Distributor', k: 'name' }, { h: 'Sep (Rp)', k: 'sep', num: true }, { h: 'Oct (Rp)', k: 'oct', num: true }], rows), a: 'Oct (Rp)', b: 'Sep (Rp)', op: '>', T: 'Grew', F: 'No growth', step: 100000,
            ask: 'Tika marks a distributor <b>Grew</b> when its <b>Oct (Rp)</b> net is more than its <b>Sep (Rp)</b> net, and <b>No growth</b> otherwise.', plant: [42, -61], flip: [-8, 23] };
        }
      ][v]();
      const S = V.S, a = S.cell(V.a, 0), b = S.cell(V.b, 0), T = `"${V.T}"`, F = `"${V.F}"`;
      const off = (d) => (x) => x[V.b] + d * V.step;
      /* Row 2 is level with the other column; rows 5 and 8 pass and fail, and the hidden sheet turns both round. */
      const base = plant(S, S.cells, [[1, { [V.a]: off(0) }], [4, { [V.a]: off(V.plant[0]) }], [7, { [V.a]: off(V.plant[1]) }]]);
      const alt = plant(S, base, [[4, { [V.a]: off(V.flip[0]) }], [7, { [V.a]: off(V.flip[1]) }]]);
      return column({
        title: V.title, file: V.file, S, base, alt, ans: V.ans, want: [V.T, V.F], rule: (x) => (test(x[V.a], V.op, x[V.b]) ? V.T : V.F),
        ask: `<p>${V.ask}</p>`,
        solution: `=IF(${a}${V.op}${b},${T},${F})`, alts: [`=IF(${a}${INV[V.op]}${b},${F},${T})`, `=IF(${b}${SWAP[V.op]}${a},${T},${F})`],
        traps: [edgeTrap(`=IF(${a}${WRONG[V.op]}${b},${T},${F})`, V.op, 'level with ' + V.b)],
        hints: ['An IF test can compare two cells, not only a cell and a number.',
          `"${WORD[V.op]}" is ${V.op}. Put ${V.a}, cell ${a}, on the left and ${V.b}, cell ${b}, on the right.`,
          `Shape: =IF(${a}${V.op}${b}, ${T}, ${F})`],
        explain: `${a}${V.op}${b} compares the two figures on the same row. Neither cell has a $ sign, so both move down together as you fill.`,
        chips: [a, b, T, F], fnKey: 'IF('
      });
    }
  });

  P.add({
    key: 'label-from-a-text-match', difficulty: 'Easy', level: 1, topics: ['IF'], count: 4,
    make(r, v) {
      /* The test checks a text column for one word. pin and swap are the row edits on each sheet. */
      const V = [
        () => ({ title: 'Food or Non-food from the product category', file: 'penjualan_q4.xlsx', ans: 'Group',
          S: H.table([C.date, C.dist, C.product, C.category], H.sales(r, 14)), h: 'Category', word: 'Soap', T: 'Non-food', F: 'Food',
          ask: 'Rasa Kenari reports soap apart from its food lines. Return <b>Non-food</b> when the <b>Category</b> is Soap, and <b>Food</b> otherwise.',
          pin: [[2, inCategory('Soap', 1)], [6, inCategory('Noodles', 2)]], swap: [[2, inCategory('Coffee', 0)], [6, inCategory('Soap', 0)]] }),
        () => ({ title: 'Same-day delivery for Jabodetabek order lines', file: 'pengiriman_q4.xlsx', ans: 'Delivery',
          S: H.table([C.date, C.dist, C.region, C.cartons], H.sales(r, 14)), h: 'Region', word: 'Jabodetabek', T: 'Same day', F: 'Scheduled',
          ask: 'Trucks reach Jabodetabek outlets on the day of the order; every other region waits for a scheduled run. Return <b>Same day</b> when the <b>Region</b> is Jabodetabek, and <b>Scheduled</b> otherwise.',
          pin: [[2, inRegion('Jabodetabek', 1)], [6, inRegion('Jawa Timur', 0)]], swap: [[2, inRegion('Jawa Barat', 0)], [6, inRegion('Jabodetabek', 0)]] }),
        () => ({ title: 'Revolving or amortising from the facility type', file: 'buku_pinjaman.xlsx', ans: 'Repayment',
          S: H.table([C.facility, C.dist, C.type, C.limit], typed(r, 14, ['Working Capital', 'Working Capital', 'Term Loan', 'Investment'])), h: 'Type', word: 'Working Capital', T: 'Revolving', F: 'Amortising',
          ask: 'Dimas splits the book by how it is repaid. A <b>Working Capital</b> line is <b>Revolving</b>; every other type is <b>Amortising</b>.',
          pin: [[1, { Type: 'Working Capital' }], [5, { Type: 'Term Loan' }]], swap: [[1, { Type: 'Investment' }], [5, { Type: 'Working Capital' }]] }),
        () => ({ title: 'Watch label for restructured facilities', file: 'buku_pinjaman.xlsx', ans: 'Watch list',
          S: H.table([C.facility, C.dist, C.outstanding, C.restr], H.loans(r, 14)), h: 'Restructured', word: 'Y', T: 'Watch', F: 'Normal',
          ask: 'Citra keeps every restructured facility on a watch list. Return <b>Watch</b> when <b>Restructured</b> is Y, and <b>Normal</b> otherwise.',
          pin: [[2, { Restructured: 'Y' }], [6, { Restructured: 'N' }], [9, { Restructured: 'N' }]], swap: [[6, { Restructured: 'Y' }], [9, { Restructured: 'Y' }]] })
      ][v]();
      const S = V.S, c = S.cell(V.h, 0), w = `"${V.word}"`, T = `"${V.T}"`, F = `"${V.F}"`;
      const base = plant(S, S.cells, V.pin), alt = plant(S, base, V.swap);
      return column({
        title: V.title, file: V.file, S, base, alt, ans: V.ans, want: [V.T, V.F], rule: (x) => (String(x[V.h]).toLowerCase() === V.word.toLowerCase() ? V.T : V.F),
        ask: `<p>${V.ask}</p>`,
        solution: `=IF(${c}=${w},${T},${F})`, alts: [`=IF(${c}<>${w},${F},${T})`, `=IF(${w}=${c},${T},${F})`],
        hints: ['The = sign also compares text: it is true when the two words are the same.',
          `Text inside a formula goes in double quotes: ${w}. Capitals do not matter to =.`,
          `Shape: =IF(${c}=${w}, ${T}, ${F})`],
        explain: `${c}=${w} is true only on rows whose ${V.h} is ${V.word}, so IF returns ${T} there and ${F} everywhere else.`,
        chips: [c, '=' + w, T, F], fnKey: 'IF('
      });
    }
  });

  /* ======================= Easy, level 2: a number as the result, a limit in B1 ======================= */

  P.add({
    key: 'amount-only-when-a-rule-holds', difficulty: 'Easy', level: 2, topics: ['IF', 'Arithmetic'], count: 4,
    make(r, v) {
      const V = [
        () => {
          const S = H.table([C.rep, C.area, C.target, C.actual], H.reps(r, 10)), t = S.cell('Target (Rp)', 0), a = S.cell('Actual (Rp)', 0);
          const off = (m) => (x) => x['Target (Rp)'] + m * 1000000;
          return { S, title: 'Target bonus of Rp 2,500,000 for reps on target', file: 'tim_sales.xlsx', ans: 'Bonus (Rp)',
            ask: 'A rep earns a flat bonus of <b>Rp 2,500,000</b> when <b>Actual (Rp)</b> is at least <b>Target (Rp)</b>. Otherwise the bonus is 0.',
            pin: [[1, { 'Actual (Rp)': off(0) }], [4, { 'Actual (Rp)': off(5) }], [7, { 'Actual (Rp)': off(-8) }]], swap: [[4, { 'Actual (Rp)': off(-1) }], [7, { 'Actual (Rp)': off(2) }]],
            rule: (x) => (x['Actual (Rp)'] >= x['Target (Rp)'] ? 2500000 : 0),
            solution: `=IF(${a}>=${t},2500000,0)`, alts: [`=IF(${a}<${t},0,2500000)`, `=(${a}>=${t})*2500000`],
            traps: [edgeTrap(`=IF(${a}>${t},2500000,0)`, '>=', 'on target')],
            last: `=IF(${a}>=${t}, 2500000, 0)`, chips: [a, t, '2500000', ',0)'],
            explain: `When ${a}>=${t} holds, IF returns 2500000; when it does not, it returns 0. Numbers go in without quotes, so the column still adds up.` };
        },
        () => {
          const S = H.table([C.date, C.dist, C.region, C.cartons], H.sales(r, 14)), c = S.cell('Cartons', 0);
          return { S, title: 'Freight charge per order line, free from 40 cartons', file: 'pengiriman_q4.xlsx', ans: 'Freight (Rp)',
            ask: 'Delivery is free for an order line of at least 40 cartons. Smaller lines pay <b>Rp 3,000</b> for every carton.',
            pin: [[1, { Cartons: 40 }], [4, { Cartons: 52 }], [7, { Cartons: 18 }]], swap: [[4, { Cartons: 38 }], [7, { Cartons: 41 }]],
            rule: (x) => (x.Cartons >= 40 ? 0 : x.Cartons * 3000),
            solution: `=IF(${c}>=40,0,${c}*3000)`, alts: [`=IF(${c}<40,${c}*3000,0)`, `=(${c}<40)*${c}*3000`],
            traps: [edgeTrap(`=IF(${c}>40,0,${c}*3000)`, '>=', '40 cartons')],
            last: `=IF(${c}>=40, 0, ${c}*3000)`, chips: [c, '>=40', `${c}*3000`],
            explain: `The free case comes first: ${c}>=40 returns 0. Every other line works out ${c}*3000, so the charge follows the cartons on its own row.` };
        },
        () => {
          const S = H.table([C.facility, C.dist, C.outstanding, C.dpd], H.loans(r, 14)), o = S.cell('Outstanding (Rp)', 0), d = S.cell('Days past due', 0);
          return { S, title: 'Late charge of 2% on facilities over 30 days past due', file: 'buku_pinjaman.xlsx', ans: 'Late charge (Rp)',
            ask: 'Bank Kenari adds a one-off late charge of <b>2%</b> of the <b>Outstanding (Rp)</b> to every facility more than 30 days past due. The rest pay nothing.',
            pin: [[1, { 'Days past due': 30 }], [4, { 'Days past due': 60 }], [7, { 'Days past due': 15 }]], swap: [[4, { 'Days past due': 0 }], [7, { 'Days past due': 45 }]],
            rule: (x) => (x['Days past due'] > 30 ? x['Outstanding (Rp)'] * 0.02 : 0),
            solution: `=IF(${d}>30,${o}*0.02,0)`, alts: [`=IF(${d}>30,${o}*2%,0)`, `=IF(${d}<=30,0,${o}*0.02)`, `=(${d}>30)*${o}*0.02`],
            traps: [edgeTrap(`=IF(${d}>=30,${o}*0.02,0)`, '>', '30 days')],
            last: `=IF(${d}>30, ${o}*0.02, 0)`, chips: [d, o, '*0.02', ',0)'],
            explain: `The test reads ${d} and the charge is worked out from ${o}: two different columns in one IF. 0.02 and 2% are the same number.` };
        },
        () => {
          const S = H.table([C.date, C.product, C.category, C.gross], H.sales(r, 14)), k = S.cell('Category', 0), g = S.cell('Gross (Rp)', 0);
          return { S, title: 'Coffee promo rebate of 5% of gross per line', file: 'promo_kopi.xlsx', ans: 'Rebate (Rp)',
            ask: 'This quarter every Coffee line earns a rebate of <b>5%</b> of its <b>Gross (Rp)</b>. Lines in other categories get 0.',
            pin: [[2, inCategory('Coffee', 1)], [6, inCategory('Sambal', 0)], [9, inCategory('Noodles', 1)]], swap: [[6, inCategory('Coffee', 2)], [2, { 'Gross (Rp)': (x) => x['Gross (Rp)'] + 372000 }]],
            rule: (x) => (x.Category === 'Coffee' ? x['Gross (Rp)'] * 0.05 : 0),
            solution: `=IF(${k}="Coffee",${g}*0.05,0)`, alts: [`=IF(${k}<>"Coffee",0,${g}*0.05)`, `=IF(${k}="Coffee",${g}*5%,0)`, `=(${k}="Coffee")*${g}*0.05`],
            traps: [],
            last: `=IF(${k}="Coffee", ${g}*0.05, 0)`, chips: [k, '"Coffee"', `${g}*0.05`, ',0)'],
            explain: `${k}="Coffee" picks the Coffee lines; for those IF works out 5% of ${g}, and every other line gets 0.` };
        }
      ][v]();
      const S = V.S, base = plant(S, S.cells, V.pin), alt = plant(S, base, V.swap);
      return column({
        title: V.title, file: V.file, S, base, alt, ans: V.ans, want: [0], rule: V.rule,
        ask: `<p>${V.ask}</p><p>Return a number on every row, 0 included.</p>`,
        solution: V.solution, alts: V.alts, traps: V.traps,
        hints: ['IF can return a number or a sum as well as a word. Numbers go in without quotes.', 'Write the test first, then the amount when it holds, then the amount when it does not.', 'Shape: ' + V.last],
        explain: V.explain, chips: V.chips, fnKey: 'IF('
      });
    }
  });

  P.add({
    key: 'flag-against-the-limit-in-b1', difficulty: 'Easy', level: 2, topics: ['IF', '$ references'], count: 4,
    make(r, v) {
      /* B1 holds the limit. The hidden sheet moves B1 past row 5's value, so a limit typed into the formula fails. */
      const V = [
        () => ({ title: 'Auto-approve order lines up to the limit in B1', file: 'persetujuan_order.xlsx', ans: 'Approval', label: 'Auto-approval limit (Rp)', b1: 7500000, b1alt: 5000000,
          S: H.table([C.date, C.dist, C.product, C.net], H.sales(r, 14)), h: 'Net (Rp)', op: '<=', T: 'Auto', F: 'Manager', what: 'the limit in B1',
          ask: 'An order line is approved by the system when its <b>Net (Rp)</b> is at most the limit in <b>B1</b>. Bigger lines go to the sales manager. Return <b>Auto</b> or <b>Manager</b>.',
          pin: [[1, 7500000], [4, 6250000], [7, 9800000]], swap: [[9, 4900000]] }),
        () => ({ title: 'Flag facilities above the borrower cap in B1', file: 'buku_pinjaman.xlsx', ans: 'Cap check', label: 'Borrower cap (Rp)', b1: 120000000, b1alt: 100000000,
          S: H.table([C.facility, C.dist, C.type, C.outstanding], H.loans(r, 14)), h: 'Outstanding (Rp)', op: '>', T: 'Over cap', F: 'OK', what: 'the cap in B1',
          ask: 'Bank Kenari caps what one facility may owe. Return <b>Over cap</b> when the <b>Outstanding (Rp)</b> is more than the cap in <b>B1</b>, and <b>OK</b> otherwise.',
          pin: [[1, 120000000], [4, 110000000], [7, 145000000]], swap: [[9, 103500000]] }),
        () => ({ title: 'Order now when stock is at or below the level in B1', file: 'stok_gudang.xlsx', ans: 'Action', label: 'Safety level (cartons)', b1: 150, b1alt: 200,
          S: H.table([C.sku, C.product, C.warehouse, C.onhand], r.shuffle(H.stock(r, 12))), h: 'On hand', op: '<=', T: 'Order now', F: 'Hold', what: 'the level in B1',
          ask: 'The buyer orders a SKU when its <b>On hand</b> is at most the safety level in <b>B1</b>. Return <b>Order now</b> or <b>Hold</b>.',
          pin: [[1, 150], [4, 180], [7, 60], [9, 320]], swap: [[9, 90]] }),
        () => ({ title: 'Coach reps with fewer visits than the minimum in B1', file: 'tim_sales.xlsx', ans: 'Coaching', label: 'Minimum visits', b1: 30, b1alt: 35,
          S: H.table([C.rep, C.area, C.actual, C.visits], H.reps(r, 10)), h: 'Visits', op: '<', T: 'Coach', F: 'OK', what: 'the minimum in B1',
          ask: 'Pak Anton coaches every rep whose <b>Visits</b> are below the minimum in <b>B1</b>. Return <b>Coach</b> or <b>OK</b>.',
          pin: [[1, 30], [4, 32], [7, 21], [8, 48]], swap: [[8, 26]] })
      ][v]();
      const S = V.S, c = S.cell(V.h, 0), T = `"${V.T}"`, F = `"${V.F}"`;
      const base = plant(S, Object.assign({}, S.cells, { A1: V.label, B1: V.b1 }), V.pin.map(([i, x]) => [i, { [V.h]: x }]));
      const alt = plant(S, Object.assign({}, base, { B1: V.b1alt }), V.swap.map(([i, x]) => [i, { [V.h]: x }]));
      const solution = `=IF(${c}${V.op}$B$1,${T},${F})`;
      return column({
        title: V.title, file: V.file, S, base, alt, ans: V.ans, want: [V.T, V.F], rule: (x, cells) => (test(x[V.h], V.op, cells.B1) ? V.T : V.F),
        ask: `<p>${V.ask}</p><p>Point at B1 rather than typing the number, because the limit can change.</p>`,
        solution, alts: [`=IF(${c}${V.op}B$1,${T},${F})`, `=IF(${c}${INV[V.op]}$B$1,${F},${T})`, `=IF($B$1${SWAP[V.op]}${c},${T},${F})`],
        traps: [unpinB1(solution), edgeTrap(`=IF(${c}${WRONG[V.op]}$B$1,${T},${F})`, V.op, 'equal to ' + V.what)],
        hints: [`Compare the row's ${V.h}, cell ${c}, with B1. "${WORD[V.op]}" is ${V.op}.`,
          'Every row reads the same limit, so pin it: $B$1 stays put when you fill down.',
          'Shape: ' + solution],
        explain: `${c} moves down a row at a time while $B$1 stays on the limit. When B1 changes, the whole column updates without editing the formula.`,
        chips: [c, '$B$1', T, F], fnKey: 'IF('
      });
    }
  });

  /* ======================= Easy, level 3: two tests joined, or a verdict on a total ======================= */

  P.add({
    key: 'both-rules-with-and', difficulty: 'Easy', level: 3, topics: ['IF', 'AND'], count: 4,
    make(r, v) {
      const V = [
        () => {
          const S = H.table([C.date, C.dist, C.category, C.cartons], H.sales(r, 14)), k = S.cell('Category', 0), c = S.cell('Cartons', 0);
          return { S, title: 'Noodle promo on lines of at least 20 cartons', file: 'promo_mi.xlsx', ans: 'Promo', T: 'Promo', F: 'Regular',
            ask: 'A line earns the noodle promo when its <b>Category</b> is Noodles <b>and</b> it has at least 20 <b>Cartons</b>. Return <b>Promo</b> or <b>Regular</b>.',
            p: [`${k}="Noodles"`, `${c}>=20`], rule: (x) => x.Category === 'Noodles' && x.Cartons >= 20,
            pin: [[1, { Category: 'Noodles', Cartons: 20 }], [4, { Category: 'Noodles', Cartons: 12 }], [7, { Category: 'Coffee', Cartons: 35 }], [10, { Category: 'Noodles', Cartons: 33 }]],
            swap: [[4, { Cartons: 26 }], [10, { Cartons: 17 }]] };
        },
        () => {
          const S = H.table([C.facility, C.dist, C.type, C.dpd], typed(r, 14, ['Working Capital', 'Working Capital', 'Term Loan'])), t = S.cell('Type', 0), d = S.cell('Days past due', 0);
          return { S, title: 'Renew working capital lines with no days past due', file: 'buku_pinjaman.xlsx', ans: 'Renewal', T: 'Renew', F: 'Review',
            ask: 'A facility renews without a review when its <b>Type</b> is Working Capital <b>and</b> it has 0 <b>Days past due</b>. Return <b>Renew</b> or <b>Review</b>.',
            p: [`${t}="Working Capital"`, `${d}=0`], rule: (x) => x.Type === 'Working Capital' && x['Days past due'] === 0,
            pin: [[1, { Type: 'Working Capital', 'Days past due': 0 }], [4, { Type: 'Working Capital', 'Days past due': 15 }], [7, { Type: 'Term Loan', 'Days past due': 0 }], [10, { Type: 'Working Capital', 'Days past due': 0 }]],
            swap: [[1, { 'Days past due': 5 }], [4, { 'Days past due': 0 }]] };
        },
        () => {
          const S = H.table([C.rep, C.target, C.actual, C.visits], H.reps(r, 10)), t = S.cell('Target (Rp)', 0), a = S.cell('Actual (Rp)', 0), n = S.cell('Visits', 0);
          const off = (m) => (x) => x['Target (Rp)'] + m * 1000000;
          return { S, title: 'Bonus for reps who hit target with at least 40 visits', file: 'tim_sales.xlsx', ans: 'Bonus', T: 'Bonus', F: 'No bonus',
            ask: 'A rep earns the quarter bonus when <b>Actual (Rp)</b> is at least <b>Target (Rp)</b> <b>and</b> the rep made at least 40 <b>Visits</b>. Return <b>Bonus</b> or <b>No bonus</b>.',
            p: [`${a}>=${t}`, `${n}>=40`], rule: (x) => x['Actual (Rp)'] >= x['Target (Rp)'] && x.Visits >= 40,
            pin: [[1, { 'Actual (Rp)': off(4), Visits: 40 }], [4, { 'Actual (Rp)': off(2), Visits: 31 }], [7, { 'Actual (Rp)': off(-5), Visits: 52 }]],
            swap: [[4, { Visits: 43 }], [1, { 'Actual (Rp)': off(-1) }]] };
        },
        () => {
          const rows = invoiceRows(r, 14);
          const S = H.table([C.invoice, C.dist, C.amount, C.paid], rows), p = S.cell('Paid on', 0), m = S.cell('Amount (Rp)', 0);
          const paid = (i, d) => H.toSerial(rows[i].date) + d;
          return { S, title: 'Chase unpaid invoices of at least Rp 5,000,000', file: 'tagihan_q4.xlsx', ans: 'Action', T: 'Chase', F: 'Wait', paidCol: true,
            ask: 'Bu Lestari chases an invoice when <b>Paid on</b> is empty <b>and</b> the <b>Amount (Rp)</b> is at least Rp 5,000,000. Return <b>Chase</b> or <b>Wait</b>.',
            p: [`${p}=""`, `${m}>=5000000`], rule: (x) => x['Paid on'] === undefined && x['Amount (Rp)'] >= 5000000,
            pin: [[1, { 'Paid on': null, 'Amount (Rp)': 5000000 }], [4, { 'Paid on': null, 'Amount (Rp)': 2400000 }], [7, { 'Paid on': paid(7, 24), 'Amount (Rp)': 7800000 }], [10, { 'Paid on': null, 'Amount (Rp)': 9100000 }]],
            swap: [[1, { 'Paid on': paid(1, 30) }], [4, { 'Amount (Rp)': 6000000 }]] };
        }
      ][v]();
      const S = V.S, base = plant(S, S.cells, V.pin), alt = plant(S, base, V.swap);
      const [p1, p2] = V.p, T = `"${V.T}"`, F = `"${V.F}"`;
      return column({
        title: V.title, file: V.file, S, base, alt, ans: V.ans, want: [V.T, V.F], rule: (x) => (V.rule(x) ? V.T : V.F),
        dateCells: V.paidCol ? filled(S, base, 'Paid on') : undefined,
        ask: `<p>${V.ask}</p>`,
        solution: `=IF(AND(${p1},${p2}),${T},${F})`,
        alts: [`=IF(AND(${p2},${p1}),${T},${F})`, `=IF(${p1},IF(${p2},${T},${F}),${F})`, `=IF((${p1})*(${p2}),${T},${F})`],
        hints: ['AND(test1, test2) is TRUE only when both tests are TRUE. Put it where the test of IF goes.',
          `The two tests here are ${p1} and ${p2}.`,
          `Shape: =IF(AND(${p1}, ${p2}), ${T}, ${F})`],
        explain: `AND returns TRUE only on rows that pass both ${p1} and ${p2}; one failed test is enough for ${F}.`,
        chips: [p1, p2, T, F], fnKey: 'AND('
      });
    }
  });

  P.add({
    key: 'either-rule-with-or', difficulty: 'Easy', level: 3, topics: ['IF', 'OR'], count: 3,
    make(r, v) {
      const V = [
        () => {
          const S = H.table([C.facility, C.dist, C.dpd, C.restr], H.loans(r, 14)), d = S.cell('Days past due', 0), x = S.cell('Restructured', 0);
          return { S, title: 'Escalate facilities over 60 days past due or restructured', file: 'buku_pinjaman.xlsx', ans: 'Handling', T: 'Escalate', F: 'Normal',
            ask: 'Citra escalates a facility to the remedial team when it is more than 60 <b>Days past due</b>, <b>or</b> when <b>Restructured</b> is Y. Return <b>Escalate</b> or <b>Normal</b>.',
            p: [`${d}>60`, `${x}="Y"`], rule: (o) => o['Days past due'] > 60 || o.Restructured === 'Y',
            pin: [[1, { 'Days past due': 60, Restructured: 'N' }], [4, { 'Days past due': 75, Restructured: 'N' }], [7, { 'Days past due': 0, Restructured: 'Y' }], [10, { 'Days past due': 15, Restructured: 'N' }]],
            swap: [[1, { 'Days past due': 75 }], [10, { Restructured: 'Y' }]] };
        },
        () => {
          const S = H.table([C.date, C.dist, C.region, C.cartons], H.sales(r, 14)), g = S.cell('Region', 0);
          return { S, title: 'Sea freight for Sumatera or Sulawesi deliveries', file: 'pengiriman_q4.xlsx', ans: 'Freight', T: 'Sea freight', F: 'Road',
            ask: 'Goods for Sumatera and Sulawesi cross by ship. Return <b>Sea freight</b> when the <b>Region</b> is Sumatera <b>or</b> Sulawesi, and <b>Road</b> otherwise.',
            p: [`${g}="Sumatera"`, `${g}="Sulawesi"`], rule: (o) => o.Region === 'Sumatera' || o.Region === 'Sulawesi',
            pin: [[1, inRegion('Sumatera', 1)], [4, inRegion('Sulawesi', 0)], [7, inRegion('Jawa Tengah', 1)], [10, inRegion('Jabodetabek', 0)]],
            swap: [[7, inRegion('Sulawesi', 1)], [1, inRegion('Jawa Barat', 1)]] };
        },
        () => {
          const S = H.table([C.rep, C.target, C.actual, C.visits], H.reps(r, 10)), t = S.cell('Target (Rp)', 0), a = S.cell('Actual (Rp)', 0), n = S.cell('Visits', 0);
          const off = (m) => (x) => x['Target (Rp)'] + m * 1000000;
          return { S, title: 'Review reps with few visits or sales below target', file: 'tim_sales.xlsx', ans: 'Review', T: 'Review', F: 'Fine',
            ask: 'Pak Anton reviews a rep who made fewer than 25 <b>Visits</b>, <b>or</b> whose <b>Actual (Rp)</b> is below <b>Target (Rp)</b>. Return <b>Review</b> or <b>Fine</b>.',
            p: [`${n}<25`, `${a}<${t}`], rule: (o) => o.Visits < 25 || o['Actual (Rp)'] < o['Target (Rp)'],
            pin: [[1, { Visits: 25, 'Actual (Rp)': off(3) }], [4, { Visits: 20, 'Actual (Rp)': off(6) }], [7, { Visits: 48, 'Actual (Rp)': off(-7) }], [8, { Visits: 44, 'Actual (Rp)': off(0) }]],
            swap: [[1, { Visits: 23 }], [7, { 'Actual (Rp)': off(1) }]] };
        }
      ][v]();
      const S = V.S, base = plant(S, S.cells, V.pin), alt = plant(S, base, V.swap);
      const [p1, p2] = V.p, T = `"${V.T}"`, F = `"${V.F}"`;
      return column({
        title: V.title, file: V.file, S, base, alt, ans: V.ans, want: [V.T, V.F], rule: (x) => (V.rule(x) ? V.T : V.F),
        ask: `<p>${V.ask}</p>`,
        solution: `=IF(OR(${p1},${p2}),${T},${F})`,
        alts: [`=IF(OR(${p2},${p1}),${T},${F})`, `=IF(${p1},${T},IF(${p2},${T},${F}))`, `=IF((${p1})+(${p2}),${T},${F})`],
        hints: ['OR(test1, test2) is TRUE when at least one of the tests is TRUE.',
          `The two tests here are ${p1} and ${p2}.`,
          `Shape: =IF(OR(${p1}, ${p2}), ${T}, ${F})`],
        explain: `OR returns TRUE as soon as one test holds, so a row needs only ${p1} or ${p2} to get ${T}. Only rows that fail both get ${F}.`,
        chips: [p1, p2, T, F], fnKey: 'OR('
      });
    }
  });

  P.add({
    key: 'verdict-on-a-column-total', difficulty: 'Easy', level: 3, topics: ['IF', 'SUM'], count: 3,
    make(r, v) {
      /*
       * One answer in B2 from a total and the figure in B1. The hidden sheet moves B1 and one row, each too little to
       * flip the verdict alone, so a typed B1 or a typed total both fail.
       */
      const V = [
        () => {
          const S = H.table([C.rep, C.area, C.target, C.actual], H.reps(r, 10)), h = 'Actual (Rp)';
          const f = Math.floor(sumCol(S, S.cells, h) / 1e6) * 1e6;
          return { S, h, title: 'Team verdict: total actual against the target in B1', file: 'tim_sales.xlsx', label: 'Team target (Rp)', b0: f - 10000000, b1: f - 2000000, d: -5000000,
            op: '>=', T: 'On track', F: 'Behind', ask: 'Pak Anton wants one word for the team. Return <b>On track</b> when the total of <b>Actual (Rp)</b> is at least the team target in <b>B1</b>, and <b>Behind</b> otherwise.' };
        },
        () => {
          const S = H.table([C.facility, C.dist, C.type, C.outstanding], H.loans(r, 12)), h = 'Outstanding (Rp)';
          const f = Math.ceil(sumCol(S, S.cells, h) / 1e6) * 1e6;
          return { S, h, title: 'Loan book verdict: total outstanding against the cap in B1', file: 'buku_pinjaman.xlsx', label: 'Book cap (Rp)', b0: f + 20000000, b1: f + 10000000, d: 15000000,
            op: '<', T: 'Room', F: 'Full', ask: 'The bank caps its lending to Rasa Kenari distributors. Return <b>Room</b> when the total of <b>Outstanding (Rp)</b> is below the cap in <b>B1</b>, and <b>Full</b> otherwise.' };
        },
        () => {
          const S = H.table([C.month, C.sales, C.returns], H.monthly(r)), h = 'Returns (Rp)';
          const f = Math.ceil(sumCol(S, S.cells, h) / 1e6) * 1e6;
          return { S, h, title: 'Returns verdict: the whole year against the budget in B1', file: 'penjualan_bulanan.xlsx', label: 'Returns budget (Rp)', b0: f + 5000000, b1: f + 2000000, d: 4000000,
            op: '<=', T: 'Within', F: 'Over', ask: 'Finance allows a yearly budget for returned goods. Return <b>Within</b> when the total of <b>Returns (Rp)</b> is at most the budget in <b>B1</b>, and <b>Over</b> otherwise.' };
        }
      ][v]();
      const S = V.S, R = S.R(V.h), T = `"${V.T}"`, F = `"${V.F}"`;
      const base = Object.assign({}, S.cells, { A1: V.label, B1: V.b0, A2: 'Verdict' });
      const alt = plant(S, Object.assign({}, base, { B1: V.b1 }), [[3, { [V.h]: (x) => x[V.h] + V.d }]]);
      const verdict = (cells) => (test(sumCol(S, cells, V.h), V.op, cells.B1) ? V.T : V.F);
      if (verdict(base) === verdict(alt)) throw new Error(V.title + ': the hidden sheet keeps the verdict');
      const solution = `=IF(SUM(${R})${V.op}B1,${T},${F})`;
      return H.problem({
        title: V.title, file: V.file, sheet: S, base, alt, answerCell: 'B2',
        statement: `<p>${V.ask}</p><p>Type the formula in <b>B2</b>. Total the column inside the IF and point at B1, so the verdict follows both the data and the figure in B1.</p>`,
        solution, alts: [`=IF(SUM(${R})${INV[V.op]}B1,${F},${T})`, `=IF(B1${SWAP[V.op]}SUM(${R}),${T},${F})`, `=IF(SUM(${S.A(V.h)})${V.op}$B$1,${T},${F})`],
        hints: ['A function can sit inside the test of IF. SUM works out the total first, then IF compares it.',
          `The total is SUM(${R}). Compare it with B1 using ${V.op} ("${WORD[V.op]}").`,
          'Shape: ' + solution],
        explain: `SUM(${R}) totals the column, the comparison with B1 gives TRUE or FALSE, and IF turns that into ${T} or ${F}. A typed total or a typed figure for B1 would miss the next change.`,
        chips: [R, 'SUM(', 'B1', T, F], fnKey: 'IF('
      });
    }
  });

  /* Nudge a ratio off a limit that is not a round binary number (1.1, 0.85, 0.9), where the engine and Excel could
     disagree in the 16th digit. num and den are headers; step is added to num until every row is clear. */
  function clearOf(S, cells, num, den, limits, step) {
    for (let i = 0; i < size(S); i++) {
      const x = row(S, cells, i);
      if (x[den] && limits.some((b) => Math.abs(x[num] / x[den] - b) < 1e-4)) cells[at(S, num, i)] = x[num] + step;
    }
    return cells;
  }

  /* ======================= Medium, level 4: three tiers, IFERROR, NOT ======================= */

  P.add({
    key: 'three-tier-label', difficulty: 'Medium', level: 4, topics: ['Nested IF', 'IF'], count: 4,
    make(r, v) {
      const order = (bad) => [{ code: 'order', title: 'The bigger limit has to come first', plain: 'IF stops at the first test that is true. A row that passes the top limit also passes the lower one, so test the top limit first.', formula: bad }];
      const V = [
        () => {
          const S = H.table([C.date, C.dist, C.product, C.cartons], H.sales(r, 14)), c = S.cell('Cartons', 0);
          return { S, title: 'Small, Medium or Large order lines by cartons', file: 'penjualan_q4.xlsx', ans: 'Size', want: ['Large', 'Medium', 'Small'],
            ask: 'Pak Anton sizes every order line by its <b>Cartons</b>: <b>Large</b> at 40 or more, <b>Medium</b> at 20 to 39, and <b>Small</b> below 20.',
            rule: (x) => (x.Cartons >= 40 ? 'Large' : x.Cartons >= 20 ? 'Medium' : 'Small'),
            pin: [[1, { Cartons: 40 }], [3, { Cartons: 20 }], [5, { Cartons: 12 }], [8, { Cartons: 55 }], [10, { Cartons: 27 }]], swap: [[8, { Cartons: 38 }], [5, { Cartons: 21 }]],
            solution: `=IF(${c}>=40,"Large",IF(${c}>=20,"Medium","Small"))`,
            alts: [`=IFS(${c}>=40,"Large",${c}>=20,"Medium",TRUE,"Small")`, `=IF(${c}<20,"Small",IF(${c}<40,"Medium","Large"))`],
            traps: order(`=IF(${c}>=20,"Medium",IF(${c}>=40,"Large","Small"))`),
            h2: `Test ${c}>=40 first. Only the rows that fail it reach the second test, ${c}>=20.`, chips: [c, '>=40,"Large"', '>=20,"Medium"', '"Small")'] };
        },
        () => {
          const S = H.table([C.rep, C.area, C.target, C.actual], H.reps(r, 10)), t = S.cell('Target (Rp)', 0), a = S.cell('Actual (Rp)', 0);
          const off = (m) => ({ 'Actual (Rp)': (x) => x['Target (Rp)'] + m * 1000000 });
          return { S, title: 'Above, on or below target for each rep', file: 'tim_sales.xlsx', ans: 'Against target', want: ['Above', 'On target', 'Below'],
            ask: 'Return <b>Above</b> when <b>Actual (Rp)</b> is more than <b>Target (Rp)</b>, <b>On target</b> when the two are equal, and <b>Below</b> otherwise.',
            rule: (x) => (x['Actual (Rp)'] > x['Target (Rp)'] ? 'Above' : x['Actual (Rp)'] === x['Target (Rp)'] ? 'On target' : 'Below'),
            pin: [[1, off(0)], [4, off(7)], [7, off(-6)]], swap: [[1, off(1)], [7, off(0)]],
            solution: `=IF(${a}>${t},"Above",IF(${a}=${t},"On target","Below"))`,
            alts: [`=IF(${a}<${t},"Below",IF(${a}=${t},"On target","Above"))`, `=IFS(${a}>${t},"Above",${a}=${t},"On target",TRUE,"Below")`, `=IF(${a}=${t},"On target",IF(${a}>${t},"Above","Below"))`],
            traps: [], h2: `The first test is ${a}>${t}. In its "false" part, a second IF tells ${a}=${t} apart from the rest.`, chips: [a, t, '"Above"', '"On target"', '"Below")'] };
        },
        () => {
          const S = H.table([C.facility, C.dist, C.limit, C.outstanding], H.loans(r, 14)), l = S.cell('Limit (Rp)', 0), o = S.cell('Outstanding (Rp)', 0), u = `${o}/${l}`;
          const lo = (lim, out) => ({ 'Limit (Rp)': lim * 1e6, 'Outstanding (Rp)': out * 1e6 });
          return { S, title: 'High, Medium or Low utilisation per facility', file: 'buku_pinjaman.xlsx', ans: 'Utilisation band', want: ['High', 'Medium', 'Low'],
            ask: 'Utilisation is <b>Outstanding (Rp)</b> divided by <b>Limit (Rp)</b>. Return <b>High</b> at 0.75 (75%) or more, <b>Medium</b> from 0.5 up to 0.75, and <b>Low</b> below 0.5.',
            rule: (x) => { const k = x['Outstanding (Rp)'] / x['Limit (Rp)']; return k >= 0.75 ? 'High' : k >= 0.5 ? 'Medium' : 'Low'; },
            pin: [[1, lo(100, 75)], [3, lo(200, 100)], [5, lo(150, 30)], [8, lo(50, 46)]], swap: [[8, { 'Outstanding (Rp)': 33000000 }], [5, { 'Outstanding (Rp)': 120000000 }]],
            solution: `=IF(${u}>=0.75,"High",IF(${u}>=0.5,"Medium","Low"))`,
            alts: [`=IF(${o}>=0.75*${l},"High",IF(${o}>=0.5*${l},"Medium","Low"))`, `=IFS(${u}>=0.75,"High",${u}>=0.5,"Medium",TRUE,"Low")`, `=IF(${u}>=75%,"High",IF(${u}>=50%,"Medium","Low"))`],
            traps: order(`=IF(${u}>=0.5,"Medium",IF(${u}>=0.75,"High","Low"))`),
            h2: `Work out the ratio inside each test: ${u}>=0.75 first, then ${u}>=0.5.`, chips: [u, '>=0.75,"High"', '>=0.5,"Medium"', '"Low")'] };
        },
        () => {
          const S = H.table([C.rep, C.area, C.visits], H.reps(r, 10)), n = S.cell('Visits', 0);
          return { S, title: 'Busy, Normal or Quiet reps by number of visits', file: 'kunjungan_sales.xlsx', ans: 'Activity', want: ['Busy', 'Normal', 'Quiet'],
            ask: 'Return <b>Busy</b> for a rep with at least 50 <b>Visits</b>, <b>Normal</b> for 30 to 49, and <b>Quiet</b> below 30.',
            rule: (x) => (x.Visits >= 50 ? 'Busy' : x.Visits >= 30 ? 'Normal' : 'Quiet'),
            pin: [[1, { Visits: 50 }], [3, { Visits: 30 }], [5, { Visits: 22 }], [8, { Visits: 57 }]], swap: [[8, { Visits: 44 }], [5, { Visits: 31 }]],
            solution: `=IF(${n}>=50,"Busy",IF(${n}>=30,"Normal","Quiet"))`,
            alts: [`=IFS(${n}>=50,"Busy",${n}>=30,"Normal",TRUE,"Quiet")`, `=IF(${n}<30,"Quiet",IF(${n}<50,"Normal","Busy"))`],
            traps: order(`=IF(${n}>=30,"Normal",IF(${n}>=50,"Busy","Quiet"))`),
            h2: `Test ${n}>=50 first, then ${n}>=30 inside the "false" part.`, chips: [n, '>=50,"Busy"', '>=30,"Normal"', '"Quiet")'] };
        }
      ][v]();
      const S = V.S, base = plant(S, S.cells, V.pin), alt = plant(S, base, V.swap);
      return column({
        title: V.title, file: V.file, S, base, alt, ans: V.ans, want: V.want, rule: V.rule, ask: `<p>${V.ask}</p>`,
        solution: V.solution, alts: V.alts, traps: V.traps,
        hints: ['Three outcomes need two tests. Put a second IF where the first IF\'s "false" answer goes.', V.h2, 'Shape: ' + V.solution],
        explain: 'The outer IF settles the first case; only rows that fail it reach the inner IF, which splits the other two. IFS does the same in one flat list.',
        chips: V.chips, fnKey: 'IF('
      });
    }
  });

  P.add({
    key: 'safe-division-with-iferror', difficulty: 'Medium', level: 4, topics: ['IFERROR'], count: 4,
    make(r, v) {
      const mo = (id, ym, k) => H.db().sales.reduce((s, x) => s + (x.distributor_id === id && x.date.startsWith(ym) ? x[k] : 0), 0);
      const V = [
        () => {
          const rows = r.sample(DIST(), 10).map((d) => ({ name: d.name, cartons: mo(d.id, '2026-10', 'cartons'), net: mo(d.id, '2026-10', 'net') }));
          const S = H.table([{ h: 'Distributor', k: 'name' }, C.cartons, C.net], rows), c = S.cell('Cartons', 0), n = S.cell('Net (Rp)', 0);
          return { S, title: 'Net per carton by distributor, 0 when nothing shipped', file: 'penjualan_oktober.xlsx', ans: 'Net per carton (Rp)', want: [0],
            ask: 'Tika wants October net per carton for each distributor: <b>Net (Rp)</b> divided by <b>Cartons</b>, rounded to whole rupiah with ROUND. A distributor that shipped nothing has 0 cartons, and its division gives #DIV/0!. Show <b>0</b> for it instead.',
            rule: (x) => (x.Cartons === 0 ? 0 : Math.round(x['Net (Rp)'] / x.Cartons)),
            pin: [[2, { Cartons: 0, 'Net (Rp)': 0 }], [6, { Cartons: 0, 'Net (Rp)': 0 }]], swap: [[2, { Cartons: 14, 'Net (Rp)': 1612800 }], [8, { 'Net (Rp)': (x) => x['Net (Rp)'] + 2000000 }]],
            div: `ROUND(${n}/${c},0)`, fb: '0', solution: `=IFERROR(ROUND(${n}/${c},0),0)`,
            alts: [`=ROUND(IFERROR(${n}/${c},0),0)`, `=IF(${c}=0,0,ROUND(${n}/${c},0))`], chips: [`${n}/${c}`, 'ROUND(', ',0)'] };
        },
        () => {
          const S = H.table([C.facility, C.dist, C.limit, C.outstanding], H.loans(r, 14)), l = S.cell('Limit (Rp)', 0), o = S.cell('Outstanding (Rp)', 0);
          return { S, title: 'Utilisation per facility, Closed when the limit is 0', file: 'buku_pinjaman.xlsx', ans: 'Utilisation', want: ['Closed'], precision: 4,
            ask: 'Dimas wants each facility\'s utilisation, <b>Outstanding (Rp)</b> divided by <b>Limit (Rp)</b>, as a decimal. A closed facility has a limit of 0, which gives #DIV/0!. Show the word <b>Closed</b> for it.',
            rule: (x) => (x['Limit (Rp)'] === 0 ? 'Closed' : x['Outstanding (Rp)'] / x['Limit (Rp)']),
            pin: [[3, { 'Limit (Rp)': 0, 'Outstanding (Rp)': 0 }], [9, { 'Limit (Rp)': 0, 'Outstanding (Rp)': 0 }]], swap: [[9, { 'Limit (Rp)': 75000000, 'Outstanding (Rp)': 30000000 }], [5, { 'Outstanding (Rp)': (x) => x['Outstanding (Rp)'] - 5000000 }]],
            div: `${o}/${l}`, fb: '"Closed"', solution: `=IFERROR(${o}/${l},"Closed")`,
            alts: [`=IF(${l}=0,"Closed",${o}/${l})`], chips: [`${o}/${l}`, '"Closed")'] };
        },
        () => {
          const rows = r.sample(DIST(), 10).map((d) => ({ name: d.name, sep: mo(d.id, '2026-09', 'net'), oct: mo(d.id, '2026-10', 'net') }));
          const S = H.table([{ h: 'Distributor', k: 'name' }, { h: 'Sep (Rp)', k: 'sep', num: true }, { h: 'Oct (Rp)', k: 'oct', num: true }], rows), s = S.cell('Sep (Rp)', 0), o = S.cell('Oct (Rp)', 0);
          return { S, title: 'Growth from September to October, New when September is 0', file: 'penjualan_distributor.xlsx', ans: 'Growth', want: ['New'], precision: 4,
            ask: 'Tika wants each distributor\'s growth in net sales: <b>Oct (Rp)</b> divided by <b>Sep (Rp)</b>, minus 1, as a decimal (0.12 means 12% up). A distributor with no September sales gives #DIV/0!. Show the word <b>New</b> for it.',
            rule: (x) => (x['Sep (Rp)'] === 0 ? 'New' : x['Oct (Rp)'] / x['Sep (Rp)'] - 1),
            pin: [[2, { 'Sep (Rp)': 0 }], [7, { 'Sep (Rp)': 0 }]], swap: [[7, { 'Sep (Rp)': 3400000 }], [4, { 'Oct (Rp)': (x) => x['Oct (Rp)'] + 2500000 }]],
            div: `${o}/${s}-1`, fb: '"New"', solution: `=IFERROR(${o}/${s}-1,"New")`,
            alts: [`=IFERROR((${o}-${s})/${s},"New")`, `=IF(${s}=0,"New",${o}/${s}-1)`], chips: [`${o}/${s}-1`, '"New")'] };
        },
        () => {
          const rows = r.shuffle(H.stock(r, 12)).map((x) => Object.assign(x, { perday: r.int(3, 25) }));
          const S = H.table([C.sku, C.product, C.onhand, { h: 'Sold per day', k: 'perday', num: true }], rows), o = S.cell('On hand', 0), d = S.cell('Sold per day', 0);
          return { S, title: 'Days of stock cover, No sales when nothing sells', file: 'stok_gudang.xlsx', ans: 'Days of cover', want: ['No sales'],
            ask: 'Days of cover is <b>On hand</b> divided by <b>Sold per day</b>, rounded to whole days with ROUND. A new SKU has sold nothing yet, so its division gives #DIV/0!. Show the words <b>No sales</b> for it.',
            rule: (x) => (x['Sold per day'] === 0 ? 'No sales' : Math.round(x['On hand'] / x['Sold per day'])),
            pin: [[4, { 'Sold per day': 0 }], [10, { 'Sold per day': 0 }]], swap: [[10, { 'Sold per day': 6 }], [2, { 'On hand': (x) => x['On hand'] + 40 }]],
            div: `ROUND(${o}/${d},0)`, fb: '"No sales"', solution: `=IFERROR(ROUND(${o}/${d},0),"No sales")`,
            alts: [`=IF(${d}=0,"No sales",ROUND(${o}/${d},0))`], chips: [`${o}/${d}`, 'ROUND(', '"No sales")'] };
        }
      ][v]();
      const S = V.S, base = plant(S, S.cells, V.pin), alt = plant(S, base, V.swap);
      return column({
        title: V.title, file: V.file, S, base, alt, ans: V.ans, want: V.want, rule: V.rule, precision: V.precision, ask: `<p>${V.ask}</p>`,
        solution: V.solution, alts: V.alts,
        traps: [{ code: 'div0', title: 'Some rows still divide by zero', plain: `Rows with a 0 underneath show #DIV/0!. Wrap the whole division in IFERROR(..., ${V.fb}).`, formula: '=' + V.div }],
        hints: ['IFERROR(value, fallback) returns the value, or the fallback when the value is an error such as #DIV/0!.',
          `Write the division first and check it on a row that works: =${V.div}. Then wrap it.`,
          'Shape: ' + V.solution],
        explain: `Rows with a number underneath divide as usual. On the others the division is #DIV/0!, and IFERROR returns ${V.fb} in its place. IF(divisor=0, ...) gives the same result and catches only the zero.`,
        chips: V.chips, fnKey: 'IFERROR('
      });
    }
  });

  P.add({
    key: 'rule-turned-around-with-not', difficulty: 'Medium', level: 4, topics: ['NOT', 'IF'], count: 4,
    make(r, v) {
      const V = [
        () => {
          const S = H.table([C.date, C.dist, C.region, C.cartons], H.sales(r, 14)), g = S.cell('Region', 0);
          const JAVA = ['Jabodetabek', 'Jawa Barat', 'Jawa Tengah', 'Jawa Timur'], inner = `OR(${JAVA.map((x) => `${g}="${x}"`).join(',')})`;
          return { S, title: 'Java or outer islands from the region of each line', file: 'pengiriman_q4.xlsx', ans: 'Island group', T: 'Outer islands', F: 'Java', inner,
            ask: 'Four regions are on Java: Jabodetabek, Jawa Barat, Jawa Tengah and Jawa Timur. Return <b>Outer islands</b> when the <b>Region</b> is <b>not</b> one of the four, and <b>Java</b> otherwise.',
            rule: (x) => !JAVA.includes(x.Region),
            pin: [[1, inRegion('Sumatera', 0)], [4, inRegion('Bali Nusra', 1)], [7, inRegion('Jawa Timur', 1)], [10, inRegion('Jabodetabek', 0)]], swap: [[7, inRegion('Sulawesi', 0)], [1, inRegion('Jawa Barat', 1)]],
            alts: [`=IF(${inner},"Java","Outer islands")`, `=IF(AND(${JAVA.map((x) => `${g}<>"${x}"`).join(',')}),"Outer islands","Java")`], chips: [g, '"Jabodetabek"', '"Jawa Barat"', '"Jawa Tengah"', '"Jawa Timur"'] };
        },
        () => {
          const S = H.table([C.rep, C.target, C.actual, C.visits], H.reps(r, 10)), t = S.cell('Target (Rp)', 0), a = S.cell('Actual (Rp)', 0), n = S.cell('Visits', 0);
          /* Not the Easy bonus rule (on target, 40 visits): here the bar is 90% of target and 30 visits. No row sits on
             exactly 90%, where a ratio and a product could differ in the last digit; the visits limit has its edge row. */
          const off = (k, vis) => ({ 'Actual (Rp)': (x) => r5(x['Target (Rp)'] * k), Visits: vis });
          const inner = `AND(${a}>=0.9*${t},${n}>=30)`;
          return { S, title: 'Follow up reps below 90% of target or under 30 visits', file: 'tim_sales.xlsx', ans: 'Rep status', T: 'Follow up', F: 'Fine', inner,
            guard: (cells) => clearOf(S, cells, 'Actual (Rp)', 'Target (Rp)', [0.9], 200000),
            ask: 'A rep is on track when <b>Actual (Rp)</b> is at least 90% of <b>Target (Rp)</b> and <b>Visits</b> are at least 30. Return <b>Follow up</b> when that is <b>not</b> so, and <b>Fine</b> otherwise.',
            rule: (x) => !(x['Actual (Rp)'] >= 0.9 * x['Target (Rp)'] && x.Visits >= 30),
            pin: [[1, off(0.96, 30)], [4, off(1.05, 26)], [7, off(0.8, 47)]], swap: [[4, { Visits: 34 }], [1, off(0.85, 30)]],
            alts: [`=IF(${inner},"Fine","Follow up")`, `=IF(OR(${a}<0.9*${t},${n}<30),"Follow up","Fine")`, `=IF(NOT(AND(${a}/${t}>=0.9,${n}>=30)),"Follow up","Fine")`, `=IF(NOT(AND(${a}>=90%*${t},${n}>=30)),"Follow up","Fine")`],
            chips: [`${a}>=0.9*${t}`, `${n}>=30`, 'NOT(', 'AND('] };
        },
        () => {
          const rows = invoiceRows(r, 14);
          const S = H.table([C.invoice, C.dist, C.amount, C.paid], rows), p = S.cell('Paid on', 0), paid = (i, d) => H.toSerial(rows[i].date) + d;
          return { S, title: 'Paid or open from the payment date column', file: 'tagihan_q4.xlsx', ans: 'Status', T: 'Paid', F: 'Open', inner: `ISBLANK(${p})`, paidCol: true,
            ask: 'Return <b>Paid</b> when <b>Paid on</b> is <b>not</b> empty, and <b>Open</b> otherwise. ISBLANK(cell) is TRUE for an empty cell.',
            rule: (x) => x['Paid on'] !== undefined,
            pin: [[2, { 'Paid on': null }], [5, { 'Paid on': paid(5, 18) }], [8, { 'Paid on': null }]], swap: [[8, { 'Paid on': paid(8, 27) }]],
            alts: [`=IF(ISBLANK(${p}),"Open","Paid")`, `=IF(${p}<>"","Paid","Open")`, `=IF(${p}="","Open","Paid")`], chips: [p, 'ISBLANK(', 'NOT('] };
        },
        () => {
          const S = H.table([C.date, C.dist, C.product, C.cartons], H.sales(r, 14)), c = S.cell('Cartons', 0), inner = `AND(${c}>=5,${c}<=60)`;
          return { S, title: 'Check order lines outside 5 to 60 cartons', file: 'penjualan_q4.xlsx', ans: 'Quantity check', T: 'Check', F: 'OK', inner,
            ask: 'A normal order line has from 5 to 60 <b>Cartons</b>, both ends included. Return <b>Check</b> when a line is <b>not</b> in that range, and <b>OK</b> otherwise.',
            rule: (x) => !(x.Cartons >= 5 && x.Cartons <= 60),
            pin: [[1, { Cartons: 5 }], [3, { Cartons: 60 }], [5, { Cartons: 3 }], [8, { Cartons: 64 }]], swap: [[8, { Cartons: 58 }], [10, { Cartons: 72 }]],
            alts: [`=IF(OR(${c}<5,${c}>60),"Check","OK")`, `=IF(${inner},"OK","Check")`], chips: [`${c}>=5`, `${c}<=60`, 'NOT(', 'AND('] };
        }
      ][v]();
      const S = V.S, g = V.guard || ((c) => c), T = `"${V.T}"`, F = `"${V.F}"`;
      const base = g(plant(S, S.cells, V.pin)), alt = g(plant(S, base, V.swap));
      const solution = `=IF(NOT(${V.inner}),${T},${F})`;
      return column({
        title: V.title, file: V.file, S, base, alt, ans: V.ans, want: [V.T, V.F], rule: (x) => (V.rule(x) ? V.T : V.F),
        dateCells: V.paidCol ? filled(S, base, 'Paid on') : undefined, ask: `<p>${V.ask}</p>`,
        solution, alts: V.alts,
        hints: ['NOT turns TRUE into FALSE and FALSE into TRUE. Write the test for the opposite case, then wrap NOT around it.',
          `The opposite case here is ${V.inner}.`,
          'Shape: ' + solution],
        explain: `${V.inner} is TRUE for the rows that should get ${F}; NOT flips it, so IF returns ${T} for every other row.`,
        chips: V.chips.concat([T, F]), fnKey: 'NOT('
      });
    }
  });

  /* ======================= Medium, level 5: IFS, SWITCH, a number inside a label, three rules at once ======================= */

  P.add({
    key: 'grade-with-ifs', difficulty: 'Medium', level: 5, topics: ['IFS'], count: 4,
    make(r, v) {
      const V = [
        () => {
          const S = H.table([C.facility, C.dist, C.outstanding, C.dpd], H.loans(r, 14)), d = S.cell('Days past due', 0), dd = (x) => ({ 'Days past due': x });
          return { S, title: 'Collectibility grade Kol 1 to 5 from days past due', file: 'buku_pinjaman.xlsx', ans: 'Kol', want: [1, 2, 3, 4, 5],
            ask: 'Bank Kenari grades each facility by its <b>Days past due</b>: <b>1</b> at 0 days, <b>2</b> from 1 to 90, <b>3</b> from 91 to 120, <b>4</b> from 121 to 180 and <b>5</b> above 180. Return the grade as a number.',
            rule: (x) => kol(x['Days past due']),
            pin: [[0, dd(0)], [2, dd(90)], [4, dd(120)], [6, dd(180)], [8, dd(200)], [10, dd(95)]], swap: [[10, dd(125)], [2, dd(95)]],
            solution: `=IFS(${d}=0,1,${d}<=90,2,${d}<=120,3,${d}<=180,4,TRUE,5)`,
            alts: [`=IF(${d}=0,1,IF(${d}<=90,2,IF(${d}<=120,3,IF(${d}<=180,4,5))))`, `=IFS(${d}>180,5,${d}>120,4,${d}>90,3,${d}>0,2,TRUE,1)`],
            h2: `Go up the grades: ${d}=0 gives 1, then ${d}<=90 gives 2, and so on. "90 days" belongs to grade 2, so use <=.`, chips: [d, '=0,1', '<=90,2', '<=120,3', '<=180,4', 'TRUE,5)'] };
        },
        () => {
          const S = H.table([C.rep, C.area, C.target, C.actual], H.reps(r, 10)), t = S.cell('Target (Rp)', 0), a = S.cell('Actual (Rp)', 0), k = `${a}/${t}`;
          const f = (m) => ({ 'Actual (Rp)': (x) => r5(x['Target (Rp)'] * m) });
          return { S, title: 'Rep grade A to D from target achievement', file: 'tim_sales.xlsx', ans: 'Grade', want: ['A', 'B', 'C', 'D'], exactCase: true,
            guard: (cells) => clearOf(S, cells, 'Actual (Rp)', 'Target (Rp)', [1.1, 0.85], 200000),
            ask: 'Achievement is <b>Actual (Rp)</b> divided by <b>Target (Rp)</b>. Grade <b>A</b> at 1.1 (110%) or more, <b>B</b> from 1 up to 1.1, <b>C</b> from 0.85 up to 1, and <b>D</b> below 0.85. Use capital letters.',
            rule: (x) => { const q = x['Actual (Rp)'] / x['Target (Rp)']; return q >= 1.1 ? 'A' : q >= 1 ? 'B' : q >= 0.85 ? 'C' : 'D'; },
            pin: [[1, f(1)], [3, f(1.2)], [5, f(0.9)], [7, f(0.7)]], swap: [[5, f(0.8)], [7, f(1)]],
            solution: `=IFS(${k}>=1.1,"A",${k}>=1,"B",${k}>=0.85,"C",TRUE,"D")`,
            alts: [`=IF(${k}>=1.1,"A",IF(${k}>=1,"B",IF(${k}>=0.85,"C","D")))`, `=IFS(${a}>=1.1*${t},"A",${a}>=${t},"B",${a}>=0.85*${t},"C",TRUE,"D")`],
            h2: `Each test works out the ratio ${k} and compares it: >=1.1 first, then >=1, then >=0.85.`, chips: [k, '>=1.1,"A"', '>=1,"B"', '>=0.85,"C"', 'TRUE,"D")'] };
        },
        () => {
          const tot = {};
          H.db().sales.forEach((s) => { tot[s.distributor] = (tot[s.distributor] || 0) + s.net; });
          const S = H.table([{ h: 'Distributor', k: 'name' }, C.region, { h: 'Net May to Oct (Rp)', k: 'net', num: true }], r.sample(DIST(), 12).map((d) => ({ name: d.name, region: d.region, net: tot[d.name] })));
          const h = 'Net May to Oct (Rp)', n = S.cell(h, 0), nn = (x) => ({ [h]: x });
          return { S, title: 'Segment distributors by six-month net sales', file: 'segmen_distributor.xlsx', ans: 'Segment', want: ['Key', 'Core', 'Growth', 'Small'],
            ask: 'Tika sorts distributors by their <b>Net May to Oct (Rp)</b>: <b>Key</b> at Rp 300,000,000 or more, <b>Core</b> from Rp 200,000,000, <b>Growth</b> from Rp 150,000,000, and <b>Small</b> below that.',
            rule: (x) => (x[h] >= 3e8 ? 'Key' : x[h] >= 2e8 ? 'Core' : x[h] >= 1.5e8 ? 'Growth' : 'Small'),
            pin: [[1, nn(300000000)], [4, nn(128400000)], [7, nn(192600000)], [9, nn(236800000)]], swap: [[7, nn(204100000)], [4, nn(151200000)]],
            solution: `=IFS(${n}>=300000000,"Key",${n}>=200000000,"Core",${n}>=150000000,"Growth",TRUE,"Small")`,
            alts: [`=IF(${n}>=300000000,"Key",IF(${n}>=200000000,"Core",IF(${n}>=150000000,"Growth","Small")))`, `=IFS(${n}<150000000,"Small",${n}<200000000,"Growth",${n}<300000000,"Core",TRUE,"Key")`],
            h2: `Start with the top band, ${n}>=300000000, and work down. Type the limits without commas or dots.`, chips: [n, '>=300000000,"Key"', '>=200000000,"Core"', '>=150000000,"Growth"', 'TRUE,"Small")'] };
        },
        () => {
          const S = H.table([C.month, C.sales, C.cost], H.monthly(r)), s = S.cell('Sales (Rp)', 0), c = S.cell('Cost (Rp)', 0), m = `(${s}-${c})/${s}`;
          const f = (k) => ({ 'Cost (Rp)': (x) => r6(x['Sales (Rp)'] * k) });
          const margin = (x) => (x['Sales (Rp)'] - x['Cost (Rp)']) / x['Sales (Rp)'];
          return { S, title: 'Strong, Fair or Weak gross margin for each month', file: 'penjualan_bulanan.xlsx', ans: 'Margin grade', want: ['Strong', 'Fair', 'Weak'],
            guard: (cells) => { for (let i = 0; i < size(S); i++) { const x = row(S, cells, i); if ([0.4, 0.35].some((b) => Math.abs(margin(x) - b) < 1e-4)) cells[at(S, 'Cost (Rp)', i)] = x['Cost (Rp)'] + 2000000; } return cells; },
            ask: 'Gross margin is <b>Sales (Rp)</b> minus <b>Cost (Rp)</b>, divided by <b>Sales (Rp)</b>. Return <b>Strong</b> at 0.4 (40%) or more, <b>Fair</b> from 0.35 up to 0.4, and <b>Weak</b> below 0.35.',
            rule: (x) => (margin(x) >= 0.4 ? 'Strong' : margin(x) >= 0.35 ? 'Fair' : 'Weak'),
            pin: [[1, f(0.57)], [4, f(0.62)], [7, f(0.68)]], swap: [[4, f(0.67)], [7, f(0.58)]],
            solution: `=IFS(${m}>=0.4,"Strong",${m}>=0.35,"Fair",TRUE,"Weak")`,
            alts: [`=IF(${m}>=0.4,"Strong",IF(${m}>=0.35,"Fair","Weak"))`, `=IFS(1-${c}/${s}>=0.4,"Strong",1-${c}/${s}>=0.35,"Fair",TRUE,"Weak")`],
            h2: `The margin is ${m}; the brackets make the subtraction happen before the division.`, chips: [m, '>=0.4,"Strong"', '>=0.35,"Fair"', 'TRUE,"Weak")'] };
        }
      ][v]();
      const S = V.S, g = V.guard || ((c) => c);
      const base = g(plant(S, S.cells, V.pin)), alt = g(plant(S, base, V.swap));
      return column({
        title: V.title, file: V.file, S, base, alt, ans: V.ans, want: V.want, rule: V.rule, exactCase: V.exactCase, ask: `<p>${V.ask}</p>`,
        solution: V.solution, alts: V.alts,
        hints: ['IFS takes pairs: a test, then the result for it. It returns the result of the first test that is TRUE.',
          V.h2 + ' End with TRUE and the last result, so every row gets an answer.',
          'Shape: ' + V.solution],
        explain: 'IFS reads its tests from left to right and stops at the first TRUE. Taking the bands in order from one end means each test needs one limit only, and the final TRUE catches every row the other tests missed.',
        chips: V.chips, fnKey: 'IFS('
      });
    }
  });

  P.add({
    key: 'map-codes-with-switch', difficulty: 'Medium', level: 5, topics: ['SWITCH'], count: 4,
    make(r, v) {
      const sw = (c, pairs, def) => `=SWITCH(${c},${pairs.map(([k, x]) => `${typeof k === 'number' ? k : `"${k}"`},"${x}"`).join(',')}${def ? `,"${def}"` : ''})`;
      const ifs = (c, pairs, def) => `=IFS(${pairs.map(([k, x]) => `${c}=${typeof k === 'number' ? k : `"${k}"`},"${x}"`).join(',')}${def ? `,TRUE,"${def}"` : ''})`;
      const V = [
        () => {
          const S = H.table([C.date, C.dist, C.region, C.cartons], H.sales(r, 14)), g = S.cell('Region', 0);
          const pairs = [['Jabodetabek', 'Gudang Bekasi'], ['Jawa Barat', 'Gudang Bandung'], ['Jawa Tengah', 'Gudang Semarang'], ['Sumatera', 'Gudang Medan']], def = 'Gudang Surabaya';
          return { S, c: g, pairs, def, title: 'Serving warehouse for each region, with a default', file: 'pengiriman_q4.xlsx', ans: 'Warehouse', h: 'Region',
            ask: 'Each region is served by one warehouse: Jabodetabek from <b>Gudang Bekasi</b>, Jawa Barat from <b>Gudang Bandung</b>, Jawa Tengah from <b>Gudang Semarang</b> and Sumatera from <b>Gudang Medan</b>. Every other region is served from <b>Gudang Surabaya</b>. Return the warehouse for each line from its <b>Region</b>.',
            pin: [[0, inRegion('Jabodetabek', 0)], [2, inRegion('Jawa Barat', 0)], [4, inRegion('Jawa Tengah', 0)], [6, inRegion('Sumatera', 0)], [8, inRegion('Bali Nusra', 0)], [10, inRegion('Sulawesi', 1)]],
            swap: [[8, inRegion('Jawa Tengah', 1)], [2, inRegion('Sulawesi', 0)]],
            more: [`=IF(${g}="Jabodetabek","Gudang Bekasi",IF(${g}="Jawa Barat","Gudang Bandung",IF(${g}="Jawa Tengah","Gudang Semarang",IF(${g}="Sumatera","Gudang Medan","Gudang Surabaya"))))`],
            h2: 'Four pairs of region and warehouse, then the default, Gudang Surabaya, on its own at the end.' };
        },
        () => {
          const S = H.table([C.facility, C.dist, C.outstanding, { h: 'Kol', k: 'kol', num: true }], H.loans(r, 14).map((x) => Object.assign(x, { kol: kol(x.dpd) }))), k = S.cell('Kol', 0);
          const NAMES = ['Current', 'Special mention', 'Substandard', 'Doubtful', 'Loss'], pairs = NAMES.map((x, i) => [i + 1, x]), kk = (x) => ({ Kol: x });
          return { S, c: k, pairs, title: 'Kol number to its collectibility name', file: 'buku_pinjaman.xlsx', ans: 'Collectibility', h: 'Kol',
            ask: 'Each <b>Kol</b> grade has a name: <b>1</b> Current, <b>2</b> Special mention, <b>3</b> Substandard, <b>4</b> Doubtful and <b>5</b> Loss. Return the name for every facility.',
            pin: [[1, kk(1)], [3, kk(2)], [5, kk(3)], [7, kk(4)], [9, kk(5)]], swap: [[3, kk(3)], [5, kk(4)]],
            more: [`=CHOOSE(${k},${NAMES.map((x) => `"${x}"`).join(',')})`],
            h2: 'The grades are numbers, so the values to match go in without quotes: 1,"Current",2,"Special mention" and so on.' };
        },
        () => {
          const S = H.table([C.facility, C.dist, C.type, C.limit], typed(r, 14, ['Working Capital', 'Term Loan', 'Trade Finance', 'Investment'])), t = S.cell('Type', 0);
          const pairs = [['Working Capital', 'WC'], ['Term Loan', 'TL'], ['Trade Finance', 'TF'], ['Investment', 'INV']], tt = (x) => ({ Type: x });
          return { S, c: t, pairs, exactCase: true, title: 'Short code in capitals for each loan type', file: 'buku_pinjaman.xlsx', ans: 'Code', h: 'Type',
            ask: 'Dimas needs a short code for each loan <b>Type</b>, in capitals: Working Capital <b>WC</b>, Term Loan <b>TL</b>, Trade Finance <b>TF</b> and Investment <b>INV</b>.',
            pin: [[1, tt('Working Capital')], [3, tt('Term Loan')], [5, tt('Trade Finance')], [7, tt('Investment')]], swap: [[7, tt('Trade Finance')], [1, tt('Term Loan')]],
            more: [`=IF(${t}="Working Capital","WC",IF(${t}="Term Loan","TL",IF(${t}="Trade Finance","TF","INV")))`],
            h2: 'Four pairs of type and code. Every type in the list has its own pair, so no default is needed.' };
        },
        () => {
          const S = H.table([C.sku, C.product, C.category, C.onhand], r.shuffle(H.stock(r, 12))), k = S.cell('Category', 0);
          const pairs = [['Noodles', 'Dry store'], ['Coffee', 'Dry store'], ['Sambal', 'Cool room'], ['Soap', 'Non-food store']];
          const find = (cat) => results(S, S.cells, (x) => x.Category).indexOf(cat);
          return { S, c: k, pairs, title: 'Storage area for each product category', file: 'stok_gudang.xlsx', ans: 'Store in', h: 'Category',
            ask: 'Noodles and Coffee go to the <b>Dry store</b>, Sambal to the <b>Cool room</b>, and Soap to the <b>Non-food store</b>. Return the storage area for every stock line from its <b>Category</b>.',
            pin: [], swap: [[find('Noodles'), inCategory('Sambal', 2)], [find('Soap'), inCategory('Coffee', 1)]],
            more: [`=SWITCH(${k},"Sambal","Cool room","Soap","Non-food store","Dry store")`, `=IFS(OR(${k}="Noodles",${k}="Coffee"),"Dry store",${k}="Sambal","Cool room",${k}="Soap","Non-food store")`],
            h2: 'Two categories share Dry store, so that result appears twice in the list. (Putting it last as the default also works.)' };
        }
      ][v]();
      const S = V.S, base = plant(S, S.cells, V.pin), alt = plant(S, base, V.swap);
      const look = (x) => { const hit = V.pairs.find(([k]) => k === x[V.h]); return hit ? hit[1] : V.def; };
      const solution = sw(V.c, V.pairs, V.def);
      return column({
        title: V.title, file: V.file, S, base, alt, ans: V.ans, want: Array.from(new Set(V.pairs.map((p) => p[1]).concat(V.def ? [V.def] : []))), rule: look, exactCase: V.exactCase,
        ask: `<p>${V.ask}</p>`,
        solution, alts: [ifs(V.c, V.pairs, V.def)].concat(V.more),
        hints: ['SWITCH(value, match1, result1, match2, result2, ..., default) compares one value with a list of matches and returns the result paired with the first one that fits.',
          V.h2, 'Shape: ' + solution],
        explain: `SWITCH reads ${V.c} once and walks the list of pairs. It is shorter than a nested IF when every test is "equals"${V.def ? ', and the last lone argument catches every value not in the list' : ''}.`,
        chips: [V.c].concat(V.pairs.slice(0, 2).map(([k, x]) => `${typeof k === 'number' ? k : `"${k}"`},"${x}"`)), fnKey: 'SWITCH('
      });
    }
  });

  P.add({
    key: 'status-with-a-number-inside', difficulty: 'Medium', level: 5, topics: ['IF', 'Text join'], count: 4,
    make(r, v) {
      const V = [
        () => {
          const S = H.table([C.sku, C.product, C.onhand, C.reorder], r.shuffle(H.stock(r, 12))), o = S.cell('On hand', 0), p = S.cell('Reorder point', 0);
          const gap = (d) => ({ 'On hand': (x) => x['Reorder point'] + d });
          return { S, title: 'OK, or how many cartons to order, for each SKU', file: 'stok_gudang.xlsx', ans: 'Order note', want: ['OK'],
            ask: 'Return <b>OK</b> when <b>On hand</b> is at least the <b>Reorder point</b>. Otherwise return a note with the gap, such as <b>Order 30 cartons</b>: the word Order, a space, Reorder point minus On hand, a space and the word cartons.',
            rule: (x) => (x['On hand'] >= x['Reorder point'] ? 'OK' : 'Order ' + (x['Reorder point'] - x['On hand']) + ' cartons'),
            pin: [[1, gap(0)], [4, gap(-30)], [7, gap(90)]], swap: [[4, gap(-50)], [7, gap(-20)]],
            solution: `=IF(${o}>=${p},"OK","Order "&(${p}-${o})&" cartons")`,
            alts: [`=IF(${o}<${p},"Order "&(${p}-${o})&" cartons","OK")`, `=IF(${o}>=${p},"OK",CONCAT("Order ",${p}-${o}," cartons"))`, `=IF(${o}>=${p},"OK","Order "&${p}-${o}&" cartons")`],
            h2: `The note is "Order "&(${p}-${o})&" cartons". The spaces sit inside the quotes.`, chips: [o, p, '"Order "&', `(${p}-${o})`, '&" cartons"'] };
        },
        () => {
          const S = H.table([C.facility, C.dist, C.outstanding, C.dpd], H.loans(r, 14)), d = S.cell('Days past due', 0), dd = (x) => ({ 'Days past due': x });
          return { S, title: 'Current, or the days late, for each facility', file: 'buku_pinjaman.xlsx', ans: 'Status note', want: ['Current'],
            ask: 'Return <b>Current</b> when <b>Days past due</b> is 0. Otherwise return the number, a space and the words days late, such as <b>45 days late</b>.',
            rule: (x) => (x['Days past due'] === 0 ? 'Current' : x['Days past due'] + ' days late'),
            pin: [[1, dd(0)], [4, dd(45)], [7, dd(0)]], swap: [[7, dd(15)], [4, dd(75)]],
            solution: `=IF(${d}=0,"Current",${d}&" days late")`,
            alts: [`=IF(${d}>0,${d}&" days late","Current")`, `=IF(${d}=0,"Current",CONCAT(${d}," days late"))`],
            h2: `The note is ${d}&" days late", with a space at the start of the quoted part.`, chips: [d, '"Current"', '&" days late"'] };
        },
        () => {
          const S = H.table([C.rep, C.area, C.visits], H.reps(r, 10)), n = S.cell('Visits', 0), vv = (x) => ({ Visits: x });
          return { S, title: 'Done, or how many more visits each rep needs', file: 'kunjungan_sales.xlsx', ans: 'Visit note', want: ['Done'],
            guard: (cells) => { for (let i = 0; i < size(S); i++) if (cells[at(S, 'Visits', i)] === 39) cells[at(S, 'Visits', i)] = 38; return cells; },
            ask: 'Each rep should make 40 <b>Visits</b> this month. Return <b>Done</b> at 40 or more. Below 40, return how many are missing, such as <b>Needs 7 more visits</b>.',
            rule: (x) => (x.Visits >= 40 ? 'Done' : 'Needs ' + (40 - x.Visits) + ' more visits'),
            pin: [[1, vv(40)], [4, vv(33)], [7, vv(51)]], swap: [[4, vv(36)], [7, vv(37)]],
            solution: `=IF(${n}>=40,"Done","Needs "&(40-${n})&" more visits")`,
            alts: [`=IF(${n}<40,"Needs "&(40-${n})&" more visits","Done")`, `=IF(${n}>=40,"Done",CONCAT("Needs ",40-${n}," more visits"))`],
            h2: `The note is "Needs "&(40-${n})&" more visits".`, chips: [n, '"Done"', '"Needs "&', `(40-${n})`, '&" more visits"'] };
        },
        () => {
          const rows = invoiceRows(r, 14), S = H.table([C.invoice, C.dist, C.date, C.paid], rows), dt = S.cell('Date', 0), p = S.cell('Paid on', 0), pd = (d) => ({ 'Paid on': d === null ? null : (x) => x.Date + d });
          return { S, title: 'Unpaid, or how many days each invoice took to pay', file: 'tagihan_q4.xlsx', ans: 'Payment note', want: ['Unpaid'], paidCol: true,
            ask: 'Return <b>Unpaid</b> when <b>Paid on</b> is empty. Otherwise return the days from <b>Date</b> to <b>Paid on</b> inside a note, such as <b>Paid in 24 days</b>.',
            rule: (x) => (x['Paid on'] === undefined ? 'Unpaid' : 'Paid in ' + (x['Paid on'] - x.Date) + ' days'),
            pin: [[2, pd(null)], [5, pd(24)], [8, pd(null)]], swap: [[8, pd(33)], [5, pd(26)]],
            solution: `=IF(${p}="","Unpaid","Paid in "&(${p}-${dt})&" days")`,
            alts: [`=IF(ISBLANK(${p}),"Unpaid","Paid in "&(${p}-${dt})&" days")`, `=IF(${p}<>"","Paid in "&(${p}-${dt})&" days","Unpaid")`],
            h2: `Two dates subtract into a number of days: ${p}-${dt}. Put it in brackets between "Paid in " and " days".`, chips: [p, dt, '"Unpaid"', '"Paid in "&', `(${p}-${dt})`, '&" days"'] };
        }
      ][v]();
      const S = V.S, g = V.guard || ((c) => c);
      const base = g(plant(S, S.cells, V.pin)), alt = g(plant(S, base, V.swap));
      return column({
        title: V.title, file: V.file, S, base, alt, ans: V.ans, want: V.want, rule: V.rule,
        dateCells: V.paidCol ? filled(S, base, 'Paid on') : undefined, ask: `<p>${V.ask}</p>`,
        solution: V.solution, alts: V.alts,
        hints: ['The & sign joins text and numbers into one piece of text: "Order "&30 gives Order 30.', V.h2, 'Shape: ' + V.solution],
        explain: 'IF picks the plain word or the note. Inside the note, & joins the quoted words and the worked-out number, so the label changes when the figures do.',
        chips: V.chips, fnKey: 'IF('
      });
    }
  });

  P.add({
    key: 'all-three-rules-with-and', difficulty: 'Medium', level: 5, topics: ['AND', 'IF'], count: 4,
    make(r, v) {
      const V = [
        () => {
          const S = H.table([C.facility, C.limit, C.outstanding, C.dpd, C.restr], H.loans(r, 14)), l = S.cell('Limit (Rp)', 0), o = S.cell('Outstanding (Rp)', 0), d = S.cell('Days past due', 0), x = S.cell('Restructured', 0);
          const f = (lim, out, dpd, y) => ({ 'Limit (Rp)': lim * 1e6, 'Outstanding (Rp)': out * 1e6, 'Days past due': dpd, Restructured: y });
          return { S, title: 'Offer a top-up: current, not restructured and 75% drawn', file: 'buku_pinjaman.xlsx', ans: 'Top-up', T: 'Offer top-up', F: 'Not now',
            ask: 'Dimas offers a bigger line to borrowers who use theirs and pay on time. Return <b>Offer top-up</b> when all three hold: 0 <b>Days past due</b>, <b>Restructured</b> is N, and <b>Outstanding (Rp)</b> is at least 75% of <b>Limit (Rp)</b>. Otherwise return <b>Not now</b>.',
            p: [`${d}=0`, `${x}="N"`, `${o}>=0.75*${l}`],
            rule: (q) => q['Days past due'] === 0 && q.Restructured === 'N' && q['Outstanding (Rp)'] >= 0.75 * q['Limit (Rp)'],
            pin: [[1, f(100, 75, 0, 'N')], [3, f(150, 138, 0, 'Y')], [5, f(200, 184, 15, 'N')], [7, f(50, 20, 0, 'N')], [9, f(200, 170, 0, 'N')]], swap: [[7, { 'Outstanding (Rp)': 40000000 }], [9, { 'Days past due': 5 }]],
            more: [`=IF(AND(${d}=0,${x}="N",${o}/${l}>=0.75),"Offer top-up","Not now")`, `=IF(AND(${d}=0,NOT(${x}="Y"),${o}>=0.75*${l}),"Offer top-up","Not now")`] };
        },
        () => {
          const S = H.table([C.rep, C.target, C.actual, C.visits, C.start], H.reps(r, 10)), t = S.cell('Target (Rp)', 0), a = S.cell('Actual (Rp)', 0), n = S.cell('Visits', 0), s = S.cell('Start date', 0);
          const f = (m, vis, iso) => ({ 'Actual (Rp)': (x) => x['Target (Rp)'] + m * 1000000, Visits: vis, 'Start date': H.toSerial(iso) });
          return { S, title: 'Promotion shortlist: target, 40 visits, joined by the date in B1', file: 'tim_sales.xlsx', ans: 'Shortlist', T: 'Shortlist', F: 'Not yet',
            top: { A1: 'Joined on or before', B1: H.toSerial('2024-12-31') }, b1alt: H.toSerial('2025-06-30'),
            ask: 'Pak Anton shortlists reps for promotion when all three hold: <b>Actual (Rp)</b> at least <b>Target (Rp)</b>, at least 40 <b>Visits</b>, and a <b>Start date</b> on or before the date in <b>B1</b>. Return <b>Shortlist</b> or <b>Not yet</b>. Point at B1, because the cut-off date can change.',
            p: [`${a}>=${t}`, `${n}>=40`, `${s}<=$B$1`],
            rule: (q, cells) => q['Actual (Rp)'] >= q['Target (Rp)'] && q.Visits >= 40 && q['Start date'] <= cells.B1,
            pin: [[1, f(3, 40, '2023-04-10')], [3, f(2, 46, '2025-03-03')], [5, f(-4, 52, '2022-08-15')], [7, f(1, 35, '2021-01-11')], [9, f(5, 44, '2024-12-31')]], swap: [[7, { Visits: 41 }]],
            more: [`=IF(AND(${a}>=${t},${n}>=40,${s}<=B$1),"Shortlist","Not yet")`, `=IF(AND(${a}>=${t},${n}>=40,$B$1>=${s}),"Shortlist","Not yet")`] };
        },
        () => {
          const S = H.table([C.date, C.dist, C.region, C.cartons, C.net], H.sales(r, 14)), g = S.cell('Region', 0), c = S.cell('Cartons', 0), nt = S.cell('Net (Rp)', 0);
          const f = (reg, k, n, net) => Object.assign(inRegion(reg, k), { Cartons: n, 'Net (Rp)': net });
          return { S, title: 'Free delivery: Jabodetabek, 20 cartons and Rp 3,000,000', file: 'pengiriman_q4.xlsx', ans: 'Delivery fee', T: 'Free', F: 'Charge',
            ask: 'An order line ships free when all three hold: the <b>Region</b> is Jabodetabek, it has at least 20 <b>Cartons</b>, and its <b>Net (Rp)</b> is at least Rp 3,000,000. Return <b>Free</b> or <b>Charge</b>.',
            p: [`${g}="Jabodetabek"`, `${c}>=20`, `${nt}>=3000000`],
            rule: (q) => q.Region === 'Jabodetabek' && q.Cartons >= 20 && q['Net (Rp)'] >= 3000000,
            pin: [[1, f('Jabodetabek', 0, 20, 3000000)], [3, f('Jabodetabek', 1, 22, 2450000)], [5, f('Jabodetabek', 0, 14, 3400000)], [7, f('Jawa Barat', 0, 35, 6200000)], [9, f('Jabodetabek', 1, 31, 5300000)]],
            swap: [[3, { 'Net (Rp)': 3150000 }], [9, { Cartons: 18, 'Net (Rp)': 3078000 }]],
            more: [`=IF(${g}<>"Jabodetabek","Charge",IF(AND(${c}>=20,${nt}>=3000000),"Free","Charge"))`, `=IF((${g}="Jabodetabek")*(${c}>=20)*(${nt}>=3000000),"Free","Charge")`] };
        },
        () => {
          const S = H.table([C.sku, C.warehouse, C.onhand, C.reorder, C.lead], r.shuffle(H.stock(r, 12))), w = S.cell('Warehouse', 0), o = S.cell('On hand', 0), p = S.cell('Reorder point', 0), l = S.cell('Lead time (days)', 0);
          const f = (wh, m, d, lead) => ({ Warehouse: wh, 'On hand': (x) => m * x['Reorder point'] + d, 'Lead time (days)': lead });
          return { S, title: 'Move excess stock: over twice reorder, not Medan, short lead', file: 'stok_gudang.xlsx', ans: 'Plan', T: 'Transfer', F: 'Keep',
            ask: 'A stock line is moved to another warehouse when all three hold: <b>On hand</b> is more than twice the <b>Reorder point</b>, the <b>Warehouse</b> is not Gudang Medan, and the <b>Lead time (days)</b> is at most 5. Return <b>Transfer</b> or <b>Keep</b>.',
            p: [`${o}>2*${p}`, `${w}<>"Gudang Medan"`, `${l}<=5`],
            rule: (q) => q['On hand'] > 2 * q['Reorder point'] && q.Warehouse !== 'Gudang Medan' && q['Lead time (days)'] <= 5,
            pin: [[1, f('Gudang Bekasi', 2, 0, 3)], [3, f('Gudang Bandung', 2, 60, 5)], [5, f('Gudang Medan', 3, 0, 3)], [7, f('Gudang Surabaya', 3, 0, 10)], [9, f('Gudang Semarang', 2, 100, 3)]],
            swap: [[1, { 'On hand': (x) => 2 * x['Reorder point'] + 20 }], [9, { 'Lead time (days)': 7 }]],
            more: [`=IF(AND(${o}>2*${p},NOT(${w}="Gudang Medan"),${l}<=5),"Transfer","Keep")`, `=IF(AND(${o}/${p}>2,${w}<>"Gudang Medan",${l}<=5),"Transfer","Keep")`] };
        }
      ][v]();
      const S = V.S, T = `"${V.T}"`, F = `"${V.F}"`;
      const base = plant(S, Object.assign({}, S.cells, V.top || {}), V.pin), alt = plant(S, Object.assign({}, base, V.top ? { B1: V.b1alt } : {}), V.swap);
      const [p1, p2, p3] = V.p, solution = `=IF(AND(${p1},${p2},${p3}),${T},${F})`;
      return column({
        title: V.title, file: V.file, S, base, alt, ans: V.ans, want: [V.T, V.F], rule: (x, cells) => (V.rule(x, cells) ? V.T : V.F),
        dateCells: V.top ? ['B1'] : undefined, ask: `<p>${V.ask}</p>`,
        solution, alts: [`=IF(AND(${p3},${p2},${p1}),${T},${F})`, `=IF(${p1},IF(${p2},IF(${p3},${T},${F}),${F}),${F})`].concat(V.more),
        traps: V.top ? [unpinB1(solution)] : [],
        hints: ['AND takes as many tests as you need, separated by commas. It is TRUE only when every one of them holds.',
          `Write each rule as its own test: ${p1}, ${p2} and ${p3}.`,
          'Shape: ' + solution],
        explain: `AND(${p1}, ${p2}, ${p3}) is TRUE only on rows that pass all three, so one failed rule is enough for ${F}. A nested IF gives the same answer with three IFs.`,
        chips: [p1, p2, p3, T, F], fnKey: 'AND('
      });
    }
  });

  /* ======================= Medium, level 6: tiered rates, AND with OR ======================= */

  P.add({
    key: 'tiered-rate-times-amount', difficulty: 'Medium', level: 6, topics: ['Nested IF', 'IFS'], count: 4,
    make(r, v) {
      const V = [
        () => {
          const S = H.table([C.rep, C.area, C.target, C.actual], H.reps(r, 10)), t = S.cell('Target (Rp)', 0), a = S.cell('Actual (Rp)', 0), k = `${a}/${t}`;
          const f = (m) => ({ 'Actual (Rp)': (x) => r5(x['Target (Rp)'] * m) });
          return { S, title: 'Commission at 0, 1, 3 or 5% by target achievement', file: 'komisi_sales.xlsx', ans: 'Commission (Rp)',
            guard: (cells) => clearOf(S, cells, 'Actual (Rp)', 'Target (Rp)', [1.2, 0.8], 300000),
            ask: 'Commission is a rate times <b>Actual (Rp)</b>. The rate depends on achievement, Actual divided by Target: <b>5%</b> at 1.2 or more, <b>3%</b> from 1 up to 1.2, <b>1%</b> from 0.8 up to 1, and <b>0</b> below 0.8.',
            rule: (x) => { const q = x['Actual (Rp)'] / x['Target (Rp)']; return x['Actual (Rp)'] * (q >= 1.2 ? 0.05 : q >= 1 ? 0.03 : q >= 0.8 ? 0.01 : 0); },
            pin: [[1, f(1)], [3, f(1.25)], [5, f(0.9)], [7, f(0.7)]], swap: [[5, f(1.02)], [7, f(0.82)]],
            solution: `=${a}*IF(${k}>=1.2,0.05,IF(${k}>=1,0.03,IF(${k}>=0.8,0.01,0)))`,
            alts: [`=${a}*IFS(${k}>=1.2,5%,${k}>=1,3%,${k}>=0.8,1%,TRUE,0)`, `=IF(${a}>=1.2*${t},${a}*0.05,IF(${a}>=${t},${a}*0.03,IF(${a}>=0.8*${t},${a}*0.01,0)))`],
            h2: `Work out the rate with a nested IF on ${k}, then multiply the whole IF by ${a}.`, chips: [a, k, '>=1.2,0.05', '>=1,0.03', '>=0.8,0.01'] };
        },
        () => {
          const rows = H.sales(r, 14), S = H.table([C.date, C.dist, C.cartons, C.gross], rows), c = S.cell('Cartons', 0), g = S.cell('Gross (Rp)', 0);
          const q = (i, n) => [i, { Cartons: n, 'Gross (Rp)': n * rows[i].price }];
          return { S, title: 'Volume discount in rupiah from carton tiers', file: 'diskon_volume.xlsx', ans: 'Discount (Rp)',
            ask: 'The volume discount is a rate times <b>Gross (Rp)</b>. The rate depends on <b>Cartons</b>: <b>7%</b> at 50 or more, <b>5%</b> at 30 to 49, <b>2%</b> at 15 to 29, and <b>0</b> below 15.',
            rule: (x) => x['Gross (Rp)'] * (x.Cartons >= 50 ? 0.07 : x.Cartons >= 30 ? 0.05 : x.Cartons >= 15 ? 0.02 : 0),
            pin: [q(1, 50), q(3, 30), q(5, 15), q(7, 9)], swap: [q(7, 16), q(3, 29)],
            solution: `=${g}*IF(${c}>=50,0.07,IF(${c}>=30,0.05,IF(${c}>=15,0.02,0)))`,
            alts: [`=${g}*IFS(${c}>=50,7%,${c}>=30,5%,${c}>=15,2%,TRUE,0)`, `=IF(${c}>=50,${g}*0.07,IF(${c}>=30,${g}*0.05,IF(${c}>=15,${g}*0.02,0)))`],
            h2: `The tests read ${c}; the rate multiplies ${g}. Start from the top tier, ${c}>=50.`, chips: [g, c, '>=50,0.07', '>=30,0.05', '>=15,0.02'] };
        },
        () => {
          const S = H.table([C.facility, C.dist, C.outstanding, C.dpd], H.loans(r, 14)), o = S.cell('Outstanding (Rp)', 0), d = S.cell('Days past due', 0), dd = (x) => ({ 'Days past due': x });
          return { S, title: 'Loan loss provision by days-past-due tier', file: 'cadangan_kerugian.xlsx', ans: 'Provision (Rp)',
            ask: 'Bank Kenari sets aside part of each <b>Outstanding (Rp)</b> against losses, by <b>Days past due</b>: <b>1%</b> at 0 days, <b>5%</b> from 1 to 90, <b>15%</b> from 91 to 120, <b>50%</b> from 121 to 180, and <b>100%</b> above 180.',
            rule: (x) => x['Outstanding (Rp)'] * [0.01, 0.05, 0.15, 0.5, 1][kol(x['Days past due']) - 1],
            pin: [[0, dd(0)], [2, dd(90)], [4, dd(120)], [6, dd(180)], [8, dd(200)], [10, dd(45)]], swap: [[2, dd(95)], [10, dd(0)]],
            solution: `=${o}*IFS(${d}=0,0.01,${d}<=90,0.05,${d}<=120,0.15,${d}<=180,0.5,TRUE,1)`,
            alts: [`=${o}*IF(${d}=0,0.01,IF(${d}<=90,0.05,IF(${d}<=120,0.15,IF(${d}<=180,0.5,1))))`, `=IFS(${d}>180,${o},${d}>120,${o}*0.5,${d}>90,${o}*0.15,${d}>0,${o}*0.05,TRUE,${o}*0.01)`],
            h2: `Five tiers: IFS keeps them flat. Go up from ${d}=0, use <= for "from ... to", and end with TRUE,1 for above 180.`, chips: [o, d, '=0,0.01', '<=90,0.05', '<=120,0.15', '<=180,0.5', 'TRUE,1)'] };
        },
        () => {
          const S = H.table([C.date, C.dist, C.region, C.cartons], H.sales(r, 14)), c = S.cell('Cartons', 0), cc = (x) => ({ Cartons: x });
          return { S, title: 'Freight with a flat fee for small drops and two carton rates', file: 'pengiriman_q4.xlsx', ans: 'Freight (Rp)',
            ask: 'The carrier charges a flat <b>Rp 75,000</b> for up to 10 <b>Cartons</b>, <b>Rp 5,000 a carton</b> for 11 to 40 cartons, and <b>Rp 4,000 a carton</b> above 40 cartons. Every carton on the line is charged at the one rate.',
            rule: (x) => (x.Cartons <= 10 ? 75000 : x.Cartons <= 40 ? x.Cartons * 5000 : x.Cartons * 4000),
            pin: [[1, cc(10)], [3, cc(40)], [5, cc(41)], [7, cc(6)]], swap: [[7, cc(12)], [3, cc(44)]],
            solution: `=IF(${c}<=10,75000,IF(${c}<=40,${c}*5000,${c}*4000))`,
            alts: [`=IFS(${c}<=10,75000,${c}<=40,${c}*5000,TRUE,${c}*4000)`, `=IF(${c}>40,${c}*4000,IF(${c}>10,${c}*5000,75000))`],
            h2: `The first tier is a fixed amount, the other two multiply ${c}. Test ${c}<=10 first, then ${c}<=40.`, chips: [c, '<=10,75000', `<=40,${c}*5000`, `${c}*4000)`] };
        }
      ][v]();
      const S = V.S, g = V.guard || ((c) => c);
      const base = g(plant(S, S.cells, V.pin)), alt = g(plant(S, base, V.swap));
      return column({
        title: V.title, file: V.file, S, base, alt, ans: V.ans, rule: V.rule, ask: `<p>${V.ask}</p><p>The results come out in whole rupiah.</p>`,
        solution: V.solution, alts: V.alts,
        hints: ['Find each row\'s tier with a nested IF (or IFS), taking the tiers in order from one end. Then apply that tier\'s rate or amount to the row.', V.h2, 'Shape: ' + V.solution],
        explain: 'Each row falls into one tier: the first test that holds picks its rate or amount, and it is applied to that row\'s own figure. Taking the tiers in order from one end means each test needs one limit only.',
        chips: V.chips, fnKey: V.solution.includes('IFS(') ? 'IFS(' : 'IF('
      });
    }
  });

  P.add({
    key: 'and-or-together', difficulty: 'Medium', level: 6, topics: ['AND', 'OR', 'IF'], count: 4,
    make(r, v) {
      const V = [
        () => {
          const S = H.table([C.date, C.dist, C.category, C.cartons], H.sales(r, 14)), k = S.cell('Category', 0), c = S.cell('Cartons', 0);
          const or = `OR(${k}="Coffee",${k}="Sambal")`, kc = (cat, n) => ({ Category: cat, Cartons: n });
          return { S, title: 'Promo for Coffee or Sambal lines of at least 25 cartons', file: 'promo_q4.xlsx', ans: 'Promo', T: 'Promo', F: 'No promo',
            ask: 'A line earns the promo when its <b>Category</b> is Coffee or Sambal, <b>and</b> it has at least 25 <b>Cartons</b>. Return <b>Promo</b> or <b>No promo</b>.',
            rule: (x) => (x.Category === 'Coffee' || x.Category === 'Sambal') && x.Cartons >= 25,
            pin: [[1, kc('Coffee', 25)], [3, kc('Sambal', 40)], [5, kc('Coffee', 18)], [7, kc('Noodles', 45)], [9, kc('Soap', 30)]], swap: [[5, { Cartons: 26 }], [3, { Cartons: 22 }]],
            test: `AND(${or},${c}>=25)`,
            alts: [`=IF(AND(${c}>=25,${or}),"Promo","No promo")`, `=IF(${c}<25,"No promo",IF(${or},"Promo","No promo"))`, `=IF(((${k}="Coffee")+(${k}="Sambal"))*(${c}>=25),"Promo","No promo")`],
            h2: `The category part is ${or}. It goes inside AND, next to ${c}>=25.` };
        },
        () => {
          const S = H.table([C.facility, C.dist, C.dpd, C.restr], H.loans(r, 14)), d = S.cell('Days past due', 0), x = S.cell('Restructured', 0);
          const and = `AND(${x}="Y",${d}>30)`, dr = (n, y) => ({ 'Days past due': n, Restructured: y });
          return { S, title: 'Escalate past 90 days, or past 30 days when restructured', file: 'buku_pinjaman.xlsx', ans: 'Action', T: 'Escalate', F: 'Monitor',
            ask: 'Citra escalates a facility when it is more than 90 <b>Days past due</b>, <b>or</b> when it is restructured (<b>Restructured</b> is Y) <b>and</b> more than 30 days past due. Return <b>Escalate</b> or <b>Monitor</b>.',
            rule: (o) => o['Days past due'] > 90 || (o.Restructured === 'Y' && o['Days past due'] > 30),
            pin: [[1, dr(120, 'N')], [3, dr(45, 'Y')], [5, dr(45, 'N')], [7, dr(15, 'Y')], [9, dr(90, 'N')], [11, dr(30, 'Y')]], swap: [[7, { 'Days past due': 60 }], [5, { 'Days past due': 95 }]],
            test: `OR(${d}>90,${and})`,
            alts: [`=IF(${d}>90,"Escalate",IF(${and},"Escalate","Monitor"))`, `=IF(OR(${and},${d}>90),"Escalate","Monitor")`],
            h2: `The second way in is ${and}. It goes inside OR, next to ${d}>90.` };
        },
        () => {
          const S = H.table([C.rep, C.area, C.target, C.actual, C.visits], H.reps(r, 10)), g = S.cell('Area', 0), t = S.cell('Target (Rp)', 0), a = S.cell('Actual (Rp)', 0), n = S.cell('Visits', 0);
          const or = `OR(${n}>=45,${g}="Sulawesi")`, rv = (area, m, vis) => ({ Area: area, 'Actual (Rp)': (x) => x['Target (Rp)'] + m * 1000000, Visits: vis });
          return { S, title: 'Bonus on target with 45 visits, no visit rule in Sulawesi', file: 'tim_sales.xlsx', ans: 'Bonus', T: 'Bonus', F: 'No bonus',
            ask: 'A rep earns the bonus when <b>Actual (Rp)</b> is at least <b>Target (Rp)</b>, <b>and</b> either the rep made at least 45 <b>Visits</b> <b>or</b> the <b>Area</b> is Sulawesi, where outlets are far apart. Return <b>Bonus</b> or <b>No bonus</b>.',
            rule: (o) => o['Actual (Rp)'] >= o['Target (Rp)'] && (o.Visits >= 45 || o.Area === 'Sulawesi'),
            pin: [[1, rv('Jawa Barat', 3, 45)], [3, rv('Sulawesi', 1, 28)], [5, rv('Jawa Timur', 2, 38)], [7, rv('Jabodetabek', -6, 55)], [9, rv('Sulawesi', -3, 30)]],
            swap: [[5, { Visits: 46 }], [3, { 'Actual (Rp)': (x) => x['Target (Rp)'] - 1000000 }]],
            test: `AND(${a}>=${t},${or})`,
            alts: [`=IF(AND(${or},${a}>=${t}),"Bonus","No bonus")`, `=IF(${a}<${t},"No bonus",IF(${or},"Bonus","No bonus"))`],
            h2: `The visits part is ${or}. It goes inside AND, next to ${a}>=${t}.` };
        },
        () => {
          const S = H.table([C.sku, C.warehouse, C.onhand, C.reorder, C.lead], r.shuffle(H.stock(r, 12))), o = S.cell('On hand', 0), p = S.cell('Reorder point', 0), l = S.cell('Lead time (days)', 0);
          const and = `AND(${o}<${p},${l}>=10)`, sl = (d, lead) => ({ 'On hand': (x) => Math.max(0, x['Reorder point'] + d), 'Lead time (days)': lead });
          return { S, title: 'Urgent stock: none left, or below reorder with a long lead', file: 'stok_gudang.xlsx', ans: 'Priority', T: 'Urgent', F: 'Normal',
            ask: 'A SKU is <b>Urgent</b> when its <b>On hand</b> is 0, <b>or</b> when it is below its <b>Reorder point</b> <b>and</b> its <b>Lead time (days)</b> is at least 10. Otherwise it is <b>Normal</b>.',
            rule: (x) => x['On hand'] === 0 || (x['On hand'] < x['Reorder point'] && x['Lead time (days)'] >= 10),
            pin: [[1, { 'On hand': 0, 'Lead time (days)': 3 }], [3, sl(-30, 14)], [5, sl(-40, 5)], [7, sl(60, 14)], [9, sl(-10, 10)]], swap: [[5, { 'Lead time (days)': 10 }], [1, { 'On hand': (x) => x['Reorder point'] + 40 }]],
            test: `OR(${o}=0,${and})`,
            alts: [`=IF(${o}=0,"Urgent",IF(${and},"Urgent","Normal"))`, `=IF(OR(${and},${o}=0),"Urgent","Normal")`],
            h2: `The second way in is ${and}. It goes inside OR, next to ${o}=0.` };
        }
      ][v]();
      const S = V.S, base = plant(S, S.cells, V.pin), alt = plant(S, base, V.swap), T = `"${V.T}"`, F = `"${V.F}"`;
      const solution = `=IF(${V.test},${T},${F})`;
      return column({
        title: V.title, file: V.file, S, base, alt, ans: V.ans, want: [V.T, V.F], rule: (x) => (V.rule(x) ? V.T : V.F), ask: `<p>${V.ask}</p>`,
        solution, alts: V.alts,
        hints: ['AND and OR can sit inside each other, the way brackets group a sum. Find the part of the rule that is grouped and write it first.', V.h2, 'Shape: ' + solution],
        explain: `${V.test} follows the grouping of the rule: the inner function settles the grouped part, and the outer one joins it with the other test.`,
        chips: [V.test.slice(0, V.test.indexOf('(') + 1)].concat(V.test.match(/(AND|OR)\([^()]*\)/g) || []).concat([T, F]), fnKey: V.test.slice(0, V.test.indexOf('(') + 1)
      });
    }
  });

  /* ======================= Medium, level 7: the first rule that applies ======================= */

  P.add({
    key: 'first-rule-that-applies', difficulty: 'Medium', level: 7, topics: ['Nested IF', 'IFS'], count: 4,
    make(r, v) {
      const asof = H.toSerial('2026-12-15');
      const V = [
        () => {
          const S = H.table([C.facility, C.limit, C.outstanding, C.dpd, C.restr], H.loans(r, 14)), l = S.cell('Limit (Rp)', 0), o = S.cell('Outstanding (Rp)', 0), d = S.cell('Days past due', 0), x = S.cell('Restructured', 0);
          const f = (lim, out, dpd, y) => ({ 'Limit (Rp)': lim * 1e6, 'Outstanding (Rp)': out * 1e6, 'Days past due': dpd, Restructured: y });
          return { S, title: 'Facility review label from the first rule that applies', file: 'buku_pinjaman.xlsx', ans: 'Review',
            guard: (cells) => clearOf(S, cells, 'Outstanding (Rp)', 'Limit (Rp)', [0.9], -1000000),
            ask: 'Dimas labels each facility with the first rule that applies, in this order: <b>Past due</b> when <b>Days past due</b> is more than 0; <b>Restructured</b> when <b>Restructured</b> is Y; <b>Near limit</b> when <b>Outstanding (Rp)</b> is more than 90% of <b>Limit (Rp)</b>; otherwise <b>Clear</b>.',
            steps: [[`${d}>0`, 'Past due'], [`${x}="Y"`, 'Restructured'], [`${o}>0.9*${l}`, 'Near limit']], last: 'Clear',
            rule: (q) => (q['Days past due'] > 0 ? 'Past due' : q.Restructured === 'Y' ? 'Restructured' : q['Outstanding (Rp)'] > 0.9 * q['Limit (Rp)'] ? 'Near limit' : 'Clear'),
            pin: [[1, f(100, 60, 30, 'Y')], [3, f(150, 90, 0, 'Y')], [5, f(100, 95, 0, 'N')], [7, f(150, 90, 0, 'N')], [9, f(200, 184, 0, 'N')]], swap: [[7, { 'Outstanding (Rp)': 140000000 }], [3, { 'Days past due': 15 }]],
            more: [`=IF(${d}>0,"Past due",IF(${x}="Y","Restructured",IF(${o}/${l}>0.9,"Near limit","Clear")))`],
            bad: [1, 0] };
        },
        () => {
          const S = H.table([C.sku, C.warehouse, C.onhand, C.reorder], r.shuffle(H.stock(r, 12))), o = S.cell('On hand', 0), p = S.cell('Reorder point', 0);
          const f = (m, d) => ({ 'On hand': (x) => m * x['Reorder point'] + d });
          return { S, title: 'Stock action: out, reorder, overstock or OK, in that order', file: 'stok_gudang.xlsx', ans: 'Action',
            ask: 'Give each stock line the first label that applies, in this order: <b>Out</b> when <b>On hand</b> is 0; <b>Reorder</b> when On hand is below the <b>Reorder point</b>; <b>Overstock</b> when On hand is more than 4 times the Reorder point; otherwise <b>OK</b>.',
            steps: [[`${o}=0`, 'Out'], [`${o}<${p}`, 'Reorder'], [`${o}>4*${p}`, 'Overstock']], last: 'OK',
            rule: (q) => (q['On hand'] === 0 ? 'Out' : q['On hand'] < q['Reorder point'] ? 'Reorder' : q['On hand'] > 4 * q['Reorder point'] ? 'Overstock' : 'OK'),
            pin: [[1, { 'On hand': 0 }], [3, f(1, -20)], [5, f(4, 0)], [7, f(4, 50)], [9, f(1, 0)]], swap: [[5, f(4, 10)], [1, { 'On hand': 60 }]],
            more: [`=IF(${o}=0,"Out",IF(${o}<${p},"Reorder",IF(${o}<=4*${p},"OK","Overstock")))`],
            bad: [1, 0] };
        },
        () => {
          const rows = H.sales(r, 14).map((x) => Object.assign(x, { credit: x.net + r.int(1, 20) * 1000000 }));
          const S = H.table([C.date, C.dist, C.region, C.cartons, C.net, { h: 'Credit left (Rp)', k: 'credit', num: true }], rows);
          const g = S.cell('Region', 0), c = S.cell('Cartons', 0), nt = S.cell('Net (Rp)', 0), cr = S.cell('Credit left (Rp)', 0);
          const f = (reg, n, net, credit) => Object.assign(reg ? inRegion(reg, 0) : {}, { Cartons: n, 'Net (Rp)': net, 'Credit left (Rp)': credit });
          return { S, title: 'Order release check with three holds in a set order', file: 'rilis_order.xlsx', ans: 'Release',
            ask: 'Each order line gets the first label that applies, in this order: <b>Check qty</b> when <b>Cartons</b> is below 5; <b>Credit hold</b> when <b>Net (Rp)</b> is more than <b>Credit left (Rp)</b>; <b>Sea freight</b> when the <b>Region</b> is Sumatera or Sulawesi; otherwise <b>Release</b>.',
            steps: [[`${c}<5`, 'Check qty'], [`${nt}>${cr}`, 'Credit hold'], [`OR(${g}="Sumatera",${g}="Sulawesi")`, 'Sea freight']], last: 'Release',
            rule: (q) => (q.Cartons < 5 ? 'Check qty' : q['Net (Rp)'] > q['Credit left (Rp)'] ? 'Credit hold' : q.Region === 'Sumatera' || q.Region === 'Sulawesi' ? 'Sea freight' : 'Release'),
            pin: [[1, f('', 3, 450000, 300000)], [3, f('', 18, 3200000, 2000000)], [5, f('Sumatera', 22, 3500000, 12000000)], [7, f('Jawa Tengah', 30, 5100000, 5100000)], [9, f('Jabodetabek', 25, 4200000, 20000000)]],
            swap: [[7, { 'Credit left (Rp)': 4000000 }], [9, inRegion('Sulawesi', 0)]],
            more: [], bad: [1, 0] };
        },
        () => {
          const rows = H.invoices(r, 14).map((x) => { const d = asof - r.int(8, 80); return Object.assign(x, { serial: d, paidOn: r.chance(0.35) ? Math.min(asof - 1, d + r.int(10, 45)) : undefined }); });
          const S = H.table([C.invoice, { h: 'Date', k: 'serial', date: true }, C.terms, C.amount, C.paid], rows);
          const b = S.cell('Date', 0), t = S.cell('Terms (days)', 0), m = S.cell('Amount (Rp)', 0), p = S.cell('Paid on', 0);
          const f = (back, terms, amt, paid) => ({ Date: asof - back, 'Terms (days)': terms, 'Amount (Rp)': amt, 'Paid on': paid === null ? null : asof - paid });
          return { S, title: 'Invoice follow-up as of the date in B1, first rule first', file: 'piutang_q4.xlsx', ans: 'Follow-up',
            top: { A1: 'As of', B1: asof }, b1alt: asof + 10, paidCol: true,
            ask: 'Bu Lestari labels each invoice with the first rule that applies, in this order: <b>Paid</b> when <b>Paid on</b> has a date; <b>Not due</b> when the due date (<b>Date</b> plus <b>Terms (days)</b>) is on or after the date in <b>B1</b>; <b>Call</b> when the <b>Amount (Rp)</b> is at least Rp 5,000,000; otherwise <b>Email</b>. Point at B1 with $ signs, because the date moves on.',
            steps: [[`${p}<>""`, 'Paid'], [`${b}+${t}>=$B$1`, 'Not due'], [`${m}>=5000000`, 'Call']], last: 'Email',
            rule: (q, cells) => (q['Paid on'] !== undefined ? 'Paid' : q.Date + q['Terms (days)'] >= cells.B1 ? 'Not due' : q['Amount (Rp)'] >= 5000000 ? 'Call' : 'Email'),
            pin: [[1, f(50, 30, 6800000, 15)], [3, f(25, 30, 6100000, null)], [5, f(50, 30, 7200000, null)], [7, f(40, 30, 2300000, null)], [9, f(45, 45, 3900000, null)], [11, f(38, 30, 2800000, null)]],
            swap: [[7, { 'Paid on': asof + 4 }]],
            more: [`=IF(ISBLANK(${p}),IF(${b}+${t}>=$B$1,"Not due",IF(${m}>=5000000,"Call","Email")),"Paid")`, `=IF(${p}<>"","Paid",IF(${b}+${t}>=B$1,"Not due",IF(${m}>=5000000,"Call","Email")))`],
            bad: null };
        }
      ][v]();
      const S = V.S, g = V.guard || ((c) => c);
      const base = g(plant(S, Object.assign({}, S.cells, V.top || {}), V.pin)), alt = g(plant(S, Object.assign({}, base, V.top ? { B1: V.b1alt } : {}), V.swap));
      const nest = (steps) => steps.reduceRight((acc, [tst, out]) => `IF(${tst},"${out}",${acc})`, `"${V.last}"`);
      const solution = '=' + nest(V.steps);
      const flat = `=IFS(${V.steps.map(([tst, out]) => `${tst},"${out}"`).join(',')},TRUE,"${V.last}")`;
      const traps = [];
      if (V.bad) { const s2 = V.steps.slice(); [s2[V.bad[0]], s2[V.bad[1]]] = [s2[V.bad[1]], s2[V.bad[0]]]; traps.push({ code: 'order', title: 'The rules must run in the order given', plain: 'Some rows meet more than one rule. A nested IF stops at the first test that is true, so keep the tests in the order of the task.', formula: '=' + nest(s2) }); }
      if (V.top) traps.push(unpinB1(solution));
      const outs = V.steps.map((s) => s[1]).concat([V.last]);
      return column({
        title: V.title, file: V.file, S, base, alt, ans: V.ans, want: outs, rule: V.rule,
        dateCells: V.top ? ['B1'].concat(filled(S, base, 'Paid on')) : undefined, ask: `<p>${V.ask}</p>`,
        solution, alts: [flat].concat(V.more), traps,
        hints: ['When rules overlap, their order decides the label. A nested IF, or IFS, stops at the first test that is TRUE.',
          `Write the tests in the order given, each inside the "false" part of the one before: ${V.steps.map((s) => s[0]).join(', then ')}.`,
          'Shape: ' + solution],
        explain: `A row reaches a test only when every test before it failed, so a row that meets two rules gets the earlier label. IFS with the same tests in the same order, ending in TRUE, gives the same answer in one flat list.`,
        chips: V.steps.map(([tst, out]) => `${tst},"${out}"`).concat([`"${V.last}"`]), fnKey: 'IF('
      });
    }
  });
});
