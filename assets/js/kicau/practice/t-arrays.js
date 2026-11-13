/*
 * Practice, Arrays (20 Medium, 30 Hard): FILTER with one or several rules and a fallback text, SORT and SORTBY,
 * UNIQUE, SEQUENCE, LET, SUMPRODUCT with several conditions and two-way grids, XMATCH, and top N with TAKE.
 * A spilled answer starts at its anchor and is judged over the whole block, on the visible and the hidden sheet.
 *
 * What the practice engine does not do, so no problem here relies on it: MONTH, YEAR, WEEKDAY, LEFT, ROUND, INT,
 * MOD, ABS, ISBLANK, ROW and COLUMN are not applied over arrays; SORT and SORTBY take one key only; COUNTIF(S) and
 * SUMIFS do not take an array as criteria; LARGE does not take an array of k; TAKE asks for no more rows than there
 * are and needs its rows argument; FILTER does not filter columns with a row of tests; LOOKUP(2,1/(...)) can miss
 * the last match. Dates are compared with input cells instead.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) factory(require('./core.js'));
  else factory(root.KicauPractice);
})(typeof globalThis !== 'undefined' ? globalThis : this, function (P) {
  'use strict';
  const H = P.H;

  /* ---------- helpers ---------- */
  const rp = (n) => Math.round(n).toLocaleString('en-US');
  const idx = (n) => Array.from({ length: n }, (_, i) => i);
  const isIn = (k, list) => (s) => list.includes(s[k]);
  const ser = (iso) => H.toSerial(iso);
  /* Sales lines from the shared ledger, part by part ([[how many, which lines], ...]), dated in Q4, sorted by date. */
  function salesFrom(r, parts) {
    const used = new Set(), out = [];
    for (const [n, keep] of parts) {
      const pool = H.db().sales.filter((s) => keep(s) && !used.has(s.invoice));
      for (const s of r.sample(pool, Math.min(n, pool.length))) {
        used.add(s.invoice);
        const p = H.PRODUCTS.find((x) => x.sku === s.sku) || {};
        out.push(Object.assign({}, s, { date: H.Q4(s.date), price: p.carton_price }));
      }
    }
    return out.sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : a.invoice < b.invoice ? -1 : 1));
  }
  /* Nudge a number column so no two rows tie: a sort or a top N then never depends on row order. */
  function untie(rows, k, step) { const seen = new Set(); for (const x of rows) { while (seen.has(x[k])) x[k] += step; seen.add(x[k]); } return rows; }
  /* The same for actual / target, by raising the actual one million at a time. */
  function untieRatio(rows) { const seen = new Set(); const key = (x) => Math.round(x.actual / x.target * 1e6); for (const x of rows) { while (seen.has(key(x))) x.actual += 1000000; seen.add(key(x)); } return rows; }
  /* Values of column k, most common first. */
  function tally(rows, k) { const m = new Map(); for (const x of rows) m.set(x[k], (m.get(x[k]) || 0) + 1); return [...m.entries()].sort((a, b) => b[1] - a[1]).map((e) => e[0]); }
  /* A loan book with exactly `tl` term loans, so every type rule has rows to find. */
  function book(r, n, tl) {
    const rows = H.loans(r, n);
    const pick = new Set(r.sample(idx(n), tl));
    rows.forEach((x, i) => { x.type = pick.has(i) ? 'Term Loan' : 'Working Capital'; });
    return rows;
  }
  const blk = (S, h1, h2) => S.col(h1) + S.first + ':' + S.col(h2) + S.last;
  const byDate = (a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : 0);

  /* ======================================================================================== */
  /* Medium                                                                                    */
  /* ======================================================================================== */

  P.add({
    key: 'filter-rows-one-rule', difficulty: 'Medium', level: 5, topics: ['FILTER'], count: 4,
    make(r, v) {
      return [
        () => {
          const names = r.sample(H.db().distributors, 4).map((d) => d.name);
          const rows = salesFrom(r, [[16, isIn('distributor', names)]]);
          const S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'Invoice', k: 'invoice' }, { h: 'Distributor', k: 'distributor' }, { h: 'Product', k: 'product' }, { h: 'Net (Rp)', k: 'net', num: true }], rows);
          const [who, next] = tally(rows, 'distributor');
          const base = Object.assign({}, S.cells, { A1: 'Distributor', B1: who });
          /* The hidden sheet: B1 names another distributor, and one line is booked to it. */
          const i = rows.findIndex((x) => x.distributor === who);
          const alt = H.edit(base, { B1: next, [S.cell('Distributor', i)]: next });
          const inv = S.R('Invoice'), dis = S.R('Distributor');
          return H.problem({
            title: "List one distributor's invoice numbers", file: 'penjualan_q4.xlsx', sheet: S, base, alt,
            spill: { cols: 1, rows: rows.length }, spillAnchor: 'G5', ansLabel: 'Invoices',
            statement: `<p>Pak Anton wants every invoice number from the distributor named in <b>B1</b>. In <b>G5</b>, type one formula that lists them. The list spills down by itself.</p>`,
            solution: `=FILTER(${inv},${dis}=B1)`,
            alts: [`=FILTER(${S.A('Invoice')},${S.A('Distributor')}=$B$1)`, `=FILTER(${inv},B1=${dis})`],
            hints: ['FILTER keeps the rows of a range that pass a test: FILTER(what to return, which rows to keep).', `Return the Invoice column ${inv}, and keep the rows where ${dis} equals B1.`, `=FILTER(${inv}, ${dis}=B1)`],
            explain: `${dis}=B1 is TRUE on that distributor's rows, and FILTER returns the invoice numbers on those rows only. It points at B1, so the list follows when B1 changes.`,
            chips: [inv, dis, '=B1'], fnKey: 'FILTER('
          });
        },
        () => {
          const rows = H.stock(r, 12);
          const [a, b, c, d, e] = r.sample(idx(12), 5);
          rows.forEach((x) => { if (x.onhand <= x.reorder) x.onhand = x.reorder + r.int(2, 30) * 10; });
          [a, b, c].forEach((i) => { rows[i].onhand = rows[i].reorder - r.int(1, 6) * 10; });
          rows[e].onhand = rows[e].reorder;
          const S = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Product', k: 'product' }, { h: 'Warehouse', k: 'warehouse' }, { h: 'On hand', k: 'onhand', num: true }, { h: 'Reorder point', k: 'reorder', num: true }], rows);
          const base = Object.assign({}, S.cells);
          /* The hidden sheet: one low line is restocked and another one runs low. */
          const alt = H.edit(base, { [S.cell('On hand', a)]: rows[a].reorder + 150, [S.cell('On hand', d)]: rows[d].reorder - 20 });
          const pr = S.R('Product'), oh = S.R('On hand'), ro = S.R('Reorder point');
          return H.problem({
            title: 'List the products below their reorder point', file: 'stok_gudang.xlsx', sheet: S, base, alt,
            spill: { cols: 1, rows: rows.length }, spillAnchor: 'G5', ansLabel: 'Reorder now',
            statement: `<p>Bu Lestari wants to know what to reorder: every line whose <b>On hand</b> is less than its <b>Reorder point</b> (equal does not count). In <b>G5</b>, list the <b>Product</b> of those lines with one formula.</p>`,
            solution: `=FILTER(${pr},${oh}<${ro})`,
            alts: [`=FILTER(${S.A('Product')},${S.A('On hand')}<${S.A('Reorder point')})`, `=FILTER(${pr},${ro}>${oh})`],
            hints: ['The test can compare two columns row by row: one range against another range of the same size.', `Keep the rows where ${oh} is less than ${ro}, and return ${pr}.`, `=FILTER(${pr}, ${oh}<${ro})`],
            explain: `${oh}<${ro} compares each line's stock with its own reorder point and gives TRUE where it is lower. FILTER returns the product names on those lines.`,
            chips: [pr, oh, ro, '<'], fnKey: 'FILTER('
          });
        },
        () => {
          const rows = book(r, 14, 3);
          const [a, b, c, d] = r.sample(idx(14), 4);
          rows[a].dpd = 45; rows[b].dpd = 95; rows[c].dpd = 30; rows[d].dpd = 61;
          const S = H.table([{ h: 'Facility', k: 'facility' }, { h: 'Distributor', k: 'distributor' }, { h: 'Type', k: 'type' }, { h: 'Outstanding (Rp)', k: 'outstanding', num: true }, { h: 'DPD', k: 'dpd', num: true }], rows);
          const base = Object.assign({}, S.cells, { A1: 'More than (days)', B1: 30 });
          /* The hidden sheet: the limit moves to 60 days, and one facility slips further. */
          const alt = H.edit(base, { B1: 60, [S.cell('DPD', c)]: 75 });
          const fa = S.R('Facility'), dp = S.R('DPD');
          return H.problem({
            title: 'List facilities past a days-past-due limit', file: 'buku_pinjaman.xlsx', sheet: S, base, alt,
            spill: { cols: 1, rows: rows.length }, spillAnchor: 'G5', ansLabel: 'Over the limit',
            statement: `<p>Citra wants the <b>Facility</b> IDs whose days past due (<b>DPD</b>) are more than the number of days in <b>B1</b>. A facility exactly at the limit is not listed. In <b>G5</b>, list them with one formula.</p>`,
            solution: `=FILTER(${fa},${dp}>B1)`,
            alts: [`=FILTER(${S.A('Facility')},${S.A('DPD')}>$B$1)`, `=FILTER(${fa},B1<${dp})`],
            hints: ['"More than" is the > sign; "at least" would be >=.', `Keep the rows where ${dp} is more than B1, and return ${fa}.`, `=FILTER(${fa}, ${dp}>B1)`],
            explain: `${dp}>B1 marks each facility above the limit in B1, and FILTER returns their IDs. Pointing at B1 lets Citra change the limit without touching the formula.`,
            chips: [fa, dp, '>B1'], fnKey: 'FILTER('
          });
        },
        () => {
          const rows = H.reps(r, 12);
          const [a, b, c] = r.sample(idx(rows.length), 3);
          rows[a].actual = rows[a].target;
          rows[b].actual = Math.round(rows[b].target * 1.1 / 100000) * 100000;
          rows[c].actual = Math.round(rows[c].target * 0.85 / 100000) * 100000;
          const S = H.table([{ h: 'Name', k: 'name' }, { h: 'Area', k: 'area' }, { h: 'Target (Rp)', k: 'target', num: true }, { h: 'Actual (Rp)', k: 'actual', num: true }], rows);
          const base = Object.assign({}, S.cells);
          /* The hidden sheet: one rep slips under target and another reaches it exactly. */
          const alt = H.edit(base, { [S.cell('Actual (Rp)', b)]: rows[b].target - 2000000, [S.cell('Actual (Rp)', c)]: rows[c].target });
          const nm = S.R('Name'), tg = S.R('Target (Rp)'), ac = S.R('Actual (Rp)');
          return H.problem({
            title: 'List the sales reps who reached their target', file: 'tim_sales.xlsx', sheet: S, base, alt,
            spill: { cols: 1, rows: rows.length }, spillAnchor: 'F5', ansLabel: 'On target',
            statement: `<p>Pak Anton wants the names of the reps whose <b>Actual</b> is at least their <b>Target</b> (reaching it exactly counts). In <b>F5</b>, list them with one formula.</p>`,
            solution: `=FILTER(${nm},${ac}>=${tg})`,
            alts: [`=FILTER(${S.A('Name')},${S.A('Actual (Rp)')}>=${S.A('Target (Rp)')})`, `=FILTER(${nm},${tg}<=${ac})`],
            hints: ['"At least" is the >= sign, so a rep exactly on target is kept.', `Compare ${ac} with ${tg} row by row and return ${nm}.`, `=FILTER(${nm}, ${ac}>=${tg})`],
            explain: `${ac}>=${tg} is TRUE for every rep at or above target, and FILTER returns their names. The >= keeps the rep who hit the target exactly.`,
            chips: [nm, ac, tg, '>='], fnKey: 'FILTER('
          });
        }
      ][v]();
    }
  });

  P.add({
    key: 'xmatch-find-position', difficulty: 'Medium', level: 5, topics: ['XMATCH', 'MATCH'], count: 3,
    make(r, v) {
      return [
        () => {
          const names = r.sample(H.db().distributors, 4).map((d) => d.name);
          const rows = salesFrom(r, [[16, isIn('distributor', names)]]);
          const S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'Invoice', k: 'invoice' }, { h: 'Distributor', k: 'distributor' }, { h: 'Net (Rp)', k: 'net', num: true }], rows);
          const [who, next] = tally(rows, 'distributor');
          const base = Object.assign({}, S.cells, { A1: 'Distributor', B1: who, A2: 'Newest order at' });
          const alt = H.edit(base, { B1: next });
          const dis = S.R('Distributor');
          return H.problem({
            title: "Position of a distributor's newest order", file: 'penjualan_q4.xlsx', sheet: S, base, alt, answerCell: 'B2',
            statement: `<p>The orders run from oldest (row 5) to newest (row ${S.last}). In <b>B2</b>, return the position in <b>${dis}</b> of the <b>newest</b> order from the distributor in <b>B1</b>, where 1 means row 5.</p>`,
            solution: `=XMATCH(B1,${dis},0,-1)`,
            alts: [`=XMATCH(B1,${S.A('Distributor')},0,-1)`, `=MATCH(2,1/(${dis}=B1))`],
            hints: ['The newest order is the last match in the list, so search from the bottom up.', 'XMATCH has a fourth argument, search_mode: 1 searches first to last, -1 searches last to first.', `=XMATCH(B1, ${dis}, 0, -1)`],
            explain: `XMATCH with match_mode 0 (exact) and search_mode -1 starts at the bottom of ${dis}, so it finds the newest order of that distributor and returns its position from the top.`,
            chips: [dis, ',0,-1)'], fnKey: 'XMATCH('
          });
        },
        () => {
          const from = [0, r.pick([10, 20, 25]), r.pick([40, 50, 60]), r.pick([100, 120]), r.pick([200, 250]), r.pick([400, 500])];
          const names = ['Standard', 'Silver', 'Gold', 'Platinum', 'Key account', 'Strategic'];
          const disc = [0, 1, 2, 3.5, 5, 7];
          const rows = from.map((f, i) => ({ from: f, tier: names[i], disc: disc[i] }));
          const S = H.table([{ h: 'From (cartons)', k: 'from', num: true }, { h: 'Tier', k: 'tier' }, { h: 'Discount (%)', k: 'disc', num: true }], rows);
          const k = r.int(1, 4), order = from[k] + r.int(1, from[k + 1] - from[k] - 1);
          const j = r.pick([1, 2, 3, 4, 5].filter((x) => x !== k));
          const base = Object.assign({}, S.cells, { A1: 'Order (cartons)', B1: order, A2: 'Tier position' });
          /* The hidden sheet: an order that sits exactly on a tier's start. */
          const alt = H.edit(base, { B1: from[j] });
          const fr = S.R('From (cartons)');
          return H.problem({
            title: 'Discount tier position for an order size', file: 'tier_diskon.xlsx', sheet: S, base, alt, answerCell: 'B2',
            statement: `<p>The discount tiers start at the carton counts in <b>${fr}</b>, smallest first. An order belongs to the last tier whose start is at or below its cartons. In <b>B2</b>, return the position of that tier in ${fr} (1 means row 5) for the order in <b>B1</b>.</p>`,
            solution: `=XMATCH(B1,${fr},-1)`,
            alts: [`=MATCH(B1,${fr},1)`, `=MATCH(B1,${fr})`, `=XMATCH(B1,${S.A('From (cartons)')},-1)`],
            hints: ['An exact match will not work: the order size is usually between two tier starts.', 'XMATCH\'s third argument, match_mode, can be -1: the exact value or the next smaller one.', `=XMATCH(B1, ${fr}, -1)`],
            explain: `With match_mode -1, XMATCH takes the exact start if the order sits on it, and otherwise the next smaller start, which is the tier the order falls in.`,
            chips: [fr, ',-1)'], fnKey: 'XMATCH('
          });
        },
        () => {
          const rows = r.shuffle(H.stock(r, 12));
          const fams = ['MLK', 'KSJ', 'SNY', 'SRW'];
          const fam = r.pick(fams), other = r.pick(fams.filter((f) => f !== fam));
          const S = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Product', k: 'product' }, { h: 'Warehouse', k: 'warehouse' }, { h: 'On hand', k: 'onhand', num: true }], rows);
          const base = Object.assign({}, S.cells, { A1: 'SKU starts with', B1: fam, A2: 'First position' });
          const alt = H.edit(base, { B1: other });
          const sk = S.R('SKU');
          return H.problem({
            title: 'First SKU of a product family, by wildcard', file: 'stok_gudang.xlsx', sheet: S, base, alt, answerCell: 'B2',
            statement: `<p>Every SKU starts with a three-letter family code, such as MLK for Mi Lekas. In <b>B2</b>, return the position in <b>${sk}</b> (1 means row 5) of the first SKU that starts with the code in <b>B1</b>.</p>`,
            solution: `=XMATCH(B1&"*",${sk},2)`,
            alts: [`=MATCH(B1&"*",${sk},0)`, `=XMATCH(B1&"-*",${sk},2)`],
            hints: ['A wildcard * stands for any characters, so "MLK*" matches every SKU that starts with MLK.', 'Join the code in B1 to the wildcard with &, then tell XMATCH to read wildcards: match_mode 2.', `=XMATCH(B1&"*", ${sk}, 2)`],
            explain: `B1&"*" builds a pattern such as KSJ*, and match_mode 2 makes XMATCH treat * as a wildcard. It returns the first SKU from the top that fits the pattern.`,
            chips: [sk, '&"*"', ',2)'], fnKey: 'XMATCH('
          });
        }
      ][v]();
    }
  });

  P.add({
    key: 'filter-two-rules-fallback', difficulty: 'Medium', level: 6, topics: ['FILTER'], count: 4,
    make(r, v) {
      return [
        () => {
          const regs = r.sample(H.REGIONS, 3);
          const rows = salesFrom(r, [[16, isIn('region', regs)]]);
          const reg = tally(rows, 'region')[0];
          const cart = rows.filter((x) => x.region === reg).map((x) => x.cartons).sort((a, b) => b - a);
          const S = H.table([{ h: 'Invoice', k: 'invoice' }, { h: 'Distributor', k: 'distributor' }, { h: 'Region', k: 'region' }, { h: 'Category', k: 'category' }, { h: 'Cartons', k: 'cartons', num: true }, { h: 'Net (Rp)', k: 'net', num: true }], rows);
          const base = Object.assign({}, S.cells, { A1: 'Region', B1: reg, A2: 'At least (cartons)', B2: cart[1] });
          /* The hidden sheet: the minimum is raised above every order of that region, so nothing passes. */
          const alt = H.edit(base, { B2: cart[0] + 5 });
          const out = blk(S, 'Invoice', 'Distributor'), re = S.R('Region'), ca = S.R('Cartons');
          return H.problem({
            title: 'Orders from a region with a carton minimum, or None', file: 'penjualan_q4.xlsx', sheet: S, base, alt,
            spill: { cols: 2, rows: rows.length }, spillAnchor: 'H5', ansLabel: 'Invoice',
            statement: `<p>Pak Anton wants the orders from the region in <b>B1</b> with at least <b>B2</b> cartons. In <b>H5</b>, return their <b>Invoice</b> and <b>Distributor</b> (${out}) with one formula. When no order passes both rules, the formula must show the text <code>None</code>.</p>`,
            solution: `=FILTER(${out},(${re}=B1)*(${ca}>=B2),"None")`,
            alts: [`=IFERROR(FILTER(${out},(${re}=B1)*(${ca}>=B2)),"None")`, `=FILTER(${out},(${ca}>=B2)*(${re}=B1),"None")`],
            hints: ['Two rules that must both hold: multiply them. TRUE*TRUE is 1, and anything with a FALSE is 0.', 'FILTER has a third argument, if_empty, shown when no row passes.', `=FILTER(${out}, (${re}=B1)*(${ca}>=B2), "None")`],
            explain: `(${re}=B1)*(${ca}>=B2) is 1 only where both rules hold, so FILTER keeps those rows of ${out}. The third argument shows None instead of the #CALC! error when nothing passes, as on the hidden sheet.`,
            chips: [out, re, ca, ')*(', ',"None")'], fnKey: 'FILTER('
          });
        },
        () => {
          const rows = book(r, 14, 4);
          const wc = idx(14).filter((i) => rows[i].type === 'Working Capital'), tl = idx(14).filter((i) => rows[i].type === 'Term Loan');
          const [w1, w2, w3] = r.sample(wc, 3);
          rows[w1].dpd = 15; rows[w2].dpd = 60; rows[w3].dpd = 0;
          tl.forEach((i) => { rows[i].dpd = 0; });
          const S = H.table([{ h: 'Facility', k: 'facility' }, { h: 'Distributor', k: 'distributor' }, { h: 'Type', k: 'type' }, { h: 'Outstanding (Rp)', k: 'outstanding', num: true }, { h: 'DPD', k: 'dpd', num: true }], rows);
          const base = Object.assign({}, S.cells, { A1: 'Type', B1: 'Working Capital' });
          /* The hidden sheet: the term loans, one of which has slipped. */
          const alt = H.edit(base, { B1: 'Term Loan', [S.cell('DPD', tl[1])]: 30 });
          const out = blk(S, 'Facility', 'Distributor'), ty = S.R('Type'), dp = S.R('DPD');
          return H.problem({
            title: 'Facilities of one type that are past due, or None', file: 'buku_pinjaman.xlsx', sheet: S, base, alt,
            spill: { cols: 2, rows: rows.length }, spillAnchor: 'G5', ansLabel: 'Facility',
            statement: `<p>Dimas wants the facilities of the type in <b>B1</b> that are past due (<b>DPD</b> more than 0). In <b>G5</b>, return their <b>Facility</b> and <b>Distributor</b> (${out}) with one formula, or the text <code>None</code> when there are none.</p>`,
            solution: `=FILTER(${out},(${ty}=B1)*(${dp}>0),"None")`,
            alts: [`=IFERROR(FILTER(${out},(${ty}=B1)*(${dp}>0)),"None")`, `=FILTER(${S.T('Facility', 'Distributor')},(${S.A('Type')}=$B$1)*(${S.A('DPD')}>0),"None")`],
            hints: ['Write each rule in brackets and multiply them, so a row is kept only when both are true.', `The rules are ${ty}=B1 and ${dp}>0. Add "None" as the third argument.`, `=FILTER(${out}, (${ty}=B1)*(${dp}>0), "None")`],
            explain: `Multiplying the two tests gives 1 only for facilities of that type with a DPD above 0, and FILTER returns their ID and distributor. "None" covers a type with no late facility.`,
            chips: [out, ty, dp, ')*(', ',"None")'], fnKey: 'FILTER('
          });
        },
        () => {
          const rows = H.invoices(r, 14).sort(byDate);
          r.sample(idx(14), 6).forEach((i) => { rows[i].paid = ''; });
          const unpaid = rows.filter((x) => x.paid === '').map((x) => x.amount).sort((a, b) => a - b);
          const lim = Math.max(1000000, Math.floor(unpaid[Math.floor(unpaid.length / 2)] / 1000000) * 1000000);
          const S = H.table([{ h: 'Ref', k: 'ref' }, { h: 'Distributor', k: 'distributor' }, { h: 'Amount (Rp)', k: 'amount', num: true }, { h: 'Date', k: 'date', date: true }, { h: 'Paid on', k: 'paid', date: true }], rows);
          const big = rows.reduce((m, x, i) => (x.paid === '' && (m < 0 || x.amount > rows[m].amount) ? i : m), -1);
          const base = Object.assign({}, S.cells, { A1: 'More than (Rp)', B1: lim });
          /* The hidden sheet: a lower limit, and the biggest unpaid invoice has been paid. */
          const alt = H.edit(base, { B1: Math.max(500000, lim - 2000000), [S.cell('Paid on', big)]: ser(rows[big].date) + 12 });
          const out = blk(S, 'Ref', 'Amount (Rp)'), am = S.R('Amount (Rp)'), pd = S.R('Paid on');
          return H.problem({
            title: 'Unpaid invoices above an amount, or None', file: 'piutang_q4.xlsx', sheet: S, base, alt,
            spill: { cols: 3, rows: rows.length }, spillAnchor: 'G5', ansLabel: 'Unpaid',
            statement: `<p>Bu Lestari wants the unpaid invoices (<b>Paid on</b> is blank) with an <b>Amount</b> of more than <b>B1</b>. In <b>G5</b>, return their <b>Ref</b>, <b>Distributor</b> and <b>Amount</b> (${out}) with one formula, or the text <code>None</code> when there are none.</p>`,
            solution: `=FILTER(${out},(${pd}="")*(${am}>B1),"None")`,
            alts: [`=FILTER(${out},(${am}>B1)*(${pd}=""),"None")`, `=IFERROR(FILTER(${out},(${pd}="")*(${am}>B1)),"None")`],
            hints: ['A blank cell is equal to the empty text "", so a test can look for unpaid rows.', `Multiply (${pd}="") by (${am}>B1) and return the block ${out}.`, `=FILTER(${out}, (${pd}="")*(${am}>B1), "None")`],
            explain: `${pd}="" is TRUE on unpaid rows and ${am}>B1 on the large ones; their product keeps rows that pass both. FILTER returns three columns at once because the first argument is three columns wide.`,
            chips: [out, pd, am, '=""', ',"None")'], fnKey: 'FILTER('
          });
        },
        () => {
          const regs = r.sample(H.REGIONS, 4);
          const rows = salesFrom(r, [[16, isIn('region', regs)]]);
          const t = tally(rows, 'region');
          const third = t[2] || regs.find((x) => x !== t[0] && x !== t[1]);
          const S = H.table([{ h: 'Invoice', k: 'invoice' }, { h: 'Distributor', k: 'distributor' }, { h: 'Region', k: 'region' }, { h: 'Net (Rp)', k: 'net', num: true }], rows);
          const base = Object.assign({}, S.cells, { A1: 'Region', B1: t[0], A2: 'Or region', B2: t[1] });
          const alt = H.edit(base, { B2: third });
          const out = blk(S, 'Invoice', 'Distributor'), re = S.R('Region');
          return H.problem({
            title: 'Orders from either of two regions, or None', file: 'penjualan_q4.xlsx', sheet: S, base, alt,
            spill: { cols: 2, rows: rows.length }, spillAnchor: 'F5', ansLabel: 'Invoice',
            statement: `<p>Tika wants the orders from the region in <b>B1</b> or the region in <b>B2</b>. In <b>F5</b>, return their <b>Invoice</b> and <b>Distributor</b> (${out}) with one formula, or the text <code>None</code> when there are none.</p>`,
            solution: `=FILTER(${out},(${re}=B1)+(${re}=B2),"None")`,
            alts: [`=FILTER(${out},((${re}=B1)+(${re}=B2))>0,"None")`, `=FILTER(${out},(${re}=B2)+(${re}=B1),"None")`],
            hints: ['"Either one" is OR. In an array test, OR is written by adding the tests: a row passes when the sum is more than 0.', `Add (${re}=B1) and (${re}=B2).`, `=FILTER(${out}, (${re}=B1)+(${re}=B2), "None")`],
            explain: `Adding the two tests gives 1 for rows of either region and 0 for the rest, and FILTER keeps every row that is not 0. Multiplying would ask for both regions on one row, which never happens.`,
            chips: [out, re, ')+(', ',"None")'], fnKey: 'FILTER('
          });
        }
      ][v]();
    }
  });

  P.add({
    key: 'sort-a-list', difficulty: 'Medium', level: 6, topics: ['SORT', 'SORTBY'], count: 3,
    make(r, v) {
      return [
        () => {
          const rows = untie(H.stock(r, 12), 'onhand', 10);
          const ord = idx(12).sort((a, b) => rows[a].onhand - rows[b].onhand);
          const S = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Product', k: 'product' }, { h: 'Warehouse', k: 'warehouse' }, { h: 'On hand', k: 'onhand', num: true }], rows);
          const base = Object.assign({}, S.cells);
          const top = rows[ord[11]].onhand;
          /* The hidden sheet: the lowest line is restocked past the top, and the top line sells down. */
          const alt = H.edit(base, { [S.cell('On hand', ord[0])]: top + 20, [S.cell('On hand', ord[11])]: 5 });
          const all = blk(S, 'SKU', 'On hand'), oh = S.R('On hand');
          return H.problem({
            title: 'Sort stock lines by on hand, lowest first', file: 'stok_gudang.xlsx', sheet: S, base, alt,
            spill: { cols: 4, rows: rows.length }, spillAnchor: 'F5', ansLabel: 'Sorted list',
            statement: `<p>Bu Lestari wants the whole stock list (${all}) sorted by <b>On hand</b>, lowest first. In <b>F5</b>, type one formula; the sorted table spills across four columns and down.</p>`,
            solution: `=SORT(${all},4,1)`,
            alts: [`=SORT(${all},4)`, `=SORTBY(${all},${oh},1)`, `=SORTBY(${all},${oh})`],
            hints: ['SORT(array, sort_index, sort_order): sort_index is the column number inside the array, not the sheet column letter.', 'On hand is the 4th column of the block, and 1 means smallest first.', `=SORT(${all}, 4, 1)`],
            explain: `SORT reorders every row of the block by its 4th column, smallest first, and keeps each row together. The data itself stays where it is.`,
            chips: [all, oh, ',4,1)'], fnKey: 'SORT('
          });
        },
        () => {
          const rows = untieRatio(H.reps(r, 12));
          const ord = idx(rows.length).sort((a, b) => rows[b].actual / rows[b].target - rows[a].actual / rows[a].target);
          const S = H.table([{ h: 'Name', k: 'name' }, { h: 'Area', k: 'area' }, { h: 'Target (Rp)', k: 'target', num: true }, { h: 'Actual (Rp)', k: 'actual', num: true }], rows);
          const base = Object.assign({}, S.cells);
          const t0 = rows[ord[0]].target, tz = rows[ord[ord.length - 1]].target;
          /* The hidden sheet: the leader slips and the last rep has a strong month. */
          const alt = H.edit(base, { [S.cell('Actual (Rp)', ord[0])]: Math.round(t0 * 0.7 / 100000) * 100000 + 300000, [S.cell('Actual (Rp)', ord[ord.length - 1])]: Math.round(tz * 1.45 / 100000) * 100000 + 300000 });
          const nm = S.R('Name'), tg = S.R('Target (Rp)'), ac = S.R('Actual (Rp)');
          return H.problem({
            title: 'Sort sales reps by share of target reached', file: 'tim_sales.xlsx', sheet: S, base, alt,
            spill: { cols: 1, rows: rows.length }, spillAnchor: 'F5', ansLabel: 'Ranking',
            statement: `<p>Pak Anton ranks reps by how much of their target they reached: <b>Actual</b> divided by <b>Target</b>. In <b>F5</b>, list the <b>Name</b>s from the highest share to the lowest with one formula. There is no share column; work it out inside the formula.</p>`,
            solution: `=SORTBY(${nm},${ac}/${tg},-1)`,
            alts: [`=CHOOSECOLS(SORT(HSTACK(${nm},${ac}/${tg}),2,-1),1)`, `=INDEX(SORT(HSTACK(${nm},${ac}/${tg}),2,-1),,1)`],
            hints: ['SORTBY sorts one range by the values of another range, and that other range can be a calculation.', `${ac}/${tg} gives each rep's share. Sort the names by it, with -1 for largest first.`, `=SORTBY(${nm}, ${ac}/${tg}, -1)`],
            explain: `${ac}/${tg} makes a list of twelve shares in memory, and SORTBY orders the names by that list, largest first. No helper column is needed.`,
            chips: [nm, ac, tg, ',-1)'], fnKey: 'SORTBY('
          });
        },
        () => {
          const rows = H.monthly(r);
          const seen = new Set();
          rows.forEach((x) => { while (seen.has(x.sales - x.cost)) x.cost -= 1000000; seen.add(x.sales - x.cost); });
          const ord = idx(12).sort((a, b) => (rows[b].sales - rows[b].cost) - (rows[a].sales - rows[a].cost));
          const S = H.table([{ h: 'Month', k: 'month' }, { h: 'Sales (Rp)', k: 'sales', num: true }, { h: 'Returns (Rp)', k: 'returns', num: true }, { h: 'Cost (Rp)', k: 'cost', num: true }], rows);
          const base = Object.assign({}, S.cells);
          /* The hidden sheet: the best month's cost is restated up, the worst month's down. */
          const alt = H.edit(base, { [S.cell('Cost (Rp)', ord[0])]: (x) => x + 157000000, [S.cell('Cost (Rp)', ord[11])]: (x) => x - 123000000 });
          const mo = S.R('Month'), sa = S.R('Sales (Rp)'), co = S.R('Cost (Rp)');
          return H.problem({
            title: 'Sort months by margin, best first', file: 'laba_bulanan.xlsx', sheet: S, base, alt,
            spill: { cols: 1, rows: 12 }, spillAnchor: 'F5', ansLabel: 'Best to worst',
            statement: `<p>Tika wants the months ordered by margin, best first, where margin is <b>Sales</b> minus <b>Cost</b> (leave Returns out). In <b>F5</b>, list the <b>Month</b> names with one formula.</p>`,
            solution: `=SORTBY(${mo},${sa}-${co},-1)`,
            alts: [`=CHOOSECOLS(SORT(HSTACK(${mo},${sa}-${co}),2,-1),1)`, `=SORTBY(${mo},${co}-${sa},1)`],
            hints: ['The sort key does not have to be a column on the sheet: it can be a calculation on two columns.', `${sa}-${co} gives the twelve margins. Sort the months by them, largest first.`, `=SORTBY(${mo}, ${sa}-${co}, -1)`],
            explain: `${sa}-${co} works out each month's margin in memory, and SORTBY orders the month names by it with -1 for largest first.`,
            chips: [mo, sa, co, ',-1)'], fnKey: 'SORTBY('
          });
        }
      ][v]();
    }
  });

  P.add({
    key: 'unique-values', difficulty: 'Medium', level: 6, topics: ['UNIQUE', 'SORT', 'FILTER'], count: 3,
    make(r, v) {
      return [
        () => {
          const ds = r.sample(H.db().distributors, 7).map((d) => d.name);
          const rows = salesFrom(r, [[20, isIn('distributor', ds.slice(0, 6))]]);
          const S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'Invoice', k: 'invoice' }, { h: 'Distributor', k: 'distributor' }, { h: 'Product', k: 'product' }, { h: 'Net (Rp)', k: 'net', num: true }], rows);
          const many = tally(rows, 'distributor')[0];
          const i = rows.findIndex((x) => x.distributor === many);
          const base = Object.assign({}, S.cells);
          /* The hidden sheet: one line moves to a distributor not on the list yet. */
          const alt = H.edit(base, { [S.cell('Distributor', i)]: ds[6] });
          const dis = S.R('Distributor');
          return H.problem({
            title: 'List the distinct distributors from A to Z', file: 'penjualan_q4.xlsx', sheet: S, base, alt,
            spill: { cols: 1, rows: rows.length }, spillAnchor: 'G5', ansLabel: 'Distributors',
            statement: `<p>Tika needs each distributor in this order list once, sorted A to Z. In <b>G5</b>, type one formula that lists them from <b>${dis}</b>.</p>`,
            solution: `=SORT(UNIQUE(${dis}))`,
            alts: [`=UNIQUE(SORT(${dis}))`, `=SORT(UNIQUE(${S.A('Distributor')}))`],
            hints: ['UNIQUE returns each value of a range once, in the order it first appears.', 'Wrap it in SORT to put the names in A to Z order.', `=SORT(UNIQUE(${dis}))`],
            explain: `UNIQUE drops the repeats from ${dis}, and SORT puts what is left in A to Z order. A new distributor on the sheet joins the list on its own.`,
            chips: [dis], fnKey: 'UNIQUE('
          });
        },
        () => {
          const prods = r.sample(H.PRODUCTS, 8);
          const rows = salesFrom(r, [[20, isIn('sku', prods.map((p) => p.sku))]]);
          const out = H.PRODUCTS.find((p) => !prods.includes(p));
          const S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'Distributor', k: 'distributor' }, { h: 'SKU', k: 'sku' }, { h: 'Product', k: 'product' }, { h: 'Cartons', k: 'cartons', num: true }], rows);
          const many = tally(rows, 'sku')[0];
          const i = rows.findIndex((x) => x.sku === many);
          const base = Object.assign({}, S.cells, { A1: 'Distinct products' });
          /* The hidden sheet: one line is corrected to a product that was not ordered before. */
          const alt = H.edit(base, { [S.cell('SKU', i)]: out.sku, [S.cell('Product', i)]: out.name });
          const pr = S.R('Product');
          return H.problem({
            title: 'Count the distinct products ordered', file: 'penjualan_q4.xlsx', sheet: S, base, alt, answerCell: 'B1',
            statement: `<p>How many different products appear in this order list? In <b>B1</b>, count each product in <b>${pr}</b> once, however many lines it has.</p>`,
            solution: `=COUNTA(UNIQUE(${pr}))`,
            alts: [`=ROWS(UNIQUE(${pr}))`, `=COUNTA(UNIQUE(${S.R('SKU')}))`],
            hints: ['COUNTA on the whole column counts lines, not products. Remove the repeats first.', 'UNIQUE gives the list of products once each; then count that list.', `=COUNTA(UNIQUE(${pr}))`],
            explain: `UNIQUE(${pr}) returns each product once, and COUNTA counts how many items that list has. ROWS works as well, since the list is one column.`,
            chips: [pr], fnKey: 'UNIQUE('
          });
        },
        () => {
          const rows = H.sales(r, 20);
          const info = H.CATEGORIES.map((c) => { const f = rows.filter((x) => x.category === c); return { c, n: f.length, u: new Set(f.map((x) => x.region)).size, list: [...new Set(f.map((x) => x.region))].join('|') }; }).filter((x) => x.n > 0);
          info.sort((a, b) => (b.n - b.u) - (a.n - a.u) || b.u - a.u);
          const cat = info[0], other = info.find((x) => x.list !== cat.list) || info[1];
          const S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'Distributor', k: 'distributor' }, { h: 'Region', k: 'region' }, { h: 'Category', k: 'category' }, { h: 'Net (Rp)', k: 'net', num: true }], rows);
          const base = Object.assign({}, S.cells, { A1: 'Category', B1: cat.c });
          const alt = H.edit(base, { B1: other.c });
          const re = S.R('Region'), ca = S.R('Category');
          return H.problem({
            title: 'List the regions that bought one category', file: 'penjualan_q4.xlsx', sheet: S, base, alt,
            spill: { cols: 1, rows: rows.length }, spillAnchor: 'G5', ansLabel: 'Regions',
            statement: `<p>Pak Anton wants the regions that bought the category in <b>B1</b>, each region once, in the order they first appear. In <b>G5</b>, type one formula.</p>`,
            solution: `=UNIQUE(FILTER(${re},${ca}=B1))`,
            alts: [`=UNIQUE(FILTER(${S.A('Region')},${S.A('Category')}=$B$1))`],
            hints: ['Two steps: first keep the rows of that category, then remove the repeated regions.', `FILTER(${re}, ${ca}=B1) gives the regions with repeats; UNIQUE around it keeps each once.`, `=UNIQUE(FILTER(${re}, ${ca}=B1))`],
            explain: `FILTER returns the region of every line in that category, and UNIQUE keeps the first appearance of each region.`,
            chips: [re, ca, '=B1'], fnKey: 'UNIQUE('
          });
        }
      ][v]();
    }
  });

  P.add({
    key: 'sumproduct-several-rules', difficulty: 'Medium', level: 7, topics: ['SUMPRODUCT'], count: 3,
    make(r, v) {
      return [
        () => {
          const rows = H.sales(r, 16);
          const combos = tally(rows.map((x) => ({ k: x.region + '|' + x.category })), 'k');
          const [rg, ct] = combos[0].split('|'), [rg2, ct2] = combos[1].split('|');
          const S = H.table([{ h: 'Region', k: 'region' }, { h: 'Category', k: 'category' }, { h: 'Product', k: 'product' }, { h: 'Cartons', k: 'cartons', num: true }, { h: 'Carton price (Rp)', k: 'price', num: true }], rows);
          const j = rows.findIndex((x) => x.region === rg2 && x.category === ct2);
          const base = Object.assign({}, S.cells, { A1: 'Region', B1: rg, A2: 'Category', B2: ct, A3: 'Value (Rp)' });
          const alt = H.edit(base, { B1: rg2, B2: ct2, [S.cell('Cartons', j)]: (x) => x + 5 });
          const re = S.R('Region'), ca = S.R('Category'), cr = S.R('Cartons'), pr = S.R('Carton price (Rp)');
          return H.problem({
            title: 'Carton value for one region and category', file: 'penjualan_q4.xlsx', sheet: S, base, alt, answerCell: 'B3',
            statement: `<p>There is no value column, only <b>Cartons</b> and <b>Carton price</b>. In <b>B3</b>, total Cartons times Carton price over the lines of the region in <b>B1</b> and the category in <b>B2</b>, in one formula.</p>`,
            solution: `=SUMPRODUCT((${re}=B1)*(${ca}=B2)*${cr}*${pr})`,
            alts: [`=SUMPRODUCT((${re}=B1)*(${ca}=B2),${cr},${pr})`, `=SUM(FILTER(${cr}*${pr},(${re}=B1)*(${ca}=B2)))`],
            hints: ['SUMIFS needs a column to add, and the line value is not on the sheet. SUMPRODUCT can multiply as it adds.', 'Each test in brackets gives 1 or 0. Multiply the tests by Cartons and by Carton price, then SUMPRODUCT adds it all.', `=SUMPRODUCT((${re}=B1)*(${ca}=B2)*${cr}*${pr})`],
            explain: `The two tests turn into 1 on matching lines and 0 elsewhere, so only those lines keep their Cartons times price. SUMPRODUCT adds the products in one pass.`,
            chips: [re, ca, cr, pr, ')*('], fnKey: 'SUMPRODUCT('
          });
        },
        () => {
          const rows = book(r, 14, 4);
          const wc = idx(14).filter((i) => rows[i].type === 'Working Capital'), tl = idx(14).filter((i) => rows[i].type === 'Term Loan');
          const [w1, w2, w3] = r.sample(wc, 3);
          rows[w1].dpd = 0; rows[w2].dpd = 0; rows[w3].dpd = 30;
          rows[tl[0]].dpd = 0; rows[tl[1]].dpd = 45;
          const S = H.table([{ h: 'Facility', k: 'facility' }, { h: 'Type', k: 'type' }, { h: 'Limit (Rp)', k: 'limit', num: true }, { h: 'Outstanding (Rp)', k: 'outstanding', num: true }, { h: 'DPD', k: 'dpd', num: true }], rows);
          const base = Object.assign({}, S.cells, { A1: 'Type', B1: 'Working Capital', A2: 'Undrawn (Rp)' });
          const alt = H.edit(base, { B1: 'Term Loan' });
          const ty = S.R('Type'), li = S.R('Limit (Rp)'), ou = S.R('Outstanding (Rp)'), dp = S.R('DPD');
          return H.problem({
            title: 'Undrawn credit of one loan type with no days past due', file: 'buku_pinjaman.xlsx', sheet: S, base, alt, answerCell: 'B2',
            statement: `<p>Dimas wants the room left to draw: <b>Limit</b> minus <b>Outstanding</b>, added up over the facilities of the type in <b>B1</b> that have a <b>DPD</b> of 0. Put it in <b>B2</b> with one formula.</p>`,
            solution: `=SUMPRODUCT((${ty}=B1)*(${dp}=0)*(${li}-${ou}))`,
            alts: [`=SUMIFS(${li},${ty},B1,${dp},0)-SUMIFS(${ou},${ty},B1,${dp},0)`, `=SUMPRODUCT((${ty}=B1)*(${dp}=0),${li}-${ou})`],
            hints: ['The undrawn amount is a calculation, Limit minus Outstanding, and SUMPRODUCT can work it out row by row.', `Multiply (${ty}=B1) by (${dp}=0) by (${li}-${ou}).`, `=SUMPRODUCT((${ty}=B1)*(${dp}=0)*(${li}-${ou}))`],
            explain: `${li}-${ou} gives each facility's undrawn room, and the two tests keep it only for current facilities of that type. SUMPRODUCT adds what is left.`,
            chips: [ty, dp, li, ou, ')*('], fnKey: 'SUMPRODUCT('
          });
        },
        () => {
          const rows = H.invoices(r, 14).sort(byDate);
          [3, 5, 7, 9].forEach((i) => { rows[i].paid = ''; });
          rows[4].paid = rows[4].date;
          const S = H.table([{ h: 'Ref', k: 'ref' }, { h: 'Distributor', k: 'distributor' }, { h: 'Date', k: 'date', date: true }, { h: 'Amount (Rp)', k: 'amount', num: true }, { h: 'Paid on', k: 'paid', date: true }], rows);
          const base = Object.assign({}, S.cells, { A1: 'From', B1: ser(rows[3].date), A2: 'To', B2: ser(rows[10].date), A3: 'Unpaid (Rp)' });
          /* The hidden sheet: an earlier window, and one unpaid invoice in it has been paid. */
          const alt = H.edit(base, { B1: ser(rows[1].date), B2: ser(rows[8].date), [S.cell('Paid on', 5)]: ser(rows[5].date) + 9 });
          const dt = S.R('Date'), am = S.R('Amount (Rp)'), pd = S.R('Paid on');
          return H.problem({
            title: 'Unpaid invoice total inside a date window', file: 'piutang_q4.xlsx', sheet: S, base, alt, answerCell: 'B3', dateCells: ['B1', 'B2'],
            statement: `<p>Bu Lestari wants the total <b>Amount</b> of unpaid invoices (<b>Paid on</b> is blank) dated from <b>B1</b> to <b>B2</b>, both days included. Put it in <b>B3</b> with one formula.</p>`,
            solution: `=SUMPRODUCT((${dt}>=B1)*(${dt}<=B2)*(${pd}="")*${am})`,
            alts: [`=SUMIFS(${am},${dt},">="&B1,${dt},"<="&B2,${pd},"")`, `=SUMPRODUCT((${dt}>=B1)*(${dt}<=B2)*(${pd}=""),${am})`],
            hints: ['A date window is two tests on the same column: on or after the start, and on or before the end.', `Add a third test for unpaid rows, ${pd}="", and multiply all three by ${am}.`, `=SUMPRODUCT((${dt}>=B1)*(${dt}<=B2)*(${pd}="")*${am})`],
            explain: `Dates are numbers, so >= and <= mark the window, and ${pd}="" marks unpaid invoices. Multiplied together they keep only the amounts that pass all three, and SUMPRODUCT adds them.`,
            chips: [dt, pd, am, '>=B1', '<=B2', '=""'], fnKey: 'SUMPRODUCT('
          });
        }
      ][v]();
    }
  });

  /* ======================================================================================== */
  /* Hard                                                                                      */
  /* ======================================================================================== */

  P.add({
    key: 'top-n-with-take', difficulty: 'Hard', level: 8, topics: ['TAKE', 'SORT', 'SORTBY', 'FILTER'], count: 4,
    make(r, v) {
      return [
        () => {
          const rows = untie(H.sales(r, 16), 'net', 100);
          const lo = idx(16).reduce((m, i) => (rows[i].net < rows[m].net ? i : m), 0);
          const hi = Math.max(...rows.map((x) => x.net));
          const S = H.table([{ h: 'Invoice', k: 'invoice' }, { h: 'Distributor', k: 'distributor' }, { h: 'Product', k: 'product' }, { h: 'Net (Rp)', k: 'net', num: true }], rows);
          const base = Object.assign({}, S.cells, { A1: 'How many', B1: 3 });
          /* The hidden sheet: top 5, and the smallest order is corrected to the biggest. */
          const alt = H.edit(base, { B1: 5, [S.cell('Net (Rp)', lo)]: hi + 1000000 });
          const all = blk(S, 'Invoice', 'Net (Rp)'), ne = S.R('Net (Rp)');
          return H.problem({
            title: 'Top N orders by net value, whole rows', file: 'penjualan_q4.xlsx', sheet: S, base, alt,
            spill: { cols: 4, rows: 5 }, spillAnchor: 'F5', ansLabel: 'Top orders',
            statement: `<p>Pak Anton wants the biggest orders: the number of rows in <b>B1</b> (at most 5), largest <b>Net</b> first, with all four columns of ${all}. In <b>F5</b>, type one formula.</p>`,
            solution: `=TAKE(SORT(${all},4,-1),B1)`,
            alts: [`=TAKE(SORTBY(${all},${ne},-1),B1)`, `=LET(s,SORT(${all},4,-1),TAKE(s,B1))`],
            hints: ['Sort the table by Net, largest first, then keep the first rows.', 'TAKE(array, rows) keeps that many rows from the top of an array.', `=TAKE(SORT(${all}, 4, -1), B1)`],
            explain: `SORT puts the largest net on top, and TAKE keeps the first B1 rows of the sorted table. Changing B1 changes how many orders show.`,
            chips: [all, ne, ',4,-1)', ',B1)'], fnKey: 'TAKE('
          });
        },
        () => {
          const rows = untie(H.stock(r, 12), 'onhand', 10);
          const ord = idx(12).sort((a, b) => rows[a].onhand - rows[b].onhand);
          const S = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Product', k: 'product' }, { h: 'Warehouse', k: 'warehouse' }, { h: 'On hand', k: 'onhand', num: true }], rows);
          const base = Object.assign({}, S.cells, { A1: 'How many', B1: 3 });
          /* The hidden sheet: the four lowest, after the lowest line is restocked. */
          const alt = H.edit(base, { B1: 4, [S.cell('On hand', ord[0])]: rows[ord[11]].onhand + 30 });
          const pr = S.R('Product'), oh = S.R('On hand');
          return H.problem({
            title: 'Bottom N products by stock on hand', file: 'stok_gudang.xlsx', sheet: S, base, alt,
            spill: { cols: 1, rows: 4 }, spillAnchor: 'F5', ansLabel: 'Lowest stock',
            statement: `<p>Bu Lestari wants the <b>Product</b> names of the lines with the least stock: as many as the number in <b>B1</b> (at most 4), lowest <b>On hand</b> first. In <b>F5</b>, type one formula.</p>`,
            solution: `=TAKE(SORTBY(${pr},${oh},1),B1)`,
            alts: [`=TAKE(SORTBY(${pr},${oh}),B1)`, `=CHOOSECOLS(TAKE(SORT(${blk(S, 'SKU', 'On hand')},4),B1),2)`],
            hints: ['Sort the product names by stock, smallest first, then keep the top of the list.', `SORTBY(${pr}, ${oh}, 1) puts the lowest stock first. TAKE keeps B1 rows.`, `=TAKE(SORTBY(${pr}, ${oh}, 1), B1)`],
            explain: `SORTBY orders the product names by their on-hand figures, lowest first, and TAKE keeps the first B1 of them.`,
            chips: [pr, oh, ',1)', ',B1)'], fnKey: 'TAKE('
          });
        },
        () => {
          const rows = untieRatio(H.reps(r, 12));
          const best = idx(rows.length).reduce((m, i) => (rows[i].actual / rows[i].target > rows[m].actual / rows[m].target ? i : m), 0);
          const S = H.table([{ h: 'Name', k: 'name' }, { h: 'Area', k: 'area' }, { h: 'Target (Rp)', k: 'target', num: true }, { h: 'Actual (Rp)', k: 'actual', num: true }], rows);
          const base = Object.assign({}, S.cells, { A1: 'How many', B1: 3 });
          /* The hidden sheet: top 2, and the leader's month is restated down. */
          const alt = H.edit(base, { B1: 2, [S.cell('Actual (Rp)', best)]: Math.round(rows[best].target * 0.66 / 100000) * 100000 + 300000 });
          const nm = S.R('Name'), tg = S.R('Target (Rp)'), ac = S.R('Actual (Rp)');
          return H.problem({
            title: 'Top N sales reps by attainment, with the ratio', file: 'tim_sales.xlsx', sheet: S, base, alt, precision: 4,
            spill: { cols: 2, rows: 3 }, spillAnchor: 'F5', ansLabel: 'Top reps',
            statement: `<p>Pak Anton wants the leading reps by <b>Actual</b> divided by <b>Target</b>: as many as <b>B1</b> (at most 3), highest first, each with its ratio as a decimal (1.12 means 112%). In <b>F5</b>, return two columns, the <b>Name</b> and the ratio, with one formula.</p>`,
            solution: `=TAKE(SORT(HSTACK(${nm},${ac}/${tg}),2,-1),B1)`,
            alts: [`=TAKE(SORTBY(HSTACK(${nm},${ac}/${tg}),${ac}/${tg},-1),B1)`, `=LET(a,${ac}/${tg},TAKE(SORTBY(HSTACK(${nm},a),a,-1),B1))`],
            hints: ['Build the two-column table first: HSTACK puts the names and the ratios side by side.', 'Sort that table by its 2nd column, largest first, then TAKE the first B1 rows.', `=TAKE(SORT(HSTACK(${nm}, ${ac}/${tg}), 2, -1), B1)`],
            explain: `HSTACK pairs each name with its ratio, SORT orders the pairs by the ratio, largest first, and TAKE keeps the first B1 pairs.`,
            chips: [nm, ac, tg, 'HSTACK(', ',2,-1)', ',B1)'], fnKey: 'TAKE('
          });
        },
        () => {
          const rows = untie(H.sales(r, 20), 'cartons', 1);
          const t = tally(rows, 'category');
          const c1 = t[0], c2 = t[1], n2 = rows.filter((x) => x.category === c2).length;
          const S = H.table([{ h: 'Invoice', k: 'invoice' }, { h: 'Distributor', k: 'distributor' }, { h: 'Category', k: 'category' }, { h: 'Cartons', k: 'cartons', num: true }], rows);
          const base = Object.assign({}, S.cells, { A1: 'Category', B1: c1, A2: 'How many', B2: 3 });
          const alt = H.edit(base, { B1: c2, B2: Math.min(3, n2) });
          const all = blk(S, 'Invoice', 'Cartons'), ca = S.R('Category'), cr = S.R('Cartons');
          return H.problem({
            title: 'Top N orders by cartons within one category', file: 'penjualan_q4.xlsx', sheet: S, base, alt,
            spill: { cols: 4, rows: 4 }, spillAnchor: 'F5', ansLabel: 'Top in category',
            statement: `<p>Pak Anton wants the biggest orders of the category in <b>B1</b>: as many as <b>B2</b>, most <b>Cartons</b> first, with all four columns of ${all}. In <b>F5</b>, type one formula. The category always has at least B2 orders.</p>`,
            solution: `=TAKE(SORT(FILTER(${all},${ca}=B1),4,-1),B2)`,
            alts: [`=LET(f,FILTER(${all},${ca}=B1),TAKE(SORT(f,4,-1),B2))`, `=TAKE(SORTBY(FILTER(${all},${ca}=B1),FILTER(${cr},${ca}=B1),-1),B2)`],
            hints: ['Three steps, inside out: keep the category, sort by cartons, keep the top rows.', `FILTER(${all}, ${ca}=B1) keeps the category. SORT it by column 4, largest first.`, `=TAKE(SORT(FILTER(${all}, ${ca}=B1), 4, -1), B2)`],
            explain: `FILTER keeps the orders of that category, SORT puts the most cartons on top, and TAKE keeps the first B2 rows.`,
            chips: [all, ca, ',4,-1)', ',B2)'], fnKey: 'TAKE('
          });
        }
      ][v]();
    }
  });

  P.add({
    key: 'let-named-steps', difficulty: 'Hard', level: 8, topics: ['LET', 'SUMIFS', 'XLOOKUP'], count: 4,
    make(r, v) {
      return [
        () => {
          const D = H.db().distributors.filter((d) => d.tier_discount > 0);
          const who = r.pick(D), next = r.pick(D.filter((d) => d.tier_discount !== who.tier_discount));
          const others = r.sample(H.db().distributors.filter((d) => d !== who && d !== next), 3).map((d) => d.name);
          const rows = salesFrom(r, [[5, (s) => s.distributor === who.name], [4, (s) => s.distributor === next.name], [7, isIn('distributor', others)]]);
          const S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'Distributor', k: 'distributor' }, { h: 'Product', k: 'product' }, { h: 'Gross (Rp)', k: 'gross', num: true }, { h: 'Net (Rp)', k: 'net', num: true }], rows);
          const base = Object.assign({}, S.cells, { A1: 'Distributor', B1: who.name, A2: 'Discount rate' });
          const alt = H.edit(base, { B1: next.name });
          const di = S.R('Distributor'), gr = S.R('Gross (Rp)'), ne = S.R('Net (Rp)');
          return H.problem({
            title: 'Effective discount rate of one distributor', file: 'penjualan_q4.xlsx', sheet: S, base, alt, answerCell: 'B2', precision: 4,
            statement: `<p>Tika wants the discount the distributor in <b>B1</b> really got: 1 minus its total <b>Net</b> divided by its total <b>Gross</b>, as a decimal (0.03 means 3%). In <b>B2</b>, use LET to name the two totals, then combine them.</p>`,
            solution: `=LET(g,SUMIFS(${gr},${di},B1),n,SUMIFS(${ne},${di},B1),1-n/g)`,
            alts: [`=1-SUMIFS(${ne},${di},B1)/SUMIFS(${gr},${di},B1)`, `=(SUMIFS(${gr},${di},B1)-SUMIFS(${ne},${di},B1))/SUMIFS(${gr},${di},B1)`, `=LET(f,${di}=B1,1-SUM(f*${ne})/SUM(f*${gr}))`],
            hints: ['LET(name, value, ..., result) works out each named value once and then uses the names in the result.', `Name g as SUMIFS of ${gr} for B1, and n as SUMIFS of ${ne} for B1. The result is 1-n/g.`, `=LET(g, SUMIFS(${gr}, ${di}, B1), n, SUMIFS(${ne}, ${di}, B1), 1-n/g)`],
            explain: `g and n are the distributor's gross and net totals, each worked out once. 1-n/g is the share of gross given away as discount, and the names make the last step easy to read.`,
            chips: [gr, ne, di, 'LET(', '1-n/g)'], fnKey: 'LET('
          });
        },
        () => {
          const rows = H.loans(r, 14);
          const [a, b] = r.sample(idx(14), 2);
          rows[a].outstanding = Math.round(rows[a].limit * 0.5 / 100000) * 100000;
          rows[b].outstanding = Math.round(rows[b].limit * 0.9 / 100000) * 100000;
          const draw = Math.round(rows[a].limit * r.pick([0.1, 0.15, 0.2, 0.25]) / 1000000) * 1000000;
          const S = H.table([{ h: 'Facility', k: 'facility' }, { h: 'Distributor', k: 'distributor' }, { h: 'Limit (Rp)', k: 'limit', num: true }, { h: 'Outstanding (Rp)', k: 'outstanding', num: true }], rows);
          const base = Object.assign({}, S.cells, { A1: 'Facility', B1: rows[a].facility, A2: 'Drawdown (Rp)', B2: draw, A3: 'Utilisation after' });
          /* The hidden sheet: a nearly full line and a drawdown that would pass its limit. */
          const alt = H.edit(base, { B1: rows[b].facility, B2: Math.round(rows[b].limit * 0.3 / 1000000) * 1000000 });
          const fa = S.R('Facility'), li = S.R('Limit (Rp)'), ou = S.R('Outstanding (Rp)');
          return H.problem({
            title: 'Utilisation after a planned drawdown, capped at 100%', file: 'buku_pinjaman.xlsx', sheet: S, base, alt, answerCell: 'B3', precision: 4,
            statement: `<p>The distributor on facility <b>B1</b> plans to draw <b>B2</b> more. In <b>B3</b>, show the utilisation after it: (<b>Outstanding</b> plus B2) divided by <b>Limit</b>, as a decimal. The bank never lends past the limit, so the answer is at most 1. Use LET to name the limit and the outstanding.</p>`,
            solution: `=LET(l,XLOOKUP(B1,${fa},${li}),o,XLOOKUP(B1,${fa},${ou}),MIN(1,(o+B2)/l))`,
            alts: [`=MIN(1,(XLOOKUP(B1,${fa},${ou})+B2)/XLOOKUP(B1,${fa},${li}))`, `=LET(l,VLOOKUP(B1,${blk(S, 'Facility', 'Outstanding (Rp)')},3,FALSE),o,VLOOKUP(B1,${blk(S, 'Facility', 'Outstanding (Rp)')},4,FALSE),MIN(1,(o+B2)/l))`, `=LET(l,XLOOKUP(B1,${fa},${li}),o,XLOOKUP(B1,${fa},${ou}),IF(o+B2>l,1,(o+B2)/l))`],
            hints: ['Look up the facility\'s limit and outstanding once each, and give them names with LET.', 'The capped ratio is MIN(1, ratio): MIN picks the smaller of 1 and the ratio.', `=LET(l, XLOOKUP(B1, ${fa}, ${li}), o, XLOOKUP(B1, ${fa}, ${ou}), MIN(1, (o+B2)/l))`],
            explain: `LET looks up l and o once, (o+B2)/l is the new utilisation, and MIN(1, ...) caps it at the limit. On the hidden sheet the drawdown would pass the limit, so the answer is 1.`,
            chips: [fa, li, ou, 'XLOOKUP(B1,', 'MIN(1,'], fnKey: 'LET('
          });
        },
        () => {
          const rows = untie(H.monthly(r), 'sales', 1000000);
          const best = idx(12).reduce((m, i) => (rows[i].sales > rows[m].sales ? i : m), 0);
          const ratio = rows[best].cost / rows[best].sales;
          const j = r.pick(idx(12).filter((i) => i !== best));
          const newSales = rows[best].sales + 25000000;
          const nr = [0.55, 0.6, 0.65, 0.7].find((x) => Math.abs(x - ratio) > 0.03);
          const S = H.table([{ h: 'Month', k: 'month' }, { h: 'Sales (Rp)', k: 'sales', num: true }, { h: 'Returns (Rp)', k: 'returns', num: true }, { h: 'Cost (Rp)', k: 'cost', num: true }], rows);
          const base = Object.assign({}, S.cells, { A1: 'Margin of best month' });
          /* The hidden sheet: another month is restated above the best, with a different cost ratio. */
          const alt = H.edit(base, { [S.cell('Sales (Rp)', j)]: newSales, [S.cell('Cost (Rp)', j)]: Math.round(newSales * nr / 1000000) * 1000000 });
          const sa = S.R('Sales (Rp)'), co = S.R('Cost (Rp)');
          return H.problem({
            title: 'Margin of the best sales month', file: 'laba_bulanan.xlsx', sheet: S, base, alt, answerCell: 'B1', precision: 4,
            statement: `<p>Tika wants the margin of the month with the highest <b>Sales</b>: (Sales minus <b>Cost</b>) divided by Sales, as a decimal. Leave Returns out. In <b>B1</b>, type one formula; LET helps you name the month's position once and reuse it.</p>`,
            solution: `=LET(i,XMATCH(MAX(${sa}),${sa}),s,INDEX(${sa},i),(s-INDEX(${co},i))/s)`,
            alts: [`=(MAX(${sa})-XLOOKUP(MAX(${sa}),${sa},${co}))/MAX(${sa})`, `=LET(s,MAX(${sa}),c,XLOOKUP(s,${sa},${co}),(s-c)/s)`, `=1-XLOOKUP(MAX(${sa}),${sa},${co})/MAX(${sa})`],
            hints: ['First find the best month: MAX gives its sales, and XMATCH of that value gives its position.', 'With the position i, INDEX(range, i) reads that month\'s sales and cost.', `=LET(i, XMATCH(MAX(${sa}), ${sa}), s, INDEX(${sa}, i), (s-INDEX(${co}, i))/s)`],
            explain: `XMATCH finds which month holds the highest sales, and LET keeps that position as i. INDEX reads the month's sales and cost, and (s - cost)/s is its margin.`,
            chips: [sa, co, 'MAX(', 'XMATCH(', 'INDEX('], fnKey: 'LET('
          });
        },
        () => {
          const rows = H.stock(r, 12);
          const [a, b] = r.sample(idx(12), 2);
          rows[a].onhand = rows[a].reorder - r.int(2, 6) * 10;
          rows[b].onhand = rows[b].reorder + r.int(1, 5) * 10;
          const S = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Product', k: 'product' }, { h: 'Warehouse', k: 'warehouse' }, { h: 'On hand', k: 'onhand', num: true }, { h: 'Reorder point', k: 'reorder', num: true }], rows);
          const base = Object.assign({}, S.cells, { A1: 'SKU', B1: rows[a].sku, A2: 'Order (cartons)' });
          /* The hidden sheet: a SKU above its reorder point but under twice it, so the order is 0. */
          const alt = H.edit(base, { B1: rows[b].sku });
          const sk = S.R('SKU'), oh = S.R('On hand'), ro = S.R('Reorder point'), all = blk(S, 'SKU', 'Reorder point');
          return H.problem({
            title: 'Reorder quantity for one SKU', file: 'stok_gudang.xlsx', sheet: S, base, alt, answerCell: 'B2',
            statement: `<p>Bu Lestari's rule for the SKU in <b>B1</b>: if <b>On hand</b> is below the <b>Reorder point</b>, order enough cartons to bring stock up to twice the reorder point; otherwise order 0. In <b>B2</b>, use LET to look up the two figures once and apply the rule.</p>`,
            solution: `=LET(oh,XLOOKUP(B1,${sk},${oh}),rp,XLOOKUP(B1,${sk},${ro}),IF(oh<rp,2*rp-oh,0))`,
            alts: [`=IF(XLOOKUP(B1,${sk},${oh})<XLOOKUP(B1,${sk},${ro}),2*XLOOKUP(B1,${sk},${ro})-XLOOKUP(B1,${sk},${oh}),0)`, `=LET(oh,VLOOKUP(B1,${all},4,FALSE),rp,VLOOKUP(B1,${all},5,FALSE),IF(oh<rp,2*rp-oh,0))`],
            hints: ['Without LET, each figure is looked up two or three times. Name them oh and rp instead.', 'The order is 2*rp-oh when oh<rp, and 0 otherwise.', `=LET(oh, XLOOKUP(B1, ${sk}, ${oh}), rp, XLOOKUP(B1, ${sk}, ${ro}), IF(oh<rp, 2*rp-oh, 0))`],
            explain: `LET looks up the stock and the reorder point once each. IF orders 2*rp-oh only when stock is below the reorder point, so a SKU above it gets 0 even when it is under twice the reorder point.`,
            chips: [sk, oh, ro, 'XLOOKUP(B1,', 'IF(oh<rp,'], fnKey: 'LET('
          });
        }
      ][v]();
    }
  });

  P.add({
    key: 'sequence-schedules', difficulty: 'Hard', level: 9, topics: ['SEQUENCE', 'LET', 'XLOOKUP'], count: 4,
    make(r, v) {
      return [
        () => {
          const L = H.loans(r, 12);
          const rows = L.map((x) => { const n = r.int(3, 10); return { facility: x.facility, distributor: x.distributor, months: n, out: n * r.int(4, 30) * 1000000 }; });
          const [a, b] = r.sample(idx(12), 2);
          rows[a].months = r.int(5, 8); rows[a].out = rows[a].months * r.int(4, 30) * 1000000;
          rows[b].months = rows[a].months + r.pick([-2, 2]); rows[b].out = rows[b].months * r.int(4, 30) * 1000000;
          const S = H.table([{ h: 'Facility', k: 'facility' }, { h: 'Distributor', k: 'distributor' }, { h: 'Outstanding (Rp)', k: 'out', num: true }, { h: 'Months left', k: 'months', num: true }], rows);
          const base = Object.assign({}, S.cells, { A1: 'Facility', B1: rows[a].facility });
          const alt = H.edit(base, { B1: rows[b].facility });
          const fa = S.R('Facility'), ou = S.R('Outstanding (Rp)'), mo = S.R('Months left'), all = blk(S, 'Facility', 'Months left');
          return H.problem({
            title: 'Balance after each equal repayment for one facility', file: 'jadwal_cicilan.xlsx', sheet: S, base, alt,
            spill: { cols: 1, rows: 10 }, spillAnchor: 'F5', ansLabel: 'Balance (Rp)',
            statement: `<p>Bank Kenari repays each facility in equal monthly amounts of principal: <b>Outstanding</b> divided by <b>Months left</b>. For the facility in <b>B1</b>, list the balance left after each month, from month 1 to the last month (when it reaches 0). In <b>F5</b>, type one formula.</p>`,
            solution: `=LET(o,XLOOKUP(B1,${fa},${ou}),n,XLOOKUP(B1,${fa},${mo}),o-o/n*SEQUENCE(n))`,
            alts: [`=XLOOKUP(B1,${fa},${ou})*(1-SEQUENCE(XLOOKUP(B1,${fa},${mo}))/XLOOKUP(B1,${fa},${mo}))`, `=LET(o,VLOOKUP(B1,${all},3,FALSE),n,VLOOKUP(B1,${all},4,FALSE),o-o/n*SEQUENCE(n))`, `=LET(o,XLOOKUP(B1,${fa},${ou}),n,XLOOKUP(B1,${fa},${mo}),SEQUENCE(n,1,o-o/n,-o/n))`],
            hints: ['SEQUENCE(n) returns 1, 2, ..., n down a column: one row per month.', 'Look up the outstanding o and the months n. After month k the balance is o minus k times o/n.', `=LET(o, XLOOKUP(B1, ${fa}, ${ou}), n, XLOOKUP(B1, ${fa}, ${mo}), o-o/n*SEQUENCE(n))`],
            explain: `SEQUENCE(n) counts the months 1 to n, and o/n is one month's principal, so o-o/n*SEQUENCE(n) is the balance after each month. The list is as long as the facility's months left.`,
            chips: [fa, ou, mo, 'SEQUENCE(n)', 'XLOOKUP(B1,'], fnKey: 'SEQUENCE('
          });
        },
        () => {
          const ds = r.sample(H.db().distributors, 12);
          const rows = ds.map((d) => ({ name: d.name, city: d.city, every: r.pick([7, 10, 14]), visits: r.int(3, 8) }));
          const [a, b] = r.sample(idx(12), 2);
          rows[a].every = 7; rows[a].visits = r.int(4, 6);
          rows[b].every = 14; rows[b].visits = r.int(3, 8);
          const S = H.table([{ h: 'Distributor', k: 'name' }, { h: 'City', k: 'city' }, { h: 'Visit every (days)', k: 'every', num: true }, { h: 'Visits planned', k: 'visits', num: true }], rows);
          const base = Object.assign({}, S.cells, { A1: 'Distributor', B1: rows[a].name, A2: 'First visit', B2: ser(r.pick(['2026-11-02', '2026-11-09', '2026-11-16'])) });
          const alt = H.edit(base, { B1: rows[b].name, B2: ser(r.pick(['2026-11-04', '2026-11-11'])) });
          const nm = S.R('Distributor'), ev = S.R('Visit every (days)'), vi = S.R('Visits planned');
          return H.problem({
            title: 'Visit dates for one distributor', file: 'jadwal_kunjungan.xlsx', sheet: S, base, alt, dateCells: ['B2'], ansDate: true,
            spill: { cols: 1, rows: 8 }, spillAnchor: 'F5', ansLabel: 'Visit dates',
            statement: `<p>Pak Anton plans visits to the distributor in <b>B1</b>: the first on the date in <b>B2</b>, then one every <b>Visit every (days)</b>, for as many visits as <b>Visits planned</b>. In <b>F5</b>, list every visit date with one formula.</p>`,
            solution: `=SEQUENCE(XLOOKUP(B1,${nm},${vi}),1,B2,XLOOKUP(B1,${nm},${ev}))`,
            alts: [`=B2+XLOOKUP(B1,${nm},${ev})*(SEQUENCE(XLOOKUP(B1,${nm},${vi}))-1)`, `=LET(n,XLOOKUP(B1,${nm},${vi}),d,XLOOKUP(B1,${nm},${ev}),SEQUENCE(n,1,B2,d))`],
            hints: ['Dates are numbers, so a list of dates a fixed number of days apart is a sequence with a start and a step.', 'SEQUENCE(rows, columns, start, step): rows is the visit count, start is B2, step is the gap in days. Look both up for B1.', `=SEQUENCE(XLOOKUP(B1, ${nm}, ${vi}), 1, B2, XLOOKUP(B1, ${nm}, ${ev}))`],
            explain: `SEQUENCE counts from the first date in B2 in steps of the visit gap, for as many rows as visits planned. Both numbers come from XLOOKUP, so the plan follows B1.`,
            chips: [nm, ev, vi, 'XLOOKUP(B1,', ',1,B2,'], fnKey: 'SEQUENCE('
          });
        },
        () => {
          const rows = H.monthly(r);
          const sales = rows.map((x) => x.sales);
          /* January targets that one month meets exactly, so "at least" matters: sales and steps are whole millions. */
          const setup = (st, k) => { const need = sales.map((s, i) => s - st * i), t1 = need.slice().sort((a, b) => b - a)[k]; return { st, t1, met: need.map((x) => x >= t1).join() }; };
          const st = r.pick([5, 8, 10, 12]) * 1000000;
          const b = setup(st, r.int(3, 6));
          let a = setup(st + 3000000, r.int(2, 7));
          for (let k = 2; a.met === b.met && k <= 8; k++) a = setup(st + 3000000, k);
          const S = H.table([{ h: 'Month', k: 'month' }, { h: 'Sales (Rp)', k: 'sales', num: true }], rows);
          const base = Object.assign({}, S.cells, { A1: 'January target (Rp)', B1: b.t1, A2: 'Rise each month (Rp)', B2: b.st });
          /* The hidden sheet: another starting target and a steeper rise. */
          const alt = H.edit(base, { B1: a.t1, B2: a.st });
          const mo = S.R('Month'), sa = S.R('Sales (Rp)');
          return H.problem({
            title: 'Months that met a sales target rising every month', file: 'laba_bulanan.xlsx', sheet: S, base, alt,
            spill: { cols: 1, rows: 12 }, spillAnchor: 'D5', ansLabel: 'Met target',
            statement: `<p>The sales target is <b>B1</b> in January and rises by <b>B2</b> every month after that, so February's target is B1+B2. In <b>D5</b>, list the <b>Month</b> names whose <b>Sales</b> are at least that month's target (reaching it exactly counts), January first, with one formula. There is no target column.</p>`,
            solution: `=FILTER(${mo},${sa}>=SEQUENCE(12,1,B1,B2))`,
            alts: [`=FILTER(${mo},${sa}-SEQUENCE(12,1,B1,B2)>=0)`, `=FILTER(${mo},${sa}>=B1+B2*(SEQUENCE(12)-1))`, `=LET(t,SEQUENCE(12,1,B1,B2),FILTER(${mo},${sa}>=t))`],
            hints: ['The twelve targets are a sequence: start at B1 and step by B2. SEQUENCE(rows, columns, start, step) builds it in memory.', `Compare ${sa} with SEQUENCE(12, 1, B1, B2) row by row, and FILTER the month names on that test.`, `=FILTER(${mo}, ${sa}>=SEQUENCE(12, 1, B1, B2))`],
            explain: `SEQUENCE(12,1,B1,B2) builds the rising targets, one per month, and ${sa}>= it is TRUE for each month that reached its own target. FILTER returns those month names in sheet order.`,
            chips: [mo, sa, 'SEQUENCE(12,1,B1,B2)', '>='], fnKey: 'SEQUENCE('
          });
        },
        () => {
          const names = r.sample(H.db().distributors, 4).map((d) => d.name);
          const rows = salesFrom(r, [[16, isIn('distributor', names)]]);
          const [who, next] = tally(rows, 'distributor');
          const S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'Invoice', k: 'invoice' }, { h: 'Distributor', k: 'distributor' }, { h: 'Net (Rp)', k: 'net', num: true }], rows);
          const base = Object.assign({}, S.cells, { A1: 'Distributor', B1: who });
          const alt = H.edit(base, { B1: next });
          const inv = S.R('Invoice'), dis = S.R('Distributor');
          return H.problem({
            title: "Numbered list of one distributor's invoices", file: 'penjualan_q4.xlsx', sheet: S, base, alt,
            spill: { cols: 2, rows: rows.length }, spillAnchor: 'F5', ansLabel: 'No.',
            statement: `<p>Citra wants a numbered list for the distributor in <b>B1</b>: in <b>F5</b>, two columns, a running number 1, 2, 3 and so on, and beside it each <b>Invoice</b> of that distributor, in the order they appear on the sheet. Use one formula.</p>`,
            solution: `=LET(f,FILTER(${inv},${dis}=B1),HSTACK(SEQUENCE(ROWS(f)),f))`,
            alts: [`=HSTACK(SEQUENCE(COUNTIF(${dis},B1)),FILTER(${inv},${dis}=B1))`, `=HSTACK(SEQUENCE(ROWS(FILTER(${inv},${dis}=B1))),FILTER(${inv},${dis}=B1))`],
            hints: ['FILTER gives the invoices. The numbers need to be as many as the invoices: ROWS counts them.', 'HSTACK puts two arrays side by side. Name the filtered list with LET so you can use it twice.', `=LET(f, FILTER(${inv}, ${dis}=B1), HSTACK(SEQUENCE(ROWS(f)), f))`],
            explain: `f is the filtered invoice list. SEQUENCE(ROWS(f)) numbers it 1 to its length, and HSTACK puts the numbers to the left of the invoices.`,
            chips: [inv, dis, 'HSTACK(', 'SEQUENCE(ROWS(f))'], fnKey: 'SEQUENCE('
          });
        }
      ][v]();
    }
  });

  P.add({
    key: 'two-way-sumproduct', difficulty: 'Hard', level: 9, topics: ['SUMPRODUCT', 'XLOOKUP'], count: 4,
    make(r, v) {
      const MON = ['Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
      return [
        () => {
          /* Week numbers as headers, so the header test can be a >= and not only an exact match. */
          const WK = [40, 41, 42, 43, 44, 45];
          const rows = r.shuffle(H.PRODUCTS).map((p) => { const o = { sku: p.sku, cat: p.category }; WK.forEach((w) => { o['w' + w] = r.int(4, 60) * 5; }); return o; });
          const cat = r.pick(H.CATEGORIES), wk = r.pick([41, 42, 43, 44]);
          const S = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Category', k: 'cat' }].concat(WK.map((w) => ({ h: w, k: 'w' + w, num: true }))), rows);
          const base = Object.assign({}, S.cells, { A1: 'Category', B1: cat, A2: 'From week', B2: wk, A3: 'Cartons' });
          /* The hidden sheet: another category from another week. */
          const alt = H.edit(base, { B1: r.pick(H.CATEGORIES.filter((c) => c !== cat)), B2: r.pick([41, 42, 43, 44].filter((w) => w !== wk)) });
          const ca = S.R('Category'), hd = 'C4:H4', grid = 'C' + S.first + ':H' + S.last;
          return H.problem({
            title: 'Cartons for a category from a chosen week onward, from a grid', file: 'penjualan_mingguan.xlsx', sheet: S, base, alt, answerCell: 'B3',
            statement: `<p>The grid shows cartons sold per SKU (rows) in weeks 40 to 45 of 2026; the week numbers are the headers in <b>${hd}</b>. In <b>B3</b>, total the cartons of the SKUs in the category in <b>B1</b> from the week in <b>B2</b> to week 45, both included, in one formula.</p>`,
            solution: `=SUMPRODUCT((${ca}=B1)*(${hd}>=B2)*${grid})`,
            alts: [`=SUMPRODUCT((${ca}=B1)*(${hd}>=B2),${grid})`, `=SUM(FILTER(${grid},${ca}=B1)*(${hd}>=B2))`, `=SUM((${ca}=B1)*(${hd}>=B2)*${grid})`],
            hints: ['A test on the header row can use any comparison, not only =. The week numbers are numbers, so >= works on them.', `(${ca}=B1) picks the rows and (${hd}>=B2) picks the weeks from B2 on. Multiply both by ${grid}.`, `=SUMPRODUCT((${ca}=B1)*(${hd}>=B2)*${grid})`],
            explain: `The category test is a column of 1s and 0s and the week test is a row of them, 1 from week B2 on. Multiplied with the grid they keep only the wanted cells, and SUMPRODUCT adds them.`,
            chips: [ca, hd, grid, '>=B2', ')*('], fnKey: 'SUMPRODUCT('
          });
        },
        () => {
          const WH = ['Bekasi', 'Bandung', 'Semarang', 'Surabaya', 'Medan'];
          const rows = r.shuffle(H.PRODUCTS).map((p) => { const o = { sku: p.sku, cat: p.category, cost: Math.round(p.carton_price * 0.72 / 100) * 100 }; WH.forEach((w) => { o[w] = r.int(0, 40) * 5; }); return o; });
          const wh = r.pick(WH), cat = r.pick(H.CATEGORIES);
          const S = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Category', k: 'cat' }].concat(WH.map((w) => ({ h: w, k: w, num: true })), [{ h: 'Cost (Rp)', k: 'cost', num: true }]), rows);
          const base = Object.assign({}, S.cells, { A1: 'Warehouse', B1: wh, A2: 'Category', B2: cat, A3: 'Value at cost (Rp)' });
          const alt = H.edit(base, { B1: r.pick(WH.filter((w) => w !== wh)), B2: r.pick(H.CATEGORIES.filter((c) => c !== cat)) });
          const ca = S.R('Category'), co = S.R('Cost (Rp)'), hd = 'C4:G4', grid = 'C' + S.first + ':G' + S.last;
          return H.problem({
            title: 'Stock value at cost for a warehouse and category', file: 'stok_gudang.xlsx', sheet: S, base, alt, answerCell: 'B3',
            statement: `<p>Cartons on hand are laid out by SKU (rows) and warehouse (columns, headers in <b>${hd}</b>), with each SKU's <b>Cost</b> per carton in column H. In <b>B3</b>, find the value at cost of the stock in the warehouse in <b>B1</b> for the category in <b>B2</b>: cartons times cost, added up.</p>`,
            solution: `=SUMPRODUCT((${ca}=B2)*(${hd}=B1)*${grid}*${co})`,
            alts: [`=SUMPRODUCT((${ca}=B2)*XLOOKUP(B1,${hd},${grid})*${co})`, `=SUMPRODUCT((${ca}=B2)*XLOOKUP(B1,${hd},${grid}),${co})`],
            hints: ['Pick the warehouse column with a test on the header row, and the SKUs with a test on the Category column.', `Multiply both tests by the grid ${grid} and by the cost column ${co}: each cost multiplies across its own row.`, `=SUMPRODUCT((${ca}=B2)*(${hd}=B1)*${grid}*${co})`],
            explain: `The header test keeps one warehouse column and the category test keeps the right SKUs; multiplying by ${co} turns each kept carton count into rupiah before SUMPRODUCT adds them.`,
            chips: [ca, hd, grid, co, ')*('], fnKey: 'SUMPRODUCT('
          });
        },
        () => {
          const C = H.CATEGORIES;
          const rows = H.db().distributors.map((d) => { const o = { name: d.name, region: d.region }; C.forEach((c) => { const s = H.db().sales.filter((x) => x.distributor_id === d.id && x.category === c).reduce((a, x) => a + x.net, 0); o[c] = Math.round(s * r.pick([0.9, 1, 1.1]) / 100000) * 100000; }); return o; });
          const reg = r.pick(H.REGIONS), cat = r.pick(C);
          const S = H.table([{ h: 'Distributor', k: 'name' }, { h: 'Region', k: 'region' }].concat(C.map((c) => ({ h: c, k: c, num: true }))), rows);
          const base = Object.assign({}, S.cells, { A1: 'Region', B1: reg, A2: 'Leave out', B2: cat, A3: 'Net (Rp)' });
          const alt = H.edit(base, { B1: r.pick(H.REGIONS.filter((x) => x !== reg)), B2: r.pick(C.filter((x) => x !== cat)) });
          const re = S.R('Region'), hd = 'C4:F4', grid = 'C' + S.first + ':F' + S.last;
          return H.problem({
            title: 'Net for a region, leaving out one category', file: 'penjualan_kategori.xlsx', sheet: S, base, alt, answerCell: 'B3',
            statement: `<p>Rows are distributors and columns are categories (headers in <b>${hd}</b>), with net sales in rupiah. In <b>B3</b>, total the net of the distributors in the region in <b>B1</b> across every category except the one in <b>B2</b>.</p>`,
            solution: `=SUMPRODUCT((${re}=B1)*(${hd}<>B2)*${grid})`,
            alts: [`=SUMPRODUCT((${re}=B1)*${grid})-SUMIFS(XLOOKUP(B2,${hd},${grid}),${re},B1)`, `=SUMPRODUCT((${re}=B1)*${grid})-SUMPRODUCT((${re}=B1)*(${hd}=B2)*${grid})`],
            hints: ['"Every category except one" is a test on the header row with <> (not equal).', `(${re}=B1) picks the rows and (${hd}<>B2) the columns. Multiply both by ${grid}.`, `=SUMPRODUCT((${re}=B1)*(${hd}<>B2)*${grid})`],
            explain: `(${hd}<>B2) is 1 for the three categories to keep, and (${re}=B1) is 1 for that region's distributors. Their product with the grid leaves only the wanted cells for SUMPRODUCT to add.`,
            chips: [re, hd, grid, '<>B2'], fnKey: 'SUMPRODUCT('
          });
        },
        () => {
          const rows = r.shuffle(H.PRODUCTS).map((p) => { const o = { sku: p.sku, product: p.name }; MON.forEach((m) => { o[m] = r.int(4, 60) * 5; }); return o; });
          const vals = (x) => MON.map((m) => x[m]).sort((a, b) => a - b);
          const above = (x, lv) => MON.filter((m) => x[m] > lv).length;
          const [a, b] = r.sample(idx(12), 2);
          const lv = vals(rows[a])[2];
          let lv2 = vals(rows[b])[1];
          if (above(rows[b], lv2) === above(rows[a], lv)) lv2 = vals(rows[b])[3];
          const S = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Product', k: 'product' }].concat(MON.map((m) => ({ h: m, k: m, num: true }))), rows);
          const base = Object.assign({}, S.cells, { A1: 'SKU', B1: rows[a].sku, A2: 'More than (cartons)', B2: lv, A3: 'Months' });
          const alt = H.edit(base, { B1: rows[b].sku, B2: lv2 });
          const sk = S.R('SKU'), grid = 'C' + S.first + ':H' + S.last;
          return H.problem({
            title: 'Count the months a SKU sold above a level', file: 'penjualan_bulanan.xlsx', sheet: S, base, alt, answerCell: 'B3',
            statement: `<p>The grid shows cartons sold per SKU (rows) and month (columns C to H). In <b>B3</b>, count how many months the SKU in <b>B1</b> sold more than <b>B2</b> cartons, in one formula.</p>`,
            solution: `=SUMPRODUCT((${sk}=B1)*(${grid}>B2))`,
            alts: [`=COUNTIF(XLOOKUP(B1,${sk},${grid}),">"&B2)`, `=SUM(--(XLOOKUP(B1,${sk},${grid})>B2))`],
            hints: ['Here the answer is a count, so the grid itself becomes a test: (grid>B2) is 1 where a month beat the level.', `Multiply (${sk}=B1) by (${grid}>B2) so only that SKU's row counts.`, `=SUMPRODUCT((${sk}=B1)*(${grid}>B2))`],
            explain: `(${grid}>B2) turns every cell into 1 or 0, and (${sk}=B1) keeps that SKU's row only. SUMPRODUCT adds the 1s, which is the number of months above the level.`,
            chips: [sk, grid, '>B2'], fnKey: 'SUMPRODUCT('
          });
        }
      ][v]();
    }
  });

  P.add({
    key: 'filter-sort-report', difficulty: 'Hard', level: 9, topics: ['FILTER', 'SORT', 'SORTBY', 'LET'], count: 4,
    make(r, v) {
      return [
        () => {
          const rows = untie(H.invoices(r, 14), 'amount', 100);
          let asof = '2026-12-15';
          const early = (d) => idx(14).filter((i) => ser(rows[i].date) + rows[i].terms < ser(d));
          if (early(asof).length < 4) asof = '2026-12-31';
          const E = early(asof);
          r.sample(E, Math.min(3, E.length)).forEach((i) => { rows[i].paid = ''; });
          if (E.every((i) => rows[i].paid === '') && E.length > 3) rows[E[0]].paid = rows[E[0]].date;
          const due = E.filter((i) => rows[i].paid === '');
          const big = due.reduce((m, i) => (m < 0 || rows[i].amount > rows[m].amount ? i : m), -1);
          const S = H.table([{ h: 'Ref', k: 'ref' }, { h: 'Distributor', k: 'distributor' }, { h: 'Amount (Rp)', k: 'amount', num: true }, { h: 'Date', k: 'date', date: true }, { h: 'Terms (days)', k: 'terms', num: true }, { h: 'Paid on', k: 'paid', date: true }], rows);
          const base = Object.assign({}, S.cells, { A1: 'As of', B1: ser(asof) });
          /* The hidden sheet: sixteen days later, and the biggest overdue invoice has been paid. */
          const alt = H.edit(base, { B1: ser(asof) + 16, [S.cell('Paid on', big)]: ser(asof) - 2 });
          const out = blk(S, 'Ref', 'Amount (Rp)'), dt = S.R('Date'), te = S.R('Terms (days)'), pd = S.R('Paid on'), am = S.R('Amount (Rp)');
          const rule = `(${pd}="")*(${dt}+${te}<B1)`;
          return H.problem({
            title: 'Overdue invoices as of a date, largest first', file: 'piutang_q4.xlsx', sheet: S, base, alt, dateCells: ['B1'],
            spill: { cols: 3, rows: rows.length }, spillAnchor: 'H5', ansLabel: 'Overdue',
            statement: `<p>An invoice is overdue when it is unpaid (<b>Paid on</b> is blank) and its due date, <b>Date</b> plus <b>Terms</b>, is before the as-of date in <b>B1</b>. In <b>H5</b>, list the overdue invoices' <b>Ref</b>, <b>Distributor</b> and <b>Amount</b> (${out}), largest amount first, with one formula.</p>`,
            solution: `=SORT(FILTER(${out},${rule}),3,-1)`,
            alts: [`=LET(f,FILTER(${out},${rule}),SORT(f,3,-1))`, `=SORTBY(FILTER(${out},${rule}),FILTER(${am},${rule}),-1)`],
            hints: ['Filter first, then sort what is left: SORT(FILTER(...), column, order).', `The due date is ${dt}+${te}, a number you can compare with B1. Multiply it with the unpaid test (${pd}="").`, `=SORT(FILTER(${out}, ${rule}), 3, -1)`],
            explain: `FILTER keeps the unpaid invoices whose date plus terms falls before B1, and SORT orders those rows by their 3rd column, Amount, largest first.`,
            chips: [out, pd, dt, te, '=""', '<B1', ',3,-1)'], fnKey: 'FILTER('
          });
        },
        () => {
          const rows = book(r, 14, 4);
          const wc = idx(14).filter((i) => rows[i].type === 'Working Capital'), tl = idx(14).filter((i) => rows[i].type === 'Term Loan');
          const [w1, w2, w3, w4, w5] = r.sample(wc, 5);
          /* One facility sits exactly on the 30-day floor, so "at least" matters; one is late but under it. */
          rows[w1].dpd = 30; rows[w2].dpd = 46; rows[w3].dpd = 93; rows[w4].dpd = 12; rows[w5].dpd = 61;
          tl.forEach((i) => { rows[i].dpd = 0; });
          untie(rows, 'outstanding', 100000);
          const S = H.table([{ h: 'Facility', k: 'facility' }, { h: 'Distributor', k: 'distributor' }, { h: 'DPD', k: 'dpd', num: true }, { h: 'Outstanding (Rp)', k: 'outstanding', num: true }, { h: 'Type', k: 'type' }], rows);
          const base = Object.assign({}, S.cells, { A1: 'Type', B1: 'Working Capital', A2: 'DPD at least', B2: 30 });
          /* The hidden sheet: the term loans with a 20-day floor; two have slipped past it and one is late but under it. */
          const alt = H.edit(base, { B1: 'Term Loan', B2: 20, [S.cell('DPD', tl[0])]: 33, [S.cell('DPD', tl[1])]: 15, [S.cell('DPD', tl[2])]: 71 });
          const out = blk(S, 'Facility', 'Outstanding (Rp)'), ty = S.R('Type'), dp = S.R('DPD'), ou = S.R('Outstanding (Rp)');
          const rule = `(${ty}=B1)*(${dp}>=B2)`;
          return H.problem({
            title: 'Facilities of one type past a DPD floor, largest balance first', file: 'buku_pinjaman.xlsx', sheet: S, base, alt,
            spill: { cols: 4, rows: rows.length }, spillAnchor: 'G5', ansLabel: 'Watch list',
            statement: `<p>Citra wants a watch list: the facilities of the type in <b>B1</b> whose <b>DPD</b> is at least the number of days in <b>B2</b> (a facility exactly on B2 is listed), showing <b>Facility</b>, <b>Distributor</b>, <b>DPD</b> and <b>Outstanding</b> (${out}), largest Outstanding first. In <b>G5</b>, type one formula.</p>`,
            solution: `=SORT(FILTER(${out},${rule}),4,-1)`,
            alts: [`=LET(f,FILTER(${out},${rule}),SORT(f,4,-1))`, `=SORTBY(FILTER(${out},${rule}),FILTER(${ou},${rule}),-1)`, `=SORT(FILTER(${out},(${dp}>=B2)*(${ty}=B1)),4,-1)`],
            hints: ['Keep the rows first, then order them: SORT(FILTER(...), column, order).', `FILTER(${out}, (${ty}=B1)*(${dp}>=B2)) keeps the facilities of that type on or past the floor. Outstanding is column 4 of the block.`, `=SORT(FILTER(${out}, ${rule}), 4, -1)`],
            explain: `FILTER keeps the facilities of that type with a DPD of at least B2, and SORT orders them by Outstanding, the 4th column, largest first, so the biggest exposure is on top.`,
            chips: [out, ty, dp, ')*(', '>=B2', ',4,-1)'], fnKey: 'FILTER('
          });
        },
        () => {
          const rows = H.reps(r, 12);
          const low = r.sample(idx(rows.length), 4);
          low.forEach((i) => { rows[i].actual = Math.round(rows[i].target * r.pick([0.6, 0.74, 0.8, 0.88]) / 100000) * 100000; });
          rows.forEach((x, i) => { if (!low.includes(i) && x.actual < x.target) x.actual = x.target + r.int(1, 20) * 1000000; });
          const seen = new Set();
          low.forEach((i) => { while (seen.has(rows[i].target - rows[i].actual)) rows[i].actual -= 100000; seen.add(rows[i].target - rows[i].actual); });
          const up = idx(rows.length).find((i) => !low.includes(i));
          const S = H.table([{ h: 'Name', k: 'name' }, { h: 'Area', k: 'area' }, { h: 'Target (Rp)', k: 'target', num: true }, { h: 'Actual (Rp)', k: 'actual', num: true }], rows);
          const base = Object.assign({}, S.cells);
          /* The hidden sheet: one rep below target catches up, and one above it falls short. */
          const alt = H.edit(base, { [S.cell('Actual (Rp)', low[0])]: rows[low[0]].target + 1000000, [S.cell('Actual (Rp)', up)]: rows[up].target - 7300000 });
          const nm = S.R('Name'), tg = S.R('Target (Rp)'), ac = S.R('Actual (Rp)');
          return H.problem({
            title: 'Reps below target, biggest gap first', file: 'tim_sales.xlsx', sheet: S, base, alt,
            spill: { cols: 2, rows: rows.length }, spillAnchor: 'F5', ansLabel: 'Below target',
            statement: `<p>Pak Anton wants the reps whose <b>Actual</b> is below <b>Target</b>, each with its gap (Target minus Actual), biggest gap first. In <b>F5</b>, return two columns, the <b>Name</b> and the gap, with one formula.</p>`,
            solution: `=LET(g,${tg}-${ac},SORT(FILTER(HSTACK(${nm},g),g>0),2,-1))`,
            alts: [`=SORT(FILTER(HSTACK(${nm},${tg}-${ac}),${ac}<${tg}),2,-1)`, `=SORTBY(FILTER(HSTACK(${nm},${tg}-${ac}),${ac}<${tg}),FILTER(${tg}-${ac},${ac}<${tg}),-1)`],
            hints: ['The gap is not on the sheet. Work it out once with LET and use it twice: to filter and to show.', 'HSTACK(names, gap) builds the two columns; keep the rows where the gap is above 0, then sort by column 2.', `=LET(g, ${tg}-${ac}, SORT(FILTER(HSTACK(${nm}, g), g>0), 2, -1))`],
            explain: `g holds every rep's gap. FILTER keeps the name and gap pairs with a positive gap, and SORT orders them by the gap, largest first.`,
            chips: [nm, tg, ac, 'HSTACK(', ',2,-1)'], fnKey: 'LET('
          });
        },
        () => {
          const rows = untie(H.sales(r, 16), 'net', 100);
          const S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'Invoice', k: 'invoice' }, { h: 'Distributor', k: 'distributor' }, { h: 'Net (Rp)', k: 'net', num: true }], rows);
          const base = Object.assign({}, S.cells, { A1: 'From', B1: ser(rows[3].date), A2: 'To', B2: ser(rows[11].date) });
          const alt = H.edit(base, { B1: ser(rows[1].date), B2: ser(rows[8].date) });
          const out = blk(S, 'Invoice', 'Net (Rp)'), dt = S.R('Date'), ne = S.R('Net (Rp)');
          const rule = `(${dt}>=B1)*(${dt}<=B2)`;
          return H.problem({
            title: 'Orders in a date window, largest net first', file: 'penjualan_q4.xlsx', sheet: S, base, alt, dateCells: ['B1', 'B2'],
            spill: { cols: 3, rows: rows.length }, spillAnchor: 'F5', ansLabel: 'In the window',
            statement: `<p>Tika wants the orders dated from <b>B1</b> to <b>B2</b> (both days included), showing <b>Invoice</b>, <b>Distributor</b> and <b>Net</b> (${out}), largest net first. In <b>F5</b>, type one formula.</p>`,
            solution: `=SORT(FILTER(${out},${rule}),3,-1)`,
            alts: [`=LET(f,FILTER(${out},${rule}),SORT(f,3,-1))`, `=SORTBY(FILTER(${out},${rule}),FILTER(${ne},${rule}),-1)`],
            hints: ['The window is two tests on the Date column, multiplied: on or after B1, and on or before B2.', `FILTER the block ${out} with that rule, then SORT it by column 3.`, `=SORT(FILTER(${out}, ${rule}), 3, -1)`],
            explain: `The two date tests keep the orders inside the window, and SORT orders the kept rows by Net, the 3rd column of the block, largest first.`,
            chips: [out, dt, '>=B1', '<=B2', ',3,-1)'], fnKey: 'FILTER('
          });
        }
      ][v]();
    }
  });

  P.add({
    key: 'first-or-last-match-rule', difficulty: 'Hard', level: 9, topics: ['XLOOKUP', 'XMATCH'], count: 3,
    make(r, v) {
      return [
        () => {
          const rows = H.sales(r, 16);
          const rule = (k) => {
            const after = rows[k].date, R = rows.filter((x) => x.date > after);
            const min = Math.ceil((R[0].net + 1) / 100000) * 100000, hit = R.find((x) => x.net >= min);
            return hit ? { after, min, ans: hit.invoice } : { after, min: R[0].net, ans: R[0].invoice };
          };
          const b = rule(3);
          let a2 = rule(7);
          if (a2.ans === b.ans) a2 = rule(9);
          const S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'Invoice', k: 'invoice' }, { h: 'Distributor', k: 'distributor' }, { h: 'Net (Rp)', k: 'net', num: true }], rows);
          const base = Object.assign({}, S.cells, { A1: 'After', B1: ser(b.after), A2: 'Net at least (Rp)', B2: b.min, A3: 'First invoice' });
          const alt = H.edit(base, { B1: ser(a2.after), B2: a2.min });
          const dt = S.R('Date'), inv = S.R('Invoice'), ne = S.R('Net (Rp)');
          const test = `(${dt}>B1)*(${ne}>=B2)`;
          return H.problem({
            title: 'First order after a date that meets a net minimum', file: 'penjualan_q4.xlsx', sheet: S, base, alt, answerCell: 'B3', dateCells: ['B1'],
            statement: `<p>The orders are in date order. In <b>B3</b>, return the <b>Invoice</b> of the first order dated after <b>B1</b> (not on it) with a <b>Net</b> of at least <b>B2</b>. If there is none, show <code>None</code>.</p>`,
            solution: `=XLOOKUP(1,${test},${inv},"None")`,
            alts: [`=INDEX(${inv},XMATCH(1,${test}))`, `=INDEX(FILTER(${inv},${test}),1)`, `=TAKE(FILTER(${inv},${test}),1)`],
            hints: ['XLOOKUP can search a calculated array: build a list that is 1 where both rules hold, then look up 1.', `(${dt}>B1)*(${ne}>=B2) is that list. XLOOKUP returns the first match from the top.`, `=XLOOKUP(1, ${test}, ${inv}, "None")`],
            explain: `Multiplying the two tests gives a column of 1s and 0s, and XLOOKUP finds the first 1 from the top, which is the earliest order that passes both rules.`,
            chips: [dt, ne, inv, ')*(', ',"None")'], fnKey: 'XLOOKUP('
          });
        },
        () => {
          const rows = H.monthly(r);
          const mg = rows.map((x) => x.sales - x.cost);
          const sorted = mg.slice().sort((a, b) => a - b);
          const first = (fl) => mg.findIndex((m) => m < fl);
          const fl = sorted[3] + 500000;
          let fl2 = fl;
          for (const k of [1, 0, 2, 5, 6, 7]) { const f = sorted[k] + 500000; if (first(f) !== first(fl)) { fl2 = f; break; } }
          const S = H.table([{ h: 'Month', k: 'month' }, { h: 'Sales (Rp)', k: 'sales', num: true }, { h: 'Returns (Rp)', k: 'returns', num: true }, { h: 'Cost (Rp)', k: 'cost', num: true }], rows);
          const base = Object.assign({}, S.cells, { A1: 'Margin floor (Rp)', B1: fl, A2: 'First month below' });
          const alt = H.edit(base, { B1: fl2 });
          const mo = S.R('Month'), sa = S.R('Sales (Rp)'), co = S.R('Cost (Rp)');
          return H.problem({
            title: 'First month the margin fell below a floor', file: 'laba_bulanan.xlsx', sheet: S, base, alt, answerCell: 'B2',
            statement: `<p>Margin is <b>Sales</b> minus <b>Cost</b> (leave Returns out). In <b>B2</b>, return the name of the first month, from January, whose margin is less than the floor in <b>B1</b>. If no month is below it, show <code>None</code>. There is no margin column.</p>`,
            solution: `=XLOOKUP(TRUE,${sa}-${co}<B1,${mo},"None")`,
            alts: [`=INDEX(${mo},XMATCH(TRUE,${sa}-${co}<B1))`, `=XLOOKUP(1,--(${sa}-${co}<B1),${mo},"None")`, `=INDEX(FILTER(${mo},${sa}-${co}<B1),1)`],
            hints: ['The lookup array can be a test: it gives TRUE for each month below the floor.', `${sa}-${co}<B1 is that test. Look up TRUE in it and return ${mo}.`, `=XLOOKUP(TRUE, ${sa}-${co}<B1, ${mo}, "None")`],
            explain: `${sa}-${co}<B1 makes twelve TRUE or FALSE values, and XLOOKUP returns the month of the first TRUE from the top.`,
            chips: [sa, co, mo, 'TRUE,', '<B1', ',"None")'], fnKey: 'XLOOKUP('
          });
        },
        () => {
          const names = r.sample(H.db().distributors, 4).map((d) => d.name);
          const rows = salesFrom(r, [[16, isIn('distributor', names)]]);
          const [who, next] = tally(rows, 'distributor');
          const rule = (name) => {
            const R = rows.filter((x) => x.distributor === name), last = R[R.length - 1];
            if (!R.slice(0, -1).some((x) => x.net > last.net)) last.net = Math.max(100000, Math.min(...R.map((x) => x.net)) - 200000);
            return last.net;
          };
          const lim = rule(who), lim2 = rule(next);
          const S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'Invoice', k: 'invoice' }, { h: 'Distributor', k: 'distributor' }, { h: 'Net (Rp)', k: 'net', num: true }], rows);
          const base = Object.assign({}, S.cells, { A1: 'Distributor', B1: who, A2: 'Net more than (Rp)', B2: lim, A3: 'Last invoice' });
          const alt = H.edit(base, { B1: next, B2: lim2 });
          const di = S.R('Distributor'), inv = S.R('Invoice'), ne = S.R('Net (Rp)');
          const test = `(${di}=B1)*(${ne}>B2)`;
          return H.problem({
            title: 'Last invoice of a distributor above an amount', file: 'penjualan_q4.xlsx', sheet: S, base, alt, answerCell: 'B3',
            statement: `<p>The orders are in date order. In <b>B3</b>, return the <b>Invoice</b> of the most recent order from the distributor in <b>B1</b> with a <b>Net</b> of more than <b>B2</b>. If there is none, show <code>None</code>.</p>`,
            solution: `=XLOOKUP(1,${test},${inv},"None",0,-1)`,
            alts: [`=INDEX(${inv},XMATCH(1,${test},0,-1))`, `=TAKE(FILTER(${inv},${test}),-1)`],
            hints: ['Build a list that is 1 where both rules hold, then search it from the bottom.', 'XLOOKUP\'s sixth argument, search_mode, is -1 for last to first. The fifth, match_mode, stays 0.', `=XLOOKUP(1, ${test}, ${inv}, "None", 0, -1)`],
            explain: `The product of the two tests is 1 on that distributor's orders above B2. Searching from the bottom with search_mode -1 finds the newest of them.`,
            chips: [di, ne, inv, ')*(', ',"None",0,-1)'], fnKey: 'XLOOKUP('
          });
        }
      ][v]();
    }
  });

  P.add({
    key: 'rank-and-run-within-group', difficulty: 'Hard', level: 10, topics: ['SUMPRODUCT', 'COUNTIFS', 'SUMIFS'], count: 3,
    make(r, v) {
      return [
        () => {
          const regs = r.sample(H.REGIONS, 3);
          const rows = untie(salesFrom(r, [[16, isIn('region', regs)]]), 'net', 100);
          const hi = Math.max(...rows.map((x) => x.net)), lo = idx(16).reduce((m, i) => (rows[i].net < rows[m].net ? i : m), 0);
          const S = H.table([{ h: 'Invoice', k: 'invoice' }, { h: 'Distributor', k: 'distributor' }, { h: 'Region', k: 'region' }, { h: 'Net (Rp)', k: 'net', num: true }], rows);
          const base = Object.assign({}, S.cells);
          /* The hidden sheet: the smallest order is corrected to the biggest, and one more is revised. */
          const alt = H.edit(base, { [S.cell('Net (Rp)', lo)]: hi + 1000000, [S.cell('Net (Rp)', (lo + 5) % 16)]: (x) => x + 2300000 });
          const re = S.A('Region'), ne = S.A('Net (Rp)'), c = S.cell('Region', 0), n = S.cell('Net (Rp)', 0);
          return H.problem({
            title: 'Rank each order within its region by net', file: 'penjualan_q4.xlsx', sheet: S, base, alt, fill: true, ansLabel: 'Rank in region',
            statement: `<p>In <b>E5</b>, give each order its rank by <b>Net</b> among the orders of the same <b>Region</b>: 1 for the biggest net in that region, 2 for the next, and so on. Fill it down to <b>E${S.last}</b>.</p>`,
            solution: `=SUMPRODUCT((${re}=${c})*(${ne}>${n}))+1`,
            alts: [`=COUNTIFS(${re},${c},${ne},">"&${n})+1`, `=SUMPRODUCT(--(${re}=${c}),--(${ne}>${n}))+1`],
            hints: ['An order\'s rank is 1 plus the number of orders in the same region with a bigger net.', `Count those with SUMPRODUCT: (${re}=${c}) for the same region, (${ne}>${n}) for a bigger net. Pin the ranges with $ so they stay put when you fill down.`, `=SUMPRODUCT((${re}=${c})*(${ne}>${n}))+1`],
            explain: `For each row, the product is 1 for orders in the same region with a bigger net. SUMPRODUCT counts them and +1 turns the count into a rank; the pinned ranges keep the whole list in view on every row.`,
            chips: [re, ne, c, n, ')*(', ')+1'], fnKey: 'SUMPRODUCT('
          });
        },
        () => {
          /* Orders as keyed, not sorted by date, and never two orders of one distributor on one day, so "to date" is exact. */
          const names = r.sample(H.db().distributors, 3).map((d) => d.name);
          const seenDay = new Set(), rows = [];
          for (const s of r.shuffle(H.db().sales.filter((x) => names.includes(x.distributor)))) {
            const date = H.Q4(s.date), k = s.distributor + '|' + date;
            if (seenDay.has(k)) continue;
            seenDay.add(k); rows.push(Object.assign({}, s, { date }));
            if (rows.length === 16) break;
          }
          const S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'Distributor', k: 'distributor' }, { h: 'Product', k: 'product' }, { h: 'Cartons', k: 'cartons', num: true }], rows);
          const base = Object.assign({}, S.cells);
          /* The hidden sheet: two orders are corrected. */
          const alt = H.edit(base, { [S.cell('Cartons', 2)]: (x) => x + 7, [S.cell('Cartons', 9)]: (x) => x + 11 });
          const f = S.first, dt = S.A('Date'), di = S.A('Distributor'), ca = S.A('Cartons');
          return H.problem({
            title: 'Cartons to date for each distributor, rows not in date order', file: 'penjualan_q4.xlsx', sheet: S, base, alt, fill: true, ansLabel: 'Cartons to date',
            statement: `<p>The orders were keyed as they came in, so the rows are <b>not</b> in date order. In <b>E5</b>, show the cartons to date for the row's <b>Distributor</b>: the <b>Cartons</b> of every order from that distributor dated on or before this row's <b>Date</b>, this row included (a distributor never has two orders on one day). Fill it down to <b>E${S.last}</b>.</p>`,
            solution: `=SUMIFS(${ca},${di},B${f},${dt},"<="&A${f})`,
            alts: [`=SUMPRODUCT((${di}=B${f})*(${dt}<=A${f})*${ca})`, `=SUM(FILTER(${ca},(${di}=B${f})*(${dt}<=A${f})))`, `=SUMIFS(${ca},${dt},"<="&A${f},${di},B${f})`],
            traps: [{ code: 'rows-not-dates', formula: `=SUMIFS($D$${f}:D${f},$B$${f}:B${f},B${f})`, title: 'The rows above are not the earlier orders', plain: `A growing range like $D$${f}:D${f} adds the rows above this one, but these rows are not in date order. Compare each order's Date with this row's date instead.` }],
            hints: ['A range that grows as you fill down adds the rows above, and here the rows above are not always the earlier orders. Test the dates instead.', `SUMIFS can take two rules over the whole pinned columns: the same distributor as B${f}, and a date on or before A${f}, written "<="&A${f}.`, `=SUMIFS(${ca}, ${di}, B${f}, ${dt}, "<="&A${f})`],
            explain: `The pinned ranges look at every order on each row. SUMIFS keeps the orders of this row's distributor dated on or before this row's date and adds their cartons, wherever those orders sit in the list.`,
            chips: [ca, di, dt, `"<="&A${f}`], fnKey: 'SUMIFS('
          });
        },
        () => {
          const rows = book(r, 14, 5);
          const seen = new Set(), key = (x) => Math.round(x.outstanding / x.limit * 1e6);
          rows.forEach((x) => { while (seen.has(key(x))) x.outstanding -= 100000; seen.add(key(x)); });
          const [a, b] = r.sample(idx(14), 2);
          const S = H.table([{ h: 'Facility', k: 'facility' }, { h: 'Distributor', k: 'distributor' }, { h: 'Type', k: 'type' }, { h: 'Limit (Rp)', k: 'limit', num: true }, { h: 'Outstanding (Rp)', k: 'outstanding', num: true }], rows);
          const base = Object.assign({}, S.cells);
          /* The hidden sheet: one line is drawn almost to its limit, another is mostly repaid. */
          const alt = H.edit(base, { [S.cell('Outstanding (Rp)', a)]: rows[a].limit - 300000, [S.cell('Outstanding (Rp)', b)]: Math.round(rows[b].limit * 0.07 / 100000) * 100000 + 100000 });
          const ty = S.A('Type'), li = S.A('Limit (Rp)'), ou = S.A('Outstanding (Rp)'), c = S.cell('Type', 0), l = S.cell('Limit (Rp)', 0), o = S.cell('Outstanding (Rp)', 0);
          return H.problem({
            title: "Rank each facility's utilisation within its type", file: 'buku_pinjaman.xlsx', sheet: S, base, alt, fill: true, ansLabel: 'Rank in type',
            statement: `<p>Utilisation is <b>Outstanding</b> divided by <b>Limit</b>. In <b>F5</b>, rank each facility by utilisation among the facilities of the same <b>Type</b>: 1 for the most used. Fill it down to <b>F${S.last}</b>. There is no utilisation column.</p>`,
            solution: `=SUMPRODUCT((${ty}=${c})*(${ou}/${li}>${o}/${l}))+1`,
            alts: [`=SUMPRODUCT(--(${ty}=${c}),--(${ou}/${li}>${o}/${l}))+1`, `=SUMPRODUCT((${c}=${ty})*(${o}/${l}<${ou}/${li}))+1`],
            hints: ['COUNTIFS cannot compare a ratio it has to work out. SUMPRODUCT can: divide the two pinned columns inside it.', `Count the facilities of the same type with ${ou}/${li} greater than this row's ${o}/${l}, then add 1.`, `=SUMPRODUCT((${ty}=${c})*(${ou}/${li}>${o}/${l}))+1`],
            explain: `${ou}/${li} works out every facility's utilisation in memory. The product keeps those of the same type that are more used than this row, and their count plus 1 is the rank.`,
            chips: [ty, ou, li, c, `${o}/${l}`, ')+1'], fnKey: 'SUMPRODUCT('
          });
        }
      ][v]();
    }
  });

  P.add({
    key: 'let-filter-statistics', difficulty: 'Hard', level: 9, topics: ['LET', 'FILTER', 'UNIQUE', 'LARGE'], count: 4,
    make(r, v) {
      return [
        () => {
          const ds = r.sample(H.db().distributors, 5).map((d) => d.name);
          const rows = untie(salesFrom(r, [[6, (s) => s.distributor === ds[0]], [5, (s) => s.distributor === ds[1]], [5, isIn('distributor', ds.slice(2))]]), 'net', 100);
          const S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'Invoice', k: 'invoice' }, { h: 'Distributor', k: 'distributor' }, { h: 'Net (Rp)', k: 'net', num: true }], rows);
          const base = Object.assign({}, S.cells, { A1: 'Distributor', B1: ds[0], A2: 'Average of top 3' });
          const alt = H.edit(base, { B1: ds[1] });
          const di = S.R('Distributor'), ne = S.R('Net (Rp)'), f = `FILTER(${ne},${di}=B1)`;
          return H.problem({
            title: "Average of one distributor's three biggest orders", file: 'penjualan_q4.xlsx', sheet: S, base, alt, answerCell: 'B2',
            statement: `<p>Citra wants the average <b>Net</b> of the three biggest orders from the distributor in <b>B1</b> (it always has at least three). Put it in <b>B2</b> with one formula.</p>`,
            solution: `=AVERAGE(TAKE(SORT(${f},,-1),3))`,
            alts: [`=LET(f,${f},(LARGE(f,1)+LARGE(f,2)+LARGE(f,3))/3)`, `=AVERAGE(TAKE(SORTBY(${f},${f},-1),3))`],
            hints: ['Keep that distributor\'s net values, put the biggest on top, keep three, then average.', `SORT(${f}, , -1) sorts the values largest first; the empty second argument means the first column.`, `=AVERAGE(TAKE(SORT(${f}, , -1), 3))`],
            explain: `FILTER keeps the distributor's net values, SORT puts the largest first, TAKE keeps three and AVERAGE averages them. With LET you can name the filtered list once and use LARGE three times instead.`,
            chips: [ne, di, ',,-1)', ',3))'], fnKey: 'TAKE('
          });
        },
        () => {
          const rows = untie(H.sales(r, 16), 'net', 100);
          const hi = idx(16).reduce((m, i) => (rows[i].net > rows[m].net ? i : m), 0), lo = idx(16).reduce((m, i) => (rows[i].net < rows[m].net ? i : m), 0);
          const S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'Distributor', k: 'distributor' }, { h: 'Region', k: 'region' }, { h: 'Net (Rp)', k: 'net', num: true }], rows);
          const base = Object.assign({}, S.cells, { A1: 'Share above median' });
          const alt = H.edit(base, { [S.cell('Net (Rp)', hi)]: Math.round(rows[hi].net * 1.6 / 100) * 100, [S.cell('Net (Rp)', lo)]: (x) => x + 500000 });
          const ne = S.R('Net (Rp)');
          return H.problem({
            title: 'Share of net from orders above the median', file: 'penjualan_q4.xlsx', sheet: S, base, alt, answerCell: 'B1', precision: 4,
            statement: `<p>Tika wants to know how much of the total <b>Net</b> comes from the bigger orders: the net of the orders above the median order, divided by the total net, as a decimal. Put it in <b>B1</b>; LET lets you name the column once.</p>`,
            solution: `=LET(n,${ne},SUM(FILTER(n,n>MEDIAN(n)))/SUM(n))`,
            alts: [`=SUMPRODUCT((${ne}>MEDIAN(${ne}))*${ne})/SUM(${ne})`, `=SUMIF(${ne},">"&MEDIAN(${ne}))/SUM(${ne})`],
            hints: ['MEDIAN gives the middle value. "Above the median" is n>MEDIAN(n).', 'Name the column n with LET, then FILTER the values above the median, SUM them, and divide by SUM(n).', `=LET(n, ${ne}, SUM(FILTER(n, n>MEDIAN(n)))/SUM(n))`],
            explain: `LET names the Net column n. FILTER keeps the values above the median, and their sum over the whole sum is the share the bigger half of the orders brings in.`,
            chips: [ne, 'MEDIAN(n)', 'SUM(n)'], fnKey: 'LET('
          });
        },
        () => {
          const rows = H.sales(r, 20);
          const stat = (c, k) => { const f = rows.filter((x) => x.category === c && x.date >= rows[k].date); return { c, k, n: f.length, u: new Set(f.map((x) => x.distributor)).size }; };
          const all = [];
          for (const c of H.CATEGORIES) for (let k = 3; k <= 12; k++) all.push(stat(c, k));
          const good = all.filter((x) => x.u >= 2).sort((a, b) => (b.n - b.u) - (a.n - a.u) || b.u - a.u || a.k - b.k);
          const b0 = good[0] || all.find((x) => x.u >= 1);
          const a0 = all.find((x) => x.u >= 1 && x.u !== b0.u && x.c !== b0.c) || all.find((x) => x.u >= 1 && x.u !== b0.u);
          const S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'Distributor', k: 'distributor' }, { h: 'Region', k: 'region' }, { h: 'Category', k: 'category' }, { h: 'Net (Rp)', k: 'net', num: true }], rows);
          const base = Object.assign({}, S.cells, { A1: 'Category', B1: b0.c, A2: 'On or after', B2: ser(rows[b0.k].date), A3: 'Distributors' });
          const alt = H.edit(base, { B1: a0.c, B2: ser(rows[a0.k].date) });
          const di = S.R('Distributor'), ca = S.R('Category'), dt = S.R('Date');
          const keep = `FILTER(${di},(${ca}=B1)*(${dt}>=B2))`;
          return H.problem({
            title: 'Distinct distributors buying a category since a date', file: 'penjualan_q4.xlsx', sheet: S, base, alt, answerCell: 'B3', dateCells: ['B2'],
            statement: `<p>Pak Anton wants to know how many different distributors bought the category in <b>B1</b> on or after the date in <b>B2</b>. A distributor with several such orders counts once. Put the count in <b>B3</b>.</p>`,
            solution: `=COUNTA(UNIQUE(${keep}))`,
            alts: [`=ROWS(UNIQUE(${keep}))`, `=LET(d,${keep},COUNTA(UNIQUE(d)))`],
            hints: ['Three steps, inside out: keep the matching rows, remove repeated distributors, count.', `FILTER(${di}, (${ca}=B1)*(${dt}>=B2)) gives the distributors with repeats. UNIQUE removes them.`, `=COUNTA(UNIQUE(${keep}))`],
            explain: `FILTER returns the distributor of every matching order, UNIQUE keeps each name once, and COUNTA counts the names.`,
            chips: [di, ca, dt, ')*(', '>=B2'], fnKey: 'UNIQUE('
          });
        },
        () => {
          const regs = r.sample(H.REGIONS, 4);
          const rows = untie(salesFrom(r, [[6, (s) => s.region === regs[0]], [5, (s) => s.region === regs[1]], [5, isIn('region', regs.slice(2))]]), 'net', 100);
          const S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'Distributor', k: 'distributor' }, { h: 'Region', k: 'region' }, { h: 'Net (Rp)', k: 'net', num: true }], rows);
          const base = Object.assign({}, S.cells, { A1: 'Region', B1: regs[0], A2: 'Middle spread (Rp)' });
          const alt = H.edit(base, { B1: regs[1] });
          const re = S.R('Region'), ne = S.R('Net (Rp)'), f = `FILTER(${ne},${re}=B1)`;
          return H.problem({
            title: 'Second largest minus second smallest order in a region', file: 'penjualan_q4.xlsx', sheet: S, base, alt, answerCell: 'B2',
            statement: `<p>To ignore the one extreme order at each end, Dimas measures a region's spread as its second largest <b>Net</b> minus its second smallest. In <b>B2</b>, work it out for the region in <b>B1</b>; LET lets you filter the region once.</p>`,
            solution: `=LET(f,${f},LARGE(f,2)-SMALL(f,2))`,
            alts: [`=LARGE(${f},2)-SMALL(${f},2)`, `=LET(s,SORT(${f}),INDEX(s,ROWS(s)-1)-INDEX(s,2))`],
            hints: ['LARGE(list, 2) is the second largest value and SMALL(list, 2) the second smallest.', `The list is ${f}. Name it with LET so FILTER runs once.`, `=LET(f, ${f}, LARGE(f, 2)-SMALL(f, 2))`],
            explain: `f holds the region's net values. LARGE(f,2) and SMALL(f,2) skip the single biggest and smallest order, so their difference is the spread of the middle orders.`,
            chips: [ne, re, 'LARGE(f,2)', 'SMALL(f,2)'], fnKey: 'LET('
          });
        }
      ][v]();
    }
  });
});
