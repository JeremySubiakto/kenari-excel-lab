/*
 * Practice, Finance (25 Medium, 15 Hard): loans and money on Bank Kenari and Rasa Kenari data. PMT, IPMT, PPMT, FV,
 * PV, NPER, RATE, NPV (plus year 0), IRR, SLN, DDB, simple and compound interest, growth, CAGR, break-even, margin
 * and markup.
 *
 * Rates sit in the sheet as yearly percents (9.5 means 9.5%), the way the bank's loan book stores them, because the
 * practice grid shows numbers with two decimals and would print 9.25% as 0.09. Every formula divides by 100.
 * Money answers are positive: Excel counts a payment as money going out (negative), so each statement says which
 * sign it wants and the solution carries the minus sign. The engine's PMT, IPMT and DDB do not take whole ranges
 * (Excel 365 would spill them), so every problem uses them one row at a time.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) factory(require('./core.js'));
  else factory(root.KicauPractice);
})(typeof globalThis !== 'undefined' ? globalThis : this, function (P) {
  'use strict';
  const H = P.H;
  const M = 1000000;
  const DIST = () => H.db().distributors;
  /* A parameter block: Item and Value from row 5, for problems with more inputs than rows 1 to 3 can hold. */
  const block = (items) => H.table([{ h: 'Item', k: 'item' }, { h: 'Value', k: 'value', num: true }], items.map(([item, value]) => ({ item, value })));
  /* The internal rate of return of yearly flows (first one negative), for setting a hurdle on either side of it. */
  function irr(cf) {
    const npv = (x) => cf.reduce((a, c, i) => a + c / Math.pow(1 + x, i), 0);
    let lo = -0.9, hi = 2;
    for (let k = 0; k < 200; k++) { const mid = (lo + hi) / 2; if (npv(mid) > 0) lo = mid; else hi = mid; }
    return (lo + hi) / 2;
  }
  /* A date some months after a first due date, as ISO text, without Date. */
  const addMonths = (y, m, d, k) => { const t = (m - 1) + k; return (y + Math.floor(t / 12)) + '-' + String(t % 12 + 1).padStart(2, '0') + '-' + String(d).padStart(2, '0'); };
  const RATE_NOTE = 'Rates are yearly percents: 9.5 means 9.5%, so divide by 100.';

  /* ===================================== MEDIUM ===================================== */

  P.add({
    key: 'installment-per-loan', difficulty: 'Medium', level: 4, topics: ['PMT'], count: 3,
    make(r, v) {
      /* Three loan books, three payment rhythms: years to months, months as given, years to quarters. */
      const V = [
        { title: 'Monthly installment per term loan, tenor in years', file: 'pinjaman_berjangka.xlsx', id: 'TL-', tenorH: 'Tenor (years)', tenors: [2, 3, 4, 5], lo: 4, hi: 24, step: 25, rates: [8.5, 9, 9.25, 9.75, 10.5, 11, 12],
          rate: 'D5/100/12', n: 'E5*12', label: 'Installment (Rp)', lead: 'Bank Kenari gives distributors term loans, paid back in equal <b>monthly</b> installments.', tenor: 'Column E is the tenor in <b>years</b>.',
          alts: ['=-PMT(D5/100/12,E5*12,C5)', '=PMT(D5/1200,12*E5,-C5)', '=ABS(PMT(D5/1200,E5*12,C5))'], conv: 'the monthly rate is D5/100/12 and the number of payments is E5*12' },
        { title: 'Monthly installment per truck loan, tenor in months', file: 'kredit_truk.xlsx', id: 'TR-', tenorH: 'Tenor (months)', tenors: [24, 30, 36, 48, 60], lo: 18, hi: 48, step: 10, rates: [7.5, 8, 8.25, 8.75, 9.5, 10],
          rate: 'D5/100/12', n: 'E5', label: 'Installment (Rp)', lead: 'Distributors buy delivery trucks with a Bank Kenari vehicle loan, paid back in equal <b>monthly</b> installments.', tenor: 'Column E is already the number of <b>months</b>.',
          alts: ['=-PMT(D5/100/12,E5,C5)', '=PMT(D5/1200,E5,-C5)', '=ABS(PMT(D5/1200,E5,C5))'], conv: 'the monthly rate is D5/100/12 and the number of payments is E5' },
        { title: 'Quarterly installment per warehouse investment loan', file: 'pinjaman_investasi.xlsx', id: 'IL-', tenorH: 'Tenor (years)', tenors: [3, 4, 5, 6, 7], lo: 8, hi: 30, step: 50, rates: [9, 9.5, 10, 10.25, 10.75, 11.5],
          rate: 'D5/100/4', n: 'E5*4', label: 'Quarterly installment (Rp)', lead: 'Bank Kenari funds new warehouses with investment loans, paid back in equal installments every <b>quarter</b>.', tenor: 'Column E is the tenor in <b>years</b>, with four quarters in a year.',
          alts: ['=-PMT(D5/100/4,E5*4,C5)', '=PMT(D5/400,4*E5,-C5)', '=ABS(PMT(D5/400,E5*4,C5))'], conv: 'the quarterly rate is D5/100/4 and the number of payments is E5*4' }
      ][v];
      const rows = r.sample(DIST(), r.int(10, 12)).map((d, i) => ({ id: V.id + (301 + i), dist: d.name, amt: r.int(V.lo, V.hi) * V.step * M, rate: r.pick(V.rates), tenor: r.pick(V.tenors) }));
      const S = H.table([{ h: 'Facility', k: 'id' }, { h: 'Distributor', k: 'dist' }, { h: 'Amount (Rp)', k: 'amt', num: true }, { h: 'Rate (% a year)', k: 'rate', num: true }, { h: V.tenorH, k: 'tenor', num: true }], rows);
      /* The hidden sheet: two amounts revised, one rate repriced, one tenor extended. */
      const alt = H.edit(S.cells, { C5: (x) => x + V.step * M, D7: (x) => x + 0.5, E9: (x) => V.tenors[(V.tenors.indexOf(x) + 1) % V.tenors.length], ['C' + S.last]: (x) => x - V.step * M });
      return H.problem({
        title: V.title, file: V.file, sheet: S, alt, fill: true, ansLabel: V.label,
        statement: '<p>' + V.lead + ' In <b>F5</b>, work out each loan\'s installment as a <b>positive</b> number of rupiah, then fill down to <b>F' + S.last + '</b>. ' + RATE_NOTE + ' ' + V.tenor + ' Leave the result unrounded.</p>',
        solution: '=PMT(' + V.rate + ',' + V.n + ',-C5)', alts: V.alts,
        traps: [
          { formula: '=PMT(' + V.rate + ',' + V.n + ',C5)', code: 'sign', title: 'The installment came out negative', plain: 'PMT counts a payment as money going out, so it returns a negative number. Put a minus sign in front of the amount: -C5.' },
          { formula: '=PMT(D5/100,' + V.n + ',-C5)', code: 'yearly-rate', title: 'The rate is still yearly', plain: 'PMT needs the rate for one payment period. Divide the yearly rate by the number of payments in a year too.' },
          { formula: '=PMT(' + V.rate.replace('/100', '') + ',' + V.n + ',-C5)', code: 'percent', title: 'The rate is a percent number', plain: 'Column D holds 9.5 for 9.5%. Divide it by 100 as well.' }
        ],
        hints: ['PMT(rate, number of payments, amount) gives an equal installment. The rate and the count must use the same period as the payments.', 'Here ' + V.conv + '. PMT returns money going out as a negative number, so use -C5 for the amount.', 'The shape: =PMT(' + V.rate + ', ' + V.n + ', -C5)'],
        explain: 'PMT spreads the amount and its interest into equal payments. ' + V.rate + ' turns the yearly percent into the rate for one payment period, ' + V.n + ' counts the payments, and -C5 makes the installment come out positive.',
        chips: [V.rate, V.n, '-C5'], fnKey: 'PMT('
      });
    }
  });

  P.add({
    key: 'deposit-interest', difficulty: 'Medium', level: 5, topics: ['FV', 'Interest'], count: 3,
    make(r, v) {
      /* Time deposits that distributors keep at Bank Kenari; the variants change how the interest grows. */
      const V = [
        { title: 'Simple interest earned on each time deposit', label: 'Interest (Rp)', sol: '=C5*D5/100*E5', alts: ['=C5*(D5/100)*E5', '=E5*C5*D5/100'],
          ask: 'the <b>simple interest</b> each deposit earns over its whole term: the deposit times the yearly rate times the years, with no interest on interest',
          hints: ['Simple interest is amount x rate x years.', 'The rate in D5 is a percent number, so use D5/100.', 'The shape: =C5*D5/100*E5'],
          explain: 'Simple interest pays the same amount every year, so the total is the deposit times the yearly rate times the years. D5/100 turns 4.5 into 4.5%.',
          traps: [{ formula: '=C5*D5*E5', code: 'percent', title: 'The rate is a percent number', plain: 'D5 holds 4.5 for 4.5%. Divide it by 100.' }], chips: ['D5/100', '*E5'], fnKey: 'C5*D5/100' },
        { title: 'Time deposit value with interest compounded yearly', label: 'Value at maturity (Rp)', sol: '=C5*(1+D5/100)^E5', alts: ['=FV(D5/100,E5,0,-C5)', '=-FV(D5/100,E5,0,C5)', '=C5*POWER(1+D5/100,E5)'],
          ask: 'what each deposit is worth at maturity when interest is <b>compounded once a year</b> (each year\'s interest earns interest the next year)',
          hints: ['Each year the balance grows by (1 + rate). After E5 years it has grown by (1 + rate) to the power E5.', 'Write the rate as D5/100. FV(rate, years, 0, -amount) gives the same result.', 'The shape: =C5*(1+D5/100)^E5'],
          explain: 'Compounding multiplies the balance by (1 + rate) every year, so after E5 years the deposit is C5*(1+D5/100)^E5. FV(D5/100,E5,0,-C5) works it out the same way.',
          traps: [{ formula: '=C5*D5/100*E5', code: 'simple', title: 'That is the simple interest', plain: 'This is only the interest, and without compounding. The task wants the whole value at maturity, with interest earning interest.' }, { formula: '=C5*(1+D5/100*E5)', code: 'simple-value', title: 'No compounding yet', plain: 'This grows the deposit by simple interest. Raise (1 + rate) to the power of the years instead.' }],
          chips: ['C5*', '(1+D5/100)', '^E5'], fnKey: 'FV(' },
        { title: 'Interest earned with monthly compounding', label: 'Interest (Rp)', sol: '=FV(D5/100/12,E5*12,0,-C5)-C5', alts: ['=C5*(1+D5/1200)^(E5*12)-C5', '=C5*((1+D5/100/12)^(12*E5)-1)', '=-FV(D5/1200,E5*12,0,C5)-C5'],
          ask: 'the <b>interest</b> each deposit earns over its term when interest is <b>compounded every month</b> at a twelfth of the yearly rate. Give only the interest, not the final balance',
          hints: ['Monthly compounding uses the monthly rate D5/100/12 and E5*12 months.', 'FV(rate, periods, 0, -C5) gives the balance at the end; take away the deposit to leave the interest.', 'The shape: =FV(D5/100/12, E5*12, 0, -C5) - C5'],
          explain: 'FV grows the deposit month by month at D5/100/12 for E5*12 months. Taking away C5 leaves only the interest earned.',
          traps: [{ formula: '=FV(D5/100/12,E5*12,0,-C5)', code: 'balance', title: 'That is the whole balance', plain: 'This is the deposit plus the interest. Take away C5 to leave the interest only.' }, { formula: '=C5*(1+D5/100)^E5-C5', code: 'yearly', title: 'That compounds once a year', plain: 'The task compounds every month: use D5/100/12 as the rate and E5*12 periods.' }],
          chips: ['D5/100/12', 'E5*12', '-C5'], fnKey: 'FV(' }
      ][v];
      const rows = r.sample(DIST(), r.int(10, 12)).map((d, i) => ({ id: 'DEP-' + (5101 + i * 3), dist: d.name, amt: r.int(2, 30) * 25 * M, rate: r.pick([3, 3.25, 3.5, 4, 4.25, 4.5, 5, 5.5]), yrs: r.int(1, 5) }));
      const S = H.table([{ h: 'Deposit', k: 'id' }, { h: 'Distributor', k: 'dist' }, { h: 'Amount (Rp)', k: 'amt', num: true }, { h: 'Rate (% a year)', k: 'rate', num: true }, { h: 'Years', k: 'yrs', num: true }], rows);
      const alt = H.edit(S.cells, { C6: (x) => x + 50 * M, D8: (x) => x + 0.75, E10: (x) => (x >= 5 ? 2 : x + 1), ['C' + S.last]: (x) => x + 25 * M });
      return H.problem({
        title: V.title, file: 'deposito_distributor.xlsx', sheet: S, alt, fill: true, ansLabel: V.label,
        statement: '<p>Distributors keep time deposits at Bank Kenari. In <b>F5</b>, work out ' + V.ask + '. Fill it down to <b>F' + S.last + '</b>. ' + RATE_NOTE + ' Leave the result unrounded.</p>',
        solution: V.sol, alts: V.alts, traps: V.traps, hints: V.hints, explain: V.explain, chips: V.chips, fnKey: V.fnKey
      });
    }
  });

  P.add({
    key: 'month-on-month-growth', difficulty: 'Medium', level: 4, topics: ['Growth'], count: 2,
    make(r, v) {
      /* Real figures from the shared world, where three distributors are fading in the last months. */
      const sales = H.db().sales;
      const sum = (id, m, k) => sales.filter((x) => x.distributor_id === id && x.date.startsWith(m)).reduce((a, x) => a + x[k], 0);
      const ds = r.sample(DIST(), 12);
      if (v === 0) {
        const S = H.table([{ h: 'Distributor', k: 'name' }, { h: 'Region', k: 'region' }, { h: 'Sep net (Rp)', v: (d) => sum(d.id, '2026-09', 'net'), num: true }, { h: 'Oct net (Rp)', v: (d) => sum(d.id, '2026-10', 'net'), num: true }], ds);
        const alt = H.edit(S.cells, { D5: (x) => x + 12500000, D8: (x) => x - 8000000, C10: (x) => x + 6400000, ['D' + S.last]: (x) => x + 3300000 });
        return H.problem({
          title: 'Month-on-month growth in net sales per distributor', file: 'penjualan_bulanan.xlsx', sheet: S, alt, fill: true, ansLabel: 'Growth', precision: 4,
          statement: '<p>Pak Anton wants to see which distributors are slowing down. In <b>E5</b>, work out each distributor\'s growth in net sales from September to October as a decimal (0.08 is 8% up, -0.25 is 25% down). Fill it down to <b>E' + S.last + '</b>.</p>',
          solution: '=D5/C5-1', alts: ['=(D5-C5)/C5', '=D5/C5-100%'],
          traps: [{ formula: '=C5/D5-1', code: 'upside-down', title: 'The months are the wrong way round', plain: 'Growth compares the new month with the old one: October divided by September, minus 1.' }, { formula: '=D5-C5', code: 'rupiah', title: 'That is the change in rupiah', plain: 'Divide the change by September\'s figure to make it a growth rate.' }],
          hints: ['Growth is (new - old) / old.', 'October is in D5 and September in C5. D5/C5 - 1 gives the same thing.', 'The shape: =D5/C5-1'],
          explain: 'D5/C5 is October as a multiple of September; minus 1 leaves the growth. A negative result means sales fell.',
          chips: ['C5', '-1'], fnKey: 'D5/C5'
        });
      }
      const S = H.table([{ h: 'Distributor', k: 'name' }, { h: 'Aug cartons', v: (d) => sum(d.id, '2026-08', 'cartons'), num: true }, { h: 'Sep cartons', v: (d) => sum(d.id, '2026-09', 'cartons'), num: true }, { h: 'Oct cartons', v: (d) => sum(d.id, '2026-10', 'cartons'), num: true }], ds);
      const alt = H.edit(S.cells, { D5: (x) => x + 40, B7: (x) => x + 25, D9: (x) => Math.max(5, x - 30), ['B' + S.last]: (x) => x + 18 });
      return H.problem({
        title: 'Average monthly growth in cartons, August to October', file: 'karton_bulanan.xlsx', sheet: S, alt, fill: true, ansLabel: 'Avg monthly growth', precision: 4,
        statement: '<p>Tika wants one growth figure per distributor for the last two months. In <b>E5</b>, work out the <b>average monthly growth</b> in cartons from August to October: the steady monthly rate that would turn August\'s cartons into October\'s in two months (compound growth, as a decimal). Fill it down to <b>E' + S.last + '</b>.</p>',
        solution: '=(D5/B5)^(1/2)-1', alts: ['=SQRT(D5/B5)-1', '=RRI(2,B5,D5)', '=(D5/B5)^0.5-1'],
        traps: [{ formula: '=D5/C5-1', code: 'one-month', title: 'That is only September to October', plain: 'The task covers two months, from August (column B) to October (column D).' }, { formula: '=(D5/B5-1)/2', code: 'simple-average', title: 'That splits the growth evenly', plain: 'Growth compounds: take the square root of D5/B5, then subtract 1.' }],
        hints: ['Two months of growth at rate g turn August into October: B5*(1+g)^2 = D5.', 'Solve for g: (D5/B5)^(1/2) - 1. RRI(2, B5, D5) does the same.', 'The shape: =(D5/B5)^(1/2)-1'],
        explain: 'October over August is two months of growth multiplied together, so its square root is one month\'s growth factor. Subtracting 1 leaves the rate. September does not change the answer.',
        chips: ['D5/B5', '-1', 'RRI('], fnKey: '^(1/2)'
      });
    }
  });

  P.add({
    key: 'margin-and-markup', difficulty: 'Medium', level: 4, topics: ['Margin', 'Markup', 'ROUNDUP'], count: 3,
    make(r, v) {
      const prods = r.shuffle(H.PRODUCTS).map((p) => ({ sku: p.sku, name: p.name, cat: p.category, price: p.carton_price, cost: Math.round(p.carton_price * r.pick([0.6, 0.64, 0.66, 0.7, 0.72, 0.75, 0.78, 0.81]) / 100) * 100 }));
      if (v < 2) {
        const S = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Product', k: 'name' }, { h: 'Cost per carton (Rp)', k: 'cost', num: true }, { h: 'Price per carton (Rp)', k: 'price', num: true }], prods);
        const alt = H.edit(S.cells, { C5: (x) => x + 4000, C8: (x) => x - 6000, D11: (x) => x + 8000, ['C' + S.last]: (x) => x + 5000 });
        const margin = { formula: '=(D5-C5)/D5', code: 'margin', title: 'That is the margin', plain: 'Margin divides the profit by the price. Markup divides it by the cost.' };
        const markup = { formula: '=(D5-C5)/C5', code: 'markup', title: 'That is the markup', plain: 'Markup divides the profit by the cost. Margin divides it by the price.' };
        return H.problem(v === 0 ? {
          title: 'Gross margin per product as a share of the price', file: 'harga_produk.xlsx', sheet: S, alt, fill: true, ansLabel: 'Margin', precision: 4,
          statement: '<p>In <b>E5</b>, work out each product\'s <b>gross margin</b>: the profit on a carton (price minus cost) as a share of the <b>price</b>, as a decimal (0.3 is 30%). Fill it down to <b>E' + S.last + '</b>.</p>',
          solution: '=(D5-C5)/D5', alts: ['=1-C5/D5', '=(D5-C5)/D5*1'], traps: [markup],
          hints: ['Profit on a carton is price minus cost: D5 - C5.', 'Margin divides that profit by the price in D5. Mind the brackets.', 'The shape: =(D5-C5)/D5'],
          explain: 'Margin is the share of every rupiah of price that is profit, so the profit D5-C5 is divided by the price D5. The brackets make the subtraction happen first.',
          chips: ['/D5'], fnKey: '(D5-C5)'
        } : {
          title: 'Markup per product as a share of the cost', file: 'harga_produk.xlsx', sheet: S, alt, fill: true, ansLabel: 'Markup', precision: 4,
          statement: '<p>In <b>E5</b>, work out each product\'s <b>markup</b>: the profit on a carton (price minus cost) as a share of the <b>cost</b>, as a decimal (0.4 is 40%). Fill it down to <b>E' + S.last + '</b>.</p>',
          solution: '=(D5-C5)/C5', alts: ['=D5/C5-1'], traps: [margin],
          hints: ['Profit on a carton is price minus cost: D5 - C5.', 'Markup divides that profit by the cost in C5.', 'The shape: =(D5-C5)/C5'],
          explain: 'Markup says how much is added on top of the cost, so the profit D5-C5 is divided by the cost C5. It is always larger than the margin on the same product.',
          chips: ['/C5'], fnKey: '(D5-C5)'
        });
      }
      const S = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Product', k: 'name' }, { h: 'Category', k: 'cat' }, { h: 'Cost per carton (Rp)', k: 'cost', num: true }], prods);
      const tm = r.pick([25, 28, 30, 32, 35]);
      const base = Object.assign({}, S.cells, { A1: 'Target margin (%)', B1: tm });
      const alt = H.edit(base, { B1: tm + r.pick([-4, 3, 5]), D6: (x) => x + 7000, D10: (x) => x - 3000 });
      return H.problem({
        title: 'Carton price for a target margin, rounded up to Rp1,000', file: 'harga_produk.xlsx', sheet: S, base, alt, fill: true, ansLabel: 'Price (Rp)',
        statement: '<p>Rasa Kenari sets carton prices so the <b>margin</b> (profit as a share of the price) hits the target in <b>B1</b>, a percent number (30 means 30%). In <b>E5</b>, work out the price for each product, <b>rounded up</b> to a multiple of 1,000 rupiah (a price already on a round thousand stays as it is). Fill it down to <b>E' + S.last + '</b>.</p>',
        solution: '=ROUNDUP(D5/(1-$B$1/100),-3)', alts: ['=CEILING(D5/(1-$B$1/100),1000)', '=ROUNDUP(D5/(100-$B$1)*100,-3)'],
        traps: [{ formula: '=ROUNDUP(D5*(1+$B$1/100),-3)', code: 'markup', title: 'That adds a markup, not a margin', plain: 'Cost times (1 + 30%) gives a 30% markup, and the margin comes out lower. The margin is a share of the price: price = cost / (1 - margin).' }, { formula: '=ROUNDUP(D5/(1-B1/100),-3)', code: 'no-pin', title: 'B1 slid down when you filled', plain: 'Pin the target margin with $B$1 so every row reads the same cell.' }],
        hints: ['If the margin is m, the cost is the rest of the price: cost = price x (1 - m). So price = cost / (1 - m).', 'm is $B$1/100. ROUNDUP(x, -3) rounds up to the next 1,000; pin B1 so it does not slide.', 'The shape: =ROUNDUP(D5/(1-$B$1/100), -3)'],
        explain: 'With a margin m, cost is the share 1-m of the price, so dividing the cost by 1-m gives the price. ROUNDUP with -3 digits lifts it to the next thousand, and $B$1 keeps every row on the same target.',
        chips: ['D5/(1-$B$1/100)', ',-3)', '$B$1'], fnKey: 'ROUNDUP('
      });
    }
  });

  P.add({
    key: 'break-even-cartons', difficulty: 'Medium', level: 5, topics: ['Break-even', 'ROUNDUP'], count: 2,
    make(r, v) {
      const prods = r.sample(H.PRODUCTS, 10).map((p) => ({ name: p.name, fixed: r.pick([45, 60, 75, 90, 120, 150, 180, 240]) * M, price: p.carton_price, varc: Math.round(p.carton_price * r.pick([0.55, 0.6, 0.64, 0.68, 0.72, 0.76]) / 1000) * 1000 }));
      const S = H.table([{ h: 'Product', k: 'name' }, { h: 'Fixed cost a month (Rp)', k: 'fixed', num: true }, { h: 'Price per carton (Rp)', k: 'price', num: true }, { h: 'Variable cost per carton (Rp)', k: 'varc', num: true }], prods);
      if (v === 0) {
        const alt = H.edit(S.cells, { B5: (x) => x + 15 * M, D7: (x) => x + 6000, C9: (x) => x - 4000, ['B' + S.last]: (x) => x + 30 * M });
        return H.problem({
          title: 'Break-even cartons a month per product, rounded up', file: 'titik_impas.xlsx', sheet: S, alt, fill: true, ansLabel: 'Break-even cartons',
          statement: '<p>Each product line carries its own fixed cost every month. In <b>E5</b>, work out how many cartons a month the line must sell to <b>break even</b> (cover its fixed cost), rounded <b>up</b> to a whole carton. Each carton contributes its price minus its variable cost. Fill it down to <b>E' + S.last + '</b>.</p>',
          solution: '=ROUNDUP(B5/(C5-D5),0)', alts: ['=CEILING(B5/(C5-D5),1)', '=-INT(-B5/(C5-D5))'],
          traps: [{ formula: '=ROUND(B5/(C5-D5),0)', code: 'round', title: 'ROUND can go down', plain: 'Selling a little under break-even still loses money. Use ROUNDUP so the count never falls short.' }, { formula: '=B5/(C5-D5)', code: 'unrounded', title: 'Round it up to a whole carton', plain: 'Wrap it in ROUNDUP(..., 0).' }, { formula: '=ROUNDUP(B5/C5,0)', code: 'price', title: 'The variable cost is missing', plain: 'Each carton only pays toward the fixed cost with what is left after its variable cost: C5 - D5.' }],
          hints: ['Each carton leaves price minus variable cost (C5 - D5) to pay the fixed cost.', 'Fixed cost divided by that contribution is the break-even count. ROUNDUP(x, 0) lifts it to a whole carton.', 'The shape: =ROUNDUP(B5/(C5-D5), 0)'],
          explain: 'A carton contributes C5-D5 toward the fixed cost, so B5 divided by it is the break-even count. ROUNDUP makes it a whole carton without falling short.',
          chips: ['B5/(C5-D5)', ',0)'], fnKey: 'ROUNDUP('
        });
      }
      const tp = r.pick([20, 25, 30, 40, 50]) * M;
      const base = Object.assign({}, S.cells, { A1: 'Target profit a month (Rp)', B1: tp });
      const alt = H.edit(base, { B1: tp + 15 * M, D6: (x) => x + 5000, B8: (x) => x - 15 * M });
      return H.problem({
        title: 'Cartons needed for a target monthly profit', file: 'titik_impas.xlsx', sheet: S, base, alt, fill: true, ansLabel: 'Cartons for target',
        statement: '<p>Pak Anton wants every product line to earn the profit in <b>B1</b> each month on top of its fixed cost. In <b>E5</b>, work out how many cartons a month each line must sell to reach it, rounded <b>up</b> to a whole carton. Each carton contributes its price minus its variable cost. Fill it down to <b>E' + S.last + '</b>.</p>',
        solution: '=ROUNDUP((B5+$B$1)/(C5-D5),0)', alts: ['=CEILING((B5+$B$1)/(C5-D5),1)', '=-INT(-(B5+$B$1)/(C5-D5))', '=ROUNDUP(($B$1+B5)/(C5-D5),0)'],
        traps: [{ formula: '=ROUNDUP(B5/(C5-D5),0)', code: 'no-target', title: 'That is only break-even', plain: 'Add the target profit in $B$1 to the fixed cost before dividing.' }, { formula: '=ROUNDUP((B5+B1)/(C5-D5),0)', code: 'no-pin', title: 'B1 slid down when you filled', plain: 'Pin the target with $B$1 so every row reads it.' }],
        hints: ['The cartons must now cover the fixed cost and the target profit together: B5 + $B$1.', 'Divide by what one carton contributes, C5 - D5, and round up with ROUNDUP(x, 0). Pin B1.', 'The shape: =ROUNDUP((B5+$B$1)/(C5-D5), 0)'],
        explain: 'The contribution from the cartons has to pay the fixed cost and leave the target profit, so (B5+$B$1) is divided by the contribution per carton and rounded up.',
        chips: ['(B5+$B$1)', '/(C5-D5)', ',0)'], fnKey: 'ROUNDUP('
      });
    }
  });

  P.add({
    key: 'depreciation-per-asset', difficulty: 'Medium', level: 5, topics: ['SLN', 'DDB'], count: 3,
    make(r, v) {
      const ASSETS = [['Delivery truck', 380, 520], ['Delivery van', 220, 320], ['Forklift', 260, 420], ['Cold room', 300, 650], ['Packing machine', 450, 900], ['Carton sealer', 60, 140], ['Generator set', 150, 300], ['Warehouse racking', 90, 220], ['Pallet jack', 15, 40], ['Office server', 45, 120], ['Sales motorbike', 20, 35], ['Air conditioning', 25, 80]];
      const rows = r.sample(ASSETS, 10).map(([name, lo, hi]) => {
        const cost = r.int(lo, hi) * M, life = r.pick([4, 5, 6, 8, 10]);
        return { name, cost, salv: Math.round(cost * r.pick([0.05, 0.1, 0.15, 0.2]) / M) * M, life, used: r.int(1, life) };
      });
      const cols = [{ h: 'Asset', k: 'name' }, { h: 'Cost (Rp)', k: 'cost', num: true }, { h: 'Salvage (Rp)', k: 'salv', num: true }, { h: 'Life (years)', k: 'life', num: true }];
      if (v === 0) {
        const S = H.table(cols, rows);
        const alt = H.edit(S.cells, { B5: (x) => x + 20 * M, C7: (x) => x + 5 * M, D9: (x) => (x === 10 ? 8 : x + 1), ['B' + S.last]: (x) => x + 10 * M });
        return H.problem({
          title: 'Straight-line depreciation a year per asset', file: 'aset_tetap.xlsx', sheet: S, alt, fill: true, ansLabel: 'Depreciation a year (Rp)',
          statement: '<p>Tika keeps Rasa Kenari\'s asset register. In <b>E5</b>, work out each asset\'s <b>straight-line</b> depreciation for one year: the cost less the salvage value, spread evenly over the life. Fill it down to <b>E' + S.last + '</b>. Leave the result unrounded.</p>',
          solution: '=SLN(B5,C5,D5)', alts: ['=(B5-C5)/D5'],
          traps: [{ formula: '=B5/D5', code: 'salvage', title: 'The salvage value is missing', plain: 'Only the part of the cost that is used up is depreciated: cost minus salvage.' }],
          hints: ['Straight-line depreciation is (cost - salvage) / life.', 'SLN(cost, salvage, life) does exactly that: SLN(B5, C5, D5).'],
          explain: 'SLN takes the part of the cost that will be used up, B5-C5, and spreads it evenly over the D5 years of life.',
          chips: ['B5', 'C5', 'D5'], fnKey: 'SLN('
        });
      }
      const S = H.table(cols.concat([{ h: v === 1 ? 'Years used' : 'Year', k: 'used', num: true }]), rows);
      const alt = H.edit(S.cells, { B6: (x) => x + 30 * M, C8: (x) => x + 4 * M, E10: (x) => Math.max(1, x - 1), ['B' + S.last]: (x) => x + 15 * M });
      if (v === 1) {
        return H.problem({
          title: 'Book value per asset after the years used, straight-line', file: 'aset_tetap.xlsx', sheet: S, alt, fill: true, ansLabel: 'Book value (Rp)',
          statement: '<p>In <b>F5</b>, work out each asset\'s <b>book value</b> at the end of the number of years in column E, with <b>straight-line</b> depreciation: the cost less the salvage value, spread evenly over the life. Fill it down to <b>F' + S.last + '</b>. Leave the result unrounded.</p>',
          solution: '=B5-SLN(B5,C5,D5)*E5', alts: ['=B5-(B5-C5)/D5*E5', '=B5-E5*SLN(B5,C5,D5)'],
          traps: [{ formula: '=SLN(B5,C5,D5)*E5', code: 'accumulated', title: 'That is the depreciation so far', plain: 'Book value is what is left: the cost minus the depreciation so far.' }],
          hints: ['One year of straight-line depreciation is SLN(B5, C5, D5).', 'After E5 years, E5 of those have been taken off the cost.', 'The shape: =B5 - SLN(B5, C5, D5) * E5'],
          explain: 'SLN gives one year of depreciation; E5 years of it are taken off the cost B5, which leaves the book value. At the end of the life it equals the salvage value.',
          chips: ['B5-', 'SLN(B5,C5,D5)', '*E5'], fnKey: 'SLN('
        });
      }
      return H.problem({
        title: 'Double-declining depreciation in a given year per asset', file: 'aset_tetap.xlsx', sheet: S, alt, fill: true, ansLabel: 'Depreciation (Rp)',
        statement: '<p>The tax adviser wants the <b>double-declining balance</b> method: each year takes 2/life of the value still on the books, and never goes below the salvage value. In <b>F5</b>, work out each asset\'s depreciation charge in the year shown in column E (1 is the first year). Fill it down to <b>F' + S.last + '</b>. Leave the result unrounded.</p>',
        solution: '=DDB(B5,C5,D5,E5)', alts: ['=DDB(B5,C5,D5,E5,2)'],
        traps: [{ formula: '=SLN(B5,C5,D5)', code: 'sln', title: 'That is straight-line', plain: 'Straight-line charges the same every year. Double-declining charges more in the early years: use DDB.' }],
        hints: ['Excel has a function for this method: DDB(cost, salvage, life, period).', 'The period is the year in column E: DDB(B5, C5, D5, E5).'],
        explain: 'DDB applies 2/life to the value left after the earlier years and stops at the salvage value, so its charge falls year by year. E5 tells it which year to return.',
        chips: ['B5', 'C5', 'D5', 'E5'], fnKey: 'DDB('
      });
    }
  });

  P.add({
    key: 'savings-and-borrowing', difficulty: 'Medium', level: 6, topics: ['FV', 'PV'], count: 2,
    make(r, v) {
      if (v === 0) {
        const items = [['Starting balance (Rp)', r.int(2, 20) * 10 * M], ['Deposit a month (Rp)', r.int(2, 16) * 2.5 * M], ['Rate (% a year)', r.pick([3.5, 4, 4.5, 5, 5.5, 6])], ['Years', r.pick([2, 3, 4, 5])]];
        const S = block(items);
        const wh = r.pick(H.WAREHOUSES);
        const base = Object.assign({}, S.cells, { A1: 'Fund', B1: 'Truck replacement, ' + wh, A2: 'Balance at the end (Rp)' });
        const alt = H.edit(base, { B6: (x) => x + 2.5 * M, B7: (x) => x + 0.5, B8: (x) => (x >= 5 ? 3 : x + 1) });
        return H.problem({
          title: 'Truck fund balance after monthly deposits', file: 'dana_truk.xlsx', sheet: S, base, alt, answerCell: 'B2', precision: 0,
          statement: '<p>Rasa Kenari saves for a new truck at ' + wh + '. The fund starts with the balance in <b>B5</b> and gets the deposit in <b>B6</b> at the <b>end of every month</b>. Interest is the yearly percent in <b>B7</b> (5 means 5%), compounded monthly, for the years in <b>B8</b>. In <b>B2</b>, work out the balance at the end, as a positive number of rupiah.</p>',
          solution: '=FV(B7/100/12,B8*12,-B6,-B5)', alts: ['=-FV(B7/1200,B8*12,B6,B5)', '=FV(B7/1200,12*B8,-B6,-B5)'],
          traps: [{ formula: '=FV(B7/100/12,B8*12,-B6)', code: 'start', title: 'The starting balance is missing', plain: 'The balance in B5 grows too. Give it to FV as the fourth argument, with a minus sign: -B5.' }, { formula: '=FV(B7/100/12,B8*12,B6,B5)', code: 'sign', title: 'The balance came out negative', plain: 'FV treats money you pay in as negative. Use -B6 and -B5 so the balance comes out positive.' }],
          hints: ['FV(rate, periods, payment, starting amount) grows deposits and a starting balance together.', 'Monthly rate is B7/100/12, periods are B8*12. Money paid into the fund goes in with a minus sign: -B6 and -B5.', 'The shape: =FV(B7/100/12, B8*12, -B6, -B5)'],
          explain: 'FV compounds the starting balance and every monthly deposit at the monthly rate for B8*12 months. The minus signs mark the deposits as money paid in, so the balance comes out positive.',
          chips: ['B7/100/12', 'B8*12', '-B6', '-B5'], fnKey: 'FV('
        });
      }
      const d = r.pick(DIST());
      const items = [['Installment it can pay a month (Rp)', r.int(8, 40) * 0.5 * M], ['Rate (% a year)', r.pick([9, 9.5, 10, 10.5, 11, 12])], ['Tenor (months)', r.pick([24, 36, 48, 60])]];
      const S = block(items);
      const base = Object.assign({}, S.cells, { A1: 'Distributor', B1: d.name, A2: 'Can borrow (Rp)' });
      const alt = H.edit(base, { B5: (x) => x + 1.5 * M, B6: (x) => x - 1, B7: (x) => (x >= 60 ? 36 : x + 12) });
      return H.problem({
        title: 'Largest loan a distributor can afford on its installment', file: 'kapasitas_pinjaman.xlsx', sheet: S, base, alt, answerCell: 'B2', precision: 0,
        statement: '<p>Citra at Bank Kenari checks how much ' + d.name + ' can borrow. It can pay the installment in <b>B5</b> at the end of every month, at the yearly percent in <b>B6</b> (10 means 10%), for the months in <b>B7</b>. In <b>B2</b>, work out the largest loan those installments repay, as a positive number of rupiah.</p>',
        solution: '=PV(B6/100/12,B7,-B5)', alts: ['=-PV(B6/1200,B7,B5)', '=PV(B6/1200,B7,-B5)'],
        traps: [{ formula: '=B5*B7', code: 'no-interest', title: 'That ignores the interest', plain: 'Part of every installment pays interest, so the loan is smaller than installment x months. PV takes the interest out.' }, { formula: '=PV(B6/100,B7,-B5)', code: 'yearly-rate', title: 'The rate is still yearly', plain: 'The installments are monthly, so the rate must be monthly: B6/100/12.' }],
        hints: ['The loan is today\'s value of all the future installments: PV(rate, periods, payment).', 'The rate is monthly, B6/100/12, and the installment goes in as -B5 so the loan comes out positive.', 'The shape: =PV(B6/100/12, B7, -B5)'],
        explain: 'PV discounts every installment back to today at the monthly rate and adds them up, which is the loan they can repay. The minus sign on B5 makes the loan positive.',
        chips: ['B6/100/12', 'B7', '-B5'], fnKey: 'PV('
      });
    }
  });

  P.add({
    key: 'cagr-of-a-series', difficulty: 'Medium', level: 6, topics: ['CAGR', 'RRI', 'XLOOKUP'], count: 2,
    make(r, v) {
      const years = [2019, 2020, 2021, 2022, 2023, 2024, 2025];
      const grow = (start, lo, hi, step) => { let x = start; return years.map((y, i) => { if (i) x = Math.round(x * (1 + r.int(lo, hi) / 100) / step) * step; return x; }); };
      if (v === 0) {
        const net = grow(r.int(40, 80) * 1000 * M, -3, 18, 100 * M), outlets = grow(r.int(800, 1400), 2, 12, 1);
        const S = H.table([{ h: 'Year', k: 'y', num: true }, { h: 'Net sales (Rp)', k: 'net', num: true }, { h: 'Outlets', k: 'out', num: true }], years.map((y, i) => ({ y, net: net[i], out: outlets[i] })));
        const base = Object.assign({}, S.cells, { A2: 'CAGR' });
        const alt = H.edit(base, { B11: (x) => x + 4000 * M, B5: (x) => x - 1500 * M, C11: (x) => x + 60 });
        return H.problem({
          title: 'Compound annual growth of net sales, 2019 to 2025', file: 'kinerja_tahunan.xlsx', sheet: S, base, alt, answerCell: 'B2', precision: 4,
          statement: '<p>For the annual report, work out in <b>B2</b> the <b>compound annual growth rate</b> (CAGR) of net sales from the first year in the table to the last: the steady yearly rate that turns 2019\'s figure into 2025\'s. Give it as a decimal (0.085 is 8.5%). Point at the cells so the rate updates when a figure is corrected.</p>',
          solution: '=(B11/B5)^(1/(A11-A5))-1', alts: ['=RRI(A11-A5,B5,B11)', '=(B11/B5)^(1/6)-1', '=RATE(A11-A5,0,-B5,B11)'],
          traps: [{ formula: '=(B11/B5)^(1/7)-1', code: 'seven', title: 'That counts 7 years of growth', plain: 'Seven yearly figures have only six steps of growth between them: 2025 - 2019 = 6.' }, { formula: '=(B11/B5-1)/6', code: 'average', title: 'That splits the growth evenly', plain: 'Growth compounds, so take the sixth root of B11/B5 instead of dividing by 6.' }],
          hints: ['CAGR = (last / first)^(1 / number of years between them) - 1.', 'There are A11 - A5 = 6 years of growth between 2019 and 2025. RRI(years, first, last) does the same.', 'The shape: =(B11/B5)^(1/(A11-A5))-1'],
          explain: 'B11/B5 is six years of growth multiplied together, so its sixth root is one year\'s growth factor, and minus 1 leaves the rate. RRI(6, B5, B11) gives the same number.',
          chips: ['B11/B5', '^(1/(A11-A5))', '-1'], fnKey: 'RRI('
        });
      }
      const book = grow(r.int(150, 300) * 10 * 1000 * M, 4, 22, 1000 * M), borrowers = grow(r.int(300, 600), 1, 10, 1);
      const S = H.table([{ h: 'Year', k: 'y', num: true }, { h: 'Loan book (Rp)', k: 'book', num: true }, { h: 'Borrowers', k: 'bor', num: true }], years.map((y, i) => ({ y, book: book[i], bor: borrowers[i] })));
      const y1 = r.pick([2019, 2020, 2021]), y2 = r.pick([2023, 2024, 2025]);
      const base = Object.assign({}, S.cells, { A1: 'From year', B1: y1, A2: 'To year', B2: y2, A3: 'CAGR' });
      const alt = H.edit(base, { B1: y1 === 2021 ? 2020 : y1 + 1, B2: y2 === 2023 ? 2025 : y2 - 1, B8: (x) => x + 150 * 1000 * M });
      return H.problem({
        title: 'Loan book CAGR between two chosen years', file: 'buku_pinjaman_tahunan.xlsx', sheet: S, base, alt, answerCell: 'B3', precision: 4,
        statement: '<p>Dimas compares Bank Kenari\'s loan book between the two years in <b>B1</b> and <b>B2</b>. In <b>B3</b>, work out the <b>compound annual growth rate</b> of the loan book between those years, as a decimal. Look the two figures up from the table so the answer changes when B1 or B2 changes.</p>',
        solution: '=(XLOOKUP(B2,A5:A11,B5:B11)/XLOOKUP(B1,A5:A11,B5:B11))^(1/(B2-B1))-1',
        alts: ['=RRI(B2-B1,XLOOKUP(B1,A5:A11,B5:B11),XLOOKUP(B2,A5:A11,B5:B11))', '=(INDEX(B5:B11,MATCH(B2,A5:A11,0))/INDEX(B5:B11,MATCH(B1,A5:A11,0)))^(1/(B2-B1))-1', '=(VLOOKUP(B2,A5:B11,2,FALSE)/VLOOKUP(B1,A5:B11,2,FALSE))^(1/(B2-B1))-1'],
        traps: [{ formula: '=(XLOOKUP(B2,A5:A11,C5:C11)/XLOOKUP(B1,A5:A11,C5:C11))^(1/(B2-B1))-1', code: 'borrowers', title: 'That is the number of borrowers', plain: 'The loan book in rupiah is in column B.' }],
        hints: ['Find each year\'s loan book with XLOOKUP(year, A5:A11, B5:B11).', 'CAGR = (later / earlier)^(1 / (B2 - B1)) - 1. RRI(B2-B1, earlier, later) works too.', 'The shape: =(XLOOKUP(B2,...)/XLOOKUP(B1,...))^(1/(B2-B1))-1'],
        explain: 'The two XLOOKUPs fetch the loan book for each chosen year, and B2-B1 counts the years of growth between them. The root of the ratio, minus 1, is the steady yearly rate.',
        chips: ['A5:A11', 'B5:B11', 'XLOOKUP(B1,', '^(1/(B2-B1))'], fnKey: 'XLOOKUP('
      });
    }
  });

  P.add({
    key: 'amortization-schedule', difficulty: 'Medium', level: 6, topics: ['IPMT', 'PPMT', 'FV'], count: 3,
    make(r, v) {
      const tenor = r.pick([36, 48, 60]), start = r.pick(tenor >= 48 ? [1, 13, 25] : [1, 13]);
      const amt = r.int(6, 30) * 25 * M, rate = r.pick([9, 9.5, 10, 10.75, 11.25, 12]);
      const y = 2025, m = r.int(1, 12), day = r.pick([5, 10, 15, 25]);
      const S = H.table([{ h: 'Month', k: 'k', num: true }, { h: 'Due date', k: 'due', date: true }], Array.from({ length: 12 }, (_, i) => ({ k: start + i, due: addMonths(y, m, day, start + i - 1) })));
      const base = Object.assign({}, S.cells, { A1: 'Loan amount (Rp)', B1: amt, A2: 'Rate (% a year)', B2: rate, A3: 'Tenor (months)', B3: tenor });
      const alt = H.edit(base, { B1: amt + 50 * M, B2: rate + 0.5 });
      const r_ = '$B$2/100/12', who = r.pick(DIST()).name;
      const lead = '<p>' + who + ' repays the Bank Kenari loan in <b>B1</b> in equal installments at the end of every month, at the yearly percent in <b>B2</b> (10 means 10%), over the months in <b>B3</b>. Column A is the month number. ';
      const V = [
        { title: 'Interest part of each installment in a repayment schedule', label: 'Interest (Rp)', ask: 'In <b>C5</b>, work out the <b>interest</b> part of that month\'s installment',
          sol: '=IPMT(' + r_ + ',A5,$B$3,-$B$1)', alts: ['=-IPMT($B$2/1200,A5,$B$3,$B$1)', '=IPMT($B$2/1200,A5,$B$3,-$B$1)'],
          hints: ['IPMT(rate, month, months in all, amount) returns the interest inside one installment.', 'Use the monthly rate $B$2/100/12 and pin the inputs with $ so they stay put when you fill. Use -$B$1 so the result is positive.', 'The shape: =IPMT($B$2/100/12, A5, $B$3, -$B$1)'],
          explain: 'IPMT works out the interest on the balance still owed before that month\'s payment. As the balance falls, so does the interest, while the installment stays the same.', fnKey: 'IPMT(' },
        { title: 'Principal part of each installment in a repayment schedule', label: 'Principal (Rp)', ask: 'In <b>C5</b>, work out the <b>principal</b> part of that month\'s installment (the part that pays down the loan)',
          sol: '=PPMT(' + r_ + ',A5,$B$3,-$B$1)', alts: ['=-PPMT($B$2/1200,A5,$B$3,$B$1)', '=PMT($B$2/1200,$B$3,-$B$1)-IPMT($B$2/1200,A5,$B$3,-$B$1)'],
          hints: ['PPMT(rate, month, months in all, amount) returns the principal inside one installment.', 'Use the monthly rate $B$2/100/12, pin the inputs with $, and use -$B$1 so the result is positive.', 'The shape: =PPMT($B$2/100/12, A5, $B$3, -$B$1)'],
          explain: 'PPMT is the installment minus that month\'s interest. It grows month by month because the interest on a smaller balance shrinks.', fnKey: 'PPMT(' },
        { title: 'Balance still owed after each installment', label: 'Balance (Rp)', ask: 'In <b>C5</b>, work out the <b>balance still owed</b> right after that month\'s installment is paid',
          sol: '=FV(' + r_ + ',A5,PMT(' + r_ + ',$B$3,-$B$1),-$B$1)', alts: ['=PV($B$2/1200,$B$3-A5,PMT($B$2/1200,$B$3,$B$1))', '=$B$1+CUMPRINC($B$2/1200,$B$3,$B$1,1,A5,0)', '=-FV($B$2/1200,A5,PMT($B$2/1200,$B$3,$B$1),$B$1)'],
          hints: ['The balance after A5 payments is the loan grown with interest for A5 months, less the installments paid and their interest.', 'FV(rate, A5, installment, -$B$1) does that, with the installment from PMT($B$2/100/12, $B$3, -$B$1). Or: PV of the installments still to come.', 'The shape: =FV($B$2/100/12, A5, PMT($B$2/100/12, $B$3, -$B$1), -$B$1)'],
          explain: 'FV rolls the loan forward A5 months and takes off the installments already paid, which leaves the balance owed. The same number is the PV of the installments still to come.', fnKey: 'FV(' }
      ][v];
      return H.problem({
        title: V.title, file: 'jadwal_angsuran.xlsx', sheet: S, base, alt, fill: true, ansLabel: V.label,
        statement: lead + V.ask + ' as a positive number of rupiah, and fill it down to <b>C' + S.last + '</b>. Leave the result unrounded.</p>',
        solution: V.sol, alts: V.alts,
        traps: [{ formula: V.sol.replace(/\$B\$/g, 'B'), code: 'no-pin', title: 'The inputs slid down when you filled', plain: 'B1, B2 and B3 hold one loan. Pin them with $ ($B$1, $B$2, $B$3) so every row reads the same cells.' }],
        hints: V.hints, explain: V.explain, chips: ['$B$2/100/12', 'A5', '$B$3', '-$B$1'], fnKey: V.fnKey
      });
    }
  });

  P.add({
    key: 'project-npv', difficulty: 'Medium', level: 7, topics: ['NPV', 'SUMPRODUCT'], count: 2,
    make(r, v) {
      const rate = r.pick([8, 9, 10, 11, 12, 13]);
      if (v === 0) {
        const flows = [-r.int(18, 36) * 50 * M].concat(Array.from({ length: 5 }, () => r.int(20, 50) * 10 * M));
        const S = H.table([{ h: 'Year', k: 'y', num: true }, { h: 'Cash flow (Rp)', k: 'cf', num: true }], flows.map((cf, y) => ({ y, cf })));
        const base = Object.assign({}, S.cells, { A1: 'Discount rate (% a year)', B1: rate, A2: 'NPV (Rp)' });
        const alt = H.edit(base, { B1: rate + 2, B8: (x) => x + 60 * M });
        return H.problem({
          title: 'NPV of a cold room, with the year-0 outlay', file: 'investasi_cold_room.xlsx', sheet: S, base, alt, answerCell: 'B2', precision: 0,
          statement: '<p>Rasa Kenari is weighing a cold room. Year 0 is the outlay today (negative); years 1 to 5 are the cash it brings in at the end of each year. In <b>B2</b>, work out the <b>net present value</b> at the yearly discount rate in <b>B1</b> (10 means 10%).</p>',
          solution: '=NPV(B1/100,B6:B10)+B5', alts: ['=B5+NPV(B1/100,B6:B10)', '=SUMPRODUCT(B5:B10/(1+B1/100)^A5:A10)'],
          traps: [{ formula: '=NPV(B1/100,B5:B10)', code: 'year0', title: 'Year 0 was discounted', plain: 'NPV treats its first value as one year away. The year-0 outlay happens today, so leave it out of NPV and add B5 on its own.' }, { formula: '=NPV(B1,B6:B10)+B5', code: 'percent', title: 'The rate is a percent number', plain: 'B1 holds 10 for 10%. Divide it by 100.' }],
          hints: ['NPV(rate, values) discounts each value as if it arrives one more year away, starting at one year.', 'So give NPV only years 1 to 5 (B6:B10), and add the year-0 outlay B5 outside it. The rate is B1/100.', 'The shape: =NPV(B1/100, B6:B10) + B5'],
          explain: 'NPV discounts B6:B10 as years 1 to 5. The year-0 outlay is already in today\'s money, so it is added as it is.',
          chips: ['B1/100', 'B6:B10', '+B5'], fnKey: 'NPV('
        });
      }
      const ins = [0].concat(Array.from({ length: 6 }, () => r.int(40, 75) * 10 * M)), outs = [r.int(30, 50) * 50 * M].concat(Array.from({ length: 6 }, () => r.int(12, 26) * 10 * M));
      const S = H.table([{ h: 'Year', k: 'y', num: true }, { h: 'Cash in (Rp)', k: 'i', num: true }, { h: 'Cash out (Rp)', k: 'o', num: true }], ins.map((x, y) => ({ y, i: x, o: outs[y] })));
      const base = Object.assign({}, S.cells, { A1: 'Discount rate (% a year)', B1: rate, A2: 'NPV (Rp)' });
      const alt = H.edit(base, { B1: rate - 1, C8: (x) => x + 40 * M, B11: (x) => x + 50 * M });
      return H.problem({
        title: 'NPV of a delivery fleet from cash in and cash out', file: 'investasi_armada.xlsx', sheet: S, base, alt, answerCell: 'B2', precision: 0,
        statement: '<p>Pak Anton proposes a delivery fleet. Each year has cash in (column B) and cash out (column C); year 0 is today, with the purchase in cash out. In <b>B2</b>, work out the <b>net present value</b> of the net cash flow (in minus out) at the yearly discount rate in <b>B1</b> (10 means 10%), in one formula.</p>',
        solution: '=NPV(B1/100,B6:B11-C6:C11)+B5-C5', alts: ['=B5-C5+NPV(B1/100,B6:B11-C6:C11)', '=SUMPRODUCT((B5:B11-C5:C11)/(1+B1/100)^A5:A11)', '=NPV(B1/100,B6:B11)-NPV(B1/100,C6:C11)+B5-C5'],
        traps: [{ formula: '=NPV(B1/100,B5:B11-C5:C11)', code: 'year0', title: 'Year 0 was discounted', plain: 'NPV treats its first value as one year away. Leave year 0 out of NPV and add B5-C5 on its own.' }, { formula: '=NPV(B1/100,B6:B11)+B5', code: 'no-out', title: 'Cash out is missing', plain: 'Each year\'s net flow is cash in minus cash out: B6:B11-C6:C11.' }],
        hints: ['The net flow of a year is cash in minus cash out. NPV accepts a calculated range like B6:B11-C6:C11.', 'NPV discounts its first value one year, so give it years 1 to 6 only and add year 0 (B5-C5) outside.', 'The shape: =NPV(B1/100, B6:B11-C6:C11) + B5 - C5'],
        explain: 'B6:B11-C6:C11 builds the six net yearly flows in one go, and NPV discounts them as years 1 to 6. The year-0 net flow is already today\'s money, so it is added undiscounted.',
        chips: ['B1/100', 'B6:B11-C6:C11', '+B5-C5'], fnKey: 'NPV('
      });
    }
  });

  /* ====================================== HARD ====================================== */

  P.add({
    key: 'irr-of-an-investment', difficulty: 'Hard', level: 8, topics: ['IRR', 'IF'], count: 3,
    make(r, v) {
      if (v === 0) {
        const ins = [0].concat(Array.from({ length: 6 }, () => r.int(45, 80) * 10 * M)), outs = [r.int(28, 44) * 50 * M].concat(Array.from({ length: 6 }, () => r.int(10, 24) * 10 * M));
        const S = H.table([{ h: 'Year', k: 'y', num: true }, { h: 'Cash in (Rp)', k: 'i', num: true }, { h: 'Cash out (Rp)', k: 'o', num: true }], ins.map((x, y) => ({ y, i: x, o: outs[y] })));
        /* The resale sits in B1, outside the table, so the formula has to add it to the last year only. */
        const resale = r.int(4, 10) * 25 * M, wh = r.pick(H.WAREHOUSES);
        const base = Object.assign({}, S.cells, { A1: 'Resale at the end (Rp)', B1: resale, A2: 'IRR' });
        const alt = H.edit(base, { B1: resale + 75 * M, C5: (x) => x + 150 * M, B8: (x) => x - 50 * M });
        return H.problem({
          title: 'IRR of a packing line with a resale value at the end', file: 'investasi_lini.xlsx', sheet: S, base, alt, answerCell: 'B2', precision: 4,
          statement: '<p>Rasa Kenari plans a packing line at ' + wh + '. Each year has cash in (column B) and cash out (column C); year 0 is today and holds the purchase. At the end of the last year in the table (year 6) the line is sold for the resale value in <b>B1</b>, which is not in the table. In <b>B2</b>, work out the <b>internal rate of return</b> of the net cash flow (cash in minus cash out, plus the resale in year 6 only), as a decimal, in one formula with no helper column.</p>',
          solution: '=IRR(B5:B11-C5:C11+(A5:A11=A11)*B1)',
          alts: ['=IRR(B5:B11-C5:C11+(A5:A11=6)*B1)', '=IRR(B5:B11-C5:C11+IF(A5:A11=A11,B1,0))', '=LET(cf,B5:B11-C5:C11+(A5:A11=MAX(A5:A11))*B1,IRR(cf))', '=IRR(VSTACK(B5:B10-C5:C10,B11-C11+B1))'],
          traps: [
            { formula: '=IRR(B5:B11-C5:C11)', code: 'no-resale', title: 'The resale value is missing', plain: 'The line is sold at the end of year 6 for the value in B1. Add B1 to the last year\'s net flow.' },
            { formula: '=IRR(B5:B11-C5:C11+B1)', code: 'every-year', title: 'The resale was added to every year', plain: 'The resale comes in once, at the end of year 6. Add B1 only where the year is the last one: (A5:A11=A11)*B1.' },
            { formula: '=IRR(B6:B11-C6:C11+(A6:A11=A11)*B1)', code: 'year0', title: 'Year 0 is missing', plain: 'IRR needs the outlay too. Start the ranges at row 5, year 0.' }
          ],
          hints: ['IRR finds the rate at which the net present value of the flows is zero. It needs every year, year 0 included, and accepts a calculated range such as B5:B11-C5:C11.', 'To add B1 to the last year only, multiply it by a TRUE/FALSE test: (A5:A11=A11)*B1 is B1 in year 6 and 0 in every other year.', 'The shape: =IRR(B5:B11-C5:C11+(A5:A11=A11)*B1)'],
          explain: 'B5:B11-C5:C11 builds each year\'s net flow, and (A5:A11=A11)*B1 adds the resale to year 6 only, because TRUE counts as 1 and FALSE as 0. IRR then finds the rate that makes the NPV of the seven flows zero.',
          chips: ['B5:B11-C5:C11', '(A5:A11=A11)*B1'], fnKey: 'IRR('
        });
      }
      if (v === 1) {
        /* Accept on the visible sheet and Reject on the hidden one, so a typed word cannot pass. Accept comes first so
           the common slip (comparing with 12 instead of 0.12, always Reject) shows its own message. */
        const make = () => [-r.int(16, 30) * 50 * M].concat(Array.from({ length: 5 }, () => r.int(25, 60) * 10 * M));
        let cf = make(); while (irr(cf) < 0.08 || irr(cf) > 0.4) cf = make();
        const cf2 = cf.slice(); cf2[3] = cf[3] + (r.chance(0.5) ? 1 : -1) * 80 * M; cf2[5] = cf[5] - 40 * M;
        const acceptFirst = true;
        const hurdle = (x, accept) => (accept ? Math.floor(x * 100) - r.int(1, 3) : Math.ceil(x * 100) + r.int(1, 3));
        const S = H.table([{ h: 'Year', k: 'y', num: true }, { h: 'Cash flow (Rp)', k: 'cf', num: true }], cf.map((x, y) => ({ y, cf: x })));
        const base = Object.assign({}, S.cells, { A1: 'Hurdle rate (% a year)', B1: hurdle(irr(cf), acceptFirst), A2: 'Decision' });
        const alt = H.edit(base, { B1: hurdle(irr(cf2), !acceptFirst), B8: cf2[3], B10: cf2[5] });
        return H.problem({
          title: 'Accept or reject a project against a hurdle rate', file: 'keputusan_investasi.xlsx', sheet: S, base, alt, answerCell: 'B2',
          statement: '<p>Rasa Kenari only funds projects that earn at least the hurdle rate in <b>B1</b> (12 means 12%). The table holds the project\'s yearly cash flows, year 0 first. In <b>B2</b>, return <code>Accept</code> when the project\'s IRR is <b>at least</b> the hurdle rate, otherwise <code>Reject</code>.</p>',
          solution: '=IF(IRR(B5:B10)>=B1/100,"Accept","Reject")', alts: ['=IF(IRR(B5:B10)<B1/100,"Reject","Accept")', '=IF(IRR(B5:B10)*100>=B1,"Accept","Reject")', '=IF(NPV(B1/100,B6:B10)+B5>=0,"Accept","Reject")'],
          traps: [{ formula: '=IF(IRR(B5:B10)>=B1,"Accept","Reject")', code: 'percent', title: 'The hurdle is a percent number', plain: 'IRR returns a decimal like 0.14, and B1 holds 12 for 12%. Compare with B1/100.' }],
          hints: ['IRR(B5:B10) gives the project\'s return as a decimal.', 'Compare it with B1/100 inside IF, and return the two words in double quotes.', 'The shape: =IF(IRR(B5:B10)>=B1/100, "Accept", "Reject")'],
          explain: 'IRR gives the rate the project earns, and IF compares it with the hurdle as a decimal. An IRR at or above the hurdle means the project is worth at least what the money costs.',
          chips: ['IRR(B5:B10)', '>=B1/100', '"Accept"', '"Reject"'], fnKey: 'IF('
        });
      }
      const price = r.int(8, 14) * 50 * M, rent = r.int(5, 9) * 10 * M, resale = r.int(4, 8) * 50 * M;
      const flows = [-price].concat(Array.from({ length: 12 }, (_, i) => rent + (i === 11 ? resale : 0)));
      const S = H.table([{ h: 'Quarter', k: 'q', num: true }, { h: 'Cash flow (Rp)', k: 'cf', num: true }], flows.map((cf, q) => ({ q, cf })));
      const base = Object.assign({}, S.cells, { A1: 'Lease', B1: 'Truck lease to ' + r.pick(DIST()).name, A2: 'Yearly return' });
      const alt = H.edit(base, { B5: (x) => x - 50 * M, B17: (x) => x - 100 * M });
      return H.problem({
        title: 'Effective yearly return on a quarterly truck lease', file: 'sewa_truk.xlsx', sheet: S, base, alt, answerCell: 'B2', precision: 4,
        statement: '<p>Bank Kenari\'s leasing arm buys a truck (quarter 0) and collects rent every <b>quarter</b> for three years, plus the resale value in the last quarter. In <b>B2</b>, work out the <b>effective yearly</b> rate of return: the quarterly IRR compounded over four quarters, as a decimal.</p>',
        solution: '=(1+IRR(B5:B17))^4-1', alts: ['=EFFECT(IRR(B5:B17)*4,4)', '=LET(q,IRR(B5:B17),(1+q)^4-1)'],
        traps: [{ formula: '=IRR(B5:B17)*4', code: 'nominal', title: 'That is the nominal rate', plain: 'Four times the quarterly rate ignores compounding. Grow (1 + quarterly IRR) to the power 4, then subtract 1.' }, { formula: '=IRR(B5:B17)', code: 'quarterly', title: 'That is a quarterly rate', plain: 'The flows are quarterly, so IRR returns a rate per quarter. Turn it into a yearly rate.' }],
        hints: ['IRR(B5:B17) gives the return per quarter, because the flows are a quarter apart.', 'A year has four quarters of compounding: (1 + quarterly rate)^4 - 1.', 'The shape: =(1+IRR(B5:B17))^4-1'],
        explain: 'IRR works in the period of the flows, here a quarter. Compounding that rate four times and subtracting 1 gives the effective yearly return.',
        chips: ['IRR(B5:B17)', '(1+', ')^4-1'], fnKey: 'IRR('
      });
    }
  });

  P.add({
    key: 'loan-balance-and-interest', difficulty: 'Hard', level: 8, topics: ['FV', 'CUMIPMT', 'PMT'], count: 3,
    make(r, v) {
      const amt = r.int(8, 40) * 25 * M, rate = r.pick([9, 9.5, 10, 10.5, 11, 11.75, 12.5]), tenor = r.pick([36, 48, 60, 72]);
      const S = block([['Loan amount (Rp)', amt], ['Rate (% a year)', rate], ['Tenor (months)', tenor]]);
      const d = r.pick(DIST()).name;
      const lead = '<p>' + d + ' has the Bank Kenari loan in <b>B5</b>, repaid in equal installments at the end of every month at the yearly percent in <b>B6</b> (10 means 10%) over the months in <b>B7</b>. ';
      if (v === 0) {
        const k = r.int(7, tenor - 12);
        const base = Object.assign({}, S.cells, { A1: 'Installments paid', B1: k, A2: 'Balance owed (Rp)' });
        const alt = H.edit(base, { B1: k + r.pick([5, 9, -4]), B5: amt + 50 * M });
        return H.problem({
          title: 'Balance still owed after a number of installments', file: 'saldo_pinjaman.xlsx', sheet: S, base, alt, answerCell: 'B2', precision: 0,
          statement: lead + 'Citra needs the balance for a top-up request. In <b>B2</b>, work out the <b>balance still owed</b> right after the number of installments in <b>B1</b> has been paid, as a positive number of rupiah.</p>',
          solution: '=FV(B6/1200,B1,PMT(B6/1200,B7,-B5),-B5)', alts: ['=PV(B6/1200,B7-B1,PMT(B6/1200,B7,B5))', '=B5+CUMPRINC(B6/1200,B7,B5,1,B1,0)', '=FV(B6/100/12,B1,PMT(B6/100/12,B7,-B5),-B5)'],
          traps: [{ formula: '=B5-PMT(B6/1200,B7,-B5)*B1', code: 'no-interest', title: 'Part of each installment was interest', plain: 'Only the principal part of each installment pays down the loan, so the balance is higher than amount minus installments paid.' }],
          hints: ['Two ways in: roll the loan forward B1 months less the installments (FV), or value the installments still to come (PV).', 'The installment is PMT(B6/1200, B7, -B5). FV(B6/1200, B1, installment, -B5) gives the balance after B1 payments.', 'The shape: =FV(B6/1200, B1, PMT(B6/1200, B7, -B5), -B5)'],
          explain: 'FV grows the loan for B1 months and takes off the B1 installments with their interest, which leaves what is still owed. It equals the PV of the B7-B1 installments that remain.',
          chips: ['B6/1200', 'PMT(B6/1200,B7,-B5)', '-B5', 'B1'], fnKey: 'FV('
        });
      }
      if (v === 1) {
        const fee = r.pick([0.5, 0.75, 1, 1.25, 1.5, 2]);
        const base = Object.assign({}, S.cells, { A1: 'Upfront fee (% of loan)', B1: fee, A2: 'Cost of credit (Rp)' });
        const alt = H.edit(base, { B1: fee + 0.5, B6: rate + 1, B5: amt + 50 * M });
        return H.problem({
          title: 'Total cost of credit: interest plus an upfront fee', file: 'biaya_kredit.xlsx', sheet: S, base, alt, answerCell: 'B2', precision: 0,
          statement: lead + 'The bank also charges an upfront fee, a percent of the loan in <b>B1</b> (1 means 1%). In <b>B2</b>, work out the <b>total cost of credit</b>: all the interest paid over the whole loan plus the fee, as a positive number of rupiah.</p>',
          solution: '=PMT(B6/1200,B7,-B5)*B7-B5+B5*B1/100', alts: ['=-CUMIPMT(B6/1200,B7,B5,1,B7,0)+B5*B1/100', '=B7*PMT(B6/100/12,B7,-B5)-B5*(1-B1/100)'],
          traps: [{ formula: '=PMT(B6/1200,B7,-B5)*B7-B5', code: 'no-fee', title: 'The fee is missing', plain: 'Add the upfront fee: B5*B1/100.' }, { formula: '=PMT(B6/1200,B7,-B5)*B7+B5*B1/100', code: 'principal', title: 'That counts the loan itself as a cost', plain: 'The installments repay the loan as well as the interest. Take B5 away to leave only the interest.' }],
          hints: ['Over the whole loan the installments add up to PMT x B7. Everything above the loan amount is interest.', 'Interest = PMT(B6/1200, B7, -B5)*B7 - B5. The fee is B5*B1/100.', 'The shape: =PMT(B6/1200,B7,-B5)*B7 - B5 + B5*B1/100'],
          explain: 'All the installments together repay the loan plus every rupiah of interest, so subtracting B5 leaves the interest. The fee is added on top as a share of the loan.',
          chips: ['PMT(B6/1200,B7,-B5)', '*B7-B5', '+B5*B1/100'], fnKey: 'PMT('
        });
      }
      /* Year 2 or later on the visible sheet, so interest-to-date (from installment 1) differs from one year's. */
      const yr = r.int(2, tenor / 12 - 1);
      const base = Object.assign({}, S.cells, { A1: 'Loan year', B1: yr, A2: 'Interest that year (Rp)' });
      const alt = H.edit(base, { B1: yr === 2 ? 3 : yr - 1, B5: amt + 75 * M });
      return H.problem({
        title: 'Interest paid during a chosen year of the loan', file: 'bunga_tahunan.xlsx', sheet: S, base, alt, answerCell: 'B2', precision: 0,
        statement: lead + 'For the distributor\'s tax return, work out in <b>B2</b> the <b>interest paid during loan year B1</b>: year 1 is installments 1 to 12, year 2 is 13 to 24, and so on. Give it as a positive number of rupiah. (The practice sheet\'s IPMT takes a single month, not a list of months.)</p>',
        solution: '=-CUMIPMT(B6/1200,B7,B5,12*B1-11,12*B1,0)', alts: ['=-CUMIPMT(B6/100/12,B7,B5,12*(B1-1)+1,12*B1,0)', '=12*PMT(B6/1200,B7,-B5)-(FV(B6/1200,12*B1-12,PMT(B6/1200,B7,-B5),-B5)-FV(B6/1200,12*B1,PMT(B6/1200,B7,-B5),-B5))'],
        traps: [{ formula: '=-CUMIPMT(B6/1200,B7,B5,1,12*B1,0)', code: 'to-date', title: 'That is all the interest so far', plain: 'The task wants one year only: start at installment 12*B1-11, not at 1.' }],
        hints: ['CUMIPMT(rate, months, amount, first month, last month, 0) adds up the interest between two installments.', 'Loan year B1 runs from installment 12*B1-11 to 12*B1. CUMIPMT returns a negative number, so put a minus in front.', 'The shape: =-CUMIPMT(B6/1200, B7, B5, 12*B1-11, 12*B1, 0)'],
        explain: 'CUMIPMT adds up the interest inside installments 12*B1-11 to 12*B1, which is loan year B1. The last argument 0 says payments come at the end of each month.',
        chips: ['B6/1200', '12*B1-11', '12*B1', ',0)'], fnKey: 'CUMIPMT('
      });
    }
  });

  P.add({
    key: 'months-and-rates-from-an-offer', difficulty: 'Hard', level: 9, topics: ['NPER', 'RATE', 'ROUNDUP'], count: 3,
    make(r, v) {
      const d = r.pick(DIST()).name;
      if (v === 0) {
        const owed = r.int(8, 30) * 10 * M, rate = r.pick([10, 12, 14, 15, 18]), pay = r.int(6, 20) * M;
        const S = block([['Amount owed (Rp)', owed], ['Rate (% a year)', rate], ['Payment a month (Rp)', pay]]);
        const base = Object.assign({}, S.cells, { A1: 'Distributor', B1: d, A2: 'Months to clear' });
        const alt = H.edit(base, { B7: pay + 2 * M, B5: owed + 20 * M });
        return H.problem({
          title: 'Months to clear an overdue balance on a payment plan', file: 'rencana_pelunasan.xlsx', sheet: S, base, alt, answerCell: 'B2',
          statement: '<p>Bu Lestari agrees a payment plan with ' + d + ' for the overdue balance in <b>B5</b>. Interest runs at the yearly percent in <b>B6</b> (12 means 12%), charged monthly, and the distributor pays <b>B7</b> at the end of every month. In <b>B2</b>, work out how many months it takes to clear the balance, counting the last, smaller payment as a full month (round <b>up</b>).</p>',
          solution: '=ROUNDUP(NPER(B6/1200,-B7,B5),0)', alts: ['=CEILING(NPER(B6/1200,-B7,B5),1)', '=ROUNDUP(NPER(B6/100/12,B7,-B5),0)', '=-INT(-NPER(B6/1200,-B7,B5))'],
          traps: [{ formula: '=ROUNDUP(B5/B7,0)', code: 'no-interest', title: 'That ignores the interest', plain: 'Interest keeps adding to the balance, so it takes longer than amount / payment. NPER counts it in.' }, { formula: '=NPER(B6/1200,-B7,B5)', code: 'unrounded', title: 'Round up to whole months', plain: 'NPER gives a fraction of a month. Wrap it in ROUNDUP(..., 0).' }],
          hints: ['NPER(rate, payment, amount) counts how many payments clear a balance.', 'Use the monthly rate B6/1200, the payment as -B7 (money going out) and the balance B5. Then ROUNDUP(..., 0).', 'The shape: =ROUNDUP(NPER(B6/1200, -B7, B5), 0)'],
          explain: 'NPER finds how many payments of B7 clear B5 while interest runs at the monthly rate. The answer has a fraction for the smaller last payment, so ROUNDUP counts it as a full month.',
          chips: ['B6/1200', '-B7', 'B5', ',0)'], fnKey: 'NPER('
        });
      }
      if (v === 1) {
        const price = r.int(18, 40) * 10 * M, months = r.pick([24, 36, 48]), hidden = r.pick([9, 10.5, 12, 13.5, 15]) / 1200;
        const inst = Math.round(price * hidden / (1 - Math.pow(1 + hidden, -months)) / 1000) * 1000;
        const S = block([['Van price financed (Rp)', price], ['Installment a month (Rp)', inst], ['Months', months]]);
        const base = Object.assign({}, S.cells, { A1: 'Dealer', B1: 'Van dealer, ' + r.pick(H.REGIONS), A2: 'Effective yearly rate' });
        const alt = H.edit(base, { B6: inst + r.int(3, 8) * 100000, B5: price - 5 * M });
        return H.problem({
          title: 'Effective yearly rate hidden in a van installment plan', file: 'cicilan_van.xlsx', sheet: S, base, alt, answerCell: 'B2', precision: 4,
          statement: '<p>A dealer offers ' + d + ' a delivery van: the price in <b>B5</b> is financed and repaid with the installment in <b>B6</b> at the end of every month for the months in <b>B7</b>. The dealer quotes no rate. In <b>B2</b>, work out the <b>effective yearly rate</b>: the monthly rate the plan charges, compounded over 12 months, as a decimal.</p>',
          solution: '=(1+RATE(B7,-B6,B5))^12-1', alts: ['=EFFECT(RATE(B7,-B6,B5)*12,12)', '=(1+RATE(B7,B6,-B5))^12-1', '=LET(m,RATE(B7,-B6,B5),(1+m)^12-1)'],
          traps: [{ formula: '=RATE(B7,-B6,B5)*12', code: 'nominal', title: 'That is the nominal rate', plain: 'Twelve times the monthly rate ignores compounding. Use (1 + monthly rate)^12 - 1.' }, { formula: '=RATE(B7,-B6,B5)', code: 'monthly', title: 'That is a monthly rate', plain: 'RATE works per payment period, here a month. Turn it into an effective yearly rate.' }],
          hints: ['RATE(months, installment, price) finds the monthly rate. The installment is money going out, so use -B6.', 'Compound the monthly rate over a year: (1 + rate)^12 - 1. EFFECT(rate*12, 12) gives the same.', 'The shape: =(1+RATE(B7, -B6, B5))^12-1'],
          explain: 'RATE finds the monthly rate at which B7 installments of B6 repay B5. Compounding it over 12 months gives the effective yearly rate the plan really charges.',
          chips: ['RATE(B7,-B6,B5)', '(1+', ')^12-1'], fnKey: 'RATE('
        });
      }
      const amt = r.int(8, 30) * 25 * M, rate = r.pick([9.5, 10, 10.5, 11, 12]), tenor = r.pick([48, 60, 72]), extra = r.int(2, 8) * 0.5 * M;
      const S = block([['Loan amount (Rp)', amt], ['Rate (% a year)', rate], ['Tenor (months)', tenor]]);
      const base = Object.assign({}, S.cells, { A1: 'Extra payment a month (Rp)', B1: extra, A2: 'Months saved' });
      const alt = H.edit(base, { B1: extra + 1.5 * M, B5: amt + 50 * M });
      return H.problem({
        title: 'Months saved by paying extra on top of each installment', file: 'pelunasan_dipercepat.xlsx', sheet: S, base, alt, answerCell: 'B2',
        statement: '<p>' + d + ' has the loan in <b>B5</b> at the yearly percent in <b>B6</b> (10 means 10%), with equal monthly installments over the months in <b>B7</b>. It decides to pay the extra amount in <b>B1</b> on top of every installment from the start. In <b>B2</b>, work out how many months sooner the loan is cleared: B7 minus the new number of months, with the new count rounded <b>up</b> to whole months.</p>',
        solution: '=B7-ROUNDUP(NPER(B6/1200,-(PMT(B6/1200,B7,-B5)+B1),B5),0)', alts: ['=B7-ROUNDUP(NPER(B6/1200,PMT(B6/1200,B7,B5)-B1,B5),0)', '=LET(p,PMT(B6/1200,B7,-B5)+B1,B7-ROUNDUP(NPER(B6/1200,-p,B5),0))', '=B7-CEILING(NPER(B6/1200,-(PMT(B6/1200,B7,-B5)+B1),B5),1)'],
        traps: [{ formula: '=ROUNDUP(NPER(B6/1200,-(PMT(B6/1200,B7,-B5)+B1),B5),0)', code: 'new-count', title: 'That is the new number of months', plain: 'The task wants the months saved: B7 minus this.' }],
        hints: ['First the normal installment: PMT(B6/1200, B7, -B5). The new monthly payment is that plus B1.', 'NPER(B6/1200, -(new payment), B5) counts the months at the new payment. Round it up, then take it from B7.', 'The shape: =B7 - ROUNDUP(NPER(B6/1200, -(PMT(B6/1200,B7,-B5)+B1), B5), 0)'],
        explain: 'PMT gives the installment the loan was built on; adding B1 gives the larger payment. NPER counts how many of those clear the loan, and B7 minus that count is the months saved.',
        chips: ['PMT(B6/1200,B7,-B5)', '+B1', 'NPER(B6/1200,', ',B5)', ',0)'], fnKey: 'NPER('
      });
    }
  });

  P.add({
    key: 'flat-rate-offers', difficulty: 'Hard', level: 9, topics: ['RATE', 'PMT'], count: 2,
    make(r, v) {
      /* Flat-rate credit charges interest on the full price for the whole tenor, common in vehicle finance. */
      const lenders = ['Bank Kenari', 'Truck dealer', 'Leasing company', 'Cooperative', 'Dealer promo', 'Rural bank'];
      const rows = Array.from({ length: r.int(8, 10) }, (_, i) => ({ id: 'OF-' + String(i + 1).padStart(2, '0'), who: r.pick(lenders), flat: r.pick([4.5, 5, 5.5, 6, 6.5, 7, 7.5, 8]), months: r.pick([12, 24, 36, 48]) }));
      const S = H.table([{ h: 'Offer', k: 'id' }, { h: 'Lender', k: 'who' }, { h: 'Flat rate (% a year)', k: 'flat', num: true }, { h: 'Months', k: 'months', num: true }], rows);
      const price = r.int(30, 60) * 10 * M;
      const base = Object.assign({}, S.cells, { A1: 'Price financed (Rp)', B1: price });
      const alt = H.edit(base, { B1: price + 50 * M, C6: (x) => x + 1, D8: (x) => (x === 48 ? 24 : x + 12) });
      const lead = '<p>Distributors get truck finance quoted at a <b>flat rate</b>: interest is charged on the full price in <b>B1</b> for the whole tenor, so the monthly installment is B1 * (1 + flat rate * months / 12) / months. Column C is the flat rate as a yearly percent (6 means 6%) and column D the months. ';
      if (v === 0) {
        return H.problem({
          title: 'True yearly rate behind each flat-rate truck offer', file: 'penawaran_kredit.xlsx', sheet: S, base, alt, fill: true, ansLabel: 'True rate a year', precision: 4,
          statement: lead + 'In <b>E5</b>, work out the <b>true yearly rate</b> of each offer: the monthly rate at which those installments repay B1 (as RATE finds it), times 12, as a decimal. Fill it down to <b>E' + S.last + '</b>.</p>',
          solution: '=RATE(D5,-$B$1*(1+C5/100*D5/12)/D5,$B$1)*12', alts: ['=12*RATE(D5,-($B$1+$B$1*C5/100*D5/12)/D5,$B$1)', '=LET(i,$B$1*(1+C5/100*D5/12)/D5,RATE(D5,-i,$B$1)*12)', '=RATE(D5,$B$1*(1+C5/100*D5/12)/D5,-$B$1)*12'],
          traps: [{ formula: '=C5/100', code: 'flat', title: 'That is the flat rate itself', plain: 'The flat rate charges interest on money already repaid, so the true rate is higher. Build the installment, then let RATE find the real monthly rate.' }, { formula: '=RATE(D5,-$B$1*(1+C5/100*D5/12)/D5,$B$1)', code: 'monthly', title: 'That is a monthly rate', plain: 'Multiply by 12 for the yearly rate the task asks for.' }],
          hints: ['Build the installment first: $B$1*(1+C5/100*D5/12)/D5.', 'RATE(months, -installment, price) finds the monthly rate that makes those installments repay the price. Times 12 gives the yearly rate. Pin B1.', 'The shape: =RATE(D5, -$B$1*(1+C5/100*D5/12)/D5, $B$1)*12'],
          explain: 'A flat rate charges interest on the full price even after much of it is repaid, so the installments hide a higher true rate. RATE finds the monthly rate those installments actually pay, and 12 times it is the yearly rate.',
          chips: ['$B$1*(1+C5/100*D5/12)/D5', 'RATE(D5,', ',$B$1)', '*12'], fnKey: 'RATE('
        });
      }
      return H.problem({
        title: 'Extra interest from a flat rate versus a declining balance', file: 'penawaran_kredit.xlsx', sheet: S, base, alt, fill: true, ansLabel: 'Extra interest (Rp)',
        statement: lead + 'An ordinary loan at the same yearly percent charges interest only on the balance still owed (equal installments from PMT). In <b>E5</b>, work out how much <b>more interest</b> the flat-rate offer costs than that ordinary loan over the whole tenor. Fill it down to <b>E' + S.last + '</b>. Leave the result unrounded.</p>',
        solution: '=$B$1*C5/100*D5/12-(PMT(C5/1200,D5,-$B$1)*D5-$B$1)', alts: ['=$B$1*(1+C5/100*D5/12)-PMT(C5/1200,D5,-$B$1)*D5', '=$B$1*C5/100*D5/12+CUMIPMT(C5/1200,D5,$B$1,1,D5,0)'],
        traps: [{ formula: '=$B$1*C5/100*D5/12', code: 'flat-only', title: 'That is the flat interest alone', plain: 'Take away the interest of the ordinary loan: PMT(C5/1200, D5, -$B$1)*D5 - $B$1.' }],
        hints: ['Flat interest over the tenor is $B$1 * C5/100 * D5/12.', 'Ordinary-loan interest is all its installments minus the price: PMT(C5/1200, D5, -$B$1)*D5 - $B$1. Pin B1.', 'The shape: =$B$1*C5/100*D5/12 - (PMT(C5/1200,D5,-$B$1)*D5 - $B$1)'],
        explain: 'The flat offer charges interest on the whole price for every year, while the ordinary loan charges it only on what is still owed. The gap between the two interest totals is what the flat quote really costs extra.',
        chips: ['$B$1*C5/100*D5/12', 'PMT(C5/1200,D5,-$B$1)', '*D5-$B$1'], fnKey: 'PMT('
      });
    }
  });

  P.add({
    key: 'declining-balance-book-value', difficulty: 'Hard', level: 9, topics: ['DDB', 'MAX'], count: 2,
    make(r, v) {
      const ASSETS = [['Delivery truck', 380, 520], ['Forklift', 260, 420], ['Cold room', 300, 650], ['Packing machine', 450, 900], ['Generator set', 150, 300], ['Warehouse racking', 90, 220], ['Delivery van', 220, 320], ['Carton sealer', 60, 140], ['Office server', 45, 120], ['Labelling machine', 120, 260]];
      const rows = r.sample(ASSETS, 10).map(([name, lo, hi], i) => {
        const cost = r.int(lo, hi) * M, life = r.pick([4, 5, 6, 8]);
        /* Some rows run late into the life, where the salvage floor takes over. */
        return { name, cost, salv: Math.round(cost * r.pick([0.1, 0.15, 0.2, 0.25]) / M) * M, life, used: i % 3 === 0 ? life : r.int(1, life) };
      });
      const S = H.table([{ h: 'Asset', k: 'name' }, { h: 'Cost (Rp)', k: 'cost', num: true }, { h: 'Salvage (Rp)', k: 'salv', num: true }, { h: 'Life (years)', k: 'life', num: true }, { h: 'Years used', k: 'used', num: true }], rows);
      /* The hidden sheet moves costs, salvage values and years, so rows on and off the salvage floor both change. */
      const alt = H.edit(S.cells, { B5: (x) => x + 20 * M, C5: (x) => x + 3 * M, B6: (x) => x + 40 * M, E7: (x) => Math.max(1, x - 1), B8: (x) => x + 30 * M, C9: (x) => x + 8 * M, E11: (x) => Math.max(1, x - 1), ['B' + S.last]: (x) => x + 25 * M, ['C' + S.last]: (x) => x + 2 * M });
      if (v === 0) {
        return H.problem({
          title: 'Book value per asset under double-declining balance', file: 'aset_saldo_menurun.xlsx', sheet: S, alt, fill: true, ansLabel: 'Book value (Rp)',
          statement: '<p>Under the <b>double-declining balance</b> method (Excel\'s DDB), each year takes 2/life of the value still on the books, and the value never falls below salvage. In <b>F5</b>, work out each asset\'s <b>book value</b> at the end of the years in column E, in one formula per row (no column per year; the practice sheet\'s DDB takes a single year, not a list of years). Fill it down to <b>F' + S.last + '</b>. Leave the result unrounded.</p>',
          solution: '=MAX(C5,B5*(1-2/D5)^E5)', alts: ['=MAX(B5*(1-2/D5)^E5,C5)', '=B5-MIN(B5-C5,B5-B5*(1-2/D5)^E5)', '=IF(B5*(1-2/D5)^E5<C5,C5,B5*(1-2/D5)^E5)'],
          traps: [{ formula: '=B5*(1-2/D5)^E5', code: 'floor', title: 'It fell below the salvage value', plain: 'DDB stops at salvage. Take the larger of this and the salvage value with MAX.' }, { formula: '=B5-DDB(B5,C5,D5,E5)', code: 'one-year', title: 'That takes off one year only', plain: 'DDB returns the charge for a single year. The book value has lost every year up to E5.' }],
          hints: ['Each year keeps (1 - 2/life) of the value, so after E5 years the value is B5*(1-2/D5)^E5.', 'DDB never takes the value under salvage, so the book value is the larger of that and C5.', 'The shape: =MAX(C5, B5*(1-2/D5)^E5)'],
          explain: 'Every year DDB keeps 1-2/D5 of the value, so E5 years multiply the cost by that factor E5 times. Once that would drop under salvage, DDB stops at salvage, which MAX enforces.',
          chips: ['B5*(1-2/D5)^E5', 'MAX(C5,'], fnKey: 'MAX('
        });
      }
      return H.problem({
        title: 'Depreciation so far per asset at a 150% declining rate', file: 'aset_saldo_menurun.xlsx', sheet: S, alt, fill: true, ansLabel: 'Depreciation so far (Rp)',
        statement: '<p>The auditor asks for the <b>150% declining balance</b> method (Excel\'s DDB with factor 1.5): each year takes 1.5/life of the value still on the books, and the value never falls below salvage. In <b>F5</b>, work out each asset\'s <b>total depreciation so far</b>, at the end of the years in column E, in one formula per row (the practice sheet\'s DDB takes a single year, not a list of years). Fill it down to <b>F' + S.last + '</b>. Leave the result unrounded.</p>',
        solution: '=B5-MAX(C5,B5*(1-1.5/D5)^E5)', alts: ['=MIN(B5-C5,B5-B5*(1-1.5/D5)^E5)', '=MIN(B5-C5,B5*(1-(1-1.5/D5)^E5))'],
        traps: [{ formula: '=MAX(C5,B5*(1-1.5/D5)^E5)', code: 'book-value', title: 'That is the book value', plain: 'The task wants what has been taken off so far: cost minus book value.' }, { formula: '=B5-MAX(C5,B5*(1-2/D5)^E5)', code: 'factor', title: 'That uses the double rate', plain: 'This method takes 1.5/life a year, not 2/life.' }],
        hints: ['Each year keeps (1 - 1.5/life) of the value, so the book value after E5 years is B5*(1-1.5/D5)^E5, but not below salvage.', 'Book value is MAX(C5, B5*(1-1.5/D5)^E5). Depreciation so far is the cost minus that.', 'The shape: =B5 - MAX(C5, B5*(1-1.5/D5)^E5)'],
        explain: 'The book value shrinks by the factor 1-1.5/D5 every year until it reaches salvage, which MAX keeps as a floor. Everything taken off the cost so far is B5 minus that book value.',
        chips: ['B5*(1-1.5/D5)^E5', 'MAX(C5,', 'B5-'], fnKey: 'MAX('
      });
    }
  });

  P.add({
    key: 'payback-period', difficulty: 'Hard', level: 9, topics: ['XLOOKUP', 'MATCH', 'NPER'], count: 2,
    make(r, v) {
      if (v === 0) {
        /* Cash flows with their running total; the payback year is where the total first turns non-negative. */
        const build = () => {
          const cf = [-r.int(20, 40) * 50 * M].concat(Array.from({ length: 6 }, () => r.int(25, 70) * 10 * M));
          let c = 0; const cum = cf.map((x) => (c += x));
          return { cf, cum };
        };
        const fine = (b) => b.cum[6] > 0 && b.cum.indexOf(b.cum.find((x) => x >= 0)) >= 2 && !b.cum.includes(0);
        let A = build(); while (!fine(A)) A = build();
        /* The hidden sheet: a weaker year 2 pushes the payback later. */
        const cf2 = A.cf.slice(); cf2[2] = Math.max(100 * M, cf2[2] - 150 * M); cf2[1] = cf2[1] - 50 * M;
        let c2 = 0; const cum2 = cf2.map((x) => (c2 += x));
        if (!(cum2[6] > 0) || cum2.includes(0)) { cf2[6] += 400 * M; c2 = 0; cum2.splice(0, 7, ...cf2.map((x) => (c2 += x))); }
        const S = H.table([{ h: 'Year', k: 'y', num: true }, { h: 'Cash flow (Rp)', k: 'cf', num: true }, { h: 'Cumulative (Rp)', k: 'cum', num: true }], A.cf.map((cf, y) => ({ y, cf, cum: A.cum[y] })));
        const base = Object.assign({}, S.cells, { A1: 'Project', B1: 'Second packing line', A2: 'Payback (years)' });
        const edits = {}; cf2.forEach((x, i) => { edits['B' + (5 + i)] = x; edits['C' + (5 + i)] = cum2[i]; });
        const alt = H.edit(base, edits);
        return H.problem({
          title: 'Payback period in years, with the fraction of the year', file: 'periode_balik_modal.xlsx', sheet: S, base, alt, answerCell: 'B2', precision: 4,
          statement: '<p>Column C is the running total of the cash flows. The project pays back during the first year in which the running total turns zero or positive. Cash arrives evenly through each year, so the payback in years is that year minus the share of its cash flow still unused: <b>payback year - cumulative that year / cash flow that year</b>. In <b>B2</b>, work out the payback in years, with the fraction, in one formula that still works when the figures change.</p>',
          solution: '=LET(k,MATCH(TRUE,C5:C11>=0,0),INDEX(A5:A11,k)-INDEX(C5:C11,k)/INDEX(B5:B11,k))',
          alts: ['=XLOOKUP(TRUE,C5:C11>=0,A5:A11)-XLOOKUP(TRUE,C5:C11>=0,C5:C11)/XLOOKUP(TRUE,C5:C11>=0,B5:B11)', '=LET(k,MATCH(TRUE,C5:C11>=0,0),INDEX(A5:A11,k-1)-INDEX(C5:C11,k-1)/INDEX(B5:B11,k))', '=COUNTIF(C5:C11,"<0")-1-INDEX(C5:C11,COUNTIF(C5:C11,"<0"))/INDEX(B5:B11,COUNTIF(C5:C11,"<0")+1)'],
          traps: [{ formula: '=XLOOKUP(TRUE,C5:C11>=0,A5:A11)', code: 'whole-year', title: 'That is only the payback year', plain: 'Take off the share of that year\'s cash flow not needed: cumulative that year / cash flow that year.' }],
          hints: ['C5:C11>=0 is TRUE from the payback year on. XLOOKUP(TRUE, C5:C11>=0, A5:A11) finds that year; MATCH(TRUE, C5:C11>=0, 0) finds its position.', 'With the position k, take INDEX(A5:A11,k) - INDEX(C5:C11,k)/INDEX(B5:B11,k). LET keeps k in one place.', 'The shape: =LET(k, MATCH(TRUE, C5:C11>=0, 0), INDEX(A5:A11,k) - INDEX(C5:C11,k)/INDEX(B5:B11,k))'],
          explain: 'MATCH(TRUE, C5:C11>=0, 0) finds the first year the running total is not negative. The running total that year is the cash not needed, so its share of the year\'s flow comes off the year number.',
          chips: ['C5:C11>=0', 'MATCH(TRUE,', 'INDEX(A5:A11,k)', 'INDEX(C5:C11,k)', 'INDEX(B5:B11,k)'], fnKey: 'LET('
        });
      }
      const inv = r.int(10, 30) * 50 * M, save = Math.round(inv * r.pick([0.18, 0.22, 0.25, 0.3]) / (10 * M)) * 10 * M, rate = r.pick([8, 9, 10, 11, 12]);
      const S = block([['Investment today (Rp)', inv], ['Saving a year (Rp)', save], ['Discount rate (% a year)', rate]]);
      const wh = r.pick(H.WAREHOUSES);
      const base = Object.assign({}, S.cells, { A1: 'Project', B1: 'Solar roof, ' + wh, A2: 'Discounted payback (years)' });
      const alt = H.edit(base, { B6: save + 30 * M, B7: rate + 2 });
      return H.problem({
        title: 'Discounted payback of a solar roof with equal yearly savings', file: 'atap_surya.xlsx', sheet: S, base, alt, answerCell: 'B2', precision: 4,
        statement: '<p>A solar roof at ' + wh + ' costs <b>B5</b> today and saves <b>B6</b> at the end of every year. At the yearly discount rate in <b>B7</b> (10 means 10%), the <b>discounted payback</b> is the number of years n at which the present value of n yearly savings equals the investment: <code>B6 * (1 - (1 + rate)^-n) / rate = B5</code>, where rate is B7/100. Solve that equation for n, which need not be a whole number (do not interpolate between whole years). In <b>B2</b>, work out n, unrounded.</p>',
        solution: '=NPER(B7/100,B6,-B5)', alts: ['=NPER(B7/100,-B6,B5)', '=LN(B6/(B6-B5*B7/100))/LN(1+B7/100)', '=-LN(1-B5*B7/100/B6)/LN(1+B7/100)'],
        traps: [{ formula: '=B5/B6', code: 'simple', title: 'That is the simple payback', plain: 'It ignores discounting. Later savings are worth less today, so the discounted payback is longer.' }],
        hints: ['Equal yearly savings are an annuity: find how many of them, discounted, add up to B5.', 'NPER(rate, payment, present value) counts the periods. Give the saving and the investment opposite signs.', 'The shape: =NPER(B7/100, B6, -B5)'],
        explain: 'NPER counts how many yearly savings of B6, discounted at B7/100, it takes to repay -B5 today. That count is the discounted payback; it is longer than B5/B6 because later savings are worth less.',
        chips: ['B7/100', 'B6', '-B5'], fnKey: 'NPER('
      });
    }
  });
});
