/*
 * Practice, Dates (30 Medium, 15 Hard): DATE, YEAR, MONTH, EOMONTH, EDATE, WEEKDAY, NETWORKDAYS, WORKDAY, DATEDIF,
 * due dates from terms, days past due against an as-of date, aging buckets and late interest.
 * "Today" is always a fixed as-of date in an input cell, never TODAY(), so the hidden sheet can move it.
 * The engine does not run YEAR, MONTH, DAY, WEEKDAY or EOMONTH over a whole range, nor a lookup with many lookup
 * values at once, so the Hard problems work with date comparisons, SUMIFS, SUMPRODUCT on plain date arithmetic,
 * MAXIFS and FILTER, which give the same results in Excel 365.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) factory(require('./core.js'));
  else factory(root.KicauPractice);
})(typeof globalThis !== 'undefined' ? globalThis : this, function (P) {
  'use strict';
  const H = P.H;

  /* ---------- date helpers, all on Excel date numbers ---------- */
  const pad = (n) => String(n).padStart(2, '0');
  const ser = (y, m, d) => H.toSerial(y + '-' + pad(m) + '-' + pad(d));
  /* Like WEEKDAY(s,2): Monday 1 to Sunday 7. */
  const wk = (s) => ((s + 5) % 7) + 1;
  const long = (s) => { const t = H.iso(s); return +t.slice(8, 10) + ' ' + H.MONTHS[+t.slice(5, 7) - 1] + ' ' + t.slice(0, 4); };
  /* Like EOMONTH(s,k). */
  const eom = (s, k) => { const t = H.iso(s); return Math.round((Date.UTC(+t.slice(0, 4), +t.slice(5, 7) + k, 1) - Date.UTC(1899, 11, 30)) / 864e5) - 1; };
  /* Like NETWORKDAYS(a,b,hol), only used to choose data. */
  const nwd = (a, b, hol) => { let n = 0; for (let s = a; s <= b; s++) if (wk(s) < 6 && !hol.includes(s)) n++; return n; };
  const span = (lo, hi) => Array.from({ length: hi - lo + 1 }, (_, i) => lo + i);
  const amount = (r) => r.int(15, 480) * 100000;
  const DIST = () => H.db().distributors;
  function refs(r, prefix, n, from) { let x = from + r.int(0, 60); const out = []; for (let i = 0; i < n; i++) { x += r.int(2, 13); out.push(prefix + x); } return out; }
  const ASOF = [ser(2026, 11, 30), ser(2026, 12, 31), ser(2027, 1, 15), ser(2027, 1, 31)];

  /* Company holidays, every one on a weekday. */
  const HOL_Q4 = [['2026-11-27', 'Stock count day'], ['2026-12-18', 'Company outing'], ['2026-12-24', 'Christmas Eve leave'], ['2026-12-25', 'Christmas'], ['2026-12-28', 'Year-end leave'], ['2026-12-31', 'Year-end closing'], ['2027-01-01', 'New Year']];
  const HOL_YEAR = [['2026-11-27', 'Stock count day'], ['2026-12-18', 'Company outing'], ['2026-12-24', 'Christmas Eve leave'], ['2026-12-25', 'Christmas'], ['2026-12-31', 'Year-end closing'], ['2027-01-01', 'New Year'], ['2027-01-29', 'Stock count day'], ['2027-02-15', 'Company leave'], ['2027-03-11', 'Public holiday'], ['2027-03-31', 'Quarter closing']];
  const holidayRows = (r, n) => r.sample(HOL_Q4, n).map(([d, name]) => ({ date: H.toSerial(d), name })).sort((a, b) => a.date - b.date);
  const holidayTable = (rows, startCol) => H.table([{ h: 'Holiday', k: 'date', date: true }, { h: 'Name', k: 'name' }], rows, { startCol });
  /* A sheet with a small table to its right, shown as one grid. */
  function beside(S, T) {
    const cols = S.cols.slice(), last = H.LETTERS.indexOf(T.cols[T.cols.length - 1]);
    for (let j = H.LETTERS.indexOf(cols[cols.length - 1]) + 1; j <= last; j++) cols.push(H.LETTERS[j]);
    return Object.assign({}, S, { cols, num: S.num.concat(T.num), dateCols: S.dateCols.concat(T.dateCols) });
  }

  /* Invoices whose due dates sit `off` days before the as-of date (a negative off is not yet due). */
  function aged(r, asof, offs, o) {
    const opt = o || {};
    const pool = opt.pool || DIST();
    const key = opt.sortBy || 'due';
    const rows = offs.map((off) => {
      const d = r.pick(pool), due = asof - off;
      const terms = r.chance(.75) ? d.terms_days : r.pick([14, 60]);
      const date = due - terms;
      let paid = '';
      if (opt.paid && date < asof - 1 && r.chance(opt.paid)) paid = Math.max(date + 1, Math.min(asof - r.int(0, 3), due + r.int(-8, 6)));
      return { dist: d.name, due, terms, date, amount: amount(r), paid, off };
    }).sort((a, b) => a[key] - b[key]);
    refs(r, 'INV-', rows.length, 4100).forEach((x, i) => { rows[i].ref = x; });
    return rows;
  }
  /* The hidden sheet's payment news: one overdue unpaid invoice gets paid, one paid invoice is reversed. */
  function payNews(cells, S, rows, paidOn) {
    const c = Object.assign({}, cells), col = S.col('Paid on');
    const unpaid = rows.findIndex((x) => x.paid === '' && x.off > 10);
    const paid = rows.findIndex((x) => x.paid !== '' && x.off > 10);
    if (unpaid >= 0) c[col + (S.first + unpaid)] = paidOn;
    if (paid >= 0) delete c[col + (S.first + paid)];
    return c;
  }
  const C_INV = { h: 'Invoice', k: 'ref' }, C_DIST = { h: 'Distributor', k: 'dist' }, C_DUE = { h: 'Due date', k: 'due', date: true }, C_DATE = { h: 'Date', k: 'date', date: true };
  const C_TERMS = { h: 'Terms (days)', k: 'terms', num: true }, C_AMT = { h: 'Amount (Rp)', k: 'amount', num: true }, C_PAID = { h: 'Paid on', k: 'paid', date: true };

  /* ======================= Medium ======================= */

  P.add({
    key: 'date-from-parts', difficulty: 'Medium', level: 4, topics: ['DATE', 'YEAR', 'MONTH', 'LEFT'], count: 3,
    make(r, v) {
      const n = 12;
      if (v === 0) {
        /* Planned deliveries: a month and a day, in the year of the order. */
        const rows = span(1, n).map(() => {
          const y = r.chance(.3) ? 2027 : 2026, om = r.int(1, 11), od = r.int(1, 24);
          const dm = om + (r.chance(.5) ? 1 : 0), dd = dm === om ? r.int(od + 1, 28) : r.int(1, 28);
          return { dist: r.pick(DIST()).name, date: ser(y, om, od), dm, dd };
        }).sort((a, b) => a.date - b.date);
        refs(r, 'SO-', n, 5200).forEach((x, i) => { rows[i].ref = x; });
        const S = H.table([{ h: 'Order', k: 'ref' }, C_DIST, { h: 'Order date', k: 'date', date: true }, { h: 'Month', k: 'dm', num: true }, { h: 'Day', k: 'dd', num: true }], rows);
        /* The hidden sheet: two orders were keyed in the wrong year, one delivery day moved. */
        const alt = H.edit(S.cells, { C6: (x) => x + 365, C11: (x) => x + 365, E8: (x) => (x > 26 ? x - 1 : x + 2) });
        return H.problem({
          title: 'Delivery date from a month and day in the order year', file: 'pesanan_2026.xlsx', sheet: S, alt, fill: true, ansLabel: 'Delivery date', ansDate: true,
          statement: '<p>Each order has a planned delivery <b>Month</b> (D) and <b>Day</b> (E), in the same year as its <b>Order date</b> (C). In <b>F5</b>, build the delivery date and fill down to <b>F' + S.last + '</b>.</p><p>The grid shows dates as day and month; each date also holds its year, 2026 or 2027.</p>',
          solution: '=DATE(YEAR(C5),D5,E5)', alts: ['=DATE(YEAR(C5),D5,1)+E5-1', '=DATE(YEAR($C5),$D5,$E5)'],
          hints: ['DATE(year, month, day) builds a date from three numbers.', 'The year comes from the order date: YEAR(C5).', 'The shape is =DATE(YEAR(C5),D5,E5).'],
          explain: 'YEAR(C5) takes the year from the order date, and DATE joins it with the month in D5 and the day in E5. Every part comes from a cell, so a corrected order year gives a new delivery date.',
          chips: ['YEAR(C5)', 'D5', 'E5'], fnKey: 'DATE('
        });
      }
      if (v === 1) {
        const months = [[2026, 8], [2026, 9], [2026, 10], [2026, 11], [2026, 12], [2027, 1], [2027, 2], [2027, 3]];
        const rows = span(1, n).map(() => { const [y, m] = r.pick(months); return { dist: r.pick(DIST()).name, date: ser(y, m, r.int(2, 28)), amount: amount(r) }; }).sort((a, b) => a.date - b.date);
        refs(r, 'INV-', n, 3300).forEach((x, i) => { rows[i].ref = x; });
        const S = H.table([C_INV, C_DIST, C_DATE, C_AMT], rows);
        /* The hidden sheet: two invoices were re-dated into the following month. */
        const alt = H.edit(S.cells, { C7: (x) => x + 31, C13: (x) => x + 33 });
        return H.problem({
          title: "First day of each invoice's month", file: 'piutang_q4.xlsx', sheet: S, alt, fill: true, ansLabel: 'Month start', ansDate: true,
          statement: '<p>The monthly report groups invoices by the first day of their month. In <b>E5</b>, return the first day of the month of the invoice <b>Date</b> in C5 (17 Nov 2026 gives 1 Nov 2026), then fill down to <b>E' + S.last + '</b>.</p>',
          solution: '=DATE(YEAR(C5),MONTH(C5),1)', alts: ['=EOMONTH(C5,-1)+1', '=C5-DAY(C5)+1'],
          hints: ['Keep the year and the month of C5, and set the day to 1.', 'YEAR(C5) and MONTH(C5) give the parts you keep.', 'The shape is =DATE(YEAR(C5),MONTH(C5),1).'],
          explain: 'DATE(YEAR(C5), MONTH(C5), 1) rebuilds the date with day 1, the first day of the same month. EOMONTH(C5,-1)+1, the day after last month ended, is the same date.',
          chips: ['YEAR(C5)', 'MONTH(C5)', ',1)'], fnKey: 'DATE('
        });
      }
      const shift = (c, k) => { const m = +c.slice(4) - 1 + k; return String((+c.slice(0, 4) + Math.floor(m / 12)) * 100 + (m % 12) + 1); };
      const rows = r.sample(span(0, 23), n).sort((a, b) => a - b).map((k) => ({ code: String((2026 + Math.floor(k / 12)) * 100 + (k % 12) + 1), dist: r.pick(DIST()).name, target: r.pick([4, 5, 6, 8, 10, 12, 15]) * 100000000 }));
      const S = H.table([{ h: 'Period', k: 'code' }, C_DIST, { h: 'Target (Rp)', k: 'target', num: true }], rows);
      /* The hidden sheet: two periods were corrected. */
      const alt = H.edit(S.cells, { A6: (c) => shift(c, 2), A12: (c) => shift(c, 13) });
      return H.problem({
        title: 'First day of a text period code like 202611', file: 'rencana_target.xlsx', sheet: S, alt, fill: true, ansLabel: 'Period start', ansDate: true,
        statement: '<p>The planning file stores each <b>Period</b> as text: four digits of year, then two of month (202611 is November 2026). In <b>D5</b>, turn the code in A5 into a real date, the first day of that month, and fill down to <b>D' + S.last + '</b>.</p>',
        solution: '=DATE(LEFT(A5,4),RIGHT(A5,2),1)', alts: ['=DATE(VALUE(LEFT(A5,4)),VALUE(RIGHT(A5,2)),1)', '=DATE(LEFT(A5,4),MID(A5,5,2),1)'],
        hints: ['LEFT(A5,4) gives the year digits and RIGHT(A5,2) the month digits.', 'DATE accepts those pieces of text and reads them as numbers.', 'The shape is =DATE(LEFT(A5,4),RIGHT(A5,2),1).'],
        explain: 'LEFT and RIGHT cut the year and the month out of the code, and DATE turns them into a date number. Once it is a real date, EOMONTH, sorting and date filters work on it.',
        chips: ['LEFT(A5,4)', 'RIGHT(A5,2)', ',1)'], fnKey: 'DATE('
      });
    }
  });

  P.add({
    key: 'month-and-quarter', difficulty: 'Medium', level: 4, topics: ['MONTH', 'ROUNDUP', 'TEXT'], count: 3,
    make(r, v) {
      const n = 14;
      let dates;
      if (v === 2) dates = span(1, n).map(() => r.pick([ser(2026, 11, 1), ser(2026, 12, 1), ser(2027, 1, 1), ser(2027, 2, 1)]) + r.int(0, 27));
      else dates = r.shuffle(span(1, 12)).concat(r.sample(span(1, 12), 2)).map((m) => ser(2026, m, r.int(1, 28)));
      const rows = dates.sort((a, b) => a - b).map((date) => ({ date, dist: r.pick(DIST()).name, net: amount(r) }));
      refs(r, 'INV-', n, 2600).forEach((x, i) => { rows[i].ref = x; });
      const S = H.table([{ h: 'Date', k: 'date', date: true }, C_INV, C_DIST, { h: 'Net (Rp)', k: 'net', num: true }], rows);
      /* The hidden sheet: three sales were re-dated. */
      const alt = H.edit(S.cells, { A5: (x) => x + 92, A10: (x) => x + 40, A14: (x) => x + 45 });
      const last = 'E' + S.last;
      const V = [
        { title: 'Quarter label Q1 to Q4 from a sale date', label: 'Quarter',
          statement: '<p>In <b>E5</b>, label each sale with its calendar quarter: <b>Q1</b> for January to March, <b>Q2</b> for April to June, <b>Q3</b> for July to September and <b>Q4</b> for October to December. Fill down to <b>' + last + '</b>. Write the label exactly, with a capital Q.</p>',
          solution: '="Q"&ROUNDUP(MONTH(A5)/3,0)', alts: ['="Q"&INT((MONTH(A5)+2)/3)', '="Q"&CHOOSE(MONTH(A5),1,1,1,2,2,2,3,3,3,4,4,4)'],
          hints: ['MONTH(A5) gives 1 to 12. Divided by 3 and rounded up, it gives 1 to 4.', 'ROUNDUP(x,0) rounds up to a whole number; join "Q" in front with &.', 'The shape is ="Q"&ROUNDUP(MONTH(A5)/3,0).'],
          explain: 'Months 1 to 3 divided by 3 give 0.33 to 1, which round up to 1; months 10 to 12 round up to 4. Joining "Q" in front gives the label.',
          chips: ['"Q"&', 'MONTH(A5)', '/3,0)'], fnKey: 'ROUNDUP(' },
        { title: 'Short month name from a sale date', label: 'Month',
          statement: '<p>In <b>E5</b>, write the short English month name of the sale <b>Date</b>: Jan, Feb, Mar and so on, three letters with a capital first letter. Fill down to <b>' + last + '</b>.</p><p>On Indonesian settings TEXT writes Indonesian names (Okt, Des), so there use CHOOSE with the twelve English names.</p>',
          solution: '=TEXT(A5,"mmm")', alts: ['=CHOOSE(MONTH(A5),"Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")'],
          hints: ['TEXT shows a date in a format you choose and returns it as text.', 'The format code "mmm" gives the three-letter month name.', 'The shape is =TEXT(A5,"mmm").'],
          explain: 'TEXT(A5,"mmm") formats the date as its short month name. CHOOSE(MONTH(A5),"Jan",...,"Dec") gives the same names on any regional setting.',
          chips: ['A5', '"mmm"'], fnKey: 'TEXT(' },
        { title: 'Year-month key for sales across New Year', label: 'Period key',
          statement: '<p>These sales run from November 2026 into early 2027. A pivot table needs a period key written as year, dash, two-digit month: <b>2026-11</b>, <b>2027-01</b>. In <b>E5</b>, build the key from the date in A5 and fill down to <b>' + last + '</b>.</p>',
          solution: '=TEXT(A5,"yyyy-mm")', alts: ['=YEAR(A5)&"-"&TEXT(MONTH(A5),"00")', '=YEAR(A5)&"-"&IF(MONTH(A5)<10,"0","")&MONTH(A5)'],
          hints: ['The key needs the year, because January 2027 and January 2026 are different periods.', 'TEXT with the code "yyyy-mm" writes both parts, with a leading zero on the month.', 'The shape is =TEXT(A5,"yyyy-mm").'],
          explain: 'The "yyyy-mm" code writes four digits of year and two of month, so the keys also sort in time order as text. YEAR(A5)&"-"&TEXT(MONTH(A5),"00") builds the same key piece by piece.',
          chips: ['A5', '"yyyy-mm"'], fnKey: 'TEXT(' }
      ][v];
      return H.problem(Object.assign({ file: 'penjualan_2026.xlsx', sheet: S, alt, fill: true, ansLabel: V.label, exactCase: true }, V, { label: undefined }));
    }
  });

  P.add({
    key: 'due-date-from-terms', difficulty: 'Medium', level: 5, topics: ['EOMONTH', 'DATE', 'RIGHT'], count: 3,
    make(r, v) {
      const n = 12;
      const rows = r.sample(span(ser(2026, 10, 1), ser(2026, 12, 28)), n).sort((a, b) => a - b).map((date) => {
        const d = r.pick(DIST());
        return { date, dist: d.name, terms: 'Net ' + (r.chance(.7) ? d.terms_days : r.pick([14, 60])), amount: amount(r) };
      });
      refs(r, 'INV-', n, 3800).forEach((x, i) => { rows[i].ref = x; });
      if (v === 0) {
        const S = H.table([C_INV, C_DIST, C_DATE, { h: 'Terms', k: 'terms' }, C_AMT], rows);
        /* The hidden sheet: one distributor's terms were renegotiated, one invoice re-dated. */
        const alt = H.edit(S.cells, { D7: (t) => (t === 'Net 30' ? 'Net 45' : 'Net 30'), C12: (x) => x + 9 });
        return H.problem({
          title: 'Due date from text terms like Net 30', file: 'piutang_q4.xlsx', sheet: S, alt, fill: true, ansLabel: 'Due date', ansDate: true,
          statement: '<p>Each invoice\'s <b>Terms</b> are written as text, such as <b>Net 30</b>: payment is due that many days after the invoice <b>Date</b>. Every term has two digits. In <b>F5</b>, return the due date and fill down to <b>F' + S.last + '</b>.</p>',
          solution: '=C5+RIGHT(D5,2)', alts: ['=C5+VALUE(RIGHT(D5,2))', '=C5+SUBSTITUTE(D5,"Net ","")', '=C5+MID(D5,5,2)'],
          hints: ['A date is a count of days, so adding 30 to it moves it 30 days on.', 'RIGHT(D5,2) takes the two digits; Excel reads them as a number when you add them to a date.', 'The shape is =C5+RIGHT(D5,2).'],
          explain: 'RIGHT(D5,2) pulls the number of days out of the text, and adding it to the invoice date moves the date forward that many days. VALUE(RIGHT(D5,2)) does the same with the conversion written out.',
          chips: ['C5+', 'RIGHT(D5,2)'], fnKey: 'RIGHT('
        });
      }
      const S = H.table([C_INV, C_DIST, C_DATE, C_AMT], rows);
      if (v === 1) {
        /* The hidden sheet: two invoices were re-dated into the next month. */
        const alt = H.edit(S.cells, { C6: (x) => x + 31, C11: (x) => x + 31 });
        return H.problem({
          title: 'Due at the end of the month after the invoice', file: 'piutang_modern_trade.xlsx', sheet: S, alt, fill: true, ansLabel: 'Due date', ansDate: true,
          statement: '<p>Modern-trade buyers pay at the end of the month after the invoice month: an invoice dated 12 Oct 2026 is due 30 Nov 2026. In <b>E5</b>, return each due date and fill down to <b>E' + S.last + '</b>.</p>',
          solution: '=EOMONTH(C5,1)', alts: ['=DATE(YEAR(C5),MONTH(C5)+2,0)', '=EOMONTH(EOMONTH(C5,0)+1,0)'],
          hints: ['EOMONTH(date, months) returns the last day of a month, counted from the month of the date.', 'Months 0 is the same month; 1 is the month after.', 'The shape is =EOMONTH(C5,1).'],
          explain: 'EOMONTH(C5,1) moves one month on from the invoice month and returns that month\'s last day, whatever the invoice day. DATE(YEAR(C5),MONTH(C5)+2,0) reaches the same day, because day 0 of a month is the last day of the month before.',
          chips: ['C5', ',1)'], fnKey: 'EOMONTH('
        });
      }
      const days = r.pick([30, 45]);
      const base = Object.assign({}, S.cells, { A1: 'Days after month end', B1: days });
      /* The hidden sheet: the terms grow by 15 days, and one invoice is re-dated. */
      const alt = H.edit(base, { B1: days + 15, C8: (x) => x + 31 });
      return H.problem({
        title: 'Due a set number of days after the invoice month ends', file: 'piutang_q4.xlsx', sheet: S, base, alt, fill: true, ansLabel: 'Due date', ansDate: true,
        statement: '<p>Some distributors pay on end-of-month terms: the number of days in <b>B1</b> is counted from the last day of the invoice month. With 30 days, an invoice dated 12 Oct 2026 is due 31 Oct plus 30 days, which is 30 Nov 2026. In <b>E5</b>, return each due date and fill down to <b>E' + S.last + '</b>.</p>',
        solution: '=EOMONTH(C5,0)+$B$1', alts: ['=DATE(YEAR(C5),MONTH(C5)+1,0)+$B$1', '=$B$1+EOMONTH(C5,0)'],
        hints: ['First find the last day of the invoice month with EOMONTH(C5,0).', 'Then add the days in B1, written $B$1 so it stays put as you fill down.', 'The shape is =EOMONTH(C5,0)+$B$1.'],
        explain: 'EOMONTH(C5,0) is the last day of the invoice month, and adding $B$1 counts the terms from there. The hidden sheet changes the days in B1, so pointing at B1 keeps every due date right.',
        chips: ['C5,0)', '$B$1'], fnKey: 'EOMONTH('
      });
    }
  });

  P.add({
    key: 'days-past-due', difficulty: 'Medium', level: 5, topics: ['MAX', 'IF'], count: 3,
    make(r, v) {
      const asof = r.pick(ASOF), asofAlt = asof + r.int(7, 12);
      const offs = r.sample([-25, -12, -4, 0, 3, 9, 16, 24, 30, 33, 41, 52, 60, 68, 75, 88, 97, 110, 126], 12);
      const rows = aged(r, asof, offs, { paid: v === 2 ? 0.4 : 0, sortBy: v === 1 ? 'date' : 'due' });
      const top = { A1: 'As of', B1: asof };
      const note = ' The as-of date in B1 is ' + long(asof) + '.';
      if (v === 0) {
        const S = H.table([C_INV, C_DIST, C_DUE, C_AMT], rows);
        const base = Object.assign({}, S.cells, top);
        const alt = H.edit(base, { B1: asofAlt, C7: (x) => x - 5 });
        return H.problem({
          title: 'Days past due against an as-of date', file: 'piutang_q4.xlsx', sheet: S, base, alt, fill: true, ansLabel: 'Days past due', dateCells: ['B1'],
          statement: '<p>Days past due is how many days the as-of date in <b>B1</b> comes after the <b>Due date</b>. An invoice that is not yet due, or due on B1 itself, has <b>0</b>. In <b>E5</b>, work it out for each invoice and fill down to <b>E' + S.last + '</b>.' + note + '</p>',
          solution: '=MAX(0,$B$1-C5)', alts: ['=IF($B$1>C5,$B$1-C5,0)', '=MAX($B$1-C5,0)'],
          hints: ['Subtracting two dates gives the days between them: $B$1-C5.', 'A negative result means not yet due; MAX(0, ...) turns it into 0.', 'The shape is =MAX(0,$B$1-C5).'],
          explain: '$B$1-C5 counts the days from the due date to the as-of date, and MAX(0, ...) puts invoices that are not yet due at 0. The $ keeps every row measured against the same as-of date.',
          chips: ['$B$1', 'C5', 'MAX(0,'], fnKey: 'MAX('
        });
      }
      if (v === 1) {
        const S = H.table([C_INV, C_DIST, C_DATE, C_TERMS, C_AMT], rows);
        const base = Object.assign({}, S.cells, top);
        const alt = H.edit(base, { B1: asofAlt, D8: (t) => (t === 30 ? 45 : 30) });
        return H.problem({
          title: 'Days past due from invoice date plus terms', file: 'piutang_q4.xlsx', sheet: S, base, alt, fill: true, ansLabel: 'Days past due', dateCells: ['B1'],
          statement: '<p>This sheet has the invoice <b>Date</b> and the <b>Terms</b> in days, but no due date: the due date is Date plus Terms. In <b>F5</b>, return the days past due as of <b>B1</b> (0 when not yet due) and fill down to <b>F' + S.last + '</b>.' + note + '</p>',
          solution: '=MAX(0,$B$1-(C5+D5))', alts: ['=MAX(0,$B$1-C5-D5)', '=IF($B$1>C5+D5,$B$1-C5-D5,0)'],
          hints: ['The due date of row 5 is C5+D5.', 'Days past due are $B$1 minus that due date, and never below 0.', 'The shape is =MAX(0,$B$1-(C5+D5)).'],
          explain: 'C5+D5 rebuilds the due date inside the formula, $B$1 minus it counts the days late, and MAX(0, ...) keeps invoices that are not yet due at 0.',
          chips: ['$B$1', 'C5+D5', 'MAX(0,'], fnKey: 'MAX('
        });
      }
      const S = H.table([C_INV, C_DIST, C_DUE, C_AMT, C_PAID], rows);
      const base = Object.assign({}, S.cells, top);
      const alt = H.edit(payNews(base, S, rows, asof - 1), { B1: asofAlt });
      return H.problem({
        title: 'Days past due, counting unpaid invoices only', file: 'piutang_q4.xlsx', sheet: S, base, alt, fill: true, ansLabel: 'Days past due', dateCells: ['B1'],
        statement: '<p>Only unpaid invoices can be past due. An invoice with a date in <b>Paid on</b> (E) scores <b>0</b>. For the rest, return the days the as-of date in <b>B1</b> comes after the <b>Due date</b>, or 0 when not yet due. Type it in <b>F5</b> and fill down to <b>F' + S.last + '</b>.' + note + '</p>',
        solution: '=IF(E5="",MAX(0,$B$1-C5),0)', alts: ['=IF(E5<>"",0,MAX(0,$B$1-C5))', '=IF(ISBLANK(E5),MAX(0,$B$1-C5),0)', '=(E5="")*MAX(0,$B$1-C5)'],
        hints: ['An empty Paid on cell means the invoice is still open: E5="".', 'For an open invoice use MAX(0,$B$1-C5); for a paid one return 0.', 'The shape is =IF(E5="",MAX(0,$B$1-C5),0).'],
        explain: 'IF first checks whether the invoice was paid. Open invoices get their days past due, floored at 0 by MAX, and paid ones get 0.',
        chips: ['E5=""', 'MAX(0,$B$1-C5)', ',0)'], fnKey: 'IF('
      });
    }
  });

  P.add({
    key: 'weekday-checks', difficulty: 'Medium', level: 6, topics: ['WEEKDAY', 'IF', 'MOD'], count: 3,
    make(r, v) {
      const n = 12, lo = ser(2026, 11, 2), hi = ser(2026, 12, 20);
      const dates = span(0, n - 1).map((i) => { let s = lo + r.int(0, hi - lo); if (i % 4 === 0 && wk(s) < 6) s += (r.chance(.5) ? 6 : 7) - wk(s); return s; }).sort((a, b) => a - b);
      const rows = dates.map((date) => ({ date, dist: r.pick(DIST()).name, cartons: r.int(8, 60) }));
      refs(r, 'DO-', n, 7100).forEach((x, i) => { rows[i].ref = x; });
      const S = H.table([{ h: 'Delivery', k: 'ref' }, C_DIST, C_DATE, { h: 'Cartons', k: 'cartons', num: true }], rows);
      /* The hidden sheet: one weekday delivery slipped to Saturday, another moved a week. */
      const i = rows.findIndex((x) => wk(x.date) < 6), j = (i + 5) % n;
      const alt = H.edit(S.cells, { ['C' + (S.first + i)]: (x) => x + 6 - wk(x), ['C' + (S.first + j)]: (x) => x + 7 });
      const last = 'E' + S.last;
      const V = [
        { title: 'Flag deliveries that land on a weekend', ansLabel: 'Day type', exactCase: true,
          statement: '<p>Deliveries on Saturday or Sunday are paid an overtime rate. In <b>E5</b>, write <b>Weekend</b> when the delivery <b>Date</b> falls on a Saturday or Sunday and <b>Weekday</b> otherwise, then fill down to <b>' + last + '</b>.</p>',
          solution: '=IF(WEEKDAY(C5,2)>5,"Weekend","Weekday")', alts: ['=IF(WEEKDAY(C5,2)<6,"Weekday","Weekend")', '=IF(OR(WEEKDAY(C5)=1,WEEKDAY(C5)=7),"Weekend","Weekday")'],
          hints: ['WEEKDAY(C5,2) numbers the days Monday 1 to Sunday 7.', 'With that numbering, Saturday and Sunday are the days above 5.', 'The shape is =IF(WEEKDAY(C5,2)>5,"Weekend","Weekday").'],
          explain: 'With return type 2, WEEKDAY gives 6 for Saturday and 7 for Sunday, so one test, >5, catches both. Without the 2, Sunday is 1 and Saturday is 7, which needs an OR.',
          chips: ['WEEKDAY(C5,2)', '>5', '"Weekend"', '"Weekday"'], fnKey: 'WEEKDAY(' },
        { title: 'Monday that starts the week of each delivery', ansLabel: 'Week of', ansDate: true,
          statement: '<p>The weekly delivery report groups by week, and every week starts on Monday. In <b>E5</b>, return the Monday of the week that holds the delivery <b>Date</b>: a Monday returns itself, a Sunday returns the Monday six days before. Fill down to <b>' + last + '</b>.</p>',
          solution: '=C5-WEEKDAY(C5,2)+1', alts: ['=C5-WEEKDAY(C5,3)', '=C5+1-WEEKDAY(C5,2)'],
          hints: ['WEEKDAY(C5,2) is 1 on Monday, 2 on Tuesday, and so on.', 'Step back that many days less one: a Wednesday (3) steps back 2 days.', 'The shape is =C5-WEEKDAY(C5,2)+1.'],
          explain: 'WEEKDAY(C5,2)-1 is how many days have passed since Monday, so subtracting it lands on that Monday. WEEKDAY(C5,3) counts Monday as 0, which does the same in one step.',
          chips: ['C5-', 'WEEKDAY(C5,2)', '+1'], fnKey: 'WEEKDAY(' },
        { title: 'Collection on the first Friday on or after delivery', ansLabel: 'Collect on', ansDate: true,
          statement: '<p>Collectors visit every Friday. Payment for a delivery is collected on the first Friday on or after the delivery <b>Date</b>: a Friday delivery is collected the same day, a Saturday delivery six days later. In <b>E5</b>, return the collection date and fill down to <b>' + last + '</b>.</p>',
          solution: '=C5+MOD(5-WEEKDAY(C5,2),7)', alts: ['=C5+MOD(6-WEEKDAY(C5),7)', '=C5+CHOOSE(WEEKDAY(C5,2),4,3,2,1,0,6,5)'],
          hints: ['WEEKDAY(C5,2) gives Friday as 5, so the days still to go are 5 minus the weekday.', 'After Friday that difference is negative; MOD(..., 7) wraps it into 0 to 6.', 'The shape is =C5+MOD(5-WEEKDAY(C5,2),7).'],
          explain: 'MOD(5-WEEKDAY(C5,2),7) is the number of days until the next Friday: 0 on a Friday, 6 on a Saturday. Adding it to the date gives the collection day.',
          chips: ['C5+', 'WEEKDAY(C5,2)', ',7)'], fnKey: 'MOD(' }
      ][v];
      return H.problem(Object.assign({ file: 'pengiriman_nov_des.xlsx', sheet: S, alt, fill: true }, V));
    }
  });

  P.add({
    key: 'month-end-and-anniversary', difficulty: 'Medium', level: 6, topics: ['EDATE', 'EOMONTH'], count: 3,
    make(r, v) {
      const rows = span(1, 12).map(() => {
        const y = r.pick([2024, 2025, 2025, 2026]), m = y === 2026 ? r.int(1, 10) : r.int(1, 12);
        return { start: ser(y, m, r.int(1, 28)), dist: r.pick(DIST()).name, tenor: r.pick([12, 18, 24, 36, 48]), limit: r.pick([50, 75, 100, 150, 200]) * 1000000 };
      }).sort((a, b) => a.start - b.start);
      refs(r, 'F', 12, 300).forEach((x, i) => { rows[i].ref = x; });
      const S = H.table([{ h: 'Facility', k: 'ref' }, C_DIST, { h: 'Start', k: 'start', date: true }, { h: 'Tenor (months)', k: 'tenor', num: true }, { h: 'Limit (Rp)', k: 'limit', num: true }], rows);
      /* The hidden sheet: two start dates corrected and one tenor changed. */
      const alt = H.edit(S.cells, { C6: (x) => x + 40, D9: (t) => (t === 24 ? 36 : 24), C12: (x) => x + 17 });
      const last = 'F' + S.last;
      const V = [
        { title: 'Maturity date from a start date and tenor in months', ansLabel: 'Maturity', ansDate: true,
          statement: '<p>Each Bank Kenari facility runs for its <b>Tenor</b> in months from its <b>Start</b> date and matures on the same day of the month that many months later. In <b>F5</b>, return the maturity date and fill down to <b>' + last + '</b>.</p>',
          solution: '=EDATE(C5,D5)', alts: ['=DATE(YEAR(C5),MONTH(C5)+D5,DAY(C5))'],
          hints: ['Adding days does not work here, because months have different lengths.', 'EDATE(start, months) moves a date by whole months.', 'The shape is =EDATE(C5,D5).'],
          explain: 'EDATE moves the start date forward D5 calendar months and keeps the day of the month. DATE(YEAR(C5),MONTH(C5)+D5,DAY(C5)) does the same, because DATE rolls month 14 into February of the next year.',
          chips: ['C5', 'D5'], fnKey: 'EDATE(' },
        { title: 'Annual review on the month end a year after start', ansLabel: 'Review date', ansDate: true,
          statement: '<p>Bank Kenari reviews each facility on the last day of the month one year after its <b>Start</b> month. A facility that started 9 Mar 2025 is reviewed 31 Mar 2026. In <b>F5</b>, return each review date and fill down to <b>' + last + '</b>.</p>',
          solution: '=EOMONTH(C5,12)', alts: ['=EOMONTH(EDATE(C5,12),0)', '=DATE(YEAR(C5)+1,MONTH(C5)+1,0)'],
          hints: ['EOMONTH(date, months) returns a month end, counted from the month of the date.', 'One year is 12 months; the tenor does not matter here.', 'The shape is =EOMONTH(C5,12).'],
          explain: 'EOMONTH(C5,12) moves twelve months on from the start month and returns that month\'s last day, whatever the start day.',
          chips: ['C5', ',12)'], fnKey: 'EOMONTH(' },
        { title: 'Interest days in the start month, both ends counted', ansLabel: 'First-month days',
          statement: '<p>Interest for the first month runs from the <b>Start</b> date to the last day of that month, counting both days. A facility that starts 28 Feb 2025 has 1 day; one that starts 1 Apr 2025 has 30. In <b>F5</b>, return the interest days in the start month and fill down to <b>' + last + '</b>.</p>',
          solution: '=EOMONTH(C5,0)-C5+1', alts: ['=DAY(EOMONTH(C5,0))-DAY(C5)+1', '=EOMONTH(C5,0)+1-C5'],
          hints: ['EOMONTH(C5,0) is the last day of the start month.', 'One date minus another counts the days between them; add 1 to count both ends.', 'The shape is =EOMONTH(C5,0)-C5+1.'],
          explain: 'Subtracting the start date from the month end counts the days after the start, and +1 adds the start day itself.',
          chips: ['EOMONTH(C5,0)', '-C5', '+1'], fnKey: 'EOMONTH(' }
      ][v];
      return H.problem(Object.assign({ file: 'fasilitas_bank_kenari.xlsx', sheet: S, alt, fill: true }, V));
    }
  });

  P.add({
    key: 'working-days-between', difficulty: 'Medium', level: 5, topics: ['NETWORKDAYS', 'NETWORKDAYS.INTL'], count: 3,
    make(r, v) {
      const n = 12;
      const hol = v === 1 ? holidayRows(r, 5) : [];
      const rows = span(0, n - 1).map((i) => {
        if (v === 1 && i < 5) { const h = hol[i].date; return { ordered: h - r.int(1, 5), delivered: h + r.int(0, 6), dist: r.pick(DIST()).name }; }
        const s = ser(2026, 11, 16) + r.int(0, 52);
        return { ordered: s, delivered: s + r.int(2, 12), dist: r.pick(DIST()).name };
      }).sort((a, b) => a.ordered - b.ordered);
      refs(r, 'SO-', n, 6100).forEach((x, i) => { rows[i].ref = x; });
      const S = H.table([{ h: 'Order', k: 'ref' }, C_DIST, { h: 'Ordered', k: 'ordered', date: true }, { h: 'Delivered', k: 'delivered', date: true }], rows);
      const T = v === 1 ? holidayTable(hol, 6) : null;
      const sheet = T ? beside(S, T) : S;
      const base = Object.assign({}, S.cells, T ? T.cells : {});
      /* The hidden sheet: one delivery came three days later, one order two days earlier. */
      const alt = H.edit(base, { D6: (x) => x + 3, C11: (x) => x - 2 });
      const last = 'E' + S.last;
      const V = [
        { title: 'Working days from order to delivery',
          statement: '<p>Count the working days (Monday to Friday) from the <b>Ordered</b> date to the <b>Delivered</b> date, counting both days. Type it in <b>E5</b> and fill down to <b>' + last + '</b>.</p>',
          solution: '=NETWORKDAYS(C5,D5)', alts: ['=NETWORKDAYS.INTL(C5,D5)', '=NETWORKDAYS.INTL(C5,D5,1)'],
          hints: ['D5-C5 would count Saturdays and Sundays too.', 'NETWORKDAYS(start, end) counts Monday to Friday, both ends included.', 'The shape is =NETWORKDAYS(C5,D5).'],
          explain: 'NETWORKDAYS counts the days from C5 to D5 that fall Monday to Friday, both ends included, and leaves out the weekends in between.',
          chips: ['C5', 'D5'], fnKey: 'NETWORKDAYS(' },
        { title: 'Working days to delivery, skipping company holidays',
          statement: '<p>Count the working days from the <b>Ordered</b> date to the <b>Delivered</b> date, both days included: Monday to Friday, leaving out the company holidays listed in <b>G5:G9</b>. Type it in <b>E5</b> and fill down to <b>' + last + '</b>.</p>',
          solution: '=NETWORKDAYS(C5,D5,$G$5:$G$9)', alts: ['=NETWORKDAYS.INTL(C5,D5,1,$G$5:$G$9)'],
          hints: ['NETWORKDAYS takes a third argument: a range of holiday dates to skip.', 'Pin the holiday range with $ so it stays in place as you fill down.', 'The shape is =NETWORKDAYS(C5,D5,$G$5:$G$9).'],
          explain: 'The third argument removes any listed holiday that falls on a weekday between the two dates. The $ signs keep every row pointing at the same holiday list.',
          chips: ['C5', 'D5', '$G$5:$G$9'], fnKey: 'NETWORKDAYS(' },
        { title: 'Working days to delivery on a Monday to Saturday week',
          statement: '<p>Distributors in this region work Monday to Saturday, so only Sunday is a day off. Count the working days from <b>Ordered</b> to <b>Delivered</b>, both days included, in <b>E5</b> and fill down to <b>' + last + '</b>.</p>',
          solution: '=NETWORKDAYS.INTL(C5,D5,11)', alts: ['=NETWORKDAYS.INTL(C5,D5,"0000001")'],
          hints: ['NETWORKDAYS always takes Saturday and Sunday off. NETWORKDAYS.INTL lets you choose the weekend.', 'Weekend code 11 means Sunday only. The text "0000001" (Monday to Sunday, 1 is a day off) means the same.', 'The shape is =NETWORKDAYS.INTL(C5,D5,11).'],
          explain: 'NETWORKDAYS.INTL with weekend 11 treats only Sunday as a day off, so Saturdays count as working days.',
          chips: ['C5', 'D5', ',11)'], fnKey: 'NETWORKDAYS.INTL(' }
      ][v];
      return H.problem(Object.assign({ file: 'pesanan_q4.xlsx', sheet, base, alt, fill: true, fillCol: 'E', ansLabel: 'Working days' }, V));
    }
  });

  P.add({
    key: 'ship-date-workday', difficulty: 'Medium', level: 5, topics: ['WORKDAY'], count: 3,
    make(r, v) {
      const n = 12;
      if (v < 2) {
        const hol = v === 1 ? holidayRows(r, 5) : [];
        const rows = span(0, n - 1).map((i) => {
          if (v === 1 && i < 5) return { ordered: hol[i].date - r.int(1, 4), lead: r.pick([3, 5, 7]), dist: r.pick(DIST()).name };
          return { ordered: ser(2026, 11, 23) + r.int(0, 37), lead: r.pick([2, 3, 5, 7, 10]), dist: r.pick(DIST()).name };
        }).sort((a, b) => a.ordered - b.ordered);
        refs(r, 'SO-', n, 6400).forEach((x, i) => { rows[i].ref = x; });
        const S = H.table([{ h: 'Order', k: 'ref' }, C_DIST, { h: 'Ordered', k: 'ordered', date: true }, { h: 'Lead (days)', k: 'lead', num: true }], rows);
        const T = v === 1 ? holidayTable(hol, 6) : null;
        const base = Object.assign({}, S.cells, T ? T.cells : {});
        /* The hidden sheet: one lead time grew, one order is a week later. */
        const alt = H.edit(base, { D7: (x) => x + 2, C10: (x) => x + 7 });
        const V = [
          { title: 'Ship date after a lead time in working days',
            statement: '<p>Each order ships <b>Lead</b> working days (Monday to Friday) after its <b>Ordered</b> date. The order day itself does not count: an order placed on a Friday with a lead of 3 ships on Wednesday. In <b>E5</b>, return the ship date and fill down to <b>E' + S.last + '</b>.</p>',
            solution: '=WORKDAY(C5,D5)', alts: ['=WORKDAY.INTL(C5,D5,1)'],
            hints: ['C5+D5 would count Saturdays and Sundays as working days.', 'WORKDAY(start, days) steps forward that many working days, skipping weekends.', 'The shape is =WORKDAY(C5,D5).'],
            explain: 'WORKDAY starts counting on the day after C5 and skips Saturdays and Sundays, so a 5-day lead from a Friday ships the next Friday.',
            chips: ['C5', 'D5'], fnKey: 'WORKDAY(' },
          { title: 'Ship date after lead time, skipping company holidays',
            statement: '<p>Each order ships <b>Lead</b> working days after its <b>Ordered</b> date, and the order day itself does not count. Working days are Monday to Friday, without the company holidays in <b>G5:G9</b>: with no holiday in the way, an order placed on a Friday with a lead of 3 ships on Wednesday. In <b>E5</b>, return the ship date and fill down to <b>E' + S.last + '</b>.</p>',
            solution: '=WORKDAY(C5,D5,$G$5:$G$9)', alts: ['=WORKDAY.INTL(C5,D5,1,$G$5:$G$9)'],
            hints: ['WORKDAY takes a third argument: a range of holiday dates to skip.', 'Pin the holiday range with $ so it does not slide as you fill down.', 'The shape is =WORKDAY(C5,D5,$G$5:$G$9).'],
            explain: 'WORKDAY skips weekends and every date in the holiday list, so an order placed before Christmas ships later than the lead time alone suggests.',
            chips: ['C5', 'D5', '$G$5:$G$9'], fnKey: 'WORKDAY(' }
        ][v];
        return H.problem(Object.assign({ file: 'pesanan_q4.xlsx', sheet: T ? beside(S, T) : S, base, alt, fill: true, fillCol: 'E', ansLabel: 'Ship date', ansDate: true }, V));
      }
      const rows = span(1, n).map(() => ({ due: ser(2026, 11, 2) + r.int(0, 58), dist: r.pick(DIST()).name, amount: amount(r) })).sort((a, b) => a.due - b.due);
      refs(r, 'INV-', n, 4500).forEach((x, i) => { rows[i].ref = x; });
      const S = H.table([C_INV, C_DIST, C_DUE, C_AMT], rows);
      const days = r.pick([2, 3]);
      const base = Object.assign({}, S.cells, { A1: 'Working days before', B1: days });
      /* The hidden sheet: reminders go out two working days earlier, and one due date moves a week. */
      const alt = H.edit(base, { B1: days + 2, C8: (x) => x + 7 });
      return H.problem({
        title: 'Reminder a set number of working days before due', file: 'piutang_q4.xlsx', sheet: S, base, alt, fill: true, ansLabel: 'Reminder', ansDate: true,
        statement: '<p>Credit Control sends a payment reminder a number of working days before each <b>Due date</b>; the number is in <b>B1</b>. Count back over Monday to Friday only: with 2 in B1, a Wednesday due date gets its reminder on Monday, and a Monday due date on the Thursday before. In <b>E5</b>, return the reminder date and fill down to <b>E' + S.last + '</b>.</p>',
        solution: '=WORKDAY(C5,-$B$1)', alts: ['=WORKDAY.INTL(C5,-$B$1,1)', '=WORKDAY(C5,-1*$B$1)'],
        hints: ['WORKDAY also counts backward when the number of days is negative.', 'Use -$B$1 so every row reads the same cell.', 'The shape is =WORKDAY(C5,-$B$1).'],
        explain: 'A negative day count makes WORKDAY step back over working days, skipping weekends. The hidden sheet changes B1, so the formula must point at it.',
        chips: ['C5', '-$B$1'], fnKey: 'WORKDAY('
      });
    }
  });

  P.add({
    key: 'tenure-in-years', difficulty: 'Medium', level: 7, topics: ['DATEDIF', 'YEAR'], count: 3,
    make(r, v) {
      const asof = r.pick([ser(2026, 10, 31), ser(2026, 12, 31), ser(2027, 3, 31)]), asofAlt = asof + r.int(120, 200);
      const rows = H.reps(r, 12).map((x) => ({ name: x.name, area: x.area, start: H.toSerial(x.start) }));
      /* One rep's sixth anniversary is the day after the as-of date: dividing the days by 365 counts 6, the right answer is 5. */
      const t = H.iso(asof + 1), k1 = r.int(0, rows.length - 1);
      rows[k1].start = ser(+t.slice(0, 4) - 6, +t.slice(5, 7), +t.slice(8, 10));
      /* Another rep's second anniversary is the as-of date itself, with no 29 February in between: 730 days over
         365.25 counts 1, the right answer is 2. */
      const a = H.iso(asof);
      rows[(k1 + 1 + r.int(0, rows.length - 2)) % rows.length].start = ser(+a.slice(0, 4) - 2, +a.slice(5, 7), +a.slice(8, 10));
      rows.sort((a, b) => a.start - b.start);
      const S = H.table([{ h: 'Rep', k: 'name' }, { h: 'Area', k: 'area' }, { h: 'Start date', k: 'start', date: true }], rows);
      const base = Object.assign({}, S.cells, { A1: 'As of', B1: asof });
      /* The hidden sheet: a later as-of date, and one start date corrected. */
      const alt = H.edit(base, { B1: asofAlt, C8: (x) => x + 100 });
      const note = '<p>Start dates run from 2019 to 2025 and the as-of date is ' + long(asof) + '; the grid shows day and month only.</p>';
      const last = 'D' + S.last;
      const V = [
        { title: 'Tenure in whole years as of a date', ansLabel: 'Years',
          statement: '<p>In <b>D5</b>, return each rep\'s tenure in whole years: the number of full years from the <b>Start date</b> to the as-of date in <b>B1</b>. A rep who started 15 Mar 2022 has 4 years on 15 Mar 2026 and still 3 the day before. Fill down to <b>' + last + '</b>.</p>' + note,
          solution: '=DATEDIF(C5,$B$1,"y")', alts: ['=YEAR($B$1)-YEAR(C5)-(DATE(YEAR($B$1),MONTH(C5),DAY(C5))>$B$1)'],
          hints: ['Dividing the days by 365 goes wrong near anniversaries and leap years.', 'DATEDIF(start, end, "y") counts complete years.', 'The shape is =DATEDIF(C5,$B$1,"y").'],
          explain: 'DATEDIF with "y" counts only full years, so the number goes up on the anniversary itself. Excel leaves DATEDIF out of its function menu, but it works in every version.',
          chips: ['C5', '$B$1', '"y")'], fnKey: 'DATEDIF(' },
        { title: 'Tenure in whole months as of a date', ansLabel: 'Months',
          statement: '<p>In <b>D5</b>, return each rep\'s tenure in whole months from the <b>Start date</b> to the as-of date in <b>B1</b>. From 15 Mar 2026, 14 Apr 2026 is still 0 months and 15 Apr 2026 is 1. Fill down to <b>' + last + '</b>.</p>' + note,
          solution: '=DATEDIF(C5,$B$1,"m")', alts: ['=(YEAR($B$1)-YEAR(C5))*12+MONTH($B$1)-MONTH(C5)-(DAY($B$1)<DAY(C5))'],
          hints: ['A whole month is complete only when the day of the month is reached again.', 'DATEDIF(start, end, "m") counts complete months.', 'The shape is =DATEDIF(C5,$B$1,"m").'],
          explain: 'DATEDIF with "m" counts the full months between the dates. The long form counts calendar months and takes one away when the as-of day has not reached the start day.',
          chips: ['C5', '$B$1', '"m")'], fnKey: 'DATEDIF(' },
        { title: 'Tenure written as years and months, like 3y 4m', ansLabel: 'Tenure', exactCase: true,
          statement: '<p>The HR card shows tenure as whole years and the whole months left over, in the form <b>3y 4m</b>: the years, a y, one space, the months, an m. Work it out from the <b>Start date</b> to the as-of date in <b>B1</b>, in <b>D5</b>, and fill down to <b>' + last + '</b>.</p>' + note,
          solution: '=DATEDIF(C5,$B$1,"y")&"y "&DATEDIF(C5,$B$1,"ym")&"m"', alts: ['=INT(DATEDIF(C5,$B$1,"m")/12)&"y "&MOD(DATEDIF(C5,$B$1,"m"),12)&"m"'],
          hints: ['DATEDIF with "y" gives the full years.', 'DATEDIF with "ym" gives the months left over after the full years.', 'Join the pieces with &: years, "y ", months, "m".'],
          explain: 'The "y" unit counts full years and the "ym" unit counts the months past the last anniversary. Joining them with the letters gives the card format.',
          chips: ['DATEDIF(C5,$B$1,"y")', 'DATEDIF(C5,$B$1,"ym")', '&"y "&', '&"m"'], fnKey: 'DATEDIF(' }
      ][v];
      return H.problem(Object.assign({ file: 'tim_sales.xlsx', sheet: S, base, alt, fill: true, dateCells: ['B1'] }, V));
    }
  });

  const BUCKETS = [{ from: 0, name: 'Current' }, { from: 1, name: '1-30' }, { from: 31, name: '31-60' }, { from: 61, name: '61-90' }, { from: 91, name: 'Over 90' }];
  const IFS_C = 'IFS($B$1-C5<=0,"Current",$B$1-C5<=30,"1-30",$B$1-C5<=60,"31-60",$B$1-C5<=90,"61-90",TRUE,"Over 90")';
  const NEST_C = 'IF($B$1-C5<=0,"Current",IF($B$1-C5<=30,"1-30",IF($B$1-C5<=60,"31-60",IF($B$1-C5<=90,"61-90","Over 90"))))';
  const BUCKET_RULE = '<b>Current</b> for 0 days or less, <b>1-30</b>, <b>31-60</b>, <b>61-90</b>, and <b>Over 90</b> for 91 days or more';

  P.add({
    key: 'aging-bucket-label', difficulty: 'Medium', level: 7, topics: ['IFS', 'XLOOKUP', 'IF'], count: 3,
    make(r, v) {
      const asof = r.pick(ASOF), asofAlt = asof + r.int(8, 15);
      /* The edges of every bucket are in the sheet, so a < where <= belongs shows up. */
      const offs = r.shuffle([0, 30, 31, 60, 61, 90, 91].concat(r.sample([-20, -6, 1, 9, 18, 44, 58, 77, 105, 140], 7)));
      const rows = aged(r, asof, offs, { paid: v === 2 ? 0.35 : 0, sortBy: v === 1 ? 'date' : 'due' });
      const top = { A1: 'As of', B1: asof };
      const note = ' The as-of date in B1 is ' + long(asof) + '.';
      if (v === 1) {
        const S = H.table([C_INV, C_DIST, C_DATE, C_TERMS, C_AMT], rows);
        const T = H.table([{ h: 'From (days)', k: 'from', num: true }, { h: 'Bucket', k: 'name' }], BUCKETS, { startCol: 7 });
        const base = Object.assign({}, S.cells, T.cells, top);
        const alt = H.edit(base, { B1: asofAlt, D9: (t) => (t === 30 ? 45 : 30) });
        return H.problem({
          title: 'Aging bucket from terms with a bucket table', file: 'piutang_q4.xlsx', sheet: beside(S, T), base, alt, fill: true, fillCol: 'F', ansLabel: 'Bucket', exactCase: true, dateCells: ['B1'],
          statement: '<p>The due date is the invoice <b>Date</b> plus its <b>Terms</b>. Days past due as of <b>B1</b> are B1 minus the due date, or 0 when not yet due. The table in <b>H5:I9</b> lists the day each bucket starts from. In <b>F5</b>, return the bucket name from the table and fill down to <b>F' + S.last + '</b>.' + note + '</p>',
          solution: '=XLOOKUP(MAX(0,$B$1-(C5+D5)),$H$5:$H$9,$I$5:$I$9,,-1)',
          alts: ['=VLOOKUP(MAX(0,$B$1-C5-D5),$H$5:$I$9,2,TRUE)', '=INDEX($I$5:$I$9,MATCH(MAX(0,$B$1-C5-D5),$H$5:$H$9,1))', '=LOOKUP(MAX(0,$B$1-C5-D5),$H$5:$H$9,$I$5:$I$9)'],
          hints: ['Days past due: MAX(0,$B$1-(C5+D5)).', 'An approximate lookup finds the largest "From" value that is not above the days: XLOOKUP with match mode -1, or VLOOKUP with TRUE.', 'The shape is =XLOOKUP(days,$H$5:$H$9,$I$5:$I$9,,-1).'],
          explain: 'XLOOKUP with match mode -1 returns the bucket whose start is the largest one at or below the days past due. With the limits in a table, a new bucket rule needs no new formula.',
          chips: ['MAX(0,$B$1-(C5+D5))', '$H$5:$H$9', '$I$5:$I$9', ',,-1)'], fnKey: 'XLOOKUP('
        });
      }
      if (v === 0) {
        const S = H.table([C_INV, C_DIST, C_DUE, C_AMT], rows);
        const base = Object.assign({}, S.cells, top);
        const alt = H.edit(base, { B1: asofAlt });
        return H.problem({
          title: 'Aging bucket for each invoice as of a date', file: 'piutang_q4.xlsx', sheet: S, base, alt, fill: true, ansLabel: 'Bucket', exactCase: true, dateCells: ['B1'],
          statement: '<p>Sort each invoice into an aging bucket by its days past due as of <b>B1</b> (B1 minus the <b>Due date</b>): ' + BUCKET_RULE + '. Write the labels exactly. Type it in <b>E5</b> and fill down to <b>E' + S.last + '</b>.' + note + '</p>',
          solution: '=' + IFS_C, alts: ['=' + NEST_C, '=LOOKUP(MAX(0,$B$1-C5),{0,1,31,61,91},{"Current","1-30","31-60","61-90","Over 90"})'],
          hints: ['Work out the days past due first: $B$1-C5.', 'IFS returns the result of the first test that is true, so test the smallest bucket first and end with TRUE for Over 90.', 'The shape is =IFS($B$1-C5<=0,"Current",$B$1-C5<=30,"1-30", ... ,TRUE,"Over 90").'],
          explain: 'IFS stops at the first true test, so each later test only needs the upper limit of its bucket. TRUE at the end catches everything over 90 days.',
          chips: ['$B$1-C5', '<=0,"Current"', '<=30,"1-30"', '<=60,"31-60"', '<=90,"61-90"', 'TRUE,"Over 90")'], fnKey: 'IFS('
        });
      }
      const S = H.table([C_INV, C_DIST, C_DUE, C_AMT, C_PAID], rows);
      const base = Object.assign({}, S.cells, top);
      const alt = H.edit(payNews(base, S, rows, asof - 1), { B1: asofAlt });
      return H.problem({
        title: 'Aging bucket that labels paid invoices Paid', file: 'piutang_q4.xlsx', sheet: S, base, alt, fill: true, ansLabel: 'Bucket', exactCase: true, dateCells: ['B1'],
        statement: '<p>An invoice with a date in <b>Paid on</b> gets the label <b>Paid</b>. Every other invoice goes into an aging bucket by its days past due as of <b>B1</b> (B1 minus the <b>Due date</b>): ' + BUCKET_RULE + '. Type it in <b>F5</b> and fill down to <b>F' + S.last + '</b>.' + note + '</p>',
        solution: '=IF(E5<>"","Paid",' + IFS_C + ')', alts: ['=IF(E5="",' + IFS_C + ',"Paid")', '=IF(E5<>"","Paid",' + NEST_C + ')'],
        hints: ['Check the payment first: E5<>"" means the invoice was paid.', 'For open invoices, IFS tests the days past due from the smallest bucket up and ends with TRUE.', 'The shape is =IF(E5<>"","Paid",IFS($B$1-C5<=0,"Current", ... ,TRUE,"Over 90")).'],
        explain: 'The outer IF settles paid invoices before any aging happens, so only open invoices reach the IFS that picks the bucket.',
        chips: ['E5<>""', '"Paid"', '$B$1-C5', 'TRUE,"Over 90")'], fnKey: 'IF('
      });
    }
  });

  /* ======================= Hard ======================= */

  P.add({
    key: 'roll-due-to-working-day', difficulty: 'Hard', level: 8, topics: ['WORKDAY', 'WEEKDAY'], count: 3,
    make(r, v) {
      const withHol = v > 0;
      const hol = withHol ? holidayRows(r, 5) : [];
      const hs = hol.map((h) => h.date);
      const days = span(ser(2026, 11, 23), ser(2027, 1, 8));
      /* Due dates on purpose: three Saturdays, three Sundays, some holidays and some plain weekdays. */
      const sat = r.sample(days.filter((s) => wk(s) === 6), 3), sun = r.sample(days.filter((s) => wk(s) === 7), 3);
      const hd = withHol ? r.sample(hs, 3) : [];
      const plain = r.sample(days.filter((s) => wk(s) < 6 && !hs.includes(s)), 12 - 6 - hd.length);
      const rows = sat.concat(sun, hd, plain).map((due) => { const terms = r.pick([14, 30, 45, 60]); return { due, terms, date: due - terms, dist: r.pick(DIST()).name, amount: amount(r) }; }).sort((a, b) => a.date - b.date);
      refs(r, 'INV-', rows.length, 5300).forEach((x, i) => { rows[i].ref = x; });
      const S = H.table([C_INV, C_DIST, C_DATE, C_TERMS, C_AMT], rows);
      const T = withHol ? holidayTable(hol, 7) : null;
      const base = Object.assign({}, S.cells, T ? T.cells : {});
      /* The hidden sheet: one invoice a week earlier, one with new terms. */
      const alt = H.edit(base, { C5: (x) => x - 7, D10: (t) => (t === 30 ? 45 : 30) });
      const last = 'F' + S.last;
      const V = [
        { title: 'Due date moved off a weekend to the next Monday',
          statement: '<p>The due date is the invoice <b>Date</b> plus its <b>Terms</b> in days. When it lands on a Saturday or Sunday, payment is due the next Monday instead; a weekday due date stays. In <b>F5</b>, return the due date after this rule and fill down to <b>' + last + '</b>.</p>',
          solution: '=WORKDAY(C5+D5-1,1)', alts: ['=WORKDAY.INTL(C5+D5-1,1,1)', '=C5+D5+IF(WEEKDAY(C5+D5,2)>5,8-WEEKDAY(C5+D5,2),0)'],
          hints: ['WORKDAY(start, 1) returns the first working day after start.', 'Start one day early, at C5+D5-1: when the due date is a working day, WORKDAY returns it unchanged.', 'The shape is =WORKDAY(C5+D5-1,1).'],
          explain: 'WORKDAY(due-1, 1) looks for the first working day after the day before the due date. That is the due date itself on a weekday, and the next Monday on a weekend.',
          chips: ['C5+D5', '-1,1)'], fnKey: 'WORKDAY(' },
        { title: 'Due date moved past weekends and company holidays',
          statement: '<p>The due date is the invoice <b>Date</b> plus its <b>Terms</b> in days. When it lands on a Saturday, a Sunday or one of the company holidays in <b>H5:H9</b>, payment moves to the next working day; a working-day due date stays. In <b>F5</b>, return the due date after this rule and fill down to <b>' + last + '</b>.</p>',
          solution: '=WORKDAY(C5+D5-1,1,$H$5:$H$9)', alts: ['=WORKDAY.INTL(C5+D5-1,1,1,$H$5:$H$9)'],
          hints: ['WORKDAY(start, 1, holidays) returns the first working day after start, skipping weekends and holidays.', 'Start one day early, at C5+D5-1, so a due date that is already a working day comes back unchanged.', 'The shape is =WORKDAY(C5+D5-1,1,$H$5:$H$9).'],
          explain: 'Starting from the day before the due date, one working day forward is the due date when it is a working day. Otherwise WORKDAY walks past the weekend and any listed holidays, so 25 Dec moves to the next open day.',
          chips: ['C5+D5', '-1,1,', '$H$5:$H$9'], fnKey: 'WORKDAY(' },
        { title: 'Due date pulled back to the working day before',
          statement: '<p>The distributor\'s auto-debit runs early. The due date is the invoice <b>Date</b> plus its <b>Terms</b>; when it lands on a Saturday, a Sunday or a company holiday in <b>H5:H9</b>, the debit runs on the last working day before it. A working-day due date stays. In <b>F5</b>, return the debit date and fill down to <b>' + last + '</b>.</p>',
          solution: '=WORKDAY(C5+D5+1,-1,$H$5:$H$9)', alts: ['=WORKDAY.INTL(C5+D5+1,-1,1,$H$5:$H$9)'],
          hints: ['WORKDAY(start, -1) returns the working day before start.', 'Start one day late, at C5+D5+1, so a due date that is already a working day comes back unchanged.', 'The shape is =WORKDAY(C5+D5+1,-1,$H$5:$H$9).'],
          explain: 'One working day back from the day after the due date is the due date itself when it is open. A weekend or holiday due date is skipped backward to the last open day, so a Saturday becomes the Friday before unless that Friday is a holiday too.',
          chips: ['C5+D5', '+1,-1,', '$H$5:$H$9'], fnKey: 'WORKDAY(' }
      ][v];
      return H.problem(Object.assign({ file: 'piutang_q4.xlsx', sheet: T ? beside(S, T) : S, base, alt, fill: true, fillCol: 'F', ansLabel: v === 2 ? 'Debit date' : 'Pay by', ansDate: true }, V));
    }
  });

  P.add({
    key: 'working-days-in-month', difficulty: 'Hard', level: 8, topics: ['NETWORKDAYS', 'EOMONTH'], count: 3,
    make(r, v) {
      const rows = HOL_YEAR.map(([d, name]) => ({ date: H.toSerial(d), name }));
      const hs = rows.map((x) => x.date);
      const S = H.table([{ h: 'Holiday', k: 'date', date: true }, { h: 'Name', k: 'name' }], rows);
      const months = [[2026, 11], [2026, 12], [2027, 1], [2027, 2], [2027, 3]];
      const left = v === 1;
      const count = (s) => (left ? nwd(s + 1, eom(s, 0), hs) : nwd(eom(s, -1) + 1, eom(s, 0), hs));
      const at = (ym) => ser(ym[0], ym[1], r.int(5, 24));
      let day, dayAlt;
      if (left) {
        day = at(r.pick(months.slice(1, 2).concat(months.slice(0, 1))));
        /* The hidden sheet asks about another date, chosen so the count is different. */
        dayAlt = day;
        for (const ym of r.shuffle(months)) { const s = at(ym); if (count(s) !== count(day) && Math.abs(s - day) > 20) { dayAlt = s; break; } }
      } else {
        /* One sheet asks about December: the day before 1 December, 30 November 2026, is a Monday, so a span that
           starts at EOMONTH(B1,-1) instead of EOMONTH(B1,-1)+1 counts a day too many there. The other sheet asks
           about November or March, which have more working days than December. */
        const pair = r.shuffle([[2026, 12], r.pick([[2026, 11], [2027, 3]])]);
        day = at(pair[0]); dayAlt = at(pair[1]);
      }
      const hol = 'A5:A' + S.last;
      const range = 'NETWORKDAYS(EOMONTH(B1,-1)+1,EOMONTH(B1,0),' + hol + ')';
      if (v === 2) {
        const target = r.pick([12, 15, 18, 24]) * 100000000;
        const base = Object.assign({}, S.cells, { A1: 'Month of', B1: day, A2: 'Monthly target (Rp)', B2: target, A3: 'Daily target (Rp)' });
        const alt = H.edit(base, { B1: dayAlt, B2: target + r.int(1, 3) * 100000000 });
        return H.problem({
          title: 'Daily sales target over the working days of a month', file: 'kalender_kerja_2027.xlsx', sheet: S, base, alt, answerCell: 'B3', dateCells: ['B1'],
          statement: '<p>The monthly sales target in <b>B2</b> is spread evenly over the working days of the month that holds the date in <b>B1</b>: Monday to Friday, less the holidays in <b>' + hol + '</b>. In <b>B3</b>, return the daily target rounded to whole rupiah. B1 is ' + long(day) + '.</p>',
          solution: '=ROUND(B2/' + range + ',0)', alts: ['=ROUND(B2/NETWORKDAYS(DATE(YEAR(B1),MONTH(B1),1),EOMONTH(B1,0),' + hol + '),0)'],
          hints: ['The month runs from EOMONTH(B1,-1)+1 to EOMONTH(B1,0).', 'NETWORKDAYS over that span, with the holiday range, gives the working days to divide by.', 'The shape is =ROUND(B2/NETWORKDAYS(EOMONTH(B1,-1)+1,EOMONTH(B1,0),' + hol + '),0).'],
          explain: 'EOMONTH finds both ends of the month from any date in it, NETWORKDAYS counts the open days between them, and ROUND keeps the daily target in whole rupiah.',
          chips: ['B2/', 'EOMONTH(B1,-1)+1', 'EOMONTH(B1,0)', hol, ',0)'], fnKey: 'NETWORKDAYS('
        });
      }
      const base = Object.assign({}, S.cells, { A1: left ? 'Planning date' : 'Month of', B1: day, A2: left ? 'Days left' : 'Working days' });
      const alt = H.edit(base, { B1: dayAlt });
      if (left) {
        return H.problem({
          title: 'Working days left in the month after a date', file: 'kalender_kerja_2027.xlsx', sheet: S, base, alt, answerCell: 'B2', dateCells: ['B1'],
          statement: '<p>Sales planning needs the working days still left in the month after the date in <b>B1</b>: from the day after B1 to the last day of its month, Monday to Friday, without the holidays in <b>' + hol + '</b>. Return the count in <b>B2</b>. B1 is ' + long(day) + '.</p>',
          solution: '=NETWORKDAYS(B1+1,EOMONTH(B1,0),' + hol + ')', alts: ['=NETWORKDAYS.INTL(B1+1,EOMONTH(B1,0),1,' + hol + ')'],
          hints: ['The span starts the day after B1, which is B1+1.', 'It ends on the last day of the month: EOMONTH(B1,0).', 'The shape is =NETWORKDAYS(B1+1,EOMONTH(B1,0),' + hol + ').'],
          explain: 'B1+1 leaves the planning day itself out, EOMONTH(B1,0) closes the month, and NETWORKDAYS counts the weekdays between that are not on the holiday list.',
          chips: ['B1+1', 'EOMONTH(B1,0)', hol], fnKey: 'NETWORKDAYS('
        });
      }
      return H.problem({
        title: 'Working days in the month of a date, less holidays', file: 'kalender_kerja_2027.xlsx', sheet: S, base, alt, answerCell: 'B2', dateCells: ['B1'],
        statement: '<p>In <b>B2</b>, count the working days (Monday to Friday) in the whole month that holds the date in <b>B1</b>, leaving out the holidays listed in <b>' + hol + '</b>. B1 is ' + long(day) + '.</p>',
        solution: '=' + range, alts: ['=NETWORKDAYS(DATE(YEAR(B1),MONTH(B1),1),EOMONTH(B1,0),' + hol + ')', '=NETWORKDAYS(EOMONTH(B1,-1)+1,EOMONTH(B1,0),$A$5:$A$' + S.last + ')'],
        hints: ['The month runs from its first day, EOMONTH(B1,-1)+1, to its last day, EOMONTH(B1,0).', 'NETWORKDAYS(start, end, holidays) counts Monday to Friday and skips the holiday dates.', 'The shape is =NETWORKDAYS(EOMONTH(B1,-1)+1,EOMONTH(B1,0),' + hol + ').'],
        explain: 'EOMONTH(B1,-1)+1 is the day after last month ended, the first of this month, and EOMONTH(B1,0) is this month\'s last day. NETWORKDAYS counts the weekdays between them and removes the listed holidays.',
        chips: ['EOMONTH(B1,-1)+1', 'EOMONTH(B1,0)', hol], fnKey: 'NETWORKDAYS('
      });
    }
  });

  P.add({
    key: 'last-order-gap', difficulty: 'Hard', level: 8, topics: ['MAXIFS', 'MINIFS', 'COUNTIF'], count: 3,
    make(r, v) {
      const ds = r.sample(DIST(), 5), counts = r.shuffle([5, 4, 4, 4, 3]);
      const days = span(ser(2026, 9, 1), ser(2026, 12, 18));
      const rows = [];
      ds.forEach((d, i) => r.sample(days, counts[i]).forEach((date) => rows.push({ date, dist: d.name, net: amount(r) })));
      rows.sort((a, b) => a.date - b.date);
      refs(r, 'INV-', rows.length, 2900).forEach((x, i) => { rows[i].ref = x; });
      const S = H.table([{ h: 'Date', k: 'date', date: true }, C_INV, C_DIST, { h: 'Net (Rp)', k: 'net', num: true }], rows);
      const who = ds[counts.indexOf(4)].name;
      const others = ds.filter((d, i) => counts[i] !== 4).map((d) => d.name);
      const whoAlt = r.pick(others);
      const A = 'A5:A' + S.last, C = 'C5:C' + S.last;
      const MX = 'MAXIFS(' + A + ',' + C + ',B1)', MN = 'MINIFS(' + A + ',' + C + ',B1)', N = 'COUNTIF(' + C + ',B1)';
      const FX = 'FILTER(' + A + ',' + C + '=B1)';
      /* The hidden sheet names another distributor, and one of its orders is re-dated. */
      const lastRow = 'A' + S.last;
      if (v === 0) {
        /* The as-of date falls between the distributor's last two orders, so the last one is booked ahead and a
           MAXIFS without the date test gives a negative count. */
        const mine = rows.filter((x) => x.dist === who).map((x) => x.date).sort((a, b) => a - b);
        let k = mine.length - 1;
        while (k > 0 && mine[k] - mine[k - 1] < 2) k--;
        const asof = k > 0 ? r.int(mine[k - 1] + 1, mine[k] - 1) : mine[0];
        /* The hidden sheet names another distributor and sets B2 on its second-to-last order date: that order
           counts, so a test of < where <= belongs gives a different answer, and its last order is booked ahead. */
        const theirs = rows.filter((x) => x.dist === whoAlt).map((x) => x.date).sort((a, b) => a - b);
        const who0 = whoAlt, asofAlt = theirs.length > 1 ? theirs[theirs.length - 2] : asof + r.int(5, 10);
        const base = Object.assign({}, S.cells, { A1: 'Distributor', B1: who, A2: 'As of', B2: asof, A3: 'Days since last order' });
        const alt = H.edit(base, { B1: who0, B2: asofAlt, [lastRow]: (x) => x - 2 });
        const MX2 = 'MAXIFS(' + A + ',' + C + ',B1,' + A + ',"<="&B2)';
        return H.problem({
          title: "Days since a distributor's last order, ignoring later bookings", file: 'penjualan_q4.xlsx', sheet: S, base, alt, answerCell: 'B3', dateCells: ['B2'],
          statement: '<p>In <b>B3</b>, return how many days have passed from the latest order of the distributor named in <b>B1</b> to the as-of date in <b>B2</b>. The orders are in rows 5 to ' + S.last + ', and some are booked ahead with a date after B2: leave those out. An order dated on B2 itself counts. B2 is ' + long(asof) + '.</p>',
          solution: '=B2-' + MX2,
          alts: ['=B2-MAX(FILTER(' + A + ',(' + C + '=B1)*(' + A + '<=B2)))', '=B2-MAXIFS($A$5:$A$' + S.last + ',$C$5:$C$' + S.last + ',$B$1,$A$5:$A$' + S.last + ',"<="&$B$2)', '=B2-MAXIFS(' + A + ',' + A + ',"<="&B2,' + C + ',B1)'],
          hints: ['The latest order counts only among that distributor\'s rows dated on or before B2.', 'MAXIFS takes more than one pair of range and test: the distributor in ' + C + ', and the date in ' + A + ' with "<="&B2.', 'The shape is =B2-MAXIFS(' + A + ',' + C + ',B1,' + A + ',"<="&B2).'],
          explain: 'MAXIFS keeps the rows that pass both tests, the distributor in B1 and a date on or before B2, and returns the newest of those dates. Subtracting it from B2 counts the days since, and orders booked for later dates never enter.',
          chips: [A, C, 'B1', '"<="&B2', 'B2-'], fnKey: 'MAXIFS('
        });
      }
      const base = Object.assign({}, S.cells, { A1: 'Distributor', B1: who, A2: v === 1 ? 'Average gap (days)' : 'Next order' });
      const alt = H.edit(base, { B1: whoAlt, [lastRow]: (x) => x - 2 });
      if (v === 1) {
        return H.problem({
          title: "Average days between a distributor's orders", file: 'penjualan_q4.xlsx', sheet: S, base, alt, answerCell: 'B2',
          statement: '<p>For the distributor in <b>B1</b>, the average gap between orders is the days from its first order to its last, divided by the number of gaps (its number of orders minus one). Return it in <b>B2</b>, unrounded.</p>',
          solution: '=(' + MX + '-' + MN + ')/(' + N + '-1)', alts: ['=(MAX(' + FX + ')-MIN(' + FX + '))/(' + N + '-1)', '=(' + MX + '-' + MN + ')/(COUNTIFS(' + C + ',B1)-1)'],
          hints: ['First order: MINIFS of the dates. Last order: MAXIFS of the dates.', 'With 4 orders there are 3 gaps, so divide by COUNTIF(' + C + ',B1)-1.', 'The shape is =(MAXIFS(...)-MINIFS(...))/(COUNTIF(...)-1).'],
          explain: 'The gaps between orders add up to last minus first, so dividing that span by the number of gaps gives the average gap without listing each one.',
          chips: [A, C, 'B1', '-1)'], fnKey: 'MAXIFS('
        });
      }
      return H.problem({
        title: 'Next expected order date for a distributor', file: 'penjualan_q4.xlsx', sheet: S, base, alt, answerCell: 'B2', ansDate: true,
        statement: '<p>Forecast the next order of the distributor in <b>B1</b>: its last order date plus its average gap between orders, rounded to whole days. The average gap is (last order minus first order) divided by (number of orders minus one). Return the date in <b>B2</b>.</p>',
        solution: '=' + MX + '+ROUND((' + MX + '-' + MN + ')/(' + N + '-1),0)',
        alts: ['=LET(f,' + MN + ',l,' + MX + ',l+ROUND((l-f)/(' + N + '-1),0))', '=MAX(' + FX + ')+ROUND((MAX(' + FX + ')-MIN(' + FX + '))/(' + N + '-1),0)'],
        hints: ['Find the last order (MAXIFS), the first order (MINIFS) and the number of orders (COUNTIF).', 'The average gap is (last - first)/(orders - 1); ROUND(..., 0) makes it whole days.', 'The shape is =MAXIFS(...)+ROUND((MAXIFS(...)-MINIFS(...))/(COUNTIF(...)-1),0).'],
        explain: 'The last order plus the typical gap is a simple forecast of the next one. LET can name the first and last dates once, which keeps the formula shorter.',
        chips: [A, C, 'B1', 'ROUND('], fnKey: 'MAXIFS('
      });
    }
  });

  P.add({
    key: 'aging-amount-as-of', difficulty: 'Hard', level: 9, topics: ['SUMIFS', 'SUMPRODUCT', 'FILTER'], count: 3,
    make(r, v) {
      const asof = r.pick(ASOF), asofAlt = asof + r.int(9, 14);
      /* Invoices exactly 30, 31, 60 and 61 days past due stay unpaid, so a > where >= belongs, or 30 where 31
         belongs, gives a different answer. */
      const EDGE = [30, 31, 60, 61];
      const offs = EDGE.concat(r.sample([-15, -4, 3, 10, 22, 29, 35, 44, 52, 58, 66, 79, 95, 120, 150], 12));
      const rows = aged(r, asof, offs, { paid: 0.35 });
      rows.forEach((x) => { if (EDGE.includes(x.off)) x.paid = ''; });
      const S = H.table([C_INV, C_DIST, C_DUE, C_AMT, C_PAID], rows);
      const L = S.last, Dr = 'D5:D' + L, Cr = 'C5:C' + L, Er = 'E5:E' + L, Ar = 'A5:A' + L;
      const note = ' B1 is ' + long(asof) + '.';
      if (v === 2) {
        const base = Object.assign({}, S.cells, { A1: 'As of', B1: asof });
        const alt = H.edit(payNews(base, S, rows, asof - 1), { B1: asofAlt });
        return H.problem({
          title: 'List unpaid invoices more than 60 days past due', file: 'piutang_q4.xlsx', sheet: S, base, alt, spill: { cols: 1, rows: rows.length }, ansLabel: 'Over 60 days', dateCells: ['B1'],
          statement: '<p>In <b>F5</b>, list the <b>Invoice</b> numbers of unpaid invoices (no date in <b>Paid on</b>) that are more than 60 days past due as of <b>B1</b>, in the order of the sheet. Days past due are B1 minus the <b>Due date</b>. One formula fills the whole list downward.' + note + '</p>',
          solution: '=FILTER(' + Ar + ',(B1-' + Cr + '>60)*(' + Er + '=""))', alts: ['=FILTER(' + Ar + ',(' + Cr + '<B1-60)*(' + Er + '=""))', '=FILTER(' + Ar + ',(' + Er + '="")*(B1-' + Cr + '>60))'],
          hints: ['FILTER(what, condition) returns every row where the condition is true.', 'Two conditions that must both hold are multiplied: (B1-' + Cr + '>60)*(' + Er + '="").', 'The shape is =FILTER(' + Ar + ',(B1-' + Cr + '>60)*(' + Er + '="")).'],
          explain: 'B1-' + Cr + ' gives the days past due of every row at once. Multiplying the two tests keeps only rows where both are true, and FILTER returns their invoice numbers.',
          chips: [Ar, Cr, Er, '>60', '=""'], fnKey: 'FILTER('
        });
      }
      const base = Object.assign({}, S.cells, { A1: 'As of', B1: asof, A2: v === 0 ? '31-60 days (Rp)' : 'Share over 30 days' });
      const alt = H.edit(payNews(base, S, rows, asof - 1), { B1: asofAlt });
      if (v === 0) {
        return H.problem({
          title: 'Unpaid amount 31 to 60 days past due', file: 'piutang_q4.xlsx', sheet: S, base, alt, answerCell: 'B2', dateCells: ['B1'],
          statement: '<p>In <b>B2</b>, total the <b>Amount</b> of unpaid invoices (no date in <b>Paid on</b>) that are 31 to 60 days past due as of <b>B1</b>, both limits included. Days past due are B1 minus the <b>Due date</b>.' + note + '</p>',
          solution: '=SUMIFS(' + Dr + ',' + Cr + ',">="&B1-60,' + Cr + ',"<="&B1-31,' + Er + ',"")',
          alts: ['=SUMPRODUCT((B1-' + Cr + '>=31)*(B1-' + Cr + '<=60)*(' + Er + '="")*' + Dr + ')', '=SUMIFS(' + Dr + ',' + Er + ',"",' + Cr + ',">="&B1-60,' + Cr + ',"<="&B1-31)'],
          hints: ['31 to 60 days past due means the due date is between B1-60 and B1-31.', 'SUMIFS takes a date limit as text joined to a value: ">="&B1-60. An empty Paid on cell matches the criterion "".', 'The shape is =SUMIFS(' + Dr + ',' + Cr + ',">="&B1-60,' + Cr + ',"<="&B1-31,' + Er + ',"").'],
          explain: 'Turning the day limits into due-date limits lets SUMIFS do the work: B1-60 is the oldest due date in the bucket and B1-31 the newest. The criterion "" keeps only rows with an empty Paid on cell.',
          chips: [Dr, Cr, Er, '">="&B1-60', '"<="&B1-31', '""'], fnKey: 'SUMIFS('
        });
      }
      return H.problem({
        title: 'Share of unpaid amount more than 30 days past due', file: 'piutang_q4.xlsx', sheet: S, base, alt, answerCell: 'B2', precision: 4, dateCells: ['B1'],
        statement: '<p>Of all unpaid <b>Amount</b> (no date in <b>Paid on</b>), what share is more than 30 days past due as of <b>B1</b>? Days past due are B1 minus the <b>Due date</b>. Return it in <b>B2</b> as a decimal (0.25 for 25%).' + note + '</p>',
        solution: '=SUMIFS(' + Dr + ',' + Cr + ',"<"&B1-30,' + Er + ',"")/SUMIFS(' + Dr + ',' + Er + ',"")',
        alts: ['=SUMPRODUCT((B1-' + Cr + '>30)*(' + Er + '="")*' + Dr + ')/SUMPRODUCT((' + Er + '="")*' + Dr + ')', '=SUMIFS(' + Dr + ',' + Er + ',"",' + Cr + ',"<"&B1-30)/SUMIFS(' + Dr + ',' + Er + ',"")'],
        hints: ['More than 30 days past due means the due date is before B1-30.', 'The top of the fraction is the unpaid amount past that limit; the bottom is all unpaid amount.', 'The shape is =SUMIFS(' + Dr + ',' + Cr + ',"<"&B1-30,' + Er + ',"")/SUMIFS(' + Dr + ',' + Er + ',"").'],
        explain: 'Both totals keep only unpaid rows, so the fraction compares like with like. The date limit B1-30 turns "more than 30 days late" into a test on the due date.',
        chips: [Dr, Cr, Er, '"<"&B1-30', '""'], fnKey: 'SUMIFS('
      });
    }
  });

  P.add({
    key: 'late-interest-as-of', difficulty: 'Hard', level: 10, topics: ['SUMPRODUCT', 'ROUND', 'SUMIFS'], count: 3,
    make(r, v) {
      const asof = r.pick(ASOF), asofAlt = asof + r.int(5, 9);
      const offs = r.sample([-12, -3, 2, 6, 9, 15, 21, 28, 33, 40, 47, 55, 63, 72, 86, 104, 131], 16);
      const note = ' B1 is ' + long(asof) + '.';
      if (v === 1) {
        const pool = r.sample(DIST(), 4);
        const rows = aged(r, asof, offs, { paid: 0.3, pool, sortBy: 'date' });
        const late = (name) => rows.filter((x) => x.dist === name && x.paid === '' && x.off > 0).length;
        const names = pool.map((d) => d.name).sort((a, b) => late(b) - late(a) || (a < b ? -1 : 1));
        const who = names[0], whoAlt = names.find((n) => n !== who && late(n) > 0) || names[1];
        /* The distributor also has an unpaid invoice that is not yet due, so leaving out the late test adds
           negative days and gives a different answer. */
        if (!rows.some((x) => x.dist === who && x.paid === '' && x.off < 0)) {
          const k = rows.findIndex((x) => x.off < 0);
          if (k >= 0) { rows[k].dist = who; rows[k].paid = ''; }
        }
        const S = H.table([C_INV, C_DIST, C_DATE, C_TERMS, C_AMT, C_PAID], rows);
        const L = S.last, B = 'B5:B' + L, C = 'C5:C' + L, Dd = 'D5:D' + L, E = 'E5:E' + L, F = 'F5:F' + L;
        const base = Object.assign({}, S.cells, { A1: 'As of', B1: asof, A2: 'Distributor', B2: who, A3: 'Late interest (Rp)' });
        const alt = H.edit(payNews(base, S, rows, asof - 1), { B1: asofAlt, B2: whoAlt });
        return H.problem({
          title: "Late interest on one distributor's overdue invoices", file: 'piutang_q4.xlsx', sheet: S, base, alt, answerCell: 'B3', dateCells: ['B1'],
          statement: '<p>For the distributor in <b>B2</b>, total the late interest on its unpaid invoices (no date in <b>Paid on</b>) as of <b>B1</b>. The due date is <b>Date</b> plus <b>Terms</b>; every day past due costs <b>0.05% of the Amount</b>. Return the total in <b>B3</b>, rounded to whole rupiah.' + note + '</p>',
          solution: '=ROUND(SUMPRODUCT((' + B + '=B2)*(' + F + '="")*(B1-' + C + '-' + Dd + '>0)*(B1-' + C + '-' + Dd + ')*' + E + ')*0.0005,0)',
          alts: ['=ROUND(SUMPRODUCT((' + B + '=B2)*(' + F + '="")*(' + C + '+' + Dd + '<B1)*(B1-(' + C + '+' + Dd + '))*' + E + '*0.0005),0)', '=ROUND(0.0005*SUMPRODUCT((' + F + '="")*(' + B + '=B2)*(B1-' + C + '-' + Dd + '>0)*(B1-' + C + '-' + Dd + ')*' + E + '),0)'],
          hints: ['For one row, the days past due are B1-C5-D5, and they count only when above 0.', 'SUMPRODUCT multiplies tests and numbers row by row: the distributor test, the unpaid test, the late test, the days and the amount.', 'Multiply the total by 0.0005 and wrap it all in ROUND(..., 0).'],
          explain: 'Each test is 1 or 0, so rows for other distributors, paid rows and rows not yet due add nothing. The rest add amount times days late, and 0.0005 turns that into rupiah of interest.',
          chips: [B + '=B2', F + '=""', 'B1-' + C + '-' + Dd, E, '*0.0005'], fnKey: 'SUMPRODUCT('
        });
      }
      const rows = aged(r, asof, offs, { paid: 0.3 });
      if (v === 2) {
        /* One unpaid invoice falls due on B1 itself. It is not past due, so a test of <= where < belongs adds its
           amount to the bottom of the fraction. The rows stay in due-date order: no due date lies between. */
        const k = rows.reduce((b, x, i) => (Math.abs(x.off) < Math.abs(rows[b].off) ? i : b), 0);
        Object.assign(rows[k], { due: asof, off: 0, paid: '' });
      }
      const S = H.table([C_INV, C_DIST, C_DUE, C_AMT, C_PAID], rows);
      const L = S.last, C = 'C5:C' + L, Dr = 'D5:D' + L, E = 'E5:E' + L;
      if (v === 0) {
        const grace = r.pick([7, 10, 14]);
        const base = Object.assign({}, S.cells, { A1: 'As of', B1: asof, A2: 'Grace days', B2: grace, A3: 'Late interest (Rp)' });
        const alt = H.edit(payNews(base, S, rows, asof - 1), { B1: asofAlt, B2: grace === 7 ? 10 : 7 });
        return H.problem({
          title: 'Late interest after a grace period, as of a date', file: 'piutang_q4.xlsx', sheet: S, base, alt, answerCell: 'B3', dateCells: ['B1'],
          statement: '<p>Unpaid invoices (no date in <b>Paid on</b>) that are more than <b>B2</b> days past due as of <b>B1</b> owe late interest of <b>0.05% of the Amount per day</b>, for each day beyond those grace days. An invoice 20 days past due with 7 grace days pays for 13 days. In <b>B3</b>, return the total late interest, rounded to whole rupiah.' + note + '</p>',
          solution: '=ROUND(SUMPRODUCT((' + E + '="")*(B1-' + C + '>B2)*(B1-' + C + '-B2)*' + Dr + ')*0.0005,0)',
          alts: ['=ROUND(SUMPRODUCT((' + E + '="")*(B1-' + C + '>B2)*(B1-' + C + '-B2)*' + Dr + '*0.0005),0)', '=ROUND(0.0005*SUMPRODUCT((' + E + '="")*(' + C + '<B1-B2)*(B1-B2-' + C + ')*' + Dr + '),0)'],
          hints: ['For one row the days charged are B1-C5-B2, and only when that is above 0.', 'SUMPRODUCT multiplies tests and numbers row by row: (' + E + '="")*(B1-' + C + '>B2)*(B1-' + C + '-B2)*' + Dr + '.', 'Multiply the total by 0.0005 and wrap it in ROUND(..., 0).'],
          explain: 'Each test is 1 or 0, so paid rows and rows still inside the grace period add nothing. The rest add amount times days charged, and 0.0005 turns that into rupiah of interest.',
          chips: [E + '=""', 'B1-' + C, '-B2', Dr, '*0.0005'], fnKey: 'SUMPRODUCT('
        });
      }
      const base = Object.assign({}, S.cells, { A1: 'As of', B1: asof, A2: 'Weighted days' });
      const alt = H.edit(payNews(base, S, rows, asof - 1), { B1: asofAlt });
      return H.problem({
        title: 'Amount-weighted average days past due', file: 'piutang_q4.xlsx', sheet: S, base, alt, answerCell: 'B2', dateCells: ['B1'],
        statement: '<p>Credit Control tracks the amount-weighted average days past due. Take the unpaid invoices (no date in <b>Paid on</b>) whose <b>Due date</b> is before <b>B1</b>. Multiply each one\'s days past due (B1 minus Due date) by its <b>Amount</b>, add these up, and divide by the total Amount of those same invoices. Return it in <b>B2</b>, unrounded.' + note + '</p>',
        solution: '=SUMPRODUCT((' + E + '="")*(' + C + '<B1)*(B1-' + C + ')*' + Dr + ')/SUMIFS(' + Dr + ',' + E + ',"",' + C + ',"<"&B1)',
        alts: ['=SUMPRODUCT((' + E + '="")*(' + C + '<B1)*(B1-' + C + ')*' + Dr + ')/SUMPRODUCT((' + E + '="")*(' + C + '<B1)*' + Dr + ')'],
        hints: ['The top of the fraction is SUMPRODUCT of the two tests, the days past due and the amount.', 'The bottom is the amount of the same invoices: SUMIFS with an empty Paid on and a due date before B1.', 'The shape is =SUMPRODUCT((' + E + '="")*(' + C + '<B1)*(B1-' + C + ')*' + Dr + ')/SUMIFS(' + Dr + ',' + E + ',"",' + C + ',"<"&B1).'],
        explain: 'Weighting by amount makes a large invoice that is long overdue count more than a small one. Both halves use the same two tests, so the average covers exactly the overdue, unpaid invoices.',
        chips: [E + '=""', C + '<B1', 'B1-' + C, Dr, '"<"&B1'], fnKey: 'SUMPRODUCT('
      });
    }
  });
});
