/*
 * Practice, Real cases (20 Hard): reports a junior analyst at Rasa Kenari or Bank Kenari builds from several workbooks
 * at once. The learner's own file holds the main list; the loan book, the collection report, the product master, price
 * lists, targets, rate tables, policies and schedules sit in other open workbooks (and sometimes on a second sheet of the
 * same file). Every formula looks things up in those files first, then does the math: a credit decision, an
 * early-warning flag, a margin, a commission, a collection queue, a reorder.
 * Each template makes `count` problems; v (0 to count-1) picks the variant and r is its seeded random.
 * Levels: 8 is a lookup feeding a SUMPRODUCT or one IF around arithmetic, 9 joins two or three workbooks inside a
 * decision, 10 builds a score or a queue from three or four workbooks.
 *
 * Ranges in other workbooks are written as Excel writes them from the main file, pinned with $:
 * [buku_pinjaman.xlsx]Pinjaman!$B$5:$B$18 (H.ext). The hidden sheet changes values in the other workbooks as well, so a
 * typed or copied number fails. What the practice engine does not do, so no problem here relies on it: a lookup that
 * searches for a whole range at once (XLOOKUP(B5:B18, ...)) inside SUMPRODUCT, SUMIFS with an array as criteria, INDEX
 * with an array of row numbers, and ISBLANK on a range. No row sits exactly on a boundary that floating-point arithmetic
 * could tip (such as 70% of a sales figure); exact boundaries are kept to whole numbers and exact divisions. LET names
 * avoid c and r, which Excel refuses as names.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) factory(require('./core.js'));
  else factory(root.KicauPractice);
})(typeof globalThis !== 'undefined' ? globalThis : this, function (P) {
  'use strict';
  const H = P.H;
  const L = H.LETTERS;
  const DIST = () => H.db().distributors;
  const NAMES = () => DIST().map((d) => d.name);
  const regionOf = (name) => DIST().find((d) => d.name === name).region;
  const nums = (n) => Array.from({ length: n }, (_, i) => i);
  const rp = (n) => 'Rp ' + Math.round(n).toLocaleString('en-US');
  const ser = (iso) => H.toSerial(iso);
  const r100k = (x) => Math.round(x / 100000) * 100000;
  const day2 = (n) => String(n).padStart(2, '0');

  /* The made-up world, summed: a distributor's net sales in one month ('2026-05'). */
  const netIn = (name, m) => H.db().sales.filter((s) => s.distributor === name && s.date.startsWith(m)).reduce((a, s) => a + s.net, 0);
  /* Keep x[top]/x[bottom] at least `gap` away from every limit, by raising the top figure Rp 100,000 at a time. */
  function clear(x, top, bottom, limits, gap) { const ls = [].concat(limits); while (ls.some((l) => Math.abs(x[top] / x[bottom] - l) < gap)) x[top] += 100000; }

  /*
   * One sheet of another open workbook (file null: a second sheet of the main workbook). w.x writes its ranges as Excel
   * does from the main file, such as [buku_pinjaman.xlsx]Pinjaman!$B$5:$B$18; w.alt is its hidden-test version.
   */
  function sheetOf(file, name, S) { return { file, name, S, x: H.ext(file, name, S), alt: S.cells }; }
  /* A problem across workbooks: o.open lists the other sheets, grouped here into books (other files) and sheets (same file). */
  function multi(o) {
    const books = [], sheets = [];
    for (const w of o.open) {
      const entry = { name: w.name, sheet: w.S, base: w.S.cells, alt: w.alt };
      if (!w.file) { sheets.push(entry); continue; }
      let b = books.find((x) => x.file === w.file);
      if (!b) { b = { file: w.file, sheets: [] }; books.push(b); }
      b.sheets.push(entry);
    }
    const p = Object.assign({}, o, { books, sheets });
    delete p.open;
    return H.problem(p);
  }
  /* Where row i of a list of objects keyed by `k` holds `v`. */
  const at = (list, v, k) => list.findIndex((x) => x[k || 'd'] === v);

  /* A grid with numbers across its header row (row 4), starting at column startCol: one labelled line per row. */
  function grid(corner, heads, lines, startCol) {
    const sc = startCol, cells = {};
    cells[L[sc] + '4'] = corner;
    heads.forEach((h, j) => { cells[L[sc + 1 + j] + '4'] = h; });
    lines.forEach((ln, i) => { cells[L[sc] + (5 + i)] = ln.h; ln.vals.forEach((x, j) => { cells[L[sc + 1 + j] + (5 + i)] = x; }); });
    const cols = L.slice(sc, sc + heads.length + 1);
    return { cells, cols, num: cols.slice(1), dateCols: [], first: 5, last: 4 + lines.length, map: {}, next: L[sc + heads.length + 1] };
  }

  /* ======================= Credit decisions: the loan book, the collection report, ratings, collateral ======================= */

  P.add({
    key: 'credit-decision-rules', difficulty: 'Hard', level: 9, topics: ['IF', 'OR', 'XLOOKUP', 'MAXIFS', 'SUMIFS', 'credit', 'Several workbooks'], count: 4, multi: true,
    make(r, v) {
      return [
        () => {
          /* A drawdown request: the days late from the collection report, the restructuring flag and the room under the
             limit from the loan book. */
          const ds = r.shuffle(NAMES());
          const asked = ds.slice(0, 12);
          const info = {};
          ds.forEach((d) => {
            const limit = r.pick([100, 150, 200, 250]) * 1000000;
            info[d] = { limit, out: r100k(limit * r.pick([0.35, 0.5, 0.62, 0.75, 0.85])), rs: r.chance(0.12) ? 'Y' : 'N', late: r.pick([0, 0, 0, 0, 5, 12, 30, 45, 95]) };
          });
          const room = (d) => info[d].limit - info[d].out;
          const fits = (d) => Math.max(1, Math.floor(room(d) * 0.5 / 5000000)) * 5000000;
          const req = {};
          asked.forEach((d) => { req[d] = Math.max(1, Math.floor(room(d) * r.pick([0.3, 0.6, 1.4]) / 5000000)) * 5000000; });
          /* Planted: days late exactly 60 and 61, restructured with no arrears, a request that fills the limit exactly, one
             over it, and three that change on the hidden test. */
          const [p0, p1, p2, p3, p4, p5, p6] = asked;
          Object.assign(info[p0], { late: 60, rs: 'N' }); req[p0] = fits(p0);
          Object.assign(info[p1], { late: 61, rs: 'N' }); req[p1] = fits(p1);
          Object.assign(info[p2], { late: 0, rs: 'Y' }); req[p2] = fits(p2);
          Object.assign(info[p3], { late: r.pick([0, 5, 12]), rs: 'N' }); req[p3] = r.int(4, 10) * 5000000; info[p3].out = info[p3].limit - req[p3];
          Object.assign(info[p4], { late: r.pick([0, 12]), rs: 'N' }); req[p4] = room(p4) + r.int(1, 4) * 5000000;
          Object.assign(info[p5], { late: 0, rs: 'N' }); req[p5] = fits(p5);
          Object.assign(info[p6], { late: 0, rs: 'Y' }); req[p6] = fits(p6);
          const rows = r.shuffle(asked).map((d, i) => ({ req: 'CR-' + (411 + i * 3), d, amt: req[d], tenor: r.pick([6, 12, 18, 24]) }));
          const S = H.table([{ h: 'Request', k: 'req' }, { h: 'Distributor', k: 'd' }, { h: 'Requested (Rp)', k: 'amt', num: true }, { h: 'Tenor (months)', k: 'tenor', num: true }], rows);
          const book = r.shuffle(ds).map((d, i) => ({ f: 'F' + (101 + i), d, limit: info[d].limit, out: info[d].out, rs: info[d].rs }));
          const PJ = H.table([{ h: 'Facility', k: 'f' }, { h: 'Distributor', k: 'd' }, { h: 'Limit (Rp)', k: 'limit', num: true }, { h: 'Outstanding (Rp)', k: 'out', num: true }, { h: 'Restructured', k: 'rs' }], book);
          const due = r.shuffle(ds).map((d) => ({ d, late: info[d].late, who: r.pick(['Nanda', 'Rizal', 'Putri', 'Bayu']) }));
          const TG = H.table([{ h: 'Distributor', k: 'd' }, { h: 'Collector', k: 'who' }, { h: 'Days late', k: 'late', num: true }], due);
          const pj = sheetOf('buku_pinjaman.xlsx', 'Pinjaman', PJ), tg = sheetOf('laporan_tagihan.xlsx', 'Tunggakan', TG);
          /* The hidden test: the bank raises one limit and corrects a restructuring flag to N, one borrower falls 75 days
             behind in the collection report, and one request grows. */
          pj.alt = H.edit(PJ.cells, { [PJ.cell('Limit (Rp)', at(book, p4))]: (x) => x + 50000000, [PJ.cell('Restructured', at(book, p6))]: 'N' });
          tg.alt = H.edit(TG.cells, { [TG.cell('Days late', at(due, p5))]: 75 });
          const alt = H.edit(S.cells, { [S.cell('Requested (Rp)', at(rows, p0))]: (x) => x + 5000000 });
          const X = pj.x, Y = tg.x;
          const lk = (w, h) => `XLOOKUP(B5,${w.A('Distributor')},${w.A(h)})`;
          const im = (w, h) => `INDEX(${w.A(h)},MATCH(B5,${w.A('Distributor')},0))`;
          const PT = X.T('Distributor', 'Restructured'), TT = Y.T('Distributor', 'Days late');
          const f = S.next + S.first, l = S.next + S.last;
          return multi({
            title: 'Drawdown decision from the loan book and the collection report', file: 'komite_kredit.xlsx', sheetName: 'Pengajuan', sheet: S, base: S.cells, alt, fill: true, ansLabel: 'Decision', open: [pj, tg],
            statement: `<p>Three workbooks are open: Dimas's committee list (<b>komite_kredit.xlsx</b>), Bank Kenari's loan book (<b>buku_pinjaman.xlsx</b>, sheet Pinjaman) and the collection team's report (<b>laporan_tagihan.xlsx</b>, sheet Tunggakan). Each row of the committee list is a distributor asking to draw more on its working-capital line. The other two files list every distributor, each in its own order.</p><p>The rule: <b>Decline</b> when the distributor's <b>Days late</b> in the collection report is more than 60, <b>or</b> the loan book shows its loan was <b>Restructured</b> (Y). Otherwise <b>Approve</b> when its <b>Outstanding</b> plus the amount <b>Requested</b> is at most its <b>Limit</b>, and <b>Refer</b> when it is above. In <b>${f}</b>, return Approve, Refer or Decline and fill it down to <b>${l}</b>.</p>`,
            solution: `=IF(OR(${lk(Y, 'Days late')}>60,${lk(X, 'Restructured')}="Y"),"Decline",IF(${lk(X, 'Outstanding (Rp)')}+C5<=${lk(X, 'Limit (Rp)')},"Approve","Refer"))`,
            alts: [`=IF(OR(VLOOKUP(B5,${TT},3,FALSE)>60,VLOOKUP(B5,${PT},4,FALSE)="Y"),"Decline",IF(VLOOKUP(B5,${PT},3,FALSE)+C5<=VLOOKUP(B5,${PT},2,FALSE),"Approve","Refer"))`,
              `=IF(OR(${im(Y, 'Days late')}>60,${im(X, 'Restructured')}="Y"),"Decline",IF(${im(X, 'Outstanding (Rp)')}+C5<=${im(X, 'Limit (Rp)')},"Approve","Refer"))`,
              `=LET(late,${lk(Y, 'Days late')},room,${lk(X, 'Limit (Rp)')}-${lk(X, 'Outstanding (Rp)')},IF(OR(late>60,${lk(X, 'Restructured')}="Y"),"Decline",IF(C5<=room,"Approve","Refer")))`],
            hints: ['Four facts come from the other two files: the days late from laporan_tagihan.xlsx, and the restructuring flag, the outstanding and the limit from buku_pinjaman.xlsx. Each is a lookup of the distributor in B5.',
              `Tap a workbook tab, then the column, and Kicau writes the range for you: ${Y.A('Distributor')}. The red flags go inside OR: XLOOKUP(B5, ..., Days late)>60 and XLOOKUP(B5, ..., Restructured)="Y".`,
              'Shape: =IF(OR(days late>60, restructured="Y"), "Decline", IF(outstanding + C5 <= limit, "Approve", "Refer")), with each fact an XLOOKUP into its own file.'],
            explain: 'The collection report and the loan book each list the distributors in their own order, so every fact is found by name. OR declines on either red flag first. Only a clean borrower reaches the inner IF, which adds the request to what it already owes and compares that with its limit. Exactly 60 days late is not more than 60, and a request that fills the limit to the rupiah is still approved.',
            chips: ['B5', 'C5', Y.A('Distributor'), Y.A('Days late'), X.A('Distributor'), X.A('Restructured'), X.A('Outstanding (Rp)'), X.A('Limit (Rp)'), '>60', '="Y"', '"Decline"', '"Approve"', '"Refer"'], fnKey: 'IF('
          });
        },
        () => {
          /* Three outcomes: Decline on bad DPD (loan book), Refer on a large request against average monthly sales (sales
             file, AVERAGEIFS) or any arrears, Approve otherwise. */
          const ds = r.shuffle(NAMES());
          const asked = ds.slice(0, 10), sold = ds.slice(0, 12);
          const MON = [['Aug', '2026-08'], ['Sep', '2026-09'], ['Oct', '2026-10']];
          const net = {};
          sold.forEach((d) => {
            const x = MON.map(([, m]) => Math.max(2000000, r100k(netIn(d, m) * (0.9 + r() * 0.2))));
            /* A three-month total that divides by 3 into whole hundreds of thousands, so twice the average is exact. */
            while ((x[0] + x[1] + x[2]) % 300000) x[2] += 100000;
            net[d] = x;
          });
          const avg = (d) => (net[d][0] + net[d][1] + net[d][2]) / 3;
          const dpd = {};
          ds.forEach((d) => { dpd[d] = r.pick([0, 0, 0, 0, 0, 5, 12, 30, 45, 75, 95]); });
          const q = {};
          asked.forEach((d) => { q[d] = r.pick([0.8, 1.1, 1.3, 1.6, 1.8, 2.4, 2.8]); });
          /* Planted [DPD, request as a multiple of average sales]: exactly 2 times, DPD exactly 60 and 61, well over 2 times,
             and two that change on the hidden test. */
          [[0, 2], [60, 1.2], [61, 1.1], [0, 2.6], [0, 1.2], [0, 1.6]].forEach(([dp, k], i) => { dpd[asked[i]] = dp; q[asked[i]] = k; });
          const amt = {};
          asked.forEach((d) => { amt[d] = q[d] === 2 ? 2 * avg(d) : Math.round(avg(d) * q[d] / 1000000) * 1000000; });
          const lines = [];
          MON.forEach(([m], k) => r.shuffle(sold).forEach((d) => lines.push({ m, d, net: net[d][k] })));
          const BL = H.table([{ h: 'Month', k: 'm' }, { h: 'Distributor', k: 'd' }, { h: 'Net (Rp)', k: 'net', num: true }], lines);
          const book = r.shuffle(ds).map((d, i) => ({ f: 'F' + (201 + i), d, out: r.int(20, 90) * 1000000, dpd: dpd[d] }));
          const PJ = H.table([{ h: 'Facility', k: 'f' }, { h: 'Distributor', k: 'd' }, { h: 'Outstanding (Rp)', k: 'out', num: true }, { h: 'DPD', k: 'dpd', num: true }], book);
          const rows = r.shuffle(asked).map((d, i) => ({ req: 'CR-' + (430 + i * 2), d, amt: amt[d] }));
          const S = H.table([{ h: 'Request', k: 'req' }, { h: 'Distributor', k: 'd' }, { h: 'Requested (Rp)', k: 'amt', num: true }], rows);
          const bl = sheetOf('penjualan_bulanan.xlsx', 'Bulanan', BL), pj = sheetOf('buku_pinjaman.xlsx', 'Pinjaman', PJ);
          /* The hidden test: returns cut one distributor's September and October sales, so its request is now more than twice
             the average, and one clean borrower falls 75 days behind at the bank. */
          const cut = asked[5], li = (m) => lines.findIndex((x) => x.m === m && x.d === cut);
          const sep2 = r100k(net[cut][1] * 0.6);
          let oct2 = r100k(net[cut][2] * 0.3);
          while (2 * (net[cut][0] + sep2 + oct2) / 3 >= amt[cut] - 1000000 && oct2 > 100000) oct2 -= 100000;
          bl.alt = H.edit(BL.cells, { [BL.cell('Net (Rp)', li('Sep'))]: sep2, [BL.cell('Net (Rp)', li('Oct'))]: oct2 });
          pj.alt = H.edit(PJ.cells, { [PJ.cell('DPD', at(book, asked[4]))]: 75 });
          const X = pj.x, B = bl.x;
          const dp = `XLOOKUP(B5,${X.A('Distributor')},${X.A('DPD')})`, av = `AVERAGEIFS(${B.A('Net (Rp)')},${B.A('Distributor')},B5)`;
          const sm = `SUMIFS(${B.A('Net (Rp)')},${B.A('Distributor')},B5)/COUNTIFS(${B.A('Distributor')},B5)`;
          const vd = `VLOOKUP(B5,${X.T('Distributor', 'DPD')},3,FALSE)`, id = `INDEX(${X.A('DPD')},MATCH(B5,${X.A('Distributor')},0))`;
          const f = S.next + S.first, l = S.next + S.last;
          return multi({
            title: 'Approve, refer or decline on bank DPD and average monthly sales', file: 'pengajuan_kredit.xlsx', sheetName: 'Pengajuan', sheet: S, base: S.cells, alt: S.cells, fill: true, ansLabel: 'Decision', open: [bl, pj],
            statement: `<p>Three workbooks are open: the loan requests (<b>pengajuan_kredit.xlsx</b>), Rasa Kenari's monthly sales (<b>penjualan_bulanan.xlsx</b>, sheet Bulanan, one row per distributor per month from August to October) and Bank Kenari's loan book (<b>buku_pinjaman.xlsx</b>, sheet Pinjaman, with each distributor's days past due, <b>DPD</b>).</p><p>The rule, in this order: <b>Decline</b> when the distributor's DPD is more than 60. Otherwise <b>Refer</b> when the request is more than 2 times its average monthly <b>Net</b> sales (the average of its three months), <b>or</b> its DPD is more than 0. Otherwise <b>Approve</b>. In <b>${f}</b>, return the decision and fill it down to <b>${l}</b>.</p>`,
            solution: `=IF(${dp}>60,"Decline",IF(OR(C5>2*${av},${dp}>0),"Refer","Approve"))`,
            alts: [`=IFS(${dp}>60,"Decline",OR(C5>2*${av},${dp}>0),"Refer",TRUE,"Approve")`,
              `=IF(${vd}>60,"Decline",IF(OR(C5>2*${sm},${vd}>0),"Refer","Approve"))`,
              `=LET(d,${dp},s,${av},IF(d>60,"Decline",IF(OR(C5>2*s,d>0),"Refer","Approve")))`,
              `=IF(${id}>60,"Decline",IF(OR(C5>2*${av},${id}>0),"Refer","Approve"))`],
            hints: ['Two files, two kinds of answer: the DPD is one value per distributor, so look it up; the sales are three rows per distributor, so average them with AVERAGEIFS.',
              `${av} is the average monthly sales of the distributor in B5. ${dp} is its DPD.`,
              'Shape: =IF(DPD>60, "Decline", IF(OR(C5>2*average, DPD>0), "Refer", "Approve")), with DPD an XLOOKUP into buku_pinjaman.xlsx and average an AVERAGEIFS over penjualan_bulanan.xlsx.'],
            explain: 'AVERAGEIFS averages the three monthly rows of the distributor in the sales file, and XLOOKUP brings its DPD from the loan book. The outer IF settles the worst case first, so a borrower 61 days behind is declined even though it would also be referred. A request of exactly twice the average is not more than twice, so with no arrears it is approved, and a DPD of exactly 60 is referred, not declined.',
            chips: ['B5', 'C5', B.A('Net (Rp)'), B.A('Distributor'), X.A('Distributor'), X.A('DPD'), 'AVERAGEIFS(', '2*', '>60', '>0', '"Decline"', '"Refer"', '"Approve"'], fnKey: 'IF('
          });
        },
        () => {
          /* A borrower across all its facilities (MAXIFS, COUNTIFS, SUMIFS on the loan book), against a cap that depends on
             its grade: the grade from the rating file, the cap from the policy sheet of the main workbook. */
          const CAP = { A: 300000000, B: 225000000, C: 150000000, D: 100000000 };
          const NEXT = { A: 'B', B: 'C', C: 'D' };
          const ds = r.sample(NAMES(), 10);
          const asked = ds.slice(0, 8);
          const [b0, b1, b2, b3, b4, b5, b6, b7] = asked;
          const grade = {};
          ds.forEach((d) => { grade[d] = r.pick(['A', 'B', 'B', 'C', 'C', 'D']); });
          if (grade[b5] === 'D') grade[b5] = r.pick(['A', 'B', 'C']);
          grade[b7] = grade[b5];
          const TY = ['Working Capital', 'Term Loan', 'Trade Finance'];
          const fac = [];
          ds.forEach((d, k) => {
            const n = k < 4 ? r.int(2, 3) : r.int(1, 3), share = r.pick([0.45, 0.55, 0.65, 0.75]);
            r.sample(TY, n).forEach((t) => fac.push({ d, t, out: Math.max(1, Math.round(CAP[grade[d]] * share / n / 2500000)) * 2500000, dpd: r.pick([0, 0, 0, 7, 15]), rs: 'N' }));
          });
          const of = (d) => fac.filter((x) => x.d === d);
          const room = (d) => CAP[grade[d]] - of(d).reduce((a, x) => a + x.out, 0);
          const half = (d) => Math.max(1, Math.floor(room(d) * 0.5 / 5000000)) * 5000000;
          /* Planted: worst DPD exactly 30; a request that fills the cap exactly; restructured and over the cap (the red flag
             decides first); 45 days behind; over the cap; and three that change on the hidden test. */
          r.pick(of(b0)).dpd = 30;
          r.pick(of(b2)).rs = 'Y';
          r.pick(of(b3)).dpd = 45;
          if (r.chance(0.5)) r.pick(of(ds[8])).rs = 'Y';
          const req = { [b0]: half(b0), [b1]: room(b1), [b2]: room(b2) + r.int(1, 4) * 5000000, [b3]: half(b3), [b4]: room(b4) + r.int(1, 4) * 5000000, [b5]: room(b5) - 5000000, [b6]: half(b6), [b7]: room(b7) + 10000000 };
          const rows = r.shuffle(asked).map((d, i) => ({ req: 'CR-' + (601 + i * 2), d, amt: req[d] }));
          const S = H.table([{ h: 'Request', k: 'req' }, { h: 'Borrower', k: 'd' }, { h: 'New request (Rp)', k: 'amt', num: true }], rows);
          const KB = H.table([{ h: 'Grade', k: 'g' }, { h: 'Max exposure (Rp)', k: 'cap', num: true }], ['A', 'B', 'C', 'D'].map((g) => ({ g, cap: CAP[g] })));
          const list = r.shuffle(fac).map((x, i) => Object.assign(x, { f: 'F' + (311 + i) }));
          const FS = H.table([{ h: 'Facility', k: 'f' }, { h: 'Borrower', k: 'd' }, { h: 'Type', k: 't' }, { h: 'Outstanding (Rp)', k: 'out', num: true }, { h: 'DPD', k: 'dpd', num: true }, { h: 'Restructured', k: 'rs' }], list);
          const rated = r.shuffle(ds).map((d) => ({ d, on: '2026-' + day2(r.int(4, 9)) + '-' + day2(r.int(1, 28)), g: grade[d] }));
          const RT = H.table([{ h: 'Borrower', k: 'd' }, { h: 'Last review', k: 'on', date: true }, { h: 'Grade', k: 'g' }], rated);
          const kb = sheetOf(null, 'Kebijakan', KB), fs = sheetOf('buku_pinjaman.xlsx', 'Fasilitas', FS), rt = sheetOf('rating_debitur.xlsx', 'Rating', RT);
          /* The hidden test: one borrower is downgraded a grade, one facility slips to 35 days, and the policy raises one cap. */
          rt.alt = H.edit(RT.cells, { [RT.cell('Grade', at(rated, b5))]: NEXT[grade[b5]] });
          fs.alt = H.edit(FS.cells, { [FS.cell('DPD', list.indexOf(of(b6)[0]))]: 35 });
          kb.alt = H.edit(KB.cells, { [KB.cell('Max exposure (Rp)', ['A', 'B', 'C', 'D'].indexOf(grade[b7]))]: (x) => x + 50000000 });
          const F = fs.x, BR = F.A('Borrower'), DP = F.A('DPD'), RS = F.A('Restructured'), OU = F.A('Outstanding (Rp)');
          const cap1 = `XLOOKUP(XLOOKUP(B5,${rt.x.A('Borrower')},${rt.x.A('Grade')}),${kb.x.A('Grade')},${kb.x.A('Max exposure (Rp)')})`;
          const cap2 = `VLOOKUP(VLOOKUP(B5,${rt.x.T('Borrower', 'Grade')},3,FALSE),${kb.x.T('Grade', 'Max exposure (Rp)')},2,FALSE)`;
          const cap3 = `INDEX(${kb.x.A('Max exposure (Rp)')},MATCH(INDEX(${rt.x.A('Grade')},MATCH(B5,${rt.x.A('Borrower')},0)),${kb.x.A('Grade')},0))`;
          const flags = `OR(MAXIFS(${DP},${BR},B5)>30,COUNTIFS(${BR},B5,${RS},"Y")>0)`;
          const f = S.next + S.first, l = S.next + S.last;
          return multi({
            title: 'Credit decision on group exposure against the rating cap', file: 'eksposur_debitur.xlsx', sheetName: 'Pengajuan', sheet: S, base: S.cells, alt: S.cells, fill: true, ansLabel: 'Decision', open: [kb, fs, rt],
            statement: `<p>Three workbooks are open: the credit desk's file (<b>eksposur_debitur.xlsx</b>), with these requests on sheet Pengajuan and the bank's policy on sheet <b>Kebijakan</b> (the most a borrower of each grade may owe in all); Bank Kenari's loan book (<b>buku_pinjaman.xlsx</b>, sheet Fasilitas), where a borrower can hold several facilities; and the rating file (<b>rating_debitur.xlsx</b>, sheet Rating) with each borrower's <b>Grade</b>.</p><p>The rule: <b>Decline</b> when any of the borrower's facilities is more than 30 days past due (<b>DPD</b>) <b>or</b> any of them was restructured (Y). Otherwise <b>Approve</b> when its total <b>Outstanding</b> plus the <b>New request</b> is at most the <b>Max exposure</b> for its grade, and <b>Refer</b> when it is above. In <b>${f}</b>, return the decision and fill it down to <b>${l}</b>.</p>`,
            solution: `=IF(${flags},"Decline",IF(SUMIFS(${OU},${BR},B5)+C5<=${cap1},"Approve","Refer"))`,
            alts: [`=IF(COUNTIFS(${BR},B5,${DP},">30")+COUNTIFS(${BR},B5,${RS},"Y")>0,"Decline",IF(SUMIFS(${OU},${BR},B5)+C5>${cap2},"Refer","Approve"))`,
              `=IF(SUMPRODUCT((${BR}=B5)*((${DP}>30)+(${RS}="Y")))>0,"Decline",IF(SUMPRODUCT((${BR}=B5)*${OU})+C5<=${cap3},"Approve","Refer"))`,
              `=LET(cap,${cap1},IFS(${flags},"Decline",SUMIFS(${OU},${BR},B5)+C5<=cap,"Approve",TRUE,"Refer"))`],
            hints: ['The loan book has several rows per borrower, so ask it questions with conditions: MAXIFS for the worst DPD, COUNTIFS for any Y, SUMIFS for the total outstanding.',
              `The cap takes two lookups: the grade from rating_debitur.xlsx, then that grade's cap on the Kebijakan sheet. ${cap1} gives the cap.`,
              'Shape: =IF(OR(MAXIFS(DPD, borrower, B5)>30, COUNTIFS(borrower, B5, restructured, "Y")>0), "Decline", IF(SUMIFS(outstanding, borrower, B5)+C5 <= cap, "Approve", "Refer"))'],
            explain: 'MAXIFS finds the borrower\'s worst arrears and COUNTIFS counts its restructured facilities across the whole loan book, so OR declines on either flag. A clean borrower reaches the inner IF, where SUMIFS adds up what it owes. The inner XLOOKUP turns the borrower into a grade, and the outer one turns the grade into a cap from the Kebijakan sheet, a sheet of your own file, so its range has no file name in brackets. A worst DPD of exactly 30 is not more than 30.',
            chips: ['B5', 'C5', BR, DP, RS, OU, rt.x.A('Borrower'), rt.x.A('Grade'), kb.x.A('Grade'), kb.x.A('Max exposure (Rp)'), '"Y"', '"Decline"', '"Approve"', '"Refer"'], fnKey: 'IF('
          });
        },
        () => {
          /* Collateral cover: the appraised value from the valuation file, the haircut by type from the policy file, and the
             DPD from the loan book: four workbooks. */
          const TYPES = ['Land', 'Building', 'Vehicle', 'Inventory', 'Receivables'];
          const HC = { Land: r.pick([25, 30]), Building: r.pick([35, 40]), Vehicle: r.pick([45, 50]), Inventory: r.pick([55, 60]), Receivables: r.pick([40, 45]) };
          const up = r.pick(['Building', 'Vehicle', 'Inventory', 'Receivables']);
          const HC2 = Object.assign({}, HC, { [up]: HC[up] + 10 });
          const cov = (w, h) => w * (100 - h) / 100;
          const ds = r.shuffle(NAMES());
          const rows = ds.slice(0, 12).map((d) => ({ d, amt: r.int(6, 30) * 5000000, c: r.pick(TYPES), q: r.pick([0.7, 0.85, 1.15, 1.3, 1.6]), dpd: r.pick([0, 0, 0, 0, 8, 21, 45, 92]) }));
          /* Planted [cover ratio, DPD]: covered at DPD exactly 30 and 31, short of cover, clean, and one that is covered only
             under the old haircut. */
          [[1.3, 30], [1.3, 31], [0.85, 0], [1.3, 0], [1.12, 0]].forEach(([q, dp], k) => { rows[k].q = q; rows[k].dpd = dp; });
          if (rows[2].c === up) rows[2].c = 'Land';
          rows[4].c = up; rows[4].amt = r.int(16, 30) * 5000000;
          /* An appraised value near q times the request after haircut, at least Rp 1,000,000 away from the request under
             both the old and the new haircut. */
          const valueFor = (x, q) => {
            let w = Math.round(x.amt / ((100 - HC[x.c]) / 100) * q / 5000000) * 5000000;
            const bad = (y) => (q < 1 ? cov(y, HC[x.c]) > x.amt - 1000000 : cov(y, HC[x.c]) < x.amt + 1000000) || Math.abs(cov(y, HC2[x.c]) - x.amt) < 1000000;
            while (bad(w)) w += q < 1 ? -5000000 : 5000000;
            return w;
          };
          rows.forEach((x) => { x.val = valueFor(x, x.q); });
          const short = rows[2], clean = rows[3];
          const list = r.shuffle(rows).map((x, i) => Object.assign(x, { req: 'CR-' + (520 + i * 2) }));
          const S = H.table([{ h: 'Request', k: 'req' }, { h: 'Distributor', k: 'd' }, { h: 'Requested (Rp)', k: 'amt', num: true }, { h: 'Collateral', k: 'c' }], list);
          const valued = r.shuffle(list).map((x) => ({ req: x.req, on: '2026-' + r.pick(['08', '09', '10']) + '-' + day2(r.int(1, 28)), val: x.val }));
          const PN = H.table([{ h: 'Request', k: 'req' }, { h: 'Valued on', k: 'on', date: true }, { h: 'Appraised value (Rp)', k: 'val', num: true }], valued);
          const HK = H.table([{ h: 'Collateral', k: 'c' }, { h: 'Haircut (%)', k: 'h', num: true }], TYPES.map((c) => ({ c, h: HC[c] })));
          const dpd = {};
          ds.forEach((d) => { dpd[d] = r.pick([0, 0, 0, 5, 14, 40]); });
          rows.forEach((x) => { dpd[x.d] = x.dpd; });
          const book = r.shuffle(ds).map((d, i) => ({ f: 'F' + (401 + i), d, out: r.int(10, 80) * 1000000, dpd: dpd[d] }));
          const PJ = H.table([{ h: 'Facility', k: 'f' }, { h: 'Distributor', k: 'd' }, { h: 'Outstanding (Rp)', k: 'out', num: true }, { h: 'DPD', k: 'dpd', num: true }], book);
          const pn = sheetOf('penilaian_agunan.xlsx', 'Penilaian', PN), hk = sheetOf('kebijakan_haircut.xlsx', 'Haircut', HK), pj = sheetOf('buku_pinjaman.xlsx', 'Pinjaman', PJ);
          /* The hidden test: a fresh valuation covers the short request, the policy haircut on one type rises 10 points, and a
             clean borrower falls 45 days behind. */
          pn.alt = H.edit(PN.cells, { [PN.cell('Appraised value (Rp)', at(valued, short.req, 'req'))]: valueFor(short, 1.3) });
          hk.alt = H.edit(HK.cells, { [HK.cell('Haircut (%)', TYPES.indexOf(up))]: HC2[up] });
          pj.alt = H.edit(PJ.cells, { [PJ.cell('DPD', at(book, clean.d))]: 45 });
          const val = `XLOOKUP(A5,${pn.x.A('Request')},${pn.x.A('Appraised value (Rp)')})`, hc = `XLOOKUP(D5,${hk.x.A('Collateral')},${hk.x.A('Haircut (%)')})`, dp = `XLOOKUP(B5,${pj.x.A('Distributor')},${pj.x.A('DPD')})`;
          const vv = `VLOOKUP(A5,${pn.x.T('Request', 'Appraised value (Rp)')},3,FALSE)`, vh = `VLOOKUP(D5,${hk.x.T('Collateral', 'Haircut (%)')},2,FALSE)`, vd = `VLOOKUP(B5,${pj.x.T('Distributor', 'DPD')},3,FALSE)`;
          const iv = `INDEX(${pn.x.A('Appraised value (Rp)')},MATCH(A5,${pn.x.A('Request')},0))`, ih = `INDEX(${hk.x.A('Haircut (%)')},MATCH(D5,${hk.x.A('Collateral')},0))`, id = `INDEX(${pj.x.A('DPD')},MATCH(B5,${pj.x.A('Distributor')},0))`;
          const f = S.next + S.first, l = S.next + S.last;
          return multi({
            title: 'Credit decision on appraised collateral after the policy haircut', file: 'agunan_kredit.xlsx', sheetName: 'Pengajuan', sheet: S, base: S.cells, alt: S.cells, fill: true, ansLabel: 'Decision', open: [pn, hk, pj],
            statement: `<p>Four workbooks are open: the requests with their collateral type (<b>agunan_kredit.xlsx</b>), the appraisers' values by request number (<b>penilaian_agunan.xlsx</b>, sheet Penilaian), the bank's haircut policy by collateral type (<b>kebijakan_haircut.xlsx</b>, sheet Haircut; a haircut of 30 means only 70% of the value counts) and the loan book with each distributor's days past due (<b>buku_pinjaman.xlsx</b>, sheet Pinjaman).</p><p>The rule: <b>Decline</b> when the <b>Appraised value</b> after its haircut (value × (1 - haircut/100)) is less than the amount <b>Requested</b>, <b>or</b> the distributor's <b>DPD</b> is more than 30. Otherwise <b>Approve</b>. In <b>${f}</b>, return the decision and fill it down to <b>${l}</b>.</p>`,
            solution: `=IF(OR(${val}*(1-${hc}/100)<C5,${dp}>30),"Decline","Approve")`,
            alts: [`=IF(OR(${vv}*(1-${vh}/100)<C5,${vd}>30),"Decline","Approve")`,
              `=IF(AND(${val}*(100-${hc})/100>=C5,${dp}<=30),"Approve","Decline")`,
              `=IF(OR(${iv}*(1-${ih}/100)<C5,${id}>30),"Decline","Approve")`,
              `=LET(cover,${val}*(1-${hc}/100),IF(OR(cover<C5,${dp}>30),"Decline","Approve"))`],
            hints: ['Three lookups, each in its own file and each on a different key: the value by the request number in A5, the haircut by the collateral type in D5, the DPD by the distributor in B5.',
              `The cover is ${val}*(1-${hc}/100). Compare it with C5, and test the DPD from the loan book, inside OR.`,
              'Shape: =IF(OR(value*(1-haircut/100) < C5, DPD > 30), "Decline", "Approve")'],
            explain: 'The valuation, the policy and the loan book each answer one question, so each needs its own lookup with its own key. Taking the haircut off the appraised value gives the part the bank counts, and OR declines when that falls short of the request or the arrears pass 30 days. A DPD of exactly 30 is still approved.',
            chips: ['A5', 'B5', 'C5', 'D5', pn.x.A('Request'), pn.x.A('Appraised value (Rp)'), hk.x.A('Collateral'), hk.x.A('Haircut (%)'), pj.x.A('Distributor'), pj.x.A('DPD'), '(1-', '/100)', '"Decline"', '"Approve"'], fnKey: 'IF('
          });
        }
      ][v]();
    }
  });

  /* ======================= Early-warning flags: Rasa Kenari's sales and Bank Kenari's loan book ======================= */

  P.add({
    key: 'early-warning-flags', difficulty: 'Hard', level: 9, topics: ['IF', 'AND', 'OR', 'XLOOKUP', 'SUMIFS', 'early warning', 'Several workbooks'], count: 4, multi: true,
    make(r, v) {
      return [
        () => {
          /* Watch only when both signals fire: the sales drop from Rasa Kenari's file and the DPD from the loan book. */
          const M6 = ['May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct'];
          const all = DIST().map((d) => {
            const x = { d: d.name, region: d.region, tr: d.trouble, rep: r.pick(H.REPS) };
            M6.forEach((m, i) => { x[m] = Math.max(1000000, r100k(netIn(d.name, '2026-' + day2(5 + i)) * (0.9 + r() * 0.2))); });
            return x;
          });
          all.forEach((x) => {
            if (x.tr) x.Oct = r100k(x.May * r.pick([0.3, 0.45, 0.55]));
            clear(x, 'Oct', 'May', 0.7, 0.02);
            x.dpd = x.tr ? r.pick([61, 75, 94, 95]) : r.pick([0, 0, 0, 0, 0, 4, 9, 15]);
          });
          const drop = (x) => x.Oct <= 0.7 * x.May;
          const normal = all.filter((x) => !x.tr);
          /* Planted: a sales drop with DPD exactly 30 (OK), and one behind at the bank with steady sales (OK). */
          let flipA = normal.find(drop);
          if (!flipA) { flipA = normal[0]; flipA.Oct = r100k(flipA.May * 0.55); }
          flipA.dpd = 30;
          const steady = normal.filter((x) => x !== flipA && !drop(x));
          if (steady.length) r.pick(steady).dpd = r.pick([45, 60]);
          const flipB = all.find((x) => x.tr);
          const S = H.table([{ h: 'Distributor', k: 'd' }, { h: 'Region', k: 'region' }, { h: 'Sales rep', k: 'rep' }], all);
          const sales = r.shuffle(all);
          const SL = H.table([{ h: 'Distributor', k: 'd' }].concat(M6.map((m) => ({ h: 'Net ' + m + ' (Rp)', k: m, num: true }))), sales);
          const book = r.shuffle(all).map((x, i) => { const limit = r.pick([100, 150, 200]) * 1000000; return { f: 'F' + (301 + i), d: x.d, limit, out: r100k(limit * r.pick([0.4, 0.6, 0.8, 0.95])), dpd: x.dpd }; });
          const PJ = H.table([{ h: 'Facility', k: 'f' }, { h: 'Distributor', k: 'd' }, { h: 'Limit (Rp)', k: 'limit', num: true }, { h: 'Outstanding (Rp)', k: 'out', num: true }, { h: 'DPD', k: 'dpd', num: true }], book);
          const sl = sheetOf('penjualan_2026.xlsx', 'Bulanan', SL), pj = sheetOf('buku_pinjaman.xlsx', 'Pinjaman', PJ);
          /* The hidden test: the distributor at DPD 30 slips to 45, and one troubled distributor's October recovers. */
          pj.alt = H.edit(PJ.cells, { [PJ.cell('DPD', at(book, flipA.d))]: 45 });
          sl.alt = H.edit(SL.cells, { [SL.cell('Net Oct (Rp)', at(sales, flipB.d))]: r100k(flipB.May * 0.9) });
          const sk = sl.x.A('Distributor'), oc = sl.x.A('Net Oct (Rp)'), my = sl.x.A('Net May (Rp)'), pk = pj.x.A('Distributor'), dp = pj.x.A('DPD');
          const ST = sl.x.T('Distributor', 'Net Oct (Rp)'), PT = pj.x.T('Distributor', 'DPD');
          const xo = `XLOOKUP(A5,${sk},${oc})`, xm = `XLOOKUP(A5,${sk},${my})`, xd = `XLOOKUP(A5,${pk},${dp})`;
          const f = S.next + S.first, l = S.next + S.last;
          return multi({
            title: 'Early-warning flag from the sales file and the loan book', file: 'peringatan_dini.xlsx', sheetName: 'Daftar', sheet: S, base: S.cells, alt: S.cells, fill: true, ansLabel: 'Flag', open: [sl, pj],
            statement: `<p>Rasa Kenari and Bank Kenari share an early-warning list. Three workbooks are open: the list itself (<b>peringatan_dini.xlsx</b>), Rasa Kenari's monthly net sales per distributor from May to October (<b>penjualan_2026.xlsx</b>, sheet Bulanan) and Bank Kenari's loan book with each distributor's days past due (<b>buku_pinjaman.xlsx</b>, sheet Pinjaman). The two other files list the distributors in their own order.</p><p>Flag <b>Watch</b> when <b>both</b> signals fire: <b>Net Oct</b> is at most 70% of <b>Net May</b> (a fall of 30% or more) <b>and</b> the bank's <b>DPD</b> is more than 30. Otherwise <b>OK</b>. In <b>${f}</b>, return the flag and fill it down to <b>${l}</b>.</p>`,
            solution: `=IF(AND(${xo}<=0.7*${xm},${xd}>30),"Watch","OK")`,
            alts: [`=IF(AND(VLOOKUP(A5,${ST},7,FALSE)<=0.7*VLOOKUP(A5,${ST},2,FALSE),VLOOKUP(A5,${PT},4,FALSE)>30),"Watch","OK")`,
              `=IF(AND(${xo}/${xm}<=0.7,${xd}>30),"Watch","OK")`,
              `=IF(OR(${xo}>0.7*${xm},${xd}<=30),"OK","Watch")`,
              `=IF(AND(INDEX(${oc},MATCH(A5,${sk},0))<=0.7*INDEX(${my},MATCH(A5,${sk},0)),INDEX(${dp},MATCH(A5,${pk},0))>30),"Watch","OK")`],
            hints: ['Three facts, from two files: the May and October sales from penjualan_2026.xlsx and the DPD from buku_pinjaman.xlsx, each a lookup of the distributor in A5.',
              `${xo} is October's sales. The sales test is that, <=0.7 times May's. Both tests must hold, so join them with AND.`,
              'Shape: =IF(AND(XLOOKUP(A5, ..., Net Oct)<=0.7*XLOOKUP(A5, ..., Net May), XLOOKUP(A5, ..., DPD)>30), "Watch", "OK")'],
            explain: 'Each file lists the distributors in its own order, so XLOOKUP matches each name instead of reading the row beside it. AND fires only when the sales fall and the arrears agree, which keeps a one-off dip or a short delay off the list. A DPD of exactly 30 is not more than 30.',
            chips: ['A5', sk, oc, my, pk, dp, '<=0.7*', '>30', '"Watch"', '"OK"'], fnKey: 'IF('
          });
        },
        () => {
          /* Red, amber or green: DPD trend on the bank's own facility list, sales from Rasa Kenari's file, and a sales floor
             that the risk policy sets per region. */
          const trouble = (name) => DIST().find((d) => d.name === name).trouble;
          const snap = (id, m) => H.db().snapshots.find((s) => s.facility_id === id && s.month === m).dpd;
          const rows = H.db().facilities.map((f) => ({ f: f.facility_id, d: f.distributor, region: f.region, may: trouble(f.distributor) ? snap(f.facility_id, '2026-05') : r.pick([0, 0, 0, 5, 10]), oct: trouble(f.distributor) ? snap(f.facility_id, '2026-10') : r.pick([0, 0, 0, 0, 5, 10, 15]) }));
          const FL = {};
          H.REGIONS.forEach((g) => { FL[g] = r.pick([70, 75, 80]); });
          /* The hidden test tightens one region's floor to 85; one steady distributor there sits at 82.5% of May. */
          const tight = r.pick(DIST().filter((d) => !d.trouble)).name, upReg = regionOf(tight);
          const FL2 = Object.assign({}, FL, { [upReg]: 85 });
          const side = r.shuffle(DIST()).map((d) => {
            const x = { d: d.name, may: r100k(netIn(d.name, '2026-05') * (0.9 + r() * 0.2)), oct: r100k(netIn(d.name, '2026-10') * (0.9 + r() * 0.2)) };
            if (d.trouble) x.oct = r100k(x.may * r.pick([0.35, 0.5, 0.6]));
            if (d.name === tight) x.oct = r100k(x.may * 0.825);
            clear(x, 'oct', 'may', [FL[d.region] / 100, FL2[d.region] / 100], 0.01);
            return x;
          });
          const sd = (name) => side.find((x) => x.d === name);
          const s2 = (name) => sd(name).oct < FL[regionOf(name)] / 100 * sd(name).may;
          /* Planted: a Green facility that turns Amber later, and an Amber one (DPD rose, sales steady) that turns Red. */
          const calm = rows.filter((x) => !trouble(x.d) && !s2(x.d) && x.d !== tight);
          const green = calm[0] || rows.find((x) => !trouble(x.d) && x.d !== tight);
          green.oct = green.may;
          if (s2(green.d)) sd(green.d).oct = sd(green.d).may;
          const amber = calm.find((x) => x !== green && x.d !== green.d) || rows.find((x) => !trouble(x.d) && x.d !== green.d && x.d !== tight);
          amber.oct = amber.may + r.pick([5, 10]);
          if (s2(amber.d)) sd(amber.d).oct = sd(amber.d).may;
          const S = H.table([{ h: 'Facility', k: 'f' }, { h: 'Distributor', k: 'd' }, { h: 'Region', k: 'region' }, { h: 'DPD May', k: 'may', num: true }, { h: 'DPD Oct', k: 'oct', num: true }], rows);
          const SL = H.table([{ h: 'Distributor', k: 'd' }, { h: 'Net May (Rp)', k: 'may', num: true }, { h: 'Net Oct (Rp)', k: 'oct', num: true }], side);
          const amb = r.shuffle(H.REGIONS).map((g) => ({ g, fl: FL[g] }));
          const AM = H.table([{ h: 'Region', k: 'g' }, { h: 'Sales floor (%)', k: 'fl', num: true }], amb);
          const sl = sheetOf('penjualan_2026.xlsx', 'Bulanan', SL), am = sheetOf('kebijakan_risiko.xlsx', 'Ambang', AM);
          /* The hidden test: the Green facility's DPD rises, the Amber distributor's October falls to 60% of May, and one
             region's floor is tightened. */
          const alt = H.edit(S.cells, { [S.cell('DPD Oct', rows.indexOf(green))]: green.may + 12 });
          sl.alt = H.edit(SL.cells, { [SL.cell('Net Oct (Rp)', side.indexOf(sd(amber.d)))]: r100k(sd(amber.d).may * 0.6) });
          am.alt = H.edit(AM.cells, { [AM.cell('Sales floor (%)', at(amb, upReg, 'g'))]: 85 });
          const K = sl.x.A('Distributor'), MY = sl.x.A('Net May (Rp)'), OC = sl.x.A('Net Oct (Rp)'), RG = am.x.A('Region'), FLr = am.x.A('Sales floor (%)');
          const fl = `XLOOKUP(C5,${RG},${FLr})`;
          const fell = `XLOOKUP(B5,${K},${OC})<${fl}/100*XLOOKUP(B5,${K},${MY})`;
          const ST = sl.x.T('Distributor', 'Net Oct (Rp)'), AT = am.x.T('Region', 'Sales floor (%)');
          const vfell = `VLOOKUP(B5,${ST},3,FALSE)<VLOOKUP(C5,${AT},2,FALSE)/100*VLOOKUP(B5,${ST},2,FALSE)`;
          const ifell = `INDEX(${OC},MATCH(B5,${K},0))<INDEX(${FLr},MATCH(C5,${RG},0))/100*INDEX(${MY},MATCH(B5,${K},0))`;
          const f = S.next + S.first, l = S.next + S.last;
          return multi({
            title: 'Red, amber or green with a sales floor set per region', file: 'pantauan_fasilitas.xlsx', sheetName: 'Fasilitas', sheet: S, base: S.cells, alt, fill: true, ansLabel: 'Signal', open: [sl, am],
            statement: `<p>Dimas grades every Bank Kenari facility for the monthly risk meeting. Three workbooks are open: his facility list with the days past due in May and October (<b>pantauan_fasilitas.xlsx</b>), Rasa Kenari's net sales per distributor (<b>penjualan_2026.xlsx</b>, sheet Bulanan) and the risk policy (<b>kebijakan_risiko.xlsx</b>, sheet Ambang), which sets a <b>Sales floor (%)</b> for each region.</p><p>Sign 1: the DPD rose (<b>DPD Oct</b> is more than <b>DPD May</b>). Sign 2: the distributor's <b>Net Oct</b> is less than its region's floor, as a percentage of its <b>Net May</b> (a floor of 75 means less than 75% of May). Return <b>Red</b> when both signs show, <b>Amber</b> when one shows, and <b>Green</b> when neither does. Start in <b>${f}</b> and fill down to <b>${l}</b>.</p>`,
            solution: `=IF(AND(E5>D5,${fell}),"Red",IF(OR(E5>D5,${fell}),"Amber","Green"))`,
            alts: [`=LET(s,${fell},d,E5>D5,IF(AND(s,d),"Red",IF(OR(s,d),"Amber","Green")))`,
              `=CHOOSE((E5>D5)+(${fell})+1,"Green","Amber","Red")`,
              `=IFS(AND(E5>D5,${vfell}),"Red",OR(E5>D5,${vfell}),"Amber",TRUE,"Green")`,
              `=IF(AND(E5>D5,${ifell}),"Red",IF(OR(E5>D5,${ifell}),"Amber","Green"))`],
            hints: ['Sign 1 is on the row: E5>D5. Sign 2 needs three lookups: the distributor\'s two months of sales by B5 in the sales file, and the floor by the region in C5 in the policy file.',
              `Sign 2 is ${fell}. Test AND of the two signs for Red first, then OR of them for Amber.`,
              'Shape: =IF(AND(sign1, sign2), "Red", IF(OR(sign1, sign2), "Amber", "Green")). LET can name the two signs so each is written once.'],
            explain: 'The floor depends on the region and the sales on the distributor, so the two lookups use different keys from the same row. Dividing the floor by 100 turns 75 into 75%. AND catches the facilities showing both signs, and only the rest reach OR. A distributor with two facilities gets the same sales sign on both rows, because the lookup goes by name.',
            chips: ['B5', 'C5', 'E5>D5', K, MY, OC, RG, FLr, '/100*', '"Red"', '"Amber"', '"Green"'], fnKey: 'IF('
          });
        },
        () => {
          /* Pak Anton's route: monthly totals from raw sales lines in one file (SUMIFS by month), the DPD from the loan book. */
          const ds = r.sample(NAMES(), 6);
          /* Planted, by position in ds: [October as a share of September, DPD]. Sales drop alone; DPD alone; both; neither (it
             drops on the hidden test); a share slightly above 80% (its DPD rises on the hidden test); neither, at DPD exactly 30. */
          const plan = [[r.pick([0.55, 0.62, 0.68]), r.pick([0, 9, 14])], [r.pick([1.02, 1.08, 1.15]), r.pick([45, 61])], [r.pick([0.45, 0.6]), r.pick([75, 95])],
            [r.pick([0.9, 0.95]), 0], [r.pick([0.84, 0.86]), r.pick([0, 12])], [r.pick([1, 1.05, 1.1]), 30]];
          const lines = [], octOf = {};
          ds.forEach((d, k) => {
            const sep = nums(r.int(1, 2)).map(() => r.int(40, 160) * 100000);
            const target = r100k(sep.reduce((a, x) => a + x, 0) * plan[k][0]);
            const n = k === 3 ? 1 : r.int(1, 2), first = n === 2 ? r100k(target * r.pick([0.3, 0.45, 0.6])) : 0;
            sep.forEach((x) => lines.push({ m: 'Sep', d, net: x }));
            if (n === 2) lines.push({ m: 'Oct', d, net: first });
            lines.push(octOf[d] = { m: 'Oct', d, net: target - first });
          });
          const sepTotal = (d) => lines.filter((x) => x.m === 'Sep' && x.d === d).reduce((a, x) => a + x.net, 0);
          const sales = r.shuffle(lines.filter((x) => x.m === 'Sep')).concat(r.shuffle(lines.filter((x) => x.m === 'Oct'))).map((x, i) => Object.assign(x, { inv: 'INV-' + (1801 + i * 3 + r.int(0, 2)) }));
          const DAYS = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
          const route = ds.slice().sort().map((d, i) => ({ d, city: DIST().find((x) => x.name === d).city, day: DAYS[i] }));
          const S = H.table([{ h: 'Distributor', k: 'd' }, { h: 'City', k: 'city' }, { h: 'Visit day', k: 'day' }], route);
          const SL = H.table([{ h: 'Month', k: 'm' }, { h: 'Invoice', k: 'inv' }, { h: 'Distributor', k: 'd' }, { h: 'Net (Rp)', k: 'net', num: true }], sales);
          const others = r.sample(NAMES().filter((n) => !ds.includes(n)), 4);
          const book = r.shuffle(ds.map((d, k) => ({ d, dpd: plan[k][1] })).concat(others.map((d) => ({ d, dpd: r.pick([0, 0, 15, 45]) })))).map((x, i) => Object.assign(x, { f: 'F' + (501 + i), out: r.int(20, 120) * 1000000 }));
          const PJ = H.table([{ h: 'Facility', k: 'f' }, { h: 'Distributor', k: 'd' }, { h: 'Outstanding (Rp)', k: 'out', num: true }, { h: 'DPD', k: 'dpd', num: true }], book);
          const sl = sheetOf('penjualan_sep_okt.xlsx', 'Penjualan', SL), pj = sheetOf('buku_pinjaman.xlsx', 'Pinjaman', PJ);
          /* The hidden test: a steady distributor's October line comes in much lower, and the one slightly above 80% falls 40
             days behind. */
          sl.alt = H.edit(SL.cells, { [SL.cell('Net (Rp)', sales.indexOf(octOf[ds[3]]))]: r100k(sepTotal(ds[3]) * 0.6) });
          pj.alt = H.edit(PJ.cells, { [PJ.cell('DPD', at(book, ds[4]))]: 40 });
          const MO = sl.x.A('Month'), DI = sl.x.A('Distributor'), NE = sl.x.A('Net (Rp)'), K = pj.x.A('Distributor'), DP = pj.x.A('DPD'), TA = pj.x.T('Distributor', 'DPD');
          const oct = `SUMIFS(${NE},${DI},A5,${MO},"Oct")`, sep = `SUMIFS(${NE},${DI},A5,${MO},"Sep")`, xd = `XLOOKUP(A5,${K},${DP})`;
          const f = S.next + S.first, l = S.next + S.last;
          return multi({
            title: 'Visit list flag from monthly sales lines and bank DPD', file: 'rute_kunjungan.xlsx', sheetName: 'Rute', sheet: S, base: S.cells, alt: S.cells, fill: true, ansLabel: 'Flag', open: [sl, pj],
            statement: `<p>Pak Anton is planning next week's visits to the six distributors on his route. Three workbooks are open: the route (<b>rute_kunjungan.xlsx</b>), the September and October sales lines (<b>penjualan_sep_okt.xlsx</b>, sheet Penjualan, one or more lines per distributor in each month) and Bank Kenari's loan book (<b>buku_pinjaman.xlsx</b>, sheet Pinjaman) with each distributor's days past due.</p><p>In <b>${f}</b>, return <b>Watch</b> when the distributor's total October <b>Net</b> sales are less than 80% of its total September net sales, <b>or</b> its <b>DPD</b> is more than 30. Otherwise return <b>OK</b>. Fill it down to <b>${l}</b>.</p>`,
            solution: `=IF(OR(${oct}<0.8*${sep},${xd}>30),"Watch","OK")`,
            alts: [`=IF(OR(SUMPRODUCT((${DI}=A5)*(${MO}="Oct")*${NE})<0.8*SUMPRODUCT((${DI}=A5)*(${MO}="Sep")*${NE}),VLOOKUP(A5,${TA},3,FALSE)>30),"Watch","OK")`,
              `=IF(AND(${oct}>=0.8*${sep},${xd}<=30),"OK","Watch")`,
              `=LET(o,${oct},s,${sep},d,${xd},IF(OR(o<0.8*s,d>30),"Watch","OK"))`,
              `=IF(OR(${oct}<0.8*${sep},INDEX(${DP},MATCH(A5,${K},0))>30),"Watch","OK")`],
            hints: ['Each month\'s total is a SUMIFS over the sales lines in penjualan_sep_okt.xlsx with two conditions: the distributor in A5 and the month. The DPD is one lookup in the loan book.',
              `${oct} is October's total. Compare it with 0.8 times September's, and test ${xd}>30, inside OR.`,
              'Shape: =IF(OR(SUMIFS(net, distributor, A5, month, "Oct") < 0.8*SUMIFS(net, distributor, A5, month, "Sep"), XLOOKUP(A5, ..., DPD)>30), "Watch", "OK")'],
            explain: 'SUMIFS adds each distributor\'s lines month by month, so the test compares whole-month totals rather than single invoices. XLOOKUP reads the bank\'s DPD from the loan book, which lists more distributors than the route, and OR raises Watch on either signal. A DPD of exactly 30 is not more than 30.',
            chips: ['A5', NE, DI, MO, '"Oct"', '"Sep"', K, DP, '0.8*', '"Watch"', '"OK"'], fnKey: 'IF('
          });
        },
        () => {
          /* Bank facilities (main), limits from the limit file, overdue trade debt from Rasa Kenari's invoice file (SUMIFS on
             days overdue): Watch on high use OR overdue debt. */
          const rows = H.loans(r, 14).map((x) => ({ f: x.facility, d: x.distributor, limit: x.limit, out: x.outstanding }));
          const names = [...new Set(rows.map((x) => x.d))];
          const od = {};
          names.forEach((n) => { od[n] = r.pick([0, 0, 0, 3500000, 8000000, 12500000, 31000000, 48000000]); });
          const use = (x) => x.out / x.limit;
          const quiet = (n) => od[n] <= 25000000 && rows.filter((x) => x.d === n).every((x) => use(x) < 0.9);
          /* Planted: a facility at exactly 90% of its limit, a distributor owing exactly Rp 25,000,000 overdue plus an invoice
             not yet due, and three that change later. */
          const pool = r.shuffle(names);
          const atLimit = rows.find((x) => x.d === pool[0]);
          od[pool[0]] = 0;
          atLimit.out = Math.round(0.9 * atLimit.limit);
          od[pool[1]] = 25000000;
          rows.filter((x) => x.d === pool[1]).forEach((x) => { if (use(x) >= 0.9) x.out = r100k(x.limit * 0.62); });
          od[pool[2]] = 0;
          rows.filter((x) => x.d === pool[2]).forEach((x) => { if (use(x) >= 0.9) x.out = r100k(x.limit * 0.5); });
          const riser = rows.find((x) => x.d !== pool[2] && x.d !== pool[0] && quiet(x.d));
          const cutF = rows.find((x) => x !== riser && x.d !== pool[2] && quiet(x.d) && use(x) >= 0.45 && use(x) < 0.85);
          const lines = [];
          names.forEach((n) => {
            let left = od[n];
            if (left > 0) {
              if (left >= 8000000 && r.chance(0.6)) { const a = r100k(left * r.pick([0.3, 0.45, 0.6])); lines.push({ d: n, amt: a, days: r.int(5, 60) }); left -= a; }
              lines.push({ d: n, amt: left, days: r.int(1, 90) });
            }
            if (n === pool[1]) lines.push({ d: n, amt: r.int(20, 80) * 100000, days: 0 });
            else if (n === pool[2]) lines.push({ d: n, amt: 30000000, days: 0 });
            else if (r.chance(0.4)) lines.push({ d: n, amt: r.int(10, 150) * 100000, days: 0 });
          });
          const inv = r.shuffle(lines).map((x, i) => Object.assign(x, { inv: 'INV-' + (2401 + i * 3 + r.int(0, 2)) }));
          const S = H.table([{ h: 'Facility', k: 'f' }, { h: 'Distributor', k: 'd' }, { h: 'Outstanding (Rp)', k: 'out', num: true }], rows);
          const lims = r.shuffle(rows).map((x) => ({ f: x.f, limit: x.limit, on: '2025-' + day2(r.int(1, 12)) + '-' + day2(r.int(1, 28)) }));
          const LM = H.table([{ h: 'Facility', k: 'f' }, { h: 'Limit (Rp)', k: 'limit', num: true }, { h: 'Approved on', k: 'on', date: true }], lims);
          const PI = H.table([{ h: 'Invoice', k: 'inv' }, { h: 'Distributor', k: 'd' }, { h: 'Amount (Rp)', k: 'amt', num: true }, { h: 'Days overdue', k: 'days', num: true }], inv);
          const lm = sheetOf('limit_kredit.xlsx', 'Limit', LM), pi = sheetOf('piutang_dagang.xlsx', 'Piutang', PI);
          /* The hidden test: a week later, one facility is drawn to 95%, the bank cuts another's limit, and an invoice of
             Rp 30,000,000 falls past its due date. */
          const alt = riser ? H.edit(S.cells, { [S.cell('Outstanding (Rp)', rows.indexOf(riser))]: r100k(riser.limit * 0.95) }) : S.cells;
          lm.alt = cutF ? H.edit(LM.cells, { [LM.cell('Limit (Rp)', at(lims, cutF.f, 'f'))]: Math.floor(cutF.out / 0.95 / 1000000) * 1000000 }) : LM.cells;
          pi.alt = H.edit(PI.cells, { [PI.cell('Days overdue', inv.findIndex((x) => x.d === pool[2] && x.days === 0))]: r.int(3, 12) });
          const lim = `XLOOKUP(A5,${lm.x.A('Facility')},${lm.x.A('Limit (Rp)')})`;
          const PA = pi.x.A('Amount (Rp)'), PD = pi.x.A('Distributor'), PO = pi.x.A('Days overdue');
          const odue = `SUMIFS(${PA},${PD},B5,${PO},">0")`;
          const f = S.next + S.first, l = S.next + S.last;
          return multi({
            title: 'Watchlist from limit use and overdue invoices in two files', file: 'daftar_pantauan.xlsx', sheetName: 'Fasilitas', sheet: S, base: S.cells, alt, fill: true, ansLabel: 'Flag', open: [lm, pi],
            statement: `<p>Citra and Bu Lestari compare notes each month. Three workbooks are open: Bank Kenari's facilities with what each one owes (<b>daftar_pantauan.xlsx</b>; a distributor can hold more than one), the approved limits by facility (<b>limit_kredit.xlsx</b>, sheet Limit) and Rasa Kenari's open invoices (<b>piutang_dagang.xlsx</b>, sheet Piutang), where <b>Days overdue</b> is 0 for an invoice not yet due.</p><p>Flag <b>Watch</b> when the facility's <b>Outstanding</b> is at least 90% of its <b>Limit</b>, <b>or</b> the distributor's invoices that are overdue (Days overdue more than 0) add up to more than ${rp(25000000)}. Otherwise <b>OK</b>. In <b>${f}</b>, return the flag and fill it down to <b>${l}</b>.</p>`,
            solution: `=IF(OR(C5>=0.9*${lim},${odue}>25000000),"Watch","OK")`,
            alts: [`=IF(OR(C5/VLOOKUP(A5,${lm.x.T('Facility', 'Limit (Rp)')},2,FALSE)>=0.9,${odue}>25000000),"Watch","OK")`,
              `=IF(AND(C5<0.9*${lim},SUMPRODUCT((${PD}=B5)*(${PO}>0)*${PA})<=25000000),"OK","Watch")`,
              `=IF(OR(C5>=0.9*INDEX(${lm.x.A('Limit (Rp)')},MATCH(A5,${lm.x.A('Facility')},0)),${odue}>25000000),"Watch","OK")`],
            hints: ['The limit is one value per facility, so look up the facility in A5 in limit_kredit.xlsx. The overdue debt is several invoices per distributor, so add them with SUMIFS on the distributor in B5.',
              `${lim} is the limit. ${odue} adds only the invoices past their due date, which is what the ">0" condition does.`,
              'Shape: =IF(OR(C5>=0.9*limit, SUMIFS(amount, distributor, B5, days overdue, ">0")>25000000), "Watch", "OK")'],
            explain: 'A line drawn close to its limit and unpaid trade invoices are two early signs of a cash squeeze, from two different companies. XLOOKUP brings the limit by facility, and SUMIFS adds the distributor\'s overdue invoices while leaving out the ones not yet due. A facility at exactly 90% counts, and exactly Rp 25,000,000 overdue does not.',
            chips: ['A5', 'B5', 'C5>=0.9*', lm.x.A('Facility'), lm.x.A('Limit (Rp)'), PA, PD, PO, '">0"', '>25000000', '"Watch"', '"OK"'], fnKey: 'IF('
          });
        }
      ][v]();
    }
  });

  /* ======================= Margins and price effects: SUMPRODUCT and SUMIFS over another workbook ======================= */

  P.add({
    key: 'weighted-average-cases', difficulty: 'Hard', level: 8, topics: ['SUMPRODUCT', 'SUMIFS', 'XLOOKUP', 'weighted average', 'Several workbooks'], count: 3, multi: true,
    make(r, v) {
      return [
        () => {
          /* Net interest margin per loan type: the rate weighted by outstanding over the loan book, less the cost of funds
             from the treasury file. */
          const TY = ['Working Capital', 'Investment', 'Trade Finance', 'Term Loan'];
          const rows = H.loans(r, 16).map((x) => ({ f: x.facility, d: x.distributor, out: x.outstanding, rate: x.rate }));
          r.shuffle(nums(16)).forEach((k, j) => { rows[k].type = TY[j % 4]; });
          const COF = { 'Working Capital': r.pick([5.5, 5.75, 6]), Investment: r.pick([6.25, 6.5]), 'Trade Finance': r.pick([5, 5.25]), 'Term Loan': r.pick([6, 6.25, 6.5]) };
          const SEG = { 'Working Capital': 'Commercial', Investment: 'Corporate', 'Trade Finance': 'Trade', 'Term Loan': 'Commercial' };
          const S = H.table([{ h: 'Loan type', k: 't' }, { h: 'Segment', k: 's' }], TY.map((t) => ({ t, s: SEG[t] })));
          const PJ = H.table([{ h: 'Facility', k: 'f' }, { h: 'Distributor', k: 'd' }, { h: 'Type', k: 'type' }, { h: 'Outstanding (Rp)', k: 'out', num: true }, { h: 'Rate (% a year)', k: 'rate', num: true }], rows);
          const dana = r.shuffle(TY).map((t) => ({ t, cof: COF[t] }));
          const DN = H.table([{ h: 'Loan type', k: 't' }, { h: 'Cost of funds (% a year)', k: 'cof', num: true }], dana);
          const pj = sheetOf('buku_pinjaman.xlsx', 'Pinjaman', PJ), dn = sheetOf('biaya_dana.xlsx', 'Dana', DN);
          /* The hidden test: a corrected balance and a repriced facility in the loan book, and a higher cost of funds. */
          const k1 = rows.findIndex((x) => x.type === TY[0]), k2 = rows.findIndex((x) => x.type === TY[2]);
          pj.alt = H.edit(PJ.cells, { [PJ.cell('Outstanding (Rp)', k1)]: (x) => x + 20000000, [PJ.cell('Rate (% a year)', k2)]: (x) => x + 0.5 });
          dn.alt = H.edit(DN.cells, { [DN.cell('Cost of funds (% a year)', at(dana, TY[3], 't'))]: (x) => x + 0.25 });
          const T = pj.x.A('Type'), O = pj.x.A('Outstanding (Rp)'), R = pj.x.A('Rate (% a year)'), LT = dn.x.A('Loan type'), CF = dn.x.A('Cost of funds (% a year)');
          const cf = `XLOOKUP(A5,${LT},${CF})`;
          const f = S.next + S.first, l = S.next + S.last;
          return multi({
            title: 'Net interest margin per loan type from two bank files', file: 'ringkasan_portofolio.xlsx', sheetName: 'Ringkasan', sheet: S, base: S.cells, alt: S.cells, fill: true, ansLabel: 'Net margin (%)', precision: 4, open: [pj, dn],
            statement: `<p>Dimas reports what the bank earns on each loan type after paying for the money it lends. Three workbooks are open: his summary (<b>ringkasan_portofolio.xlsx</b>), Bank Kenari's loan book (<b>buku_pinjaman.xlsx</b>, sheet Pinjaman, one row per facility with its <b>Type</b>, <b>Outstanding</b> and <b>Rate</b>) and the treasury's cost of funds per loan type (<b>biaya_dana.xlsx</b>, sheet Dana).</p><p>In <b>${f}</b>, work out the net margin for the loan type in A5: the average <b>Rate (% a year)</b> of that type weighted by Outstanding (the sum of Outstanding × Rate for the type, divided by the type's total Outstanding), minus that type's <b>Cost of funds</b>. Fill it down to <b>${l}</b>. Leave it unrounded; it is checked to 4 decimals.</p>`,
            solution: `=SUMPRODUCT((${T}=A5)*${O}*${R})/SUMIFS(${O},${T},A5)-${cf}`,
            alts: [`=SUMPRODUCT(--(${T}=A5),${O},${R})/SUMIFS(${O},${T},A5)-VLOOKUP(A5,${dn.x.T('Loan type', 'Cost of funds (% a year)')},2,FALSE)`,
              `=LET(w,(${T}=A5)*${O},SUMPRODUCT(w,${R})/SUM(w)-${cf})`,
              `=SUMPRODUCT((${T}=A5)*${O},${R})/SUMPRODUCT((${T}=A5)*${O})-INDEX(${CF},MATCH(A5,${LT},0))`],
            hints: ['A weighted average is the sum of weight times value, divided by the sum of the weights. Here the weight is Outstanding and the value is Rate, both in the loan book; (Type=A5) keeps only the rows of this row\'s type.',
              `SUMPRODUCT((${T}=A5)*${O}*${R}) is the top and SUMIFS(${O}, ${T}, A5) the bottom. Then take away ${cf}.`,
              'Shape: =SUMPRODUCT((type=A5)*outstanding*rate)/SUMIFS(outstanding, type, A5) - XLOOKUP(A5, loan type, cost of funds)'],
            explain: 'The test (Type=A5) zeroes out every other loan type, so SUMPRODUCT adds Outstanding times Rate for this type only, and dividing by its total Outstanding gives a rate where a large facility counts for more. XLOOKUP brings the matching cost of funds from the treasury file, whose rows are in another order. The ranges in the loan book are pinned with $, so filling down keeps them in place while A5 moves to the next type.',
            chips: ['A5', T, O, R, LT, CF, '(' + T + '=A5)', 'SUMIFS(', '/'], fnKey: 'SUMPRODUCT('
          });
        },
        () => {
          /* What the 2027 price list adds at last quarter's volume: the new price from the price list, the current price from
             the product master, and the cartons from the sales lines (SUMIFS): four workbooks. */
          const PR = H.PRODUCTS;
          const nw = {};
          PR.forEach((p) => { nw[p.sku] = Math.round(p.carton_price * (1 + r.pick([-0.03, 0, 0.02, 0.04, 0.05, 0.07, 0.1])) / 1000) * 1000; });
          const q3 = H.db().sales.filter((s) => s.date >= '2026-07-01' && s.date <= '2026-09-30');
          const lines = r.sample(q3, 22).sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : 0)).map((s) => ({ date: s.date, inv: s.invoice, d: s.distributor, sku: s.sku, cartons: s.cartons }));
          const sold = [...new Set(lines.map((x) => x.sku))];
          const S = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Product', k: 'name' }], r.shuffle(PR));
          const PM = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Product', k: 'name' }, { h: 'Category', k: 'category' }, { h: 'Carton price (Rp)', k: 'carton_price', num: true }], PR);
          const harga = r.shuffle(PR).map((p) => ({ sku: p.sku, nw: nw[p.sku] }));
          const HG = H.table([{ h: 'SKU', k: 'sku' }, { h: 'New price (Rp)', k: 'nw', num: true }], harga);
          const PQ = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'Invoice', k: 'inv' }, { h: 'Distributor', k: 'd' }, { h: 'SKU', k: 'sku' }, { h: 'Cartons', k: 'cartons', num: true }], lines);
          const pm = sheetOf('master_produk.xlsx', 'Produk', PM), hg = sheetOf('harga_2027.xlsx', 'Harga', HG), pq = sheetOf('penjualan_q3.xlsx', 'Penjualan', PQ);
          /* The hidden test: one new price is revised, one current price is corrected in the master, and one sales line's
             cartons are corrected. */
          const [s1, s2] = r.sample(sold, 2);
          hg.alt = H.edit(HG.cells, { [HG.cell('New price (Rp)', at(harga, s1, 'sku'))]: (x) => x + 4000 });
          pm.alt = H.edit(PM.cells, { [PM.cell('Carton price (Rp)', at(PR, s2, 'sku'))]: (x) => x + 2000 });
          pq.alt = H.edit(PQ.cells, { [PQ.cell('Cartons', r.int(0, lines.length - 1))]: (x) => x + 10 });
          const HS = hg.x.A('SKU'), HN = hg.x.A('New price (Rp)'), MS = pm.x.A('SKU'), MP = pm.x.A('Carton price (Rp)'), QS = pq.x.A('SKU'), QC = pq.x.A('Cartons');
          const xn = `XLOOKUP(A5,${HS},${HN})`, xo = `XLOOKUP(A5,${MS},${MP})`, qty = `SUMIFS(${QC},${QS},A5)`;
          const f = S.next + S.first, l = S.next + S.last;
          return multi({
            title: 'Revenue effect of the 2027 price list at last quarter\'s volume', file: 'dampak_harga.xlsx', sheetName: 'Dampak', sheet: S, base: S.cells, alt: S.cells, fill: true, ansLabel: 'Effect (Rp)', open: [pm, hg, pq],
            statement: `<p>Tika has the new carton price list, and Pak Anton asks what it would have added to last quarter's sales. Four workbooks are open: her summary by SKU (<b>dampak_harga.xlsx</b>), the product master with today's <b>Carton price</b> (<b>master_produk.xlsx</b>, sheet Produk), the 2027 price list (<b>harga_2027.xlsx</b>, sheet Harga) and the July to September sales lines (<b>penjualan_q3.xlsx</b>, sheet Penjualan, one row per order line).</p><p>In <b>${f}</b>, work out the effect for the SKU in A5: its <b>New price</b> minus its current Carton price, times the total <b>Cartons</b> of that SKU in the sales lines. A SKU that did not sell gives 0. Fill it down to <b>${l}</b>.</p>`,
            solution: `=(${xn}-${xo})*${qty}`,
            alts: [`=(VLOOKUP(A5,${hg.x.T('SKU', 'New price (Rp)')},2,FALSE)-VLOOKUP(A5,${pm.x.T('SKU', 'Carton price (Rp)')},4,FALSE))*${qty}`,
              `=(${xn}-${xo})*SUMPRODUCT((${QS}=A5)*${QC})`,
              `=LET(gap,${xn}-INDEX(${MP},MATCH(A5,${MS},0)),gap*${qty})`],
            hints: ['Two prices and one volume, from three files: the new price from harga_2027.xlsx and the current price from master_produk.xlsx (lookups), and the cartons from penjualan_q3.xlsx (a SUMIFS, because a SKU sells on several lines).',
              `${xn}-${xo} is the change per carton. ${qty} is the cartons sold.`,
              'Shape: =(XLOOKUP(A5, ..., New price) - XLOOKUP(A5, ..., Carton price)) * SUMIFS(cartons, SKU, A5)'],
            explain: 'Each lookup finds the SKU in its own file, whatever the order, and the brackets take the difference before multiplying. SUMIFS adds every line of the SKU in the quarter, so the effect follows the real volume. A price cut gives a negative effect, and a SKU with no sales gives 0.',
            chips: ['A5', HS, HN, MS, MP, QS, QC, 'SUMIFS(', ')*'], fnKey: 'XLOOKUP('
          });
        },
        () => {
          /* The margin of each category's mix, weighted by net revenue computed inside SUMPRODUCT over the SKU file, against
             the budget margin from the budget file. */
          const MG = { Noodles: [18, 20, 22], Coffee: [26, 28, 30], Sambal: [32, 34, 36], Soap: [24, 26] };
          const sku = r.shuffle(H.PRODUCTS).map((p) => ({ sku: p.sku, cat: p.category, qty: r.int(5, 80) * 10, price: p.carton_price, disc: r.pick([0, 0, 2, 3, 5, 8, 10]), m: r.pick(MG[p.category]) + r.pick([0, 0, 0.5]) }));
          /* Planted: in each category the SKU with the most cartons has the deepest discount, so leaving the discount out
             moves the answer. */
          H.CATEGORIES.forEach((c) => { sku.filter((x) => x.cat === c).reduce((a, x) => (x.qty > a.qty ? x : a)).disc = 12; });
          const BRAND = { Noodles: 'Mi Lekas', Coffee: 'Kopi Sekejap', Sambal: 'Sambal Nyala', Soap: 'Sabun Rimba Wangi' };
          const S = H.table([{ h: 'Category', k: 'c' }, { h: 'Brand', k: 'b' }], H.CATEGORIES.map((c) => ({ c, b: BRAND[c] })));
          const SK = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Category', k: 'cat' }, { h: 'Cartons sold', k: 'qty', num: true }, { h: 'Carton price (Rp)', k: 'price', num: true }, { h: 'Discount (%)', k: 'disc', num: true }, { h: 'Margin (%)', k: 'm', num: true }], sku);
          const bud = r.shuffle(H.CATEGORIES).map((c) => ({ c, m: MG[c][1] + r.pick([-1, 0, 0.5, 1]) }));
          const AG = H.table([{ h: 'Category', k: 'c' }, { h: 'Budget margin (%)', k: 'm', num: true }], bud);
          const sk = sheetOf('ringkasan_sku.xlsx', 'Sku', SK), ag = sheetOf('anggaran_2026.xlsx', 'Anggaran', AG);
          /* The hidden test: a corrected margin, a larger volume and a new discount in the SKU file, and a revised budget. */
          sk.alt = H.edit(SK.cells, { [SK.cell('Margin (%)', 2)]: (x) => x + 2, [SK.cell('Cartons sold', 6)]: (x) => x + 150, [SK.cell('Discount (%)', 4)]: (x) => x + 5 });
          ag.alt = H.edit(AG.cells, { [AG.cell('Budget margin (%)', 1)]: (x) => x - 1 });
          const C = sk.x.A('Category'), Q = sk.x.A('Cartons sold'), Pr = sk.x.A('Carton price (Rp)'), Dc = sk.x.A('Discount (%)'), Mg = sk.x.A('Margin (%)');
          const AC = ag.x.A('Category'), AM = ag.x.A('Budget margin (%)');
          const net = `(${C}=A5)*${Q}*${Pr}*(1-${Dc}/100)`, bm = `XLOOKUP(A5,${AC},${AM})`;
          const f = S.next + S.first, l = S.next + S.last;
          return multi({
            title: 'Mix margin per category against budget, from two files', file: 'laporan_margin.xlsx', sheetName: 'Margin', sheet: S, base: S.cells, alt: S.cells, fill: true, ansLabel: 'Gap to budget (points)', precision: 4, open: [sk, ag],
            statement: `<p>Tika needs last month's margin by category for the board pack. Three workbooks are open: her report (<b>laporan_margin.xlsx</b>), last month's figures per SKU (<b>ringkasan_sku.xlsx</b>, sheet Sku) and the 2026 budget (<b>anggaran_2026.xlsx</b>, sheet Anggaran). A SKU's net revenue is <b>Cartons sold</b> × <b>Carton price</b> × (1 - <b>Discount (%)</b>/100); there is no net revenue column, so work it out inside the formula.</p><p>In <b>${f}</b>, return the category's <b>Margin (%)</b> weighted by net revenue (the sum of net revenue × Margin over the category's SKUs, divided by the sum of their net revenue), minus the category's <b>Budget margin (%)</b>. A negative answer means below budget. Fill it down to <b>${l}</b>. Leave it unrounded; it is checked to 4 decimals.</p>`,
            solution: `=SUMPRODUCT(${net}*${Mg})/SUMPRODUCT(${net})-${bm}`,
            alts: [`=SUMPRODUCT(--(${C}=A5),${Q},${Pr},1-${Dc}/100,${Mg})/SUMPRODUCT(--(${C}=A5),${Q},${Pr},1-${Dc}/100)-VLOOKUP(A5,${ag.x.T('Category', 'Budget margin (%)')},2,FALSE)`,
              `=LET(rev,${net},SUMPRODUCT(rev,${Mg})/SUM(rev)-${bm})`,
              `=SUMPRODUCT((${C}=A5)*${Q}*${Pr}*(100-${Dc})*${Mg})/SUMPRODUCT((${C}=A5)*${Q}*${Pr}*(100-${Dc}))-INDEX(${AM},MATCH(A5,${AC},0))`],
            hints: ['The weight is not a column anywhere: it is net revenue, Cartons × Price × (1 - Discount/100), row by row in ringkasan_sku.xlsx, and (Category=A5) keeps only this category\'s SKUs.',
              `${net} is the category's net revenue per SKU. SUMPRODUCT of that times ${Mg} is the top; SUMPRODUCT of it alone is the bottom. Then take away ${bm}.`,
              'Shape: =SUMPRODUCT(net revenue * margin) / SUMPRODUCT(net revenue) - XLOOKUP(A5, budget category, budget margin)'],
            explain: 'The top adds net revenue times margin for the category\'s SKUs and the bottom adds their net revenue, so the result is the margin of the whole category. Leaving out the discount would give a heavily discounted SKU more weight than the money it brings in. XLOOKUP brings the budget from the other file, and the difference says how far the category is above or below plan.',
            chips: ['A5', C, Q, Pr, Dc, Mg, AC, AM, '(1-' + Dc + '/100)', 'SUMPRODUCT('], fnKey: 'SUMPRODUCT('
          });
        }
      ][v]();
    }
  });

  /* ======================= Commission and bonus: sales, targets, visits and rate tables in other files ======================= */

  P.add({
    key: 'commission-rate-tables', difficulty: 'Hard', level: 9, topics: ['XLOOKUP', 'VLOOKUP', 'INDEX', 'MATCH', 'approximate match', 'commission', 'Several workbooks'], count: 3, multi: true,
    make(r, v) {
      return [
        () => {
          /* A progressive table (base plus a rate on the excess) in the rate file, applied to sales from the sales file. */
          const t1 = r.pick([75, 100]) * 1e6, t2 = r.pick([150, 175]) * 1e6, t3 = r.pick([225, 250]) * 1e6;
          const rate = r.pick([[1, 1.5, 2, 3], [0.5, 1, 2, 2.5], [1, 2, 2.5, 3.5]]);
          const from = [0, t1, t2, t3], base = [0];
          for (let k = 1; k < 4; k++) base.push(base[k - 1] + (from[k] - from[k - 1]) * rate[k - 1] / 100);
          const reps = H.reps(r, 10).map((x) => ({ name: x.name, area: x.area, sales: x.actual, inv: r.int(18, 70) }));
          /* Planted: sales exactly on the second tier's From, one in the top tier, one in the third. */
          reps[0].sales = t1;
          reps[1].sales = t3 + r.int(1, 40) * 1e6;
          reps[2].sales = t2 + r.int(1, 40) * 1e6;
          const roster = r.shuffle(reps), sales = r.shuffle(reps);
          const S = H.table([{ h: 'Rep', k: 'name' }, { h: 'Area', k: 'area' }], roster);
          const RS = H.table([{ h: 'Rep', k: 'name' }, { h: 'Invoices', k: 'inv', num: true }, { h: 'Net sales (Rp)', k: 'sales', num: true }], sales);
          const TK = H.table([{ h: 'From (Rp)', k: 'f', num: true }, { h: 'Base (Rp)', k: 'b', num: true }, { h: 'Rate on excess (%)', k: 'rt', num: true }], from.map((f, k) => ({ f, b: base[k], rt: rate[k] })));
          const rs = sheetOf('penjualan_rep.xlsx', 'Ringkasan', RS), tk = sheetOf('tarif_komisi.xlsx', 'Progresif', TK);
          /* The hidden test: the top rate goes up half a point, and one rep's sales cross into the next tier. */
          const mover = reps.find((x) => x.sales < t2 - 20e6 && x.sales > t1);
          tk.alt = H.edit(TK.cells, { [TK.cell('Rate on excess (%)', 3)]: rate[3] + 0.5 });
          rs.alt = mover ? H.edit(RS.cells, { [RS.cell('Net sales (Rp)', sales.indexOf(mover))]: t2 + 5e6 }) : RS.cells;
          const F = tk.x.A('From (Rp)'), Bs = tk.x.A('Base (Rp)'), Rt = tk.x.A('Rate on excess (%)'), TA = tk.x.T('From (Rp)', 'Rate on excess (%)');
          const s = `XLOOKUP(A5,${rs.x.A('Rep')},${rs.x.A('Net sales (Rp)')})`, vs = `VLOOKUP(A5,${rs.x.T('Rep', 'Net sales (Rp)')},3,FALSE)`;
          const xl = (x, col) => `XLOOKUP(${x},${F},${col},,-1)`;
          const f = S.next + S.first, l = S.next + S.last;
          return multi({
            title: 'Progressive commission with sales and rates in other files', file: 'komisi_progresif.xlsx', sheetName: 'Komisi', sheet: S, base: S.cells, alt: S.cells, fill: true, ansLabel: 'Commission (Rp)', open: [rs, tk],
            statement: `<p>Rasa Kenari pays sales commission like a tax table, so a rep never earns less by selling more. Three workbooks are open: the commission sheet with the reps (<b>komisi_progresif.xlsx</b>), each rep's Q4 <b>Net sales</b> (<b>penjualan_rep.xlsx</b>, sheet Ringkasan) and the rate table (<b>tarif_komisi.xlsx</b>, sheet Progresif). In the rate table each row starts at <b>From</b>; sales that reach a row (at least its From) get that row's <b>Base</b>, plus its <b>Rate on excess</b> on the part of sales above its From.</p><p>In <b>${f}</b>, work out the rep's commission: Base + (Net sales - From) × Rate / 100, all from the row the sales fall in, rounded to whole rupiah with ROUND. Fill it down to <b>${l}</b>.</p>`,
            solution: `=LET(s,${s},ROUND(${xl('s', Bs)}+(s-${xl('s', F)})*${xl('s', Rt)}/100,0))`,
            alts: [`=ROUND(VLOOKUP(${vs},${TA},2,TRUE)+(${vs}-VLOOKUP(${vs},${TA},1,TRUE))*VLOOKUP(${vs},${TA},3,TRUE)/100,0)`,
              `=LET(s,INDEX(${rs.x.A('Net sales (Rp)')},MATCH(A5,${rs.x.A('Rep')},0)),k,MATCH(s,${F},1),ROUND(INDEX(${Bs},k)+(s-INDEX(${F},k))*INDEX(${Rt},k)/100,0))`,
              `=ROUND(${xl(s, Bs)}+(${s}-${xl(s, F)})*${xl(s, Rt)}/100,0)`],
            hints: ['First the sales: an exact lookup of the rep in A5 in penjualan_rep.xlsx. Then three numbers from the same row of the rate table, the last From the sales reach: an approximate match.',
              `LET(s, ${s}, ...) names the sales once. ${xl('s', Bs)} is the Base; swap the last range for ${F} or ${Rt} to get the From and the Rate. match_mode -1 means exact or next smaller.`,
              'Shape: =LET(s, sales, ROUND(Base + (s - From) * Rate / 100, 0)), with Base, From and Rate each an XLOOKUP(s, From column, ..., , -1).'],
            explain: 'The exact lookup brings the rep\'s sales across, and the approximate lookups pick the tier those sales fall in. The formula charges that tier\'s rate only on the part above its starting point; the Base already holds the commission earned on the tiers below, so the total rises smoothly. LET saves writing the sales lookup four times.',
            chips: ['A5', rs.x.A('Rep'), rs.x.A('Net sales (Rp)'), F, Bs, Rt, ',,-1)', 'LET(', 'ROUND('], fnKey: 'XLOOKUP('
          });
        },
        () => {
          /* A two-way rate grid in the rate file: the category (looked up from the product master by SKU) picks the row, the
             cartons pick the column (approximate). */
          const rows = H.sales(r, 14).map((s) => ({ invoice: s.invoice, sku: s.sku, cat: s.category, cartons: s.cartons, price: s.price, disc: s.discount_pct }));
          const th = r.pick([[0, 20, 40], [0, 15, 30], [0, 25, 50]]);
          const RT = { Noodles: [1, 1.5, 2], Coffee: [1.5, 2, 3], Sambal: [2, 2.5, 3.5], Soap: [1, 2, 2.5] };
          const cats = r.shuffle(H.CATEGORIES);
          const vals = cats.map((c) => RT[c].map((x, j) => x + (j && r.chance(0.3) ? 0.5 : 0)));
          const netOf = (x) => r100k(x.cartons * x.price * (1 - x.disc / 100)) || 100000;
          /* Planted: one line exactly on the top carton limit. */
          rows[3].cartons = th[2];
          rows.forEach((x) => { x.net = netOf(x); });
          const S = H.table([{ h: 'Invoice', k: 'invoice' }, { h: 'SKU', k: 'sku' }, { h: 'Cartons', k: 'cartons', num: true }, { h: 'Net (Rp)', k: 'net', num: true }], rows);
          const PM = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Product', k: 'name' }, { h: 'Category', k: 'category' }, { h: 'Carton price (Rp)', k: 'carton_price', num: true }], H.PRODUCTS);
          const G = grid('Category', th, cats.map((c, i) => ({ h: c, vals: vals[i] })), 0);
          const pm = sheetOf('master_produk.xlsx', 'Produk', PM), gk = sheetOf('tarif_komisi.xlsx', 'Kategori', G);
          /* The hidden test: a line moves to exactly the middle limit, and one category's rates rise in the rate file. */
          const k = Math.max(0, rows.findIndex((x, i) => i !== 3 && x.cartons < th[1]));
          const moved = Object.assign({}, rows[k], { cartons: th[1] });
          const alt = H.edit(S.cells, { [S.cell('Cartons', k)]: th[1], [S.cell('Net (Rp)', k)]: netOf(moved) });
          const ci = cats.indexOf(rows[0].cat);
          gk.alt = H.edit(G.cells, { ['D' + (5 + ci)]: (x) => x + 0.5, ['B' + (5 + ci)]: (x) => x + 0.25 });
          const p = gk.x.prefix, last = G.last;
          const GR = `${p}$B$5:$D$${last}`, CK = `${p}$A$5:$A$${last}`, TH = `${p}$B$4:$D$4`, ALL = `${p}$A$5:$D$${last}`, WIDE = `${p}$B$4:$D$${last}`;
          const MS = pm.x.A('SKU'), MC = pm.x.A('Category'), cat = `XLOOKUP(B5,${MS},${MC})`;
          const f = S.next + S.first, l = S.next + S.last;
          return multi({
            title: 'Line commission from a category grid, category from the master', file: 'komisi_kategori.xlsx', sheetName: 'Penjualan', sheet: S, base: S.cells, alt, fill: true, ansLabel: 'Commission (Rp)', open: [pm, gk],
            statement: `<p>The trade team pays commission on each order line. Three workbooks are open: the order lines (<b>komisi_kategori.xlsx</b>), the product master that gives each SKU its <b>Category</b> (<b>master_produk.xlsx</b>, sheet Produk) and the commission grid (<b>tarif_komisi.xlsx</b>, sheet Kategori). The grid has a row per category and carton limits across row 4; a line takes the column of the highest limit its <b>Cartons</b> reach (at least that many cartons).</p><p>In <b>${f}</b>, return the line's commission: its <b>Net (Rp)</b> times the rate (%) where its category's row meets its carton column, divided by 100 and rounded to whole rupiah with ROUND. Fill it down to <b>${l}</b>.</p>`,
            solution: `=ROUND(D5*INDEX(${GR},MATCH(${cat},${CK},0),MATCH(C5,${TH},1))/100,0)`,
            alts: [`=ROUND(D5*XLOOKUP(C5,${TH},XLOOKUP(${cat},${CK},${GR}),,-1)/100,0)`,
              `=ROUND(D5*VLOOKUP(VLOOKUP(B5,${pm.x.T('SKU', 'Category')},3,FALSE),${ALL},MATCH(C5,${TH},1)+1,FALSE)/100,0)`,
              `=ROUND(D5*HLOOKUP(C5,${WIDE},MATCH(INDEX(${MC},MATCH(B5,${MS},0)),${CK},0)+1,TRUE)/100,0)`],
            hints: ['The order lines have a SKU, not a category, so the first step is a lookup in master_produk.xlsx. The category and the cartons then give a row and a column in the grid.',
              `${cat} is the category. MATCH(that, ${CK}, 0) is the row and MATCH(C5, ${TH}, 1) the column: with 1, MATCH returns the last limit that is not above C5. INDEX takes both.`,
              `Shape: =ROUND(D5 * INDEX(grid, MATCH(category, categories, 0), MATCH(C5, limits, 1)) / 100, 0), with the category an XLOOKUP into the product master.`],
            explain: 'The XLOOKUP into the master turns the SKU into a category, the exact MATCH finds that category\'s row in the grid and the approximate MATCH finds the carton tier, so INDEX returns the one rate that applies. A line exactly on a limit takes the higher column, because MATCH with 1 accepts an equal value.',
            chips: ['B5', 'C5', 'D5*', MS, MC, GR, CK, TH, ',0)', ',1)'], fnKey: 'INDEX('
          });
        },
        () => {
          /* The quarter bonus per rep: the target, the visits and the tier rate each from its own file: four workbooks. */
          const reps = H.reps(r, 10);
          const tiers = [[0, 0], [0.9, r.pick([0.5, 0.75])], [1, r.pick([1, 1.25])], [1.1, r.pick([1.5, 2])]];
          const ach = (x) => x.actual / x.target;
          const [a, b, c2, e, g] = r.sample(nums(10), 5);
          /* Planted: rep a has exactly 30 visits (paid); rep b reaches target with too few visits; rep c2 is exactly on target;
             rep e has 28 visits until the CRM is corrected on the hidden test; rep g's target is revised. */
          reps[a].visits = 30;
          if (ach(reps[a]) < 0.9) reps[a].actual = r100k(reps[a].target * r.pick([0.95, 1.04, 1.12]));
          reps[b].visits = r.pick([24, 26, 29]);
          if (ach(reps[b]) < 1) reps[b].actual = r100k(reps[b].target * r.pick([1.04, 1.12]));
          reps[c2].visits = Math.max(reps[c2].visits, 34);
          reps[c2].actual = reps[c2].target;
          reps[e].visits = 28;
          if (ach(reps[e]) < 0.9) reps[e].actual = r100k(reps[e].target * 1.04);
          reps[g].visits = Math.max(reps[g].visits, 32);
          if (ach(reps[g]) < 1) reps[g].actual = r100k(reps[g].target * 1.04);
          const S = H.table([{ h: 'Rep', k: 'name' }, { h: 'Area', k: 'area' }, { h: 'Actual (Rp)', k: 'actual', num: true }], reps);
          const tg = r.shuffle(reps), kj = r.shuffle(reps);
          const TG = H.table([{ h: 'Rep', k: 'name' }, { h: 'Q4 target (Rp)', k: 'target', num: true }], tg);
          const BN = H.table([{ h: 'From achievement', v: (t) => t[0], num: true }, { h: 'Bonus rate (%)', v: (t) => t[1], num: true }], tiers);
          const KJ = H.table([{ h: 'Rep', k: 'name' }, { h: 'Visits', k: 'visits', num: true }], kj);
          const tgs = sheetOf('target_rep_2026.xlsx', 'Target', TG), bn = sheetOf('skema_bonus.xlsx', 'Bonus', BN), kjs = sheetOf('crm_kunjungan.xlsx', 'Kunjungan', KJ);
          /* The hidden test: the CRM adds three late-logged visits, one target is revised up, and the top tier pays more. */
          kjs.alt = H.edit(KJ.cells, { [KJ.cell('Visits', kj.indexOf(reps[e]))]: 31 });
          tgs.alt = H.edit(TG.cells, { [TG.cell('Q4 target (Rp)', tg.indexOf(reps[g]))]: (x) => x + 20000000 });
          bn.alt = H.edit(BN.cells, { [BN.cell('Bonus rate (%)', 3)]: (x) => x + 0.25 });
          const KR = kjs.x.A('Rep'), KV = kjs.x.A('Visits'), TR = tgs.x.A('Rep'), TT = tgs.x.A('Q4 target (Rp)'), Fr = bn.x.A('From achievement'), Br = bn.x.A('Bonus rate (%)');
          const vis = `XLOOKUP(A5,${KR},${KV})`, tgt = `XLOOKUP(A5,${TR},${TT})`;
          const f = S.next + S.first, l = S.next + S.last;
          return multi({
            title: 'Quarter bonus per rep from targets, tiers and visit counts', file: 'bonus_sales.xlsx', sheetName: 'Kinerja', sheet: S, base: S.cells, alt: S.cells, fill: true, ansLabel: 'Bonus (Rp)', open: [tgs, bn, kjs],
            statement: `<p>Pak Anton is preparing the quarter bonus. Four workbooks are open: each rep's <b>Actual</b> sales (<b>bonus_sales.xlsx</b>), the Q4 targets (<b>target_rep_2026.xlsx</b>, sheet Target), the bonus scheme (<b>skema_bonus.xlsx</b>, sheet Bonus) and the customer visits logged in the CRM (<b>crm_kunjungan.xlsx</b>, sheet Kunjungan).</p><p>Achievement is Actual divided by the rep's <b>Q4 target</b> (1 means 100% of target). The bonus rate is the one for the highest <b>From achievement</b> the rep reaches (at least that much), and the bonus is Actual × that rate / 100, rounded to whole rupiah with ROUND. One gate: a rep with fewer than 30 <b>Visits</b> gets 0, whatever the achievement. In <b>${f}</b>, return the bonus and fill it down to <b>${l}</b>.</p>`,
            solution: `=IF(${vis}<30,0,ROUND(C5*XLOOKUP(C5/${tgt},${Fr},${Br},,-1)/100,0))`,
            alts: [`=IF(VLOOKUP(A5,${kjs.x.T('Rep', 'Visits')},2,FALSE)<30,0,ROUND(C5*VLOOKUP(C5/VLOOKUP(A5,${tgs.x.T('Rep', 'Q4 target (Rp)')},2,FALSE),${bn.x.T('From achievement', 'Bonus rate (%)')},2,TRUE)/100,0))`,
              `=LET(t,${tgt},k,${vis},IF(k<30,0,ROUND(C5*XLOOKUP(C5/t,${Fr},${Br},,-1)/100,0)))`,
              `=IF(${vis}>=30,ROUND(C5*INDEX(${Br},MATCH(C5/INDEX(${TT},MATCH(A5,${TR},0)),${Fr},1))/100,0),0)`],
            hints: ['Three files, three lookups: the visits from crm_kunjungan.xlsx and the target from target_rep_2026.xlsx are exact lookups of the rep in A5; the rate from skema_bonus.xlsx is an approximate one, on the achievement.',
              `C5/${tgt} is the achievement. Look it up in ${Fr} with match_mode -1 to get the rate, and put the visit test in an IF around everything.`,
              'Shape: =IF(visits<30, 0, ROUND(C5 * XLOOKUP(C5/target, From achievement, Bonus rate, , -1) / 100, 0))'],
            explain: 'The two exact lookups bring the rep\'s target and visits from their own files, and the approximate XLOOKUP turns the achievement into the tier rate. IF checks the visit gate first, so a rep who sold well but visited too few customers gets 0. Exactly 30 visits passes the gate, and an achievement of exactly 1 takes the 100% tier.',
            chips: ['A5', 'C5', KR, KV, TR, TT, Fr, Br, ',,-1)', '<30'], fnKey: 'IF('
          });
        }
      ][v]();
    }
  });

  /* ======================= Collection: scores and queues from terms, rules, rosters and the loan book ======================= */

  P.add({
    key: 'collection-priority', difficulty: 'Hard', level: 10, topics: ['XLOOKUP', 'COUNTIF', 'INT', 'IF', 'SORTBY', 'TAKE', 'collection', 'Several workbooks'], count: 3, multi: true,
    make(r, v) {
      return [
        () => {
          /* A score from four workbooks: the due date needs the terms from the distributor master, the points come from the
             scoring table, plus the amount (INT), plus the watchlist (COUNTIF). */
          const inv = H.invoices(r, 14).map((x) => ({ ref: x.ref, d: x.distributor, date: ser(x.date), terms: x.terms, amount: x.amount }));
          inv.forEach((x) => { x.due = x.date + x.terms; });
          inv.sort((p, q) => p.due - q.due || (p.ref < q.ref ? -1 : 1));
          const asof = inv[inv.length - 1].due - r.int(2, 5);
          const pts = r.pick([[[0, 0], [1, 1], [31, 2], [61, 4], [91, 6]], [[0, 0], [1, 1], [31, 3], [61, 5], [91, 8]]]);
          const names = [...new Set(inv.map((x) => x.d))];
          const watch = r.sample(names, 3);
          const out = names.filter((n) => !watch.includes(n));
          /* Planted: an invoice exactly 30 days overdue (still the 1 to 30 tier). */
          const k30 = inv[r.int(2, 5)];
          k30.date = asof - 30 - k30.terms;
          inv.sort((p, q) => p.date - q.date || (p.ref < q.ref ? -1 : 1));
          const S = H.table([{ h: 'Invoice', k: 'ref' }, { h: 'Distributor', k: 'd' }, { h: 'Invoice date', k: 'date', date: true }, { h: 'Amount (Rp)', k: 'amount', num: true }], inv);
          const base = Object.assign({}, S.cells, { A1: 'As of', B1: asof });
          const master = r.shuffle(DIST()).map((d) => ({ d: d.name, city: d.city, terms: d.terms_days }));
          const MD = H.table([{ h: 'Distributor', k: 'd' }, { h: 'City', k: 'city' }, { h: 'Terms (days)', k: 'terms', num: true }], master);
          const PN = H.table([{ h: 'From days', v: (t) => t[0], num: true }, { h: 'Points', v: (t) => t[1], num: true }], pts);
          const REASON = ['Two late payments', 'Bank DPD above 30', 'Disputed delivery', 'Cheque returned'];
          const WL = H.table([{ h: 'Distributor', k: 'd' }, { h: 'Reason', k: 'why' }], watch.map((d) => ({ d, why: r.pick(REASON) })));
          const md = sheetOf('master_distributor.xlsx', 'Distributor', MD), pn = sheetOf('skor_penagihan.xlsx', 'Poin', PN), wl = sheetOf('daftar_pantau.xlsx', 'Pantau', WL);
          /* The hidden test: twelve days later, one distributor moves to new terms, and one name on the watchlist is replaced. */
          const alt = H.edit(base, { B1: asof + 12 });
          const moved = r.pick(names.filter((n) => n !== k30.d));
          md.alt = H.edit(MD.cells, { [MD.cell('Terms (days)', at(master, moved))]: (x) => (x === 45 ? 30 : 45) });
          wl.alt = H.edit(WL.cells, { [WL.cell('Distributor', 1)]: r.pick(out) });
          const tm = `XLOOKUP(B5,${md.x.A('Distributor')},${md.x.A('Terms (days)')})`, vt = `VLOOKUP(B5,${md.x.T('Distributor', 'Terms (days)')},3,FALSE)`;
          const Fd = pn.x.A('From days'), Pt = pn.x.A('Points'), TA = pn.x.T('From days', 'Points'), W = wl.x.A('Distributor');
          const f = S.next + S.first, l = S.next + S.last;
          return multi({
            title: 'Collection score from terms, points and watchlist files', file: 'antrian_penagihan.xlsx', sheetName: 'Tagihan', sheet: S, base, alt, fill: true, ansLabel: 'Score', dateCells: ['B1'], open: [md, pn, wl],
            statement: `<p>Bu Lestari's team calls the highest scores first. Four workbooks are open: the open invoices (<b>antrian_penagihan.xlsx</b>, with the as-of date in <b>B1</b>), the distributor master with each distributor's payment <b>Terms (days)</b> (<b>master_distributor.xlsx</b>, sheet Distributor), the scoring table (<b>skor_penagihan.xlsx</b>, sheet Poin) and the watchlist (<b>daftar_pantau.xlsx</b>, sheet Pantau).</p><p>An invoice is due its <b>Invoice date</b> plus the distributor's terms, and its days overdue are B1 minus that due date (an invoice not yet due counts as 0 days). Score = the <b>Points</b> for its days overdue, where a row of the scoring table applies from its <b>From days</b> on; <b>plus</b> 1 point for each full ${rp(5000000)} of <b>Amount</b>; <b>plus</b> 3 points when the distributor is on the watchlist. In <b>${f}</b>, return the score and fill it down to <b>${l}</b>.</p>`,
            solution: `=XLOOKUP(MAX(0,$B$1-C5-${tm}),${Fd},${Pt},,-1)+INT(D5/5000000)+IF(COUNTIF(${W},B5)>0,3,0)`,
            alts: [`=VLOOKUP(MAX(0,$B$1-(C5+${vt})),${TA},2,TRUE)+INT(D5/5000000)+3*COUNTIF(${W},B5)`,
              `=LOOKUP(MAX(0,$B$1-C5-${tm}),${Fd},${Pt})+ROUNDDOWN(D5/5000000,0)+IF(ISNUMBER(MATCH(B5,${W},0)),3,0)`,
              `=LET(late,MAX(0,$B$1-(C5+${tm})),XLOOKUP(late,${Fd},${Pt},,-1)+INT(D5/5000000)+IF(COUNTIF(${W},B5)>0,3,0))`,
              `=INDEX(${Pt},MATCH(MAX(0,$B$1-C5-INDEX(${md.x.A('Terms (days)')},MATCH(B5,${md.x.A('Distributor')},0))),${Fd},1))+INT(D5/5000000)+IF(COUNTIF(${W},B5)>0,3,0)`],
            hints: ['Start with the due date: C5 plus the terms of the distributor in B5, looked up in master_distributor.xlsx. Days overdue is $B$1 minus that, kept at 0 or more with MAX(0, ...).',
              `Part 1: XLOOKUP(days overdue, ${Fd}, ${Pt}, , -1). Part 2: INT(D5/5000000) counts full Rp 5,000,000 blocks. Part 3: COUNTIF(${W}, B5)>0 says whether the name is on the watchlist.`,
              `Shape: =XLOOKUP(MAX(0,$B$1-C5-XLOOKUP(B5, ..., Terms)), From days, Points, , -1) + INT(D5/5000000) + IF(COUNTIF(watchlist, B5)>0, 3, 0)`],
            explain: 'The terms lookup turns the invoice date into a due date, the approximate XLOOKUP turns the days overdue into points by tier, INT drops the part of the amount below the next full block, and COUNTIF checks the name against the watchlist in its own file. $B$1 stays pinned so every row is scored on the same date. An invoice exactly 30 days overdue is still in the 1 to 30 tier.',
            chips: ['$B$1-C5-', 'MAX(0,', md.x.A('Distributor'), md.x.A('Terms (days)'), Fd, Pt, ',,-1)', 'INT(D5/5000000)', W, 'COUNTIF('], fnKey: 'XLOOKUP('
          });
        },
        () => {
          /* A queue from a chain across files: the outstanding from the loan book, DPD to Kol (approximate) and Kol to
             provision rate (exact) from two sheets of the rules workbook, then expected loss in an OR. */
          const rows = H.loans(r, 14).map((x) => ({ f: x.facility, d: x.distributor, out: x.outstanding, limit: x.limit, dpd: x.dpd }));
          const KOL = [[0, 1], [1, 2], [91, 3], [121, 4], [181, 5]];
          const prov = [1, 5, r.pick([15, 20]), 50, 100];
          /* Planted: DPD exactly 90 on a small line, DPD 91, a Kol 2 line between Rp 62.5 and 100 million, a clean line, a
             Kol 2 line of exactly Rp 100 million (expected loss exactly Rp 5,000,000, exact in floating point too), and a small
             Kol 3 line past 90 days whose expected loss stays under Rp 5,000,000, so only the DPD test calls it. */
          const [p0, p1, p2, p3, p4, p5] = r.sample(nums(14), 6);
          rows[p0].dpd = 90; rows[p0].out = r.int(20, 50) * 1000000;
          rows[p1].dpd = 91;
          rows[p2].dpd = r.pick([15, 30, 45]); rows[p2].out = r.int(65, 95) * 1000000;
          rows[p3].dpd = 0;
          rows[p4].dpd = r.pick([7, 21, 60]); rows[p4].out = 100000000;
          rows[p5].dpd = r.pick([95, 105, 115]); rows[p5].out = r.int(10, 20) * 1000000;
          rows.forEach((x) => { if (x.out > x.limit) x.limit = Math.ceil(x.out / 50000000) * 50000000; });
          const S = H.table([{ h: 'Facility', k: 'f' }, { h: 'Distributor', k: 'd' }, { h: 'DPD', k: 'dpd', num: true }], rows);
          const book = r.shuffle(rows);
          const PJ = H.table([{ h: 'Facility', k: 'f' }, { h: 'Distributor', k: 'd' }, { h: 'Limit (Rp)', k: 'limit', num: true }, { h: 'Outstanding (Rp)', k: 'out', num: true }], book);
          const KL = H.table([{ h: 'From DPD', v: (t) => t[0], num: true }, { h: 'Kol', v: (t) => t[1], num: true }], KOL);
          /* The provision sheet lists the worst grade first, so it does not line up row for row with the DPD table: skipping
             the Kol step and reading the rate beside the matched DPD row gives the wrong rate. */
          const provRows = [5, 4, 3, 2, 1].map((k) => ({ k, p: prov[k - 1] }));
          const PV = H.table([{ h: 'Kol', k: 'k', num: true }, { h: 'Provision (%)', k: 'p', num: true }], provRows);
          const pj = sheetOf('buku_pinjaman.xlsx', 'Pinjaman', PJ), kl = sheetOf('aturan_kol.xlsx', 'Kol', KL), pv = sheetOf('aturan_kol.xlsx', 'Provisi', PV);
          /* The hidden test: Kol 2 provisions rise to 8% in the rules workbook, the Rp 100 million line is paid down to
             Rp 50 million in the loan book (its expected loss falls below Rp 5,000,000 at either rate), and the clean line
             falls 15 days behind. */
          pv.alt = H.edit(PV.cells, { [PV.cell('Provision (%)', at(provRows, 2, 'k'))]: 8 });
          pj.alt = H.edit(PJ.cells, { [PJ.cell('Outstanding (Rp)', book.indexOf(rows[p4]))]: 50000000 });
          const alt = H.edit(S.cells, { [S.cell('DPD', p3)]: 15 });
          const Fd = kl.x.A('From DPD'), Ko = kl.x.A('Kol'), K2 = pv.x.A('Kol'), Pv = pv.x.A('Provision (%)');
          const out = `XLOOKUP(A5,${pj.x.A('Facility')},${pj.x.A('Outstanding (Rp)')})`, kol = `XLOOKUP(C5,${Fd},${Ko},,-1)`;
          const el = `${out}*XLOOKUP(${kol},${K2},${Pv})/100`;
          const vel = `VLOOKUP(A5,${pj.x.T('Facility', 'Outstanding (Rp)')},4,FALSE)*VLOOKUP(VLOOKUP(C5,${kl.x.T('From DPD', 'Kol')},2,TRUE),${pv.x.T('Kol', 'Provision (%)')},2,FALSE)/100`;
          const f = S.next + S.first, l = S.next + S.last;
          return multi({
            title: 'Collection queue from DPD, Kol rules and the loan book', file: 'daftar_tagih.xlsx', sheetName: 'Fasilitas', sheet: S, base: S.cells, alt, fill: true, ansLabel: 'Priority', open: [pj, kl, pv],
            statement: `<p>Citra sorts the bank's facilities into three call lists. Three workbooks are open: her list with each facility's days past due (<b>daftar_tagih.xlsx</b>), the loan book with the <b>Outstanding</b> by facility (<b>buku_pinjaman.xlsx</b>, sheet Pinjaman) and the bank's rules (<b>aturan_kol.xlsx</b>), which has two sheets: <b>Kol</b> turns a DPD into a Kol grade (a row applies from its From DPD on) and <b>Provisi</b> gives each Kol its provision rate.</p><p>Expected loss = Outstanding × provision rate / 100. Return <b>Call today</b> when the expected loss is at least ${rp(5000000)} <b>or</b> DPD is more than 90; otherwise <b>This week</b> when DPD is more than 0; otherwise <b>Monitor</b>. Start in <b>${f}</b> and fill down to <b>${l}</b>.</p>`,
            solution: `=IF(OR(${el}>=5000000,C5>90),"Call today",IF(C5>0,"This week","Monitor"))`,
            alts: [`=LET(k,${kol},e,${out}*XLOOKUP(k,${K2},${Pv})/100,IF(OR(e>=5000000,C5>90),"Call today",IF(C5>0,"This week","Monitor")))`,
              `=IF(OR(${vel}>=5000000,C5>90),"Call today",IF(C5>0,"This week","Monitor"))`,
              `=IFS(OR(${el}>=5000000,C5>90),"Call today",C5>0,"This week",TRUE,"Monitor")`,
              `=IF(OR(INDEX(${pj.x.A('Outstanding (Rp)')},MATCH(A5,${pj.x.A('Facility')},0))*INDEX(${Pv},MATCH(INDEX(${Ko},MATCH(C5,${Fd},1)),${K2},0))/100>=5000000,C5>90),"Call today",IF(C5>0,"This week","Monitor"))`],
            hints: ['Work in a chain across files: the facility in A5 gives the outstanding (loan book), the DPD in C5 gives the Kol (sheet Kol, approximate), and the Kol gives the provision rate (sheet Provisi, exact).',
              `${kol} is the Kol. Use it as the value to find in ${K2}, and multiply the rate by the outstanding / 100: ${el}.`,
              'Shape: =IF(OR(expected loss>=5000000, C5>90), "Call today", IF(C5>0, "This week", "Monitor"))'],
            explain: 'The first XLOOKUP brings the balance from the loan book. The inner XLOOKUP on the Kol sheet grades the facility by its arrears, and the one around it turns the grade into a provision rate from the Provisi sheet of the same rules workbook, so the expected loss follows the bank\'s own tables. OR puts a large expected loss or a long arrear on today\'s list, and the second IF splits the rest between this week and monitoring.',
            chips: ['A5', 'C5', pj.x.A('Facility'), pj.x.A('Outstanding (Rp)'), Fd, Ko, K2, Pv, ',,-1)', '>=5000000', '"Call today"', '"This week"', '"Monitor"'], fnKey: 'IF('
          });
        },
        () => {
          /* The three invoices one collector should chase first: the region from the roster, the minimum from the policy,
             then a computed key inside SORTBY, cut with TAKE. */
          const COLL = ['Nanda', 'Rizal', 'Putri', 'Bayu', 'Ika', 'Joko', 'Sari'];
          const roster = r.shuffle(H.REGIONS).map((g, i) => ({ who: COLL[i], g }));
          /* On the hidden test the roster swaps the collector's region with a colleague's, so a region typed into the formula
             fails: the collector's three calls then come from the colleague's region. */
          const me = r.pick(roster), g = me.g;
          const mate = r.pick(roster.filter((x) => x !== me)), g2 = mate.g;
          const MIN = {};
          H.REGIONS.forEach((x) => { MIN[x] = r.pick([5, 7.5, 10]) * 1000000; });
          MIN[g] = 10000000; MIN[g2] = 10000000;
          const asof = ser('2026-11-' + day2(r.int(10, 20))), later = asof + 10;
          const inRegion = (x) => DIST().filter((d) => d.region === x).map((d) => d.name);
          const mine = inRegion(g), next = inRegion(g2);
          const others = DIST().filter((d) => d.region !== g && d.region !== g2);
          const list = [];
          const add = (d, amt, late, paid) => list.push({ d, region: regionOf(d), amt, due: asof - late, paid: paid ? asof - r.int(1, 5) : '' });
          /* In the collector's region: four open invoices at or above the minimum, one more at 12 to 14.5 million that ranks
             first, one below the minimum, one paid, and one not yet due. */
          for (let i = 0; i < 4; i++) add(r.pick(mine), r.int(16, 30) * 1000000, r.int(10, 45));
          const keyOf = (x, d) => (d - x.due) * x.amt;
          const ranked = () => list.slice(0, 4).map((x) => keyOf(x, asof)).sort((p, q) => q - p);
          const e1 = r.int(121, 145) * 100000;
          add(r.pick(mine), e1, Math.ceil(ranked()[0] / e1) + r.int(2, 6));
          const below = r.int(60, 90) * 100000;
          add(r.pick(mine), below, Math.ceil(ranked()[1] / below) + r.int(3, 10));
          const top = keyOf(list[4], asof), big = r.int(40, 60) * 1000000;
          add(r.pick(mine), big, Math.ceil(top / big) + r.int(3, 10), true);
          add(r.pick(mine), r.int(20, 30) * 1000000, -r.int(5, 20));
          /* In the colleague's region: three open invoices at Rp 16 to 30 million, one at Rp 12 to 14 million that would rank
             first ten days later but drops out when the policy raises that region's minimum to Rp 15,000,000, and one below
             any minimum. On the visible sheet they are other-region invoices with large keys. */
          const from2 = list.length;
          for (let i = 0; i < 3; i++) add(r.pick(next), r.int(16, 30) * 1000000, r.int(10, 45));
          const top2 = Math.max(...list.slice(from2).map((x) => keyOf(x, later)));
          const e2 = r.int(121, 140) * 100000;
          add(r.pick(next), e2, Math.ceil(top2 / e2) - 10 + r.int(2, 6));
          const below2 = r.int(60, 90) * 100000;
          add(r.pick(next), below2, Math.ceil(top2 / below2) - 10 + r.int(3, 10));
          /* An invoice in a third region that outranks every open one above, on both dates. */
          const far = r.int(30, 50) * 1000000;
          add(r.pick(others).name, far, Math.ceil(Math.max(...list.filter((x) => x.paid === '').map((x) => keyOf(x, later))) / far) + r.int(3, 10));
          for (let i = 0; i < 4; i++) add(r.pick(others).name, r.int(5, 60) * 1000000, r.int(-10, 120), r.chance(0.3));
          /* No ties among the open, eligible keys on either date, so the order never depends on the row order. */
          const seen = new Set();
          list.forEach((x) => {
            if (x.paid !== '') return;
            while (seen.has(keyOf(x, asof)) || seen.has(keyOf(x, asof + 10))) x.amt += 100000;
            seen.add(keyOf(x, asof)); seen.add(keyOf(x, asof + 10));
          });
          const rows = r.shuffle(list).map((x, i) => Object.assign(x, { inv: 'INV-' + (3101 + i * 4 + r.int(0, 3)) }));
          const S = H.table([{ h: 'Invoice', k: 'inv' }, { h: 'Distributor', k: 'd' }, { h: 'Region', k: 'region' }, { h: 'Due date', k: 'due', date: true }, { h: 'Amount (Rp)', k: 'amt', num: true }, { h: 'Paid on', v: (x) => (x.paid === '' ? undefined : x.paid) }], rows);
          const base = Object.assign({}, S.cells, { A1: 'As of', B1: asof, A2: 'Collector', B2: me.who });
          const paidRefs = rows.map((x, i) => (x.paid === '' ? null : S.cell('Paid on', i))).filter(Boolean);
          const RO = H.table([{ h: 'Collector', k: 'who' }, { h: 'Region', k: 'g' }], roster);
          const amb = r.shuffle(H.REGIONS).map((x) => ({ g: x, m: MIN[x] }));
          const AM = H.table([{ h: 'Region', k: 'g' }, { h: 'Minimum amount (Rp)', k: 'm', num: true }], amb);
          const ro = sheetOf('roster_penagihan.xlsx', 'Roster', RO), am = sheetOf('kebijakan_penagihan.xlsx', 'Ambang', AM);
          /* The hidden test: ten days later, the roster swaps the collector's region with the colleague's, and the policy
             raises the colleague's region's minimum to Rp 15,000,000. */
          const alt = H.edit(base, { B1: later });
          ro.alt = H.edit(RO.cells, { [RO.cell('Region', roster.indexOf(me))]: g2, [RO.cell('Region', roster.indexOf(mate))]: g });
          am.alt = H.edit(AM.cells, { [AM.cell('Minimum amount (Rp)', at(amb, g2, 'g'))]: 15000000 });
          const Iv = S.R('Invoice'), Rg = S.R('Region'), Du = S.R('Due date'), Am = S.R('Amount (Rp)'), Pd = S.R('Paid on');
          const RC = ro.x.A('Collector'), RR = ro.x.A('Region'), AR = am.x.A('Region'), AMm = am.x.A('Minimum amount (Rp)');
          const gx = `XLOOKUP(B2,${RC},${RR})`, mx = (gv) => `XLOOKUP(${gv},${AR},${AMm})`;
          const key = (gv, mv) => `(${Rg}=${gv})*(${Pd}="")*(${Am}>=${mv})*(B1-${Du})*${Am}`;
          return multi({
            title: 'Top three invoices for one collector, region from the roster', file: 'piutang_terbuka.xlsx', sheetName: 'Piutang', sheet: S, base, alt, spill: { cols: 1, rows: 3 }, ansLabel: 'Call first', dateCells: ['B1'].concat(paidRefs), open: [ro, am],
            statement: `<p>Bu Lestari plans each collector's first three calls. Three workbooks are open: Rasa Kenari's invoices (<b>piutang_terbuka.xlsx</b>, with the as-of date in <b>B1</b> and the collector in <b>B2</b>), the collection roster that gives each collector a region (<b>roster_penagihan.xlsx</b>, sheet Roster) and the collection policy with the smallest amount worth a call in each region (<b>kebijakan_penagihan.xlsx</b>, sheet Ambang).</p><p>An invoice counts for the collector when it is in the collector's <b>Region</b>, still open (<b>Paid on</b> is empty) and its <b>Amount</b> is at least the region's <b>Minimum amount</b>. Its urgency is Amount × days overdue as of B1 (B1 minus <b>Due date</b>). In <b>${S.next + S.first}</b>, list the <b>Invoice</b> numbers of the collector's three most urgent invoices, most urgent first, in one formula that spills down.</p>`,
            solution: `=LET(g,${gx},m,${mx('g')},TAKE(SORTBY(${Iv},${key('g', 'm')},-1),3))`,
            alts: [`=TAKE(SORTBY(${Iv},${key(gx, mx(gx))},-1),3)`,
              `=LET(g,VLOOKUP(B2,${ro.x.T('Collector', 'Region')},2,FALSE),m,VLOOKUP(g,${am.x.T('Region', 'Minimum amount (Rp)')},2,FALSE),TAKE(SORTBY(${Iv},${key('g', 'm')},-1),3))`,
              `=LET(g,${gx},m,${mx('g')},k,(${Rg}=g)*(${Pd}="")*(${Am}>=m),TAKE(SORTBY(FILTER(${Iv},k=1),FILTER((B1-${Du})*${Am},k=1),-1),3))`,
              `=LET(g,INDEX(${RR},MATCH(B2,${RC},0)),m,INDEX(${AMm},MATCH(g,${AR},0)),TAKE(SORTBY(${Iv},${key('g', 'm')},-1),3))`],
            hints: ['Two lookups first, one feeding the other: the collector in B2 gives a region (roster file), and that region gives the minimum amount (policy file). LET can name them g and m.',
              `Then a key for every row at once: ${key('g', 'm')}. Each test in brackets is 1 or 0, so an invoice that fails any test scores 0.`,
              `Shape: =LET(g, XLOOKUP(B2, ...), m, XLOOKUP(g, ...), TAKE(SORTBY(${Iv}, key, -1), 3))`],
            explain: 'The roster turns the collector into a region and the policy turns the region into a minimum, so the same formula serves any collector typed in B2. The key multiplies three tests by the days overdue and the amount, so invoices in other regions, paid ones and small ones score 0. SORTBY orders the invoice numbers by that key without a helper column, and TAKE keeps the top three.',
            chips: ['B2', RC, RR, AR, AMm, Iv, Rg, Pd, Am, Du, '(' + Pd + '="")', ',-1)', ',3)'], fnKey: 'TAKE('
          });
        }
      ][v]();
    }
  });

  /* ======================= Reorder plans: forecasts, suppliers, schedules and caps in other files ======================= */

  P.add({
    key: 'reorder-plans', difficulty: 'Hard', level: 8, topics: ['IF', 'ROUNDUP', 'MAX', 'MIN', 'INT', 'SUMPRODUCT', 'XLOOKUP', 'stock', 'Several workbooks'], count: 3, multi: true,
    make(r, v) {
      return [
        () => {
          /* Reorder when stock is at or below the reorder point; top up for lead time plus cover, in full pallets. The daily
             sales, the lead time and the pallet size each come from their own file: four workbooks. */
          const rows = H.stock(r, 12).map((x) => ({ sku: x.sku, wh: x.warehouse }));
          const daily = {}, lead = {}, sup = {};
          rows.forEach((x) => { daily[x.sku] = r.int(3, 25); lead[x.sku] = r.pick([3, 5, 7, 10, 14]); sup[x.sku] = 'SP-' + day2(r.int(1, 9)); });
          const PAL = { Noodles: r.pick([30, 36]), Coffee: r.pick([20, 24]), Sambal: r.pick([24, 30]), Soap: r.pick([20, 24]) };
          const low = new Set(r.sample(nums(12), 5));
          rows.forEach((x, i) => { const rop = daily[x.sku] * lead[x.sku]; x.oh = low.has(i) ? Math.max(0, rop - r.int(1, 12) * 5) : rop + r.int(2, 30) * 10; });
          /* Planted: one SKU exactly on its reorder point (it orders), one that crosses it when the forecast rises, and one
             whose supplier needs longer. */
          const lowI = [...low], highI = nums(12).filter((i) => !low.has(i));
          /* The slow supplier serves the busiest SKU that orders (other than the one on its reorder point), so a few more
             days of lead time add at least one pallet on their own. */
          const eq = lowI[0], dip = highI[0], slow = lowI.slice(1).reduce((a, i) => (daily[rows[i].sku] > daily[rows[a].sku] ? i : a));
          rows[eq].oh = daily[rows[eq].sku] * lead[rows[eq].sku];
          const cover = r.pick([14, 21]);
          const S = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Warehouse', k: 'wh' }, { h: 'On hand', k: 'oh', num: true }], rows);
          const base = Object.assign({}, S.cells, { A1: 'Cover days', B1: cover });
          const fc = r.shuffle(rows).map((x) => ({ sku: x.sku, d: daily[x.sku] }));
          const PR = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Daily sales (cartons)', k: 'd', num: true }], fc);
          const ps = r.shuffle(rows).map((x) => ({ sku: x.sku, sup: sup[x.sku], lt: lead[x.sku] }));
          const PS = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Supplier', k: 'sup' }, { h: 'Lead time (days)', k: 'lt', num: true }], ps);
          const PM = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Product', k: 'name' }, { h: 'Category', k: 'category' }, { h: 'Carton price (Rp)', k: 'carton_price', num: true }, { h: 'Cartons per pallet', v: (p) => PAL[p.category], num: true }], H.PRODUCTS);
          const pr = sheetOf('prakiraan_penjualan.xlsx', 'Prakiraan', PR), sp = sheetOf('data_pemasok.xlsx', 'Pemasok', PS), pm = sheetOf('master_produk.xlsx', 'Produk', PM);
          /* The hidden test: a longer cover, a higher forecast that takes one SKU below its reorder point, and a supplier
             that needs a few more days (enough for at least one more pallet). */
          const alt = H.edit(base, { B1: cover + 7 });
          const dx = rows[dip], sx = rows[slow];
          const extra = Math.max(3, Math.ceil(PAL[H.PRODUCTS.find((p) => p.sku === sx.sku).category] / daily[sx.sku]));
          pr.alt = H.edit(PR.cells, { [PR.cell('Daily sales (cartons)', at(fc, dx.sku, 'sku'))]: Math.ceil(dx.oh / lead[dx.sku]) + r.int(1, 3) });
          sp.alt = H.edit(PS.cells, { [PS.cell('Lead time (days)', at(ps, sx.sku, 'sku'))]: (x) => x + extra });
          const xd = `XLOOKUP(A5,${pr.x.A('SKU')},${pr.x.A('Daily sales (cartons)')})`, xt = `XLOOKUP(A5,${sp.x.A('SKU')},${sp.x.A('Lead time (days)')})`, xp = `XLOOKUP(A5,${pm.x.A('SKU')},${pm.x.A('Cartons per pallet')})`;
          const vd = `VLOOKUP(A5,${pr.x.T('SKU', 'Daily sales (cartons)')},2,FALSE)`, vt = `VLOOKUP(A5,${sp.x.T('SKU', 'Lead time (days)')},3,FALSE)`, vp = `VLOOKUP(A5,${pm.x.T('SKU', 'Cartons per pallet')},5,FALSE)`;
          const f = S.next + S.first, l = S.next + S.last;
          return multi({
            title: 'Reorder in full pallets from forecast, supplier and pallet files', file: 'rencana_pesan.xlsx', sheetName: 'Stok', sheet: S, base, alt, fill: true, ansLabel: 'Order (cartons)', open: [pr, sp, pm],
            statement: `<p>Rasa Kenari's stock planner reviews every SKU each morning. Four workbooks are open: the stock on hand (<b>rencana_pesan.xlsx</b>, with the cover days in <b>B1</b>), the sales forecast (<b>prakiraan_penjualan.xlsx</b>, sheet Prakiraan), the supplier list with each SKU's <b>Lead time</b> (<b>data_pemasok.xlsx</b>, sheet Pemasok) and the product master with the <b>Cartons per pallet</b> (<b>master_produk.xlsx</b>, sheet Produk).</p><p>A SKU needs an order when its <b>On hand</b> is at or below its reorder point, <b>Daily sales</b> × Lead time. It then orders Daily sales × (Lead time + B1) - On hand, rounded <b>up</b> to full pallets. A SKU above its reorder point orders 0. In <b>${f}</b>, return the order in cartons and fill it down to <b>${l}</b>.</p>`,
            solution: `=LET(d,${xd},t,${xt},p,${xp},IF(C5<=d*t,ROUNDUP((d*(t+$B$1)-C5)/p,0)*p,0))`,
            alts: [`=IF(C5<=${xd}*${xt},ROUNDUP((${xd}*(${xt}+$B$1)-C5)/${xp},0)*${xp},0)`,
              `=LET(d,${vd},t,${vt},p,${vp},IF(C5<=d*t,CEILING(d*(t+$B$1)-C5,p),0))`,
              `=LET(d,${xd},t,${xt},p,INDEX(${pm.x.A('Cartons per pallet')},MATCH(A5,${pm.x.A('SKU')},0)),IF(C5>d*t,0,CEILING(d*(t+$B$1)-C5,p)))`],
            hints: ['Three lookups of the SKU in A5, each in its own file: the daily sales, the lead time and the pallet size. LET(d, ..., t, ..., p, ..., ...) names them so each is written once.',
              'The test is C5<=d*t. The need is d*(t+$B$1)-C5 cartons; to round up to full pallets, divide by p, ROUNDUP(..., 0), and multiply back by p. CEILING(need, p) does it in one step.',
              'Shape: =LET(d, daily, t, lead time, p, pallet, IF(C5<=d*t, ROUNDUP((d*(t+$B$1)-C5)/p, 0)*p, 0))'],
            explain: 'Each file answers one question about the SKU, so the three lookups run first and LET keeps the formula readable. IF orders only when stock has reached the reorder point, so a SKU sitting exactly on it still orders. The need covers the days until the truck arrives plus the cover days, and rounding up to whole pallets means the warehouse never receives a part pallet. $B$1 stays pinned as you fill.',
            chips: ['A5', 'C5', pr.x.A('SKU'), pr.x.A('Daily sales (cartons)'), sp.x.A('SKU'), sp.x.A('Lead time (days)'), pm.x.A('SKU'), pm.x.A('Cartons per pallet'), '$B$1', 'ROUNDUP(', 'LET('], fnKey: 'IF('
          });
        },
        () => {
          /* The latest order date: full days of cover (INT, from the forecast file) less the warehouse's lead time (from the
             delivery schedule), never before the as-of date. */
          const rows = H.stock(r, 12).map((x) => ({ sku: x.sku, wh: x.warehouse, oh: r.int(4, 60) * 10 }));
          const daily = {};
          rows.forEach((x) => { daily[x.sku] = r.int(4, 25); });
          const whs = r.shuffle(H.WAREHOUSES);
          const lead = {};
          whs.forEach((w) => { lead[w] = r.pick([3, 5, 7, 10, 14]); });
          /* Planted: two SKUs already late (cover below lead time), so the answer is the as-of date itself. */
          r.sample(nums(12), 2).forEach((i) => { const x = rows[i]; x.oh = Math.max(10, Math.floor((lead[x.wh] - 2) * daily[x.sku] / 10) * 10); });
          const asof = ser('2026-11-02') + r.int(0, 10);
          const S = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Warehouse', k: 'wh' }, { h: 'On hand', k: 'oh', num: true }], rows);
          const base = Object.assign({}, S.cells, { A1: 'As of', B1: asof });
          const fc = r.shuffle(rows).map((x) => ({ sku: x.sku, d: daily[x.sku] }));
          const PR = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Daily sales (cartons)', k: 'd', num: true }], fc);
          const TRUCK = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri'];
          const jd = whs.map((w, i) => ({ w, day: TRUCK[i], lt: lead[w] }));
          const JD = H.table([{ h: 'Warehouse', k: 'w' }, { h: 'Truck day', k: 'day' }, { h: 'Lead time (days)', k: 'lt', num: true }], jd);
          const pr = sheetOf('prakiraan_penjualan.xlsx', 'Prakiraan', PR), jk = sheetOf('jadwal_kirim.xlsx', 'Jadwal', JD);
          /* The hidden test: a week later, one warehouse's supplier needs three more days, and one SKU's forecast is revised. */
          const alt = H.edit(base, { B1: asof + 7 });
          jk.alt = H.edit(JD.cells, { [JD.cell('Lead time (days)', at(jd, rows[0].wh, 'w'))]: (x) => x + 3 });
          pr.alt = H.edit(PR.cells, { [PR.cell('Daily sales (cartons)', at(fc, rows[5].sku, 'sku'))]: (x) => x + 5 });
          const xd = `XLOOKUP(A5,${pr.x.A('SKU')},${pr.x.A('Daily sales (cartons)')})`, xl = `XLOOKUP(B5,${jk.x.A('Warehouse')},${jk.x.A('Lead time (days)')})`;
          const vd = `VLOOKUP(A5,${pr.x.T('SKU', 'Daily sales (cartons)')},2,FALSE)`, vl = `VLOOKUP(B5,${jk.x.T('Warehouse', 'Lead time (days)')},3,FALSE)`;
          const f = S.next + S.first, l = S.next + S.last;
          return multi({
            title: 'Latest order date from the forecast and the delivery schedule', file: 'jadwal_pesan.xlsx', sheetName: 'Stok', sheet: S, base, alt, fill: true, ansLabel: 'Order by', ansDate: true, dateCells: ['B1'], open: [pr, jk],
            statement: `<p>The planner wants an order-by date for every SKU. Three workbooks are open: the stock on hand (<b>jadwal_pesan.xlsx</b>, with the as-of date in <b>B1</b>), the sales forecast with each SKU's <b>Daily sales</b> (<b>prakiraan_penjualan.xlsx</b>, sheet Prakiraan) and the delivery schedule with each warehouse's <b>Lead time</b> (<b>jadwal_kirim.xlsx</b>, sheet Jadwal).</p><p>A SKU's stock lasts the full days of <b>On hand</b> divided by its Daily sales (drop any part day). In <b>${f}</b>, return the latest date to place the order: B1 plus the full days of stock, minus the lead time of the SKU's <b>Warehouse</b>. When that date is before B1, the order is already late, so return B1 itself. The answer is a date; fill it down to <b>${l}</b>.</p>`,
            solution: `=MAX($B$1,$B$1+INT(C5/${xd})-${xl})`,
            alts: [`=MAX($B$1,$B$1+INT(C5/${vd})-${vl})`,
              `=IF(INT(C5/${xd})<${xl},$B$1,$B$1+INT(C5/${xd})-${xl})`,
              `=MAX($B$1,$B$1+ROUNDDOWN(C5/INDEX(${pr.x.A('Daily sales (cartons)')},MATCH(A5,${pr.x.A('SKU')},0)),0)-${xl})`],
            hints: ['Two lookups on two different keys: the daily sales by the SKU in A5 in the forecast file, and the lead time by the warehouse in B5 in the schedule.',
              `INT(C5/${xd}) is the full days of stock. $B$1 plus that, minus ${xl}, is the order-by date; MAX with $B$1 stops it falling before the as-of date.`,
              'Shape: =MAX($B$1, $B$1 + INT(C5/XLOOKUP(A5, ..., Daily sales)) - XLOOKUP(B5, ..., Lead time))'],
            explain: 'Dates are day numbers, so adding the days of stock and taking away the lead time gives a date. INT keeps whole days only, because a part day of stock cannot wait for a delivery, and MAX turns any date in the past into B1, the day the order must go out. The forecast goes by SKU and the schedule by warehouse, so each lookup uses its own column of the row.',
            chips: ['$B$1', 'A5', 'B5', 'C5', pr.x.A('SKU'), pr.x.A('Daily sales (cartons)'), jk.x.A('Warehouse'), jk.x.A('Lead time (days)'), 'INT(', 'MAX('], fnKey: 'MAX('
          });
        },
        () => {
          /* Each warehouse's purchase budget: top every SKU at or below its reorder point up to its maximum, at cost
             (SUMPRODUCT over the stock file), capped by finance's weekly limit (lookup in the budget file). */
          const whs = H.WAREHOUSES.slice();
          const REGW = { 'Gudang Bekasi': 'Jabodetabek', 'Gudang Bandung': 'Jawa Barat', 'Gudang Semarang': 'Jawa Tengah', 'Gudang Surabaya': 'Jawa Timur', 'Gudang Medan': 'Sumatera' };
          const pattern = r.shuffle([0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4]);
          const st = H.stock(r, 20).map((x, i) => ({ sku: x.sku, wh: whs[pattern[i]], rop: x.reorder, cost: x.cost }));
          st.forEach((x) => { x.max = Math.round(x.rop * r.pick([2, 2.5, 3]) / 10) * 10; x.oh = r.chance(0.45) ? Math.max(0, x.rop - r.int(0, 8) * 10) : x.rop + r.int(1, 30) * 10; });
          const of = (w) => st.filter((x) => x.wh === w);
          /* Planted: in every warehouse one SKU exactly at its reorder point (it orders) and one above it. */
          whs.forEach((w) => { const g = of(w); g[0].oh = g[0].rop; g[1].oh = g[1].rop + r.int(2, 20) * 10; });
          const need = (w) => of(w).reduce((a, x) => a + (x.oh <= x.rop ? (x.max - x.oh) * x.cost : 0), 0);
          const bind = r.sample(whs, 2);
          const cap = {};
          whs.forEach((w) => { cap[w] = Math.max(5000000, Math.round(need(w) * (bind.includes(w) ? 0.7 : 1.4) / 5000000) * 5000000); if (cap[w] === need(w)) cap[w] += 5000000; });
          const free = whs.filter((w) => !bind.includes(w));
          const [wA, wB] = r.sample(free, 2);
          const S = H.table([{ h: 'Warehouse', k: 'w' }, { h: 'Region', k: 'g' }], whs.map((w) => ({ w, g: REGW[w] })));
          const list = st.map((x) => x);
          const ST = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Warehouse', k: 'wh' }, { h: 'On hand', k: 'oh', num: true }, { h: 'Reorder point', k: 'rop', num: true }, { h: 'Max stock', k: 'max', num: true }, { h: 'Cost per carton (Rp)', k: 'cost', num: true }], list);
          const plaf = r.shuffle(whs).map((w) => ({ w, cap: cap[w] }));
          const PF = H.table([{ h: 'Warehouse', k: 'w' }, { h: 'Weekly cap (Rp)', k: 'cap', num: true }], plaf);
          const sg = sheetOf('stok_gudang.xlsx', 'Stok', ST), pf = sheetOf('anggaran_gudang.xlsx', 'Plafon', PF);
          /* The hidden test: finance cuts one warehouse's cap below its need, and in another a SKU sells down to its reorder point. */
          pf.alt = H.edit(PF.cells, { [PF.cell('Weekly cap (Rp)', at(plaf, wA, 'w'))]: Math.max(5000000, Math.round(need(wA) * 0.6 / 5000000) * 5000000) });
          sg.alt = H.edit(ST.cells, { [ST.cell('On hand', list.indexOf(of(wB)[1]))]: of(wB)[1].rop });
          const Wr = sg.x.A('Warehouse'), OH = sg.x.A('On hand'), RP = sg.x.A('Reorder point'), MX = sg.x.A('Max stock'), CO = sg.x.A('Cost per carton (Rp)');
          const nd = `SUMPRODUCT((${Wr}=A5)*(${OH}<=${RP})*(${MX}-${OH})*${CO})`;
          const PW = pf.x.A('Warehouse'), PC = pf.x.A('Weekly cap (Rp)'), cp = `XLOOKUP(A5,${PW},${PC})`;
          const f = S.next + S.first, l = S.next + S.last;
          return multi({
            title: 'Weekly purchase budget per warehouse, capped by finance', file: 'anggaran_pembelian.xlsx', sheetName: 'Ringkasan', sheet: S, base: S.cells, alt: S.cells, fill: true, ansLabel: 'Budget (Rp)', open: [sg, pf],
            statement: `<p>Finance asks for this week's purchase budget per warehouse. Three workbooks are open: the summary (<b>anggaran_pembelian.xlsx</b>), the stock file with every SKU in every warehouse (<b>stok_gudang.xlsx</b>, sheet Stok) and finance's weekly caps (<b>anggaran_gudang.xlsx</b>, sheet Plafon).</p><p>In a warehouse, every SKU whose <b>On hand</b> is at or below its <b>Reorder point</b> is topped up to its <b>Max stock</b>, so it orders Max stock minus On hand cartons at <b>Cost per carton</b>; SKUs above their reorder point order nothing. In <b>${f}</b>, return the cost of those orders for the warehouse in A5, but never more than its <b>Weekly cap</b>. Fill it down to <b>${l}</b>.</p>`,
            solution: `=MIN(${cp},${nd})`,
            alts: [`=MIN(VLOOKUP(A5,${pf.x.T('Warehouse', 'Weekly cap (Rp)')},2,FALSE),SUMPRODUCT(--(${Wr}=A5),--(${OH}<=${RP}),${MX}-${OH},${CO}))`,
              `=LET(n,${nd},k,${cp},IF(n>k,k,n))`,
              `=MIN(INDEX(${PC},MATCH(A5,${PW},0)),${nd})`],
            hints: ['Two parts from two files: the need, a SUMPRODUCT over stok_gudang.xlsx for the warehouse in A5, and the cap, a lookup of A5 in anggaran_gudang.xlsx. The budget is the smaller of the two.',
              `For one row the cost is (Max stock - On hand) × Cost. (${Wr}=A5) and (${OH}<=${RP}) are 1 or 0, so SUMPRODUCT can multiply them in: ${nd}.`,
              'Shape: =MIN(XLOOKUP(A5, warehouse, cap), SUMPRODUCT((warehouse=A5)*(on hand<=reorder point)*(max-on hand)*cost))'],
            explain: 'Each test turns into 1 or 0, so rows from other warehouses and SKUs above their reorder point add nothing, and SUMPRODUCT adds the cost of the rest without a helper column. XLOOKUP brings the cap from finance\'s file, and MIN pays the need when it fits and the cap when it does not. A SKU exactly on its reorder point orders.',
            chips: ['A5', Wr, OH, RP, MX, CO, PW, PC, '(' + Wr + '=A5)', '<=', 'MIN('], fnKey: 'SUMPRODUCT('
          });
        }
      ][v]();
    }
  });
});
