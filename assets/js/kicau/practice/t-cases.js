/*
 * Practice, Real cases (20 Hard): requests from Rasa Kenari and Bank Kenari that need several functions in one
 * formula. Credit decisions with lookups and OR, early-warning flags from two tables on one sheet, weighted averages
 * with SUMPRODUCT, commission from rate tables, collection priority and reorder quantities.
 * Each template makes `count` problems; v (0 to count-1) picks the variant and r is its seeded random.
 * Levels: 8 is one SUMPRODUCT or one IF around arithmetic, 9 joins two tables or chains lookups inside a decision,
 * 10 builds a score or a queue from three tables, or ranks rows by a computed key.
 *
 * Side tables sit to the right of the main list with one empty column between, so a column answer fills down beside
 * the main list. What the practice engine does not do, so no problem here relies on it: a lookup that searches for a
 * whole range at once (XLOOKUP(B5:B18, ...)) inside SUMPRODUCT, SUMIFS with an array as criteria, INDEX with an array of
 * row numbers (INDEX(..., SEQUENCE(3))), and ISBLANK on a range. No row sits exactly on a boundary that floating-point
 * arithmetic could tip (such as 70% of a sales figure); exact boundaries are kept to whole numbers (days past due, Kol,
 * cartons) and to exact divisions.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) factory(require('./core.js'));
  else factory(root.KicauPractice);
})(typeof globalThis !== 'undefined' ? globalThis : this, function (P) {
  'use strict';
  const H = P.H;
  const L = H.LETTERS;
  const DIST = () => H.db().distributors;
  const nums = (n) => Array.from({ length: n }, (_, i) => i);
  const rp = (n) => 'Rp ' + Math.round(n).toLocaleString('en-US');
  const ser = (iso) => H.toSerial(iso);
  const r100k = (x) => Math.round(x / 100000) * 100000;

  /*
   * Several tables on one sheet, side by side (as in t-lookup.js). The first is the main one: a column answer fills
   * its rows. The grid shows every table down to the longest one, and tapping a column header inserts that table's range.
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

  /* A grid with numbers across its header row (row 4), starting at column startCol: one labelled line per row. */
  function grid(corner, heads, lines, startCol) {
    const sc = startCol, cells = {};
    cells[L[sc] + '4'] = corner;
    heads.forEach((h, j) => { cells[L[sc + 1 + j] + '4'] = h; });
    lines.forEach((ln, i) => { cells[L[sc] + (5 + i)] = ln.h; ln.vals.forEach((x, j) => { cells[L[sc + 1 + j] + (5 + i)] = x; }); });
    const cols = L.slice(sc, sc + heads.length + 1);
    return { cells, cols, num: cols.slice(1), dateCols: [], first: 5, last: 4 + lines.length, map: {}, next: L[sc + heads.length + 1] };
  }

  /* The made-up world, summed: a distributor's net sales in one month ('2026-05'), and its six-month average. */
  const netIn = (name, m) => H.db().sales.filter((s) => s.distributor === name && s.date.startsWith(m)).reduce((a, s) => a + s.net, 0);
  const avgMonthly = (name) => r100k(H.db().sales.filter((s) => s.distributor === name).reduce((a, s) => a + s.net, 0) / 6);
  /* Keep a ratio at least `gap` away from a limit, by raising the top figure one million at a time. */
  function clear(x, top, bottom, limit, gap) { while (Math.abs(x[top] / x[bottom] - limit) < gap) x[top] += 1000000; }

  /* ======================= Credit decisions: lookups inside OR and IF ======================= */

  P.add({
    key: 'credit-decision-rules', difficulty: 'Hard', level: 9, topics: ['IF', 'OR', 'XLOOKUP', 'MAXIFS', 'credit'], count: 4,
    make(r, v) {
      return [
        () => {
          /* Approve or Decline from two facts in the bank check: the Kol grade and the restructured flag. */
          const ds = r.shuffle(DIST());
          const info = {};
          ds.forEach((d) => { info[d.name] = { kol: r.pick([1, 1, 1, 1, 2, 2, 2, 3, 4, 5]), rs: r.chance(0.15) ? 'Y' : 'N' }; });
          const asked = ds.slice(0, 12);
          /* Planted: Kol exactly 3, restructured with a good Kol, and two rows that change on the hidden sheet. */
          info[asked[0].name] = { kol: 3, rs: 'N' };
          info[asked[1].name] = { kol: r.pick([1, 2]), rs: 'Y' };
          info[asked[2].name] = { kol: 2, rs: 'N' };
          info[asked[3].name] = { kol: 1, rs: 'Y' };
          info[asked[4].name] = { kol: 1, rs: 'N' };
          const rows = r.shuffle(asked).map((d, i) => ({ req: 'CR-' + (411 + i * 3), d: d.name, amt: r.int(8, 40) * 5000000, tenor: r.pick([12, 18, 24, 36]) }));
          const side = r.shuffle(ds).map((d) => ({ d: d.name, kol: info[d.name].kol, rs: info[d.name].rs }));
          const S = H.table([{ h: 'Request', k: 'req' }, { h: 'Distributor', k: 'd' }, { h: 'Requested (Rp)', k: 'amt', num: true }, { h: 'Tenor (months)', k: 'tenor', num: true }], rows);
          const M = H.table([{ h: 'Distributor', k: 'd' }, { h: 'Kol', k: 'kol', num: true }, { h: 'Restructured', k: 'rs' }], side, { startCol: S.cols.length + 2 });
          const B = board([S, M]);
          const at = (name) => side.findIndex((x) => x.d === name);
          /* The hidden sheet: one borrower slips to Kol 3, another's restructuring flag is corrected to N. */
          const alt = H.edit(B.cells, { [M.cell('Kol', at(asked[2].name))]: 3, [M.cell('Restructured', at(asked[3].name))]: 'N', [S.cell('Requested (Rp)', 5)]: (x) => x + 10000000 });
          const K = M.A('Distributor'), KO = M.A('Kol'), RS = M.A('Restructured'), TA = M.T('Distributor', 'Restructured');
          const f = S.next + S.first, l = S.next + S.last, side1 = M.col('Distributor') + M.first + ':' + M.col('Restructured') + M.last;
          return build(B, {
            title: 'Credit decision from Kol grade or restructuring', file: 'komite_kredit.xlsx', base: B.cells, alt, fill: true, ansLabel: 'Decision',
            statement: `<p>Dimas is preparing Friday's credit committee at Bank Kenari. Each row in A:D is a distributor asking for a bigger working-capital line. The bank check in <b>${side1}</b> gives every distributor's collectibility grade (<b>Kol</b>, 1 is best and 5 is worst) and whether its loan was ever <b>Restructured</b> (Y or N).</p><p>The rule: <b>Decline</b> when the distributor's Kol is 3 or worse (3, 4 or 5) <b>or</b> its loan was restructured; otherwise <b>Approve</b>. In <b>${f}</b>, return Approve or Decline for the request, and fill it down to <b>${l}</b>.</p>`,
            solution: `=IF(OR(XLOOKUP(B5,${K},${KO})>=3,XLOOKUP(B5,${K},${RS})="Y"),"Decline","Approve")`,
            alts: [`=IF(OR(VLOOKUP(B5,${TA},2,FALSE)>=3,VLOOKUP(B5,${TA},3,FALSE)="Y"),"Decline","Approve")`,
              `=IF(AND(XLOOKUP(B5,${K},${KO})<3,XLOOKUP(B5,${K},${RS})="N"),"Approve","Decline")`,
              `=IF(OR(INDEX(${KO},MATCH(B5,${K},0))>=3,INDEX(${RS},MATCH(B5,${K},0))="Y"),"Decline","Approve")`],
            hints: ['Two facts come from the bank check, the Kol and the Restructured flag. Each one is a lookup of the distributor in B5.',
              `OR(test1, test2) is TRUE when either test holds. Put both lookups inside it: XLOOKUP(B5, ${K}, ${KO})>=3 and XLOOKUP(B5, ${K}, ${RS})="Y". The $ signs keep the bank check in place as you fill.`,
              `Shape: =IF(OR(XLOOKUP(B5, ${K}, ${KO})>=3, XLOOKUP(B5, ${K}, ${RS})="Y"), "Decline", "Approve")`],
            explain: 'Each XLOOKUP finds the distributor in the bank check and returns one fact about it. OR turns either warning into a Decline, and IF writes the word. Kol 3 itself must decline, so the test is >=3, not >3.',
            chips: ['B5', K, KO, RS, '>=3', '="Y"', '"Decline"', '"Approve"'], fnKey: 'IF('
          });
        },
        () => {
          /* Three outcomes: Decline on bad DPD, Refer on a large request or any arrears, Approve otherwise. */
          const ds = r.shuffle(DIST());
          const side = ds.map((d) => ({ d: d.name, s: avgMonthly(d.name), dpd: r.pick([0, 0, 0, 0, 0, 5, 12, 30, 45, 75, 95]) }));
          const by = (name) => side.find((x) => x.d === name);
          const asked = r.sample(ds, 12).map((d) => ({ d: d.name, f: r.pick([0.8, 1.1, 1.3, 1.6, 1.8, 2.4, 2.8]) }));
          /* Planted [DPD, request as a multiple of sales]: exactly 2 times, DPD exactly 60 and 61, and two that change later. */
          [[0, 2], [60, 1.2], [61, 1.1], [0, 2.6], [0, 1.2], [0, 1.5]].forEach(([dp, f], k) => { by(asked[k].d).dpd = dp; asked[k].f = f; });
          asked.forEach((a) => { a.amt = a.f === 2 ? 2 * by(a.d).s : Math.round(by(a.d).s * a.f / 1000000) * 1000000; });
          const rows = r.shuffle(asked).map((a, i) => ({ req: 'CR-' + (430 + i * 2), d: a.d, amt: a.amt }));
          const S = H.table([{ h: 'Request', k: 'req' }, { h: 'Distributor', k: 'd' }, { h: 'Requested (Rp)', k: 'amt', num: true }], rows);
          const M = H.table([{ h: 'Distributor', k: 'd' }, { h: 'Avg monthly sales (Rp)', k: 's', num: true }, { h: 'DPD', k: 'dpd', num: true }], side, { startCol: S.cols.length + 2 });
          const B = board([S, M]);
          const sIdx = (name) => side.findIndex((x) => x.d === name), mIdx = (name) => rows.findIndex((x) => x.d === name);
          /* The hidden sheet: one clean borrower falls 75 days behind, and one request grows past twice its sales. */
          const alt = H.edit(B.cells, { [M.cell('DPD', sIdx(asked[4].d))]: 75, [S.cell('Requested (Rp)', mIdx(asked[5].d))]: Math.round(by(asked[5].d).s * 2.5 / 1000000) * 1000000 });
          const K = M.A('Distributor'), SA = M.A('Avg monthly sales (Rp)'), DP = M.A('DPD'), TA = M.T('Distributor', 'DPD');
          const f = S.next + S.first, l = S.next + S.last, side1 = M.col('Distributor') + M.first + ':' + M.col('DPD') + M.last;
          const lk = (x) => `XLOOKUP(B5,${K},${x})`;
          return build(B, {
            title: 'Approve, refer or decline on DPD and sales cover', file: 'komite_kredit.xlsx', base: B.cells, alt, fill: true, ansLabel: 'Decision',
            statement: `<p>Bank Kenari screens new loan requests before the committee sees them. The distributor table in <b>${side1}</b> gives each distributor's average monthly net sales at Rasa Kenari and its days past due (<b>DPD</b>) at the bank today.</p><p>The rule, in this order: <b>Decline</b> when the distributor's DPD is more than 60. Otherwise <b>Refer</b> when the request is more than 2 times its average monthly sales, <b>or</b> its DPD is more than 0. Otherwise <b>Approve</b>. In <b>${f}</b>, return the decision and fill it down to <b>${l}</b>.</p>`,
            solution: `=IF(${lk(DP)}>60,"Decline",IF(OR(C5>2*${lk(SA)},${lk(DP)}>0),"Refer","Approve"))`,
            alts: [`=IFS(${lk(DP)}>60,"Decline",OR(C5>2*${lk(SA)},${lk(DP)}>0),"Refer",TRUE,"Approve")`,
              `=LET(s,${lk(SA)},d,${lk(DP)},IF(d>60,"Decline",IF(OR(C5>2*s,d>0),"Refer","Approve")))`,
              `=IF(VLOOKUP(B5,${TA},3,FALSE)>60,"Decline",IF(OR(C5>2*VLOOKUP(B5,${TA},2,FALSE),VLOOKUP(B5,${TA},3,FALSE)>0),"Refer","Approve"))`],
            hints: ['Check the rules in order: the Decline test first, then the Refer test, then Approve for everything left.',
              `Two lookups do the work: ${lk(SA)} is the sales and ${lk(DP)} the DPD. The Refer test is OR(C5>2*sales, DPD>0).`,
              `Shape: =IF(DPD>60, "Decline", IF(OR(C5>2*sales, DPD>0), "Refer", "Approve")), with each name replaced by its XLOOKUP.`],
            explain: 'The outer IF settles the worst case first, so a borrower 61 days behind is declined even though it would also be referred. The inner IF uses OR to refer a large request or any arrears, and approves the rest. A request of exactly twice the sales is not more than twice, so it can still be approved.',
            chips: ['B5', 'C5', K, SA, DP, '>60', '>0', '"Refer"', '"Decline"', '"Approve"'], fnKey: 'IF('
          });
        },
        () => {
          /* One borrower across all its facilities: MAXIFS and COUNTIFS for the red flags, SUMIFS for the exposure cap. */
          const ds = r.sample(DIST(), 5).map((d) => d.name);
          const TY = ['Working Capital', 'Term Loan', 'Trade Finance'];
          const rows = [];
          ds.forEach((d, k) => {
            const m = k < 2 ? 3 : r.pick([2, 3]);
            r.sample(TY, m).forEach((t) => rows.push({ d, t, out: (k === 0 ? r.int(24, 32) : r.int(8, 30)) * 2500000, dpd: r.pick([0, 0, 0, 7, 15]), rs: 'N' }));
          });
          const of = (d) => rows.filter((x) => x.d === d);
          /* Borrower A: worst DPD exactly 30, never restructured, close to the cap, so the request tips it into Refer. */
          const A = of(ds[0]);
          A[r.int(0, A.length - 1)].dpd = 30;
          let tA = A.reduce((a, x) => a + x.out, 0);
          if (tA > 230000000) { A[0].out -= tA - 230000000; tA = 230000000; }
          const reqA = Math.ceil((250000000 - tA) / 5000000) * 5000000 + r.pick([5, 10, 20]) * 1000000;
          /* Borrower B: low DPD, but one facility was restructured, so it is declined. C has one facility 45 days behind. */
          r.pick(of(ds[1])).rs = 'Y';
          r.pick(of(ds[2])).dpd = 45;
          if (r.chance(0.5)) r.pick(of(ds[3])).rs = 'Y';
          /* B's request would also take it past the limit, so a formula that tests the limit before the red flags says Refer. */
          const tB = of(ds[1]).reduce((a, x) => a + x.out, 0);
          const reqB = Math.max(0, Math.ceil((250000000 - tB) / 5000000)) * 5000000 + r.int(2, 6) * 5000000;
          const list = r.shuffle(rows).map((x, i) => Object.assign(x, { f: 'F' + (311 + i) }));
          const S = H.table([{ h: 'Facility', k: 'f' }, { h: 'Borrower', k: 'd' }, { h: 'Type', k: 't' }, { h: 'Outstanding (Rp)', k: 'out', num: true }, { h: 'DPD', k: 'dpd', num: true }, { h: 'Restructured', k: 'rs' }], list);
          const base = Object.assign({}, S.cells, { A1: 'Borrower', B1: ds[0], A2: 'New request (Rp)', B2: reqA, A3: 'Decision' });
          /* The hidden sheet: the restructured borrower, whose request also breaks the limit; the red flag decides first. */
          const alt = H.edit(base, { B1: ds[1], B2: reqB });
          const BR = S.R('Borrower'), OUT = S.R('Outstanding (Rp)'), DP = S.R('DPD'), RS = S.R('Restructured');
          return H.problem({
            title: 'Credit decision for one borrower across its facilities', file: 'eksposur_debitur.xlsx', sheet: S, base, alt, answerCell: 'B3',
            statement: `<p>Dimas needs a quick answer for the credit desk on a new request. A borrower can hold several facilities, so the whole list counts. For the borrower in <b>B1</b>, asking for the amount in <b>B2</b>:</p><p><b>Decline</b> when any of its facilities is more than 30 days past due (<b>DPD</b>) <b>or</b> any of them was restructured (Y). Otherwise <b>Approve</b> when its total <b>Outstanding</b> plus the new request is at most ${rp(250000000)} (the single-borrower limit), and <b>Refer</b> when it is above. Return the decision in <b>B3</b>.</p>`,
            solution: `=IF(OR(MAXIFS(${DP},${BR},B1)>30,COUNTIFS(${BR},B1,${RS},"Y")>0),"Decline",IF(SUMIFS(${OUT},${BR},B1)+B2<=250000000,"Approve","Refer"))`,
            alts: [`=IFS(OR(MAXIFS(${DP},${BR},B1)>30,COUNTIFS(${BR},B1,${RS},"Y")>0),"Decline",SUMIFS(${OUT},${BR},B1)+B2<=250000000,"Approve",TRUE,"Refer")`,
              `=IF(SUMPRODUCT((${BR}=B1)*((${DP}>30)+(${RS}="Y")))>0,"Decline",IF(SUMPRODUCT((${BR}=B1)*${OUT})+B2<=250000000,"Approve","Refer"))`,
              `=IF(COUNTIFS(${BR},B1,${DP},">30")+COUNTIFS(${BR},B1,${RS},"Y")>0,"Decline",IF(SUMIFS(${OUT},${BR},B1)+B2>250000000,"Refer","Approve"))`],
            hints: ['Three questions about the borrower\'s rows: its worst DPD, whether any row says Y, and its total outstanding. MAXIFS, COUNTIFS and SUMIFS answer them.',
              `MAXIFS(${DP}, ${BR}, B1)>30 and COUNTIFS(${BR}, B1, ${RS}, "Y")>0 are the two red flags; join them with OR. The exposure test is SUMIFS(${OUT}, ${BR}, B1)+B2<=250000000.`,
              'Shape: =IF(OR(flag1, flag2), "Decline", IF(total + B2 <= 250000000, "Approve", "Refer"))'],
            explain: 'MAXIFS finds the borrower\'s worst arrears and COUNTIFS counts its restructured facilities, so OR declines on either flag across all its rows. Only a clean borrower reaches the inner IF, where SUMIFS adds its outstanding to the request and compares it with the limit. A worst DPD of exactly 30 is not more than 30.',
            chips: [BR, DP, RS, OUT, 'B1', 'B2', '"Y"', '250000000', '"Decline"', '"Approve"', '"Refer"'], fnKey: 'IF('
          });
        },
        () => {
          /* Collateral cover after a haircut looked up by collateral type, joined with a DPD limit by OR. */
          const HC = { Land: r.pick([25, 30]), Building: r.pick([35, 40]), Vehicle: r.pick([45, 50]), Inventory: r.pick([55, 60]), Receivables: r.pick([40, 45]) };
          const types = Object.keys(HC);
          const cover = (x) => x.val * (100 - HC[x.c]) / 100;
          const ds = r.sample(DIST(), 12);
          const rows = ds.map((d) => ({ d: d.name, amt: r.int(6, 30) * 5000000, c: r.pick(types), ratio: r.pick([0.7, 0.85, 1.15, 1.3, 1.6]), dpd: r.pick([0, 0, 0, 0, 8, 21, 45, 92]) }));
          /* Planted [ratio, DPD]: covered at DPD exactly 30 and 31, short of cover at DPD 0, covered at DPD 0. */
          [[1.3, 30], [1.3, 31], [0.85, 0], [1.3, 0]].forEach(([q, dp], k) => { rows[k].ratio = q; rows[k].dpd = dp; });
          const valueFor = (x, q) => {
            let val = Math.round(x.amt / ((100 - HC[x.c]) / 100) * q / 5000000) * 5000000;
            const probe = { val, c: x.c };
            while (q < 1 && cover(probe) > x.amt - 1000000) { val -= 5000000; probe.val = val; }
            while (q > 1 && cover(probe) < x.amt + 1000000) { val += 5000000; probe.val = val; }
            return val;
          };
          rows.forEach((x) => { x.val = valueFor(x, x.ratio); });
          const short = rows[2], clean = rows[3];
          const list = r.shuffle(rows).map((x, i) => Object.assign(x, { req: 'CR-' + (520 + i * 2) }));
          const S = H.table([{ h: 'Request', k: 'req' }, { h: 'Distributor', k: 'd' }, { h: 'Requested (Rp)', k: 'amt', num: true }, { h: 'Collateral', k: 'c' }, { h: 'Collateral value (Rp)', k: 'val', num: true }, { h: 'DPD', k: 'dpd', num: true }], list);
          const M = H.table([{ h: 'Collateral', k: 'c' }, { h: 'Haircut (%)', k: 'h', num: true }], types.map((c) => ({ c, h: HC[c] })), { startCol: S.cols.length + 2 });
          const B = board([S, M]);
          /* The hidden sheet: the short request brings a fresh valuation that covers it, and a clean borrower falls 45 days behind. */
          const alt = H.edit(B.cells, { [S.cell('Collateral value (Rp)', list.indexOf(short))]: valueFor(short, 1.3), [S.cell('DPD', list.indexOf(clean))]: 45 });
          const K = M.A('Collateral'), HA = M.A('Haircut (%)'), TA = M.T('Collateral', 'Haircut (%)');
          const f = S.next + S.first, l = S.next + S.last, side1 = M.col('Collateral') + M.first + ':' + M.col('Haircut (%)') + M.last;
          return build(B, {
            title: 'Credit decision on collateral cover after haircut', file: 'agunan_kredit.xlsx', base: B.cells, alt, fill: true, ansLabel: 'Decision',
            statement: `<p>The bank does not lend against the full value of collateral. The table in <b>${side1}</b> gives the <b>Haircut (%)</b> for each type of collateral: a haircut of 30 means only 70% of the value counts.</p><p>The rule: <b>Decline</b> when the collateral value after its haircut (<b>Collateral value</b> times (1 minus haircut/100)) is less than the amount requested, <b>or</b> the distributor is more than 30 days past due (<b>DPD</b>). Otherwise <b>Approve</b>. In <b>${f}</b>, return the decision and fill it down to <b>${l}</b>.</p>`,
            solution: `=IF(OR(E5*(1-XLOOKUP(D5,${K},${HA})/100)<C5,F5>30),"Decline","Approve")`,
            alts: [`=IF(OR(E5*(1-VLOOKUP(D5,${TA},2,FALSE)/100)<C5,F5>30),"Decline","Approve")`,
              `=IF(AND(E5*(100-XLOOKUP(D5,${K},${HA}))/100>=C5,F5<=30),"Approve","Decline")`,
              `=IF(OR(E5*(1-INDEX(${HA},MATCH(D5,${K},0))/100)<C5,F5>30),"Decline","Approve")`],
            hints: ['First the cover: look up the haircut for the collateral type in D5, then take that share off the collateral value.',
              `E5*(1-XLOOKUP(D5, ${K}, ${HA})/100) is the value that counts. Compare it with C5, and test F5>30, inside OR.`,
              `Shape: =IF(OR(E5*(1-XLOOKUP(D5, ${K}, ${HA})/100)<C5, F5>30), "Decline", "Approve")`],
            explain: 'XLOOKUP finds the haircut for the row\'s collateral type, and 1 minus haircut/100 is the share of the value the bank counts. OR declines when that cover falls short of the request or the arrears pass 30 days; everything else is approved.',
            chips: ['C5', 'D5', 'E5', 'F5>30', K, HA, '(1-', '/100)', '"Decline"', '"Approve"'], fnKey: 'IF('
          });
        }
      ][v]();
    }
  });

  /* ======================= Early-warning flags: two tables on one sheet ======================= */

  P.add({
    key: 'early-warning-flags', difficulty: 'Hard', level: 9, topics: ['IF', 'AND', 'OR', 'XLOOKUP', 'SUMIFS', 'early warning'], count: 4,
    make(r, v) {
      return [
        () => {
          /* Rasa Kenari's sales (main) and Bank Kenari's DPD (side): Watch only when both signals fire. */
          const rows = DIST().map((d) => ({ d: d.name, tr: d.trouble, may: r100k(netIn(d.name, '2026-05') * (0.9 + r() * 0.2)), oct: r100k(netIn(d.name, '2026-10') * (0.9 + r() * 0.2)) }));
          rows.forEach((x) => {
            if (x.tr) x.oct = r100k(x.may * r.pick([0.3, 0.45, 0.55]));
            clear(x, 'oct', 'may', 0.7, 0.02);
            x.dpd = x.tr ? r.pick([61, 75, 94, 95]) : r.pick([0, 0, 0, 0, 0, 4, 9, 15]);
          });
          const drop = (x) => x.oct <= 0.7 * x.may;
          const normal = rows.filter((x) => !x.tr);
          /* Planted: a distributor with a sales drop and DPD exactly 30 (OK), and one behind at the bank with steady sales (OK). */
          let flipA = normal.find(drop);
          if (!flipA) { flipA = normal[0]; flipA.oct = r100k(flipA.may * 0.55); }
          flipA.dpd = 30;
          const dpdOnly = r.pick(normal.filter((x) => x !== flipA && !drop(x)));
          if (dpdOnly) dpdOnly.dpd = r.pick([45, 60]);
          const flipB = rows.find((x) => x.tr);
          const side = r.shuffle(rows).map((x) => ({ d: x.d, dpd: x.dpd }));
          const S = H.table([{ h: 'Distributor', k: 'd' }, { h: 'Net May (Rp)', k: 'may', num: true }, { h: 'Net Oct (Rp)', k: 'oct', num: true }], rows);
          const M = H.table([{ h: 'Distributor', k: 'd' }, { h: 'DPD', k: 'dpd', num: true }], side, { startCol: S.cols.length + 2 });
          const B = board([S, M]);
          /* The hidden sheet: the distributor at DPD 30 slips to 45, and one troubled distributor's October recovers. */
          const alt = H.edit(B.cells, { [M.cell('DPD', side.findIndex((x) => x.d === flipA.d))]: 45, [S.cell('Net Oct (Rp)', rows.indexOf(flipB))]: r100k(flipB.may * 0.9) });
          const K = M.A('Distributor'), DP = M.A('DPD'), TA = M.T('Distributor', 'DPD');
          const f = S.next + S.first, l = S.next + S.last, side1 = M.col('Distributor') + M.first + ':' + M.col('DPD') + M.last;
          return build(B, {
            title: 'Early-warning flag: sales drop and bank DPD together', file: 'peringatan_dini.xlsx', base: B.cells, alt, fill: true, ansLabel: 'Flag',
            statement: `<p>Rasa Kenari and Bank Kenari share an early-warning list. Columns A:C are Rasa Kenari's net sales to each distributor in May and October; <b>${side1}</b> is the bank's days past due (<b>DPD</b>) for the same distributors, in a different order.</p><p>Flag <b>Watch</b> when <b>both</b> signals fire: October net sales are at most 70% of May (a fall of 30% or more) <b>and</b> the bank's DPD is more than 30. Otherwise <b>OK</b>. In <b>${f}</b>, return the flag and fill it down to <b>${l}</b>.</p>`,
            solution: `=IF(AND(C5<=0.7*B5,XLOOKUP(A5,${K},${DP})>30),"Watch","OK")`,
            alts: [`=IF(AND(C5/B5<=0.7,VLOOKUP(A5,${TA},2,FALSE)>30),"Watch","OK")`,
              `=IF(OR(C5>0.7*B5,XLOOKUP(A5,${K},${DP})<=30),"OK","Watch")`,
              `=IF(AND(C5<=0.7*B5,INDEX(${DP},MATCH(A5,${K},0))>30),"Watch","OK")`],
            hints: ['The sales test uses this row only: C5<=0.7*B5. The DPD sits in the other table, so it needs a lookup on the distributor in A5.',
              `XLOOKUP(A5, ${K}, ${DP}) brings the bank's DPD across. Both tests must hold, so join them with AND.`,
              `Shape: =IF(AND(C5<=0.7*B5, XLOOKUP(A5, ${K}, ${DP})>30), "Watch", "OK")`],
            explain: 'The two tables list the distributors in different orders, so XLOOKUP matches each name to its DPD instead of reading the row beside it. AND fires only when the sales fall and the arrears agree, which keeps one-off dips off the list.',
            chips: ['A5', 'C5<=0.7*B5', K, DP, '>30', '"Watch"', '"OK"'], fnKey: 'IF('
          });
        },
        () => {
          /* Bank facilities (main) and Rasa Kenari's sales (side): Red for both signs, Amber for one, Green for none. */
          const trouble = (name) => DIST().find((d) => d.name === name).trouble;
          const snap = (id, m) => H.db().snapshots.find((s) => s.facility_id === id && s.month === m).dpd;
          const rows = H.db().facilities.map((f) => ({ f: f.facility_id, d: f.distributor, may: trouble(f.distributor) ? snap(f.facility_id, '2026-05') : r.pick([0, 0, 0, 5, 10]), oct: trouble(f.distributor) ? snap(f.facility_id, '2026-10') : r.pick([0, 0, 0, 0, 5, 10, 15]) }));
          const side = r.shuffle(DIST()).map((d) => {
            const x = { d: d.name, may: r100k(netIn(d.name, '2026-05') * (0.9 + r() * 0.2)), oct: r100k(netIn(d.name, '2026-10') * (0.9 + r() * 0.2)) };
            if (d.trouble) x.oct = r100k(x.may * r.pick([0.35, 0.5, 0.6]));
            clear(x, 'oct', 'may', 0.75, 0.02);
            return x;
          });
          const sd = (name) => side.find((x) => x.d === name);
          const s2 = (name) => sd(name).oct < 0.75 * sd(name).may;
          /* Planted: a Green facility that turns Amber later, and an Amber one (DPD rose, sales steady) that turns Red. */
          const calm = rows.filter((x) => !trouble(x.d) && !s2(x.d));
          const green = calm[0] || rows.find((x) => !trouble(x.d));
          green.oct = green.may;
          if (s2(green.d)) { sd(green.d).oct = sd(green.d).may; }
          const amber = calm.find((x) => x !== green && x.d !== green.d) || rows.find((x) => !trouble(x.d) && x.d !== green.d);
          amber.oct = amber.may + r.pick([5, 10]);
          if (s2(amber.d)) { sd(amber.d).oct = sd(amber.d).may; }
          const S = H.table([{ h: 'Facility', k: 'f' }, { h: 'Distributor', k: 'd' }, { h: 'DPD May', k: 'may', num: true }, { h: 'DPD Oct', k: 'oct', num: true }], rows);
          const M = H.table([{ h: 'Distributor', k: 'd' }, { h: 'Net May (Rp)', k: 'may', num: true }, { h: 'Net Oct (Rp)', k: 'oct', num: true }], side, { startCol: S.cols.length + 2 });
          const B = board([S, M]);
          /* The hidden sheet: the Green facility's DPD rises, and the Amber distributor's October sales fall to 60% of May. */
          const alt = H.edit(B.cells, { [S.cell('DPD Oct', rows.indexOf(green))]: green.may + 12, [M.cell('Net Oct (Rp)', side.indexOf(sd(amber.d)))]: r100k(sd(amber.d).may * 0.6) });
          const K = M.A('Distributor'), MY = M.A('Net May (Rp)'), OC = M.A('Net Oct (Rp)'), TA = M.T('Distributor', 'Net Oct (Rp)');
          const f = S.next + S.first, l = S.next + S.last, side1 = M.col('Distributor') + M.first + ':' + M.col('Net Oct (Rp)') + M.last;
          const fell = `XLOOKUP(B5,${K},${OC})<0.75*XLOOKUP(B5,${K},${MY})`;
          return build(B, {
            title: 'Red, amber or green from DPD trend and sales drop', file: 'peringatan_dini.xlsx', base: B.cells, alt, fill: true, ansLabel: 'Signal',
            statement: `<p>Dimas grades every Bank Kenari facility for the monthly risk meeting, using two warning signs. Sign 1: the facility's days past due rose (<b>DPD Oct</b> is more than <b>DPD May</b>). Sign 2: the distributor's net sales at Rasa Kenari fell by more than 25% (<b>Net Oct</b> is less than 75% of <b>Net May</b>, from the sales table in <b>${side1}</b>).</p><p>Return <b>Red</b> when both signs show, <b>Amber</b> when one shows, and <b>Green</b> when neither does. Start in <b>${f}</b> and fill down to <b>${l}</b>.</p>`,
            solution: `=IF(AND(D5>C5,${fell}),"Red",IF(OR(D5>C5,${fell}),"Amber","Green"))`,
            alts: [`=LET(s,${fell},d,D5>C5,IF(AND(s,d),"Red",IF(OR(s,d),"Amber","Green")))`,
              `=CHOOSE((D5>C5)+(${fell})+1,"Green","Amber","Red")`,
              `=IFS(AND(D5>C5,VLOOKUP(B5,${TA},3,FALSE)<0.75*VLOOKUP(B5,${TA},2,FALSE)),"Red",OR(D5>C5,VLOOKUP(B5,${TA},3,FALSE)<0.75*VLOOKUP(B5,${TA},2,FALSE)),"Amber",TRUE,"Green")`],
            hints: ['Sign 1 is on the row: D5>C5. Sign 2 needs two lookups of the distributor in B5, one for each month\'s sales.',
              `Sign 2 is ${fell}. Test AND of the two signs for Red first, then OR of them for Amber.`,
              'Shape: =IF(AND(sign1, sign2), "Red", IF(OR(sign1, sign2), "Amber", "Green")). LET can name the two signs so each is written once.'],
            explain: 'AND catches the facilities showing both signs, and only the rest reach OR, which catches one sign. A distributor with two facilities gets the same sales sign on both rows, because the lookup goes by name.',
            chips: ['B5', 'D5>C5', K, MY, OC, '<0.75*', '"Red"', '"Amber"', '"Green"'], fnKey: 'IF('
          });
        },
        () => {
          /* Pak Anton's route: each distributor's monthly totals come from raw lines (SUMIFS by month), its DPD from a third table. */
          const ds = r.sample(DIST(), 6).map((d) => d.name);
          /* Planted, by position in ds: [October as a share of September, DPD]. Sales drop alone; DPD alone; both; neither (it
             drops on the hidden sheet); a share slightly above 80% (its DPD rises on the hidden sheet); neither, at DPD exactly 30. */
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
          const rows = r.shuffle(lines.filter((x) => x.m === 'Sep')).concat(r.shuffle(lines.filter((x) => x.m === 'Oct'))).map((x, i) => Object.assign(x, { inv: 'INV-' + (1801 + i * 3 + r.int(0, 2)) }));
          const route = ds.slice().sort();
          const S = H.table([{ h: 'Distributor', k: 'd' }], route.map((d) => ({ d })));
          const T = H.table([{ h: 'Month', k: 'm' }, { h: 'Invoice', k: 'inv' }, { h: 'Distributor', k: 'd' }, { h: 'Net (Rp)', k: 'net', num: true }], rows, { startCol: 3 });
          const side = r.shuffle(ds.map((d, k) => ({ d, dpd: plan[k][1] })));
          const M = H.table([{ h: 'Distributor', k: 'd' }, { h: 'DPD', k: 'dpd', num: true }], side, { startCol: 8 });
          const B = board([S, T, M]);
          /* The hidden sheet: a steady distributor's October line comes in much lower, and the one slightly above 80% falls 40 days behind. */
          const alt = H.edit(B.cells, { [T.cell('Net (Rp)', rows.indexOf(octOf[ds[3]]))]: r100k(sepTotal(ds[3]) * 0.6), [M.cell('DPD', side.findIndex((x) => x.d === ds[4]))]: 40 });
          const MO = T.A('Month'), DI = T.A('Distributor'), NE = T.A('Net (Rp)'), K = M.A('Distributor'), DP = M.A('DPD'), TA = M.T('Distributor', 'DPD');
          const oct = `SUMIFS(${NE},${DI},A5,${MO},"Oct")`, sep = `SUMIFS(${NE},${DI},A5,${MO},"Sep")`;
          const f = S.next + S.first, l = S.next + S.last;
          const span = (t, a, b) => t.col(a) + t.first + ':' + t.col(b) + t.last;
          return build(B, {
            title: 'Early-warning flag per distributor from monthly sales lines', file: 'peringatan_dini.xlsx', base: B.cells, alt, fill: true, ansLabel: 'Flag',
            statement: `<p>Pak Anton is planning visits to the six distributors in <b>${S.R('Distributor')}</b>. The sales lines in <b>${span(T, 'Month', 'Net (Rp)')}</b> cover September and October, with one or more lines per distributor in each month; Bank Kenari's days past due (<b>DPD</b>) are in <b>${span(M, 'Distributor', 'DPD')}</b>.</p><p>In <b>${f}</b>, return <b>Watch</b> when the distributor's total October net sales are less than 80% of its total September net sales, <b>or</b> its DPD is more than 30. Otherwise return <b>OK</b>. Fill it down to <b>${l}</b>.</p>`,
            solution: `=IF(OR(${oct}<0.8*${sep},XLOOKUP(A5,${K},${DP})>30),"Watch","OK")`,
            alts: [`=IF(OR(SUMPRODUCT((${DI}=A5)*(${MO}="Oct")*${NE})<0.8*SUMPRODUCT((${DI}=A5)*(${MO}="Sep")*${NE}),VLOOKUP(A5,${TA},2,FALSE)>30),"Watch","OK")`,
              `=IF(AND(${oct}>=0.8*${sep},XLOOKUP(A5,${K},${DP})<=30),"OK","Watch")`,
              `=LET(o,${oct},s,${sep},d,XLOOKUP(A5,${K},${DP}),IF(OR(o<0.8*s,d>30),"Watch","OK"))`],
            hints: ['Each month\'s total is a SUMIFS over the sales lines with two conditions: the distributor in A5 and the month. Pin the ranges with $ so they stay put as you fill.',
              `${oct} is October's total. Compare it with 0.8 times September's, and test the DPD from the bank table with XLOOKUP, inside OR.`,
              `Shape: =IF(OR(${oct} < 0.8*${sep}, XLOOKUP(A5, ${K}, ${DP})>30), "Watch", "OK")`],
            explain: 'SUMIFS adds each distributor\'s lines month by month, so the test compares whole-month totals rather than single invoices. XLOOKUP reads the bank\'s DPD from the third table, and OR raises Watch on either signal. A DPD of exactly 30 is not more than 30.',
            chips: ['A5', NE, DI, MO, '"Oct"', '"Sep"', K, DP, '"Watch"', '"OK"'], fnKey: 'IF('
          });
        },
        () => {
          /* Bank facilities (main) and Rasa Kenari's overdue invoices (side): Watch on high use OR on overdue trade debt. */
          const rows = H.loans(r, 14);
          const names = [...new Set(rows.map((x) => x.distributor))];
          const od = {};
          names.forEach((n) => { od[n] = r.pick([0, 0, 0, 3500000, 8000000, 12500000, 31000000, 48000000]); });
          const use = (x) => x.outstanding / x.limit;
          const quiet = (n) => od[n] <= 25000000 && rows.filter((x) => x.distributor === n).every((x) => use(x) < 0.9);
          /* Planted: one facility at exactly 90% of its limit, a distributor owing exactly Rp 25,000,000, and two that change later. */
          const pool = r.shuffle(names);
          const atLimit = rows.find((x) => x.distributor === pool[0]);
          od[pool[0]] = 0;
          atLimit.outstanding = Math.round(0.9 * atLimit.limit);
          od[pool[1]] = 25000000;
          rows.filter((x) => x.distributor === pool[1]).forEach((x) => { if (use(x) >= 0.9) x.outstanding = r100k(x.limit * 0.62); });
          od[pool[2]] = 0;
          rows.filter((x) => x.distributor === pool[2]).forEach((x) => { if (use(x) >= 0.9) x.outstanding = r100k(x.limit * 0.5); });
          const riser = rows.find((x) => x.distributor !== pool[2] && quiet(x.distributor));
          const side = r.shuffle(names).map((n) => ({ d: n, od: od[n] }));
          /* Reading the overdue figure on the same row, instead of looking the name up, must give a wrong flag on this sheet. */
          const flag = (x, o) => use(x) >= 0.9 || o > 25000000;
          if (rows.every((x, i) => flag(x, od[x.distributor]) === flag(x, side[i] ? side[i].od : 0))) {
            const i = rows.findIndex((x, k) => k < side.length && !flag(x, od[x.distributor]));
            const j = side.findIndex((s) => s.od > 25000000);
            if (i >= 0 && j >= 0) [side[i], side[j]] = [side[j], side[i]];
          }
          const S = H.table([{ h: 'Facility', k: 'facility' }, { h: 'Distributor', k: 'distributor' }, { h: 'Limit (Rp)', k: 'limit', num: true }, { h: 'Outstanding (Rp)', k: 'outstanding', num: true }], rows);
          const M = H.table([{ h: 'Distributor', k: 'd' }, { h: 'Overdue to Rasa Kenari (Rp)', k: 'od', num: true }], side, { startCol: S.cols.length + 2 });
          const B = board([S, M]);
          /* The hidden sheet: one distributor's overdue trade debt jumps past the limit, and another facility is drawn to 95%. */
          const e = { [M.cell('Overdue to Rasa Kenari (Rp)', side.findIndex((x) => x.d === pool[2]))]: 30000000 };
          if (riser) e[S.cell('Outstanding (Rp)', rows.indexOf(riser))] = r100k(riser.limit * 0.95);
          const alt = H.edit(B.cells, e);
          const K = M.A('Distributor'), OD = M.A('Overdue to Rasa Kenari (Rp)'), TA = M.T('Distributor', 'Overdue to Rasa Kenari (Rp)');
          const f = S.next + S.first, l = S.next + S.last, side1 = M.col('Distributor') + M.first + ':' + M.col('Overdue to Rasa Kenari (Rp)') + M.last;
          return build(B, {
            title: 'Watchlist flag from utilisation or overdue invoices', file: 'daftar_pantauan.xlsx', base: B.cells, alt, fill: true, ansLabel: 'Flag',
            statement: `<p>Citra and Bu Lestari compare notes each month. Columns A:D are Bank Kenari's facilities, and a distributor can hold more than one; <b>${side1}</b> is what each distributor owes Rasa Kenari on invoices past their due date, one row per distributor, in a different order.</p><p>Flag <b>Watch</b> when the facility's <b>Outstanding</b> is at least 90% of its <b>Limit</b>, <b>or</b> the distributor owes Rasa Kenari more than ${rp(25000000)} overdue. Otherwise <b>OK</b>. In <b>${f}</b>, return the flag and fill it down to <b>${l}</b>.</p>`,
            solution: `=IF(OR(D5>=0.9*C5,XLOOKUP(B5,${K},${OD})>25000000),"Watch","OK")`,
            alts: [`=IF(OR(D5/C5>=0.9,VLOOKUP(B5,${TA},2,FALSE)>25000000),"Watch","OK")`,
              `=IF(AND(D5<0.9*C5,XLOOKUP(B5,${K},${OD})<=25000000),"OK","Watch")`,
              `=IF(OR(D5>=0.9*C5,INDEX(${OD},MATCH(B5,${K},0))>25000000),"Watch","OK")`],
            hints: ['The first signal is on the row: D5>=0.9*C5. The second sits in Rasa Kenari\'s table, so look up the distributor in B5.',
              `XLOOKUP(B5, ${K}, ${OD}) is the overdue amount. Either signal is enough, so join the two tests with OR.`,
              `Shape: =IF(OR(D5>=0.9*C5, XLOOKUP(B5, ${K}, ${OD})>25000000), "Watch", "OK")`],
            explain: 'A line drawn close to its limit and unpaid trade invoices are two early signs of a cash squeeze, from two different companies. XLOOKUP brings Rasa Kenari\'s figure onto the bank\'s row by name, and OR flags the facility on either one. A facility at exactly 90% counts, and exactly Rp 25,000,000 overdue does not.',
            chips: ['B5', 'D5>=0.9*C5', K, OD, '>25000000', '"Watch"', '"OK"'], fnKey: 'IF('
          });
        }
      ][v]();
    }
  });

  /* ======================= Weighted averages with SUMPRODUCT ======================= */

  P.add({
    key: 'weighted-average-cases', difficulty: 'Hard', level: 8, topics: ['SUMPRODUCT', 'SUMIFS', 'weighted average'], count: 3,
    make(r, v) {
      return [
        () => {
          /* The rate on one loan type, weighted by outstanding: a condition inside SUMPRODUCT over SUMIFS. */
          const rows = H.loans(r, 14);
          const TY = ['Working Capital', 'Term Loan', 'Trade Finance'];
          r.shuffle(nums(14)).forEach((k, j) => { rows[k].type = TY[j % 3]; });
          const [a, b] = r.sample(TY, 2);
          const S = H.table([{ h: 'Facility', k: 'facility' }, { h: 'Distributor', k: 'distributor' }, { h: 'Type', k: 'type' }, { h: 'Outstanding (Rp)', k: 'outstanding', num: true }, { h: 'Rate (% a year)', k: 'rate', num: true }], rows);
          const base = Object.assign({}, S.cells, { A1: 'Loan type', B1: a, A2: 'Weighted rate (%)' });
          const kb = rows.findIndex((x) => x.type === b);
          /* The hidden sheet: the other loan type, and a correction to one of its balances. */
          const alt = H.edit(base, { B1: b, [S.cell('Outstanding (Rp)', kb)]: (x) => x + 20000000 });
          const T = S.R('Type'), O = S.R('Outstanding (Rp)'), R = S.R('Rate (% a year)');
          return H.problem({
            title: 'Weighted average loan rate for one loan type', file: 'buku_pinjaman.xlsx', sheet: S, base, alt, answerCell: 'B2', precision: 4,
            statement: `<p>Dimas reports the average rate the bank earns on each loan type. A plain average would treat a small facility like a large one, so he weights each rate by its <b>Outstanding</b>.</p><p>In <b>B2</b>, return the weighted average <b>Rate (% a year)</b> for the loan type in <b>B1</b>: the sum of Outstanding times Rate for that type, divided by the total Outstanding of that type. Leave it unrounded; it is checked to 4 decimals.</p>`,
            solution: `=SUMPRODUCT((${T}=B1)*${O}*${R})/SUMIFS(${O},${T},B1)`,
            alts: [`=SUMPRODUCT(--(${T}=B1),${O},${R})/SUMIFS(${O},${T},B1)`, `=SUMPRODUCT((${T}=B1)*${O},${R})/SUMPRODUCT((${T}=B1)*${O})`, `=LET(w,(${T}=B1)*${O},SUMPRODUCT(w,${R})/SUM(w))`],
            hints: ['A weighted average is the sum of weight times value, divided by the sum of the weights. Here the weight is Outstanding and the value is Rate.',
              `(${T}=B1) is 1 on the rows of the chosen type and 0 elsewhere. Multiply it into SUMPRODUCT, and divide by SUMIFS(${O}, ${T}, B1).`,
              `Shape: =SUMPRODUCT((${T}=B1)*${O}*${R}) / SUMIFS(${O}, ${T}, B1)`],
            explain: 'The test (Type=B1) zeroes out every other loan type, so SUMPRODUCT adds Outstanding times Rate for the chosen type only. Dividing by that type\'s total Outstanding turns the sum into an average where a large facility counts for more. AVERAGEIFS would give the plain average, which is a different number.',
            chips: [T, O, R, '(' + T + '=B1)', 'SUMIFS(', '*'], fnKey: 'SUMPRODUCT('
          });
        },
        () => {
          /* A price change per SKU (new / old - 1), weighted by the cartons sold: an array of ratios inside SUMPRODUCT. */
          const rows = r.sample(H.PRODUCTS, 10).map((p) => ({ sku: p.sku, name: p.name, old: p.carton_price, qty: r.int(4, 60) * 10 }));
          rows.forEach((x) => { x.nw = Math.round(x.old * (1 + r.pick([-0.03, 0, 0.02, 0.04, 0.05, 0.07, 0.1])) / 1000) * 1000; });
          const S = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Product', k: 'name' }, { h: 'Old price (Rp)', k: 'old', num: true }, { h: 'New price (Rp)', k: 'nw', num: true }, { h: 'Cartons last quarter', k: 'qty', num: true }], rows);
          const base = Object.assign({}, S.cells, { A1: 'Weighted price change' });
          /* The hidden sheet: a revised new price on one SKU and a corrected volume on another. */
          const alt = H.edit(base, { [S.cell('New price (Rp)', 1)]: (x) => x + 6000, [S.cell('Cartons last quarter', 4)]: (x) => x + 120 });
          const O = S.R('Old price (Rp)'), N = S.R('New price (Rp)'), Q = S.R('Cartons last quarter');
          return H.problem({
            title: 'Volume-weighted average price change on a new price list', file: 'daftar_harga_baru.xlsx', sheet: S, base, alt, answerCell: 'B1', precision: 4,
            statement: `<p>Tika has the new carton price list. Pak Anton asks how much prices go up on average, counting each SKU by how much of it was sold last quarter.</p><p>Each SKU's change is <b>New price</b> divided by <b>Old price</b>, minus 1 (0.05 is a 5% rise). In <b>B1</b>, return the average of these changes weighted by <b>Cartons last quarter</b>, as a decimal. It is checked to 4 decimals.</p>`,
            solution: `=SUMPRODUCT(${Q},${N}/${O}-1)/SUM(${Q})`,
            alts: [`=SUMPRODUCT(${Q}*(${N}/${O}-1))/SUM(${Q})`, `=SUMPRODUCT(${Q},${N}/${O})/SUM(${Q})-1`, `=LET(c,${N}/${O}-1,SUMPRODUCT(${Q},c)/SUM(${Q}))`],
            hints: ['SUMPRODUCT can multiply a column by a calculation on two other columns: ' + N + '/' + O + '-1 is a list of ten changes.',
              `Weight each change by the cartons: SUMPRODUCT(${Q}, ${N}/${O}-1). Then divide by the total cartons.`,
              `Shape: =SUMPRODUCT(${Q}, ${N}/${O}-1) / SUM(${Q})`],
            explain: `${N}/${O}-1 works out every SKU's change at once, and SUMPRODUCT multiplies each one by its cartons and adds them. Dividing by the total cartons gives an average where the best sellers count most, which is what the trade will feel.`,
            chips: [Q, N, O, '/', '-1', 'SUM('], fnKey: 'SUMPRODUCT('
          });
        },
        () => {
          /* Margin weighted by net revenue, where the weight itself is cartons times price after the line discount: a
             computed weight inside SUMPRODUCT, on top and bottom. */
          const MG = { Noodles: [18, 20, 22], Coffee: [26, 28, 30], Sambal: [32, 34, 36], Soap: [24, 26] };
          const rows = r.sample(H.PRODUCTS, 10).map((p) => ({ sku: p.sku, cat: p.category, qty: r.int(5, 80) * 10, price: p.carton_price, disc: r.pick([0, 0, 2, 3, 5, 8, 10]), m: r.pick(MG[p.category]) + r.pick([0, 0.5]) }));
          /* Planted: the largest-volume SKU has the deepest discount, so leaving the discount out moves the answer. */
          rows.reduce((a, x) => (x.qty > a.qty ? x : a)).disc = 12;
          const S = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Category', k: 'cat' }, { h: 'Cartons sold', k: 'qty', num: true }, { h: 'Carton price (Rp)', k: 'price', num: true }, { h: 'Discount (%)', k: 'disc', num: true }, { h: 'Margin (%)', k: 'm', num: true }], rows);
          const base = Object.assign({}, S.cells, { A1: 'Mix margin (%)' });
          /* The hidden sheet: a corrected margin on one SKU, a larger volume on another, and a new discount on a third. */
          const alt = H.edit(base, { [S.cell('Margin (%)', 2)]: (x) => x + 2, [S.cell('Cartons sold', 6)]: (x) => x + 150, [S.cell('Discount (%)', 4)]: (x) => x + 5 });
          const Q = S.R('Cartons sold'), Pr = S.R('Carton price (Rp)'), Dc = S.R('Discount (%)'), Mg = S.R('Margin (%)');
          const net = `${Q}*${Pr}*(1-${Dc}/100)`;
          return H.problem({
            title: 'Margin of the product mix weighted by net revenue', file: 'margin_mix.xlsx', sheet: S, base, alt, answerCell: 'B1', precision: 4,
            statement: `<p>Tika needs the margin of last month's product mix for the board pack. Each SKU's net revenue is <b>Cartons sold</b> × <b>Carton price</b> × (1 - <b>Discount (%)</b>/100), and a SKU with more net revenue should weigh more. There is no net revenue column, so work it out inside the formula.</p><p>In <b>B1</b>, return the average <b>Margin (%)</b> weighted by net revenue: the sum of net revenue times Margin, divided by the sum of net revenue. Leave it unrounded; it is checked to 4 decimals.</p>`,
            solution: `=SUMPRODUCT(${net},${Mg})/SUMPRODUCT(${net})`,
            alts: [`=SUMPRODUCT(${Q},${Pr},1-${Dc}/100,${Mg})/SUMPRODUCT(${Q},${Pr},1-${Dc}/100)`, `=LET(rev,${net},SUMPRODUCT(rev,${Mg})/SUM(rev))`,
              `=SUMPRODUCT(${Q}*${Pr}*(100-${Dc})*${Mg})/SUMPRODUCT(${Q}*${Pr}*(100-${Dc}))`],
            hints: ['The weight is not a column on the sheet: it is net revenue, Cartons times Price times (1 - Discount/100), row by row.',
              `${net} is the net revenue of every SKU at once. SUMPRODUCT(that, ${Mg}) is the top; SUMPRODUCT(that) on its own is the bottom.`,
              `Shape: =SUMPRODUCT(${net}, ${Mg}) / SUMPRODUCT(${net})`],
            explain: 'The top adds net revenue times margin for every SKU and the bottom adds the net revenue, so the result is the margin of the whole mix. Leaving out the discount would give a heavily discounted SKU more weight than the money it brings in, and weighting by cartons alone would make a cheap noodle carton count as much as a soap carton priced at more than twice as much.',
            chips: [Q, Pr, Dc, Mg, '(1-' + Dc + '/100)', 'SUMPRODUCT(', '/'], fnKey: 'SUMPRODUCT('
          });
        }
      ][v]();
    }
  });

  /* ======================= Commission and bonus from rate tables ======================= */

  P.add({
    key: 'commission-rate-tables', difficulty: 'Hard', level: 9, topics: ['XLOOKUP', 'VLOOKUP', 'INDEX', 'MATCH', 'approximate match', 'commission'], count: 3,
    make(r, v) {
      return [
        () => {
          /* A progressive table (base plus a rate on the excess): three approximate lookups on the same row. */
          const t1 = r.pick([75, 100]) * 1e6, t2 = r.pick([150, 175]) * 1e6, t3 = r.pick([225, 250]) * 1e6;
          const rate = r.pick([[1, 1.5, 2, 3], [0.5, 1, 2, 2.5], [1, 2, 2.5, 3.5]]);
          const from = [0, t1, t2, t3], base = [0];
          for (let k = 1; k < 4; k++) base.push(base[k - 1] + (from[k] - from[k - 1]) * rate[k - 1] / 100);
          const reps = H.reps(r, 10).map((x) => ({ name: x.name, area: x.area, sales: x.actual }));
          /* Planted: sales exactly on the second limit, one in the top tier, one in the third. */
          reps[0].sales = t1;
          reps[1].sales = t3 + r.int(1, 40) * 1e6;
          reps[2].sales = t2 + r.int(1, 40) * 1e6;
          const rows = r.shuffle(reps);
          const S = H.table([{ h: 'Rep', k: 'name' }, { h: 'Area', k: 'area' }, { h: 'Net sales (Rp)', k: 'sales', num: true }], rows);
          const M = H.table([{ h: 'From (Rp)', k: 'f', num: true }, { h: 'Base (Rp)', k: 'b', num: true }, { h: 'Rate on excess (%)', k: 'r', num: true }], from.map((f, k) => ({ f, b: base[k], r: rate[k] })), { startCol: S.cols.length + 2 });
          const B = board([S, M]);
          /* The hidden sheet: the top rate goes up half a point, and one rep's sales cross into the next tier. */
          const mover = rows.find((x) => x.sales < t2 - 20e6 && x.sales > t1);
          const e = { [M.cell('Rate on excess (%)', 3)]: rate[3] + 0.5 };
          if (mover) e[S.cell('Net sales (Rp)', rows.indexOf(mover))] = t2 + 5e6;
          const alt = H.edit(B.cells, e);
          const F = M.A('From (Rp)'), Bs = M.A('Base (Rp)'), Rt = M.A('Rate on excess (%)'), TA = M.T('From (Rp)', 'Rate on excess (%)');
          const f = S.next + S.first, l = S.next + S.last, side1 = M.col('From (Rp)') + M.first + ':' + M.col('Rate on excess (%)') + M.last;
          const xl = (c) => `XLOOKUP(C5,${F},${c},,-1)`;
          return build(B, {
            title: 'Commission from a base-plus-excess rate table', file: 'komisi_progresif.xlsx', base: B.cells, alt, fill: true, ansLabel: 'Commission (Rp)',
            statement: `<p>Rasa Kenari pays sales commission like a tax table, so a rep never earns less by selling more. In <b>${side1}</b>, each row starts at <b>From</b>; a rep whose sales reach a row (at least its From) gets that row's <b>Base</b>, plus its <b>Rate on excess</b> on the part of sales above its From.</p><p>In <b>${f}</b>, work out each rep's commission: Base + (Net sales - From) × Rate / 100, all from the row the sales fall in, rounded to whole rupiah with ROUND. Fill it down to <b>${l}</b>.</p>`,
            solution: `=ROUND(${xl(Bs)}+(C5-${xl(F)})*${xl(Rt)}/100,0)`,
            alts: [`=ROUND(VLOOKUP(C5,${TA},2,TRUE)+(C5-VLOOKUP(C5,${TA},1,TRUE))*VLOOKUP(C5,${TA},3,TRUE)/100,0)`,
              `=LET(k,MATCH(C5,${F},1),ROUND(INDEX(${Bs},k)+(C5-INDEX(${F},k))*INDEX(${Rt},k)/100,0))`,
              `=ROUND(LOOKUP(C5,${F},${Bs})+(C5-LOOKUP(C5,${F},${F}))*LOOKUP(C5,${F},${Rt})/100,0)`],
            hints: ['All three numbers come from the same row of the table: the last From that the sales reach. That is an approximate match, not an exact one.',
              `${xl(Bs)} returns the Base; swap the last range for ${F} or ${Rt} to get the From and the Rate. match_mode -1 means exact or next smaller.`,
              `Shape: =ROUND(Base + (C5 - From) * Rate / 100, 0), with each of Base, From and Rate an XLOOKUP(C5, ${F}, ..., , -1).`],
            explain: 'The approximate match picks the tier the sales fall in, and the formula charges that tier\'s rate only on the part above its starting point. The Base already holds the commission earned on the tiers below, so the total rises smoothly as sales grow.',
            chips: ['C5', F, Bs, Rt, ',,-1)', TA, ',TRUE)'], fnKey: 'XLOOKUP('
          });
        },
        () => {
          /* A two-way rate grid: the category picks the row (exact) and the cartons pick the column (approximate). */
          const rows = H.sales(r, 14).map((s) => Object.assign({}, s, { disc: s.discount_pct }));
          const th = r.pick([[0, 20, 40], [0, 15, 30], [0, 25, 50]]);
          const RT = { Noodles: [1, 1.5, 2], Coffee: [1.5, 2, 3], Sambal: [2, 2.5, 3.5], Soap: [1, 2, 2.5] };
          const cats = r.shuffle(H.CATEGORIES);
          const vals = cats.map((c) => RT[c].map((x, j) => x + (j && r.chance(0.3) ? 0.5 : 0)));
          const netOf = (x) => r100k(x.cartons * x.price * (1 - x.disc / 100)) || 100000;
          /* Planted: one line exactly on the top carton limit. */
          rows[3].cartons = th[2];
          rows.forEach((x) => { x.net = netOf(x); });
          const S = H.table([{ h: 'Invoice', k: 'invoice' }, { h: 'Category', k: 'category' }, { h: 'Cartons', k: 'cartons', num: true }, { h: 'Net (Rp)', k: 'net', num: true }], rows);
          const sc = S.cols.length + 2;
          const G = grid('Category', th, cats.map((c, i) => ({ h: c, vals: vals[i] })), sc);
          const B = board([S, G]);
          /* The hidden sheet: a line moves to exactly the middle limit, and one category's top rate rises. */
          const k = Math.max(0, rows.findIndex((x, i) => i !== 3 && x.cartons < th[1]));
          const moved = Object.assign({}, rows[k], { cartons: th[1] });
          const ci = cats.indexOf(rows[0].category);
          const e = { [S.cell('Cartons', k)]: th[1], [S.cell('Net (Rp)', k)]: netOf(moved), [L[sc + 3] + (5 + ci)]: (x) => x + 0.5, [L[sc + 1] + (5 + ci)]: (x) => x + 0.25 };
          const alt = H.edit(B.cells, e);
          const c0 = L[sc], c1 = L[sc + 1], c3 = L[sc + 3], last = 4 + cats.length;
          const GR = `$${c1}$5:$${c3}$${last}`, CK = `$${c0}$5:$${c0}$${last}`, TH = `$${c1}$4:$${c3}$4`, ALL = `$${c0}$5:$${c3}$${last}`, WIDE = `$${c1}$4:$${c3}$${last}`;
          const f = S.next + S.first, l = S.next + S.last;
          return build(B, {
            title: 'Commission rate by category and carton tier', file: 'komisi_kategori.xlsx', base: B.cells, alt, fill: true, ansLabel: 'Commission (Rp)',
            statement: `<p>The trade team pays commission on each order line. The grid in <b>${c0}4:${c3}${last}</b> has a row per <b>Category</b>; across the top are carton limits, and a line takes the column of the highest limit its <b>Cartons</b> reach (at least that many cartons).</p><p>In <b>${f}</b>, return each line's commission: its <b>Net (Rp)</b> times the rate (%) where its category's row meets its carton column, divided by 100 and rounded to whole rupiah with ROUND. Fill it down to <b>${l}</b>.</p>`,
            solution: `=ROUND(D5*INDEX(${GR},MATCH(B5,${CK},0),MATCH(C5,${TH},1))/100,0)`,
            alts: [`=ROUND(D5*XLOOKUP(C5,${TH},XLOOKUP(B5,${CK},${GR}),,-1)/100,0)`,
              `=ROUND(D5*VLOOKUP(B5,${ALL},MATCH(C5,${TH},1)+1,FALSE)/100,0)`,
              `=ROUND(D5*HLOOKUP(C5,${WIDE},MATCH(B5,${CK},0)+1,TRUE)/100,0)`],
            hints: ['Two positions locate the rate: the category gives the row (an exact match) and the cartons give the column (an approximate match on the limits).',
              `MATCH(B5, ${CK}, 0) is the row and MATCH(C5, ${TH}, 1) the column: with 1, MATCH returns the last limit that is not above C5. INDEX takes both.`,
              `Shape: =ROUND(D5 * INDEX(${GR}, MATCH(B5, ${CK}, 0), MATCH(C5, ${TH}, 1)) / 100, 0)`],
            explain: 'The exact MATCH finds the category\'s row and the approximate MATCH finds the carton tier, so INDEX returns the one rate that applies to the line. A line exactly on a limit takes the higher column, because MATCH with 1 accepts an equal value.',
            chips: ['D5*', 'B5', 'C5', GR, CK, TH, ',0)', ',1)'], fnKey: 'INDEX('
          });
        },
        () => {
          /* One rep's bonus: lookups for the rep's figures, an approximate lookup for the tier, and a visit gate. */
          const reps = H.reps(r, 10);
          const tiers = [[0, 0], [0.9, r.pick([0.5, 0.75])], [1, r.pick([1, 1.25])], [1.1, r.pick([1.5, 2])]];
          const ach = (x) => x.actual / x.target;
          const [a, b] = r.sample(nums(10), 2);
          /* Planted: rep a has exactly 30 visits (paid) and at least 90% achievement; rep b reaches target but made 26 visits. */
          reps[a].visits = 30;
          if (ach(reps[a]) < 0.9) reps[a].actual = r100k(reps[a].target * r.pick([0.95, 1.04, 1.12]));
          reps[b].visits = r.pick([24, 26, 29]);
          if (ach(reps[b]) < 1) reps[b].actual = r100k(reps[b].target * r.pick([1.04, 1.12]));
          const S = H.table([{ h: 'Rep', k: 'name' }, { h: 'Area', k: 'area' }, { h: 'Target (Rp)', k: 'target', num: true }, { h: 'Actual (Rp)', k: 'actual', num: true }, { h: 'Visits', k: 'visits', num: true }], reps);
          const M = H.table([{ h: 'From achievement', v: (t) => t[0], num: true }, { h: 'Bonus rate (%)', v: (t) => t[1], num: true }], tiers, { startCol: S.cols.length + 1 });
          const B = board([S, M]);
          const base = Object.assign({}, B.cells, { A1: 'Rep', B1: reps[a].name, A2: 'Bonus (Rp)' });
          /* The hidden sheet: the rep who hit target with too few visits, so the bonus is 0. */
          const alt = H.edit(base, { B1: reps[b].name });
          const N = S.R('Rep'), Tg = S.R('Target (Rp)'), Ac = S.R('Actual (Rp)'), Vi = S.R('Visits'), Fr = M.R('From achievement'), Br = M.R('Bonus rate (%)');
          const act = `XLOOKUP(B1,${N},${Ac})`, tgt = `XLOOKUP(B1,${N},${Tg})`, vis = `XLOOKUP(B1,${N},${Vi})`;
          const va = (c) => `VLOOKUP(B1,${N.split(':')[0]}:${S.col('Visits') + S.last},${c},FALSE)`;
          return build(B, {
            title: 'Quarter bonus for one rep with a visit gate', file: 'bonus_sales.xlsx', base, alt, answerCell: 'B2',
            statement: `<p>Pak Anton is answering a rep who asked about the quarter bonus. For the rep in <b>B1</b>: achievement is <b>Actual</b> divided by <b>Target</b> (1 means 100% of target), and the bonus rate is the one in <b>${Fr.split(':')[0]}:${Br.split(':')[1]}</b> for the highest <b>From achievement</b> reached (at least that much). The bonus is Actual times that rate / 100, rounded to whole rupiah with ROUND.</p><p>One gate: a rep with fewer than 30 <b>Visits</b> gets no bonus (0), whatever the achievement. Return the bonus in <b>B2</b>.</p>`,
            solution: `=IF(${vis}<30,0,ROUND(${act}*XLOOKUP(${act}/${tgt},${Fr},${Br},,-1)/100,0))`,
            alts: [`=LET(a,${act},t,${tgt},IF(${vis}<30,0,ROUND(a*XLOOKUP(a/t,${Fr},${Br},,-1)/100,0)))`,
              `=IF(${va(5)}<30,0,ROUND(${va(4)}*VLOOKUP(${va(4)}/${va(3)},${Fr.split(':')[0]}:${Br.split(':')[1]},2,TRUE)/100,0))`,
              `=IF(${vis}>=30,ROUND(${act}*LOOKUP(${act}/${tgt},${Fr},${Br})/100,0),0)`],
            hints: ['Three facts about the rep come from the main table (Target, Actual, Visits), each an exact lookup on B1. The rate comes from the tier table, an approximate one.',
              `${act}/${tgt} is the achievement. Look it up in ${Fr} with match_mode -1 to get the rate, and put the visit test in an IF around everything.`,
              `Shape: =IF(visits<30, 0, ROUND(actual * XLOOKUP(actual/target, ${Fr}, ${Br}, , -1) / 100, 0))`],
            explain: 'The exact lookups pull the rep\'s figures, and the approximate XLOOKUP turns the achievement into the tier rate. The IF checks the visit gate first, so a rep who sold well but visited too few customers gets 0. Exactly 30 visits passes the gate.',
            chips: ['B1', N, Ac, Tg, Vi, Fr, Br, ',,-1)', '<30'], fnKey: 'IF('
          });
        }
      ][v]();
    }
  });

  /* ======================= Collection priority ======================= */

  P.add({
    key: 'collection-priority', difficulty: 'Hard', level: 10, topics: ['XLOOKUP', 'COUNTIF', 'INT', 'IF', 'SORTBY', 'TAKE', 'collection'], count: 3,
    make(r, v) {
      return [
        () => {
          /* A score from three tables: points for days overdue (approximate lookup), the amount (INT), the watchlist (COUNTIF). */
          const inv = H.invoices(r, 14).map((x) => Object.assign(x, { due: ser(x.date) + x.terms })).sort((p, q) => p.due - q.due || (p.ref < q.ref ? -1 : 1));
          const asof = inv[inv.length - 1].due - r.int(2, 5);
          const pts = r.pick([[[0, 0], [1, 1], [31, 2], [61, 4], [91, 6]], [[0, 0], [1, 1], [31, 3], [61, 5], [91, 8]]]);
          const names = [...new Set(inv.map((x) => x.distributor))];
          const watch = r.sample(names, 3);
          const out = names.filter((n) => !watch.includes(n));
          /* Planted: an invoice exactly 30 days overdue (still the 1 to 30 tier). */
          inv[r.int(2, 5)].due = asof - 30;
          const S = H.table([{ h: 'Invoice', k: 'ref' }, { h: 'Distributor', k: 'distributor' }, { h: 'Due date', k: 'due', date: true }, { h: 'Amount (Rp)', k: 'amount', num: true }], inv);
          const M = H.table([{ h: 'From days', v: (t) => t[0], num: true }, { h: 'Points', v: (t) => t[1], num: true }], pts, { startCol: S.cols.length + 2 });
          const W = H.table([{ h: 'Watchlist', k: 'd' }], watch.map((d) => ({ d })), { startCol: S.cols.length + 5 });
          const B = board([S, M, W]);
          const base = Object.assign({}, B.cells, { A1: 'As of', B1: asof });
          /* The hidden sheet: twelve days later, and one name on the watchlist replaced. */
          const alt = H.edit(base, { B1: asof + 12, [W.cell('Watchlist', 1)]: r.pick(out) });
          const Fd = M.A('From days'), Pt = M.A('Points'), WL = W.A('Watchlist'), TA = M.T('From days', 'Points');
          const f = S.next + S.first, l = S.next + S.last;
          return build(B, {
            title: 'Collection priority score per overdue invoice', file: 'antrian_penagihan.xlsx', base, alt, fill: true, ansLabel: 'Score', dateCells: ['B1'],
            statement: `<p>Bu Lestari's team calls the highest scores first. For each open invoice, as of the date in <b>B1</b>:</p><p>Score = the <b>Points</b> for its days overdue (B1 minus <b>Due date</b>; an invoice not yet due counts as 0 days) from the table in <b>${M.col('From days') + M.first}:${M.col('Points') + M.last}</b>, where a row applies from its <b>From days</b> on; <b>plus</b> 1 point for each full ${rp(5000000)} of <b>Amount</b>; <b>plus</b> 3 points when the distributor is on the <b>Watchlist</b> in <b>${W.col('Watchlist') + W.first}:${W.col('Watchlist') + W.last}</b>. In <b>${f}</b>, return the score and fill it down to <b>${l}</b>.</p>`,
            solution: `=XLOOKUP(MAX(0,$B$1-C5),${Fd},${Pt},,-1)+INT(D5/5000000)+IF(COUNTIF(${WL},B5)>0,3,0)`,
            alts: [`=VLOOKUP(MAX(0,$B$1-C5),${TA},2,TRUE)+INT(D5/5000000)+3*COUNTIF(${WL},B5)`,
              `=LOOKUP(MAX(0,$B$1-C5),${Fd},${Pt})+ROUNDDOWN(D5/5000000,0)+IF(ISNUMBER(MATCH(B5,${WL},0)),3,0)`,
              `=LET(d,MAX(0,$B$1-C5),XLOOKUP(d,${Fd},${Pt},,-1)+INT(D5/5000000)+IF(COUNTIF(${WL},B5)>0,3,0))`],
            hints: ['Build the score in three parts and add them. Days overdue is $B$1-C5, kept at 0 or more with MAX(0, ...).',
              `Part 1: XLOOKUP(MAX(0,$B$1-C5), ${Fd}, ${Pt}, , -1). Part 2: INT(D5/5000000) counts full Rp 5,000,000 blocks. Part 3: COUNTIF(${WL}, B5)>0 says whether the name is on the list.`,
              `Shape: =XLOOKUP(MAX(0,$B$1-C5), ${Fd}, ${Pt}, , -1) + INT(D5/5000000) + IF(COUNTIF(${WL}, B5)>0, 3, 0)`],
            explain: 'The approximate XLOOKUP turns days overdue into points by tier, INT drops the part of the amount below the next full block, and COUNTIF checks the name against the watchlist. $B$1 stays pinned so every row is scored on the same date.',
            chips: ['$B$1-C5', 'MAX(0,', Fd, Pt, ',,-1)', 'INT(D5/5000000)', WL, 'COUNTIF('], fnKey: 'XLOOKUP('
          });
        },
        () => {
          /* A queue from a chain of tables: DPD to Kol (approximate), Kol to provision rate (exact), then expected loss in an OR. */
          const rows = H.loans(r, 14).map((x) => ({ f: x.facility, d: x.distributor, out: x.outstanding, dpd: x.dpd }));
          const KOL = [[0, 1], [1, 2], [91, 3], [121, 4], [181, 5]];
          const prov = [1, 5, r.pick([15, 20]), 50, 100];
          /* Planted: DPD exactly 90 on a small line, DPD 91, a Kol 2 line between Rp 62.5 and 100 million, a clean line, and a
             Kol 2 line of exactly Rp 100 million, whose expected loss is exactly Rp 5,000,000 (exact in floating point too),
             and a small Kol 3 line past 90 days whose expected loss stays under Rp 5,000,000, so only the DPD test calls it. */
          const [p0, p1, p2, p3, p4, p5] = r.sample(nums(14), 6);
          rows[p0].dpd = 90; rows[p0].out = r.int(20, 50) * 1000000;
          rows[p1].dpd = 91;
          rows[p2].dpd = r.pick([15, 30, 45]); rows[p2].out = r.int(65, 95) * 1000000;
          rows[p3].dpd = 0;
          rows[p4].dpd = r.pick([7, 21, 60]); rows[p4].out = 100000000;
          rows[p5].dpd = r.pick([95, 105, 115]); rows[p5].out = r.int(10, 20) * 1000000;
          const S = H.table([{ h: 'Facility', k: 'f' }, { h: 'Distributor', k: 'd' }, { h: 'Outstanding (Rp)', k: 'out', num: true }, { h: 'DPD', k: 'dpd', num: true }], rows);
          const sc = S.cols.length + 2;
          const M = H.table([{ h: 'From DPD', v: (t) => t[0], num: true }, { h: 'Kol', v: (t) => t[1], num: true }], KOL, { startCol: sc });
          /* The provision table lists the worst grade first, so it does not line up row for row with the DPD table: a lookup
             that skips the Kol step and reads the rate beside the matched DPD row gives the wrong rate. */
          const provRows = [5, 4, 3, 2, 1].map((k) => ({ k, p: prov[k - 1] }));
          const N = H.table([{ h: 'Kol', k: 'k', num: true }, { h: 'Provision (%)', k: 'p', num: true }], provRows, { startCol: sc + 3 });
          const B = board([S, M, N]);
          /* The hidden sheet: Kol 2 provisions rise to 8%, and the clean line falls 15 days behind. */
          const alt = H.edit(B.cells, { [N.cell('Provision (%)', provRows.findIndex((x) => x.k === 2))]: 8, [S.cell('DPD', p3)]: 15 });
          const Fd = M.A('From DPD'), Ko = M.A('Kol'), K2 = N.A('Kol'), Pv = N.A('Provision (%)');
          const kol = `XLOOKUP(D5,${Fd},${Ko},,-1)`, el = `C5*XLOOKUP(${kol},${K2},${Pv})/100`;
          const T1 = M.T('From DPD', 'Kol'), T2 = N.T('Kol', 'Provision (%)');
          const f = S.next + S.first, l = S.next + S.last;
          return build(B, {
            title: 'Collection queue from DPD, Kol and provision rate', file: 'antrian_penagihan.xlsx', base: B.cells, alt, fill: true, ansLabel: 'Priority',
            statement: `<p>Citra sorts the bank's facilities into three call lists. A facility's Kol grade comes from its <b>DPD</b> through the table in <b>${M.col('From DPD') + M.first}:${M.col('Kol') + M.last}</b> (a row applies from its From DPD on), and each Kol has a provision rate in <b>${N.col('Kol') + N.first}:${N.col('Provision (%)') + N.last}</b>. Expected loss = <b>Outstanding</b> × provision rate / 100.</p><p>Return <b>Call today</b> when the expected loss is at least ${rp(5000000)} <b>or</b> DPD is more than 90; otherwise <b>This week</b> when DPD is more than 0; otherwise <b>Monitor</b>. Start in <b>${f}</b> and fill down to <b>${l}</b>.</p>`,
            solution: `=IF(OR(${el}>=5000000,D5>90),"Call today",IF(D5>0,"This week","Monitor"))`,
            alts: [`=LET(k,${kol},e,C5*XLOOKUP(k,${K2},${Pv})/100,IF(OR(e>=5000000,D5>90),"Call today",IF(D5>0,"This week","Monitor")))`,
              `=IF(OR(C5*VLOOKUP(VLOOKUP(D5,${T1},2,TRUE),${T2},2,FALSE)/100>=5000000,D5>90),"Call today",IF(D5>0,"This week","Monitor"))`,
              `=IFS(OR(${el}>=5000000,D5>90),"Call today",D5>0,"This week",TRUE,"Monitor")`],
            hints: ['Work in a chain: DPD gives Kol (an approximate lookup), Kol gives the provision rate (an exact lookup), and the rate gives the expected loss.',
              `${kol} is the Kol. Use it as the value to find in ${K2}, and multiply the rate by C5/100: ${el}.`,
              'Shape: =IF(OR(expected loss>=5000000, D5>90), "Call today", IF(D5>0, "This week", "Monitor"))'],
            explain: 'The inner XLOOKUP grades the facility by its arrears, and the outer one turns the grade into a provision rate, so the expected loss follows the bank\'s own tables. OR puts a large expected loss or a long arrear on today\'s list, and the second IF splits the rest between this week and monitoring.',
            chips: ['C5', 'D5', Fd, Ko, K2, Pv, ',,-1)', '>=5000000', '"Call today"', '"This week"', '"Monitor"'], fnKey: 'IF('
          });
        },
        () => {
          /* The three invoices to chase first: a computed key (unpaid x days overdue x amount) inside SORTBY, cut with TAKE. */
          const inv = H.invoices(r, 14).map((x) => Object.assign(x, { due: ser(x.date) + x.terms })).sort((p, q) => p.due - q.due || (p.ref < q.ref ? -1 : 1));
          const asof = inv[inv.length - 1].due - r.int(0, 6);
          inv.forEach((x, i) => { x.paid = i % 3 === 1 || (i % 3 === 2 && r.chance(0.3)) ? Math.min(asof - 1, x.due + r.int(-6, 9)) : ''; });
          const score = (x, d) => (x.paid === '' ? (d - x.due) * x.amount : 0);
          /* No ties among the scores, so the order never depends on the row order. */
          const seen = new Set();
          inv.forEach((x) => { while (x.paid === '' && (seen.has(score(x, asof)) || seen.has(score(x, asof + 15)))) x.amount += 100000; if (x.paid === '') { seen.add(score(x, asof)); seen.add(score(x, asof + 15)); } });
          /* Planted: an early paid invoice large enough to make the top three if the Paid on test were left out. */
          const third = inv.map((x) => score(x, asof)).sort((p, q) => q - p)[2];
          while (asof > inv[1].due && (asof - inv[1].due) * inv[1].amount <= third) inv[1].amount += 1000000;
          const S = H.table([{ h: 'Invoice', k: 'ref' }, { h: 'Distributor', k: 'distributor' }, { h: 'Due date', k: 'due', date: true }, { h: 'Amount (Rp)', k: 'amount', num: true }, { h: 'Paid on', v: (x) => (x.paid === '' ? undefined : x.paid) }], inv);
          const top = inv.slice().sort((p, q) => score(q, asof) - score(p, asof))[0];
          const base = Object.assign({}, S.cells, { A1: 'As of', B1: asof });
          /* The hidden sheet: fifteen days later, and the top invoice has been paid. */
          const alt = H.edit(base, { B1: asof + 15, [S.cell('Paid on', inv.indexOf(top))]: asof + 2 });
          const paidRefs = inv.map((x, i) => (x.paid === '' ? null : S.cell('Paid on', i))).filter(Boolean);
          const A = S.R('Invoice'), C = S.R('Due date'), Dm = S.R('Amount (Rp)'), E = S.R('Paid on');
          const key = `(${E}="")*(B1-${C})*${Dm}`;
          return H.problem({
            title: 'Top three invoices to chase, by amount times days late', file: 'antrian_penagihan.xlsx', sheet: S, base, alt, spill: { cols: 1, rows: 3 }, ansLabel: 'Chase first', dateCells: ['B1'].concat(paidRefs),
            statement: `<p>Bu Lestari has time for three calls this morning. An invoice is still open when <b>Paid on</b> is empty. Its urgency is its <b>Amount</b> times its days overdue as of <b>B1</b> (B1 minus <b>Due date</b>); a paid invoice has no urgency.</p><p>In <b>${S.next + S.first}</b>, list the <b>Invoice</b> numbers of the three most urgent open invoices, most urgent first, in one formula that spills down.</p>`,
            solution: `=TAKE(SORTBY(${A},${key},-1),3)`,
            alts: [`=LET(u,${key},TAKE(SORTBY(${A},u,-1),3))`, `=TAKE(SORTBY(${A},${Dm}*(B1-${C})*(${E}=""),-1),3)`,
              `=TAKE(SORTBY(FILTER(${A},${E}=""),FILTER((B1-${C})*${Dm},${E}=""),-1),3)`],
            hints: [`Build the urgency for every row at once: (${E}="") is 1 for open invoices, B1-${C} is the days overdue, and ${Dm} is the amount. Multiply them.`,
              'SORTBY sorts one range by another; -1 puts the largest key first. TAKE then keeps the first rows.',
              `Shape: =TAKE(SORTBY(${A}, ${key}, -1), 3)`],
            explain: 'The key multiplies three columns row by row, so a paid invoice scores 0 and a large, long-overdue one scores highest. SORTBY orders the invoice numbers by that key without a helper column, and TAKE keeps the top three.',
            chips: [A, C, Dm, E, '(' + E + '="")', 'B1-' + C, ',-1)', ',3)'], fnKey: 'SORTBY('
          });
        }
      ][v]();
    }
  });

  /* ======================= Reorder plans from stock and lead time ======================= */

  P.add({
    key: 'reorder-plans', difficulty: 'Hard', level: 8, topics: ['IF', 'ROUNDUP', 'MAX', 'INT', 'SUMPRODUCT', 'XLOOKUP', 'stock'], count: 3,
    make(r, v) {
      return [
        () => {
          /* Reorder when stock is at or below the reorder point; top up for lead time plus cover, in full pallets. */
          const rows = H.stock(r, 12).map((x) => ({ sku: x.sku, wh: x.warehouse, daily: r.int(3, 25), lead: x.lead }));
          const low = new Set(r.sample(nums(12), 5));
          rows.forEach((x, i) => { const rop = x.daily * x.lead; x.oh = low.has(i) ? Math.max(0, rop - r.int(1, 12) * 5) : rop + r.int(2, 30) * 10; });
          /* Planted: one SKU exactly on its reorder point (it orders), and one that dips below it on the hidden sheet. */
          const [eq, dip] = [...low].slice(0, 1).concat(nums(12).filter((i) => !low.has(i)).slice(0, 1));
          rows[eq].oh = rows[eq].daily * rows[eq].lead;
          const cover = r.pick([14, 21]), pallet = r.pick([20, 24, 30]);
          const S = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Warehouse', k: 'wh' }, { h: 'On hand', k: 'oh', num: true }, { h: 'Daily sales', k: 'daily', num: true }, { h: 'Lead time (days)', k: 'lead', num: true }], rows);
          const base = Object.assign({}, S.cells, { A1: 'Cover days', B1: cover, A2: 'Cartons per pallet', B2: pallet });
          /* The hidden sheet: a longer cover, another pallet size, and one SKU sells down below its reorder point. */
          const alt = H.edit(base, { B1: cover + 7, B2: pallet === 24 ? 30 : 24, [S.cell('On hand', dip)]: Math.max(0, rows[dip].daily * rows[dip].lead - 15) });
          const f = S.next + S.first, l = S.next + S.last;
          return H.problem({
            title: 'Reorder quantity rounded up to full pallets', file: 'rencana_pesan.xlsx', sheet: S, base, alt, fill: true, ansLabel: 'Order (cartons)',
            statement: `<p>Rasa Kenari's stock planner reviews every SKU each morning. A SKU needs an order when its <b>On hand</b> is at or below its reorder point, which is <b>Daily sales</b> × <b>Lead time</b>. It then orders enough to last the lead time plus the cover days in <b>B1</b>: Daily sales × (Lead time + B1) - On hand, rounded <b>up</b> to full pallets of <b>B2</b> cartons. A SKU above its reorder point orders 0.</p><p>In <b>${f}</b>, return the order in cartons and fill it down to <b>${l}</b>.</p>`,
            solution: '=IF(C5<=D5*E5,ROUNDUP((D5*(E5+$B$1)-C5)/$B$2,0)*$B$2,0)',
            alts: ['=IF(C5<=D5*E5,CEILING(D5*(E5+$B$1)-C5,$B$2),0)', '=IF(C5>D5*E5,0,CEILING.MATH(D5*(E5+$B$1)-C5,$B$2))', '=(C5<=D5*E5)*ROUNDUP((D5*(E5+$B$1)-C5)/$B$2,0)*$B$2'],
            hints: ['The test is C5<=D5*E5. The need is D5*(E5+$B$1)-C5 cartons; pin B1 and B2 with $ so they stay put as you fill.',
              'To round up to a full pallet, divide by $B$2, round up with ROUNDUP(..., 0), and multiply back by $B$2. CEILING(need, $B$2) does it in one step.',
              'Shape: =IF(C5<=D5*E5, ROUNDUP((D5*(E5+$B$1)-C5)/$B$2, 0)*$B$2, 0)'],
            explain: 'IF orders only when stock has reached the reorder point, so a SKU sitting exactly on it still orders. The need covers the days until the truck arrives plus the cover days, and rounding up to whole pallets means the warehouse never receives a part pallet or runs short.',
            chips: ['C5<=D5*E5', 'D5*(E5+$B$1)-C5', '$B$2', 'ROUNDUP(', ',0)', 'CEILING('], fnKey: 'IF('
          });
        },
        () => {
          /* The latest order date: full days of cover (INT) less the warehouse's lead time (lookup), never before the as-of date. */
          const rows = H.stock(r, 12).map((x) => ({ sku: x.sku, wh: x.warehouse, oh: r.int(4, 60) * 10, daily: r.int(4, 25) }));
          const whs = r.shuffle(H.WAREHOUSES);
          const lead = whs.map(() => r.pick([3, 5, 7, 10, 14]));
          const leadOf = (w) => lead[whs.indexOf(w)];
          /* Planted: two SKUs already late (cover below lead time), so the answer is the as-of date itself. */
          r.sample(nums(12), 2).forEach((i) => { rows[i].oh = Math.max(10, Math.floor((leadOf(rows[i].wh) - 2) * rows[i].daily / 10) * 10); });
          const asof = ser('2026-11-02') + r.int(0, 10);
          const S = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Warehouse', k: 'wh' }, { h: 'On hand', k: 'oh', num: true }, { h: 'Daily sales', k: 'daily', num: true }], rows);
          const M = H.table([{ h: 'Warehouse', k: 'w' }, { h: 'Lead time (days)', k: 'l', num: true }], whs.map((w, i) => ({ w, l: lead[i] })), { startCol: S.cols.length + 2 });
          const B = board([S, M]);
          const base = Object.assign({}, B.cells, { A1: 'As of', B1: asof });
          /* The hidden sheet: a week later, and one warehouse's supplier needs three more days. */
          const wk = whs.indexOf(rows[0].wh);
          const alt = H.edit(base, { B1: asof + 7, [M.cell('Lead time (days)', wk)]: lead[wk] + 3 });
          const K = M.A('Warehouse'), LT = M.A('Lead time (days)'), TA = M.T('Warehouse', 'Lead time (days)');
          const f = S.next + S.first, l = S.next + S.last;
          return build(B, {
            title: 'Latest order date per SKU from cover and lead time', file: 'rencana_pesan.xlsx', base, alt, fill: true, ansLabel: 'Order by', ansDate: true, dateCells: ['B1'],
            statement: `<p>The planner wants an order-by date for every SKU. Each warehouse's supplier needs a set number of days to deliver (<b>${M.col('Warehouse') + M.first}:${M.col('Lead time (days)') + M.last}</b>). A SKU's stock lasts the full days of <b>On hand</b> divided by <b>Daily sales</b> (drop any part day).</p><p>In <b>${f}</b>, return the latest date to place the order: the date in <b>B1</b> plus the full days of stock, minus the lead time of the SKU's <b>Warehouse</b>. When that date is before B1, the order is already late, so return B1 itself. The answer is a date; fill it down to <b>${l}</b>.</p>`,
            solution: `=MAX($B$1,$B$1+INT(C5/D5)-XLOOKUP(B5,${K},${LT}))`,
            alts: [`=MAX($B$1,$B$1+INT(C5/D5)-VLOOKUP(B5,${TA},2,FALSE))`, `=IF(INT(C5/D5)<XLOOKUP(B5,${K},${LT}),$B$1,$B$1+INT(C5/D5)-XLOOKUP(B5,${K},${LT}))`, `=MAX($B$1,$B$1+ROUNDDOWN(C5/D5,0)-XLOOKUP(B5,${K},${LT}))`],
            hints: [`Full days of stock is INT(C5/D5). The lead time depends on the warehouse, so look up B5 in ${K}.`,
              `$B$1+INT(C5/D5)-XLOOKUP(B5, ${K}, ${LT}) is the order-by date. MAX with $B$1 stops it falling before the as-of date.`,
              `Shape: =MAX($B$1, $B$1+INT(C5/D5)-XLOOKUP(B5, ${K}, ${LT}))`],
            explain: 'Dates are day numbers, so adding the days of stock and taking away the lead time gives a date. INT keeps whole days only, because a part day of stock cannot wait for a delivery, and MAX turns any date in the past into B1, the day the order must go out.',
            chips: ['$B$1', 'INT(C5/D5)', 'B5', K, LT, 'MAX('], fnKey: 'MAX('
          });
        },
        () => {
          /* One warehouse's purchase budget: top every SKU at or below its reorder point up to its maximum, at cost. */
          const whs = r.sample(H.WAREHOUSES, 4);
          const pattern = r.shuffle([0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 3, 3, 3]);
          const rows = H.stock(r, 14).map((x, i) => ({ sku: x.sku, wh: whs[pattern[i]], rop: x.reorder, cost: x.cost }));
          rows.forEach((x) => { x.max = Math.round(x.rop * r.pick([2, 2.5, 3]) / 10) * 10; x.oh = r.chance(0.45) ? Math.max(0, x.rop - r.int(0, 8) * 10) : x.rop + r.int(1, 30) * 10; });
          const of = (w) => rows.filter((x) => x.wh === w);
          /* Planted: in each of the first two warehouses, one SKU exactly at its reorder point, one below, and one above. */
          [0, 1].forEach((k) => { const g = of(whs[k]); g[0].oh = g[0].rop; g[1].oh = Math.max(0, g[1].rop - r.int(2, 8) * 10); g[2].oh = g[2].rop + r.int(2, 20) * 10; });
          const budget = (w) => of(w).reduce((a, x) => a + (x.oh <= x.rop ? (x.max - x.oh) * x.cost : 0), 0);
          const g2 = of(whs[1])[2];
          if (budget(whs[0]) === budget(whs[1]) + (g2.max - g2.rop) * g2.cost) of(whs[1])[1].oh = Math.max(0, of(whs[1])[1].oh - 10);
          const S = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Warehouse', k: 'wh' }, { h: 'On hand', k: 'oh', num: true }, { h: 'Reorder point', k: 'rop', num: true }, { h: 'Max stock', k: 'max', num: true }, { h: 'Cost per carton (Rp)', k: 'cost', num: true }], rows);
          const base = Object.assign({}, S.cells, { A1: 'Warehouse', B1: whs[0], A2: 'Order budget (Rp)' });
          /* The hidden sheet: the second warehouse, after one of its SKUs above the reorder point sells down to it. */
          const g1 = of(whs[1]);
          const alt = H.edit(base, { B1: whs[1], [S.cell('On hand', rows.indexOf(g1[2]))]: g1[2].rop });
          const Wr = S.R('Warehouse'), OH = S.R('On hand'), RP = S.R('Reorder point'), MX = S.R('Max stock'), CO = S.R('Cost per carton (Rp)');
          return H.problem({
            title: 'Reorder budget for one warehouse at cost', file: 'anggaran_pembelian.xlsx', sheet: S, base, alt, answerCell: 'B2',
            statement: `<p>Finance asks each warehouse for this week's purchase budget. For the warehouse in <b>B1</b>, every SKU whose <b>On hand</b> is at or below its <b>Reorder point</b> is topped up to its <b>Max stock</b>, so it orders Max stock minus On hand cartons, paid at <b>Cost per carton</b>. SKUs above their reorder point order nothing.</p><p>In <b>B2</b>, return the total cost in rupiah of the warehouse's orders, in one formula.</p>`,
            solution: `=SUMPRODUCT((${Wr}=B1)*(${OH}<=${RP})*(${MX}-${OH})*${CO})`,
            alts: [`=SUMPRODUCT(--(${Wr}=B1),--(${OH}<=${RP}),${MX}-${OH},${CO})`, `=LET(need,(${Wr}=B1)*(${OH}<=${RP})*(${MX}-${OH}),SUMPRODUCT(need,${CO}))`],
            hints: ['Two tests pick the rows: the warehouse is B1, and On hand is at or below the reorder point. As numbers, each test is 1 or 0.',
              `For one row, the cost is (${MX.split(':')[0]}-${OH.split(':')[0]})*${CO.split(':')[0]}. SUMPRODUCT can multiply the two tests by that for every row and add the results.`,
              `Shape: =SUMPRODUCT((${Wr}=B1)*(${OH}<=${RP})*(${MX}-${OH})*${CO})`],
            explain: 'Each test turns into 1 or 0, so rows from other warehouses and SKUs above their reorder point add nothing. For the rest, SUMPRODUCT multiplies the cartons to order by the cost and adds them up, all without a helper column.',
            chips: [Wr, OH, RP, MX, CO, '(' + Wr + '=B1)', '<='], fnKey: 'SUMPRODUCT('
          });
        }
      ][v]();
    }
  });
});
