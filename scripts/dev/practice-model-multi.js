/*
 * A model for practice problems that work across several workbooks, as real reports do. Not loaded by the app.
 *   node --no-warnings scripts/dev/check-practice.mjs scripts/dev/practice-model-multi.js --show
 * The learner's own file holds the sales lines; the product master and the targets are separate workbooks. The answer
 * looks up in both, then does the math: each line's value at list price, as a share of the distributor's Q4 target.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) factory(require('../../assets/js/kicau/practice/core.js'));
  else factory(root.KicauPractice);
})(typeof globalThis !== 'undefined' ? globalThis : this, function (P) {
  'use strict';
  const H = P.H;
  P.add({
    key: 'model-line-share-of-target', difficulty: 'Hard', level: 9, topics: ['XLOOKUP', 'Several workbooks'], count: 2, multi: true,
    make(r, v) {
      const lines = H.sales(r, 12);
      const S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'Distributor', k: 'distributor' }, { h: 'SKU', k: 'sku' }, { h: 'Cartons', k: 'cartons', num: true }], lines);
      /* Workbook 2: the product master (SKU, product, carton price). */
      const PM = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Product', k: 'name' }, { h: 'Carton price (Rp)', k: 'carton_price', num: true }], H.PRODUCTS);
      /* Workbook 3: Q4 targets per distributor. */
      const dists = H.db().distributors.map((d, i) => ({ name: d.name, target: [180, 150, 120, 140, 100, 90, 110, 80, 75, 160, 95, 130, 70, 105][i] * 1000000 }));
      const TG = H.table([{ h: 'Distributor', k: 'name' }, { h: 'Q4 target (Rp)', k: 'target', num: true }], dists);
      const prod = H.ext('master_produk.xlsx', 'Produk', PM), targ = H.ext('target_2026.xlsx', 'Target', TG);
      /* The hidden test: the price list and two targets change, so typed numbers fail. */
      const pmAlt = H.edit(PM.cells, { C6: (x) => x + 6000, C9: (x) => x + 8000, C14: (x) => x + 4000 });
      const tgAlt = H.edit(TG.cells, { B5: (x) => x + 20000000, B9: (x) => x - 10000000, B12: (x) => x + 15000000 });
      const base = S.cells, alt = H.edit(S.cells, { D6: (x) => x + 9 });
      const solution = '=D5*XLOOKUP(C5,' + prod.A('SKU') + ',' + prod.A('Carton price (Rp)') + ')/XLOOKUP(B5,' + targ.A('Distributor') + ',' + targ.A('Q4 target (Rp)') + ')';
      return H.problem({
        title: ['Each line as a share of its Q4 target, from two files', 'Line value over target, prices and targets elsewhere'][v],
        file: 'penjualan_q4.xlsx', sheetName: 'Penjualan', sheet: S, base, alt, fill: true, fillCol: 'E', ansLabel: 'Share of target', precision: 4,
        books: [{ file: 'master_produk.xlsx', sheets: [{ name: 'Produk', sheet: PM, base: PM.cells, alt: pmAlt }] }, { file: 'target_2026.xlsx', sheets: [{ name: 'Target', sheet: TG, base: TG.cells, alt: tgAlt }] }],
        statement: '<p>Three workbooks are open: your sales lines (<b>penjualan_q4.xlsx</b>), the product master (<b>master_produk.xlsx</b>) and the Q4 targets (<b>target_2026.xlsx</b>).</p><p>In <b>E5</b>, work out what the line is worth at list price (Cartons × the SKU\'s carton price from the product master), divided by the distributor\'s Q4 target from the targets file. Fill it down to <b>E' + S.last + '</b>. Leave it as a decimal.</p>',
        solution,
        alts: ['=D5*VLOOKUP(C5,' + prod.T('SKU', 'Carton price (Rp)') + ',3,FALSE)/VLOOKUP(B5,' + targ.T('Distributor', 'Q4 target (Rp)') + ',2,FALSE)'],
        hints: ['Two lookups, one in each other workbook: the carton price by SKU, and the target by distributor.', 'Tap a workbook tab, then the column: Kicau writes ' + prod.A('SKU') + ' for you.', 'Shape: =D5*XLOOKUP(C5, SKU column, price column)/XLOOKUP(B5, distributor column, target column).'],
        explain: 'The first XLOOKUP brings the carton price from master_produk.xlsx, the second the Q4 target from target_2026.xlsx. Cartons × price ÷ target gives the share. The ranges in other files are pinned with $, as Excel writes them, so filling down keeps them in place.',
        chips: ['D5*', prod.A('SKU'), prod.A('Carton price (Rp)'), targ.A('Distributor'), targ.A('Q4 target (Rp)'), 'C5', 'B5'], fnKey: 'XLOOKUP('
      });
    }
  });
});
