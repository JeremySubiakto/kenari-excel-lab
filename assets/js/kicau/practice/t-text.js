/*
 * Practice, Text (20 Easy, 25 Medium, 10 Hard): LEFT, RIGHT, MID, FIND, SEARCH, LEN, TRIM, UPPER, LOWER, PROPER,
 * SUBSTITUTE, REPLACE, TEXT, joining with & and TEXTJOIN, VALUE. Codes from invoice references, messy distributor
 * names and, at the hard end, splitting at the second delimiter, counting characters and building references.
 * The engine's text functions do not run over a whole range (LEN(A5:A9) gives one value, not five), so every
 * answer here works one cell at a time and is filled down, or uses COUNTIF, SUMIF, FILTER and TEXTJOIN, which do.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) factory(require('./core.js'));
  else factory(root.KicauPractice);
})(typeof globalThis !== 'undefined' ? globalThis : this, function (P) {
  'use strict';
  const H = P.H;

  /* ---------- small helpers for codes and messy text ---------- */
  const DISTS = () => H.db().distributors;
  const cityCode = (city) => city.slice(0, 3).toUpperCase();
  const pickOther = (r, arr, same) => r.pick(arr.filter((x) => !same(x)));
  /* "In E5 ... then fill down to E18": the answer column of a sheet. */
  const AT = (S) => '<b>' + S.next + S.first + '</b>';
  const DOWN = (S) => ', then fill down to <b>' + S.next + S.last + '</b>';
  const ENGLISH = '<p>Kicau uses the English format codes and month names. On Indonesian settings Excel may show Okt for Oct and swap the separators.</p>';
  /* The same words in the wrong capitals, the way people type them. Never already in Proper Case. */
  function messCase(r, s) {
    const k = r.int(0, 3);
    if (k === 0) return s.toLowerCase();
    if (k === 1) return s.toUpperCase();
    if (k === 2) return s.split(' ').map((w) => w.charAt(0).toLowerCase() + w.slice(1).toUpperCase()).join(' ');
    return s.split('').map((c, i) => (i % 2 ? c.toUpperCase() : c.toLowerCase())).join('');
  }
  /* Extra spaces at the ends and between words; always at least one more than the clean text. */
  function messSpace(r, s) {
    const words = s.split(' ');
    const gap = r.int(0, Math.max(0, words.length - 2));
    const twice = r.chance(0.7);
    let out = words.map((w, i) => w + (i < words.length - 1 ? (i === gap && twice ? '  ' : ' ') : '')).join('');
    out = ' '.repeat(r.pick([0, 1, 2])) + out + ' '.repeat(r.pick([0, 1, 2]));
    return out === s ? s + ' ' : out;
  }
  /* A typing slip: two letters near the end swapped (or the last letter dropped). */
  function typo(s) {
    const i = s.length - 3;
    const t = s.slice(0, i) + s[i + 1] + s[i] + s.slice(i + 2);
    return t === s ? s.slice(0, -1) : t;
  }
  const phone = (r, sep) => '08' + r.int(11, 59) + sep + r.int(1000, 9999) + sep + r.int(1000, 9999);
  const SURNAMES = ['Hartono', 'Pratiwi', 'Saputra', 'Lestari', 'Wibowo', 'Kurniawan', 'Anggraini', 'Nugroho', 'Permata', 'Setiawan', 'Rahayu', 'Siregar'];
  /* Sales lines from a few regions only, so the hidden sheet can bring in a region that was not there. */
  function salesIn(r, regions, n) {
    const rows = r.sample(H.db().sales.filter((s) => regions.includes(s.region)), n).map((s) => Object.assign({}, s, { date: H.Q4(s.date) }));
    return rows.sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : 0));
  }

  /* ==================== Easy ==================== */

  P.add({
    key: 'letters-from-codes', difficulty: 'Easy', level: 1, topics: ['LEFT', 'RIGHT', 'MID'], count: 4,
    make(r, v) {
      return [
        () => {
          const rows = H.sales(r, 14);
          const S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'Distributor', k: 'distributor' }, { h: 'SKU', k: 'sku' }, { h: 'Cartons', k: 'cartons', num: true }], rows);
          const other = (i) => pickOther(r, H.PRODUCTS, (p) => p.sku.slice(0, 3) === rows[i].sku.slice(0, 3)).sku;
          const alt = H.edit(S.cells, { [S.cell('SKU', 1)]: other(1), [S.cell('SKU', 8)]: other(8) });
          return H.problem({
            title: 'Product family from the first three letters of a SKU', file: 'penjualan_q4.xlsx', sheet: S, alt, fill: true, ansLabel: 'Family', exactCase: true,
            statement: '<p>Every SKU starts with a 3-letter product family, such as <b>MLK</b> for Mi Lekas or <b>KSJ</b> for Kopi Sekejap. In ' + AT(S) + ', take the first 3 characters of the SKU in <b>C5</b>' + DOWN(S) + '.</p>',
            solution: '=LEFT(C5,3)', alts: ['=MID(C5,1,3)', '=TEXTBEFORE(C5,"-")'],
            hints: ['LEFT takes characters from the start of a text.', 'You need 3 characters from the start of C5.', 'It looks like =LEFT(C5,3).'],
            explain: 'LEFT(C5,3) returns the first three characters of the SKU. C5 has no $ sign, so it moves down a row as you fill and every line gets its own family.',
            chips: ['C5', ',3)'], fnKey: 'LEFT('
          });
        },
        () => {
          const rows = H.stock(r, 12);
          const S = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Product', k: 'product' }, { h: 'Warehouse', k: 'warehouse' }, { h: 'On hand', k: 'onhand', num: true }], rows);
          const swap = (i) => pickOther(r, H.PRODUCTS, (p) => p.sku.slice(4) === rows[i].sku.slice(4));
          const p2 = swap(2), p9 = swap(9);
          const alt = H.edit(S.cells, { [S.cell('SKU', 2)]: p2.sku, [S.cell('Product', 2)]: p2.name, [S.cell('SKU', 9)]: p9.sku, [S.cell('Product', 9)]: p9.name });
          return H.problem({
            title: 'Flavour code from the last three letters of a SKU', file: 'stok_gudang.xlsx', sheet: S, alt, fill: true, ansLabel: 'Flavour', exactCase: true,
            statement: '<p>The last 3 letters of a SKU name the flavour or scent, such as <b>GOR</b> in MLK-GOR (Goreng). In ' + AT(S) + ', take the last 3 characters of the SKU in <b>A5</b>' + DOWN(S) + '.</p>',
            solution: '=RIGHT(A5,3)', alts: ['=MID(A5,5,3)', '=TEXTAFTER(A5,"-")'],
            hints: ['RIGHT takes characters from the end of a text.', 'You need the last 3 characters of A5.', 'It looks like =RIGHT(A5,3).'],
            explain: 'RIGHT(A5,3) reads three characters from the end of the SKU, which is the flavour part after the dash.',
            chips: ['A5', ',3)'], fnKey: 'RIGHT('
          });
        },
        () => {
          const rows = H.invoices(r, 12);
          const S = H.table([{ h: 'Reference', k: 'ref' }, { h: 'Distributor', k: 'distributor' }, { h: 'Date', k: 'date', date: true }, { h: 'Amount (Rp)', k: 'amount', num: true }], rows);
          const renum = (x) => x.replace(/-(\d+)-/, (m, n) => '-' + (+n + 400) + '-');
          const alt = H.edit(S.cells, { [S.cell('Reference', 3)]: renum, [S.cell('Reference', 10)]: renum });
          return H.problem({
            title: 'Invoice number from the middle of each reference', file: 'faktur_q4.xlsx', sheet: S, alt, fill: true, ansLabel: 'Invoice no.', exactCase: true,
            statement: '<p>Each invoice reference is INV, a dash, a 4-digit invoice number, a dash and a 3-letter city code: <b>INV-2104-BEK</b>. In ' + AT(S) + ', take the invoice number: the 4 characters that start at character 5 of the reference in <b>A5</b>' + DOWN(S) + '.</p>',
            solution: '=MID(A5,5,4)', alts: ['=LEFT(RIGHT(A5,8),4)', '=TEXTBEFORE(TEXTAFTER(A5,"-"),"-")', '=VALUE(MID(A5,5,4))'],
            hints: ['MID reads characters from the middle of a text: from a start position, for a number of characters.', 'INV and the dash are characters 1 to 4, so the number starts at character 5 and is 4 characters long.', 'It looks like =MID(A5,5,4).'],
            explain: 'MID(A5,5,4) starts at character 5, right after INV and its dash, and returns 4 characters: the invoice number, without the city code after it.',
            chips: ['A5', ',5,4)'], fnKey: 'MID('
          });
        },
        () => {
          const rows = r.sample(DISTS(), 12);
          const S = H.table([{ h: 'ID', k: 'id' }, { h: 'Distributor', k: 'name' }, { h: 'City', k: 'city' }, { h: 'Region', k: 'region' }], rows);
          const flip = (s) => (s.startsWith('CV') ? 'UD' + s.slice(2) : 'CV' + s.slice(2));
          const alt = H.edit(S.cells, { [S.cell('Distributor', 1)]: flip, [S.cell('Distributor', 7)]: flip });
          return H.problem({
            title: 'Legal form (CV or UD) at the start of a distributor name', file: 'master_distributor.xlsx', sheet: S, alt, fill: true, ansLabel: 'Legal form', exactCase: true,
            statement: '<p>Every distributor name starts with its legal form, <b>CV</b> or <b>UD</b>. In ' + AT(S) + ', take the first 2 characters of the name in <b>B5</b>' + DOWN(S) + '.</p>',
            solution: '=LEFT(B5,2)', alts: ['=MID(B5,1,2)', '=TEXTBEFORE(B5," ")'],
            hints: ['The legal form is at the start of the name.', 'LEFT with 2 characters from B5.', 'It looks like =LEFT(B5,2).'],
            explain: 'LEFT(B5,2) keeps the first two characters of the name, which is the legal form. On the hidden sheet two distributors changed their legal form and the answer follows.',
            chips: ['B5', ',2)'], fnKey: 'LEFT('
          });
        }
      ][v]();
    }
  });

  P.add({
    key: 'change-letter-case', difficulty: 'Easy', level: 1, topics: ['UPPER', 'LOWER', 'PROPER'], count: 4,
    make(r, v) {
      return [
        () => {
          const rows = r.shuffle(H.PRODUCTS);
          const S = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Product', k: 'name' }, { h: 'Category', k: 'category' }, { h: 'Carton price (Rp)', k: 'carton_price', num: true }], rows);
          const alt = H.edit(S.cells, { [S.cell('Product', 2)]: (x) => x + ' Pedas', [S.cell('Product', 8)]: (x) => x + ' Hemat' });
          return H.problem({
            title: 'Product names in capital letters for carton labels', file: 'produk.xlsx', sheet: S, alt, fill: true, ansLabel: 'Label', exactCase: true,
            statement: '<p>The carton labels print product names in capitals. In ' + AT(S) + ', write the product name from <b>B5</b> in capital letters' + DOWN(S) + '. The capitals must match exactly.</p>',
            solution: '=UPPER(B5)',
            hints: ['One function turns every letter into a capital.', 'UPPER takes the text to change: the name in B5.', 'It looks like =UPPER(B5).'],
            explain: 'UPPER(B5) returns the same name with every letter in capitals. Spaces and dashes stay as they are.',
            chips: ['B5'], fnKey: 'UPPER('
          });
        },
        () => {
          const rows = H.reps(r, 10);
          const S = H.table([{ h: 'Rep', k: 'name' }, { h: 'Area', k: 'area' }, { h: 'Target (Rp)', k: 'target', num: true }, { h: 'Visits', k: 'visits', num: true }], rows);
          const unused = H.REPS.filter((n) => !rows.some((x) => x.name === n));
          const alt = H.edit(S.cells, { [S.cell('Rep', 3)]: unused[0], [S.cell('Rep', 8)]: unused[1] });
          return H.problem({
            title: 'Login IDs in small letters from rep names', file: 'tim_sales.xlsx', sheet: S, alt, fill: true, ansLabel: 'Login ID', exactCase: true,
            statement: '<p>The new sales app needs every login ID in small letters. In ' + AT(S) + ', write the rep name from <b>A5</b> in small letters (Rudi becomes <b>rudi</b>)' + DOWN(S) + '.</p>',
            solution: '=LOWER(A5)',
            hints: ['UPPER makes capitals. Its opposite makes small letters.', 'LOWER takes the text to change: the name in A5.', 'It looks like =LOWER(A5).'],
            explain: 'LOWER(A5) turns every capital into a small letter, so Rudi becomes rudi. Two new reps joined on the hidden sheet and get their IDs too.',
            chips: ['A5'], fnKey: 'LOWER('
          });
        },
        () => {
          const rows = H.sales(r, 14).map((s) => Object.assign(s, { typed: messCase(r, s.product) }));
          const S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'Distributor', k: 'distributor' }, { h: 'Product (typed)', k: 'typed' }, { h: 'Cartons', k: 'cartons', num: true }], rows);
          const retype = (i) => messCase(r, pickOther(r, H.PRODUCTS, (p) => p.name === rows[i].product).name);
          const alt = H.edit(S.cells, { [S.cell('Product (typed)', 2)]: retype(2), [S.cell('Product (typed)', 11)]: retype(11) });
          return H.problem({
            title: 'Fix the capitals in hand-typed product names', file: 'form_pesanan.xlsx', sheet: S, alt, fill: true, ansLabel: 'Product', exactCase: true,
            statement: '<p>Product names on the order form were typed in random capitals. In ' + AT(S) + ', give each word of the name in <b>C5</b> a capital first letter and small letters after it (<b>Mi Lekas Goreng</b>)' + DOWN(S) + '.</p>',
            solution: '=PROPER(C5)',
            hints: ['UPPER and LOWER change every letter. A third function capitalises only the first letter of each word.', 'PROPER(C5) gives every word a capital first letter.', 'It looks like =PROPER(C5).'],
            explain: 'PROPER puts a capital on the first letter of each word and makes the rest small, which turns mI LEKAS goreng into Mi Lekas Goreng.',
            chips: ['C5'], fnKey: 'PROPER('
          });
        },
        () => {
          const pool = r.shuffle(DISTS());
          const rows = pool.slice(0, 12).map((d) => Object.assign({}, d, { typed: messCase(r, d.city) }));
          const S = H.table([{ h: 'ID', k: 'id' }, { h: 'Distributor', k: 'name' }, { h: 'City (typed)', k: 'typed' }, { h: 'Region', k: 'region' }], rows);
          const put = (i, d) => ({ [S.cell('ID', i)]: d.id, [S.cell('Distributor', i)]: d.name, [S.cell('City (typed)', i)]: messCase(r, d.city), [S.cell('Region', i)]: d.region });
          const alt = H.edit(S.cells, Object.assign(put(4, pool[12]), put(9, pool[13])));
          return H.problem({
            title: 'Fix the capitals in hand-typed city names', file: 'master_distributor.xlsx', sheet: S, alt, fill: true, ansLabel: 'City', exactCase: true,
            statement: '<p>The city column was typed by several people, in any capitals: <b>bEKASI</b>, <b>SURABAYA</b>, <b>malang</b>. In ' + AT(S) + ', write the city from <b>C5</b> with a capital first letter and small letters after it (<b>Bekasi</b>)' + DOWN(S) + '.</p>',
            solution: '=PROPER(C5)',
            hints: ['You need a capital first letter and small letters after it.', 'PROPER does exactly that for each word.', 'It looks like =PROPER(C5).'],
            explain: 'PROPER(C5) rewrites the city with one capital at the start of each word, however it was typed.',
            chips: ['C5'], fnKey: 'PROPER('
          });
        }
      ][v]();
    }
  });

  P.add({
    key: 'count-characters', difficulty: 'Easy', level: 2, topics: ['LEN', 'IF'], count: 3,
    make(r, v) {
      return [
        () => {
          const pool = r.shuffle(DISTS());
          const rows = pool.slice(0, 10);
          const S = H.table([{ h: 'ID', k: 'id' }, { h: 'Distributor', k: 'name' }, { h: 'City', k: 'city' }, { h: 'Region', k: 'region' }], rows);
          /* The hidden sheet: two rows now hold distributors whose names have a different length. */
          const spare = pool.slice(10);
          const a = spare.find((d) => d.name.length !== rows[2].name.length) || spare[0];
          const b = spare.find((d) => d !== a && d.name.length !== rows[7].name.length) || spare.find((d) => d !== a);
          const alt = H.edit(S.cells, { [S.cell('Distributor', 2)]: a.name, [S.cell('Distributor', 7)]: b.name });
          return H.problem({
            title: 'Number of characters in each distributor name', file: 'master_distributor.xlsx', sheet: S, alt, fill: true, ansLabel: 'Characters',
            statement: '<p>The bank\'s system cuts names that are too long. In ' + AT(S) + ', count the characters in the distributor name in <b>B5</b>, spaces included' + DOWN(S) + '.</p>',
            solution: '=LEN(B5)',
            hints: ['One function counts the characters in a text.', 'LEN counts every character, spaces too.', 'It looks like =LEN(B5).'],
            explain: 'LEN(B5) returns how many characters the name has, counting letters and spaces alike.',
            chips: ['B5'], fnKey: 'LEN('
          });
        },
        () => {
          const rows = r.shuffle(H.PRODUCTS).slice(0, 12);
          const S = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Product', k: 'name' }, { h: 'Category', k: 'category' }], rows);
          const w = r.pick([24, 26, 28]);
          const base = Object.assign({}, S.cells, { A1: 'Label width (characters)', B1: w });
          const alt = H.edit(base, { B1: w + 4, [S.cell('Product', 5)]: (x) => x + ' Pedas' });
          return H.problem({
            title: 'Characters left on a shelf label of fixed width', file: 'produk.xlsx', sheet: S, base, alt, fill: true, ansLabel: 'Characters left',
            statement: '<p>A shelf label holds the number of characters in <b>B1</b>. In ' + AT(S) + ', work out how many characters are left after the product name in <b>B5</b>: B1 minus the length of the name. A negative result means the name does not fit. Fill down to <b>' + S.next + S.last + '</b>, and pin B1 with $ so it stays put.</p>',
            solution: '=$B$1-LEN(B5)', alts: ['=B$1-LEN(B5)', '=-LEN(B5)+$B$1'],
            hints: ['LEN(B5) is the length of the name.', 'Take it away from the width in B1, and write B1 as $B$1 so it does not slide when you fill down.', 'It looks like =$B$1-LEN(B5).'],
            explain: 'LEN counts the characters of each name and the label width in $B$1 stays fixed for every row. On the hidden sheet the label is wider, so every answer moves.',
            chips: ['$B$1', 'LEN(B5)'], fnKey: 'LEN('
          });
        },
        () => {
          const rows = r.sample(DISTS(), 12);
          const S = H.table([{ h: 'ID', k: 'id' }, { h: 'Distributor', k: 'name' }, { h: 'City', k: 'city' }, { h: 'Region', k: 'region' }], rows);
          const lim = r.pick([15, 16]);
          const base = Object.assign({}, S.cells, { A1: 'Name limit (characters)', B1: lim });
          const alt = H.edit(base, { B1: lim + 2 });
          return H.problem({
            title: 'Flag distributor names longer than the system limit', file: 'master_distributor.xlsx', sheet: S, base, alt, fill: true, ansLabel: 'Name check', exactCase: true,
            statement: '<p>The payment system takes names of up to <b>B1</b> characters. In ' + AT(S) + ', write <b>Too long</b> when the name in <b>B5</b> has more than B1 characters, otherwise <b>OK</b>' + DOWN(S) + '.</p>',
            solution: '=IF(LEN(B5)>$B$1,"Too long","OK")', alts: ['=IF(LEN(B5)<=$B$1,"OK","Too long")', '=IF($B$1<LEN(B5),"Too long","OK")'],
            hints: ['LEN(B5) gives the length of the name.', 'Compare it with $B$1 inside IF: more than the limit means Too long.', 'It looks like =IF(LEN(B5)>$B$1,"Too long","OK").'],
            explain: 'IF tests whether LEN(B5) is greater than the limit in $B$1. Exactly at the limit is still OK, because the rule is more than.',
            chips: ['LEN(B5)', '>$B$1', ',"Too long","OK")'], fnKey: 'IF('
          });
        }
      ][v]();
    }
  });

  P.add({
    key: 'remove-extra-spaces', difficulty: 'Easy', level: 2, topics: ['TRIM', 'LEN', 'IF'], count: 3,
    make(r, v) {
      return [
        () => {
          const pool = r.shuffle(DISTS());
          const rows = pool.slice(0, 12).map((d) => Object.assign({}, d, { typed: messSpace(r, d.name) }));
          const S = H.table([{ h: 'ID', k: 'id' }, { h: 'Distributor (typed)', k: 'typed' }, { h: 'City', k: 'city' }], rows);
          const put = (i, d) => ({ [S.cell('ID', i)]: d.id, [S.cell('Distributor (typed)', i)]: messSpace(r, d.name), [S.cell('City', i)]: d.city });
          const alt = H.edit(S.cells, Object.assign(put(3, pool[12]), put(10, pool[13])));
          return H.problem({
            title: 'Remove extra spaces from distributor names', file: 'master_distributor.xlsx', sheet: S, alt, fill: true, ansLabel: 'Clean name', exactCase: true,
            statement: '<p>The names in column B have spaces at the start, at the end and doubled between words. In ' + AT(S) + ', write each name from <b>B5</b> with no spaces at the ends and single spaces between words' + DOWN(S) + '.</p>',
            solution: '=TRIM(B5)',
            hints: ['One function removes extra spaces.', 'TRIM cuts spaces at both ends and squeezes double spaces inside to one.', 'It looks like =TRIM(B5).'],
            explain: 'TRIM removes leading and trailing spaces and turns any run of spaces between words into one, which is what a lookup needs to match.',
            chips: ['B5'], fnKey: 'TRIM('
          });
        },
        () => {
          const rows = r.shuffle(H.PRODUCTS).map((p) => Object.assign({}, p, { typed: r.chance(0.75) ? messSpace(r, p.name) : p.name }));
          const S = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Product (typed)', k: 'typed' }, { h: 'Category', k: 'category' }], rows);
          const alt = H.edit(S.cells, { [S.cell('Product (typed)', 1)]: (x) => x + '  ', [S.cell('Product (typed)', 6)]: (x) => '   ' + x });
          return H.problem({
            title: 'Count the extra spaces in each product name', file: 'produk.xlsx', sheet: S, alt, fill: true, ansLabel: 'Extra spaces',
            statement: '<p>Some product names were typed with extra spaces. In ' + AT(S) + ', count how many characters TRIM would remove from the name in <b>B5</b>: its length minus the length of the trimmed name' + DOWN(S) + '. A clean name gives 0.</p>',
            solution: '=LEN(B5)-LEN(TRIM(B5))', alts: ['=-LEN(TRIM(B5))+LEN(B5)'],
            hints: ['LEN(B5) counts every character, extra spaces included.', 'LEN(TRIM(B5)) counts the characters after the extra spaces are gone.', 'It looks like =LEN(B5)-LEN(TRIM(B5)).'],
            explain: 'The difference between the length before and after TRIM is the number of extra spaces. A name that was typed cleanly gives 0.',
            chips: ['LEN(B5)', 'TRIM(B5)'], fnKey: 'LEN('
          });
        },
        () => {
          const rows = r.sample(DISTS(), 12).map((d) => Object.assign({}, d, { bad: r.chance(0.35) }));
          rows[2].bad = true; rows[5].bad = false;
          rows.forEach((d) => { d.typed = messSpace(r, d.bad ? typo(d.name) : d.name); });
          const S = H.table([{ h: 'ID', k: 'id' }, { h: 'Typed name', k: 'typed' }, { h: 'Master name', k: 'name' }], rows);
          const alt = H.edit(S.cells, { [S.cell('Typed name', 2)]: messSpace(r, rows[2].name), [S.cell('Typed name', 5)]: messSpace(r, typo(rows[5].name)) });
          return H.problem({
            title: 'Check typed names against the master list', file: 'cek_nama.xlsx', sheet: S, alt, fill: true, ansLabel: 'Result', exactCase: true,
            statement: '<p>Column B holds names typed by the sales team, with stray spaces and the odd typing slip. In ' + AT(S) + ', write <b>Match</b> when the name in <b>B5</b>, once its extra spaces are removed, equals the master name in <b>C5</b>, otherwise <b>Check</b>' + DOWN(S) + '.</p>',
            solution: '=IF(TRIM(B5)=C5,"Match","Check")', alts: ['=IF(TRIM(B5)<>C5,"Check","Match")', '=IF(C5=TRIM(B5),"Match","Check")'],
            hints: ['Compare the cleaned typed name with the master name.', 'TRIM(B5)=C5 is TRUE when they agree. Put that test inside IF.', 'It looks like =IF(TRIM(B5)=C5,"Match","Check").'],
            explain: 'TRIM removes the stray spaces, so only a real difference in letters gives Check. Comparing B5 with C5 directly would flag every name with an extra space.',
            chips: ['TRIM(B5)', '=C5', ',"Match","Check")'], fnKey: 'IF('
          });
        }
      ][v]();
    }
  });

  P.add({
    key: 'join-text-pieces', difficulty: 'Easy', level: 3, topics: ['&', 'CONCAT', 'TEXTJOIN'], count: 3,
    make(r, v) {
      return [
        () => {
          const rows = r.shuffle(H.PRODUCTS);
          const S = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Product', k: 'name' }, { h: 'Category', k: 'category' }], rows);
          const alt = H.edit(S.cells, { [S.cell('Product', 3)]: (x) => x + ' Pedas', [S.cell('SKU', 9)]: (x) => x.slice(0, 4) + 'NEW' });
          return H.problem({
            title: 'Join each SKU and product name with a dash', file: 'produk.xlsx', sheet: S, alt, fill: true, ansLabel: 'Catalogue line', exactCase: true,
            statement: '<p>The price list shows each product as <b>SKU - Product</b>, for example <b>MLK-GOR - Mi Lekas Goreng</b>: a space, a dash and a space between them. In ' + AT(S) + ', join the SKU in <b>A5</b> and the name in <b>B5</b> that way' + DOWN(S) + '.</p>',
            solution: '=A5&" - "&B5', alts: ['=CONCAT(A5," - ",B5)', '=TEXTJOIN(" - ",TRUE,A5,B5)'],
            hints: ['The & sign joins pieces of text.', 'The separator is a text of its own: " - " in double quotes.', 'It looks like =A5&" - "&B5.'],
            explain: 'The & sign glues the SKU, the separator text " - " and the name into one text, in that order.',
            chips: ['A5', '&" - "&', 'B5'], fnKey: '&'
          });
        },
        () => {
          const pool = r.shuffle(DISTS());
          const rows = pool.slice(0, 12);
          const S = H.table([{ h: 'ID', k: 'id' }, { h: 'Distributor', k: 'name' }, { h: 'City', k: 'city' }, { h: 'Region', k: 'region' }], rows);
          const put = (i, d) => ({ [S.cell('ID', i)]: d.id, [S.cell('Distributor', i)]: d.name, [S.cell('City', i)]: d.city, [S.cell('Region', i)]: d.region });
          const alt = H.edit(S.cells, Object.assign(put(0, pool[12]), put(6, pool[13])));
          return H.problem({
            title: 'Distributor name with its city in brackets', file: 'master_distributor.xlsx', sheet: S, alt, fill: true, ansLabel: 'Label', exactCase: true,
            statement: '<p>The statement header shows each distributor as <b>CV Nuri Makmur (Bekasi)</b>: the name, a space, then the city in round brackets. In ' + AT(S) + ', build that from <b>B5</b> and <b>C5</b>' + DOWN(S) + '.</p>',
            solution: '=B5&" ("&C5&")"', alts: ['=CONCAT(B5," (",C5,")")'],
            hints: ['Three joins: the name, " (" with its space, the city, then ")".', 'Each piece of fixed text goes in double quotes, joined with &.', 'It looks like =B5&" ("&C5&")".'],
            explain: 'The & signs join the name, the text " (" (a space and an open bracket), the city and the closing bracket.',
            chips: ['B5', '&" ("&', 'C5', '&")"'], fnKey: '&'
          });
        },
        () => {
          const rows = H.reps(r, 10);
          const S = H.table([{ h: 'Rep', k: 'name' }, { h: 'Area', k: 'area' }, { h: 'Visits', k: 'visits', num: true }], rows);
          const alt = H.edit(S.cells, { [S.cell('Visits', 2)]: (x) => x + 6, [S.cell('Visits', 7)]: (x) => x - 4 });
          return H.problem({
            title: 'Rep name and visit count in one label', file: 'tim_sales.xlsx', sheet: S, alt, fill: true, ansLabel: 'Label', exactCase: true,
            statement: '<p>The weekly chat update lists each rep as <b>Rudi: 42 visits</b>. In ' + AT(S) + ', join the name in <b>A5</b>, a colon and a space, the visits in <b>C5</b>, and the word <b>visits</b> after a space' + DOWN(S) + '.</p>',
            solution: '=A5&": "&C5&" visits"', alts: ['=CONCAT(A5,": ",C5," visits")'],
            hints: ['A number joins to text with & like any other piece.', 'The pieces are A5, ": ", C5 and " visits".', 'It looks like =A5&": "&C5&" visits".'],
            explain: 'The & sign turns the number of visits into text as it joins, so the label follows the visit count when it changes.',
            chips: ['A5', '&": "&', 'C5', '&" visits"'], fnKey: '&'
          });
        }
      ][v]();
    }
  });

  P.add({
    key: 'swap-characters', difficulty: 'Easy', level: 3, topics: ['SUBSTITUTE'], count: 3,
    make(r, v) {
      return [
        () => {
          const rows = H.invoices(r, 12);
          const S = H.table([{ h: 'Reference', k: 'ref' }, { h: 'Distributor', k: 'distributor' }, { h: 'Date', k: 'date', date: true }, { h: 'Amount (Rp)', k: 'amount', num: true }], rows);
          const renum = (x) => x.replace(/-(\d+)-/, (m, n) => '-' + (+n + 300) + '-');
          const alt = H.edit(S.cells, { [S.cell('Reference', 2)]: renum, [S.cell('Reference', 8)]: renum });
          return H.problem({
            title: 'Turn the dashes in invoice references into slashes', file: 'faktur_q4.xlsx', sheet: S, alt, fill: true, ansLabel: 'Bank reference', exactCase: true,
            statement: '<p>The bank\'s upload file wants references with slashes: <b>INV/2104/BEK</b> instead of INV-2104-BEK. In ' + AT(S) + ', replace every dash in the reference in <b>A5</b> with a slash' + DOWN(S) + '.</p>',
            solution: '=SUBSTITUTE(A5,"-","/")',
            hints: ['SUBSTITUTE swaps one piece of text for another, everywhere it appears.', 'SUBSTITUTE(text, old, new): the old text is "-", the new one "/".', 'It looks like =SUBSTITUTE(A5,"-","/").'],
            explain: 'SUBSTITUTE(A5,"-","/") replaces every dash with a slash. Without a fourth argument it changes all of them.',
            chips: ['A5', ',"-"', ',"/")'], fnKey: 'SUBSTITUTE('
          });
        },
        () => {
          const rows = r.sample(DISTS(), 12).map((d) => Object.assign({}, d, { phone: (r.chance(0.3) ? ' ' : '') + phone(r, ' ') }));
          const S = H.table([{ h: 'Distributor', k: 'name' }, { h: 'City', k: 'city' }, { h: 'Phone (typed)', k: 'phone' }], rows);
          const alt = H.edit(S.cells, { [S.cell('Phone (typed)', 1)]: phone(r, ' '), [S.cell('Phone (typed)', 9)]: phone(r, '  ') });
          return H.problem({
            title: 'Remove the spaces inside contact phone numbers', file: 'kontak_distributor.xlsx', sheet: S, alt, fill: true, ansLabel: 'Phone', exactCase: true,
            statement: '<p>The messaging tool rejects phone numbers with spaces. In ' + AT(S) + ', remove every space from the number in <b>C5</b>, so <b>0812 3456 7890</b> becomes <b>081234567890</b>' + DOWN(S) + '. Keep the result as text so the leading 0 stays.</p>',
            solution: '=SUBSTITUTE(C5," ","")',
            hints: ['Replace each space with nothing.', 'Nothing is written as two double quotes: "".', 'It looks like =SUBSTITUTE(C5," ","").'],
            explain: 'SUBSTITUTE swaps every space for an empty text, which removes them all, the ones at the start included. The result is text, so the leading 0 survives.',
            chips: ['C5', '," "', ',"")'], fnKey: 'SUBSTITUTE('
          });
        },
        () => {
          const rows = H.stock(r, 12);
          const S = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Product', k: 'product' }, { h: 'Warehouse', k: 'warehouse' }, { h: 'On hand', k: 'onhand', num: true }], rows);
          const move = (i) => pickOther(r, H.WAREHOUSES, (w) => w === rows[i].warehouse);
          const alt = H.edit(S.cells, { [S.cell('Warehouse', 0)]: move(0), [S.cell('Warehouse', 7)]: move(7) });
          return H.problem({
            title: 'Shorten Gudang to WH in warehouse names', file: 'stok_gudang.xlsx', sheet: S, alt, fill: true, ansLabel: 'Short name', exactCase: true,
            statement: '<p>The handheld scanner shows short warehouse names. In ' + AT(S) + ', replace the word <b>Gudang</b> with <b>WH</b> in the warehouse name in <b>C5</b>, so Gudang Bekasi becomes <b>WH Bekasi</b>' + DOWN(S) + '.</p>',
            solution: '=SUBSTITUTE(C5,"Gudang","WH")', alts: ['=SUBSTITUTE(C5,"Gudang ","WH ")'],
            hints: ['SUBSTITUTE can swap a whole word, not only one character.', 'The old text is "Gudang" and the new text is "WH".', 'It looks like =SUBSTITUTE(C5,"Gudang","WH").'],
            explain: 'SUBSTITUTE finds the word Gudang inside the name and puts WH in its place. Its search is case-sensitive, so the capital G matters.',
            chips: ['C5', ',"Gudang"', ',"WH")'], fnKey: 'SUBSTITUTE('
          });
        }
      ][v]();
    }
  });

  /* ==================== Medium ==================== */

  P.add({
    key: 'split-at-first-delimiter', difficulty: 'Medium', level: 4, topics: ['LEFT', 'MID', 'FIND'], count: 3,
    make(r, v) {
      return [
        () => {
          const rows = r.sample(H.db().outlets, 14);
          const S = H.table([{ h: 'Outlet ID', k: 'outlet_id' }, { h: 'Outlet', k: 'outlet' }, { h: 'City', k: 'city' }], rows);
          const retype = (i) => { const first = rows[i].outlet.split(' ')[0]; return pickOther(r, ['Warung', 'Toko', 'Kios', 'Grosir'], (w) => w === first) + rows[i].outlet.slice(first.length); };
          const alt = H.edit(S.cells, { [S.cell('Outlet', 2)]: retype(2), [S.cell('Outlet', 10)]: retype(10) });
          return H.problem({
            title: 'Outlet type from the first word of an outlet name', file: 'outlet.xlsx', sheet: S, alt, fill: true, ansLabel: 'Type', exactCase: true,
            statement: '<p>An outlet\'s name starts with its type: <b>Warung</b>, <b>Toko</b>, <b>Kios</b> and so on, all of different lengths. In ' + AT(S) + ', take everything before the first space of the name in <b>B5</b>' + DOWN(S) + '.</p>',
            solution: '=LEFT(B5,FIND(" ",B5)-1)', alts: ['=TEXTBEFORE(B5," ")', '=MID(B5,1,FIND(" ",B5)-1)'],
            hints: ['The type is not a fixed number of characters, so find where the first space is.', 'FIND(" ",B5) gives the position of the first space. The type is one character shorter.', 'It looks like =LEFT(B5,FIND(" ",B5)-1).'],
            explain: 'FIND returns the position of the first space; LEFT takes one character less than that, which is the whole first word whatever its length.',
            chips: ['B5', 'FIND(" ",B5)', '-1)'], fnKey: 'FIND('
          });
        },
        () => {
          const pool = r.shuffle(DISTS());
          const rows = pool.slice(0, 12).map((d) => Object.assign({}, d, { label: d.name + ' - ' + d.city }));
          const S = H.table([{ h: 'Customer label', k: 'label' }, { h: 'Region', k: 'region' }, { h: 'Terms (days)', k: 'terms_days', num: true }], rows);
          const put = (i, d) => ({ [S.cell('Customer label', i)]: d.name + ' - ' + d.city, [S.cell('Region', i)]: d.region, [S.cell('Terms (days)', i)]: d.terms_days });
          const alt = H.edit(S.cells, Object.assign(put(1, pool[12]), put(8, pool[13])));
          return H.problem({
            title: 'City after the dash in a distributor label', file: 'label_pelanggan.xlsx', sheet: S, alt, fill: true, ansLabel: 'City', exactCase: true,
            statement: '<p>Each label is a name, then <b>space dash space</b>, then the city: <b>CV Nuri Makmur - Bekasi</b>. In ' + AT(S) + ', take everything after the <code> - </code> separator in <b>A5</b>' + DOWN(S) + '.</p>',
            solution: '=MID(A5,FIND(" - ",A5)+3,99)', alts: ['=RIGHT(A5,LEN(A5)-FIND(" - ",A5)-2)', '=TEXTAFTER(A5," - ")', '=MID(A5,FIND("-",A5)+2,99)'],
            hints: ['Find where the separator starts, then read from after it.', 'FIND(" - ",A5) is where the separator starts. The city begins 3 characters later.', 'It looks like =MID(A5,FIND(" - ",A5)+3,99).'],
            explain: 'FIND locates the separator, +3 steps over its three characters, and MID with a large length such as 99 reads to the end of the text.',
            chips: ['A5', 'FIND(" - ",A5)', '+3', ',99)'], fnKey: 'MID('
          });
        },
        () => {
          const firsts = r.sample(H.REPS, 10), lasts = r.shuffle(SURNAMES);
          const rows = firsts.map((n, i) => ({ full: n + ' ' + lasts[i], area: r.pick(H.REGIONS), visits: r.int(18, 64) }));
          const S = H.table([{ h: 'Full name', k: 'full' }, { h: 'Area', k: 'area' }, { h: 'Visits', k: 'visits', num: true }], rows);
          const alt = H.edit(S.cells, { [S.cell('Full name', 2)]: firsts[2] + ' ' + lasts[10], [S.cell('Full name', 6)]: firsts[6] + ' ' + lasts[11] });
          return H.problem({
            title: 'Surname after the first space of a full name', file: 'tim_sales.xlsx', sheet: S, alt, fill: true, ansLabel: 'Surname', exactCase: true,
            statement: '<p>The HR file wants the surname on its own. In ' + AT(S) + ', take everything after the first space of the full name in <b>A5</b> (Wulan Pratiwi gives <b>Pratiwi</b>)' + DOWN(S) + '.</p>',
            solution: '=MID(A5,FIND(" ",A5)+1,99)', alts: ['=RIGHT(A5,LEN(A5)-FIND(" ",A5))', '=TEXTAFTER(A5," ")'],
            hints: ['First names have different lengths, so find the space first.', 'The surname starts one character after FIND(" ",A5).', 'It looks like =MID(A5,FIND(" ",A5)+1,99).'],
            explain: 'FIND gives the position of the space; MID starts one character after it and reads to the end of the name.',
            chips: ['A5', 'FIND(" ",A5)', '+1', ',99)'], fnKey: 'MID('
          });
        }
      ][v]();
    }
  });

  P.add({
    key: 'find-words-in-text', difficulty: 'Medium', level: 4, topics: ['SEARCH', 'ISNUMBER', 'COUNTIF', 'SUMIF'], count: 3,
    make(r, v) {
      const RET = ['Return of 2 damaged cartons', 'Partial RETURN, wrong flavour sent', 'Outlet returned 1 carton of sambal', 'Expired stock sent back as a return', 'Asked to return 3 cartons of soap'];
      const OK = ['Delivered on time', 'Paid in full by transfer', 'New outlet opened near the market', 'Asked for promo posters', 'Late delivery, no complaint', 'Owner wants a bigger order next week', 'Display rack moved to the front'];
      return [
        () => {
          const rows = H.sales(r, 14).map((s) => Object.assign(s, { note: r.chance(0.4) ? r.pick(RET) : r.pick(OK) }));
          rows[1].note = RET[0]; rows[4].note = OK[0];
          const S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'Distributor', k: 'distributor' }, { h: 'Visit note', k: 'note' }], rows);
          const alt = H.edit(S.cells, { [S.cell('Visit note', 1)]: r.pick(OK), [S.cell('Visit note', 4)]: r.pick(RET) });
          return H.problem({
            title: 'Flag visit notes that mention a return', file: 'catatan_kunjungan.xlsx', sheet: S, alt, fill: true, ansLabel: 'Type', exactCase: true,
            statement: '<p>In ' + AT(S) + ', write <b>Return</b> when the note in <b>C5</b> contains the letters <b>return</b> anywhere, in any capitals (Returned and RETURN count too), otherwise <b>Sale</b>' + DOWN(S) + '.</p>',
            solution: '=IF(ISNUMBER(SEARCH("return",C5)),"Return","Sale")', alts: ['=IF(ISERROR(SEARCH("return",C5)),"Sale","Return")', '=IF(COUNTIF(C5,"*return*"),"Return","Sale")', '=IFERROR(IF(SEARCH("return",C5),"Return"),"Sale")'],
            hints: ['SEARCH finds a word inside a text and ignores capitals. FIND does not ignore them.', 'SEARCH returns a position when it finds the word and an error when it does not, so ISNUMBER turns that into TRUE or FALSE.', 'It looks like =IF(ISNUMBER(SEARCH("return",C5)),"Return","Sale").'],
            explain: 'SEARCH gives a number only when return appears somewhere in the note, in any capitals. ISNUMBER turns that into TRUE or FALSE for IF.',
            chips: ['SEARCH("return",C5)', 'ISNUMBER(', ',"Return","Sale")'], fnKey: 'SEARCH('
          });
        },
        () => {
          const rows = H.sales(r, 14);
          const S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'Distributor', k: 'distributor' }, { h: 'Product', k: 'product' }, { h: 'Cartons', k: 'cartons', num: true }], rows);
          const n = (w) => rows.filter((x) => x.product.toLowerCase().includes(w.toLowerCase())).length;
          const words = r.shuffle(['Kopi', 'Sambal', 'Sabun', 'Lekas', 'Wangi', 'Nyala']).filter((w) => n(w) > 0);
          const w1 = words[0], w2 = words.find((w) => n(w) !== n(w1)) || words[1];
          const base = Object.assign({}, S.cells, { A1: 'Word to find', B1: w1, A2: 'Lines' });
          const alt = H.edit(base, { B1: w2 });
          return H.problem({
            title: 'Count order lines whose product contains a word', file: 'penjualan_q4.xlsx', sheet: S, base, alt, answerCell: 'B2',
            statement: '<p>In <b>B2</b>, count the order lines whose product name in <b>' + S.R('Product') + '</b> contains the word in <b>B1</b> anywhere in the name, in any capitals. Use a condition with the * wildcard, and point at B1 rather than typing the word, because it will change.</p>',
            solution: '=COUNTIF(' + S.R('Product') + ',"*"&B1&"*")', alts: ['=COUNTIFS(' + S.R('Product') + ',"*"&B1&"*")'],
            hints: ['In a COUNTIF condition, * stands for any number of characters.', 'Build the condition by joining: "*"&B1&"*" means B1 with anything before or after it.', 'It looks like =COUNTIF(' + S.R('Product') + ',"*"&B1&"*").'],
            explain: 'The stars on both sides let the word sit anywhere in the name, and joining them to B1 keeps the count right when the word changes.',
            chips: [S.R('Product'), '"*"&B1&"*"'], fnKey: 'COUNTIF('
          });
        },
        () => {
          const rows = H.sales(r, 14);
          const S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'Product', k: 'product' }, { h: 'Cartons', k: 'cartons', num: true }, { h: 'Net (Rp)', k: 'net', num: true }], rows);
          const n = (w) => rows.filter((x) => x.product.toLowerCase().includes(w.toLowerCase())).length;
          const words = r.shuffle(['Sekejap', 'Nyala', 'Rimba', 'Lekas', 'Susu', 'Bawang', 'Wangi']).filter((w) => n(w) > 0);
          const base = Object.assign({}, S.cells, { A1: 'Word to find', B1: words[0], A2: 'Net (Rp)' });
          const alt = H.edit(base, { B1: words[1] });
          return H.problem({
            title: 'Net sales of products whose name contains a word', file: 'penjualan_q4.xlsx', sheet: S, base, alt, answerCell: 'B2',
            statement: '<p>In <b>B2</b>, add up <b>Net (Rp)</b> (' + S.R('Net (Rp)') + ') for the lines whose product name in <b>' + S.R('Product') + '</b> contains the word in <b>B1</b> anywhere, in any capitals. Use a condition with the * wildcard so the word can sit in the middle of a name.</p>',
            solution: '=SUMIF(' + S.R('Product') + ',"*"&B1&"*",' + S.R('Net (Rp)') + ')', alts: ['=SUMIFS(' + S.R('Net (Rp)') + ',' + S.R('Product') + ',"*"&B1&"*")'],
            hints: ['SUMIF adds the rows that meet a condition.', 'The condition "*"&B1&"*" matches names with the word anywhere in them.', 'It looks like =SUMIF(' + S.R('Product') + ',"*"&B1&"*",' + S.R('Net (Rp)') + ').'],
            explain: 'SUMIF checks each product name against the pattern built from B1 and adds the net of the rows that match. The stars make the word match in the middle of a name.',
            chips: [S.R('Product'), '"*"&B1&"*"', S.R('Net (Rp)')], fnKey: 'SUMIF('
          });
        }
      ][v]();
    }
  });

  P.add({
    key: 'format-with-text', difficulty: 'Medium', level: 5, topics: ['TEXT', '&'], count: 4,
    make(r, v) {
      return [
        () => {
          const rows = r.sample(H.db().outlets, 12).map((o, i) => Object.assign({}, o, { no: [r.int(1, 9), r.int(10, 99), r.int(100, 240)][i % 3] }));
          const S = H.table([{ h: 'Outlet no.', k: 'no', num: true }, { h: 'Outlet', k: 'outlet' }, { h: 'City', k: 'city' }], rows);
          const alt = H.edit(S.cells, { [S.cell('Outlet no.', 0)]: (x) => x + 30, [S.cell('Outlet no.', 7)]: (x) => x + 5 });
          return H.problem({
            title: 'Outlet codes padded to three digits with TEXT', file: 'outlet.xlsx', sheet: S, alt, fill: true, ansLabel: 'Code', exactCase: true,
            statement: '<p>Outlet codes are <b>W</b> followed by the outlet number in 3 digits, with leading zeros: 7 becomes <b>W007</b>, 45 becomes <b>W045</b>. In ' + AT(S) + ', build the code from the number in <b>A5</b>' + DOWN(S) + '.</p>',
            solution: '="W"&TEXT(A5,"000")', alts: ['=CONCAT("W",TEXT(A5,"000"))', '="W"&RIGHT("00"&A5,3)'],
            hints: ['TEXT writes a number in a format you choose, and the format "000" always shows 3 digits.', 'Join the letter W in front with &.', 'It looks like ="W"&TEXT(A5,"000").'],
            explain: 'The format "000" pads the number with zeros to three digits. Joining "W" in front gives codes that sort and match correctly.',
            chips: ['"W"&', 'TEXT(A5,', '"000")'], fnKey: 'TEXT('
          });
        },
        () => {
          const rows = H.sales(r, 14);
          const S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'Distributor', k: 'distributor' }, { h: 'Net (Rp)', k: 'net', num: true }], rows);
          const alt = H.edit(S.cells, { [S.cell('Date', 1)]: (x) => x + 35, [S.cell('Date', 12)]: (x) => x - 40 });
          return H.problem({
            title: 'Month and year label from each order date', file: 'penjualan_q4.xlsx', sheet: S, alt, fill: true, ansLabel: 'Month', exactCase: true,
            statement: '<p>The pivot needs a month label as text, like <b>Oct 2026</b>: the short month name, a space, the 4-digit year. In ' + AT(S) + ', make it from the date in <b>A5</b>' + DOWN(S) + '.</p>' + ENGLISH,
            solution: '=TEXT(A5,"mmm yyyy")', alts: ['=TEXT(A5,"mmm")&" "&TEXT(A5,"yyyy")', '=TEXT(A5,"mmm")&" "&YEAR(A5)'],
            hints: ['A date is a number underneath. TEXT shows it in any date format you name.', 'In format codes, mmm is the short month name and yyyy the 4-digit year.', 'It looks like =TEXT(A5,"mmm yyyy").'],
            explain: 'TEXT with the code "mmm yyyy" writes the date as its short month and year, so every date in October gives the same label.',
            chips: ['A5', ',"mmm yyyy")'], fnKey: 'TEXT('
          });
        },
        () => {
          const rows = H.sales(r, 14);
          const S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'Distributor', k: 'distributor' }, { h: 'Cartons', k: 'cartons', num: true }], rows);
          const alt = H.edit(S.cells, { [S.cell('Date', 3)]: (x) => x + 9, [S.cell('Date', 9)]: (x) => x + 31 });
          return H.problem({
            title: 'Sales order number from the date as yyyymmdd', file: 'penjualan_q4.xlsx', sheet: S, alt, fill: true, ansLabel: 'Order no.', exactCase: true,
            statement: '<p>Sales order numbers are <b>SO-</b> and the date written as year, month and day with no gaps: 5 October 2026 becomes <b>SO-20261005</b>. In ' + AT(S) + ', build it from the date in <b>A5</b>' + DOWN(S) + '.</p>' + ENGLISH,
            solution: '="SO-"&TEXT(A5,"yyyymmdd")', alts: ['=CONCAT("SO-",TEXT(A5,"yyyymmdd"))', '="SO-"&YEAR(A5)&TEXT(MONTH(A5),"00")&TEXT(DAY(A5),"00")'],
            hints: ['Joining a date directly gives its serial number, such as 46300. TEXT is needed.', 'The code yyyymmdd writes the year, then a 2-digit month, then a 2-digit day.', 'It looks like ="SO-"&TEXT(A5,"yyyymmdd").'],
            explain: 'TEXT turns the date into the digits 20261005, with leading zeros on single-digit months and days. Joining it to "SO-" gives the order number.',
            chips: ['"SO-"&', 'TEXT(A5,', '"yyyymmdd")'], fnKey: 'TEXT('
          });
        },
        () => {
          const rows = H.reps(r, 10);
          /* Two reps off a round figure, so the visible sheet already needs rounding; no result ends in exactly .5. */
          rows[1].actual += 1700000; rows[5].actual -= 2300000;
          rows.forEach((x) => { if (Math.abs((x.actual / x.target * 100) % 1 - 0.5) < 1e-9) x.actual += 100000; });
          const S = H.table([{ h: 'Rep', k: 'name' }, { h: 'Area', k: 'area' }, { h: 'Target (Rp)', k: 'target', num: true }, { h: 'Actual (Rp)', k: 'actual', num: true }], rows);
          const alt = H.edit(S.cells, { [S.cell('Actual (Rp)', 2)]: (x) => x + 9000000, [S.cell('Actual (Rp)', 6)]: (x) => x - 13000000 });
          return H.problem({
            title: 'Rep achievement written as a whole percentage', file: 'tim_sales.xlsx', sheet: S, alt, fill: true, ansLabel: 'Message', exactCase: true,
            statement: '<p>For the team chat, write each rep\'s achievement as <b>Rudi reached 95%</b>: the name in <b>A5</b>, the word <b>reached</b> with a space on each side, then <b>Actual</b> (D5) divided by <b>Target</b> (C5) as a percentage rounded to the nearest whole number, with a % sign (94.6% shows as <b>95%</b>). Type it in ' + AT(S) + DOWN(S) + '.</p>',
            solution: '=A5&" reached "&TEXT(D5/C5,"0%")', alts: ['=CONCAT(A5," reached ",TEXT(D5/C5,"0%"))', '=A5&" reached "&ROUND(D5/C5*100,0)&"%"'],
            hints: ['D5/C5 is the achievement as a decimal, like 0.95.', 'TEXT with the format "0%" writes 0.95 as 95%.', 'It looks like =A5&" reached "&TEXT(D5/C5,"0%").'],
            explain: 'TEXT(D5/C5,"0%") turns the ratio into a rounded whole percentage. Joining the number without TEXT would give a long decimal such as 0.9466.',
            chips: ['A5', '&" reached "&', 'TEXT(D5/C5,', '"0%")'], fnKey: 'TEXT('
          });
        }
      ][v]();
    }
  });

  P.add({
    key: 'numbers-inside-codes', difficulty: 'Medium', level: 5, topics: ['MID', 'FIND', 'VALUE'], count: 3,
    make(r, v) {
      return [
        () => {
          const rows = H.sales(r, 14).map((s) => Object.assign(s, { code: s.sku + '/' + s.packs }));
          const S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'Distributor', k: 'distributor' }, { h: 'Pack code', k: 'code' }, { h: 'Cartons', k: 'cartons', num: true }], rows);
          const p = pickOther(r, H.PRODUCTS, (x) => x.packs_per_carton === rows[5].packs);
          const alt = H.edit(S.cells, { [S.cell('Pack code', 5)]: p.sku + '/' + p.packs_per_carton, [S.cell('Cartons', 11)]: (x) => x + 8 });
          return H.problem({
            title: 'Total packs using the pack size inside a code', file: 'penjualan_q4.xlsx', sheet: S, alt, fill: true, ansLabel: 'Packs',
            statement: '<p>Each pack code ends with the packs per carton after a slash: <b>MLK-GOR/40</b> means 40 packs, <b>KSJ-SUS/120</b> means 120. In ' + AT(S) + ', multiply the <b>Cartons</b> in D5 by the number after the slash in <b>C5</b>' + DOWN(S) + '.</p>',
            solution: '=D5*VALUE(MID(C5,FIND("/",C5)+1,9))', alts: ['=D5*MID(C5,FIND("/",C5)+1,9)', '=D5*TEXTAFTER(C5,"/")', '=VALUE(TEXTAFTER(C5,"/"))*D5', '=D5*RIGHT(C5,LEN(C5)-FIND("/",C5))'],
            hints: ['The number has 2 or 3 digits, so RIGHT(C5,2) is not safe. Start after the slash instead.', 'MID(C5,FIND("/",C5)+1,9) is the text after the slash. VALUE turns it into a number.', 'It looks like =D5*VALUE(MID(C5,FIND("/",C5)+1,9)).'],
            explain: 'FIND locates the slash and MID reads everything after it, whether it is 24 or 120. VALUE makes that text a number you can multiply by the cartons.',
            chips: ['D5*', 'VALUE(', 'MID(C5,FIND("/",C5)+1,9)'], fnKey: 'VALUE('
          });
        },
        () => {
          const rows = H.invoices(r, 12).map((x) => Object.assign(x, { termsText: 'NET ' + r.pick([7, 14, 30, 45, 60]) }));
          const S = H.table([{ h: 'Reference', k: 'ref' }, { h: 'Distributor', k: 'distributor' }, { h: 'Date', k: 'date', date: true }, { h: 'Terms', k: 'termsText' }], rows);
          const retime = (i) => 'NET ' + pickOther(r, [7, 14, 30, 45, 60], (d) => 'NET ' + d === rows[i].termsText);
          const alt = H.edit(S.cells, { [S.cell('Terms', 2)]: retime(2), [S.cell('Terms', 9)]: retime(9) });
          return H.problem({
            title: 'Due date from payment terms written as text', file: 'faktur_q4.xlsx', sheet: S, alt, fill: true, ansLabel: 'Due date', ansDate: true,
            statement: '<p>Terms are typed as text: <b>NET 30</b> means payment is due 30 days after the invoice date. The number can be 1 or 2 digits. In ' + AT(S) + ', add the number of days in <b>D5</b> to the date in <b>C5</b>' + DOWN(S) + '.</p>',
            solution: '=C5+VALUE(MID(D5,5,9))', alts: ['=C5+MID(D5,5,9)', '=C5+VALUE(SUBSTITUTE(D5,"NET ",""))', '=C5+TEXTAFTER(D5," ")', '=C5+VALUE(RIGHT(D5,LEN(D5)-4))'],
            hints: ['NET and the space are always the first 4 characters, so the days start at character 5.', 'MID(D5,5,9) is the number as text. VALUE makes it a number you can add to a date.', 'It looks like =C5+VALUE(MID(D5,5,9)).'],
            explain: 'MID reads the days after "NET " and VALUE turns them into a number. A date is a count of days, so adding the number moves it forward that many days.',
            chips: ['C5+', 'VALUE(', 'MID(D5,5,9)'], fnKey: 'VALUE('
          });
        },
        () => {
          const rows = H.sales(r, 14).map((s) => Object.assign(s, { promo: 'DISC' + r.pick([2, 5, 10, 15]) }));
          const S = H.table([{ h: 'Distributor', k: 'distributor' }, { h: 'Product', k: 'product' }, { h: 'Promo code', k: 'promo' }, { h: 'Gross (Rp)', k: 'gross', num: true }], rows);
          const recode = (i) => 'DISC' + pickOther(r, [2, 5, 10, 15], (d) => 'DISC' + d === rows[i].promo);
          const alt = H.edit(S.cells, { [S.cell('Promo code', 0)]: recode(0), [S.cell('Promo code', 8)]: recode(8) });
          return H.problem({
            title: 'Net after the discount written inside a promo code', file: 'promo_q4.xlsx', sheet: S, alt, fill: true, ansLabel: 'Net (Rp)',
            statement: '<p>The promo code holds the discount percentage after the letters DISC: <b>DISC5</b> is 5% off, <b>DISC15</b> is 15% off. In ' + AT(S) + ', work out the net: <b>Gross</b> in D5 times (1 minus the percentage / 100)' + DOWN(S) + '.</p>',
            solution: '=D5*(1-VALUE(MID(C5,5,9))/100)', alts: ['=D5*(1-MID(C5,5,9)/100)', '=D5-D5*VALUE(SUBSTITUTE(C5,"DISC",""))/100', '=D5*(100-MID(C5,5,9))/100'],
            hints: ['DISC is always 4 letters, so the percentage starts at character 5.', 'VALUE(MID(C5,5,9)) is the percentage as a number. Divide it by 100.', 'It looks like =D5*(1-VALUE(MID(C5,5,9))/100).'],
            explain: 'MID reads the digits after DISC and VALUE turns them into a number. Dividing by 100 makes it a rate, so the gross times (1 minus the rate) is the net.',
            chips: ['D5*', '(1-', 'VALUE(MID(C5,5,9))', '/100)'], fnKey: 'VALUE('
          });
        }
      ][v]();
    }
  });

  P.add({
    key: 'replace-by-position', difficulty: 'Medium', level: 5, topics: ['REPLACE', 'SUBSTITUTE'], count: 3,
    make(r, v) {
      return [
        () => {
          const rows = r.sample(DISTS(), 12).map((d) => Object.assign({}, d, { ph: phone(r, '-') }));
          const S = H.table([{ h: 'Distributor', k: 'name' }, { h: 'City', k: 'city' }, { h: 'Phone', k: 'ph' }], rows);
          const alt = H.edit(S.cells, { [S.cell('Phone', 3)]: phone(r, '-'), [S.cell('Phone', 9)]: phone(r, '-') });
          return H.problem({
            title: 'Hide the middle digits of each phone number', file: 'kontak_distributor.xlsx', sheet: S, alt, fill: true, ansLabel: 'Masked phone', exactCase: true,
            statement: '<p>The shared report must hide part of each phone number. Every number looks like <b>0812-3456-7890</b>. In ' + AT(S) + ', replace the 4 digits in the middle (characters 6 to 9 of <b>C5</b>) with <b>****</b>, giving 0812-****-7890' + DOWN(S) + '.</p>',
            solution: '=REPLACE(C5,6,4,"****")', alts: ['=LEFT(C5,5)&"****"&RIGHT(C5,5)', '=LEFT(C5,5)&"****"&MID(C5,10,5)'],
            hints: ['REPLACE swaps characters by position, whatever they are.', 'REPLACE(text, start, how many, new text): start at 6, replace 4.', 'It looks like =REPLACE(C5,6,4,"****").'],
            explain: 'REPLACE removes 4 characters starting at position 6 and puts **** in their place. SUBSTITUTE would not work here, because it looks for known text, not a position.',
            chips: ['C5', ',6,4', ',"****")'], fnKey: 'REPLACE('
          });
        },
        () => {
          const rows = H.invoices(r, 12).map((x, i) => Object.assign(x, { po: 'PO-' + r.pick([2024, 2025]) + '-' + String(r.int(1, 9999)).padStart(4, '0') }));
          const S = H.table([{ h: 'PO number', k: 'po' }, { h: 'Distributor', k: 'distributor' }, { h: 'Amount (Rp)', k: 'amount', num: true }], rows);
          const yr = r.pick([2026, 2027]);
          const base = Object.assign({}, S.cells, { A1: 'New year', B1: yr });
          const alt = H.edit(base, { B1: yr + 1, [S.cell('PO number', 4)]: (x) => x.slice(0, 8) + String(r.int(1, 9999)).padStart(4, '0') });
          return H.problem({
            title: 'Change the year inside purchase order numbers', file: 'po_list.xlsx', sheet: S, base, alt, fill: true, ansLabel: 'New PO number', exactCase: true,
            statement: '<p>PO numbers carried over to the new year must show the year in <b>B1</b>. Every PO number looks like <b>PO-2025-0412</b>, with the year in characters 4 to 7. In ' + AT(S) + ', replace those 4 characters of <b>A5</b> with the year in B1' + DOWN(S) + '.</p>',
            solution: '=REPLACE(A5,4,4,$B$1)', alts: ['=LEFT(A5,3)&$B$1&MID(A5,8,99)', '=REPLACE(A5,4,4,B$1)', '=LEFT(A5,3)&$B$1&RIGHT(A5,5)'],
            hints: ['The old years differ, so swap by position instead of searching for a year.', 'REPLACE(A5,4,4, new text) replaces characters 4 to 7. The new text is the year in $B$1.', 'It looks like =REPLACE(A5,4,4,$B$1).'],
            explain: 'REPLACE swaps the four characters from position 4, whatever year they hold, for the year in $B$1. The $ keeps B1 fixed as you fill down.',
            chips: ['A5', ',4,4', ',$B$1)'], fnKey: 'REPLACE('
          });
        },
        () => {
          const rows = H.invoices(r, 12);
          const S = H.table([{ h: 'Reference', k: 'ref' }, { h: 'Distributor', k: 'distributor' }, { h: 'Amount (Rp)', k: 'amount', num: true }], rows);
          const renum = (x) => x.replace(/-(\d+)-/, (m, n) => '-' + (+n + 517) + '-');
          const alt = H.edit(S.cells, { [S.cell('Reference', 1)]: renum, [S.cell('Reference', 7)]: renum });
          return H.problem({
            title: 'Bank-style invoice references with a slash before the city', file: 'faktur_q4.xlsx', sheet: S, alt, fill: true, ansLabel: 'Bank reference', exactCase: true,
            statement: '<p>The bank\'s statement writes references with a slash before the city code: <b>INV-2104/BEK</b>. In ' + AT(S) + ', replace only the <b>second</b> dash of the reference in <b>A5</b> with a slash' + DOWN(S) + '.</p>',
            solution: '=SUBSTITUTE(A5,"-","/",2)', alts: ['=REPLACE(A5,9,1,"/")', '=LEFT(A5,8)&"/"&RIGHT(A5,3)'],
            hints: ['SUBSTITUTE has an optional fourth argument: which occurrence to change.', 'With 2 as the fourth argument, only the second dash changes.', 'It looks like =SUBSTITUTE(A5,"-","/",2).'],
            explain: 'The fourth argument of SUBSTITUTE picks one occurrence. Here 2 changes the second dash and leaves the first one alone.',
            chips: ['A5', ',"-"', ',"/"', ',2)'], fnKey: 'SUBSTITUTE('
          });
        }
      ][v]();
    }
  });

  P.add({
    key: 'clean-messy-names', difficulty: 'Medium', level: 6, topics: ['TRIM', 'PROPER', 'UPPER', 'SUBSTITUTE'], count: 3,
    make(r, v) {
      return [
        () => {
          const rows = H.sales(r, 14).map((s) => Object.assign(s, { typed: messSpace(r, messCase(r, s.product)) }));
          const S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'Product (typed)', k: 'typed' }, { h: 'Cartons', k: 'cartons', num: true }], rows);
          const retype = (i) => messSpace(r, messCase(r, pickOther(r, H.PRODUCTS, (p) => p.name === rows[i].product).name));
          const alt = H.edit(S.cells, { [S.cell('Product (typed)', 3)]: retype(3), [S.cell('Product (typed)', 10)]: retype(10) });
          return H.problem({
            title: 'Clean product names: extra spaces and capitals', file: 'form_pesanan.xlsx', sheet: S, alt, fill: true, ansLabel: 'Product', exactCase: true,
            statement: '<p>The typed product names have stray spaces and random capitals. In ' + AT(S) + ', remove the extra spaces from <b>B5</b> and give every word a capital first letter and small letters after it (<b>Kopi Sekejap Susu</b>)' + DOWN(S) + '.</p>',
            solution: '=PROPER(TRIM(B5))', alts: ['=TRIM(PROPER(B5))'],
            hints: ['Two jobs: spaces and capitals. One function for each.', 'TRIM fixes the spaces and PROPER the capitals. One can sit inside the other.', 'It looks like =PROPER(TRIM(B5)).'],
            explain: 'TRIM removes the extra spaces and PROPER rewrites the capitals. Nesting them does both jobs in one cell, in either order.',
            chips: ['TRIM(B5)', 'PROPER('], fnKey: 'PROPER('
          });
        },
        () => {
          const firsts = r.sample(H.REPS, 10), lasts = r.shuffle(SURNAMES);
          const rows = firsts.map((n, i) => ({ login: n.toLowerCase() + '_' + lasts[i].toLowerCase(), area: r.pick(H.REGIONS), visits: r.int(18, 64) }));
          const S = H.table([{ h: 'Login ID', k: 'login' }, { h: 'Area', k: 'area' }, { h: 'Visits', k: 'visits', num: true }], rows);
          const alt = H.edit(S.cells, { [S.cell('Login ID', 0)]: firsts[0].toLowerCase() + '_' + lasts[10].toLowerCase(), [S.cell('Login ID', 5)]: firsts[5].toLowerCase() + '_' + lasts[11].toLowerCase() });
          return H.problem({
            title: 'Full names from underscore login IDs', file: 'tim_sales.xlsx', sheet: S, alt, fill: true, ansLabel: 'Full name', exactCase: true,
            statement: '<p>The app exports only login IDs such as <b>wulan_pratiwi</b>. In ' + AT(S) + ', turn the login in <b>A5</b> into a name: the underscore becomes a space and each word gets a capital first letter (<b>Wulan Pratiwi</b>)' + DOWN(S) + '.</p>',
            solution: '=PROPER(SUBSTITUTE(A5,"_"," "))', alts: ['=SUBSTITUTE(PROPER(A5),"_"," ")'],
            hints: ['First swap the underscore for a space, then fix the capitals.', 'SUBSTITUTE(A5,"_"," ") gives wulan pratiwi. PROPER does the capitals.', 'It looks like =PROPER(SUBSTITUTE(A5,"_"," ")).'],
            explain: 'SUBSTITUTE turns the underscore into a space, so the login becomes two words, and PROPER gives each word its capital.',
            chips: ['A5', 'SUBSTITUTE(', ',"_"," ")', 'PROPER('], fnKey: 'SUBSTITUTE('
          });
        },
        () => {
          const pool = r.shuffle(DISTS());
          const rows = pool.slice(0, 12).map((d) => Object.assign({}, d, { typed: messSpace(r, messCase(r, d.name)) }));
          const S = H.table([{ h: 'ID', k: 'id' }, { h: 'Distributor (typed)', k: 'typed' }, { h: 'City', k: 'city' }], rows);
          const put = (i, d) => ({ [S.cell('ID', i)]: d.id, [S.cell('Distributor (typed)', i)]: messSpace(r, messCase(r, d.name)), [S.cell('City', i)]: d.city });
          const alt = H.edit(S.cells, Object.assign(put(2, pool[12]), put(9, pool[13])));
          return H.problem({
            title: 'Clean distributor names and keep CV or UD in capitals', file: 'master_distributor.xlsx', sheet: S, alt, fill: true, ansLabel: 'Clean name', exactCase: true,
            statement: '<p>The names in column B have stray spaces and random capitals. PROPER alone would write <b>Cv</b> and <b>Ud</b>. In ' + AT(S) + ', write each name from <b>B5</b> with single spaces, the 2-letter legal form in capitals and the rest in Proper Case: <b>CV Nuri Makmur</b>' + DOWN(S) + '.</p>',
            solution: '=UPPER(LEFT(TRIM(B5),2))&PROPER(MID(TRIM(B5),3,99))', alts: ['=UPPER(LEFT(TRIM(B5),2))&MID(PROPER(TRIM(B5)),3,99)', '=SUBSTITUTE(SUBSTITUTE(PROPER(TRIM(B5)),"Cv ","CV "),"Ud ","UD ")'],
            hints: ['Clean the spaces first with TRIM(B5); the legal form is then the first 2 characters.', 'Treat the two parts apart: UPPER for LEFT(TRIM(B5),2), PROPER for MID(TRIM(B5),3,99). Join them with &.', 'It looks like =UPPER(LEFT(TRIM(B5),2))&PROPER(MID(TRIM(B5),3,99)).'],
            explain: 'After TRIM the legal form is always the first two characters, so UPPER fixes that part and PROPER the rest. Joining the two gives CV Nuri Makmur instead of Cv Nuri Makmur.',
            chips: ['TRIM(B5)', 'UPPER(LEFT(', ',2))', '&PROPER(MID(', ',3,99))'], fnKey: 'PROPER('
          });
        }
      ][v]();
    }
  });

  P.add({
    key: 'text-between-markers', difficulty: 'Medium', level: 6, topics: ['MID', 'FIND', 'TRIM'], count: 3,
    make(r, v) {
      return [
        () => {
          const pool = r.shuffle(DISTS());
          const lab = (d) => d.name + ' (' + d.city + ') INV-' + r.int(2100, 2999);
          const rows = pool.slice(0, 12).map((d) => ({ label: lab(d), amount: r.int(20, 160) * 100000 }));
          const S = H.table([{ h: 'Customer label', k: 'label' }, { h: 'Amount (Rp)', k: 'amount', num: true }], rows);
          const alt = H.edit(S.cells, { [S.cell('Customer label', 3)]: lab(pool[12]), [S.cell('Customer label', 8)]: lab(pool[13]) });
          return H.problem({
            title: 'City inside the brackets of a customer label', file: 'label_pelanggan.xlsx', sheet: S, alt, fill: true, ansLabel: 'City', exactCase: true,
            statement: '<p>Each label holds the city in round brackets: <b>CV Nuri Makmur (Bekasi) INV-2104</b>. Names and cities have different lengths. In ' + AT(S) + ', take the text between <b>(</b> and <b>)</b> in <b>A5</b>' + DOWN(S) + '.</p>',
            solution: '=MID(A5,FIND("(",A5)+1,FIND(")",A5)-FIND("(",A5)-1)', alts: ['=TEXTBEFORE(TEXTAFTER(A5,"("),")")', '=TEXTAFTER(TEXTBEFORE(A5,")"),"(")'],
            hints: ['MID needs a start and a length. Both come from FIND.', 'The city starts at FIND("(",A5)+1. Its length is the gap between the two brackets, less one.', 'It looks like =MID(A5,FIND("(",A5)+1,FIND(")",A5)-FIND("(",A5)-1).'],
            explain: 'The start is one character after the open bracket, and the length is the distance between the brackets minus one, so MID returns exactly the city.',
            chips: ['A5', 'FIND("(",A5)', 'FIND(")",A5)', '+1', '-1)'], fnKey: 'MID('
          });
        },
        () => {
          const pool = r.shuffle(DISTS());
          /* Tellers type a space or two inside the brackets now and then; padded forces one in. */
          const pad = () => ' '.repeat(r.pick([0, 0, 1, 2]));
          const memo = (d, padded) => r.pick(['Payment', 'Transfer from', 'Deposit by']) + ' [' + (padded ? ' ' : pad()) + d.name + pad() + '] ' + r.pick(['inv ' + r.int(2100, 2999), 'Oct', 'Nov', 'partial', 'final']);
          const rows = H.invoices(r, 12).map((x, i) => Object.assign(x, { memo: memo(pool[i], i === 1 || i === 4) }));
          const S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'Bank memo', k: 'memo' }, { h: 'Amount (Rp)', k: 'amount', num: true }], rows);
          const alt = H.edit(S.cells, { [S.cell('Bank memo', 0)]: memo(pool[12], true), [S.cell('Bank memo', 6)]: memo(pool[13], false) });
          return H.problem({
            title: 'Payer name between square brackets, extra spaces trimmed', file: 'mutasi_bank.xlsx', sheet: S, alt, fill: true, ansLabel: 'Distributor', exactCase: true,
            statement: '<p>Every bank memo carries the payer between square brackets, and tellers sometimes leave spaces inside them: <b>Transfer from [ UD Jalak Sentosa ] inv 2104</b>. The words around the brackets change. In ' + AT(S) + ', take the text between <b>[</b> and <b>]</b> in <b>B5</b>, without spaces at its start or end, so it gives <b>UD Jalak Sentosa</b>' + DOWN(S) + '.</p>',
            solution: '=TRIM(MID(B5,FIND("[",B5)+1,FIND("]",B5)-FIND("[",B5)-1))', alts: ['=TRIM(TEXTBEFORE(TEXTAFTER(B5,"["),"]"))', '=TRIM(TEXTAFTER(TEXTBEFORE(B5,"]"),"["))'],
            hints: ['Find both brackets first, then cut out what lies between them.', 'MID(B5,FIND("[",B5)+1,FIND("]",B5)-FIND("[",B5)-1) is everything between the brackets, spaces included. TRIM removes the spaces at its ends.', 'It looks like =TRIM(MID(B5,FIND("[",B5)+1,FIND("]",B5)-FIND("[",B5)-1)).'],
            explain: 'Two FINDs fix where the brackets are, so MID cuts out what lies between them whatever words come before or after. TRIM then drops the spaces next to the brackets, so the name is clean enough for a lookup.',
            chips: ['B5', 'FIND("[",B5)', 'FIND("]",B5)', '+1', '-1)', 'TRIM('], fnKey: 'MID('
          });
        },
        () => {
          const note = (region) => r.pick(['Region', 'Area', 'Territory']) + ': ' + region + '; Rep: ' + r.pick(H.REPS);
          const rows = H.sales(r, 12).map((s) => Object.assign(s, { note: note(s.region) }));
          const S = H.table([{ h: 'Visit tag', k: 'note' }, { h: 'Distributor', k: 'distributor' }, { h: 'Cartons', k: 'cartons', num: true }], rows);
          /* The hidden sheet: two visits go to a distributor in another region, and the tag follows. */
          const moveTo = (i) => {
            const d = pickOther(r, DISTS(), (x) => x.region === rows[i].region);
            return { [S.cell('Visit tag', i)]: S.cells[S.cell('Visit tag', i)].replace(/: [^;]+;/, ': ' + d.region + ';'), [S.cell('Distributor', i)]: d.name };
          };
          const alt = H.edit(S.cells, Object.assign(moveTo(2), moveTo(9)));
          return H.problem({
            title: 'Region between the colon and the semicolon of a note', file: 'catatan_kunjungan.xlsx', sheet: S, alt, fill: true, ansLabel: 'Region', exactCase: true,
            statement: '<p>Each visit tag reads like <b>Area: Jawa Barat; Rep: Wulan</b>. The first word changes (Region, Area, Territory). In ' + AT(S) + ', take the region: the text after the first <b>colon and space</b> up to the <b>semicolon</b> in <b>A5</b>' + DOWN(S) + '.</p>',
            solution: '=MID(A5,FIND(":",A5)+2,FIND(";",A5)-FIND(":",A5)-2)', alts: ['=TEXTBEFORE(TEXTAFTER(A5,": "),";")', '=TRIM(MID(A5,FIND(":",A5)+1,FIND(";",A5)-FIND(":",A5)-1))'],
            hints: ['FIND returns the first colon, so the one after Rep does not get in the way.', 'The region starts 2 characters after the colon (a colon and a space) and stops before the semicolon.', 'It looks like =MID(A5,FIND(":",A5)+2,FIND(";",A5)-FIND(":",A5)-2).'],
            explain: 'FIND(":",A5) finds the first colon, +2 skips it and its space, and the gap to the semicolon minus 2 is the length of the region.',
            chips: ['A5', 'FIND(":",A5)', 'FIND(";",A5)', '+2', '-2)'], fnKey: 'MID('
          });
        }
      ][v]();
    }
  });

  P.add({
    key: 'join-a-list-in-one-cell', difficulty: 'Medium', level: 7, topics: ['TEXTJOIN', 'FILTER', 'UNIQUE', 'SORT'], count: 3,
    make(r, v) {
      return [
        () => {
          const rows = r.shuffle(DISTS());
          const S = H.table([{ h: 'ID', k: 'id' }, { h: 'Distributor', k: 'name' }, { h: 'City', k: 'city' }, { h: 'Region', k: 'region' }], rows);
          const regs = r.sample(H.REGIONS, 2);
          const base = Object.assign({}, S.cells, { A1: 'Region', B1: regs[0], A2: 'Distributors' });
          const alt = H.edit(base, { B1: regs[1] });
          return H.problem({
            title: 'List the distributors of one region in one cell', file: 'master_distributor.xlsx', sheet: S, base, alt, answerCell: 'B2', exactCase: true,
            statement: '<p>In <b>B2</b>, list every distributor in <b>' + S.R('Distributor') + '</b> whose region (' + S.R('Region') + ') is the one in <b>B1</b>, in sheet order, joined by a comma and a space: <b>CV Kutilang Abadi, UD Murai Niaga</b>.</p>',
            solution: '=TEXTJOIN(", ",TRUE,FILTER(' + S.R('Distributor') + ',' + S.R('Region') + '=B1))', alts: ['=TEXTJOIN(", ",TRUE,IF(' + S.R('Region') + '=B1,' + S.R('Distributor') + ',""))', '=ARRAYTOTEXT(FILTER(' + S.R('Distributor') + ',' + S.R('Region') + '=B1))'],
            hints: ['FILTER returns the names whose region equals B1.', 'TEXTJOIN(", ",TRUE, list) joins a list into one text with the separator between items.', 'It looks like =TEXTJOIN(", ",TRUE,FILTER(' + S.R('Distributor') + ',' + S.R('Region') + '=B1)).'],
            explain: 'FILTER keeps the names from the chosen region and TEXTJOIN glues them with ", ". TRUE tells TEXTJOIN to skip empty items.',
            chips: [S.R('Distributor'), S.R('Region') + '=B1', 'FILTER(', 'TEXTJOIN(", ",TRUE,'], fnKey: 'TEXTJOIN('
          });
        },
        () => {
          const rows = H.stock(r, 12);
          rows[3].onhand = Math.max(0, rows[3].reorder - 40); rows[6].onhand = rows[6].reorder + 120;
          const S = H.table([{ h: 'SKU', k: 'sku' }, { h: 'Product', k: 'product' }, { h: 'Warehouse', k: 'warehouse' }, { h: 'On hand', k: 'onhand', num: true }, { h: 'Reorder point', k: 'reorder', num: true }], rows);
          /* The hidden sheet: row 7 runs low, and row 1 flips the other way. Row 4 stays below, so the list is never empty. */
          const flip = rows[0].onhand < rows[0].reorder ? rows[0].reorder + 60 : Math.max(0, rows[0].reorder - 50);
          const base = Object.assign({}, S.cells, { A1: 'Reorder now' });
          const alt = H.edit(base, { [S.cell('On hand', 6)]: rows[6].reorder - 30, [S.cell('On hand', 0)]: flip });
          return H.problem({
            title: 'SKUs below their reorder point, joined with slashes', file: 'stok_gudang.xlsx', sheet: S, base, alt, answerCell: 'B1', exactCase: true,
            statement: '<p>In <b>B1</b>, list the SKUs (' + S.R('SKU') + ') whose <b>On hand</b> (' + S.R('On hand') + ') is less than their own <b>Reorder point</b> (' + S.R('Reorder point') + '), in sheet order, joined by a space, a slash and a space: <b>MLK-GOR / SNY-IJO</b>.</p>',
            solution: '=TEXTJOIN(" / ",TRUE,FILTER(' + S.R('SKU') + ',' + S.R('On hand') + '<' + S.R('Reorder point') + '))', alts: ['=TEXTJOIN(" / ",TRUE,IF(' + S.R('On hand') + '<' + S.R('Reorder point') + ',' + S.R('SKU') + ',""))'],
            hints: ['The test compares two columns row by row: ' + S.R('On hand') + '<' + S.R('Reorder point') + '.', 'FILTER the SKUs with that test, then join them with TEXTJOIN and the separator " / ".', 'It looks like =TEXTJOIN(" / ",TRUE,FILTER(' + S.R('SKU') + ',' + S.R('On hand') + '<' + S.R('Reorder point') + ')).'],
            explain: 'The comparison gives TRUE for each row below its own reorder point; FILTER keeps those SKUs and TEXTJOIN writes them as one line for the purchasing chat.',
            chips: [S.R('SKU'), S.R('On hand'), '<', S.R('Reorder point'), 'TEXTJOIN(" / ",TRUE,'], fnKey: 'TEXTJOIN('
          });
        },
        () => {
          const regs = r.sample(H.REGIONS, 6);
          const rows = salesIn(r, regs.slice(0, 4), 14);
          const S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'Distributor', k: 'distributor' }, { h: 'Region', k: 'region' }, { h: 'Net (Rp)', k: 'net', num: true }], rows);
          const d4 = DISTS().find((d) => d.region === regs[4]), d5 = DISTS().find((d) => d.region === regs[5]);
          const base = Object.assign({}, S.cells, { A1: 'Regions' });
          const alt = H.edit(base, { [S.cell('Distributor', 2)]: d4.name, [S.cell('Region', 2)]: d4.region, [S.cell('Distributor', 9)]: d5.name, [S.cell('Region', 9)]: d5.region });
          return H.problem({
            title: 'Every region that ordered, once each and sorted', file: 'penjualan_q4.xlsx', sheet: S, base, alt, answerCell: 'B1', exactCase: true,
            statement: '<p>In <b>B1</b>, list the regions in <b>' + S.R('Region') + '</b>, each one once, sorted A to Z, joined by a comma and a space: <b>Bali Nusra, Jawa Barat, Sumatera</b>.</p>',
            solution: '=TEXTJOIN(", ",TRUE,SORT(UNIQUE(' + S.R('Region') + ')))', alts: ['=TEXTJOIN(", ",TRUE,UNIQUE(SORT(' + S.R('Region') + ')))', '=ARRAYTOTEXT(SORT(UNIQUE(' + S.R('Region') + ')))'],
            hints: ['UNIQUE removes the repeats, SORT puts them A to Z.', 'TEXTJOIN then joins the sorted list with ", ".', 'It looks like =TEXTJOIN(", ",TRUE,SORT(UNIQUE(' + S.R('Region') + '))).'],
            explain: 'UNIQUE keeps each region once, SORT orders them, and TEXTJOIN writes the list in one cell. On the hidden sheet two lines come from new regions and the list grows.',
            chips: [S.R('Region'), 'UNIQUE(', 'SORT(', 'TEXTJOIN(", ",TRUE,'], fnKey: 'TEXTJOIN('
          });
        }
      ][v]();
    }
  });

  /* ==================== Hard ==================== */

  P.add({
    key: 'split-at-second-delimiter', difficulty: 'Hard', level: 8, topics: ['MID', 'FIND', 'VALUE'], count: 3,
    make(r, v) {
      return [
        () => {
          const ref = (city) => r.pick(['INV', 'INV', 'INV', 'CN', 'DN']) + '-' + r.pick([r.int(10, 99), r.int(100, 999), r.int(1000, 9999), r.int(10000, 99999)]) + '-' + cityCode(city);
          const rows = H.invoices(r, 12).map((x) => Object.assign(x, { doc: ref(x.city) }));
          const S = H.table([{ h: 'Document', k: 'doc' }, { h: 'Distributor', k: 'distributor' }, { h: 'Amount (Rp)', k: 'amount', num: true }], rows);
          const alt = H.edit(S.cells, { [S.cell('Document', 1)]: ref(rows[1].city), [S.cell('Document', 8)]: ref(rows[8].city) });
          return H.problem({
            title: 'Document number between the first and second dash', file: 'dokumen_q4.xlsx', sheet: S, alt, fill: true, ansLabel: 'Number',
            statement: '<p>Documents are invoices (INV), credit notes (CN) or debit notes (DN): <b>INV-2104-BEK</b>, <b>CN-87-SUR</b>, <b>DN-10432-MED</b>. The prefix and the number change length. In ' + AT(S) + ', take the number between the first and second dash of <b>A5</b> as a number' + DOWN(S) + '.</p>',
            solution: '=VALUE(MID(A5,FIND("-",A5)+1,FIND("-",A5,FIND("-",A5)+1)-FIND("-",A5)-1))', alts: ['=--TEXTBEFORE(TEXTAFTER(A5,"-"),"-")', '=VALUE(TEXTBEFORE(TEXTAFTER(A5,"-"),"-"))', '=VALUE(MID(A5,FIND("-",A5)+1,FIND("#",SUBSTITUTE(A5,"-","#",2))-FIND("-",A5)-1))'],
            hints: ['FIND has a third argument: where to start looking. Starting one past the first dash finds the second dash.', 'The second dash is FIND("-",A5,FIND("-",A5)+1). The number runs from the first dash +1 up to the second dash -1.', 'It looks like =VALUE(MID(A5,FIND("-",A5)+1,FIND("-",A5,FIND("-",A5)+1)-FIND("-",A5)-1)).'],
            explain: 'The nested FIND starts searching after the first dash, so it returns the second one. The gap between the two dashes is the length of the number, and VALUE turns the text into a number.',
            chips: ['A5', 'FIND("-",A5)', 'FIND("-",A5,FIND("-",A5)+1)', 'VALUE('], fnKey: 'FIND('
          });
        },
        () => {
          const rows = H.sales(r, 14);
          const S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'Product', k: 'product' }, { h: 'Cartons', k: 'cartons', num: true }], rows);
          const other = (i) => pickOther(r, H.PRODUCTS, (p) => p.name === rows[i].product).name;
          const alt = H.edit(S.cells, { [S.cell('Product', 4)]: other(4), [S.cell('Product', 11)]: other(11) });
          return H.problem({
            title: 'Flavour words after the second space of a product name', file: 'penjualan_q4.xlsx', sheet: S, alt, fill: true, ansLabel: 'Variant', exactCase: true,
            statement: '<p>Every product name is two brand words, then the variant: <b>Mi Lekas Goreng</b> gives <b>Goreng</b>, <b>Sabun Rimba Wangi Serai</b> gives <b>Wangi Serai</b>. In ' + AT(S) + ', take everything after the <b>second</b> space of the name in <b>B5</b>' + DOWN(S) + '.</p>',
            solution: '=MID(B5,FIND(" ",B5,FIND(" ",B5)+1)+1,99)', alts: ['=TEXTAFTER(B5," ",2)', '=RIGHT(B5,LEN(B5)-FIND(" ",B5,FIND(" ",B5)+1))', '=MID(B5,FIND("#",SUBSTITUTE(B5," ","#",2))+1,99)'],
            hints: ['The brand words have different lengths, so find the second space.', 'FIND(" ",B5,FIND(" ",B5)+1) starts looking after the first space and returns the second.', 'It looks like =MID(B5,FIND(" ",B5,FIND(" ",B5)+1)+1,99).'],
            explain: 'The inner FIND finds the first space, the outer FIND starts one character later and finds the second, and MID reads everything after it, one word or two.',
            chips: ['B5', 'FIND(" ",B5)', 'FIND(" ",B5,FIND(" ",B5)+1)', '+1,99)'], fnKey: 'FIND('
          });
        },
        () => {
          const memo = (paid, name) => r.pick(['TRF', 'SETORAN', 'RTGS']) + ' ' + paid + ' ' + name.toUpperCase();
          const rows = H.invoices(r, 12).map((x) => Object.assign(x, { memo: memo(x.amount + r.pick([0, 0, 0, -6500, -2500, -100000, 50000]), x.distributor) }));
          const S = H.table([{ h: 'Bank memo', k: 'memo' }, { h: 'Invoice (Rp)', k: 'amount', num: true }], rows);
          const alt = H.edit(S.cells, { [S.cell('Bank memo', 2)]: memo(rows[2].amount - 250000, rows[2].distributor), [S.cell('Bank memo', 9)]: memo(rows[9].amount + 15000, rows[9].distributor) });
          return H.problem({
            title: 'Gap between the amount in a bank memo and the invoice', file: 'mutasi_bank.xlsx', sheet: S, alt, fill: true, ansLabel: 'Difference (Rp)',
            statement: '<p>Each bank memo is a code, a space, the amount paid, a space and the payer: <b>SETORAN 12493500 CV NURI MAKMUR</b>. The code changes length. In ' + AT(S) + ', take the amount between the first and second space of <b>A5</b> and subtract the invoice in <b>B5</b> (negative means short paid)' + DOWN(S) + '.</p>',
            solution: '=VALUE(MID(A5,FIND(" ",A5)+1,FIND(" ",A5,FIND(" ",A5)+1)-FIND(" ",A5)-1))-B5', alts: ['=MID(A5,FIND(" ",A5)+1,FIND(" ",A5,FIND(" ",A5)+1)-FIND(" ",A5)-1)-B5', '=VALUE(TEXTBEFORE(TEXTAFTER(A5," ")," "))-B5', '=-B5+TEXTBEFORE(TEXTAFTER(A5," ")," ")'],
            hints: ['The amount sits between the first and the second space.', 'The second space is FIND(" ",A5,FIND(" ",A5)+1). MID from the first space +1, for the gap between the spaces minus 1.', 'It looks like =VALUE(MID(A5,FIND(" ",A5)+1,FIND(" ",A5,FIND(" ",A5)+1)-FIND(" ",A5)-1))-B5.'],
            explain: 'Two FINDs locate the spaces on each side of the amount, MID cuts it out and VALUE turns it into a number, so subtracting the invoice shows any short or over payment.',
            chips: ['A5', 'FIND(" ",A5)', 'FIND(" ",A5,FIND(" ",A5)+1)', 'VALUE(', '-B5'], fnKey: 'FIND('
          });
        }
      ][v]();
    }
  });

  P.add({
    key: 'count-a-character', difficulty: 'Hard', level: 9, topics: ['LEN', 'SUBSTITUTE', 'TRIM', 'UPPER'], count: 3,
    make(r, v) {
      const NOTES = ['Outlet asked for more Kopi Sekejap', 'Display rack moved to the front', 'Owner wants a promo next month', 'Stock of sambal almost empty', 'New fridge installed', 'Paid two old invoices', 'Asked for the price list', 'Competitor gave a bigger discount', 'Shop closed for renovation', 'Order doubled for the holiday'];
      const messNote = (s) => (r.chance(0.5) ? ' ' : '') + s.split(' ').map((w) => w + ' '.repeat(r.pick([1, 1, 2, 3]))).join('').replace(/ +$/, r.pick(['', ' ', '  ']));
      return [
        () => {
          const list = () => r.sample(H.PRODUCTS, r.int(1, 5)).map((p) => p.sku).join(', ');
          const rows = H.invoices(r, 12).map((x) => Object.assign(x, { list: list() }));
          const S = H.table([{ h: 'Order', k: 'ref' }, { h: 'Distributor', k: 'distributor' }, { h: 'SKUs ordered', k: 'list' }], rows);
          /* The hidden sheet adds one SKU to two orders, always one the order did not have, so no list repeats a SKU. */
          const more = (i) => { const x = rows[i].list; return x + ', ' + pickOther(r, H.PRODUCTS, (p) => x.includes(p.sku)).sku; };
          const alt = H.edit(S.cells, { [S.cell('SKUs ordered', 1)]: more(1), [S.cell('SKUs ordered', 7)]: more(7) });
          return H.problem({
            title: 'Number of SKUs in a comma-separated order list', file: 'pesanan_q4.xlsx', sheet: S, alt, fill: true, ansLabel: 'SKUs',
            statement: '<p>Each order lists its SKUs in one cell, each SKU once, separated by a comma and a space: <b>MLK-GOR, KSJ-SUS, SNY-IJO</b> is 3 SKUs. In ' + AT(S) + ', count the SKUs listed in <b>C5</b>' + DOWN(S) + '.</p>',
            solution: '=LEN(C5)-LEN(SUBSTITUTE(C5,",",""))+1', alts: ['=(LEN(C5)-LEN(SUBSTITUTE(C5,", ","")))/2+1','=COUNTA(TEXTSPLIT(C5,","))', '=COLUMNS(TEXTSPLIT(C5,","))'],
            hints: ['Removing every comma makes the text shorter by one character per comma.', 'LEN(C5)-LEN(SUBSTITUTE(C5,",","")) is the number of commas. There is one more SKU than commas.', 'It looks like =LEN(C5)-LEN(SUBSTITUTE(C5,",",""))+1.'],
            explain: 'SUBSTITUTE deletes the commas, so the drop in length is the number of commas. A list with n commas has n+1 items.',
            chips: ['LEN(C5)', 'SUBSTITUTE(C5,",","")', '+1'], fnKey: 'SUBSTITUTE('
          });
        },
        () => {
          const rows = H.sales(r, 12).map((s) => Object.assign(s, { note: messNote(r.pick(NOTES)), rep: r.pick(H.REPS) }));
          const S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'Rep', k: 'rep' }, { h: 'Visit note', k: 'note' }], rows);
          const alt = H.edit(S.cells, { [S.cell('Visit note', 3)]: (x) => x + ' and  paid', [S.cell('Visit note', 8)]: messNote('No order') });
          return H.problem({
            title: 'Count the words in hand-typed visit notes', file: 'catatan_kunjungan.xlsx', sheet: S, alt, fill: true, ansLabel: 'Words',
            statement: '<p>Reps type notes with stray spaces: at the start, at the end and two or three between words. In ' + AT(S) + ', count the words in the note in <b>C5</b>; words are separated by one or more spaces' + DOWN(S) + '.</p>',
            solution: '=LEN(TRIM(C5))-LEN(SUBSTITUTE(TRIM(C5)," ",""))+1', alts: ['=LEN(TRIM(C5))-LEN(SUBSTITUTE(C5," ",""))+1', '=COUNTA(TEXTSPLIT(TRIM(C5)," "))'],
            hints: ['Counting spaces in the raw note counts the extra ones too. Clean it first.', 'In TRIM(C5) each gap is one space, so the words are the spaces plus 1. Count spaces with LEN and SUBSTITUTE.', 'It looks like =LEN(TRIM(C5))-LEN(SUBSTITUTE(TRIM(C5)," ",""))+1.'],
            explain: 'TRIM leaves exactly one space between words, so the drop in length when the spaces are removed is the number of gaps, and there is one more word than gaps.',
            chips: ['TRIM(C5)', 'LEN(', 'SUBSTITUTE(TRIM(C5)," ","")', '+1'], fnKey: 'SUBSTITUTE('
          });
        },
        () => {
          const line = () => r.sample(H.PRODUCTS, r.int(2, 5)).map((p) => messCase(r, p.sku) + ' x' + r.int(2, 30)).join(', ');
          const rows = H.invoices(r, 12).map((x) => Object.assign(x, { lines: line() }));
          const S = H.table([{ h: 'Order', k: 'ref' }, { h: 'Order lines', k: 'lines' }], rows);
          const fams = r.sample(['MLK', 'KSJ', 'SNY', 'SRW'], 2);
          const base = Object.assign({}, S.cells, { A1: 'Family code', B1: fams[0] });
          const alt = H.edit(base, { B1: fams[1].toLowerCase() });
          return H.problem({
            title: 'How often a product family appears in an order string', file: 'pesanan_q4.xlsx', sheet: S, base, alt, fill: true, ansLabel: 'Lines of family',
            statement: '<p>Order lines were typed in any capitals: <b>mlk-gor x10, KSJ-SUS x4, Mlk-Sot x2</b>. In ' + AT(S) + ', count how many times the family code in <b>B1</b> appears in <b>B5</b>, ignoring capitals (MLK appears twice above)' + DOWN(S) + '. The code in B1 may change, and may be typed in small letters.</p>',
            solution: '=(LEN(B5)-LEN(SUBSTITUTE(UPPER(B5),UPPER($B$1),"")))/LEN($B$1)', alts: ['=(LEN(B5)-LEN(SUBSTITUTE(LOWER(B5),LOWER($B$1),"")))/LEN($B$1)', '=(LEN(UPPER(B5))-LEN(SUBSTITUTE(UPPER(B5),UPPER($B$1),"")))/LEN($B$1)'],
            hints: ['SUBSTITUTE is case-sensitive, so make both texts capitals first with UPPER.', 'Removing every copy of the code shortens the text by LEN($B$1) per copy. Divide the drop by LEN($B$1).', 'It looks like =(LEN(B5)-LEN(SUBSTITUTE(UPPER(B5),UPPER($B$1),"")))/LEN($B$1).'],
            explain: 'UPPER on both sides makes the search ignore capitals. Each removed copy of the code shortens the text by its length, so the drop divided by LEN($B$1) is the count.',
            chips: ['UPPER(B5)', 'UPPER($B$1)', 'SUBSTITUTE(', 'LEN($B$1)'], fnKey: 'SUBSTITUTE('
          });
        }
      ][v]();
    }
  });

  P.add({
    key: 'build-reference-codes', difficulty: 'Hard', level: 8, topics: ['TEXT', 'UPPER', 'LEFT', 'LOWER', 'SUBSTITUTE', '&'], count: 2,
    make(r, v) {
      return [
        () => {
          const rows = H.sales(r, 12).map((s) => Object.assign(s, { seq: r.pick([r.int(1, 9), r.int(10, 99), r.int(100, 999), r.int(1000, 2500)]), city: DISTS().find((d) => d.id === s.distributor_id).city }));
          const S = H.table([{ h: 'Date', k: 'date', date: true }, { h: 'Seq no.', k: 'seq', num: true }, { h: 'City', k: 'city' }, { h: 'Distributor', k: 'distributor' }], rows);
          const alt = H.edit(S.cells, { [S.cell('Seq no.', 2)]: (x) => x + 40, [S.cell('Date', 7)]: (x) => x + 32 });
          return H.problem({
            title: 'Build RK/year/month/number/city references', file: 'register_faktur.xlsx', sheet: S, alt, fill: true, ansLabel: 'Reference', exactCase: true,
            statement: '<p>Rasa Kenari\'s new references read <b>RK/2026/10/0042/BEK</b>: RK, the 4-digit year and 2-digit month of the date in A5, the sequence number in B5 in 4 digits with leading zeros, and the first 3 letters of the city in C5 in capitals, all separated by slashes. Build it in ' + AT(S) + DOWN(S) + '.</p>' + ENGLISH,
            solution: '="RK/"&TEXT(A5,"yyyy")&"/"&TEXT(A5,"mm")&"/"&TEXT(B5,"0000")&"/"&UPPER(LEFT(C5,3))', alts: ['=TEXTJOIN("/",TRUE,"RK",TEXT(A5,"yyyy"),TEXT(A5,"mm"),TEXT(B5,"0000"),UPPER(LEFT(C5,3)))', '="RK/"&YEAR(A5)&"/"&TEXT(MONTH(A5),"00")&"/"&TEXT(B5,"0000")&"/"&UPPER(LEFT(C5,3))'],
            hints: ['Build five pieces and join them with "/": RK, year, month, number, city.', 'TEXT(A5,"yyyy") and TEXT(A5,"mm") give the year and month, TEXT(B5,"0000") pads the number, UPPER(LEFT(C5,3)) gives the city code.', 'It looks like ="RK/"&TEXT(A5,"yyyy")&"/"&TEXT(A5,"mm")&"/"&TEXT(B5,"0000")&"/"&UPPER(LEFT(C5,3)).'],
            explain: 'Each part is formatted on its own: TEXT pads the month and the number with zeros, and UPPER(LEFT(C5,3)) turns Bekasi into BEK. TEXTJOIN with "/" is a shorter way to join the same five parts.',
            chips: ['"RK/"&', 'TEXT(A5,"yyyy")', 'TEXT(A5,"mm")', 'TEXT(B5,"0000")', 'UPPER(LEFT(C5,3))', '&"/"&'], fnKey: 'TEXT('
          });
        },
        () => {
          const ends = ['2026-10-31', '2026-11-30', '2026-12-31'].map(H.toSerial);
          const pool = r.shuffle(DISTS());
          const rows = pool.slice(0, 12).map((d) => ({ date: r.pick(ends), typed: messSpace(r, messCase(r, d.name)) }));
          const S = H.table([{ h: 'Statement date', k: 'date', date: true }, { h: 'Distributor (typed)', k: 'typed' }], rows);
          const shift = (x) => ends[(ends.indexOf(x) + 1) % 3];
          const alt = H.edit(S.cells, { [S.cell('Statement date', 1)]: shift, [S.cell('Distributor (typed)', 6)]: messSpace(r, messCase(r, pool[12].name)) });
          return H.problem({
            title: 'File names from a distributor name and a statement date', file: 'kirim_statement.xlsx', sheet: S, alt, fill: true, ansLabel: 'File name', exactCase: true,
            statement: '<p>Each statement is saved as <b>cv_nuri_makmur_20261031.xlsx</b>: the name in <b>B5</b> with its extra spaces removed, in small letters, with an underscore for each space, then an underscore, the date in <b>A5</b> as yyyymmdd, and <b>.xlsx</b>. Build it in ' + AT(S) + DOWN(S) + '.</p>' + ENGLISH,
            solution: '=LOWER(SUBSTITUTE(TRIM(B5)," ","_"))&"_"&TEXT(A5,"yyyymmdd")&".xlsx"', alts: ['=SUBSTITUTE(LOWER(TRIM(B5))," ","_")&"_"&TEXT(A5,"yyyymmdd")&".xlsx"', '=LOWER(SUBSTITUTE(TRIM(B5)," ","_")&"_"&TEXT(A5,"yyyymmdd")&".xlsx")'],
            hints: ['Work from the inside out: TRIM first, or the doubled spaces become doubled underscores.', 'SUBSTITUTE(TRIM(B5)," ","_") gives the underscores and LOWER the small letters. TEXT(A5,"yyyymmdd") gives the date part.', 'It looks like =LOWER(SUBSTITUTE(TRIM(B5)," ","_"))&"_"&TEXT(A5,"yyyymmdd")&".xlsx".'],
            explain: 'TRIM leaves single spaces, SUBSTITUTE turns each into an underscore and LOWER fixes the capitals. TEXT writes the date as digits, and & adds the separator and the extension.',
            chips: ['TRIM(B5)', 'SUBSTITUTE(', '," ","_")', 'LOWER(', 'TEXT(A5,"yyyymmdd")', '&".xlsx"'], fnKey: 'SUBSTITUTE('
          });
        }
      ][v]();
    }
  });

  P.add({
    key: 'nth-field-of-a-line', difficulty: 'Hard', level: 10, topics: ['MID', 'SUBSTITUTE', 'REPT', 'TRIM', 'LEN'], count: 2,
    make(r, v) {
      return [
        () => {
          const city = (s) => DISTS().find((d) => d.id === s.distributor_id).city;
          const rows = H.sales(r, 12).map((s) => Object.assign(s, { line: [s.region, city(s), s.distributor, s.sku].join('|') }));
          const S = H.table([{ h: 'Export line', k: 'line' }, { h: 'Cartons', k: 'cartons', num: true }], rows);
          const n = r.int(2, 4);
          const base = Object.assign({}, S.cells, { A1: 'Field number', B1: n });
          const alt = H.edit(base, { B1: n === 2 ? 3 : n - 1 });
          return H.problem({
            title: 'Pick the field number given in B1 from a pipe-separated line', file: 'export_sistem.xlsx', sheet: S, base, alt, fill: true, ansLabel: 'Field', exactCase: true,
            statement: '<p>The old system exports one line per sale: <b>region|city|distributor|SKU</b>, separated by the pipe sign |. In ' + AT(S) + ', return field number <b>B1</b> of the line in <b>A5</b> (1 is the region, 4 the SKU)' + DOWN(S) + '. B1 will change, so the formula must work for any field.</p>',
            solution: '=TRIM(MID(SUBSTITUTE(A5,"|",REPT(" ",99)),($B$1-1)*99+1,99))', alts: ['=INDEX(TEXTSPLIT(A5,"|"),$B$1)', '=CHOOSECOLS(TEXTSPLIT(A5,"|"),$B$1)', '=TEXTAFTER(TEXTBEFORE("|"&A5&"|","|",$B$1+1),"|",-1)'],
            hints: ['A trick: replace every | with 99 spaces. Then field n sits somewhere inside characters (n-1)*99+1 to n*99.', 'SUBSTITUTE(A5,"|",REPT(" ",99)) spreads the fields out. MID that from ($B$1-1)*99+1 for 99 characters, then TRIM the padding away.', 'It looks like =TRIM(MID(SUBSTITUTE(A5,"|",REPT(" ",99)),($B$1-1)*99+1,99)). In Excel 365, INDEX(TEXTSPLIT(A5,"|"),$B$1) also works.'],
            explain: 'Swapping each pipe for 99 spaces puts every field in its own 99-character block, so MID can cut block number B1 and TRIM removes the padding. The spaces inside a field such as Jawa Barat survive, because TRIM keeps single spaces.',
            chips: ['SUBSTITUTE(A5,"|",REPT(" ",99))', '($B$1-1)*99+1', ',99))', 'TRIM(MID('], fnKey: 'SUBSTITUTE('
          });
        },
        () => {
          const path = (s, depth) => ['Rasa Kenari', s.region, DISTS().find((d) => d.id === s.distributor_id).city, s.distributor, s.sku].slice(0, depth).join('/');
          const rows = H.sales(r, 12).map((s) => Object.assign(s, { depth: r.int(2, 5) }));
          rows.forEach((s) => { s.path = path(s, s.depth); });
          const S = H.table([{ h: 'Sales path', k: 'path' }, { h: 'Net (Rp)', k: 'net', num: true }], rows);
          const deeper = (i) => path(rows[i], rows[i].depth === 5 ? 3 : rows[i].depth + 1);
          const alt = H.edit(S.cells, { [S.cell('Sales path', 1)]: deeper(1), [S.cell('Sales path', 7)]: deeper(7) });
          return H.problem({
            title: 'Last part of a sales path of any depth', file: 'hirarki_penjualan.xlsx', sheet: S, alt, fill: true, ansLabel: 'Last part', exactCase: true,
            statement: '<p>Each path goes from the company down to the level where the sale was booked, separated by slashes: <b>Rasa Kenari/Jawa Barat/Bandung</b> ends at the city, others go down to the distributor or the SKU. In ' + AT(S) + ', return the text after the <b>last</b> slash in <b>A5</b>' + DOWN(S) + '.</p>',
            solution: '=MID(A5,FIND("#",SUBSTITUTE(A5,"/","#",LEN(A5)-LEN(SUBSTITUTE(A5,"/",""))))+1,99)', alts: ['=TEXTAFTER(A5,"/",-1)', '=TRIM(RIGHT(SUBSTITUTE(A5,"/",REPT(" ",99)),99))'],
            hints: ['The number of slashes changes. Count them first: LEN(A5)-LEN(SUBSTITUTE(A5,"/","")).', 'SUBSTITUTE with that count as the fourth argument marks only the last slash, for example with #. FIND("#",...) is its position.', 'It looks like =MID(A5,FIND("#",SUBSTITUTE(A5,"/","#",LEN(A5)-LEN(SUBSTITUTE(A5,"/",""))))+1,99). In Excel 365, TEXTAFTER(A5,"/",-1) does the same.'],
            explain: 'Counting the slashes tells SUBSTITUTE which one is last, so only that slash becomes #. FIND locates the # and MID reads everything after it, at any depth.',
            chips: ['LEN(A5)-LEN(SUBSTITUTE(A5,"/",""))', 'SUBSTITUTE(A5,"/","#",', 'FIND("#",', '+1,99)'], fnKey: 'SUBSTITUTE('
          });
        }
      ][v]();
    }
  });
});
