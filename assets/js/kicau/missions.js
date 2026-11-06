/*
 * Chapter 1 of Kicau, "Your first week": seven formula requests and one macro, from five colleagues.
 *
 * Every formula mission is judged by core/check.js on two sheets. The second sheet ("alt") is the
 * same file after a normal working day: two cancelled lines deleted at the bottom, a few numbers
 * corrected, and, when the request points at B1, a different name in B1. A typed answer, or a
 * formula that only works on today's numbers, passes the first sheet and fails the second, and the
 * colleague says so in their own words.
 *
 * Text placeholders, filled in by the app: {name} the learner, {value} what their formula gave,
 * {target} and {other} the two distributor names, {altD5} row 5's cartons on the second sheet.
 * Classic script in the browser (window.KicauMissions), CommonJS in Node.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('../core/data.js'));
  else root.KicauMissions = factory(root.KenariData);
})(typeof globalThis !== 'undefined' ? globalThis : this, function (D) {
  'use strict';

  const { rows, target, other, example } = D.missionSales();
  const FIRST = 5, LAST = FIRST + rows.length - 1;
  const R = (col) => col + FIRST + ':' + col + LAST;
  const FILE = 'penjualan_oktober.xlsx';

  /** The October sheet: a header on row 4, twenty sales lines from row 5. `top` fills rows 1 to 2. */
  function sheet(top) {
    const c = { A4: 'Date', B4: 'Distributor', C4: 'Product', D4: 'Cartons', E4: 'Net (Rp)' };
    rows.forEach((r, i) => {
      const n = FIRST + i;
      c['A' + n] = D.toSerial(r.date);
      c['B' + n] = r.distributor;
      c['C' + n] = r.product;
      c['D' + n] = r.cartons;
      c['E' + n] = r.net;
    });
    return Object.assign(c, top);
  }

  const ALT_D5 = rows[0].cartons >= 30 ? 12 : 41;

  /** The same file after a working day. Every formula mission must survive it. */
  function nextDay(c) {
    const a = Object.assign({}, c);
    for (const col of 'ABCDE') { delete a[col + (LAST - 1)]; delete a[col + LAST]; }
    a.D5 = ALT_D5;
    a.E6 = c.E6 + 1250000;
    a.D7 = c.D7 + 5;
    /* A price correction on the biggest sale, so MAX moves too. */
    let top = FIRST + 2;
    for (let n = FIRST + 2; n <= LAST - 2; n++) if (a['E' + n] > a['E' + top]) top = n;
    a['E' + top] += 400000;
    return a;
  }

  const GENERIC = {
    typed: { title: 'That is a typed value', plain: 'A typed answer stays the same when the sheet changes. Start with = so Excel works it out.' },
    value: { title: 'Not the number they need', plain: 'Check which column you used, and which rows.' },
    frozen: { title: 'Right today, wrong tomorrow', plain: 'When the sheet changed, your answer did not follow. Point at cells and ranges instead of typing numbers or names.' }
  };

  /** Fill in the shared parts of a formula mission. */
  function formula(o) {
    const top = typeof o.top === 'function' ? o.top : () => o.top;
    return Object.assign({
      kind: 'formula', file: FILE, answerCell: 'B2', precision: 2,
      variants: [{ id: 'alt', label: 'the next day' }],
      variantCells: { base: sheet(top('base')), alt: nextDay(sheet(top('alt'))) },
      cells: sheet(top('base')),
      colRange: { A: R('A'), B: R('B'), C: R('C'), D: R('D'), E: R('E') },
      first: FIRST, last: LAST,
      traps: [], requires: [],
      msgs: Object.assign({}, GENERIC, o.msgs || {})
    }, o, { msgs: Object.assign({}, GENERIC, o.msgs || {}) });
  }

  const LIST = [
    formula({
      id: '1.1', from: 'tika', due: '09:30', title: 'Total cartons',
      intro: ['Pagi, {name}! Welcome to Sales Finance 👋', 'An easy one to start your week. How many cartons did we ship to these five distributors in October?', '@file', 'Put the total in <b>B2</b> please.'],
      ask: 'Total of the <b>Cartons</b> column, in <b>B2</b>.',
      top: { A1: 'October 2026', A2: 'Total cartons' },
      lesson: {
        idea: 'SUM adds up every number in a range.',
        teach: '<p>A <b>range</b> is a block of cells: the first cell, a colon, the last cell. <code>D5:D24</code> means D5 down to D24.</p><p><code>=SUM(D5:D24)</code> adds all twenty. Change any number and the total follows.</p>',
        ex: [['=SUM(', 'Add up'], [R('E'), 'everything from E5 down to E24'], [')', 'and close the bracket.']],
        exFormula: '=SUM(' + R('E') + ')', exLabel: 'rupiah of net sales',
        bridge: 'Like footing a ledger column, except it re-foots itself when a number changes.'
      },
      starter: '=SUM(', fnKey: 'SUM(', chips: [R('D'), R('E')],
      solution: '=SUM(' + R('D') + ')',
      alts: ['=SUM($D$5:$D$24)', '=SUM(D:D)'],
      traps: [{ code: 'net', formula: '=SUM(' + R('E') + ')', title: 'That is rupiah, not cartons', plain: 'Column E holds Net (Rp). Cartons are in column D.', reply: 'Hmm, {value} cartons? That would fill every warehouse in Java 😅 I think that is the <b>Net (Rp)</b> column. I need <b>Cartons</b>.' }],
      requires: [{ re: /SUM\s*\(/i, title: 'Right total, long way round', plain: 'Adding cells one by one works today and breaks when a row is added. SUM takes the whole range.', reply: 'That is the right number! Mbak Sari says adding cells one by one breaks when rows change though. Can you use SUM?' }],
      replies: { pass: ['Mantap, {value} cartons! 🙏', 'That is exactly my number for the morning report.'], frozen: 'I deleted two cancelled lines at the bottom and your total did not change. Did you type the number?', typed: 'Did you type the number in? When I fix a line it will not update 😅 Can you make it a formula?' },
      note: { fn: 'SUM', what: 'Adds every number in a range.', shape: '=SUM(D5:D24)' },
      hints: ['SUM needs one thing inside its brackets: the range to add.', 'Cartons are in column D, rows 5 to 24. Tap the D at the top of the grid.', 'It looks like =SUM(D5:D24).']
    }),

    formula({
      id: '1.2', from: 'bayu', due: '10:15', title: 'Count the deliveries',
      intro: ['Halo {name}, Bayu from the gudang 👷', 'Tika forwarded me her October file.', '@file', 'How many invoice lines are in it? Each line is one drop for my drivers. Answer in <b>B2</b>.'],
      ask: 'How many invoice lines, in <b>B2</b>.',
      top: { A1: 'October 2026', A2: 'Invoice lines' },
      lesson: {
        idea: 'COUNT counts the cells that hold numbers.',
        teach: '<p><code>COUNT</code> counts cells with numbers and skips blanks. <code>COUNTA</code> counts every cell that is not empty, words included.</p><p>So <code>=COUNT(D5:D24)</code> and <code>=COUNTA(B5:B24)</code> both count the lines here.</p>',
        ex: [['=COUNT(', 'Count the numbers'], ['D5:D14', 'in the first ten rows of Cartons'], [')', 'and close.']],
        exFormula: '=COUNT(D5:D14)', exLabel: 'lines in the first ten rows',
        bridge: 'Like counting vouchers in a batch before you post them.'
      },
      starter: '=COUNT(', fnKey: 'COUNT(', chips: [R('B'), R('D')],
      solution: '=COUNT(' + R('D') + ')',
      alts: ['=COUNTA(' + R('B') + ')', '=COUNTA(' + R('C') + ')', '=COUNT(' + R('E') + ')'],
      traps: [{ code: 'summed', formula: '=SUM(' + R('D') + ')', title: 'That adds them up', plain: 'SUM adds the numbers. COUNT says how many there are.', reply: '{value} drops? My trucks would never sleep 😂 That looks like the carton total. I need how many <b>lines</b>.' }],
      requires: [{ re: /COUNTA?\s*\(/i, title: 'Right number, other tool', plain: 'This mission practices COUNT or COUNTA.', reply: 'Number is right! Mbak Sari wants you to practice COUNT on this one though 🙂' }],
      replies: { pass: ['{value} lines, makasih! That is {value} drops for my team 🚚'], frozen: 'Two lines at the bottom were cancelled, so I deleted them, and your count stayed at {value}. Can it count only the filled rows?', typed: 'Did you count by hand? 😄 Make it a formula, the file changes every day.' },
      note: { fn: 'COUNT, COUNTA', what: 'COUNT counts numbers. COUNTA counts anything that is not blank.', shape: '=COUNT(D5:D24)' },
      hints: ['You want how many lines, not their total.', 'COUNT counts the cells with numbers in a range, like the Cartons column.', 'It looks like =COUNT(D5:D24).']
    }),

    formula({
      id: '1.3', from: 'tika', due: '11:00', title: 'Average sale',
      intro: ['Thanks for this morning!', 'Pak Anton now wants the <b>average sale per line</b>, in rupiah.', '@file', 'Same file, answer in <b>B2</b>.'],
      ask: 'Average <b>Net (Rp)</b> per line, in <b>B2</b>.',
      top: { A1: 'October 2026', A2: 'Average sale (Rp)' },
      lesson: {
        idea: 'AVERAGE adds a range and divides by how many numbers it holds.',
        teach: '<p><code>=AVERAGE(E5:E24)</code> is the total divided by the count, worked out together.</p><p>Blank cells are skipped, so deleted lines do not drag the average down.</p>',
        ex: [['=AVERAGE(', 'Find the average'], [R('D'), 'of the Cartons column'], [')', 'and close.']],
        exFormula: '=AVERAGE(' + R('D') + ')', exLabel: 'cartons per line',
        bridge: 'Like the average balance on a statement: total over count.'
      },
      starter: '=AVERAGE(', fnKey: 'AVERAGE(', chips: [R('D'), R('E')],
      solution: '=AVERAGE(' + R('E') + ')',
      alts: ['=SUM(' + R('E') + ')/COUNT(' + R('E') + ')'],
      traps: [
        { code: 'total', formula: '=SUM(' + R('E') + ')', title: 'That is the total', plain: 'SUM gives the total. AVERAGE divides it by the number of lines.', reply: 'Rp {value} per line? That is the whole month in one sale 😄 I think that is the total. I need the <b>average</b>.' },
        { code: 'cartons', formula: '=AVERAGE(' + R('D') + ')', title: 'That is cartons', plain: 'This averages the Cartons column. Pak Anton asked for rupiah, column E.', reply: '{value}? That looks like cartons, not rupiah. Pak Anton wants <b>Net (Rp)</b>.' }
      ],
      replies: { pass: ['Rp {value} per line on average. Pak Anton will like that 👍'], frozen: 'I removed two cancelled lines and the average went wrong. Dividing by 20 stays 20 even when lines disappear. Can AVERAGE do the counting?', typed: 'Is that typed in? The numbers change every day, so I need a formula 🙏' },
      note: { fn: 'AVERAGE', what: 'The total of a range divided by how many numbers it holds. Blanks are skipped.', shape: '=AVERAGE(E5:E24)' },
      hints: ['Average is the total divided by the count. One function does both.', 'Net (Rp) is column E, rows 5 to 24.', 'It looks like =AVERAGE(E5:E24).']
    }),

    formula({
      id: '1.4', from: 'anton', due: '13:30', title: 'Biggest single sale',
      intro: ['Siang, {name}. Anton here, Area Sales.', 'I meet the distributors on Thursday.', '@file', 'What was our <b>single biggest sale</b> in October, in rupiah? <b>B2</b> please.'],
      ask: 'The biggest single <b>Net (Rp)</b>, in <b>B2</b>.',
      top: { A1: 'October 2026', A2: 'Biggest sale (Rp)' },
      lesson: {
        idea: 'MAX finds the biggest number in a range. MIN finds the smallest.',
        teach: '<p><code>=MAX(E5:E24)</code> looks through every cell and returns the largest. <code>=MIN</code> returns the smallest.</p><p>Both ignore blanks and words.</p>',
        ex: [['=MIN(', 'Find the smallest'], [R('D'), 'in the Cartons column'], [')', 'and close.']],
        exFormula: '=MIN(' + R('D') + ')', exLabel: 'cartons in the smallest order',
        bridge: 'Like scanning a ledger for the largest single posting.'
      },
      starter: '=MAX(', fnKey: 'MAX(', chips: [R('D'), R('E')],
      solution: '=MAX(' + R('E') + ')',
      alts: ['=LARGE(' + R('E') + ',1)'],
      traps: [
        { code: 'min', formula: '=MIN(' + R('E') + ')', title: 'That is the smallest', plain: 'MIN returns the smallest number. You want MAX.', reply: 'Rp {value}? That is our smallest one 😄 I asked for the <b>biggest</b>.' },
        { code: 'cartons', formula: '=MAX(' + R('D') + ')', title: 'That is cartons', plain: 'This is the biggest order in cartons. Pak Anton asked for rupiah.', reply: '{value} what, cartons? I need it in <b>rupiah</b>, the Net (Rp) column.' }
      ],
      replies: { pass: ['Rp {value} in one order. Good to know who to thank first.'], frozen: 'Bu Lestari corrected a few lines this morning and your number stayed the same. Is it typed?', typed: 'Did you look for it by eye? 😄 Make it a formula, the file changes.' },
      after: 'Want to know <i>which</i> distributor placed it? That needs XLOOKUP. It is coming in a later chapter.',
      note: { fn: 'MAX, MIN', what: 'The biggest, or the smallest, number in a range.', shape: '=MAX(E5:E24)' },
      hints: ['There is a function for the biggest number in a range.', 'MAX, on the Net (Rp) column, E5 to E24.', 'It looks like =MAX(E5:E24).']
    }),

    formula({
      id: '1.5', from: 'bayu', due: '14:45', title: 'Deliveries for one distributor',
      intro: ['Truck planning again 😅', 'How many deliveries went to the distributor in <b>B1</b>?', '@file', 'Point at B1 please. I will change the name for each truck.'],
      ask: 'Count the lines for the distributor in <b>B1</b>, in <b>B2</b>.',
      top: (v) => ({ A1: 'Distributor', B1: v === 'alt' ? other : target, A2: 'Deliveries' }),
      lesson: {
        idea: 'COUNTIF counts only the cells that match what you ask for.',
        teach: '<p><code>COUNTIF</code> takes a range to look in, then what to look for.</p><p>Point at a cell for what to look for, like <code>B1</code>. When B1 changes, the count follows.</p>',
        ex: [['=COUNTIF(', 'Count only some cells:'], [R('C') + ',', 'look in the Product column'], ['"' + rows[0].product + '")', 'and count this product.']],
        exFormula: '=COUNTIF(' + R('C') + ',"' + rows[0].product + '")', exLabel: 'lines of ' + rows[0].product,
        bridge: 'Like counting the vouchers for one account in a batch.'
      },
      starter: '=COUNTIF(', fnKey: 'COUNTIF(', chips: [R('B'), 'B1'],
      solution: '=COUNTIF(' + R('B') + ',B1)',
      alts: ['=COUNTIFS(' + R('B') + ',B1)', '=COUNTIF($B$5:$B$24,$B$1)'],
      traps: [{ code: 'everyone', formula: '=COUNTA(' + R('B') + ')', title: 'That counts every distributor', plain: 'This counts all the lines. COUNTIF counts only the ones that match B1.', reply: '{value}? That is every line in the file. I only need <b>{target}</b>, the name in B1.' }],
      requires: [{ re: /COUNTIFS?\s*\(/i, title: 'Right number, other tool', plain: 'This mission practices COUNTIF.', reply: 'Right number! Can you do it with COUNTIF though? Mbak Sari says I will need it every day.' }],
      replies: { pass: ['{value} drops for {target}. One truck should do it. Makasih! 🚚'], frozen: 'I changed B1 to {other} and still got {value}. Did you type the name? Point at <b>B1</b> instead.', typed: 'Did you count by hand? 😄 I change B1 all day, so it has to be a formula.' },
      note: { fn: 'COUNTIF', what: 'Counts the cells in a range that match one condition.', shape: '=COUNTIF(B5:B24,B1)' },
      hints: ['COUNTIF needs two things: where to look, and what to look for.', 'Look in the Distributor column, B5 to B24. What to look for is in B1.', 'It looks like =COUNTIF(B5:B24,B1).']
    }),

    formula({
      id: '1.6', from: 'tika', due: '15:30', title: 'Total for one distributor',
      intro: ['Last big one today 🙏', 'In <b>B2</b>, put the total <b>Net (Rp)</b> for the distributor named in <b>B1</b>.', '@file', 'Point at B1 please, I will change the name myself later.'],
      ask: 'Total <b>Net (Rp)</b> for the distributor in <b>B1</b>, in <b>B2</b>.',
      top: (v) => ({ A1: 'Distributor', B1: v === 'alt' ? other : target, A2: 'Total net (Rp)' }),
      lesson: {
        idea: 'SUMIFS adds up only the rows that match what you ask for.',
        teach: '<p><code>SUMIFS</code> takes the column to add first, then a column to test, then what to look for. Rows that fail the test are left out.</p><p>Point at a cell for what to look for. When that cell changes, the total follows.</p>',
        ex: [['=SUMIFS(', 'Add up, but only some rows.'], [R('D') + ',', 'Add the cartons in this column,'], [R('B') + ',', 'testing the distributor in this column,'], ['"' + example + '")', 'and keep the rows for this name.']],
        exFormula: '=SUMIFS(' + R('D') + ',' + R('B') + ',"' + example + '")', exLabel: 'cartons for ' + example,
        bridge: 'Like filtering a ledger to one account and reading the subtotal, with no filter left on the sheet.'
      },
      starter: '=SUMIFS(', fnKey: 'SUMIFS(', chips: [R('E'), R('B'), 'B1'],
      solution: '=SUMIFS(' + R('E') + ',' + R('B') + ',B1)',
      alts: ['=SUMIFS($E$5:$E$24,$B$5:$B$24,$B$1)', '=SUMIFS(' + R('E') + ';' + R('B') + ';B1)', '=SUMIF(' + R('B') + ',B1,' + R('E') + ')'],
      traps: [
        { code: 'cartons', formula: '=SUMIFS(' + R('D') + ',' + R('B') + ',B1)', title: 'You added cartons', plain: 'Your total comes from the Cartons column. Tika asked for Net (Rp), so the first range needs to be column E.', reply: 'Hmm, {value}? That looks like cartons, not rupiah 😅 I need the <b>Net (Rp)</b> column for the report.' },
        { code: 'all', formula: '=SUM(' + R('E') + ')', title: 'That is every distributor', plain: 'This adds the whole column. Use SUMIFS, so only the rows for B1 count.', reply: 'Rp {value} is everyone together, I think. I only need <b>{target}</b>.' }
      ],
      requires: [{ re: /SUMIFS?\s*\(/i, title: 'Right number, older tool', plain: 'This total is correct. This mission practices SUMIFS.', reply: 'The number is right! Mbak Sari says SUMIFS is the one to learn, it can test more columns later.' }],
      replies: { pass: ['Mantap, that is it! I changed B1 to {other} and it updated by itself.', 'You just saved me an hour every month 🙌'], frozen: 'I changed B1 to {other} and the total did not move. Did you type the name in the formula? Point at <b>B1</b>.', typed: 'Is that typed in? I change B1 all the time, so it needs to be a formula 🙏' },
      note: { fn: 'SUMIFS', what: 'Adds the rows that match. Column to add, column to test, what to look for.', shape: '=SUMIFS(E5:E24,B5:B24,B1)' },
      hints: ['SUMIFS wants three things here: what to add, where to look, and what to look for.', 'What to add is column E. Where to look is column B. What to look for sits in B1.', 'It looks like =SUMIFS(E5:E24,B5:B24,B1).']
    }),

    formula({
      id: '1.7', from: 'lestari', due: '16:15', title: 'Flag a big order',
      intro: ['Hai {name}, Bu Lestari from Credit Control.', 'Big orders need my sign-off. If row 5 is <b>30 cartons or more</b>, write <b>Big</b> in B2. Otherwise <b>Normal</b>.', '@file'],
      ask: '<b>Big</b> if D5 is 30 or more, else <b>Normal</b>, in <b>B2</b>.',
      top: { A1: 'Row 5 check', A2: 'Order size' },
      lesson: {
        idea: 'IF picks one of two answers, depending on a test.',
        teach: '<p><code>IF</code> takes a test, the answer when it is true, and the answer when it is false.</p><p>Words go in double quotes: <code>"Big"</code>. Tests use signs like <code>&gt;=</code>, which means greater than or equal to.</p>',
        ex: [['=IF(', 'If'], ['E5>4000000,', 'the net sale in E5 is over 4 million,'], ['"Large",', 'say Large,'], ['"Small")', 'otherwise say Small.']],
        exFormula: '=IF(E5>4000000,"Large","Small")', exLabel: 'for row 5',
        bridge: 'Like an approval limit: above it, the order goes up for sign-off.'
      },
      starter: '=IF(', fnKey: 'IF(', chips: ['D5', '>=', '30', '"Big"', '"Normal"'],
      solution: '=IF(D5>=30,"Big","Normal")',
      alts: ['=IF(D5<30,"Normal","Big")', '=IF(D5>29,"Big","Normal")'],
      traps: [{ code: 'flipped', formula: '=IF(D5>=30,"Normal","Big")', title: 'The answers are swapped', plain: 'The true answer comes first, right after the test.', reply: 'Hmm, it says {value} but row 5 is the other way round, I think. The answer for <b>true</b> comes first.' }],
      requires: [{ re: /IF\s*\(/i, title: 'Use IF', plain: 'This mission practices IF.', reply: 'Right answer! Can you build it with IF? I will copy it down every row.' }],
      replies: { pass: ['Thank you. Row 5 is {value}.', 'I will copy this down the whole column myself 👍'], frozen: 'Bayu corrected row 5 to {altD5} cartons and it still says {value}. It has to check <b>D5</b> itself.', typed: 'Did you type the word in? When row 5 changes it will not update. It has to be a formula.' },
      after: 'Copying a formula down is how one formula serves a thousand rows. That is the next chapter.',
      note: { fn: 'IF', what: 'A test, the answer if true, the answer if false. Words go in double quotes.', shape: '=IF(D5>=30,"Big","Normal")' },
      hints: ['IF needs three parts: the test, the true answer, and the false answer.', 'The test is D5>=30. The answers are "Big" and "Normal", with quotes.', 'It looks like =IF(D5>=30,"Big","Normal").']
    })
  ];

  /* The macro request: step through it, then answer how many rows it flagged. */
  const macroRows = D.macroSheet();
  const flagged = macroRows.filter((r) => r.dpd > 30).length;
  LIST.push({
    id: '1.8', kind: 'macro', from: 'lestari', due: '17:00', title: 'Read a macro', file: 'piutang_senin.xlsm',
    intro: ['One more before you go home 😩', 'Every Monday I mark late distributors by hand.', 'Mbak Sari wrote a macro for it, but I do not trust what I do not understand.', '@file', 'Can you step through it and tell me <b>how many need review</b>?'],
    ask: 'Step through the macro, then answer how many need review.',
    lesson: {
      idea: 'A macro is a list of steps Excel follows, one line at a time.',
      teach: '<p><code>For r = 2 To lastRow</code> repeats the lines inside it once per row. <code>If ... Then</code> decides what to write on each one.</p><p>Step through it line by line and watch the sheet change.</p>',
      bridge: 'Like a written procedure: the same steps for every voucher in the batch.'
    },
    rows: macroRows, answer: flagged,
    choices: [flagged - 1, flagged, flagged + 1].filter((n) => n >= 0),
    replies: { pass: ['{value}, and now I know why. Terima kasih! 🙏', 'I will run it myself next Monday.'], wrong: 'Hmm, I count differently. Can you step through it to the end and look at F1?' },
    note: { fn: 'Macros: For and If', what: 'For repeats lines once per row. If decides what happens on each one.', shape: 'For r = 2 To lastRow ... Next r' }
  });

  /* ---------- Chapter 2: Bank Kenari ----------
     The bank's October collection list: one row per credit facility. Days late are the collection team's own
     count for this list, spread a little wider than the month-end snapshot so every rule gets exercised. */
  const FAC = D.build().facilities;
  const B_FIRST = 5, B_LAST = B_FIRST + FAC.length - 1;
  const BR = (col) => col + B_FIRST + ':' + col + B_LAST;
  const LATE = { 2: 35, 6: 140, 10: 90, 13: 210 };
  const UNPAID = [4, 8, 12, 13];
  const BANK_FILE = 'koleksi_oktober.xlsx';
  const BCOLS = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'];

  /** The collection list. `top` fills rows 1 to 2, `ans` names column I. */
  function bankSheet(top, ans) {
    const c = { A4: 'Facility', B4: 'Distributor', C4: 'Type', D4: 'Limit (Rp)', E4: 'Outstanding (Rp)', F4: 'Days late', G4: 'Rate %', H4: 'Paid (Rp)' };
    if (ans) c.I4 = ans;
    FAC.forEach((f, i) => {
      const n = B_FIRST + i;
      c['A' + n] = f.facility_id;
      c['B' + n] = f.distributor;
      c['C' + n] = f.type === 'Term Loan' ? 'TL' : 'WC';
      c['D' + n] = f.limit;
      c['E' + n] = f.outstanding;
      c['F' + n] = LATE[i] !== undefined ? LATE[i] : f.dpd;
      c['G' + n] = f.rate_pct;
      c['H' + n] = UNPAID.includes(i) ? 0 : Math.round(f.outstanding * 0.04 / 100000) * 100000;
    });
    return Object.assign(c, top);
  }
  /** The same list a day later: a few loans paid up or slipped, a payment missed, a balance corrected. */
  function bankNextDay(c) {
    const a = Object.assign({}, c);
    a['F' + B_FIRST] = 45;
    a['F' + (B_FIRST + 2)] = 0;
    a['F' + (B_FIRST + 6)] = 185;
    a['E' + (B_FIRST + 1)] = c['E' + (B_FIRST + 1)] + 3000000;
    a['E' + (B_FIRST + 4)] = c['E' + (B_FIRST + 4)] + 2000000;
    a['G' + (B_FIRST + 4)] = c['G' + (B_FIRST + 4)] + 1.5;
    a['H' + (B_FIRST + 3)] = 0;
    a['H' + (B_FIRST + 4)] = 1500000;
    return a;
  }
  /** A Chapter 2 request. Column requests (`fill`) are typed in I5 and copied down to the last row. */
  function bank(o) {
    const top = typeof o.top === 'function' ? o.top : () => (o.top || {});
    const fill = (o.fill !== undefined ? o.fill : !!o.ansLabel) ? { col: "I", first: B_FIRST, last: B_LAST } : null;
    return formula(Object.assign({
      chapter: 2, file: BANK_FILE, cols: BCOLS, num: ['D', 'E', 'H'], dateCols: [],
      first: B_FIRST, last: B_LAST, fill, answerCell: fill ? 'I' + B_FIRST : 'B2',
      colRange: Object.fromEntries(BCOLS.map((c) => [c, BR(c)])),
      variantCells: { base: bankSheet(top('base'), o.ansLabel), alt: bankNextDay(bankSheet(top('alt'), o.ansLabel)) },
      cells: bankSheet(top('base'), o.ansLabel)
    }, o, { top: undefined }));
  }

  LIST.push(
    bank({
      id: '2.1', from: 'dimas', due: '09:15', title: 'Utilization, every row', ansLabel: 'Utilization',
      intro: ['Pagi, {name}! Dimas, Portfolio Analyst at Bank Kenari 👋', 'Citra says you are quick. Here is our October collection list.', '@file', 'For every facility, how much of its limit is used? Outstanding divided by Limit, in column <b>I</b>. Write it in <b>I5</b> and copy it down.'],
      ask: 'Outstanding ÷ Limit for every row, in <b>I5</b>, copied down.',
      lesson: {
        idea: 'Write a formula once, then copy it down the column.',
        teach: '<p>When Excel copies a formula down, the row numbers move with it: <code>=E5/D5</code> in I5 becomes <code>=E6/D6</code> in I6. These are <b>relative</b> references.</p><p>Here, Kicau copies I5 down to the last row for you when you send, like double-clicking the fill handle in Excel.</p>',
        ex: [['=E5', 'Outstanding on this row'], ['/D5', 'divided by the Limit on this row.']],
        exFormula: '=E5/D5', exLabel: 'for ' + FAC[0].facility_id + ', so ' + Math.round(FAC[0].outstanding / FAC[0].limit * 100) + '% used',
        bridge: 'Like the same ratio worked out on every line of a schedule.'
      },
      starter: '=', fnKey: 'E5/D5', chips: ['E5', 'D5', '/'],
      solution: '=E5/D5',
      alts: ['=E5/D5*1', '=$E5/$D5'],
      traps: [{ code: 'flipped', formula: '=D5/E5', title: 'Upside down', plain: 'Limit divided by Outstanding is the other way round. Outstanding goes on top.', reply: 'Hmm, every facility is over 100%? 😅 I think it is upside down. <b>Outstanding</b> on top, Limit below.' }],
      replies: { pass: ['Perfect, the whole column in one go 🙏', 'CV Kepodang and UD Srigunting are above 90%. I will keep an eye on them.'], frozen: 'I updated one balance and your column stayed the same. Is it typed in?', typed: 'Is that typed? It has to be a formula so it follows the balances.' },
      note: { fn: 'Copying down', what: 'Relative references move with the formula: E5 becomes E6 on the next row.', shape: '=E5/D5' },
      hints: ['Utilization is Outstanding divided by Limit.', 'Outstanding is column E, Limit is column D, both on row 5.', 'Type =E5/D5 in I5. Kicau copies it down for you.']
    }),
    bank({
      id: '2.2', from: 'citra', due: '10:00', title: 'Provision with a fixed rate', ansLabel: 'Provision (Rp)',
      intro: ['Halo {name}, Citra here. Selamat datang di bank 👋', 'Policy says we set aside a provision on every loan. The rate is in <b>B1</b>.', '@file', 'Outstanding × the rate in B1, for every row, in column I. Careful: every row uses the <b>same</b> B1.'],
      ask: 'Outstanding × the rate in <b>B1</b>, for every row, in column I.',
      top: (v) => ({ A1: 'Provision rate', B1: v === 'alt' ? 0.015 : 0.01 }),
      lesson: {
        idea: 'A $ pins a reference so it does not move when copied.',
        teach: '<p>Copied down, <code>B1</code> would become B2, B3 and so on, which are empty. <code>$B$1</code> stays on B1 in every row. That is an <b>absolute</b> reference.</p><p>On a laptop, put the caret on B1 and press F4 to add the $ signs. On a phone, tap the $ key.</p>',
        ex: [['=E5', 'Outstanding on this row,'], ['*$B$1', 'times the rate in B1, pinned.']],
        exFormula: '=E5*$B$1', exLabel: 'set aside for ' + FAC[0].facility_id,
        bridge: 'Like one tax rate in a cell at the top of a schedule, used by every line.'
      },
      starter: '=E5*', fnKey: '$B$1', chips: ['E5', 'B1', '$'],
      solution: '=E5*$B$1',
      alts: ['=$B$1*E5', '=E5*B$1'],
      frozenTry: '=E5*0.01',
      traps: [{ code: 'nodollar', formula: '=E5*B1', title: 'B1 slid down the column', plain: 'Copied down, B1 became B2, B3 and so on, which are empty. Pin it as $B$1.', reply: 'Only the first loan has a provision, the rest are 0 🤔 When the formula is copied, B1 slides to B2, B3… Pin it with <b>$B$1</b>.' }],
      replies: { pass: ['That is it. One rate, every row 👍', 'When policy changes I only touch B1.'], frozen: 'Policy changed the rate in B1 to 1.5% and your column did not follow. Did you type 1% into the formula? Point at <b>$B$1</b>.', typed: 'Typed numbers will not follow B1. Make it a formula 🙏' },
      note: { fn: '$ absolute references', what: '$B$1 stays on B1 when copied. F4 adds the $ signs.', shape: '=E5*$B$1' },
      hints: ['Every row needs the same rate, from B1.', 'Copied down, B1 would move. A $ before the letter and the number stops that.', 'It looks like =E5*$B$1.']
    }),
    bank({
      id: '2.3', from: 'citra', due: '11:00', title: 'Review list with OR', ansLabel: 'Status',
      intro: ['Next one.', 'A facility goes on the review list if it is <b>more than 30 days late, OR</b> it uses <b>more than 90%</b> of its limit.', '@file', 'Write <b>Review</b> or <b>OK</b> in column I, every row.'],
      ask: '<b>Review</b> if days late &gt; 30 OR utilization &gt; 90%, else <b>OK</b>.',
      lesson: {
        idea: 'OR is true when at least one of its tests is true.',
        teach: '<p><code>OR(F5&gt;30, E5/D5&gt;0.9)</code> is TRUE if the loan is late, or nearly maxed out, or both.</p><p>Put it inside IF as the test: <code>=IF(OR(…),"Review","OK")</code>.</p>',
        ex: [['=OR(', 'Is at least one true:'], ['F5>30,', 'more than 30 days late,'], ['E5/D5>0.9)', 'or more than 90% used?']],
        exFormula: '=OR(F5>30,E5/D5>0.9)', exLabel: 'for ' + FAC[0].facility_id,
        bridge: 'Like an audit flag: any one red flag puts the file on the list.'
      },
      starter: '=IF(OR(', fnKey: 'IF(OR(', chips: ['F5>30', 'E5/D5>0.9', '"Review"', '"OK"'],
      solution: '=IF(OR(F5>30,E5/D5>0.9),"Review","OK")',
      alts: ['=IF(OR(E5/D5>0.9,F5>30),"Review","OK")', '=IF(AND(F5<=30,E5/D5<=0.9),"OK","Review")'],
      traps: [{ code: 'and', formula: '=IF(AND(F5>30,E5/D5>0.9),"Review","OK")', title: 'AND needs both', plain: 'AND is only true when both tests are true. Citra wants either one, which is OR.', reply: 'Hmm, CV Kepodang is 94 days late but says {value}? Either reason is enough for review, so it is <b>OR</b>, not AND.' }],
      requires: [{ re: /IF\s*\(/i, title: 'Use IF', plain: 'Write Review or OK with IF.', reply: 'I need the words Review or OK please, with IF.' }],
      replies: { pass: ['Mantap. That is my review list for Monday 📋'], frozen: 'Two loans changed overnight and your column did not move. It has to read F and E on each row.', typed: 'Typed words will not change when the loans do 🙏' },
      note: { fn: 'OR', what: 'TRUE when at least one test is true. Usually inside IF.', shape: '=IF(OR(F5>30,E5/D5>0.9),"Review","OK")' },
      hints: ['Either reason is enough, so the test is OR.', 'The two tests are F5>30 and E5/D5>0.9.', 'It looks like =IF(OR(F5>30,E5/D5>0.9),"Review","OK").']
    }),
    bank({
      id: '2.4', from: 'dimas', due: '11:45', title: 'Healthy loans with AND', ansLabel: 'Health',
      intro: ['The opposite list now 😄', 'A facility is <b>Healthy</b> only if it is <b>0 days late AND</b> uses <b>less than 70%</b> of its limit. Otherwise <b>Watch</b>.', '@file'],
      ask: '<b>Healthy</b> if days late = 0 AND utilization &lt; 70%, else <b>Watch</b>.',
      lesson: {
        idea: 'AND is true only when every test is true.',
        teach: '<p><code>AND(F5=0, E5/D5&lt;0.7)</code> needs both: not late, and under 70% used.</p><p>One failed test makes the whole AND false.</p>',
        ex: [['=AND(', 'Are both true:'], ['F5=0,', 'not late at all,'], ['E5/D5<0.7)', 'and under 70% used?']],
        exFormula: '=AND(F5=0,E5/D5<0.7)', exLabel: 'for ' + FAC[0].facility_id,
        bridge: 'Like a clean audit: every check has to pass.'
      },
      starter: '=IF(AND(', fnKey: 'IF(AND(', chips: ['F5=0', 'E5/D5<0.7', '"Healthy"', '"Watch"'],
      solution: '=IF(AND(F5=0,E5/D5<0.7),"Healthy","Watch")',
      alts: ['=IF(AND(E5/D5<0.7,F5=0),"Healthy","Watch")', '=IF(OR(F5>0,E5/D5>=0.7),"Watch","Healthy")'],
      traps: [{ code: 'or', formula: '=IF(OR(F5=0,E5/D5<0.7),"Healthy","Watch")', title: 'OR lets one pass', plain: 'With OR, a loan that is not late but nearly maxed out counts as healthy. Healthy needs both, which is AND.', reply: 'Hmm, a loan at 93% of its limit is {value}? Healthy needs <b>both</b> tests, so it is AND.' }],
      requires: [{ re: /IF\s*\(/i, title: 'Use IF', plain: 'Write Healthy or Watch with IF.', reply: 'Words please, Healthy or Watch, with IF 🙏' }],
      replies: { pass: ['Clean. Healthy and Watch, every row ✅'], frozen: 'A loan went late overnight and it is still marked the same. It has to read the row.', typed: 'If it is typed, it will not follow the loans 😅' },
      note: { fn: 'AND', what: 'TRUE only when every test is true. Usually inside IF.', shape: '=IF(AND(F5=0,E5/D5<0.7),"Healthy","Watch")' },
      hints: ['Both conditions must hold, so the test is AND.', 'The tests are F5=0 and E5/D5<0.7.', 'It looks like =IF(AND(F5=0,E5/D5<0.7),"Healthy","Watch").']
    }),
    bank({
      id: '2.5', from: 'citra', due: '13:30', title: 'Collectibility grade with IFS', ansLabel: 'Kol',
      intro: ['Important one.', 'Bank Kenari grades every loan from 1 to 5 by days late. Our rule:', '<b>0 days</b> = 1 · <b>up to 90</b> = 2 · <b>up to 120</b> = 3 · <b>up to 180</b> = 4 · <b>more</b> = 5', '@file', 'Grade every row in column I please.'],
      ask: 'Grade 1 to 5 by days late, every row, in column I.',
      lesson: {
        idea: 'IFS checks tests in order and returns the answer for the first true one.',
        teach: '<p><code>IFS(F5=0,1, F5&lt;=90,2, …)</code> reads pairs: a test, then its answer. The first true test wins, so put them in order.</p><p>End with <code>TRUE,5</code> to catch everything left. Older Excel uses nested IFs for the same job.</p>',
        ex: [['=IFS(F5=0,1,', 'not late: grade 1,'], ['F5<=90,2,', 'up to 90 days: grade 2,'], ['TRUE,"later")', 'anything else.']],
        exFormula: '=IFS(F5=0,1,F5<=90,2,TRUE,"later")', exLabel: 'for ' + FAC[0].facility_id,
        bridge: 'Like an aging schedule: current, 1 to 90, 91 to 120, and so on.'
      },
      starter: '=IFS(', fnKey: 'IFS(', chips: ['F5=0,1', 'F5<=90,2', 'F5<=120,3', 'F5<=180,4', 'TRUE,5'],
      solution: '=IFS(F5=0,1,F5<=90,2,F5<=120,3,F5<=180,4,TRUE,5)',
      alts: ['=IF(F5=0,1,IF(F5<=90,2,IF(F5<=120,3,IF(F5<=180,4,5))))', '=IFS(F5>180,5,F5>120,4,F5>90,3,F5>0,2,TRUE,1)'],
      traps: [{ code: 'strict', formula: '=IFS(F5=0,1,F5<90,2,F5<120,3,F5<180,4,TRUE,5)', title: 'Exactly 90 days is grade 2', plain: 'The rule says up to 90, which includes 90. Use <= instead of <.', reply: 'Close! ' + FAC[10].distributor + ' is exactly 90 days late, and the rule says <b>up to 90</b> is grade 2. Use <b>&lt;=</b>.' }],
      requires: [{ re: /\bIFS?\s*\(/i, title: 'Use IFS or IF', plain: 'Grade it with IFS, or nested IFs.', reply: 'Can you build it with IFS? Mbak Sari says it reads like our policy.' }],
      replies: { pass: ['Perfect grades 🙏', 'This column is what OJK-style reports are built on. You will see it a lot.'], frozen: 'Two loans moved overnight and the grades did not. It has to read the days late on each row.', typed: 'Typed grades will be wrong tomorrow 😅 Formula please.' },
      note: { fn: 'IFS', what: 'Pairs of test and answer. The first true test wins. TRUE catches the rest.', shape: '=IFS(F5=0,1,F5<=90,2,F5<=120,3,F5<=180,4,TRUE,5)' },
      hints: ['Five grades means five test-and-answer pairs, in order.', 'Start with F5=0,1 then F5<=90,2 and so on. Use <= so 90 counts as up to 90.', 'It looks like =IFS(F5=0,1,F5<=90,2,F5<=120,3,F5<=180,4,TRUE,5).']
    }),
    bank({
      id: '2.6', from: 'dimas', due: '14:15', title: 'Months to clear, with IFERROR', ansLabel: 'Months to clear',
      intro: ['Quick one.', 'At this month\'s payment, how many months until each loan is cleared? Outstanding ÷ Paid.', '@file', 'Some paid nothing, so show <b>No payment</b> there instead of an error.'],
      ask: 'Outstanding ÷ Paid, every row. <b>No payment</b> where Paid is 0.',
      lesson: {
        idea: 'IFERROR shows something friendly when a formula fails.',
        teach: '<p>Dividing by 0 gives <code>#DIV/0!</code>. <code>=IFERROR(E5/H5,"No payment")</code> shows the division when it works, and the words when it fails.</p><p>Use it on purpose, for an error you expect. It also hides mistakes you did not expect.</p>',
        ex: [['=IFERROR(', 'Try this:'], ['E5/H5,', 'outstanding over paid,'], ['"No payment")', 'and if it fails, say this.']],
        exFormula: '=IFERROR(E5/H5,"No payment")', exLabel: 'months for ' + FAC[0].facility_id,
        bridge: 'Like writing "n/a" on a schedule line instead of leaving a broken figure.'
      },
      starter: '=IFERROR(', fnKey: 'IFERROR(', chips: ['E5/H5', '"No payment"'],
      solution: '=IFERROR(E5/H5,"No payment")',
      alts: ['=IF(H5=0,"No payment",E5/H5)', '=IF(H5>0,E5/H5,"No payment")'],
      traps: [{ code: 'flipped', formula: '=IFERROR(H5/E5,"No payment")', title: 'Upside down', plain: 'Paid divided by Outstanding is a fraction of a month. Outstanding goes on top.', reply: 'Clearing in {value} months? I wish 😂 It is upside down: <b>Outstanding</b> on top.' }],
      replies: { pass: ['Nice, no ugly errors on my sheet now 👌'], frozen: 'One distributor missed a payment overnight and your column did not change.', typed: 'Typed numbers will not follow the payments 🙏' },
      note: { fn: 'IFERROR', what: 'Shows the formula, or your fallback when it fails. Use it for errors you expect.', shape: '=IFERROR(E5/H5,"No payment")' },
      hints: ['First the division, then what to show if it fails.', 'Outstanding is E5 and Paid is H5. The fallback is "No payment".', 'It looks like =IFERROR(E5/H5,"No payment").']
    }),
    bank({
      id: '2.7', from: 'citra', due: '15:00', title: 'Monthly interest, rounded', ansLabel: 'Interest (Rp)',
      intro: ['For the billing letters:', 'Monthly interest is Outstanding × Rate % ÷ 100 ÷ 12. Round it to the <b>nearest thousand</b> rupiah, every row.', '@file'],
      ask: 'Outstanding × Rate ÷ 100 ÷ 12, rounded to thousands, every row.',
      lesson: {
        idea: 'ROUND cuts a number to the digits you choose.',
        teach: '<p><code>ROUND(x, 2)</code> keeps two decimals. <code>ROUND(x, 0)</code> rounds to whole rupiah. Negative digits round to the left of the point: <code>ROUND(x, -3)</code> is the nearest thousand.</p><p>Formatting only changes how a number looks. ROUND changes the number itself, so totals add up.</p>',
        ex: [['=ROUND(', 'Round'], ['E5*G5/100/12,', 'this month\'s interest'], ['-3)', 'to the nearest thousand.']],
        exFormula: '=ROUND(E5*G5/100/12,-3)', exLabel: 'rupiah for ' + FAC[0].facility_id,
        bridge: 'Like rounding a journal to thousands so the statement foots.'
      },
      starter: '=ROUND(', fnKey: 'ROUND(', chips: ['E5*G5/100/12', ',-3)', ',0)'],
      solution: '=ROUND(E5*G5/100/12,-3)',
      alts: ['=ROUND(E5*G5/1200,-3)', '=ROUND(E5*(G5/100)/12,-3)', '=MROUND(E5*G5/1200,1000)'],
      traps: [
        { code: 'unrounded', formula: '=E5*G5/100/12', title: 'Not rounded', plain: 'The interest is right, but not rounded. Wrap it in ROUND(…,-3).', reply: 'Rp {value}… with cents? 😅 The letters need the nearest <b>thousand</b>.' },
        { code: 'digits', formula: '=ROUND(E5*G5/100/12,0)', title: 'Rounded to rupiah', plain: '0 rounds to whole rupiah. Thousands is -3.', reply: 'Rounded, but to the rupiah. I need the nearest <b>thousand</b>, that is -3.' }
      ],
      requires: [{ re: /ROUND\s*\(/i, title: 'Use ROUND', plain: 'Round it with ROUND.', reply: 'Right numbers! Can you use ROUND? Mbak Sari will ask 😄' }],
      replies: { pass: ['Perfect, ready for the letters 📨'], frozen: 'A balance was corrected overnight and the interest stayed the same.', typed: 'Typed numbers will not follow the balances 🙏' },
      note: { fn: 'ROUND', what: 'ROUND(x, digits). 2 for cents, 0 for whole, -3 for thousands.', shape: '=ROUND(E5*G5/100/12,-3)' },
      hints: ['Work out the interest, then wrap it in ROUND.', 'The interest is E5*G5/100/12. The nearest thousand is -3 digits.', 'It looks like =ROUND(E5*G5/100/12,-3).']
    }),
    bank({
      id: '2.8', from: 'dimas', due: '15:45', title: 'Count with two conditions', ansLabel: '',
      intro: ['One number for my memo:', 'How many facilities of the type in <b>B1</b> are <b>more than 30 days late</b>?', '@file', 'Answer in <b>B2</b>. Point at B1, I will switch between WC and TL.'],
      ask: 'Count facilities of type <b>B1</b> more than 30 days late, in <b>B2</b>.',
      top: (v) => ({ A1: 'Type', B1: v === 'alt' ? 'TL' : 'WC', A2: 'Late facilities' }),
      lesson: {
        idea: 'COUNTIFS counts rows that pass every condition.',
        teach: '<p><code>COUNTIFS</code> takes pairs: a range, then what it must match. A row counts only when every pair matches.</p><p>A condition can be a comparison in quotes, like <code>"&gt;30"</code>. AVERAGEIFS, SUMIFS, MINIFS and MAXIFS work the same way.</p>',
        ex: [['=COUNTIFS(', 'Count the rows where'], [BR('C') + ',"TL",', 'the type is TL,'], [BR('F') + ',">0")', 'and days late is above 0.']],
        exFormula: '=COUNTIFS(' + BR('C') + ',"TL",' + BR('F') + ',">0")', exLabel: 'late term loans',
        bridge: 'Like counting vouchers that match two fields in a batch.'
      },
      starter: '=COUNTIFS(', fnKey: 'COUNTIFS(', chips: [BR('C'), 'B1', BR('F'), '">30"'],
      solution: '=COUNTIFS(' + BR('C') + ',B1,' + BR('F') + ',">30")',
      alts: ['=COUNTIFS(' + BR('F') + ',">30",' + BR('C') + ',B1)', '=COUNTIFS($C$5:$C$21,$B$1,$F$5:$F$21,">"&30)'],
      frozenTry: '=COUNTIFS(' + BR('C') + ',"WC",' + BR('F') + ',">30")',
      traps: [{ code: 'onecond', formula: '=COUNTIF(' + BR('C') + ',B1)', title: 'Only one condition', plain: 'That counts every facility of the type, late or not. Add the days-late pair.', reply: '{value}? That is every facility of that type. I only need the ones <b>more than 30 days late</b>.' }],
      requires: [{ re: /COUNTIFS\s*\(/i, title: 'Use COUNTIFS', plain: 'Two conditions call for COUNTIFS.', reply: 'Right number! Mbak Sari wants COUNTIFS on this one 🙂' }],
      replies: { pass: ['{value}, noted for the memo. Makasih! 📝'], frozen: 'I switched B1 to TL and the count did not change. Did you type "WC"? Point at <b>B1</b>.', typed: 'Counted by hand? 😄 It needs to follow B1.' },
      note: { fn: 'COUNTIFS', what: 'Range and condition pairs. Counts rows that pass them all. Conditions like ">30" go in quotes.', shape: '=COUNTIFS(C5:C21,B1,F5:F21,">30")' },
      hints: ['Two conditions: the type, and days late.', 'Type is in column C, compared with B1. Days late is column F, compared with ">30".', 'It looks like =COUNTIFS(C5:C21,B1,F5:F21,">30").']
    }),
    bank({
      id: '2.9', from: 'citra', due: '16:15', title: 'Biggest late exposure', ansLabel: '',
      intro: ['For the risk committee:', 'Of the facilities <b>more than 30 days late</b>, what is the <b>biggest outstanding</b>?', '@file', 'One number in <b>B2</b>.'],
      ask: 'Biggest outstanding among loans more than 30 days late, in <b>B2</b>.',
      top: { A1: 'Late = more than 30 days', A2: 'Biggest late (Rp)' },
      lesson: {
        idea: 'MAXIFS finds the biggest number among the rows that match.',
        teach: '<p><code>MAXIFS(what, range, condition)</code> works like SUMIFS: the column to look at first, then the conditions.</p><p>MINIFS finds the smallest, AVERAGEIFS the average.</p>',
        ex: [['=AVERAGEIFS(', 'Average'], [BR('E') + ',', 'the outstanding'], [BR('F') + ',">30")', 'of the late loans.']],
        exFormula: '=AVERAGEIFS(' + BR('E') + ',' + BR('F') + ',">30")', exLabel: 'rupiah average for late loans',
        bridge: 'Like finding the single largest past-due balance in an aging list.'
      },
      starter: '=MAXIFS(', fnKey: 'MAXIFS(', chips: [BR('E'), BR('F'), '">30"'],
      solution: '=MAXIFS(' + BR('E') + ',' + BR('F') + ',">30")',
      alts: ['=MAXIFS(' + BR('E') + ',' + BR('F') + ',">"&30)'],
      traps: [
        { code: 'all', formula: '=MAX(' + BR('E') + ')', title: 'That is every loan', plain: 'MAX looks at every loan. MAXIFS looks only at the late ones.', reply: 'Rp {value} is our biggest loan overall, but it is not late. I need the biggest among the <b>late</b> ones.' },
        { code: 'minifs', formula: '=MINIFS(' + BR('E') + ',' + BR('F') + ',">30")', title: 'That is the smallest', plain: 'MINIFS gives the smallest. You want MAXIFS.', reply: 'Rp {value} is the smallest late one 😄 I asked for the <b>biggest</b>.' }
      ],
      replies: { pass: ['Rp {value}. That goes on page one of the committee pack.'], frozen: 'A loan slipped past 30 days overnight and the number did not change.', typed: 'Typed in? The list changes daily 🙏' },
      note: { fn: 'MAXIFS, MINIFS, AVERAGEIFS', what: 'The biggest, smallest or average of the rows that match.', shape: '=MAXIFS(E5:E21,F5:F21,">30")' },
      hints: ['Biggest, but only among late loans: MAXIFS.', 'Look at Outstanding, column E. The condition is on days late, column F, ">30".', 'It looks like =MAXIFS(E5:E21,F5:F21,">30").']
    }),
    bank({
      id: '2.10', from: 'dimas', due: '16:45', title: 'Weighted average rate', ansLabel: '',
      intro: ['Last formula today 🙏', 'What is our <b>average rate</b>, weighted by outstanding? Big loans should count more than small ones.', '@file', 'Answer in <b>B2</b>.'],
      ask: 'Average rate weighted by outstanding, in <b>B2</b>.',
      top: { A1: 'Weighted by outstanding', A2: 'Average rate %' },
      lesson: {
        idea: 'SUMPRODUCT multiplies two columns row by row and adds the results.',
        teach: '<p><code>SUMPRODUCT(E5:E21, G5:G21)</code> is E5×G5 + E6×G6 + … in one step.</p><p>Divide by the total outstanding and you have a weighted average: big loans pull it more.</p>',
        ex: [['=SUMPRODUCT(', 'Multiply row by row and add:'], [BR('E') + ',', 'outstanding'], [BR('H') + ')', 'times paid.']],
        exFormula: '=SUMPRODUCT(' + BR('D') + ',' + BR('G') + ')/SUM(' + BR('D') + ')', exLabel: '% rate weighted by limit',
        bridge: 'Like a weighted average cost: each line counts by its size.'
      },
      starter: '=SUMPRODUCT(', fnKey: 'SUMPRODUCT(', chips: [BR('E'), BR('G'), '/SUM(' + BR('E') + ')'],
      solution: '=SUMPRODUCT(' + BR('E') + ',' + BR('G') + ')/SUM(' + BR('E') + ')',
      alts: ['=SUMPRODUCT(' + BR('G') + ',' + BR('E') + ')/SUM(' + BR('E') + ')'],
      traps: [{ code: 'simple', formula: '=AVERAGE(' + BR('G') + ')', title: 'Every loan counts the same', plain: 'AVERAGE treats a small loan like a big one. Weight each rate by its outstanding.', reply: '{value}% is the plain average. A Rp 50 juta loan should not count as much as a big one. <b>Weighted</b> please.' }],
      requires: [{ re: /SUMPRODUCT\s*\(/i, title: 'Use SUMPRODUCT', plain: 'This mission practices SUMPRODUCT.', reply: 'Right number! Mbak Sari wants SUMPRODUCT on this one.' }],
      replies: { pass: ['{value}% weighted. That is the number for the treasury team 👍'], frozen: 'A balance was corrected overnight and the rate did not move.', typed: 'Typed? It has to follow the balances 🙏' },
      note: { fn: 'SUMPRODUCT', what: 'Multiplies ranges row by row and adds. Divide by the total for a weighted average.', shape: '=SUMPRODUCT(E5:E21,G5:G21)/SUM(E5:E21)' },
      hints: ['Weighted average = sum of outstanding × rate, divided by total outstanding.', 'SUMPRODUCT(E5:E21,G5:G21) does the multiply and add.', 'It looks like =SUMPRODUCT(E5:E21,G5:G21)/SUM(E5:E21).']
    })
  );

  /* The chapter ends in real Excel: download, grade every loan, upload the file back. */
  const kolSolution = '=IFS(F5=0,1,F5<=90,2,F5<=120,3,F5<=180,4,TRUE,5)';
  LIST.push(Object.assign(bank({
    id: '2.11', from: 'citra', due: '17:30', title: 'Grade it in real Excel', ansLabel: 'Kol',
    intro: ['Last one, and this one is in <b>real Excel</b> 💻', 'Download the collection list, grade every loan in column I with our 1 to 5 rule, and send the file back.', '@file', 'Same rule as before: 0 days = 1, up to 90 = 2, up to 120 = 3, up to 180 = 4, more = 5.'],
    ask: 'In Excel: grade column I, save, and upload the .xlsx.',
    lesson: {
      idea: 'Same formula, real Excel.',
      teach: '<p>Tap <b>Download for Excel</b>, open the file, type the IFS in I5 and double-click the fill handle to copy it down. Save, then tap <b>Upload</b>.</p><p>Kicau reads your file, checks every grade, and checks that column I holds formulas, not typed numbers.</p>',
      bridge: 'Exactly what you will do at work: get a file, fix it, send it back.'
    },
    starter: kolSolution, fnKey: 'IFS(', chips: [],
    solution: kolSolution, alts: [],
    replies: { pass: ['Received, and every grade is right. Terima kasih! 🙏', 'Chapter 2 is done. Welcome to the bank side for real.'], wrong: '', typed: 'The grades are typed in, not formulas. Next month they will be wrong 😅 Can you use IFS in I5 and copy it down?' },
    after: 'Chapter 3 is lookups: VLOOKUP, INDEX and MATCH, and XLOOKUP. It arrives in the next build.',
    note: { fn: 'Real Excel round trip', what: 'Download, fill a column with one formula and the fill handle, save and upload.', shape: kolSolution }
  }), { kind: 'excel' }));

  /* ---------- Chapter 3: Lookups ----------
     One sheet, two tables, as in real work: October's orders on the left (A to D, answers in E), Rasa Kenari's
     price list on the right (G to K), and a small discount table across the top (G1 to K2). */
  const P_FIRST = 5, P_LAST = P_FIRST + D.PRODUCTS.length - 1;
  const PR = (col) => '$' + col + '$' + P_FIRST + ':$' + col + '$' + P_LAST;
  const TABLE = '$G$' + P_FIRST + ':$K$' + P_LAST;
  const SCOLS = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K'];
  const CATS = ['Noodles', 'Coffee', 'Sambal', 'Soap'];
  const DISC = { Noodles: 2, Coffee: 3, Sambal: 0, Soap: 5 };
  const SHOP_FILE = 'pesanan_oktober.xlsx';
  /** Orders, price list and discounts. `unknownRow` puts a brand-new SKU into the orders, for if_not_found. */
  function shopSheet(top, ans, opts) {
    const o = opts || {};
    const c = { A4: 'Date', B4: 'Distributor', C4: 'SKU', D4: 'Cartons', G1: 'Category', G2: 'Discount %', G4: 'SKU', H4: 'Product', I4: 'Category', J4: 'Packs', K4: 'Carton price (Rp)' };
    if (ans) c.E4 = ans;
    CATS.forEach((cat, i) => { const col = 'HIJK'[i]; c[col + '1'] = cat; c[col + '2'] = (o.disc || DISC)[cat]; });
    D.PRODUCTS.forEach((p, i) => {
      const n = P_FIRST + i;
      c['G' + n] = p.sku; c['H' + n] = p.name; c['I' + n] = p.category; c['J' + n] = p.packs_per_carton;
      c['K' + n] = p.carton_price + (o.priceBump && i % 3 === 0 ? 4000 : 0);
    });
    rows.forEach((r, i) => {
      const n = FIRST + i;
      c['A' + n] = D.toSerial(r.date); c['B' + n] = r.distributor; c['C' + n] = r.sku; c['D' + n] = r.cartons;
    });
    if (o.swap) { c['C' + (FIRST + 1)] = 'SRW-SER'; c['C' + (FIRST + 6)] = 'MLK-GOR'; c['D' + (FIRST + 2)] += 6; }
    if (o.unknownRow) c['C' + (FIRST + 4)] = 'KSJ-DGN';
    return Object.assign(c, top);
  }
  /** A Chapter 3 request. Column requests are typed in E5 and copied down to E24. */
  function shop(o) {
    const top = typeof o.top === 'function' ? o.top : () => (o.top || {});
    const fill = o.ansLabel ? { col: 'E', first: FIRST, last: LAST } : null;
    const base = shopSheet(top('base'), o.ansLabel, { unknownRow: o.unknownRow });
    const alt = shopSheet(top('alt'), o.ansLabel, { unknownRow: o.unknownRow, swap: true, priceBump: true, disc: { Noodles: 2, Coffee: 4, Sambal: 1, Soap: 5 } });
    return formula(Object.assign({
      chapter: 3, file: SHOP_FILE, cols: SCOLS, num: ['D', 'J', 'K'], dateCols: ['A'],
      first: FIRST, last: LAST, fill, answerCell: fill ? 'E' + FIRST : 'B2',
      colRange: Object.fromEntries(SCOLS.map((c) => [c, ['G', 'H', 'I', 'J', 'K'].includes(c) ? PR(c) : R(c)])),
      variantCells: { base, alt }, cells: base
    }, o, { top: undefined }));
  }

  LIST.push(
    shop({
      id: '3.1', from: 'bayu', due: '08:45', title: 'Product names with VLOOKUP', ansLabel: 'Product',
      intro: ['Pagi {name}, Bayu again 👷', 'My drivers only see SKU codes on the orders. Can you put the <b>product name</b> next to each one?', '@file', 'The price list on the right has every SKU and its name. Column <b>E</b>, every order.'],
      ask: 'Product name for every order\'s SKU, in column E.',
      lesson: {
        idea: 'VLOOKUP finds a value in the first column of a table and returns a value from the same row.',
        teach: '<p><code>VLOOKUP(what, table, column number, FALSE)</code>: look for the SKU in the first column of the table, then bring back column 2 of that row.</p><p>Always end with <code>FALSE</code> for an exact match. Pin the table with $ so it does not slide when copied.</p>',
        ex: [['=VLOOKUP(C5,', 'Look for this SKU'], [TABLE + ',', 'in the price list,'], ['5,', 'bring back column 5, the price,'], ['FALSE)', 'exact match only.']],
        exFormula: '=VLOOKUP(C5,' + TABLE + ',5,FALSE)', exLabel: 'rupiah per carton for row 5',
        bridge: 'Like finding an account in the chart of accounts and reading its name.'
      },
      starter: '=VLOOKUP(', fnKey: 'VLOOKUP(', chips: ['C5', TABLE, ',2', ',FALSE)'],
      solution: '=VLOOKUP(C5,' + TABLE + ',2,FALSE)',
      alts: ['=VLOOKUP(C5,' + TABLE + ',2,0)', '=XLOOKUP(C5,' + PR('G') + ',' + PR('H') + ')', '=INDEX(' + PR('H') + ',MATCH(C5,' + PR('G') + ',0))'],
      traps: [
        { code: 'col', formula: '=VLOOKUP(C5,' + TABLE + ',3,FALSE)', title: 'Column 3 is the category', plain: 'Count from the table\'s first column: SKU is 1, Product is 2.', reply: '{value}? That is the category. I need the <b>product name</b>, column 2 of the price list.' },
        { code: 'nodollar', formula: '=VLOOKUP(C5,G5:K16,2,FALSE)', title: 'The table slid down', plain: 'Copied down, G5:K16 became G6:K17 and so on, and the top products fell out. Pin it as $G$5:$K$16.', reply: 'The first orders are fine, but further down it says #N/A 😅 When copied, the price list slides down. Pin it with <b>$</b>.' }
      ],
      replies: { pass: ['Mantap! My drivers can read the orders now 🚚'], frozen: 'Two orders were changed this morning and the names did not follow.', typed: 'Typed names will not follow the SKUs 😅' },
      note: { fn: 'VLOOKUP', what: 'Find a value in the first column of a table, return a column from that row. End with FALSE.', shape: '=VLOOKUP(C5,$G$5:$K$16,2,FALSE)' },
      hints: ['Look up the SKU in C5 in the price list, and bring back the name.', 'The table is $G$5:$K$16. The name is its 2nd column. End with FALSE.', 'It looks like =VLOOKUP(C5,$G$5:$K$16,2,FALSE).']
    }),
    shop({
      id: '3.2', from: 'tika', due: '09:30', title: 'Order value from the price list', ansLabel: 'Value (Rp)',
      intro: ['Pagi {name}! 🙏', 'For each order, what is it worth? <b>Cartons × the carton price</b> from the price list.', '@file', 'Column E, every order.'],
      ask: 'Cartons × carton price from the price list, every order.',
      lesson: {
        idea: 'A lookup can sit inside a bigger formula.',
        teach: '<p><code>VLOOKUP(C5, table, 5, FALSE)</code> brings back the carton price. Multiply it by the cartons in D5.</p><p>Excel works out the lookup first, then the multiplication.</p>',
        ex: [['=D5*', 'Cartons times'], ['VLOOKUP(C5,' + TABLE + ',5,FALSE)', 'the price for this SKU.']],
        exFormula: '=D5*VLOOKUP(C5,' + TABLE + ',5,FALSE)', exLabel: 'rupiah for row 5',
        bridge: 'Like quantity × list price on an invoice line.'
      },
      starter: '=D5*VLOOKUP(', fnKey: 'VLOOKUP(', chips: ['D5*', 'C5', TABLE, ',5', ',FALSE)'],
      solution: '=D5*VLOOKUP(C5,' + TABLE + ',5,FALSE)',
      alts: ['=VLOOKUP(C5,' + TABLE + ',5,FALSE)*D5', '=D5*XLOOKUP(C5,' + PR('G') + ',' + PR('K') + ')'],
      traps: [
        { code: 'price', formula: '=VLOOKUP(C5,' + TABLE + ',5,FALSE)', title: 'That is the price of one carton', plain: 'Multiply the price by the cartons in D5.', reply: 'Rp {value} for the whole order? That is one carton 😄 Times the <b>cartons</b> please.' },
        { code: 'packs', formula: '=D5*VLOOKUP(C5,' + TABLE + ',4,FALSE)', title: 'Column 4 is packs', plain: 'Column 4 of the price list is packs per carton. The price is column 5.', reply: '{value} is not rupiah, I think that is packs. The <b>price</b> is column 5.' }
      ],
      replies: { pass: ['Perfect, every order valued 💰'], frozen: 'The price list changed this morning and your values did not.', typed: 'Typed values will not follow the price list 🙏' },
      note: { fn: 'Lookup in a formula', what: 'A lookup can be one part of a calculation: =D5*VLOOKUP(...).', shape: '=D5*VLOOKUP(C5,$G$5:$K$16,5,FALSE)' },
      hints: ['Price × cartons. The price comes from the list.', 'The carton price is column 5 of $G$5:$K$16.', 'It looks like =D5*VLOOKUP(C5,$G$5:$K$16,5,FALSE).']
    }),
    shop({
      id: '3.3', from: 'anton', due: '10:30', title: 'Category with XLOOKUP', ansLabel: 'Category',
      intro: ['Siang {name}. Anton.', 'I want each order\'s <b>category</b>: Noodles, Coffee, Sambal or Soap.', '@file', 'Mbak Sari says you should learn XLOOKUP for this. Column E.'],
      ask: 'Category for every order, with XLOOKUP, in column E.',
      lesson: {
        idea: 'XLOOKUP looks in one column and returns from another, in any direction.',
        teach: '<p><code>XLOOKUP(what, where to look, what to return)</code>. No column numbers to count, no FALSE to remember, and the return column can be left of the lookup column.</p><p>It is in Excel 365 and Excel 2021 and later.</p>',
        ex: [['=XLOOKUP(C5,', 'Find this SKU'], [PR('G') + ',', 'in the SKU column,'], [PR('K') + ')', 'and return the price.']],
        exFormula: '=XLOOKUP(C5,' + PR('G') + ',' + PR('K') + ')', exLabel: 'rupiah per carton for row 5',
        bridge: 'Like matching two columns of a ledger by reference number.'
      },
      starter: '=XLOOKUP(', fnKey: 'XLOOKUP(', chips: ['C5', PR('G'), PR('I')],
      solution: '=XLOOKUP(C5,' + PR('G') + ',' + PR('I') + ')',
      alts: ['=XLOOKUP(C5,$G$5:$G$16,$I$5:$I$16,"")', '=XLOOKUP($C5,' + PR('G') + ',' + PR('I') + ')'],
      traps: [{ code: 'col', formula: '=XLOOKUP(C5,' + PR('G') + ',' + PR('H') + ')', title: 'That returns the name', plain: 'The third part says what to return. Category is column I.', reply: '{value} is the product name. I need the <b>category</b>, column I.' }],
      requires: [{ re: /XLOOKUP\s*\(/i, title: 'Try XLOOKUP', plain: 'This request practices XLOOKUP.', reply: 'Right answers! Sari wants XLOOKUP on this one though 😉' }],
      replies: { pass: ['Good. Now I can see which category each order is 📊'], frozen: 'Orders changed this morning and the categories did not move.', typed: 'Typed? It has to follow the SKUs.' },
      note: { fn: 'XLOOKUP', what: 'XLOOKUP(what, where to look, what to return). Any direction, exact match by default.', shape: '=XLOOKUP(C5,$G$5:$G$16,$I$5:$I$16)' },
      hints: ['Look for the SKU in the SKU column and return the category column.', 'Where to look is $G$5:$G$16, what to return is $I$5:$I$16.', 'It looks like =XLOOKUP(C5,$G$5:$G$16,$I$5:$I$16).']
    }),
    shop({
      id: '3.4', from: 'bayu', due: '11:15', title: 'A new SKU, with if_not_found', ansLabel: 'Product', unknownRow: true,
      intro: ['Problem 😬', 'One order has a SKU that is not on the price list yet, a new coffee.', '@file', 'Names for every order again, but show <b>Not in list</b> instead of #N/A for the new one.'],
      ask: 'Product name for every order, <b>Not in list</b> when the SKU is missing.',
      lesson: {
        idea: 'XLOOKUP has a fourth part: what to show when nothing is found.',
        teach: '<p><code>XLOOKUP(what, where, return, "Not in list")</code> shows your words instead of <code>#N/A</code>.</p><p>With VLOOKUP you would wrap it in IFNA or IFERROR to do the same.</p>',
        ex: [['=XLOOKUP("ABC",', 'Look for a SKU that does not exist'], [PR('G') + ',' + PR('H') + ',', 'in the price list,'], ['"Not in list")', 'and say so.']],
        exFormula: '=XLOOKUP("ABC",' + PR('G') + ',' + PR('H') + ',"Not in list")', exLabel: '',
        bridge: 'Like marking an unmatched item in a reconciliation instead of leaving it blank.'
      },
      starter: '=XLOOKUP(', fnKey: 'XLOOKUP(', chips: ['C5', PR('G'), PR('H'), '"Not in list"'],
      solution: '=XLOOKUP(C5,' + PR('G') + ',' + PR('H') + ',"Not in list")',
      alts: ['=IFNA(VLOOKUP(C5,' + TABLE + ',2,FALSE),"Not in list")', '=IFERROR(XLOOKUP(C5,' + PR('G') + ',' + PR('H') + '),"Not in list")'],
      traps: [{ code: 'nofound', formula: '=XLOOKUP(C5,' + PR('G') + ',' + PR('H') + ')', title: 'The new SKU shows #N/A', plain: 'Add a fourth part to XLOOKUP: what to show when nothing matches.', reply: 'Row ' + (FIRST + 4) + ' says #N/A 😅 Can it say <b>Not in list</b> instead? XLOOKUP has a spot for that.' }],
      replies: { pass: ['Perfect. I will ask Tika to add the new coffee to the list 👍'], frozen: 'Orders changed this morning and your names did not follow.', typed: 'Typed names will not follow the SKUs 😅' },
      note: { fn: 'XLOOKUP if_not_found', what: 'The fourth part of XLOOKUP is shown when nothing matches. IFNA does the same for VLOOKUP.', shape: '=XLOOKUP(C5,$G$5:$G$16,$H$5:$H$16,"Not in list")' },
      hints: ['Same lookup as before, plus what to show when the SKU is missing.', 'Add ,"Not in list" as the fourth part of XLOOKUP.', 'It looks like =XLOOKUP(C5,$G$5:$G$16,$H$5:$H$16,"Not in list").']
    }),
    shop({
      id: '3.5', from: 'bayu', due: '13:15', title: 'SKU from a name, with INDEX and MATCH',
      intro: ['The other way round now.', 'I know the product <b>name</b> in B1, and I need its <b>SKU</b>. The SKU is to the <b>left</b> of the name, so VLOOKUP cannot do it.', '@file', 'Answer in <b>B2</b>.'],
      ask: 'The SKU for the product name in <b>B1</b>, in <b>B2</b>.',
      top: (v) => ({ A1: 'Product', B1: v === 'alt' ? D.PRODUCTS[9].name : D.PRODUCTS[4].name, A2: 'SKU' }),
      lesson: {
        idea: 'MATCH finds the position. INDEX returns what is at that position.',
        teach: '<p><code>MATCH(B1, names, 0)</code> says which row the name is on, like 5. <code>INDEX(SKUs, 5)</code> returns the 5th SKU.</p><p>Together they look left, right, anywhere. This pair works in every Excel version.</p>',
        ex: [['=MATCH(B1,', 'Where is this name'], [PR('H') + ',', 'in the name column,'], ['0)', 'exact match?']],
        exFormula: '=MATCH(B1,' + PR('H') + ',0)', exLabel: 'is the position of the name in B1',
        bridge: 'Like finding a line number in a register, then reading across to another column.'
      },
      starter: '=INDEX(', fnKey: 'INDEX(', chips: [PR('G'), 'MATCH(', 'B1', PR('H'), ',0))'],
      solution: '=INDEX(' + PR('G') + ',MATCH(B1,' + PR('H') + ',0))',
      alts: ['=INDEX(G5:G16,MATCH(B1,H5:H16,0))'],
      traps: [{ code: 'position', formula: '=MATCH(B1,' + PR('H') + ',0)', title: 'That is the position', plain: 'MATCH gives the position. Wrap it in INDEX to get the SKU at that position.', reply: '{value}? That is a row number, not a SKU 😄 Use it inside <b>INDEX</b> to pick the SKU.' }],
      requires: [{ re: /INDEX\s*\(/i, title: 'Try INDEX and MATCH', plain: 'This request practices INDEX with MATCH.', reply: 'Right! XLOOKUP works too, but older laptops in the gudang only have INDEX/MATCH 😅' }],
      replies: { pass: ['{value}, makasih! Now I can type the SKU on the picking list 📦'], frozen: 'I changed B1 to another product and the SKU stayed the same. Did you type it?', typed: 'Typed? I change B1 all day 😄' },
      note: { fn: 'INDEX and MATCH', what: 'MATCH finds the position, INDEX returns the value there. Looks in any direction.', shape: '=INDEX($G$5:$G$16,MATCH(B1,$H$5:$H$16,0))' },
      hints: ['First find which row the name is on, then read the SKU on that row.', 'MATCH(B1,$H$5:$H$16,0) gives the row. INDEX($G$5:$G$16, that row) gives the SKU.', 'It looks like =INDEX($G$5:$G$16,MATCH(B1,$H$5:$H$16,0)).']
    }),
    shop({
      id: '3.6', from: 'tika', due: '14:00', title: 'Discount with HLOOKUP',
      intro: ['The discount table at the top runs <b>across</b>, not down 😅', 'For the category in <b>B1</b>, what discount % do distributors get?', '@file', 'Answer in <b>B2</b>.'],
      ask: 'The discount % for the category in <b>B1</b>, from the table across the top, in <b>B2</b>.',
      top: (v) => ({ A1: 'Category', B1: v === 'alt' ? 'Sambal' : 'Coffee', A2: 'Discount %' }),
      lesson: {
        idea: 'HLOOKUP is VLOOKUP turned sideways: it searches the top row.',
        teach: '<p><code>HLOOKUP(what, table, row number, FALSE)</code> looks across the first row of the table and returns the row you ask for.</p><p>XLOOKUP does this too, with a row to look in and a row to return.</p>',
        ex: [['=HLOOKUP("Soap",', 'Find Soap'], ['$H$1:$K$2,', 'across the top table,'], ['2,FALSE)', 'and return row 2.']],
        exFormula: '=HLOOKUP("Soap",$H$1:$K$2,2,FALSE)', exLabel: '% discount for Soap',
        bridge: 'Like reading a rate off a table with the headings across the top.'
      },
      starter: '=HLOOKUP(', fnKey: 'HLOOKUP(', chips: ['B1', '$H$1:$K$2', ',2', ',FALSE)'],
      solution: '=HLOOKUP(B1,$H$1:$K$2,2,FALSE)',
      alts: ['=XLOOKUP(B1,$H$1:$K$1,$H$2:$K$2)', '=INDEX($H$2:$K$2,MATCH(B1,$H$1:$K$1,0))'],
      traps: [{ code: 'row', formula: '=HLOOKUP(B1,$H$1:$K$2,1,FALSE)', title: 'Row 1 is the category', plain: 'Row 1 of the table is the category itself. The discount is row 2.', reply: '{value}? That is just the category again 😄 The discount is <b>row 2</b>.' }],
      replies: { pass: ['{value}%, thanks! That goes on the price letter.'], frozen: 'I switched B1 to Sambal and the discount stayed the same. Typed?', typed: 'Typed? The discounts change every quarter 🙏' },
      note: { fn: 'HLOOKUP', what: 'Like VLOOKUP, but searches the top row and returns a row number.', shape: '=HLOOKUP(B1,$H$1:$K$2,2,FALSE)' },
      hints: ['The table runs across: categories in row 1, discounts in row 2.', 'Look for B1 in $H$1:$K$2 and return row 2, exact match.', 'It looks like =HLOOKUP(B1,$H$1:$K$2,2,FALSE).']
    }),
    shop({
      id: '3.7', from: 'bayu', due: '14:45', title: 'Shelf position with XMATCH',
      intro: ['Small one.', 'The price list is in shelf order. Which <b>position</b> is the SKU in <b>B1</b>?', '@file', 'Answer in <b>B2</b>.'],
      ask: 'The position of the SKU in <b>B1</b> in the price list, in <b>B2</b>.',
      top: (v) => ({ A1: 'SKU', B1: v === 'alt' ? 'SRW-SER' : 'SNY-IJO', A2: 'Shelf position' }),
      lesson: {
        idea: 'XMATCH returns where a value sits in a list.',
        teach: '<p><code>XMATCH(what, list)</code> returns 1 for the first item, 2 for the second, and so on. It is exact by default.</p><p>The older <code>MATCH(what, list, 0)</code> does the same. Do not forget the 0.</p>',
        ex: [['=XMATCH("MLK-KAR",', 'Where is this SKU'], [PR('G') + ')', 'in the SKU list?']],
        exFormula: '=XMATCH("MLK-KAR",' + PR('G') + ')', exLabel: 'for MLK-KAR',
        bridge: 'Like the line number of an entry in a register.'
      },
      starter: '=XMATCH(', fnKey: 'XMATCH(', chips: ['B1', PR('G')],
      solution: '=XMATCH(B1,' + PR('G') + ')',
      alts: ['=MATCH(B1,' + PR('G') + ',0)', '=XMATCH(B1,G5:G16)'],
      traps: [{ code: 'namecol', formula: '=XMATCH(B1,' + PR('H') + ')', title: 'That list is names', plain: 'B1 holds a SKU, so look in the SKU column, G.', reply: 'It says #N/A? The SKU is in column <b>G</b>, not the names.' }],
      replies: { pass: ['Shelf {value}. My pickers will love you 📦'], frozen: 'I changed B1 and the position did not move. Typed?', typed: 'Typed? It has to follow B1 😄' },
      note: { fn: 'XMATCH, MATCH', what: 'The position of a value in a list. MATCH needs 0 for exact.', shape: '=XMATCH(B1,$G$5:$G$16)' },
      hints: ['You want the position, not a value.', 'XMATCH looks for B1 in the SKU list $G$5:$G$16.', 'It looks like =XMATCH(B1,$G$5:$G$16).']
    }),
    shop({
      id: '3.8', from: 'anton', due: '15:30', title: 'Pick a measure with CHOOSE',
      intro: ['One cell, three answers.', 'In <b>B1</b> I type 1, 2 or 3. 1 means total cartons, 2 the number of orders, 3 the average cartons per order.', '@file', 'B2 should show whichever I picked.'],
      ask: '<b>B2</b>: 1 = total cartons, 2 = number of orders, 3 = average cartons, by the number in B1.',
      top: (v) => ({ A1: 'Measure (1-3)', B1: v === 'alt' ? 3 : 1, A2: 'Result' }),
      lesson: {
        idea: 'CHOOSE picks one of several answers by number.',
        teach: '<p><code>CHOOSE(n, first, second, third)</code> returns the first answer when n is 1, the second when n is 2, and so on.</p><p>Each answer can be a whole formula.</p>',
        ex: [['=CHOOSE(2,', 'Pick the 2nd of'], ['"Noodles","Coffee","Soap")', 'these three words.']],
        exFormula: '=CHOOSE(2,"Noodles","Coffee","Soap")', exLabel: '',
        bridge: 'Like a report option: the same cell shows a different measure.'
      },
      starter: '=CHOOSE(B1,', fnKey: 'CHOOSE(', chips: ['B1,', 'SUM(' + R('D') + '),', 'COUNT(' + R('D') + '),', 'AVERAGE(' + R('D') + '))'],
      solution: '=CHOOSE(B1,SUM(' + R('D') + '),COUNT(' + R('D') + '),AVERAGE(' + R('D') + '))',
      alts: ['=CHOOSE(B1,SUM(' + R('D') + '),COUNTA(' + R('D') + '),AVERAGE(' + R('D') + '))', '=CHOOSE($B$1,SUM($D$5:$D$24),COUNT($D$5:$D$24),AVERAGE($D$5:$D$24))'],
      traps: [{ code: 'order', formula: '=CHOOSE(B1,COUNT(' + R('D') + '),SUM(' + R('D') + '),AVERAGE(' + R('D') + '))', title: 'The order is mixed up', plain: 'The first answer is for 1, the second for 2. Total cartons comes first.', reply: '{value} for option 1? I think the order is mixed up. <b>1 = total cartons</b>.' }],
      requires: [{ re: /CHOOSE\s*\(/i, title: 'Try CHOOSE', plain: 'This request practices CHOOSE.', reply: 'It works! Sari wants CHOOSE on this one 😉' }],
      replies: { pass: ['Exactly. One cell for my whole dashboard 👍'], frozen: 'I typed 3 into B1 and it still shows the same. It has to follow B1.', typed: 'Typed? It must change when B1 changes.' },
      note: { fn: 'CHOOSE', what: 'CHOOSE(n, answer 1, answer 2, …) returns the answer at position n.', shape: '=CHOOSE(B1,SUM(D5:D24),COUNT(D5:D24),AVERAGE(D5:D24))' },
      hints: ['B1 says which answer. CHOOSE lists the answers in order.', 'The three answers are SUM, COUNT and AVERAGE of D5:D24.', 'It looks like =CHOOSE(B1,SUM(D5:D24),COUNT(D5:D24),AVERAGE(D5:D24)).']
    })
  );
  LIST.push(Object.assign(shop({
    id: '3.9', from: 'tika', due: '17:00', title: 'Lookups in real Excel', ansLabel: 'Product',
    intro: ['Last one this chapter, in <b>real Excel</b> 💻', 'Download the orders, fill column E with each order\'s <b>product name</b>, and send it back.', '@file', 'VLOOKUP, XLOOKUP or INDEX/MATCH, your choice.'],
    ask: 'In Excel: product name for every order in column E, then upload.',
    lesson: {
      idea: 'Lookups, in real Excel.',
      teach: '<p>Tap <b>Download for Excel</b>, type your lookup in E5, double-click the fill handle, save and upload.</p><p>Excel 365 suggests the table range as you drag. Check that the $ signs are there before you copy down.</p>',
      bridge: 'The same job you will do with any master file at work.'
    },
    starter: '=XLOOKUP(', fnKey: 'XLOOKUP(', chips: [],
    solution: '=XLOOKUP(C5,' + PR('G') + ',' + PR('H') + ')', alts: [],
    rowReply: 'Row {row} says {got}, but SKU {key} is {want} on the price list.', keyCol: 'C',
    replies: { pass: ['Received, every name is right. Makasih! 🙏', 'That is Chapter 3. You can read a master file now.'], wrong: '', typed: 'The names are typed in, not formulas. When the orders change, they will be wrong 😅' },
    after: 'Chapter 4 is text and dates: LEFT, MID, TRIM, TEXT, EOMONTH, NETWORKDAYS and aging by days. It arrives in the next build.',
    note: { fn: 'Lookups in real Excel', what: 'One lookup in E5, the fill handle, save, upload.', shape: '=XLOOKUP(C5,$G$5:$G$16,$H$5:$H$16)' }
  }), { kind: 'excel' }));

  /* ---------- Chapter 4: Text and dates ----------
     Rasa Kenari's invoice register for October, as it comes out of the old billing system: reference codes
     with a branch at the end, distributor names typed by hand, invoice dates and payment terms. */
  const CITY = { Bekasi: 'BKS', Tangerang: 'TNG', Bandung: 'BDG', Cirebon: 'CRB', Semarang: 'SMG' };
  const DIST = Object.fromEntries(D.DISTRIBUTORS.map((d) => [d.name, d]));
  const messy = (name, i) => [name, name.toLowerCase().replace(' ', '  '), '  ' + name.toUpperCase() + ' ', name + '  '][i % 4];
  const ASOF = D.toSerial('2026-11-20');
  const ICOLS = ['A', 'B', 'C', 'D', 'E', 'F'];
  const INV_FILE = 'register_faktur_oktober.xlsx';
  function invSheet(top, ans, alt) {
    const c = { A4: 'Reference', B4: 'Distributor (as typed)', C4: 'Invoice date', D4: 'Terms (days)', E4: 'Amount (Rp)' };
    if (ans) c.F4 = ans;
    rows.forEach((r, i) => {
      const n = FIRST + i, d = DIST[r.distributor];
      c['A' + n] = r.invoice + '-' + CITY[d.city];
      c['B' + n] = messy(r.distributor, i);
      c['C' + n] = D.toSerial(r.date);
      c['D' + n] = d.terms_days;
      c['E' + n] = r.net;
    });
    if (alt) {
      c['A' + (FIRST + 1)] = 'INV-2207-CRB';
      c['B' + (FIRST + 2)] = ' ud  seriti   jaya ';
      c['C' + (FIRST + 3)] = c['C' + (FIRST + 3)] - 9;
      c['D' + (FIRST + 4)] = 60;
      c['E' + (FIRST + 5)] = c['E' + (FIRST + 5)] + 1234500;
    }
    return Object.assign(c, top);
  }
  function inv(o) {
    const top = typeof o.top === 'function' ? o.top : () => (o.top || {});
    const base = invSheet(top('base'), o.ansLabel, false), alt = invSheet(top('alt'), o.ansLabel, true);
    return formula(Object.assign({
      chapter: 4, file: INV_FILE, cols: ICOLS, num: ['D', 'E'], dateCols: ['C'], dateCells: ['B1'],
      first: FIRST, last: LAST, fill: { col: 'F', first: FIRST, last: LAST }, answerCell: 'F' + FIRST,
      colRange: Object.fromEntries(ICOLS.map((c) => [c, R(c)])),
      variantCells: { base, alt }, cells: base
    }, o, { top: undefined }));
  }
  const asOf = (v) => ({ A1: 'As of', B1: v === 'alt' ? ASOF + 1 : ASOF });

  LIST.push(
    inv({
      id: '4.1', from: 'tika', due: '08:30', title: 'Branch code with RIGHT', ansLabel: 'Branch',
      intro: ['Pagi {name}! The old billing system is messy 😅', 'Every invoice reference ends with a 3-letter branch code, like <b>BKS</b> for Bekasi.', '@file', 'Pull the <b>branch code</b> into column F, every row.'],
      ask: 'The last 3 letters of each reference, in column F.',
      lesson: {
        idea: 'RIGHT takes characters from the end of a text. LEFT takes them from the start.',
        teach: '<p><code>RIGHT(A5, 3)</code> gives the last 3 characters. <code>LEFT(A5, 3)</code> gives the first 3.</p><p>Text functions work on anything typed as text, including codes made of numbers.</p>',
        ex: [['=LEFT(A5,', 'From the start of the reference,'], ['3)', 'take 3 characters.']],
        exFormula: '=LEFT(A5,3)', exLabel: '',
        bridge: 'Like reading the branch from the last part of a voucher number.'
      },
      starter: '=RIGHT(', fnKey: 'RIGHT(', chips: ['A5', ',3)'],
      solution: '=RIGHT(A5,3)', alts: ['=MID(A5,LEN(A5)-2,3)'], exactCase: true,
      traps: [{ code: 'left', formula: '=LEFT(A5,3)', title: 'That is the start', plain: 'LEFT reads from the start. The branch is at the end, which is RIGHT.', reply: '{value}? That is the start of every reference 😄 The branch is at the <b>end</b>.' }],
      replies: { pass: ['Mantap. Now I can sort by branch 🙏'], frozen: 'One reference was corrected and your branch did not change.', typed: 'Typed codes will not follow the references 😅' },
      note: { fn: 'LEFT, RIGHT', what: 'Characters from the start, or the end, of a text.', shape: '=RIGHT(A5,3)' },
      hints: ['The branch is the last 3 characters.', 'RIGHT reads from the end of A5.', 'It looks like =RIGHT(A5,3).']
    }),
    inv({
      id: '4.2', from: 'tika', due: '09:15', title: 'Invoice number with MID', ansLabel: 'Number',
      intro: ['Next.', 'The 4-digit invoice number sits in the middle: <b>INV-<u>1023</u>-BKS</b>.', '@file', 'Just the 4 digits in column F please.'],
      ask: 'The 4 digits in the middle of each reference, in column F.',
      lesson: {
        idea: 'MID takes characters from the middle, from a start position.',
        teach: '<p><code>MID(A5, 5, 4)</code>: start at character 5, take 4. Count every character: I, N, V and the dash are 1 to 4.</p><p>The result is text. Wrap it in VALUE if you need a number.</p>',
        ex: [['=MID(A5,', 'From the reference,'], ['1,3)', 'start at 1, take 3.']],
        exFormula: '=MID(A5,1,3)', exLabel: '',
        bridge: 'Like reading digits 5 to 8 of an account number.'
      },
      starter: '=MID(', fnKey: 'MID(', chips: ['A5', ',5', ',4)'],
      solution: '=MID(A5,5,4)', alts: ['=LEFT(RIGHT(A5,8),4)'], exactCase: true,
      traps: [{ code: 'offbyone', formula: '=MID(A5,4,4)', title: 'One character early', plain: 'Position 4 is the dash. The digits start at 5.', reply: '{value}? There is a dash in front 😅 Start one character <b>later</b>.' }],
      replies: { pass: ['Perfect, clean numbers 👍'], frozen: 'A reference was fixed this morning and your number stayed the same.', typed: 'Typed? It has to follow the references.' },
      note: { fn: 'MID', what: 'MID(text, start, how many). Count every character, dashes too.', shape: '=MID(A5,5,4)' },
      hints: ['Start after INV- and take 4 characters.', 'INV- is 4 characters, so start at 5.', 'It looks like =MID(A5,5,4).']
    }),
    inv({
      id: '4.3', from: 'bayu', due: '10:00', title: 'Clean names with TRIM and UPPER', ansLabel: 'Clean name',
      intro: ['Help 😩', 'The distributor names were typed by hand: extra spaces, small letters, big letters.', '@file', 'Column F: each name with <b>no extra spaces</b>, in <b>CAPITALS</b>.'],
      ask: 'Each name trimmed and in capitals, in column F.',
      lesson: {
        idea: 'TRIM removes extra spaces. UPPER turns text into capitals.',
        teach: '<p><code>TRIM(B5)</code> removes spaces at both ends and squeezes double spaces inside to one. <code>UPPER</code>, <code>LOWER</code> and <code>PROPER</code> change the capitals.</p><p>Functions can sit inside each other: <code>UPPER(TRIM(B5))</code>.</p>',
        ex: [['=LOWER(', 'Small letters'], ['TRIM(B5))', 'after removing the extra spaces.']],
        exFormula: '=LOWER(TRIM(B5))', exLabel: '',
        bridge: 'Like cleaning a customer master before you reconcile against it.'
      },
      starter: '=UPPER(', fnKey: 'UPPER(TRIM(', chips: ['B5', '))'],
      solution: '=UPPER(TRIM(B5))', alts: ['=TRIM(UPPER(B5))'], exactCase: true,
      traps: [
        { code: 'notrim', formula: '=UPPER(B5)', title: 'The spaces are still there', plain: 'UPPER changes the letters only. Wrap the name in TRIM first.', reply: 'Capitals, good! But the extra spaces are still there 😅 Look at row 6.' },
        { code: 'noupper', formula: '=TRIM(B5)', title: 'Not in capitals', plain: 'TRIM fixed the spaces. Now wrap it in UPPER.', reply: 'Spaces are gone 👍 but some names are still small letters. I need <b>CAPITALS</b>.' }
      ],
      replies: { pass: ['Beautiful. Now my lookups will actually match 🙌'], frozen: 'One name was retyped and your clean name did not follow.', typed: 'Typed names will not follow 😅' },
      note: { fn: 'TRIM, UPPER, LOWER, PROPER', what: 'TRIM removes extra spaces. UPPER, LOWER and PROPER change the capitals.', shape: '=UPPER(TRIM(B5))' },
      hints: ['Two jobs: remove spaces, make capitals.', 'TRIM(B5) removes the spaces. UPPER(...) makes capitals.', 'It looks like =UPPER(TRIM(B5)).']
    }),
    inv({
      id: '4.4', from: 'tika', due: '10:45', title: 'Amount as text with TEXT', ansLabel: 'For the letter',
      intro: ['For the reminder letters:', 'I need each amount written like <b>Rp 4,351,200</b>, with the comma groups.', '@file', 'Column F, every row.'],
      ask: '"Rp " and the amount with thousands commas, in column F.',
      lesson: {
        idea: 'TEXT turns a number into text with the format you choose.',
        teach: '<p><code>TEXT(E5, "#,##0")</code> writes 4351200 as 4,351,200. Join it to other text with <code>&amp;</code>.</p><p><b>On Indonesian settings</b> Excel may need <code>"#.##0"</code> instead, because the separators swap. Kicau uses the English codes.</p>',
        ex: [['="Total: "&', 'The word Total, then'], ['TEXT(E5,"0")', 'the amount as plain text.']],
        exFormula: '="Total: "&TEXT(E5,"0")', exLabel: '',
        bridge: 'Like the amount line on a printed statement.'
      },
      starter: '="Rp "&', fnKey: 'TEXT(', chips: ['E5', ',"#,##0")', '&'],
      solution: '="Rp "&TEXT(E5,"#,##0")', alts: ['=CONCAT("Rp ",TEXT(E5,"#,##0"))'], exactCase: true,
      traps: [{ code: 'notext', formula: '="Rp "&E5', title: 'No commas', plain: 'Joining a number gives the raw digits. TEXT adds the commas.', reply: 'It says {value}, no commas 😅 Use TEXT with "#,##0".' }],
      replies: { pass: ['Perfect, the letters look professional now ✉️'], frozen: 'An amount was corrected and your text did not follow.', typed: 'Typed amounts will not follow 🙏' },
      note: { fn: 'TEXT, &', what: 'TEXT(number, format) writes a number as text. & joins texts.', shape: '="Rp "&TEXT(E5,"#,##0")' },
      hints: ['"Rp " joined with the amount, formatted.', 'TEXT(E5,"#,##0") gives the commas. Join with &.', 'It looks like ="Rp "&TEXT(E5,"#,##0").']
    }),
    inv({
      id: '4.5', from: 'citra', due: '11:30', title: 'Payment reference by joining', ansLabel: 'Payment ref',
      intro: ['Citra here. For our virtual accounts:', 'Each invoice needs a payment reference like <b>KNR-1023-BKS</b>: KNR, the invoice number, the branch, with dashes.', '@file'],
      ask: 'KNR-number-branch for every invoice, in column F.',
      lesson: {
        idea: 'Join pieces of text with &, CONCAT or TEXTJOIN.',
        teach: '<p><code>"KNR-"&amp;MID(A5,5,4)</code> joins the pieces. <code>CONCAT</code> does the same with commas between the parts.</p><p><code>TEXTJOIN("-", TRUE, …)</code> puts a dash between every part for you.</p>',
        ex: [['=TEXTJOIN("/",TRUE,', 'Join with slashes:'], ['RIGHT(A5,3),D5)', 'the branch and the terms.']],
        exFormula: '=TEXTJOIN("/",TRUE,RIGHT(A5,3),D5)', exLabel: '',
        bridge: 'Like building a journal reference from its parts.'
      },
      starter: '="KNR-"&', fnKey: 'MID(A5,5,4)', chips: ['&"-"&', 'RIGHT(A5,3)', 'TEXTJOIN('],
      solution: '="KNR-"&MID(A5,5,4)&"-"&RIGHT(A5,3)', exactCase: true,
      alts: ['=CONCAT("KNR-",MID(A5,5,4),"-",RIGHT(A5,3))', '=TEXTJOIN("-",TRUE,"KNR",MID(A5,5,4),RIGHT(A5,3))'],
      traps: [{ code: 'nodash', formula: '="KNR"&MID(A5,5,4)&RIGHT(A5,3)', title: 'The dashes are missing', plain: 'Add "-" between the parts.', reply: '{value}? The bank system needs the <b>dashes</b> 😅' }],
      replies: { pass: ['Perfect, the virtual accounts will match 🏦'], frozen: 'A reference changed and your payment ref did not follow.', typed: 'Typed? It has to follow the references 🙏' },
      note: { fn: '&, CONCAT, TEXTJOIN', what: 'Join texts. TEXTJOIN puts a separator between every part.', shape: '="KNR-"&MID(A5,5,4)&"-"&RIGHT(A5,3)' },
      hints: ['Three parts with dashes: KNR, the number, the branch.', 'The number is MID(A5,5,4), the branch RIGHT(A5,3). Join with &.', 'It looks like ="KNR-"&MID(A5,5,4)&"-"&RIGHT(A5,3).']
    }),
    inv({
      id: '4.6', from: 'lestari', due: '13:00', title: 'Due date by adding days', ansLabel: 'Due date', ansDate: true,
      intro: ['Hai {name}, Bu Lestari.', 'Each invoice is due <b>Terms days</b> after the invoice date.', '@file', 'The due date in column F, every row.'],
      ask: 'Invoice date + terms, every row, in column F.',
      lesson: {
        idea: 'A date is a number of days, so you can add days to it.',
        teach: '<p>Excel stores 31 Oct 2026 as a day number. <code>C5+D5</code> adds the terms and gives the due date.</p><p>If you see a number like 46000 instead of a date, format the cell as a date. The value is right.</p>',
        ex: [['=C5', 'The invoice date'], ['+7', 'plus 7 days.']],
        exFormula: '=C5+7', exLabel: '', exDate: true,
        bridge: 'Like a due-date column on an aging schedule.'
      },
      starter: '=C5+', fnKey: 'D5', chips: ['C5', '+'],
      solution: '=C5+D5', alts: ['=D5+C5'],
      traps: [{ code: 'fixed', formula: '=C5+30', title: 'Not every term is 30 days', plain: 'Some distributors get 45 days. Use the terms in column D.', reply: 'Some distributors have 45 days, not 30 😅 Use the <b>Terms</b> column.' }],
      replies: { pass: ['Thank you. Now I know when to chase 📅'], frozen: 'A distributor got 60 days terms this morning and the due date did not move.', typed: 'Typed dates will not follow 🙏' },
      note: { fn: 'Date + days', what: 'Dates are day numbers, so date + days gives a later date.', shape: '=C5+D5' },
      hints: ['Due date = invoice date + terms.', 'The date is C5 and the terms are D5.', 'It looks like =C5+D5.']
    }),
    inv({
      id: '4.7', from: 'lestari', due: '13:45', title: 'Month end with EOMONTH', ansLabel: 'Month end', ansDate: true,
      intro: ['For the monthly close:', 'Which <b>month-end</b> does each invoice belong to? The last day of the invoice\'s month.', '@file'],
      ask: 'The last day of each invoice\'s month, in column F.',
      lesson: {
        idea: 'EOMONTH gives the last day of a month.',
        teach: '<p><code>EOMONTH(C5, 0)</code> is the last day of C5\'s month. <code>EOMONTH(C5, 1)</code> is the end of next month, and -1 the end of last month.</p><p>Banks and finance teams use it for every month-end report.</p>',
        ex: [['=EOMONTH(C5,', 'From the invoice date,'], ['-1)', 'the end of the month before.']],
        exFormula: '=EOMONTH(C5,-1)', exLabel: '', exDate: true,
        bridge: 'Like the cut-off date of a monthly close.'
      },
      starter: '=EOMONTH(', fnKey: 'EOMONTH(', chips: ['C5', ',0)'],
      solution: '=EOMONTH(C5,0)', alts: [],
      traps: [{ code: 'nextmonth', formula: '=EOMONTH(C5,1)', title: 'That is next month', plain: '1 means one month later. This month is 0.', reply: '{value}? That is next month\'s end 😄 This month is <b>0</b>.' }],
      replies: { pass: ['Perfect, every invoice sits in its month 📆'], frozen: 'A date was corrected and your month end did not move.', typed: 'Typed dates will not follow 🙏' },
      note: { fn: 'EOMONTH', what: 'EOMONTH(date, months): the last day of the month, months later or earlier.', shape: '=EOMONTH(C5,0)' },
      hints: ['You want the last day of the same month.', 'EOMONTH with 0 months.', 'It looks like =EOMONTH(C5,0).']
    }),
    inv({
      id: '4.8', from: 'citra', due: '14:30', title: 'Days past due', ansLabel: 'Days late', top: asOf,
      intro: ['This is the one collection lives on 💪', 'As of the date in <b>B1</b>, how many days is each invoice <b>past its due date</b>? If it is not due yet, show <b>0</b>.', '@file'],
      ask: 'Days past due as of <b>B1</b>, 0 if not due yet, every row.',
      lesson: {
        idea: 'Subtract dates to count days. MAX(0, …) stops negatives.',
        teach: '<p>Days late = as-of date − due date, and the due date is <code>C5+D5</code>. Pin B1 with $ for copying.</p><p>Invoices not due yet give a negative number. <code>MAX(0, …)</code> turns those into 0.</p>',
        ex: [['=$B$1', 'The as-of date'], ['-C5', 'minus the invoice date: the invoice\'s age.']],
        exFormula: '=$B$1-C5', exLabel: 'days old for row 5',
        bridge: 'Exactly the aging buckets on a receivables report.'
      },
      starter: '=MAX(0,', fnKey: '$B$1', chips: ['-(C5+D5)', ')'],
      solution: '=MAX(0,$B$1-(C5+D5))', alts: ['=MAX(0,$B$1-C5-D5)', '=IF($B$1>C5+D5,$B$1-C5-D5,0)'],
      traps: [
        { code: 'negative', formula: '=$B$1-(C5+D5)', title: 'Negative for invoices not due', plain: 'An invoice not due yet comes out negative. Wrap it in MAX(0, …).', reply: 'Some say {value}, negative days late? 😄 Those are not due yet. Show <b>0</b> with MAX.' },
        { code: 'nodollar', formula: '=MAX(0,B1-(C5+D5))', title: 'B1 slid down the column', plain: 'Copied down, B1 became B2, B3 and so on. Pin it as $B$1.', reply: 'Only the first invoice has a number 🤔 B1 slides when copied. Pin it with <b>$B$1</b>.' }
      ],
      replies: { pass: ['That is my aging list. Terima kasih! 💪'], frozen: 'B1 moved to tomorrow and the days did not change. Did you type the date in? Point at <b>$B$1</b>.', typed: 'Typed days will be wrong tomorrow 😅' },
      note: { fn: 'Days between dates', what: 'Later date − earlier date = days. MAX(0, …) stops negatives.', shape: '=MAX(0,$B$1-(C5+D5))' },
      hints: ['Days late = as-of date − due date, never below 0.', 'The due date is C5+D5. The as-of date is $B$1.', 'It looks like =MAX(0,$B$1-(C5+D5)).']
    }),
    inv({
      id: '4.9', from: 'citra', due: '15:15', title: 'Working days with NETWORKDAYS', ansLabel: 'Working days', top: asOf,
      intro: ['One more for collection.', 'How many <b>working days</b>, Monday to Friday, from each invoice date up to the date in <b>B1</b>?', '@file'],
      ask: 'Working days from the invoice date to <b>B1</b>, every row.',
      lesson: {
        idea: 'NETWORKDAYS counts Monday to Friday between two dates.',
        teach: '<p><code>NETWORKDAYS(start, end)</code> counts working days, counting both ends. A third part can list holidays.</p><p>Plain subtraction counts calendar days, weekends included.</p>',
        ex: [['=NETWORKDAYS(C5,', 'From the invoice date'], ['C5+14)', 'to two weeks later.']],
        exFormula: '=NETWORKDAYS(C5,C5+14)', exLabel: 'working days in two weeks',
        bridge: 'Like counting business days for a payment SLA.'
      },
      starter: '=NETWORKDAYS(', fnKey: 'NETWORKDAYS(', chips: ['C5', ',$B$1)'],
      solution: '=NETWORKDAYS(C5,$B$1)', alts: ['=NETWORKDAYS(C5,B$1)'],
      traps: [{ code: 'calendar', formula: '=$B$1-C5', title: 'Those are calendar days', plain: 'Subtraction counts weekends too. NETWORKDAYS skips them.', reply: '{value} days includes weekends 😄 I need <b>working</b> days.' }],
      replies: { pass: ['Exactly. That goes into the SLA report ✅'], frozen: 'B1 moved and your days did not. Point at <b>$B$1</b>.', typed: 'Typed? It must follow B1 🙏' },
      note: { fn: 'NETWORKDAYS', what: 'Working days between two dates, both ends counted. Add holidays as a third part.', shape: '=NETWORKDAYS(C5,$B$1)' },
      hints: ['Working days, not calendar days.', 'NETWORKDAYS from C5 to $B$1.', 'It looks like =NETWORKDAYS(C5,$B$1).']
    }),
    inv({
      id: '4.10', from: 'lestari', due: '16:00', title: 'Review date with EDATE', ansLabel: 'Review date', ansDate: true,
      intro: ['Last formula today 🙏', 'Every distributor gets a credit review <b>3 months</b> after the invoice date. Same day of the month.', '@file'],
      ask: 'The invoice date plus 3 months, every row.',
      lesson: {
        idea: 'EDATE moves a date by whole months.',
        teach: '<p><code>EDATE(C5, 3)</code> is the same day, 3 months later. Months have different lengths, so adding 90 days drifts.</p><p>Negative months go back in time.</p>',
        ex: [['=EDATE(C5,', 'From the invoice date,'], ['-1)', 'one month earlier.']],
        exFormula: '=EDATE(C5,-1)', exLabel: '', exDate: true,
        bridge: 'Like a quarterly review cycle.'
      },
      starter: '=EDATE(', fnKey: 'EDATE(', chips: ['C5', ',3)'],
      solution: '=EDATE(C5,3)', alts: [],
      traps: [{ code: 'days90', formula: '=C5+90', title: '90 days is not 3 months', plain: 'Months are 28 to 31 days long. EDATE keeps the same day of the month.', reply: '{value}? That drifted off the date 😅 3 <b>months</b>, not 90 days. EDATE does it.' }],
      replies: { pass: ['Perfect, the review calendar is ready 📅'], frozen: 'A date changed and your review date did not.', typed: 'Typed dates will not follow 🙏' },
      note: { fn: 'EDATE', what: 'EDATE(date, months): the same day of the month, months later.', shape: '=EDATE(C5,3)' },
      hints: ['Move the date by 3 whole months.', 'EDATE from C5, 3 months.', 'It looks like =EDATE(C5,3).']
    })
  );
  LIST.push(Object.assign(inv({
    id: '4.11', from: 'citra', due: '17:00', title: 'Aging in real Excel', ansLabel: 'Days late', top: asOf,
    intro: ['Last one, in <b>real Excel</b> 💻', 'Download the register, work out days past due as of <b>B1</b> in column F, 0 when not due yet, and send it back.', '@file'],
    ask: 'In Excel: days past due as of B1, every row, then upload.',
    lesson: {
      idea: 'Dates in real Excel.',
      teach: '<p>Download, type your formula in F5, copy it down, save, upload.</p><p>If Excel shows a date instead of a number in column F, change the format to Number. The value is still right.</p>',
      bridge: 'The aging you will do every month-end.'
    },
    starter: '=MAX(0,', fnKey: '', chips: [],
    solution: '=MAX(0,$B$1-(C5+D5))', alts: [],
    rowReply: 'Row {row} says {got}, but invoice {key} is {want} days past due as of B1.', keyCol: 'A',
    replies: { pass: ['Received, every invoice aged right. Makasih! 🙏', 'That is Chapter 4. You can clean text and work with dates now.'], wrong: '', typed: 'The days are typed in, not formulas. Tomorrow they will be wrong 😅' },
    after: 'Chapter 5 is dynamic arrays: FILTER, SORT, UNIQUE, SEQUENCE and LET. It arrives in the next build.',
    note: { fn: 'Aging in real Excel', what: 'One formula, copied down, as of a date you can change.', shape: '=MAX(0,$B$1-(C5+D5))' }
  }), { kind: 'excel' }));

  /* ---------- Chapter 5: Dynamic arrays ----------
     October's sales again, with room on the right. One formula in G5 spills its answer down and across. */
  const ACOLS = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'];
  function arrSheet(top, head, alt) {
    const c = { A4: 'Date', B4: 'Distributor', C4: 'Product', D4: 'Cartons', E4: 'Net (Rp)' };
    (head || []).forEach((h, i) => { c['GHI'[i] + '4'] = h; });
    rows.forEach((r, i) => { const n = FIRST + i; c['A' + n] = D.toSerial(r.date); c['B' + n] = r.distributor; c['C' + n] = r.product; c['D' + n] = r.cartons; c['E' + n] = r.net; });
    if (alt) { c['B' + (FIRST + 1)] = 'UD Seriti Jaya'; c['E' + (FIRST + 2)] += 2500000; c['D' + (FIRST + 3)] += 15; c['C' + (FIRST + 6)] = 'Mi Lekas Goreng'; }
    return Object.assign(c, top);
  }
  function arr(o) {
    const top = typeof o.top === 'function' ? o.top : () => (o.top || {});
    const spill = o.spillSize ? { anchor: 'G' + FIRST, cols: o.spillSize[0], rows: o.spillSize[1] } : null;
    const base = arrSheet(top('base'), o.heads, false), alt = arrSheet(top('alt'), o.heads, true);
    return formula(Object.assign({
      chapter: 5, file: 'penjualan_oktober_v2.xlsx', cols: ACOLS, num: ['D', 'E'], dateCols: ['A'],
      first: FIRST, last: LAST, spill, answerCell: spill ? 'G' + FIRST : 'B2',
      colRange: Object.fromEntries(ACOLS.map((c) => [c, R(c)])),
      variantCells: { base, alt }, cells: base
    }, o, { top: undefined }));
  }
  const pick = (v) => ({ A1: 'Distributor', B1: v === 'alt' ? other : target });

  LIST.push(
    arr({
      id: '5.1', from: 'dimas', due: '09:00', title: 'Every distributor once, with UNIQUE', heads: ['Distributor'], spillSize: [1, 20],
      intro: ['Pagi {name}! Dimas 👋', 'Which distributors appear in October\'s sales? Each name <b>once</b>.', '@file', 'One formula in <b>G5</b>, and let it spill.'],
      ask: 'Each distributor once, from one formula in <b>G5</b>.',
      lesson: {
        idea: 'Some formulas return many values, and they spill into the cells below.',
        teach: '<p><code>UNIQUE(B5:B24)</code> returns each name once. Type it in one cell and the list spills down by itself.</p><p>If something blocks the cells it needs, Excel shows <code>#SPILL!</code>. Clear the way and it comes back.</p>',
        ex: [['=UNIQUE(', 'Each value once'], [R('C') + ')', 'from the Product column.']],
        exFormula: '=ROWS(UNIQUE(' + R('C') + '))', exLabel: 'different products',
        bridge: 'Like the list of accounts in a trial balance: each one once.'
      },
      starter: '=UNIQUE(', fnKey: 'UNIQUE(', chips: [R('B'), ')'],
      solution: '=UNIQUE(' + R('B') + ')', alts: ['=UNIQUE($B$5:$B$24)'],
      traps: [{ code: 'dupes', formula: '=' + R('B'), title: 'Every row, duplicates and all', plain: 'That copies the whole column. UNIQUE keeps each name once.', reply: 'Some names come up three or four times 😅 I need each one <b>once</b>.' }],
      requires: [{ re: /UNIQUE\s*\(/i, title: 'Use UNIQUE', plain: 'This request practices UNIQUE.', reply: 'Right list! Sari wants UNIQUE on this one 😉' }],
      replies: { pass: ['That is the list, from one formula 😮'], frozen: 'A distributor changed this morning and your list did not.', typed: 'Typed names will not follow the sales 🙏' },
      note: { fn: 'UNIQUE, spilling', what: 'UNIQUE(range) returns each value once. One formula spills into the cells below.', shape: '=UNIQUE(B5:B24)' },
      hints: ['One function returns each value once.', 'UNIQUE on the Distributor column, B5:B24.', 'It looks like =UNIQUE(B5:B24).']
    }),
    arr({
      id: '5.2', from: 'anton', due: '09:45', title: 'Sorted with SORT', heads: ['Distributor'], spillSize: [1, 20],
      intro: ['The same list, but in <b>A to Z</b> order please.', '@file', 'One formula in G5.'],
      ask: 'Each distributor once, A to Z, from one formula in <b>G5</b>.',
      lesson: {
        idea: 'SORT puts a list in order. Formulas can go inside each other.',
        teach: '<p><code>SORT(list)</code> sorts A to Z, or smallest first. <code>SORT(list, 1, -1)</code> sorts the other way.</p><p><code>SORT(UNIQUE(B5:B24))</code> first removes duplicates, then sorts.</p>',
        ex: [['=SORT(', 'Sort'], [R('D') + ',1,-1)', 'the cartons, biggest first.']],
        exFormula: '=INDEX(SORT(' + R('D') + ',1,-1),1)', exLabel: 'is the biggest order, the top of that sorted list',
        bridge: 'Like sorting a ledger by account name.'
      },
      starter: '=SORT(', fnKey: 'SORT(', chips: ['UNIQUE(', R('B'), '))'],
      solution: '=SORT(UNIQUE(' + R('B') + '))', alts: ['=UNIQUE(SORT(' + R('B') + '))', '=SORT(UNIQUE(' + R('B') + '),1,1)'],
      traps: [{ code: 'unsorted', formula: '=UNIQUE(' + R('B') + ')', title: 'Not sorted yet', plain: 'Wrap UNIQUE in SORT.', reply: 'The names are right, but not in A to Z order 😄 Wrap it in <b>SORT</b>.' }],
      requires: [{ re: /SORT\s*\(/i, title: 'Use SORT', plain: 'This request practices SORT.', reply: 'Right order! Use SORT though 😉' }],
      replies: { pass: ['Clean, A to Z 👍'], frozen: 'The sales changed and your list stayed the same.', typed: 'Typed? It has to follow the sales.' },
      note: { fn: 'SORT', what: 'SORT(list, column, 1 or -1). Nest it: SORT(UNIQUE(…)).', shape: '=SORT(UNIQUE(B5:B24))' },
      hints: ['Remove duplicates, then sort.', 'SORT around UNIQUE(B5:B24).', 'It looks like =SORT(UNIQUE(B5:B24)).']
    }),
    arr({
      id: '5.3', from: 'tika', due: '10:30', title: 'One distributor\'s rows with FILTER', heads: ['Product', 'Cartons', 'Net (Rp)'], spillSize: [3, 20], top: pick,
      intro: ['I need every sale for the distributor in <b>B1</b>: product, cartons and net.', '@file', 'One formula in G5. I will change B1 later.'],
      ask: 'Product, cartons and net for the distributor in <b>B1</b>, from one formula in <b>G5</b>.',
      lesson: {
        idea: 'FILTER keeps only the rows that pass a test.',
        teach: '<p><code>FILTER(C5:E24, B5:B24=B1)</code> returns columns C to E, only for rows where the distributor is B1.</p><p>It is like an AutoFilter that lives in a formula and updates by itself.</p>',
        ex: [['=FILTER(' + R('C') + ',', 'The products'], [R('D') + '>40)', 'of orders over 40 cartons.']],
        exFormula: '=ROWS(FILTER(' + R('C') + ',' + R('D') + '>40))', exLabel: 'orders over 40 cartons',
        bridge: 'Like filtering a ledger to one account, but the filter is a formula.'
      },
      starter: '=FILTER(', fnKey: 'FILTER(', chips: ['C5:E24', ',B5:B24=B1', ')'],
      solution: '=FILTER(C5:E24,B5:B24=B1)', alts: ['=FILTER($C$5:$E$24,$B$5:$B$24=$B$1)'],
      frozenTry: '=FILTER(C5:E24,B5:B24="' + target + '")',
      traps: [{ code: 'wide', formula: '=FILTER(A5:E24,B5:B24=B1)', title: 'Too many columns', plain: 'Start from column C: product, cartons and net.', reply: 'It starts with the date and the name 😅 I only need <b>product, cartons and net</b>, columns C to E.' }],
      replies: { pass: ['Exactly. And it changes when I change B1 🙌'], frozen: 'I changed B1 to {other} and still see the same rows. Did you type the name? Use <b>B5:B24=B1</b>.', typed: 'Typed rows will not follow 🙏' },
      note: { fn: 'FILTER', what: 'FILTER(what to return, test). Returns only the rows that pass.', shape: '=FILTER(C5:E24,B5:B24=B1)' },
      hints: ['Return columns C to E, only where the distributor is B1.', 'The test is B5:B24=B1.', 'It looks like =FILTER(C5:E24,B5:B24=B1).']
    }),
    arr({
      id: '5.4', from: 'dimas', due: '11:15', title: 'Two conditions, and None', heads: ['Product'], spillSize: [1, 20],
      top: (v) => ({ A1: 'Distributor', B1: target, A2: 'Min cartons', B2: v === 'alt' ? 60 : 20 }),
      intro: ['A stricter list:', 'Products sold to the distributor in <b>B1</b> with <b>at least B2 cartons</b>.', '@file', 'If nothing matches, show <b>None</b> instead of an error.'],
      ask: 'Products for B1 with cartons ≥ B2, or <b>None</b>, from G5.',
      lesson: {
        idea: 'Multiply tests to mean AND. A third part shows when nothing matches.',
        teach: '<p><code>(B5:B24=B1)*(D5:D24&gt;=B2)</code> is 1 only where both are true. That is AND inside FILTER.</p><p><code>FILTER(…, …, "None")</code> shows None instead of <code>#CALC!</code> when no row passes.</p>',
        ex: [['=FILTER(' + R('C') + ',', 'Products where'], ['(' + R('D') + '>30)*(' + R('E') + '>5000000),', 'both tests pass,'], ['"None")', 'or say None.']],
        exFormula: '=INDEX(FILTER(' + R('C') + ',(' + R('D') + '>30)*(' + R('E') + '>5000000),"None"),1)', exLabel: 'is the first of them',
        bridge: 'Like a two-field filter on an exceptions report.'
      },
      starter: '=FILTER(', fnKey: 'FILTER(', chips: ['C5:C24', ',(B5:B24=B1)*(D5:D24>=B2)', ',"None")'],
      solution: '=FILTER(C5:C24,(B5:B24=B1)*(D5:D24>=B2),"None")', alts: ['=FILTER(C5:C24,(D5:D24>=B2)*(B5:B24=B1),"None")'],
      traps: [{ code: 'onecond', formula: '=FILTER(C5:C24,B5:B24=B1,"None")', title: 'Only one condition', plain: 'Add the cartons test with *.', reply: 'That is every product for them, small orders too. Only orders of at least <b>B2</b> cartons.' }],
      replies: { pass: ['Mantap. And None when there is nothing 👌'], frozen: 'I raised B2 to 60, nothing matched, and it broke 😅 Add <b>"None"</b> as the third part, and point at B1 and B2.', typed: 'Typed? It has to follow B1 and B2.' },
      note: { fn: 'FILTER with AND, if_empty', what: 'Multiply tests for AND. The third part shows when nothing matches.', shape: '=FILTER(C5:C24,(B5:B24=B1)*(D5:D24>=B2),"None")' },
      hints: ['Two tests together, and something to show when empty.', 'Multiply (B5:B24=B1)*(D5:D24>=B2). Third part "None".', 'It looks like =FILTER(C5:C24,(B5:B24=B1)*(D5:D24>=B2),"None").']
    }),
    arr({
      id: '5.5', from: 'tika', due: '13:00', title: 'Biggest sales first with SORTBY', heads: ['Product'], spillSize: [1, 20],
      intro: ['The products of every sale, ordered by <b>net, biggest first</b>.', '@file', 'Just the product names, sorted by column E.'],
      ask: 'Products, sorted by net, biggest first, from G5.',
      lesson: {
        idea: 'SORTBY sorts one list by another column.',
        teach: '<p><code>SORTBY(C5:C24, E5:E24, -1)</code> returns the products in order of their net, biggest first. The column you sort by does not have to be shown.</p><p>1 is smallest first, -1 biggest first.</p>',
        ex: [['=SORTBY(' + R('B') + ',', 'The distributors,'], [R('D') + ',1)', 'by cartons, smallest first.']],
        exFormula: '=INDEX(SORTBY(' + R('B') + ',' + R('D') + ',1),1)', exLabel: 'placed the smallest order',
        bridge: 'Like ranking invoices by amount.'
      },
      starter: '=SORTBY(', fnKey: 'SORTBY(', chips: [R('C'), ',' + R('E'), ',-1)'],
      solution: '=SORTBY(' + R('C') + ',' + R('E') + ',-1)', alts: ['=SORTBY($C$5:$C$24,$E$5:$E$24,-1)'],
      traps: [{ code: 'asc', formula: '=SORTBY(' + R('C') + ',' + R('E') + ',1)', title: 'Smallest first', plain: '1 sorts smallest first. Biggest first is -1.', reply: 'The smallest sale is on top 😄 I need the <b>biggest first</b>, -1.' }],
      requires: [{ re: /SORT(BY)?\s*\(/i, title: 'Use SORTBY', plain: 'This request practices SORTBY.', reply: 'Right order! SORTBY please 😉' }],
      replies: { pass: ['Perfect ranking 🏆'], frozen: 'A sale was corrected and the order did not change.', typed: 'Typed? It has to follow the sales 🙏' },
      note: { fn: 'SORTBY', what: 'SORTBY(list, sort by, 1 or -1). Sort by a column you do not show.', shape: '=SORTBY(C5:C24,E5:E24,-1)' },
      hints: ['Sort the products by another column.', 'SORTBY(C5:C24, by E5:E24, biggest first is -1).', 'It looks like =SORTBY(C5:C24,E5:E24,-1).']
    }),
    arr({
      id: '5.6', from: 'anton', due: '13:45', title: 'Line numbers with SEQUENCE', heads: ['No.'], spillSize: [1, 20],
      intro: ['For the printout:', 'Number the sales lines 1, 2, 3… down column G.', '@file'],
      ask: 'Numbers 1 to the last line, from one formula in G5.',
      lesson: {
        idea: 'SEQUENCE makes a list of numbers.',
        teach: '<p><code>SEQUENCE(20)</code> spills 1 to 20 down. <code>SEQUENCE(ROWS(A5:A24))</code> counts the rows for you, so it follows if the table grows.</p><p>SEQUENCE(rows, columns, start, step) can also go across or count by 5.</p>',
        ex: [['=SEQUENCE(3,1,', 'Three numbers,'], ['10,5)', 'from 10, counting by 5.']],
        exFormula: '=SUM(SEQUENCE(3,1,10,5))', exLabel: 'is 10 + 15 + 20',
        bridge: 'Like voucher numbers on a batch.'
      },
      starter: '=SEQUENCE(', fnKey: 'SEQUENCE(', chips: ['ROWS(' + R('A') + ')', '20', ')'],
      solution: '=SEQUENCE(ROWS(' + R('A') + '))', alts: ['=SEQUENCE(20)', '=SEQUENCE(COUNTA(' + R('B') + '))'],
      traps: [{ code: 'row', formula: '=ROW(' + R('A') + ')', title: 'Those are row numbers', plain: 'ROW gives the sheet\'s row numbers, 5 to 24. SEQUENCE starts at 1.', reply: 'It starts at {value}, not 1 😄 Those are the <b>sheet</b> row numbers.' }],
      requires: [{ re: /SEQUENCE\s*\(/i, title: 'Use SEQUENCE', plain: 'This request practices SEQUENCE.', reply: 'Right numbers! SEQUENCE please 😉' }],
      replies: { pass: ['Numbered. The printout looks tidy 🖨️'], frozen: 'Hmm, the numbers did not follow.', typed: 'Typed numbers break when lines are added 🙏' },
      note: { fn: 'SEQUENCE', what: 'SEQUENCE(rows, columns, start, step) spills a list of numbers.', shape: '=SEQUENCE(ROWS(A5:A24))' },
      hints: ['One function makes 1, 2, 3…', 'SEQUENCE with the number of rows, ROWS(A5:A24).', 'It looks like =SEQUENCE(ROWS(A5:A24)).']
    }),
    arr({
      id: '5.7', from: 'citra', due: '14:30', title: 'Name the parts with LET', top: (v) => ({ ...pick(v), A2: 'Average sale (Rp)' }),
      intro: ['Citra here. A readable formula, please 🙏', 'For the distributor in <b>B1</b>: total net divided by number of sales, in <b>B2</b>.', '@file', 'Mbak Sari says LET makes it readable. Name the total and the count.'],
      ask: 'Average sale for the distributor in <b>B1</b>, written with LET, in <b>B2</b>.',
      lesson: {
        idea: 'LET gives names to parts of a formula.',
        teach: '<p><code>LET(total, SUMIFS(…), count, COUNTIFS(…), total/count)</code>: name a value, use the name later. The last part is the answer.</p><p>Long formulas become readable, and each part is worked out once.</p>',
        ex: [['=LET(x,', 'Call it x:'], ['SUM(' + R('D') + '),', 'the total cartons,'], ['x/1000)', 'then use x.']],
        exFormula: '=LET(x,SUM(' + R('D') + '),x/1000)', exLabel: 'thousand cartons',
        bridge: 'Like labelling the subtotals on a working paper before the final line.'
      },
      starter: '=LET(', fnKey: 'LET(', chips: ['total,', 'SUMIFS(E5:E24,B5:B24,B1),', 'count,', 'COUNTIFS(B5:B24,B1),', 'total/count)'],
      solution: '=LET(total,SUMIFS(E5:E24,B5:B24,B1),count,COUNTIFS(B5:B24,B1),total/count)',
      alts: ['=LET(c,COUNTIFS(B5:B24,B1),t,SUMIFS(E5:E24,B5:B24,B1),t/c)'],
      traps: [{ code: 'total', formula: '=LET(total,SUMIFS(E5:E24,B5:B24,B1),total)', title: 'That is the total', plain: 'Divide the total by the count.', reply: 'Rp {value} is their total. I need the <b>average</b>: divided by the number of sales.' }],
      requires: [{ re: /LET\s*\(/i, title: 'Use LET', plain: 'This request practices LET.', reply: 'Right number! But Sari wants it written with LET 😉' }],
      replies: { pass: ['Rp {value}, and I can read the formula too 🙏'], frozen: 'I changed B1 and the average did not move. Point at <b>B1</b>.', typed: 'Typed? It has to follow B1.' },
      note: { fn: 'LET', what: 'LET(name, value, …, result) names parts of a formula.', shape: '=LET(total,SUMIFS(…),count,COUNTIFS(…),total/count)' },
      hints: ['Name the total, name the count, then divide.', 'total is SUMIFS(E5:E24,B5:B24,B1). count is COUNTIFS(B5:B24,B1).', 'It looks like =LET(total,SUMIFS(E5:E24,B5:B24,B1),count,COUNTIFS(B5:B24,B1),total/count).']
    })
  );
  LIST.push(Object.assign(arr({
    id: '5.8', from: 'tika', due: '16:00', title: 'FILTER in real Excel', heads: ['Product', 'Cartons', 'Net (Rp)'], spillSize: [3, 20], top: pick,
    intro: ['Last one, in <b>real Excel</b> 💻', 'Download, put a FILTER in <b>G5</b> for the distributor in B1 (product, cartons, net), save and send it back.', '@file'],
    ask: 'In Excel: FILTER in G5 for the distributor in B1, then upload.',
    lesson: { idea: 'Spilling, in real Excel.', teach: '<p>Type the FILTER in G5 only. Excel fills the rest. Save and upload.</p><p>Excel 2019 and older do not spill. You need Excel 365 or 2021.</p>', bridge: 'A live report in one cell.' },
    starter: '=FILTER(', fnKey: '', chips: [],
    solution: '=FILTER(C5:E24,B5:B24=B1)', alts: [],
    rowReply: 'Row {row} of your FILTER shows {got}, but it should be {want}.',
    replies: { pass: ['Received, and it spills perfectly. Makasih! 🙏', 'That is Chapter 5. One formula, whole reports.'], wrong: '', typed: 'G5 has no formula, the rows are typed in 😅 Put the FILTER in G5.' },
    after: 'Chapter 6 is reporting: PivotTables, charts, conditional formatting and data validation. It arrives in the next build.',
    note: { fn: 'Spilling in real Excel', what: 'One formula in one cell, and Excel 365 fills the rest.', shape: '=FILTER(C5:E24,B5:B24=B1)' }
  }), { kind: 'excel' }));

  /* ---------- Chapter 6: Reporting ----------
     PivotTables, charts, conditional formatting and data validation, built in Kicau's report builders and
     checked setting by setting. The chapter ends with a real PivotTable made in Excel. */
  const MONS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const DATA = rows.map((r) => ({ Date: r.date.slice(8) + ' ' + MONS[+r.date.slice(5, 7) - 1], Distributor: r.distributor, Product: r.product, Category: r.category, Cartons: r.cartons, Net: r.net }));
  function repSheet() {
    const c = { A4: 'Date', B4: 'Distributor', C4: 'Product', D4: 'Category', E4: 'Cartons', F4: 'Net (Rp)' };
    rows.forEach((r, i) => { const n = FIRST + i; c['A' + n] = D.toSerial(r.date); c['B' + n] = r.distributor; c['C' + n] = r.product; c['D' + n] = r.category; c['E' + n] = r.cartons; c['F' + n] = r.net; });
    return c;
  }
  const PIVOT_NUDGE = {
    rows: 'Look at what goes down the side. That field belongs in <b>Rows</b>.',
    cols: 'Check the <b>Columns</b> box. Remove what should not be there, or add what should.',
    val: 'The numbers come from <b>Values</b>. Check which column is in there.',
    agg: 'Check <b>Sum, Count or Average</b> under the boxes.',
    pct: 'Tap <b>% of grand total</b> under the boxes.',
    other: 'Compare your PivotTable with what was asked, box by box.'
  };
  function report(o) {
    const cells = repSheet();
    return Object.assign({
      chapter: 6, file: 'laporan_oktober.xlsx', data: DATA, cells, cols: ['A', 'B', 'C', 'D', 'E', 'F'], num: ['E', 'F'], dateCols: ['A'],
      first: FIRST, last: LAST, answerCell: 'H5', nudges: PIVOT_NUDGE, startCfg: {}
    }, o);
  }
  LIST.push(
    report({
      id: '6.1', kind: 'pivot', from: 'tika', due: '09:00', title: 'First PivotTable', noCols: false,
      intro: ['Pagi {name}! Big week: reports 📊', 'I need <b>total net per distributor</b> for October. No formulas this time, a <b>PivotTable</b>.', '@file', 'Distributors down the side, sum of Net next to them.'],
      ask: 'PivotTable: Distributor in Rows, Sum of Net in Values.',
      lesson: { idea: 'A PivotTable groups rows and adds them up, without a single formula.', teach: '<p>Drag a field into <b>Rows</b> to list each value once, and a number field into <b>Values</b> to add it up. Excel does the SUMIFS for you.</p><p>Here: tap a field, then tap where it goes.</p>', bridge: 'Like a trial balance from the general ledger: one line per account, with the total.' },
      want: { rows: 'Distributor', cols: null, val: 'Net', agg: 'Sum' },
      replies: { pass: ['Mantap, that is my October summary in ten seconds 🙌'], rows: 'Hmm, I need the <b>distributors</b> down the side.', cols: 'Keep it simple please, nothing across the top yet.', val: 'Those numbers are not rupiah 😅 I need <b>Net</b>.', agg: 'I need the <b>total</b>, so Sum.', other: 'Not quite what I need.' },
      hints: ['Distributor goes in Rows, Net in Values.', 'Tap Distributor, then Rows. Tap Net, then Values.', 'Values should say Sum of Net.'],
      note: { fn: 'PivotTable', what: 'Rows list each value once, Values add up a number field. No formulas needed.', shape: 'Rows: Distributor · Values: Sum of Net' }
    }),
    report({
      id: '6.2', kind: 'pivot', from: 'bayu', due: '09:45', title: 'Cartons by category',
      intro: ['For the warehouse plan:', 'How many <b>cartons per category</b>: noodles, coffee, sambal, soap?', '@file'],
      ask: 'PivotTable: Category in Rows, Sum of Cartons.',
      lesson: { idea: 'Change one field and the same PivotTable answers a new question.', teach: '<p>Swap Distributor for Category in Rows, and Net for Cartons in Values.</p><p>A PivotTable never changes the data. It only changes the view.</p>', bridge: 'Like the same ledger summarised by department instead of by account.' },
      want: { rows: 'Category', cols: null, val: 'Cartons', agg: 'Sum' },
      replies: { pass: ['That is my shelf plan 📦 Makasih!'], rows: 'I need it by <b>category</b> 🙏', cols: 'Nothing across the top please.', val: 'I count <b>cartons</b>, not rupiah 😄', agg: 'Total cartons please, Sum.', other: 'Not quite.' },
      hints: ['Category in Rows, Cartons in Values.', 'Values should be Sum of Cartons.', 'Remove anything in Columns.'],
      note: { fn: 'PivotTable rows', what: 'Any field can go in Rows. Change it to answer a new question.', shape: 'Rows: Category · Values: Sum of Cartons' }
    }),
    report({
      id: '6.3', kind: 'pivot', from: 'anton', due: '10:30', title: 'A cross-tab with Columns',
      intro: ['I want the full picture:', '<b>Distributors</b> down the side, <b>categories</b> across the top, <b>net</b> in the middle.', '@file'],
      ask: 'Rows: Distributor, Columns: Category, Values: Sum of Net.',
      lesson: { idea: 'Columns split the numbers across the top.', teach: '<p>A field in <b>Columns</b> makes one column per value, with grand totals on both sides.</p><p>This is a cross-tab: two ways of grouping at once.</p>', bridge: 'Like a matrix of cost centre by expense type.' },
      want: { rows: 'Distributor', cols: 'Category', val: 'Net', agg: 'Sum' },
      replies: { pass: ['Exactly the table for Thursday\'s meeting 👍'], rows: 'Distributors <b>down the side</b>, please.', cols: 'I need the <b>categories across the top</b>.', val: 'In rupiah please, Net.', agg: 'Totals, so Sum.', other: 'Not quite.' },
      hints: ['Three boxes: Rows, Columns, Values.', 'Distributor in Rows, Category in Columns.', 'Net in Values as Sum.'],
      note: { fn: 'PivotTable columns', what: 'A field in Columns splits the values across the top. Two groupings at once.', shape: 'Rows: Distributor · Columns: Category · Values: Sum of Net' }
    }),
    report({
      id: '6.4', kind: 'pivot', from: 'bayu', due: '11:15', title: 'Count of deliveries',
      intro: ['Not rupiah this time:', 'How many <b>deliveries</b> did each distributor get? Just count the rows.', '@file'],
      ask: 'Rows: Distributor, Values: Count.',
      lesson: { idea: 'Values can count instead of add.', teach: '<p>Switch Values from <b>Sum</b> to <b>Count</b> and the PivotTable counts rows, like COUNTIF for every distributor at once.</p><p>Average works the same way.</p>', bridge: 'Like counting vouchers per account.' },
      want: { rows: 'Distributor', cols: null, agg: 'Count' },
      replies: { pass: ['That is my truck plan for the month 🚚'], rows: 'Per <b>distributor</b> please.', cols: 'Nothing across the top.', agg: 'Those are totals, not how many. Switch to <b>Count</b> 😄', other: 'Not quite.' },
      hints: ['You want how many, not how much.', 'Put Distributor in Rows and anything in Values, then pick Count.', 'Values should say Count.'],
      note: { fn: 'Count in a PivotTable', what: 'Switch Values to Count to count rows per group.', shape: 'Rows: Distributor · Values: Count' }
    }),
    report({
      id: '6.5', kind: 'pivot', from: 'tika', due: '13:00', title: 'Share of total', showPct: true,
      intro: ['Pak Anton wants <b>shares</b>, not rupiah:', 'Each distributor\'s net as a <b>% of the grand total</b>.', '@file'],
      ask: 'Rows: Distributor, Sum of Net, shown as % of grand total.',
      lesson: { idea: 'Show Values As changes how a number is shown.', teach: '<p>In Excel, right-click a value, <b>Show Values As</b>, <b>% of Grand Total</b>. The totals become shares that add up to 100%.</p><p>Here: tap <b>% of grand total</b>.</p>', bridge: 'Like a common-size income statement: every line as a % of revenue.' },
      want: { rows: 'Distributor', cols: null, val: 'Net', agg: 'Sum', pct: true },
      replies: { pass: ['Perfect, shares that add up to 100% 💯'], rows: 'Per distributor please.', cols: 'Nothing across the top.', val: 'Shares of <b>net</b> please.', agg: 'Shares of the total, so Sum.', pct: 'Those are rupiah. I need <b>% of the grand total</b>.', other: 'Not quite.' },
      hints: ['Same as the first PivotTable, shown as percentages.', 'Distributor in Rows, Sum of Net in Values.', 'Tap % of grand total.'],
      note: { fn: 'Show Values As', what: '% of grand total turns totals into shares.', shape: 'Sum of Net · % of grand total' }
    }),
    report({
      id: '6.6', kind: 'pivot', from: 'anton', due: '13:45', title: 'Average order by product',
      intro: ['Which products bring the <b>biggest orders on average</b>?', 'Average net per sale, by product.', '@file'],
      ask: 'Rows: Product, Values: Average of Net.',
      lesson: { idea: 'Average in a PivotTable gives the mean per group.', teach: '<p>Switch Values to <b>Average</b>: each product shows its average net per sale.</p>', bridge: 'Like average ticket size by product line.' },
      want: { rows: 'Product', cols: null, val: 'Net', agg: 'Average' },
      replies: { pass: ['Interesting, sambal orders are big 🌶️'], rows: 'By <b>product</b> please.', cols: 'Nothing across the top.', val: 'Average of <b>net</b>.', agg: 'I need the <b>average</b> per sale, not the total.', other: 'Not quite.' },
      hints: ['By product, an average.', 'Product in Rows, Net in Values.', 'Switch Sum to Average.'],
      note: { fn: 'Average in a PivotTable', what: 'Average shows the mean per group.', shape: 'Rows: Product · Values: Average of Net' }
    }),
    report({
      id: '6.7', kind: 'chart', from: 'anton', due: '14:30', title: 'A column chart',
      intro: ['For my slides:', 'A chart comparing <b>net by distributor</b>. Easy to compare at a glance.', '@file'],
      ask: 'A column chart of Net by Distributor.',
      lesson: { idea: 'Pick the chart that fits the question.', teach: '<p><b>Column or bar</b>: compare amounts between groups. <b>Line</b>: change over time. <b>Pie</b>: parts of one whole, only with a few slices.</p><p>In Excel: select the data or a PivotTable, Insert, then the chart.</p>', bridge: 'Like choosing the right exhibit for an audit report.' },
      want: { type: 'Column', cat: 'Distributor', val: 'Net' },
      nudges: { type: 'Comparing amounts between groups is a <b>Column</b> chart.', cat: 'The groups along the bottom are the <b>distributors</b>.', val: 'The heights should be <b>Net</b>.', other: 'Check the type, the categories and the values.' },
      replies: { pass: ['Clear at a glance. Going on slide 2 📈'], type: 'Hmm, that chart type makes it hard to compare. A <b>column</b> chart please.', cat: 'I need the <b>distributors</b> along the bottom.', val: 'In rupiah please, <b>Net</b>.', other: 'Not quite.' },
      hints: ['Comparing groups: column chart.', 'Categories: Distributor. Values: Net.', 'Column, Distributor, Net.'],
      note: { fn: 'Column charts', what: 'Column or bar to compare groups, line for time, pie for a few parts of a whole.', shape: 'Column · Net by Distributor' }
    }),
    report({
      id: '6.8', kind: 'chart', from: 'tika', due: '15:15', title: 'A pie for shares',
      intro: ['One more chart:', 'How October\'s net <b>splits between the four categories</b>. Parts of one whole.', '@file'],
      ask: 'A pie chart of Net by Category.',
      lesson: { idea: 'A pie shows parts of one whole.', teach: '<p>Use a pie only for a few parts that add up to 100%. Four categories is fine. Twenty products is not.</p>', bridge: 'Like a revenue mix chart in an annual report.' },
      want: { type: 'Pie', cat: 'Category', val: 'Net' },
      nudges: { type: 'Parts of one whole, with only four parts, is a <b>Pie</b>.', cat: 'The slices are the <b>categories</b>.', val: 'The slices should be sized by <b>Net</b>.', other: 'Check the type, the categories and the values.' },
      replies: { pass: ['Lovely, the mix is clear 🥧'], type: 'I want to see the split of the whole, so a <b>pie</b>, please.', cat: 'Split by <b>category</b> please.', val: 'Sized by rupiah, <b>Net</b>.', other: 'Not quite.' },
      hints: ['Parts of a whole: pie.', 'Categories: Category. Values: Net.', 'Pie, Category, Net.'],
      note: { fn: 'Pie charts', what: 'Parts of one whole, only for a few slices.', shape: 'Pie · Net by Category' }
    }),
    report({
      id: '6.9', kind: 'rule', rule: 'cf', from: 'lestari', due: '16:00', title: 'Highlight big sales', startCfg: { style: 'green' },
      intro: ['Bu Lestari here.', 'On the sales list, make every <b>Net above 5,000,000</b> stand out in <b>green</b>.', '@file'],
      ask: 'Conditional formatting: Net greater than 5,000,000, green fill.',
      lesson: { idea: 'Conditional formatting colors cells by their value.', teach: '<p>In Excel: select the column, Home, <b>Conditional Formatting</b>, Highlight Cells Rules, Greater Than. The color follows the numbers when they change.</p>', bridge: 'Like ticking the big items on a listing for review.' },
      want: { col: 'Net', op: '>', a: 5000000, style: 'green' },
      nudges: { col: 'The rule goes on the <b>Net</b> column.', op: 'Above means <b>greater than</b>.', a: 'The threshold is <b>5000000</b>.', style: 'Pick <b>Green fill</b>.', other: 'Check the column, the condition, the number and the color.' },
      replies: { pass: ['Perfect, the big ones jump out now ✅'], col: 'The rule should be on <b>Net</b>, not cartons.', op: 'Above 5 million, so <b>greater than</b>.', a: 'The line is <b>5,000,000</b>.', style: 'Green please 🙂', other: 'Not quite.' },
      hints: ['Net, greater than, 5000000, green.', 'Pick Net, then greater than.', 'Type 5000000 and pick Green fill.'],
      note: { fn: 'Conditional formatting', what: 'Colors cells by a rule, and follows the values when they change.', shape: 'Net > 5,000,000 · green fill' }
    }),
    report({
      id: '6.10', kind: 'rule', rule: 'dv', from: 'bayu', due: '16:30', title: 'Stop bad entries with validation', startCfg: { op: 'between' },
      intro: ['People keep typing 1000 cartons by mistake 😩', 'Make the <b>Cartons</b> column accept only whole numbers <b>from 1 to 100</b>.', '@file'],
      ask: 'Data validation on Cartons: between 1 and 100.',
      lesson: { idea: 'Data validation refuses entries that break a rule.', teach: '<p>In Excel: select the column, Data, <b>Data Validation</b>, Allow: Whole number, between 1 and 100. A wrong entry gets an error message.</p><p>It also powers dropdown lists, with Allow: List.</p>', bridge: 'Like an input control in a system: wrong values never get in.' },
      want: { col: 'Cartons', op: 'between', a: 1, b: 100 },
      nudges: { col: 'The rule goes on <b>Cartons</b>.', op: 'Pick <b>between</b>.', a: 'The lowest allowed is <b>1</b>.', b: 'The highest allowed is <b>100</b>.', other: 'Check the column and the two limits.' },
      replies: { pass: ['No more 1000-carton typos 🙌'], col: 'It is the <b>Cartons</b> column that gets typos.', op: 'Between 1 and 100.', a: 'The lowest is <b>1</b>.', b: 'The highest is <b>100</b>.', other: 'Not quite.' },
      hints: ['Cartons, between 1 and 100.', 'Pick Cartons, then type 1 and 100.', 'From 1 to 100.'],
      note: { fn: 'Data validation', what: 'Refuses entries that break a rule. Also used for dropdown lists.', shape: 'Cartons · whole number between 1 and 100' }
    })
  );
  LIST.push(report({
    id: '6.11', kind: 'excel', pivotCheck: true, from: 'tika', due: '17:00', title: 'A real PivotTable in Excel',
    intro: ['Last one, in <b>real Excel</b> 💻', 'Download the sales, insert a <b>PivotTable</b> with total net per distributor, save and send it back.', '@file'],
    ask: 'In Excel: a PivotTable of Sum of Net by Distributor, then upload.',
    lesson: { idea: 'A PivotTable in real Excel.', teach: '<p>Click inside the data, Insert, <b>PivotTable</b>, OK. Tick Distributor and Net in the field list. Save and upload.</p><p>Kicau finds your PivotTable on any sheet and checks every distributor\'s total.</p>', bridge: 'The report you will build every month-end.' },
    replies: { pass: ['Received, every total matches. Makasih! 🙏', 'That is Chapter 6. You can build reports now.'], typed: '', wrong: '' },
    after: 'Chapter 7 is Power Query: import, clean, merge and unpivot. It arrives in the next build.',
    note: { fn: 'PivotTable in real Excel', what: 'Insert PivotTable, tick the fields, done.', shape: 'Rows: Distributor · Values: Sum of Net' }
  }));

  /* ---------- Chapter 7: Power Query ----------
     Raw files that need cleaning before anyone can report on them. The learner adds applied steps, and the
     finished table is compared with the table the right steps produce, so any path that gets there passes. */
  const QN = { cols: 'Look at the column headers. Something is missing, extra, or not promoted yet.', rows: 'Count the rows. Too many means a step is missing; too few means a filter is too strict.', values: 'The shape is right, but some values differ. Check text, spaces and types.', other: 'Compare your table with what was asked.' };
  const sales5 = rows.map((r) => [r.date.slice(8) + ' ' + MONS[+r.date.slice(5, 7) - 1], r.distributor, r.product, r.cartons, r.net]);
  function query(o) {
    return Object.assign({ chapter: 7, kind: 'query', file: o.sources.main.name, nudges: QN, startCfg: { steps: [] }, cells: repSheet(), data: DATA, cols: ['A', 'B', 'C', 'D', 'E', 'F'], num: ['E', 'F'], dateCols: ['A'], first: FIRST, last: LAST, answerCell: 'H5' }, o);
  }
  const SNAP = D.build().snapshots;
  const dpdTable = ['D02', 'D05', 'D08', 'D12', 'D14'].map((id) => [D.DISTRIBUTORS.find((d) => d.id === id).name, ...D.MONTHS.map((mo) => Math.max(...SNAP.filter((x) => x.distributor_id === id && x.month === mo).map((x) => x.dpd)))]);
  LIST.push(
    query({
      id: '7.1', from: 'bayu', due: '08:30', title: 'Import a messy CSV',
      sources: { main: { name: 'export_penjualan.csv', head: null, rows: [['Date', 'Distributor', 'Product', 'Cartons', 'Net'], ...sales5.flatMap((r, i) => (i % 5 === 4 ? [r, ['', '', '', '', '']] : [r]))] } },
      ops: ['promote', 'blank', 'trim', 'remove'], solutionSteps: [{ op: 'promote' }, { op: 'blank' }],
      intro: ['Pagi {name} 👷 The old system exports a CSV every morning.', 'The headers come in as the first row, and there are blank lines everywhere.', '@file', 'Clean it with <b>Power Query</b>: proper headers, no blank rows.'],
      ask: 'Use the first row as headers, then remove blank rows.',
      lesson: { idea: 'Power Query cleans data in steps, and replays them every time.', teach: '<p>Each thing you do becomes an <b>applied step</b>. Next month, click Refresh and every step runs again on the new file.</p><p>In Excel: Data, <b>From Text/CSV</b>, then Transform Data.</p>', bridge: 'Like a written procedure for preparing a ledger extract, done by the machine.' },
      replies: { pass: ['Clean! And tomorrow I just click Refresh? 😮'], cols: 'The headers still say Column1, Column2… 😅', rows: 'There are still empty lines in it.', values: 'Something changed in the values that should not have.', other: 'Not quite.' },
      hints: ['Two steps: headers, then blanks.', 'Add "Use first row as headers".', 'Then add "Remove blank rows".'],
      note: { fn: 'Power Query basics', what: 'Every action is an applied step that replays on Refresh.', shape: 'Use first row as headers → Remove blank rows' }
    }),
    query({
      id: '7.2', from: 'tika', due: '09:15', title: 'Trim text and fix a number column',
      sources: { main: { name: 'penjualan_kotor.xlsx', head: ['Date', 'Distributor', 'Product', 'Cartons', 'Net'], rows: sales5.map((r, i) => [r[0], messy(r[1], i), r[2], r[3], 'Rp ' + r[4].toLocaleString('de-DE')]) } },
      ops: ['trim', 'number', 'remove', 'blank'], solutionSteps: [{ op: 'trim', col: 'Distributor' }, { op: 'number', col: 'Net' }],
      intro: ['This file came from a branch 😩', 'The distributor names have extra spaces, and Net is text like "Rp 4.351.200", so it will not add up.', '@file'],
      ask: 'Trim Distributor, and change Net to a number.',
      lesson: { idea: 'Clean text and set the right data type.', teach: '<p><b>Trim</b> removes extra spaces. <b>Change type</b> to a whole number makes Net add up again.</p><p>A dot shows each space here, so you can see them.</p>', bridge: 'Like re-keying amounts from a printout into proper numbers.' },
      replies: { pass: ['Perfect, it adds up now 🙏'], cols: 'Some columns are gone or renamed 🤔', rows: 'The number of rows changed. I did not ask for that.', values: 'Still messy: check the spaces in Distributor and the text in Net.', other: 'Not quite.' },
      hints: ['Two columns need fixing.', 'Trim on Distributor.', 'Change type to number on Net.'],
      note: { fn: 'Trim and change type', what: 'Trim removes extra spaces. Change type makes text into numbers.', shape: 'Trim Distributor → Change type Net' }
    }),
    query({
      id: '7.3', from: 'anton', due: '10:00', title: 'Filter rows, remove a column',
      sources: { main: { name: 'penjualan_kategori.xlsx', head: ['Date', 'Distributor', 'Product', 'Category', 'Cartons', 'Net'], rows: rows.map((r) => [r.date.slice(8) + ' ' + MONS[+r.date.slice(5, 7) - 1], r.distributor, r.product, r.category, r.cartons, r.net]) } },
      ops: ['filter', 'remove', 'trim'], solutionSteps: [{ op: 'filter', col: 'Category', val: 'Coffee' }, { op: 'remove', col: 'Date' }],
      intro: ['I only care about <b>Coffee</b> today ☕', 'Keep only the coffee rows, and drop the <b>Date</b> column.', '@file'],
      ask: 'Keep only Category = Coffee, and remove the Date column.',
      lesson: { idea: 'Filter rows and remove columns, as steps.', teach: '<p>A filter in Power Query is a step: it keeps working when new rows arrive. Removing a column is a step too.</p>', bridge: 'Like a standing report that only shows one cost centre.' },
      replies: { pass: ['Coffee only, nice and short ☕'], cols: 'The Date column is still there, or another one went missing.', rows: 'I see more than just coffee 😄', values: 'Something is off in the values.', other: 'Not quite.' },
      hints: ['One filter, one column removed.', 'Keep rows where Category is Coffee.', 'Remove column Date.'],
      note: { fn: 'Filter and remove', what: 'Filters and removed columns are steps that keep working on new data.', shape: 'Keep Category = Coffee → Remove Date' }
    }),
    query({
      id: '7.4', from: 'tika', due: '10:45', title: 'Append two branch files',
      sources: { main: { name: 'jakarta.xlsx', head: ['Date', 'Distributor', 'Product', 'Cartons', 'Net'], rows: sales5.slice(0, 10) }, other: { name: 'bandung.xlsx', head: ['Date', 'Distributor', 'Product', 'Cartons', 'Net'], rows: sales5.slice(10) } },
      ops: ['append', 'remove', 'blank'], solutionSteps: [{ op: 'append' }],
      intro: ['Two branches, two files 😅', 'Jakarta and Bandung send the same columns. Stack them into <b>one table</b>.', '@file'],
      ask: 'Append bandung.xlsx under jakarta.xlsx.',
      lesson: { idea: 'Append stacks tables with the same columns, one under the other.', teach: '<p>Home, <b>Append Queries</b>. Next month, new branch files stack automatically.</p><p>Merge is different: it adds columns side by side.</p>', bridge: 'Like consolidating two branch ledgers into one.' },
      replies: { pass: ['One table, both branches 👍'], cols: 'The columns changed. I only wanted them stacked.', rows: 'I only see one branch.', values: 'Something is off.', other: 'Not quite.' },
      hints: ['Stack the second file under the first.', 'Add the Append step.', 'Append adds bandung.xlsx.'],
      note: { fn: 'Append', what: 'Stacks tables with the same columns.', shape: 'jakarta.xlsx + bandung.xlsx' }
    }),
    Object.assign(query({
      id: '7.5', from: 'bayu', due: '11:30', title: 'Merge in the category',
      sources: { main: { name: 'pesanan.xlsx', head: ['Date', 'Distributor', 'SKU', 'Cartons'], rows: rows.map((r) => [r.date.slice(8) + ' ' + MONS[+r.date.slice(5, 7) - 1], r.distributor, r.sku, r.cartons]) }, lookup: { name: 'daftar_harga.xlsx', head: ['SKU', 'Product', 'Category'], rows: D.PRODUCTS.map((x) => [x.sku, x.name, x.category]) } },
      ops: ['merge', 'remove', 'trim'], solutionSteps: [{ op: 'merge', col: 'SKU' }],
      intro: ['The orders only have SKUs.', 'Bring in each order\'s <b>Category</b> from the price list, matching on <b>SKU</b>.', '@file'],
      ask: 'Merge with the price list on SKU, bringing in Category.',
      lesson: { idea: 'Merge is a lookup for whole tables.', teach: '<p>Home, <b>Merge Queries</b>: pick the matching column in both tables, then expand the column you need. It is VLOOKUP for an entire table.</p>', bridge: 'Like matching a sub-ledger to the master file by code.' },
      replies: { pass: ['Every order has its category now 📦'], cols: 'I need a <b>Category</b> column next to the orders.', rows: 'Some orders went missing.', values: 'Some categories look wrong. Did you match on SKU?', other: 'Not quite.' },
      hints: ['Match on the code both tables share.', 'Merge, matching on SKU.', 'Merge on SKU brings in Category.'],
      note: { fn: 'Merge', what: 'Joins two tables on a matching column, like VLOOKUP for a whole table.', shape: 'Merge pesanan with daftar_harga on SKU' }
    }), { mergeCol: 'Category' }),
    Object.assign(query({
      id: '7.6', from: 'citra', due: '13:30', title: 'Unpivot months into rows',
      sources: { main: { name: 'dpd_bulanan.xlsx', head: ['Distributor', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct'], rows: dpdTable } },
      ops: ['unpivot', 'remove', 'filter'], solutionSteps: [{ op: 'unpivot', col: 'Distributor' }],
      intro: ['Our days-late report is <b>wide</b>: one column per month.', 'PivotTables and charts need it <b>long</b>: Distributor, Month, Days late.', '@file'],
      ask: 'Unpivot every month column into Month and Days late rows.',
      lesson: { idea: 'Unpivot turns columns into rows.', teach: '<p>Select the column to keep, then <b>Unpivot Other Columns</b>. Each month becomes its own row, with the month name in one column and the number in another.</p><p>Long data is what PivotTables love.</p>', bridge: 'Like turning a monthly schedule into a journal with one line per month.' },
      replies: { pass: ['Now I can pivot and chart it any way 📊'], cols: 'I need three columns: <b>Distributor, Month, Days late</b>.', rows: 'There should be one row per distributor per month.', values: 'Some values do not match.', other: 'Not quite.' },
      hints: ['Keep Distributor, turn the months into rows.', 'Unpivot other columns, keeping Distributor.', 'Pick Unpivot, then Distributor.'],
      note: { fn: 'Unpivot', what: 'Turns month columns into Month and value rows: long data for PivotTables.', shape: 'Unpivot other columns than Distributor' }
    }), { unpivotNames: ['Month', 'Days late'] }),
    Object.assign(query({
      id: '7.7', from: 'dimas', due: '14:15', title: 'Group by distributor',
      sources: { main: { name: 'fasilitas.xlsx', head: ['Facility', 'Distributor', 'Type', 'Outstanding'], rows: FAC.map((f) => [f.facility_id, f.distributor, f.type === 'Term Loan' ? 'TL' : 'WC', f.outstanding]) } },
      ops: ['group', 'remove', 'filter'], solutionSteps: [{ op: 'group', col: 'Distributor' }],
      intro: ['Some distributors have two facilities.', 'One line per <b>distributor</b>, with their total <b>outstanding</b>.', '@file'],
      ask: 'Group by Distributor, summing Outstanding.',
      lesson: { idea: 'Group By summarises rows into one line per group.', teach: '<p>Transform, <b>Group By</b>: pick the column to group on, and Sum for the number. Like a PivotTable, but as a clean table you can load anywhere.</p>', bridge: 'Like exposure per borrower, adding up every facility.' },
      replies: { pass: ['Total exposure per distributor. That is the credit view 🏦'], cols: 'I need <b>Distributor</b> and the total, nothing else.', rows: 'One line per distributor, please.', values: 'Some totals look off.', other: 'Not quite.' },
      hints: ['One line per distributor.', 'Group By Distributor.', 'It sums Outstanding for you.'],
      note: { fn: 'Group By', what: 'One row per group, with a total. A PivotTable as a clean table.', shape: 'Group by Distributor, Sum of Outstanding' }
    }), { groupSum: 'Outstanding' })
  );
  LIST.push(report({
    id: '7.8', chapter: 7, kind: 'excel', pivotCheck: true, from: 'tika', due: '16:30', title: 'Power Query in real Excel', file: 'penjualan_oktober.xlsx',
    intro: ['Last one, in <b>real Excel</b> 💻', 'Load the sales into Power Query, <b>Group By Distributor</b> with Sum of Net, Close &amp; Load, save and send it back.', '@file'],
    ask: 'In Excel: Power Query Group By Distributor, Sum of Net, then upload.',
    lesson: { idea: 'Power Query in real Excel.', teach: '<p>Click in the data, Data, <b>From Table/Range</b>. In the editor: Transform, <b>Group By</b> Distributor, Sum of Net. Close &amp; Load puts the result on a new sheet.</p><p>On a Mac, Power Query works in Excel 365, with fewer buttons than on Windows.</p>', bridge: 'The same steps, every month, with one Refresh.' },
    replies: { pass: ['Received, every total matches. Makasih! 🙏', 'That is Chapter 7. You can clean any file now.'], typed: '', wrong: '' },
    after: 'Chapter 8 is what-if and loans: Goal Seek, Data Tables, PMT, IPMT, NPV and IRR. It arrives in the next build.',
    note: { fn: 'Power Query in Excel', what: 'From Table/Range, Group By, Close & Load.', shape: 'Group By Distributor, Sum of Net' }
  }));

  /* ---------- Chapter 8: What-if and loans ----------
     Bank Kenari's loan calculator: amount in B1, yearly rate in B2, months in B3. A distributor's warehouse
     project has yearly cash flows in B8:B13. */
  const LCOLS = ['A', 'B', 'C'];
  function loanSheet(top, alt, extra) {
    const c = { A1: 'Loan (Rp)', B1: alt ? 600000000 : 500000000, A2: 'Rate per year', B2: alt ? 0.13 : 0.12, A3: 'Months', B3: alt ? 48 : 36 };
    if (extra === 'cash') { c.A7 = 'Year'; c.B7 = 'Cash flow (Rp)'; [-400000000, 90000000, 110000000, 120000000, 130000000, alt ? 150000000 : 140000000].forEach((v, i) => { c['A' + (8 + i)] = i; c['B' + (8 + i)] = v; }); c.A5 = 'Discount rate'; c.B5 = alt ? 0.11 : 0.1; }
    if (extra === 'rates') { c.A7 = 'Rate'; c.B7 = 'Installment (Rp)'; [0.08, 0.09, 0.1, 0.11, 0.12, 0.13, 0.14].forEach((v, i) => { c['A' + (8 + i)] = v; }); }
    return Object.assign(c, top);
  }
  function loan(o) {
    const top = typeof o.top === 'function' ? o.top : () => (o.top || {});
    const base = loanSheet(top('base'), false, o.extra), alt = loanSheet(top('alt'), true, o.extra);
    return formula(Object.assign({
      chapter: 8, file: 'kalkulator_kredit.xlsx', cols: LCOLS, num: ['B'], dateCols: [], first: 1, last: o.extra ? 14 : 6,
      answerCell: o.answerCell || 'B4', colRange: { A: 'A8:A14', B: 'B8:B13', C: 'C8:C14' },
      variantCells: { base, alt }, cells: base
    }, o, { top: undefined }));
  }
  const ans = (label, cell) => () => ({ ['A' + cell.slice(1)]: label });
  LIST.push(
    loan({
      id: '8.1', from: 'citra', due: '09:00', title: 'Monthly installment with PMT', top: ans('Installment (Rp)', 'B4'),
      intro: ['Pagi {name}! Loans today 🏦', 'A distributor wants the loan in <b>B1</b> at the yearly rate in <b>B2</b> over the months in <b>B3</b>.', '@file', 'What is the <b>monthly installment</b>? In <b>B4</b>, as a positive number.'],
      ask: 'Monthly installment in <b>B4</b>, positive.',
      lesson: { idea: 'PMT works out a fixed installment for a loan.', teach: '<p><code>PMT(rate per period, number of periods, loan)</code>. The rate is yearly, so divide by 12 for monthly payments.</p><p>Excel shows money you pay out as negative. Put a minus in front of the loan to get a positive installment.</p>', ex: [['=PMT(0.12/12,', '1% a month,'], ['12,', 'for 12 months,'], ['-12000000)', 'on 12 million.']], exFormula: '=PMT(0.12/12,12,-12000000)', exLabel: 'rupiah a month', bridge: 'Like an annuity table in your finance class, done in one cell.' },
      starter: '=PMT(', fnKey: 'PMT(', chips: ['B2/12', ',B3', ',-B1)'],
      solution: '=PMT(B2/12,B3,-B1)', alts: ['=-PMT(B2/12,B3,B1)', '=PMT(B2/12,B3,B1)*-1'],
      traps: [{ code: 'annual', formula: '=PMT(B2,B3,-B1)', title: 'The rate is yearly', plain: 'Payments are monthly, so the rate must be monthly too: B2/12.', reply: 'Rp {value} a month? That would bankrupt them 😅 The rate in B2 is <b>per year</b>. Divide by 12.' }, { code: 'sign', formula: '=PMT(B2/12,B3,B1)', title: 'Negative installment', plain: 'Excel shows payments as negative. Use -B1 to make it positive.', reply: 'The number is right but negative 😄 Put a minus before <b>B1</b>.' }],
      replies: { pass: ['Rp {value} a month. I will put it in the offer letter ✉️'], frozen: 'They asked for a bigger loan over 48 months and B4 did not move. Point at B1, B2 and B3.', typed: 'Typed? The terms change in every meeting 🙏' },
      note: { fn: 'PMT', what: 'PMT(rate per period, periods, -loan). Divide a yearly rate by 12.', shape: '=PMT(B2/12,B3,-B1)' },
      hints: ['Rate per month, number of months, the loan.', 'The rate per month is B2/12. The loan is -B1 for a positive result.', 'It looks like =PMT(B2/12,B3,-B1).']
    }),
    loan({
      id: '8.2', from: 'dimas', due: '09:45', title: 'Interest in month 1 with IPMT', top: ans('Interest, month 1 (Rp)', 'B4'),
      intro: ['How much of the <b>first</b> installment is interest?', '@file', 'In <b>B4</b>, positive please.'],
      ask: 'The interest part of installment 1, in <b>B4</b>.',
      lesson: { idea: 'IPMT gives the interest part of one installment.', teach: '<p><code>IPMT(rate, period, periods, -loan)</code>. Early installments are mostly interest, later ones mostly principal. <code>PPMT</code> gives the principal part.</p>', ex: [['=IPMT(0.12/12,', '1% a month,'], ['1,12,', 'month 1 of 12,'], ['-12000000)', 'on 12 million.']], exFormula: '=IPMT(0.12/12,1,12,-12000000)', exLabel: 'rupiah of interest in month 1', bridge: 'Like the first line of an amortization schedule.' },
      starter: '=IPMT(', fnKey: 'IPMT(', chips: ['B2/12', ',1', ',B3', ',-B1)'],
      solution: '=IPMT(B2/12,1,B3,-B1)', alts: ['=-IPMT(B2/12,1,B3,B1)', '=B1*B2/12'],
      traps: [{ code: 'pmt', formula: '=PMT(B2/12,B3,-B1)', title: 'That is the whole installment', plain: 'PMT is interest plus principal. IPMT is the interest part only.', reply: 'That is the whole installment 😄 I need just the <b>interest</b> part.' }],
      replies: { pass: ['Rp {value} of interest in month 1. Good for the schedule 📄'], frozen: 'The terms changed and B4 did not.', typed: 'Typed? It has to follow the terms 🙏' },
      note: { fn: 'IPMT, PPMT', what: 'IPMT is the interest part of one installment, PPMT the principal part.', shape: '=IPMT(B2/12,1,B3,-B1)' },
      hints: ['Interest part of installment number 1.', 'IPMT(rate per month, 1, months, -loan).', 'It looks like =IPMT(B2/12,1,B3,-B1).']
    }),
    loan({
      id: '8.3', from: 'citra', due: '10:30', title: 'Total interest over the loan', top: ans('Total interest (Rp)', 'B4'),
      intro: ['The distributor asks: how much <b>interest in total</b> will they pay over the whole loan?', '@file', 'In <b>B4</b>.'],
      ask: 'All installments minus the loan, in <b>B4</b>.',
      lesson: { idea: 'Total interest is everything paid minus what was borrowed.', teach: '<p>Installment × number of months is everything they pay. Take away the loan, and what is left is interest.</p>', ex: [['=PMT(0.12/12,12,-12000000)', 'One installment'], ['*12-12000000', 'times 12, minus the loan.']], exFormula: '=PMT(0.12/12,12,-12000000)*12-12000000', exLabel: 'rupiah of interest in a year', bridge: 'Like finance charges in a loan disclosure.' },
      starter: '=PMT(', fnKey: 'PMT(', chips: ['B2/12,B3,-B1)', '*B3', '-B1'],
      solution: '=PMT(B2/12,B3,-B1)*B3-B1', alts: ['=B3*PMT(B2/12,B3,-B1)-B1', '=-CUMIPMT(B2/12,B3,B1,1,B3,0)'],
      traps: [{ code: 'total', formula: '=PMT(B2/12,B3,-B1)*B3', title: 'That is everything they pay', plain: 'That includes the loan itself. Subtract B1.', reply: 'That is everything they pay back, loan included. Take away the <b>loan</b>.' }],
      replies: { pass: ['Rp {value} in interest. They will want a discount 😄'], frozen: 'The terms changed and the total did not.', typed: 'Typed? 🙏' },
      note: { fn: 'Total interest', what: 'Installment × months − loan.', shape: '=PMT(B2/12,B3,-B1)*B3-B1' },
      hints: ['Everything paid, minus the loan.', 'Everything paid is the installment times B3.', 'It looks like =PMT(B2/12,B3,-B1)*B3-B1.']
    }),
    loan({
      id: '8.4', from: 'dimas', due: '11:15', title: 'Is the project worth it? NPV', extra: 'cash', answerCell: 'B4', top: ans('NPV (Rp)', 'B4'),
      intro: ['A distributor wants to build a warehouse.', 'It costs Rp 400 juta now (year 0) and brings cash for 5 years, in <b>B8:B13</b>. Our discount rate is in <b>B5</b>.', '@file', 'What is it worth today? <b>NPV</b> in B4.'],
      ask: 'NPV of the cash flows at the rate in B5, year 0 included, in <b>B4</b>.',
      lesson: { idea: 'NPV discounts future cash to today.', teach: '<p><code>NPV(rate, year 1 to year 5)</code> discounts each year back to today. Excel\'s NPV assumes the first value is one year away, so add year 0 on its own: <code>NPV(…)+B8</code>.</p><p>Above 0 means the project earns more than the rate.</p>', ex: [['=NPV(B5,', 'Discount at B5'], ['B9:B13)', 'years 1 to 5.']], exFormula: '=NPV(B5,B9:B13)', exLabel: 'rupiah: the future cash, in today\'s money', bridge: 'Like the present value of a lease liability.' },
      starter: '=NPV(', fnKey: 'NPV(', chips: ['B5', ',B9:B13)', '+B8'],
      solution: '=NPV(B5,B9:B13)+B8', alts: ['=B8+NPV(B5,B9:B13)'],
      traps: [{ code: 'year0', formula: '=NPV(B5,B8:B13)', title: 'Year 0 was discounted too', plain: 'NPV treats its first value as one year away. Leave year 0 out and add it after.', reply: 'Close! But the Rp 400 juta is paid <b>now</b>, so it should not be discounted. Add <b>B8</b> after the NPV.' }],
      replies: { pass: ['Rp {value} today. Positive, so the warehouse pays off 👍'], frozen: 'The rate and year 5 changed and B4 stayed the same.', typed: 'Typed? 🙏' },
      note: { fn: 'NPV', what: 'NPV(rate, future cash flows) + the cash flow today.', shape: '=NPV(B5,B9:B13)+B8' },
      hints: ['Discount years 1 to 5, then add year 0.', 'NPV(B5,B9:B13), then +B8.', 'It looks like =NPV(B5,B9:B13)+B8.']
    }),
    loan({
      id: '8.5', from: 'citra', due: '13:00', title: 'Its return with IRR', extra: 'cash', answerCell: 'B4', top: ans('IRR', 'B4'),
      intro: ['Same warehouse.', 'What yearly <b>return</b> does it earn? <b>IRR</b> in B4.', '@file'],
      ask: 'The IRR of the cash flows in B8:B13, in <b>B4</b>.',
      lesson: { idea: 'IRR is the rate where NPV is zero.', teach: '<p><code>IRR(all cash flows)</code>, year 0 included. If IRR is above our discount rate, the project is worth doing.</p><p>It needs at least one negative and one positive value.</p>', ex: [['=IRR(', 'The return of'], ['B8:B12)', 'the first four years only.']], exFormula: '=IRR(B8:B12)', exLabel: 'a year, if it stopped after year 4', bridge: 'Like the effective rate of a bond.' },
      starter: '=IRR(', fnKey: 'IRR(', chips: ['B8:B13', ')'],
      solution: '=IRR(B8:B13)', alts: ['=IRR(B8:B13,0.1)'], precision: 4,
      traps: [{ code: 'noinvest', formula: '=IRR(B9:B13)', title: 'Year 0 is missing', plain: 'IRR needs the investment too, year 0, as a negative value.', reply: 'It says #NUM! 😅 IRR needs the Rp 400 juta at the start. Include <b>B8</b>.' }],
      replies: { pass: ['{value}, above our 10%. Approved in principle ✅'], frozen: 'Year 5 changed and the IRR did not.', typed: 'Typed? 🙏' },
      note: { fn: 'IRR', what: 'IRR(all cash flows, year 0 included): the rate where NPV is zero.', shape: '=IRR(B8:B13)' },
      hints: ['All cash flows, year 0 included.', 'IRR of B8:B13.', 'It looks like =IRR(B8:B13).']
    }),
    Object.assign(loan({
      id: '8.6', from: 'dimas', due: '13:45', title: 'Installments for every rate', extra: 'rates',
      intro: ['The distributor wants to see the installment at <b>every rate</b> from 8% to 14%.', '@file', 'Column B, next to each rate in A8:A14. Same loan and months for all.'],
      ask: 'Installment for each rate in A8:A14, in B8, copied down.',
      lesson: { idea: 'A what-if table: one formula, copied down, each row trying a different input.', teach: '<p>Point at the rate on the row (A8, which moves) and pin the loan and months ($B$1, $B$3).</p><p>Excel also has a built-in <b>Data Table</b> (Data, What-If Analysis) that does the same from one formula.</p>', ex: [['=PMT(A8/12,', 'The rate on this row,'], ['$B$3,-$B$1)', 'the same months and loan.']], exFormula: '=PMT(A8/12,$B$3,-$B$1)', exLabel: 'rupiah at 8%', bridge: 'Like a sensitivity table in a valuation.' },
      starter: '=PMT(', fnKey: 'PMT(', chips: ['A8/12', ',$B$3', ',-$B$1)'],
      solution: '=PMT(A8/12,$B$3,-$B$1)', alts: ['=-PMT(A8/12,$B$3,$B$1)'],
      traps: [{ code: 'nodollar', formula: '=PMT(A8/12,B3,-B1)', title: 'B1 and B3 slid down', plain: 'Copied down, B1 and B3 moved too. Pin them with $.', reply: 'Only the first rate works, the rest break 🤔 The loan and months should not move. Pin <b>$B$1</b> and <b>$B$3</b>.' }],
      replies: { pass: ['Perfect, a rate table for the meeting 📋'], frozen: 'The loan changed and the table did not follow.', typed: 'Typed? 🙏' },
      note: { fn: 'What-if table', what: 'One formula, copied down, with the input that changes relative and the rest pinned.', shape: '=PMT(A8/12,$B$3,-$B$1)' },
      hints: ['The rate moves with the row, the loan and months do not.', 'Use A8 for the rate, $B$3 and $B$1 for the rest.', 'It looks like =PMT(A8/12,$B$3,-$B$1).']
    }), { fill: { col: 'B', first: 8, last: 14 }, answerCell: 'B8' })
  );
  LIST.push({
    id: '8.7', chapter: 8, kind: 'goal', from: 'citra', due: '14:30', title: 'Goal Seek the loan size', file: 'kalkulator_kredit.xlsx',
    cells: loanSheet({ A4: 'Installment (Rp)', B4: '=PMT(B2/12,B3,-B1)' }, false), cols: LCOLS, num: ['B'], dateCols: [], first: 1, last: 6, answerCell: 'B4',
    intro: ['The distributor can only pay <b>Rp 15,000,000 a month</b>.', 'B4 already has the PMT formula. How big a loan can they take? Use <b>Goal Seek</b>.', '@file'],
    ask: 'Goal Seek: set B4 to 15,000,000 by changing B1.',
    lesson: { idea: 'Goal Seek works a formula backwards.', teach: '<p>Data, What-If Analysis, <b>Goal Seek</b>: set the cell with the formula, to the value you want, by changing one input. Excel tries values until it fits.</p>', bridge: 'Like working out the gross amount from a net payment.' },
    want: { set: 'B4', to: 15000000, change: 'B1' }, startCfg: {},
    nudges: { set: 'Set the cell with the installment formula, <b>B4</b>.', to: 'The target is <b>15000000</b>.', change: 'The loan size is in <b>B1</b>.', other: 'Check all three boxes.' },
    replies: { pass: ['So the most they can borrow is in B1 now. Perfect 🎯'], set: 'The installment is in <b>B4</b>.', to: 'They can pay <b>15,000,000</b> a month.', change: 'We are changing the <b>loan size</b>, B1.', other: 'Not quite.' },
    hints: ['Set B4, to 15000000, by changing B1.', 'Set cell is the formula cell, B4.', 'By changing B1, the loan.'],
    note: { fn: 'Goal Seek', what: 'Set a formula cell to a value by changing one input.', shape: 'Set B4 to 15,000,000 by changing B1' }
  });
  LIST.push(Object.assign(loan({
    id: '8.8', from: 'citra', due: '16:00', title: 'PMT in real Excel', top: ans('Installment (Rp)', 'B4'),
    intro: ['Last one, in <b>real Excel</b> 💻', 'Download the calculator, put the monthly installment in <b>B4</b> with PMT, save and send it back.', '@file'],
    ask: 'In Excel: PMT in B4, then upload.',
    lesson: { idea: 'PMT in real Excel.', teach: '<p>Type the PMT in B4, save, upload. Excel may show it in red with brackets if the sign is negative. Use -B1.</p>', bridge: 'The calculator you will use in every credit meeting.' },
    starter: '=PMT(', fnKey: '', chips: [], solution: '=PMT(B2/12,B3,-B1)', alts: [],
    rowReply: 'B4 says {got}, but the installment should be {want}.',
    replies: { pass: ['Received, the installment is right. Makasih! 🙏', 'That is Chapter 8. You can price a loan now.'], wrong: '', typed: 'B4 has a typed number, not a formula 😅' },
    after: 'Chapter 9 is macros: recording, variables, loops and buttons. It arrives in the next build.',
    note: { fn: 'PMT in real Excel', what: 'One formula, the calculator you will use at work.', shape: '=PMT(B2/12,B3,-B1)' }
  }), { kind: 'excel' }));

  /* ---------- Chapter 9: Writing macros ----------
     Each request shows a working macro with one line missing. The learner writes the line (or picks one),
     the whole macro runs on the sheet, and the result is compared with the right macro's. */
  const MSHEET = { A1: 'Distributor', B1: 'Outstanding (Rp)', C1: 'Days late', D1: 'Status' };
  macroRows.forEach((r, i) => { const n = i + 2; MSHEET['A' + n] = r.distributor; MSHEET['B' + n] = r.outstanding; MSHEET['C' + n] = r.dpd; });
  const LAST_M = macroRows.length + 1;
  function vba(o) {
    return Object.assign({ chapter: 9, kind: 'vbafix', from: 'lestari', file: 'piutang_senin.xlsm', sheet: MSHEET, sheetCols: ['A', 'B', 'C', 'D', 'E', 'F'], sheetRows: LAST_M,
      nudge: 'Read the lines around the blank: what does the macro need at that point? Then press Run it to see the sheet change.' }, o);
  }
  LIST.push(
    vba({
      id: '9.1', due: '09:00', title: 'Add up with a variable', macroName: 'TotalOutstanding',
      code: ['Sub TotalOutstanding()', '    Dim r As Long, total As Double', '    For r = 2 To ' + LAST_M, '        ___', '    Next r', '    Range("F1").Value = total', 'End Sub'],
      solutionLine: 'total = total + Cells(r, 2).Value', choices: ['total = Cells(r, 2).Value', 'total = total + Cells(r, 2).Value', 'total = total + Cells(r, 3).Value'],
      intro: ['Pagi {name}! Bu Lestari. You read my macro last time. Now <b>write</b> one 💪', 'This one should add up all the <b>outstanding</b> balances into F1. One line is missing.', '@file'],
      ask: 'Write the line that adds each row\'s outstanding to total.',
      lesson: { idea: 'A variable remembers a value while the macro runs.', teach: '<p><code>total = total + Cells(r, 2).Value</code> takes the old total, adds this row\'s number, and stores it back. Run the loop and total grows row by row.</p><p>Without <code>total +</code>, each row would replace the last one.</p>', bridge: 'Like a running balance column, kept in the macro\'s head.' },
      replies: { pass: ['That is the total. My first macro from you 🙏'], wrong: 'F1 does not show the total of every balance 🤔' },
      hints: ['Add this row\'s balance to what total already has.', 'Outstanding is column 2.', 'total = total + Cells(r, 2).Value'],
      note: { fn: 'VBA variables', what: 'A running total: total = total + value, inside a loop.', shape: 'total = total + Cells(r, 2).Value' }
    }),
    vba({
      id: '9.2', due: '09:45', title: 'Loop to the last row', macroName: 'MarkLate',
      code: ['Sub MarkLate()', '    Dim r As Long, lastRow As Long', '    lastRow = Cells(Rows.Count, 1).End(xlUp).Row', '    ___', '        If Cells(r, 3).Value > 30 Then', '            Cells(r, 4).Value = "LATE"', '        Else', '            Cells(r, 4).Value = "OK"', '        End If', '    Next r', 'End Sub'],
      solutionLine: 'For r = 2 To lastRow', choices: ['For r = 2 To 5', 'For r = 2 To lastRow', 'For r = 1 To lastRow'],
      intro: ['This one marks late distributors. The loop line is missing.', 'It must go from the first distributor to the <b>last row</b>, however many there are next week.', '@file'],
      ask: 'Write the For line: from row 2 to the last row.',
      lesson: { idea: 'For … To repeats lines, once per row.', teach: '<p><code>For r = 2 To lastRow</code> runs the lines down to <code>Next r</code> with r = 2, 3, 4 … up to lastRow. Row 1 is the header, so start at 2.</p><p>lastRow is found by jumping up from the bottom of column A, so it follows the data.</p>', bridge: 'Like going down a batch voucher by voucher.' },
      replies: { pass: ['Every row marked, from top to bottom ✅'], wrong: 'Hmm, some rows are missing a status, or the header changed 😅' },
      hints: ['Start after the header, end at the last row.', 'The header is row 1. The last row is in lastRow.', 'For r = 2 To lastRow'],
      note: { fn: 'For loops', what: 'For r = 2 To lastRow … Next r: once per data row.', shape: 'For r = 2 To lastRow' }
    }),
    vba({
      id: '9.3', due: '10:30', title: 'Decide with If', macroName: 'FlagKol3',
      code: ['Sub FlagKol3()', '    Dim r As Long', '    For r = 2 To ' + LAST_M, '        ___', '            Cells(r, 4).Value = "KOL 3+"', '        Else', '            Cells(r, 4).Value = "OK"', '        End If', '    Next r', 'End Sub'],
      solutionLine: 'If Cells(r, 3).Value > 90 Then', choices: ['If Cells(r, 3).Value > 90 Then', 'If Cells(r, 2).Value > 90 Then', 'If Cells(r, 3).Value < 90 Then'],
      intro: ['Citra wants loans <b>more than 90 days late</b> marked KOL 3+.', 'The If line is missing.', '@file'],
      ask: 'Write the If line: days late (column 3) more than 90.',
      lesson: { idea: 'If … Then decides what happens on each row.', teach: '<p><code>If Cells(r, 3).Value &gt; 90 Then</code> runs the next lines only when the test is true. <code>Else</code> runs when it is false. <code>End If</code> closes it.</p>', bridge: 'Like an approval rule applied to each file.' },
      replies: { pass: ['Exactly the KOL 3+ list 👍'], wrong: 'The wrong loans are marked 🤔 It is <b>days late</b>, more than 90.' },
      hints: ['Test days late, column 3.', 'More than 90 is > 90.', 'If Cells(r, 3).Value > 90 Then'],
      note: { fn: 'If … Then … Else', what: 'Runs lines only when a test is true; Else covers the rest.', shape: 'If Cells(r, 3).Value > 90 Then' }
    }),
    vba({
      id: '9.4', due: '11:15', title: 'Write text with &', macroName: 'ReviewNotes',
      code: ['Sub ReviewNotes()', '    Dim r As Long', '    For r = 2 To ' + LAST_M, '        If Cells(r, 3).Value > 30 Then', '            ___', '        End If', '    Next r', 'End Sub'],
      solutionLine: 'Cells(r, 4).Value = Cells(r, 1).Value & " - REVIEW"', choices: ['Cells(r, 4).Value = "REVIEW"', 'Cells(r, 4).Value = Cells(r, 1).Value & " - REVIEW"', 'Cells(r, 1).Value = Cells(r, 1).Value & " - REVIEW"'],
      intro: ['For the late ones, column D should say the <b>name</b> and "- REVIEW", like <b>CV Kepodang Niaga - REVIEW</b>.', '@file'],
      ask: 'Write the line that puts the name & " - REVIEW" in column 4.',
      lesson: { idea: '& joins text in VBA too.', teach: '<p><code>Cells(r, 1).Value &amp; " - REVIEW"</code> joins the name with the words. Put the result in column 4 so the name in column 1 stays clean.</p>', bridge: 'Like writing a note next to each exception.' },
      replies: { pass: ['Clear notes, and the names untouched 📝'], wrong: 'Hmm, column D does not look right, or the names changed 😅' },
      hints: ['Column 4 gets the name joined with " - REVIEW".', 'Join with &: Cells(r, 1).Value & " - REVIEW".', 'Cells(r, 4).Value = Cells(r, 1).Value & " - REVIEW"'],
      note: { fn: 'Joining text in VBA', what: '& joins texts. Write to another column so the source stays clean.', shape: 'Cells(r, 4).Value = Cells(r, 1).Value & " - REVIEW"' }
    }),
    vba({
      id: '9.5', due: '13:00', title: 'Loop until a blank with Do While', macroName: 'CountList',
      code: ['Sub CountList()', '    Dim r As Long, n As Long', '    r = 2', '    ___', '        n = n + 1', '        r = r + 1', '    Loop', '    MsgBox n & " distributors on the list"', 'End Sub'],
      solutionLine: 'Do While Cells(r, 1).Value <> ""', choices: ['Do While r < 5', 'Do While Cells(r, 1).Value <> ""', 'Do While n < 10'],
      intro: ['Count the distributors on the list: keep going <b>until column A is empty</b>.', '@file'],
      ask: 'Write the Do While line: keep going while column A is not empty.',
      lesson: { idea: 'Do While repeats until a test turns false.', teach: '<p><code>Do While Cells(r, 1).Value &lt;&gt; ""</code> keeps looping while the cell is not empty. <code>&lt;&gt;</code> means not equal, and <code>""</code> is empty.</p><p>Always move r forward inside, or the loop never ends.</p>', bridge: 'Like reading a list until the first blank line.' },
      replies: { pass: ['Right count, and it will work when the list grows 👍'], wrong: 'The count in the message box is off 🤔' },
      hints: ['Stop at the first empty cell in column A.', 'Not empty is <> "".', 'Do While Cells(r, 1).Value <> ""'],
      note: { fn: 'Do While … Loop', what: 'Repeats while a test is true. Move forward inside the loop.', shape: 'Do While Cells(r, 1).Value <> ""' }
    }),
    vba({
      id: '9.6', due: '13:45', title: 'Color the risky rows', macroName: 'ColorRisk',
      code: ['Sub ColorRisk()', '    Dim r As Long', '    For r = 2 To ' + LAST_M, '        If Cells(r, 3).Value > 90 Then', '            ___', '        End If', '    Next r', 'End Sub'],
      solutionLine: 'Cells(r, 1).Interior.Color = vbYellow', choices: ['Cells(r, 1).Interior.Color = vbYellow', 'Cells(r, 3).Value = vbYellow', 'Cells(1, 1).Interior.Color = vbYellow'],
      intro: ['Make the <b>name</b> of every loan more than 90 days late stand out in <b>yellow</b>.', '@file'],
      ask: 'Write the line that colors column 1 yellow on that row.',
      lesson: { idea: '.Interior.Color sets a cell\'s fill.', teach: '<p><code>Cells(r, 1).Interior.Color = vbYellow</code> colors the fill. <code>.Value</code> changes the content, <code>.Interior</code> the background.</p>', bridge: 'Like highlighting with a marker, row by row.' },
      replies: { pass: ['Now I can see them at a glance 🟨'], wrong: 'The wrong cells changed 😅 Color the <b>name</b> on that row.' },
      hints: ['Color, not value. Column 1, this row.', 'Use .Interior.Color with vbYellow.', 'Cells(r, 1).Interior.Color = vbYellow'],
      note: { fn: 'Interior.Color', what: '.Interior.Color sets the fill; .Value sets the content.', shape: 'Cells(r, 1).Interior.Color = vbYellow' }
    })
  );
  LIST.push({
    id: '9.7', chapter: 9, kind: 'excel', valuesOnly: true, from: 'lestari', due: '15:30', title: 'Run a macro in real Excel', file: 'piutang_senin.xlsx',
    cells: MSHEET, cols: ['A', 'B', 'C', 'D'], num: ['B', 'C'], dateCols: [], first: 2, last: LAST_M, answerCell: 'D2', fill: { col: 'D', first: 2, last: LAST_M },
    wantValues: macroRows.map((r) => (r.dpd > 30 ? 'LATE' : 'OK')), solution: '', rowReply: 'Row {row} says {got}, but {key} should be {want}.', keyCol: 'A',
    intro: ['Last one, in <b>real Excel</b> 💻', 'Download the list, add the <b>MarkLate</b> macro (LATE above 30 days, otherwise OK in column D), run it, save and send it back.', '@file'],
    ask: 'In Excel: run MarkLate so column D says LATE or OK, then upload.',
    lesson: { idea: 'A macro in real Excel.', teach: '<p>Windows: Alt+F11. Mac: Tools, Macro, Visual Basic Editor. Insert, Module, paste the macro, run it with F5.</p><p>Save as .xlsx: the macro is dropped, its results stay, and that is what Kicau checks. Keep an .xlsm copy to reuse it.</p>', bridge: 'Your Monday job, done in one click.' },
    replies: { pass: ['Received, every status is right. Terima kasih! 🙏', 'That is Chapter 9. You can automate the boring parts now.'], wrong: '', typed: '' },
    after: 'Chapter 10 is the final report: join Rasa Kenari\'s falling sales with Bank Kenari\'s days late into an early-warning list. It arrives in the next build.',
    note: { fn: 'Macros in real Excel', what: 'VBA editor, Module, F5. Save .xlsm to keep the macro.', shape: 'Sub MarkLate() … End Sub' }
  });

  /* ---------- Chapter 10: The early-warning report ----------
     Both companies on one sheet: Rasa Kenari's net sales per distributor in May and October, and Bank Kenari's
     days late for the same distributors. A distributor is a warning when sales fell more than 10% AND it is more
     than 30 days late. The three planted in data.js since week one are the ones it finds. */
  const EW = D.earlyWarning();
  const E_FIRST = 5, E_LAST = E_FIRST + EW.length - 1;
  const ER = (c) => c + E_FIRST + ':' + c + E_LAST;
  const ECOLS = ['A', 'B', 'C', 'D', 'E', 'F', 'G'];
  const RULE = '(' + ER('C') + '/' + ER('B') + '-1<-0.1)*(' + ER('E') + '>30)';
  function ewSheet(top, ans, alt) {
    const c = { A4: 'Distributor', B4: 'Net May (Rp)', C4: 'Net Oct (Rp)', D4: 'Days late May', E4: 'Days late Oct' };
    if (ans) c.F4 = ans;
    EW.forEach((e, i) => { const n = E_FIRST + i; c['A' + n] = e.distributor; c['B' + n] = e.net_may; c['C' + n] = e.net_oct; c['D' + n] = e.dpd_may; c['E' + n] = e.dpd_oct; });
    /* Kedasi Prima is late but growing: late alone is not a warning. */
    c['E' + (E_FIRST + EW.findIndex((e) => e.distributor === 'UD Kedasi Prima'))] = 40;
    if (alt) { const t = EW.findIndex((e) => e.distributor === 'CV Nuri Makmur'); c['E' + (E_FIRST + t)] = 35; c['C' + (E_FIRST + 2)] = Math.round(c['C' + (E_FIRST + 2)] * 0.5); }
    return Object.assign(c, top);
  }
  function ew(o) {
    const top = typeof o.top === 'function' ? o.top : () => (o.top || {});
    const fill = o.ansLabel ? { col: 'F', first: E_FIRST, last: E_LAST } : null;
    const spill = o.spillRows ? { anchor: 'G' + E_FIRST, cols: 1, rows: o.spillRows } : null;
    const base = ewSheet(top('base'), o.ansLabel, false), alt = ewSheet(top('alt'), o.ansLabel, true);
    return formula(Object.assign({
      chapter: 10, file: 'early_warning.xlsx', cols: ECOLS, num: ['B', 'C', 'D', 'E'], dateCols: [], first: E_FIRST, last: E_LAST,
      fill, spill, answerCell: spill ? 'G' + E_FIRST : fill ? 'F' + E_FIRST : 'B2',
      colRange: Object.fromEntries(ECOLS.map((c) => [c, ER(c)])), variantCells: { base, alt }, cells: base
    }, o, { top: undefined }));
  }
  LIST.push(
    ew({
      id: '10.1', from: 'citra', due: '09:00', title: 'Sales change, May to October', ansLabel: 'Sales change',
      intro: ['This is the big one 🏦🤝🏭', 'Rasa Kenari shared its sales with us. For each distributor: how much did net sales <b>change</b> from May to October, as a %?', '@file', 'Column F, every row. A fall shows as negative.'],
      ask: 'October ÷ May − 1, every row, in column F.',
      lesson: { idea: 'A change in % is new ÷ old − 1.', teach: '<p><code>C5/B5-1</code>: −0.35 means sales fell 35%. Format it as a percentage in Excel.</p>', ex: [['=C5', 'October'], ['/B5-1', 'over May, minus 1.']], exFormula: '=C5/B5-1', exLabel: 'for the first distributor', bridge: 'Like a variance analysis: actual over prior, minus one.' },
      starter: '=', fnKey: 'C5/B5-1', chips: ['C5', '/B5', '-1'],
      solution: '=C5/B5-1', alts: ['=(C5-B5)/B5'],
      traps: [{ code: 'flipped', formula: '=B5/C5-1', title: 'Upside down', plain: 'New over old: October on top.', reply: 'Srigunting\'s sales <b>grew</b> 400%? 😅 October goes on top.' }, { code: 'rupiah', formula: '=C5-B5', title: 'That is rupiah', plain: 'That is the change in rupiah. Divide by May for a %.', reply: 'That is the change in rupiah. I need it as a <b>%</b> of May.' }],
      replies: { pass: ['Clear. Three of them fell hard 📉'], frozen: 'October sales were corrected for one distributor and your % did not move.', typed: 'Typed? 🙏' },
      note: { fn: '% change', what: 'new ÷ old − 1. Negative means a fall.', shape: '=C5/B5-1' },
      hints: ['New divided by old, minus 1.', 'October is C5, May is B5.', 'It looks like =C5/B5-1.']
    }),
    ew({
      id: '10.2', from: 'dimas', due: '09:45', title: 'Days late, the change', ansLabel: 'Days late change',
      intro: ['Now our side.', 'How many <b>more days late</b> is each distributor in October than in May?', '@file'],
      ask: 'October days late − May days late, every row.',
      lesson: { idea: 'A change in days is a plain subtraction.', teach: '<p><code>E5-D5</code>: positive means they slipped further behind.</p>', ex: [['=E5', 'October days late'], ['-D5', 'minus May.']], exFormula: '=E5-D5', exLabel: 'days for the first distributor', bridge: 'Like the movement column in an aging roll-forward.' },
      starter: '=', fnKey: 'E5-D5', chips: ['E5', '-D5'],
      solution: '=E5-D5', alts: ['=-(D5-E5)'],
      traps: [{ code: 'flipped', formula: '=D5-E5', title: 'Upside down', plain: 'October minus May.', reply: 'They got <b>better</b> by 95 days? 😅 October minus May.' }],
      replies: { pass: ['Kepodang, Pipit, Srigunting slipped badly. Kedasi Prima too, a little 😬'], frozen: 'Two distributors slipped this morning and your column did not move.', typed: 'Typed? 🙏' },
      note: { fn: 'Change in days', what: 'October − May. Positive means worse.', shape: '=E5-D5' },
      hints: ['October minus May.', 'E5 minus D5.', 'It looks like =E5-D5.']
    }),
    ew({
      id: '10.3', from: 'citra', due: '10:30', title: 'The warning flag', ansLabel: 'Flag',
      intro: ['Our rule:', 'A distributor is a <b>WARNING</b> when sales fell <b>more than 10%</b> AND they are <b>more than 30 days late</b> in October. Otherwise <b>OK</b>.', '@file'],
      ask: 'WARNING if sales change < −10% AND October days late > 30, else OK.',
      lesson: { idea: 'Put everything together: IF, AND, a % change and a comparison.', teach: '<p><code>IF(AND(C5/B5-1&lt;-0.1, E5&gt;30), "WARNING", "OK")</code>. Both signals must agree: one alone could be noise.</p>', ex: [['=AND(', 'Both true:'], ['C5/B5-1<-0.1,', 'sales fell more than 10%,'], ['E5>30)', 'and more than 30 days late?']], exFormula: '=AND(C5/B5-1<-0.1,E5>30)', exLabel: 'for the first distributor', bridge: 'Like an audit red flag that needs two indicators.' },
      starter: '=IF(AND(', fnKey: 'IF(AND(', chips: ['C5/B5-1<-0.1', ',E5>30)', ',"WARNING","OK")'],
      solution: '=IF(AND(C5/B5-1<-0.1,E5>30),"WARNING","OK")', alts: ['=IF(AND(C5/B5<0.9,E5>30),"WARNING","OK")', '=IF(OR(C5/B5-1>=-0.1,E5<=30),"OK","WARNING")'],
      traps: [{ code: 'or', formula: '=IF(OR(C5/B5-1<-0.1,E5>30),"WARNING","OK")', title: 'OR flags too many', plain: 'With OR, one signal alone is a warning. The rule needs both.', reply: 'CV Tekukur Rejeki and UD Kedasi Prima are warnings? One has falling sales but pays on time, the other is late but growing 🤔 It needs <b>both</b>: AND.' }, { code: 'onlylate', formula: '=IF(E5>30,"WARNING","OK")', title: 'Only one signal', plain: 'Late alone is not enough. Add the sales test with AND.', reply: 'UD Kedasi Prima is late, but their sales grew 26%. Not a warning yet. It needs the <b>sales fall too</b>.' }],
      replies: { pass: ['WARNING, WARNING, WARNING. Kepodang, Pipit and Srigunting 😬', 'That is exactly what we were afraid of.'], frozen: 'CV Nuri Makmur slipped to 35 days late this morning, and their sales fell 19%. They should now be a WARNING, but your flag did not change. Read both signals from the row.', typed: 'Typed flags will be wrong next month 🙏' },
      note: { fn: 'Early-warning flag', what: 'IF(AND(sales fell > 10%, days late > 30), "WARNING", "OK").', shape: '=IF(AND(C5/B5-1<-0.1,E5>30),"WARNING","OK")' },
      hints: ['Two tests, both must be true.', 'C5/B5-1<-0.1 and E5>30, inside AND, inside IF.', 'It looks like =IF(AND(C5/B5-1<-0.1,E5>30),"WARNING","OK").']
    }),
    ew({
      id: '10.4', from: 'dimas', due: '11:15', title: 'Count the warnings in one cell', top: { A2: 'Warnings' },
      intro: ['For the committee\'s first slide:', 'How many distributors are warnings? <b>One formula in B2</b>, without a helper column.', '@file'],
      ask: 'The number of warnings, in one formula, in <b>B2</b>.',
      lesson: { idea: 'SUMPRODUCT can count rows that pass several tests at once.', teach: '<p>Each test gives TRUE or FALSE per row. Multiplying two lists gives 1 only where both are true, and SUMPRODUCT adds them.</p><p><code>SUMPRODUCT((test 1)*(test 2))</code>.</p>', ex: [['=SUMPRODUCT(', 'Add up'], ['(' + ER('E') + '>30)*1)', '1 for each late distributor.']], exFormula: '=SUMPRODUCT((' + ER('E') + '>30)*1)', exLabel: 'distributors are more than 30 days late', bridge: 'Like a count of exceptions that meet two criteria.' },
      starter: '=SUMPRODUCT(', fnKey: 'SUMPRODUCT(', chips: ['(' + ER('C') + '/' + ER('B') + '-1<-0.1)', '*(' + ER('E') + '>30)', ')'],
      solution: '=SUMPRODUCT(' + RULE + ')', alts: ['=ROWS(FILTER(' + ER('A') + ',' + RULE + '))', '=SUM(' + RULE + ')'],
      traps: [{ code: 'onlylate', formula: '=COUNTIF(' + ER('E') + ',">30")', title: 'Only one signal', plain: 'That counts late distributors only. Add the sales test.', reply: '{value}? That counts everyone late. A warning needs the <b>sales fall too</b>.' }],
      replies: { pass: ['{value} warnings. Slide one, done 📊'], frozen: 'Two distributors changed this morning and the count did not.', typed: 'Typed? 🙏' },
      note: { fn: 'Counting with two tests', what: 'SUMPRODUCT((test 1)*(test 2)) counts the rows where both are true.', shape: '=SUMPRODUCT((C5:C18/B5:B18-1<-0.1)*(E5:E18>30))' },
      hints: ['Multiply two lists of TRUE/FALSE and add them.', '(C5:C18/B5:B18-1<-0.1)*(E5:E18>30) inside SUMPRODUCT.', 'It looks like =SUMPRODUCT((C5:C18/B5:B18-1<-0.1)*(E5:E18>30)).']
    }),
    ew({
      id: '10.5', from: 'anton', due: '13:00', title: 'List them with FILTER', spillRows: 14,
      intro: ['Anton here. Citra told me 😟', 'I need the <b>names</b> of the warning distributors, as a live list. I visit them first.', '@file', 'One formula in G5. "None" if the list is ever empty.'],
      ask: 'FILTER the names of the warning distributors into G5, "None" if empty.',
      lesson: { idea: 'The same two tests, used to filter a list.', teach: '<p><code>FILTER(names, (test 1)*(test 2), "None")</code> spills only the names that pass both tests.</p>', ex: [['=FILTER(' + ER('A') + ',', 'The names'], ['(' + ER('E') + '>30),"None")', 'of those late.']], exFormula: '=ROWS(FILTER(' + ER('A') + ',(' + ER('E') + '>30),"None"))', exLabel: 'names', bridge: 'Your watch list, always up to date.' },
      starter: '=FILTER(', fnKey: 'FILTER(', chips: [ER('A'), ',' + RULE, ',"None")'],
      solution: '=FILTER(' + ER('A') + ',' + RULE + ',"None")', alts: ['=FILTER(' + ER('A') + ',' + RULE + ')'],
      traps: [{ code: 'onlylate', formula: '=FILTER(' + ER('A') + ',' + ER('E') + '>30,"None")', title: 'Only one signal', plain: 'Add the sales test with *.', reply: 'That is everyone late. I visit only the ones whose <b>sales fell too</b>.' }],
      replies: { pass: ['Three names. I am booking flights to Semarang, Malang and Palembang ✈️'], frozen: 'The data changed and the list did not.', typed: 'Typed names will not update 🙏' },
      note: { fn: 'FILTER with two tests', what: 'FILTER(names, (test 1)*(test 2), "None"): a live watch list.', shape: '=FILTER(A5:A18,(C5:C18/B5:B18-1<-0.1)*(E5:E18>30),"None")' },
      hints: ['FILTER the names with both tests.', 'Names A5:A18; tests multiplied; "None" as the third part.', 'It looks like =FILTER(A5:A18,(C5:C18/B5:B18-1<-0.1)*(E5:E18>30),"None").']
    })
  );
  LIST.push(Object.assign(ew({
    id: '10.6', from: 'citra', due: '16:00', title: 'The final report, in real Excel', ansLabel: 'Flag',
    intro: ['Last one. The real report, in <b>real Excel</b> 💻', 'Download the sheet, put the WARNING / OK flag in column F with one formula copied down, save and send it to the committee.', '@file'],
    ask: 'In Excel: the warning flag in F5:F18, then upload.',
    lesson: { idea: 'Everything you learned, in one file.', teach: '<p>One formula in F5, the fill handle, save, upload. Add conditional formatting for WARNING if you like: it will survive the upload.</p>', bridge: 'This is the report a real analyst sends.' },
    starter: '', fnKey: '', chips: [], solution: '=IF(AND(C5/B5-1<-0.1,E5>30),"WARNING","OK")', alts: [],
    rowReply: 'Row {row} says {got}, but {key} should be {want}.', keyCol: 'A',
    replies: { pass: ['Received. The committee has it. Terima kasih 🙏', 'You started with SUM. You just warned a bank about three borrowers before they defaulted.'], wrong: '', typed: 'The flags are typed in, not a formula 😅 Next month they will be wrong.' },
    note: { fn: 'The early-warning report', what: 'Sales from one company, days late from another, one rule, one formula.', shape: '=IF(AND(C5/B5-1<-0.1,E5>30),"WARNING","OK")' }
  }), { kind: 'excel', final: true }));

  /* ---------- Chapter 11: Real cases, across workbooks ----------
     Year-end at Rasa Kenari and Bank Kenari. Real work arrives in more than one file: the sales lines, the
     product master, the targets, the bank's loan book and the shipping rate card are separate workbooks, and a
     formula reaches into the others as Excel writes it: [master_produk.xlsx]Produk!$A$5:$D$16.
     Each sheet below has a base version and a next-day version. A request picks its own sheet (`main`, where
     the answer goes) and the other workbooks it has open (`open`). The engine gets the other sheets through
     the SHEETS key; the app draws a tab for every open workbook and sheet. */
  const SHEETS = Symbol.for('kenari.sheets'), MAIN = Symbol.for('kenari.main');
  const DB = D.build();
  const quote = (name) => (/^[A-Za-z_][A-Za-z0-9_]*$/.test(name) ? name : "'" + name.replace(/'/g, "''") + "'");
  /** How a formula in `fromBook` names a range on another sheet, the way Excel writes it. */
  const sheetRef = (fromBook, book, sheet) => (book === fromBook ? quote(sheet) + '!' : (/[^A-Za-z0-9_.]/.test(book + sheet) ? "'[" + book + ']' + sheet + "'!" : '[' + book + ']' + sheet + '!'));
  const engineName = (mainBook, book, sheet) => (book === mainBook ? sheet : book + '|' + sheet);

  const DIST11 = DB.distributors.map((d) => d.name);
  const REGIONS = ['Jabodetabek', 'Jawa Barat', 'Jawa Tengah', 'Jawa Timur', 'Bali Nusra', 'Sumatera', 'Sulawesi'];
  const REPS = ['Rudi', 'Wulan', 'Agus', 'Sinta'];
  /* Q4 achievement per distributor (net ÷ target). Exactly 1.2, 1 and 0.8 sit on the tier edges on purpose. */
  const ACH = [1.32, 1.2, 1.08, 1, 0.95, 0.86, 0.8, 0.74, 0.6, 1.15, 0.9, 1.02, 0.79, 1.25];
  const TARGET = [180, 150, 120, 140, 100, 90, 110, 80, 75, 160, 95, 130, 70, 105].map((m) => m * 1000000);

  const SHEET11 = {
    /* Product master: one row per SKU, and the category discounts on a second sheet. */
    Produk: { book: 'master_produk.xlsx', cols: ['A', 'B', 'C', 'D'], num: ['D'], first: 5, last: 16, build: (alt) => {
      const c = { A4: 'SKU', B4: 'Product', C4: 'Category', D4: 'Carton price (Rp)' };
      D.PRODUCTS.forEach((p, i) => { const n = 5 + i; c['A' + n] = p.sku; c['B' + n] = p.name; c['C' + n] = p.category; c['D' + n] = p.carton_price + (alt && i % 4 === 1 ? 6000 : 0); });
      return c;
    } },
    Diskon: { book: 'master_produk.xlsx', cols: ['A', 'B'], num: ['B'], first: 5, last: 8, build: (alt) => {
      const c = { A4: 'Category', B4: 'Discount %' };
      [['Noodles', 2], ['Coffee', alt ? 4 : 3], ['Sambal', alt ? 1 : 0], ['Soap', 5]].forEach(([k, v], i) => { c['A' + (5 + i)] = k; c['B' + (5 + i)] = v; });
      return c;
    } },
    /* Q4 targets by distributor, with the area and the sales rep. */
    Target: { book: 'target_2026.xlsx', cols: ['A', 'B', 'C', 'D'], num: ['D'], first: 5, last: 18, build: (alt) => {
      const c = { A4: 'Distributor', B4: 'Area', C4: 'Sales rep', D4: 'Q4 target (Rp)' };
      DB.distributors.forEach((d, i) => { const n = 5 + i; c['A' + n] = d.name; c['B' + n] = d.region; c['C' + n] = REPS[i % 4]; c['D' + n] = TARGET[i] + (alt && i === 4 ? 10000000 : 0); });
      return c;
    } },
    /* Q4 net sales per distributor, the summary everyone works from. */
    Ringkasan: { book: 'ringkasan_q4.xlsx', cols: ['A', 'B', 'C'], num: ['B'], first: 5, last: 18, build: (alt) => {
      const c = { A4: 'Distributor', B4: 'Net Q4 (Rp)' };
      DB.distributors.forEach((d, i) => { const n = 5 + i; c['A' + n] = d.name; c['B' + n] = Math.round(TARGET[i] * ACH[i]) + (alt && i === 2 ? 15000000 : 0) - (alt && i === 7 ? 6000000 : 0); });
      return c;
    } },
    /* Bank Kenari's loan book, one facility per distributor. */
    Pinjaman: { book: 'buku_pinjaman.xlsx', cols: ['A', 'B', 'C', 'D', 'E', 'F', 'G'], num: ['C', 'D', 'E'], first: 5, last: 18, build: (alt) => {
      const c = { A4: 'Facility', B4: 'Distributor', C4: 'Limit (Rp)', D4: 'Outstanding (Rp)', E4: 'Days late', F4: 'Restructured', G4: 'Region' };
      const DPD = [0, 0, 45, 0, 95, 15, 130, 0, 200, 0, 75, 0, 35, 150];
      const RES = ['N', 'N', 'N', 'Y', 'N', 'N', 'Y', 'N', 'N', 'N', 'Y', 'N', 'N', 'Y'];
      const USE = [0.66, 0.46, 0.92, 0.55, 0.97, 0.71, 0.88, 0.3, 0.99, 0.62, 0.81, 0.97, 0.58, 0.9];
      DB.distributors.forEach((d, i) => {
        const n = 5 + i, f = DB.facilities.find((x) => x.distributor === d.name);
        const limit = f ? f.limit : 75000000;
        c['A' + n] = 'F' + String(101 + i); c['B' + n] = d.name; c['C' + n] = limit; c['D' + n] = Math.round(limit * USE[i] / 100000) * 100000;
        c['E' + n] = DPD[i] + (alt && i === 5 ? 50 : 0) + (alt && i === 0 ? 31 : 0); c['F' + n] = RES[i]; c['G' + n] = d.region;
      });
      return c;
    } },
    /* Shipping rate card: rupiah per carton, by region down and category across. */
    Tarif: { book: 'tarif_kirim.xlsx', cols: ['A', 'B', 'C', 'D', 'E'], num: ['B', 'C', 'D', 'E'], first: 5, last: 11, build: (alt) => {
      const c = { A4: 'Region', B4: 'Noodles', C4: 'Coffee', D4: 'Sambal', E4: 'Soap' };
      const base = [1500, 2000, 2500, 2500, 4500, 5000, 6500];
      REGIONS.forEach((r, i) => { const n = 5 + i; c['A' + n] = r; ['B', 'C', 'D', 'E'].forEach((col, k) => { c[col + n] = base[i] + k * 500 + (alt && i >= 5 ? 750 : 0); }); });
      return c;
    } },
    /* Q4 sales lines: date, distributor, SKU, cartons. */
    Penjualan: { book: 'penjualan_q4.xlsx', cols: ['A', 'B', 'C', 'D', 'E'], num: ['D'], dateCols: ['A'], first: 5, last: 22, build: (alt) => {
      const c = { A4: 'Date', B4: 'Distributor', C4: 'SKU', D4: 'Cartons' };
      const picks = [];
      for (let i = 0; picks.length < 18; i += 7) { const sale = DB.sales[(i * 3) % DB.sales.length]; if (!picks.some((x) => x.invoice === sale.invoice)) picks.push(sale); }
      picks.sort((a, b) => ((+a.date.slice(5, 7) - 5) % 3 * 100 + +a.date.slice(8, 10)) - ((+b.date.slice(5, 7) - 5) % 3 * 100 + +b.date.slice(8, 10)));
      picks.forEach((sale, i) => { const n = 5 + i; c['A' + n] = D.toSerial('2026-' + String(10 + ((+sale.date.slice(5, 7) - 5) % 3)) + '-' + sale.date.slice(8, 10)); c['B' + n] = sale.distributor; c['C' + n] = sale.sku; c['D' + n] = sale.cartons; });
      if (alt) { c.C7 = c.C7 === 'SRW-SER' ? 'MLK-GOR' : 'SRW-SER'; c.D9 += 12; c.C14 = c.C14 === 'KSJ-ARE' ? 'SNY-IJO' : 'KSJ-ARE'; }
      return c;
    } },
    /* Bank Kenari's late-fee rules, a second sheet in the loan book's workbook. */
    Aturan: { book: 'buku_pinjaman.xlsx', cols: ['A', 'B'], num: ['B'], first: 5, last: 7, build: (alt) => {
      return { A4: 'Rule', B4: 'Value', A5: 'Grace days', B5: alt ? 5 : 7, A6: 'Fee per day late (%)', B6: 0.1, A7: 'Fee cap (Rp)', B7: alt ? 4000000 : 5000000 };
    } },
    /* A blank memo: the inputs go in rows 1 to 2 and the answer in B3 or below. */
    Memo: { book: 'memo_risiko.xlsx', cols: ['A', 'B', 'C', 'D'], num: [], first: 5, last: 5, build: () => ({}) },
    /* Orders waiting for credit approval. */
    Order: { book: 'persetujuan_order.xlsx', cols: ['A', 'B', 'C'], num: ['B'], first: 5, last: 14, build: (alt) => {
      const c = { A4: 'Distributor', B4: 'Order value (Rp)' };
      [0, 1, 2, 3, 4, 5, 7, 9, 11, 12].forEach((di, i) => { c['A' + (5 + i)] = DB.distributors[di].name; c['B' + (5 + i)] = [12, 8, 4, 30, 6, 9, 15, 20, 2, 11][i] * 1000000 + (alt && i === 1 ? 25000000 : 0); });
      return c;
    } }
  };

  /** Where a sheet's data starts and ends, as an absolute range Excel would write from `fromBook`. */
  const range11 = (fromBook, name, col) => { const S = SHEET11[name]; return sheetRef(fromBook, S.book, name) + '$' + col + '$' + S.first + ':$' + col + '$' + S.last; };
  const table11 = (fromBook, name, c1, c2) => { const S = SHEET11[name]; return sheetRef(fromBook, S.book, name) + '$' + c1 + '$' + S.first + ':$' + c2 + '$' + S.last; };

  /**
   * A Chapter 11 request. `main` is the sheet the answer goes on, `open` the other sheets it can see.
   * `ansLabel` makes it a column request, typed in the first row and filled down; otherwise the answer is a
   * single cell (`answerCell`, default B2) with `top` filling rows 1 to 3.
   */
  function real(o) {
    const M = SHEET11[o.main];
    const top = typeof o.top === 'function' ? o.top : () => (o.top || {});
    const ansCol = o.ansCol || M.cols[M.cols.length - 1];
    const fill = o.ansLabel ? { col: ansCol, first: M.first, last: M.last } : null;
    const mainCols = fill && !M.cols.includes(ansCol) ? M.cols.concat(ansCol) : (o.cols || M.cols);
    const make = (v) => {
      const c = Object.assign(M.build(v === 'alt'), top(v));
      if (o.ansLabel) c[ansCol + '4'] = o.ansLabel;
      if (o.edit) o.edit(c, v);
      c[MAIN] = o.main;
      c[SHEETS] = (o.open || []).map((name) => { const S = SHEET11[name]; const sc = S.build(v === 'alt'); if (o.editOpen) o.editOpen(name, sc, v); return { name: engineName(M.book, S.book, name), cells: sc }; });
      return c;
    };
    const base = make('base'), alt = make('alt');
    const books = [];
    for (const name of [o.main].concat(o.open || [])) {
      const S = SHEET11[name];
      let b = books.find((x) => x.file === S.book);
      if (!b) { b = { file: S.book, sheets: [] }; books.push(b); }
      b.sheets.push({ name, cols: name === o.main ? mainCols : S.cols, num: S.num, dateCols: S.dateCols || [], first: S.first, last: S.last, prefix: sheetRef(M.book, S.book, name), main: name === o.main });
    }
    const cols = mainCols;
    return formula(Object.assign({
      chapter: 11, file: M.book, sheetName: o.main, books, cols, num: M.num, dateCols: M.dateCols || [], first: M.first, last: M.last,
      fill, answerCell: fill ? fill.col + M.first : (o.answerCell || 'B2'),
      colRange: Object.fromEntries(M.cols.map((c) => [c, c + M.first + ':' + c + M.last])),
      variantCells: { base, alt }, cells: base
    }, o, { top: undefined, main: undefined, open: undefined, edit: undefined, editOpen: undefined }));
  }

  const PROD = (fromBook) => table11(fromBook, 'Produk', 'A', 'D');
  LIST.push(
    real({
      id: '11.1', from: 'bayu', due: '08:30', title: 'Prices from another workbook', main: 'Penjualan', open: ['Produk'], ansLabel: 'Value (Rp)',
      intro: ['Pagi {name} 👷 Year-end, so everything comes in separate files now.', 'The Q4 sales lines are in <b>penjualan_q4.xlsx</b>, and the carton prices live in <b>master_produk.xlsx</b>. Both are open.', '@file', 'Column E: what is each line worth? Cartons × the carton price from the product master.'],
      ask: 'Cartons × the carton price from master_produk.xlsx, every line, in column E.',
      lesson: {
        idea: 'A formula can read another open workbook. Excel writes the file name in square brackets.',
        teach: '<p>Start the formula in this file, then tap the other workbook\'s tab and pick the range. Excel writes it as <code>[master_produk.xlsx]Produk!$A$5:$D$16</code>: the file, the sheet, then the cells, already pinned with $.</p><p>If the other file is closed, Excel adds its folder too. The formula still works, and Excel asks to update the link when you open the file.</p>',
        ex: [['=VLOOKUP("KSJ-ARE",', 'Look for this SKU'], [PROD('penjualan_q4.xlsx') + ',', 'in the product master, another workbook,'], ['4,', 'bring back column 4, the carton price,'], ['FALSE)', 'exact match only.']],
        exFormula: '=VLOOKUP("KSJ-ARE",' + PROD('penjualan_q4.xlsx') + ',4,FALSE)', exLabel: 'rupiah per carton of Kopi Sekejap Gula Aren',
        bridge: 'Like a journal that takes its prices from a separate price list file instead of copying them in.'
      },
      starter: '=D5*VLOOKUP(', fnKey: 'VLOOKUP(', chips: ['D5*', 'C5', PROD('penjualan_q4.xlsx'), ',4', ',FALSE)'],
      solution: '=D5*VLOOKUP(C5,' + PROD('penjualan_q4.xlsx') + ',4,FALSE)',
      alts: ['=VLOOKUP(C5,' + PROD('penjualan_q4.xlsx') + ',4,FALSE)*D5', '=D5*XLOOKUP(C5,' + range11('penjualan_q4.xlsx', 'Produk', 'A') + ',' + range11('penjualan_q4.xlsx', 'Produk', 'D') + ')', '=D5*INDEX(' + range11('penjualan_q4.xlsx', 'Produk', 'D') + ',MATCH(C5,' + range11('penjualan_q4.xlsx', 'Produk', 'A') + ',0))'],
      traps: [
        { code: 'price', formula: '=VLOOKUP(C5,' + PROD('penjualan_q4.xlsx') + ',4,FALSE)', title: 'That is the price of one carton', plain: 'Multiply the carton price by the cartons in D5.', reply: 'Rp {value} for the whole line? That is one carton 😄 Times the <b>cartons</b> in column D please.' },
        { code: 'col', formula: '=D5*VLOOKUP(C5,' + PROD('penjualan_q4.xlsx') + ',3,FALSE)', title: 'Column 3 is the category', plain: 'Count the columns of the product master from A: SKU 1, Product 2, Category 3, Carton price 4.', reply: 'Hmm, #VALUE! everywhere. Column 3 of the master is the <b>category</b>, a word. The price is column <b>4</b>.' },
        { code: 'nodollar', formula: '=D5*VLOOKUP(C5,[master_produk.xlsx]Produk!A5:D16,4,FALSE)', title: 'The product master slid down', plain: 'Copied down, A5:D16 became A6:D17 and so on, and the first products fell out of the table. Keep the $ Excel adds: $A$5:$D$16.', reply: 'Some lines say #N/A further down 😅 When you copy it, the range in the other file slides too. Keep the <b>$</b>.' }
      ],
      replies: { pass: ['Mantap! Every line valued, straight from the master file 💰', 'When marketing changes a price, I only fix it in one place now.'], frozen: 'Marketing updated three prices in the master this morning and your values stayed the same.', typed: 'Typed values will not follow the master file 🙏' },
      note: { fn: 'Another workbook', what: 'Refer to an open workbook by tapping it. Excel writes [file.xlsx]Sheet!$A$1, pinned with $.', shape: '=D5*VLOOKUP(C5,[master_produk.xlsx]Produk!$A$5:$D$16,4,FALSE)' },
      hints: ['Cartons in D5 times the carton price of the SKU in C5. The price is in the other workbook.', 'Tap the master_produk.xlsx tab: the table is $A$5:$D$16 and the price is its 4th column.', 'It looks like =D5*VLOOKUP(C5,[master_produk.xlsx]Produk!$A$5:$D$16,4,FALSE).']
    })
  );
  /* Requests 11.2 to 11.10: each written, then solved blind and checked by a second reviewer against Excel. */
  LIST.push(
    (() => {
      const P = PROD('penjualan_q4.xlsx');
      const DK = table11('penjualan_q4.xlsx', 'Diskon', 'A', 'B');
      const PA = range11('penjualan_q4.xlsx', 'Produk', 'A'), PC = range11('penjualan_q4.xlsx', 'Produk', 'C'), PD = range11('penjualan_q4.xlsx', 'Produk', 'D');
      const DA = range11('penjualan_q4.xlsx', 'Diskon', 'A'), DBC = range11('penjualan_q4.xlsx', 'Diskon', 'B');
      const PRICE = 'VLOOKUP(C5,' + P + ',4,FALSE)';
      const CAT = 'VLOOKUP(C5,' + P + ',3,FALSE)';
      const DISC = 'VLOOKUP(' + CAT + ',' + DK + ',2,FALSE)';
      const SOL = '=D5*' + PRICE + '*(1-' + DISC + '/100)';
      return real({
        id: '11.2', from: 'tika', due: '09:15', title: 'Net after the category discount', main: 'Penjualan', open: ['Produk', 'Diskon'], ansLabel: 'Net (Rp)',
        intro: [
          'Pagi {name}! 🙏 Tika here, closing Q4 sales.',
          'Cartons × carton price is only the gross. My report needs the <b>net</b>, after the category discount.',
          'The discounts are on the <b>Diskon</b> sheet of master_produk.xlsx, one row per category. A 3 there means 3%.',
          '@file',
          'The sales lines only have a SKU. On <b>Produk</b>, the SKU gives the category (3rd column) and the carton price (4th). On <b>Diskon</b>, the category gives the discount (2nd column).',
          'Column E, every line: <b>cartons × carton price × (1 − discount % ÷ 100)</b> 😊'
        ],
        ask: 'Net (Rp) = cartons × carton price × (1 − the category\'s discount % ÷ 100), every line, in column E.',
        lesson: {
          idea: 'A lookup can feed another lookup: the answer of the first one is what the second one looks for.',
          teach: '<p>A sales line only knows its SKU. <code>VLOOKUP(C5,' + P + ',3,FALSE)</code> turns the SKU into its category.</p>' +
            '<p>Put that whole lookup where the second VLOOKUP expects the value to find: <code>VLOOKUP(VLOOKUP(C5,…,3,FALSE),' + DK + ',2,FALSE)</code>. Excel works from the inside out: SKU to category, category to discount.</p>' +
            '<p>Diskon stores 3 for 3%, so divide by 100. Keep the brackets so the discount comes off the whole line: <code>D5*price*(1-discount/100)</code>.</p>',
          ex: [['=VLOOKUP(', 'Look for'], ['VLOOKUP("KSJ-ARE",' + P + ',3,FALSE),', 'the category of this SKU,'], [DK + ',', 'in the discount table,'], ['2,FALSE)', 'and bring back column 2, the discount %.']],
          exFormula: '=VLOOKUP(VLOOKUP("KSJ-ARE",' + P + ',3,FALSE),' + DK + ',2,FALSE)', exLabel: '% discount on Kopi Sekejap Gula Aren, a Coffee',
          bridge: 'Like a two-step mapping in the ledger: the item code gives the product group, and the group gives its trade discount rate.'
        },
        starter: '=D5*VLOOKUP(', fnKey: 'VLOOKUP(', chips: ['D5*', 'C5', P, ',4', ',3', ',FALSE)', '*(1-', DK, ',2', '/100)'],
        solution: SOL,
        alts: [
          '=' + PRICE + '*D5*(1-' + DISC + '/100)',
          '=D5*' + PRICE + '*(100-' + DISC + ')/100',
          '=D5*VLOOKUP(C5,' + P + ',4,0)*(1-VLOOKUP(VLOOKUP(C5,' + P + ',3,0),' + DK + ',2,0)/100)',
          '=D5*XLOOKUP(C5,' + PA + ',' + PD + ')*(1-XLOOKUP(XLOOKUP(C5,' + PA + ',' + PC + '),' + DA + ',' + DBC + ')/100)',
          '=D5*INDEX(' + PD + ',MATCH(C5,' + PA + ',0))*(1-INDEX(' + DBC + ',MATCH(INDEX(' + PC + ',MATCH(C5,' + PA + ',0)),' + DA + ',0))/100)',
          '=LET(price,' + PRICE + ',cat,' + CAT + ',disc,VLOOKUP(cat,' + DK + ',2,FALSE),D5*price*(1-disc/100))'
        ],
        traps: [
          { code: 'nodiv', formula: '=D5*' + PRICE + '*(1-' + DISC + ')', title: 'The % was not divided by 100', plain: 'Diskon stores 3 for 3%. 1 − 3 is −2, so the lines turn negative. Divide the discount by 100 first: (1-…/100).', reply: 'Rp {value}? A minus? 😅 The Diskon sheet says <b>3</b> for 3%, so it needs <b>÷ 100</b> before you take it off.' },
          { code: 'gross', formula: '=D5*' + PRICE, title: 'That is the gross, before the discount', plain: 'Cartons × carton price is the gross. Take the category discount off: ×(1 − discount % ÷ 100).', reply: 'Rp {value} is the gross 😊 Net means <b>after</b> the category discount from the Diskon sheet.' },
          { code: 'sku', formula: '=D5*' + PRICE + '*(1-VLOOKUP(C5,' + DK + ',2,FALSE)/100)', title: 'Diskon found no match', plain: 'Diskon lists categories (Noodles, Coffee, Sambal, Soap), not SKU codes. Find the category first with VLOOKUP(C5,…,3,FALSE), then look that up in Diskon. End every VLOOKUP with FALSE so it matches exactly.', reply: '#N/A 😬 The Diskon sheet only knows <b>categories</b> like Coffee, not SKU codes. Is the category coming from Produk first?' },
          { code: 'brackets', formula: '=D5*' + PRICE + '*1-' + DISC + '/100', title: 'The brackets went missing', plain: 'Without brackets, Excel multiplies first and then subtracts the discount ÷ 100, a few hundredths of a rupiah, from the total. Wrap it: *(1-…/100).', reply: 'Rp {value}, the gross minus a few hundredths of a rupiah? 🤔 Put <b>(1 − discount ÷ 100)</b> in brackets so the whole line gets the discount.' }
        ],
        replies: { pass: ['Mantap! The first line nets Rp {value} after the 3% Coffee discount 💰', 'When marketing moves a discount in the master, my net moves with it. Makasih! 🙏'], frozen: 'Marketing moved Coffee to 4% and Sambal to 1% this morning, and your net did not follow 🤔', typed: 'Typed values will not follow the discounts 🙏' },
        note: { fn: 'Lookup inside a lookup', what: 'The inner VLOOKUP turns the SKU into a category. The outer one turns the category into its discount.', shape: SOL },
        hints: [
          'Three things multiply: the cartons, the carton price, and (1 − discount ÷ 100). The discount needs the category first.',
          'The price is VLOOKUP(C5,' + P + ',4,FALSE) and the category is the same lookup with 3. Look the category up in ' + DK + ', column 2, then ÷ 100.',
          'It looks like ' + SOL + '.'
        ]
      });
    })(),
    (() => {
      const BK = 'penjualan_q4.xlsx';
      const T = table11(BK, 'Pinjaman', 'A', 'B');
      const T7 = table11(BK, 'Pinjaman', 'A', 'G');
      const FA = range11(BK, 'Pinjaman', 'A'), FB = range11(BK, 'Pinjaman', 'B');
      const LOOK = 'VLOOKUP(B1,' + T + ',2,FALSE)';
      return real({
        id: '11.3', from: 'citra', due: '10:15', title: 'A borrower\'s real sales', main: 'Penjualan', open: ['Pinjaman'],
        top: (v) => ({ A1: 'Facility', B1: v === 'alt' ? 'F114' : 'F111', A2: 'Cartons in Q4' }),
        /* Next day only: line 5 is rebooked to UD Kedasi Prima, the F114 borrower, so its lines sit on row 5 and row 22. A range that drops the first or last row then misses cartons. */
        edit: (c, v) => { if (v === 'alt') c.B5 = 'UD Kedasi Prima'; },
        intro: [
          'Pagi {name}! Annual review season at the bank 🏦',
          'Before we renew a facility, I check that the borrower really sells.',
          'Rasa Kenari shared their Q4 sales lines, <b>penjualan_q4.xlsx</b>. Our loan book, <b>buku_pinjaman.xlsx</b>, is open too.',
          '@file',
          'I type a facility number in <b>B1</b>. The sales lines only know distributor names, so the loan book tells you who owns that facility.',
          'In <b>B2</b>: the total <b>cartons</b> in the sales lines for that distributor. One formula, so I can change B1 and read the next borrower.'
        ],
        ask: 'Total cartons in the sales lines for the distributor that owns the facility in B1, in <b>B2</b>.',
        lesson: {
          idea: 'The criteria of a SUMIF can be another formula. A VLOOKUP turns the facility number into the name SUMIF needs.',
          teach: '<p><code>SUMIF(B5:B22, name, D5:D22)</code> adds the cartons on the rows where column B shows that name. Typed in, the name answers one borrower only.</p><p>Now put a lookup where the name goes: <code>=SUMIF(names, VLOOKUP(…), cartons)</code>. The VLOOKUP reads the facility in B1 and brings back column 2 of the loan book, the distributor. Excel works from the inside out: first the VLOOKUP finds the name, then SUMIF adds the cartons for it.</p>',
          ex: [['=SUMIF(B5:B22,', 'In the distributor column of the sales lines,'], ['"UD Kedasi Prima",', 'find this name,'], ['D5:D22)', 'and add the cartons on those rows.']],
          exFormula: '=SUMIF(B5:B22,"UD Kedasi Prima",D5:D22)', exLabel: 'cartons for UD Kedasi Prima, with the name typed in',
          bridge: 'Like tracing a loan account number to the customer name, then footing that customer in the sales subledger.'
        },
        starter: '=SUMIF(', fnKey: 'SUMIF(', chips: ['B5:B22,', 'VLOOKUP(B1,', T, ',2,FALSE)', ',D5:D22)'],
        solution: '=SUMIF(B5:B22,' + LOOK + ',D5:D22)',
        alts: [
          '=SUMIF($B$5:$B$22,VLOOKUP($B$1,' + T7 + ',2,FALSE),$D$5:$D$22)',
          '=SUMIFS(D5:D22,B5:B22,' + LOOK + ')',
          '=SUMIF(B5:B22,XLOOKUP(B1,' + FA + ',' + FB + '),D5:D22)',
          '=SUMIF(B5:B22,INDEX(' + FB + ',MATCH(B1,' + FA + ',0)),D5:D22)',
          '=LET(owner,' + LOOK + ',SUMIF(B5:B22,owner,D5:D22))',
          '=SUMPRODUCT((B5:B22=' + LOOK + ')*D5:D22)',
          '=SUM(FILTER(D5:D22,B5:B22=' + LOOK + '))'
        ],
        frozenTry: '=SUMIF(B5:B22,"CV Bangau Sejahtera",D5:D22)',
        traps: [
          { code: 'col', formula: '=SUMIF(B5:B22,VLOOKUP(B1,' + T7 + ',3,FALSE),D5:D22)', title: 'SUMIF found nothing to add', plain: 'A total of 0 means no line matched, or nothing was added. The criteria has to be a distributor name: in the loan book, Facility is column 1 and Distributor is column 2. Then check the order: names B5:B22 first, cartons D5:D22 last.', reply: '{value} cartons? I nearly called their director 😅 Check two things. The criteria is a distributor <b>name</b>, column <b>2</b> of the loan book. And the cartons, <b>D5:D22</b>, go last.' },
          { code: 'dates', formula: '=SUMIF(B5:B22,' + LOOK + ',A5:A22)', title: 'Those are dates added up', plain: 'The third part of SUMIF is what gets added. Column A holds the dates. The cartons are in D5:D22.', reply: '{value} cartons? Their warehouse is not that big 😄 You added the <b>dates</b> in column A. Cartons are column <b>D</b>.' },
          { code: 'count', formula: '=COUNTIF(B5:B22,' + LOOK + ')', title: 'That counts the lines', plain: 'COUNTIF counts how many sales lines match. Citra needs the cartons on those lines added up: SUMIF, with D5:D22 as the third part.', reply: '{value} cartons for a whole quarter? 😅 That is how many <b>sales lines</b> they had. I need the <b>cartons</b> on those lines, added up.' },
          { code: 'name', formula: '=' + LOOK, title: 'That is the borrower, not the cartons', plain: 'This VLOOKUP finds who owns the facility, which is the first step. Now use it as the criteria of a SUMIF that adds their cartons.', reply: 'Yes, {value} owns F111 👍 That is half of it. Now I need their <b>cartons</b> added up, with this name as the SUMIF criteria.' }
        ],
        replies: {
          pass: ['Restructured, 75 days late, and still {value} cartons in Q4. They sell. The problem is cash, not demand 🏦', 'Now I type any facility in B1 and the answer follows. Terima kasih.'],
          frozen: 'I put F114 in B1 and got the wrong total 🤔 A typed name works for one borrower only. Let a VLOOKUP read B1, and add up every sales line, rows 5 to 22.',
          typed: 'A typed total is right for one borrower only 🙏 I need a formula that reads B1.'
        },
        note: { fn: 'SUMIF with a lookup inside', what: 'VLOOKUP turns a code into a name, and SUMIF adds the rows for that name. Excel works from the inside out.', shape: '=SUMIF(B5:B22,VLOOKUP(B1,[buku_pinjaman.xlsx]Pinjaman!$A$5:$B$18,2,FALSE),D5:D22)' },
        hints: [
          'Two steps in one formula: find the distributor that owns the facility in B1, then add their cartons.',
          'The name comes from VLOOKUP(B1,[buku_pinjaman.xlsx]Pinjaman!$A$5:$B$18,2,FALSE). Use it as the criteria of a SUMIF over B5:B22 that adds D5:D22.',
          'It looks like =SUMIF(B5:B22,VLOOKUP(B1,[buku_pinjaman.xlsx]Pinjaman!$A$5:$B$18,2,FALSE),D5:D22).'
        ]
      });
    })(),
    (() => {
      const BOOK = 'ringkasan_q4.xlsx';
      const T = table11(BOOK, 'Target', 'A', 'D');
      const ACH = 'B5/VLOOKUP(A5,' + T + ',4,FALSE)';
      const FLIP = 'VLOOKUP(A5,' + T + ',4,FALSE)/B5';
      const tiers = (a, ge) => 'IF(' + a + ge + '1.2,5%,IF(' + a + ge + '1,3%,IF(' + a + ge + '0.8,1%,0)))';
      const SOLUTION = '=B5*' + tiers(ACH, '>=');
      return real({
        id: '11.4', from: 'anton', due: '11:45', title: 'Commission tiers with nested IF', main: 'Ringkasan', open: ['Target'], ansLabel: 'Commission (Rp)',
        intro: [
          'Siang {name}. Anton.',
          'Q4 commission for the distributors. Net Q4 is in <b>ringkasan_q4.xlsx</b>, the Q4 targets are in <b>target_2026.xlsx</b>. Both open.',
          '@file',
          '<b>Achievement</b> = Net Q4 ÷ that distributor\'s Q4 target.',
          'Rate: at least <b>120%</b> = 5% · at least <b>100%</b> = 3% · at least <b>80%</b> = 1% · below 80% = 0.',
          '<b>Commission = Net Q4 × rate</b>, in rupiah. Column C, every distributor.',
          'Three of them landed exactly on a line. On the line counts.'
        ],
        ask: 'Commission in column C: Net Q4 × 5%, 3%, 1% or 0, for an achievement (Net Q4 ÷ Q4 target) of at least 120%, at least 100%, at least 80%, or below 80%.',
        lesson: {
          idea: 'A nested IF picks a tier. Test the highest tier first, because the first true test wins.',
          teach: '<p><b>1. The achievement.</b> <code>B5/VLOOKUP(A5,' + T + ',4,FALSE)</code>: this file\'s net over the other file\'s target.</p><p><b>2. The rate, top tier first.</b> <code>IF(ach&gt;=1.2, 5%, IF(ach&gt;=1, 3%, IF(ach&gt;=0.8, 1%, 0)))</code>, with the achievement where it says ach. A distributor at 130% passes the first test and stops there. Start from 80% instead and everyone above it stops at 1%. "At least" means <code>&gt;=</code>, so exactly 100% earns 3%.</p><p><b>3. The rupiah.</b> Put <code>B5*</code> in front of the IF. If the formula feels long, <code>LET(ach, …, …)</code> names the achievement once.</p>',
          ex: [['=50000000*', 'Rp 50 million of sales, times'], ['IF(0.95>=1.2,5%,', 'at least 120%? No, go on,'], ['IF(0.95>=1,3%,', 'at least 100%? No, go on,'], ['IF(0.95>=0.8,1%,0)))', 'at least 80%? Yes: 1%.']],
          exFormula: '=50000000*IF(0.95>=1.2,5%,IF(0.95>=1,3%,IF(0.95>=0.8,1%,0)))', exLabel: 'rupiah of commission on Rp 50 million at 95% of target',
          bridge: 'Like accruing a sales bonus: find the band from achievement, then apply its rate to the sales.'
        },
        starter: '=B5*IF(', fnKey: 'IF(', chips: [ACH, '>=1.2,5%,', '>=1,3%,', '>=0.8,1%,0)))', T],
        solution: SOLUTION,
        alts: [
          '=LET(ach,' + ACH + ',B5*IF(ach>=1.2,5%,IF(ach>=1,3%,IF(ach>=0.8,1%,0))))',
          '=B5*IFS(' + ACH + '>=1.2,5%,' + ACH + '>=1,3%,' + ACH + '>=0.8,1%,TRUE,0)',
          '=LET(ach,' + ACH + ',B5*IFS(ach>=1.2,5%,ach>=1,3%,ach>=0.8,1%,TRUE,0))',
          '=IF(' + ACH + '>=1.2,B5*5%,IF(' + ACH + '>=1,B5*3%,IF(' + ACH + '>=0.8,B5*1%,0)))',
          '=B5*IF(' + ACH + '>=1.2,0.05,IF(' + ACH + '>=1,0.03,IF(' + ACH + '>=0.8,0.01,0)))',
          '=B5*IF(' + ACH + '<0.8,0,IF(' + ACH + '<1,1%,IF(' + ACH + '<1.2,3%,5%)))',
          '=LET(ach,B5/XLOOKUP(A5,' + range11(BOOK, 'Target', 'A') + ',' + range11(BOOK, 'Target', 'D') + '),B5*IF(ach>=1.2,5%,IF(ach>=1,3%,IF(ach>=0.8,1%,0))))',
          '=LET(ach,B5/INDEX(' + range11(BOOK, 'Target', 'D') + ',MATCH(A5,' + range11(BOOK, 'Target', 'A') + ',0)),B5*IF(ach>=1.2,5%,IF(ach>=1,3%,IF(ach>=0.8,1%,0))))',
          '=B5*IF(' + ACH + '>=120%,5%,IF(' + ACH + '>=100%,3%,IF(' + ACH + '>=80%,1%,0)))',
          '=B5*XLOOKUP(' + ACH + ',{0,0.8,1,1.2},{0,0.01,0.03,0.05},,-1)'
        ],
        traps: [
          { code: 'order', formula: '=B5*IF(' + ACH + '>=0.8,1%,IF(' + ACH + '>=1,3%,IF(' + ACH + '>=1.2,5%,0)))', title: 'Everyone from 80% up gets 1%', plain: 'IF stops at the first test that is true. 132% is also at least 80%, so testing 80% first gives everyone above it 1%. Test the highest tier first: 120%, then 100%, then 80%.', reply: DIST11[0] + ' hit 132% and gets 1%? My best distributor. Check the top tier first: <b>120%</b>, then 100%, then 80%.' },
          { code: 'edge', formula: '=B5*' + tiers(ACH, '>'), title: 'On the line counts', plain: '"At least 120%" includes exactly 120%. With > the three distributors on the line drop a tier. Use >= in every test.', reply: DIST11[1] + ' hit exactly 120% and gets 3%. The scheme says <b>at least</b>. Use <b>&gt;=</b>.' },
          { code: 'rate', formula: '=' + tiers(ACH, '>='), title: 'That is the rate, not rupiah', plain: 'The IF gives the rate: 5%, 3%, 1% or 0. Multiply it by Net Q4 in B5 for the commission in rupiah.', reply: 'I see 0.05 and 0.03. That is the rate. I pay in rupiah: times <b>Net Q4</b> in column B.' },
          { code: 'flipped', formula: '=B5*' + tiers(FLIP, '>='), title: 'Achievement upside down', plain: 'Target ÷ net gives the weakest distributor the biggest number. Achievement is Net Q4 in B5 divided by the target.', reply: DIST11[8] + ' made 60% of target and gets the top rate? Achievement is <b>net ÷ target</b>, not the other way.' }
        ],
        replies: {
          pass: ['Good. Eleven distributors earn commission this quarter 💰', DIST11[1] + ', ' + DIST11[3] + ' and ' + DIST11[6] + ' sat exactly on the line, and they are paid right. Terima kasih.'],
          frozen: 'This morning ' + DIST11[2] + ' booked a late Q4 invoice and finance corrected one target. Your column did not move.',
          typed: 'Typed amounts will be wrong when a target changes. Formula please.'
        },
        note: { fn: 'Tiered commission', what: 'Achievement from two workbooks, tiers tested from the top with >=, then the rate × the sales.', shape: SOLUTION },
        hints: [
          'Two steps: the achievement, Net Q4 ÷ the Q4 target from target_2026.xlsx, then the rate for its tier. Commission is B5 × that rate.',
          'The achievement is ' + ACH + '. Test it with >=1.2 first, then >=1, then >=0.8, else 0.',
          'It looks like ' + SOLUTION + '.'
        ]
      });
    })(),
    (() => {
      const FB = 'persetujuan_order.xlsx';
      const T = table11(FB, 'Pinjaman', 'B', 'G');
      const V = (n) => 'VLOOKUP(A5,' + T + ',' + n + ',FALSE)';
      const X = (c) => 'XLOOKUP(A5,' + range11(FB, 'Pinjaman', 'B') + ',' + range11(FB, 'Pinjaman', c) + ')';
      const HOLD = 'OR(' + V(4) + '>60,' + V(3) + '+B5>' + V(2) + ')';
      return real({
        id: '11.5', from: 'lestari', due: '13:30', title: 'Approve, review or hold', main: 'Order', open: ['Pinjaman'], ansLabel: 'Decision',
        /* Row 11: a distributor more than 60 days late whose order still fits its limit, so AND and a missing order give different answers.
           Rows 6 and 14 swap whole orders, so UD Jalak Sentosa (row 6 of the loan book) sits on row 14 and a range without $ slides past it to #N/A. */
        edit: (c) => { c.A11 = 'CV Bangau Sejahtera'; [c.A6, c.B6, c.A14, c.B14] = [c.A14, c.B14, c.A6, c.B6]; },
        intro: [
          'Siang {name}! Bu Lestari from Credit Control.',
          'Ten orders are waiting for my sign-off before the year-end cut-off. They are in <b>persetujuan_order.xlsx</b>.',
          'Our distributors pay through their facility at Bank Kenari, so I check each one against the bank\'s loan book, <b>buku_pinjaman.xlsx</b>. Both are open.',
          '@file',
          'The rule, in this order:',
          '<b>HOLD</b> if the distributor is <b>more than 60 days late</b>, OR its outstanding plus this order would be <b>more than its limit</b>.',
          'Otherwise <b>APPROVE</b> if it is <b>0 days late</b>, and <b>REVIEW</b> for everyone else.',
          'Find each distributor by name in column B of the loan book. The decision goes in column C, in capitals exactly like that please. The approval system reads them.'
        ],
        ask: 'HOLD if more than 60 days late OR outstanding + this order is more than the limit; else APPROVE if 0 days late; else REVIEW. Every order, in column C.',
        lesson: {
          idea: 'One IF can sit inside another, and OR lets either of two tests trigger the first answer.',
          teach: '<p>First the facts. The loan book table is <code>' + T + '</code>. Counting from column B: limit is column 2, outstanding 3, days late 4. Each fact is its own <code>VLOOKUP(A5, table, n, FALSE)</code>.</p>' +
            '<p>Then the rule: <code>IF(OR(test 1, test 2), "HOLD", IF(test 3, "APPROVE", "REVIEW"))</code>. OR is TRUE when either test is true. The inner IF only runs for orders that are not a HOLD, so it never has to check them again.</p>' +
            '<p>Read "more than" as <code>&gt;</code>. An order that lands exactly on the limit is not over it.</p>',
          ex: [
            ['=IF(OR(', 'If either test is true:'],
            [V(4) + '>60,', 'more than 60 days late,'],
            [V(3) + '+B5>' + V(2) + '),', 'or outstanding plus this order is more than the limit,'],
            ['"HOLD","not yet")', 'say HOLD, otherwise not yet.']
          ],
          exFormula: '=IF(' + HOLD + ',"HOLD","not yet")', exLabel: 'for CV Nuri Makmur, so the second IF decides its order',
          bridge: 'Like the credit check before a sales order is released: an overdue account or a blown credit limit blocks it, a clean account goes through.'
        },
        starter: '=IF(OR(', fnKey: 'IF(OR(',
        chips: ['VLOOKUP(A5,', T, ',4,FALSE)', ',3,FALSE)', ',2,FALSE)', '>60,', '+B5>', '),"HOLD",IF(', '=0,"APPROVE","REVIEW"))'],
        solution: '=IF(' + HOLD + ',"HOLD",IF(' + V(4) + '=0,"APPROVE","REVIEW"))',
        alts: [
          '=IFS(' + HOLD + ',"HOLD",' + V(4) + '=0,"APPROVE",TRUE,"REVIEW")',
          '=LET(late,' + V(4) + ',IF(OR(late>60,' + V(3) + '+B5>' + V(2) + '),"HOLD",IF(late=0,"APPROVE","REVIEW")))',
          '=IF(OR(' + X('E') + '>60,' + X('D') + '+B5>' + X('C') + '),"HOLD",IF(' + X('E') + '=0,"APPROVE","REVIEW"))',
          '=IF(OR(' + V(4) + '>60,B5>' + V(2) + '-' + V(3) + '),"HOLD",IF(' + V(4) + '>0,"REVIEW","APPROVE"))'
        ],
        exactCase: true,
        traps: [
          { code: 'and', formula: '=IF(AND(' + V(4) + '>60,' + V(3) + '+B5>' + V(2) + '),"HOLD",IF(' + V(4) + '=0,"APPROVE","REVIEW"))', title: 'AND needs both tests to be true', plain: 'The rule holds an order when either test is true. That is OR, not AND.', reply: 'UD Cucak Berkah would go Rp 1,000,000 over its limit, and it says APPROVE? 😟 And CV Bangau Sejahtera is 75 days late. Either one is enough for a HOLD: <b>OR</b>, not AND.' },
          { code: 'noorder', formula: '=IF(OR(' + V(4) + '>60,' + V(3) + '>' + V(2) + '),"HOLD",IF(' + V(4) + '=0,"APPROVE","REVIEW"))', title: 'The new order is missing', plain: 'The limit test is outstanding plus this order, B5, against the limit.', reply: 'UD Srigunting Utama has Rp 1,500,000 of room left and this order is Rp 2,000,000. Approved? 😟 Add <b>this order</b> in B5 to the outstanding before you compare it with the limit.' },
          { code: 'swap', formula: '=IF(' + HOLD + ',"HOLD",IF(' + V(4) + '=0,"REVIEW","APPROVE"))', title: 'APPROVE and REVIEW are swapped', plain: '0 days late is APPROVE. The ones paying late, but not more than 60 days, get REVIEW.', reply: 'CV Nuri Makmur pays on time, 0 days late, and it goes to REVIEW? 🤔 <b>0 days late</b> is APPROVE. REVIEW is for the ones paying a little late.' },
          { code: 'atlimit', formula: '=IF(OR(' + V(4) + '>60,' + V(3) + '+B5>=' + V(2) + '),"HOLD",IF(' + V(4) + '=0,"APPROVE","REVIEW"))', title: 'Exactly at the limit is not over it', plain: 'The rule says more than the limit. Use > instead of >=.', reply: 'CV Kutilang Abadi lands on exactly Rp 50,000,000, its limit. That is not <b>more than</b> the limit, so it is not a HOLD. Use <b>&gt;</b>, not &gt;=.' }
        ],
        replies: {
          pass: ['Terima kasih. Four on HOLD, three go straight through, and three for me to read 🙏', 'CV Bangau Sejahtera still has room on its facility, but 75 days late is a HOLD. I would have missed that one by eye.'],
          frozen: 'Bank Kenari sent a new loan book this morning. UD Seriti Jaya is now 65 days late and your column still says REVIEW. It has to read the loan book every time.',
          typed: 'Typed decisions will not follow the loan book 🙏 A formula please.'
        },
        note: { fn: 'Nested IF with OR', what: 'IF(OR(test 1, test 2), first answer, IF(test 3, second answer, third answer)). Each fact comes from its own VLOOKUP.', shape: '=IF(OR(VLOOKUP(A5,[buku_pinjaman.xlsx]Pinjaman!$B$5:$G$18,4,FALSE)>60,VLOOKUP(A5,[buku_pinjaman.xlsx]Pinjaman!$B$5:$G$18,3,FALSE)+B5>VLOOKUP(A5,[buku_pinjaman.xlsx]Pinjaman!$B$5:$G$18,2,FALSE)),"HOLD",IF(VLOOKUP(A5,[buku_pinjaman.xlsx]Pinjaman!$B$5:$G$18,4,FALSE)=0,"APPROVE","REVIEW"))' },
        hints: [
          'Three facts from the loan book: limit, outstanding and days late. Look up A5 in $B$5:$G$18. Counting from column B, they are columns 2, 3 and 4.',
          'The HOLD test is OR(days late>60, outstanding+B5>limit). When it is false, a second IF gives APPROVE if days late=0, otherwise REVIEW.',
          'It looks like =IF(OR(' + V(4) + '>60,' + V(3) + '+B5>' + V(2) + '),"HOLD",IF(' + V(4) + '=0,"APPROVE","REVIEW")).'
        ]
      });
    })(),
    (() => {
      const G = 'IFS(E5=0,1,E5<=90,2,E5<=120,3,E5<=180,4,TRUE,5)';
      const GS = 'IFS(E5=0,1,E5<90,2,E5<120,3,E5<180,4,TRUE,5)';
      const NG = 'IF(E5=0,1,IF(E5<=90,2,IF(E5<=120,3,IF(E5<=180,4,5))))';
      return real({
        id: '11.6', from: 'citra', due: '14:00', title: 'Grades with a restructuring floor', main: 'Pinjaman', open: [], ansCol: 'H', ansLabel: 'Kol',
        /* Two facilities sit exactly on an edge (90 and 180 days), so < instead of <= shows on either edge. UD Jalak Sentosa
           carries the 90: no other request cites its days late (CV Kepodang Niaga's 95 is quoted in 11.7 and 11.9).
           Overnight, UD Pipit Karya is restructured, so a formula that singles out F104 is caught. */
        edit: (c, v) => { c.E6 = 90; c.E18 = 180; if (v === 'alt') c.F12 = 'Y'; },
        intro: [
          'Siang {name}. Year-end grading of our own loan book, <b>buku_pinjaman.xlsx</b>.',
          'Same grade as before, by days late in column E:',
          '<b>0 days</b> = 1 · <b>up to 90</b> = 2 · <b>up to 120</b> = 3 · <b>up to 180</b> = 4 · <b>more</b> = 5',
          '@file',
          'Plus Bank Kenari\'s restructuring rule: a facility with <b>Y</b> in column F is <b>never better than grade 2</b>.',
          'Restructured and on time is still a 2. A restructured 4 stays a 4.',
          'Column H, every facility please.'
        ],
        ask: 'Grade 1 to 5 by days late, but at least 2 when Restructured is Y. Every row, in column H.',
        lesson: {
          idea: 'MAX can put a floor under a number: it can go up, never below the floor.',
          teach: '<p><code>MAX(grade, 2)</code> gives the bigger of the two. A 1 becomes 2, a 4 stays 4. In collectibility a bigger grade is worse, so "never better than 2" means "at least 2", and that is MAX. MIN does the opposite: every restructured loan worse than 2 would show as a 2.</p><p>Only restructured facilities get the floor, so wrap it in IF: <code>IF(F5="Y", MAX(grade,2), grade)</code>. The grade is the IFS from Chapter 2, written in both places.</p>',
          ex: [['=MAX(', 'The bigger of'], ['IFS(E8=0,1,E8<=90,2,E8<=120,3,E8<=180,4,TRUE,5),', 'F104\'s grade by days late, a 1,'], ['2)', 'and the floor, 2.']],
          exFormula: '=MAX(IFS(E8=0,1,E8<=90,2,E8<=120,3,E8<=180,4,TRUE,5),2)', exLabel: 'for F104, on time but restructured',
          bridge: 'Like prudence in accounting: when two measures disagree, you book the more cautious one.'
        },
        starter: '=IF(', fnKey: 'MAX(', chips: ['F5="Y"', ',', G, ',2)', ')'],
        solution: '=IF(F5="Y",MAX(' + G + ',2),' + G + ')',
        alts: [
          '=MAX(' + G + ',IF(F5="Y",2,1))',
          '=LET(g,' + G + ',IF(F5="Y",MAX(g,2),g))',
          '=IF(F5="Y",MAX(' + NG + ',2),' + NG + ')',
          '=IF(F5="N",' + G + ',MAX(2,' + G + '))',
          '=IF(AND(F5="Y",E5<=90),2,' + G + ')',
          '=IFS(E5=0,IF(F5="Y",2,1),E5<=90,2,E5<=120,3,E5<=180,4,TRUE,5)',
          '=MAX(IF(F5="Y",2,1),XLOOKUP(E5,{0,1,91,121,181},{1,2,3,4,5},,-1))'
        ],
        traps: [
          { code: 'noflag', formula: '=' + G, title: 'The restructuring rule is missing', plain: 'By days late F104 is a 1, since it pays on time. It is restructured (Y in column F), so Bank Kenari\'s rule lifts it to 2. Wrap the grade: IF(F5="Y", MAX(grade,2), grade).', reply: 'F104 is a 1? It pays on time, but it is <b>restructured</b>. Our rule: never better than 2.' },
          { code: 'min', formula: '=IF(F5="Y",MIN(' + G + ',2),' + G + ')', title: 'MIN made the bad loans look better', plain: 'MIN keeps the smaller number, so F107 and F114, both grade 4, dropped to 2. Never better than 2 means at least 2: the bigger number, MAX.', reply: 'F107 is 130 days late and you made it a 2? 😳 Restructuring never makes a loan look <b>better</b>. MAX, not MIN.' },
          { code: 'strict', formula: '=IF(F5="Y",MAX(' + GS + ',2),' + GS + ')', title: 'Exactly 90 days is grade 2', plain: 'Up to 90 includes 90, and up to 180 includes 180. Use <= on every edge, not <.', reply: 'F102 is exactly 90 days late, and up to 90 is grade 2. F114 at 180 is a 4, not a 5. <b>&lt;=</b> on the edges please.' },
          { code: 'flat', formula: '=IF(F5="Y",2,' + G + ')', title: 'A floor, not a fixed grade', plain: 'The rule puts a floor under the grade: at least 2. F107 and F114 are grade 4 by days late and stay 4. Compare the grade with 2 and keep the bigger one.', reply: 'All four restructured loans are a 2? F114 is 180 days late, that is a 4 🤔 Our rule is a <b>floor</b>. A restructured loan can be worse than 2.' }
        ],
        replies: {
          pass: ['Clean. F104 pays on time and still reads 2, as it should ✅', 'This column goes straight into the year-end review pack. Terima kasih.'],
          frozen: 'Overnight CV Nuri Makmur slipped to 31 days late and UD Pipit Karya was restructured. Your grades did not move.',
          typed: 'Typed grades will not follow tomorrow\'s loan book 😅 Formula please.'
        },
        note: { fn: 'MAX as a floor', what: 'MAX(grade, 2) is never below 2. Inside IF, only the rows that need the floor get it.', shape: '=IF(F5="Y",MAX(' + G + ',2),' + G + ')' },
        hints: [
          'First the grade by days late, the IFS from before. Then the restructuring rule on top.',
          'If F5 is "Y", take MAX of the grade and 2. Otherwise the grade as it is.',
          'It looks like =IF(F5="Y",MAX(' + G + ',2),' + G + ').'
        ]
      });
    })(),
    (() => {
      const P = (c) => range11('memo_risiko.xlsx', 'Pinjaman', c);
      return real({
        id: '11.7', from: 'dimas', due: '14:30', title: 'Late exposure by region, from a cell', main: 'Memo', open: ['Pinjaman'], answerCell: 'B3', frozenTry: '=SUMIFS(' + P('D') + ',' + P('G') + ',B1,' + P('E') + ',">30")',
        top: (v) => ({ A1: 'Region', B1: v === 'alt' ? 'Sulawesi' : 'Jawa Tengah', A2: 'Days late over', B2: v === 'alt' ? 35 : 30, A3: 'Outstanding (Rp)' }),
        intro: [
          'Siang {name}! Dimas here, working on the year-end risk memo 📝',
          'The loan book is <b>buku_pinjaman.xlsx</b> and my memo is <b>memo_risiko.xlsx</b>. Both are open.',
          '@file',
          'The memo has a region in <b>B1</b> and a days-late cut-off in <b>B2</b>.',
          'In <b>B3</b> I need the total <b>Outstanding</b> of the facilities in that region that are <b>more than B2 days late</b>. Exactly B2 days does not count.',
          'I will change B1 and B2 for every region in the memo, so read both from the cells 🙏'
        ],
        ask: 'Total Outstanding in the region in <b>B1</b>, facilities more than <b>B2</b> days late, in <b>B3</b>.',
        lesson: {
          idea: 'To use a cell inside a condition, join the sign to the cell with &.',
          teach: '<p>A condition with a sign is text in quotes: <code>"&gt;30"</code> works. <code>"&gt;B2"</code> does not: inside the quotes, B2 is two letters, not a cell. Excel compares the days late with the text "B2" and finds nothing. And <code>B2</code> with no sign asks for exactly 30 days.</p><p>Keep the sign in quotes and join the cell outside them with &amp;: <code>"&gt;"&amp;B2</code>. Excel reads B2 first, so with 30 in B2 the condition becomes <code>"&gt;30"</code>. Change B2 and the condition follows.</p><p>One catch across workbooks: SUMIFS, SUMIF and COUNTIFS need the other file <b>open</b>. If buku_pinjaman.xlsx is closed, they show #VALUE! the next time Excel recalculates. SUMPRODUCT works either way.</p>',
          ex: [['=COUNTIFS(', 'Count the facilities'], [P('E') + ',', 'in the loan book whose days late'], ['">"&B2)', 'are more than the number in B2.']],
          exFormula: '=COUNTIFS(' + P('E') + ',">"&B2)', exLabel: 'facilities more than B2 days late, every region',
          bridge: 'Like an aging report with the cut-off in one cell: change the cell and the whole report follows.'
        },
        starter: '=SUMIFS(', fnKey: 'SUMIFS(', chips: [P('D'), ',' + P('G'), ',B1', ',' + P('E'), ',">"', '&B2', ')'],
        solution: '=SUMIFS(' + P('D') + ',' + P('G') + ',B1,' + P('E') + ',">"&B2)',
        alts: [
          '=SUMIFS(' + P('D') + ',' + P('E') + ',">"&B2,' + P('G') + ',B1)',
          '=SUMIFS(' + P('D') + ',' + P('G') + ',$B$1,' + P('E') + ',">"&$B$2)',
          '=SUMPRODUCT((' + P('G') + '=B1)*(' + P('E') + '>B2)*' + P('D') + ')',
          '=SUM(FILTER(' + P('D') + ',(' + P('G') + '=B1)*(' + P('E') + '>B2),0))',
          '=LET(late,">"&B2,SUMIFS(' + P('D') + ',' + P('G') + ',B1,' + P('E') + ',late))'
        ],
        traps: [
          { code: 'quoted', formula: '=SUMIFS(' + P('D') + ',' + P('G') + ',B1,' + P('E') + ',">B2")', title: 'No facility passed the conditions', plain: 'A total of 0 means nothing matched. Inside quotes, ">B2" compares the days late with the letters B2, and B2 with no sign asks for exactly the number in B2. Keep the sign in quotes and join the cell outside: ">"&B2.', reply: 'Rp {value}? CV Kepodang Niaga is 95 days late, so it should count 😅 Check the days-late condition. Inside quotes, B2 is only two letters.' },
          { code: 'region', formula: '=SUMIF(' + P('G') + ',B1,' + P('D') + ')', title: 'Every facility in the region', plain: 'With one condition, every facility in the region is in the total, late or not. Use SUMIFS and add the days-late pair.', reply: 'Rp {value} is all of Jawa Tengah. UD Seriti Jaya is only 15 days late. I need the ones <b>more than B2 days late</b>.' },
          { code: 'allregions', formula: '=SUMIF(' + P('E') + ',">"&B2,' + P('D') + ')', title: 'Every region', plain: 'That is every late facility in the loan book. Use SUMIFS and add the region pair, so only the region in B1 counts.', reply: 'Rp {value}? That is the late book for the whole country 😄 Only the region in <b>B1</b> please.' },
          { code: 'limit', formula: '=SUMIFS(' + P('C') + ',' + P('G') + ',B1,' + P('E') + ',">"&B2)', title: 'That is the limit, not the outstanding', plain: 'Column C is the limit the bank agreed to. What they owe today is Outstanding, column D. The first range in SUMIFS is the one it adds.', reply: 'Rp {value} is the <b>limit</b>. The risk is what they still owe: <b>Outstanding</b>, column D.' }
        ],
        replies: {
          pass: ['Rp {value} more than 30 days late in Jawa Tengah. Mantap, that goes in the memo 📝', 'Now I change two cells and the next region is done. Makasih!'],
          frozen: 'I switched B1 to Sulawesi and B2 to 35, and your total went wrong 😅 CV Kasuari Sentral is exactly 35 days late, so it stays out. Read both cells, and use <b>"&gt;"</b>, not "&gt;=".',
          typed: 'Typed in? I change B1 and B2 for every region, so it has to follow them 🙏'
        },
        note: { fn: 'A condition from a cell', what: 'Join the sign and the cell: ">"&B2. Inside quotes, B2 is only text. SUMIFS on another workbook needs that file open.', shape: '=SUMIFS([buku_pinjaman.xlsx]Pinjaman!$D$5:$D$18,[buku_pinjaman.xlsx]Pinjaman!$G$5:$G$18,B1,[buku_pinjaman.xlsx]Pinjaman!$E$5:$E$18,">"&B2)' },
        hints: [
          'Two conditions, region and days late, so SUMIFS. It adds Outstanding, column D of the loan book.',
          'Region is column G, compared with B1. Days late is column E, compared with ">"&B2, the sign in quotes joined to the cell.',
          'It looks like =SUMIFS([buku_pinjaman.xlsx]Pinjaman!$D$5:$D$18,[buku_pinjaman.xlsx]Pinjaman!$G$5:$G$18,B1,[buku_pinjaman.xlsx]Pinjaman!$E$5:$E$18,">"&B2).'
        ]
      });
    })(),
    (() => {
      const BK = 'penjualan_q4.xlsx';
      const TBL = table11(BK, 'Tarif', 'B', 'E');
      const REG = range11(BK, 'Tarif', 'A');
      const HDR = sheetRef(BK, SHEET11.Tarif.book, 'Tarif') + '$B$4:$E$4';
      const TGT = table11(BK, 'Target', 'A', 'D');
      const PRD = PROD(BK);
      const AREA = 'VLOOKUP(B5,' + TGT + ',2,FALSE)';
      const CAT = 'VLOOKUP(C5,' + PRD + ',3,FALSE)';
      const ROW = 'MATCH(' + AREA + ',' + REG + ',0)';
      const COL = 'MATCH(' + CAT + ',' + HDR + ',0)';
      const SOL = '=D5*INDEX(' + TBL + ',' + ROW + ',' + COL + ')';
      return real({
        id: '11.8', from: 'bayu', due: '15:45', title: 'Shipping from the rate card', main: 'Penjualan', open: ['Target', 'Produk', 'Tarif'], ansLabel: 'Shipping (Rp)',
        intro: [
          'Sore {name} 👷 The shipping company sent its new rate card, <b>tarif_kirim.xlsx</b>.',
          'It is rupiah per carton. The <b>regions</b> run down the side (A5:A11) and the <b>categories</b> across the top (B4:E4).',
          'The sales lines show neither one. The distributor\'s region is its <b>Area</b> in target_2026.xlsx. The SKU\'s <b>category</b> is in master_produk.xlsx.',
          '@file',
          'Column E: the shipping cost of every line. Cartons × the rate where that Area and that category cross.'
        ],
        ask: 'Cartons × the rate per carton for the distributor\'s Area (row) and the SKU\'s category (column) in tarif_kirim.xlsx, every line, in column E.',
        lesson: {
          idea: 'A rate card has two keys. One MATCH finds the row, a second MATCH finds the column, and INDEX reads the box where they cross.',
          teach: '<p><b>1. Find the row.</b> The sales line has no region, so look it up first: <code>VLOOKUP(B5,[target_2026.xlsx]Target!$A$5:$D$18,2,FALSE)</code> gives the Area. Then <code>MATCH(that Area,[tarif_kirim.xlsx]Tarif!$A$5:$A$11,0)</code> says which row of the rate card it is. The first line is UD Murai Niaga, in Jawa Barat: row 2.</p>'
            + '<p><b>2. Find the column.</b> Same idea for the category: <code>VLOOKUP(C5,[master_produk.xlsx]Produk!$A$5:$D$16,3,FALSE)</code>, then <code>MATCH(that category,[tarif_kirim.xlsx]Tarif!$B$4:$E$4,0)</code>. KSJ-ARE is Coffee: column 2.</p>'
            + '<p><b>3. INDEX.</b> <code>INDEX([tarif_kirim.xlsx]Tarif!$B$5:$E$11, row, column)</code> reads the rate where they cross. Row first, then column. Then times the cartons.</p>'
            + '<p>End every MATCH with <code>0</code> for an exact match. Without it, Excel assumes the list is sorted A to Z, and this rate card is not. VLOOKUP ends with <code>FALSE</code>, MATCH ends with <code>0</code>: both mean exact.</p>'
            + '<p>Keep the three rate card ranges lined up: the rates <code>$B$5:$E$11</code>, the regions on the same rows <code>$A$5:$A$11</code>, the categories in the same columns <code>$B$4:$E$4</code>.</p>',
          ex: [['=INDEX(' + TBL + ',', 'In the rates,'], ['MATCH("Sumatera",' + REG + ',0),', 'go down to the Sumatera row,'], ['MATCH("Soap",' + HDR + ',0))', 'then across to the Soap column.']],
          exFormula: '=INDEX(' + TBL + ',MATCH("Sumatera",' + REG + ',0),MATCH("Soap",' + HDR + ',0))', exLabel: 'rupiah per carton of soap to Sumatera',
          bridge: 'Like a depreciation rate table: find the asset group down the side, the method across the top, and read the rate where they meet.'
        },
        starter: '=D5*INDEX(', fnKey: 'INDEX(',
        chips: ['MATCH(', 'VLOOKUP(', TBL, REG, HDR, TGT, PRD, ',2', ',3', ',FALSE)', ',0)'],
        solution: SOL,
        alts: [
          '=INDEX(' + TBL + ',' + ROW + ',' + COL + ')*D5',
          '=D5*INDEX(' + TBL + ',XMATCH(' + AREA + ',' + REG + '),XMATCH(' + CAT + ',' + HDR + '))',
          '=LET(area,' + AREA + ',cat,' + CAT + ',D5*INDEX(' + TBL + ',MATCH(area,' + REG + ',0),MATCH(cat,' + HDR + ',0)))',
          '=D5*XLOOKUP(XLOOKUP(B5,' + range11(BK, 'Target', 'A') + ',' + range11(BK, 'Target', 'B') + '),' + REG + ',XLOOKUP(XLOOKUP(C5,' + range11(BK, 'Produk', 'A') + ',' + range11(BK, 'Produk', 'C') + '),' + HDR + ',' + TBL + '))'
        ],
        traps: [
          { code: 'approx', formula: '=D5*INDEX(' + TBL + ',MATCH(' + AREA + ',' + REG + '),MATCH(' + CAT + ',' + HDR + '))', title: 'Each MATCH needs its 0', plain: 'Without the 0, MATCH assumes the list is sorted A to Z. The rate card is not sorted, so the practice sheet finds nothing (#N/A), and desktop Excel can land on a wrong rate without any warning. End both MATCHes with ,0. Both 0s already there? Then check that the VLOOKUPs bring back the Area (column 2) and the category (column 3), and that every range keeps its $.', reply: 'Column E says #N/A 🤔 Does each MATCH end with <b>,0</b>? Without it, MATCH reads the rate card as if it were sorted A to Z, and it is not. In Excel it can even pick a wrong rate quietly.' },
          { code: 'swap', formula: '=D5*INDEX(' + TBL + ',' + COL + ',' + ROW + ')', title: 'Row first, then column', plain: 'INDEX takes the row number first and the column number second. The regions are the rows and the categories are the columns. Here the category went in as the row and the region as the column. A region past the 4th column gives #REF!, and the other lines read the wrong box. If the order is right, check that the categories are $B$4:$E$4, the same columns as the rates $B$5:$E$11.', reply: '#REF! on some lines 😅 The rate card has 7 regions down but only 4 categories across. INDEX wants the <b>row</b> first, the region, then the <b>column</b>, the category.' },
          { code: 'shift', formula: '=D5*INDEX(' + sheetRef(BK, SHEET11.Tarif.book, 'Tarif') + '$A$5:$E$11,' + ROW + ',' + COL + ')', title: 'The ranges do not line up with the rates', plain: 'INDEX counts rows and columns from the top-left box of its own range. The rates are $B$5:$E$11. Start INDEX at column A or at row 4 and every read moves one box, so some lines land on a region name or a heading (#VALUE!). Keep the regions on the same rows, $A$5:$A$11, and the categories in the same columns, $B$4:$E$4.', reply: '#VALUE! on some lines 🤔 INDEX is reading one box off. The rates start at <b>B5</b>, so the regions start on row 5 and the categories in column B.' },
          { code: 'rate', formula: '=INDEX(' + TBL + ',' + ROW + ',' + COL + ')', title: 'That is the rate for one carton', plain: 'INDEX gives the rate per carton. Multiply it by the cartons in D5.', reply: 'Rp {value} to ship 53 cartons? That is one carton 😄 Times the <b>cartons</b> in column D please.' }
        ],
        replies: {
          pass: ['Mantap! Every line has its shipping cost 🚚', 'Four files, one formula. Next time the shipper changes a rate, I fix one box and column E follows.'],
          frozen: 'The shipper raised the Sumatera and Sulawesi rates this morning and your column did not move 😅 It has to read the rate card.',
          typed: 'Typed costs will be wrong the next time the rate card changes 🙏'
        },
        note: { fn: 'Two-way lookup across workbooks', what: 'VLOOKUP finds the Area and the category, MATCH turns each into a position on the rate card, INDEX reads the rate where row and column cross.', shape: SOL },
        hints: [
          'Work in steps: the distributor\'s Area from target_2026.xlsx, the SKU\'s category from master_produk.xlsx, then the rate where they cross in tarif_kirim.xlsx, times the cartons.',
          'The row: MATCH the Area, VLOOKUP(B5,' + TGT + ',2,FALSE), in ' + REG + '. The column: MATCH the category, VLOOKUP(C5,' + PRD + ',3,FALSE), in ' + HDR + '. Both MATCHes end with 0, and INDEX takes the row first.',
          'It looks like ' + SOL + '.'
        ]
      });
    })(),
    real({
      id: '11.9', from: 'lestari', due: '16:15', title: 'Late fees with a floor and a cap', main: 'Pinjaman', open: ['Aturan'], ansCol: 'H', ansLabel: 'Late fee (Rp)',
      /* Overnight the bank also raises the fee per day, so a typed 0.1 is caught as frozen like the other two rules. */
      editOpen: (name, c, v) => { if (name === 'Aturan' && v === 'alt') c.B6 = 0.15; },
      intro: [
        'Sore {name}, Bu Lestari again.',
        'Before I sign more orders, I want to see what each distributor owes Bank Kenari in late fees.',
        'The loan book <b>buku_pinjaman.xlsx</b> is open. The bank\'s late-fee rules are on its second sheet, <b>Aturan</b>.',
        '@file',
        'Late fee = outstanding × the fee per day late (%) ÷ 100 × the days late <b>beyond the grace days</b>.',
        'It is <b>never below 0</b>, so a loan inside the grace days pays nothing. And it is <b>never above the fee cap</b>.',
        'Column H please. Take the grace days, the fee % and the cap from <b>Aturan</b>, because the bank changes them.'
      ],
      ask: 'Outstanding × fee % ÷ 100 × (days late − grace days), at least 0 and at most the cap from Aturan, every row, in column H.',
      lesson: {
        idea: 'MAX sets a floor and MIN sets a ceiling. Together they keep a number between the two.',
        teach: '<p><code>MAX(0, fee)</code> takes the bigger of 0 and the fee, so it never drops below 0. <code>MIN(cap, fee)</code> takes the smaller, so it never goes above the cap. Put one inside the other: <code>MIN(cap, MAX(0, fee))</code>.</p><p>Build it from the inside out: the fee first, then <code>MAX(0, ...)</code> around it, then <code>MIN(cap, ...)</code> around that.</p><p>The rules sit on the <b>Aturan</b> sheet of this same workbook. When you click a cell on another sheet, Excel writes <code>Aturan!B5</code> with no $, unlike a cell in another workbook. Press <b>F4</b> (Cmd+T on a Mac) to pin it: <code>Aturan!$B$5</code>.</p>',
        ex: [['=MAX(0,', 'The bigger of 0'], ['E7-Aturan!$B$5)', 'and the days late minus the grace days.']],
        exFormula: '=MAX(0,E7-Aturan!$B$5)', exLabel: 'days to charge for F103',
        bridge: 'Like a penalty clause in a contract: a grace period first, then a daily rate, up to a maximum penalty.'
      },
      starter: '=MIN(', fnKey: 'MIN(', chips: ['Aturan!$B$7', ',MAX(0,', 'D5', '*Aturan!$B$6/100', '*(E5-Aturan!$B$5)', '))'],
      solution: '=MIN(Aturan!$B$7,MAX(0,D5*Aturan!$B$6/100*(E5-Aturan!$B$5)))',
      alts: [
        '=MAX(0,MIN(Aturan!$B$7,D5*Aturan!$B$6/100*(E5-Aturan!$B$5)))',
        '=MIN(Aturan!$B$7,D5*Aturan!$B$6/100*MAX(0,E5-Aturan!$B$5))',
        '=IF(E5<=Aturan!$B$5,0,MIN(Aturan!$B$7,D5*Aturan!$B$6/100*(E5-Aturan!$B$5)))',
        '=MIN(Aturan!$B$7,MAX(0,D5*Aturan!$B$6%*(E5-Aturan!$B$5)))',
        '=MEDIAN(0,D5*Aturan!$B$6/100*(E5-Aturan!$B$5),Aturan!$B$7)',
        '=LET(fee,D5*Aturan!$B$6/100*(E5-Aturan!$B$5),MIN(Aturan!$B$7,MAX(0,fee)))'
      ],
      traps: [
        { code: 'nomax', formula: '=MIN(Aturan!$B$7,D5*Aturan!$B$6/100*(E5-Aturan!$B$5))', title: 'On-time loans get a negative fee', plain: 'A loan with fewer days late than the grace days gives a minus fee. Wrap the fee in MAX(0, ...) so it never goes below 0.', reply: 'F101, CV Nuri Makmur, shows <b>minus</b> Rp 231,000. They paid on time, and the bank does not pay them for it 😅 Never below 0 please: <b>MAX(0, ...)</b>.' },
        { code: 'nocap', formula: '=MAX(0,D5*Aturan!$B$6/100*(E5-Aturan!$B$5))', title: 'Some fees went over the cap', plain: 'Very late loans pass the cap in Aturan!$B$7. Take the smaller of the cap and the fee with MIN.', reply: 'F109, CV Tekukur Rejeki, at Rp 19,107,000? The bank stops at the cap, <b>Rp 5,000,000</b>. Take the smaller one with <b>MIN</b>.' },
        { code: 'nograce', formula: '=MIN(Aturan!$B$7,MAX(0,D5*Aturan!$B$6/100*E5))', title: 'The grace days were charged too', plain: 'Only the days beyond the grace days count. Subtract Aturan!$B$5 from the days late in E5.', reply: 'F106, UD Seriti Jaya, is 15 days late. The first 7 are free, so only 8 days count. Please subtract the <b>grace days</b>.' },
        { code: 'nodollar', formula: '=MIN(Aturan!B7,MAX(0,D5*Aturan!B6/100*(E5-Aturan!B5)))', title: 'The rule cells slid down', plain: 'Copied down, Aturan!B5 became Aturan!B6, then B7, and below the rules the cells are empty. Pin them with $: Aturan!$B$5, Aturan!$B$6, Aturan!$B$7.', reply: 'Every fee is 0? CV Kepodang Niaga is 95 days late 🤔 When you copy down, the Aturan cells move too. Press <b>F4</b> to add the <b>$</b>.' }
      ],
      frozenTry: '=MIN(Aturan!$B$7,MAX(0,D5*0.1/100*(E5-Aturan!$B$5)))',
      replies: {
        pass: ['Terima kasih. Five distributors already hit the full Rp 5,000,000 cap.', 'Kepodang, Rangkong, Tekukur, Bangau and Kedasi. I will call them before I approve new orders 🙏'],
        frozen: 'Bank Kenari changed all three rules this morning: 5 grace days, 0.15% a day and a cap of Rp 4,000,000. Your fees still follow the old ones. Please read them from the Aturan sheet, not typed numbers.',
        typed: 'Typed fees will not follow the bank\'s rules 🙏'
      },
      note: { fn: 'MIN and MAX together', what: 'MAX(0, x) is a floor and MIN(cap, x) a ceiling. Rules on another sheet are pinned with $.', shape: '=MIN(Aturan!$B$7,MAX(0,D5*Aturan!$B$6/100*(E5-Aturan!$B$5)))' },
      hints: [
        'The fee is D5 × the fee % ÷ 100 × the days beyond the grace days. Then keep it between 0 and the cap.',
        'MAX(0, fee) stops it below 0, and MIN(Aturan!$B$7, ...) stops it above the cap. Pin the Aturan cells with $.',
        'It looks like =MIN(Aturan!$B$7,MAX(0,D5*Aturan!$B$6/100*(E5-Aturan!$B$5))).'
      ]
    }),
    (() => {
      const FB = 'ringkasan_q4.xlsx';
      const T = table11(FB, 'Pinjaman', 'B', 'G');
      const V = (n) => 'VLOOKUP(A5,' + T + ',' + n + ',FALSE)';
      const X = (c) => 'XLOOKUP(A5,' + range11(FB, 'Pinjaman', 'B') + ',' + range11(FB, 'Pinjaman', c) + ')';
      const M = (c) => 'INDEX(' + range11(FB, 'Pinjaman', c) + ',MATCH(A5,' + range11(FB, 'Pinjaman', 'B') + ',0))';
      const SOLUTION = '=IF(AND(' + V(4) + '>30,' + V(3) + '>B5/2),"WATCH","OK")';
      /* Collection keeps its loan book sorted by days late, worst first, so its rows do not line up with the summary
         and a row-by-row link ([buku_pinjaman.xlsx]Pinjaman!E5) gives wrong statuses. Only a lookup by name works. */
      const P = SHEET11.Pinjaman;
      const worstFirst = (name, c) => {
        if (name !== 'Pinjaman') return;
        const rows = [];
        for (let r = P.first; r <= P.last; r++) rows.push(P.cols.map((k) => c[k + r]));
        rows.sort((a, b) => b[4] - a[4] || (a[0] < b[0] ? -1 : 1));
        rows.forEach((row, i) => P.cols.forEach((k, j) => { c[k + (P.first + i)] = row[j]; }));
      };
      return Object.assign(real({
        id: '11.10', from: 'citra', due: '16:30', title: 'The watch list, across two real workbooks', main: 'Ringkasan', open: ['Pinjaman'], ansLabel: 'Status', editOpen: worstFirst,
        intro: [
          'Sore {name}. Last one of the year, and it goes to the credit committee in <b>real Excel</b> 💻',
          'Our watch list for 2027 needs two files: Rasa Kenari\'s Q4 summary, <b>ringkasan_q4.xlsx</b>, and our loan book, <b>buku_pinjaman.xlsx</b>.',
          '@file',
          'The rule: a distributor is <b>WATCH</b> when it is <b>more than 30 days late</b> AND its outstanding is <b>more than half</b> of its Q4 net sales (column B of the summary). Otherwise <b>OK</b>.',
          'Days late and outstanding are in the loan book. We sort it by days late, worst first, so match each distributor by name in its column B.',
          'Download both files and open both in Excel.',
          'In ringkasan_q4.xlsx, type the formula in <b>C5</b>. For the lookups, click into the loan book so Excel writes the link itself.',
          'Fill it down to C18, with WATCH and OK in capitals. Save ringkasan_q4.xlsx and upload it here.'
        ],
        ask: 'In Excel: WATCH if more than 30 days late AND outstanding is more than half of Q4 net sales, else OK, in C5:C18 of ringkasan_q4.xlsx, linked to buku_pinjaman.xlsx. Then upload.',
        lesson: {
          idea: 'A real link between two workbooks: one file reads the other, and the link lives in the formula.',
          teach: '<p>Open both files. In C5 of ringkasan_q4.xlsx, type <code>=IF(AND(VLOOKUP(A5,</code>, then switch to the buku_pinjaman.xlsx window and drag across B5:G18. Excel writes the link for you, like <code>' + T + '</code>, already pinned with $. Counting from column B: outstanding is column 3, days late is column 4. Keep typing <code>,4,FALSE)&gt;30,</code>, add the second lookup the same way, and Enter takes you back to ringkasan_q4.xlsx.</p>' +
            '<p>Why a lookup, and not a link to the same row? The loan book is in another order. Row 5 there is a different distributor from row 5 here.</p>' +
            '<p>Then the rule: <code>IF(AND(days late&gt;30, outstanding&gt;B5/2), "WATCH", "OK")</code>. Both tests must be true. "More than" is <code>&gt;</code>, so exactly half is not a WATCH. Type <code>B5</code> yourself: a click while the loan book is in front would point at the loan book\'s B5.</p>' +
            '<p>When you close the loan book, Excel shows its full folder path inside the formula. That is normal. VLOOKUP keeps working with the file closed, and Excel asks to update the link the next time you open ringkasan_q4.xlsx. SUMIF and COUNTIF are different: they need the other file open, or they show #VALUE!.</p>',
          bridge: 'Like cross-referencing to another working paper file: the link shows where each number came from, and it follows when that file changes.'
        },
        starter: '', fnKey: '', chips: [],
        solution: SOLUTION,
        alts: [
          '=IF(AND(' + X('E') + '>30,' + X('D') + '>B5/2),"WATCH","OK")',
          '=IF(AND(' + M('E') + '>30,' + M('D') + '>B5/2),"WATCH","OK")',
          '=LET(late,' + V(4) + ',owed,' + V(3) + ',IF(AND(late>30,owed>B5/2),"WATCH","OK"))',
          '=IF(OR(' + V(4) + '<=30,' + V(3) + '<=B5/2),"OK","WATCH")',
          '=IF(AND(' + V(4) + '>30,' + V(3) + '/B5>0.5),"WATCH","OK")'
        ],
        exactCase: true,
        rowReply: 'Row {row} says {got}, but {key} should be {want}. Find {key} by name in the loan book and check its days late and outstanding.', keyCol: 'A',
        replies: {
          pass: ['Received. Six on WATCH, eight OK. The committee has it 🙏', 'CV Kasuari Sentral is only 35 days late, but owes more than half its Q4 sales. CV Kutilang Abadi is 45 days late, but owes about a third, so it stays OK. Both tests matter.'],
          wrong: '',
          typed: 'The statuses are typed in, not formulas 😅 When the bank updates the loan book, they will be wrong. One formula in C5, filled down.'
        },
        note: { fn: 'Linked workbooks in real Excel', what: 'Both files open, a lookup that clicks into the other workbook, IF(AND(...)) on top. Save the file that holds the formula.', shape: SOLUTION },
        hints: [
          'Two facts from the loan book for the distributor in A5: days late and outstanding. Then compare outstanding with half of B5.',
          'Look up A5 in the loan book\'s B5:G18 by clicking into it: outstanding is column 3, days late column 4. Put both tests inside AND, inside IF.',
          'It looks like =IF(AND(' + V(4) + '>30,' + V(3) + '>B5/2),"WATCH","OK").'
        ]
      }), { kind: 'excel' });
    })()
  );
  /* CH11-REQUESTS */

  /* Ranks along the way. Chapter 1 ends with a promotion to Analyst. */
  const RANKS = [
    { at: 0, title: 'Junior Analyst' },
    { at: 8, title: 'Analyst' },
    { at: 30, title: 'Senior Analyst' },
    { at: 70, title: 'Lead Analyst' }
  ];

  const CHAPTERS = [{ n: 1, title: 'Your first week', ids: LIST.filter((m) => !m.chapter).map((m) => m.id) }, { n: 2, title: 'Bank Kenari', ids: LIST.filter((m) => m.chapter === 2).map((m) => m.id) }, { n: 3, title: 'Lookups', ids: LIST.filter((m) => m.chapter === 3).map((m) => m.id) }, { n: 4, title: 'Text and dates', ids: LIST.filter((m) => m.chapter === 4).map((m) => m.id) }, { n: 5, title: 'Dynamic arrays', ids: LIST.filter((m) => m.chapter === 5).map((m) => m.id) }, { n: 6, title: 'Reporting', ids: LIST.filter((m) => m.chapter === 6).map((m) => m.id) }, { n: 7, title: 'Power Query', ids: LIST.filter((m) => m.chapter === 7).map((m) => m.id) }, { n: 8, title: 'What-if and loans', ids: LIST.filter((m) => m.chapter === 8).map((m) => m.id) }, { n: 9, title: 'Macros', ids: LIST.filter((m) => m.chapter === 9).map((m) => m.id) }, { n: 10, title: 'The early-warning report', ids: LIST.filter((m) => m.chapter === 10).map((m) => m.id) }, { n: 11, title: 'Real cases', ids: LIST.filter((m) => m.chapter === 11).map((m) => m.id) }];
  const DATA_TOTALS = Object.fromEntries([...new Set(DATA.map((d) => d.Distributor))].map((k) => [k, DATA.filter((d) => d.Distributor === k).reduce((a, d) => a + d.Net, 0)]));
  return { DATA_TOTALS, LIST, RANKS, CHAPTERS, target, other, ALT_D5, FIRST, LAST, rows, FAC, B_FIRST, B_LAST, SHEET11, real, range11, table11, sheetRef };
});
