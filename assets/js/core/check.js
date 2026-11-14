/*
 * Answer checking and plain-English errors for formula missions.
 * Pure functions with no DOM and no engine, so the app and the tests share them.
 * Classic script in the browser (window.KenariCheck), CommonJS in Node.
 *
 * The idea: a formula passes when it gives the right number now AND keeps giving the
 * right number when the sheet changes. A typed-in number, or a sum of hand-picked cells,
 * can be right once. It cannot survive the second sheet.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.KenariCheck = factory();
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';

  const KNOWN = ['ABS', 'ACCRINT', 'ADDRESS', 'AGGREGATE', 'AND', 'ARRAYTOTEXT', 'AVERAGE', 'AVERAGEA', 'AVERAGEIF', 'AVERAGEIFS', 'BYCOL', 'BYROW', 'CEILING', 'CEILING.MATH', 'CHAR', 'CHOOSE', 'CHOOSECOLS', 'CHOOSEROWS', 'CLEAN', 'CODE', 'COLUMN', 'COLUMNS', 'CONCAT', 'CONCATENATE', 'CORREL', 'COUNT', 'COUNTA', 'COUNTBLANK', 'COUNTIF', 'COUNTIFS', 'COUPDAYS', 'CUMIPMT', 'CUMPRINC', 'DATE', 'DATEDIF', 'DATEVALUE', 'DAY', 'DAYS', 'DAYS360', 'DB', 'DDB', 'DOLLAR', 'DROP', 'EDATE', 'EFFECT', 'EOMONTH', 'EVEN', 'EXACT', 'EXP', 'EXPAND', 'FALSE', 'FILTER', 'FIND', 'FIXED', 'FLOOR', 'FLOOR.MATH', 'FORECAST', 'FORECAST.LINEAR', 'FREQUENCY', 'FV', 'FVSCHEDULE', 'GEOMEAN', 'HLOOKUP', 'HOUR', 'HSTACK', 'HYPERLINK', 'IF', 'IFERROR', 'IFNA', 'IFS', 'INDEX', 'INDIRECT', 'INT', 'INTERCEPT', 'IPMT', 'IRR', 'ISBLANK', 'ISERR', 'ISERROR', 'ISEVEN', 'ISFORMULA', 'ISLOGICAL', 'ISNA', 'ISNONTEXT', 'ISNUMBER', 'ISODD', 'ISOWEEKNUM', 'ISPMT', 'ISTEXT', 'LAMBDA', 'LARGE', 'LEFT', 'LEN', 'LET', 'LINEST', 'LN', 'LOG', 'LOG10', 'LOOKUP', 'LOWER', 'MAKEARRAY', 'MAP', 'MATCH', 'MAX', 'MAXA', 'MAXIFS', 'MEDIAN', 'MID', 'MIN', 'MINA', 'MINIFS', 'MINUTE', 'MIRR', 'MOD', 'MODE', 'MONTH', 'MROUND', 'N', 'NA', 'NETWORKDAYS', 'NETWORKDAYS.INTL', 'NOMINAL', 'NOT', 'NOW', 'NPER', 'NPV', 'NUMBERVALUE', 'ODD', 'OFFSET', 'OR', 'PERCENTILE', 'PERCENTRANK', 'PI', 'PMT', 'POWER', 'PPMT', 'PRODUCT', 'PROPER', 'PV', 'QUARTILE', 'QUOTIENT', 'RAND', 'RANDBETWEEN', 'RANK', 'RANK.AVG', 'RANK.EQ', 'RATE', 'REDUCE', 'REPLACE', 'REPT', 'RIGHT', 'ROUND', 'ROUNDDOWN', 'ROUNDUP', 'ROW', 'ROWS', 'RRI', 'SCAN', 'SEARCH', 'SECOND', 'SEQUENCE', 'SIGN', 'SLN', 'SLOPE', 'SMALL', 'SORT', 'SORTBY', 'SQRT', 'STDEV', 'STDEV.P', 'STDEV.S', 'SUBSTITUTE', 'SUBTOTAL', 'SUM', 'SUMIF', 'SUMIFS', 'SUMPRODUCT', 'SUMSQ', 'SWITCH', 'SYD', 'T', 'TAKE', 'TEXT', 'TEXTAFTER', 'TEXTBEFORE', 'TEXTJOIN', 'TEXTSPLIT', 'TIME', 'TIMEVALUE', 'TOCOL', 'TODAY', 'TOROW', 'TRANSPOSE', 'TREND', 'TRIM', 'TRUE', 'TRUNC', 'TYPE', 'UNIQUE', 'UPPER', 'VALUE', 'VAR', 'VAR.P', 'VAR.S', 'VLOOKUP', 'VSTACK', 'WEEKDAY', 'WEEKNUM', 'WORKDAY', 'WORKDAY.INTL', 'WRAPCOLS', 'WRAPROWS', 'XIRR', 'XLOOKUP', 'XMATCH', 'XNPV', 'XOR', 'YEAR', 'YEARFRAC'];

  /** Walk a formula, calling fn for each character outside double-quoted text. */
  function outsideStrings(f, fn) {
    let inStr = false, out = '';
    for (let i = 0; i < f.length; i++) {
      const c = f[i];
      if (c === '"') { inStr = !inStr; out += c; continue; }
      out += inStr ? c : fn(c);
    }
    return out;
  }

  /**
   * What the learner typed, ready for the engine. Excel on an Indonesian Windows or Mac setup
   * uses ; between arguments, so accept it here and say so in the lesson.
   */
  function normalize(text) {
    const t = String(text == null ? '' : text).trim();
    return outsideStrings(t, (c) => (c === ';' ? ',' : c));
  }

  function bracketGap(f) {
    let n = 0;
    outsideStrings(f, (c) => { if (c === '(') n++; else if (c === ')') n--; return c; });
    return n;
  }

  function editDistance(a, b) {
    const d = Array.from({ length: a.length + 1 }, (_, i) => [i]);
    for (let j = 1; j <= b.length; j++) d[0][j] = j;
    for (let i = 1; i <= a.length; i++) for (let j = 1; j <= b.length; j++) {
      d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    }
    return d[a.length][b.length];
  }

  function functionNames(f) {
    const plain = f.replace(/"[^"]*"/g, '""');
    const out = [];
    const re = /([A-Za-z][A-Za-z0-9.]*)\s*\(/g;
    let m;
    while ((m = re.exec(plain))) out.push(m[1]);
    return out;
  }

  /** The function name that looks like a typo, with the name it was probably meant to be. */
  function misspelled(f) {
    for (const name of functionNames(f)) {
      const up = name.toUpperCase();
      if (KNOWN.includes(up)) continue;
      /* Ties go to the name of the same length: SUMIFZ means SUMIFS, not SUMIF. */
      let best = null, bestD = 3;
      for (const k of KNOWN) { const d = editDistance(up, k) + (k.length === up.length ? 0 : 0.5); if (d < bestD) { bestD = d; best = k; } }
      return { name, suggest: best };
    }
    return null;
  }

  const ERRORS = {
    '#NAME?': { title: 'Excel does not know that word', plain: 'One word in your formula is not a function or a cell Excel can find. Check the spelling, and put text inside double quotes.' },
    '#VALUE!': { title: 'The ranges do not line up', plain: 'In SUMIFS, COUNTIFS or SUMPRODUCT every range must have the same number of rows. Check that each one starts and ends on the same rows.' },
    '#REF!': { title: 'A cell in this formula does not exist', plain: 'The formula points at a cell that is not there.' },
    '#DIV/0!': { title: 'This divides by zero', plain: 'The number under the division is 0 or empty.' },
    '#N/A': { title: 'Nothing matched', plain: 'A lookup searched for a value and did not find it.' },
    '#NUM!': { title: 'That number cannot be worked out', plain: 'A function got a number outside what it can handle.' },
    '#SPILL!': { title: 'The answer has no room to spread', plain: 'This formula returns several cells, and something is sitting in the way. Clear the cells below or beside it.' },
    '#CALC!': { title: 'Nothing came back', plain: 'The formula ran and found no rows to return.' },
    '#N/IMPL!': { title: 'The practice sheet cannot run this function yet', plain: 'Your formula may be fine in real Excel. The practice sheet does not have this one.' },
    '#CIRC!': { title: 'This formula includes its own cell', plain: 'One of your ranges covers the cell the formula sits in, so the answer would depend on itself. Whole columns like B:B often do this. Use a range that stops above or below it. Real Excel shows a circular reference warning here.' },
    '#PARSE': { title: 'Excel cannot read this formula yet', plain: 'Something in it is not finished. Look for a stray comma, an empty argument or a missing range.' }
  };

  const isError = (v) => typeof v === 'string' && /^#[A-Z/0-9!?]+/.test(v);

  /** Turn a broken formula, or an error value, into {title, plain, fix?}. */
  function explain(formula, value) {
    const gap = bracketGap(formula);
    if (gap > 0) return { title: gap === 1 ? 'One bracket is still open' : gap + ' brackets are still open', plain: 'Every ( needs a ) to close it. Add the closing bracket at the end.', fix: { label: 'Add ' + ')'.repeat(gap), formula: formula + ')'.repeat(gap) } };
    if (gap < 0) return { title: 'One bracket too many', plain: 'There is a ) with no ( to match it.' };
    const bad = misspelled(formula);
    if (bad && bad.suggest) {
      const re = new RegExp('\\b' + bad.name.replace(/[.]/g, '\\.') + '(?=\\s*\\()');
      return { title: 'Right idea, different spelling', plain: 'Excel does not have a function called ' + bad.name + '. It looks like you meant ' + bad.suggest + '.', fix: { label: 'Change ' + bad.name + ' to ' + bad.suggest, formula: formula.replace(re, bad.suggest) } };
    }
    /* #VALUE! has two common causes: ranges of different sizes in SUMIFS-type functions, or text where a number is needed. */
    if (value === '#VALUE!' && !/\b(SUMIFS|COUNTIFS|AVERAGEIFS|MAXIFS|MINIFS|SUMPRODUCT)\s*\(/i.test(formula)) return { title: 'Something here is not a number', plain: 'Part of this formula got text, or a whole range, where it needed one number. Check the cells you pointed at: a header row, a word, or a range like E5:E21 inside a calculation that needs one cell.' };
    return ERRORS[value] || { title: 'Excel returned ' + value, plain: 'This formula gave an error value.' };
  }

  /** Numbers match to `precision` decimals, text ignores capitals, the way Excel's = does. */
  function same(a, b, precision, exactCase) {
    const p = precision == null ? 2 : precision;
    if (typeof a === 'number' && typeof b === 'number') return Math.abs(a - b) < Math.pow(10, -p) / 2;
    if (typeof a === 'string' && typeof b === 'string') return exactCase ? a === b : a.toLowerCase() === b.toLowerCase();
    return a === b;
  }

  /**
   * Judge one formula. `run(formula, variant)` evaluates on a fresh sheet and returns the value.
   * Variants come from mission.variants, plus 'base'.
   * Returns {state:'pass'|'wrong'|'error'|'empty', title, plain, ...}.
   */
  function judge(mission, typed, run) {
    const f = normalize(typed);
    if (!f || f === '=') return { state: 'empty', title: 'The formula bar is empty', plain: 'Start with = and build the formula.' };
    if (f[0] !== '=') {
      const num = Number(f.replace(/[.,\s]/g, ''));
      return { state: 'wrong', code: 'typed', title: mission.msgs.typed.title, plain: mission.msgs.typed.plain, typed: Number.isNaN(num) ? f : num };
    }
    const gap = bracketGap(f);
    if (gap !== 0) return Object.assign({ state: 'error' }, explain(f));
    const mine = run(f, 'base');
    if (isError(mine)) {
      /* A known mistake can end in an error too, such as a missing $ that slides onto a header. Name the mistake, not the error. */
      for (const t of mission.traps || []) {
        if (run(t.formula, 'base') === mine) return { state: 'wrong', code: t.code, title: t.title, plain: t.plain, value: mine };
      }
      return Object.assign({ state: 'error', value: mine }, explain(f, mine));
    }
    const target = run(mission.solution, 'base');
    if (!same(mine, target, mission.precision, mission.exactCase)) {
      for (const t of mission.traps || []) {
        if (same(mine, run(t.formula, 'base'), mission.precision, mission.exactCase)) return { state: 'wrong', code: t.code, title: t.title, plain: t.plain, value: mine, target };
      }
      return { state: 'wrong', code: 'value', title: mission.msgs.value.title, plain: mission.msgs.value.plain, value: mine, target };
    }
    for (const v of mission.variants || []) {
      const a = run(f, v.id), b = run(mission.solution, v.id);
      if (!same(a, b, mission.precision, mission.exactCase)) return { state: 'wrong', code: 'frozen', title: mission.msgs.frozen.title, plain: mission.msgs.frozen.plain, value: mine, target };
    }
    for (const r of mission.requires || []) {
      if (!r.re.test(f)) return { state: 'wrong', code: 'technique', title: r.title, plain: r.plain, value: mine, target };
    }
    return { state: 'pass', value: mine, target, formula: f };
  }

  /**
   * The formula Excel's fill handle would put dr rows lower: relative row numbers move, $-pinned ones stay.
   * Text in double quotes and function names such as LOG10( are left alone.
   */
  function shiftRows(f, dr) {
    if (!dr) return f;
    return String(f).split(/("[^"]*"?)/).map((part, i) => (i % 2 ? part : part.replace(/(^|[^A-Za-z0-9_$.])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d(A-Za-z_])/g,
      (m, pre, d1, col, d2, row) => pre + d1 + col + d2 + (d2 ? row : String(+row + dr))))).join('');
  }

  /** Several values as one comparable key, so a whole filled column is judged like a single answer. */
  function columnKey(values, exactCase) {
    const err = values.find(isError);
    if (err) return err;
    return values.map((v) => (typeof v === 'number' ? (Math.round(v * 1e4) / 1e4).toString() : exactCase ? String(v) : String(v).toLowerCase())).join(' | ');
  }

  return { normalize, bracketGap, misspelled, explain, isError, same, judge, editDistance, shiftRows, columnKey };
});
