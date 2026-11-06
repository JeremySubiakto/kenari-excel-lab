/*
 * The formula engine: Formualizer 0.13.0 (MIT or Apache-2.0), shipped in assets/engine/ (or a pinned CDN URL).
 * Its npm package is built for bundlers, so this wires up the WebAssembly by hand.
 * docs/ENGINE.md says why this engine was chosen and what it gets wrong.
 */
const VERSION = '0.13.0';
const CDN = `https://cdn.jsdelivr.net/npm/formualizer@${VERSION}/pkg/`;
/* A page can ship the two engine files itself (a private artifact cannot fetch the .wasm from a CDN).
   It sets window.KENARI_ENGINE_BASE to the folder that holds them. */
const BASE = (typeof window !== 'undefined' && window.KENARI_ENGINE_BASE) || CDN;

let loading = null;

async function load() {
  if (loading) return loading;
  loading = (async () => {
    const started = performance.now();
    const bg = await import(BASE + 'formualizer_wasm_bg.js');
    const imports = { './formualizer_wasm_bg.js': bg };
    const resp = fetch(BASE + 'formualizer_wasm_bg.wasm');
    let instance;
    try {
      ({ instance } = await WebAssembly.instantiateStreaming(resp, imports));
    } catch (e) {
      /* Some proxies send the wrong content type, which streaming refuses. */
      const bytes = await (await fetch(BASE + 'formualizer_wasm_bg.wasm')).arrayBuffer();
      ({ instance } = await WebAssembly.instantiate(bytes, imports));
    }
    bg.__wbg_set_wasm(instance.exports);
    if (instance.exports.__wbindgen_start) instance.exports.__wbindgen_start();
    return { Workbook: bg.Workbook, loadMs: Math.round(performance.now() - started) };
  })();
  return loading;
}

const EPOCH = Date.UTC(1899, 11, 30);

/** Engine values to plain JS: numbers, text, booleans, '' for empty, and error strings like '#N/A'. */
function plain(v) {
  if (v === undefined || v === null) return '';
  if (v instanceof Date) return Math.round((v.getTime() - EPOCH) / 864e5 * 1e6) / 1e6;
  if (typeof v === 'object') {
    if (v.type === 'Error' || v.kind === 'Error' || v.error) return String(v.error || v.message || v.value || v.code || '#ERROR');
    if (v.value !== undefined) return plain(v.value);
    return JSON.stringify(v);
  }
  return v;
}

const parseRef = (a1) => {
  const m = /^\$?([A-Z]+)\$?(\d+)$/.exec(a1);
  return { col: m[1].split('').reduce((a, c) => a * 26 + c.charCodeAt(0) - 64, 0), row: +m[2] };
};

/*
 * More than one sheet, and more than one workbook. A request's cells can carry extra sheets under SHEETS:
 * [{ name, cells }]. A sheet of the same workbook is named as in Excel ("Targets"); a sheet of another open
 * workbook is named "file.xlsx|Sheet", because the engine does not read Excel's external form
 * [file.xlsx]Sheet!A1. toEngine() rewrites that form before a formula reaches the engine, so the learner
 * types and sees exactly what Excel shows. MAIN names the request's own sheet (default Sheet1).
 * Symbol keys are skipped by Object.entries and copied by Object.assign, so plain cell code is unaffected.
 */
const SHEETS = Symbol.for('kenari.sheets');
const MAIN = Symbol.for('kenari.main');
/** MATCH(x, range, FALSE): Excel reads FALSE as 0, an exact match; the engine reads it as approximate. */
function exactMatch(f) {
  let out = '', i = 0;
  while (i < f.length) {
    const at = f.slice(i).search(/(^|[^A-Za-z.])MATCH\(/i);
    if (at < 0) { out += f.slice(i); break; }
    const open = i + at + f.slice(i + at).indexOf('(') + 1;
    out += f.slice(i, open);
    let depth = 1, j = open, inQ = false, comma = [];
    for (; j < f.length && depth; j++) {
      const c = f[j];
      if (c === '"') inQ = !inQ;
      else if (!inQ && c === '(') depth++;
      else if (!inQ && c === ')') depth--;
      else if (!inQ && c === ',' && depth === 1) comma.push(j);
    }
    const close = j - 1;
    let inner = f.slice(open, close);
    if (comma.length === 2 && /^\s*FALSE\s*$/i.test(f.slice(comma[1] + 1, close))) inner = f.slice(open, comma[1] + 1) + '0';
    out += exactMatch(inner) + ')';
    i = close + 1;
  }
  return out;
}
const toEngine = (f) => exactMatch(String(f).replace(/'?\[([^\]]+)\]([^'!]+)'?!/g, (m, book, sheet) => "'" + book + '|' + sheet + "'!"));

/** A workbook with the request's sheet, plus any extra sheets. `cells` is {A1: value or '=formula'}. */
class Book {
  constructor(api, cells) {
    this.wb = new api.Workbook();
    this.main = cells[MAIN] || 'Sheet1';
    this.wb.addSheet(this.main);
    for (const sh of cells[SHEETS] || []) {
      this.wb.addSheet(sh.name);
      for (const [ref, v] of Object.entries(sh.cells)) this.set(ref, v, sh.name);
    }
    for (const [ref, v] of Object.entries(cells)) this.set(ref, v);
  }
  set(ref, v, sheet) {
    const { row, col } = parseRef(ref);
    if (typeof v === 'string' && v[0] === '=') this.wb.setFormula(sheet || this.main, row, col, toEngine(v));
    else this.wb.setValue(sheet || this.main, row, col, v);
  }
  calc() { this.wb.evaluateAll(); }
  get(ref) { const { row, col } = parseRef(ref); return plain(this.wb.sheet(this.main).getValue(row, col)); }
  close() { try { this.wb.free(); } catch (e) { /* already gone */ } }
}

/**
 * Put one formula on a fresh copy of the sheet and read it back.
 * A formula the engine refuses to read comes back as '#PARSE'.
 */
function evalOn(api, cells, at, formula) {
  const b = new Book(api, cells);
  try {
    b.set(at, formula);
    b.calc();
    return b.get(at);
  } catch (e) {
    return '#PARSE';
  } finally {
    b.close();
  }
}

/**
 * Fill a formula down a column, as Excel's fill handle does, and read every cell back.
 * shift(formula, dr) moves the relative rows. Returns an array of values, or ['#PARSE'] if unreadable.
 */
function evalFill(api, cells, col, first, last, formula, shift) {
  const b = new Book(api, cells);
  try {
    for (let r = first; r <= last; r++) b.set(col + r, shift(formula, r - first));
    b.calc();
    const out = [];
    for (let r = first; r <= last; r++) out.push(b.get(col + r));
    return out;
  } catch (e) {
    return ['#PARSE'];
  } finally {
    b.close();
  }
}

/**
 * Put one formula at `at` and read the block it spills into: `rows` by `cols` cells, row by row.
 * Cells the result does not reach come back as ''.
 */
function evalSpill(api, cells, at, formula, cols, rows) {
  const b = new Book(api, cells);
  const { row, col } = parseRef(at);
  try {
    b.set(at, formula);
    b.calc();
    const out = [];
    for (let r = 0; r < rows; r++) for (let c = 0; c < cols; c++) out.push(plain(b.wb.sheet(b.main).getValue(row + r, col + c)));
    return out;
  } catch (e) {
    return ['#PARSE'];
  } finally {
    b.close();
  }
}

export { load, Book, evalOn, evalFill, evalSpill, plain, toEngine, SHEETS, MAIN, VERSION };
