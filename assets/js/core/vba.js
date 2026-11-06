/*
 * A small VBA stepper for teaching. It runs a macro one line at a time on a fake sheet.
 * It covers the first macros a learner meets. It is NOT a full interpreter:
 *   Sub, Dim, assignment, For/Next, If/ElseIf/Else, Do While/Loop, MsgBox, comments,
 *   Cells(r, c).Value, Range("A1").Value, .Interior.Color, Cells(Rows.Count, c).End(xlUp).Row
 * Anything else stops with a plain message naming the line. There is no eval here.
 *
 * Classic script in the browser (window.KenariVBA), CommonJS in Node.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.KenariVBA = factory();
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';

  const CONSTANTS = { vbyellow: 'yellow', vbred: 'red', vbgreen: 'green', vbwhite: 'white', true: true, false: false, xlup: 'xlUp' };
  const colName = (n) => { let s = ''; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = (n - r - 1) / 26; } return s; };
  const colNum = (s) => s.toUpperCase().split('').reduce((a, c) => a * 26 + c.charCodeAt(0) - 64, 0);

  /* ---------- Expressions ---------- */
  function tokenize(src) {
    const out = [];
    let i = 0;
    while (i < src.length) {
      const c = src[i];
      if (/\s/.test(c)) { i++; continue; }
      if (c === '"') { let j = i + 1, s = ''; while (j < src.length) { if (src[j] === '"') { if (src[j + 1] === '"') { s += '"'; j += 2; continue; } break; } s += src[j++]; } out.push({ t: 'str', v: s }); i = j + 1; continue; }
      const num = /^\d+(\.\d+)?/.exec(src.slice(i));
      if (num) { out.push({ t: 'num', v: Number(num[0]) }); i += num[0].length; continue; }
      const id = /^[A-Za-z_][A-Za-z0-9_]*/.exec(src.slice(i));
      if (id) { out.push({ t: 'id', v: id[0] }); i += id[0].length; continue; }
      const two = src.slice(i, i + 2);
      if (two === '<>' || two === '<=' || two === '>=') { out.push({ t: 'op', v: two }); i += 2; continue; }
      if ('+-*/\\^&=<>(),.'.includes(c)) { out.push({ t: 'op', v: c }); i++; continue; }
      throw new Error('This demo cannot read the character ' + c);
    }
    return out;
  }

  /** Parses and evaluates at once, since the sheet can change between lines anyway. */
  function evaluate(src, st) {
    const toks = tokenize(src);
    let p = 0;
    const peek = () => toks[p];
    const isOp = (v) => peek() && peek().t === 'op' && peek().v === v;
    const isWord = (v) => peek() && peek().t === 'id' && peek().v.toLowerCase() === v;
    const expect = (v) => { if (!isOp(v)) throw new Error('Expected ' + v); p++; };

    function orE() { let a = andE(); while (isWord('or')) { p++; const b = andE(); a = a || b; } return a; }
    function andE() { let a = notE(); while (isWord('and')) { p++; const b = notE(); a = a && b; } return a; }
    function notE() { if (isWord('not')) { p++; return !notE(); } return cmp(); }
    function cmp() {
      let a = concat();
      while (peek() && peek().t === 'op' && ['=', '<>', '<', '>', '<=', '>='].includes(peek().v)) {
        const op = toks[p++].v, b = concat();
        const x = a === '' || a == null ? (typeof b === 'number' ? 0 : '') : a;
        const y = b === '' || b == null ? (typeof a === 'number' ? 0 : '') : b;
        a = op === '=' ? x === y : op === '<>' ? x !== y : op === '<' ? x < y : op === '>' ? x > y : op === '<=' ? x <= y : x >= y;
      }
      return a;
    }
    function concat() { let a = add(); while (isOp('&')) { p++; const b = add(); a = String(a == null ? '' : a) + String(b == null ? '' : b); } return a; }
    function add() { let a = mul(); while (isOp('+') || isOp('-')) { const op = toks[p++].v, b = mul(); a = op === '+' ? num(a) + num(b) : num(a) - num(b); } return a; }
    function mul() {
      let a = pow();
      for (;;) {
        if (isOp('*')) { p++; a = num(a) * num(pow()); }
        else if (isOp('/')) { p++; a = num(a) / num(pow()); }
        else if (isOp('\\')) { p++; a = Math.trunc(num(a) / num(pow())); }
        else if (isWord('mod')) { p++; a = num(a) % num(pow()); }
        else return a;
      }
    }
    function pow() { const a = unary(); if (isOp('^')) { p++; return Math.pow(num(a), num(pow())); } return a; }
    function unary() { if (isOp('-')) { p++; return -num(unary()); } return prim(); }
    function num(v) { if (v === '' || v == null) return 0; if (typeof v === 'boolean') return v ? -1 : 0; const n = Number(v); if (Number.isNaN(n)) throw new Error('"' + v + '" is text, so it cannot be used as a number here'); return n; }
    function args() { expect('('); const a = []; if (!isOp(')')) { a.push(orE()); while (isOp(',')) { p++; a.push(orE()); } } expect(')'); return a; }
    function prim() {
      const t = toks[p];
      if (!t) throw new Error('This line stops too early');
      if (t.t === 'num' || t.t === 'str') { p++; return t.v; }
      if (t.t === 'op' && t.v === '(') { p++; const v = orE(); expect(')'); return v; }
      if (t.t === 'id') {
        p++;
        const low = t.v.toLowerCase();
        if (low === 'rows' && isOp('.')) { p++; if (!isWord('count')) throw new Error('This demo only knows Rows.Count'); p++; return 1048576; }
        if (low === 'cells') {
          const [r, c] = args();
          const ref = colName(num(c)) + num(r);
          return afterCell(ref, num(c));
        }
        if (low === 'range') {
          const [a] = args();
          const ref = String(a).toUpperCase();
          return afterCell(ref, colNum(ref.replace(/\d+/g, '')));
        }
        if (low === 'len') return String(args()[0]).length;
        if (low === 'ucase') return String(args()[0]).toUpperCase();
        if (low === 'lcase') return String(args()[0]).toLowerCase();
        if (low === 'trim') return String(args()[0]).trim();
        if (low === 'round') { const [x, d] = args(); const k = Math.pow(10, d || 0); return Math.round(num(x) * k) / k; }
        if (Object.prototype.hasOwnProperty.call(CONSTANTS, low)) return CONSTANTS[low];
        if (Object.prototype.hasOwnProperty.call(st.vars, low)) return st.vars[low].v;
        throw new Error(t.v + ' has not been declared. Add a Dim line for it, or check the spelling');
      }
      throw new Error('This demo cannot read ' + t.v + ' here');
    }
    /* .Value is optional in VBA. .End(xlUp).Row finds the last filled row of the column. */
    function afterCell(ref, col) {
      if (isOp('.')) {
        p++;
        if (isWord('value')) { p++; return st.sheet[ref] === undefined ? '' : st.sheet[ref]; }
        if (isWord('end')) {
          p++; args(); expect('.');
          if (!isWord('row')) throw new Error('This demo only knows .End(xlUp).Row');
          p++;
          let last = 1;
          for (const k of Object.keys(st.sheet)) {
            const m = /^([A-Z]+)(\d+)$/.exec(k);
            if (m && colNum(m[1]) === col && st.sheet[k] !== '' && +m[2] > last) last = +m[2];
          }
          return last;
        }
        throw new Error('This demo only reads .Value and .End(xlUp).Row');
      }
      return st.sheet[ref] === undefined ? '' : st.sheet[ref];
    }
    const v = orE();
    if (p < toks.length) throw new Error('This demo cannot read "' + (toks[p].v) + '" here');
    return v;
  }

  /* ---------- Compile: every code line becomes one step ---------- */
  function compile(source) {
    const lines = source.replace(/\r/g, '').split('\n');
    const ops = lines.map((raw, i) => classify(raw, i));
    const stack = [];
    ops.forEach((op, i) => {
      if (!op) return;
      if (op.k === 'for' || op.k === 'do') stack.push(i);
      if (op.k === 'next') { const s = stack.pop(); if (s === undefined || ops[s].k !== 'for') throw new SyntaxError('Line ' + (i + 1) + ': Next has no For'); ops[s].end = i; op.start = s; }
      if (op.k === 'loop') { const s = stack.pop(); if (s === undefined || ops[s].k !== 'do') throw new SyntaxError('Line ' + (i + 1) + ': Loop has no Do'); ops[s].end = i; op.start = s; }
      if (op.k === 'if') stack.push(i);
      if (op.k === 'elseif' || op.k === 'else') {
        const s = stack[stack.length - 1];
        if (s === undefined || ops[s].k !== 'if') throw new SyntaxError('Line ' + (i + 1) + ': Else has no If');
        (ops[s].arms = ops[s].arms || []).push(i);
      }
      if (op.k === 'endif') {
        const s = stack.pop();
        if (s === undefined || ops[s].k !== 'if') throw new SyntaxError('Line ' + (i + 1) + ': End If has no If');
        ops[s].end = i;
        (ops[s].arms || []).forEach((a) => { ops[a].end = i; });
      }
    });
    if (stack.length) throw new SyntaxError('Line ' + (stack[stack.length - 1] + 1) + ' is not closed');
    return { lines, ops };
  }

  function classify(raw, i) {
    const line = raw.replace(/'.*$/, (m, off) => (raw.slice(0, off).split('"').length % 2 ? '' : m)).trim();
    if (!line) return null;
    let m;
    if (/^Sub\s+/i.test(line)) return { k: 'sub', name: line.replace(/^Sub\s+/i, '').replace(/\(.*/, '') };
    if (/^End\s+Sub$/i.test(line)) return { k: 'endsub' };
    if (/^Dim\s+/i.test(line)) return { k: 'dim', names: line.replace(/^Dim\s+/i, '').split(',').map((p) => { const q = p.trim().split(/\s+/); return { name: q[0], type: (q[2] || 'Variant') }; }) };
    if ((m = /^For\s+(\w+)\s*=\s*(.+?)\s+To\s+(.+?)(?:\s+Step\s+(.+))?$/i.exec(line))) return { k: 'for', v: m[1], from: m[2], to: m[3], step: m[4] || '1' };
    if (/^Next\b/i.test(line)) return { k: 'next' };
    if ((m = /^Do\s+While\s+(.+)$/i.exec(line))) return { k: 'do', cond: m[1] };
    if (/^Loop$/i.test(line)) return { k: 'loop' };
    if ((m = /^If\s+(.+?)\s+Then$/i.exec(line))) return { k: 'if', cond: m[1] };
    if ((m = /^ElseIf\s+(.+?)\s+Then$/i.exec(line))) return { k: 'elseif', cond: m[1] };
    if (/^Else$/i.test(line)) return { k: 'else' };
    if (/^End\s+If$/i.test(line)) return { k: 'endif' };
    if (/^MsgBox\b/i.test(line)) return { k: 'msg', expr: line.replace(/^MsgBox\s*/i, '') };
    if ((m = /^(Cells\s*\(.+\)|Range\s*\(.+\))\.Interior\.Color\s*=\s*(.+)$/i.exec(line))) return { k: 'color', target: m[1], expr: m[2] };
    if ((m = /^(Cells\s*\(.+\)|Range\s*\(.+\))(?:\.Value)?\s*=\s*(.+)$/i.exec(line))) return { k: 'setcell', target: m[1], expr: m[2] };
    if ((m = /^(\w+)\s*=\s*(.+)$/.exec(line))) return { k: 'setvar', v: m[1], expr: m[2] };
    throw new SyntaxError('Line ' + (i + 1) + ': this demo cannot run "' + line + '" yet');
  }

  function targetRef(target, st) {
    let m;
    if ((m = /^Cells\s*\((.+),(.+)\)$/i.exec(target))) return colName(Number(evaluate(m[2], st))) + Number(evaluate(m[1], st));
    if ((m = /^Range\s*\(\s*"([A-Za-z]+\d+)"\s*\)$/i.exec(target))) return m[1].toUpperCase();
    throw new Error('This demo cannot write to ' + target);
  }
  const show = (v) => (typeof v === 'string' ? '"' + v + '"' : v === true ? 'True' : v === false ? 'False' : String(v));
  const clone = (st) => ({ pc: st.pc, done: st.done, sheet: Object.assign({}, st.sheet), fill: Object.assign({}, st.fill), vars: JSON.parse(JSON.stringify(st.vars)), loops: JSON.parse(JSON.stringify(st.loops)), msgs: st.msgs.slice(), say: st.say, changed: st.changed, line: st.line, steps: st.steps });

  /**
   * machine = create(source, sheet). machine.state is a plain object.
   * step() runs the current line. back() returns to the state before it. reset() starts again.
   */
  function create(source, sheet) {
    const prog = compile(source);
    const next = (i) => { let j = i + 1; while (j < prog.ops.length && !prog.ops[j]) j++; return j < prog.ops.length ? j : -1; };
    const first = () => { let j = 0; while (j < prog.ops.length && !prog.ops[j]) j++; return j; };
    const initial = () => ({ pc: first(), line: first(), done: false, sheet: Object.assign({}, sheet), fill: {}, vars: {}, loops: {}, msgs: [], say: 'Ready. Nothing has run yet.', changed: [], steps: 0 });
    let st = initial();
    const history = [];
    const STEP_LIMIT = 5000;

    function exec() {
      const i = st.pc, op = prog.ops[i];
      st.changed = [];
      st.line = i;
      st.steps++;
      if (st.steps > STEP_LIMIT) throw new Error('This macro ran more than ' + STEP_LIMIT + ' lines, so the demo stopped it. Check that the loop can end.');
      const setv = (name, v) => { st.vars[name.toLowerCase()] = { name: (st.vars[name.toLowerCase()] || {}).name || name, v, type: (st.vars[name.toLowerCase()] || {}).type || 'Variant' }; };
      switch (op.k) {
        case 'sub': st.say = 'The macro ' + op.name + ' starts.'; st.pc = next(i); break;
        case 'endsub': st.say = 'The macro is finished.'; st.done = true; break;
        case 'dim':
          op.names.forEach((d) => { const number = /long|integer|double|single|currency/i.test(d.type); st.vars[d.name.toLowerCase()] = { name: d.name, type: d.type, v: number ? 0 : '' }; });
          st.say = 'Excel sets aside ' + op.names.length + (op.names.length === 1 ? ' variable, ' : ' variables, ') + op.names.map((d) => d.name).join(' and ') + '. Numbers start at 0.';
          st.pc = next(i); break;
        case 'setvar': { const v = evaluate(op.expr, st); if (!Object.prototype.hasOwnProperty.call(st.vars, op.v.toLowerCase())) throw new Error(op.v + ' has not been declared. Add a Dim line for it, or check the spelling'); setv(op.v, v); st.say = op.v + ' is now ' + show(v) + '.'; st.pc = next(i); break; }
        case 'setcell': { const ref = targetRef(op.target, st); const v = evaluate(op.expr, st); st.sheet[ref] = v; st.changed = [ref]; st.say = ref + ' now holds ' + show(v) + '.'; st.pc = next(i); break; }
        case 'color': { const ref = targetRef(op.target, st); const v = evaluate(op.expr, st); st.fill[ref] = v; st.changed = [ref]; st.say = ref + ' is now colored ' + v + '.'; st.pc = next(i); break; }
        case 'msg': { const v = evaluate(op.expr, st); st.msgs.push(String(v)); st.say = 'A message box opens: ' + v; st.pc = next(i); break; }
        case 'for': {
          const from = Number(evaluate(op.from, st)), to = Number(evaluate(op.to, st)), step = Number(evaluate(op.step, st));
          st.loops[i] = { to, step };
          setv(op.v, from);
          if ((step > 0 && from > to) || (step < 0 && from < to)) { st.say = op.v + ' starts past the end, so the loop is skipped.'; st.pc = next(op.end); }
          else { st.say = 'The loop starts with ' + op.v + ' = ' + from + '. It runs until ' + op.v + ' passes ' + to + '.'; st.pc = next(i); }
          break;
        }
        case 'next': {
          const f = prog.ops[op.start], lp = st.loops[op.start];
          const v = st.vars[f.v.toLowerCase()].v + lp.step;
          setv(f.v, v);
          if ((lp.step > 0 && v > lp.to) || (lp.step < 0 && v < lp.to)) { st.say = f.v + ' is now ' + v + ', past ' + lp.to + ', so the loop ends.'; st.pc = next(i); }
          else { st.say = f.v + ' is now ' + v + '. Back to the top of the loop.'; st.pc = next(op.start); }
          break;
        }
        case 'do': { const c = evaluate(op.cond, st); st.say = op.cond + ' is ' + show(!!c) + (c ? ', so the loop runs.' : ', so the loop ends.'); st.pc = c ? next(i) : next(op.end); break; }
        case 'loop': st.say = 'Back to the Do line.'; st.pc = op.start; break;
        case 'if': case 'elseif': {
          const c = evaluate(op.cond, st);
          st.say = describe(op.cond, st) + (c ? ', so this branch runs.' : ', so this branch is skipped.');
          if (c) { st.pc = next(i); break; }
          const owner = op.k === 'if' ? op : prog.ops.find((o, n) => o && o.k === 'if' && (o.arms || []).includes(i));
          const arms = owner.arms || [];
          const after = arms.find((a) => a > i);
          st.pc = after !== undefined ? after : next(op.end) === -1 ? op.end : op.end;
          break;
        }
        case 'else': st.say = 'No earlier branch ran, so the Else branch runs.'; st.pc = next(i); break;
        case 'endif': st.say = 'The If block ends.'; st.pc = next(i); break;
        default: throw new Error('Unknown step');
      }
      /* A branch that has run skips the rest of its If when it reaches the next arm. */
      if (!st.done && prog.ops[st.pc] && (prog.ops[st.pc].k === 'else' || prog.ops[st.pc].k === 'elseif') && op.k !== 'if' && op.k !== 'elseif') st.pc = prog.ops[st.pc].end;
    }

    function describe(cond, s) {
      const m = /^(.+?)\s*(<>|<=|>=|=|<|>)\s*(.+)$/.exec(cond);
      if (!m) return cond + ' is ' + show(!!evaluate(cond, s));
      return show(evaluate(m[1], s)) + ' ' + m[2] + ' ' + show(evaluate(m[3], s)) + ' is ' + show(!!evaluate(cond, s));
    }

    return {
      lines: prog.lines,
      get state() { return st; },
      get canBack() { return history.length > 0; },
      step() {
        if (st.done || st.error) return st;
        history.push(clone(st));
        try { exec(); } catch (e) { st.error = e.message; st.say = 'The macro stopped here: ' + e.message; st.done = true; }
        return st;
      },
      back() { if (history.length) st = history.pop(); return st; },
      reset() { history.length = 0; st = initial(); return st; },
      runToEnd() { while (!st.done) this.step(); return st; }
    };
  }

  return { create, compile, colName, colNum };
});
