// Gridwork formula engine: tokenizer, Pratt parser, evaluator, workbooks with cross-sheet refs, dates, ~70 functions.
// Plain script: sets window.Engine in the browser, module.exports in Node.
(function (root) {
'use strict';

// ---------- values ----------
const ERR_HELP = {
  '#DIV/0!': 'Division by zero (or by a blank cell).',
  '#VALUE!': 'A value has the wrong type, e.g. text where a number is expected.',
  '#REF!': 'The formula points at a cell that does not exist.',
  '#NAME?': 'Unknown function or name.',
  '#N/A': 'The lookup value was not found.',
  '#NUM!': 'The result is not a valid number.',
  '#CYCLE!': 'This formula depends on itself (circular reference).',
  '#ERROR!': 'The formula could not be read.',
};
class Err {
  constructor(code, msg) { this.code = code; this.msg = msg || ERR_HELP[code] || ''; }
  toString() { return this.code; }
}
const isErr = v => v instanceof Err;
class RangeVal {
  constructor(rows) { this.rows = rows; }
  flat() { return this.rows.flat(); }
}

// ---------- addresses ----------
function colName(c) { let s = ''; for (c++; c > 0; c = (c - 1) / 26 | 0) s = String.fromCharCode(65 + (c - 1) % 26) + s; return s; }
function colIndex(s) { let c = 0; for (const ch of s.toUpperCase()) c = c * 26 + ch.charCodeAt(0) - 64; return c - 1; }
const key = (r, c) => colName(c) + (r + 1);
function parseKey(k) {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(String(k).trim());
  return m && +m[2] > 0 ? { r: +m[2] - 1, c: colIndex(m[1]) } : null;
}

// ---------- literals & coercion ----------
function literal(raw) {
  if (raw[0] === "'") return raw.slice(1);
  const s = raw.trim();
  if (/^(true|false)$/i.test(s)) return s.toUpperCase() === 'TRUE';
  const t = s.replace(/,(?=\d{3}(\D|$))/g, ''), pct = t.endsWith('%'), n = pct ? t.slice(0, -1) : t;
  if (n !== '' && /^[+-]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?$/i.test(n)) return pct ? +n / 100 : +n;
  const d = parseDate(s);
  return d == null ? raw : d;
}

// ---------- dates: serial numbers like Excel (1 = 1900-01-01, whole days, time as the fraction) ----------
const config = { dateOrder: 'mdy' }; // how 3/4/2026 reads; the UI sets 'dmy' for day-first locales
const DAY = 864e5, DAY0 = Date.UTC(1899, 11, 30);
const MONTHS = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'];
const WEEKDAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
const serial = (y, m, d) => (Date.UTC(y, m - 1, d) - DAY0) / DAY;
const toDate = n => new Date(DAY0 + Math.round(n * DAY)); // read it with getUTC*
const pad2 = x => String(x).padStart(2, '0');
function monthOf(name) {
  const n = name.toLowerCase().replace(/\.$/, '');
  return n.length >= 3 ? MONTHS.findIndex(m => m.startsWith(n)) + 1 : 0;
}
// "2026-09-22", "22/09/2026" (day-first locales), "9/22/2026", "22 Sep 2026", "Sep 22, 2026", optional " 14:30" or " 2:30 pm"
function parseDate(s) {
  s = String(s).trim();
  let m, y, mo, d, time = 0;
  const tm = /\s+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(am|pm)?$/i.exec(s);
  if (tm) {
    let h = +tm[1];
    if (tm[4]) h = h % 12 + (tm[4].toLowerCase() === 'pm' ? 12 : 0);
    if (h > 23 || +tm[2] > 59 || +(tm[3] || 0) > 59) return null;
    time = (h * 3600 + +tm[2] * 60 + +(tm[3] || 0)) / 86400;
    s = s.slice(0, tm.index);
  }
  if ((m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(s))) [y, mo, d] = [+m[1], +m[2], +m[3]];
  else if ((m = /^(\d{1,2})[/.](\d{1,2})[/.](\d{4}|\d{2})$/.exec(s))) [mo, d, y] = config.dateOrder === 'dmy' ? [+m[2], +m[1], +m[3]] : [+m[1], +m[2], +m[3]];
  else if ((m = /^(\d{1,2})[ -]([A-Za-z]{3,9}\.?)[ -]+(\d{4})$/.exec(s))) [d, mo, y] = [+m[1], monthOf(m[2]), +m[3]];
  else if ((m = /^([A-Za-z]{3,9}\.?) (\d{1,2}),? (\d{4})$/.exec(s))) [mo, d, y] = [monthOf(m[1]), +m[2], +m[3]];
  else return null;
  if (y < 100) y += y < 50 ? 2000 : 1900;
  if (!mo || mo > 12 || d < 1 || new Date(Date.UTC(y, mo - 1, d)).getUTCDate() !== d) return null;
  return serial(y, mo, d) + time;
}
function isoDate(n) {
  const t = toDate(n), day = `${t.getUTCFullYear()}-${pad2(t.getUTCMonth() + 1)}-${pad2(t.getUTCDate())}`;
  return n % 1 ? `${day} ${pad2(t.getUTCHours())}:${pad2(t.getUTCMinutes())}` : day;
}
// TEXT()'s date codes: yyyy yy mmmm mmm mm m dddd ddd dd d hh h mm(after h) ss am/pm, "quoted literals"
function fmtDate(n, fmt) {
  const t = toDate(n), h = t.getUTCHours(), ampm = /am\/pm/i.test(fmt);
  let afterHour = false;
  return fmt.replace(/"([^"]*)"|yyyy|yy|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|am\/pm/gi, (tok, lit) => {
    if (lit !== undefined) return lit;
    const k = tok.toLowerCase();
    if (k[0] === 'h') { afterHour = true; const hh = ampm ? h % 12 || 12 : h; return k === 'hh' ? pad2(hh) : String(hh); }
    if ((k === 'mm' || k === 'm') && afterHour) { afterHour = false; return k === 'mm' ? pad2(t.getUTCMinutes()) : String(t.getUTCMinutes()); }
    const M = t.getUTCMonth(), D = t.getUTCDate(), W = WEEKDAYS[t.getUTCDay()], name = MONTHS[M][0].toUpperCase() + MONTHS[M].slice(1);
    return { yyyy: t.getUTCFullYear(), yy: pad2(t.getUTCFullYear() % 100), mmmm: name, mmm: name.slice(0, 3), mm: pad2(M + 1), m: M + 1,
      dddd: W, ddd: W.slice(0, 3), dd: pad2(D), d: D, ss: pad2(t.getUTCSeconds()), s: t.getUTCSeconds(), 'am/pm': h < 12 ? 'AM' : 'PM' }[k];
  });
}
function todaySerial() { const t = new Date(); return serial(t.getFullYear(), t.getMonth() + 1, t.getDate()); }
function fmtGeneral(n) {
  if (Number.isInteger(n) && Math.abs(n) < 1e15) return String(n);
  return String(+n.toPrecision(10));
}
function scalar(v) {
  if (!(v instanceof RangeVal)) return v;
  return v.rows.length === 1 && v.rows[0].length === 1 ? v.rows[0][0] : new Err('#VALUE!', 'Expected one cell but got a range. Wrap it in a function like SUM().');
}
function toNum(v) {
  v = scalar(v);
  if (typeof v === 'number' || isErr(v)) return v;
  if (v == null) return 0;
  if (typeof v === 'boolean') return +v;
  const n = literal(v);
  return typeof n === 'number' ? n : new Err('#VALUE!', `"${v}" is not a number.`);
}
function toStr(v) {
  v = scalar(v);
  return v == null ? '' : typeof v === 'number' ? fmtGeneral(v) : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v);
}
function toBool(v) {
  v = scalar(v);
  if (isErr(v) || typeof v === 'boolean') return v;
  if (v == null) return false;
  if (typeof v === 'number') return v !== 0;
  const u = v.toUpperCase();
  return u === 'TRUE' ? true : u === 'FALSE' ? false : new Err('#VALUE!', `"${v}" is not TRUE or FALSE.`);
}
const RANK = { number: 0, string: 1, boolean: 2 };
function compare(a, b) {
  if (a == null) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
  if (b == null) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
  if (typeof a !== typeof b) return RANK[typeof a] - RANK[typeof b];
  if (typeof a === 'string') { a = a.toLowerCase(); b = b.toLowerCase(); }
  return a < b ? -1 : a > b ? 1 : 0;
}

// ---------- tokenizer ----------
const TOKEN = /\s*(?:(?<num>\d+\.?\d*(?:[eE][+-]?\d+)?|\.\d+(?:[eE][+-]?\d+)?)|(?<str>"(?:[^"]|"")*")|(?<ref>(?<pre>(?:[A-Za-z_][\w.]*|'(?:[^']|'')+')!)?\$?[A-Za-z]{1,3}\$?\d+)(?![\w.(!])|(?<id>[A-Za-z_][\w.]*)|(?<err>#(?:REF!|DIV\/0!|N\/A|VALUE!|NAME\?|NUM!|CYCLE!|ERROR!))|(?<op><=|>=|<>|[-+*/^&=<>%(),:]))/y;
const KINDS = ['num', 'str', 'ref', 'id', 'err', 'op'];
const unquote = s => s[0] === "'" ? s.slice(1, -1).replace(/''/g, "'") : s;
// Sheet names need quotes in formulas unless they look like a plain identifier (and not like a cell).
const quoteSheet = name => /^[A-Za-z_][\w.]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name) ? name : "'" + name.replace(/'/g, "''") + "'";
function tokenize(src) {
  const out = [];
  TOKEN.lastIndex = 0;
  while (TOKEN.lastIndex < src.length) {
    const at = TOKEN.lastIndex;
    if (!src.slice(at).trim()) break;
    const m = TOKEN.exec(src);
    if (!m) {
      const ch = src.slice(at).trim()[0];
      throw new SyntaxError(ch === '"' ? 'Text is missing its closing quote.' : `Unexpected "${ch}".`);
    }
    const g = m.groups, k = KINDS.find(n => g[n] !== undefined), v = g[k];
    const tok = { k, v, s: TOKEN.lastIndex - v.length, e: TOKEN.lastIndex };
    if (k === 'ref') { tok.pre = g.pre || ''; tok.cell = v.slice(tok.pre.length); if (g.pre) tok.sheet = unquote(g.pre.slice(0, -1)); }
    out.push(tok);
  }
  return out;
}

// ---------- parser (Pratt; Excel precedence) ----------
const BIN = { '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '&': 2, '+': 3, '-': 3, '*': 4, '/': 4, '^': 5 };
function refNode(tok) {
  const m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/.exec(tok.cell);
  return +m[4] < 1 ? { t: 'err', v: '#REF!' } : { t: 'ref', r: +m[4] - 1, c: colIndex(m[2]), ac: !!m[1], ar: !!m[3], sheet: tok.sheet };
}
function parse(src) {
  const toks = tokenize(src);
  let i = 0;
  const peek = () => toks[i], next = () => toks[i++];
  const fail = msg => { throw new SyntaxError(msg); };
  const expect = v => { const t = next(); if (!t || t.v !== v) fail(t ? `Expected "${v}" but found "${t.v}".` : `Missing "${v}".`); };

  function prefix() {
    const t = next();
    if (!t) fail('The formula ends too early.');
    if (t.k === 'num') return { t: 'num', v: +t.v };
    if (t.k === 'str') return { t: 'str', v: t.v.slice(1, -1).replace(/""/g, '"') };
    if (t.k === 'err') return { t: 'err', v: t.v };
    if (t.k === 'ref') return refNode(t);
    if (t.k === 'id') {
      const name = t.v.toUpperCase();
      if (peek() && peek().v === '(') {
        next();
        const args = [];
        if (peek() && peek().v === ')') next();
        else { do args.push(expr(0)); while (peek() && peek().v === ',' && next()); expect(')'); }
        return { t: 'call', name, args };
      }
      if (name === 'TRUE' || name === 'FALSE') return { t: 'bool', v: name === 'TRUE' };
      return { t: 'err', v: '#NAME?', msg: `Unknown name "${t.v}". Text needs "quotes".` };
    }
    if (t.v === '(') { const e = expr(0); expect(')'); return e; }
    if (t.v === '-') return { t: 'neg', e: expr(7) };
    if (t.v === '+') return expr(7);
    fail(`Unexpected "${t.v}".`);
  }
  function expr(min) {
    let left = prefix();
    for (let t; (t = peek()) && t.k === 'op';) {
      if (t.v === ':') {
        if (min > 8) break;
        next();
        const right = prefix();
        if (left.t !== 'ref' || right.t !== 'ref') fail('A range needs a cell on both sides of ":".');
        if (right.sheet && right.sheet.toLowerCase() !== (left.sheet || '').toLowerCase()) fail('A range has to stay on one sheet.');
        left = { t: 'range', a: left, b: right, sheet: left.sheet };
      } else if (t.v === '%') {
        if (min > 6) break;
        next();
        left = { t: 'pct', e: left };
      } else {
        const bp = BIN[t.v];
        if (!bp || bp < min) break;
        next();
        left = { t: 'bin', op: t.v, l: left, r: expr(bp + 1) };
      }
    }
    return left;
  }
  const ast = expr(0);
  if (i < toks.length) fail(`Unexpected "${toks[i].v}".`);
  return ast;
}

// ---------- evaluator ----------
function rangeOf(sh, a, b) {
  const r0 = Math.min(a.r, b.r), r1 = Math.max(a.r, b.r), c0 = Math.min(a.c, b.c), c1 = Math.max(a.c, b.c);
  const rows = [];
  for (let r = r0; r <= r1; r++) {
    const row = [];
    for (let c = c0; c <= c1; c++) row.push(sh.value(key(r, c)));
    rows.push(row);
  }
  return new RangeVal(rows);
}
function binop(op, a, b) {
  a = scalar(a); b = scalar(b);
  if (isErr(a)) return a;
  if (isErr(b)) return b;
  if (op === '&') return toStr(a) + toStr(b);
  if (BIN[op] === 1) {
    const c = compare(a, b);
    return op === '=' ? c === 0 : op === '<>' ? c !== 0 : op === '<' ? c < 0 : op === '>' ? c > 0 : op === '<=' ? c <= 0 : c >= 0;
  }
  const x = toNum(a), y = toNum(b);
  if (isErr(x)) return x;
  if (isErr(y)) return y;
  if (op === '+') return x + y;
  if (op === '-') return x - y;
  if (op === '*') return x * y;
  if (op === '/') return y === 0 ? new Err('#DIV/0!') : x / y;
  return x ** y;
}
function evaluate(n, sh) {
  switch (n.t) {
    case 'num': case 'str': case 'bool': return n.v;
    case 'err': return new Err(n.v, n.msg);
    case 'ref': case 'range': {
      const target = n.sheet ? sh.book.byName(n.sheet) : sh;
      if (!target) return new Err('#REF!', `There is no sheet called "${n.sheet}".`);
      return n.t === 'ref' ? target.value(key(n.r, n.c)) : rangeOf(target, n.a, n.b);
    }
    case 'neg': { const x = toNum(evaluate(n.e, sh)); return isErr(x) ? x : -x; }
    case 'pct': { const x = toNum(evaluate(n.e, sh)); return isErr(x) ? x : x / 100; }
    case 'bin': return binop(n.op, evaluate(n.l, sh), evaluate(n.r, sh));
    case 'call': {
      if (LAZY[n.name]) return LAZY[n.name](n.args, a => evaluate(a, sh));
      const fn = FUNCS[n.name];
      if (!fn) return new Err('#NAME?', `There is no function called ${n.name}().`);
      const min = fn.min ?? fn.length;
      if (n.args.length < min) return new Err('#VALUE!', `${n.name}() needs at least ${min} argument${min > 1 ? 's' : ''}.`);
      // A plain reference is passed as a 1x1 range, so SUM(A1) ignores text like SUM(A1:A1) does.
      return fn(...n.args.map(a => a.t === 'ref' ? new RangeVal([[evaluate(a, sh)]]) : evaluate(a, sh)));
    }
  }
}
function finalValue(v) {
  v = scalar(v);
  if (v == null) return 0;
  if (typeof v === 'number') return Number.isFinite(v) ? v + 0 : new Err('#NUM!');
  return v;
}

// ---------- functions ----------
const sum = xs => xs.reduce((a, b) => a + b, 0);
// Excel rule: typed-in arguments are coerced; inside ranges only real numbers count.
function nums(args) {
  const out = [];
  for (const a of args) {
    if (a instanceof RangeVal) {
      for (const v of a.flat()) { if (isErr(v)) return v; if (typeof v === 'number') out.push(v); }
    } else {
      const n = toNum(a);
      if (isErr(n)) return n;
      out.push(n);
    }
  }
  return out;
}
const agg = f => { const w = (...a) => { const xs = nums(a); return isErr(xs) ? xs : f(xs); }; w.min = 1; return w; };
const CONV = { n: toNum, s: v => (v = scalar(v), isErr(v) ? v : toStr(v)), b: toBool, v: scalar };
function typed(sig, f) {
  const w = (...args) => {
    const xs = args.map((a, i) => CONV[sig[Math.min(i, sig.length - 1)]](a));
    return xs.find(isErr) || f(...xs);
  };
  w.min = f.length;
  return w;
}
function bools(args, f) {
  const xs = [];
  for (const a of args) {
    if (a instanceof RangeVal) {
      for (const v of a.flat()) { if (isErr(v)) return v; if (typeof v === 'boolean' || typeof v === 'number') xs.push(!!v); }
    } else {
      const b = toBool(a);
      if (isErr(b)) return b;
      xs.push(b);
    }
  }
  return xs.length ? f(xs) : new Err('#VALUE!', 'No TRUE/FALSE values to test.');
}
function roundTo(x, d, f) {
  const s = Math.sign(x), a = Math.abs(x);
  d = Math.trunc(d);
  let r = Number(f(+(a + 'e' + d)) + 'e' + -d); // string exponent dodges 1.005*100 = 100.4999…
  if (Number.isNaN(r)) r = f(a * 10 ** d) / 10 ** d;
  return s * r;
}
const vec = r => r instanceof RangeVal ? r.flat() : [scalar(r)];
function wild(s) {
  return new RegExp('^' + s.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.') + '$', 'i');
}
// COUNTIF-style criteria: 5, "apple", "app*", ">=10", "<>done", "" (blank), "<>" (not blank)
function criterion(c) {
  c = scalar(c);
  if (typeof c !== 'string') return v => v != null && compare(v, c) === 0;
  const m = /^(<=|>=|<>|<|>|=)?([\s\S]*)$/.exec(c), op = m[1] || '=', rhs = m[2] === '' ? '' : literal(m[2]);
  if (typeof rhs === 'string' && (op === '=' || op === '<>')) {
    const re = wild(rhs);
    const hit = v => rhs === '' ? v == null || v === '' : v != null && re.test(toStr(v));
    return op === '=' ? hit : v => !hit(v);
  }
  return v => {
    if (v == null || isErr(v) || (typeof rhs === 'number' && typeof v !== 'number')) return false;
    const k = compare(v, rhs);
    return op === '=' ? k === 0 : op === '<>' ? k !== 0 : op === '<' ? k < 0 : op === '>' ? k > 0 : op === '<=' ? k <= 0 : k >= 0;
  };
}
function matches(pairs) {
  if (!pairs.length || pairs.length % 2) return new Err('#VALUE!', 'Criteria come in range, criterion pairs.');
  let ok = null;
  for (let i = 0; i < pairs.length; i += 2) {
    const v = vec(pairs[i]), t = criterion(pairs[i + 1]);
    if (ok && v.length !== ok.length) return new Err('#VALUE!', 'All criteria ranges must be the same size.');
    ok = v.map((x, j) => (!ok || ok[j]) && t(x));
  }
  return ok;
}
function pick(sumRange, pairs) {
  const ok = matches(pairs);
  if (isErr(ok)) return ok;
  return vec(sumRange).filter((x, i) => ok[i] && typeof x === 'number');
}
// mode 0: exact (wildcards ok) · 1: largest <= v in ascending data · -1: smallest >= v in descending data
function lookup(list, v, mode) {
  if (mode === 0) {
    if (typeof v === 'string' && /[*?]/.test(v)) { const re = wild(v); return list.findIndex(x => typeof x === 'string' && re.test(x)); }
    return list.findIndex(x => x != null && !isErr(x) && compare(x, v) === 0);
  }
  let best = -1;
  for (let i = 0; i < list.length; i++) {
    if (list[i] == null) continue;
    const c = compare(list[i], v);
    if (mode > 0 ? c <= 0 : c >= 0) best = i; else break;
  }
  return best;
}
const notFound = v => new Err('#N/A', `"${toStr(v)}" was not found.`);
const withMin = (min, f) => (f.min = min, f);

const FUNCS = {
  SUM: agg(sum),
  AVERAGE: agg(xs => xs.length ? sum(xs) / xs.length : new Err('#DIV/0!', 'AVERAGE() of no numbers.')),
  MIN: agg(xs => xs.length ? xs.reduce((a, b) => Math.min(a, b)) : 0),
  MAX: agg(xs => xs.length ? xs.reduce((a, b) => Math.max(a, b)) : 0),
  PRODUCT: agg(xs => xs.reduce((a, b) => a * b, 1)),
  MEDIAN: agg(xs => {
    if (!xs.length) return new Err('#NUM!');
    const s = [...xs].sort((a, b) => a - b), m = s.length >> 1;
    return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
  }),
  STDEV: agg(xs => {
    if (xs.length < 2) return new Err('#DIV/0!', 'STDEV() needs at least two numbers.');
    const m = sum(xs) / xs.length;
    return Math.sqrt(sum(xs.map(x => (x - m) ** 2)) / (xs.length - 1));
  }),
  COUNT: withMin(1, (...a) => a.reduce((n, x) => n + (x instanceof RangeVal ? x.flat().filter(v => typeof v === 'number').length : typeof toNum(x) === 'number' ? 1 : 0), 0)),
  COUNTA: withMin(1, (...a) => a.reduce((n, x) => n + (x instanceof RangeVal ? x.flat().filter(v => v != null && v !== '').length : 1), 0)),
  COUNTBLANK: r => vec(r).filter(v => v == null || v === '').length,
  SUMPRODUCT: withMin(1, (...rs) => {
    const vs = rs.map(vec), n = vs[0].length;
    if (vs.some(v => v.length !== n)) return new Err('#VALUE!', 'SUMPRODUCT() ranges must be the same size.');
    let s = 0;
    for (let i = 0; i < n; i++) {
      let p = 1;
      for (const v of vs) { if (isErr(v[i])) return v[i]; p *= typeof v[i] === 'number' ? v[i] : 0; }
      s += p;
    }
    return s;
  }),
  SUMIF: withMin(2, (range, crit, sumRange) => { const xs = pick(sumRange ?? range, [range, crit]); return isErr(xs) ? xs : sum(xs); }),
  SUMIFS: withMin(3, (sumRange, ...pairs) => { const xs = pick(sumRange, pairs); return isErr(xs) ? xs : sum(xs); }),
  AVERAGEIF: withMin(2, (range, crit, avgRange) => {
    const xs = pick(avgRange ?? range, [range, crit]);
    return isErr(xs) ? xs : xs.length ? sum(xs) / xs.length : new Err('#DIV/0!', 'No cells matched.');
  }),
  COUNTIF: (range, crit) => { const ok = matches([range, crit]); return isErr(ok) ? ok : ok.filter(Boolean).length; },
  COUNTIFS: withMin(2, (...pairs) => { const ok = matches(pairs); return isErr(ok) ? ok : ok.filter(Boolean).length; }),

  ROUND: typed('nn', (x, d = 0) => roundTo(x, d, Math.round)),
  ROUNDUP: typed('nn', (x, d = 0) => roundTo(x, d, Math.ceil)),
  ROUNDDOWN: typed('nn', (x, d = 0) => roundTo(x, d, Math.floor)),
  INT: typed('n', Math.floor),
  ABS: typed('n', Math.abs),
  SIGN: typed('n', Math.sign),
  SQRT: typed('n', x => x < 0 ? new Err('#NUM!', 'SQRT() of a negative number.') : Math.sqrt(x)),
  POWER: typed('nn', (x, y) => x ** y),
  MOD: typed('nn', (x, y) => y === 0 ? new Err('#DIV/0!') : x - y * Math.floor(x / y)),
  EXP: typed('n', Math.exp),
  LN: typed('n', x => x <= 0 ? new Err('#NUM!') : Math.log(x)),
  LOG10: typed('n', x => x <= 0 ? new Err('#NUM!') : Math.log10(x)),
  PI: () => Math.PI,
  RAND: () => Math.random(),
  RANDBETWEEN: typed('nn', (a, b) => { a = Math.ceil(a); b = Math.floor(b); return b < a ? new Err('#NUM!') : a + Math.floor(Math.random() * (b - a + 1)); }),

  AND: withMin(1, (...a) => bools(a, xs => xs.every(Boolean))),
  OR: withMin(1, (...a) => bools(a, xs => xs.some(Boolean))),
  NOT: typed('b', b => !b),
  ISBLANK: v => scalar(v) == null,
  ISNUMBER: v => typeof scalar(v) === 'number',
  ISTEXT: v => typeof scalar(v) === 'string',
  ISERROR: v => isErr(scalar(v)),

  CONCAT: withMin(1, (...a) => {
    let s = '';
    for (const v of a.flatMap(vec)) { if (isErr(v)) return v; s += toStr(v); }
    return s;
  }),
  LEN: typed('s', s => s.length),
  UPPER: typed('s', s => s.toUpperCase()),
  LOWER: typed('s', s => s.toLowerCase()),
  PROPER: typed('s', s => s.toLowerCase().replace(/(^|[^a-z])([a-z])/g, (m, a, b) => a + b.toUpperCase())),
  TRIM: typed('s', s => s.trim().replace(/ +/g, ' ')),
  LEFT: typed('sn', (s, n = 1) => n < 0 ? new Err('#VALUE!') : s.slice(0, n)),
  RIGHT: typed('sn', (s, n = 1) => n < 0 ? new Err('#VALUE!') : n ? s.slice(-n) : ''),
  MID: typed('snn', (s, start, n) => start < 1 || n < 0 ? new Err('#VALUE!') : s.substr(start - 1, n)),
  FIND: typed('ssn', (needle, hay, start = 1) => { const i = hay.indexOf(needle, start - 1); return i < 0 ? new Err('#VALUE!', `"${needle}" was not found.`) : i + 1; }),
  SUBSTITUTE: typed('sss', (s, a, b) => a ? s.split(a).join(b) : s),
  REPT: typed('sn', (s, n) => n < 0 ? new Err('#VALUE!') : s.repeat(n)),
  VALUE: typed('n', x => x),

  TODAY: () => todaySerial(),
  NOW: () => { const t = new Date(); return todaySerial() + (t.getHours() * 3600 + t.getMinutes() * 60 + t.getSeconds()) / 86400; },
  DATE: typed('nnn', (y, m, d) => { y = Math.trunc(y); if (y < 100) y += 1900; const v = serial(y, Math.trunc(m), Math.trunc(d)); return v < 0 ? new Err('#NUM!') : v; }),
  YEAR: typed('n', n => toDate(n).getUTCFullYear()),
  MONTH: typed('n', n => toDate(n).getUTCMonth() + 1),
  DAY: typed('n', n => toDate(n).getUTCDate()),
  WEEKDAY: typed('nn', (n, type = 1) => { const w = toDate(n).getUTCDay(); return type === 2 ? (w + 6) % 7 + 1 : type === 3 ? (w + 6) % 7 : w + 1; }),
  EDATE: typed('nn', (n, months) => {
    const t = toDate(n), y = t.getUTCFullYear(), m = t.getUTCMonth() + Math.trunc(months);
    return serial(y, m + 1, Math.min(t.getUTCDate(), new Date(Date.UTC(y, m + 1, 0)).getUTCDate()));
  }),
  EOMONTH: typed('nn', (n, months) => { const t = toDate(n); return (Date.UTC(t.getUTCFullYear(), t.getUTCMonth() + Math.trunc(months) + 1, 0) - DAY0) / DAY; }),
  DAYS: typed('nn', (end, start) => Math.floor(end) - Math.floor(start)),
  DATEDIF: typed('nns', (a, b, unit) => {
    if (b < a) return new Err('#NUM!', 'The start date is after the end date.');
    const u = unit.toUpperCase(), x = toDate(a), y = toDate(b);
    if (u === 'D') return Math.floor(b) - Math.floor(a);
    const months = (y.getUTCFullYear() - x.getUTCFullYear()) * 12 + y.getUTCMonth() - x.getUTCMonth() - (y.getUTCDate() < x.getUTCDate() ? 1 : 0);
    return u === 'M' ? months : u === 'Y' ? Math.floor(months / 12) : new Err('#NUM!', 'The unit must be "Y", "M" or "D".');
  }),
  NETWORKDAYS: typed('nn', (a, b) => {
    const s = Math.floor(Math.min(a, b)), e = Math.floor(Math.max(a, b));
    if (e - s > 100000) return new Err('#NUM!');
    let n = 0;
    for (let d = s; d <= e; d++) { const w = toDate(d).getUTCDay(); if (w > 0 && w < 6) n++; }
    return a <= b ? n : -n;
  }),
  TEXT: withMin(2, (v, fmt) => {
    v = scalar(v); fmt = CONV.s(fmt);
    if (isErr(v)) return v;
    if (isErr(fmt)) return fmt;
    const n = toNum(v);
    if (isErr(n)) return toStr(v);
    if (/[ymdhs]/i.test(fmt.replace(/"[^"]*"/g, ''))) return fmtDate(n, fmt);
    const m = /^([^#0,.%]*)([#0,]*)(?:\.(0+|#+))?(%?)(.*)$/.exec(fmt);
    if (!m || (!m[2] && !m[3])) return toStr(v);
    const pct = m[4] === '%', dec = m[3] ? m[3].length : 0;
    return m[1] + new Intl.NumberFormat('en-US', {
      minimumIntegerDigits: Math.max(1, (m[2].match(/0/g) || []).length),
      minimumFractionDigits: m[3] && m[3][0] === '0' ? dec : 0, maximumFractionDigits: dec, useGrouping: m[2].includes(','),
    }).format(pct ? n * 100 : n) + (pct ? '%' : '') + m[5];
  }),

  VLOOKUP: (v, table, col, approx = true) => {
    v = scalar(v); col = toNum(col); approx = toBool(approx);
    if (isErr(v)) return v;
    if (isErr(col)) return col;
    if (isErr(approx)) return approx;
    if (!(table instanceof RangeVal)) return new Err('#VALUE!', 'VLOOKUP() needs a range to search.');
    if (col < 1 || col > table.rows[0].length) return new Err('#REF!', 'The column number is outside the range.');
    const i = lookup(table.rows.map(r => r[0]), v, approx ? 1 : 0);
    return i < 0 ? notFound(v) : table.rows[i][Math.trunc(col) - 1];
  },
  XLOOKUP: withMin(3, (v, look, ret, ifNone) => {
    v = scalar(v);
    if (isErr(v)) return v;
    const i = lookup(vec(look), v, 0);
    if (i < 0) return ifNone !== undefined ? ifNone : notFound(v);
    const out = vec(ret)[i];
    return out === undefined ? new Err('#REF!', 'The return range is shorter than the lookup range.') : out;
  }),
  MATCH: (v, range, type = 1) => {
    v = scalar(v); type = toNum(type);
    if (isErr(v)) return v;
    if (isErr(type)) return type;
    const i = lookup(vec(range), v, Math.sign(type));
    return i < 0 ? notFound(v) : i + 1;
  },
  INDEX: withMin(2, (range, r, c) => {
    if (!(range instanceof RangeVal)) return new Err('#VALUE!', 'INDEX() needs a range.');
    if (c === undefined && range.rows.length === 1) [r, c] = [1, r];
    r = toNum(r); c = c === undefined ? 1 : toNum(c);
    if (isErr(r)) return r;
    if (isErr(c)) return c;
    const v = (range.rows[Math.trunc(r) - 1] || [])[Math.trunc(c) - 1];
    return v === undefined ? new Err('#REF!', 'The position is outside the range.') : v;
  }),
};
FUNCS.CONCATENATE = FUNCS.CONCAT;
FUNCS.AVG = FUNCS.AVERAGE;

// IF / IFERROR only evaluate the branch they need, so IF(B1=0, 0, A1/B1) never divides by zero.
const LAZY = {
  IF(args, ev) {
    if (args.length < 2) return new Err('#VALUE!', 'IF() needs a condition and a value.');
    const c = toBool(ev(args[0]));
    if (isErr(c)) return c;
    return c ? ev(args[1]) : args.length > 2 ? ev(args[2]) : false;
  },
  IFERROR(args, ev) {
    if (args.length < 2) return new Err('#VALUE!', 'IFERROR() needs a value and a fallback.');
    const v = ev(args[0]);
    return isErr(scalar(v)) ? ev(args[1]) : v;
  },
};

// Signature + one-line description, used by autocomplete and the "build with Claude" prompt.
const HELP = {
  SUM: ['number1, [number2, …]', 'Adds numbers and ranges.'],
  AVERAGE: ['number1, [number2, …]', 'Arithmetic mean.'],
  MIN: ['number1, [number2, …]', 'Smallest number.'],
  MAX: ['number1, [number2, …]', 'Largest number.'],
  PRODUCT: ['number1, [number2, …]', 'Multiplies numbers.'],
  MEDIAN: ['number1, [number2, …]', 'Middle value.'],
  STDEV: ['number1, [number2, …]', 'Sample standard deviation.'],
  COUNT: ['value1, [value2, …]', 'Counts cells holding numbers.'],
  COUNTA: ['value1, [value2, …]', 'Counts non-empty cells.'],
  COUNTBLANK: ['range', 'Counts empty cells.'],
  SUMPRODUCT: ['range1, [range2, …]', 'Sum of element-wise products.'],
  SUMIF: ['range, criterion, [sum_range]', 'Sums cells that meet a criterion like ">100".'],
  SUMIFS: ['sum_range, range1, criterion1, …', 'Sums cells meeting every criterion.'],
  AVERAGEIF: ['range, criterion, [average_range]', 'Averages cells that meet a criterion.'],
  COUNTIF: ['range, criterion', 'Counts cells that meet a criterion.'],
  COUNTIFS: ['range1, criterion1, …', 'Counts rows meeting every criterion.'],
  ROUND: ['number, [digits]', 'Rounds half away from zero.'],
  ROUNDUP: ['number, [digits]', 'Rounds away from zero.'],
  ROUNDDOWN: ['number, [digits]', 'Rounds toward zero.'],
  INT: ['number', 'Rounds down to an integer.'],
  ABS: ['number', 'Absolute value.'],
  SIGN: ['number', '1, 0 or -1.'],
  SQRT: ['number', 'Square root.'],
  POWER: ['number, power', 'Raises to a power (same as ^).'],
  MOD: ['number, divisor', 'Remainder; takes the sign of the divisor.'],
  EXP: ['number', 'e raised to a power.'],
  LN: ['number', 'Natural logarithm.'],
  LOG10: ['number', 'Base-10 logarithm.'],
  PI: ['', '3.14159…'],
  RAND: ['', 'Random number in [0, 1); changes on every edit.'],
  RANDBETWEEN: ['low, high', 'Random integer between two numbers.'],
  IF: ['condition, value_if_true, [value_if_false]', 'Picks a value based on a condition.'],
  IFERROR: ['value, value_if_error', 'Replaces an error with a fallback.'],
  AND: ['logical1, [logical2, …]', 'TRUE if all are true.'],
  OR: ['logical1, [logical2, …]', 'TRUE if any is true.'],
  NOT: ['logical', 'Flips TRUE and FALSE.'],
  ISBLANK: ['value', 'TRUE if the cell is empty.'],
  ISNUMBER: ['value', 'TRUE if the value is a number.'],
  ISTEXT: ['value', 'TRUE if the value is text.'],
  ISERROR: ['value', 'TRUE if the value is an error.'],
  CONCAT: ['text1, [text2, …]', 'Joins text (same as &).'],
  CONCATENATE: ['text1, [text2, …]', 'Joins text (same as &).'],
  LEN: ['text', 'Number of characters.'],
  UPPER: ['text', 'UPPERCASE.'],
  LOWER: ['text', 'lowercase.'],
  PROPER: ['text', 'Capitalizes Each Word.'],
  TRIM: ['text', 'Removes extra spaces.'],
  LEFT: ['text, [count]', 'First characters.'],
  RIGHT: ['text, [count]', 'Last characters.'],
  MID: ['text, start, count', 'Characters from the middle.'],
  FIND: ['find_text, within_text, [start]', 'Position of text (case-sensitive).'],
  SUBSTITUTE: ['text, old, new', 'Replaces every occurrence of old with new.'],
  REPT: ['text, times', 'Repeats text.'],
  VALUE: ['text', 'Converts text to a number.'],
  VLOOKUP: ['value, table, column, [approximate]', 'Looks down the first column; use FALSE for exact match.'],
  XLOOKUP: ['value, lookup_range, return_range, [if_not_found]', 'Exact-match lookup returning the matching item.'],
  MATCH: ['value, range, [type]', 'Position of a value; type 0 = exact.'],
  INDEX: ['range, row, [column]', 'Value at a position in a range.'],
  AVG: ['number1, [number2, …]', 'Same as AVERAGE.'],
  TODAY: ['', 'Today’s date; updates every day.'],
  NOW: ['', 'The current date and time.'],
  DATE: ['year, month, day', 'Builds a date; months and days roll over.'],
  YEAR: ['date', 'Year of a date.'],
  MONTH: ['date', 'Month of a date, 1–12.'],
  DAY: ['date', 'Day of the month, 1–31.'],
  WEEKDAY: ['date, [type]', 'Day of the week; type 1 = Sun…Sat as 1–7, 2 = Mon…Sun as 1–7.'],
  EDATE: ['date, months', 'Same day some months later (or earlier).'],
  EOMONTH: ['date, months', 'Last day of the month, some months away.'],
  DAYS: ['end_date, start_date', 'Days between two dates.'],
  DATEDIF: ['start_date, end_date, unit', 'Whole "Y" years, "M" months or "D" days between dates.'],
  NETWORKDAYS: ['start_date, end_date', 'Working days (Mon–Fri) between dates, inclusive.'],
  TEXT: ['value, format', 'Formats a number or date as text, e.g. "#,##0.00", "0%", "dd mmm yyyy".'],
};

// ---------- sheet ----------
function compile(raw) {
  if (raw[0] === '=' && raw.length > 1) {
    try { return { raw, node: parse(raw.slice(1)) }; }
    catch (e) { return { raw, node: { t: 'err', v: '#ERROR!', msg: e.message } }; }
  }
  return { raw, lit: literal(raw) };
}
class Sheet {
  // A Sheet made on its own gets a private one-sheet Workbook, so `new Sheet()` just works.
  constructor(book, id = 's1', name = 'Sheet1') {
    this.book = book || new Workbook();
    this.id = id; this.name = name;
    this.cells = new Map(); this.vals = new Map();
    if (!book) this.book.sheets.push(this);
  }
  set(k, raw) {
    raw = raw == null ? '' : String(raw);
    if (raw === '') this.cells.delete(k); else this.cells.set(k, compile(raw));
  }
  raw(k) { const c = this.cells.get(k); return c ? c.raw : ''; }
  recalc() { this.book.recalc(); }
  value(k) { return this.book.value(this, k); }
}
class Workbook {
  constructor() { this.sheets = []; this.busy = new Set(); }
  add(name, id = 's' + Math.random().toString(36).slice(2, 9)) { const sh = new Sheet(this, id, name); this.sheets.push(sh); return sh; }
  remove(id) { this.sheets = this.sheets.filter(s => s.id !== id); }
  get(id) { return this.sheets.find(s => s.id === id); }
  byName(name) { const n = name.toLowerCase(); return this.sheets.find(s => s.name.toLowerCase() === n); }
  // ponytail: full recalc of every sheet on each edit (~ms for thousands of cells); dependency graph if books get huge.
  recalc() {
    for (const s of this.sheets) s.vals.clear();
    for (const s of this.sheets) {
      // Row-major order keeps long fill-down chains (A2=A1+1 …) shallow on the call stack.
      const keys = [...s.cells.keys()].map(k => [k, parseKey(k)]).sort((a, b) => a[1].r - b[1].r || a[1].c - b[1].c);
      for (const [k] of keys) this.value(s, k);
    }
  }
  value(sh, k) {
    const cached = sh.vals.get(k);
    if (cached !== undefined) return cached;
    const cell = sh.cells.get(k);
    if (!cell) return null;
    const id = sh.id + '!' + k;
    let v;
    if (!cell.node) v = cell.lit;
    else if (this.busy.has(id)) return new Err('#CYCLE!');
    else {
      this.busy.add(id);
      try { v = finalValue(evaluate(cell.node, sh)); }
      catch (e) { v = new Err('#ERROR!', e instanceof RangeError ? 'Formula chain is too deep.' : e.message); }
      finally { this.busy.delete(id); }
    }
    sh.vals.set(k, v);
    return v;
  }
}

// Move relative references by (dr, dc); $-anchored parts stay put. Used by copy/paste and fill.
function shift(raw, dr, dc) {
  if (raw[0] !== '=') return raw;
  let toks;
  try { toks = tokenize(raw.slice(1)); } catch { return raw; }
  let out = raw.slice(1);
  for (let i = toks.length - 1; i >= 0; i--) {
    const t = toks[i];
    if (t.k !== 'ref') continue;
    const m = /^(\$?)([A-Za-z]+)(\$?)(\d+)$/.exec(t.cell);
    const c = colIndex(m[2]) + (m[1] ? 0 : dc), r = +m[4] - 1 + (m[3] ? 0 : dr);
    out = out.slice(0, t.s) + (r < 0 || c < 0 ? '#REF!' : t.pre + m[1] + colName(c) + m[3] + (r + 1)) + out.slice(t.e);
  }
  return '=' + out;
}

// What Excel does on Enter: uppercase names and references, close any parentheses left open.
function tidy(raw) {
  if (raw[0] !== '=') return raw;
  let toks;
  try { toks = tokenize(raw.slice(1)); } catch { return raw; }
  let out = raw.slice(1), depth = 0;
  for (let i = toks.length - 1; i >= 0; i--) {
    const t = toks[i];
    if (t.k === 'ref') out = out.slice(0, t.s) + t.pre + t.cell.toUpperCase() + out.slice(t.e);
    if (t.k === 'id') out = out.slice(0, t.s) + t.v.toUpperCase() + out.slice(t.e);
    if (t.k === 'op') depth += t.v === '(' ? 1 : t.v === ')' ? -1 : 0;
  }
  return '=' + out + ')'.repeat(Math.max(0, depth));
}

// Point formulas at a renamed sheet: Old!A1 → New!A1 ('Quoted Names' handled both ways).
function renameRefs(raw, oldName, newName) {
  if (raw[0] !== '=') return raw;
  let toks;
  try { toks = tokenize(raw.slice(1)); } catch { return raw; }
  let out = raw.slice(1);
  const o = oldName.toLowerCase();
  for (let i = toks.length - 1; i >= 0; i--) {
    const t = toks[i];
    if (t.k === 'ref' && t.sheet && t.sheet.toLowerCase() === o) out = out.slice(0, t.s) + quoteSheet(newName) + '!' + t.cell + out.slice(t.e);
  }
  return '=' + out;
}

// Cell and range references inside a formula, with their text positions (for colored highlights).
function refSpans(f) {
  if (f[0] !== '=') return [];
  let toks;
  try { toks = tokenize(f.slice(1)); } catch { return []; }
  const out = [];
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i], a = t.k === 'ref' && parseKey(t.cell);
    if (!a) continue;
    let b = a, e = t.e;
    const u = toks[i + 2], bb = toks[i + 1]?.v === ':' && u?.k === 'ref' && parseKey(u.cell);
    if (bb) { b = bb; e = u.e; i += 2; }
    out.push({ s: t.s + 1, e: e + 1, r0: Math.min(a.r, b.r), c0: Math.min(a.c, b.c), r1: Math.max(a.r, b.r), c1: Math.max(a.c, b.c), sheet: t.sheet || null });
  }
  return out;
}

const api = { Sheet, Workbook, Err, RangeVal, parse, tokenize, shift, tidy, refSpans, renameRefs, quoteSheet, literal, parseDate, isoDate, fmtGeneral, colName, colIndex, key, parseKey, config, FUNCS, HELP };
if (typeof module === 'object' && module.exports) module.exports = api; else root.Engine = api;
})(this);
