/**
 * formula.js — tokenizer, parser and evaluator for workbook formulas.
 *
 * Language (Excel-like, deliberately small):
 *   numbers      12  0.35  .35  1e3
 *   strings      "text"
 *   booleans     TRUE FALSE
 *   cell refs    C3  $C$3  C$3      (columns A–Z, AA… ; rows from 1)
 *   ranges       C1:C20
 *   other sheet  [Ground slab]!H5   [Ground slab]!C1:C9
 *   names        wastage  gst       (workbook-level named values)
 *   operators    + - * / ^ %  &(join text)  = <> < > <= >=   (unary -, postfix %)
 *   functions    SUM AVERAGE MIN MAX COUNT ROUND ROUNDUP ROUNDDOWN CEILING FLOOR
 *                ABS SQRT PI IF AND OR NOT IFERROR MEAS("measurement name")
 *   link         inside a measurement-link modifier, `link` is the measured value
 *                (e.g.  =link*1.05  for 5% waste)
 *
 * Errors are values (they propagate through arithmetic, like Excel):
 *   #REF!  deleted/invalid reference      #DIV/0!  divide by zero
 *   #NAME? unknown name/function          #VALUE!  wrong type
 *   #CIRC! circular reference             #ERROR!  syntax error
 *
 * The parser keeps every token's source span, so formulas can be REWRITTEN
 * (row insert/delete, copy-paste offsets) without re-printing and losing the
 * user's formatting.
 */

export class Err {
  constructor(code) {
    this.code = code;
  }
  toString() {
    return this.code;
  }
}
export const ERR = {
  REF: new Err('#REF!'),
  DIV0: new Err('#DIV/0!'),
  NAME: new Err('#NAME?'),
  VALUE: new Err('#VALUE!'),
  CIRC: new Err('#CIRC!'),
  SYNTAX: new Err('#ERROR!'),
  NA: new Err('#N/A'),
};
export const isErr = (v) => v instanceof Err;

// ------------------------------------------------------------------ A1 helpers

export function colToIndex(col) {
  let n = 0;
  for (let i = 0; i < col.length; i++) n = n * 26 + (col.charCodeAt(i) - 64);
  return n - 1; // A → 0
}
export function indexToCol(i) {
  let s = '';
  i += 1;
  while (i > 0) {
    const m = (i - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    i = Math.floor((i - 1) / 26);
  }
  return s;
}
/** "C12" → {col:'C', row:12} */
export function parseA1(a1) {
  const m = /^\$?([A-Z]+)\$?(\d+)$/.exec(a1);
  return m ? { col: m[1], row: +m[2] } : null;
}
export const a1 = (col, row) => `${col}${row}`;

// ------------------------------------------------------------------ tokenizer

const RE_REF = /\$?[A-Za-z]{1,3}\$?\d+/y; // case-insensitive: users type c3 as often as C3
const RE_NUM = /(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/y;
const RE_IDENT = /[A-Za-z_][A-Za-z0-9_.]*/y;

/**
 * @returns {Array<{t:string, v:any, s:number, e:number}>}
 *   t: num | str | ref | sheet | ident | op | ( | ) | , | : | end
 */
export function tokenize(src) {
  const out = [];
  let i = 0;
  const n = src.length;
  while (i < n) {
    const c = src[i];
    if (c === ' ' || c === '\t' || c === '\n') {
      i++;
      continue;
    }
    const s = i;
    if (c === '"') {
      let j = i + 1,
        v = '';
      while (j < n) {
        if (src[j] === '"') {
          if (src[j + 1] === '"') {
            v += '"';
            j += 2;
            continue;
          }
          break;
        }
        v += src[j++];
      }
      if (j >= n) throw new SyntaxError('unterminated string');
      out.push({ t: 'str', v, s, e: j + 1 });
      i = j + 1;
      continue;
    }
    if (src.startsWith('#REF!', i)) {
      // left behind when a referenced row is deleted
      out.push({ t: 'err', v: '#REF!', s, e: i + 5 });
      i += 5;
      continue;
    }
    if (c === '[') {
      // [Sheet name]!
      const j = src.indexOf(']', i);
      if (j < 0 || src[j + 1] !== '!') throw new SyntaxError('bad sheet reference');
      out.push({ t: 'sheet', v: src.slice(i + 1, j).trim(), s, e: j + 2 });
      i = j + 2;
      continue;
    }
    RE_REF.lastIndex = i;
    let m = RE_REF.exec(src);
    if (m && !/[A-Za-z0-9_(]/.test(src[i + m[0].length] ?? '')) {
      const t = m[0];
      const mm = /^(\$?)([A-Za-z]+)(\$?)(\d+)$/.exec(t);
      out.push({ t: 'ref', v: { col: mm[2].toUpperCase(), row: +mm[4], absC: !!mm[1], absR: !!mm[3] }, s, e: i + t.length });
      i += t.length;
      continue;
    }
    RE_NUM.lastIndex = i;
    m = RE_NUM.exec(src);
    if (m && /[\d.]/.test(c)) {
      out.push({ t: 'num', v: parseFloat(m[0]), s, e: i + m[0].length });
      i += m[0].length;
      continue;
    }
    RE_IDENT.lastIndex = i;
    m = RE_IDENT.exec(src);
    if (m) {
      out.push({ t: 'ident', v: m[0], s, e: i + m[0].length });
      i += m[0].length;
      continue;
    }
    const two = src.slice(i, i + 2);
    if (two === '<>' || two === '<=' || two === '>=') {
      out.push({ t: 'op', v: two, s, e: i + 2 });
      i += 2;
      continue;
    }
    if ('+-*/^%&=<>'.includes(c)) out.push({ t: 'op', v: c, s, e: i + 1 });
    else if ('(),:'.includes(c)) out.push({ t: c, v: c, s, e: i + 1 });
    else throw new SyntaxError(`unexpected '${c}'`);
    i++;
  }
  out.push({ t: 'end', s: n, e: n });
  return out;
}

// ------------------------------------------------------------------ parser (Pratt)

const BIN = {
  '=': [1, 'L'],
  '<>': [1, 'L'],
  '<': [1, 'L'],
  '>': [1, 'L'],
  '<=': [1, 'L'],
  '>=': [1, 'L'],
  '&': [2, 'L'],
  '+': [3, 'L'],
  '-': [3, 'L'],
  '*': [4, 'L'],
  '/': [4, 'L'],
  '^': [6, 'R'], // above unary minus: -2^2 = -4 (maths), unlike Excel's 4
};

/**
 * Parse a formula body (without the leading '='). Returns an AST, or throws SyntaxError.
 * Node kinds: num str bool ref range name call neg pct bin
 */
export function parse(src) {
  const toks = tokenize(src);
  let p = 0;
  const peek = () => toks[p];
  const next = () => toks[p++];
  const expect = (t) => {
    if (toks[p].t !== t) throw new SyntaxError(`expected '${t}'`);
    return toks[p++];
  };

  const refNode = (sheet) => {
    const a = expect('ref');
    if (peek().t === ':') {
      next();
      const b = expect('ref');
      return { k: 'range', sheet, a: a.v, b: b.v };
    }
    return { k: 'ref', sheet, ...a.v };
  };

  const primary = () => {
    const tk = next();
    switch (tk.t) {
      case 'num':
        return { k: 'num', v: tk.v };
      case 'str':
        return { k: 'str', v: tk.v };
      case 'ref':
        p--;
        return refNode(null);
      case 'sheet':
        if (peek().t === 'err') return next(), { k: 'err', v: ERR.REF };
        return refNode(tk.v);
      case 'err':
        return { k: 'err', v: ERR.REF };
      case '(': {
        const e = expr(0);
        expect(')');
        return e;
      }
      case 'op':
        if (tk.v === '-') return { k: 'neg', a: expr(5) };
        if (tk.v === '+') return expr(5);
        break;
      case 'ident': {
        const up = tk.v.toUpperCase();
        if (peek().t === '(') {
          next();
          const args = [];
          if (peek().t !== ')') {
            for (;;) {
              args.push(expr(0));
              if (peek().t === ',') {
                next();
                continue;
              }
              break;
            }
          }
          expect(')');
          return { k: 'call', fn: up, args };
        }
        if (up === 'TRUE' || up === 'FALSE') return { k: 'bool', v: up === 'TRUE' };
        return { k: 'name', v: tk.v };
      }
    }
    throw new SyntaxError(`unexpected ${tk.t === 'end' ? 'end of formula' : `'${tk.v}'`}`);
  };

  const expr = (minPrec) => {
    let left = primary();
    for (;;) {
      const tk = peek();
      if (tk.t === 'op' && tk.v === '%') {
        next();
        left = { k: 'pct', a: left };
        continue;
      }
      if (tk.t !== 'op' || !BIN[tk.v]) break;
      const [prec, assoc] = BIN[tk.v];
      if (prec < minPrec) break;
      next();
      const right = expr(assoc === 'L' ? prec + 1 : prec);
      left = { k: 'bin', op: tk.v, a: left, b: right };
    }
    return left;
  };

  const ast = expr(0);
  if (peek().t !== 'end') throw new SyntaxError(`unexpected '${peek().v}'`);
  return ast;
}

const parseCache = new Map();
/** Parse with a small cache (formulas are re-evaluated far more often than edited). */
export function parseCached(src) {
  let r = parseCache.get(src);
  if (r === undefined) {
    try {
      r = parse(src);
    } catch {
      r = null;
    }
    if (parseCache.size > 5000) parseCache.clear();
    parseCache.set(src, r);
  }
  return r;
}

// ------------------------------------------------------------------ evaluator

/**
 * ctx = {
 *   cell(sheetName|null, col, row) → value     (null sheet = current sheet)
 *   name(id) → value | undefined
 *   meas(key) → value | undefined
 *   link → value | undefined                   (inside link modifiers)
 * }
 * Values: number | string | boolean | null (empty) | Err
 */
export function evaluate(ast, ctx) {
  if (!ast) return ERR.SYNTAX;
  return ev(ast, ctx);
}

const toNum = (v) => {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (isErr(v)) return v;
  const n = Number(String(v).replace(/,/g, ''));
  return Number.isFinite(n) ? n : ERR.VALUE;
};
const toBool = (v) => {
  if (isErr(v)) return v;
  if (typeof v === 'string') {
    const u = v.toUpperCase();
    if (u === 'TRUE') return true;
    if (u === 'FALSE') return false;
    return v !== '';
  }
  return !!toNum(v);
};
const num = (x) => (Number.isFinite(x) ? x : ERR.VALUE);

function rangeCells(node) {
  const c0 = colToIndex(node.a.col),
    c1 = colToIndex(node.b.col);
  const r0 = Math.min(node.a.row, node.b.row),
    r1 = Math.max(node.a.row, node.b.row);
  const out = [];
  for (let r = r0; r <= r1; r++) for (let c = Math.min(c0, c1); c <= Math.max(c0, c1); c++) out.push([indexToCol(c), r]);
  return out;
}

/** Flatten args for aggregate functions (ranges expand; text/empty in ranges are skipped, like Excel). */
function flat(args, ctx) {
  const out = [];
  for (const a of args) {
    if (a.k === 'range') {
      for (const [c, r] of rangeCells(a)) {
        const v = ctx.cell(a.sheet, c, r);
        if (isErr(v)) return v;
        if (typeof v === 'number') out.push(v);
        else if (typeof v === 'boolean') out.push(v ? 1 : 0);
      }
    } else {
      const v = ev(a, ctx);
      if (isErr(v)) return v;
      if (v === null || v === '') continue;
      const n = toNum(v);
      if (isErr(n)) return n;
      out.push(n);
    }
  }
  return out;
}

const roundTo = (x, d, mode) => {
  const f = Math.pow(10, d);
  const y = x * f;
  // nudge by a tiny epsilon so 2.675 → 2.68 (binary float would give 2.67)
  const eps = Math.sign(y) * 1e-9;
  const r = mode === 'up' ? Math.sign(y) * Math.ceil(Math.abs(y) - 1e-9) : mode === 'down' ? Math.trunc(y + eps) : Math.round(Math.abs(y) + 1e-9) * Math.sign(y);
  return r / f;
};

const FN = {
  SUM: (a, c) => {
    const v = flat(a, c);
    return isErr(v) ? v : v.reduce((s, x) => s + x, 0);
  },
  AVERAGE: (a, c) => {
    const v = flat(a, c);
    return isErr(v) ? v : v.length ? v.reduce((s, x) => s + x, 0) / v.length : ERR.DIV0;
  },
  MIN: (a, c) => {
    const v = flat(a, c);
    return isErr(v) ? v : v.length ? Math.min(...v) : 0;
  },
  MAX: (a, c) => {
    const v = flat(a, c);
    return isErr(v) ? v : v.length ? Math.max(...v) : 0;
  },
  COUNT: (a, c) => {
    const v = flat(a, c);
    return isErr(v) ? v : v.length;
  },
  ROUND: (a, c) => round2(a, c, 'half'),
  ROUNDUP: (a, c) => round2(a, c, 'up'),
  ROUNDDOWN: (a, c) => round2(a, c, 'down'),
  CEILING: (a, c) => sig(a, c, (q) => Math.ceil(q - 1e-9)),
  FLOOR: (a, c) => sig(a, c, (q) => Math.floor(q + 1e-9)),
  ABS: (a, c) => un(a, c, Math.abs),
  SQRT: (a, c) => un(a, c, (x) => (x < 0 ? NaN : Math.sqrt(x))),
  PI: () => Math.PI,
  IF: (a, c) => {
    if (a.length < 2 || a.length > 3) return ERR.VALUE;
    const t = toBool(ev(a[0], c));
    if (isErr(t)) return t;
    return t ? ev(a[1], c) : a.length === 3 ? ev(a[2], c) : false;
  },
  IFERROR: (a, c) => {
    if (a.length !== 2) return ERR.VALUE;
    const v = ev(a[0], c);
    return isErr(v) ? ev(a[1], c) : v;
  },
  AND: (a, c) => logic(a, c, true),
  OR: (a, c) => logic(a, c, false),
  NOT: (a, c) => {
    const b = toBool(ev(a[0], c));
    return isErr(b) ? b : !b;
  },
  MEAS: (a, c) => {
    if (a.length !== 1) return ERR.VALUE;
    const key = ev(a[0], c);
    if (isErr(key)) return key;
    const v = c.meas?.(String(key));
    return v === undefined ? ERR.REF : v;
  },
};
function round2(a, c, mode) {
  const x = toNum(ev(a[0], c));
  const d = a[1] ? toNum(ev(a[1], c)) : 0;
  if (isErr(x)) return x;
  if (isErr(d)) return d;
  return roundTo(x, Math.trunc(d), mode);
}
function sig(a, c, f) {
  const x = toNum(ev(a[0], c));
  const s = a[1] ? toNum(ev(a[1], c)) : 1;
  if (isErr(x)) return x;
  if (isErr(s)) return s;
  if (s === 0) return 0;
  return +(f(x / s) * s).toPrecision(15); // epsilon inside f: 2.0/0.5 must stay 4; toPrecision: 12×0.6 → 7.2 not 7.1999…
}
function un(a, c, f) {
  const x = toNum(ev(a[0], c));
  return isErr(x) ? x : num(f(x));
}
function logic(a, c, isAnd) {
  for (const x of a) {
    const b = toBool(ev(x, c));
    if (isErr(b)) return b;
    if (isAnd && !b) return false;
    if (!isAnd && b) return true;
  }
  return isAnd;
}

function cmp(op, x, y) {
  // compare numbers numerically, text case-insensitively
  if (typeof x === 'string' || typeof y === 'string') {
    x = String(x ?? '').toLowerCase();
    y = String(y ?? '').toLowerCase();
  } else {
    x = toNum(x);
    y = toNum(y);
  }
  switch (op) {
    case '=':
      return x === y;
    case '<>':
      return x !== y;
    case '<':
      return x < y;
    case '>':
      return x > y;
    case '<=':
      return x <= y;
    case '>=':
      return x >= y;
  }
}

function ev(n, ctx) {
  switch (n.k) {
    case 'num':
    case 'str':
    case 'bool':
    case 'err':
      return n.v;
    case 'ref': {
      const v = ctx.cell(n.sheet, n.col, n.row);
      return v;
    }
    case 'range':
      return ERR.VALUE; // a bare range outside an aggregate function
    case 'name': {
      if (n.v.toLowerCase() === 'link' && ctx.link !== undefined) return ctx.link;
      const v = ctx.name?.(n.v);
      return v === undefined ? ERR.NAME : v;
    }
    case 'call': {
      const f = FN[n.fn];
      return f ? f(n.args, ctx) : ERR.NAME;
    }
    case 'neg': {
      const x = toNum(ev(n.a, ctx));
      return isErr(x) ? x : -x;
    }
    case 'pct': {
      const x = toNum(ev(n.a, ctx));
      return isErr(x) ? x : x / 100;
    }
    case 'bin': {
      const a = ev(n.a, ctx);
      if (isErr(a)) return a;
      const b = ev(n.b, ctx);
      if (isErr(b)) return b;
      if (n.op === '&') return `${a ?? ''}${b ?? ''}`;
      if (BIN[n.op][0] === 1) return cmp(n.op, a, b);
      const x = toNum(a),
        y = toNum(b);
      if (isErr(x)) return x;
      if (isErr(y)) return y;
      switch (n.op) {
        case '+':
          return x + y;
        case '-':
          return x - y;
        case '*':
          return x * y;
        case '/':
          return y === 0 ? ERR.DIV0 : x / y;
        case '^':
          return num(Math.pow(x, y));
      }
    }
  }
  return ERR.SYNTAX;
}

// ------------------------------------------------------------------ rewriting

/**
 * Rewrite every cell reference in a formula body. `fn(ref, sheet)` receives
 * {col,row,absC,absR} and the sheet prefix (null = own sheet) and returns a
 * new ref object, or null for #REF!. Ranges are passed as `fn(range)` with
 * {a,b} via `rangeFn` when supplied. Untouched text is kept verbatim.
 */
export function rewriteRefs(body, refFn, rangeFn) {
  let toks;
  try {
    toks = tokenize(body);
  } catch {
    return body;
  }
  let out = '';
  let last = 0;
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    if (t.t !== 'ref') continue;
    const sheet = i > 0 && toks[i - 1].t === 'sheet' && toks[i - 1].e === t.s ? toks[i - 1].v : null;
    // range?
    if (toks[i + 1]?.t === ':' && toks[i + 2]?.t === 'ref') {
      const b = toks[i + 2];
      let r = rangeFn ? rangeFn({ a: t.v, b: b.v }, sheet) : { a: refFn(t.v, sheet), b: refFn(b.v, sheet) };
      out += body.slice(last, t.s);
      out += r && r.a && r.b ? `${fmtRef(r.a)}:${fmtRef(r.b)}` : '#REF!';
      last = b.e;
      i += 2;
      continue;
    }
    const r = refFn(t.v, sheet);
    out += body.slice(last, t.s) + (r ? fmtRef(r) : '#REF!');
    last = t.e;
  }
  return out + body.slice(last);
}
const fmtRef = (r) => `${r.absC ? '$' : ''}${r.col}${r.absR ? '$' : ''}${r.row}`;

/** Rename a sheet everywhere it's referenced: [Old]! → [New]! */
export function renameSheetRefs(body, oldName, newName) {
  let toks;
  try {
    toks = tokenize(body);
  } catch {
    return body;
  }
  let out = '',
    last = 0;
  for (const t of toks) {
    if (t.t === 'sheet' && t.v.toLowerCase() === oldName.toLowerCase()) {
      out += body.slice(last, t.s) + `[${newName}]!`;
      last = t.e;
    }
  }
  return out + body.slice(last);
}

/** Copy-paste offset: relative refs move by (dCol, dRow); $-absolute parts don't. */
export function offsetFormula(body, dCol, dRow) {
  return rewriteRefs(body, (r) => {
    const c = r.absC ? colToIndex(r.col) : colToIndex(r.col) + dCol;
    const row = r.absR ? r.row : r.row + dRow;
    if (c < 0 || row < 1) return null;
    return { ...r, col: indexToCol(c), row };
  });
}
