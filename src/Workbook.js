/**
 * Workbook.js — the estimating model + calculation engine.
 *
 * A Workbook is a TREE of sheets: the estimate at the root, and every
 * drill-down a child sheet that belongs to exactly one parent cell.
 *
 * WHY THIS NEVER GOES STALE (the key difference from Nezo)
 * --------------------------------------------------------
 * Nezo stores each drill-down's total in the parent cell (`rolled`) and asks
 * the user to press Recalculate when they disagree. Here nothing is stored:
 * every value is computed on demand and memoised, and the engine records
 * which values each computation READ (cells, sheet totals, measurements,
 * names). When anything changes, everything that (transitively) read it is
 * invalidated — across sheets, through drill-downs, up to the estimate
 * total. The next read recomputes just those. So:
 *
 *   measurement changes  →  linked cell  →  row total  →  sheet total
 *                        →  parent drilled cell  →  …  →  estimate total
 *
 * happens automatically, in microseconds, with no Recalculate button.
 * Dependencies are discovered while evaluating (like a reactive system), so
 * IF(), ranges and cross-sheet refs are handled without static analysis.
 *
 * Every edit goes through a transaction → one undo step.
 */

import { parseCached, evaluate, rewriteRefs, renameSheetRefs, offsetFormula, ERR, Err, isErr, parseA1, colToIndex, indexToCol } from './formula.js';
import { columnsFor, KIND, drillTitle } from './columns.js';

const uid = () => (globalThis.crypto?.randomUUID?.() ?? Math.random().toString(36).slice(2) + Date.now().toString(36));
const NUM_RE = /^\s*([-+]?)\$?\s*((?:\d{1,3}(?:,\d{3})+|\d*)(?:\.\d+)?(?:e[-+]?\d+)?)\s*(%?)\s*$/i;

/** "1,234.5" → 1234.5 · "$12" → 12 · "15%" → 0.15 · "abc" → null */
export function parseLiteral(raw) {
  const m = NUM_RE.exec(raw);
  if (!m || m[2] === '' || m[2] === '.') return null;
  let v = parseFloat(m[2].replace(/,/g, ''));
  if (!Number.isFinite(v)) return null;
  if (m[1] === '-') v = -v;
  if (m[3]) v /= 100;
  return v;
}

const isFormula = (raw) => typeof raw === 'string' && raw.trimStart().startsWith('=');
const body = (raw) => raw.trimStart().slice(1);

function cloneSheet(s) {
  return {
    id: s.id,
    name: s.name,
    kind: s.kind,
    parent: s.parent ? { ...s.parent } : null,
    rowCount: s.rowCount,
    cells: new Map([...s.cells].map(([k, c]) => [k, structuredClone(c)])),
    rowTypes: new Map(s.rowTypes),
  };
}
const sheetToJSON = (s) => ({
  id: s.id,
  name: s.name,
  kind: s.kind,
  parent: s.parent,
  rowCount: s.rowCount,
  cells: Object.fromEntries(s.cells),
  rowTypes: Object.fromEntries(s.rowTypes),
});
const sheetFromJSON = (j) => ({
  id: j.id,
  name: j.name,
  kind: j.kind,
  parent: j.parent ?? null,
  rowCount: j.rowCount ?? 0,
  cells: new Map(Object.entries(j.cells ?? {})),
  rowTypes: new Map(Object.entries(j.rowTypes ?? {}).map(([r, t]) => [+r, t])),
});

export class Workbook extends EventTarget {
  /**
   * @param {object} [o]
   * @param {string} [o.name='Estimate']   root sheet name
   * @param {number} [o.userColumns=6]     free columns after H (I, J, …)
   * @param {string[]} [o.userColumnLabels]
   * @param {number} [o.rows=30]           initial rows on the root sheet
   */
  constructor(o = {}) {
    super();
    this.userColumns = o.userColumns ?? 6;
    this.userColumnLabels = o.userColumnLabels ?? [];
    /** @type {Map<string, any>} */
    this.sheets = new Map();
    /** named values: lower-case name → { name, raw } */
    this.names = new Map();
    /** measurements (owned by the app): id → { value, unit, name } */
    this.measurements = new Map();

    this._cache = new Map();
    this._deps = new Map(); // key → Set(keys it read)
    this._rdeps = new Map(); // key → Set(keys that read it)
    this._stack = [];
    this._evaluating = new Set();
    this._undo = [];
    this._redo = [];
    this._tx = null;
    this._colsCache = new Map();

    if (!o._empty) {
      const root = { id: uid(), name: o.name ?? 'Estimate', kind: KIND.COST, parent: null, rowCount: o.rows ?? 30, cells: new Map(), rowTypes: new Map() };
      this.sheets.set(root.id, root);
      this.rootId = root.id;
    }
  }

  // ================================================================ structure / lookup

  get root() {
    return this.sheets.get(this.rootId);
  }
  sheet(id) {
    return this.sheets.get(id);
  }
  sheetByName(name) {
    const n = name.trim().toLowerCase();
    for (const s of this.sheets.values()) if (s.name.toLowerCase() === n) return s;
    return null;
  }
  /** Column definitions (labels, roles) for a sheet. */
  columns(sheetId) {
    const kind = this.sheets.get(sheetId)?.kind ?? KIND.COST;
    let c = this._colsCache.get(kind);
    if (!c) {
      c = columnsFor(kind, this.userColumns, this.userColumnLabels);
      this._colsCache.set(kind, c);
    }
    return c;
  }
  columnLetters(sheetId) {
    return Object.keys(this.columns(sheetId));
  }
  /** Sheets from the root down to `sheetId`. */
  path(sheetId) {
    const out = [];
    for (let s = this.sheets.get(sheetId); s; s = s.parent ? this.sheets.get(s.parent.sheetId) : null) out.unshift(s);
    return out;
  }
  /** Direct child sheets of a sheet, in row order. */
  children(sheetId) {
    const out = [];
    for (const s of this.sheets.values()) if (s.parent?.sheetId === sheetId) out.push(s);
    return out.sort((a, b) => parseA1(a.parent.a1).row - parseA1(b.parent.a1).row || a.parent.a1.localeCompare(b.parent.a1));
  }
  rowType(sheetId, row) {
    return this.sheets.get(sheetId)?.rowTypes.get(row) ?? 'normal';
  }

  /** The stored cell object (read-only copy) or null. */
  cell(sheetId, a1) {
    const c = this.sheets.get(sheetId)?.cells.get(a1);
    return c ? structuredClone(c) : null;
  }

  // ================================================================ calculation engine

  /** Live value of a cell: number | string | boolean | null (empty) | Err */
  value(sheetId, a1) {
    return this._read(`c|${sheetId}|${a1}`, () => this._computeCell(sheetId, a1));
  }

  /** Live total of a sheet (sum of the Total/Quantity column, excluding subtotal rows). */
  total(sheetId) {
    return this._read(`t|${sheetId}`, () => this._computeTotal(sheetId));
  }

  /** What's in a cell, for UI hints. */
  info(sheetId, a1) {
    const s = this.sheets.get(sheetId);
    const c = s?.cells.get(a1);
    const p = parseA1(a1);
    const def = this.columns(sheetId)[p.col];
    const rt = s?.rowTypes.get(p.row);
    if (rt === 'subtotal' && p.col === 'H') return { type: 'subtotal' };
    if (c?.drill) return { type: 'drill', sheetId: c.drill, kind: this.sheets.get(c.drill)?.kind };
    if (c?.link) {
      const m = this.measurements.get(c.link.id);
      return { type: 'link', measurementId: c.link.id, expr: c.link.expr ?? null, name: m?.name ?? null, missing: !m };
    }
    if (def?.computed && !c?.raw) return { type: 'computed' };
    if (def?.computed && c?.raw) return { type: 'override', raw: c.raw };
    if (isFormula(c?.raw)) return { type: 'formula', raw: c.raw };
    if (c?.raw) return { type: 'value', raw: c.raw };
    return { type: 'empty' };
  }

  /** Value formatted for display in the grid. */
  display(sheetId, a1) {
    const v = this.value(sheetId, a1);
    return this.format(sheetId, a1, v);
  }
  format(sheetId, a1, v) {
    if (v === null || v === undefined || v === '') return '';
    if (isErr(v)) return v.code;
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (typeof v === 'string') return v;
    const def = this.columns(sheetId)[parseA1(a1).col];
    const dp = def?.dp;
    if (dp === undefined) return fmtNum(v, 0, 4);
    return fmtNum(v, dp, dp);
  }

  /**
   * Memoised read with dependency recording.
   * `key` kinds: c|sheet|A1  t|sheet  n|name  m|id  rows|sheet
   */
  _read(key, compute) {
    const top = this._stack[this._stack.length - 1];
    if (top) this._link(top, key);
    if (this._cache.has(key)) return this._cache.get(key);
    if (this._evaluating.has(key)) return ERR.CIRC;
    this._clearDeps(key);
    this._evaluating.add(key);
    this._stack.push(key);
    let v;
    try {
      v = compute();
    } finally {
      this._stack.pop();
      this._evaluating.delete(key);
    }
    this._cache.set(key, v);
    return v;
  }
  /** Record "currently-evaluating value depends on `key`" for a non-cached source (measurements, row structure). */
  _track(key) {
    const top = this._stack[this._stack.length - 1];
    if (top) this._link(top, key);
  }
  _link(from, to) {
    let d = this._deps.get(from);
    if (!d) this._deps.set(from, (d = new Set()));
    d.add(to);
    let r = this._rdeps.get(to);
    if (!r) this._rdeps.set(to, (r = new Set()));
    r.add(from);
  }
  _clearDeps(key) {
    const d = this._deps.get(key);
    if (!d) return;
    for (const to of d) this._rdeps.get(to)?.delete(key);
    this._deps.delete(key);
  }
  /** Invalidate `keys` and everything that depends on them. Returns the set invalidated. */
  _invalidate(keys) {
    const seen = new Set();
    const queue = [...keys];
    while (queue.length) {
      const k = queue.pop();
      if (seen.has(k)) continue;
      seen.add(k);
      this._cache.delete(k);
      const r = this._rdeps.get(k);
      if (r) for (const x of r) queue.push(x);
    }
    return seen;
  }
  _invalidateAll() {
    this._cache.clear();
    this._deps.clear();
    this._rdeps.clear();
  }

  _computeCell(sheetId, a1) {
    const s = this.sheets.get(sheetId);
    if (!s) return ERR.REF;
    const p = parseA1(a1);
    const def = this.columns(sheetId)[p.col];
    if (!def) return ERR.REF;
    const c = s.cells.get(a1);
    const rt = s.rowTypes.get(p.row);

    if (rt === 'subtotal') {
      if (p.col !== 'H') return c?.raw && !def.computed ? this._literal(sheetId, c.raw) : null;
      this._track(`rows|${sheetId}`);
      let sum = 0;
      for (let r = p.row - 1; r >= 1 && s.rowTypes.get(r) !== 'subtotal'; r--) {
        if (s.rowTypes.get(r) === 'heading') continue;
        const v = this.value(sheetId, `H${r}`);
        if (isErr(v)) return v;
        if (typeof v === 'number') sum += v;
      }
      return sum;
    }
    if (c?.drill) {
      if (!this.sheets.has(c.drill)) return ERR.REF;
      return this.total(c.drill);
    }
    if (def.computed && !c?.raw) return this._computeRow(s, p.col, p.row);
    if (c?.link) {
      this._track(`m|${c.link.id}`);
      const m = this.measurements.get(c.link.id);
      const lv = m ? m.value : ERR.REF;
      if (!c.link.expr) return lv;
      if (isErr(lv)) return lv;
      return evaluate(parseCached(body(c.link.expr)), this._ctx(sheetId, lv));
    }
    if (c?.raw === undefined || c.raw === '') return null;
    // text columns (Code, Description, Unit) keep what was typed: "01" stays "01"
    if (def.type === 'text' && !isFormula(c.raw)) return c.raw;
    return this._literal(sheetId, c.raw);
  }

  _literal(sheetId, raw) {
    if (isFormula(raw)) return evaluate(parseCached(body(raw)), this._ctx(sheetId));
    const n = parseLiteral(raw);
    return n === null ? raw : n;
  }

  /** The built-in row maths (see columns.js). */
  _computeRow(s, col, row) {
    const v = (c) => this.value(s.id, `${c}${row}`);
    const n = (x, dflt) => (x === null ? dflt : typeof x === 'number' ? x : isErr(x) ? x : typeof x === 'boolean' ? +x : ERR.VALUE);
    if (s.kind === KIND.QUANTITY) {
      if (col !== 'H') return null;
      let prod = 1,
        any = false;
      for (const c of ['C', 'D', 'E', 'F']) {
        const x = v(c);
        if (x === null) continue;
        const y = n(x);
        if (isErr(y)) return y;
        prod *= y;
        any = true;
      }
      if (!any) return null;
      const g = n(v('G'), 1);
      return isErr(g) ? g : prod * g;
    }
    if (col === 'F') {
      const q = v('C'),
        r = v('E');
      if (q === null && r === null) return null;
      const a = n(q, 0),
        b = n(r, 0);
      if (isErr(a)) return a;
      if (isErr(b)) return b;
      return a * b;
    }
    if (col === 'H') {
      const f = v('F');
      if (f === null) return null;
      const a = n(f);
      if (isErr(a)) return a;
      const g = n(v('G'), 1);
      return isErr(g) ? g : a * g;
    }
    return null;
  }

  _computeTotal(sheetId) {
    const s = this.sheets.get(sheetId);
    if (!s) return ERR.REF;
    this._track(`rows|${sheetId}`);
    let sum = 0;
    for (let r = 1; r <= s.rowCount; r++) {
      const t = s.rowTypes.get(r);
      if (t === 'subtotal' || t === 'heading') continue;
      const v = this.value(sheetId, `H${r}`);
      if (isErr(v)) return v; // an error anywhere below is surfaced, not silently dropped
      if (typeof v === 'number') sum += v;
    }
    return sum;
  }

  /** Formula evaluation context for a sheet. */
  _ctx(sheetId, link) {
    return {
      link,
      cell: (sheetName, col, row) => {
        let target = sheetId;
        if (sheetName) {
          const t = this.sheetByName(sheetName);
          if (!t) return ERR.REF;
          target = t.id;
        }
        if (!this.columns(target)[col] || row < 1) return ERR.REF;
        return this.value(target, `${col}${row}`);
      },
      name: (id) => {
        const k = id.toLowerCase();
        if (!this.names.has(k)) return undefined;
        return this._read(`n|${k}`, () => {
          const raw = this.names.get(k).raw;
          return raw === '' ? null : this._literal(this.rootId, raw);
        });
      },
      meas: (key) => {
        const m = this._findMeasurement(key);
        if (!m) return undefined;
        this._track(`m|${m.id}`);
        return m.value;
      },
    };
  }
  _findMeasurement(key) {
    if (this.measurements.has(key)) return { id: key, ...this.measurements.get(key) };
    const k = key.toLowerCase();
    for (const [id, m] of this.measurements) if ((m.name ?? '').toLowerCase() === k) return { id, ...m };
    return null;
  }

  // ================================================================ measurements (live, not undoable)

  /** Create or update a measurement. Everything linked to it updates. */
  setMeasurement(id, { value, unit, name } = {}) {
    const prev = this.measurements.get(id) ?? {};
    this.measurements.set(id, { value: value ?? prev.value ?? 0, unit: unit ?? prev.unit ?? '', name: name ?? prev.name ?? id });
    this._emit(this._invalidate([`m|${id}`]));
  }
  removeMeasurement(id) {
    this.measurements.delete(id);
    this._emit(this._invalidate([`m|${id}`]));
  }
  /** Cells linked to a measurement: [{sheetId, a1}] */
  linksTo(measurementId) {
    const out = [];
    for (const s of this.sheets.values()) for (const [a1, c] of s.cells) if (c.link?.id === measurementId) out.push({ sheetId: s.id, a1 });
    return out;
  }

  // ================================================================ transactions / undo

  /** Group several edits into one undo step. */
  batch(label, fn) {
    return this._transact(label, fn);
  }

  _transact(label, fn) {
    if (this._tx) return fn(); // nested → part of the outer transaction
    this._tx = { label, before: new Map(), names: null, structural: false, keys: new Set() };
    let result;
    try {
      result = fn();
    } catch (e) {
      // roll back partial changes
      const tx = this._tx;
      this._tx = null;
      this._restore(tx.before, tx.names);
      this._invalidateAll();
      this._emit(null);
      throw e;
    }
    const tx = this._tx;
    this._tx = null;
    if (tx.before.size || tx.names) {
      const after = new Map();
      for (const id of tx.before.keys()) after.set(id, this.sheets.has(id) ? cloneSheet(this.sheets.get(id)) : null);
      this._undo.push({ label, before: tx.before, after, namesBefore: tx.names, namesAfter: tx.names ? new Map(this.names) : null });
      if (this._undo.length > 500) this._undo.shift();
      this._redo.length = 0;
      if (tx.structural) {
        this._invalidateAll();
        this._emit(null);
      } else this._emit(this._invalidate(tx.keys));
    }
    return result;
  }
  /** Call before modifying a sheet inside a transaction. */
  _touch(sheetId) {
    const tx = this._tx;
    if (!tx) throw new Error('edit outside a transaction');
    if (!tx.before.has(sheetId)) tx.before.set(sheetId, this.sheets.has(sheetId) ? cloneSheet(this.sheets.get(sheetId)) : null);
    return this.sheets.get(sheetId);
  }
  _touchNames() {
    if (!this._tx.names) this._tx.names = new Map(this.names);
  }
  _restore(snap, names) {
    for (const [id, s] of snap) {
      if (s) this.sheets.set(id, cloneSheet(s));
      else this.sheets.delete(id);
    }
    if (names) this.names = new Map(names);
  }
  get canUndo() {
    return this._undo.length > 0;
  }
  get canRedo() {
    return this._redo.length > 0;
  }
  undo() {
    const step = this._undo.pop();
    if (!step) return false;
    this._restore(step.before, step.namesBefore);
    this._redo.push(step);
    this._invalidateAll();
    this._emit(null);
    return step.label;
  }
  redo() {
    const step = this._redo.pop();
    if (!step) return false;
    this._restore(step.after, step.namesAfter);
    this._undo.push(step);
    this._invalidateAll();
    this._emit(null);
    return step.label;
  }

  _emit(keys) {
    this.dispatchEvent(new CustomEvent('change', { detail: { keys } }));
  }

  // ================================================================ editing

  /**
   * Set a cell from what the user typed: a number, text, or a formula ("=...").
   * On a computed column (Subtotal/Total/Quantity) a value is an OVERRIDE;
   * clearing it ('') returns to the built-in calculation.
   */
  setCell(sheetId, a1, raw) {
    return this.setCells(sheetId, [{ a1, raw }]);
  }

  /** @param {{a1:string, raw:string}[]} edits */
  setCells(sheetId, edits, label = 'Edit') {
    return this._transact(label, () => {
      const s = this._touch(sheetId);
      for (const { a1, raw } of edits) {
        const p = parseA1(a1);
        if (!p || !this.columns(sheetId)[p.col]) throw new Error(`no such cell ${a1}`);
        const c = s.cells.get(a1) ?? {};
        if (c.drill) throw new Error(`${a1} has a breakdown — remove it first`);
        const next = { ...c };
        delete next.link; // typing over a link replaces it
        const v = (raw ?? '').toString();
        if (v === '') delete next.raw;
        else next.raw = v;
        if (Object.keys(next).length) s.cells.set(a1, next);
        else s.cells.delete(a1);
        if (p.row > s.rowCount) {
          s.rowCount = p.row;
          this._tx.keys.add(`rows|${sheetId}`);
        }
        this._tx.keys.add(`c|${sheetId}|${a1}`);
      }
    });
  }

  /** Set a cell's note (shown as a tooltip). */
  setNote(sheetId, a1, note) {
    return this._transact('Note', () => {
      const s = this._touch(sheetId);
      const c = { ...(s.cells.get(a1) ?? {}) };
      if (note) c.note = note;
      else delete c.note;
      Object.keys(c).length ? s.cells.set(a1, c) : s.cells.delete(a1);
    });
  }

  /** Mark a row as 'heading' (bold, excluded), 'subtotal' (auto-sums the rows above) or 'normal'. */
  setRowType(sheetId, row, type) {
    return this._transact('Row type', () => {
      const s = this._touch(sheetId);
      if (type === 'normal' || !type) s.rowTypes.delete(row);
      else s.rowTypes.set(row, type);
      if (row > s.rowCount) s.rowCount = row;
      this._tx.structural = true;
    });
  }

  /**
   * Drill down into a cell (creating the breakdown sheet if needed).
   * Quantity → quantity sheet, Rate → rate build-up, Total → cost breakdown.
   * The cell's current figure is carried into the first row so nothing changes until you edit.
   * @returns {string} child sheet id
   */
  drill(sheetId, a1) {
    const s = this.sheets.get(sheetId);
    const c = s?.cells.get(a1);
    if (c?.drill) return c.drill;
    const p = parseA1(a1);
    const def = this.columns(sheetId)[p.col];
    if (!def?.drill) throw new Error(`${def?.label ?? a1} can't be drilled`);
    if (s.rowTypes.get(p.row) === 'subtotal') throw new Error('subtotal rows can’t be drilled');
    const current = this.value(sheetId, a1);
    const keep = typeof current === 'number' && (c?.raw !== undefined || c?.link) ? current : null;

    return this._transact('Drill down', () => {
      const parentDesc = String(this.value(sheetId, `B${p.row}`) ?? '').trim();
      const child = {
        id: uid(),
        // named after the row it breaks down, e.g. "Concrete slab — Quantity" (use it in [..]! refs)
        name: this._uniqueName(`${parentDesc || `${s.name} ${a1}`} — ${drillTitle[def.drill]}`),
        kind: def.drill,
        parent: { sheetId, a1 },
        rowCount: 12,
        cells: new Map(),
        rowTypes: new Map(),
      };
      this._touch(child.id);
      this.sheets.set(child.id, child);
      // carry the existing figure into row 1 so the parent value doesn't jump
      const unit = s.kind !== KIND.QUANTITY ? this.value(sheetId, `D${p.row}`) : null;
      if (def.drill === KIND.QUANTITY) {
        if (c?.link) child.cells.set('C1', { link: { ...c.link } });
        else if (keep !== null) child.cells.set('C1', { raw: c?.raw ?? String(keep) });
        if (keep !== null || c?.link) child.cells.set('B1', { raw: 'Existing quantity' });
      } else if (def.drill === KIND.RATE) {
        if (keep !== null) {
          child.cells.set('B1', { raw: 'Existing rate' });
          child.cells.set('C1', { raw: '1' });
          child.cells.set('E1', { raw: c.raw ?? String(keep) });
          if (unit) child.cells.set('D1', { raw: String(unit) });
        }
      } else {
        // cost breakdown of a Total: start with the row as it stood
        const q = this.value(sheetId, `C${p.row}`),
          r = this.value(sheetId, `E${p.row}`);
        if (typeof q === 'number' || typeof r === 'number') {
          child.cells.set('B1', { raw: parentDesc || 'Existing item' });
          if (typeof q === 'number') child.cells.set('C1', { raw: String(q) });
          if (unit) child.cells.set('D1', { raw: String(unit) });
          if (typeof r === 'number') child.cells.set('E1', { raw: String(r) });
          const g = this.value(sheetId, `G${p.row}`);
          if (typeof g === 'number') child.cells.set('G1', { raw: String(g) });
        }
      }
      const ps = this._touch(sheetId);
      const note = c?.note;
      ps.cells.set(a1, note ? { drill: child.id, note } : { drill: child.id });
      this._tx.structural = true;
      return child.id;
    });
  }

  /** Remove a breakdown (and everything under it). keepValue=true leaves its figure as a typed value. */
  removeDrill(sheetId, a1, { keepValue = true } = {}) {
    const c = this.sheets.get(sheetId)?.cells.get(a1);
    if (!c?.drill) return;
    const v = this.value(sheetId, a1);
    return this._transact('Remove breakdown', () => {
      this._deleteSubtree(c.drill);
      const s = this._touch(sheetId);
      const next = { ...c };
      delete next.drill;
      if (keepValue && typeof v === 'number') next.raw = String(roundClean(v));
      Object.keys(next).length ? s.cells.set(a1, next) : s.cells.delete(a1);
      this._tx.structural = true;
    });
  }

  _deleteSubtree(sheetId) {
    for (const ch of this.children(sheetId)) this._deleteSubtree(ch.id);
    this._touch(sheetId);
    this.sheets.delete(sheetId);
  }

  /**
   * Link a cell to a measurement. Optional modifier formula uses `link`:
   *   link(sheet, 'C4', 'm-17', '=link*1.05')   // +5% waste
   */
  link(sheetId, a1, measurementId, expr) {
    const p = parseA1(a1);
    const def = this.columns(sheetId)[p.col];
    if (!def || def.computed || (def.type === 'text' && !def.user)) throw new Error(`${def?.label ?? a1} can't hold a measurement`);
    return this._transact('Link measurement', () => {
      const s = this._touch(sheetId);
      const c = { ...(s.cells.get(a1) ?? {}) };
      if (c.drill) throw new Error(`${a1} has a breakdown — remove it first`);
      delete c.raw;
      c.link = { id: measurementId, ...(expr ? { expr } : {}) };
      s.cells.set(a1, c);
      // fill the unit from the measurement if the row has none
      const m = this.measurements.get(measurementId);
      if (s.kind !== KIND.QUANTITY && p.col === 'C' && m?.unit && !s.cells.get(`D${p.row}`)?.raw) s.cells.set(`D${p.row}`, { raw: m.unit });
      if (p.row > s.rowCount) s.rowCount = p.row;
      this._tx.structural = true;
    });
  }

  /** Change a link's modifier, e.g. '=link*1.1' (or '' to remove the modifier). */
  setLinkExpr(sheetId, a1, expr) {
    const c = this.sheets.get(sheetId)?.cells.get(a1);
    if (!c?.link) throw new Error(`${a1} is not linked`);
    return this._transact('Link modifier', () => {
      const s = this._touch(sheetId);
      const link = { id: c.link.id };
      if (expr) link.expr = expr.trimStart().startsWith('=') ? expr : `=${expr}`;
      s.cells.set(a1, { ...c, link });
      this._tx.keys.add(`c|${sheetId}|${a1}`);
    });
  }

  unlink(sheetId, a1, { keepValue = true } = {}) {
    const c = this.sheets.get(sheetId)?.cells.get(a1);
    if (!c?.link) return;
    const v = this.value(sheetId, a1);
    return this._transact('Unlink', () => {
      const s = this._touch(sheetId);
      const next = { ...c };
      delete next.link;
      if (keepValue && typeof v === 'number') next.raw = String(roundClean(v));
      Object.keys(next).length ? s.cells.set(a1, next) : s.cells.delete(a1);
      this._tx.keys.add(`c|${sheetId}|${a1}`);
    });
  }

  renameSheet(sheetId, name) {
    const s = this.sheets.get(sheetId);
    const clean = name.replace(/[[\]]/g, '').trim();
    if (!s || !clean || clean === s.name) return;
    if (this.sheetByName(clean) && this.sheetByName(clean) !== s) throw new Error(`a sheet called "${clean}" already exists`);
    return this._transact('Rename sheet', () => {
      const old = s.name;
      this._touch(sheetId).name = clean;
      this._rewriteAll((raw) => renameSheetRefs(raw, old, clean));
      this._tx.structural = true;
    });
  }

  /** Workbook-level named value, e.g. setName('wastage', '1.05') then use =C3*wastage */
  setName(name, raw) {
    if (!/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) || /^[A-Za-z]{1,3}\d+$/.test(name) || /^(true|false|link)$/i.test(name))
      throw new Error(`"${name}" isn't a valid name (letters/digits/_ and not like a cell reference)`);
    return this._transact('Name', () => {
      this._touchNames();
      if (raw === null || raw === undefined) this.names.delete(name.toLowerCase());
      else this.names.set(name.toLowerCase(), { name, raw: String(raw) });
      this._tx.structural = true;
    });
  }

  _uniqueName(base) {
    const clean = base.replace(/[[\]]/g, '').slice(0, 80);
    let n = clean,
      i = 2;
    while (this.sheetByName(n)) n = `${clean} (${i++})`;
    return n;
  }

  // ================================================================ rows

  /** Insert `n` blank rows before row `at`. Formulas everywhere are adjusted, like Excel. */
  insertRows(sheetId, at, n = 1) {
    return this._transact('Insert rows', () => {
      const s = this._touch(sheetId);
      s.cells = shiftCells(s.cells, (row) => (row >= at ? row + n : row));
      s.rowTypes = new Map([...s.rowTypes].map(([r, t]) => [r >= at ? r + n : r, t]));
      s.rowCount += n;
      this._moveChildren(sheetId, (row) => (row >= at ? row + n : row));
      this._rewriteRefsTo(sheetId, (row) => (row >= at ? row + n : row), (r0, r1) => [r0 >= at ? r0 + n : r0, r1 >= at ? r1 + n : r1]);
      this._tx.structural = true;
    });
  }

  /** Delete rows at..at+n-1 (and their breakdowns). References to them become #REF!. */
  deleteRows(sheetId, at, n = 1) {
    return this._transact('Delete rows', () => {
      const s = this._touch(sheetId);
      const end = at + n - 1;
      for (const ch of this.children(sheetId)) {
        const r = parseA1(ch.parent.a1).row;
        if (r >= at && r <= end) this._deleteSubtree(ch.id);
      }
      const map = (row) => (row < at ? row : row > end ? row - n : null);
      s.cells = shiftCells(s.cells, map);
      s.rowTypes = new Map([...s.rowTypes].filter(([r]) => map(r) !== null).map(([r, t]) => [map(r), t]));
      s.rowCount = Math.max(0, s.rowCount - Math.max(0, Math.min(end, s.rowCount) - at + 1));
      this._moveChildren(sheetId, map);
      this._rewriteRefsTo(sheetId, map, (r0, r1) => {
        const a = r0 < at ? r0 : r0 > end ? r0 - n : at;
        const b = r1 < at ? r1 : r1 > end ? r1 - n : at - 1;
        return b < a ? null : [a, b];
      });
      this._tx.structural = true;
    });
  }

  _moveChildren(sheetId, mapRow) {
    for (const ch of [...this.sheets.values()]) {
      if (ch.parent?.sheetId !== sheetId) continue;
      const p = parseA1(ch.parent.a1);
      const r = mapRow(p.row);
      if (r !== null && r !== p.row) this._touch(ch.id).parent = { sheetId, a1: `${p.col}${r}` };
    }
  }

  /** Apply fn to every formula (cells, link modifiers, names); touches only what changes. */
  _rewriteAll(fn) {
    for (const s of [...this.sheets.values()]) {
      for (const [a1, c] of s.cells) {
        const raw2 = isFormula(c.raw) ? '=' + fn(body(c.raw), s) : c.raw;
        const ex2 = c.link?.expr ? '=' + fn(body(c.link.expr), s) : c.link?.expr;
        if (raw2 !== c.raw || ex2 !== c.link?.expr) {
          const t = this._touch(s.id);
          const nc = { ...t.cells.get(a1) };
          if (raw2 !== c.raw) nc.raw = raw2;
          if (ex2 !== c.link?.expr) nc.link = { ...nc.link, expr: ex2 };
          t.cells.set(a1, nc);
        }
      }
    }
    for (const [k, v] of [...this.names]) {
      if (!isFormula(v.raw)) continue;
      const r2 = '=' + fn(body(v.raw), this.root);
      if (r2 !== v.raw) {
        this._touchNames();
        this.names.set(k, { ...v, raw: r2 });
      }
    }
  }

  /** Rewrite references (in any sheet) that point INTO `sheetId`. */
  _rewriteRefsTo(sheetId, mapRow, mapRange) {
    const target = this.sheets.get(sheetId);
    const pointsHere = (prefix, owner) => (prefix ? prefix.toLowerCase() === target.name.toLowerCase() : owner.id === sheetId);
    this._rewriteAll((b, owner) =>
      rewriteRefs(
        b,
        (r, prefix) => {
          if (!pointsHere(prefix, owner)) return r;
          const row = mapRow(r.row);
          return row === null ? null : { ...r, row };
        },
        (rg, prefix) => {
          if (!pointsHere(prefix, owner)) return rg;
          const top = rg.a.row <= rg.b.row ? 'a' : 'b';
          const bot = top === 'a' ? 'b' : 'a';
          const m = mapRange(rg[top].row, rg[bot].row);
          if (!m) return null;
          return { [top]: { ...rg[top], row: m[0] }, [bot]: { ...rg[bot], row: m[1] } };
        },
      ),
    );
  }

  // ================================================================ copy / paste

  /**
   * Copy a rectangle. Returns a clip with the cells (incl. breakdowns as
   * sub-trees and links) plus TSV text for the system clipboard.
   */
  copy(sheetId, r0, c0, r1, c1) {
    const s = this.sheets.get(sheetId);
    const cols = this.columnLetters(sheetId);
    const cells = [];
    const lines = [];
    for (let r = r0; r <= r1; r++) {
      const line = [];
      for (let ci = c0; ci <= c1; ci++) {
        const a1 = `${cols[ci]}${r}`;
        const c = s.cells.get(a1);
        if (c) {
          const entry = { dr: r - r0, dc: ci - c0, cell: structuredClone(c) };
          if (c.drill) entry.subtree = this._exportSubtree(c.drill);
          cells.push(entry);
        }
        const v = this.value(sheetId, a1);
        line.push(c?.raw && isFormula(c.raw) ? c.raw : isErr(v) ? v.code : v === null ? '' : String(v));
      }
      lines.push(line.join('\t'));
    }
    return { kind: 'takeoff-workbook/clip', fromSheet: sheetId, r0, c0, rows: r1 - r0 + 1, cols: c1 - c0 + 1, cells, text: lines.join('\n') };
  }

  /** Paste a clip at (row, colIndex). Formulas shift relatively; breakdowns are deep-copied. */
  paste(sheetId, row, colIndex, clip) {
    const cols = this.columnLetters(sheetId);
    return this._transact('Paste', () => {
      const s = this._touch(sheetId);
      for (let dr = 0; dr < clip.rows; dr++)
        for (let dc = 0; dc < clip.cols; dc++) {
          const col = cols[colIndex + dc];
          if (!col) continue;
          const a1 = `${col}${row + dr}`;
          const old = s.cells.get(a1);
          if (old?.drill) this._deleteSubtree(old.drill);
          s.cells.delete(a1);
        }
      for (const e of clip.cells) {
        const col = cols[colIndex + e.dc];
        if (!col) continue;
        const r = row + e.dr;
        const a1 = `${col}${r}`;
        const c = structuredClone(e.cell);
        const dCol = colIndex - clip.c0,
          dRow = row - clip.r0;
        if (isFormula(c.raw)) c.raw = '=' + offsetFormula(body(c.raw), dCol, dRow);
        if (c.link?.expr) c.link.expr = '=' + offsetFormula(body(c.link.expr), dCol, dRow);
        if (c.drill) {
          const def = this.columns(sheetId)[col];
          if (def?.drill && e.subtree) c.drill = this._importSubtree(e.subtree, { sheetId, a1 });
          else delete c.drill; // pasted somewhere that can't hold a breakdown
        }
        if (Object.keys(c).length) s.cells.set(a1, c);
        if (r > s.rowCount) s.rowCount = r;
      }
      this._tx.structural = true;
    });
  }

  /** Paste plain text (from Excel/Sheets): tab-separated columns, newline-separated rows. */
  pasteText(sheetId, row, colIndex, text) {
    const cols = this.columnLetters(sheetId);
    const rows = text.replace(/\r\n?/g, '\n').replace(/\n$/, '').split('\n').map((l) => l.split('\t'));
    const edits = [];
    rows.forEach((line, dr) =>
      line.forEach((raw, dc) => {
        const col = cols[colIndex + dc];
        if (!col) return;
        const a1 = `${col}${row + dr}`;
        if (this.sheets.get(sheetId).cells.get(a1)?.drill) return; // never overwrite a breakdown with text
        edits.push({ a1, raw: raw.trim() });
      }),
    );
    return this.setCells(sheetId, edits, 'Paste');
  }

  _exportSubtree(rootSheetId) {
    const out = [];
    const walk = (id) => {
      out.push(sheetToJSON(this.sheets.get(id)));
      for (const ch of this.children(id)) walk(ch.id);
    };
    walk(rootSheetId);
    return out;
  }
  _importSubtree(list, parent) {
    const idMap = new Map(list.map((j) => [j.id, uid()]));
    for (const j of list) {
      const s = sheetFromJSON(structuredClone(j));
      s.id = idMap.get(j.id);
      s.name = this._uniqueName(j.name);
      s.parent = j.id === list[0].id ? parent : { ...s.parent, sheetId: idMap.get(s.parent.sheetId) };
      for (const [a1, c] of s.cells) if (c.drill) s.cells.set(a1, { ...c, drill: idMap.get(c.drill) });
      this._touch(s.id);
      this.sheets.set(s.id, s);
    }
    return idMap.get(list[0].id);
  }

  // ================================================================ persistence

  /** Plain JSON (one document for the whole tree — no separately-stored totals to drift). */
  toJSON() {
    return {
      format: 'takeoff-workbook',
      version: 1,
      rootId: this.rootId,
      userColumns: this.userColumns,
      userColumnLabels: this.userColumnLabels,
      names: [...this.names.values()],
      sheets: [...this.sheets.values()].map(sheetToJSON),
      // last known measurement values, so a reopened estimate shows the right
      // numbers before the takeoff app pushes live ones
      measurements: Object.fromEntries(this.measurements),
    };
  }

  static fromJSON(j) {
    if (j?.format !== 'takeoff-workbook') throw new Error('not a takeoff-workbook document');
    const wb = new Workbook({ userColumns: j.userColumns, userColumnLabels: j.userColumnLabels, _empty: true });
    wb.rootId = j.rootId;
    for (const s of j.sheets) wb.sheets.set(s.id, sheetFromJSON(s));
    for (const n of j.names ?? []) wb.names.set(n.name.toLowerCase(), n);
    for (const [id, m] of Object.entries(j.measurements ?? {})) wb.measurements.set(id, m);
    return wb;
  }
}

// ---------------------------------------------------------------- helpers

function shiftCells(cells, mapRow) {
  const out = new Map();
  for (const [a1, c] of cells) {
    const p = parseA1(a1);
    const r = mapRow(p.row);
    if (r !== null) out.set(`${p.col}${r}`, c);
  }
  return out;
}

const fmtCache = new Map();
function fmtNum(v, min, max) {
  const k = `${min}|${max}`;
  let f = fmtCache.get(k);
  if (!f) fmtCache.set(k, (f = new Intl.NumberFormat('en-NZ', { minimumFractionDigits: min, maximumFractionDigits: max })));
  return f.format(v);
}

/** 12.300000000000001 → 12.3 (for values written back as text) */
const roundClean = (v) => +v.toPrecision(12);

export { ERR, Err, isErr };
