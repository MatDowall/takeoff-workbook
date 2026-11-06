/**
 * WorkbookView — the grid UI for a Workbook. Framework-free DOM, one class.
 *
 *   const view = new WorkbookView(el, workbook, {
 *     onRequestLink: (sheetId, a1) => { ...pick a measurement...; view.linkActive(id) },
 *   });
 *
 * Spreadsheet keys: arrows/Tab/Enter move · type to replace · F2 or double-click to edit ·
 * Delete clears · Ctrl+Z / Ctrl+Y · Ctrl+C / Ctrl+V (Excel/Sheets text or internal with breakdowns) ·
 * Ctrl+Enter drills down (or opens a breakdown) · Alt+↑ goes up a level · right-click for the menu.
 *
 * Rendering is a plain <table>: estimate sheets are hundreds of rows, not
 * hundreds of thousands, and the engine is memoised, so a full value refresh
 * on every change is ~1 ms. (Swap in row virtualisation if you ever need 50k rows.)
 */
import { parseA1, isErr } from './formula.js';

const CSS = `
.twb{--twb-line:#dfe4e6;--twb-head:#f1f4f5;--twb-sel:#1a73e8;--twb-comp:#f7faf9;--twb-drill:#3d8179;--twb-link:#c77b30;--twb-err:#d32f2f;
  font:13px system-ui,sans-serif;color:#1d2a30;display:flex;flex-direction:column;height:100%;outline:none;background:#fff;min-height:0;min-width:0;overflow:hidden}
.twb-bar{display:flex;gap:8px;align-items:center;padding:6px 8px;border-bottom:1px solid var(--twb-line);flex-wrap:wrap}
.twb-crumbs{display:flex;gap:4px;align-items:center;flex:1;min-width:0;overflow:hidden;white-space:nowrap}
.twb-crumbs button{border:0;background:none;color:var(--twb-drill);cursor:pointer;font:inherit;padding:2px 4px;border-radius:4px;max-width:260px;overflow:hidden;text-overflow:ellipsis}
.twb-crumbs button:hover{background:#eef5f4}.twb-crumbs button[aria-current]{color:#1d2a30;font-weight:600;cursor:default;background:none}
.twb-crumbs .kind{font-size:11px;color:#7a8a90;margin-left:2px}
.twb-total{font-weight:700;font-variant-numeric:tabular-nums}
.twb-btn{border:1px solid var(--twb-line);background:#fff;border-radius:4px;padding:2px 8px;cursor:pointer;font:inherit}.twb-btn:disabled{opacity:.4;cursor:default}
.twb-fx{display:flex;border-bottom:1px solid var(--twb-line)}
.twb-fx .addr{width:64px;padding:4px 8px;border-right:1px solid var(--twb-line);color:#56646a;font-variant-numeric:tabular-nums}
.twb-fx input{flex:1;border:0;padding:4px 8px;font:inherit;outline:none}
.twb-fx .hint{padding:4px 8px;color:#7a8a90;font-size:12px;white-space:nowrap}
.twb-scroll{flex:1;overflow:auto;min-height:0;position:relative}
.twb table{border-collapse:separate;border-spacing:0;table-layout:fixed}
.twb th,.twb td{border-right:1px solid var(--twb-line);border-bottom:1px solid var(--twb-line);height:24px;padding:0 6px;overflow:hidden;white-space:nowrap;text-overflow:ellipsis;box-sizing:border-box}
.twb thead th{position:sticky;top:0;background:var(--twb-head);z-index:2;font-weight:600;text-align:left;font-size:12px}
.twb thead th small{color:#98a6ab;font-weight:400;margin-right:4px}
.twb th.rn{position:sticky;left:0;background:var(--twb-head);z-index:1;width:40px;text-align:right;color:#7a8a90;font-weight:400;font-size:12px}
.twb thead th.rn{z-index:3}
.twb td{cursor:cell;position:relative}
.twb td.num{text-align:right;font-variant-numeric:tabular-nums}
.twb td.comp{background:var(--twb-comp)}
.twb td.ovr{font-style:italic}.twb td.ovr::after{content:'';position:absolute;right:0;top:0;border:4px solid transparent;border-top-color:#8e24aa;border-right-color:#8e24aa}
.twb td.drill{color:var(--twb-drill);font-weight:600;padding-left:18px}
.twb td.drill .open{position:absolute;left:2px;top:3px;width:14px;height:16px;border:0;background:none;color:var(--twb-drill);cursor:pointer;padding:0;font-size:12px}
.twb td.link{color:var(--twb-link)}.twb td.link::before{content:'⛓';position:absolute;left:3px;font-size:10px;opacity:.7}
.twb td.err,.twb td.missing{color:var(--twb-err)}
.twb td.note::after{content:'';position:absolute;right:0;top:0;border:4px solid transparent;border-top-color:#f9a825;border-right-color:#f9a825}
.twb tr.heading td{font-weight:700;background:#f4f7f8}
.twb tr.subtotal td{font-weight:700;border-top:2px solid #9fb0b6}
.twb td.insel{background:#e8f0fe}
.twb td.active{outline:2px solid var(--twb-sel);outline-offset:-2px;z-index:1}
.twb td.editing{padding:0;outline:2px solid var(--twb-sel);outline-offset:-2px}
.twb td.editing input{width:100%;height:100%;border:0;padding:0 6px;font:inherit;outline:none;box-sizing:border-box;background:#fff;text-align:inherit}
.twb tfoot td{position:sticky;bottom:0;background:#eef2f3;font-weight:700;border-top:2px solid #9fb0b6}
.twb-status{display:flex;gap:14px;padding:3px 8px;border-top:1px solid var(--twb-line);color:#56646a;font-size:12px;min-height:18px}
.twb-status .msg{color:#b35c00;margin-left:auto}
.twb-menu{position:fixed;z-index:1000;background:#fff;border:1px solid #cfd8dc;border-radius:6px;box-shadow:0 6px 24px rgba(0,0,0,.18);padding:4px;min-width:220px;font:13px system-ui,sans-serif}
.twb-menu button{display:block;width:100%;text-align:left;border:0;background:none;padding:6px 10px;border-radius:4px;cursor:pointer;font:inherit}
.twb-menu button:hover:not(:disabled){background:#eef5f4}.twb-menu button:disabled{color:#aab4b8;cursor:default}
.twb-menu hr{border:0;border-top:1px solid #eceff1;margin:4px 0}
`;
let cssInjected = false;

const KIND_LABEL = { cost: 'cost', rate: 'rate build-up', quantity: 'quantity' };

export class WorkbookView extends EventTarget {
  /**
   * @param {HTMLElement} container
   * @param {import('./Workbook.js').Workbook} wb
   * @param {object} [o]
   * @param {(sheetId:string, a1:string)=>void} [o.onRequestLink]  "Link measurement…" menu item
   * @param {number} [o.spareRows=5]  blank rows shown below the last used row
   * @param {number} [o.minRows=20]
   */
  constructor(container, wb, o = {}) {
    super();
    if (!cssInjected && typeof document !== 'undefined') {
      const st = document.createElement('style');
      st.textContent = CSS;
      document.head.appendChild(st);
      cssInjected = true;
    }
    this.wb = wb;
    this.opts = { spareRows: 5, minRows: 20, ...o };
    this.sheetId = wb.rootId;
    this.sel = { ar: 1, ac: 0, r0: 1, c0: 0, r1: 1, c1: 0 };
    this.editing = null;
    this._clip = null;
    this._build(container);
    this._onChange = (e) => this._changed(e.detail.keys);
    wb.addEventListener('change', this._onChange);
    this.render();
  }

  // ---------------------------------------------------------------- public API

  /** Show a sheet (e.g. after drilling, or from your own navigation). */
  openSheet(sheetId, a1) {
    if (!this.wb.sheet(sheetId)) return;
    this._commitEdit();
    this.sheetId = sheetId;
    const p = a1 ? parseA1(a1) : { row: 1, col: 'A' };
    const ci = Math.max(0, this.wb.columnLetters(sheetId).indexOf(p.col));
    this.sel = { ar: p.row, ac: ci, r0: p.row, c0: ci, r1: p.row, c1: ci };
    this.render();
    this.root.focus();
  }
  /** The active cell: { sheetId, a1, col, row } */
  activeCell() {
    const col = this._cols[this.sel.ac];
    return { sheetId: this.sheetId, a1: `${col}${this.sel.ar}`, col, row: this.sel.ar };
  }
  /** Link the active cell to a measurement (call from your takeoff UI). */
  linkActive(measurementId, expr) {
    const { a1 } = this.activeCell();
    try {
      this.wb.link(this.sheetId, a1, measurementId, expr);
      this._flash(`Linked ${a1}`);
    } catch (e) {
      this._flash(e.message);
    }
  }
  focus() {
    this.root.focus();
  }
  destroy() {
    this.wb.removeEventListener('change', this._onChange);
    this._menu?.remove();
    this.root.remove();
  }

  // ---------------------------------------------------------------- DOM

  _build(container) {
    const r = (this.root = document.createElement('div'));
    r.className = 'twb';
    r.tabIndex = 0;
    r.innerHTML = `
      <div class="twb-bar">
        <button class="twb-btn" data-act="up" title="Up a level (Alt+↑)">↑</button>
        <div class="twb-crumbs"></div>
        <span class="twb-total"></span>
        <button class="twb-btn" data-act="undo" title="Undo (Ctrl+Z)">↶</button>
        <button class="twb-btn" data-act="redo" title="Redo (Ctrl+Y)">↷</button>
      </div>
      <div class="twb-fx"><span class="addr"></span><input spellcheck="false" placeholder="value, text or =formula"/><span class="hint"></span></div>
      <div class="twb-scroll"><table><colgroup></colgroup><thead></thead><tbody></tbody><tfoot></tfoot></table></div>
      <div class="twb-status"><span class="sum"></span><span class="msg"></span></div>`;
    container.appendChild(r);
    this.$ = (s) => r.querySelector(s);
    this.table = this.$('table');
    this.tbody = this.$('tbody');
    this.fxInput = this.$('.twb-fx input');

    r.addEventListener('click', (e) => {
      const act = e.target.closest('[data-act]')?.dataset.act;
      if (act === 'up') this._up();
      if (act === 'undo') this.wb.undo();
      if (act === 'redo') this.wb.redo();
      const crumb = e.target.closest('[data-sheet]');
      if (crumb) this.openSheet(crumb.dataset.sheet, crumb.dataset.a1);
      const open = e.target.closest('.open');
      if (open) {
        const td = open.closest('td');
        this.openSheet(this.wb.sheet(this.sheetId).cells.get(td.dataset.a1).drill);
      }
    });
    this.tbody.addEventListener('mousedown', (e) => this._mouseDown(e));
    this.tbody.addEventListener('dblclick', (e) => {
      const td = e.target.closest('td');
      if (td) this._enterCell(false);
    });
    r.addEventListener('keydown', (e) => this._key(e));
    r.addEventListener('contextmenu', (e) => {
      const td = e.target.closest('tbody td');
      if (!td) return;
      e.preventDefault();
      if (!this._inSel(+td.dataset.r, +td.dataset.c)) this._select(+td.dataset.r, +td.dataset.c);
      this._showMenu(e.clientX, e.clientY);
    });
    r.addEventListener('copy', (e) => this._copy(e));
    r.addEventListener('paste', (e) => this._paste(e));
    this.fxInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        this._commitRaw(this.fxInput.value);
        this.root.focus();
      } else if (e.key === 'Escape') {
        this._renderFx();
        this.root.focus();
      }
      e.stopPropagation();
    });
    document.addEventListener('mousedown', (this._docDown = (e) => this._menu && !this._menu.contains(e.target) && this._hideMenu()));
  }

  render() {
    const wb = this.wb;
    const s = wb.sheet(this.sheetId);
    if (!s) {
      this.sheetId = wb.rootId;
      return this.render();
    }
    const defs = wb.columns(this.sheetId);
    this._cols = Object.keys(defs);
    this._defs = defs;
    const nRows = Math.max(this.opts.minRows, s.rowCount + this.opts.spareRows, this.sel.r1 + 1);
    this._nRows = nRows;

    this.$('colgroup').innerHTML = `<col style="width:40px">` + this._cols.map((c) => `<col style="width:${defs[c].width}px">`).join('');
    this.table.style.width = 40 + this._cols.reduce((a, c) => a + defs[c].width, 0) + 'px';
    this.$('thead').innerHTML =
      `<tr><th class="rn"></th>` + this._cols.map((c) => `<th title="${esc(defs[c].label)}${defs[c].computed ? ' (calculated)' : ''}"><small>${c}</small>${esc(defs[c].label)}</th>`).join('') + '</tr>';
    let html = '';
    for (let r = 1; r <= nRows; r++) {
      html += `<tr data-r="${r}"><th class="rn">${r}</th>`;
      for (let ci = 0; ci < this._cols.length; ci++) html += `<td data-r="${r}" data-c="${ci}" data-a1="${this._cols[ci]}${r}"></td>`;
      html += '</tr>';
    }
    this.tbody.innerHTML = html;
    const totalCol = this._cols.indexOf('H');
    this.$('tfoot').innerHTML =
      `<tr><th class="rn"></th>` +
      this._cols.map((c, i) => (i === totalCol - 1 ? `<td style="text-align:right">${s.parent ? '→ parent' : 'Total'}</td>` : i === totalCol ? `<td class="num tot"></td>` : '<td></td>')).join('') +
      '</tr>';
    this._renderCrumbs();
    this._refresh();
  }

  _renderCrumbs() {
    const path = this.wb.path(this.sheetId);
    this.$('.twb-crumbs').innerHTML = path
      .map((s, i) => {
        const cur = i === path.length - 1;
        // when clicking a parent crumb, land on the cell that drills into the next level
        const back = path[i + 1]?.parent?.a1 ?? '';
        return `<button data-sheet="${s.id}" data-a1="${back}" ${cur ? 'aria-current="page"' : ''} title="${esc(s.name)}">${esc(s.name)}</button>`;
      })
      .join('<span>›</span>');
    this.$('[data-act="up"]').disabled = path.length < 2;
  }

  /** Refresh all values/classes (cheap: memoised engine). */
  _refresh() {
    const wb = this.wb;
    const s = wb.sheet(this.sheetId);
    for (const tr of this.tbody.rows) {
      const r = +tr.dataset.r;
      const rt = s.rowTypes.get(r);
      tr.className = rt && rt !== 'normal' ? rt : '';
      for (const td of tr.cells) {
        if (td.tagName !== 'TD' || td === this.editing?.td) continue;
        const ci = +td.dataset.c;
        const col = this._cols[ci];
        const a1 = td.dataset.a1;
        const def = this._defs[col];
        const v = wb.value(this.sheetId, a1);
        const info = wb.info(this.sheetId, a1);
        const cls = [];
        if (typeof v === 'number' || def.type === 'num') cls.push('num');
        if (def.computed) cls.push('comp');
        if (info.type === 'override') cls.push('ovr');
        if (info.type === 'link') cls.push('link', ...(info.missing ? ['missing'] : []));
        if (isErr(v)) cls.push('err');
        const c = s.cells.get(a1);
        if (c?.note) cls.push('note');
        const text = wb.format(this.sheetId, a1, v);
        if (info.type === 'drill') {
          cls.push('drill');
          td.innerHTML = `<button class="open" tabindex="-1" title="Open breakdown">▸</button>${esc(text)}`;
        } else td.textContent = text;
        td.title = this._tooltip(info, c, v);
        if (this._inSel(r, ci)) cls.push('insel');
        if (r === this.sel.ar && ci === this.sel.ac) cls.push('active');
        td.className = cls.join(' ');
      }
    }
    const tot = wb.total(this.sheetId);
    const totText = wb.format(this.sheetId, 'H1', tot);
    const tf = this.$('tfoot .tot');
    if (tf) {
      tf.textContent = totText;
      tf.classList.toggle('err', isErr(tot));
    }
    const s0 = this.wb.sheet(this.sheetId);
    this.$('.twb-total').textContent = s0.parent ? '' : `Total ${totText}`;
    this.$('[data-act="undo"]').disabled = !wb.canUndo;
    this.$('[data-act="redo"]').disabled = !wb.canRedo;
    this._renderFx();
    this._renderStatus();
  }

  _tooltip(info, c, v) {
    const parts = [];
    if (info.type === 'drill') parts.push(`Breakdown (${KIND_LABEL[info.kind]}) — Ctrl+Enter or ▸ to open`);
    if (info.type === 'link') parts.push(info.missing ? '⚠ Linked measurement no longer exists' : `Linked to measurement "${info.name}"${info.expr ? ` · modified by ${info.expr}` : ''}`);
    if (info.type === 'override') parts.push(`Overrides the calculation (${info.raw}) — clear the cell to restore`);
    if (info.type === 'formula') parts.push(info.raw);
    if (isErr(v)) parts.push(ERR_HELP[v.code] ?? v.code);
    if (c?.note) parts.push(`Note: ${c.note}`);
    return parts.join('\n');
  }

  _renderFx() {
    const { a1 } = this.activeCell();
    this.$('.twb-fx .addr').textContent = a1;
    const info = this.wb.info(this.sheetId, a1);
    const c = this.wb.sheet(this.sheetId).cells.get(a1);
    let val = c?.raw ?? '';
    let hint = '';
    if (info.type === 'drill') {
      val = '';
      hint = `breakdown: ${this.wb.sheet(info.sheetId)?.name ?? '?'}`;
    } else if (info.type === 'link') {
      val = c.link.expr ?? '=link';
      hint = info.missing ? '⚠ measurement deleted' : `⛓ ${info.name}`;
    } else if (info.type === 'computed') hint = 'calculated';
    else if (info.type === 'subtotal') hint = 'subtotal of rows above';
    this.fxInput.value = val;
    this.fxInput.disabled = info.type === 'drill' || info.type === 'subtotal';
    this.$('.twb-fx .hint').textContent = hint;
  }

  _renderStatus() {
    const { r0, r1, c0, c1 } = this._rect();
    let sum = 0,
      n = 0;
    if (r1 - r0 + c1 - c0 > 0) {
      for (let r = r0; r <= r1; r++)
        for (let c = c0; c <= c1; c++) {
          const v = this.wb.value(this.sheetId, `${this._cols[c]}${r}`);
          if (typeof v === 'number') {
            sum += v;
            n++;
          }
        }
    }
    this.$('.twb-status .sum').textContent = n ? `Sum ${this.wb.format(this.sheetId, 'H1', sum)} · Count ${n}` : '';
  }

  _flash(msg) {
    const el = this.$('.twb-status .msg');
    el.textContent = msg;
    clearTimeout(this._flashT);
    this._flashT = setTimeout(() => (el.textContent = ''), 4000);
  }

  _changed(keys) {
    const s = this.wb.sheet(this.sheetId);
    if (!s) return this.render(); // sheet deleted (e.g. undo of a drill) → back to root
    if (keys === null) {
      // structural: rebuild if row count changed, else just refresh
      const need = Math.max(this.opts.minRows, s.rowCount + this.opts.spareRows, this.sel.r1 + 1);
      if (need !== this._nRows) return this.render();
      this._renderCrumbs();
    }
    this._refresh();
  }

  // ---------------------------------------------------------------- selection

  _rect() {
    const { r0, r1, c0, c1 } = this.sel;
    return { r0: Math.min(r0, r1), r1: Math.max(r0, r1), c0: Math.min(c0, c1), c1: Math.max(c0, c1) };
  }
  _inSel(r, c) {
    const q = this._rect();
    return r >= q.r0 && r <= q.r1 && c >= q.c0 && c <= q.c1;
  }
  _select(r, c, extend = false) {
    r = Math.max(1, r);
    c = Math.max(0, Math.min(this._cols.length - 1, c));
    if (extend) this.sel = { ...this.sel, r1: r, c1: c };
    else this.sel = { ar: r, ac: c, r0: r, c0: c, r1: r, c1: c };
    if (r > this._nRows - 1) return this.render();
    this._refresh();
    this.tbody.querySelector(`td[data-r="${extend ? r : this.sel.ar}"][data-c="${extend ? c : this.sel.ac}"]`)?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
    this.dispatchEvent(new CustomEvent('select', { detail: this.activeCell() }));
  }
  _mouseDown(e) {
    const td = e.target.closest('td');
    if (!td || e.button !== 0 || e.target.closest('.open')) return;
    if (this.editing && this.editing.td === td) return;
    this._commitEdit();
    this._select(+td.dataset.r, +td.dataset.c, e.shiftKey);
    const move = (ev) => {
      const t = document.elementFromPoint(ev.clientX, ev.clientY)?.closest?.('tbody td');
      if (t) {
        const r = +t.dataset.r,
          c = +t.dataset.c;
        if (r !== this.sel.r1 || c !== this.sel.c1) {
          this.sel = { ...this.sel, r1: r, c1: c };
          this._refresh();
        }
      }
    };
    const up = () => {
      removeEventListener('mousemove', move);
      removeEventListener('mouseup', up);
    };
    addEventListener('mousemove', move);
    addEventListener('mouseup', up);
    this.root.focus();
    e.preventDefault();
  }

  // ---------------------------------------------------------------- keyboard

  _key(e) {
    if (this.editing) return; // the editor input handles its own keys
    const mod = e.ctrlKey || e.metaKey;
    const { ar, ac } = this.sel;
    const ext = e.shiftKey;
    const mv = (dr, dc) => {
      e.preventDefault();
      const r = (ext ? this.sel.r1 : ar) + dr,
        c = (ext ? this.sel.c1 : ac) + dc;
      this._select(r, c, ext);
    };
    if (e.altKey && e.key === 'ArrowUp') return e.preventDefault(), this._up();
    switch (e.key) {
      case 'ArrowUp':
        return mv(-1, 0);
      case 'ArrowDown':
        return mv(1, 0);
      case 'ArrowLeft':
        return mv(0, -1);
      case 'ArrowRight':
        return mv(0, 1);
      case 'Tab':
        e.preventDefault();
        return this._select(ar, ac + (e.shiftKey ? -1 : 1));
      case 'Enter':
        e.preventDefault();
        if (mod) return this._drillActive();
        return this._enterCell(false);
      case 'F2':
        e.preventDefault();
        return this._enterCell(false);
      case 'Delete':
      case 'Backspace':
        e.preventDefault();
        return this._clearSel();
      case 'Escape':
        return this._hideMenu();
    }
    if (mod && (e.key === 'z' || e.key === 'Z')) {
      e.preventDefault();
      return e.shiftKey ? this.wb.redo() : this.wb.undo();
    }
    if (mod && (e.key === 'y' || e.key === 'Y')) return e.preventDefault(), this.wb.redo();
    if (mod && e.key === 'a') {
      e.preventDefault();
      this.sel = { ...this.sel, r0: 1, c0: 0, r1: this.wb.sheet(this.sheetId).rowCount || 1, c1: this._cols.length - 1 };
      return this._refresh();
    }
    if (!mod && !e.altKey && e.key.length === 1) {
      e.preventDefault();
      this._enterCell(true, e.key);
    }
  }

  // ---------------------------------------------------------------- editing

  _enterCell(replace, firstChar = '') {
    const { a1 } = this.activeCell();
    const info = this.wb.info(this.sheetId, a1);
    if (info.type === 'drill') {
      if (replace) return this._flash('This figure comes from its breakdown — Ctrl+Enter to open, or right-click → Remove breakdown');
      return this.openSheet(info.sheetId);
    }
    if (info.type === 'subtotal') return this._flash('Subtotal rows add up the rows above automatically');
    const td = this.tbody.querySelector(`td[data-a1="${a1}"]`);
    if (!td) return;
    const c = this.wb.sheet(this.sheetId).cells.get(a1);
    const initial = replace ? firstChar : info.type === 'link' ? (c.link.expr ?? '=link') : (c?.raw ?? '');
    const input = document.createElement('input');
    input.value = initial;
    input.spellcheck = false;
    td.className += ' editing';
    td.textContent = '';
    td.appendChild(input);
    this.editing = { td, a1, input, wasLink: info.type === 'link' && !replace };
    input.focus();
    input.setSelectionRange(input.value.length, input.value.length);
    input.addEventListener('input', () => (this.fxInput.value = input.value));
    input.addEventListener('keydown', (e) => {
      e.stopPropagation();
      const go = (dr, dc) => {
        e.preventDefault();
        this._commitEdit();
        this._select(this.sel.ar + dr, this.sel.ac + dc);
        this.root.focus();
      };
      if (e.key === 'Enter') go(e.shiftKey ? -1 : 1, 0);
      else if (e.key === 'Tab') go(0, e.shiftKey ? -1 : 1);
      else if (e.key === 'Escape') {
        e.preventDefault();
        this._cancelEdit();
      } else if ((e.key === 'ArrowUp' || e.key === 'ArrowDown') && replace) go(e.key === 'ArrowUp' ? -1 : 1, 0);
    });
    input.addEventListener('blur', () => setTimeout(() => this.editing?.input === input && this._commitEdit(), 0));
  }

  _commitEdit() {
    const ed = this.editing;
    if (!ed) return;
    this.editing = null;
    const v = ed.input.value;
    if (ed.wasLink && /^\s*=.*\blink\b/i.test(v)) {
      // editing a linked cell's modifier keeps the link
      try {
        this.wb.setLinkExpr(this.sheetId, ed.a1, v.trim() === '=link' ? '' : v);
      } catch (e) {
        this._flash(e.message);
      }
      return this._refresh();
    }
    this._commitRaw(v, ed.a1);
  }
  _cancelEdit() {
    this.editing = null;
    this._refresh();
    this.root.focus();
  }
  _commitRaw(raw, a1 = this.activeCell().a1) {
    const cur = this.wb.sheet(this.sheetId).cells.get(a1);
    if ((cur?.raw ?? '') === raw && !cur?.link) return this._refresh();
    try {
      this.wb.setCell(this.sheetId, a1, raw);
    } catch (e) {
      this._flash(e.message);
      this._refresh();
    }
  }

  _clearSel() {
    const { r0, r1, c0, c1 } = this._rect();
    const s = this.wb.sheet(this.sheetId);
    const edits = [];
    let skipped = 0;
    for (let r = r0; r <= r1; r++)
      for (let c = c0; c <= c1; c++) {
        const a1 = `${this._cols[c]}${r}`;
        const cell = s.cells.get(a1);
        if (cell?.drill) {
          skipped++;
          continue;
        }
        if (cell?.raw !== undefined || cell?.link) edits.push({ a1, raw: '' });
      }
    if (edits.length) this.wb.setCells(this.sheetId, edits, 'Clear');
    if (skipped) this._flash(`${skipped} cell(s) with a breakdown left alone — right-click → Remove breakdown`);
  }

  // ---------------------------------------------------------------- drill / navigation

  _drillActive() {
    const { a1 } = this.activeCell();
    try {
      const id = this.wb.drill(this.sheetId, a1);
      this.openSheet(id);
    } catch (e) {
      this._flash(e.message);
    }
  }
  _up() {
    const s = this.wb.sheet(this.sheetId);
    if (s.parent) this.openSheet(s.parent.sheetId, s.parent.a1);
  }

  // ---------------------------------------------------------------- clipboard

  _copy(e) {
    if (this.editing) return;
    e.preventDefault();
    const { r0, r1, c0, c1 } = this._rect();
    this._clip = this.wb.copy(this.sheetId, r0, c0, r1, c1);
    e.clipboardData.setData('text/plain', this._clip.text);
    this._flash(`Copied ${r1 - r0 + 1}×${c1 - c0 + 1}`);
  }
  _paste(e) {
    if (this.editing) return;
    e.preventDefault();
    const text = e.clipboardData.getData('text/plain');
    const { ar, ac } = this.sel;
    try {
      // our own copy → rich paste (formulas shift, breakdowns + links come along)
      if (this._clip && norm(text) === norm(this._clip.text)) this.wb.paste(this.sheetId, ar, ac, this._clip);
      else if (text) this.wb.pasteText(this.sheetId, ar, ac, text);
    } catch (err) {
      this._flash(err.message);
    }
  }

  // ---------------------------------------------------------------- context menu

  _showMenu(x, y) {
    this._hideMenu();
    const { a1, col, row } = this.activeCell();
    const def = this._defs[col];
    const info = this.wb.info(this.sheetId, a1);
    const { r0, r1 } = this._rect();
    const nRows = r1 - r0 + 1;
    const rt = this.wb.rowType(this.sheetId, row);
    const items = [
      info.type === 'drill'
        ? ['Open breakdown', () => this.openSheet(info.sheetId), 'Ctrl+Enter']
        : [`Drill down${def.drill ? ` (${KIND_LABEL[def.drill]})` : ''}`, () => this._drillActive(), 'Ctrl+Enter', !def.drill || rt === 'subtotal'],
      ['Remove breakdown (keep figure)', () => this.wb.removeDrill(this.sheetId, a1), '', info.type !== 'drill'],
      '-',
      ['Link measurement…', () => this.opts.onRequestLink?.(this.sheetId, a1), '', !this.opts.onRequestLink || def.computed || (def.type === 'text' && !def.user) || info.type === 'drill'],
      ['Edit link modifier…', () => this._enterCell(false), 'F2', info.type !== 'link'],
      ['Unlink (keep figure)', () => this.wb.unlink(this.sheetId, a1), '', info.type !== 'link'],
      '-',
      ['Insert row above', () => this.wb.insertRows(this.sheetId, r0, nRows)],
      ['Insert row below', () => this.wb.insertRows(this.sheetId, r1 + 1, nRows)],
      [`Delete row${nRows > 1 ? `s ${r0}–${r1}` : ''}`, () => this.wb.deleteRows(this.sheetId, r0, nRows)],
      '-',
      ['Row: normal', () => this.wb.setRowType(this.sheetId, row, 'normal'), '', rt === 'normal'],
      ['Row: heading', () => this.wb.setRowType(this.sheetId, row, 'heading'), '', rt === 'heading'],
      ['Row: subtotal', () => this.wb.setRowType(this.sheetId, row, 'subtotal'), '', rt === 'subtotal'],
      '-',
      ['Add / edit note…', () => {
        const n = prompt('Note for ' + a1, this.wb.sheet(this.sheetId).cells.get(a1)?.note ?? '');
        if (n !== null) this.wb.setNote(this.sheetId, a1, n);
      }],
    ];
    const m = (this._menu = document.createElement('div'));
    m.className = 'twb-menu';
    for (const it of items) {
      if (it === '-') {
        m.appendChild(document.createElement('hr'));
        continue;
      }
      const [label, fn, key, disabled] = it;
      const b = document.createElement('button');
      b.innerHTML = `${esc(label)}${key ? `<span style="float:right;color:#98a6ab">${key}</span>` : ''}`;
      b.disabled = !!disabled;
      b.onclick = () => {
        this._hideMenu();
        try {
          fn();
        } catch (e) {
          this._flash(e.message);
        }
        this.root.focus();
      };
      m.appendChild(b);
    }
    document.body.appendChild(m);
    const r = m.getBoundingClientRect();
    m.style.left = Math.min(x, innerWidth - r.width - 4) + 'px';
    m.style.top = Math.min(y, innerHeight - r.height - 4) + 'px';
  }
  _hideMenu() {
    this._menu?.remove();
    this._menu = null;
  }
}

const ERR_HELP = {
  '#REF!': '#REF! — refers to something that no longer exists (deleted row, sheet or measurement)',
  '#DIV/0!': '#DIV/0! — division by zero',
  '#NAME?': '#NAME? — unknown name or function',
  '#VALUE!': '#VALUE! — text where a number was expected',
  '#CIRC!': '#CIRC! — circular reference (this value depends on itself)',
  '#ERROR!': '#ERROR! — the formula has a syntax error',
};
const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]);
const norm = (t) => t.replace(/\r\n?/g, '\n').replace(/\n$/, '');
