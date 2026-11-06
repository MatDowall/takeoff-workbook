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
 * FORMULA EDITING (Excel-style)
 *  - Autocomplete: typing a function/name shows matching entries at the cursor
 *    (↑↓ choose, Tab/Enter insert). Inside a call, the argument hint shows the
 *    signature with the current argument in bold. Inside MEAS(" it lists measurements.
 *  - Point mode: wherever a reference can go (after = ( , + - * / ^ & < > :) a
 *    click on a cell inserts its reference; drag → range; arrow keys move the
 *    pointed reference, Shift+arrows extend it. Navigate to ANOTHER sheet (breadcrumb
 *    or ▸) mid-formula and clicks insert [Sheet]!A1 — Enter commits back to the
 *    original cell. F2 toggles between point mode and caret-moving edit mode.
 *  - Every reference in the formula is outlined on the grid in its own colour.
 *
 * LAYOUT: drag a column header's right edge / a row number's bottom edge to
 * resize; double-click the edge to auto-fit (columns) or reset (rows).
 * Column widths apply to every sheet of the same kind and are saved + undoable.
 *
 * Rendering is a plain <table>: estimate sheets are hundreds of rows, and the
 * engine is memoised, so a full value refresh on every change is ~1 ms.
 */
import { parseA1, isErr, tokenize, caretContext, FUNCTIONS, FUNCTION_INFO, colToIndex, indexToCol } from './formula.js';

const ROW_H = 24;
const HL_COLOURS = ['#1a73e8', '#d93025', '#188038', '#a142f4', '#e37400', '#00897b'];

const CSS = `
.twb{--twb-line:#dfe4e6;--twb-head:#f1f4f5;--twb-sel:#1a73e8;--twb-comp:#f7faf9;--twb-drill:#3d8179;--twb-link:#c77b30;--twb-err:#d32f2f;
  font:13px system-ui,sans-serif;color:#1d2a30;display:flex;flex-direction:column;height:100%;outline:none;background:#fff;min-height:0;min-width:0;overflow:hidden}
.twb-bar{display:flex;gap:8px;align-items:center;padding:6px 8px;border-bottom:1px solid var(--twb-line);flex-wrap:wrap}
.twb-crumbs{display:flex;gap:4px;align-items:center;flex:1;min-width:0;overflow:hidden;white-space:nowrap}
.twb-crumbs button{border:0;background:none;color:var(--twb-drill);cursor:pointer;font:inherit;padding:2px 4px;border-radius:4px;max-width:260px;overflow:hidden;text-overflow:ellipsis}
.twb-crumbs button:hover{background:#eef5f4}.twb-crumbs button[aria-current]{color:#1d2a30;font-weight:600;cursor:default;background:none}
.twb-total{font-weight:700;font-variant-numeric:tabular-nums}
.twb-btn{border:1px solid var(--twb-line);background:#fff;border-radius:4px;padding:2px 8px;cursor:pointer;font:inherit}.twb-btn:disabled{opacity:.4;cursor:default}
.twb-fx{display:flex;border-bottom:1px solid var(--twb-line)}
.twb-fx .addr{width:64px;padding:4px 8px;border-right:1px solid var(--twb-line);color:#56646a;font-variant-numeric:tabular-nums;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.twb-fx input{flex:1;border:0;padding:4px 8px;font:inherit;outline:none;min-width:0}
.twb-fx input:focus{box-shadow:inset 0 -2px 0 var(--twb-sel)}
.twb-fx .hint{padding:4px 8px;color:#7a8a90;font-size:12px;white-space:nowrap}
.twb-fx .hint.away{color:#b35c00;font-weight:600}
.twb-scroll{flex:1;overflow:auto;min-height:0;position:relative}
.twb table{border-collapse:separate;border-spacing:0;table-layout:fixed}
.twb tr{height:${ROW_H}px}
.twb th,.twb td{border-right:1px solid var(--twb-line);border-bottom:1px solid var(--twb-line);padding:0 6px;overflow:hidden;white-space:nowrap;text-overflow:ellipsis;box-sizing:border-box}
.twb thead th{position:sticky;top:0;background:var(--twb-head);z-index:3;font-weight:600;text-align:left;font-size:12px;height:${ROW_H}px;user-select:none}
.twb thead th small{color:#98a6ab;font-weight:400;margin-right:4px}
.twb th.rn{position:sticky;left:0;background:var(--twb-head);z-index:2;text-align:right;color:#7a8a90;font-weight:400;font-size:12px;user-select:none;cursor:e-resize}
.twb thead th.rn{z-index:4;cursor:default}
.twb thead th:not(.rn){cursor:s-resize}
.twb .rz{position:absolute;right:0;top:0;width:6px;height:100%;cursor:col-resize;z-index:5}
.twb .rzr{position:absolute;left:0;bottom:0;height:5px;width:100%;cursor:row-resize;z-index:5}
.twb .rz:hover,.twb .rz.on{background:var(--twb-sel);opacity:.5}.twb .rzr:hover,.twb .rzr.on{background:var(--twb-sel);opacity:.5}
.twb thead th,.twb th.rn{position:sticky}
.twb td{cursor:cell;position:relative}
.twb tr.wrap td{white-space:normal;vertical-align:top;line-height:17px;padding-top:3px;text-overflow:clip}
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
.twb td.editing{padding:0;outline:2px solid var(--twb-sel);outline-offset:-2px;overflow:visible;z-index:2}
.twb td.editing input{position:absolute;left:0;top:0;min-width:100%;height:100%;border:0;padding:0 6px;font:inherit;outline:2px solid var(--twb-sel);outline-offset:-2px;box-sizing:border-box;background:#fff;text-align:left;box-shadow:0 2px 8px rgba(0,0,0,.15)}
.twb tfoot td{position:sticky;bottom:0;background:#eef2f3;font-weight:700;border-top:2px solid #9fb0b6}
.twb-hl{position:absolute;pointer-events:none;box-sizing:border-box;border:2px solid;z-index:1}
.twb-hl.pt{border-style:dashed;animation:twb-march .6s linear infinite}
@keyframes twb-march{50%{border-color:transparent}}
.twb-status{display:flex;gap:14px;padding:3px 8px;border-top:1px solid var(--twb-line);color:#56646a;font-size:12px;min-height:18px}
.twb-status .msg{color:#b35c00;margin-left:auto}
.twb-menu,.twb-ac,.twb-sig{position:fixed;z-index:1000;background:#fff;border:1px solid #cfd8dc;border-radius:6px;box-shadow:0 6px 24px rgba(0,0,0,.18);font:13px system-ui,sans-serif}
.twb-menu{padding:4px;min-width:220px}
.twb-menu button{display:block;width:100%;text-align:left;border:0;background:none;padding:6px 10px;border-radius:4px;cursor:pointer;font:inherit}
.twb-menu button:hover:not(:disabled){background:#eef5f4}.twb-menu button:disabled{color:#aab4b8;cursor:default}
.twb-menu hr{border:0;border-top:1px solid #eceff1;margin:4px 0}
.twb-ac{min-width:280px;max-width:420px;padding:4px 0}
.twb-ac .it{display:flex;gap:8px;padding:4px 10px;cursor:pointer;align-items:baseline}
.twb-ac .it b{font-weight:600}.twb-ac .it span{color:#7a8a90;font-size:12px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.twb-ac .it.on{background:#e8f0fe}
.twb-ac .desc{border-top:1px solid #eceff1;margin-top:4px;padding:6px 10px 2px;color:#56646a;font-size:12px}
.twb-ac .kind{margin-left:auto;font-size:10px;text-transform:uppercase;color:#98a6ab}
.twb-sig{padding:5px 10px;font-size:12px;color:#56646a;max-width:460px}
.twb-sig b{color:#1d2a30}.twb-sig .fn{color:#1d2a30;font-weight:600}
`;
let cssInjected = false;

const KIND_LABEL = { cost: 'cost', rate: 'rate build-up', quantity: 'quantity' };
const isFormulaText = (t) => typeof t === 'string' && t.trimStart().startsWith('=');

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
    /** active edit, or null — see _beginEdit */
    this.ed = null;
    this._clip = null;
    this._build(container);
    this._onChange = (e) => this._changed(e.detail);
    wb.addEventListener('change', this._onChange);
    this.render();
  }

  // ================================================================ public API

  /** Show a sheet. While a formula is being edited this keeps the edit open (cross-sheet pointing). */
  openSheet(sheetId, a1) {
    if (!this.wb.sheet(sheetId)) return;
    if (this.ed && !this._isFormulaEdit()) this._commitEdit({ stay: true });
    this.sheetId = sheetId;
    const p = a1 ? parseA1(a1) : { row: 1, col: 'A' };
    const ci = Math.max(0, this.wb.columnLetters(sheetId).indexOf(p.col));
    if (!this.ed) this.sel = { ar: p.row, ac: ci, r0: p.row, c0: ci, r1: p.row, c1: ci };
    if (this.ed) {
      // formula edit continues in the formula bar while we look at another sheet
      this.ed.point = null;
      this._moveEditorTo(this.ed.sheetId === sheetId ? 'cell' : 'fx');
    }
    this.render();
    if (!this.ed) this.root.focus();
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
  /** Begin editing the active cell programmatically (e.g. a toolbar "fx" button). */
  edit(initialText) {
    this._beginEdit({ replace: initialText !== undefined, text: initialText });
  }
  focus() {
    this.root.focus();
  }
  destroy() {
    this.wb.removeEventListener('change', this._onChange);
    document.removeEventListener('mousedown', this._docDown, true);
    this._menu?.remove();
    this._acEl?.remove();
    this._sigEl?.remove();
    this.root.remove();
  }

  // ================================================================ DOM

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
      <div class="twb-fx"><span class="addr"></span><input spellcheck="false" autocomplete="off" placeholder="value, text or =formula"/><span class="hint"></span></div>
      <div class="twb-scroll"><table><colgroup></colgroup><thead></thead><tbody></tbody><tfoot></tfoot></table></div>
      <div class="twb-status"><span class="sum"></span><span class="msg"></span></div>`;
    container.appendChild(r);
    this.$ = (s) => r.querySelector(s);
    this.table = this.$('table');
    this.tbody = this.$('tbody');
    this.scroll = this.$('.twb-scroll');
    this.fxInput = this.$('.twb-fx input');

    // buttons that must not steal focus from an open formula edit
    r.addEventListener('mousedown', (e) => {
      if (this.ed && e.target.closest('.twb-crumbs button, .open, [data-act]')) e.preventDefault();
    });
    r.addEventListener('click', (e) => {
      const act = e.target.closest('[data-act]')?.dataset.act;
      if (act === 'up') this._up();
      if (act === 'undo') this.wb.undo();
      if (act === 'redo') this.wb.redo();
      const crumb = e.target.closest('[data-sheet]');
      if (crumb) this.openSheet(crumb.dataset.sheet, crumb.dataset.a1);
      const open = e.target.closest('.open');
      if (open) this.openSheet(this.wb.sheet(this.sheetId).cells.get(open.closest('td').dataset.a1).drill);
    });
    this.table.addEventListener('mousedown', (e) => this._mouseDown(e));
    this.tbody.addEventListener('dblclick', (e) => {
      if (e.target.closest('td') && !this.ed) this._beginEdit({ replace: false });
    });
    this.$('thead').addEventListener('dblclick', (e) => {
      const rz = e.target.closest('.rz');
      if (rz) this._autofit(rz.dataset.col);
    });
    this.tbody.addEventListener('dblclick', (e) => {
      const rzr = e.target.closest('.rzr');
      if (rzr) this.wb.setRowHeight(this.sheetId, +rzr.dataset.row, null);
    });
    r.addEventListener('keydown', (e) => this._key(e));
    r.addEventListener('contextmenu', (e) => {
      const td = e.target.closest('tbody td');
      if (!td || this.ed) return;
      e.preventDefault();
      if (!this._inSel(+td.dataset.r, +td.dataset.c)) this._select(+td.dataset.r, +td.dataset.c);
      this._showMenu(e.clientX, e.clientY);
    });
    r.addEventListener('copy', (e) => this._copy(e));
    r.addEventListener('paste', (e) => this._paste(e));
    this.scroll.addEventListener('scroll', () => this._placePopups());

    // the formula bar is a second view of the same editor
    this.fxInput.addEventListener('focus', () => {
      if (!this.ed) this._beginEdit({ replace: false, via: 'fx' });
      else if (this.ed.input !== this.fxInput) this._moveEditorTo('fx');
    });
    this._wireInput(this.fxInput);

    // commit when focus leaves the grid entirely (clicking elsewhere on the page)
    this._docDown = (e) => {
      if (this._menu && !this._menu.contains(e.target)) this._hideMenu();
      if (!this.ed) return;
      const inside = this.root.contains(e.target) || this._acEl?.contains(e.target);
      if (!inside) this._commitEdit({ stay: true });
    };
    document.addEventListener('mousedown', this._docDown, true);
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

    this.$('thead').innerHTML =
      `<tr><th class="rn"></th>` +
      this._cols
        .map((c) => `<th data-col="${c}" title="${esc(defs[c].label)}${defs[c].computed ? ' (calculated)' : ''}"><small>${c}</small>${esc(defs[c].label)}<span class="rz" data-col="${c}"></span></th>`)
        .join('') +
      '</tr>';
    let html = '';
    for (let r = 1; r <= nRows; r++) {
      html += `<tr data-r="${r}"><th class="rn" data-row="${r}">${r}<span class="rzr" data-row="${r}"></span></th>`;
      for (let ci = 0; ci < this._cols.length; ci++) html += `<td data-r="${r}" data-c="${ci}" data-a1="${this._cols[ci]}${r}"></td>`;
      html += '</tr>';
    }
    this.tbody.innerHTML = html;
    const totalCol = this._cols.indexOf('H');
    this.$('tfoot').innerHTML =
      `<tr><th class="rn"></th>` +
      this._cols.map((c, i) => (i === totalCol - 1 ? `<td style="text-align:right">${s.parent ? '→ parent' : 'Total'}</td>` : i === totalCol ? `<td class="num tot"></td>` : '<td></td>')).join('') +
      '</tr>';
    this._applyLayout();
    this._renderCrumbs();
    if (this.ed && this.ed.sheetId === this.sheetId && this.ed.input !== this.fxInput) this._mountCellEditor();
    this._refresh();
  }

  /** Column widths + row heights from the workbook. */
  _applyLayout() {
    const wb = this.wb;
    this.$('colgroup').innerHTML = `<col style="width:44px">` + this._cols.map((c) => `<col data-col="${c}" style="width:${wb.columnWidth(this.sheetId, c)}px">`).join('');
    this.table.style.width = 44 + this._cols.reduce((a, c) => a + wb.columnWidth(this.sheetId, c), 0) + 'px';
    for (const tr of this.tbody.rows) {
      const h = wb.rowHeight(this.sheetId, +tr.dataset.r);
      tr.style.height = h ? `${h}px` : '';
      tr.classList.toggle('wrap', !!h && h > ROW_H);
    }
    this._highlight();
  }

  _renderCrumbs() {
    const path = this.wb.path(this.sheetId);
    this.$('.twb-crumbs').innerHTML = path
      .map((s, i) => {
        const cur = i === path.length - 1;
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
    const editTd = this.ed && this.ed.sheetId === this.sheetId && this.ed.input !== this.fxInput ? this.ed.td : null;
    for (const tr of this.tbody.rows) {
      const r = +tr.dataset.r;
      const rt = s.rowTypes.get(r);
      tr.classList.toggle('heading', rt === 'heading');
      tr.classList.toggle('subtotal', rt === 'subtotal');
      for (const td of tr.cells) {
        if (td.tagName !== 'TD' || td === editTd) continue;
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
        if (!this.ed && this._inSel(r, ci)) cls.push('insel');
        if (!this.ed && r === this.sel.ar && ci === this.sel.ac) cls.push('active');
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
    this.$('.twb-total').textContent = s.parent ? '' : `Total ${totText}`;
    this.$('[data-act="undo"]').disabled = !wb.canUndo;
    this.$('[data-act="redo"]').disabled = !wb.canRedo;
    if (!this.ed) this._renderFx();
    this._renderStatus();
    this._highlight();
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
    const h = this.$('.twb-fx .hint');
    h.textContent = hint;
    h.classList.remove('away');
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

  _changed(detail) {
    const s = this.wb.sheet(this.sheetId);
    if (!s) return this.render(); // sheet deleted (e.g. undo of a drill) → back to root
    if (this.ed && !this.wb.sheet(this.ed.sheetId)) this._cancelEdit();
    if (detail.keys === null) {
      const need = Math.max(this.opts.minRows, s.rowCount + this.opts.spareRows, this.sel.r1 + 1);
      if (need !== this._nRows) return this.render();
      this._renderCrumbs();
      this._applyLayout();
    } else if (detail.layout) this._applyLayout();
    this._refresh();
  }

  // ================================================================ selection

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
    this._cellEl(extend ? r : this.sel.ar, extend ? c : this.sel.ac)?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
    this.dispatchEvent(new CustomEvent('select', { detail: this.activeCell() }));
  }
  _cellEl(r, c) {
    return this.tbody.querySelector(`td[data-r="${r}"][data-c="${c}"]`);
  }

  _mouseDown(e) {
    if (e.button !== 0) return;
    // --- resizing
    const rz = e.target.closest('.rz');
    if (rz) return this._startColResize(e, rz);
    const rzr = e.target.closest('.rzr');
    if (rzr) return this._startRowResize(e, rzr);
    if (e.target.closest('.open')) return;

    // --- row / column header selection
    const rn = e.target.closest('tbody th.rn');
    const ch = e.target.closest('thead th[data-col]');
    if ((rn || ch) && !this.ed) {
      e.preventDefault();
      const last = Math.max(this.wb.sheet(this.sheetId).rowCount, 1);
      if (rn) this.sel = { ar: +rn.dataset.row, ac: 0, r0: +rn.dataset.row, c0: 0, r1: +rn.dataset.row, c1: this._cols.length - 1 };
      else {
        const ci = this._cols.indexOf(ch.dataset.col);
        this.sel = { ar: 1, ac: ci, r0: 1, c0: ci, r1: last, c1: ci };
      }
      this._refresh();
      this.root.focus();
      return;
    }

    const td = e.target.closest('tbody td');
    if (!td) return;
    const r = +td.dataset.r,
      c = +td.dataset.c;

    // --- formula editing: clicking a cell inserts a reference (point mode)
    if (this.ed) {
      if (this.ed.td === td && this.ed.input !== this.fxInput) return; // clicking inside the editor
      if (this._isFormulaEdit() && (this.ed.point || this._ctx().canInsertRef)) {
        e.preventDefault(); // keep focus (and the caret) in the editor
        this._pointAt(r, c, r, c, true);
        const move = (ev) => {
          const t = document.elementFromPoint(ev.clientX, ev.clientY)?.closest?.('tbody td');
          if (t && this.ed?.point) this._pointAt(this.ed.point.anchor.r, this.ed.point.anchor.c, +t.dataset.r, +t.dataset.c, false);
        };
        const up = () => {
          removeEventListener('mousemove', move);
          removeEventListener('mouseup', up);
        };
        addEventListener('mousemove', move);
        addEventListener('mouseup', up);
        return;
      }
      this._commitEdit({ stay: true });
    }

    this._select(r, c, e.shiftKey);
    const move = (ev) => {
      const t = document.elementFromPoint(ev.clientX, ev.clientY)?.closest?.('tbody td');
      if (t) {
        const r2 = +t.dataset.r,
          c2 = +t.dataset.c;
        if (r2 !== this.sel.r1 || c2 !== this.sel.c1) {
          this.sel = { ...this.sel, r1: r2, c1: c2 };
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

  // ================================================================ keyboard (not editing)

  _key(e) {
    if (this.ed) return; // the editor input handles its own keys
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
        return this._beginEdit({ replace: false });
      case 'F2':
        e.preventDefault();
        return this._beginEdit({ replace: false });
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
      this._beginEdit({ replace: true, text: e.key });
    }
  }

  // ================================================================ editor

  /**
   * this.ed = {
   *   sheetId, a1, r, c      the cell being edited (its "home")
   *   mode                   'enter' (typed over: arrows point / commit) | 'edit' (F2: arrows move the caret)
   *   wasLink                editing a linked cell's modifier
   *   input                  the active <input>: in-cell, or the formula bar
   *   td                     the home cell element (when mounted)
   *   text                   current text
   *   point                  { start, end, anchor:{r,c}, cur:{r,c}, sheetId } — the reference being pointed at
   * }
   */
  _beginEdit({ replace, text, via }) {
    const { a1, row } = this.activeCell();
    const info = this.wb.info(this.sheetId, a1);
    if (info.type === 'drill') {
      if (replace) return this._flash('This figure comes from its breakdown — Ctrl+Enter to open, or right-click → Remove breakdown');
      return this.openSheet(info.sheetId);
    }
    if (info.type === 'subtotal') return this._flash('Subtotal rows add up the rows above automatically');
    const c = this.wb.sheet(this.sheetId).cells.get(a1);
    const initial = replace ? (text ?? '') : info.type === 'link' ? (c.link.expr ?? '=link') : (c?.raw ?? '');
    this.ed = { sheetId: this.sheetId, a1, r: row, c: this.sel.ac, mode: replace ? 'enter' : 'edit', wasLink: info.type === 'link' && !replace, text: initial, point: null, input: null, td: null };
    this.fxInput.disabled = false;
    if (via === 'fx') {
      this.ed.input = this.fxInput;
      this.fxInput.value = initial;
    } else this._mountCellEditor();
    this._refresh();
    this._updateAssist();
  }

  _mountCellEditor() {
    const ed = this.ed;
    const td = this.tbody.querySelector(`td[data-a1="${ed.a1}"]`);
    if (!td) return;
    const input = document.createElement('input');
    input.spellcheck = false;
    input.autocomplete = 'off';
    input.value = ed.text;
    td.className = 'editing';
    td.textContent = '';
    td.appendChild(input);
    ed.td = td;
    ed.input = input;
    this._wireInput(input);
    this.fxInput.value = ed.text;
    input.focus();
    input.setSelectionRange(input.value.length, input.value.length);
    this._grow();
  }

  /** Switch the active editor between the cell and the formula bar, keeping text + caret. */
  _moveEditorTo(where) {
    const ed = this.ed;
    if (!ed) return;
    const caret = ed.input?.selectionStart ?? ed.text.length;
    if (where === 'fx') {
      ed.input = this.fxInput;
      ed.td = null;
      this.fxInput.value = ed.text;
      this.fxInput.focus();
      this.fxInput.setSelectionRange(caret, caret);
      const h = this.$('.twb-fx .hint');
      const home = this.wb.sheet(ed.sheetId);
      if (ed.sheetId !== this.sheetId) {
        h.textContent = `editing ${home.name} › ${ed.a1} — click cells to reference them, Enter to finish`;
        h.classList.add('away');
      }
    } else {
      ed.input = null; // render() mounts the in-cell editor
    }
  }

  /** Like Excel, the in-cell editor grows over the neighbouring cells as you type. */
  _grow() {
    const ed = this.ed;
    if (!ed?.td || ed.input === this.fxInput) return;
    const max = this.scroll.clientWidth - (ed.td.offsetLeft - this.scroll.scrollLeft) - 8;
    ed.input.style.width = `${Math.max(ed.td.clientWidth, Math.min(max, measure(ed.text) + 24))}px`;
  }

  _wireInput(input) {
    input.addEventListener('input', () => {
      if (!this.ed || this.ed.input !== input) return;
      this.ed.text = input.value;
      this.ed.point = null; // typing ends pointing
      if (input !== this.fxInput) this.fxInput.value = input.value;
      else if (this.ed.td?.firstChild && this.ed.td.firstChild !== input) this.ed.td.firstChild.value = input.value;
      this._grow();
      this._updateAssist();
      this._highlight();
    });
    input.addEventListener('keydown', (e) => this._editKey(e, input));
    input.addEventListener('keyup', (e) => {
      if (/^(Arrow|Home|End)/.test(e.key) && this.ed && !this.ed.point) this._updateAssist();
    });
    input.addEventListener('click', () => this.ed && this._updateAssist());
  }

  _isFormulaEdit() {
    return !!this.ed && isFormulaText(this.ed.text);
  }
  _ctx() {
    const ed = this.ed;
    const caret = ed.input?.selectionStart ?? ed.text.length;
    return caretContext(ed.text, caret);
  }

  _editKey(e, input) {
    const ed = this.ed;
    if (!ed || ed.input !== input) return;
    e.stopPropagation();
    const ac = this._ac;

    // --- autocomplete list has priority
    if (ac?.items.length) {
      if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
        e.preventDefault();
        ac.index = (ac.index + (e.key === 'ArrowDown' ? 1 : -1) + ac.items.length) % ac.items.length;
        return this._renderAc();
      }
      if (e.key === 'Tab' || (e.key === 'Enter' && !e.ctrlKey)) {
        e.preventDefault();
        return this._acceptAc(ac.items[ac.index]);
      }
      if (e.key === 'Escape') {
        e.preventDefault();
        return this._closeAc();
      }
    }

    const arrows = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
    if (arrows[e.key]) {
      const [dr, dc] = arrows[e.key];
      // point mode: arrows move (Shift: extend) the reference being inserted
      if (this._isFormulaEdit() && (ed.point || (ed.mode === 'enter' && this._ctx().canInsertRef))) {
        e.preventDefault();
        const p = ed.point;
        const onHome = this.sheetId === ed.sheetId;
        const base = p && p.sheetId === this.sheetId ? p.cur : onHome ? { r: ed.r, c: ed.c } : { r: this.sel.ar, c: this.sel.ac };
        const cur = { r: Math.max(1, base.r + dr), c: Math.max(0, Math.min(this._cols.length - 1, base.c + dc)) };
        if (e.shiftKey && p) this._pointAt(p.anchor.r, p.anchor.c, cur.r, cur.c, false);
        else this._pointAt(cur.r, cur.c, cur.r, cur.c, !p);
        return;
      }
      if (ed.mode === 'enter' && input !== this.fxInput) {
        // typed-over value: arrows commit and move, like Excel
        e.preventDefault();
        this._commitEdit({ move: [dr, dc] });
        return;
      }
      return; // edit mode: let the caret move
    }
    switch (e.key) {
      case 'Enter':
        e.preventDefault();
        return this._commitEdit({ move: [e.shiftKey ? -1 : 1, 0] });
      case 'Tab':
        e.preventDefault();
        return this._commitEdit({ move: [0, e.shiftKey ? -1 : 1] });
      case 'Escape':
        e.preventDefault();
        return this._cancelEdit();
      case 'F2':
        e.preventDefault();
        ed.mode = ed.mode === 'enter' ? 'edit' : 'enter';
        ed.point = null;
        this._flash(ed.mode === 'enter' ? 'Point mode — arrows pick cells' : 'Edit mode — arrows move the cursor');
        return;
    }
  }

  /**
   * Insert (or replace the currently pointed) reference for r0,c0 : r1,c1 on
   * the DISPLAYED sheet. Adds a [Sheet]! prefix when it's not the home sheet.
   */
  _pointAt(r0, c0, r1, c1, fresh) {
    const ed = this.ed;
    const input = ed.input;
    const a = `${this._cols[c0]}${r0}`;
    const b = `${this._cols[c1]}${r1}`;
    let ref = a === b ? a : `${this._cols[Math.min(c0, c1)]}${Math.min(r0, r1)}:${this._cols[Math.max(c0, c1)]}${Math.max(r0, r1)}`;
    if (this.sheetId !== ed.sheetId) ref = `[${this.wb.sheet(this.sheetId).name}]!${ref}`;
    let start, end;
    if (ed.point && !fresh) {
      start = ed.point.start;
      end = ed.point.end;
    } else if (ed.point && fresh) {
      // clicking again right after a pointed ref replaces it (Excel behaviour)
      start = ed.point.start;
      end = ed.point.end;
    } else {
      start = end = input.selectionStart ?? ed.text.length;
      // if the caret sits right after an operator with trailing spaces, keep them
    }
    ed.text = ed.text.slice(0, start) + ref + ed.text.slice(end);
    ed.point = { start, end: start + ref.length, anchor: { r: r0, c: c0 }, cur: { r: r1, c: c1 }, sheetId: this.sheetId };
    input.value = ed.text;
    if (input !== this.fxInput) this.fxInput.value = ed.text;
    else if (ed.td?.firstChild) ed.td.firstChild.value = ed.text;
    input.focus();
    input.setSelectionRange(ed.point.end, ed.point.end);
    this._grow();
    this._closeAc();
    this._updateAssist();
    this._highlight();
    this._cellEl(r1, c1)?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  }

  _commitEdit({ move, stay } = {}) {
    const ed = this.ed;
    if (!ed) return;
    this.ed = null;
    this._closeAc();
    const v = ed.text;
    const home = ed.sheetId;
    const sheetChanged = this.sheetId !== home;
    try {
      if (ed.wasLink && /^\s*=.*\blink\b/i.test(v)) this.wb.setLinkExpr(home, ed.a1, v.trim() === '=link' ? '' : v);
      else {
        const cur = this.wb.sheet(home)?.cells.get(ed.a1);
        if ((cur?.raw ?? '') !== v || cur?.link) this.wb.setCell(home, ed.a1, v);
      }
    } catch (err) {
      this._flash(err.message);
    }
    if (sheetChanged && !stay) {
      // finished a cross-sheet formula: go back to the cell it belongs to
      this.sheetId = home;
      this.sel = { ar: ed.r, ac: ed.c, r0: ed.r, c0: ed.c, r1: ed.r, c1: ed.c };
      this.render();
    } else if (!sheetChanged) {
      this.sel = { ar: ed.r, ac: ed.c, r0: ed.r, c0: ed.c, r1: ed.r, c1: ed.c };
    }
    if (move && !sheetChanged) this._select(ed.r + move[0], ed.c + move[1]);
    else this._refresh();
    if (!stay) this.root.focus();
  }

  _cancelEdit() {
    const ed = this.ed;
    if (!ed) return;
    this.ed = null;
    this._closeAc();
    if (this.sheetId !== ed.sheetId && this.wb.sheet(ed.sheetId)) {
      this.sheetId = ed.sheetId;
      this.sel = { ar: ed.r, ac: ed.c, r0: ed.r, c0: ed.c, r1: ed.r, c1: ed.c };
      this.render();
    } else this._refresh();
    this.root.focus();
  }

  // ================================================================ reference highlighting

  /** Outline each reference of the formula being edited, in its own colour (Excel-style). */
  _highlight() {
    for (const el of this.scroll.querySelectorAll('.twb-hl')) el.remove();
    const ed = this.ed;
    if (!ed || !isFormulaText(ed.text)) return;
    let toks;
    try {
      toks = tokenize(ed.text.trimStart().slice(1));
    } catch {
      return;
    }
    const shown = this.wb.sheet(this.sheetId);
    let colour = 0;
    const offset = ed.text.length - ed.text.trimStart().length + 1;
    for (let i = 0; i < toks.length; i++) {
      const t = toks[i];
      if (t.t !== 'ref') continue;
      const prefix = toks[i - 1]?.t === 'sheet' && toks[i - 1].e === t.s ? toks[i - 1].v : null;
      let a = t.v,
        b = t.v,
        end = t.e;
      if (toks[i + 1]?.t === ':' && toks[i + 2]?.t === 'ref') {
        b = toks[i + 2].v;
        end = toks[i + 2].e;
        i += 2;
      }
      const target = prefix ? this.wb.sheetByName(prefix) : this.wb.sheet(ed.sheetId);
      const col = HL_COLOURS[colour++ % HL_COLOURS.length];
      if (target !== shown) continue;
      const pointed = ed.point && ed.point.end === end + offset && ed.point.sheetId === this.sheetId;
      this._drawBox(a, b, col, pointed);
    }
  }
  _drawBox(a, b, colour, pointed) {
    const c0 = Math.min(colToIndex(a.col), colToIndex(b.col)),
      c1 = Math.max(colToIndex(a.col), colToIndex(b.col));
    const r0 = Math.min(a.row, b.row),
      r1 = Math.max(a.row, b.row);
    const ci0 = this._cols.indexOf(indexToCol(c0)),
      ci1 = this._cols.indexOf(indexToCol(c1));
    if (ci0 < 0 || ci1 < 0) return;
    const tl = this._cellEl(r0, ci0),
      br = this._cellEl(Math.min(r1, this._nRows), ci1);
    if (!tl || !br) return;
    const box = document.createElement('div');
    box.className = 'twb-hl' + (pointed ? ' pt' : '');
    Object.assign(box.style, {
      left: `${tl.offsetLeft}px`,
      top: `${tl.offsetTop}px`, // offsets are relative to the table, which sits at the scroll origin
      width: `${br.offsetLeft + br.offsetWidth - tl.offsetLeft}px`,
      height: `${br.offsetTop + br.offsetHeight - tl.offsetTop}px`,
      borderColor: colour,
      background: colour + '14',
    });
    this.scroll.appendChild(box);
  }

  // ================================================================ autocomplete + argument hint

  _updateAssist() {
    const ed = this.ed;
    if (!ed || !isFormulaText(ed.text) || !ed.input) {
      this._closeAc();
      this._hideSig();
      return;
    }
    const ctx = this._ctx();
    const items = [];
    if (ctx.measPrefix !== null) {
      const q = ctx.measPrefix.toLowerCase();
      for (const [id, m] of this.wb.measurements)
        if ((m.name ?? id).toLowerCase().includes(q)) items.push({ kind: 'measurement', label: m.name ?? id, detail: `${fmt(m.value)} ${m.unit ?? ''}`, replace: ctx.measPrefix.length, insert: `${m.name ?? id}")` });
    } else if (ctx.word && !/^[A-Za-z]{1,3}\d+$/.test(ctx.word.text)) {
      const w = ctx.word.text.toUpperCase();
      for (const f of FUNCTIONS) if (f.name.startsWith(w)) items.push({ kind: 'function', label: f.name, detail: `(${f.args.join(', ')})`, desc: f.desc, replace: ctx.word.text.length, insert: f.args.length ? `${f.name}(` : `${f.name}()` });
      for (const n of this.wb.names.values())
        if (n.name.toUpperCase().startsWith(w)) items.push({ kind: 'name', label: n.name, detail: `= ${n.raw}`, replace: ctx.word.text.length, insert: n.name });
      if (ed.wasLink && 'LINK'.startsWith(w)) items.push({ kind: 'name', label: 'link', detail: 'the measured value', replace: ctx.word.text.length, insert: 'link' });
      // a word that already exactly matches a single item isn't worth a popup
      if (items.length === 1 && items[0].label.toUpperCase() === w && items[0].kind !== 'function') items.length = 0;
    }
    if (items.length) {
      const prev = this._ac;
      this._ac = { items: items.slice(0, 12), index: prev && prev.items[prev.index] ? Math.min(prev.index, items.length - 1) : 0 };
      this._hideSig();
      this._renderAc();
    } else {
      this._closeAc();
      if (ctx.call && FUNCTION_INFO[ctx.call.name]) this._showSig(FUNCTION_INFO[ctx.call.name], ctx.call.arg);
      else this._hideSig();
    }
  }

  _renderAc() {
    const ac = this._ac;
    if (!this._acEl) {
      this._acEl = document.createElement('div');
      this._acEl.className = 'twb-ac';
      // mousedown, not click, and prevent default → the editor keeps focus
      this._acEl.addEventListener('mousedown', (e) => {
        e.preventDefault();
        const it = e.target.closest('.it');
        if (it) this._acceptAc(this._ac.items[+it.dataset.i]);
      });
      document.body.appendChild(this._acEl);
    }
    const cur = ac.items[ac.index];
    this._acEl.innerHTML =
      ac.items.map((it, i) => `<div class="it${i === ac.index ? ' on' : ''}" data-i="${i}"><b>${esc(it.label)}</b><span>${esc(it.detail ?? '')}</span><span class="kind">${it.kind}</span></div>`).join('') +
      (cur?.desc ? `<div class="desc">${esc(cur.desc)}</div>` : '');
    this._acEl.style.display = '';
    this._placePopups();
  }

  _acceptAc(item) {
    const ed = this.ed;
    if (!ed || !item) return;
    const input = ed.input;
    const caret = input.selectionStart ?? ed.text.length;
    const start = caret - item.replace;
    ed.text = ed.text.slice(0, start) + item.insert + ed.text.slice(caret);
    input.value = ed.text;
    if (input !== this.fxInput) this.fxInput.value = ed.text;
    const pos = start + item.insert.length - (item.insert.endsWith('()') ? 1 : 0);
    input.setSelectionRange(pos, pos);
    this._grow();
    ed.point = null;
    ed.mode = 'enter'; // after "SUM(" a click or arrow should point at cells
    this._closeAc();
    this._updateAssist();
    this._highlight();
  }

  _closeAc() {
    this._ac = null;
    if (this._acEl) this._acEl.style.display = 'none';
  }

  _showSig(f, argIndex) {
    if (!this._sigEl) {
      this._sigEl = document.createElement('div');
      this._sigEl.className = 'twb-sig';
      document.body.appendChild(this._sigEl);
    }
    const repeating = f.args[f.args.length - 1] === '…';
    const idx = repeating && argIndex >= f.args.length - 1 ? f.args.length - 2 : argIndex;
    const args = f.args.map((a, i) => (i === idx ? `<b>${esc(a)}</b>` : esc(a))).join(', ');
    this._sigEl.innerHTML = `<span class="fn">${f.name}</span>(${args})<br>${esc(f.desc)}`;
    this._sigEl.style.display = '';
    this._placePopups();
  }
  _hideSig() {
    if (this._sigEl) this._sigEl.style.display = 'none';
  }

  /** Position popups under the active editor (cell or formula bar). */
  _placePopups() {
    const input = this.ed?.input;
    if (!input) return;
    const r = input.getBoundingClientRect();
    const caretX = r.left + Math.min(r.width - 20, measure(input.value.slice(0, input.selectionStart ?? 0)) + 6);
    for (const el of [this._acEl, this._sigEl]) {
      if (!el || el.style.display === 'none') continue;
      const h = el.offsetHeight;
      const below = r.bottom + 2;
      el.style.left = `${Math.max(4, Math.min(caretX - 12, innerWidth - el.offsetWidth - 4))}px`;
      el.style.top = `${below + h > innerHeight ? r.top - h - 2 : below}px`;
    }
  }

  // ================================================================ resizing

  _startColResize(e, handle) {
    e.preventDefault();
    e.stopPropagation();
    const col = handle.dataset.col;
    const colEl = this.$(`col[data-col="${col}"]`);
    const startX = e.clientX;
    const startW = this.wb.columnWidth(this.sheetId, col);
    const baseTable = this.table.offsetWidth - startW;
    handle.classList.add('on');
    let w = startW;
    const move = (ev) => {
      w = Math.max(24, startW + ev.clientX - startX);
      colEl.style.width = `${w}px`; // live preview; the model is updated once on release
      this.table.style.width = `${baseTable + w}px`;
      this._highlight();
    };
    const up = () => {
      removeEventListener('mousemove', move);
      removeEventListener('mouseup', up);
      handle.classList.remove('on');
      if (Math.round(w) !== startW) this.wb.setColumnWidth(this.sheetId, col, w);
    };
    addEventListener('mousemove', move);
    addEventListener('mouseup', up);
  }

  _startRowResize(e, handle) {
    e.preventDefault();
    e.stopPropagation();
    const row = +handle.dataset.row;
    const tr = this.tbody.querySelector(`tr[data-r="${row}"]`);
    const startY = e.clientY;
    const startH = tr.offsetHeight;
    handle.classList.add('on');
    let h = startH;
    const move = (ev) => {
      h = Math.max(16, startH + ev.clientY - startY);
      tr.style.height = `${h}px`;
      tr.classList.toggle('wrap', h > ROW_H);
      this._highlight();
    };
    const up = () => {
      removeEventListener('mousemove', move);
      removeEventListener('mouseup', up);
      handle.classList.remove('on');
      if (Math.round(h) !== startH) this.wb.setRowHeight(this.sheetId, row, h);
    };
    addEventListener('mousemove', move);
    addEventListener('mouseup', up);
  }

  /** Double-click a column edge: fit the widest content on this sheet (and the header). */
  _autofit(col) {
    const s = this.wb.sheet(this.sheetId);
    const def = this._defs[col];
    let w = measure(def.label, '600 12px system-ui,sans-serif') + measure(col, '12px system-ui,sans-serif') + 20;
    for (let r = 1; r <= s.rowCount; r++) {
      const a1 = `${col}${r}`;
      const t = this.wb.display(this.sheetId, a1);
      if (!t) continue;
      const bold = s.rowTypes.get(r) === 'heading' || s.rowTypes.get(r) === 'subtotal' || this.wb.info(this.sheetId, a1).type === 'drill';
      w = Math.max(w, measure(t, `${bold ? '600 ' : ''}13px system-ui,sans-serif`) + 16 + (s.cells.get(a1)?.drill ? 14 : 0) + (s.cells.get(a1)?.link ? 10 : 0));
    }
    this.wb.setColumnWidth(this.sheetId, col, Math.ceil(Math.min(800, w)));
  }

  // ================================================================ edits, drill, clipboard, menu

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

  _copy(e) {
    if (this.ed) return;
    e.preventDefault();
    const { r0, r1, c0, c1 } = this._rect();
    this._clip = this.wb.copy(this.sheetId, r0, c0, r1, c1);
    e.clipboardData.setData('text/plain', this._clip.text);
    this._flash(`Copied ${r1 - r0 + 1}×${c1 - c0 + 1}`);
  }
  _paste(e) {
    if (this.ed) return;
    e.preventDefault();
    const text = e.clipboardData.getData('text/plain');
    const { ar, ac } = this.sel;
    try {
      if (this._clip && norm(text) === norm(this._clip.text)) this.wb.paste(this.sheetId, ar, ac, this._clip);
      else if (text) this.wb.pasteText(this.sheetId, ar, ac, text);
    } catch (err) {
      this._flash(err.message);
    }
  }

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
      ['Edit link modifier…', () => this._beginEdit({ replace: false }), 'F2', info.type !== 'link'],
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
      ['Auto-fit column width', () => this._autofit(col)],
      ['Reset column width', () => this.wb.setColumnWidth(this.sheetId, col, null)],
      ['Reset row height', () => this.wb.setRowHeight(this.sheetId, row, null), '', !this.wb.rowHeight(this.sheetId, row)],
      '-',
      [
        'Add / edit note…',
        () => {
          const n = prompt('Note for ' + a1, this.wb.sheet(this.sheetId).cells.get(a1)?.note ?? '');
          if (n !== null) this.wb.setNote(this.sheetId, a1, n);
        },
      ],
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
        if (!this.ed) this.root.focus();
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
const fmt = (v) => (typeof v === 'number' ? +v.toFixed(3) : (v ?? ''));
let measureCtx;
/** Text width in px (canvas measureText) — for auto-fit and caret-anchored popups. */
function measure(text, font = '13px system-ui,sans-serif') {
  measureCtx ??= document.createElement('canvas').getContext('2d');
  measureCtx.font = font;
  return measureCtx.measureText(text).width;
}
