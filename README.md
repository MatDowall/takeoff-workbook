# takeoff-workbook

A drill-down estimating workbook for takeoff apps: CostX-style **Cost / Quantity / Rate** sheets in a tree, **live links to measurements**, and a calculation engine that **never goes stale**. Framework-free JavaScript, with a grid UI included.

Designed to pair with [`pdf-viewport`](https://github.com/MatDowall/pdf-viewport) and [`pdf-snap`](https://github.com/MatDowall/pdf-snap), but it has no dependency on them.

```
Estimate                                    Total 18,706.57
  1.01  Excavate to reduced level   ▸ 60.00 m3 × 38.50 = 2,310.00
          └─ Excavate… — Quantity          (2 × 10 × 3 = 60)
  1.02  Concrete slab 100mm        ⛓ 120.00 m2 × 96.00 …        ← linked to a measurement
        Substructure subtotal                       17,446.20    ← auto subtotal row
```

---

## Quick start

```js
import { Workbook, WorkbookView } from 'takeoff-workbook';

const wb = new Workbook({ userColumns: 4, userColumnLabels: ['Trade', 'Notes'] });
const view = new WorkbookView(document.getElementById('wb'), wb, {
  onRequestLink: (sheetId, a1) => view.linkActive(selectedMeasurementId),   // right-click → Link measurement…
});

// from your takeoff tool, whenever a measurement is created or changes:
wb.setMeasurement('m17', { value: 42.6, unit: 'm2', name: 'Ground floor slab' });

// save / load: one JSON document for the whole tree
localStorage.estimate = JSON.stringify(wb.toJSON());
const again = Workbook.fromJSON(JSON.parse(localStorage.estimate));
```

**Demo:** put `pdf-viewport`, `pdf-snap` and `takeoff-workbook` side by side. Run `npx.cmd serve .` in the parent folder and open `/takeoff-workbook/demo/`. Measure on the PDF (length or area, with snapping), select a measurement and a Quantity cell, then click **Link**. Change the drawing scale and watch the estimate update. The workbook side also runs without a PDF.

### Keys (grid)

| | |
|---|---|
| arrows / Tab / Enter | move (Shift = extend selection or move back) |
| type | replace the cell · **F2** / double-click to edit |
| **Ctrl+Enter** | **drill down** (or open an existing breakdown) · ▸ in the cell also opens it |
| **Alt+↑** | up a level · the breadcrumb jumps to any level |
| Delete | clear (breakdowns are protected) |
| Ctrl+Z / Ctrl+Y | undo / redo (500 steps, including drills and row ops) |
| Ctrl+C / Ctrl+V | internal copy keeps breakdowns, links and relative formulas. Text from Excel or Sheets pastes as values and formulas. |
| right-click | drill / remove breakdown, link / unlink / edit modifier, insert / delete rows, heading / subtotal rows, notes |

The status bar shows the Sum and Count of the selection.

---

## The model

### Rows are "quantity × rate", not free-form

| Sheet | A | B | C | D | E | F | G | H |
|---|---|---|---|---|---|---|---|---|
| **Cost** (estimate, cost breakdowns) | Code | Description | Quantity | Unit | Rate | Subtotal **= C×E** | Factor | Total **= F×G** |
| **Rate** build-up | Code | Resource | Usage | Unit | Rate | Subtotal **= C×E** | Factor | Rate **= F×G** |
| **Quantity** (dimension lines) | Code | Description | Number | Length | Width | Depth | Factor | Quantity **= product of filled C–F × G** |

Columns I onward are free user columns (default 6; any number, e.g. 30 like CostX). A blank Factor counts as 1.

- **Overrides:** type into a calculated cell (Subtotal/Total/Quantity) to override it. It shows in italics with a purple corner. Clear it to go back to the calculation.
- **Heading rows** are bold and excluded from totals. **Subtotal rows** add up the rows above (back to the previous subtotal) and aren't double-counted in the sheet total.

### Drill-down: the same three types as CostX

| Drill from | Creates | Its total becomes |
|---|---|---|
| **Quantity** (C; or C–F on a quantity sheet) | Quantity sheet of dimension lines | the quantity |
| **Rate** (E) | Rate build-up (labour, plant, materials…) | the rate |
| **Total** (H) | Cost breakdown | the line total |

Drill-downs nest to any depth, e.g. a length inside a quantity inside an estimate line. When you drill an existing figure, it's **carried into row 1** of the new sheet, so the parent doesn't change until you edit. "Remove breakdown" keeps the figure as a typed value.

Each breakdown belongs to **exactly one cell**. Copying a drilled cell **deep-copies** the breakdown, so two lines can never share one and be "priced twice" (a warning state Nezo needs).

### Measurement links

```js
wb.link(sheetId, 'C4', 'm17');               // the cell shows the measured value
wb.setLinkExpr(sheetId, 'C4', '=link*1.05'); // optional modifier: +5% waste
wb.setMeasurement('m17', { value: 44.1 });   // re-measure → cell, row, sheet, parents, estimate update
wb.removeMeasurement('m17');                 // the cell shows #REF! (never silently 0) and so does the total
wb.linksTo('m17');                           // [{sheetId, a1}] e.g. to show "used in 2 places"
```

Any formula can also read a measurement by name: `=MEAS("Ground floor slab")/2`. Linking into an empty Unit column fills the unit from the measurement. Measurements belong to your app. The workbook stores the links, plus a snapshot of the last known values in its JSON so a reopened estimate shows the right numbers straight away.

### Formulas

`+ − * / ^ %` · `&` (join text) · `= <> < > <= >=` · refs `C3 $C$3 c3` · ranges `H1:H20` · other sheets `[Sheet name]!H5` · workbook names (`wb.setName('wastage', '1.05')` → `=C3*wastage`) · `link`

Functions: `SUM AVERAGE MIN MAX COUNT ROUND ROUNDUP ROUNDDOWN CEILING FLOOR ABS SQRT PI IF AND OR NOT IFERROR MEAS`

Errors are values that propagate (like Excel): `#REF!` `#DIV/0!` `#NAME?` `#VALUE!` `#CIRC!` (circular, e.g. a breakdown referring to the total it feeds) `#ERROR!` (syntax). Hover a cell for an explanation. An error below **surfaces in the totals above it**, so a broken breakdown can't quietly understate the estimate.

**Inserting or deleting rows** rewrites references everywhere, across sheets, in link modifiers and in names, exactly like Excel. Ranges grow and shrink. References to deleted rows become `#REF!`. Breakdowns move with their row, and are deleted (undoably) with it. Renaming a sheet updates every `[Name]!` reference to it.

---

## How the engine works (and why there's no Recalculate button)

Nezo copies each breakdown's total into the parent cell (`rolled`). Its UI then needs warnings like "This total disagrees with its breakdown, press Recalculate", "This breakdown no longer exists", and "priced twice" whenever the copy and the source disagree.

Here **nothing derived is stored**:

1. Every value (cell, sheet total, name) is computed on demand and memoised.
2. While computing, the engine records what each value **read**: other cells, sheet totals, measurements, names, a sheet's row structure. These dependencies are discovered at runtime (like a reactive system), so `IF()`, ranges and cross-sheet references need no static analysis.
3. When something changes, the engine invalidates it and everything that transitively read it. The path runs from measurement to cell, row, sheet total, the parent's drilled cell, and on up to the estimate total. The next read recomputes only those values.
4. A value can't depend on itself: if it's met again while still being calculated, it returns `#CIRC!`.

Structural edits (drill, row insert/delete, undo) clear the memo and recompute lazily, which takes milliseconds.

**Timings** (Node, this machine): 5,000-row estimate with 1,000 breakdowns: full calculation 265 ms, and an edit deep in a breakdown flows through to a new estimate total in **~10 ms**. Typical sheets (hundreds of rows) are under 1 ms.

**Undo** is transactional: the first time an edit touches a sheet, the engine snapshots it. A step can span several sheets (a drill creates one; deleting a row removes a subtree). `wb.batch(label, fn)` groups edits into one undo step. A failed edit rolls back.

**Persistence** is one JSON document for the whole tree (`format: 'takeoff-workbook'`). Nezo stores workbooks as separate documents, which is how totals and children drift apart. For a multi-user server, store the document per estimate, or the sheets as rows. Nothing derived needs saving, so there's nothing to keep in sync.

---

## API summary

**Workbook**

| | |
|---|---|
| `value(sheetId, a1)` / `display(sheetId, a1)` / `total(sheetId)` | live values |
| `info(sheetId, a1)` | `{type: drill \| link \| override \| computed \| formula \| value \| subtotal \| empty, …}` |
| `setCell` / `setCells(sheetId, [{a1, raw}])` | edit (`''` clears; a value on a calculated column overrides it) |
| `drill(sheetId, a1)` → childId · `removeDrill(…, {keepValue})` | drill-downs |
| `link` / `setLinkExpr` / `unlink` · `setMeasurement` / `removeMeasurement` / `linksTo` | measurements |
| `insertRows(sheetId, at, n)` · `deleteRows(…)` · `setRowType(sheetId, row, 'normal'\|'heading'\|'subtotal')` | structure |
| `copy(sheetId, r0, c0, r1, c1)` → clip · `paste(sheetId, row, colIndex, clip)` · `pasteText(…, tsv)` | clipboard |
| `renameSheet` · `setName` · `setNote` | |
| `undo()` / `redo()` / `batch(label, fn)` · `canUndo` / `canRedo` | history |
| `root` / `sheet(id)` / `path(id)` / `children(id)` / `columns(id)` | navigation |
| `toJSON()` / `Workbook.fromJSON(json)` | persistence |
| `'change'` event | `detail.keys` = the invalidated keys, or `null` for structural changes |

**WorkbookView** (`new WorkbookView(el, wb, { onRequestLink, minRows, spareRows })`): `openSheet(id, a1?)`, `activeCell()`, `linkActive(measurementId, expr?)`, `focus()`, `destroy()`. It fires `'select'` events.

---

## Notes and limits

- The grid is a plain `<table>`, re-filled from the memoised engine on each change (about 1 ms for hundreds of rows). Add row virtualisation if sheets ever reach tens of thousands of rows.
- There's one formula language, Excel-like but small; adding a function is one entry in the `FN` table in `formula.js`. Names must not look like cell refs (`gst15` would be read as a reference).
- Collaborative editing (cell locks, like Nezo's) isn't included. The transaction log in `_transact` is the natural hook for sending edits to a server.
- There's no styling, number-format UI or column resizing yet. Column widths and decimal places are in `columns.js`.

## Tests

```
npm test              # formula engine + workbook model (19 tests incl. 5,000-row perf)
npm run test:browser  # headless Chrome on the demo: typing, Ctrl+Enter drill, Alt+↑, undo/redo,
                      # PDF measure → link → change scale → live totals, F2 link modifier, deleted measurement → #REF!
```

## Files

```
src/formula.js       tokenizer, Pratt parser, evaluator, reference rewriting
src/columns.js       sheet kinds & column roles (edit labels/widths/decimals here)
src/Workbook.js      model, dependency-tracking engine, drill tree, links, rows, undo, JSON
src/WorkbookView.js  grid UI (CSS injected, framework-free)
demo/index.html      takeoff + workbook demo (needs ../pdf-viewport and ../pdf-snap)
```
