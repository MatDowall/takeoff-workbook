import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/Workbook.js';
import { ERR, isErr } from '../src/formula.js';

const close = (a, b, eps = 1e-9) => assert.ok(typeof a === 'number' && Math.abs(a - b) < eps, `${a} ≠ ${b}`);

function estimate() {
  const wb = new Workbook();
  const R = wb.rootId;
  wb.setCells(R, [
    { a1: 'A1', raw: '01' }, { a1: 'B1', raw: 'Concrete slab' }, { a1: 'C1', raw: '24.5' }, { a1: 'D1', raw: 'm3' }, { a1: 'E1', raw: '320' },
    { a1: 'B2', raw: 'Reinforcing' }, { a1: 'C2', raw: '1.8' }, { a1: 'D2', raw: 't' }, { a1: 'E2', raw: '2,950' }, { a1: 'G2', raw: '1.1' },
  ]);
  return { wb, R };
}

test('row maths: Subtotal = Qty × Rate, Total = Subtotal × Factor; text columns stay text', () => {
  const { wb, R } = estimate();
  assert.equal(wb.value(R, 'A1'), '01');
  close(wb.value(R, 'F1'), 7840);
  close(wb.value(R, 'H1'), 7840); // factor blank = 1
  close(wb.value(R, 'H2'), 1.8 * 2950 * 1.1);
  close(wb.total(R), 7840 + 1.8 * 2950 * 1.1);
  assert.equal(wb.display(R, 'E2'), '2,950.00');
  assert.equal(wb.value(R, 'H5'), null); // empty row → blank, not 0
});

test('override a computed cell, clear to restore', () => {
  const { wb, R } = estimate();
  wb.setCell(R, 'H1', '8000');
  close(wb.value(R, 'H1'), 8000);
  assert.equal(wb.info(R, 'H1').type, 'override');
  wb.setCell(R, 'H1', '');
  close(wb.value(R, 'H1'), 7840);
});

test('3-level drill-down rolls up live — no recalculate', () => {
  const { wb, R } = estimate();
  // drill slab quantity → dimension lines
  const Q = wb.drill(R, 'C1');
  close(wb.value(R, 'C1'), 24.5); // existing figure carried into row 1
  assert.equal(wb.sheet(Q).kind, 'quantity');
  wb.setCells(Q, [
    { a1: 'C1', raw: '' }, { a1: 'B1', raw: 'Slab' }, { a1: 'D1', raw: '12' }, { a1: 'E1', raw: '8' }, { a1: 'F1', raw: '0.2' }, // 19.2
    { a1: 'B2', raw: 'Thickenings' }, { a1: 'C2', raw: '4' }, { a1: 'D2', raw: '8' }, { a1: 'E2', raw: '0.4' }, { a1: 'F2', raw: '0.3' }, // 3.84
  ]);
  close(wb.total(Q), 19.2 + 3.84);
  close(wb.value(R, 'C1'), 23.04);
  close(wb.value(R, 'H1'), 23.04 * 320);
  // drill a dimension inside the quantity sheet (nested quantity)
  const L = wb.drill(Q, 'D1');
  wb.setCells(L, [{ a1: 'B1', raw: 'grid 1-3' }, { a1: 'D1', raw: '6.5' }, { a1: 'B2', raw: 'grid 3-5' }, { a1: 'D2', raw: '6.5' }, { a1: 'C1', raw: '' }]);
  close(wb.value(Q, 'D1'), 13);
  // drill the slab RATE → build-up
  const Rt = wb.drill(R, 'E1');
  close(wb.value(R, 'E1'), 320); // carried
  wb.setCells(Rt, [{ a1: 'B1', raw: 'Concrete 25MPa' }, { a1: 'C1', raw: '1.05' }, { a1: 'E1', raw: '265' }, { a1: 'B2', raw: 'Place & finish' }, { a1: 'C2', raw: '1' }, { a1: 'E2', raw: '48' }, { a1: 'E1', raw: '265' }]);
  wb.setCell(Rt, 'B1', 'Concrete 25MPa');
  // remove the carried "Existing rate" row? it's row 1 → we overwrote it. total = 1.05*265 + 48
  close(wb.value(R, 'E1'), 1.05 * 265 + 48);
  // deepest edit flows all the way up
  const before = wb.total(R);
  wb.setCell(L, 'D2', '7.5');
  const slabQty = (6.5 + 7.5) * 8 * 0.2 + 3.84;
  close(wb.value(R, 'C1'), slabQty);
  close(wb.total(R), slabQty * (1.05 * 265 + 48) + 1.8 * 2950 * 1.1);
  assert.notEqual(wb.total(R), before);
  assert.deepEqual(wb.path(L).map((s) => s.id), [R, Q, L]);
});

test('measurement links update live, with modifier and missing measurement', () => {
  const { wb, R } = estimate();
  wb.setMeasurement('m1', { value: 30, unit: 'm2', name: 'Ground floor' });
  wb.link(R, 'C3', 'm1');
  wb.setCells(R, [{ a1: 'B3', raw: 'Floor coating' }, { a1: 'E3', raw: '18' }]);
  assert.equal(wb.value(R, 'D3'), 'm2'); // unit filled from measurement
  close(wb.value(R, 'H3'), 540);
  wb.setMeasurement('m1', { value: 32 }); // re-measured in the takeoff
  close(wb.value(R, 'H3'), 576);
  wb.setLinkExpr(R, 'C3', 'link*1.05');
  close(wb.value(R, 'C3'), 33.6);
  // MEAS() by name from any formula
  wb.setCell(R, 'I3', '=MEAS("ground floor")/2');
  close(wb.value(R, 'I3'), 16);
  wb.removeMeasurement('m1');
  assert.equal(wb.value(R, 'C3'), ERR.REF);
  assert.ok(isErr(wb.total(R))); // surfaced, not silently zero
  assert.equal(wb.info(R, 'C3').missing, true);
});

test('cross-sheet refs, names, cycles', () => {
  const { wb, R } = estimate();
  const Q = wb.drill(R, 'C2');
  wb.renameSheet(Q, 'Rebar');
  wb.setCell(Q, 'C1', '2');
  wb.setCell(R, 'I1', '=[Rebar]!H1*10');
  close(wb.value(R, 'I1'), 20);
  wb.renameSheet(Q, 'Reo'); // references follow the rename
  assert.equal(wb.cell(R, 'I1').raw, '=[Reo]!H1*10');
  wb.setName('wastage', '1.05');
  wb.setCell(R, 'G1', '=wastage');
  close(wb.value(R, 'H1'), 7840 * 1.05);
  wb.setName('wastage', '1.1');
  close(wb.value(R, 'H1'), 7840 * 1.1);
  // cycle: child quantity refers to the parent total that depends on it
  wb.setCell(Q, 'C1', '=[Estimate]!H2');
  assert.equal(wb.value(R, 'C2'), ERR.CIRC);
  wb.setCell(Q, 'C1', '3'); // breaking the cycle recovers
  close(wb.value(R, 'C2'), 3);
});

test('subtotal and heading rows', () => {
  const wb = new Workbook();
  const R = wb.rootId;
  wb.setCells(R, [{ a1: 'B1', raw: 'SUBSTRUCTURE' }, { a1: 'C2', raw: '2' }, { a1: 'E2', raw: '10' }, { a1: 'C3', raw: '1' }, { a1: 'E3', raw: '5' }, { a1: 'B4', raw: 'Subtotal' }, { a1: 'C5', raw: '1' }, { a1: 'E5', raw: '100' }]);
  wb.setRowType(R, 1, 'heading');
  wb.setRowType(R, 4, 'subtotal');
  close(wb.value(R, 'H4'), 25);
  close(wb.total(R), 125); // subtotal not double-counted
});

test('insert/delete rows rewrite formulas everywhere (Excel-style)', () => {
  const { wb, R } = estimate();
  wb.setCell(R, 'I10', '=SUM(H1:H2)');
  const Q = wb.drill(R, 'C2');
  wb.setCell(Q, 'C5', '=[Estimate]!C1');
  wb.insertRows(R, 1, 2);
  assert.equal(wb.cell(R, 'I12').raw, '=SUM(H3:H4)');
  assert.equal(wb.cell(Q, 'C5').raw, '=[Estimate]!C3');
  assert.equal(wb.sheet(Q).parent.a1, 'C4'); // breakdown moved with its row
  close(wb.value(R, 'I12'), 7840 + wb.value(R, 'H4'));
  wb.deleteRows(R, 3, 1); // delete the slab row
  assert.equal(wb.cell(R, 'I11').raw, '=SUM(H3:H3)'); // range shrinks, like Excel
  assert.equal(wb.cell(Q, 'C5').raw, '=[Estimate]!#REF!');
  assert.equal(wb.value(Q, 'C5'), ERR.REF);
  wb.deleteRows(R, 3, 1); // delete the reo row → its breakdown goes too
  assert.equal(wb.sheet(Q), undefined);
  wb.undo(); // and comes back
  assert.ok(wb.sheet(Q));
});

test('undo / redo restores values and breakdowns', () => {
  const { wb, R } = estimate();
  const t0 = wb.total(R);
  const Q = wb.drill(R, 'C1');
  wb.setCell(Q, 'C1', '30');
  const t1 = wb.total(R);
  assert.notEqual(t0, t1);
  wb.undo(); // cell edit
  wb.undo(); // drill
  assert.equal(wb.sheet(Q), undefined);
  close(wb.total(R), t0);
  wb.redo();
  wb.redo();
  close(wb.total(R), t1);
  wb.batch('two edits', () => {
    wb.setCell(R, 'C2', '1');
    wb.setCell(R, 'E2', '1');
  });
  assert.equal(wb.undo(), 'two edits');
});

test('copy/paste shifts formulas and deep-copies breakdowns', () => {
  const { wb, R } = estimate();
  const Q = wb.drill(R, 'C1');
  wb.setCell(Q, 'C1', '10');
  wb.setCell(R, 'I1', '=C1*2');
  const clip = wb.copy(R, 1, 0, 1, 9); // A1:J1
  wb.paste(R, 5, 0, clip);
  assert.equal(wb.cell(R, 'I5').raw, '=C5*2');
  const Q2 = wb.cell(R, 'C5').drill;
  assert.ok(Q2 && Q2 !== Q);
  wb.setCell(Q2, 'C1', '20'); // independent copy — no Nezo-style "priced twice"
  close(wb.value(R, 'C1'), 10);
  close(wb.value(R, 'C5'), 20);
  close(wb.value(R, 'I5'), 40);
  // plain text from Excel
  wb.pasteText(R, 8, 1, 'Formwork\t12\tm2\t45\nStripping\t12\tm2\t=5+3');
  close(wb.value(R, 'H8'), 540);
  close(wb.value(R, 'H9'), 96);
});

test('JSON round-trip keeps everything (and stays live)', () => {
  const { wb, R } = estimate();
  wb.setMeasurement('m1', { value: 5, unit: 'm', name: 'Edge' });
  const Q = wb.drill(R, 'C2');
  wb.link(Q, 'D1', 'm1');
  wb.setCell(Q, 'C1', '2');
  wb.setName('gst', '15%');
  const json = JSON.parse(JSON.stringify(wb.toJSON()));
  const wb2 = Workbook.fromJSON(json);
  close(wb2.total(R), wb.total(R));
  wb2.setMeasurement('m1', { value: 6 });
  close(wb2.value(R, 'C2'), 12);
});

test('performance: 5,000-row estimate with 1,000 breakdowns', () => {
  const wb = new Workbook({ rows: 5000 });
  const R = wb.rootId;
  wb.batch('build', () => {
    const edits = [];
    for (let r = 1; r <= 5000; r++) edits.push({ a1: `C${r}`, raw: String(r % 7) }, { a1: `E${r}`, raw: String(10 + (r % 13)) });
    wb.setCells(R, edits);
  });
  for (let r = 1; r <= 1000; r++) {
    const q = wb.drill(R, `C${r * 5}`);
    wb.setCell(q, 'D1', String(r));
  }
  let t = performance.now();
  const total = wb.total(R);
  const full = performance.now() - t;
  t = performance.now();
  const leaf = wb.cell(R, 'C2500').drill;
  wb.setCell(leaf, 'D1', '999');
  const after = wb.total(R);
  const incr = performance.now() - t;
  console.log(`# full calc ${full.toFixed(1)} ms · edit deep cell + new total ${incr.toFixed(2)} ms`);
  assert.ok(after !== total);
  assert.ok(full < 500 && incr < 50);
});
