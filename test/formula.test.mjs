import { test } from 'node:test';
import assert from 'node:assert/strict';
import { parse, evaluate, ERR, offsetFormula, rewriteRefs, isErr } from '../src/formula.js';

const cells = { A1: 2, A2: 3, A3: 4, B1: 'text', C1: null, D1: 0 };
const ctx = { cell: (sh, c, r) => (sh === 'Other' ? 100 : sh ? ERR.REF : cells[`${c}${r}`] ?? null), name: (n) => (n === 'gst' ? 0.15 : undefined), meas: (k) => (k === 'Wall A' ? 12.5 : undefined) };
const ev = (f, x = ctx) => evaluate(parse(f), x);

test('arithmetic & precedence', () => {
  assert.equal(ev('1+2*3'), 7);
  assert.equal(ev('(1+2)*3'), 9);
  assert.equal(ev('2^3^2'), 512); // right-assoc
  assert.equal(ev('-2^2'), -4);
  assert.equal(ev('50%*A1'), 1);
  assert.equal(ev('.35*100'), 35);
  assert.equal(ev('1e3/4'), 250);
});
test('refs, ranges, functions', () => {
  assert.equal(ev('A1*A2'), 6);
  assert.equal(ev('a1+a2'), 5); // lower-case refs
  assert.equal(ev('$A$1+A$2'), 5);
  assert.equal(ev('SUM(A1:A3)'), 9);
  assert.equal(ev('SUM(A1:A3, 10)'), 19);
  assert.equal(ev('AVERAGE(A1:A3)'), 3);
  assert.equal(ev('MAX(A1:A3)-MIN(A1:A3)'), 2);
  assert.equal(ev('COUNT(A1:C1)'), 1); // text & empty skipped
  assert.equal(ev('C1*5'), 0); // empty = 0
  assert.equal(ev('[Other]!H5'), 100);
});
test('rounding helpers (QS favourites)', () => {
  assert.equal(ev('ROUND(2.675,2)'), 2.68);
  assert.equal(ev('ROUNDUP(12.01,0)'), 13);
  assert.equal(ev('ROUNDDOWN(12.99,0)'), 12);
  assert.equal(ev('CEILING(7.1,0.6)'), 7.2);
  assert.equal(ev('CEILING(2,0.5)'), 2);
  assert.equal(ev('FLOOR(2,0.5)'), 2);
  assert.ok(Math.abs(ev('FLOOR(7.1,0.6)') - 6.6) < 1e-12);
});
test('logic', () => {
  assert.equal(ev('IF(A1>1,"big","small")'), 'big');
  assert.equal(ev('IF(AND(A1=2,NOT(A2=2)),1,0)'), 1);
  assert.equal(ev('IFERROR(1/D1,0)'), 0);
  assert.equal(ev('"a"&"b"'), 'ab');
});
test('names, measurements, link', () => {
  assert.ok(Math.abs(ev('100*(1+gst)') - 115) < 1e-9);
  assert.equal(ev('MEAS("Wall A")*2'), 25);
  assert.equal(evaluate(parse('link*1.05'), { ...ctx, link: 10 }), 10.5);
});
test('errors propagate', () => {
  assert.equal(ev('1/D1'), ERR.DIV0);
  assert.equal(ev('B1*2'), ERR.VALUE);
  assert.equal(ev('nope+1'), ERR.NAME);
  assert.equal(ev('FOO(1)'), ERR.NAME);
  assert.equal(ev('[Missing]!A1'), ERR.REF);
  assert.equal(ev('MEAS("nope")'), ERR.REF);
  assert.ok(isErr(ev('SUM(A1:A3)+1/D1')));
  assert.throws(() => parse('1+'));
  assert.throws(() => parse('SUM(1,'));
});
test('rewriting: copy offset respects $ and keeps formatting', () => {
  assert.equal(offsetFormula('C3*E3 + $G$1', 0, 2), 'C5*E5 + $G$1');
  assert.equal(offsetFormula('SUM(H1:H9)', 1, 0), 'SUM(I1:I9)');
  assert.equal(offsetFormula('A1', -1, 0), '#REF!');
  assert.equal(rewriteRefs('[Slab]!C3+C3', (r, sh) => (sh ? { ...r, row: 9 } : r)), '[Slab]!C9+C3');
});
