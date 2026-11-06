/**
 * columns.js — what each column MEANS on each kind of sheet.
 *
 * This is the heart of a takeoff workbook (and where it differs from a
 * generic spreadsheet): rows aren't free-form, they're "quantity × rate".
 * Fixed roles mean the maths is built in, totals can't be broken by a bad
 * formula, and every drill-down knows what it's breaking down.
 *
 *  COST sheet (the estimate, and "cost" drill-downs)
 *    A Code · B Description · C Quantity · D Unit · E Rate · F Subtotal = C×E · G Factor · H Total = F×G
 *  RATE sheet (build-up of one rate — resources × usage)
 *    same columns and maths as COST; its total becomes the parent's RATE
 *  QUANTITY sheet (dimension lines that make up one quantity)
 *    A Code · B Description · C Number · D Length · E Width · F Depth · G Factor · H Quantity
 *    H = product of whichever of C–F are filled × G     (e.g. 2 × 5.4 × 0.3 = 3.24)
 *
 *  User columns (I onwards) are free: notes, trade, cost codes, formulas.
 *
 *  DRILL-DOWN (same three types as CostX):
 *    Quantity cell  → QUANTITY sheet      (cost/rate sheets: C;  quantity sheets: C–F)
 *    Rate cell      → RATE sheet          (E)
 *    Total cell     → COST sheet          (H on cost/rate sheets)
 *  A drilled cell's value is ALWAYS its child sheet's live total.
 */

import { indexToCol } from './formula.js';

export const KIND = { COST: 'cost', RATE: 'rate', QUANTITY: 'quantity' };

const COST_COLS = {
  A: { label: 'Code', type: 'text', width: 70 },
  B: { label: 'Description', type: 'text', width: 260 },
  C: { label: 'Quantity', type: 'num', width: 90, drill: KIND.QUANTITY, link: true, dp: 2 },
  D: { label: 'Unit', type: 'text', width: 52 },
  E: { label: 'Rate', type: 'num', width: 90, drill: KIND.RATE, money: true, dp: 2 },
  F: { label: 'Subtotal', type: 'num', width: 104, computed: true, money: true, dp: 2 },
  G: { label: 'Factor', type: 'num', width: 64, dp: 3, defaultValue: 1 },
  H: { label: 'Total', type: 'num', width: 112, computed: true, drill: KIND.COST, money: true, dp: 2, total: true },
};

const QTY_COLS = {
  A: { label: 'Code', type: 'text', width: 70 },
  B: { label: 'Description', type: 'text', width: 260 },
  C: { label: 'Number', type: 'num', width: 80, drill: KIND.QUANTITY, link: true, dp: 3, factor: true },
  D: { label: 'Length', type: 'num', width: 80, drill: KIND.QUANTITY, link: true, dp: 3, factor: true },
  E: { label: 'Width', type: 'num', width: 80, drill: KIND.QUANTITY, link: true, dp: 3, factor: true },
  F: { label: 'Depth', type: 'num', width: 80, drill: KIND.QUANTITY, link: true, dp: 3, factor: true },
  G: { label: 'Factor', type: 'num', width: 64, dp: 3, defaultValue: 1 },
  H: { label: 'Quantity', type: 'num', width: 104, computed: true, dp: 3, total: true },
};

/** Column definitions for a sheet kind (+ user columns). */
export function columnsFor(kind, userColumns = 6, userLabels = []) {
  const base = kind === KIND.QUANTITY ? QTY_COLS : COST_COLS;
  const cols = { ...base };
  if (kind === KIND.RATE) {
    cols.B = { ...cols.B, label: 'Resource' };
    cols.C = { ...cols.C, label: 'Usage' };
    cols.H = { ...cols.H, label: 'Rate' };
  }
  for (let i = 0; i < userColumns; i++) {
    const col = indexToCol(8 + i); // I, J, K … Z, AA … (CostX allows 30)
    cols[col] = { label: userLabels[i] ?? `User ${i + 1}`, type: 'any', width: 90, user: true };
  }
  return cols;
}

/** Title for a child sheet created by drilling a given column. */
export const drillTitle = { quantity: 'Quantity', rate: 'Rate build-up', cost: 'Cost breakdown' };
