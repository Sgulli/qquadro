import type { FormulaValue } from "./types.js";
import { colLetter } from "./utils.js";

type Ref = string | number;

function resolveCol(col: Ref): string {
  return typeof col === "number" ? colLetter(col) : col;
}

function esc(v: string | number): string {
  return typeof v === "string" ? `"${v}"` : String(v);
}

export function ref(col: Ref, row: number): string {
  return `${resolveCol(col)}${row}`;
}

export function range(col: Ref, fromRow: number, toRow: number): string {
  const c = resolveCol(col);
  return `${c}${fromRow}:${c}${toRow}`;
}

export function rect(
  fromCol: Ref,
  fromRow: number,
  toCol: Ref,
  toRow: number,
): string {
  return `${resolveCol(fromCol)}${fromRow}:${resolveCol(toCol)}${toRow}`;
}

export function sum(rangeStr: string, result?: number): FormulaValue {
  return { formula: `SUM(${rangeStr})`, result };
}

export function average(rangeStr: string, result?: number): FormulaValue {
  return { formula: `AVERAGE(${rangeStr})`, result };
}

export function count(rangeStr: string, result?: number): FormulaValue {
  return { formula: `COUNT(${rangeStr})`, result };
}

export function max(rangeStr: string, result?: number): FormulaValue {
  return { formula: `MAX(${rangeStr})`, result };
}

export function min(rangeStr: string, result?: number): FormulaValue {
  return { formula: `MIN(${rangeStr})`, result };
}

export function add(...cells: (string | number)[]): FormulaValue {
  return { formula: cells.join("+") };
}

export function sub(a: string | number, b: string | number): FormulaValue {
  return { formula: `${a}-${b}` };
}

export function mul(a: string | number, b: string | number): FormulaValue {
  return { formula: `${a}*${b}` };
}

export function div(a: string | number, b: string | number): FormulaValue {
  return { formula: `${a}/${b}` };
}

export function pct(
  current: string | number,
  previous: string | number,
): FormulaValue {
  return { formula: `(${current}-${previous})/${previous}` };
}

export function ifExpr(
  condition: string,
  ifTrue: string | number | FormulaValue,
  ifFalse: string | number | FormulaValue,
): FormulaValue {
  const thenVal = typeof ifTrue === "object" && "formula" in ifTrue ? ifTrue.formula : esc(ifTrue);
  const elseVal = typeof ifFalse === "object" && "formula" in ifFalse ? ifFalse.formula : esc(ifFalse);
  return { formula: `IF(${condition},${thenVal},${elseVal})` };
}

export const F = {
  ref,
  range,
  rect,
  sum,
  average,
  count,
  max,
  min,
  add,
  sub,
  mul,
  div,
  pct,
  if: ifExpr,
} as const;
