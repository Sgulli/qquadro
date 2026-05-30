import type { CellFormulaValue } from "@cj-tech-master/excelts";
import { colLetter } from "./coords.js";
import type { CellPrimitive, CellValue, FormulaValue } from "./types.js";

type Ref = string | number;

function resolveCol(col: Ref): string {
  return typeof col === "number" ? colLetter(col) : col;
}

function esc(v: string | number): string {
  if (typeof v === "string") {
    const escaped = v.replace(/"/g, '""');
    return /^[=+\-@]/.test(v) ? `"'${escaped}"` : `"${escaped}"`;
  }
  return String(v);
}

/** Create a cell reference like "A1" from column letter/number and row. */
export function ref(col: Ref, row: number): string {
  return `${resolveCol(col)}${row}`;
}

/** Build a range string like "A1:A10" from column and row bounds. */
export function range(col: Ref, fromRow: number, toRow: number): string {
  const c = resolveCol(col);
  return `${c}${fromRow}:${c}${toRow}`;
}

/** Build a rectangular range like "A1:C10" from two corners. */
export function rect(fromCol: Ref, fromRow: number, toCol: Ref, toRow: number): string {
  return `${resolveCol(fromCol)}${fromRow}:${resolveCol(toCol)}${toRow}`;
}

/** Build a SUM formula. */
export function sum(rangeStr: string, result?: number): FormulaValue {
  return { formula: `SUM(${rangeStr})`, result };
}

/** Build an AVERAGE formula. */
export function average(rangeStr: string, result?: number): FormulaValue {
  return { formula: `AVERAGE(${rangeStr})`, result };
}

/** Build a COUNT formula. */
export function count(rangeStr: string, result?: number): FormulaValue {
  return { formula: `COUNT(${rangeStr})`, result };
}

/** Build a MAX formula. */
export function max(rangeStr: string, result?: number): FormulaValue {
  return { formula: `MAX(${rangeStr})`, result };
}

/** Build a MIN formula. */
export function min(rangeStr: string, result?: number): FormulaValue {
  return { formula: `MIN(${rangeStr})`, result };
}

/** Build an addition formula from multiple operands. */
export function add(...cells: (string | number)[]): FormulaValue {
  return { formula: cells.join("+") };
}

/** Build a subtraction formula. */
export function sub(a: string | number, b: string | number): FormulaValue {
  return { formula: `${a}-${b}` };
}

/** Build a multiplication formula. */
export function mul(a: string | number, b: string | number): FormulaValue {
  return { formula: `${a}*${b}` };
}

/** Build a division formula. */
export function div(a: string | number, b: string | number): FormulaValue {
  return { formula: `${a}/${b}` };
}

/** Build a percentage change formula: (current-previous)/previous. */
export function pct(current: string | number, previous: string | number): FormulaValue {
  return { formula: `(${current}-${previous})/${previous}` };
}

/** Build an IF formula. */
export function ifExpr(
  condition: string,
  ifTrue: string | number | FormulaValue,
  ifFalse: string | number | FormulaValue,
): FormulaValue {
  const thenVal = typeof ifTrue === "object" && "formula" in ifTrue ? ifTrue.formula : esc(ifTrue);
  const elseVal =
    typeof ifFalse === "object" && "formula" in ifFalse ? ifFalse.formula : esc(ifFalse);
  return { formula: `IF(${condition},${thenVal},${elseVal})` };
}

/** Build a VLOOKUP formula. */
export function vlookup(
  lookupValue: string | number,
  tableArray: string,
  colIndex: number,
  rangeLookup = false,
): FormulaValue {
  return {
    formula: `VLOOKUP(${esc(lookupValue)},${tableArray},${colIndex},${String(rangeLookup)})`,
  };
}

/** Build an HLOOKUP formula. */
export function hlookup(
  lookupValue: string | number,
  tableArray: string,
  rowIndex: number,
  rangeLookup = false,
): FormulaValue {
  return {
    formula: `HLOOKUP(${esc(lookupValue)},${tableArray},${rowIndex},${String(rangeLookup)})`,
  };
}

/** Build an INDEX formula. */
export function index(array: string, row: number, col?: number): FormulaValue {
  return { formula: col ? `INDEX(${array},${row},${col})` : `INDEX(${array},${row})` };
}

/** Build a MATCH formula. */
export function match(
  lookupValue: string | number,
  lookupArray: string,
  matchType = 0,
): FormulaValue {
  return { formula: `MATCH(${esc(lookupValue)},${lookupArray},${matchType})` };
}

/** Build an XLOOKUP formula. */
export function xlookup(
  lookupValue: string | number,
  lookupArray: string,
  returnArray: string,
  ifNotFound?: string,
  matchMode = 0,
): FormulaValue {
  const args = [esc(lookupValue), lookupArray, returnArray];
  if (ifNotFound !== undefined) args.push(esc(ifNotFound));
  if (matchMode !== 0) args.push(String(matchMode));
  return { formula: `XLOOKUP(${args.join(",")})` };
}

/** Build an OFFSET formula. */
export function offset(
  reference: string,
  rows: number,
  cols: number,
  height?: number,
  width?: number,
): FormulaValue {
  const args = [reference, String(rows), String(cols)];
  if (height !== undefined) args.push(String(height));
  if (width !== undefined) args.push(String(width));
  return { formula: `OFFSET(${args.join(",")})` };
}

/** Build an INDIRECT formula. */
export function indirect(refText: string): FormulaValue {
  return { formula: `INDIRECT(${esc(refText)})` };
}

/** Build an AND formula. */
export function and(...conditions: string[]): FormulaValue {
  return { formula: `AND(${conditions.join(",")})` };
}

/** Build an OR formula. */
export function or(...conditions: string[]): FormulaValue {
  return { formula: `OR(${conditions.join(",")})` };
}

/** Build a NOT formula. */
export function not(condition: string): FormulaValue {
  return { formula: `NOT(${condition})` };
}

/** Build a SWITCH formula. */
export function switchExpr(expression: string, ...cases: (string | number)[]): FormulaValue {
  return { formula: `SWITCH(${[expression, ...cases.map(esc)].join(",")})` };
}

/** Build an IFS formula. */
export function ifs(...conditions: string[]): FormulaValue {
  return { formula: `IFS(${conditions.join(",")})` };
}

/** Build an IFERROR formula. */
export function iferror(value: string | number, valueIfError: string | number): FormulaValue {
  return { formula: `IFERROR(${value},${esc(valueIfError)})` };
}

/** Build an IFNA formula. */
export function ifna(value: string | number, valueIfNA: string | number): FormulaValue {
  return { formula: `IFNA(${value},${esc(valueIfNA)})` };
}

/** Build a ROUND formula. */
export function round(number: string | number, digits = 0): FormulaValue {
  return { formula: `ROUND(${number},${digits})` };
}

/** Build a ROUNDUP formula. */
export function roundup(number: string | number, digits = 0): FormulaValue {
  return { formula: `ROUNDUP(${number},${digits})` };
}

/** Build a ROUNDDOWN formula. */
export function rounddown(number: string | number, digits = 0): FormulaValue {
  return { formula: `ROUNDDOWN(${number},${digits})` };
}

/** Build an ABS formula. */
export function abs(number: string | number): FormulaValue {
  return { formula: `ABS(${number})` };
}

/** Build a TRUNC formula. */
export function trunc(number: string | number, digits = 0): FormulaValue {
  return { formula: `TRUNC(${number},${digits})` };
}

/** Build an INT formula. */
export function int(number: string | number): FormulaValue {
  return { formula: `INT(${number})` };
}

/** Build a MOD formula. */
export function mod(number: string | number, divisor: string | number): FormulaValue {
  return { formula: `MOD(${number},${divisor})` };
}

/** Build a CEILING formula. */
export function ceiling(number: string | number, significance: string | number): FormulaValue {
  return { formula: `CEILING(${number},${significance})` };
}

/** Build a FLOOR formula. */
export function floor(number: string | number, significance: string | number): FormulaValue {
  return { formula: `FLOOR(${number},${significance})` };
}

/** Build a POWER formula. */
export function power(number: string | number, power: string | number): FormulaValue {
  return { formula: `POWER(${number},${power})` };
}

/** Build a SQRT formula. */
export function sqrt(number: string | number): FormulaValue {
  return { formula: `SQRT(${number})` };
}

/** Build a CONCAT formula. */
export function concat(...texts: (string | number)[]): FormulaValue {
  return { formula: `CONCAT(${texts.map(esc).join(",")})` };
}

/** Build a TEXT formula. */
export function text(value: string | number, format: string): FormulaValue {
  return { formula: `TEXT(${esc(value)},${esc(format)})` };
}

/** Build a LEFT formula. */
export function left(text: string | number, chars = 1): FormulaValue {
  return { formula: `LEFT(${esc(text)},${chars})` };
}

/** Build a RIGHT formula. */
export function right(text: string | number, chars = 1): FormulaValue {
  return { formula: `RIGHT(${esc(text)},${chars})` };
}

/** Build a MID formula. */
export function mid(text: string | number, start: number, chars: number): FormulaValue {
  return { formula: `MID(${esc(text)},${start},${chars})` };
}

/** Build a LEN formula. */
export function len(text: string | number): FormulaValue {
  return { formula: `LEN(${esc(text)})` };
}

/** Build a TRIM formula. */
export function trim(text: string | number): FormulaValue {
  return { formula: `TRIM(${esc(text)})` };
}

/** Build an UPPER formula. */
export function upper(text: string | number): FormulaValue {
  return { formula: `UPPER(${esc(text)})` };
}

/** Build a LOWER formula. */
export function lower(text: string | number): FormulaValue {
  return { formula: `LOWER(${esc(text)})` };
}

/** Build a PROPER formula. */
export function proper(text: string | number): FormulaValue {
  return { formula: `PROPER(${esc(text)})` };
}

/** Build a SUMIF formula. */
export function sumif(range: string, criteria: string | number, sumRange?: string): FormulaValue {
  return {
    formula: sumRange
      ? `SUMIF(${range},${esc(criteria)},${sumRange})`
      : `SUMIF(${range},${esc(criteria)})`,
  };
}

/** Build a SUMIFS formula. */
export function sumifs(
  sumRange: string,
  criteriaRange1: string,
  criteria1: string | number,
  ...more: (string | number)[]
): FormulaValue {
  const args = [sumRange, criteriaRange1, esc(criteria1)];
  for (let i = 0; i < more.length; i++) args.push(i % 2 === 0 ? String(more[i]) : esc(more[i]));
  return { formula: `SUMIFS(${args.join(",")})` };
}

/** Build a COUNTIF formula. */
export function countif(range: string, criteria: string | number): FormulaValue {
  return { formula: `COUNTIF(${range},${esc(criteria)})` };
}

/** Build a COUNTIFS formula. */
export function countifs(
  criteriaRange1: string,
  criteria1: string | number,
  ...more: (string | number)[]
): FormulaValue {
  const args = [criteriaRange1, esc(criteria1)];
  for (let i = 0; i < more.length; i++) args.push(i % 2 === 0 ? String(more[i]) : esc(more[i]));
  return { formula: `COUNTIFS(${args.join(",")})` };
}

/** Build an AVERAGEIF formula. */
export function averageif(
  range: string,
  criteria: string | number,
  averageRange?: string,
): FormulaValue {
  return {
    formula: averageRange
      ? `AVERAGEIF(${range},${esc(criteria)},${averageRange})`
      : `AVERAGEIF(${range},${esc(criteria)})`,
  };
}

/** Build an AVERAGEIFS formula. */
export function averageifs(
  averageRange: string,
  criteriaRange1: string,
  criteria1: string | number,
  ...more: (string | number)[]
): FormulaValue {
  const args = [averageRange, criteriaRange1, esc(criteria1)];
  for (let i = 0; i < more.length; i++) args.push(i % 2 === 0 ? String(more[i]) : esc(more[i]));
  return { formula: `AVERAGEIFS(${args.join(",")})` };
}

/** Build a SUBTOTAL formula. */
export function subtotal(functionNum: number, ref1: string, ...refs: string[]): FormulaValue {
  return { formula: `SUBTOTAL(${functionNum},${[ref1, ...refs].join(",")})` };
}

/** Build a NOW formula (current date/time). */
export function now(): FormulaValue {
  return { formula: "NOW()" };
}

/** Build a TODAY formula (current date). */
export function today(): FormulaValue {
  return { formula: "TODAY()" };
}

/** Build a DATE formula. */
export function date(year: number, month: number, day: number): FormulaValue {
  return { formula: `DATE(${year},${month},${day})` };
}

/** Build a YEAR formula. */
export function year(date: string | number): FormulaValue {
  return { formula: `YEAR(${date})` };
}

/** Build a MONTH formula. */
export function month(date: string | number): FormulaValue {
  return { formula: `MONTH(${date})` };
}

/** Build a DAY formula. */
export function day(date: string | number): FormulaValue {
  return { formula: `DAY(${date})` };
}

/** Build an EOMONTH formula. */
export function eomonth(startDate: string | number, months: number): FormulaValue {
  return { formula: `EOMONTH(${startDate},${months})` };
}

/** Build a NETWORKDAYS formula. */
export function networkdays(
  startDate: string | number,
  endDate: string | number,
  holidays?: string,
): FormulaValue {
  return {
    formula: holidays
      ? `NETWORKDAYS(${startDate},${endDate},${holidays})`
      : `NETWORKDAYS(${startDate},${endDate})`,
  };
}

/** Build an ISNUMBER formula. */
export function isnumber(value: string | number): FormulaValue {
  return { formula: `ISNUMBER(${value})` };
}

/** Build an ISTEXT formula. */
export function istext(value: string | number): FormulaValue {
  return { formula: `ISTEXT(${value})` };
}

/** Build an ISBLANK formula. */
export function isblank(value: string | number): FormulaValue {
  return { formula: `ISBLANK(${value})` };
}

/** Build an ISERROR formula. */
export function iserror(value: string | number): FormulaValue {
  return { formula: `ISERROR(${value})` };
}

/** Build a RANK formula. */
export function rank(number: string | number, ref: string, order = 0): FormulaValue {
  return { formula: `RANK(${number},${ref},${order})` };
}

/** Build a LARGE formula. */
export function large(array: string, k: number): FormulaValue {
  return { formula: `LARGE(${array},${k})` };
}

/** Build a SMALL formula. */
export function small(array: string, k: number): FormulaValue {
  return { formula: `SMALL(${array},${k})` };
}

/** Build a FILTER formula. */
export function filter(array: string, include: string, ifEmpty?: string): FormulaValue {
  return {
    formula: ifEmpty
      ? `FILTER(${array},${include},${esc(ifEmpty)})`
      : `FILTER(${array},${include})`,
  };
}

/** Build a SORT formula. */
export function sort(array: string, sortBy?: number, sortOrder = 1): FormulaValue {
  const args = [array];
  if (sortBy !== undefined) args.push(String(sortBy));
  if (sortOrder !== 1) args.push(String(sortOrder));
  return { formula: `SORT(${args.join(",")})` };
}

/** Build a UNIQUE formula. */
export function unique(array: string, byCol = false): FormulaValue {
  return byCol ? { formula: `UNIQUE(${array},TRUE)` } : { formula: `UNIQUE(${array})` };
}

/** Wrap a formula string in array braces `{}` for CSE entry. */
export function cse(formulaStr: string): FormulaValue {
  return { formula: `{${formulaStr}}` };
}

/** Check if a CellValue is a FormulaValue (object, not null, not Date). */
export function isFormula(val: CellValue): val is FormulaValue {
  return val !== null && typeof val === "object" && !(val instanceof Date);
}

function normalizeFormula(f: string): string {
  return f.startsWith("=") ? f.slice(1) : f;
}

/** Convert a raw {formula, result?} to a CellFormulaValue (strips leading `=`). */
export function toFormulaValue(v: { formula: string; result?: CellPrimitive }): CellFormulaValue {
  const fv: CellFormulaValue = { formula: normalizeFormula(v.formula) };
  if (v.result) fv.result = v.result;
  return fv;
}

/** Convert a CellValue to the excelts-compatible CellPrimitive | CellFormulaValue. */
export function toExcelValue(val: CellValue): CellPrimitive | CellFormulaValue {
  if (isFormula(val)) return toFormulaValue(val);
  return val;
}

/** Namespace object grouping all formula builders for convenient access. */
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
  ifExpr,
  vlookup,
  hlookup,
  index,
  match,
  xlookup,
  offset,
  indirect,
  and,
  or,
  not,
  switch: switchExpr,
  switchExpr,
  ifs,
  iferror,
  ifna,
  round,
  roundup,
  rounddown,
  abs,
  trunc,
  int,
  mod,
  ceiling,
  floor,
  power,
  sqrt,
  concat,
  text,
  left,
  right,
  mid,
  len,
  trim,
  upper,
  lower,
  proper,
  sumif,
  sumifs,
  countif,
  countifs,
  averageif,
  averageifs,
  subtotal,
  now,
  today,
  date,
  year,
  month,
  day,
  eomonth,
  networkdays,
  isnumber,
  istext,
  isblank,
  iserror,
  rank,
  large,
  small,
  filter,
  sort,
  unique,
  cse,
} as const;
