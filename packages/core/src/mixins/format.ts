import type { Worksheet as ExcelWorksheet } from "@cj-tech-master/excelts";
import { colLetter, resolveRange } from "../coords.js";
import { isFormula, toFormulaValue } from "../formulas.js";
import { applyStyle } from "../style-presets.js";
import type { CellRange, CellStyle, CellValue } from "../types.js";

export function applyFreeze(ws: ExcelWorksheet, row: number, col: number): void {
  const prev = ws.views?.[0];
  const { state: _, style: _1, ...rest } = prev ?? {};
  ws.views = [
    {
      ...rest,
      state: "frozen",
      xSplit: col,
      ySplit: row,
      topLeftCell: `${colLetter(col + 1)}${row + 1}`,
      activeCell: `${colLetter(col + 1)}${row + 1}`,
    },
  ];
}

export function applyAutoFilter(ws: ExcelWorksheet, range: string): void {
  ws.autoFilter = range;
}

export function applyStyleRange(ws: ExcelWorksheet, range: CellRange, style: CellStyle): void {
  const [tl, br] = resolveRange(range).split(":");
  if (!tl) return;
  const tlCell = ws.getCell(tl);
  const brCell = ws.getCell(br ?? tl);
  for (let r = tlCell.fullAddress.row; r <= brCell.fullAddress.row; r++) {
    for (let c = tlCell.fullAddress.col; c <= brCell.fullAddress.col; c++) {
      applyStyle(ws.getCell(r, c), style);
    }
  }
}

export function applyMerge(
  ws: ExcelWorksheet,
  range: string,
  value?: CellValue,
  style?: CellStyle,
  height?: number,
): void {
  ws.mergeCells(range);
  const [tl] = range.split(":");
  if (!tl) return;
  const cell = ws.getCell(tl);
  if (value) {
    if (isFormula(value)) {
      (cell as unknown as { value: object }).value = toFormulaValue(value);
    } else {
      cell.value = value;
    }
  }
  if (style) applyStyle(cell, style);
  if (height) ws.getRow(cell.fullAddress.row).height = height;
}
