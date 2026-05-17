/**
 * @module excel-wrapper
 *
 * Fluent, fully-typed Excel builder for Node.js.
 * Powered by ExcelTS — supports styles, formulas, merged cells,
 * headers, freeze panes, auto-filter, and streaming (memory-safe) output.
 *
 * @example
 * ```ts
 * import { WorkbookBuilder, Styles } from "./index";
 *
 * await new WorkbookBuilder({ author: "Acme Corp" })
 *   .addSheet({ name: "Sales" }, (sheet) => {
 *     sheet
 *       .columns([
 *         { key: "product", header: "Product",    width: 22, headerStyle: Styles.header },
 *         { key: "units",   header: "Units Sold", width: 15, headerStyle: Styles.header },
 *         { key: "revenue", header: "Revenue",    width: 18, style: Styles.currency, headerStyle: Styles.header },
 *       ])
 *       .writeHeaders()
 *       .addRows([
 *         { product: "Widget A", units: 1_200, revenue: 36_000 },
 *         { product: "Widget B", units:   850, revenue: 25_500 },
 *       ])
 *       .addRow(
 *         { product: "Total", units: { formula: "=SUM(B2:B3)" }, revenue: { formula: "=SUM(C2:C3)" } },
 *         { style: Styles.totalRow }
 *       )
 *       .autoFitColumns()
 *       .freeze(1)
 *       .autoFilter();
 *   })
 *   .write("./output/sales.xlsx");
 * ```
 */

// ── Builders ────────────────────────────────────────────────────────────────

export { WorkbookBuilder } from "./workbook-builder.js";
export { SheetBuilder } from "./sheet-builder.js";

// ── Formula helpers ───────────────────────────────────────────────────────────

export { F, ref, range, rect, sum, average, count, max, min, add, sub, mul, div, pct, ifExpr } from "./formulas.js";

// ── Style utilities ──────────────────────────────────────────────────────────

export {
  Styles,
  applyStyle,
  formatHeaderFooterSection,
  toExcelFont,
  toExcelFill,
  toExcelBorder,
  toExcelAlignment,
  resolveNumberFormat,
  colLetter,
} from "./utils.js";

// ── Types ────────────────────────────────────────────────────────────────────

export type {
  // Cell value types
  CellPrimitive,
  CellValue,
  FormulaValue,
  // Style types
  CellStyle,
  Font,
  Fill,
  SolidFill,
  GradientFill,
  Border,
  BorderSide,
  BorderStyle,
  Alignment,
  HorizontalAlignment,
  VerticalAlignment,
  NumberFormat,
  BuiltinFormat,
  // Column / row / sheet
  ColumnDef,
  MergeRange,
  RowData,
  RowOptions,
  SheetOptions,
  SheetHeaderFooter,
  HeaderFooterSection,
  // Workbook
  WorkbookOptions,
  WriteResult,
} from "./types.js";

// ── Constants ────────────────────────────────────────────────────────────────

export { BUILTIN_FORMATS } from "./types.js";
