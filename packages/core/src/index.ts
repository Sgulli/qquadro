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
 *       .headers([
 *         { key: "product", header: "Product",  width: 22, headerStyle: Styles.header },
 *         { key: "revenue", header: "Revenue",    width: 18, style: Styles.currency, headerStyle: Styles.header },
 *       ])
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

export { SheetBuilder } from "./sheet-builder.js";
export { WorkbookBuilder } from "./workbook-builder.js";

// ── Formula helpers ───────────────────────────────────────────────────────────

export {
  add,
  average,
  count,
  div,
  F,
  ifExpr,
  max,
  min,
  mul,
  pct,
  range,
  rect,
  ref,
  sub,
  sum,
} from "./formulas.js";

// ── Style utilities ──────────────────────────────────────────────────────────

export {
  accounting,
  align,
  applyStyle,
  border,
  cellRef,
  colLetter,
  colRange,
  currency,
  fill,
  font,
  formatHeaderFooterSection,
  numFmt,
  rangeRef,
  resolveNumberFormat,
  Styles,
  style,
  toExcelAlignment,
  toExcelBorder,
  toExcelFill,
  toExcelFont,
} from "./utils.js";

// ── Types ────────────────────────────────────────────────────────────────────

export type {
  // Chart types
  AddAxisOptions,
  AddBarChartOptions,
  AddChartExOptions,
  AddChartFromTableOptions,
  AddChartMarkerOptions,
  AddChartOptions,
  AddChartRange,
  AddComboChartOptions,
  AddDataLabelsOptions,
  AddDataPointOptions,
  AddErrorBarsOptions,
  // Image types
  AddImageRange,
  AddLegendOptions,
  AddPieChartOptions,
  AddPlotAreaOptions,
  AddScatterChartOptions,
  // Sparkline types
  AddSparklineGroupOptions,
  AddSurfaceChartOptions,
  AddTitleOptions,
  AddTrendlineOptions,
  Alignment,
  Border,
  BorderSide,
  BorderStyle,
  BuiltinFormat,
  // Cell value types
  CellHyperlinkValue,
  CellHyperlinkValueInput,
  // Cell value types
  CellPrimitive,
  CellRichTextValue,
  // Style types
  CellStyle,
  CellValue,
  ChartType,
  // Column / row / sheet
  ColumnDef,
  ExcelChartExPreset,
  ExcelChartPreset,
  Fill,
  Font,
  FormulaValue,
  GradientFill,
  HeaderFooterSection,
  HorizontalAlignment,
  ImageData,
  MergeRange,
  // Comment / Note types
  NoteConfig,
  NoteText,
  NumberFormat,
  RichTextRun,
  RowData,
  RowOptions,
  SeriesFromColumnsOptions,
  SheetHeaderFooter,
  SheetOptions,
  SolidFill,
  Sparkline,
  SparklineAxisType,
  SparklineColor,
  SparklineGroup,
  SparklineType,
  ThreadedComment,
  ThreadedCommentMention,
  ThreadedCommentPerson,
  VerticalAlignment,
  WatermarkMode,
  WatermarkOptions,
  // Workbook
  WorkbookOptions,
  WriteResult,
} from "./types.js";

// ── Constants ────────────────────────────────────────────────────────────────

export { BUILTIN_FORMATS } from "./types.js";

// ── Data Validation & Conditional Formatting (re-exported from excelts) ──────

export type {
  CellIsOperators,
  ConditionalFormattingOptions,
  ConditionalFormattingRule,
  ContainsTextOperators,
  DataValidation,
  DataValidationOperator,
  DefinedNameModel,
  IconSetTypes,
  TableColumnProperties,
  TableStyleProperties,
  TimePeriodTypes,
} from "@cj-tech-master/excelts";
