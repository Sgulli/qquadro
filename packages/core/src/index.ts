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

export { RangeBuilder } from "./range-builder.js";
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

// ── Range & Coordinate utilities ──────────────────────────────────────────────

export { cellRef, colLetter, colRange, rangeRef } from "./coords.js";
export { col, Range, row } from "./range.js";
export {
  accounting,
  align,
  border,
  color,
  currency,
  fill,
  font,
  numFmt,
  Styles,
  style,
} from "./style-presets.js";

// ── Types ────────────────────────────────────────────────────────────────────

export type {
  Addr,
  AddSparklineGroupOptions,
  Alignment,
  Border,
  BorderSide,
  BorderStyle,
  BuiltinFormat,
  CellPrimitive,
  CellRange,
  CellStyle,
  CellValue,
  ColumnDef,
  ExternalLinkInput,
  Fill,
  Font,
  FormulaValue,
  GradientFill,
  HeaderFooterSection,
  HorizontalAlignment,
  IgnoredErrorDef,
  MergeRange,
  NoteConfig,
  NoteText,
  NumberFormat,
  RangeValidationDef,
  RowData,
  RowOptions,
  SheetHeaderFooter,
  SheetOptions,
  SolidFill,
  Sparkline,
  SparklineAxisType,
  SparklineColor,
  SparklineGroup,
  SparklineType,
  VerticalAlignment,
  WorkbookOptions,
  WriteResult,
} from "./types.js";

// ── Constants ────────────────────────────────────────────────────────────────

export { BUILTIN_FORMATS } from "./types.js";

// ── Re-exports from excelts ──────────────────────────────────────────────────

export type {
  AddImageRange,
  CellHyperlinkValue,
  CellHyperlinkValueInput,
  CellIsOperators,
  CellRichTextValue,
  ConditionalFormattingOptions,
  ConditionalFormattingRule,
  ContainsTextOperators,
  DataValidation,
  DataValidationOperator,
  DefinedNameModel,
  ExternalLinkModel,
  FormCheckboxOptions,
  FormControlRange,
  IconSetTypes,
  ImageData,
  RichText as RichTextRun,
  TableColumnProperties,
  TableStyleProperties,
  ThreadedComment,
  ThreadedCommentMention,
  ThreadedCommentPerson,
  TimePeriodTypes,
  WatermarkMode,
  WatermarkOptions,
} from "@cj-tech-master/excelts";

export type {
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
  AddLegendOptions,
  AddPieChartOptions,
  AddPlotAreaOptions,
  AddScatterChartOptions,
  AddSurfaceChartOptions,
  AddTitleOptions,
  AddTrendlineOptions,
  ChartType,
  ExcelChartExPreset,
  ExcelChartPreset,
  SeriesFromColumnsOptions,
} from "@cj-tech-master/excelts/chart";
