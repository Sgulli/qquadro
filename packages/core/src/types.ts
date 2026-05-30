// ─── Overload helper types ────────────────────────────────────────────────────

/** Single cell address: A1 string or `[col, row]` numeric pair. */
export type Addr = string | number[];

/** Cell range: A1 string or `[col1, row1, col2, row2]` numeric quadruple. */
export type CellRange = string | number[];

// ─── Data validation ─────────────────────────────────────────────────────────

/** Grouped input for `SheetBuilder.addRangeValidation()`. */
export interface RangeValidationDef {
  type: "whole" | "decimal" | "date" | "textLength";
  operator:
    | "between"
    | "notBetween"
    | "equal"
    | "notEqual"
    | "greaterThan"
    | "lessThan"
    | "greaterThanOrEqual"
    | "lessThanOrEqual";
  formulae: (string | number | Date)[];
  allowBlank?: boolean;
  error?: string;
  errorTitle?: string;
  prompt?: string;
  promptTitle?: string;
  showErrorMessage?: boolean;
  showInputMessage?: boolean;
}

// ─── Alignment ───────────────────────────────────────────────────────────────

export type HorizontalAlignment =
  | "left"
  | "center"
  | "right"
  | "fill"
  | "justify"
  | "centerContinuous"
  | "distributed";

export type VerticalAlignment = "top" | "middle" | "bottom" | "distributed" | "justify";

export interface Alignment {
  horizontal?: HorizontalAlignment;
  vertical?: VerticalAlignment;
  wrapText?: boolean;
  shrinkToFit?: boolean;
  indent?: number;
  textRotation?: number | "vertical";
}

// ─── Border ──────────────────────────────────────────────────────────────────

export type BorderStyle =
  | "thin"
  | "medium"
  | "thick"
  | "dotted"
  | "dashed"
  | "hair"
  | "mediumDashed"
  | "dashDot"
  | "mediumDashDot"
  | "dashDotDot"
  | "mediumDashDotDot"
  | "slantDashDot"
  | "double";

export interface BorderSide {
  style?: BorderStyle;
  color?: string; // hex without "#", e.g. "FF0000"
}

export interface Border {
  top?: BorderSide;
  bottom?: BorderSide;
  left?: BorderSide;
  right?: BorderSide;
  diagonal?: BorderSide & { up?: boolean; down?: boolean };
}

// ─── Fill ────────────────────────────────────────────────────────────────────

export interface SolidFill {
  type: "solid";
  color: string; // hex ARGB, e.g. "FFFFFF00" for yellow
}

export interface GradientFill {
  type: "gradient";
  gradient: "angle" | "path";
  degree?: number;
  stops: Array<{ position: number; color: string }>;
}

export type Fill = SolidFill | GradientFill;

// ─── Font ────────────────────────────────────────────────────────────────────

export interface Color {
  argb: string;
  theme: number;
}

export interface Font {
  name?: string;
  size?: number;
  family?: number;
  scheme?: "minor" | "major" | "none";
  charset?: number;
  color?: string | Partial<Color>;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean | "none" | "single" | "double" | "singleAccounting" | "doubleAccounting";
  vertAlign?: "superscript" | "subscript";
  strike?: boolean;
  outline?: boolean;
  condense?: boolean;
  extend?: boolean;
  shadow?: boolean;
}

// ─── Number Format ───────────────────────────────────────────────────────────

export type BuiltinFormat =
  | "general"
  | "number"
  | "integer"
  | "float"
  | "currency"
  | "accountingUSD"
  | "percent"
  | "percentDecimal"
  | "date"
  | "datetime"
  | "time"
  | "scientific"
  | "text";

export type NumberFormat = BuiltinFormat | (string & {}); // custom format string passthrough

export const BUILTIN_FORMATS: Record<BuiltinFormat, string> = {
  general: "General",
  number: "#,##0",
  integer: "0",
  float: "#,##0.00",
  currency: "#,##0.00",
  accountingUSD: '#,##0.00;(#,##0.00);"-"',
  percent: "0%",
  percentDecimal: "0.00%",
  date: "yyyy-mm-dd",
  datetime: "yyyy-mm-dd hh:mm:ss",
  time: "hh:mm:ss",
  scientific: "0.00E+00",
  text: "@",
};

// ─── Cell Style ──────────────────────────────────────────────────────────────

export interface CellStyle {
  font?: Font;
  fill?: Fill;
  border?: Border;
  alignment?: Alignment;
  numberFormat?: NumberFormat;
  protection?: { locked?: boolean; hidden?: boolean };
}

// ─── Cell Value ──────────────────────────────────────────────────────────────

export type CellPrimitive = string | number | boolean | Date | null;
export type FormulaValue = { formula: string; result?: CellPrimitive };
export type CellValue = CellPrimitive | FormulaValue;

// ─── Column Definition ───────────────────────────────────────────────────────

export interface ColumnDef {
  /** Header label shown in row 1 */
  header: string;
  /** Unique key used when writing rows via key-value objects */
  key: string;
  /** Column width in character units (default: auto / 15) */
  width?: number;
  /** Style applied to every cell in the column (overridden per-cell) */
  style?: CellStyle;
  /** Style applied only to the header cell */
  headerStyle?: CellStyle;
  /** Whether this column is hidden */
  hidden?: boolean;
  /** Number format shorthand — wraps a `numberFormat` in the column style. */
  format?: NumberFormat;
}

// ─── Merge Region ────────────────────────────────────────────────────────────

export interface MergeRange {
  /** e.g. "A1:C3" or computed from row/col numbers */
  range: string;
  /** Value/formula to write into the top-left cell before merging */
  value?: CellValue;
  /** Style for the merged cell */
  style?: CellStyle;
  /** Row height in points */
  height?: number;
}

// ─── Row ─────────────────────────────────────────────────────────────────────

export type RowData = CellValue[] | Record<string, CellValue>;

export interface RowOptions {
  height?: number;
  style?: CellStyle;
  hidden?: boolean;
  outlineLevel?: number;
}

// ─── Sheet-level Header / Footer ─────────────────────────────────────────────

export interface HeaderFooterSection {
  left?: string;
  center?: string;
  right?: string;
}

export interface SheetHeaderFooter {
  oddHeader?: HeaderFooterSection;
  oddFooter?: HeaderFooterSection;
  evenHeader?: HeaderFooterSection;
  evenFooter?: HeaderFooterSection;
}

// ─── Sheet Options ───────────────────────────────────────────────────────────

export interface SheetOptions {
  name: string;
  /** Default column style */
  defaultColStyle?: CellStyle;
  /** Default row height */
  defaultRowHeight?: number;
  /** Tab color (hex ARGB) */
  tabColor?: string;
  /** Freeze panes */
  freeze?: { row?: number; col?: number };
  /** Print / page setup */
  pageSetup?: {
    paperSize?: number; // 9 = A4, 1 = Letter
    orientation?: "portrait" | "landscape";
    fitToPage?: boolean;
    fitToWidth?: number;
    fitToHeight?: number;
    margins?: {
      left?: number;
      right?: number;
      top?: number;
      bottom?: number;
      header?: number;
      footer?: number;
    };
    /** Rows to repeat at top when printing, e.g. "1:3" */
    printTitlesRow?: string;
    /** Columns to repeat at left when printing, e.g. "A:C" */
    printTitlesColumn?: string;
  };
  headerFooter?: SheetHeaderFooter;
  /** Whether the sheet is protected */
  protection?: { password?: string; sheet?: boolean };
  /** Show/hide grid lines in view */
  showGridLines?: boolean;
  /** Zoom level (10–400) */
  zoom?: number;
  /** Sheet visibility state */
  state?: "visible" | "hidden" | "veryHidden";
}

// ─── Workbook Options ────────────────────────────────────────────────────────

export interface WorkbookOptions {
  /** Author name embedded in the file metadata */
  author?: string;
  /** Company embedded in the file metadata */
  company?: string;
  /** Created date (defaults to now) */
  created?: Date;
  /** Document title */
  title?: string;
  /** Document subject */
  subject?: string;
  /** Document keywords */
  keywords?: string;
  /** Document category */
  category?: string;
  /** Document manager */
  manager?: string;
  /** Document description */
  description?: string;
  /** Document language */
  language?: string;
  /** Whether to use shared strings table (reduces file size for repeated strings) */
  useSharedStrings?: boolean;
  /**
   * Stream mode: write directly to a file path.
   * When set, the workbook operates in streaming mode for
   * memory-efficient output of large datasets.
   */
  useStreaming?: boolean;
  /** Base directory for output path containment (prevents path traversal) */
  allowedBase?: string;
}

// ─── Builder Return Types ────────────────────────────────────────────────────

export interface WriteResult {
  /** Absolute path of the file written */
  filePath: string;
  /** Approximate file size in bytes (-1 when streaming, size reported after flush) */
  sizeBytes: number;
}

// ─── Comment / Note Types ────────────────────────────────────────────────────

export interface NoteText {
  text: string;
  font?: Partial<Font>;
}

export interface NoteConfig {
  texts?: NoteText[];
  margins?: { insetmode?: string; inset?: number[] };
  protection?: { locked?: string; lockText?: string };
  editAs?: string;
  anchor?: string;
}

// ─── Sparkline Types (defined locally; not exported from excelts index) ─────

export interface SparklineColor {
  theme?: number;
  rgb?: string;
  tint?: number;
  auto?: boolean;
}

export interface Sparkline {
  /** Data reference (e.g. "Sheet1!B2:G2") */
  dataRef: string;
  /** Anchor cell reference (e.g. "H2") */
  cellRef: string;
}

export interface SparklineGroup {
  type?: SparklineType;
  lineWeight?: number;
  markers?: boolean;
  high?: boolean;
  low?: boolean;
  first?: boolean;
  last?: boolean;
  negative?: boolean;
  displayXAxis?: boolean;
  displayHidden?: boolean;
  minAxisType?: SparklineAxisType;
  maxAxisType?: SparklineAxisType;
  manualMin?: number;
  manualMax?: number;
  rightToLeft?: boolean;
  colorSeries?: SparklineColor;
  colorNegative?: SparklineColor;
  colorAxis?: SparklineColor;
  colorMarkers?: SparklineColor;
  colorFirst?: SparklineColor;
  colorLast?: SparklineColor;
  colorHigh?: SparklineColor;
  colorLow?: SparklineColor;
  dateAxis?: string;
  sparklines: Sparkline[];
}

export type SparklineType = "line" | "column" | "stacked";

export type SparklineAxisType = "individual" | "group" | "custom";

export interface AddSparklineGroupOptions {
  type: SparklineType;
  sparklines: Sparkline[];
  lineWeight?: number;
  markers?: boolean;
  high?: boolean;
  low?: boolean;
  first?: boolean;
  last?: boolean;
  negative?: boolean;
  lineColor?: string;
  negativeColor?: string;
  axisColor?: string;
  markerColor?: string;
  highColor?: string;
  lowColor?: string;
  firstColor?: string;
  lastColor?: string;
  minAxisType?: SparklineAxisType;
  maxAxisType?: SparklineAxisType;
  manualMin?: number;
  manualMax?: number;
  displayXAxis?: boolean;
  rightToLeft?: boolean;
  displayEmptyCellsAs?: "gap" | "zero" | "span";
  dateAxis?: string;
}

// ─── Ignored Errors ──────────────────────────────────────────────────────────

export interface IgnoredErrorDef {
  /** Cell reference range, e.g. "A1:B10" */
  ref: string;
  /** Ignore "Number Stored as Text" errors */
  numberStoredAsText?: boolean;
  /** Ignore formula errors */
  formula?: boolean;
  /** Ignore formula range errors */
  formulaRange?: boolean;
  /** Ignore unlocked formula errors */
  unlockedFormula?: boolean;
  /** Ignore empty cell reference errors */
  emptyCellReference?: boolean;
  /** Ignore list data validation errors */
  listDataValidation?: boolean;
  /** Ignore calculated column errors */
  calculatedColumn?: boolean;
  /** Ignore eval errors */
  evalError?: boolean;
  /** Ignore two-digit text year errors */
  twoDigitTextYear?: boolean;
}

// ─── SheetBuilder Extension Interface ────────────────────────────────────────

import type {
  AddImageRange,
  CellIsOperators,
  ConditionalFormattingOptions,
  ContainsTextOperators,
  DataValidation,
  DataValidationWithFormulae,
  FormCheckbox as ExcelFormCheckbox,
  Style as ExcelStyle,
  FormCheckboxOptions,
  FormControlRange,
  IconSetTypes,
  RichText as RichTextRun,
  ThreadedComment,
  TimePeriodTypes,
  WatermarkOptions,
} from "@cj-tech-master/excelts";
import type {
  AddBarChartOptions,
  AddChartExOptions,
  AddChartOptions,
  AddChartRange,
  AddComboChartOptions,
  AddPieChartOptions,
  AddScatterChartOptions,
  AddSurfaceChartOptions,
} from "@cj-tech-master/excelts/chart";

export interface SheetBuilderExtension {
  addChart(options: AddChartOptions, range: AddChartRange): unknown;
  addColumnChart(
    options: Omit<AddBarChartOptions, "type" | "barDir">,
    range: AddChartRange,
  ): unknown;
  addBarChart(options: Omit<AddBarChartOptions, "type" | "barDir">, range: AddChartRange): unknown;
  addLineChart(options: Omit<AddChartOptions, "type">, range: AddChartRange): unknown;
  addAreaChart(options: Omit<AddChartOptions, "type">, range: AddChartRange): unknown;
  addPieChart(options: Omit<AddPieChartOptions, "type">, range: AddChartRange): unknown;
  addDoughnutChart(options: Omit<AddPieChartOptions, "type">, range: AddChartRange): unknown;
  addScatterChart(options: Omit<AddScatterChartOptions, "type">, range: AddChartRange): unknown;
  addBubbleChart(options: Omit<AddChartOptions, "type">, range: AddChartRange): unknown;
  addRadarChart(options: Omit<AddChartOptions, "type">, range: AddChartRange): unknown;
  addStockChart(options: Omit<AddChartOptions, "type">, range: AddChartRange): unknown;
  addSurfaceChart(options: Omit<AddSurfaceChartOptions, "type">, range: AddChartRange): unknown;
  addHistogramChart(options: Omit<AddChartExOptions, "type">, range: AddChartRange): unknown;
  addParetoChart(options: Omit<AddChartExOptions, "type">, range: AddChartRange): unknown;
  addWaterfallChart(options: Omit<AddChartExOptions, "type">, range: AddChartRange): unknown;
  addFunnelChart(options: Omit<AddChartExOptions, "type">, range: AddChartRange): unknown;
  addTreemapChart(options: Omit<AddChartExOptions, "type">, range: AddChartRange): unknown;
  addSunburstChart(options: Omit<AddChartExOptions, "type">, range: AddChartRange): unknown;
  addBoxWhiskerChart(options: Omit<AddChartExOptions, "type">, range: AddChartRange): unknown;
  addRegionMapChart(options: Omit<AddChartExOptions, "type">, range: AddChartRange): unknown;
  addComboChart(options: AddComboChartOptions, range: AddChartRange): unknown;
  addConditionalFormatting(cf: ConditionalFormattingOptions): unknown;
  removeConditionalFormatting(
    filter?:
      | number
      | ((
          value: ConditionalFormattingOptions,
          index: number,
          array: ConditionalFormattingOptions[],
        ) => boolean),
  ): unknown;
  addCellIsRule(
    range: CellRange,
    operator: CellIsOperators,
    formulae: (string | number)[],
    style?: Partial<ExcelStyle>,
  ): unknown;
  addExpressionRule(range: CellRange, formula: string, style?: Partial<ExcelStyle>): unknown;
  addDataBar(range: CellRange, color?: { argb?: string; theme?: number }): unknown;
  addColorScale(
    range: CellRange,
    cfvo: {
      type: "min" | "max" | "num" | "percent" | "percentile" | "formula";
      value?: number | string;
    }[],
    colors?: { argb?: string; theme?: number }[],
  ): unknown;
  addIconSet(
    range: CellRange,
    iconSet?: IconSetTypes,
    cfvo?: { type: "percent" | "num" | "percentile" | "formula"; value?: number | string }[],
    options?: { showValue?: boolean; reverse?: boolean },
  ): unknown;
  addTop10Rule(
    range: CellRange,
    rank: number,
    options?: { percent?: boolean; bottom?: boolean; style?: Partial<ExcelStyle> },
  ): unknown;
  addAboveAverageRule(
    range: CellRange,
    options?: { aboveAverage?: boolean; style?: Partial<ExcelStyle> },
  ): unknown;
  addContainsTextRule(
    range: CellRange,
    text: string,
    operator?: ContainsTextOperators,
    style?: Partial<ExcelStyle>,
  ): unknown;
  addTimePeriodRule(
    range: CellRange,
    timePeriod: TimePeriodTypes,
    style?: Partial<ExcelStyle>,
  ): unknown;
  addDataValidation(addr: Addr, validation: DataValidation): unknown;
  removeDataValidation(address: string): unknown;
  addListValidation(
    range: CellRange,
    list: (string | number | Date)[],
    options?: Omit<DataValidationWithFormulae, "type" | "formulae" | "operator">,
  ): unknown;
  addRangeValidation(range: CellRange, validation: RangeValidationDef): unknown;
  addNote(addr: Addr, text: string): unknown;
  addNote(addr: Addr, config: NoteConfig): unknown;
  addThreadedComment(ref: string, comment: ThreadedComment): unknown;
  setCellHyperlink(addr: Addr, hyperlink: string, text?: string, tooltip?: string): unknown;
  setCellRichText(addr: Addr, richText: RichTextRun[]): unknown;
  addImage(imageId: string | number, range: AddImageRange): unknown;
  addBackgroundImage(imageId: string | number): unknown;
  addWatermark(options: WatermarkOptions): unknown;
  removeWatermark(): unknown;
  addSparklineGroup(options: AddSparklineGroupOptions): unknown;
  addFormCheckbox(range: FormControlRange, options?: FormCheckboxOptions): unknown;
  getFormCheckboxes(): ExcelFormCheckbox[];
  // ── Format mixin ─────────────────────────────────────────────────────
  freeze: (rowOrOpts: number | { row?: number; col?: number }, maybeCol?: number) => unknown;
  autoFilter: (range?: string) => unknown;
  autoFitColumns: (startCol?: number | string, endCol?: number | string) => unknown;
  rowHeight: (rowNumber: number, height: number) => unknown;
  colWidth: (col: string | number, width: number) => unknown;
  merge: (
    rangeOrRegion: string | MergeRange,
    options?: { value?: CellValue; style?: CellStyle; height?: number },
  ) => unknown;
  mergeAll: (regions: MergeRange[]) => unknown;
  styleRange: (range: CellRange, style: CellStyle) => unknown;
  clear: () => unknown;
  // ── Structure mixin ──────────────────────────────────────────────────
  insertRow: (pos: number, data: unknown[], options?: RowOptions) => unknown;
  duplicateRow: (rowNum: number, count: number, insert?: boolean) => unknown;
  removeRow: (start: number, count: number) => unknown;
  insertColumn: (start: number, count: number, ...inserts: unknown[][]) => unknown;
  removeColumn: (start: number, count: number) => unknown;
  addPageBreak: (rowNum: number) => unknown;
  addColumnPageBreak: (colNum: number) => unknown;
  // ── Data-io mixin ────────────────────────────────────────────────────
  eachRow: (callback: (row: unknown, rowNumber: number) => void) => void;
  toJSON: (opts?: unknown) => unknown;
  toAOA: () => unknown[][];
  addJSON: (data: Record<string, unknown>[], opts?: unknown) => unknown;
  addAOA: (data: unknown[][], opts?: unknown) => unknown;
}

// ─── External Workbook References ────────────────────────────────────────────

export interface ExternalLinkInput {
  /** Path or URL to external workbook */
  target: string;
  /** Exposed sheet names in the external workbook */
  sheetNames?: string[];
  /** Cached values for formula references */
  cachedValues?: Record<string, Record<string, string | number | boolean | null>>;
  /** Link mode (default: "External") */
  targetMode?: "External" | "Internal";
}
