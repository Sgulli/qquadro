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
  };
  headerFooter?: SheetHeaderFooter;
  /** Whether the sheet is protected */
  protection?: { password?: string; sheet?: boolean };
  /** Show/hide grid lines in view */
  showGridLines?: boolean;
  /** Zoom level (10–400) */
  zoom?: number;
}

// ─── Workbook Options ────────────────────────────────────────────────────────

export interface WorkbookOptions {
  /** Author name embedded in the file metadata */
  author?: string;
  /** Company embedded in the file metadata */
  company?: string;
  /** Created date (defaults to now) */
  created?: Date;
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

export type ThreadedComment = import("@cj-tech-master/excelts").ThreadedComment;
export type ThreadedCommentPerson = import("@cj-tech-master/excelts").ThreadedCommentPerson;
export type ThreadedCommentMention = import("@cj-tech-master/excelts").ThreadedCommentMention;

// ─── Cell Value Types (re-exported from excelts) ─────────────────────────────

export type RichTextRun = import("@cj-tech-master/excelts").RichText;
export type CellRichTextValue = import("@cj-tech-master/excelts").CellRichTextValue;
export type CellHyperlinkValue = import("@cj-tech-master/excelts").CellHyperlinkValue;
export type CellHyperlinkValueInput = import("@cj-tech-master/excelts").CellHyperlinkValueInput;

// ─── Image Types ─────────────────────────────────────────────────────────────

export type ImageData = import("@cj-tech-master/excelts").ImageData;
export type AddImageRange = import("@cj-tech-master/excelts").AddImageRange;
export type WatermarkOptions = import("@cj-tech-master/excelts").WatermarkOptions;
export type WatermarkMode = import("@cj-tech-master/excelts").WatermarkMode;

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

// ─── Chart Types (from excelts chart sub-module) ────────────────────────────

export type ChartType = import("@cj-tech-master/excelts/chart").ChartType;
export type AddChartRange = import("@cj-tech-master/excelts/chart").AddChartRange;
export type AddChartOptions = import("@cj-tech-master/excelts/chart").AddChartOptions;
export type AddBarChartOptions = import("@cj-tech-master/excelts/chart").AddBarChartOptions;
export type AddPieChartOptions = import("@cj-tech-master/excelts/chart").AddPieChartOptions;
export type AddScatterChartOptions = import("@cj-tech-master/excelts/chart").AddScatterChartOptions;
export type AddSurfaceChartOptions = import("@cj-tech-master/excelts/chart").AddSurfaceChartOptions;
export type AddComboChartOptions = import("@cj-tech-master/excelts/chart").AddComboChartOptions;
export type AddChartExOptions = import("@cj-tech-master/excelts/chart").AddChartExOptions;
export type AddChartSeriesOptions = import("@cj-tech-master/excelts/chart").AddChartSeriesOptions;
export type AddAxisOptions = import("@cj-tech-master/excelts/chart").AddAxisOptions;
export type AddDataLabelsOptions = import("@cj-tech-master/excelts/chart").AddDataLabelsOptions;
export type AddChartMarkerOptions = import("@cj-tech-master/excelts/chart").AddChartMarkerOptions;
export type AddTrendlineOptions = import("@cj-tech-master/excelts/chart").AddTrendlineOptions;
export type AddErrorBarsOptions = import("@cj-tech-master/excelts/chart").AddErrorBarsOptions;
export type AddDataPointOptions = import("@cj-tech-master/excelts/chart").AddDataPointOptions;
export type AddTitleOptions = import("@cj-tech-master/excelts/chart").AddTitleOptions;
export type AddLegendOptions = import("@cj-tech-master/excelts/chart").AddLegendOptions;
export type AddPlotAreaOptions = import("@cj-tech-master/excelts/chart").AddPlotAreaOptions;
export type AddChartFromTableOptions =
  import("@cj-tech-master/excelts/chart").AddChartFromTableOptions;
export type SeriesFromColumnsOptions =
  import("@cj-tech-master/excelts/chart").SeriesFromColumnsOptions;
export type ExcelChartPreset = import("@cj-tech-master/excelts/chart").ExcelChartPreset;
export type ExcelChartExPreset = import("@cj-tech-master/excelts/chart").ExcelChartExPreset;
