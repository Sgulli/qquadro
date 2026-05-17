// ─── Alignment ───────────────────────────────────────────────────────────────

export type HorizontalAlignment =
  | "left"
  | "center"
  | "right"
  | "fill"
  | "justify"
  | "centerContinuous"
  | "distributed";

export type VerticalAlignment =
  | "top"
  | "middle"
  | "bottom"
  | "distributed"
  | "justify";

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

export interface Font {
  name?: string;
  size?: number;
  bold?: boolean;
  italic?: boolean;
  underline?:
    | boolean
    | "single"
    | "double"
    | "singleAccounting"
    | "doubleAccounting";
  strike?: boolean;
  color?: string; // hex ARGB
  vertAlign?: "superscript" | "subscript";
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
  currency: "$#,##0.00",
  accountingUSD: '$#,##0.00;($#,##0.00);"-"',
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
}

// ─── Builder Return Types ────────────────────────────────────────────────────

export interface WriteResult {
  /** Absolute path of the file written */
  filePath: string;
  /** Approximate file size in bytes (-1 when streaming, size reported after flush) */
  sizeBytes: number;
}
