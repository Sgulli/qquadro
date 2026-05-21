import type {
  Cell,
  Alignment as ExcelAlignment,
  Border as ExcelBorder,
  Borders as ExcelBorders,
  Fill as ExcelFill,
  Font as ExcelFont,
} from "@cj-tech-master/excelts";
import {
  type Alignment,
  type Border,
  type BorderStyle,
  BUILTIN_FORMATS,
  type CellStyle,
  type Fill,
  type Font,
  type HeaderFooterSection,
  type NumberFormat,
} from "./types.js";

// ─── Font ────────────────────────────────────────────────────────────────────

export function toExcelFont(font: Font): Partial<ExcelFont> {
  return {
    name: font.name,
    size: font.size,
    bold: font.bold,
    italic: font.italic,
    underline: font.underline,
    strike: font.strike,
    color: font.color
      ? typeof font.color === "string"
        ? { argb: normalizeArgb(font.color) }
        : { argb: font.color.argb, theme: font.color.theme }
      : undefined,
    vertAlign: font.vertAlign,
  };
}

// ─── Fill ────────────────────────────────────────────────────────────────────

export function toExcelFill(fill: Fill): ExcelFill {
  if (fill.type === "solid") {
    return {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: normalizeArgb(fill.color) },
    };
  }

  const stops = fill.stops.map((s) => ({
    position: s.position,
    color: { argb: normalizeArgb(s.color) },
  }));

  if (fill.gradient === "angle") {
    return {
      type: "gradient",
      gradient: "angle",
      degree: fill.degree ?? 0,
      stops,
    };
  }
  return {
    type: "gradient",
    gradient: "path",
    center: { left: 0.5, top: 0.5 },
    stops,
  };
}

// ─── Border ──────────────────────────────────────────────────────────────────

function toExcelBorderSide(side: NonNullable<Border["top"]>): Partial<ExcelBorder> {
  return {
    style: side.style,
    color: side.color ? { argb: normalizeArgb(side.color) } : undefined,
  };
}

export function toExcelBorder(border: Border): Partial<ExcelBorders> {
  const result: Partial<ExcelBorders> = {};
  if (border.top) result.top = toExcelBorderSide(border.top);
  if (border.bottom) result.bottom = toExcelBorderSide(border.bottom);
  if (border.left) result.left = toExcelBorderSide(border.left);
  if (border.right) result.right = toExcelBorderSide(border.right);
  if (border.diagonal) {
    result.diagonal = {
      ...toExcelBorderSide(border.diagonal),
      up: border.diagonal.up ?? false,
      down: border.diagonal.down ?? false,
    };
  }
  return result;
}

// ─── Alignment ───────────────────────────────────────────────────────────────

export function toExcelAlignment(a: Alignment): Partial<ExcelAlignment> {
  return {
    horizontal: a.horizontal,
    vertical: a.vertical,
    wrapText: a.wrapText,
    shrinkToFit: a.shrinkToFit,
    indent: a.indent,
    textRotation: a.textRotation,
  };
}

// ─── Number Format ───────────────────────────────────────────────────────────

export function resolveNumberFormat(fmt: NumberFormat): string {
  return BUILTIN_FORMATS[fmt as keyof typeof BUILTIN_FORMATS] ?? fmt;
}

// ─── Apply Style ─────────────────────────────────────────────────────────────

export function applyStyle(cell: Cell, style: CellStyle | undefined): void {
  if (!style) return;
  if (style.font) cell.font = toExcelFont(style.font);
  if (style.fill) cell.fill = toExcelFill(style.fill);
  if (style.border) cell.border = toExcelBorder(style.border);
  if (style.alignment) cell.alignment = toExcelAlignment(style.alignment);
  if (style.numberFormat) cell.numFmt = resolveNumberFormat(style.numberFormat);
  if (style.protection) cell.protection = style.protection;
}

// ─── Header / Footer ─────────────────────────────────────────────────────────

/** Builds an Excel header/footer string from left, center, and right sections. */
export function formatHeaderFooterSection(section: HeaderFooterSection): string {
  return [
    section.left && `&L${section.left}`,
    section.center && `&C${section.center}`,
    section.right && `&R${section.right}`,
  ]
    .filter(Boolean)
    .join("");
}

// ─── Helpers ─────────────────────────────────────────────────────────────────

import type { Color } from "./types.js";

/** Create a color from an ARGB hex string (e.g. "FFFFFFFF" for white). */
export function color(argb: string): Partial<Color> {
  return { argb };
}

function normalizeArgb(hex: string): string {
  const clean = hex.replace(/^#/, "");
  if (clean.length === 6) return `FF${clean}`;
  if (clean.length === 8) return clean;
  throw new Error(`Invalid colour "${hex}" — must be 6-char RGB or 8-char ARGB hex.`);
}

export function colLetter(n: number): string {
  let r = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    r = String.fromCharCode(65 + rem) + r;
    n = Math.floor((n - 1) / 26);
  }
  return r;
}

/** Build an A1-style cell reference from 1‑based column and row numbers. */
export function cellRef(col: number, row: number): string {
  return `${colLetter(col)}${row}`;
}

/** Build a range string like `"B2:B10"` from a column index and row range. */
export function colRange(col: number, startRow: number, endRow: number): string {
  const c = colLetter(col);
  return `${c}${startRow}:${c}${endRow}`;
}

/** Build a range string like `"B2:D10"` from two corners. */
export function rangeRef(col1: number, row1: number, col2: number, row2: number): string {
  return `${cellRef(col1, row1)}:${cellRef(col2, row2)}`;
}

// ─── style() compose helper ──────────────────────────────────────────────────

/**
 * Compose a `CellStyle` from multiple partial style objects.
 * Later parts override earlier ones. Sub-objects (font, border, alignment)
 * are deep-merged at the leaf level so you can mix-and-match partials.
 *
 * @example
 * ```ts
 * const myStyle = style(
 *   { font: { bold: true, color: "FFFFFFFF" } },
 *   { fill: { type: "solid", color: "FF2B579A" } },
 *   align.centerWrap,
 *   border.thinBlack,
 * );
 * ```
 */
export function style(...parts: CellStyle[]): CellStyle {
  const out: CellStyle = {};
  for (const p of parts) {
    if (p.font) out.font = { ...out.font, ...p.font };
    if (p.fill) out.fill = p.fill;
    if (p.border) out.border = { ...out.border, ...p.border };
    if (p.alignment) out.alignment = { ...out.alignment, ...p.alignment };
    if (p.numberFormat) out.numberFormat = p.numberFormat;
    if (p.protection) out.protection = { ...out.protection, ...p.protection };
  }
  return out;
}

// ─── border presets ──────────────────────────────────────────────────────────

function borderAll(styleName: BorderStyle, color: string): Border {
  return {
    top: { style: styleName, color },
    bottom: { style: styleName, color },
    left: { style: styleName, color },
    right: { style: styleName, color },
  };
}

function borderPart(styleName: BorderStyle, color: string): CellStyle {
  return { border: borderAll(styleName, color) };
}

/**
 * Pre-built border partials — pass into `style()` or use standalone.
 *
 * @example
 * ```ts
 * style({ font: { bold: true } }, border.thinBlack)
 * ```
 */
export const border = {
  /** Thin black border on all four sides ("FF000000") */
  thinBlack: borderPart("thin", "FF000000"),

  /** Thin grey border on all four sides ("FFBFBFBF") */
  thinGrey: borderPart("thin", "FFBFBFBF"),

  /** Medium black border on all four sides */
  mediumBlack: borderPart("medium", "FF000000"),

  /** Thin border with a custom ARGB color */
  thin(color: string): CellStyle {
    return borderPart("thin", color);
  },

  /** Custom border applied to all four sides */
  all(styleName: BorderStyle, color: string): CellStyle {
    return borderPart(styleName, color);
  },
};

// ─── alignment presets ───────────────────────────────────────────────────────

/**
 * Pre-built alignment partials — pass into `style()` or use standalone.
 *
 * @example
 * ```ts
 * style({ font: { bold: true } }, align.centerWrap)
 * ```
 */
export const align = {
  center: { alignment: { horizontal: "center", vertical: "middle" } } as const satisfies CellStyle,
  centerWrap: {
    alignment: { horizontal: "center", vertical: "middle", wrapText: true },
  } as const satisfies CellStyle,
  left: { alignment: { horizontal: "left", vertical: "middle" } } as const satisfies CellStyle,
  leftWrap: {
    alignment: { horizontal: "left", vertical: "middle", wrapText: true },
  } as const satisfies CellStyle,
  right: { alignment: { horizontal: "right", vertical: "middle" } } as const satisfies CellStyle,
  rightWrap: {
    alignment: { horizontal: "right", vertical: "middle", wrapText: true },
  } as const satisfies CellStyle,
};

// ─── Number format helpers ───────────────────────────────────────────────────

/**
 * Generate a currency number format string for any symbol.
 *
 * @example
 * ```ts
 * style({ numberFormat: currency("€") })
 * style({ numberFormat: currency("$") })
 * style({ numberFormat: accounting("€") })
 * ```
 */
export function currency(symbol: string): string {
  return `"${symbol}"#,##0.00`;
}

/** Accounting-style format (negatives in parens) for any symbol. */
export function accounting(symbol: string): string {
  return `"${symbol}"#,##0.00;("${symbol}"#,##0.00);"-"`;
}

// ─── Style-part wrappers ────────────────────────────────────────────────────

/**
 * Shorthand for `{ font: opts }` — pass into `style()`.
 *
 * @example
 * ```ts
 * style(font({ bold: true, size: 10 }), align.center)
 * ```
 */
export function font(opts: Font): CellStyle {
  return { font: opts };
}

/**
 * Shorthand for `{ fill: opts }` — pass into `style()`.
 *
 * @example
 * ```ts
 * style(fill({ type: "solid", color: "FFDCE6F1" }), border.thinBlack)
 * ```
 */
export function fill(opts: Fill): CellStyle {
  return { fill: opts };
}

/**
 * Shorthand for `{ numberFormat: opts }` — pass into `style()`.
 *
 * @example
 * ```ts
 * style(numFmt(currency("€")), align.right)
 * ```
 */
export function numFmt(opts: NumberFormat): CellStyle {
  return { numberFormat: opts };
}

// ─── Preset Styles ───────────────────────────────────────────────────────────

export const Styles = {
  /** Bold white text on blue — ideal for column headers */
  header: {
    font: { bold: true, color: { argb: "FFFFFFFF" }, name: "Arial", size: 11 },
    fill: { type: "solid", color: "FF2B579A" },
    alignment: { horizontal: "center", vertical: "middle" },
    border: { bottom: { style: "medium", color: "FF1A3A6B" } },
  } as const satisfies CellStyle,

  /** Lighter blue — sub-section headers */
  subHeader: {
    font: { bold: true, name: "Arial", size: 10, color: { argb: "FF1F497D" } },
    fill: { type: "solid", color: "FFDCE6F1" },
    alignment: { horizontal: "left", vertical: "middle" },
  } as const satisfies CellStyle,

  /** Yellow-tinted bold — totals / summary rows */
  totalRow: {
    font: { bold: true, name: "Arial", size: 10 },
    fill: { type: "solid", color: "FFFFF2CC" },
    border: {
      top: { style: "thin", color: "FFBFBFBF" },
      bottom: { style: "double", color: "FF000000" },
    },
  } as const satisfies CellStyle,

  /** Right-aligned currency format */
  currency: {
    numberFormat: "currency",
    font: { name: "Arial", size: 10 },
    alignment: { horizontal: "right" },
  } as const satisfies CellStyle,

  /** Right-aligned percentage format (two decimals) */
  percent: {
    numberFormat: "percentDecimal",
    font: { name: "Arial", size: 10 },
    alignment: { horizontal: "right" },
  } as const satisfies CellStyle,

  /** Centered date format */
  date: {
    numberFormat: "date",
    font: { name: "Arial", size: 10 },
    alignment: { horizontal: "center" },
  } as const satisfies CellStyle,

  /** Thin grey border on all four sides */
  boxBorder: {
    border: {
      top: { style: "thin", color: "FFBFBFBF" },
      bottom: { style: "thin", color: "FFBFBFBF" },
      left: { style: "thin", color: "FFBFBFBF" },
      right: { style: "thin", color: "FFBFBFBF" },
    },
  } as const satisfies CellStyle,

  /** Blue font — financial-model convention for hardcoded inputs */
  inputCell: {
    font: { color: { argb: "FF0000FF" }, name: "Arial", size: 10 },
  } as const satisfies CellStyle,

  /** Black font — financial-model convention for formula cells */
  formulaCell: {
    font: { color: { argb: "FF000000" }, name: "Arial", size: 10 },
  } as const satisfies CellStyle,

  /** Green font — financial-model convention for cross-sheet links */
  linkCell: {
    font: { color: { argb: "FF008000" }, name: "Arial", size: 10 },
  } as const satisfies CellStyle,
} as const;
