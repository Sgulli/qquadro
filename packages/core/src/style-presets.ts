import type {
  Cell,
  Alignment as ExcelAlignment,
  Border as ExcelBorder,
  Borders as ExcelBorders,
  Fill as ExcelFill,
  Font as ExcelFont,
} from "@cj-tech-master/excelts";
import type { Color } from "./types.js";
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

function normalizeArgb(hex: string): string {
  const clean = hex.replace(/^#/, "");
  if (clean.length === 6) return `FF${clean}`;
  if (clean.length === 8) return clean;
  throw new Error(`Invalid colour "${hex}" — must be 6-char RGB or 8-char ARGB hex.`);
}

export function color(argb: string): Partial<Color> {
  normalizeArgb(argb);
  return { argb };
}

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

export function resolveNumberFormat(fmt: NumberFormat): string {
  return BUILTIN_FORMATS[fmt as keyof typeof BUILTIN_FORMATS] ?? fmt;
}

export function applyStyle(cell: Cell, style: CellStyle | undefined): void {
  if (!style) return;
  if (style.font) cell.font = toExcelFont(style.font);
  if (style.fill) cell.fill = toExcelFill(style.fill);
  if (style.border) cell.border = toExcelBorder(style.border);
  if (style.alignment) cell.alignment = toExcelAlignment(style.alignment);
  if (style.numberFormat) cell.numFmt = resolveNumberFormat(style.numberFormat);
  if (style.protection) cell.protection = style.protection;
}

export function formatHeaderFooterSection(section: HeaderFooterSection): string {
  return [
    section.left && `&L${section.left}`,
    section.center && `&C${section.center}`,
    section.right && `&R${section.right}`,
  ]
    .filter(Boolean)
    .join("");
}

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

export const border = {
  thinBlack: borderPart("thin", "FF000000"),
  thinGrey: borderPart("thin", "FFBFBFBF"),
  mediumBlack: borderPart("medium", "FF000000"),
  thin(color: string): CellStyle {
    return borderPart("thin", color);
  },
  all(styleName: BorderStyle, color: string): CellStyle {
    return borderPart(styleName, color);
  },
};

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

export function currency(symbol: string): string {
  return `"${symbol}"#,##0.00`;
}

export function accounting(symbol: string): string {
  return `"${symbol}"#,##0.00;("${symbol}"#,##0.00);"-"`;
}

export function font(opts: Font): CellStyle {
  return { font: opts };
}

export function fill(opts: Fill): CellStyle {
  return { fill: opts };
}

export function numFmt(opts: NumberFormat): CellStyle {
  return { numberFormat: opts };
}

export const Styles = {
  header: {
    font: { bold: true, color: { argb: "FFFFFFFF" }, name: "Arial", size: 11 },
    fill: { type: "solid", color: "FF2B579A" },
    alignment: { horizontal: "center", vertical: "middle" },
    border: { bottom: { style: "medium", color: "FF1A3A6B" } },
  } as const satisfies CellStyle,

  subHeader: {
    font: { bold: true, name: "Arial", size: 10, color: { argb: "FF1F497D" } },
    fill: { type: "solid", color: "FFDCE6F1" },
    alignment: { horizontal: "left", vertical: "middle" },
  } as const satisfies CellStyle,

  totalRow: {
    font: { bold: true, name: "Arial", size: 10 },
    fill: { type: "solid", color: "FFFFF2CC" },
    border: {
      top: { style: "thin", color: "FFBFBFBF" },
      bottom: { style: "double", color: "FF000000" },
    },
  } as const satisfies CellStyle,

  currency: {
    numberFormat: "currency",
    font: { name: "Arial", size: 10 },
    alignment: { horizontal: "right" },
  } as const satisfies CellStyle,

  percent: {
    numberFormat: "percentDecimal",
    font: { name: "Arial", size: 10 },
    alignment: { horizontal: "right" },
  } as const satisfies CellStyle,

  date: {
    numberFormat: "date",
    font: { name: "Arial", size: 10 },
    alignment: { horizontal: "center" },
  } as const satisfies CellStyle,

  boxBorder: {
    border: {
      top: { style: "thin", color: "FFBFBFBF" },
      bottom: { style: "thin", color: "FFBFBFBF" },
      left: { style: "thin", color: "FFBFBFBF" },
      right: { style: "thin", color: "FFBFBFBF" },
    },
  } as const satisfies CellStyle,

  inputCell: {
    font: { color: { argb: "FF0000FF" }, name: "Arial", size: 10 },
  } as const satisfies CellStyle,

  formulaCell: {
    font: { color: { argb: "FF000000" }, name: "Arial", size: 10 },
  } as const satisfies CellStyle,

  linkCell: {
    font: { color: { argb: "FF008000" }, name: "Arial", size: 10 },
  } as const satisfies CellStyle,
} as const;
