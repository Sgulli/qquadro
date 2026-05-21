import { describe, expect, it } from "vitest";
import {
  accounting,
  align,
  applyStyle,
  border,
  colLetter,
  currency,
  fill,
  font,
  formatHeaderFooterSection,
  numFmt,
  resolveNumberFormat,
  Styles,
  style,
  toExcelAlignment,
  toExcelBorder,
  toExcelFill,
  toExcelFont,
} from "../utils.js";

describe("colLetter", () => {
  it("returns empty string for 0", () => {
    expect(colLetter(0)).toBe("");
  });

  it("returns A for 1, Z for 26", () => {
    expect(colLetter(1)).toBe("A");
    expect(colLetter(26)).toBe("Z");
  });

  it("wraps to AA, AB, ZZ", () => {
    expect(colLetter(27)).toBe("AA");
    expect(colLetter(28)).toBe("AB");
    expect(colLetter(702)).toBe("ZZ");
  });

  it("handles three-letter columns", () => {
    expect(colLetter(703)).toBe("AAA");
  });
});

describe("toExcelFont", () => {
  it("returns empty object for empty input", () => {
    expect(toExcelFont({})).toEqual({});
  });

  it("maps basic font properties", () => {
    expect(toExcelFont({ bold: true, italic: true, size: 12, name: "Arial" })).toMatchObject({
      bold: true,
      italic: true,
      size: 12,
      name: "Arial",
    });
  });

  it("normalizes ARGB color", () => {
    expect(toExcelFont({ color: "FF0000" })).toEqual({
      color: { argb: "FFFF0000" },
    });
  });

  it("preserves ARGB color when 8-char", () => {
    expect(toExcelFont({ color: "FF2B579A" })).toEqual({
      color: { argb: "FF2B579A" },
    });
  });

  it("handles underline and strike", () => {
    expect(toExcelFont({ underline: "double", strike: true })).toMatchObject({
      underline: "double",
      strike: true,
    });
  });
});

describe("toExcelFill", () => {
  it("converts solid fill to pattern fill", () => {
    expect(toExcelFill({ type: "solid", color: "FF2B579A" })).toEqual({
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF2B579A" },
    });
  });

  it("strips hash from color", () => {
    expect(toExcelFill({ type: "solid", color: "#FF0000" })).toEqual({
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFF0000" },
    });
  });

  it("converts angle gradient", () => {
    const result = toExcelFill({
      type: "gradient",
      gradient: "angle",
      degree: 45,
      stops: [
        { position: 0, color: "FFFFFF" },
        { position: 1, color: "000000" },
      ],
    });
    expect(result).toEqual({
      type: "gradient",
      gradient: "angle",
      degree: 45,
      stops: [
        { position: 0, color: { argb: "FFFFFFFF" } },
        { position: 1, color: { argb: "FF000000" } },
      ],
    });
  });

  it("converts path gradient", () => {
    const result = toExcelFill({
      type: "gradient",
      gradient: "path",
      stops: [{ position: 0, color: "FFFFFF" }],
    });
    expect(result).toEqual({
      type: "gradient",
      gradient: "path",
      center: { left: 0.5, top: 0.5 },
      stops: [{ position: 0, color: { argb: "FFFFFFFF" } }],
    });
  });

  it("throws on invalid colour", () => {
    expect(() => toExcelFill({ type: "solid", color: "XYZ" })).toThrow("Invalid colour");
  });
});

describe("toExcelBorder", () => {
  it("handles empty border", () => {
    expect(toExcelBorder({})).toEqual({});
  });

  it("converts all four sides", () => {
    const result = toExcelBorder({
      top: { style: "thin", color: "000000" },
      bottom: { style: "medium", color: "FF000000" },
      left: { style: "thick", color: "FFFFFF" },
      right: { style: "double" },
    });
    expect(result.top).toEqual({ style: "thin", color: { argb: "FF000000" } });
    expect(result.bottom).toEqual({ style: "medium", color: { argb: "FF000000" } });
    expect(result.left).toEqual({ style: "thick", color: { argb: "FFFFFFFF" } });
    expect(result.right).toEqual({ style: "double", color: undefined });
  });

  it("converts diagonal border", () => {
    const result = toExcelBorder({
      diagonal: { style: "thin", color: "000000", up: true, down: false },
    });
    expect(result.diagonal).toMatchObject({
      style: "thin",
      up: true,
      down: false,
    });
  });
});

describe("toExcelAlignment", () => {
  it("maps all alignment properties", () => {
    expect(
      toExcelAlignment({
        horizontal: "center",
        vertical: "middle",
        wrapText: true,
        shrinkToFit: true,
        indent: 2,
        textRotation: 45,
      }),
    ).toEqual({
      horizontal: "center",
      vertical: "middle",
      wrapText: true,
      shrinkToFit: true,
      indent: 2,
      textRotation: 45,
    });
  });

  it("returns empty for empty alignment", () => {
    expect(toExcelAlignment({})).toEqual({});
  });
});

describe("resolveNumberFormat", () => {
  it("resolves builtin format names", () => {
    expect(resolveNumberFormat("currency")).toBe("#,##0.00");
    expect(resolveNumberFormat("percent")).toBe("0%");
    expect(resolveNumberFormat("date")).toBe("yyyy-mm-dd");
    expect(resolveNumberFormat("general")).toBe("General");
  });

  it("passes through custom format strings", () => {
    expect(resolveNumberFormat("#,##0.00")).toBe("#,##0.00");
  });
});

describe("formatHeaderFooterSection", () => {
  it("formats all three sections", () => {
    expect(formatHeaderFooterSection({ left: "Title", center: "&D", right: "Page &P" })).toBe(
      "&LTitle&C&D&RPage &P",
    );
  });

  it("omits empty sections", () => {
    expect(formatHeaderFooterSection({ left: "Only Left" })).toBe("&LOnly Left");
    expect(formatHeaderFooterSection({})).toBe("");
  });
});

describe("style()", () => {
  it("merges font properties", () => {
    const result = style({ font: { bold: true } }, { font: { italic: true } });
    expect(result.font).toEqual({ bold: true, italic: true });
  });

  it("later parts override earlier parts", () => {
    const result = style({ font: { bold: true, color: "FFFFFF" } }, { font: { bold: false } });
    expect(result.font).toEqual({ bold: false, color: "FFFFFF" });
  });

  it("fill is replaced, not merged", () => {
    const result = style(
      { fill: { type: "solid", color: "FFFFFF" } },
      { fill: { type: "solid", color: "000000" } },
    );
    expect(result.fill).toEqual({ type: "solid", color: "000000" });
  });

  it("merges alignment, border, protection", () => {
    const result = style(
      { alignment: { horizontal: "left" } },
      { alignment: { vertical: "middle" } },
      { border: { top: { style: "thin" } } },
      { protection: { locked: false } },
    );
    expect(result.alignment).toEqual({ horizontal: "left", vertical: "middle" });
    expect(result.border).toEqual({ top: { style: "thin" } });
    expect(result.protection).toEqual({ locked: false });
  });

  it("handles single part", () => {
    expect(style({ font: { bold: true } })).toEqual({ font: { bold: true } });
  });

  it("handles empty parts", () => {
    expect(style()).toEqual({});
  });
});

describe("border presets", () => {
  it("thinBlack has thin black on all four sides", () => {
    const b = border.thinBlack.border;
    expect(b?.top).toEqual({ style: "thin", color: "FF000000" });
    expect(b?.bottom).toEqual(b?.top);
    expect(b?.left).toEqual(b?.top);
    expect(b?.right).toEqual(b?.top);
  });

  it("thinGrey uses grey color", () => {
    expect(border.thinGrey.border?.top?.color).toBe("FFBFBFBF");
  });

  it("mediumBlack uses medium style", () => {
    expect(border.mediumBlack.border?.top?.style).toBe("medium");
  });

  it("thin() factory creates custom color", () => {
    const b = border.thin("FF00FF00").border;
    expect(b?.top?.color).toBe("FF00FF00");
    expect(b?.top?.style).toBe("thin");
  });

  it("all() factory creates custom style + color", () => {
    const b = border.all("dashed", "FF0000FF").border;
    expect(b?.top?.style).toBe("dashed");
    expect(b?.top?.color).toBe("FF0000FF");
  });
});

describe("align presets", () => {
  it("center", () => {
    expect(align.center.alignment).toMatchObject({ horizontal: "center", vertical: "middle" });
  });

  it("centerWrap", () => {
    expect(align.centerWrap.alignment.wrapText).toBe(true);
  });

  it("left, leftWrap, right, rightWrap all have matching orientation", () => {
    expect(align.left.alignment.horizontal).toBe("left");
    expect(align.leftWrap.alignment.horizontal).toBe("left");
    expect(align.leftWrap.alignment.wrapText).toBe(true);
    expect(align.right.alignment.horizontal).toBe("right");
    expect(align.rightWrap.alignment.wrapText).toBe(true);
  });
});

describe("format helpers", () => {
  it("currency() wraps symbol", () => {
    expect(currency("$")).toBe('"$"#,##0.00');
    expect(currency("€")).toBe('"€"#,##0.00');
  });

  it("accounting() wraps symbol with parens for negatives", () => {
    expect(accounting("$")).toBe('"$"#,##0.00;("$"#,##0.00);"-"');
  });
});

describe("factory wrappers", () => {
  it("font() creates cell style with only font", () => {
    expect(font({ bold: true, size: 12 })).toEqual({ font: { bold: true, size: 12 } });
  });

  it("fill() creates cell style with only fill", () => {
    expect(fill({ type: "solid", color: "FF0000" })).toEqual({
      fill: { type: "solid", color: "FF0000" },
    });
  });

  it("numFmt() creates cell style with only numberFormat", () => {
    expect(numFmt("percent")).toEqual({ numberFormat: "percent" });
  });
});

describe("Styles presets", () => {
  it("header has blue fill and white font", () => {
    expect(Styles.header.fill).toEqual({ type: "solid", color: "FF2B579A" });
    expect(Styles.header.font).toMatchObject({ bold: true, color: { argb: "FFFFFFFF" } });
  });

  it("totalRow has yellow fill and double bottom border", () => {
    expect(Styles.totalRow.fill).toEqual({ type: "solid", color: "FFFFF2CC" });
    expect(Styles.totalRow.border?.bottom?.style).toBe("double");
  });

  it("currency has right alignment", () => {
    expect(Styles.currency.alignment?.horizontal).toBe("right");
  });

  it("percent uses percentDecimal format", () => {
    expect(Styles.percent.numberFormat).toBe("percentDecimal");
  });

  it("all preset styles exist", () => {
    expect(Styles.header).toBeDefined();
    expect(Styles.subHeader).toBeDefined();
    expect(Styles.totalRow).toBeDefined();
    expect(Styles.currency).toBeDefined();
    expect(Styles.percent).toBeDefined();
    expect(Styles.date).toBeDefined();
    expect(Styles.boxBorder).toBeDefined();
    expect(Styles.inputCell).toBeDefined();
    expect(Styles.formulaCell).toBeDefined();
    expect(Styles.linkCell).toBeDefined();
  });
});

describe("applyStyle", () => {
  it("does nothing for undefined style", () => {
    const cell = { font: {} } as Record<string, unknown>;
    applyStyle(cell as never, undefined);
    expect(cell.font).toEqual({});
  });

  it("applies font, fill, border, alignment, numberFormat, protection", () => {
    const cell: Record<string, unknown> = {};
    applyStyle(cell as never, {
      font: { bold: true },
      fill: { type: "solid", color: "FFFFFF" },
      border: { top: { style: "thin" } },
      alignment: { horizontal: "center" },
      numberFormat: "currency",
      protection: { locked: true },
    });
    expect(cell.font).toBeDefined();
    expect(cell.fill).toBeDefined();
    expect(cell.border).toBeDefined();
    expect(cell.alignment).toBeDefined();
    expect(cell.numFmt).toBe("#,##0.00");
    expect(cell.protection).toEqual({ locked: true });
  });

  it("skips undefined style sub-objects", () => {
    const cell: Record<string, unknown> = {};
    applyStyle(cell as never, {});
    expect(cell.font).toBeUndefined();
  });
});
