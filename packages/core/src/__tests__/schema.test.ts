import { describe, expect, it } from "vitest";
import { defineSheet } from "../define-sheet.js";
import { Schema, validateRow } from "../schema.js";
import { WorkbookBuilder } from "../workbook-builder.js";

describe("Schema field builders", () => {
  it("creates text field", () => {
    const f = Schema.text({ width: 25 });
    expect(f.type).toBe("text");
    expect(f.width).toBe(25);
  });

  it("creates number field", () => {
    const f = Schema.number({ min: 0, max: 100 });
    expect(f.type).toBe("number");
    expect(f.min).toBe(0);
    expect(f.max).toBe(100);
  });

  it("creates date field", () => {
    const f = Schema.date();
    expect(f.type).toBe("date");
  });

  it("creates boolean field", () => {
    const f = Schema.boolean();
    expect(f.type).toBe("boolean");
  });

  it("creates enum field", () => {
    const f = Schema.enum(["A", "B", "C"] as const);
    expect(f.type).toBe("enum");
    expect(f.values).toEqual(["A", "B", "C"]);
  });

  it("creates currency field", () => {
    const f = Schema.currency({ width: 20 });
    expect(f.type).toBe("currency");
    expect(f.width).toBe(20);
  });

  it("creates percent field", () => {
    const f = Schema.percent();
    expect(f.type).toBe("percent");
  });
});

describe("validateRow()", () => {
  const schema = {
    name: Schema.text(),
    score: Schema.number({ min: 0, max: 100 }),
    grade: Schema.enum(["A", "B", "C", "D", "F"] as const),
  };

  it("passes for valid row", () => {
    const errors = validateRow({ name: "Alice", score: 85, grade: "A" }, schema, 1);
    expect(errors).toEqual([]);
  });

  it("catches wrong type for text", () => {
    const errors = validateRow({ name: 42, score: 85, grade: "A" }, schema, 1);
    expect(errors.length).toBe(1);
    expect(errors[0]).toContain("expected string");
  });

  it("catches wrong type for number", () => {
    const errors = validateRow({ name: "Alice", score: "85", grade: "A" }, schema, 1);
    expect(errors.length).toBe(1);
    expect(errors[0]).toContain("expected number");
  });

  it("catches number below min", () => {
    const errors = validateRow({ name: "Alice", score: -5, grade: "A" }, schema, 1);
    expect(errors.length).toBe(1);
    expect(errors[0]).toContain("below min");
  });

  it("catches number above max", () => {
    const errors = validateRow({ name: "Alice", score: 150, grade: "A" }, schema, 1);
    expect(errors.length).toBe(1);
    expect(errors[0]).toContain("above max");
  });

  it("catches invalid enum value", () => {
    const errors = validateRow({ name: "Alice", score: 85, grade: "X" }, schema, 1);
    expect(errors.length).toBe(1);
    expect(errors[0]).toContain("must be one of");
  });

  it("skips null/undefined values", () => {
    const errors = validateRow({ name: null, score: undefined, grade: "A" }, schema, 1);
    expect(errors).toEqual([]);
  });
});

describe("defineSheet()", () => {
  it("creates a sheet with headers from schema", async () => {
    const wb = new WorkbookBuilder();
    const { sheet } = defineSheet(wb, "Scores", {
      name: Schema.text({ width: 20 }),
      score: Schema.number({ width: 15 }),
    });

    expect(sheet.name).toBe("Scores");
    const result = await wb.toBuffer();
    expect(result).toBeInstanceOf(Buffer);
    expect(result.length).toBeGreaterThan(0);
  });

  it("adds typed rows via addRows()", async () => {
    const wb = new WorkbookBuilder();
    const { addRows } = defineSheet(wb, "Scores", {
      name: Schema.text(),
      score: Schema.number(),
    });

    addRows([
      { name: "Alice", score: 85 },
      { name: "Bob", score: 92 },
    ]);

    const result = await wb.toBuffer();
    expect(result).toBeInstanceOf(Buffer);
  });

  it("adds a single typed row via addRow()", async () => {
    const wb = new WorkbookBuilder();
    const { addRow } = defineSheet(wb, "Scores", {
      name: Schema.text(),
      score: Schema.number(),
    });

    addRow({ name: "Alice", score: 85 });

    const result = await wb.toBuffer();
    expect(result).toBeInstanceOf(Buffer);
  });

  it("throws on invalid row data", () => {
    const wb = new WorkbookBuilder();
    const { addRow } = defineSheet(wb, "Scores", {
      name: Schema.text(),
      score: Schema.number({ min: 0, max: 100 }),
    });

    expect(() => addRow({ name: "Alice", score: 150 })).toThrow("Validation failed");
  });

  it("throws on wrong type in addRows()", () => {
    const wb = new WorkbookBuilder();
    const { addRows } = defineSheet(wb, "Scores", {
      name: Schema.text(),
      score: Schema.number(),
    });

    expect(() =>
      addRows([
        { name: "Alice", score: 85 },
        { name: 42 as unknown as string, score: 92 },
      ]),
    ).toThrow("Validation failed");
  });

  it("provides column references via columns map", async () => {
    const wb = new WorkbookBuilder();
    const { columns, addRows, sheet } = defineSheet(wb, "Scores", {
      name: Schema.text(),
      score: Schema.number(),
    });

    addRows([
      { name: "Alice", score: 85 },
      { name: "Bob", score: 92 },
    ]);

    expect(columns.name.letter()).toBe("A");
    expect(columns.score.letter()).toBe("B");
    expect(columns.score.index()).toBe(2);
    expect(columns.score.range()).toBe("B2:B3");

    sheet.range(columns.score.range()).rangeValidation({
      type: "whole",
      operator: "between",
      formulae: [0, 100],
    });

    const result = await wb.toBuffer();
    expect(result).toBeInstanceOf(Buffer);
  });

  it("works with enum fields", async () => {
    const wb = new WorkbookBuilder();
    const { addRows } = defineSheet(wb, "Grades", {
      name: Schema.text(),
      grade: Schema.enum(["A", "B", "C", "D", "F"] as const),
    });

    addRows([
      { name: "Alice", grade: "A" },
      { name: "Bob", grade: "B" },
    ]);

    const result = await wb.toBuffer();
    expect(result).toBeInstanceOf(Buffer);
  });

  it("works with currency and percent fields", async () => {
    const wb = new WorkbookBuilder();
    const { addRows } = defineSheet(wb, "Financial", {
      item: Schema.text(),
      amount: Schema.currency(),
      rate: Schema.percent(),
    });

    addRows([
      { item: "Widget", amount: 1500, rate: 0.15 },
      { item: "Gadget", amount: 2500, rate: 0.22 },
    ]);

    const result = await wb.toBuffer();
    expect(result).toBeInstanceOf(Buffer);
  });

  it("creates sheet without headers when writeHeaders=false", async () => {
    const wb = new WorkbookBuilder();
    const { sheet, addRows } = defineSheet(
      wb,
      "Data",
      {
        x: Schema.number(),
        y: Schema.number(),
      },
      { writeHeaders: false },
    );

    addRows([
      { x: 1, y: 2 },
      { x: 3, y: 4 },
    ]);

    expect(sheet.rowCount).toBe(2);
    const result = await wb.toBuffer();
    expect(result).toBeInstanceOf(Buffer);
  });
});
