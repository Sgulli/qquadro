import { describe, expect, it } from "vitest";
import { WorkbookBuilder } from "../workbook-builder.js";

describe("SheetBuilder.defineColumns()", () => {
  it("returns a ColumnMap with correct indices", () => {
    const wb = new WorkbookBuilder();
    let scoreIndex = 0;
    let nameIndex = 0;
    wb.addSheet({ name: "Test" }, (sheet) => {
      const cols = sheet.defineColumns({
        name: { header: "Name", width: 20 },
        score: { header: "Score", width: 15 },
      });
      nameIndex = cols.name.index();
      scoreIndex = cols.score.index();
    });
    expect(nameIndex).toBe(1);
    expect(scoreIndex).toBe(2);
  });

  it("returns correct column letters", () => {
    const wb = new WorkbookBuilder();
    let nameLetter = "";
    let scoreLetter = "";
    wb.addSheet({ name: "Test" }, (sheet) => {
      const cols = sheet.defineColumns({
        name: { header: "Name" },
        score: { header: "Score" },
      });
      nameLetter = cols.name.letter();
      scoreLetter = cols.score.letter();
    });
    expect(nameLetter).toBe("A");
    expect(scoreLetter).toBe("B");
  });

  it("computes range without headers", () => {
    const wb = new WorkbookBuilder();
    let scoreRange = "";
    wb.addSheet({ name: "Test" }, (sheet) => {
      const cols = sheet.defineColumns({
        name: { header: "Name" },
        score: { header: "Score" },
      });
      sheet.addRows([
        { name: "Alice", score: 85 },
        { name: "Bob", score: 92 },
      ]);
      scoreRange = cols.score.range();
    });
    expect(scoreRange).toBe("B1:B2");
  });

  it("computes range with headers", () => {
    const wb = new WorkbookBuilder();
    let scoreRange = "";
    wb.addSheet({ name: "Test" }, (sheet) => {
      const cols = sheet.defineColumns({
        name: { header: "Name" },
        score: { header: "Score" },
      });
      sheet.headers([cols.name.toColumnDef(), cols.score.toColumnDef()]).addRows([
        { name: "Alice", score: 85 },
        { name: "Bob", score: 92 },
      ]);
      scoreRange = cols.score.range();
    });
    expect(scoreRange).toBe("B2:B3");
  });

  it("computes explicit range with custom rows", () => {
    const wb = new WorkbookBuilder();
    let scoreRange = "";
    wb.addSheet({ name: "Test" }, (sheet) => {
      const cols = sheet.defineColumns({
        score: { header: "Score" },
      });
      scoreRange = cols.score.range(5, 20);
    });
    expect(scoreRange).toBe("A5:A20");
  });

  it("returns cell references", () => {
    const wb = new WorkbookBuilder();
    let cell = "";
    wb.addSheet({ name: "Test" }, (sheet) => {
      const cols = sheet.defineColumns({
        score: { header: "Score" },
      });
      cell = cols.score.cell(10);
    });
    expect(cell).toBe("A10");
  });

  it("works with RangeBuilder via range()", async () => {
    const wb = new WorkbookBuilder();
    wb.addSheet({ name: "Test" }, (sheet) => {
      const cols = sheet.defineColumns({
        name: { header: "Name", width: 20 },
        score: { header: "Score", width: 15 },
      });
      sheet
        .headers([cols.name.toColumnDef(), cols.score.toColumnDef()])
        .addRows([
          { name: "Alice", score: 85 },
          { name: "Bob", score: 92 },
          { name: "Carol", score: 78 },
        ])
        .range(cols.score.range())
        .rangeValidation({
          type: "whole",
          operator: "between",
          formulae: [0, 100],
        });
    });
    const result = await wb.toBuffer();
    expect(result).toBeInstanceOf(Buffer);
    expect(result.length).toBeGreaterThan(0);
  });

  it("preserves column key", () => {
    const wb = new WorkbookBuilder();
    let key = "";
    wb.addSheet({ name: "Test" }, (sheet) => {
      const cols = sheet.defineColumns({
        revenue: { header: "Revenue" },
      });
      key = cols.revenue.key;
    });
    expect(key).toBe("revenue");
  });
});
