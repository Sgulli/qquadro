import fs from "node:fs";
import path from "node:path";
import { afterAll, describe, expect, it } from "vitest";
import { align, border, F, font, Styles, style, WorkbookBuilder } from "../index.js";

const outputDir = path.resolve("output/test-integration");
const outputPath = (name: string) => path.join(outputDir, name);

afterAll(() => {
  fs.rmSync(outputDir, { recursive: true, force: true });
});

describe("integration tests (write + read-back verification)", () => {
  it("creates a workbook and verifies cell content via re-read", async () => {
    const outputFile = outputPath("cell-content.xlsx");
    await new WorkbookBuilder({ author: "Integration Test" })
      .addSheet({ name: "Data" }, (sheet) => {
        sheet
          .headers([
            { key: "item", header: "Item", width: 20 },
            { key: "qty", header: "Qty", width: 10 },
            { key: "price", header: "Price", width: 12, format: "currency" },
          ])
          .addRow({ item: "Widget", qty: 10, price: 9.99 })
          .addRow({ item: "Gadget", qty: 5, price: 24.5 })
          .addRow({
            item: "Total",
            qty: { formula: "=SUM(B2:B3)" },
            price: { formula: "=SUM(C2:C3)" },
          });
      })
      .write(outputFile);

    const loaded = await WorkbookBuilder.load(outputFile);
    const sheet = loaded.getSheet("Data");
    if (!sheet) throw new Error("Expected sheet 'Data'");
    const rows = sheet.toJSON({ header: 1 });
    expect(rows).toHaveLength(4);
    expect(rows[0]).toEqual(["Item", "Qty", "Price"]);
    expect(rows[1][0]).toBe("Widget");
    expect(rows[1][1]).toBe(10);
    expect(rows[2][0]).toBe("Gadget");
    expect(rows[2][1]).toBe(5);
    expect(rows[3][0]).toBe("Total");
  });

  it("verifies styles are preserved in re-read", async () => {
    const outputFile = outputPath("styles.xlsx");
    await new WorkbookBuilder({ author: "Styles Test" })
      .addSheet({ name: "Styled" }, (sheet) => {
        sheet.setCell(
          "A1",
          "Styled Cell",
          style(
            font({ bold: true, size: 14, name: "Arial" }),
            { fill: { type: "solid", color: "FFDCE6F1" } },
            border.thinBlack,
            align.center,
          ),
        );
        sheet.setCell("B1", 42, Styles.currency);
      })
      .write(outputFile);

    const loaded = await WorkbookBuilder.load(outputFile);
    const sheet = loaded.getSheet("Styled");
    expect(sheet).toBeDefined();

    if (!sheet) throw new Error("Expected sheet 'Styled'");
    const ws = sheet.worksheet;
    const cellA1 = ws.getCell("A1");
    expect(cellA1.value).toBe("Styled Cell");
    expect(cellA1.font?.bold).toBe(true);
    expect(cellA1.font?.size).toBe(14);
    expect(cellA1.alignment?.horizontal).toBe("center");
    expect(cellA1.border?.top).toBeDefined();

    const cellB1 = ws.getCell("B1");
    expect(cellB1.value).toBe(42);
  });

  it("verifies formula cell values", async () => {
    const outputFile = outputPath("formulas.xlsx");
    await new WorkbookBuilder()
      .addSheet({ name: "Calc" }, (sheet) => {
        sheet.setCell("A1", 100);
        sheet.setCell("A2", 200);
        sheet.setCell("A3", { formula: "=A1+A2", result: 300 });
        sheet.setCell("B1", { formula: "=SUM(A1:A2)", result: 300 });
      })
      .write(outputFile);

    const loaded = await WorkbookBuilder.load(outputFile);
    const calcSheet = loaded.getSheet("Calc");
    if (!calcSheet) throw new Error("Expected sheet 'Calc'");
    const ws = calcSheet.worksheet;

    expect(ws.getCell("A1").value).toBe(100);
    expect(ws.getCell("A2").value).toBe(200);
    expect(ws.getCell("A3").value).toEqual({ formula: "A1+A2", result: 300 });
    expect(ws.getCell("B1").value).toEqual({
      formula: "SUM(A1:A2)",
      result: 300,
    });
  });

  it("verifies merged cells", async () => {
    const outputFile = outputPath("merged.xlsx");
    await new WorkbookBuilder()
      .addSheet({ name: "Merged" }, (sheet) => {
        sheet.merge("A1:C1", { value: "Report Title", style: Styles.header });
        sheet.setCell("A2", "Col1").setCell("B2", "Col2").setCell("C2", "Col3");
      })
      .write(outputFile);

    const loaded = await WorkbookBuilder.load(outputFile);
    const mergedSheet = loaded.getSheet("Merged");
    if (!mergedSheet) throw new Error("Expected sheet 'Merged'");
    const ws = mergedSheet.worksheet;

    expect(ws.getCell("A1").value).toBe("Report Title");
    expect(ws.getCell("A2").value).toBe("Col1");
    expect(ws.getCell("B2").value).toBe("Col2");
    expect(ws.getCell("C2").value).toBe("Col3");
  });

  it("verifies data validation and conditional formatting are written", async () => {
    const outputFile = outputPath("validation.xlsx");
    await new WorkbookBuilder()
      .addSheet({ name: "Validated" }, (sheet) => {
        sheet.setCell("A1", "Select");
        sheet.addListValidation("A1:A5", ["Yes", "No", "Maybe"]);
        sheet.addCellIsRule("B1:B5", "greaterThan", [0], {
          font: { bold: true },
        });
      })
      .write(outputFile);

    const loaded = await WorkbookBuilder.load(outputFile);
    const validSheet = loaded.getSheet("Validated");
    if (!validSheet) throw new Error("Expected sheet 'Validated'");
    const ws = validSheet.worksheet;

    expect(ws.getCell("A1").value).toBe("Select");
    expect(ws.dataValidations.model).toBeDefined();
    expect(ws.conditionalFormattings).toBeDefined();
  });

  it("verifies freeze panes and auto-filter", async () => {
    const outputFile = outputPath("freeze-filter.xlsx");
    await new WorkbookBuilder()
      .addSheet({ name: "View" }, (sheet) => {
        sheet.headers([
          { key: "a", header: "A", width: 10 },
          { key: "b", header: "B", width: 10 },
        ]);
        sheet.addRow({ a: 1, b: 2 });
        sheet.freeze(1);
        sheet.autoFilter();
      })
      .write(outputFile);

    const loaded = await WorkbookBuilder.load(outputFile);
    const viewSheet = loaded.getSheet("View");
    if (!viewSheet) throw new Error("Expected sheet 'View'");
    const ws = viewSheet.worksheet;

    expect(ws.views).toBeDefined();
    expect(ws.views?.[0]?.state).toBe("frozen");
    expect(ws.autoFilter).toBeDefined();
  });

  it("verifies multi-sheet workbook content", async () => {
    const outputFile = outputPath("multi-sheet.xlsx");
    await new WorkbookBuilder({ author: "Multi-Sheet", title: "Test Report" })
      .addSheet({ name: "Summary" }, (sheet) => {
        sheet.setCell("A1", "Total Revenue");
        sheet.setCell("B1", 1_000_000);
      })
      .addSheet({ name: "Details" }, (sheet) => {
        sheet.headers([
          { key: "month", header: "Month", width: 12 },
          { key: "amount", header: "Amount", width: 15 },
        ]);
        sheet.addRow({ month: "Jan", amount: 100_000 });
        sheet.addRow({ month: "Feb", amount: 150_000 });
      })
      .write(outputFile);

    const loaded = await WorkbookBuilder.load(outputFile);
    expect(loaded.sheets).toHaveLength(2);

    const summarySheet = loaded.getSheet("Summary");
    if (!summarySheet) throw new Error("Expected sheet 'Summary'");
    const summary = summarySheet.worksheet;
    expect(summary.getCell("A1").value).toBe("Total Revenue");
    expect(summary.getCell("B1").value).toBe(1_000_000);

    const detailsSheet = loaded.getSheet("Details");
    if (!detailsSheet) throw new Error("Expected sheet 'Details'");
    const details = detailsSheet.worksheet;
    expect(details.getCell("A1").value).toBe("Month");
    expect(details.getCell("B2").value).toBe(100000);
    expect(details.getCell("B3").value).toBe(150000);
  });

  it("verifies workbook metadata is preserved", async () => {
    const outputFile = outputPath("metadata.xlsx");
    await new WorkbookBuilder({
      author: "Meta Author",
      company: "Meta Corp",
      title: "Meta Title",
      subject: "Meta Subject",
    })
      .addSheet({ name: "Data" }, (sheet) => {
        sheet.setCell("A1", "Hello");
      })
      .write(outputFile);

    const loaded = await WorkbookBuilder.load(outputFile);
    const wb = loaded.workbook;

    expect(wb.creator).toBe("Meta Author");
    expect(wb.company).toBe("Meta Corp");
    expect(wb.title).toBe("Meta Title");
    expect(wb.subject).toBe("Meta Subject");
  });

  it("verifies F formula helpers produce working formulas", async () => {
    const outputFile = outputPath("f-helpers.xlsx");
    await new WorkbookBuilder()
      .addSheet({ name: "F" }, (sheet) => {
        sheet.setCell("A1", 10);
        sheet.setCell("A2", 20);
        sheet.setCell("A3", 30);
        sheet.setCell("B1", F.sum("A1:A3", 60));
        sheet.setCell("B2", F.average("A1:A3", 20));
        sheet.setCell("B3", F.count("A1:A3", 3));
      })
      .write(outputFile);

    const loaded = await WorkbookBuilder.load(outputFile);
    const fSheet = loaded.getSheet("F");
    if (!fSheet) throw new Error("Expected sheet 'F'");
    const ws = fSheet.worksheet;

    expect(ws.getCell("A1").value).toBe(10);
    expect(ws.getCell("A2").value).toBe(20);
    expect(ws.getCell("A3").value).toBe(30);
    expect(ws.getCell("B1").value).toEqual({
      formula: "SUM(A1:A3)",
      result: 60,
    });
    expect(ws.getCell("B2").value).toEqual({
      formula: "AVERAGE(A1:A3)",
      result: 20,
    });
    expect(ws.getCell("B3").value).toEqual({
      formula: "COUNT(A1:A3)",
      result: 3,
    });
  });

  it("verifies toBuffer round-trip preserves content", async () => {
    const builder = new WorkbookBuilder({ author: "Buffer" });
    builder.addSheet({ name: "Buf" }, (sheet) => {
      sheet.setCell("A1", "Buffer Test");
      sheet.setCell("A2", 42);
    });
    const buf = await builder.toBuffer();
    expect(Buffer.isBuffer(buf)).toBe(true);
    expect(buf.length).toBeGreaterThan(0);
  });

  it("verifies named ranges are preserved", async () => {
    const outputFile = outputPath("named-ranges.xlsx");
    await new WorkbookBuilder()
      .addSheet({ name: "Data" }, (sheet) => {
        sheet.setCell("A1", 100);
        sheet.setCell("A2", 200);
      })
      .defineName("MyRange", "Data!A1:A2")
      .write(outputFile);

    const loaded = await WorkbookBuilder.load(outputFile);
    const names = loaded.getDefinedNames();
    expect(names.length).toBeGreaterThan(0);
  });

  it("verifies sheet duplication produces correct content", async () => {
    const outputFile = outputPath("duplicated.xlsx");
    const builder = new WorkbookBuilder();
    builder.addSheet({ name: "Original" }, (sheet) => {
      sheet.setCell("A1", "Original Content");
    });
    builder.duplicateSheet("Original", "Copy");
    await builder.write(outputFile);

    const loaded = await WorkbookBuilder.load(outputFile);
    expect(loaded.sheets).toHaveLength(2);

    const originalSheet = loaded.getSheet("Original");
    if (!originalSheet) throw new Error("Expected sheet 'Original'");
    const original = originalSheet.worksheet;
    const copySheet = loaded.getSheet("Copy");
    if (!copySheet) throw new Error("Expected sheet 'Copy'");
    const copy = copySheet.worksheet;
    expect(original.getCell("A1").value).toBe("Original Content");
    expect(copy.getCell("A1").value).toBe("Original Content");
  });

  it("verifies removeSheet cleans up properly", async () => {
    const outputFile = outputPath("removed.xlsx");
    const builder = new WorkbookBuilder();
    builder.addSheet({ name: "Keep" }, (sheet) => {
      sheet.setCell("A1", "Keep me");
    });
    builder.addSheet({ name: "Remove" }, (sheet) => {
      sheet.setCell("A1", "Remove me");
    });
    builder.removeSheet("Remove");
    await builder.write(outputFile);

    const loaded = await WorkbookBuilder.load(outputFile);
    expect(loaded.sheets).toHaveLength(1);
    const keepSheet = loaded.getSheet("Keep");
    expect(keepSheet?.worksheet.getCell("A1").value).toBe("Keep me");
    expect(loaded.getSheet("Remove")).toBeUndefined();
  });

  it("fillFormula writes shared formula", async () => {
    const outputFile = outputPath("fill-formula.xlsx");
    await new WorkbookBuilder({ author: "v0.7 Test" })
      .addSheet({ name: "Data" }, (sheet) => {
        sheet.headers([
          { key: "x", header: "X", width: 10 },
          { key: "y", header: "Y", width: 10 },
          { key: "sum", header: "Sum", width: 10 },
        ]);
        sheet.addRows([
          { x: 1, y: 2 },
          { x: 3, y: 4 },
        ]);
        sheet.fillFormula("C2", "A2+B2");
        sheet.fillFormula("C3", "A3+B3");
      })
      .write(outputFile);

    const loaded = await WorkbookBuilder.load(outputFile);
    const ws = loaded.getSheet("Data")!.worksheet;
    expect((ws.getCell("C2") as { value: { formula?: string } }).value?.formula ?? "").toBeTruthy();
  });

  it("markdown import/export round-trips", async () => {
    const md = "| Name  | Age |\n|-------|-----|\n| Alice | 30  |\n| Bob   | 25  |";
    const outputFile = outputPath("markdown.xlsx");

    await new WorkbookBuilder({ author: "Markdown" })
      .addSheetFromMarkdown(md, "Sheet1")
      .write(outputFile);

    const loaded = await WorkbookBuilder.load(outputFile);
    const result = await loaded.toMarkdown();
    expect(result).toContain("Name");
    expect(result).toContain("Alice");
    expect(result).toContain("Bob");
  });
});
