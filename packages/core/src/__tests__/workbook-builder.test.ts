import fs from "node:fs";
import path from "node:path";
import { afterAll, describe, expect, it } from "vitest";
import { F, Styles, WorkbookBuilder } from "../index.js";

const outputDir = path.resolve("output/test");
const outputPath = (name: string) => path.join(outputDir, name);

afterAll(() => {
  fs.rmSync(outputDir, { recursive: true, force: true });
});

describe("WorkbookBuilder", () => {
  it("creates a workbook with a single sheet", async () => {
    const result = await new WorkbookBuilder({ author: "Test" })
      .addSheet({ name: "Sheet1" }, (sheet) => {
        sheet
          .headers([
            { key: "name", header: "Name", width: 20, headerStyle: Styles.header },
            {
              key: "value",
              header: "Value",
              width: 15,
              style: Styles.currency,
              headerStyle: Styles.header,
            },
          ])

          .addRow({ name: "Alpha", value: 100 })
          .addRow({ name: "Beta", value: 200 });
      })
      .write(outputPath("single-sheet.xlsx"));

    expect(result.filePath).toMatch(/single-sheet\.xlsx$/);
    expect(result.sizeBytes).toBeGreaterThan(0);
    expect(fs.existsSync(result.filePath)).toBe(true);
  });

  it("creates a multi-sheet workbook", async () => {
    const result = await new WorkbookBuilder({ author: "Test" })
      .addSheet({ name: "Summary" }, (sheet) => {
        sheet
          .headers([
            { key: "item", header: "Item", width: 15 },
            { key: "qty", header: "Qty", width: 10 },
          ])

          .addRow({ item: "Widget", qty: 42 });
      })
      .addSheet({ name: "Details" }, (sheet) => {
        sheet.setCell("A1", "Hello");
      })
      .write(outputPath("multi-sheet.xlsx"));

    expect(result.filePath).toMatch(/multi-sheet\.xlsx$/);
    expect(result.sizeBytes).toBeGreaterThan(0);
  });

  it("supports formula cells", async () => {
    const result = await new WorkbookBuilder({ author: "Test" })
      .addSheet({ name: "Formulas" }, (sheet) => {
        sheet.setCell("A1", 10);
        sheet.setCell("A2", 20);
        sheet.setCell("A3", { formula: "=SUM(A1:A2)", result: 30 });
      })
      .write(outputPath("formulas.xlsx"));

    expect(result.sizeBytes).toBeGreaterThan(0);
  });

  it("supports merged cells", async () => {
    const result = await new WorkbookBuilder({ author: "Test" })
      .addSheet({ name: "Merged" }, (sheet) => {
        sheet.merge({ range: "A1:C1", value: "Title", style: Styles.header });
      })
      .write(outputPath("merged.xlsx"));

    expect(result.sizeBytes).toBeGreaterThan(0);
  });

  it("supports freeze panes", async () => {
    const result = await new WorkbookBuilder({ author: "Test" })
      .addSheet({ name: "Frozen", freeze: { row: 1 } }, (sheet) => {
        sheet.setCell("A1", "Header").setCell("A2", "Data");
      })
      .write(outputPath("frozen.xlsx"));

    expect(result.sizeBytes).toBeGreaterThan(0);
  });

  it("supports auto-filter", async () => {
    const result = await new WorkbookBuilder()
      .addSheet({ name: "Filtered" }, (sheet) => {
        sheet
          .headers([
            { key: "x", header: "X", width: 10 },
            { key: "y", header: "Y", width: 10 },
          ])

          .addRow({ x: 1, y: 2 })
          .autoFilter();
      })
      .write(outputPath("filter.xlsx"));

    expect(result.sizeBytes).toBeGreaterThan(0);
  });

  it("supports row options (height, style)", async () => {
    const result = await new WorkbookBuilder()
      .addSheet({ name: "Styled" }, (sheet) => {
        sheet
          .headers([{ key: "a", header: "A", width: 10 }])

          .addRow({ a: "Tall" }, { height: 30, style: Styles.totalRow });
      })
      .write(outputPath("row-options.xlsx"));

    expect(result.sizeBytes).toBeGreaterThan(0);
  });

  it("supports styleRange", async () => {
    const result = await new WorkbookBuilder()
      .addSheet({ name: "Range" }, (sheet) => {
        sheet.setCell("A1", "Cell 1").setCell("B1", "Cell 2").styleRange("A1:B1", Styles.boxBorder);
      })
      .write(outputPath("range-style.xlsx"));

    expect(result.sizeBytes).toBeGreaterThan(0);
  });

  it("supports F formula shortcuts", async () => {
    const result = await new WorkbookBuilder()
      .addSheet({ name: "F-Sheet" }, (sheet) => {
        sheet
          .headers([
            { key: "a", header: "A", width: 10 },
            { key: "b", header: "B", width: 10 },
            { key: "total", header: "Total", width: 15 },
          ])

          .addRow({ a: 5, b: 10, total: F.add(F.ref("A", 2), F.ref("B", 2)) });
      })
      .write(outputPath("f-formulas.xlsx"));

    expect(result.sizeBytes).toBeGreaterThan(0);
  });

  it("supports header/footer", async () => {
    const result = await new WorkbookBuilder()
      .addSheet(
        {
          name: "HF",
          headerFooter: {
            oddHeader: { left: "Title", right: "&D" },
            oddFooter: { center: "Page &P" },
          },
        },
        (sheet) => {
          sheet.setCell("A1", "Data");
        },
      )
      .write(outputPath("header-footer.xlsx"));

    expect(result.sizeBytes).toBeGreaterThan(0);
  });

  it("returns buffer via toBuffer()", async () => {
    const buf = await new WorkbookBuilder({ author: "BufferTest" })
      .addSheet({ name: "Buf" }, (sheet) => {
        sheet.setCell("A1", "Hello");
      })
      .toBuffer();

    expect(Buffer.isBuffer(buf)).toBe(true);
    expect(buf.length).toBeGreaterThan(0);
  });

  it("exposes excelts workbook via getter", () => {
    const builder = new WorkbookBuilder();
    expect(builder.workbook).toBeDefined();
    expect(typeof builder.workbook.addWorksheet).toBe("function");
  });

  it("addSheet without callback returns SheetBuilder", () => {
    const builder = new WorkbookBuilder();
    const sheet = builder.addSheet({ name: "Imperative" });
    expect(typeof sheet.setCell).toBe("function");
    expect(typeof sheet.addRow).toBe("function");
  });

  it("tabColor applies to sheet", async () => {
    const result = await new WorkbookBuilder()
      .addSheet({ name: "Colored", tabColor: "FFFF0000" }, (sheet) => {
        sheet.setCell("A1", "Red tab");
      })
      .write(outputPath("tab-color.xlsx"));

    expect(result.sizeBytes).toBeGreaterThan(0);
  });

  it("writes object-array rows", async () => {
    const result = await new WorkbookBuilder()
      .addSheet({ name: "Batch" }, (sheet) => {
        sheet
          .headers([
            { key: "name", header: "Name", width: 15 },
            { key: "age", header: "Age", width: 8 },
          ])

          .addRows([
            { name: "Alice", age: 30 },
            { name: "Bob", age: 25 },
          ]);
      })
      .write(outputPath("batch-rows.xlsx"));

    expect(result.sizeBytes).toBeGreaterThan(0);
  });

  it("writes positional array rows", async () => {
    const result = await new WorkbookBuilder()
      .addSheet({ name: "Array" }, (sheet) => {
        sheet
          .headers([
            { key: "a", header: "A", width: 10 },
            { key: "b", header: "B", width: 10 },
          ])

          .addRows([
            [1, 2],
            [3, 4],
          ]);
      })
      .write(outputPath("array-rows.xlsx"));

    expect(result.sizeBytes).toBeGreaterThan(0);
  });

  it("colWidth and rowHeight work", async () => {
    const result = await new WorkbookBuilder()
      .addSheet({ name: "Sizing" }, (sheet) => {
        sheet.setCell("A1", "Wide column").colWidth("A", 40).rowHeight(1, 50);
      })
      .write(outputPath("sizing.xlsx"));

    expect(result.sizeBytes).toBeGreaterThan(0);
  });

  it("autoFitColumns works", async () => {
    const result = await new WorkbookBuilder()
      .addSheet({ name: "AutoFit" }, (sheet) => {
        sheet.setCell("A1", "Long text here").autoFitColumns();
      })
      .write(outputPath("autofit.xlsx"));

    expect(result.sizeBytes).toBeGreaterThan(0);
  });

  it("pageSetup applies", async () => {
    const result = await new WorkbookBuilder()
      .addSheet(
        {
          name: "Page",
          pageSetup: { orientation: "landscape", paperSize: 9, fitToPage: true },
        },
        (sheet) => {
          sheet.setCell("A1", "Page setup");
        },
      )
      .write(outputPath("page-setup.xlsx"));

    expect(result.sizeBytes).toBeGreaterThan(0);
  });

  it("zoom applies", async () => {
    const result = await new WorkbookBuilder()
      .addSheet({ name: "Zoomed", zoom: 150 }, (sheet) => {
        sheet.setCell("A1", "Zoom 150%");
      })
      .write(outputPath("zoom.xlsx"));

    expect(result.sizeBytes).toBeGreaterThan(0);
  });

  it("hide gridlines", async () => {
    const result = await new WorkbookBuilder()
      .addSheet({ name: "NoGrid", showGridLines: false }, (sheet) => {
        sheet.setCell("A1", "No grid");
      })
      .write(outputPath("no-grid.xlsx"));

    expect(result.sizeBytes).toBeGreaterThan(0);
  });

  it("default row height", async () => {
    const result = await new WorkbookBuilder()
      .addSheet({ name: "DefaultH", defaultRowHeight: 30 }, (sheet) => {
        sheet.setCell("A1", "Taller rows");
      })
      .write(outputPath("default-row-height.xlsx"));

    expect(result.sizeBytes).toBeGreaterThan(0);
  });

  it("returns toCsv as string", async () => {
    const csv = await new WorkbookBuilder()
      .addSheet({ name: "Data" }, (sheet) => {
        sheet
          .headers([
            { key: "name", header: "Name" },
            { key: "age", header: "Age" },
          ])

          .addRows([
            { name: "Alice", age: 30 },
            { name: "Bob", age: 25 },
          ]);
      })
      .toCsv();

    expect(typeof csv).toBe("string");
    expect(csv).toContain("Alice");
    expect(csv).toContain("30");
  });

  it("toCsv writes to file", async () => {
    await new WorkbookBuilder()
      .addSheet({ name: "Data" }, (sheet) => {
        sheet
          .headers([{ key: "x", header: "X" }])

          .addRow({ x: 42 });
      })
      .toCsv(outputPath("test-out.csv"));

    expect(fs.existsSync(outputPath("test-out.csv"))).toBe(true);
  });

  it("load and read back", async () => {
    const written = outputPath("roundtrip.xlsx");
    await new WorkbookBuilder({ author: "Roundtrip" })
      .addSheet({ name: "ReadTest" }, (sheet) => {
        sheet
          .headers([{ key: "val", header: "Value" }])

          .addRow({ val: 99 });
      })
      .write(written);

    const loaded = await WorkbookBuilder.load(written);
    expect(loaded.sheets.length).toBe(1);
  });

  it("template workflow: load → modify → re-save", async () => {
    const original = outputPath("template-original.xlsx");
    const modified = outputPath("template-modified.xlsx");

    await new WorkbookBuilder({ author: "Template" })
      .addSheet({ name: "Sheet1" }, (sheet) => {
        sheet
          .headers([
            { key: "item", header: "Item", width: 20 },
            { key: "qty", header: "Qty", width: 10 },
          ])

          .addRow({ item: "Widget", qty: 10 });
      })
      .write(original);

    const loaded = await WorkbookBuilder.load(original);
    const sheet = loaded.sheets[0];

    expect(sheet).toBeDefined();

    await loaded.write(modified);

    const reloaded = await WorkbookBuilder.load(modified);
    expect(reloaded.sheets.length).toBe(1);
    expect(reloaded.sheets[0]).toBeDefined();
  });

  it("fromFile is an alias for load", async () => {
    const written = outputPath("fromfile.xlsx");
    await new WorkbookBuilder()
      .addSheet({ name: "S" }, (sheet) => sheet.setCell("A1", "hi"))
      .write(written);

    const loaded = await WorkbookBuilder.fromFile(written);
    expect(loaded.sheets.length).toBe(1);
  });

  it("custom author and company metadata", async () => {
    const result = await new WorkbookBuilder({
      author: "Quadro Test",
      company: "Quadro Corp",
    })
      .addSheet({ name: "Meta" }, (sheet) => {
        sheet.setCell("A1", "Meta data");
      })
      .write(outputPath("metadata.xlsx"));

    expect(result.sizeBytes).toBeGreaterThan(0);
  });
});
