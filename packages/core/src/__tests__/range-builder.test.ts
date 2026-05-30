import { describe, expect, it } from "vitest";
import { WorkbookBuilder } from "../workbook-builder.js";

describe("SheetBuilder.range()", () => {
  it("returns a RangeBuilder with the resolved ref", () => {
    const wb = new WorkbookBuilder();
    let rangeRef = "";
    wb.addSheet({ name: "Test" }, (sheet) => {
      const r = sheet.range("A1:C10");
      rangeRef = r.ref;
    });
    expect(rangeRef).toBe("A1:C10");
  });

  it("resolves tuple ranges", () => {
    const wb = new WorkbookBuilder();
    let rangeRef = "";
    wb.addSheet({ name: "Test" }, (sheet) => {
      const r = sheet.range([1, 1, 3, 10]);
      rangeRef = r.ref;
    });
    expect(rangeRef).toBe("A1:C10");
  });

  it("chains style and validation", async () => {
    const wb = new WorkbookBuilder();
    wb.addSheet({ name: "Test" }, (sheet) => {
      sheet
        .headers([
          { key: "name", header: "Name", width: 20 },
          { key: "score", header: "Score", width: 15 },
        ])
        .addRows([
          { name: "Alice", score: 85 },
          { name: "Bob", score: 92 },
        ])
        .range("B2:B3")
        .style({ font: { bold: true } })
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

  it("applies conditional formatting via range builder", async () => {
    const wb = new WorkbookBuilder();
    wb.addSheet({ name: "Test" }, (sheet) => {
      sheet
        .headers([{ key: "val", header: "Value", width: 15 }])
        .addRows([{ val: 10 }, { val: 20 }, { val: 30 }])
        .range("A2:A4")
        .cellIs("greaterThan", [15], {
          fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FF00FF00" } },
        });
    });
    const result = await wb.toBuffer();
    expect(result).toBeInstanceOf(Buffer);
  });

  it("applies list validation via range builder", async () => {
    const wb = new WorkbookBuilder();
    wb.addSheet({ name: "Test" }, (sheet) => {
      sheet
        .headers([{ key: "status", header: "Status", width: 15 }])
        .addRows([{ status: "Active" }])
        .range("A2:A10")
        .listValidation(["Active", "Inactive", "Pending"]);
    });
    const result = await wb.toBuffer();
    expect(result).toBeInstanceOf(Buffer);
  });

  it("applies data bar via range builder", async () => {
    const wb = new WorkbookBuilder();
    wb.addSheet({ name: "Test" }, (sheet) => {
      sheet
        .headers([{ key: "val", header: "Value", width: 15 }])
        .addRows([{ val: 10 }, { val: 20 }, { val: 30 }])
        .range("A2:A4")
        .dataBar({ argb: "FF4472C4" });
    });
    const result = await wb.toBuffer();
    expect(result).toBeInstanceOf(Buffer);
  });

  it("applies color scale via range builder", async () => {
    const wb = new WorkbookBuilder();
    wb.addSheet({ name: "Test" }, (sheet) => {
      sheet
        .headers([{ key: "val", header: "Value", width: 15 }])
        .addRows([{ val: 10 }, { val: 20 }, { val: 30 }])
        .range("A2:A4")
        .colorScale(
          [{ type: "min" }, { type: "max" }],
          [{ argb: "FFF8696B" }, { argb: "FF63BE7B" }],
        );
    });
    const result = await wb.toBuffer();
    expect(result).toBeInstanceOf(Buffer);
  });

  it("merges cells via range builder", async () => {
    const wb = new WorkbookBuilder();
    wb.addSheet({ name: "Test" }, (sheet) => {
      sheet.range("A1:C1").merge({
        value: "Merged Header",
        style: { font: { bold: true }, alignment: { horizontal: "center" } },
      });
    });
    const result = await wb.toBuffer();
    expect(result).toBeInstanceOf(Buffer);
  });

  it("applies expression rule via range builder", async () => {
    const wb = new WorkbookBuilder();
    wb.addSheet({ name: "Test" }, (sheet) => {
      sheet
        .headers([{ key: "val", header: "Value", width: 15 }])
        .addRows([{ val: 10 }, { val: 20 }])
        .range("A2:A3")
        .expression("$A2>15", {
          fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } },
        });
    });
    const result = await wb.toBuffer();
    expect(result).toBeInstanceOf(Buffer);
  });
});
