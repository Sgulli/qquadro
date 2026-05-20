import { Workbook } from "@cj-tech-master/excelts";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { SheetBuilder } from "../sheet-builder.js";
import type { SheetOptions } from "../types.js";

function makeSheet(opts: SheetOptions) {
  const wb = new Workbook();
  const ws = wb.addWorksheet(opts.name);
  const sheet = new SheetBuilder(ws, opts);
  return { ws, sheet };
}

beforeEach(() => {
  vi.restoreAllMocks();
});

describe("SheetBuilder", () => {
  describe("headers", () => {
    it("sets column definitions and writes header row", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addRow");
      sheet.headers([{ key: "a", header: "A" }]);
      expect(ws.columns).toHaveLength(1);
      expect(ws.columns[0].key).toBe("a");
      expect(ws.columns[0].width).toBe(15);
      expect(ws.columns[0].hidden).toBe(false);
      expect(ws.addRow).toHaveBeenCalledWith(["A"]);
    });

    it("calls addRow on the underlying worksheet", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addRow");
      sheet.headers([{ key: "x", header: "X" }]);
      expect(ws.addRow).toHaveBeenCalledWith(["X"]);
    });
  });

  describe("addRow", () => {
    it("writes array data", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addRow");
      sheet.addRow(["a", "b"]);
      expect(ws.addRow).toHaveBeenCalledWith(["a", "b"]);
    });

    it("writes object data", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addRow");
      sheet.columns([
        { key: "name", header: "Name" },
        { key: "val", header: "Val" },
      ]);
      sheet.addRow({ name: "Alice", val: 42 });
      expect(ws.addRow).toHaveBeenCalledWith({ name: "Alice", val: 42 });
    });

    it("writes formula values", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addRow");
      sheet.addRow([{ formula: "SUM(A1:A10)" }]);
      expect(ws.addRow).toHaveBeenCalledWith([{ formula: "SUM(A1:A10)" }]);
    });
  });

  describe("addRows", () => {
    it("adds multiple rows", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addRow");
      sheet.addRows([[1], [2], [3]]);
      expect(ws.addRow).toHaveBeenCalledTimes(3);
    });
  });

  describe("setCell / setCellRC", () => {
    it("setCellRC delegates with cellRef", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.setCellRC(4, 5, 42);
      expect(ws.getCell("D5").value).toBe(42);
    });

    it("sets a value at the given address", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.setCell("B3", 42);
      expect(ws.getCell("B3").value).toBe(42);
    });

    it("applies style when provided", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.setCell("A1", "hello", { font: { bold: true } });
      expect(ws.getCell("A1").font?.bold).toBe(true);
    });

    it("skips undefined value", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.setCell("A1");
      expect(ws.getCell("A1").value).toBeFalsy();
    });

    it("writes formula values", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.setCell("C5", { formula: "SUM(A1:A10)" });
      expect(ws.getCell("C5").value).toEqual({ formula: "SUM(A1:A10)" });
    });
  });

  describe("styleRange / styleRangeRC", () => {
    it("applies style to a range", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.setCell("A1", "a");
      sheet.setCell("B2", "b");
      sheet.styleRange("A1:B2", { font: { bold: true } });
      expect(ws.getCell("A1").font?.bold).toBe(true);
      expect(ws.getCell("B2").font?.bold).toBe(true);
    });

    it("handles single cell range", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.setCell("A1", "a");
      const result = sheet.styleRange("A1:A1", { font: { bold: true } });
      expect(ws.getCell("A1").font?.bold).toBe(true);
      expect(result).toBe(sheet);
    });
  });

  describe("merge / mergeRC", () => {
    it("merges cells and writes value", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "mergeCells");
      sheet.merge({ range: "A1:C1", value: "Title", style: { font: { bold: true } } });
      expect(ws.mergeCells).toHaveBeenCalledWith("A1:C1");
      expect(ws.getCell("A1").value).toBe("Title");
      expect(ws.getCell("A1").font?.bold).toBe(true);
    });

    it("mergeRC delegates with rangeRef", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "mergeCells");
      sheet.mergeRC(1, 1, 3, 1, { value: "Title" });
      expect(ws.mergeCells).toHaveBeenCalledWith("A1:C1");
    });
  });

  describe("rowHeight / colWidth", () => {
    it("sets row height", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.rowHeight(1, 30);
      expect(ws.getRow(1).height).toBe(30);
    });

    it("sets column width", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.colWidth("A", 40);
      expect(ws.getColumn("A").width).toBe(40);
    });
  });

  describe("freeze", () => {
    it("sets freeze pane", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.freeze(3);
      const view = ws.views?.[0] ?? null;
      expect(view?.state).toBe("frozen");
      if (view?.state === "frozen") {
        expect(view.ySplit).toBe(3);
      }
    });
  });

  describe("autoFilter", () => {
    it("sets auto filter range", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.autoFilter("A1:C1");
      expect(ws.autoFilter).toBe("A1:C1");
    });
  });

  describe("autoFitColumns", () => {
    it("auto-fits columns", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "autoFitColumns");
      sheet.autoFitColumns();
      expect(ws.autoFitColumns).toHaveBeenCalledWith(undefined, undefined);
    });
  });

  describe("rowCount / columnRange", () => {
    it("tracks row count after addRow", () => {
      const { sheet } = makeSheet({ name: "Test" });
      sheet.addRow([1]);
      expect(sheet.rowCount).toBe(1);
      sheet.addRow([2]);
      expect(sheet.rowCount).toBe(2);
    });

    it("tracks row count after addRows", () => {
      const { sheet } = makeSheet({ name: "Test" });
      sheet.addRows([[1], [2], [3]]);
      expect(sheet.rowCount).toBe(3);
    });

    it("columnRange resolves column key to range string", () => {
      const { sheet } = makeSheet({ name: "Test" });
      sheet
        .headers([
          { key: "name", header: "Name" },
          { key: "score", header: "Score" },
        ])
        .addRows([
          { name: "Alice", score: 85 },
          { name: "Bob", score: 42 },
        ]);
      expect(sheet.columnRange("name")).toBe("A2:A3");
      expect(sheet.columnRange("score")).toBe("B2:B3");
    });

    it("columnRange accepts custom startRow", () => {
      const { sheet } = makeSheet({ name: "Test" });
      sheet.headers([{ key: "x", header: "X" }]);
      sheet.addRow(["a"]);
      expect(sheet.columnRange("x", 1)).toBe("A1:A1");
    });

    it("columnRange throws for unknown key", () => {
      const { sheet } = makeSheet({ name: "Test" });
      expect(() => sheet.columnRange("missing")).toThrow("No column with key");
    });
  });

  describe("auto-inferred columns", () => {
    it("infers columns from first object row", () => {
      const { sheet } = makeSheet({ name: "Test" });
      sheet.addRow({ name: "Alice", score: 85 });
      expect(sheet.columnRange("name")).toBe("A1:A1");
      expect(sheet.columnRange("score")).toBe("B1:B1");
    });

    it("columnRange works after addRows with inferred columns", () => {
      const { sheet } = makeSheet({ name: "Test" });
      sheet.addRows([
        { product: "Widget", quantity: 10 },
        { product: "Gadget", quantity: 20 },
        { product: "Doohickey", quantity: 30 },
      ]);
      expect(sheet.rowCount).toBe(3);
      expect(sheet.columnRange("product")).toBe("A1:A3");
      expect(sheet.columnRange("quantity")).toBe("B1:B3");
    });

    it("explicit columns take precedence over inference", () => {
      const { sheet } = makeSheet({ name: "Test" });
      sheet.headers([{ key: "custom", header: "Custom" }]);
      sheet.addRow({ name: "Alice", score: 85 });
      expect(sheet.columnRange("custom")).toBe("A2:A2");
      expect(() => sheet.columnRange("name")).toThrow("No column with key");
    });
  });

  describe("data validation", () => {
    it("addDataValidation delegates to ws.dataValidations.add", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws.dataValidations, "add");
      sheet.addDataValidation("A1", { type: "any" });
      expect(ws.dataValidations.add).toHaveBeenCalledWith("A1", { type: "any" });
    });

    it("addListValidation creates list type validation", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws.dataValidations, "add");
      sheet.addListValidation("B2:B10", ["Option A", "Option B"]);
      expect(ws.dataValidations.add).toHaveBeenCalledWith("B2:B10", {
        type: "list",
        formulae: ['"Option A"', '"Option B"'],
      });
    });

    it("addListValidation forwards options", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws.dataValidations, "add");
      sheet.addListValidation("C1", ["X"], { allowBlank: true, prompt: "Pick one" });
      expect(ws.dataValidations.add).toHaveBeenCalledWith("C1", {
        type: "list",
        formulae: ['"X"'],
        allowBlank: true,
        prompt: "Pick one",
      });
    });

    it("addRangeValidation with between operator", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws.dataValidations, "add");
      sheet.addRangeValidation("D1:D10", "whole", "between", [1, 100]);
      expect(ws.dataValidations.add).toHaveBeenCalledWith("D1:D10", {
        type: "whole",
        operator: "between",
        formulae: [1, 100],
      });
    });

    it("addRangeValidation forwards options", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws.dataValidations, "add");
      sheet.addRangeValidation("E1", "decimal", "lessThan", [50], { error: "Too high" });
      expect(ws.dataValidations.add).toHaveBeenCalledWith("E1", {
        type: "decimal",
        operator: "lessThan",
        formulae: [50],
        error: "Too high",
      });
    });
  });

  describe("conditional formatting", () => {
    it("addConditionalFormatting delegates to ws.addConditionalFormatting", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addConditionalFormatting");
      sheet.addConditionalFormatting({
        ref: "A1:A10",
        rules: [{ type: "cellIs", operator: "greaterThan", formulae: [100] }],
      });
      expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
        ref: "A1:A10",
        rules: [{ type: "cellIs", operator: "greaterThan", formulae: [100] }],
      });
    });

    it("removeConditionalFormatting delegates to ws", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "removeConditionalFormatting");
      sheet.removeConditionalFormatting(0);
      expect(ws.removeConditionalFormatting).toHaveBeenCalledWith(0);
    });

    it("addCellIsRule adds a cellIs rule", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addConditionalFormatting");
      sheet.addCellIsRule("A1:A10", "between", [10, 100]);
      expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
        ref: "A1:A10",
        rules: [{ type: "cellIs", operator: "between", formulae: [10, 100] }],
      });
    });

    it("addCellIsRule with style", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addConditionalFormatting");
      sheet.addCellIsRule("A1", "equal", ["Yes"], { font: { bold: true } });
      expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
        ref: "A1",
        rules: [
          { type: "cellIs", operator: "equal", formulae: ["Yes"], style: { font: { bold: true } } },
        ],
      });
    });

    it("addExpressionRule adds an expression rule", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addConditionalFormatting");
      sheet.addExpressionRule("A1:A10", "A1>100");
      expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
        ref: "A1:A10",
        rules: [{ type: "expression", formulae: ["A1>100"] }],
      });
    });

    it("addDataBar adds a data bar rule", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addConditionalFormatting");
      sheet.addDataBar("B1:B10");
      expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
        ref: "B1:B10",
        rules: [{ type: "dataBar" }],
      });
    });

    it("addColorScale adds a color scale rule", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addConditionalFormatting");
      sheet.addColorScale("C1:C10", [{ type: "min" }, { type: "max" }]);
      expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
        ref: "C1:C10",
        rules: [{ type: "colorScale", cfvo: [{ type: "min" }, { type: "max" }] }],
      });
    });

    it("addIconSet adds an icon set rule", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addConditionalFormatting");
      sheet.addIconSet("D1:D10", "3TrafficLights1");
      expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
        ref: "D1:D10",
        rules: [{ type: "iconSet", iconSet: "3TrafficLights1" }],
      });
    });

    it("addTop10Rule adds a top10 rule", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addConditionalFormatting");
      sheet.addTop10Rule("E1:E10", 5);
      expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
        ref: "E1:E10",
        rules: [{ type: "top10", rank: 5, percent: false }],
      });
    });

    it("addAboveAverageRule adds an aboveAverage rule", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addConditionalFormatting");
      sheet.addAboveAverageRule("F1:F10");
      expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
        ref: "F1:F10",
        rules: [{ type: "aboveAverage" }],
      });
    });

    it("addContainsTextRule adds a containsText rule", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addConditionalFormatting");
      sheet.addContainsTextRule("G1:G10", "urgent");
      expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
        ref: "G1:G10",
        rules: [{ type: "containsText", text: "urgent" }],
      });
    });

    it("addTimePeriodRule adds a timePeriod rule", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addConditionalFormatting");
      sheet.addTimePeriodRule("H1:H10", "thisMonth");
      expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
        ref: "H1:H10",
        rules: [{ type: "timePeriod", timePeriod: "thisMonth" }],
      });
    });

    it("methods are chainable", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws.dataValidations, "add");
      vi.spyOn(ws, "addConditionalFormatting");
      sheet
        .addDataValidation("A1", { type: "any" })
        .addConditionalFormatting({ ref: "B1", rules: [{ type: "dataBar" }] })
        .addCellIsRule("C1", "greaterThan", [0]);
      expect(ws.dataValidations.add).toHaveBeenCalledTimes(1);
      expect(ws.addConditionalFormatting).toHaveBeenCalledTimes(2);
    });
  });

  describe("data validation RC", () => {
    it("addDataValidationRC delegates with cellRef", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws.dataValidations, "add");
      sheet.addDataValidationRC(3, 5, { type: "any" });
      expect(ws.dataValidations.add).toHaveBeenCalledWith("C5", { type: "any" });
    });

    it("addListValidationRC delegates with colRange", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws.dataValidations, "add");
      sheet.addListValidationRC(2, 3, 10, ["X", "Y"]);
      expect(ws.dataValidations.add).toHaveBeenCalledWith("B3:B10", {
        type: "list",
        formulae: ['"X"', '"Y"'],
      });
    });

    it("addRangeValidationRC delegates with colRange", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws.dataValidations, "add");
      sheet.addRangeValidationRC(4, 2, 100, "whole", "between", [1, 100]);
      expect(ws.dataValidations.add).toHaveBeenCalledWith("D2:D100", {
        type: "whole",
        operator: "between",
        formulae: [1, 100],
      });
    });
  });

  describe("conditional formatting RC", () => {
    it("addCellIsRuleRC delegates with rangeRef", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addConditionalFormatting");
      sheet.addCellIsRuleRC(2, 3, 4, 10, "greaterThan", [100]);
      expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
        ref: "B3:D10",
        rules: [{ type: "cellIs", operator: "greaterThan", formulae: [100] }],
      });
    });

    it("addExpressionRuleRC delegates with rangeRef", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addConditionalFormatting");
      sheet.addExpressionRuleRC(1, 1, 5, 5, "A1>0");
      expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
        ref: "A1:E5",
        rules: [{ type: "expression", formulae: ["A1>0"] }],
      });
    });

    it("addDataBarRC delegates with rangeRef", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addConditionalFormatting");
      sheet.addDataBarRC(2, 2, 2, 20);
      expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
        ref: "B2:B20",
        rules: [{ type: "dataBar" }],
      });
    });

    it("addColorScaleRC delegates with rangeRef", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addConditionalFormatting");
      sheet.addColorScaleRC(3, 1, 3, 10, [{ type: "min" }, { type: "max" }]);
      expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
        ref: "C1:C10",
        rules: [{ type: "colorScale", cfvo: [{ type: "min" }, { type: "max" }] }],
      });
    });

    it("addIconSetRC delegates with rangeRef", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addConditionalFormatting");
      sheet.addIconSetRC(4, 1, 4, 10, "3TrafficLights1");
      expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
        ref: "D1:D10",
        rules: [{ type: "iconSet", iconSet: "3TrafficLights1" }],
      });
    });

    it("addTop10RuleRC delegates with rangeRef", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addConditionalFormatting");
      sheet.addTop10RuleRC(5, 1, 5, 10, 5);
      expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
        ref: "E1:E10",
        rules: [{ type: "top10", rank: 5, percent: false }],
      });
    });

    it("addAboveAverageRuleRC delegates with rangeRef", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addConditionalFormatting");
      sheet.addAboveAverageRuleRC(6, 1, 6, 10);
      expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
        ref: "F1:F10",
        rules: [{ type: "aboveAverage" }],
      });
    });

    it("addContainsTextRuleRC delegates with rangeRef", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addConditionalFormatting");
      sheet.addContainsTextRuleRC(7, 1, 7, 10, "urgent");
      expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
        ref: "G1:G10",
        rules: [{ type: "containsText", text: "urgent" }],
      });
    });

    it("addTimePeriodRuleRC delegates with rangeRef", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addConditionalFormatting");
      sheet.addTimePeriodRuleRC(8, 1, 8, 10, "thisMonth");
      expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
        ref: "H1:H10",
        rules: [{ type: "timePeriod", timePeriod: "thisMonth" }],
      });
    });
  });

  describe("tables", () => {
    it("addTable delegates to ws.addTable", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addTable");
      sheet.addTable("MyTable", "A1:C10", [{ name: "Name" }, { name: "Score" }]);
      expect(ws.addTable).toHaveBeenCalled();
      const args = vi.mocked(ws.addTable).mock.calls[0][0];
      expect(args).toMatchObject({
        name: "MyTable",
        ref: "A1:C10",
        headerRow: true,
      });
    });

    it("addTable with rows and options", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addTable");
      sheet.addTable("Sales", "A1:D5", [{ name: "Product" }, { name: "Revenue" }], {
        rows: [
          ["Widget", 100],
          ["Gadget", 200],
        ],
        totalsRow: true,
        style: { theme: "TableStyleMedium9", showRowStripes: true },
      });
      expect(ws.addTable).toHaveBeenCalled();
      const args = vi.mocked(ws.addTable).mock.calls[0][0];
      expect(args).toMatchObject({
        name: "Sales",
        ref: "A1:D5",
        totalsRow: true,
        style: { theme: "TableStyleMedium9", showRowStripes: true },
      });
    });

    it("addTableRC delegates with rangeRef", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addTable");
      sheet.addTableRC("Data", 1, 1, 3, 10, [{ name: "A" }, { name: "B" }, { name: "C" }]);
      expect(ws.addTable).toHaveBeenCalled();
      const args = vi.mocked(ws.addTable).mock.calls[0][0];
      expect(args).toMatchObject({ name: "Data", ref: "A1:C10" });
    });

    it("methods are chainable", () => {
      const { sheet } = makeSheet({ name: "Test" });
      const result = sheet
        .addTable("T1", "A1:B2", [{ name: "X" }])
        .addTable("T2", "D1:E2", [{ name: "Y" }]);
      expect(result).toBe(sheet);
    });
  });
});
