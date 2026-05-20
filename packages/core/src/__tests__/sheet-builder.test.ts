import { beforeEach, describe, expect, it, vi } from "vitest";
import { SheetBuilder } from "../sheet-builder.js";
import type { SheetOptions } from "../types.js";

function createMockRow() {
  const callbacks: Array<(cell: unknown, colNumber: number) => void> = [];
  return {
    eachCell: vi.fn((_opts: unknown, cb?: (cell: unknown, colNumber: number) => void) => {
      if (cb) callbacks.push(cb);
    }),
    commit: vi.fn(),
    height: undefined as number | undefined,
    hidden: false,
    outlineLevel: undefined as number | undefined,
    _callbacks: callbacks,
  };
}

function parseAddress(addr: string): { row: number; col: number } {
  const m = addr.match(/^([A-Z]+)(\d+)$/);
  if (!m) return { row: 1, col: 1 };
  const [, letters, digits] = m;
  let col = 0;
  for (const ch of letters) col = col * 26 + ch.charCodeAt(0) - 64;
  return { row: Number.parseInt(digits, 10), col };
}

function createMockCell(address?: string) {
  const addr = address ? parseAddress(address) : { row: 1, col: 1 };
  return {
    fullAddress: { row: addr.row, col: addr.col },
    font: undefined,
    fill: undefined,
    border: undefined,
    alignment: undefined,
    numFmt: undefined,
    protection: undefined,
    value: undefined,
  };
}

// biome-ignore format: mock return shape must stay compact
function createMockWorksheet() {
  // biome-ignore format: cells map
  const cells = new Map<string, object>();
  return {
    addRow: vi.fn(() => createMockRow()),
    getCell: vi.fn((a: string | number, b?: number) => {
      let addr: string | undefined;
      if (typeof a === "number" && b !== undefined) {
        addr = `${String.fromCharCode(64 + b)}${a}`;
      } else {
        addr = String(a);
      }
      if (!cells.has(addr)) cells.set(addr, createMockCell(addr));
      return cells.get(addr);
    }),
    mergeCells: vi.fn(),
    getRow: vi.fn(() => ({ height: undefined })),
    getColumn: vi.fn(() => ({ width: undefined })),
    protect: vi.fn(),
    autoFitColumns: vi.fn(),
    views: undefined,
    properties: { tabColor: undefined, defaultRowHeight: undefined },
    pageSetup: {},
    headerFooter: { oddHeader: "", oddFooter: "", evenHeader: "", evenFooter: "" },
    autoFilter: undefined as string | undefined,
    columns: undefined,
    dataValidations: { add: vi.fn(), find: vi.fn(), remove: vi.fn() },
    addConditionalFormatting: vi.fn(),
    removeConditionalFormatting: vi.fn(),
    conditionalFormattings: [],
  };
}

function makeSheet(opts: SheetOptions) {
  // biome-ignore lint/suspicious/noExplicitAny: mock Worksheet
  const ws = createMockWorksheet() as any;
  const sheet = new SheetBuilder(ws, opts);
  return { ws, sheet };
}

beforeEach(() => {
  vi.restoreAllMocks();
});

describe("SheetBuilder", () => {
  describe("columns / addColumn", () => {
    it("columns() sets column definitions", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.columns([{ key: "a", header: "A" }]);
      sheet.writeHeaders();
      expect(ws.columns).toEqual([{ key: "a", width: 15, hidden: false }]);
    });

    it("addColumn() appends a column", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.addColumn({ key: "a", header: "A" });
      sheet.addColumn({ key: "b", header: "B" });
      sheet.writeHeaders();
      expect(ws.columns).toHaveLength(2);
    });
  });

  describe("writeHeaders", () => {
    it("throws if no columns defined", () => {
      const { sheet } = makeSheet({ name: "Test" });
      expect(() => sheet.writeHeaders()).toThrow("Call columns() before writeHeaders()");
    });

    it("throws if called twice", () => {
      const { sheet } = makeSheet({ name: "Test" });
      sheet.columns([{ key: "a", header: "A" }]).writeHeaders();
      expect(() => sheet.writeHeaders()).toThrow("writeHeaders() already called");
    });

    it("calls addRow on the underlying worksheet", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.columns([{ key: "x", header: "X" }]).writeHeaders();
      expect(ws.addRow).toHaveBeenCalledWith(["X"]);
    });
  });

  describe("addRow", () => {
    it("writes array data", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.addRow(["a", "b"]);
      expect(ws.addRow).toHaveBeenCalledWith(["a", "b"]);
    });

    it("writes object data", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.columns([
        { key: "name", header: "Name" },
        { key: "val", header: "Val" },
      ]);
      sheet.addRow({ name: "Alice", val: 42 });
      expect(ws.addRow).toHaveBeenCalledWith({ name: "Alice", val: 42 });
    });

    it("writes formula values", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.addRow([{ formula: "SUM(A1:A10)" }]);
      expect(ws.addRow).toHaveBeenCalledWith([{ formula: "SUM(A1:A10)" }]);
    });
  });

  describe("addRows", () => {
    it("adds multiple rows", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.addRows([[1], [2], [3]]);
      expect(ws.addRow).toHaveBeenCalledTimes(3);
    });
  });

  describe("setCell", () => {
    it("sets a value at the given address", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.setCell("B3", 42);
      const cell = ws.getCell("B3") as Record<string, unknown>;
      expect(cell.value).toBe(42);
    });

    it("applies style when provided", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.setCell("A1", "hello", { font: { bold: true } });
      const cell = ws.getCell("A1") as Record<string, unknown>;
      expect(cell.font).toBeDefined();
    });

    it("skips undefined value", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.setCell("A1");
      const cell = ws.getCell("A1") as Record<string, unknown>;
      expect(cell.value).toBeUndefined();
    });

    it("writes formula values", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.setCell("C5", { formula: "=B3*2", result: 100 });
      const cell = ws.getCell("C5") as Record<string, unknown>;
      expect(cell.value).toEqual({ formula: "B3*2", result: 100 });
    });
  });

  describe("styleRange", () => {
    it("applies style to cells in range", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.styleRange("A1:B2", { font: { bold: true } });
      expect((ws.getCell("A1") as Record<string, unknown>).font).toBeDefined();
      expect((ws.getCell("B2") as Record<string, unknown>).font).toBeDefined();
    });

    it("returns this for chaining", () => {
      const { sheet } = makeSheet({ name: "Test" });
      const result = sheet.styleRange("A1:A1", {});
      expect(result).toBe(sheet);
    });
  });

  describe("merge", () => {
    it("merges cells and writes value", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.merge({ range: "A1:C1", value: "Title", style: { font: { bold: true } } });
      const cell = ws.getCell("A1") as Record<string, unknown>;
      expect(cell.value).toBe("Title");
      expect((cell.font as Record<string, unknown>)?.bold).toBe(true);
    });
  });

  describe("mergeAll", () => {
    it("merges multiple regions", () => {
      const { sheet } = makeSheet({ name: "Test" });
      sheet.mergeAll([{ range: "A1:C1" }, { range: "A2:C2" }]);
      expect(sheet).toBeDefined();
    });
  });

  describe("rowHeight / colWidth", () => {
    it("sets row height", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.rowHeight(1, 30);
      expect(ws.getRow).toHaveBeenCalledWith(1);
    });

    it("sets column width", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.colWidth("A", 40);
      expect(ws.getColumn).toHaveBeenCalledWith("A");
    });
  });

  describe("autoFitColumns", () => {
    it("delegates to underlying worksheet", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.autoFitColumns();
      expect(ws.autoFitColumns).toHaveBeenCalled();
    });
  });

  describe("freeze", () => {
    it("sets frozen view state", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.freeze(1, 0);
      const views = ws.views as Array<Record<string, unknown>>;
      expect(views[0].state).toBe("frozen");
      expect(views[0].ySplit).toBe(1);
    });

    it("preserves existing view properties", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      ws.views = [{ showGridLines: false, zoomScale: 80 }] as never;
      sheet.freeze(2);
      const views = ws.views as Array<Record<string, unknown>>;
      expect(views[0].showGridLines).toBe(false);
      expect(views[0].zoomScale).toBe(80);
      expect(views[0].state).toBe("frozen");
    });
  });

  describe("autoFilter", () => {
    it("sets range based on columns", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.columns([
        { key: "a", header: "A" },
        { key: "b", header: "B" },
      ]);
      sheet.autoFilter();
      expect(ws.autoFilter).toBe("A1:B1");
    });

    it("accepts explicit range", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.autoFilter("A1:F1");
      expect(ws.autoFilter).toBe("A1:F1");
    });
  });

  describe("_finalize", () => {
    it("protects sheet when protection configured", async () => {
      const { ws, sheet } = makeSheet({ name: "Test", protection: { password: "secret" } });
      await sheet._finalize();
      expect(ws.protect).toHaveBeenCalledWith("secret");
    });

    it("skips when no protection", async () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      await sheet._finalize();
      expect(ws.protect).not.toHaveBeenCalled();
    });
  });

  describe("sheet options", () => {
    it("applies tab color", () => {
      const { ws } = makeSheet({ name: "Test", tabColor: "FFFF0000" });
      expect(ws.properties.tabColor).toEqual({ argb: "FFFF0000" });
    });

    it("applies default row height", () => {
      const { ws } = makeSheet({ name: "Test", defaultRowHeight: 25 });
      expect(ws.properties.defaultRowHeight).toBe(25);
    });

    it("applies zoom", () => {
      const { ws } = makeSheet({ name: "Test", zoom: 120 });
      const views = ws.views as Array<Record<string, unknown>>;
      expect(views[0].zoomScale).toBe(120);
    });

    it("hides gridlines", () => {
      const { ws } = makeSheet({ name: "Test", showGridLines: false });
      const views = ws.views as Array<Record<string, unknown>>;
      expect(views[0].showGridLines).toBe(false);
    });

    it("applies freeze via options", () => {
      const { ws } = makeSheet({ name: "Test", freeze: { row: 2, col: 1 } });
      const views = ws.views as Array<Record<string, unknown>>;
      expect(views[0].state).toBe("frozen");
      expect(views[0].ySplit).toBe(2);
    });

    it("applies header/footer", () => {
      const { ws } = makeSheet({
        name: "Test",
        headerFooter: {
          oddHeader: { left: "Left", right: "Right" },
          oddFooter: { center: "Page &P" },
        },
      });
      expect(ws.headerFooter.oddHeader).toContain("&LLeft");
      expect(ws.headerFooter.oddFooter).toContain("&CPage &P");
    });

    it("applies page orientation", () => {
      const { ws } = makeSheet({
        name: "Test",
        pageSetup: { orientation: "landscape" },
      });
      expect(ws.pageSetup.orientation).toBe("landscape");
    });

    it("applies page margins", () => {
      const { ws } = makeSheet({
        name: "Test",
        pageSetup: { margins: { left: 1.0, right: 1.0 } },
      });
      expect(ws.pageSetup.margins.left).toBe(1.0);
    });
  });

  describe("data validation", () => {
    it("addDataValidation delegates to ws.dataValidations.add", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.addDataValidation("A1", { type: "any" });
      expect(ws.dataValidations.add).toHaveBeenCalledWith("A1", { type: "any" });
    });

    it("addListValidation creates list type validation", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.addListValidation("B2:B10", ["Option A", "Option B"]);
      expect(ws.dataValidations.add).toHaveBeenCalledWith("B2:B10", {
        type: "list",
        formulae: ['"Option A"', '"Option B"'],
      });
    });

    it("addListValidation forwards options", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
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
      sheet.addRangeValidation("D1:D10", "whole", "between", [1, 100]);
      expect(ws.dataValidations.add).toHaveBeenCalledWith("D1:D10", {
        type: "whole",
        operator: "between",
        formulae: [1, 100],
      });
    });

    it("addRangeValidation forwards options", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
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
      sheet.removeConditionalFormatting(0);
      expect(ws.removeConditionalFormatting).toHaveBeenCalledWith(0);
    });

    it("addCellIsRule adds a cellIs rule", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.addCellIsRule("A1:A10", "between", [10, 100]);
      expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
        ref: "A1:A10",
        rules: [{ type: "cellIs", operator: "between", formulae: [10, 100] }],
      });
    });

    it("addCellIsRule with style", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
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
      sheet.addExpressionRule("A1:A10", "A1>100");
      expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
        ref: "A1:A10",
        rules: [{ type: "expression", formulae: ["A1>100"] }],
      });
    });

    it("addDataBar adds a data bar rule", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.addDataBar("B1:B10");
      expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
        ref: "B1:B10",
        rules: [{ type: "dataBar" }],
      });
    });

    it("addColorScale adds a color scale rule", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.addColorScale("C1:C10", [{ type: "min" }, { type: "max" }]);
      expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
        ref: "C1:C10",
        rules: [{ type: "colorScale", cfvo: [{ type: "min" }, { type: "max" }] }],
      });
    });

    it("addIconSet adds an icon set rule", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.addIconSet("D1:D10", "3TrafficLights1");
      expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
        ref: "D1:D10",
        rules: [{ type: "iconSet", iconSet: "3TrafficLights1" }],
      });
    });

    it("addTop10Rule adds a top10 rule", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.addTop10Rule("E1:E10", 5);
      expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
        ref: "E1:E10",
        rules: [{ type: "top10", rank: 5, percent: false }],
      });
    });

    it("addAboveAverageRule adds an aboveAverage rule", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.addAboveAverageRule("F1:F10");
      expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
        ref: "F1:F10",
        rules: [{ type: "aboveAverage" }],
      });
    });

    it("addContainsTextRule adds a containsText rule", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.addContainsTextRule("G1:G10", "urgent");
      expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
        ref: "G1:G10",
        rules: [{ type: "containsText", text: "urgent" }],
      });
    });

    it("addTimePeriodRule adds a timePeriod rule", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.addTimePeriodRule("H1:H10", "thisMonth");
      expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
        ref: "H1:H10",
        rules: [{ type: "timePeriod", timePeriod: "thisMonth" }],
      });
    });

    it("methods are chainable", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet
        .addDataValidation("A1", { type: "any" })
        .addConditionalFormatting({ ref: "B1", rules: [{ type: "dataBar" }] })
        .addCellIsRule("C1", "greaterThan", [0]);
      expect(ws.dataValidations.add).toHaveBeenCalledTimes(1);
      expect(ws.addConditionalFormatting).toHaveBeenCalledTimes(2);
    });
  });
});
