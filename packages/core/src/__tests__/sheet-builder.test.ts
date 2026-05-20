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
  describe("columns / addColumn", () => {
    it("columns() sets column definitions", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.columns([{ key: "a", header: "A" }]);
      sheet.writeHeaders();
      expect(ws.columns).toHaveLength(1);
      expect(ws.columns[0].key).toBe("a");
      expect(ws.columns[0].width).toBe(15);
      expect(ws.columns[0].hidden).toBe(false);
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
      vi.spyOn(ws, "addRow");
      sheet.columns([{ key: "x", header: "X" }]).writeHeaders();
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

  describe("setCell", () => {
    it("sets a value at the given address", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.setCell("B3", 42);
      const cell = ws.getCell("B3");
      expect(cell.value).toBe(42);
    });

    it("applies style when provided", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.setCell("A1", "hello", { font: { bold: true } });
      const cell = ws.getCell("A1");
      expect(cell.font).toBeDefined();
    });

    it("skips undefined value", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.setCell("A1");
      const cell = ws.getCell("A1");
      expect(cell.value).toBeNull();
    });

    it("writes formula values", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.setCell("C5", { formula: "=B3*2", result: 100 });
      const cell = ws.getCell("C5");
      expect(cell.value).toEqual({ formula: "B3*2", result: 100 });
    });
  });

  describe("styleRange", () => {
    it("applies style to cells in range", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.styleRange("A1:B2", { font: { bold: true } });
      expect(ws.getCell("A1").font).toBeDefined();
      expect(ws.getCell("B2").font).toBeDefined();
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
      sheet.merge({
        range: "A1:C1",
        value: "Title",
        style: { font: { bold: true } },
      });
      const cell = ws.getCell("A1");
      expect(cell.value).toBe("Title");
      expect(cell.font?.bold).toBe(true);
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
      vi.spyOn(ws, "getRow");
      sheet.rowHeight(1, 30);
      expect(ws.getRow).toHaveBeenCalledWith(1);
    });

    it("sets column width", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "getColumn");
      sheet.colWidth("A", 40);
      expect(ws.getColumn).toHaveBeenCalledWith("A");
    });
  });

  describe("autoFitColumns", () => {
    it("delegates to underlying worksheet", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "autoFitColumns");
      sheet.autoFitColumns();
      expect(ws.autoFitColumns).toHaveBeenCalled();
    });
  });

  describe("freeze", () => {
    it("sets frozen view state", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.freeze(1, 0);
      const view = ws.views[0];
      if (view.state !== "frozen") throw new Error("expected frozen view");
      expect(view.ySplit).toBe(1);
    });

    it("preserves existing view properties", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      ws.views = [{ showGridLines: false, zoomScale: 80 }];
      sheet.freeze(2);
      const views = ws.views;
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
      const { ws, sheet } = makeSheet({
        name: "Test",
        protection: { password: "secret" },
      });
      vi.spyOn(ws, "protect");
      await sheet._finalize();
      expect(ws.protect).toHaveBeenCalledWith("secret");
    });

    it("skips when no protection", async () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "protect");
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
      const views = ws.views;
      expect(views[0].zoomScale).toBe(120);
    });

    it("hides gridlines", () => {
      const { ws } = makeSheet({ name: "Test", showGridLines: false });
      const views = ws.views;
      expect(views[0].showGridLines).toBe(false);
    });

    it("applies freeze via options", () => {
      const { ws } = makeSheet({
        name: "Test",
        freeze: { row: 2, col: 1 },
      });
      const view = ws.views[0];
      if (view.state !== "frozen") throw new Error("expected frozen view");
      expect(view.ySplit).toBe(2);
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
});
