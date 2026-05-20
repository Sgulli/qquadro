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
});
