import type {
  AddAOAOptions,
  AddJSONOptions,
  Cell as ExcelCell,
  CellValue as ExcelCellValue,
  Row as ExcelRow,
  Worksheet as ExcelWorksheet,
  SheetToJSONOptions,
  TableColumnProperties,
  TableStyleProperties,
  WorksheetView,
} from "@cj-tech-master/excelts";
import { installChartSupport } from "@cj-tech-master/excelts/chart";

installChartSupport();

import { type ColumnMap, type ColumnSchemaMap, createColumnMap } from "./column-map.js";
import { colLetter, resolveAddr, resolveRange } from "./coords.js";
import { isFormula, toExcelValue, toFormulaValue } from "./formulas.js";
import { applyChartMixin } from "./mixins/charts.js";
import { applyConditionalFormattingMixin } from "./mixins/conditional-formatting.js";
import { applyDataValidationMixin } from "./mixins/data-validation.js";
import { applyAutoFilter, applyFreeze, applyMerge, applyStyleRange } from "./mixins/format.js";
import { applyMediaMixin } from "./mixins/media.js";
import { RangeBuilder } from "./range-builder.js";
import { applyStyle, formatHeaderFooterSection } from "./style-presets.js";
import type {
  Addr,
  CellRange,
  CellStyle,
  CellValue,
  ColumnDef,
  IgnoredErrorDef,
  MergeRange,
  RowData,
  RowOptions,
  SheetBuilderExtension,
  SheetOptions,
} from "./types.js";

export const _sheetFinalizers = new WeakMap<SheetBuilder, () => Promise<void>>();

/** Fluent builder for constructing a single worksheet with column definitions, data rows, and formatting. */
export class SheetBuilder {
  private readonly _ws: ExcelWorksheet;

  private _columns: ColumnDef[] = [];
  private _headerWritten = false;
  private _protectionConfig?: { password?: string };
  private _rowCount = 0;
  private _columnsInferred = false;

  /** Number of data rows written (excluding headers). */
  get rowCount(): number {
    return this._rowCount;
  }

  /** Sheet name. */
  get name(): string {
    return this._opts.name;
  }

  /** Underlying excelts Worksheet instance. */
  get worksheet(): ExcelWorksheet {
    return this._ws;
  }

  /** 1-based index of a column by key. Throws if the key is not found. */
  columnIndex(key: string): number {
    const col = this._columns.find((c) => c.key === key);
    if (!col) throw new Error(`[SheetBuilder] No column with key "${key}".`);
    return this._columns.indexOf(col) + 1;
  }

  /** A1-style range string covering a column across all data rows (e.g. "B2:B10"). */
  columnRange(key: string, startRow?: number): string {
    if (startRow === undefined) startRow = this._headerWritten ? 2 : 1;
    const col = this._columns.find((c) => c.key === key);
    if (!col) throw new Error(`[SheetBuilder] No column with key "${key}".`);
    const idx = this._columns.indexOf(col);
    const letter = colLetter(idx + 1);
    return `${letter}${startRow}:${letter}${startRow + this._rowCount - 1}`;
  }

  constructor(
    ws: ExcelWorksheet,
    private readonly _opts: SheetOptions,
  ) {
    this._ws = ws;
    this._applySheetOptions();
    _sheetFinalizers.set(this, () => this._doFinalize());
  }

  /** Set column definitions (replaces any existing columns). */
  columns(defs: ColumnDef[]): this {
    this._columns = defs;
    return this;
  }

  /** Define columns explicitly and return a typed ColumnMap for safe references. */
  defineColumns<T extends ColumnSchemaMap>(schema: T): ColumnMap<T> {
    const map = createColumnMap(schema, () => ({
      rowCount: this._rowCount,
      headerWritten: this._headerWritten,
    }));
    this._columns = Object.values(map).map((ref) =>
      (ref as import("./column-map.js").ColumnRef).toColumnDef(),
    );
    return map;
  }

  /** Append a single column definition. */
  addColumn(def: ColumnDef): this {
    this._columns.push(def);
    return this;
  }

  /** Set column definitions and write the header row in one call. */
  headers(defs: ColumnDef[], globalStyle?: CellStyle, height?: number): this {
    this._columns = defs;
    this._writeHeaders(globalStyle, height);
    return this;
  }

  /** Append one data row (object or array) to the sheet. */
  addRow(data: RowData, options?: RowOptions): this {
    this._flushRow(data, options);
    return this;
  }

  /** Append multiple data rows with optional per-row overrides or a shared options bag. */
  addRows(rows: RowData[], sharedOptions?: RowOptions): this;
  addRows(rows: Array<{ data: RowData; options?: RowOptions }>): this;
  addRows(
    rows: RowData[] | Array<{ data: RowData; options?: RowOptions }>,
    sharedOptions?: RowOptions,
  ): this {
    if (rows.length === 0) return this;
    const errors: Array<{ index: number; error: Error }> = [];
    if ("data" in rows[0]) {
      for (let i = 0; i < rows.length; i++) {
        const item = rows[i] as { data: RowData; options?: RowOptions };
        try {
          this.addRow(item.data, item.options);
        } catch (e) {
          errors.push({
            index: i,
            error: e instanceof Error ? e : new Error(String(e)),
          });
        }
      }
    } else {
      for (let i = 0; i < rows.length; i++) {
        try {
          this.addRow(rows[i] as RowData, sharedOptions);
        } catch (e) {
          errors.push({
            index: i,
            error: e instanceof Error ? e : new Error(String(e)),
          });
        }
      }
    }
    if (errors.length > 0) {
      throw new Error(
        `[SheetBuilder] Failed to add ${errors.length} of ${rows.length} rows:\n${errors.map((e) => `  row[${e.index}]: ${e.error.message}`).join("\n")}`,
      );
    }
    return this;
  }

  /** Set a cell's value and optional style by address (e.g. "A1" or { col: 1, row: 1 }). */
  setCell(addr: Addr, value?: CellValue, style?: CellStyle): this {
    const address = resolveAddr(addr);
    const cell = this._ws.getCell(address);
    this._writeValue(cell, value);
    if (style) applyStyle(cell, style);
    return this;
  }

  /** Apply a style to a range of cells. */
  styleRange(range: CellRange, style: CellStyle): this {
    applyStyleRange(this._ws, range, style);
    return this;
  }

  /** Merge cells over a range, optionally setting a value, style, and row height. */
  merge(
    range: CellRange,
    options?: { value?: CellValue; style?: CellStyle; height?: number },
  ): this;
  merge(region: MergeRange): this;
  merge(
    rangeOrRegion: CellRange | MergeRange,
    options?: { value?: CellValue; style?: CellStyle; height?: number },
  ): this {
    let range: string;
    let region: { value?: CellValue; style?: CellStyle; height?: number };
    if (typeof rangeOrRegion === "object" && "range" in rangeOrRegion) {
      range = rangeOrRegion.range;
      region = rangeOrRegion;
    } else {
      range = resolveRange(rangeOrRegion as CellRange);
      region = options ?? {};
    }
    applyMerge(this._ws, range, region.value, region.style, region.height);
    return this;
  }

  /** Merge multiple regions in batch. */
  mergeAll(regions: MergeRange[]): this {
    for (const r of regions) this.merge(r);
    return this;
  }

  /** Set the height of a specific row by number. */
  rowHeight(rowNumber: number, height: number): this {
    this._ws.getRow(rowNumber).height = height;
    return this;
  }

  /** Set the width of a column (by letter or index). */
  colWidth(col: string | number, width: number): this {
    this._ws.getColumn(col).width = width;
    return this;
  }

  /** Auto-fit column widths within an optional range. */
  autoFitColumns(startCol?: number | string, endCol?: number | string): this {
    this._ws.autoFitColumns(startCol, endCol);
    return this;
  }

  /** Freeze rows and/or columns at a given split point. */
  freeze(opts: { row?: number; col?: number }): this;
  freeze(row: number, col?: number): this;
  freeze(rowOrOpts: number | { row?: number; col?: number }, maybeCol = 0): this {
    const col: number = typeof rowOrOpts === "object" ? (rowOrOpts.col ?? 0) : maybeCol;
    const row: number = typeof rowOrOpts === "object" ? (rowOrOpts.row ?? 0) : rowOrOpts;
    applyFreeze(this._ws, row, col);
    return this;
  }

  /** Enable auto-filter on the header row (or a custom range). */
  autoFilter(range?: string): this {
    const r = range ?? (this._columns.length ? `A1:${colLetter(this._columns.length)}1` : "A1");
    applyAutoFilter(this._ws, r);
    return this;
  }

  /** Create a RangeBuilder for fluent range operations (style, validation, conditional formatting). */
  range(spec: CellRange): RangeBuilder {
    return new RangeBuilder(this as unknown as import("./range-builder.js").SheetLike, spec);
  }

  /** Add a structured table to the sheet. */
  addTable(
    name: string,
    ref: CellRange,
    columns: TableColumnProperties[],
    options?: {
      rows?: ExcelCellValue[][];
      headerRow?: boolean;
      totalsRow?: boolean;
      style?: TableStyleProperties;
    },
  ): this {
    this._ws.addTable({
      name,
      ref: resolveRange(ref),
      columns,
      rows: options?.rows ?? [],
      headerRow: options?.headerRow ?? true,
      totalsRow: options?.totalsRow,
      style: options?.style,
    });
    return this;
  }

  /** Insert a row at a given position with cell values and optional row-level options. */
  insertRow(pos: number, data: CellValue[], options?: RowOptions): this {
    const excelRow = this._ws.insertRow(pos, data.map(toExcelValue));
    if (options) {
      if (options.height !== undefined) excelRow.height = options.height;
      if (options.hidden) excelRow.hidden = true;
      if (options.outlineLevel !== undefined) excelRow.outlineLevel = options.outlineLevel;
      if (options.style) {
        excelRow.eachCell({ includeEmpty: true }, (cell) => applyStyle(cell, options.style));
      }
    }
    excelRow.commit();
    this._rowCount++;
    return this;
  }

  /** Duplicate a row one or more times, optionally inserting rather than appending. */
  duplicateRow(rowNum: number, count: number, insert?: boolean): this {
    this._ws.duplicateRow(rowNum, count, insert);
    this._rowCount += count;
    return this;
  }

  /** Remove a range of contiguous rows by starting index and count. */
  removeRow(start: number, count: number): this {
    this._ws.spliceRows(start, count);
    this._rowCount = Math.max(0, this._rowCount - count);
    return this;
  }

  /** Insert columns at a given position. */
  insertColumn(start: number, count: number, ...inserts: ExcelCellValue[][]): this {
    this._ws.spliceColumns(start, count, ...inserts);
    return this;
  }

  /** Remove a range of columns by starting index and count. */
  removeColumn(start: number, count: number): this {
    this._ws.spliceColumns(start, count);
    return this;
  }

  /** Insert a page break at a specific row number. */
  addPageBreak(rowNum: number): this {
    this._ws.getRow(rowNum).addPageBreak();
    return this;
  }

  /** Insert a page break at a specific column number. */
  addColumnPageBreak(colNum: number): this {
    this._ws.getColumn(colNum).addPageBreak();
    return this;
  }

  /** Add an ignored error entry (e.g. number-stored-as-text) for the given reference. */
  addIgnoredError(ref: string, options?: Omit<IgnoredErrorDef, "ref">): this {
    this._ws.ignoredErrors.push({ ref, ...options });
    return this;
  }

  /** Iterate over every row in the sheet. */
  eachRow(callback: (row: ExcelRow, rowNumber: number) => void): void {
    this._ws.eachRow(callback);
  }

  /** Export sheet data as an array of objects or arrays (delegates to excelts). */
  toJSON(): Record<string, ExcelCellValue>[];
  toJSON(opts: SheetToJSONOptions & { header: 1 }): ExcelCellValue[][];
  toJSON(opts: SheetToJSONOptions & { header: "A" }): Record<string, ExcelCellValue>[];
  toJSON(opts: SheetToJSONOptions & { header: string[] }): Record<string, ExcelCellValue>[];
  toJSON(opts?: SheetToJSONOptions): Record<string, ExcelCellValue>[] | ExcelCellValue[][] {
    return this._ws.toJSON(opts);
  }

  /** Export sheet data as an array-of-arrays. */
  toAOA(): ExcelCellValue[][] {
    return this._ws.toAOA();
  }

  /** Import data from an array of objects (delegates to excelts). */
  addJSON(data: Record<string, ExcelCellValue>[], opts?: AddJSONOptions): this {
    this._ws.addJSON(data, opts);
    return this;
  }

  /** Import data from an array-of-arrays (delegates to excelts). */
  addAOA(data: ExcelCellValue[][], opts?: AddAOAOptions): this {
    this._ws.addAOA(data, opts);
    return this;
  }

  private async _doFinalize(): Promise<void> {
    if (this._protectionConfig) {
      await this._ws.protect(this._protectionConfig.password ?? "");
    }
  }

  private _writeHeaders(globalStyle?: CellStyle, height?: number): void {
    if (this._headerWritten) {
      throw new Error(`[SheetBuilder] headers() already called on sheet "${this._opts.name}".`);
    }
    if (this._columns.length === 0) {
      throw new Error(`[SheetBuilder] headers() requires at least one column definition.`);
    }

    const row = this._ws.addRow(this._columns.map((c) => c.header));
    row.eachCell((cell, colNumber) => {
      const col = this._columns[colNumber - 1];
      if (globalStyle) applyStyle(cell, globalStyle);
      if (col?.headerStyle) applyStyle(cell, col.headerStyle);
    });
    row.commit();
    if (height) row.height = height;

    this._headerWritten = true;

    this._ws.columns = this._columns.map((c) => ({
      key: c.key,
      width: c.width ?? 15,
      hidden: c.hidden ?? false,
    }));
  }

  private _applySheetOptions(): void {
    const opts = this._opts;

    if (opts.tabColor) {
      this._ws.properties.tabColor = { argb: opts.tabColor };
    }
    if (opts.defaultRowHeight !== undefined) {
      this._ws.properties.defaultRowHeight = opts.defaultRowHeight;
    }
    if (opts.showGridLines === false || opts.zoom !== undefined) {
      const view: Partial<WorksheetView> = {};
      if (opts.showGridLines === false) view.showGridLines = false;
      if (opts.zoom !== undefined) view.zoomScale = opts.zoom;
      this._ws.views = [view];
    }
    if (opts.state) {
      this._ws.state = opts.state;
    }
    if (opts.freeze) {
      const { row = 0, col = 0 } = opts.freeze;
      this.freeze(row, col);
    }
    if (opts.pageSetup) {
      Object.assign(this._ws.pageSetup, {
        paperSize: opts.pageSetup.paperSize ?? 9,
        orientation: opts.pageSetup.orientation ?? "portrait",
        fitToPage: opts.pageSetup.fitToPage ?? false,
        fitToWidth: opts.pageSetup.fitToWidth ?? 1,
        fitToHeight: opts.pageSetup.fitToHeight ?? 0,
      });
      if (opts.pageSetup.margins) {
        this._ws.pageSetup.margins = {
          left: 0.7,
          right: 0.7,
          top: 0.75,
          bottom: 0.75,
          header: 0.3,
          footer: 0.3,
          ...opts.pageSetup.margins,
        };
      }
      if (opts.pageSetup.printTitlesRow) {
        this._ws.pageSetup.printTitlesRow = opts.pageSetup.printTitlesRow;
      }
      if (opts.pageSetup.printTitlesColumn) {
        this._ws.pageSetup.printTitlesColumn = opts.pageSetup.printTitlesColumn;
      }
    }
    if (opts.headerFooter) {
      const hf = opts.headerFooter;
      if (hf.oddHeader) this._ws.headerFooter.oddHeader = formatHeaderFooterSection(hf.oddHeader);
      if (hf.oddFooter) this._ws.headerFooter.oddFooter = formatHeaderFooterSection(hf.oddFooter);
      if (hf.evenHeader)
        this._ws.headerFooter.evenHeader = formatHeaderFooterSection(hf.evenHeader);
      if (hf.evenFooter)
        this._ws.headerFooter.evenFooter = formatHeaderFooterSection(hf.evenFooter);
    }
    if (opts.protection) {
      this._protectionConfig = { password: opts.protection.password };
    }
  }

  private _flushRow(data: RowData, options?: RowOptions): void {
    let excelRow: ExcelRow;

    if (Array.isArray(data)) {
      const rowValues = data.map(toExcelValue);
      excelRow = this._ws.addRow(rowValues);
    } else {
      if (this._columns.length === 0 && !this._columnsInferred) {
        this._columns = Object.keys(data).map((key) => ({ key, header: key }));
        this._columnsInferred = true;
      }
      const objValues: Record<string, unknown> = {};
      for (const [k, v] of Object.entries(data)) {
        objValues[k] = toExcelValue(v);
      }
      excelRow = this._ws.addRow(objValues);
    }

    if (this._columns.length > 0) {
      excelRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        const col = this._columns[colNumber - 1];
        if (col?.style) applyStyle(cell, col.style);
        if (col?.format) applyStyle(cell, { numberFormat: col.format });
      });
    }

    if (options) {
      if (options.height !== undefined) excelRow.height = options.height;
      if (options.hidden) excelRow.hidden = true;
      if (options.outlineLevel !== undefined) excelRow.outlineLevel = options.outlineLevel;
      if (options.style) {
        const s = options.style;
        excelRow.eachCell({ includeEmpty: true }, (cell) => applyStyle(cell, s));
      }
    }

    excelRow.commit();
    this._rowCount++;
  }

  private _writeValue(cell: ExcelCell, value: CellValue | undefined): void {
    if (value === undefined || value === null) return;
    if (isFormula(value)) {
      cell.value = toFormulaValue(value);
    } else {
      cell.value = value;
    }
  }

  /** Write a formula string into a specific cell by A1 reference. */
  fillFormula(refStr: string, formula: string): this {
    (this._ws.getCell(refStr) as { value: { formula: string } }).value = { formula };
    return this;
  }

  /** Write a formula into a cell by column number and row number (1-based). */
  fillFormulaRC(col: number, row: number, formula: string): this {
    return this.fillFormula(`${colLetter(col)}${row}`, formula);
  }
}

const proto = SheetBuilder.prototype as unknown as SheetBuilderExtension;
applyChartMixin(proto);
applyConditionalFormattingMixin(proto);
applyDataValidationMixin(proto);
applyMediaMixin(proto);
