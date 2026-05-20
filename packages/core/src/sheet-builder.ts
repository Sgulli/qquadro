import type {
  AboveAverageRuleType,
  AddAOAOptions,
  AddJSONOptions,
  CellFormulaValue,
  CellIsOperators,
  CellIsRuleType,
  ColorScaleRuleType,
  ConditionalFormattingOptions,
  ContainsTextOperators,
  ContainsTextRuleType,
  DataBarRuleType,
  DataValidation,
  DataValidationOperator,
  DataValidationWithFormulae,
  Cell as ExcelCell,
  CellValue as ExcelCellValue,
  Row as ExcelRow,
  Style as ExcelStyle,
  Worksheet as ExcelWorksheet,
  ExpressionRuleType,
  IconSetRuleType,
  IconSetTypes,
  SheetToJSONOptions,
  TableColumnProperties,
  TableStyleProperties,
  TimePeriodRuleType,
  TimePeriodTypes,
  Top10RuleType,
  WorksheetView,
} from "@cj-tech-master/excelts";
import type {
  CellPrimitive,
  CellStyle,
  CellValue,
  ColumnDef,
  MergeRange,
  RowData,
  RowOptions,
  SheetOptions,
} from "./types.js";
import {
  applyStyle,
  cellRef,
  colLetter,
  colRange,
  formatHeaderFooterSection,
  rangeRef,
} from "./utils.js";

export class SheetBuilder {
  private readonly _ws: ExcelWorksheet;
  private _columns: ColumnDef[] = [];
  private _headerWritten = false;
  private _protectionConfig?: { password?: string };
  private _rowCount = 0;
  private _columnsInferred = false;

  /** Current number of data rows written (excludes header row). */
  get rowCount(): number {
    return this._rowCount;
  }

  /**
   * Resolve a column key to a range string like `"B2:B5"`.
   * Automatically accounts for header row when headers were written.
   * @param key ColumnDef key
   * @param startRow First data row (defaults to 1 if no headers, 2 after header row)
   */
  /** Get the 1‑based column index for a column key. */
  columnIndex(key: string): number {
    const col = this._columns.find((c) => c.key === key);
    if (!col) throw new Error(`[SheetBuilder] No column with key "${key}".`);
    return this._columns.indexOf(col) + 1;
  }

  columnRange(key: string, startRow?: number): string {
    if (startRow === undefined) startRow = this._headerWritten ? 2 : 1;
    const col = this._columns.find((c) => c.key === key);
    if (!col) throw new Error(`[SheetBuilder] No column with key "${key}".`);
    const idx = this._columns.indexOf(col);
    const letter = colLetter(idx + 1);
    return `${letter}${startRow}:${letter}${startRow + this._rowCount - 1}`;
  }

  /** @internal */
  constructor(
    ws: ExcelWorksheet,
    private readonly _opts: SheetOptions,
  ) {
    this._ws = ws;
    this._applySheetOptions();
  }

  columns(defs: ColumnDef[]): this {
    this._columns = defs;
    return this;
  }

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

  addRow(data: RowData, options?: RowOptions): this {
    this._flushRow(data, options);
    return this;
  }

  addRows(rows: RowData[], options?: RowOptions): this {
    for (const data of rows) this.addRow(data, options);
    return this;
  }

  setCell(address: string, value?: CellValue, style?: CellStyle): this {
    const cell = this._ws.getCell(address);
    this._writeValue(cell, value);
    if (style) applyStyle(cell, style);
    return this;
  }

  /** Set a cell by 1‑based column and row numbers instead of A1 notation. */
  setCellRC(col: number, row: number, value?: CellValue, style?: CellStyle): this {
    return this.setCell(cellRef(col, row), value, style);
  }

  /** Style a range by 1‑based corner coordinates. */
  styleRangeRC(col1: number, row1: number, col2: number, row2: number, style: CellStyle): this {
    return this.styleRange(rangeRef(col1, row1, col2, row2), style);
  }

  styleRange(range: string, style: CellStyle): this {
    const [tl, br] = range.split(":");
    if (!tl) return this;
    const tlCell = this._ws.getCell(tl);
    const brCell = this._ws.getCell(br ?? tl);

    for (let r = tlCell.fullAddress.row; r <= brCell.fullAddress.row; r++) {
      for (let c = tlCell.fullAddress.col; c <= brCell.fullAddress.col; c++) {
        applyStyle(this._ws.getCell(r, c), style);
      }
    }
    return this;
  }

  merge(region: MergeRange): this {
    this._ws.mergeCells(region.range);

    const [tl] = region.range.split(":");
    if (!tl) return this;

    const cell = this._ws.getCell(tl);
    if (region.value) this._writeValue(cell, region.value);
    if (region.style) applyStyle(cell, region.style);
    if (region.height) this._ws.getRow(cell.fullAddress.row).height = region.height;

    return this;
  }

  /** Merge cells by 1‑based coordinates. */
  mergeRC(
    col1: number,
    row1: number,
    col2: number,
    row2: number,
    options?: { value?: CellValue; style?: CellStyle; height?: number },
  ): this {
    return this.merge({
      range: `${cellRef(col1, row1)}:${cellRef(col2, row2)}`,
      ...options,
    });
  }

  mergeAll(regions: MergeRange[]): this {
    for (const r of regions) this.merge(r);
    return this;
  }

  rowHeight(rowNumber: number, height: number): this {
    this._ws.getRow(rowNumber).height = height;
    return this;
  }

  colWidth(col: string | number, width: number): this {
    this._ws.getColumn(col).width = width;
    return this;
  }

  /** Set column width by 1‑based column number. */
  colWidthRC(col: number, width: number): this {
    return this.colWidth(col, width);
  }

  autoFitColumns(startCol?: number | string, endCol?: number | string): this {
    this._ws.autoFitColumns(startCol, endCol);
    return this;
  }

  freeze(row: number, col = 0): this {
    const prev = this._ws.views?.[0];
    this._ws.views = [
      {
        ...(prev ? { showGridLines: prev.showGridLines, zoomScale: prev.zoomScale } : {}),
        state: "frozen",
        xSplit: col,
        ySplit: row,
        topLeftCell: `${colLetter(col + 1)}${row + 1}`,
        activeCell: `${colLetter(col + 1)}${row + 1}`,
      },
    ];
    return this;
  }

  autoFilter(range?: string): this {
    const r = range ?? (this._columns.length ? `A1:${colLetter(this._columns.length)}1` : "A1");
    this._ws.autoFilter = r;
    return this;
  }

  // ── Data Validation ─────────────────────────────────────────────────────────

  addDataValidation(address: string, validation: DataValidation): this {
    this._ws.dataValidations.add(address, validation);
    return this;
  }

  addDataValidationRC(col: number, row: number, validation: DataValidation): this {
    return this.addDataValidation(cellRef(col, row), validation);
  }

  removeDataValidation(address: string): this {
    this._ws.dataValidations.remove(address);
    return this;
  }

  addListValidation(
    address: string,
    list: (string | number | Date)[],
    options?: Omit<DataValidationWithFormulae, "type" | "formulae" | "operator">,
  ): this {
    this._ws.dataValidations.add(address, {
      type: "list",
      formulae: list.map(formatListValue),
      ...options,
    });
    return this;
  }

  addListValidationRC(
    col: number,
    startRow: number,
    endRow: number,
    list: (string | number | Date)[],
    options?: {
      allowBlank?: boolean;
      error?: string;
      errorTitle?: string;
      prompt?: string;
      promptTitle?: string;
      showErrorMessage?: boolean;
      showInputMessage?: boolean;
    },
  ): this {
    return this.addListValidation(colRange(col, startRow, endRow), list, options);
  }

  addRangeValidation(
    address: string,
    type: "whole" | "decimal" | "date" | "textLength",
    operator: DataValidationOperator,
    formulae: (string | number | Date)[],
    options?: Omit<DataValidationWithFormulae, "type" | "formulae" | "operator">,
  ): this {
    this._ws.dataValidations.add(address, {
      type,
      operator,
      formulae,
      ...options,
    });
    return this;
  }

  addRangeValidationRC(
    col: number,
    startRow: number,
    endRow: number,
    type: "whole" | "decimal" | "date" | "textLength",
    operator: DataValidationOperator,
    formulae: (string | number | Date)[],
    options?: {
      allowBlank?: boolean;
      error?: string;
      errorTitle?: string;
      prompt?: string;
      promptTitle?: string;
      showErrorMessage?: boolean;
      showInputMessage?: boolean;
    },
  ): this {
    return this.addRangeValidation(
      colRange(col, startRow, endRow),
      type,
      operator,
      formulae,
      options,
    );
  }

  // ── Conditional Formatting ─────────────────────────────────────────────────

  addConditionalFormatting(cf: ConditionalFormattingOptions): this {
    this._ws.addConditionalFormatting(cf);
    return this;
  }

  removeConditionalFormatting(
    filter?:
      | number
      | ((
          value: ConditionalFormattingOptions,
          index: number,
          array: ConditionalFormattingOptions[],
        ) => boolean),
  ): this {
    this._ws.removeConditionalFormatting(filter);
    return this;
  }

  addCellIsRule(
    ref: string,
    operator: CellIsOperators,
    formulae: (string | number)[],
    style?: Partial<ExcelStyle>,
  ): this {
    const rule: CellIsRuleType = { type: "cellIs", operator, formulae };
    if (style) rule.style = style;
    this._ws.addConditionalFormatting({ ref, rules: [rule] });
    return this;
  }

  addCellIsRuleRC(
    col1: number,
    row1: number,
    col2: number,
    row2: number,
    operator: CellIsOperators,
    formulae: (string | number)[],
    style?: Partial<ExcelStyle>,
  ): this {
    return this.addCellIsRule(rangeRef(col1, row1, col2, row2), operator, formulae, style);
  }

  addExpressionRule(ref: string, formula: string, style?: Partial<ExcelStyle>): this {
    const rule: ExpressionRuleType = {
      type: "expression",
      formulae: [formula],
    };
    if (style) rule.style = style;
    this._ws.addConditionalFormatting({ ref, rules: [rule] });
    return this;
  }

  addExpressionRuleRC(
    col1: number,
    row1: number,
    col2: number,
    row2: number,
    formula: string,
    style?: Partial<ExcelStyle>,
  ): this {
    return this.addExpressionRule(rangeRef(col1, row1, col2, row2), formula, style);
  }

  addDataBar(ref: string, color?: { argb?: string; theme?: number }): this {
    const rule: DataBarRuleType = { type: "dataBar" };
    if (color) rule.color = color;
    this._ws.addConditionalFormatting({ ref, rules: [rule] });
    return this;
  }

  addDataBarRC(
    col1: number,
    row1: number,
    col2: number,
    row2: number,
    color?: { argb?: string; theme?: number },
  ): this {
    return this.addDataBar(rangeRef(col1, row1, col2, row2), color);
  }

  addColorScale(
    ref: string,
    cfvo: {
      type: "min" | "max" | "num" | "percent" | "percentile" | "formula";
      value?: number | string;
    }[],
    colors?: { argb?: string; theme?: number }[],
  ): this {
    const rule: ColorScaleRuleType = {
      type: "colorScale",
      cfvo,
      color: colors,
    };
    this._ws.addConditionalFormatting({ ref, rules: [rule] });
    return this;
  }

  addColorScaleRC(
    col1: number,
    row1: number,
    col2: number,
    row2: number,
    cfvo: {
      type: "min" | "max" | "num" | "percent" | "percentile" | "formula";
      value?: number | string;
    }[],
    colors?: { argb?: string; theme?: number }[],
  ): this {
    return this.addColorScale(rangeRef(col1, row1, col2, row2), cfvo, colors);
  }

  addIconSet(
    ref: string,
    iconSet?: IconSetTypes,
    cfvo?: {
      type: "percent" | "num" | "percentile" | "formula";
      value?: number | string;
    }[],
    options?: { showValue?: boolean; reverse?: boolean },
  ): this {
    const rule: IconSetRuleType = {
      type: "iconSet",
      iconSet,
      cfvo,
      ...options,
    };
    this._ws.addConditionalFormatting({ ref, rules: [rule] });
    return this;
  }

  addIconSetRC(
    col1: number,
    row1: number,
    col2: number,
    row2: number,
    iconSet?: IconSetTypes,
    cfvo?: {
      type: "percent" | "num" | "percentile" | "formula";
      value?: number | string;
    }[],
    options?: { showValue?: boolean; reverse?: boolean },
  ): this {
    return this.addIconSet(rangeRef(col1, row1, col2, row2), iconSet, cfvo, options);
  }

  addTop10Rule(
    ref: string,
    rank: number,
    options?: {
      percent?: boolean;
      bottom?: boolean;
      style?: Partial<ExcelStyle>;
    },
  ): this {
    const rule: Top10RuleType = {
      type: "top10",
      rank,
      percent: options?.percent ?? false,
      bottom: options?.bottom,
    };
    if (options?.style) rule.style = options.style;
    this._ws.addConditionalFormatting({ ref, rules: [rule] });
    return this;
  }

  addTop10RuleRC(
    col1: number,
    row1: number,
    col2: number,
    row2: number,
    rank: number,
    options?: {
      percent?: boolean;
      bottom?: boolean;
      style?: Partial<ExcelStyle>;
    },
  ): this {
    return this.addTop10Rule(rangeRef(col1, row1, col2, row2), rank, options);
  }

  addAboveAverageRule(
    ref: string,
    options?: { aboveAverage?: boolean; style?: Partial<ExcelStyle> },
  ): this {
    const rule: AboveAverageRuleType = {
      type: "aboveAverage",
      aboveAverage: options?.aboveAverage,
    };
    if (options?.style) rule.style = options.style;
    this._ws.addConditionalFormatting({ ref, rules: [rule] });
    return this;
  }

  addAboveAverageRuleRC(
    col1: number,
    row1: number,
    col2: number,
    row2: number,
    options?: { aboveAverage?: boolean; style?: Partial<ExcelStyle> },
  ): this {
    return this.addAboveAverageRule(rangeRef(col1, row1, col2, row2), options);
  }

  addContainsTextRule(
    ref: string,
    text: string,
    operator?: ContainsTextOperators,
    style?: Partial<ExcelStyle>,
  ): this {
    const rule: ContainsTextRuleType = {
      type: "containsText",
      operator,
      text,
    };
    if (style) rule.style = style;
    this._ws.addConditionalFormatting({ ref, rules: [rule] });
    return this;
  }

  addContainsTextRuleRC(
    col1: number,
    row1: number,
    col2: number,
    row2: number,
    text: string,
    operator?: ContainsTextOperators,
    style?: Partial<ExcelStyle>,
  ): this {
    return this.addContainsTextRule(rangeRef(col1, row1, col2, row2), text, operator, style);
  }

  addTimePeriodRule(ref: string, timePeriod: TimePeriodTypes, style?: Partial<ExcelStyle>): this {
    const rule: TimePeriodRuleType = {
      type: "timePeriod",
      timePeriod,
    };
    if (style) rule.style = style;
    this._ws.addConditionalFormatting({ ref, rules: [rule] });
    return this;
  }

  addTimePeriodRuleRC(
    col1: number,
    row1: number,
    col2: number,
    row2: number,
    timePeriod: TimePeriodTypes,
    style?: Partial<ExcelStyle>,
  ): this {
    return this.addTimePeriodRule(rangeRef(col1, row1, col2, row2), timePeriod, style);
  }

  // ── Tables ──────────────────────────────────────────────────────────────────

  addTable(
    name: string,
    ref: string,
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
      ref,
      columns,
      rows: options?.rows ?? [],
      headerRow: options?.headerRow ?? true,
      totalsRow: options?.totalsRow,
      style: options?.style,
    });
    return this;
  }

  addTableRC(
    name: string,
    col1: number,
    row1: number,
    col2: number,
    row2: number,
    columns: TableColumnProperties[],
    options?: {
      rows?: ExcelCellValue[][];
      headerRow?: boolean;
      totalsRow?: boolean;
      style?: TableStyleProperties;
    },
  ): this {
    const ref = rangeRef(col1, row1, col2, row2);
    return this.addTable(name, ref, columns, options);
  }

  // ── Reading / Export ───────────────────────────────────────────────────────

  eachRow(callback: (row: ExcelRow, rowNumber: number) => void): void {
    this._ws.eachRow(callback);
  }

  toJSON(): Record<string, ExcelCellValue>[];
  toJSON(opts: SheetToJSONOptions & { header: 1 }): ExcelCellValue[][];
  toJSON(opts: SheetToJSONOptions & { header: "A" }): Record<string, ExcelCellValue>[];
  toJSON(opts: SheetToJSONOptions & { header: string[] }): Record<string, ExcelCellValue>[];
  toJSON(opts?: SheetToJSONOptions): Record<string, ExcelCellValue>[] | ExcelCellValue[][] {
    return this._ws.toJSON(opts);
  }

  toAOA(): ExcelCellValue[][] {
    return this._ws.toAOA();
  }

  addJSON(data: Record<string, ExcelCellValue>[], opts?: AddJSONOptions): this {
    this._ws.addJSON(data, opts);
    return this;
  }

  addAOA(data: ExcelCellValue[][], opts?: AddAOAOptions): this {
    this._ws.addAOA(data, opts);
    return this;
  }

  // ── Finalization ───────────────────────────────────────────────────────────

  /** @internal */
  async _finalize(): Promise<void> {
    if (this._protectionConfig) {
      await this._ws.protect(this._protectionConfig.password ?? "");
    }
  }

  // ── Private helpers ────────────────────────────────────────────────────────

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
      });
    }

    if (options) {
      if (options.height !== undefined) excelRow.height = options.height;
      if (options.hidden) excelRow.hidden = true;
      if (options.outlineLevel !== undefined) excelRow.outlineLevel = options.outlineLevel;
      if (options.style) {
        excelRow.eachCell({ includeEmpty: true }, (cell) => applyStyle(cell, options.style!));
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
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function formatListValue(v: string | number | Date): string | number | Date {
  if (typeof v === "string") return `"${v.replace(/"/g, '""')}"`;
  if (v instanceof Date) {
    const y = v.getFullYear();
    const m = String(v.getMonth() + 1).padStart(2, "0");
    const d = String(v.getDate()).padStart(2, "0");
    return `"${y}-${m}-${d}"`;
  }
  return v;
}

function isFormula(val: CellValue): val is { formula: string; result?: CellPrimitive } {
  return (
    typeof val === "object" &&
    val !== null &&
    !(val instanceof Date) &&
    "formula" in val &&
    typeof (val as { formula: unknown }).formula === "string"
  );
}

function normalizeFormula(f: string): string {
  return f.startsWith("=") ? f.slice(1) : f;
}

function toFormulaValue(v: { formula: string; result?: CellPrimitive }): CellFormulaValue {
  const fv: CellFormulaValue = { formula: normalizeFormula(v.formula) };
  if (v.result !== undefined && v.result !== null) fv.result = v.result;
  return fv;
}

function toExcelValue(val: CellValue): CellPrimitive | CellFormulaValue {
  if (isFormula(val)) return toFormulaValue(val);
  return val as CellPrimitive;
}
