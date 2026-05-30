import fs from "node:fs";
import path from "node:path";
import {
  type DefinedNameModel,
  Workbook as ExcelWorkbook,
  type ExternalLinkModel,
  type ImageData,
} from "@cj-tech-master/excelts";
import { colLetter } from "./coords.js";
import { _sheetFinalizers, SheetBuilder } from "./sheet-builder.js";
import type { ExternalLinkInput, SheetOptions, WorkbookOptions, WriteResult } from "./types.js";

export class WorkbookBuilder {
  private _wb: ExcelWorkbook;
  private _opts: WorkbookOptions;
  private _sheets: SheetBuilder[] = [];

  constructor(opts: WorkbookOptions = {}) {
    this._opts = opts;
    this._wb = new ExcelWorkbook();
    this._applyWorkbookMeta();
  }

  static async load(filePath: string): Promise<WorkbookBuilder> {
    const resolved = path.resolve(filePath);
    const wb = new ExcelWorkbook();
    await wb.xlsx.readFile(resolved);
    return WorkbookBuilder._withWb(wb);
  }

  static fromFile(filePath: string): Promise<WorkbookBuilder> {
    return WorkbookBuilder.load(filePath);
  }

  static async fromCsv(data: string | ArrayBuffer): Promise<WorkbookBuilder> {
    const wb = new ExcelWorkbook();
    await wb.readCsv(data);
    return WorkbookBuilder._withWb(wb);
  }

  static async fromCsvFile(filePath: string): Promise<WorkbookBuilder> {
    const wb = new ExcelWorkbook();
    await wb.readCsvFile(filePath);
    return WorkbookBuilder._withWb(wb);
  }

  private static _withWb(wb: ExcelWorkbook): WorkbookBuilder {
    const builder = new WorkbookBuilder();
    builder._wb = wb;
    for (let i = 0; i < wb.worksheets.length; i++) {
      const ws = wb.getWorksheet(i + 1);
      if (!ws) continue;
      builder._sheets.push(
        new SheetBuilder(ws, { name: wb.worksheets[i]?.name ?? `Sheet${i + 1}` }),
      );
    }
    return builder;
  }

  addSheet(opts: SheetOptions, configure: (sheet: SheetBuilder) => void): this {
    const ws = this._wb.addWorksheet(opts.name);
    const builder = new SheetBuilder(ws, opts);
    this._sheets.push(builder);

    if (configure) {
      configure(builder);
    }

    return this;
  }

  getSheet(name: string): SheetBuilder | undefined {
    return this._sheets.find((s) => s.name === name);
  }

  sheet(index: number): SheetBuilder {
    const s = this._sheets[index];
    if (!s) throw new Error(`[WorkbookBuilder] No sheet at index ${index}.`);
    return s;
  }

  async write(outputPath: string): Promise<WriteResult> {
    const resolved = this._safeResolve(outputPath);
    const dir = path.dirname(resolved);

    await this._finalizeAll();
    await fs.promises.mkdir(dir, { recursive: true });

    if (this._opts.useStreaming) {
      const stream = fs.createWriteStream(resolved);
      await this._wb.xlsx.write(stream, {
        useSharedStrings: this._opts.useSharedStrings,
      });
      await new Promise<void>((ok, err) => {
        stream.on("finish", ok);
        stream.on("error", err);
      });
      const { size } = await fs.promises.stat(resolved);
      return { filePath: resolved, sizeBytes: size };
    }

    await this._wb.xlsx.writeFile(resolved, {
      useSharedStrings: this._opts.useSharedStrings,
    });
    const { size } = await fs.promises.stat(resolved);
    return { filePath: resolved, sizeBytes: size };
  }

  async toBuffer(): Promise<Buffer> {
    await this._finalizeAll();
    const buf = await this._wb.xlsx.writeBuffer({
      useSharedStrings: this._opts.useSharedStrings,
    });
    return Buffer.from(buf);
  }

  async toCsvString(): Promise<string> {
    return this._wb.writeCsv();
  }

  async writeCsv(outputPath: string): Promise<void> {
    await this._wb.writeCsvFile(path.resolve(outputPath));
  }

  private _safeResolve(filePath: string): string {
    const resolved = path.resolve(filePath);
    if (this._opts.allowedBase) {
      const base = path.resolve(this._opts.allowedBase);
      if (!resolved.startsWith(base + path.sep) && resolved !== base) {
        throw new Error(`Path "${filePath}" is outside allowed directory "${base}"`);
      }
    }
    return resolved;
  }

  get sheets(): SheetBuilder[] {
    return this._sheets;
  }

  get workbook(): ExcelWorkbook {
    return this._wb;
  }

  /** Register a named range (workbook‑level by default). */
  defineName(
    name: string,
    address: string,
    sheetName?: string,
    options?: { hidden?: boolean },
  ): this {
    const loc = sheetName ? `${sheetName}!${address}` : address;
    if (options?.hidden) {
      this._wb.definedNames.addHidden(loc, name);
    } else {
      this._wb.definedNames.add(loc, name);
    }
    return this;
  }

  /** Register a named range by 1‑based coordinates. */
  defineNameRC(
    name: string,
    col1: number,
    row1: number,
    col2: number,
    row2: number,
    sheetName?: string,
    options?: { hidden?: boolean },
  ): this {
    const address = `${colLetter(col1)}${row1}:${colLetter(col2)}${row2}`;
    return this.defineName(name, address, sheetName, options);
  }

  /** Register a formula‑based defined name (e.g. `"LAMBDA(x,y,x+y)"`). */
  defineFormula(name: string, expression: string): this {
    this._wb.definedNames.addFormula(name, expression);
    return this;
  }

  /** Remove a defined name by address and name. */
  removeDefinedName(name: string, address: string): this {
    this._wb.definedNames.remove(address, name);
    return this;
  }

  /** Get all defined names. */
  getDefinedNames(): DefinedNameModel[] {
    return this._wb.definedNames.getAllEntries();
  }

  /** Add an image to the workbook and return its numeric id. */
  addImage(image: ImageData): number {
    return this._wb.addImage(image);
  }

  /** Register a person for threaded comments. Returns the person id. */
  registerPerson(displayName: string, userId?: string, providerId?: string): string {
    return this._wb.registerPerson(displayName, userId, providerId);
  }

  /** Remove a sheet by name. */
  removeSheet(name: string): this {
    const sheet = this.getSheet(name);
    if (!sheet) throw new Error(`[WorkbookBuilder] No sheet named "${name}".`);
    this._wb.removeWorksheetEx(sheet.worksheet);
    this._sheets = this._sheets.filter((s) => s.name !== name);
    return this;
  }

  /** Duplicate an existing sheet with an optional new name. */
  duplicateSheet(name: string, newName?: string): this {
    const source = this.getSheet(name);
    if (!source) throw new Error(`[WorkbookBuilder] No sheet named "${name}".`);
    const targetName =
      newName ?? `${name} (${this._sheets.filter((s) => s.name.startsWith(name)).length})`;
    const newWs = this._wb.importSheet(source.worksheet, targetName);
    const builder = new SheetBuilder(newWs, { name: targetName });
    this._sheets.push(builder);
    return this;
  }

  /** Register an external workbook reference for cross-workbook formulas. */
  addExternalLink(input: ExternalLinkInput): ExternalLinkModel {
    return this._wb.addExternalLink(input);
  }

  /** Get all registered external links. */
  getExternalLinks(): ExternalLinkModel[] {
    return this._wb.externalLinks;
  }

  /** Trigger full formula recalculation on next open in Excel. */
  calculate(): this {
    this._wb.calcProperties.fullCalcOnLoad = true;
    return this;
  }

  /** Register a custom function (UDF) for use in formulas. */
  registerFunction(name: string, fn: (...args: unknown[]) => unknown): this {
    const map = this._wb.userFunctions as
      | Map<string, { invoke: (...args: unknown[]) => unknown }>
      | undefined;
    if (map) map.set(name, { invoke: fn });
    return this;
  }

  /** Import a markdown table string as a new sheet. */
  addSheetFromMarkdown(markdown: string, sheetName: string): this {
    const lines = markdown.split("\n").filter((l) => l.trim());
    if (lines.length < 2) return this;

    const headers = lines[0]
      .split("|")
      .map((h) => h.trim())
      .filter(Boolean);
    const dataRows = lines.slice(2);

    this.addSheet({ name: sheetName }, (sheet) => {
      sheet.headers(headers.map((h) => ({ key: h.toLowerCase().replace(/\s+/g, "_"), header: h })));
      for (const row of dataRows) {
        const vals = row
          .split("|")
          .map((c) => c.trim())
          .filter(Boolean);
        if (vals.length === 0) continue;
        const obj: Record<string, string> = {};
        for (let i = 0; i < Math.min(vals.length, headers.length); i++) {
          obj[headers[i].toLowerCase().replace(/\s+/g, "_")] = vals[i];
        }
        sheet.addRow(obj);
      }
    });
    return this;
  }

  /** Export first sheet as a markdown table. */
  async toMarkdown(sheetIndex = 0): Promise<string> {
    const sheet = this._sheets[sheetIndex];
    if (!sheet) return "";

    const allRows: string[][] = [];
    sheet.eachRow((row, _rowNumber) => {
      const vals: string[] = [];
      row.eachCell({ includeEmpty: true }, (cell: { value: unknown }) => {
        vals.push(String(cell.value ?? ""));
      });
      allRows.push(vals);
    });

    if (allRows.length === 0) return "";

    const colCount = Math.max(...allRows.map((r) => r.length));
    const colWidths = Array.from({ length: colCount }, (_, ci) =>
      Math.max(...allRows.map((r) => (r[ci] ?? "").length), 3),
    );

    const fmt = (row: string[]) => `| ${row.map((v, ci) => v.padEnd(colWidths[ci])).join(" | ")} |`;

    const sep = `| ${colWidths.map((w) => "-".repeat(w)).join(" | ")} |`;

    return [fmt(allRows[0]), sep, ...allRows.slice(1).map(fmt)].join("\n");
  }

  private async _finalizeAll(): Promise<void> {
    const results = await Promise.allSettled(this._sheets.map((s) => _sheetFinalizers.get(s)?.()));
    const errors = results
      .filter((r): r is PromiseRejectedResult => r.status === "rejected")
      .map((r) => r.reason);
    if (errors.length > 0) {
      throw new AggregateError(
        errors,
        `[WorkbookBuilder] ${errors.length} sheet(s) failed to finalize.`,
      );
    }
  }

  private _applyWorkbookMeta(): void {
    const {
      author,
      company,
      created,
      title,
      subject,
      keywords,
      category,
      manager,
      description,
      language,
    } = this._opts;
    if (author) this._wb.creator = author;
    if (company) this._wb.company = company;
    if (title) this._wb.title = title;
    if (subject) this._wb.subject = subject;
    if (keywords) this._wb.keywords = keywords;
    if (category) this._wb.category = category;
    if (manager) this._wb.manager = manager;
    if (description) this._wb.description = description;
    if (language) this._wb.language = language;
    this._wb.created = created ?? new Date();
    this._wb.calcProperties.fullCalcOnLoad = true;
  }
}
