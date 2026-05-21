import fs from "node:fs";
import path from "node:path";
import {
  type DefinedNameModel,
  Workbook as ExcelWorkbook,
  type ImageData,
} from "@cj-tech-master/excelts";
import { SheetBuilder } from "./sheet-builder.js";
import type { SheetOptions, WorkbookOptions, WriteResult } from "./types.js";
import { colLetter } from "./utils.js";

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
    const builder = new WorkbookBuilder();
    builder._wb = wb;
    for (let i = 0; i < wb.worksheets.length; i++) {
      const ws = wb.getWorksheet(i + 1);
      if (!ws) continue;
      builder._sheets.push(new SheetBuilder(ws, { name: wb.worksheets[i]?.name }));
    }
    return builder;
  }

  static fromFile(filePath: string): Promise<WorkbookBuilder> {
    return WorkbookBuilder.load(filePath);
  }

  static async fromCsv(data: string | ArrayBuffer): Promise<WorkbookBuilder> {
    const wb = new ExcelWorkbook();
    await wb.readCsv(data);
    const builder = new WorkbookBuilder();
    builder._wb = wb;
    for (let i = 0; i < wb.worksheets.length; i++) {
      const ws = wb.getWorksheet(i + 1);
      if (!ws) continue;
      builder._sheets.push(new SheetBuilder(ws, { name: wb.worksheets[i].name }));
    }
    return builder;
  }

  static async fromCsvFile(filePath: string): Promise<WorkbookBuilder> {
    const wb = new ExcelWorkbook();
    await wb.readCsvFile(filePath);
    const builder = new WorkbookBuilder();
    builder._wb = wb;
    for (let i = 0; i < wb.worksheets.length; i++) {
      const ws = wb.getWorksheet(i + 1);
      if (!ws) continue;
      builder._sheets.push(new SheetBuilder(ws, { name: wb.worksheets[i].name }));
    }
    return builder;
  }

  addSheet(opts: SheetOptions, configure: (sheet: SheetBuilder) => void): this;
  addSheet(opts: SheetOptions): SheetBuilder;
  addSheet(opts: SheetOptions, configure?: (sheet: SheetBuilder) => void): this | SheetBuilder {
    const ws = this._wb.addWorksheet(opts.name);
    const builder = new SheetBuilder(ws, opts);
    this._sheets.push(builder);

    if (configure) {
      configure(builder);
      return this;
    }

    return builder;
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

  async toCsv(outputPath?: string): Promise<string | undefined> {
    if (outputPath) {
      await this._wb.writeCsvFile(path.resolve(outputPath));
      return;
    }
    return this._wb.writeCsv();
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

  private async _finalizeAll(): Promise<void> {
    await Promise.all(this._sheets.map((s) => s._finalize()));
  }

  private _applyWorkbookMeta(): void {
    const { author, company, created } = this._opts;
    if (author) this._wb.creator = author;
    if (company) this._wb.company = company;
    this._wb.created = created ?? new Date();
    this._wb.calcProperties.fullCalcOnLoad = true;
  }
}
