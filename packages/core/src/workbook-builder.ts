import fs from "node:fs";
import path from "node:path";
import type { Worksheet as ExcelWorksheet } from "@cj-tech-master/excelts";
import { Workbook as ExcelWorkbook } from "@cj-tech-master/excelts";
import { SheetBuilder } from "./sheet-builder.js";
import type { SheetOptions, WorkbookOptions, WriteResult } from "./types.js";

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
      builder._sheets.push(new SheetBuilder(ws, { name: wb.worksheets[i]!.name }));
    }
    return builder;
  }

  static fromFile(filePath: string): Promise<WorkbookBuilder> {
    return WorkbookBuilder.load(filePath);
  }

  static async fromCsv(
    input: string | ArrayBuffer,
    options?: { fileName?: string },
  ): Promise<WorkbookBuilder> {
    const wb = new ExcelWorkbook();
    if (options?.fileName) {
      await wb.readCsvFile(options.fileName);
    } else {
      await wb.readCsv(input);
    }
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
    const resolved = path.resolve(outputPath);
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

  async toCsv(outputPath?: string): Promise<string | void> {
    if (outputPath) {
      await this._wb.writeCsvFile(path.resolve(outputPath));
      return;
    }
    return this._wb.writeCsv();
  }

  get sheets(): SheetBuilder[] {
    return this._sheets;
  }

  get workbook(): ExcelWorkbook {
    return this._wb;
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
