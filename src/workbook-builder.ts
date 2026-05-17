import { Workbook as ExcelWorkbook } from "@cj-tech-master/excelts";
import fs from "node:fs";
import path from "node:path";
import { SheetBuilder } from "./sheet-builder.js";
import type { SheetOptions, WorkbookOptions, WriteResult } from "./types.js";

/**
 * Top-level fluent builder for an Excel workbook.
 *
 * Supports two write modes:
 * - **Buffer mode** (default): builds in memory, then writes to disk.
 * - **Stream mode**: writes directly to a file via a Node.js write stream —
 *   constant memory regardless of file size, ideal for large datasets.
 *
 * @example
 * ```ts
 * import { WorkbookBuilder, Styles } from "./index";
 *
 * const result = await new WorkbookBuilder({ author: "Acme Corp" })
 *   .addSheet({ name: "Sales" }, (sheet) => {
 *     sheet
 *       .columns([
 *         { key: "region", header: "Region", width: 20, headerStyle: Styles.header },
 *         { key: "revenue", header: "Revenue ($)", width: 18, style: Styles.currency, headerStyle: Styles.header },
 *       ])
 *       .writeHeaders()
 *       .addRows([
 *         { region: "EMEA",  revenue: 120_000 },
 *         { region: "APAC",  revenue: 98_500  },
 *         { region: "AMER",  revenue: 210_000 },
 *       ])
 *       .addRow({ region: "Total", revenue: { formula: "=SUM(B2:B4)" } }, { style: Styles.totalRow })
 *       .autoFitColumns()
 *       .freeze(1);
 *   })
 *   .write("./output/report.xlsx");
 *
 * console.log("Written to", result.filePath, `(${result.sizeBytes} bytes)`);
 * ```
 */
export class WorkbookBuilder {
  private readonly _wb: ExcelWorkbook;
  private readonly _opts: WorkbookOptions;
  private readonly _sheets: SheetBuilder[] = [];

  constructor(opts: WorkbookOptions = {}) {
    this._opts = opts;
    this._wb = new ExcelWorkbook();
    this._applyWorkbookMeta();
  }

  // ── Sheet API ──────────────────────────────────────────────────────────────

  /**
   * Add a sheet using a callback-based builder pattern.
   *
   * The callback receives a fully-typed `SheetBuilder`. No need to call
   * `.done()` — the parent regains control automatically after the callback.
   *
   * @example
   * builder.addSheet({ name: "Summary" }, (sheet) => {
   *   sheet.columns([...]).writeHeaders().addRows([...]);
   * })
   */
  addSheet(opts: SheetOptions, configure: (sheet: SheetBuilder) => void): this;

  /**
   * Add a sheet and get back the `SheetBuilder` for imperative use.
   */
  addSheet(opts: SheetOptions): SheetBuilder;

  addSheet(
    opts: SheetOptions,
    configure?: (sheet: SheetBuilder) => void,
  ): this | SheetBuilder {
    const ws = this._wb.addWorksheet(opts.name);
    const builder = new SheetBuilder(ws, opts);
    this._sheets.push(builder);

    if (configure) {
      configure(builder);
      return this;
    }

    return builder;
  }

  // ── Write API ──────────────────────────────────────────────────────────────

  /**
   * Write the workbook to disk.
   *
   * - Uses **streaming mode** when `opts.useStreaming` was set in the
   *   constructor — O(1) memory for arbitrarily large files.
   * - Otherwise writes via an in-memory buffer.
   *
   * @param outputPath  Destination file path (.xlsx).
   */
  async write(outputPath: string): Promise<WriteResult> {
    const resolved = path.resolve(outputPath);
    const dir = path.dirname(resolved);

    await this._finalizeAll();

    // Ensure the target directory exists
    await fs.promises.mkdir(dir, { recursive: true });

    if (this._opts.useStreaming) {
      // ── Streaming path ─────────────────────────────────────────────────────
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

    // ── Buffer path ────────────────────────────────────────────────────────
    await this._wb.xlsx.writeFile(resolved, {
      useSharedStrings: this._opts.useSharedStrings,
    });
    const { size } = await fs.promises.stat(resolved);
    return { filePath: resolved, sizeBytes: size };
  }

  /**
   * Return the workbook as a `Buffer` (e.g. for HTTP responses).
   */
  async toBuffer(): Promise<Buffer> {
    await this._finalizeAll();
    const buf = await this._wb.xlsx.writeBuffer({
      useSharedStrings: this._opts.useSharedStrings,
    });
    return Buffer.from(buf);
  }

  /**
   * Return the underlying ExcelTS `Workbook` for advanced operations not
   * covered by this wrapper.
   */
  get workbook(): ExcelWorkbook {
    return this._wb;
  }

  // ── Private ────────────────────────────────────────────────────────────────

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
