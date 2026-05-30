import { execFileSync } from "node:child_process";
import fs from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { defineCommand } from "citty";
import { XMLParser } from "fast-xml-parser";

export interface DebugResult {
  calcPr: { calcId: string; fullCalcOnLoad: string };
  sheets: Array<{ name: string; cells: Array<{ ref: string; formula: string; cached: string }> }>;
}

const ALLOWED_EXTENSIONS = [".xlsx", ".xlsm", ".xltx"];

export function handler(filePath: string): DebugResult {
  const xlsxPath = path.resolve(filePath);

  const ext = path.extname(xlsxPath).toLowerCase();
  if (!ALLOWED_EXTENSIONS.includes(ext)) {
    throw new Error(`Unsupported file type: "${ext}". Expected .xlsx`);
  }

  const buf = Buffer.alloc(4);
  const fd = fs.openSync(xlsxPath, "r");
  fs.readSync(fd, buf, 0, 4, 0);
  fs.closeSync(fd);
  if (buf[0] !== 0x50 || buf[1] !== 0x4b || buf[2] !== 0x03 || buf[3] !== 0x04) {
    throw new Error(
      `Invalid file: "${filePath}" is not a valid ZIP/XLSX file (missing PK\x03\x04 header).`,
    );
  }

  const tmp = fs.mkdtempSync(path.join(tmpdir(), "quadro-debug-"));

  try {
    execFileSync(
      "unzip",
      [
        "-o",
        xlsxPath,
        "xl/worksheets/sheet*.xml",
        "xl/workbook.xml",
        "xl/sharedStrings.xml",
        "-d",
        tmp,
      ],
      { stdio: "ignore" },
    );

    const parser = new XMLParser({
      ignoreAttributes: false,
      attributeNamePrefix: "",
      isArray: (name: string) => name === "row" || name === "c" || name === "si",
    });

    const wb = parser.parse(fs.readFileSync(path.join(tmp, "xl/workbook.xml"), "utf-8")).workbook;
    const calc = wb.calcPr ?? {};

    const ss: string[] = [];
    const ssPath = path.join(tmp, "xl/sharedStrings.xml");
    if (fs.existsSync(ssPath)) {
      const items = parser.parse(fs.readFileSync(ssPath, "utf-8"))?.sst?.si ?? [];
      for (const si of items) {
        ss.push(si?.t ?? si?.r?.t ?? "");
      }
    }

    const sheets: DebugResult["sheets"] = [];
    const sheetFiles = fs.readdirSync(path.join(tmp, "xl/worksheets")).sort();
    for (const sf of sheetFiles) {
      const sheetName = path.basename(sf, ".xml");
      const rows =
        parser.parse(fs.readFileSync(path.join(tmp, "xl/worksheets", sf), "utf-8")).worksheet
          ?.sheetData?.row ?? [];
      const cells: DebugResult["sheets"][number]["cells"] = [];

      for (const row of Array.isArray(rows) ? rows : [rows]) {
        const rowCells = Array.isArray(row.c) ? row.c : row.c ? [row.c] : [];
        for (const c of rowCells) {
          const formula =
            c.f != null ? (typeof c.f === "object" ? (c.f["#text"] ?? "") : c.f) : null;
          if (formula == null) continue;
          const cached =
            c.v != null ? (c.t === "s" ? `"${ss[Number(c.v)] ?? "???"}"` : c.v) : "(none)";
          cells.push({ ref: c.r, formula, cached });
        }
      }
      sheets.push({ name: sheetName, cells });
    }

    return {
      calcPr: { calcId: calc.calcId ?? "N/A", fullCalcOnLoad: calc.fullCalcOnLoad ?? "N/A" },
      sheets,
    };
  } finally {
    fs.rmSync(tmp, { recursive: true, force: true });
  }
}

export const debug = defineCommand({
  meta: {
    name: "debug",
    description: "Inspect formulas in a .xlsx file",
  },
  args: {
    file: {
      type: "positional",
      description: "Path to .xlsx file (default: output/demo-report.xlsx)",
      default: "output/demo-report.xlsx",
      required: false,
    },
  },
  run({ args }) {
    try {
      const result = handler(args.file);
      console.log(
        `calcPr: calcId=${result.calcPr.calcId} fullCalcOnLoad=${result.calcPr.fullCalcOnLoad}\n`,
      );
      for (const sheet of result.sheets) {
        console.log(`=== ${sheet.name} ===`);
        for (const cell of sheet.cells) {
          console.log(`  ${cell.ref}  ${cell.formula}  [cached=${cell.cached}]`);
        }
        console.log();
      }
    } catch (err) {
      console.error("Debug failed:", err instanceof Error ? err.message : String(err));
      process.exit(1);
    }
  },
});
