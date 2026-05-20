import path from "node:path";
import type { WriteResult } from "@quadro/core";
import { align, border, F, style, WorkbookBuilder } from "@quadro/core";
import { defineCommand } from "citty";
import {
  mockClubs,
  mockMacroareas,
  mockResultFreeSeats,
  mockSeatRequests,
} from "../data/mock-data.js";

function colLetter(n: number): string {
  let r = "";
  let num = n + 1;
  while (num > 0) {
    const rem = (num - 1) % 26;
    r = String.fromCharCode(65 + rem) + r;
    num = Math.floor((num - 1) / 26);
  }
  return r;
}

function diffFormula(cl: string, rowA: number, rowB: number): { formula: string } {
  return { formula: `${cl}${rowA}-${cl}${rowB}` };
}

function colRange(colIndex: number, startRow: number, endRow: number): string {
  const c = colLetter(colIndex);
  return `${c}${startRow}:${c}${endRow}`;
}

const darkHeader = style(
  { font: { bold: true, size: 11, color: "FFFFFFFF", name: "Arial" } },
  { fill: { type: "solid", color: "FF2B579A" } },
  align.centerWrap,
  border.thinBlack,
);

const subHeader = style(
  { font: { bold: true, size: 10, color: "FF1F497D", name: "Arial" } },
  { fill: { type: "solid", color: "FFDCE6F1" } },
  align.centerWrap,
  border.thinBlack,
);

const dataCell = style({ font: { name: "Arial", size: 10 } }, align.center, border.thinBlack);
const clubNameCell = style(dataCell, { font: { bold: true } }, align.left);
const footerLabel = style({ font: { bold: true, size: 10, name: "Arial" } }, align.right);
const footerCell = style(dataCell, { font: { bold: true } });

const HEADER_ROWS = 3;
const FIXED_COLS = 3;
const DATA_ROW_START = 3;
const EMPTY_SEPARATOR = 1;

export interface AssignmentsOptions {
  outputDir?: string;
}

export async function handler(options: AssignmentsOptions = {}): Promise<WriteResult> {
  const macroareasWithSectors = mockMacroareas.map((ma) => ({
    id: ma.id,
    name: ma.name,
    sectors: ma.sectors.map((s) => ({ id: s.id, name: s.name })),
  }));

  const clubData = new Map<string, { name: string; totalRequests: number }>();
  const clubsById = new Map(mockClubs.map((c) => [c.id, c]));
  for (const club of mockClubs) {
    if (!club.name) continue;
    clubData.set(club.id, { name: club.name, totalRequests: 0 });
  }
  for (const { clubId, macroAreaId, sectorId, quantity } of mockSeatRequests) {
    if (!clubId || !macroAreaId) continue;
    const club = clubData.get(clubId);
    if (!club) continue;
    const fullClub = clubsById.get(clubId);
    const clubImaSeats =
      (
        fullClub?.seats as Array<{
          originalClubId: string;
          sectorId?: string;
          sector?: { macroAreaId: string };
        }>
      ).filter((s) => {
        if (s.originalClubId !== clubId) return false;
        if (sectorId) return s.sectorId === sectorId;
        return s.sector?.macroAreaId === macroAreaId;
      }) ?? [];
    club.totalRequests += Math.max(0, quantity - clubImaSeats.length);
  }

  const orderedColumns: Array<{
    macroareaName: string;
    macroareaId: string;
    sectorName: string | null;
    sectorId: string | null;
    hasSectors: boolean;
  }> = [];
  for (const ma of macroareasWithSectors) {
    if (ma.sectors.length > 0) {
      for (const sector of ma.sectors) {
        orderedColumns.push({
          macroareaName: ma.name,
          macroareaId: ma.id,
          sectorName: sector.name,
          sectorId: sector.id,
          hasSectors: true,
        });
      }
    } else {
      orderedColumns.push({
        macroareaName: ma.name,
        macroareaId: ma.id,
        sectorName: null,
        sectorId: null,
        hasSectors: false,
      });
    }
  }

  const dataCols = orderedColumns.length;
  const sortedClubs = Array.from(clubData.values()).sort((a, b) => a.name.localeCompare(b.name));
  const clubNameToId = new Map<string, string>();
  for (const [clubId, info] of clubData.entries()) {
    clubNameToId.set(info.name, clubId);
  }

  const dataRowEnd = DATA_ROW_START + sortedClubs.length - 1;
  const assignedRow = dataRowEnd + 2;
  const availableRow = assignedRow + 2;

  const totalRows = HEADER_ROWS + sortedClubs.length + EMPTY_SEPARATOR * 2 + 3;
  const totalCols = FIXED_COLS + dataCols;
  console.log(`Layout: ${totalRows} rows x ${totalCols} cols`);

  const outputPath = path.join(options.outputDir ?? process.cwd(), "output", "assignments.xlsx");

  return new WorkbookBuilder({ author: "Quadro", useSharedStrings: true })
    .addSheet(
      {
        name: "Riepilogo Assegnazioni",
        pageSetup: { paperSize: 9, orientation: "landscape", fitToPage: true, fitToWidth: 1 },
      },
      (sheet) => {
        sheet.setCell("A1", F.sum(colRange(0, DATA_ROW_START + 1, dataRowEnd + 1)), darkHeader);
        sheet.merge({ range: "A1:A2", style: darkHeader });

        sheet.setCell("B1", F.sum(colRange(1, DATA_ROW_START + 1, dataRowEnd + 1)), darkHeader);
        sheet.merge({ range: "B1:B2", style: darkHeader });

        sheet.setCell("C1", "TOTALE PER CLUB", darkHeader);
        sheet.merge({ range: "C1:C3", style: darkHeader });

        let currentCol = FIXED_COLS;
        for (const ma of macroareasWithSectors) {
          const cl1 = colLetter(currentCol);
          if (ma.sectors.length > 0) {
            sheet.setCell(`${cl1}1`, ma.name, darkHeader);
            if (ma.sectors.length > 1) {
              sheet.merge({
                range: `${cl1}1:${colLetter(currentCol + ma.sectors.length - 1)}1`,
                value: ma.name,
                style: darkHeader,
              });
            }
            for (let si = 0; si < ma.sectors.length; si++) {
              sheet.setCell(`${colLetter(currentCol + si)}2`, ma.sectors[si].name, subHeader);
            }
            currentCol += ma.sectors.length;
          } else {
            sheet.setCell(`${cl1}1`, ma.name, darkHeader);
            sheet.merge({ range: `${cl1}1:${cl1}2`, value: ma.name, style: darkHeader });
            currentCol++;
          }
        }

        sheet.setCell("A3", "Richieste", darkHeader);
        sheet.setCell("B3", "Distribuiti", darkHeader);

        for (let i = 0; i < dataCols; i++) {
          sheet.setCell(
            `${colLetter(FIXED_COLS + i)}3`,
            diffFormula(colLetter(FIXED_COLS + i), assignedRow + 1, availableRow + 1),
            darkHeader,
          );
        }

        for (let i = 0; i < sortedClubs.length; i++) {
          const club = sortedClubs[i];
          if (!club) continue;
          const excelRow = DATA_ROW_START + i + 1;
          const clubId = clubNameToId.get(club.name);
          const fullClub = clubId ? clubsById.get(clubId) : undefined;

          sheet.setCell(`A${excelRow}`, club.totalRequests > 0 ? club.totalRequests : "", dataCell);

          const bEndCol = colLetter(FIXED_COLS + dataCols - 1);
          sheet.setCell(
            `B${excelRow}`,
            F.sum(`${colLetter(FIXED_COLS)}${excelRow}:${bEndCol}${excelRow}`),
            dataCell,
          );
          sheet.setCell(`C${excelRow}`, club.name, clubNameCell);

          for (let ci = 0; ci < orderedColumns.length; ci++) {
            const col = orderedColumns[ci];
            if (!col) continue;
            const cl = colLetter(FIXED_COLS + ci);

            if (col.hasSectors && col.sectorId) {
              const sectorImaSeats =
                fullClub?.seats.filter(
                  (s: { originalClubId: string; sectorId?: string }) =>
                    s.originalClubId === clubId && s.sectorId === col.sectorId,
                ) ?? [];
              const requests =
                mockSeatRequests
                  .filter((r) => r.clubId === clubId && r.sectorId === col.sectorId)
                  .reduce((sum, r) => sum + r.quantity, 0) - sectorImaSeats.length;
              sheet.setCell(`${cl}${excelRow}`, requests > 0 ? requests : "", dataCell);
            } else {
              const macroAreaImaSeats =
                fullClub?.seats.filter(
                  (s: { originalClubId: string; sector?: { macroAreaId: string } }) =>
                    s.originalClubId === clubId && s.sector?.macroAreaId === col.macroareaId,
                ) ?? [];
              const requests =
                mockSeatRequests
                  .filter(
                    (r) => r.clubId === clubId && r.macroAreaId === col.macroareaId && !r.sectorId,
                  )
                  .reduce((sum, r) => sum + r.quantity, 0) - macroAreaImaSeats.length;
              sheet.setCell(`${cl}${excelRow}`, requests > 0 ? requests : "", dataCell);
            }
          }
        }

        const assignedExcelRow = assignedRow + 1;
        sheet.setCell(`C${assignedExcelRow}`, "Assegnati:", footerLabel);
        for (let i = 0; i < dataCols; i++) {
          sheet.setCell(
            `${colLetter(FIXED_COLS + i)}${assignedExcelRow}`,
            F.sum(colRange(FIXED_COLS + i, DATA_ROW_START + 1, dataRowEnd + 1)),
            footerCell,
          );
        }

        const availExcelRow = availableRow + 1;
        sheet.setCell(`C${availExcelRow}`, "Disponibili:", footerLabel);
        for (let i = 0; i < dataCols; i++) {
          sheet.setCell(
            `${colLetter(FIXED_COLS + i)}${availExcelRow}`,
            mockResultFreeSeats[i]?.freeSeats ?? 0,
            footerCell,
          );
        }

        sheet.colWidth("A", 14);
        sheet.colWidth("B", 14);
        sheet.colWidth("C", 28);
        for (let i = 0; i < dataCols; i++) {
          sheet.colWidth(colLetter(FIXED_COLS + i), 16);
        }
      },
    )
    .write(outputPath);
}

export const assignments = defineCommand({
  meta: {
    name: "assignments",
    description: "Generate a seat assignment summary report (assignments.xlsx)",
  },
  run() {
    handler()
      .then(({ filePath, sizeBytes }) => {
        console.log(`Workbook written to: ${filePath}`);
        console.log(`File size: ${(sizeBytes / 1024).toFixed(1)} KB`);
      })
      .catch((err) => {
        console.error("Failed:", err);
        process.exit(1);
      });
  },
});
