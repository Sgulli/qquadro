import path from "node:path";
import type { WriteResult } from "@quadro/core";
import {
  align,
  border,
  cellRef,
  colRange,
  F,
  rangeRef,
  style,
  WorkbookBuilder,
} from "@quadro/core";
import { defineCommand } from "citty";
import {
  mockClubs,
  mockMacroareas,
  mockResultFreeSeats,
  mockSeatRequests,
} from "../data/mock-data.js";

const darkHeader = style(
  { font: { bold: true, size: 11, color: { argb: "FFFFFFFF" }, name: "Arial" } },
  { fill: { type: "solid", color: "FF2B579A" } },
  align.centerWrap,
  border.thinBlack,
);

const subHeader = style(
  { font: { bold: true, size: 10, color: { argb: "FF1F497D" }, name: "Arial" } },
  { fill: { type: "solid", color: "FFDCE6F1" } },
  align.centerWrap,
  border.thinBlack,
);

const dataCell = style({ font: { name: "Arial", size: 10 } }, align.center, border.thinBlack);
const clubNameCell = style(dataCell, { font: { bold: true } }, align.left);
const footerLabel = style({ font: { bold: true, size: 10, name: "Arial" } }, align.right);
const footerCell = style(dataCell, { font: { bold: true } });

const FIXED_COLS = 4;
const DATA_ROW_START = 3;

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
      fullClub?.seats.filter((s) => {
        if (s.originalClubId !== clubId) return false;
        if (sectorId) return "sectorId" in s && s.sectorId === sectorId;
        return "sector" in s && s.sector?.macroAreaId === macroAreaId;
      }) ?? [];
    club.totalRequests += Math.max(0, quantity - clubImaSeats.length);
  }

  const orderedColumns: {
    macroareaName: string;
    macroareaId: string;
    sectorName: string | null;
    sectorId: string | null;
    hasSectors: boolean;
  }[] = [];
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

  const outputPath = path.join(options.outputDir ?? process.cwd(), "output", "assignments.xlsx");

  return new WorkbookBuilder({ author: "Quadro", useSharedStrings: true })
    .addSheet(
      {
        name: "Riepilogo Assegnazioni",
        pageSetup: {
          paperSize: 9,
          orientation: "landscape",
          fitToPage: true,
          fitToWidth: 1,
        },
      },
      (sheet) => {
        const dr = DATA_ROW_START;
        const drEnd = dataRowEnd;
        const assignR = assignedRow;
        const availR = availableRow;

        sheet.setCellRC(1, 1, F.sum(colRange(1, dr + 1, drEnd + 1)), darkHeader);
        sheet.mergeRC(1, 1, 1, 2, { style: darkHeader });

        sheet.setCellRC(2, 1, F.sum(colRange(2, dr + 1, drEnd + 1)), darkHeader);
        sheet.mergeRC(2, 1, 2, 2, { style: darkHeader });

        sheet.setCellRC(3, 1, "TOTALE PER CLUB", darkHeader);
        sheet.mergeRC(3, 1, 3, 3, { style: darkHeader });

        let currentCol = FIXED_COLS;
        for (const ma of macroareasWithSectors) {
          if (ma.sectors.length > 0) {
            sheet.setCellRC(currentCol, 1, ma.name, darkHeader);
            if (ma.sectors.length > 1) {
              sheet.mergeRC(currentCol, 1, currentCol + ma.sectors.length - 1, 1, {
                value: ma.name,
                style: darkHeader,
              });
            }
            for (let si = 0; si < ma.sectors.length; si++) {
              sheet.setCellRC(currentCol + si, 2, ma.sectors[si].name, subHeader);
            }
            currentCol += ma.sectors.length;
          } else {
            sheet.setCellRC(currentCol, 1, ma.name, darkHeader);
            sheet.mergeRC(currentCol, 1, currentCol, 2, {
              value: ma.name,
              style: darkHeader,
            });
            currentCol++;
          }
        }

        sheet.setCellRC(1, 3, "Richieste", darkHeader);
        sheet.setCellRC(2, 3, "Distribuiti", darkHeader);

        for (let ci = 0; ci < dataCols; ci++) {
          const c = FIXED_COLS + ci;
          sheet.setCellRC(
            c,
            3,
            { formula: `${cellRef(c, assignR + 1)}-${cellRef(c, availR + 1)}` },
            darkHeader,
          );
        }

        for (let ri = 0; ri < sortedClubs.length; ri++) {
          const club = sortedClubs[ri];
          if (!club) continue;
          const row = dr + ri + 1;
          const clubId = clubNameToId.get(club.name);
          const fullClub = clubId ? clubsById.get(clubId) : undefined;

          sheet.setCellRC(1, row, club.totalRequests > 0 ? club.totalRequests : "", dataCell);
          sheet.setCellRC(
            2,
            row,
            F.sum(rangeRef(FIXED_COLS, row, FIXED_COLS + dataCols - 1, row)),
            dataCell,
          );
          sheet.setCellRC(3, row, club.name, clubNameCell);

          for (let ci = 0; ci < orderedColumns.length; ci++) {
            const col = orderedColumns[ci];
            if (!col) continue;
            const c = FIXED_COLS + ci;

            if (col.hasSectors && col.sectorId) {
              const sectorImaSeats =
                fullClub?.seats.filter(
                  (s) =>
                    s.originalClubId === clubId && "sectorId" in s && s.sectorId === col.sectorId,
                ) ?? [];
              const requests =
                mockSeatRequests
                  .filter((r) => r.clubId === clubId && r.sectorId === col.sectorId)
                  .reduce((sum, r) => sum + r.quantity, 0) - sectorImaSeats.length;
              sheet.setCellRC(c, row, requests > 0 ? requests : "", dataCell);
            } else {
              const macroAreaImaSeats =
                fullClub?.seats.filter(
                  (s) =>
                    s.originalClubId === clubId &&
                    "sector" in s &&
                    s.sector?.macroAreaId === col.macroareaId,
                ) ?? [];
              const requests =
                mockSeatRequests
                  .filter(
                    (r) => r.clubId === clubId && r.macroAreaId === col.macroareaId && !r.sectorId,
                  )
                  .reduce((sum, r) => sum + r.quantity, 0) - macroAreaImaSeats.length;
              sheet.setCellRC(c, row, requests > 0 ? requests : "", dataCell);
            }
          }
        }

        sheet.setCellRC(3, assignR, "Assegnati:", footerLabel);
        for (let ci = 0; ci < dataCols; ci++) {
          const c = FIXED_COLS + ci;
          sheet.setCellRC(c, assignR, F.sum(colRange(c, dr + 1, drEnd + 1)), footerCell);
        }

        sheet.setCellRC(3, availR, "Disponibili:", footerLabel);
        for (let ci = 0; ci < dataCols; ci++) {
          sheet.setCellRC(
            FIXED_COLS + ci,
            availR,
            mockResultFreeSeats[ci]?.freeSeats ?? 0,
            footerCell,
          );
        }

        sheet.colWidth("A", 14);
        sheet.colWidth("B", 14);
        sheet.colWidth("C", 28);
        for (let i = 0; i < dataCols; i++) {
          sheet.colWidth(cellRef(FIXED_COLS + i, 1).replace(/\d+$/, ""), 16);
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
