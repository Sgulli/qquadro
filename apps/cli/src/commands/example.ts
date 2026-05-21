import path from "node:path";
import type { CellStyle, ColumnDef, WriteResult } from "@qquadro/core";
import { F, Styles, WorkbookBuilder } from "@qquadro/core";
import { defineCommand } from "citty";

const sectionTitle: CellStyle = {
  font: { bold: true, size: 13, color: { argb: "FF1F497D" }, name: "Arial" },
  fill: { type: "solid", color: "FFD9E1F2" },
  alignment: { horizontal: "center", vertical: "middle" },
  border: {
    bottom: { style: "medium", color: "FF1F497D" },
    top: { style: "medium", color: "FF1F497D" },
  },
};

const altRow: CellStyle = {
  fill: { type: "solid", color: "FFF2F2F2" },
};

const highlight: CellStyle = {
  font: { bold: true, color: { argb: "FF9C0006" } },
  fill: { type: "solid", color: "FFFFC7CE" },
};

const salesColumns: ColumnDef[] = [
  { key: "region", header: "Region", width: 20, headerStyle: Styles.header },
  { key: "product", header: "Product", width: 24, headerStyle: Styles.header },
  {
    key: "q1",
    header: "Q1 Revenue",
    width: 16,
    style: Styles.currency,
    headerStyle: Styles.header,
  },
  {
    key: "q2",
    header: "Q2 Revenue",
    width: 16,
    style: Styles.currency,
    headerStyle: Styles.header,
  },
  {
    key: "q3",
    header: "Q3 Revenue",
    width: 16,
    style: Styles.currency,
    headerStyle: Styles.header,
  },
  {
    key: "q4",
    header: "Q4 Revenue",
    width: 16,
    style: Styles.currency,
    headerStyle: Styles.header,
  },
  {
    key: "total",
    header: "Annual Total",
    width: 18,
    style: { ...Styles.currency, font: { bold: true } },
    headerStyle: Styles.header,
  },
  {
    key: "growth",
    header: "QoQ Growth",
    width: 14,
    style: Styles.percent,
    headerStyle: Styles.header,
  },
];

const salesData = [
  {
    region: "EMEA",
    product: "Enterprise Suite",
    q1: 120_000,
    q2: 135_000,
    q3: 142_000,
    q4: 158_000,
  },
  { region: "EMEA", product: "SMB Package", q1: 48_000, q2: 51_000, q3: 49_500, q4: 53_000 },
  {
    region: "APAC",
    product: "Enterprise Suite",
    q1: 98_000,
    q2: 104_000,
    q3: 117_500,
    q4: 122_000,
  },
  { region: "APAC", product: "SMB Package", q1: 32_000, q2: 35_500, q3: 37_000, q4: 39_200 },
  {
    region: "AMER",
    product: "Enterprise Suite",
    q1: 215_000,
    q2: 228_000,
    q3: 241_500,
    q4: 260_000,
  },
  { region: "AMER", product: "SMB Package", q1: 72_000, q2: 78_000, q3: 81_000, q4: 88_500 },
];

const kpiColumns: ColumnDef[] = [
  { key: "metric", header: "KPI Metric", width: 28, headerStyle: Styles.header },
  {
    key: "target",
    header: "Target",
    width: 16,
    style: Styles.currency,
    headerStyle: Styles.header,
  },
  {
    key: "actual",
    header: "Actual",
    width: 16,
    style: Styles.currency,
    headerStyle: Styles.header,
  },
  {
    key: "variance",
    header: "Variance",
    width: 16,
    style: Styles.currency,
    headerStyle: Styles.header,
  },
  {
    key: "pct",
    header: "Achievement",
    width: 16,
    style: Styles.percent,
    headerStyle: Styles.header,
  },
  { key: "status", header: "Status", width: 12, headerStyle: Styles.header },
];

const formulaColumns: ColumnDef[] = [
  { key: "label", header: "Description", width: 32, headerStyle: Styles.header },
  { key: "value", header: "Value", width: 20, style: Styles.currency, headerStyle: Styles.header },
  { key: "note", header: "Note", width: 40, headerStyle: Styles.header },
];

export interface ExampleOptions {
  outputDir?: string;
}

export async function handler(options: ExampleOptions = {}): Promise<WriteResult> {
  const outputPath = path.join(options.outputDir ?? process.cwd(), "output", "demo-report.xlsx");

  return new WorkbookBuilder({
    author: "Quadro CLI",
    company: "Quadro Demo",
    useSharedStrings: true,
  })
    .addSheet(
      {
        name: "Sales Report",
        tabColor: "FF2B579A",
        freeze: { row: 2 },
        pageSetup: {
          paperSize: 9,
          orientation: "landscape",
          fitToPage: true,
          fitToWidth: 1,
        },
        headerFooter: {
          oddHeader: { left: "Quadro — Demo Report", right: "&D" },
          oddFooter: { center: "Page &P of &N" },
        },
      },
      (sheet) => {
        sheet.merge({
          range: "A1:H1",
          value: "2024 Annual Sales Report",
          style: sectionTitle,
          height: 32,
        });
        sheet.headers(salesColumns, undefined, 22);

        salesData.forEach((row, i) => {
          const r = i + 3;
          sheet.addRow(
            {
              ...row,
              total: F.add(F.ref("C", r), F.ref("D", r), F.ref("E", r), F.ref("F", r)),
              growth: F.pct(F.ref("F", r), F.ref("C", r)),
            },
            { style: i % 2 === 1 ? altRow : undefined },
          );
        });

        const lastDataRow = salesData.length + 2;
        const totalsRow = lastDataRow + 1;
        sheet.addRow(
          {
            region: "GRAND TOTAL",
            product: "",
            q1: F.sum(F.range("C", 3, lastDataRow)),
            q2: F.sum(F.range("D", 3, lastDataRow)),
            q3: F.sum(F.range("E", 3, lastDataRow)),
            q4: F.sum(F.range("F", 3, lastDataRow)),
            total: F.sum(F.range("G", 3, lastDataRow)),
            growth: F.pct(F.ref("F", totalsRow), F.ref("C", totalsRow)),
          },
          { style: Styles.totalRow, height: 20 },
        );
        sheet.autoFilter("A2:H2");
      },
    )
    .addSheet(
      {
        name: "KPI Dashboard",
        tabColor: "FF70AD47",
        freeze: { row: 3 },
      },
      (sheet) => {
        sheet.merge({
          range: "A1:F1",
          value: "Q4 2024 — KPI Dashboard",
          style: sectionTitle,
          height: 36,
        });
        sheet.merge({
          range: "A2:F2",
          value: "Revenue & Pipeline Metrics",
          style: Styles.subHeader,
          height: 20,
        });
        sheet.headers(kpiColumns);

        const kpiRows = [
          { metric: "New ARR", target: 4_800_000, actual: 5_120_000 },
          { metric: "Expansion ARR", target: 1_200_000, actual: 1_098_000 },
          { metric: "Gross Revenue", target: 6_000_000, actual: 6_218_000 },
          { metric: "Net Revenue", target: 5_400_000, actual: 5_580_000 },
          { metric: "Churn (negative)", target: -240_000, actual: -198_000 },
          { metric: "Pipeline Generated", target: 18_000_000, actual: 17_200_000 },
        ];

        kpiRows.forEach((kpi, i) => {
          const r = i + 4;
          sheet.addRow({
            metric: kpi.metric,
            target: kpi.target,
            actual: kpi.actual,
            variance: F.sub(F.ref("C", r), F.ref("B", r)),
            pct: F.div(F.ref("C", r), F.ref("B", r)),
            status: F.if(`${F.ref("C", r)}/${F.ref("B", r)}>=1`, "On Track", "At Risk"),
          });
        });

        sheet.styleRange("D5:D9", { ...Styles.boxBorder });
        const lastKpi = kpiRows.length + 3;
        sheet.addRow(
          {
            metric: "Total Revenue",
            target: F.sum(F.range("B", 4, lastKpi)),
            actual: F.sum(F.range("C", 4, lastKpi)),
            variance: F.sub(F.ref("C", lastKpi + 1), F.ref("B", lastKpi + 1)),
            pct: F.div(F.ref("C", lastKpi + 1), F.ref("B", lastKpi + 1)),
            status: "",
          },
          { style: Styles.totalRow },
        );
        sheet.autoFitColumns();
      },
    )
    .addSheet(
      {
        name: "Formula Reference",
        tabColor: "FFED7D31",
      },
      (sheet) => {
        sheet.merge({
          range: "A1:C1",
          value: "Excel Formula Examples",
          style: sectionTitle,
          height: 30,
        });
        sheet.headers(formulaColumns);

        sheet.setCell("B3", 50_000, Styles.inputCell);
        sheet.setCell("B4", 75_000, Styles.inputCell);
        sheet.setCell("B5", 92_000, Styles.inputCell);

        sheet.addRow(
          { label: "Base Value A", value: 50_000, note: "Hardcoded input" },
          { style: { font: { color: { argb: "FF0000FF" } } } },
        );
        sheet.addRow(
          { label: "Base Value B", value: 75_000, note: "Hardcoded input" },
          { style: { font: { color: { argb: "FF0000FF" } } } },
        );
        sheet.addRow(
          { label: "Base Value C", value: 92_000, note: "Hardcoded input" },
          { style: { font: { color: { argb: "FF0000FF" } } } },
        );
        sheet.addRow({ label: "SUM(A+B+C)", value: F.sum("B3:B5", 217_000), note: "=SUM(B3:B5)" });
        sheet.addRow({
          label: "AVERAGE",
          value: F.average("B3:B5", 72_333),
          note: "=AVERAGE(B3:B5)",
        });
        sheet.addRow({ label: "MAX", value: F.max("B3:B5", 92_000), note: "=MAX(B3:B5)" });
        sheet.addRow({ label: "MIN", value: F.min("B3:B5", 50_000), note: "=MIN(B3:B5)" });
        sheet.addRow({ label: "COUNT", value: F.count("B3:B5", 3), note: "=COUNT(B3:B5)" });
        sheet.addRow(
          { label: "YoY Growth %", value: F.pct("B5", "B3"), note: "=(B5-B3)/B3" },
          { style: Styles.percent },
        );
        sheet.addRow({
          label: "Nested IF",
          value: F.if("B6>200000", "Exceeds Budget", "Within Budget"),
          note: "Conditional text",
        });

        sheet.merge({ range: "A13:C13", value: "Merge & Style Examples", style: Styles.subHeader });
        sheet.merge({
          range: "A14:B14",
          value: "This cell spans two columns",
          style: { ...Styles.boxBorder, alignment: { horizontal: "center" } },
        });
        sheet.setCell("C14", "merged A14:B14", {
          font: { italic: true, color: { argb: "FF595959" } },
        });
        sheet.merge({ range: "A15:C15", value: "Full-width highlight", style: highlight });
        sheet.autoFitColumns();
      },
    )
    .write(outputPath);
}

export const example = defineCommand({
  meta: {
    name: "example",
    description: "Generate a demo multi-sheet workbook (demo-report.xlsx)",
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
