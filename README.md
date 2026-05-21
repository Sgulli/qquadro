# quadro

Fluent, fully-typed Excel workbook builder for Node.js. Built on [excelts](https://github.com/cjnoname/excelts) by [@cjnoname](https://github.com/cjnoname). Created by [@sgulli](https://github.com/Sgulli).

```ts
import { WorkbookBuilder, Styles, F } from "@qquadro/core";

await new WorkbookBuilder({ author: "Acme Corp" })
  .addSheet({ name: "Sales" }, (sheet) => {
    sheet
      .headers([
        { key: "product", header: "Product", width: 22, headerStyle: Styles.header },
        { key: "revenue", header: "Revenue ($)", width: 18, style: Styles.currency },
      ])
      .addRows([
        { product: "Widget A", revenue: 120_000 },
        { product: "Widget B", revenue: 98_500 },
      ])
      .addRow(
        { product: "Total", revenue: F.sum("B2:B3") },
        { style: Styles.totalRow },
      )
      .autoFilter();
  })
  .write("./report.xlsx");
```

## Install

```sh
npm install @qquadro/core
```

```sh
pnpm add @qquadro/core
```

```sh
yarn add @qquadro/core
```

```sh
bun add @qquadro/core
```

Requires Node 18+.

## Features

- **Fluent API** — chainable `.headers()`, `.addRow()`, `.merge()`, `.freeze()`, etc.
- **Numeric RC API** — reference cells by `(col, row)` numbers, no A1 strings
- **Data validation** — dropdown lists, number/date ranges, custom formulas
- **Conditional formatting** — cell rules, data bars, color scales, icon sets, top N
- **Named ranges** — workbook-level defined names with RC support
- **Tables** — structured references, totals row, banded rows, table styles
- **Workbook reading** — load, modify, re-save existing XLSX files
- **CSV import/export** — read/write CSV to files or strings
- **Sheet-to-JSON** — convert worksheet data to objects or arrays
- **Formulas** — typed helpers (`F.sum()`, `F.pct()`, `F.add()`, etc.) or raw `{ formula: "..." }`
- **Styles** — presets (`Styles.header`, `Styles.currency`, `Styles.totalRow`) or composable custom
- **Merged cells** — `.merge()` with value, style & height in one call
- **Freeze panes** + **auto-filter**
- **Streaming** — constant memory for large files via `useStreaming` option
- **ESM + CJS** — dual package with exports map
- **TypeScript** — full type declarations

## Quick Tour

### Creating Workbooks

```ts
const wb = new WorkbookBuilder({ author: "Acme Corp" });

wb.addSheet({ name: "Sales", freeze: { row: 2 } }, (sheet) => {
  sheet
    .headers([
      { key: "region", header: "Region", width: 18, headerStyle: Styles.header },
      { key: "q1", header: "Q1", width: 14, style: Styles.currency },
      { key: "q2", header: "Q2", width: 14, style: Styles.currency },
    ])
    .addRows([
      { region: "EMEA", q1: 120_000, q2: 135_000 },
      { region: "APAC", q1: 98_000, q2: 104_000 },
    ])
    .addRow(
      { region: "Total", q1: F.sum("B2:B3"), q2: F.sum("C2:C3") },
      { style: Styles.totalRow },
    )
    .autoFilter("A1:C1");
});

const { filePath, sizeBytes } = await wb.write("./output/report.xlsx");
```

### Numeric RC API — No A1 strings

Use `for` loops with column/row numbers:

```ts
const COL = { name: 1, score: 2, status: 3 };
const dataStart = 2;

sheet
  .headers([
    { key: "name", header: "Name" },
    { key: "score", header: "Score" },
    { key: "status", header: "Status" },
  ])
  .addRows([
    { name: "Alice", score: 85, status: "Pass" },
    { name: "Bob", score: 42, status: "Fail" },
  ]);

const dr = { start: dataStart, end: dataStart + sheet.rowCount - 1 };
sheet
  .addDataBarRC(COL.score, dr.start, COL.score, dr.end, { argb: "FF5B9BD5" })
  .addCellIsRuleRC(COL.score, dr.start, COL.score, dr.end, "greaterThanOrEqual", [70], {
    font: { bold: true, color: { argb: "FF006100" } },
  });
```

### Named Ranges

```ts
wb.defineName("MyRange", "A1:B10", "Sheet1");
wb.defineNameRC("Data", 1, 1, 5, 10, "Sheet1");
wb.defineFormula("Double", "LAMBDA(x,x*2)");
wb.getDefinedNames(); // → DefinedNameModel[]
```

### Tables

```ts
sheet.addTable("SalesTable", "A1:D20", [
  { name: "Product" },
  { name: "Revenue", totalsRowFunction: "sum" },
  { name: "Region" },
], {
  totalsRow: true,
  style: { theme: "TableStyleMedium9", showRowStripes: true },
});
```

### Data Validation

```ts
sheet
  .addListValidation(sheet.columnRange("product"), ["Widget A", "Widget B"])
  .addRangeValidation(sheet.columnRange("quantity"), "whole", "between", [1, 1000], {
    error: "Must be between 1 and 1000",
  });
```

### Conditional Formatting

```ts
sheet
  .addCellIsRule(sheet.columnRange("score"), "greaterThanOrEqual", [70], {
    font: { color: { argb: "FF006100" } },
    fill: { type: "solid", color: { argb: "FFC6EFCE" } },
  })
  .addDataBar(sheet.columnRange("score"), { argb: "FF5B9BD5" })
  .addIconSet(sheet.columnRange("score"), "3TrafficLights1");
```

### Reading Workbooks

```ts
const wb = await WorkbookBuilder.load("./input.xlsx");

const data = wb.sheets[0].toJSON();
// => [{ Name: "Alice", Age: 30 }, { Name: "Bob", Age: 25 }]

const rows = wb.sheets[0].toJSON({ header: 1 });
// => [["Name", "Age"], ["Alice", 30], ["Bob", 25]]

const aoa = wb.sheets[0].toAOA();
// => [["Name", "Age"], ["Alice", 30], ["Bob", 25]]
```

### Template Workflow

```ts
const wb = await WorkbookBuilder.load("./template.xlsx");
// modify existing sheets or add new ones
await wb.write("./output.xlsx");
```

### CSV Import/Export

```ts
// Write CSV
const csvString = await wb.toCsv();
await wb.toCsv("./output.csv");

// Read CSV from inline data
const csvWb = await WorkbookBuilder.fromCsv("name,age\nAlice,30\nBob,25");

// Read CSV from file
const csvFileWb = await WorkbookBuilder.fromCsvFile("./data.csv");
```

### Streaming (large files)

```ts
const wb = new WorkbookBuilder({ useStreaming: true });
// add rows — memory usage stays constant
```

### Adding Data from JSON

```ts
wb.sheets[0].addJSON([
  { Name: "Alice", Age: 30 },
  { Name: "Bob", Age: 25 },
]);

wb.sheets[0].addAOA([
  ["Name", "Age"],
  ["Alice", 30],
  ["Bob", 25],
]);
```

## Formula Helpers

| Helper | Output | Example |
|--------|--------|---------|
| `F.sum(range)` | `SUM(C3:C8)` | aggregate |
| `F.average(range)` | `AVERAGE(C3:C8)` | |
| `F.count(range)` | `COUNT(C3:C8)` | |
| `F.max(range)` | `MAX(C3:C8)` | |
| `F.min(range)` | `MIN(C3:C8)` | |
| `F.pct(current, prev)` | `(C3-D3)/D3` | growth % |
| `F.add(a, b, ...)` | `C3+D3+E3` | arithmetic |
| `F.sub(a, b)` | `C3-D3` | |
| `F.mul(a, b)` | `C3*D3` | |
| `F.div(a, b)` | `C3/D3` | |
| `F.if(cond, t, f)` | `IF(C3>0,"ok","bad")` | conditional |

Raw formulas also work: `{ formula: "SUM(C3:C8)", result: 217000 }`.

## Styles

### Presets

| Preset | Use |
|--------|-----|
| `Styles.header` | Column headers — bold white on blue, center aligned |
| `Styles.subHeader` | Sub-section headers — blue on light blue |
| `Styles.totalRow` | Totals — bold on yellow with double bottom border |
| `Styles.currency` | Right-aligned number format |
| `Styles.percent` | Right-aligned `%` format (2 decimals) |
| `Styles.date` | Centered date format |
| `Styles.boxBorder` | Thin grey border all sides |
| `Styles.inputCell` | Blue font — hardcoded input convention |
| `Styles.formulaCell` | Black font — formula convention |
| `Styles.linkCell` | Green font — cross-sheet link convention |

### Composable styles

```ts
import { style, font, fill, numFmt, border, align, currency } from "@qquadro/core";

const euroCell = style(
  font({ bold: true, size: 10, name: "Arial" }),
  fill({ type: "solid", color: "FFFFF2CC" }),
  border.thinBlack,
  align.right,
  numFmt(currency("€")),
);
```

| Part | Description |
|------|-------------|
| `style(...parts)` | Deep-merge multiple partials into a `CellStyle` |
| `font({...})` | `{ font: Font }` — bold, size, color, name, etc. |
| `fill({...})` | `{ fill: Fill }` — solid or gradient |
| `numFmt(fmt)` | `{ numberFormat: NumberFormat }` — e.g. `currency("€")` |
| `border.thinBlack` | Thin black border on all 4 sides |
| `border.thinGrey` | Thin grey border on all 4 sides |
| `border.thin(color)` | Thin border with custom ARGB color |
| `border.all(style, color)` | Any border style/color on all sides |
| `align.center` | Horizontal center, vertical middle |
| `align.centerWrap` | Center + `wrapText` |
| `align.left` / `align.right` | Left/right + middle |
| `align.leftWrap` / `align.rightWrap` | Left/right + `wrapText` |

### Number format helpers

`currency("€")` returns `"€"#,##0.00` — works with any symbol.

```ts
style({ numberFormat: currency("$") })      // "$"#,##0.00
style({ numberFormat: currency("€") })      // "€"#,##0.00
style({ numberFormat: accounting("€") })    // "€"#,##0.00;("€"#,##0.00);"-"
```

## License

MIT
