# quadro

Fluent, fully-typed Excel builder for Node.js. Built on [ExcelTS](https://github.com/cjnoname/excelts).

```ts
import { WorkbookBuilder, Styles, F } from "@quadro/core";

await new WorkbookBuilder({ author: "Acme Corp" })
  .addSheet({ name: "Sales" }, (sheet) => {
    sheet
      .columns([
        { key: "product", header: "Product", width: 22, headerStyle: Styles.header },
        { key: "revenue", header: "Revenue ($)", width: 18, style: Styles.currency },
        { key: "growth", header: "Growth", width: 14, style: Styles.percent },
      ])
      .writeHeaders()
      .addRows([
        { product: "Widget A", revenue: 36_000, growth: F.pct("B3", "B2") },
        { product: "Widget B", revenue: 25_500, growth: F.pct("B4", "B3") },
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
npm install @quadro/core
```

Requires Node 18+.

## Features

- **Fluent API** — chain `.columns()`, `.addRow()`, `.merge()`, `.freeze()`, etc.
- **Workbook reading** — load, modify, re-save existing XLSX files
- **CSV import/export** — read/write CSV to files or strings
- **Sheet-to-JSON** — convert worksheet data to objects or arrays
- **Formulas** — typed helpers (`F.sum()`, `F.pct()`, `F.add()`, etc.) or raw `{ formula: "..." }`
- **Styles** — presets (`Styles.header`, `Styles.currency`, `Styles.totalRow`) or custom
- **Merged cells** — `.merge()` with value & style in one call
- **Freeze panes** + **auto-filter**
- **Streaming** — constant memory for large files via `useStreaming` option
- **ESM + CJS** — dual package with exports map
- **TypeScript** — full type declarations

## Quick Tour

### Creating Workbooks

```ts
const wb = new WorkbookBuilder({ author: "Acme Corp", useSharedStrings: true });

wb.addSheet({ name: "Sales", freeze: { row: 3 } }, (sheet) => {
  sheet.merge({ range: "A1:F1", value: "Report", style: Styles.header });
  sheet.columns([
    { key: "region", header: "Region", width: 18, headerStyle: Styles.header },
    { key: "q1", header: "Q1", width: 14, style: Styles.currency },
    { key: "q2", header: "Q2", width: 14, style: Styles.currency },
    { key: "total", header: "Total", width: 16, style: { ...Styles.currency, font: { bold: true } } },
  ]).writeHeaders();

  const data = [
    { region: "EMEA", q1: 120_000, q2: 135_000 },
    { region: "APAC", q1: 98_000, q2: 104_000 },
  ];
  data.forEach((row, i) => {
    const r = i + 3;
    sheet.addRow({ ...row, total: F.add(F.ref("C", r), F.ref("D", r)) });
  });
  sheet.addRow(
    { region: "Total", q1: F.sum(F.range("C", 3, data.length + 2)), q2: F.sum(F.range("D", 3, data.length + 2)) },
    { style: Styles.totalRow },
  );
  sheet.autoFilter("A2:D2");
});

const { filePath, sizeBytes } = await wb.write("./output/report.xlsx");
```

### Reading Workbooks

```ts
import { WorkbookBuilder } from "@quadro/core";

const wb = await WorkbookBuilder.load("./input.xlsx");

// Read data as objects (first row = headers)
const data = wb.sheets[0].toJSON();
// => [{ Name: "Alice", Age: 30 }, { Name: "Bob", Age: 25 }]

// Read as array of arrays
const rows = wb.sheets[0].toJSON({ header: 1 });
// => [["Name", "Age"], ["Alice", 30], ["Bob", 25]]

// Read as array of arrays with toAOA
const aoa = wb.sheets[0].toAOA();
// => [["Name", "Age"], ["Alice", 30], ["Bob", 25]]
```

### Template Workflow

Load an existing file, modify it, and save a new version:

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

// Read CSV
const csvWb = await WorkbookBuilder.fromCsv("name,age\nAlice,30\nBob,25", {
  fileName: "data.csv",
});
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
| `F.ref(col, row)` | `"C3"` | cell reference |
| `F.range(col, from, to)` | `"C3:C8"` | column range |
| `F.rect(col1, row1, col2, row2)` | `"A1:D10"` | rectangular range |
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

Build styles from partials with `style()` — deep-merges font, fill, border, alignment, and numberFormat so you can mix reusable pieces.

```ts
import { style, font, fill, numFmt, border, align, currency } from "@quadro/core";

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
