# quadro

Fluent, fully-typed Excel builder for Node.js. Powered by [ExcelTS](https://github.com/cjnoname/excelts).

```ts
import { WorkbookBuilder, Styles, F } from "quadro";

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

```
npm install quadro
```

Requires Node 18+.

## Features

- **Fluent API** — chain `.columns()`, `.addRow()`, `.merge()`, `.freeze()`, etc.
- **Formulas** — typed helpers (`F.sum()`, `F.pct()`, `F.add()`, etc.) or raw `{ formula: "..." }`
- **Styles** — presets (`Styles.header`, `Styles.currency`, `Styles.totalRow`) or custom
- **Merged cells** — `.merge()` with value & style in one call
- **Freeze panes** + **auto-filter**
- **Streaming** — constant memory for large files via `useStreaming` option
- **ESM + CJS** — dual package with exports map
- **TypeScript** — full type declarations

## Quick Tour

```ts
const wb = new WorkbookBuilder({ author: "Acme Corp", useSharedStrings: true });
```

**Sheet with columns & formulas:**

```ts
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
```

**Write to disk:**

```ts
const { filePath, sizeBytes } = await wb.write("./output/report.xlsx");
```

### Streaming (large files)

```ts
const wb = new WorkbookBuilder({ useStreaming: true });
// add rows — memory usage stays constant
```

## Formula Helpers

| Helper | Output | Example |
|--------|--------|---------|
| `F.ref(col, row)` | `"C3"` | cell reference |
| `F.range(col, from, to)` | `"C3:C8"` | column range |
| `F.sum(range)` | `SUM(C3:C8)` | aggregate |
| `F.average(range)` | `AVERAGE(C3:C8)` | |
| `F.pct(current, prev)` | `(C3-D3)/D3` | growth % |
| `F.add(a, b, ...)` | `C3+D3+E3` | arithmetic |
| `F.sub(a, b)` | `C3-D3` | |
| `F.div(a, b)` | `C3/D3` | |
| `F.if(cond, t, f)` | `IF(C3>0,"ok","bad")` | conditional |

Raw formulas also work: `{ formula: "SUM(C3:C8)", result: 217000 }`.

## Styles

Presets: `Styles.header`, `Styles.currency`, `Styles.percent`, `Styles.totalRow`, `Styles.inputCell`, `Styles.subHeader`, `Styles.boxBorder`.

Custom: use `CellStyle` interface (font, fill, border, alignment, numberFormat).

## License

MIT
