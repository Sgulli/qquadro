# @quadro/core

Fluent, fully-typed Excel workbook builder for Node.js. Built on [excelts](https://github.com/cj-tech-master/excelts).

```ts
import { WorkbookBuilder, Styles, F } from "@quadro/core";

await new WorkbookBuilder({ author: "Acme Corp" })
  .addSheet({ name: "Sales" }, (sheet) => {
    sheet
      .headers([
        { key: "product", header: "Product", width: 24, headerStyle: Styles.header },
        { key: "revenue", header: "Revenue ($)", width: 18, style: Styles.currency, headerStyle: Styles.header },
      ])
      .addRows([
        { product: "Widget Pro", revenue: 120_000 },
        { product: "Gadget X",   revenue: 98_500 },
      ])
      .addRow(
        { product: "Total", revenue: F.sum("B2:B3") },
        { style: Styles.totalRow },
      )
      .autoFitColumns()
      .freeze(1);
  })
  .write("./report.xlsx");
```

## Features

- **Fluent API** — chainable `addSheet` → `headers` → `addRows` → `write`
- **Full styling** — fonts, fills, borders, alignment, number formats with preset styles
- **Formulas** — typed helpers for SUM, AVERAGE, IF, ADD, PCT, and more
- **Data validation** — dropdown lists, number/date ranges, custom rules
- **Conditional formatting** — cell rules, data bars, color scales, icon sets, top N
- **Merged cells** — merge ranges with value and style
- **Freeze panes & auto-filter** — header rows stay visible, one-liner filter
- **Headers & footers** — per-sheet odd/even header/footer sections
- **Page setup** — orientation, paper size, fit-to-page, margins
- **Streaming mode** — constant memory for large datasets
- **Dual ESM/CJS** — works with `import` and `require`

## Data Validation

Restrict cell input with dropdown lists, number ranges, dates, or custom formulas.

Columns are inferred from your data — no need to define them explicitly:

```ts
import { WorkbookBuilder } from "@quadro/core";

await new WorkbookBuilder()
  .addSheet({ name: "Validation" }, (sheet) => {
    sheet
      .addRows([
        { product: "Widget A", quantity: 100, date: new Date("2026-01-15") },
        { product: "Widget B", quantity: 50, date: new Date("2026-02-01") },
      ])
      // Columns "product", "quantity", "date" are auto-inferred from the first row.
      // sheet.columnRange("quantity") → "B2:B3"
      .addListValidation(sheet.columnRange("product"), ["Widget A", "Widget B", "Gadget X"])
      .addRangeValidation(sheet.columnRange("quantity"), "whole", "between", [1, 1000], {
        error: "Must be between 1 and 1000",
        errorTitle: "Invalid quantity",
      })
      .addRangeValidation(sheet.columnRange("date"), "date", "greaterThanOrEqual", [new Date("2026-01-01")]);
  })
  .write("./validation.xlsx");
```

## Conditional Formatting

Highlight cells, add data bars, color scales, icon sets, and top N rules. Columns are inferred from data, so you can skip `.columns()` and `.writeHeaders()`:

```ts
import { WorkbookBuilder } from "@quadro/core";

await new WorkbookBuilder()
  .addSheet({ name: "Formatting" }, (sheet) => {
    sheet
      .addRows([
        { name: "Alice", score: 85, status: "Pass" },
        { name: "Bob", score: 42, status: "Fail" },
        { name: "Charlie", score: 73, status: "Pass" },
      ])
      .addCellIsRule(sheet.columnRange("score"), "greaterThanOrEqual", [70], {
        font: { bold: true, color: { argb: "FF006100" } },
        fill: { type: "solid", color: { argb: "FFC6EFCE" } },
      })
      .addCellIsRule(sheet.columnRange("score"), "lessThan", [50], {
        font: { color: { argb: "FF9C0006" } },
        fill: { type: "solid", color: { argb: "FFFFC7CE" } },
      })
      .addDataBar(sheet.columnRange("score"), { argb: "FF5B9BD5" })
      .addIconSet(sheet.columnRange("score"), "3TrafficLights1")
      .addContainsTextRule(sheet.columnRange("status"), "Pass", "containsText", {
        font: { color: { argb: "FF006100" } },
      })
      .addContainsTextRule(sheet.columnRange("status"), "Fail", "containsText", {
        font: { color: { argb: "FF9C0006" } },
      });
  })
  .write("./formatting.xlsx");
```

## Installation

```bash
npm install @quadro/core
```

## Documentation

See the [monorepo README](https://github.com/Sgulli/quadro) for full API reference, examples, and style system documentation.

## License

MIT
