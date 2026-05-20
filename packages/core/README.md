# @quadro/core

Fluent, fully-typed Excel workbook builder for Node.js. Built on [excelts](https://github.com/cj-tech-master/excelts).

```ts
import { WorkbookBuilder, Styles, F } from "@quadro/core";

await new WorkbookBuilder({ author: "Acme Corp" })
  .addSheet({ name: "Sales" }, (sheet) => {
    sheet
      .columns([
        { key: "product", header: "Product", width: 24, headerStyle: Styles.header },
        { key: "revenue", header: "Revenue ($)", width: 18, style: Styles.currency, headerStyle: Styles.header },
      ])
      .writeHeaders()
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

- **Fluent API** — chainable `addSheet` → `columns` → `writeHeaders` → `addRows` → `write`
- **Full styling** — fonts, fills, borders, alignment, number formats with preset styles
- **Formulas** — typed helpers for SUM, AVERAGE, IF, ADD, PCT, and more
- **Merged cells** — merge ranges with value and style
- **Freeze panes & auto-filter** — header rows stay visible, one-liner filter
- **Headers & footers** — per-sheet odd/even header/footer sections
- **Page setup** — orientation, paper size, fit-to-page, margins
- **Streaming mode** — constant memory for large datasets
- **Dual ESM/CJS** — works with `import` and `require`

## Installation

```bash
npm install @quadro/core
```

## Documentation

See the [monorepo README](https://github.com/Sgulli/quadro) for full API reference, examples, and style system documentation.

## License

MIT
