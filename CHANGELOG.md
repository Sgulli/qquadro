# Changelog

## 0.3.0 (unreleased)

- Data validation: `sheet.addDataValidation()`, `sheet.addListValidation()`, `sheet.addRangeValidation()`
- Conditional formatting: `sheet.addConditionalFormatting()`, `sheet.addCellIsRule()`, `sheet.addExpressionRule()`
- Conditional formatting: `sheet.addDataBar()`, `sheet.addColorScale()`, `sheet.addIconSet()`
- Conditional formatting: `sheet.addTop10Rule()`, `sheet.addAboveAverageRule()`, `sheet.addContainsTextRule()`, `sheet.addTimePeriodRule()`
- Exported `DataValidation`, `ConditionalFormattingOptions`, and related types from `@quadro/core`

## 0.2.0 (unreleased)

- Workbook reading: `WorkbookBuilder.load()`, `WorkbookBuilder.fromFile()`
- Sheet-to-JSON: `sheet.toJSON()`, `sheet.toAOA()`, `sheet.addJSON()`, `sheet.addAOA()`
- CSV import/export: `workbook.toCsv()`, `WorkbookBuilder.fromCsv()`, `fromCsvFile()`
- Template workflow: load → modify → re-save
- Command injection fix in debug CLI command (`execFileSync`)
- Path containment via `allowedBase` option
- CLI now has lint + test scripts and handler tests

## 0.1.0 (unreleased)

- Initial release of `@quadro/core`
- Workbook builder with fluent API
- Sheet builder with columns, headers, rows, cells, merged cells
- Full style system (font, fill, border, alignment, number formats)
- Preset styles (header, subHeader, totalRow, currency, percent, date, etc.)
- Formula helpers (SUM, AVERAGE, COUNT, MAX, MIN, ADD, SUB, MUL, DIV, PCT, IF)
- Freeze panes, auto-filter
- Page setup, headers/footers
- Streaming mode for large datasets
- Dual ESM/CJS builds
- CLI app (`@quadro/cli`) with `example`, `assignments`, `debug` commands
