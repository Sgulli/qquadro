# Changelog

## 0.3.0 (2026-05-20)

- Data validation: `sheet.addDataValidation()`, `sheet.addListValidation()`, `sheet.addRangeValidation()`
- Conditional formatting: `sheet.addConditionalFormatting()`, `sheet.addCellIsRule()`, `sheet.addExpressionRule()`
- Conditional formatting: `sheet.addDataBar()`, `sheet.addColorScale()`, `sheet.addIconSet()`
- Conditional formatting: `sheet.addTop10Rule()`, `sheet.addAboveAverageRule()`, `sheet.addContainsTextRule()`, `sheet.addTimePeriodRule()`
- Exported `DataValidation`, `ConditionalFormattingOptions`, and related types from `@quadro/core`
- Simplified API: `headers()` replaces `columns()` + `writeHeaders()` in one call
- `merge()` accepts optional `height` property (no separate `rowHeight()` call needed)
- `headers()` accepts optional `height` parameter
- Columns auto-inferred from first object row when not explicitly defined
- `columnRange()` auto-detects header offset (row 1 vs row 2)
- CLI commands exposed via `pnpm cli` and `pnpm example`
- Numeric coordinate API: `setCellRC()`, `mergeRC()`, `styleRangeRC()`, `colWidthRC()`
- Numeric validation/formatting helpers: `cellRef()`, `colRange()`, `rangeRef()`
- `addListValidationRC()`, `addRangeValidationRC()`, `addDataValidationRC()`
- `addCellIsRuleRC()`, `addExpressionRuleRC()`, `addDataBarRC()`, `addColorScaleRC()`
- `addIconSetRC()`, `addTop10RuleRC()`, `addAboveAverageRuleRC()`
- `addContainsTextRuleRC()`, `addTimePeriodRuleRC()`
- Full zero-A1-string sheet building with `for` loops
- Conditional formatting removal: `sheet.removeConditionalFormatting()`
- Rewrote assignments example using pure numeric coordinates (no A1 strings)
- Fixed pre-existing merge crash in assignments (off-by-one `FIXED_COLS`)
- Fixed pre-existing TypeScript errors in seat filter type guards
- `sheet.columnIndex()` resolves column key to 1‑based index for RC API
- `sheet.removeDataValidation()` to remove validation rules
- Updated `@quadro/core` README with RC API docs, package managers, author attribution

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
