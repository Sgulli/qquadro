# Roadmap

## v0.1 — Foundation ✅

- [x] Add test framework (vitest)
- [x] Unit tests for WorkbookBuilder, SheetBuilder, formulas, styles
- [x] Unit tests for all core modules (131 tests, 94.62% statement coverage)
- [x] GitHub Actions CI (lint + typecheck + test on push/PR)
- [x] LICENSE file
- [x] `packages/core/README.md` (npm-facing, separate from monorepo root)
- [x] CHANGELOG.md
- [x] CONTRIBUTING.md

## v0.2 — Workbook Reading ✅

- [x] Read existing XLSX files (`WorkbookBuilder.load()`, `WorkbookBuilder.fromFile()`)
- [x] Template workflow: load → modify → re-save
- [x] Sheet-to-JSON conversion (`sheet.toJSON()`, `sheet.toAOA()`)
- [x] CSV import/export (`WorkbookBuilder.toCsv()`, `WorkbookBuilder.fromCsv()`)

## v0.3 — Data & Validation Layer ✅ (2026-05-20)

- [x] Data Validation (dropdown lists, number ranges, date constraints, custom formulas)
- [x] Conditional Formatting (color scales, data bars, icon sets, cell rules, expression rules, top N, text, time period)
- [x] `headers()` replaces `columns()` + `writeHeaders()`
- [x] Merge/header height inline (no separate `rowHeight()` calls)
- [x] Columns auto-inferred from first object row
- [x] Full numeric RC API (`setCellRC`, `mergeRC`, `styleRangeRC`, `colWidthRC`)
- [x] RC overloads for all validation and conditional formatting methods
- [x] `cellRef()`, `colRange()`, `rangeRef()` coordinate utilities
- [x] `columnIndex()` / `columnRange()` for column-keyed references
- [x] `removeDataValidation()` / `removeConditionalFormatting()`
- [x] CLI commands exposed via `pnpm cli` / `pnpm example`
- [x] assignments.ts rewritten with pure numeric coordinates

## v0.4 — Named Ranges, Tables & Charting ✅ (2026-05-21)

- [x] Named Ranges / Defined Names
- [x] Tables (structured references, totals row, banded rows, table styles)
- [x] Charts API (bar, line, pie, scatter, bubble, radar, stock, surface, waterfall, treemap, sunburst, funnel, histogram, pareto, boxWhisker, regionMap, combo)
- [x] Sparklines (in-cell mini charts)
- [x] Images (embed PNG/JPEG in cells, backgrounds, watermarks)
- [x] Hyperlinks
- [x] Rich Text (mixed bold/italic/color within a single cell)

## v0.5 — Comments, Controls & Sheet Management ✅ (2026-05-30)

- [x] Comments / Notes (classic VML + threaded)
- [x] Form Controls (checkboxes)
- [x] Sheet state (hidden / veryHidden)
- [x] Sheet operations: insert/remove/duplicate rows, columns, sheets
- [x] Print titles (repeat rows/columns on every printed page)
- [x] Page breaks (manual row/column breaks)
- [x] Sheet zoom
- [x] Custom workbook properties (title, subject, keywords, category, manager)
- [x] Ignored errors (suppress green triangles)
- [x] External workbook references

## v0.6 — Formula Expansion ✅ (2026-05-30)

- [x] Lookup & reference: VLOOKUP, HLOOKUP, INDEX, MATCH, XLOOKUP, OFFSET, INDIRECT
- [x] Conditional logic: AND, OR, NOT, SWITCH, IFS, IFERROR, IFNA
- [x] Math: ROUND, ROUNDUP, ROUNDDOWN, ABS, TRUNC, INT, MOD, CEILING, FLOOR, POWER, SQRT
- [x] Text: CONCAT, TEXT, LEFT, RIGHT, MID, LEN, TRIM, UPPER, LOWER, PROPER
- [x] Aggregate: SUMIF, SUMIFS, COUNTIF, COUNTIFS, AVERAGEIF, AVERAGEIFS, SUBTOTAL
- [x] Date: NOW, TODAY, DATE, YEAR, MONTH, DAY, EOMONTH, NETWORKDAYS
- [x] Info: ISNUMBER, ISTEXT, ISBLANK, ISERROR
- [x] Rank: RANK, LARGE, SMALL
- [x] Array formulas (FILTER, SORT, UNIQUE, CSE)

## v0.7 — Performance & Polish

- [ ] Reactive formula calculation (`workbook.calculate()` via excelts 433-function engine)
- [ ] Custom function registration (user-defined UDFs)
- [ ] Async row streaming (`WorkbookWriter` — event-based, low memory)
- [ ] Shared formulas (efficient formula replication via `fillFormula()`)
- [ ] Browser support (generate XLSX in-browser without Node.js)
- [ ] PDF export from Excel
- [ ] Markdown table import/export

## v0.8 — DX & Design Improvements

- [ ] `style()` deep-merge fills consistently with font/border
- [ ] `_finalizeAll()` use `Promise.allSettled` with error collection
- [ ] Row-level error handling in `addRows()`
- [ ] Sheet name validation with user-facing messages
- [ ] `freeze()` preserve existing view state reliably
- [ ] JSDoc for all public exports

## v1.0 — Ship

- [ ] End-to-end integration tests
- [ ] Performance benchmarks
- [ ] 90%+ test coverage
- [ ] npm publish configuration (`publishConfig`, provenance)
- [ ] API documentation site or typedoc
- [ ] Release checklist
