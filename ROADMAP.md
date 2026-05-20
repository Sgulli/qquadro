# Roadmap

## v0.1 — Foundation

- [x] Add test framework (vitest)
- [x] Unit tests for WorkbookBuilder, SheetBuilder, formulas, styles
- [x] Unit tests for all core modules (131 tests, 94.62% statement coverage)
- [x] GitHub Actions CI (lint + typecheck + test on push/PR)
- [x] LICENSE file
- [x] `packages/core/README.md` (npm-facing, separate from monorepo root)
- [x] CHANGELOG.md
- [x] CONTRIBUTING.md

## v0.2 — Workbook Reading

- [x] Read existing XLSX files (`WorkbookBuilder.load()`, `WorkbookBuilder.fromFile()`)
- [x] Template workflow: load → modify → re-save
- [x] Sheet-to-JSON conversion (`sheet.toJSON()`, `sheet.toAOA()`)
- [x] CSV import/export (`WorkbookBuilder.toCsv()`, `WorkbookBuilder.fromCsv()`)

## v0.3 — Data & Validation Layer

- [ ] Data Validation (dropdown lists, number ranges, date constraints, custom formulas)
- [ ] Conditional Formatting (color scales, data bars, icon sets, cell rules)
- [ ] Named Ranges / Defined Names
- [ ] Tables (structured references, totals row, banded rows, table styles)

## v0.4 — Charting & Visuals

- [ ] Charts API (bar, line, pie, scatter, bubble, radar, stock, surface, waterfall, treemap, sunburst, funnel, histogram, pareto, boxWhisker, regionMap, combo)
- [ ] Sparklines (in-cell mini charts)
- [ ] Images (embed PNG/JPEG in cells, backgrounds, watermarks)
- [ ] Hyperlinks

## v0.5 — Rich Content

- [ ] Rich Text (mixed bold/italic/color within a single cell)
- [ ] Comments / Notes (classic VML + threaded)
- [ ] Form Controls (checkboxes)
- [ ] Sheet state (hidden / veryHidden)

## v0.6 — Formula Expansion

- [ ] Lookup & reference: VLOOKUP, HLOOKUP, INDEX, MATCH, XLOOKUP, OFFSET, INDIRECT
- [ ] Conditional logic: AND, OR, NOT, SWITCH, IFS, IFERROR, IFNA
- [ ] Math: ROUND, ROUNDUP, ROUNDDOWN, ABS, TRUNC, INT, MOD, CEILING, FLOOR, POWER, SQRT
- [ ] Text: CONCAT, TEXT, LEFT, RIGHT, MID, LEN, TRIM, UPPER, LOWER, PROPER
- [ ] Aggregate: SUMIF, SUMIFS, COUNTIF, COUNTIFS, AVERAGEIF, AVERAGEIFS, SUBTOTAL
- [ ] Date: NOW, TODAY, DATE, YEAR, MONTH, DAY, EOMONTH, NETWORKDAYS
- [ ] Info: ISNUMBER, ISTEXT, ISBLANK, ISERROR
- [ ] Rank: RANK, LARGE, SMALL
- [ ] Array formulas (FILTER, SORT, UNIQUE, CSE)

## v0.7 — Advanced Workbook

- [ ] Sheet operations: insert/remove/duplicate rows, columns, sheets
- [ ] Print titles (repeat rows/columns on every printed page)
- [ ] Page breaks (manual row/column breaks)
- [ ] Sheet zoom
- [ ] Custom workbook properties (title, subject, keywords, category, manager)
- [ ] Ignored errors (suppress green triangles)
- [ ] External workbook references

## v0.8 — Performance & Polish

- [ ] Reactive formula calculation (`workbook.calculate()` via excelts 433-function engine)
- [ ] Custom function registration (user-defined UDFs)
- [ ] Async row streaming (`WorkbookWriter` — event-based, low memory)
- [ ] Shared formulas (efficient formula replication via `fillFormula()`)
- [ ] Browser support (generate XLSX in-browser without Node.js)
- [ ] PDF export from Excel
- [ ] Markdown table import/export

## v0.9 — DX & Design Improvements

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
