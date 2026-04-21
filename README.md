# DSS

`dss-go` is a Go command-line tool for loading DSS, CSV, XLS, XLSX, and XLSM workbooks, computing formulas, editing them in a full-screen terminal spreadsheet view with bordered cells, visualizing workbook summaries in a termdash dashboard, and exporting back to DSS.

## Current Commands

```text
dss <input> [--output output.dss]
dss tui <input> [--output output.dss]
dss dashboard <input>
dss convert <input> <output.dss>
dss set <input> <output.dss> --sheet NAME --cell A1 --value VALUE
```

Passing a workbook path directly opens the interactive TUI.

The TUI shows a bordered spreadsheet grid, a dedicated formula/value bar for the selected cell, color-coded formula/error states, and an inline editor with cursor movement.

## TUI Controls

- Move with arrow keys or `h`, `j`, `k`, `l`
- Edit the current cell with `Enter` or `e`
- Move inside the edit buffer with arrow keys, `Home`, `End`, `Ctrl+A`, and `Ctrl+E`
- Clear edit content with `Ctrl+U`
- Save DSS with `Ctrl+S`
- Recalculate with `r`
- Switch sheets with `[` and `]`
- Cycle display mode with `m`
- Open the dashboard for the current file with `d`
- Quit with `q`

## Dashboard

- `dss dashboard <input>` opens a workbook summary dashboard built with `termdash`
- The dashboard shows workbook totals, sheet ranges, formula/error counts, and quick usage hints
- Close it with `q`, `Ctrl+Q`, or `Ctrl+C`

## Current Formula Support

- Arithmetic: `+`, `-`, `*`, `/`, parentheses
- Comparisons: `=`, `<>`, `<`, `<=`, `>`, `>=`
- Cell references: `A1`
- Ranges: `A1:B10`
- Boolean literals: `TRUE`, `FALSE`
- Functions: `SUM`, `AVG`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `COUNTA`, `IF`, `ABS`, `ROUND`, `POWER`, `MOD`, `SQRT`, `AND`, `OR`, `NOT`, `LEN`, `CONCAT`, `CONCATENATE`
- Imported XLSX/XLSM formulas with unsupported functions fall back to the cached workbook value when one is available.

I'm still working on adding more functions and improving error handling, but the current set covers many common use cases.

## Format Notes

- `.xlsx` and `.xlsm` are imported with `excelize`.
- `.xls` is imported with `extrame/xls` and currently loads cell values only.
- `.xlsm` macros are ignored.
- DSS export writes sparse multi-anchor blocks instead of flattening each sheet into one padded rectangle.

## Example

```text
go run ./cmd/dss ../sample.xlsx
go run ./cmd/dss tui ../sample.xlsx --output ../sample.dss
go run ./cmd/dss dashboard ../sample.xlsx
go run ./cmd/dss convert ../sample.xlsx ../sample.dss
go run ./cmd/dss set ../sample.dss ../edited.dss --sheet Sheet1 --cell C3 --value "=SUM(A1:B2)"
```