# ExcelTask 0.14.0

ExcelTask is a clean-sheet, Copilot-first Excel automation engine. The selected
client model calls one high-level `excel_task` tool; deterministic code handles
workbook inspection, planning, Excel execution, verification, saving, and
cleanup. The server never chooses or invokes a model.

This repository does not depend on or preserve the former ExcelMcp interface.
There is intentionally no broad CLI or low-level tool catalog.

See the [changelog](CHANGELOG.md), [roadmap](ROADMAP.md), and
[latest release](https://github.com/rjgold98/excel-task-mcp/releases/latest).

## What works

One request can perform exactly one operation:

1. copy an explicitly named reference worksheet, including from another
   workbook, into an existing target workbook;
2. repair blank cells on an existing worksheet only when matching neighboring
   patterns make the intended formula unambiguous; or
3. extend a proven formula series right or down from two adjacent evidence
   periods into an immediately adjacent blank destination; or
4. replace one named standard-module VBA procedure in an `.xlsm`, guarded by the
   expected current hash, optionally running it afterwards;
5. report a read-only map of the workbook's worksheets, tables, defined names,
   queries, connections, macro procedures, data model, pivots, and external
   links; or
6. read the contents of one bounded worksheet range, as displayed values or as
   R1C1 formulas; or
7. write constants into named cells, never formula text; or
8. find the cells whose text matches, and rewrite the constants among them,
   leaving any cell whose text comes from a formula reported but untouched; or
9. create an empty workbook, or add an empty worksheet, never overwriting
   either; or
10. set one number format code across a bounded range, changing how numbers
    display and never the numbers themselves; or
11. map a workbook's structure by reading the file directly - no Excel process
    at all: sheets, dimensions, formula/constant counts, and the constant
    islands that mark manual overrides inside calculated columns; then
12. recalculate, save, close owned Excel, reopen the saved workbook, and verify
    the worksheet, repairs, procedure, written values, replacements, or format;
    and
13. return a compact, structured receipt.

Operations 5 and 6 never write, and say so from evidence: their receipts carry a
check proving the workbook's size and timestamp were identical before and after.
The receipt withholds workbook values and formula text everywhere except the
range read, where the contents are the entire request rather than incidental.

Every inspection and execution runs in a short-lived private worker. The MCP
host enforces a two-minute deadline, reports interrupted mutations as
`Unknown`, and never kills Excel based on worker-reported process data.

The MVP accepts `.xlsx` and `.xlsm`. Copy output must keep the target file's
extension. Macros stay disabled when ExcelTask opens a workbook unless the
request explicitly asks to run the procedure it just edited; workbook events are
suppressed either way.

## Safe choices

- Begin with `workbookBinding: "AskIfOpen"`.
- If the exact target is open, choose `UseOpen` with `save: "Same"` to edit that
  live workbook, or `Isolated` with `save: "Copy"` to use the saved on-disk
  file without controlling the user's Excel process.
- `overwriteConfirmed: true` is explicit authorization to replace the selected
  save destination. ExcelTask otherwise returns `NeedsConfirmation` before
  mutation.
- Use `mode: "Plan"` for a non-mutating preview. Normal execution uses
  `mode: "Apply"`.
- The request has one `operation` union: `CopyExhibit`,
  `RepairExistingWorksheet`, `ExtendFormulaSeries`, `EditMacroProcedure`,
  `AuditWorkbookFlows`, `ReadWorksheetRange`, `WriteWorksheetValues`,
  `FindReplace`, `Create`, `SetNumberFormat`, or `ScanWorkbookStructure`. Supply exactly the one matching payload. It never
  accepts formula text or `FormulaR1C1`: `WriteWorksheetValues` takes constants
  only and rejects any value starting with `=`, and `FindReplace` refuses a
  replacement that would leave a cell starting with `=` even when the
  replacement text itself is a legal constant.
- Start with `AuditWorkbookFlows` when the worksheet names are unknown. Every
  other operation requires a worksheet name it otherwise has no way to discover.
- `FindReplace` matches plain text on what the cell displays; `*` and `?` are
  literal rather than wildcards, which is where it deliberately differs from
  Excel's own Find. It searches at most 10,000 cells, and refuses rather than
  silently searching part of a larger used range.
- `Create` writes the target it names, so it takes no save destination and
  requires binding `Isolated`. It never overwrites: an existing file, or an
  existing worksheet name, is refused outright.
- `SetNumberFormat` sets only the number format, on at most 10,000 cells. It
  changes no cell values and sets no fonts, fills, borders, widths, or
  conditional formats. Codes are not trimmed, because a format's leading and
  trailing spaces are what align parenthesised negatives under positives.
- `ScanWorkbookStructure` is the one operation that never starts Excel: it reads
  the file as the ZIP of XML it physically is. Counts and addresses only, never
  contents. Encrypted workbooks cannot be scanned this way; the Excel-based
  operations still open them.
- `ReadWorksheetRange` returns at most 400 cells, omits blank ones, and caps each
  cell's text. It rejects a range larger than that rather than truncating to a
  partial answer that would read as a complete one.
- The `EditMacroProcedure` operation is deliberately narrow: only an isolated
  `.xlsm` saved as a `Copy`, one named standard-module procedure, full
  replacement guarded by the expected current hash, and an optional no-argument
  run. Plan returns only that requested procedure's bounded source and hash;
  Apply never returns source. Signed or locked VBA projects and automatic-entry
  procedures such as `Auto_Open` are refused. Reading or editing VBA requires
  "Trust access to the VBA project object model"; ExcelTask reports when that is
  blocked and never changes the setting itself.
- Repair and copy-exhibit operations accept at most 16 non-overlapping A1
  ranges and scan at most 10,000 cells. Series extension accepts two evidence
  periods, 1–24 adjacent destination periods, and no more than 2,000 planned
  mutations.
- `mode: "Plan"` analyzes only; it never changes, saves, or recalculates a
  workbook.

Example tool arguments for a new copy:

```json
{
  "request": {
    "targetWorkbookPath": "C:\\Work\\Target.xlsx",
    "operation": {
      "kind": "CopyExhibit",
      "copyExhibit": {
        "referenceWorkbookPath": "C:\\Work\\Reference.xlsx",
        "referenceWorksheet": "Template",
        "newWorksheetName": "Exhibit A",
        "repairRanges": ["A1:A20"]
      }
    },
    "mode": "Apply",
    "workbookBinding": "AskIfOpen",
    "save": "Copy",
    "outputWorkbookPath": "C:\\Work\\Target - updated.xlsx",
    "overwriteConfirmed": false
  }
}
```

## Build and test

Requires Windows, 64-bit Microsoft Excel, and the .NET 10 SDK.

```powershell
dotnet restore ExcelTask.slnx -p:NuGetAudit=false
dotnet test ExcelTask.slnx --filter "RunType!=OnDemand" -p:NuGetAudit=false
dotnet run --project src\ExcelTask.McpServer
```

The full gate, including the tests that drive real desktop Excel, runs one project at a time:

```powershell
.\tools\gate.ps1
```

Use it rather than `dotnet test ExcelTask.slnx` with no filter. The solution form runs test
assemblies in parallel, and the two real-Excel assemblies each assert that no Excel process was
left behind - concurrently, each sees the other's.

The disposable desktop-Excel acceptance tests are serial and opt-in:

```powershell
.\scripts\Test-Mvp.ps1 -IncludeExcel
```

## Side-by-side MCP setup

Use a unique server key and the full executable path. Do not put it on `PATH`
or replace an existing `excel-mcp` entry. A published local build can be
configured in `.vscode/mcp.json` as:

```json
{
  "servers": {
    "excel-task": {
      "command": "C:\\Tools\\ExcelTask\\excel-task-mcp.exe"
    }
  }
}
```

This leaves the original Excel MCP installation, command, server key, and
client cache untouched.

## Current boundary

Version 0.14.0 is the current stable release: formula, exhibit, macro editing,
discovery, range reading, constant writes, find/replace, creation, and number
formats.
It does not yet set fonts, fills, borders or widths, refresh Power Query or data
models, attach to unsaved workbooks, edit sheet or class modules, or expose a
general automation surface.
Authentication or IRM can still require a person; a macro dialog or timeout is
`Unknown` and must be reconciled before retrying.

Macro editing is verified against desktop Excel on a machine where VBA project
access is trusted. On a managed computer that policy is set by the organization,
so the first work-computer run should be treated as the real gate.
