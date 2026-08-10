# ExcelTask 0.3.0

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
   periods into an immediately adjacent blank destination; then
4. recalculate, save, close owned Excel, reopen the saved workbook, and verify
   the worksheet and repairs; and
5. return a compact, structured receipt without workbook values or formula
   text.

Every inspection and execution runs in a short-lived private worker. The MCP
host enforces a two-minute deadline, reports interrupted mutations as
`Unknown`, and never kills Excel based on worker-reported process data.

The MVP accepts `.xlsx` and `.xlsm`. Copy output must keep the target file's
extension. Macro execution is disabled when ExcelTask opens a workbook.

Version 0.4 macro editing is in progress; 0.3 remains the stable release.

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
  `RepairExistingWorksheet`, or `ExtendFormulaSeries`. Supply exactly the one
  matching payload. It never accepts formula text or `FormulaR1C1`.
- The upcoming `EditMacroProcedure` operation is deliberately narrow: only an
  isolated `.xlsm` saved as a `Copy`, one named procedure, full replacement
  guarded by the expected current hash, and an optional no-argument run. Plan
  returns only that requested procedure's bounded source and hash; Apply never
  returns source. Excel Trust access remains user-controlled.
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

Version 0.3.0 is the current stable formula/exhibit release. Version 0.4 macro
editing is in progress and is not yet a stable claim. It does not yet refresh
Power Query or data models, attach to unsaved workbooks, or expose a general
automation surface. Authentication or IRM can still require a person; a macro
dialog or timeout is `Unknown` and must be reconciled before retrying.
