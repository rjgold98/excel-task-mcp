# ExcelTask 0.1.0

ExcelTask is a clean-sheet, Copilot-first Excel automation engine. The selected
client model calls one high-level `excel_task` tool; deterministic code handles
workbook inspection, planning, Excel execution, verification, saving, and
cleanup. The server never chooses or invokes a model.

This repository does not depend on or preserve the former ExcelMcp interface.
There is intentionally no broad CLI or low-level tool catalog.

See the [changelog](CHANGELOG.md), [roadmap](ROADMAP.md), and
[latest release](https://github.com/rjgold98/excel-task-mcp/releases/latest).

## What works

One request can:

1. copy an explicitly named reference worksheet, including from another
   workbook, into an existing target workbook;
2. repair blank cells only when matching neighboring FormulaR1C1 patterns make
   the intended formula unambiguous;
3. recalculate, save, close owned Excel, reopen the saved workbook, and verify
   the worksheet and repairs; and
4. return a compact, structured receipt without workbook values or formula
   text.

The MVP accepts `.xlsx` and `.xlsm`. Copy output must keep the target file's
extension. Macro execution is disabled when ExcelTask opens a workbook.

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
- Use `formulaRepairRanges: []` when only the worksheet copy is needed. A repair
  range must include the neighboring formula evidence and all ranges together
  are capped at 10,000 cells.

Example tool arguments for a new copy:

```json
{
  "request": {
    "targetWorkbookPath": "C:\\Work\\Target.xlsx",
    "referenceWorkbookPath": "C:\\Work\\Reference.xlsx",
    "referenceWorksheet": "Template",
    "newWorksheetName": "Exhibit A",
    "formulaRepairRanges": ["A1:A20"],
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

Version 0.1.0 proves the formula/exhibit vertical slice. It does not yet edit VBA,
refresh Power Query or data models, attach to unsaved workbooks, or expose a
general automation surface. A blocked Office authentication or IRM dialog can
still stall the in-process COM worker; hard worker-process deadlines are a
post-MVP reliability gate, not a claimed capability of this build.
