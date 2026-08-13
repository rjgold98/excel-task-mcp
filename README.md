# ExcelTask 0.21.0

ExcelTask is a clean-sheet, Copilot-first Excel automation engine. The selected
client model calls one high-level `excel_task` tool; deterministic code handles
workbook inspection, planning, Excel execution, verification, saving, and
cleanup. The server never chooses or invokes a model.

This repository does not depend on or preserve the former ExcelMcp interface.
There is intentionally no broad CLI or low-level tool catalog.

See the [changelog](CHANGELOG.md), [roadmap](ROADMAP.md), [privacy](PRIVACY.md), and
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
8. write bounded caller-supplied A1 formulas through the explicit
   `WriteWorksheetFormulas` operation, with immediate read-back and save/reopen
   verification; or
9. find the cells whose text matches, and rewrite the constants among them,
   leaving any cell whose text comes from a formula reported but untouched; or
10. create an empty workbook, or add an empty worksheet, never overwriting
   either; or
11. set how a bounded range looks - number format, bold, italic, font size, name
    and colour, fill, borders, column width, row height - changing no cell value;
    or
12. map a workbook's structure by reading the file directly - no Excel process
    at all: sheets, dimensions, formula/constant counts, the constant islands
    that mark manual overrides inside calculated columns, plus defined names,
    tables, and external links by file name; or
13. create, rename, restyle, resize, or convert back to plain cells one Excel
    table, keeping every cell when the table goes; or
14. create, replace, or delete one Power Query, guarded by the fingerprint a Plan
    reports. Plan returns the M expression too - and an M expression usually names
    a server, so PRIVACY.md section 4 states what that puts on the wire; or
15. create, replace, or delete one Data Model measure, guarded the same way, with
    Plan returning the DAX; or
16. create or delete one Data Model relationship, naming the many side and the
    one side. Replace is refused: a relationship has no editable middle, so the
    honest instruction is to delete it and create the new one; then
17. recalculate, save, close owned Excel, reopen the saved workbook, and verify
    the worksheet, repairs, procedure, written values, formulas, replacements, format,
    table, query, measure, or relationship; and
18. return a compact, structured receipt.

Operations 5 and 6 never write, and say so from evidence: their receipts carry a
check proving the workbook's size and timestamp were identical before and after.
The receipt withholds workbook values and formula text everywhere except the
range read, where the contents are the entire request rather than incidental.
Formula text is accepted only by `WriteWorksheetFormulas` and is still withheld
from the receipt.

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
  `WriteWorksheetFormulas`,
  `FindReplace`, `Create`, `SetRangeFormat`, `ScanWorkbookStructure`, `ManageTable`,
  `ManageQuery`, `ManageModelMeasure`, or `ManageModelRelationship`. Supply exactly the one
  matching payload. `WriteWorksheetValues` takes constants only and rejects any
  value starting with `=`. `WriteWorksheetFormulas` is the explicit path for
  bounded A1 formulas; `FindReplace` refuses a
  replacement that would leave a cell starting with `=` even when the
  replacement text itself is a legal constant. `RepairExistingWorksheet` and
  `ExtendFormulaSeries` write formulas, but only ones inferred from formulas
  already on the sheet - never text a caller supplied.
- Start with `ScanWorkbookStructure` when the worksheet names are unknown: it
  names them and starts no Excel process at all. Every other operation requires
  a worksheet name it otherwise has no way to discover.
- `FindReplace` matches plain text on what the cell displays; `*` and `?` are
  literal rather than wildcards, which is where it deliberately differs from
  Excel's own Find. It searches at most 10,000 cells, and refuses rather than
  silently searching part of a larger used range.
- `Create` writes the target it names, so it takes no save destination and
  requires binding `Isolated`. It never overwrites: an existing file, or an
  existing worksheet name, is refused outright.
- `SetRangeFormat` sets only how a range looks, on at most 10,000 cells: number
  format, bold, italic, font size, name and colour, fill, borders and their
  weight, column width, row height. Every field is optional and independent, and
  at least one must be supplied. It changes no cell values and sets no
  conditional formats. Number-format codes are not trimmed, because a format's
  leading and trailing spaces are what align parenthesised negatives under
  positives. Colours are `#RRGGBB`. A misspelled font name cannot be caught by
  verification: Excel stores whatever name it is given and substitutes only when
  rendering, so the read-back always agrees.
- `CopyExhibit` binds the copied worksheet's references to the workbook it
  landed in. Excel rewrites `=Data!A1` into a link back to the source during the
  copy, because the sheet it names is not yet in the destination; the copy still
  calculates, which is what makes that quiet. A sheet the destination does not
  have is deliberately left alone rather than turned into `#REF`, and the
  `copy-rebind` check names it.
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
- `WriteWorksheetFormulas` accepts at most 200 single-cell A1 formulas, each no
  longer than 8,192 characters, within a 400-cell span and 768 KiB of UTF-8
  formula text per request. Excel reads them back before saving and after reopen.
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

Fifteen operations ship today: formula, exhibit, macro editing, discovery, range
reading, constant writes, find/replace, creation, worksheet repair, structure
scanning without Excel, range formatting, table management, Power Query
mutation, and Data Model measures and relationships. The version is the heading
above, and the build is always the
[latest release](https://github.com/rjgold98/excel-task-mcp/releases/latest).
It does not yet refresh Power Query or data models, attach to unsaved workbooks,
edit sheet or class modules, or expose a general automation surface.
Authentication or IRM can still require a person; a macro dialog or timeout is
`Unknown` and must be reconciled before retrying.

Macro editing is verified against desktop Excel on a machine where VBA project
access is trusted. On a managed computer that policy is set by the organization,
so the first work-computer run should be treated as the real gate.

## Where Microsoft stands on this, stated plainly

ExcelTask automates desktop Excel through COM. Microsoft's
[KB 257757](https://support.microsoft.com/en-us/help/257757/considerations-for-server-side-automation-of-office)
says outright that it "does not recommend or support server-side Automation of Office," and points
at Open XML file manipulation as "the recommended and supported method for handling changes to
Office files from a service."

That guidance is aimed at unattended server-side use. ExcelTask's intended shape - an interactive
desktop, a signed-in person, a chat in front of them - is client-side automation, which is not what
the KB withholds support from. Anyone considering it as an unattended service should read the KB as
written and decide accordingly.

The same KB names the hazard this design spends most of its complexity on: "A modal dialog box on a
non-interactive desktop cannot be dismissed. Therefore, that thread stops responding (hangs)
indefinitely." That is the failure the supervised worker, the hard deadline, the dialog sentry, and
the proof-of-exit exist to bound - not to eliminate, since nothing can, but to make it a reported
`Unknown` rather than a hung process holding your workbook.

`ScanWorkbookStructure` is the one operation on Microsoft's recommended path: it reads the Open XML
package directly and starts no Excel at all. See [docs/LANDSCAPE.md](docs/LANDSCAPE.md) for how
that trade-off compares against the other spreadsheet MCP servers.
