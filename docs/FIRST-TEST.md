# First real test through a client

The shortest path from a downloaded release to watching an AI client actually
change a workbook. About ten minutes, and it touches nothing that matters: the
workbooks are generated, disposable, and deleted afterwards.

## 1. Install

Download the latest release ZIP and extract it somewhere writable **without
administrator rights**. A managed computer usually refuses `C:\Tools`.

```powershell
%USERPROFILE%\ExcelTask\excel-task-mcp.exe --field-check
```

Run that first. It takes about a minute and proves Excel automation works on this
machine before any client is involved. If it ends `result=PASS`, continue. If not,
the digest it prints says which step refused and why.

## 2. Make something to work on

Every operation needs an existing workbook, and a locked-down machine usually
cannot script Excel to produce one, so the server can produce them itself:

```powershell
%USERPROFILE%\ExcelTask\excel-task-mcp.exe --make-fixture %USERPROFILE%\ExcelTaskDemo
```

That writes three workbooks and prints two ready-made prompts:

- `target.xlsx` - sheet `Model`, with `A1:D1` = 1,2,3,4 and a formula pattern in
  `A2:B2`
- `reference.xlsx` - sheet `Reference`, holding `=ROW()` in `A1` and `A3` with
  `A2` deliberately blank
- `macro-target.xlsm` - module `DemoModule` with a procedure `StampRun`, when the
  machine permits VBA project access

## 3. Point the client at the server

Add it under its own server key, with the full path, so any existing Excel MCP
install is untouched:

```json
{
  "servers": {
    "excel-task": {
      "command": "C:\\Users\\<you>\\ExcelTask\\excel-task-mcp.exe"
    }
  }
}
```

Restart the client so it picks the server up, then confirm it advertises exactly
one tool named `excel_task`.

## 4. Ask for the work

In a fresh session:

> Using the excel-task server, copy the `Reference` worksheet from
> `<path>\reference.xlsx` into `<path>\target.xlsx` as a new worksheet named
> `Exhibit A`, and repair the blank formula in `A1:A3`.

Expect one `excel_task` call. The first attempt commonly returns
`NeedsConfirmation` because the default binding asks what to do about an open
workbook - that is the design working, not a failure. The client should resubmit
with `workbookBinding: "Isolated"`.

## 5. Judge the result

Open `target.xlsx` yourself. Success is:

- a worksheet `Exhibit A` exists;
- `A1:A3` on it hold `=ROW()`, including `A2`, which was blank;
- the receipt says `Completed` with a `reopen-verification` check that passed;
- no stray `EXCEL.EXE` in Task Manager afterwards.

The number worth writing down is wall-clock time from pressing enter to the
workbook being correct, and how many tool calls it took. That is the measurement
the roadmap still lacks.

## Afterwards

Delete `%USERPROFILE%\ExcelTaskDemo`. Nothing in it is needed again.

If anything failed, the receipt names the step that refused. Send that text back
verbatim rather than a description of it.
