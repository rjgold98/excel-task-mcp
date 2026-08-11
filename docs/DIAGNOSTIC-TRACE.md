# Diagnostic trace — temporary, development only

A narration of every step the server takes, written to a file, so a failure on the work computer
can travel to the machine where fixes happen. **Off unless you turn it on.**

## Turning it on

Set one environment variable to a file path before the MCP client starts the server:

```powershell
$env:EXCELTASK_TRACE = "C:\Users\<you>\Documents\exceltask-trace.log"
```

To set it for a client that launches the server itself, add it to the server entry in
`.vscode/mcp.json` (or the equivalent Copilot config):

```json
{
  "servers": {
    "excel-task": {
      "command": "C:\\path\\to\\excel-task-mcp.exe",
      "env": { "EXCELTASK_TRACE": "C:\\Users\\<you>\\Documents\\exceltask-trace.log" }
    }
  }
}
```

Turn it off by clearing the variable. There is no other switch, no CLI flag, and no schema field —
the model never sees that tracing exists.

The file is appended to, never truncated, so one file can hold a whole session. Delete it when you
want a clean run.

## What it records, and why that matters

The trace is designed to be **safe to paste into a chat**, and every file it writes opens by saying
so. That constraint is the point: the work computer holds real client workbooks, and a debug log
that carried their contents would be worse than no log at all.

**Records:**

- every phase and how long it took
- the operation kind, mode, workbook binding, save mode, and overwrite flag
- worksheet names and A1 ranges
- workbook **file names only** — never the directory they sit in
- owned Excel process ids, and when each one started
- the final status, and every check with its detail
- the exception type and message behind an unhandled worker failure

**Never records:**

- cell values, formulas, or any workbook contents
- VBA source
- connection strings, server names, or query text
- full paths, user names, or machine names

## Reading one

```
15:23:54.968 [7e3175a6] === worker start ===
15:23:54.972 [7e3175a6]   execute: SetNumberFormat mode=Apply binding=Isolated save=Same ...
15:23:54.986 [7e3175a6]   phase start session-open
15:23:55.293 [7e3175a6]   owned Excel started, pid 70452
15:23:55.384 [7e3175a6]   phase end   session-open (398 ms)
...
15:24:00.542 [7e3175a6] === worker end: Completed: ... (5573 ms total) ===
```

The bracketed id is the task. A tool call and its worker have different ids — the tool call brackets
the whole round trip, and one or two worker blocks sit inside it (inspection first, then execution).

**A phase that never ends is where it hung.** That is the single most useful thing in the file: if
the last line is `phase start reopen-verification` with no matching end, Excel did not come back
from opening the saved workbook, and the modal-dialog or file-lock paths are where to look.

## What the first run already showed

On a `SetNumberFormat` apply against a 12-row sheet:

| phase | duration |
|---|---:|
| session-open | 398 ms |
| format-preflight | 5 ms |
| number-format | 22 ms |
| save | 17 ms |
| **primary-cleanup** | **2,814 ms** |
| **reopen-verification** | **2,297 ms** |
| total | 5,573 ms |

The Excel work is 44 ms. Closing owned Excel and proving it exited, plus reopening the saved file to
verify, is 5.1 of the 5.6 seconds — **92%**.

This is direct evidence on the roadmap's longest-standing open question, which recorded a ~4 s
unaccounted gap in a macro Apply and speculated it was "worker startup, MCP round trips, and model
coordination." On this operation it is neither: it is owned-Excel teardown and the verification
reopen. Both are load-bearing (proving the process exited is the product's central claim, and
verification is what makes a receipt mean anything), so neither is simply removable — but they are
now measured rather than guessed, which is what the roadmap said had to happen before anything here
was optimized.

## Removing it

It was built to come out cleanly:

1. Delete `src/ExcelTask.Excel/DiagnosticTrace.cs`.
2. Delete this file.
3. In `src/ExcelTask.Excel/WorkbookWorkerHost.cs`, remove the `trace` local, the `DiagnosticTrace?`
   parameter on `ProtocolObserver`, the `trace?.` calls, and the `DescribeRequest` / `DescribeResult`
   methods.
4. In `src/ExcelTask.McpServer/ExcelTaskTool.cs`, remove the two `trace` lines in `RunAsync`.

Nothing else references it. No test depends on it, no receipt carries it, and the published tool
schema is byte-identical with it present or absent.
