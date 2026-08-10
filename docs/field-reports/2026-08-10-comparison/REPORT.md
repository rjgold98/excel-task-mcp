# Field report - tool-surface comparison, 2026-08-10

Field task 001, run on the managed work computer with ExcelTask v0.6.1 and
relayed by the owner. The field agent cannot write to this repository, so this
file is transcribed from its report; the digest is verbatim.

## Result

Exit code 0, `result=PASS`. Release ZIP SHA-256 matched
`191b72549ddd543a9d139790a33ce67ae948070359bc3a625931b80dbffcfa3b`.

```text
----- EXCELTASK FIELD DIGEST -----
excel=16.0.20228 vbom=1 vbawarn=1
self  v0.6.1 tools=1 bytes=7164
other tools=25 bytes=58324 ratio=8.1x
CopyExhibit (Plan)         Planned             8.1s L0
CopyExhibit (Apply)        Completed          14.3s L0
ExtendFormulaSeries (Apply Completed          16.8s L0
EditMacroProcedure (Plan)  Planned             8.0s L0
EditMacroProcedure (Apply+ Completed          16.2s L0
leaked=0 result=PASS
----- END DIGEST -----
```

## Tool surface

| Server | Tools | `tools/list` bytes |
|---|---|---|
| ExcelTask v0.6.1 | 1 | 7,164 |
| Original Excel MCP | 25 | 58,324 |

The original advertises **8.1x** the context cost before any work is requested.
Note that 25 is a count of MCP *tools*, not of the operations behind them; the
earlier figure of roughly 234 described operations and is not comparable.

## Owned-process boundary, observed in the field

Four `EXCEL.EXE` processes were already running before the check and the same
four were running after. ExcelTask created and cleaned up its own instances,
terminated nothing, and reported `leaked=0`. This is the first observation of
that boundary holding on a machine where the user's own Excel was live.

## Timings

Operations ran roughly two to three times slower than on the development
machine - 8-17s against 3-7s - which is the expected shape for a managed
computer with security tooling, four live Excel instances, and synced folders.
This is a baseline, not a regression: there is no earlier work-computer figure
to compare against.

## What was not measured

The six owner-run prompt-to-completion comparisons were not performed. Verbatim
blocker:

> Owner-controlled fresh-session client interaction was unavailable in this
> field session; no owner-run prompts were available, so the six runs were not
> started.

So this run measures **context cost only**. Nothing here supports a claim about
end-to-end speed, token consumption during real work, or error rates against the
original server. Those remain unmeasured.

## Environment and history

- Excel 16.0.20228, `AccessVBOM=1`, `VBAWarnings=1`.
- The original MCP was found in the global client configuration with no launch
  arguments. Its path was withheld deliberately under the reporting privacy
  boundary, which was the correct call.
- The intended install location `C:\Tools\ExcelTask` was **not writable**; a
  user-writable temporary location was used instead. No permission was changed.
- The PowerShell field-check script could not run: both installed PowerShell
  hosts are governed by Constrained Language Mode. No bypass was attempted.
- v0.5.0 compiled runs failed at Excel startup with
  `COMException: Server execution failed (0x80080005 CO_E_SERVER_EXEC_FAILURE)`
  and exited 1 without producing measurements. v0.6.1 completed cleanly.

## Confirmations recorded by the field agent

No real workbook was opened. No workbook contents, cell values, formulas, VBA
source, real workbook names, customer or patient information, credentials,
account names, internal server names, or machine paths were transmitted. No
Excel, Trust Center, Constrained Language Mode, Group Policy, permission, or
system setting was changed. The original MCP was not modified. No git write,
repository creation, or account switch was performed.
