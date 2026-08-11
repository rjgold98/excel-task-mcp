# What Excel configuration is actually worth

Measured on the development machine against real desktop Excel, 2026-08-10, with
`tools/excel-config-probe.ps1` and `tools/excel-calc-probe.ps1`. Reproduce before trusting; these
are wall-clock numbers from one machine, and the shape of the answer matters more than the digits.

## The question

ExcelTask configures a private Excel instance before doing any work:

```
Visible = false
DisplayAlerts = false
EnableEvents = false
AutomationSecurity = ForceDisable
```

Those were chosen from reputation. The obvious next candidates - manual calculation, screen
updating, print communication - are recommended everywhere for Excel automation. The question was
whether adding them would make ExcelTask faster. The answer for every one of them is no, and the
reason is worth writing down because it will otherwise be proposed again.

## Where the time goes

Median of interleaved trials, a 2,000-row repair on a workbook with a dependency chain:

| Phase | Cost |
| --- | --- |
| Excel launch | 274-482 ms |
| Workbook open | 60-75 ms |
| Write 2,000 formulas (43 batched calls) | 58 ms |
| CalculateFull | ~1 ms |
| Save | 18 ms |

Launching Excel is the entire cost. Everything ExcelTask does *inside* Excel is rounding error
beside starting it. That single fact decides every tuning question below.

## Manual calculation makes it slower

| Calculation during writes | Write | Calc | Restore | Save | Total |
| --- | --- | --- | --- | --- | --- |
| automatic (shipping) | 58 ms | 1 ms | 0 ms | 18 ms | **76 ms** |
| manual, restored before save | 83 ms | 1 ms | 1 ms | 18 ms | **103 ms** |

Manual calculation exists to stop Excel recalculating once per write. ExcelTask does not write once
per cell: repairs are grouped by identical R1C1 formula and written as multi-area ranges, so 2,000
cells cost 43 calls, and there is no per-cell recalculation left to suppress. What remains is the
mode switching and the dirty-chain bookkeeping a later `CalculateFull` has to redo, which is why the
"optimization" costs 27 ms rather than saving any.

The per-cell loop tells the opposite story - 586 ms automatic against 489 ms manual - which is
exactly why it must not be used to justify the setting. That loop is not the code we ship.

**Do not add `Calculation = xlCalculationManual`.** If the write strategy ever stops batching, this
conclusion expires with it.

### The safety question it raised, answered

Excel records a calculation mode in the workbook file. Saving while the application is in manual
mode could hand the user back a model that silently stops recalculating - quiet damage of exactly
the kind this project exists not to do. Measured directly: a file saved from an instance left in
manual mode reopens in **automatic** calculation in a fresh instance. The hazard is not real on this
Excel build. Recorded because the question is worth more than the answer: if manual calculation is
ever adopted, restore the prior mode before saving anyway.

## The settings that do nothing

`ScreenUpdating = false`, `PrintCommunication = false`, and `DisplayStatusBar = false` all measured
within noise of the baseline. They are advice for automation driving a *visible* Excel. ExcelTask's
instance is invisible and has no window to update, no status bar to repaint, and no page layout to
negotiate with a printer driver. Adding them would be three more settings to reason about, three
more things to restore, and no faster.

## The macro path: one second, and 59% of it is launching Excel

`tools/excel-macro-probe.ps1`, median of 3 trials, the full edit-run-save-verify sequence:

| Step | Cost | Share |
| --- | --- | --- |
| launch | 284 ms | 27% |
| configure | 29 ms | 3% |
| open .xlsm | 63 ms | 6% |
| **first touch of `VBProject`** | 77 ms | 7% |
| find component | 6 ms | 1% |
| read procedure | 6 ms | 1% |
| replace procedure | 4 ms | 0% |
| `Application.Run` | 5 ms | 0% |
| save as .xlsm | 29 ms | 3% |
| close and quit | 17 ms | 2% |
| verification launch | 326 ms | 31% |
| verification open | 82 ms | 8% |
| verification read | 107 ms | 10% |
| **total** | **1,035 ms** | |

The VBIDE work that gets blamed for macro slowness - finding the component, reading the procedure,
replacing it, running it - is **21 ms of 1,035**. There is nothing to optimize there. The two Excel
launches are 610 ms, 59% of the whole sequence.

This corrects the roadmap's standing attribution. The macro workflow was measured at 28.1s against
the original server's 26.4s and recorded as a regression caused by "Plan and Apply each opening
their own Excel". Four launches is about 1.2 seconds. The other 27 seconds were never Excel: they
are worker startup, MCP round trips, and model coordination. Optimizing COM further cannot move
that number - fewer round trips can, which is what the one-deep-tool design is already for.

## Where the remaining speed actually is

A mutating Apply launches Excel **twice**: once to make the change, and once more - after proving
the first process exited and released its file lock - to reopen the saved file and verify it. On
the macro path that second launch alone is 31% of all Excel time, and the pair is 59%.

It cannot simply be removed. Verifying in the process that just did the writing would be verifying
against the same memory that produced them; the reopen is what makes the receipt evidence rather
than intent.

The version worth building is to start the verification instance **early**, so its launch overlaps
the write and save phases instead of following them. That keeps every property intact - still a
separate process, still freshly launched, still opening the file only after the primary closed - and
costs only latency. It is not built yet, because a pre-launched instance is a new way to leak an
Excel process on every early-return path, and this project's central claim is that it never does.
That work needs the leak test to come first.

`ReadWorksheetRange` and `AuditWorkbookFlows` already launch once. The most-requested operation is
therefore already at the floor.
