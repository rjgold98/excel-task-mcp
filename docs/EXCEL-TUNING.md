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

So it is started **early** instead, while the primary session is still writing. Every property is
intact - still a separate process, still freshly launched, still opening the saved file only after
the primary closed and released its lock. Only when the clock starts changed. Shipped in v0.9.1.

### What it was actually worth

A/B on one successful macro Apply through the real runtime, 5 trials each, measured with a harness
that calls `ExecuteAsync` directly rather than through the test suite:

| Build | Median | Range |
| --- | --- | --- |
| serial launch | 5,398 ms | 5,383 - 5,528 |
| pre-launched | **5,244 ms** | 5,125 - 5,343 |

**154 ms, about 3%.** The ranges do not overlap across five trials each, so the effect is real, but
it is well under the 326 ms the launch costs in isolation: two Excels starting at once contend for
the same disk and CPU, so overlapping a launch does not make it free.

The test suite cannot measure this at all. Those runs are dominated by test-host startup and by the
leak detector's settle waits - both constant, both far larger - and an A/B there returned an
11 ms difference on a 24-second test, which is noise.

## The four seconds, found: Excel takes 2.3 seconds to die

Bisected by timing four operation shapes that skip different parts of the lifecycle, since the
phase observer is internal:

| Shape | Median |
| --- | --- |
| read - 1 launch, no save, no verify | 2,619 ms |
| macro Plan - 1 launch, VBA read, no save | 2,817 ms |
| formula Apply - 2 launches, save, verify | 5,197 ms |
| macro Apply - 2 launches, save, verify, VBA | 5,802 ms |

Cost tracks **sessions**, not work. One session is ~2.6 s; two are ~5.2 s. The VBA that the macro
path is blamed for adds 200 ms to a Plan and 600 ms to an Apply.

Measured directly in C#, with proxies actually released - the earlier PowerShell attempt is recorded
below as invalid:

| Step | Median |
| --- | --- |
| launch | 419 ms |
| `Quit` + release proxies | **13 ms** |
| wait for the process to actually exit | **2,267 ms** |

So a session is 419 ms to start, 13 ms to dismiss, and **2,267 ms to die**. The runtime waits for
that death because "no Excel is ever left behind" is the product's central claim, and the wait is
what earns it. A mutating Apply pays it twice: 4.5 of its 5.2 seconds is starting and burying Excel.

### Overlapping the death was tried, and made it slower

The wait is a Win32 process wait, not a COM call, so it can leave the STA thread. The primary's
death was moved to run alongside the verification and joined afterwards, with an explicit bounded
wait for the file lock replacing the release that the death wait used to provide for free.

| Build | formula Apply | macro Apply |
| --- | --- | --- |
| serial death wait | 5,288 ms | 5,987 ms |
| overlapped | 5,528 ms | 6,230 ms |

**About 240 ms slower, both shapes.** There is almost nothing to hide the wait behind: the
verification's actual work is ~300 ms against a 2,267 ms death, and a dying Excel and a working one
contend for the same disk. The change was reverted. It added a concurrency hazard to the mutation
path in exchange for a measured regression.

### What would actually move it

Not launching the second Excel at all. Verification exists to read the saved file back and confirm
what reached disk - and for `.xlsx` that file is a zip of XML, readable without Excel. That would
remove a launch *and* a death, ~2.6 s of a 5.2 s Apply, where every scheduling trick has now
returned less than nothing.

It is a real change in meaning, not just mechanism: reopening in Excel also proves Excel can still
open the file, which reading the bytes does not. That tradeoff deserves its own decision rather than
being smuggled in as an optimization. `.xlsm` macro verification is a separate question again.

### An invalid measurement, kept

Excel's teardown was first measured from PowerShell, which reported the process alive at the
10-second timeout in every trial. PowerShell holds its own references to the COM object, so Excel
could not exit: that measured the probe. It is recorded so the next person does not repeat it.

`ReadWorksheetRange` and `AuditWorkbookFlows` launch once and never verify, so neither pays any of
this. The most-requested operation is already at the floor.
