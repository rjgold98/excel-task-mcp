# Changelog

## Unreleased

### Added - the first operation that never starts Excel

- **`ScanWorkbookStructure` maps a workbook by reading the file directly.** An
  .xlsx or .xlsm is physically a ZIP of XML, and everything structural is
  legible from it without Excel: every sheet's name and dimension, and per cell
  whether a formula or a constant put the value there. The trace had already
  measured that 92% of a small task's wall time is owned-Excel teardown and
  verification; a scan pays none of it.

  The reason it exists is planning - "algorithmically scan the workbook so the
  fixes can be planned before the file is ever opened." Its signature report is
  the **constant island**: a column that is overwhelmingly formulas holding a
  scattering of constants, which is the shape of a manual override sitting in a
  calculated column. On a 20,000-row spike fixture it named all 37 hardcoded
  overrides in a formula column by exact address - a fact no bounded 400-cell
  read could affordably discover, and the audit could not see at all.

  It reports shape, never contents: names, dimensions, counts and addresses -
  no cell values, no formula text. It handles Excel's shared-formula encoding
  (member cells whose f element is empty), resolves sheet parts through the
  relationships file rather than assuming their order, prohibits DTDs, and
  carries a five-million-cell budget so a hostile file exhausts a disposable
  worker rather than the machine. An encrypted workbook is an OLE container
  rather than a ZIP; the rejection says so and names the road still open - the
  Excel-based operations can open what the scan cannot.

  It still runs inside the supervised worker deliberately: a malformed file
  takes down a subprocess with a deadline, never the server. But its behaviour
  tests need no Excel and run in the fast tier in milliseconds - hand-authored
  OOXML for the parser's claims, plus one integration test against a workbook
  real Excel wrote, with the observer proving zero owned processes started.

  Schema budget rises 15 KB to 16 KB - the eleventh operation, and the only
  reason the budget is allowed to grow.

### Changed - measured against fresh assistants on plain-English prompts

Round 12 of the interface study stopped testing schema comprehension and tested
the product: a fresh assistant, a sentence a finance user would actually type,
and the real server driving real Excel. Four scenarios, scored by their own
friction reports and by a COM verifier that reopens every workbook afterwards.
Every task was correct in both rounds; what changed is what it cost.

| scenario | before | after |
|---|---:|---:|
| build a workbook from nothing | 4 calls | 3 |
| change one named figure | 4 calls | 2 |
| hardcode a calculated cell | - | 1 |
| recover from a mistyped filename | - | 2 |

- **A mistyped path now says so.** A doubled letter in a filename came back as
  "Workbook inspection could not be completed before execution" - infrastructure
  answering the most ordinary user error there is. Inspection carries an
  `InfeasibleReason` instead of throwing, and the rejection names the file it
  looked for, so the next attempt can be the right one.
- **Reads report `isFormula` per cell.** An assistant spent an entire Excel
  launch reading a range twice and diffing it to learn whether the cell it was
  about to overwrite held a formula. One extra array read per range now answers
  that, and the scenario went from four calls to two.
- **A write that replaces a formula names those cells, and says what each cell
  held before.** `FindReplace` has always refused to rewrite a formula while
  `WriteWorksheetValues` would silently destroy one. Overwriting a formula with
  a constant stays legal - hardcoding a figure is real finance work - but it can
  no longer happen behind a receipt that says only "wrote 1 constant", and the
  caller can now report "469,750.25 to 471,000" rather than "something changed".
- **`Create` with kind `Workbook` can name its starting sheet.** "Set me up a
  workbook with a Summary tab" was two operations and two Excel launches; it is
  one, and the result now contains exactly one sheet with the requested name
  rather than an orphan `Sheet1` beside it.
- **Two capabilities that already existed are now advertised.** Assistants asked
  for a way to see a range's current number format before changing it, and a way
  to find a label's address without a prior read. `SetNumberFormat` Plan did the
  first; `FindReplace` Plan does the second. Six words each. This is a different
  failure class from the one the study has chased for eleven rounds - not a rule
  the schema fails to state, but a capability it fails to mention.

All of it landed inside the existing 15 KB schema budget, which bit three times
and each time forced a cut or a tightening rather than a bigger budget.

### Added - development-time diagnostics

- **`EXCELTASK_TRACE` narrates every step to a file.** Temporary and built to be
  deleted: off unless the variable names a path, one file plus a handful of call
  sites, no schema bytes, no CLI flag, and the model never learns it exists.
  `docs/DIAGNOSTIC-TRACE.md` lists the four edits that remove it.

  It is safe to paste into a chat by construction, and every file says so in its
  own header: phases and durations, operation kind and policy, worksheet names
  and A1 ranges, workbook file names only, owned Excel process ids, statuses and
  checks. Never cell values, formulas, VBA source, connection strings, server
  names, or full paths. A phase with no matching end is where it hung, which is
  the one thing no receipt could ever say.

  It earned itself on its first run. A `SetNumberFormat` apply on a 12-row
  sheet: 44 ms of Excel work, 2,814 ms closing owned Excel and proving it
  exited, 2,297 ms reopening to verify - 92% teardown and verification. The
  roadmap's longest-standing open question recorded a ~4 s unaccounted gap and
  guessed worker startup and MCP round trips; on this operation it is neither.
  Both costs are load-bearing, so this is not yet an optimization - but the
  roadmap required a measurement before anything here was touched, and this is
  one.

- **`tools/Measure-WorkbookCorpus.ps1` audits a folder of real workbooks into
  one anonymized shape report.** Built to answer "what do my actual exhibits
  look like" with a count rather than a recollection, which is what the roadmap
  gates `range_format`, tables and Power Query mutation on. It calls the shipped
  read-only audit and emits pseudonyms, size bands, sheet counts, used-range
  dimensions and item counts by kind - never real names, values, paths, or
  connection details. The name map stays local and is discarded unless asked
  for.

## 0.13.0 - 2026-08-11

An architecture release: no new operations, one closed leak, three deep modules
where twelve hand-kept copies were, and two defects that only consolidation
could surface. Short-term churn accepted deliberately for long-term stability -
every change re-proven against real Excel, 231 tests green including the full
serial desktop-Excel gate.

### Fixed - a real leak in the product's central claim

- **The pre-launched verification Excel could be left running.** Verify marks
  itself consumed on entry, so once Excel had started for verification, an
  attach failure (Workbooks.Open throwing on the just-saved file) left a
  process nothing would ever shut down: the catch only closed a session that
  was still null, and Dispose skips its abandon branch once consumed. The
  catch now abandons the prepared instance, and the correct-but-orphaned
  `OpenForVerification` wrapper - built, documented, zero call sites - is
  deleted rather than left to mislead. `Abandon` itself now escalates the way
  `Close` always has: Quit, wait, terminate by identity, then report.
- **A staging promotion failure destroyed verified output silently.**
  `IOException` from the post-verification file move was in no operation's
  catch filter, so it escaped through the `finally` that deletes the verified
  staging file and reached the caller as an Unknown with no receipt. Caught
  and named once, in the pipeline.
- **Rejection summaries reached the model cut mid-word.** The MCP tool's
  96-character cap silently truncated engine-authored summaries - including
  the four-part macro policy rejection written so one resubmission can fix
  everything, a rule the interface study paid for, which arrived missing its
  fourth requirement and the instruction. The engine's own 256-character test
  passed the whole time because it read the engine directly. The model-facing
  cap is now 256 everywhere.
- **An oversized macro source is omitted at every seam, never truncated.** The
  worker protocol used to truncate one to exactly the limit, which then
  arrived downstream measuring within it and passed as complete - defeating
  the two layers that had decided partial VBA is more dangerous than none.

### Changed - three deep modules

- **`ExecuteMutation` owns the mutate-save-verify choreography.** Twelve steps
  in a fixed order were restated in every operation (91% of the write path was
  scaffolding shared with another operation) and had drifted six ways - one
  path skipped the writability preflight, catch filters disagreed, one verify
  block was inverted. Write, FindReplace, NumberFormat and worksheet creation
  now supply only their own middle; the tail is not their code, so it cannot
  run out of order. Macro and the shared formula path stay bespoke
  deliberately - their extra machinery (hash preconditions, dialog
  containment, two-phase revalidation) would have made the pipeline interface
  as wide as the bodies it replaced.
- **`ComAccess.IsComFailure` is the one definition of "a COM call failed",**
  replacing 29 catch sites in five mutually inconsistent forms. The narrowest
  omitted the exception `ComAccess` itself throws, so a preflight fault on
  four operations was attributed to the wrong phase in the receipt.
- **`ReceiptBounds` is the one implementation of receipt bounding,** replacing
  three modules' private opinions: four independently declared item caps, a
  dead 256-character cap behind the protocol's 128, and the truncate-vs-omit
  disagreement above. The three seams still bound - each has a reason to
  distrust the layers behind it - but none decides what bounding means.

### Tested

- **The choreography is asserted at the observer seam.** The runtime always
  reported every phase, owned process, and staging path; every test discarded
  it through the null observer. A recording adapter now asserts the ordering
  invariants - cleanup proven before verification opens the file, staging
  announced while the supervisor can act on it, both owned processes announced
  - against real Excel, before and after the rewrite they guard.
- **The engine's inspection-driven decision surface is reached.** Every Core
  test ran against a closed target, leaving both confirmations, both
  rejections, the inspection-failure receipt and the retry coercion untested -
  including the arm whose wrong message was the v0.11.0 bug. Eight tests hold
  it now, and `WorkbookExecutionOutcome` finally states that a runtime's
  CanRetry/RetryReason are advisory: the retry policy belongs to the engine.

### Not done, deliberately

- The operation union's registration shape stays hand-rolled. The schema it
  generates is a measured product surface (15 KB budget, wording validated at
  p = 0.0012), the worker protocol already needs zero edits per operation, and
  the arity checklist test catches the one silent failure that actually
  shipped. Revisit only if a second class of silent miss appears.

## 0.12.0 - 2026-08-11

### Added

- **`SetNumberFormat` sets one number format code across one bounded range.**
  This closes a gap the server made for itself: `WriteWorksheetValues` could put
  1000.5 into a cell and nothing could make it read as `1,000.50` or
  `(1,000.50)`, so a correct number could still be presented wrongly - which for
  a financial exhibit is the difference between usable and not.

  It is deliberately *only* the number format. Fonts, fills, borders, widths and
  conditional formats are not here. The measured demand names the `range_format`
  tool, not which of its operations were used, and building all of it on a
  tool-level count is how a server ends up with 230 operations and no idea which
  matter. The operation is named for what it sets so a caller does not spend a
  round trip discovering the rest is missing, and its description says outright
  what is absent.

  Unlike every other mutation here it is a single COM assignment no matter how
  many cells it covers, since Excel applies a format to a whole range at once.
  What it does need is the read-back: an unrecognized code can be kept verbatim,
  coerced, or rejected outright, and all three look identical from the caller's
  side of the assignment. The code is read back in session and again from the
  reopened file, and a mismatch is reported rather than called success.

  Format codes are **not trimmed**. Leading and trailing spaces are meaningful -
  `_)` and a trailing space are what make parenthesised negatives line up under
  positives in a column - so trimming would silently apply a different format
  from the one asked for.

  Plan reports the range's current format before changing anything, because a
  format is destructive in a way a value write is not: it replaces whatever was
  there and the old code is not recoverable from the sheet afterwards.

  Bounded at 10,000 cells, the same bound find/replace and repair use, and at
  Excel's own 255-character ceiling on a format code.

- Schema is 15,029 bytes against the 15 KB budget set in 0.11.0, so the tenth
  operation fits without raising it.

### Measured

- **A model can write these format codes.** The operation's real risk was never
  the interface - operation selection has been perfect in 484 of 484 decisions
  across the study - but the format code itself, a small hostile language where
  `_)` is padding and `[Red]` is a colour section. Round 11 put four requests
  phrased the way a finance user phrases them, never naming a code, through
  three runs, and applied every proposed code to real Excel: **12 of 12 accepted,
  12 of 12 round-tripped verbatim, all three runs byte-identical**, including the
  `_)` padding that is the usual hand-written mistake. Details in
  `docs/INTERFACE-AB-STUDY.md`. The read-back stays: zero failures in twelve
  tries is not evidence the guard is unnecessary.

## 0.11.0 - 2026-08-10

### Added - the three faithful-rebuild gaps

Until this release ExcelTask could only change a workbook that already existed,
with sheets that already existed. Two operations from the demand data closed
that, and together with the existing write they are the first time the server
can compose new work from nothing: create a workbook, add a sheet, write values
into it, change them. That sequence is now a test, end to end through the real
MCP boundary.

- **`FindReplace` finds cells whose text matches, and on Apply rewrites the
  constants among them.** `range_edit` appeared in 13 of 46 measured sessions.
  Plan lists the matches and changes nothing, because replace-across-a-sheet is
  the operation most likely to be regretted and the caller should see the cells
  before authorizing them. A cell whose text comes from a formula is reported as
  a match and never rewritten, the same refusal every other operation makes.

  Excel's own `Find` and `Replace` are deliberately not used, for three separate
  reasons. `Find` treats `*`, `?` and `~` as wildcards, so a search for `Q1*`
  would quietly match text nobody asked about. Its omitted arguments inherit
  from the last search performed anywhere in the application - including one a
  person ran by hand - so the same request can mean different things on
  different machines. And `Replace` reports how many cells it changed and never
  which, so a receipt built on it could not name what moved. Reading the range
  once and matching in code costs two COM calls regardless of how many cells
  match, against three per match for a `Find`/`FindNext` walk, and it makes the
  matching rule something this repository states and tests rather than something
  Excel decides.

  Every replacement is composed and checked before any is written. A partial
  replacement can turn a label into a formula - `x=1` losing its `x` leaves
  `=1`, which Excel stores as a formula - so the whole request is refused before
  a single cell changes rather than half a sheet being rewritten and then
  stopped.

- **`Create` makes an empty workbook, or adds an empty worksheet.** `file
  create` appeared in 9 sessions and `worksheet create` in 7; both were hard
  walls, since every other operation requires a target that already exists.
  Creation never overwrites: an existing file or an existing sheet name is
  refused outright, and there is no confirmation that unlocks it, because a
  caller who wants to replace a workbook should say so with a save. A new sheet
  is added after the last one, so it never displaces the sheet the workbook
  opens on. Creation is deliberately empty - no template, no seeded content -
  because the operations that fill a sheet already exist and each verifies its
  own work.

  Creating a workbook is the one operation whose target must *not* exist, so
  inspection now carries `TargetMustExist` and this is the only request that
  sets it false. It also never asks to confirm an overwrite, because it is
  refused outright if anything is there - asking would authorize nothing and
  teach the caller to set the flag reflexively.

### Fixed - all three found by the A/B run, before any user hit them

- **Two operations shipped unreachable.** The operation union counts its
  payloads by hand, and `FindReplace` and `Create` were absent from that count,
  so a request carrying only one of them failed the arity check and never
  reached its own validation. Both looked complete in every other respect. The
  count is now a list rather than a chain of additions, and a test builds one
  request per `ExcelOperationKind` and asserts none fails on arity, so the next
  operation cannot ship the same way.
- **`NeedsConfirmation` described the wrong thing.** A same-file Apply missing
  its overwrite flag came back saying "The requested copy output already exists"
  - a sentence about a file the request never mentioned. The requirement's own
  prompt was correct all along; only the summary lied, and the summary is the
  line a caller reads first. It is now built from the requirements themselves.
- **`Create` enforced a weaker rule than it stated.** The schema and the
  rejection message both said Isolated; the check only rejected `UseOpen`, so
  `AskIfOpen` reached a confirmation whose two answers are an option creation
  refuses and an option inspection then rejects. Now Isolated exactly, matching
  macro editing.
- **The real-Excel MCP tests could report leaks that never happened.** They
  snapshotted the process table once, the instant the test body returned, while
  the Excel project's equivalent has settled and retried since 0.9.0. Excel
  exits asynchronously, and running both assemblies at once - which
  `dotnet test` on the solution does and `scripts/Test-Mvp.ps1` deliberately
  does not - made one assembly's fixtures look like the other's leaks. The
  assertion no longer depends on the caller remembering to serialize.

### Interface

- Every rule the engine rejects on for the two new operations is stated in the
  schema. Measured this release, not assumed: on the creation tasks the arm
  without those clauses scored **0 of 6** against 5 of 6 with them, every failure
  the identical binding rejection (Fisher exact two-sided p = 0.015). On
  find/replace the 10,000-cell bound accounted for all three failures of a
  546,000-cell used range. The `* and ? are literal` clause is recorded as
  **unvalidated** - both arms got it right without being told.
- **The overwrite gate's wording is the study's most persistent defect and is
  now rewritten.** Six of seven remaining failures were a same-file Apply sent
  without `overwriteConfirmed` - round 1's headline failure recurring on a rule
  that *was* stated. "An existing save destination" does not read as *the
  workbook you are editing*. It now names the rule per save mode, and
  `auditWorkbookFlows` now says to supply `{}` rather than only that it takes no
  options. With both, the surface scored **18/18 clean on every task and every
  dimension**, against 14/18 before. Directional rather than proven (p = 0.104,
  and all four failures came from one run), applied on the same mechanical
  grounds as the round 5 live-binding clause: most frequent failure in the whole
  study, 180 bytes, nothing regressed.
- The schema budget rises from 13 KB to 15 KB, which is what two operations with
  29 sessions of measured demand between them buy.

## 0.10.1 - 2026-08-10

- Closed the interface study's follow-on audit: the eight rules the engine
  enforced but the schema never stated are now all in the descriptions - the
  VBA source bounds (8,192 characters, 200 lines), the blocking constructs a
  run refuses, auto-entry procedures, the `.xlsx`/`.xlsm` path rule, the copy
  output's differ-from-target and same-extension rules, and the Isolated+Same
  rejection while the target is open. The fix class is measured: stating a rule
  the server rejects on eliminated every related failure across two models and
  cut end-to-end calls 32%.
- The standing rules now match the shipped product: contents appear in a receipt
  only when they are the explicit request, never incidentally - the earlier
  absolute wording predated the bounded read and was already untrue of macro
  Plan.
- Roadmap carries the faithful-rebuild gaps the demand data surfaced: workbook
  creation, blank worksheet creation, and open-workbook discovery.

## 0.10.0 - 2026-08-10

### Added - the second-most-requested operation

- **`WriteWorksheetValues` writes constants into named cells.** `set-values`
  appeared in 19 of 46 measured sessions and was a standing refusal. The refusal
  is now split, because the two halves are not the same risk:

  - A **constant** cannot be plausibly and silently wrong. It is exactly what the
    caller named, and reading it back proves it character for character.
  - **Formula text** still is, and is still refused. `ExtendFormulaSeries` and
    `RepairExistingWorksheet` infer formulas from evidence already in the sheet,
    which is what makes an ExcelTask edit safe; accepting composed formula text
    would discard exactly that. A value starting with `=` is rejected, and the
    rejection names the two operations to use instead - storing a formula as a
    text label would be the worst outcome available.

  Numbers and TRUE/FALSE are converted rather than stored as text: a model that
  sends "1000" means the number, and a text "1000" leaves a cell that looks
  right and breaks every SUM above it. Dates are deliberately not parsed - "3-4"
  is March 4th to Excel and a label to a person, and nothing in the request says
  which, so it is written as text where it can be seen and corrected.

  Bounded at 200 cells that must fit inside a 400-cell region, so a write cannot
  quietly scatter across a model. Every cell is read back in-session and again
  from the reopened file; the same cell twice in one request is rejected, since
  such a request does not say what it wants.

## 0.9.1 - 2026-08-10

### Changed - speed

- **The verification Excel now starts before the work instead of after it.** A
  mutating Apply launches Excel twice: once to make the change, once more to
  reopen the saved file and prove it. The second launch cannot be deleted -
  verifying in the process that did the writing would be verifying against the
  memory that produced it - so it is started early and its launch overlaps the
  writes and the save. Every property is unchanged: a separate process, freshly
  launched, opening the file only after the primary closed and released its lock.

  Measured A/B on one successful macro Apply, 5 trials each: **5,398 ms to
  5,244 ms**, about 3%, with non-overlapping ranges. Less than the 326 ms a
  launch costs alone, because two Excels starting at once contend for the same
  disk and CPU.

### Fixed

- **The deadline watchdog would have killed the wrong Excel.** It tracked a
  single owned process identity, which was correct while only one owned Excel
  could exist at a time. With the verification instance alive alongside the
  primary, the later registration silently replaced the earlier one - so on a
  deadline the watchdog would have ended the idle verification instance and left
  the one holding the user's workbook mid-write running. It now tracks every
  owned process and terminates all of them.

### Measured

- **A macro Apply is 5,244 ms; the COM it performs is 1,035 ms.** Roughly four
  seconds happens inside the runtime and has never been attributed. That is now
  the largest known unexplained cost in the product, and the next measurement is
  aimed at it rather than at another three hundred milliseconds. An attempt to
  attribute it to Excel's process teardown is recorded as invalid: measuring that
  from PowerShell measures PowerShell's own COM references. See
  `docs/EXCEL-TUNING.md`.

## 0.9.0 - 2026-08-10

### Added - the most-requested operation

- **`ReadWorksheetRange` returns the contents of one bounded range**, as
  displayed values or as R1C1 formulas. Reading cell values was the single
  largest gap the demand data showed: 31 of 46 measured sessions called for it,
  and every one of them had to go to a different server to get it.

  This is deliberately the one operation that returns workbook contents. Every
  other receipt withholds them, because there they would be incidental - nobody
  asked an audit for cell values, so carrying them would be leakage. Here the
  contents are the entire request, and refusing would just send the caller
  elsewhere, which is what real use showed happening. The protection is a hard
  bound rather than a refusal: at most 400 cells, capped cell text, blanks
  omitted, and a range larger than the cap rejected rather than silently
  truncated to a partial answer that would read as a complete one.

  The whole range is fetched in a single array. Blank cells cost nothing.

### Fixed

- **A read would have demanded permission to overwrite the file it only reads.**
  Save mode `Same` requires `overwriteConfirmed` on Apply; the audit was exempted
  when it shipped and the range read inherited the rule. Beyond being nonsense,
  a caller taught to set `overwriteConfirmed` reflexively to get a read through
  carries it into the next call, which does write. Both read-only operations are
  now exempt.

- **A full-size read would have been lost, not truncated.** The worker frame
  budget was 16 KB, set when every receipt was metadata. 400 cells of capped text
  do not fit, and a frame over budget is replaced wholesale with a fatal code -
  so the largest reads, the ones a real model produces, would have failed
  entirely while small ones passed. The budget is now 64 KB, still far below the
  MCP response bound that decides what the caller actually sees, and a test
  serializes the largest legitimate result and measures it.

- **The full local gate could report a leaked Excel process that was not there,
  and could pass without having tested for one.** `dotnet test` on the solution
  runs test assemblies in parallel; two of them drive real desktop Excel and then
  assert no Excel was left behind, so each saw the other's. `tools/gate.ps1` runs
  one project at a time and is now the documented gate.

### Changed - measured, and mostly a decision not to change anything

- Recorded in `docs/EXCEL-TUNING.md`, from `tools/excel-config-probe.ps1` and
  `tools/excel-calc-probe.ps1`. Launching Excel is 274-482 ms; writing 2,000
  formulas is 58 ms. Starting Excel is the entire cost, and everything done
  inside it is rounding error beside that.

- **Manual calculation was measured and rejected: it is slower here.** 103 ms
  against 76 ms. It exists to stop a recalculation per write, and ExcelTask does
  not write per cell - repairs are grouped by identical R1C1 formula, so 2,000
  cells cost 43 calls and there is nothing left to suppress. A per-cell loop
  shows the opposite, which is exactly why it must not be the thing measured.
  `ScreenUpdating`, `PrintCommunication`, and `DisplayStatusBar` all measured
  within noise: they are advice for driving a visible Excel, and this one has no
  window.

- A range read launches Excel once, like the audit. The most-requested operation
  is already at the floor. A mutating Apply still launches twice, and the reason
  that second launch cannot simply be deleted - along with the version worth
  building instead - is written down rather than left as a to-do.

## 0.8.1 - 2026-08-10

### Fixed

- **A macro compile error could take the whole server process down.** The dialog
  sentry ends the owned Excel to break a call blocked behind a compile-error
  dialog, and the code then closed the session normally - sending Quit and Close
  to a server that no longer existed and releasing its proxies. Releasing a proxy
  whose server has been terminated can fault hard enough to kill the host process,
  reliably so once the VBA editor had been materialized by an earlier macro run.
  Those references are now abandoned rather than released, and the receipt reports
  whether the process genuinely exited instead of assuming it.

### Changed - speed

- **Formula verification reads the repaired block in one array instead of one COM
  call per cell.** Measured on this machine: 3,000 individual cell reads cost
  4,864 ms, the same cells read as a single range cost 13 ms. Verification was
  the last chatty path; writes were already batched.

### Added - discovery, from the demand data

- The audit lists **tables** and **defined names**, which appeared in 11 and 12
  of the owner's 46 real Excel sessions. With worksheets and macro procedures in
  0.8.0, the discovery layer now covers what real work looks up before acting.

### Measured, and deliberately not changed

Excel's application settings were tested against a workload with a real
dependency chain: `ScreenUpdating`, `Calculation = manual`, `PrintCommunication`
and `AskToUpdateLinks` all landed within noise of the 425 ms baseline. They do
not help because the architecture already avoids what they protect against -
`Visible = false` subsumes screen updating, and writes are grouped rather than
per-cell so calculation thrash never happens. `AskToUpdateLinks` is also
unnecessary for correctness: opening with `UpdateLinks:=0` was verified to return
the cached value with no prompt for both a stale link and a missing one.

## 0.8.0 - 2026-08-10

Two product defects found by driving the built server against real Excel, the
discovery gap that sent a real task to a different server, and an architectural
pass that removes the duplication behind three of this month's defects.

### Fixed - found end to end against real Excel

- **A repair could be rejected with no reason depending on where it sat in the
  sheet.** Formula writes were batched into fixed groups of 64 and joined into a
  single `Range("B10,B20,...")` address, but Excel rejects that argument beyond
  255 characters - so identical work succeeded near row 1 and failed near row
  2500. Batches are now bounded by joined address length.
- **A chunked repair could silently skip the last row of each chunk.** Inference
  reads the neighbour on each side of a blank, and a gap on a chunk's edge had
  its neighbour outside the range that was read - so the cell was skipped with a
  `Completed` receipt and no warning. Evidence is now read one cell beyond the
  request on each side while writes stay inside it. The natural chunking at round
  thousands was the dangerous one.
- **A catch-all failure now names its phase and fault.** Any COM fault previously
  surfaced as "execution was rejected before changes were attempted" with the
  reason discarded.

### Added

- **The audit lists what a caller needs before it can act**: macro components and
  procedures, and every worksheet with its visibility and used range. Both close
  discovery gaps that forced real work to a different server - `EditMacroProcedure`
  demands names ExcelTask could not supply, and every formula operation demands a
  worksheet name.
- Pivot sources are classified rather than quoted, because `PivotCache.SourceData`
  is a connection string for externally backed pivots.

### Changed - measured interface work

- Guardrails the engine enforced but never stated are now in the schema: the range
  and cell caps, the `UseOpen`+`Copy` conflict, audit-never-writes, and the
  already-open exception. Proven across two models (p = 0.0012) and end to end
  (32% fewer calls, p = 0.0032); every failure in the study traced to a rule the
  server rejected on but the description never mentioned.
- Macro policy states every unmet requirement in one rejection instead of teaching
  one rule per round trip.

### Internal

- One `ComAccess` module holds the late-bound rules that nine helper sets used to
  restate. `Item` is the only member that tries both bindings, since a retried
  write could land somewhere nobody asked for. This is the trap behind three
  defects this month.
- `CloseAndProve` replaces fourteen hand-written copies of close, prove exit, and
  map failure to `Unknown` - previously nine different wordings of one check.
- The field check's leak figure and the test fixtures now stand on the product's
  own process identity rather than raw process ids, which are recyclable.
- Leak assertions settle before failing and exclude harness-owned processes, so a
  test cannot blame the product for Excel exiting asynchronously.

## 0.7.0 - 2026-08-10

- Added the fifth operation, `AuditWorkbookFlows`: a read-only report of how one
  workbook's data flows fit together - Power Query queries and where each loads,
  connections, Data Model tables, relationships and measures, PivotTables and
  their sources, and links to other workbooks. It returns names and shapes only:
  no cell values, no M formulas, and no connection strings, because those carry
  servers and credentials. External links are reported as file names, never
  paths.
- Read-only is proven, not promised: the session opens read-only, a save
  destination is refused at validation, and the receipt compares the file's size
  and timestamp before and after so it can state the workbook was not changed.
- Every collection entry is resolved through an accessor that tolerates both the
  property and method bindings of `Item`, the difference that produced two
  separate defects earlier in the day.
- The field check now exercises the audit as its fourth operation, and its
  fixtures carry an external link and a Power Query for it to find.
- The schema budget test grew from 8 KB to 9 KB for the fifth operation. The
  advertised surface is 8,182 bytes - still 7.1x smaller than the original
  server's 58,324.

## 0.6.4 - 2026-08-10

- Made the macro rules visible in the schema instead of only in the rejection.
  Field measurement showed a caller losing two round trips to rules it could not
  see: it sent Apply-only fields on a Plan, and used the default binding where
  only `Isolated` is permitted - which the tool description encouraged by saying
  "Start with AskIfOpen" without exception. A protocol test now holds both rules
  in the schema.
- Recorded the six-session client comparison under
  `docs/field-reports/2026-08-10-comparison/CLIENT-SESSIONS.md`: 74% fewer input
  tokens, 73% fewer model requests, 84% fewer MCP calls, 53% less wall time, all
  workbooks correct after reopening. Two qualifications are recorded with it -
  ExcelTask's own Excel execution was 13% slower, and both tool catalogs were
  registered during those sessions, so they measure orchestration rather than
  schema loading.

## 0.6.3 - 2026-08-10

- Added `--make-fixture <directory>`, which writes a few disposable workbooks to
  try the server against and prints ready-made prompts. Every operation needs an
  existing workbook, and a managed computer usually cannot script Excel to produce
  one, so the first real test otherwise depended on building a workbook by hand.
- Added `docs/FIRST-TEST.md`: the shortest path from a downloaded release to
  watching a client actually change a workbook.

## 0.6.2 - 2026-08-10

- Released every COM interface the environment probe obtains. A retained one
  keeps the probe's Excel alive past `Quit`, and the next activation then meets a
  half-dead instance - the likely shape of the `CO_E_SERVER_EXEC_FAILURE` that
  stopped 0.5.0 on the work computer.
- Moved the documented install location and `.vscode/mcp.json` out of `C:\Tools`
  and under the home directory. The work computer refused to let the field agent
  write to `C:\Tools`, which is ordinary on a managed machine.
- Recorded the first work-computer comparison under
  `docs/field-reports/2026-08-10-comparison/`: ExcelTask advertises 1 tool and
  7,164 bytes against the original's 25 tools and 58,324 bytes, a ratio of 8.1x.

## 0.6.1 - 2026-08-10

- `--field-check` now prints and writes a compact digest: about ten dense lines
  carrying the Excel build, VBA trust values, both servers' tool-surface sizes,
  each operation's status and elapsed time, and the leak count. A managed computer
  often cannot move a file off itself, so the numbers that decide what to do next
  have to be short enough to retype or photograph.
- Fixed the environment probe overwriting a good Excel version with an error
  message. Enumerating COM add-ins can fail on its own, and that failure was
  discarding the version alongside it. The failure is now recorded separately.
- Fixed COM add-in enumeration. Office exposes `COMAddIns.Item` as a method rather
  than a parameterized property, so binding it as a property lost the whole list -
  the same trap that broke the VBA object model in 0.4.0.

## 0.6.0 - 2026-08-10

- A VBA modal dialog no longer stalls a macro run. `RunAfterEdit` now calls the
  procedure through a generated `On Error` wrapper, so a run-time error comes back
  as a named VBA error and number instead of opening a dialog nobody can answer.
  The edit itself is still saved and verified; the outcome is `Partial`.
- A replacement containing `MsgBox`, `InputBox`, `GetOpenFilename`,
  `GetSaveAsFilename`, `FileDialog`, or `Stop` is refused before Excel opens, but
  only when the request also asks to run it. Editing such a procedure is unaffected.
- Dialogs that a wrapper cannot catch - a compile error, or a message box inside a
  procedure the replacement calls - are answered by a sentry watching the owned
  Excel process from outside the blocked call. It acts only on a process ExcelTask
  created and re-verifies that identity by id, start time, and image path before
  every action, so a workbook bound with `UseOpen` is never touched.
- A compile error is distinguished from a message box by window ownership rather
  than by controls. The two are otherwise identical - `MsgBox "x", vbInformation +
  vbMsgBoxHelpButton` has the same control identifiers as a compile error - and
  only the compile error path may end the Excel process, so the difference decides
  whether a user's own dialog could trigger a termination. It cannot.
- Message boxes are recognized by button identity, not control count. The previous
  count rule matched only the bare form, so `MsgBox "x", vbInformation` was ignored
  and the run blocked until the deadline. Dialogs offering a real choice - Yes/No,
  OK/Cancel, Retry/Cancel - are still left alone, because on those identifier 2 is
  Cancel rather than OK and answering would choose for the user.
- The refusal screen no longer mistakes a method call for the `Stop` statement.
  A word boundary alone matched the `Stop` in `timer.Stop`, so ordinary macros
  that never wait for anyone were refused.
- Compile errors end the owned instance rather than clicking through, because
  measurement showed answering that dialog leaves VBA in break mode with the
  blocked call never returning. Nothing has been written at that point, so the
  outcome is a plain `Rejected` carrying the compiler's own words.

## 0.5.0 - 2026-08-10

- Moved the work-computer field check into the released executable as
  `--field-check`, replacing `scripts/Invoke-FieldCheck.ps1`. Managed computers
  commonly run PowerShell in Constrained Language Mode, which forbids the COM and
  reflection the script needed; compiled code is unaffected. It is not an MCP tool
  and not a public CLI - the model-facing surface is still exactly one tool.
- The tool-surface comparison now measures the raw `tools/list` wire response over
  JSON-RPC rather than a re-serialization, so the reported bytes are the server's
  own. Use `--compare` and `--compare-arg` to measure another MCP server.
- Field-check fixture processes are tracked and excluded from the leak figures, so
  a stranded-Excel count can never be the harness mistaken for the product.
- Added `.vscode/mcp.json` so clients that sync a repository's MCP configuration
  pick up the `excel-task` server automatically.
- Added `docs/AGENT-BRIDGE.md`, defining how the lead and field agents divide
  work, task each other through issues, and report with evidence.

## 0.4.1 - 2026-08-10

- Added `scripts/Invoke-FieldCheck.ps1` and `docs/FIELD-CHECK.md` so the
  work-computer gates the roadmap has always required can actually be run. It
  reports the environment, measures the advertised tool surface against another
  MCP server, and times each operation against disposable workbooks. It changes
  no Excel or security setting.
- Fixed the test harness for per-user .NET installs. The MCP transports launch
  the server with environment inheritance disabled, so the framework-dependent
  apphost could not see `DOTNET_ROOT` and failed with `0x80008083` on machines
  with no machine-wide .NET, which is how the first work-computer run failed. The
  child environment now carries the runtime root that hosts the tests. Found by
  the 2026-08-10 work-computer validation run.
- Recorded the first work-computer validation: the complete suite passed 146/146
  there, including desktop Excel integration, macro editing with execution, and
  process cleanup. VBA project access is permitted on that machine.

## 0.4.0 - 2026-08-10

- Added bounded macro editing to the existing operation union: `Plan` returns one
  selected standard-module procedure with its hash, and `Apply` replaces exactly
  that procedure only while the hash still matches, then saves, reopens, and
  verifies an `.xlsm` copy. Signed and locked VBA projects are refused, automatic-entry
  procedures cannot be edited, and Trust Center settings are never changed.
- Fixed VBA object-model binding. The extensibility model exposes `VBComponents.Item`
  as a method and `CodeModule.ProcStartLine`/`ProcCountLines`/`Lines` as parameterized
  properties, which is the opposite of Excel's own collections; the previous binding
  raised `DISP_E_MEMBERNOTFOUND` and made every macro operation fail preflight.
- Fixed optional post-edit macro execution. Owned Excel force-disables macros, so
  `Application.Run` could never succeed; macros are now enabled only when a request
  explicitly asks to run the edited procedure, with workbook events still suppressed.
- Replaced the single generic macro-access failure with distinct reasons for a Trust
  Center block, a locked project, a missing component, a non-standard module, a missing
  procedure, and an oversized procedure.
- Accepted array parameters such as `values() As Variant` in procedure signatures.
- Stopped re-checking plan source against the replacement grammar, which silently
  returned a successful plan with no source for valid but unparseable procedures.
  Oversized source is still omitted rather than truncated.
- Verified through the real one-tool MCP path against desktop Excel, including
  plan, hash-guarded replacement, post-edit run, save/reopen, and process cleanup.
- Split the Excel runtime into one file per type plus focused partials for the
  formula and session areas. This is a move-only change: the set of code lines is
  unchanged, and the same tests pass before and after.
- Added an explicit owned-process cleanup assertion to the desktop macro test,
  because editing VBA materializes the VBE and that is the most likely way this
  workflow could strand an Excel process.

## 0.3.0 - 2026-08-09

- Added one closed operation union for copy-exhibit, existing-sheet formula
  repair, and stable right/down formula-series extension without adding a tool.
- Added bounded pre-mutation formula planning and revalidation, exact
  save/reopen verification, and a final isolated same-file live-binding check.
- Preserved all valid range changes and terminal verification checks across the
  private-worker and MCP receipt boundaries.
- Fixed the acceptance script to propagate failed test exit codes.
- Verified six serial desktop-Excel workflows and the real one-tool MCP path
  from an empty Excel baseline with no process left afterward.

## 0.2.0 - 2026-08-09

- Moved every Excel operation into a short-lived private worker so blocked COM
  cannot stall the MCP host indefinitely.
- Added bounded worker protocol, hard deadlines, worker-owned Excel recovery,
  exact staging cleanup, and non-retryable `Unknown` outcomes after uncertain
  dispatch.
- Verified the real one-tool MCP path through save, reopen, file-lock release,
  and owned-process cleanup.

## 0.1.0 - 2026-08-08

- First formula/exhibit MVP: copy a named worksheet, repair only unambiguous
  blank formulas, save, reopen, and verify.
- One `excel_task` MCP tool with compact structured receipts.
- Safety baseline: exact open-workbook handling, explicit overwrite approval,
  owned-process cleanup, staged saves, and truthful uncertain outcomes.
