# Changelog

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
