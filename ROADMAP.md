# Roadmap

Direction, not a feature promise. Nothing ships without its gate, and gates are
evidence - a test that passed, a field report, a measured need. The full
evidence trail lives in `docs/field-reports/` and the changelog.

## Delivered

All four original phases landed and were field-validated on the work computer on
2026-08-10:

1. **Reliability** (v0.2.0): supervised worker, hard deadline, truthful
   interruption, proven process cleanup. Field-confirmed with the user's own
   Excel instances live and untouched.
2. **Formula/exhibit depth** (v0.3.0): bounded gap repair and series extension.
3. **Macro editing** (v0.4.0, hardened v0.6.0): hash-preconditioned
   whole-procedure replacement on an isolated copy, optional run, dialog
   containment. Field-confirmed including execution.
4. **Read-only audit** (v0.7.0, completed v0.8.1): one workbook's worksheets,
   tables, defined names, queries, connections, model, pivots, and external
   links - names and shapes only, unchanged-proof in the receipt.
   Field-confirmed safe on a real business workbook.
5. **Bounded range read** (v0.9.0): the contents of one range, as values or as
   R1C1 formulas. The most-requested operation in the whole history.
6. **Overlapped verification launch** (v0.9.1): the verification Excel starts
   while the primary is still writing, measured at 3% of a macro Apply.
7. **Constant writes** (v0.10.0): values into named cells, read back in session
   and again from the reopened file. Formula text is still refused.
8. **Find/replace and creation** (v0.11.0): the three faithful-rebuild gaps.
   `FindReplace` lists matches on Plan and rewrites only constants on Apply;
   `Create` makes an empty workbook or adds an empty worksheet and never
   overwrites either. Together with the write, this is the first release that
   can compose new work from nothing rather than only edit what exists.
9. **Number formats** (v0.12.0): one format code across one bounded range, read
   back in session and again from the reopened file. Only the number format -
   the one part of `range_format` the write operation itself made necessary.
10. **Writability preflight** (v0.11.0): a read-only same-file target is refused
    before Excel is ever started, so it is a clean `Rejected` rather than the
    `Unknown` a failed save produced. On every one of the five write paths.
11. **Structure scan** (v0.14.0, extended v0.15.0, corrected v0.15.2): the first
    operation that never starts Excel - the workbook read as the ZIP of XML it
    is. Sheets, dimensions, formula/constant counts, constant islands (manual
    overrides inside calculated columns) by address, plus defined names, tables,
    and external links by file name. Measured end to end at 1.2 s against 4.6 s
    for the Excel-based audit on the same 80,000-cell workbook, with zero owned
    processes. Built for planning: what to read, what to fix, before any file
    opens.
12. **OneDrive and SharePoint identity** (v0.16.0): a workbook opened from a
    synced folder reports a service URL as its `FullName`, so exact-path
    matching refused every `UseOpen` against the storage the owner actually
    uses. The caller's path is now resolved through the sync client's own
    registry mapping and then compared exactly. **The registry lookup itself is
    still unproven** - see the open gates below.
13. **A field check that states its own coverage** (v0.16.0, completed v0.17.0):
    it once validated five of eleven operations and printed PASS. It now names
    any operation it did not exercise - and on its first run named a real one,
    `RepairExistingWorksheet`, which is now a step. All eleven are covered.

### Current test counts

Updated every release; `docs/FIELD-TASK.md` step 4 checks against these.

| Suite | Count |
|---|---|
| Core | 133 |
| Excel (fast) | 85 |
| McpServer (fast) | 18 |
| **Fast total** | **236** |
| Excel (OnDemand, real Excel) | 36 |
| McpServer (OnDemand) | 4 |
| **Full gate total** | **276** |

Measured against the original server on the work computer: 8.1x smaller tool
surface; 74% fewer input tokens, 73% fewer model requests, 84% fewer MCP calls,
53% less wall time across three client workflows - with ExcelTask's own Excel
execution 13% slower, the advantage being entirely the removal of model
coordination. One run per workflow; not yet a benchmark.

### What the 2026-08-11 review and field log corrected

Both were treated as claims to test rather than conclusions to adopt, and that
mattered: seven adversarial verifiers were run against the architecture review's
findings and **all seven returned partly-true**. The structural inventories held;
the stated consequences were routinely overstated. One headline finding was
refuted outright, and an entire category of five "seam violations" proved to be
five true observations with five false conclusions - acting on one would have
recreated a defect the code already records as having shipped once. None of that
was implemented. What survived is in the 0.16.0 changelog.

Two defects the review did **not** find were caught by running the shipped
binary against hostile fixtures: the scan reported every second defined name
under a summary that read as complete, and it carried external paths and stored
values into receipts from the one operation documented as reporting shape and
never contents. Both were in code shipped an hour earlier, and the test that
should have caught them asserted a single benign example rather than the
guarantee.

## From the 2026-08-11 field log, not yet addressed

A full day's real work produced four frictions. Two are fixed in 0.16.0 - the
SharePoint identity refusals and the audit truncating its worksheet list, the
latter already answered by `ScanWorkbookStructure`. These two are not.

- **`CopyExhibit` leaves the copied worksheet pointing at the source workbook.**
  Copying a sheet from a reference workbook produced 305 formula cells on the
  new tab carrying external references back to the source - and the follow-up
  repair fixed none of them, because Excel normalized the proposed internal
  references straight back into external ones. The workbook was left with an
  external link the owner did not ask for and could not remove, and the session
  ended with its state unknown. This is the flagship operation, and it is the
  largest product gap the log exposed. Worth measuring first: whether writing
  the formulas as R1C1 after the copy, or copying cell contents rather than the
  sheet, avoids the normalization.

- **`UseOpen` bound to a stray `Book1` instead of the named workbook.** Live
  verification reported an active workbook of `Book1` containing `iwe_getinst`
  and `Sheet1` - an add-in artifact, not the target. `RotWorkbookLocator.Find`
  and `HasExternalWorkbookAtPath` both bind the moniker and re-read `FullName`;
  `ContainsPath` matches on the ROT display name alone and is what
  `InspectCore` uses to set `TargetIsOpen`, which becomes "The exact target
  workbook is open." The word *exact* is carrying weight that code path does not
  supply. Fix is to bind and confirm, as its two siblings already do.

## Open field gates (small, when convenient)

- **One `UseOpen` against a OneDrive-backed workbook.** 0.16.0 resolves a
  reported SharePoint URL back to the local path through the sync client's
  registry mapping, and the mapping and comparison are pinned by tests - but a
  machine that syncs nothing registers no providers, so the lookup itself has
  never run. This is the only part of that change still unproven, and it is the
  highest-value unknown in the product. `docs/FIELD-TASK.md` step 3 settles it;
  the check now reports `syncRootsRegistered` and `syncPathsResolving` so half
  the answer arrives with step 1.
- Audit one workbook the owner *knows* contains Power Query and Data Model
  flows; owner confirms the reported categories. Closes phase 4's last gap.
- The repeated benchmark: one MCP catalog per client profile, three or more
  repetitions per workflow, median and spread, order alternated. Until then no
  percentage above is quoted as characteristic.
- **The schema budget is nearly spent: 16,012 bytes of 16,384 at 0.15.1.** The
  next description that states a rule will fail the pin test, and the strategy
  this project chose - no new tools, richer descriptions on the one that exists
  - spends exactly that budget. So the next addition has to buy its space from
  text already there rather than append. Two candidates to reclaim it: the
  eleven payload descriptions each re-state "all other payloads must be null"
  (~400 bytes of pure repetition, and the enum plus the `required` list already
  carry the rule), and several restate their kind name, which the property name
  already gives. Do that reclamation *before* the next feature, not during one,
  so a budget failure never arrives disguised as a feature bug.

## Candidates, ranked by measured demand

Five weeks of the owner's real history: 46 Excel sessions, 7,873 calls, ranked
by sessions rather than calls because one session made 2,515 of them. Full data
in `docs/field-reports/2026-08-10-demand/`.

**1. Finish the discovery layer.** Shipped in v0.8.1: the audit lists tables and
defined names alongside everything else.

**2. Reading cell values and formulas.** Shipped in v0.9.0. The design pass this
entry asked for concluded against the clever answer: returning *what differs from
a pattern* rather than the contents would have kept a promise nobody had asked
for, at the cost of not answering the question. The promise worth keeping is that
contents are never carried *incidentally* - and in a read they are the entire
request. So it returns them, under a hard bound rather than a refusal.

**3. Writing values and formulas.** `set-values` 19 sessions, `set-formulas` 15.
Now the top open item, and the two halves should be separated rather than decided
together:

- **Values shipped in v0.10.0**, within the existing stance rather than against
  it: a constant is exactly what the caller named, and a read-back proves it.
- **Formula text remains a genuine collision, and remains refused.** Inference
  plus verification is what makes an ExcelTask edit safe, and accepting composed
  formula text discards exactly that. `ExtendFormulaSeries` and
  `RepairExistingWorksheet` already serve the cases where the intended formula is
  derivable from evidence. What is still unknown is how much of the 15 sessions
  those two already cover - measurable rather than arguable, and not yet
  measured.

**4. Find and replace shipped in v0.11.0** (`range_edit`, 13 sessions). What
remains in demand order is below.

**The faithful-rebuild gaps are closed.** Workbook creation (9 sessions) and
blank worksheet creation (7) shipped in v0.11.0. One of the three remains, and it
is the weakest of them:

- **Discover open workbooks.** `file list` appeared in 35 sessions on the
  session-based server. ExcelTask's equivalent question - "what is open in Excel
  right now" - is answerable per-target through `AskIfOpen` but not as a survey.
  Demand here is largely an artifact of the other server's session model, which
  forced an open call before any work at all. Still unbuilt, and deliberately:
  measure it against real ExcelTask use before building to a number that may be
  entirely the other design's overhead.

**What is left, and what each one needs before it is built.** Nothing here is
blocked on effort; each is blocked on evidence, which is the standing rule.

- **`range_format`, 12 sessions - the number-format tenth shipped in v0.12.0, the
  rest is still gated.** `SetNumberFormat` was built on an argument rather than a
  measurement, and the argument is narrow enough to state exactly: the write
  operation *created* that gap, since a correct number that reads wrong is not
  finished work. Nothing else in `range_format` has that property.

  The demand data records the *tool*, not which of its operations were called, so
  "12 sessions used formatting" still says nothing about whether they set fonts,
  widths, borders, or conditional formats. Building all of it would be the
  dead-weight problem this project exists to avoid; building the wrong tenth of it
  is worse than building none. **Gate for anything beyond the number format:** an
  operation-level count from the session history, the same way
  `docs/field-reports/2026-08-10-demand/` produced the tool-level one.

- **`screenshot`, 13 sessions (11 of them `capture`).** Worth stating plainly
  rather than leaving on the list: most of this demand is *verification*, and
  ExcelTask already answers it by a different means - it reopens the saved file
  and reads back what it wrote, which is stronger evidence than a picture of a
  window. What a screenshot does that nothing here does is show a **person** the
  result. That is a real need and a different one, and it would put workbook
  contents into an image, which the receipt bounds cannot inspect. **Gate:** an
  observed case where reopen-verification was insufficient.

- **`table` beyond listing (11), Data Model and Power Query mutation (10 each).**
  Each is a large surface with a small measured slice. **Gate:** the same
  operation-level count.

**Not to be built.** `chart_config`, `pivottable_calc` and `worksheet_style` were
never called once in five weeks; `chart`, `slicer` and `table_column` appeared in
one session each. Roughly half of the original server's 230 operations show no
demand at all - which is most of the 8.1x schema ExcelTask does not carry.

**Open from earlier evidence, unranked by this data:**

- ~~**The four seconds inside a macro Apply that nobody has accounted for.**~~
  **Attributed, v0.14.0.** This entry demanded a measurement before anything was
  optimized; `EXCELTASK_TRACE` supplied one. A `SetNumberFormat` apply on a
  12-row sheet: 44 ms of Excel work, 398 ms session open, **2,814 ms closing
  owned Excel and proving it exited, 2,297 ms reopening to verify** - 92% of
  5,573 ms is teardown plus verification. The old guess in this entry (worker
  startup, MCP round trips, model coordination) was wrong for this operation.

  Neither cost is waste: proving the process exited is the product's central
  claim, and the reopen is what makes a receipt mean anything. So this becomes a
  narrower question rather than a closed one - **why does proving exit take 2.8
  seconds?** `WaitForExitOrTerminate` allows Excel 10 s before escalating, so the
  observed 2.8 s is Excel genuinely taking that long to die, not a timeout.
  Whether that is reducible without weakening the proof is untested. **Gate:**
  the same trace, run across the macro and copy paths, to confirm the split holds
  where the work is heavier before anything is tuned.

  This also retires the "macro session sharing" entry that used to sit here: the
  28.1s against 26.4s regression was blamed on Plan and Apply each opening their
  own Excel, but four launches is about 1.2 seconds. The other 27 were never
  Excel - they are worker startup, MCP round trips, and model coordination, and
  the only lever on those is fewer calls. See `docs/EXCEL-TUNING.md`.
- **Multi-workbook audit.** Follow external links through several workbooks into
  one dependency report.

## What the field survey changes

`docs/LANDSCAPE.md` compared ExcelTask against every spreadsheet MCP server the
survey could find. Most of it argues for holding course, and that is worth
stating before the parts that argue for change - a competitive survey that only
ever generates work has not been read critically.

**Confirmed, no action.** The incumbent's own README benchmark reports ~163K
tokens in MCP mode against ~59K in CLI mode, attributing the gap to "MCP sends
26 tool schemas to the LLM (~100K+ tokens)". That is this project's founding bet,
measured and published by the other side. Likewise the headless C# path (EPPlus,
ClosedXML) evaluates formulas in .NET but as partial re-implementations of Excel
semantics, which is the fidelity gap that keeps mutations on live COM. Neither
finding asks for a change; both retire a recurring doubt.

**1. Let the audit answer from the file when the workbook is closed.**
`ScanWorkbookStructure` measured 1.2 s against `AuditWorkbookFlows`'s 4.6 s on
the same workbook, because it starts no Excel. Much of what the audit reports -
worksheets, their dimensions, defined names, tables, external link targets - is
in the Open XML package too. The audit's *interface* would not change; only its
implementation, on the closed-workbook path.

  The gate this entry originally set - a part-by-part inventory of what the
  package can answer - **has been run.** A workbook was built carrying a table, a
  defined name, an external link, a Power Query and a VBA module, then opened as
  a ZIP and read part by part:

  | audit category | package part | verdict |
  |---|---|---|
  | `worksheet` | `xl/worksheets/*.xml` | plain XML - **answerable** |
  | `table` | `xl/tables/table1.xml` | plain XML - **answerable** |
  | `named-range` | `xl/workbook.xml` `<definedNames>` | plain XML - **answerable** |
  | `external-link` | `xl/externalLinks/*.xml` | plain XML - **answerable** |
  | `pivot`, `connection` | `xl/pivotTables/`, `xl/connections.xml` | plain XML when present; absent from the probe, so **untested** |
  | `query` | `customXml/item1.xml` | a `<DataMashup>` element wrapping **base64 of a nested ZIP** - reachable in principle, undocumented and brittle in practice |
  | `macro-component`, `macro-procedure` | `xl/vbaProject.bin` | **binary OLE compound file** - needs a real parser |
  | `model-table/-relationship/-measure` | binary model part | **not answerable** |

  So the honest shape of the win is narrower than "make the audit fast": four
  categories are cleanly answerable, and they happen to be the four a caller most
  often needs to plan with. Macros, the data model and Power Query are not, and
  chasing them means writing an OLE parser and depending on an undocumented
  nested-archive format to report on the two things - VBA and connections - where
  a wrong answer is most expensive.

  **Therefore the recommendation is not a hybrid audit.** A single operation that
  silently answered a subset while keeping the same summary would be a receipt
  that reads complete and is not - the worst outcome available, and this project
  has already shipped one defect from two sources of truth disagreeing. The
  better shape is what already exists: `ScanWorkbookStructure` stays a separate,
  honestly-named fast path, and gains the three cheap categories it does not yet
  report (tables, defined names, external links) so a caller can plan from one
  1.2-second call and reach for the audit only when it needs macros, queries or
  the model. **Gate:** none - this is additive to an operation that already
  exists and carries no risk of disagreeing with the audit, because it would stop
  short of every category the audit alone can answer.

**2. The formula refusal has a middle, and the field has mapped it.** Every
surveyed competitor accepts model-written formula text; ExcelTask refuses and
infers instead. The survey found the design space is not binary: SheetMind
constrains generation to a closed BNF grammar of seven operations, and
Microsoft's SheetBrain accepts model-written Python but runs it sandboxed. The
grammar option is the one compatible with this project's stance - a caller could
name a *shape* (sum this contiguous range into that cell) which the engine
composes and verifies, without any model-authored text ever reaching a cell.

  This remains the largest open product question, and the gate has not moved:
  `set-formulas` was 15 of 46 sessions, and how much of it `ExtendFormulaSeries`
  and `RepairExistingWorksheet` already cover is still unmeasured. **Gate:** that
  measurement first. A grammar built before it would be guessing at which shapes
  matter, which is the same error as building all of `range_format`.

**3. Verification now has outside grounding, and should be cited.** Nothing in
the survey has a counterpart to reopen-and-verify or proof-of-exit. Independently,
a 2026 study of verified tool calls found a postcondition-wrapped agent held 100%
task success under injected faults while an unverified baseline fell to 64%. That
is the closest academic parallel to what this server does on every Apply, and it
belongs in the design rationale rather than only in a survey appendix.

**4. Nothing here argues for more tools.** The description-quality study found
97.1% of 856 real-world tool descriptions carry a defect and 56% never state the
tool's purpose. ExcelTask's answer to that is one tool whose every rejection rule
is stated and A/B tested. The survey's one apparent contradiction - that richer
descriptions raised execution steps 67% - was checked at the source and measures
model invocations against wild-caught defective descriptions, not tool calls
against an already-rich schema. It is a warning about a road not yet taken
(adding examples and usage guidance), not evidence against the road taken.

## Standing rules

One tool. No CLI. No model selection anywhere. Schema bytes are budgeted, every
operation earns its own, and a rule the engine rejects on must be stated in the
schema - every failure in the interface study traced to one that was not.

Receipts stay bounded and truthful: a truncated report says so, and an uncertain
outcome is `Unknown`. Workbook contents appear in a receipt only when they are
the explicit request - the range a read asked for, the one procedure a macro
Plan named - and never incidentally. No receipt ever carries connection
strings, machine paths, or content nobody asked for. Model-written formula text
is never accepted; formulas are inferred from evidence in the sheet and
verified after reopening.
