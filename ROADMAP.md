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

Measured against the original server on the work computer: 8.1x smaller tool
surface; 74% fewer input tokens, 73% fewer model requests, 84% fewer MCP calls,
53% less wall time across three client workflows - with ExcelTask's own Excel
execution 13% slower, the advantage being entirely the removal of model
coordination. One run per workflow; not yet a benchmark.

## Open field gates (small, when convenient)

- Audit one workbook the owner *knows* contains Power Query and Data Model
  flows; owner confirms the reported categories. Closes phase 4's last gap.
- The repeated benchmark: one MCP catalog per client profile, three or more
  repetitions per workflow, median and spread, order alternated. Until then no
  percentage above is quoted as characteristic.

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

- **The four seconds inside a macro Apply that nobody has accounted for.** One
  Apply is 5,244 ms end to end; the COM it performs is 1,035 ms. That gap is now
  the largest known cost in the product, and it is twelve times the launch
  overlap just shipped. Attributing it needs timing inside the runtime, where the
  phase observer already sits - and nothing here should be optimized before that
  exists. An attempt to blame Excel's process teardown was confounded by the
  probe itself and is recorded as invalid rather than quietly dropped.

  This also retires the "macro session sharing" entry that used to sit here: the
  28.1s against 26.4s regression was blamed on Plan and Apply each opening their
  own Excel, but four launches is about 1.2 seconds. The other 27 were never
  Excel - they are worker startup, MCP round trips, and model coordination, and
  the only lever on those is fewer calls. See `docs/EXCEL-TUNING.md`.
- **Multi-workbook audit.** Follow external links through several workbooks into
  one dependency report.

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
