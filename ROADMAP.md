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

Measured against the original server on the work computer: 8.1x smaller tool
surface; 74% fewer input tokens, 73% fewer model requests, 84% fewer MCP calls,
53% less wall time across three client workflows - with ExcelTask's own Excel
execution 13% slower, the advantage being entirely the removal of model
coordination. One run per workflow; not yet a benchmark.

## Next release: v0.8.0 (built, gated, held)

Held only for coordination with in-flight interface work, then ships:

- **Macro discovery.** The audit lists macro components and procedures, and the
  schema routes unknown names to it. Built the day the first real task split
  across both servers exactly at this gap - the field agent's verdict was "if
  excel-task added a read-only list of modules and procedures, it'd stand
  alone." See `docs/field-reports/2026-08-10-mixed-server-macro.md`.
- **One-rejection macro policy.** Every unmet requirement in a single message;
  field use paid one round trip per rule to learn them one at a time.
- **Measured schema improvements** from the interface A/B study: bounds stated
  in descriptions, binding and save rules made explicit.

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

- **Values are writable within the existing stance.** The refusal protects
  against model-authored *formulas*, which can be plausibly and silently wrong in
  a way no receipt would catch. A constant - a label, a number, a date - has no
  such failure mode: it is exactly what was asked for, and a read-back proves it
  byte for byte. This is 19 sessions of demand available without giving anything
  up.
- **Formula text remains a genuine collision.** Inference plus verification is
  what makes an ExcelTask edit safe, and accepting formula text discards exactly
  that. `ExtendFormulaSeries` and `RepairExistingWorksheet` already serve the
  cases where the intended formula is derivable from evidence. What is not yet
  known is how much of the 15 sessions those two already cover, and that is
  measurable rather than arguable.

**4. Then, in demand order.** `range_edit` find and replace (13 sessions),
`screenshot` for verification (13), `range_format` (12), `table` beyond listing
(11), Data Model and Power Query mutation (10 each).

**Not to be built.** `chart_config`, `pivottable_calc` and `worksheet_style` were
never called once in five weeks; `chart`, `slicer` and `table_column` appeared in
one session each. Roughly half of the original server's 230 operations show no
demand at all - which is most of the 8.1x schema ExcelTask does not carry.

**Open from earlier evidence, unranked by this data:**

- **Writability preflight for same-file saves.** A read-only target is found only
  after Excel is open, producing `Unknown` - the worst answer for a caller, since
  it means the file may or may not have changed. Preflight makes it a clean
  `Rejected`.
- **The second Excel launch on a mutating Apply.** Measured at 274-482 ms of a
  roughly one-second operation, and the largest remaining cost by far: launching
  Excel is the entire budget, and everything done inside it is rounding error.
  It cannot be deleted - verifying in the process that did the writing would be
  verifying against the memory that produced it - but it can be started early so
  it overlaps the write and save. Not built yet because a pre-launched instance
  is a new way to leak an Excel process on every early-return path, which is the
  one thing this project claims never happens. See `docs/EXCEL-TUNING.md`.
- **Multi-workbook audit.** Follow external links through several workbooks into
  one dependency report.

## Standing rules

One tool. No CLI. No model selection anywhere. Schema bytes are budgeted and
every operation earns its own. Receipts stay bounded and truthful - a truncated
report says so, an uncertain outcome is `Unknown`, and no receipt ever carries
cell values, formulas, VBA source, connection strings, or machine paths.
