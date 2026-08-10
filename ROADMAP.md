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
4. **Read-only audit** (v0.7.0): one workbook's queries, connections, model,
   pivots, and external links - names and shapes only, unchanged-proof in the
   receipt. Field-confirmed safe on a real business workbook.

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

**1. Finish the discovery layer.** `table list` and `namedrange list` appear in
11 and 12 sessions. The audit already covers worksheets, macros, queries,
connections, the model, pivots and links; these two complete it, cost almost
nothing, and introduce no new concept.

**2. Reading cell values and formulas.** `range get-values` is the most-used
operation in the entire history - 31 of 46 sessions - with `get-formulas` at 21.
This corrects an earlier entry here that treated value reading as one incident's
blocker; it is the most frequent thing this work does, and ExcelTask returns no
cell data by design. The open question is whether a bounded read can answer the
caller's question without becoming a general data pipe: returning *what differs
from a pattern* rather than *the contents* would keep the promise while serving
most of the demand. Needs a design pass, not a ticket.

**3. Writing values and formulas.** `set-values` 19 sessions, `set-formulas` 15.
Both are deliberate refusals - inference plus verification is what makes an
ExcelTask edit safe, and accepting model-written formula text discards exactly
that. A genuine collision between a design stance and observed demand. It should
stay unresolved until reading is settled, because a good read may remove much of
the reason to write.

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
- **Macro session sharing.** The one measured regression: 28.1s against 26.4s,
  because Plan and Apply each open their own Excel.
- **Multi-workbook audit.** Follow external links through several workbooks into
  one dependency report.

## Standing rules

One tool. No CLI. No model selection anywhere. Schema bytes are budgeted and
every operation earns its own. Receipts stay bounded and truthful - a truncated
report says so, an uncertain outcome is `Unknown`, and no receipt ever carries
cell values, formulas, VBA source, connection strings, or machine paths.
