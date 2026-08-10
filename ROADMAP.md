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

## Candidates, strictly demand-gated

Built only when real use shows a recurring need, in the order demand appears:

- **Bounded read-only value inspection.** The largest open question. The first
  real task needed range reads to find the bug, and ExcelTask returns no cell
  values by design; today the original server fills that role. Either the
  two-server split is the end state - the original for reading and exploring,
  ExcelTask for edits that must not go wrong - or ExcelTask grows a bounded
  read. Widening what the product promises never to return is not done lightly;
  more real tasks decide.
- **Macro session sharing.** The one measured regression: 28.1s against 26.4s,
  because Plan and Apply each open their own Excel. Worth building only if
  macro editing turns out to be frequent.
- **Module-level edits.** Whole-procedure replacement cannot introduce a
  module-level constant; one field occurrence, clean workaround, not yet demand.
- **Multi-workbook audit.** Follow external links through several workbooks
  into one dependency report. The natural growth of phase 4 once single-workbook
  audits prove themselves on rich content.

## Standing rules

One tool. No CLI. No model selection anywhere. Schema bytes are budgeted and
every operation earns its own. Receipts stay bounded and truthful - a truncated
report says so, an uncertain outcome is `Unknown`, and no receipt ever carries
cell values, formulas, VBA source, connection strings, or machine paths.
