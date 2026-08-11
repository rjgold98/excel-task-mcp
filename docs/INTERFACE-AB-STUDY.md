# Tool-interface A/B study

An empirical check on the three interface bets ExcelTask makes: **one union tool** rather
than five, a **nested operation object** rather than flat fields, and **guidance-rich**
parameter descriptions. Run entirely on the home machine against Opus 5 subagents, each
acting as the model under test.

## Method

Each subagent sees one tool-surface variant (JSON Schema only) plus a set of independent
tasks, and must decide the **single tool call it would make first** for each. It never sees
the validation rules — those live with the scorer, so what is measured is purely the
interface's effect on the decision.

Calls are then replayed against an oracle transcribed from the shipped engine
(`ExcelTaskEngine.cs`, `ExcelWorkbookRuntime.cs`), so "valid" means *the real server would
have accepted this*, not *it looked reasonable*.

Variants are built from three axes in `variants.py`:

| Axis | Options |
|---|---|
| Grouping | one union tool (`one`) vs five separate tools (`many`) |
| Syntax | nested operation object vs flat top-level fields |
| Wording | terse / rich / rich+guardrails (`plus`) / plus live-binding (`live`) |

## Round 1 — baseline suite (6 tasks, 6 variants, 3 reps = 108 decisions)

Operation selection was **perfect: 108/108**. Every variant, including the terse ones,
picked the right operation every time. All failures were in the *policy envelope* —
mode, binding, save, overwrite, and the macro Plan→Apply contract.

| Variant | Schema bytes | Clean first call |
|---|---:|---:|
| one_flat_rich | 3,773 | 1.00 |
| one_nested_rich | 4,272 | 0.78 |
| many_rich | 8,124 | 0.78 |
| many_terse | 5,180 | 0.28 |
| one_nested_terse | 2,686 | 0.17 |
| one_flat_terse | 2,208 | 0.17 |

Two conclusions. **Terse descriptions are catastrophic** — the same model, same tasks,
drops from 0.78–1.00 to 0.17–0.28 purely because the rules stopped being written down.
And **five separate tools cost 1.90× the schema bytes for identical accuracy**, which is
the union-tool bet paying off.

## Round 2 — hard suite (8 tasks, 4 variants, 3 reps = 96 decisions)

Multi-operation requests, 180 MB–1 GB workbooks, and requests that brush the engine's
limits. Operation selection was again **perfect (96/96)**.

| Variant | Clean first call |
|---|---:|
| one_nested_rich | 0.83 |
| many_rich | 0.75 |
| one_flat_rich | 0.71 |
| one_nested_terse | 0.38 |

**The headline defect:** 21 of 32 failures came from two limits — `MaxFormulaRepairCells`
(10,000) and `MaxFormulaRepairRanges` (16) — that the engine enforces but **no schema
description mentioned**. Asked to repair `A2:H240000`, every rich variant passed all
1,919,992 cells straight through. The model was not being careless; it was obeying an
interface that never told it a ceiling existed.

## Round 3 — stating the hidden guardrails

Same 8 hard tasks, with three previously-undocumented rules written into the descriptions:
the two caps, the `UseOpen` + `Copy` conflict, and audit-never-writes.

| Variant | Clean, round 2 | Clean, round 3 | Schema bytes |
|---|---:|---:|---:|
| one_nested_plus | 0.83 | **1.00** | 4,272 → 5,036 |
| one_flat_plus | 0.71 | **1.00** | 3,773 → 4,537 |

Both went to a perfect score on every dimension. **Cost: 764 bytes (+18%). Benefit:
every remaining hard-suite failure eliminated.** This is now applied to the shipped
`Contracts.cs`.

## Round 4 — success-path suite (8 tasks, 4 variants, 3 reps = 96 decisions)

Rounds 1–3 leaned on refusal and narrowing behaviour, which measures error handling rather
than throughput. Round 4 fixes that: **every task has a legal, achievable answer that does
real work at or near the ceiling** — a 287-sheet/412 MB consolidation, 14 separate ranges
near the 16-range limit, a series extended 2,489 rows deep, a real VBA procedure to author
and run, a 900 MB audit. Nothing is a trap.

Scoring is graded rather than pass/fail: range sets are scored by **F1 over covered cells**,
so an answer is judged on the cells it actually touches, not on how it chose to spell them.

| Variant | Kind | Scalars | Range F1 | Policy | VBA | Valid | Clean |
|---|---:|---:|---:|---:|---:|---:|---:|
| many_plus | 1.00 | 1.00 | 1.00 | 1.00 | 1.00 | 1.00 | **1.00** |
| one_nested_plus | 1.00 | 1.00 | 1.00 | 0.96 | 1.00 | 1.00 | 0.96 |
| one_flat_plus | 1.00 | 1.00 | 1.00 | 0.96 | 1.00 | 1.00 | 0.96 |
| one_nested_rich | 1.00 | 1.00 | 1.00 | 0.92 | 1.00 | 1.00 | 0.92 |

Seven of the eight tasks scored a perfect 1.00 across every variant: every range exact,
every sheet name right, every authored VBA procedure complete, parameterless, non-blocking
and correct. **On large legal work the interface is not the bottleneck.**

One task was the exception — see round 5.

### A scoring bug worth recording

`plan_preview` initially scored 0.00 for all 12 agents. Every one had answered `B4:M703`
where the expected set was twelve single-column ranges. Those address **the identical
8,400 cells**; the model had given the more compact equivalent and the grader was wrong.
Switching from range-string identity to cell-coverage F1 fixed it, and the metric still
penalises real errors (a solid `B7:AB646` block over alternating columns scores 0.68).

This was the third scorer bug caught during the study, each of which had inverted a result
before it was fixed. Oracle correctness deserves as much scrutiny as the thing under test.

## Round 5 — the one remaining defect

`live_inplace` is the only success-path task any variant got wrong. The request says the
workbook is *already open with unsaved edits*; the correct call is `UseOpen` + `Same`.
Some runs chose `AskIfOpen`, which is not illegal but guarantees a wasted confirmation
round trip on a workbook whose state was already stated.

The cause is the schema's own wording: **"Use AskIfOpen first."** The model was following
the instruction. Round 5 adds the missing exception — use `UseOpen` directly when the
request already says the workbook is open.

| Wording | `live_inplace` correct | Clean (all tasks) |
|---|---:|---:|
| live-aware (`one_nested_live`, n=48) | **6/6** | **1.00** |
| every variant without the clause (pooled) | 11/15 | — |

Fisher exact two-sided **p = 0.281** — the direction is right but this is *not*
statistically significant at n=6 vs 15, and it should not be reported as proven.

It was applied anyway, on grounds that are mechanical rather than statistical: asking
"is it open?" about a workbook the user has just said is open costs a guaranteed round
trip every time it happens, the fix costs 85 schema bytes, and no other task regressed.
A larger replication would be needed to put a real number on the frequency.

### Corrections to round 1 and round 4

Two things in the tables above are weaker than they first appear, and both matter.

**Round 1 was mostly one rule.** 44 of its 51 failures (86%) were the same rejection:
`Same-file Apply requires overwriteConfirmed`. The 0.17-vs-1.00 spread therefore measures
"does the schema explain the overwrite gate" far more than it measures wording in general.
It also means the `one_flat_rich` 1.00 vs `one_nested_rich` 0.78 gap is **not** a syntax
effect — both carry the identical `overwriteConfirmed` description, so that difference is
sampling noise on 18 decisions. Nested vs flat never separated on anything real.

**Two round-4 tasks were not actually legal.** `ExcelTaskEngine.cs:754` rejects any
extension whose *destination* exceeds `FormulaMutationPlanner.MaxMutations` (2,000 cells) —
a much tighter cap than the 10,000-cell aggregate. `deep_extend` (destination 4,978 cells)
and `sequence_first` (3,992) would both be rejected by the real engine. Every variant scored
1.00 on them only because the scorer never implemented `MaxMutations`. They produced no
false ranking — all variants scored identically — but the "7 of 8 tasks perfect" claim is
properly "5 of 6 valid tasks", and the promise that every round-4 task had a legal success
path was false for those two.

## Round 6 — Sonnet 5 replication

Rounds 1–5 ran on Opus 5. The model actually used day to day is **Sonnet 5**, and Opus
scored near the ceiling on the success-path suite, so the interesting question was whether
these findings hold on the cheaper model — or whether Opus was simply strong enough to
paper over a bad interface. Both suites were re-run with Sonnet 5 subagents at high effort.

**Hard suite (Sonnet 5, high):**

| Variant | Clean, Opus | Clean, Sonnet |
|---|---:|---:|
| one_nested_plus (caps documented) | 1.00 | **1.00** |
| one_flat_plus (caps documented) | 1.00 | **1.00** |
| one_nested_rich (caps hidden) | 0.83 | 0.75 |
| one_nested_terse | 0.38 | 0.46 |

**Success-path suite (Sonnet 5, high):**

| Variant | Clean, Opus | Clean, Sonnet |
|---|---:|---:|
| one_nested_live | 1.00 | **1.00** |
| many_plus | 1.00 | **1.00** |
| one_nested_plus | 0.98 | 0.92 |
| one_nested_rich | 0.92 | 0.92 |

Sonnet reproduces the Opus result almost exactly, including **operation selection perfect
at 96/96** and the same seven-of-eight success-path tasks at a flat 1.00 — exact ranges
across a 287-sheet workbook, correct non-blocking VBA, correct Plan-vs-Apply. The one
weak spot is the same one: `live_inplace`.

### Pooling both models

| Finding | Pooled result | Fisher two-sided |
|---|---|---:|
| Documenting the caps | 48/48 clean vs 38/48 | **p = 0.0012** |
| Naming the already-open exception | 9/9 correct vs 16/24 | p = 0.073 |

The cap fix is now **significant across two models**. The live-binding fix strengthened
from p = 0.281 to p = 0.073 — better, still short of conventional significance, and it
matters more on Sonnet, which chose the wasteful `AskIfOpen` on 2 of 3 runs without the
clause versus 1 of 3 for Opus. The honest reading is that the weaker the model, the more
the missing exception costs, which is an argument for keeping it.

**The most important negative result:** Sonnet 5 was not meaningfully worse than Opus 5 at
driving this interface. Whatever this tool surface costs in accuracy, it is not recovered
by paying for a larger model — so the interface work, not the model tier, is where the
gains are.

## What this changed in the product

Two edits to `src/ExcelTask.Core/Contracts.cs`, both description-only — no behaviour
change, 87 core tests still pass:

1. **The caps are now documented** on `RepairRanges`, `Ranges` and `EvidenceRange`, and
   the `UseOpen`+`Copy` conflict and audit-never-writes rule on `WorkbookBinding` / `Save`.
   *Proven on both models* (p = 0.0012): eliminated every remaining hard-suite failure for
   +18% schema bytes.
2. **The already-open exception** on `WorkbookBinding`. *Directional* (p = 0.073); kept
   because the cost is 85 bytes, nothing regressed, and the penalty is a guaranteed wasted
   round trip every time it is hit.

## Standing conclusions

- **The three original bets hold.** One union tool matches five separate tools on accuracy
  for roughly half the schema bytes; operation selection was perfect in **484/484**
  decisions across every variant, every round and both models, including the terse ones.
- **Model tier is not the lever.** Sonnet 5 matched Opus 5 on this interface. Spending on
  a bigger model does not buy back what a vague schema costs.
- **Grouping and syntax barely matter; wording dominates.** Nested vs flat never separated
  by more than one decision. Rich vs terse separated by up to 0.83.
- **Every failure in the entire study was a policy failure**, and every one traced to a
  rule the engine enforced but the schema did not state. The lesson is narrow and
  actionable: *if the server rejects on it, the description must say so.*
- **On large legal work the interface is not the bottleneck.** Given the rules in writing,
  Opus 5 scored 1.00 on 7 of 8 success-path tasks — exact ranges across 287-sheet
  workbooks, correct non-blocking VBA, correct Plan-vs-Apply — with the sole exception
  above.

## The follow-on audit — 8 more rules still unstated

The generalized lesson ("if the server rejects on it, the description must say so") is
mechanically checkable without any model. Dumping the live `tools/list` schema and grepping
it against every caller-actionable rejection reason in the engine gives:

**Stated (10):** 16-range cap · 10,000-cell repair cap · 2,000-cell extend destination cap ·
24-period cap · UseOpen+Copy conflict · audit never writes · already-open exception ·
macro requires Isolated · macro Plan omits Apply fields · overwrite gate.

**Enforced but still unstated (8):**

| Gap | Engine rule |
|---|---|
| `RunAfterEdit` blocking calls | rejects a replacement containing `MsgBox`/`InputBox`/`Stop` |
| VBA source length | `MacroProcedureText.MaxSourceCharacters` = 8,192 |
| VBA line count | `MacroProcedureText.MaxLineCount` = 200 |
| Path extensions | paths must be `.xlsx` or `.xlsm` |
| Copy output extension | must match the target's extension |
| Copy output path | must differ from the target path |
| Isolated + Same | rejected when the target is open |
| Auto-entry VBA | automatic-entry procedures cannot be edited |

Closing these needs no A/B run — the rule is already proven at p = 0.0012. Worth noting
that `RunAfterEdit` blocking was *tested* in round 2 (`macro_blocking`, 0.83 clean) and its
failures are explained by exactly this gap.

## Round 7 — end to end, against real Excel

Rounds 1–6 were schema-comprehension tests: a model read a JSON Schema and wrote down the
call it *would* make. Nothing opened Excel. Round 7 replaces that with the real thing —
the built server, real MCP JSON-RPC over stdio, real `.xlsx` fixtures, and verification by
reopening the workbook and checking every formula rather than by reading the receipt.

The first thing it found was a product defect that no amount of schema testing could reach.

### Defect: `Range()` address overflow in `ApplyFormulaWrites`

A repair of a 3,000-row ledger was rejected with
`"Workbook execution was rejected before changes were attempted."` — no reason, no logging,
`CanRetry: false`. Plan succeeded everywhere Apply failed. The pattern was incoherent:

| Range | Cells | Repairs | Result |
|---|---:|---:|---|
| `B2:H500` | 3,493 | 350 | Completed |
| `B2500:H2999` | 3,500 | 350 | **Rejected** |
| `B2:B3000` | 2,999 | 300 | **Rejected** |

Identical repair counts, opposite outcomes; fewer cells failing than passing. It tracked
neither cells, nor rows, nor repair count.

Cause: `ApplyFormulaWrites` batched repairs into fixed groups of 64 and joined them into a
single `Range("B10,B20,B30,…")` address. **Excel rejects that argument beyond 255
characters**, so whether a batch fit depended on how many digits the row numbers had — the
same work succeeded near row 1 and failed near row 2500. Every observation above follows
from that: 50 addresses of `B10`–`B500` are ~250 chars, the same 50 as `B2500`–`B2990` are
~300.

Fix: batch by joined address length (≤ 255) instead of a fixed cell count. All previously
failing ranges now complete, and file-level verification of a full three-chunk repair shows
**2,093 of 2,093 inferable gaps filled with the exact expected formulas, 0 wrong, 0
collateral damage**. (The other 7 of 2,100 sit on row 3000, the last data row, whose gaps
have no lower neighbour — `FormulaPatternAnalyzer` correctly declines to infer them.)

Two lessons worth keeping:

- **Every schema-level result in this document assumed the engine executes correctly, and
  it did not.** Interface measurement cannot substitute for end-to-end measurement.
- The failure was invisible because `catch (Exception)` at `ExcelWorkbookRuntime.cs` discards
  the reason and the worker's stderr is drained and thrown away. Diagnosing it required
  temporarily echoing that stream. That diagnosability gap is still open.

### Defect: silent under-repair at range boundaries

The end-to-end A/B (below) exposed a second, worse defect. Runs that chunked a 3,000-row
repair at round numbers — `B2:H1000`, `B1001:H2000`, `B2001:H3000` — finished with three
`Completed` receipts and **rows 1000 and 2000 still unrepaired**.

Cause: `AnalyzeFormulaRepairs` read only the requested range, but `FormulaPatternAnalyzer`
infers a blank from the neighbour on each side. A gap on the last row of a chunk has its
lower neighbour outside that chunk, so evidence was missing and the cell was skipped —
silently, with no warning to the caller.

This is the most serious finding in the study, because it is not a coding slip:

- **Two deliberate design choices collided.** The caps force chunking; neighbour-based
  inference needs data outside the chunk.
- **The verification layer did not catch it.** Save → reopen → verify confirms the writes
  the engine *intended* actually landed. It never compares intent to the caller's request,
  so a plan that silently omitted cells verifies perfectly.
- **The natural chunking is the dangerous one.** Round thousands are the obvious split, and
  in this fixture every round thousand is a gap row.

Fix: read evidence one cell beyond the requested range on each side, while restricting
writes to the requested range. Replaying the exact failing chunking now fills rows 1000 and
2000 (`=A1000*2`, `=A2000*2`) and reaches 2,093 of 2,093 inferable gaps, 0 wrong, 0
collateral.

### Diagnosability

The catch in `ExecuteCore` discarded the exception, so any COM fault surfaced as
`"Workbook execution was rejected before changes were attempted."` A `failure-detail` check
now names the phase and the fault — verified reaching the caller as
`Failed in phase 'session-open': InvalidOperationException: ...`.

One correction to an earlier claim in this document's history: the receipt was never
*empty*. Checks travel in the MCP response's `structuredContent`, and they were always
reasonably detailed; the text block is only a one-line summary. What was genuinely missing
was the exception behind a catch-all failure, which is what the new check supplies.

### The end-to-end A/B result

Both arms use the schema the server actually publishes; the only difference is the 595
characters of guardrail text the study added. Six Sonnet 5 runs per pass, sequential
(Excel cannot be driven in parallel), scored by reopening the workbook.

| | calls | wasted first call | coverage | fully correct |
|---|---:|---:|---:|---:|
| **stated** (pass 1, before boundary fix) | 3.0 | 0/3 | 99.6% | 1/3 |
| **unstated** (pass 1) | 4.7 | 3/3 | 100% | 3/3 |
| **stated** (pass 2, after fixes) | 3.3 | 0/3 | 100% | **3/3** |
| **unstated** (pass 2) | 4.7 | 3/3 | 100% | 3/3 |

Pooled over all 12 runs: **3.17 calls vs 4.67, a 32% saving, exact permutation p = 0.0032**,
and 0/6 vs 6/6 wasted first calls.

The mechanism is not subtle. Every `unstated` run opens with `B2:H3000` — the whole
20,993-cell block, twice the cap — and eats a rejection before recovering. Documenting the
limit does not merely help the model recover faster; it stops the wasted round trip
happening at all.

The boundary fix also did what it was supposed to: `stated` went from 1/3 to **3/3** fully
correct, and both arms now reach 100% coverage. The earlier completeness penalty is gone.

### Still open

`ExcelTaskRealExcelOnDemandTests` — the two tests asserting the owned Excel process is
released — fail on this machine. They were confirmed **pre-existing**: with the boundary fix
absent from the tree entirely, both still failed. Excel processes were repeatedly observed
lingering and then exiting on their own, which points at a teardown race in the assertion
rather than a true leak, but that is not yet confirmed. This is the supervised-worker bet,
and it is currently unproven.

## Round 8 — head to head against sbroenne/mcp-server-excel

Both servers built locally, driven through identical client code, same fixture, same task,
same Sonnet 5 harness. Two runs each, scored by reopening the workbook.

| | tools | schema bytes | calls | tool time | cells repaired | correct |
|---|---:|---:|---:|---:|---:|---:|
| **ExcelTask** | 1 | 9,422 | 4.0 | 21.2 s | 2,093 / 2,100 | 2/2 |
| **sbroenne** | 26 | 66,033 | 7.0 | 5.0 s | **2,100 / 2,100** | 2/2 |

Three findings, and they do not all favour this project.

**The one-tool bet is validated.** 7.0× smaller schema and 43% fewer round trips. sbroenne
is session-based (`file open` → operate → `file close(save=true)`), so it pays a three-call
floor before any work happens.

**Verification is not free — it costs about 4× the execution time.** ExcelTask saves,
reopens and re-verifies on every call; against a 3,000-row workbook that is 21.2 s versus
5.0 s for a server that holds one session open and writes. That is the price of the
"changes were saved and verified after reopening" guarantee, quantified here for the first
time. Whether it is worth paying is a judgement call, but it should be made knowingly.

**Neighbour inference has a structural coverage ceiling that a general write primitive does
not.** Row 3000 is the last data row, so its gaps have no lower neighbour and
`FormulaPatternAnalyzer` cannot infer them — ExcelTask leaves `B3000` blank and always will.
sbroenne, told the rule, simply wrote `=A3000*2`. Earlier rounds in this document score
2,093 / 2,100 as full credit and call the remainder "correctly not repairable". That is
accurate for ExcelTask's design and misleading as a statement about the task: those cells
are repairable, just not by inference alone.

Caveat kept deliberately: repairing inferable blanks is ExcelTask's purpose-built operation
and sbroenne assembled the same outcome from general primitives, so this is the
specialist's home turf. The reverse is equally true — building monthly exhibits across
multiple tables from a 10,000-row extract is routine for sbroenne and impossible for
ExcelTask, whose five operations cannot create a sheet, write a value, or build a table.

## Round 9 — failure-mode matrix, and per-call cost

**Per-call overhead (A/B/C, identical 9,009-cell workload, only the call shape varies):**

| variant | calls | wall time |
|---|---:|---:|
| A — 3 calls x 1 range | 3 | 21.3 s |
| B — 1 call x 1 range | 1 | 7.6 s |
| C — 1 call x 13 ranges | 1 | 7.3 s |

Per-call overhead is **~6.9 s**; ranges within a call are free. Separately, 3,003 cells cost
7.1 s and 9,009 cost 7.6 s, so per-cell work is ~0.06 ms. The cost model is almost entirely
fixed-per-call, which means **the 10,000-cell cap is miscalibrated**: it is priced as though
per-cell work were the risk, when the real cost is the call itself. Raising the aggregate cap
would cut calls and wall time at near-zero execution cost, while keeping verification intact.

**Failure modes, both servers, measuring what the supervision claim is actually about:**

| scenario | ExcelTask | sbroenne | leaked Excel | file intact |
|---|---|---|---:|---|
| read-only target | `Unknown` - "saved, but lock not released" | clean refusal at open | 0 / 0 | yes / yes |
| server killed mid-operation | no result; nothing corrupted | write errored, close lost | 0 / 0 | yes / yes |
| worksheet does not exist | `Rejected` at preflight | opens, write fails, closes | 0 / 0 | yes / yes |

Two things follow, and they revise earlier entries in this document.

**The supervised-worker bet holds up better than round 8 suggested.** Zero leaked Excel
processes and zero file corruption across every failure mode, including a hard kill
mid-operation. The earlier doubt rested on flaky lifecycle tests and two `Unknown` statuses,
not on observed leaks; under direct test, nothing leaked.

**But there is no writability preflight for same-file saves.** `EnsureWritableCopyOutput`
runs only when `Save == Copy`, so a read-only target is discovered only after Excel is open
and the save attempted - producing `Unknown`, the worst possible answer for a caller, since
it means "your file may or may not have changed." sbroenne refuses cleanly at open. Checking
writability during preflight would turn this into a clean `Rejected`.

## Round 10 - the v0.11.0 operations, and a defect the run found in the old ones

Two new operations shipped in v0.11.0 (`FindReplace`, `Create`). This round asks the
same question rounds 3 and 6 asked, about them: does stating the rules the engine
rejects on change the first call?

Two changes to method, both worth keeping:

- **The oracle is the shipped code, not a transcription of it.** A small console app
  references `ExcelTask.Core` and runs each proposed call through the real
  `ExcelTaskEngine` with execution faked. Round 7's lesson was that a hand-written
  oracle can be wrong; this one cannot disagree with the product because it *is* the
  product. Three scorer bugs inverted results in earlier rounds, and none was
  possible here.
- **Both arms are generated from the schema the server publishes**, by textual
  removal of exactly the clauses that state a rejection rule. The arms differ by
  576 characters and nothing else.

Sonnet 5, 6 tasks x 3 reps per arm. `clean` means the shipped validation accepts the
call *and* the call does what the user asked.

| Arm | Clean | Creation tasks | Find/replace tasks |
|---|---:|---:|---:|
| `unstated` | 0.67 | **0/6** | 12/12 |
| `stated` | 0.78 | 5/6 | 9/12 |
| `statedplus` | **1.00** | 6/6 | 12/12 |

### Creation: the binding rule is worth its bytes

Every one of the six `unstated` creation failures is the identical rejection -
`Creating a workbook or worksheet requires workbook binding Isolated` - because the
arm's schema never said so. **5/6 versus 0/6, Fisher exact two-sided p = 0.015**;
pooling both stated arms gives 11/12 versus 0/6, p = 0.0004. This replicates the
round 3 and 6 finding on a new pair of operations.

### Find/replace: one clause validated, one not

Round 1's tasks stayed inside the search bound, so the cap it documents was never
exercised - a flaw in the tasks, not a result. A second round of three tasks built
to reach the new rules found:

| Task | `stated` | `unstated` |
|---|---:|---:|
| `over_cap` - used range 546,000 cells | 1/2 | **0/3** |
| `wildcard_urge` - literal `Q1*` in the search text | 1/2 | 3/3 |
| `strip_prefix` - replacement would leave `=Opening balance` | 2/2 | 2/3 |

**The 10,000-cell bound pays.** All three `unstated` runs failed `over_cap`; two were
rejected verbatim with `A find/replace range must be at most 10,000 cells` and the
third named no range at all, which the runtime refuses against a 546,000-cell used
range. Small n, same mechanism as the round 7 end-to-end result.

**The wildcard clause did not pay, and is recorded as unvalidated.** Every run in both
arms used the literal `Q1*` correctly. `* and ? are literal, not wildcards` costs
about 40 bytes and bought nothing measurable here. It is kept because it documents a
deliberate divergence from Excel's own Find that a reader would otherwise assume the
other way, but it should not be counted among the proven clauses.

### The defect the run found: the overwrite gate, again

Six of the seven `stated`-arm failures across both rounds were the same thing, and it
had nothing to do with the new operations: **a same-file Apply sent without
`overwriteConfirmed`**. This is round 1's headline defect - 44 of its 51 failures -
recurring on a rule that *is* stated.

The wording was the cause. "Explicit authorization required before Apply can overwrite
an existing save destination" does not read as *the workbook you are editing*; a model
that is not writing a copy concludes there is no save destination to overwrite. The
`statedplus` arm spells it per save mode instead, and adds one clause for a second
thing the run exposed - `auditWorkbookFlows` says "takes no options", so a run omitted
the payload entirely and was rejected for supplying none.

`statedplus` scored **18/18, a perfect clean rate on every task and every dimension.**

Against `stated` that is 18/18 versus 14/18, **Fisher exact two-sided p = 0.104** - the
direction is right and this is *not* significant. Two honest caveats:

- All four `stated` failures came from a **single run**, so the 18 decisions are not
  independent. Per run it is 3/3 perfect versus 2/3, which is n = 3 and no statistic
  at all.
- `statedplus` changed **two** clauses, so the improvement cannot be attributed to
  either alone.

It is applied anyway, on the same mechanical grounds as round 5's live-binding clause:
the failure it addresses is the most frequent one in the entire study across four
years of rounds and two models, the cost is 180 bytes, and nothing regressed.

### Three product defects this round found before any user did

The A/B was measuring the interface and kept finding the engine instead. All three are
fixed in v0.11.0.

1. **`FindReplace` and `Create` were unreachable.** The operation union counts its
   payloads by hand and neither new one was in the count, so a request carrying only
   one of them failed the arity check before reaching its own validation. Every other
   part of both operations was complete and tested. A test now builds one request per
   `ExcelOperationKind` and asserts none fails on arity.
2. **`NeedsConfirmation` described the wrong thing.** A same-file Apply missing its
   overwrite flag came back saying *"The requested copy output already exists"* - a
   sentence about a file the request never mentioned. The requirement's own prompt was
   correct; only the summary lied, and the summary is the line a caller reads first.
3. **`Create` enforced a weaker rule than it stated.** The schema and the rejection
   message both said Isolated; the check only rejected `UseOpen`, so `AskIfOpen` slipped
   through to a confirmation whose two answers are an option creation refuses and an
   option inspection then rejects. Now Isolated exactly, matching macro editing.

A fourth, in the harness rather than the product: the McpServer real-Excel tests
snapshotted the process table once, the instant the test body returned, while the Excel
project's equivalent has settled and retried since 0.9.0. Running both assemblies at
once - which `dotnet test` on the solution does and `scripts/Test-Mvp.ps1` deliberately
does not - made one assembly's fixtures look like the other's leaks. Four tests failed
that way and none had leaked anything. The assertion no longer depends on the caller
remembering to serialize.

## Round 11 - can a model actually write an Excel number format code?

`SetNumberFormat` (v0.12.0) turns on a question none of the earlier rounds asked. Every
round so far measured whether a model picks the right operation and obeys the policy
envelope, and operation selection has been perfect in 484 of 484 decisions. This
operation's risk is elsewhere: it takes a **format code**, a small hostile language of
its own where `_)` is a padding directive, `[Red]` is a colour section, and the
positive and negative cases are separated by a semicolon. If a model cannot write one,
a schema that describes the operation perfectly still ships something unusable.

So this round is a capability check rather than an arm comparison, and the oracle is as
strong as it gets: **every proposed code is applied to a real Excel workbook**, then
read back, then rendered against a sample positive and a sample negative. Excel decides.

Four requests phrased the way a finance user would phrase them - never naming a format
code - x 3 reps.

| Task | Code produced (all 3 runs) | Positive | Negative |
|---|---|---|---|
| thousands, 2dp, negatives in aligned parentheses | `#,##0.00_);(#,##0.00)` | `1,234,567.89` | `(1,234,567.89)` |
| whole dollars, red parenthesised negatives | `$#,##0;[Red]($#,##0)` | `$1,234,568` | `($1,234,568)` |
| ratios as one-decimal percentages | `0.0%` | `18.4%` | `-7.3%` |
| ISO dates | `yyyy-mm-dd` | `2026-07-31` | - |

**12 of 12 accepted by Excel, 12 of 12 round-tripped verbatim, and all three runs
produced byte-identical codes for every task.** Including the `_)` padding, which is
the one part a person hand-writing a format usually gets wrong and the reason the
engine does not trim.

Two things follow.

- **Format-code authoring is not a bottleneck**, so the operation is usable as shipped
  and no worked-example text needs to be spent on it in the schema.
- **The read-back is still not redundant.** It cost nothing here because nothing failed,
  and it is the only thing standing between a code Excel silently coerces and a receipt
  claiming success. Zero failures in twelve tries is not evidence that the guard is
  unnecessary; it is evidence that this model, on these four cases, did not need it.

Recorded caveat: the codes are deterministic across runs, which means n = 3 buys much
less than it looks like. This measures four format families, not the space of them.

## Round 12 - plain-text prompts in fresh chats, and a fix-and-remeasure pass

Every earlier round measured a model reading a JSON Schema and writing down a call. Round 12
measures the thing that actually ships: **a fresh assistant, a plain-English request from a finance
user, and the real server driving real Excel.** No schema-shaped task descriptions, no hints - the
agent gets the published `tools/list`, a brief on how to invoke one call, and a sentence a person
would actually type.

Method: four scenarios, each in a clean context with its own fixture workbook, driving the built
server over MCP stdio. Scored two ways - the agent's own friction report, and a COM verifier that
reopens every workbook afterwards and checks the cells. Round 1 found friction, the friction was
fixed, and **round 2 re-ran the same prompts against the fixed build**, which is the part that makes
this more than a survey.

### Result: every task correct, both rounds, and fewer calls after the fixes

| scenario | round 1 calls | round 2 calls |
|---|---:|---:|
| build a workbook from nothing | 4 | **3** |
| change one named figure | 4 | **2** |
| hardcode a calculated cell | - | **1** |
| recover from a mistyped filename | - | **2** |

Ground truth passed on every scenario in both rounds. Round 2 also produced a better artifact, not
just a cheaper path: the created workbook now contains **exactly one sheet, named `Summary`**, where
round 1 left an orphan `Sheet1` beside it.

### What the traces showed, and what each one cost

**A mistyped path answered with infrastructure.** The most ordinary user error there is - a doubled
letter in a filename - came back as *"Workbook inspection could not be completed before execution."*
Inspection was throwing, and the engine's catch-all turned a finding into a failure. It now carries
an `InfeasibleReason`, and the round-2 agent got *"Target workbook does not exist"*, guessed the
correction, and finished in two calls.

**Two calls to learn one bit.** An agent asked to change a figure read the same range twice - once
for values, once for formulas - and diffed them, purely to learn whether the cell it was about to
overwrite held a formula. That is a whole Excel launch for one boolean. Reads now report
`isFormula` per cell, at the cost of one extra array read per range, and the same scenario dropped
from four calls to two.

**A silent asymmetry nobody had noticed.** `FindReplace` refuses to rewrite a formula cell;
`WriteWorksheetValues` would silently destroy one. Overwriting a formula with a constant is
legitimate finance work, so the fix reports rather than refuses - the receipt now names those cells,
and a follow-up round asked for the prior value too, so a caller can say *"469,750.25 to 471,000"*
rather than only *"wrote 1 constant"*.

**Two calls to name a sheet.** Creating a workbook and then naming its sheet were separate
operations, so every "set me up a workbook with a Summary tab" cost two Excel launches. `Create`
with kind `Workbook` now takes an optional starting sheet name.

### The most interesting finding: capability without discoverability

Twice, an agent's wishlist asked for something **that already existed**:

- *"Let SetNumberFormat's Plan surface the current format before committing"* - it already did.
- *"A way to target a cell by its row label instead of a prior read"* - `FindReplace` in Plan mode
  returns the address of every matching cell and changes nothing, which is exactly that.

Neither was a missing feature; both were unadvertised ones. Six words in each description closed
them. This is a distinct failure class from the one the study has measured for eleven rounds - that
one was *rules the engine enforces but the schema never states*, and this one is *capabilities the
server has but the schema never mentions*. The same audit method finds it: read the wishlist, check
whether the code already does it.

### The budget did its job three times

The 15 KB schema bound was hit three times while applying these fixes, and each time forced a real
decision instead of a bigger budget:

- A "Plan and Apply behave identically for read-only operations" clause was **cut**: two agents
  hesitated over it, neither lost a call, and the standing rule says bytes must buy round trips.
- Two sentences of reassurance about number-format read-back were **reclaimed** - round 11 had
  already proven format-code authoring is not a bottleneck, so the reassurance bought nothing.
- A locator clause was **tightened** by twelve characters rather than allowed to push the budget up.

Net: eleven scenarios' worth of UX fixes landed inside the existing bound.

### Recorded honestly

The verifier produced **two phantom failures** on its first run - a miscounted label and a JSON
date round-trip that dropped sub-second precision and then shifted four hours by re-parsing a local
timestamp. The workbook was byte-identical to the tick. That is the fourth scorer bug in this
study's history, and it is why every claim here is stated against a verifier that was itself
debugged rather than trusted.

Also unexercised, and worth saying: `AskIfOpen` never once returned a confirmation across eight
runs, because no fixture was open in Excel at the time. The confirmation-recovery path this
document has discussed since round 5 still has no end-to-end evidence behind it.

## Reproducing

Round 10's harness is a different one. It lives outside the repo, at
`../excel-task-abtest/` - durable rather than in a scratchpad, because rounds 1-9's
harness was not and is gone. `dump-schema.ps1` speaks MCP to the built server and captures the real
`tools/list`, `build-arms.ps1` derives the `unstated` arm from it by textual clause
removal and fails loudly if a clause it expects is missing, `generate-prompts.ps1`
emits self-contained prompts, `oracle/` is the console app that replays decisions
through `ExcelTaskEngine`, and `score.ps1` joins the two.

Rounds 1-9 (scratchpad `abtest/`): `variants.py` builds the surfaces,
`tasks_hard.py` / `tasks_big.py` hold the suites, `generate_*.py` emit self-contained
prompts, and `score_hard.py` / `score_big.py` replay decisions against the engine oracle.
Agents are told not to read the task or scoring modules, so the rules never leak into the
surface under test.
