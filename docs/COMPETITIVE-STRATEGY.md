# Beating the field on its own axes without giving up the design

`docs/LANDSCAPE.md` is the map: every surveyed competitor beats ExcelTask on **breadth**, **speed**,
or **portability**, and ExcelTask beats all of them on not being wrong. This is the route. The
claim it argues is that all three axes are reachable *from* the design rather than against it,
because one move unlocks all three — and that move already shipped.

## Status of this document's claims, as of v0.17.0

Read this first; the rest was written before any of it was tested.

| Claim | Status |
|---|---|
| ExcelTask operates in both camps (live COM and file format) | **Holds.** `ScanWorkbookStructure` shipped v0.14.0 and was extended v0.15.0. |
| The measured phase timings below (44 ms Excel work, 2,814 ms teardown, 2,297 ms verify) | **Holds**, on the home machine. Never reproduced on managed hardware. |
| ~44% off the four accelerable operations from file-based verification | **Still projected.** Arithmetic from the timings, not an observed result. Nothing has been built. |
| "Correctness nobody has measured on a real machine" is the durable advantage | **Weakened by our own evidence.** Two correctness defects shipped and were found by running the binary against hostile fixtures, not by the suite. The advantage is real only to the extent the checking is real; see below. |
| The competitor comparison numbers (8.1x surface, 74% fewer tokens…) | **One run per workflow**, on the work computer, at v0.10.x. Not a benchmark, and now four releases stale. |

The fourth row is the one that changed. This document argues that the field is
uniformly untested and that this is ExcelTask's opening. That is still true of
the field — but on 2026-08-11 two defects shipped here that a full green gate did
not catch, both found only by running the released binary against inputs designed
to break it. The lesson is not that the claim is wrong; it is that "measured
correctness" is a practice to keep performing, not a property already owned. What
made the difference was adversarial fixtures and verifying findings rather than
adopting them, and both are now in the process rather than in this argument.

## The move: reading the file is not a compromise, it is the second mechanism

`ScanWorkbookStructure` (v0.14.0) reads the workbook as the ZIP of XML it physically is, with no
Excel process. That was built for planning, but its real significance is that **ExcelTask now
operates in both camps** — the live-COM camp for fidelity, the file-format camp for everything that
does not need Excel's opinion.

Every competitor is stuck in exactly one camp. The file-format servers cannot recalculate, run a
macro, or refresh a query, ever. The COM servers cannot answer a question without paying for an
Excel process. ExcelTask can now choose per question, and that choice is what the three axes turn
on.

## Speed: the measurement, and where the win actually is

Traced on a 0.72 MB, 20,000-row workbook, applying a number format across 2,000 cells:

| phase | ms | share |
|---|---:|---:|
| session-open | 522 | 10% |
| format-preflight | 13 | <1% |
| **the actual work** | **26** | **<1%** |
| save | 63 | 1% |
| primary-cleanup (close + prove exit) | 2,325 | 44% |
| **reopen-verification (second Excel)** | **2,359** | **44%** |
| total | 5,328 | |

The work is 26 ms. Everything else is Excel lifecycle. Competitors are not faster because their
code is better; they are faster because they never pay this. xlwings-mcp keeps sessions alive to
avoid the launch; the file-format servers avoid the whole thing.

Two of those rows are addressable and one is not:

- **`primary-cleanup` (2,325 ms) is not addressable.** It is Close, Quit, and prove the process
  exited. That proof is the product's central claim. It stays.
- **`reopen-verification` (2,359 ms) is addressable, and this is the whole opportunity.** The
  question it answers — "does the saved file really contain what we asked for" — is a question
  about *the file*, and the file is right there. Answering it by launching a second Excel is one
  valid way, not the only one.

### Proven, with the trap it exposes

A spike read the same assertion straight out of the saved package: resolve `xl/styles.xml` to a map
of style index → number format code, stream the sheet part, and check every cell's `s` attribute.
Result: **2,000 of 2,000 cells VERIFIED**, and a negative control on a never-formatted range
correctly reported `General` and refused. The file genuinely holds the answer.

It also produced the finding that makes this a bounded strategy rather than a free win. The first
attempt reported *every* cell as wrong, because `#,##0.00` **is built-in format id 4** — Excel
writes only the id and no custom entry, so a verifier reading custom codes alone sees nothing.
Resolving it requires the ECMA-376 built-in table, and **ids 14–22 and 45–47 are
locale-dependent**. That is precisely the "partial re-implementation of Excel semantics" the
landscape survey identified as EPPlus's and ClosedXML's weakness. Adopting it wholesale would import
the weakness ExcelTask exists to avoid.

### So the rule is: accelerate where the file is unambiguous, fall back where it is not

Verification never weakens. It gets cheaper only where the package answers definitively:

| operation | file-verifiable? | why |
|---|---|---|
| `WriteWorksheetValues` | yes | the constant is literally in the XML |
| `FindReplace` | yes | same - constants only, by design |
| `SetNumberFormat` | yes, **except** locale-dependent ids 14–22, 45–47 | built-in table is spec, not guesswork; the date/time ids are not |
| `Create` | yes | the sheet either exists in the package or does not |
| `RepairExistingWorksheet`, `ExtendFormulaSeries` | **no** | formulas need recalculation; cached values are not proof |
| `EditMacroProcedure` | **no** | VBA lives in a binary OLE container, and a run must actually run |

Expected effect on the four accelerable operations: roughly **5.3 s → 3.0 s, about 44%**, with the
cleanup proof and the receipt untouched. A fallback to Excel verification on any ambiguity means
the guarantee is unchanged; only its cost is.

**Second speed lever, already justified and not yet taken.** Round 9 measured per-call overhead at
~6.9 s fixed and per-cell work at ~0.06 ms. The 400-cell read bound is priced as though per-cell
work were the risk. On a 20,000-row exhibit it forces 50 calls where 5 would do. Raising it costs
almost nothing in execution and removes 45 round trips - and the receipt bound, not the read bound,
is what actually protects the response.

## Breadth: the axis is mostly fake, and the real part is small

The incumbent's 234 operations sound like a rout until the demand data is applied: across 46 real
sessions, roughly half the original server's surface was **never called once**, and three whole
tools scored zero. Matching 234 operations would be matching mostly dead weight — and the incumbent
now self-reports what that weight costs, at ~163 K tokens of tool schemas per request against
ExcelTask's ~16 KB.

The genuinely missing breadth is short and known:

1. **Formula writing** (`set-formulas`, 15 sessions) — the one real gap, discussed below.
2. **Formatting beyond number formats** (12 sessions) — gated on an operation-level count, not on
   effort.
3. **Tables beyond listing** (11 sessions).

That is three items, not 220. "Breadth" is won by covering the measured demand completely, not by
matching a catalog.

**And the one-tool design does not cap breadth — the schema budget does.** Those are separable.
Operation payload descriptions are the bulk of the 16 KB, and MCP offers resources and prompts as
places to put detail that a caller fetches only when it needs it. That is worth designing *before*
the budget forces a bad cut, not after.

### The formula gap, and the option never evaluated

Refusing model-written formula text is the correct call and the landscape confirms how unusual it
is — every competitor accepts them, one writing `=`-prefixed strings verbatim with no validation and
no recalculation. But refusal is the *strict end* of a spectrum the survey found, not the only
principled point on it:

- **Refuse** (ExcelTask today)
- **Constrain to a grammar** — SheetMind restricts generation to a closed BNF over seven operations,
  so what the model emits is executable by construction
- **Embrace with a sandbox** — Microsoft's own SheetBrain runs model-written Python in one

A grammar-constrained formula operation would keep every property that matters: the model cannot
emit arbitrary text, the result is verifiable, and a bad formula is rejected structurally rather
than discovered later. **It should not be built until the prerequisite measurement is done**, which
the roadmap already names: how much of those 15 sessions `ExtendFormulaSeries` and
`RepairExistingWorksheet` already cover. If inference covers 12 of 15, the refusal costs almost
nothing and stays. If it covers 3, this is the most valuable unbuilt thing in the project.

## Portability: partial, honest, and already half true

Full portability contradicts the design — mutations need real Excel, and that is the point. But the
scan already proved that a meaningful subset does not:

- **Runs anywhere today, in principle:** `ScanWorkbookStructure` — pure ZIP and XML, no COM, no
  Windows API. It is Windows-only right now only because the assembly is.
- **Could join it:** `ReadWorksheetRange` against a saved file, and most of `AuditWorkbookFlows`
  (sheets, tables, defined names, and external links are all plain XML; queries, the data model,
  and VBA are not, and must stay with Excel).

The honest position is not "cross-platform" but something better-defined and more defensible:
**every read works anywhere; every mutation is verified by real Excel.** No competitor offers that
combination — the file-format servers cannot verify with Excel at all, and the COM servers cannot
answer without one.

## Beating each one specifically

The axes above are the mechanism. This is the scoreboard, one competitor at a time. Every row
marked *projected* is arithmetic from the measurements in this document, not an observed result.

### sbroenne/mcp-server-excel — the incumbent

Same mechanism (C#, live COM), so fidelity is a tie and neither side wins on "does Excel really do
it." It wins today on **capability** (Power Query, DAX, PivotTables, charts) and on offering a CLI
mode its own benchmark says costs 64% fewer tokens than its MCP mode.

- **Win on context cost — already true, and it publishes the number itself.** ~163 K tokens of tool
  schemas per request against ~16 KB. Nothing needs building; it needs *showing*, which is the
  field gate's job.
- **Win on trust — already true.** It has no verification after write, no proof the Excel process
  exited, and its lifecycle model could not even be characterised from its documentation.
- **Do not chase its capability list.** Half of it was never called once in 46 sessions. Chasing it
  is how you acquire its token problem.
- **Take its one good idea.** The CLI-vs-MCP token gap is real and is an argument for putting
  detail behind MCP resources rather than in the tool schema.

### xlwings-mcp-server — the closest architecture

Python, live COM, Windows-only, same fundamental bet. It wins on **per-call speed** by keeping Excel
sessions alive, and on **extensibility** because Python is easier to move in than C# plus COM.

- **Win on speed without adopting sessions.** Its advantage is avoiding the launch; yours is
  avoiding the *verification* launch, which the trace shows is the bigger number — 2,359 ms against
  522 ms. File-based verification takes roughly 44% off the four accelerable operations *(projected)*
  while keeping statelessness, which is what makes a leak impossible in the first place.
- **Win on safety — already true.** Sessions held open with TTL eviction is a leak surface with no
  proof-of-exit behind it.

### haris-musa/excel-mcp-server — openpyxl

Wins outright on **portability** (no Excel, any OS, headless, CI) and on raw speed for file edits.

- **Do not try to beat it at being headless.** It will always win that, because it gave up Excel.
- **Beat it on the claim it cannot make.** It cannot recalculate, run a macro, or refresh a query -
  ever. Your position is *every read works anywhere; every mutation is verified by real Excel*, and
  it can only ever offer the first half.
- **Narrow the gap where it costs nothing.** The scan is already pure ZIP and XML; extending
  file-based reads means the portable half of that sentence becomes literally true.

### negokaz/excel-mcp-server — Go, 7 tools

The closest philosophical relative: it also believes in a small surface, and its single-binary npm
distribution is genuinely better packaging than a 39 MB zip.

- **Win decisively on correctness.** It writes any `=`-prefixed string verbatim to `SetFormula`,
  with no validation, no recalculation, and no read-back, then reports success. That is the exact
  silent-wrongness failure this project exists to prevent, and it is worth stating plainly in any
  comparison.
- **Learn from its distribution.** Install friction is a real axis and it beats you there.

### The competitor not yet surveyed

Microsoft's own first-party direction — Copilot in Excel, or a Graph workbook MCP — would be
supported, cross-platform and cloud-native, and would undercut both camps at once. The survey found
nothing, but it also did not look hard. **This is the one that should be checked before another
multi-week investment**, because it is the only one that could make the whole category moot.

## The order

1. **Run the work-computer field gate.** Every claim here is theory until the numbers come from the
   machine that matters. Nothing else on this list should start first.
2. **File-based verification for the four accelerable operations**, with a fallback to Excel on any
   ambiguity and an explicit refusal to resolve locale-dependent format ids. ~44% off those
   operations, guarantee unchanged.
3. **Recalibrate the read bound** against the measured cost model. Cheapest real win available.
4. **Measure formula-inference coverage.** It decides whether the grammar-constrained operation is
   the most valuable unbuilt thing or an unnecessary risk.
5. **Corpus-audit the real exhibits.** Unblocks formatting, tables, and Power Query in one pass.

## What would make this strategy wrong

- If the field gate shows the work computer cannot run macro editing or COM at all, the fidelity
  argument for live Excel weakens sharply, and the file-format camp becomes more attractive than
  this document assumes.
- If the formula-coverage measurement shows inference already covers most real demand, item 4
  disappears and the refusal is simply correct.
- If file-based verification is ever caught disagreeing with Excel verification on the same
  assertion, item 2 must be reverted whole — two implementations that can disagree about what a
  workbook contains is the exact defect class this project spent v0.13.0 removing.
