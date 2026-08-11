# Where ExcelTask sits in the field

Deep-research survey, 2026-08-11. Method: 5 search angles, 23 sources fetched, 115 claims
extracted, top 25 adversarially verified by three independent votes each — 22 confirmed, 3
refuted (listed at the end, and not used). Repository facts are as of this date in a fast-moving
ecosystem.

Spot-check, same day, independent of the original run: eight of the load-bearing citations were
re-fetched at source. All eight held, including the two carrying the most specific numbers — the
tool-description study's 97.1%/856/103 and the verified-tool-calls 100%-vs-64%, the latter found
in the body (Section 6.1) rather than the abstract, where checking only the abstract would have
produced a false refusal. Two attributions were softer than the sources supported and have been
narrowed in place: EPPlus's function count carried a version the wiki does not state, and
SpreadsheetBench's venue is not named on its own project page.

## The fault line: live Excel vs the file format

Every spreadsheet MCP server picks one of two mechanisms.

**Live desktop Excel over COM** (Windows-only, Excel required): ExcelTask; the incumbent it
replaced, [sbroenne/mcp-server-excel](https://github.com/sbroenne/mcp-server-excel) (C#,
PIA-typed interop where ExcelTask is pure late-bound IDispatch — same mechanism, different
binding style); and [xlwings-mcp-server](https://pypi.org/project/xlwings-mcp-server/)
(Python, xlwings/pywin32).

**The file format directly** (cross-platform, headless, no Excel):
[haris-musa/excel-mcp-server](https://github.com/haris-musa/excel-mcp-server) (Python,
openpyxl) and [negokaz/excel-mcp-server](https://github.com/negokaz/excel-mcp-server) (Go,
excelize, distributed as npm-bundled binaries). Trade-off inherent to the camp: no live
recalculation, no macros, no real Excel semantics.

Microsoft's own position is against ExcelTask's camp:
[KB 257757](https://support.microsoft.com/en-us/help/257757/considerations-for-server-side-automation-of-office)
states verbatim that Microsoft "does not currently recommend, and does not support, Automation of
Microsoft Office applications from any unattended, non-interactive client application or
component," and that Open XML file manipulation "is the recommended and supported method for
handling changes to Office files from a service." The same KB documents the canonical hazard: "A
modal dialog box on a non-interactive desktop cannot be dismissed. Therefore, that thread stops
responding (hangs) indefinitely." That is precisely the failure class the supervised worker, hard
deadline, and proof-of-exit exist to bound. Qualification that matters here: ExcelTask's typical
use — an interactive desktop with a person in the chat loop — is closer to supported client-side
automation; the unsupported status attaches to unattended operation.

As of v0.14.0, ExcelTask is the only surveyed server in **both** camps: `ScanWorkbookStructure`
reads the OOXML ZIP directly — the other camp's mechanism, adopted for the one job it is
unambiguously better at (read-only structure, no process, no teardown cost) — while every
mutation stays on live Excel for fidelity.

## Every defining choice is the outlier

| Choice | ExcelTask | The field |
|---|---|---|
| Tool surface | 1 deep tool, 11 operations, ~16 KB | 7 (negokaz) to 26 tools / 234 operations (incumbent); xlwings ~19 documented, ~29 registered |
| State | Stateless per call | xlwings-mcp is explicitly session-based (`open_workbook` → session_id → `close_workbook`, TTL/LRU eviction); the incumbent's model could not be characterized (a daemon/named-pipe claim was refuted 0-3) |
| Model-written formulas | Refused; inferred from evidence, verified after reopening | Accepted by every competitor that writes formulas — haris-musa syntax-checks plus a 5-function blocklist; negokaz writes `=`-prefixed strings verbatim to `SetFormula` with no validation and no recalculation; xlwings-mcp same accept-with-syntax-check pattern |
| Verification | Save, reopen in a separate Excel, read back; uncertain = `Unknown` | No surveyed system has any counterpart (absence of evidence in this survey, not proven uniqueness) |
| Process lifecycle | Supervised worker subprocess, deadline, identity-checked cleanup, proof-of-exit in the receipt | No counterpart found |
| Schema wording | A/B-tested against models, every enforced rule stated | An empirical study ([arXiv 2602.14878](https://arxiv.org/pdf/2602.14878)) found 97.1% of 856 tool descriptions across 103 MCP servers carry at least one quality defect; 56% fail to state the tool's purpose |

Outlier does not mean wrong. The two loudest corroborations come from outside:

- **The incumbent now documents ExcelTask's thesis itself.** Its README benchmark reports ~163K
  tokens in MCP-server mode versus ~59K in its CLI mode, attributing the gap to "MCP sends 26
  tool schemas to the LLM (~100K+ tokens)." Self-reported, task/model unspecified, prompt caching
  ignored — but it is the token-overhead problem the single deep tool was designed to remove,
  measured by the other side.
- **Anthropic's published guidance**
  ([Writing effective tools for agents](https://www.anthropic.com/engineering/writing-tools-for-agents))
  recommends consolidating discrete operations into fewer comprehensive tools, and attributes a
  state-of-the-art SWE-bench result in part to "precise refinements to tool descriptions."
  Honest scale note: Anthropic's examples consolidate ~3 tools into 1; 25-into-1 extends the
  direction of the guidance, it is not restated by it. And the description-quality study found
  fixing descriptions raised task success a median +5.85pp — while increasing execution steps
  +67% median, which supports "richer schemas raise success" and complicates "richer schemas cut
  round trips."

## The road not taken, fairly stated

The headless C# path (EPPlus, ClosedXML) is not formula-blind:
[EPPlus](https://github.com/EPPlusSoftware/EPPlus/wiki/Formula-Calculation) "evaluates formulas
entirely in .NET" (478 built-in functions; the wiki page states the count without attributing it
to a version), and
[ClosedXML](https://github.com/ClosedXML/ClosedXML/blob/develop/docs/concepts/formula-calculation.rst)
walks a calculation chain with dirty tracking and cycle detection. Both are partial
re-implementations of Excel semantics — function subsets, no iterative calculation — which is the
fidelity gap that justifies live COM despite its support-status and lifecycle costs. A claim that
ClosedXML saves stale cached values by default was refuted (1-2) and is not relied on here.

## Evidence around the formula refusal (surfaced, not all adversarially verified)

The survey's verified tier established the *design contrast* (everyone else accepts model
formulas); the *empirical grounding* below was extracted from primary sources but fell outside
the 25-claim verification budget — treat as sourced leads, not confirmed findings:

- [SpreadsheetBench](https://spreadsheetbench.github.io/): best reported score 70.48% on 912
  questions derived from real-world scenarios. (The project page links an OpenReview paper but
  does not itself name the venue; a NeurIPS 2024 attribution was not confirmed from the page.)
- FLARE benchmark ([arXiv 2506.17330](https://arxiv.org/pdf/2506.17330)): "most models fail
  silently: they produce plausible outputs that conceal" errors — the exact failure mode the
  refusal-plus-verification stance targets.
- Microsoft Research
  ([Excel formula repair benchmark](https://www.microsoft.com/en-us/research/publication/benchmark-dataset-generation-and-evaluation-for-excel-formula-repair-with-llms/)):
  automated correction of semantic formula errors described as unsolved as of Aug 2025.
- [ICAEW's 2025 comparison](https://www.icaew.com/insights/viewpoints-on-the-news/2025/oct-2025/the-best-and-worst-genais-for-spreadsheets):
  best LLMs got roughly two-thirds of spreadsheet tasks right.
- A verified-tool-calls study ([arXiv 2608.02645](https://arxiv.org/html/2608.02645)): wrapping
  agent tool calls in postcondition verification held 100% task success under injected faults
  while the unverified baseline fell to 64% — the closest academic parallel to reopen-and-verify.
- Middle positions exist: SheetMind ([arXiv 2506.12339](https://arxiv.org/html/2506.12339v2))
  constrains generation to a closed BNF grammar of seven operations; Microsoft's SheetBrain
  ([arXiv 2510.19247](https://arxiv.org/pdf/2510.19247)) embraces model-written Python but runs
  it in a sandbox. The design space is refusal → grammar constraint → sandboxed embrace, and
  ExcelTask holds the strict end.

## What this survey could not settle

- **Whether reopen-and-verify / proof-of-exit exists anywhere else.** Nothing surveyed has it;
  the survey did not exhaustively rule out a counterpart.
- **The incumbent's real lifecycle model** — the session-daemon characterization failed
  verification 0-3, so the stateless-vs-session comparison against it is uncharacterized.
- **Google Sheets MCPs, official Microsoft MCP/Copilot work, Office.js, Graph workbook API,
  LibreOffice UNO** — not covered by surviving claims.
- **haris-musa's tool count** — a "33 tools" claim failed 0-3; its granular-catalog *pattern* is
  confirmed, its number is not.

## Refuted during verification — do not reuse

1. "ExcelMcp is a persistent session daemon over a named pipe" — 0-3.
2. "haris-musa exposes 33 granular tools" — 0-3.
3. "ClosedXML saves stale cached formula values by default" — 1-2.
