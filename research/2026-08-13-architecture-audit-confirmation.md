# Architecture audit confirmation — 2026-08-13

## Question

Do the three architecture-review candidates improve MCP token efficiency for complicated
tasks, end-to-end speed, or reliability?

## Method

Primary sources only:

- current source and tests in `src/` and `tests/`;
- `docs/ARCHITECTURE.md`, `docs/DIAGNOSTIC-TRACE.md`,
  `docs/INTERFACE-AB-STUDY.md`;
- field evidence in `docs/field-reports/`;
- commit `3c1b472` ("One implementation of the mutate-save-verify choreography").

Basic verification run sequentially:

```powershell
dotnet test tests\ExcelTask.Core.Tests\ExcelTask.Core.Tests.csproj --no-restore --filter "RunType!=OnDemand" -p:NuGetAudit=false
# Passed 138, Failed 0, Skipped 0

dotnet test tests\ExcelTask.Excel.Tests\ExcelTask.Excel.Tests.csproj --no-restore --filter "RunType!=OnDemand" -p:NuGetAudit=false
# Passed 88, Failed 0, Skipped 0

dotnet test tests\ExcelTask.McpServer.Tests\ExcelTask.McpServer.Tests.csproj --no-restore --filter "RunType!=OnDemand" -p:NuGetAudit=false
# Passed 19, Failed 0, Skipped 0

dotnet test tests\ExcelTask.Excel.Tests\ExcelTask.Excel.Tests.csproj --no-restore --filter "FullyQualifiedName~FormulaCopyPromotionFailureIsReportedAsUnknown" -p:NuGetAudit=false
# Before the refactor: Failed (Expected Unknown, Actual Partial)
# After the refactor: Passed 1, Failed 0, Skipped 0

# The full Excel on-demand set was then run serially: 39 passed, 0 failed,
# 0 skipped; the new direct-formula test passed in 16 seconds.

dotnet test tests\ExcelTask.Excel.Tests\ExcelTask.Excel.Tests.csproj --no-restore --filter "FullyQualifiedName~ExcelWorkbookRuntimeIntegrationTests" -p:NuGetAudit=false
# Passed 39, Failed 0, Skipped 0 (serial desktop Excel integration class)
```

The first three are non-desktop slices. The last command is the complete opt-in
real-Excel integration class and was run serially.

## Findings

### 1. Finish the private mutation-transaction module

**Structural claim: confirmed and implemented for formula/exhibit Apply, with an
important qualification.**

`ExcelWorkbookRuntime.Mutation.cs:74-224` owns the generic mutation lifecycle:
input rechecks, staging save, close/prove, file-lock check, reopen verification,
promotion, and status classification. Before this change, formula execution in
`ExcelWorkbookRuntime.cs` and macro execution in
`ExcelWorkbookRuntime.Macro.cs` contained parallel lifecycle code. Formula now
delegates that tail to `ExecuteMutation`; macro remains bespoke. A static source
check found `SaveAs`, `CloseAndProve`, lock checks, promotion, and `Unknown`
classification in all three paths before the extraction.

However, commit `3c1b472` explicitly kept formula and macro paths bespoke:
formula has two-phase revalidation, while macro has hash preconditions,
dialog containment, and an abandon path. The right candidate is therefore a
narrow private transaction seam with operation-specific adapters, not a
generic executor that absorbs those semantics.

**Token efficiency:** no direct improvement. The seam is private and changes
neither the MCP schema nor receipt size. The measured token/call wins come from
the existing deep-call shape and model-visible guardrails:
`docs/field-reports/2026-08-10-comparison/CLIENT-SESSIONS.md:11-42` reports
-73.9% total tokens and -83.8% MCP calls versus the comparison server, while
`docs/INTERFACE-AB-STUDY.md:359-365` attributes a 32% call saving to
stating enforced limits in the schema. A reliability fix could indirectly
avoid retry conversations, but that effect is unmeasured.

**End-to-end speed:** no immediate improvement proven. The diagnostic trace
measures 5.1 seconds of 5.6 seconds (92%) in owned-Excel teardown and
reopen verification, and calls both load-bearing
(`docs/DIAGNOSTIC-TRACE.md:80-99`). A transaction refactor does not remove
those steps. The field comparison also shows active MCP execution was
48.0s vs 42.6s (+12.7%) even though prompt-to-completion was faster because
model coordination was reduced (`docs/field-reports/2026-08-10-comparison/CLIENT-SESSIONS.md:19-42`).

**Reliability:** high potential, medium confidence. Centralizing the repeated
lifecycle can prevent drift in cleanup, promotion, and `Unknown` handling.
The 3c1b472 commit is direct evidence that the generic copies had already
drifted in six ways before consolidation. Confidence is medium because
formula and macro semantics are intentionally different. The seam should still
be covered by a broader cross-family transaction test matrix as macro behavior
evolves; it should not be sold as a speed or token win.

**Implementation status:** the worthwhile narrow seam is now implemented. The
formula/exhibit Apply path delegates its save/cleanup/lock/reopen/promotion
tail to `ExecuteMutation`; formula preflight, two-phase revalidation, mutation,
recalculation, and verification remain local. Macro remains bespoke. A real
Excel regression holds the existing output locked at `copy-promotion` and
proves the shared path returns `Unknown` with a failed `formula-save` check.

### 2. Separate supervision policy from worker transport

**Structural claim: confirmed, but conditional.**

`SupervisedWorkbookRuntime.cs:60-272` mixes acceptance, deadlines,
cancellation, `Unknown` classification, identity authorization, cleanup
sweeps, and result handling. The same file contains the JSON-lines process
client/session at `361-702`. Existing `IWorkbookWorkerClient` and
`IWorkbookWorkerSession` seams, plus fakes in
`tests/ExcelTask.Excel.Tests/SupervisedWorkbookRuntimeTests.cs`, already
make much of the policy testable. There is only one production transport.

**Token efficiency:** neutral. A private policy/transport split does not alter
the one `excel_task` schema or its descriptions.

**End-to-end speed:** neutral today. The worker is still short-lived per
operation and the architecture explicitly defers warm-process reuse until
work-computer measurements justify it (`docs/ARCHITECTURE.md:119-131`).
Separating the code does not reduce process startup, teardown, or verification.

**Reliability:** medium potential, conditional. It could make policy tests
independent of JSON framing and reduce risk when a second transport or
protocol version arrives. With one production adapter and existing seams,
the immediate leverage is limited; splitting now may reduce locality. Revisit
when transport churn or a second adapter is real.

### 3. Isolate request compilation from Task Engine orchestration

**Structural claim: confirmed, but the defer decision is correct.**

`ExcelTaskEngine.cs:23-215` owns orchestration; normalization and the
twelve-operation dispatch live at `362-472` and `696-790`. But
`OperationCatalog.cs:16-64` already centralizes payload mapping with an
exhaustive switch, so adding a private compiler would create another seam
without a second adapter or variation.

**Token efficiency:** neutral to very low indirect benefit. Internal
organization cannot change model-visible schema bytes. The interface study
shows the direct gains came from wording and guardrails, not compiler
placement (`docs/INTERFACE-AB-STUDY.md:27-60` and `:618-667`).

**End-to-end speed:** neutral. Request normalization is local CPU work; the
measured cost is Excel lifecycle and model coordination, not this dispatch
(`docs/DIAGNOSTIC-TRACE.md:80-99`;
`docs/field-reports/2026-08-10-comparison/CLIENT-SESSIONS.md:29-42`).

**Reliability:** low now, possible future benefit. The closed-union mapping is
compile-time checked and the current non-desktop suite passed 245 tests.
Split it only if operation churn makes normalization changes repeatedly
collide with orchestration; do not introduce a generic registry preemptively.

## Verdict

The audit's three structural observations hold. The impact ranking is:

1. **Mutation transaction:** implemented for reliability; no token or speed
   gain is claimed without a separately measured optimization.
2. **Supervision/transport:** leave as-is until a second transport or protocol
   change creates real leverage.
3. **Request compiler:** defer; current exhaustive catalog and tests provide
   better locality than another private module.

No candidate is currently a demonstrated token-efficiency or end-to-end-speed
optimization. The existing evidence says those outcomes come from fewer model
round trips, explicit schema guardrails, and measured lifecycle decisions—not
from these private refactors alone.

## Follow-up implementation: direct formula writes

The user-requested breadth gap was implemented after this audit. The closed
operation union now includes `WriteWorksheetFormulas`, separate from the
constants-only `WriteWorksheetValues` operation. The new path accepts at most
200 single-cell A1 formulas, 8,192 characters per formula, one 400-cell span,
and 768 KiB of UTF-8 formula text per request. It reads formulas back before
saving, uses the shared `ExecuteMutation` transaction, and verifies them after
reopening. Formula text is never placed in receipts. Core normalization,
MCP-schema round-trip, desktop-Excel integration, and real MCP-boundary tests
cover the operation; managed-work-computer field execution remains the next
release gate.
