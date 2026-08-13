# Mutation Transaction Refactor Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Move the formula/exhibit mutation lifecycle onto the existing private mutation transaction module so save, cleanup, lock, reopen verification, promotion, and uncertain-outcome classification have one implementation.

**Architecture:** Keep `ExcelWorkbookRuntime.Mutation.cs` as the private transaction module. Formula/exhibit code remains responsible for workbook-specific preflight, formula analysis, two-phase revalidation, mutation, recalculation, and verification. Macro editing remains bespoke because its hash, VBA dialog, and abandoned-process semantics are materially different. No public interface, MCP schema, or model-facing receipt shape changes.

**Tech Stack:** C#/.NET 10, desktop Excel COM, xUnit, existing serial on-demand integration tests.

---

### Task 1: Lock the lifecycle error contract with a regression test

**Files:**
- Modify: `tests/ExcelTask.Excel.Tests/ExcelWorkbookRuntimeIntegrationTests.cs`
- Modify: `tests/ExcelTask.Excel.Tests/RecordingRuntimeObserver.cs` only if a lock-holding observer helper is needed

- [x] Add an on-demand integration test that creates a valid target workbook, creates an existing copy output, enables overwrite, and uses an observer that opens the output with `FileShare.None` when the `copy-promotion` phase begins.
- [x] Assert the formula copy returns `Unknown`, includes a failed `formula-save` check, and leaves the locked output present for cleanup.
- [x] Run only this test with `dotnet test tests\\ExcelTask.Excel.Tests\\ExcelTask.Excel.Tests.csproj --no-restore --filter "FullyQualifiedName~FormulaCopyPromotionFailureIsReportedAsUnknown"`; it failed before the extraction (`Expected: Unknown`, `Actual: Partial`) and passed after it.

### Task 2: Route formula/exhibit execution through `ExecuteMutation`

**Files:**
- Modify: `src/ExcelTask.Excel/ExcelWorkbookRuntime.cs`
- Modify: `src/ExcelTask.Excel/ExcelWorkbookRuntime.Mutation.cs` only for the minimal save callback/status data needed by formula/exhibit
- Test: `tests/ExcelTask.Excel.Tests/ExcelWorkbookRuntimeIntegrationTests.cs`

- [x] Preserve the existing input checks, reference workbook behavior, preflight, formula analysis, Plan early exit, formula revalidation, isolated-target revalidation, copy/exhibit mutation, recalculation, and formula-specific saved-workbook verifier.
- [x] Replace the formula path's hand-written save/cleanup/lock/reopen/promotion/catch/finally tail with `ExecuteMutation`.
- [x] Add the smallest private transaction hook for formula's normal `Save`/`SaveAs` choice and its existing verification delegate; do not add a generic operation registry.
- [x] Keep macro editing outside this seam, including `MacroCompilationException` abandonment and macro receipt behavior.
- [x] Run the failing regression test and confirm it passes.

### Task 3: Verify behavior and documentation

**Files:**
- Modify: `CHANGELOG.md`
- Modify: `docs/ARCHITECTURE.md`
- Modify: `research/2026-08-13-architecture-audit-confirmation.md`

- [x] Run the focused Excel integration test, the existing formula/mutation tests, and the serial integration class (38 passed).
- [x] Run the three non-desktop slices sequentially: Core (138 passed), Excel (88 passed), and MCP-server (19 passed) with `RunType!=OnDemand`.
- [x] Run `git diff --check`, inspect the complete diff, and confirm no MCP schema or public contract changed.
- [x] Record the implemented scope honestly: reliability-focused lifecycle consolidation; no measured token or speed improvement; macro and transport/request-compiler candidates remain deferred.
