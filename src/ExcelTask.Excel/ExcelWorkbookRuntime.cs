using System.Collections.Concurrent;
using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.Versioning;
using System.Security.Cryptography;
using System.Text;
using ExcelTask.Core;

namespace ExcelTask.Excel;

/// <summary>Runs all desktop Excel automation on one message-pumping STA thread.</summary>
[SupportedOSPlatform("windows")]
public sealed partial class ExcelWorkbookRuntime : IWorkbookRuntime, IDisposable
{
    private readonly StaComDispatcher _dispatcher = new();
    private readonly IExcelWorkbookRuntimeObserver _observer;
    private bool _disposed;

    public ExcelWorkbookRuntime() : this(NullExcelWorkbookRuntimeObserver.Instance) { }

    internal ExcelWorkbookRuntime(IExcelWorkbookRuntimeObserver observer) =>
        _observer = observer ?? throw new ArgumentNullException(nameof(observer));

    public Task<WorkbookInspection> InspectAsync(WorkbookInspectionRequest request, CancellationToken cancellationToken)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        ArgumentNullException.ThrowIfNull(request);
        return _dispatcher.InvokeAsync(() => InspectCore(request, _observer), cancellationToken);
    }

    public Task<WorkbookExecutionOutcome> ExecuteAsync(ExcelTaskPlan plan, CancellationToken cancellationToken)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        ArgumentNullException.ThrowIfNull(plan);
        return _dispatcher.InvokeAsync(() => ExecuteCore(plan, _observer), cancellationToken);
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        _dispatcher.Dispose();
    }

    private static WorkbookInspection InspectCore(WorkbookInspectionRequest request, IExcelWorkbookRuntimeObserver observer)
    {
        observer.OnPhase("inspection");
        // A path that cannot be used is a finding, not a failure. Thrown, it reached the caller as
        // "Workbook inspection could not be completed before execution" - an infrastructure-sounding
        // answer to the most ordinary user error there is, a mistyped path.
        try
        {
            var targetPath = WorkbookRuntimeHelpers.NormalizePath(request.TargetWorkbookPath);
            var referencePath = string.IsNullOrWhiteSpace(request.ReferenceWorkbookPath) ? null : WorkbookRuntimeHelpers.NormalizePath(request.ReferenceWorkbookPath);
            if (request.TargetMustExist)
            {
                WorkbookRuntimeHelpers.EnsureReadableWorkbook(targetPath, "Target workbook");
            }
            else
            {
                WorkbookRuntimeHelpers.EnsureCreatableWorkbook(targetPath);
            }

            if (referencePath is not null) WorkbookRuntimeHelpers.EnsureReadableWorkbook(referencePath, "Reference workbook");
            var copyOutputExists = request.Save == SaveMode.Copy &&
                                   !string.IsNullOrWhiteSpace(request.OutputWorkbookPath) &&
                                   File.Exists(WorkbookRuntimeHelpers.NormalizePath(request.OutputWorkbookPath));
            if (request.Save == SaveMode.Copy)
            {
                WorkbookRuntimeHelpers.EnsureWritableCopyOutput(request.OutputWorkbookPath);
            }
            var targetIsOpen = RotWorkbookLocator.ContainsPath(targetPath);
            return new WorkbookInspection(
                TargetIsOpen: targetIsOpen,
                CopyOutputExists: copyOutputExists,
                OpenWorkbookDescription: targetIsOpen ? "The exact target workbook is open." : null,
                Checks: [new TaskCheck("target-path", true, "Target path was normalized and checked against the running object table.")]);
        }
        catch (InvalidOperationException exception)
        {
            return new WorkbookInspection(
                TargetIsOpen: false,
                Checks: [new TaskCheck("workbook-inputs", false, exception.Message)],
                InfeasibleReason: exception.Message);
        }
    }

    /// <summary>
    /// One bounded line naming the phase that failed and the fault behind it. Without this the
    /// caller only ever saw "rejected before changes were attempted", which names neither, and a
    /// COM fault could only be diagnosed by rebuilding the worker to print its own stderr - which
    /// the supervisor drains and discards. Excel's own message is kept short and single-line; it
    /// describes the fault, never workbook contents.
    /// </summary>
    private static string DescribeFailure(string phase, Exception exception)
    {
        var fault = exception is TargetInvocationException { InnerException: not null } wrapper ? wrapper.InnerException! : exception;
        var message = fault.Message.ReplaceLineEndings(" ").Trim();
        if (message.Length > 60) message = string.Concat(message.AsSpan(0, 60), "...");
        return string.Create(CultureInfo.InvariantCulture, $"Failed in phase '{phase}': {fault.GetType().Name}: {message}");
    }

    private static WorkbookExecutionOutcome ExecuteCore(ExcelTaskPlan plan, IExcelWorkbookRuntimeObserver observer)
    {
        if (plan.Request.Operation.Kind == ExcelOperationKind.EditMacroProcedure)
        {
            return ExecuteMacroCore(plan, observer);
        }

        // Dispatched before the write-oriented gates below: neither reading operation has a save to
        // confirm, so the overwrite and copy-output gates would reject them for lacking a policy
        // they never needed.
        if (plan.Request.Operation.Kind == ExcelOperationKind.AuditWorkbookFlows)
        {
            return ExecuteAuditCore(plan, observer);
        }

        // First in spirit as well as in order: the one operation that never starts Excel at all.
        if (plan.Request.Operation.Kind == ExcelOperationKind.ScanWorkbookStructure)
        {
            return ExecuteScanCore(plan, observer);
        }

        if (plan.Request.Operation.Kind == ExcelOperationKind.ReadWorksheetRange)
        {
            return ExecuteReadCore(plan, observer);
        }

        // Dispatched before the shared gates too: it writes, but it carries its own save and verify
        // rather than the formula plan the code below assumes.
        if (plan.Request.Operation.Kind == ExcelOperationKind.WriteWorksheetValues)
        {
            if (plan.Request.WorkbookBinding == WorkbookBinding.UseOpen && plan.Request.Save == SaveMode.Copy)
            {
                return new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected,
                    "Copy saves are not supported when applying to a live workbook.",
                    Checks: [new TaskCheck("live-copy-save", false, "Use the confirmed same-file save mode or isolated copy mode.")]);
            }

            if (plan.Request.Mode == ExcelTaskMode.Apply && plan.Request.Save == SaveMode.Same && !plan.Request.OverwriteConfirmed)
            {
                return new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected,
                    "Same-file saves require explicit overwrite confirmation.",
                    Checks: [new TaskCheck("same-file-overwrite", false, "Apply with save Same requires overwrite confirmation.")]);
            }

            return ExecuteWriteCore(plan, observer);
        }

        // Same reason as the write above: each carries its own save, verification and gates, and
        // none of them wants the formula plan the shared path below builds.
        if (plan.Request.Operation.Kind is ExcelOperationKind.FindReplace or ExcelOperationKind.SetNumberFormat)
        {
            if (plan.Request.WorkbookBinding == WorkbookBinding.UseOpen && plan.Request.Save == SaveMode.Copy)
            {
                return new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected,
                    "Copy saves are not supported when applying to a live workbook.",
                    Checks: [new TaskCheck("live-copy-save", false, "Use the confirmed same-file save mode or isolated copy mode.")]);
            }

            if (plan.Request.Mode == ExcelTaskMode.Apply && plan.Request.Save == SaveMode.Same && !plan.Request.OverwriteConfirmed)
            {
                return new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected,
                    "Same-file saves require explicit overwrite confirmation.",
                    Checks: [new TaskCheck("same-file-overwrite", false, "Apply with save Same requires overwrite confirmation.")]);
            }

            return plan.Request.Operation.Kind == ExcelOperationKind.FindReplace
                ? ExecuteFindReplaceCore(plan, observer)
                : ExecuteNumberFormatCore(plan, observer);
        }

        // A creation writes the target it names and is refused a copy destination during validation,
        // so the overwrite gate below would ask it to confirm replacing a file it has already
        // guaranteed does not exist.
        if (plan.Request.Operation.Kind == ExcelOperationKind.Create)
        {
            return ExecuteCreateCore(plan, observer);
        }

        try
        {
            WorkbookRuntimeHelpers.EnsureReadableWorkbook(WorkbookRuntimeHelpers.NormalizePath(plan.Request.TargetWorkbookPath), "Target workbook");
            if (NeedsReferenceWorkbook(plan.Request))
                WorkbookRuntimeHelpers.EnsureReadableWorkbook(WorkbookRuntimeHelpers.NormalizePath(plan.Request.Operation.CopyExhibit!.ReferenceWorkbookPath), "Reference workbook");
            if (plan.Request.Save == SaveMode.Copy) WorkbookRuntimeHelpers.EnsureWritableCopyOutput(plan.Request.OutputWorkbookPath);
            // A read-only target found here is a clean rejection; found after the save it is Unknown.
            if (plan.Request.Save == SaveMode.Same && plan.Request.Mode == ExcelTaskMode.Apply)
                WorkbookRuntimeHelpers.EnsureWritableSameTarget(WorkbookRuntimeHelpers.NormalizePath(plan.Request.TargetWorkbookPath));
        }
        catch (InvalidOperationException exception)
        {
            return new WorkbookExecutionOutcome(
                ExcelTaskStatus.Rejected,
                "Workbook inputs cannot be safely executed.",
                Checks: [new TaskCheck("workbook-inputs", false, exception.Message)]);
        }

        if (plan.Request.WorkbookBinding == WorkbookBinding.UseOpen && plan.Request.Save == SaveMode.Copy)
        {
            return new WorkbookExecutionOutcome(
                ExcelTaskStatus.Rejected,
                "Copy saves are not supported when applying to a live workbook.",
                Checks: [new TaskCheck("live-copy-save", false, "Use the confirmed same-file save mode or isolated copy mode.")]);
        }

        if (plan.Request.Mode == ExcelTaskMode.Apply && plan.Request.Save == SaveMode.Same && !plan.Request.OverwriteConfirmed)
        {
            return new WorkbookExecutionOutcome(
                ExcelTaskStatus.Rejected,
                "Same-file saves require explicit overwrite confirmation.",
                Checks: [new TaskCheck("same-file-overwrite", false, "Apply with save Same requires overwrite confirmation.")]);
        }

        ExcelSession? session = null;
        PendingVerification? pendingVerification = null;
        var mutationAttempted = false;
        var verified = false;
        var changes = new List<TaskChange>();
        var checks = new List<TaskCheck>();
        FormulaExecutionPlan? formulaPlan = null;
        var phase = "input-validation";
        void SetPhase(string value)
        {
            phase = value;
            observer.OnPhase(value);
        }
        SetPhase(phase);
        var savedPath = WorkbookRuntimeHelpers.NormalizePath(plan.Request.Save == SaveMode.Copy
            ? plan.Request.OutputWorkbookPath ?? throw new InvalidOperationException("Copy output path is required.")
            : plan.Request.TargetWorkbookPath);
        string? stagingPath = null;

        try
        {
            if (plan.Request.Mode == ExcelTaskMode.Apply && plan.Request.Save == SaveMode.Copy && File.Exists(savedPath) && !plan.Request.OverwriteConfirmed)
            {
                return new WorkbookExecutionOutcome(
                    ExcelTaskStatus.Rejected,
                    "The copy output already exists and was not authorized for overwrite.",
                    Checks: [new TaskCheck("copy-output", false, "Existing output requires overwrite confirmation.")]);
            }

            if (plan.Request.Mode == ExcelTaskMode.Apply && plan.Request.WorkbookBinding == WorkbookBinding.Isolated && plan.Request.Save == SaveMode.Same)
            {
                using var openTarget = RotWorkbookLocator.Find(WorkbookRuntimeHelpers.NormalizePath(plan.Request.TargetWorkbookPath));
                if (openTarget is not null)
                {
                    return new WorkbookExecutionOutcome(
                        ExcelTaskStatus.Rejected,
                        "The target workbook is already open and cannot be safely overwritten in isolated mode.",
                        Checks: [new TaskCheck("isolated-target", false, "Use the exact open-workbook binding or save a copy.")]);
                }
            }

            // Started here rather than after the save, so its launch runs alongside the analysis,
            // the writes, and the save instead of after them. Only a mutating run ever verifies,
            // so only a mutating run pays for one. Disposed unconditionally in the outer finally.
            if (plan.Request.Mode == ExcelTaskMode.Apply) pendingVerification = PendingVerification.Begin(observer);

            SetPhase("session-open");
            session = ExcelSession.Open(plan.Request, observer, readOnlyTarget: plan.Request.Mode == ExcelTaskMode.Plan);
            SetPhase("preflight");
            var preflight = Preflight(session, plan.Request.Operation);
            checks.AddRange(preflight.Checks);
            if (!preflight.IsFeasible)
            {
                return ExcelSession.CloseAndProve(ref session, "preflight", checks)
                    ?? new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected, "Workbook preflight did not permit the requested formula operation.", Checks: checks);
            }

            SetPhase("formula-analysis");
            formulaPlan = AnalyzeFormulaPlan(session, plan.Request.Operation);
            checks.Add(new TaskCheck("formula-plan", true,
                $"Planned {formulaPlan.Repairs.Count} formula changes across {formulaPlan.RangeResults.Count} requested range targets; fingerprint {formulaPlan.Fingerprint}."));

            if (plan.Request.Mode == ExcelTaskMode.Plan)
            {
                return ExcelSession.CloseAndProve(ref session, "planning", checks)
                    ?? new WorkbookExecutionOutcome(
                        ExcelTaskStatus.Planned,
                        "Workbook plan is feasible; no Excel changes were made.",
                        CreateFormulaChanges(formulaPlan, planning: true),
                        checks);
            }

            SetPhase("formula-revalidation");
            // The same preflight, deliberately repeated: it is evidence gathered immediately before
            // mutation, not a cached result from before the plan was built.
            var revalidatedPreflight = Preflight(session, plan.Request.Operation);
            if (!revalidatedPreflight.IsFeasible || !FormulaPlansEqual(formulaPlan, AnalyzeFormulaPlan(session, plan.Request.Operation)))
            {
                checks.Add(new TaskCheck("formula-revalidation", false, "Workbook formula evidence changed before mutation; no changes were made."));
                return ExcelSession.CloseAndProve(ref session, "revalidation", checks)
                    ?? new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected, "Workbook evidence changed before mutation; no changes were made.", Checks: checks);
            }

            if (RotWorkbookLocator.RequiresPreMutationIsolatedSameApplyRevalidation(
                    plan.Request.Mode,
                    plan.Request.WorkbookBinding,
                    plan.Request.Save) &&
                session.HasExternalTargetOpen(WorkbookRuntimeHelpers.NormalizePath(plan.Request.TargetWorkbookPath)))
            {
                checks.Add(new TaskCheck("isolated-target-revalidation", false,
                    "The exact target workbook was opened in another Excel application before mutation; no changes were made."));
                return ExcelSession.CloseAndProve(ref session, "isolated target revalidation", checks)
                    ?? new WorkbookExecutionOutcome(
                        ExcelTaskStatus.Rejected,
                        "The target workbook was opened before isolated same-file apply could begin; no changes were made.",
                        Checks: checks);
            }

            var operation = plan.Request.Operation;
            var worksheetName = formulaPlan.WorksheetName;
            if (operation.Kind == ExcelOperationKind.CopyExhibit)
            {
                SetPhase("worksheet-copy");
                mutationAttempted = true;
                CopyReferenceWorksheet(session, operation.CopyExhibit!.ReferenceWorksheet, operation.CopyExhibit.NewWorksheetName, SetPhase);
                changes.Add(new TaskChange("worksheet-copy", worksheetName, "Copied the requested reference worksheet."));
            }

            SetPhase(operation.Kind == ExcelOperationKind.ExtendFormulaSeries ? "formula-extension" : "formula-repair");
            ApplyFormulaWrites(session, formulaPlan, () => mutationAttempted = true);
            changes.AddRange(CreateFormulaChanges(formulaPlan, planning: false));
            checks.Add(new TaskCheck("formula-change-count", true,
                $"Applied {formulaPlan.Repairs.Count} planned formula changes; fingerprint {formulaPlan.Fingerprint}."));

            var noFormulaChanges = operation.Kind != ExcelOperationKind.CopyExhibit && formulaPlan.Repairs.Count == 0;
            if (noFormulaChanges && plan.Request.Save == SaveMode.Same)
            {
                SetPhase("no-change-cleanup");
                var noChangeFailure = ExcelSession.CloseAndProve(ref session, "finding no changes", checks, changes);
                if (noChangeFailure is not null) return noChangeFailure;

                checks.Add(new TaskCheck("no-formula-changes", true, "No formula changes were required; Excel was not recalculated or saved."));
                return new WorkbookExecutionOutcome(ExcelTaskStatus.Completed, "Workbook analysis found no formula changes; no Excel changes were made.", changes, checks);
            }

            if (!noFormulaChanges)
            {
                SetPhase("recalculate");
                Invoke(session.Application, "CalculateFull");
            }
            SetPhase("save");
            mutationAttempted = true;
            if (plan.Request.Save == SaveMode.Copy)
            {
                stagingPath = WorkbookRuntimeHelpers.CreateStagingPath(savedPath, plan.TaskId);
                observer.OnStagingPathCreated(stagingPath);
                Invoke(session.TargetWorkbook, "SaveAs", stagingPath);
            }
            else
            {
                Invoke(session.TargetWorkbook, "Save");
            }

            checks.Add(new TaskCheck("save", true, "Excel completed the requested save operation."));
            SetPhase("primary-cleanup");
            var saveCleanupFailure = ExcelSession.CloseAndProve(ref session, "the primary save", checks, changes);
            if (saveCleanupFailure is not null)
            {
                AddStagingCleanupCheck(stagingPath, checks);
                return saveCleanupFailure with { RetryReason = "Inspect workbook state before retrying." };
            }

            if (plan.Request.WorkbookBinding != WorkbookBinding.UseOpen && !WorkbookRuntimeHelpers.CanOpenExclusively(stagingPath ?? savedPath))
            {
                checks.Add(new TaskCheck("file-lock", false, "The saved workbook remained locked after owned Excel cleanup."));
                AddStagingCleanupCheck(stagingPath, checks);
                return new WorkbookExecutionOutcome(
                    ExcelTaskStatus.Unknown,
                    "Excel saved the workbook, but the owned Excel session did not release its file lock for verification.",
                    changes,
                    checks,
                    CanRetry: false,
                    RetryReason: "Inspect the Excel process and file lock before retrying.");
            }

            var verificationPath = stagingPath ?? savedPath;
            SetPhase("reopen-verification");
            if (!VerifySavedWorkbook(verificationPath, worksheetName, formulaPlan.Repairs, pendingVerification!, observer, out var verificationCheck))
            {
                checks.Add(verificationCheck);
                AddStagingCleanupCheck(stagingPath, checks);
                return new WorkbookExecutionOutcome(
                    ExcelTaskStatus.Unknown,
                    "Excel saved the workbook, but reopen verification did not confirm all requested changes.",
                    changes,
                    checks,
                    CanRetry: false,
                    RetryReason: "Inspect the saved workbook before attempting another apply operation.");
            }

            checks.Add(verificationCheck);
            if (noFormulaChanges)
            {
                checks.Add(new TaskCheck("no-formula-changes", true, "No formula changes were required; Excel was not recalculated."));
            }
            verified = true;
            if (stagingPath is not null)
            {
                SetPhase("copy-promotion");
                WorkbookRuntimeHelpers.PromoteStaging(stagingPath, savedPath, plan.Request.OverwriteConfirmed);
                stagingPath = null;
                changes.Add(new TaskChange("copy-promotion", "workbook", "Promoted the verified staging workbook to the requested output path."));
            }
            return new WorkbookExecutionOutcome(ExcelTaskStatus.Completed, "Workbook changes were saved and verified after reopening.", changes, checks);
        }
        catch (Exception exception)
        {
            var ownedCleanupFailed = false;
            if (session is not null)
            {
                try { ownedCleanupFailed = !session.Close(); }
                catch (Exception cleanupException) when (ComAccess.IsComFailure(cleanupException)) { ownedCleanupFailed = true; }
                session = null;
            }
            var stagingCleanupFailed = stagingPath is not null && !WorkbookRuntimeHelpers.TryDeleteStaging(stagingPath);
            if (!stagingCleanupFailed) stagingPath = null;
            var status = ownedCleanupFailed || stagingCleanupFailed ? ExcelTaskStatus.Unknown : verified ? ExcelTaskStatus.Partial : mutationAttempted ? ExcelTaskStatus.Unknown : ExcelTaskStatus.Rejected;
            checks.Add(new TaskCheck("execution", false, mutationAttempted
                ? "Excel execution did not complete during the current phase; the change was not fully verified."
                : "Excel did not complete the requested read-only preflight."));
            checks.Add(new TaskCheck("failure-detail", false, DescribeFailure(phase, exception)));
            if (ownedCleanupFailed) checks.Add(new TaskCheck("owned-process-exit", false, "The owned Excel process did not exit during cleanup."));
            if (stagingCleanupFailed) checks.Add(new TaskCheck("staging-cleanup", false, "A staging workbook could not be deleted; inspect the output directory before retrying."));
            return new WorkbookExecutionOutcome(
                status,
                verified
                    ? "Workbook changes were verified, but final delivery did not complete."
                    : mutationAttempted ? "Excel attempted workbook changes, but their final state is unknown." : "Workbook execution was rejected before changes were attempted.",
                changes,
                checks,
                CanRetry: false,
                RetryReason: mutationAttempted ? "Inspect workbook state before retrying." : "Correct the workbook preflight issue before retrying.");
        }
        finally
        {
            session?.Close();
            // Unconditional: every early return above - rejection, preflight failure, exception,
            // a plan that turned out to need no changes - leaves a started Excel that nothing else
            // will close. This one line is the whole containment for the pre-launch.
            pendingVerification?.Dispose();
            if (stagingPath is not null) _ = WorkbookRuntimeHelpers.TryDeleteStaging(stagingPath);
        }
    }

    /// <summary>
    /// Checks that the worksheets this operation needs exist and that its new name is free. Stated
    /// once because it runs twice - before planning and again immediately before mutation - and two
    /// copies of the same switch could drift into checking different things at those two moments.
    /// </summary>
    private static WorksheetCopyPreflight Preflight(ExcelSession session, NormalizedExcelOperation operation) => operation.Kind switch
    {
        ExcelOperationKind.RepairExistingWorksheet => PreflightWorksheetExists(session, operation.RepairExistingWorksheet!.WorksheetName),
        ExcelOperationKind.ExtendFormulaSeries => PreflightWorksheetExists(session, operation.ExtendFormulaSeries!.WorksheetName),
        _ => PreflightWorksheetCopy(session, operation.CopyExhibit!.ReferenceWorksheet, operation.CopyExhibit.NewWorksheetName)
    };

    // Late-bound access lives in ComAccess so the binding rules are stated once.
    private static object Get(object target, string member, params object?[] arguments) => ComAccess.Get(target, member, arguments);

    private static object? GetOrNull(object target, string member, params object?[] arguments) => ComAccess.GetOrNull(target, member, arguments);

    private static void Set(object target, string member, object? value) => ComAccess.Set(target, member, value);

    private static object? Invoke(object target, string member, params object?[] arguments) => ComAccess.Invoke(target, member, arguments);

    /// <summary>Resolves a collection entry regardless of how the owning object model binds Item.</summary>
    private static object Item(object collection, object index) => ComAccess.Item(collection, index);

}
