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
        var targetPath = WorkbookRuntimeHelpers.NormalizePath(request.TargetWorkbookPath);
        var referencePath = string.IsNullOrWhiteSpace(request.ReferenceWorkbookPath) ? null : WorkbookRuntimeHelpers.NormalizePath(request.ReferenceWorkbookPath);
        WorkbookRuntimeHelpers.EnsureReadableWorkbook(targetPath, "Target workbook");
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

    private static WorkbookExecutionOutcome ExecuteCore(ExcelTaskPlan plan, IExcelWorkbookRuntimeObserver observer)
    {
        if (plan.Request.Operation.Kind == ExcelOperationKind.EditMacroProcedure)
        {
            return ExecuteMacroCore(plan, observer);
        }

        // Dispatched before the write-oriented gates below: an audit has no save to confirm.
        if (plan.Request.Operation.Kind == ExcelOperationKind.AuditWorkbookFlows)
        {
            return ExecuteAuditCore(plan, observer);
        }

        try
        {
            WorkbookRuntimeHelpers.EnsureReadableWorkbook(WorkbookRuntimeHelpers.NormalizePath(plan.Request.TargetWorkbookPath), "Target workbook");
            if (NeedsReferenceWorkbook(plan.Request))
                WorkbookRuntimeHelpers.EnsureReadableWorkbook(WorkbookRuntimeHelpers.NormalizePath(plan.Request.Operation.CopyExhibit!.ReferenceWorkbookPath), "Reference workbook");
            if (plan.Request.Save == SaveMode.Copy) WorkbookRuntimeHelpers.EnsureWritableCopyOutput(plan.Request.OutputWorkbookPath);
        }
        catch (InvalidOperationException)
        {
            return new WorkbookExecutionOutcome(
                ExcelTaskStatus.Rejected,
                "Workbook inputs cannot be safely executed.",
                Checks: [new TaskCheck("workbook-inputs", false, "Workbook paths or output location are not ready for execution.")]);
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

            SetPhase("session-open");
            session = ExcelSession.Open(plan.Request, observer, readOnlyTarget: plan.Request.Mode == ExcelTaskMode.Plan);
            SetPhase("preflight");
            var preflight = plan.Request.Operation.Kind switch
            {
                ExcelOperationKind.RepairExistingWorksheet => PreflightWorksheetExists(session, plan.Request.Operation.RepairExistingWorksheet!.WorksheetName),
                ExcelOperationKind.ExtendFormulaSeries => PreflightWorksheetExists(session, plan.Request.Operation.ExtendFormulaSeries!.WorksheetName),
                _ => PreflightWorksheetCopy(session, plan.Request.Operation.CopyExhibit!.ReferenceWorksheet, plan.Request.Operation.CopyExhibit.NewWorksheetName)
            };
            checks.AddRange(preflight.Checks);
            if (!preflight.IsFeasible)
            {
                var preflightCleanupVerified = session.Close();
                session = null;
                if (!preflightCleanupVerified)
                {
                    checks.Add(new TaskCheck("owned-process-exit", false, "The owned preflight Excel process did not exit."));
                    return new WorkbookExecutionOutcome(ExcelTaskStatus.Unknown, "Workbook preflight could not prove owned Excel cleanup.", Checks: checks, CanRetry: false, RetryReason: "Inspect the owned Excel process before retrying.");
                }

                return new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected, "Workbook preflight did not permit the requested formula operation.", Checks: checks);
            }

            SetPhase("formula-analysis");
            formulaPlan = AnalyzeFormulaPlan(session, plan.Request.Operation);
            checks.Add(new TaskCheck("formula-plan", true,
                $"Planned {formulaPlan.Repairs.Count} formula changes across {formulaPlan.RangeResults.Count} requested range targets; fingerprint {formulaPlan.Fingerprint}."));

            if (plan.Request.Mode == ExcelTaskMode.Plan)
            {
                var planCleanupVerified = session.Close();
                session = null;
                if (!planCleanupVerified)
                {
                    checks.Add(new TaskCheck("owned-process-exit", false, "The read-only planning session did not prove owned Excel exited."));
                    return new WorkbookExecutionOutcome(
                        ExcelTaskStatus.Unknown,
                        "Workbook plan could not prove owned Excel cleanup.",
                        Checks: checks,
                        CanRetry: false,
                        RetryReason: "Inspect the owned Excel process before retrying.");
                }

                return new WorkbookExecutionOutcome(
                    ExcelTaskStatus.Planned,
                    "Workbook plan is feasible; no Excel changes were made.",
                    CreateFormulaChanges(formulaPlan, planning: true),
                    checks);
            }

            SetPhase("formula-revalidation");
            var revalidatedPreflight = plan.Request.Operation.Kind switch
            {
                ExcelOperationKind.RepairExistingWorksheet => PreflightWorksheetExists(session, plan.Request.Operation.RepairExistingWorksheet!.WorksheetName),
                ExcelOperationKind.ExtendFormulaSeries => PreflightWorksheetExists(session, plan.Request.Operation.ExtendFormulaSeries!.WorksheetName),
                _ => PreflightWorksheetCopy(session, plan.Request.Operation.CopyExhibit!.ReferenceWorksheet, plan.Request.Operation.CopyExhibit.NewWorksheetName)
            };
            if (!revalidatedPreflight.IsFeasible || !FormulaPlansEqual(formulaPlan, AnalyzeFormulaPlan(session, plan.Request.Operation)))
            {
                checks.Add(new TaskCheck("formula-revalidation", false, "Workbook formula evidence changed before mutation; no changes were made."));
                var cleanupVerified = session.Close();
                session = null;
                if (!cleanupVerified)
                {
                    checks.Add(new TaskCheck("owned-process-exit", false, "The owned revalidation Excel process did not exit."));
                    return new WorkbookExecutionOutcome(ExcelTaskStatus.Unknown, "Workbook revalidation could not prove owned Excel cleanup.", Checks: checks, CanRetry: false, RetryReason: "Inspect the owned Excel process before retrying.");
                }
                return new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected, "Workbook evidence changed before mutation; no changes were made.", Checks: checks);
            }

            if (RotWorkbookLocator.RequiresPreMutationIsolatedSameApplyRevalidation(
                    plan.Request.Mode,
                    plan.Request.WorkbookBinding,
                    plan.Request.Save) &&
                session.HasExternalTargetOpen(WorkbookRuntimeHelpers.NormalizePath(plan.Request.TargetWorkbookPath)))
            {
                checks.Add(new TaskCheck("isolated-target-revalidation", false,
                    "The exact target workbook was opened in another Excel application before mutation; no changes were made."));
                var cleanupVerified = session.Close();
                session = null;
                if (!cleanupVerified)
                {
                    checks.Add(new TaskCheck("owned-process-exit", false, "The owned Excel process did not exit after isolated target revalidation."));
                    return new WorkbookExecutionOutcome(ExcelTaskStatus.Unknown, "Workbook target revalidation could not prove owned Excel cleanup.", Checks: checks, CanRetry: false, RetryReason: "Inspect the owned Excel process before retrying.");
                }

                return new WorkbookExecutionOutcome(
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
                var noChangeCleanupVerified = session.Close();
                session = null;
                if (!noChangeCleanupVerified)
                {
                    checks.Add(new TaskCheck("owned-process-exit", false, "The owned no-change Excel process did not exit."));
                    return new WorkbookExecutionOutcome(
                        ExcelTaskStatus.Unknown,
                        "Workbook analysis found no changes, but owned Excel cleanup could not be verified.",
                        changes,
                        checks,
                        CanRetry: false,
                        RetryReason: "Inspect the owned Excel process before retrying.");
                }

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
            var primaryCleanupVerified = session.Close();
            session = null;
            if (!primaryCleanupVerified)
            {
                checks.Add(new TaskCheck("owned-process-exit", false, "The owned Excel process did not exit after the primary save."));
                AddStagingCleanupCheck(stagingPath, checks);
                return new WorkbookExecutionOutcome(ExcelTaskStatus.Unknown, "Workbook changes could not be verified after owned Excel cleanup.", changes, checks, CanRetry: false, RetryReason: "Inspect workbook state before retrying.");
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
            if (!VerifySavedWorkbook(verificationPath, worksheetName, formulaPlan.Repairs, observer, out var verificationCheck))
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
        catch (Exception)
        {
            var ownedCleanupFailed = false;
            if (session is not null)
            {
                try { ownedCleanupFailed = !session.Close(); }
                catch (Exception cleanupException) when (cleanupException is COMException or InvalidOperationException or TargetInvocationException) { ownedCleanupFailed = true; }
                session = null;
            }
            var stagingCleanupFailed = stagingPath is not null && !WorkbookRuntimeHelpers.TryDeleteStaging(stagingPath);
            if (!stagingCleanupFailed) stagingPath = null;
            var status = ownedCleanupFailed || stagingCleanupFailed ? ExcelTaskStatus.Unknown : verified ? ExcelTaskStatus.Partial : mutationAttempted ? ExcelTaskStatus.Unknown : ExcelTaskStatus.Rejected;
            checks.Add(new TaskCheck("execution", false, mutationAttempted
                ? "Excel execution did not complete during the current phase; the change was not fully verified."
                : "Excel did not complete the requested read-only preflight."));
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
            if (stagingPath is not null) _ = WorkbookRuntimeHelpers.TryDeleteStaging(stagingPath);
        }
    }

    private static object Get(object target, string member, params object?[] arguments) => target.GetType().InvokeMember(
        member,
        BindingFlags.GetProperty,
        null,
        target,
        arguments,
        CultureInfo.InvariantCulture) ?? throw new InvalidOperationException($"Excel did not return '{member}'.");

    private static void Set(object target, string member, object? value) => target.GetType().InvokeMember(
        member,
        BindingFlags.SetProperty,
        null,
        target,
        [value],
        CultureInfo.InvariantCulture);

    private static object? Invoke(object target, string member, params object?[] arguments) => target.GetType().InvokeMember(
        member,
        BindingFlags.InvokeMethod,
        null,
        target,
        arguments,
        CultureInfo.InvariantCulture);

}
