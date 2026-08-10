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
public sealed class ExcelWorkbookRuntime : IWorkbookRuntime, IDisposable
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

    private static void AddStagingCleanupCheck(string? stagingPath, List<TaskCheck> checks)
    {
        if (stagingPath is not null && !WorkbookRuntimeHelpers.TryDeleteStaging(stagingPath))
        {
            checks.Add(new TaskCheck("staging-cleanup", false, "A staging workbook could not be deleted; inspect the output directory before retrying."));
        }
    }

    private static bool NeedsReferenceWorkbook(NormalizedExcelTaskRequest request) =>
        request.Operation.Kind == ExcelOperationKind.CopyExhibit;

    private static WorksheetCopyPreflight PreflightWorksheetCopy(ExcelSession session, string referenceSheetName, string newSheetName)
    {
        using var references = new ComReferenceScope();
        try
        {
            var referenceSheets = references.Add(Get(session.ReferenceWorkbook, "Worksheets"));
            var targetSheets = references.Add(Get(session.TargetWorkbook, "Worksheets"));
            var referenceExists = WorksheetExists(referenceSheets, referenceSheetName, references);
            var destinationExists = WorksheetExists(targetSheets, newSheetName, references);
            var checks = new List<TaskCheck>
            {
                new("reference-worksheet", referenceExists, referenceExists ? "The requested reference worksheet is available." : "The requested reference worksheet is unavailable."),
                new("destination-worksheet", !destinationExists, destinationExists ? "The destination worksheet name is already in use." : "The destination worksheet name is available.")
            };
            return new WorksheetCopyPreflight(referenceExists && !destinationExists, checks);
        }
        catch (Exception exception) when (exception is COMException or TargetInvocationException or InvalidComObjectException)
        {
            return new WorksheetCopyPreflight(false, [new TaskCheck("worksheet-preflight", false, "Workbook worksheet feasibility could not be read.")]);
        }
    }

    private static WorksheetCopyPreflight PreflightWorksheetExists(ExcelSession session, string worksheetName)
    {
        using var references = new ComReferenceScope();
        try
        {
            var sheets = references.Add(Get(session.TargetWorkbook, "Worksheets"));
            var exists = WorksheetExists(sheets, worksheetName, references);
            return new WorksheetCopyPreflight(exists,
                [new TaskCheck("target-worksheet", exists,
                    exists ? "The requested target worksheet is available." : "The requested target worksheet is unavailable.")]);
        }
        catch (Exception exception) when (exception is COMException or TargetInvocationException or InvalidComObjectException)
        {
            return new WorksheetCopyPreflight(false, [new TaskCheck("target-worksheet", false, "Target worksheet feasibility could not be read.")]);
        }
    }

    private static bool WorksheetExists(object worksheets, string name, ComReferenceScope references)
    {
        var count = Convert.ToInt32(Get(worksheets, "Count"), CultureInfo.InvariantCulture);
        for (var index = 1; index <= count; index++)
        {
            var worksheet = references.Add(Get(worksheets, "Item", index));
            var worksheetName = Get(worksheet, "Name") as string;
            if (string.Equals(worksheetName, name, StringComparison.OrdinalIgnoreCase)) return true;
        }

        return false;
    }

    private static void CopyReferenceWorksheet(ExcelSession session, string referenceSheetName, string newSheetName, Action<string> updatePhase)
    {
        using var references = new ComReferenceScope();
        updatePhase("worksheet-copy-reference-sheets");
        var referenceSheets = references.Add(Get(session.ReferenceWorkbook, "Worksheets"));
        updatePhase("worksheet-copy-reference-sheet");
        var referenceSheet = references.Add(Get(referenceSheets, "Item", referenceSheetName));
        updatePhase("worksheet-copy-target-sheets");
        var targetSheets = references.Add(Get(session.TargetWorkbook, "Worksheets"));
        var count = Convert.ToInt32(Get(targetSheets, "Count"), CultureInfo.InvariantCulture);
        updatePhase("worksheet-copy-target-anchor");
        var afterSheet = references.Add(Get(targetSheets, "Item", count));
        updatePhase("worksheet-copy-copy");
        Invoke(referenceSheet, "Copy", Type.Missing, afterSheet);
        updatePhase("worksheet-copy-copied-sheet");
        var copiedSheet = references.Add(Get(targetSheets, "Item", count + 1));
        updatePhase("worksheet-copy-rename");
        Set(copiedSheet, "Name", newSheetName);
    }

    private static FormulaExecutionPlan AnalyzeFormulaPlan(ExcelSession session, NormalizedExcelOperation operation) => operation.Kind switch
    {
        ExcelOperationKind.CopyExhibit => AnalyzeFormulaRepairs(
            session.ReferenceWorkbook,
            operation.CopyExhibit!.ReferenceWorksheet,
            operation.CopyExhibit.RepairRanges,
            operation.Kind,
            operation.CopyExhibit.NewWorksheetName),
        ExcelOperationKind.RepairExistingWorksheet => AnalyzeFormulaRepairs(
            session.TargetWorkbook,
            operation.RepairExistingWorksheet!.WorksheetName,
            operation.RepairExistingWorksheet.Ranges,
            operation.Kind,
            operation.RepairExistingWorksheet.WorksheetName),
        ExcelOperationKind.ExtendFormulaSeries => AnalyzeFormulaExtension(session.TargetWorkbook, operation.ExtendFormulaSeries!),
        _ => throw new InvalidOperationException("The requested operation kind is unsupported.")
    };

    private static FormulaExecutionPlan AnalyzeFormulaRepairs(
        object workbook,
        string sourceWorksheetName,
        IReadOnlyList<FormulaRepairRange> ranges,
        ExcelOperationKind kind,
        string targetWorksheetName)
    {
        using var references = new ComReferenceScope();
        var worksheetCollection = references.Add(Get(workbook, "Worksheets"));
        var worksheet = references.Add(Get(worksheetCollection, "Item", sourceWorksheetName));
        var repairs = new List<ExpectedFormula>();
        var rangeResults = new List<RepairRangeResult>();

        foreach (var requestedRange in ranges)
        {
            var bounds = WorkbookRuntimeHelpers.GetBounds(requestedRange);
            var range = references.Add(Get(worksheet, "Range", requestedRange.ToString()));
            var formulaGrid = WorkbookRuntimeHelpers.CreateFormulaGrid(Get(range, "FormulaR1C1"), bounds.RowCount, bounds.ColumnCount);
            var inferred = FormulaPatternAnalyzer.InferRepairs(formulaGrid);
            rangeResults.Add(new RepairRangeResult(requestedRange, inferred.Count));
            foreach (var repair in inferred)
            {
                repairs.Add(new ExpectedFormula(bounds.StartRow + repair.RowIndex, bounds.StartColumn + repair.ColumnIndex, repair.FormulaR1C1));
            }
        }

        return FormulaExecutionPlan.Create(kind, targetWorksheetName, repairs, rangeResults);
    }

    private static FormulaExecutionPlan AnalyzeFormulaExtension(object workbook, NormalizedExtendFormulaSeriesOperation task)
    {
        using var references = new ComReferenceScope();
        var sheets = references.Add(Get(workbook, "Worksheets"));
        var sheet = references.Add(Get(sheets, "Item", task.WorksheetName));
        var evidence = WorkbookRuntimeHelpers.GetBounds(task.EvidenceRange);
        var destination = WorkbookRuntimeHelpers.GetBounds(task.DestinationRange);
        var combinedAddress = $"{WorkbookRuntimeHelpers.ToA1Address(Math.Min(evidence.StartRow, destination.StartRow), Math.Min(evidence.StartColumn, destination.StartColumn))}:{WorkbookRuntimeHelpers.ToA1Address(Math.Max(evidence.EndRow, destination.EndRow), Math.Max(evidence.EndColumn, destination.EndColumn))}";
        var range = references.Add(Get(sheet, "Range", combinedAddress));
        var grid = WorkbookRuntimeHelpers.CreateFormulaGrid(Get(range, "FormulaR1C1"),
            destination.EndRow - evidence.StartRow + 1,
            destination.EndColumn - evidence.StartColumn + 1);
        var periods = task.Direction == FormulaExtensionDirection.Right
            ? destination.EndColumn - destination.StartColumn + 1
            : destination.EndRow - destination.StartRow + 1;
        var planned = FormulaMutationPlanner.Plan(grid, task.Direction, periods);
        var repairs = planned.Mutations.Select(m => new ExpectedFormula(
            evidence.StartRow + m.RowIndex, evidence.StartColumn + m.ColumnIndex, m.FormulaR1C1)).ToList();
        return FormulaExecutionPlan.Create(ExcelOperationKind.ExtendFormulaSeries, task.WorksheetName, repairs,
            [new RepairRangeResult(task.DestinationRange, repairs.Count)]);
    }

    private static void ApplyFormulaWrites(ExcelSession session, FormulaExecutionPlan plan, Action markMutationAttempted)
    {
        if (plan.Repairs.Count == 0) return;
        using var references = new ComReferenceScope();
        var sheets = references.Add(Get(session.TargetWorkbook, "Worksheets"));
        var sheet = references.Add(Get(sheets, "Item", plan.WorksheetName));
        foreach (var group in plan.Repairs.GroupBy(repair => repair.FormulaR1C1, StringComparer.Ordinal))
        {
            foreach (var batch in group.Chunk(64))
            {
                var address = string.Join(",", batch.Select(repair => WorkbookRuntimeHelpers.ToA1Address(repair.Row, repair.Column)));
                var target = references.Add(Get(sheet, "Range", address));
                markMutationAttempted();
                Set(target, "FormulaR1C1", group.Key);
            }
        }
    }

    private static bool FormulaPlansEqual(FormulaExecutionPlan left, FormulaExecutionPlan right) =>
        left.Kind == right.Kind &&
        string.Equals(left.WorksheetName, right.WorksheetName, StringComparison.Ordinal) &&
        left.Fingerprint == right.Fingerprint &&
        left.Repairs.SequenceEqual(right.Repairs) &&
        left.RangeResults.SequenceEqual(right.RangeResults);

    private static TaskChange[] CreateFormulaChanges(FormulaExecutionPlan plan, bool planning)
    {
        var kind = plan.Kind == ExcelOperationKind.ExtendFormulaSeries ? "formula-extension" : "formula-repair";
        var verb = planning ? "Planned" : "Applied";
        return plan.RangeResults.Select(result => new TaskChange(
            kind,
            $"{plan.WorksheetName}!{result.Range}",
            $"{verb} {result.RepairCount} formula changes.")).ToArray();
    }

    private static bool VerifySavedWorkbook(
        string path,
        string worksheetName,
        IReadOnlyList<ExpectedFormula> expectedRepairs,
        IExcelWorkbookRuntimeObserver observer,
        out TaskCheck check)
    {
        ExcelSession? verification = null;
        var contentVerified = false;
        check = new TaskCheck("reopen-verification", false, "Saved workbook verification did not complete.");
        try
        {
            verification = ExcelSession.OpenForVerification(path, observer);
            using var references = new ComReferenceScope();
            var sheets = references.Add(Get(verification.TargetWorkbook, "Worksheets"));
            var sheet = references.Add(Get(sheets, "Item", worksheetName));
            foreach (var expected in expectedRepairs)
            {
                var cell = references.Add(Get(sheet, "Cells", expected.Row, expected.Column));
                var actual = Get(cell, "FormulaR1C1") as string;
                if (!string.Equals(actual, expected.FormulaR1C1, StringComparison.Ordinal))
                {
                    check = new TaskCheck("reopen-verification", false, "A repaired formula was not present after reopening the saved workbook.");
                    return false;
                }
            }

            contentVerified = true;
            check = new TaskCheck("reopen-verification", true, $"Saved workbook reopened with the requested worksheet and {expectedRepairs.Count} requested formulas.");
        }
        finally
        {
            if (verification is not null && !verification.Close())
            {
                check = new TaskCheck("verification-process-exit", false, "The owned verification Excel process did not exit.");
                contentVerified = false;
            }
        }

        return contentVerified;
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

    private sealed record ExpectedFormula(int Row, int Column, string FormulaR1C1);

    private sealed record RepairRangeResult(FormulaRepairRange Range, int RepairCount);

    private sealed record FormulaExecutionPlan(
        ExcelOperationKind Kind,
        string WorksheetName,
        IReadOnlyList<ExpectedFormula> Repairs,
        IReadOnlyList<RepairRangeResult> RangeResults,
        string Fingerprint)
    {
        public static FormulaExecutionPlan Create(ExcelOperationKind kind, string worksheetName, IReadOnlyList<ExpectedFormula> repairs, IReadOnlyList<RepairRangeResult> rangeResults)
        {
            var orderedRepairs = repairs.OrderBy(repair => repair.Row).ThenBy(repair => repair.Column).ToArray();
            var orderedRanges = rangeResults.OrderBy(result => result.Range.StartCell, StringComparer.Ordinal).ThenBy(result => result.Range.EndCell, StringComparer.Ordinal).ToArray();
            var material = new StringBuilder()
                .Append(kind).Append('|').Append(worksheetName).Append('|');
            foreach (var result in orderedRanges) material.Append(result.Range).Append(':').Append(result.RepairCount).Append('|');
            foreach (var repair in orderedRepairs) material.Append(repair.Row).Append(',').Append(repair.Column).Append(':').Append(repair.FormulaR1C1).Append('|');
            var fingerprint = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(material.ToString()))).ToLowerInvariant();
            return new FormulaExecutionPlan(kind, worksheetName, Array.AsReadOnly(orderedRepairs), Array.AsReadOnly(orderedRanges), fingerprint);
        }
    }

    private sealed record WorksheetCopyPreflight(bool IsFeasible, IReadOnlyList<TaskCheck> Checks);

    private sealed class ExcelSession : IDisposable
    {
        private readonly ComReferenceScope _references = new();
        private readonly bool _ownsApplication;
        private readonly bool _closeTarget;
        private readonly bool _closeReference;
        private readonly OwnedExcelProcess? _ownedProcess;
        private bool _closed;

        private ExcelSession(object application, object targetWorkbook, object referenceWorkbook, bool ownsApplication, bool closeTarget, bool closeReference, OwnedExcelProcess? ownedProcess)
        {
            Application = _references.Add(application);
            TargetWorkbook = _references.Add(targetWorkbook);
            ReferenceWorkbook = _references.Add(referenceWorkbook);
            _ownsApplication = ownsApplication;
            _closeTarget = closeTarget;
            _closeReference = closeReference;
            _ownedProcess = ownedProcess;
        }

        public object Application { get; }

        public object TargetWorkbook { get; }

        public object ReferenceWorkbook { get; }

        public bool HasExternalTargetOpen(string targetPath) =>
            RotWorkbookLocator.HasExternalWorkbookAtPath(targetPath, GetApplicationHwnd(Application));

        public static ExcelSession Open(NormalizedExcelTaskRequest request, IExcelWorkbookRuntimeObserver observer, bool readOnlyTarget = false)
        {
            var needsReference = NeedsReferenceWorkbook(request);
            var referencePath = request.Operation.CopyExhibit?.ReferenceWorkbookPath;
            if (request.WorkbookBinding == WorkbookBinding.UseOpen)
            {
                using var found = RotWorkbookLocator.Find(WorkbookRuntimeHelpers.NormalizePath(request.TargetWorkbookPath));
                if (found is null) throw new InvalidOperationException("The requested open target workbook was not found in the running object table.");
                var target = found.Detach();
                var application = Get(target, "Application");
                object? reference = null;
                var closeReference = false;
                try
                {
                    if (!needsReference)
                    {
                        return new ExcelSession(application, target, target, ownsApplication: false, closeTarget: false, closeReference: false, ownedProcess: null);
                    }

                    using var openReference = RotWorkbookLocator.Find(WorkbookRuntimeHelpers.NormalizePath(referencePath!));
                    if (openReference is not null)
                    {
                        var openReferenceApplication = Get(openReference.Workbook, "Application");
                        try
                        {
                            var sameApplication = AreSameApplication(application, openReferenceApplication);
                            if (sameApplication)
                            {
                                reference = openReference.Detach();
                            }

                            if (!sameApplication) ComReferences.Release(openReferenceApplication);
                        }
                        catch
                        {
                            ComReferences.Release(openReferenceApplication);
                            throw;
                        }
                    }

                    if (reference is null && WorkbookRuntimeHelpers.PathsEqual(request.TargetWorkbookPath, referencePath!))
                    {
                        reference = target;
                    }

                    if (reference is null)
                    {
                        var workbooks = Get(application, "Workbooks");
                        try
                        {
                            var priorAutomationSecurity = Get(application, "AutomationSecurity");
                            try
                            {
                                Set(application, "AutomationSecurity", WorkbookRuntimeHelpers.AutomationSecurityForceDisable);
                                reference = OpenWorkbook(workbooks, referencePath!, readOnly: true);
                            }
                            finally
                            {
                                Set(application, "AutomationSecurity", priorAutomationSecurity);
                            }
                            closeReference = true;
                        }
                        finally
                        {
                            ComReferences.Release(workbooks);
                        }
                    }

                    return new ExcelSession(application, target, reference!, ownsApplication: false, closeTarget: false, closeReference, ownedProcess: null);
                }
                catch
                {
                    ComReferences.Release(reference);
                    ComReferences.Release(target);
                    ComReferences.Release(application);
                    throw;
                }
            }

            var beforeStart = OwnedExcelProcess.SnapshotExcelProcesses();
            var app = CreateApplication();
            try
            {
                var ownedProcess = OwnedExcelProcess.CaptureNew(app, beforeStart);
                observer.OnOwnedProcessCaptured(ownedProcess.Identity);
                ConfigureOwnedApplication(app);
                var workbooks = Get(app, "Workbooks");
                try
                {
                    var target = OpenWorkbook(workbooks, request.TargetWorkbookPath, readOnly: readOnlyTarget);
                    object? reference = null;
                    try
                    {
                        reference = !needsReference || WorkbookRuntimeHelpers.PathsEqual(request.TargetWorkbookPath, referencePath!)
                            ? target
                            : OpenWorkbook(workbooks, referencePath!, readOnly: true);
                        return new ExcelSession(app, target, reference, ownsApplication: true, closeTarget: true, closeReference: !ReferenceEquals(target, reference), ownedProcess);
                    }
                    catch
                    {
                        if (!ReferenceEquals(reference, target)) ComReferences.Release(reference);
                        ComReferences.Release(target);
                        throw;
                    }
                }
                finally
                {
                    ComReferences.Release(workbooks);
                }
            }
            catch
            {
                if (OwnedExcelProcess.IsNewlyOwned(app, beforeStart)) TryQuit(app);
                ComReferences.Release(app);
                throw;
            }
        }

        public static ExcelSession OpenForVerification(string path, IExcelWorkbookRuntimeObserver observer)
        {
            var beforeStart = OwnedExcelProcess.SnapshotExcelProcesses();
            var app = CreateApplication();
            try
            {
                var ownedProcess = OwnedExcelProcess.CaptureNew(app, beforeStart);
                observer.OnOwnedProcessCaptured(ownedProcess.Identity);
                ConfigureOwnedApplication(app);
                var workbooks = Get(app, "Workbooks");
                try
                {
                    var workbook = OpenWorkbook(workbooks, path, readOnly: true);
                    return new ExcelSession(app, workbook, workbook, ownsApplication: true, closeTarget: true, closeReference: false, ownedProcess);
                }
                finally
                {
                    ComReferences.Release(workbooks);
                }
            }
            catch
            {
                if (OwnedExcelProcess.IsNewlyOwned(app, beforeStart)) TryQuit(app);
                ComReferences.Release(app);
                throw;
            }
        }

        public bool Close()
        {
            if (_closed) return _ownedProcess is null || !_ownedProcess.IsRunning;
            _closed = true;
            var ownedProcessExited = true;
            try
            {
                try
                {
                    if (_closeReference && !ReferenceEquals(ReferenceWorkbook, TargetWorkbook)) Invoke(ReferenceWorkbook, "Close", false);
                }
                finally
                {
                    try
                    {
                        if (_closeTarget) Invoke(TargetWorkbook, "Close", false);
                    }
                    finally
                    {
                        if (_ownsApplication) Invoke(Application, "Quit");
                    }
                }
            }
            finally
            {
                _references.Dispose();
                if (_ownedProcess is not null)
                {
                    ownedProcessExited = _ownedProcess.WaitForExitOrTerminate();
                }
            }

            return ownedProcessExited;
        }

        public void Dispose() => _ = Close();

        private static object CreateApplication()
        {
            var excelType = Type.GetTypeFromProgID("Excel.Application", throwOnError: true)
                ?? throw new InvalidOperationException("Microsoft Excel is not registered on this machine.");
            return Activator.CreateInstance(excelType) ?? throw new InvalidOperationException("Microsoft Excel could not be started.");
        }

        private static void ConfigureOwnedApplication(object application)
        {
            Set(application, "Visible", false);
            Set(application, "DisplayAlerts", false);
            Set(application, "AutomationSecurity", WorkbookRuntimeHelpers.AutomationSecurityForceDisable);
        }

        private static object OpenWorkbook(object workbooks, string path, bool readOnly) => Invoke(
            workbooks,
            "Open",
            WorkbookRuntimeHelpers.NormalizePath(path),
            0,
            readOnly,
            Type.Missing,
            Type.Missing,
            Type.Missing,
            true,
            Type.Missing,
            Type.Missing,
            Type.Missing,
            false,
            false,
            Type.Missing,
            false,
            0) ?? throw new InvalidOperationException("Excel did not open the workbook.");

        private static bool AreSameApplication(object left, object right) =>
            GetApplicationHwnd(left) == GetApplicationHwnd(right);

        private static long GetApplicationHwnd(object application) =>
            Convert.ToInt64(Get(application, "Hwnd"), CultureInfo.InvariantCulture);

        private static void TryQuit(object application)
        {
            try { Invoke(application, "Quit"); }
            catch (COMException) { }
        }
    }
}

internal static class WorkbookRuntimeHelpers
{
    public const int AutomationSecurityForceDisable = 3;

    private static readonly HashSet<string> SupportedWorkbookExtensions = new(StringComparer.OrdinalIgnoreCase) { ".xlsx", ".xlsm" };

    public static string NormalizePath(string path) => Path.GetFullPath(path ?? throw new ArgumentNullException(nameof(path)));

    public static bool PathsEqual(string left, string right) => string.Equals(
        NormalizePath(left).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar),
        NormalizePath(right).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar),
        StringComparison.OrdinalIgnoreCase);

    public static bool CanOpenExclusively(string path)
    {
        try
        {
            using var stream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            return true;
        }
        catch (IOException)
        {
            return false;
        }
        catch (UnauthorizedAccessException)
        {
            return false;
        }
    }

    public static void EnsureReadableWorkbook(string path, string description)
    {
        if (!SupportedWorkbookExtensions.Contains(Path.GetExtension(path)))
        {
            throw new InvalidOperationException($"{description} must be an .xlsx or .xlsm file.");
        }

        if (!File.Exists(path)) throw new InvalidOperationException($"{description} does not exist.");
    }

    public static void EnsureWritableCopyOutput(string? outputPath)
    {
        if (string.IsNullOrWhiteSpace(outputPath)) throw new InvalidOperationException("Copy output path is required.");
        var normalized = NormalizePath(outputPath);
        if (!SupportedWorkbookExtensions.Contains(Path.GetExtension(normalized)))
        {
            throw new InvalidOperationException("Copy output must be an .xlsx or .xlsm file.");
        }

        var parent = Directory.GetParent(normalized)?.FullName;
        if (parent is null || !Directory.Exists(parent)) throw new InvalidOperationException("Copy output directory does not exist.");
        if ((File.GetAttributes(parent) & FileAttributes.ReadOnly) != 0)
        {
            throw new InvalidOperationException("Copy output directory is read-only.");
        }

        if (File.Exists(normalized) && !CanOpenExclusively(normalized))
        {
            throw new InvalidOperationException("Existing copy output is locked.");
        }
    }

    public static string CreateStagingPath(string finalPath, string taskId)
    {
        var normalized = NormalizePath(finalPath);
        var directory = Path.GetDirectoryName(normalized) ?? throw new InvalidOperationException("Copy output directory is required.");
        var extension = Path.GetExtension(normalized);
        var name = Path.GetFileNameWithoutExtension(normalized);
        if (!IsSafeTaskId(taskId))
        {
            throw new InvalidOperationException("Task identity cannot be used for staging.");
        }

        return Path.Combine(directory, $".{name}.excel-task-{taskId}-{Guid.NewGuid():N}{extension}");
    }

    public static bool IsSafeTaskId(string? taskId) => taskId is { Length: > 0 and <= 128 } &&
        taskId.All(character => char.IsAsciiLetterOrDigit(character) || character is '-' or '_');

    public static void PromoteStaging(string stagingPath, string finalPath, bool overwrite)
    {
        File.Move(stagingPath, finalPath, overwrite);
    }

    public static bool TryDeleteStaging(string stagingPath)
    {
        try
        {
            if (File.Exists(stagingPath)) File.Delete(stagingPath);
            return !File.Exists(stagingPath);
        }
        catch (IOException) { return false; }
        catch (UnauthorizedAccessException) { return false; }
    }

    public static FormulaRangeBounds GetBounds(FormulaRepairRange range)
    {
        ArgumentNullException.ThrowIfNull(range);
        var start = ParseCell(range.StartCell);
        var end = ParseCell(range.EndCell);
        if (start.Row > end.Row || start.Column > end.Column) throw new InvalidOperationException("Formula repair range is not rectangular.");
        return new FormulaRangeBounds(start.Row, start.Column, end.Row, end.Column);
    }

    public static FormulaGridCell[,] CreateFormulaGrid(object value, int rowCount, int columnCount)
    {
        var grid = new FormulaGridCell[rowCount, columnCount];
        for (var row = 0; row < rowCount; row++)
        {
            for (var column = 0; column < columnCount; column++)
            {
                var cellValue = value is Array values ? values.GetValue(row + values.GetLowerBound(0), column + values.GetLowerBound(1)) : value;
                grid[row, column] = cellValue switch
                {
                    null => FormulaGridCell.Blank,
                    string { Length: 0 } => FormulaGridCell.Blank,
                    string formula when formula.StartsWith('=') => FormulaGridCell.Formula(formula),
                    _ => FormulaGridCell.Constant
                };
            }
        }

        return grid;
    }

    public static string ToA1Address(int row, int column)
    {
        if (row is < 1 or > 1_048_576 || column is < 1 or > 16_384)
        {
            throw new ArgumentOutOfRangeException(nameof(row));
        }

        Span<char> letters = stackalloc char[3];
        var index = letters.Length;
        var remaining = column;
        while (remaining > 0)
        {
            remaining--;
            letters[--index] = (char)('A' + (remaining % 26));
            remaining /= 26;
        }

        return new string(letters[index..]) + row.ToString(CultureInfo.InvariantCulture);
    }

    private static CellAddress ParseCell(string text)
    {
        var index = 0;
        var column = 0;
        while (index < text.Length && char.IsLetter(text[index]))
        {
            column = checked((column * 26) + char.ToUpperInvariant(text[index]) - 'A' + 1);
            index++;
        }

        if (index == 0 || !int.TryParse(text[index..], CultureInfo.InvariantCulture, out var row) ||
            row is < 1 or > 1_048_576 || column is < 1 or > 16_384)
        {
            throw new InvalidOperationException("Formula repair range contains an invalid cell address.");
        }

        return new CellAddress(row, column);
    }

    private sealed record CellAddress(int Row, int Column);
}

internal sealed record FormulaRangeBounds(int StartRow, int StartColumn, int EndRow, int EndColumn)
{
    public int RowCount => EndRow - StartRow + 1;

    public int ColumnCount => EndColumn - StartColumn + 1;
}

internal sealed class RotWorkbookLocator : IDisposable
{
    private object? _workbook;

    private RotWorkbookLocator(object workbook) => _workbook = workbook;

    public object Workbook => _workbook ?? throw new ObjectDisposedException(nameof(RotWorkbookLocator));

    public object Detach()
    {
        var workbook = Workbook;
        _workbook = null;
        return workbook;
    }

    public static RotWorkbookLocator? Find(string targetPath)
    {
        var result = GetRunningObjectTable(0, out var table);
        if (result < 0 || table is null) Marshal.ThrowExceptionForHR(result);
        var runningTable = table ?? throw new InvalidOperationException("The running object table was unavailable.");
        try
        {
            runningTable.EnumRunning(out var enumerator);
            try
            {
                var monikers = new IMoniker[1];
                while (enumerator.Next(1, monikers, IntPtr.Zero) == 0)
                {
                    var moniker = monikers[0];
                    try
                    {
                        var bindResult = CreateBindCtx(0, out var bindContext);
                        if (bindResult < 0 || bindContext is null) continue;
                        try
                        {
                            moniker.GetDisplayName(bindContext, null, out var displayName);
                            if (!MatchesDisplayName(displayName, targetPath)) continue;
                            moniker.BindToObject(bindContext, null, ref WorkbookInterfaceId, out var candidate);
                            if (candidate is not null && HasMatchingFullName(candidate, targetPath)) return new RotWorkbookLocator(candidate);
                            ComReferences.Release(candidate);
                        }
                        catch (Exception exception) when (IsExpectedBindingNonmatch(exception)) { }
                        finally { ComReferences.Release(bindContext); }
                    }
                    finally { ComReferences.Release(moniker); }
                }
            }
            finally { ComReferences.Release(enumerator); }
        }
        finally { ComReferences.Release(runningTable); }

        return null;
    }

    public static bool ContainsPath(string targetPath)
    {
        var result = GetRunningObjectTable(0, out var table);
        if (result < 0 || table is null) Marshal.ThrowExceptionForHR(result);
        var runningTable = table ?? throw new InvalidOperationException("The running object table was unavailable.");
        try
        {
            runningTable.EnumRunning(out var enumerator);
            try
            {
                var monikers = new IMoniker[1];
                while (enumerator.Next(1, monikers, IntPtr.Zero) == 0)
                {
                    var moniker = monikers[0];
                    try
                    {
                        var bindResult = CreateBindCtx(0, out var bindContext);
                        if (bindResult < 0 || bindContext is null) continue;
                        try
                        {
                            moniker.GetDisplayName(bindContext, null, out var displayName);
                            if (MatchesDisplayName(displayName, targetPath)) return true;
                        }
                        catch (Exception exception) when (IsExpectedBindingNonmatch(exception)) { }
                        finally { ComReferences.Release(bindContext); }
                    }
                    finally { ComReferences.Release(moniker); }
                }
            }
            finally { ComReferences.Release(enumerator); }
        }
        finally { ComReferences.Release(runningTable); }

        return false;
    }

    /// <summary>
    /// Detects an exact target workbook in the ROT that belongs to a different Excel application.
    /// The caller's application RCWs are intentionally never released here: ROT can return the
    /// same RCW instance used by the active session, and final-releasing it would invalidate that session.
    /// </summary>
    public static bool HasExternalWorkbookAtPath(string targetPath, long sessionApplicationHwnd)
    {
        var result = GetRunningObjectTable(0, out var table);
        if (result < 0 || table is null) Marshal.ThrowExceptionForHR(result);
        var runningTable = table ?? throw new InvalidOperationException("The running object table was unavailable.");
        try
        {
            runningTable.EnumRunning(out var enumerator);
            try
            {
                var monikers = new IMoniker[1];
                while (enumerator.Next(1, monikers, IntPtr.Zero) == 0)
                {
                    var moniker = monikers[0];
                    try
                    {
                        var bindResult = CreateBindCtx(0, out var bindContext);
                        if (bindResult < 0 || bindContext is null) continue;
                        try
                        {
                            moniker.GetDisplayName(bindContext, null, out var displayName);
                            if (!MatchesDisplayName(displayName, targetPath)) continue;

                            object? candidate = null;
                            object? candidateApplication = null;
                            var unrelatedCandidate = false;
                            try
                            {
                                moniker.BindToObject(bindContext, null, ref WorkbookInterfaceId, out candidate);
                                if (candidate is null || !HasMatchingFullName(candidate, targetPath)) continue;

                                candidateApplication = GetApplication(candidate);
                                var candidateApplicationHwnd = GetApplicationHwnd(candidateApplication);
                                unrelatedCandidate = IsExternalApplicationHwnd(sessionApplicationHwnd, candidateApplicationHwnd);
                                if (unrelatedCandidate) return true;
                            }
                            catch (Exception exception) when (IsExpectedBindingNonmatch(exception)) { }
                            finally
                            {
                                if (unrelatedCandidate)
                                {
                                    ComReferences.Release(candidateApplication);
                                    ComReferences.Release(candidate);
                                }
                            }
                        }
                        finally { ComReferences.Release(bindContext); }
                    }
                    finally { ComReferences.Release(moniker); }
                }
            }
            finally { ComReferences.Release(enumerator); }
        }
        finally { ComReferences.Release(runningTable); }

        return false;
    }

    public void Dispose() => ComReferences.Release(_workbook);

    internal static bool IsExpectedBindingNonmatch(Exception exception) => exception is COMException or ArgumentException;

    internal static bool RequiresPreMutationIsolatedSameApplyRevalidation(
        ExcelTaskMode mode,
        WorkbookBinding binding,
        SaveMode save) =>
        mode == ExcelTaskMode.Apply &&
        binding == WorkbookBinding.Isolated &&
        save == SaveMode.Same;

    internal static bool IsExternalApplicationHwnd(long sessionApplicationHwnd, long candidateApplicationHwnd) =>
        sessionApplicationHwnd != candidateApplicationHwnd;

    internal static bool MatchesDisplayName(string? displayName, string targetPath)
    {
        if (string.IsNullOrWhiteSpace(displayName)) return false;
        var candidate = displayName[0] == '!' ? displayName[1..] : displayName;
        return WorkbookRuntimeHelpers.PathsEqual(candidate, targetPath);
    }

    private static bool HasMatchingFullName(object candidate, string targetPath)
    {
        try
        {
            var fullName = candidate.GetType().InvokeMember("FullName", BindingFlags.GetProperty, null, candidate, null, CultureInfo.InvariantCulture) as string;
            return fullName is not null && WorkbookRuntimeHelpers.PathsEqual(fullName, targetPath);
        }
        catch (COMException)
        {
            return false;
        }
    }

    private static object GetApplication(object workbook) => workbook.GetType().InvokeMember(
        "Application",
        BindingFlags.GetProperty,
        null,
        workbook,
        null,
        CultureInfo.InvariantCulture) ?? throw new InvalidOperationException("Excel workbook did not return its application.");

    private static long GetApplicationHwnd(object application) => Convert.ToInt64(application.GetType().InvokeMember(
        "Hwnd",
        BindingFlags.GetProperty,
        null,
        application,
        null,
        CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);

    private static Guid WorkbookInterfaceId = new("00000000-0000-0000-C000-000000000046");

    [DllImport("ole32.dll")]
    private static extern int GetRunningObjectTable(uint reserved, out IRunningObjectTable? runningObjectTable);

    [DllImport("ole32.dll")]
    private static extern int CreateBindCtx(uint reserved, out IBindCtx? bindContext);
}

internal sealed class OwnedExcelProcess
{
    private readonly ProcessIdentity _identity;

    private OwnedExcelProcess(ProcessIdentity identity) => _identity = identity;

    internal ProcessIdentity Identity => _identity;

    public static HashSet<ProcessIdentity> SnapshotExcelProcesses()
    {
        var identities = new HashSet<ProcessIdentity>();
        foreach (var process in Process.GetProcessesByName("EXCEL"))
        {
            using (process)
            {
                try { identities.Add(ProcessIdentity.Capture(process.Id)); }
                catch (InvalidOperationException) { }
                catch (System.ComponentModel.Win32Exception) { }
            }
        }

        return identities;
    }

    public static OwnedExcelProcess CaptureNew(object application, ISet<ProcessIdentity> preExisting)
    {
        var identity = GetApplicationIdentity(application);
        if (preExisting.Contains(identity))
        {
            throw new InvalidOperationException("Excel activation did not create a new owned process.");
        }

        return new OwnedExcelProcess(identity);
    }

    public static bool IsNewlyOwned(object application, ISet<ProcessIdentity> preExisting)
    {
        try { return !preExisting.Contains(GetApplicationIdentity(application)); }
        catch (COMException) { return false; }
        catch (InvalidOperationException) { return false; }
    }

    private static ProcessIdentity GetApplicationIdentity(object application)
    {
        var hwndValue = application.GetType().InvokeMember("Hwnd", BindingFlags.GetProperty, null, application, null, CultureInfo.InvariantCulture);
        var hwnd = new IntPtr(Convert.ToInt64(hwndValue, CultureInfo.InvariantCulture));
        _ = GetWindowThreadProcessId(hwnd, out var processId);
        if (processId == 0) throw new InvalidOperationException("Excel process identity could not be captured.");
        return ProcessIdentity.Capture((int)processId);
    }

    public bool WaitForExitOrTerminate()
    {
        if (!ProcessIdentity.TryOpenMatching(_identity, out var process)) return true;
        using (process)
        {
            if (process.WaitForExit(10_000)) return HasExited();
            if (!ProcessIdentity.TryOpenMatching(_identity, out var stillMatching)) return true;
            using (stillMatching)
            {
                stillMatching.Kill(entireProcessTree: false);
                _ = stillMatching.WaitForExit(5_000);
            }

            return HasExited();
        }
    }

    private bool HasExited()
    {
        if (!ProcessIdentity.TryOpenMatching(_identity, out var process)) return true;
        process.Dispose();
        return false;
    }

    public bool IsRunning => ProcessIdentity.TryOpenMatching(_identity, out var process) && DisposeProcess(process);

    private static bool DisposeProcess(Process process)
    {
        process.Dispose();
        return true;
    }

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
}

internal sealed record ProcessIdentity(int ProcessId, DateTime StartTimeUtc, string ExecutablePath)
{
    public static ProcessIdentity Capture(int processId)
    {
        using var process = Process.GetProcessById(processId);
        return new ProcessIdentity(processId, process.StartTime.ToUniversalTime(), GetExecutablePath(process));
    }

    public static bool TryOpenMatching(ProcessIdentity identity, out Process process)
    {
        process = null!;
        try
        {
            var candidate = Process.GetProcessById(identity.ProcessId);
            if (candidate.StartTime.ToUniversalTime() != identity.StartTimeUtc ||
                !string.Equals(GetExecutablePath(candidate), identity.ExecutablePath, StringComparison.OrdinalIgnoreCase))
            {
                candidate.Dispose();
                return false;
            }

            process = candidate;
            return true;
        }
        catch (ArgumentException)
        {
            return false;
        }
        catch (InvalidOperationException)
        {
            return false;
        }
    }

    private static string GetExecutablePath(Process process) => process.MainModule?.FileName
        ?? throw new InvalidOperationException("Excel executable path could not be captured.");
}

internal sealed class ComReferenceScope : IDisposable
{
    private readonly HashSet<object> _references = new(ReferenceEqualityComparer.Instance);

    public T Add<T>(T value) where T : class
    {
        _references.Add(value);
        return value;
    }

    public void Dispose()
    {
        foreach (var reference in _references) ComReferences.Release(reference);
        _references.Clear();
    }
}

internal static class ComReferences
{
    public static void Release(object? value)
    {
        if (value is not null && Marshal.IsComObject(value)) Marshal.FinalReleaseComObject(value);
    }
}

internal sealed class StaComDispatcher : IDisposable
{
    private readonly BlockingCollection<IWorkItem> _queue = new(boundedCapacity: 32);
    private readonly Thread _thread;
    private bool _disposed;

    public StaComDispatcher()
    {
        _thread = new Thread(Run) { IsBackground = true, Name = "ExcelTask COM STA" };
        _thread.SetApartmentState(ApartmentState.STA);
        _thread.Start();
    }

    public Task<T> InvokeAsync<T>(Func<T> callback, CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(callback);
        cancellationToken.ThrowIfCancellationRequested();
        var item = new WorkItem<T>(callback, cancellationToken);
        if (!_queue.TryAdd(item))
        {
            ObjectDisposedException.ThrowIf(_queue.IsAddingCompleted, this);
            throw new InvalidOperationException("Excel runtime queue is full; retry after queued tasks complete.");
        }
        return item.Task;
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        _queue.CompleteAdding();
        _thread.Join();
        _queue.Dispose();
    }

    private void Run()
    {
        _ = PeekMessage(out _, IntPtr.Zero, 0, 0, 0);
        foreach (var item in _queue.GetConsumingEnumerable())
        {
            item.Run();
            PumpMessages();
        }
    }

    private static void PumpMessages()
    {
        while (PeekMessage(out var message, IntPtr.Zero, 0, 0, 1))
        {
            _ = TranslateMessage(ref message);
            _ = DispatchMessage(ref message);
        }
    }

    [DllImport("user32.dll")]
    private static extern bool PeekMessage(out Message message, IntPtr window, uint min, uint max, uint remove);

    [DllImport("user32.dll")]
    private static extern bool TranslateMessage(ref Message message);

    [DllImport("user32.dll")]
    private static extern IntPtr DispatchMessage(ref Message message);

    [StructLayout(LayoutKind.Sequential)]
    private struct Message
    {
        public IntPtr Hwnd;
        public uint MessageId;
        public IntPtr WParam;
        public IntPtr LParam;
        public uint Time;
        public int PointX;
        public int PointY;
        public uint Private;
    }

    private interface IWorkItem { void Run(); }

    private sealed class WorkItem<T>(Func<T> callback, CancellationToken cancellationToken) : IWorkItem
    {
        private readonly TaskCompletionSource<T> _completion = new(TaskCreationOptions.RunContinuationsAsynchronously);

        public Task<T> Task => _completion.Task;

        public void Run()
        {
            if (cancellationToken.IsCancellationRequested)
            {
                _completion.TrySetCanceled(cancellationToken);
                return;
            }

            try { _completion.SetResult(callback()); }
            catch (Exception exception) { _completion.SetException(exception); }
        }
    }
}
