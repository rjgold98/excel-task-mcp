using System.Collections.Concurrent;
using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.Versioning;
using ExcelTask.Core;

namespace ExcelTask.Excel;

/// <summary>Runs all desktop Excel automation on one message-pumping STA thread.</summary>
[SupportedOSPlatform("windows")]
public sealed class ExcelWorkbookRuntime : IWorkbookRuntime, IDisposable
{
    private readonly StaComDispatcher _dispatcher = new();
    private bool _disposed;

    public Task<WorkbookInspection> InspectAsync(WorkbookInspectionRequest request, CancellationToken cancellationToken)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        ArgumentNullException.ThrowIfNull(request);
        return _dispatcher.InvokeAsync(() => InspectCore(request), cancellationToken);
    }

    public Task<WorkbookExecutionOutcome> ExecuteAsync(ExcelTaskPlan plan, CancellationToken cancellationToken)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        ArgumentNullException.ThrowIfNull(plan);
        return _dispatcher.InvokeAsync(() => ExecuteCore(plan), cancellationToken);
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        _dispatcher.Dispose();
    }

    private static WorkbookInspection InspectCore(WorkbookInspectionRequest request)
    {
        var targetPath = WorkbookRuntimeHelpers.NormalizePath(request.TargetWorkbookPath);
        var referencePath = WorkbookRuntimeHelpers.NormalizePath(request.ReferenceWorkbookPath);
        WorkbookRuntimeHelpers.EnsureReadableWorkbook(targetPath, "Target workbook");
        WorkbookRuntimeHelpers.EnsureReadableWorkbook(referencePath, "Reference workbook");
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

    private static WorkbookExecutionOutcome ExecuteCore(ExcelTaskPlan plan)
    {
        try
        {
            WorkbookRuntimeHelpers.EnsureReadableWorkbook(WorkbookRuntimeHelpers.NormalizePath(plan.Request.TargetWorkbookPath), "Target workbook");
            WorkbookRuntimeHelpers.EnsureReadableWorkbook(WorkbookRuntimeHelpers.NormalizePath(plan.Request.ReferenceWorkbookPath), "Reference workbook");
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

        ExcelSession? session = null;
        var mutationAttempted = false;
        var verified = false;
        var changes = new List<TaskChange>();
        var checks = new List<TaskCheck>();
        var repairs = new RepairApplication([], []);
        var phase = "input-validation";
        var savedPath = WorkbookRuntimeHelpers.NormalizePath(plan.Request.Save == SaveMode.Copy
            ? plan.Request.OutputWorkbookPath ?? throw new InvalidOperationException("Copy output path is required.")
            : plan.Request.TargetWorkbookPath);
        string? stagingPath = null;

        try
        {
            if (plan.Request.Save == SaveMode.Copy && File.Exists(savedPath) && !plan.Request.OverwriteConfirmed)
            {
                return new WorkbookExecutionOutcome(
                    ExcelTaskStatus.Rejected,
                    "The copy output already exists and was not authorized for overwrite.",
                    Checks: [new TaskCheck("copy-output", false, "Existing output requires overwrite confirmation.")]);
            }

            if (plan.Request.WorkbookBinding == WorkbookBinding.Isolated && plan.Request.Save == SaveMode.Same)
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

            phase = "session-open";
            session = ExcelSession.Open(plan.Request, readOnlyTarget: plan.Request.Mode == ExcelTaskMode.Plan);
            phase = "preflight";
            var preflight = PreflightWorksheetCopy(session, plan.Request.ReferenceWorksheet, plan.Request.NewWorksheetName);
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

                return new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected, "Workbook preflight did not permit the requested worksheet copy.", Checks: checks);
            }

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
                    [new TaskChange("plan", "workbook", "Reference worksheet copy and bounded formula repair are feasible.")],
                    checks);
            }

            mutationAttempted = true;
            phase = "worksheet-copy";
            CopyReferenceWorksheet(session, plan.Request.ReferenceWorksheet, plan.Request.NewWorksheetName, value => phase = value);
            changes.Add(new TaskChange("worksheet-copy", "workbook", "Copied the requested reference worksheet."));

            phase = "formula-repair";
            repairs = ApplyFormulaRepairs(session, plan.Request.NewWorksheetName, plan.Request.FormulaRepairRanges);
            changes.AddRange(repairs.RangeResults.Select(result => new TaskChange(
                "formula-repair",
                result.Range.ToString(),
                $"Applied {result.RepairCount} safely inferred blank repairs in the requested range.")));
            checks.Add(new TaskCheck("formula-repair-count", true, $"Applied {repairs.Repairs.Count} repairs across {repairs.RangeResults.Count} requested ranges."));

            phase = "recalculate";
            Invoke(session.Application, "CalculateFull");
            phase = "save";
            if (plan.Request.Save == SaveMode.Copy)
            {
                stagingPath = WorkbookRuntimeHelpers.CreateStagingPath(savedPath);
                Invoke(session.TargetWorkbook, "SaveAs", stagingPath);
            }
            else
            {
                Invoke(session.TargetWorkbook, "Save");
            }

            checks.Add(new TaskCheck("save", true, "Excel completed the requested save operation."));
            phase = "primary-cleanup";
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
            phase = "reopen-verification";
            if (!VerifySavedWorkbook(verificationPath, plan.Request.NewWorksheetName, repairs.Repairs, out var verificationCheck))
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
            verified = true;
            if (stagingPath is not null)
            {
                phase = "copy-promotion";
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
                catch (Exception cleanupException) when (cleanupException is COMException or InvalidOperationException or TargetInvocationException) { ownedCleanupFailed = true; }
                session = null;
            }
            var stagingCleanupFailed = stagingPath is not null && !WorkbookRuntimeHelpers.TryDeleteStaging(stagingPath);
            if (!stagingCleanupFailed) stagingPath = null;
            var status = ownedCleanupFailed || stagingCleanupFailed ? ExcelTaskStatus.Unknown : verified ? ExcelTaskStatus.Partial : mutationAttempted ? ExcelTaskStatus.Unknown : ExcelTaskStatus.Rejected;
            checks.Add(new TaskCheck("execution", false, mutationAttempted
                ? $"Excel reported {exception.GetType().Name} during {phase}; the change was not fully verified."
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

    private static RepairApplication ApplyFormulaRepairs(
        ExcelSession session,
        string worksheetName,
        IReadOnlyList<FormulaRepairRange> ranges)
    {
        using var references = new ComReferenceScope();
        var worksheetCollection = references.Add(Get(session.TargetWorkbook, "Worksheets"));
        var worksheet = references.Add(Get(worksheetCollection, "Item", worksheetName));
        var repairs = new List<ExpectedFormula>();
        var rangeResults = new List<RepairRangeResult>();

        foreach (var requestedRange in ranges)
        {
            var bounds = WorkbookRuntimeHelpers.GetBounds(requestedRange);
            var range = references.Add(Get(worksheet, "Range", requestedRange.ToString()));
            var formulaGrid = WorkbookRuntimeHelpers.CreateFormulaGrid(Get(range, "FormulaR1C1"), bounds.RowCount, bounds.ColumnCount);
            var inferred = FormulaPatternAnalyzer.InferRepairs(formulaGrid);
            rangeResults.Add(new RepairRangeResult(requestedRange, inferred.Count));
            if (inferred.Count == 0) continue;

            foreach (var repair in inferred)
            {
                repairs.Add(new ExpectedFormula(bounds.StartRow + repair.RowIndex, bounds.StartColumn + repair.ColumnIndex, repair.FormulaR1C1));
            }

            foreach (var formulaGroup in inferred.GroupBy(repair => repair.FormulaR1C1, StringComparer.Ordinal))
            {
                foreach (var repairBatch in formulaGroup.Chunk(64))
                {
                    var targetAddress = string.Join(",", repairBatch.Select(repair => WorkbookRuntimeHelpers.ToA1Address(
                        bounds.StartRow + repair.RowIndex,
                        bounds.StartColumn + repair.ColumnIndex)));
                    var repairRange = references.Add(Get(worksheet, "Range", targetAddress));
                    Set(repairRange, "FormulaR1C1", formulaGroup.Key);
                }
            }
        }

        return new RepairApplication(repairs, rangeResults);
    }

    private static bool VerifySavedWorkbook(
        string path,
        string worksheetName,
        IReadOnlyList<ExpectedFormula> expectedRepairs,
        out TaskCheck check)
    {
        ExcelSession? verification = null;
        var contentVerified = false;
        check = new TaskCheck("reopen-verification", false, "Saved workbook verification did not complete.");
        try
        {
            verification = ExcelSession.OpenForVerification(path);
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
            check = new TaskCheck("reopen-verification", true, $"Saved workbook reopened with the copied worksheet and {expectedRepairs.Count} expected repairs.");
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

    private sealed record RepairApplication(List<ExpectedFormula> Repairs, List<RepairRangeResult> RangeResults);

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

        public static ExcelSession Open(NormalizedExcelTaskRequest request, bool readOnlyTarget = false)
        {
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
                    using var openReference = RotWorkbookLocator.Find(WorkbookRuntimeHelpers.NormalizePath(request.ReferenceWorkbookPath));
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

                    if (reference is null && WorkbookRuntimeHelpers.PathsEqual(request.TargetWorkbookPath, request.ReferenceWorkbookPath))
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
                                reference = OpenWorkbook(workbooks, request.ReferenceWorkbookPath, readOnly: true);
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
                var identity = OwnedExcelProcess.CaptureNew(app, beforeStart);
                ConfigureOwnedApplication(app);
                var workbooks = Get(app, "Workbooks");
                try
                {
                    var target = OpenWorkbook(workbooks, request.TargetWorkbookPath, readOnly: readOnlyTarget);
                    object? reference = null;
                    try
                    {
                        reference = WorkbookRuntimeHelpers.PathsEqual(request.TargetWorkbookPath, request.ReferenceWorkbookPath)
                            ? target
                            : OpenWorkbook(workbooks, request.ReferenceWorkbookPath, readOnly: true);
                        return new ExcelSession(app, target, reference, ownsApplication: true, closeTarget: true, closeReference: !ReferenceEquals(target, reference), identity);
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

        public static ExcelSession OpenForVerification(string path)
        {
            var beforeStart = OwnedExcelProcess.SnapshotExcelProcesses();
            var app = CreateApplication();
            try
            {
                var identity = OwnedExcelProcess.CaptureNew(app, beforeStart);
                ConfigureOwnedApplication(app);
                var workbooks = Get(app, "Workbooks");
                try
                {
                    var workbook = OpenWorkbook(workbooks, path, readOnly: true);
                    return new ExcelSession(app, workbook, workbook, ownsApplication: true, closeTarget: true, closeReference: false, identity);
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
            Convert.ToInt64(Get(left, "Hwnd"), CultureInfo.InvariantCulture) ==
            Convert.ToInt64(Get(right, "Hwnd"), CultureInfo.InvariantCulture);

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

    public static string CreateStagingPath(string finalPath)
    {
        var normalized = NormalizePath(finalPath);
        var directory = Path.GetDirectoryName(normalized) ?? throw new InvalidOperationException("Copy output directory is required.");
        var extension = Path.GetExtension(normalized);
        var name = Path.GetFileNameWithoutExtension(normalized);
        return Path.Combine(directory, $".{name}.excel-task-{Guid.NewGuid():N}{extension}");
    }

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

        if (index == 0 || !int.TryParse(text[index..], CultureInfo.InvariantCulture, out var row) || row < 1 || column < 1)
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

    public void Dispose() => ComReferences.Release(_workbook);

    internal static bool IsExpectedBindingNonmatch(Exception exception) => exception is COMException or ArgumentException;

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
