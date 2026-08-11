using System.Reflection;
using System.Runtime.InteropServices;
using ExcelTask.Core;

namespace ExcelTask.Excel;

public sealed partial class ExcelWorkbookRuntime
{
    /// <summary>Excel's xlWorkbookDefault format, so a new .xlsx is saved as one.</summary>
    private const int XlWorkbookDefault = 51;

    /// <summary>
    /// Creates an empty workbook, or adds an empty worksheet to an existing one.
    ///
    /// Every other operation in this server starts from a workbook that exists, which was a hard
    /// wall: a caller with nowhere to put a result had to leave and come back. Creation is the
    /// smallest thing that removes it, and it stays small by refusing to do anything else - no
    /// starting content, no template, no position, and never an overwrite. What comes next is a
    /// write or a copy, which already exist and already verify.
    /// </summary>
    private static WorkbookExecutionOutcome ExecuteCreateCore(ExcelTaskPlan plan, IExcelWorkbookRuntimeObserver observer)
    {
        var operation = plan.Request.Operation.Create!;
        var targetPath = WorkbookRuntimeHelpers.NormalizePath(plan.Request.TargetWorkbookPath);

        try
        {
            if (operation.Kind == CreateKind.Workbook)
            {
                WorkbookRuntimeHelpers.EnsureCreatableWorkbook(targetPath);
            }
            else
            {
                WorkbookRuntimeHelpers.EnsureReadableWorkbook(targetPath, "Target workbook");
                if (plan.Request.Mode == ExcelTaskMode.Apply) WorkbookRuntimeHelpers.EnsureWritableSameTarget(targetPath);
            }
        }
        catch (InvalidOperationException exception)
        {
            return new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected, "The creation target cannot be used.",
                Checks: [new TaskCheck("workbook-inputs", false, exception.Message)]);
        }

        // Plan reaches its answer from the filesystem alone, without starting Excel. Everything it
        // could report - whether the path is free, whether the sheet name is taken - is either
        // already settled above or needs the workbook open, and for a creation the first is the part
        // a caller wants to check cheaply.
        if (plan.Request.Mode == ExcelTaskMode.Plan && operation.Kind == CreateKind.Workbook)
        {
            return new WorkbookExecutionOutcome(ExcelTaskStatus.Planned,
                $"The path is free; applying would create an empty workbook at it. Nothing was created.",
                [new TaskChange("workbook-create", Path.GetFileName(targetPath), "Planned an empty workbook.")],
                [new TaskCheck("create-path", true, "No file exists at the requested path and its directory is writable.")]);
        }

        return operation.Kind == CreateKind.Workbook
            ? CreateWorkbook(targetPath, observer)
            : CreateWorksheet(plan, operation, targetPath, observer);
    }

    private static WorkbookExecutionOutcome CreateWorkbook(string targetPath, IExcelWorkbookRuntimeObserver observer)
    {
        var checks = new List<TaskCheck>();
        var changes = new List<TaskChange>();
        ExcelSession? session = null;
        var mutationAttempted = false;

        try
        {
            observer.OnPhase("session-open");
            session = ExcelSession.OpenNewWorkbook(observer);

            observer.OnPhase("save");
            mutationAttempted = true;
            Invoke(session.TargetWorkbook, "SaveAs", targetPath,
                string.Equals(Path.GetExtension(targetPath), ".xlsm", StringComparison.OrdinalIgnoreCase)
                    ? XlOpenXmlWorkbookMacroEnabled
                    : XlWorkbookDefault);
            checks.Add(new TaskCheck("save", true, "Excel saved the new workbook to the requested path."));
            changes.Add(new TaskChange("workbook-create", Path.GetFileName(targetPath), "Created an empty workbook."));

            observer.OnPhase("primary-cleanup");
            var cleanupFailure = ExcelSession.CloseAndProve(ref session, "the workbook creation", checks, changes);
            if (cleanupFailure is not null) return cleanupFailure with { RetryReason = "Inspect the created workbook before retrying." };

            // Proof from the filesystem rather than from Excel's word for it. A creation is the one
            // operation where "it exists now" is the entire result, so reopening it to read a cell
            // would be verifying something nobody asked to be there.
            observer.OnPhase("create-verification");
            if (!File.Exists(targetPath))
            {
                checks.Add(new TaskCheck("create-verification", false, "Excel reported a successful save but no file exists at the path."));
                return new WorkbookExecutionOutcome(ExcelTaskStatus.Unknown, "The new workbook could not be found after saving.",
                    changes, checks, CanRetry: false, RetryReason: "Inspect the requested path before retrying.");
            }

            checks.Add(new TaskCheck("create-verification", true, "A workbook file exists at the requested path after owned Excel exited."));
            return new WorkbookExecutionOutcome(ExcelTaskStatus.Completed,
                $"Created an empty workbook at the requested path and confirmed it exists.", changes, checks);
        }
        catch (Exception exception) when (exception is COMException or TargetInvocationException or InvalidOperationException or InvalidComObjectException)
        {
            checks.Add(new TaskCheck("workbook-create", false, DescribeFailure("workbook-create", exception)));
            var cleanupFailure = ExcelSession.CloseAndProve(ref session, "the failed workbook creation", checks, changes);
            if (cleanupFailure is not null) return cleanupFailure;
            return new WorkbookExecutionOutcome(
                mutationAttempted ? ExcelTaskStatus.Unknown : ExcelTaskStatus.Rejected,
                mutationAttempted
                    ? "The workbook creation failed after the save was attempted."
                    : "The workbook creation was rejected before anything was written.",
                changes, checks, CanRetry: !mutationAttempted,
                RetryReason: mutationAttempted ? "Inspect the requested path before retrying." : null);
        }
        finally
        {
            session?.Close();
        }
    }

    private static WorkbookExecutionOutcome CreateWorksheet(
        ExcelTaskPlan plan,
        NormalizedCreateOperation operation,
        string targetPath,
        IExcelWorkbookRuntimeObserver observer)
    {
        var apply = plan.Request.Mode == ExcelTaskMode.Apply;
        var checks = new List<TaskCheck>();
        var changes = new List<TaskChange>();
        ExcelSession? session = null;
        PendingVerification? pendingVerification = null;
        var mutationAttempted = false;

        try
        {
            if (apply) pendingVerification = PendingVerification.Begin(observer);

            observer.OnPhase("session-open");
            session = ExcelSession.Open(plan.Request, observer, readOnlyTarget: !apply);

            observer.OnPhase("worksheet-preflight");
            if (WorksheetExists(session, operation.WorksheetName!))
            {
                checks.Add(new TaskCheck("worksheet-name", false, "A worksheet with that name already exists; creating one never replaces it."));
                return ExcelSession.CloseAndProve(ref session, "worksheet preflight", checks)
                    ?? new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected, "That worksheet name is already in use.", Checks: checks);
            }

            checks.Add(new TaskCheck("worksheet-name", true, "No worksheet in the target workbook carries that name."));
            if (!apply)
            {
                changes.Add(new TaskChange("worksheet-create", operation.WorksheetName!, "Planned an empty worksheet."));
                return ExcelSession.CloseAndProve(ref session, "worksheet planning", checks, changes)
                    ?? new WorkbookExecutionOutcome(ExcelTaskStatus.Planned,
                        $"The name is free; applying would add an empty worksheet. Nothing was changed.", changes, checks);
            }

            observer.OnPhase("worksheet-create");
            mutationAttempted = true;
            AddWorksheet(session, operation.WorksheetName!);
            changes.Add(new TaskChange("worksheet-create", operation.WorksheetName!, "Added an empty worksheet."));

            observer.OnPhase("save");
            Invoke(session.TargetWorkbook, "Save");
            checks.Add(new TaskCheck("save", true, "Excel completed the requested save operation."));

            observer.OnPhase("primary-cleanup");
            var cleanupFailure = ExcelSession.CloseAndProve(ref session, "the worksheet creation save", checks, changes);
            if (cleanupFailure is not null) return cleanupFailure with { RetryReason = "Inspect workbook state before retrying." };

            if (plan.Request.WorkbookBinding != WorkbookBinding.UseOpen && !WorkbookRuntimeHelpers.CanOpenExclusively(targetPath))
            {
                checks.Add(new TaskCheck("file-lock", false, "The saved workbook remained locked after owned Excel cleanup."));
                return new WorkbookExecutionOutcome(ExcelTaskStatus.Unknown,
                    "Excel saved the workbook, but the owned Excel session did not release its file lock for verification.",
                    changes, checks, CanRetry: false, RetryReason: "Inspect the Excel process and file lock before retrying.");
            }

            observer.OnPhase("reopen-verification");
            var verified = pendingVerification!.Verify(targetPath, observer, verification =>
                WorksheetExists(verification, operation.WorksheetName!)
                    ? (true, new TaskCheck("reopen-verification", true, "The saved workbook reopened with the new worksheet present."))
                    : (false, new TaskCheck("reopen-verification", false, "The new worksheet was not present after reopening the saved workbook.")),
                out var reopenCheck);

            checks.Add(reopenCheck);
            if (!verified)
            {
                return new WorkbookExecutionOutcome(ExcelTaskStatus.Unknown,
                    "Excel saved the workbook, but reopen verification did not find the new worksheet.",
                    changes, checks, CanRetry: false, RetryReason: "Inspect the saved workbook before attempting another apply operation.");
            }

            return new WorkbookExecutionOutcome(ExcelTaskStatus.Completed,
                $"Added an empty worksheet, saved, and confirmed it after reopening.", changes, checks);
        }
        catch (Exception exception) when (exception is COMException or TargetInvocationException or InvalidOperationException or InvalidComObjectException)
        {
            checks.Add(new TaskCheck("worksheet-create", false, DescribeFailure("worksheet-create", exception)));
            var cleanupFailure = ExcelSession.CloseAndProve(ref session, "the failed worksheet creation", checks, changes);
            if (cleanupFailure is not null) return cleanupFailure;
            return new WorkbookExecutionOutcome(
                mutationAttempted ? ExcelTaskStatus.Unknown : ExcelTaskStatus.Rejected,
                mutationAttempted
                    ? "The worksheet creation failed after the change was attempted."
                    : "The worksheet creation was rejected before any change was attempted.",
                changes, checks, CanRetry: !mutationAttempted,
                RetryReason: mutationAttempted ? "Inspect workbook state before retrying." : null);
        }
        finally
        {
            session?.Close();
            pendingVerification?.Dispose();
        }
    }

    private static bool WorksheetExists(ExcelSession session, string worksheetName)
    {
        using var references = new ComReferenceScope();
        var sheets = references.Add(Get(session.TargetWorkbook, "Worksheets"));
        try
        {
            references.Add(Item(sheets, worksheetName));
            return true;
        }
        catch (Exception exception) when (exception is COMException or TargetInvocationException or InvalidOperationException)
        {
            return false;
        }
    }

    /// <summary>
    /// Adds one sheet after the last, so an added sheet never displaces the one the workbook opens
    /// on. Excel places a new sheet before the active one by default, which would quietly change
    /// what a person sees when they open the file.
    /// </summary>
    private static void AddWorksheet(ExcelSession session, string worksheetName)
    {
        using var references = new ComReferenceScope();
        var sheets = references.Add(Get(session.TargetWorkbook, "Worksheets"));
        var count = Convert.ToInt32(Get(sheets, "Count"), System.Globalization.CultureInfo.InvariantCulture);
        var last = references.Add(Item(sheets, count));
        var added = references.Add(ComAccess.Invoke(sheets, "Add", Type.Missing, last)
            ?? throw new InvalidOperationException("Excel did not return the new worksheet."));
        Set(added, "Name", worksheetName);
    }
}
