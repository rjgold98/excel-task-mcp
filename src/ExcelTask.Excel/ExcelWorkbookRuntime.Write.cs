using System.Globalization;
using ExcelTask.Core;

namespace ExcelTask.Excel;

public sealed partial class ExcelWorkbookRuntime
{
    /// <summary>
    /// Turns the caller's text into the value Excel should hold.
    ///
    /// This matters more than it looks. A model that sends "1000" means the number, and storing it
    /// as text would leave a cell that looks right and breaks every SUM above it - the exact quiet
    /// wrongness this design exists to avoid. So numbers and booleans are converted, and anything
    /// else is written as text.
    ///
    /// Dates are deliberately *not* parsed. "3-4" is March 4th to Excel and a label to a person, and
    /// nothing in the request says which was meant. Writing it as text is the answer that can be
    /// seen and corrected; guessing is the one that cannot.
    /// </summary>
    private static object? ToCellValue(string value)
    {
        if (value.Length == 0) return null;
        if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) &&
            !double.IsNaN(number) && !double.IsInfinity(number))
        {
            return number;
        }

        if (string.Equals(value, "TRUE", StringComparison.OrdinalIgnoreCase)) return true;
        if (string.Equals(value, "FALSE", StringComparison.OrdinalIgnoreCase)) return false;
        return value;
    }

    /// <summary>
    /// Writes the requested constants, then reads every one of them back and compares.
    ///
    /// The read-back is the whole point. Excel coerces on assignment - a text value can become a
    /// date, a number can pick up a format - and a receipt that reported "written" without looking
    /// would be reporting the request rather than the result. Comparing against what Excel actually
    /// holds turns any coercion into a visible mismatch instead of a silent one.
    /// </summary>
    private static bool ApplyWorksheetValues(
        ExcelSession session,
        NormalizedWriteWorksheetValuesOperation operation,
        Action markMutationAttempted,
        out IReadOnlyList<WorksheetCell> written,
        out TaskCheck check)
    {
        using var references = new ComReferenceScope();
        var sheets = references.Add(Get(session.TargetWorkbook, "Worksheets"));
        var sheet = references.Add(Item(sheets, operation.WorksheetName));

        foreach (var cell in operation.Cells)
        {
            var target = references.Add(Get(sheet, "Range", cell.Address));
            markMutationAttempted();
            Set(target, "Value2", ToCellValue(cell.Value));
        }

        var readBack = new List<WorksheetCell>(operation.Cells.Count);
        foreach (var cell in operation.Cells)
        {
            var target = references.Add(Get(sheet, "Range", cell.Address));
            var actual = Render(GetOrNull(target, "Value2"));
            if (!string.Equals(actual, cell.Value, StringComparison.Ordinal))
            {
                written = [];
                check = new TaskCheck("value-write", false,
                    $"Excel stored something other than the requested value in {cell.Address}; no further cells were confirmed.");
                return false;
            }

            readBack.Add(new WorksheetCell(cell.Address, actual));
        }

        written = readBack;
        check = new TaskCheck("value-write", true, $"Wrote {operation.Cells.Count} cell(s) and read every one back unchanged.");
        return true;
    }

    /// <summary>
    /// Writes constants, saves, and proves the result by reopening the saved file - the same
    /// save-and-verify shape as a formula repair, because a write that is not read back off disk
    /// has only been attempted.
    /// </summary>
    private static WorkbookExecutionOutcome ExecuteWriteCore(ExcelTaskPlan plan, IExcelWorkbookRuntimeObserver observer)
    {
        var operation = plan.Request.Operation.WriteWorksheetValues!;
        var checks = new List<TaskCheck>();
        var changes = new List<TaskChange>();
        ExcelSession? session = null;
        PendingVerification? pendingVerification = null;
        string? stagingPath = null;
        var mutationAttempted = false;

        try
        {
            WorkbookRuntimeHelpers.EnsureReadableWorkbook(WorkbookRuntimeHelpers.NormalizePath(plan.Request.TargetWorkbookPath), "Target workbook");
            if (plan.Request.Save == SaveMode.Copy) WorkbookRuntimeHelpers.EnsureWritableCopyOutput(plan.Request.OutputWorkbookPath);
        }
        catch (InvalidOperationException)
        {
            return new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected, "Workbook inputs cannot be safely executed.",
                Checks: [new TaskCheck("workbook-inputs", false, "Workbook paths or output location are not ready for execution.")]);
        }

        var savedPath = WorkbookRuntimeHelpers.NormalizePath(plan.Request.Save == SaveMode.Copy
            ? plan.Request.OutputWorkbookPath ?? throw new InvalidOperationException("Copy output path is required.")
            : plan.Request.TargetWorkbookPath);

        try
        {
            if (plan.Request.Mode == ExcelTaskMode.Apply) pendingVerification = PendingVerification.Begin(observer);

            observer.OnPhase("session-open");
            session = ExcelSession.Open(plan.Request, observer, readOnlyTarget: plan.Request.Mode == ExcelTaskMode.Plan);

            observer.OnPhase("write-preflight");
            var preflight = PreflightWorksheetExists(session, operation.WorksheetName);
            checks.AddRange(preflight.Checks);
            if (!preflight.IsFeasible)
            {
                return ExcelSession.CloseAndProve(ref session, "write preflight", checks)
                    ?? new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected, "The requested worksheet was not found.", Checks: checks);
            }

            // Plan says what would be written without touching the workbook, so a caller can see the
            // exact cells before authorizing an overwrite.
            if (plan.Request.Mode == ExcelTaskMode.Plan)
            {
                changes.Add(new TaskChange("value-write", $"{operation.WorksheetName}!{operation.Cells[0].Address}",
                    $"Planned {operation.Cells.Count} constant writes."));
                return ExcelSession.CloseAndProve(ref session, "write planning", checks, changes)
                    ?? new WorkbookExecutionOutcome(ExcelTaskStatus.Planned,
                        $"Planned {operation.Cells.Count} constant writes; no Excel changes were made.", changes, checks);
            }

            observer.OnPhase("value-write");
            if (!ApplyWorksheetValues(session, operation, () => mutationAttempted = true, out var written, out var writeCheck))
            {
                checks.Add(writeCheck);
                return ExcelSession.CloseAndProve(ref session, "the value write", checks, changes)
                    ?? new WorkbookExecutionOutcome(ExcelTaskStatus.Unknown,
                        "Excel did not store the requested values as written.", changes, checks,
                        CanRetry: false, RetryReason: "Inspect the workbook before retrying.");
            }

            checks.Add(writeCheck);
            changes.Add(new TaskChange("value-write", $"{operation.WorksheetName}!{written[0].Address}",
                $"Wrote {written.Count} constants."));

            observer.OnPhase("save");
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
            observer.OnPhase("primary-cleanup");
            var cleanupFailure = ExcelSession.CloseAndProve(ref session, "the value write save", checks, changes);
            if (cleanupFailure is not null)
            {
                AddStagingCleanupCheck(stagingPath, checks);
                return cleanupFailure with { RetryReason = "Inspect workbook state before retrying." };
            }

            var verificationPath = stagingPath ?? savedPath;
            if (plan.Request.WorkbookBinding != WorkbookBinding.UseOpen && !WorkbookRuntimeHelpers.CanOpenExclusively(verificationPath))
            {
                checks.Add(new TaskCheck("file-lock", false, "The saved workbook remained locked after owned Excel cleanup."));
                AddStagingCleanupCheck(stagingPath, checks);
                return new WorkbookExecutionOutcome(ExcelTaskStatus.Unknown,
                    "Excel saved the workbook, but the owned Excel session did not release its file lock for verification.",
                    changes, checks, CanRetry: false, RetryReason: "Inspect the Excel process and file lock before retrying.");
            }

            observer.OnPhase("reopen-verification");
            if (!VerifySavedValues(verificationPath, operation, pendingVerification!, observer, out var reopenCheck))
            {
                checks.Add(reopenCheck);
                AddStagingCleanupCheck(stagingPath, checks);
                return new WorkbookExecutionOutcome(ExcelTaskStatus.Unknown,
                    "Excel saved the workbook, but reopen verification did not confirm every written value.",
                    changes, checks, CanRetry: false, RetryReason: "Inspect the saved workbook before attempting another apply operation.");
            }

            checks.Add(reopenCheck);
            if (stagingPath is not null)
            {
                observer.OnPhase("copy-promotion");
                WorkbookRuntimeHelpers.PromoteStaging(stagingPath, savedPath, plan.Request.OverwriteConfirmed);
                stagingPath = null;
                changes.Add(new TaskChange("copy-promotion", "workbook", "Promoted the verified staging workbook to the requested output path."));
            }

            return new WorkbookExecutionOutcome(ExcelTaskStatus.Completed,
                $"Wrote {operation.Cells.Count} constants, saved, and verified them after reopening.",
                changes, checks,
                Range: new WorksheetRangeReceipt(operation.WorksheetName, WrittenSpan(operation), false,
                    operation.Cells.Count, operation.Cells.Count,
                    [.. operation.Cells.Select(cell => new WorksheetCell(cell.Address, cell.Value))], Truncated: false));
        }
        catch (Exception exception) when (ComAccess.IsComFailure(exception))
        {
            checks.Add(new TaskCheck("value-write", false, DescribeFailure("value-write", exception)));
            var cleanupFailure = ExcelSession.CloseAndProve(ref session, "the failed value write", checks, changes);
            if (cleanupFailure is not null) return cleanupFailure;
            return new WorkbookExecutionOutcome(
                mutationAttempted ? ExcelTaskStatus.Unknown : ExcelTaskStatus.Rejected,
                mutationAttempted
                    ? "The value write failed after changes were attempted."
                    : "The value write was rejected before any change was attempted.",
                changes, checks, CanRetry: !mutationAttempted,
                RetryReason: mutationAttempted ? "Inspect workbook state before retrying." : null);
        }
        finally
        {
            session?.Close();
            pendingVerification?.Dispose();
            if (stagingPath is not null) _ = WorkbookRuntimeHelpers.TryDeleteStaging(stagingPath);
        }
    }

    /// <summary>The A1 span the written cells occupy, for the receipt's range field.</summary>
    private static string WrittenSpan(NormalizedWriteWorksheetValuesOperation operation)
    {
        var first = operation.Cells[0].Address;
        var last = operation.Cells[^1].Address;
        return string.Equals(first, last, StringComparison.Ordinal) ? first : $"{first}:{last}";
    }

    private static bool VerifySavedValues(
        string path,
        NormalizedWriteWorksheetValuesOperation operation,
        PendingVerification verification,
        IExcelWorkbookRuntimeObserver observer,
        out TaskCheck check) =>
        verification.Verify(path, observer, session =>
        {
            using var references = new ComReferenceScope();
            var sheets = references.Add(Get(session.TargetWorkbook, "Worksheets"));
            var sheet = references.Add(Item(sheets, operation.WorksheetName));
            foreach (var cell in operation.Cells)
            {
                var target = references.Add(Get(sheet, "Range", cell.Address));
                if (!string.Equals(Render(GetOrNull(target, "Value2")), cell.Value, StringComparison.Ordinal))
                {
                    return (false, new TaskCheck("reopen-verification", false,
                        "A written value was not present after reopening the saved workbook."));
                }
            }

            return (true, new TaskCheck("reopen-verification", true,
                $"Saved workbook reopened with all {operation.Cells.Count} written values."));
        }, out check);
}
