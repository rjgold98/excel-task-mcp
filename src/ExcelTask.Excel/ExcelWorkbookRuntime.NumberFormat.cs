using System.Reflection;
using System.Runtime.InteropServices;
using ExcelTask.Core;

namespace ExcelTask.Excel;

public sealed partial class ExcelWorkbookRuntime
{
    /// <summary>
    /// Sets one number format code across one bounded range, and proves it took.
    ///
    /// Unlike every other mutation here, this one is a single COM assignment no matter how many
    /// cells it covers - Excel applies a format to a whole range at once - so there is no batching
    /// and no per-cell cost. What it does need is the read-back, because Excel is free to store
    /// something other than what it was given: an unrecognized code can be kept verbatim, coerced,
    /// or rejected outright, and all three look identical from the caller's side of the assignment.
    /// Reading it back is what turns "we sent a format" into "the sheet holds this format".
    /// </summary>
    private static WorkbookExecutionOutcome ExecuteNumberFormatCore(ExcelTaskPlan plan, IExcelWorkbookRuntimeObserver observer)
    {
        var operation = plan.Request.Operation.SetNumberFormat!;
        var apply = plan.Request.Mode == ExcelTaskMode.Apply;
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
            if (apply && plan.Request.Save == SaveMode.Same)
            {
                WorkbookRuntimeHelpers.EnsureWritableSameTarget(WorkbookRuntimeHelpers.NormalizePath(plan.Request.TargetWorkbookPath));
            }
        }
        catch (InvalidOperationException exception)
        {
            return new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected, "Workbook inputs cannot be safely executed.",
                Checks: [new TaskCheck("workbook-inputs", false, exception.Message)]);
        }

        var savedPath = WorkbookRuntimeHelpers.NormalizePath(plan.Request.Save == SaveMode.Copy
            ? plan.Request.OutputWorkbookPath ?? throw new InvalidOperationException("Copy output path is required.")
            : plan.Request.TargetWorkbookPath);
        var target = $"{operation.WorksheetName}!{operation.Range}";

        try
        {
            if (apply) pendingVerification = PendingVerification.Begin(observer);

            observer.OnPhase("session-open");
            session = ExcelSession.Open(plan.Request, observer, readOnlyTarget: !apply);

            observer.OnPhase("format-preflight");
            var preflight = PreflightWorksheetExists(session, operation.WorksheetName);
            checks.AddRange(preflight.Checks);
            if (!preflight.IsFeasible)
            {
                return ExcelSession.CloseAndProve(ref session, "number format preflight", checks)
                    ?? new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected, "The requested worksheet was not found.", Checks: checks);
            }

            // What is there now, so the caller can see what an Apply would replace. Excel returns
            // null for a range whose cells do not share one format, which is itself the answer.
            var existing = ReadNumberFormat(session, operation);
            checks.Add(new TaskCheck("current-format", true, existing is null
                ? "The range does not currently share one number format."
                : $"The range currently uses the number format {existing}."));

            if (!apply)
            {
                changes.Add(new TaskChange("number-format", target, $"Planned the number format {operation.NumberFormat}."));
                return ExcelSession.CloseAndProve(ref session, "number format planning", checks, changes)
                    ?? new WorkbookExecutionOutcome(ExcelTaskStatus.Planned,
                        $"Applying would set {target} to the number format {operation.NumberFormat}. Nothing was changed.",
                        changes, checks);
            }

            observer.OnPhase("number-format");
            mutationAttempted = true;
            if (!ApplyNumberFormat(session, operation, out var stored, out var formatCheck))
            {
                checks.Add(formatCheck);
                return ExcelSession.CloseAndProve(ref session, "the number format", checks, changes)
                    ?? new WorkbookExecutionOutcome(ExcelTaskStatus.Unknown,
                        $"Excel stored the number format {stored ?? "of mixed cells"} rather than the one requested; nothing was saved.",
                        changes, checks, CanRetry: false, RetryReason: "Correct the format code and inspect the workbook before retrying.");
            }

            checks.Add(formatCheck);
            changes.Add(new TaskChange("number-format", target, $"Set the number format to {operation.NumberFormat}."));

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
            var cleanupFailure = ExcelSession.CloseAndProve(ref session, "the number format save", checks, changes);
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
            var verified = pendingVerification!.Verify(verificationPath, observer, verification =>
                string.Equals(ReadNumberFormat(verification, operation), operation.NumberFormat, StringComparison.Ordinal)
                    ? (true, new TaskCheck("reopen-verification", true, "The saved workbook reopened with the requested number format across the range."))
                    : (false, new TaskCheck("reopen-verification", false, "The requested number format was not present across the range after reopening the saved workbook.")),
                out var reopenCheck);

            checks.Add(reopenCheck);
            if (!verified)
            {
                AddStagingCleanupCheck(stagingPath, checks);
                return new WorkbookExecutionOutcome(ExcelTaskStatus.Unknown,
                    "Excel saved the workbook, but reopen verification did not confirm the number format.",
                    changes, checks, CanRetry: false, RetryReason: "Inspect the saved workbook before attempting another apply operation.");
            }

            if (stagingPath is not null)
            {
                observer.OnPhase("copy-promotion");
                WorkbookRuntimeHelpers.PromoteStaging(stagingPath, savedPath, plan.Request.OverwriteConfirmed);
                stagingPath = null;
                changes.Add(new TaskChange("copy-promotion", "workbook", "Promoted the verified staging workbook to the requested output path."));
            }

            return new WorkbookExecutionOutcome(ExcelTaskStatus.Completed,
                $"Set {target} to the number format {operation.NumberFormat}, saved, and confirmed it after reopening.",
                changes, checks);
        }
        catch (Exception exception) when (ComAccess.IsComFailure(exception))
        {
            checks.Add(new TaskCheck("number-format", false, DescribeFailure("number-format", exception)));
            var cleanupFailure = ExcelSession.CloseAndProve(ref session, "the failed number format", checks, changes);
            if (cleanupFailure is not null) return cleanupFailure;
            return new WorkbookExecutionOutcome(
                mutationAttempted ? ExcelTaskStatus.Unknown : ExcelTaskStatus.Rejected,
                mutationAttempted
                    ? "The number format failed after the change was attempted."
                    : "The number format was rejected before any change was attempted.",
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

    /// <summary>
    /// The range's number format, or null when its cells do not all share one. Excel answers a
    /// mixed range with null rather than an error, so null is a fact about the sheet rather than a
    /// failure to read it.
    /// </summary>
    private static string? ReadNumberFormat(ExcelSession session, NormalizedSetNumberFormatOperation operation)
    {
        using var references = new ComReferenceScope();
        var sheets = references.Add(Get(session.TargetWorkbook, "Worksheets"));
        var sheet = references.Add(Item(sheets, operation.WorksheetName));
        var range = references.Add(Get(sheet, "Range", operation.Range.ToString()));
        return GetOrNull(range, "NumberFormat") as string;
    }

    private static bool ApplyNumberFormat(
        ExcelSession session,
        NormalizedSetNumberFormatOperation operation,
        out string? stored,
        out TaskCheck check)
    {
        using (var references = new ComReferenceScope())
        {
            var sheets = references.Add(Get(session.TargetWorkbook, "Worksheets"));
            var sheet = references.Add(Item(sheets, operation.WorksheetName));
            var range = references.Add(Get(sheet, "Range", operation.Range.ToString()));
            Set(range, "NumberFormat", operation.NumberFormat);
        }

        stored = ReadNumberFormat(session, operation);
        if (!string.Equals(stored, operation.NumberFormat, StringComparison.Ordinal))
        {
            check = new TaskCheck("number-format", false, stored is null
                ? "Excel did not apply one number format across the whole range."
                : $"Excel stored the number format {stored} rather than the one requested.");
            return false;
        }

        check = new TaskCheck("number-format", true, $"Applied the number format {operation.NumberFormat} and read it back unchanged across the range.");
        return true;
    }
}
