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
        var target = $"{operation.WorksheetName}!{operation.Range}";
        return ExecuteMutation(plan, observer, "number-format", "The number format", context =>
        {
            context.OnPhase("format-preflight");
            var preflight = PreflightWorksheetExists(context.Session, operation.WorksheetName);
            context.Checks.AddRange(preflight.Checks);
            if (!preflight.IsFeasible)
            {
                return new MutationStep.Finish(
                    new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected, "The requested worksheet was not found.", Checks: context.Checks),
                    "number format preflight");
            }

            // What is there now, so the caller can see what an Apply would replace. A format is
            // destructive in a way a value write is not - the old code is not recoverable from the
            // sheet afterwards. Excel returns null for a range whose cells do not share one format,
            // which is itself the answer.
            var existing = ReadNumberFormat(context.Session, operation);
            context.Checks.Add(new TaskCheck("current-format", true, existing is null
                ? "The range does not currently share one number format."
                : $"The range currently uses the number format {existing}."));

            if (!context.Apply)
            {
                context.Changes.Add(new TaskChange("number-format", target, $"Planned the number format {operation.NumberFormat}."));
                return new MutationStep.Finish(
                    new WorkbookExecutionOutcome(ExcelTaskStatus.Planned,
                        $"Applying would set {target} to the number format {operation.NumberFormat}. Nothing was changed.",
                        context.Changes, context.Checks),
                    "number format planning");
            }

            context.OnPhase("number-format");
            context.MarkMutationAttempted();
            if (!ApplyNumberFormat(context.Session, operation, out var stored, out var formatCheck))
            {
                context.Checks.Add(formatCheck);
                return new MutationStep.Finish(
                    new WorkbookExecutionOutcome(ExcelTaskStatus.Unknown,
                        $"Excel stored the number format {stored ?? "of mixed cells"} rather than the one requested; nothing was saved.",
                        context.Changes, context.Checks,
                        CanRetry: false, RetryReason: "Correct the format code and inspect the workbook before retrying."),
                    "the number format");
            }

            context.Checks.Add(formatCheck);
            context.Changes.Add(new TaskChange("number-format", target, $"Set the number format to {operation.NumberFormat}."));

            return new MutationStep.SaveAndVerify(
                verification => string.Equals(ReadNumberFormat(verification, operation), operation.NumberFormat, StringComparison.Ordinal)
                    ? (true, new TaskCheck("reopen-verification", true, "The saved workbook reopened with the requested number format across the range."))
                    : (false, new TaskCheck("reopen-verification", false, "The requested number format was not present across the range after reopening the saved workbook.")),
                $"Set {target} to the number format {operation.NumberFormat}, saved, and confirmed it after reopening.",
                "Excel saved the workbook, but reopen verification did not confirm the number format.");
        });
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
