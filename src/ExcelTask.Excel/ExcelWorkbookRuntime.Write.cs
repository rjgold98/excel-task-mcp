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
    /// Writes constants, saves, and proves the result by reopening the saved file. Everything
    /// around the write itself - the save, the cleanup proof, the verification, the promotion - is
    /// the pipeline's; this method is only what writing constants actually means.
    /// </summary>
    private static WorkbookExecutionOutcome ExecuteWriteCore(ExcelTaskPlan plan, IExcelWorkbookRuntimeObserver observer)
    {
        var operation = plan.Request.Operation.WriteWorksheetValues!;
        return ExecuteMutation(plan, observer, "value-write", "The value write", context =>
        {
            context.OnPhase("write-preflight");
            var preflight = PreflightWorksheetExists(context.Session, operation.WorksheetName);
            context.Checks.AddRange(preflight.Checks);
            if (!preflight.IsFeasible)
            {
                return new MutationStep.Finish(
                    new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected, "The requested worksheet was not found.", Checks: context.Checks),
                    "write preflight");
            }

            // Plan says what would be written without touching the workbook, so a caller can see
            // the exact cells before authorizing an overwrite.
            if (!context.Apply)
            {
                context.Changes.Add(new TaskChange("value-write", $"{operation.WorksheetName}!{operation.Cells[0].Address}",
                    $"Planned {operation.Cells.Count} constant writes."));
                return new MutationStep.Finish(
                    new WorkbookExecutionOutcome(ExcelTaskStatus.Planned,
                        $"Planned {operation.Cells.Count} constant writes; no Excel changes were made.", context.Changes, context.Checks),
                    "write planning");
            }

            context.OnPhase("value-write");
            if (!ApplyWorksheetValues(context.Session, operation, context.MarkMutationAttempted, out var written, out var writeCheck))
            {
                context.Checks.Add(writeCheck);
                return new MutationStep.Finish(
                    new WorkbookExecutionOutcome(ExcelTaskStatus.Unknown,
                        "Excel did not store the requested values as written.", context.Changes, context.Checks,
                        CanRetry: false, RetryReason: "Inspect the workbook before retrying."),
                    "the value write");
            }

            context.Checks.Add(writeCheck);
            context.Changes.Add(new TaskChange("value-write", $"{operation.WorksheetName}!{written[0].Address}",
                $"Wrote {written.Count} constants."));

            return new MutationStep.SaveAndVerify(
                verification => VerifySavedValuesBody(verification, operation),
                $"Wrote {operation.Cells.Count} constants, saved, and verified them after reopening.",
                "Excel saved the workbook, but reopen verification did not confirm every written value.",
                Range: new WorksheetRangeReceipt(operation.WorksheetName, WrittenSpan(operation), false,
                    operation.Cells.Count, operation.Cells.Count,
                    [.. operation.Cells.Select(cell => new WorksheetCell(cell.Address, cell.Value))], Truncated: false));
        });
    }
    /// <summary>The A1 span the written cells occupy, for the receipt's range field.</summary>
    private static string WrittenSpan(NormalizedWriteWorksheetValuesOperation operation)
    {
        var first = operation.Cells[0].Address;
        var last = operation.Cells[^1].Address;
        return string.Equals(first, last, StringComparison.Ordinal) ? first : $"{first}:{last}";
    }

    private static (bool Verified, TaskCheck Check) VerifySavedValuesBody(
        ExcelSession session,
        NormalizedWriteWorksheetValuesOperation operation)
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
    }
}
