using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using ExcelTask.Core;

namespace ExcelTask.Excel;

public sealed partial class ExcelWorkbookRuntime
{
    /// <summary>
    /// Returns the contents of one bounded range.
    ///
    /// The whole range is fetched in a single array rather than a call per cell, which is the
    /// difference between milliseconds and seconds: 3,000 individual cell reads measured 4,864 ms
    /// against 13 ms for the same cells read at once. Blank cells are dropped rather than reported
    /// as empty, so a sparse range costs almost nothing to return.
    /// </summary>
    private static WorkbookExecutionOutcome ExecuteReadCore(ExcelTaskPlan plan, IExcelWorkbookRuntimeObserver observer)
    {
        var operation = plan.Request.Operation.ReadWorksheetRange!;
        var checks = new List<TaskCheck>();

        string targetPath;
        try
        {
            targetPath = WorkbookRuntimeHelpers.NormalizePath(plan.Request.TargetWorkbookPath);
            WorkbookRuntimeHelpers.EnsureReadableWorkbook(targetPath, "Target workbook");
        }
        catch (InvalidOperationException)
        {
            return new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected,
                "The target workbook cannot be read.",
                Checks: [new TaskCheck("workbook-inputs", false, "The target workbook path is not readable.")]);
        }

        var stampBefore = ReadFileStamp(targetPath);
        ExcelSession? session = null;
        WorksheetRangeReceipt? receipt = null;
        try
        {
            observer.OnPhase("session-open");
            // A read attaches to a live workbook when asked to, because reading what the user is
            // looking at right now is a legitimate request and opening a second copy would answer
            // about the file on disk instead.
            session = ExcelSession.Open(plan.Request, observer, readOnlyTarget: plan.Request.WorkbookBinding != WorkbookBinding.UseOpen);
            observer.OnPhase("range-read");

            using (var references = new ComReferenceScope())
            {
                var sheets = references.Add(Get(session.TargetWorkbook, "Worksheets"));
                object sheet;
                try
                {
                    sheet = references.Add(Item(sheets, operation.WorksheetName));
                }
                catch (Exception exception) when (ComAccess.IsComFailure(exception))
                {
                    checks.Add(new TaskCheck("worksheet", false, "The requested worksheet does not exist in this workbook. Run AuditWorkbookFlows to list the worksheets."));
                    return ExcelSession.CloseAndProve(ref session, "the read", checks)
                        ?? new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected, "The requested worksheet was not found.", Checks: checks);
                }

                var bounds = WorkbookRuntimeHelpers.GetBounds(operation.Range);
                var range = references.Add(Get(sheet, "Range", operation.Range.ToString()));
                // Value2 rather than Text: Text is the displayed string of a single cell and
                // returns null for a multi-cell range, and it would also hand back "####" for a
                // column too narrow to show its number.
                var value = Get(range, operation.Formulas ? "FormulaR1C1" : "Value2");

                var cells = new List<WorksheetCell>();
                var nonEmpty = 0;
                for (var row = 0; row < bounds.RowCount; row++)
                {
                    for (var column = 0; column < bounds.ColumnCount; column++)
                    {
                        var cell = value is Array values
                            ? values.GetValue(row + values.GetLowerBound(0), column + values.GetLowerBound(1))
                            : value;
                        var text = Render(cell);
                        if (text.Length == 0) continue;

                        nonEmpty++;
                        cells.Add(new WorksheetCell(
                            WorkbookRuntimeHelpers.ToA1Address(bounds.StartRow + row, bounds.StartColumn + column),
                            text));
                    }
                }

                receipt = new WorksheetRangeReceipt(
                    operation.WorksheetName,
                    operation.Range.ToString(),
                    operation.Formulas,
                    bounds.RowCount * bounds.ColumnCount,
                    nonEmpty,
                    cells,
                    Truncated: false);
                checks.Add(new TaskCheck("range-read", true,
                    $"Read {bounds.RowCount * bounds.ColumnCount} cell(s); {nonEmpty} held content."));
            }

            observer.OnPhase("read-cleanup");
            var cleanupFailure = ExcelSession.CloseAndProve(ref session, "the read", checks);
            if (cleanupFailure is not null) return cleanupFailure;
        }
        catch (Exception exception) when (ComAccess.IsComFailure(exception))
        {
            var cleanupFailed = false;
            if (session is not null)
            {
                try { cleanupFailed = !session.Close(); }
                catch (Exception cleanupException) when (ComAccess.IsComFailure(cleanupException)) { cleanupFailed = true; }
                session = null;
            }

            checks.Add(new TaskCheck("range-read", false, "The requested range could not be read."));
            return cleanupFailed
                ? new WorkbookExecutionOutcome(ExcelTaskStatus.Unknown, "The read failed and could not prove owned Excel cleanup.", Checks: checks, CanRetry: false, RetryReason: "Inspect the owned Excel process before retrying.")
                : new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected, "The requested range could not be read.", Checks: checks);
        }
        finally
        {
            session?.Close();
        }

        // A read must leave the file exactly as it found it, and says so from evidence rather than
        // from intent - the same proof the audit gives.
        var unchanged = ReadFileStamp(targetPath) == stampBefore;
        checks.Add(new TaskCheck("workbook-unchanged", unchanged, unchanged
            ? "The workbook's size and timestamp are identical before and after the read."
            : "The workbook's size or timestamp changed during the read; something else wrote to it."));

        return new WorkbookExecutionOutcome(
            plan.Request.Mode == ExcelTaskMode.Plan ? ExcelTaskStatus.Planned : ExcelTaskStatus.Completed,
            $"Read {receipt!.CellsInRange} cell(s) from {operation.WorksheetName}!{operation.Range}; {receipt.NonEmptyCells} held content.",
            Checks: checks,
            Range: receipt);
    }

    /// <summary>
    /// Renders a cell the way a person reading the sheet would see it, without locale surprises:
    /// dates as ISO, numbers round-tripped, everything else as its string form. Blank and empty
    /// string both render empty so the caller sees one notion of "nothing here".
    /// </summary>
    private static string Render(object? value) => value switch
    {
        null => string.Empty,
        string text => text,
        bool flag => flag ? "TRUE" : "FALSE",
        DateTime moment => moment.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture).Replace(" 00:00:00", string.Empty, StringComparison.Ordinal),
        double number => number.ToString("R", CultureInfo.InvariantCulture),
        _ => Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty
    };
}
