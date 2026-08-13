using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using ExcelTask.Core;

namespace ExcelTask.Excel;

public sealed partial class ExcelWorkbookRuntime
{
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
        catch (Exception exception) when (ComAccess.IsComFailure(exception))
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
        catch (Exception exception) when (ComAccess.IsComFailure(exception))
        {
            return new WorksheetCopyPreflight(false, [new TaskCheck("target-worksheet", false, "Target worksheet feasibility could not be read.")]);
        }
    }

    private static bool WorksheetExists(object worksheets, string name, ComReferenceScope references)
    {
        var count = Convert.ToInt32(Get(worksheets, "Count"), CultureInfo.InvariantCulture);
        for (var index = 1; index <= count; index++)
        {
            var worksheet = references.Add(Item(worksheets, index));
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
        var referenceSheet = references.Add(Item(referenceSheets, referenceSheetName));
        updatePhase("worksheet-copy-target-sheets");
        var targetSheets = references.Add(Get(session.TargetWorkbook, "Worksheets"));
        var count = Convert.ToInt32(Get(targetSheets, "Count"), CultureInfo.InvariantCulture);
        updatePhase("worksheet-copy-target-anchor");
        var afterSheet = references.Add(Item(targetSheets, count));
        updatePhase("worksheet-copy-copy");
        Invoke(referenceSheet, "Copy", Type.Missing, afterSheet);
        updatePhase("worksheet-copy-copied-sheet");
        var copiedSheet = references.Add(Item(targetSheets, count + 1));
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
        var worksheet = references.Add(Item(worksheetCollection, sourceWorksheetName));
        var repairs = new List<ExpectedFormula>();
        var rangeResults = new List<RepairRangeResult>();

        foreach (var requestedRange in ranges)
        {
            var bounds = WorkbookRuntimeHelpers.GetBounds(requestedRange);

            var evidence = EvidenceBoundsFor(bounds);
            var evidenceAddress = $"{WorkbookRuntimeHelpers.ToA1Address(evidence.StartRow, evidence.StartColumn)}:{WorkbookRuntimeHelpers.ToA1Address(evidence.EndRow, evidence.EndColumn)}";

            var range = references.Add(Get(worksheet, "Range", evidenceAddress));
            var formulaGrid = WorkbookRuntimeHelpers.CreateFormulaGrid(Get(range, "FormulaR1C1"), evidence.RowCount, evidence.ColumnCount);
            var inferred = FormulaPatternAnalyzer.InferRepairs(formulaGrid);

            var repairCount = 0;
            foreach (var repair in inferred)
            {
                var row = evidence.StartRow + repair.RowIndex;
                var column = evidence.StartColumn + repair.ColumnIndex;
                if (row < bounds.StartRow || row > bounds.EndRow || column < bounds.StartColumn || column > bounds.EndColumn) continue;
                repairs.Add(new ExpectedFormula(row, column, repair.FormulaR1C1));
                repairCount++;
            }

            rangeResults.Add(new RepairRangeResult(requestedRange, repairCount));
        }

        return FormulaExecutionPlan.Create(kind, targetWorksheetName, repairs, rangeResults);
    }

    /// <summary>
    /// The rectangle read as evidence for a requested repair range: one cell wider on every side,
    /// clamped to the sheet.
    ///
    /// Inference needs the neighbours on each side, and a blank on the very edge of the requested
    /// range has one of them outside it. Reading only the requested range made such a cell
    /// unrepairable, so a caller who split a large area into chunks silently lost every gap that
    /// landed on a chunk boundary and still got Completed. Writes stay inside the requested range;
    /// only the reading widens.
    ///
    /// Separated from the COM path so the widening itself can be asserted. In the widest fixture in
    /// the suite every gap already has both neighbours inside the requested range, so deleting the
    /// +/-1 left the whole suite green while silent data loss came back.
    /// </summary>
    internal static FormulaRangeBounds EvidenceBoundsFor(FormulaRangeBounds bounds) => new(
        Math.Max(1, bounds.StartRow - 1),
        Math.Max(1, bounds.StartColumn - 1),
        Math.Min(1_048_576, bounds.EndRow + 1),
        Math.Min(16_384, bounds.EndColumn + 1));

    private static FormulaExecutionPlan AnalyzeFormulaExtension(object workbook, NormalizedExtendFormulaSeriesOperation task)
    {
        using var references = new ComReferenceScope();
        var sheets = references.Add(Get(workbook, "Worksheets"));
        var sheet = references.Add(Item(sheets, task.WorksheetName));
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
        var sheet = references.Add(Item(sheets, plan.WorksheetName));
        foreach (var group in plan.Repairs.GroupBy(repair => repair.FormulaR1C1, StringComparer.Ordinal))
        {
            foreach (var address in BatchAddresses(group))
            {
                var target = references.Add(Get(sheet, "Range", address));
                markMutationAttempted();
                Set(target, "FormulaR1C1", group.Key);
            }
        }
    }

    /// <summary>
    /// Excel rejects a Range address argument longer than 255 characters, so batches are bounded by
    /// the joined address length rather than by a cell count. A fixed count cannot be safe: the same
    /// number of cells produces a longer address further down the sheet, which made identical repairs
    /// succeed near row 1 and fail near row 2500.
    ///
    /// Internal rather than private so the bound can be asserted directly. Every repair fixture in
    /// the suite is small enough to produce one batch, so the split this exists for never ran under
    /// test: reintroducing a fixed cell count would have left the whole suite green while the
    /// recorded failure came back exactly as described above.
    /// </summary>
    internal static IEnumerable<string> BatchAddresses(IEnumerable<ExpectedFormula> repairs)
    {
        const int MaxAddressLength = 255;
        var builder = new StringBuilder();
        foreach (var repair in repairs)
        {
            var address = WorkbookRuntimeHelpers.ToA1Address(repair.Row, repair.Column);
            if (builder.Length > 0 && builder.Length + 1 + address.Length > MaxAddressLength)
            {
                yield return builder.ToString();
                builder.Clear();
            }
            if (builder.Length > 0) builder.Append(',');
            builder.Append(address);
        }
        if (builder.Length > 0) yield return builder.ToString();
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

    private static (bool Verified, TaskCheck Check) VerifySavedWorkbook(
        ExcelSession session,
        string worksheetName,
        IReadOnlyList<ExpectedFormula> expectedRepairs)
    {
        using var references = new ComReferenceScope();
        var sheets = references.Add(Get(session.TargetWorkbook, "Worksheets"));
        var sheet = references.Add(Item(sheets, worksheetName));

        // Read the whole bounding box of the repaired cells in one array rather than fetching
        // each cell across the COM boundary. Measured on this machine: 3,000 individual cell
        // reads cost about 4.9 seconds; the same cells as one range read cost 13 ms. The
        // per-call cost of COM, not the work, was the entire verification budget.
        ReadFormulaBox(references, sheet, expectedRepairs, out var box, out var origin);

        foreach (var expected in expectedRepairs)
        {
            var actual = box[expected.Row - origin.Row, expected.Column - origin.Column];
            if (!string.Equals(actual, expected.FormulaR1C1, StringComparison.Ordinal))
            {
                return (false, new TaskCheck("reopen-verification", false, "A repaired formula was not present after reopening the saved workbook."));
            }
        }

        return (true, new TaskCheck("reopen-verification", true, $"Saved workbook reopened with the requested worksheet and {expectedRepairs.Count} requested formulas."));
    }

    /// <summary>
    /// Reads the R1C1 formulas of the smallest rectangle covering every expected repair, as one
    /// COM call, into a zero-based array indexed off <paramref name="origin"/>. A single bulk read
    /// is hundreds of times cheaper than one call per cell. The rectangle is bounded by the same
    /// caps the repairs are, so it cannot marshal an unbounded array.
    ///
    /// Void, not bool. It returned bool and the caller reported "could not be read back" on false -
    /// but both exits returned true, so that branch had never executed. The only thing that can go
    /// wrong here is a COM throw from the two reads below, which the outer handler already turns
    /// into a failure check carrying the phase and the exception; converting it to a bool would be
    /// a worse diagnostic for the same event. Void keeps it honest: a genuine failure condition
    /// added later cannot be introduced without changing the signature and forcing the caller to
    /// decide what to do about it.
    /// </summary>
    private static void ReadFormulaBox(
        ComReferenceScope references,
        object sheet,
        IReadOnlyList<ExpectedFormula> expectedRepairs,
        out string?[,] box,
        out (int Row, int Column) origin)
    {
        box = new string?[0, 0];
        origin = (1, 1);
        if (expectedRepairs.Count == 0) return;

        var minRow = expectedRepairs.Min(repair => repair.Row);
        var maxRow = expectedRepairs.Max(repair => repair.Row);
        var minColumn = expectedRepairs.Min(repair => repair.Column);
        var maxColumn = expectedRepairs.Max(repair => repair.Column);
        origin = (minRow, minColumn);

        var address = $"{WorkbookRuntimeHelpers.ToA1Address(minRow, minColumn)}:{WorkbookRuntimeHelpers.ToA1Address(maxRow, maxColumn)}";
        var range = references.Add(Get(sheet, "Range", address));
        var value = Get(range, "FormulaR1C1");

        var rows = maxRow - minRow + 1;
        var columns = maxColumn - minColumn + 1;
        box = new string?[rows, columns];
        if (value is Array values)
        {
            for (var row = 0; row < rows; row++)
            {
                for (var column = 0; column < columns; column++)
                {
                    box[row, column] = values.GetValue(row + values.GetLowerBound(0), column + values.GetLowerBound(1)) as string;
                }
            }
        }
        else
        {
            // A single-cell range returns the scalar rather than a one-element array.
            box[0, 0] = value as string;
        }
    }

    internal sealed record ExpectedFormula(int Row, int Column, string FormulaR1C1);

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
}
