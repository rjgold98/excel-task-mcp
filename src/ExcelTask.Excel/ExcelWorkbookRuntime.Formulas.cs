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

    private static CopyRebindResult CopyReferenceWorksheet(ExcelSession session, string referenceSheetName, string newSheetName, Action<string> updatePhase)
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

        updatePhase("worksheet-copy-rebind");
        return RebindCopiedReferences(session, targetSheets, copiedSheet, references);
    }

    /// <summary>
    /// What happened to the references Excel rewrote when it copied the sheet.
    /// <paramref name="Rebound"/> names sheets pointed back inside the target;
    /// <paramref name="StillExternal"/> names sheets the target does not have, which therefore
    /// cannot be made internal and are reported instead.
    /// </summary>
    private sealed record CopyRebindResult(IReadOnlyList<string> Rebound, IReadOnlyList<string> StillExternal);

    /// <summary>
    /// Points the copied sheet's formulas back at the target workbook.
    ///
    /// Excel does not copy a worksheet's formulas verbatim. A formula reading another sheet in the
    /// source - =Data!A1 - becomes ='[source.xlsx]Data'!A1 in the destination, because the sheet it
    /// named is not there. The copy still calculates, which is what makes this so quiet: the numbers
    /// look right and are read out of the workbook the exhibit was copied FROM, so the new sheet
    /// reports the template's figures rather than the workbook it now lives in, and it breaks the
    /// day the template moves.
    ///
    /// Measured in the field before it was measured here: one copy produced 305 externally linked
    /// formula cells, and the follow-up repair fixed none of them.
    ///
    /// Where the target has a sheet of the same name, the prefix is removed and the formula binds
    /// locally - verified: the values change to the target's own numbers and the external link
    /// disappears. Where it does not, nothing is rewritten, because pointing a formula at a sheet
    /// that does not exist would turn a wrong number into a #REF. Those are named in the receipt so
    /// the caller learns it rather than discovering it a quarter later.
    /// </summary>
    private static CopyRebindResult RebindCopiedReferences(
        ExcelSession session,
        object targetSheets,
        object copiedSheet,
        ComReferenceScope references)
    {
        var sourceName = Get(session.ReferenceWorkbook, "Name") as string;
        if (string.IsNullOrEmpty(sourceName)) return new CopyRebindResult([], []);

        var used = references.Add(Get(copiedSheet, "UsedRange"));
        var formulas = GetOrNull(used, "Formula");
        var referenced = ExternalSheetNames(formulas, sourceName);
        if (referenced.Count == 0) return new CopyRebindResult([], []);

        List<string> rebound = [];
        List<string> stillExternal = [];
        foreach (var sheetName in referenced)
        {
            if (!WorksheetExists(targetSheets, sheetName, references))
            {
                stillExternal.Add(sheetName);
                continue;
            }

            // Text replacement rather than per-cell assignment. Assigning a formula back cell by
            // cell is what the field repair tried, and Excel re-resolved each one to the external
            // form; replacing the prefix in place leaves Excel nothing to re-resolve.
            // The quoted form carries doubled apostrophes, because that is how Excel wrote it; the
            // bare form only ever appears for a name that needed no quoting, so it has none.
            var quoted = $"'[{sourceName}]{sheetName.Replace("'", "''", StringComparison.Ordinal)}'!";
            var bare = $"[{sourceName}]{sheetName}!";
            var replacement = $"{QuoteSheetIfNeeded(sheetName)}!";
            Invoke(used, "Replace", quoted, replacement, XlPart, XlByRows, false);
            Invoke(used, "Replace", bare, replacement, XlPart, XlByRows, false);
            rebound.Add(sheetName);
        }

        return new CopyRebindResult(rebound, stillExternal);
    }

    private const int XlPart = 2;
    private const int XlByRows = 1;

    /// <summary>
    /// A sheet name written back into a formula, quoted exactly when Excel requires it.
    ///
    /// "Every character is alphanumeric" is not Excel's rule and reading it as one produced two
    /// wrong formulas. A name that starts with a digit needs quotes - <c>2024</c> and <c>2025</c>
    /// are ordinary fiscal-year tabs, and <c>=2024!A1</c> is not a formula Excel accepts. So does a
    /// name that is itself a cell or row/column reference: <c>A1</c>, <c>R1</c>, <c>C5</c>. And an
    /// apostrophe inside a quoted name has to be doubled, or the quotes close early.
    /// </summary>
    internal static string QuoteSheetIfNeeded(string sheetName) =>
        NeedsQuoting(sheetName)
            ? $"'{sheetName.Replace("'", "''", StringComparison.Ordinal)}'"
            : sheetName;

    private static bool NeedsQuoting(string sheetName)
    {
        if (sheetName.Length == 0) return true;
        if (!sheetName.All(character => char.IsLetterOrDigit(character) || character == '_')) return true;
        if (char.IsDigit(sheetName[0])) return true;
        return LooksLikeCellReference(sheetName);
    }

    /// <summary>
    /// Whether the name would be read as a reference rather than a sheet: A1-style, or the R/C
    /// forms Excel reserves whole. Deliberately generous - quoting a name that did not need it is
    /// still a valid formula, while failing to quote one that did is not.
    /// </summary>
    private static bool LooksLikeCellReference(string name)
    {
        if (name.Length == 1 && (name[0] is 'R' or 'C' or 'r' or 'c')) return true;

        var index = 0;
        while (index < name.Length && char.IsLetter(name[index])) index++;
        if (index == 0 || index > 3 || index == name.Length) return false;
        return name[index..].All(char.IsDigit);
    }

    /// <summary>
    /// Every sheet of the source workbook that the copied formulas now point at, read from the
    /// bulk formula array rather than cell by cell.
    /// </summary>
    internal static SortedSet<string> ExternalSheetNames(object? formulas, string sourceWorkbookName)
    {
        SortedSet<string> names = new(StringComparer.OrdinalIgnoreCase);
        var marker = $"[{sourceWorkbookName}]";

        void Scan(string? formula)
        {
            if (formula is null) return;
            var from = 0;
            while (true)
            {
                var start = formula.IndexOf(marker, from, StringComparison.OrdinalIgnoreCase);
                if (start < 0) return;
                var nameStart = start + marker.Length;
                var end = formula.IndexOf('!', nameStart);
                if (end < 0) return;
                // Un-double the apostrophes Excel doubled on the way in. A sheet named Payer's Data
                // arrives as '[source.xlsx]Payer''s Data'!, and reading it literally produced a
                // name that could never match the real sheet - so a fixable exhibit was left
                // external and the receipt named a worksheet that does not exist.
                var sheet = formula[nameStart..end].TrimEnd('\'').Replace("''", "'", StringComparison.Ordinal);
                if (sheet.Length > 0) names.Add(sheet);
                from = end + 1;
            }
        }

        if (formulas is Array array)
        {
            foreach (var value in array) Scan(value as string);
        }
        else
        {
            Scan(formulas as string);
        }

        return names;
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

    private static bool VerifySavedWorkbook(
        string path,
        string worksheetName,
        IReadOnlyList<ExpectedFormula> expectedRepairs,
        PendingVerification verification,
        IExcelWorkbookRuntimeObserver observer,
        out TaskCheck check) =>
        verification.Verify(path, observer, session =>
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
        }, out check);

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
