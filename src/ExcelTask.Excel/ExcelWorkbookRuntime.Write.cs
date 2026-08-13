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
        out IReadOnlyList<string> replacedFormulas,
        out IReadOnlyList<string> changedConstants,
        out TaskCheck check)
    {
        using var references = new ComReferenceScope();
        var sheets = references.Add(Get(session.TargetWorkbook, "Worksheets"));
        var sheet = references.Add(Item(sheets, operation.WorksheetName));
        var priorConstants = new List<string>();
        changedConstants = priorConstants;

        // Looked at before anything is written, because afterwards the evidence is gone. Replacing
        // a formula with a constant is legitimate - hardcoding a figure is a real thing finance
        // does - so this reports rather than refuses. What it will not do is let a caller destroy a
        // live calculation and read a receipt that says only "wrote 1 constant".
        //
        // The prior value travels with it. A simulation could tell its user a formula had not been
        // replaced but not what had actually changed, which is the question a person asks next:
        // "469750.25 to 471000" is the answer, and the cell was named in the request, so its former
        // contents are responsive rather than incidental.
        var overwritten = new List<string>();
        foreach (var cell in operation.Cells)
        {
            var target = references.Add(Get(sheet, "Range", cell.Address));
            var priorFormula = GetOrNull(target, "Formula") as string;
            var priorText = Render(GetOrNull(target, "Value2"));
            if (priorFormula is not null && priorFormula.StartsWith('='))
            {
                overwritten.Add($"{cell.Address} (was {Abbreviate(priorFormula)})");
            }
            else if (!string.Equals(priorText, cell.Value, StringComparison.Ordinal))
            {
                priorConstants.Add($"{cell.Address}: {Abbreviate(priorText)} -> {Abbreviate(cell.Value)}");
            }
        }

        replacedFormulas = overwritten;

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

    /// <summary>Names the cells whose formulas the write replaced, or says none were.</summary>
    private static TaskCheck DescribeReplacedFormulas(IReadOnlyList<string> replaced) => replaced.Count == 0
        ? new TaskCheck("formulas-replaced", true, "No cell written to held a formula.")
        : new TaskCheck("formulas-replaced", false,
            $"{replaced.Count} cell(s) held a formula that this write replaced with a constant: {string.Join("; ", replaced.Take(10))}.");

    /// <summary>What each cell actually changed from, so a caller can say what moved and not only that something did.</summary>
    private static TaskCheck DescribeChangedConstants(IReadOnlyList<string> changed) => changed.Count == 0
        ? new TaskCheck("prior-values", true, "Every cell written already held the value it was given.")
        : new TaskCheck("prior-values", true, $"Replaced {changed.Count} existing value(s): {string.Join("; ", changed.Take(10))}.");

    /// <summary>Bounded so one pasted paragraph in a cell cannot crowd out the rest of the receipt.</summary>
    private static string Abbreviate(string value) => value.Length <= 40 ? value : $"{value[..40]}...";

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
            if (!ApplyWorksheetValues(context.Session, operation, context.MarkMutationAttempted, out var written, out var replacedFormulas, out var changedConstants, out var writeCheck))
            {
                context.Checks.Add(writeCheck);
                context.Checks.Add(DescribeReplacedFormulas(replacedFormulas));
                context.Checks.Add(DescribeChangedConstants(changedConstants));
                return new MutationStep.Finish(
                    new WorkbookExecutionOutcome(ExcelTaskStatus.Unknown,
                        "Excel did not store the requested values as written.", context.Changes, context.Checks,
                        CanRetry: false, RetryReason: "Inspect the workbook before retrying."),
                    "the value write");
            }

            context.Checks.Add(writeCheck);
            context.Checks.Add(DescribeReplacedFormulas(replacedFormulas));
            context.Checks.Add(DescribeChangedConstants(changedConstants));
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

    /// <summary>
    /// Writes caller-supplied A1 formulas as a distinct operation from constant writes. The
    /// formula text is never placed in the receipt; the immediate read-back and the shared
    /// save/reopen transaction prove only that Excel stored the requested formulas.
    /// </summary>
    private static WorkbookExecutionOutcome ExecuteFormulaWriteCore(ExcelTaskPlan plan, IExcelWorkbookRuntimeObserver observer)
    {
        var operation = plan.Request.Operation.WriteWorksheetFormulas!;
        return ExecuteMutation(plan, observer, "formula-write", "The formula write", context =>
        {
            context.OnPhase("formula-write-preflight");
            var preflight = PreflightWorksheetExists(context.Session, operation.WorksheetName);
            context.Checks.AddRange(preflight.Checks);
            if (!preflight.IsFeasible)
            {
                return new MutationStep.Finish(
                    new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected, "The requested worksheet was not found.", Checks: context.Checks),
                    "formula write preflight");
            }

            if (!context.Apply)
            {
                context.Changes.Add(new TaskChange("formula-write", $"{operation.WorksheetName}!{operation.Cells[0].Address}",
                    $"Planned {operation.Cells.Count} caller-supplied formula writes."));
                return new MutationStep.Finish(
                    new WorkbookExecutionOutcome(
                        ExcelTaskStatus.Planned,
                        $"Planned {operation.Cells.Count} formula writes; no Excel changes were made.",
                        context.Changes,
                        context.Checks),
                    "formula write planning");
            }

            context.OnPhase("formula-write");
            var formulaCells = CreateFormulaWriteCells(operation);
            if (!ApplyWorksheetFormulas(context.Session, formulaCells, context.MarkMutationAttempted, out var formulaCheck))
            {
                context.Checks.Add(formulaCheck);
                return new MutationStep.Finish(
                    new WorkbookExecutionOutcome(
                        ExcelTaskStatus.Unknown,
                        "Excel did not store every requested formula as written.",
                        context.Changes,
                        context.Checks,
                        CanRetry: false,
                        RetryReason: "Inspect the workbook before retrying."),
                    "the formula write");
            }

            context.Checks.Add(formulaCheck);
            context.Changes.Add(new TaskChange("formula-write", $"{operation.WorksheetName}!{operation.Cells[0].Address}",
                $"Wrote {operation.Cells.Count} caller-supplied formulas."));

            return new MutationStep.SaveAndVerify(
                verification => VerifySavedFormulasBody(verification, operation.WorksheetName, formulaCells),
                $"Wrote {operation.Cells.Count} formulas, saved, and verified them after reopening.",
                "Excel saved the workbook, but reopen verification did not confirm every requested formula.");
        });
    }

    private static bool ApplyWorksheetFormulas(
        ExcelSession session,
        IReadOnlyList<FormulaWriteCell> cells,
        Action markMutationAttempted,
        out TaskCheck check)
    {
        using var references = new ComReferenceScope();
        var sheets = references.Add(Get(session.TargetWorkbook, "Worksheets"));
        var sheet = references.Add(Item(sheets, cells[0].WorksheetName));

        foreach (var cell in cells)
        {
            var target = references.Add(Get(sheet, "Range", cell.Address));
            markMutationAttempted();
            Set(target, "Formula", cell.Formula);
        }

        ReadFormulaWriteBox(references, sheet, cells, out var box, out var origin);
        foreach (var cell in cells)
        {
            var actual = box[cell.Row - origin.Row, cell.Column - origin.Column];
            if (!FormulaEquivalent(actual, cell.Formula))
            {
                check = new TaskCheck("formula-write", false,
                    $"Excel stored a different formula in {cell.Address}; no further cells were confirmed.");
                return false;
            }
        }

        check = new TaskCheck("formula-write", true,
            $"Wrote {cells.Count} formula(s) and read every one back from Excel.");
        return true;
    }

    private static (bool Verified, TaskCheck Check) VerifySavedFormulasBody(
        ExcelSession session,
        string worksheetName,
        IReadOnlyList<FormulaWriteCell> cells)
    {
        using var references = new ComReferenceScope();
        var sheets = references.Add(Get(session.TargetWorkbook, "Worksheets"));
        var sheet = references.Add(Item(sheets, worksheetName));
        ReadFormulaWriteBox(references, sheet, cells, out var box, out var origin);
        foreach (var cell in cells)
        {
            var actual = box[cell.Row - origin.Row, cell.Column - origin.Column];
            if (!FormulaEquivalent(actual, cell.Formula))
            {
                return (false, new TaskCheck("reopen-verification", false,
                    $"A requested formula was not present after reopening at {cell.Address}."));
            }
        }

        return (true, new TaskCheck("reopen-verification", true,
            $"Saved workbook reopened with all {cells.Count} requested formulas."));
    }

    private static FormulaWriteCell[] CreateFormulaWriteCells(NormalizedWriteWorksheetFormulasOperation operation) =>
        operation.Cells.Select(cell =>
        {
            var bounds = WorkbookRuntimeHelpers.GetBounds(new FormulaRepairRange(cell.Address, cell.Address));
            return new FormulaWriteCell(operation.WorksheetName, cell.Address, bounds.StartRow, bounds.StartColumn, cell.Formula);
        }).ToArray();

    private static void ReadFormulaWriteBox(
        ComReferenceScope references,
        object sheet,
        IReadOnlyList<FormulaWriteCell> cells,
        out string?[,] box,
        out (int Row, int Column) origin)
    {
        var minRow = cells.Min(cell => cell.Row);
        var maxRow = cells.Max(cell => cell.Row);
        var minColumn = cells.Min(cell => cell.Column);
        var maxColumn = cells.Max(cell => cell.Column);
        origin = (minRow, minColumn);

        var address = $"{WorkbookRuntimeHelpers.ToA1Address(minRow, minColumn)}:{WorkbookRuntimeHelpers.ToA1Address(maxRow, maxColumn)}";
        var range = references.Add(Get(sheet, "Range", address));
        var value = Get(range, "Formula");
        var rows = maxRow - minRow + 1;
        var columns = maxColumn - minColumn + 1;
        box = new string?[rows, columns];
        for (var row = 0; row < rows; row++)
        {
            for (var column = 0; column < columns; column++)
            {
                box[row, column] = FormulaCellOf(value, row, column);
            }
        }
    }

    private static string? FormulaCellOf(object? value, int row, int column) => value is Array values
        ? values.GetValue(row + values.GetLowerBound(0), column + values.GetLowerBound(1)) as string
        : value as string;

    private static bool FormulaEquivalent(string? actual, string expected) =>
        actual is not null && string.Equals(actual.Trim(), expected.Trim(), StringComparison.OrdinalIgnoreCase);

    private sealed record FormulaWriteCell(string WorksheetName, string Address, int Row, int Column, string Formula);
}
