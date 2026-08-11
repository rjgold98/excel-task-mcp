using System.Reflection;
using System.Runtime.InteropServices;
using ExcelTask.Core;

namespace ExcelTask.Excel;

public sealed partial class ExcelWorkbookRuntime
{
    /// <summary>One cell the search matched, and whether a formula is what put the text there.</summary>
    private sealed record FoundCell(string Address, string Text, bool IsFormula);

    /// <summary>
    /// Finds the cells whose text matches, and on Apply rewrites the constants among them.
    ///
    /// Excel's own Find and Replace are deliberately not used, for three separate reasons. Find
    /// treats <c>*</c>, <c>?</c> and <c>~</c> as wildcards, so a search for "Q1*" would quietly
    /// match text nobody asked about. Its omitted arguments inherit from the last search performed
    /// anywhere in the application - including one a person ran by hand - so the same request can
    /// mean different things on different machines. And Replace reports how many cells it changed
    /// and never which, so a receipt built on it could not name what moved.
    ///
    /// Reading the range once and matching in C# costs two COM calls regardless of how many cells
    /// match, against three per match for a Find/FindNext walk, and it makes the matching rule
    /// something this repository states and tests rather than something Excel decides.
    /// </summary>
    private static WorkbookExecutionOutcome ExecuteFindReplaceCore(ExcelTaskPlan plan, IExcelWorkbookRuntimeObserver observer)
    {
        var operation = plan.Request.Operation.FindReplace!;
        return ExecuteMutation(plan, observer, "find-replace", "The find/replace", context =>
        {
            context.OnPhase("find-preflight");
            var preflight = PreflightWorksheetExists(context.Session, operation.WorksheetName);
            context.Checks.AddRange(preflight.Checks);
            if (!preflight.IsFeasible)
            {
                return new MutationStep.Finish(
                    new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected, "The requested worksheet was not found.", Checks: context.Checks),
                    "find preflight");
            }

            context.OnPhase("find");
            if (!ScanForMatches(context.Session, operation, out var searched, out var matches, out var scanCheck))
            {
                context.Checks.Add(scanCheck);
                return new MutationStep.Finish(
                    new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected, "The requested range could not be searched.", Checks: context.Checks),
                    "the find");
            }

            context.Checks.Add(scanCheck);
            var editable = matches.Where(match => !match.IsFormula).ToArray();
            if (matches.Count > editable.Length)
            {
                context.Checks.Add(new TaskCheck("formula-cells-untouched", true,
                    $"{matches.Count - editable.Length} matching cell(s) get their text from a formula; they are listed and never rewritten."));
            }

            var receipt = MatchReceipt(operation, searched, matches);
            if (!context.Apply)
            {
                context.Changes.Add(new TaskChange("find", $"{operation.WorksheetName}!{searched.Address}",
                    $"Planned replacement of {editable.Length} constant cell(s)."));
                return new MutationStep.Finish(
                    new WorkbookExecutionOutcome(ExcelTaskStatus.Planned,
                        $"Found {matches.Count} matching cell(s) in {operation.WorksheetName}!{searched.Address}; {editable.Length} would be rewritten. Nothing was changed.",
                        context.Changes, context.Checks, Range: receipt),
                    "the find");
            }

            // Every replacement is computed and checked before any of them is written, so a request
            // that would produce something unwritable changes nothing at all rather than half a
            // sheet. A partial replacement can turn a label into a formula - "x=1" losing its "x"
            // leaves "=1" - and that is the one thing this server never writes.
            if (!TryComposeReplacements(operation, editable, out var replacements, out var composeCheck))
            {
                context.Checks.Add(composeCheck);
                return new MutationStep.Finish(
                    new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected,
                        "The requested replacement would write formula text; nothing was changed.", Checks: context.Checks),
                    "the find");
            }

            if (replacements.Count == 0)
            {
                return new MutationStep.Finish(
                    new WorkbookExecutionOutcome(ExcelTaskStatus.Completed,
                        matches.Count == 0
                            ? $"No cell in {operation.WorksheetName}!{searched.Address} matched; nothing was changed."
                            : $"All {matches.Count} matching cell(s) get their text from formulas; nothing was changed.",
                        context.Changes, context.Checks, Range: receipt),
                    "the find");
            }

            context.OnPhase("replace");
            if (!ApplyReplacements(context.Session, operation, replacements, context.MarkMutationAttempted, out var replaceCheck))
            {
                context.Checks.Add(replaceCheck);
                return new MutationStep.Finish(
                    new WorkbookExecutionOutcome(ExcelTaskStatus.Unknown,
                        "Excel did not store the replacement text as written.", context.Changes, context.Checks,
                        CanRetry: false, RetryReason: "Inspect the workbook before retrying."),
                    "the replace");
            }

            context.Checks.Add(replaceCheck);
            context.Changes.Add(new TaskChange("find-replace", $"{operation.WorksheetName}!{replacements[0].Address}",
                $"Replaced text in {replacements.Count} constant cell(s)."));

            return new MutationStep.SaveAndVerify(
                verification => VerifySavedReplacementsBody(verification, operation, replacements),
                $"Replaced text in {replacements.Count} of {matches.Count} matching cell(s), saved, and verified them after reopening.",
                "Excel saved the workbook, but reopen verification did not confirm every replacement.",
                Range: receipt with { Cells = [.. replacements.Select(item => new WorksheetCell(item.Address, item.Replacement))] });
        });
    }

    /// <summary>The area a search covered: its A1 span and how many cells that is.</summary>
    private sealed record SearchedArea(string Address, int CellCount);

    /// <summary>
    /// Reads the search area once and returns the cells whose text matches. Two array reads - values
    /// and formulas - answer both questions the receipt needs, at a fixed cost no matter how dense
    /// the sheet is or how many cells match.
    /// </summary>
    private static bool ScanForMatches(
        ExcelSession session,
        NormalizedFindReplaceOperation operation,
        out SearchedArea searched,
        out IReadOnlyList<FoundCell> matches,
        out TaskCheck check)
    {
        searched = new SearchedArea(string.Empty, 0);
        matches = [];
        using var references = new ComReferenceScope();
        var sheets = references.Add(Get(session.TargetWorkbook, "Worksheets"));
        var sheet = references.Add(Item(sheets, operation.WorksheetName));

        FormulaRangeBounds bounds;
        object area;
        if (operation.Range is not null)
        {
            bounds = WorkbookRuntimeHelpers.GetBounds(operation.Range);
            area = references.Add(Get(sheet, "Range", operation.Range.ToString()));
        }
        else
        {
            var used = references.Add(Get(sheet, "UsedRange"));
            bounds = ReadBounds(used, references);
            // An omitted range means "the sheet", and the sheet can be far larger than the bound the
            // engine enforces on an explicit one. Refusing here, by name, is the answer that lets the
            // caller narrow the request; silently searching part of it would not be.
            var cells = (long)bounds.RowCount * bounds.ColumnCount;
            if (cells > ExcelTaskEngine.MaxFormulaRepairCells)
            {
                check = new TaskCheck("find", false,
                    $"The worksheet's used range spans {cells:N0} cells, more than the {ExcelTaskEngine.MaxFormulaRepairCells:N0} a search covers. Name a range to search.");
                return false;
            }

            area = used;
        }

        searched = new SearchedArea(
            $"{WorkbookRuntimeHelpers.ToA1Address(bounds.StartRow, bounds.StartColumn)}:{WorkbookRuntimeHelpers.ToA1Address(bounds.EndRow, bounds.EndColumn)}",
            bounds.RowCount * bounds.ColumnCount);

        var values = Get(area, "Value2");
        var formulas = Get(area, "Formula");
        var found = new List<FoundCell>();
        for (var row = 0; row < bounds.RowCount; row++)
        {
            for (var column = 0; column < bounds.ColumnCount; column++)
            {
                var text = Render(CellOf(values, row, column));
                if (text.Length == 0 || !MatchesFind(text, operation)) continue;

                found.Add(new FoundCell(
                    WorkbookRuntimeHelpers.ToA1Address(bounds.StartRow + row, bounds.StartColumn + column),
                    text,
                    CellOf(formulas, row, column) is string formula && formula.StartsWith('=')));
            }
        }

        matches = found;
        check = new TaskCheck("find", true,
            $"Searched {searched.CellCount:N0} cell(s) in {operation.WorksheetName}!{searched.Address} and matched {found.Count}.");
        return true;
    }

    private static object? CellOf(object value, int row, int column) => value is Array values
        ? values.GetValue(row + values.GetLowerBound(0), column + values.GetLowerBound(1))
        : value;

    private static FormulaRangeBounds ReadBounds(object area, ComReferenceScope references)
    {
        var rows = references.Add(Get(area, "Rows"));
        var columns = references.Add(Get(area, "Columns"));
        var startRow = Convert.ToInt32(Get(area, "Row"), System.Globalization.CultureInfo.InvariantCulture);
        var startColumn = Convert.ToInt32(Get(area, "Column"), System.Globalization.CultureInfo.InvariantCulture);
        var rowCount = Convert.ToInt32(Get(rows, "Count"), System.Globalization.CultureInfo.InvariantCulture);
        var columnCount = Convert.ToInt32(Get(columns, "Count"), System.Globalization.CultureInfo.InvariantCulture);
        return new FormulaRangeBounds(startRow, startColumn, startRow + rowCount - 1, startColumn + columnCount - 1);
    }

    private static bool MatchesFind(string text, NormalizedFindReplaceOperation operation)
    {
        var comparison = operation.MatchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        return operation.WholeCell
            ? string.Equals(text, operation.Find, comparison)
            : text.Contains(operation.Find, comparison);
    }

    private sealed record PlannedReplacement(string Address, string Replacement);

    /// <summary>
    /// Works out what each matching cell would become, and refuses the whole request if any result
    /// would be formula text. Composing before writing is what makes that refusal free of cost:
    /// nothing has been written when it happens.
    /// </summary>
    private static bool TryComposeReplacements(
        NormalizedFindReplaceOperation operation,
        IReadOnlyList<FoundCell> editable,
        out IReadOnlyList<PlannedReplacement> replacements,
        out TaskCheck check)
    {
        var comparison = operation.MatchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        var planned = new List<PlannedReplacement>(editable.Count);
        foreach (var match in editable)
        {
            var replaced = operation.WholeCell
                ? operation.ReplaceWith!
                : match.Text.Replace(operation.Find, operation.ReplaceWith!, comparison);
            if (replaced.StartsWith('='))
            {
                replacements = [];
                check = new TaskCheck("formula-text-refused", false,
                    $"Replacing in {match.Address} would leave text starting with '=', which Excel stores as a formula. No cell was changed.");
                return false;
            }

            planned.Add(new PlannedReplacement(match.Address, replaced));
        }

        replacements = planned;
        check = new TaskCheck("replacement-plan", true, $"Composed {planned.Count} replacement(s), none of which is formula text.");
        return true;
    }

    /// <summary>
    /// Writes each replacement and reads every one back, for the same reason the constant write
    /// does: Excel coerces on assignment, and a receipt that reported the request rather than the
    /// result would be reporting an intention.
    /// </summary>
    private static bool ApplyReplacements(
        ExcelSession session,
        NormalizedFindReplaceOperation operation,
        IReadOnlyList<PlannedReplacement> replacements,
        Action markMutationAttempted,
        out TaskCheck check)
    {
        using var references = new ComReferenceScope();
        var sheets = references.Add(Get(session.TargetWorkbook, "Worksheets"));
        var sheet = references.Add(Item(sheets, operation.WorksheetName));

        foreach (var item in replacements)
        {
            var target = references.Add(Get(sheet, "Range", item.Address));
            markMutationAttempted();
            Set(target, "Value2", ToCellValue(item.Replacement));
        }

        foreach (var item in replacements)
        {
            var target = references.Add(Get(sheet, "Range", item.Address));
            if (!string.Equals(Render(GetOrNull(target, "Value2")), item.Replacement, StringComparison.Ordinal))
            {
                check = new TaskCheck("find-replace", false,
                    $"Excel stored something other than the replacement text in {item.Address}; no further cells were confirmed.");
                return false;
            }
        }

        check = new TaskCheck("find-replace", true, $"Replaced {replacements.Count} cell(s) and read every one back unchanged.");
        return true;
    }

    private static (bool Verified, TaskCheck Check) VerifySavedReplacementsBody(
        ExcelSession session,
        NormalizedFindReplaceOperation operation,
        IReadOnlyList<PlannedReplacement> replacements)
    {
        using var references = new ComReferenceScope();
        var sheets = references.Add(Get(session.TargetWorkbook, "Worksheets"));
        var sheet = references.Add(Item(sheets, operation.WorksheetName));
        foreach (var item in replacements)
        {
            var target = references.Add(Get(sheet, "Range", item.Address));
            if (!string.Equals(Render(GetOrNull(target, "Value2")), item.Replacement, StringComparison.Ordinal))
            {
                return (false, new TaskCheck("reopen-verification", false,
                    "A replaced value was not present after reopening the saved workbook."));
            }
        }

        return (true, new TaskCheck("reopen-verification", true,
            $"Saved workbook reopened with all {replacements.Count} replacement(s) in place."));
    }

    private static WorksheetRangeReceipt MatchReceipt(
        NormalizedFindReplaceOperation operation,
        SearchedArea searched,
        IReadOnlyList<FoundCell> matches) =>
        new(operation.WorksheetName,
            searched.Address,
            Formulas: false,
            searched.CellCount,
            matches.Count,
            [.. matches.Select(match => new WorksheetCell(match.Address, match.Text))],
            Truncated: false);
}
