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

        try
        {
            if (apply) pendingVerification = PendingVerification.Begin(observer);

            observer.OnPhase("session-open");
            session = ExcelSession.Open(plan.Request, observer, readOnlyTarget: !apply);

            observer.OnPhase("find-preflight");
            var preflight = PreflightWorksheetExists(session, operation.WorksheetName);
            checks.AddRange(preflight.Checks);
            if (!preflight.IsFeasible)
            {
                return ExcelSession.CloseAndProve(ref session, "find preflight", checks)
                    ?? new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected, "The requested worksheet was not found.", Checks: checks);
            }

            observer.OnPhase("find");
            if (!ScanForMatches(session, operation, out var searched, out var matches, out var scanCheck))
            {
                checks.Add(scanCheck);
                return ExcelSession.CloseAndProve(ref session, "the find", checks)
                    ?? new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected, "The requested range could not be searched.", Checks: checks);
            }

            checks.Add(scanCheck);
            var editable = matches.Where(match => !match.IsFormula).ToArray();
            if (matches.Count > editable.Length)
            {
                checks.Add(new TaskCheck("formula-cells-untouched", true,
                    $"{matches.Count - editable.Length} matching cell(s) get their text from a formula; they are listed and never rewritten."));
            }

            var receipt = MatchReceipt(operation, searched, matches);
            if (!apply)
            {
                changes.Add(new TaskChange("find", $"{operation.WorksheetName}!{searched.Address}",
                    $"Planned replacement of {editable.Length} constant cell(s)."));
                return ExcelSession.CloseAndProve(ref session, "the find", checks, changes)
                    ?? new WorkbookExecutionOutcome(ExcelTaskStatus.Planned,
                        $"Found {matches.Count} matching cell(s) in {operation.WorksheetName}!{searched.Address}; {editable.Length} would be rewritten. Nothing was changed.",
                        changes, checks, Range: receipt);
            }

            // Every replacement is computed and checked before any of them is written, so a request
            // that would produce something unwritable changes nothing at all rather than half a
            // sheet. A partial replacement can turn a label into a formula - "x=1" losing its "x"
            // leaves "=1" - and that is the one thing this server never writes.
            if (!TryComposeReplacements(operation, editable, out var replacements, out var composeCheck))
            {
                checks.Add(composeCheck);
                return ExcelSession.CloseAndProve(ref session, "the find", checks)
                    ?? new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected,
                        "The requested replacement would write formula text; nothing was changed.", Checks: checks);
            }

            if (replacements.Count == 0)
            {
                return ExcelSession.CloseAndProve(ref session, "the find", checks)
                    ?? new WorkbookExecutionOutcome(ExcelTaskStatus.Completed,
                        matches.Count == 0
                            ? $"No cell in {operation.WorksheetName}!{searched.Address} matched; nothing was changed."
                            : $"All {matches.Count} matching cell(s) get their text from formulas; nothing was changed.",
                        changes, checks, Range: receipt);
            }

            observer.OnPhase("replace");
            if (!ApplyReplacements(session, operation, replacements, () => mutationAttempted = true, out var replaceCheck))
            {
                checks.Add(replaceCheck);
                return ExcelSession.CloseAndProve(ref session, "the replace", checks, changes)
                    ?? new WorkbookExecutionOutcome(ExcelTaskStatus.Unknown,
                        "Excel did not store the replacement text as written.", changes, checks,
                        CanRetry: false, RetryReason: "Inspect the workbook before retrying.");
            }

            checks.Add(replaceCheck);
            changes.Add(new TaskChange("find-replace", $"{operation.WorksheetName}!{replacements[0].Address}",
                $"Replaced text in {replacements.Count} constant cell(s)."));

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
            var cleanupFailure = ExcelSession.CloseAndProve(ref session, "the find/replace save", checks, changes);
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
            if (!VerifySavedReplacements(verificationPath, operation, replacements, pendingVerification!, observer, out var reopenCheck))
            {
                checks.Add(reopenCheck);
                AddStagingCleanupCheck(stagingPath, checks);
                return new WorkbookExecutionOutcome(ExcelTaskStatus.Unknown,
                    "Excel saved the workbook, but reopen verification did not confirm every replacement.",
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
                $"Replaced text in {replacements.Count} of {matches.Count} matching cell(s), saved, and verified them after reopening.",
                changes, checks,
                Range: receipt with { Cells = [.. replacements.Select(item => new WorksheetCell(item.Address, item.Replacement))] });
        }
        catch (Exception exception) when (ComAccess.IsComFailure(exception))
        {
            checks.Add(new TaskCheck("find-replace", false, DescribeFailure("find-replace", exception)));
            var cleanupFailure = ExcelSession.CloseAndProve(ref session, "the failed find/replace", checks, changes);
            if (cleanupFailure is not null) return cleanupFailure;
            return new WorkbookExecutionOutcome(
                mutationAttempted ? ExcelTaskStatus.Unknown : ExcelTaskStatus.Rejected,
                mutationAttempted
                    ? "The find/replace failed after changes were attempted."
                    : "The find/replace was rejected before any change was attempted.",
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

    private static bool VerifySavedReplacements(
        string path,
        NormalizedFindReplaceOperation operation,
        IReadOnlyList<PlannedReplacement> replacements,
        PendingVerification verification,
        IExcelWorkbookRuntimeObserver observer,
        out TaskCheck check) =>
        verification.Verify(path, observer, session =>
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
        }, out check);

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
