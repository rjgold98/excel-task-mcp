using ExcelTask.Core;

namespace ExcelTask.Excel;

public sealed partial class ExcelWorkbookRuntime
{
    // ListObjects.Add(SourceType, Source, LinkSource, XlYesNoGuess). 1 is xlSrcRange and 1 is
    // xlYes - the header row is in the range, which is why the schema says to include it.
    private const int XlSrcRange = 1;
    private const int XlYes = 1;

    /// <summary>
    /// Creates or changes one table, and proves the change by reading the table back.
    ///
    /// Every action here except ConvertToRange is undone by another call, which is why this
    /// operation is willing to perform them at all. ConvertToRange is the exception and is treated
    /// as one: it keeps every cell exactly where it was and drops only the table over them, so the
    /// worst case is a caller who has to create the table again rather than one who has lost data.
    ///
    /// The read-back is not ceremony. Excel renames a table it is given a name it dislikes, refuses
    /// a style it does not have without saying so, and resizes to the nearest thing it can fit -
    /// and all three look like success from the caller's side of the assignment.
    /// </summary>
    private static WorkbookExecutionOutcome ExecuteManageTableCore(ExcelTaskPlan plan, IExcelWorkbookRuntimeObserver observer)
    {
        var operation = plan.Request.Operation.ManageTable!;
        var target = $"{operation.WorksheetName}!{operation.TableName}";

        return ExecuteMutation(plan, observer, "table", "The table change", context =>
        {
            context.OnPhase("table-preflight");
            var preflight = PreflightWorksheetExists(context.Session, operation.WorksheetName);
            context.Checks.AddRange(preflight.Checks);
            if (!preflight.IsFeasible)
            {
                return new MutationStep.Finish(
                    new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected, "The requested worksheet was not found.", Checks: context.Checks),
                    "table preflight");
            }

            var existing = ReadTable(context.Session, operation.WorksheetName, operation.TableName);

            // Create needs the name free; everything else needs it taken. Saying which, before
            // Excel starts changing anything, is the difference between a clean rejection and an
            // Unknown that leaves the caller wondering what happened to their sheet.
            if (operation.Action == TableAction.Create && existing is not null)
            {
                context.Checks.Add(new TaskCheck("table", false, $"A table named {operation.TableName} already exists on {operation.WorksheetName}."));
                return new MutationStep.Finish(
                    new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected,
                        "A table with that name already exists; creating one never replaces it.",
                        context.Changes, context.Checks, CanRetry: true,
                        RetryReason: "Choose another name, or use Resize or Restyle to change the existing table."),
                    "table creation");
            }

            if (operation.Action != TableAction.Create && existing is null)
            {
                context.Checks.Add(new TaskCheck("table", false, $"No table named {operation.TableName} exists on {operation.WorksheetName}."));
                return new MutationStep.Finish(
                    new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected,
                        "The named table was not found on that worksheet.",
                        context.Changes, context.Checks, CanRetry: true,
                        RetryReason: "Run AuditWorkbookFlows or ScanWorkbookStructure to list the tables that exist."),
                    "table lookup");
            }

            context.Checks.Add(new TaskCheck("current-table", true, existing is null
                ? $"No table named {operation.TableName} exists on {operation.WorksheetName} yet."
                : $"{operation.TableName} currently covers {existing.Range} with style {existing.Style ?? "none"}."));

            var described = Describe(operation);
            if (!context.Apply)
            {
                context.Changes.Add(new TaskChange("table", target, $"Planned {described}."));
                return new MutationStep.Finish(
                    new WorkbookExecutionOutcome(ExcelTaskStatus.Planned,
                        $"Applying would {described}. Nothing was changed.", context.Changes, context.Checks),
                    "table planning");
            }

            context.OnPhase("table");
            context.MarkMutationAttempted();
            ApplyTable(context.Session, operation);

            var expectedName = operation.Action == TableAction.Rename ? operation.NewName! : operation.TableName;
            var stored = ReadTable(context.Session, operation.WorksheetName, expectedName);
            var mismatch = FirstTableMismatch(operation, expectedName, stored);
            if (mismatch is not null)
            {
                context.Checks.Add(new TaskCheck("table", false, mismatch));
                return new MutationStep.Finish(
                    new WorkbookExecutionOutcome(ExcelTaskStatus.Unknown,
                        "Excel did not store the table change as requested; nothing was saved.",
                        context.Changes, context.Checks,
                        CanRetry: false, RetryReason: "Inspect the worksheet before retrying."),
                    "the table change");
            }

            context.Checks.Add(new TaskCheck("table", true, $"Applied {described} and read the table back unchanged."));
            context.Changes.Add(new TaskChange("table", target, $"Did {described}."));

            return new MutationStep.SaveAndVerify(
                verification => FirstTableMismatch(operation, expectedName, ReadTable(verification, operation.WorksheetName, expectedName)) is { } detail
                    ? (false, new TaskCheck("reopen-verification", false, $"After reopening the saved workbook: {detail}"))
                    : (true, new TaskCheck("reopen-verification", true, "The saved workbook reopened with the table as requested.")),
                $"Did {described} on {operation.WorksheetName}, saved, and confirmed it after reopening.",
                "Excel saved the workbook, but reopen verification did not confirm the table.");
        });
    }

    private static string Describe(NormalizedManageTableOperation operation) => operation.Action switch
    {
        TableAction.Create => $"create the table {operation.TableName} over {operation.Range}",
        TableAction.Rename => $"rename the table {operation.TableName} to {operation.NewName}",
        TableAction.Restyle => $"restyle the table {operation.TableName} as {operation.TableStyle}",
        TableAction.Resize => $"resize the table {operation.TableName} to {operation.Range}",
        _ => $"convert the table {operation.TableName} back to plain cells, keeping every one of them"
    };

    /// <summary>One table as Excel currently holds it, or null when no table by that name is there.</summary>
    private sealed record TableSnapshot(string Name, string Range, string? Style);

    private static TableSnapshot? ReadTable(ExcelSession session, string worksheetName, string tableName)
    {
        using var references = new ComReferenceScope();
        var sheets = references.Add(Get(session.TargetWorkbook, "Worksheets"));
        var sheet = references.Add(Item(sheets, worksheetName));
        var tables = references.Add(Get(sheet, "ListObjects"));

        // Walked rather than fetched by name, because Item throws for a missing table and a throw
        // is not how "it is not there" should reach the caller.
        var count = Convert.ToInt32(Get(tables, "Count"), System.Globalization.CultureInfo.InvariantCulture);
        for (var index = 1; index <= count; index++)
        {
            var table = references.Add(Item(tables, index));
            var name = GetOrNull(table, "Name") as string;
            if (!string.Equals(name, tableName, StringComparison.OrdinalIgnoreCase)) continue;

            var range = references.Add(Get(table, "Range"));
            return new TableSnapshot(
                name!,
                GetOrNull(range, "Address") as string ?? "(unknown)",
                GetOrNull(table, "TableStyle") is { } style ? StyleName(references, style) : null);
        }

        return null;
    }

    /// <summary>
    /// TableStyle answers with a style object on a styled table and an empty string on an unstyled
    /// one, so the name has to be pulled out of whichever it is.
    /// </summary>
    private static string? StyleName(ComReferenceScope references, object style)
    {
        if (style is string text) return string.IsNullOrEmpty(text) ? null : text;
        references.Add(style);
        return GetOrNull(style, "Name") as string;
    }

    private static void ApplyTable(ExcelSession session, NormalizedManageTableOperation operation)
    {
        using var references = new ComReferenceScope();
        var sheets = references.Add(Get(session.TargetWorkbook, "Worksheets"));
        var sheet = references.Add(Item(sheets, operation.WorksheetName));
        var tables = references.Add(Get(sheet, "ListObjects"));

        if (operation.Action == TableAction.Create)
        {
            var source = references.Add(Get(sheet, "Range", operation.Range!.ToString()));
            var created = references.Add(Invoke(tables, "Add", XlSrcRange, source, Type.Missing, XlYes)!);
            Set(created, "Name", operation.TableName);
            if (!string.IsNullOrEmpty(operation.TableStyle)) SetTableStyle(created, operation.TableStyle);
            return;
        }

        var table = references.Add(Item(tables, operation.TableName));
        switch (operation.Action)
        {
            case TableAction.Rename:
                Set(table, "Name", operation.NewName!);
                break;

            case TableAction.Restyle:
                SetTableStyle(table, operation.TableStyle!);
                break;

            case TableAction.Resize:
                var resized = references.Add(Get(sheet, "Range", operation.Range!.ToString()));
                Invoke(table, "Resize", resized);
                break;

            default:
                // Unlist keeps every cell and its contents; only the table over them goes.
                Invoke(table, "Unlist");
                break;
        }
    }

    /// <summary>None removes the style rather than looking one up by that name.</summary>
    private static void SetTableStyle(object table, string style) =>
        Set(table, "TableStyle", string.Equals(style, "None", StringComparison.OrdinalIgnoreCase) ? string.Empty : style);

    private static string? FirstTableMismatch(NormalizedManageTableOperation operation, string expectedName, TableSnapshot? stored)
    {
        if (operation.Action == TableAction.ConvertToRange)
        {
            return stored is null ? null : $"the table {operation.TableName} is still a table after being converted.";
        }

        if (stored is null) return $"no table named {expectedName} was present after the change.";

        if (operation.Action is TableAction.Create or TableAction.Restyle && !string.IsNullOrEmpty(operation.TableStyle))
        {
            var wantsNone = string.Equals(operation.TableStyle, "None", StringComparison.OrdinalIgnoreCase);
            if (wantsNone && stored.Style is not null) return "the table still carries a style after being asked for none.";

            // A style Excel does not have is not an error to it - it keeps the one it had, which
            // reads as success unless the name is compared.
            if (!wantsNone && !string.Equals(stored.Style, operation.TableStyle, StringComparison.OrdinalIgnoreCase))
            {
                return stored.Style is null
                    ? $"Excel left the table unstyled rather than applying {operation.TableStyle}; that style is probably not in this workbook."
                    : $"Excel stored the style {stored.Style} rather than {operation.TableStyle}.";
            }
        }

        return null;
    }
}
