using System.Globalization;
using ExcelTask.Core;

namespace ExcelTask.Excel;

public sealed partial class ExcelWorkbookRuntime
{
    /// <summary>
    /// Creates, replaces, or deletes one Data Model measure, and proves it by reading it back.
    ///
    /// A measure is read by every pivot over the model, so a wrong one is wrong everywhere at once
    /// and quietly. It therefore carries the same precondition as a query or a macro edit: Plan
    /// reports the current DAX and its fingerprint, and Apply must send that fingerprint back.
    ///
    /// Unlike an M expression, the DAX is returned in the Plan receipt. DAX names tables and columns
    /// inside the model rather than servers and databases, so there is nothing here the audit would
    /// refuse to return - and a caller writing a replacement genuinely needs to see what they are
    /// replacing. The two operations differ because the risk differs, not by accident.
    ///
    /// Two things about the COM surface are worth knowing, because both look like the capability is
    /// missing rather than the call being wrong. ModelMeasures.Add requires a real format object -
    /// null produces "value does not fall within the expected range" - and the expression must carry
    /// no leading equals sign, though Excel's own editor shows one.
    /// </summary>
    private static WorkbookExecutionOutcome ExecuteModelMeasureCore(ExcelTaskPlan plan, IExcelWorkbookRuntimeObserver observer)
    {
        var operation = plan.Request.Operation.ManageModelMeasure!;
        var target = $"{operation.TableName}[{operation.MeasureName}]";

        return ExecuteMutation(plan, observer, "model-measure", "The measure change", context =>
        {
            context.OnPhase("model-preflight");
            if (!TryFindModelTable(context.Session, operation.TableName, out var tableMissingDetail))
            {
                context.Checks.Add(new TaskCheck("model-measure", false, tableMissingDetail));
                return new MutationStep.Finish(
                    new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected,
                        "The Data Model table was not found.",
                        context.Changes, context.Checks, CanRetry: true,
                        RetryReason: "Run AuditWorkbookFlows to list the model tables, or load a query into the model with ManageQuery first."),
                    "model table lookup");
            }

            var existing = ReadMeasure(context.Session, operation.MeasureName);

            if (operation.Action == QueryAction.Create && existing is not null)
            {
                context.Checks.Add(new TaskCheck("model-measure", false, $"A measure named {operation.MeasureName} already exists."));
                return new MutationStep.Finish(
                    new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected,
                        "A measure with that name already exists; creating one never replaces it.",
                        context.Changes, context.Checks, CanRetry: true,
                        RetryReason: "Choose another name, or use Replace with the fingerprint from a Plan."),
                    "measure creation");
            }

            if (operation.Action != QueryAction.Create && existing is null)
            {
                context.Checks.Add(new TaskCheck("model-measure", false, $"No measure named {operation.MeasureName} exists in the model."));
                return new MutationStep.Finish(
                    new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected,
                        "The named measure was not found in the Data Model.",
                        context.Changes, context.Checks, CanRetry: true,
                        RetryReason: "Run AuditWorkbookFlows to list the measures this model has."),
                    "measure lookup");
            }

            context.Checks.Add(new TaskCheck("current-measure", true, existing is null
                ? $"No measure named {operation.MeasureName} exists yet."
                : $"{operation.MeasureName} is currently {existing.Formula}, fingerprint {existing.Sha256}."));

            if (!context.Apply)
            {
                var planned = operation.Action switch
                {
                    QueryAction.Create => $"create the measure {target}",
                    QueryAction.Replace => $"replace the measure {target}",
                    _ => $"delete the measure {target}"
                };
                context.Changes.Add(new TaskChange("model-measure", target, $"Planned to {planned}."));
                return new MutationStep.Finish(
                    new WorkbookExecutionOutcome(ExcelTaskStatus.Planned,
                        existing is null
                            ? $"Applying would {planned}. Nothing was changed."
                            : $"Applying would {planned}. Send fingerprint {existing.Sha256} with the Apply. Nothing was changed.",
                        context.Changes, context.Checks),
                    "measure planning");
            }

            if (existing is not null && !string.Equals(existing.Sha256, operation.ExpectedFormulaSha256, StringComparison.OrdinalIgnoreCase))
            {
                context.Checks.Add(new TaskCheck("measure-fingerprint", false,
                    "The measure changed since the Plan that produced this fingerprint; nothing was written."));
                return new MutationStep.Finish(
                    new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected,
                        "The expected measure fingerprint does not match the measure in the model.",
                        context.Changes, context.Checks, CanRetry: true,
                        RetryReason: "Run Plan again to take a current fingerprint, then resubmit."),
                    "the measure fingerprint");
            }

            if (existing is not null)
            {
                context.Checks.Add(new TaskCheck("measure-fingerprint", true, "The measure matches the fingerprint the Plan reported."));
            }

            context.OnPhase("model-measure");
            context.MarkMutationAttempted();
            ApplyMeasure(context.Session, operation);

            var mismatch = FirstMeasureMismatch(operation, ReadMeasure(context.Session, operation.MeasureName));
            if (mismatch is not null)
            {
                context.Checks.Add(new TaskCheck("model-measure", false, mismatch));
                return new MutationStep.Finish(
                    new WorkbookExecutionOutcome(ExcelTaskStatus.Unknown,
                        "Excel did not store the measure change as requested; nothing was saved.",
                        context.Changes, context.Checks,
                        CanRetry: false, RetryReason: "Inspect the workbook's Data Model before retrying."),
                    "the measure change");
            }

            var done = operation.Action switch
            {
                QueryAction.Create => $"Created the measure {target}.",
                QueryAction.Replace => $"Replaced the measure {target}.",
                _ => $"Deleted the measure {target}."
            };
            context.Checks.Add(new TaskCheck("model-measure", true, done));
            context.Changes.Add(new TaskChange("model-measure", target, done));

            return new MutationStep.SaveAndVerify(
                verification => FirstMeasureMismatch(operation, ReadMeasure(verification, operation.MeasureName)) is { } detail
                    ? (false, new TaskCheck("reopen-verification", false, $"After reopening the saved workbook: {detail}"))
                    : (true, new TaskCheck("reopen-verification", true, "The saved workbook reopened with the measure as requested.")),
                $"{done} Saved and confirmed it after reopening.",
                "Excel saved the workbook, but reopen verification did not confirm the measure.");
        });
    }

    private sealed record MeasureSnapshot(string Name, string Formula, string Sha256);

    private static bool TryFindModelTable(ExcelSession session, string tableName, out string detail)
    {
        using var references = new ComReferenceScope();
        var model = GetOrNull(session.TargetWorkbook, "Model");
        if (model is null)
        {
            detail = "This workbook has no Data Model.";
            return false;
        }

        references.Add(model);
        var tables = references.Add(Get(model, "ModelTables"));
        var count = Convert.ToInt32(Get(tables, "Count"), CultureInfo.InvariantCulture);
        for (var index = 1; index <= count; index++)
        {
            var table = references.Add(Item(tables, index));
            if (string.Equals(GetOrNull(table, "Name") as string, tableName, StringComparison.OrdinalIgnoreCase))
            {
                detail = $"The model table {tableName} is present.";
                return true;
            }
        }

        detail = count == 0
            ? "This workbook's Data Model holds no tables; load a query into the model first."
            : $"The Data Model has {count} table(s), and none of them is named {tableName}.";
        return false;
    }

    private static MeasureSnapshot? ReadMeasure(ExcelSession session, string measureName)
    {
        using var references = new ComReferenceScope();
        var model = GetOrNull(session.TargetWorkbook, "Model");
        if (model is null) return null;

        references.Add(model);
        var measures = references.Add(Get(model, "ModelMeasures"));
        var count = Convert.ToInt32(Get(measures, "Count"), CultureInfo.InvariantCulture);
        for (var index = 1; index <= count; index++)
        {
            var measure = references.Add(Item(measures, index));
            var name = GetOrNull(measure, "Name") as string;
            if (!string.Equals(name, measureName, StringComparison.OrdinalIgnoreCase)) continue;

            var formula = (GetOrNull(measure, "Formula") as string ?? string.Empty).Trim();
            return new MeasureSnapshot(name!, formula, MacroProcedureText.ComputeSha256(formula));
        }

        return null;
    }

    private static void ApplyMeasure(ExcelSession session, NormalizedManageModelMeasureOperation operation)
    {
        using var references = new ComReferenceScope();
        var model = references.Add(Get(session.TargetWorkbook, "Model"));
        var measures = references.Add(Get(model, "ModelMeasures"));

        if (operation.Action == QueryAction.Delete)
        {
            var existing = references.Add(Item(measures, operation.MeasureName));
            Invoke(existing, "Delete");
            return;
        }

        if (operation.Action == QueryAction.Replace)
        {
            // Replace in place rather than delete-and-add: the measure keeps its identity, so
            // anything already pointing at it keeps working.
            var existing = references.Add(Item(measures, operation.MeasureName));
            Set(existing, "Formula", operation.Formula!);
            return;
        }

        var tables = references.Add(Get(model, "ModelTables"));
        var table = references.Add(Item(tables, operation.TableName));
        // The format argument is required. Passing null is what made this look impossible.
        var format = references.Add(Get(model, "ModelFormatGeneral"));
        references.Add(Invoke(measures, "Add", operation.MeasureName, table, operation.Formula!, format)!);
    }

    private static string? FirstMeasureMismatch(NormalizedManageModelMeasureOperation operation, MeasureSnapshot? stored)
    {
        if (operation.Action == QueryAction.Delete)
        {
            return stored is null ? null : $"the measure {operation.MeasureName} is still present after being deleted.";
        }

        if (stored is null) return $"no measure named {operation.MeasureName} was present after the change.";

        // Excel normalizes whitespace inside a DAX expression, so an exact comparison would fail a
        // change it actually made. Comparing without whitespace still catches a different formula.
        return string.Equals(Squeeze(stored.Formula), Squeeze(operation.Formula!), StringComparison.OrdinalIgnoreCase)
            ? null
            : $"Excel stored the expression {stored.Formula} rather than the one requested.";
    }

    private static string Squeeze(string value) => string.Concat(value.Where(character => !char.IsWhiteSpace(character)));
}
