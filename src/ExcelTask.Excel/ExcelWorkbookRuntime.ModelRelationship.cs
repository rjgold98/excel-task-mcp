using System.Globalization;
using ExcelTask.Core;

namespace ExcelTask.Excel;

public sealed partial class ExcelWorkbookRuntime
{
    /// <summary>
    /// Creates or deletes one Data Model relationship, and proves it by reading it back.
    ///
    /// This was deliberately absent while measures shipped, on the grounds that a wrong relationship
    /// silently changes every number the model produces and the operation that adds one should be
    /// able to show what it would join first. That is the condition this meets rather than waives:
    /// there is no expression to fingerprint, so the Plan is the precondition. It confirms both
    /// tables and both columns exist by name and lists every relationship already joining those two
    /// tables, so a caller sees the join - and any join it would sit beside - before making it.
    ///
    /// Direction is the thing callers get wrong, so it is named rather than inferred: From is the
    /// many side, To is the one side, and <c>ModelRelationships.Add</c> takes them in that order as
    /// foreign key then primary key. Excel refuses a relationship whose one side is not unique, and
    /// that refusal is reported as what it is rather than as a fault.
    /// </summary>
    private static WorkbookExecutionOutcome ExecuteModelRelationshipCore(ExcelTaskPlan plan, IExcelWorkbookRuntimeObserver observer)
    {
        var operation = plan.Request.Operation.ManageModelRelationship!;
        var target = $"{operation.FromTable}[{operation.FromColumn}] -> {operation.ToTable}[{operation.ToColumn}]";

        return ExecuteMutation(plan, observer, "model-relationship", "The relationship change", context =>
        {
            context.OnPhase("model-preflight");
            foreach (var (table, column) in new[]
            {
                (operation.FromTable, operation.FromColumn),
                (operation.ToTable, operation.ToColumn)
            })
            {
                if (TryFindModelColumn(context.Session, table, column, out var missing)) continue;

                context.Checks.Add(new TaskCheck("model-relationship", false, missing));
                return new MutationStep.Finish(
                    new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected,
                        "The Data Model column named for the join was not found.",
                        context.Changes, context.Checks, CanRetry: true,
                        RetryReason: "Run AuditWorkbookFlows to list the model tables, or load a query into the model with ManageQuery first."),
                    "model column lookup");
            }

            var existing = ReadRelationships(context.Session);
            var match = existing.FirstOrDefault(relationship => relationship.Matches(operation));
            var between = existing
                .Where(relationship => relationship.JoinsTheSameTables(operation))
                .Select(relationship => relationship.Describe())
                .ToList();

            // Emitted BEFORE the rejections below, not after. It used to sit under them, so a
            // rejected Delete told the caller to "run Plan to list the relationships joining those
            // two tables" - and Plan took the same early return and produced the same bare
            // rejection with no list. The advice named a remedy that could not be reached.
            //
            // A second relationship between the same pair is legal and Excel stores it inactive,
            // which is exactly the state that makes a model produce numbers nobody can account for,
            // so what is already there is named here rather than discovered later.
            context.Checks.Add(new TaskCheck("current-relationships", true, between.Count == 0
                ? $"Nothing currently joins {operation.FromTable} and {operation.ToTable}, in either direction."
                : $"Already joining those tables: {string.Join("; ", between)}."));

            if (operation.Action == QueryAction.Create && match is not null)
            {
                context.Checks.Add(new TaskCheck("model-relationship", false, $"That relationship already exists: {match.Describe()}."));
                return new MutationStep.Finish(
                    new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected,
                        "That relationship already exists; creating one never replaces it.",
                        context.Changes, context.Checks, CanRetry: false,
                        RetryReason: "Delete it first if the join is meant to change."),
                    "relationship creation");
            }

            if (operation.Action == QueryAction.Delete && match is null)
            {
                context.Checks.Add(new TaskCheck("model-relationship", false, $"No relationship {target} exists in the model."));
                return new MutationStep.Finish(
                    new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected,
                        "The named relationship was not found in the Data Model.",
                        context.Changes, context.Checks, CanRetry: true,
                        RetryReason: "The current-relationships check lists what does join those tables; the many side is named first."),
                    "relationship lookup");
            }

            if (!context.Apply)
            {
                var planned = operation.Action == QueryAction.Create
                    ? $"create the relationship {target}, with {operation.ToTable} as the one side"
                    : $"delete the relationship {target}";
                context.Changes.Add(new TaskChange("model-relationship", target, $"Planned to {planned}."));
                return new MutationStep.Finish(
                    new WorkbookExecutionOutcome(ExcelTaskStatus.Planned,
                        $"Applying would {planned}. Nothing was changed.",
                        context.Changes, context.Checks),
                    "relationship planning");
            }

            context.OnPhase("model-relationship");
            context.MarkMutationAttempted();
            try
            {
                ApplyRelationship(context.Session, operation);
            }
            catch (Exception exception) when (ComAccess.IsComFailure(exception))
            {
                // ModelRelationships.Add is one atomic call, so a refusal means nothing was written -
                // and the model can be re-read to prove it. Without this the commonest mistake there
                // is, reversing the many and one sides so the one side is not unique, came back as
                // Unknown with CanRetry false: a receipt telling a finance user their Data Model
                // might be in an indeterminate state, about a workbook Excel never touched.
                if (ReadRelationships(context.Session).Any(relationship => relationship.Matches(operation))) throw;

                context.Checks.Add(new TaskCheck("model-relationship", false,
                    $"Excel refused the relationship and the model is unchanged. The usual cause is that {operation.ToTable}[{operation.ToColumn}] is not unique, so it cannot be the one side."));
                return new MutationStep.Finish(
                    new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected,
                        "Excel refused the relationship; nothing in the Data Model was changed.",
                        context.Changes, context.Checks, CanRetry: true,
                        RetryReason: "The one side must hold each key once. If the sides are the wrong way round, swap them and submit again."),
                    "the relationship change");
            }

            var mismatch = FirstRelationshipMismatch(operation, ReadRelationships(context.Session));
            if (mismatch is not null)
            {
                context.Checks.Add(new TaskCheck("model-relationship", false, mismatch));
                return new MutationStep.Finish(
                    new WorkbookExecutionOutcome(ExcelTaskStatus.Unknown,
                        "Excel did not store the relationship change as requested; nothing was saved.",
                        context.Changes, context.Checks,
                        CanRetry: false, RetryReason: "Inspect the workbook's Data Model before retrying."),
                    "the relationship change");
            }

            var done = operation.Action == QueryAction.Create
                ? $"Created the relationship {target}."
                : $"Deleted the relationship {target}.";
            context.Checks.Add(new TaskCheck("model-relationship", true, done));
            context.Changes.Add(new TaskChange("model-relationship", target, done));

            return new MutationStep.SaveAndVerify(
                verification => FirstRelationshipMismatch(operation, ReadRelationships(verification)) is { } detail
                    ? (false, new TaskCheck("reopen-verification", false, $"After reopening the saved workbook: {detail}"))
                    : (true, new TaskCheck("reopen-verification", true, "The saved workbook reopened with the relationship as requested.")),
                $"{done} Saved and confirmed it after reopening.",
                "Excel saved the workbook, but reopen verification did not confirm the relationship.");
        });
    }

    private sealed record RelationshipSnapshot(string FromTable, string FromColumn, string ToTable, string ToColumn, bool Active)
    {
        /// <summary>The one relationship the caller named, in the direction they named it.</summary>
        public bool Matches(NormalizedManageModelRelationshipOperation operation) =>
            Same(FromTable, operation.FromTable) && Same(FromColumn, operation.FromColumn) &&
            Same(ToTable, operation.ToTable) && Same(ToColumn, operation.ToColumn);

        /// <summary>
        /// Any relationship between the same two tables, in EITHER direction. Deliberately
        /// unordered while <see cref="Matches"/> stays ordered: the caller who has the many and one
        /// sides the wrong way round is exactly the caller who most needs to be shown the join that
        /// already exists, and a direction-sensitive test told them "nothing currently joins these
        /// tables" - a false statement about the model, in a passed check.
        /// </summary>
        public bool JoinsTheSameTables(NormalizedManageModelRelationshipOperation operation) =>
            (Same(FromTable, operation.FromTable) && Same(ToTable, operation.ToTable)) ||
            (Same(FromTable, operation.ToTable) && Same(ToTable, operation.FromTable));

        public string Describe() =>
            $"{FromTable}[{FromColumn}] -> {ToTable}[{ToColumn}]{(Active ? string.Empty : " (inactive)")}";

        private static bool Same(string left, string right) => string.Equals(left, right, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// How many column names a not-found message lists before it reports a count instead. Chosen
    /// against the 128-character cut the worker applies to a check detail, so the message arrives
    /// whole rather than sliced mid-name.
    /// </summary>
    private const int MaxNamedColumns = 4;

    private static bool TryFindModelColumn(ExcelSession session, string tableName, string columnName, out string detail)
    {
        if (!TryFindModelTable(session, tableName, out detail)) return false;

        using var references = new ComReferenceScope();
        var model = references.Add(Get(session.TargetWorkbook, "Model"));
        var tables = references.Add(Get(model, "ModelTables"));
        var table = references.Add(Item(tables, tableName));
        var columns = references.Add(Get(table, "ModelTableColumns"));
        var count = Convert.ToInt32(Get(columns, "Count"), CultureInfo.InvariantCulture);
        var names = new List<string>();
        for (var index = 1; index <= count; index++)
        {
            var column = references.Add(Item(columns, index));
            var name = GetOrNull(column, "Name") as string ?? string.Empty;
            if (string.Equals(name, columnName, StringComparison.OrdinalIgnoreCase))
            {
                detail = $"{tableName}[{columnName}] is present.";
                return true;
            }

            names.Add(name);
        }

        // The columns are named, because the caller cannot guess them and the audit lists tables
        // rather than their columns. Column names inside a model are structure, not workbook data.
        //
        // Bounded here rather than at the worker seam, which hard-cuts a check detail at 128
        // characters with no ellipsis: a 22-column table produced four names and no sign that
        // anything was missing, reading as a complete list. Saying how many there are, and that
        // this is the first few, is a true short answer instead of a false complete one.
        detail = names.Count <= MaxNamedColumns
            ? $"The model table {tableName} has no column named {columnName}. It has: {string.Join(", ", names)}."
            : $"The model table {tableName} has no column named {columnName}. It has {names.Count} columns, starting {string.Join(", ", names.Take(MaxNamedColumns))} - run AuditWorkbookFlows for the rest.";
        return false;
    }

    private static List<RelationshipSnapshot> ReadRelationships(ExcelSession session)
    {
        var found = new List<RelationshipSnapshot>();
        using var references = new ComReferenceScope();
        var model = GetOrNull(session.TargetWorkbook, "Model");
        if (model is null) return found;

        references.Add(model);
        var relationships = references.Add(Get(model, "ModelRelationships"));
        var count = Convert.ToInt32(Get(relationships, "Count"), CultureInfo.InvariantCulture);
        for (var index = 1; index <= count; index++)
        {
            found.Add(SnapshotOf(references, references.Add(Item(relationships, index))));
        }

        return found;
    }

    /// <summary>
    /// One relationship as four names and whether it is in effect.
    ///
    /// Active is read because presence does not prove a join was made. Excel accepts a second
    /// relationship between an already-joined pair of tables and stores it inactive; without this,
    /// creating one and reading it back agrees, and the receipt says Completed for a relationship
    /// no pivot uses.
    /// </summary>
    private static RelationshipSnapshot SnapshotOf(ComReferenceScope references, object relationship)
    {
        var foreignColumn = references.Add(Get(relationship, "ForeignKeyColumn"));
        var primaryColumn = references.Add(Get(relationship, "PrimaryKeyColumn"));
        return new RelationshipSnapshot(
            NameOfTable(references, relationship, "ForeignKeyTable"),
            GetOrNull(foreignColumn, "Name") as string ?? string.Empty,
            NameOfTable(references, relationship, "PrimaryKeyTable"),
            GetOrNull(primaryColumn, "Name") as string ?? string.Empty,
            GetOrNull(relationship, "Active") is not bool active || active);
    }

    private static string NameOfTable(ComReferenceScope references, object relationship, string member) =>
        GetOrNull(references.Add(Get(relationship, member)), "Name") as string ?? string.Empty;

    private static void ApplyRelationship(ExcelSession session, NormalizedManageModelRelationshipOperation operation)
    {
        using var references = new ComReferenceScope();
        var model = references.Add(Get(session.TargetWorkbook, "Model"));
        var relationships = references.Add(Get(model, "ModelRelationships"));

        if (operation.Action == QueryAction.Delete)
        {
            var count = Convert.ToInt32(Get(relationships, "Count"), CultureInfo.InvariantCulture);
            for (var index = 1; index <= count; index++)
            {
                var relationship = references.Add(Item(relationships, index));
                if (!SnapshotOf(references, relationship).Matches(operation)) continue;

                // Returning immediately rather than continuing the walk: the collection reindexes
                // on Delete, so anything after this point would be reading a moved target.
                Invoke(relationship, "Delete");
                return;
            }

            return;
        }

        // Foreign key first, primary key second - the many side then the one side. Reversing them
        // is the mistake the payload's field names exist to prevent.
        references.Add(Invoke(relationships, "Add",
            ModelColumn(references, model, operation.FromTable, operation.FromColumn),
            ModelColumn(references, model, operation.ToTable, operation.ToColumn))!);
    }

    private static object ModelColumn(ComReferenceScope references, object model, string tableName, string columnName)
    {
        var tables = references.Add(Get(model, "ModelTables"));
        var table = references.Add(Item(tables, tableName));
        var columns = references.Add(Get(table, "ModelTableColumns"));
        return references.Add(Item(columns, columnName));
    }

    private static string? FirstRelationshipMismatch(
        NormalizedManageModelRelationshipOperation operation,
        List<RelationshipSnapshot> stored)
    {
        var match = stored.FirstOrDefault(relationship => relationship.Matches(operation));
        if (operation.Action == QueryAction.Delete)
        {
            return match is null ? null : "the relationship is still present after being deleted.";
        }

        if (match is null) return "no relationship joining those columns was present after the change.";

        // Presence is not proof the join is in effect. Excel accepts a second relationship between
        // an already-joined pair and stores it inactive, so a create that read back as present
        // reported Completed while every pivot went on aggregating through the older join.
        return match.Active
            ? null
            : $"Excel stored the relationship inactive, because {operation.FromTable} and {operation.ToTable} are already joined; no pivot will use it.";
    }
}
