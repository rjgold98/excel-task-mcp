using System.Globalization;
using ExcelTask.Core;

namespace ExcelTask.Excel;

public sealed partial class ExcelWorkbookRuntime
{
    /// <summary>
    /// Creates, replaces, or deletes one Power Query, and proves the change by reading the query
    /// back and fingerprinting it.
    ///
    /// A query is where a workbook's numbers come from, so this borrows the macro operation's whole
    /// posture: Plan changes nothing and reports a fingerprint, Apply must carry that fingerprint
    /// back, and a mismatch stops before anything is written. The precondition is the point - it is
    /// what makes "replace this query" mean the query the caller looked at rather than whatever is
    /// there now.
    ///
    /// The expression never enters a receipt. An M expression usually names a server and a database,
    /// which is precisely what AuditWorkbookFlows refuses to return, and a fingerprint carries the
    /// proof without carrying the secret. That does mean a caller replacing a query must read it in
    /// Excel first; that is the trade, and it is the same one the audit already makes.
    /// </summary>
    private static WorkbookExecutionOutcome ExecuteManageQueryCore(ExcelTaskPlan plan, IExcelWorkbookRuntimeObserver observer)
    {
        var operation = plan.Request.Operation.ManageQuery!;

        return ExecuteMutation(plan, observer, "query", "The query change", context =>
        {
            context.OnPhase("query-preflight");
            var existing = ReadQuery(context.Session, operation.QueryName);

            if (operation.Action == QueryAction.Create && existing is not null)
            {
                context.Checks.Add(new TaskCheck("query", false, $"A query named {operation.QueryName} already exists."));
                return new MutationStep.Finish(
                    new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected,
                        "A query with that name already exists; creating one never replaces it.",
                        context.Changes, context.Checks, CanRetry: true,
                        RetryReason: "Choose another name, or use Replace with the fingerprint from a Plan."),
                    "query creation");
            }

            if (operation.Action != QueryAction.Create && existing is null)
            {
                context.Checks.Add(new TaskCheck("query", false, $"No query named {operation.QueryName} exists in this workbook."));
                return new MutationStep.Finish(
                    new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected,
                        "The named query was not found.",
                        context.Changes, context.Checks, CanRetry: true,
                        RetryReason: "Run AuditWorkbookFlows to list the queries this workbook has."),
                    "query lookup");
            }

            // Length and fingerprint, never the expression. Both are facts about the query that a
            // caller needs and neither carries what the query connects to.
            context.Checks.Add(new TaskCheck("current-query", true, existing is null
                ? $"No query named {operation.QueryName} exists yet."
                : $"{operation.QueryName} is {existing.Length:N0} characters, fingerprint {existing.Sha256}."));

            if (!context.Apply)
            {
                var planned = operation.Action switch
                {
                    QueryAction.Create => $"create the query {operation.QueryName}",
                    QueryAction.Replace => $"replace the query {operation.QueryName}",
                    _ => $"delete the query {operation.QueryName}"
                };
                context.Changes.Add(new TaskChange("query", operation.QueryName, $"Planned to {planned}."));
                return new MutationStep.Finish(
                    new WorkbookExecutionOutcome(ExcelTaskStatus.Planned,
                        existing is null
                            ? $"Applying would {planned}. Nothing was changed."
                            : $"Applying would {planned}. Send fingerprint {existing.Sha256} with the Apply. Nothing was changed.",
                        context.Changes, context.Checks),
                    "query planning");
            }

            // The precondition, checked before anything is written. A query that moved since the
            // Plan is a different query, and replacing it would be replacing something the caller
            // never saw.
            if (existing is not null && !string.Equals(existing.Sha256, operation.ExpectedFormulaSha256, StringComparison.OrdinalIgnoreCase))
            {
                context.Checks.Add(new TaskCheck("query-fingerprint", false,
                    "The query changed since the Plan that produced this fingerprint; nothing was written."));
                return new MutationStep.Finish(
                    new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected,
                        "The expected query fingerprint does not match the query in the workbook.",
                        context.Changes, context.Checks, CanRetry: true,
                        RetryReason: "Run Plan again to take a current fingerprint, then resubmit."),
                    "the query fingerprint");
            }

            if (existing is not null)
            {
                context.Checks.Add(new TaskCheck("query-fingerprint", true, "The query matches the fingerprint the Plan reported."));
            }

            context.OnPhase("query");
            context.MarkMutationAttempted();
            ApplyQuery(context.Session, operation);

            var stored = ReadQuery(context.Session, operation.QueryName);
            var mismatch = FirstQueryMismatch(operation, stored);
            if (mismatch is not null)
            {
                context.Checks.Add(new TaskCheck("query", false, mismatch));
                return new MutationStep.Finish(
                    new WorkbookExecutionOutcome(ExcelTaskStatus.Unknown,
                        "Excel did not store the query change as requested; nothing was saved.",
                        context.Changes, context.Checks,
                        CanRetry: false, RetryReason: "Inspect the workbook's queries before retrying."),
                    "the query change");
            }

            var done = operation.Action switch
            {
                QueryAction.Create => $"Created the query {operation.QueryName}.",
                QueryAction.Replace => $"Replaced the query {operation.QueryName}.",
                _ => $"Deleted the query {operation.QueryName}."
            };
            context.Checks.Add(new TaskCheck("query", true, done));
            context.Changes.Add(new TaskChange("query", operation.QueryName, done));

            return new MutationStep.SaveAndVerify(
                verification => FirstQueryMismatch(operation, ReadQuery(verification, operation.QueryName)) is { } detail
                    ? (false, new TaskCheck("reopen-verification", false, $"After reopening the saved workbook: {detail}"))
                    : (true, new TaskCheck("reopen-verification", true, "The saved workbook reopened with the query as requested.")),
                $"{done} Saved and confirmed it after reopening.",
                "Excel saved the workbook, but reopen verification did not confirm the query.");
        });
    }

    /// <summary>
    /// One query's shape - never its expression. The fingerprint is taken over the normalized text
    /// so a line-ending difference between what Excel stores and what a caller sends cannot read as
    /// a changed query.
    /// </summary>
    private sealed record QuerySnapshot(string Name, int Length, string Sha256);

    private static QuerySnapshot? ReadQuery(ExcelSession session, string queryName)
    {
        using var references = new ComReferenceScope();
        var queries = GetOrNull(session.TargetWorkbook, "Queries");
        if (queries is null) return null;
        references.Add(queries);

        var count = Convert.ToInt32(Get(queries, "Count"), CultureInfo.InvariantCulture);
        for (var index = 1; index <= count; index++)
        {
            var query = references.Add(Item(queries, index));
            var name = GetOrNull(query, "Name") as string;
            if (!string.Equals(name, queryName, StringComparison.OrdinalIgnoreCase)) continue;

            var formula = GetOrNull(query, "Formula") as string ?? string.Empty;
            var normalized = MacroProcedureText.NormalizeLineEndings(formula);
            return new QuerySnapshot(name!, normalized.Length, MacroProcedureText.ComputeSha256(normalized));
        }

        return null;
    }

    private static void ApplyQuery(ExcelSession session, NormalizedManageQueryOperation operation)
    {
        using var references = new ComReferenceScope();
        var queries = references.Add(Get(session.TargetWorkbook, "Queries"));

        if (operation.Action == QueryAction.Create)
        {
            references.Add(Invoke(queries, "Add", operation.QueryName, operation.Formula)!);
            return;
        }

        var query = references.Add(Item(queries, operation.QueryName));
        if (operation.Action == QueryAction.Replace) Set(query, "Formula", operation.Formula!);
        else Invoke(query, "Delete");
    }

    private static string? FirstQueryMismatch(NormalizedManageQueryOperation operation, QuerySnapshot? stored)
    {
        if (operation.Action == QueryAction.Delete)
        {
            return stored is null ? null : $"the query {operation.QueryName} is still present after being deleted.";
        }

        if (stored is null) return $"no query named {operation.QueryName} was present after the change.";

        var expected = MacroProcedureText.ComputeSha256(operation.Formula!);
        return string.Equals(stored.Sha256, expected, StringComparison.OrdinalIgnoreCase)
            ? null
            : "Excel stored a different expression from the one requested.";
    }
}
