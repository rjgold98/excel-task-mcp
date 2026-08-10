using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using ExcelTask.Core;

namespace ExcelTask.Excel;

public sealed partial class ExcelWorkbookRuntime
{
    /// <summary>Items carried in the receipt. The real total is always reported alongside.</summary>
    private const int MaxAuditReceiptItems = 20;
    private const int XlExcelLinks = 1;
    private const int DispIdMemberNotFound = unchecked((int)0x80020003);

    /// <summary>
    /// Reads how one workbook's data flows fit together and reports names and shapes only. The
    /// session is opened read-only, nothing is saved, and the file's size and timestamp are compared
    /// before and after so the receipt can state that the workbook was not changed rather than
    /// merely promise it.
    /// </summary>
    private static WorkbookExecutionOutcome ExecuteAuditCore(ExcelTaskPlan plan, IExcelWorkbookRuntimeObserver observer)
    {
        var checks = new List<TaskCheck>();
        if (plan.Request.WorkbookBinding == WorkbookBinding.UseOpen)
        {
            return new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected,
                "Auditing opens its own read-only Excel instance.",
                Checks: [new TaskCheck("audit-binding", false, "Resubmit with workbook binding Isolated; the audit never attaches to a live session.")]);
        }

        string targetPath;
        try
        {
            targetPath = WorkbookRuntimeHelpers.NormalizePath(plan.Request.TargetWorkbookPath);
            WorkbookRuntimeHelpers.EnsureReadableWorkbook(targetPath, "Target workbook");
        }
        catch (InvalidOperationException)
        {
            return new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected,
                "The audit target workbook cannot be read.",
                Checks: [new TaskCheck("workbook-inputs", false, "The target workbook path is not readable.")]);
        }

        var stampBefore = ReadFileStamp(targetPath);
        ExcelSession? session = null;
        var items = new List<WorkbookFlowItem>();
        var totalFound = 0;
        try
        {
            observer.OnPhase("session-open");
            session = ExcelSession.Open(plan.Request, observer, readOnlyTarget: true);
            observer.OnPhase("audit-scan");
            totalFound = ScanWorkbookFlows(session.TargetWorkbook, items, checks);

            observer.OnPhase("audit-cleanup");
            var cleanupVerified = session.Close();
            session = null;
            if (!cleanupVerified)
            {
                checks.Add(new TaskCheck("owned-process-exit", false, "The owned audit Excel process did not exit."));
                return new WorkbookExecutionOutcome(ExcelTaskStatus.Unknown, "The audit completed but could not prove owned Excel cleanup.", Checks: checks, CanRetry: false, RetryReason: "Inspect the owned Excel process before retrying.");
            }
        }
        catch (Exception exception) when (exception is COMException or TargetInvocationException or InvalidOperationException or InvalidComObjectException)
        {
            var cleanupFailed = false;
            if (session is not null)
            {
                try { cleanupFailed = !session.Close(); }
                catch (Exception cleanupException) when (cleanupException is COMException or InvalidOperationException or TargetInvocationException) { cleanupFailed = true; }
                session = null;
            }

            checks.Add(new TaskCheck("audit-scan", false, "The workbook could not be opened or read for auditing."));
            if (cleanupFailed)
            {
                checks.Add(new TaskCheck("owned-process-exit", false, "The owned audit Excel process did not exit."));
                return new WorkbookExecutionOutcome(ExcelTaskStatus.Unknown, "The audit failed and could not prove owned Excel cleanup.", Checks: checks, CanRetry: false, RetryReason: "Inspect the owned Excel process before retrying.");
            }

            // Nothing was written and the owned instance is gone, so this is an ordinary rejection.
            return new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected, "The workbook could not be audited.", Checks: checks);
        }
        finally
        {
            session?.Close();
        }

        var stampAfter = ReadFileStamp(targetPath);
        var unchanged = stampBefore == stampAfter;
        checks.Add(new TaskCheck("workbook-unchanged", unchanged, unchanged
            ? "The workbook's size and timestamp are identical before and after the read-only audit."
            : "The workbook's size or timestamp changed during the audit; something else wrote to it."));

        var receipt = new WorkbookAuditReceipt(items, totalFound, Truncated: totalFound > items.Count, WorkbookUnchanged: unchanged);
        var summary = $"Audited {totalFound} data-flow item(s); the workbook was not modified.";
        return new WorkbookExecutionOutcome(
            plan.Request.Mode == ExcelTaskMode.Plan ? ExcelTaskStatus.Planned : ExcelTaskStatus.Completed,
            summary,
            Checks: checks,
            Audit: receipt);
    }

    private static (long Length, long WriteTimeTicks) ReadFileStamp(string path)
    {
        var info = new FileInfo(path);
        return (info.Length, info.LastWriteTimeUtc.Ticks);
    }

    /// <summary>
    /// Walks every flow surface the workbook exposes. Each section is guarded on its own: an Excel
    /// build without a Data Model, or a policy that hides one collection, costs that section and a
    /// named check rather than the whole audit.
    /// </summary>
    private static int ScanWorkbookFlows(object workbook, List<WorkbookFlowItem> items, List<TaskCheck> checks)
    {
        using var references = new ComReferenceScope();
        var totalFound = 0;
        var failedSections = new List<string>();
        // Connection load destinations, gathered first so queries can report where they load.
        var connectionInModel = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);

        void Report(string kind, string name, string detail, string? dependsOn = null)
        {
            totalFound++;
            if (items.Count < MaxAuditReceiptItems) items.Add(new WorkbookFlowItem(kind, name, detail, dependsOn));
        }

        void Section(string name, Action scan)
        {
            try { scan(); }
            catch (Exception exception) when (exception is COMException or TargetInvocationException or InvalidOperationException or InvalidCastException or OverflowException)
            {
                failedSections.Add(name);
            }
        }

        Section("connections", () =>
        {
            var connections = references.Add(Get(workbook, "Connections"));
            var count = Convert.ToInt32(Get(connections, "Count"), CultureInfo.InvariantCulture);
            for (var index = 1; index <= count; index++)
            {
                var connection = references.Add(AuditItem(connections, index));
                var name = (string)Get(connection, "Name");
                var inModel = false;
                try { inModel = Convert.ToBoolean(Get(connection, "InModel"), CultureInfo.InvariantCulture); }
                catch (Exception exception) when (exception is COMException or TargetInvocationException or InvalidOperationException) { }
                connectionInModel[name] = inModel;
                // The connection string is deliberately never read: it carries servers and credentials.
                Report("connection", name, inModel ? "feeds the data model" : "feeds the workbook");
            }
        });

        Section("queries", () =>
        {
            var queries = references.Add(Get(workbook, "Queries"));
            var count = Convert.ToInt32(Get(queries, "Count"), CultureInfo.InvariantCulture);
            for (var index = 1; index <= count; index++)
            {
                var query = references.Add(AuditItem(queries, index));
                var name = (string)Get(query, "Name");
                // The M formula is deliberately never read; only where the query's output goes.
                var detail = connectionInModel.TryGetValue($"Query - {name}", out var inModel)
                    ? (inModel ? "loads to the data model" : "loads to the workbook")
                    : "connection only";
                Report("query", name, detail);
            }
        });

        Section("data-model", () =>
        {
            var model = references.Add(Get(workbook, "Model"));
            var tables = references.Add(Get(model, "ModelTables"));
            var tableCount = Convert.ToInt32(Get(tables, "Count"), CultureInfo.InvariantCulture);
            for (var index = 1; index <= tableCount; index++)
            {
                var table = references.Add(AuditItem(tables, index));
                var name = (string)Get(table, "Name");
                var detail = "model table";
                try { detail = $"{Convert.ToInt64(Get(table, "RecordCount"), CultureInfo.InvariantCulture):N0} rows"; }
                catch (Exception exception) when (exception is COMException or TargetInvocationException or InvalidOperationException) { }
                Report("model-table", name, detail);
            }

            var relationships = references.Add(Get(model, "ModelRelationships"));
            var relationshipCount = Convert.ToInt32(Get(relationships, "Count"), CultureInfo.InvariantCulture);
            for (var index = 1; index <= relationshipCount; index++)
            {
                var relationship = references.Add(AuditItem(relationships, index));
                var foreignTable = (string)Get(references.Add(Get(references.Add(Get(relationship, "ForeignKeyColumn")), "Parent")), "Name");
                var primaryTable = (string)Get(references.Add(Get(references.Add(Get(relationship, "PrimaryKeyColumn")), "Parent")), "Name");
                var active = true;
                try { active = Convert.ToBoolean(Get(relationship, "Active"), CultureInfo.InvariantCulture); }
                catch (Exception exception) when (exception is COMException or TargetInvocationException or InvalidOperationException) { }
                Report("model-relationship", $"{foreignTable} -> {primaryTable}", active ? "active" : "inactive", primaryTable);
            }

            // Measures arrived later than tables and relationships, so they get their own guard.
            try
            {
                var measures = references.Add(Get(model, "ModelMeasures"));
                var measureCount = Convert.ToInt32(Get(measures, "Count"), CultureInfo.InvariantCulture);
                for (var index = 1; index <= measureCount; index++)
                {
                    var measure = references.Add(AuditItem(measures, index));
                    var name = (string)Get(measure, "Name");
                    string? table = null;
                    try { table = (string)Get(references.Add(Get(measure, "AssociatedTable")), "Name"); }
                    catch (Exception exception) when (exception is COMException or TargetInvocationException or InvalidOperationException) { }
                    Report("model-measure", name, "measure", table);
                }
            }
            catch (Exception exception) when (exception is COMException or TargetInvocationException or InvalidOperationException) { }
        });

        Section("pivot-tables", () =>
        {
            var sheets = references.Add(Get(workbook, "Worksheets"));
            var sheetCount = Convert.ToInt32(Get(sheets, "Count"), CultureInfo.InvariantCulture);
            for (var sheetIndex = 1; sheetIndex <= sheetCount; sheetIndex++)
            {
                var sheet = references.Add(AuditItem(sheets, sheetIndex));
                var sheetName = (string)Get(sheet, "Name");
                var pivots = references.Add(Invoke(sheet, "PivotTables") ?? throw new InvalidOperationException("Excel did not return 'PivotTables'."));
                var pivotCount = Convert.ToInt32(Get(pivots, "Count"), CultureInfo.InvariantCulture);
                for (var pivotIndex = 1; pivotIndex <= pivotCount; pivotIndex++)
                {
                    var pivot = references.Add(AuditItem(pivots, pivotIndex));
                    var name = (string)Get(pivot, "Name");
                    string? source = null;
                    try { source = Invoke(references.Add(Invoke(pivot, "PivotCache")!), "SourceData") as string; }
                    catch (Exception exception) when (exception is COMException or TargetInvocationException or InvalidOperationException)
                    {
                        // A Data Model pivot has no worksheet source range; that is itself the answer.
                        source = "data model";
                    }
                    Report("pivot", name, $"on sheet {sheetName}", source);
                }
            }
        });

        Section("external-links", () =>
        {
            // Returns a variant array of full paths, or nothing at all when the workbook links to
            // no other workbook. Only the file name is reported: the directory part carries server
            // and share names that must not leave the machine in a receipt.
            if (Invoke(workbook, "LinkSources", XlExcelLinks) is Array sources)
            {
                foreach (var source in sources)
                {
                    if (source is string path && path.Length > 0)
                    {
                        Report("external-link", Path.GetFileName(path), "linked workbook");
                    }
                }
            }
        });

        checks.Add(failedSections.Count == 0
            ? new TaskCheck("audit-scan", true, $"All flow surfaces were read; {totalFound} item(s) found.")
            : new TaskCheck("audit-scan", false, $"These surfaces could not be read and are missing from the report: {string.Join(", ", failedSections)}."));
        return totalFound;
    }

    /// <summary>
    /// Resolves one collection entry, tolerating both bindings. Excel's own collections expose Item
    /// as a parameterized property, but Office collections expose it as a method - a difference that
    /// produced two separate defects in one day - so this tries the property first and falls back on
    /// the exact missing-member failure.
    /// </summary>
    private static object AuditItem(object collection, object index)
    {
        try
        {
            return collection.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, collection, [index], CultureInfo.InvariantCulture)
                ?? throw new InvalidOperationException("The collection returned no entry.");
        }
        catch (Exception exception) when (IsMissingMember(exception))
        {
            return collection.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, collection, [index], CultureInfo.InvariantCulture)
                ?? throw new InvalidOperationException("The collection returned no entry.");
        }
    }

    private static bool IsMissingMember(Exception exception) =>
        (exception as COMException)?.HResult == DispIdMemberNotFound ||
        (exception is TargetInvocationException { InnerException: COMException inner } && inner.HResult == DispIdMemberNotFound) ||
        exception is MissingMethodException or MissingMemberException;
}
