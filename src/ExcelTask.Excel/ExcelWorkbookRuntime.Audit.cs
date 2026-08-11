using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
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
            var cleanupFailure = ExcelSession.CloseAndProve(ref session, "the audit", checks);
            if (cleanupFailure is not null) return cleanupFailure;
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
                var connection = references.Add(Item(connections, index));
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
                var query = references.Add(Item(queries, index));
                var name = (string)Get(query, "Name");
                // The M formula is deliberately never read; only where the query's output goes.
                var detail = connectionInModel.TryGetValue($"Query - {name}", out var inModel)
                    ? (inModel ? "loads to the data model" : "loads to the workbook")
                    : "connection only";
                Report("query", name, detail);
            }
        });

        // Scanned ahead of the model and pivots deliberately: when a caller was told to "fix this
        // macro", the procedure names are the items it cannot proceed without, so they must survive
        // the receipt's item cap ahead of anything merely descriptive. The first real field use
        // failed exactly here - EditMacroProcedure demands names the caller had no way to discover.
        Section("macros", () =>
        {
            var hasProject = false;
            try { hasProject = Convert.ToBoolean(Get(workbook, "HasVBProject"), CultureInfo.InvariantCulture); }
            catch (Exception exception) when (exception is COMException or TargetInvocationException or InvalidOperationException) { }
            if (!hasProject) return; // An .xlsx cannot hold VBA; absence is not a failure.

            var project = references.Add(Get(workbook, "VBProject"));
            var components = references.Add(Get(project, "VBComponents"));
            var componentCount = Convert.ToInt32(Get(components, "Count"), CultureInfo.InvariantCulture);
            for (var index = 1; index <= componentCount; index++)
            {
                var component = references.Add(Item(components, index));
                var componentName = (string)Get(component, "Name");
                var componentType = Convert.ToInt32(Get(component, "Type"), CultureInfo.InvariantCulture);
                var module = references.Add(Get(component, "CodeModule"));
                var totalLines = Convert.ToInt32(Get(module, "CountOfLines"), CultureInfo.InvariantCulture);
                var declarationLines = Convert.ToInt32(Get(module, "CountOfDeclarationLines"), CultureInfo.InvariantCulture);

                if (componentType == VbextCtStdModule)
                {
                    Report("macro-component", componentName, "standard module");
                    // Procedure names come from scanning the module's own text; the source itself
                    // never reaches the receipt. Only Sub and Function are listed, because those
                    // are the only shapes EditMacroProcedure can act on.
                    if (totalLines > 0)
                    {
                        var text = (string)Get(module, "Lines", 1, totalLines);
                        foreach (var line in text.Split('\n'))
                        {
                            var match = ProcedureNameRegex().Match(line.TrimEnd('\r'));
                            if (!match.Success) continue;
                            var procedureName = match.Groups["name"].Value;
                            var detail = "procedure";
                            try { detail = $"{Convert.ToInt32(Get(module, "ProcCountLines", procedureName, VbextPkProc), CultureInfo.InvariantCulture)} lines"; }
                            catch (Exception exception) when (exception is COMException or TargetInvocationException or InvalidOperationException) { }
                            Report("macro-procedure", $"{componentName}.{procedureName}", detail, componentName);
                        }
                    }
                }
                else if (totalLines > declarationLines)
                {
                    // A document, class, or form module carrying real code is worth knowing about
                    // even though EditMacroProcedure cannot touch it. Empty ones are noise.
                    Report("macro-component", componentName, componentType switch
                    {
                        2 => "class module with code, not editable",
                        3 => "form module with code, not editable",
                        _ => "document module with code, not editable"
                    });
                }
            }
        });

        Section("data-model", () =>
        {
            var model = references.Add(Get(workbook, "Model"));
            var tables = references.Add(Get(model, "ModelTables"));
            var tableCount = Convert.ToInt32(Get(tables, "Count"), CultureInfo.InvariantCulture);
            for (var index = 1; index <= tableCount; index++)
            {
                var table = references.Add(Item(tables, index));
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
                var relationship = references.Add(Item(relationships, index));
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
                    var measure = references.Add(Item(measures, index));
                    var name = (string)Get(measure, "Name");
                    string? table = null;
                    try { table = (string)Get(references.Add(Get(measure, "AssociatedTable")), "Name"); }
                    catch (Exception exception) when (exception is COMException or TargetInvocationException or InvalidOperationException) { }
                    Report("model-measure", name, "measure", table);
                }
            }
            catch (Exception exception) when (exception is COMException or TargetInvocationException or InvalidOperationException) { }
        });

        // Worksheets are walked once, reporting each sheet and the pivots on it. Naming the sheets
        // is what lets a caller aim the formula operations, which all take a worksheet name it
        // otherwise has no way to learn - the same discovery gap that sent a real macro task to a
        // different server.
        Section("worksheets", () =>
        {
            var sheets = references.Add(Get(workbook, "Worksheets"));
            var sheetCount = Convert.ToInt32(Get(sheets, "Count"), CultureInfo.InvariantCulture);
            for (var sheetIndex = 1; sheetIndex <= sheetCount; sheetIndex++)
            {
                var sheet = references.Add(Item(sheets, sheetIndex));
                var sheetName = (string)Get(sheet, "Name");
                var detail = "worksheet";
                try
                {
                    // Visibility matters to a caller choosing a target: xlSheetVisible is -1.
                    if (Convert.ToInt32(Get(sheet, "Visible"), CultureInfo.InvariantCulture) != -1) detail = "hidden worksheet";
                }
                catch (Exception exception) when (exception is COMException or TargetInvocationException or InvalidOperationException) { }

                // The used range says how much is on the sheet without reporting any of it.
                try
                {
                    var used = references.Add(Get(sheet, "UsedRange"));
                    detail += $", used range {(string)Get(used, "Address")}";
                }
                catch (Exception exception) when (exception is COMException or TargetInvocationException or InvalidOperationException) { }

                Report("worksheet", sheetName, detail);

                var pivots = references.Add(Invoke(sheet, "PivotTables") ?? throw new InvalidOperationException("Excel did not return 'PivotTables'."));
                var pivotCount = Convert.ToInt32(Get(pivots, "Count"), CultureInfo.InvariantCulture);
                for (var pivotIndex = 1; pivotIndex <= pivotCount; pivotIndex++)
                {
                    var pivot = references.Add(Item(pivots, pivotIndex));
                    var name = (string)Get(pivot, "Name");
                    // Classified, never quoted: SourceData is a connection string for an externally
                    // backed pivot, and connection strings carry servers and credentials. The kind
                    // of source is the flow information; the string itself is not ours to hand out.
                    var source = "pivot table";
                    try
                    {
                        var cache = references.Add(Invoke(pivot, "PivotCache") ?? throw new InvalidOperationException("Excel did not return 'PivotCache'."));
                        source = Convert.ToInt32(Get(cache, "SourceType"), CultureInfo.InvariantCulture) switch
                        {
                            1 => "from a worksheet range",
                            2 => "from an external connection",
                            3 => "from another pivot table",
                            4 => "from consolidated ranges",
                            _ => "pivot table"
                        };
                    }
                    catch (Exception exception) when (exception is COMException or TargetInvocationException or InvalidOperationException)
                    {
                        // A Data Model pivot commonly refuses these; that refusal is itself the answer.
                        source = "from the data model";
                    }

                    Report("pivot", name, source, sheetName);
                }

                // Tables are the structured targets an analyst aims at by name; 11 of 46 real
                // sessions listed them before working. The sheet-local address is safe to report.
                try
                {
                    var tables = references.Add(Get(sheet, "ListObjects"));
                    var tableCount = Convert.ToInt32(Get(tables, "Count"), CultureInfo.InvariantCulture);
                    for (var tableIndex = 1; tableIndex <= tableCount; tableIndex++)
                    {
                        var table = references.Add(Item(tables, tableIndex));
                        var tableName = (string)Get(table, "Name");
                        var tableDetail = "table";
                        try
                        {
                            var tableRange = references.Add(Get(table, "Range"));
                            tableDetail = $"table at {(string)Get(tableRange, "Address")}";
                        }
                        catch (Exception exception) when (exception is COMException or TargetInvocationException or InvalidOperationException) { }
                        Report("table", tableName, tableDetail, sheetName);
                    }
                }
                catch (Exception exception) when (exception is COMException or TargetInvocationException or InvalidOperationException) { }
            }
        });

        // Defined names, in 12 of 46 real sessions the way a workbook's important ranges are found.
        // The referred-to address is sheet-local and safe; an external reference is classified and
        // reduced to its file name, because the full RefersTo string carries directories.
        Section("named-ranges", () =>
        {
            var names = references.Add(Get(workbook, "Names"));
            var nameCount = Convert.ToInt32(Get(names, "Count"), CultureInfo.InvariantCulture);
            for (var index = 1; index <= nameCount; index++)
            {
                var definedName = references.Add(Item(names, index));
                // Hidden names are Excel plumbing - filter databases, print areas gone stale - and
                // reporting them would drown the twenty-item receipt in noise nobody asked for.
                try
                {
                    if (!Convert.ToBoolean(Get(definedName, "Visible"), CultureInfo.InvariantCulture)) continue;
                }
                catch (Exception exception) when (exception is COMException or TargetInvocationException or InvalidOperationException) { }

                var name = (string)Get(definedName, "Name");
                var detail = "defined name";
                string? dependsOn = null;
                try
                {
                    var refersTo = (string)Get(definedName, "RefersTo");
                    var bracket = refersTo.IndexOf('[', StringComparison.Ordinal);
                    if (bracket >= 0)
                    {
                        var close = refersTo.IndexOf(']', bracket);
                        detail = "refers to another workbook";
                        if (close > bracket) dependsOn = refersTo[(bracket + 1)..close];
                    }
                    else
                    {
                        detail = $"refers to {refersTo.TrimStart('=')}";
                    }
                }
                catch (Exception exception) when (exception is COMException or TargetInvocationException or InvalidOperationException) { }

                Report("named-range", name, detail, dependsOn);
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

    // Discovery only needs the names of Sub and Function procedures; Property accessors are not
    // editable and are deliberately not listed.
    [GeneratedRegex(@"^\s*(?:(?:Public|Private|Friend|Static)\s+)*(?<kind>Sub|Function)\s+(?<name>[A-Za-z_][A-Za-z0-9_]*)", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex ProcedureNameRegex();
}
