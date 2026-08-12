using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using ExcelTask.Core;
using ExcelTask.Excel;
using Microsoft.Win32;
using ModelContextProtocol.Client;

namespace ExcelTask.McpServer;

internal sealed record ToolSurface(
    string Label,
    string Path,
    string? ServerName,
    string? ServerVersion,
    int ToolCount,
    string ToolNames,
    int ToolListBytes,
    double HandshakeSeconds,
    string? Error);

/// <summary>
/// Reports any operation this run never exercised.
///
/// The check shipped once covering five of eleven operations - everything from the first four
/// releases and nothing from the last four - so a field session validated the half already proven
/// and reported PASS. The step list is written by hand and always will be, because each step needs
/// its own fixture and arguments; what it must not do is stay silent about what it skipped. Labels
/// carry the operation name, so the set of kinds exercised can be read back off the results and
/// compared against the catalog the engine dispatches on.
/// </summary>
internal static class FieldCheckCoverage
{
    public static IReadOnlyList<string> UncoveredKinds(IEnumerable<OperationResult> operations)
    {
        var labels = operations.Select(operation => operation.Label).ToArray();
        return [.. OperationCatalog.AllKinds
            .Select(kind => kind.ToString())
            .Where(name => !labels.Any(label => label.StartsWith(name, StringComparison.Ordinal)))];
    }
}

internal sealed record OperationResult(
    string Label,
    string Status,
    double ElapsedSeconds,
    int LeakedExcel,
    string Summary,
    string Checks,
    string? Error,
    // Excel processes still up when this operation's own wait expired. Kept out of the report
    // because it is working data: the run's final snapshot turns it into the LeakedExcel above.
    [property: JsonIgnore] IReadOnlyList<ProcessIdentity>? StillUpAtDeadline = null);

/// <summary>
/// Validates ExcelTask on a real machine and records comparable measurements. This is compiled
/// rather than scripted on purpose: managed computers commonly run PowerShell in Constrained
/// Language Mode, which forbids the COM and reflection a scripted equivalent needs. It is not an
/// MCP tool and not a public CLI - the model-facing surface is still exactly one tool.
/// </summary>
[SupportedOSPlatform("windows")]
internal static class FieldCheck
{
    private static readonly JsonSerializerOptions ReportJsonOptions = new() { WriteIndented = true };

    private const string MacroComponent = "FieldModule";
    private const string MacroProcedure = "FieldMarker";
    private const string MacroBefore = "Public Sub FieldMarker()\n    ThisWorkbook.Worksheets(1).Range(\"A1\").Value2 = \"before\"\nEnd Sub";
    private const string MacroAfter = "Public Sub FieldMarker()\n    ThisWorkbook.Worksheets(1).Range(\"A1\").Value2 = \"after\"\nEnd Sub";
    private const string ModelTable = "FieldModel";
    private const string ModelLookupTable = "FieldLookup";
    private const string ModelMeasure = "FieldTotalK";

    public static async Task<int> RunAsync(string[] args)
    {
        string? serverPath = null, comparePath = null, outputDirectory = null;
        var compareArguments = new List<string>();
        for (var index = 1; index < args.Length; index++)
        {
            switch (args[index])
            {
                case "--server" when index + 1 < args.Length: serverPath = args[++index]; break;
                case "--compare" when index + 1 < args.Length: comparePath = args[++index]; break;
                case "--compare-arg" when index + 1 < args.Length: compareArguments.Add(args[++index]); break;
                case "--output" when index + 1 < args.Length: outputDirectory = args[++index]; break;
                default:
                    Console.Error.WriteLine($"Unrecognized argument '{args[index]}'.");
                    Console.Error.WriteLine("Usage: excel-task-mcp --field-check [--server <exe>] [--compare <exe>] [--compare-arg <arg>]... [--output <dir>]");
                    return 2;
            }
        }

        serverPath = System.IO.Path.GetFullPath(serverPath ?? Environment.ProcessPath
            ?? throw new InvalidOperationException("The running executable path could not be determined."));
        outputDirectory = System.IO.Path.GetFullPath(outputDirectory ?? System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "ExcelTask-FieldCheck"));
        Directory.CreateDirectory(outputDirectory);

        var work = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "ExcelTaskField-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(work);

        var notes = new List<string>();
        var operations = new List<OperationResult>();
        var surfaces = new List<ToolSurface>();
        var fixtures = new FieldCheckFixtures();
        var preExisting = FieldCheckFixtures.SnapshotExcelProcesses();
        // Probed once and reused. Probing again after cleanup would start an untracked Excel and
        // strand it, which is exactly the thing this report exists to measure honestly.
        var environment = new Dictionary<string, string>(StringComparer.Ordinal);

        try
        {
            Write("[1/4] Reading environment (read-only)...");
            environment = ReadEnvironment(notes);

            Write("[2/4] Building disposable workbooks...");
            fixtures.CreateFormulaFixtures(
                System.IO.Path.Combine(work, "target.xlsx"),
                System.IO.Path.Combine(work, "reference.xlsx"));
            var macroNote = fixtures.TryCreateMacroFixture(
                System.IO.Path.Combine(work, "macro-target.xlsm"), MacroComponent, MacroBefore);
            if (macroNote is not null) notes.Add(macroNote);
            var modelNote = fixtures.TryCreateModelFixture(
                System.IO.Path.Combine(work, "model-target.xlsx"), ModelTable, ModelLookupTable);
            if (modelNote is not null) notes.Add(modelNote);

            Write("[3/4] Measuring the advertised tool surface...");
            var mine = await FieldCheckProbe.MeasureAsync("ExcelTask", serverPath, [], work, CancellationToken.None);
            surfaces.Add(mine);
            Write(mine.Error is null
                ? $"      ExcelTask: {mine.ToolCount} tool(s), tools/list payload = {mine.ToolListBytes:N0} bytes"
                : $"      ExcelTask surface failed: {mine.Error}");

            if (comparePath is not null)
            {
                var other = await FieldCheckProbe.MeasureAsync("comparison server", System.IO.Path.GetFullPath(comparePath), compareArguments, work, CancellationToken.None);
                surfaces.Add(other);
                if (other.Error is null)
                {
                    var ratio = mine.ToolListBytes > 0 ? Math.Round((double)other.ToolListBytes / mine.ToolListBytes, 1) : 0;
                    Write($"      Comparison: {other.ToolCount} tool(s), {other.ToolListBytes:N0} bytes ({ratio.ToString(CultureInfo.InvariantCulture)}x ExcelTask)");
                }
                else
                {
                    Write($"      Comparison server did not report: {other.Error}");
                }
            }

            Write("[4/4] Exercising each operation on disposable workbooks...");
            await RunOperationsAsync(serverPath, work, macroNote is null, modelNote is null, fixtures, operations, notes);
            ReportCoverageGaps(operations, notes);
        }
        catch (Exception exception)
        {
            notes.Add($"The run stopped early: {exception.GetType().Name}: {exception.Message.ReplaceLineEndings(" ").Trim()}");
        }
        finally
        {
            fixtures.TerminateOwnedProcesses();
            await Task.Delay(800);
            try { if (Directory.Exists(work)) Directory.Delete(work, recursive: true); } catch (IOException) { }
        }

        // The headline number, so it gets the longest wait: teardown time scales with what the
        // operation did, and this figure is the product's central claim.
        var leaked = FieldCheckFixtures.LeakedAfterSettling(preExisting, fixtures.OwnedProcesses, TimeSpan.FromSeconds(30));
        ReconcilePerOperationLeaks(operations, leaked, notes);
        return WriteReport(outputDirectory, serverPath, environment, surfaces, operations, notes, leaked.Count);
    }

    /// <summary>
    /// Turns each operation's provisional figure into what actually leaked.
    ///
    /// An operation waits twenty seconds for the Excel processes it started to finish exiting. That
    /// is ample on a clean machine and not always enough on a managed one: the work computer loads
    /// four connected COM add-ins into every Excel instance and unloads them again on every exit.
    /// The run's own thirty-second snapshot is taken after everything has stopped and the harness
    /// has terminated what it owns, so it is the only authoritative answer. A process in an
    /// operation's list that is no longer running never leaked - it was still shutting down when
    /// that operation stopped looking.
    ///
    /// Without this the field report read "Leaked Excel: 2" against three operations while stating
    /// excelLeakedByProduct = 0 in the same file. Both cannot be true, and the one the eye lands on
    /// was the wrong one: a false accusation of the single defect this product exists to rule out,
    /// raised only on the class of machine whose verdict actually matters.
    /// </summary>
    internal static void ReconcilePerOperationLeaks(
        List<OperationResult> operations,
        IReadOnlyList<ProcessIdentity> leaked,
        List<string> notes)
    {
        var stillRunning = leaked.ToHashSet();
        var corrected = 0;
        for (var index = 0; index < operations.Count; index++)
        {
            var operation = operations[index];
            if (operation.StillUpAtDeadline is not { Count: > 0 } suspected) continue;

            var actual = suspected.Count(stillRunning.Contains);
            if (actual == operation.LeakedExcel) continue;
            operations[index] = operation with { LeakedExcel = actual };
            corrected++;
        }

        if (corrected > 0)
        {
            notes.Add($"{corrected} operation(s) still had an Excel process shutting down when their own twenty-second wait expired. Those processes had exited by the end of the run, so the per-operation counts above report only what was still running at the end. Excel teardown is slower here than that wait assumes, which is ordinary on a machine that loads COM add-ins into every instance.");
        }
    }

    /// <summary>
    /// Says out loud what this run did not exercise. A check that silently covers a subset reports
    /// PASS for the whole product, which is exactly how it once validated five of eleven operations
    /// and said nothing.
    /// </summary>
    private static void ReportCoverageGaps(List<OperationResult> operations, List<string> notes)
    {
        var total = OperationCatalog.AllKinds.Count;
        var uncovered = FieldCheckCoverage.UncoveredKinds(operations);
        if (uncovered.Count == 0)
        {
            Write($"      Coverage: all {total} operations exercised.");
            return;
        }

        var list = string.Join(", ", uncovered);
        notes.Add($"This run exercised {total - uncovered.Count} of {total} operations. Not exercised: {list}. A PASS covers only what ran.");
        Write($"      Coverage: {uncovered.Count} of {total} operation(s) not exercised - {list}");
    }

    private static async Task RunOperationsAsync(
        string serverPath,
        string work,
        bool macroReady,
        bool modelReady,
        FieldCheckFixtures fixtures,
        List<OperationResult> operations,
        List<string> notes)
    {
        var transport = new StdioClientTransport(new StdioClientTransportOptions
        {
            Name = "ExcelTask-FieldCheck",
            Command = serverPath,
            WorkingDirectory = work,
            ShutdownTimeout = TimeSpan.FromSeconds(5)
        });
        await using var client = await McpClient.CreateAsync(
            transport,
            new McpClientOptions { ClientInfo = new() { Name = "ExcelTask-FieldCheck", Version = "1.0.0" } });

        var target = System.IO.Path.Combine(work, "target.xlsx");
        var reference = System.IO.Path.Combine(work, "reference.xlsx");

        static ExcelOperation Copy(string reference) => new(
            ExcelOperationKind.CopyExhibit,
            CopyExhibit: new CopyExhibitOperation(reference, "Reference", "Exhibit A", ["A1:A3"]));

        await RunAsync(client, fixtures, operations, "CopyExhibit (Plan)", new ExcelTaskRequest(
            target, Copy(reference), ExcelTaskMode.Plan, WorkbookBinding.Isolated, SaveMode.Copy,
            System.IO.Path.Combine(work, "out-copy.xlsx"), OverwriteConfirmed: false));

        await RunAsync(client, fixtures, operations, "CopyExhibit (Apply)", new ExcelTaskRequest(
            target, Copy(reference), ExcelTaskMode.Apply, WorkbookBinding.Isolated, SaveMode.Copy,
            System.IO.Path.Combine(work, "out-copy.xlsx"), OverwriteConfirmed: true));

        await RunAsync(client, fixtures, operations, "ExtendFormulaSeries (Apply)", new ExcelTaskRequest(
            target,
            new ExcelOperation(
                ExcelOperationKind.ExtendFormulaSeries,
                ExtendFormulaSeries: new ExtendFormulaSeriesOperation("Model", FormulaExtensionDirection.Right, "A2:B2", "C2:D2")),
            ExcelTaskMode.Apply, WorkbookBinding.Isolated, SaveMode.Copy,
            System.IO.Path.Combine(work, "out-extend.xlsx"), OverwriteConfirmed: true));

        await RunAsync(client, fixtures, operations, "AuditWorkbookFlows (Apply)", new ExcelTaskRequest(
            target,
            new ExcelOperation(ExcelOperationKind.AuditWorkbookFlows, AuditWorkbookFlows: new AuditWorkbookFlowsOperation()),
            ExcelTaskMode.Apply, WorkbookBinding.Isolated, SaveMode.Same, null, OverwriteConfirmed: false));

        // The operations added since the first field check. Without these the check validated five
        // of eleven operations and nothing shipped in the four most recent releases - a work-computer
        // session spent proving the half that was already proven.
        await RunAsync(client, fixtures, operations, "ScanWorkbookStructure (Plan)", new ExcelTaskRequest(
            target,
            new ExcelOperation(ExcelOperationKind.ScanWorkbookStructure, ScanWorkbookStructure: new ScanWorkbookStructureOperation()),
            ExcelTaskMode.Plan, WorkbookBinding.Isolated, SaveMode.Same, null, OverwriteConfirmed: false));

        await RunAsync(client, fixtures, operations, "ReadWorksheetRange (Apply)", new ExcelTaskRequest(
            target,
            new ExcelOperation(
                ExcelOperationKind.ReadWorksheetRange,
                ReadWorksheetRange: new ReadWorksheetRangeOperation("Model", "A1:D4")),
            ExcelTaskMode.Apply, WorkbookBinding.Isolated, SaveMode.Same, null, OverwriteConfirmed: false));

        // Everything below writes, so each one works on its own copy: a failure in one must not
        // decide the next one's result.
        var writeTarget = System.IO.Path.Combine(work, "write-target.xlsx");
        System.IO.File.Copy(target, writeTarget, overwrite: true);
        await RunAsync(client, fixtures, operations, "WriteWorksheetValues (Apply)", new ExcelTaskRequest(
            writeTarget,
            new ExcelOperation(
                ExcelOperationKind.WriteWorksheetValues,
                WriteWorksheetValues: new WriteWorksheetValuesOperation("Model", [new WorksheetCellValue("A4", "FieldCheck")])),
            ExcelTaskMode.Apply, WorkbookBinding.Isolated, SaveMode.Same, null, OverwriteConfirmed: true));

        // Its own copy, because it repairs the same row SetNumberFormat later formats. Found by the
        // coverage reporter on its first run: this was the last operation the check never exercised,
        // which is precisely the gap that reporter exists to name.
        var repairTarget = System.IO.Path.Combine(work, "repair-target.xlsx");
        System.IO.File.Copy(target, repairTarget, overwrite: true);
        await RunAsync(client, fixtures, operations, "RepairExistingWorksheet (Apply)", new ExcelTaskRequest(
            repairTarget,
            new ExcelOperation(
                ExcelOperationKind.RepairExistingWorksheet,
                RepairExistingWorksheet: new RepairExistingWorksheetOperation("Model", ["A2:D2"])),
            ExcelTaskMode.Apply, WorkbookBinding.Isolated, SaveMode.Same, null, OverwriteConfirmed: true));

        await RunAsync(client, fixtures, operations, "FindReplace (Plan)", new ExcelTaskRequest(
            writeTarget,
            new ExcelOperation(
                ExcelOperationKind.FindReplace,
                FindReplace: new FindReplaceOperation("Model", "FieldCheck", Range: "A1:D10")),
            ExcelTaskMode.Plan, WorkbookBinding.Isolated, SaveMode.Same, null, OverwriteConfirmed: false));

        await RunAsync(client, fixtures, operations, "FindReplace (Apply)", new ExcelTaskRequest(
            writeTarget,
            new ExcelOperation(
                ExcelOperationKind.FindReplace,
                FindReplace: new FindReplaceOperation("Model", "FieldCheck", "FieldChecked", "A1:D10")),
            ExcelTaskMode.Apply, WorkbookBinding.Isolated, SaveMode.Same, null, OverwriteConfirmed: true));

        await RunAsync(client, fixtures, operations, "SetRangeFormat (Apply)", new ExcelTaskRequest(
            writeTarget,
            new ExcelOperation(
                ExcelOperationKind.SetRangeFormat,
                // More than the number format, deliberately: the fonts, fills and borders half is
                // the newest surface and the least exercised anywhere else.
                SetRangeFormat: new SetRangeFormatOperation(
                    "Model", "A1:B2", "#,##0.00", Bold: true, FillColor: "#EAF2ED", Borders: "Outline")),
            ExcelTaskMode.Apply, WorkbookBinding.Isolated, SaveMode.Same, null, OverwriteConfirmed: true));

        // Its own copy, because a table over A1:B2 would fight the formatting the step above put
        // there. Create then rename covers both the add and the rename paths in one Excel launch.
        var tableTarget = System.IO.Path.Combine(work, "table-target.xlsx");
        System.IO.File.Copy(target, tableTarget, overwrite: true);
        await RunAsync(client, fixtures, operations, "ManageTable (Create)", new ExcelTaskRequest(
            tableTarget,
            new ExcelOperation(
                ExcelOperationKind.ManageTable,
                ManageTable: new ManageTableOperation("Model", TableAction.Create, "FieldTable", Range: "A1:D2", TableStyle: "TableStyleMedium2")),
            ExcelTaskMode.Apply, WorkbookBinding.Isolated, SaveMode.Same, null, OverwriteConfirmed: true));

        await RunAsync(client, fixtures, operations, "ManageTable (Rename)", new ExcelTaskRequest(
            tableTarget,
            new ExcelOperation(
                ExcelOperationKind.ManageTable,
                ManageTable: new ManageTableOperation("Model", TableAction.Rename, "FieldTable", NewName: "FieldTableRenamed")),
            ExcelTaskMode.Apply, WorkbookBinding.Isolated, SaveMode.Same, null, OverwriteConfirmed: true));

        // Its own copy again, and self-contained: it creates the query it then deletes, so the step
        // does not depend on whether this Excel build let the fixture add one.
        var queryTarget = System.IO.Path.Combine(work, "query-target.xlsx");
        System.IO.File.Copy(target, queryTarget, overwrite: true);
        await RunAsync(client, fixtures, operations, "ManageQuery (Create)", new ExcelTaskRequest(
            queryTarget,
            new ExcelOperation(
                ExcelOperationKind.ManageQuery,
                ManageQuery: new ManageQueryOperation("FieldCheckQuery", QueryAction.Create,
                    Formula: "let Source = #table({\"A\"}, {{1}}) in Source")),
            ExcelTaskMode.Apply, WorkbookBinding.Isolated, SaveMode.Same, null, OverwriteConfirmed: true));

        // The last operation the coverage reporter still named as never exercised. It needs a model
        // table, which only exists once a query has been loaded into the model - so it runs on its
        // own fixture and only when this machine let that fixture be built. Create then delete,
        // because the delete is what proves the fingerprint precondition round-trips.
        if (modelReady)
        {
            var modelTarget = System.IO.Path.Combine(work, "model-target.xlsx");
            await RunAsync(client, fixtures, operations, "ManageModelMeasure (Create)", new ExcelTaskRequest(
                modelTarget,
                new ExcelOperation(
                    ExcelOperationKind.ManageModelMeasure,
                    ManageModelMeasure: new ManageModelMeasureOperation(
                        ModelTable, ModelMeasure, QueryAction.Create, Formula: $"SUM({ModelTable}[K])")),
                ExcelTaskMode.Apply, WorkbookBinding.Isolated, SaveMode.Same, null, OverwriteConfirmed: true));

            await RunAsync(client, fixtures, operations, "ManageModelMeasure (Delete)", new ExcelTaskRequest(
                modelTarget,
                new ExcelOperation(
                    ExcelOperationKind.ManageModelMeasure,
                    ManageModelMeasure: new ManageModelMeasureOperation(
                        ModelTable, ModelMeasure, QueryAction.Delete,
                        ExpectedFormulaSha256: MacroProcedureText.ComputeSha256($"SUM({ModelTable}[K])"))),
                ExcelTaskMode.Apply, WorkbookBinding.Isolated, SaveMode.Same, null, OverwriteConfirmed: true));

            // Create then delete on the same fixture, because a relationship left behind would
            // change what a re-run of the measure steps above is measuring.
            await RunAsync(client, fixtures, operations, "ManageModelRelationship (Create)", new ExcelTaskRequest(
                modelTarget,
                new ExcelOperation(
                    ExcelOperationKind.ManageModelRelationship,
                    ManageModelRelationship: new ManageModelRelationshipOperation(
                        ModelTable, "K", ModelLookupTable, "K", QueryAction.Create)),
                ExcelTaskMode.Apply, WorkbookBinding.Isolated, SaveMode.Same, null, OverwriteConfirmed: true));

            await RunAsync(client, fixtures, operations, "ManageModelRelationship (Delete)", new ExcelTaskRequest(
                modelTarget,
                new ExcelOperation(
                    ExcelOperationKind.ManageModelRelationship,
                    ManageModelRelationship: new ManageModelRelationshipOperation(
                        ModelTable, "K", ModelLookupTable, "K", QueryAction.Delete)),
                ExcelTaskMode.Apply, WorkbookBinding.Isolated, SaveMode.Same, null, OverwriteConfirmed: true));
        }

        // Creation names a path that must not exist, so it is the one operation the check must be
        // careful never to leave behind for a second run.
        var createdWorkbook = System.IO.Path.Combine(work, "created.xlsx");
        if (System.IO.File.Exists(createdWorkbook)) System.IO.File.Delete(createdWorkbook);
        await RunAsync(client, fixtures, operations, "Create (Workbook+Sheet)", new ExcelTaskRequest(
            createdWorkbook,
            new ExcelOperation(ExcelOperationKind.Create, Create: new CreateOperation(CreateKind.Workbook, "Summary")),
            ExcelTaskMode.Apply, WorkbookBinding.Isolated, SaveMode.Same, null, OverwriteConfirmed: false));

        if (!macroReady) return;

        var macroTarget = System.IO.Path.Combine(work, "macro-target.xlsm");
        var macroOutput = System.IO.Path.Combine(work, "out-macro.xlsm");
        var plan = await RunAsync(client, fixtures, operations, "EditMacroProcedure (Plan)", new ExcelTaskRequest(
            macroTarget,
            new ExcelOperation(
                ExcelOperationKind.EditMacroProcedure,
                EditMacroProcedure: new EditMacroProcedureOperation(MacroComponent, MacroProcedure)),
            ExcelTaskMode.Plan, WorkbookBinding.Isolated, SaveMode.Copy, macroOutput, OverwriteConfirmed: false));

        var hash = plan is { } receipt &&
                   receipt.TryGetProperty("macroProcedure", out var macro) &&
                   macro.ValueKind == JsonValueKind.Object &&
                   macro.TryGetProperty("sha256", out var sha)
            ? sha.GetString()
            : null;

        if (string.IsNullOrWhiteSpace(hash))
        {
            notes.Add("Macro Plan returned no procedure hash, so the Apply step was skipped. The Plan row's checks explain why; a blocked Trust Center setting is the usual cause.");
            return;
        }

        await RunAsync(client, fixtures, operations, "EditMacroProcedure (Apply+Run)", new ExcelTaskRequest(
            macroTarget,
            new ExcelOperation(
                ExcelOperationKind.EditMacroProcedure,
                EditMacroProcedure: new EditMacroProcedureOperation(MacroComponent, MacroProcedure, hash, MacroAfter, RunAfterEdit: true)),
            ExcelTaskMode.Apply, WorkbookBinding.Isolated, SaveMode.Copy, macroOutput, OverwriteConfirmed: true));
    }

    private static async Task<JsonElement?> RunAsync(
        McpClient client,
        FieldCheckFixtures fixtures,
        List<OperationResult> operations,
        string label,
        ExcelTaskRequest request)
    {
        var before = FieldCheckFixtures.SnapshotExcelProcesses();
        var timer = Stopwatch.StartNew();
        JsonElement? structured = null;
        string status = "Error", summary = "", checks = "";
        string? error = null;
        try
        {
            var result = await client.CallToolAsync("excel_task", new Dictionary<string, object?> { ["request"] = request });
            if (result.StructuredContent is { } content)
            {
                structured = content;
                status = content.TryGetProperty("status", out var s) ? s.GetString() ?? "Error" : "Error";
                summary = content.TryGetProperty("summary", out var m) ? m.GetString() ?? "" : "";
                checks = content.TryGetProperty("checks", out var c) && c.ValueKind == JsonValueKind.Array
                    ? string.Join("; ", c.EnumerateArray().Select(check =>
                        $"{check.GetProperty("name").GetString()}={check.GetProperty("passed").GetBoolean()}"))
                    : "";
            }
            else
            {
                error = "The tool returned no structured content.";
            }
        }
        catch (Exception exception)
        {
            error = $"{exception.GetType().Name}: {exception.Message.ReplaceLineEndings(" ").Trim()}";
        }

        timer.Stop();
        // Waits for a dying Excel rather than counting it. Bounded, so a genuine leak still reports.
        // Provisional: whatever is still up here is reconciled against the run's final snapshot
        // before the report is written, so this figure never reaches the reader as a leak claim.
        var stillUp = FieldCheckFixtures.LeakedAfterSettling(before, fixtures.OwnedProcesses, TimeSpan.FromSeconds(20));

        operations.Add(new OperationResult(
            label, status, Math.Round(timer.Elapsed.TotalSeconds, 2), stillUp.Count, summary, checks, error, stillUp));
        Write($"      {label,-34} {status,-16} {timer.Elapsed.TotalSeconds,6:F1}s  excelStillUp={stillUp.Count}");
        return structured;
    }

    /// <summary>
    /// The two things only a managed, synced machine can answer - measured here so the field run
    /// stays one command, and reported so that neither answer names the machine it came from.
    ///
    /// The first is the sync mapping. v0.16.0 resolves a workbook Excel reports as a service URL
    /// back to the local path the caller named, through the sync client's own registry entries, and
    /// that lookup has never executed: a machine that syncs nothing registers no providers, so only
    /// the arithmetic around it is covered by tests. Counts are reported and values are not, because
    /// a UrlNamespace is the tenant and site - an internal server name - and a MountPoint is a
    /// person's directory layout.
    ///
    /// The second is folder writability. Until v0.16.0 a directory carrying the ReadOnly attribute
    /// was treated as unwritable, and Windows sets that attribute on Documents, Downloads, Desktop
    /// and the OneDrive root of an ordinary profile - so every copy-save and every create into the
    /// folders workbooks actually live in was refused, before Excel started, with a reason that was
    /// not true. Folders are named by label, never by path.
    /// </summary>
    private static void ProbeStorage(Dictionary<string, string> environment, List<string> notes)
    {
        var (registered, resolving) = OneDriveSyncMap.SelfTest();
        environment["syncRootsRegistered"] = registered.ToString(CultureInfo.InvariantCulture);
        environment["syncPathsResolving"] = registered == 0
            ? "n/a - nothing synced on this machine"
            : $"{resolving} of {registered}";

        if (registered == 0)
        {
            notes.Add("No OneDrive sync roots are registered, so the SharePoint-URL identity path could not be exercised. A UseOpen against a synced workbook is what proves it.");
        }
        else if (resolving < registered)
        {
            notes.Add($"{registered - resolving} of {registered} sync roots did not resolve a path beneath themselves back to their own namespace. UseOpen against a workbook in one of those will still refuse; the mapping shape differs from what the resolver expects.");
        }

        // Labels, never paths. A OneDrive root names the tenant in its own folder name.
        (string Label, string? Path)[] folders =
        [
            ("documents", Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)),
            ("desktop", Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)),
            ("downloads", System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads")),
            ("oneDriveRoot", Environment.GetEnvironmentVariable("OneDrive"))
        ];

        var refusedBefore = new List<string>();
        foreach (var (label, path) in folders)
        {
            if (string.IsNullOrWhiteSpace(path) || !Directory.Exists(path))
            {
                environment[$"folder:{label}"] = "absent";
                continue;
            }

            var readOnly = false;
            try { readOnly = (File.GetAttributes(path) & FileAttributes.ReadOnly) != 0; }
            catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) { }

            var accepts = WorkbookRuntimeHelpers.DirectoryAcceptsNewFile(path);
            environment[$"folder:{label}"] = $"readOnlyAttribute={(readOnly ? "yes" : "no")} acceptsNewFile={(accepts ? "yes" : "no")}";
            if (readOnly && accepts) refusedBefore.Add(label);
        }

        if (refusedBefore.Count > 0)
        {
            notes.Add($"These folders carry the ReadOnly attribute and are nevertheless writable: {string.Join(", ", refusedBefore)}. Before v0.16.0 every copy-save and create into them was refused; this run confirms the attribute test is gone.");
        }
    }

    private static Dictionary<string, string> ReadEnvironment(List<string> notes)
    {
        var environment = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["computerName"] = Environment.MachineName,
            ["osVersion"] = Environment.OSVersion.VersionString,
            ["runtime"] = RuntimeInformation.FrameworkDescription,
            ["dotnetRoot"] = WithoutUserProfile(Environment.GetEnvironmentVariable("DOTNET_ROOT")),
            // Named for the one thing it reads. AppLocker and WDAC also put PowerShell into
            // Constrained Language Mode without ever setting this variable, so "not set" is not
            // "no lockdown" - and a security control reported absent while it is being enforced is
            // the more damaging of the two mistakes this line can make.
            ["psLockdownPolicyVariable"] = Environment.GetEnvironmentVariable("__PSLockdownPolicy") ?? "not set"
        };

        ProbeStorage(environment, notes);

        environment["excelRunningBefore"] = FieldCheckFixtures.SnapshotExcelProcesses().Count.ToString(CultureInfo.InvariantCulture);
        try
        {
            var type = Type.GetTypeFromProgID("Excel.Application", throwOnError: true)!;
            var application = Activator.CreateInstance(type)!;
            try
            {
                environment["excelVersion"] = Read(application, "Version");
                environment["excelBuild"] = Read(application, "Build");
                var addins = new List<string>();
                // Every interface handed out here is released before the probe ends. A retained one
                // keeps the probe's Excel alive past Quit, and the next activation then meets a
                // half-dead instance - the likely shape of the CO_E_SERVER_EXEC_FAILURE seen on the
                // work computer in 0.5.0.
                var comAddins = ComAccess.Get(application, "COMAddIns");
                try
                {
                    if (comAddins is not null)
                    {
                        var count = Convert.ToInt32(ComAccess.Get(comAddins, "Count"), CultureInfo.InvariantCulture);
                        for (var index = 1; index <= count; index++)
                        {
                            // ComAccess.Item resolves whichever binding this object model uses: Office
                            // exposes COMAddIns.Item as a method where Excel exposes its own collections'
                            // Item as a parameterized property.
                            var addin = ComAccess.Item(comAddins, index);
                            try
                            {
                                if (Convert.ToBoolean(ComAccess.Get(addin, "Connect"), CultureInfo.InvariantCulture))
                                {
                                    addins.Add(Read(addin, "ProgId"));
                                }
                            }
                            finally
                            {
                                ComAccess.Release(addin);
                            }
                        }
                    }
                }
                finally
                {
                    if (comAddins is not null && Marshal.IsComObject(comAddins)) Marshal.FinalReleaseComObject(comAddins);
                }

                environment["connectedComAddins"] = addins.Count > 0 ? string.Join("; ", addins) : "none";
            }
            finally
            {
                try { application.GetType().InvokeMember("Quit", BindingFlags.InvokeMethod, null, application, null, CultureInfo.InvariantCulture); }
                catch (COMException) { }
                catch (TargetInvocationException) { }
                if (Marshal.IsComObject(application)) Marshal.FinalReleaseComObject(application);
            }
        }
        catch (Exception exception) when (exception is COMException or TargetInvocationException or InvalidOperationException or NotSupportedException)
        {
            // Recorded alongside whatever was already read rather than over it. Enumerating add-ins
            // can fail on its own, and losing a good Excel version to that would be a poor trade.
            environment["excelProbeError"] = exception.Message.ReplaceLineEndings(" ").Trim();
            environment.TryAdd("excelVersion", "unavailable");
        }

        // Read last, because it is keyed on the Excel that actually answered the probe.
        ReadOfficeSecurity(environment, notes);

        return environment;

        static string Read(object target, string member) =>
            target.GetType().InvokeMember(member, BindingFlags.GetProperty, null, target, null, CultureInfo.InvariantCulture)?.ToString() ?? "";
    }

    /// <summary>
    /// Rewrites the user profile directory as %USERPROFILE%.
    ///
    /// The Windows account name is an employee number at most large employers, and this report
    /// already carries the machine name, so leaving the account name in a path names a person as
    /// well as a computer. It earns nothing in exchange: what matters about DOTNET_ROOT or the
    /// server path is whether it is set and where it sits relative to the profile, never whose
    /// profile it is. The sibling artifact - the diagnostic trace - already promises never to
    /// record a full path or a user name; this is the same promise kept in the other file.
    /// </summary>
    internal static string WithoutUserProfile(string? value)
    {
        if (string.IsNullOrEmpty(value)) return "not set";
        var profile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        return string.IsNullOrEmpty(profile)
            ? value
            : value.Replace(profile, "%USERPROFILE%", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Excel's macro-trust settings, read from the registry hive belonging to the Excel that is
    /// actually running.
    ///
    /// This used to walk 16.0 then 15.0 and overwrite as it went, so the last hive present won.
    /// A machine carrying a stale Office 15.0 key reported officeVersion 15.0 while Excel 16.0 was
    /// the version answering every probe - and, because the same loop assigned them, accessVBOM and
    /// vbaWarnings came from that dead hive too. Those two values are the entire reason this section
    /// exists: they are what decides whether macro editing can work on a managed computer. Read from
    /// the wrong Office, they answer for an Excel that is not the one about to open the workbook,
    /// and the report states them as plainly as if they were right.
    /// </summary>
    private static void ReadOfficeSecurity(Dictionary<string, string> environment, List<string> notes)
    {
        // Excel reports its version as "16.0", which is already how the hive is named.
        var running = environment.GetValueOrDefault("excelVersion");
        string[] candidates = running is { Length: > 0 } and not "unavailable"
            ? [running, "16.0", "15.0"]
            : ["16.0", "15.0"];

        foreach (var version in candidates)
        {
            using var security = Registry.CurrentUser.OpenSubKey($@"Software\Microsoft\Office\{version}\Excel\Security");
            if (security is null) continue;

            environment["officeVersion"] = version;
            // 1 means programmatic VBA project access is trusted; without it macro work is impossible.
            environment["accessVBOM"] = security.GetValue("AccessVBOM")?.ToString() ?? "not set";
            // 1 enable all, 2 disable with notification, 3 signed only, 4 disable all.
            environment["vbaWarnings"] = security.GetValue("VBAWarnings")?.ToString() ?? "not set";

            using var policy = Registry.CurrentUser.OpenSubKey($@"Software\Policies\Microsoft\Office\{version}\Excel\Security");
            if (policy is not null)
            {
                environment["policyAccessVBOM"] = policy.GetValue("AccessVBOM")?.ToString() ?? "not set";
                environment["policyVBAWarnings"] = policy.GetValue("VBAWarnings")?.ToString() ?? "not set";
                notes.Add("Group Policy sets Excel macro security on this machine; the organization controls it, not this product.");
            }

            if (running is not null && version != running)
            {
                notes.Add($"Excel {running} is running but carries no macro-security key of its own, so accessVBOM and vbaWarnings were read from the Office {version} hive and describe that version. The EditMacroProcedure steps below are what actually settle whether macro work is permitted here.");
            }

            return;
        }

        notes.Add("No Excel macro-security registry key was found. accessVBOM and vbaWarnings are unknown rather than permissive; the EditMacroProcedure steps are the only evidence in this report about macro trust.");
    }

    private static int WriteReport(
        string outputDirectory,
        string serverPath,
        Dictionary<string, string> environment,
        List<ToolSurface> surfaces,
        List<OperationResult> operations,
        List<string> notes,
        int leaked)
    {
        environment["excelLeakedByProduct"] = leaked.ToString(CultureInfo.InvariantCulture);
        var stamp = DateTime.Now.ToString("yyyy-MM-dd-HHmm", CultureInfo.InvariantCulture);
        var jsonPath = System.IO.Path.Combine(outputDirectory, $"exceltask-fieldcheck-{stamp}.json");
        var markdownPath = System.IO.Path.Combine(outputDirectory, $"exceltask-fieldcheck-{stamp}.md");

        File.WriteAllText(jsonPath, JsonSerializer.Serialize(new
        {
            completedUtc = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture),
            serverPath = WithoutUserProfile(serverPath),
            environment,
            // Same rule as serverPath: the probe is told where each server lives, the report is not.
            toolSurfaces = surfaces.Select(surface => surface with { Path = WithoutUserProfile(surface.Path) }),
            operations,
            notes
        }, ReportJsonOptions), Encoding.UTF8);

        var markdown = new StringBuilder()
            .AppendLine("# ExcelTask field check").AppendLine()
            .AppendLine(CultureInfo.InvariantCulture, $"Run {DateTime.UtcNow:u} on {environment.GetValueOrDefault("computerName", "unknown")}.").AppendLine()
            .AppendLine("## Environment").AppendLine()
            .AppendLine("| Item | Value |").AppendLine("|---|---|");
        foreach (var (key, value) in environment) markdown.AppendLine(CultureInfo.InvariantCulture, $"| {key} | {value} |");

        markdown.AppendLine().AppendLine("## Tool surface").AppendLine()
            .AppendLine("| Server | Name | Version | Tools | tools/list bytes | Handshake s | Error |")
            .AppendLine("|---|---|---|---|---|---|---|");
        foreach (var surface in surfaces)
        {
            markdown.AppendLine(CultureInfo.InvariantCulture,
                $"| {surface.Label} | {surface.ServerName} | {surface.ServerVersion} | {surface.ToolCount} | {surface.ToolListBytes:N0} | {surface.HandshakeSeconds} | {surface.Error} |");
        }

        if (surfaces.Count > 1 && surfaces[0].ToolListBytes > 0 && surfaces[1].Error is null)
        {
            var ratio = Math.Round((double)surfaces[1].ToolListBytes / surfaces[0].ToolListBytes, 1);
            markdown.AppendLine().AppendLine(CultureInfo.InvariantCulture,
                $"The comparison server's `tools/list` payload is **{ratio.ToString(CultureInfo.InvariantCulture)}x** the size of ExcelTask's. That payload is carried in context every session, before any work is requested.");
        }

        markdown.AppendLine().AppendLine("## Operations").AppendLine()
            .AppendLine("| Operation | Status | Seconds | Leaked Excel | Summary |").AppendLine("|---|---|---|---|---|");
        foreach (var operation in operations)
        {
            markdown.AppendLine(CultureInfo.InvariantCulture,
                $"| {operation.Label} | {operation.Status} | {operation.ElapsedSeconds} | {operation.LeakedExcel} | {operation.Error ?? operation.Summary} |");
        }

        if (notes.Count > 0)
        {
            markdown.AppendLine().AppendLine("## Notes").AppendLine();
            foreach (var note in notes) markdown.AppendLine(CultureInfo.InvariantCulture, $"- {note}");
        }

        markdown.AppendLine().AppendLine("## Checks in full").AppendLine();
        foreach (var operation in operations) markdown.AppendLine(CultureInfo.InvariantCulture, $"- **{operation.Label}**: {operation.Checks}");
        File.WriteAllText(markdownPath, markdown.ToString(), Encoding.UTF8);

        var failed = operations.Count(operation => operation.Status is not ("Completed" or "Planned"));
        var passed = failed == 0 && leaked == 0 && operations.Count > 0;
        var digestPath = Path.Combine(outputDirectory, $"exceltask-digest-{stamp}.txt");
        var digest = BuildDigest(environment, surfaces, operations, leaked, passed);
        File.WriteAllText(digestPath, digest, Encoding.UTF8);

        Console.WriteLine();
        Write("Report written:");
        Write("  " + markdownPath);
        Write("  " + jsonPath);
        Write("  " + digestPath);
        Console.WriteLine();
        // Printed as well as written because a managed computer often cannot move a file off itself.
        // Everything needed to decide what to do next is here, short enough to retype or photograph.
        Write(digest);

        if (!passed)
        {
            Write("Some operations did not complete, or an Excel process was left behind.");
            return 1;
        }

        Write("All operations completed and no Excel process was left behind.");
        return 0;
    }

    /// <summary>A dense, transcribable summary for machines that cannot send a file anywhere.</summary>
    private static string BuildDigest(
        Dictionary<string, string> environment,
        List<ToolSurface> surfaces,
        List<OperationResult> operations,
        int leaked,
        bool passed)
    {
        // Every value is length-capped: the digest is only useful if it stays transcribable, and an
        // unbounded probe error would otherwise swallow the line it appears on.
        static string Get(Dictionary<string, string> source, string key)
        {
            var value = source.GetValueOrDefault(key, "?");
            return value.Length <= 24 ? value : value[..24];
        }

        var digest = new StringBuilder()
            .AppendLine("----- EXCELTASK FIELD DIGEST -----")
            .AppendLine(CultureInfo.InvariantCulture,
                $"excel={Get(environment, "excelVersion")}.{Get(environment, "excelBuild")} vbom={Get(environment, "accessVBOM")} vbawarn={Get(environment, "vbaWarnings")}");
        if (environment.ContainsKey("excelProbeError")) digest.AppendLine("NOTE excel probe partly failed, see report");

        var mine = surfaces.FirstOrDefault(surface => surface.Error is null);
        if (mine is null)
        {
            digest.AppendLine("tools self=FAILED");
        }
        else
        {
            digest.AppendLine(CultureInfo.InvariantCulture,
                $"self  v{mine.ServerVersion} tools={mine.ToolCount} bytes={mine.ToolListBytes}");
            var other = surfaces.Skip(1).FirstOrDefault();
            digest.AppendLine(other switch
            {
                null => "other none",
                { Error: not null } => "other FAILED",
                _ => string.Create(CultureInfo.InvariantCulture,
                    $"other tools={other.ToolCount} bytes={other.ToolListBytes} ratio={(mine.ToolListBytes > 0 ? Math.Round((double)other.ToolListBytes / mine.ToolListBytes, 1) : 0)}x")
            });
        }

        foreach (var operation in operations)
        {
            var label = operation.Label.Length <= 26 ? operation.Label : operation.Label[..26];
            digest.AppendLine(CultureInfo.InvariantCulture,
                $"{label,-26} {operation.Status,-16} {operation.ElapsedSeconds,6:F1}s L{operation.LeakedExcel}");
        }

        return digest
            .AppendLine(CultureInfo.InvariantCulture, $"leaked={leaked} result={(passed ? "PASS" : "FAIL")}")
            .AppendLine("----- END DIGEST -----")
            .ToString();
    }

    private static void Write(string message) => Console.WriteLine(message);
}
