using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using ExcelTask.Excel;

namespace ExcelTask.McpServer;

/// <summary>
/// Builds the disposable workbooks the field check operates on, and owns every Excel process
/// it starts. Those processes are tracked and force-closed so the report can subtract them:
/// a leak figure is only meaningful if the harness cannot be mistaken for the product.
/// </summary>
[SupportedOSPlatform("windows")]
internal sealed class FieldCheckFixtures
{
    private const int XlOpenXmlWorkbookMacroEnabled = 52;

    private readonly List<ProcessIdentity> _ownedProcesses = [];

    public IReadOnlyList<ProcessIdentity> OwnedProcesses => _ownedProcesses;

    /// <summary>
    /// Identities, not process ids. The leak figure this check exists to report is the difference
    /// between two snapshots, and a process id alone can be recycled - the same hazard the runtime's
    /// dialog sentry refuses to accept before acting on a process. Reusing the product's identity
    /// means the reported number carries the same id, start time, and image path check the product
    /// itself demands.
    /// </summary>
    public static HashSet<ProcessIdentity> SnapshotExcelProcesses() => OwnedExcelProcess.SnapshotExcelProcesses();

    /// <summary>
    /// Counts the Excel processes the product left behind, waiting for a dying one to finish dying.
    ///
    /// Excel exits asynchronously, and the diagnostic trace measured its teardown at about 2.8
    /// seconds. A single snapshot taken shortly after an operation therefore counts an Excel on its
    /// way out as one that was abandoned - which made this check report leaked=2 result=FAIL on a
    /// machine whose full gate proves zero leaks across 241 tests. The tell was that the running
    /// count fell back to zero between operations. A genuine leak never clears, so waiting can turn
    /// a slow exit into a pass but can never hide a real leak.
    /// </summary>
    public static int CountLeakedAfterSettling(
        ISet<ProcessIdentity> before,
        IReadOnlyCollection<ProcessIdentity> harnessOwned,
        TimeSpan timeout) => LeakedAfterSettling(before, harnessOwned, timeout).Count;

    /// <summary>
    /// The same wait, returning which processes were still up rather than only how many.
    ///
    /// The identities are what let a per-operation figure be corrected later. An operation's own
    /// wait can expire while Excel is genuinely still shutting down - four connected COM add-ins on
    /// the work computer push teardown past twenty seconds where a clean machine finishes in three -
    /// and the run then prints "Leaked Excel: 2" beside an operation that leaked nothing. Keeping the
    /// identities means the final snapshot can settle the question exactly: a process still alive at
    /// the end leaked, one that has since exited was only slow.
    /// </summary>
    public static IReadOnlyList<ProcessIdentity> LeakedAfterSettling(
        ISet<ProcessIdentity> before,
        IReadOnlyCollection<ProcessIdentity> harnessOwned,
        TimeSpan timeout)
    {
        var deadline = DateTime.UtcNow + timeout;
        while (true)
        {
            var leaked = SnapshotExcelProcesses()
                .Where(identity => !before.Contains(identity) && !harnessOwned.Contains(identity))
                .ToArray();
            if (leaked.Length == 0 || DateTime.UtcNow >= deadline) return leaked;
            Thread.Sleep(250);
        }
    }

    public void CreateFormulaFixtures(string targetPath, string referencePath)
    {
        var application = Start();
        try
        {
            var workbooks = Get(application, "Workbooks");
            // Reference first, so the target can carry an external link to it for the audit to find.
            var reference = Invoke(workbooks, "Add")!;
            var referenceSheets = Get(reference, "Worksheets");
            var referenceSheet = Item(referenceSheets, 1);
            Set(referenceSheet, "Name", "Reference");
            Set(Get(referenceSheet, "Range", "A1"), "Formula", "=ROW()");
            Set(Get(referenceSheet, "Range", "A3"), "Formula", "=ROW()");
            Invoke(reference, "SaveAs", referencePath);
            Invoke(reference, "Close", false);

            var target = Invoke(workbooks, "Add")!;
            var sheets = Get(target, "Worksheets");
            var model = Item(sheets, 1);
            Set(model, "Name", "Model");
            Set(Get(model, "Range", "A1:D1"), "Formula", new object[,] { { "=1", "=2", "=3", "=4" } });
            Set(Get(model, "Range", "A2"), "Formula", "=A1*2");
            Set(Get(model, "Range", "B2"), "Formula", "=B1*2");
            // Audit surfaces, both tolerated as absent: a link into the reference workbook, in F1
            // where no formula operation looks, and one Power Query on builds that permit adding one.
            var referenceDirectory = Path.GetDirectoryName(Path.GetFullPath(referencePath));
            var referenceFile = Path.GetFileName(referencePath);
            Set(Get(model, "Range", "F1"), "Formula", $"='{referenceDirectory}\\[{referenceFile}]Reference'!$A$1");
            try
            {
                var queries = Get(target, "Queries");
                Invoke(queries, "Add", "FieldQuery", "let Source = #table({\"A\"}, {{1}}) in Source");
            }
            catch (Exception exception) when (exception is COMException or TargetInvocationException or InvalidOperationException)
            {
                // Older Excel or policy without Power Query; the audit reports what exists.
            }

            Invoke(target, "SaveAs", targetPath);
            Invoke(target, "Close", false);
        }
        finally
        {
            Close(application);
        }
    }

    /// <summary>Creates the macro workbook, or returns why the machine would not allow it.</summary>
    public string? TryCreateMacroFixture(string path, string componentName, string source)
    {
        var application = Start();
        try
        {
            var workbooks = Get(application, "Workbooks");
            var workbook = Invoke(workbooks, "Add")!;
            Invoke(workbook, "SaveAs", path, XlOpenXmlWorkbookMacroEnabled);
            var project = Get(workbook, "VBProject");
            var components = Get(project, "VBComponents");
            var component = Invoke(components, "Add", 1)!;
            Set(component, "Name", componentName);
            Invoke(Get(component, "CodeModule"), "AddFromString", source.ReplaceLineEndings("\r\n"));
            Invoke(workbook, "Save");
            Invoke(workbook, "Close", false);
            return null;
        }
        catch (Exception exception) when (exception is COMException or TargetInvocationException or InvalidOperationException)
        {
            return "The macro fixture could not be created, so macro editing was not measured. " +
                   "This usually means Trust Center does not permit programmatic access to the VBA project on this machine. " +
                   $"Underlying error: {exception.Message.ReplaceLineEndings(" ").Trim()}";
        }
        finally
        {
            Close(application);
        }
    }

    /// <summary>
    /// Creates a workbook whose Data Model holds two joinable tables, or returns why this machine
    /// would not. The fact table's key repeats and the lookup table's key is unique, because that is
    /// what a one-to-many relationship requires and Excel refuses one whose one side is not.
    ///
    /// A measure and a relationship both have nothing to act on without model tables, and a model
    /// table only exists once a query has been loaded into the model - so these steps depend on
    /// Power Query being permitted here. When it is not, the reason is reported and those rows are
    /// skipped, rather than a policy refusal being counted against the product. Both fixtures come
    /// from one Excel launch, which is most of what a step costs.
    /// </summary>
    public string? TryCreateModelFixture(string path, string factTable, string lookupTable)
    {
        var application = Start();
        try
        {
            var workbooks = Get(application, "Workbooks");
            var workbook = Invoke(workbooks, "Add")!;
            var queries = Get(workbook, "Queries");
            Invoke(queries, "Add", lookupTable, "let Source = #table({\"K\",\"Label\"}, {{1,\"a\"},{2,\"b\"}}) in Source");
            Invoke(queries, "Add", factTable, "let Source = #table({\"K\",\"Amount\"}, {{1,10},{1,20},{2,30}}) in Source");

            // CreateModelConnection is the argument that lands these in the model rather than on a
            // sheet; 6 is xlCmdExcel, and the Mashup provider is how a query is addressed as a
            // data source. Excel names each model table after its query, which is what the measure
            // and relationship steps then ask for.
            var connections = Get(workbook, "Connections");
            foreach (var name in new[] { lookupTable, factTable })
            {
                Invoke(connections, "Add2", name + "Conn", "field check model connection",
                    $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={name}",
                    name, 6, true, false);
            }

            var tables = Get(Get(workbook, "Model"), "ModelTables");
            var count = Convert.ToInt32(Get(tables, "Count"), CultureInfo.InvariantCulture);
            Invoke(workbook, "SaveAs", path);
            Invoke(workbook, "Close", false);
            return count >= 2
                ? null
                : $"Excel accepted the queries but produced {count} Data Model table(s) rather than two, so Data Model measures and relationships were not measured.";
        }
        catch (Exception exception) when (exception is COMException or TargetInvocationException or InvalidOperationException)
        {
            return "The Data Model fixture could not be created, so Data Model measures and relationships were not measured. " +
                   "This usually means Power Query or the Data Model is unavailable or not permitted on this machine. " +
                   $"Underlying error: {exception.Message.ReplaceLineEndings(" ").Trim()}";
        }
        finally
        {
            Close(application);
        }
    }

    public void TerminateOwnedProcesses()
    {
        foreach (var identity in _ownedProcesses)
        {
            // Identity-checked before the kill, so a recycled id cannot redirect it at a process
            // this check never started.
            if (!ProcessIdentity.TryOpenMatching(identity, out var process)) continue;
            using (process)
            {
                try
                {
                    if (!process.HasExited) process.Kill();
                }
                catch (InvalidOperationException) { }
                catch (System.ComponentModel.Win32Exception) { }
            }
        }
    }

    private object Start()
    {
        var before = SnapshotExcelProcesses();
        var type = Type.GetTypeFromProgID("Excel.Application", throwOnError: true)
            ?? throw new InvalidOperationException("Microsoft Excel is not registered on this machine.");
        var application = Activator.CreateInstance(type)
            ?? throw new InvalidOperationException("Microsoft Excel could not be started.");

        // Registered before it is configured, not after. A throw from either property write - and
        // RPC_E_SERVERCALL_RETRYLATER off a mid-launch Excel is exactly the kind this machine
        // produces - used to escape with a running Excel that TerminateOwnedProcesses could not
        // find and the leak count then charged to the product. This is the harness; being mistaken
        // for the thing it measures is the one failure it must not have.
        foreach (var identity in SnapshotExcelProcesses().Where(identity => !before.Contains(identity)))
        {
            if (!_ownedProcesses.Contains(identity)) _ownedProcesses.Add(identity);
        }

        Set(application, "Visible", false);
        Set(application, "DisplayAlerts", false);
        return application;
    }

    private static void Close(object application)
    {
        try { Invoke(application, "Quit"); }
        catch (COMException) { }
        catch (TargetInvocationException) { }
        finally
        {
            if (Marshal.IsComObject(application)) Marshal.FinalReleaseComObject(application);
        }
    }

    // The same late-bound rules the runtime uses, so a fixture cannot bind a member differently
    // from the product it exists to exercise.
    private static object Get(object target, string member, params object?[] arguments) => ComAccess.Get(target, member, arguments);

    private static void Set(object target, string member, object? value) => ComAccess.Set(target, member, value);

    private static object? Invoke(object target, string member, params object?[] arguments) => ComAccess.Invoke(target, member, arguments);

    private static object Item(object collection, object index) => ComAccess.Item(collection, index);
}
