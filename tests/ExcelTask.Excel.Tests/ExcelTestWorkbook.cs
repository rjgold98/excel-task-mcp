using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;

namespace ExcelTask.Excel.Tests;

internal static class ExcelTestWorkbook
{
    public static void CreateTarget(string path) => Create(path, null);

    public static void CreateFormulaTarget(string path, string range, object?[,] formulas, string? constantCell = null, object? constantValue = null) => Create(path, workbook =>
    {
        var sheets = Get(workbook, "Worksheets");
        var sheet = Item(sheets, 1);
        var target = Get(sheet, "Range", range);
        Set(target, "FormulaR1C1", formulas);
        if (constantCell is not null)
        {
            var constant = Get(sheet, "Range", constantCell);
            Set(constant, "Value2", constantValue);
            Release(constant);
        }
        Release(target);
        Release(sheet);
        Release(sheets);
    });

    public static void CreateReference(string path) => Create(path, workbook =>
    {
        var sheets = Get(workbook, "Worksheets");
        var sheet = Item(sheets, 1);
        Set(sheet, "Name", "Reference");
        var range = Get(sheet, "Range", "A1:A3");
        Set(range, "FormulaR1C1", new object?[,] { { "=ROW()" }, { null }, { "=ROW()" } });
        Release(range);
        Release(sheet);
        Release(sheets);
    });

    /// <summary>
    /// A workbook with audit surfaces: one external-link formula into the reference workbook, and
    /// one Power Query when this Excel permits adding one. Returns whether the query was created.
    /// </summary>
    public static bool CreateAuditTarget(string path, string referencePath)
    {
        var hasQuery = false;
        Create(path, workbook =>
        {
            var sheets = Get(workbook, "Worksheets");
            var sheet = Item(sheets, 1);
            var directory = Path.GetDirectoryName(Path.GetFullPath(referencePath));
            var file = Path.GetFileName(referencePath);
            Set(Get(sheet, "Range", "F1"), "Formula", $"='{directory}\\[{file}]Reference'!$A$1");
            try
            {
                var queries = Get(workbook, "Queries");
                Invoke(queries, "Add", "AuditQuery", "let Source = #table({\"A\"}, {{1}}) in Source");
                hasQuery = true;
                Release(queries);
            }
            catch (Exception exception) when (exception is System.Runtime.InteropServices.COMException or System.Reflection.TargetInvocationException)
            {
                // This Excel build or policy does not permit adding queries; the audit reports what exists.
            }
            Release(sheet);
            Release(sheets);
        });
        return hasQuery;
    }

    /// <summary>
    /// A workbook whose Data Model holds one table, which is the only way to get a model table to
    /// exist: a model table is not created directly, it appears when a query is loaded into the
    /// model. Returns false when this Excel build or policy will not do it, so the test can say so
    /// rather than fail for the wrong reason.
    /// </summary>
    public static bool CreateModelTarget(string path, string queryName)
    {
        var hasModel = false;
        Create(path, workbook =>
        {
            try
            {
                var queries = Get(workbook, "Queries");
                Invoke(queries, "Add", queryName, "let Source = #table({\"K\",\"V\"}, {{1,\"a\"},{2,\"b\"}}) in Source");
                Release(queries);

                // CreateModelConnection is the argument that makes this land in the model rather
                // than on a sheet; 6 is xlCmdExcel, and the Mashup provider is how a query is
                // addressed as a data source.
                var connections = Get(workbook, "Connections");
                Invoke(connections, "Add2", queryName + "Conn", "test model connection",
                    $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={queryName}",
                    queryName, 6, true, false);
                Release(connections);

                var model = Get(workbook, "Model");
                var tables = Get(model, "ModelTables");
                hasModel = Convert.ToInt32(Get(tables, "Count"), System.Globalization.CultureInfo.InvariantCulture) > 0;
                Release(tables);
                Release(model);
            }
            catch (Exception exception) when (exception is System.Runtime.InteropServices.COMException or System.Reflection.TargetInvocationException)
            {
                // No Power Query, no Data Model, or policy forbids the connection.
            }
        });
        return hasModel;
    }

    /// <summary>
    /// A workbook whose Data Model holds two tables that can legally be joined: the fact table's
    /// key repeats and the lookup table's key is unique, which is what a one-to-many relationship
    /// requires. Excel refuses the relationship outright if the one side is not unique, so the
    /// fixture has to be built this way for the operation to have anything real to do.
    /// </summary>
    public static bool CreateModelPair(string path, string factTable, string lookupTable)
    {
        var hasBoth = false;
        Create(path, workbook =>
        {
            try
            {
                var queries = Get(workbook, "Queries");
                Invoke(queries, "Add", lookupTable, "let Source = #table({\"K\",\"Label\"}, {{1,\"a\"},{2,\"b\"}}) in Source");
                Invoke(queries, "Add", factTable, "let Source = #table({\"K\",\"Amount\"}, {{1,10},{1,20},{2,30}}) in Source");
                Release(queries);

                var connections = Get(workbook, "Connections");
                foreach (var name in new[] { lookupTable, factTable })
                {
                    Invoke(connections, "Add2", name + "Conn", "test model connection",
                        $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={name}",
                        name, 6, true, false);
                }

                Release(connections);

                var model = Get(workbook, "Model");
                var tables = Get(model, "ModelTables");
                hasBoth = Convert.ToInt32(Get(tables, "Count"), System.Globalization.CultureInfo.InvariantCulture) >= 2;
                Release(tables);
                Release(model);
            }
            catch (Exception exception) when (exception is System.Runtime.InteropServices.COMException or System.Reflection.TargetInvocationException)
            {
                // No Power Query, no Data Model, or policy forbids the connection.
            }
        });
        return hasBoth;
    }

    /// <summary>
    /// A workbook with a Data sheet and an Exhibit sheet whose formulas read it. The two workbooks
    /// in a copy test must hold different numbers, because the value is the only way to tell a
    /// formula bound to this workbook from one still reading the other.
    /// </summary>
    public static void CreateCrossSheetExhibit(string path, double sourceValue, bool withoutExhibit = false) => Create(path, workbook =>
    {
        var sheets = Get(workbook, "Worksheets");
        var data = Item(sheets, 1);
        Set(data, "Name", "Data");
        Set(Get(data, "Range", "A1"), "Value2", sourceValue);
        Set(Get(data, "Range", "A2"), "Value2", sourceValue * 2);

        if (!withoutExhibit)
        {
            var exhibit = Invoke(sheets, "Add")!;
            Set(exhibit, "Name", "Exhibit");
            Set(Get(exhibit, "Range", "A1"), "Formula", "=Data!A1*2");
            Set(Get(exhibit, "Range", "A2"), "Formula", "=SUM(Data!A1:A2)");
            Release(exhibit);
        }

        Release(data);
        Release(sheets);
    });

    /// <summary>How many other workbooks this one links to. Zero is the point of the copy rebind.</summary>
    public static int CountExternalLinks(string path)
    {
        using var application = TestExcelApplication.Start();
        var workbooks = Get(application.Value, "Workbooks");
        var workbook = Invoke(workbooks, "Open", Path.GetFullPath(path))!;
        try
        {
            var links = Invoke(workbook, "LinkSources", 1);
            return links is Array array ? array.Length : 0;
        }
        finally
        {
            Invoke(workbook, "Close", false);
            Release(workbook);
            Release(workbooks);
        }
    }

    public static void CreateMacroTarget(string path, string componentName, string source)
    {
        using var application = TestExcelApplication.Start();
        object? workbook = null;
        try
        {
            var workbooks = Get(application.Value, "Workbooks");
            try { workbook = Invoke(workbooks, "Add"); }
            finally { Release(workbooks); }
            Invoke(workbook, "SaveAs", path, 52);
            var project = Get(workbook, "VBProject");
            var components = Get(project, "VBComponents");
            var component = Invoke(components, "Add", 1);
            var module = Get(component, "CodeModule");
            Set(component, "Name", componentName);
            Invoke(module, "AddFromString", source.Replace("\n", "\r\n", StringComparison.Ordinal));
            Invoke(workbook, "Save");
            Release(module);
            Release(component);
            Release(components);
            Release(project);
        }
        finally
        {
            if (workbook is not null) Invoke(workbook, "Close", false);
            Release(workbook);
        }
    }

    /// <summary>Reads the whole module so a test can prove no generated helper was left behind.</summary>
    public static string ReadModuleText(string path, string componentName)
    {
        using var application = TestExcelApplication.Start();
        object? workbook = null;
        try
        {
            var workbooks = Get(application.Value, "Workbooks");
            try { workbook = Invoke(workbooks, "Open", path, 0, true); }
            finally { Release(workbooks); }
            var project = Get(workbook, "VBProject");
            var components = Get(project, "VBComponents");
            var component = Item(components, componentName);
            var module = Get(component, "CodeModule");
            var count = Convert.ToInt32(Get(module, "CountOfLines"), CultureInfo.InvariantCulture);
            var text = count == 0 ? string.Empty : (string)Get(module, "Lines", 1, count);
            Release(module);
            Release(component);
            Release(components);
            Release(project);
            return text.Replace("\r\n", "\n", StringComparison.Ordinal).Replace('\r', '\n');
        }
        finally
        {
            if (workbook is not null) Invoke(workbook, "Close", false);
            Release(workbook);
        }
    }

    public static string ReadMacroProcedure(string path, string componentName, string procedureName)
    {
        using var application = TestExcelApplication.Start();
        object? workbook = null;
        try
        {
            var workbooks = Get(application.Value, "Workbooks");
            try { workbook = Invoke(workbooks, "Open", path, 0, true); }
            finally { Release(workbooks); }
            var project = Get(workbook, "VBProject");
            var components = Get(project, "VBComponents");
            // VBIDE exposes Item as a method, unlike Excel's own collections.
            var component = Item(components, componentName);
            var module = Get(component, "CodeModule");
            // Parameterized VBIDE properties; they are not callable with InvokeMethod binding.
            var start = Convert.ToInt32(Get(module, "ProcStartLine", procedureName, 0), CultureInfo.InvariantCulture);
            var count = Convert.ToInt32(Get(module, "ProcCountLines", procedureName, 0), CultureInfo.InvariantCulture);
            var source = (string)Get(module, "Lines", start, count);
            Release(module);
            Release(component);
            Release(components);
            Release(project);
            return source.Replace("\r\n", "\n", StringComparison.Ordinal).Replace('\r', '\n').TrimEnd('\n');
        }
        finally
        {
            if (workbook is not null) Invoke(workbook, "Close", false);
            Release(workbook);
        }
    }

    public static OpenUserWorkbook OpenAsUser(string path)
    {
        var application = TestExcelApplication.Start();
        try
        {
            var workbooks = Get(application.Value, "Workbooks");
            try { return new OpenUserWorkbook(application, Invoke(workbooks, "Open", path, 0, false)); }
            finally { Release(workbooks); }
        }
        catch
        {
            application.Dispose();
            throw;
        }
    }

    public static bool HasExpectedSheetAndRepair(string path)
    {
        using var application = TestExcelApplication.Start();
        object? workbook = null;
        try
        {
            var workbooks = Get(application.Value, "Workbooks");
            try { workbook = Invoke(workbooks, "Open", path, 0, true); }
            finally { Release(workbooks); }
            var sheets = Get(workbook, "Worksheets");
            var sheet = Item(sheets, "Imported");
            var range = Get(sheet, "Range", "A1:A3");
            var formulas = Get(range, "FormulaR1C1") as Array;
            var expected = formulas is not null &&
                           Enumerable.Range(1, 3).All(row => string.Equals(formulas.GetValue(row, 1) as string, "=ROW()", StringComparison.Ordinal));
            Release(range);
            Release(sheet);
            Release(sheets);
            return expected;
        }
        finally
        {
            if (workbook is not null) Invoke(workbook, "Close", false);
            Release(workbook);
        }
    }

    public static bool HasFormula(string path, string range, string expected)
    {
        using var application = TestExcelApplication.Start();
        object? workbook = null;
        try
        {
            var workbooks = Get(application.Value, "Workbooks");
            try { workbook = Invoke(workbooks, "Open", path, 0, true); }
            finally { Release(workbooks); }
            var sheets = Get(workbook, "Worksheets");
            var sheet = Item(sheets, 1);
            var target = Get(sheet, "Range", range);
            var actual = Get(target, "FormulaR1C1") as string;
            Release(target);
            Release(sheet);
            Release(sheets);
            return string.Equals(actual, expected, StringComparison.Ordinal);
        }
        finally
        {
            if (workbook is not null) Invoke(workbook, "Close", false);
            Release(workbook);
        }
    }

    /// <summary>The range's number format off disk, or null when its cells do not share one.</summary>
    public static string? ReadNumberFormat(string path, string range)
    {
        using var application = TestExcelApplication.Start();
        object? workbook = null;
        try
        {
            var workbooks = Get(application.Value, "Workbooks");
            try { workbook = Invoke(workbooks, "Open", path, 0, true); }
            finally { Release(workbooks); }
            var sheets = Get(workbook, "Worksheets");
            var sheet = Item(sheets, 1);
            var target = Get(sheet, "Range", range);
            var format = ComAccess.GetOrNull(target, "NumberFormat") as string;
            Release(target);
            Release(sheet);
            Release(sheets);
            return format;
        }
        finally
        {
            if (workbook is not null) Invoke(workbook, "Close", false);
            Release(workbook);
        }
    }

    public static bool HasValue(string path, string range, object expected, string? worksheet = null)
    {
        using var application = TestExcelApplication.Start();
        object? workbook = null;
        try
        {
            var workbooks = Get(application.Value, "Workbooks");
            try { workbook = Invoke(workbooks, "Open", path, 0, true); }
            finally { Release(workbooks); }
            var sheets = Get(workbook, "Worksheets");
            var sheet = Item(sheets, worksheet ?? (object)1);
            var target = Get(sheet, "Range", range);
            var actual = Get(target, "Value2");
            Release(target);
            Release(sheet);
            Release(sheets);
            return Equals(actual, expected);
        }
        finally
        {
            if (workbook is not null) Invoke(workbook, "Close", false);
            Release(workbook);
        }
    }

    public static bool HasExpectedCells(
        string path,
        IReadOnlyDictionary<string, string> expectedFormulas,
        IReadOnlyDictionary<string, object>? expectedValues = null)
    {
        using var application = TestExcelApplication.Start();
        object? workbook = null;
        try
        {
            var workbooks = Get(application.Value, "Workbooks");
            try { workbook = Invoke(workbooks, "Open", path, 0, true); }
            finally { Release(workbooks); }
            var sheets = Get(workbook, "Worksheets");
            var sheet = Item(sheets, 1);
            try
            {
                foreach (var (range, expectedFormula) in expectedFormulas)
                {
                    var target = Get(sheet, "Range", range);
                    try
                    {
                        if (!string.Equals(Get(target, "FormulaR1C1") as string, expectedFormula, StringComparison.Ordinal)) return false;
                    }
                    finally { Release(target); }
                }

                if (expectedValues is not null)
                {
                    foreach (var (range, expectedValue) in expectedValues)
                    {
                        var target = Get(sheet, "Range", range);
                        try
                        {
                            if (!Equals(Get(target, "Value2"), expectedValue)) return false;
                        }
                        finally { Release(target); }
                    }
                }

                return true;
            }
            finally
            {
                Release(sheet);
                Release(sheets);
            }
        }
        finally
        {
            if (workbook is not null) Invoke(workbook, "Close", false);
            Release(workbook);
        }
    }

    private static void Create(string path, Action<object>? configure)
    {
        using var application = TestExcelApplication.Start();
        object? workbook = null;
        try
        {
            var workbooks = Get(application.Value, "Workbooks");
            try { workbook = Invoke(workbooks, "Add"); }
            finally { Release(workbooks); }
            configure?.Invoke(workbook);
            Invoke(workbook, "SaveAs", path);
        }
        finally
        {
            if (workbook is not null) Invoke(workbook, "Close", false);
            Release(workbook);
        }
    }

    internal sealed class OpenUserWorkbook(TestExcelApplication application, object workbook) : IDisposable
    {
        private readonly TestExcelApplication _application = application;
        private readonly object _workbook = workbook;
        private bool _disposed;

        public bool IsApplicationRunning => _application.IsRunning;

        public bool IsWorkbookOpen
        {
            get
            {
                try
                {
                    _ = Get(_workbook, "Name");
                    return true;
                }
                catch (COMException) { return false; }
            }
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            try { Invoke(_workbook, "Close", false); }
            finally
            {
                Release(_workbook);
                _application.Dispose();
            }
        }
    }

    internal sealed class TestExcelApplication : IDisposable
    {
        private readonly OwnedExcelProcess _ownedProcess;
        private bool _disposed;

        private TestExcelApplication(object value, OwnedExcelProcess ownedProcess)
        {
            Value = value;
            _ownedProcess = ownedProcess;
        }

        public object Value { get; }

        public bool IsRunning => _ownedProcess.IsRunning;

        public static TestExcelApplication Start()
        {
            var beforeStart = OwnedExcelProcess.SnapshotExcelProcesses();
            var type = Type.GetTypeFromProgID("Excel.Application", throwOnError: true)!;
            var app = Activator.CreateInstance(type)!;
            try
            {
                var process = OwnedExcelProcess.CaptureNew(app, beforeStart);
                RecordFixtureProcess(process.Identity);
                Set(app, "Visible", false);
                Set(app, "DisplayAlerts", false);
                return new TestExcelApplication(app, process);
            }
            catch
            {
                if (OwnedExcelProcess.IsNewlyOwned(app, beforeStart)) Invoke(app, "Quit");
                Release(app);
                throw;
            }
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            try { Invoke(Value, "Quit"); }
            finally
            {
                Release(Value);
                _ownedProcess.WaitForExitOrTerminate();
            }
        }
    }

    private static readonly Lock FixtureProcessGate = new();
    private static readonly HashSet<ProcessIdentity> FixtureProcesses = [];

    /// <summary>
    /// Captures the "before" set once the process table has stopped moving.
    ///
    /// Excel exits asynchronously, so an instance from the previous test can still be dying when
    /// the next one starts. Snapshotting immediately omits it from that test's baseline, and it
    /// then surfaces as a leak attributed to a test that never created it. These tests are already
    /// serial; this makes their accounting serial too. Two identical readings in a row mean nothing
    /// is mid-exit, and the wait is bounded so a genuinely busy machine still proceeds.
    /// </summary>
    public static HashSet<ProcessIdentity> SnapshotSettledExcel()
    {
        var previous = OwnedExcelProcess.SnapshotExcelProcesses();
        for (var attempt = 0; attempt < 20; attempt++)
        {
            Thread.Sleep(250);
            var current = OwnedExcelProcess.SnapshotExcelProcesses();
            if (current.SetEquals(previous)) return current;
            previous = current;
        }

        return previous;
    }

    private static void RecordFixtureProcess(ProcessIdentity identity)
    {
        lock (FixtureProcessGate) FixtureProcesses.Add(identity);
    }

    /// <summary>
    /// Asserts the product left no Excel process behind.
    ///
    /// Two things are excluded, and only these two: processes that existed before the test, and
    /// processes this fixture started itself. The second exclusion is the important one - a test
    /// opens Excel several times of its own accord to verify results, and counting those as product
    /// leaks measures the harness rather than the product. That is exactly the false positive a
    /// field report produced when both servers appeared to strand a process on the same workflow.
    ///
    /// Identity, not process id: the fixture's own processes are matched on id, start time, and
    /// image path, so a recycled id can never launder a genuine leak into an exclusion.
    ///
    /// A brief settle first, because Excel exits asynchronously and can linger a beat after Quit
    /// returns. A real leak never clears, so waiting cannot hide one.
    /// </summary>
    public static void AssertNoLeakedExcel(ISet<ProcessIdentity> existingExcel)
    {
        ProcessIdentity[] leaked = [];
        // Thirty seconds, because teardown time scales with what the operation did: an audit that
        // walked every worksheet, table and defined name takes measurably longer to exit than a
        // three-cell repair. The bound still holds the assertion honest - a genuine leak never
        // clears, so a longer wait cannot turn a failure into a pass, only a slow exit into one.
        for (var attempt = 0; attempt < 120; attempt++)
        {
            HashSet<ProcessIdentity> fixtureOwned;
            lock (FixtureProcessGate) fixtureOwned = [.. FixtureProcesses];

            leaked = [.. OwnedExcelProcess.SnapshotExcelProcesses()
                .Where(process => !existingExcel.Contains(process) && !fixtureOwned.Contains(process))];
            if (leaked.Length == 0) return;
            Thread.Sleep(250);
        }

        Assert.Empty(leaked);
    }

    // The product's own late-bound rules, so a fixture cannot bind a member differently from the
    // runtime it exists to exercise - which is how three binding defects reached a release.
    private static object Get(object target, string member, params object?[] arguments) => ComAccess.Get(target, member, arguments);

    private static void Set(object target, string member, object? value) => ComAccess.Set(target, member, value);

    private static object Invoke(object target, string member, params object?[] arguments) => ComAccess.Invoke(target, member, arguments)!;

    private static object Item(object collection, object index) => ComAccess.Item(collection, index);

    private static void Release(object? value) => ComAccess.Release(value);
}
