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
        var sheet = Get(sheets, "Item", 1);
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
        var sheet = Get(sheets, "Item", 1);
        Set(sheet, "Name", "Reference");
        var range = Get(sheet, "Range", "A1:A3");
        Set(range, "FormulaR1C1", new object?[,] { { "=ROW()" }, { null }, { "=ROW()" } });
        Release(range);
        Release(sheet);
        Release(sheets);
    });

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
            var component = Invoke(components, "Item", componentName);
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
            var component = Invoke(components, "Item", componentName);
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
            var sheet = Get(sheets, "Item", "Imported");
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
            var sheet = Get(sheets, "Item", 1);
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

    public static bool HasValue(string path, string range, object expected)
    {
        using var application = TestExcelApplication.Start();
        object? workbook = null;
        try
        {
            var workbooks = Get(application.Value, "Workbooks");
            try { workbook = Invoke(workbooks, "Open", path, 0, true); }
            finally { Release(workbooks); }
            var sheets = Get(workbook, "Worksheets");
            var sheet = Get(sheets, "Item", 1);
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
            var sheet = Get(sheets, "Item", 1);
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

    private static object Get(object target, string member, params object?[] arguments) => target.GetType().InvokeMember(member, BindingFlags.GetProperty, null, target, arguments, CultureInfo.InvariantCulture)!;

    private static void Set(object target, string member, object? value) => target.GetType().InvokeMember(member, BindingFlags.SetProperty, null, target, [value], CultureInfo.InvariantCulture);

    private static object Invoke(object target, string member, params object?[] arguments) => target.GetType().InvokeMember(member, BindingFlags.InvokeMethod, null, target, arguments, CultureInfo.InvariantCulture)!;

    private static void Release(object? value)
    {
        if (value is not null && Marshal.IsComObject(value)) Marshal.FinalReleaseComObject(value);
    }
}
