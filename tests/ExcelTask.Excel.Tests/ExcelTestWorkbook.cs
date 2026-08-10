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
