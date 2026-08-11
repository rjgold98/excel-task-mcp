using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using ExcelTask.Excel;

namespace ExcelTask.McpServer.Tests;

internal static class ExcelFixtureWorkbook
{
    public static void CreateTarget(string path, ICollection<ExcelProcessIdentity> ownedProcesses) =>
        Create(path, ownedProcesses, null);

    public static void CreateReference(string path, ICollection<ExcelProcessIdentity> ownedProcesses) =>
        Create(path, ownedProcesses, workbook =>
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

    public static void CreateMacroTarget(string path, string componentName, string source, ICollection<ExcelProcessIdentity> ownedProcesses)
    {
        using var application = FixtureExcelApplication.Start(ownedProcesses);
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

    public static string ReadMacroProcedure(string path, string componentName, string procedureName, ICollection<ExcelProcessIdentity> ownedProcesses)
    {
        using var application = FixtureExcelApplication.Start(ownedProcesses);
        object? workbook = null;
        try
        {
            var workbooks = Get(application.Value, "Workbooks");
            try { workbook = Invoke(workbooks, "Open", path, 0, true); }
            finally { Release(workbooks); }
            var project = Get(workbook, "VBProject");
            var components = Get(project, "VBComponents");
            // VBIDE exposes Item as a method and its line accessors as parameterized properties,
            // the opposite of Excel's own collections.
            var component = Item(components, componentName);
            var module = Get(component, "CodeModule");
            var start = Convert.ToInt32(Get(module, "ProcStartLine", procedureName, 0), CultureInfo.InvariantCulture);
            var count = Convert.ToInt32(Get(module, "ProcCountLines", procedureName, 0), CultureInfo.InvariantCulture);
            var text = (string)Get(module, "Lines", start, count);
            Release(module);
            Release(component);
            Release(components);
            Release(project);
            return text.Replace("\r\n", "\n", StringComparison.Ordinal).Replace('\r', '\n').TrimEnd('\n');
        }
        finally
        {
            if (workbook is not null) Invoke(workbook, "Close", false);
            Release(workbook);
        }
    }

    public static bool HasExpectedSheetAndRepair(string path, ICollection<ExcelProcessIdentity> ownedProcesses)
    {
        using var application = FixtureExcelApplication.Start(ownedProcesses);
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

    private static void Create(string path, ICollection<ExcelProcessIdentity> ownedProcesses, Action<object>? configure)
    {
        using var application = FixtureExcelApplication.Start(ownedProcesses);
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

    // The product's own late-bound rules, so a fixture cannot bind a member differently from the
    // runtime it exists to exercise.
    private static object Get(object target, string member, params object?[] arguments) => ComAccess.Get(target, member, arguments);

    private static void Set(object target, string member, object? value) => ComAccess.Set(target, member, value);

    private static object Invoke(object target, string member, params object?[] arguments) => ComAccess.Invoke(target, member, arguments)!;

    private static object Item(object collection, object index) => ComAccess.Item(collection, index);

    private static void Release(object? value) => ComAccess.Release(value);

    private sealed class FixtureExcelApplication : IDisposable
    {
        private readonly ExcelProcessIdentity _process;
        private bool _disposed;

        private FixtureExcelApplication(object value, ExcelProcessIdentity process)
        {
            Value = value;
            _process = process;
        }

        public object Value { get; }

        public static FixtureExcelApplication Start(ICollection<ExcelProcessIdentity> ownedProcesses)
        {
            var beforeStart = ExcelProcessIdentity.SnapshotExcelProcesses();
            var type = Type.GetTypeFromProgID("Excel.Application", throwOnError: true)!;
            var application = Activator.CreateInstance(type)!;
            ExcelProcessIdentity? ownedProcess = null;
            try
            {
                var process = ExcelProcessIdentity.CaptureApplication(application);
                if (beforeStart.Contains(process))
                {
                    throw new InvalidOperationException("Excel activation did not create a new fixture-owned process.");
                }

                ownedProcess = process;
                ownedProcesses.Add(process);
                Set(application, "Visible", false);
                Set(application, "DisplayAlerts", false);
                return new FixtureExcelApplication(application, process);
            }
            catch
            {
                if (ownedProcess is not null && ExcelProcessIdentity.TryOpenMatching(ownedProcess, out var owned))
                {
                    using (owned) Invoke(application, "Quit");
                }

                Release(application);
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
                _process.WaitForExitOrTerminate();
            }
        }
    }
}

/// <summary>
/// Process identity for fixtures, standing on the product's own <see cref="ProcessIdentity"/> rather
/// than a second implementation of it. This was a near-verbatim copy of the product's triple-check -
/// process id, start time, and image path - which meant a test could prove a weaker fact about
/// process ownership than the runtime it was testing.
/// </summary>
internal sealed record ExcelProcessIdentity(ProcessIdentity Identity)
{
    public static HashSet<ExcelProcessIdentity> SnapshotExcelProcesses() =>
        [.. OwnedExcelProcess.SnapshotExcelProcesses().Select(identity => new ExcelProcessIdentity(identity))];

    public static ExcelProcessIdentity CaptureApplication(object application)
    {
        var hwnd = new IntPtr(Convert.ToInt64(ComAccess.Get(application, "Hwnd"), CultureInfo.InvariantCulture));
        _ = GetWindowThreadProcessId(hwnd, out var processId);
        if (processId == 0) throw new InvalidOperationException("Excel fixture process identity could not be captured.");
        return Capture((int)processId);
    }

    public static ExcelProcessIdentity Capture(int processId) => new(ProcessIdentity.Capture(processId));

    /// <summary>
    /// Snapshots once the process table has stopped moving, so a dying Excel from the previous test
    /// is not omitted from this one's baseline and then reported as this one's leak.
    /// </summary>
    public static HashSet<ExcelProcessIdentity> SnapshotSettledExcel()
    {
        var previous = SnapshotExcelProcesses();
        for (var attempt = 0; attempt < 20; attempt++)
        {
            Thread.Sleep(250);
            var current = SnapshotExcelProcesses();
            if (current.SetEquals(previous)) return current;
            previous = current;
        }

        return previous;
    }

    /// <summary>
    /// Asserts the product left no Excel process behind, waiting for one to clear.
    ///
    /// The Excel-project fixture has done this since 0.9.0 and this one did not: it snapshotted
    /// once, the instant the test body returned. Excel exits asynchronously, so an instance the
    /// product had already told to quit could still be alive for that single reading - and the whole
    /// suite run under `dotnet test` on the solution runs both Excel assemblies at once, which makes
    /// the other assembly's fixtures look like this one's leaks. Four tests failed that way and none
    /// of them had leaked anything. `scripts/Test-Mvp.ps1` serializes the two projects for exactly
    /// this reason; the assertion should not depend on the caller remembering to.
    ///
    /// A genuine leak never clears, so waiting cannot turn a failure into a pass.
    /// </summary>
    public static void AssertNoLeakedExcel(ISet<ExcelProcessIdentity> existingExcel)
    {
        ExcelProcessIdentity[] leaked = [];
        for (var attempt = 0; attempt < 120; attempt++)
        {
            leaked = [.. SnapshotExcelProcesses().Where(process => !existingExcel.Contains(process))];
            if (leaked.Length == 0) return;
            Thread.Sleep(250);
        }

        Assert.Empty(leaked);
    }

    public static bool TryOpenMatching(ExcelProcessIdentity identity, out Process process) =>
        ProcessIdentity.TryOpenMatching(identity.Identity, out process);

    public bool IsRunning => TryOpenMatching(this, out var process) && DisposeProcess(process);

    public void WaitForExitOrTerminate()
    {
        if (!TryOpenMatching(this, out var process)) return;
        using (process)
        {
            if (process.WaitForExit(10_000)) return;
            if (!TryOpenMatching(this, out var stillMatching)) return;
            using (stillMatching)
            {
                stillMatching.Kill(entireProcessTree: false);
                _ = stillMatching.WaitForExit(5_000);
            }
        }
    }

    private static bool DisposeProcess(Process process)
    {
        process.Dispose();
        return true;
    }

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
}
