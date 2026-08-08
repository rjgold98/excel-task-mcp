using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;

namespace ExcelTask.McpServer.Tests;

internal static class ExcelFixtureWorkbook
{
    public static void CreateTarget(string path, ICollection<ExcelProcessIdentity> ownedProcesses) =>
        Create(path, ownedProcesses, null);

    public static void CreateReference(string path, ICollection<ExcelProcessIdentity> ownedProcesses) =>
        Create(path, ownedProcesses, workbook =>
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

    private static object Get(object target, string member, params object?[] arguments) =>
        target.GetType().InvokeMember(member, BindingFlags.GetProperty, null, target, arguments, CultureInfo.InvariantCulture)!;

    private static void Set(object target, string member, object? value) =>
        target.GetType().InvokeMember(member, BindingFlags.SetProperty, null, target, [value], CultureInfo.InvariantCulture);

    private static object Invoke(object target, string member, params object?[] arguments) =>
        target.GetType().InvokeMember(member, BindingFlags.InvokeMethod, null, target, arguments, CultureInfo.InvariantCulture)!;

    private static void Release(object? value)
    {
        if (value is not null && Marshal.IsComObject(value)) Marshal.FinalReleaseComObject(value);
    }

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

internal sealed record ExcelProcessIdentity(int ProcessId, DateTime StartTimeUtc, string ExecutablePath)
{
    public static HashSet<ExcelProcessIdentity> SnapshotExcelProcesses()
    {
        var identities = new HashSet<ExcelProcessIdentity>();
        foreach (var process in Process.GetProcessesByName("EXCEL"))
        {
            using (process)
            {
                try { identities.Add(Capture(process.Id)); }
                catch (InvalidOperationException) { }
                catch (System.ComponentModel.Win32Exception) { }
            }
        }

        return identities;
    }

    public static ExcelProcessIdentity CaptureApplication(object application)
    {
        var hwndValue = application.GetType().InvokeMember("Hwnd", BindingFlags.GetProperty, null, application, null, CultureInfo.InvariantCulture);
        var hwnd = new IntPtr(Convert.ToInt64(hwndValue, CultureInfo.InvariantCulture));
        _ = GetWindowThreadProcessId(hwnd, out var processId);
        if (processId == 0) throw new InvalidOperationException("Excel fixture process identity could not be captured.");
        return Capture((int)processId);
    }

    public static ExcelProcessIdentity Capture(int processId)
    {
        using var process = Process.GetProcessById(processId);
        return new ExcelProcessIdentity(processId, process.StartTime.ToUniversalTime(), GetExecutablePath(process));
    }

    public static bool TryOpenMatching(ExcelProcessIdentity identity, out Process process)
    {
        process = null!;
        try
        {
            var candidate = Process.GetProcessById(identity.ProcessId);
            if (candidate.StartTime.ToUniversalTime() != identity.StartTimeUtc ||
                !string.Equals(GetExecutablePath(candidate), identity.ExecutablePath, StringComparison.OrdinalIgnoreCase))
            {
                candidate.Dispose();
                return false;
            }

            process = candidate;
            return true;
        }
        catch (ArgumentException) { return false; }
        catch (InvalidOperationException) { return false; }
    }

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

    private static string GetExecutablePath(Process process) => process.MainModule?.FileName
        ?? throw new InvalidOperationException("Excel executable path could not be captured.");

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
}
