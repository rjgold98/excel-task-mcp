using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;

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

    private readonly List<int> _ownedProcesses = [];

    public IReadOnlyList<int> OwnedProcesses => _ownedProcesses;

    public static int[] SnapshotExcelProcesses()
    {
        var processes = Process.GetProcessesByName("EXCEL");
        try
        {
            return [.. processes.Select(process => process.Id)];
        }
        finally
        {
            foreach (var process in processes) process.Dispose();
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
            var referenceSheet = Get(referenceSheets, "Item", 1);
            Set(referenceSheet, "Name", "Reference");
            Set(Get(referenceSheet, "Range", "A1"), "Formula", "=ROW()");
            Set(Get(referenceSheet, "Range", "A3"), "Formula", "=ROW()");
            Invoke(reference, "SaveAs", referencePath);
            Invoke(reference, "Close", false);

            var target = Invoke(workbooks, "Add")!;
            var sheets = Get(target, "Worksheets");
            var model = Get(sheets, "Item", 1);
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

    public void TerminateOwnedProcesses()
    {
        foreach (var id in _ownedProcesses)
        {
            try
            {
                using var process = Process.GetProcessById(id);
                if (!process.HasExited) process.Kill();
            }
            catch (ArgumentException) { }
            catch (InvalidOperationException) { }
        }
    }

    private object Start()
    {
        var before = SnapshotExcelProcesses();
        var type = Type.GetTypeFromProgID("Excel.Application", throwOnError: true)
            ?? throw new InvalidOperationException("Microsoft Excel is not registered on this machine.");
        var application = Activator.CreateInstance(type)
            ?? throw new InvalidOperationException("Microsoft Excel could not be started.");
        Set(application, "Visible", false);
        Set(application, "DisplayAlerts", false);
        foreach (var id in SnapshotExcelProcesses().Where(id => !before.Contains(id)))
        {
            if (!_ownedProcesses.Contains(id)) _ownedProcesses.Add(id);
        }

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

    private static object Get(object target, string member, params object?[] arguments) =>
        target.GetType().InvokeMember(member, BindingFlags.GetProperty, null, target, arguments, CultureInfo.InvariantCulture)
        ?? throw new InvalidOperationException($"Excel did not return '{member}'.");

    private static void Set(object target, string member, object? value) =>
        target.GetType().InvokeMember(member, BindingFlags.SetProperty, null, target, [value], CultureInfo.InvariantCulture);

    private static object? Invoke(object target, string member, params object?[] arguments) =>
        target.GetType().InvokeMember(member, BindingFlags.InvokeMethod, null, target, arguments, CultureInfo.InvariantCulture);
}
