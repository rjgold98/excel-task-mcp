using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Text;
using System.Text.Json;
using ExcelTask.Core;
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

internal sealed record OperationResult(
    string Label,
    string Status,
    double ElapsedSeconds,
    int LeakedExcel,
    string Summary,
    string Checks,
    string? Error);

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
            await RunOperationsAsync(serverPath, work, macroNote is null, fixtures, operations, notes);
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

        var leaked = FieldCheckFixtures.SnapshotExcelProcesses()
            .Count(id => !preExisting.Contains(id) && !fixtures.OwnedProcesses.Contains(id));
        return WriteReport(outputDirectory, serverPath, environment, surfaces, operations, notes, leaked);
    }

    private static async Task RunOperationsAsync(
        string serverPath,
        string work,
        bool macroReady,
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
        await Task.Delay(700);
        var leaked = FieldCheckFixtures.SnapshotExcelProcesses()
            .Count(id => !before.Contains(id) && !fixtures.OwnedProcesses.Contains(id));

        operations.Add(new OperationResult(label, status, Math.Round(timer.Elapsed.TotalSeconds, 2), leaked, summary, checks, error));
        Write($"      {label,-34} {status,-16} {timer.Elapsed.TotalSeconds,6:F1}s  leaked={leaked}");
        return structured;
    }

    private static Dictionary<string, string> ReadEnvironment(List<string> notes)
    {
        var environment = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["computerName"] = Environment.MachineName,
            ["osVersion"] = Environment.OSVersion.VersionString,
            ["runtime"] = RuntimeInformation.FrameworkDescription,
            ["dotnetRoot"] = Environment.GetEnvironmentVariable("DOTNET_ROOT") ?? "not set",
            ["powerShellLockdownPolicy"] = Environment.GetEnvironmentVariable("__PSLockdownPolicy") ?? "not set"
        };

        foreach (var version in (string[])["16.0", "15.0"])
        {
            using var security = Registry.CurrentUser.OpenSubKey($@"Software\Microsoft\Office\{version}\Excel\Security");
            if (security is not null)
            {
                environment["officeVersion"] = version;
                // 1 means programmatic VBA project access is trusted; without it macro work is impossible.
                environment["accessVBOM"] = security.GetValue("AccessVBOM")?.ToString() ?? "not set";
                // 1 enable all, 2 disable with notification, 3 signed only, 4 disable all.
                environment["vbaWarnings"] = security.GetValue("VBAWarnings")?.ToString() ?? "not set";
            }

            using var policy = Registry.CurrentUser.OpenSubKey($@"Software\Policies\Microsoft\Office\{version}\Excel\Security");
            if (policy is not null)
            {
                environment["policyAccessVBOM"] = policy.GetValue("AccessVBOM")?.ToString() ?? "not set";
                environment["policyVBAWarnings"] = policy.GetValue("VBAWarnings")?.ToString() ?? "not set";
                notes.Add("Group Policy sets Excel macro security on this machine; the organization controls it, not this product.");
            }
        }

        environment["excelRunningBefore"] = FieldCheckFixtures.SnapshotExcelProcesses().Length.ToString(CultureInfo.InvariantCulture);
        try
        {
            var type = Type.GetTypeFromProgID("Excel.Application", throwOnError: true)!;
            var application = Activator.CreateInstance(type)!;
            try
            {
                environment["excelVersion"] = Read(application, "Version");
                environment["excelBuild"] = Read(application, "Build");
                var addins = new List<string>();
                var comAddins = application.GetType().InvokeMember("COMAddIns", BindingFlags.GetProperty, null, application, null, CultureInfo.InvariantCulture);
                if (comAddins is not null)
                {
                    var count = Convert.ToInt32(comAddins.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, comAddins, null, CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
                    for (var index = 1; index <= count; index++)
                    {
                        var addin = comAddins.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, comAddins, [index], CultureInfo.InvariantCulture);
                        if (addin is null) continue;
                        if (Convert.ToBoolean(addin.GetType().InvokeMember("Connect", BindingFlags.GetProperty, null, addin, null, CultureInfo.InvariantCulture), CultureInfo.InvariantCulture))
                        {
                            addins.Add(Read(addin, "ProgId"));
                        }
                    }
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
            environment["excelVersion"] = $"COM probe failed: {exception.Message.ReplaceLineEndings(" ").Trim()}";
        }

        return environment;

        static string Read(object target, string member) =>
            target.GetType().InvokeMember(member, BindingFlags.GetProperty, null, target, null, CultureInfo.InvariantCulture)?.ToString() ?? "";
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
            serverPath,
            environment,
            toolSurfaces = surfaces,
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

        Console.WriteLine();
        Write("Report written:");
        Write("  " + markdownPath);
        Write("  " + jsonPath);

        var failed = operations.Count(operation => operation.Status is not ("Completed" or "Planned"));
        if (failed > 0 || leaked > 0 || operations.Count == 0)
        {
            Write("Some operations did not complete, or an Excel process was left behind. Send both files back.");
            return 1;
        }

        Write("All operations completed and no Excel process was left behind.");
        return 0;
    }

    private static void Write(string message) => Console.WriteLine(message);
}
