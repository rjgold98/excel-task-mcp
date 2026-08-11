using System.Text.Json;
using ExcelTask.Core;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;

namespace ExcelTask.McpServer.Tests;

[CollectionDefinition("ExcelTask COM serial", DisableParallelization = true)]
public sealed class ExcelTaskComSerialFixture;

[Collection("ExcelTask COM serial")]
[Trait("RunType", "OnDemand")]
public sealed class ExcelTaskRealExcelOnDemandTests
{
    [Fact]
    public async Task ReadReturnsCellContentsThroughTheRealToolBoundaryAndReleasesOwnedExcel()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        var target = Path.Combine(directory, "read.xlsx");
        var existingExcel = ExcelProcessIdentity.SnapshotExcelProcesses();
        var fixtureProcesses = new List<ExcelProcessIdentity>();
        Directory.CreateDirectory(directory);

        try
        {
            // The reference fixture is used as the target here because it has content: =ROW() in
            // A1 and A3 with A2 left blank, which is what makes the blank-omission observable.
            ExcelFixtureWorkbook.CreateReference(target, fixtureProcesses);
            var stampBefore = (new FileInfo(target).Length, new FileInfo(target).LastWriteTimeUtc);

            var transport = new StdioClientTransport(new StdioClientTransportOptions
            {
                Name = "ExcelTask-real-read-test",
                Command = TestServer.ServerPath,
                WorkingDirectory = directory,
                InheritEnvironmentVariables = false,
                EnvironmentVariables = TestServer.EnvironmentVariables(),
                ShutdownTimeout = TimeSpan.FromSeconds(2)
            });
            await using var client = await McpClient.CreateAsync(
                transport,
                new McpClientOptions { ClientInfo = new() { Name = "ExcelTask-RealExcel-TestClient", Version = "1.0.0" } });

            var result = await client.CallToolAsync(
                "excel_task",
                new Dictionary<string, object?>
                {
                    ["request"] = new ExcelTaskRequest(
                        target,
                        new ExcelOperation(
                            ExcelOperationKind.ReadWorksheetRange,
                            ReadWorksheetRange: new ReadWorksheetRangeOperation("Reference", "A1:C3")),
                        ExcelTaskMode.Apply,
                        WorkbookBinding.Isolated,
                        SaveMode.Same,
                        null,
                        OverwriteConfirmed: false)
                });

            Assert.False(result.IsError);
            var receipt = result.StructuredContent!.Value;
            Assert.Equal(nameof(ExcelTaskStatus.Completed), receipt.GetProperty("status").GetString());

            // Reading is the one thing this server does that returns workbook contents, so the
            // whole point is that they survive the worker protocol and both bounding seams intact.
            var range = receipt.GetProperty("range");
            Assert.Equal(9, range.GetProperty("cellsInRange").GetInt32());
            Assert.Equal(2, range.GetProperty("nonEmptyCells").GetInt32());
            Assert.Equal(2, range.GetProperty("cells").GetArrayLength());
            Assert.False(range.GetProperty("truncated").GetBoolean());
            var addresses = range.GetProperty("cells").EnumerateArray()
                .Select(cell => cell.GetProperty("address").GetString() ?? string.Empty)
                .ToArray();
            Assert.Equal(["A1", "A3"], addresses);

            // Reading asks for no overwrite authorization, because it never writes - and proves it.
            Assert.False(receipt.GetProperty("confirmation").GetProperty("required").GetBoolean());
            Assert.Equal(stampBefore, (new FileInfo(target).Length, new FileInfo(target).LastWriteTimeUtc));
        }
        finally
        {
            Assert.All(fixtureProcesses, process => Assert.False(process.IsRunning));
            var remainingExcel = ExcelProcessIdentity.SnapshotExcelProcesses();
            Assert.DoesNotContain(remainingExcel, process => !existingExcel.Contains(process));
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public async Task ApplyIsolatedCopyPersistsWorkbookAndReleasesOwnedExcel()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        var target = Path.Combine(directory, "target.xlsx");
        var reference = Path.Combine(directory, "reference.xlsx");
        var output = Path.Combine(directory, "output.xlsx");
        var existingExcel = ExcelProcessIdentity.SnapshotExcelProcesses();
        var fixtureProcesses = new List<ExcelProcessIdentity>();
        Directory.CreateDirectory(directory);

        try
        {
            ExcelFixtureWorkbook.CreateTarget(target, fixtureProcesses);
            ExcelFixtureWorkbook.CreateReference(reference, fixtureProcesses);

            var environment = TestServer.EnvironmentVariables();
            var transport = new StdioClientTransport(new StdioClientTransportOptions
            {
                Name = "ExcelTask-real-worker-test",
                Command = TestServer.ServerPath,
                WorkingDirectory = directory,
                InheritEnvironmentVariables = false,
                EnvironmentVariables = environment,
                ShutdownTimeout = TimeSpan.FromSeconds(2)
            });
            await using var client = await McpClient.CreateAsync(
                transport,
                new McpClientOptions { ClientInfo = new() { Name = "ExcelTask-RealExcel-TestClient", Version = "1.0.0" } });
            var tools = await client.ListToolsAsync(new ListToolsRequestParams());
            var tool = Assert.Single(tools.Tools);
            Assert.Equal("excel_task", tool.Name);

            var result = await client.CallToolAsync(
                "excel_task",
                new Dictionary<string, object?>
                {
                    ["request"] = new ExcelTaskRequest(
                        target,
                        new ExcelOperation(
                            ExcelOperationKind.CopyExhibit,
                            CopyExhibit: new CopyExhibitOperation(reference, "Reference", "Imported", ["A1:A3"])),
                        ExcelTaskMode.Apply,
                        WorkbookBinding.Isolated,
                        SaveMode.Copy,
                        output,
                        OverwriteConfirmed: false)
                });

            Assert.False(result.IsError);
            var receipt = result.StructuredContent!.Value;
            Assert.Equal(nameof(ExcelTaskStatus.Completed), receipt.GetProperty("status").GetString());
            Assert.Equal(nameof(SaveMode.Copy), receipt.GetProperty("save").GetProperty("mode").GetString());
            Assert.False(receipt.GetProperty("confirmation").GetProperty("required").GetBoolean());
            Assert.True(File.Exists(output));
            Assert.True(ExcelFixtureWorkbook.HasExpectedSheetAndRepair(output, fixtureProcesses));
            using (new FileStream(output, FileMode.Open, FileAccess.ReadWrite, FileShare.None)) { }
        }
        finally
        {
            Assert.All(fixtureProcesses, process => Assert.False(process.IsRunning));
            var remainingExcel = ExcelProcessIdentity.SnapshotExcelProcesses();
            Assert.DoesNotContain(remainingExcel, process => !existingExcel.Contains(process));
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public async Task MacroPlanAndApplyRunThroughTheRealToolBoundaryAndReleaseOwnedExcel()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        var target = Path.Combine(directory, "macro-target.xlsm");
        var output = Path.Combine(directory, "macro-output.xlsm");
        const string component = "SafeModule";
        const string procedure = "WriteMarker";
        const string originalSource = "Public Sub WriteMarker()\n    ThisWorkbook.Worksheets(1).Range(\"A1\").Value2 = \"original\"\nEnd Sub";
        const string replacementSource = "Public Sub WriteMarker()\n    ThisWorkbook.Worksheets(1).Range(\"A1\").Value2 = \"ran\"\nEnd Sub";
        var existingExcel = ExcelProcessIdentity.SnapshotExcelProcesses();
        var fixtureProcesses = new List<ExcelProcessIdentity>();
        Directory.CreateDirectory(directory);

        try
        {
            ExcelFixtureWorkbook.CreateMacroTarget(target, component, originalSource, fixtureProcesses);

            var environment = TestServer.EnvironmentVariables();
            var transport = new StdioClientTransport(new StdioClientTransportOptions
            {
                Name = "ExcelTask-real-macro-test",
                Command = TestServer.ServerPath,
                WorkingDirectory = directory,
                InheritEnvironmentVariables = false,
                EnvironmentVariables = environment,
                ShutdownTimeout = TimeSpan.FromSeconds(2)
            });
            await using var client = await McpClient.CreateAsync(
                transport,
                new McpClientOptions { ClientInfo = new() { Name = "ExcelTask-RealMacro-TestClient", Version = "1.0.0" } });

            var planned = await client.CallToolAsync("excel_task", MacroArguments(target, output, component, procedure, ExcelTaskMode.Plan));
            Assert.False(planned.IsError);
            var planReceipt = planned.StructuredContent!.Value;
            Assert.Equal(nameof(ExcelTaskStatus.Planned), planReceipt.GetProperty("status").GetString());
            var macro = planReceipt.GetProperty("macroProcedure");
            Assert.Equal(originalSource, macro.GetProperty("source").GetString());
            var hash = macro.GetProperty("sha256").GetString()!;
            Assert.False(File.Exists(output));

            var applied = await client.CallToolAsync(
                "excel_task",
                MacroArguments(target, output, component, procedure, ExcelTaskMode.Apply, hash, replacementSource, runAfterEdit: true));
            Assert.False(applied.IsError);
            var applyReceipt = applied.StructuredContent!.Value;
            Assert.Equal(nameof(ExcelTaskStatus.Completed), applyReceipt.GetProperty("status").GetString());
            var appliedMacro = applyReceipt.GetProperty("macroProcedure");
            Assert.True(appliedMacro.GetProperty("runCompleted").GetBoolean());
            // Apply must never echo VBA source back through the tool boundary.
            Assert.Equal(JsonValueKind.Null, appliedMacro.GetProperty("source").ValueKind);

            Assert.True(File.Exists(output));
            Assert.Equal(originalSource, ExcelFixtureWorkbook.ReadMacroProcedure(target, component, procedure, fixtureProcesses));
            Assert.Equal(replacementSource, ExcelFixtureWorkbook.ReadMacroProcedure(output, component, procedure, fixtureProcesses));
            using (new FileStream(output, FileMode.Open, FileAccess.ReadWrite, FileShare.None)) { }
        }
        finally
        {
            Assert.All(fixtureProcesses, process => Assert.False(process.IsRunning));
            var remainingExcel = ExcelProcessIdentity.SnapshotExcelProcesses();
            Assert.DoesNotContain(remainingExcel, process => !existingExcel.Contains(process));
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        }
    }

    private static Dictionary<string, object?> MacroArguments(
        string target,
        string output,
        string component,
        string procedure,
        ExcelTaskMode mode,
        string? expectedHash = null,
        string? replacementSource = null,
        bool runAfterEdit = false) => new()
    {
        ["request"] = new ExcelTaskRequest(
            target,
            new ExcelOperation(
                ExcelOperationKind.EditMacroProcedure,
                EditMacroProcedure: new EditMacroProcedureOperation(component, procedure, expectedHash, replacementSource, runAfterEdit)),
            mode,
            WorkbookBinding.Isolated,
            SaveMode.Copy,
            output,
            OverwriteConfirmed: mode == ExcelTaskMode.Apply)
    };
}
