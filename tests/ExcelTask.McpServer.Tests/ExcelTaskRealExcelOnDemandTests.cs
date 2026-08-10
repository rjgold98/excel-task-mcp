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

            var environment = StdioClientTransportOptions.GetDefaultEnvironmentVariables();
            var transport = new StdioClientTransport(new StdioClientTransportOptions
            {
                Name = "ExcelTask-real-worker-test",
                Command = Path.Combine(AppContext.BaseDirectory, "excel-task-mcp.exe"),
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

}
