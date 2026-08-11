using ExcelTask.Core;
using ModelContextProtocol.Client;
using System.Diagnostics;
using System.Text.Json;

namespace ExcelTask.McpServer.Tests;

public sealed class PrivateWorkerMcpProcessTests
{
    private static readonly JsonSerializerOptions WorkerJsonOptions = new(JsonSerializerDefaults.Web)
    {
        Converters = { new System.Text.Json.Serialization.JsonStringEnumConverter(JsonNamingPolicy.CamelCase) }
    };

    [Fact]
    public void ChildEnvironmentCarriesTheHostingRuntimeRoot()
    {
        // Guards the per-user .NET install scenario: with inheritance disabled, the launched
        // apphost must still be told where the runtime lives.
        var environment = TestServer.EnvironmentVariables();

        Assert.True(environment.TryGetValue("DOTNET_ROOT", out var root));
        Assert.True(Directory.Exists(Path.Combine(root!, "shared")));
    }

    [Fact]
    public async Task PrivateWorkerModeCompletesOneBoundedInspection()
    {
        var server = TestServer.ServerPath;
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var target = Path.Combine(directory, "target.xlsx");
        var reference = Path.Combine(directory, "reference.xlsx");
        File.WriteAllBytes(target, []);
        File.WriteAllBytes(reference, []);

        try
        {
            var startInfo = new ProcessStartInfo
            {
                FileName = server,
                UseShellExecute = false,
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };
            startInfo.ArgumentList.Add("--excel-worker");
            using var process = Process.Start(startInfo)!;
            var request = JsonSerializer.Serialize(new
            {
                version = 1,
                taskId = "process_probe",
                operation = "inspect",
                inspection = new WorkbookInspectionRequest(target, reference, WorkbookBinding.Isolated, SaveMode.Same, null)
            }, WorkerJsonOptions);
            await process.StandardInput.WriteLineAsync(request);
            process.StandardInput.Close();
            var stdout = await process.StandardOutput.ReadToEndAsync();
            var stderr = await process.StandardError.ReadToEndAsync();
            await process.WaitForExitAsync();

            Assert.True(process.ExitCode == 0, $"exit={process.ExitCode}; stdout={stdout}; stderr={stderr}");
            Assert.Empty(stderr);
            var lines = stdout.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries);
            Assert.Equal(4, lines.Length);
            var frameTypes = new List<string>(lines.Length);
            foreach (var line in lines)
            {
                Assert.InRange(System.Text.Encoding.UTF8.GetByteCount(line), 1, 16 * 1024);
                using var frame = JsonDocument.Parse(line);
                Assert.Equal(1, frame.RootElement.GetProperty("version").GetInt32());
                Assert.Equal("process_probe", frame.RootElement.GetProperty("taskId").GetString());
                frameTypes.Add(frame.RootElement.GetProperty("type").GetString()!);
            }

            Assert.Equal(["accepted", "phase", "phase", "result"], frameTypes);
        }
        finally
        {
            TempDirectory.Remove(directory);
        }
    }

    [Fact]
    public async Task StdioServerUsesPrivateInspectionWorkerFromAnUntrustedWorkingDirectory()
    {
        var server = TestServer.ServerPath;
        Assert.True(File.Exists(server), $"Expected MCP server apphost at {server}.");

        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var target = Path.Combine(directory, "target.xlsx");
        var reference = Path.Combine(directory, "reference.xlsx");
        File.WriteAllBytes(target, []);
        File.WriteAllBytes(reference, []);
        File.WriteAllText(Path.Combine(directory, "excel-task-mcp.exe"), "not an executable");

        try
        {
            var environment = TestServer.EnvironmentVariables();
            var transport = new StdioClientTransport(new StdioClientTransportOptions
            {
                Name = "ExcelTask-private-worker-test",
                Command = server,
                WorkingDirectory = directory,
                InheritEnvironmentVariables = false,
                EnvironmentVariables = environment,
                ShutdownTimeout = TimeSpan.FromSeconds(1)
            });

            await using var client = await McpClient.CreateAsync(
                transport,
                new McpClientOptions { ClientInfo = new() { Name = "ExcelTask-TestClient", Version = "1.0.0" } });
            var request = new ExcelTaskRequest(
                target,
                new ExcelOperation(
                    ExcelOperationKind.CopyExhibit,
                    CopyExhibit: new CopyExhibitOperation(reference, "Reference", "Imported", [])),
                ExcelTaskMode.Apply,
                WorkbookBinding.AskIfOpen,
                SaveMode.Same,
                OverwriteConfirmed: false);

            var result = await client.CallToolAsync(
                "excel_task",
                new Dictionary<string, object?> { ["request"] = request });

            Assert.False(result.IsError, JsonSerializer.Serialize(result));
            Assert.Equal(nameof(ExcelTaskStatus.NeedsConfirmation), result.StructuredContent!.Value.GetProperty("status").GetString());
            Assert.Equal(0, new FileInfo(target).Length);
            Assert.Equal(0, new FileInfo(reference).Length);
        }
        finally
        {
            TempDirectory.Remove(directory);
        }
    }

}
