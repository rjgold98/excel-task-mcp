using System.IO.Pipelines;
using ExcelTask.Core;
using ExcelTask.Excel;
using ExcelTask.McpServer;
using Microsoft.Extensions.DependencyInjection;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using Server = ModelContextProtocol.Server;

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

            await using var host = await InProcessExcelTaskServer.StartAsync();
            var tools = await host.Client.ListToolsAsync(new ListToolsRequestParams());
            var tool = Assert.Single(tools.Tools);
            Assert.Equal("excel_task", tool.Name);

            var result = await host.Client.CallToolAsync(
                "excel_task",
                new Dictionary<string, object?>
                {
                    ["request"] = new ExcelTaskRequest(
                        target,
                        reference,
                        "Reference",
                        "Imported",
                        ["A1:A3"],
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

    private sealed class InProcessExcelTaskServer : IAsyncDisposable
    {
        private readonly Pipe _clientToServer;
        private readonly Pipe _serverToClient;
        private readonly CancellationTokenSource _cancellation;
        private readonly ServiceProvider _services;
        private readonly Task _serverTask;
        private bool _disposed;

        private InProcessExcelTaskServer(
            Pipe clientToServer,
            Pipe serverToClient,
            CancellationTokenSource cancellation,
            ServiceProvider services,
            Task serverTask,
            McpClient client)
        {
            _clientToServer = clientToServer;
            _serverToClient = serverToClient;
            _cancellation = cancellation;
            _services = services;
            _serverTask = serverTask;
            Client = client;
        }

        public McpClient Client { get; }

        public static async Task<InProcessExcelTaskServer> StartAsync()
        {
            var clientToServer = new Pipe();
            var serverToClient = new Pipe();
            var cancellation = new CancellationTokenSource();
            var services = new ServiceCollection();
            services.AddSingleton<IWorkbookRuntime, ExcelWorkbookRuntime>();
            services.AddSingleton<IExcelTaskEngine, ExcelTaskEngine>();
            services
                .AddMcpServer(options => options.ServerInfo = new() { Name = "ExcelTask-RealExcel-Test", Version = "1.0.0" })
                .WithStreamServerTransport(clientToServer.Reader.AsStream(), serverToClient.Writer.AsStream())
                .WithTools<ExcelTaskTool>();

            var provider = services.BuildServiceProvider(validateScopes: true);
            var serverTask = provider.GetRequiredService<Server.McpServer>().RunAsync(cancellation.Token);
            try
            {
                var client = await McpClient.CreateAsync(
                    new StreamClientTransport(clientToServer.Writer.AsStream(), serverToClient.Reader.AsStream()),
                    new McpClientOptions { ClientInfo = new() { Name = "ExcelTask-RealExcel-TestClient", Version = "1.0.0" } },
                    cancellationToken: cancellation.Token);
                return new InProcessExcelTaskServer(clientToServer, serverToClient, cancellation, provider, serverTask, client);
            }
            catch
            {
                await DisposeStartupFailureAsync(clientToServer, serverToClient, cancellation, provider, serverTask);
                throw;
            }
        }

        public async ValueTask DisposeAsync()
        {
            if (_disposed) return;
            _disposed = true;
            await _cancellation.CancelAsync();
            _clientToServer.Writer.Complete();
            _serverToClient.Writer.Complete();
            await Client.DisposeAsync();
            try { await _serverTask; }
            catch (OperationCanceledException) { }
            await _services.DisposeAsync();
            _cancellation.Dispose();
        }

        private static async Task DisposeStartupFailureAsync(
            Pipe clientToServer,
            Pipe serverToClient,
            CancellationTokenSource cancellation,
            ServiceProvider services,
            Task serverTask)
        {
            await cancellation.CancelAsync();
            clientToServer.Writer.Complete();
            serverToClient.Writer.Complete();
            try { await serverTask; }
            catch (OperationCanceledException) { }
            await services.DisposeAsync();
            cancellation.Dispose();
        }
    }
}
