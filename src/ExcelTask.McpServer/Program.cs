using ExcelTask.Core;
using ExcelTask.Excel;
using ExcelTask.McpServer;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;

if (args is ["--excel-worker"])
{
    Environment.ExitCode = await WorkbookWorkerHost.RunAsync(Console.In, Console.Out);
    return;
}

var builder = Host.CreateApplicationBuilder(args);

builder.Logging.ClearProviders();

builder.Services.AddSingleton<IWorkbookRuntime, SupervisedWorkbookRuntime>();
builder.Services.AddSingleton<IExcelTaskEngine, ExcelTaskEngine>();
builder.Services
    .AddMcpServer(options => options.ServerInfo = new()
    {
        Name = "ExcelTask",
        Version = typeof(ExcelTaskTool).Assembly.GetName().Version?.ToString(3) ?? "0.2.0"
    })
    .WithStdioServerTransport()
    .WithTools<ExcelTaskTool>();

await builder.Build().RunAsync();
