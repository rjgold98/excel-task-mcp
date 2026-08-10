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

// Field validation. Compiled rather than scripted because managed computers commonly run
// PowerShell in Constrained Language Mode. Neither this nor --excel-worker is an MCP tool.
if (args is ["--field-check", ..])
{
    Environment.ExitCode = await FieldCheck.RunAsync(args);
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
        Version = typeof(ExcelTaskTool).Assembly.GetName().Version?.ToString(3) ?? "0.3.0"
    })
    .WithStdioServerTransport()
    .WithTools<ExcelTaskTool>();

await builder.Build().RunAsync();
