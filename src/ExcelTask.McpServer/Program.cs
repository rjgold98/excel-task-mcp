using ExcelTask.Core;
using ExcelTask.Excel;
using ExcelTask.McpServer;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;

var builder = Host.CreateApplicationBuilder(args);

builder.Logging.ClearProviders();

builder.Services.AddSingleton<IWorkbookRuntime, ExcelWorkbookRuntime>();
builder.Services.AddSingleton<IExcelTaskEngine, ExcelTaskEngine>();
builder.Services
    .AddMcpServer(options => options.ServerInfo = new()
    {
        Name = "ExcelTask",
        Version = typeof(ExcelTaskTool).Assembly.GetName().Version?.ToString(3) ?? "0.1.0"
    })
    .WithStdioServerTransport()
    .WithTools<ExcelTaskTool>();

await builder.Build().RunAsync();
