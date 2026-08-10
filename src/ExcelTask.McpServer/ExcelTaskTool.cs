using System.ComponentModel;
using System.Text.Json;
using System.Text.Json.Serialization;
using ExcelTask.Core;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;

namespace ExcelTask.McpServer;

[McpServerToolType]
public sealed class ExcelTaskTool(IExcelTaskEngine engine)
{
    private const int MaxReceiptStringLength = 96;
    private const int MaxReceiptChanges = 20;
    private const int MaxReceiptChecks = 20;
    private const int MaxMcpResultBytes = 30 * 1024;
    private static readonly JsonSerializerOptions ReceiptJsonOptions = new(JsonSerializerDefaults.Web)
    {
        Converters = { new JsonStringEnumConverter() }
    };

    [McpServerTool(
        Name = "excel_task",
        Title = "Excel Task",
        Destructive = true,
        UseStructuredContent = true,
        OutputSchemaType = typeof(ExcelTaskReceipt))]
    [Description("Perform one bounded formula or exhibit operation in an existing .xlsx or .xlsm workbook: copy a named exhibit, repair safely inferable blank formulas, or extend a proven formula series right or down. Plan previews without mutation; Apply saves, reopens, and verifies. Start with AskIfOpen.")]
    public async Task<CallToolResult> RunAsync(
        [Description("The complete Excel task request.")] ExcelTaskRequest request,
        CancellationToken cancellationToken = default)
    {
        var receipt = BoundReceipt(await engine.RunAsync(request, cancellationToken));
        var result = CreateResult(receipt);

        return JsonSerializer.SerializeToUtf8Bytes(result, ReceiptJsonOptions).Length <= MaxMcpResultBytes
            ? result
            : CreateResult(MinimalReceipt(receipt));
    }

    private static CallToolResult CreateResult(ExcelTaskReceipt receipt) => new()
    {
        Content = [new TextContentBlock { Text = $"{receipt.Status}: {receipt.Summary}" }],
        StructuredContent = JsonSerializer.SerializeToElement(receipt, ReceiptJsonOptions),
        IsError = receipt.Status is ExcelTaskStatus.Rejected or ExcelTaskStatus.Partial or ExcelTaskStatus.Unknown
    };

    private static ExcelTaskReceipt MinimalReceipt(ExcelTaskReceipt receipt) => receipt with
    {
        Summary = "Receipt details were omitted to preserve the MCP response size bound.",
        Changes = [],
        Checks = [],
        Save = receipt.Save with { OutputWorkbookPath = Bound(receipt.Save.OutputWorkbookPath) },
        Retry = receipt.Retry with { Reason = Bound(receipt.Retry.Reason) },
        Confirmation = receipt.Confirmation with { Requirements = [] }
    };

    private static ExcelTaskReceipt BoundReceipt(ExcelTaskReceipt receipt) => receipt with
    {
        TaskId = BoundRequired(receipt.TaskId),
        Summary = BoundRequired(receipt.Summary),
        Changes = receipt.Changes.Take(MaxReceiptChanges).Select(change => change with
        {
            Kind = BoundRequired(change.Kind),
            Target = BoundRequired(change.Target),
            Summary = BoundRequired(change.Summary)
        }).ToArray(),
        Checks = receipt.Checks.Take(MaxReceiptChecks).Select(check => check with
        {
            Name = BoundRequired(check.Name),
            Detail = BoundRequired(check.Detail)
        }).ToArray(),
        Save = receipt.Save with { OutputWorkbookPath = Bound(receipt.Save.OutputWorkbookPath) },
        Retry = receipt.Retry with { Reason = Bound(receipt.Retry.Reason) },
        Confirmation = receipt.Confirmation with
        {
            Requirements = receipt.Confirmation.Requirements.Take(MaxReceiptChecks).Select(requirement => requirement with
            {
                Code = BoundRequired(requirement.Code),
                Prompt = BoundRequired(requirement.Prompt)
            }).ToArray()
        }
    };

    private static string? Bound(string? value) => value is { Length: > MaxReceiptStringLength }
        ? value[..MaxReceiptStringLength]
        : value;

    private static string BoundRequired(string value) => Bound(value) ?? string.Empty;
}
