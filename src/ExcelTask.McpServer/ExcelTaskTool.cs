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
    private const int MaxMacroMetadataLength = 96;
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
    [Description("Perform one bounded formula, exhibit, or macro-procedure operation in an existing .xlsx or .xlsm workbook. Plan previews without mutation; Apply saves, reopens, and verifies. Start with AskIfOpen, except EditMacroProcedure, which requires Isolated binding and Copy save to an .xlsm output, and whose Plan must omit every Apply-only field.")]
    public async Task<CallToolResult> RunAsync(
        [Description("The complete Excel task request.")] ExcelTaskRequest request,
        CancellationToken cancellationToken = default)
    {
        var receipt = BoundReceipt(await engine.RunAsync(request, cancellationToken), request.Mode == ExcelTaskMode.Plan);
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
        Confirmation = receipt.Confirmation with { Requirements = [] },
        MacroProcedure = BoundMacroProcedure(receipt.MacroProcedure, includeSource: false)
    };

    private static ExcelTaskReceipt BoundReceipt(ExcelTaskReceipt receipt, bool includeMacroSource) => receipt with
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
        },
        MacroProcedure = BoundMacroProcedure(receipt.MacroProcedure, includeMacroSource)
    };

    private static string? Bound(string? value) => value is { Length: > MaxReceiptStringLength }
        ? value[..MaxReceiptStringLength]
        : value;

    private static string BoundRequired(string value) => Bound(value) ?? string.Empty;

    private static MacroProcedureReceipt? BoundMacroProcedure(MacroProcedureReceipt? receipt, bool includeSource) => receipt is null
        ? null
        : receipt with
        {
            ComponentName = BoundMacroMetadata(receipt.ComponentName),
            ProcedureName = BoundMacroMetadata(receipt.ProcedureName),
            Sha256 = BoundMacroMetadata(receipt.Sha256),
            Source = includeSource ? BoundMacroSource(receipt.Source) : null
        };

    private static string BoundMacroMetadata(string value) => value.Length <= MaxMacroMetadataLength
        ? value
        : value[..MaxMacroMetadataLength];

    // Omitted rather than truncated, for the same reason as the engine boundary guard: partial VBA is
    // more dangerous to return than none, because a replacement written from it would destroy the
    // part the caller never saw.
    private static string? BoundMacroSource(string? source) => source is { Length: > MacroProcedureText.MaxSourceCharacters }
        ? null
        : source;
}
