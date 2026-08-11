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
    private const int MaxReceiptItems = 20;
    private const int MaxMcpResultBytes = 30 * 1024;
    private const int MaxMacroMetadataLength = 96;
    private const int MaxReceiptCellTextLength = 64;
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
        if (WithinResponseBound(result)) return result;

        // Range cells are the answer to a read, not incidental detail, so an oversized read gives
        // back as many as fit instead of being emptied like the fields below. Halving terminates:
        // the count strictly decreases and an empty list leaves the loop.
        while (receipt.Range is { Cells.Count: > 0 } range)
        {
            receipt = receipt with
            {
                Range = range with { Cells = range.Cells.Take(range.Cells.Count / 2).ToArray(), Truncated = true }
            };
            result = CreateResult(receipt);
            if (WithinResponseBound(result)) return result;
        }

        // The minimal receipt is measured too. Claiming details were omitted while returning an
        // oversized response would be both untrue and useless to the caller.
        var minimal = CreateResult(MinimalReceipt(receipt));
        return WithinResponseBound(minimal) ? minimal : CreateResult(EmptyReceipt(receipt));
    }

    private static bool WithinResponseBound(CallToolResult result) =>
        JsonSerializer.SerializeToUtf8Bytes(result, ReceiptJsonOptions).Length <= MaxMcpResultBytes;

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
        MacroProcedure = BoundMacroProcedure(receipt.MacroProcedure, includeSource: false),
        // The audit report is the largest field a receipt can carry, so a receipt that had to be
        // minimized is precisely the one that must drop it. Keeping it while the summary says
        // details were omitted would be untrue as well as oversized.
        Audit = null,
        // The cells are already gone by the time this runs, but the range's shape is kept and
        // marked truncated: that is what tells the caller to ask again for a narrower range
        // instead of concluding the sheet was empty.
        Range = receipt.Range is null ? null : receipt.Range with { Cells = [], Truncated = true }
    };

    /// <summary>Last resort: status and identity only, for a receipt oversized without any detail.</summary>
    private static ExcelTaskReceipt EmptyReceipt(ExcelTaskReceipt receipt) => MinimalReceipt(receipt) with
    {
        Summary = "The receipt exceeded the MCP response size bound and was reduced to its status.",
        Save = receipt.Save with { OutputWorkbookPath = null },
        Retry = receipt.Retry with { Reason = null },
        MacroProcedure = null,
        Range = null
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
        MacroProcedure = BoundMacroProcedure(receipt.MacroProcedure, includeMacroSource),
        Audit = BoundAudit(receipt.Audit),
        Range = BoundRange(receipt.Range)
    };

    /// <summary>
    /// Cell text is bounded far shorter than the range is deep: a long cell is usually a note or a
    /// pasted paragraph, and one of them must not crowd out the four hundred cells around it.
    /// </summary>
    private static WorksheetRangeReceipt? BoundRange(WorksheetRangeReceipt? range) => range is null ? null : range with
    {
        WorksheetName = BoundRequired(range.WorksheetName),
        Range = BoundRequired(range.Range),
        Cells = range.Cells.Take(ExcelTaskEngine.MaxReadCells).Select(cell => cell with
        {
            Address = BoundRequired(cell.Address),
            Text = cell.Text.Length > MaxReceiptCellTextLength ? cell.Text[..MaxReceiptCellTextLength] : cell.Text
        }).ToArray(),
        Truncated = range.Truncated || range.Cells.Count > ExcelTaskEngine.MaxReadCells
    };

    /// <summary>
    /// The audit report is the one receipt field that grows with the workbook, so the model-facing
    /// seam bounds it like every other field rather than trusting the layers behind it.
    /// </summary>
    private static WorkbookAuditReceipt? BoundAudit(WorkbookAuditReceipt? audit) => audit is null ? null : audit with
    {
        Items = audit.Items.Take(MaxReceiptItems).Select(item => item with
        {
            Kind = BoundRequired(item.Kind),
            Name = BoundRequired(item.Name),
            Detail = BoundRequired(item.Detail),
            DependsOn = Bound(item.DependsOn)
        }).ToArray(),
        Truncated = audit.Truncated || audit.Items.Count > MaxReceiptItems
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
