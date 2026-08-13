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
    // This sentence is the whole of what a routing model reads before it decides whether to open the
    // schema. Keep the operation list and the direct-formula route visible: a caller that learns
    // either only from a rejection has already spent a round trip.
    [Description("Perform one bounded Excel operation on an existing .xlsx or .xlsm workbook: scan or audit structure, read a range, write constants or bounded A1 formulas, find/replace, format, create, copy a worksheet, repair or extend formulas, manage tables, Power Query, or Data Model measures and relationships, or replace one VBA procedure. Plan previews without mutation; Apply saves, reopens, and verifies. Start with AskIfOpen, except EditMacroProcedure, which requires Isolated binding and Copy save to an .xlsm output, and whose Plan must omit every Apply-only field.")]
    public async Task<CallToolResult> RunAsync(
        [Description("The complete Excel task request.")] ExcelTaskRequest request,
        CancellationToken cancellationToken = default)
    {
        // The caller's half of the round trip. The worker traces what Excel did; this traces what
        // the model asked for and what it got back, so one file reads as the whole interaction.
        var trace = ExcelTask.Excel.DiagnosticTrace.Begin(Guid.NewGuid().ToString("N"), "tool call");
        trace?.Note($"request: {request.Operation?.Kind.ToString() ?? "(no operation)"} mode={request.Mode} " +
                    $"binding={request.WorkbookBinding} save={request.Save} overwriteConfirmed={request.OverwriteConfirmed}");

        var receipt = BoundReceipt(await engine.RunAsync(request, cancellationToken), request.Mode == ExcelTaskMode.Plan);
        trace?.End("tool call", $"{receipt.Status}: {receipt.Summary}");
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
        MacroProcedure = ReceiptBounds.MacroProcedure(receipt.MacroProcedure, includeSource: false, ReceiptBounds.MaxModelTextLength),
        // The audit report is the largest field a receipt can carry, so a receipt that had to be
        // minimized is precisely the one that must drop it. Keeping it while the summary says
        // details were omitted would be untrue as well as oversized.
        Audit = null,
        // The cells are already gone by the time this runs, but the range's shape is kept and
        // marked truncated: that is what tells the caller to ask again for a narrower range
        // instead of concluding the sheet was empty.
        Range = receipt.Range is null ? null : receipt.Range with { Cells = [], Truncated = true },
        // An M expression runs to 8,192 characters, so it goes the way the range's cells do rather
        // than the way the audit does: the shape stays and the bulk leaves. Name, fingerprint and
        // length are what an Apply needs, and a length with no expression beside it is already how
        // "too long to send" reads, so this needs no truncation flag of its own.
        Query = receipt.Query is null ? null : receipt.Query with { Formula = null }
    };

    /// <summary>Last resort: status and identity only, for a receipt oversized without any detail.</summary>
    private static ExcelTaskReceipt EmptyReceipt(ExcelTaskReceipt receipt) => MinimalReceipt(receipt) with
    {
        Summary = "The receipt exceeded the MCP response size bound and was reduced to its status.",
        Save = receipt.Save with { OutputWorkbookPath = null },
        Retry = receipt.Retry with { Reason = null },
        MacroProcedure = null,
        Range = null,
        Query = null
    };

    // The model-facing seam bounds again rather than trusting the layers behind it - defence in
    // depth is deliberate. What it no longer has is its own opinion of what bounding means: caps,
    // truncation flags and the macro-source omit rule all come from ReceiptBounds.
    private static ExcelTaskReceipt BoundReceipt(ExcelTaskReceipt receipt, bool isPlan) => receipt with
    {
        TaskId = BoundRequired(receipt.TaskId),
        Summary = BoundRequired(receipt.Summary),
        Changes = ReceiptBounds.Changes(receipt.Changes, ReceiptBounds.MaxModelTextLength),
        Checks = ReceiptBounds.Checks(receipt.Checks, ReceiptBounds.MaxModelTextLength),
        Save = receipt.Save with { OutputWorkbookPath = Bound(receipt.Save.OutputWorkbookPath) },
        Retry = receipt.Retry with { Reason = Bound(receipt.Retry.Reason) },
        Confirmation = receipt.Confirmation with
        {
            Requirements = ReceiptBounds.Requirements(receipt.Confirmation.Requirements, ReceiptBounds.MaxModelTextLength)
        },
        MacroProcedure = ReceiptBounds.MacroProcedure(receipt.MacroProcedure, isPlan, ReceiptBounds.MaxModelTextLength),
        Audit = ReceiptBounds.Audit(receipt.Audit, ReceiptBounds.MaxModelTextLength),
        Range = ReceiptBounds.Range(receipt.Range, ReceiptBounds.MaxModelTextLength),
        Query = ReceiptBounds.Query(receipt.Query, isPlan, ReceiptBounds.MaxModelTextLength)
    };

    private static string? Bound(string? value) => ReceiptBounds.Text(value, ReceiptBounds.MaxModelTextLength);

    private static string BoundRequired(string value) => ReceiptBounds.RequiredText(value, ReceiptBounds.MaxModelTextLength);
}
