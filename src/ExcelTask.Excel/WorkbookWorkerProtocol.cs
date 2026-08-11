using System.Text.Json;
using System.Text.Json.Serialization;
using ExcelTask.Core;

namespace ExcelTask.Excel;

/// <summary>Private, versioned protocol used only between the supervisor and its Excel worker process.</summary>
internal static class WorkbookWorkerProtocol
{
    internal const int Version = 1;
    internal const int MaxRequestBytes = 1024 * 1024;

    /// <summary>
    /// A worker frame over this size is replaced with a fatal code, so the budget has to fit the
    /// largest legitimate result. It was 16 KB when every receipt was metadata. A full range read is
    /// 400 cells of up to 64 characters each and does not fit, and the failure would have appeared
    /// only on large reads - the case a real model produces - as a lost result rather than a
    /// truncated one. It stays far below the MCP response bound, which is what actually decides what
    /// the caller sees.
    /// </summary>
    internal const int MaxFrameBytes = 64 * 1024;
    internal const int MaxTextLength = ReceiptBounds.MaxFrameTextLength;

    /// <summary>
    /// Phase labels are progress annotations, not data. The supervisor both parses and validates
    /// against this one bound, so a long label can never make a healthy run look malformed.
    /// </summary>
    internal const int MaxPhaseLength = 64;

    internal static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web)
    {
        PropertyNameCaseInsensitive = false,
        Converters = { new JsonStringEnumConverter(JsonNamingPolicy.CamelCase) }
    };

    internal static bool TryParseRequest(string line, out WorkbookWorkerRequest? request)
    {
        request = null;
        try
        {
            using var document = JsonDocument.Parse(line, new JsonDocumentOptions { MaxDepth = 16 });
            var root = document.RootElement;
            if (root.ValueKind != JsonValueKind.Object ||
                !root.TryGetProperty("version", out var version) || version.ValueKind != JsonValueKind.Number || version.GetInt32() != Version ||
                !root.TryGetProperty("taskId", out var taskId) || taskId.ValueKind != JsonValueKind.String || !WorkbookRuntimeHelpers.IsSafeTaskId(taskId.GetString()) ||
                !root.TryGetProperty("operation", out var operation) || operation.ValueKind != JsonValueKind.String)
            {
                return false;
            }

            var hasInspection = root.TryGetProperty("inspection", out var inspectionElement) && inspectionElement.ValueKind != JsonValueKind.Null;
            var hasPlan = root.TryGetProperty("plan", out var planElement) && planElement.ValueKind != JsonValueKind.Null;
            if (hasInspection == hasPlan) return false;

            var taskIdValue = taskId.GetString()!;
            switch (operation.GetString())
            {
                case "inspect" when hasInspection:
                    var inspection = JsonSerializer.Deserialize<WorkbookInspectionRequest>(inspectionElement.GetRawText(), JsonOptions);
                    if (inspection is null) return false;
                    request = new WorkbookWorkerRequest(Version, taskIdValue, "inspect", inspection, null);
                    return true;

                case "execute" when hasPlan:
                    var plan = JsonSerializer.Deserialize<ExcelTaskPlan>(planElement.GetRawText(), JsonOptions);
                    if (plan is null || !string.Equals(plan.TaskId, taskIdValue, StringComparison.Ordinal)) return false;
                    request = new WorkbookWorkerRequest(Version, taskIdValue, "execute", null, plan);
                    return true;

                default:
                    return false;
            }
        }
        catch (JsonException)
        {
            return false;
        }
        catch (NotSupportedException)
        {
            return false;
        }
        catch (FormatException)
        {
            return false;
        }
        catch (InvalidOperationException)
        {
            return false;
        }
        catch (OverflowException)
        {
            return false;
        }
    }

    internal static WorkbookInspection Bound(WorkbookInspection inspection) => inspection with
    {
        OpenWorkbookDescription = ReceiptBounds.Text(inspection.OpenWorkbookDescription, MaxTextLength),
        Checks = ReceiptBounds.Checks(inspection.Checks, MaxTextLength)
    };

    // Bounded for the frame budget, using the one implementation of what bounding means. The macro
    // source rule matters most here: this layer used to TRUNCATE an oversized source while the two
    // layers downstream deliberately omit one - so a source clipped to exactly the limit arrived
    // measuring within it and passed as complete, defeating both. ReceiptBounds omits, everywhere.
    internal static WorkbookExecutionOutcome Bound(WorkbookExecutionOutcome outcome) => outcome with
    {
        Summary = ReceiptBounds.RequiredText(outcome.Summary, MaxTextLength),
        Changes = ReceiptBounds.Changes(outcome.Changes, MaxTextLength),
        Checks = ReceiptBounds.Checks(outcome.Checks, MaxTextLength),
        RetryReason = ReceiptBounds.Text(outcome.RetryReason, MaxTextLength),
        MacroProcedure = ReceiptBounds.MacroProcedure(outcome.MacroProcedure, includeSource: true, MaxTextLength),
        Audit = ReceiptBounds.Audit(outcome.Audit, MaxTextLength),
        Range = ReceiptBounds.Range(outcome.Range, MaxTextLength)
    };
}

internal sealed record WorkbookWorkerRequest(
    int Version,
    string TaskId,
    string Operation,
    WorkbookInspectionRequest? Inspection,
    ExcelTaskPlan? Plan);

internal interface IExcelWorkbookRuntimeObserver
{
    void OnPhase(string phase);

    void OnOwnedProcessCaptured(ProcessIdentity identity);

    void OnStagingPathCreated(string stagingPath);
}

internal sealed class NullExcelWorkbookRuntimeObserver : IExcelWorkbookRuntimeObserver
{
    public static readonly NullExcelWorkbookRuntimeObserver Instance = new();

    private NullExcelWorkbookRuntimeObserver() { }

    public void OnPhase(string phase) { }

    public void OnOwnedProcessCaptured(ProcessIdentity identity) { }

    public void OnStagingPathCreated(string stagingPath) { }
}
