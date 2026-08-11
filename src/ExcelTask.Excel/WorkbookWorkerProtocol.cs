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
    internal const int MaxTextLength = 128;


    /// <summary>
    /// Phase labels are progress annotations, not data. The supervisor both parses and validates
    /// against this one bound, so a long label can never make a healthy run look malformed.
    /// </summary>
    internal const int MaxPhaseLength = 64;
    internal const int MaxResultItems = 20;

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
        OpenWorkbookDescription = Bound(inspection.OpenWorkbookDescription),
        Checks = BoundChecks(inspection.Checks)
    };

    internal static WorkbookExecutionOutcome Bound(WorkbookExecutionOutcome outcome) => outcome with
    {
        Summary = BoundRequired(outcome.Summary),
        Changes = (outcome.Changes ?? []).Take(MaxResultItems).Select(change => change with
        {
            Kind = BoundRequired(change.Kind),
            Target = BoundRequired(change.Target),
            Summary = BoundRequired(change.Summary)
        }).ToArray(),
        Checks = BoundChecks(outcome.Checks),
        RetryReason = Bound(outcome.RetryReason),
        MacroProcedure = Bound(outcome.MacroProcedure),
        Audit = Bound(outcome.Audit),
        Range = Bound(outcome.Range)
    };

    /// <summary>
    /// Capped at the engine's read limit rather than at MaxResultItems, because these cells are the
    /// answer the caller asked for: trimming them to twenty would not be a bound, it would be a
    /// wrong result. The frame budget is what makes that safe to carry.
    /// </summary>
    private static WorksheetRangeReceipt? Bound(WorksheetRangeReceipt? range) => range is null ? null : range with
    {
        WorksheetName = BoundRequired(range.WorksheetName),
        Range = BoundRequired(range.Range),
        Cells = range.Cells.Take(ExcelTaskEngine.MaxReadCells).Select(cell => cell with
        {
            Address = BoundRequired(cell.Address),
            Text = cell.Text.Length > ExcelTaskEngine.MaxReadCellTextLength
                ? cell.Text[..ExcelTaskEngine.MaxReadCellTextLength]
                : cell.Text
        }).ToArray(),
        Truncated = range.Truncated || range.Cells.Count > ExcelTaskEngine.MaxReadCells
    };

    private static TaskCheck[] BoundChecks(IReadOnlyList<TaskCheck>? checks) => (checks ?? [])
        .Take(MaxResultItems)
        .Select(check => check with { Name = BoundRequired(check.Name), Detail = BoundRequired(check.Detail) })
        .ToArray();

    private static string? Bound(string? value) => value is { Length: > MaxTextLength } ? value[..MaxTextLength] : value;

    private static string BoundRequired(string? value) => Bound(value) ?? string.Empty;

    private static MacroProcedureReceipt? Bound(MacroProcedureReceipt? receipt) => receipt is null ? null : receipt with
    {
        ComponentName = BoundRequired(receipt.ComponentName),
        ProcedureName = BoundRequired(receipt.ProcedureName),
        Sha256 = BoundRequired(receipt.Sha256),
        Source = receipt.Source is { Length: > MacroProcedureText.MaxSourceCharacters } source ? source[..MacroProcedureText.MaxSourceCharacters] : receipt.Source
    };

    private static WorkbookAuditReceipt? Bound(WorkbookAuditReceipt? audit) => audit is null ? null : audit with
    {
        Items = audit.Items.Take(MaxResultItems).Select(item => item with
        {
            Kind = BoundRequired(item.Kind),
            Name = BoundRequired(item.Name),
            Detail = BoundRequired(item.Detail),
            DependsOn = Bound(item.DependsOn)
        }).ToArray(),
        // The real total survives the cap, so a trimmed report still says what it is not showing.
        Truncated = audit.Truncated || audit.Items.Count > MaxResultItems
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
