using System.Text.Json;
using System.Text.Json.Serialization;
using ExcelTask.Core;

namespace ExcelTask.Excel;

/// <summary>Private, versioned protocol used only between the supervisor and its Excel worker process.</summary>
internal static class WorkbookWorkerProtocol
{
    internal const int Version = 1;
    internal const int MaxRequestBytes = 1024 * 1024;
    internal const int MaxFrameBytes = 16 * 1024;
    internal const int MaxTextLength = 128;
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
        RetryReason = Bound(outcome.RetryReason)
    };

    private static TaskCheck[] BoundChecks(IReadOnlyList<TaskCheck>? checks) => (checks ?? [])
        .Take(MaxResultItems)
        .Select(check => check with { Name = BoundRequired(check.Name), Detail = BoundRequired(check.Detail) })
        .ToArray();

    private static string? Bound(string? value) => value is { Length: > MaxTextLength } ? value[..MaxTextLength] : value;

    private static string BoundRequired(string? value) => Bound(value) ?? string.Empty;

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
