using System.ComponentModel;

namespace ExcelTask.Core;

public enum ExcelTaskMode
{
    Plan,
    Apply
}

public enum WorkbookBinding
{
    AskIfOpen,
    UseOpen,
    Isolated
}

public enum SaveMode
{
    Same,
    Copy
}

public enum ExcelTaskStatus
{
    Planned,
    NeedsConfirmation,
    Completed,
    Rejected,
    Partial,
    Unknown
}

public sealed record ExcelTaskRequest(
    [property: Description("Existing target workbook path.")]
    string TargetWorkbookPath,
    [property: Description("Workbook containing the named reference worksheet.")]
    string ReferenceWorkbookPath,
    [property: Description("Reference worksheet name to copy from.")]
    string ReferenceWorksheet,
    [property: Description("Name for the new worksheet in the target workbook.")]
    string NewWorksheetName,
    [property: Description("Bounded A1 ranges for blank-formula repair; use [] when no repairs are needed.")]
    IReadOnlyList<string> FormulaRepairRanges,
    [property: Description("Plan previews without mutation; Apply performs the task after required confirmations.")]
    ExcelTaskMode Mode,
    [property: Description("Use AskIfOpen first; if confirmation is returned, resubmit with UseOpen or Isolated.")]
    WorkbookBinding WorkbookBinding,
    [property: Description("Same saves to the target; Copy saves only to outputWorkbookPath.")]
    SaveMode Save,
    [property: Description("Required destination path when save is Copy; omit for Same.")]
    string? OutputWorkbookPath = null,
    [property: Description("Explicit authorization required before Apply can overwrite an existing save destination.")]
    bool OverwriteConfirmed = false);

public sealed record WorkbookInspectionRequest(
    string TargetWorkbookPath,
    string ReferenceWorkbookPath,
    WorkbookBinding Binding,
    SaveMode Save,
    string? OutputWorkbookPath);

public sealed record WorkbookInspection(
    bool TargetIsOpen,
    bool CopyOutputExists = false,
    string? OpenWorkbookDescription = null,
    IReadOnlyList<TaskCheck>? Checks = null);

public interface IWorkbookRuntime
{
    Task<WorkbookInspection> InspectAsync(WorkbookInspectionRequest request, CancellationToken cancellationToken);

    Task<WorkbookExecutionOutcome> ExecuteAsync(ExcelTaskPlan plan, CancellationToken cancellationToken);
}

public interface IExcelTaskEngine
{
    Task<ExcelTaskReceipt> RunAsync(ExcelTaskRequest request, CancellationToken cancellationToken);
}

public sealed record FormulaRepairRange(string StartCell, string EndCell)
{
    public override string ToString() => StartCell == EndCell ? StartCell : $"{StartCell}:{EndCell}";
}

public sealed record NormalizedExcelTaskRequest(
    string TargetWorkbookPath,
    string ReferenceWorkbookPath,
    string ReferenceWorksheet,
    string NewWorksheetName,
    IReadOnlyList<FormulaRepairRange> FormulaRepairRanges,
    ExcelTaskMode Mode,
    WorkbookBinding WorkbookBinding,
    SaveMode Save,
    string? OutputWorkbookPath,
    bool OverwriteConfirmed);

public sealed record ExcelTaskPlan(string TaskId, NormalizedExcelTaskRequest Request);

public sealed record TaskChange(string Kind, string Target, string Summary);

public sealed record TaskCheck(string Name, bool Passed, string Detail);

public sealed record WorkbookExecutionOutcome(
    ExcelTaskStatus Status,
    string Summary,
    IReadOnlyList<TaskChange>? Changes = null,
    IReadOnlyList<TaskCheck>? Checks = null,
    bool CanRetry = false,
    string? RetryReason = null);

public sealed record SaveReceipt(SaveMode Mode, string? OutputWorkbookPath, bool OverwriteConfirmed);

public sealed record RetryReceipt(bool CanRetry, string? Reason);

public sealed record ConfirmationRequirement(string Code, string Prompt);

public sealed record ConfirmationReceipt(
    bool Required,
    IReadOnlyList<ConfirmationRequirement> Requirements);

public sealed record PhaseTimings(
    TimeSpan Validation,
    TimeSpan Inspection,
    TimeSpan Execution,
    TimeSpan Total);

public sealed record ExcelTaskReceipt(
    string TaskId,
    ExcelTaskStatus Status,
    string Summary,
    IReadOnlyList<TaskChange> Changes,
    IReadOnlyList<TaskCheck> Checks,
    SaveReceipt Save,
    RetryReceipt Retry,
    ConfirmationReceipt Confirmation,
    PhaseTimings Timings);
