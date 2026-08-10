using System.ComponentModel;

namespace ExcelTask.Core;

public enum ExcelTaskMode { Plan, Apply }
public enum WorkbookBinding { AskIfOpen, UseOpen, Isolated }
public enum SaveMode { Same, Copy }
public enum ExcelTaskStatus { Planned, NeedsConfirmation, Completed, Rejected, Partial, Unknown }
public enum ExcelOperationKind { CopyExhibit, RepairExistingWorksheet, ExtendFormulaSeries, EditMacroProcedure, AuditWorkbookFlows }
public enum FormulaExtensionDirection { Right, Down }

public sealed record CopyExhibitOperation(
    [property: Description("Existing reference workbook path containing the worksheet to copy.")] string ReferenceWorkbookPath,
    [property: Description("Worksheet name in the reference workbook to copy.")] string ReferenceWorksheet,
    [property: Description("New worksheet name in the target workbook.")] string NewWorksheetName,
    [property: Description("Bounded A1 ranges on the copied worksheet where blank formulas may be repaired; use [] when none are needed.")] IReadOnlyList<string> RepairRanges);

public sealed record RepairExistingWorksheetOperation(
    [property: Description("Existing target worksheet name to repair.")] string WorksheetName,
    [property: Description("One or more bounded A1 ranges where blank formulas may be repaired.")] IReadOnlyList<string> Ranges);

public sealed record ExtendFormulaSeriesOperation(
    [property: Description("Existing target worksheet name containing the formula series.")] string WorksheetName,
    [property: Description("Right extends a horizontal series; Down extends a vertical series.")] FormulaExtensionDirection Direction,
    [property: Description("Exactly two adjacent evidence columns for Right or rows for Down, expressed as one A1 range.")] string EvidenceRange,
    [property: Description("Immediately adjacent blank destination columns for Right or rows for Down, expressed as one A1 range.")] string DestinationRange);

/// <summary>
/// Macro editing always uses workbookBinding Isolated and save Copy, on an .xlsm target and output.
/// Plan is inspect-only and must carry none of the Apply fields.
/// </summary>
public sealed record EditMacroProcedureOperation(
    [property: Description("Existing VBA component name containing the procedure to inspect or replace.")] string ComponentName,
    [property: Description("Existing VBA procedure name to inspect or replace.")] string ProcedureName,
    [property: Description("Apply only, and must be omitted for Plan: SHA-256 fingerprint of the existing procedure, taken from the Plan receipt.")] string? ExpectedProcedureSha256 = null,
    [property: Description("Apply only, and must be omitted for Plan: one complete replacement Sub or Function procedure with the requested name.")] string? ReplacementSource = null,
    [property: Description("Apply only, and must be omitted for Plan. When true, Apply runs the replacement procedure after the edit; the replacement must have zero parameters.")] bool RunAfterEdit = false);

/// <summary>
/// Reports how one workbook's data flows fit together: its Power Query queries and where each one
/// loads, its connections, its Data Model tables, relationships and measures, its PivotTables, and
/// the other workbooks it links to. It never changes anything, and it returns names and shapes
/// rather than data: no cell values, no query text, and no connection strings, because those carry
/// server names and credentials.
/// </summary>
public sealed record AuditWorkbookFlowsOperation();

/// <summary>Manual closed union for the operation selected by the one Excel task.</summary>
public sealed record ExcelOperation(
    [property: Description("Selects which one operation payload is supplied.")] ExcelOperationKind Kind,
    [property: Description("Required only when kind is CopyExhibit; all other payloads must be null.")] CopyExhibitOperation? CopyExhibit = null,
    [property: Description("Required only when kind is RepairExistingWorksheet; all other payloads must be null.")] RepairExistingWorksheetOperation? RepairExistingWorksheet = null,
    [property: Description("Required only when kind is ExtendFormulaSeries; all other payloads must be null.")] ExtendFormulaSeriesOperation? ExtendFormulaSeries = null,
    [property: Description("Required only when kind is EditMacroProcedure; all other payloads must be null.")] EditMacroProcedureOperation? EditMacroProcedure = null,
    [property: Description("Required only when kind is AuditWorkbookFlows; all other payloads must be null. Takes no options.")] AuditWorkbookFlowsOperation? AuditWorkbookFlows = null);

public sealed record ExcelTaskRequest(
    [property: Description("Existing target workbook path.")] string TargetWorkbookPath,
    [property: Description("The required manual operation union. Supply exactly one payload matching kind.")] ExcelOperation Operation,
    [property: Description("Plan previews without mutation; Apply performs the task after required confirmations.")] ExcelTaskMode Mode = ExcelTaskMode.Apply,
    [property: Description("Use AskIfOpen first; if confirmation is returned, resubmit with UseOpen or Isolated. EditMacroProcedure requires Isolated and rejects anything else.")] WorkbookBinding WorkbookBinding = WorkbookBinding.AskIfOpen,
    [property: Description("Same saves to the target; Copy saves only to outputWorkbookPath. EditMacroProcedure requires Copy to an .xlsm path.")] SaveMode Save = SaveMode.Same,
    [property: Description("Required destination path when save is Copy; omit for Same.")] string? OutputWorkbookPath = null,
    [property: Description("Explicit authorization required before Apply can overwrite an existing save destination.")] bool OverwriteConfirmed = false);

public sealed record WorkbookInspectionRequest(
    string TargetWorkbookPath,
    string? ReferenceWorkbookPath,
    WorkbookBinding Binding,
    SaveMode Save,
    string? OutputWorkbookPath);

public sealed record WorkbookInspection(bool TargetIsOpen, bool CopyOutputExists = false, string? OpenWorkbookDescription = null, IReadOnlyList<TaskCheck>? Checks = null);

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

public sealed record NormalizedCopyExhibitOperation(
    string ReferenceWorkbookPath,
    string ReferenceWorksheet,
    string NewWorksheetName,
    IReadOnlyList<FormulaRepairRange> RepairRanges);

public sealed record NormalizedRepairExistingWorksheetOperation(string WorksheetName, IReadOnlyList<FormulaRepairRange> Ranges);

public sealed record NormalizedExtendFormulaSeriesOperation(
    string WorksheetName,
    FormulaExtensionDirection Direction,
    FormulaRepairRange EvidenceRange,
    FormulaRepairRange DestinationRange);

public sealed record NormalizedEditMacroProcedureOperation(
    string ComponentName,
    string ProcedureName,
    string? ExpectedProcedureSha256,
    string? ReplacementSource,
    bool RunAfterEdit);

public sealed record NormalizedAuditWorkbookFlowsOperation();

/// <summary>Validated internal counterpart of <see cref="ExcelOperation"/>. It contains no legacy flat request fields.</summary>
public sealed record NormalizedExcelOperation(
    ExcelOperationKind Kind,
    NormalizedCopyExhibitOperation? CopyExhibit = null,
    NormalizedRepairExistingWorksheetOperation? RepairExistingWorksheet = null,
    NormalizedExtendFormulaSeriesOperation? ExtendFormulaSeries = null,
    NormalizedEditMacroProcedureOperation? EditMacroProcedure = null,
    NormalizedAuditWorkbookFlowsOperation? AuditWorkbookFlows = null);

public sealed record NormalizedExcelTaskRequest(
    string TargetWorkbookPath,
    ExcelTaskMode Mode,
    WorkbookBinding WorkbookBinding,
    SaveMode Save,
    string? OutputWorkbookPath,
    bool OverwriteConfirmed,
    NormalizedExcelOperation Operation);

public sealed record ExcelTaskPlan(string TaskId, NormalizedExcelTaskRequest Request);
public sealed record TaskChange(string Kind, string Target, string Summary);
public sealed record TaskCheck(string Name, bool Passed, string Detail);
public sealed record MacroProcedureReceipt(string ComponentName, string ProcedureName, string Sha256, string? Source, bool RunRequested, bool RunCompleted);

/// <summary>
/// One element of a workbook's data flow. <paramref name="Kind"/> says what it is - a query, a
/// connection, a model table, a relationship, a measure, a pivot, or an external link.
/// <paramref name="DependsOn"/> names what it reads from, which is what turns a list into a map.
/// Everything here is a name or a shape; never a value, query text, or connection string.
/// </summary>
public sealed record WorkbookFlowItem(string Kind, string Name, string Detail, string? DependsOn = null);

/// <summary>
/// A bounded description of one workbook's data flows. <paramref name="TotalFound"/> counts what
/// existed, not what fitted, so a truncated report can never be mistaken for a complete one.
/// </summary>
public sealed record WorkbookAuditReceipt(
    IReadOnlyList<WorkbookFlowItem> Items,
    int TotalFound,
    bool Truncated,
    bool WorkbookUnchanged);

public sealed record WorkbookExecutionOutcome(ExcelTaskStatus Status, string Summary, IReadOnlyList<TaskChange>? Changes = null, IReadOnlyList<TaskCheck>? Checks = null, bool CanRetry = false, string? RetryReason = null, MacroProcedureReceipt? MacroProcedure = null, WorkbookAuditReceipt? Audit = null);
public sealed record SaveReceipt(SaveMode Mode, string? OutputWorkbookPath, bool OverwriteConfirmed);
public sealed record RetryReceipt(bool CanRetry, string? Reason);
public sealed record ConfirmationRequirement(string Code, string Prompt);
public sealed record ConfirmationReceipt(bool Required, IReadOnlyList<ConfirmationRequirement> Requirements);
public sealed record PhaseTimings(TimeSpan Validation, TimeSpan Inspection, TimeSpan Execution, TimeSpan Total);
public sealed record ExcelTaskReceipt(string TaskId, ExcelTaskStatus Status, string Summary, IReadOnlyList<TaskChange> Changes, IReadOnlyList<TaskCheck> Checks, SaveReceipt Save, RetryReceipt Retry, ConfirmationReceipt Confirmation, PhaseTimings Timings, MacroProcedureReceipt? MacroProcedure = null, WorkbookAuditReceipt? Audit = null);
