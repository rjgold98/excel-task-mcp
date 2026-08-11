using System.ComponentModel;

namespace ExcelTask.Core;

public enum ExcelTaskMode { Plan, Apply }
public enum WorkbookBinding { AskIfOpen, UseOpen, Isolated }
public enum SaveMode { Same, Copy }
public enum ExcelTaskStatus { Planned, NeedsConfirmation, Completed, Rejected, Partial, Unknown }
public enum ExcelOperationKind { CopyExhibit, RepairExistingWorksheet, ExtendFormulaSeries, EditMacroProcedure, AuditWorkbookFlows, ReadWorksheetRange, WriteWorksheetValues }
public enum FormulaExtensionDirection { Right, Down }

public sealed record CopyExhibitOperation(
    [property: Description("Existing reference workbook path containing the worksheet to copy.")] string ReferenceWorkbookPath,
    [property: Description("Worksheet name in the reference workbook to copy.")] string ReferenceWorksheet,
    [property: Description("New worksheet name in the target workbook.")] string NewWorksheetName,
    [property: Description("Bounded A1 ranges on the copied worksheet where blank formulas may be repaired; use [] when none are needed. At most 16 ranges and 10,000 cells per call; split a larger area across calls.")] IReadOnlyList<string> RepairRanges);

public sealed record RepairExistingWorksheetOperation(
    [property: Description("Existing target worksheet name to repair.")] string WorksheetName,
    [property: Description("One or more bounded A1 ranges where blank formulas may be repaired. At most 16 ranges and 10,000 cells per call; split a larger area across calls.")] IReadOnlyList<string> Ranges);

public sealed record ExtendFormulaSeriesOperation(
    [property: Description("Existing target worksheet name containing the formula series.")] string WorksheetName,
    [property: Description("Right extends a horizontal series; Down extends a vertical series.")] FormulaExtensionDirection Direction,
    [property: Description("Exactly two adjacent evidence columns for Right or rows for Down, expressed as one A1 range. Evidence and destination together must stay within 10,000 cells.")] string EvidenceRange,
    [property: Description("Immediately adjacent blank destination columns for Right or rows for Down, expressed as one A1 range. At most 24 periods and 2,000 destination cells, which binds before the 10,000-cell aggregate; split larger work across calls.")] string DestinationRange);

/// <summary>
/// Macro editing always uses workbookBinding Isolated and save Copy, on an .xlsm target and output.
/// Plan is inspect-only and must carry none of the Apply fields.
/// </summary>
public sealed record EditMacroProcedureOperation(
    [property: Description("Existing VBA component name containing the procedure to inspect or replace. If unknown, run AuditWorkbookFlows first; it lists every macro component and procedure.")] string ComponentName,
    [property: Description("Existing VBA procedure name to inspect or replace. Automatic-entry procedures such as Auto_Open cannot be edited.")] string ProcedureName,
    [property: Description("Apply only, and must be omitted for Plan: SHA-256 fingerprint of the existing procedure, taken from the Plan receipt.")] string? ExpectedProcedureSha256 = null,
    [property: Description("Apply only, and must be omitted for Plan: one complete replacement Sub or Function procedure with the requested name, at most 8,192 characters and 200 lines.")] string? ReplacementSource = null,
    [property: Description("Apply only, and must be omitted for Plan. When true, Apply runs the replacement after the edit; the replacement must have zero parameters and must not contain MsgBox, InputBox, GetOpenFilename, GetSaveAsFilename, FileDialog, or Stop, which wait for a person.")] bool RunAfterEdit = false);

/// <summary>
/// Reports how one workbook's data flows fit together: its Power Query queries and where each one
/// loads, its connections, its macro components and procedures, its Data Model tables,
/// relationships and measures, its PivotTables, and the other workbooks it links to. It never
/// changes anything, and it returns names and shapes rather than data: no cell values, no query
/// text, no VBA source, and no connection strings, because those carry server names and
/// credentials.
/// </summary>
public sealed record AuditWorkbookFlowsOperation();

/// <summary>
/// Reads one bounded range and returns what is in it.
///
/// This is the one operation that deliberately returns workbook contents. Every other receipt
/// withholds them, because there they would be incidental - nobody asked the audit for cell
/// values, so carrying them would be leakage. Here the contents are the entire request, and
/// refusing to answer would just send the caller to a different server, which is what real use
/// showed happening. The protection is therefore a hard bound rather than a refusal: a capped
/// range, capped cell text, and blanks omitted.
/// </summary>
public sealed record ReadWorksheetRangeOperation(
    [property: Description("Existing worksheet name to read. Run AuditWorkbookFlows first if the sheet names are unknown.")] string WorksheetName,
    [property: Description("One bounded A1 range to read, at most 400 cells. Narrow the range and read again if more is needed.")] string Range,
    [property: Description("False returns displayed values; true returns R1C1 formulas instead, for comparing a formula pattern.")] bool Formulas = false);

/// <summary>
/// Writes constants into named cells.
///
/// This server refuses model-written formula text, and that refusal is not relaxed here. The two
/// are not the same risk. A formula composed by a model can be plausible and silently wrong - off
/// by a row, anchored where it should be relative - and a receipt saying "it was written" would
/// confirm nothing about whether it was right. That is why ExtendFormulaSeries and
/// RepairExistingWorksheet infer formulas from evidence already in the sheet instead of accepting
/// them. A constant has no such failure mode: it is exactly what the caller named, and reading it
/// back proves it character for character. So values are writable and formulas are still not.
///
/// Anything beginning with "=" is rejected rather than written as text, because silently storing a
/// formula as a label is the kind of quiet wrongness this whole design exists to avoid.
/// </summary>
public sealed record WriteWorksheetValuesOperation(
    [property: Description("Existing worksheet name to write to. Run AuditWorkbookFlows first if the sheet names are unknown.")] string WorksheetName,
    [property: Description("Cells to write, at most 200. Each address must fall inside one bounded A1 range no larger than 400 cells.")] IReadOnlyList<WorksheetCellValue> Cells);

/// <summary>One constant to write. Text starting with = is rejected; this operation never writes formulas.</summary>
public sealed record WorksheetCellValue(
    [property: Description("Single A1 cell address, such as B7.")] string Address,
    [property: Description("Constant to write. Numbers and TRUE/FALSE are written as numbers and booleans, everything else as text. Must not start with =.")] string Value);

/// <summary>Manual closed union for the operation selected by the one Excel task.</summary>
public sealed record ExcelOperation(
    [property: Description("Selects which one operation payload is supplied.")] ExcelOperationKind Kind,
    [property: Description("Required only when kind is CopyExhibit; all other payloads must be null.")] CopyExhibitOperation? CopyExhibit = null,
    [property: Description("Required only when kind is RepairExistingWorksheet; all other payloads must be null.")] RepairExistingWorksheetOperation? RepairExistingWorksheet = null,
    [property: Description("Required only when kind is ExtendFormulaSeries; all other payloads must be null.")] ExtendFormulaSeriesOperation? ExtendFormulaSeries = null,
    [property: Description("Required only when kind is EditMacroProcedure; all other payloads must be null.")] EditMacroProcedureOperation? EditMacroProcedure = null,
    [property: Description("Required only when kind is AuditWorkbookFlows; all other payloads must be null. Takes no options. The read-only report lists worksheets, tables, defined names, queries, connections, macro components and procedures, the data model, pivots, and external links.")] AuditWorkbookFlowsOperation? AuditWorkbookFlows = null,
    [property: Description("Required only when kind is ReadWorksheetRange; all other payloads must be null. Reads one bounded range and returns its contents.")] ReadWorksheetRangeOperation? ReadWorksheetRange = null,
    [property: Description("Required only when kind is WriteWorksheetValues; all other payloads must be null. Writes constants into named cells and reads them back. Never accepts formula text.")] WriteWorksheetValuesOperation? WriteWorksheetValues = null);

public sealed record ExcelTaskRequest(
    [property: Description("Existing target workbook path, ending .xlsx or .xlsm.")] string TargetWorkbookPath,
    [property: Description("The required manual operation union. Supply exactly one payload matching kind.")] ExcelOperation Operation,
    [property: Description("Plan previews without mutation; Apply performs the task after required confirmations.")] ExcelTaskMode Mode = ExcelTaskMode.Apply,
    [property: Description("Use AskIfOpen when the workbook state is unknown; resubmit with UseOpen or Isolated if confirmation is returned. When the request already says the workbook is open, use UseOpen directly to avoid a wasted round trip. EditMacroProcedure requires Isolated. UseOpen cannot be combined with Copy, and Isolated with save Same is rejected while the target is open in Excel.")] WorkbookBinding WorkbookBinding = WorkbookBinding.AskIfOpen,
    [property: Description("Same saves to the target; Copy saves only to outputWorkbookPath. EditMacroProcedure requires Copy to an .xlsm path. Copy is rejected with UseOpen. AuditWorkbookFlows and ReadWorksheetRange never write: leave Same with no outputWorkbookPath.")] SaveMode Save = SaveMode.Same,
    [property: Description("Required destination path when save is Copy; omit for Same. Must differ from the target path and carry the target's extension.")] string? OutputWorkbookPath = null,
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

public sealed record NormalizedReadWorksheetRangeOperation(string WorksheetName, FormulaRepairRange Range, bool Formulas);

public sealed record NormalizedWorksheetCellValue(string Address, string Value);

public sealed record NormalizedWriteWorksheetValuesOperation(string WorksheetName, IReadOnlyList<NormalizedWorksheetCellValue> Cells);

/// <summary>Validated internal counterpart of <see cref="ExcelOperation"/>. It contains no legacy flat request fields.</summary>
public sealed record NormalizedExcelOperation(
    ExcelOperationKind Kind,
    NormalizedCopyExhibitOperation? CopyExhibit = null,
    NormalizedRepairExistingWorksheetOperation? RepairExistingWorksheet = null,
    NormalizedExtendFormulaSeriesOperation? ExtendFormulaSeries = null,
    NormalizedEditMacroProcedureOperation? EditMacroProcedure = null,
    NormalizedAuditWorkbookFlowsOperation? AuditWorkbookFlows = null,
    NormalizedReadWorksheetRangeOperation? ReadWorksheetRange = null,
    NormalizedWriteWorksheetValuesOperation? WriteWorksheetValues = null);

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

/// <summary>One cell that had something in it. Blanks are omitted rather than reported as empty.</summary>
public sealed record WorksheetCell(string Address, string Text);

/// <summary>
/// The contents of one bounded range. <paramref name="CellsInRange"/> counts what the range spans
/// and <paramref name="Truncated"/> says whether the list stops short, so a partial answer can
/// never read as a complete one.
/// </summary>
public sealed record WorksheetRangeReceipt(
    string WorksheetName,
    string Range,
    bool Formulas,
    int CellsInRange,
    int NonEmptyCells,
    IReadOnlyList<WorksheetCell> Cells,
    bool Truncated);

public sealed record WorkbookExecutionOutcome(ExcelTaskStatus Status, string Summary, IReadOnlyList<TaskChange>? Changes = null, IReadOnlyList<TaskCheck>? Checks = null, bool CanRetry = false, string? RetryReason = null, MacroProcedureReceipt? MacroProcedure = null, WorkbookAuditReceipt? Audit = null, WorksheetRangeReceipt? Range = null);
public sealed record SaveReceipt(SaveMode Mode, string? OutputWorkbookPath, bool OverwriteConfirmed);
public sealed record RetryReceipt(bool CanRetry, string? Reason);
public sealed record ConfirmationRequirement(string Code, string Prompt);
public sealed record ConfirmationReceipt(bool Required, IReadOnlyList<ConfirmationRequirement> Requirements);
public sealed record PhaseTimings(TimeSpan Validation, TimeSpan Inspection, TimeSpan Execution, TimeSpan Total);
public sealed record ExcelTaskReceipt(string TaskId, ExcelTaskStatus Status, string Summary, IReadOnlyList<TaskChange> Changes, IReadOnlyList<TaskCheck> Checks, SaveReceipt Save, RetryReceipt Retry, ConfirmationReceipt Confirmation, PhaseTimings Timings, MacroProcedureReceipt? MacroProcedure = null, WorkbookAuditReceipt? Audit = null, WorksheetRangeReceipt? Range = null);
