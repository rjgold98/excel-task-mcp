using System.ComponentModel;

namespace ExcelTask.Core;

public enum ExcelTaskMode { Plan, Apply }
public enum WorkbookBinding { AskIfOpen, UseOpen, Isolated }
public enum SaveMode { Same, Copy }
public enum ExcelTaskStatus { Planned, NeedsConfirmation, Completed, Rejected, Partial, Unknown }
public enum ExcelOperationKind { CopyExhibit, RepairExistingWorksheet, ExtendFormulaSeries, EditMacroProcedure, AuditWorkbookFlows, ReadWorksheetRange, WriteWorksheetValues, FindReplace, Create, SetNumberFormat, ScanWorkbookStructure }
public enum CreateKind { Workbook, Worksheet }
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

/// <summary>
/// Finds cells whose text matches, and optionally replaces it.
///
/// Plan reports the matches and changes nothing, which is the point: replace-across-a-sheet is the
/// operation most likely to be regretted, so the caller sees exactly which cells would change
/// before authorizing it. Only constants are searched and rewritten - a cell containing a formula
/// is reported as a match but never edited, because rewriting formula text is the thing this
/// server refuses everywhere else and a find/replace is no different.
/// </summary>
public sealed record FindReplaceOperation(
    [property: Description("Existing worksheet name to search. Run AuditWorkbookFlows first if the sheet names are unknown.")] string WorksheetName,
    [property: Description("Text to find, at most 200 characters. Matching is plain text on what the cell displays; * and ? are literal, not wildcards.")] string Find,
    [property: Description("Apply only, and must be omitted for Plan: replacement text, at most 200 characters. Cells holding formulas are reported but never rewritten, and a replacement that would leave a cell starting with = is refused before anything is written.")] string? ReplaceWith = null,
    [property: Description("One bounded A1 range to search, at most 10,000 cells; omit to search the worksheet's used range when it is within that bound.")] string? Range = null,
    [property: Description("True matches only whole cell contents; false matches anywhere in the cell.")] bool WholeCell = false,
    [property: Description("True makes matching case-sensitive.")] bool MatchCase = false);

/// <summary>
/// Creates an empty workbook, or adds an empty worksheet to an existing one.
///
/// Every other operation requires a target that already exists, which made the first step of real
/// work impossible: nine of forty-six measured sessions began by creating a workbook and seven by
/// adding a sheet. Creation is deliberately empty - no template, no seeded content - because the
/// operations that fill a sheet already exist and each verifies its own work.
/// </summary>
public sealed record CreateOperation(
    [property: Description("Workbook creates an empty workbook at targetWorkbookPath, which must not already exist. Worksheet adds an empty sheet to the existing target. Either way Create writes the target itself: it requires binding Isolated and takes no save destination.")] CreateKind Kind,
    [property: Description("Required for Worksheet, and must not already be in use. Optional for Workbook: names its one starting sheet, so a single call yields a workbook ready to write to. Omitted, the receipt reports the default name Excel chose.")] string? WorksheetName = null);

/// <summary>
/// Sets one number format code across one bounded range.
///
/// This closes a gap the server made for itself. WriteWorksheetValues can put 1000.5 into a cell
/// and had no way to make it read as 1,000.50 or (1,000.50), so a correct number could still be
/// presented wrongly - which for a financial exhibit is the difference between usable and not.
///
/// It is deliberately only the number format. Fonts, fills, borders, widths and conditional formats
/// are not here and are not coming without evidence: the measured demand names the tool, not which
/// of its operations were used, and building all of it on a tool-level count is how a server ends
/// up with 230 operations and no idea which matter. The name says exactly what it sets so a caller
/// does not spend a round trip discovering the rest is missing.
/// </summary>
public sealed record SetNumberFormatOperation(
    [property: Description("Existing worksheet name to format. Run AuditWorkbookFlows first if the sheet names are unknown.")] string WorksheetName,
    [property: Description("One bounded A1 range to format, at most 10,000 cells. The format applies to every cell in it, including blank ones.")] string Range,
    [property: Description("Excel number format code in US-English form, at most 255 characters, such as #,##0.00 or #,##0;(#,##0) or 0.0% or yyyy-mm-dd or General to clear formatting.")] string NumberFormat);

/// <summary>
/// Maps a workbook's structure by reading the file directly - no Excel process at all.
///
/// An .xlsx or .xlsm is a ZIP of XML, and everything structural is legible from it: sheets,
/// dimensions, and per-cell whether a formula or a constant put the value there. Measured against
/// the same workbook, the direct read answered in under a third of the audit's time with zero
/// Excel launches - and the trace had already shown that on small tasks 92% of wall time was
/// Excel teardown and verification, none of which a scan pays.
///
/// The reason it exists is planning. The mixed-column report - a column that is mostly formulas
/// with a scattering of constants - is the shape of a manual override sitting inside a calculated
/// column, which is precisely what a caller needs to see before deciding what to fix and where.
/// Twenty thousand rows with 37 hardcoded cells is a fact this returns in one fast call and that
/// no bounded read could affordably discover.
///
/// It reports shape, never contents: counts, dimensions, and row numbers - no cell values, no
/// formula text. Stored results reflect the file's last save, which for structure is exact.
/// </summary>
public sealed record ScanWorkbookStructureOperation();

/// <summary>Manual closed union for the operation selected by the one Excel task.</summary>
public sealed record ExcelOperation(
    [property: Description("Selects which one operation payload is supplied.")] ExcelOperationKind Kind,
    [property: Description("Required only when kind is CopyExhibit; all other payloads must be null.")] CopyExhibitOperation? CopyExhibit = null,
    [property: Description("Required only when kind is RepairExistingWorksheet; all other payloads must be null.")] RepairExistingWorksheetOperation? RepairExistingWorksheet = null,
    [property: Description("Required only when kind is ExtendFormulaSeries; all other payloads must be null.")] ExtendFormulaSeriesOperation? ExtendFormulaSeries = null,
    [property: Description("Required only when kind is EditMacroProcedure; all other payloads must be null.")] EditMacroProcedureOperation? EditMacroProcedure = null,
    [property: Description("Required only when kind is AuditWorkbookFlows; all other payloads must be null. It takes no options, so supply the empty object {} - omitting it entirely is rejected. The read-only report lists worksheets, tables, defined names, queries, connections, macro components and procedures, the data model, pivots, and external links.")] AuditWorkbookFlowsOperation? AuditWorkbookFlows = null,
    [property: Description("Required only when kind is ReadWorksheetRange; all other payloads must be null. Reads one bounded range and returns its contents.")] ReadWorksheetRangeOperation? ReadWorksheetRange = null,
    [property: Description("Required only when kind is WriteWorksheetValues; all other payloads must be null. Writes constants into named cells and reads them back. Never accepts formula text.")] WriteWorksheetValuesOperation? WriteWorksheetValues = null,
    [property: Description("Required only when kind is FindReplace; all other payloads must be null. Plan lists the matching cells and changes nothing - also how to locate a known label; Apply rewrites the constants among them.")] FindReplaceOperation? FindReplace = null,
    [property: Description("Required only when kind is Create; all other payloads must be null. Creates an empty workbook or adds an empty worksheet.")] CreateOperation? Create = null,
    [property: Description("Required only when kind is SetNumberFormat; all other payloads must be null. Sets how a range displays its numbers. Plan reports the range's current format and changes nothing. It changes no cell values, and it sets nothing else: no fonts, fills, borders, widths, or conditional formats.")] SetNumberFormatOperation? SetNumberFormat = null,
    [property: Description("Required only when kind is ScanWorkbookStructure; all other payloads must be null, and supply the empty object {}. Reads the file directly without starting Excel - fast on any size. Reports each sheet's dimension and formula/constant counts, defined names, tables, and external links by file name, and flags mostly-formula columns holding scattered constants: the shape of a manual override. Counts only, never contents. It reports nothing about macros, queries, connections, or the data model, which are unreadable without Excel - absence here is not evidence they are absent, so use AuditWorkbookFlows when those matter. Use it before deciding what to read or fix.")] ScanWorkbookStructureOperation? ScanWorkbookStructure = null);

public sealed record ExcelTaskRequest(
    [property: Description("Target workbook path, ending .xlsx or .xlsm. It must already exist for every operation except Create with kind Workbook, where it must not.")] string TargetWorkbookPath,
    [property: Description("The required manual operation union. Supply exactly one payload matching kind.")] ExcelOperation Operation,
    [property: Description("Plan previews without mutation; Apply performs the task after required confirmations.")] ExcelTaskMode Mode = ExcelTaskMode.Apply,
    [property: Description("Use AskIfOpen when the workbook state is unknown; resubmit with UseOpen or Isolated if confirmation is returned. When the request already says the workbook is open, use UseOpen directly to avoid a wasted round trip. EditMacroProcedure and Create require Isolated. UseOpen cannot be combined with Copy, and Isolated with save Same is rejected while the target is open in Excel.")] WorkbookBinding WorkbookBinding = WorkbookBinding.AskIfOpen,
    [property: Description("Same saves to the target; Copy saves only to outputWorkbookPath. EditMacroProcedure requires Copy to an .xlsm path. Copy is rejected with UseOpen. AuditWorkbookFlows, ReadWorksheetRange, ScanWorkbookStructure and Create never take a save destination: leave Same with no outputWorkbookPath.")] SaveMode Save = SaveMode.Same,
    [property: Description("Required destination path when save is Copy; omit for Same. Must differ from the target path and carry the target's extension.")] string? OutputWorkbookPath = null,
    [property: Description("Explicit authorization to overwrite. Apply with save Same requires it, because saving in place overwrites the target workbook itself. Apply with save Copy requires it only when outputWorkbookPath already exists. Plan never requires it, and neither does Create with kind Workbook, which is refused outright if anything is already at the path.")] bool OverwriteConfirmed = false);

/// <summary>
/// What inspection needs to know before execution. <paramref name="TargetMustExist"/> is false for
/// the one operation whose whole purpose is a target that does not exist yet - creating a workbook -
/// and true for every other, so a missing file stays a clean rejection everywhere else.
/// </summary>
public sealed record WorkbookInspectionRequest(
    string TargetWorkbookPath,
    string? ReferenceWorkbookPath,
    WorkbookBinding Binding,
    SaveMode Save,
    string? OutputWorkbookPath,
    bool TargetMustExist = true);

/// <summary>
/// What inspection learned before execution. <paramref name="InfeasibleReason"/> carries the
/// caller-actionable reason the task cannot run at all - a target that does not exist, a copy
/// output whose directory is missing. It used to be thrown instead, and the engine's catch-all
/// turned "Target workbook does not exist" into "Workbook inspection could not be completed before
/// execution" - an infrastructure-sounding answer to the most ordinary user error there is, a
/// mistyped path. A reason is a finding, not a failure, so it travels as data.
/// </summary>
public sealed record WorkbookInspection(bool TargetIsOpen, bool CopyOutputExists = false, string? OpenWorkbookDescription = null, IReadOnlyList<TaskCheck>? Checks = null, string? InfeasibleReason = null);

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

public sealed record NormalizedFindReplaceOperation(string WorksheetName, string Find, string? ReplaceWith, FormulaRepairRange? Range, bool WholeCell, bool MatchCase);

public sealed record NormalizedCreateOperation(CreateKind Kind, string? WorksheetName);

public sealed record NormalizedSetNumberFormatOperation(string WorksheetName, FormulaRepairRange Range, string NumberFormat);

public sealed record NormalizedScanWorkbookStructureOperation();

/// <summary>Validated internal counterpart of <see cref="ExcelOperation"/>. It contains no legacy flat request fields.</summary>
public sealed record NormalizedExcelOperation(
    ExcelOperationKind Kind,
    NormalizedCopyExhibitOperation? CopyExhibit = null,
    NormalizedRepairExistingWorksheetOperation? RepairExistingWorksheet = null,
    NormalizedExtendFormulaSeriesOperation? ExtendFormulaSeries = null,
    NormalizedEditMacroProcedureOperation? EditMacroProcedure = null,
    NormalizedAuditWorkbookFlowsOperation? AuditWorkbookFlows = null,
    NormalizedReadWorksheetRangeOperation? ReadWorksheetRange = null,
    NormalizedWriteWorksheetValuesOperation? WriteWorksheetValues = null,
    NormalizedFindReplaceOperation? FindReplace = null,
    NormalizedCreateOperation? Create = null,
    NormalizedSetNumberFormatOperation? SetNumberFormat = null,
    NormalizedScanWorkbookStructureOperation? ScanWorkbookStructure = null);

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

/// <summary>
/// One cell that had something in it. Blanks are omitted rather than reported as empty.
///
/// <paramref name="IsFormula"/> says whether a formula put the text there. Without it the only way
/// to know was to read the same range twice - once for values, once for formulas - and diff the
/// results, which a UX simulation did, spending a whole Excel launch to find out whether a cell it
/// was about to overwrite was a formula. It costs one extra array read per range and answers the
/// question every caller about to write already has.
/// </summary>
public sealed record WorksheetCell(string Address, string Text, bool IsFormula = false);

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

/// <summary>
/// What a runtime hands back. One fact of this interface was previously stated nowhere: the engine
/// coerces <paramref name="CanRetry"/> and <paramref name="RetryReason"/> by status - Rejected is
/// always retryable, Unknown and Partial never are, with the engine's own reasons - so a runtime's
/// values are honoured only for the statuses in between. The retry policy belongs to the engine;
/// what a runtime supplies is advisory and, today, decorative for every status a mutation path
/// actually produces.
/// </summary>
public sealed record WorkbookExecutionOutcome(ExcelTaskStatus Status, string Summary, IReadOnlyList<TaskChange>? Changes = null, IReadOnlyList<TaskCheck>? Checks = null, bool CanRetry = false, string? RetryReason = null, MacroProcedureReceipt? MacroProcedure = null, WorkbookAuditReceipt? Audit = null, WorksheetRangeReceipt? Range = null);
public sealed record SaveReceipt(SaveMode Mode, string? OutputWorkbookPath, bool OverwriteConfirmed);
public sealed record RetryReceipt(bool CanRetry, string? Reason);
public sealed record ConfirmationRequirement(string Code, string Prompt);
public sealed record ConfirmationReceipt(bool Required, IReadOnlyList<ConfirmationRequirement> Requirements);
public sealed record PhaseTimings(TimeSpan Validation, TimeSpan Inspection, TimeSpan Execution, TimeSpan Total);
public sealed record ExcelTaskReceipt(string TaskId, ExcelTaskStatus Status, string Summary, IReadOnlyList<TaskChange> Changes, IReadOnlyList<TaskCheck> Checks, SaveReceipt Save, RetryReceipt Retry, ConfirmationReceipt Confirmation, PhaseTimings Timings, MacroProcedureReceipt? MacroProcedure = null, WorkbookAuditReceipt? Audit = null, WorksheetRangeReceipt? Range = null);
