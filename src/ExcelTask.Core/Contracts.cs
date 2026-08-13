using System.ComponentModel;

namespace ExcelTask.Core;

public enum ExcelTaskMode { Plan, Apply }
public enum WorkbookBinding { AskIfOpen, UseOpen, Isolated }
public enum SaveMode { Same, Copy }
public enum ExcelTaskStatus { Planned, NeedsConfirmation, Completed, Rejected, Partial, Unknown }
public enum ExcelOperationKind { CopyExhibit, RepairExistingWorksheet, ExtendFormulaSeries, EditMacroProcedure, AuditWorkbookFlows, ReadWorksheetRange, WriteWorksheetValues, WriteWorksheetFormulas, FindReplace, Create, SetRangeFormat, ScanWorkbookStructure, ManageTable, ManageQuery, ManageModelMeasure, ManageModelRelationship }
public enum CreateKind { Workbook, Worksheet }

/// <summary>What a table request does. ConvertToRange keeps every cell and drops only the table.</summary>
public enum TableAction { Create, Rename, Restyle, Resize, ConvertToRange }

/// <summary>What a Power Query request does.</summary>
public enum QueryAction { Create, Replace, Delete }
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
    [property: Description("Existing worksheet name to read. If unknown, run ScanWorkbookStructure - it names them and starts no Excel.")] string WorksheetName,
    [property: Description("One bounded A1 range to read, at most 400 cells. Narrow the range and read again if more is needed.")] string Range,
    [property: Description("False returns displayed values; true returns R1C1 formulas instead, for comparing a formula pattern.")] bool Formulas = false);

/// <summary>
/// Writes constants into named cells.
///
/// This operation refuses model-written formula text, and that refusal is not relaxed here. The two
/// are not the same risk. A formula composed by a model can be plausible and silently wrong - off
/// by a row, anchored where it should be relative - and a receipt saying "it was written" would
/// confirm nothing about whether it was right. That is why ExtendFormulaSeries and
/// RepairExistingWorksheet infer formulas from evidence already in the sheet, while the separate
/// WriteWorksheetFormulas operation is the explicit opt-in for callers that really do have formula
/// text. A constant has no such failure mode: it is exactly what the caller named, and reading it
/// back proves it character for character.
///
/// Anything beginning with "=" is rejected rather than written as text, because silently storing a
/// formula as a label is the kind of quiet wrongness this whole design exists to avoid; use the
/// separate formula-write operation when a formula is intentional.
/// </summary>
public sealed record WriteWorksheetValuesOperation(
    [property: Description("Existing worksheet name to write to. If unknown, run ScanWorkbookStructure - it names them and starts no Excel.")] string WorksheetName,
    [property: Description("Cells to write, at most 200. Each address must fall inside one bounded A1 range no larger than 400 cells.")] IReadOnlyList<WorksheetCellValue> Cells);

/// <summary>One constant to write. Text starting with = is rejected; this operation never writes formulas.</summary>
public sealed record WorksheetCellValue(
    [property: Description("Single A1 cell address, such as B7.")] string Address,
    [property: Description("Constant to write. Numbers and TRUE/FALSE are written as numbers and booleans, everything else as text. Must not start with =.")] string Value);

/// <summary>
/// Writes caller-supplied A1 formulas into named cells.
///
/// Formula text is deliberately a separate operation from <see cref="WriteWorksheetValuesOperation"/>
/// so a model cannot turn a constant write into a formula by accident. Formula writes are accepted
/// only when they begin with '=' and stay inside the same bounded cell/span limits as value writes.
/// Apply reads each formula back before saving and verifies the saved formulas after reopening; the
/// receipt reports counts and fingerprints, never the formula text itself. When a formula can be
/// inferred from neighbouring evidence, ExtendFormulaSeries or RepairExistingWorksheet remains the
/// stronger choice because it can prove the intended pattern rather than only the stored formula.
/// </summary>
public sealed record WriteWorksheetFormulasOperation(
    [property: Description("Worksheet.")] string WorksheetName,
    [property: Description("200 A1 formulas; starts =; max 8,192 chars; span 400; UTF-8 768 KiB.")] IReadOnlyList<WorksheetCellFormula> Cells);

/// <summary>One caller-supplied A1 formula. Formula text is stored in the workbook but withheld from receipts.</summary>
public sealed record WorksheetCellFormula(
    [property: Description("A1 cell.")] string Address,
    [property: Description("Starts =; never returned.")] string Formula);

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
    [property: Description("Existing worksheet name to search. If unknown, run ScanWorkbookStructure - it names them and starts no Excel.")] string WorksheetName,
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
/// It shipped as number format alone, deliberately, because the measured demand named the tool and
/// not which of its operations were used - and building all of a tool on a tool-level count is how a
/// server ends up with 230 operations and no idea which matter. The owner has since asked for the
/// rest by name, which is the evidence that entry was waiting for.
///
/// Every appearance field is optional and independent. Omitted means "leave it as it is", which
/// matters more here than anywhere else in the product: formatting is the one mutation with no
/// recoverable prior state on the sheet, so a caller changing a fill must not have to restate a
/// font it never looked at. At least one field must be supplied, because a request that changes
/// nothing is a mistake worth naming rather than a no-op worth performing.
///
/// Still absent, and still on the same evidence rule: conditional formats, merged cells, alignment,
/// and styles by name.
/// </summary>
public sealed record SetRangeFormatOperation(
    [property: Description("Existing worksheet name to format. If unknown, run ScanWorkbookStructure - it names them and starts no Excel.")] string WorksheetName,
    [property: Description("One bounded A1 range to format, at most 10,000 cells. Formatting applies to every cell in it, including blank ones.")] string Range,
    [property: Description("Excel number format code in US-English form, at most 255 characters, such as #,##0.00 or #,##0;(#,##0) or 0.0% or yyyy-mm-dd or General to clear formatting.")] string? NumberFormat = null,
    [property: Description("True bolds, false clears bold; omit to leave it as it is.")] bool? Bold = null,
    [property: Description("True italicises, false clears it; omit to leave it as it is.")] bool? Italic = null,
    [property: Description("Font point size, 1 to 409.")] double? FontSize = null,
    [property: Description("Font name. Excel stores whatever it is given and renders a substitute when the font is not installed, so a wrong name cannot be detected here - check the spelling.")] string? FontName = null,
    [property: Description("Text colour as #RRGGBB.")] string? FontColor = null,
    [property: Description("Cell background as #RRGGBB, or None to clear it.")] string? FillColor = null,
    [property: Description("Which edges to draw: All, Outline, Top, Bottom, Left, Right, or None to clear them.")] string? Borders = null,
    [property: Description("Border weight, used only with borders: Hairline, Thin, Medium, or Thick. Thin when omitted.")] string? BorderStyle = null,
    [property: Description("Column width in characters, 0 to 255. Applies to every column the range touches, not only the cells in it.")] double? ColumnWidth = null,
    [property: Description("Row height in points, 0 to 409. Applies to every row the range touches, not only the cells in it.")] double? RowHeight = null);

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

/// <summary>
/// Creates or changes one Excel table (a ListObject), beyond the listing the audit and the scan
/// already give.
///
/// A table is the unit a lot of real work is organised around - structured references, a growing
/// range that formulas follow, a style that survives inserted rows - and until now this server could
/// tell you one existed and nothing else. Every action here is reversible by another call except
/// ConvertToRange, which is why that one keeps the cells and only drops the table over them.
///
/// Deliberately absent: adding or removing columns, and sorting or filtering. Those change what the
/// data says rather than how it is organised, and belong with the operations that already prove a
/// value round trip rather than here.
/// </summary>
public sealed record ManageTableOperation(
    [property: Description("Existing worksheet name holding the table, or where a new one goes.")] string WorksheetName,
    [property: Description("Create, Rename, Restyle, Resize, or ConvertToRange - which keeps every cell and drops only the table over them.")] TableAction Action,
    [property: Description("The table to act on, as the audit or scan reports it. For Create, the name the new table takes.")] string TableName,
    [property: Description("A1 range for Create and Resize, including the header row. Ignored by the others.")] string? Range = null,
    [property: Description("The new name, for Rename only. Must not already be in use.")] string? NewName = null,
    [property: Description("Style name for Create and Restyle, such as TableStyleMedium2, or None for an unstyled table.")] string? TableStyle = null);

/// <summary>
/// Creates, replaces, or deletes one Power Query, guarded the way a macro edit is.
///
/// A query decides where a workbook's data comes from, which makes replacing one at least as
/// consequential as replacing a macro procedure: a plausible-looking M expression pointing at the
/// wrong source produces a workbook full of numbers that are all wrong and all confident. So it
/// borrows the macro operation's precondition exactly. Plan reports the current query's fingerprint,
/// and Apply must carry that fingerprint back, which fails if anything changed in between.
///
/// Plan returns the current expression alongside the fingerprint, the way macro Plan returns
/// bounded VBA source. It did not until 0.20.0. An M expression very often names a server and a
/// database and sometimes a key - Sql.Database(...) and Web.Contents(...) are the ordinary way to
/// write one - so the expression was withheld and the caller was told to go read it in Excel. That
/// withheld it from the one reader that has to understand a query in order to replace it correctly,
/// and it made every replacement a blind edit guarded only by a hash. The owner asked for the text;
/// PRIVACY.md section 4 states plainly what now crosses the boundary.
///
/// It is bounded the way macro source is: at most 8,192 characters, and OMITTED rather than
/// truncated past that. Half an M expression reads exactly like a whole one, and a caller that sent
/// one back as a replacement would destroy the query it meant to edit.
/// </summary>
public sealed record ManageQueryOperation(
    [property: Description("The query to act on, as AuditWorkbookFlows lists it. For Create, the name the new query takes.")] string QueryName,
    [property: Description("Create, Replace, or Delete. Delete removes the definition and leaves any worksheet it already loaded to.")] QueryAction Action,
    [property: Description("Apply only, for Create and Replace: the complete M expression, at most 8,192 characters, as the Power Query editor shows it.")] string? Formula = null,
    [property: Description("Apply only for Replace and Delete, and must be omitted for Plan: SHA-256 fingerprint of the query being changed, taken from the Plan receipt.")] string? ExpectedFormulaSha256 = null);

/// <summary>
/// Creates, replaces, or deletes one Data Model measure, guarded by the same fingerprint the query
/// and macro operations use.
///
/// A measure is a DAX expression every pivot over the model reads through, so a wrong one is wrong
/// everywhere at once and silently - which is why this carries a precondition rather than trusting
/// the name alone. Unlike an M expression, DAX names columns in the model rather than servers, so
/// the Plan receipt reports the expression as well as its fingerprint; there is nothing here that
/// the audit would refuse to return.
///
/// Only measures. Model tables come from loading a query into the model - use ManageQuery - and
/// relationships are not here yet, because a wrong relationship silently changes every number the
/// model produces and the operation that adds one should be able to show what it would join before
/// it does it.
/// </summary>
public sealed record ManageModelMeasureOperation(
    [property: Description("Data Model table the measure belongs to, as AuditWorkbookFlows lists it. Load a query into the model with ManageQuery first if there is none.")] string TableName,
    [property: Description("The measure to create, replace, or delete.")] string MeasureName,
    [property: Description("Create, Replace, or Delete.")] QueryAction Action,
    [property: Description("Apply only, for Create and Replace: the DAX expression, with NO leading equals sign, at most 8,192 characters. For example SUM(Sales[Amount]).")] string? Formula = null,
    [property: Description("Apply only for Replace and Delete, and must be omitted for Plan: SHA-256 fingerprint of the measure being changed, taken from the Plan receipt.")] string? ExpectedFormulaSha256 = null);

/// <summary>
/// Creates or deletes one Data Model relationship.
///
/// A relationship decides how every pivot over the model aggregates, so a wrong one is wrong
/// everywhere at once and produces numbers rather than errors. There is no expression to fingerprint
/// here, so the precondition is the Plan itself: it confirms both tables and both columns exist and
/// lists the relationships already joining them, so the join can be read before it is made. Replace
/// is refused by name - a relationship is not edited in place, and delete-then-create says so.
/// </summary>
public sealed record ManageModelRelationshipOperation(
    [property: Description("Data Model table on the MANY side - the fact table, whose column repeats.")] string FromTable,
    [property: Description("Column on the many side, the foreign key.")] string FromColumn,
    [property: Description("Data Model table on the ONE side - the lookup table, whose column is unique. Excel refuses the relationship if it is not.")] string ToTable,
    [property: Description("Column on the one side, the primary key.")] string ToColumn,
    [property: Description("Create or Delete. Replace is rejected: delete the relationship and create the new one.")] QueryAction Action);

/// <summary>Manual closed union for the operation selected by the one Excel task.</summary>
public sealed record ExcelOperation(
    [property: Description("Selects which one operation payload is supplied.")] ExcelOperationKind Kind,
    [property: Description("Required when kind is CopyExhibit.")] CopyExhibitOperation? CopyExhibit = null,
    [property: Description("Required when kind is RepairExistingWorksheet.")] RepairExistingWorksheetOperation? RepairExistingWorksheet = null,
    [property: Description("Required when kind is ExtendFormulaSeries.")] ExtendFormulaSeriesOperation? ExtendFormulaSeries = null,
    [property: Description("Required when kind is EditMacroProcedure.")] EditMacroProcedureOperation? EditMacroProcedure = null,
    [property: Description("Required when kind is AuditWorkbookFlows. It takes no options, so supply the empty object {} - omitting it entirely is rejected. The read-only report lists worksheets, tables, defined names, queries, connections, macro components and procedures, the data model, pivots, and external links.")] AuditWorkbookFlowsOperation? AuditWorkbookFlows = null,
    [property: Description("Required when kind is ReadWorksheetRange. Reads one bounded range and returns its contents.")] ReadWorksheetRangeOperation? ReadWorksheetRange = null,
    [property: Description("Required when kind is WriteWorksheetValues. Writes constants into named cells and reads them back. Never accepts formula text.")] WriteWorksheetValuesOperation? WriteWorksheetValues = null,
    [property: Description("Formula-write payload.")] WriteWorksheetFormulasOperation? WriteWorksheetFormulas = null,
    [property: Description("Required when kind is FindReplace. Plan lists the matching cells and changes nothing - also how to locate a known label; Apply rewrites the constants among them.")] FindReplaceOperation? FindReplace = null,
    [property: Description("Required when kind is Create. Creates an empty workbook or adds an empty worksheet.")] CreateOperation? Create = null,
    [property: Description("Required when kind is SetRangeFormat. Sets how a range looks: number format, bold, italic, font size/name/colour, fill, borders, column width, row height. Every field is optional and independent - omit one to leave it as it is - and at least one must be supplied. Plan reports what is there now and changes nothing. It never changes a cell value.")] SetRangeFormatOperation? SetRangeFormat = null,
    [property: Description("Required when kind is ScanWorkbookStructure; supply the empty object {}. Reads the file directly without starting Excel - fast on any size. Reports each sheet's dimension and formula/constant counts, defined names, tables, and external links by file name, and flags mostly-formula columns holding scattered constants: the shape of a manual override. Counts only, never contents. It reports nothing about macros, queries, connections, or the data model, which are unreadable without Excel; absence here is not evidence, so use AuditWorkbookFlows when those matter.")] ScanWorkbookStructureOperation? ScanWorkbookStructure = null,
    [property: Description("Required when kind is ManageTable. Creates or changes one Excel table: create over a range, rename, restyle, resize, or convert back to plain cells. Plan reports it as it is now and changes nothing.")] ManageTableOperation? ManageTable = null,
    [property: Description("Required when kind is ManageQuery. Creates, replaces, or deletes one Power Query. Plan reports the query''s current M expression and fingerprint; Apply must carry the fingerprint back.")] ManageQueryOperation? ManageQuery = null,
    [property: Description("Required when kind is ManageModelMeasure. Creates, replaces, or deletes one Data Model measure. Plan reports the DAX and its fingerprint; Apply must carry the fingerprint back.")] ManageModelMeasureOperation? ManageModelMeasure = null,
    [property: Description("Required when kind is ManageModelRelationship. Creates or deletes one Data Model relationship, from the many side to the one side. Plan names the join and lists what already joins those tables.")] ManageModelRelationshipOperation? ManageModelRelationship = null);

public sealed record ExcelTaskRequest(
    [property: Description("Target workbook path, ending .xlsx or .xlsm. It must already exist for every operation except Create with kind Workbook, where it must not.")] string TargetWorkbookPath,
    [property: Description("The required manual operation union. Supply exactly one payload matching kind.")] ExcelOperation Operation,
    [property: Description("Plan previews without mutation; Apply performs the task after required confirmations.")] ExcelTaskMode Mode = ExcelTaskMode.Apply,
    [property: Description("AskIfOpen when unknown; resubmit UseOpen or Isolated if confirmation returns. Use UseOpen when the workbook is known open. EditMacroProcedure/Create require Isolated. UseOpen+Copy and Isolated+Same while open are rejected.")] WorkbookBinding WorkbookBinding = WorkbookBinding.AskIfOpen,
    [property: Description("Same saves the target; Copy saves outputWorkbookPath. Macro editing requires .xlsm Copy. UseOpen+Copy is rejected. Audit, Read, Scan, and Create never take output; leave Same.")] SaveMode Save = SaveMode.Same,
    [property: Description("Required destination path when save is Copy; omit for Same. Must differ from the target path and carry the target's extension.")] string? OutputWorkbookPath = null,
    [property: Description("Explicit overwrite authorization. Apply+Same requires true; Apply+Copy requires true only when output exists. Plan and Create Workbook do not require it; Create Workbook refuses an existing file.")] bool OverwriteConfirmed = false);

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

public sealed record NormalizedWorksheetCellFormula(string Address, string Formula);

public sealed record NormalizedWriteWorksheetFormulasOperation(string WorksheetName, IReadOnlyList<NormalizedWorksheetCellFormula> Cells);

public sealed record NormalizedFindReplaceOperation(string WorksheetName, string Find, string? ReplaceWith, FormulaRepairRange? Range, bool WholeCell, bool MatchCase);

public sealed record NormalizedCreateOperation(CreateKind Kind, string? WorksheetName);

public sealed record NormalizedSetRangeFormatOperation(string WorksheetName, FormulaRepairRange Range, string? NumberFormat, bool? Bold, bool? Italic, double? FontSize, string? FontName, int? FontColor, int? FillColor, RangeBorderEdges Borders, RangeBorderWeight BorderStyle, double? ColumnWidth, double? RowHeight);

/// <summary>Which edges a border request draws. None clears them.</summary>
public enum RangeBorderEdges { Unspecified, None, All, Outline, Top, Bottom, Left, Right }

/// <summary>Border weight, mapped to Excel's own weights rather than exposing its numbers.</summary>
public enum RangeBorderWeight { Thin, Hairline, Medium, Thick }

public sealed record NormalizedScanWorkbookStructureOperation();

public sealed record NormalizedManageTableOperation(string WorksheetName, TableAction Action, string TableName, FormulaRepairRange? Range, string? NewName, string? TableStyle);

public sealed record NormalizedManageQueryOperation(string QueryName, QueryAction Action, string? Formula, string? ExpectedFormulaSha256);

public sealed record NormalizedManageModelMeasureOperation(string TableName, string MeasureName, QueryAction Action, string? Formula, string? ExpectedFormulaSha256);

public sealed record NormalizedManageModelRelationshipOperation(string FromTable, string FromColumn, string ToTable, string ToColumn, QueryAction Action);

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
    NormalizedWriteWorksheetFormulasOperation? WriteWorksheetFormulas = null,
    NormalizedFindReplaceOperation? FindReplace = null,
    NormalizedCreateOperation? Create = null,
    NormalizedSetRangeFormatOperation? SetRangeFormat = null,
    NormalizedScanWorkbookStructureOperation? ScanWorkbookStructure = null,
    NormalizedManageTableOperation? ManageTable = null,
    NormalizedManageQueryOperation? ManageQuery = null,
    NormalizedManageModelMeasureOperation? ManageModelMeasure = null,
    NormalizedManageModelRelationshipOperation? ManageModelRelationship = null);

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
/// One Power Query as Plan found it. <paramref name="Formula"/> is the stored M expression, present
/// on Plan and null on Apply - the same split macro source uses, because Apply's answer is what
/// changed, not what was there.
///
/// Null also means "too long to send whole": past 8,192 characters the expression is omitted rather
/// than cut, so <paramref name="Length"/> is the field that always tells the truth about size.
/// </summary>
public sealed record QueryReceipt(string QueryName, string Sha256, int Length, string? Formula);

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
public sealed record WorkbookExecutionOutcome(ExcelTaskStatus Status, string Summary, IReadOnlyList<TaskChange>? Changes = null, IReadOnlyList<TaskCheck>? Checks = null, bool CanRetry = false, string? RetryReason = null, MacroProcedureReceipt? MacroProcedure = null, WorkbookAuditReceipt? Audit = null, WorksheetRangeReceipt? Range = null, QueryReceipt? Query = null);
public sealed record SaveReceipt(SaveMode Mode, string? OutputWorkbookPath, bool OverwriteConfirmed);
public sealed record RetryReceipt(bool CanRetry, string? Reason);
public sealed record ConfirmationRequirement(string Code, string Prompt);
public sealed record ConfirmationReceipt(bool Required, IReadOnlyList<ConfirmationRequirement> Requirements);
public sealed record PhaseTimings(TimeSpan Validation, TimeSpan Inspection, TimeSpan Execution, TimeSpan Total);
public sealed record ExcelTaskReceipt(string TaskId, ExcelTaskStatus Status, string Summary, IReadOnlyList<TaskChange> Changes, IReadOnlyList<TaskCheck> Checks, SaveReceipt Save, RetryReceipt Retry, ConfirmationReceipt Confirmation, PhaseTimings Timings, MacroProcedureReceipt? MacroProcedure = null, WorkbookAuditReceipt? Audit = null, WorksheetRangeReceipt? Range = null, QueryReceipt? Query = null);
