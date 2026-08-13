using ExcelTask.Core;

namespace ExcelTask.Excel.Tests;

internal static class ExcelTaskPlans
{
    public static NormalizedExcelTaskRequest Audit(
        string target,
        ExcelTaskMode mode = ExcelTaskMode.Apply,
        WorkbookBinding binding = WorkbookBinding.Isolated) => new(
        target,
        mode,
        binding,
        SaveMode.Same,
        null,
        false,
        new NormalizedExcelOperation(
            ExcelOperationKind.AuditWorkbookFlows,
            AuditWorkbookFlows: new NormalizedAuditWorkbookFlowsOperation()));

    public static NormalizedExcelTaskRequest Read(
        string target,
        string worksheet,
        string range,
        bool formulas = false,
        ExcelTaskMode mode = ExcelTaskMode.Apply,
        WorkbookBinding binding = WorkbookBinding.Isolated) => new(
        target,
        mode,
        binding,
        SaveMode.Same,
        null,
        false,
        new NormalizedExcelOperation(
            ExcelOperationKind.ReadWorksheetRange,
            ReadWorksheetRange: new NormalizedReadWorksheetRangeOperation(worksheet, ToRange(range), formulas)));

    public static NormalizedExcelTaskRequest Write(
        string target,
        string worksheet,
        (string Address, string Value)[] cells,
        ExcelTaskMode mode = ExcelTaskMode.Apply,
        SaveMode save = SaveMode.Same,
        string? output = null,
        WorkbookBinding binding = WorkbookBinding.Isolated) => new(
        target,
        mode,
        binding,
        save,
        output,
        mode == ExcelTaskMode.Apply,
        new NormalizedExcelOperation(
            ExcelOperationKind.WriteWorksheetValues,
            WriteWorksheetValues: new NormalizedWriteWorksheetValuesOperation(
                worksheet,
                [.. cells.Select(cell => new NormalizedWorksheetCellValue(cell.Address, cell.Value))])));

    public static NormalizedExcelTaskRequest WriteFormulas(
        string target,
        string worksheet,
        (string Address, string Formula)[] cells,
        ExcelTaskMode mode = ExcelTaskMode.Apply,
        SaveMode save = SaveMode.Same,
        string? output = null,
        WorkbookBinding binding = WorkbookBinding.Isolated) => new(
        target,
        mode,
        binding,
        save,
        output,
        mode == ExcelTaskMode.Apply,
        new NormalizedExcelOperation(
            ExcelOperationKind.WriteWorksheetFormulas,
            WriteWorksheetFormulas: new NormalizedWriteWorksheetFormulasOperation(
                worksheet,
                [.. cells.Select(cell => new NormalizedWorksheetCellFormula(cell.Address, cell.Formula))])));

    public static NormalizedExcelTaskRequest FindReplace(
        string target,
        string worksheet,
        string find,
        string? replaceWith = null,
        string? range = null,
        bool wholeCell = false,
        bool matchCase = false,
        ExcelTaskMode mode = ExcelTaskMode.Apply,
        SaveMode save = SaveMode.Same,
        string? output = null,
        WorkbookBinding binding = WorkbookBinding.Isolated) => new(
        target,
        mode,
        binding,
        save,
        output,
        mode == ExcelTaskMode.Apply,
        new NormalizedExcelOperation(
            ExcelOperationKind.FindReplace,
            FindReplace: new NormalizedFindReplaceOperation(
                worksheet, find, replaceWith, range is null ? null : ToRange(range), wholeCell, matchCase)));

    public static NormalizedExcelTaskRequest NumberFormat(
        string target,
        string worksheet,
        string range,
        string numberFormat,
        ExcelTaskMode mode = ExcelTaskMode.Apply,
        WorkbookBinding binding = WorkbookBinding.Isolated) => new(
        target,
        mode,
        binding,
        SaveMode.Same,
        null,
        mode == ExcelTaskMode.Apply,
        new NormalizedExcelOperation(
            ExcelOperationKind.SetNumberFormat,
            SetNumberFormat: new NormalizedSetNumberFormatOperation(worksheet, ToRange(range), numberFormat)));

    public static NormalizedExcelTaskRequest Scan(
        string target,
        ExcelTaskMode mode = ExcelTaskMode.Plan) => new(
        target,
        mode,
        WorkbookBinding.Isolated,
        SaveMode.Same,
        null,
        false,
        new NormalizedExcelOperation(
            ExcelOperationKind.ScanWorkbookStructure,
            ScanWorkbookStructure: new NormalizedScanWorkbookStructureOperation()));

    public static NormalizedExcelTaskRequest Create(
        string target,
        CreateKind kind,
        string? worksheet = null,
        ExcelTaskMode mode = ExcelTaskMode.Apply) => new(
        target,
        mode,
        WorkbookBinding.Isolated,
        SaveMode.Same,
        null,
        mode == ExcelTaskMode.Apply,
        new NormalizedExcelOperation(
            ExcelOperationKind.Create,
            Create: new NormalizedCreateOperation(kind, worksheet)));

    public static NormalizedExcelTaskRequest Copy(
        string target,
        string reference,
        string referenceWorksheet,
        string newWorksheet,
        IReadOnlyList<FormulaRepairRange> repairRanges,
        ExcelTaskMode mode,
        WorkbookBinding binding,
        SaveMode save,
        string? output,
        bool overwrite) => new(
        target,
        mode,
        binding,
        save,
        output,
        overwrite,
        new NormalizedExcelOperation(
            ExcelOperationKind.CopyExhibit,
            new NormalizedCopyExhibitOperation(reference, referenceWorksheet, newWorksheet, repairRanges)));

    public static NormalizedExcelTaskRequest Repair(
        string target,
        string worksheet,
        IReadOnlyList<FormulaRepairRange> ranges,
        ExcelTaskMode mode,
        WorkbookBinding binding,
        SaveMode save = SaveMode.Same) => new(
        target,
        mode,
        binding,
        save,
        null,
        mode == ExcelTaskMode.Apply,
        new NormalizedExcelOperation(
            ExcelOperationKind.RepairExistingWorksheet,
            RepairExistingWorksheet: new NormalizedRepairExistingWorksheetOperation(worksheet, ranges)));

    public static NormalizedExcelTaskRequest Extend(
        string target,
        string worksheet,
        FormulaExtensionDirection direction,
        string evidence,
        string destination,
        ExcelTaskMode mode,
        WorkbookBinding binding) => new(
        target,
        mode,
        binding,
        SaveMode.Same,
        null,
        mode == ExcelTaskMode.Apply,
        new NormalizedExcelOperation(
            ExcelOperationKind.ExtendFormulaSeries,
            ExtendFormulaSeries: new NormalizedExtendFormulaSeriesOperation(
                worksheet,
                direction,
                ToRange(evidence),
                ToRange(destination))));

    public static NormalizedExcelTaskRequest Macro(
        string target,
        string output,
        string component,
        string procedure,
        ExcelTaskMode mode,
        string? expectedHash = null,
        string? replacementSource = null,
        bool runAfterEdit = false) => new(
        target,
        mode,
        WorkbookBinding.Isolated,
        SaveMode.Copy,
        output,
        mode == ExcelTaskMode.Apply,
        new NormalizedExcelOperation(
            ExcelOperationKind.EditMacroProcedure,
            EditMacroProcedure: new NormalizedEditMacroProcedureOperation(
                component,
                procedure,
                expectedHash,
                replacementSource,
                runAfterEdit)));

    private static FormulaRepairRange ToRange(string value)
    {
        var pieces = value.Split(':', StringSplitOptions.TrimEntries);
        return pieces.Length == 1 ? new FormulaRepairRange(pieces[0], pieces[0]) : new FormulaRepairRange(pieces[0], pieces[1]);
    }
}
