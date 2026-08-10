using ExcelTask.Core;

namespace ExcelTask.Excel.Tests;

internal static class ExcelTaskPlans
{
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

    private static FormulaRepairRange ToRange(string value)
    {
        var pieces = value.Split(':', StringSplitOptions.TrimEntries);
        return pieces.Length == 1 ? new FormulaRepairRange(pieces[0], pieces[0]) : new FormulaRepairRange(pieces[0], pieces[1]);
    }
}
