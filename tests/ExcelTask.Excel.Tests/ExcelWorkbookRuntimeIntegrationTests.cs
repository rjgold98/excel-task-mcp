using ExcelTask.Core;

namespace ExcelTask.Excel.Tests;

[CollectionDefinition("Excel COM serial", DisableParallelization = true)]
public sealed class ExcelComSerialFixture;

[Collection("Excel COM serial")]
[Trait("RunType", "OnDemand")]
public sealed class ExcelWorkbookRuntimeIntegrationTests
{
    [Fact]
    public async Task ExecuteCopiesRepairsSavesReopensAndReleasesOwnedExcel()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var target = Path.Combine(directory, "target.xlsx");
        var reference = Path.Combine(directory, "reference.xlsx");
        var output = Path.Combine(directory, "output.xlsx");

        try
        {
            ExcelTestWorkbook.CreateTarget(target);
            ExcelTestWorkbook.CreateReference(reference);
            using var runtime = new ExcelWorkbookRuntime();
            var plan = new ExcelTaskPlan("test", new NormalizedExcelTaskRequest(
                target,
                reference,
                "Reference",
                "Imported",
                [new FormulaRepairRange("A1", "A3")],
                ExcelTaskMode.Apply,
                WorkbookBinding.Isolated,
                SaveMode.Copy,
                output,
                OverwriteConfirmed: false));

            var outcome = await runtime.ExecuteAsync(plan, CancellationToken.None);

            Assert.True(outcome.Status == ExcelTaskStatus.Completed, $"{outcome.Summary} {string.Join("; ", outcome.Checks?.Select(check => check.Detail) ?? [])}");
            Assert.True(File.Exists(output));
            Assert.Contains(outcome.Changes ?? [], change => change.Kind == "formula-repair" && change.Target == "A1:A3" && change.Summary.Contains('1'));
            Assert.Contains(outcome.Checks ?? [], check => check.Name == "formula-repair-count" && check.Detail.Contains("1 repairs", StringComparison.Ordinal));
            Assert.True(ExcelTestWorkbook.HasExpectedSheetAndRepair(output));
            using var stream = new FileStream(output, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        }
        finally
        {
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public async Task AskDetectsExactOpenWorkbookAndUseOpenLeavesUserExcelRunning()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var target = Path.Combine(directory, "target.xlsx");
        var reference = Path.Combine(directory, "reference.xlsx");

        try
        {
            ExcelTestWorkbook.CreateTarget(target);
            ExcelTestWorkbook.CreateReference(reference);
            using var userWorkbook = ExcelTestWorkbook.OpenAsUser(target);
            using var runtime = new ExcelWorkbookRuntime();

            var inspection = await runtime.InspectAsync(
                new WorkbookInspectionRequest(target, reference, WorkbookBinding.AskIfOpen, SaveMode.Same, null),
                CancellationToken.None);
            var plan = new ExcelTaskPlan("attached", new NormalizedExcelTaskRequest(
                target,
                reference,
                "Reference",
                "Imported",
                [new FormulaRepairRange("A1", "A3")],
                ExcelTaskMode.Apply,
                WorkbookBinding.UseOpen,
                SaveMode.Same,
                OutputWorkbookPath: null,
                OverwriteConfirmed: true));

            var outcome = await runtime.ExecuteAsync(plan, CancellationToken.None);

            Assert.True(inspection.TargetIsOpen);
            Assert.True(outcome.Status == ExcelTaskStatus.Completed, $"{outcome.Summary} {string.Join("; ", outcome.Checks?.Select(check => check.Detail) ?? [])}");
            Assert.True(userWorkbook.IsApplicationRunning);
            Assert.True(userWorkbook.IsWorkbookOpen);
            Assert.True(ExcelTestWorkbook.HasExpectedSheetAndRepair(target));
        }
        finally
        {
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        }
    }
}
