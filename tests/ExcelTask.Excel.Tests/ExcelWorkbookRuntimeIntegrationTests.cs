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
            var plan = new ExcelTaskPlan("test", ExcelTaskPlans.Copy(
                target, reference, "Reference", "Imported", [new FormulaRepairRange("A1", "A3")],
                ExcelTaskMode.Apply, WorkbookBinding.Isolated, SaveMode.Copy, output, overwrite: false));

            var outcome = await runtime.ExecuteAsync(plan, CancellationToken.None);

            Assert.True(outcome.Status == ExcelTaskStatus.Completed, $"{outcome.Summary} {string.Join("; ", outcome.Checks?.Select(check => check.Detail) ?? [])}");
            Assert.True(File.Exists(output));
            Assert.Contains(outcome.Changes ?? [], change => change.Kind == "formula-repair" && change.Target == "Imported!A1:A3" && change.Summary.Contains('1'));
            Assert.Contains(outcome.Checks ?? [], check => check.Name == "formula-change-count" && check.Detail.Contains("1 planned formula changes", StringComparison.Ordinal));
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
            var plan = new ExcelTaskPlan("attached", ExcelTaskPlans.Copy(
                target, reference, "Reference", "Imported", [new FormulaRepairRange("A1", "A3")],
                ExcelTaskMode.Apply, WorkbookBinding.UseOpen, SaveMode.Same, output: null, overwrite: true));

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

    [Fact]
    public async Task ExecuteRepairsAnExistingWorksheet()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var target = Path.Combine(directory, "target.xlsx");
        try
        {
            ExcelTestWorkbook.CreateFormulaTarget(target, "A1:A3", new object?[,] { { "=ROW()" }, { null }, { "=ROW()" } });
            using var runtime = new ExcelWorkbookRuntime();
            var outcome = await runtime.ExecuteAsync(new ExcelTaskPlan("repair", ExcelTaskPlans.Repair(
                target, "Sheet1", [new FormulaRepairRange("A1", "A3")], ExcelTaskMode.Apply, WorkbookBinding.Isolated)), CancellationToken.None);

            Assert.Equal(ExcelTaskStatus.Completed, outcome.Status);
            Assert.Contains(outcome.Changes ?? [], change => change.Kind == "formula-repair" && change.Target == "Sheet1!A1:A3");
            Assert.True(ExcelTestWorkbook.HasFormula(target, "A2", "=ROW()"));
        }
        finally
        {
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public async Task ExecuteExtendsFormulaSeriesRight()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var target = Path.Combine(directory, "right.xlsx");
        try
        {
            ExcelTestWorkbook.CreateFormulaTarget(target, "C3:D4", new object?[,] { { "=RC[-1]", "=RC[-1]" }, { "=RC[-1]", "=RC[-1]" } }, "G3", "unchanged");
            using var runtime = new ExcelWorkbookRuntime();
            var outcome = await runtime.ExecuteAsync(new ExcelTaskPlan("right", ExcelTaskPlans.Extend(
                target, "Sheet1", FormulaExtensionDirection.Right, "C3:D4", "E3:F4", ExcelTaskMode.Apply, WorkbookBinding.Isolated)), CancellationToken.None);

            Assert.Equal(ExcelTaskStatus.Completed, outcome.Status);
            Assert.Contains(outcome.Changes ?? [], change => change.Kind == "formula-extension" && change.Target == "Sheet1!E3:F4");
            Assert.True(ExcelTestWorkbook.HasExpectedCells(target,
                new Dictionary<string, string>
                {
                    ["C3"] = "=RC[-1]",
                    ["E3"] = "=RC[-1]",
                    ["F3"] = "=RC[-1]",
                    ["E4"] = "=RC[-1]",
                    ["F4"] = "=RC[-1]"
                },
                new Dictionary<string, object> { ["G3"] = "unchanged" }));
        }
        finally
        {
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public async Task ExecuteExtendsFormulaSeriesDown()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var target = Path.Combine(directory, "down.xlsx");
        try
        {
            ExcelTestWorkbook.CreateFormulaTarget(target, "C3:D4", new object?[,] { { "=R[-1]C", "=R[-1]C" }, { "=R[-1]C", "=R[-1]C" } }, "E3", "unchanged");
            using var runtime = new ExcelWorkbookRuntime();
            var outcome = await runtime.ExecuteAsync(new ExcelTaskPlan("down", ExcelTaskPlans.Extend(
                target, "Sheet1", FormulaExtensionDirection.Down, "C3:D4", "C5:D6", ExcelTaskMode.Apply, WorkbookBinding.Isolated)), CancellationToken.None);

            Assert.Equal(ExcelTaskStatus.Completed, outcome.Status);
            Assert.Contains(outcome.Changes ?? [], change => change.Kind == "formula-extension" && change.Target == "Sheet1!C5:D6");
            Assert.True(ExcelTestWorkbook.HasExpectedCells(target,
                new Dictionary<string, string>
                {
                    ["C3"] = "=R[-1]C",
                    ["C5"] = "=R[-1]C",
                    ["D5"] = "=R[-1]C",
                    ["C6"] = "=R[-1]C",
                    ["D6"] = "=R[-1]C"
                },
                new Dictionary<string, object> { ["E3"] = "unchanged" }));
        }
        finally
        {
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public async Task PlanAnalyzesExistingRepairWithoutMutation()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var target = Path.Combine(directory, "plan.xlsx");
        try
        {
            ExcelTestWorkbook.CreateFormulaTarget(target, "A1:A3", new object?[,] { { "=ROW()" }, { null }, { "=ROW()" } });
            using var runtime = new ExcelWorkbookRuntime();
            var outcome = await runtime.ExecuteAsync(new ExcelTaskPlan("plan", ExcelTaskPlans.Repair(
                target, "Sheet1", [new FormulaRepairRange("A1", "A3")], ExcelTaskMode.Plan, WorkbookBinding.Isolated)), CancellationToken.None);

            Assert.Equal(ExcelTaskStatus.Planned, outcome.Status);
            Assert.False(ExcelTestWorkbook.HasFormula(target, "A2", "=ROW()"));
        }
        finally
        {
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public async Task MacroPlanAndApplyReplaceOnlyTheRequestedProcedureAndSaveAVerifiedCopy()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var target = Path.Combine(directory, "macro-target.xlsm");
        var output = Path.Combine(directory, "macro-output.xlsm");
        const string component = "SafeModule";
        const string procedure = "WriteMarker";
        const string originalSource = "Public Sub WriteMarker()\n    ThisWorkbook.Worksheets(1).Range(\"A1\").Value2 = \"original\"\nEnd Sub";
        const string replacementSource = "Public Sub WriteMarker()\n    ThisWorkbook.Worksheets(1).Range(\"A1\").Value2 = \"ran\"\nEnd Sub";

        try
        {
            try { ExcelTestWorkbook.CreateMacroTarget(target, component, originalSource); }
            catch (Exception exception) when (IsAccessVbomUnavailable(exception))
            {
                throw Xunit.Sdk.SkipException.ForSkip("Excel Trust Center does not permit programmatic VBA project access on this machine.");
            }

            using var runtime = new ExcelWorkbookRuntime();
            var planned = await runtime.ExecuteAsync(new ExcelTaskPlan("macro-plan", ExcelTaskPlans.Macro(
                target, output, component, procedure, ExcelTaskMode.Plan)), CancellationToken.None);

            Assert.Equal(ExcelTaskStatus.Planned, planned.Status);
            Assert.NotNull(planned.MacroProcedure);
            Assert.Equal(originalSource, planned.MacroProcedure.Source);
            Assert.False(File.Exists(output));
            Assert.Equal(originalSource, ExcelTestWorkbook.ReadMacroProcedure(target, component, procedure));

            var applied = await runtime.ExecuteAsync(new ExcelTaskPlan("macro-apply", ExcelTaskPlans.Macro(
                target, output, component, procedure, ExcelTaskMode.Apply,
                planned.MacroProcedure.Sha256, replacementSource, runAfterEdit: true)), CancellationToken.None);

            Assert.Equal(ExcelTaskStatus.Completed, applied.Status);
            Assert.NotNull(applied.MacroProcedure);
            Assert.Null(applied.MacroProcedure.Source);
            Assert.True(applied.MacroProcedure.RunCompleted);
            Assert.Equal(originalSource, ExcelTestWorkbook.ReadMacroProcedure(target, component, procedure));
            Assert.Equal(replacementSource, ExcelTestWorkbook.ReadMacroProcedure(output, component, procedure));
            Assert.True(ExcelTestWorkbook.HasValue(output, "A1", "ran"));
            using var stream = new FileStream(output, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        }
        finally
        {
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        }
    }

    private static bool IsAccessVbomUnavailable(Exception exception)
    {
        for (Exception? current = exception; current is not null; current = current.InnerException)
        {
            if (current.Message.Contains("programmatic access", StringComparison.OrdinalIgnoreCase) ||
                current.Message.Contains("trust", StringComparison.OrdinalIgnoreCase)) return true;
        }

        return false;
    }
}
