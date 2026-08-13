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
            TempDirectory.Remove(directory);
        }
    }

    [Fact]
    public async Task FormulaCopyPromotionFailureIsReportedAsUnknown()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var target = Path.Combine(directory, "target.xlsx");
        var reference = Path.Combine(directory, "reference.xlsx");
        var output = Path.Combine(directory, "output.xlsx");
        LockingRuntimeObserver? observer = null;

        try
        {
            ExcelTestWorkbook.CreateTarget(target);
            ExcelTestWorkbook.CreateReference(reference);
            // Promotion replaces an existing destination. Keeping that destination open with an
            // exclusive handle is the real filesystem failure the transaction must classify.
            ExcelTestWorkbook.CreateTarget(output);
            observer = new LockingRuntimeObserver(output);
            using var runtime = new ExcelWorkbookRuntime(observer);
            var plan = new ExcelTaskPlan("promotion-failure", ExcelTaskPlans.Copy(
                target, reference, "Reference", "Imported", [new FormulaRepairRange("A1", "A3")],
                ExcelTaskMode.Apply, WorkbookBinding.Isolated, SaveMode.Copy, output, overwrite: true));

            var outcome = await runtime.ExecuteAsync(plan, CancellationToken.None);

            Assert.Equal(ExcelTaskStatus.Unknown, outcome.Status);
            Assert.Contains(outcome.Checks ?? [], check => check.Name == "formula-save" && !check.Passed);
            Assert.True(File.Exists(output));
        }
        finally
        {
            observer?.Dispose();
            TempDirectory.Remove(directory);
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
            TempDirectory.Remove(directory);
        }
    }

    [Fact]
    public async Task VerificationReadsTheRepairedBlockAsOneOffsetBox()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var target = Path.Combine(directory, "box.xlsx");
        try
        {
            // Verification now reads the bounding box of the repairs in one array instead of one
            // COM call per cell, which is hundreds of times cheaper but introduces index
            // arithmetic that a single-cell or A1-anchored range would not exercise. This block
            // starts at C3 and spans three columns, so a wrong origin or a 1-based array bound
            // would read the wrong cell and fail verification rather than pass silently.
            ExcelTestWorkbook.CreateFormulaTarget(target, "C3:E6", new object?[,]
            {
                { "=ROW()", "=ROW()", "=ROW()" },
                { "=ROW()", null,     "=ROW()" },
                { null,     "=ROW()", null     },
                { "=ROW()", "=ROW()", "=ROW()" }
            });

            using var runtime = new ExcelWorkbookRuntime();
            var outcome = await runtime.ExecuteAsync(new ExcelTaskPlan("box-repair", ExcelTaskPlans.Repair(
                target, "Sheet1", [new FormulaRepairRange("C3", "E6")], ExcelTaskMode.Apply, WorkbookBinding.Isolated)), CancellationToken.None);

            Assert.True(outcome.Status == ExcelTaskStatus.Completed,
                $"{outcome.Summary} {string.Join("; ", outcome.Checks?.Select(check => check.Detail) ?? [])}");
            Assert.Contains(outcome.Checks ?? [], check => check.Name == "reopen-verification" && check.Passed);

            // The three interior gaps, at three different offsets from the box origin.
            Assert.True(ExcelTestWorkbook.HasFormula(target, "D4", "=ROW()"));
            Assert.True(ExcelTestWorkbook.HasFormula(target, "C5", "=ROW()"));
            Assert.True(ExcelTestWorkbook.HasFormula(target, "E5", "=ROW()"));
        }
        finally
        {
            TempDirectory.Remove(directory);
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
            TempDirectory.Remove(directory);
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
            TempDirectory.Remove(directory);
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
            TempDirectory.Remove(directory);
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
            TempDirectory.Remove(directory);
        }
    }

    [Fact]
    public async Task MacroPlanAndApplyReplaceOnlyTheRequestedProcedureAndSaveAVerifiedCopy()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var existingExcel = ExcelTestWorkbook.SnapshotSettledExcel();
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
            // Editing VBA materializes the VBE, which is the most likely way this workflow could
            // strand an Excel process, so the macro case asserts cleanup explicitly.
            TempDirectory.Remove(directory);
            ExcelTestWorkbook.AssertNoLeakedExcel(existingExcel);
        }
    }

    [Fact]
    public async Task MacroRunErrorIsTrappedAndReportedInsteadOfBlockingOnADialog()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var target = Path.Combine(directory, "macro-target.xlsm");
        var output = Path.Combine(directory, "macro-output.xlsm");
        const string component = "SafeModule";
        const string procedure = "WriteMarker";
        const string originalSource = "Public Sub WriteMarker()\n    ThisWorkbook.Worksheets(1).Range(\"A1\").Value2 = \"original\"\nEnd Sub";
        // Raises VBA error 9 at run time. Called directly this opens a modal "Run-time error '9'"
        // dialog inside the owned Excel instance and blocks until the watchdog kills the process.
        const string failingSource = "Public Sub WriteMarker()\n    Dim values(1 To 2) As Long\n    values(99) = 1\nEnd Sub";
        var existingExcel = ExcelTestWorkbook.SnapshotSettledExcel();

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

            var started = System.Diagnostics.Stopwatch.StartNew();
            var applied = await runtime.ExecuteAsync(new ExcelTaskPlan("macro-apply", ExcelTaskPlans.Macro(
                target, output, component, procedure, ExcelTaskMode.Apply,
                planned.MacroProcedure!.Sha256, failingSource, runAfterEdit: true)), CancellationToken.None);
            started.Stop();

            // The whole point: a returned result, not a 110-second watchdog kill.
            Assert.True(started.Elapsed < TimeSpan.FromSeconds(90), $"Run took {started.Elapsed}, which suggests it blocked on a dialog.");
            Assert.Equal(ExcelTaskStatus.Partial, applied.Status);
            Assert.False(applied.MacroProcedure!.RunCompleted);

            var runCheck = Assert.Single(applied.Checks!, check => check.Name == "macro-run");
            Assert.False(runCheck.Passed);
            Assert.Contains("9", runCheck.Detail, StringComparison.Ordinal);

            // The edit is still delivered and verified, and carries no trace of the generated helper.
            Assert.True(File.Exists(output));
            Assert.Equal(failingSource, ExcelTestWorkbook.ReadMacroProcedure(output, component, procedure));
            Assert.DoesNotContain("ExcelTaskRun", ExcelTestWorkbook.ReadModuleText(output, component), StringComparison.OrdinalIgnoreCase);
            Assert.Equal(originalSource, ExcelTestWorkbook.ReadMacroProcedure(target, component, procedure));
        }
        finally
        {
            TempDirectory.Remove(directory);
            ExcelTestWorkbook.AssertNoLeakedExcel(existingExcel);
        }
    }

    [Fact]
    public async Task MacroCompileErrorIsAnsweredAndReportedInsteadOfStallingTheTask()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var target = Path.Combine(directory, "macro-target.xlsm");
        var output = Path.Combine(directory, "macro-output.xlsm");
        const string component = "SafeModule";
        const string procedure = "WriteMarker";
        const string originalSource = "Public Sub WriteMarker()\n    ThisWorkbook.Worksheets(1).Range(\"A1\").Value2 = \"original\"\nEnd Sub";
        // Structurally one valid procedure, so it passes validation, but VBA cannot compile it.
        // A compile error happens before any On Error handler exists, so only the sentry can clear it.
        const string uncompilableSource = "Public Sub WriteMarker()\n    Call NoSuchProcedureExists\nEnd Sub";
        var existingExcel = ExcelTestWorkbook.SnapshotSettledExcel();

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

            var started = System.Diagnostics.Stopwatch.StartNew();
            var applied = await runtime.ExecuteAsync(new ExcelTaskPlan("macro-apply", ExcelTaskPlans.Macro(
                target, output, component, procedure, ExcelTaskMode.Apply,
                planned.MacroProcedure!.Sha256, uncompilableSource, runAfterEdit: true)), CancellationToken.None);
            started.Stop();

            Assert.True(started.Elapsed < TimeSpan.FromSeconds(90), $"Run took {started.Elapsed}, which suggests the dialog was never answered.");
            Assert.Equal(ExcelTaskStatus.Rejected, applied.Status);

            // The caller gets the compiler's own words, which is what makes a retry possible.
            var runCheck = Assert.Single(applied.Checks!, check => check.Name == "macro-run");
            Assert.False(runCheck.Passed);
            Assert.Contains("did not compile", runCheck.Detail, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Sub or Function not defined", runCheck.Detail, StringComparison.OrdinalIgnoreCase);

            // Nothing reached disk, so the retry the caller is invited to make starts from a clean slate.
            Assert.False(File.Exists(output));
            Assert.Equal(originalSource, ExcelTestWorkbook.ReadMacroProcedure(target, component, procedure));
        }
        finally
        {
            TempDirectory.Remove(directory);
            ExcelTestWorkbook.AssertNoLeakedExcel(existingExcel);
        }
    }

    [Theory]
    // An icon adds a Static control, which is what a real analyst macro almost always uses. The
    // count-based rule this replaced matched only the bare form and ignored the rest in silence.
    [InlineData("MsgBox \"field message\"")]
    [InlineData("MsgBox \"field message\", vbInformation")]
    [InlineData("MsgBox \"field message\", vbCritical, \"Title\"")]
    public async Task MacroMessageBoxRaisedByACalledProcedureIsAnsweredSoTheRunCanFinish(string call)
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var target = Path.Combine(directory, "macro-target.xlsm");
        var output = Path.Combine(directory, "macro-output.xlsm");
        const string component = "SafeModule";
        const string procedure = "WriteMarker";
        // The dialog comes from Announce, not from the replacement, so pre-flight source screening
        // cannot see it. Only the sentry can clear it once the automation thread is blocked.
        const string originalSource = "Public Sub WriteMarker()\n    ThisWorkbook.Worksheets(1).Range(\"A1\").Value2 = \"original\"\nEnd Sub";
        var helper = $"\nPublic Sub Announce()\n    {call}\nEnd Sub";
        const string replacementSource = "Public Sub WriteMarker()\n    Announce\n    ThisWorkbook.Worksheets(1).Range(\"A1\").Value2 = \"after\"\nEnd Sub";
        var existingExcel = ExcelTestWorkbook.SnapshotSettledExcel();

        try
        {
            try { ExcelTestWorkbook.CreateMacroTarget(target, component, originalSource + helper); }
            catch (Exception exception) when (IsAccessVbomUnavailable(exception))
            {
                throw Xunit.Sdk.SkipException.ForSkip("Excel Trust Center does not permit programmatic VBA project access on this machine.");
            }

            using var runtime = new ExcelWorkbookRuntime();
            var planned = await runtime.ExecuteAsync(new ExcelTaskPlan("macro-plan", ExcelTaskPlans.Macro(
                target, output, component, procedure, ExcelTaskMode.Plan)), CancellationToken.None);
            Assert.Equal(ExcelTaskStatus.Planned, planned.Status);

            var started = System.Diagnostics.Stopwatch.StartNew();
            var applied = await runtime.ExecuteAsync(new ExcelTaskPlan("macro-apply", ExcelTaskPlans.Macro(
                target, output, component, procedure, ExcelTaskMode.Apply,
                planned.MacroProcedure!.Sha256, replacementSource, runAfterEdit: true)), CancellationToken.None);
            started.Stop();

            Assert.True(started.Elapsed < TimeSpan.FromSeconds(90), $"Run took {started.Elapsed}, so the message box was never answered.");

            // Answering a message box is reported, not hidden: the run is not a clean success.
            Assert.Equal(ExcelTaskStatus.Partial, applied.Status);
            var runCheck = Assert.Single(applied.Checks!, check => check.Name == "macro-run");
            Assert.False(runCheck.Passed);
            Assert.Contains("message box", runCheck.Detail, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("field message", runCheck.Detail, StringComparison.Ordinal);

            // The macro ran past the dialog, which is the proof the button was really pressed.
            Assert.True(ExcelTestWorkbook.HasValue(output, "A1", "after"));
            Assert.Equal(replacementSource, ExcelTestWorkbook.ReadMacroProcedure(output, component, procedure));
            Assert.DoesNotContain("ExcelTaskRun", ExcelTestWorkbook.ReadModuleText(output, component), StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            TempDirectory.Remove(directory);
            ExcelTestWorkbook.AssertNoLeakedExcel(existingExcel);
        }
    }

    // A dialog offering a real choice - Yes/No, OK/Cancel, Retry/Cancel - is deliberately left
    // alone, so the blocked automation call never returns. Only the supervised runtime bounds that,
    // by its deadline; this direct-runtime layer has none and would hang forever.
    //
    // That exclusion is therefore UNTESTED. Driving it needs a real modal dialog and a real blocked
    // COM call, and any test asserting it here would hang rather than fail. The rule it protects -
    // identifier 2 is OK on a one-button dialog but Cancel on OK/Cancel and Retry/Cancel, so a
    // count-based match presses exactly the wrong button - is enforced only by ModalDialogSentry's
    // own structure. Measured dialog layouts are recorded in the 0.6.0 release notes.

    [Fact]
    public async Task WriteStoresConstantsWithTheirTypesAndProvesThemAfterReopening()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var target = Path.Combine(directory, "write.xlsx");
        var existingExcel = ExcelTestWorkbook.SnapshotSettledExcel();

        try
        {
            ExcelTestWorkbook.CreateTarget(target);
            using var runtime = new ExcelWorkbookRuntime();
            var outcome = await runtime.ExecuteAsync(new ExcelTaskPlan("write", ExcelTaskPlans.Write(
                target, "Sheet1", [("A1", "Revenue"), ("B1", "1000.5"), ("C1", "TRUE"), ("A2", "")])),
                CancellationToken.None);

            Assert.True(outcome.Status == ExcelTaskStatus.Completed,
                $"{outcome.Summary} {string.Join("; ", outcome.Checks?.Select(check => check.Detail) ?? [])}");
            Assert.Contains(outcome.Checks ?? [], check => check.Name == "reopen-verification" && check.Passed);

            // A number sent as text must land as a number, or every SUM above it silently breaks -
            // which is the failure this operation exists to not have.
            Assert.True(ExcelTestWorkbook.HasValue(target, "B1", 1000.5d));
            Assert.True(ExcelTestWorkbook.HasValue(target, "A1", "Revenue"));
            Assert.True(ExcelTestWorkbook.HasValue(target, "C1", true));
        }
        finally
        {
            TempDirectory.Remove(directory);
            ExcelTestWorkbook.AssertNoLeakedExcel(existingExcel);
        }
    }

    [Fact]
    public async Task FormulaWriteStoresCallerSuppliedA1FormulaAndProvesItAfterReopening()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var target = Path.Combine(directory, "formula-write.xlsx");
        var existingExcel = ExcelTestWorkbook.SnapshotSettledExcel();

        try
        {
            ExcelTestWorkbook.CreateTarget(target);
            using var runtime = new ExcelWorkbookRuntime();
            var outcome = await runtime.ExecuteAsync(new ExcelTaskPlan("formula-write", ExcelTaskPlans.WriteFormulas(
                target, "Sheet1", [("B1", "=SUM(1,2)")])), CancellationToken.None);

            Assert.True(outcome.Status == ExcelTaskStatus.Completed,
                $"{outcome.Summary} {string.Join("; ", outcome.Checks?.Select(check => check.Detail) ?? [])}");
            Assert.Contains(outcome.Checks ?? [], check => check.Name == "formula-write" && check.Passed);
            Assert.Contains(outcome.Checks ?? [], check => check.Name == "reopen-verification" && check.Passed);
            Assert.True(ExcelTestWorkbook.HasFormula(target, "B1", "=SUM(1,2)"));
            Assert.True(ExcelTestWorkbook.HasValue(target, "B1", 3d));
        }
        finally
        {
            TempDirectory.Remove(directory);
            ExcelTestWorkbook.AssertNoLeakedExcel(existingExcel);
        }
    }

    [Fact]
    public async Task WritePlanChangesNothingOnDisk()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var target = Path.Combine(directory, "write-plan.xlsx");
        var existingExcel = ExcelTestWorkbook.SnapshotSettledExcel();

        try
        {
            ExcelTestWorkbook.CreateTarget(target);
            var stampBefore = (new FileInfo(target).Length, new FileInfo(target).LastWriteTimeUtc);

            using var runtime = new ExcelWorkbookRuntime();
            var outcome = await runtime.ExecuteAsync(new ExcelTaskPlan("write-plan", ExcelTaskPlans.Write(
                target, "Sheet1", [("A1", "Revenue")], ExcelTaskMode.Plan)), CancellationToken.None);

            Assert.Equal(ExcelTaskStatus.Planned, outcome.Status);
            Assert.Equal(stampBefore, (new FileInfo(target).Length, new FileInfo(target).LastWriteTimeUtc));
        }
        finally
        {
            TempDirectory.Remove(directory);
            ExcelTestWorkbook.AssertNoLeakedExcel(existingExcel);
        }
    }

    [Fact]
    public async Task AnApplyThatFailsBeforeVerifyingStillReleasesThePreLaunchedExcel()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var target = Path.Combine(directory, "prelaunch.xlsx");
        var existingExcel = ExcelTestWorkbook.SnapshotSettledExcel();

        try
        {
            // The verification Excel is now started before the work rather than after it, which
            // buys back a launch but creates a way to leak a process: every path that returns
            // before verifying leaves one running that nothing else would close. This asks for a
            // worksheet that does not exist, so preflight rejects and the run returns long before
            // the verification would have been used. It is the reason the pre-launch is allowed to
            // exist, and it is worth more than the milliseconds it protects.
            ExcelTestWorkbook.CreateFormulaTarget(target, "A1:A3", new object?[,] { { "=ROW()" }, { null }, { "=ROW()" } });
            using var runtime = new ExcelWorkbookRuntime();
            var outcome = await runtime.ExecuteAsync(new ExcelTaskPlan("prelaunch-reject", ExcelTaskPlans.Repair(
                target, "NoSuchSheet", [new FormulaRepairRange("A1", "A3")], ExcelTaskMode.Apply, WorkbookBinding.Isolated)),
                CancellationToken.None);

            Assert.Equal(ExcelTaskStatus.Rejected, outcome.Status);
        }
        finally
        {
            TempDirectory.Remove(directory);
            ExcelTestWorkbook.AssertNoLeakedExcel(existingExcel);
        }
    }

    [Fact]
    public async Task ReadReturnsCellContentsAndOmitsBlanksWithoutChangingTheWorkbook()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var target = Path.Combine(directory, "read-target.xlsx");
        var existingExcel = ExcelTestWorkbook.SnapshotSettledExcel();

        try
        {
            // Two of the nine cells are blank, so the receipt can be checked for omitting them
            // rather than padding the answer with empties the caller would have to filter.
            ExcelTestWorkbook.CreateFormulaTarget(target, "A1:C3", new object?[,]
            {
                { "=ROW()", "=ROW()*10", null       },
                { "=ROW()", null,        "=ROW()*3" },
                { "=ROW()", "=ROW()*10", "=ROW()*3" }
            });
            var stampBefore = (new FileInfo(target).Length, new FileInfo(target).LastWriteTimeUtc);

            using var runtime = new ExcelWorkbookRuntime();
            var outcome = await runtime.ExecuteAsync(
                new ExcelTaskPlan("read", ExcelTaskPlans.Read(target, "Sheet1", "A1:C3")), CancellationToken.None);

            Assert.True(outcome.Status == ExcelTaskStatus.Completed,
                $"{outcome.Summary} {string.Join("; ", outcome.Checks?.Select(check => check.Detail) ?? [])}");
            Assert.NotNull(outcome.Range);
            Assert.Equal(9, outcome.Range.CellsInRange);
            Assert.Equal(7, outcome.Range.NonEmptyCells);
            Assert.Equal(7, outcome.Range.Cells.Count);
            Assert.False(outcome.Range.Truncated);
            Assert.DoesNotContain(outcome.Range.Cells, cell => cell.Address is "C1" or "B2");

            // Values by default: =ROW() in row 3 comes back as 3, not as its formula.
            Assert.Equal("3", Assert.Single(outcome.Range.Cells, cell => cell.Address == "A3").Text);
            Assert.Equal("30", Assert.Single(outcome.Range.Cells, cell => cell.Address == "B3").Text);

            // Read-only means provably unchanged on disk, the same proof the audit gives.
            Assert.Equal(stampBefore, (new FileInfo(target).Length, new FileInfo(target).LastWriteTimeUtc));
            Assert.Contains(outcome.Checks ?? [], check => check.Name == "workbook-unchanged" && check.Passed);
        }
        finally
        {
            TempDirectory.Remove(directory);
            ExcelTestWorkbook.AssertNoLeakedExcel(existingExcel);
        }
    }

    [Fact]
    public async Task AFullDenseReadReturnsEveryOneOfThePermittedCells()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var target = Path.Combine(directory, "read-dense.xlsx");
        var existingExcel = ExcelTestWorkbook.SnapshotSettledExcel();

        try
        {
            // The worst case the caller can ask for: every one of the 400 permitted cells full, and
            // each holding text past the per-cell cap. The runtime is the innermost layer and does
            // not truncate - the engine, the worker, and the tool each bound what they pass on -
            // so what this proves is that a dense read at the limit returns every cell.
            // WorkbookWorkerProtocolTests covers the frame budget that carries it.
            var formulas = new object?[20, 20];
            for (var row = 0; row < 20; row++)
            {
                for (var column = 0; column < 20; column++) formulas[row, column] = "=REPT(\"x\",120)";
            }

            ExcelTestWorkbook.CreateFormulaTarget(target, "A1:T20", formulas);
            using var runtime = new ExcelWorkbookRuntime();
            var outcome = await runtime.ExecuteAsync(
                new ExcelTaskPlan("read-dense", ExcelTaskPlans.Read(target, "Sheet1", "A1:T20")), CancellationToken.None);

            Assert.True(outcome.Status == ExcelTaskStatus.Completed,
                $"{outcome.Summary} {string.Join("; ", outcome.Checks?.Select(check => check.Detail) ?? [])}");
            Assert.NotNull(outcome.Range);
            Assert.Equal(400, outcome.Range.CellsInRange);
            Assert.Equal(400, outcome.Range.Cells.Count);
            Assert.All(outcome.Range.Cells, cell => Assert.Equal(120, cell.Text.Length));
        }
        finally
        {
            TempDirectory.Remove(directory);
            ExcelTestWorkbook.AssertNoLeakedExcel(existingExcel);
        }
    }

    [Fact]
    public async Task ReadReturnsR1C1FormulasWhenAskedAndRejectsAnUnknownWorksheet()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var target = Path.Combine(directory, "read-formulas.xlsx");
        var existingExcel = ExcelTestWorkbook.SnapshotSettledExcel();

        try
        {
            ExcelTestWorkbook.CreateFormulaTarget(target, "A1:A2", new object?[,] { { "=ROW()" }, { "=ROW()" } });
            using var runtime = new ExcelWorkbookRuntime();

            var outcome = await runtime.ExecuteAsync(
                new ExcelTaskPlan("read-formulas", ExcelTaskPlans.Read(target, "Sheet1", "A1:A2", formulas: true)),
                CancellationToken.None);

            Assert.Equal(ExcelTaskStatus.Completed, outcome.Status);
            Assert.NotNull(outcome.Range);
            Assert.True(outcome.Range.Formulas);
            Assert.All(outcome.Range.Cells, cell => Assert.Equal("=ROW()", cell.Text));

            // A misspelled sheet name must say so, not return an empty range that reads as an
            // answer. The caller is told where the real names come from.
            var missing = await runtime.ExecuteAsync(
                new ExcelTaskPlan("read-missing", ExcelTaskPlans.Read(target, "NoSuchSheet", "A1:A2")),
                CancellationToken.None);

            Assert.Equal(ExcelTaskStatus.Rejected, missing.Status);
            Assert.Null(missing.Range);
            Assert.Contains(missing.Checks ?? [], check => check.Name == "worksheet" && !check.Passed);
        }
        finally
        {
            TempDirectory.Remove(directory);
            ExcelTestWorkbook.AssertNoLeakedExcel(existingExcel);
        }
    }

    [Fact]
    public async Task AuditReportsFlowsWithoutChangingTheWorkbookOrLeakingPaths()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var target = Path.Combine(directory, "audit-target.xlsx");
        var reference = Path.Combine(directory, "audit-reference.xlsx");
        var existingExcel = ExcelTestWorkbook.SnapshotSettledExcel();

        try
        {
            ExcelTestWorkbook.CreateReference(reference);
            var hasQuery = ExcelTestWorkbook.CreateAuditTarget(target, reference);
            var stampBefore = (new FileInfo(target).Length, new FileInfo(target).LastWriteTimeUtc);

            using var runtime = new ExcelWorkbookRuntime();
            var outcome = await runtime.ExecuteAsync(
                new ExcelTaskPlan("audit", ExcelTaskPlans.Audit(target)), CancellationToken.None);

            Assert.True(outcome.Status == ExcelTaskStatus.Completed,
                $"{outcome.Summary} {string.Join("; ", outcome.Checks?.Select(check => check.Detail) ?? [])}");
            Assert.NotNull(outcome.Audit);
            Assert.True(outcome.Audit.WorkbookUnchanged);
            Assert.True(outcome.Audit.TotalFound >= 1);

            // Worksheets are named so a caller can aim the formula operations, every one of which
            // requires a worksheet name it otherwise has no way to discover.
            var sheet = Assert.Single(outcome.Audit.Items, item => item.Kind == "worksheet");
            Assert.Contains("worksheet", sheet.Detail, StringComparison.Ordinal);

            // The external link is reported as a file name only; the directory would carry
            // machine-specific path segments a receipt must never contain.
            var link = Assert.Single(outcome.Audit.Items, item => item.Kind == "external-link");
            Assert.Equal("audit-reference.xlsx", link.Name);
            Assert.All(outcome.Audit.Items, item => Assert.DoesNotContain(directory, item.Name + item.Detail + item.DependsOn, StringComparison.OrdinalIgnoreCase));

            if (hasQuery)
            {
                var query = Assert.Single(outcome.Audit.Items, item => item.Kind == "query");
                Assert.Equal("AuditQuery", query.Name);
            }

            // Read-only means provably unchanged on disk, not just unpersisted.
            var stampAfter = (new FileInfo(target).Length, new FileInfo(target).LastWriteTimeUtc);
            Assert.Equal(stampBefore, stampAfter);
            Assert.Contains(outcome.Checks ?? [], check => check.Name == "workbook-unchanged" && check.Passed);
        }
        finally
        {
            TempDirectory.Remove(directory);
            ExcelTestWorkbook.AssertNoLeakedExcel(existingExcel);
        }
    }

    [Fact]
    public async Task AuditListsMacroProceduresSoAnEditCanBeDiscovered()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var target = Path.Combine(directory, "macro-target.xlsm");
        const string component = "SafeModule";
        const string source = "Public Sub WriteMarker()\n    ThisWorkbook.Worksheets(1).Range(\"A1\").Value2 = \"original\"\nEnd Sub";
        var existingExcel = ExcelTestWorkbook.SnapshotSettledExcel();

        try
        {
            try { ExcelTestWorkbook.CreateMacroTarget(target, component, source); }
            catch (Exception exception) when (IsAccessVbomUnavailable(exception))
            {
                throw Xunit.Sdk.SkipException.ForSkip("Excel Trust Center does not permit programmatic VBA project access on this machine.");
            }

            using var runtime = new ExcelWorkbookRuntime();
            var outcome = await runtime.ExecuteAsync(
                new ExcelTaskPlan("audit-macros", ExcelTaskPlans.Audit(target)), CancellationToken.None);

            Assert.Equal(ExcelTaskStatus.Completed, outcome.Status);
            Assert.NotNull(outcome.Audit);
            Assert.True(outcome.Audit.WorkbookUnchanged);

            // The discovery a caller needs before EditMacroProcedure: the component and the
            // qualified procedure name, with its size, and never its source.
            Assert.Contains(outcome.Audit.Items, item => item.Kind == "macro-component" && item.Name == component);
            var procedure = Assert.Single(outcome.Audit.Items, item => item.Kind == "macro-procedure");
            Assert.Equal($"{component}.WriteMarker", procedure.Name);
            Assert.Equal("3 lines", procedure.Detail);
            Assert.Equal(component, procedure.DependsOn);
            Assert.All(outcome.Audit.Items, item => Assert.DoesNotContain("Value2", item.Name + item.Detail, StringComparison.OrdinalIgnoreCase));
        }
        finally
        {
            TempDirectory.Remove(directory);
            ExcelTestWorkbook.AssertNoLeakedExcel(existingExcel);
        }
    }

    [Fact]
    public async Task FindReplaceRewritesOnlyConstantsAndProvesThemAfterReopening()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var target = Path.Combine(directory, "find.xlsx");
        var existingExcel = ExcelTestWorkbook.SnapshotSettledExcel();

        try
        {
            // A2 holds a formula whose *result* contains the search text. It must be reported as a
            // match and left exactly as it was, which is the rule this operation turns on.
            ExcelTestWorkbook.CreateFormulaTarget(target, "A1:A3", new object?[,]
            {
                { "FY25 Budget" },
                { "=\"FY25 \" & \"Actual\"" },
                { "Prior FY25" }
            });

            using var runtime = new ExcelWorkbookRuntime();
            var outcome = await runtime.ExecuteAsync(new ExcelTaskPlan("find", ExcelTaskPlans.FindReplace(
                target, "Sheet1", "FY25", "FY26", "A1:A3")), CancellationToken.None);

            Assert.True(outcome.Status == ExcelTaskStatus.Completed,
                $"{outcome.Summary} {string.Join("; ", outcome.Checks?.Select(check => check.Detail) ?? [])}");
            Assert.Contains(outcome.Checks ?? [], check => check.Name == "reopen-verification" && check.Passed);
            Assert.Contains(outcome.Checks ?? [], check => check.Name == "formula-cells-untouched" && check.Passed);

            Assert.True(ExcelTestWorkbook.HasValue(target, "A1", "FY26 Budget"));
            Assert.True(ExcelTestWorkbook.HasValue(target, "A3", "Prior FY26"));
            Assert.True(ExcelTestWorkbook.HasFormula(target, "A2", "=\"FY25 \" & \"Actual\""));
        }
        finally
        {
            TempDirectory.Remove(directory);
            ExcelTestWorkbook.AssertNoLeakedExcel(existingExcel);
        }
    }

    [Fact]
    public async Task FindReplacePlanListsTheMatchesAndChangesNothingOnDisk()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var target = Path.Combine(directory, "find-plan.xlsx");
        var existingExcel = ExcelTestWorkbook.SnapshotSettledExcel();

        try
        {
            ExcelTestWorkbook.CreateFormulaTarget(target, "A1:A3", new object?[,]
            {
                { "FY25 Budget" }, { "Other" }, { "Prior FY25" }
            });
            var stampBefore = (new FileInfo(target).Length, new FileInfo(target).LastWriteTimeUtc);

            using var runtime = new ExcelWorkbookRuntime();
            var outcome = await runtime.ExecuteAsync(new ExcelTaskPlan("find-plan", ExcelTaskPlans.FindReplace(
                target, "Sheet1", "FY25", range: "A1:A3", mode: ExcelTaskMode.Plan)), CancellationToken.None);

            Assert.Equal(ExcelTaskStatus.Planned, outcome.Status);
            Assert.NotNull(outcome.Range);
            Assert.Equal(2, outcome.Range.NonEmptyCells);
            Assert.Equal(["A1", "A3"], outcome.Range.Cells.Select(cell => cell.Address));
            Assert.Equal(stampBefore, (new FileInfo(target).Length, new FileInfo(target).LastWriteTimeUtc));
        }
        finally
        {
            TempDirectory.Remove(directory);
            ExcelTestWorkbook.AssertNoLeakedExcel(existingExcel);
        }
    }

    [Fact]
    public async Task FindReplaceHonoursWholeCellAndCaseWithoutTreatingAsteriskAsAWildcard()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var target = Path.Combine(directory, "find-exact.xlsx");
        var existingExcel = ExcelTestWorkbook.SnapshotSettledExcel();

        try
        {
            // Excel's own Find would treat "NET*" as a wildcard and match all four. Matching in code
            // means the asterisk is a character, which is what the caller wrote.
            ExcelTestWorkbook.CreateFormulaTarget(target, "A1:A4", new object?[,]
            {
                { "NET*" }, { "NETTING" }, { "net*" }, { "NET* extra" }
            });

            using var runtime = new ExcelWorkbookRuntime();
            var outcome = await runtime.ExecuteAsync(new ExcelTaskPlan("find-exact", ExcelTaskPlans.FindReplace(
                target, "Sheet1", "NET*", "REV", "A1:A4", wholeCell: true, matchCase: true)), CancellationToken.None);

            Assert.True(outcome.Status == ExcelTaskStatus.Completed,
                $"{outcome.Summary} {string.Join("; ", outcome.Checks?.Select(check => check.Detail) ?? [])}");
            Assert.True(ExcelTestWorkbook.HasValue(target, "A1", "REV"));
            Assert.True(ExcelTestWorkbook.HasValue(target, "A2", "NETTING"));
            Assert.True(ExcelTestWorkbook.HasValue(target, "A3", "net*"));
            Assert.True(ExcelTestWorkbook.HasValue(target, "A4", "NET* extra"));
        }
        finally
        {
            TempDirectory.Remove(directory);
            ExcelTestWorkbook.AssertNoLeakedExcel(existingExcel);
        }
    }

    [Fact]
    public async Task FindReplaceRefusesTheWholeRequestWhenOneResultWouldBecomeAFormula()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var target = Path.Combine(directory, "find-formula.xlsx");
        var existingExcel = ExcelTestWorkbook.SnapshotSettledExcel();

        try
        {
            // Removing "x" from "x=1" leaves "=1", which Excel stores as a formula. The replacement
            // itself is a legal constant, so only composing the result can catch it - and it must
            // stop A1 changing too, or a refusal would leave the sheet half rewritten.
            ExcelTestWorkbook.CreateFormulaTarget(target, "A1:A2", new object?[,] { { "xy" }, { "x=1" } });

            using var runtime = new ExcelWorkbookRuntime();
            var outcome = await runtime.ExecuteAsync(new ExcelTaskPlan("find-formula", ExcelTaskPlans.FindReplace(
                target, "Sheet1", "x", string.Empty, "A1:A2")), CancellationToken.None);

            Assert.Equal(ExcelTaskStatus.Rejected, outcome.Status);
            Assert.Contains(outcome.Checks ?? [], check => check.Name == "formula-text-refused" && !check.Passed);
            Assert.True(ExcelTestWorkbook.HasValue(target, "A1", "xy"));
            Assert.True(ExcelTestWorkbook.HasValue(target, "A2", "x=1"));
        }
        finally
        {
            TempDirectory.Remove(directory);
            ExcelTestWorkbook.AssertNoLeakedExcel(existingExcel);
        }
    }

    [Fact]
    public async Task CreateMakesAnEmptyWorkbookAndRefusesToOverwriteOne()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var target = Path.Combine(directory, "created.xlsx");
        var existingExcel = ExcelTestWorkbook.SnapshotSettledExcel();

        try
        {
            using var runtime = new ExcelWorkbookRuntime();
            var created = await runtime.ExecuteAsync(
                new ExcelTaskPlan("create", ExcelTaskPlans.Create(target, CreateKind.Workbook)), CancellationToken.None);

            Assert.True(created.Status == ExcelTaskStatus.Completed,
                $"{created.Summary} {string.Join("; ", created.Checks?.Select(check => check.Detail) ?? [])}");
            Assert.True(File.Exists(target));

            // The second attempt is the point: "create" must never quietly mean "replace".
            var again = await runtime.ExecuteAsync(
                new ExcelTaskPlan("create-again", ExcelTaskPlans.Create(target, CreateKind.Workbook)), CancellationToken.None);

            Assert.Equal(ExcelTaskStatus.Rejected, again.Status);
            Assert.Contains(again.Checks ?? [], check => check.Name == "workbook-inputs" && !check.Passed);
        }
        finally
        {
            TempDirectory.Remove(directory);
            ExcelTestWorkbook.AssertNoLeakedExcel(existingExcel);
        }
    }

    [Fact]
    public async Task CreateRefusesAWorksheetNameExcelReservesWithoutClaimingAnythingWasWritten()
    {
        // Excel reserves a handful of worksheet names, and they pass the engine's validation, which
        // only knows about length and forbidden characters. The rename therefore fails inside the
        // runtime with the target path still free and nothing on disk.
        //
        // That used to report Unknown and not-retryable, sending the caller to reconcile a file
        // that was never created. The fix is one line - marking the mutation attempted after the
        // rename rather than before - and nothing observed it, which is why this exists.
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var target = Path.Combine(directory, "reserved.xlsx");
        var existingExcel = ExcelTestWorkbook.SnapshotSettledExcel();

        try
        {
            using var runtime = new ExcelWorkbookRuntime();
            var outcome = await runtime.ExecuteAsync(
                new ExcelTaskPlan("create-reserved", ExcelTaskPlans.Create(target, CreateKind.Workbook, "History")),
                CancellationToken.None);

            // Excel builds vary in which names they refuse. If this one was accepted the operation
            // simply succeeds, and the assertion below would be about the wrong thing.
            if (outcome.Status == ExcelTaskStatus.Completed)
            {
                Assert.True(File.Exists(target));
                return;
            }

            Assert.Equal(ExcelTaskStatus.Rejected, outcome.Status);
            Assert.True(outcome.CanRetry);
            // A retryable rejection has to say what to change, or the caller resubmits unchanged.
            Assert.False(string.IsNullOrWhiteSpace(outcome.RetryReason));
            // The whole basis for calling it Rejected rather than Unknown: nothing was written.
            Assert.False(File.Exists(target));
        }
        finally
        {
            TempDirectory.Remove(directory);
            ExcelTestWorkbook.AssertNoLeakedExcel(existingExcel);
        }
    }

    [Fact]
    public async Task CreateAddsAWorksheetAfterTheLastAndProvesItAfterReopening()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var target = Path.Combine(directory, "add-sheet.xlsx");
        var existingExcel = ExcelTestWorkbook.SnapshotSettledExcel();

        try
        {
            ExcelTestWorkbook.CreateTarget(target);
            using var runtime = new ExcelWorkbookRuntime();
            var outcome = await runtime.ExecuteAsync(new ExcelTaskPlan(
                "add-sheet", ExcelTaskPlans.Create(target, CreateKind.Worksheet, "Q3 Actuals")), CancellationToken.None);

            Assert.True(outcome.Status == ExcelTaskStatus.Completed,
                $"{outcome.Summary} {string.Join("; ", outcome.Checks?.Select(check => check.Detail) ?? [])}");
            Assert.Contains(outcome.Checks ?? [], check => check.Name == "reopen-verification" && check.Passed);

            // Added after the last sheet, so it never displaces the one the workbook opens on.
            var audit = await runtime.ExecuteAsync(
                new ExcelTaskPlan("audit", ExcelTaskPlans.Audit(target)), CancellationToken.None);
            var worksheets = (audit.Audit?.Items ?? []).Where(item => item.Kind == "worksheet").ToArray();
            Assert.Equal("Q3 Actuals", worksheets[^1].Name);

            var again = await runtime.ExecuteAsync(new ExcelTaskPlan(
                "add-again", ExcelTaskPlans.Create(target, CreateKind.Worksheet, "Q3 Actuals")), CancellationToken.None);
            Assert.Equal(ExcelTaskStatus.Rejected, again.Status);
            Assert.Contains(again.Checks ?? [], check => check.Name == "worksheet-name" && !check.Passed);
        }
        finally
        {
            TempDirectory.Remove(directory);
            ExcelTestWorkbook.AssertNoLeakedExcel(existingExcel);
        }
    }

    [Fact]
    public async Task SetNumberFormatChangesHowTheNumberReadsAndNotTheNumber()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var target = Path.Combine(directory, "format.xlsx");
        var existingExcel = ExcelTestWorkbook.SnapshotSettledExcel();
        // Parenthesised negatives with the padding that aligns them under positives: the format a
        // financial exhibit actually uses, and the one whose trailing spaces must survive intact.
        const string accounting = "#,##0.00_);(#,##0.00)";

        try
        {
            ExcelTestWorkbook.CreateFormulaTarget(target, "A1:A2", new object?[,] { { 1000.5 }, { -250.25 } });

            using var runtime = new ExcelWorkbookRuntime();
            var outcome = await runtime.ExecuteAsync(new ExcelTaskPlan("format", ExcelTaskPlans.NumberFormat(
                target, "Sheet1", "A1:A2", accounting)), CancellationToken.None);

            Assert.True(outcome.Status == ExcelTaskStatus.Completed,
                $"{outcome.Summary} {string.Join("; ", outcome.Checks?.Select(check => check.Detail) ?? [])}");
            Assert.Contains(outcome.Checks ?? [], check => check.Name == "reopen-verification" && check.Passed);
            Assert.Equal(accounting, ExcelTestWorkbook.ReadNumberFormat(target, "A1:A2"));

            // The point of the whole operation: presentation moved, the numbers did not. A format
            // that quietly rounded or retyped the value would be worse than no format at all.
            Assert.True(ExcelTestWorkbook.HasValue(target, "A1", 1000.5d));
            Assert.True(ExcelTestWorkbook.HasValue(target, "A2", -250.25d));
        }
        finally
        {
            TempDirectory.Remove(directory);
            ExcelTestWorkbook.AssertNoLeakedExcel(existingExcel);
        }
    }

    [Fact]
    public async Task SetNumberFormatPlanReportsWhatIsThereAndChangesNothingOnDisk()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var target = Path.Combine(directory, "format-plan.xlsx");
        var existingExcel = ExcelTestWorkbook.SnapshotSettledExcel();

        try
        {
            ExcelTestWorkbook.CreateFormulaTarget(target, "A1:A2", new object?[,] { { 1000.5 }, { 2000.25 } });
            var stampBefore = (new FileInfo(target).Length, new FileInfo(target).LastWriteTimeUtc);

            using var runtime = new ExcelWorkbookRuntime();
            var outcome = await runtime.ExecuteAsync(new ExcelTaskPlan("format-plan", ExcelTaskPlans.NumberFormat(
                target, "Sheet1", "A1:A2", "0.0%", ExcelTaskMode.Plan)), CancellationToken.None);

            Assert.Equal(ExcelTaskStatus.Planned, outcome.Status);
            // A format is destructive in a way a value write is not - it replaces whatever was there
            // and the old code is not recoverable from the sheet - so Plan names what it would replace.
            Assert.Contains(outcome.Checks ?? [], check => check.Name == "current-format" && check.Detail.Contains("General", StringComparison.Ordinal));
            Assert.Equal(stampBefore, (new FileInfo(target).Length, new FileInfo(target).LastWriteTimeUtc));
        }
        finally
        {
            TempDirectory.Remove(directory);
            ExcelTestWorkbook.AssertNoLeakedExcel(existingExcel);
        }
    }

    [Fact]
    public async Task SetNumberFormatClearsFormattingWithGeneralAndRejectsAMissingWorksheet()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var target = Path.Combine(directory, "format-general.xlsx");
        var existingExcel = ExcelTestWorkbook.SnapshotSettledExcel();

        try
        {
            ExcelTestWorkbook.CreateFormulaTarget(target, "A1:A2", new object?[,] { { 1000.5 }, { 2000.25 } });
            using var runtime = new ExcelWorkbookRuntime();

            var formatted = await runtime.ExecuteAsync(new ExcelTaskPlan("format-set", ExcelTaskPlans.NumberFormat(
                target, "Sheet1", "A1:A2", "0.0%")), CancellationToken.None);
            Assert.Equal(ExcelTaskStatus.Completed, formatted.Status);
            Assert.Equal("0.0%", ExcelTestWorkbook.ReadNumberFormat(target, "A1:A2"));

            var cleared = await runtime.ExecuteAsync(new ExcelTaskPlan("format-clear", ExcelTaskPlans.NumberFormat(
                target, "Sheet1", "A1:A2", "General")), CancellationToken.None);
            Assert.Equal(ExcelTaskStatus.Completed, cleared.Status);
            Assert.Equal("General", ExcelTestWorkbook.ReadNumberFormat(target, "A1:A2"));

            var missing = await runtime.ExecuteAsync(new ExcelTaskPlan("format-missing", ExcelTaskPlans.NumberFormat(
                target, "NoSuchSheet", "A1:A2", "0.0%")), CancellationToken.None);
            Assert.Equal(ExcelTaskStatus.Rejected, missing.Status);
        }
        finally
        {
            TempDirectory.Remove(directory);
            ExcelTestWorkbook.AssertNoLeakedExcel(existingExcel);
        }
    }

    [Fact]
    public async Task TheScanReadsAWorkbookExcelActuallyWroteWithoutStartingExcel()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var target = Path.Combine(directory, "scan-real.xlsx");
        var existingExcel = ExcelTestWorkbook.SnapshotSettledExcel();

        try
        {
            // Written by real Excel - the fixture launch is the ONLY Excel in this test - so the
            // scan is proven against the XML Excel actually produces (shared formulas as Excel
            // writes them), not only against the hand-authored packages in the fast tier.
            ExcelTestWorkbook.CreateFormulaTarget(target, "A1:A3", new object?[,] { { "=ROW()" }, { "=ROW()" }, { "=ROW()" } }, "B1", 42d);
            var before = OwnedExcelProcess.SnapshotExcelProcesses();

            using var runtime = new ExcelWorkbookRuntime();
            var outcome = await runtime.ExecuteAsync(
                new ExcelTaskPlan("scan-real", ExcelTaskPlans.Scan(target)), CancellationToken.None);

            Assert.True(outcome.Status == ExcelTaskStatus.Planned,
                $"{outcome.Summary} {string.Join("; ", outcome.Checks?.Select(check => check.Detail) ?? [])}");
            var sheet = Assert.Single(outcome.Audit!.Items, item => item.Kind == "scan-sheet");
            Assert.Contains("3 formula cell(s)", sheet.Detail, StringComparison.Ordinal);
            Assert.Contains("1 constant cell(s)", sheet.Detail, StringComparison.Ordinal);

            // The operation's entire promise: the process table is untouched.
            Assert.Equal(before, OwnedExcelProcess.SnapshotExcelProcesses());
        }
        finally
        {
            TempDirectory.Remove(directory);
            ExcelTestWorkbook.AssertNoLeakedExcel(existingExcel);
        }
    }

    [Fact]
    public async Task AReadSaysWhichCellsAreFormulasSoNobodyHasToReadTheRangeTwice()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var target = Path.Combine(directory, "isformula.xlsx");
        var existingExcel = ExcelTestWorkbook.SnapshotSettledExcel();

        try
        {
            // A2 is a formula whose result is a number; A3 is that same number as a constant. They
            // are indistinguishable in a values read, which is exactly why a caller about to
            // overwrite one had to read the range twice and diff. A UX simulation did precisely that.
            ExcelTestWorkbook.CreateFormulaTarget(target, "A1:A3", new object?[,] { { 20 }, { "=R[-1]C*2" }, { 40 } });

            using var runtime = new ExcelWorkbookRuntime();
            var outcome = await runtime.ExecuteAsync(
                new ExcelTaskPlan("isformula", ExcelTaskPlans.Read(target, "Sheet1", "A1:A3")), CancellationToken.None);

            Assert.True(outcome.Status == ExcelTaskStatus.Completed,
                $"{outcome.Summary} {string.Join("; ", outcome.Checks?.Select(check => check.Detail) ?? [])}");
            var cells = outcome.Range!.Cells;
            Assert.Equal("40", Assert.Single(cells, cell => cell.Address == "A2").Text);
            Assert.Equal("40", Assert.Single(cells, cell => cell.Address == "A3").Text);
            Assert.True(Assert.Single(cells, cell => cell.Address == "A2").IsFormula);
            Assert.False(Assert.Single(cells, cell => cell.Address == "A3").IsFormula);
            Assert.False(Assert.Single(cells, cell => cell.Address == "A1").IsFormula);
        }
        finally
        {
            TempDirectory.Remove(directory);
            ExcelTestWorkbook.AssertNoLeakedExcel(existingExcel);
        }
    }

    [Fact]
    public async Task AWriteThatReplacesAFormulaWithAConstantSaysSoByName()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var target = Path.Combine(directory, "clobber.xlsx");
        var existingExcel = ExcelTestWorkbook.SnapshotSettledExcel();

        try
        {
            ExcelTestWorkbook.CreateFormulaTarget(target, "A1:A3", new object?[,] { { 20 }, { "=R[-1]C*2" }, { 40 } });

            using var runtime = new ExcelWorkbookRuntime();
            var outcome = await runtime.ExecuteAsync(new ExcelTaskPlan("clobber", ExcelTaskPlans.Write(
                target, "Sheet1", [("A2", "99"), ("A3", "77")])), CancellationToken.None);

            // Replacing a formula with a constant is legitimate - hardcoding a figure is real finance
            // work - so this reports rather than refuses. What it must never do is destroy a live
            // calculation behind a receipt that says only "wrote 2 constants".
            Assert.True(outcome.Status == ExcelTaskStatus.Completed,
                $"{outcome.Summary} {string.Join("; ", outcome.Checks?.Select(check => check.Detail) ?? [])}");
            var replaced = Assert.Single(outcome.Checks!, check => check.Name == "formulas-replaced");
            Assert.False(replaced.Passed);
            Assert.Contains("A2", replaced.Detail, StringComparison.Ordinal);
            Assert.DoesNotContain("A3", replaced.Detail, StringComparison.Ordinal);
            Assert.True(ExcelTestWorkbook.HasValue(target, "A2", 99d));

            // What actually moved, not merely that something did. A simulation could tell its user a
            // formula had not been replaced but not what the cell had been, which is the next
            // question a person asks.
            var prior = Assert.Single(outcome.Checks!, check => check.Name == "prior-values");
            Assert.Contains("A3: 40 -> 77", prior.Detail, StringComparison.Ordinal);
        }
        finally
        {
            TempDirectory.Remove(directory);
            ExcelTestWorkbook.AssertNoLeakedExcel(existingExcel);
        }
    }

    [Fact]
    public async Task ACopySaveApplyReportsItsWholeChoreographyInTheOnlySafeOrder()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var target = Path.Combine(directory, "observed.xlsx");
        var output = Path.Combine(directory, "observed-out.xlsx");
        var existingExcel = ExcelTestWorkbook.SnapshotSettledExcel();

        try
        {
            ExcelTestWorkbook.CreateTarget(target);
            var observer = new RecordingRuntimeObserver();
            using var runtime = new ExcelWorkbookRuntime(observer);
            var outcome = await runtime.ExecuteAsync(new ExcelTaskPlan("observed", ExcelTaskPlans.Write(
                target, "Sheet1", [("A1", "1")], save: SaveMode.Copy, output: output)), CancellationToken.None);

            Assert.True(outcome.Status == ExcelTaskStatus.Completed,
                $"{outcome.Summary} {string.Join("; ", outcome.Checks?.Select(check => check.Detail) ?? [])}");

            // The ordering constraints that keep Excel and staging files from leaking, stated as
            // facts rather than trusted to convention. Every mutation path must report this same
            // shape, which is what makes the sequence safe to rewrite into one module.
            AssertMutationChoreography(observer, "value-write", expectStaging: true);
        }
        finally
        {
            TempDirectory.Remove(directory);
            ExcelTestWorkbook.AssertNoLeakedExcel(existingExcel);
        }
    }

    [Fact]
    public async Task ASameSaveApplyReportsTheSameChoreographyWithoutStaging()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var target = Path.Combine(directory, "observed-same.xlsx");
        var existingExcel = ExcelTestWorkbook.SnapshotSettledExcel();

        try
        {
            ExcelTestWorkbook.CreateFormulaTarget(target, "A1:A2", new object?[,] { { 1000.5 }, { 2000.25 } });
            var observer = new RecordingRuntimeObserver();
            using var runtime = new ExcelWorkbookRuntime(observer);
            var outcome = await runtime.ExecuteAsync(new ExcelTaskPlan("observed-same", ExcelTaskPlans.NumberFormat(
                target, "Sheet1", "A1:A2", "0.0%")), CancellationToken.None);

            Assert.True(outcome.Status == ExcelTaskStatus.Completed,
                $"{outcome.Summary} {string.Join("; ", outcome.Checks?.Select(check => check.Detail) ?? [])}");
            AssertMutationChoreography(observer, "number-format", expectStaging: false);
        }
        finally
        {
            TempDirectory.Remove(directory);
            ExcelTestWorkbook.AssertNoLeakedExcel(existingExcel);
        }
    }

    /// <summary>
    /// The invariants every Apply shares, whichever operation ran:
    /// the mutation happens after the session opens; the save follows the mutation; owned Excel
    /// cleanup is proven before verification opens the saved file (two processes on one file
    /// otherwise); a staging path, when one exists, is announced before anything saves into it,
    /// because the supervisor can only delete orphans it was told about; and both owned Excel
    /// processes - primary and pre-launched verification - are announced the moment they exist.
    /// </summary>
    private static void AssertMutationChoreography(RecordingRuntimeObserver observer, string mutatePhase, bool expectStaging)
    {
        Assert.True(observer.IndexOf("phase:session-open") < observer.IndexOf($"phase:{mutatePhase}"),
            "The mutation phase ran before the session opened.");
        Assert.True(observer.IndexOf($"phase:{mutatePhase}") < observer.IndexOf("phase:save"),
            "The save ran before the mutation.");
        Assert.True(observer.IndexOf("phase:save") < observer.IndexOf("phase:primary-cleanup"),
            "Cleanup began before the save.");
        Assert.True(observer.IndexOf("phase:primary-cleanup") < observer.IndexOf("phase:reopen-verification"),
            "Verification opened the file before owned Excel cleanup was proven.");

        if (expectStaging)
        {
            Assert.True(observer.IndexOf("staging-path") <= observer.IndexOf("phase:save") + 1 &&
                        observer.IndexOf("staging-path") < observer.IndexOf("phase:primary-cleanup"),
                "The staging path was not announced when the save began - an orphan there would be invisible to the supervisor.");
        }
        else
        {
            Assert.Equal(0, observer.CountOf("staging-path"));
        }

        // Exactly two owned processes for an Apply: the primary and the pre-launched verification.
        Assert.Equal(2, observer.CountOf("owned-process"));
    }

    private sealed class LockingRuntimeObserver(string outputPath) : IExcelWorkbookRuntimeObserver, IDisposable
    {
        private FileStream? _outputLock;

        public void OnPhase(string phase)
        {
            if (phase == "copy-promotion")
            {
                _outputLock = new FileStream(outputPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            }
        }

        public void OnOwnedProcessCaptured(ProcessIdentity identity) { }

        public void OnStagingPathCreated(string stagingPath) { }

        public void Dispose() => _outputLock?.Dispose();
    }

    [Fact]
    public async Task ScanReadsAWorkbookExcelItselfWroteAndStartsNoExcelToDoIt()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var target = Path.Combine(directory, "scan-real.xlsx");
        var existingExcel = ExcelTestWorkbook.SnapshotSettledExcel();

        try
        {
            // Built by real Excel - shared formulas, its namespaces, its quirks - because the fast
            // tests pin the parser against hand-authored XML and this pins it against the real thing.
            ExcelTestWorkbook.CreateFormulaTarget(target, "A1:A30", new object?[,]
            {
                { "=ROW()*2" }, { "=ROW()*2" }, { "=ROW()*2" }, { "=ROW()*2" }, { "=ROW()*2" },
                { "=ROW()*2" }, { "=ROW()*2" }, { "=ROW()*2" }, { "=ROW()*2" }, { "=ROW()*2" },
                { "=ROW()*2" }, { "=ROW()*2" }, { "=ROW()*2" }, { "=ROW()*2" }, { "=ROW()*2" },
                { "=ROW()*2" }, { "=ROW()*2" }, { "=ROW()*2" }, { "=ROW()*2" }, { "=ROW()*2" },
                { "=ROW()*2" }, { "=ROW()*2" }, { "=ROW()*2" }, { "=ROW()*2" }, { "=ROW()*2" },
                { "=ROW()*2" }, { "=ROW()*2" }, { "=ROW()*2" }, { "=ROW()*2" }, { "=ROW()*2" }
            }, constantCell: "A15", constantValue: 12345.67d);

            var observer = new RecordingRuntimeObserver();
            using var runtime = new ExcelWorkbookRuntime(observer);
            var outcome = await runtime.ExecuteAsync(
                new ExcelTaskPlan("scan-real", ExcelTaskPlans.Scan(target)), CancellationToken.None);

            Assert.True(outcome.Status == ExcelTaskStatus.Planned,
                $"{outcome.Summary} {string.Join("; ", outcome.Checks?.Select(check => check.Detail) ?? [])}");

            var sheet = Assert.Single(outcome.Audit!.Items, item => item.Kind == "scan-sheet");
            Assert.Contains("29 formula cell(s)", sheet.Detail, StringComparison.Ordinal);
            var island = Assert.Single(outcome.Audit.Items, item => item.Kind == "constant-island");
            Assert.Contains("A15", island.Detail, StringComparison.Ordinal);

            // The operation's whole claim, held at the observer seam: the scan started nothing.
            Assert.Equal(0, observer.CountOf("owned-process"));
        }
        finally
        {
            TempDirectory.Remove(directory);
            ExcelTestWorkbook.AssertNoLeakedExcel(existingExcel);
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
