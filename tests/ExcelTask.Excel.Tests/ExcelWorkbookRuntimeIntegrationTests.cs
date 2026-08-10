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
        var existingExcel = OwnedExcelProcess.SnapshotExcelProcesses();
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
            var remainingExcel = OwnedExcelProcess.SnapshotExcelProcesses();
            var leaked = remainingExcel.Where(process => !existingExcel.Contains(process)).ToArray();
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
            Assert.Empty(leaked);
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
        var existingExcel = OwnedExcelProcess.SnapshotExcelProcesses();

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
            var remainingExcel = OwnedExcelProcess.SnapshotExcelProcesses();
            var leaked = remainingExcel.Where(process => !existingExcel.Contains(process)).ToArray();
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
            Assert.Empty(leaked);
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
        var existingExcel = OwnedExcelProcess.SnapshotExcelProcesses();

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
            var remainingExcel = OwnedExcelProcess.SnapshotExcelProcesses();
            var leaked = remainingExcel.Where(process => !existingExcel.Contains(process)).ToArray();
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
            Assert.Empty(leaked);
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
        var existingExcel = OwnedExcelProcess.SnapshotExcelProcesses();

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
            var remainingExcel = OwnedExcelProcess.SnapshotExcelProcesses();
            var leaked = remainingExcel.Where(process => !existingExcel.Contains(process)).ToArray();
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
            Assert.Empty(leaked);
        }
    }

    // A dialog offering a real choice - Yes/No, OK/Cancel, Retry/Cancel - is deliberately left
    // alone, so the blocked automation call never returns. Only the supervised runtime bounds that,
    // by its deadline; this direct-runtime layer has none and would hang forever. The exclusion is
    // therefore covered by ModalDialogSentryTests against real dialogs rather than here.

    [Fact]
    public async Task AuditReportsFlowsWithoutChangingTheWorkbookOrLeakingPaths()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var target = Path.Combine(directory, "audit-target.xlsx");
        var reference = Path.Combine(directory, "audit-reference.xlsx");
        var existingExcel = OwnedExcelProcess.SnapshotExcelProcesses();

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
            var remainingExcel = OwnedExcelProcess.SnapshotExcelProcesses();
            var leaked = remainingExcel.Where(process => !existingExcel.Contains(process)).ToArray();
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
            Assert.Empty(leaked);
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
