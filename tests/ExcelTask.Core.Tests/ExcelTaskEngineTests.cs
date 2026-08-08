using System.Text.Json;
using ExcelTask.Core;

namespace ExcelTask.Core.Tests;

public sealed class ExcelTaskEngineTests
{
    [Fact]
    public async Task RunAsyncPlanCompilesNormalizedPlanAndPropagatesRuntimeReceipt()
    {
        var runtime = new FakeRuntime { Outcome = new(ExcelTaskStatus.Planned, "Plan ready", [new("worksheet", "Summary", "will add")]) };
        var engine = new ExcelTaskEngine(runtime);

        var receipt = await engine.RunAsync(Request(mode: ExcelTaskMode.Plan, save: SaveMode.Copy, output: ".\\out.xlsx"), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Planned, receipt.Status);
        Assert.Equal("Plan ready", receipt.Summary);
        Assert.Single(receipt.Changes);
        Assert.Equal(Path.GetFullPath(".\\target.xlsx"), runtime.Plan!.Request.TargetWorkbookPath);
        Assert.Equal("A1:C3", runtime.Plan.Request.FormulaRepairRanges[0].ToString());
        Assert.Equal(Path.GetFullPath(".\\out.xlsx"), runtime.Plan.Request.OutputWorkbookPath);
        Assert.False(receipt.Confirmation.Required);
    }

    [Fact]
    public async Task RunAsyncApplySameRequiresExplicitOverwriteConfirmation()
    {
        var runtime = new FakeRuntime();
        var engine = new ExcelTaskEngine(runtime);

        var receipt = await engine.RunAsync(Request(mode: ExcelTaskMode.Apply), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.NeedsConfirmation, receipt.Status);
        Assert.Contains(receipt.Confirmation.Requirements, requirement => requirement.Code == "overwrite-same");
        Assert.Null(runtime.Plan);
    }

    [Fact]
    public async Task RunAsyncOpenTargetWithAskBindingRequiresChoiceEvenWhenOverwriteIsConfirmed()
    {
        var runtime = new FakeRuntime { Inspection = new(true, OpenWorkbookDescription: "target.xlsx is open") };
        var engine = new ExcelTaskEngine(runtime);

        var receipt = await engine.RunAsync(Request(mode: ExcelTaskMode.Apply, overwriteConfirmed: true), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.NeedsConfirmation, receipt.Status);
        Assert.Contains(receipt.Confirmation.Requirements, requirement => requirement.Code == "target-open");
        Assert.Null(runtime.Plan);
    }

    [Fact]
    public async Task RunAsyncApplyDelegatesAfterConfirmationsAndPreservesPartialOutcome()
    {
        var runtime = new FakeRuntime
        {
            Outcome = new(ExcelTaskStatus.Partial, "One repair was skipped", [], [new("repairs", false, "C5 contains a constant")], true, "Review C5")
        };
        var engine = new ExcelTaskEngine(runtime);

        var receipt = await engine.RunAsync(Request(mode: ExcelTaskMode.Apply, overwriteConfirmed: true), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Partial, receipt.Status);
        Assert.NotNull(runtime.Plan);
        Assert.False(receipt.Retry.CanRetry);
        Assert.Contains("reconcile", receipt.Retry.Reason, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(receipt.Checks, check => check.Name == "repairs");
    }

    [Theory]
    [InlineData("", "Reference", "New", "A1")]
    [InlineData("target.xlsx", "Bad/Sheet", "New", "A1")]
    [InlineData("target.xlsx", "Reference", "ThisNameIsLongerThanThirtyOneChars", "A1")]
    [InlineData("target.xlsx", "Reference", "New", "A0")]
    [InlineData("target.xlsx", "Reference", "New", "B2:A1")]
    public async Task RunAsyncRejectsInvalidPathsNamesAndRanges(string target, string referenceSheet, string newSheet, string range)
    {
        var runtime = new FakeRuntime();
        var engine = new ExcelTaskEngine(runtime);
        var request = Request() with { TargetWorkbookPath = target, ReferenceWorksheet = referenceSheet, NewWorksheetName = newSheet, FormulaRepairRanges = [range] };

        var receipt = await engine.RunAsync(request, CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Rejected, receipt.Status);
        Assert.Null(runtime.InspectionRequest);
        Assert.False(receipt.Checks.Single().Passed);
    }

    [Fact]
    public async Task RunAsyncReturnsUnknownWhenRuntimeExecutionThrows()
    {
        var runtime = new FakeRuntime { ExecuteException = new InvalidOperationException("Excel disconnected") };
        var engine = new ExcelTaskEngine(runtime);

        var receipt = await engine.RunAsync(Request(mode: ExcelTaskMode.Plan), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Unknown, receipt.Status);
        Assert.False(receipt.Retry.CanRetry);
        Assert.Contains("Reconcile", receipt.Retry.Reason);
        Assert.Contains(receipt.Checks, check => check.Name == "runtime-execution");
    }

    [Fact]
    public async Task RunAsyncForcesUnknownRuntimeOutcomesToBeNonRetryable()
    {
        var runtime = new FakeRuntime { Outcome = new(ExcelTaskStatus.Unknown, "Runtime did not verify completion", CanRetry: true, RetryReason: "Retry it") };
        var engine = new ExcelTaskEngine(runtime);

        var receipt = await engine.RunAsync(Request(), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Unknown, receipt.Status);
        Assert.False(receipt.Retry.CanRetry);
        Assert.Contains("Reconcile", receipt.Retry.Reason);
    }

    [Fact]
    public async Task RunAsyncReturnsRejectedSafeToRetryWhenInspectionFailsBeforeDispatch()
    {
        var runtime = new FakeRuntime { InspectionException = new InvalidOperationException("Sensitive runtime detail") };
        var engine = new ExcelTaskEngine(runtime);

        var receipt = await engine.RunAsync(Request(), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Rejected, receipt.Status);
        Assert.True(receipt.Retry.CanRetry);
        Assert.DoesNotContain("Sensitive runtime detail", receipt.Summary);
        Assert.DoesNotContain(receipt.Checks, check => check.Detail.Contains("Sensitive runtime detail", StringComparison.Ordinal));
        Assert.Null(runtime.Plan);
    }

    [Fact]
    public async Task RunAsyncApplyCopyRequiresConfirmationOnlyWhenOutputAlreadyExists()
    {
        var runtime = new FakeRuntime { Inspection = new(false, CopyOutputExists: true) };
        var engine = new ExcelTaskEngine(runtime);

        var receipt = await engine.RunAsync(Request(mode: ExcelTaskMode.Apply, save: SaveMode.Copy, output: ".\\existing-copy.xlsx"), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.NeedsConfirmation, receipt.Status);
        Assert.Contains(receipt.Confirmation.Requirements, requirement => requirement.Code == "overwrite-copy");
        Assert.Equal(SaveMode.Copy, runtime.InspectionRequest!.Save);
        Assert.Equal(Path.GetFullPath(".\\existing-copy.xlsx"), runtime.InspectionRequest.OutputWorkbookPath);
        Assert.Null(runtime.Plan);
    }

    [Fact]
    public async Task RunAsyncApplyCopyDoesNotRequireOverwriteConfirmationForNewOutput()
    {
        var runtime = new FakeRuntime { Inspection = new(false, CopyOutputExists: false) };
        var engine = new ExcelTaskEngine(runtime);

        var receipt = await engine.RunAsync(Request(mode: ExcelTaskMode.Apply, save: SaveMode.Copy, output: ".\\new-copy.xlsx"), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Completed, receipt.Status);
        Assert.NotNull(runtime.Plan);
    }

    [Fact]
    public async Task RunAsyncRejectsUseOpenWhenTargetIsNotOpen()
    {
        var runtime = new FakeRuntime { Inspection = new(false) };
        var engine = new ExcelTaskEngine(runtime);

        var receipt = await engine.RunAsync(Request(binding: WorkbookBinding.UseOpen), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Rejected, receipt.Status);
        Assert.Contains("requires", receipt.Summary, StringComparison.OrdinalIgnoreCase);
        Assert.Null(runtime.Plan);
    }

    [Fact]
    public async Task RunAsyncRejectsIsolatedApplySameWhenTargetIsOpen()
    {
        var runtime = new FakeRuntime { Inspection = new(true) };
        var engine = new ExcelTaskEngine(runtime);

        var receipt = await engine.RunAsync(Request(mode: ExcelTaskMode.Apply, binding: WorkbookBinding.Isolated, overwriteConfirmed: true), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Rejected, receipt.Status);
        Assert.Contains("UseOpen or save a Copy", receipt.Summary);
        Assert.Null(runtime.Plan);
    }

    [Theory]
    [InlineData("mode")]
    [InlineData("binding")]
    [InlineData("save")]
    public async Task RunAsyncRejectsUndefinedEnumValuesBeforeInspection(string invalidField)
    {
        var runtime = new FakeRuntime();
        var engine = new ExcelTaskEngine(runtime);
        var request = invalidField switch
        {
            "mode" => Request() with { Mode = (ExcelTaskMode)99 },
            "binding" => Request() with { WorkbookBinding = (WorkbookBinding)99 },
            _ => Request() with { Save = (SaveMode)99 }
        };

        var receipt = await engine.RunAsync(request, CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Rejected, receipt.Status);
        Assert.Null(runtime.InspectionRequest);
    }

    [Fact]
    public async Task RunAsyncBoundsRuntimeSuppliedReceiptDataAndDoesNotExposeOutputPath()
    {
        var longText = new string('x', 600);
        var runtime = new FakeRuntime
        {
            Outcome = new(
                ExcelTaskStatus.Completed,
                longText,
                Enumerable.Range(0, 21).Select(index => new TaskChange(longText, longText, longText)).ToArray(),
                Enumerable.Range(0, 21).Select(index => new TaskCheck(longText, true, longText)).ToArray(),
                true,
                longText)
        };
        var engine = new ExcelTaskEngine(runtime);

        var receipt = await engine.RunAsync(Request(save: SaveMode.Copy, output: ".\\private\\out.xlsx"), CancellationToken.None);

        Assert.Equal(256, receipt.Summary.Length);
        Assert.Equal(20, receipt.Changes.Count);
        Assert.Equal(20, receipt.Checks.Count);
        Assert.All(receipt.Changes, change => Assert.All([change.Kind, change.Target, change.Summary], value => Assert.Equal(256, value.Length)));
        Assert.All(receipt.Checks, check => Assert.All([check.Name, check.Detail], value => Assert.InRange(value.Length, 0, 256)));
        Assert.Equal(256, receipt.Retry.Reason!.Length);
        Assert.Equal("out.xlsx", receipt.Save.OutputWorkbookPath);
        Assert.True(JsonSerializer.SerializeToUtf8Bytes(receipt).Length < 32 * 1024);
    }

    [Theory]
    [InlineData(".\\target.xlsb", ".\\reference.xlsx", null, "MVP workbook paths")]
    [InlineData(".\\target.xlsx", ".\\reference.xlsb", null, "MVP workbook paths")]
    [InlineData(".\\target.xlsx", ".\\reference.xlsx", ".\\out.xlsm", "Copy output extension")]
    public async Task RunAsyncRejectsUnsupportedOrMismatchedWorkbookExtensions(string target, string reference, string? output, string expectedError)
    {
        var runtime = new FakeRuntime();
        var engine = new ExcelTaskEngine(runtime);
        var request = Request(save: output is null ? SaveMode.Same : SaveMode.Copy, output: output) with
        {
            TargetWorkbookPath = target,
            ReferenceWorkbookPath = reference
        };

        var receipt = await engine.RunAsync(request, CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Rejected, receipt.Status);
        Assert.Contains(expectedError, receipt.Summary);
        Assert.Null(runtime.InspectionRequest);
    }

    [Fact]
    public async Task RunAsyncRejectsUseOpenCopyBeforeInspection()
    {
        var runtime = new FakeRuntime();
        var engine = new ExcelTaskEngine(runtime);

        var receipt = await engine.RunAsync(Request(save: SaveMode.Copy, output: ".\\out.xlsx", binding: WorkbookBinding.UseOpen), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Rejected, receipt.Status);
        Assert.Null(runtime.InspectionRequest);
    }

    [Fact]
    public async Task RunAsyncAcceptsFormulaRepairRangesAtTheMvpCellCap()
    {
        var runtime = new FakeRuntime();
        var engine = new ExcelTaskEngine(runtime);
        var request = Request() with { FormulaRepairRanges = ["A1:CV100"] };

        var receipt = await engine.RunAsync(request, CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Completed, receipt.Status);
        Assert.NotNull(runtime.Plan);
        Assert.Equal(ExcelTaskEngine.MaxFormulaRepairCells, 100 * 100);
    }

    [Fact]
    public async Task RunAsyncRejectsFormulaRepairRangesOverTheAggregateMvpCellCap()
    {
        var runtime = new FakeRuntime();
        var engine = new ExcelTaskEngine(runtime);
        var request = Request() with { FormulaRepairRanges = ["A1:CV100", "CW1"] };

        var receipt = await engine.RunAsync(request, CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Rejected, receipt.Status);
        Assert.Contains("10,000 aggregate cells", receipt.Summary);
        Assert.Null(runtime.InspectionRequest);
    }

    private static ExcelTaskRequest Request(
        ExcelTaskMode mode = ExcelTaskMode.Plan,
        SaveMode save = SaveMode.Same,
        string? output = null,
        bool overwriteConfirmed = false,
        WorkbookBinding binding = WorkbookBinding.AskIfOpen) => new(
            ".\\target.xlsx",
            ".\\reference.xlsx",
            "Reference",
            "New sheet",
            ["$a$1:$c$3"],
            mode,
            binding,
            save,
            output,
            overwriteConfirmed);

    private sealed class FakeRuntime : IWorkbookRuntime
    {
        public WorkbookInspection Inspection { get; init; } = new(false, Checks: [new("runtime-inspection", true, "available")]);
        public Exception? InspectionException { get; init; }
        public WorkbookExecutionOutcome Outcome { get; init; } = new(ExcelTaskStatus.Completed, "Completed");
        public Exception? ExecuteException { get; init; }
        public WorkbookInspectionRequest? InspectionRequest { get; private set; }
        public ExcelTaskPlan? Plan { get; private set; }

        public Task<WorkbookInspection> InspectAsync(WorkbookInspectionRequest request, CancellationToken cancellationToken)
        {
            InspectionRequest = request;
            if (InspectionException is not null) throw InspectionException;
            return Task.FromResult(Inspection);
        }

        public Task<WorkbookExecutionOutcome> ExecuteAsync(ExcelTaskPlan plan, CancellationToken cancellationToken)
        {
            Plan = plan;
            if (ExecuteException is not null) throw ExecuteException;
            return Task.FromResult(Outcome);
        }
    }
}
