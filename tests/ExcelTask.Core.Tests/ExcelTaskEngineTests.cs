using System.Text.Json;
using System.ComponentModel;
using ExcelTask.Core;

namespace ExcelTask.Core.Tests;

public sealed class ExcelTaskEngineTests
{
    [Fact]
    public void EveryModelFacingInputFieldHasADescription()
    {
        var modelTypes = new[]
        {
            typeof(ExcelTaskRequest),
            typeof(ExcelOperation),
            typeof(CopyExhibitOperation),
            typeof(RepairExistingWorksheetOperation),
            typeof(ExtendFormulaSeriesOperation)
        };

        foreach (var property in modelTypes.SelectMany(type => type.GetProperties()))
        {
            var description = property.GetCustomAttributes(typeof(DescriptionAttribute), inherit: true)
                .Cast<DescriptionAttribute>()
                .SingleOrDefault();
            Assert.False(string.IsNullOrWhiteSpace(description?.Description), $"{property.DeclaringType!.Name}.{property.Name} is missing a description.");
        }
    }

    [Fact]
    public async Task PlanCopyExhibitDispatchesCleanNormalizedCopyShape()
    {
        var runtime = new FakeRuntime { Outcome = new(ExcelTaskStatus.Planned, "Plan ready") };
        var receipt = await new ExcelTaskEngine(runtime).RunAsync(Request(save: SaveMode.Copy, output: ".\\out.xlsx"), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Planned, receipt.Status);
        var copy = Assert.IsType<NormalizedCopyExhibitOperation>(runtime.Plan!.Request.Operation.CopyExhibit);
        Assert.Equal(Path.GetFullPath(".\\reference.xlsx"), copy.ReferenceWorkbookPath);
        Assert.Equal("A1:C3", copy.RepairRanges[0].ToString());
        Assert.Null(runtime.Plan.Request.Operation.RepairExistingWorksheet);
        Assert.Equal(Path.GetFullPath(".\\reference.xlsx"), runtime.InspectionRequest!.ReferenceWorkbookPath);
    }

    [Fact]
    public async Task PlanRepairExistingWorksheetDispatchesNoReferenceWorkbook()
    {
        var runtime = new FakeRuntime();
        var task = new ExcelOperation(ExcelOperationKind.RepairExistingWorksheet,
            RepairExistingWorksheet: new("Model", ["$B$2:$C$3"]));

        var receipt = await new ExcelTaskEngine(runtime).RunAsync(Request(task), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Completed, receipt.Status);
        var repair = Assert.IsType<NormalizedRepairExistingWorksheetOperation>(runtime.Plan!.Request.Operation.RepairExistingWorksheet);
        Assert.Equal("Model", repair.WorksheetName);
        Assert.Equal("B2:C3", repair.Ranges.Single().ToString());
        Assert.Null(runtime.InspectionRequest!.ReferenceWorkbookPath);
    }

    [Fact]
    public async Task PlanExtensionDispatchesNormalizedRangesAndDirection()
    {
        var runtime = new FakeRuntime();
        var task = new ExcelOperation(ExcelOperationKind.ExtendFormulaSeries,
            ExtendFormulaSeries: new("Model", FormulaExtensionDirection.Right, "$B$2:$C$4", "D2:F4"));

        var receipt = await new ExcelTaskEngine(runtime).RunAsync(Request(task), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Completed, receipt.Status);
        var extension = Assert.IsType<NormalizedExtendFormulaSeriesOperation>(runtime.Plan!.Request.Operation.ExtendFormulaSeries);
        Assert.Equal(FormulaExtensionDirection.Right, extension.Direction);
        Assert.Equal("B2:C4", extension.EvidenceRange.ToString());
        Assert.Equal("D2:F4", extension.DestinationRange.ToString());
    }

    [Fact]
    public async Task ApplySameRequiresExplicitOverwriteConfirmation()
    {
        var runtime = new FakeRuntime();
        var receipt = await new ExcelTaskEngine(runtime).RunAsync(Request(mode: ExcelTaskMode.Apply), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.NeedsConfirmation, receipt.Status);
        Assert.Contains(receipt.Confirmation.Requirements, requirement => requirement.Code == "overwrite-same");
        Assert.Null(runtime.Plan);
    }

    [Fact]
    public async Task OmittedPolicyDefaultsNormalizeToApplyAskIfOpenSameAndFalse()
    {
        var runtime = new FakeRuntime();
        var request = new ExcelTaskRequest(".\\target.xlsx", new ExcelOperation(ExcelOperationKind.CopyExhibit, CopyExhibit: Copy()));

        var receipt = await new ExcelTaskEngine(runtime).RunAsync(request, CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.NeedsConfirmation, receipt.Status);
        Assert.Contains(receipt.Confirmation.Requirements, requirement => requirement.Code == "overwrite-same");
        Assert.Equal(WorkbookBinding.AskIfOpen, runtime.InspectionRequest!.Binding);
        Assert.Equal(SaveMode.Same, runtime.InspectionRequest.Save);
        Assert.Equal(SaveMode.Same, receipt.Save.Mode);
        Assert.False(receipt.Save.OverwriteConfirmed);
    }

    [Theory]
    [MemberData(nameof(InvalidOperations))]
    public async Task RejectsMissingUnknownMultipleAndMismatchedOperationUnion(ExcelOperation? operation)
    {
        var runtime = new FakeRuntime();
        var request = operation is null ? Request() with { Operation = null! } : Request(operation);
        var receipt = await new ExcelTaskEngine(runtime).RunAsync(request, CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Rejected, receipt.Status);
        Assert.Null(runtime.InspectionRequest);
        Assert.False(receipt.Checks.Single().Passed);
    }

    public static IEnumerable<object?[]> InvalidOperations()
    {
        yield return [null];
        yield return [new ExcelOperation((ExcelOperationKind)99, CopyExhibit: Copy())];
        yield return [new ExcelOperation(ExcelOperationKind.CopyExhibit)];
        yield return [new ExcelOperation(ExcelOperationKind.CopyExhibit, RepairExistingWorksheet: new("Model", []))];
        yield return [new ExcelOperation(ExcelOperationKind.CopyExhibit, CopyExhibit: Copy(), RepairExistingWorksheet: new("Model", []))];
        yield return [new ExcelOperation(ExcelOperationKind.ExtendFormulaSeries, CopyExhibit: Copy())];
    }

    [Theory]
    [InlineData("", "Reference", "New", "A1")]
    [InlineData(".\\reference.xlsx", "Bad/Sheet", "New", "A1")]
    [InlineData(".\\reference.xlsx", "Reference", "ThisNameIsLongerThanThirtyOneChars", "A1")]
    [InlineData(".\\reference.xlsx", "Reference", "New", "B2:A1")]
    public async Task RejectsInvalidCopyPathNamesAndRanges(string reference, string referenceSheet, string newSheet, string range)
    {
        var runtime = new FakeRuntime();
        var task = new ExcelOperation(ExcelOperationKind.CopyExhibit, CopyExhibit: new(reference, referenceSheet, newSheet, [range]));

        var receipt = await new ExcelTaskEngine(runtime).RunAsync(Request(task), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Rejected, receipt.Status);
        Assert.Null(runtime.InspectionRequest);
    }

    [Fact]
    public async Task RejectsOverlappingRepairRangesBeforeInspection()
    {
        var runtime = new FakeRuntime();
        var task = new ExcelOperation(ExcelOperationKind.RepairExistingWorksheet,
            RepairExistingWorksheet: new("Model", ["A1:C3", "C3:D4"]));

        var receipt = await new ExcelTaskEngine(runtime).RunAsync(Request(task), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Rejected, receipt.Status);
        Assert.Contains("must not overlap", receipt.Summary);
        Assert.Null(runtime.InspectionRequest);
    }

    [Fact]
    public async Task RejectsEmptyRepairExistingWorksheetRangesBeforeInspection()
    {
        var runtime = new FakeRuntime();
        var operation = new ExcelOperation(ExcelOperationKind.RepairExistingWorksheet,
            RepairExistingWorksheet: new("Model", []));

        var receipt = await new ExcelTaskEngine(runtime).RunAsync(Request(operation), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Rejected, receipt.Status);
        Assert.Contains("one or more repair ranges", receipt.Summary);
        Assert.Null(runtime.InspectionRequest);
    }

    [Fact]
    public async Task RejectsMoreThanSixteenRepairRangesBeforeInspection()
    {
        var runtime = new FakeRuntime();
        var ranges = Enumerable.Range(1, 17).Select(row => $"A{row}").ToArray();
        var task = new ExcelOperation(ExcelOperationKind.RepairExistingWorksheet,
            RepairExistingWorksheet: new("Model", ranges));

        var receipt = await new ExcelTaskEngine(runtime).RunAsync(Request(task), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Rejected, receipt.Status);
        Assert.Contains("16 ranges", receipt.Summary);
        Assert.Null(runtime.InspectionRequest);
    }

    [Fact]
    public async Task AcceptsRepairRangesAtAggregateScanCap()
    {
        var runtime = new FakeRuntime();
        var task = new ExcelOperation(ExcelOperationKind.RepairExistingWorksheet,
            RepairExistingWorksheet: new("Model", ["A1:CV100"]));

        var receipt = await new ExcelTaskEngine(runtime).RunAsync(Request(task), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Completed, receipt.Status);
        Assert.Equal("A1:CV100", runtime.Plan!.Request.Operation.RepairExistingWorksheet!.Ranges.Single().ToString());
    }

    [Fact]
    public async Task RejectsRepairRangesAboveAggregateScanCap()
    {
        var runtime = new FakeRuntime();
        var task = new ExcelOperation(ExcelOperationKind.RepairExistingWorksheet,
            RepairExistingWorksheet: new("Model", ["A1:CV100", "CW1"]));

        var receipt = await new ExcelTaskEngine(runtime).RunAsync(Request(task), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Rejected, receipt.Status);
        Assert.Contains("10,000 aggregate cells", receipt.Summary);
    }

    [Theory]
    [InlineData(FormulaExtensionDirection.Right, "B2:C4", "D2:AA4")]
    [InlineData(FormulaExtensionDirection.Down, "B2:D3", "B4:D27")]
    public async Task ExtensionAcceptsMaximumTwentyFourPeriods(FormulaExtensionDirection direction, string evidence, string destination)
    {
        var runtime = new FakeRuntime();
        var task = new ExcelOperation(ExcelOperationKind.ExtendFormulaSeries,
            ExtendFormulaSeries: new("Model", direction, evidence, destination));

        var receipt = await new ExcelTaskEngine(runtime).RunAsync(Request(task), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Completed, receipt.Status);
    }

    [Theory]
    [InlineData(FormulaExtensionDirection.Right, "B2:C3", "E2:F3")]
    [InlineData(FormulaExtensionDirection.Right, "B2:D3", "E2:F3")]
    [InlineData(FormulaExtensionDirection.Down, "B2:C3", "B5:C6")]
    [InlineData(FormulaExtensionDirection.Down, "B2:C3", "B4:D5")]
    public async Task ExtensionRejectsNonAdjacentOrInvalidGeometry(FormulaExtensionDirection direction, string evidence, string destination)
    {
        var runtime = new FakeRuntime();
        var task = new ExcelOperation(ExcelOperationKind.ExtendFormulaSeries,
            ExtendFormulaSeries: new("Model", direction, evidence, destination));

        var receipt = await new ExcelTaskEngine(runtime).RunAsync(Request(task), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Rejected, receipt.Status);
        Assert.Null(runtime.InspectionRequest);
    }

    [Fact]
    public async Task ExtensionRejectsDestinationAboveTwoThousandCells()
    {
        var runtime = new FakeRuntime();
        var task = new ExcelOperation(ExcelOperationKind.ExtendFormulaSeries,
            ExtendFormulaSeries: new("Model", FormulaExtensionDirection.Right, "B1:C100", "D1:AA100"));

        var receipt = await new ExcelTaskEngine(runtime).RunAsync(Request(task), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Rejected, receipt.Status);
        Assert.Contains("2,000", receipt.Summary);
    }

    [Fact]
    public async Task ExtensionRejectsUnknownDirectionBeforeInspection()
    {
        var runtime = new FakeRuntime();
        var task = new ExcelOperation(ExcelOperationKind.ExtendFormulaSeries,
            ExtendFormulaSeries: new("Model", (FormulaExtensionDirection)99, "B1:C1", "D1:E1"));

        var receipt = await new ExcelTaskEngine(runtime).RunAsync(Request(task), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Rejected, receipt.Status);
        Assert.Null(runtime.InspectionRequest);
    }

    [Theory]
    [InlineData("mode")]
    [InlineData("binding")]
    [InlineData("save")]
    public async Task RejectsUndefinedOuterEnumsBeforeInspection(string invalidField)
    {
        var runtime = new FakeRuntime();
        var request = invalidField switch
        {
            "mode" => Request() with { Mode = (ExcelTaskMode)99 },
            "binding" => Request() with { WorkbookBinding = (WorkbookBinding)99 },
            _ => Request() with { Save = (SaveMode)99 }
        };

        var receipt = await new ExcelTaskEngine(runtime).RunAsync(request, CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Rejected, receipt.Status);
        Assert.Null(runtime.InspectionRequest);
    }

    [Fact]
    public async Task ReturnsUnknownWhenRuntimeExecutionThrows()
    {
        var runtime = new FakeRuntime { ExecuteException = new InvalidOperationException("Excel disconnected") };
        var receipt = await new ExcelTaskEngine(runtime).RunAsync(Request(), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Unknown, receipt.Status);
        Assert.False(receipt.Retry.CanRetry);
        Assert.Contains("Reconcile", receipt.Retry.Reason);
    }

    [Fact]
    public async Task BoundsRuntimeReceiptDataAndDoesNotExposeOutputDirectory()
    {
        var longText = new string('x', 600);
        var runtime = new FakeRuntime
        {
            Outcome = new(ExcelTaskStatus.Completed, longText,
            Enumerable.Range(0, 21).Select(_ => new TaskChange(longText, longText, longText)).ToArray(),
            Enumerable.Range(0, 21).Select(_ => new TaskCheck(longText, true, longText)).ToArray(), true, longText)
        };

        var receipt = await new ExcelTaskEngine(runtime).RunAsync(Request(save: SaveMode.Copy, output: ".\\private\\out.xlsx"), CancellationToken.None);

        Assert.Equal(256, receipt.Summary.Length);
        Assert.Equal(20, receipt.Changes.Count);
        Assert.Equal("out.xlsx", receipt.Save.OutputWorkbookPath);
        Assert.True(JsonSerializer.SerializeToUtf8Bytes(receipt).Length < 32 * 1024);
    }

    private static ExcelTaskRequest Request(
        ExcelOperation? operation = null,
        ExcelTaskMode mode = ExcelTaskMode.Plan,
        SaveMode save = SaveMode.Same,
        string? output = null,
        bool overwriteConfirmed = false,
        WorkbookBinding binding = WorkbookBinding.AskIfOpen) => new(
            TargetWorkbookPath: ".\\target.xlsx",
            Operation: operation ?? new ExcelOperation(ExcelOperationKind.CopyExhibit, CopyExhibit: Copy()),
            Mode: mode,
            WorkbookBinding: binding,
            Save: save,
            OutputWorkbookPath: output,
            OverwriteConfirmed: overwriteConfirmed);

    private static CopyExhibitOperation Copy() => new(".\\reference.xlsx", "Reference", "New sheet", ["$a$1:$c$3"]);

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
