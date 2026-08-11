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
            typeof(ExtendFormulaSeriesOperation),
            typeof(EditMacroProcedureOperation),
            typeof(ReadWorksheetRangeOperation),
            typeof(WriteWorksheetValuesOperation),
            typeof(WorksheetCellValue),
            typeof(FindReplaceOperation),
            typeof(CreateOperation)
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
    public async Task PlanMacroProcedureDispatchesCleanInspectOnlyNormalizedShape()
    {
        var runtime = new FakeRuntime { Outcome = new(ExcelTaskStatus.Planned, "Macro plan ready") };
        var receipt = await new ExcelTaskEngine(runtime).RunAsync(MacroRequest(), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Planned, receipt.Status);
        var macro = Assert.IsType<NormalizedEditMacroProcedureOperation>(runtime.Plan!.Request.Operation.EditMacroProcedure);
        Assert.Equal("StandardModule", macro.ComponentName);
        Assert.Equal("RefreshModel", macro.ProcedureName);
        Assert.Null(macro.ExpectedProcedureSha256);
        Assert.Null(macro.ReplacementSource);
        Assert.False(macro.RunAfterEdit);
        Assert.Null(runtime.InspectionRequest!.ReferenceWorkbookPath);
        Assert.Equal(WorkbookBinding.Isolated, runtime.InspectionRequest.Binding);
        Assert.Equal(SaveMode.Copy, runtime.InspectionRequest.Save);
    }

    [Fact]
    public async Task PlanReturnsTheProcedureSourceEvenWhenItsSignatureIsNotReplacementParseable()
    {
        // A line-continuation signature is valid VBA that the bounded replacement parser rejects.
        // Plan source is evidence read out of the workbook, not model input, so dropping it would
        // leave the caller with a successful plan and nothing to edit.
        const string source = "Public Sub RefreshModel( _\n    ByVal scope As Long)\n    Debug.Print scope\nEnd Sub";
        var runtime = new FakeRuntime
        {
            Outcome = new(
                ExcelTaskStatus.Planned,
                "Macro plan ready",
                MacroProcedure: new MacroProcedureReceipt("StandardModule", "RefreshModel", new string('a', 64), source, false, false))
        };

        var receipt = await new ExcelTaskEngine(runtime).RunAsync(MacroRequest(), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Planned, receipt.Status);
        Assert.Equal(source, receipt.MacroProcedure!.Source);
    }

    [Fact]
    public async Task ApplyNeverEchoesMacroSourceBackToTheCaller()
    {
        var runtime = new FakeRuntime
        {
            Outcome = new(
                ExcelTaskStatus.Completed,
                "Macro applied",
                MacroProcedure: new MacroProcedureReceipt("StandardModule", "RefreshModel", new string('a', 64), "Public Sub RefreshModel()\nEnd Sub", false, false))
        };

        var receipt = await new ExcelTaskEngine(runtime).RunAsync(
            MacroRequest(
                mode: ExcelTaskMode.Apply,
                operation: Macro(expectedHash: new string('A', 64), replacementSource: "Public Sub RefreshModel()\nEnd Sub")),
            CancellationToken.None);

        Assert.Null(receipt.MacroProcedure!.Source);
    }

    [Fact]
    public async Task ApplyMacroProcedureNormalizesHashAndReplacementSource()
    {
        const string source = "Public Sub RefreshModel()\r\nEnd Sub\r\n";
        var runtime = new FakeRuntime();
        var receipt = await new ExcelTaskEngine(runtime).RunAsync(
            MacroRequest(
                mode: ExcelTaskMode.Apply,
                operation: Macro(expectedHash: new string('A', 64), replacementSource: source)),
            CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Completed, receipt.Status);
        var macro = Assert.IsType<NormalizedEditMacroProcedureOperation>(runtime.Plan!.Request.Operation.EditMacroProcedure);
        Assert.Equal(new string('a', 64), macro.ExpectedProcedureSha256);
        Assert.Equal("Public Sub RefreshModel()\nEnd Sub", macro.ReplacementSource);
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
        yield return [new ExcelOperation(ExcelOperationKind.EditMacroProcedure, CopyExhibit: Copy())];
        yield return [new ExcelOperation(ExcelOperationKind.EditMacroProcedure, EditMacroProcedure: Macro().EditMacroProcedure, ExtendFormulaSeries: new("Model", FormulaExtensionDirection.Right, "A1:B1", "C1:D1"))];
    }

    [Theory]
    [InlineData(true, false, false)]
    [InlineData(false, true, false)]
    [InlineData(false, false, true)]
    public async Task MacroPlanRejectsAnyMutationInputsBeforeInspection(bool expectedHash, bool replacementSource, bool runAfterEdit)
    {
        var operation = Macro(
            expectedHash: expectedHash ? new string('a', 64) : null,
            replacementSource: replacementSource ? "Sub RefreshModel()\nEnd Sub" : null,
            runAfterEdit: runAfterEdit);
        var runtime = new FakeRuntime();

        var receipt = await new ExcelTaskEngine(runtime).RunAsync(MacroRequest(operation: operation), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Rejected, receipt.Status);
        Assert.Null(runtime.InspectionRequest);
    }

    [Theory]
    [InlineData(".\\target.xlsx", WorkbookBinding.Isolated, SaveMode.Copy)]
    [InlineData(".\\target.xlsm", WorkbookBinding.AskIfOpen, SaveMode.Copy)]
    [InlineData(".\\target.xlsm", WorkbookBinding.Isolated, SaveMode.Same)]
    public async Task MacroPolicyRejectsWorkbookFormatBindingAndSaveBeforeInspection(string target, WorkbookBinding binding, SaveMode save)
    {
        var runtime = new FakeRuntime();
        var request = MacroRequest(
            mode: ExcelTaskMode.Apply,
            operation: Macro(expectedHash: new string('a', 64), replacementSource: "Sub RefreshModel()\nEnd Sub"),
            binding: binding,
            save: save) with
        {
            TargetWorkbookPath = target,
            OutputWorkbookPath = save == SaveMode.Copy ? ".\\out.xlsm" : null
        };

        var receipt = await new ExcelTaskEngine(runtime).RunAsync(request, CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Rejected, receipt.Status);
        Assert.Null(runtime.InspectionRequest);
    }

    [Fact]
    public async Task MacroPolicyNamesEveryUnmetRequirementInOneRejection()
    {
        // Field use showed a caller rejected twice, learning one rule per round trip. Every unmet
        // requirement must arrive in the first rejection so the second attempt is the correct one.
        var runtime = new FakeRuntime();
        var request = MacroRequest(
            mode: ExcelTaskMode.Apply,
            operation: Macro(expectedHash: new string('a', 64), replacementSource: "Sub RefreshModel()\nEnd Sub"),
            binding: WorkbookBinding.AskIfOpen,
            save: SaveMode.Same) with
        {
            TargetWorkbookPath = ".\\target.xlsx",
            OutputWorkbookPath = null
        };

        var receipt = await new ExcelTaskEngine(runtime).RunAsync(request, CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Rejected, receipt.Status);
        var check = Assert.Single(receipt.Checks, check => check.Name == "request");
        Assert.Contains(".xlsm target", check.Detail, StringComparison.Ordinal);
        Assert.Contains("Isolated", check.Detail, StringComparison.Ordinal);
        Assert.Contains("save mode Copy", check.Detail, StringComparison.Ordinal);
        Assert.Contains("outputWorkbookPath", check.Detail, StringComparison.Ordinal);
    }

    [Fact]
    public async Task MacroApplyRejectsAutomaticEntryProcedureAndParameterizedRunBeforeInspection()
    {
        var autoRuntime = new FakeRuntime();
        var auto = await new ExcelTaskEngine(autoRuntime).RunAsync(
            MacroRequest(mode: ExcelTaskMode.Apply, operation: new ExcelOperation(
                ExcelOperationKind.EditMacroProcedure,
                EditMacroProcedure: new("Module1", "Auto_Open", new string('a', 64), "Sub Auto_Open()\nEnd Sub"))),
            CancellationToken.None);
        var parametersRuntime = new FakeRuntime();
        var parameters = await new ExcelTaskEngine(parametersRuntime).RunAsync(
            MacroRequest(mode: ExcelTaskMode.Apply, operation: Macro(
                expectedHash: new string('a', 64),
                replacementSource: "Sub RefreshModel(value As Long)\nEnd Sub",
                runAfterEdit: true)),
            CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Rejected, auto.Status);
        Assert.Null(autoRuntime.InspectionRequest);
        Assert.Equal(ExcelTaskStatus.Rejected, parameters.Status);
        Assert.Null(parametersRuntime.InspectionRequest);
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
    public async Task ReturnsUnknownWithoutEchoingRuntimeExceptionText()
    {
        const string privateVbaAndPath = "Sub SecretProcedure() C:\\private\\workbook.xlsm";
        var runtime = new FakeRuntime { ExecuteException = new InvalidOperationException(privateVbaAndPath) };
        var receipt = await new ExcelTaskEngine(runtime).RunAsync(Request(), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Unknown, receipt.Status);
        Assert.False(receipt.Retry.CanRetry);
        Assert.Contains("Reconcile", receipt.Retry.Reason);
        Assert.DoesNotContain(privateVbaAndPath, receipt.Checks.Single(check => check.Name == "runtime-execution").Detail, StringComparison.Ordinal);
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

    [Fact]
    public async Task MacroReceiptPassesPlanEvidenceAndOmitsOversizedRuntimeSource()
    {
        const string source = "Sub RefreshModel()\nEnd Sub";
        var runtime = new FakeRuntime
        {
            Outcome = new(
                ExcelTaskStatus.Planned,
                "Macro plan ready",
                MacroProcedure: new MacroProcedureReceipt("StandardModule", "RefreshModel", MacroProcedureText.ComputeSha256(source), source, false, false))
        };

        var receipt = await new ExcelTaskEngine(runtime).RunAsync(MacroRequest(), CancellationToken.None);

        Assert.NotNull(receipt.MacroProcedure);
        Assert.Equal(source, receipt.MacroProcedure.Source);
        Assert.Equal(MacroProcedureText.ComputeSha256(source), receipt.MacroProcedure.Sha256);

        var oversizedSource = $"Sub RefreshModel()\n'{new string('x', MacroProcedureText.MaxSourceCharacters)}\nEnd Sub";
        runtime = new FakeRuntime
        {
            Outcome = new(
                ExcelTaskStatus.Planned,
                "Macro plan ready",
                MacroProcedure: new MacroProcedureReceipt(new string('c', 600), "RefreshModel", new string('a', 600), oversizedSource, false, false))
        };

        receipt = await new ExcelTaskEngine(runtime).RunAsync(MacroRequest(), CancellationToken.None);

        Assert.NotNull(receipt.MacroProcedure);
        Assert.Equal(256, receipt.MacroProcedure.ComponentName.Length);
        Assert.Equal(256, receipt.MacroProcedure.Sha256.Length);
        Assert.Null(receipt.MacroProcedure.Source);
    }

    [Fact]
    public async Task AuditDispatchesWithNoSaveDestinationAndReturnsItsReport()
    {
        var runtime = new FakeRuntime
        {
            Outcome = new(
                ExcelTaskStatus.Completed,
                "Audited",
                Audit: new WorkbookAuditReceipt(
                    [
                        new WorkbookFlowItem("query", "Trial Balance", "loads to model"),
                        new WorkbookFlowItem("relationship", "Accounts -> TB", "many to one", "Accounts")
                    ],
                    2,
                    Truncated: false,
                    WorkbookUnchanged: true))
        };

        var receipt = await new ExcelTaskEngine(runtime).RunAsync(AuditRequest(), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Completed, receipt.Status);
        Assert.NotNull(runtime.Plan!.Request.Operation.AuditWorkbookFlows);
        Assert.Null(runtime.Plan.Request.OutputWorkbookPath);
        Assert.Equal(SaveMode.Same, runtime.Plan.Request.Save);
        Assert.Equal(2, receipt.Audit!.Items.Count);
        Assert.True(receipt.Audit.WorkbookUnchanged);
        Assert.Equal("Accounts", receipt.Audit.Items[1].DependsOn);
    }

    [Theory]
    [InlineData(SaveMode.Copy, ".\\out.xlsx")]
    [InlineData(SaveMode.Same, ".\\out.xlsx")]
    public async Task AuditRefusesAnySaveDestinationBecauseItNeverWrites(SaveMode save, string? output)
    {
        var runtime = new FakeRuntime();

        var receipt = await new ExcelTaskEngine(runtime).RunAsync(AuditRequest(save: save, output: output), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Rejected, receipt.Status);
        Assert.Null(runtime.Plan);
    }

    [Fact]
    public async Task AuditReportBoundsItsItemsWhileStillReportingTheRealTotal()
    {
        var many = Enumerable.Range(0, 60)
            .Select(index => new WorkbookFlowItem("query", $"Query {index}", new string('d', 400)))
            .ToArray();
        var runtime = new FakeRuntime
        {
            Outcome = new(ExcelTaskStatus.Completed, "Audited",
                Audit: new WorkbookAuditReceipt(many, many.Length, Truncated: false, WorkbookUnchanged: true))
        };

        var receipt = await new ExcelTaskEngine(runtime).RunAsync(AuditRequest(), CancellationToken.None);

        Assert.True(receipt.Audit!.Items.Count < many.Length);
        Assert.True(receipt.Audit.Truncated);
        // The count of what exists survives the cap, so a partial map cannot read as a whole one.
        Assert.Equal(60, receipt.Audit.TotalFound);
        Assert.All(receipt.Audit.Items, item => Assert.True(item.Detail.Length <= 256));
    }

    [Fact]
    public async Task ReadRejectsARangeLargerThanItWillReturnRatherThanTruncatingSilently()
    {
        var runtime = new FakeRuntime();

        // 21 x 21 is 441 cells, just past the cap. Accepting it and returning 400 would hand back a
        // partial answer that looks complete, which is worse than a rejection naming the limit.
        var receipt = await new ExcelTaskEngine(runtime).RunAsync(
            ReadRequest(range: "A1:U21"), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Rejected, receipt.Status);
        Assert.Null(receipt.Range);
        Assert.Contains("400", receipt.Summary, StringComparison.Ordinal);
    }

    [Fact]
    public async Task ReadRejectsASaveDestinationBecauseItNeverWrites()
    {
        var runtime = new FakeRuntime();

        var receipt = await new ExcelTaskEngine(runtime).RunAsync(
            ReadRequest(save: SaveMode.Copy, output: ".\\copy.xlsx"), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Rejected, receipt.Status);
    }

    [Fact]
    public async Task ReadDoesNotDemandOverwriteConfirmationForAWorkbookItOnlyReads()
    {
        var runtime = new FakeRuntime();

        var receipt = await new ExcelTaskEngine(runtime).RunAsync(ReadRequest(), CancellationToken.None);

        // Save mode Same normally requires overwrite confirmation on Apply. A read never writes, so
        // demanding it would be asking permission for something that cannot happen - and a caller
        // taught to set overwriteConfirmed to get a read through carries it into the call that does
        // write.
        Assert.False(receipt.Confirmation.Required);
        Assert.Empty(receipt.Confirmation.Requirements);
        Assert.Equal(ExcelTaskStatus.Completed, receipt.Status);
    }

    [Fact]
    public async Task ReadReceiptBoundsItsCellsAndTheirText()
    {
        var many = Enumerable.Range(1, 500)
            .Select(index => new WorksheetCell($"A{index}", new string('x', 400)))
            .ToArray();
        var runtime = new FakeRuntime
        {
            Outcome = new(ExcelTaskStatus.Completed, "Read",
                Range: new WorksheetRangeReceipt("Sheet1", "A1:A500", false, 500, 500, many, Truncated: false))
        };

        var receipt = await new ExcelTaskEngine(runtime).RunAsync(ReadRequest(), CancellationToken.None);

        Assert.True(receipt.Status == ExcelTaskStatus.Completed, receipt.Summary);
        Assert.True(receipt.Range!.Cells.Count < many.Length);
        Assert.True(receipt.Range.Truncated);
        // The counts describe the range that was asked for, so a bounded answer still says how much
        // of the sheet it covered rather than implying the rest was blank.
        Assert.Equal(500, receipt.Range.CellsInRange);
        Assert.All(receipt.Range.Cells, cell => Assert.True(cell.Text.Length <= 64));
    }

    [Fact]
    public async Task WriteRefusesFormulaTextAndSaysWhatToUseInstead()
    {
        var runtime = new FakeRuntime();

        var receipt = await new ExcelTaskEngine(runtime).RunAsync(
            WriteRequest([("B7", "=SUM(A1:A6)")]), CancellationToken.None);

        // The whole basis for allowing writes at all is that a constant cannot be plausibly and
        // silently wrong the way a model-composed formula can. Storing "=SUM(...)" as a text label
        // would be the worst of both: it looks written and computes nothing.
        Assert.Equal(ExcelTaskStatus.Rejected, receipt.Status);
        Assert.Contains("constants only", receipt.Summary, StringComparison.Ordinal);
        Assert.Contains("ExtendFormulaSeries", receipt.Summary, StringComparison.Ordinal);
    }

    [Fact]
    public async Task WriteRefusesTheSameCellTwiceBecauseTheRequestWouldNotSayWhatItWants()
    {
        var runtime = new FakeRuntime();

        var receipt = await new ExcelTaskEngine(runtime).RunAsync(
            WriteRequest([("B7", "1"), ("b7", "2")]), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Rejected, receipt.Status);
        Assert.Contains("more than once", receipt.Summary, StringComparison.Ordinal);
    }

    [Fact]
    public async Task WriteRefusesCellsScatteredAcrossTheSheet()
    {
        var runtime = new FakeRuntime();

        // Two cells, well inside the count limit, but spanning far more of the sheet than a write
        // is allowed to touch - so the count alone cannot be the only bound.
        var receipt = await new ExcelTaskEngine(runtime).RunAsync(
            WriteRequest([("A1", "1"), ("Z900", "2")]), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Rejected, receipt.Status);
        Assert.Contains("span", receipt.Summary, StringComparison.Ordinal);
    }

    [Fact]
    public async Task WriteRequiresOverwriteConfirmationBecauseItChangesTheWorkbook()
    {
        var runtime = new FakeRuntime();

        var receipt = await new ExcelTaskEngine(runtime).RunAsync(
            WriteRequest([("B7", "1")], overwriteConfirmed: false), CancellationToken.None);

        // The opposite of the read: this one does write, so the confirmation it asks for is real.
        Assert.Equal(ExcelTaskStatus.NeedsConfirmation, receipt.Status);
        Assert.Contains(receipt.Confirmation.Requirements, requirement => requirement.Code == "overwrite-same");
    }

    [Fact]
    public async Task EveryOperationKindHasACountedPayload()
    {
        // The union's arity check counts payloads by hand. An operation missing from that list is
        // not just miscounted, it is unreachable: the request fails arity before reaching its own
        // validation. FindReplace and Create both shipped that way, and both looked complete.
        ExcelOperation[] oneOfEach =
        [
            new(ExcelOperationKind.CopyExhibit, CopyExhibit: Copy()),
            new(ExcelOperationKind.RepairExistingWorksheet, RepairExistingWorksheet: new("Sheet1", ["A1:B2"])),
            new(ExcelOperationKind.ExtendFormulaSeries, ExtendFormulaSeries: new("Sheet1", FormulaExtensionDirection.Down, "A1:A2", "A3:A4")),
            new(ExcelOperationKind.EditMacroProcedure, EditMacroProcedure: new("StandardModule", "RefreshModel")),
            new(ExcelOperationKind.AuditWorkbookFlows, AuditWorkbookFlows: new()),
            new(ExcelOperationKind.ReadWorksheetRange, ReadWorksheetRange: new("Sheet1", "A1:B2")),
            new(ExcelOperationKind.WriteWorksheetValues, WriteWorksheetValues: new("Sheet1", [new("A1", "1")])),
            new(ExcelOperationKind.FindReplace, FindReplace: new("Sheet1", "FY25")),
            new(ExcelOperationKind.Create, Create: new(CreateKind.Workbook))
        ];

        // The list above is the checklist: a new operation that is not added here fails this line
        // before it can fail silently in production.
        Assert.Equal(Enum.GetValues<ExcelOperationKind>(), [.. oneOfEach.Select(operation => operation.Kind)]);

        foreach (var operation in oneOfEach)
        {
            var receipt = await new ExcelTaskEngine(new FakeRuntime()).RunAsync(Request(operation), CancellationToken.None);
            Assert.DoesNotContain("exactly one payload", receipt.Summary, StringComparison.Ordinal);
        }
    }

    [Fact]
    public async Task FindReplacePlanMustNotCarryTheReplacementText()
    {
        var receipt = await new ExcelTaskEngine(new FakeRuntime()).RunAsync(
            FindReplaceRequest(replaceWith: "FY26", mode: ExcelTaskMode.Plan), CancellationToken.None);

        // Plan lists matches; a replacement on it would read as authorization the caller never gave.
        Assert.Equal(ExcelTaskStatus.Rejected, receipt.Status);
        Assert.Contains("omit replaceWith", receipt.Summary, StringComparison.Ordinal);
    }

    [Fact]
    public async Task FindReplaceApplyRequiresTheReplacementText()
    {
        var receipt = await new ExcelTaskEngine(new FakeRuntime()).RunAsync(
            FindReplaceRequest(replaceWith: null), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Rejected, receipt.Status);
        Assert.Contains("Apply requires replaceWith", receipt.Summary, StringComparison.Ordinal);
    }

    [Fact]
    public async Task FindReplaceRefusesFormulaTextAsTheReplacement()
    {
        var receipt = await new ExcelTaskEngine(new FakeRuntime()).RunAsync(
            FindReplaceRequest(replaceWith: "=SUM(A1:A9)"), CancellationToken.None);

        // The same rule the constant write keeps: model-composed formula text is never accepted.
        Assert.Equal(ExcelTaskStatus.Rejected, receipt.Status);
        Assert.Contains("never writes formula text", receipt.Summary, StringComparison.Ordinal);
    }

    [Fact]
    public async Task FindReplaceRejectsARangeLargerThanTheSearchCap()
    {
        var receipt = await new ExcelTaskEngine(new FakeRuntime()).RunAsync(
            FindReplaceRequest(range: "A1:Z10000"), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Rejected, receipt.Status);
        Assert.Contains("10,000 cells", receipt.Summary, StringComparison.Ordinal);
    }

    [Fact]
    public async Task FindReplaceWithoutARangeReachesTheRuntimeSoTheUsedRangeCanBeMeasured()
    {
        var runtime = new FakeRuntime();

        var receipt = await new ExcelTaskEngine(runtime).RunAsync(FindReplaceRequest(range: null), CancellationToken.None);

        // The cap on an omitted range can only be applied against the real used range, so validation
        // must let it through rather than guessing.
        Assert.Equal(ExcelTaskStatus.Completed, receipt.Status);
        Assert.Null(runtime.Plan!.Request.Operation.FindReplace!.Range);
    }

    [Fact]
    public async Task FindReplaceApplyRequiresOverwriteConfirmationBecauseItChangesTheWorkbook()
    {
        var receipt = await new ExcelTaskEngine(new FakeRuntime()).RunAsync(
            FindReplaceRequest(overwriteConfirmed: false), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.NeedsConfirmation, receipt.Status);
        Assert.Contains(receipt.Confirmation.Requirements, requirement => requirement.Code == "overwrite-same");
    }

    [Fact]
    public async Task CreatingAWorkbookIsNotGivenASaveDestination()
    {
        var receipt = await new ExcelTaskEngine(new FakeRuntime()).RunAsync(
            CreateRequest(save: SaveMode.Copy, output: ".\\other.xlsx"), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Rejected, receipt.Status);
        Assert.Contains("must not be given a save destination", receipt.Summary, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(WorkbookBinding.UseOpen)]
    [InlineData(WorkbookBinding.AskIfOpen)]
    public async Task CreatingRequiresIsolatedBindingAndNothingElse(WorkbookBinding binding)
    {
        var receipt = await new ExcelTaskEngine(new FakeRuntime()).RunAsync(
            CreateRequest(binding: binding), CancellationToken.None);

        // AskIfOpen used to be accepted, and it led nowhere: on an open target its confirmation
        // offers UseOpen, which creation refuses, and Isolated, which inspection then rejects.
        Assert.Equal(ExcelTaskStatus.Rejected, receipt.Status);
        Assert.Contains("requires workbook binding Isolated", receipt.Summary, StringComparison.Ordinal);
    }

    [Fact]
    public async Task CreatingAWorksheetRequiresItsName()
    {
        var receipt = await new ExcelTaskEngine(new FakeRuntime()).RunAsync(
            CreateRequest(kind: CreateKind.Worksheet, worksheetName: null), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Rejected, receipt.Status);
        Assert.Contains("Worksheet name", receipt.Summary, StringComparison.Ordinal);
    }

    [Fact]
    public async Task CreatingAWorkbookTakesNoWorksheetName()
    {
        var receipt = await new ExcelTaskEngine(new FakeRuntime()).RunAsync(
            CreateRequest(worksheetName: "Summary"), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Rejected, receipt.Status);
        Assert.Contains("takes no worksheet name", receipt.Summary, StringComparison.Ordinal);
    }

    [Fact]
    public async Task CreatingAWorkbookNeverAsksToConfirmAnOverwriteItCannotPerform()
    {
        var runtime = new FakeRuntime();

        var receipt = await new ExcelTaskEngine(runtime).RunAsync(
            CreateRequest(overwriteConfirmed: false), CancellationToken.None);

        // Creation is refused outright if anything exists at the path, so an overwrite confirmation
        // would authorize nothing - and teach the caller to set the flag reflexively.
        Assert.Equal(ExcelTaskStatus.Completed, receipt.Status);
        Assert.False(receipt.Confirmation.Required);
        Assert.False(runtime.InspectionRequest!.TargetMustExist);
    }

    [Fact]
    public async Task CreatingAWorksheetStillConfirmsTheOverwriteBecauseItChangesAnExistingFile()
    {
        var runtime = new FakeRuntime();

        var receipt = await new ExcelTaskEngine(runtime).RunAsync(
            CreateRequest(kind: CreateKind.Worksheet, worksheetName: "Summary", overwriteConfirmed: false),
            CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.NeedsConfirmation, receipt.Status);
        Assert.Contains(receipt.Confirmation.Requirements, requirement => requirement.Code == "overwrite-same");
        Assert.True(runtime.InspectionRequest!.TargetMustExist);
    }

    private static ExcelTaskRequest FindReplaceRequest(
        string? replaceWith = "FY26",
        string? range = "A1:D50",
        ExcelTaskMode mode = ExcelTaskMode.Apply,
        bool overwriteConfirmed = true) => new(
            TargetWorkbookPath: ".\\model.xlsx",
            Operation: new ExcelOperation(
                ExcelOperationKind.FindReplace,
                FindReplace: new FindReplaceOperation("Sheet1", "FY25", replaceWith, range)),
            Mode: mode,
            WorkbookBinding: WorkbookBinding.Isolated,
            Save: SaveMode.Same,
            OutputWorkbookPath: null,
            OverwriteConfirmed: overwriteConfirmed);

    private static ExcelTaskRequest CreateRequest(
        CreateKind kind = CreateKind.Workbook,
        string? worksheetName = null,
        WorkbookBinding binding = WorkbookBinding.Isolated,
        SaveMode save = SaveMode.Same,
        string? output = null,
        bool overwriteConfirmed = false) => new(
            TargetWorkbookPath: ".\\new-model.xlsx",
            Operation: new ExcelOperation(ExcelOperationKind.Create, Create: new CreateOperation(kind, worksheetName)),
            Mode: ExcelTaskMode.Apply,
            WorkbookBinding: binding,
            Save: save,
            OutputWorkbookPath: output,
            OverwriteConfirmed: overwriteConfirmed);

    private static ExcelTaskRequest WriteRequest(
        (string Address, string Value)[] cells,
        bool overwriteConfirmed = true) => new(
            TargetWorkbookPath: ".\\model.xlsx",
            Operation: new ExcelOperation(
                ExcelOperationKind.WriteWorksheetValues,
                WriteWorksheetValues: new WriteWorksheetValuesOperation(
                    "Sheet1",
                    [.. cells.Select(cell => new WorksheetCellValue(cell.Address, cell.Value))])),
            Mode: ExcelTaskMode.Apply,
            WorkbookBinding: WorkbookBinding.Isolated,
            Save: SaveMode.Same,
            OutputWorkbookPath: null,
            OverwriteConfirmed: overwriteConfirmed);

    private static ExcelTaskRequest ReadRequest(
        string range = "A1:C3",
        SaveMode save = SaveMode.Same,
        string? output = null) => new(
            TargetWorkbookPath: ".\\model.xlsx",
            Operation: new ExcelOperation(
                ExcelOperationKind.ReadWorksheetRange,
                ReadWorksheetRange: new ReadWorksheetRangeOperation("Sheet1", range)),
            Mode: ExcelTaskMode.Apply,
            WorkbookBinding: WorkbookBinding.Isolated,
            Save: save,
            OutputWorkbookPath: output,
            OverwriteConfirmed: false);

    private static ExcelTaskRequest AuditRequest(
        ExcelTaskMode mode = ExcelTaskMode.Apply,
        SaveMode save = SaveMode.Same,
        string? output = null) => new(
            TargetWorkbookPath: ".\\model.xlsx",
            Operation: new ExcelOperation(ExcelOperationKind.AuditWorkbookFlows, AuditWorkbookFlows: new AuditWorkbookFlowsOperation()),
            Mode: mode,
            WorkbookBinding: WorkbookBinding.Isolated,
            Save: save,
            OutputWorkbookPath: output,
            OverwriteConfirmed: false);

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

    private static ExcelOperation Macro(
        string? expectedHash = null,
        string? replacementSource = null,
        bool runAfterEdit = false) => new(
            ExcelOperationKind.EditMacroProcedure,
            EditMacroProcedure: new EditMacroProcedureOperation("StandardModule", "RefreshModel", expectedHash, replacementSource, runAfterEdit));

    private static ExcelTaskRequest MacroRequest(
        ExcelOperation? operation = null,
        ExcelTaskMode mode = ExcelTaskMode.Plan,
        WorkbookBinding binding = WorkbookBinding.Isolated,
        SaveMode save = SaveMode.Copy) => new(
            TargetWorkbookPath: ".\\target.xlsm",
            Operation: operation ?? Macro(),
            Mode: mode,
            WorkbookBinding: binding,
            Save: save,
            OutputWorkbookPath: save == SaveMode.Copy ? ".\\out.xlsm" : null,
            OverwriteConfirmed: false);

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
