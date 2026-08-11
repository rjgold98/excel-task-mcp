using System.IO.Pipelines;
using System.Text.Json;
using ExcelTask.Core;
using ExcelTask.McpServer;
using Microsoft.Extensions.DependencyInjection;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using Server = ModelContextProtocol.Server;

namespace ExcelTask.McpServer.Tests;

public sealed class ExcelTaskToolProtocolTests : IAsyncLifetime, IAsyncDisposable
{
    private readonly Pipe _clientToServer = new();
    private readonly Pipe _serverToClient = new();
    private readonly CancellationTokenSource _cancellation = new();
    private ServiceProvider? _services;
    private Task? _serverTask;
    private McpClient? _client;
    private FakeRuntime? _runtime;

    public async Task InitializeAsync()
    {
        _runtime = new FakeRuntime();
        var services = new ServiceCollection();
        services.AddSingleton<IWorkbookRuntime>(_runtime);
        services.AddSingleton<IExcelTaskEngine, ExcelTaskEngine>();
        services
            .AddMcpServer(options => options.ServerInfo = new() { Name = "ExcelTask-Test", Version = "1.0.0" })
            .WithStreamServerTransport(_clientToServer.Reader.AsStream(), _serverToClient.Writer.AsStream())
            .WithTools<ExcelTaskTool>();

        _services = services.BuildServiceProvider(validateScopes: true);
        _serverTask = _services.GetRequiredService<Server.McpServer>().RunAsync(_cancellation.Token);
        _client = await McpClient.CreateAsync(
            new StreamClientTransport(_clientToServer.Writer.AsStream(), _serverToClient.Reader.AsStream()),
            new McpClientOptions { ClientInfo = new() { Name = "ExcelTask-TestClient", Version = "1.0.0" } },
            cancellationToken: _cancellation.Token);
    }

    public async Task DisposeAsync()
    {
        await _cancellation.CancelAsync();
        _clientToServer.Writer.Complete();
        _serverToClient.Writer.Complete();

        if (_client is not null) await _client.DisposeAsync();
        if (_serverTask is not null)
        {
            try
            {
                await _serverTask;
            }
            catch (OperationCanceledException)
            {
                // Expected while stopping the in-process transport.
            }
        }

        if (_services is not null) await _services.DisposeAsync();
        _cancellation.Dispose();
    }

    async ValueTask IAsyncDisposable.DisposeAsync()
    {
        await DisposeAsync();
        GC.SuppressFinalize(this);
    }

    [Fact]
    public async Task InitializeAndListToolsExposeExactlyOneBoundedExcelTaskSchema()
    {
        Assert.Equal("ExcelTask-Test", _client!.ServerInfo!.Name);

        var listed = await _client.ListToolsAsync(new ListToolsRequestParams(), _cancellation.Token);

        var tool = Assert.Single(listed.Tools);
        Assert.Equal("excel_task", tool.Name);
        // The budget forces every operation to earn its bytes: 8 KB for the first four, 9 KB when
        // the audit became the fifth, 11 KB for the range read, 13 KB for the value write, 15 KB
        // for find/replace and create together, 16 KB when the structure scan became the eleventh.
        // The two reads and writes of cells also carry an output schema, being the only operations
        // that return workbook contents. It must not grow to make room for wordier prose - only for
        // a new operation, or for a rule the caller cannot act without, which field measurement
        // showed cost two round trips when it was left out of the schema. The UX round proved the
        // bound works in the other direction too: three fixes landed only because it forced cuts.
        Assert.InRange(JsonSerializer.SerializeToUtf8Bytes(tool).Length, 1, 16 * 1024);

        var schema = tool.InputSchema.GetRawText();
        Assert.Contains("request", schema, StringComparison.Ordinal);

        // Field measurement showed a caller losing two round trips to macro rules it could not see:
        // it sent Apply-only fields on a Plan, and used the default binding where only Isolated is
        // permitted. Both rules must stay visible in the schema itself, not only in the rejection.
        Assert.Contains("Isolated", schema, StringComparison.Ordinal);
        Assert.Contains("omitted for Plan", schema, StringComparison.Ordinal);
        Assert.DoesNotContain("session", schema, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("workbookData", schema, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("confirmationToken", schema, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("formulaR1C1", schema, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("formulaText", schema, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("sourceText", schema, StringComparison.OrdinalIgnoreCase);

        // The guard against a model-selection input bans property NAMES, not prose: Excel's Data
        // Model now legitimately appears in the audit's description, and the thing this protects
        // against is a field, not a word.
        Assert.DoesNotContain(
            EnumeratePropertyNames(tool.InputSchema),
            name => name.Contains("model", StringComparison.OrdinalIgnoreCase));

        var request = ResolveReference(tool.InputSchema.GetProperty("properties").GetProperty("request"), tool.InputSchema);
        var properties = request.GetProperty("properties");
        Assert.Equal(
            ["targetWorkbookPath", "operation"],
            request.GetProperty("required").EnumerateArray().Select(item => item.GetString()));
        AssertDescription(properties, "targetWorkbookPath", "Target workbook path, ending .xlsx or .xlsm. It must already exist for every operation except Create with kind Workbook, where it must not.");
        AssertDescription(properties, "operation", "The required manual operation union. Supply exactly one payload matching kind.");
        AssertDescription(properties, "mode", "Plan previews without mutation; Apply performs the task after required confirmations.");
        // A/B measurement (docs/INTERFACE-AB-STUDY.md) showed both of these rules cost round trips
        // when they were enforced but unstated: UseOpen+Copy is rejected, and asking AskIfOpen about
        // a workbook the caller already said was open only ever returns a confirmation.
        AssertDescription(properties, "workbookBinding", "Use AskIfOpen when the workbook state is unknown; resubmit with UseOpen or Isolated if confirmation is returned. When the request already says the workbook is open, use UseOpen directly to avoid a wasted round trip. EditMacroProcedure and Create require Isolated. UseOpen cannot be combined with Copy, and Isolated with save Same is rejected while the target is open in Excel.");
        AssertDescription(properties, "save", "Same saves to the target; Copy saves only to outputWorkbookPath. EditMacroProcedure requires Copy to an .xlsm path. Copy is rejected with UseOpen. AuditWorkbookFlows, ReadWorksheetRange, ScanWorkbookStructure and Create never take a save destination: leave Same with no outputWorkbookPath.");
        AssertDescription(properties, "outputWorkbookPath", "Required destination path when save is Copy; omit for Same. Must differ from the target path and carry the target's extension.");
        // The single most frequent failure in every A/B round, in both the original study and the
        // v0.11.0 run: a same-file Apply sent without this flag. "An existing save destination" did
        // not read as "the target workbook you are editing", so the rule is now spelled per save mode.
        AssertDescription(properties, "overwriteConfirmed", "Explicit authorization to overwrite. Apply with save Same requires it, because saving in place overwrites the target workbook itself. Apply with save Copy requires it only when outputWorkbookPath already exists. Plan never requires it, and neither does Create with kind Workbook, which is refused outright if anything is already at the path.");

        var operation = ResolveReference(properties.GetProperty("operation"), tool.InputSchema);
        var operationProperties = operation.GetProperty("properties");
        AssertDescription(operationProperties, "kind", "Selects which one operation payload is supplied.");
        AssertDescription(operationProperties, "copyExhibit", "Required only when kind is CopyExhibit; all other payloads must be null.");
        AssertDescription(operationProperties, "repairExistingWorksheet", "Required only when kind is RepairExistingWorksheet; all other payloads must be null.");
        AssertDescription(operationProperties, "extendFormulaSeries", "Required only when kind is ExtendFormulaSeries; all other payloads must be null.");
        AssertDescription(operationProperties, "editMacroProcedure", "Required only when kind is EditMacroProcedure; all other payloads must be null.");
        AssertDescription(operationProperties, "auditWorkbookFlows", "Required only when kind is AuditWorkbookFlows; all other payloads must be null. It takes no options, so supply the empty object {} - omitting it entirely is rejected. The read-only report lists worksheets, tables, defined names, queries, connections, macro components and procedures, the data model, pivots, and external links.");
        AssertDescription(operationProperties, "readWorksheetRange", "Required only when kind is ReadWorksheetRange; all other payloads must be null. Reads one bounded range and returns its contents.");
        AssertDescription(operationProperties, "writeWorksheetValues", "Required only when kind is WriteWorksheetValues; all other payloads must be null. Writes constants into named cells and reads them back. Never accepts formula text.");
        AssertDescription(operationProperties, "findReplace", "Required only when kind is FindReplace; all other payloads must be null. Plan lists the matching cells and changes nothing - also how to locate a known label; Apply rewrites the constants among them.");
        AssertDescription(operationProperties, "create", "Required only when kind is Create; all other payloads must be null. Creates an empty workbook or adds an empty worksheet.");
        AssertDescription(operationProperties, "setNumberFormat", "Required only when kind is SetNumberFormat; all other payloads must be null. Sets how a range displays its numbers. Plan reports the range's current format and changes nothing. It changes no cell values, and it sets nothing else: no fonts, fills, borders, widths, or conditional formats.");

        // The name says number format and the description says what is absent, because a caller who
        // reads "format" and sends a font spends a round trip finding out - the exact cost the
        // interface study measures.
        var numberFormat = ResolveReference(operationProperties.GetProperty("setNumberFormat"), tool.InputSchema);
        AssertDescription(numberFormat.GetProperty("properties"), "range", "One bounded A1 range to format, at most 10,000 cells. The format applies to every cell in it, including blank ones.");
        AssertDescription(numberFormat.GetProperty("properties"), "numberFormat", "Excel number format code in US-English form, at most 255 characters, such as #,##0.00 or #,##0;(#,##0) or 0.0% or yyyy-mm-dd or General to clear formatting.");

        // Every rule the engine rejects on has to be readable here, which the A/B study measured at
        // p = 0.0012. For these two that means the search cap, the two text caps, the Plan/Apply
        // split, the formula refusal, and the binding and save policy creation enforces.
        var findReplace = ResolveReference(operationProperties.GetProperty("findReplace"), tool.InputSchema);
        AssertDescription(findReplace.GetProperty("properties"), "find", "Text to find, at most 200 characters. Matching is plain text on what the cell displays; * and ? are literal, not wildcards.");
        AssertDescription(findReplace.GetProperty("properties"), "replaceWith", "Apply only, and must be omitted for Plan: replacement text, at most 200 characters. Cells holding formulas are reported but never rewritten, and a replacement that would leave a cell starting with = is refused before anything is written.");
        AssertDescription(findReplace.GetProperty("properties"), "range", "One bounded A1 range to search, at most 10,000 cells; omit to search the worksheet's used range when it is within that bound.");

        AssertDescription(operationProperties, "scanWorkbookStructure", "Required only when kind is ScanWorkbookStructure; all other payloads must be null, and supply the empty object {}. Reads the file directly without starting Excel - fast on any size. Reports each sheet's dimension and formula/constant counts, and flags mostly-formula columns holding scattered constants: the shape of a manual override. Counts only, never contents. Use it before deciding what to read or fix.");

        AssertDescription(operationProperties, "scanWorkbookStructure", "Required only when kind is ScanWorkbookStructure; all other payloads must be null, and supply the empty object {}. Reads the file directly without starting Excel - fast on any size. Reports each sheet's dimension and formula/constant counts, and flags mostly-formula columns holding scattered constants: the shape of a manual override. Counts only, never contents. Use it before deciding what to read or fix.");

        var create = ResolveReference(operationProperties.GetProperty("create"), tool.InputSchema);
        AssertDescription(create.GetProperty("properties"), "kind", "Workbook creates an empty workbook at targetWorkbookPath, which must not already exist. Worksheet adds an empty sheet to the existing target. Either way Create writes the target itself: it requires binding Isolated and takes no save destination.");
        AssertDescription(create.GetProperty("properties"), "worksheetName", "Required for Worksheet, and must not already be in use. Optional for Workbook: names its one starting sheet, so a single call yields a workbook ready to write to. Omitted, the receipt reports the default name Excel chose.");

        // The one operation that returns workbook contents states its own cap, because a caller
        // that does not know the bound asks for a whole sheet and gets a truncated answer back.
        var readRange = ResolveReference(operationProperties.GetProperty("readWorksheetRange"), tool.InputSchema);
        AssertDescription(readRange.GetProperty("properties"), "range", "One bounded A1 range to read, at most 400 cells. Narrow the range and read again if more is needed.");
        AssertDescription(readRange.GetProperty("properties"), "formulas", "False returns displayed values; true returns R1C1 formulas instead, for comparing a formula pattern.");

        var copyExhibit = ResolveReference(operationProperties.GetProperty("copyExhibit"), tool.InputSchema);
        AssertDescription(copyExhibit.GetProperty("properties"), "repairRanges", "Bounded A1 ranges on the copied worksheet where blank formulas may be repaired; use [] when none are needed. At most 16 ranges and 10,000 cells per call; split a larger area across calls.");

        var repairExisting = ResolveReference(operationProperties.GetProperty("repairExistingWorksheet"), tool.InputSchema);
        AssertDescription(repairExisting.GetProperty("properties"), "ranges", "One or more bounded A1 ranges where blank formulas may be repaired. At most 16 ranges and 10,000 cells per call; split a larger area across calls.");

        var extendSeries = ResolveReference(operationProperties.GetProperty("extendFormulaSeries"), tool.InputSchema);
        AssertDescription(extendSeries.GetProperty("properties"), "evidenceRange", "Exactly two adjacent evidence columns for Right or rows for Down, expressed as one A1 range. Evidence and destination together must stay within 10,000 cells.");
        AssertDescription(extendSeries.GetProperty("properties"), "destinationRange", "Immediately adjacent blank destination columns for Right or rows for Down, expressed as one A1 range. At most 24 periods and 2,000 destination cells, which binds before the 10,000-cell aggregate; split larger work across calls.");

        var editMacro = ResolveReference(operationProperties.GetProperty("editMacroProcedure"), tool.InputSchema);
        var editMacroProperties = editMacro.GetProperty("properties");
        // The routing sentence exists because the first real field use failed without it: the model
        // was asked to fix a macro, had no way to learn the names this operation demands, and fell
        // back to a different server entirely.
        AssertDescription(editMacroProperties, "componentName", "Existing VBA component name containing the procedure to inspect or replace. If unknown, run AuditWorkbookFlows first; it lists every macro component and procedure.");
        // The A/B study's follow-on audit found eight rules the engine enforced but the schema never
        // stated; the fix class is proven (p = 0.0012, and end to end at 32% fewer calls). These
        // pins hold the remaining ones in the schema: source bounds, blocked constructs for a run,
        // auto-entry procedures, and the path rules asserted at the request level above.
        AssertDescription(editMacroProperties, "procedureName", "Existing VBA procedure name to inspect or replace. Automatic-entry procedures such as Auto_Open cannot be edited.");
        AssertDescription(editMacroProperties, "expectedProcedureSha256", "Apply only, and must be omitted for Plan: SHA-256 fingerprint of the existing procedure, taken from the Plan receipt.");
        AssertDescription(editMacroProperties, "replacementSource", "Apply only, and must be omitted for Plan: one complete replacement Sub or Function procedure with the requested name, at most 8,192 characters and 200 lines.");
        AssertDescription(editMacroProperties, "runAfterEdit", "Apply only, and must be omitted for Plan. When true, Apply runs the replacement after the edit; the replacement must have zero parameters and must not contain MsgBox, InputBox, GetOpenFilename, GetSaveAsFilename, FileDialog, or Stop, which wait for a person.");
    }

    [Fact]
    public async Task CallToolPlanReturnsSuccessfulStructuredReceipt()
    {
        _runtime!.Outcome = new WorkbookExecutionOutcome(ExcelTaskStatus.Planned, "Plan ready");

        var result = await CallAsync(Request(ExcelTaskMode.Plan));
        var tool = Assert.Single((await _client!.ListToolsAsync(new ListToolsRequestParams(), _cancellation.Token)).Tools);

        Assert.False(result.IsError);
        Assert.Equal(nameof(ExcelTaskStatus.Planned), result.StructuredContent!.Value.GetProperty("status").GetString());
        Assert.Equal(nameof(SaveMode.Same), result.StructuredContent.Value.GetProperty("save").GetProperty("mode").GetString());
        Assert.Contains("Planned: Plan ready", result.Content.OfType<TextContentBlock>().Single().Text, StringComparison.Ordinal);
        AssertConformsToSchema(tool.OutputSchema!.Value, result.StructuredContent.Value, tool.OutputSchema!.Value, "$");
    }

    [Fact]
    public async Task CallToolOpenWorkbookReturnsNeedsConfirmationWithoutProtocolError()
    {
        _runtime!.Inspection = new WorkbookInspection(true, OpenWorkbookDescription: "Target workbook is open.");

        var result = await CallAsync(Request(ExcelTaskMode.Apply, overwriteConfirmed: true));

        Assert.False(result.IsError);
        Assert.Equal(nameof(ExcelTaskStatus.NeedsConfirmation), result.StructuredContent!.Value.GetProperty("status").GetString());
        Assert.True(result.StructuredContent.Value.GetProperty("confirmation").GetProperty("required").GetBoolean());
        Assert.Null(_runtime.Plan);
    }

    [Fact]
    public async Task CallToolRejectedRequestSetsProtocolErrorFlag()
    {
        var result = await CallAsync(Request(ExcelTaskMode.Plan) with { TargetWorkbookPath = string.Empty });

        Assert.True(result.IsError);
        Assert.Equal(nameof(ExcelTaskStatus.Rejected), result.StructuredContent!.Value.GetProperty("status").GetString());
        Assert.Null(_runtime!.InspectionRequest);
    }

    [Fact]
    public async Task CallToolPlanRepairExistingWorksheetRoundTripsOperation()
    {
        _runtime!.Outcome = new WorkbookExecutionOutcome(ExcelTaskStatus.Planned, "Repair plan ready");
        var operation = new ExcelOperation(
            ExcelOperationKind.RepairExistingWorksheet,
            RepairExistingWorksheet: new RepairExistingWorksheetOperation("Model", ["B2:C3"]));

        var result = await CallAsync(Request(ExcelTaskMode.Plan, operation));

        Assert.False(result.IsError);
        var repaired = Assert.IsType<NormalizedRepairExistingWorksheetOperation>(_runtime.Plan!.Request.Operation.RepairExistingWorksheet);
        Assert.Equal("Model", repaired.WorksheetName);
        Assert.Equal("B2:C3", repaired.Ranges.Single().ToString());
        Assert.Null(_runtime.InspectionRequest!.ReferenceWorkbookPath);
    }

    [Fact]
    public async Task CallToolPlanExtendFormulaSeriesRoundTripsOperation()
    {
        _runtime!.Outcome = new WorkbookExecutionOutcome(ExcelTaskStatus.Planned, "Extension plan ready");
        var operation = new ExcelOperation(
            ExcelOperationKind.ExtendFormulaSeries,
            ExtendFormulaSeries: new ExtendFormulaSeriesOperation(
                "Model", FormulaExtensionDirection.Right, "B2:C4", "D2:F4"));

        var result = await CallAsync(Request(ExcelTaskMode.Plan, operation));

        Assert.False(result.IsError);
        var extension = Assert.IsType<NormalizedExtendFormulaSeriesOperation>(_runtime.Plan!.Request.Operation.ExtendFormulaSeries);
        Assert.Equal(FormulaExtensionDirection.Right, extension.Direction);
        Assert.Equal("B2:C4", extension.EvidenceRange.ToString());
        Assert.Equal("D2:F4", extension.DestinationRange.ToString());
    }

    [Fact]
    public async Task CallToolPlanMacroProcedureRoundTripsOperationAndReturnsRequestedSource()
    {
        const string source = "Public Sub RefreshModel()\nEnd Sub";
        _runtime!.Outcome = new WorkbookExecutionOutcome(
            ExcelTaskStatus.Planned,
            "Macro plan ready",
            MacroProcedure: new("ModelCode", "RefreshModel", "a".PadLeft(64, 'a'), source, false, false));
        var operation = new ExcelOperation(
            ExcelOperationKind.EditMacroProcedure,
            EditMacroProcedure: new EditMacroProcedureOperation("ModelCode", "RefreshModel"));

        var result = await CallAsync(MacroRequest(ExcelTaskMode.Plan, operation));

        Assert.False(result.IsError);
        var macro = Assert.IsType<NormalizedEditMacroProcedureOperation>(_runtime.Plan!.Request.Operation.EditMacroProcedure);
        Assert.Equal("ModelCode", macro.ComponentName);
        Assert.Equal("RefreshModel", macro.ProcedureName);
        Assert.Equal(source, result.StructuredContent!.Value.GetProperty("macroProcedure").GetProperty("source").GetString());
        Assert.Equal("a".PadLeft(64, 'a'), result.StructuredContent.Value.GetProperty("macroProcedure").GetProperty("sha256").GetString());
    }

    [Fact]
    public async Task CallToolRejectsMismatchedOperationUnionBeforeInspection()
    {
        var mismatch = new ExcelOperation(
            ExcelOperationKind.CopyExhibit,
            RepairExistingWorksheet: new RepairExistingWorksheetOperation("Model", ["B2:C3"]));

        var result = await CallAsync(Request(ExcelTaskMode.Plan, mismatch));

        Assert.True(result.IsError);
        Assert.Equal(nameof(ExcelTaskStatus.Rejected), result.StructuredContent!.Value.GetProperty("status").GetString());
        Assert.Null(_runtime!.InspectionRequest);
    }

    [Fact]
    public async Task CallToolRejectsMismatchedMacroOperationUnionBeforeInspection()
    {
        var mismatch = new ExcelOperation(
            ExcelOperationKind.EditMacroProcedure,
            CopyExhibit: new CopyExhibitOperation("reference.xlsx", "Reference", "New sheet", []));

        var result = await CallAsync(Request(ExcelTaskMode.Plan, mismatch));

        Assert.True(result.IsError);
        Assert.Equal(nameof(ExcelTaskStatus.Rejected), result.StructuredContent!.Value.GetProperty("status").GetString());
        Assert.Null(_runtime!.InspectionRequest);
    }

    [Fact]
    public async Task CallToolBoundsMacroReceiptAndSuppressesSourceForApply()
    {
        var longMetadata = new string('m', 160);
        var boundedSource = CompleteProcedure(longMetadata, MacroProcedureText.MaxSourceCharacters);
        _runtime!.Outcome = new WorkbookExecutionOutcome(
            ExcelTaskStatus.Planned,
            "Macro planned",
            MacroProcedure: new(longMetadata, longMetadata, longMetadata, boundedSource, true, true));
        var planOperation = new ExcelOperation(
            ExcelOperationKind.EditMacroProcedure,
            EditMacroProcedure: new EditMacroProcedureOperation("ModelCode", "RefreshModel"));
        var applyOperation = new ExcelOperation(
            ExcelOperationKind.EditMacroProcedure,
            EditMacroProcedure: new EditMacroProcedureOperation(
                "ModelCode", "RefreshModel", new string('a', 64), "Public Sub RefreshModel()\nEnd Sub"));

        var planResult = await CallAsync(MacroRequest(ExcelTaskMode.Plan, planOperation));
        var plannedMacro = planResult.StructuredContent!.Value.GetProperty("macroProcedure");

        // 160-character metadata survives whole: the model-facing cap is ReceiptBounds's 256, not
        // the 96 this seam used to impose. A 160-character component name is legal VBA, and the old
        // cap cut it - along with any engine rejection summary long enough to name every unmet
        // requirement, which was the cap's real cost.
        Assert.Equal(160, plannedMacro.GetProperty("componentName").GetString()!.Length);
        Assert.Equal(160, plannedMacro.GetProperty("procedureName").GetString()!.Length);
        Assert.Equal(160, plannedMacro.GetProperty("sha256").GetString()!.Length);
        Assert.Equal(MacroProcedureText.MaxSourceCharacters, plannedMacro.GetProperty("source").GetString()!.Length);
        Assert.InRange(JsonSerializer.SerializeToUtf8Bytes(planResult).Length, 1, 30 * 1024);
        Assert.InRange(JsonSerializer.SerializeToUtf8Bytes(new { jsonrpc = "2.0", id = 1, result = planResult }).Length, 1, 32 * 1024);

        _runtime.Outcome = new WorkbookExecutionOutcome(
            ExcelTaskStatus.Completed,
            "Macro applied",
            MacroProcedure: new(longMetadata, longMetadata, longMetadata, boundedSource, true, true));
        var result = await CallAsync(MacroRequest(ExcelTaskMode.Apply, applyOperation, overwriteConfirmed: true));
        var macro = result.StructuredContent!.Value.GetProperty("macroProcedure");

        Assert.Equal(160, macro.GetProperty("componentName").GetString()!.Length);
        Assert.Equal(160, macro.GetProperty("procedureName").GetString()!.Length);
        Assert.Equal(160, macro.GetProperty("sha256").GetString()!.Length);
        Assert.Equal(JsonValueKind.Null, macro.GetProperty("source").ValueKind);
        Assert.True(macro.GetProperty("runRequested").GetBoolean());
        Assert.True(macro.GetProperty("runCompleted").GetBoolean());
    }

    [Fact]
    public async Task CallToolLimitsFinalJsonRpcEnvelopeTo32KiB()
    {
        const int maxEnvelopeBytes = 32 * 1024;
        var longText = new string('x', 4_096);
        _runtime!.Outcome = new WorkbookExecutionOutcome(
            ExcelTaskStatus.Completed,
            longText,
            Enumerable.Range(0, 64).Select(index => new TaskChange(longText, longText, longText)).ToArray(),
            Enumerable.Range(0, 64).Select(index => new TaskCheck(longText, true, longText)).ToArray(),
            true,
            longText);

        var result = await CallAsync(Request(ExcelTaskMode.Plan));
        var envelope = JsonSerializer.SerializeToUtf8Bytes(new { jsonrpc = "2.0", id = 1, result });

        Assert.InRange(envelope.Length, 1, maxEnvelopeBytes);
        Assert.Equal(20, result.StructuredContent!.Value.GetProperty("changes").GetArrayLength());
        Assert.Equal(20, result.StructuredContent.Value.GetProperty("checks").GetArrayLength());
    }

    [Fact]
    public async Task CallToolPreservesTerminalReopenVerificationCheck()
    {
        _runtime!.Outcome = new WorkbookExecutionOutcome(
            ExcelTaskStatus.Completed,
            "Completed",
            Checks:
            [
                new TaskCheck("target-path", true, "Target is readable."),
                new TaskCheck("reference-worksheet", true, "Reference exists."),
                new TaskCheck("destination-worksheet", true, "Destination is available."),
                new TaskCheck("formula-plan", true, "Formula plan is valid."),
                new TaskCheck("formula-change-count", true, "Formula changes were applied."),
                new TaskCheck("save", true, "Workbook was saved."),
                new TaskCheck("reopen-verification", true, "Saved workbook was reopened and verified.")
            ]);

        var result = await CallAsync(Request(ExcelTaskMode.Apply, overwriteConfirmed: true));
        var checks = result.StructuredContent!.Value.GetProperty("checks");

        Assert.Equal(7, checks.GetArrayLength());
        Assert.Equal("reopen-verification", checks[6].GetProperty("name").GetString());
        Assert.True(checks[6].GetProperty("passed").GetBoolean());
    }

    private async Task<CallToolResult> CallAsync(ExcelTaskRequest request) => await _client!.CallToolAsync(
        "excel_task",
        new Dictionary<string, object?> { ["request"] = request },
        cancellationToken: _cancellation.Token);

    private static ExcelTaskRequest Request(
        ExcelTaskMode mode,
        ExcelOperation? operation = null,
        bool overwriteConfirmed = false) => new(
        "target.xlsx",
        operation ?? new ExcelOperation(
            ExcelOperationKind.CopyExhibit,
            CopyExhibit: new CopyExhibitOperation("reference.xlsx", "Reference", "New sheet", ["A1:C3"])),
        mode,
        WorkbookBinding.AskIfOpen,
        SaveMode.Same,
        OverwriteConfirmed: overwriteConfirmed);

    private static ExcelTaskRequest MacroRequest(
        ExcelTaskMode mode,
        ExcelOperation operation,
        bool overwriteConfirmed = false) => new(
            TargetWorkbookPath: "target.xlsm",
            Operation: operation,
            Mode: mode,
            WorkbookBinding: WorkbookBinding.Isolated,
            Save: SaveMode.Copy,
            OutputWorkbookPath: "output.xlsm",
            OverwriteConfirmed: overwriteConfirmed);

    private static string CompleteProcedure(string procedureName, int totalLength)
    {
        var prefix = $"Public Sub {procedureName}()\n'";
        const string suffix = "\nEnd Sub";
        return prefix + new string('s', totalLength - prefix.Length - suffix.Length) + suffix;
    }

    private static void AssertDescription(JsonElement properties, string propertyName, string expected) =>
        Assert.Equal(expected, properties.GetProperty(propertyName).GetProperty("description").GetString());

    private static void AssertConformsToSchema(JsonElement schema, JsonElement value, JsonElement root, string path)
    {
        schema = ResolveReference(schema, root);

        if (schema.TryGetProperty("enum", out var enumValues))
        {
            Assert.Contains(enumValues.EnumerateArray(), allowed => allowed.GetRawText() == value.GetRawText());
        }

        if (schema.TryGetProperty("type", out var type))
        {
            Assert.True(TypeAllows(type, value.ValueKind), $"{path} has JSON type {value.ValueKind}, which is not allowed by {type.GetRawText()}.");
        }

        if (value.ValueKind == JsonValueKind.Object && schema.TryGetProperty("properties", out var properties))
        {
            if (schema.TryGetProperty("required", out var required))
            {
                foreach (var property in required.EnumerateArray())
                {
                    Assert.True(value.TryGetProperty(property.GetString()!, out _), $"{path} is missing required property {property.GetString()}.");
                }
            }

            foreach (var property in value.EnumerateObject())
            {
                if (properties.TryGetProperty(property.Name, out var propertySchema))
                {
                    AssertConformsToSchema(propertySchema, property.Value, root, $"{path}.{property.Name}");
                }
            }
        }

        if (value.ValueKind == JsonValueKind.Array && schema.TryGetProperty("items", out var itemSchema))
        {
            var index = 0;
            foreach (var item in value.EnumerateArray())
            {
                AssertConformsToSchema(itemSchema, item, root, $"{path}[{index++}]");
            }
        }
    }

    /// <summary>Every property name a schema defines, at any depth, through objects and arrays.</summary>
    private static IEnumerable<string> EnumeratePropertyNames(JsonElement element)
    {
        if (element.ValueKind == JsonValueKind.Object)
        {
            foreach (var property in element.EnumerateObject())
            {
                if (property.Name == "properties" && property.Value.ValueKind == JsonValueKind.Object)
                {
                    foreach (var defined in property.Value.EnumerateObject())
                    {
                        yield return defined.Name;
                        foreach (var nested in EnumeratePropertyNames(defined.Value)) yield return nested;
                    }
                }
                else
                {
                    foreach (var nested in EnumeratePropertyNames(property.Value)) yield return nested;
                }
            }
        }
        else if (element.ValueKind == JsonValueKind.Array)
        {
            foreach (var item in element.EnumerateArray())
            {
                foreach (var nested in EnumeratePropertyNames(item)) yield return nested;
            }
        }
    }

    private static JsonElement ResolveReference(JsonElement schema, JsonElement root)
    {
        while (schema.TryGetProperty("$ref", out var reference))
        {
            const string definitionsPrefix = "#/$defs/";
            var value = reference.GetString();
            Assert.StartsWith(definitionsPrefix, value, StringComparison.Ordinal);
            schema = root.GetProperty("$defs").GetProperty(value![definitionsPrefix.Length..]);
        }

        return schema;
    }

    private static bool TypeAllows(JsonElement type, JsonValueKind valueKind)
    {
        if (type.ValueKind == JsonValueKind.String) return TypeAllows(type.GetString()!, valueKind);
        return type.EnumerateArray().Any(candidate => TypeAllows(candidate.GetString()!, valueKind));
    }

    private static bool TypeAllows(string schemaType, JsonValueKind valueKind) => (schemaType, valueKind) switch
    {
        ("object", JsonValueKind.Object) => true,
        ("array", JsonValueKind.Array) => true,
        ("string", JsonValueKind.String) => true,
        ("boolean", JsonValueKind.True or JsonValueKind.False) => true,
        ("integer", JsonValueKind.Number) => true,
        ("number", JsonValueKind.Number) => true,
        ("null", JsonValueKind.Null) => true,
        _ => false
    };

    private sealed class FakeRuntime : IWorkbookRuntime
    {
        public WorkbookInspection Inspection { get; set; } = new(false);
        public WorkbookExecutionOutcome Outcome { get; set; } = new(ExcelTaskStatus.Completed, "Completed");
        public WorkbookInspectionRequest? InspectionRequest { get; private set; }
        public ExcelTaskPlan? Plan { get; private set; }

        public Task<WorkbookInspection> InspectAsync(WorkbookInspectionRequest request, CancellationToken cancellationToken)
        {
            InspectionRequest = request;
            return Task.FromResult(Inspection);
        }

        public Task<WorkbookExecutionOutcome> ExecuteAsync(ExcelTaskPlan plan, CancellationToken cancellationToken)
        {
            Plan = plan;
            return Task.FromResult(Outcome);
        }
    }
}
