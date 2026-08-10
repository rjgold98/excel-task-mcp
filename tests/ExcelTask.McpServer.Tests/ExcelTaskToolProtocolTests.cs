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
        Assert.InRange(JsonSerializer.SerializeToUtf8Bytes(tool).Length, 1, 8 * 1024);

        var schema = tool.InputSchema.GetRawText();
        Assert.Contains("request", schema, StringComparison.Ordinal);

        // Field measurement showed a caller losing two round trips to macro rules it could not see:
        // it sent Apply-only fields on a Plan, and used the default binding where only Isolated is
        // permitted. Both rules must stay visible in the schema itself, not only in the rejection.
        Assert.Contains("Isolated", schema, StringComparison.Ordinal);
        Assert.Contains("omitted for Plan", schema, StringComparison.Ordinal);
        Assert.DoesNotContain("session", schema, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("workbookData", schema, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("model", schema, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("confirmationToken", schema, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("formulaR1C1", schema, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("formulaText", schema, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("sourceText", schema, StringComparison.OrdinalIgnoreCase);

        var request = ResolveReference(tool.InputSchema.GetProperty("properties").GetProperty("request"), tool.InputSchema);
        var properties = request.GetProperty("properties");
        Assert.Equal(
            ["targetWorkbookPath", "operation"],
            request.GetProperty("required").EnumerateArray().Select(item => item.GetString()));
        AssertDescription(properties, "targetWorkbookPath", "Existing target workbook path.");
        AssertDescription(properties, "operation", "The required manual operation union. Supply exactly one payload matching kind.");
        AssertDescription(properties, "mode", "Plan previews without mutation; Apply performs the task after required confirmations.");
        AssertDescription(properties, "workbookBinding", "Use AskIfOpen first; if confirmation is returned, resubmit with UseOpen or Isolated. EditMacroProcedure requires Isolated and rejects anything else.");
        AssertDescription(properties, "save", "Same saves to the target; Copy saves only to outputWorkbookPath. EditMacroProcedure requires Copy to an .xlsm path.");
        AssertDescription(properties, "outputWorkbookPath", "Required destination path when save is Copy; omit for Same.");
        AssertDescription(properties, "overwriteConfirmed", "Explicit authorization required before Apply can overwrite an existing save destination.");

        var operation = ResolveReference(properties.GetProperty("operation"), tool.InputSchema);
        var operationProperties = operation.GetProperty("properties");
        AssertDescription(operationProperties, "kind", "Selects which one operation payload is supplied.");
        AssertDescription(operationProperties, "copyExhibit", "Required only when kind is CopyExhibit; all other payloads must be null.");
        AssertDescription(operationProperties, "repairExistingWorksheet", "Required only when kind is RepairExistingWorksheet; all other payloads must be null.");
        AssertDescription(operationProperties, "extendFormulaSeries", "Required only when kind is ExtendFormulaSeries; all other payloads must be null.");
        AssertDescription(operationProperties, "editMacroProcedure", "Required only when kind is EditMacroProcedure; all other payloads must be null.");

        var copyExhibit = ResolveReference(operationProperties.GetProperty("copyExhibit"), tool.InputSchema);
        AssertDescription(copyExhibit.GetProperty("properties"), "repairRanges", "Bounded A1 ranges on the copied worksheet where blank formulas may be repaired; use [] when none are needed.");

        var repairExisting = ResolveReference(operationProperties.GetProperty("repairExistingWorksheet"), tool.InputSchema);
        AssertDescription(repairExisting.GetProperty("properties"), "ranges", "One or more bounded A1 ranges where blank formulas may be repaired.");

        var extendSeries = ResolveReference(operationProperties.GetProperty("extendFormulaSeries"), tool.InputSchema);
        AssertDescription(extendSeries.GetProperty("properties"), "evidenceRange", "Exactly two adjacent evidence columns for Right or rows for Down, expressed as one A1 range.");
        AssertDescription(extendSeries.GetProperty("properties"), "destinationRange", "Immediately adjacent blank destination columns for Right or rows for Down, expressed as one A1 range.");

        var editMacro = ResolveReference(operationProperties.GetProperty("editMacroProcedure"), tool.InputSchema);
        var editMacroProperties = editMacro.GetProperty("properties");
        AssertDescription(editMacroProperties, "componentName", "Existing VBA component name containing the procedure to inspect or replace.");
        AssertDescription(editMacroProperties, "procedureName", "Existing VBA procedure name to inspect or replace.");
        AssertDescription(editMacroProperties, "expectedProcedureSha256", "Apply only, and must be omitted for Plan: SHA-256 fingerprint of the existing procedure, taken from the Plan receipt.");
        AssertDescription(editMacroProperties, "replacementSource", "Apply only, and must be omitted for Plan: one complete replacement Sub or Function procedure with the requested name.");
        AssertDescription(editMacroProperties, "runAfterEdit", "Apply only, and must be omitted for Plan. When true, Apply runs the replacement procedure after the edit; the replacement must have zero parameters.");
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

        Assert.Equal(96, plannedMacro.GetProperty("componentName").GetString()!.Length);
        Assert.Equal(96, plannedMacro.GetProperty("procedureName").GetString()!.Length);
        Assert.Equal(96, plannedMacro.GetProperty("sha256").GetString()!.Length);
        Assert.Equal(MacroProcedureText.MaxSourceCharacters, plannedMacro.GetProperty("source").GetString()!.Length);
        Assert.InRange(JsonSerializer.SerializeToUtf8Bytes(planResult).Length, 1, 30 * 1024);
        Assert.InRange(JsonSerializer.SerializeToUtf8Bytes(new { jsonrpc = "2.0", id = 1, result = planResult }).Length, 1, 32 * 1024);

        _runtime.Outcome = new WorkbookExecutionOutcome(
            ExcelTaskStatus.Completed,
            "Macro applied",
            MacroProcedure: new(longMetadata, longMetadata, longMetadata, boundedSource, true, true));
        var result = await CallAsync(MacroRequest(ExcelTaskMode.Apply, applyOperation, overwriteConfirmed: true));
        var macro = result.StructuredContent!.Value.GetProperty("macroProcedure");

        Assert.Equal(96, macro.GetProperty("componentName").GetString()!.Length);
        Assert.Equal(96, macro.GetProperty("procedureName").GetString()!.Length);
        Assert.Equal(96, macro.GetProperty("sha256").GetString()!.Length);
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
