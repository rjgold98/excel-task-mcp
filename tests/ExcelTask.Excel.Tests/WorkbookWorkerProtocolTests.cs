using System.Text;
using System.Text.Json;
using ExcelTask.Core;

namespace ExcelTask.Excel.Tests;

public sealed class WorkbookWorkerProtocolTests
{
    [Fact]
    public async Task HostEmitsAcceptedObservedLifecycleAndResultForOneInspection()
    {
        const string taskId = "worker_test_1";
        var request = JsonSerializer.Serialize(new
        {
            version = WorkbookWorkerProtocol.Version,
            taskId,
            operation = "inspect",
            inspection = new WorkbookInspectionRequest("target.xlsx", "reference.xlsx", WorkbookBinding.Isolated, SaveMode.Same, null)
        }, WorkbookWorkerProtocol.JsonOptions);
        using var input = new StringReader(request);
        using var output = new StringWriter();
        ObservingFakeRuntime? runtime = null;

        var exitCode = await WorkbookWorkerHost.RunAsync(input, output, observer => runtime = new ObservingFakeRuntime(observer));

        Assert.Equal(0, exitCode);
        Assert.True(runtime!.Disposed);
        var frames = ReadFrames(output);
        Assert.Collection(frames,
            frame => Assert.Equal("accepted", frame.GetProperty("type").GetString()),
            frame => Assert.Equal("phase", frame.GetProperty("type").GetString()),
            frame => Assert.Equal("phase", frame.GetProperty("type").GetString()),
            frame => Assert.Equal("owned-process", frame.GetProperty("type").GetString()),
            frame => Assert.Equal("artifact-staged", frame.GetProperty("type").GetString()),
            frame => Assert.Equal("result", frame.GetProperty("type").GetString()));
        Assert.All(frames, frame => Assert.Equal(taskId, frame.GetProperty("taskId").GetString()));
        Assert.True(frames[3].GetProperty("processId").GetInt32() > 0);
        Assert.Equal("C:\\worker-staging.xlsx", frames[4].GetProperty("stagingPath").GetString());
        Assert.True(frames[5].GetProperty("result").GetProperty("targetIsOpen").GetBoolean());
    }

    [Theory]
    [InlineData("{not-json")]
    [InlineData("{\"version\":1,\"taskId\":\"task\",\"operation\":\"inspect\",\"inspection\":{},\"plan\":{}}")]
    [InlineData("{\"version\":1e100,\"taskId\":\"task\",\"operation\":\"inspect\",\"inspection\":{\"targetWorkbookPath\":\"target.xlsx\",\"referenceWorkbookPath\":\"reference.xlsx\",\"binding\":\"isolated\",\"save\":\"same\"}}")]
    public async Task HostRejectsMalformedOrAmbiguousInput(string request)
    {
        using var input = new StringReader(request);
        using var output = new StringWriter();

        var exitCode = await WorkbookWorkerHost.RunAsync(input, output, _ => throw new InvalidOperationException("must not construct runtime"));

        Assert.Equal(2, exitCode);
        var frame = Assert.Single(ReadFrames(output));
        Assert.Equal("fatal", frame.GetProperty("type").GetString());
        Assert.Equal("invalid-request", frame.GetProperty("code").GetString());
    }

    [Fact]
    public async Task HostRejectsOversizeInputWithoutConstructingRuntime()
    {
        using var input = new StringReader(new string('x', WorkbookWorkerProtocol.MaxRequestBytes + 1));
        using var output = new StringWriter();

        var exitCode = await WorkbookWorkerHost.RunAsync(input, output, _ => throw new InvalidOperationException("must not construct runtime"));

        Assert.Equal(2, exitCode);
        var frame = Assert.Single(ReadFrames(output));
        Assert.Equal("input-too-large", frame.GetProperty("code").GetString());
    }

    [Fact]
    public async Task HostBoundsResultFramesAndNeverWritesThrownErrorText()
    {
        const string secret = "formula=TOP_SECRET()";
        var request = JsonSerializer.Serialize(new
        {
            version = WorkbookWorkerProtocol.Version,
            taskId = "worker_test_2",
            operation = "execute",
            plan = Plan("worker_test_2")
        }, WorkbookWorkerProtocol.JsonOptions);
        using var input = new StringReader(request);
        using var output = new StringWriter();

        var exitCode = await WorkbookWorkerHost.RunAsync(input, output, _ => new ThrowingFakeRuntime(secret));

        Assert.Equal(1, exitCode);
        var text = output.ToString();
        Assert.DoesNotContain(secret, text, StringComparison.Ordinal);
        var frame = ReadFrames(output).Last();
        Assert.Equal("fatal", frame.GetProperty("type").GetString());
        Assert.Equal("worker-failure", frame.GetProperty("code").GetString());
    }

    [Fact]
    public async Task HostBoundsLargeResultPayload()
    {
        const string taskId = "worker_test_3";
        var request = JsonSerializer.Serialize(new
        {
            version = WorkbookWorkerProtocol.Version,
            taskId,
            operation = "execute",
            plan = Plan(taskId)
        }, WorkbookWorkerProtocol.JsonOptions);
        using var input = new StringReader(request);
        using var output = new StringWriter();

        var exitCode = await WorkbookWorkerHost.RunAsync(input, output, _ => new LargeResultFakeRuntime());

        Assert.Equal(0, exitCode);
        var resultLine = output.ToString().Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries).Last();
        Assert.True(Encoding.UTF8.GetByteCount(resultLine) <= WorkbookWorkerProtocol.MaxFrameBytes);
        var result = JsonDocument.Parse(resultLine).RootElement.GetProperty("result");
        Assert.True(result.GetProperty("summary").GetString()!.Length <= WorkbookWorkerProtocol.MaxTextLength);
        Assert.True(result.GetProperty("changes").GetArrayLength() <= WorkbookWorkerProtocol.MaxResultItems);
    }

    [Fact]
    public async Task HostPreservesTwentyBoundedChangesWithinFrameLimit()
    {
        const string taskId = "worker_test_4";
        var changes = Enumerable.Range(1, WorkbookWorkerProtocol.MaxResultItems)
            .Select(index => new TaskChange("formula-repaired", $"Sheet1!A{index}", $"Repaired range {index}"))
            .ToArray();
        var request = JsonSerializer.Serialize(new
        {
            version = WorkbookWorkerProtocol.Version,
            taskId,
            operation = "execute",
            plan = Plan(taskId)
        }, WorkbookWorkerProtocol.JsonOptions);
        using var input = new StringReader(request);
        using var output = new StringWriter();

        var exitCode = await WorkbookWorkerHost.RunAsync(input, output, _ => new BoundedChangeFakeRuntime(changes));

        Assert.Equal(0, exitCode);
        var resultLine = output.ToString().Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries).Last();
        Assert.True(Encoding.UTF8.GetByteCount(resultLine) <= WorkbookWorkerProtocol.MaxFrameBytes);
        using var frame = JsonDocument.Parse(resultLine);
        var outcome = JsonSerializer.Deserialize<WorkbookExecutionOutcome>(
            frame.RootElement.GetProperty("result").GetRawText(),
            WorkbookWorkerProtocol.JsonOptions);
        Assert.NotNull(outcome);
        Assert.Equal(changes, outcome.Changes);
    }

    [Fact]
    public void BoundMacroReceiptStaysWithinTheWorkerFrameLimit()
    {
        var outcome = new WorkbookExecutionOutcome(
            ExcelTaskStatus.Planned,
            new string('s', WorkbookWorkerProtocol.MaxTextLength * 2),
            MacroProcedure: new MacroProcedureReceipt(
                new string('m', WorkbookWorkerProtocol.MaxTextLength * 2),
                new string('p', WorkbookWorkerProtocol.MaxTextLength * 2),
                new string('h', WorkbookWorkerProtocol.MaxTextLength * 2),
                new string('x', MacroProcedureText.MaxSourceCharacters * 2),
                RunRequested: false,
                RunCompleted: false));

        var bounded = WorkbookWorkerProtocol.Bound(outcome);
        var frame = JsonSerializer.Serialize(new { version = WorkbookWorkerProtocol.Version, type = "result", taskId = "worker_test_5", operation = "execute", result = bounded }, WorkbookWorkerProtocol.JsonOptions);

        Assert.True(Encoding.UTF8.GetByteCount(frame) <= WorkbookWorkerProtocol.MaxFrameBytes);
        Assert.NotNull(bounded.MacroProcedure);
        Assert.Equal(WorkbookWorkerProtocol.MaxTextLength, bounded.MacroProcedure.ComponentName.Length);
        Assert.Equal(WorkbookWorkerProtocol.MaxTextLength, bounded.MacroProcedure.ProcedureName.Length);
        Assert.Equal(WorkbookWorkerProtocol.MaxTextLength, bounded.MacroProcedure.Sha256.Length);
        Assert.Equal(MacroProcedureText.MaxSourceCharacters, bounded.MacroProcedure.Source!.Length);
    }

    [Fact]
    public async Task WorkerWatchdogRecoversOnlyTheIdentityCapturedInsideTheWorker()
    {
        var identity = new ProcessIdentity(1234, DateTime.UtcNow, "C:\\Excel\\EXCEL.EXE");
        var terminator = new RecordingTerminator();
        var recovery = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        await using var watchdog = new WorkerOwnedProcessWatchdog(
            TimeSpan.FromMilliseconds(20),
            terminator,
            () => recovery.TrySetResult());

        watchdog.Track(identity);
        await recovery.Task.WaitAsync(TimeSpan.FromSeconds(1));

        Assert.Equal(identity, Assert.Single(terminator.Identities));
    }

    [Fact]
    public async Task CompletedWorkerStopsTheWatchdogBeforeItsDeadline()
    {
        var terminator = new RecordingTerminator();
        var watchdog = new WorkerOwnedProcessWatchdog(
            TimeSpan.FromMilliseconds(50),
            terminator,
            () => { });
        watchdog.Track(new ProcessIdentity(1234, DateTime.UtcNow, "C:\\Excel\\EXCEL.EXE"));

        await watchdog.DisposeAsync();
        await Task.Delay(75);

        Assert.Empty(terminator.Identities);
    }

    private static ExcelTaskPlan Plan(string taskId) => new(taskId, ExcelTaskPlans.Copy(
        "target.xlsx", "reference.xlsx", "Reference", "New Sheet", [], ExcelTaskMode.Apply, WorkbookBinding.Isolated, SaveMode.Same, null, true));

    private static List<JsonElement> ReadFrames(StringWriter output) => output.ToString()
        .Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries)
        .Select(line => JsonDocument.Parse(line).RootElement.Clone())
        .ToList();

    private sealed class ObservingFakeRuntime(IExcelWorkbookRuntimeObserver observer) : IWorkbookRuntime, IDisposable
    {
        public bool Disposed { get; private set; }

        public Task<WorkbookInspection> InspectAsync(WorkbookInspectionRequest request, CancellationToken cancellationToken)
        {
            observer.OnPhase("fake-inspection");
            using var current = System.Diagnostics.Process.GetCurrentProcess();
            observer.OnOwnedProcessCaptured(ProcessIdentity.Capture(current.Id));
            observer.OnStagingPathCreated("C:\\worker-staging.xlsx");
            return Task.FromResult(new WorkbookInspection(true, Checks: [new TaskCheck("fake", true, "safe")]));
        }

        public Task<WorkbookExecutionOutcome> ExecuteAsync(ExcelTaskPlan plan, CancellationToken cancellationToken) => throw new NotSupportedException();

        public void Dispose() => Disposed = true;
    }

    private sealed class ThrowingFakeRuntime(string message) : IWorkbookRuntime
    {
        public Task<WorkbookInspection> InspectAsync(WorkbookInspectionRequest request, CancellationToken cancellationToken) => throw new NotSupportedException();

        public Task<WorkbookExecutionOutcome> ExecuteAsync(ExcelTaskPlan plan, CancellationToken cancellationToken) => throw new InvalidOperationException(message);
    }

    private sealed class LargeResultFakeRuntime : IWorkbookRuntime
    {
        public Task<WorkbookInspection> InspectAsync(WorkbookInspectionRequest request, CancellationToken cancellationToken) => throw new NotSupportedException();

        public Task<WorkbookExecutionOutcome> ExecuteAsync(ExcelTaskPlan plan, CancellationToken cancellationToken) => Task.FromResult(new WorkbookExecutionOutcome(
            ExcelTaskStatus.Completed,
            new string('s', WorkbookWorkerProtocol.MaxTextLength * 20),
            Enumerable.Range(0, WorkbookWorkerProtocol.MaxResultItems * 3)
                .Select(index => new TaskChange(new string('k', 600), new string('t', 600), new string('d', 600)))
                .ToArray(),
            Enumerable.Range(0, WorkbookWorkerProtocol.MaxResultItems * 3)
                .Select(index => new TaskCheck(new string('n', 600), true, new string('d', 600)))
                .ToArray()));
    }

    private sealed class BoundedChangeFakeRuntime(IReadOnlyList<TaskChange> changes) : IWorkbookRuntime
    {
        public Task<WorkbookInspection> InspectAsync(WorkbookInspectionRequest request, CancellationToken cancellationToken) => throw new NotSupportedException();

        public Task<WorkbookExecutionOutcome> ExecuteAsync(ExcelTaskPlan plan, CancellationToken cancellationToken) => Task.FromResult(new WorkbookExecutionOutcome(
            ExcelTaskStatus.Completed,
            "Completed",
            changes));
    }

    private sealed class RecordingTerminator : IWorkerOwnedProcessTerminator
    {
        public List<ProcessIdentity> Identities { get; } = [];

        public bool TryTerminateExact(ProcessIdentity identity)
        {
            Identities.Add(identity);
            return true;
        }
    }
}
