using ExcelTask.Core;

namespace ExcelTask.Excel.Tests;

public sealed class SupervisedWorkbookRuntimeTests
{
    [Fact]
    public void LaunchSpecUsesAppHostWithoutPassingTheAssemblyAgain()
    {
        var provider = new CurrentProcessWorkerLaunchSpecProvider(
            () => "C:\\apps\\ExcelTask.McpServer.exe",
            () => "C:\\apps\\ExcelTask.McpServer.dll");

        var launch = provider.GetLaunchSpec();

        Assert.Equal("C:\\apps\\ExcelTask.McpServer.exe", launch.FileName);
        Assert.Empty(launch.PrefixArguments);
    }

    [Fact]
    public void LaunchSpecUsesDotnetHostWithTheEntryAssembly()
    {
        var provider = new CurrentProcessWorkerLaunchSpecProvider(
            () => "C:\\Program Files\\dotnet\\dotnet.exe",
            () => "C:\\apps\\ExcelTask.McpServer.dll");

        var launch = provider.GetLaunchSpec();

        Assert.Equal("C:\\Program Files\\dotnet\\dotnet.exe", launch.FileName);
        Assert.Equal(["C:\\apps\\ExcelTask.McpServer.dll"], launch.PrefixArguments);
    }

    [Fact]
    public void LaunchSpecDoesNotDependOnTheCurrentWorkingDirectory()
    {
        var originalDirectory = Directory.GetCurrentDirectory();
        var alternateDirectory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(alternateDirectory);
        try
        {
            Directory.SetCurrentDirectory(alternateDirectory);
            var provider = new CurrentProcessWorkerLaunchSpecProvider(
                () => "C:\\apps\\ExcelTask.McpServer.exe",
                () => "C:\\apps\\ExcelTask.McpServer.dll");

            var launch = provider.GetLaunchSpec();

            Assert.Equal("C:\\apps\\ExcelTask.McpServer.exe", launch.FileName);
            Assert.Empty(launch.PrefixArguments);
        }
        finally
        {
            Directory.SetCurrentDirectory(originalDirectory);
            Directory.Delete(alternateDirectory, recursive: true);
        }
    }

    [Fact]
    public async Task ExecutePassesThroughTheVerifiedWorkerOutcome()
    {
        var expected = new WorkbookExecutionOutcome(ExcelTaskStatus.Completed, "Verified by worker.");
        var session = new FakeWorkerSession(
            new SupervisedWorkerFrame(Accepted: true, Phase: "accepted"),
            new SupervisedWorkerFrame(Execution: expected));
        using var runtime = CreateRuntime(new FakeWorkerClient(session));

        var actual = await runtime.ExecuteAsync(CreatePlan(), CancellationToken.None);

        Assert.Equal(expected, actual);
        Assert.Equal(1, session.StopCount);
    }

    [Fact]
    public async Task ExecuteReturnsUnknownWhenWorkerResultArrivesButProcessExitIsUnproven()
    {
        var session = new FakeWorkerSession(
            new SupervisedWorkerFrame(Accepted: true),
            new SupervisedWorkerFrame(Execution: new WorkbookExecutionOutcome(ExcelTaskStatus.Completed, "Verified by worker.")))
        {
            StopResult = false
        };
        using var runtime = CreateRuntime(new FakeWorkerClient(session));

        var outcome = await runtime.ExecuteAsync(CreatePlan(), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Unknown, outcome.Status);
        Assert.False(outcome.CanRetry);
        Assert.Contains(outcome.Checks!, check => check.Name == "worker-process-exit" && !check.Passed);
    }

    [Fact]
    public async Task ExecuteRejectsWhenWorkerCannotStartBeforeAcceptance()
    {
        using var runtime = CreateRuntime(new ThrowingWorkerClient());

        var outcome = await runtime.ExecuteAsync(CreatePlan(), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Rejected, outcome.Status);
        Assert.True(outcome.CanRetry);
    }

    [Fact]
    public async Task ExecuteReturnsUnknownWhenWorkerRequestMayHaveBeenDispatchedBeforeStartFailed()
    {
        using var runtime = CreateRuntime(new AmbiguousDispatchWorkerClient());

        var outcome = await runtime.ExecuteAsync(CreatePlan(), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Unknown, outcome.Status);
        Assert.False(outcome.CanRetry);
    }

    [Fact]
    public async Task ExecuteReturnsUnknownWhenStartedWorkerEndsBeforeAcceptance()
    {
        var session = new FakeWorkerSession(_ => Task.FromResult<SupervisedWorkerFrame?>(null));
        using var runtime = CreateRuntime(new FakeWorkerClient(session));

        var outcome = await runtime.ExecuteAsync(CreatePlan(), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Unknown, outcome.Status);
        Assert.False(outcome.CanRetry);
        Assert.Equal(1, session.StopCount);
    }

    [Fact]
    public async Task ExecuteReturnsUnknownWhenAcceptedWorkerTimesOut()
    {
        var session = new FakeWorkerSession(
            new SupervisedWorkerFrame(Accepted: true, Phase: "accepted"),
            cancellationToken => Task.Delay(Timeout.InfiniteTimeSpan, cancellationToken).ContinueWith<SupervisedWorkerFrame?>(_ => null));
        using var runtime = CreateRuntime(new FakeWorkerClient(session), deadline: TimeSpan.FromMilliseconds(20));

        var outcome = await runtime.ExecuteAsync(CreatePlan(), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Unknown, outcome.Status);
        Assert.False(outcome.CanRetry);
        Assert.Equal(1, session.StopCount);
    }

    [Fact]
    public async Task AWorkerTimeoutDoesNotPoisonTheNextTask()
    {
        var timedOut = new FakeWorkerSession(
            new SupervisedWorkerFrame(Accepted: true),
            cancellationToken => Task.Delay(Timeout.InfiniteTimeSpan, cancellationToken).ContinueWith<SupervisedWorkerFrame?>(_ => null));
        var completed = new FakeWorkerSession(
            new SupervisedWorkerFrame(Accepted: true),
            new SupervisedWorkerFrame(Execution: new WorkbookExecutionOutcome(ExcelTaskStatus.Completed, "Verified.")));
        using var runtime = CreateRuntime(new SequenceWorkerClient(timedOut, completed), deadline: TimeSpan.FromMilliseconds(20));

        var first = await runtime.ExecuteAsync(CreatePlan(), CancellationToken.None);
        var second = await runtime.ExecuteAsync(CreatePlan(), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Unknown, first.Status);
        Assert.Equal(ExcelTaskStatus.Completed, second.Status);
    }

    [Fact]
    public async Task ExecuteReturnsUnknownWhenCancelledAfterWorkerAcceptance()
    {
        var accepted = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var session = new FakeWorkerSession(
            new SupervisedWorkerFrame(Accepted: true, Phase: "accepted"),
            cancellationToken =>
            {
                accepted.TrySetResult();
                return Task.Delay(Timeout.InfiniteTimeSpan, cancellationToken).ContinueWith<SupervisedWorkerFrame?>(_ => null);
            });
        using var runtime = CreateRuntime(new FakeWorkerClient(session), deadline: TimeSpan.FromSeconds(1));
        using var cancellation = new CancellationTokenSource();

        var execution = runtime.ExecuteAsync(CreatePlan(), cancellation.Token);
        await accepted.Task;
        cancellation.Cancel();
        var outcome = await execution;

        Assert.Equal(ExcelTaskStatus.Unknown, outcome.Status);
        Assert.False(outcome.CanRetry);
    }

    [Fact]
    public async Task ExecuteReturnsUnknownForMalformedResultAfterAcceptance()
    {
        var session = new FakeWorkerSession(
            new SupervisedWorkerFrame(Accepted: true, Phase: "accepted"),
            new SupervisedWorkerFrame(Phase: new string('x', 65)));
        using var runtime = CreateRuntime(new FakeWorkerClient(session));

        var outcome = await runtime.ExecuteAsync(CreatePlan(), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Unknown, outcome.Status);
        Assert.False(outcome.CanRetry);
    }

    [Fact]
    public async Task ExecuteReturnsUnknownForTruncatedResultAfterAcceptance()
    {
        var session = new FakeWorkerSession(
            new SupervisedWorkerFrame(Accepted: true, Phase: "accepted"),
            _ => Task.FromResult<SupervisedWorkerFrame?>(null));
        using var runtime = CreateRuntime(new FakeWorkerClient(session));

        var outcome = await runtime.ExecuteAsync(CreatePlan(), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Unknown, outcome.Status);
        Assert.False(outcome.CanRetry);
    }

    [Fact]
    public async Task IdentityMismatchIsNotAcceptedAsExited()
    {
        var processes = new ExactMatchOnlyProcessObserver(new SupervisedExcelIdentity(9876, DateTime.UtcNow, "C:\\Excel\\EXCEL.EXE"));
        var session = new FakeWorkerSession(
            new SupervisedWorkerFrame(
                Accepted: true,
                Phase: "accepted",
                ExcelIdentity: new SupervisedExcelIdentity(1234, DateTime.UtcNow, "C:\\Excel\\EXCEL.EXE")),
            new SupervisedWorkerFrame(Fatal: true));
        using var runtime = CreateRuntime(new FakeWorkerClient(session), processes);

        var outcome = await runtime.ExecuteAsync(CreatePlan(binding: WorkbookBinding.Isolated), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Unknown, outcome.Status);
        Assert.Equal(1, processes.AttemptCount);
    }

    [Fact]
    public async Task NonExcelProcessIdentityIsNeverSentToTheExitObserver()
    {
        var processes = new ExactMatchOnlyProcessObserver(null);
        var session = new FakeWorkerSession(
            new SupervisedWorkerFrame(
                Accepted: true,
                ExcelIdentity: new SupervisedExcelIdentity(1234, DateTime.UtcNow, "C:\\Windows\\System32\\notepad.exe")),
            new SupervisedWorkerFrame(Fatal: true));
        using var runtime = CreateRuntime(new FakeWorkerClient(session), processes);

        var outcome = await runtime.ExecuteAsync(CreatePlan(), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Unknown, outcome.Status);
        Assert.Equal(0, processes.AttemptCount);
    }

    [Fact]
    public void PreExistingExcelIdentityIsNotAuthorizedAsWorkerOwned()
    {
        var identity = new SupervisedExcelIdentity(1234, DateTime.UtcNow, "C:\\Excel\\EXCEL.EXE");

        Assert.False(JsonLineWorkbookWorkerSession.IsAuthorizedExcelIdentity(identity, [identity]));
        Assert.True(JsonLineWorkbookWorkerSession.IsAuthorizedExcelIdentity(
            identity,
            [identity with { StartTimeUtc = identity.StartTimeUtc.AddSeconds(-1) }]));
    }

    [Fact]
    public async Task OwnedExcelStillRunningAfterWorkerExitIsReported()
    {
        var identity = new SupervisedExcelIdentity(1234, DateTime.UtcNow, "C:\\Excel\\EXCEL.EXE");
        var processes = new ExactMatchOnlyProcessObserver(identity, exitResult: false);
        var session = new FakeWorkerSession(
            new SupervisedWorkerFrame(Accepted: true),
            new SupervisedWorkerFrame(ExcelIdentity: identity),
            new SupervisedWorkerFrame(Fatal: true));
        using var runtime = CreateRuntime(new FakeWorkerClient(session), processes);

        var outcome = await runtime.ExecuteAsync(CreatePlan(), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Unknown, outcome.Status);
        Assert.Contains(outcome.Checks!, check => check.Name == "owned-process-exit" && !check.Passed);
    }

    [Fact]
    public async Task UseOpenNeverInspectsTheReportedExcelProcessForCleanup()
    {
        var identity = new SupervisedExcelIdentity(1234, DateTime.UtcNow, "C:\\Excel\\EXCEL.EXE");
        var processes = new ExactMatchOnlyProcessObserver(identity);
        var session = new FakeWorkerSession(
            new SupervisedWorkerFrame(Accepted: true),
            new SupervisedWorkerFrame(ExcelIdentity: identity),
            new SupervisedWorkerFrame(Fatal: true));
        using var runtime = CreateRuntime(new FakeWorkerClient(session), processes);

        var outcome = await runtime.ExecuteAsync(CreatePlan(binding: WorkbookBinding.UseOpen), CancellationToken.None);

        Assert.Equal(ExcelTaskStatus.Unknown, outcome.Status);
        Assert.Equal(0, processes.AttemptCount);
    }

    [Fact]
    public async Task CleanupDeletesOnlyTheExactReportedStagingArtifact()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var output = Path.Combine(directory, "output.xlsx");
        var exact = Path.Combine(directory, ".output.excel-task-task-" + Guid.NewGuid().ToString("N") + ".xlsx");
        var unrelated = Path.Combine(directory, ".output.excel-task-other-" + Guid.NewGuid().ToString("N") + ".xlsx");
        File.WriteAllText(exact, "exact");
        File.WriteAllText(unrelated, "unrelated");
        try
        {
            var session = new FakeWorkerSession(
                new SupervisedWorkerFrame(Accepted: true, Phase: "accepted"),
                new SupervisedWorkerFrame(StagingArtifactPath: exact),
                new SupervisedWorkerFrame(StagingArtifactPath: unrelated),
                new SupervisedWorkerFrame(Fatal: true));
            using var runtime = CreateRuntime(new FakeWorkerClient(session));

            var outcome = await runtime.ExecuteAsync(CreatePlan(output: output), CancellationToken.None);

            Assert.Equal(ExcelTaskStatus.Unknown, outcome.Status);
            Assert.False(File.Exists(exact));
            Assert.True(File.Exists(unrelated));
        }
        finally
        {
            if (File.Exists(exact)) File.Delete(exact);
            if (File.Exists(unrelated)) File.Delete(unrelated);
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public async Task LockedTaskStagingArtifactIsReported()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var output = Path.Combine(directory, "output.xlsx");
        var exact = Path.Combine(directory, ".output.excel-task-task-" + Guid.NewGuid().ToString("N") + ".xlsx");
        File.WriteAllText(exact, "exact");
        try
        {
            using (var lockStream = new FileStream(exact, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
            {
                var session = new FakeWorkerSession(
                    new SupervisedWorkerFrame(Accepted: true),
                    new SupervisedWorkerFrame(StagingArtifactPath: exact),
                    new SupervisedWorkerFrame(Fatal: true));
                using var runtime = CreateRuntime(new FakeWorkerClient(session));

                var outcome = await runtime.ExecuteAsync(CreatePlan(output: output), CancellationToken.None);

                Assert.Equal(ExcelTaskStatus.Unknown, outcome.Status);
                Assert.Contains(outcome.Checks!, check => check.Name == "staging-cleanup" && !check.Passed);
                Assert.True(File.Exists(exact));
            }
        }
        finally
        {
            if (File.Exists(exact)) File.Delete(exact);
            Directory.Delete(directory, recursive: true);
        }
    }

    private static SupervisedWorkbookRuntime CreateRuntime(
        IWorkbookWorkerClient client,
        IReportedExcelProcessObserver? processes = null,
        TimeSpan? deadline = null) => new(client, processes ?? new ExactMatchOnlyProcessObserver(null), deadline ?? TimeSpan.FromSeconds(1));

    private static ExcelTaskPlan CreatePlan(string? output = null, WorkbookBinding binding = WorkbookBinding.Isolated) => new(
        "task",
        ExcelTaskPlans.Copy(
            "target.xlsx",
            "reference.xlsx",
            "Reference",
            "Imported",
            [],
            ExcelTaskMode.Apply,
            binding,
            output is null ? SaveMode.Same : SaveMode.Copy,
            output,
            overwrite: false));

    private sealed class ThrowingWorkerClient : IWorkbookWorkerClient
    {
        public Task<IWorkbookWorkerSession> StartAsync(WorkbookWorkerOperation operation, WorkbookInspectionRequest? inspectionRequest, ExcelTaskPlan? plan, CancellationToken cancellationToken) =>
            throw new InvalidOperationException("not exposed");
    }

    private sealed class AmbiguousDispatchWorkerClient : IWorkbookWorkerClient
    {
        public Task<IWorkbookWorkerSession> StartAsync(WorkbookWorkerOperation operation, WorkbookInspectionRequest? inspectionRequest, ExcelTaskPlan? plan, CancellationToken cancellationToken) =>
            throw new WorkbookWorkerDispatchException(new IOException("not exposed"));
    }

    private sealed class FakeWorkerClient(FakeWorkerSession session) : IWorkbookWorkerClient
    {
        public Task<IWorkbookWorkerSession> StartAsync(WorkbookWorkerOperation operation, WorkbookInspectionRequest? inspectionRequest, ExcelTaskPlan? plan, CancellationToken cancellationToken) =>
            Task.FromResult<IWorkbookWorkerSession>(session);
    }

    private sealed class SequenceWorkerClient(params FakeWorkerSession[] sessions) : IWorkbookWorkerClient
    {
        private readonly Queue<FakeWorkerSession> _sessions = new(sessions);

        public Task<IWorkbookWorkerSession> StartAsync(WorkbookWorkerOperation operation, WorkbookInspectionRequest? inspectionRequest, ExcelTaskPlan? plan, CancellationToken cancellationToken) =>
            Task.FromResult<IWorkbookWorkerSession>(_sessions.Dequeue());
    }

    private sealed class FakeWorkerSession : IWorkbookWorkerSession
    {
        private readonly Queue<Func<CancellationToken, Task<SupervisedWorkerFrame?>>> _reads;

        public FakeWorkerSession(params SupervisedWorkerFrame[] frames)
            : this(frames.Select(frame => new Func<CancellationToken, Task<SupervisedWorkerFrame?>>(_ => Task.FromResult<SupervisedWorkerFrame?>(frame))).ToArray())
        {
        }

        public FakeWorkerSession(SupervisedWorkerFrame first, Func<CancellationToken, Task<SupervisedWorkerFrame?>> second)
            : this([_ => Task.FromResult<SupervisedWorkerFrame?>(first), second])
        {
        }

        public FakeWorkerSession(Func<CancellationToken, Task<SupervisedWorkerFrame?>> read)
            : this([read])
        {
        }

        private FakeWorkerSession(params Func<CancellationToken, Task<SupervisedWorkerFrame?>>[] reads) => _reads = new(reads);

        public int StopCount { get; private set; }

        public bool StopResult { get; init; } = true;

        public Task<SupervisedWorkerFrame?> ReadAsync(CancellationToken cancellationToken) => _reads.Count == 0
            ? Task.FromResult<SupervisedWorkerFrame?>(null)
            : _reads.Dequeue()(cancellationToken);

        public Task<bool> StopAsync()
        {
            StopCount++;
            return Task.FromResult(StopResult);
        }

        public ValueTask DisposeAsync() => ValueTask.CompletedTask;
    }

    private sealed class ExactMatchOnlyProcessObserver(
        SupervisedExcelIdentity? expected,
        bool exitResult = true) : IReportedExcelProcessObserver
    {
        public int AttemptCount { get; private set; }

        public bool HasExitedExact(SupervisedExcelIdentity identity)
        {
            AttemptCount++;
            if (identity != expected) return false;
            return exitResult;
        }
    }
}
