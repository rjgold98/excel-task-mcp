using System.Diagnostics;
using System.Runtime.Versioning;
using System.Text.Json;
using ExcelTask.Core;

namespace ExcelTask.Excel;

/// <summary>
/// Contains one workbook operation in a short-lived worker process. The host never
/// waits on Excel directly, so an accepted operation that loses its worker is unknown.
/// </summary>
[SupportedOSPlatform("windows")]
public sealed class SupervisedWorkbookRuntime : IWorkbookRuntime, IDisposable
{
    private static readonly TimeSpan DefaultDeadline = TimeSpan.FromMinutes(2);
    private readonly IWorkbookWorkerClient _workers;
    private readonly IReportedExcelProcessObserver _excelProcessObserver;
    private readonly TimeSpan _deadline;
    private bool _disposed;

    public SupervisedWorkbookRuntime()
        : this(new JsonLineWorkbookWorkerClient(), new ReportedExcelProcessObserver(), DefaultDeadline)
    {
    }

    internal SupervisedWorkbookRuntime(
        IWorkbookWorkerClient workers,
        IReportedExcelProcessObserver excelProcessObserver,
        TimeSpan deadline)
    {
        _workers = workers ?? throw new ArgumentNullException(nameof(workers));
        _excelProcessObserver = excelProcessObserver ?? throw new ArgumentNullException(nameof(excelProcessObserver));
        ArgumentOutOfRangeException.ThrowIfLessThanOrEqual(deadline, TimeSpan.Zero);
        _deadline = deadline;
    }

    public async Task<WorkbookInspection> InspectAsync(WorkbookInspectionRequest request, CancellationToken cancellationToken)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        ArgumentNullException.ThrowIfNull(request);

        var result = await RunAsync(WorkbookWorkerOperation.Inspect, request, null, cancellationToken).ConfigureAwait(false);
        if (result.Inspection is not null) return result.Inspection;

        // IWorkbookRuntime cannot express an unknown inspection. Do not leak worker failures.
        throw new InvalidOperationException("Workbook inspection could not be completed by the private worker.");
    }

    public async Task<WorkbookExecutionOutcome> ExecuteAsync(ExcelTaskPlan plan, CancellationToken cancellationToken)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        ArgumentNullException.ThrowIfNull(plan);

        var result = await RunAsync(WorkbookWorkerOperation.Execute, null, plan, cancellationToken).ConfigureAwait(false);
        return result.Execution ?? Rejected("The private workbook worker did not return an execution result.");
    }

    public void Dispose() => _disposed = true;

    private async Task<SupervisedRunResult> RunAsync(
        WorkbookWorkerOperation operation,
        WorkbookInspectionRequest? inspectionRequest,
        ExcelTaskPlan? plan,
        CancellationToken cancellationToken)
    {
        IWorkbookWorkerSession? session = null;
        var accepted = false;
        var excelIdentities = new List<SupervisedExcelIdentity>(4);
        var stagingArtifacts = new List<string>(4);
        var cleanupRequired = false;
        var sessionStopped = false;
        List<TaskCheck>? unknownChecks = null;

        SupervisedRunResult UnknownResult(string summary)
        {
            unknownChecks ??=
            [
                new TaskCheck("worker-supervision", false, "The accepted operation requires workbook reconciliation.")
            ];
            return new SupervisedRunResult(null, new WorkbookExecutionOutcome(
                ExcelTaskStatus.Unknown,
                summary,
                Checks: unknownChecks,
                CanRetry: false,
                RetryReason: "Inspect workbook state before retrying."));
        }

        void AddCleanupFailure(string name, string message)
        {
            if (unknownChecks is null || unknownChecks.Any(check => string.Equals(check.Name, name, StringComparison.Ordinal))) return;
            unknownChecks.Add(new TaskCheck(name, false, message));
        }

        try
        {
            try
            {
                session = await _workers.StartAsync(operation, inspectionRequest, plan, cancellationToken).ConfigureAwait(false);
            }
            catch (WorkbookWorkerDispatchException)
            {
                return operation == WorkbookWorkerOperation.Execute
                    ? UnknownResult("The private workbook worker may have received execution, but acceptance was not observed.")
                    : BeforeAcceptance(operation, "The private workbook worker could not complete inspection startup.");
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
            {
                return BeforeAcceptance(operation, "The workbook operation was cancelled before the private worker accepted it.");
            }
            catch
            {
                return BeforeAcceptance(operation, "The private workbook worker could not be started.");
            }

            using var deadline = new CancellationTokenSource(_deadline);
            using var readCancellation = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, deadline.Token);
            for (var frameCount = 0; frameCount < 128; frameCount++)
            {
                SupervisedWorkerFrame? frame;
                try
                {
                    frame = await session.ReadAsync(readCancellation.Token).ConfigureAwait(false);
                }
                catch (OperationCanceledException) when (readCancellation.IsCancellationRequested)
                {
                    cleanupRequired = true;
                    return operation == WorkbookWorkerOperation.Execute
                        ? UnknownResult(accepted
                            ? "The private workbook worker stopped before its accepted operation was verified."
                            : "The private workbook worker may have received execution, but acceptance was not observed.")
                        : BeforeAcceptance(operation, "The private workbook worker did not complete inspection.");
                }
                catch
                {
                    cleanupRequired = true;
                    return operation == WorkbookWorkerOperation.Execute
                        ? UnknownResult(accepted
                            ? "The private workbook worker ended without a verified result."
                            : "The private workbook worker may have received execution, but no acceptance or result was observed.")
                        : BeforeAcceptance(operation, "The private workbook worker did not complete inspection.");
                }

                if (frame is null || !frame.IsWellFormed())
                {
                    cleanupRequired = true;
                    return operation == WorkbookWorkerOperation.Execute
                        ? UnknownResult(accepted
                            ? "The private workbook worker returned an incomplete result after accepting the operation."
                            : "The private workbook worker may have received execution, but its response was incomplete.")
                        : BeforeAcceptance(operation, "The private workbook worker did not complete inspection.");
                }

                if (frame.ExcelIdentity is not null && excelIdentities.Count < 4 && !excelIdentities.Contains(frame.ExcelIdentity))
                {
                    excelIdentities.Add(frame.ExcelIdentity);
                }
                if (IsExactStagingArtifact(frame.StagingArtifactPath, plan) && stagingArtifacts.Count < 4 && !stagingArtifacts.Contains(frame.StagingArtifactPath!, StringComparer.OrdinalIgnoreCase))
                {
                    stagingArtifacts.Add(frame.StagingArtifactPath!);
                }
                if (frame.Accepted) accepted = true;

                if (frame.Fatal)
                {
                    cleanupRequired = accepted;
                    return accepted
                        ? UnknownResult("The private workbook worker could not verify its accepted operation.")
                        : BeforeAcceptance(operation, "The private workbook worker rejected the operation before acceptance.");
                }

                if (!accepted && !frame.Accepted)
                {
                    cleanupRequired = true;
                    return operation == WorkbookWorkerOperation.Execute
                        ? UnknownResult("The private workbook worker may have received execution, but its acceptance handshake was invalid.")
                        : BeforeAcceptance(operation, "The private workbook worker did not complete inspection.");
                }

                if (operation == WorkbookWorkerOperation.Inspect && frame.Inspection is not null)
                {
                    sessionStopped = await session.StopAsync().ConfigureAwait(false);
                    if (!sessionStopped)
                    {
                        cleanupRequired = true;
                        return BeforeAcceptance(operation, "The private workbook worker did not prove process exit after inspection.");
                    }
                    return new SupervisedRunResult(frame.Inspection, null);
                }

                if (operation == WorkbookWorkerOperation.Execute && frame.Execution is not null)
                {
                    sessionStopped = await session.StopAsync().ConfigureAwait(false);
                    if (!sessionStopped)
                    {
                        cleanupRequired = true;
                        return UnknownResult("The private workbook worker returned a result but did not prove process exit.");
                    }
                    return new SupervisedRunResult(null, NormalizeExecution(frame.Execution));
                }
            }

            cleanupRequired = true;
            return operation == WorkbookWorkerOperation.Execute
                ? UnknownResult(accepted
                    ? "The private workbook worker did not finish its accepted operation."
                    : "The private workbook worker may have received execution, but acceptance was not observed.")
                : BeforeAcceptance(operation, "The private workbook worker did not complete inspection.");
        }
        finally
        {
            if (session is not null && !sessionStopped)
            {
                try { sessionStopped = await session.StopAsync().ConfigureAwait(false); }
                catch { sessionStopped = false; }
                if (!sessionStopped)
                {
                    AddCleanupFailure("worker-process-exit", "The private workbook worker did not prove process exit.");
                }
            }

            if (cleanupRequired)
            {
                if (plan?.Request.WorkbookBinding != WorkbookBinding.UseOpen)
                {
                    foreach (var excelIdentity in excelIdentities)
                    {
                        var exited = false;
                        try { exited = _excelProcessObserver.HasExitedExact(excelIdentity); }
                        catch { }
                        if (!exited)
                        {
                            AddCleanupFailure("owned-process-exit", "An owned Excel process did not prove exit.");
                        }
                    }
                }

                foreach (var stagingArtifact in stagingArtifacts)
                {
                    if (!TryDeleteExactStagingArtifact(stagingArtifact, plan))
                    {
                        AddCleanupFailure("staging-cleanup", "A task staging workbook could not be deleted.");
                    }
                }
            }

            if (session is not null)
            {
                try { await session.DisposeAsync().ConfigureAwait(false); }
                catch { }
            }
        }
    }

    private static SupervisedRunResult BeforeAcceptance(WorkbookWorkerOperation operation, string message) => operation == WorkbookWorkerOperation.Execute
        ? new(null, Rejected(message))
        : new(null, null);

    private static WorkbookExecutionOutcome NormalizeExecution(WorkbookExecutionOutcome outcome) => outcome.Status == ExcelTaskStatus.Unknown
        ? outcome with { CanRetry = false, RetryReason = "Inspect workbook state before retrying." }
        : outcome;

    private static WorkbookExecutionOutcome Rejected(string summary) => new(
        ExcelTaskStatus.Rejected,
        summary,
        Checks: [new TaskCheck("worker-launch", false, "The private worker did not accept execution.")],
        CanRetry: true,
        RetryReason: "Correct the worker startup issue and retry.");

    private static bool IsExactStagingArtifact(string? artifact, ExcelTaskPlan? plan)
    {
        if (artifact is null || plan?.Request.Save != SaveMode.Copy || string.IsNullOrWhiteSpace(plan.Request.OutputWorkbookPath) ||
            !WorkbookRuntimeHelpers.IsSafeTaskId(plan.TaskId)) return false;
        try
        {
            var output = Path.GetFullPath(plan.Request.OutputWorkbookPath);
            var candidate = Path.GetFullPath(artifact);
            var directory = Path.GetDirectoryName(output);
            var extension = Path.GetExtension(output);
            var name = Path.GetFileNameWithoutExtension(output);
            var prefix = $".{name}.excel-task-{plan.TaskId}-";
            return string.Equals(Path.GetDirectoryName(candidate), directory, StringComparison.OrdinalIgnoreCase) &&
                   string.Equals(Path.GetExtension(candidate), extension, StringComparison.OrdinalIgnoreCase) &&
                   Path.GetFileNameWithoutExtension(candidate).StartsWith(prefix, StringComparison.OrdinalIgnoreCase) &&
                   Guid.TryParseExact(Path.GetFileNameWithoutExtension(candidate)[prefix.Length..], "N", out _);
        }
        catch (ArgumentException) { return false; }
        catch (NotSupportedException) { return false; }
    }

    private static bool TryDeleteExactStagingArtifact(string? artifact, ExcelTaskPlan? plan)
    {
        return IsExactStagingArtifact(artifact, plan) && WorkbookRuntimeHelpers.TryDeleteStaging(artifact!);
    }
}

internal enum WorkbookWorkerOperation { Inspect, Execute }

internal sealed record SupervisedExcelIdentity(int ProcessId, DateTime StartTimeUtc, string ExecutablePath)
{
    public bool IsWellFormed() => ProcessId > 0 && StartTimeUtc != default &&
        !string.IsNullOrWhiteSpace(ExecutablePath) && Path.IsPathFullyQualified(ExecutablePath);

    public bool IsExcelProcess() => IsWellFormed() &&
        string.Equals(Path.GetFileName(ExecutablePath), "EXCEL.EXE", StringComparison.OrdinalIgnoreCase);

    public bool MatchesExactly(SupervisedExcelIdentity other) => ProcessId == other.ProcessId &&
        StartTimeUtc == other.StartTimeUtc &&
        string.Equals(ExecutablePath, other.ExecutablePath, StringComparison.OrdinalIgnoreCase);
}

internal sealed record SupervisedWorkerFrame(
    bool Accepted = false,
    string? Phase = null,
    SupervisedExcelIdentity? ExcelIdentity = null,
    string? StagingArtifactPath = null,
    WorkbookInspection? Inspection = null,
    WorkbookExecutionOutcome? Execution = null,
    bool Fatal = false)
{
    public bool IsWellFormed() =>
        (Phase is null || (Phase.Length is > 0 and <= WorkbookWorkerProtocol.MaxPhaseLength)) &&
        (ExcelIdentity is null || ExcelIdentity.IsExcelProcess()) &&
        !(Inspection is not null && Execution is not null) &&
        !(Fatal && (Inspection is not null || Execution is not null));
}

internal sealed record SupervisedRunResult(WorkbookInspection? Inspection, WorkbookExecutionOutcome? Execution);

internal sealed class WorkbookWorkerDispatchException(Exception innerException)
    : Exception("The private worker request may have been dispatched.", innerException);

internal interface IWorkbookWorkerClient
{
    Task<IWorkbookWorkerSession> StartAsync(
        WorkbookWorkerOperation operation,
        WorkbookInspectionRequest? inspectionRequest,
        ExcelTaskPlan? plan,
        CancellationToken cancellationToken);
}

internal interface IWorkbookWorkerSession : IAsyncDisposable
{
    Task<SupervisedWorkerFrame?> ReadAsync(CancellationToken cancellationToken);

    Task<bool> StopAsync();
}

internal interface IReportedExcelProcessObserver
{
    bool HasExitedExact(SupervisedExcelIdentity identity);
}

internal sealed class ReportedExcelProcessObserver : IReportedExcelProcessObserver
{
    public bool HasExitedExact(SupervisedExcelIdentity identity)
    {
        if (!identity.IsExcelProcess()) return false;
        try
        {
            using var process = Process.GetProcessById(identity.ProcessId);
            if (process.StartTime.ToUniversalTime() != identity.StartTimeUtc ||
                !string.Equals(process.MainModule?.FileName, identity.ExecutablePath, StringComparison.OrdinalIgnoreCase)) return true;
            return process.HasExited;
        }
        catch (ArgumentException) { return true; }
        catch (InvalidOperationException) { return true; }
        catch (System.ComponentModel.Win32Exception) { return false; }
    }
}

/// <summary>Private JSON-lines client for <see cref="WorkbookWorkerProtocol"/>.</summary>
internal sealed class JsonLineWorkbookWorkerClient : IWorkbookWorkerClient
{
    private readonly IWorkbookWorkerLaunchSpecProvider _launchSpecs;

    public JsonLineWorkbookWorkerClient()
        : this(new CurrentProcessWorkerLaunchSpecProvider())
    {
    }

    internal JsonLineWorkbookWorkerClient(IWorkbookWorkerLaunchSpecProvider launchSpecs) =>
        _launchSpecs = launchSpecs ?? throw new ArgumentNullException(nameof(launchSpecs));

    public async Task<IWorkbookWorkerSession> StartAsync(
        WorkbookWorkerOperation operation,
        WorkbookInspectionRequest? inspectionRequest,
        ExcelTaskPlan? plan,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var taskId = plan?.TaskId ?? Guid.NewGuid().ToString("N");
        var request = new WorkbookWorkerRequest(
            WorkbookWorkerProtocol.Version,
            taskId,
            operation == WorkbookWorkerOperation.Inspect ? "inspect" : "execute",
            inspectionRequest,
            plan);
        var requestJson = JsonSerializer.Serialize(request, WorkbookWorkerProtocol.JsonOptions);
        if (System.Text.Encoding.UTF8.GetByteCount(requestJson) > WorkbookWorkerProtocol.MaxRequestBytes)
        {
            throw new InvalidOperationException("Worker request is too large.");
        }

        var launch = _launchSpecs.GetLaunchSpec();
        if (string.IsNullOrWhiteSpace(launch.FileName) || !Path.IsPathFullyQualified(launch.FileName))
        {
            throw new InvalidOperationException("Worker executable is unavailable.");
        }

        var startInfo = new ProcessStartInfo
        {
            FileName = launch.FileName,
            UseShellExecute = false,
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };
        foreach (var argument in launch.PrefixArguments) startInfo.ArgumentList.Add(argument);
        startInfo.ArgumentList.Add("--excel-worker");
        var preExistingExcel = OwnedExcelProcess.SnapshotExcelProcesses()
            .Select(identity => new SupervisedExcelIdentity(identity.ProcessId, identity.StartTimeUtc, identity.ExecutablePath))
            .ToArray();
        var process = Process.Start(startInfo) ?? throw new InvalidOperationException("Worker process could not be started.");

        JsonLineWorkbookWorkerSession session;
        try
        {
            session = new JsonLineWorkbookWorkerSession(process, taskId, request.Operation, preExistingExcel);
        }
        catch
        {
            TryStopUninitializedWorker(process);
            process.Dispose();
            throw;
        }

        try
        {
            await process.StandardInput.WriteLineAsync(requestJson.AsMemory(), cancellationToken).ConfigureAwait(false);
            await process.StandardInput.FlushAsync(cancellationToken).ConfigureAwait(false);
            return session;
        }
        catch (Exception exception)
        {
            await session.StopAsync().ConfigureAwait(false);
            await session.DisposeAsync().ConfigureAwait(false);
            throw new WorkbookWorkerDispatchException(exception);
        }
    }

    private static void TryStopUninitializedWorker(Process process)
    {
        try
        {
            if (!process.HasExited)
            {
                process.Kill(entireProcessTree: true);
                _ = process.WaitForExit(5_000);
            }
        }
        catch (InvalidOperationException) { }
        catch (System.ComponentModel.Win32Exception) { }
    }
}

internal sealed class JsonLineWorkbookWorkerSession : IWorkbookWorkerSession
{
    private readonly Process _process;
    private readonly SupervisedExcelIdentity _identity;
    private readonly string _taskId;
    private readonly string _operation;
    private readonly SupervisedExcelIdentity[] _preExistingExcel;
    private readonly Task _stderrDrain;

    public JsonLineWorkbookWorkerSession(
        Process process,
        string taskId,
        string operation,
        SupervisedExcelIdentity[] preExistingExcel)
    {
        _process = process;
        _identity = new SupervisedExcelIdentity(process.Id, process.StartTime.ToUniversalTime(), process.MainModule?.FileName ?? string.Empty);
        _taskId = taskId;
        _operation = operation;
        _preExistingExcel = preExistingExcel ?? throw new ArgumentNullException(nameof(preExistingExcel));
        _stderrDrain = DrainBoundedAsync(process.StandardError);
    }

    public async Task<SupervisedWorkerFrame?> ReadAsync(CancellationToken cancellationToken)
    {
        var line = await ReadBoundedLineAsync(_process.StandardOutput, WorkbookWorkerProtocol.MaxFrameBytes, cancellationToken).ConfigureAwait(false);
        return line is null ? null : ParseFrame(line, _taskId, _operation, _preExistingExcel);
    }

    public async Task<bool> StopAsync()
    {
        var stopped = false;
        try
        {
            if (_process.HasExited)
            {
                stopped = true;
            }
            else if (_process.StartTime.ToUniversalTime() == _identity.StartTimeUtc &&
                     string.Equals(_process.MainModule?.FileName, _identity.ExecutablePath, StringComparison.OrdinalIgnoreCase))
            {
                _process.Kill(entireProcessTree: true);
                stopped = _process.WaitForExit(5_000);
            }
        }
        catch (InvalidOperationException) { }
        catch (System.ComponentModel.Win32Exception) { }
        await AwaitStderrDrainAsync().ConfigureAwait(false);
        if (!stopped)
        {
            try { stopped = _process.HasExited; }
            catch (InvalidOperationException) { }
        }
        return stopped;
    }

    public ValueTask DisposeAsync()
    {
        _process.Dispose();
        return ValueTask.CompletedTask;
    }

    private async Task AwaitStderrDrainAsync()
    {
        try { await _stderrDrain.WaitAsync(TimeSpan.FromSeconds(2)).ConfigureAwait(false); }
        catch (TimeoutException) { }
        catch { }
    }

    private static async Task DrainBoundedAsync(StreamReader reader)
    {
        var buffer = new char[256];
        while (true)
        {
            var read = await reader.ReadAsync(buffer.AsMemory()).ConfigureAwait(false);
            if (read == 0) return;
        }
    }

    private static async Task<string?> ReadBoundedLineAsync(StreamReader reader, int maxBytes, CancellationToken cancellationToken)
    {
        var buffer = new System.Text.StringBuilder(Math.Min(1024, maxBytes));
        var character = new char[1];
        while (true)
        {
            var read = await reader.ReadAsync(character.AsMemory(), cancellationToken).ConfigureAwait(false);
            if (read == 0) return buffer.Length == 0 ? null : buffer.ToString();
            if (character[0] == '\n') break;
            if (character[0] == '\r') continue;
            buffer.Append(character[0]);
            if (buffer.Length > maxBytes || System.Text.Encoding.UTF8.GetByteCount(buffer.ToString()) > maxBytes)
            {
                throw new InvalidDataException("Worker frame is too large.");
            }
        }

        return buffer.ToString();
    }

    private static SupervisedWorkerFrame ParseFrame(
        string line,
        string taskId,
        string operation,
        SupervisedExcelIdentity[] preExistingExcel)
    {
        try
        {
            using var document = JsonDocument.Parse(line, new JsonDocumentOptions { MaxDepth = 16 });
            var root = document.RootElement;
            if (root.ValueKind != JsonValueKind.Object ||
                !root.TryGetProperty("version", out var version) || version.ValueKind != JsonValueKind.Number || version.GetInt32() != WorkbookWorkerProtocol.Version ||
                !root.TryGetProperty("type", out var typeElement) || typeElement.ValueKind != JsonValueKind.String)
            {
                return new SupervisedWorkerFrame(Fatal: true);
            }

            var type = typeElement.GetString();
            if (type == "fatal") return new SupervisedWorkerFrame(Fatal: true);
            if (!MatchesRequest(root, taskId, operation)) return new SupervisedWorkerFrame(Fatal: true);

            return type switch
            {
                "accepted" => new SupervisedWorkerFrame(Accepted: true),
                // Bounded at the same length the frame is later checked against. Parsing a longer
                // phase and then classifying the frame malformed would turn a healthy run into
                // Unknown over a label nobody reads.
                "phase" when TryGetBoundedText(root, "phase", WorkbookWorkerProtocol.MaxPhaseLength, out var phase) => new SupervisedWorkerFrame(Phase: phase),
                "owned-process" when TryGetExcelIdentity(root, out var identity) && IsAuthorizedExcelIdentity(identity!, preExistingExcel) => new SupervisedWorkerFrame(ExcelIdentity: identity),
                "artifact-staged" when TryGetBoundedPath(root, "stagingPath", out var stagingPath) => new SupervisedWorkerFrame(StagingArtifactPath: stagingPath),
                "result" when TryGetResult(root, operation, out var inspection, out var execution) => new SupervisedWorkerFrame(Inspection: inspection, Execution: execution),
                _ => new SupervisedWorkerFrame(Fatal: true)
            };
        }
        catch (JsonException) { return new SupervisedWorkerFrame(Fatal: true); }
        catch (NotSupportedException) { return new SupervisedWorkerFrame(Fatal: true); }
        catch (FormatException) { return new SupervisedWorkerFrame(Fatal: true); }
        catch (InvalidOperationException) { return new SupervisedWorkerFrame(Fatal: true); }
        catch (OverflowException) { return new SupervisedWorkerFrame(Fatal: true); }
    }

    private static bool MatchesRequest(JsonElement root, string taskId, string operation)
    {
        if (!root.TryGetProperty("taskId", out var frameTaskId) || frameTaskId.ValueKind != JsonValueKind.String ||
            !string.Equals(frameTaskId.GetString(), taskId, StringComparison.Ordinal)) return false;
        return !root.TryGetProperty("operation", out var frameOperation) ||
               (frameOperation.ValueKind == JsonValueKind.String && string.Equals(frameOperation.GetString(), operation, StringComparison.Ordinal));
    }

    private static bool TryGetBoundedText(JsonElement root, string property, out string? value) =>
        TryGetBoundedText(root, property, WorkbookWorkerProtocol.MaxTextLength, out value);

    private static bool TryGetBoundedText(JsonElement root, string property, int maximumLength, out string? value)
    {
        value = null;
        if (!root.TryGetProperty(property, out var text) || text.ValueKind != JsonValueKind.String) return false;
        value = text.GetString();
        return value is { Length: > 0 } && value.Length <= maximumLength;
    }

    private static bool TryGetBoundedPath(JsonElement root, string property, out string? value)
    {
        value = null;
        if (!root.TryGetProperty(property, out var path) || path.ValueKind != JsonValueKind.String) return false;
        var raw = path.GetString();
        if (raw is not { Length: > 0 and <= 4096 } || !Path.IsPathFullyQualified(raw)) return false;
        try
        {
            value = Path.GetFullPath(raw);
            return true;
        }
        catch (ArgumentException) { return false; }
        catch (NotSupportedException) { return false; }
    }

    private static bool TryGetExcelIdentity(JsonElement root, out SupervisedExcelIdentity? identity)
    {
        identity = null;
        if (!root.TryGetProperty("processId", out var processId) || !processId.TryGetInt32(out var id) ||
            !root.TryGetProperty("startTimeUtc", out var startTime) || startTime.ValueKind != JsonValueKind.String || !DateTime.TryParse(startTime.GetString(), out var start) ||
            !TryGetBoundedPath(root, "executablePath", out var path)) return false;
        identity = new SupervisedExcelIdentity(id, start.ToUniversalTime(), path!);
        return identity.IsExcelProcess();
    }

    internal static bool IsAuthorizedExcelIdentity(
        SupervisedExcelIdentity identity,
        IReadOnlyList<SupervisedExcelIdentity> preExistingExcel) => identity.IsExcelProcess() &&
        !preExistingExcel.Any(existing => existing.MatchesExactly(identity));

    private static bool TryGetResult(JsonElement root, string operation, out WorkbookInspection? inspection, out WorkbookExecutionOutcome? execution)
    {
        inspection = null;
        execution = null;
        if (!root.TryGetProperty("result", out var result)) return false;
        try
        {
            if (operation == "inspect") inspection = JsonSerializer.Deserialize<WorkbookInspection>(result.GetRawText(), WorkbookWorkerProtocol.JsonOptions);
            else execution = JsonSerializer.Deserialize<WorkbookExecutionOutcome>(result.GetRawText(), WorkbookWorkerProtocol.JsonOptions);
            return inspection is not null || execution is not null;
        }
        catch (JsonException) { return false; }
        catch (NotSupportedException) { return false; }
        catch (FormatException) { return false; }
        catch (InvalidOperationException) { return false; }
        catch (OverflowException) { return false; }
    }
}

internal sealed record WorkbookWorkerLaunchSpec(string FileName, IReadOnlyList<string> PrefixArguments);

internal interface IWorkbookWorkerLaunchSpecProvider
{
    WorkbookWorkerLaunchSpec GetLaunchSpec();
}

internal sealed class CurrentProcessWorkerLaunchSpecProvider : IWorkbookWorkerLaunchSpecProvider
{
    private readonly Func<string?> _processPath;
    private readonly Func<string?> _entryAssemblyPath;

    public CurrentProcessWorkerLaunchSpecProvider()
        : this(() => Environment.ProcessPath, ResolveEntryAssemblyPath)
    {
    }

    internal CurrentProcessWorkerLaunchSpecProvider(Func<string?> processPath, Func<string?> entryAssemblyPath)
    {
        _processPath = processPath ?? throw new ArgumentNullException(nameof(processPath));
        _entryAssemblyPath = entryAssemblyPath ?? throw new ArgumentNullException(nameof(entryAssemblyPath));
    }

    public WorkbookWorkerLaunchSpec GetLaunchSpec()
    {
        var executable = _processPath();
        var entryAssembly = _entryAssemblyPath();
        if (string.IsNullOrWhiteSpace(executable) || !Path.IsPathFullyQualified(executable))
        {
            throw new InvalidOperationException("Worker executable is unavailable.");
        }

        return IsDotnetHost(executable) && entryAssembly is { Length: > 0 } && Path.IsPathFullyQualified(entryAssembly) && string.Equals(Path.GetExtension(entryAssembly), ".dll", StringComparison.OrdinalIgnoreCase)
            ? new WorkbookWorkerLaunchSpec(executable, [entryAssembly])
            : new WorkbookWorkerLaunchSpec(executable, []);
    }

    private static bool IsDotnetHost(string executable) =>
        string.Equals(Path.GetFileNameWithoutExtension(executable), "dotnet", StringComparison.OrdinalIgnoreCase);

    private static string? ResolveEntryAssemblyPath()
    {
        var assemblyName = System.Reflection.Assembly.GetEntryAssembly()?.GetName().Name;
        return string.IsNullOrWhiteSpace(assemblyName)
            ? null
            : Path.Combine(AppContext.BaseDirectory, assemblyName + ".dll");
    }
}
