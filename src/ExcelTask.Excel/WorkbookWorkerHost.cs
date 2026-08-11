using System.Text;
using System.Text.Json;
using System.Diagnostics;
using ExcelTask.Core;

namespace ExcelTask.Excel;

/// <summary>Runs one private worker request over line-delimited JSON streams.</summary>
public static class WorkbookWorkerHost
{
    private static readonly TimeSpan OwnedProcessRecoveryDeadline = TimeSpan.FromSeconds(110);

    public static Task<int> RunAsync(TextReader input, TextWriter output, CancellationToken cancellationToken = default) =>
        RunAsync(input, output, observer => new ExcelWorkbookRuntime(observer), cancellationToken);

    internal static async Task<int> RunAsync(
        TextReader input,
        TextWriter output,
        Func<IExcelWorkbookRuntimeObserver, IWorkbookRuntime> runtimeFactory,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(input);
        ArgumentNullException.ThrowIfNull(output);
        ArgumentNullException.ThrowIfNull(runtimeFactory);

        var writer = new WorkerFrameWriter(output);
        var read = await ReadBoundedLineAsync(input, cancellationToken).ConfigureAwait(false);
        if (read.TooLarge)
        {
            await writer.WriteFatalAsync("input-too-large").ConfigureAwait(false);
            return 2;
        }

        if (read.Line is null || !WorkbookWorkerProtocol.TryParseRequest(read.Line, out var request))
        {
            await writer.WriteFatalAsync("invalid-request").ConfigureAwait(false);
            return 2;
        }

        var parsedRequest = request!;
        await using var watchdog = new WorkerOwnedProcessWatchdog(
            OwnedProcessRecoveryDeadline,
            new WorkerOwnedProcessTerminator(),
            () => writer.Write(new { version = WorkbookWorkerProtocol.Version, type = "phase", taskId = parsedRequest.TaskId, phase = "owned-process-recovery" }));
        var observer = new ProtocolObserver(writer, parsedRequest.TaskId, watchdog);
        try
        {
            await writer.WriteAsync(new { version = WorkbookWorkerProtocol.Version, type = "accepted", taskId = parsedRequest.TaskId, operation = parsedRequest.Operation }).ConfigureAwait(false);
            observer.OnPhase("runtime-dispatch");

            object result;
            var runtime = runtimeFactory(observer);
            if (runtime is null) throw new InvalidOperationException();
            try
            {
                result = parsedRequest.Inspection is not null
                    ? WorkbookWorkerProtocol.Bound(await runtime.InspectAsync(parsedRequest.Inspection, cancellationToken).ConfigureAwait(false))
                    : WorkbookWorkerProtocol.Bound(await runtime.ExecuteAsync(parsedRequest.Plan!, cancellationToken).ConfigureAwait(false));
            }
            finally
            {
                (runtime as IDisposable)?.Dispose();
            }

            await writer.WriteAsync(new { version = WorkbookWorkerProtocol.Version, type = "result", taskId = parsedRequest.TaskId, operation = parsedRequest.Operation, result }).ConfigureAwait(false);
            return 0;
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            await writer.WriteFatalAsync("cancelled", parsedRequest.TaskId).ConfigureAwait(false);
            return 3;
        }
        catch (Exception)
        {
            await writer.WriteFatalAsync("worker-failure", parsedRequest.TaskId).ConfigureAwait(false);
            return 1;
        }
    }

    private static async Task<(string? Line, bool TooLarge)> ReadBoundedLineAsync(TextReader input, CancellationToken cancellationToken)
    {
        var buffer = new StringBuilder(Math.Min(1024, WorkbookWorkerProtocol.MaxRequestBytes));
        var character = new char[1];
        var byteCount = 0;
        while (true)
        {
            var count = await input.ReadAsync(character.AsMemory(), cancellationToken).ConfigureAwait(false);
            if (count == 0) return (buffer.Length == 0 ? null : buffer.ToString(), false);
            if (character[0] == '\n') return (buffer.ToString(), false);
            if (character[0] == '\r') continue;
            byteCount += Encoding.UTF8.GetByteCount(character);
            if (byteCount > WorkbookWorkerProtocol.MaxRequestBytes) return (null, true);
            buffer.Append(character[0]);
        }
    }

    private sealed class ProtocolObserver(
        WorkerFrameWriter writer,
        string taskId,
        WorkerOwnedProcessWatchdog watchdog) : IExcelWorkbookRuntimeObserver
    {
        public void OnPhase(string phase) => writer.Write(new { version = WorkbookWorkerProtocol.Version, type = "phase", taskId, phase });

        public void OnOwnedProcessCaptured(ProcessIdentity identity)
        {
            watchdog.Track(identity);
            writer.Write(new
            {
                version = WorkbookWorkerProtocol.Version,
                type = "owned-process",
                taskId,
                processId = identity.ProcessId,
                startTimeUtc = identity.StartTimeUtc,
                executablePath = identity.ExecutablePath
            });
        }

        public void OnStagingPathCreated(string stagingPath) => writer.Write(new
        {
            version = WorkbookWorkerProtocol.Version,
            type = "artifact-staged",
            taskId,
            stagingPath
        });
    }

    private sealed class WorkerFrameWriter(TextWriter output)
    {
        private readonly object _gate = new();

        public Task WriteAsync(object frame)
        {
            Write(frame);
            return Task.CompletedTask;
        }

        public void Write(object frame)
        {
            lock (_gate)
            {
                WriteCore(frame);
            }
        }

        public Task WriteFatalAsync(string code, string? taskId = null) => WriteAsync(new
        {
            version = WorkbookWorkerProtocol.Version,
            type = "fatal",
            taskId,
            code
        });

        private void WriteCore(object frame)
        {
            var json = JsonSerializer.Serialize(frame, WorkbookWorkerProtocol.JsonOptions);
            if (Encoding.UTF8.GetByteCount(json) > WorkbookWorkerProtocol.MaxFrameBytes)
            {
                json = JsonSerializer.Serialize(new { version = WorkbookWorkerProtocol.Version, type = "fatal", code = "frame-too-large" }, WorkbookWorkerProtocol.JsonOptions);
            }

            output.WriteLine(json);
            output.Flush();
        }
    }
}

internal interface IWorkerOwnedProcessTerminator
{
    bool TryTerminateExact(ProcessIdentity identity);
}

internal sealed class WorkerOwnedProcessTerminator : IWorkerOwnedProcessTerminator
{
    public bool TryTerminateExact(ProcessIdentity identity)
    {
        if (!string.Equals(Path.GetFileName(identity.ExecutablePath), "EXCEL.EXE", StringComparison.OrdinalIgnoreCase)) return false;
        if (!ProcessIdentity.TryOpenMatching(identity, out var process)) return true;
        using (process)
        {
            try
            {
                if (process.HasExited) return true;
                process.Kill(entireProcessTree: false);
                return process.WaitForExit(5_000);
            }
            catch (InvalidOperationException) { return true; }
            catch (System.ComponentModel.Win32Exception) { return false; }
        }
    }
}

internal sealed class WorkerOwnedProcessWatchdog : IAsyncDisposable
{
    private readonly object _gate = new();
    private readonly IWorkerOwnedProcessTerminator _terminator;
    private readonly Action _onRecovery;
    private readonly CancellationTokenSource _stop = new();
    private readonly Task _deadlineTask;

    /// <summary>
    /// Every owned Excel, not the most recent one. This was a single identity while only one owned
    /// process could exist at a time. Once the verification instance starts before the primary
    /// closes, a single field would be overwritten by the second launch - and on a deadline the
    /// watchdog would kill the idle verification Excel and leave the one holding the user's
    /// workbook mid-write running, which is the exact opposite of what it exists to do.
    /// </summary>
    private readonly HashSet<ProcessIdentity> _identities = [];
    private bool _deadlineElapsed;
    private bool _stopped;

    public WorkerOwnedProcessWatchdog(
        TimeSpan deadline,
        IWorkerOwnedProcessTerminator terminator,
        Action onRecovery)
    {
        ArgumentOutOfRangeException.ThrowIfLessThanOrEqual(deadline, TimeSpan.Zero);
        _terminator = terminator ?? throw new ArgumentNullException(nameof(terminator));
        _onRecovery = onRecovery ?? throw new ArgumentNullException(nameof(onRecovery));
        _deadlineTask = WaitForDeadlineAsync(deadline);
    }

    public void Track(ProcessIdentity identity)
    {
        ArgumentNullException.ThrowIfNull(identity);
        var recoverNow = false;
        lock (_gate)
        {
            if (_stopped) return;
            _identities.Add(identity);
            recoverNow = _deadlineElapsed;
        }

        // A process registered after the deadline has already passed is terminated immediately:
        // the run is over, and nothing owned may outlive it.
        if (recoverNow) Recover([identity]);
    }

    public async ValueTask DisposeAsync()
    {
        lock (_gate) _stopped = true;
        _stop.Cancel();
        try { await _deadlineTask.ConfigureAwait(false); }
        catch (OperationCanceledException) { }
        _stop.Dispose();
    }

    private async Task WaitForDeadlineAsync(TimeSpan deadline)
    {
        await Task.Delay(deadline, _stop.Token).ConfigureAwait(false);
        ProcessIdentity[] identities;
        lock (_gate)
        {
            if (_stopped) return;
            _deadlineElapsed = true;
            identities = [.. _identities];
        }

        if (identities.Length > 0) Recover(identities);
    }

    /// <summary>
    /// Recovery runs once and then terminates every owned process. One failing to die must not
    /// stop the others being tried, so each is attempted independently.
    /// </summary>
    private void Recover(IReadOnlyList<ProcessIdentity> identities)
    {
        try { _onRecovery(); }
        catch { }
        foreach (var identity in identities)
        {
            try { _ = _terminator.TryTerminateExact(identity); }
            catch { }
        }
    }
}
