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
        var trace = DiagnosticTrace.Begin(parsedRequest.TaskId, "worker");
        DescribeRequest(trace, parsedRequest);
        var observer = new ProtocolObserver(writer, parsedRequest.TaskId, watchdog, trace);
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

            DescribeResult(trace, result);
            await writer.WriteAsync(new { version = WorkbookWorkerProtocol.Version, type = "result", taskId = parsedRequest.TaskId, operation = parsedRequest.Operation, result }).ConfigureAwait(false);
            return 0;
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            trace?.End("worker", "cancelled by the host deadline");
            await writer.WriteFatalAsync("cancelled", parsedRequest.TaskId).ConfigureAwait(false);
            return 3;
        }
        catch (Exception exception)
        {
            // The one place the worker's own fault is legible. The supervisor drains stderr and
            // discards it, so without this a worker crash reached the caller as "Unknown" and
            // nothing else - which is what made a work-computer failure impossible to diagnose here.
            trace?.Note($"UNHANDLED {exception.GetType().Name}: {exception.Message.ReplaceLineEndings(" ")}");
            trace?.End("worker", "unhandled worker failure");
            await writer.WriteFatalAsync("worker-failure", parsedRequest.TaskId).ConfigureAwait(false);
            return 1;
        }
    }

    /// <summary>The request's shape - kind, mode, policy, sheets and ranges. Never its contents.</summary>
    private static void DescribeRequest(DiagnosticTrace? trace, WorkbookWorkerRequest request)
    {
        if (trace is null) return;

        if (request.Inspection is not null)
        {
            trace.Note($"inspect: target={DiagnosticTrace.FileNameOnly(request.Inspection.TargetWorkbookPath)} " +
                       $"binding={request.Inspection.Binding} save={request.Inspection.Save} mustExist={request.Inspection.TargetMustExist}");
            return;
        }

        var plan = request.Plan!.Request;
        var operation = plan.Operation;
        trace.Note($"execute: {operation.Kind} mode={plan.Mode} binding={plan.WorkbookBinding} save={plan.Save} " +
                   $"overwriteConfirmed={plan.OverwriteConfirmed} target={DiagnosticTrace.FileNameOnly(plan.TargetWorkbookPath)} " +
                   $"output={DiagnosticTrace.FileNameOnly(plan.OutputWorkbookPath)}");

        var detail = operation.Kind switch
        {
            ExcelOperationKind.ReadWorksheetRange => $"sheet={operation.ReadWorksheetRange!.WorksheetName} range={operation.ReadWorksheetRange.Range} formulas={operation.ReadWorksheetRange.Formulas}",
            ExcelOperationKind.WriteWorksheetValues => $"sheet={operation.WriteWorksheetValues!.WorksheetName} cells={operation.WriteWorksheetValues.Cells.Count}",
            ExcelOperationKind.FindReplace => $"sheet={operation.FindReplace!.WorksheetName} range={operation.FindReplace.Range?.ToString() ?? "(used range)"} wholeCell={operation.FindReplace.WholeCell} matchCase={operation.FindReplace.MatchCase}",
            ExcelOperationKind.SetNumberFormat => $"sheet={operation.SetNumberFormat!.WorksheetName} range={operation.SetNumberFormat.Range}",
            ExcelOperationKind.Create => $"createKind={operation.Create!.Kind} sheet={operation.Create.WorksheetName ?? "(default)"}",
            ExcelOperationKind.RepairExistingWorksheet => $"sheet={operation.RepairExistingWorksheet!.WorksheetName} ranges={operation.RepairExistingWorksheet.Ranges.Count}",
            ExcelOperationKind.ExtendFormulaSeries => $"sheet={operation.ExtendFormulaSeries!.WorksheetName} direction={operation.ExtendFormulaSeries.Direction} evidence={operation.ExtendFormulaSeries.EvidenceRange} destination={operation.ExtendFormulaSeries.DestinationRange}",
            ExcelOperationKind.CopyExhibit => $"referenceSheet={operation.CopyExhibit!.ReferenceWorksheet} newSheet={operation.CopyExhibit.NewWorksheetName} repairRanges={operation.CopyExhibit.RepairRanges.Count}",
            ExcelOperationKind.EditMacroProcedure => $"component={operation.EditMacroProcedure!.ComponentName} procedure={operation.EditMacroProcedure.ProcedureName} run={operation.EditMacroProcedure.RunAfterEdit}",
            _ => "no options"
        };
        trace.Note($"  {detail}");
    }

    /// <summary>Status and every check, which is where a failure explains itself.</summary>
    private static void DescribeResult(DiagnosticTrace? trace, object result)
    {
        if (trace is null) return;

        if (result is WorkbookExecutionOutcome outcome)
        {
            foreach (var check in outcome.Checks ?? [])
            {
                trace.Note($"check {(check.Passed ? "PASS" : "FAIL")} {check.Name}: {check.Detail}");
            }

            trace.End("worker", $"{outcome.Status}: {outcome.Summary}");
            return;
        }

        if (result is WorkbookInspection inspection)
        {
            foreach (var check in inspection.Checks ?? [])
            {
                trace.Note($"check {(check.Passed ? "PASS" : "FAIL")} {check.Name}: {check.Detail}");
            }

            trace.End("worker", inspection.InfeasibleReason ?? $"inspected, targetIsOpen={inspection.TargetIsOpen}");
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
        WorkerOwnedProcessWatchdog watchdog,
        DiagnosticTrace? trace) : IExcelWorkbookRuntimeObserver
    {
        public void OnPhase(string phase)
        {
            trace?.Phase(phase);
            writer.Write(new { version = WorkbookWorkerProtocol.Version, type = "phase", taskId, phase });
        }

        public void OnOwnedProcessCaptured(ProcessIdentity identity)
        {
            watchdog.Track(identity);
            trace?.Note($"owned Excel started, pid {identity.ProcessId}");
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

        public void OnStagingPathCreated(string stagingPath)
        {
            trace?.Note($"staging file created: {DiagnosticTrace.FileNameOnly(stagingPath)}");
            writer.Write(new
            {
                version = WorkbookWorkerProtocol.Version,
                type = "artifact-staged",
                taskId,
                stagingPath
            });
        }
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
