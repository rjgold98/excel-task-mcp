using System.Diagnostics;
using System.Globalization;
using System.Text;

namespace ExcelTask.Excel;

/// <summary>
/// A development-time narration of what the server did, and when.
///
/// TEMPORARY, and built to be deleted: it is off unless EXCELTASK_TRACE names a file, it is one
/// file plus a handful of call sites, and nothing depends on it. Deleting this file and the lines
/// that call <see cref="Begin"/> removes the feature entirely.
///
/// It exists because the work computer is where the real workbooks are and the home computer is
/// where the fixes happen, so a failure there has to travel here as evidence. The supervisor drains
/// the worker's stderr and discards it, receipts are deliberately bounded, and phases were reported
/// to an observer that threw them away - which left "it did not work" as the whole bug report.
///
/// WHAT IT RECORDS:
///   - phase names and how long each took
///   - operation kind, mode, binding, save mode
///   - worksheet names and A1 ranges (needed to reproduce anything)
///   - workbook FILE NAMES only, never directories
///   - owned Excel process ids and whether they exited
///   - status, check names, and check details
/// WHAT IT NEVER RECORDS:
///   - cell values, formulas, or any workbook contents
///   - VBA source
///   - connection strings, server names, queries
///   - full paths, user names, or machine names
/// Every file it writes opens with that same statement, so the person sending it can see what they
/// are sending before they send it.
///
/// What the header no longer says is "safe to share". It said that for eight versions while the
/// first list plainly includes workbook and worksheet names - and a healthcare finance workbook is
/// routinely called something like a payer, a facility, or a deal. The file cannot know whether
/// naming those is acceptable where it is going, so it states its contents and leaves the
/// conclusion to the person holding it. This is the same rule the receipts follow: report the
/// evidence, never assert the verdict the evidence does not reach.
/// </summary>
public sealed class DiagnosticTrace
{
    private const string EnvironmentVariable = "EXCELTASK_TRACE";

    private static readonly string[] Header =
    [
        "ExcelTask diagnostic trace - DEVELOPMENT ONLY. Read both lists before sharing this file.",
        "Records: phases and durations, operation kind/mode/binding/save, worksheet names, A1 ranges,",
        "         workbook file names (never directories), owned Excel process ids, statuses and checks.",
        "Never records: cell values, formulas, workbook contents, VBA source, connection strings,",
        "         server names, full paths, user names, or machine names.",
        "A workbook or worksheet name can name a client, a payer, a deal or a facility by itself.",
        "Whether this file may leave your machine is your judgement, not this file's to assert.",
        "Turn it off by clearing the EXCELTASK_TRACE environment variable.",
        ""
    ];

    private readonly string _path;
    private readonly string _taskId;
    private readonly Stopwatch _sinceStart = Stopwatch.StartNew();
    private readonly Lock _gate = new();
    private string? _openPhase;
    private long _phaseStartedMs;

    private DiagnosticTrace(string path, string taskId)
    {
        _path = path;
        _taskId = taskId;
    }

    /// <summary>The trace for this run, or null when tracing is off - which is the default.</summary>
    public static DiagnosticTrace? Begin(string taskId, string role)
    {
        var path = Environment.GetEnvironmentVariable(EnvironmentVariable);
        if (string.IsNullOrWhiteSpace(path)) return null;

        try
        {
            var directory = Path.GetDirectoryName(Path.GetFullPath(path));
            if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
            var trace = new DiagnosticTrace(Path.GetFullPath(path), taskId);
            if (!File.Exists(trace._path)) trace.WriteLines(Header);
            trace.Write($"=== {role} start ===");
            return trace;
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or ArgumentException or NotSupportedException)
        {
            // Tracing must never be able to break a task. A trace that cannot be written is simply
            // not written.
            return null;
        }
    }

    /// <summary>Closes the previous phase with its duration and opens the next.</summary>
    public void Phase(string phase)
    {
        lock (_gate)
        {
            if (_openPhase is not null)
            {
                Write($"  phase end   {_openPhase} ({_sinceStart.ElapsedMilliseconds - _phaseStartedMs} ms)");
            }

            _openPhase = phase;
            _phaseStartedMs = _sinceStart.ElapsedMilliseconds;
            Write($"  phase start {phase}");
        }
    }

    public void Note(string message) => Write($"  {message}");

    /// <summary>Closes any open phase and stamps the total, so a truncated file is still readable.</summary>
    public void End(string role, string summary)
    {
        lock (_gate)
        {
            if (_openPhase is not null)
            {
                Write($"  phase end   {_openPhase} ({_sinceStart.ElapsedMilliseconds - _phaseStartedMs} ms)");
                _openPhase = null;
            }

            Write($"=== {role} end: {summary} ({_sinceStart.ElapsedMilliseconds} ms total) ===");
            Write(string.Empty);
        }
    }

    /// <summary>
    /// A file name with no directory, because a full path names the machine and the person.
    /// One rule, owned by <see cref="WorkbookRuntimeHelpers"/> - the receipts redact with the same
    /// call, and a guarantee they depend on must not live in a module built to be deleted.
    /// </summary>
    public static string FileNameOnly(string? path) => WorkbookRuntimeHelpers.FileNameOnly(path);

    private void Write(string line) => WriteLines([
        line.Length == 0
            ? string.Empty
            : $"{DateTime.UtcNow.ToString("HH:mm:ss.fff", CultureInfo.InvariantCulture)} [{_taskId[..Math.Min(8, _taskId.Length)]}] {line}"
    ]);

    private void WriteLines(IReadOnlyList<string> lines)
    {
        var payload = new StringBuilder();
        foreach (var line in lines) payload.AppendLine(line);

        // Appended with sharing on, because the supervisor and its worker are two processes writing
        // one file. A brief contention retry is enough; losing a trace line must never fail a task.
        for (var attempt = 0; attempt < 5; attempt++)
        {
            try
            {
                using var stream = new FileStream(_path, FileMode.Append, FileAccess.Write, FileShare.ReadWrite);
                using var writer = new StreamWriter(stream, Encoding.UTF8);
                writer.Write(payload.ToString());
                return;
            }
            catch (IOException) { Thread.Sleep(20); }
            catch (UnauthorizedAccessException) { return; }
        }
    }
}
