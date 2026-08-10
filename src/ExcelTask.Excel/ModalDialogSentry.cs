using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Text;

namespace ExcelTask.Excel;

/// <summary>What kind of blocking dialog was found, which decides how it is answered.</summary>
internal enum ModalDialogKind
{
    VbaCompileError,
    VbaRuntimeError,
    MacroMessageBox
}

internal sealed record DismissedDialog(ModalDialogKind Kind, string Message);

/// <summary>
/// Answers the modal dialogs that VBA raises inside one owned Excel process.
///
/// A compile error, an unhandled run-time error, or a MsgBox in macro code opens a modal window on
/// Excel's own thread. The automation call is already blocked inside COM at that point, so nothing
/// on the calling thread can react; without help the operation stalls until the watchdog kills
/// Excel and the caller receives an uncertain result carrying no reason. This sentry watches from a
/// separate thread, answers only dialogs it positively recognizes, and records what it answered so
/// the failure can be reported precisely.
///
/// It is deliberately narrow. It only ever looks at the single process identity it was given, which
/// is always an Excel instance ExcelTask created and verified by process id, start time, and image
/// path. A workbook bound with <see cref="Core.WorkbookBinding.UseOpen"/> has no owned process, so
/// no sentry is started and a dialog in the user's own Excel is never touched. Any dialog whose
/// control layout is not recognized is left alone.
/// </summary>
[SupportedOSPlatform("windows")]
internal sealed class ModalDialogSentry : IDisposable
{
    private const string DialogClassName = "#32770";
    private const int MaxMessageLength = 200;

    // Control identifiers are stable across Office language packs, unlike button captions.
    private const int IdOk = 2;
    private const int IdHelp = 9;
    private const int IdCompileErrorDecoration = 20;
    private const int IdStaticText = 65535;
    private const int IdRuntimeEnd = 4800;
    private const int IdRuntimeDebug = 4801;
    private const int IdRuntimeContinue = 4802;
    private const int IdRuntimeText = 4803;

    private static readonly TimeSpan PollInterval = TimeSpan.FromMilliseconds(500);

    private readonly ProcessIdentity _identity;
    private readonly CancellationTokenSource _stop = new();
    private readonly Task _loop;
    private readonly Lock _gate = new();
    private readonly List<DismissedDialog> _dismissed = [];

    private ModalDialogSentry(ProcessIdentity identity)
    {
        _identity = identity;
        _loop = Task.Run(WatchAsync);
    }

    /// <summary>Starts watching an owned Excel process, or returns null when there is nothing owned to watch.</summary>
    public static ModalDialogSentry? Watch(ProcessIdentity? identity) =>
        identity is null ? null : new ModalDialogSentry(identity);

    /// <summary>Dialogs answered during the watch, in the order they were answered.</summary>
    public IReadOnlyList<DismissedDialog> Dismissed
    {
        get { lock (_gate) return _dismissed.ToArray(); }
    }

    public void Dispose()
    {
        _stop.Cancel();
        try { _loop.Wait(TimeSpan.FromSeconds(2)); }
        catch (AggregateException) { }
        _stop.Dispose();
    }

    private async Task WatchAsync()
    {
        while (!_stop.IsCancellationRequested)
        {
            try { await Task.Delay(PollInterval, _stop.Token).ConfigureAwait(false); }
            catch (OperationCanceledException) { return; }

            try { InspectOnce(); }
            catch (Exception exception) when (exception is Win32Exception or InvalidOperationException or ArgumentException)
            {
                // The process may exit at any moment; that is not a sentry failure.
            }
        }
    }

    private void InspectOnce()
    {
        // Re-verify identity every pass. A process id alone can be recycled, and answering a dialog
        // in some other process would be exactly the mistake this class exists to avoid.
        if (!ProcessIdentity.TryOpenMatching(_identity, out var process)) return;
        using (process)
        {
            if (process.HasExited) return;
        }

        foreach (var dialog in EnumerateDialogs((uint)_identity.ProcessId))
        {
            if (!TryClassify(dialog, out var kind, out var message, out var buttonToClick)) continue;

            lock (_gate) _dismissed.Add(new DismissedDialog(kind, message));

            if (kind == ModalDialogKind.VbaCompileError)
            {
                // Measured behaviour: answering a compile error does not hand control back. VBA
                // drops into break mode with the module open in the editor, and the automation call
                // that is already blocked inside COM never returns. Ending this instance is the only
                // way to release it. That is safe here because a compile error happens before the
                // macro runs and before anything is saved, so the workbook on disk is untouched.
                TerminateOwnedProcess();
                continue;
            }

            // A timeout keeps a wedged UI thread from blocking the sentry itself, which would
            // defeat the whole point of watching from outside the blocked call.
            _ = SendMessageTimeout(buttonToClick, BM_CLICK, IntPtr.Zero, IntPtr.Zero, SMTO_ABORTIFHUNG, 2000, out _);
        }
    }

    /// <summary>
    /// Ends the one Excel instance this sentry was given. The identity is re-verified by process id,
    /// start time, and image path first, so a recycled id cannot redirect this at another process,
    /// and a session bound to a workbook the user already had open never has an identity to begin with.
    /// </summary>
    private void TerminateOwnedProcess()
    {
        if (!ProcessIdentity.TryOpenMatching(_identity, out var process)) return;
        using (process)
        {
            try
            {
                process.Kill(entireProcessTree: false);
                _ = process.WaitForExit(5_000);
            }
            catch (Exception exception) when (exception is Win32Exception or InvalidOperationException or NotSupportedException)
            {
                // Already gone, or not ours to end. Either way the blocked call will surface it.
            }
        }
    }

    private static List<IntPtr> EnumerateDialogs(uint processId)
    {
        var dialogs = new List<IntPtr>();
        EnumWindows((handle, state) =>
        {
            _ = GetWindowThreadProcessId(handle, out var owner);
            if (owner == processId && IsWindowVisible(handle) && ReadClassName(handle) == DialogClassName)
            {
                dialogs.Add(handle);
            }

            return true;
        }, IntPtr.Zero);
        return dialogs;
    }

    /// <summary>
    /// Recognizes a dialog by the set of control identifiers it contains. Nothing is answered on a
    /// caption or button caption, so an Office language pack cannot cause the wrong button press.
    /// </summary>
    private static bool TryClassify(IntPtr dialog, out ModalDialogKind kind, out string message, out IntPtr buttonToClick)
    {
        kind = default;
        message = string.Empty;
        buttonToClick = IntPtr.Zero;

        var controls = new Dictionary<int, IntPtr>();
        EnumChildWindows(dialog, (child, state) =>
        {
            controls.TryAdd(GetDlgCtrlID(child), child);
            return true;
        }, IntPtr.Zero);

        // The VBA run-time error dialog. End stops the macro; Debug would leave Excel sitting in the
        // VBE with no way back, so it is never pressed.
        if (controls.TryGetValue(IdRuntimeEnd, out var endButton) &&
            controls.ContainsKey(IdRuntimeDebug) &&
            controls.ContainsKey(IdRuntimeContinue))
        {
            kind = ModalDialogKind.VbaRuntimeError;
            message = ReadControlText(controls, IdRuntimeText);
            buttonToClick = endButton;
            return true;
        }

        if (!controls.TryGetValue(IdOk, out var okButton)) return false;

        // A compile error carries a Help button and a decoration control alongside OK.
        if (controls.ContainsKey(IdHelp) && controls.ContainsKey(IdCompileErrorDecoration))
        {
            kind = ModalDialogKind.VbaCompileError;
            message = ReadControlText(controls, IdStaticText);
            buttonToClick = okButton;
            return true;
        }

        // A MsgBox offering only OK. Anything offering a real choice - Yes/No, OK/Cancel, Retry -
        // is left alone, because picking an answer on the user's behalf is not the sentry's call.
        if (controls.Count == 2 && controls.ContainsKey(IdStaticText))
        {
            kind = ModalDialogKind.MacroMessageBox;
            message = ReadControlText(controls, IdStaticText);
            buttonToClick = okButton;
            return true;
        }

        return false;
    }

    private static string ReadControlText(Dictionary<int, IntPtr> controls, int id) =>
        controls.TryGetValue(id, out var handle) ? Bound(ReadWindowText(handle)) : string.Empty;

    private static string Bound(string value)
    {
        var collapsed = new string(value.Select(character => char.IsControl(character) ? ' ' : character).ToArray())
            .Replace("  ", " ", StringComparison.Ordinal)
            .Trim();
        return collapsed.Length <= MaxMessageLength ? collapsed : collapsed[..MaxMessageLength];
    }

    private static string ReadClassName(IntPtr handle)
    {
        var buffer = new char[64];
        var length = GetClassName(handle, buffer, buffer.Length);
        return length <= 0 ? string.Empty : new string(buffer, 0, length);
    }

    private static string ReadWindowText(IntPtr handle)
    {
        var buffer = new char[512];
        var length = GetWindowText(handle, buffer, buffer.Length);
        return length <= 0 ? string.Empty : new string(buffer, 0, length);
    }

    private const uint BM_CLICK = 0x00F5;
    private const uint SMTO_ABORTIFHUNG = 0x0002;

    private delegate bool EnumWindowProc(IntPtr handle, IntPtr parameter);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowProc callback, IntPtr parameter);

    [DllImport("user32.dll")]
    private static extern bool EnumChildWindows(IntPtr parent, EnumWindowProc callback, IntPtr parameter);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr handle, out uint processId);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "GetClassNameW")]
    private static extern int GetClassName(IntPtr handle, [Out] char[] buffer, int maxCount);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "GetWindowTextW")]
    private static extern int GetWindowText(IntPtr handle, [Out] char[] buffer, int maxCount);

    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr handle);

    [DllImport("user32.dll")]
    private static extern int GetDlgCtrlID(IntPtr handle);

    [DllImport("user32.dll", EntryPoint = "SendMessageTimeoutW")]
    private static extern IntPtr SendMessageTimeout(
        IntPtr handle,
        uint message,
        IntPtr wParam,
        IntPtr lParam,
        uint flags,
        uint timeoutMilliseconds,
        out IntPtr result);
}
