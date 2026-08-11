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
    private const string ButtonClassName = "Button";
    // The VBA editor's own frame window. A compile error belongs to the editor; a MsgBox raised by
    // running macro code has no owner at all. Measured, and unlike a caption it is not translated.
    private const string VbaEditorClassName = "wndclass_desked_gsk";
    private const int MaxMessageLength = 200;

    // Control identifiers are stable across Office language packs, unlike button captions.
    // Beware: identifier 2 is the OK button on a one-button dialog but the Cancel button on
    // OK/Cancel and Retry/Cancel, so it is only ever pressed when no other choice button exists.
    private const int IdOk = 2;
    private const int IdHelp = 9;
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

    /// <summary>
    /// Dialog windows already acted on, so a dialog that takes a moment to close is not counted
    /// again on the next pass. Touched only from the watch loop, so it needs no gate;
    /// <see cref="_dismissed"/> does, because the macro path reads it from another thread.
    /// </summary>
    private readonly HashSet<IntPtr> _acted = [];

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

    /// <summary>
    /// Stops watching, and does not return while a pass could still act.
    ///
    /// The wait is generous because a pass already inside SendMessageTimeout budgets two seconds
    /// per dialog on top of a process open for the identity re-check, and the old two-second wait
    /// could expire underneath it. That was worse than a slow Dispose: it disposed the token source
    /// while the loop was live, so the next Delay threw ObjectDisposedException past a catch that
    /// only handles cancellation and faulted the loop unobserved - and it returned control to a
    /// caller that goes straight into SaveAs while a sentry pass could still terminate owned Excel.
    /// In the ordinary case the loop is parked in Delay and cancels in microseconds, so this costs
    /// nothing; the wait only materialises when something really is wedged, which is when it should.
    /// </summary>
    public void Dispose()
    {
        _stop.Cancel();

        bool finished;
        try { finished = _loop.Wait(TimeSpan.FromSeconds(15)); }
        catch (AggregateException) { finished = true; }

        if (!finished)
        {
            // Never yank the token source out from under a live loop. An undisposed source holds no
            // timer and no handle here, so leaving it is strictly cheaper than the race.
            _ = _loop.ContinueWith(static task => _ = task.Exception, TaskScheduler.Default);
            return;
        }

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

        // Handles whose windows are gone are dropped, so a recycled handle belonging to a genuinely
        // new dialog is answered rather than mistaken for one already handled.
        _acted.RemoveWhere(handle => !IsWindow(handle));

        foreach (var dialog in EnumerateDialogs((uint)_identity.ProcessId))
        {
            // Already acted on. Without this the same window was re-counted every 500 ms, so one
            // MsgBox that took 1.6 s to close arrived in the receipt as three dialogs answered.
            if (!_acted.Add(dialog)) continue;
            if (!TryClassify(dialog, out var kind, out var message, out var buttonToClick))
            {
                // Not ours to answer - a Cancel-bearing dialog, say. Do not hold the handle, or a
                // dialog that later becomes answerable would be skipped.
                _acted.Remove(dialog);
                continue;
            }

            if (kind == ModalDialogKind.VbaCompileError)
            {
                lock (_gate) _dismissed.Add(new DismissedDialog(kind, message));
                // Measured behaviour: answering a compile error does not hand control back. VBA
                // drops into break mode with the module open in the editor, and the automation call
                // that is already blocked inside COM never returns. Ending this instance is the only
                // way to release it. That is safe here because a compile error happens before the
                // macro runs and before anything is saved, so the workbook on disk is untouched.
                TerminateOwnedProcess();
                continue;
            }

            // A timeout keeps a wedged UI thread from blocking the sentry itself, which would defeat
            // the whole point of watching from outside the blocked call. Its return is evidence,
            // not decoration: non-zero means the window procedure processed BM_CLICK, and zero is
            // what SMTO_ABORTIFHUNG returns against exactly the hung thread this class exists for.
            // Recording before the call, and discarding its result, is how a dialog nobody ever
            // clicked reached the receipt as "answered".
            //
            // Delivery is the evidence, deliberately, rather than the window later disappearing:
            // the macro path reads Dismissed the instant the run returns, so evidence that arrives
            // a poll later arrives after the receipt is built and the dialog goes unreported.
            var delivered = SendMessageTimeout(
                buttonToClick, BM_CLICK, IntPtr.Zero, IntPtr.Zero, SMTO_ABORTIFHUNG, 2000, out _) != IntPtr.Zero;
            if (delivered) lock (_gate) _dismissed.Add(new DismissedDialog(kind, message));
            else _acted.Remove(dialog);
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
        var buttonIds = new HashSet<int>();
        EnumChildWindows(dialog, (child, state) =>
        {
            var id = GetDlgCtrlID(child);
            controls.TryAdd(id, child);
            if (ReadClassName(child) == ButtonClassName) buttonIds.Add(id);
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

        // A compile error and a MsgBox carrying a help button have identical control identifiers, so
        // ownership is what separates them. Only the editor's own dialog may end this process; a
        // dialog raised by macro code never can, however much it looks like one.
        if (IsOwnedByVbaEditor(dialog))
        {
            if (!controls.ContainsKey(IdHelp)) return false;
            kind = ModalDialogKind.VbaCompileError;
            message = ReadControlText(controls, IdStaticText);
            buttonToClick = okButton;
            return true;
        }

        // A MsgBox whose only real choice is OK. Help is auxiliary, and an icon adds a Static rather
        // than a Button, so both are tolerated. Anything offering a genuine choice - Yes/No,
        // OK/Cancel, Retry/Cancel - is left alone, because answering for the user is not this
        // class's call. That exclusion is by button identity, not by control count: on those dialogs
        // identifier 2 is Cancel, so counting alone would press the wrong button.
        if (controls.ContainsKey(IdStaticText) && buttonIds.All(id => id is IdOk or IdHelp))
        {
            kind = ModalDialogKind.MacroMessageBox;
            message = ReadControlText(controls, IdStaticText);
            buttonToClick = okButton;
            return true;
        }

        return false;
    }

    /// <summary>True when the dialog belongs to the VBA editor frame rather than to macro code.</summary>
    private static bool IsOwnedByVbaEditor(IntPtr dialog)
    {
        var owner = GetWindow(dialog, GwOwner);
        return owner != IntPtr.Zero && ReadClassName(owner) == VbaEditorClassName;
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
    private const uint GwOwner = 4;

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
    private static extern bool IsWindow(IntPtr handle);

    [DllImport("user32.dll")]
    private static extern int GetDlgCtrlID(IntPtr handle);

    [DllImport("user32.dll")]
    private static extern IntPtr GetWindow(IntPtr handle, uint command);

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
