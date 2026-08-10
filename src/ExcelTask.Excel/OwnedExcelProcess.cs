using System.Collections.Concurrent;
using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.Versioning;
using System.Security.Cryptography;
using System.Text;
using ExcelTask.Core;

namespace ExcelTask.Excel;
internal sealed class OwnedExcelProcess
{
    private readonly ProcessIdentity _identity;

    private OwnedExcelProcess(ProcessIdentity identity) => _identity = identity;

    internal ProcessIdentity Identity => _identity;

    public static HashSet<ProcessIdentity> SnapshotExcelProcesses()
    {
        var identities = new HashSet<ProcessIdentity>();
        foreach (var process in Process.GetProcessesByName("EXCEL"))
        {
            using (process)
            {
                try { identities.Add(ProcessIdentity.Capture(process.Id)); }
                catch (InvalidOperationException) { }
                catch (System.ComponentModel.Win32Exception) { }
            }
        }

        return identities;
    }

    public static OwnedExcelProcess CaptureNew(object application, ISet<ProcessIdentity> preExisting)
    {
        var identity = GetApplicationIdentity(application);
        if (preExisting.Contains(identity))
        {
            throw new InvalidOperationException("Excel activation did not create a new owned process.");
        }

        return new OwnedExcelProcess(identity);
    }

    public static bool IsNewlyOwned(object application, ISet<ProcessIdentity> preExisting)
    {
        try { return !preExisting.Contains(GetApplicationIdentity(application)); }
        catch (COMException) { return false; }
        catch (InvalidOperationException) { return false; }
    }

    private static ProcessIdentity GetApplicationIdentity(object application)
    {
        var hwndValue = application.GetType().InvokeMember("Hwnd", BindingFlags.GetProperty, null, application, null, CultureInfo.InvariantCulture);
        var hwnd = new IntPtr(Convert.ToInt64(hwndValue, CultureInfo.InvariantCulture));
        _ = GetWindowThreadProcessId(hwnd, out var processId);
        if (processId == 0) throw new InvalidOperationException("Excel process identity could not be captured.");
        return ProcessIdentity.Capture((int)processId);
    }

    public bool WaitForExitOrTerminate()
    {
        if (!ProcessIdentity.TryOpenMatching(_identity, out var process)) return true;
        using (process)
        {
            if (process.WaitForExit(10_000)) return HasExited();
            if (!ProcessIdentity.TryOpenMatching(_identity, out var stillMatching)) return true;
            using (stillMatching)
            {
                stillMatching.Kill(entireProcessTree: false);
                _ = stillMatching.WaitForExit(5_000);
            }

            return HasExited();
        }
    }

    private bool HasExited()
    {
        if (!ProcessIdentity.TryOpenMatching(_identity, out var process)) return true;
        process.Dispose();
        return false;
    }

    public bool IsRunning => ProcessIdentity.TryOpenMatching(_identity, out var process) && DisposeProcess(process);

    private static bool DisposeProcess(Process process)
    {
        process.Dispose();
        return true;
    }

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
}

internal sealed record ProcessIdentity(int ProcessId, DateTime StartTimeUtc, string ExecutablePath)
{
    public static ProcessIdentity Capture(int processId)
    {
        using var process = Process.GetProcessById(processId);
        return new ProcessIdentity(processId, process.StartTime.ToUniversalTime(), GetExecutablePath(process));
    }

    public static bool TryOpenMatching(ProcessIdentity identity, out Process process)
    {
        process = null!;
        try
        {
            var candidate = Process.GetProcessById(identity.ProcessId);
            if (candidate.StartTime.ToUniversalTime() != identity.StartTimeUtc ||
                !string.Equals(GetExecutablePath(candidate), identity.ExecutablePath, StringComparison.OrdinalIgnoreCase))
            {
                candidate.Dispose();
                return false;
            }

            process = candidate;
            return true;
        }
        catch (ArgumentException)
        {
            return false;
        }
        catch (InvalidOperationException)
        {
            return false;
        }
    }

    private static string GetExecutablePath(Process process) => process.MainModule?.FileName
        ?? throw new InvalidOperationException("Excel executable path could not be captured.");
}
