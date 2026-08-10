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
internal sealed class RotWorkbookLocator : IDisposable
{
    private object? _workbook;

    private RotWorkbookLocator(object workbook) => _workbook = workbook;

    public object Workbook => _workbook ?? throw new ObjectDisposedException(nameof(RotWorkbookLocator));

    public object Detach()
    {
        var workbook = Workbook;
        _workbook = null;
        return workbook;
    }

    public static RotWorkbookLocator? Find(string targetPath)
    {
        var result = GetRunningObjectTable(0, out var table);
        if (result < 0 || table is null) Marshal.ThrowExceptionForHR(result);
        var runningTable = table ?? throw new InvalidOperationException("The running object table was unavailable.");
        try
        {
            runningTable.EnumRunning(out var enumerator);
            try
            {
                var monikers = new IMoniker[1];
                while (enumerator.Next(1, monikers, IntPtr.Zero) == 0)
                {
                    var moniker = monikers[0];
                    try
                    {
                        var bindResult = CreateBindCtx(0, out var bindContext);
                        if (bindResult < 0 || bindContext is null) continue;
                        try
                        {
                            moniker.GetDisplayName(bindContext, null, out var displayName);
                            if (!MatchesDisplayName(displayName, targetPath)) continue;
                            moniker.BindToObject(bindContext, null, ref WorkbookInterfaceId, out var candidate);
                            if (candidate is not null && HasMatchingFullName(candidate, targetPath)) return new RotWorkbookLocator(candidate);
                            ComReferences.Release(candidate);
                        }
                        catch (Exception exception) when (IsExpectedBindingNonmatch(exception)) { }
                        finally { ComReferences.Release(bindContext); }
                    }
                    finally { ComReferences.Release(moniker); }
                }
            }
            finally { ComReferences.Release(enumerator); }
        }
        finally { ComReferences.Release(runningTable); }

        return null;
    }

    public static bool ContainsPath(string targetPath)
    {
        var result = GetRunningObjectTable(0, out var table);
        if (result < 0 || table is null) Marshal.ThrowExceptionForHR(result);
        var runningTable = table ?? throw new InvalidOperationException("The running object table was unavailable.");
        try
        {
            runningTable.EnumRunning(out var enumerator);
            try
            {
                var monikers = new IMoniker[1];
                while (enumerator.Next(1, monikers, IntPtr.Zero) == 0)
                {
                    var moniker = monikers[0];
                    try
                    {
                        var bindResult = CreateBindCtx(0, out var bindContext);
                        if (bindResult < 0 || bindContext is null) continue;
                        try
                        {
                            moniker.GetDisplayName(bindContext, null, out var displayName);
                            if (MatchesDisplayName(displayName, targetPath)) return true;
                        }
                        catch (Exception exception) when (IsExpectedBindingNonmatch(exception)) { }
                        finally { ComReferences.Release(bindContext); }
                    }
                    finally { ComReferences.Release(moniker); }
                }
            }
            finally { ComReferences.Release(enumerator); }
        }
        finally { ComReferences.Release(runningTable); }

        return false;
    }

    /// <summary>
    /// Detects an exact target workbook in the ROT that belongs to a different Excel application.
    /// The caller's application RCWs are intentionally never released here: ROT can return the
    /// same RCW instance used by the active session, and final-releasing it would invalidate that session.
    /// </summary>
    public static bool HasExternalWorkbookAtPath(string targetPath, long sessionApplicationHwnd)
    {
        var result = GetRunningObjectTable(0, out var table);
        if (result < 0 || table is null) Marshal.ThrowExceptionForHR(result);
        var runningTable = table ?? throw new InvalidOperationException("The running object table was unavailable.");
        try
        {
            runningTable.EnumRunning(out var enumerator);
            try
            {
                var monikers = new IMoniker[1];
                while (enumerator.Next(1, monikers, IntPtr.Zero) == 0)
                {
                    var moniker = monikers[0];
                    try
                    {
                        var bindResult = CreateBindCtx(0, out var bindContext);
                        if (bindResult < 0 || bindContext is null) continue;
                        try
                        {
                            moniker.GetDisplayName(bindContext, null, out var displayName);
                            if (!MatchesDisplayName(displayName, targetPath)) continue;

                            object? candidate = null;
                            object? candidateApplication = null;
                            var unrelatedCandidate = false;
                            try
                            {
                                moniker.BindToObject(bindContext, null, ref WorkbookInterfaceId, out candidate);
                                if (candidate is null || !HasMatchingFullName(candidate, targetPath)) continue;

                                candidateApplication = GetApplication(candidate);
                                var candidateApplicationHwnd = GetApplicationHwnd(candidateApplication);
                                unrelatedCandidate = IsExternalApplicationHwnd(sessionApplicationHwnd, candidateApplicationHwnd);
                                if (unrelatedCandidate) return true;
                            }
                            catch (Exception exception) when (IsExpectedBindingNonmatch(exception)) { }
                            finally
                            {
                                if (unrelatedCandidate)
                                {
                                    ComReferences.Release(candidateApplication);
                                    ComReferences.Release(candidate);
                                }
                            }
                        }
                        finally { ComReferences.Release(bindContext); }
                    }
                    finally { ComReferences.Release(moniker); }
                }
            }
            finally { ComReferences.Release(enumerator); }
        }
        finally { ComReferences.Release(runningTable); }

        return false;
    }

    public void Dispose() => ComReferences.Release(_workbook);

    internal static bool IsExpectedBindingNonmatch(Exception exception) => exception is COMException or ArgumentException;

    internal static bool RequiresPreMutationIsolatedSameApplyRevalidation(
        ExcelTaskMode mode,
        WorkbookBinding binding,
        SaveMode save) =>
        mode == ExcelTaskMode.Apply &&
        binding == WorkbookBinding.Isolated &&
        save == SaveMode.Same;

    internal static bool IsExternalApplicationHwnd(long sessionApplicationHwnd, long candidateApplicationHwnd) =>
        sessionApplicationHwnd != candidateApplicationHwnd;

    internal static bool MatchesDisplayName(string? displayName, string targetPath)
    {
        if (string.IsNullOrWhiteSpace(displayName)) return false;
        var candidate = displayName[0] == '!' ? displayName[1..] : displayName;
        return WorkbookRuntimeHelpers.PathsEqual(candidate, targetPath);
    }

    private static bool HasMatchingFullName(object candidate, string targetPath)
    {
        try
        {
            var fullName = candidate.GetType().InvokeMember("FullName", BindingFlags.GetProperty, null, candidate, null, CultureInfo.InvariantCulture) as string;
            return fullName is not null && WorkbookRuntimeHelpers.PathsEqual(fullName, targetPath);
        }
        catch (COMException)
        {
            return false;
        }
    }

    private static object GetApplication(object workbook) => workbook.GetType().InvokeMember(
        "Application",
        BindingFlags.GetProperty,
        null,
        workbook,
        null,
        CultureInfo.InvariantCulture) ?? throw new InvalidOperationException("Excel workbook did not return its application.");

    private static long GetApplicationHwnd(object application) => Convert.ToInt64(application.GetType().InvokeMember(
        "Hwnd",
        BindingFlags.GetProperty,
        null,
        application,
        null,
        CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);

    private static Guid WorkbookInterfaceId = new("00000000-0000-0000-C000-000000000046");

    [DllImport("ole32.dll")]
    private static extern int GetRunningObjectTable(uint reserved, out IRunningObjectTable? runningObjectTable);

    [DllImport("ole32.dll")]
    private static extern int CreateBindCtx(uint reserved, out IBindCtx? bindContext);
}
