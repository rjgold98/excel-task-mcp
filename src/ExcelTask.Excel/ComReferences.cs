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
internal sealed class ComReferenceScope : IDisposable
{
    private readonly HashSet<object> _references = new(ReferenceEqualityComparer.Instance);
    private bool _abandoned;

    public T Add<T>(T value) where T : class
    {
        _references.Add(value);
        return value;
    }

    /// <summary>
    /// Forgets every reference without releasing it, for the one case where releasing is unsafe:
    /// the Excel process these proxies point at has already been terminated.
    ///
    /// Releasing a proxy asks the COM runtime to talk to a server that no longer exists. Usually
    /// that fails quietly, but after the VBA editor has been materialized in this process it can
    /// fault hard enough to take the whole process down - observed as a test host crash when a
    /// macro run was followed by a compile error that ended its Excel. Abandoning leaks nothing
    /// that matters: the server is gone, so the resources the release would have freed are gone
    /// with it, and the runtime cleans up the local wrappers.
    /// </summary>
    public void Abandon()
    {
        _abandoned = true;
        _references.Clear();
    }

    public void Dispose()
    {
        if (_abandoned) return;
        foreach (var reference in _references) ComReferences.Release(reference);
        _references.Clear();
    }
}

internal static class ComReferences
{
    public static void Release(object? value)
    {
        if (value is not null && Marshal.IsComObject(value)) Marshal.FinalReleaseComObject(value);
    }
}
