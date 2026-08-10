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

    public T Add<T>(T value) where T : class
    {
        _references.Add(value);
        return value;
    }

    public void Dispose()
    {
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
