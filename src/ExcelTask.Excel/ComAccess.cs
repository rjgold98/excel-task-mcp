using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;

namespace ExcelTask.Excel;

/// <summary>
/// Late-bound access to the Office object models.
///
/// Everything here is knowledge about IDispatch that was previously restated in nine places and
/// disagreed with itself in three of them. It is one module because the alternative was measured:
/// the same member kind is bound differently by different object models - Excel exposes
/// <c>Worksheets.Item</c> as a parameterized property while Office and the VBA extensibility model
/// expose <c>Item</c> as a method - and choosing wrong raises DISP_E_MEMBERNOTFOUND at run time
/// only, on a machine with Excel installed. That single confusion produced three separate shipped
/// defects in one day.
///
/// <see cref="Item"/> is the only member that tries both bindings, and deliberately so. Reads may
/// be retried harmlessly; a write must never be retried against a different member, because
/// succeeding on the wrong one would silently put a value somewhere nobody asked for.
/// </summary>
internal static class ComAccess
{
    private const int DispIdMemberNotFound = unchecked((int)0x80020003);

    /// <summary>Reads a property, with arguments for the parameterized ones such as Range.</summary>
    public static object Get(object target, string member, params object?[] arguments) =>
        target.GetType().InvokeMember(member, BindingFlags.GetProperty, null, target, arguments, CultureInfo.InvariantCulture)
        ?? throw new InvalidOperationException($"Excel did not return '{member}'.");

    /// <summary>
    /// Reads a property whose null is an answer rather than a fault. An empty cell's Value2 is
    /// null, and treating that as a failure would make "the cell is blank" indistinguishable from
    /// "Excel would not answer".
    /// </summary>
    public static object? GetOrNull(object target, string member, params object?[] arguments) =>
        target.GetType().InvokeMember(member, BindingFlags.GetProperty, null, target, arguments, CultureInfo.InvariantCulture);

    /// <summary>Writes a property. Never retried against another binding.</summary>
    public static void Set(object target, string member, object? value) =>
        target.GetType().InvokeMember(member, BindingFlags.SetProperty, null, target, [value], CultureInfo.InvariantCulture);

    /// <summary>Calls a method. Returns null for the void ones.</summary>
    public static object? Invoke(object target, string member, params object?[] arguments) =>
        target.GetType().InvokeMember(member, BindingFlags.InvokeMethod, null, target, arguments, CultureInfo.InvariantCulture);

    /// <summary>
    /// Resolves one collection entry whichever way the owning object model exposes Item. Tries the
    /// property binding first because Excel's own collections are the common case, then the method
    /// binding on the exact missing-member failure and nothing else.
    /// </summary>
    public static object Item(object collection, object index)
    {
        try
        {
            return collection.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, collection, [index], CultureInfo.InvariantCulture)
                ?? throw new InvalidOperationException("The collection returned no entry.");
        }
        catch (Exception exception) when (IsMissingMember(exception))
        {
            return collection.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, collection, [index], CultureInfo.InvariantCulture)
                ?? throw new InvalidOperationException("The collection returned no entry.");
        }
    }

    /// <summary>True only for DISP_E_MEMBERNOTFOUND, so a real failure is never retried as a binding mismatch.</summary>
    public static bool IsMissingMember(Exception exception) =>
        (exception as COMException)?.HResult == DispIdMemberNotFound ||
        (exception is TargetInvocationException { InnerException: COMException inner } && inner.HResult == DispIdMemberNotFound) ||
        exception is MissingMethodException or MissingMemberException;

    public static void Release(object? value)
    {
        if (value is not null && Marshal.IsComObject(value)) Marshal.FinalReleaseComObject(value);
    }
}
