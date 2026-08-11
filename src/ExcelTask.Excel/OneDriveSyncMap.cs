using Microsoft.Win32;

namespace ExcelTask.Excel;

/// <summary>
/// Resolves a local synced path to the URL Excel reports for the same workbook.
///
/// A workbook opened from a synced OneDrive or SharePoint folder does not report the path the
/// person sees in Explorer. Its FullName is the service URL -
/// https://tenant.sharepoint.com/sites/Team/Shared Documents/Book.xlsm - while the caller names
/// C:\Users\...\OneDrive - Tenant\Team\Book.xlsm. Exact-path matching then finds nothing open and
/// refuses. That refusal is correct for the question it asked and useless for the one that
/// mattered: a field session produced four consecutive refusals in a row, each reporting that the
/// workbook name matched and the path did not, against the storage the owner keeps everything in.
///
/// This does not loosen the identity check. It resolves the caller's path through the sync client's
/// own published mapping and then compares exactly, so the workbook is still proven rather than
/// guessed - a same-named file in a different library still does not match. Where no mapping
/// applies, because nothing is synced or the path sits outside every sync root, nothing changes.
///
/// The mapping is the one the sync engine registers for the shell: HKCU\Software\SyncEngines\
/// Providers\OneDrive\{id} carries MountPoint, the local root, and UrlNamespace, its service URL.
/// </summary>
internal static class OneDriveSyncMap
{
    private const string ProvidersKeyPath = @"Software\SyncEngines\Providers\OneDrive";

    /// <summary>Whether a URL Excel reported denotes the same workbook as a local path.</summary>
    public static bool MatchesLocalPath(string reportedUrl, string localPath)
    {
        foreach (var (mountPoint, urlNamespace) in ReadProviders())
        {
            if (TryBuildUrl(mountPoint, urlNamespace, localPath, out var candidate) &&
                UrlsEqual(reportedUrl, candidate))
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// The service URL for a path inside a sync root, or false when the path is outside it.
    /// Separated from the registry so the mapping itself is testable on a machine that syncs
    /// nothing - which is every machine but the one where this matters.
    /// </summary>
    internal static bool TryBuildUrl(string mountPoint, string urlNamespace, string localPath, out string url)
    {
        url = string.Empty;
        if (string.IsNullOrWhiteSpace(mountPoint) || string.IsNullOrWhiteSpace(urlNamespace)) return false;

        string root;
        string full;
        try
        {
            root = Path.GetFullPath(mountPoint).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            full = Path.GetFullPath(localPath);
        }
        catch (Exception exception) when (exception is ArgumentException or NotSupportedException or PathTooLongException)
        {
            return false;
        }

        // A separator is required after the root so that a sibling folder sharing a name prefix -
        // "OneDrive - Contoso Ltd" against a root of "OneDrive - Contoso" - cannot resolve into it.
        if (full.Length <= root.Length + 1) return false;
        if (!full.StartsWith(root, StringComparison.OrdinalIgnoreCase)) return false;
        if (full[root.Length] is not ('\\' or '/')) return false;

        var relative = full[(root.Length + 1)..].Replace('\\', '/');
        url = $"{urlNamespace.TrimEnd('/')}/{relative}";
        return true;
    }

    /// <summary>
    /// Compares two service URLs for the same workbook. Excel reports spaces literally in some
    /// paths and percent-encoded in others, so both sides are decoded before comparison; SharePoint
    /// treats the path case-insensitively, so this does too.
    /// </summary>
    internal static bool UrlsEqual(string? left, string? right)
    {
        if (string.IsNullOrWhiteSpace(left) || string.IsNullOrWhiteSpace(right)) return false;
        return string.Equals(NormalizeUrl(left), NormalizeUrl(right), StringComparison.OrdinalIgnoreCase);
    }

    private static string NormalizeUrl(string url)
    {
        var trimmed = url.Trim().TrimEnd('/');
        try { return Uri.UnescapeDataString(trimmed); }
        catch (UriFormatException) { return trimmed; }
    }

    /// <summary>
    /// A redacted self-test of the half that only a synced machine can exercise: how many sync
    /// roots are registered, and how many of them resolve a path beneath themselves back to a URL
    /// under their own namespace.
    ///
    /// Counts only, deliberately. A UrlNamespace names the tenant and the site collection - an
    /// internal server name - and a MountPoint names the person. Neither leaves the machine. What
    /// leaves is whether the mapping worked, which is the whole question.
    /// </summary>
    internal static (int Registered, int Resolving) SelfTest()
    {
        const string ProbeRelative = "ExcelTask-probe/workbook.xlsx";
        var providers = ReadProviders();
        var resolving = 0;

        foreach (var (mountPoint, urlNamespace) in providers)
        {
            // The probe path never has to exist; only the arithmetic is under test.
            var probe = Path.Combine(mountPoint, "ExcelTask-probe", "workbook.xlsx");
            if (TryBuildUrl(mountPoint, urlNamespace, probe, out var url) &&
                url.StartsWith(urlNamespace.TrimEnd('/'), StringComparison.OrdinalIgnoreCase) &&
                url.EndsWith(ProbeRelative, StringComparison.OrdinalIgnoreCase))
            {
                resolving++;
            }
        }

        return (providers.Count, resolving);
    }

    /// <summary>
    /// The sync roots this user has registered. A machine with no sync client, or a policy that
    /// denies the key, yields none - and every caller then behaves exactly as it did before.
    /// </summary>
    private static List<(string MountPoint, string UrlNamespace)> ReadProviders()
    {
        List<(string MountPoint, string UrlNamespace)> providers = [];
        try
        {
            using var root = Registry.CurrentUser.OpenSubKey(ProvidersKeyPath);
            if (root is null) return providers;

            foreach (var name in root.GetSubKeyNames())
            {
                using var provider = root.OpenSubKey(name);
                if (provider?.GetValue("MountPoint") is string mountPoint &&
                    provider.GetValue("UrlNamespace") is string urlNamespace)
                {
                    providers.Add((mountPoint, urlNamespace));
                }
            }
        }
        catch (Exception exception) when (exception is System.Security.SecurityException or UnauthorizedAccessException or IOException)
        {
            return providers;
        }

        return providers;
    }
}
