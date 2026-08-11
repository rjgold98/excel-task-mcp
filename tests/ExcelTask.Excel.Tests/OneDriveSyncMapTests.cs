namespace ExcelTask.Excel.Tests;

/// <summary>
/// The synced-path to service-URL mapping, asserted without a sync client present.
///
/// A workbook opened from a synced OneDrive or SharePoint folder reports its service URL as
/// FullName, not the local path the caller named, so exact-path matching refused every UseOpen
/// against the storage the owner actually keeps workbooks in. The registry half of this cannot be
/// exercised here - this machine registers no sync roots - so the mapping and the comparison are
/// separated from the lookup and pinned directly. The end-to-end path still needs one field run.
/// </summary>
public sealed class OneDriveSyncMapTests
{
    private const string MountPoint = @"C:\Users\person\OneDrive - Contoso";
    private const string UrlNamespace = "https://contoso.sharepoint.com/sites/Team/Shared Documents";

    [Fact]
    public void APathInsideTheSyncRootResolvesToItsServiceUrl()
    {
        Assert.True(OneDriveSyncMap.TryBuildUrl(
            MountPoint, UrlNamespace, @"C:\Users\person\OneDrive - Contoso\Reporting\Book.xlsm", out var url));

        Assert.Equal("https://contoso.sharepoint.com/sites/Team/Shared Documents/Reporting/Book.xlsm", url);
    }

    [Fact]
    public void APathOutsideEverySyncRootResolvesToNothing()
    {
        // The whole guarantee rests on this: an unsynced path must keep behaving exactly as before,
        // which means producing no URL rather than a plausible-looking one.
        Assert.False(OneDriveSyncMap.TryBuildUrl(MountPoint, UrlNamespace, @"C:\Work\Book.xlsm", out _));
        Assert.False(OneDriveSyncMap.TryBuildUrl(MountPoint, UrlNamespace, MountPoint, out _));
    }

    [Fact]
    public void ASiblingFolderSharingANamePrefixDoesNotResolveIntoTheSyncRoot()
    {
        // "OneDrive - Contoso Ltd" starts with "OneDrive - Contoso". Matching on the prefix alone
        // would map a different library's workbook onto this one's URL and call it the same file -
        // which is the precise mistake the exact-path check exists to prevent.
        Assert.False(OneDriveSyncMap.TryBuildUrl(
            MountPoint, UrlNamespace, @"C:\Users\person\OneDrive - Contoso Ltd\Reporting\Book.xlsm", out _));
    }

    [Theory]
    [InlineData(@"C:\Users\person\OneDrive - Contoso\Reporting\Book.xlsm")]
    [InlineData(@"C:\Users\person\onedrive - contoso\reporting\book.xlsm")]
    [InlineData(@"C:\Users\person\OneDrive - Contoso\.\Reporting\Book.xlsm")]
    [InlineData(@"C:\Users\person\OneDrive - Contoso\Other\..\Reporting\Book.xlsm")]
    public void TheSameWorkbookResolvesTheSameWayThroughEverySpellingOfItsPath(string localPath)
    {
        Assert.True(OneDriveSyncMap.TryBuildUrl(MountPoint, UrlNamespace, localPath, out var url));
        Assert.True(OneDriveSyncMap.UrlsEqual(
            "https://contoso.sharepoint.com/sites/Team/Shared Documents/Reporting/Book.xlsm", url));
    }

    [Theory]
    // Excel reports the space literally in some paths and percent-encoded in others.
    [InlineData("https://contoso.sharepoint.com/sites/Team/Shared%20Documents/Book.xlsm",
                "https://contoso.sharepoint.com/sites/Team/Shared Documents/Book.xlsm")]
    // SharePoint treats the path case-insensitively, and a trailing slash names the same resource.
    [InlineData("https://Contoso.sharepoint.com/Sites/Team/Book.xlsm",
                "https://contoso.sharepoint.com/sites/team/Book.xlsm")]
    [InlineData("https://contoso.sharepoint.com/sites/Team/Book.xlsm/",
                "https://contoso.sharepoint.com/sites/Team/Book.xlsm")]
    public void UrlsThatNameTheSameWorkbookCompareEqual(string reported, string resolved) =>
        Assert.True(OneDriveSyncMap.UrlsEqual(reported, resolved));

    [Theory]
    [InlineData("https://contoso.sharepoint.com/sites/Team/Book.xlsm",
                "https://contoso.sharepoint.com/sites/Other/Book.xlsm")]
    [InlineData("https://contoso.sharepoint.com/sites/Team/Book.xlsm",
                "https://contoso.sharepoint.com/sites/Team/Budget.xlsm")]
    [InlineData("", "https://contoso.sharepoint.com/sites/Team/Book.xlsm")]
    public void UrlsThatNameDifferentWorkbooksDoNot(string reported, string resolved) =>
        Assert.False(OneDriveSyncMap.UrlsEqual(reported, resolved));

    [Fact]
    public void LocalIdentityStillMatchesExactlyAndNothingElseDoes()
    {
        var path = Path.Combine(Path.GetTempPath(), "ExcelTask", "Book.xlsx");

        Assert.True(WorkbookRuntimeHelpers.IdentifiesSameWorkbook(path, path));
        Assert.False(WorkbookRuntimeHelpers.IdentifiesSameWorkbook(
            Path.Combine(Path.GetTempPath(), "ExcelTask", "Other.xlsx"), path));

        // An unsaved workbook reports a bare name, and a URL that resolves to nothing on this
        // machine reports no match - neither may throw out of the search that would have found it.
        Assert.False(WorkbookRuntimeHelpers.IdentifiesSameWorkbook("Book1", path));
        Assert.False(WorkbookRuntimeHelpers.IdentifiesSameWorkbook(null, path));
        Assert.False(WorkbookRuntimeHelpers.IdentifiesSameWorkbook(
            "https://contoso.sharepoint.com/sites/Team/Book.xlsx", path));
    }
}
