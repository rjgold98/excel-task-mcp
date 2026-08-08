using ExcelTask.Core;

namespace ExcelTask.Excel.Tests;

public sealed class WorkbookRuntimeHelpersTests
{
    [Fact]
    public void StagingPathUsesTheFinalWorkbookDirectoryAndExtension()
    {
        var finalPath = Path.Combine(Path.GetTempPath(), "ExcelTask", "output.xlsm");

        var stagingPath = WorkbookRuntimeHelpers.CreateStagingPath(finalPath);

        Assert.Equal(Path.GetDirectoryName(Path.GetFullPath(finalPath)), Path.GetDirectoryName(stagingPath));
        Assert.Equal(".xlsm", Path.GetExtension(stagingPath));
        Assert.NotEqual(Path.GetFullPath(finalPath), stagingPath);
    }

    [Fact]
    public void PromoteStagingDoesNotReplaceDestinationThatAppearsLateWithoutAuthorization()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var finalPath = Path.Combine(directory, "output.xlsx");
        var stagingPath = WorkbookRuntimeHelpers.CreateStagingPath(finalPath);
        try
        {
            File.WriteAllText(stagingPath, "staged");
            File.WriteAllText(finalPath, "late-destination");

            Assert.Throws<IOException>(() => WorkbookRuntimeHelpers.PromoteStaging(stagingPath, finalPath, overwrite: false));
            Assert.Equal("late-destination", File.ReadAllText(finalPath));
            Assert.Equal("staged", File.ReadAllText(stagingPath));
        }
        finally
        {
            WorkbookRuntimeHelpers.TryDeleteStaging(stagingPath);
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void PromoteStagingReplacesDestinationOnlyWhenAuthorized()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var finalPath = Path.Combine(directory, "output.xlsx");
        var stagingPath = WorkbookRuntimeHelpers.CreateStagingPath(finalPath);
        try
        {
            File.WriteAllText(stagingPath, "staged");
            File.WriteAllText(finalPath, "old");

            WorkbookRuntimeHelpers.PromoteStaging(stagingPath, finalPath, overwrite: true);

            Assert.False(File.Exists(stagingPath));
            Assert.Equal("staged", File.ReadAllText(finalPath));
        }
        finally
        {
            WorkbookRuntimeHelpers.TryDeleteStaging(stagingPath);
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void TryDeleteStagingReportsAFileLockedAgainstDeletion()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var stagingPath = Path.Combine(directory, ".locked.xlsx");
        try
        {
            File.WriteAllText(stagingPath, "staged");
            using (var lockStream = new FileStream(stagingPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
            {
                Assert.False(WorkbookRuntimeHelpers.TryDeleteStaging(stagingPath));
                Assert.True(File.Exists(stagingPath));
            }

            Assert.True(WorkbookRuntimeHelpers.TryDeleteStaging(stagingPath));
        }
        finally
        {
            WorkbookRuntimeHelpers.TryDeleteStaging(stagingPath);
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void AutomationSecurityForceDisableUsesOfficeConstant()
    {
        Assert.Equal(3, WorkbookRuntimeHelpers.AutomationSecurityForceDisable);
    }

    [Fact]
    public void RotBindingPolicySkipsOnlyKnownNonmatchFailures()
    {
        Assert.True(RotWorkbookLocator.IsExpectedBindingNonmatch(new ArgumentException()));
        Assert.False(RotWorkbookLocator.IsExpectedBindingNonmatch(new InvalidOperationException()));
    }

    [Fact]
    public void RotDisplayNameMatchingRequiresTheExactNormalizedWorkbookPath()
    {
        var target = Path.Combine(Path.GetTempPath(), "ExcelTask", "target.xlsx");

        Assert.True(RotWorkbookLocator.MatchesDisplayName($"!{target}", target));
        Assert.False(RotWorkbookLocator.MatchesDisplayName($"!{target}.bak", target));
    }

    [Fact]
    public async Task ExecuteRejectsExistingCopyOutputBeforeStartingExcel()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var target = Path.Combine(directory, "target.xlsx");
        var reference = Path.Combine(directory, "reference.xlsx");
        var output = Path.Combine(directory, "output.xlsx");
        File.WriteAllText(target, string.Empty);
        File.WriteAllText(reference, string.Empty);
        File.WriteAllText(output, string.Empty);
        try
        {
            using var runtime = new ExcelWorkbookRuntime();
            var plan = new ExcelTaskPlan("test", new NormalizedExcelTaskRequest(
                target,
                reference,
                "Reference",
                "Imported",
                [],
                ExcelTaskMode.Apply,
                WorkbookBinding.Isolated,
                SaveMode.Copy,
                output,
                OverwriteConfirmed: false));

            var outcome = await runtime.ExecuteAsync(plan, CancellationToken.None);

            Assert.Equal(ExcelTaskStatus.Rejected, outcome.Status);
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public async Task InspectReportsExistingCopyOutputWithoutStartingExcel()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var target = Path.Combine(directory, "target.xlsx");
        var reference = Path.Combine(directory, "reference.xlsx");
        var output = Path.Combine(directory, "output.xlsx");
        File.WriteAllText(target, string.Empty);
        File.WriteAllText(reference, string.Empty);
        File.WriteAllText(output, string.Empty);
        try
        {
            using var runtime = new ExcelWorkbookRuntime();

            var inspection = await runtime.InspectAsync(
                new WorkbookInspectionRequest(target, reference, WorkbookBinding.Isolated, SaveMode.Copy, output),
                CancellationToken.None);

            Assert.True(inspection.CopyOutputExists);
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void PathsEqualNormalizesRelativeAndCaseInsensitivePaths()
    {
        var root = Path.GetTempPath();
        var left = Path.Combine(root, "ExcelTask", "Book.xlsx");
        var right = Path.Combine(root, "ExcelTask", ".", "book.xlsx");

        Assert.True(WorkbookRuntimeHelpers.PathsEqual(left, right));
    }

    [Fact]
    public void GetBoundsReturnsInclusiveRowsAndColumns()
    {
        var bounds = WorkbookRuntimeHelpers.GetBounds(new FormulaRepairRange("B2", "D4"));

        Assert.Equal(2, bounds.StartRow);
        Assert.Equal(2, bounds.StartColumn);
        Assert.Equal(3, bounds.RowCount);
        Assert.Equal(3, bounds.ColumnCount);
    }

    [Fact]
    public void CreateFormulaGridPreservesOnlyBlanksFormulasAndConstants()
    {
        object?[,] values = { { "=RC[-1]", null }, { 42, string.Empty } };

        var grid = WorkbookRuntimeHelpers.CreateFormulaGrid(values, 2, 2);

        Assert.Equal(FormulaCellKind.Formula, grid[0, 0].Kind);
        Assert.Equal(FormulaCellKind.Blank, grid[0, 1].Kind);
        Assert.Equal(FormulaCellKind.Constant, grid[1, 0].Kind);
        Assert.Equal(FormulaCellKind.Blank, grid[1, 1].Kind);
    }

    [Fact]
    public void ToA1AddressReturnsExpectedColumnsAtBoundaries()
    {
        Assert.Equal("A1", WorkbookRuntimeHelpers.ToA1Address(1, 1));
        Assert.Equal("Z10", WorkbookRuntimeHelpers.ToA1Address(10, 26));
        Assert.Equal("AA3", WorkbookRuntimeHelpers.ToA1Address(3, 27));
        Assert.Equal("XFD1048576", WorkbookRuntimeHelpers.ToA1Address(1_048_576, 16_384));
    }

    [Fact]
    public void ProcessIdentityMatchesTheCapturedCurrentProcess()
    {
        using var current = System.Diagnostics.Process.GetCurrentProcess();
        var identity = ProcessIdentity.Capture(current.Id);

        Assert.True(ProcessIdentity.TryOpenMatching(identity, out var matched));
        using (matched) Assert.Equal(current.Id, matched.Id);
    }

    [Fact]
    public async Task DispatcherSkipsCallbackWhenQueuedRequestIsCancelled()
    {
        using var dispatcher = new StaComDispatcher();
        var started = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var release = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var first = dispatcher.InvokeAsync(() =>
        {
            started.SetResult();
            release.Task.GetAwaiter().GetResult();
            return 1;
        }, CancellationToken.None);
        await started.Task;

        var callbackRan = false;
        using var cancellation = new CancellationTokenSource();
        var queued = dispatcher.InvokeAsync(() =>
        {
            callbackRan = true;
            return 2;
        }, cancellation.Token);
        cancellation.Cancel();
        release.SetResult();

        Assert.Equal(1, await first);
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => queued);
        Assert.False(callbackRan);
    }
}
