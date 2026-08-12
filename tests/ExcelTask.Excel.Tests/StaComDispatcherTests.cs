namespace ExcelTask.Excel.Tests;

/// <summary>
/// Runs alone, because the count these tests read is process-global and every
/// <c>ExcelWorkbookRuntime</c> constructs a dispatcher. Two other fast-tier classes build runtimes,
/// so without this the assertions race their setup and report a defect that is not there - which
/// they did on the first run.
/// </summary>
[CollectionDefinition("sta-dispatcher", DisableParallelization = true)]
public sealed class SerialStaComDispatcherTests;

/// <summary>
/// The STA dispatcher's own lifecycle, asserted by count.
///
/// A redundant field initializer in <c>PendingVerification</c> constructed a second dispatcher on
/// every mutating Apply and orphaned it immediately: its thread parked forever in the consuming
/// enumerable, it never received CompleteAdding, and its queue was never disposed. That survived
/// for its entire life because nothing observed how many existed - the full gate was green
/// throughout, and it was found by reading, lost, and re-found hours later by an independent scan.
/// These assert the property that was missing rather than the line that was wrong.
/// </summary>
[Collection("sta-dispatcher")]
public sealed class StaComDispatcherTests
{
    [Fact]
    public void ConstructingAndDisposingLeavesTheCountWhereItStarted()
    {
        var before = StaComDispatcher.LiveInstances;

        for (var round = 0; round < 5; round++)
        {
            using var dispatcher = new StaComDispatcher();
            Assert.Equal(before + 1, StaComDispatcher.LiveInstances);
        }

        Assert.Equal(before, StaComDispatcher.LiveInstances);
    }

    [Fact]
    public void DisposingTwiceDecrementsOnce()
    {
        // Dispose is idempotent, so the count must be too - otherwise the instrumentation would
        // read negative and a future leak would hide behind it.
        var before = StaComDispatcher.LiveInstances;
        var dispatcher = new StaComDispatcher();

        dispatcher.Dispose();
        dispatcher.Dispose();

        Assert.Equal(before, StaComDispatcher.LiveInstances);
    }

    [Fact]
    public async Task AnUndisposedDispatcherIsVisibleInTheCount()
    {
        // The shape of the defect, made observable: something constructs a dispatcher, does work on
        // it, and never disposes it. Without this the count could only ever prove the happy path.
        var before = StaComDispatcher.LiveInstances;
        var orphan = new StaComDispatcher();

        Assert.Equal(42, await orphan.InvokeAsync(() => 42, CancellationToken.None));
        Assert.Equal(before + 1, StaComDispatcher.LiveInstances);

        orphan.Dispose();
        Assert.Equal(before, StaComDispatcher.LiveInstances);
    }
}
