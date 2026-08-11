using System.Collections.Concurrent;

namespace ExcelTask.Excel.Tests;

/// <summary>
/// The second adapter at the observer seam, and the reason the seam is now real.
///
/// The runtime already reports every phase, every owned Excel process the moment it exists, and
/// every staging path before it is written - which is exactly the ordering evidence the
/// mutate-save-verify choreography depends on and, until this adapter, nothing asserted. The
/// production adapter discards all of it. This one records it, so a test can state the ordering
/// constraints as facts rather than trusting that six hand-written copies of the sequence all
/// remembered them: staging is announced before anything saves to it, cleanup is proven before
/// verification opens the file, and no owned process ever goes unannounced.
/// </summary>
internal sealed class RecordingRuntimeObserver : IExcelWorkbookRuntimeObserver
{
    private readonly ConcurrentQueue<string> _events = new();

    public IReadOnlyList<string> Events => [.. _events];

    public void OnPhase(string phase) => _events.Enqueue($"phase:{phase}");

    public void OnOwnedProcessCaptured(ProcessIdentity identity) => _events.Enqueue("owned-process");

    public void OnStagingPathCreated(string stagingPath) => _events.Enqueue("staging-path");

    public int CountOf(string eventName) => Events.Count(item => item == eventName);

    /// <summary>Index of the first occurrence, or fails the test naming what never happened.</summary>
    public int IndexOf(string eventName)
    {
        var events = Events;
        for (var index = 0; index < events.Count; index++)
        {
            if (events[index] == eventName) return index;
        }

        Assert.Fail($"The runtime never reported '{eventName}'. Reported: {string.Join(", ", events)}");
        return -1;
    }
}
