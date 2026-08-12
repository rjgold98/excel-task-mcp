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
internal sealed class StaComDispatcher : IDisposable
{
    private static int _liveInstances;

    private readonly BlockingCollection<IWorkItem> _queue = new(boundedCapacity: 32);
    private readonly Thread _thread;
    private bool _disposed;

    /// <summary>
    /// Dispatchers constructed and not yet disposed, which is one STA thread each.
    ///
    /// Instrumentation rather than bookkeeping: nothing here reads it. A redundant field
    /// initializer once constructed a second dispatcher on every mutating Apply and immediately
    /// orphaned it - the thread parked forever, the queue was never disposed - and that was
    /// invisible for the defect's entire life precisely because nothing counted. A test can now
    /// assert the count returns to where it started.
    /// </summary>
    internal static int LiveInstances => Volatile.Read(ref _liveInstances);

    public StaComDispatcher()
    {
        _thread = new Thread(Run) { IsBackground = true, Name = "ExcelTask COM STA" };
        _thread.SetApartmentState(ApartmentState.STA);
        _thread.Start();
        Interlocked.Increment(ref _liveInstances);
    }

    public Task<T> InvokeAsync<T>(Func<T> callback, CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(callback);
        cancellationToken.ThrowIfCancellationRequested();
        var item = new WorkItem<T>(callback, cancellationToken);
        if (!_queue.TryAdd(item))
        {
            ObjectDisposedException.ThrowIf(_queue.IsAddingCompleted, this);
            throw new InvalidOperationException("Excel runtime queue is full; retry after queued tasks complete.");
        }
        return item.Task;
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        _queue.CompleteAdding();
        _thread.Join();
        _queue.Dispose();
        Interlocked.Decrement(ref _liveInstances);
    }

    private void Run()
    {
        _ = PeekMessage(out _, IntPtr.Zero, 0, 0, 0);
        foreach (var item in _queue.GetConsumingEnumerable())
        {
            item.Run();
            PumpMessages();
        }
    }

    private static void PumpMessages()
    {
        while (PeekMessage(out var message, IntPtr.Zero, 0, 0, 1))
        {
            _ = TranslateMessage(ref message);
            _ = DispatchMessage(ref message);
        }
    }

    [DllImport("user32.dll")]
    private static extern bool PeekMessage(out Message message, IntPtr window, uint min, uint max, uint remove);

    [DllImport("user32.dll")]
    private static extern bool TranslateMessage(ref Message message);

    [DllImport("user32.dll")]
    private static extern IntPtr DispatchMessage(ref Message message);

    [StructLayout(LayoutKind.Sequential)]
    private struct Message
    {
        public IntPtr Hwnd;
        public uint MessageId;
        public IntPtr WParam;
        public IntPtr LParam;
        public uint Time;
        public int PointX;
        public int PointY;
        public uint Private;
    }

    private interface IWorkItem { void Run(); }

    private sealed class WorkItem<T>(Func<T> callback, CancellationToken cancellationToken) : IWorkItem
    {
        private readonly TaskCompletionSource<T> _completion = new(TaskCreationOptions.RunContinuationsAsynchronously);

        public Task<T> Task => _completion.Task;

        public void Run()
        {
            if (cancellationToken.IsCancellationRequested)
            {
                _completion.TrySetCanceled(cancellationToken);
                return;
            }

            try { _completion.SetResult(callback()); }
            catch (Exception exception) { _completion.SetException(exception); }
        }
    }
}
