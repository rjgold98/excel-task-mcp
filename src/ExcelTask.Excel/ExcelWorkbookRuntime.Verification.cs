using System.Runtime.InteropServices;
using ExcelTask.Core;

namespace ExcelTask.Excel;

public sealed partial class ExcelWorkbookRuntime
{
    /// <summary>
    /// A verification Excel that has been started but has no workbook open yet.
    /// </summary>
    /// <remarks>
    /// Held separately from <see cref="ExcelSession"/> because between preparing and attaching there
    /// is a window where the instance is owned, running, and holding nothing - and something has to
    /// be able to shut it down in that state.
    /// </remarks>
    internal sealed class PreparedVerificationExcel(object application, OwnedExcelProcess ownedProcess)
    {
        public object Application { get; } = application;

        public OwnedExcelProcess OwnedProcess { get; } = ownedProcess;

        /// <summary>
        /// Shuts down an instance that will never be used, with the same escalation the primary
        /// session gets: Quit, wait, and terminate by identity if Excel ignores it. It used to stop
        /// at a best-effort Quit and then sample liveness immediately, which meant an Excel holding
        /// anything - a dialog, a slow teardown - survived the one call whose whole job is that it
        /// does not. Returns whether the process is genuinely gone, proven the same way Close proves
        /// it, so a caller reports the truth rather than assuming it.
        /// </summary>
        public bool Abandon()
        {
            try { ExcelSession.TryQuit(Application); }
            catch (Exception exception) when (ComAccess.IsComFailure(exception)) { }
            ComAccess.Release(Application);
            return OwnedProcess.WaitForExitOrTerminate();
        }
    }

    /// <summary>
    /// Starts the verification Excel while the primary session is still working, so its launch
    /// overlaps the writes and the save instead of following them.
    ///
    /// Every property of the verification stays exactly as it was: a separate process, freshly
    /// launched, opening the saved file only after the primary closed and released its lock. What
    /// changes is only when the clock on starting it begins. That launch was measured at 274-482 ms
    /// - on the macro path, 31% of all Excel time, against 6 ms for the reading it then does.
    ///
    /// It runs on its own STA thread with its own message pump, because that is what Excel COM
    /// requires and the primary instance already owns the runtime's. The two instances never touch
    /// each other, and no Excel object ever crosses between the threads: the whole verification -
    /// open, read, close, prove - runs on this one.
    /// </summary>
    private sealed class PendingVerification : IDisposable
    {
        // Assigned only by the constructor, from the dispatcher Begin already created. An initializer
        // here ran first and was then overwritten, so every Apply started two STA threads and
        // orphaned one: never disposed, never sent CompleteAdding, parked forever in
        // GetConsumingEnumerable. In the module whose whole job is proving nothing is left running.
        private readonly StaComDispatcher _dispatcher;
        private readonly Task<PreparedVerificationExcel> _prepared;
        private bool _consumed;
        private bool _disposed;

        private PendingVerification(StaComDispatcher dispatcher, Task<PreparedVerificationExcel> prepared)
        {
            _dispatcher = dispatcher;
            _prepared = prepared;
        }

        public static PendingVerification Begin(IExcelWorkbookRuntimeObserver observer)
        {
            var dispatcher = new StaComDispatcher();
            try
            {
                return new PendingVerification(
                    dispatcher,
                    dispatcher.InvokeAsync(() => ExcelSession.PrepareForVerification(observer), CancellationToken.None));
            }
            catch
            {
                dispatcher.Dispose();
                throw;
            }
        }

        /// <summary>
        /// Runs the verification body against the saved file and closes the instance, all on the
        /// verification thread. If starting Excel early failed, this falls back to starting one now,
        /// so a pre-launch that does not happen costs latency rather than the whole operation.
        /// </summary>
        public bool Verify(
            string path,
            IExcelWorkbookRuntimeObserver observer,
            Func<ExcelSession, (bool Verified, TaskCheck Check)> body,
            out TaskCheck check)
        {
            _consumed = true;
            var outcome = _dispatcher.InvokeAsync(() => RunVerification(path, observer, body), CancellationToken.None)
                .GetAwaiter()
                .GetResult();
            check = outcome.Check;
            return outcome.Verified;
        }

        private (bool Verified, TaskCheck Check) RunVerification(
            string path,
            IExcelWorkbookRuntimeObserver observer,
            Func<ExcelSession, (bool Verified, TaskCheck Check)> body)
        {
            var prepared = ResolvePreparedOrStartNow(observer);
            ExcelSession? session = null;
            (bool Verified, TaskCheck Check) outcome;
            try
            {
                session = ExcelSession.AttachForVerification(prepared, path);
                outcome = body(session);
            }
            catch
            {
                if (session is null)
                {
                    // Attach failed, so the instance is running with nothing open - and nothing else
                    // will ever shut it down. Verify marks this object consumed on entry, which means
                    // Dispose skips its abandon branch, so this catch is the last hand that can reach
                    // the process. It leaked here once: Excel started, Workbooks.Open threw, and the
                    // session-null close below was a no-op while Dispose assumed Verify had cleaned up.
                    _ = prepared.Abandon();
                }
                else
                {
                    try { _ = session.Close(); }
                    catch (Exception exception) when (ComAccess.IsComFailure(exception)) { }
                }

                throw;
            }

            // A verification that read the right thing but could not prove its own Excel exited is
            // not a pass. The instance is shut down on the thread that owns it, on every path out.
            return session.Close()
                ? outcome
                : (false, new TaskCheck("verification-process-exit", false, "The owned verification Excel process did not exit."));
        }



        /// <summary>
        /// The pre-launched instance, or a freshly started one if launching early failed. A
        /// pre-launch that does not happen must cost latency, not the operation.
        /// </summary>
        private PreparedVerificationExcel ResolvePreparedOrStartNow(IExcelWorkbookRuntimeObserver observer)
        {
            try
            {
                return _prepared.GetAwaiter().GetResult();
            }
            catch (Exception exception) when (ComAccess.IsComFailure(exception))
            {
                return ExcelSession.PrepareForVerification(observer);
            }
        }

        /// <summary>
        /// Shuts down an instance that was never used - the early returns, the rejections, the
        /// exceptions. This is the whole risk the pre-launch introduces, so it is one place rather
        /// than one per exit path, and it runs on the verification thread like everything else.
        /// </summary>
        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            try
            {
                if (!_consumed)
                {
                    _dispatcher.InvokeAsync(() =>
                    {
                        try { return _prepared.GetAwaiter().GetResult().Abandon(); }
                        catch (Exception exception) when (ComAccess.IsComFailure(exception)) { return true; }
                    }, CancellationToken.None).GetAwaiter().GetResult();
                }
            }
            catch (Exception exception) when (exception is InvalidOperationException or ObjectDisposedException) { }
            finally
            {
                _dispatcher.Dispose();
            }
        }
    }
}
