using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using ExcelTask.Core;

namespace ExcelTask.Excel;

public sealed partial class ExcelWorkbookRuntime
{
    private sealed class ExcelSession : IDisposable
    {
        private readonly ComReferenceScope _references = new();
        private readonly bool _ownsApplication;
        private readonly bool _closeTarget;
        private readonly bool _closeReference;
        private readonly OwnedExcelProcess? _ownedProcess;
        private bool _closed;

        private ExcelSession(object application, object targetWorkbook, object referenceWorkbook, bool ownsApplication, bool closeTarget, bool closeReference, OwnedExcelProcess? ownedProcess)
        {
            Application = _references.Add(application);
            TargetWorkbook = _references.Add(targetWorkbook);
            ReferenceWorkbook = _references.Add(referenceWorkbook);
            _ownsApplication = ownsApplication;
            _closeTarget = closeTarget;
            _closeReference = closeReference;
            _ownedProcess = ownedProcess;
        }

        public object Application { get; }

        public object TargetWorkbook { get; }

        public object ReferenceWorkbook { get; }

        /// <summary>
        /// The Excel process this session created, or null when the session is bound to a workbook
        /// the user already had open. Only an owned process may ever be watched or acted upon.
        /// </summary>
        public ProcessIdentity? OwnedProcessIdentity => _ownedProcess?.Identity;

        public bool HasExternalTargetOpen(string targetPath) =>
            RotWorkbookLocator.HasExternalWorkbookAtPath(targetPath, GetApplicationHwnd(Application));

        /// <summary>
        /// Closes the session and turns a failure to prove owned Excel exited into the outcome the
        /// caller must return, or null when cleanup was proven and the caller may continue.
        ///
        /// This exists because the close-and-prove sequence was written out fourteen times across
        /// three operations, in nine different wordings, and every copy had to remember four
        /// separate things: capture the result, forget the session, add the failed check, and return
        /// Unknown as non-retryable. Forgetting any one of them silently weakens the guarantee that
        /// no Excel process is left behind, which is the product's central reliability claim.
        /// </summary>
        /// <summary>
        /// Gives this session up without touching Excel, for the one case where talking to it is
        /// unsafe: its process has already been terminated to break a blocked call. Closing would
        /// send Quit and Close to a server that no longer exists and then release its proxies,
        /// which can fault hard enough to take this process down. Returns whether the owned process
        /// is genuinely gone, so the caller still reports the truth rather than assuming it.
        /// </summary>
        public bool AbandonTerminatedProcess()
        {
            _closed = true;
            _references.Abandon();
            return _ownedProcess is null || !_ownedProcess.IsRunning;
        }

        /// <param name="session">Set to null so a later close cannot double-run against a dead process.</param>
        /// <param name="what">Names the moment in the operation, for the failure detail only.</param>
        public static WorkbookExecutionOutcome? CloseAndProve(
            ref ExcelSession? session,
            string what,
            List<TaskCheck> checks,
            IReadOnlyList<TaskChange>? changes = null,
            MacroProcedureReceipt? macroProcedure = null)
        {
            if (session is null) return null;

            var proven = session.Close();
            session = null;
            if (proven) return null;

            checks.Add(new TaskCheck("owned-process-exit", false, $"The owned Excel process did not exit after {what}."));
            return new WorkbookExecutionOutcome(
                ExcelTaskStatus.Unknown,
                $"ExcelTask could not prove owned Excel cleanup after {what}.",
                changes,
                checks,
                CanRetry: false,
                RetryReason: "Inspect the owned Excel process before retrying.",
                MacroProcedure: macroProcedure);
        }

        public static ExcelSession Open(NormalizedExcelTaskRequest request, IExcelWorkbookRuntimeObserver observer, bool readOnlyTarget = false, bool enableMacros = false)
        {
            var needsReference = NeedsReferenceWorkbook(request);
            var referencePath = request.Operation.CopyExhibit?.ReferenceWorkbookPath;
            if (request.WorkbookBinding == WorkbookBinding.UseOpen)
            {
                using var found = RotWorkbookLocator.Find(WorkbookRuntimeHelpers.NormalizePath(request.TargetWorkbookPath));
                if (found is null) throw new InvalidOperationException("The requested open target workbook was not found in the running object table.");
                var target = found.Detach();
                var application = Get(target, "Application");
                object? reference = null;
                var closeReference = false;
                try
                {
                    if (!needsReference)
                    {
                        return new ExcelSession(application, target, target, ownsApplication: false, closeTarget: false, closeReference: false, ownedProcess: null);
                    }

                    using var openReference = RotWorkbookLocator.Find(WorkbookRuntimeHelpers.NormalizePath(referencePath!));
                    if (openReference is not null)
                    {
                        var openReferenceApplication = Get(openReference.Workbook, "Application");
                        try
                        {
                            var sameApplication = AreSameApplication(application, openReferenceApplication);
                            if (sameApplication)
                            {
                                reference = openReference.Detach();
                            }

                            if (!sameApplication) ComReferences.Release(openReferenceApplication);
                        }
                        catch
                        {
                            ComReferences.Release(openReferenceApplication);
                            throw;
                        }
                    }

                    if (reference is null && WorkbookRuntimeHelpers.PathsEqual(request.TargetWorkbookPath, referencePath!))
                    {
                        reference = target;
                    }

                    if (reference is null)
                    {
                        var workbooks = Get(application, "Workbooks");
                        try
                        {
                            var priorAutomationSecurity = Get(application, "AutomationSecurity");
                            try
                            {
                                Set(application, "AutomationSecurity", WorkbookRuntimeHelpers.AutomationSecurityForceDisable);
                                reference = OpenWorkbook(workbooks, referencePath!, readOnly: true);
                            }
                            finally
                            {
                                Set(application, "AutomationSecurity", priorAutomationSecurity);
                            }
                            closeReference = true;
                        }
                        finally
                        {
                            ComReferences.Release(workbooks);
                        }
                    }

                    return new ExcelSession(application, target, reference!, ownsApplication: false, closeTarget: false, closeReference, ownedProcess: null);
                }
                catch
                {
                    ComReferences.Release(reference);
                    ComReferences.Release(target);
                    ComReferences.Release(application);
                    throw;
                }
            }

            var beforeStart = OwnedExcelProcess.SnapshotExcelProcesses();
            var app = CreateApplication();
            try
            {
                var ownedProcess = OwnedExcelProcess.CaptureNew(app, beforeStart);
                observer.OnOwnedProcessCaptured(ownedProcess.Identity);
                ConfigureOwnedApplication(app, enableMacros);
                var workbooks = Get(app, "Workbooks");
                try
                {
                    var target = OpenWorkbook(workbooks, request.TargetWorkbookPath, readOnly: readOnlyTarget);
                    object? reference = null;
                    try
                    {
                        reference = !needsReference || WorkbookRuntimeHelpers.PathsEqual(request.TargetWorkbookPath, referencePath!)
                            ? target
                            : OpenWorkbook(workbooks, referencePath!, readOnly: true);
                        return new ExcelSession(app, target, reference, ownsApplication: true, closeTarget: true, closeReference: !ReferenceEquals(target, reference), ownedProcess);
                    }
                    catch
                    {
                        if (!ReferenceEquals(reference, target)) ComReferences.Release(reference);
                        ComReferences.Release(target);
                        throw;
                    }
                }
                finally
                {
                    ComReferences.Release(workbooks);
                }
            }
            catch
            {
                if (OwnedExcelProcess.IsNewlyOwned(app, beforeStart)) TryQuit(app);
                ComReferences.Release(app);
                throw;
            }
        }

        public static ExcelSession OpenForVerification(string path, IExcelWorkbookRuntimeObserver observer)
        {
            var prepared = PrepareForVerification(observer);
            try
            {
                return AttachForVerification(prepared, path);
            }
            catch
            {
                prepared.Abandon();
                throw;
            }
        }

        /// <summary>
        /// The expensive half of a verification session: start Excel and configure it, with no
        /// workbook open. Split out so it can be run early, while the primary session is still
        /// writing, rather than after it closes. Measured at 274-482 ms - on the macro path, 31% of
        /// all Excel time - against 6 ms for everything the verification then does.
        /// </summary>
        public static PreparedVerificationExcel PrepareForVerification(IExcelWorkbookRuntimeObserver observer)
        {
            var beforeStart = OwnedExcelProcess.SnapshotExcelProcesses();
            var app = CreateApplication();
            try
            {
                var ownedProcess = OwnedExcelProcess.CaptureNew(app, beforeStart);
                // Reported the moment it exists, not when it is first used. An owned process the
                // supervisor has not been told about is one it cannot clean up.
                observer.OnOwnedProcessCaptured(ownedProcess.Identity);
                ConfigureOwnedApplication(app);
                return new PreparedVerificationExcel(app, ownedProcess);
            }
            catch
            {
                if (OwnedExcelProcess.IsNewlyOwned(app, beforeStart)) TryQuit(app);
                ComReferences.Release(app);
                throw;
            }
        }

        /// <summary>The cheap half: open the saved file read-only on an instance already running.</summary>
        public static ExcelSession AttachForVerification(PreparedVerificationExcel prepared, string path)
        {
            var workbooks = Get(prepared.Application, "Workbooks");
            try
            {
                var workbook = OpenWorkbook(workbooks, path, readOnly: true);
                return new ExcelSession(prepared.Application, workbook, workbook, ownsApplication: true, closeTarget: true, closeReference: false, prepared.OwnedProcess);
            }
            finally
            {
                ComReferences.Release(workbooks);
            }
        }

        public bool Close()
        {
            if (_closed) return _ownedProcess is null || !_ownedProcess.IsRunning;
            _closed = true;
            var ownedProcessExited = true;
            try
            {
                try
                {
                    if (_closeReference && !ReferenceEquals(ReferenceWorkbook, TargetWorkbook)) Invoke(ReferenceWorkbook, "Close", false);
                }
                finally
                {
                    try
                    {
                        if (_closeTarget) Invoke(TargetWorkbook, "Close", false);
                    }
                    finally
                    {
                        if (_ownsApplication) Invoke(Application, "Quit");
                    }
                }
            }
            finally
            {
                _references.Dispose();
                if (_ownedProcess is not null)
                {
                    ownedProcessExited = _ownedProcess.WaitForExitOrTerminate();
                }
            }

            return ownedProcessExited;
        }

        public void Dispose() => _ = Close();

        private static object CreateApplication()
        {
            var excelType = Type.GetTypeFromProgID("Excel.Application", throwOnError: true)
                ?? throw new InvalidOperationException("Microsoft Excel is not registered on this machine.");
            return Activator.CreateInstance(excelType) ?? throw new InvalidOperationException("Microsoft Excel could not be started.");
        }

        private static void ConfigureOwnedApplication(object application, bool enableMacros = false)
        {
            Set(application, "Visible", false);
            Set(application, "DisplayAlerts", false);
            Set(application, "EnableEvents", false);
            // Macros stay force-disabled unless the request explicitly asked to run one. Even then
            // EnableEvents stays false, so Workbook_Open cannot fire, and a programmatic Open never
            // executes Auto_Open. AutomationSecurity only takes effect for workbooks opened after it
            // is set, so this must happen before Workbooks.Open.
            Set(application, "AutomationSecurity", enableMacros
                ? WorkbookRuntimeHelpers.AutomationSecurityLow
                : WorkbookRuntimeHelpers.AutomationSecurityForceDisable);
        }

        private static object OpenWorkbook(object workbooks, string path, bool readOnly) => Invoke(
            workbooks,
            "Open",
            WorkbookRuntimeHelpers.NormalizePath(path),
            0,
            readOnly,
            Type.Missing,
            Type.Missing,
            Type.Missing,
            true,
            Type.Missing,
            Type.Missing,
            Type.Missing,
            false,
            false,
            Type.Missing,
            false,
            0) ?? throw new InvalidOperationException("Excel did not open the workbook.");

        private static bool AreSameApplication(object left, object right) =>
            GetApplicationHwnd(left) == GetApplicationHwnd(right);

        private static long GetApplicationHwnd(object application) =>
            Convert.ToInt64(Get(application, "Hwnd"), CultureInfo.InvariantCulture);

        internal static void TryQuit(object application)
        {
            try { Invoke(application, "Quit"); }
            catch (COMException) { }
        }
    }
}
