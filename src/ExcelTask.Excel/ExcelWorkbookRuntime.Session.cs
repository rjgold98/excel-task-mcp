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

        public bool HasExternalTargetOpen(string targetPath) =>
            RotWorkbookLocator.HasExternalWorkbookAtPath(targetPath, GetApplicationHwnd(Application));

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
            var beforeStart = OwnedExcelProcess.SnapshotExcelProcesses();
            var app = CreateApplication();
            try
            {
                var ownedProcess = OwnedExcelProcess.CaptureNew(app, beforeStart);
                observer.OnOwnedProcessCaptured(ownedProcess.Identity);
                ConfigureOwnedApplication(app);
                var workbooks = Get(app, "Workbooks");
                try
                {
                    var workbook = OpenWorkbook(workbooks, path, readOnly: true);
                    return new ExcelSession(app, workbook, workbook, ownsApplication: true, closeTarget: true, closeReference: false, ownedProcess);
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

        private static void TryQuit(object application)
        {
            try { Invoke(application, "Quit"); }
            catch (COMException) { }
        }
    }
}
