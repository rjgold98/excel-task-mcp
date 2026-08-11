using System.Globalization;
using System.Runtime.ExceptionServices;
using System.Reflection;
using System.Runtime.InteropServices;
using ExcelTask.Core;

namespace ExcelTask.Excel;

public sealed partial class ExcelWorkbookRuntime
{
    private const string RunShimSuccess = "OK";
    private const string RunShimFailure = "ERR";
    private const char RunShimSeparator = '';
    private const string RunShimFailurePrefix = RunShimFailure + "";
    private const int MaxRunDetailLength = 160;
    private const int VbextCtStdModule = 1;
    private const int VbextPpNone = 0;
    private const int VbextPkProc = 0;
    private const int XlOpenXmlWorkbookMacroEnabled = 52;

    private static WorkbookExecutionOutcome ExecuteMacroCore(ExcelTaskPlan plan, IExcelWorkbookRuntimeObserver observer)
    {
        var operation = plan.Request.Operation.EditMacroProcedure!;
        var checks = new List<TaskCheck>();
        var changes = new List<TaskChange>();
        ExcelSession? session = null;
        string? stagingPath = null;
        MacroProcedureSnapshot? snapshot = null;
        var mutationAttempted = false;
        var verified = false;
        var runCompleted = false;
        string? runFailureDetail = null;

        if (!IsIsolatedMacroCopyRequest(plan.Request))
        {
            return new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected,
                "Macro edits require an isolated .xlsm copy save.",
                Checks: [new TaskCheck("macro-isolation", false, "Macro edits require an isolated .xlsm copy save.")]);
        }

        try
        {
            WorkbookRuntimeHelpers.EnsureReadableWorkbook(WorkbookRuntimeHelpers.NormalizePath(plan.Request.TargetWorkbookPath), "Target workbook");
            WorkbookRuntimeHelpers.EnsureWritableCopyOutput(plan.Request.OutputWorkbookPath);
        }
        catch (InvalidOperationException)
        {
            return new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected,
                "Macro workbook inputs cannot be safely executed.",
                Checks: [new TaskCheck("macro-workbook-inputs", false, "Macro workbook paths or output location are not ready for execution.")]);
        }

        var savedPath = WorkbookRuntimeHelpers.NormalizePath(plan.Request.OutputWorkbookPath!);
        try
        {
            if (plan.Request.Mode == ExcelTaskMode.Apply && File.Exists(savedPath) && !plan.Request.OverwriteConfirmed)
            {
                return new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected,
                    "The copy output already exists and was not authorized for overwrite.",
                    Checks: [new TaskCheck("copy-output", false, "Existing output requires overwrite confirmation.")]);
            }

            observer.OnPhase("session-open");
            session = ExcelSession.Open(
                plan.Request,
                observer,
                readOnlyTarget: plan.Request.Mode == ExcelTaskMode.Plan,
                enableMacros: plan.Request.Mode == ExcelTaskMode.Apply && operation.RunAfterEdit);
            observer.OnPhase("macro-preflight");
            if (!TryGetVbaSigned(session.TargetWorkbook, out var vbaSigned))
            {
                checks.Add(new TaskCheck("macro-procedure-access", false, "The selected macro procedure cannot be safely accessed."));
                return CloseRejectedMacroSession(ref session, checks);
            }
            if (vbaSigned)
            {
                checks.Add(new TaskCheck("vba-signature", false, "Signed VBA projects cannot be edited."));
                return CloseRejectedMacroSession(ref session, checks);
            }
            if (!TryReadMacroProcedure(session.TargetWorkbook, operation.ComponentName, operation.ProcedureName, out snapshot, out var preflightCheck))
            {
                checks.Add(preflightCheck);
                return CloseRejectedMacroSession(ref session, checks);
            }

            checks.Add(preflightCheck);
            var currentSnapshot = snapshot!;
            if (plan.Request.Mode == ExcelTaskMode.Plan)
            {
                return ExcelSession.CloseAndProve(ref session, "macro planning", checks)
                    ?? new WorkbookExecutionOutcome(ExcelTaskStatus.Planned,
                        "Macro procedure plan is feasible; no Excel changes were made.",
                        Checks: checks,
                        MacroProcedure: CreateMacroReceipt(operation, currentSnapshot, includeSource: true, runCompleted: false));
            }

            if (!string.Equals(operation.ExpectedProcedureSha256, currentSnapshot.Sha256, StringComparison.Ordinal))
            {
                checks.Add(new TaskCheck("macro-hash", false, "The selected macro procedure changed before replacement; no changes were made."));
                return CloseRejectedMacroSession(ref session, checks, CreateMacroReceipt(operation, currentSnapshot, includeSource: false, runCompleted: false));
            }

            observer.OnPhase("macro-revalidation");
            if (!TryReadMacroProcedure(session.TargetWorkbook, operation.ComponentName, operation.ProcedureName, out var revalidated, out var revalidationCheck) ||
                !string.Equals(currentSnapshot.Sha256, revalidated!.Sha256, StringComparison.Ordinal))
            {
                checks.Add(new TaskCheck("macro-revalidation", false, "Macro procedure evidence changed before mutation; no changes were made."));
                return CloseRejectedMacroSession(ref session, checks, CreateMacroReceipt(operation, currentSnapshot, includeSource: false, runCompleted: false));
            }

            observer.OnPhase("macro-replace");
            mutationAttempted = true;
            ReplaceMacroProcedure(session.TargetWorkbook, operation.ComponentName, revalidated!, operation.ReplacementSource!);
            if (!TryReadMacroProcedure(session.TargetWorkbook, operation.ComponentName, operation.ProcedureName, out var replaced, out _) ||
                !string.Equals(replaced!.Sha256, MacroProcedureText.ComputeSha256(operation.ReplacementSource!), StringComparison.Ordinal))
            {
                throw new InvalidOperationException("Macro replacement verification failed.");
            }

            snapshot = replaced;
            changes.Add(new TaskChange("macro-procedure", $"{operation.ComponentName}.{operation.ProcedureName}", "Replaced the requested macro procedure."));
            checks.Add(new TaskCheck("macro-replacement", true, "The replacement procedure was verified in memory."));

            if (operation.RunAfterEdit)
            {
                observer.OnPhase("macro-run");
                var run = RunProcedureWithErrorTrap(session, operation, plan.TaskId);
                runCompleted = run.Ok;
                runFailureDetail = run.Ok ? null : run.Detail;
                checks.Add(new TaskCheck("macro-run", run.Ok, run.Detail));
            }

            observer.OnPhase("save");
            stagingPath = WorkbookRuntimeHelpers.CreateStagingPath(savedPath, plan.TaskId);
            observer.OnStagingPathCreated(stagingPath);
            Invoke(session.TargetWorkbook, "SaveAs", stagingPath, XlOpenXmlWorkbookMacroEnabled);
            checks.Add(new TaskCheck("save", true, "Excel completed the requested macro copy save."));

            observer.OnPhase("primary-cleanup");
            var saveCleanupFailure = ExcelSession.CloseAndProve(
                ref session, "the macro save", checks, changes, CreateMacroReceipt(operation, snapshot, includeSource: false, runCompleted));
            if (saveCleanupFailure is not null)
            {
                AddStagingCleanupCheck(stagingPath, checks);
                return saveCleanupFailure with { RetryReason = "Inspect workbook state before retrying." };
            }

            if (!WorkbookRuntimeHelpers.CanOpenExclusively(stagingPath))
            {
                checks.Add(new TaskCheck("file-lock", false, "The saved macro workbook remained locked after owned Excel cleanup."));
                AddStagingCleanupCheck(stagingPath, checks);
                return new WorkbookExecutionOutcome(ExcelTaskStatus.Unknown, "Macro changes could not be verified because the staging workbook remained locked.", changes, checks, CanRetry: false, RetryReason: "Inspect workbook state before retrying.", MacroProcedure: CreateMacroReceipt(operation, snapshot, includeSource: false, runCompleted));
            }

            observer.OnPhase("reopen-verification");
            if (!VerifySavedMacroProcedure(stagingPath, operation, snapshot.Sha256, observer, out var reopenCheck))
            {
                checks.Add(reopenCheck);
                AddStagingCleanupCheck(stagingPath, checks);
                return new WorkbookExecutionOutcome(ExcelTaskStatus.Unknown, "Saved macro procedure could not be verified after reopening.", changes, checks, CanRetry: false, RetryReason: "Inspect workbook state before retrying.", MacroProcedure: CreateMacroReceipt(operation, snapshot, includeSource: false, runCompleted));
            }

            checks.Add(reopenCheck);
            verified = true;
            observer.OnPhase("copy-promotion");
            WorkbookRuntimeHelpers.PromoteStaging(stagingPath, savedPath, plan.Request.OverwriteConfirmed);
            stagingPath = null;
            changes.Add(new TaskChange("copy-promotion", "workbook", "Promoted the verified staging workbook to the requested output path."));
            // The edit itself is saved and verified either way. A trapped run error is a known,
            // complete outcome rather than an uncertain one, so it reports Partial and names the
            // VBA error instead of leaving the caller to guess.
            return runFailureDetail is null
                ? new WorkbookExecutionOutcome(ExcelTaskStatus.Completed, "Macro procedure changes were saved and verified after reopening.", changes, checks, MacroProcedure: CreateMacroReceipt(operation, snapshot, includeSource: false, runCompleted))
                : new WorkbookExecutionOutcome(ExcelTaskStatus.Partial, "Macro procedure changes were saved and verified, but the requested run did not succeed.", changes, checks, CanRetry: false, RetryReason: runFailureDetail, MacroProcedure: CreateMacroReceipt(operation, snapshot, includeSource: false, runCompleted));
        }
        catch (MacroCompilationException exception)
        {
            // The replacement never ran and nothing reached disk: the edit existed only inside the
            // Excel instance that has now ended, and the target workbook was never written. So this
            // is an ordinary rejection the caller can fix and retry, not an uncertain outcome.
            //
            // The session is abandoned rather than closed. Its process was already terminated to
            // break the blocked call, so Quit and Close would be sent to a server that is gone and
            // its proxies then released - which can fault hard enough to take this process down.
            var exited = session?.AbandonTerminatedProcess() ?? true;
            session = null;

            checks.Add(new TaskCheck("macro-run", false, $"The macro did not compile, so it was not run: {exception.Message}"));
            checks.Add(new TaskCheck("owned-process-exit", exited, exited
                ? "The isolated Excel instance was ended to release the blocked call."
                : "The isolated Excel instance was ended, but its process has not exited yet."));
            return exited
                ? new WorkbookExecutionOutcome(
                    ExcelTaskStatus.Rejected,
                    "The replacement macro did not compile, so no changes were saved.",
                    Checks: checks,
                    MacroProcedure: snapshot is null ? null : CreateMacroReceipt(operation, snapshot, includeSource: false, runCompleted: false))
                : new WorkbookExecutionOutcome(
                    ExcelTaskStatus.Unknown,
                    "The replacement macro did not compile, and owned Excel cleanup could not be proven.",
                    Checks: checks,
                    CanRetry: false,
                    RetryReason: "Inspect the owned Excel process before retrying.",
                    MacroProcedure: snapshot is null ? null : CreateMacroReceipt(operation, snapshot, includeSource: false, runCompleted: false));
        }
        catch (Exception)
        {
            var ownedCleanupFailed = false;
            if (session is not null)
            {
                try { ownedCleanupFailed = !session.Close(); }
                catch (Exception cleanupException) when (cleanupException is COMException or InvalidOperationException or TargetInvocationException) { ownedCleanupFailed = true; }
                session = null;
            }

            var stagingCleanupFailed = stagingPath is not null && !WorkbookRuntimeHelpers.TryDeleteStaging(stagingPath);
            if (!stagingCleanupFailed) stagingPath = null;
            var status = ownedCleanupFailed || stagingCleanupFailed ? ExcelTaskStatus.Unknown : verified ? ExcelTaskStatus.Partial : mutationAttempted ? ExcelTaskStatus.Unknown : ExcelTaskStatus.Rejected;
            checks.Add(new TaskCheck("macro-execution", false, mutationAttempted
                ? "Macro execution did not complete during the current phase; the change was not fully verified."
                : "Macro procedure preflight did not complete."));
            if (ownedCleanupFailed) checks.Add(new TaskCheck("owned-process-exit", false, "The owned Excel process did not exit during macro cleanup."));
            if (stagingCleanupFailed) checks.Add(new TaskCheck("staging-cleanup", false, "A macro staging workbook could not be deleted; inspect the output directory before retrying."));
            return new WorkbookExecutionOutcome(status,
                verified ? "Macro changes were verified, but final delivery did not complete." : mutationAttempted ? "Excel attempted macro changes, but their final state is unknown." : "Macro execution was rejected before changes were attempted.",
                changes, checks, CanRetry: false, RetryReason: mutationAttempted ? "Inspect workbook state before retrying." : "Correct the macro preflight issue before retrying.",
                MacroProcedure: snapshot is null ? null : CreateMacroReceipt(operation, snapshot, includeSource: false, runCompleted));
        }
        finally
        {
            session?.Close();
            if (stagingPath is not null) _ = WorkbookRuntimeHelpers.TryDeleteStaging(stagingPath);
        }
    }

    private static bool IsIsolatedMacroCopyRequest(NormalizedExcelTaskRequest request) =>
        request.WorkbookBinding == WorkbookBinding.Isolated &&
        request.Save == SaveMode.Copy &&
        string.Equals(Path.GetExtension(request.TargetWorkbookPath), ".xlsm", StringComparison.OrdinalIgnoreCase) &&
        string.Equals(Path.GetExtension(request.OutputWorkbookPath), ".xlsm", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Closes a session that preflight refused, proving cleanup the same way every other exit does.
    /// Takes the session by reference so the caller's field is nulled here rather than being left
    /// for a later close to run against a process that is already gone.
    /// </summary>
    private static WorkbookExecutionOutcome CloseRejectedMacroSession(ref ExcelSession? session, List<TaskCheck> checks, MacroProcedureReceipt? receipt = null) =>
        ExcelSession.CloseAndProve(ref session, "macro preflight", checks, macroProcedure: receipt)
        ?? new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected, "Macro procedure preflight did not permit the requested operation.", Checks: checks, MacroProcedure: receipt);

    private static bool TryReadMacroProcedure(object workbook, string componentName, string procedureName, out MacroProcedureSnapshot? snapshot, out TaskCheck check)
    {
        snapshot = null;
        // Each step reports its own reason so an operator can tell a Trust Center block apart
        // from a missing module, a locked project, or an oversized procedure.
        var reason = "Excel did not permit programmatic access to the workbook's VBA project. Enable Trust Center > Macro Settings > Trust access to the VBA project object model.";
        try
        {
            using var references = new ComReferenceScope();
            var project = references.Add(Get(workbook, "VBProject"));

            reason = "The workbook's VBA project is locked or password-protected and cannot be inspected.";
            if (Convert.ToInt32(Get(project, "Protection"), CultureInfo.InvariantCulture) != VbextPpNone) throw new InvalidOperationException();

            reason = "The requested VBA component does not exist in the workbook.";
            var components = references.Add(Get(project, "VBComponents"));
            var component = references.Add(Item(components, componentName));

            reason = "The requested VBA component is not a standard module; only standard modules can be inspected or edited.";
            if (Convert.ToInt32(Get(component, "Type"), CultureInfo.InvariantCulture) != VbextCtStdModule) throw new InvalidOperationException();

            reason = "The requested procedure does not exist as a Sub or Function in that component.";
            var module = references.Add(Get(component, "CodeModule"));
            // ProcStartLine, ProcCountLines and Lines are parameterized properties on the VBIDE
            // CodeModule interface, not methods; InvokeMethod binding raises DISP_E_MEMBERNOTFOUND.
            var startLine = Convert.ToInt32(Get(module, "ProcStartLine", procedureName, VbextPkProc), CultureInfo.InvariantCulture);
            var lineCount = Convert.ToInt32(Get(module, "ProcCountLines", procedureName, VbextPkProc), CultureInfo.InvariantCulture);
            var source = MacroProcedureText.NormalizeLineEndings((string)Get(module, "Lines", startLine, lineCount));

            reason = $"The requested procedure exceeds the bounded limits of {MacroProcedureText.MaxSourceCharacters:N0} characters and {MacroProcedureText.MaxLineCount} lines.";
            if (source.Length > MacroProcedureText.MaxSourceCharacters || CountLines(source) > MacroProcedureText.MaxLineCount) throw new InvalidOperationException();

            snapshot = new MacroProcedureSnapshot(startLine, lineCount, source, MacroProcedureText.ComputeSha256(source));
            check = new TaskCheck("macro-procedure-access", true, "The selected standard-module macro procedure is available.");
            return true;
        }
        catch (Exception exception) when (IsExpectedMacroAccessFailure(exception))
        {
            snapshot = null;
            check = new TaskCheck("macro-procedure-access", false, reason);
            return false;
        }
    }

    private static int CountLines(string source) => source.Length == 0 ? 0 : source.Count(character => character == '\n') + 1;

    private static void ReplaceMacroProcedure(object workbook, string componentName, MacroProcedureSnapshot existing, string replacementSource)
    {
        using var references = new ComReferenceScope();
        var project = references.Add(Get(workbook, "VBProject"));
        var components = references.Add(Get(project, "VBComponents"));
        var component = references.Add(Item(components, componentName));
        var module = references.Add(Get(component, "CodeModule"));
        Invoke(module, "DeleteLines", existing.StartLine, existing.LineCount);
        Invoke(module, "InsertLines", existing.StartLine, MacroProcedureText.NormalizeLineEndings(replacementSource).Replace("\n", "\r\n", StringComparison.Ordinal));
    }

    private static string QualifiedMember(object workbook, string componentName, string memberName) =>
        $"'{((string)Get(workbook, "Name")).Replace("'", "''", StringComparison.Ordinal)}'!{componentName}.{memberName}";

    /// <summary>
    /// Runs the edited procedure through a temporary generated wrapper that traps VBA errors.
    /// Calling the procedure directly lets an unhandled error raise a modal dialog inside the owned
    /// Excel instance, which blocks the automation thread until a watchdog kills the process and
    /// costs the caller a useless uncertain result. The wrapper turns that into a returned string.
    /// The wrapper is appended after the module's existing lines so no procedure shifts position,
    /// and it is always removed again before the workbook is saved.
    /// </summary>
    private static MacroRunResult RunProcedureWithErrorTrap(ExcelSession session, NormalizedEditMacroProcedureOperation operation, string taskId)
    {
        using var references = new ComReferenceScope();
        var project = references.Add(Get(session.TargetWorkbook, "VBProject"));
        var components = references.Add(Get(project, "VBComponents"));
        var component = references.Add(Item(components, operation.ComponentName));
        var module = references.Add(Get(component, "CodeModule"));

        var shimName = CreateShimName(taskId);
        var linesBefore = Convert.ToInt32(Get(module, "CountOfLines"), CultureInfo.InvariantCulture);
        var existing = linesBefore == 0 ? string.Empty : (string)Get(module, "Lines", 1, linesBefore);
        if (existing.Contains(shimName, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException("The temporary run helper name is already used in the module.");
        }

        Invoke(module, "InsertLines", linesBefore + 1, BuildRunShim(shimName, operation.ProcedureName));
        string? raw = null;
        Exception? runFailure = null;
        IReadOnlyList<DismissedDialog> dismissed;
        // The wrapper cannot catch a compile error, because VBA refuses to compile the module before
        // any handler is installed, and it cannot catch a MsgBox in a procedure the replacement
        // calls. Both raise a modal dialog while this thread is blocked inside COM, so a sentry
        // watches the owned Excel process from outside and answers what it recognizes.
        using (var sentry = ModalDialogSentry.Watch(session.OwnedProcessIdentity))
        {
            try
            {
                raw = Invoke(session.Application, "Run", QualifiedMember(session.TargetWorkbook, operation.ComponentName, shimName)) as string;
            }
            catch (Exception exception)
            {
                runFailure = exception;
            }

            dismissed = sentry?.Dismissed ?? [];
        }

        // A compile error means the sentry had to end this Excel instance, so no further COM call
        // can succeed and there is nothing left to tidy up inside it. The references are abandoned
        // rather than released: releasing a proxy whose server has been terminated can fault hard
        // enough to take this process down, and there is nothing left on the other side to free.
        var compileError = dismissed.FirstOrDefault(dialog => dialog.Kind == ModalDialogKind.VbaCompileError);
        if (compileError is not null)
        {
            references.Abandon();
            throw new MacroCompilationException(compileError.Message);
        }

        var linesAfter = Convert.ToInt32(Get(module, "CountOfLines"), CultureInfo.InvariantCulture);
        if (linesAfter > linesBefore) Invoke(module, "DeleteLines", linesBefore + 1, linesAfter - linesBefore);

        // Saving a workbook that still carried the generated helper would hand the user a macro
        // project they did not write, so an incomplete removal fails the task instead.
        if (Convert.ToInt32(Get(module, "CountOfLines"), CultureInfo.InvariantCulture) != linesBefore)
        {
            throw new InvalidOperationException("The temporary run helper could not be removed from the module.");
        }

        // An answered dialog explains the failure far better than the COM error Excel then returns,
        // so it wins over both the exception and the wrapper's own result.
        if (dismissed.Count > 0) return DescribeDismissedDialogs(dismissed);
        if (runFailure is not null) ExceptionDispatchInfo.Capture(runFailure).Throw();

        return InterpretRunResult(raw);
    }

    private static MacroRunResult DescribeDismissedDialogs(IReadOnlyList<DismissedDialog> dismissed)
    {
        var first = dismissed[0];
        var detail = first.Kind switch
        {
            ModalDialogKind.VbaCompileError => $"The macro did not compile, so it was not run: {first.Message}",
            ModalDialogKind.VbaRuntimeError => $"The macro stopped and was ended: {first.Message}",
            _ => $"The macro displayed a message box, which was acknowledged so the task could finish: {first.Message}"
        };

        return new MacroRunResult(false, dismissed.Count == 1 ? detail : $"{detail} ({dismissed.Count} dialogs were answered.)");
    }

    private static MacroRunResult InterpretRunResult(string? raw)
    {
        if (string.Equals(raw, RunShimSuccess, StringComparison.Ordinal))
        {
            return new MacroRunResult(true, "The requested no-argument macro procedure completed.");
        }

        if (raw is not null && raw.StartsWith(RunShimFailurePrefix, StringComparison.Ordinal))
        {
            var parts = raw.Split(RunShimSeparator);
            var number = parts.Length > 1 ? BoundRunDetail(parts[1]) : "unknown";
            var description = parts.Length > 2 ? BoundRunDetail(parts[2]) : string.Empty;
            return new MacroRunResult(false, $"The macro stopped on VBA error {number}: {description}");
        }

        return new MacroRunResult(false, "The macro run returned an unrecognized result.");
    }

    private static string BoundRunDetail(string value)
    {
        var collapsed = new string(value.Select(character => char.IsControl(character) ? ' ' : character).ToArray()).Trim();
        return collapsed.Length <= MaxRunDetailLength ? collapsed : collapsed[..MaxRunDetailLength];
    }

    private static string CreateShimName(string taskId)
    {
        var cleaned = new string(taskId.Where(char.IsLetterOrDigit).Take(16).ToArray());
        return "ExcelTaskRun" + (cleaned.Length == 0 ? "Shim" : cleaned);
    }

    private static string BuildRunShim(string shimName, string procedureName) => string.Join("\r\n",
        $"Public Function {shimName}() As String",
        "    On Error GoTo ExcelTaskTrapped",
        $"    {procedureName}",
        $"    {shimName} = \"{RunShimSuccess}\"",
        "    Exit Function",
        "ExcelTaskTrapped:",
        $"    {shimName} = \"{RunShimFailure}\" & Chr$({(int)RunShimSeparator}) & CStr(Err.Number) & Chr$({(int)RunShimSeparator}) & Err.Description",
        "End Function");

    private readonly record struct MacroRunResult(bool Ok, string Detail);

    /// <summary>
    /// The replacement could not be compiled, so it never ran and the owned Excel instance had to be
    /// ended to release the blocked call. Nothing was saved when this is thrown.
    /// </summary>
    private sealed class MacroCompilationException(string message) : Exception(message);

    private static bool VerifySavedMacroProcedure(string path, NormalizedEditMacroProcedureOperation operation, string expectedHash, IExcelWorkbookRuntimeObserver observer, out TaskCheck check)
    {
        ExcelSession? verification = null;
        var contentVerified = false;
        try
        {
            verification = ExcelSession.OpenForVerification(path, observer);
            if (!TryReadMacroProcedure(verification.TargetWorkbook, operation.ComponentName, operation.ProcedureName, out var reopened, out _) ||
                !string.Equals(reopened!.Sha256, expectedHash, StringComparison.Ordinal))
            {
                check = new TaskCheck("reopen-verification", false, "The saved macro procedure was not present after reopening.");
                return false;
            }

            contentVerified = true;
            check = new TaskCheck("reopen-verification", true, "Saved macro workbook reopened with the requested procedure fingerprint.");
        }
        catch (Exception)
        {
            check = new TaskCheck("reopen-verification", false, "Saved macro procedure verification did not complete.");
            return false;
        }
        finally
        {
            if (verification is not null && !verification.Close())
            {
                contentVerified = false;
                check = new TaskCheck("verification-process-exit", false, "The owned verification Excel process did not exit.");
            }
        }

        return contentVerified;
    }

    private static bool TryGetVbaSigned(object workbook, out bool signed)
    {
        try
        {
            signed = Convert.ToBoolean(Get(workbook, "VBASigned"), CultureInfo.InvariantCulture);
            return true;
        }
        catch (Exception exception) when (IsExpectedMacroAccessFailure(exception))
        {
            signed = false;
            return false;
        }
    }

    private static bool IsExpectedMacroAccessFailure(Exception exception) => exception is COMException or TargetInvocationException or InvalidComObjectException or InvalidOperationException or ArgumentException or OverflowException;

    private static MacroProcedureReceipt CreateMacroReceipt(NormalizedEditMacroProcedureOperation operation, MacroProcedureSnapshot snapshot, bool includeSource, bool runCompleted) =>
        new(operation.ComponentName, operation.ProcedureName, snapshot.Sha256, includeSource ? snapshot.Source : null, operation.RunAfterEdit, runCompleted);

    private sealed record MacroProcedureSnapshot(int StartLine, int LineCount, string Source, string Sha256);
}
