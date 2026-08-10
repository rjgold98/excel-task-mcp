using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using ExcelTask.Core;

namespace ExcelTask.Excel;

public sealed partial class ExcelWorkbookRuntime
{
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
            session = ExcelSession.Open(plan.Request, observer, readOnlyTarget: plan.Request.Mode == ExcelTaskMode.Plan);
            observer.OnPhase("macro-preflight");
            if (!TryGetVbaSigned(session.TargetWorkbook, out var vbaSigned))
            {
                checks.Add(new TaskCheck("macro-procedure-access", false, "The selected macro procedure cannot be safely accessed."));
                return CloseRejectedMacroSession(session, checks);
            }
            if (vbaSigned)
            {
                checks.Add(new TaskCheck("vba-signature", false, "Signed VBA projects cannot be edited."));
                return CloseRejectedMacroSession(session, checks);
            }
            if (!TryReadMacroProcedure(session.TargetWorkbook, operation.ComponentName, operation.ProcedureName, out snapshot, out var preflightCheck))
            {
                checks.Add(preflightCheck);
                return CloseRejectedMacroSession(session, checks);
            }

            checks.Add(preflightCheck);
            var currentSnapshot = snapshot!;
            if (plan.Request.Mode == ExcelTaskMode.Plan)
            {
                var cleanupVerified = session.Close();
                session = null;
                if (!cleanupVerified)
                {
                    checks.Add(new TaskCheck("owned-process-exit", false, "The owned macro planning Excel process did not exit."));
                    return new WorkbookExecutionOutcome(ExcelTaskStatus.Unknown, "Macro plan could not prove owned Excel cleanup.", Checks: checks, CanRetry: false, RetryReason: "Inspect the owned Excel process before retrying.");
                }

                return new WorkbookExecutionOutcome(ExcelTaskStatus.Planned,
                    "Macro procedure plan is feasible; no Excel changes were made.",
                    Checks: checks,
                    MacroProcedure: CreateMacroReceipt(operation, currentSnapshot, includeSource: true, runCompleted: false));
            }

            if (!string.Equals(operation.ExpectedProcedureSha256, currentSnapshot.Sha256, StringComparison.Ordinal))
            {
                checks.Add(new TaskCheck("macro-hash", false, "The selected macro procedure changed before replacement; no changes were made."));
                return CloseRejectedMacroSession(session, checks, CreateMacroReceipt(operation, currentSnapshot, includeSource: false, runCompleted: false));
            }

            observer.OnPhase("macro-revalidation");
            if (!TryReadMacroProcedure(session.TargetWorkbook, operation.ComponentName, operation.ProcedureName, out var revalidated, out var revalidationCheck) ||
                !string.Equals(currentSnapshot.Sha256, revalidated!.Sha256, StringComparison.Ordinal))
            {
                checks.Add(new TaskCheck("macro-revalidation", false, "Macro procedure evidence changed before mutation; no changes were made."));
                return CloseRejectedMacroSession(session, checks, CreateMacroReceipt(operation, currentSnapshot, includeSource: false, runCompleted: false));
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
                Invoke(session.Application, "Run", QualifiedProcedureName(session.TargetWorkbook, operation));
                runCompleted = true;
                checks.Add(new TaskCheck("macro-run", true, "The requested no-argument macro procedure completed."));
            }

            observer.OnPhase("save");
            stagingPath = WorkbookRuntimeHelpers.CreateStagingPath(savedPath, plan.TaskId);
            observer.OnStagingPathCreated(stagingPath);
            Invoke(session.TargetWorkbook, "SaveAs", stagingPath, XlOpenXmlWorkbookMacroEnabled);
            checks.Add(new TaskCheck("save", true, "Excel completed the requested macro copy save."));

            observer.OnPhase("primary-cleanup");
            if (!session.Close())
            {
                session = null;
                checks.Add(new TaskCheck("owned-process-exit", false, "The owned Excel process did not exit after the macro save."));
                AddStagingCleanupCheck(stagingPath, checks);
                return new WorkbookExecutionOutcome(ExcelTaskStatus.Unknown, "Macro changes could not be verified after owned Excel cleanup.", changes, checks, CanRetry: false, RetryReason: "Inspect workbook state before retrying.", MacroProcedure: CreateMacroReceipt(operation, snapshot, includeSource: false, runCompleted));
            }
            session = null;

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
            return new WorkbookExecutionOutcome(ExcelTaskStatus.Completed, "Macro procedure changes were saved and verified after reopening.", changes, checks, MacroProcedure: CreateMacroReceipt(operation, snapshot, includeSource: false, runCompleted));
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

    private static WorkbookExecutionOutcome CloseRejectedMacroSession(ExcelSession session, List<TaskCheck> checks, MacroProcedureReceipt? receipt = null)
    {
        var cleanupVerified = session.Close();
        if (!cleanupVerified)
        {
            checks.Add(new TaskCheck("owned-process-exit", false, "The owned macro preflight Excel process did not exit."));
            return new WorkbookExecutionOutcome(ExcelTaskStatus.Unknown, "Macro preflight could not prove owned Excel cleanup.", Checks: checks, CanRetry: false, RetryReason: "Inspect the owned Excel process before retrying.", MacroProcedure: receipt);
        }

        return new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected, "Macro procedure preflight did not permit the requested operation.", Checks: checks, MacroProcedure: receipt);
    }

    private static bool TryReadMacroProcedure(object workbook, string componentName, string procedureName, out MacroProcedureSnapshot? snapshot, out TaskCheck check)
    {
        snapshot = null;
        try
        {
            using var references = new ComReferenceScope();
            var project = references.Add(Get(workbook, "VBProject"));
            if (Convert.ToInt32(Get(project, "Protection"), CultureInfo.InvariantCulture) != VbextPpNone) throw new InvalidOperationException();
            var components = references.Add(Get(project, "VBComponents"));
            var component = references.Add(Get(components, "Item", componentName));
            if (Convert.ToInt32(Get(component, "Type"), CultureInfo.InvariantCulture) != VbextCtStdModule) throw new InvalidOperationException();
            var module = references.Add(Get(component, "CodeModule"));
            var startLine = Convert.ToInt32(Invoke(module, "ProcStartLine", procedureName, VbextPkProc), CultureInfo.InvariantCulture);
            var lineCount = Convert.ToInt32(Invoke(module, "ProcCountLines", procedureName, VbextPkProc), CultureInfo.InvariantCulture);
            var source = MacroProcedureText.NormalizeLineEndings((string)Invoke(module, "Lines", startLine, lineCount)!);
            if (source.Length > MacroProcedureText.MaxSourceCharacters || CountLines(source) > MacroProcedureText.MaxLineCount) throw new InvalidOperationException();
            snapshot = new MacroProcedureSnapshot(startLine, lineCount, source, MacroProcedureText.ComputeSha256(source));
            check = new TaskCheck("macro-procedure-access", true, "The selected standard-module macro procedure is available.");
            return true;
        }
        catch (Exception exception) when (IsExpectedMacroAccessFailure(exception))
        {
            snapshot = null;
            check = new TaskCheck("macro-procedure-access", false, "The selected macro procedure cannot be safely accessed.");
            return false;
        }
    }

    private static int CountLines(string source) => source.Length == 0 ? 0 : source.Count(character => character == '\n') + 1;

    private static void ReplaceMacroProcedure(object workbook, string componentName, MacroProcedureSnapshot existing, string replacementSource)
    {
        using var references = new ComReferenceScope();
        var project = references.Add(Get(workbook, "VBProject"));
        var components = references.Add(Get(project, "VBComponents"));
        var component = references.Add(Get(components, "Item", componentName));
        var module = references.Add(Get(component, "CodeModule"));
        Invoke(module, "DeleteLines", existing.StartLine, existing.LineCount);
        Invoke(module, "InsertLines", existing.StartLine, MacroProcedureText.NormalizeLineEndings(replacementSource).Replace("\n", "\r\n", StringComparison.Ordinal));
    }

    private static string QualifiedProcedureName(object workbook, NormalizedEditMacroProcedureOperation operation) =>
        $"'{((string)Get(workbook, "Name")).Replace("'", "''", StringComparison.Ordinal)}'!{operation.ComponentName}.{operation.ProcedureName}";

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
