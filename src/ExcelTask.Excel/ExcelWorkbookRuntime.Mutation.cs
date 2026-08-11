using ExcelTask.Core;

namespace ExcelTask.Excel;

public sealed partial class ExcelWorkbookRuntime
{
    /// <summary>
    /// What an operation hands the pipeline while it runs: the open session, the accumulating
    /// receipt lists, and the two facts the pipeline must know - which phase to blame and whether
    /// a mutation has been attempted yet, which is what decides Rejected against Unknown when
    /// something fails.
    /// </summary>
    private sealed class MutationContext(
        ExcelSession session,
        IExcelWorkbookRuntimeObserver observer,
        List<TaskCheck> checks,
        List<TaskChange> changes,
        bool apply)
    {
        public ExcelSession Session { get; } = session;

        public List<TaskCheck> Checks { get; } = checks;

        public List<TaskChange> Changes { get; } = changes;

        /// <summary>False on Plan: the operation must decide without mutating.</summary>
        public bool Apply { get; } = apply;

        public bool MutationAttempted { get; private set; }

        public void OnPhase(string phase) => observer.OnPhase(phase);

        /// <summary>Called at or before the first write, never after it - it is what makes a later failure Unknown rather than Rejected.</summary>
        public void MarkMutationAttempted() => MutationAttempted = true;
    }

    /// <summary>
    /// What the operation decided: stop here with this outcome, or proceed to the save-and-verify
    /// tail. Every early exit - a failed preflight, a Plan answer, a refusal, a no-op Completed -
    /// is a <see cref="Finish"/>; the pipeline closes the session with proof either way.
    /// </summary>
    private abstract record MutationStep
    {
        /// <summary>Close and prove, then return this outcome (or the cleanup failure instead).</summary>
        public sealed record Finish(WorkbookExecutionOutcome Outcome, string What) : MutationStep;

        /// <summary>
        /// Save, prove cleanup, check the lock, verify against the reopened file, promote staging,
        /// and return Completed. <paramref name="Verify"/> runs in the separate verification Excel;
        /// <paramref name="VerifyFailedSummary"/> is the Unknown summary when it does not confirm.
        /// </summary>
        public sealed record SaveAndVerify(
            Func<ExcelSession, (bool Verified, TaskCheck Check)> Verify,
            string CompletedSummary,
            string VerifyFailedSummary,
            WorksheetRangeReceipt? Range = null) : MutationStep;
    }

    /// <summary>
    /// The one implementation of the mutate-save-verify choreography.
    ///
    /// Before this module the sequence was an interface with no implementation of its own: twelve
    /// steps in a fixed order, restated in every operation, where 91% of a method was scaffolding
    /// and the order was enforced by the author remembering it. The copies had already drifted six
    /// ways - one skipped the writability preflight, two swallowed the failure reason, the catch
    /// filters disagreed, and a promotion failure escaped through the finally that deletes the
    /// verified staging file. Every one of those is now a single decision made once.
    ///
    /// The operation supplies only what is genuinely its own: everything between the session
    /// opening and the save. Its delegate ends by returning a <see cref="MutationStep"/>, so the
    /// operation cannot run the tail out of order, skip the cleanup proof, or forget the staging
    /// checks - the tail is not its code.
    /// </summary>
    private static WorkbookExecutionOutcome ExecuteMutation(
        ExcelTaskPlan plan,
        IExcelWorkbookRuntimeObserver observer,
        string failureKind,
        string operationNoun,
        Func<MutationContext, MutationStep> operate)
    {
        var apply = plan.Request.Mode == ExcelTaskMode.Apply;
        var checks = new List<TaskCheck>();
        var changes = new List<TaskChange>();

        try
        {
            WorkbookRuntimeHelpers.EnsureReadableWorkbook(WorkbookRuntimeHelpers.NormalizePath(plan.Request.TargetWorkbookPath), "Target workbook");
            if (plan.Request.Save == SaveMode.Copy) WorkbookRuntimeHelpers.EnsureWritableCopyOutput(plan.Request.OutputWorkbookPath);
            // A read-only target found here is a clean rejection; found after the save it is
            // Unknown, the worst answer a caller can get. One operation shipped without this line.
            if (apply && plan.Request.Save == SaveMode.Same)
            {
                WorkbookRuntimeHelpers.EnsureWritableSameTarget(WorkbookRuntimeHelpers.NormalizePath(plan.Request.TargetWorkbookPath));
            }
        }
        catch (InvalidOperationException exception)
        {
            return new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected, "Workbook inputs cannot be safely executed.",
                Checks: [new TaskCheck("workbook-inputs", false, exception.Message)]);
        }

        var savedPath = WorkbookRuntimeHelpers.NormalizePath(plan.Request.Save == SaveMode.Copy
            ? plan.Request.OutputWorkbookPath ?? throw new InvalidOperationException("Copy output path is required.")
            : plan.Request.TargetWorkbookPath);

        // The engine's inspection made both of these decisions already, but in a different process
        // at an earlier moment. Re-checking here costs two filesystem reads and turns a race into a
        // clean rejection instead of a failed promotion or a clobbered open workbook.
        if (apply && plan.Request.Save == SaveMode.Copy && File.Exists(savedPath) && !plan.Request.OverwriteConfirmed)
        {
            return new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected,
                "The copy output already exists and was not authorized for overwrite.",
                Checks: [new TaskCheck("copy-output", false, "Existing output requires overwrite confirmation.")]);
        }

        if (apply && plan.Request.WorkbookBinding == WorkbookBinding.Isolated && plan.Request.Save == SaveMode.Same)
        {
            using var openTarget = RotWorkbookLocator.Find(WorkbookRuntimeHelpers.NormalizePath(plan.Request.TargetWorkbookPath));
            if (openTarget is not null)
            {
                return new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected,
                    "The target workbook is already open and cannot be safely overwritten in isolated mode.",
                    Checks: [new TaskCheck("isolated-target", false, "Use the exact open-workbook binding or save a copy.")]);
            }
        }

        ExcelSession? session = null;
        PendingVerification? pendingVerification = null;
        MutationContext? context = null;
        string? stagingPath = null;

        try
        {
            // Started before the session so its launch overlaps the operation's work; only a
            // mutating run ever verifies, so only a mutating run pays for one.
            if (apply) pendingVerification = PendingVerification.Begin(observer);

            observer.OnPhase("session-open");
            session = ExcelSession.Open(plan.Request, observer, readOnlyTarget: !apply);
            context = new MutationContext(session, observer, checks, changes, apply);

            var step = operate(context);
            if (step is MutationStep.Finish finish)
            {
                return ExcelSession.CloseAndProve(ref session, finish.What, checks, changes) ?? finish.Outcome;
            }

            var save = (MutationStep.SaveAndVerify)step;
            observer.OnPhase("save");
            if (plan.Request.Save == SaveMode.Copy)
            {
                stagingPath = WorkbookRuntimeHelpers.CreateStagingPath(savedPath, plan.TaskId);
                observer.OnStagingPathCreated(stagingPath);
                Invoke(session.TargetWorkbook, "SaveAs", stagingPath);
            }
            else
            {
                Invoke(session.TargetWorkbook, "Save");
            }

            checks.Add(new TaskCheck("save", true, "Excel completed the requested save operation."));
            observer.OnPhase("primary-cleanup");
            var cleanupFailure = ExcelSession.CloseAndProve(ref session, $"the {failureKind} save", checks, changes);
            if (cleanupFailure is not null)
            {
                AddStagingCleanupCheck(stagingPath, checks);
                return cleanupFailure with { RetryReason = "Inspect workbook state before retrying." };
            }

            var verificationPath = stagingPath ?? savedPath;
            if (plan.Request.WorkbookBinding != WorkbookBinding.UseOpen && !WorkbookRuntimeHelpers.CanOpenExclusively(verificationPath))
            {
                checks.Add(new TaskCheck("file-lock", false, "The saved workbook remained locked after owned Excel cleanup."));
                AddStagingCleanupCheck(stagingPath, checks);
                return new WorkbookExecutionOutcome(ExcelTaskStatus.Unknown,
                    "Excel saved the workbook, but the owned Excel session did not release its file lock for verification.",
                    changes, checks, CanRetry: false, RetryReason: "Inspect the Excel process and file lock before retrying.");
            }

            observer.OnPhase("reopen-verification");
            var verified = pendingVerification!.Verify(verificationPath, observer, save.Verify, out var reopenCheck);
            checks.Add(reopenCheck);
            if (!verified)
            {
                AddStagingCleanupCheck(stagingPath, checks);
                return new WorkbookExecutionOutcome(ExcelTaskStatus.Unknown, save.VerifyFailedSummary,
                    changes, checks, CanRetry: false, RetryReason: "Inspect the saved workbook before attempting another apply operation.");
            }

            if (stagingPath is not null)
            {
                observer.OnPhase("copy-promotion");
                WorkbookRuntimeHelpers.PromoteStaging(stagingPath, savedPath, plan.Request.OverwriteConfirmed);
                stagingPath = null;
                changes.Add(new TaskChange("copy-promotion", "workbook", "Promoted the verified staging workbook to the requested output path."));
            }

            return new WorkbookExecutionOutcome(ExcelTaskStatus.Completed, save.CompletedSummary, changes, checks, Range: save.Range);
        }
        // IOException and UnauthorizedAccessException are the promotion failing - the one step
        // after verification that still touches the filesystem. It used to be in no operation's
        // filter, so it escaped through the finally that deletes the verified staging file and
        // reached the caller as an Unknown with no receipt at all.
        catch (Exception exception) when (ComAccess.IsComFailure(exception) || exception is IOException or UnauthorizedAccessException)
        {
            var mutationAttempted = context?.MutationAttempted ?? false;
            checks.Add(new TaskCheck(failureKind, false, DescribeFailure(failureKind, exception)));
            var cleanupFailure = ExcelSession.CloseAndProve(ref session, $"the failed {failureKind}", checks, changes);
            if (cleanupFailure is not null) return cleanupFailure;
            return new WorkbookExecutionOutcome(
                mutationAttempted ? ExcelTaskStatus.Unknown : ExcelTaskStatus.Rejected,
                mutationAttempted
                    ? $"{operationNoun} failed after changes were attempted."
                    : $"{operationNoun} was rejected before any change was attempted.",
                changes, checks, CanRetry: !mutationAttempted,
                RetryReason: mutationAttempted ? "Inspect workbook state before retrying." : null);
        }
        finally
        {
            session?.Close();
            pendingVerification?.Dispose();
            if (stagingPath is not null) _ = WorkbookRuntimeHelpers.TryDeleteStaging(stagingPath);
        }
    }
}
