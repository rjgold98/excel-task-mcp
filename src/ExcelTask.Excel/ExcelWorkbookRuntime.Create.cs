using System.Reflection;
using System.Runtime.InteropServices;
using ExcelTask.Core;

namespace ExcelTask.Excel;

public sealed partial class ExcelWorkbookRuntime
{
    /// <summary>Excel's xlWorkbookDefault format, so a new .xlsx is saved as one.</summary>
    private const int XlWorkbookDefault = 51;

    /// <summary>
    /// Creates an empty workbook, or adds an empty worksheet to an existing one.
    ///
    /// Every other operation in this server starts from a workbook that exists, which was a hard
    /// wall: a caller with nowhere to put a result had to leave and come back. Creation is the
    /// smallest thing that removes it, and it stays small by refusing to do anything else - no
    /// starting content, no template, no position, and never an overwrite. What comes next is a
    /// write or a copy, which already exist and already verify.
    /// </summary>
    private static WorkbookExecutionOutcome ExecuteCreateCore(ExcelTaskPlan plan, IExcelWorkbookRuntimeObserver observer)
    {
        var operation = plan.Request.Operation.Create!;
        var targetPath = WorkbookRuntimeHelpers.NormalizePath(plan.Request.TargetWorkbookPath);

        try
        {
            if (operation.Kind == CreateKind.Workbook)
            {
                WorkbookRuntimeHelpers.EnsureCreatableWorkbook(targetPath);
            }
            else
            {
                WorkbookRuntimeHelpers.EnsureReadableWorkbook(targetPath, "Target workbook");
                if (plan.Request.Mode == ExcelTaskMode.Apply) WorkbookRuntimeHelpers.EnsureWritableSameTarget(targetPath);
            }
        }
        catch (InvalidOperationException exception)
        {
            return new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected, "The creation target cannot be used.",
                Checks: [new TaskCheck("workbook-inputs", false, exception.Message)]);
        }

        // Plan reaches its answer from the filesystem alone, without starting Excel. Everything it
        // could report - whether the path is free, whether the sheet name is taken - is either
        // already settled above or needs the workbook open, and for a creation the first is the part
        // a caller wants to check cheaply.
        if (plan.Request.Mode == ExcelTaskMode.Plan && operation.Kind == CreateKind.Workbook)
        {
            return new WorkbookExecutionOutcome(ExcelTaskStatus.Planned,
                $"The path is free; applying would create an empty workbook at it. Nothing was created.",
                [new TaskChange("workbook-create", Path.GetFileName(targetPath), "Planned an empty workbook.")],
                [new TaskCheck("create-path", true, "No file exists at the requested path and its directory is writable.")]);
        }

        return operation.Kind == CreateKind.Workbook
            ? CreateWorkbook(targetPath, operation.WorksheetName, observer)
            : CreateWorksheet(plan, operation, observer);
    }

    private static WorkbookExecutionOutcome CreateWorkbook(string targetPath, string? worksheetName, IExcelWorkbookRuntimeObserver observer)
    {
        var checks = new List<TaskCheck>();
        var changes = new List<TaskChange>();
        ExcelSession? session = null;
        var mutationAttempted = false;

        try
        {
            observer.OnPhase("session-open");
            session = ExcelSession.OpenNewWorkbook(observer);

            observer.OnPhase("save");
            // Named before the save, so the file has never existed under the default name. The
            // receipt reports whichever name it ended up with, because a caller who did not choose
            // one still has to write to it.
            var startingSheet = NameStartingWorksheet(session, worksheetName);

            // Set here rather than before the rename, because until SaveAs runs nothing has touched
            // the disk. Excel reserves a few worksheet names - History is the usual one - and they
            // pass the engine's name validation, so the rename above can fail with the target path
            // still free and untouched. Marking the mutation attempted first reported that as
            // Unknown and not retryable, telling the caller to go reconcile a file that was never
            // created. It is a clean rejection, and a retry with another name is safe.
            mutationAttempted = true;
            Invoke(session.TargetWorkbook, "SaveAs", targetPath,
                string.Equals(Path.GetExtension(targetPath), ".xlsm", StringComparison.OrdinalIgnoreCase)
                    ? XlOpenXmlWorkbookMacroEnabled
                    : XlWorkbookDefault);
            checks.Add(new TaskCheck("save", true, "Excel saved the new workbook to the requested path."));
            checks.Add(new TaskCheck("starting-worksheet", true, $"The new workbook's sheet is named {startingSheet}."));
            changes.Add(new TaskChange("workbook-create", Path.GetFileName(targetPath), $"Created an empty workbook whose sheet is named {startingSheet}."));

            observer.OnPhase("primary-cleanup");
            var cleanupFailure = ExcelSession.CloseAndProve(ref session, "the workbook creation", checks, changes);
            if (cleanupFailure is not null) return cleanupFailure with { RetryReason = "Inspect the created workbook before retrying." };

            // Proof from the filesystem rather than from Excel's word for it. A creation is the one
            // operation where "it exists now" is the entire result, so reopening it to read a cell
            // would be verifying something nobody asked to be there.
            observer.OnPhase("create-verification");
            if (!File.Exists(targetPath))
            {
                checks.Add(new TaskCheck("create-verification", false, "Excel reported a successful save but no file exists at the path."));
                return new WorkbookExecutionOutcome(ExcelTaskStatus.Unknown, "The new workbook could not be found after saving.",
                    changes, checks, CanRetry: false, RetryReason: "Inspect the requested path before retrying.");
            }

            checks.Add(new TaskCheck("create-verification", true, "A workbook file exists at the requested path after owned Excel exited."));
            return new WorkbookExecutionOutcome(ExcelTaskStatus.Completed,
                $"Created an empty workbook at the requested path and confirmed it exists.", changes, checks);
        }
        catch (Exception exception) when (ComAccess.IsComFailure(exception))
        {
            checks.Add(new TaskCheck("workbook-create", false, DescribeFailure("workbook-create", exception)));
            var cleanupFailure = ExcelSession.CloseAndProve(ref session, "the failed workbook creation", checks, changes);
            if (cleanupFailure is not null) return cleanupFailure;
            return new WorkbookExecutionOutcome(
                mutationAttempted ? ExcelTaskStatus.Unknown : ExcelTaskStatus.Rejected,
                mutationAttempted
                    ? "The workbook creation failed after the save was attempted."
                    : "The workbook creation was rejected before anything was written.",
                changes, checks, CanRetry: !mutationAttempted,
                RetryReason: mutationAttempted ? "Inspect the requested path before retrying." : null);
        }
        finally
        {
            session?.Close();
        }
    }

    private static WorkbookExecutionOutcome CreateWorksheet(
        ExcelTaskPlan plan,
        NormalizedCreateOperation operation,
        IExcelWorkbookRuntimeObserver observer) =>
        ExecuteMutation(plan, observer, "worksheet-create", "The worksheet creation", context =>
        {
            context.OnPhase("worksheet-preflight");
            if (WorksheetExists(context.Session, operation.WorksheetName!))
            {
                context.Checks.Add(new TaskCheck("worksheet-name", false, "A worksheet with that name already exists; creating one never replaces it."));
                return new MutationStep.Finish(
                    new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected, "That worksheet name is already in use.", Checks: context.Checks),
                    "worksheet preflight");
            }

            context.Checks.Add(new TaskCheck("worksheet-name", true, "No worksheet in the target workbook carries that name."));
            if (!context.Apply)
            {
                context.Changes.Add(new TaskChange("worksheet-create", operation.WorksheetName!, "Planned an empty worksheet."));
                return new MutationStep.Finish(
                    new WorkbookExecutionOutcome(ExcelTaskStatus.Planned,
                        $"The name is free; applying would add an empty worksheet. Nothing was changed.", context.Changes, context.Checks),
                    "worksheet planning");
            }

            context.OnPhase("worksheet-create");
            context.MarkMutationAttempted();
            AddWorksheet(context.Session, operation.WorksheetName!);
            context.Changes.Add(new TaskChange("worksheet-create", operation.WorksheetName!, "Added an empty worksheet."));

            return new MutationStep.SaveAndVerify(
                verification => WorksheetExists(verification, operation.WorksheetName!)
                    ? (true, new TaskCheck("reopen-verification", true, "The saved workbook reopened with the new worksheet present."))
                    : (false, new TaskCheck("reopen-verification", false, "The new worksheet was not present after reopening the saved workbook.")),
                $"Added an empty worksheet, saved, and confirmed it after reopening.",
                "Excel saved the workbook, but reopen verification did not find the new worksheet.");
        });

    /// <summary>
    /// Renames the new workbook's single sheet, or reports the default name Excel chose. Excel's
    /// default depends on the user's locale and their "sheets in new workbook" setting, so the name
    /// is read back rather than assumed - a caller writing to "Sheet1" on a machine that calls it
    /// "Hoja1" would get a worksheet-not-found rejection it could not have predicted.
    /// </summary>
    private static string NameStartingWorksheet(ExcelSession session, string? worksheetName)
    {
        using var references = new ComReferenceScope();
        var sheets = references.Add(Get(session.TargetWorkbook, "Worksheets"));
        var first = references.Add(Item(sheets, 1));
        if (worksheetName is not null) Set(first, "Name", worksheetName);
        return (string)Get(first, "Name");
    }

    private static bool WorksheetExists(ExcelSession session, string worksheetName)
    {
        using var references = new ComReferenceScope();
        var sheets = references.Add(Get(session.TargetWorkbook, "Worksheets"));
        try
        {
            references.Add(Item(sheets, worksheetName));
            return true;
        }
        catch (Exception exception) when (ComAccess.IsComFailure(exception))
        {
            return false;
        }
    }

    /// <summary>
    /// Adds one sheet after the last, so an added sheet never displaces the one the workbook opens
    /// on. Excel places a new sheet before the active one by default, which would quietly change
    /// what a person sees when they open the file.
    /// </summary>
    private static void AddWorksheet(ExcelSession session, string worksheetName)
    {
        using var references = new ComReferenceScope();
        var sheets = references.Add(Get(session.TargetWorkbook, "Worksheets"));
        var count = Convert.ToInt32(Get(sheets, "Count"), System.Globalization.CultureInfo.InvariantCulture);
        var last = references.Add(Item(sheets, count));
        var added = references.Add(ComAccess.Invoke(sheets, "Add", Type.Missing, last)
            ?? throw new InvalidOperationException("Excel did not return the new worksheet."));
        Set(added, "Name", worksheetName);
    }
}
