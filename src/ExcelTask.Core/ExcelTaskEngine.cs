using System.Diagnostics;
using System.Text.RegularExpressions;

namespace ExcelTask.Core;

public sealed partial class ExcelTaskEngine(IWorkbookRuntime runtime) : IExcelTaskEngine
{
    /// <summary>Maximum aggregate number of cells requested for formula repair in the MVP.</summary>
    public const int MaxFormulaRepairCells = 10_000;
    public const int MaxFormulaRepairRanges = 16;

    private const int MaxReceiptItems = 20;
    private const int MaxReceiptStringLength = 256;
    private static readonly HashSet<string> AutoEntryProcedureNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "Auto_Activate", "Auto_Close", "Auto_Deactivate", "Auto_Exec", "Auto_Exit", "Auto_New", "Auto_Open"
    };

    private readonly IWorkbookRuntime _runtime = runtime ?? throw new ArgumentNullException(nameof(runtime));

    public async Task<ExcelTaskReceipt> RunAsync(ExcelTaskRequest request, CancellationToken cancellationToken)
    {
        var total = Stopwatch.StartNew();
        var taskId = Guid.NewGuid().ToString("N");
        var validation = Stopwatch.StartNew();

        if (!TryNormalize(request, out var normalized, out var validationError))
        {
            validation.Stop();
            return CreateReceipt(
                taskId,
                ExcelTaskStatus.Rejected,
                validationError!,
                [],
                [new TaskCheck("request", false, validationError!)],
                request?.Save ?? SaveMode.Same,
                request?.OutputWorkbookPath,
                request?.OverwriteConfirmed ?? false,
                false,
                [],
                false,
                null,
                validation.Elapsed,
                TimeSpan.Zero,
                TimeSpan.Zero,
                total.Elapsed);
        }

        validation.Stop();
        var normalizedRequest = normalized!;
        var inspectionTimer = Stopwatch.StartNew();
        WorkbookInspection inspection;
        try
        {
            inspection = await _runtime.InspectAsync(
                new WorkbookInspectionRequest(
                    normalizedRequest.TargetWorkbookPath,
                    normalizedRequest.Operation.CopyExhibit?.ReferenceWorkbookPath,
                    normalizedRequest.WorkbookBinding,
                    normalizedRequest.Save,
                    normalizedRequest.OutputWorkbookPath),
                cancellationToken) ?? throw new InvalidOperationException("Workbook runtime returned no inspection result.");
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            throw;
        }
        catch (Exception)
        {
            inspectionTimer.Stop();
            return CreateReceipt(
                taskId,
                ExcelTaskStatus.Rejected,
                "Workbook inspection could not be completed before execution.",
                [],
                [new TaskCheck("runtime-inspection", false, "Inspection is unavailable before execution.")],
                normalizedRequest.Save,
                normalizedRequest.OutputWorkbookPath,
                normalizedRequest.OverwriteConfirmed,
                false,
                [],
                true,
                "Retry after the workbook runtime is available; no execution was dispatched.",
                validation.Elapsed,
                inspectionTimer.Elapsed,
                TimeSpan.Zero,
                total.Elapsed);
        }

        inspectionTimer.Stop();
        if (TryGetInspectionRejection(normalizedRequest, inspection, out var inspectionRejection))
        {
            return CreateReceipt(
                taskId,
                ExcelTaskStatus.Rejected,
                inspectionRejection!,
                [],
                inspection.Checks ?? [],
                normalizedRequest.Save,
                normalizedRequest.OutputWorkbookPath,
                normalizedRequest.OverwriteConfirmed,
                false,
                [],
                true,
                "Adjust the workbook binding or save policy, then submit a new task.",
                validation.Elapsed,
                inspectionTimer.Elapsed,
                TimeSpan.Zero,
                total.Elapsed);
        }

        var requirements = GetConfirmationRequirements(normalizedRequest, inspection);
        if (requirements.Count > 0)
        {
            var description = inspection.TargetIsOpen
                ? "The target workbook is already open."
                : "The requested copy output already exists.";
            return CreateReceipt(
                taskId,
                ExcelTaskStatus.NeedsConfirmation,
                description,
                [],
                inspection.Checks ?? [],
                normalizedRequest.Save,
                normalizedRequest.OutputWorkbookPath,
                normalizedRequest.OverwriteConfirmed,
                true,
                requirements,
                false,
                null,
                validation.Elapsed,
                inspectionTimer.Elapsed,
                TimeSpan.Zero,
                total.Elapsed);
        }

        var plan = new ExcelTaskPlan(taskId, normalizedRequest);
        var executionTimer = Stopwatch.StartNew();
        try
        {
            var outcome = await _runtime.ExecuteAsync(plan, cancellationToken);
            executionTimer.Stop();
            return CreateReceipt(
                taskId,
                outcome.Status,
                outcome.Summary,
                outcome.Changes ?? [],
                CombineChecks(inspection.Checks, outcome.Checks),
                normalizedRequest.Save,
                normalizedRequest.OutputWorkbookPath,
                normalizedRequest.OverwriteConfirmed,
                false,
                [],
                outcome.CanRetry,
                outcome.RetryReason,
                validation.Elapsed,
                inspectionTimer.Elapsed,
                executionTimer.Elapsed,
                total.Elapsed,
                outcome.MacroProcedure,
                normalizedRequest.Mode == ExcelTaskMode.Plan &&
                normalizedRequest.Operation.EditMacroProcedure is not null &&
                outcome.Status == ExcelTaskStatus.Planned,
                outcome.Audit);
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            throw;
        }
        catch (Exception)
        {
            executionTimer.Stop();
            return CreateReceipt(
                taskId,
                ExcelTaskStatus.Unknown,
                "Workbook execution did not complete.",
                [],
                CombineChecks(inspection.Checks, [new TaskCheck("runtime-execution", false, "Execution failed after dispatch.")]),
                normalizedRequest.Save,
                normalizedRequest.OutputWorkbookPath,
                normalizedRequest.OverwriteConfirmed,
                false,
                [],
                false,
                "Reconcile the target workbook before attempting any new task; execution may have applied changes.",
                validation.Elapsed,
                inspectionTimer.Elapsed,
                executionTimer.Elapsed,
                total.Elapsed);
        }
    }

    private static List<ConfirmationRequirement> GetConfirmationRequirements(NormalizedExcelTaskRequest request, WorkbookInspection inspection)
    {
        var requirements = new List<ConfirmationRequirement>();
        if (inspection.TargetIsOpen && request.WorkbookBinding == WorkbookBinding.AskIfOpen)
        {
            requirements.Add(new ConfirmationRequirement(
                "target-open",
                "Target is open. Resubmit with workbook binding UseOpen or Isolated after choosing how Excel should bind it."));
        }

        // A read-only operation has no save to authorize. Asking to confirm an overwrite that
        // cannot happen would teach the caller that the confirmation means nothing.
        if (request.Operation.AuditWorkbookFlows is not null) return requirements;

        if (request.Mode == ExcelTaskMode.Apply && request.Save == SaveMode.Same && !request.OverwriteConfirmed)
        {
            requirements.Add(new ConfirmationRequirement(
                "overwrite-same",
                "Applying with save mode Same overwrites the target workbook. Set overwriteConfirmed to true to proceed."));
        }

        if (request.Mode == ExcelTaskMode.Apply && request.Save == SaveMode.Copy && inspection.CopyOutputExists && !request.OverwriteConfirmed)
        {
            requirements.Add(new ConfirmationRequirement(
                "overwrite-copy",
                "The requested copy output already exists. Set overwriteConfirmed to true to proceed."));
        }

        return requirements;
    }

    private static bool TryGetInspectionRejection(NormalizedExcelTaskRequest request, WorkbookInspection inspection, out string? rejection)
    {
        if (request.WorkbookBinding == WorkbookBinding.UseOpen && !inspection.TargetIsOpen)
        {
            rejection = "Workbook binding UseOpen requires the target workbook to already be open.";
            return true;
        }

        if (request.Mode == ExcelTaskMode.Apply &&
            request.WorkbookBinding == WorkbookBinding.Isolated &&
            request.Save == SaveMode.Same &&
            inspection.TargetIsOpen)
        {
            rejection = "An isolated task cannot safely overwrite an open target workbook. Choose UseOpen or save a Copy.";
            return true;
        }

        rejection = null;
        return false;
    }

    private static ExcelTaskReceipt CreateReceipt(
        string taskId,
        ExcelTaskStatus status,
        string summary,
        IReadOnlyList<TaskChange> changes,
        IReadOnlyList<TaskCheck> checks,
        SaveMode save,
        string? outputPath,
        bool overwriteConfirmed,
        bool confirmationRequired,
        IReadOnlyList<ConfirmationRequirement> requirements,
        bool canRetry,
        string? retryReason,
        TimeSpan validation,
        TimeSpan inspection,
        TimeSpan execution,
        TimeSpan total,
        MacroProcedureReceipt? macroProcedure = null,
        bool includeMacroSource = false,
        WorkbookAuditReceipt? audit = null)
    {
        if (status == ExcelTaskStatus.Rejected)
        {
            canRetry = true;
            retryReason ??= "Correct the rejected request or pre-dispatch condition, then submit a new task.";
        }
        else if (status == ExcelTaskStatus.Unknown)
        {
            canRetry = false;
            retryReason = "Reconcile the target workbook before attempting any new task; execution may have applied changes.";
        }
        else if (status == ExcelTaskStatus.Partial)
        {
            canRetry = false;
            retryReason = "Review and reconcile the reported partial changes before attempting any new task.";
        }

        return new ExcelTaskReceipt(
            taskId,
            status,
            Bound(summary) ?? string.Empty,
            BoundChanges(changes),
            BoundChecks(checks),
            new SaveReceipt(save, Bound(DisplayOutputPath(outputPath)), overwriteConfirmed),
            new RetryReceipt(canRetry, Bound(retryReason)),
            new ConfirmationReceipt(confirmationRequired, BoundRequirements(requirements)),
            new PhaseTimings(validation, inspection, execution, total),
            BoundMacroProcedure(macroProcedure, includeMacroSource),
            BoundAudit(audit));
    }

    /// <summary>
    /// Caps the audit at the same receipt limits as everything else. The reported total counts what
    /// the workbook actually contains, so a bounded report still says how much it is not showing.
    /// </summary>
    private static WorkbookAuditReceipt? BoundAudit(WorkbookAuditReceipt? audit)
    {
        if (audit is null) return null;

        var items = audit.Items
            .Take(MaxReceiptItems)
            .Select(item => new WorkbookFlowItem(
                Bound(item.Kind) ?? string.Empty,
                Bound(item.Name) ?? string.Empty,
                Bound(item.Detail) ?? string.Empty,
                Bound(item.DependsOn)))
            .ToArray();

        return new WorkbookAuditReceipt(items, audit.TotalFound, audit.Truncated || audit.Items.Count > items.Length, audit.WorkbookUnchanged);
    }

    private static MacroProcedureReceipt? BoundMacroProcedure(MacroProcedureReceipt? macroProcedure, bool includeSource)
    {
        if (macroProcedure is null) return null;

        // Plan source is evidence read out of the workbook and already bounded by the runtime, so it
        // is only length-capped here. Re-checking it against the replacement grammar would silently
        // return a successful plan with no source for procedures that are valid VBA but outside
        // that deliberately narrow grammar.
        string? source = null;
        if (includeSource && macroProcedure.Source is not null)
        {
            // Oversized evidence is omitted rather than truncated: a partial procedure could lead a
            // caller to write a replacement from incomplete source and destroy the remainder. The
            // runtime already rejects oversized procedures at read time, so this is a boundary guard.
            var normalized = MacroProcedureText.NormalizeLineEndings(macroProcedure.Source);
            if (normalized.Length <= MacroProcedureText.MaxSourceCharacters) source = normalized;
        }

        return new MacroProcedureReceipt(
            Bound(macroProcedure.ComponentName) ?? string.Empty,
            Bound(macroProcedure.ProcedureName) ?? string.Empty,
            Bound(macroProcedure.Sha256) ?? string.Empty,
            source,
            macroProcedure.RunRequested,
            macroProcedure.RunCompleted);
    }

    private static IReadOnlyList<TaskCheck> CombineChecks(IReadOnlyList<TaskCheck>? first, IReadOnlyList<TaskCheck>? second)
    {
        if (first is null or { Count: 0 }) return second ?? [];
        if (second is null or { Count: 0 }) return first;
        return [.. first, .. second];
    }

    private static bool TryNormalize(ExcelTaskRequest? request, out NormalizedExcelTaskRequest? normalized, out string? error)
    {
        normalized = null;
        if (request is null)
        {
            error = "A task request is required.";
            return false;
        }

        if (!Enum.IsDefined(request.Mode) || !Enum.IsDefined(request.WorkbookBinding) || !Enum.IsDefined(request.Save))
        {
            error = "Mode, workbook binding, and save mode must be defined values.";
            return false;
        }

        if (!TryNormalizeWorkbookPath(request.TargetWorkbookPath, "Target workbook path", out var target, out error))
        {
            return false;
        }
        if (!IsSupportedWorkbookPath(target!))
        {
            error = "MVP workbook paths must use a .xlsx or .xlsm extension.";
            return false;
        }
        if (!TryNormalizeOperation(request.Operation, request.Mode, out var operation, out error)) return false;

        if (request.WorkbookBinding == WorkbookBinding.UseOpen && request.Save == SaveMode.Copy)
        {
            error = "Workbook binding UseOpen supports save mode Same only; use Isolated for a copy output.";
            return false;
        }

        string? output = null;
        if (request.Save == SaveMode.Copy)
        {
            if (!TryNormalizeWorkbookPath(request.OutputWorkbookPath, "Output workbook path", out output, out error)) return false;
            if (!IsSupportedWorkbookPath(output!))
            {
                error = "MVP workbook paths must use a .xlsx or .xlsm extension.";
                return false;
            }
            if (string.Equals(target, output, StringComparison.OrdinalIgnoreCase))
            {
                error = "Copy save mode requires an output workbook path different from the target workbook path.";
                return false;
            }
            if (!string.Equals(Path.GetExtension(target), Path.GetExtension(output), StringComparison.OrdinalIgnoreCase))
            {
                error = "Copy output extension must match the target workbook extension.";
                return false;
            }
        }
        else if (!string.IsNullOrWhiteSpace(request.OutputWorkbookPath))
        {
            error = "Output workbook path is only valid with save mode Copy.";
            return false;
        }

        if (operation!.EditMacroProcedure is not null &&
            !TryValidateMacroRequestPolicy(target!, request.WorkbookBinding, request.Save, output, out error))
        {
            return false;
        }

        // An audit reads and reports; it has nothing to write. Refusing a save destination outright
        // means the read-only promise is enforced by the request shape rather than by the runtime
        // remembering not to save.
        if (operation.AuditWorkbookFlows is not null && (request.Save == SaveMode.Copy || output is not null))
        {
            error = "Auditing workbook flows never writes, so it must not be given a save destination.";
            return false;
        }

        normalized = new NormalizedExcelTaskRequest(
            target!, request.Mode, request.WorkbookBinding, request.Save, output, request.OverwriteConfirmed, operation!);
        error = null;
        return true;
    }

    private static bool TryNormalizeWorkbookPath(string? value, string name, out string? normalized, out string? error)
    {
        normalized = null;
        error = null;
        if (string.IsNullOrWhiteSpace(value))
        {
            error = $"{name} is required.";
            return false;
        }

        var trimmed = value.Trim();
        if (trimmed.IndexOfAny(Path.GetInvalidPathChars()) >= 0)
        {
            error = $"{name} contains invalid path characters.";
            return false;
        }

        try
        {
            normalized = Path.GetFullPath(trimmed);
            return true;
        }
        catch (Exception exception) when (exception is ArgumentException or NotSupportedException or PathTooLongException)
        {
            error = $"{name} is invalid.";
            return false;
        }
    }

    private static bool TryNormalizeWorksheetName(string? value, string name, out string? normalized, out string? error)
    {
        normalized = null;
        error = null;
        if (string.IsNullOrWhiteSpace(value))
        {
            error = $"{name} is required.";
            return false;
        }

        var trimmed = value.Trim();
        if (trimmed.Length > 31 || trimmed.IndexOfAny([':', '\\', '/', '?', '*', '[', ']']) >= 0)
        {
            error = $"{name} is not a valid Excel worksheet name.";
            return false;
        }

        normalized = trimmed;
        return true;
    }

    private static bool TryNormalizeRanges(IReadOnlyList<string>? values, string name, out IReadOnlyList<FormulaRepairRange>? ranges, out string? error)
    {
        ranges = null;
        error = null;
        if (values is null)
        {
            error = $"{name} are required; supply an empty list when none are needed.";
            return false;
        }
        if (values.Count > MaxFormulaRepairRanges)
        {
            error = $"{name} exceed the MVP limit of {MaxFormulaRepairRanges} ranges.";
            return false;
        }

        var parsed = new List<ParsedA1Range>(values.Count);
        var normalized = new List<FormulaRepairRange>(values.Count);
        long aggregateCellCount = 0;
        foreach (var value in values)
        {
            if (!TryNormalizeA1Range(value, out var range))
            {
                error = $"{name} contains invalid range '{value}'. Use a rectangular A1 range such as A1:C10.";
                return false;
            }
            if (parsed.Any(existing => Overlaps(existing, range)))
            {
                error = $"{name} must not overlap.";
                return false;
            }

            aggregateCellCount += CellCount(range);
            if (aggregateCellCount > MaxFormulaRepairCells)
            {
                error = $"{name} exceed the MVP limit of {MaxFormulaRepairCells:N0} aggregate cells.";
                return false;
            }
            parsed.Add(range);
            normalized.Add(ToFormulaRepairRange(range));
        }

        ranges = normalized.AsReadOnly();
        return true;
    }

    private static bool TryNormalizeOperation(ExcelOperation? operation, ExcelTaskMode mode, out NormalizedExcelOperation? normalized, out string? error)
    {
        normalized = null;
        error = null;
        if (operation is null)
        {
            error = "Operation is required.";
            return false;
        }

        var payloadCount = (operation.CopyExhibit is null ? 0 : 1) +
                           (operation.RepairExistingWorksheet is null ? 0 : 1) +
                           (operation.ExtendFormulaSeries is null ? 0 : 1) +
                           (operation.EditMacroProcedure is null ? 0 : 1) +
                           (operation.AuditWorkbookFlows is null ? 0 : 1);
        if (!Enum.IsDefined(operation.Kind))
        {
            error = "Operation kind must be a defined value.";
            return false;
        }
        if (payloadCount != 1)
        {
            error = "Operation must specify exactly one payload.";
            return false;
        }

        switch (operation.Kind)
        {
            case ExcelOperationKind.CopyExhibit when operation.CopyExhibit is not null:
                var copy = operation.CopyExhibit;
                if (!TryNormalizeWorkbookPath(copy.ReferenceWorkbookPath, "Reference workbook path", out var referencePath, out error) ||
                    !IsSupportedWorkbookPath(referencePath!))
                {
                    error ??= "MVP workbook paths must use a .xlsx or .xlsm extension.";
                    return false;
                }
                if (!TryNormalizeWorksheetName(copy.ReferenceWorksheet, "Reference worksheet", out var referenceSheet, out error) ||
                    !TryNormalizeWorksheetName(copy.NewWorksheetName, "New worksheet name", out var newSheet, out error) ||
                    !TryNormalizeRanges(copy.RepairRanges, "Copy exhibit repair ranges", out var repairRanges, out error))
                {
                    return false;
                }
                normalized = new NormalizedExcelOperation(operation.Kind,
                    new NormalizedCopyExhibitOperation(referencePath!, referenceSheet!, newSheet!, repairRanges!));
                return true;

            case ExcelOperationKind.RepairExistingWorksheet when operation.RepairExistingWorksheet is not null:
                var repair = operation.RepairExistingWorksheet;
                if (!TryNormalizeWorksheetName(repair.WorksheetName, "Worksheet name", out var worksheetName, out error) ||
                    !TryNormalizeRanges(repair.Ranges, "Repair ranges", out var ranges, out error))
                {
                    return false;
                }
                if (ranges!.Count == 0)
                {
                    error = "RepairExistingWorksheet requires one or more repair ranges.";
                    return false;
                }
                normalized = new NormalizedExcelOperation(operation.Kind,
                    RepairExistingWorksheet: new NormalizedRepairExistingWorksheetOperation(worksheetName!, ranges!));
                return true;

            case ExcelOperationKind.ExtendFormulaSeries when operation.ExtendFormulaSeries is not null:
                return TryNormalizeExtension(operation, out normalized, out error);

            case ExcelOperationKind.EditMacroProcedure when operation.EditMacroProcedure is not null:
                return TryNormalizeMacroOperation(operation.EditMacroProcedure, mode, operation.Kind, out normalized, out error);

            case ExcelOperationKind.AuditWorkbookFlows when operation.AuditWorkbookFlows is not null:
                normalized = new NormalizedExcelOperation(operation.Kind, AuditWorkbookFlows: new NormalizedAuditWorkbookFlowsOperation());
                return true;

            default:
                error = "Operation payload does not match its kind.";
                return false;
        }
    }

    private static bool TryNormalizeMacroOperation(
        EditMacroProcedureOperation macro,
        ExcelTaskMode mode,
        ExcelOperationKind kind,
        out NormalizedExcelOperation? normalized,
        out string? error)
    {
        normalized = null;
        if (!TryNormalizeVbaIdentifier(macro.ComponentName, "Component name", out var componentName, out error) ||
            !TryNormalizeVbaIdentifier(macro.ProcedureName, "Procedure name", out var procedureName, out error))
        {
            return false;
        }
        if (AutoEntryProcedureNames.Contains(procedureName!))
        {
            error = "Automatic-entry VBA procedures cannot be edited.";
            return false;
        }

        string? expectedHash = null;
        string? replacementSource = null;
        if (mode == ExcelTaskMode.Plan)
        {
            if (macro.ExpectedProcedureSha256 is not null || macro.ReplacementSource is not null || macro.RunAfterEdit)
            {
                error = "Macro Plan is inspect-only and must omit the expected hash, replacement source, and run request.";
                return false;
            }
        }
        else
        {
            var suppliedHash = macro.ExpectedProcedureSha256?.Trim();
            if (suppliedHash is null || !Sha256Regex().IsMatch(suppliedHash))
            {
                error = "Macro Apply requires a 64-character hexadecimal expected procedure SHA-256.";
                return false;
            }
            if (!MacroProcedureText.TryNormalizeProcedureSource(
                    macro.ReplacementSource,
                    procedureName!,
                    macro.RunAfterEdit,
                    out replacementSource,
                    out error))
            {
                return false;
            }
            // Refused before Excel is opened: running this would stop on a modal dialog that only a
            // person can clear. Editing such a procedure without running it stays permitted.
            if (macro.RunAfterEdit && MacroProcedureText.TryFindBlockingConstruct(replacementSource!, out var blocking))
            {
                error = $"RunAfterEdit rejects a replacement containing {blocking}, which waits for a person.";
                return false;
            }

            expectedHash = suppliedHash.ToLowerInvariant();
        }

        normalized = new NormalizedExcelOperation(
            kind,
            EditMacroProcedure: new NormalizedEditMacroProcedureOperation(
                componentName!, procedureName!, expectedHash, replacementSource, macro.RunAfterEdit));
        error = null;
        return true;
    }

    private static bool TryValidateMacroRequestPolicy(
        string targetWorkbookPath,
        WorkbookBinding workbookBinding,
        SaveMode save,
        string? outputWorkbookPath,
        out string? error)
    {
        error = null;
        // All violations are reported in one message. Field use showed a caller being rejected
        // twice, learning one rule per round trip; every unmet requirement in a single rejection
        // makes the second attempt the correct one.
        var violations = new List<string>();
        if (!string.Equals(Path.GetExtension(targetWorkbookPath), ".xlsm", StringComparison.OrdinalIgnoreCase))
        {
            violations.Add("an .xlsm target workbook path");
        }
        if (workbookBinding != WorkbookBinding.Isolated)
        {
            violations.Add("workbook binding Isolated");
        }
        if (save != SaveMode.Copy)
        {
            violations.Add("save mode Copy");
        }
        if (!string.Equals(Path.GetExtension(outputWorkbookPath), ".xlsm", StringComparison.OrdinalIgnoreCase))
        {
            violations.Add("an .xlsm outputWorkbookPath");
        }

        if (violations.Count == 0) return true;

        error = $"Macro editing requires {string.Join(", ", violations)}. Correct all of these in one resubmission.";
        return false;
    }

    private static bool TryNormalizeVbaIdentifier(string? value, string name, out string? normalized, out string? error)
    {
        normalized = null;
        error = null;
        if (string.IsNullOrWhiteSpace(value))
        {
            error = $"{name} is required.";
            return false;
        }

        var trimmed = value.Trim();
        if (trimmed.Length > 31 || !VbaIdentifierRegex().IsMatch(trimmed))
        {
            error = $"{name} must be a VBA identifier of at most 31 characters.";
            return false;
        }

        normalized = trimmed;
        return true;
    }

    private static bool TryNormalizeExtension(ExcelOperation operation, out NormalizedExcelOperation? normalized, out string? error)
    {
        normalized = null;
        error = null;
        var extension = operation.ExtendFormulaSeries!;
        if (!Enum.IsDefined(extension.Direction))
        {
            error = "Formula extension direction must be a defined value.";
            return false;
        }
        if (!TryNormalizeWorksheetName(extension.WorksheetName, "Worksheet name", out var worksheetName, out error) ||
            !TryNormalizeA1Range(extension.EvidenceRange, out var evidence) ||
            !TryNormalizeA1Range(extension.DestinationRange, out var destination))
        {
            error ??= "ExtendFormulaSeries ranges must be valid rectangular A1 ranges.";
            return false;
        }

        var evidencePeriods = extension.Direction == FormulaExtensionDirection.Right ? evidence.Width : evidence.Height;
        var destinationPeriods = extension.Direction == FormulaExtensionDirection.Right ? destination.Width : destination.Height;
        var samePerpendicular = extension.Direction == FormulaExtensionDirection.Right
            ? evidence.StartRow == destination.StartRow && evidence.EndRow == destination.EndRow
            : evidence.StartColumn == destination.StartColumn && evidence.EndColumn == destination.EndColumn;
        var adjacent = extension.Direction == FormulaExtensionDirection.Right
            ? destination.StartColumn == evidence.EndColumn + 1
            : destination.StartRow == evidence.EndRow + 1;
        if (evidencePeriods != 2 || destinationPeriods is < 1 or > FormulaMutationPlanner.MaxPeriods || !samePerpendicular || !adjacent)
        {
            error = "ExtendFormulaSeries requires exactly 2 evidence columns/rows and 1-24 immediately adjacent destination periods with matching perpendicular geometry.";
            return false;
        }
        if (CellCount(destination) > FormulaMutationPlanner.MaxMutations)
        {
            error = $"ExtendFormulaSeries destination exceeds the MVP limit of {FormulaMutationPlanner.MaxMutations:N0} cells.";
            return false;
        }
        if (CellCount(evidence) + CellCount(destination) > MaxFormulaRepairCells)
        {
            error = $"ExtendFormulaSeries ranges exceed the MVP limit of {MaxFormulaRepairCells:N0} aggregate cells.";
            return false;
        }

        normalized = new NormalizedExcelOperation(operation.Kind,
            ExtendFormulaSeries: new NormalizedExtendFormulaSeriesOperation(worksheetName!, extension.Direction, ToFormulaRepairRange(evidence), ToFormulaRepairRange(destination)));
        return true;
    }

    private readonly record struct ParsedA1Range(int StartColumn, int StartRow, int EndColumn, int EndRow)
    {
        public int Width => EndColumn - StartColumn + 1;
        public int Height => EndRow - StartRow + 1;
    }

    private static bool TryNormalizeA1Range(string? value, out ParsedA1Range range)
    {
        range = default;
        var match = A1RangeRegex().Match(value?.Trim() ?? string.Empty);
        if (!match.Success || !TryParseCell(match.Groups[1].Value, match.Groups[2].Value, out var sc, out var sr)) return false;
        var ecText = match.Groups[3].Success ? match.Groups[3].Value : match.Groups[1].Value;
        var erText = match.Groups[4].Success ? match.Groups[4].Value : match.Groups[2].Value;
        if (!TryParseCell(ecText, erText, out var ec, out var er) || sc > ec || sr > er) return false;
        range = new ParsedA1Range(sc, sr, ec, er);
        return true;
    }

    private static FormulaRepairRange ToFormulaRepairRange(ParsedA1Range range) =>
        new($"{ColumnName(range.StartColumn)}{range.StartRow}", $"{ColumnName(range.EndColumn)}{range.EndRow}");

    private static long CellCount(ParsedA1Range range) => (long)range.Width * range.Height;

    private static bool Overlaps(ParsedA1Range left, ParsedA1Range right) =>
        left.StartColumn <= right.EndColumn && right.StartColumn <= left.EndColumn &&
        left.StartRow <= right.EndRow && right.StartRow <= left.EndRow;

    private static string ColumnName(int column)
    {
        Span<char> buffer = stackalloc char[3];
        var position = buffer.Length;
        while (column > 0)
        {
            column--;
            buffer[--position] = (char)('A' + (column % 26));
            column /= 26;
        }
        return new string(buffer[position..]);
    }

    private static bool TryParseCell(string columnText, string rowText, out int column, out int row)
    {
        column = 0;
        row = 0;
        foreach (var character in columnText)
        {
            column = (column * 26) + (char.ToUpperInvariant(character) - 'A' + 1);
        }

        return column is >= 1 and <= 16384 && int.TryParse(rowText, out row) && row is >= 1 and <= 1_048_576;
    }

    private static bool IsSupportedWorkbookPath(string path) =>
        string.Equals(Path.GetExtension(path), ".xlsx", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(Path.GetExtension(path), ".xlsm", StringComparison.OrdinalIgnoreCase);

    private static string? DisplayOutputPath(string? outputPath) => string.IsNullOrWhiteSpace(outputPath) ? null : Path.GetFileName(outputPath);

    private static string? Bound(string? value) => value is null
        ? null
        : value.Length <= MaxReceiptStringLength ? value : value[..MaxReceiptStringLength];

    private static TaskChange[] BoundChanges(IReadOnlyList<TaskChange> changes) => changes
        .Take(MaxReceiptItems)
        .Select(change => new TaskChange(Bound(change.Kind) ?? string.Empty, Bound(change.Target) ?? string.Empty, Bound(change.Summary) ?? string.Empty))
        .ToArray();

    private static TaskCheck[] BoundChecks(IReadOnlyList<TaskCheck> checks) => checks
        .Take(MaxReceiptItems)
        .Select(check => new TaskCheck(Bound(check.Name) ?? string.Empty, check.Passed, Bound(check.Detail) ?? string.Empty))
        .ToArray();

    private static ConfirmationRequirement[] BoundRequirements(IReadOnlyList<ConfirmationRequirement> requirements) => requirements
        .Take(MaxReceiptItems)
        .Select(requirement => new ConfirmationRequirement(Bound(requirement.Code) ?? string.Empty, Bound(requirement.Prompt) ?? string.Empty))
        .ToArray();

    [GeneratedRegex(@"^\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6})(?::\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6}))?$", RegexOptions.CultureInvariant)]
    private static partial Regex A1RangeRegex();

    [GeneratedRegex(@"^[A-Za-z_][A-Za-z0-9_]*$", RegexOptions.CultureInvariant)]
    private static partial Regex VbaIdentifierRegex();

    [GeneratedRegex(@"^[0-9A-Fa-f]{64}$", RegexOptions.CultureInvariant)]
    private static partial Regex Sha256Regex();
}
