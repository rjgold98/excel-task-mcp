using System.Diagnostics;
using System.Text.RegularExpressions;

namespace ExcelTask.Core;

public sealed partial class ExcelTaskEngine(IWorkbookRuntime runtime) : IExcelTaskEngine
{
    /// <summary>Maximum aggregate number of cells requested for formula repair in the MVP.</summary>
    public const int MaxFormulaRepairCells = 10_000;

    private const int MaxReceiptItems = 20;
    private const int MaxReceiptStringLength = 256;

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
                    normalizedRequest.ReferenceWorkbookPath,
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
                total.Elapsed);
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            throw;
        }
        catch (Exception exception)
        {
            executionTimer.Stop();
            return CreateReceipt(
                taskId,
                ExcelTaskStatus.Unknown,
                "Workbook execution did not complete.",
                [],
                CombineChecks(inspection.Checks, [new TaskCheck("runtime-execution", false, exception.Message)]),
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
        TimeSpan total)
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
            new PhaseTimings(validation, inspection, execution, total));
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

        if (!TryNormalizeWorkbookPath(request.TargetWorkbookPath, "Target workbook path", out var target, out error) ||
            !TryNormalizeWorkbookPath(request.ReferenceWorkbookPath, "Reference workbook path", out var reference, out error) ||
            !TryNormalizeWorksheetName(request.ReferenceWorksheet, "Reference worksheet", out var referenceSheet, out error) ||
            !TryNormalizeWorksheetName(request.NewWorksheetName, "New worksheet name", out var newSheet, out error) ||
            !TryNormalizeRanges(request.FormulaRepairRanges, out var ranges, out error))
        {
            return false;
        }

        if (!IsSupportedWorkbookPath(target!) || !IsSupportedWorkbookPath(reference!))
        {
            error = "MVP workbook paths must use a .xlsx or .xlsm extension.";
            return false;
        }

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

        normalized = new NormalizedExcelTaskRequest(
            target!, reference!, referenceSheet!, newSheet!, ranges!, request.Mode, request.WorkbookBinding,
            request.Save, output, request.OverwriteConfirmed);
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

    private static bool TryNormalizeRanges(IReadOnlyList<string>? values, out IReadOnlyList<FormulaRepairRange>? ranges, out string? error)
    {
        ranges = null;
        error = null;
        if (values is null)
        {
            error = "Formula repair ranges are required; supply an empty list when no repair range is needed.";
            return false;
        }

        var normalized = new List<FormulaRepairRange>(values.Count);
        long aggregateCellCount = 0;
        foreach (var value in values)
        {
            var match = A1RangeRegex().Match(value?.Trim() ?? string.Empty);
            var endColumnText = match.Success && match.Groups[3].Success ? match.Groups[3].Value : match.Groups[1].Value;
            var endRowText = match.Success && match.Groups[4].Success ? match.Groups[4].Value : match.Groups[2].Value;
            if (!match.Success ||
                !TryParseCell(match.Groups[1].Value, match.Groups[2].Value, out var startColumn, out var startRow) ||
                !TryParseCell(endColumnText, endRowText, out var endColumn, out var endRow) ||
                startColumn > endColumn || startRow > endRow)
            {
                error = $"Formula repair range '{value}' is invalid. Use a rectangular A1 range such as A1:C10.";
                return false;
            }

            aggregateCellCount += (long)(endColumn - startColumn + 1) * (endRow - startRow + 1);
            if (aggregateCellCount > MaxFormulaRepairCells)
            {
                error = $"Formula repair ranges exceed the MVP limit of {MaxFormulaRepairCells:N0} aggregate cells.";
                return false;
            }

            normalized.Add(new FormulaRepairRange(
                $"{match.Groups[1].Value.ToUpperInvariant()}{startRow}",
                $"{endColumnText.ToUpperInvariant()}{endRow}"));
        }

        ranges = normalized;
        return true;
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
}
