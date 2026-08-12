namespace ExcelTask.Core;

/// <summary>
/// The one implementation of receipt bounding.
///
/// Before this module the same receipt was bounded in three modules - the worker protocol, the
/// engine, and the MCP tool - with four independently declared item caps of 20, three string caps
/// of which one was dead (the engine's 256 could never fire behind the protocol's 128), and one
/// genuine disagreement: the protocol truncated an oversized macro source while the engine and the
/// tool deliberately omit it, each carrying a comment saying truncation is the dangerous option.
/// A source the protocol had clipped to exactly the limit arrived downstream measuring within it
/// and passed as complete, defeating both layers that had decided otherwise.
///
/// The three call sites remain - the protocol bounds for its frame budget, the engine and the tool
/// for the model - because each seam has a reason to distrust the layers behind it. What they no
/// longer have is their own opinion about what bounding means: the caps, the truncation flags, and
/// the omit-not-truncate rule for macro source are decided here, once.
/// </summary>
public static class ReceiptBounds
{
    /// <summary>Items carried per receipt list; totals still report what existed.</summary>
    public const int MaxItems = 20;

    /// <summary>
    /// The string cap a model-facing seam applies - what the caller actually sees.
    ///
    /// 256, not the 96 the MCP tool used to enforce. Consolidating surfaced what that 96 was
    /// doing: engine-authored rejection summaries never cross the worker pipe, so the four-part
    /// macro policy rejection - written so one resubmission can fix everything at once, a rule the
    /// interface study paid for - reached the model cut mid-word at 96 characters, missing its
    /// fourth requirement and the instruction. The engine's own 256-character test passed the whole
    /// time because it read the engine directly. Runtime-authored strings still arrive pre-cut at
    /// <see cref="MaxFrameTextLength"/>; this cap governs what the engine itself writes.
    /// </summary>
    public const int MaxModelTextLength = 256;

    /// <summary>
    /// The cap the worker protocol applies inside its 64 KB frame budget, for strings authored in
    /// the worker. The protocol must never be the layer that decides what a model reads, only that
    /// a frame fits.
    /// </summary>
    public const int MaxFrameTextLength = 128;

    public static string? Text(string? value, int maxText) => value is null
        ? null
        : value.Length <= maxText ? value : value[..maxText];

    public static string RequiredText(string? value, int maxText) => Text(value, maxText) ?? string.Empty;

    public static TaskChange[] Changes(IReadOnlyList<TaskChange>? changes, int maxText) => (changes ?? [])
        .Take(MaxItems)
        .Select(change => new TaskChange(RequiredText(change.Kind, maxText), RequiredText(change.Target, maxText), RequiredText(change.Summary, maxText)))
        .ToArray();

    public static TaskCheck[] Checks(IReadOnlyList<TaskCheck>? checks, int maxText) => (checks ?? [])
        .Take(MaxItems)
        .Select(check => new TaskCheck(RequiredText(check.Name, maxText), check.Passed, RequiredText(check.Detail, maxText)))
        .ToArray();

    public static ConfirmationRequirement[] Requirements(IReadOnlyList<ConfirmationRequirement>? requirements, int maxText) => (requirements ?? [])
        .Take(MaxItems)
        .Select(requirement => new ConfirmationRequirement(RequiredText(requirement.Code, maxText), RequiredText(requirement.Prompt, maxText)))
        .ToArray();

    /// <summary>
    /// The one receipt field that deliberately carries workbook contents. Cells are capped at the
    /// read limit rather than <see cref="MaxItems"/>, because they are the answer the caller asked
    /// for - trimming them to twenty would not be a bound, it would be a wrong result. Cell text is
    /// capped far shorter than the range is deep: a long cell is usually a pasted paragraph, and one
    /// of them must not crowd out the four hundred cells around it.
    ///
    /// Bounding rewrites two fields and must carry the rest through untouched. Constructing a fresh
    /// <see cref="WorksheetCell"/> here silently reset <c>IsFormula</c> to its default of false on
    /// every cell, at all three seams - so every read told the caller that nothing in the range was
    /// a formula. `with` is used rather than a constructor precisely so that a field added later
    /// survives by default instead of being dropped by omission.
    /// </summary>
    public static WorksheetRangeReceipt? Range(WorksheetRangeReceipt? range, int maxText)
    {
        if (range is null) return null;

        var cells = range.Cells
            .Take(ExcelTaskEngine.MaxReadCells)
            .Select(cell => cell with
            {
                Address = RequiredText(cell.Address, maxText),
                Text = cell.Text.Length <= ExcelTaskEngine.MaxReadCellTextLength
                    ? cell.Text
                    : cell.Text[..ExcelTaskEngine.MaxReadCellTextLength]
            })
            .ToArray();

        return range with
        {
            WorksheetName = RequiredText(range.WorksheetName, maxText),
            Range = RequiredText(range.Range, maxText),
            Cells = cells,
            Truncated = range.Truncated || range.Cells.Count > cells.Length
        };
    }

    /// <summary>The real total survives the cap, so a trimmed report still says what it is not showing.</summary>
    public static WorkbookAuditReceipt? Audit(WorkbookAuditReceipt? audit, int maxText)
    {
        if (audit is null) return null;

        var items = audit.Items
            .Take(MaxItems)
            .Select(item => new WorkbookFlowItem(
                RequiredText(item.Kind, maxText),
                RequiredText(item.Name, maxText),
                RequiredText(item.Detail, maxText),
                Text(item.DependsOn, maxText)))
            .ToArray();

        return new WorkbookAuditReceipt(items, audit.TotalFound, audit.Truncated || audit.Items.Count > items.Length, audit.WorkbookUnchanged);
    }

    /// <summary>
    /// Macro source is omitted when oversized, never truncated - at every seam, which is the point.
    /// A partial procedure could lead a caller to write a replacement from incomplete source and
    /// destroy the remainder, so the layers agreed that no source is safer than half a source; this
    /// method is where that agreement stops being three separate opinions. Metadata is length-capped
    /// like any other text; <paramref name="includeSource"/> is the caller's policy (Plan carries
    /// evidence, Apply never echoes source back).
    /// </summary>
    public static MacroProcedureReceipt? MacroProcedure(MacroProcedureReceipt? receipt, bool includeSource, int maxText)
    {
        if (receipt is null) return null;

        string? source = null;
        if (includeSource && receipt.Source is not null)
        {
            var normalized = MacroProcedureText.NormalizeLineEndings(receipt.Source);
            if (normalized.Length <= MacroProcedureText.MaxSourceCharacters) source = normalized;
        }

        return new MacroProcedureReceipt(
            RequiredText(receipt.ComponentName, maxText),
            RequiredText(receipt.ProcedureName, maxText),
            RequiredText(receipt.Sha256, maxText),
            source,
            receipt.RunRequested,
            receipt.RunCompleted);
    }
}
