using System.Globalization;
using ExcelTask.Core;

namespace ExcelTask.Excel;

public sealed partial class ExcelWorkbookRuntime
{
    // Excel's own border indices and line weights. Named here so the mapping from the caller's
    // words to Excel's numbers lives in one place rather than inline at the assignment.
    private const int XlEdgeLeft = 7;
    private const int XlEdgeTop = 8;
    private const int XlEdgeBottom = 9;
    private const int XlEdgeRight = 10;
    private const int XlInsideVertical = 11;
    private const int XlInsideHorizontal = 12;
    private const int XlContinuous = 1;
    private const int XlLineStyleNone = -4142;
    private const int XlHairline = 1;
    private const int XlThin = 2;
    private const int XlMedium = -4138;
    private const int XlThick = 4;
    private const int XlNone = -4142;

    /// <summary>
    /// Sets how one bounded range looks, and proves every part of it took.
    ///
    /// Formatting is one COM assignment per property no matter how many cells it covers - Excel
    /// applies each to the whole range at once - so there is no batching and no per-cell cost. What
    /// it does need is the read-back, because Excel is free to store something other than what it
    /// was given: an unrecognized number format can be kept verbatim, coerced, or rejected, and a
    /// font name that is not installed is silently substituted. All of those look identical from the
    /// caller's side of the assignment. Reading each back is what turns "we sent a format" into
    /// "the sheet holds this format".
    ///
    /// Only the fields the caller supplied are written, and only those are verified. Formatting has
    /// no recoverable prior state on the sheet, so touching a property nobody asked about would be
    /// destroying something the caller never offered to lose.
    /// </summary>
    private static WorkbookExecutionOutcome ExecuteRangeFormatCore(ExcelTaskPlan plan, IExcelWorkbookRuntimeObserver observer)
    {
        var operation = plan.Request.Operation.SetRangeFormat!;
        var target = $"{operation.WorksheetName}!{operation.Range}";
        var requested = DescribeRequest(operation);

        return ExecuteMutation(plan, observer, "range-format", "The range format", context =>
        {
            context.OnPhase("format-preflight");
            var preflight = PreflightWorksheetExists(context.Session, operation.WorksheetName);
            context.Checks.AddRange(preflight.Checks);
            if (!preflight.IsFeasible)
            {
                return new MutationStep.Finish(
                    new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected, "The requested worksheet was not found.", Checks: context.Checks),
                    "range format preflight");
            }

            // What is there now, so the caller can see what an Apply would replace. Formatting is
            // destructive in a way a value write is not - the old appearance is not recoverable from
            // the sheet afterwards. Excel answers a range whose cells disagree with null, which is
            // itself the answer rather than a failure to read.
            var existing = ReadFormat(context.Session, operation);
            context.Checks.Add(new TaskCheck("current-format", true, existing.Describe()));

            if (!context.Apply)
            {
                context.Changes.Add(new TaskChange("range-format", target, $"Planned {requested}."));
                return new MutationStep.Finish(
                    new WorkbookExecutionOutcome(ExcelTaskStatus.Planned,
                        $"Applying would set {target}: {requested}. Nothing was changed.",
                        context.Changes, context.Checks),
                    "range format planning");
            }

            context.OnPhase("range-format");
            context.MarkMutationAttempted();
            ApplyFormat(context.Session, operation);

            var stored = ReadFormat(context.Session, operation);
            var mismatch = FirstMismatch(operation, stored);
            if (mismatch is not null)
            {
                context.Checks.Add(new TaskCheck("range-format", false, mismatch));
                return new MutationStep.Finish(
                    new WorkbookExecutionOutcome(ExcelTaskStatus.Unknown,
                        "Excel did not store the formatting as requested; nothing was saved.",
                        context.Changes, context.Checks,
                        CanRetry: false, RetryReason: "Correct the request and inspect the workbook before retrying."),
                    "the range format");
            }

            context.Checks.Add(new TaskCheck("range-format", true, $"Applied {requested} and read every part of it back unchanged."));
            context.Changes.Add(new TaskChange("range-format", target, $"Set {requested}."));

            return new MutationStep.SaveAndVerify(
                verification => FirstMismatch(operation, ReadFormat(verification, operation)) is { } detail
                    ? (false, new TaskCheck("reopen-verification", false, $"After reopening the saved workbook: {detail}"))
                    : (true, new TaskCheck("reopen-verification", true, "The saved workbook reopened with every requested format across the range.")),
                $"Set {target}: {requested}, saved, and confirmed it after reopening.",
                "Excel saved the workbook, but reopen verification did not confirm the formatting.");
        });
    }

    /// <summary>What the caller asked for, in their words rather than Excel's, for the receipt.</summary>
    private static string DescribeRequest(NormalizedSetRangeFormatOperation operation)
    {
        var parts = new List<string>(8);
        if (operation.NumberFormat is not null) parts.Add($"number format {operation.NumberFormat}");
        if (operation.Bold is { } bold) parts.Add(bold ? "bold" : "not bold");
        if (operation.Italic is { } italic) parts.Add(italic ? "italic" : "not italic");
        if (operation.FontSize is { } size) parts.Add($"{size.ToString("0.##", CultureInfo.InvariantCulture)}pt");
        if (operation.FontName is not null) parts.Add($"font {operation.FontName}");
        if (operation.FontColor is not null) parts.Add("text colour");
        if (operation.FillColor is { } fill) parts.Add(fill == ExcelTaskEngine.NoFillColor ? "no fill" : "fill colour");
        if (operation.Borders is not RangeBorderEdges.Unspecified) parts.Add($"{operation.Borders} borders");
        if (operation.ColumnWidth is { } width) parts.Add($"column width {width.ToString("0.##", CultureInfo.InvariantCulture)}");
        if (operation.RowHeight is { } height) parts.Add($"row height {height.ToString("0.##", CultureInfo.InvariantCulture)}");
        return string.Join(", ", parts);
    }

    /// <summary>
    /// Every property this operation can set, as Excel currently holds it. Null means the range's
    /// cells do not agree on that property, which Excel reports rather than erroring.
    /// </summary>
    private sealed record RangeFormatSnapshot(
        string? NumberFormat,
        bool? Bold,
        bool? Italic,
        double? FontSize,
        string? FontName,
        int? FontColor,
        int? FillColor,
        bool BordersDrawn,
        double? ColumnWidth,
        double? RowHeight)
    {
        public string Describe()
        {
            var parts = new List<string>(6)
            {
                NumberFormat is null ? "mixed number formats" : $"number format {NumberFormat}"
            };
            if (Bold is { } bold) parts.Add(bold ? "bold" : "not bold");
            if (FontSize is { } size) parts.Add($"{size.ToString("0.##", CultureInfo.InvariantCulture)}pt");
            if (FontName is not null) parts.Add(FontName);
            parts.Add(BordersDrawn ? "some borders" : "no borders");
            return $"The range currently has {string.Join(", ", parts)}.";
        }
    }

    private static RangeFormatSnapshot ReadFormat(ExcelSession session, NormalizedSetRangeFormatOperation operation)
    {
        using var references = new ComReferenceScope();
        var sheets = references.Add(Get(session.TargetWorkbook, "Worksheets"));
        var sheet = references.Add(Item(sheets, operation.WorksheetName));
        var range = references.Add(Get(sheet, "Range", operation.Range.ToString()));
        var font = references.Add(Get(range, "Font"));
        var interior = references.Add(Get(range, "Interior"));

        return new RangeFormatSnapshot(
            GetOrNull(range, "NumberFormat") as string,
            GetOrNull(font, "Bold") as bool?,
            GetOrNull(font, "Italic") as bool?,
            ToDouble(GetOrNull(font, "Size")),
            GetOrNull(font, "Name") as string,
            ToInt(GetOrNull(font, "Color")),
            ToInt(GetOrNull(interior, "Color")),
            AnyBorderDrawn(references, range, operation.Borders),
            ToDouble(GetOrNull(range, "ColumnWidth")),
            ToDouble(GetOrNull(range, "RowHeight")));
    }

    private static void ApplyFormat(ExcelSession session, NormalizedSetRangeFormatOperation operation)
    {
        using var references = new ComReferenceScope();
        var sheets = references.Add(Get(session.TargetWorkbook, "Worksheets"));
        var sheet = references.Add(Item(sheets, operation.WorksheetName));
        var range = references.Add(Get(sheet, "Range", operation.Range.ToString()));

        if (operation.NumberFormat is not null) Set(range, "NumberFormat", operation.NumberFormat);

        if (operation.Bold is not null || operation.Italic is not null || operation.FontSize is not null ||
            operation.FontName is not null || operation.FontColor is not null)
        {
            var font = references.Add(Get(range, "Font"));
            if (operation.Bold is { } bold) Set(font, "Bold", bold);
            if (operation.Italic is { } italic) Set(font, "Italic", italic);
            if (operation.FontSize is { } size) Set(font, "Size", size);
            if (operation.FontName is not null) Set(font, "Name", operation.FontName);
            if (operation.FontColor is { } color) Set(font, "Color", color);
        }

        if (operation.FillColor is { } fill)
        {
            var interior = references.Add(Get(range, "Interior"));
            // Clearing is a pattern change, not a colour: assigning a colour would paint white,
            // which looks the same on screen and is not the same thing at all when the sheet has
            // banding or a theme behind it.
            if (fill == ExcelTaskEngine.NoFillColor) Set(interior, "Pattern", XlNone);
            else Set(interior, "Color", fill);
        }

        if (operation.Borders is not RangeBorderEdges.Unspecified) ApplyBorders(references, range, operation);
        if (operation.ColumnWidth is { } columnWidth) Set(range, "ColumnWidth", columnWidth);
        if (operation.RowHeight is { } rowHeight) Set(range, "RowHeight", rowHeight);
    }

    private static void ApplyBorders(ComReferenceScope references, object range, NormalizedSetRangeFormatOperation operation)
    {
        var weight = operation.BorderStyle switch
        {
            RangeBorderWeight.Hairline => XlHairline,
            RangeBorderWeight.Medium => XlMedium,
            RangeBorderWeight.Thick => XlThick,
            _ => XlThin
        };

        var borders = references.Add(Get(range, "Borders"));
        foreach (var index in BorderIndices(operation.Borders))
        {
            var border = references.Add(Item(borders, index));
            if (operation.Borders == RangeBorderEdges.None)
            {
                Set(border, "LineStyle", XlLineStyleNone);
                continue;
            }

            Set(border, "LineStyle", XlContinuous);
            Set(border, "Weight", weight);
        }
    }

    /// <summary>
    /// Which of Excel's border slots the caller's word covers. All includes the interior grid; an
    /// Outline deliberately does not, because "put a box round this" and "rule every cell" are
    /// different requests and conflating them is not recoverable once saved.
    /// </summary>
    private static int[] BorderIndices(RangeBorderEdges edges) => edges switch
    {
        RangeBorderEdges.Top => [XlEdgeTop],
        RangeBorderEdges.Bottom => [XlEdgeBottom],
        RangeBorderEdges.Left => [XlEdgeLeft],
        RangeBorderEdges.Right => [XlEdgeRight],
        RangeBorderEdges.Outline => [XlEdgeTop, XlEdgeBottom, XlEdgeLeft, XlEdgeRight],
        _ => [XlEdgeTop, XlEdgeBottom, XlEdgeLeft, XlEdgeRight, XlInsideVertical, XlInsideHorizontal]
    };

    private static bool AnyBorderDrawn(ComReferenceScope references, object range, RangeBorderEdges edges)
    {
        if (edges is RangeBorderEdges.Unspecified) return false;
        var borders = references.Add(Get(range, "Borders"));
        foreach (var index in BorderIndices(edges))
        {
            var border = references.Add(Item(borders, index));
            var style = ToInt(GetOrNull(border, "LineStyle"));
            if (style is not null && style != XlLineStyleNone) return true;
        }

        return false;
    }

    /// <summary>
    /// The first thing Excel stored that is not what was asked for, or null when every requested
    /// property matches. Only requested properties are compared - the rest were never ours.
    /// </summary>
    private static string? FirstMismatch(NormalizedSetRangeFormatOperation operation, RangeFormatSnapshot stored)
    {
        if (operation.NumberFormat is not null && !string.Equals(stored.NumberFormat, operation.NumberFormat, StringComparison.Ordinal))
        {
            return stored.NumberFormat is null
                ? "the range does not share one number format."
                : $"Excel stored the number format {stored.NumberFormat} rather than the one requested.";
        }

        if (operation.Bold is { } bold && stored.Bold != bold) return "the range does not share the requested bold setting.";
        if (operation.Italic is { } italic && stored.Italic != italic) return "the range does not share the requested italic setting.";
        if (operation.FontSize is { } size && !NearlyEqual(stored.FontSize, size)) return "the range does not share the requested font size.";

        // This catches a mixed range, and nothing more. Measured: Excel stores whatever font name
        // it is given, installed or not, and substitutes only when rendering - so the read-back
        // agrees with the request even for a font that does not exist on the machine. Verification
        // cannot tell a caller they misspelled Garamond, and the schema says so rather than
        // implying a guarantee this comparison does not provide.
        if (operation.FontName is not null && !string.Equals(stored.FontName, operation.FontName, StringComparison.OrdinalIgnoreCase))
        {
            return stored.FontName is null
                ? "the range does not share one font."
                : $"Excel stored the font {stored.FontName} rather than the one requested.";
        }

        if (operation.FontColor is { } fontColor && stored.FontColor != fontColor) return "the range does not share the requested text colour.";
        if (operation.FillColor is { } fill && fill != ExcelTaskEngine.NoFillColor && stored.FillColor != fill) return "the range does not share the requested fill colour.";
        if (operation.Borders is not (RangeBorderEdges.Unspecified or RangeBorderEdges.None) && !stored.BordersDrawn) return "the requested borders were not drawn.";
        if (operation.ColumnWidth is { } width && !NearlyEqual(stored.ColumnWidth, width)) return "the columns do not share the requested width.";
        if (operation.RowHeight is { } height && !NearlyEqual(stored.RowHeight, height)) return "the rows do not share the requested height.";
        return null;
    }

    // Excel rounds a width or height to what the display can render, so an exact comparison would
    // fail a request it actually honoured.
    private static bool NearlyEqual(double? stored, double requested) =>
        stored is not null && Math.Abs(stored.Value - requested) < 0.05;

    private static double? ToDouble(object? value) => value is null ? null : Convert.ToDouble(value, CultureInfo.InvariantCulture);

    private static int? ToInt(object? value) => value is null ? null : (int)Convert.ToDouble(value, CultureInfo.InvariantCulture);
}
