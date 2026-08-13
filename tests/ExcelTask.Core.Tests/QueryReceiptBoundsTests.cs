using ExcelTask.Core;

namespace ExcelTask.Core.Tests;

/// <summary>
/// The bound on a returned M expression, which 0.20.0 made a thing that exists at all.
///
/// The rule under test is omit-not-truncate, and it is the one worth pinning: a caller reads this
/// expression in order to send a replacement back, so a clipped one is not a smaller answer but a
/// wrong one that looks complete. The worker protocol once truncated macro source while the layers
/// downstream omitted it, and a source cut to exactly the limit then measured as within it and
/// passed as whole. This is the same seam, so it gets the test that defect earned.
/// </summary>
public sealed class QueryReceiptBoundsTests
{
    private static QueryReceipt Receipt(string formula) =>
        new("SalesData", MacroProcedureText.ComputeSha256(formula), formula.Length, formula);

    [Fact]
    public void AnExpressionWithinTheBoundSurvivesWhole()
    {
        const string formula = "let Source = Sql.Database(\"reporting\", \"finance\") in Source";

        var bounded = ReceiptBounds.Query(Receipt(formula), includeFormula: true, ReceiptBounds.MaxModelTextLength);

        // Not cut to MaxModelTextLength: the expression is the answer, and it has its own bound.
        Assert.Equal(formula, bounded!.Formula);
        Assert.Equal(formula.Length, bounded.Length);
    }

    [Fact]
    public void AnOversizedExpressionIsOmittedRatherThanCutAndLengthStillTellsTheTruth()
    {
        var formula = new string('m', MacroProcedureText.MaxSourceCharacters + 1);

        var bounded = ReceiptBounds.Query(Receipt(formula), includeFormula: true, ReceiptBounds.MaxModelTextLength);

        Assert.Null(bounded!.Formula);
        // The size is what remains true when the text is gone; a caller that sees a length and no
        // expression knows to open the query in Excel rather than that the query was empty.
        Assert.Equal(MacroProcedureText.MaxSourceCharacters + 1, bounded.Length);
        Assert.NotEmpty(bounded.Sha256);
    }

    [Fact]
    public void ApplyCarriesNoExpression()
    {
        const string formula = "let Source = Web.Contents(\"https://api.example.com\") in Source";

        var bounded = ReceiptBounds.Query(Receipt(formula), includeFormula: false, ReceiptBounds.MaxModelTextLength);

        // The same split macro source uses: Apply's answer is what changed, not what was there.
        Assert.Null(bounded!.Formula);
        Assert.Equal("SalesData", bounded.QueryName);
    }

    [Fact]
    public void LineEndingsAreNormalizedSoTheExpressionMatchesWhatWasFingerprinted()
    {
        const string lf = "let\n    Source = #table({\"A\"}, {{1}})\nin\n    Source";
        var crlf = lf.Replace("\n", "\r\n", StringComparison.Ordinal);

        var bounded = ReceiptBounds.Query(Receipt(crlf), includeFormula: true, ReceiptBounds.MaxModelTextLength);

        // A caller compares the expression it received against the fingerprint it received. If this
        // seam returned CRLF while the fingerprint was taken over LF, the two would disagree and the
        // Apply would be rejected for a difference no one made.
        Assert.Equal(lf, bounded!.Formula);
    }
}
