namespace ExcelTask.Excel.Tests;

/// <summary>
/// The two string rules the copy rebind runs on, asserted directly rather than through Excel.
///
/// The rebind reads a sheet name out of an external reference Excel wrote, then writes that name
/// back into a formula. Both halves had a rule that was almost right, and neither could be caught
/// by the real-Excel fixtures, whose sheets are called Data and Exhibit - names that need no
/// quoting and carry no apostrophe. A table of names is what pins the guarantee; a fixture pins
/// one example of it.
///
/// Quoting was "every character is alphanumeric or underscore", which is not Excel's rule. A tab
/// called 2024 came back bare and produced =2024!A1, which Excel rejects - and a rejection there
/// arrives after the sheet has already been copied, so the caller gets Unknown on a workbook they
/// now have to reconcile by hand, for an ordinary fiscal-year tab.
///
/// Reading was TrimEnd('\''), which leaves the apostrophes Excel doubles inside a quoted name. A
/// sheet called Payer's Data was read as Payer''s Data, matched no worksheet in the target, and so
/// was reported as unfixable while naming a worksheet that does not exist.
/// </summary>
public sealed class CopyRebindSheetNameTests
{
    [Theory]
    // Left alone: ordinary names that Excel reads as names.
    [InlineData("Data", "Data")]
    [InlineData("Exhibit_A", "Exhibit_A")]
    [InlineData("Q1Actuals", "Q1Actuals")]
    // Quoted for punctuation or spacing.
    [InlineData("Payer Data", "'Payer Data'")]
    [InlineData("FY-2024", "'FY-2024'")]
    // Quoted because a leading digit makes the bare form a broken formula.
    [InlineData("2024", "'2024'")]
    [InlineData("2024Actuals", "'2024Actuals'")]
    // Quoted because Excel would read the bare form as a reference.
    [InlineData("A1", "'A1'")]
    [InlineData("R1", "'R1'")]
    [InlineData("C", "'C'")]
    [InlineData("ABC123", "'ABC123'")]
    // Quoted and doubled: the apostrophe would otherwise close the quoted name early.
    [InlineData("Payer's Data", "'Payer''s Data'")]
    public void ASheetNameIsQuotedExactlyWhenExcelWouldNotReadItAsAName(string sheetName, string expected) =>
        Assert.Equal(expected, ExcelWorkbookRuntime.QuoteSheetIfNeeded(sheetName));

    [Fact]
    public void ANameExcelQuotedIsReadBackAsTheNameTheWorksheetActuallyHas()
    {
        // Exactly what Excel writes into the copied sheet: the workbook in brackets, the sheet
        // after it, and every embedded apostrophe doubled.
        var formulas = new object[,]
        {
            { "='[template.xlsx]Payer''s Data'!A1*2", "=SUM('[template.xlsx]FY 2024'!A1:A9)" },
            { "=[template.xlsx]Data!B2", "=1+1" },
        };

        var names = ExcelWorkbookRuntime.ExternalSheetNames(formulas, "template.xlsx");

        // Payer's Data, not Payer''s Data - the second matches no worksheet that can exist, so the
        // sheet was left external and the receipt named a worksheet the caller could not find.
        Assert.Equal(["Data", "FY 2024", "Payer's Data"], names);
    }

    [Fact]
    public void AWorkbookWithNoExternalReferencesYieldsNothing() =>
        Assert.Empty(ExcelWorkbookRuntime.ExternalSheetNames(
            new object[,] { { "=Data!A1", "=SUM(A1:A9)" } }, "template.xlsx"));

    [Fact]
    public void ARoundTripLeavesAQuotedNameAsExcelWroteIt()
    {
        // The rebind reads a name and writes it straight back; if the two rules disagree the
        // replacement text is a formula Excel refuses. This is the property that ties them.
        const string original = "Payer's Data";
        var written = $"'[template.xlsx]{original.Replace("'", "''", StringComparison.Ordinal)}'!A1";

        var read = Assert.Single(ExcelWorkbookRuntime.ExternalSheetNames(new object[,] { { written } }, "template.xlsx"));

        Assert.Equal(original, read);
        Assert.Equal("'Payer''s Data'", ExcelWorkbookRuntime.QuoteSheetIfNeeded(read));
    }
}
