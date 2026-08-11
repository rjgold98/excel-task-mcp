namespace ExcelTask.Excel.Tests;

/// <summary>
/// The address-length bound on a repair batch, asserted directly.
///
/// Excel refuses a Range address argument longer than 255 characters. The fix for that is a bound
/// on the joined length rather than on a cell count - but every repair fixture in the suite is
/// small enough to produce a single batch, so the split never ran and the fix was not load-bearing.
/// Swapping it back for any fixed cell count would have left the whole suite green while the
/// recorded failure returned exactly as its comment describes: identical repairs succeeding near
/// row 1 and throwing near row 2500, because the same number of cells makes a longer address
/// further down the sheet.
/// </summary>
public sealed class FormulaBatchAddressTests
{
    [Fact]
    public void NoBatchExceedsTheAddressLengthExcelAccepts()
    {
        // Fifty scattered cells at four- and five-character addresses - the shape that overflows.
        var repairs = Enumerable.Range(0, 50)
            .Select(index => new ExcelWorkbookRuntime.ExpectedFormula(2400 + (index * 2), 27 + index, "=RC[-1]*2"))
            .ToArray();

        var batches = ExcelWorkbookRuntime.BatchAddresses(repairs).ToArray();

        Assert.True(batches.Length > 1, "The fixture must be large enough to force a split; otherwise it asserts nothing.");
        Assert.All(batches, batch => Assert.InRange(batch.Length, 1, 255));
    }

    [Fact]
    public void EveryRepairAppearsExactlyOnceAcrossTheBatches()
    {
        // A bound that drops or duplicates a cell would be worse than the overflow it replaced: the
        // receipt counts repairs, so a lost one is a silent partial reported as complete.
        var repairs = Enumerable.Range(0, 137)
            .Select(index => new ExcelWorkbookRuntime.ExpectedFormula(1000 + index, 1 + (index % 40), "=RC[-1]"))
            .ToArray();

        var addresses = ExcelWorkbookRuntime.BatchAddresses(repairs)
            .SelectMany(batch => batch.Split(','))
            .ToArray();

        Assert.Equal(repairs.Length, addresses.Length);
        Assert.Equal(
            repairs.Select(repair => WorkbookRuntimeHelpers.ToA1Address(repair.Row, repair.Column)),
            addresses);
    }

    [Fact]
    public void ASingleAddressLongerThanTheBoundIsStillEmittedRatherThanDropped()
    {
        // Degenerate, but the alternative is silently repairing nothing: one address can never be
        // split, so it has to go out on its own and let Excel answer for it.
        var repairs = new[] { new ExcelWorkbookRuntime.ExpectedFormula(1048576, 16384, "=1") };

        var batches = ExcelWorkbookRuntime.BatchAddresses(repairs).ToArray();

        Assert.Single(batches);
        Assert.Equal(WorkbookRuntimeHelpers.ToA1Address(1048576, 16384), batches[0]);
    }

    [Fact]
    public void NoRepairsProducesNoBatches()
    {
        Assert.Empty(ExcelWorkbookRuntime.BatchAddresses([]));
    }

    [Fact]
    public void EvidenceIsReadOneCellWiderThanTheRangeRequested()
    {
        // The chunk-boundary fix, asserted where it can actually fail. A caller splitting a large
        // area into chunks puts gaps on the seams, and a gap on a seam has one of its two
        // neighbours in the other chunk. Reading only the requested range made those cells
        // unrepairable while the receipt still said Completed - silent loss, reported as success.
        var evidence = ExcelWorkbookRuntime.EvidenceBoundsFor(new FormulaRangeBounds(10, 4, 20, 6));

        Assert.Equal(9, evidence.StartRow);
        Assert.Equal(3, evidence.StartColumn);
        Assert.Equal(21, evidence.EndRow);
        Assert.Equal(7, evidence.EndColumn);
    }

    [Fact]
    public void EvidenceStopsAtTheEdgesOfTheSheet()
    {
        // Widening must never produce an address Excel refuses, which would turn a repair that
        // worked into a rejection at the one place a caller cannot move the data away from.
        var topLeft = ExcelWorkbookRuntime.EvidenceBoundsFor(new FormulaRangeBounds(1, 1, 3, 3));
        Assert.Equal(1, topLeft.StartRow);
        Assert.Equal(1, topLeft.StartColumn);

        var bottomRight = ExcelWorkbookRuntime.EvidenceBoundsFor(new FormulaRangeBounds(1_048_575, 16_383, 1_048_576, 16_384));
        Assert.Equal(1_048_576, bottomRight.EndRow);
        Assert.Equal(16_384, bottomRight.EndColumn);
    }
}
