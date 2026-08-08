using ExcelTask.Core;

namespace ExcelTask.Core.Tests;

public sealed class FormulaPatternAnalyzerTests
{
    [Fact]
    public void InferRepairsRepairsBlankCellBracketedVerticallyByTheSameFormula()
    {
        var repairs = FormulaPatternAnalyzer.InferRepairs(new[,]
        {
            { FormulaGridCell.Formula("=RC[-1]" ) },
            { FormulaGridCell.Blank },
            { FormulaGridCell.Formula("=RC[-1]" ) }
        });

        var repair = Assert.Single(repairs);
        Assert.Equal((1, 0), (repair.RowIndex, repair.ColumnIndex));
        Assert.Equal("=RC[-1]", repair.FormulaR1C1);
        Assert.Equal(FormulaRepairEvidence.Vertical, repair.Evidence);
    }

    [Fact]
    public void InferRepairsRepairsBlankCellBracketedHorizontallyByTheSameFormula()
    {
        var repairs = FormulaPatternAnalyzer.InferRepairs(new[,]
        {
            { FormulaGridCell.Formula("=RC[-1]"), FormulaGridCell.Blank, FormulaGridCell.Formula("=RC[-1]") }
        });

        Assert.Equal(FormulaRepairEvidence.Horizontal, Assert.Single(repairs).Evidence);
    }

    [Fact]
    public void InferRepairsNeverOverwritesConstantsOrExistingDifferentFormulas()
    {
        var repairs = FormulaPatternAnalyzer.InferRepairs(new[,]
        {
            { FormulaGridCell.Formula("=R[-1]C"), FormulaGridCell.Constant, FormulaGridCell.Formula("=R[-1]C") },
            { FormulaGridCell.Formula("=R[-1]C"), FormulaGridCell.Formula("=RC[-1]"), FormulaGridCell.Formula("=R[-1]C") }
        });

        Assert.Empty(repairs);
    }

    [Fact]
    public void InferRepairsDoesNotRepairEdgesOrAmbiguousCrossPatterns()
    {
        var repairs = FormulaPatternAnalyzer.InferRepairs(new[,]
        {
            { FormulaGridCell.Blank, FormulaGridCell.Formula("=RC[-1]"), FormulaGridCell.Blank },
            { FormulaGridCell.Formula("=R[-1]C"), FormulaGridCell.Blank, FormulaGridCell.Formula("=R[-1]C") },
            { FormulaGridCell.Blank, FormulaGridCell.Formula("=RC[-1]"), FormulaGridCell.Blank }
        });

        Assert.Empty(repairs);
    }

    [Fact]
    public void InferRepairsMarksAgreeingVerticalAndHorizontalEvidence()
    {
        var repairs = FormulaPatternAnalyzer.InferRepairs(new[,]
        {
            { FormulaGridCell.Constant, FormulaGridCell.Formula("=RC[-1]"), FormulaGridCell.Constant },
            { FormulaGridCell.Formula("=RC[-1]"), FormulaGridCell.Blank, FormulaGridCell.Formula("=RC[-1]") },
            { FormulaGridCell.Constant, FormulaGridCell.Formula("=RC[-1]"), FormulaGridCell.Constant }
        });

        Assert.Equal(FormulaRepairEvidence.Both, Assert.Single(repairs).Evidence);
    }
}
