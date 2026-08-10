using ExcelTask.Core;

namespace ExcelTask.Core.Tests;

public sealed class FormulaMutationPlannerTests
{
    [Fact]
    public void RightPlansBlankPeriodsForEveryLaneWithoutMutatingGrid()
    {
        var grid = new[,] { { FormulaGridCell.Formula("=RC[-1]"), FormulaGridCell.Formula("=RC[-1]"), FormulaGridCell.Blank, FormulaGridCell.Blank } };
        var plan = FormulaMutationPlanner.Plan(grid, FormulaExtensionDirection.Right, 2);
        Assert.Equal([(0, 2), (0, 3)], plan.Mutations.Select(m => (m.RowIndex, m.ColumnIndex)));
        Assert.Equal(FormulaCellKind.Blank, grid[0, 2].Kind);
    }

    [Fact]
    public void DownRejectsMismatchedEvidence()
    {
        var grid = new[,] { { FormulaGridCell.Formula("=R[-1]C"), FormulaGridCell.Constant }, { FormulaGridCell.Formula("=R[-1]C"), FormulaGridCell.Blank }, { FormulaGridCell.Blank, FormulaGridCell.Blank } };
        Assert.Throws<ArgumentException>(() => FormulaMutationPlanner.Plan(grid, FormulaExtensionDirection.Down, 1));
    }

    [Fact]
    public void RejectsNonBlankDestination()
    {
        var grid = new[,] { { FormulaGridCell.Formula("=RC[-1]"), FormulaGridCell.Formula("=RC[-1]"), FormulaGridCell.Constant } };
        Assert.Throws<ArgumentException>(() => FormulaMutationPlanner.Plan(grid, FormulaExtensionDirection.Right, 1));
    }

    [Fact]
    public void RejectsPlansAboveTwoThousandMutations()
    {
        var oversized = new FormulaGridCell[84, 26];
        for (var row = 0; row < oversized.GetLength(0); row++)
        {
            for (var column = 0; column < oversized.GetLength(1); column++)
            {
                oversized[row, column] = FormulaGridCell.Blank;
            }
        }

        for (var row = 0; row < oversized.GetLength(0); row++)
        {
            oversized[row, 0] = FormulaGridCell.Formula("=RC[-1]");
            oversized[row, 1] = FormulaGridCell.Formula("=RC[-1]");
        }

        Assert.Throws<ArgumentException>(() => FormulaMutationPlanner.Plan(oversized, FormulaExtensionDirection.Right, 24));
    }

    [Fact]
    public void DownPlansBlankPeriodsForEveryLane()
    {
        var grid = new[,]
        {
            { FormulaGridCell.Formula("=R[-1]C"), FormulaGridCell.Formula("=R[-1]C") },
            { FormulaGridCell.Formula("=R[-1]C"), FormulaGridCell.Formula("=R[-1]C") },
            { FormulaGridCell.Blank, FormulaGridCell.Blank },
            { FormulaGridCell.Blank, FormulaGridCell.Blank }
        };

        var plan = FormulaMutationPlanner.Plan(grid, FormulaExtensionDirection.Down, 2);

        Assert.Equal([(2, 0), (3, 0), (2, 1), (3, 1)], plan.Mutations.Select(m => (m.RowIndex, m.ColumnIndex)));
    }

    [Fact]
    public void ReturnsReadOnlyMutationCollection()
    {
        var grid = new[,] { { FormulaGridCell.Formula("=RC[-1]"), FormulaGridCell.Formula("=RC[-1]"), FormulaGridCell.Blank } };
        var plan = FormulaMutationPlanner.Plan(grid, FormulaExtensionDirection.Right, 1);

        var mutable = Assert.IsAssignableFrom<IList<FormulaMutation>>(plan.Mutations);
        Assert.Throws<NotSupportedException>(() => mutable.Add(new FormulaMutation(0, 2, "=RC[-1]")));
    }
}
