namespace ExcelTask.Core;

public enum FormulaCellKind
{
    Blank,
    Constant,
    Formula
}

public sealed record FormulaGridCell(FormulaCellKind Kind, string? FormulaR1C1 = null)
{
    public static FormulaGridCell Blank { get; } = new(FormulaCellKind.Blank);

    public static FormulaGridCell Constant { get; } = new(FormulaCellKind.Constant);

    public static FormulaGridCell Formula(string formulaR1C1) => new(FormulaCellKind.Formula, formulaR1C1);
}

public enum FormulaRepairEvidence
{
    Vertical,
    Horizontal,
    Both
}

public sealed record InferredFormulaRepair(
    int RowIndex,
    int ColumnIndex,
    string FormulaR1C1,
    FormulaRepairEvidence Evidence);

public static class FormulaPatternAnalyzer
{
    public static IReadOnlyList<InferredFormulaRepair> InferRepairs(FormulaGridCell[,] grid)
    {
        ArgumentNullException.ThrowIfNull(grid);
        var repairs = new List<InferredFormulaRepair>();
        for (var row = 0; row < grid.GetLength(0); row++)
        {
            for (var column = 0; column < grid.GetLength(1); column++)
            {
                if (grid[row, column]?.Kind != FormulaCellKind.Blank) continue;

                var vertical = GetMatchingFormula(grid, row - 1, column, row + 1, column);
                var horizontal = GetMatchingFormula(grid, row, column - 1, row, column + 1);
                if (vertical is null && horizontal is null) continue;
                if (vertical is not null && horizontal is not null && !string.Equals(vertical, horizontal, StringComparison.Ordinal)) continue;

                repairs.Add(new InferredFormulaRepair(
                    row,
                    column,
                    vertical ?? horizontal!,
                    vertical is not null && horizontal is not null
                        ? FormulaRepairEvidence.Both
                        : vertical is not null ? FormulaRepairEvidence.Vertical : FormulaRepairEvidence.Horizontal));
            }
        }

        return repairs;
    }

    private static string? GetMatchingFormula(FormulaGridCell[,] grid, int firstRow, int firstColumn, int secondRow, int secondColumn)
    {
        if (firstRow < 0 || firstColumn < 0 || secondRow >= grid.GetLength(0) || secondColumn >= grid.GetLength(1)) return null;
        var first = grid[firstRow, firstColumn];
        var second = grid[secondRow, secondColumn];
        return first?.Kind == FormulaCellKind.Formula &&
               second?.Kind == FormulaCellKind.Formula &&
               !string.IsNullOrEmpty(first.FormulaR1C1) &&
               string.Equals(first.FormulaR1C1, second.FormulaR1C1, StringComparison.Ordinal)
            ? first.FormulaR1C1
            : null;
    }
}
