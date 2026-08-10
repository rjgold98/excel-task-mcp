namespace ExcelTask.Core;

public sealed record FormulaMutation(int RowIndex, int ColumnIndex, string FormulaR1C1);
public sealed record FormulaMutationPlan(IReadOnlyList<FormulaMutation> Mutations);

/// <summary>Pure planner for bounded formula-series extension. It never mutates the input grid.</summary>
public static class FormulaMutationPlanner
{
    public const int MaxPeriods = 24;
    public const int MaxMutations = 2_000;

    public static FormulaMutationPlan Plan(FormulaGridCell[,] grid, FormulaExtensionDirection direction, int periods)
    {
        ArgumentNullException.ThrowIfNull(grid);
        if (!Enum.IsDefined(direction)) throw new ArgumentOutOfRangeException(nameof(direction));
        if (periods is < 1 or > MaxPeriods) throw new ArgumentOutOfRangeException(nameof(periods));

        var lanes = direction == FormulaExtensionDirection.Right ? grid.GetLength(0) : grid.GetLength(1);
        var required = 2 + periods;
        var available = direction == FormulaExtensionDirection.Right ? grid.GetLength(1) : grid.GetLength(0);
        if (available < required) throw new ArgumentException("Grid does not contain two evidence cells and the requested destination periods.", nameof(grid));
        if ((long)lanes * periods > MaxMutations) throw new ArgumentException("Planned mutations exceed the MVP limit.", nameof(periods));

        var result = new List<FormulaMutation>(lanes * periods);
        for (var lane = 0; lane < lanes; lane++)
        {
            var first = direction == FormulaExtensionDirection.Right ? grid[lane, 0] : grid[0, lane];
            var second = direction == FormulaExtensionDirection.Right ? grid[lane, 1] : grid[1, lane];
            if (first?.Kind != FormulaCellKind.Formula || second?.Kind != FormulaCellKind.Formula ||
                string.IsNullOrEmpty(first.FormulaR1C1) || !string.Equals(first.FormulaR1C1, second.FormulaR1C1, StringComparison.Ordinal))
                throw new ArgumentException("Each lane requires exactly two matching formula evidence cells.", nameof(grid));

            for (var period = 0; period < periods; period++)
            {
                var row = direction == FormulaExtensionDirection.Right ? lane : period + 2;
                var column = direction == FormulaExtensionDirection.Right ? period + 2 : lane;
                var destination = grid[row, column];
                if (destination?.Kind != FormulaCellKind.Blank)
                    throw new ArgumentException("Formula-series destinations must be blank.", nameof(grid));
                result.Add(new FormulaMutation(row, column, first.FormulaR1C1!));
            }
        }
        return new FormulaMutationPlan(Array.AsReadOnly(result.ToArray()));
    }
}
