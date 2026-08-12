using ExcelTask.Core;
using ExcelTask.McpServer;

namespace ExcelTask.McpServer.Tests;

/// <summary>
/// The field check's own coverage arithmetic - the first test of any kind over the ~900 lines that
/// validate a release on the locked-down work computer.
///
/// It shipped covering five of eleven operations, exercising everything from the first four
/// releases and nothing from the last four, and still reported PASS. The step list is written by
/// hand and will stay that way, since each step needs its own fixture; what it must never do again
/// is stay quiet about what it skipped.
/// </summary>
public sealed class FieldCheckCoverageTests
{
    private static OperationResult Result(string label) =>
        new(label, "Completed", 1.0, 0, "summary", "checks", null);

    [Fact]
    public void EveryOperationExercisedLeavesNothingUncovered()
    {
        var operations = OperationCatalog.AllKinds
            .Select(kind => Result($"{kind} (Apply)"))
            .ToArray();

        Assert.Empty(FieldCheckCoverage.UncoveredKinds(operations));
    }

    [Fact]
    public void TheGapThatShippedIsNamedRatherThanCounted()
    {
        // The five it actually ran, reconstructed. A caller reading the report needs the names, not
        // a ratio - "5 of 11" does not tell anyone which half was proven.
        var shipped = new[]
        {
            Result("CopyExhibit (Plan)"),
            Result("CopyExhibit (Apply)"),
            Result("ExtendFormulaSeries (Apply)"),
            Result("AuditWorkbookFlows (Apply)"),
            Result("EditMacroProcedure (Plan)"),
            Result("EditMacroProcedure (Apply+Run)")
        };

        var uncovered = FieldCheckCoverage.UncoveredKinds(shipped);

        Assert.Equal(
            ["RepairExistingWorksheet", "ReadWorksheetRange", "WriteWorksheetValues", "FindReplace", "Create", "SetRangeFormat", "ScanWorkbookStructure", "ManageTable", "ManageQuery", "ManageModelMeasure"],
            uncovered);
    }

    [Fact]
    public void ARunThatExercisedNothingReportsEveryOperation()
    {
        Assert.Equal(OperationCatalog.AllKinds.Count, FieldCheckCoverage.UncoveredKinds([]).Count);
    }
}
