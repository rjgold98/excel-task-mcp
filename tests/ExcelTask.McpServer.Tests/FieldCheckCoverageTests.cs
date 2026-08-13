using ExcelTask.Core;
using ExcelTask.McpServer;

namespace ExcelTask.McpServer.Tests;

/// <summary>
/// The field check's own coverage arithmetic - the first test of any kind over the ~900 lines that
/// validate a release on the locked-down work computer.
///
/// It shipped covering five of twelve operations, exercising everything from the first four
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

        // Derived rather than restated. The expectation used to be a written-out list, which meant
        // every new operation broke this test for the one reason that is not a defect - and the
        // temptation each time is to paste the new name in, which is how a list stops being a check.
        // The property is what matters: every kind the run did not exercise is named, in declaration
        // order, so a reader gets names rather than a ratio.
        var ran = shipped.Select(result => result.Label.Split(' ')[0]).ToHashSet(StringComparer.Ordinal);
        Assert.Equal(
            OperationCatalog.AllKinds.Select(kind => kind.ToString()).Where(name => !ran.Contains(name)),
            uncovered);
    }

    [Fact]
    public void ARunThatExercisedNothingReportsEveryOperation()
    {
        Assert.Equal(OperationCatalog.AllKinds.Count, FieldCheckCoverage.UncoveredKinds([]).Count);
    }
}
