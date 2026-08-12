using ExcelTask.Excel;
using ExcelTask.McpServer;

namespace ExcelTask.McpServer.Tests;

/// <summary>
/// The two rules that decide what a field report says about the machine it ran on, asserted
/// directly, because a real run cannot reach either of them on a healthy computer.
///
/// The leak reconciliation only does anything when an Excel process outlives its operation's own
/// twenty-second wait and then exits before the run ends. On the personal machine every operation
/// reported zero and the correcting branch never executed - so the fix for a defect measured on the
/// work computer would have shipped with no evidence it works, which is the same shape as the two
/// tests the review caught this morning that could not fail. A table of process identities reaches
/// it without needing a slow machine.
///
/// The profile redaction has the same problem in reverse: it is trivially exercised by every run
/// and trivially wrong in the cases that matter - a null value, a differently-cased drive letter, a
/// path that is not under the profile at all.
/// </summary>
public sealed class FieldCheckReportingTests
{
    private static ProcessIdentity Excel(int processId) =>
        new(processId, new DateTime(2026, 8, 12, 13, 0, 0, DateTimeKind.Utc), @"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE");

    private static OperationResult Operation(string label, params ProcessIdentity[] stillUpAtDeadline) =>
        new(label, "Completed", 1.0, stillUpAtDeadline.Length, "summary", "checks", null, stillUpAtDeadline);

    [Fact]
    public void AProcessThatHasSinceExitedIsNotCountedAsALeak()
    {
        // What the work computer actually produced: three operations whose Excel was still tearing
        // down at their own deadline, and a run that ended with nothing left behind. The report
        // said "Leaked Excel: 2" and "excelLeakedByProduct: 0" in the same file.
        var operations = new List<OperationResult>
        {
            Operation("WriteWorksheetValues (Apply)", Excel(7364), Excel(34020)),
            Operation("SetRangeFormat (Apply)", Excel(29912), Excel(34404)),
        };
        var notes = new List<string>();

        FieldCheck.ReconcilePerOperationLeaks(operations, [], notes);

        Assert.All(operations, operation => Assert.Equal(0, operation.LeakedExcel));
        // The correction is stated rather than performed silently: a reader comparing this run to
        // the console output needs to know why the two disagree.
        Assert.Contains(notes, note => note.Contains("2 operation(s)", StringComparison.Ordinal));
    }

    [Fact]
    public void AProcessStillRunningAtTheEndIsStillALeak()
    {
        // The whole point of the reconciliation is that it cannot launder a real leak. 34020 never
        // exited, so it survives; 7364 did, so it does not.
        var operations = new List<OperationResult> { Operation("WriteWorksheetValues (Apply)", Excel(7364), Excel(34020)) };
        var notes = new List<string>();

        FieldCheck.ReconcilePerOperationLeaks(operations, [Excel(34020)], notes);

        Assert.Equal(1, operations[0].LeakedExcel);
    }

    [Fact]
    public void AnOperationThatLeftNothingBehindIsUntouchedAndUnremarkedOn()
    {
        var operations = new List<OperationResult> { Operation("ScanWorkbookStructure (Plan)") };
        var notes = new List<string>();

        FieldCheck.ReconcilePerOperationLeaks(operations, [Excel(34020)], notes);

        Assert.Equal(0, operations[0].LeakedExcel);
        // A note on a run where nothing was corrected would train the reader to skip the notes.
        Assert.Empty(notes);
    }

    [Fact]
    public void ARunWhereEveryFigureWasAlreadyRightAddsNoNote()
    {
        // Identity, not just zero: the count already matched, so there is nothing to say.
        var operations = new List<OperationResult> { Operation("WriteWorksheetValues (Apply)", Excel(34020)) };
        var notes = new List<string>();

        FieldCheck.ReconcilePerOperationLeaks(operations, [Excel(34020)], notes);

        Assert.Equal(1, operations[0].LeakedExcel);
        Assert.Empty(notes);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public void AnAbsentValueReadsAsNotSetRatherThanEmpty(string? value) =>
        Assert.Equal("not set", FieldCheck.WithoutUserProfile(value));

    [Fact]
    public void TheAccountNameIsRemovedFromAPathBeneathTheProfile()
    {
        var profile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        Assert.False(string.IsNullOrEmpty(profile), "this test needs a real user profile to redact");

        var redacted = FieldCheck.WithoutUserProfile(Path.Combine(profile, "ExcelTask", "excel-task-mcp.exe"));

        Assert.StartsWith("%USERPROFILE%", redacted, StringComparison.Ordinal);
        // The account name is the whole point. Paired with the computer name this report already
        // carries, leaving it in names a person and not only a machine.
        Assert.DoesNotContain(Path.GetFileName(profile), redacted, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void TheComparisonIgnoresCaseBecauseWindowsPathsDo()
    {
        var profile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);

        // DOTNET_ROOT and an MCP client's configured command are typed by people and by installers,
        // and they disagree about drive-letter and folder casing. A case-sensitive compare would
        // pass every test written from Environment's own output and redact nothing in the field.
        var redacted = FieldCheck.WithoutUserProfile(profile.ToUpperInvariant() + @"\.dotnet");

        Assert.Equal(@"%USERPROFILE%\.dotnet", redacted);
    }

    [Fact]
    public void APathOutsideTheProfileIsLeftExactlyAsItWas()
    {
        const string shared = @"C:\Program Files\dotnet";
        Assert.Equal(shared, FieldCheck.WithoutUserProfile(shared));
    }
}
