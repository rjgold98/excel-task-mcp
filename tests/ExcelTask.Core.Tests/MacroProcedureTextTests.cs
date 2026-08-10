using ExcelTask.Core;

namespace ExcelTask.Core.Tests;

public sealed class MacroProcedureTextTests
{
    [Fact]
    public void NormalizeAndHashTreatLineEndingsAndTrailingNewlinesAsEquivalent()
    {
        const string lf = "Public Function RefreshModel() As Long\n    RefreshModel = 1\nEnd Function";
        var crlfWithTrailingNewlines = lf.Replace("\n", "\r\n", StringComparison.Ordinal) + "\r\n\r\n";

        Assert.Equal(lf, MacroProcedureText.NormalizeLineEndings(crlfWithTrailingNewlines));
        Assert.Equal(MacroProcedureText.ComputeSha256(lf), MacroProcedureText.ComputeSha256(crlfWithTrailingNewlines));
    }

    [Fact]
    public void NormalizesOneCompleteRequestedProcedureAndPreservesInteriorWhitespace()
    {
        const string source = "' replace model\r\nPublic Function RefreshModel() As Long\r\n    RefreshModel = 1\r\nEnd Function\r\n";

        var valid = MacroProcedureText.TryNormalizeProcedureSource(source, "refreshmodel", requireZeroParameters: true, out var normalized, out var error);

        Assert.True(valid, error);
        Assert.Equal("' replace model\nPublic Function RefreshModel() As Long\n    RefreshModel = 1\nEnd Function", normalized);
    }

    [Theory]
    [InlineData("Option Explicit\nSub RefreshModel()\nEnd Sub")]
    [InlineData("Sub RefreshModel()\nEnd Sub\nSub AnotherProcedure()\nEnd Sub")]
    [InlineData("Sub AnotherProcedure()\nEnd Sub")]
    [InlineData("Sub RefreshModel()\nEnd Function")]
    [InlineData("Sub RefreshModel()")]
    public void RejectsModuleDeclarationsSecondProceduresAndIncompleteSourceWithoutEchoingVba(string source)
    {
        var valid = MacroProcedureText.TryNormalizeProcedureSource(source, "RefreshModel", requireZeroParameters: false, out var normalized, out var error);

        Assert.False(valid);
        Assert.Null(normalized);
        Assert.DoesNotContain(source, error, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("Public Sub RefreshModel(ByRef values() As Variant)\n    values(0) = 1\nEnd Sub")]
    [InlineData("Public Function RefreshModel(rows() As Double) As Double\n    RefreshModel = 0\nEnd Function")]
    public void AcceptsArrayParametersWhichAreOrdinaryInAnalystMacros(string source)
    {
        var valid = MacroProcedureText.TryNormalizeProcedureSource(source, "RefreshModel", requireZeroParameters: false, out var normalized, out var error);

        Assert.True(valid, error);
        Assert.Equal(source, normalized);
    }

    [Fact]
    public void ArrayParameterProcedureStillCountsAsHavingParameters()
    {
        var valid = MacroProcedureText.TryNormalizeProcedureSource(
            "Public Sub RefreshModel(values() As Variant)\nEnd Sub",
            "RefreshModel",
            requireZeroParameters: true,
            out _,
            out var error);

        Assert.False(valid);
        Assert.NotNull(error);
    }

    [Theory]
    [InlineData("Public Sub RefreshModel()\n    MsgBox \"done\"\nEnd Sub", "MsgBox")]
    [InlineData("Public Sub RefreshModel()\n    Dim answer\n    answer = InputBox(\"name?\")\nEnd Sub", "InputBox")]
    [InlineData("Public Sub RefreshModel()\n    Stop\nEnd Sub", "Stop")]
    [InlineData("Public Sub RefreshModel()\n    If Range(\"A1\").Value2 = 0 Then Stop\nEnd Sub", "Stop")]
    [InlineData("Public Sub RefreshModel()\n    Range(\"A1\").Value2 = 1: Stop\nEnd Sub", "Stop")]
    [InlineData("Public Sub RefreshModel()\n    Dim f\n    f = Application.GetOpenFilename()\nEnd Sub", "GetOpenFilename")]
    public void DetectsConstructsThatWaitForAPerson(string source, string expected)
    {
        Assert.True(MacroProcedureText.TryFindBlockingConstruct(source, out var construct));
        Assert.Equal(expected, construct, ignoreCase: true);
    }

    [Theory]
    [InlineData("Public Sub RefreshModel()\n    Range(\"A1\").Value2 = \"MsgBox is only text here\"\nEnd Sub")]
    [InlineData("Public Sub RefreshModel()\n    ' MsgBox in a comment is not code\n    Range(\"A1\").Value2 = 1\nEnd Sub")]
    [InlineData("Public Sub RefreshModel()\n    Debug.Print \"no dialog\"\nEnd Sub")]
    [InlineData("Public Sub StopwatchTotal()\n    Range(\"A1\").Value2 = 1\nEnd Sub")]
    // Stop is an ordinary method name as well as a statement, and refusing these would refuse
    // macros that never wait for anyone.
    [InlineData("Public Sub RefreshModel()\n    timer.Stop\nEnd Sub")]
    [InlineData("Public Sub RefreshModel()\n    Application.Speech.Stop\nEnd Sub")]
    [InlineData("Public Sub RefreshModel()\n    Call recorder.Stop()\nEnd Sub")]
    public void DoesNotMistakeStringsCommentsOrSimilarNamesForBlockingCode(string source)
    {
        Assert.False(MacroProcedureText.TryFindBlockingConstruct(source, out var construct));
        Assert.Null(construct);
    }

    [Fact]
    public void RunValidationRejectsParametersAndSourceBounds()
    {
        var parameters = MacroProcedureText.TryNormalizeProcedureSource(
            "Sub RefreshModel(value As Long)\nEnd Sub",
            "RefreshModel",
            requireZeroParameters: true,
            out _,
            out var parameterError);
        var tooManyLines = string.Join("\n", Enumerable.Repeat("' comment", MacroProcedureText.MaxLineCount + 1));
        var lineLimit = MacroProcedureText.TryNormalizeProcedureSource(tooManyLines, "RefreshModel", false, out _, out var lineError);
        var oversized = new string('x', MacroProcedureText.MaxSourceCharacters + 1);
        var characterLimit = MacroProcedureText.TryNormalizeProcedureSource(oversized, "RefreshModel", false, out _, out var characterError);

        Assert.False(parameters);
        Assert.NotNull(parameterError);
        Assert.False(lineLimit);
        Assert.NotNull(lineError);
        Assert.False(characterLimit);
        Assert.NotNull(characterError);
    }
}
