using System.Globalization;
using System.Runtime.Versioning;

namespace ExcelTask.McpServer;

/// <summary>
/// Writes a small set of disposable workbooks to try ExcelTask against.
///
/// This exists because every operation needs an existing workbook, and a managed computer usually
/// cannot produce one: Constrained Language Mode blocks the COM automation a script would need, so
/// without this the first real test depends on somebody building a workbook by hand. Like
/// --field-check and --excel-worker, this is an internal mode rather than an MCP tool.
/// </summary>
[SupportedOSPlatform("windows")]
internal static class MakeFixture
{
    private const string MacroComponent = "DemoModule";
    private const string MacroSource = "Public Sub StampRun()\n    ThisWorkbook.Worksheets(1).Range(\"A1\").Value2 = \"before\"\nEnd Sub";

    public static int Run(string[] args)
    {
        // Defaulting rather than refusing: the first field use of this spent a cycle on a usage
        // error, and there is an obvious right answer for where disposable workbooks should go.
        var directory = Path.GetFullPath(args.Length >= 2 && !string.IsNullOrWhiteSpace(args[1])
            ? args[1]
            : Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "ExcelTaskDemo"));
        Directory.CreateDirectory(directory);
        var target = Path.Combine(directory, "target.xlsx");
        var reference = Path.Combine(directory, "reference.xlsx");
        var macroTarget = Path.Combine(directory, "macro-target.xlsm");

        var fixtures = new FieldCheckFixtures();
        try
        {
            fixtures.CreateFormulaFixtures(target, reference);
            var macroNote = fixtures.TryCreateMacroFixture(macroTarget, MacroComponent, MacroSource);

            Console.WriteLine("Created disposable workbooks:");
            Console.WriteLine($"  {target}");
            Console.WriteLine("      sheet 'Model'  A1:D1 = 1,2,3,4   A2 and B2 hold =A1*2 and =B1*2");
            Console.WriteLine($"  {reference}");
            Console.WriteLine("      sheet 'Reference'  A1 and A3 hold =ROW(), A2 is deliberately blank");
            Console.WriteLine(macroNote is null
                ? $"  {macroTarget}{Environment.NewLine}      module '{MacroComponent}' with procedure 'StampRun'"
                : $"  (no macro workbook) {macroNote}");

            Console.WriteLine();
            Console.WriteLine("Try one of these through the excel-task server:");
            Console.WriteLine($"  Copy the Reference worksheet from {reference}");
            Console.WriteLine($"  into {target} as a new sheet named Exhibit A,");
            Console.WriteLine("  and repair the blank formula in A1:A3.");
            Console.WriteLine();
            Console.WriteLine($"  On sheet Model in {target}, extend the formula");
            Console.WriteLine("  pattern in A2:B2 two columns right into C2:D2.");
            Console.WriteLine();
            Console.WriteLine("Delete this folder when you are finished; nothing here is needed afterwards.");
            return 0;
        }
        catch (Exception exception)
        {
            Console.Error.WriteLine(string.Create(CultureInfo.InvariantCulture,
                $"Could not create the fixtures: {exception.GetType().Name}: {exception.Message.ReplaceLineEndings(" ").Trim()}"));
            return 1;
        }
        finally
        {
            fixtures.TerminateOwnedProcesses();
        }
    }
}
