using System.IO.Compression;
using System.Globalization;
using System.Text;
using ExcelTask.Core;

namespace ExcelTask.Excel.Tests;

/// <summary>
/// The scan parser pinned against hand-authored Office Open XML - which is what makes these tests
/// runnable in the fast tier: the operation never starts Excel, so its tests need none either.
/// Hand-authored deliberately: the XML here states exactly which OOXML shapes the parser claims to
/// understand, including the shared-formula member cells whose f element is empty, where a naive
/// count would misread an Excel-written file.
/// </summary>
public sealed class WorkbookStructureScanTests
{
    [Fact]
    public async Task ScansSheetsFormulasConstantsAndOverridesWithoutExcel()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var target = Path.Combine(directory, "scan.xlsx");

        try
        {
            // Ledger!D: 24 formula cells (2 explicit, 22 shared-formula members with empty f
            // elements) and 3 constants - the override signature... except the island threshold is
            // 20 formulas with constants under a tenth, so 24:3 stays BELOW the island bar and only
            // the counts assert. Notes!B is 30 formulas and 2 constants: a real island.
            WriteScanFixture(target);

            using var runtime = new ExcelWorkbookRuntime();
            var outcome = await runtime.ExecuteAsync(
                new ExcelTaskPlan("scan", ExcelTaskPlans.Scan(target)), CancellationToken.None);

            Assert.True(outcome.Status == ExcelTaskStatus.Planned,
                $"{outcome.Summary} {string.Join("; ", outcome.Checks?.Select(check => check.Detail) ?? [])}");
            Assert.Contains(outcome.Checks ?? [], check => check.Name == "structure-scan" && check.Passed);
            Assert.Contains(outcome.Checks ?? [], check => check.Name == "workbook-unchanged" && check.Passed);

            var items = outcome.Audit!.Items;
            var ledger = Assert.Single(items, item => item.Kind == "scan-sheet" && item.Name == "Ledger");
            Assert.Contains("dimension A1:D30", ledger.Detail, StringComparison.Ordinal);
            Assert.Contains("24 formula cell(s)", ledger.Detail, StringComparison.Ordinal);
            Assert.Contains("9 constant cell(s)", ledger.Detail, StringComparison.Ordinal);

            var notes = Assert.Single(items, item => item.Kind == "scan-sheet" && item.Name == "Notes");
            Assert.Contains("30 formula cell(s)", notes.Detail, StringComparison.Ordinal);

            // The reason the operation exists: the override inside the calculated column, by name.
            var island = Assert.Single(items, item => item.Kind == "constant-island");
            Assert.Equal("Notes!B", island.Name);
            Assert.Contains("B7", island.Detail, StringComparison.Ordinal);
            Assert.Contains("B19", island.Detail, StringComparison.Ordinal);
            Assert.True(outcome.Audit.WorkbookUnchanged);
        }
        finally
        {
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public async Task AnEncryptedOrDamagedFileIsACleanRejectionThatNamesTheRoadStillOpen()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ExcelTask", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var target = Path.Combine(directory, "encrypted.xlsx");

        try
        {
            // An encrypted workbook is an OLE compound file: it starts with D0 CF 11 E0, not a ZIP.
            File.WriteAllBytes(target, [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 0, 0, 0, 0]);

            using var runtime = new ExcelWorkbookRuntime();
            var outcome = await runtime.ExecuteAsync(
                new ExcelTaskPlan("scan-bad", ExcelTaskPlans.Scan(target)), CancellationToken.None);

            Assert.Equal(ExcelTaskStatus.Rejected, outcome.Status);
            var check = Assert.Single(outcome.Checks!, item => item.Name == "structure-scan");
            Assert.False(check.Passed);
            Assert.Contains("Excel-based operations can still open it", check.Detail, StringComparison.Ordinal);
        }
        finally
        {
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        }
    }

    /// <summary>A minimal but genuine OOXML package: content types, rels, workbook, two sheets.</summary>
    private static void WriteScanFixture(string path)
    {
        using var archive = ZipFile.Open(path, ZipArchiveMode.Create);

        Add(archive, "[Content_Types].xml", """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
              <Default Extension="xml" ContentType="application/xml"/>
              <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
              <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
              <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
            </Types>
            """);

        Add(archive, "_rels/.rels", """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
            </Relationships>
            """);

        Add(archive, "xl/workbook.xml", """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <sheets>
                <sheet name="Ledger" sheetId="1" r:id="rId1"/>
                <sheet name="Notes" sheetId="2" r:id="rId2"/>
              </sheets>
            </workbook>
            """);

        Add(archive, "xl/_rels/workbook.xml.rels", """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
              <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
            </Relationships>
            """);

        // Ledger: rows 1-6 constants in A (6 constants) plus B constants rows 1-3 (3), and column D
        // rows 1-24: one shared-formula master, 22 members with EMPTY f elements - the Excel-written
        // shape a naive scan misses - plus one explicit formula. Constants: 9. Formulas: 24.
        var ledger = new StringBuilder();
        ledger.Append("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:D30"/><sheetData>""");
        for (var row = 1; row <= 6; row++)
        {
            ledger.Append(CultureInfo.InvariantCulture, $"""<row r="{row}"><c r="A{row}"><v>{row * 10}</v></c>""");
            if (row <= 3) ledger.Append(CultureInfo.InvariantCulture, $"""<c r="B{row}" t="str"><v>label</v></c>""");
            if (row == 1) ledger.Append(CultureInfo.InvariantCulture, $"""<c r="D1"><f t="shared" ref="D1:D23" si="0">A1*2</f><v>20</v></c>""");
            else if (row <= 6) ledger.Append(CultureInfo.InvariantCulture, $"""<c r="D{row}"><f t="shared" si="0"/><v>0</v></c>""");
            ledger.Append("</row>");
        }
        for (var row = 7; row <= 23; row++)
        {
            ledger.Append(CultureInfo.InvariantCulture, $"""<row r="{row}"><c r="D{row}"><f t="shared" si="0"/><v>0</v></c></row>""");
        }
        ledger.Append("""<row r="24"><c r="D24"><f>SUM(D1:D23)</f><v>0</v></c></row>""");
        ledger.Append("</sheetData></worksheet>");
        Add(archive, "xl/worksheets/sheet1.xml", ledger.ToString());

        // Notes: B1:B32 - 30 formulas with constants punched in at B7 and B19. 30 formulas, 2
        // constants: 2*10 < 30, so this IS a constant island.
        var notes = new StringBuilder();
        notes.Append("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="B1:B32"/><sheetData>""");
        for (var row = 1; row <= 32; row++)
        {
            if (row is 7 or 19)
            {
                notes.Append(CultureInfo.InvariantCulture, $"""<row r="{row}"><c r="B{row}"><v>999</v></c></row>""");
            }
            else
            {
                notes.Append(CultureInfo.InvariantCulture, $"""<row r="{row}"><c r="B{row}"><f>A{row}+1</f><v>0</v></c></row>""");
            }
        }
        notes.Append("</sheetData></worksheet>");
        Add(archive, "xl/worksheets/sheet2.xml", notes.ToString());
    }

    private static void Add(ZipArchive archive, string name, string content)
    {
        var entry = archive.CreateEntry(name);
        using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
        writer.Write(content);
    }
}
