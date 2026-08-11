using System.Globalization;
using System.IO.Compression;
using System.Xml;
using ExcelTask.Core;

namespace ExcelTask.Excel;

public sealed partial class ExcelWorkbookRuntime
{
    /// <summary>Cells examined before a scan stops and says so - a zip-bomb guard, not a real limit.</summary>
    private const long MaxScannedCells = 5_000_000;

    /// <summary>Constant addresses remembered per column; counting continues past it.</summary>
    private const int MaxIslandAddressesPerColumn = 64;

    /// <summary>
    /// Maps a workbook's structure by reading the file as what it physically is - a ZIP of XML -
    /// with no Excel process at any point.
    ///
    /// This is the only operation in the server that does not automate Excel, and that is its whole
    /// value: the trace measured 92% of a small task's wall time as owned-Excel teardown and
    /// verification, none of which a scan pays. It still runs inside the supervised worker
    /// deliberately - a malformed or hostile file takes down a disposable subprocess with a
    /// deadline, never the server.
    ///
    /// It reports shape, never contents: sheet names, dimensions, counts, and cell addresses - no
    /// cell values and no formula text. The mixed-column report (a column that is mostly formulas
    /// holding a scattering of constants) is the signature of a manual override sitting inside a
    /// calculated column, which is what a caller needs to see before deciding what to read or fix.
    /// </summary>
    private static WorkbookExecutionOutcome ExecuteScanCore(ExcelTaskPlan plan, IExcelWorkbookRuntimeObserver observer)
    {
        var checks = new List<TaskCheck>();
        string targetPath;
        try
        {
            targetPath = WorkbookRuntimeHelpers.NormalizePath(plan.Request.TargetWorkbookPath);
            WorkbookRuntimeHelpers.EnsureReadableWorkbook(targetPath, "Target workbook");
        }
        catch (InvalidOperationException exception)
        {
            return new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected, "The target workbook cannot be read.",
                Checks: [new TaskCheck("workbook-inputs", false, exception.Message)]);
        }

        var stampBefore = ReadFileStamp(targetPath);
        observer.OnPhase("structure-scan");

        List<WorkbookFlowItem> items;
        var truncated = false;
        try
        {
            items = ScanArchive(targetPath, ref truncated);
        }
        catch (Exception exception) when (exception is InvalidDataException or XmlException or IOException or UnauthorizedAccessException)
        {
            // An encrypted workbook is an OLE container rather than a ZIP, which is the common way
            // this fails; the answer names the road still open rather than dead-ending.
            checks.Add(new TaskCheck("structure-scan", false,
                "The file could not be read as an Open XML workbook - it may be encrypted or damaged. The Excel-based operations can still open it."));
            return new WorkbookExecutionOutcome(ExcelTaskStatus.Rejected,
                "The workbook could not be scanned without Excel.", Checks: checks);
        }

        var totalFound = items.Count;
        if (items.Count > MaxAuditReceiptItems)
        {
            items = [.. items.Take(MaxAuditReceiptItems)];
            truncated = true;
        }

        checks.Add(new TaskCheck("structure-scan", true,
            $"Read the file directly and mapped {totalFound} structural item(s); no Excel process was started."));

        // The same proof the audit and the read give, from evidence rather than intent.
        var unchanged = ReadFileStamp(targetPath) == stampBefore;
        checks.Add(new TaskCheck("workbook-unchanged", unchanged, unchanged
            ? "The workbook's size and timestamp are identical before and after the scan."
            : "The workbook's size or timestamp changed during the scan; something else wrote to it."));

        return new WorkbookExecutionOutcome(
            plan.Request.Mode == ExcelTaskMode.Plan ? ExcelTaskStatus.Planned : ExcelTaskStatus.Completed,
            $"Scanned the workbook's structure directly - {totalFound} item(s), no Excel process started.",
            Checks: checks,
            Audit: new WorkbookAuditReceipt(items, totalFound, truncated, unchanged));
    }

    /// <summary>
    /// One pass over the archive: sheet list from xl/workbook.xml, each sheet's part resolved
    /// through the relationships file rather than assumed from its position, then a streaming read
    /// of each part. Streaming matters - a large sheet's XML is tens of megabytes, and a DOM parse
    /// of it would be the scan becoming the memory problem it exists to avoid.
    /// </summary>
    private static List<WorkbookFlowItem> ScanArchive(string path, ref bool truncated)
    {
        using var archive = ZipFile.OpenRead(path);
        var workbookEntry = archive.GetEntry("xl/workbook.xml")
            ?? throw new InvalidDataException("The archive holds no workbook part.");

        var relationshipTargets = ReadRelationshipTargets(archive);
        var sheets = ReadSheetList(workbookEntry);

        var items = new List<WorkbookFlowItem>();
        var budget = MaxScannedCells;
        foreach (var (sheetName, relationshipId) in sheets)
        {
            if (!relationshipTargets.TryGetValue(relationshipId, out var target) || !target.StartsWith("worksheets/", StringComparison.OrdinalIgnoreCase))
            {
                // A chartsheet or something else that is not a grid; named so the count of sheets
                // still matches what a person sees in Excel's tab bar.
                items.Add(new WorkbookFlowItem("scan-sheet", sheetName, "not a worksheet grid"));
                continue;
            }

            var entry = archive.GetEntry($"xl/{target}");
            if (entry is null)
            {
                items.Add(new WorkbookFlowItem("scan-sheet", sheetName, "sheet part missing from the archive"));
                continue;
            }

            var sheet = ScanSheet(entry, ref budget);
            items.Add(new WorkbookFlowItem(
                "scan-sheet",
                sheetName,
                $"dimension {sheet.Dimension}; {sheet.FormulaCells:N0} formula cell(s); {sheet.ConstantCells:N0} constant cell(s)" +
                (sheet.Truncated ? "; scan budget reached, counts partial" : "")));
            if (sheet.Truncated) truncated = true;

            foreach (var island in sheet.ConstantIslands)
            {
                items.Add(new WorkbookFlowItem(
                    "constant-island",
                    $"{sheetName}!{island.Column}",
                    $"{island.Addresses.Count}{(island.CountIsExact ? "" : "+")} constant(s) inside a {island.FormulaCount:N0}-formula column: " +
                    string.Join(", ", island.Addresses.Take(8)) + (island.Addresses.Count > 8 ? ", ..." : "")));
            }
        }

        items.AddRange(ScanDefinedNames(workbookEntry));
        items.AddRange(ScanTables(archive));
        items.AddRange(ScanExternalLinks(archive));
        return items;
    }

    /// <summary>
    /// The three categories beyond sheets that the package answers in plain XML.
    ///
    /// Deliberately only these three. A part-by-part inventory of a workbook carrying every
    /// feature found that macro components live in a binary OLE compound file, the data model in
    /// an opaque blob, and Power Query inside base64 of a nested ZIP under an undocumented
    /// DataMashup element. Those stay with AuditWorkbookFlows, which opens Excel and asks. Guessing
    /// at them here would put a wrong answer about VBA or a connection into a receipt that reads
    /// as complete - the one outcome worth more than the seconds saved.
    /// </summary>
    private static List<WorkbookFlowItem> ScanDefinedNames(ZipArchiveEntry workbookEntry)
    {
        var items = new List<WorkbookFlowItem>();
        using var reader = XmlReader.Create(workbookEntry.Open(), ScanReaderSettings());
        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "definedName") continue;
            var name = reader.GetAttribute("name") ?? "(unnamed)";
            // A defined name's value is a reference, not workbook content, and it is the whole
            // point of reporting one - a caller planning a fix needs to know where it points.
            var target = reader.ReadElementContentAsString();
            items.Add(new WorkbookFlowItem("scan-defined-name", name, target.Length == 0 ? "(empty reference)" : target));
        }

        return items;
    }

    private static List<WorkbookFlowItem> ScanTables(ZipArchive archive)
    {
        var items = new List<WorkbookFlowItem>();
        foreach (var entry in archive.Entries.Where(entry =>
                     entry.FullName.StartsWith("xl/tables/", StringComparison.Ordinal) &&
                     entry.FullName.EndsWith(".xml", StringComparison.Ordinal)))
        {
            using var reader = XmlReader.Create(entry.Open(), ScanReaderSettings());
            while (reader.Read())
            {
                if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "table") continue;
                var name = reader.GetAttribute("displayName") ?? reader.GetAttribute("name") ?? "(unnamed)";
                items.Add(new WorkbookFlowItem("scan-table", name, $"range {reader.GetAttribute("ref") ?? "(unknown)"}"));
                break;
            }
        }

        return items;
    }

    private static List<WorkbookFlowItem> ScanExternalLinks(ZipArchive archive)
    {
        var items = new List<WorkbookFlowItem>();
        foreach (var entry in archive.Entries.Where(entry =>
                     entry.FullName.StartsWith("xl/externalLinks/_rels/", StringComparison.Ordinal) &&
                     entry.FullName.EndsWith(".rels", StringComparison.Ordinal)))
        {
            using var reader = XmlReader.Create(entry.Open(), ScanReaderSettings());
            while (reader.Read())
            {
                if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "Relationship") continue;
                var target = reader.GetAttribute("Target");
                if (string.IsNullOrEmpty(target)) continue;
                // The file name only. The full target is a path on someone's machine or a share,
                // and the audit holds the same line for the same reason.
                items.Add(new WorkbookFlowItem("scan-external-link", DiagnosticTrace.FileNameOnly(target.Replace('/', '\\')), "linked workbook"));
            }
        }

        return items;
    }

    private static Dictionary<string, string> ReadRelationshipTargets(ZipArchive archive)
    {
        var targets = new Dictionary<string, string>(StringComparer.Ordinal);
        var entry = archive.GetEntry("xl/_rels/workbook.xml.rels");
        if (entry is null) return targets;

        using var reader = XmlReader.Create(entry.Open(), ScanReaderSettings());
        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "Relationship")
            {
                var id = reader.GetAttribute("Id");
                var target = reader.GetAttribute("Target");
                if (id is not null && target is not null) targets[id] = target.TrimStart('/');
            }
        }

        return targets;
    }

    private static List<(string Name, string RelationshipId)> ReadSheetList(ZipArchiveEntry workbookEntry)
    {
        var sheets = new List<(string, string)>();
        using var reader = XmlReader.Create(workbookEntry.Open(), ScanReaderSettings());
        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "sheet")
            {
                var name = reader.GetAttribute("name") ?? "(unnamed)";
                var relationshipId = reader.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships") ?? string.Empty;
                sheets.Add((name, relationshipId));
            }
        }

        return sheets;
    }

    private sealed record ConstantIsland(string Column, int FormulaCount, IReadOnlyList<string> Addresses, bool CountIsExact);

    private sealed record ScannedSheet(string Dimension, long FormulaCells, long ConstantCells, bool Truncated, IReadOnlyList<ConstantIsland> ConstantIslands);

    private static ScannedSheet ScanSheet(ZipArchiveEntry entry, ref long budget)
    {
        var dimension = "(unknown)";
        long formulaCells = 0, constantCells = 0;
        var truncatedHere = false;
        var columnFormulas = new Dictionary<string, int>(StringComparer.Ordinal);
        var columnConstantAddresses = new Dictionary<string, List<string>>(StringComparer.Ordinal);
        var columnConstantCounts = new Dictionary<string, int>(StringComparer.Ordinal);

        using (var reader = XmlReader.Create(entry.Open(), ScanReaderSettings()))
        {
            string currentRef = "";
            var currentHasFormula = false;
            var currentHasValue = false;

            void CloseCell()
            {
                if (currentRef.Length == 0 || (!currentHasFormula && !currentHasValue)) return;
                var column = ColumnOf(currentRef);
                if (currentHasFormula)
                {
                    formulaCells++;
                    columnFormulas[column] = columnFormulas.GetValueOrDefault(column) + 1;
                }
                else
                {
                    constantCells++;
                    columnConstantCounts[column] = columnConstantCounts.GetValueOrDefault(column) + 1;
                    var addresses = columnConstantAddresses.TryGetValue(column, out var list) ? list : columnConstantAddresses[column] = [];
                    if (addresses.Count < MaxIslandAddressesPerColumn) addresses.Add(currentRef);
                }
            }

            while (reader.Read())
            {
                if (reader.NodeType != XmlNodeType.Element) continue;
                switch (reader.LocalName)
                {
                    case "dimension":
                        dimension = reader.GetAttribute("ref") ?? dimension;
                        break;
                    case "c":
                        CloseCell();
                        currentRef = reader.GetAttribute("r") ?? "";
                        currentHasFormula = false;
                        currentHasValue = false;
                        if (--budget <= 0)
                        {
                            truncatedHere = true;
                            goto done;
                        }

                        break;
                    case "f":
                        currentHasFormula = true;
                        break;
                    case "v":
                    case "is":
                        currentHasValue = true;
                        break;
                }
            }

            done:
            CloseCell();
        }

        // The override signature: a column dominated by one formula pattern holding a scattering of
        // constants. Thresholds are deliberately conservative - a 50/50 column is a data column, not
        // an override, and flagging it would teach callers to ignore the flag.
        var islands = new List<ConstantIsland>();
        foreach (var (column, formulaCount) in columnFormulas)
        {
            var constantCount = columnConstantCounts.GetValueOrDefault(column);
            if (formulaCount >= 20 && constantCount > 0 && constantCount * 10 < formulaCount)
            {
                var addresses = columnConstantAddresses[column];
                islands.Add(new ConstantIsland(column, formulaCount, addresses, constantCount == addresses.Count));
            }
        }

        return new ScannedSheet(dimension, formulaCells, constantCells, truncatedHere, islands);
    }

    private static string ColumnOf(string cellReference)
    {
        var length = 0;
        while (length < cellReference.Length && char.IsAsciiLetter(cellReference[length])) length++;
        return cellReference[..length];
    }

    private static XmlReaderSettings ScanReaderSettings() => new()
    {
        DtdProcessing = DtdProcessing.Prohibit,
        XmlResolver = null,
        IgnoreWhitespace = true,
        IgnoreComments = true
    };
}
