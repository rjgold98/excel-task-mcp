using System.Collections.Concurrent;
using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.Versioning;
using System.Security.Cryptography;
using System.Text;
using ExcelTask.Core;

namespace ExcelTask.Excel;
internal static class WorkbookRuntimeHelpers
{
    public const int AutomationSecurityLow = 1;
    public const int AutomationSecurityForceDisable = 3;

    private static readonly HashSet<string> SupportedWorkbookExtensions = new(StringComparer.OrdinalIgnoreCase) { ".xlsx", ".xlsm" };

    public static string NormalizePath(string path) => Path.GetFullPath(path ?? throw new ArgumentNullException(nameof(path)));

    public static bool PathsEqual(string left, string right) => string.Equals(
        NormalizePath(left).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar),
        NormalizePath(right).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar),
        StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Whether an identity Excel reported denotes the workbook the caller named.
    ///
    /// Usually that is exact path equality, and it stays exact. The exception is a workbook opened
    /// from a synced OneDrive or SharePoint folder: Excel reports its service URL rather than the
    /// local path, so a comparison against the caller's path found nothing and refused - correctly
    /// for the question asked, uselessly for the one that mattered. The URL is resolved back
    /// through the sync client's own mapping and then compared exactly, so nothing is guessed.
    ///
    /// Deliberately separate from <see cref="PathsEqual"/>, which compares two paths the caller
    /// supplied. Only an identity Excel reported can be a URL.
    /// </summary>
    public static bool IdentifiesSameWorkbook(string? reportedIdentity, string targetPath)
    {
        if (string.IsNullOrWhiteSpace(reportedIdentity)) return false;

        if (reportedIdentity.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
            reportedIdentity.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
        {
            return OneDriveSyncMap.MatchesLocalPath(reportedIdentity, targetPath);
        }

        // A reported identity is not always a well-formed path - an unsaved workbook reports a bare
        // name - and normalising one must not take down the search that would have found the file.
        try { return PathsEqual(reportedIdentity, targetPath); }
        catch (Exception exception) when (exception is ArgumentException or NotSupportedException or PathTooLongException)
        {
            return false;
        }
    }

    /// <summary>
    /// A file name with no directory, because a full path names the machine and the person.
    ///
    /// It lives here rather than on the diagnostic tracer because it is a redaction rule the
    /// receipts depend on, and the tracer documents itself as temporary and built to be deleted.
    /// A privacy guarantee must not be reachable only through a module scheduled for removal.
    /// </summary>
    public static string FileNameOnly(string? path) => string.IsNullOrWhiteSpace(path)
        ? "(none)"
        : Path.GetFileName(path.Replace('/', Path.DirectorySeparatorChar));

    /// <summary>
    /// Whether a directory will accept a new file, answered by trying it.
    ///
    /// The attribute cannot answer this, which is what the previous version got wrong.
    /// FILE_ATTRIBUTE_READONLY on a *directory* is a shell marker for a customized folder, not a
    /// permission: on an ordinary Windows profile, Documents, Downloads, Desktop and the OneDrive
    /// root all carry it while being perfectly writable. Testing it refused every copy-save and
    /// every create into the folders people actually keep workbooks in - and still missed the real
    /// case, because a genuinely unwritable directory is ACL-denied and carries no attribute at all.
    /// False on the common case, blind on the true one.
    ///
    /// So this asks the question the save will ask, early, which is the entire point of a preflight.
    /// DeleteOnClose means a crash between create and delete leaves nothing behind.
    /// </summary>
    public static bool DirectoryAcceptsNewFile(string directory)
    {
        var probe = Path.Combine(directory, $".exceltask-probe-{Guid.NewGuid():N}.tmp");
        try
        {
            using var stream = new FileStream(
                probe, FileMode.CreateNew, FileAccess.Write, FileShare.None, 1, FileOptions.DeleteOnClose);
            return true;
        }
        catch (IOException)
        {
            return false;
        }
        catch (UnauthorizedAccessException)
        {
            return false;
        }
    }

    public static bool CanOpenExclusively(string path)
    {
        try
        {
            using var stream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            return true;
        }
        catch (IOException)
        {
            return false;
        }
        catch (UnauthorizedAccessException)
        {
            return false;
        }
    }

    public static void EnsureReadableWorkbook(string path, string description)
    {
        // The file name, never the directory: a caller correcting a typo needs to see which name was
        // checked, and a UX simulation had to cross-reference its own request to find out. The
        // directory would name the machine and the person, so it stays out.
        var name = Path.GetFileName(path);
        if (!SupportedWorkbookExtensions.Contains(Path.GetExtension(path)))
        {
            throw new InvalidOperationException($"{description} must be an .xlsx or .xlsm file; '{name}' is not.");
        }

        if (!File.Exists(path)) throw new InvalidOperationException($"{description} does not exist: no file named '{name}' at that path.");
    }

    /// <summary>
    /// Refuses a same-file save whose target cannot be written, before Excel is ever started.
    /// Without this a read-only target is discovered only when the save fails mid-operation,
    /// producing Unknown - the worst possible answer, since it tells the caller the file may or
    /// may not have changed. Measured against the original server in the failure-mode matrix:
    /// it refuses cleanly at open, and there was no reason ExcelTask should do worse.
    /// </summary>
    public static void EnsureWritableSameTarget(string path)
    {
        if ((File.GetAttributes(path) & FileAttributes.ReadOnly) != 0)
        {
            throw new InvalidOperationException("The target workbook is read-only and cannot be saved in place.");
        }
    }

    /// <summary>
    /// The mirror of <see cref="EnsureReadableWorkbook"/> for the one operation that wants the
    /// target absent. Refusing an existing file here rather than at save time is what keeps
    /// "create" from ever meaning "overwrite" - there is no confirmation that unlocks it, because a
    /// caller who wants to replace a workbook should say so with a save, not with a create.
    /// </summary>
    public static void EnsureCreatableWorkbook(string path)
    {
        if (!SupportedWorkbookExtensions.Contains(Path.GetExtension(path)))
        {
            throw new InvalidOperationException("A new workbook path must end in .xlsx or .xlsm.");
        }

        if (File.Exists(path)) throw new InvalidOperationException("A workbook already exists at that path; creating one never overwrites.");

        var parent = Directory.GetParent(path)?.FullName;
        if (parent is null || !Directory.Exists(parent)) throw new InvalidOperationException("The new workbook's directory does not exist.");
        if (!DirectoryAcceptsNewFile(parent))
        {
            throw new InvalidOperationException("The new workbook's directory will not accept a new file.");
        }
    }

    public static void EnsureWritableCopyOutput(string? outputPath)
    {
        if (string.IsNullOrWhiteSpace(outputPath)) throw new InvalidOperationException("Copy output path is required.");
        var normalized = NormalizePath(outputPath);
        if (!SupportedWorkbookExtensions.Contains(Path.GetExtension(normalized)))
        {
            throw new InvalidOperationException("Copy output must be an .xlsx or .xlsm file.");
        }

        var parent = Directory.GetParent(normalized)?.FullName;
        if (parent is null || !Directory.Exists(parent)) throw new InvalidOperationException("Copy output directory does not exist.");
        if (!DirectoryAcceptsNewFile(parent))
        {
            throw new InvalidOperationException("Copy output directory will not accept a new file.");
        }

        if (File.Exists(normalized) && !CanOpenExclusively(normalized))
        {
            throw new InvalidOperationException("Existing copy output is locked.");
        }
    }

    public static string CreateStagingPath(string finalPath, string taskId)
    {
        var normalized = NormalizePath(finalPath);
        var directory = Path.GetDirectoryName(normalized) ?? throw new InvalidOperationException("Copy output directory is required.");
        var extension = Path.GetExtension(normalized);
        var name = Path.GetFileNameWithoutExtension(normalized);
        if (!IsSafeTaskId(taskId))
        {
            throw new InvalidOperationException("Task identity cannot be used for staging.");
        }

        return Path.Combine(directory, $".{name}.excel-task-{taskId}-{Guid.NewGuid():N}{extension}");
    }

    public static bool IsSafeTaskId(string? taskId) => taskId is { Length: > 0 and <= 128 } &&
        taskId.All(character => char.IsAsciiLetterOrDigit(character) || character is '-' or '_');

    public static void PromoteStaging(string stagingPath, string finalPath, bool overwrite)
    {
        File.Move(stagingPath, finalPath, overwrite);
    }

    public static bool TryDeleteStaging(string stagingPath)
    {
        try
        {
            if (File.Exists(stagingPath)) File.Delete(stagingPath);
            return !File.Exists(stagingPath);
        }
        catch (IOException) { return false; }
        catch (UnauthorizedAccessException) { return false; }
    }

    public static FormulaRangeBounds GetBounds(FormulaRepairRange range)
    {
        ArgumentNullException.ThrowIfNull(range);
        var start = ParseCell(range.StartCell);
        var end = ParseCell(range.EndCell);
        if (start.Row > end.Row || start.Column > end.Column) throw new InvalidOperationException("Formula repair range is not rectangular.");
        return new FormulaRangeBounds(start.Row, start.Column, end.Row, end.Column);
    }

    public static FormulaGridCell[,] CreateFormulaGrid(object value, int rowCount, int columnCount)
    {
        var grid = new FormulaGridCell[rowCount, columnCount];
        for (var row = 0; row < rowCount; row++)
        {
            for (var column = 0; column < columnCount; column++)
            {
                var cellValue = value is Array values ? values.GetValue(row + values.GetLowerBound(0), column + values.GetLowerBound(1)) : value;
                grid[row, column] = cellValue switch
                {
                    null => FormulaGridCell.Blank,
                    string { Length: 0 } => FormulaGridCell.Blank,
                    string formula when formula.StartsWith('=') => FormulaGridCell.Formula(formula),
                    _ => FormulaGridCell.Constant
                };
            }
        }

        return grid;
    }

    public static string ToA1Address(int row, int column)
    {
        if (row is < 1 or > 1_048_576 || column is < 1 or > 16_384)
        {
            throw new ArgumentOutOfRangeException(nameof(row));
        }

        Span<char> letters = stackalloc char[3];
        var index = letters.Length;
        var remaining = column;
        while (remaining > 0)
        {
            remaining--;
            letters[--index] = (char)('A' + (remaining % 26));
            remaining /= 26;
        }

        return new string(letters[index..]) + row.ToString(CultureInfo.InvariantCulture);
    }

    private static CellAddress ParseCell(string text)
    {
        var index = 0;
        var column = 0;
        while (index < text.Length && char.IsLetter(text[index]))
        {
            column = checked((column * 26) + char.ToUpperInvariant(text[index]) - 'A' + 1);
            index++;
        }

        if (index == 0 || !int.TryParse(text[index..], CultureInfo.InvariantCulture, out var row) ||
            row is < 1 or > 1_048_576 || column is < 1 or > 16_384)
        {
            throw new InvalidOperationException("Formula repair range contains an invalid cell address.");
        }

        return new CellAddress(row, column);
    }

    private sealed record CellAddress(int Row, int Column);
}

internal sealed record FormulaRangeBounds(int StartRow, int StartColumn, int EndRow, int EndColumn)
{
    public int RowCount => EndRow - StartRow + 1;

    public int ColumnCount => EndColumn - StartColumn + 1;
}
