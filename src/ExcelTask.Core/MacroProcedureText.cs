using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelTask.Core;

/// <summary>Deterministic text handling and structural validation for one bounded VBA procedure.</summary>
public static partial class MacroProcedureText
{
    public const int MaxSourceCharacters = 8_192;
    public const int MaxLineCount = 200;

    /// <summary>Normalizes VBA source line endings to LF before hashing or replacement.</summary>
    public static string NormalizeLineEndings(string source)
    {
        ArgumentNullException.ThrowIfNull(source);
        return source.Replace("\r\n", "\n", StringComparison.Ordinal).Replace('\r', '\n').TrimEnd('\n');
    }

    /// <summary>Computes the lowercase SHA-256 fingerprint of line-ending-normalized source.</summary>
    public static string ComputeSha256(string source)
    {
        var bytes = Encoding.UTF8.GetBytes(NormalizeLineEndings(source));
        return Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
    }

    /// <summary>
    /// VBA constructs that open a modal dialog or enter break mode. Each one blocks the automation
    /// thread until a person clicks something, which is exactly the outcome the product must never
    /// produce, so a request that asks to run the procedure is refused before Excel is opened.
    /// This is a first line of defence only: it cannot see into procedures the replacement calls.
    /// </summary>
    public static bool TryFindBlockingConstruct(string source, out string? construct)
    {
        ArgumentNullException.ThrowIfNull(source);
        construct = null;
        foreach (var line in NormalizeLineEndings(source).Split('\n'))
        {
            var code = StripStringLiterals(StripComment(line));
            var match = BlockingConstructRegex().Match(code);
            if (match.Success)
            {
                construct = match.Value;
                return true;
            }
        }

        return false;
    }

    /// <summary>Replaces literal text with spaces so a construct name inside a string is not mistaken for code.</summary>
    private static string StripStringLiterals(string line)
    {
        if (!line.Contains('"', StringComparison.Ordinal)) return line;

        var builder = new StringBuilder(line.Length);
        var quoted = false;
        foreach (var character in line)
        {
            if (character == '"') quoted = !quoted;
            builder.Append(quoted || character == '"' ? ' ' : character);
        }

        return builder.ToString();
    }

    /// <summary>Validates and normalizes exactly one complete replacement procedure without exposing its text in errors.</summary>
    public static bool TryNormalizeProcedureSource(
        string? source,
        string requiredProcedureName,
        bool requireZeroParameters,
        out string? normalizedSource,
        out string? error)
    {
        normalizedSource = null;
        error = null;
        if (source is null)
        {
            error = "Replacement procedure source is required.";
            return false;
        }
        if (source.Length > MaxSourceCharacters)
        {
            error = $"Replacement procedure source exceeds the limit of {MaxSourceCharacters:N0} characters.";
            return false;
        }

        var normalized = NormalizeLineEndings(source);
        if (CountLines(normalized) > MaxLineCount)
        {
            error = $"Replacement procedure source exceeds the limit of {MaxLineCount} lines.";
            return false;
        }
        if (!TryParseSingleProcedure(normalized, out var procedure, out error))
        {
            return false;
        }
        if (!string.Equals(procedure.Name, requiredProcedureName, StringComparison.OrdinalIgnoreCase))
        {
            error = "Replacement procedure name must match the requested procedure.";
            return false;
        }
        if (requireZeroParameters && procedure.HasParameters)
        {
            error = "RunAfterEdit requires a replacement procedure with zero parameters.";
            return false;
        }

        normalizedSource = normalized;
        return true;
    }

    private static bool TryParseSingleProcedure(string normalizedSource, out ParsedProcedure procedure, out string? error)
    {
        procedure = default;
        error = null;
        var lines = normalizedSource.Split('\n');
        var foundStart = false;
        var foundEnd = false;
        foreach (var line in lines)
        {
            var code = StripComment(line).Trim();
            if (!foundStart)
            {
                if (code.Length == 0) continue;

                var start = ProcedureStartRegex().Match(code);
                if (!start.Success)
                {
                    error = "Replacement source must contain exactly one procedure and no module-level declarations.";
                    return false;
                }

                foundStart = true;
                procedure = new ParsedProcedure(
                    start.Groups["kind"].Value,
                    start.Groups["name"].Value,
                    HasNonWhitespaceParameterText(start.Groups["parameters"].Value));
                continue;
            }

            if (foundEnd)
            {
                if (code.Length != 0)
                {
                    error = "Replacement source must contain exactly one procedure and no module-level declarations.";
                    return false;
                }
                continue;
            }

            if (ProcedureStartRegex().IsMatch(code))
            {
                error = "Replacement source must contain exactly one procedure.";
                return false;
            }

            var end = ProcedureEndRegex().Match(code);
            if (end.Success)
            {
                if (!string.Equals(end.Groups["kind"].Value, procedure.Kind, StringComparison.OrdinalIgnoreCase))
                {
                    error = "Replacement procedure end does not match its declaration.";
                    return false;
                }
                foundEnd = true;
            }
        }

        if (!foundStart || !foundEnd)
        {
            error = "Replacement source must contain one complete Sub or Function procedure.";
            return false;
        }

        return true;
    }

    private static int CountLines(string source) => source.Length == 0 ? 0 : source.Count(character => character == '\n') + 1;

    private static bool HasNonWhitespaceParameterText(string parameters) => parameters.Length > 2 && !string.IsNullOrWhiteSpace(parameters[1..^1]);

    private static string StripComment(string line)
    {
        var quoted = false;
        for (var index = 0; index < line.Length; index++)
        {
            if (line[index] == '"')
            {
                if (quoted && index + 1 < line.Length && line[index + 1] == '"')
                {
                    index++;
                    continue;
                }
                quoted = !quoted;
                continue;
            }
            if (!quoted && line[index] == '\'') return line[..index];
        }

        var trimmed = line.TrimStart();
        return trimmed.StartsWith("Rem ", StringComparison.OrdinalIgnoreCase) || string.Equals(trimmed, "Rem", StringComparison.OrdinalIgnoreCase)
            ? string.Empty
            : line;
    }

    private readonly record struct ParsedProcedure(string Kind, string Name, bool HasParameters);

    // The parameter list admits empty () pairs so that array parameters such as
    // "ByRef values() As Variant" parse; the alternatives are disjoint on their first
    // character, so this cannot backtrack pathologically.
    [GeneratedRegex(@"^(?:(?:Public|Private|Friend|Static)\s+)*(?<kind>Sub|Function)\s+(?<name>[A-Za-z_][A-Za-z0-9_]*)\s*(?<parameters>\((?:[^()]|\(\s*\))*\))?\s*(?:As\s+[A-Za-z_][A-Za-z0-9_]*(?:\s*\(\s*\))?)?$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex ProcedureStartRegex();

    [GeneratedRegex(@"^End\s+(?<kind>Sub|Function)$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex ProcedureEndRegex();

    // MsgBox/InputBox/GetOpenFilename/GetSaveAsFilename/FileDialog raise a modal dialog; Stop drops
    // the VBA host into break mode. All of them wait for a person the product cannot summon.
    [GeneratedRegex(@"\b(MsgBox|InputBox|GetOpenFilename|GetSaveAsFilename|FileDialog|Stop)\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex BlockingConstructRegex();
}
