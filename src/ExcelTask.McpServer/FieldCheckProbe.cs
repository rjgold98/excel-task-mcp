using System.Diagnostics;
using System.Text;
using System.Text.Json;

namespace ExcelTask.McpServer;

/// <summary>
/// Handshakes with any MCP stdio server and measures the tool surface it advertises.
/// The raw <c>tools/list</c> response is what a client must carry in context every session
/// before any work is requested, so its wire size is the fairest like-for-like comparison
/// between servers. This speaks JSON-RPC directly rather than through a client library so
/// the measured bytes are the server's own, not a re-serialization.
/// </summary>
internal static class FieldCheckProbe
{
    public static async Task<ToolSurface> MeasureAsync(
        string label,
        string path,
        IReadOnlyList<string> arguments,
        string workingDirectory,
        CancellationToken cancellationToken)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = path,
            WorkingDirectory = workingDirectory,
            UseShellExecute = false,
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            StandardOutputEncoding = Encoding.UTF8
        };
        foreach (var argument in arguments) startInfo.ArgumentList.Add(argument);

        Process? process = null;
        var timer = Stopwatch.StartNew();
        try
        {
            using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            timeout.CancelAfter(TimeSpan.FromSeconds(60));

            process = Process.Start(startInfo) ?? throw new InvalidOperationException("The server process did not start.");

            await WriteAsync(process, new
            {
                jsonrpc = "2.0",
                id = 1,
                method = "initialize",
                @params = new
                {
                    protocolVersion = "2025-06-18",
                    capabilities = new { },
                    clientInfo = new { name = "ExcelTask-FieldCheck", version = "1.0.0" }
                }
            });
            var initialize = await ReadResponseAsync(process, 1, timeout.Token);
            timer.Stop();

            await WriteAsync(process, new { jsonrpc = "2.0", method = "notifications/initialized" });
            await WriteAsync(process, new { jsonrpc = "2.0", id = 2, method = "tools/list", @params = new { } });
            var toolsLine = await ReadResponseAsync(process, 2, timeout.Token);

            using var initializeDocument = JsonDocument.Parse(initialize);
            using var toolsDocument = JsonDocument.Parse(toolsLine);
            var serverInfo = initializeDocument.RootElement.GetProperty("result").TryGetProperty("serverInfo", out var info)
                ? info
                : default;
            var tools = toolsDocument.RootElement.GetProperty("result").GetProperty("tools");
            var names = tools.EnumerateArray()
                .Select(tool => tool.TryGetProperty("name", out var name) ? name.GetString() ?? "" : "")
                .Where(name => name.Length > 0)
                .ToArray();

            return new ToolSurface(
                label,
                path,
                serverInfo.ValueKind == JsonValueKind.Object && serverInfo.TryGetProperty("name", out var serverName) ? serverName.GetString() : null,
                serverInfo.ValueKind == JsonValueKind.Object && serverInfo.TryGetProperty("version", out var version) ? version.GetString() : null,
                tools.GetArrayLength(),
                string.Join(", ", names),
                Encoding.UTF8.GetByteCount(toolsLine),
                Math.Round(timer.Elapsed.TotalSeconds, 2),
                null);
        }
        catch (Exception exception)
        {
            return new ToolSurface(label, path, null, null, 0, "", 0, 0, Describe(exception));
        }
        finally
        {
            if (process is not null)
            {
                try
                {
                    if (!process.HasExited)
                    {
                        process.StandardInput.Close();
                        if (!process.WaitForExit(5000)) process.Kill(entireProcessTree: true);
                    }
                }
                catch (InvalidOperationException) { }
                process.Dispose();
            }
        }
    }

    private static async Task WriteAsync(Process process, object message)
    {
        await process.StandardInput.WriteLineAsync(JsonSerializer.Serialize(message));
        await process.StandardInput.FlushAsync();
    }

    /// <summary>Reads newline-delimited JSON-RPC, skipping notifications and any non-JSON noise.</summary>
    private static async Task<string> ReadResponseAsync(Process process, int id, CancellationToken cancellationToken)
    {
        while (true)
        {
            var line = await process.StandardOutput.ReadLineAsync(cancellationToken);
            if (line is null) throw new InvalidOperationException("The server closed its output before replying.");
            if (line.Length == 0) continue;

            try
            {
                using var document = JsonDocument.Parse(line);
                if (document.RootElement.TryGetProperty("id", out var replyId) &&
                    replyId.ValueKind == JsonValueKind.Number &&
                    replyId.GetInt32() == id)
                {
                    if (document.RootElement.TryGetProperty("error", out var error))
                    {
                        throw new InvalidOperationException($"The server returned an error: {error}");
                    }

                    return line;
                }
            }
            catch (JsonException)
            {
                // Not a JSON-RPC frame; some servers print banners or logs to stdout.
            }
        }
    }

    private static string Describe(Exception exception) => exception is OperationCanceledException
        ? "Timed out waiting for the server to complete the MCP handshake."
        : $"{exception.GetType().Name}: {exception.Message}";
}
