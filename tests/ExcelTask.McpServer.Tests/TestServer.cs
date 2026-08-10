using System.Runtime.InteropServices;
using ModelContextProtocol.Client;

namespace ExcelTask.McpServer.Tests;

/// <summary>Locates the MCP server under test and builds its child-process environment.</summary>
internal static class TestServer
{
    public static string ServerPath
    {
        get
        {
            var configured = Environment.GetEnvironmentVariable("EXCELTASK_TEST_SERVER_PATH");
            return string.IsNullOrWhiteSpace(configured)
                ? Path.Combine(AppContext.BaseDirectory, "excel-task-mcp.exe")
                : Path.GetFullPath(configured);
        }
    }

    /// <summary>
    /// The transports launch the server with environment inheritance disabled, so a
    /// framework-dependent apphost cannot see DOTNET_ROOT. On a machine where .NET is installed
    /// per-user - no HKLM registration and nothing under Program Files, which is how managed
    /// work computers commonly install it - runtime resolution then fails with 0x80008083 before
    /// the server ever starts. Point the child at the runtime hosting this test instead.
    /// </summary>
    public static Dictionary<string, string?> EnvironmentVariables()
    {
        var environment = StdioClientTransportOptions.GetDefaultEnvironmentVariables();
        // e.g. <root>\shared\Microsoft.NETCore.App\<version>\ -> <root>
        var root = Path.GetFullPath(Path.Combine(RuntimeEnvironment.GetRuntimeDirectory(), "..", "..", ".."));
        if (Directory.Exists(Path.Combine(root, "shared")))
        {
            environment.TryAdd("DOTNET_ROOT", root);
            environment.TryAdd("DOTNET_ROOT(x64)", root);
        }

        return environment;
    }
}
