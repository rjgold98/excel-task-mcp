namespace ExcelTask.McpServer.Tests;

/// <summary>
/// Removes a test's temp directory without letting the removal decide the test's verdict. See the
/// twin in ExcelTask.Excel.Tests: a throw from a finally block fails a test whose assertions all
/// passed, and the macro round-trip test did exactly that - it proved owned Excel had exited, then
/// failed on rmdir because Windows had not yet released the file behind the process it watched die.
/// </summary>
internal static class TempDirectory
{
    public static void Remove(string path)
    {
        for (var attempt = 0; attempt < 20; attempt++)
        {
            try
            {
                if (!Directory.Exists(path)) return;
                Directory.Delete(path, recursive: true);
                return;
            }
            catch (IOException) { Thread.Sleep(100); }
            catch (UnauthorizedAccessException) { Thread.Sleep(100); }
        }
    }
}
