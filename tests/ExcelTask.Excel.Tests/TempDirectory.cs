namespace ExcelTask.Excel.Tests;

/// <summary>
/// Removes a test's temp directory without letting the removal decide the test's verdict.
///
/// The plain <c>Directory.Delete(path, recursive: true)</c> in a finally block throws
/// <see cref="IOException"/> when anything still holds a handle, and a throw from finally fails a
/// test whose assertions all passed. That is exactly what happened to the macro round-trip test:
/// it proved the owned Excel process had exited, then failed on rmdir because Windows had not yet
/// released the file behind the process it had just watched die.
///
/// Excel releases handles a beat after the process goes, so this retries briefly and then gives up
/// silently. A leftover temp directory is litter; a false red is a lie about the product.
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
