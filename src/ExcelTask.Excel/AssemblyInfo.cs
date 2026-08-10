using System.Runtime.CompilerServices;

[assembly: InternalsVisibleTo("ExcelTask.Excel.Tests")]

// The field check and the fixture builders drive Excel over the same late-bound object models, so
// they use the same ComAccess rules rather than restating them. Sharing internals keeps that
// knowledge in one module without widening the product's public surface.
[assembly: InternalsVisibleTo("excel-task-mcp")]
[assembly: InternalsVisibleTo("ExcelTask.McpServer.Tests")]
