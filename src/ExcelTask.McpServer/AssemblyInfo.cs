using System.Runtime.CompilerServices;

// The field check is ~900 lines compiled into the shipped binary behind a --field-check argument,
// and it had no test of any kind - which is how it came to validate five of eleven operations and
// report PASS. Sharing internals lets its coverage arithmetic be asserted without widening the
// server's surface: none of this is reachable from the MCP client.
[assembly: InternalsVisibleTo("ExcelTask.McpServer.Tests")]
