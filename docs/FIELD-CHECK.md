# Work-computer field check

The roadmap gates every phase on real use, not on the developer machine. This is
that gate. It has two halves: an automated half that needs no AI, and a short
manual half that measures what only Copilot can show.

Nothing here changes an Excel setting, a security policy, or any workbook of
yours. If a managed policy blocks something, the report records that as a fact
about the machine rather than treating it as a product failure.

## Half 1 - automated (about three minutes)

1. Download `ExcelTask-<version>-windows-x64.zip` from the
   [latest release](https://github.com/rjgold98/excel-task-mcp/releases/latest)
   and extract it, for example to `C:\Tools\ExcelTask`.
2. Copy `scripts\Invoke-FieldCheck.ps1` next to it.
3. Run:

```powershell
.\Invoke-FieldCheck.ps1 -ServerPath C:\Tools\ExcelTask\excel-task-mcp.exe
```

To measure the original Excel MCP alongside it, add its launch command. Only its
handshake and tool list are read; no workbook operation is sent to it.

```powershell
.\Invoke-FieldCheck.ps1 -ServerPath C:\Tools\ExcelTask\excel-task-mcp.exe -CompareServerPath <path-to-the-other-server.exe>
```

Two files land on the Desktop under `ExcelTask-FieldCheck`. Send both back.

What it records:

- **Environment** - Excel version and build, connected COM add-ins, whether
  "Trust access to the VBA project object model" is permitted, and whether Group
  Policy is setting macro security.
- **Tool surface** - how many tools the server advertises and how many bytes its
  `tools/list` payload occupies. That payload is carried in context every
  session before any work is requested, so it is the fairest like-for-like
  context cost between two MCP servers.
- **Operations** - copy exhibit, extend formula series, and macro edit, each run
  against a disposable workbook, with elapsed seconds, the resulting status, every
  verification check, and whether any Excel process was left behind.

A `Rejected` or `Unknown` status is still a useful result. The check detail says
which preflight step refused and why.

## Half 2 - manual, in Copilot (about fifteen minutes)

The automated half cannot see tokens or turns, because only the client knows
those. Pick three tasks that resemble real work - one exhibit built from a
modelled worksheet, one formula repair or extension, one macro edit - and run
each one twice: once with ExcelTask configured, once with the original server.

Use whichever model is already selected. Do not switch models between runs, and
do not force a specific model; that would make the two runs incomparable.

For each of the six runs, record:

| Field | Meaning |
|---|---|
| Wall-clock time | From pressing enter to the change being finished in Excel |
| Tool calls | How many times the agent invoked the server |
| Retries or corrections | How often it had to fix its own call |
| Result correct | Did the workbook end up right, verified by opening it |
| Excel left behind | Any stray `EXCEL.EXE` in Task Manager afterwards |
| Tokens | Only if Copilot exposes a usage figure; leave blank otherwise |

The number that matters most is the first one: how long from your prompt to the
work actually being done.

## Configuring ExcelTask side by side

Use a distinct server key and the full executable path, so the original install
is untouched:

```json
{
  "servers": {
    "excel-task": {
      "command": "C:\\Tools\\ExcelTask\\excel-task-mcp.exe"
    }
  }
}
```

## If macro editing is refused

On a managed computer, reading or editing VBA requires "Trust access to the VBA
project object model", which the organization controls. The report shows the
current value. ExcelTask will never change it, and refusing to edit is the
correct behaviour when it is off - the point of recording it is so the roadmap
reflects what is actually possible on that machine.
