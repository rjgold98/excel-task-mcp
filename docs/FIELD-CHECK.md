# Work-computer field check

The roadmap gates every phase on real use, not on the developer machine. This is
that gate. It has two halves: an automated half that needs no AI, and a short
manual half that measures what only the client can show.

The automated half is built into the released executable rather than shipped as
a script, because managed computers commonly run PowerShell in Constrained
Language Mode, which forbids the COM and reflection a scripted equivalent needs.
`--field-check` is not an MCP tool and not a public CLI; the model-facing surface
is still exactly one tool.

Nothing here changes an Excel setting, a security policy, or any workbook of
yours. If a managed policy blocks something, the report records that as a fact
about the machine rather than treating it as a product failure.

## Half 1 - automated (about three minutes)

1. Download `ExcelTask-<version>-windows-x64.zip` from the
   [latest release](https://github.com/rjgold98/excel-task-mcp/releases/latest)
   and extract it somewhere you can write without administrator rights, such as
   `%USERPROFILE%\ExcelTask`. A managed computer commonly refuses `C:\Tools`.
2. Run:

```powershell
%USERPROFILE%\ExcelTask\excel-task-mcp.exe --field-check
```

To measure the original Excel MCP alongside it, add its launch command. Only its
handshake and tool list are read; no workbook operation is sent to it. Repeat
`--compare-arg` once per argument the other server needs.

```powershell
%USERPROFILE%\ExcelTask\excel-task-mcp.exe --field-check --compare <other-server.exe> --compare-arg <arg>
```

Options: `--server <exe>` to measure a different build, `--output <dir>` to
choose where reports land. The default output is
`Desktop\ExcelTask-FieldCheck`. The exit code is `0` when every operation
completed and nothing was stranded, `1` otherwise.

Three files are written, and the third is printed to the console as well.

### What it reports beyond pass and fail

**Coverage.** The last line of the operation list names any operation the run
did not exercise. The check once covered five of twelve and still printed PASS,
so a subset is now stated rather than implied. On its first run the reporter
found a real gap — `RepairExistingWorksheet` — which is now a step.

**Sync roots.** `syncRootsRegistered` and `syncPathsResolving` measure the one
part of workbook identity that a developer machine cannot: resolving a path
under a OneDrive or SharePoint sync root back to the URL Excel reports for it.
Counts only. A `UrlNamespace` is the tenant and site collection, which is an
internal server name, and a `MountPoint` is a person's directory layout;
neither leaves the machine. `0 of N` means `UseOpen` against a workbook in those
roots will still refuse, and that the mapping's shape differs from what the
resolver expects.

**Folder writability.** `folder:documents`, `folder:desktop`,
`folder:downloads` and `folder:oneDriveRoot` each report the `ReadOnly`
attribute and whether the folder actually accepts a new file. Windows sets that
attribute on all four of an ordinary profile to mark a customized folder, and
until v0.16.0 it was read as a permission — so every copy-save and every create
into the folders workbooks live in was refused before Excel started, with a
reason that was false. `readOnlyAttribute=yes acceptsNewFile=yes` is the
expected pair, and confirms the attribute test is gone. Folders are named by
label; no path is written.

## Getting the results back

A managed computer often cannot send a file anywhere, so the check ends by
printing a **digest**: about ten dense lines carrying the Excel build, the VBA
trust values, both servers' tool-surface sizes, each operation's status and
elapsed time, and the leak count. It is deliberately short enough to retype or
photograph, and it is enough to decide what to do next.

```text
----- EXCELTASK FIELD DIGEST -----
excel=16.0.20228 vbom=1 vbawarn=1
self  v0.6.1 tools=1 bytes=7164
other tools=234 bytes=64512 ratio=9.0x
CopyExhibit (Plan)         Planned      3.0s L0
...
leaked=0 result=PASS
----- END DIGEST -----
```

Send the digest first. The full Markdown and JSON reports stay on disk and are
worth relaying only if something failed and the detail is needed.

What it records:

- **Environment** - Excel version and build, connected COM add-ins, the .NET
  runtime and `DOTNET_ROOT`, any PowerShell lockdown policy, whether "Trust
  access to the VBA project object model" is permitted, and whether Group
  Policy is setting macro security.
- **Tool surface** - for each server, how many tools it advertises and the exact
  wire size of its `tools/list` response. That payload is carried in context
  every session before any work is requested, so it is the fairest like-for-like
  context cost between two MCP servers. It is measured over raw JSON-RPC rather
  than through a client library, so the bytes are the server's own.
- **Operations** - copy exhibit, extend formula series, and macro edit, each run
  against a disposable workbook through the real MCP boundary, with elapsed
  seconds, the resulting status, every verification check, and whether any Excel
  process was left behind. The check's own fixture processes are tracked and
  excluded, so a leak figure is never the harness being mistaken for the product.

A `Rejected` or `Unknown` status is still a useful result. The check detail says
which preflight step refused and why.

## Half 2 - manual, in the client (about fifteen minutes)

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
| Tokens | Only if the client exposes a usage figure; leave blank otherwise |

The number that matters most is the first one: how long from your prompt to the
work actually being done.

## Configuring ExcelTask side by side

`.vscode/mcp.json` in this repository already points at
`${userHome}\ExcelTask\excel-task-mcp.exe`, and clients that sync a repository's
MCP configuration will pick it up. The home directory is used because a managed
computer commonly refuses to let you write under `C:\`. To configure it
elsewhere, use a distinct server key and the full executable path so the original
install is untouched:

```json
{
  "servers": {
    "excel-task": {
      "command": "${userHome}\\ExcelTask\\excel-task-mcp.exe"
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

## Working directory

Real workbooks commonly live in OneDrive or SharePoint-synced folders. The field
check only ever uses disposable workbooks in the system temp directory, so sync
behaviour cannot affect its results. Measuring ExcelTask against a synced folder
is a separate, later task.
