# The current work-computer field task

The standing task for an agent running on the managed work computer. It replaces
the per-release issues that preceded it, which went stale the moment the release
they named was superseded — issue #6 asked for `ExcelTask-0.14.1-...zip` from
`releases/latest`, and by the time anyone read it `releases/latest` was four
releases further on and contained no such file.

**Target: the latest release.** Never a pinned version. `docs/FIELD-CHECK.md`
describes what the check measures; this describes what to do and report.

## Boundaries

These bind, and they come from `docs/AGENT-BRIDGE.md`.

- **Perform no git write of any kind** — no push, no branch, no PR, no issue
  comment. Report as files and give Ross the paths.
- Never change Excel, Trust Center, or any security or system setting. If a
  policy blocks a step, **record the block verbatim and continue with the
  remaining steps**. A block is a result, not an obstacle to work around.
- Touch only the disposable fixtures the check creates in temp. Never a real
  business workbook.
- Do not reconfigure or interfere with any other MCP install. Measuring is fine.
- Anything this task does not cover is a question for Ross, not an improvisation.

## Step 1 — the automated check (required, ~3 min)

Download `ExcelTask-<latest>-windows-x64.zip` from
<https://github.com/rjgold98/excel-task-mcp/releases/latest> and extract to
`%USERPROFILE%\ExcelTask`. Use the home folder; a managed machine commonly
refuses writes under `C:\`.

```powershell
%USERPROFILE%\ExcelTask\excel-task-mcp.exe --field-check
```

Compiled .NET, not a script, specifically so Constrained Language Mode does not
block it. It builds its own disposable workbooks in temp, changes no setting,
and touches nothing of yours.

Report the digest verbatim, the exit code, and the three file paths.

### Results that are interesting rather than bad

- **`vbom=0`** — VBA project access is blocked by policy. Macro Plan returns no
  hash and Apply is skipped. This is correct behaviour. Report it; change
  nothing.
- **`Coverage:`** — names any operation the run did not exercise. If it names
  any, report it. A PASS covers only what ran.
- **`syncRootsRegistered=0`** — nothing is synced, so the SharePoint identity
  path could not be exercised and step 3 is the only way to settle it.
- **`syncPathsResolving`** showing fewer than registered — the mapping on this
  machine differs from what the resolver expects. This is a finding worth
  stopping for; it means `UseOpen` on synced storage still refuses.
- **`folder:*` showing `readOnlyAttribute=yes acceptsNewFile=no`** — a genuinely
  unwritable folder, unlike the attribute-only case. Report which.

### Results that are bad

- **Any `Unknown` status.** The one result worth stopping to investigate.
  Capture the full check text for that row from the Markdown report.
- **`leaked=` anything but 0.** Real since v0.14.1, where the counter was fixed
  to wait for Excel's asynchronous exit rather than counting a dying process.
  Report it with the surrounding rows.

## Step 2 — if anything fails or hangs (only then)

```powershell
$env:EXCELTASK_TRACE = "$env:USERPROFILE\exceltask-trace.log"
%USERPROFILE%\ExcelTask\excel-task-mcp.exe --field-check
Remove-Item Env:EXCELTASK_TRACE
```

The log is safe to share by construction and states its own contract in its
header. **A phase with no matching `phase end` is where it hung** — the single
most useful line in the file, and no receipt can report it.

## Step 3 — the OneDrive binding (required if `syncRootsRegistered` > 0, ~5 min)

The highest-value unproven thing in the product. Workbook identity resolves a
service URL back to the local path through the sync client's registry mapping,
and until this runs on a synced machine only the arithmetic around it is tested.

With **a disposable copy** of any workbook inside a synced folder open in Excel,
send a `ReadWorksheetRange` with `workbookBinding: "UseOpen"` against its
**local** path.

Report: whether it bound or refused, and the **exact** refusal text if it
refused. Before v0.16.0 this refused every time, reporting that the workbook
name matched and the path did not.

## Step 4 — the full test suite (required if the SDK is present, ~10 min)

```powershell
git -C <your local clone> pull
dotnet test ExcelTask.slnx --filter "RunType!=OnDemand" -p:NuGetAudit=false
.\scripts\Test-Mvp.ps1 -IncludeExcel
```

Report the pass/fail/skip counts **per assembly, verbatim**. "Tests passed"
without counts will be sent back. Current expected counts are in `ROADMAP.md`
under Delivered, which is updated every release.

If .NET is a per-user install, the harness needs `DOTNET_ROOT` in the child
environment. That was fixed in v0.4.1; if a child-process runtime error appears,
report its exact text.

## Step 5 — optional: measure another server (~5 min)

```powershell
%USERPROFILE%\ExcelTask\excel-task-mcp.exe --field-check --compare <other-server.exe>
```

Only the handshake and tool list of the other server are read; **no workbook
operation is sent to it**. Repeat `--compare-arg <arg>` per argument it needs.
This produces the tool-surface ratio — the context cost every session pays
before any work is requested.

## Step 6 — optional: the client comparison (~15 min)

The automated half cannot see tokens or turns; only the client knows those. Pick
three tasks resembling real work — one exhibit built from a modelled worksheet,
one formula repair or extension, one macro edit — and run each twice, once with
ExcelTask and once with the other server.

Use whichever model is already selected. **Do not switch models between runs and
do not force one**, which would make the runs incomparable.

Record per run: wall-clock from pressing enter to the change being finished,
tool calls, retries or self-corrections, whether the workbook ended up correct
(verified by opening it), any stray `EXCEL.EXE` afterwards, and tokens only if
the client exposes them. The first number matters most: prompt to work done.

## Reporting

Files on the work computer, paths to Ross, **no git operations**.

1. The three generated field-check files, unedited.
2. The trace log, if step 2 ran.
3. One `SUMMARY.md` containing:
   - each command exactly as executed;
   - the digest verbatim and every exit code;
   - test counts per assembly;
   - the environment table the check reports;
   - the step 3 result, with exact text if it refused;
   - the exact text of anything that failed — pasted, never summarized;
   - anything surprising the tooling did not capture.
