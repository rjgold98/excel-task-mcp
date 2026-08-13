# The current work-computer field task

The standing task for an agent running on the managed work computer. It replaces
the per-release issues that preceded it, which went stale the moment the release
they named was superseded — issue #6 asked for `ExcelTask-0.14.1-...zip` from
`releases/latest`, and by the time anyone read it `releases/latest` was four
releases further on and contained no such file.

**Target: the latest release.** Never a pinned version.

## What this run is for

Three questions, in order of value. Everything else is confirmation.

1. **Does workbook identity work on synced storage?** ExcelTask resolves a
   workbook Excel reports as a SharePoint URL back to the local path you named.
   That resolution has never executed anywhere — this machine is the only one
   that can run it. Step 5.
2. **What does the agent reach for instead of ExcelTask?** A previous session
   made 216 PowerShell calls against 71 ExcelTask calls. Knowing which receipt
   sent it away is worth more than any feature on the roadmap. Step 6.
3. **Do all sixteen operations still work on managed hardware?** Step 3.

## Boundaries

These bind, and they come from `docs/AGENT-BRIDGE.md`.

- **Perform no git write of any kind** — no push, no branch, no PR, no issue
  comment. Report as files and give Ross the paths.
- Never change Excel, Trust Center, or any security or system setting. If a
  policy blocks a step, **record the block verbatim and continue with the
  remaining steps**. A block is a result, not an obstacle to work around.
- Touch only the disposable fixtures the check creates in temp, and the one
  disposable copy step 5 asks for. Never a real business workbook.
- Do not reconfigure or interfere with any other MCP install. Measuring is fine.
- Anything this task does not cover is a question for Ross, not an improvisation.

---

## Step 0 — record what else is loaded (~2 min)

This is new, and it comes from a finding rather than a guess. Two Excel skills
were enabled on this machine during earlier sessions. One names nineteen tools,
none of which exist here — ExcelTask advertises exactly one, `excel_task`. The
other drives a **separate** Excel automation binary with its own sessions, and
its own documentation says unclosed sessions leave Excel processes running.

That second one matters to this task specifically: an Excel process started by
something else **while the check is running** is counted as a process ExcelTask
leaked, and the run prints `result=FAIL` for a leak that is not ours. Step 1a
clears strays *before* the run; nothing protects the run from something started
*during* it.

1. List every skill currently **loaded and enabled**, whether or not it fires.
   In the Copilot app that is `/skills`. **Paste the list into the report.**
2. If any Excel-related skill is enabled, **say so and stop for Ross's decision**
   before disabling anything. Disabling a client-side skill is not the same as
   touching another MCP install, but it is close enough to the boundary above
   that the owner makes the call, not the agent.
3. Run nothing else that automates Excel for the duration of this task.

A skill that shaped behaviour without ever being invoked is invisible to an
invocation log, which is why the inventory is asked for separately from step 6.

---

## Step 1 — replace the installed build (~4 min)

This is the step most likely to go wrong quietly. Windows will not let you delete
a running `.exe`, and a self-contained build is ~231 files — so extracting over
the top of an old one leaves stale files behind and can produce a mixed install
that reports the new version while running old code. **Delete, then extract.**

### 1a. Stop the server so its files are unlocked

The MCP host keeps `excel-task-mcp.exe` running while the client is up.

1. In the GitHub Copilot app, **disable the ExcelTask MCP server** (or quit the
   app entirely — quitting is more reliable).
2. Confirm nothing is holding the files:

```powershell
Get-Process excel-task-mcp -ErrorAction SilentlyContinue
```

Expect **no output**. If a process is listed, the client is still running it —
quit the app and check again. Only if it persists after the app is closed:

```powershell
Stop-Process -Name excel-task-mcp -Force
```

3. Clear any stray Excel left from earlier runs, so the leak count means
   something. **Do not do this if you have unsaved work open in Excel** — save
   and close by hand first.

```powershell
Get-Process EXCEL -ErrorAction SilentlyContinue
```

If any are listed and none of them is yours:

```powershell
Stop-Process -Name EXCEL -Force
```

### 1b. Note what is installed now, then remove it

```powershell
(Get-Item "$env:USERPROFILE\ExcelTask\excel-task-mcp.exe").VersionInfo.ProductVersion
```

That prints something like `0.16.0+6a7d4f7…` — the version **and** the exact
commit it was built from, so there is no ambiguity about what was installed.

Record that version — it is the "upgraded from" line in your report. Then:

```powershell
Remove-Item "$env:USERPROFILE\ExcelTask" -Recurse -Force
```

If that fails with a file-in-use error, something is still running it. Go back
to 1a. **Do not** work around it by extracting to a different folder — the
Copilot MCP config points at this path, and a second copy elsewhere means you
will test one build while the client runs another.

### 1c. Download the newest release

Go to <https://github.com/rjgold98/excel-task-mcp/releases/latest> and download
the asset named `ExcelTask-<version>-windows-x64.zip`. There is exactly one zip
per release. A browser download to `%USERPROFILE%\Downloads` is fine and is
usually the least friction on a managed machine.

If the GitHub CLI is available, this is faster and avoids picking the wrong file:

```powershell
gh release download --repo rjgold98/excel-task-mcp --pattern "ExcelTask-*-windows-x64.zip" --dir "$env:USERPROFILE\Downloads" --clobber
```

### 1d. Extract

Pick the archive explicitly rather than with a wildcard. After a second run there
will be more than one release zip in `Downloads`, and `Expand-Archive` fails
outright when its path matches several — this line takes the newest and prints
which one it chose, so a stale zip cannot be extracted by accident.

```powershell
$zip = Get-ChildItem "$env:USERPROFILE\Downloads\ExcelTask-*-windows-x64.zip" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
$zip.Name
Expand-Archive -Path $zip.FullName -DestinationPath "$env:USERPROFILE\ExcelTask" -Force
```

Use the home folder. A managed machine commonly refuses writes under `C:\`.

Windows marks downloaded files as blocked, which can stop the executable from
starting. Clear that:

```powershell
Get-ChildItem "$env:USERPROFILE\ExcelTask" -Recurse | Unblock-File
```

### 1e. Confirm you are on the build you think you are

```powershell
(Get-Item "$env:USERPROFILE\ExcelTask\excel-task-mcp.exe").VersionInfo.ProductVersion
(Get-ChildItem "$env:USERPROFILE\ExcelTask" -Recurse -File | Measure-Object).Count
```

The version must match the release you downloaded, and the file count is **231**
for a v0.17.0 build — expect roughly that. A count in the single digits means the
archive did not fully extract, and a count far above it means the old install was
not removed first.

**Report: the version you removed, the version you installed with its commit
hash, and the file count.**

---

## Step 2 — re-enable the server in Copilot

Turn the ExcelTask MCP server back on (or reopen the app). Confirm the client
lists **one** tool, named `excel_task`. If it lists more than one, or none, stop
and report that — it means the client is pointed at something other than this
build.

Nothing else in the Copilot configuration should change. The install path is the
same as before precisely so no config edit is needed.

---

## Step 3 — the automated check (required, ~3 min)

This is a compiled .NET executable, not a script, specifically so Constrained
Language Mode does not block it. It builds its own disposable workbooks in temp,
changes no setting, and touches nothing of yours.

```powershell
& "$env:USERPROFILE\ExcelTask\excel-task-mcp.exe" --field-check
```

It prints a digest of about twenty lines and writes three files to
`Desktop\ExcelTask-FieldCheck`.

**What good looks like:** every operation row reads `Completed` or `Planned`,
the last line reads `Coverage: all 16 operations exercised.`, and the digest
ends `leaked=0 result=PASS`. Exit code `0`.

A live row ending `excelStillUp=1` or `=2` is **not** a leak. It means Excel was
still shutting down when that operation stopped waiting, which is ordinary here:
four connected COM add-ins load into every instance and unload again on exit, so
teardown runs past the twenty seconds the per-row wait allows. The written report
reconciles every row against a final snapshot taken after everything has stopped.
Judge leaks by the report's `Leaked Excel` column and the digest's `leaked=`,
never by the console line — the 2026-08-12 run printed `L2` against three
operations and `leaked=0 result=PASS` in the same file.

The four Data Model rows — two `ManageModelMeasure`, two `ManageModelRelationship`
— need a Data Model, which needs Power Query.
If policy forbids it here, those rows are absent and a note says why — that is
a machine answer, not a product failure, and the run can still pass. **Report
the note verbatim if it appears.**

```powershell
$LASTEXITCODE
```

**Report: the digest verbatim, the exit code, and the three file paths.**

### Results that are interesting rather than bad

Report each; change nothing.

| Reading | Means |
|---|---|
| `vbom=0` | VBA project access is blocked by policy. Macro Plan returns no hash and Apply is skipped. **Correct behaviour**, and it tells the roadmap what is possible here. |
| `syncRootsRegistered=0` | Nothing is synced. Step 5 cannot run, and that is itself the answer. |
| `syncPathsResolving` less than registered | **A finding worth stopping for.** The mapping on this machine differs from what the resolver expects, so `UseOpen` on synced storage will still refuse. |
| `folder:* readOnlyAttribute=yes acceptsNewFile=yes` | Expected, and confirms a fixed defect. Windows marks these folders; they are writable. |
| `folder:* acceptsNewFile=no` | A genuinely unwritable folder. Report which. |

### Results that are bad

- **Any `Unknown` status.** The one result worth stopping to investigate. Open
  the Markdown report and copy the full check text for that row.
- **`leaked=` anything but 0.** Real since v0.14.1, where the counter was fixed
  to wait for Excel's asynchronous exit rather than counting a dying process.
  Report it with the surrounding rows.
- **`Coverage:` naming any operation.** A PASS covers only what ran.

---

## Step 4 — only if something failed or hung

Re-run with the trace on. The log states its own contract in its header: phases,
durations, operation kinds, worksheet names, A1 ranges, workbook **file names
only**, process ids, statuses. It never records cell values, formulas, VBA
source, connection strings, or full paths.

What its header no longer says — through v0.18.0 it did — is that the result is
**safe to share**. Read that first list again: worksheet names and workbook file
names are in it. Where this run touched a real workbook, those names are real.
The file states its contents; whether it can leave this machine is your call.
See the reporting section at the end before relaying it.

```powershell
$env:EXCELTASK_TRACE = "$env:USERPROFILE\exceltask-trace.log"
& "$env:USERPROFILE\ExcelTask\excel-task-mcp.exe" --field-check
Remove-Item Env:EXCELTASK_TRACE
```

**A phase with no matching `phase end` is where it hung.** That is the single
most useful line in the file, and no receipt can report it.

---

## Step 5 — the OneDrive binding (the important one, ~5 min)

**Skip only if step 3 reported `syncRootsRegistered=0`.**

ExcelTask resolves a workbook Excel reports as a service URL back to the local
path you named, through the sync client's registry mapping. Only the arithmetic
around that is covered by tests — the lookup itself has never run anywhere.
Before v0.16.0 this refused every time, reporting that the workbook name matched
and the path did not.

1. Take **a disposable copy** of any workbook that lives inside a synced OneDrive
   or SharePoint folder. Copy it, rename the copy, and work only on the copy.
   Never the original.
2. Open the copy in Excel and leave it open.
3. Through the Copilot client, ask ExcelTask for a **`ReadWorksheetRange`** on a
   small range of that copy, using its **local** path (the one Explorer shows,
   not a URL), with `workbookBinding` set to `UseOpen`.
4. Then ask for a **`ScanWorkbookStructure`** on the same local path.

**Report, exactly:**

- whether the read bound to the open workbook or refused;
- if it refused, the **complete refusal text**, pasted;
- whether the scan succeeded (it reads the file directly, so it should work
  regardless — if it fails while the read succeeds, that is a separate finding);
- the local path's **shape only**, never the path itself: for example
  "under the OneDrive root, two folders deep".

Delete the disposable copy afterwards.

---

## Step 6 — the session analytics export (~5 min)

A previous session made **216 PowerShell calls against 71 ExcelTask calls**, and
the PowerShell was almost entirely one question asked over and over: *what
workbook does Excel currently have open?* The agent tried PowerShell, then
Python, then compiling C#. Knowing which receipt sends the agent away is the most
valuable diagnostic available, and it is invisible from the server side.

After you have finished steps 3 to 5, produce `ANALYTICS.md` from the local
session event stream. **Metadata only** — no workbook values, no formula text, no
VBA, no full paths, no prompt or reasoning content.

1. Every tool call in time order: timestamp, tool name, duration in ms, status
   (completed / errored / aborted). For `excel_task` calls also give the
   operation kind, the mode, and the receipt status returned. **Not** arguments.
2. Any `tool.execution_start` with no matching `tool.execution_complete`: which
   tool, when, and what followed it.
3. Every skill invocation: name, timestamp, duration. Separately, every skill
   **loaded and enabled** for the session, invoked or not — the step 0 list. A
   skill that shaped the agent's behaviour without ever firing leaves no
   invocation to log, and would otherwise confound exactly the question this
   step exists to answer.
4. Every subagent: label, duration, whether it completed.
5. **For each `excel_task` call, the tool names of the three calls immediately
   before and after it.** This is the important one — it shows what the agent
   reached for instead of the MCP.
6. Any error or refusal text returned **by** `excel_task`, verbatim. These are
   ExcelTask's own messages and carry no workbook content.
7. Token count at each usage checkpoint, and what triggered any compaction or
   truncation.

Note if the client cannot export some of this. Raw reasoning traces are usually
stored as opaque or encrypted blocks and are not expected to be recoverable —
their absence is not a failure, and they are not needed.

---

## Step 7 — the test suite (only if the .NET SDK and a clone are present)

Skip this entirely if either is missing, and say so. It is a developer-machine
check that happens to be worth repeating here; it is not what this run is for.

```powershell
dotnet --version
```

If that fails, the SDK is absent — skip to reporting. Otherwise, from your local
clone:

```powershell
git -C <your local clone> pull
dotnet test ExcelTask.slnx --filter "RunType!=OnDemand" -p:NuGetAudit=false
```

**Expected: 287 fast tests, zero failures** — Core 156, Excel 103, McpServer 28.
The current counts are also in `ROADMAP.md` under Delivered, updated each
release; if these disagree with the roadmap, trust the roadmap and report the
difference.

Then the full gate, which drives real Excel and takes about ten minutes:

```powershell
.\scripts\Test-Mvp.ps1 -IncludeExcel
```

**Expected: 341 total, zero failures, and no Excel process left behind.**

**Report the pass/fail/skip counts per assembly, verbatim.** "Tests passed"
without counts will be sent back.

If .NET is a per-user install, the harness needs `DOTNET_ROOT` in the child
environment. That was fixed in v0.4.1; if a child-process runtime error appears,
report its exact text.

---

## Step 8 — optional: measure another server (~5 min)

Only the handshake and tool list of the other server are read. **No workbook
operation is sent to it.**

```powershell
& "$env:USERPROFILE\ExcelTask\excel-task-mcp.exe" --field-check --compare <other-server.exe>
```

Repeat `--compare-arg <arg>` once per argument it needs. This produces the
tool-surface ratio — the context cost every session pays before any work is
requested.

---

## Reporting

Files on the work computer, paths to Ross, **no git operations**.

**Read anything before you relay it.** These artifacts are not equally
shareable, and two of them describe the machine and its real work rather than
the product. That judgement is yours to make with the facts, and nothing here
makes it for you — see [PRIVACY.md](../PRIVACY.md).

1. The **digest** first. It carries versions, per-operation status and timing,
   the leak count and the result. No machine name, no add-in list, no paths.
2. The Markdown and JSON reports, which name the computer, the Office and Excel
   builds, the macro-trust values, and **every connected COM add-in by ProgID** —
   a fair description of the finance stack this shop runs. From 0.19.0 the
   Windows account name is written as `%USERPROFILE%`; a report from an earlier
   build has the account name in it, beside the computer name.
3. The trace log, if step 4 ran — and **only after reading it**. It carries no
   workbook contents, but it does carry workbook file names, worksheet names and
   A1 ranges, by design, because a trace that cannot be matched to its run is
   useless. If this run touched a real workbook, those are real names, and a
   workbook here is routinely named for a payer, a client, a facility or a deal.
   If that is a problem, say so and send the digest alone; the phase timings are
   what the trace is wanted for and they can be quoted without the file.
4. `ANALYTICS.md` from step 6.
5. One `SUMMARY.md` containing:
   - the version removed and the version installed, with the file count;
   - each command exactly as executed;
   - the digest verbatim and every exit code;
   - the environment table the check reports;
   - **the step 5 result, with the complete refusal text if it refused**;
   - test counts per assembly, or a note that the SDK or clone was absent;
   - the exact text of anything that failed — pasted, never summarized;
   - anything surprising the tooling did not capture.

If you ran out of time, report what you completed rather than skipping ahead.
Steps 3, 5 and 6 are the ones that matter; 7 and 8 are optional.
