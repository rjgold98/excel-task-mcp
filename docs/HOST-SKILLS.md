# What host-loaded skills do to this server

Whether always-on, auto-invoked skills help or hurt an agent driving ExcelTask, and
what to load instead. Written against ExcelTask 0.18.0 (commit `b82a19f`), with the
live `tools/list` measured rather than assumed.

Every claim below cites either a file in this repository (`path:line`) or a primary
document (spec text or first-party vendor documentation). Where a claim cannot be
traced to a primary source it is marked **unverified inference**.

> **Scope note.** This was written while the working tree was ahead of `b82a19f`, with a
> fifteenth operation, `ManageModelRelationship`, in progress. That operation has since
> shipped in **0.19.0**, together with a rewritten tool description. Counts below are
> stated for **committed 0.18.0 (fourteen operations)**; where a number moved materially
> — the schema byte budget in §5 — every figure is given with its version and its
> measurement method. Line numbers in modified files may drift.
>
> Nothing in the argument turns on fourteen versus fifteen: the tool count is one either
> way. The §5 verdict — that a routing sentence must be paid for by reclamation rather
> than appended — is only sharpened by 0.19.0, which spent 369 of the remaining bytes on
> the tool description itself. §5 records what that leaves.

---

## Answer

**Yes — the two skills currently loaded on the work computer interfere, and the
first one interferes badly.** Three findings carry the verdict.

1. **The "Excel Default Skill" names 19 tools. ExcelTask exposes exactly one, and
   it is none of them.** ExcelTask advertises a single tool, `excel_task`
   (`src/ExcelTask.McpServer/ExcelTaskTool.cs:20`), pinned by a test that asserts
   the list has exactly one entry (`tests/ExcelTask.McpServer.Tests/ExcelTaskToolProtocolTests.cs:77-78`).
   The skill's tool table (`get_workbook_info`, `list_sheets`, `get_range_values`,
   `set_range_values`, `set_number_format`, `create_table`, `create_chart`,
   `analyze_data`, and eleven more) has **zero** overlap. A skill that names
   non-existent tools is not neutral: MCP makes an unknown tool a *protocol* error
   (`-32602`), and the spec itself warns those are "less likely to result in
   successful recovery" than tool-level errors — the model gets a malformed-request
   signal rather than something it can correct
   ([MCP tools spec](https://modelcontextprotocol.io/specification/2026-07-28/server/tools)).

2. **The excelcli skill drives a competing automation stack against the same desktop
   Excel, and its session model is the exact hazard ExcelTask's central guarantee is
   built on.** Its own Rule 3 admits "Unclosed sessions leave Excel processes
   running, locking files." ExcelTask binds through the Running Object Table
   (`src/ExcelTask.Excel/RotWorkbookLocator.cs:28-65`) and proves its own Excel
   exited (`src/ExcelTask.Excel/OwnedExcelProcess.cs:69-84`). A stray excelcli Excel
   holding the target path silently changes what `UseOpen` binds to, converts
   `Isolated`+`Same` applies into outright rejections
   (`src/ExcelTask.Excel/ExcelWorkbookRuntime.cs:253-261`), and — if it appears
   mid-run — is counted against ExcelTask's leak figure, turning a clean run into a
   false `result=FAIL` (`src/ExcelTask.McpServer/FieldCheckFixtures.cs:51-52`).

3. **Its Rule 1, "NEVER Ask Clarifying Questions," directly contradicts the
   product's design.** ExcelTask has a `NeedsConfirmation` outcome
   (`src/ExcelTask.Core/Contracts.cs:8`) and a Plan/Apply split
   (`Contracts.cs:5`) whose entire purpose is to stop before mutation. A loaded
   instruction telling the model never to pause is a standing instruction to route
   around the safety property in a healthcare-finance environment.

**On the proposed fixes:** a `/scan` skill is the wrong shape — the sequencing
guidance already lives in the tool schema, it just points at the wrong operation,
and the schema is the layer with the proven effect. An `/excel-mcp`
skill is a counterweight to a conflict that should simply be removed. **The highest-value
action is deletion, not addition.**

---

## Findings

### 1. Tool-name collision: 19 named, 0 exist

**What ExcelTask actually exposes.** Measured, not assumed — a real
`initialize` + `tools/list` handshake over raw JSON-RPC against the 0.18.0 Release
build (the same method `docs/FIELD-CHECK.md:99-102` prescribes, "measured over raw
JSON-RPC rather than through a client library, so the bytes are the server's own"):

```
TOOL COUNT: 1
name=excel_task  bytes=20603 (20.12 KB)
  inputSchema  16,756 bytes
  outputSchema  3,359 bytes
```

One tool. Its fourteen operations are a discriminated union on
`operation.kind` (`src/ExcelTask.Core/Contracts.cs:9`), and the enum is the single
source of truth — `OperationCatalog` switches over it with no default arm precisely
so a forgotten member is a compile error (`src/ExcelTask.Core/OperationCatalog.cs:26-49`).

The exact fourteen, quoted from `Contracts.cs:9`:

> `CopyExhibit, RepairExistingWorksheet, ExtendFormulaSeries, EditMacroProcedure,
> AuditWorkbookFlows, ReadWorksheetRange, WriteWorksheetValues, FindReplace, Create,
> SetRangeFormat, ScanWorkbookStructure, ManageTable, ManageQuery, ManageModelMeasure`

Outcome values, quoted from `Contracts.cs:8`:

> `Planned, NeedsConfirmation, Completed, Rejected, Partial, Unknown`

Modes and bindings (`Contracts.cs:5-6`): `Plan | Apply`, and
`AskIfOpen | UseOpen | Isolated`.

**The collision.** The Excel Default Skill's "High-Level Tool Guidance" table names
19 distinct tools. Intersection with ExcelTask's surface: **0**.

| Skill names | Exists in ExcelTask? |
|---|---|
| `get_workbook_info`, `list_sheets`, `list_tables` | No |
| `get_used_range`, `get_range_values`, `get_range_formulas` | No |
| `set_range_values`, `set_range_formulas` | No |
| `set_number_format`, `format_range`, `auto_fit_columns` | No |
| `create_table`, `get_table_data`, `filter_table` | No |
| `create_chart`, `set_chart_type`, `set_chart_title` | No |
| `recalculate_workbook`, `analyze_data` | No |

**What actually happens — grounded, with the limit of the evidence stated.**

The model is given the tool list in its system prompt. Anthropic publishes the
literal template: "In this environment you have access to a set of tools you can use
to answer the user's question," followed by "Here are the functions available in
JSONSchema format:" and the definitions
([define-tools](https://platform.claude.com/docs/en/agents-and-tools/tool-use/define-tools)).
MCP is the discovery half: "To discover available tools, clients send a `tools/list`
request," and tools "are designed to be **model-controlled**"
([MCP tools spec](https://modelcontextprotocol.io/specification/2026-07-28/server/tools)).

So the skill text and the tool list are two different inputs that disagree. Three
documented consequences, in descending order of certainty:

- **An unknown tool name is a protocol error, not a recoverable tool error.** The MCP
  spec gives the exact shape: code `-32602`, `"Unknown tool: invalid_tool_name"`. The
  2026-07-28 revision adds the framing that matters here — protocol errors "indicate
  issues with the request structure itself that models are less likely to be able to
  fix," and clients "**MAY** provide protocol errors to language models, though these
  are less likely to result in successful recovery." On the Anthropic API side a name
  outside the tools array is a 400: `"Tool reference 'unknown_tool' not found in
  available tools"` ([tool-search-tool](https://platform.claude.com/docs/en/agents-and-tools/tool-use/tool-search-tool)).
- **Naming a tool that isn't resolvable is a known failure mode Anthropic warns about
  in exactly this context.** The Agent Skills best-practices page tells skill authors
  to "always use fully qualified tool names to avoid 'tool not found' errors," because
  "**Without the server prefix, Claude may fail to locate the tool**, especially when
  multiple MCP servers are available"
  ([best-practices](https://platform.claude.com/docs/en/agents-and-tools/agent-skills/best-practices)).
  That is the closest primary-source phrasing to the concern; it describes a
  *prefix* problem, and this skill's problem is worse — the bare names do not exist
  under any prefix.
- **Tool descriptions, not skill prose, are the documented lever for tool selection.**
  "Provide extremely detailed descriptions. This is by far the most important factor
  in tool performance" ([define-tools](https://platform.claude.com/docs/en/agents-and-tools/tool-use/define-tools)).

**Unverified inference:** that the model, faced with a skill naming absent tools,
*hallucinates a call and then falls back to `excel_task`* — versus abandoning the MCP
for shell tooling. No primary source states either behaviour. But this project has a
field observation pointing at the second: a prior session made **216 PowerShell calls
against 71 ExcelTask calls** (`docs/FIELD-TASK.md:267`), and the agent "tried
PowerShell, then Python, then compiling C#" (`FIELD-TASK.md:265-266`). The mechanism
is not proven; the outcome shape has been seen once on this machine.

**One thing the skill gets right, and it is already free.** Its Operating Loop
(Discover → Read → Execute → Verify → Summarize) describes what ExcelTask does
*inside a single call*: "One task owns inspect, plan, execute, verify, save, and
cleanup" (`AGENTS.md:11`). The skill is asking the model to hand-orchestrate a loop
the server already performs and proves.

---

### 2. Competing automation stack: the collision is real and located

**What the second skill is.** Its command groups —
`session, batch, service, calculationmode, chart, chartconfig, conditionalformat,
connection, datamodel, datamodelrelationship, namedrange, pivottable, pivottablecalc,
pivottablefield, powerquery, pythoninexcel, range, rangeedit, rangelink, rangeformat,
screenshot, sheet, worksheetstyle, slicer, table, tablecolumn, vba, window` —
map almost one-for-one onto the incumbent Excel MCP's surface as measured in this
repo: `file, range, worksheet, vba, namedrange, range_edit, connection, screenshot,
calculation_mode, range_format, table, datamodel, powerquery, pivottable,
conditionalformat, window, range_link, pivottable_field, datamodel_relationship,
chart, slicer, table_column, chart_config, pivottable_calc, worksheet_style`
(`docs/field-reports/2026-08-10-demand/REPORT.md:24-45`). 23 of 25 match exactly;
`file` is replaced by `session`/`batch`/`service`, `worksheet` by `sheet`, and
`pythoninexcel` is new.

**Unverified inference:** that `excelcli.exe` *is* the incumbent Excel MCP in CLI
form. The correspondence is very strong and the skill self-describes as a "GitHub
Copilot `excel-cli` plugin," but I have no access to its source. Treat it as "the
same capability surface," which is all the argument needs.

Why that matters: this repo has measured what routing work to that surface costs. Six
controlled sessions on the work computer, one model held constant, three per server
(`docs/field-reports/2026-08-10-comparison/CLIENT-SESSIONS.md:11-22`):

| Metric | ExcelTask | Original MCP | Difference |
|---|---|---|---|
| MCP calls | 6 | 37 | −83.8% |
| Total tokens | 515,575 | 1,971,923 | −73.9% |
| Prompt to task complete | 124.8s | 263.4s | −52.6% |
| Correct after reopening | 3/3 | 3/3 | — |

A skill that steers Excel work onto that stack does not add a capability; it spends
the measured advantage. (Caveat kept from the source: one run per workflow, and both
catalogs were globally registered, so the token figures measure orchestration, not
schema loading — `CLIENT-SESSIONS.md:53-58`.)

#### 2a. Two automation stacks against one desktop Excel

Microsoft's position on concurrent automation of Office is explicit and currently
supported (not archived):

> "Office applications are non-reentrant, STA-based applications that are designed to
> provide diverse but resource-intensive functionality for **a single client**."

> "This can limit the number of instances that can run concurrently and can lead to
> **race conditions** if the applications are configured in a multiclient environment."

> "If you plan to run more than one instance of any Office application, you should plan
> to isolate them at the virtual machine level"

— [Considerations for unattended automation of Office](https://learn.microsoft.com/en-us/office/client-developer/integration/considerations-unattended-automation-office-microsoft-365-for-unattended-rpa).
The older KB 257757 states the same, adding that Office "may exhibit unstable
behavior and/or deadlock" ([KB 257757](https://support.microsoft.com/topic/considerations-for-server-side-automation-of-office-48bcfe93-8a89-47f1-0bce-017433ad79e2)).

#### 2b. Instance reuse and the Running Object Table

Microsoft documents the ROT as "a globally accessible look-up table on each
workstation" ([IRunningObjectTable](https://learn.microsoft.com/en-us/windows/win32/api/objidl/nn-objidl-irunningobjecttable)).
The multiple-instance rule is documented only in a support/troubleshooting article:

> "If multiple instances of Microsoft Excel are running, GetObject attaches to the
> instance that is **launched first**."

> "If you then close the first instance, another call to GetObject attaches to the
> second instance that was launched, and so forth."

— [GetObject and CreateObject behavior](https://learn.microsoft.com/en-us/troubleshoot/office/office-suite-issues/getobject-createobject-behavior).
The same article documents Excel as **SingleUse** instancing, meaning each
`CreateObject("Excel.Application")` launches another Excel process, and gives
Microsoft's own recommendation: "In general, Microsoft recommends that you use a new
instance of an Office application instead of attaching to an instance that the user
may be using."

**Two caveats that must not be dropped.** That article's instancing table stops at
Office 2007 and lives in the archived `previous-versions` tree; Microsoft publishes no
current-era equivalent for Excel 2016/2019/365. And the reference API pages
([`Marshal.GetActiveObject`](https://learn.microsoft.com/en-us/dotnet/api/system.runtime.interopservices.marshal.getactiveobject),
Win32 `GetActiveObject`, `IRunningObjectTable`) are **silent** on which instance wins.

**Crucially, ExcelTask does not use `GetActiveObject`.** It enumerates the ROT itself
and matches on moniker display name *and* the bound workbook's `FullName`
(`RotWorkbookLocator.cs:39-58`, `:181-199`). That is stronger than ProgID binding —
it disambiguates by path, which is exactly the escape hatch Microsoft's article
recommends ("`GetObject("Book2").Application` … attaches successfully to that instance
even if it isn't the earliest instance that was launched").

**But it does not survive two instances holding the *same* path.** `Find` returns the
first enumerated match and stops (`RotWorkbookLocator.cs:39-58`). Microsoft does not
document an ordering guarantee for `IEnumMoniker`, so when a user's Excel and a leaked
excelcli Excel both hold the target workbook, **which one `UseOpen` binds to is
undefined**. That is the sharpest collision, and it is a correctness risk, not a
performance one: an edit could land in the automation process's copy rather than the
one the user is looking at. *(The undefined-ordering conclusion is mine — sound from
the absence of a documented guarantee, but Microsoft nowhere states it.)*

#### 2c. The four places a leaked excelcli session lands, in this codebase

| # | Site | Effect of a stray Excel holding the target path |
|---|---|---|
| 1 | `ExcelWorkbookRuntime.cs:75` — `RotWorkbookLocator.ContainsPath` sets `TargetIsOpen` | `AskIfOpen` returns `NeedsConfirmation` ("The exact target workbook is open", `:79`). True but misleading: the "open" instance is an invisible automation process, not the user. |
| 2 | `ExcelWorkbookRuntime.Session.cs:99-103` — `UseOpen` binds the first ROT match | Binds to an undefined one of the two instances (see 2b). |
| 3 | `ExcelWorkbookRuntime.cs:253-261` and `.Mutation.cs:116-123` | `Isolated` + `Same` Apply is **rejected**: "The target workbook is already open and cannot be safely overwritten in isolated mode." In-place saves stop working entirely. |
| 4 | `ExcelWorkbookRuntime.cs:307-319` — pre-mutation revalidation via `HasExternalWorkbookAtPath` (`RotWorkbookLocator.cs:108-164`, compared by `Application.Hwnd`) | Apply aborts after planning: "The exact target workbook was opened in another Excel application before mutation; no changes were made." |

#### 2d. The leak assertion, and how a stray session makes it lie

ExcelTask's central claim is that it strands nothing: it captures the process identity
of the Excel it created — id **plus** start time **plus** image path, so a recycled
pid cannot be mistaken for it (`OwnedExcelProcess.cs:105-146`) — then waits for exit
and only force-terminates its own (`OwnedExcelProcess.cs:69-84`). `AGENTS.md:29`
states the boundary: "Close only owned workbooks. Never quit or kill a user-owned
Excel process."

The field check reports the leak figure as a set difference:

```
FieldCheckFixtures.cs:51-52
    var leaked = SnapshotExcelProcesses()
        .Count(identity => !before.Contains(identity) && !harnessOwned.Contains(identity));
```

`before` is snapshotted once at run start (`FieldCheck.cs:106`) and again per
operation (`FieldCheck.cs:403`), and `leaked == 0` is a precondition of PASS
(`FieldCheck.cs:676`). So:

- An excelcli Excel already running **before** the check starts is in `before` and is
  correctly excluded.
- An excelcli Excel started **during** the run — which is exactly what happens if the
  agent follows the excelcli skill mid-check — is in neither set and is **counted as
  ExcelTask's leak**. The digest prints `result=FAIL`, and `docs/FIELD-TASK.md:203-204`
  tells the tester to treat any non-zero leak as real.

That is a false negative in the one measurement this project treats as its central
claim. `docs/FIELD-TASK.md:66-78` already tells the tester to clear stray Excel before
starting — that guard covers case one and not case two.

#### 2e. File locking

Microsoft documents the owner file — "temporary and holds the logon name of the person
who opens the document," beginning "with a tilde (~), followed by a dollar sign ($)"
— but **only for Word**
([locked for editing](https://support.microsoft.com/en-us/topic/-the-document-is-locked-for-editing-by-another-user-error-message-when-you-try-to-open-a-document-in-word-10b92aeb-2e23-25e0-9110-370af6edb638)).
The Excel-branded lock article covers only cloud co-authoring and never mentions `~$`
files. **Excel's identical `~$` behaviour is universally observed but, as far as I can
find, undocumented by Microsoft — treat the Excel case as unverified inference.**

What *is* documented for Excel is the automation consequence, via `Workbooks.Open`'s
`Notify` parameter: if `Notify` "is False or omitted, no notification is requested, and
any attempts to open an unavailable file will **fail**"
([Workbooks.Open](https://learn.microsoft.com/en-us/office/vba/api/excel.workbooks.open)).
A locked workbook makes a default open fail rather than silently degrade — which is why
ExcelTask waits for its own process to exit *and* for the file lock to release before
the verification reopen (`docs/EXCEL-TUNING.md:107-117`).

---

### 3. Rule conflicts

#### 3a. Rule 1 — "NEVER Ask Clarifying Questions"

The skill's text: *"Execute commands to discover the answer instead… You have commands
to answer your own questions. USE THEM."*

Against this product's design that is a direct contradiction, at three levels:

- **Contract.** `NeedsConfirmation` is a first-class outcome (`Contracts.cs:8`), and
  the server is required to "Return `Completed`, `Partial`, `Unknown`, `Rejected`, or
  `NeedsConfirmation` **truthfully**" (`AGENTS.md:25-26`).
- **Product decision, locked with the user.** "If the target workbook is already open,
  ask whether to use that exact live workbook or an isolated file instance"
  (`docs/DECISIONS.md`, decision 8); "Overwriting an existing file requires explicit
  authorization" (decision 11). `AGENTS.md:14-15` adds: "Never guess from
  `ActiveWorkbook`."
- **Mechanism.** `overwriteConfirmed` is a required act of authorization for any
  same-file Apply, and the schema says why: "saving in place overwrites the target
  workbook itself" (`ExcelTaskToolProtocolTests.cs:155`).

Note the asymmetry in what the rule costs. Rule 1 is defensible for *read-only
discovery* — "which sheet?" really is answerable by a listing call. It is not
defensible for the confirmation gates, which exist because the user, not the model,
owns the decision to overwrite. A skill cannot distinguish the two cases, and the rule
as written is unconditional and capitalised. In a healthcare-finance context, the
confirmation *is* the product: the mixed-server field report records the user naming
these three guardrails — refusing to save in place, requiring the SHA-256 of existing
code, and reopening to re-verify — as "the reason to prefer ExcelTask for any edit that
matters, against a one-call alternative with none of them"
(`docs/field-reports/2026-08-10-mixed-server-macro.md:33-37`).

There is also a plain interaction failure: ExcelTask *returns* `NeedsConfirmation` and
expects a resubmission. A model instructed never to ask has been told to treat that
receipt as an obstacle.

#### 3b. Rule 7 — calculation-mode tuning for bulk writes

The skill: *"For 10+ cells, set manual calc mode, write, recalculate once, restore
automatic."*

This repo measured it and concluded the opposite. `docs/EXCEL-TUNING.md` (measured
2026-08-10 with `tools/excel-calc-probe.ps1`, median of interleaved trials):

| Phase | Cost |
|---|---|
| Excel launch | 274–482 ms |
| Workbook open | 60–75 ms |
| Write 2,000 formulas (43 batched calls) | 58 ms |
| CalculateFull | ~1 ms |
| Save | 18 ms |

> "Launching Excel is the entire cost. Everything ExcelTask does *inside* Excel is
> rounding error beside starting it." (`EXCEL-TUNING.md:35-36`)

And the direct A/B on the setting itself (`EXCEL-TUNING.md:40-44`):

| Calculation during writes | Write | Calc | Restore | Save | Total |
|---|---|---|---|---|---|
| automatic (shipping) | 58 ms | 1 ms | 0 ms | 18 ms | **76 ms** |
| manual, restored before save | 83 ms | 1 ms | 1 ms | 18 ms | **103 ms** |

> "**Do not add `Calculation = xlCalculationManual`.**" (`EXCEL-TUNING.md:54`)

The reason is structural, not incidental: ExcelTask groups repairs by identical R1C1
formula and writes them as multi-area ranges, so 2,000 cells cost 43 calls and "there
is no per-cell recalculation left to suppress" (`:45-49`). The document explicitly
records that a per-cell loop *does* show the opposite result (586 ms vs 489 ms) "which
is exactly why it must not be used to justify the setting. That loop is not the code we
ship" (`:51-52`).

So Rule 7 costs **+27 ms of overhead against a ~2,600 ms floor** — under 1% — and buys
nothing. It also adds a failure mode the repo took seriously enough to test: a workbook
saved from an instance left in manual mode could hand the user back a model that
silently stops recalculating. Measured, that hazard did not materialise on the tested
build (`:57-64`) — but the risk is real for any tool that toggles the mode and dies
before restoring it, which is precisely what an unclosed session does.

Two further notes: the document's related conclusions cover `ScreenUpdating`,
`PrintCommunication` and `DisplayStatusBar` as "within noise of the baseline" (`:66-72`),
and it records an **invalid** PowerShell-based measurement deliberately, because
"PowerShell holds its own references to the COM object" so Excel could not exit — "that
measured the probe" (`:192-196`). This is the one repository whose measurement notes
already anticipate the class of advice Rule 7 represents.

Also stale, and worth noting because Rule 7 sits beside it: `docs/EXCEL-TUNING.md:3`
scopes all of this to "the development machine." The work computer has never
reproduced these numbers.

#### 3c. Rule 5 and Rule 7's PowerShell examples under Constrained Language Mode

The brief asked whether these would even run. **They would — the assumption they fail
is wrong, and I am recording the correction rather than the convenient answer.**

Microsoft's `about_Language_Modes` establishes:

> "All cmdlets in Windows modules are fully functional and have complete access to
> system resources, except as noted."

> "The ConstrainedLanguage mode permits all cmdlets and a subset of PowerShell language
> elements, but limits the **object types** that can be used."

— [about_Language_Modes](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_language_modes)

`ConvertFrom-Json` is a cmdlet, is not named among the restricted items (`New-Object`
and `Add-Type` are), and its documented outputs are `PSCustomObject` and
`OrderedHashtable` ([ConvertFrom-Json](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/convertfrom-json)).
`[pscustomobject]` and `[psobject]` are both on CLM's allowed-types list, and "Users can
get all properties of allowed types." So `excelcli … | ConvertFrom-Json` followed by
`.sessionId` **works** under CLM.

**Marked honestly: Microsoft nowhere writes "ConvertFrom-Json works in
ConstrainedLanguage mode."** This is a two-step inference from separately-quoted
documented facts. One line settles it empirically on the machine itself:
`$ExecutionContext.SessionState.LanguageMode`, then run the pipeline. One narrow
exception is worth carrying: `OrderedHashtable` (the PS 7.3+ `-AsHashtable` output) is
**not** on the 7.6 allowed list, so `-AsHashtable` on PowerShell 7 is the one form that
may fail.

CLM also does not block running a compiled `.exe`: the doc describes even the stricter
NoLanguage mode as one where "You can only run native commands and cmdlets." It is a
*language* restriction, not a process-execution restriction — process execution is
governed separately by AppLocker/WDAC, which is what triggers CLM in the first place.
So `excelcli.exe` itself runs fine.

**What CLM does block is the thing ExcelTask deliberately avoids.** CLM permits
exactly three COM ProgIDs — `Scripting.Dictionary`, `Scripting.FileSystemObject`,
`VBScript.RegExp` — so `New-Object -ComObject Excel.Application` fails, and "`Add-Type`
… can't load arbitrary C# code or Win32 APIs." This is why `--field-check` is a
compiled executable: "compiled rather than scripted on purpose: managed computers
commonly run PowerShell in Constrained Language Mode, which forbids the COM and
reflection a scripted equivalent needs" (`src/ExcelTask.McpServer/FieldCheck.cs:57-60`;
also `Program.cs:14-15`, `MakeFixture.cs:9-12`, `docs/FIELD-CHECK.md:7-9`). The server
already reports the policy value it finds (`FieldCheck.cs:515`).

**Net for this question:** CLM is *not* the reason to reject the excelcli skill. The
reasons are the leaked-session collision (§2) and the confirmation conflict (§3a). Any
recommendation that leaned on "CLM will block it anyway" would be wrong, and would also
be routing around employer policy rather than respecting it.

---

### 4. How skills load, per host — and the two hosts differ

#### 4a. Claude Code / Anthropic Agent Skills

Progressive disclosure is documented as a table with explicit costs
([Agent Skills overview](https://platform.claude.com/docs/en/agents-and-tools/agent-skills/overview)):

| Level | When loaded | Token cost | Content |
|---|---|---|---|
| 1 — Metadata | "Always (at startup)" | **"~100 tokens per Skill"** | `name` and `description` from frontmatter |
| 2 — Instructions | "When Skill is triggered" | **"Under 5k tokens"** | SKILL.md body |
| 3+ — Resources | "As needed" | **"None until accessed"** | Bundled files; scripts' output only |

> "Claude loads this metadata at startup and includes it in the system prompt."
> "until a Skill is triggered, only its name and description occupy context."
> "When you request something that matches a Skill's description, Claude reads SKILL.md
> from the filesystem using bash. Only then does this content enter the context window."
> "Claude accesses these files only when referenced."

Frontmatter: `name` and `description` required; `name` max 64 characters, `description`
max 1,024. Triggering is model-decided on the description — "The `description` is what
Claude matches your request against when determining whether to trigger the Skill" —
and the description "is injected into the system prompt."

**There is no documented always-on/deterministic skill mode.** The docs say so
plainly: if a skill stops influencing behaviour, "the model is choosing other tools or
approaches. Strengthen the skill's `description` and instructions … or **use hooks to
enforce behavior deterministically**." Hooks are the documented escape hatch, not skills.

In Claude Code specifically ([Claude Code skills](https://code.claude.com/docs/en/skills)):
explicit invocation exists — "Claude uses skills when relevant, or you can invoke one
directly with `/skill-name`" — and the invocation matrix is documented:

| Frontmatter | You invoke | Claude invokes | Loading |
|---|---|---|---|
| (default) | Yes | Yes | "Description always in context, full skill loads when invoked" |
| `disable-model-invocation: true` | Yes | No | **"Description not in context**, full skill loads when you invoke" |
| `user-invocable: false` | No | Yes | "Description always in context, full skill loads when invoked" |

Locations: `~/.claude/skills/<name>/SKILL.md` (personal), `.claude/skills/` (project),
plugin skills namespaced `plugin:skill`. Disabling: `skillOverrides` in settings with
states `"on" | "name-only" | "user-invocable-only" | "off"`; or denying the `Skill` tool
in `/permissions`; or `disable-model-invocation: true`, which "**removes the skill from
Claude's context entirely.**" Note: "**Plugin skills are not affected by
`skillOverrides`. Manage those through `/plugin` instead.**"

Claude Code also deviates on persistence: the rendered SKILL.md "enters the conversation
as a single message and stays there for the rest of the session"; "Claude Code does not
re-read the skill file on later turns."

**Portability trap, relevant if a skill is ever written here and used in both hosts:**
outside Claude Code only six frontmatter fields are legal — `name`, `description`,
`license`, `compatibility`, `metadata`, `allowed-tools`. Anything else is a hard error:
"Unexpected key(s) in SKILL.md frontmatter: argument-hint."

#### 4b. GitHub Copilot — and this is where Ross actually is

GitHub now documents seven customization mechanisms with different loading semantics.
The [customization cheat sheet](https://docs.github.com/en/copilot/reference/customization-cheat-sheet)
tabulates the trigger column directly:

| Feature | How to trigger (verbatim) | Location |
|---|---|---|
| Custom instructions | **"Automatic"** | `.github/copilot-instructions.md`, `.github/instructions/*.instructions.md`, `AGENTS.md`, personal/org UI |
| Prompt files | "Manual: reference directly in chat or use the prompt file picker" | `.github/prompts/*.prompt.md` |
| Custom agents | "Manual: select from the agent dropdown" | `.github/agents/NAME.md` |
| **Agent skills** | **"Automatic: chosen by Copilot when relevant to your prompt"** | `.github/skills/<name>/SKILL.md`, `.claude/skills/…`, `.agents/skills/…`; `~/.copilot/skills/…`, `~/.agents/skills/…` |
| Hooks | "Automatic: at configured lifecycle events" | `.github/hooks/*.json` |
| MCP servers | "Automatic, or ask for a specific tool by name" | `mcp.json` |

Same page: custom instructions are "Always-on context that automatically applies to
every interaction within its defined scope."

**The decisive fact: Copilot's agent skills are the same open standard.**

> "Agent skills are folders of instructions, scripts, and resources that Copilot can
> load when relevant to improve its performance in specialized tasks. **The Agent Skills
> specification is an open standard**, used by a range of different AI systems."

> "Agent skills work with Copilot cloud agent, Copilot code review, the Copilot CLI,
> **the GitHub Copilot app**, and agent mode in Visual Studio Code and JetBrains IDEs."

— [About agent skills](https://docs.github.com/en/copilot/concepts/agents/about-agent-skills)

Triggering is model-decided on the description, exactly as with Anthropic: "When
performing tasks, Copilot will decide when to use your skills based on your prompt and
the skill's description," and "When a skill is invoked, Copilot automatically discovers
all of the files in the skill's directory and makes them available alongside the skill's
instructions"
([add-skills](https://docs.github.com/en/copilot/how-tos/copilot-cli/customize-copilot/add-skills)).

**Explicit invocation exists**, and answers the "would a `/`-invoked skill work?"
question for the Copilot side:

> "To tell Copilot to use a specific skill, include the skill name in your prompt,
> preceded by a forward slash. For example … `Use the /frontend-design skill to create a
> responsive navigation bar in React.`"

**Disabling exists, and needs no admin rights:**

> "use the command `/skills` and then use the up and down keys on your keyboard, and the
> space bar, to toggle skills on or off"

`/skills` is confirmed in the desktop app's own command list, along with `/skills reload`
([Copilot app slash commands](https://docs.github.com/en/copilot/reference/github-copilot-app-reference/slash-commands)).
For plugin-supplied skills: "To remove skills added as part of a plugin you must manage
the plugin itself" — which is the likely route for the excelcli skill, since it
self-describes as a "GitHub Copilot `excel-cli` plugin."

**How they got there without Ross installing them:** the desktop app inherits
configuration. "Any MCP servers configured for your repositories or Copilot CLI are
automatically available in the GitHub Copilot app," with the same sentence pattern for
skills ([customize the Copilot app](https://docs.github.com/en/copilot/how-tos/github-copilot-app/customize-github-copilot-app)).
That is consistent with Ross's account of "just letting it handle" the `.mcp` config, and
with `docs/FIELD-CHECK.md:136-142`, which notes that "clients that sync a repository's
MCP configuration will pick it up" from `.vscode/mcp.json`.

**Precedence, and a documented contradiction worth knowing.** The global order is
"Personal instructions, Repository custom instructions (Path-specific, Repository-wide,
Agent instructions), Organization custom instructions"
([response customization](https://docs.github.com/en/copilot/concepts/prompting/response-customization)),
with the important qualifier: "However, **all sets of relevant instructions are provided
to Copilot**." Precedence is a hint to the model, not a filter — nothing is dropped.
GitHub's own advice is exactly the recommendation below: "Whenever possible, try to avoid
providing conflicting sets of instructions." But the Copilot CLI page states it
"**does not define a general precedence order between these files**," which contradicts
the global list for that surface. Do not rely on precedence to resolve a conflict.

#### 4c. Where the hosts differ (this matters if a skill is written)

| | Claude Code | GitHub Copilot |
|---|---|---|
| Skill format | SKILL.md, same open standard | SKILL.md, same open standard |
| Auto-trigger | Model-decided on description | Model-decided on description |
| Always-in-context cost | "~100 tokens per Skill" (documented) | **Not documented** |
| Body-load size guidance | "Under 5k tokens" / "under 500 lines" | **Not documented** for SKILL.md |
| Explicit invocation | `/skill-name` | Skill name after `/` in the prompt |
| Disable | `skillOverrides`, `/permissions` deny, frontmatter | `/skills` toggle; plugin management for plugin skills |
| Personal path | `~/.claude/skills/` | `~/.copilot/skills/`, `~/.agents/skills/` |
| Project path | `.claude/skills/` | `.github/skills/`, `.claude/skills/`, `.agents/skills/` |
| Extra frontmatter | ~20 fields accepted | Only `name`, `description`, `license` documented |
| `/`-invoked prompt files | Commands and skills unified | `.github/prompts/*.prompt.md` — **VS Code / VS / JetBrains only**, not the app or CLI |

Two consequences. First, **`.claude/skills/` is read by both hosts** — so a skill added
to this repository would load in Copilot on the work machine too, which is either
convenient or a hazard depending on intent. Second, **prompt files are not available in
the Copilot desktop app**, so "write a `/scan` prompt file" is not an option there;
a skill invoked by name is.

---

### 5. The `/scan` idea — right instinct, wrong layer, and the schema is already 90% there

`ScanWorkbookStructure` reads the package directly with no Excel process. That is real:
`ExcelWorkbookRuntime.Scan.cs` imports `System.IO.Compression` (`:2`) and the operation
is documented as mapping structure "by reading the file as what it physically is - a ZIP
of XML - with no Excel process at any point" (`Scan.cs:17-18`). Microsoft's OPC
documentation supports the premise: OPC "provides a mapping of package concepts to the
ZIP archive format," and names `.xlsx` explicitly
([OPC fundamentals](https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview));
Microsoft's own unattended-automation guidance goes further — "Direct editing of the
file formats is the recommended and supported method for handling changes to Office
files from a service."

The cost argument is quantified in-repo: a session is ~2.6 s and "cost tracks
**sessions**, not work" (`EXCEL-TUNING.md:141-150`); a scan pays none of it, and
`Contracts.cs:163-166` records that "the direct read answered in under a third of the
audit's time with zero Excel launches."

**So should a skill mandate "scan first"? No — because sequencing guidance is already in
the schema, and the schema is the documented lever.** Dumping the live schema and
extracting every description containing sequencing language returns 15 hits. Five say
some form of *"Run AuditWorkbookFlows first if unknown"* — on `worksheetName` for read,
write, find/replace and format, and on `componentName` for macro editing. A sixth routes
Data Model work through `AuditWorkbookFlows` and `ManageQuery`.

**Zero point at `ScanWorkbookStructure` as a first step.** Its own description routes
*outward* — "absence here is not evidence, so use AuditWorkbookFlows when those matter"
(`ExcelTaskToolProtocolTests.cs:199`) — and nothing routes *inward*.

That is the actual defect, and it is one sentence of description text, not a skill. Today
a model asked to repair formulas is told five times to run `AuditWorkbookFlows`, which
launches Excel and costs a full session, when for pure structure the free operation would
do. The repo already knows why this class of fix works: the A/B study's standing
conclusion is "*if the server rejects on it, the description must say so*," proven across
two models at **p = 0.0012** for the caps fix, eliminating every remaining hard-suite
failure for +18% schema bytes (`docs/INTERFACE-AB-STUDY.md`). It also records that
**operation selection was perfect in 484/484 decisions across every variant, every round
and both models, including the terse ones** — models are not failing to pick the right
operation; they fail on policy and sequencing, which is exactly what description text
fixes and skill prose does not.

**Is there room? Yes, but less than it looks, and the margin is closing as this is
written.**

| | Bytes | Headroom |
|---|---|---|
| Budget (`ExcelTaskToolProtocolTests.cs:102`, `22 * 1024`) | 22,528 | — |
| Committed 0.18.0, the repo's own figure (`ROADMAP.md:219`) | 20,694 | 1,834 (8.1%) |
| Committed 0.18.0, measured here over raw JSON-RPC | 20,603 | 1,925 (8.5%) |
| Working tree, with the fifteenth operation | 21,618 | 910 (4.0%) |
| **Shipped 0.19.0, the pin's own measure** | **22,214** | **314 (1.4%)** |
| Shipped 0.19.0, raw JSON-RPC (field check) | 22,171 | 357 (1.6%) |

The 91-byte gap between the two 0.18.0 figures is method, not disagreement: the pin
measures `JsonSerializer.SerializeToUtf8Bytes(tool)` in C#, this measured the wire JSON
of the same object. Either way the fifteenth operation consumed roughly half the
remaining slack, so **the honest reading is that a routing sentence must now be paid for
by reclamation, not appended.** That is the pin working as designed rather than failing.

**What happened next is the direct test of that reading, and it went the other way.**
0.19.0 spent 369 more bytes — not on an operation, but on the tool's own one-line
description, which still named three kinds of work while fifteen shipped and had been
measured sending ordinary formula work to a different Excel tool. By this section's own
standard that qualifies: it is "a rule the caller cannot act without", since a caller
that learns the formula refusal from a rejection has already paid a round trip. It was
still appended rather than reclaimed, and the margin is now **314 bytes**. The next
sentence, routing or otherwise, has to buy its space.

The pin's own comment sets the standard a new
sentence must meet: it "must not grow to make room for wordier prose - only for a new
operation, or for a rule the caller cannot act without"
(`ExcelTaskToolProtocolTests.cs:86-89`). A routing clause that saves an entire Excel
session on the most common discovery path is arguably such a rule, but it is a judgement
call, and the comment also warns "Reclaim before raising."

**Verdict on `/scan`:** not worth a skill. The same effect, more reliably and at zero
per-session token cost, comes from ~100 bytes in `ScanWorkbookStructureOperation`'s
description. A skill costs ~100 tokens of always-on context per session *in Claude Code*
(undocumented in Copilot), fires only when the model decides its description matches, and
is not portable to the field-test machine's guarantees.

---

### 6. The `/excel-mcp` idea — a counterweight to a conflict that should be deleted

If ExcelTask were the only Excel tool present, a skill saying "route Excel work through
`excel_task`" would be pure redundancy: MCP already tells the model what exists via
`tools/list`, and Anthropic documents that "Claude determines when to call a tool based
on the user's request and the tool's description."

It is not the only tool present. Both catalogs were globally registered during the
comparison run (`CLIENT-SESSIONS.md:53-55`), the incumbent advertises 25 tools / 230
operations / 57,641 bytes (`docs/field-reports/2026-08-10-demand/REPORT.md:12-13`), and
the excelcli skill actively instructs the model to use a *third* path — with the closing
line "Do not use MCP call syntax, snake_case parameters, or underscore tool names," which
is a direct instruction not to call `excel_task`.

So the honest framing is: **`/excel-mcp` would be a counterweight, and counterweights are
the weaker move.**

- It is not deterministic. Anthropic's docs say so outright — when a skill stops
  influencing behaviour, "the model is choosing other tools or approaches," and hooks,
  not skills, are the deterministic mechanism. Adding instruction volume to beat other
  instruction volume is a coin-flip.
- GitHub's own guidance is to remove the conflict: "Whenever possible, try to avoid
  providing conflicting sets of instructions." And precedence will not save you —
  "all sets of relevant instructions are provided to Copilot," and the CLI surface
  "does not define a general precedence order between these files."
- It would pollute the very measurement that matters. `docs/FIELD-TASK.md:19-21` states
  the second-most-valuable question of the whole field run: "**What does the agent reach
  for instead of ExcelTask?** … Knowing which receipt sent it away is worth more than any
  feature on the roadmap." A skill that forces routing destroys that signal. The right
  time for a routing skill is *after* a clean run has shown what the agent does unaided.

**Verdict:** do not add `/excel-mcp` now. Remove the competing skill, run the field task
clean, and only then decide whether routing needs reinforcement. If it turns out it does,
the finding belongs in the tool description or an `AGENTS.md`-style always-on instruction —
both of which apply unconditionally — rather than in a skill that fires on a description match.

One genuinely useful variant does exist, and it is not a routing skill: a **short
always-on instruction** stating the two facts a model cannot discover — that `excel_task`
is the only Excel tool that verifies its own work, and that `NeedsConfirmation` is a
normal, expected receipt to be relayed to the user rather than worked around. In Copilot
that belongs in repository custom instructions or `AGENTS.md`, which are documented as
"Automatic" and always-on, not in a skill.

---

## Recommendation

Prioritised. "Where" names the layer the change belongs in, which is the load-bearing
part of each row.

| # | Action | Where | Why | Effort / risk |
|---|---|---|---|---|
| **1** | **Disable the "Excel Default Skill."** In the Copilot app, run `/skills`, arrow to it, press space to toggle off. | Copilot client, no admin rights | 19 named tools, 0 exist (§1). It cannot succeed and can only misdirect. Highest damage, lowest cost to remove. | Minutes. None. |
| **2** | **Disable the "Excel Automation with excelcli" skill**, same `/skills` toggle. If it is plugin-supplied, manage the plugin — "To remove skills added as part of a plugin you must manage the plugin itself." | Copilot client | Rule 1 contradicts the confirmation design (§3a); Rule 7 contradicts measured tuning (§3b); a leaked session corrupts `UseOpen` binding, blocks in-place Apply, and falsifies the leak assertion (§2c–2d). | Minutes. **Check with Ross first** — `docs/AGENT-BRIDGE.md:88-89` says never to interfere with the original Excel MCP install. Disabling a *client-side skill* is not that, but the boundary is close enough that the owner should make the call, not a field agent. |
| **3** | **Do not add `/scan` or `/excel-mcp` skills.** | — | §5 and §6. Both are counterweights to a problem that step 1–2 removes; the second also destroys the field task's most valuable signal. | None. |
| **4** | **Add a "Before you start" section to `docs/FIELD-TASK.md`**: list the enabled skills and paste that list into the report; confirm no Excel-related skill is enabled; state that no other Excel automation may run during the check. | Repo docs | §2d — an excelcli Excel started *mid-run* is counted as ExcelTask's leak and prints `result=FAIL`. Step 1a only clears strays *before* the run (`docs/FIELD-TASK.md:66-78`). | Small. **Highest-value doc change here.** |
| **5** | **Extend `docs/FIELD-TASK.md` step 6** to require the skill *inventory*, not just invocations. It already asks for "Every skill invocation: name, timestamp, duration" (`:282`) — add: every skill *loaded and enabled*, whether invoked or not. | Repo docs | A skill that shaped behaviour without firing is invisible to the current export, and would silently confound the "what did the agent reach for instead" analysis (`:19-21`). | Small. |
| **6** | **Add one sentence to `ScanWorkbookStructureOperation`'s description** routing structural discovery to the scan before the audit. | **MCP schema** — `src/ExcelTask.Core/Contracts.cs` | Five descriptions already say "Run AuditWorkbookFlows first"; none mention the scan (§5). Fixes sequencing at the layer with proven effect (p = 0.0012) and zero per-session token cost. | Small, but **no longer free**: the fifteenth operation cut headroom to ~910 bytes, so this must be paid for by reclamation. Must pass `ExcelTaskToolProtocolTests`. |
| **7** | **Correct the CLM assumption wherever it is relied on.** `ConvertFrom-Json` pipelines *do* run under CLM (§3c). | Repo docs / this file | Rejecting the excelcli skill "because CLM blocks it" would be wrong and would look like routing around employer policy. The real reasons are §2 and §3a. | Done here; worth a line in `docs/FIELD-CHECK.md` if that claim ever hardens. |
| **8** | **Consider a short always-on instruction** (repository custom instructions or `AGENTS.md` — documented "Automatic"), stating that `excel_task` verifies its own work and that `NeedsConfirmation` is expected and must be relayed, not bypassed. **Only after a clean field run.** | Copilot instructions, *not* a skill | Instructions are unconditional; skills fire on a description match (§4b, §6). But adding it before a clean run pollutes the measurement `docs/FIELD-TASK.md:19-21` exists to take. | Small; deliberately deferred. |
| **9** | **Re-run `--field-check` with skills off and compare** the digest and `leaked=` figure against the last run. | Work computer | Turns "skills interfere" from an argument into a measurement, using the harness already built for it. | ~3 min. |

**Explicitly not recommended:** editing the two skill files directly, or anything
touching the incumbent MCP install. Disabling via `/skills` is reversible, needs no
admin rights, and stays inside `docs/AGENT-BRIDGE.md`'s boundaries.

### Where the evidence is thin

Recorded rather than dressed up.

1. **What a model does with a skill naming absent tools** — hallucinate, fall back, or
   abandon the MCP — is **not documented by anyone**. The error *shapes* are documented
   (§1); the behavioural response is not. The 216-vs-71 PowerShell observation
   (`FIELD-TASK.md:267`) is one data point on one machine with no controlled comparison.
2. **`excelcli.exe` is the incumbent Excel MCP in CLI form** — strong correspondence
   (23 of 25 command groups), not proof.
3. **ROT enumeration order** — Microsoft documents "first launched" for
   `GetObject`/`GetActiveObject`, but ExcelTask enumerates the ROT directly. No ordering
   guarantee is documented for `IEnumMoniker`, so "undefined which instance wins" is my
   inference from an absence.
4. **Excel's `~$` owner file** is documented by Microsoft **for Word only**. The Excel
   case is universally observed and undocumented.
5. **Excel's SingleUse instancing** is documented only in an archived article covering
   Office 97–2007. No current-era Microsoft statement exists for Excel 2016/2019/365.
6. **`ConvertFrom-Json` under CLM** is a sound two-step inference, not a Microsoft
   statement. One command on the machine settles it.
7. **Copilot documents no token cost or size limit for SKILL.md**, so the "~100 tokens
   per skill" figure is Anthropic's and must not be quoted as Copilot's.
8. **Copilot's own precedence docs contradict each other** (global order vs. the CLI's
   "does not define a general precedence order"). Do not rely on precedence.
9. **`docs/EXCEL-TUNING.md:3-5`** scopes all timing to the development machine and says
   "Reproduce before trusting." None of it has been reproduced on the work computer.

---

## Sources

**This repository** (0.18.0, `b82a19f`)

- `src/ExcelTask.Core/Contracts.cs` — `:5-9` modes, bindings, statuses, the 14 operation kinds; `:163-177` scan rationale; `:246-270` request/operation descriptions
- `src/ExcelTask.Core/OperationCatalog.cs` — `:26-49` no-default-arm switch over all kinds
- `src/ExcelTask.McpServer/ExcelTaskTool.cs` — `:19-25` single tool registration and description
- `src/ExcelTask.McpServer/FieldCheck.cs` — `:106`, `:403` leak snapshots; `:515` PowerShell lockdown probe; `:676` PASS condition; `:57-60` why compiled
- `src/ExcelTask.McpServer/FieldCheckFixtures.cs` — `:43-56` leak arithmetic
- `src/ExcelTask.McpServer/Program.cs:14-15`, `MakeFixture.cs:9-12` — CLM rationale
- `src/ExcelTask.Excel/RotWorkbookLocator.cs` — `:28-65` first-match ROT find; `:108-164` external-application detection; `:181-199` path matching
- `src/ExcelTask.Excel/OwnedExcelProcess.cs` — `:21-40` snapshot; `:69-84` proof of exit; `:105-146` process identity
- `src/ExcelTask.Excel/ExcelWorkbookRuntime.cs` — `:75-80` ROT open check; `:253-261` isolated-target rejection; `:307-319` pre-mutation revalidation
- `src/ExcelTask.Excel/ExcelWorkbookRuntime.Session.cs:99-103` — `UseOpen` binding
- `src/ExcelTask.Excel/ExcelWorkbookRuntime.Mutation.cs:116-123` — same rejection, mutation path
- `src/ExcelTask.Excel/ExcelWorkbookRuntime.Scan.cs:1-30` — ZIP-of-XML scan
- `tests/ExcelTask.McpServer.Tests/ExcelTaskToolProtocolTests.cs` — `:77-78` exactly one tool; `:79-102` schema byte budget and its history; `:143-234` every pinned description
- `AGENTS.md` — `:9-16` product contract; `:25-31` runtime invariants
- `docs/DECISIONS.md` — decisions 8, 10, 11
- `docs/EXCEL-TUNING.md` — `:3-5` scope; `:27-36` where time goes; `:40-64` calculation mode; `:66-72` no-op settings; `:107-117` verification reopen; `:137-163` session cost; `:192-196` the invalid PowerShell measurement
- `docs/INTERFACE-AB-STUDY.md` — 484/484 operation selection; p = 0.0012 for documented caps; "if the server rejects on it, the description must say so"
- `docs/FIELD-CHECK.md` — `:7-9` CLM; `:99-102` how tool-surface bytes are measured; `:136-142` MCP config sync
- `docs/FIELD-TASK.md` — `:19-22` the field run's questions; `:66-78` clearing stray Excel; `:267-271` 216-vs-71; `:282` skill invocations in analytics
- `docs/AGENT-BRIDGE.md:82-95` — boundaries binding both agents
- `docs/field-reports/2026-08-10-demand/REPORT.md` — incumbent surface and measured demand
- `docs/field-reports/2026-08-10-comparison/CLIENT-SESSIONS.md` — six-session token/latency comparison and its caveats
- `docs/field-reports/2026-08-10-mixed-server-macro.md` — the user's own verdict on the guardrails
- `ROADMAP.md:219` (current schema-budget figure, 20,694 of 22,528 at 0.18.0), `CHANGELOG.md:74-79` (the 16→22 KB rise)

**Measured for this document** — `initialize` + `tools/list` over raw JSON-RPC against
the Release build: 1 tool, `excel_task` — **20,603 bytes** at committed 0.18.0
(inputSchema 16,756, outputSchema 3,359), and **21,618 bytes** in the working tree once
`ManageModelRelationship` landed (inputSchema 17,771). `dotnet test` on
`ExcelTask.McpServer.Tests`: 18 passed, 0 failed.

**Anthropic**

- [Agent Skills overview](https://platform.claude.com/docs/en/agents-and-tools/agent-skills/overview) — progressive-disclosure table, ~100 tokens/skill, frontmatter limits, description-based triggering
- [Agent Skills best practices](https://platform.claude.com/docs/en/agents-and-tools/agent-skills/best-practices) — 500-line guidance; fully-qualified MCP tool names
- [Claude Code skills](https://code.claude.com/docs/en/skills) — invocation matrix, `skillOverrides`, locations, session persistence
- [Tool use: define tools](https://platform.claude.com/docs/en/agents-and-tools/tool-use/define-tools) — system-prompt template; "by far the most important factor"
- [Tool search tool](https://platform.claude.com/docs/en/agents-and-tools/tool-use/tool-search-tool) — unknown-tool 400; 30–50 tool degradation

**Model Context Protocol**

- [Server tools specification, 2026-07-28](https://modelcontextprotocol.io/specification/2026-07-28/server/tools) — `tools/list`, model-controlled tools, `-32602` unknown tool, protocol-vs-tool error framing

**GitHub**

- [Customization cheat sheet](https://docs.github.com/en/copilot/reference/customization-cheat-sheet) — trigger semantics for all seven mechanisms
- [About agent skills](https://docs.github.com/en/copilot/concepts/agents/about-agent-skills) — open standard; supported surfaces incl. the desktop app
- [Adding agent skills (CLI)](https://docs.github.com/en/copilot/how-tos/copilot-cli/customize-copilot/add-skills) — locations, `/skills` toggle, slash invocation, plugin-supplied skills
- [Copilot app slash commands](https://docs.github.com/en/copilot/reference/github-copilot-app-reference/slash-commands) — `/skills`, `/skills reload`
- [Customize the GitHub Copilot app](https://docs.github.com/en/copilot/how-tos/github-copilot-app/customize-github-copilot-app) — inherited MCP servers and skills
- [Repository custom instructions](https://docs.github.com/en/copilot/how-tos/configure-custom-instructions/add-repository-instructions) — "automatically added to requests"; `AGENTS.md`
- [Response customization](https://docs.github.com/en/copilot/concepts/prompting/response-customization) — precedence order; "all sets … are provided"
- [Copilot CLI custom instructions](https://docs.github.com/en/copilot/how-tos/copilot-cli/customize-copilot/add-custom-instructions) — "does not define a general precedence order"
- [Prompt files](https://docs.github.com/en/copilot/tutorials/customization-library/prompt-files/your-first-prompt-file) — `.github/prompts/*.prompt.md`, slash invocation, IDE-only availability

**Microsoft**

- [Considerations for unattended automation of Office](https://learn.microsoft.com/en-us/office/client-developer/integration/considerations-unattended-automation-office-microsoft-365-for-unattended-rpa) — non-reentrant, single client, race conditions; direct file-format editing recommended
- [KB 257757 — server-side Automation of Office](https://support.microsoft.com/topic/considerations-for-server-side-automation-of-office-48bcfe93-8a89-47f1-0bce-017433ad79e2) — unstable behaviour and deadlock
- [GetObject and CreateObject behavior](https://learn.microsoft.com/en-us/troubleshoot/office/office-suite-issues/getobject-createobject-behavior) — first-launched attachment; Excel SingleUse *(archived, Office 97–2007)*
- [IRunningObjectTable](https://learn.microsoft.com/en-us/windows/win32/api/objidl/nn-objidl-irunningobjecttable), [Marshal.GetActiveObject](https://learn.microsoft.com/en-us/dotnet/api/system.runtime.interopservices.marshal.getactiveobject) — ROT definition; both silent on multi-instance
- [Workbooks.Open](https://learn.microsoft.com/en-us/office/vba/api/excel.workbooks.open) — `Notify` and open-failure behaviour on a locked file
- [Document locked for editing](https://support.microsoft.com/en-us/topic/-the-document-is-locked-for-editing-by-another-user-error-message-when-you-try-to-open-a-document-in-word-10b92aeb-2e23-25e0-9110-370af6edb638) — owner file / `~$` *(Word only)*
- [about_Language_Modes](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_language_modes) — CLM allowed types, three permitted COM ProgIDs, `Add-Type` limits, native commands permitted
- [ConvertFrom-Json](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/convertfrom-json) — documented output types
- [Open Packaging Conventions fundamentals](https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview) — `.xlsx` as a ZIP package
