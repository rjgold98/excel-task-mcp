# Privacy

ExcelTask automates desktop Excel on the machine it runs on. This document says exactly
what leaves that machine, and by which route.

It exists because a shorter sentence kept getting written instead — "ExcelTask sends no
private data" — and that sentence is not true. There are four separate channels here with
four different contracts, and collapsing them into one claim makes the strong parts
unbelievable along with the loose one.

| Channel | Leaves the machine? | Carries workbook contents? |
|---|---|---|
| Telemetry | Nothing. There is none. | — |
| Diagnostic trace file | Only if you send it | No contents; yes, workbook and worksheet **names** |
| Ordinary receipts | To your MCP client, and so to the model | Only where the content is what you asked for |
| The request you send | To your MCP client, and so to the model | Whatever you put in it |
| Field-check report | Only if you send it | No contents; identifies the machine |

## 1. Telemetry: there is none

There is no telemetry provider, no analytics SDK, no crash reporter, no model SDK, and no
external formatting service. The server does not make network calls in normal operation.

`Program.cs` calls `builder.Logging.ClearProviders()` before anything is registered, so the
standard .NET logging providers — console, debug, event log — are removed rather than left
at their defaults.

This is the one place a flat claim is safe, and it is the only place this document makes
one.

## 2. The diagnostic trace file

Off unless the `EXCELTASK_TRACE` environment variable names a file. It is a development
aid; clear the variable and the feature is inert.

When on, it is written to the path you named, on your machine, and **appends until you
delete it**. Nothing rotates or expires it.

**It records:** phase names and durations; operation kind, mode, workbook binding, save
mode, and the overwrite flag; worksheet names; A1 ranges; workbook **file** names, never
their directories; owned Excel process ids; the final status and every check with its
detail; and the type and message of an unhandled worker failure.

**It never records:** cell values, formulas, or any workbook contents; VBA source;
connection strings, server names, or query text; full paths, user names, or machine names.

**What that means in practice.** A trace of real work contains lines like
`sheet=Control range=I24:K33` against a workbook file name. Where the workbooks are real,
those names are real: a workbook can be named for a payer, a client, a facility, or a deal,
and a worksheet name can be just as telling. The file cannot know whether that is
acceptable where you are about to send it. Up to and including v0.18.0 its header said
"safe to share" anyway. It now states its contents and leaves the conclusion to you.

## 3. Ordinary receipts

A receipt goes back through MCP to your client, which means the model sees it. Receipts are
bounded and deliberately withhold workbook contents — with the exceptions below, which are
exceptions by design and not by oversight.

**`WriteWorksheetValues` returns what it overwrote.** Its `prior-values` check names each
changed cell and its former value — `A4: 469750.25 -> 471000` — and its `formulas-replaced`
check names any cell whose formula the write destroyed, including that formula's text.
Both are capped at ten cells and forty characters each.

This is intentional: you named those cells in the request, so what was in them is the
answer to "what did I just change", not incidental spill. But it is workbook content
crossing the boundary during an operation that is not a read, so it is stated here rather
than left to be discovered.

**Checks describe structure.** `current-table` reports a table's address and style;
`current-format` reports the number format and font state a range currently has;
`find` reports how many cells were searched and how many matched, never their text.

**Audits return names, not contents.** `AuditWorkbookFlows` reports worksheets, tables,
defined names, queries, connections, macro procedure names, data model objects, pivots, and
external links — as names, counts, and addresses. It omits connection strings, and it does not
read stored M.

That last one is a size decision, not a secrecy one, and 0.20.0 is where the difference started
to matter. An audit walks every query in the workbook; `ManageQuery` in Plan walks the one you
named, and returns its full expression. If you want to see a query, ask for that query.

## 4. Content you explicitly asked for

Four operations return workbook content because the content **is** the request:

- **`ReadWorksheetRange`** returns cell values, or R1C1 formulas if you asked for formulas.
  At most 400 cells, blanks omitted, each cell's text capped.
- **`EditMacroProcedure` in Plan** returns the bounded source and hash of the one procedure
  you named. Apply never returns source.
- **`ManageModelMeasure` in Plan** returns the measure's DAX.
- **`ManageQuery` in Plan returns the stored M expression**, with its fingerprint and length.
  Apply never returns it. At most 8,192 characters, and **omitted rather than truncated** past
  that — half an M expression reads exactly like a whole one, and sending it back as a
  replacement would destroy the query.

**Read this one before you decide it is fine.** An M expression is where a workbook says what
it connects to: `Sql.Database("server", "db")`, `Web.Contents("https://…")`, and — if whoever
wrote the query hardcoded one instead of using Excel's credential store — an API key or token
in plain text. All of that now goes to your MCP client and to the model behind it.

Until 0.20.0 it did not. The expression was withheld and Plan returned only a fingerprint, on
the reasoning that a hash proves you are replacing what you looked at without carrying what the
query connects to. What that cost was the ability to look at all: a caller either edited blind
against a hash or left the tool to read the query in Excel. The owner of this project asked for
the text, on the basis that the model reading it is enterprise Copilot, which does not train on
or access their data. **That is a property of one deployment, not of this software.** If you
run ExcelTask against a model with different terms, this section is the paragraph to weigh.

## 5. What *you* send is not covered by any of the above

The request travels the same MCP channel in the other direction. If you supply M to
`ManageQuery`, DAX to `ManageModelMeasure`, or VBA source to `EditMacroProcedure`, that text
crosses the MCP boundary and reaches the model, whatever the read side promises.

Before 0.20.0 there was an asymmetry worth stating here — Plan would not show you a stored M
expression, but Apply required you to send one. There is no asymmetry now: M crosses the
boundary in both directions.

## 6. The field-check report

`--field-check` writes three artifacts. They are not equally shareable and should not be
treated as one thing.

The **`.md` and `.json` reports** describe the machine, because that is their whole purpose.
They contain the computer name, OS build, .NET runtime version, Office and Excel versions
and build, Excel macro-trust registry values, OneDrive sync-root counts, whether well-known
folders accept a new file, the number of Excel processes already running, and **the ProgIDs
of every connected COM add-in**. That last one is more revealing than it looks: an add-in
list names the finance stack a shop runs.

As of 0.19.0 the user profile directory is rewritten as `%USERPROFILE%` in the server path
and `DOTNET_ROOT`, so the Windows account name is no longer in the report. It used to be,
alongside the computer name — together, those name a person as well as a machine. If you
hold a report generated before 0.19.0, it has the account name in it.

The **compact digest** carries versions, per-operation status and timing, the leak count,
and the pass/fail result. No machine name, no add-in list, no paths. **Send the digest.**

The default output directory is your Desktop, which on a managed machine is usually
OneDrive-synced — meaning these reports sync to corporate storage as soon as they are
written. That is normally what you want. It is not nothing, so it is written down.

## Wording this project does not use

- ~~"ExcelTask sends no private data."~~ Section 3 and section 4 are both counterexamples.
- ~~"The AI never sees your workbook."~~ It sees exactly what section 3 and section 4 describe.
- ~~"The trace file is safe to share."~~ It contains workbook and worksheet names; whether
  that is safe depends on what yours are called.
- ~~"Your server names never reach the model."~~ True until 0.20.0, false now: a `ManageQuery`
  Plan returns the M expression, and M is where a query names its source. Section 4.

What is accurate: **ExcelTask emits no incidental workbook contents, and sends nothing
anywhere on its own. Explicit reads, selected Plans, and the record of what a write
overwrote return the content that was asked for.**

## Checking any of this yourself

Everything above is a property of code in this repository, not a policy statement about it:

- telemetry — [`Program.cs`](src/ExcelTask.McpServer/Program.cs)
- the trace contract — [`DiagnosticTrace.cs`](src/ExcelTask.Excel/DiagnosticTrace.cs) and
  [`docs/DIAGNOSTIC-TRACE.md`](docs/DIAGNOSTIC-TRACE.md)
- receipt bounding — `ReceiptBounds` in [`src/ExcelTask.Core`](src/ExcelTask.Core)
- what a write returns — [`ExcelWorkbookRuntime.Write.cs`](src/ExcelTask.Excel/ExcelWorkbookRuntime.Write.cs)
- the field report — [`FieldCheck.cs`](src/ExcelTask.McpServer/FieldCheck.cs) and
  [`docs/FIELD-CHECK.md`](docs/FIELD-CHECK.md)
