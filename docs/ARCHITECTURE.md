# Architecture

## Task flow

```text
GitHub Copilot (currently selected model)
  -> MCP adapter: one excel_task interface
  -> Task Engine: inspect, validate, plan, policy, receipt
  -> Workbook Runtime port
  -> Supervisor: one bounded private worker per operation
  -> Excel adapter: one STA owner, exact workbook identity, COM, save/verify
  -> compact structured receipt
```

The model decides what outcome to request. Deterministic code decides how to
perform and verify it. The server never delegates planning to a hidden model.

## Deep modules

### Task Engine

The external module interface is `IExcelTaskEngine.RunAsync`. Its request has a
manual closed operation union of fourteen kinds, so the MCP schema stays small
without generic action language:

| Operation | Reads or writes | Notes |
|---|---|---|
| `CopyExhibit` | writes | Copies a modelled worksheet into the target, then binds its references home |
| `RepairExistingWorksheet` | writes | Infers blank formulas from their neighbours |
| `ExtendFormulaSeries` | writes | Extends a series right or down |
| `EditMacroProcedure` | writes | One named procedure; isolated `.xlsm` + `Save=Copy` only |
| `AuditWorkbookFlows` | reads | Opens Excel; the only operation that sees queries, the model, and macros |
| `ReadWorksheetRange` | reads | One bounded range |
| `WriteWorksheetValues` | writes | Constants only; never formula text |
| `FindReplace` | both | Plan locates, Apply rewrites constants |
| `Create` | writes | Empty workbook or worksheet; never overwrites |
| `SetRangeFormat` | writes | Appearance only: number format, font, fill, borders, width, height |
| `ScanWorkbookStructure` | reads | **Starts no Excel** — reads the package directly |
| `ManageTable` | writes | Create over a range, rename, restyle, resize, or convert back to cells |
| `ManageQuery` | writes | One Power Query, under a fingerprint; Plan never returns the M expression |
| `ManageModelMeasure` | writes | One Data Model measure, under a fingerprint; Plan does return the DAX |
| `ManageModelRelationship` | writes | One Data Model relationship, many side to one side; Create or Delete, never Replace |

The engine hides normalization, overwrite and live-workbook confirmation, plan
compilation, outcome classification, and receipt construction. Tests use an
in-memory workbook runtime adapter through the same interface.

`OperationCatalog` is the one place that maps a kind to its payload. Its switch
has no default arm, so adding an operation without handling it fails the build.
That exists because the hand-kept array it replaced once shipped two operations
unreachable: a payload missing from the list did not merely go uncounted, the
request failed the arity check before reaching its own validation.

`ScanWorkbookStructure` is the only operation that never starts Excel. It reads
the `.xlsx` package as what it physically is, a ZIP of XML, and reports sheets,
dimensions, formula and constant counts, constant islands, defined names,
tables, and external links by file name. It reports **nothing** about macros,
queries, connections, or the data model, because those are not readable from the
package — VBA is a binary OLE compound file, Power Query is base64 of a nested
ZIP, the model is opaque — and a receipt that answered some categories while
reading as complete would be a receipt that lies. Its description says so, and a
test asserts the silence against a fixture that contains a VBA project.

`EditMacroProcedure` has no generic VBE surface: it selects one procedure by
name, Plan returns only that bounded source/hash, and Apply requires a complete
replacement with the expected hash. It is isolated `.xlsm` + `Save=Copy` only;
an optional no-argument run remains bounded by the existing worker deadline.
Trust access is intentionally controlled by the user, while a dialog or timeout
after dispatch is `Unknown` rather than retried.

### Excel adapter

The Excel adapter satisfies the workbook-runtime port. It owns one STA thread,
all COM references, exact open/owned workbook identity, formula/exhibit
execution, save/reopen verification, and process cleanup. COM types never cross
the seam.

The MCP host never loads Excel COM. A short-lived private worker owns the STA
adapter, and a worker-local watchdog may terminate only the exact Excel process
captured from the application it created. The supervisor treats worker process
reports as evidence only, enforces a two-minute deadline, and returns `Unknown`
when completion or cleanup cannot be proved.

The adapter has two ownership modes:

- `isolated`: ExcelTask owns Excel and may close/quit it;
- `use_open`: ExcelTask attaches to one exact workbook after confirmation and
  must not close or quit the user's Excel process.

**Workbook identity is exact, and a synced path has two exact spellings.** A
workbook opened from a OneDrive or SharePoint folder reports a service URL as
its `FullName`, not the local path the caller named, so a straight path
comparison found nothing and refused every `use_open` against synced storage.
`OneDriveSyncMap` resolves the caller's path through the sync client's own
registry mapping and then compares exactly — the identity is still proven, never
guessed, and a same-named workbook in a different library still does not match.
Where nothing is synced, the comparison is unchanged.

**The supervisor's cleanup sweep triggers on what the worker reported**, not on
whether it reported cleanly. A worker that completes its protocol and returns
`Unknown`, or any failed check, has said it could not finish cleaning up; that
is when an independent exit re-check and orphaned-staging deletion are wanted.
Gating them on a silent worker meant they never ran on the path that knows.

### MCP adapter

The MCP adapter exposes exactly one tool and maps typed requests/receipts to
MCP structured content and protocol-level tool errors. It contains no workbook
behavior and no model configuration.

## Outcome truth

- `Planned`: validated preview; no mutation.
- `NeedsConfirmation`: no mutation; user choice is required.
- `Completed`: requested mutation, verification, and persistence succeeded.
- `Rejected`: nothing started; correcting the request is safe.
- `Partial`: known subset changed; receipt identifies the verified subset.
- `Unknown`: dispatch occurred but final workbook state cannot be proved.

Only `Rejected` is automatically safe to retry after correction. A repeated
`Unknown` mutation must first reconcile the workbook state.

## Performance strategy

- Keep the model-facing schema small and stable.
- Combine discovery, rectangular reads, writes, recalculation, and verification
  inside one task call.
- Create isolated Excel only for a task, then prove process exit; do not add
  warm-process reuse until work-computer measurements justify it.
- Bound every task and return changed ranges/checks rather than workbook data.
- Limit formula work to 16 requested ranges, 10,000 scanned cells, 2,000
  planned mutations, and 24 extension periods; never put formula text on the
  MCP wire. The macro operation bounds procedure names/hashes to 96 characters
  and Plan source to 8,192 characters; Apply returns no source.
- Measure actual Copilot tokens and turns separately from server/COM timing.
