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
manual closed operation union—copy exhibit, repair existing worksheet, extend
formula series, and the in-progress `EditMacroProcedure`—so the MCP schema
stays small without generic action language. It hides normalization, overwrite
and live-workbook confirmation, plan compilation, outcome classification, and
receipt construction. Tests use an in-memory workbook runtime adapter through
the same interface.

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
  MCP wire. The in-progress macro operation bounds procedure names/hashes to
  96 characters and Plan source to 8,192 characters; Apply returns no source.
- Measure actual Copilot tokens and turns separately from server/COM timing.
