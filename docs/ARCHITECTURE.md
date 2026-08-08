# Architecture

## Task flow

```text
GitHub Copilot (currently selected model)
  -> MCP adapter: one excel_task interface
  -> Task Engine: inspect, validate, plan, policy, receipt
  -> Workbook Runtime port
  -> Excel adapter: one STA owner, exact workbook identity, COM, save/verify
  -> compact structured receipt
```

The model decides what outcome to request. Deterministic code decides how to
perform and verify it. The server never delegates planning to a hidden model.

## Deep modules

### Task Engine

The external module interface is `IExcelTaskEngine.RunAsync`. It hides request
normalization, overwrite and live-workbook confirmation, plan compilation,
outcome classification, and receipt construction. Tests use an in-memory
workbook runtime adapter through the same interface.

### Excel adapter

The Excel adapter satisfies the workbook-runtime port. It owns one STA thread,
all COM references, exact open/owned workbook identity, formula/exhibit
execution, save/reopen verification, and process cleanup. COM types never cross
the seam.

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
- Measure actual Copilot tokens and turns separately from server/COM timing.
