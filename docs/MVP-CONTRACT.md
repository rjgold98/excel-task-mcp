# MVP contract

## User outcome

From one `excel_task` request, create a new worksheet in an existing target
workbook using an explicitly named reference worksheet, optionally from a
different workbook. Repair only safely inferable blank formulas, verify the
saved result after reopening it, and report exactly what happened.

## Model-facing interface

The single tool accepts:

- target workbook path;
- reference workbook path and worksheet name;
- new worksheet name;
- zero or more bounded formula-repair ranges;
- `plan` or `apply` mode;
- `ask`, `use_open`, or `isolated` workbook binding;
- save-in-place or save-copy policy;
- explicit overwrite authorization. A `NeedsConfirmation` result is continued
  by resubmitting with the explicit workbook-binding choice and/or overwrite
  authorization; there is no server session or confirmation token.

The tool does not expose sessions, COM objects, low-level command names,
checkpoint switches, idempotency keys, model selection, or CLI behavior.

## Acceptance evidence

1. `tools/list` returns exactly one ExcelTask tool with a bounded schema.
2. A model-free MCP call completes the cross-workbook exhibit task.
3. The copied worksheet exists after save and reopen.
4. Every reported formula repair is verified against the intended FormulaR1C1
   value without returning formula text or cell contents in the receipt.
5. Plan mode makes no workbook change.
6. An open workbook with `ask` returns `NeedsConfirmation` without mutation.
7. Save-in-place without explicit authorization is rejected before Excel
   mutation.
8. The engine never closes or quits a user-owned Excel instance.
9. An isolated task runs in a supervised private worker and leaves no owned
   Excel process or workbook file lock.
10. The receipt distinguishes rejected, completed, partial, and unknown
    outcomes and states whether retry is safe.
