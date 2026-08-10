# MVP contract

## User outcome

From one `excel_task` request, perform one formula/exhibit operation in an
existing target workbook: copy an explicitly named reference worksheet,
repair safely inferable blank formulas on an existing worksheet, or extend a
proven formula series right or down. Version 0.4 is in progress to add one
bounded macro procedure edit; it is not yet a stable or Excel-verified claim.
Verify an applied saved result after reopening it, and report exactly what
happened.

## Model-facing interface

The single tool accepts:

- target workbook path;
- one manual `operation` union: `CopyExhibit`, `RepairExistingWorksheet`,
  `ExtendFormulaSeries`, or the in-progress `EditMacroProcedure`; exactly one
  matching payload is required;
- A1 ranges only; repair/copy ranges are capped at 16 ranges and 10,000 scanned
  cells, while extension is capped at two evidence periods, 1–24 destination
  periods, and 2,000 planned mutations;
- `plan` or `apply` mode;
- `ask`, `use_open`, or `isolated` workbook binding;
- save-in-place or save-copy policy;
- explicit overwrite authorization. A `NeedsConfirmation` result is continued
  by resubmitting with the explicit workbook-binding choice and/or overwrite
  authorization; there is no server session or confirmation token.

`EditMacroProcedure` is isolated `.xlsm` + save-copy only. Plan returns the
explicitly requested procedure's bounded source and hash. Apply requires the
expected hash plus a complete replacement, returns no source, and can only
optionally run that no-argument procedure. Excel Trust access is user-managed;
dialogs and timeouts after dispatch are `Unknown`. The tool exposes no
arbitrary VBE API, sessions, COM objects, low-level command names, checkpoint
switches, idempotency keys, model selection, or CLI behavior. It never accepts
or returns formula text, `FormulaR1C1`, or cell values.

## Acceptance evidence

1. `tools/list` returns exactly one ExcelTask tool with a bounded schema.
2. Model-free MCP calls normalize each operation payload and reject a mismatched
   union before inspection.
3. A copied worksheet or in-place formula mutation is verified after save and
   reopen.
4. Every reported formula repair or extension is verified without returning
   formula text or cell contents in the receipt.
5. Plan mode analyzes but makes no workbook change, save, or recalculation.
6. An open workbook with `ask` returns `NeedsConfirmation` without mutation.
7. Save-in-place without explicit authorization is rejected before Excel
   mutation.
8. The engine never closes or quits a user-owned Excel instance.
9. An isolated task runs in a supervised private worker and leaves no owned
   Excel process or workbook file lock.
10. The receipt distinguishes rejected, completed, partial, and unknown
    outcomes and states whether retry is safe.
