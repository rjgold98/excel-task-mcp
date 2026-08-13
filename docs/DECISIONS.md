# Product decisions

These decisions were locked with the user on 2026-08-08.

1. ExcelTask is a ground-up product. The previous ExcelMcp architecture and
   public interface are not compatibility requirements.
2. Version 1 is optimized solely for GitHub Copilot on the user's managed work
   computer while remaining an MCP implementation.
3. The primary goal is the fastest verified saved workbook result with the
   lowest practical client-token and tool-call cost.
4. Formula and exhibit work is the first mastered workflow.
5. New exhibits normally model an explicitly chosen existing worksheet and are
   usually added to an existing target workbook.
6. Multiple source/reference workbooks are required.
7. Formula repair may infer blank formulas from neighboring/comparable formula
   patterns and must verify every applied repair.
8. If the target workbook is already open, ask whether to use that exact live
   workbook or an isolated file instance.
9. Agent mode executes and verifies; Plan mode previews without mutation.
10. The MCP never selects a model. The client uses the user's current/default
    model and may change it independently.
11. Overwriting an existing file requires explicit authorization.
12. No full CLI, saved template preference, workbook memory, compatibility
    facade, specialist packs, or broad feature parity belongs in the MVP.
13. Caller-supplied formulas are supported only through a separate bounded
    `WriteWorksheetFormulas` operation. `WriteWorksheetValues` remains
    constants-only; direct formulas are read back before save and verified after
    reopen, and formula text is never returned in receipts.
