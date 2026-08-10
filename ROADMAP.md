# Roadmap

Each phase starts only after its gate is met. This is direction, not a feature
promise.

1. **Work-computer validation and reliability** - v0.2.0 delivered the private
   worker, hard deadline, truthful interruption handling, and process cleanup.
   Remaining gate: three controlled Copilot tasks save correctly with measured
   time/tokens and no orphaned Excel process.
2. **Formula/exhibit depth** - v0.3.0 delivered bounded in-place gap repair and
   stable right/down formula-series extension through the existing one-tool
   operation union. Remaining field gate: three controlled Copilot tasks finish
   with verified outputs and no new tool.
3. **Macro editing** - v0.4.0 delivered the `EditMacroProcedure` operation for an
   isolated `.xlsm` saved only as a `Copy`: Plan returns a bounded
   requested-procedure source and hash, Apply requires that hash plus a complete
   replacement, returns no source, and may optionally run the no-argument
   procedure. Trust access remains user-controlled. Remaining field gate: a
   controlled work-computer edit where Trust Center policy, not the product,
   decides what is permitted.
4. **Read-first multi-workbook audit** - inspect Power Query and Data Model
   flows before considering mutation. Gate: the agreed fixture set produces a
   correct bounded dependency report without modifying a workbook.
