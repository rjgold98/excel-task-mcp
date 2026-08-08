# Roadmap

Each phase starts only after its gate is met. This is direction, not a feature
promise.

1. **Work-computer validation and reliability** - validate real Copilot use,
   measure task time/tokens, and close reliability gaps exposed by evidence.
   Gate: three controlled tasks save correctly with telemetry and no orphaned
   Excel process.
2. **Formula/exhibit depth** - improve the existing worksheet-copy and safe
   formula-repair workflow only where real use requires it. Gate: three
   representative user tasks finish with verified outputs and no new tool.
3. **Macro editing** - add one narrow macro-editing workflow. Gate: a
   disposable `.xlsm` edit, run, save, and reopen test fails closed on dialogs
   and leaves no owned Excel process.
4. **Read-first multi-workbook audit** - inspect Power Query and Data Model
   flows before considering mutation. Gate: the agreed fixture set produces a
   correct bounded dependency report without modifying a workbook.
