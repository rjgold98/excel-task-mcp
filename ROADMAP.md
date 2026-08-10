# Roadmap

Each phase starts only after its gate is met. This is direction, not a feature
promise.

1. **Work-computer validation and reliability** - v0.2.0 delivered the private
   worker, hard deadline, truthful interruption handling, and process cleanup.
   On 2026-08-10 the complete suite passed 146/146 on the work computer, and the
   field check then passed there against real Excel with four of the user's own
   Excel processes running and untouched. The tool-surface comparison is done:
   1 tool and 7,164 bytes against the original's 25 tools and 58,324 bytes, a
   ratio of 8.1x. See `docs/field-reports/2026-08-10-comparison/`.
   Remaining gate: the three controlled client tasks with measured prompt-to-done
   time. Until those run, nothing is known about end-to-end speed or token use
   during real work - only about context cost before work begins.
2. **Formula/exhibit depth** - v0.3.0 delivered bounded in-place gap repair and
   stable right/down formula-series extension through the existing one-tool
   operation union. Remaining field gate: three controlled Copilot tasks finish
   with verified outputs and no new tool.
3. **Macro editing** - v0.4.0 delivered the `EditMacroProcedure` operation for an
   isolated `.xlsm` saved only as a `Copy`: Plan returns a bounded
   requested-procedure source and hash, Apply requires that hash plus a complete
   replacement, returns no source, and may optionally run the no-argument
   procedure. v0.6.0 added dialog containment, so a run-time error, a compile
   error, or a message box returns a named outcome instead of stalling.
   Trust access remains user-controlled. Field gate met 2026-08-10:
   the work computer ran the disposable `.xlsm` edit, run, save, and reopen
   tests through the real MCP boundary, with VBA project access permitted by
   that machine's own policy and no Excel process left behind.
4. **Read-first multi-workbook audit** - inspect Power Query and Data Model
   flows before considering mutation. Gate: the agreed fixture set produces a
   correct bounded dependency report without modifying a workbook.
