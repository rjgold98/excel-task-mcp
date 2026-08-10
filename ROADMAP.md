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
   The three controlled client tasks then ran as six fresh sessions, three per
   server: 74% fewer input tokens, 73% fewer model requests, 84% fewer MCP calls,
   and 53% less wall time to a verified workbook, with all six correct after
   reopening. See `docs/field-reports/2026-08-10-comparison/CLIENT-SESSIONS.md`.
   Two qualifications belong with those figures: ExcelTask's own Excel execution
   was 13% *slower*, and the whole advantage came from removing model
   coordination between calls; and both tool catalogs were registered during
   those sessions, so they measure orchestration, not schema loading. This phase
   is met. What remains is repetition - one run per workflow is evidence, not a
   benchmark.
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
   flows before considering mutation. v0.7.0 delivered the single-workbook
   slice: `AuditWorkbookFlows` reports queries and their load destinations,
   connections, model tables, relationships and measures, pivots, and links to
   other workbooks - names and shapes only, never values, M text, connection
   strings, or paths - with the receipt proving by size and timestamp that the
   workbook was not changed. The development gate is met: the fixture set
   produces a correct bounded report, verified against real Excel. On
   2026-08-10 the first real-workbook audit ran on the work computer: one call,
   nothing changed by the receipt's own proof and by independent metadata, no
   process left behind. See `docs/field-reports/2026-08-10-audit/`. That proved
   the audit safe on real content but not yet complete on rich content - the
   chosen workbook had no Power Query or Data Model flows, so those surfaces
   have only been exercised against fixtures. Remaining field gate: one audit
   of a workbook the owner knows to contain query and model flows, with the
   owner confirming the reported categories. Following multiple workbooks
   through their links into one report remains open.
