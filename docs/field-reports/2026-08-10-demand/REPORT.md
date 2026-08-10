# Field report - measured demand, 2026-08-10

Field task 005. The original Excel MCP's complete advertised surface, and which
of it the owner's real work actually reached for. Relayed by the owner from the
field agent; no Excel workflow was run, only `tools/list` and a read of the
session history.

**History:** 264 session logs, 339,779 events, 2026-06-22 to 2026-08-10. The
original MCP was used in **46 sessions, 7,873 calls**.

## The surface

25 tools, **230 operations**, 57,641 bytes of `tools/list`. ExcelTask advertises
one tool and 5 operations in 9,422 bytes.

## Sessions, not calls

Calls are unusable as a demand signal here: the median Excel session made 32
calls and the largest made 2,515, so a single unusual session would dominate any
ranking. Sessions-per-tool is used throughout.

| Tool | Sessions | Covered by ExcelTask |
|---|---:|---|
| `file` | **46 (all)** | internally, never exposed |
| `range` | 36 | partly - infers formulas, never reads or writes values |
| `worksheet` | 30 | listing added in 0.8.0 |
| `vba` | 18 | listing added in 0.8.0; one-procedure edit and run |
| `namedrange` | 13 | no |
| `range_edit` | 13 | no |
| `connection` | 13 | listing only |
| `screenshot` | 13 | no |
| `calculation_mode` | 12 | internally, never exposed |
| `range_format` | 12 | no |
| `table` | 11 | no |
| `datamodel` | 10 | listing only |
| `powerquery` | 10 | listing only |
| `pivottable` | 7 | listing only |
| `conditionalformat` | 4 | no |
| `window` | 4 | no |
| `range_link` | 3 | no |
| `pivottable_field`, `datamodel_relationship` | 2 | relationships listed |
| `chart`, `slicer`, `table_column` | 1 | no |
| **`chart_config`, `pivottable_calc`, `worksheet_style`** | **0** | no |

## The top operations are overwhelmingly discovery and reading

By sessions:

| Operation | Sessions | Status |
|---|---:|---|
| `range get-values` | **31** | **the largest gap** |
| `worksheet list` | 28 | added 0.8.0 |
| `range get-used-range` | 25 | added 0.8.0, per worksheet |
| `range get-formulas` | 21 | no |
| `range set-values` | 19 | no, by design |
| `vba list` | 18 | added 0.8.0 |
| `range set-formulas` | 15 | no, by design - never accepts model-written formula text |
| `namedrange list` | 12 | no |
| `connection list` | 12 | yes |
| `vba update` | 12 | yes |
| `vba view` | 11 | Plan, for one named procedure |
| `table list` | 11 | no |
| `screenshot capture` | 11 | no |
| `powerquery list` | 10 | yes |
| `range_edit find` | 10 | no |
| `vba run` | 10 | yes |
| `calculation_mode calculate` | 10 | internal |

## What this settles

**The deep-call bet is confirmed from the demand side.** Every session that used
`range`, `worksheet`, `vba`, `connection`, `datamodel` or `powerquery` *also*
used `file` - lifecycle is pure overhead the other server forces onto the model,
and 1,079 of its 7,873 calls were spent on it. That is the cost ExcelTask's one
deep call removes, measured rather than argued.

**Most of the 230 operations are dead weight.** Three tools were never called at
all, and `chart`, `slicer` and `table_column` appeared in one session each. The
8.1x schema ExcelTask does not carry is largely paid for capability this owner
does not use.

**Reading is the demand, not writing.** The single most-used operation in five
weeks of real work is `range get-values` in 31 of 46 sessions, followed by
`worksheet list` and `get-used-range`. This corrects an assumption in the
roadmap: value reading was recorded as one incident's blocker, and it is in fact
the most frequent thing the owner's work does.

**Two design stances now sit against measured demand.** `range set-values` (19
sessions) and `range set-formulas` (15) are both deliberate refusals - ExcelTask
infers formulas and never accepts model-written text. That refusal is what makes
its edits safe, and it is also the second-largest block of observed demand. That
tension is real and is not resolved by this report.

## Caveats kept

- Counts come from `tool.execution_start`, so rejected attempts still count as
  demand - which is correct for measuring intent.
- The window includes setup, field testing, and the controlled comparison runs
  for this project, so some volume is this work rather than the owner's.
- Two malformed operation identifiers were redacted; no prompts, workbook names,
  paths, sheet names, ranges, values, formulas, VBA source, or connection details
  appear anywhere in the report.
