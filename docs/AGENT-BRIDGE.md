# Agent bridge

How the two AI agents on this project communicate. The owner does not carry
technical detail between them; the repository carries it.

## Roles

- **Lead agent (Claude, home computer)** - owns architecture, code, releases,
  and tasking. Writes field tasks as GitHub issues. Reviews and merges pull
  requests. Cannot see the work computer, so it depends entirely on what the
  field agent reports.
- **Field agent (GitHub Copilot, managed work computer)** - the lead's eyes and
  ears in the real environment: corporate policy, per-user runtimes, add-ins,
  authentication, the original Excel MCP install. Executes field tasks exactly
  as written and reports with full evidence. Does not change product code on
  `main` and does not merge.
- **Owner (Ross)** - facilitates timing only. Tells the field agent "read and
  follow issue #N on rjgold98/excel-task-mcp", and tells the lead when results
  are back. Decides anything either agent escalates.

## Channels

The channel is deliberately one-way. This repository is personal, and the field
agent runs inside an enterprise that does not permit contributing to repositories
outside it. So the field agent **reads** from GitHub and **never writes** to it:
no pushes, no branches, no pull requests, no issue comments. Attempting one and
reporting success without checking is the specific failure this section exists to
prevent.

- **Issues, outbound** - tasking. Each field task is one issue authored by the
  lead, labeled `field-task`, self-contained: context, exact commands, expected
  outputs, and the report format. The field agent reads it.
- **Files, inbound** - reporting. The field agent writes its report as files on
  the work computer and tells the owner where they are. The owner relays them to
  the lead, who commits them under `docs/field-reports/` and answers on the
  issue. A round trip is therefore: lead writes an issue, owner points the field
  agent at it, field agent writes files, owner relays, lead responds.
- **The owner is the only write path.** If something must reach GitHub from the
  work computer, it goes through the owner as file contents, not as a git
  operation.

## Field task lifecycle

1. Lead opens an issue labeled `field-task` with numbered steps.
2. Owner points the field agent at it. If the field agent cannot read the
   repository either, the owner pastes the issue text to it directly.
3. Field agent executes the steps. If a step cannot be executed - blocked by
   policy, missing tool, unexpected state - it records exactly what happened
   and continues with the remaining steps rather than improvising a workaround.
4. Field agent writes its report to files on the work computer, using the
   evidence standard below, and tells the owner the exact paths. It does not
   attempt any git write operation.
5. Owner relays those files to the lead.
6. Lead commits them under `docs/field-reports/`, responds on the issue, fixes
   what needs fixing, and closes the issue with a disposition.

## Reporting format

Everything the lead needs must survive being copied out of the work computer as
plain files, so a report is:

- the generated field-check files, unedited; plus
- one `SUMMARY.md` containing anything the tooling could not capture: what was
  run, what was observed, timings, and the exact text of any failure.

Prefer a small number of complete files to many fragments. Never summarize an
error; paste it.

## Evidence standard

The lead cannot re-run anything on the work computer, so a claim without its
evidence is unusable. Every reported step includes:

- the exact command as executed;
- the relevant output verbatim - counts, exit codes, error text, not summaries;
- the machine context when it matters: paths, versions, policy values;
- timings with what they measure, since "fast" is not a number.

"Tests passed" without the pass/fail/skip counts, or "it errored" without the
error text, will be sent back for the missing detail.

## Boundaries

These bind both agents:

- Never change Excel, Trust Center, or any security or system setting.
- Touch only disposable fixture workbooks; never a real business workbook.
- Never uninstall, reconfigure, or interfere with the original Excel MCP
  installation; measuring it is fine.
- Never force a specific model anywhere, including in tests.
- Never attempt to bypass an employer control - Constrained Language Mode,
  contribution restrictions, Trust Center policy. Report the block and stop.
- The field agent performs no git write operation of any kind.
- Anything a task does not cover is a question for the owner to relay, not an
  improvisation.
