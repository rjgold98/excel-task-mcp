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

- **Issues** - tasking. Each field task is one issue authored by the lead,
  labeled `field-task`, self-contained: context, exact commands, expected
  outputs, and the report format. The field agent works it and reports back on
  that same issue.
- **Pull requests** - anything with files: field reports committed under
  `docs/field-reports/`, or proposed code changes. Always referenced back to
  the issue that asked for them.
- **Comments** - discussion. One thread per finding. Every response ends with a
  question or a disposition: accepted, fixed in `<commit>`, or declined with the
  reason.

## Field task lifecycle

1. Lead opens an issue labeled `field-task` with numbered steps.
2. Owner points the field agent at it.
3. Field agent executes the steps. If a step cannot be executed - blocked by
   policy, missing tool, unexpected state - it records exactly what happened
   and continues with the remaining steps rather than improvising a workaround.
4. Field agent reports on the issue using the evidence standard below, and
   commits any report files to a branch named `field/<issue-number>-<topic>`,
   opening a pull request that references the issue. If pushing a branch is
   blocked, it pastes full file contents into issue comments instead.
5. Lead reads everything with the GitHub CLI, responds on the same threads,
   fixes what needs fixing, and closes the issue with a disposition.

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
- The field agent does not push to `main`, merge, tag, or release.
- Anything a task does not cover is a question on the issue, not an
  improvisation.
