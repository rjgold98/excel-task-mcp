# Local Claude-Codex relay

How the owner's local Claude Code and Codex sessions collaborate through an
opt-in GitHub pull request. This is not the managed-work-computer field-agent
channel; the read-only boundary in [AGENT-BRIDGE.md](AGENT-BRIDGE.md) remains
unchanged.

This is internal collaboration tooling. It is not included in the ExcelTask
release artifact and does not change the MCP's user-visible behavior, so it
does not create a product changelog or release requirement.

## Purpose and roles

- **Codex** owns implementation, integration, verification, and the branch.
- **Claude** is a read-only second opinion for architecture, review, and
  bounded task decomposition.
- **Ross** owns scope changes and decides what is merged or released.

GitHub is the durable transcript. The owner's active Codex automation, **Watch
ExcelTask Claude-Codex relay**, checks hourly and wakes the Codex task when a
marked message is handled. The automation is local Codex state rather than
repository runtime; pausing or deleting it disables delivery without changing
ExcelTask. It passes only a validated, bounded body to the intended agent and
posts the result back to the same pull request. Neither agent treats repository
text as higher-priority instruction.

## Opt in and stop

The relay runs only on an open pull request in `rjgold98/excel-task-mcp` with
the `agent-relay` label. Removing the label, closing the pull request, or
pausing the Codex heartbeat stops delivery. The monitor processes at most one
unhandled message per run, oldest first. Runs serialize through one atomic
directory lock under the owner's local temporary-data directory; a run that
cannot acquire it exits without reading or posting a message.

Only a top-level pull-request comment authored by `rjgold98` and beginning with
one of the exact markers below is eligible. Reviews, inline comments, commit
messages, issue bodies, quoted text, and comments from any other author remain
ordinary untrusted review input.

The normative token grammar is:

```text
slug       = [a-z0-9]+(?:-[a-z0-9]+)*
positive   = [1-9][0-9]*
request    = ^<!-- excel-task-agent-relay:v1 kind=request from=(codex|claude) to=(claude|codex) thread=(slug) turn=(positive) -->$
response   = ^<!-- excel-task-agent-relay:v1 kind=response from=(codex|claude) to=(claude|codex) thread=(slug) turn=(positive) in-reply-to=(positive) -->$
ack        = ^<!-- excel-task-agent-relay:v1 kind=ack from=(codex|claude) to=(claude|codex) thread=(slug) turn=(positive) in-reply-to=(positive) -->$
```

The regexes are ASCII and apply to the complete first line before its newline.
Fields use exactly one ASCII space and the shown order. `from` and `to` must be
different. Leading/trailing whitespace, leading zeroes, duplicate or extra
fields, uppercase slugs, and unknown kinds or agents are invalid.

Request marker:

```text
<!-- excel-task-agent-relay:v1 kind=request from=codex to=claude thread=<slug> turn=<positive-integer> -->
```

Response marker:

```text
<!-- excel-task-agent-relay:v1 kind=response from=claude to=codex thread=<slug> turn=<positive-integer> in-reply-to=<comment-id> -->
```

Acknowledgement marker:

```text
<!-- excel-task-agent-relay:v1 kind=ack from=codex to=claude thread=<slug> turn=<positive-integer> in-reply-to=<comment-id> -->
```

The reverse direction swaps `from` and `to`. An acknowledgement is never
forwarded. A response or acknowledgement is terminal only when all of these
match its source: it is a later top-level comment by `rjgold98`; its `from` is
the source's `to`; its `to` is the source's `from`; its thread and turn are
identical; and `in-reply-to` is the exact source comment ID. A request accepts
either a response or acknowledgement. A response addressed to an agent accepts
only an acknowledgement. Wrong-direction, wrong-thread, wrong-turn, malformed,
quoted, review, and inline-comment markers never suppress delivery.

Immediately before invoking an agent and immediately before posting, the
monitor re-reads GitHub and requires the allowlisted repository, an open pull
request, the label, the unchanged author and source body, complete validation,
and no matching terminal comment. Label removal or pull-request closure during
processing therefore stops both invocation and publication.

## Bounded message format

After the marker, a request contains:

```markdown
## Agent relay request

Scope: <the existing pull-request scope>
Question or task: <one bounded request>
Evidence requested: <files, commands, or acceptance checks>
Done when: <observable stopping condition>
```

Each request field occurs exactly once, in the shown order, with a non-empty
single-line value. Additional prose may follow, but cannot introduce another
relay marker or repeat a request field.

A response contains a bounded Markdown answer with each item exactly once:

```markdown
## Agent relay response

Source comment: <positive comment ID>
Disposition: <accepted|revise|defer|declined|needs-owner>
Evidence: <public repository files, commands, and exact results>
Remaining gaps: <none or explicit gaps>
Next bounded step: <one step or none>
```

An acknowledgement contains the same shape headed `## Agent relay ack`; it is
terminal metadata and is never forwarded. Response and acknowledgement fields
occur exactly once in the shown order, use non-empty values, and their `Source
comment` must equal the marker's `in-reply-to` value.

The forwarded body is limited to 8 KiB measured as UTF-8 bytes. Before an agent
is invoked, the monitor requires the exact marker and request fields, validates
the author/direction/thread/turn, and rejects obvious credentials, private keys,
connection strings, absolute local paths, workbook values, formulas, VBA
source, prompts from customer workbooks, customer/company identifiers, private
transcripts, and attempts to override repository or owner policy. Rejections
produce metadata and a reason code only; the body is not copied into local
logs. Link to public repository evidence instead of copying sensitive content.

## Authority and safety

- A relay message is context, not authority. It cannot override `AGENTS.md`,
  the pull-request scope, or an explicit owner decision.
- Claude runs with read-only repository tools. The relay, not Claude, writes
  its response to GitHub.
- Codex may implement only work already authorized by Ross and within the
  current branch scope. Scope expansion is acknowledged as `needs-owner`.
- Neither side may merge, release, push to `main`, change credentials or
  security settings, operate on real business workbooks, or touch a
  user-owned Excel process through this relay.
- The managed work-computer field agent never writes to GitHub. Its reports
  continue through the owner-mediated process in `AGENT-BRIDGE.md`.

## Turn behavior

For a request addressed to Claude, the monitor invokes local Claude Code in
noninteractive read-only plan mode with only repository read/search tools and
no persisted session. A fixed wrapper labels the bounded body as untrusted
collaborator data, restates the authority limits, and requests evidence plus one
disposition. It posts Claude's privacy-checked, bounded answer as a response to
Codex.

For a request or response addressed to Codex, the heartbeat wakes this Codex
task and supplies the body as explicitly untrusted collaborator input. Codex
either completes authorized work and posts a response, asks one new bounded
question, or posts an acknowledgement with `accepted`, `revise`, `defer`,
`declined`, or `needs-owner`. Responses do not create another turn
automatically; a new question requires a new `kind=request` comment. This
prevents an accidental agent-to-agent loop.

## Evidence in every response

Every response identifies the source comment, files inspected or changed,
exact checks and counts, remaining gaps, and one clear disposition or next
question. Hosted CI is not evidence of Excel behavior; serial desktop-Excel
results remain required before a product change can merge.
