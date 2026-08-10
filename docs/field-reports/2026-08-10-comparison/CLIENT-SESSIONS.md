# Field report - client session comparison, 2026-08-10

Six fresh client sessions on the work computer, one workflow each, three per
server, run sequentially against disposable workbook copies. Model
`gpt-5.6-luna` at maximum reasoning, default context tier, held constant across
all six. Telemetry cut at each session's first `session.task_complete`, so
follow-up messages could not inflate the totals.

## Result

| Metric | ExcelTask | Original MCP | Difference |
|---|---|---|---|
| Model requests | 11 | 40 | -72.5% |
| MCP calls | 6 | 37 | -83.8% |
| Input tokens | 509,883 | 1,961,596 | -74.0% |
| Output tokens | 5,692 | 10,327 | -44.9% |
| Total tokens | 515,575 | 1,971,923 | -73.9% |
| Reasoning tokens | 3,483 | 7,101 | -51.0% |
| Model duration | 74.8s | 216.2s | -65.4% |
| Prompt to task complete | 124.8s | 263.4s | -52.6% |
| **Active MCP execution** | **48.0s** | **42.6s** | **+12.7%** |
| Correct after reopening | 3/3 | 3/3 | - |

## Where the time actually goes

This is the finding that explains the rest, and it is not the one the product
markets itself on.

| | ExcelTask | Original MCP |
|---|---|---|
| Workflow span | 62.5s | 170.5s |
| Active MCP execution | 48.0s | 42.6s |
| Model coordination between calls | 14.5s | 127.9s |

**ExcelTask's Excel work is slower**, by about 13%. Its deep calls average 8.0
seconds against the original's 1.15 seconds per low-level call, because one call
inspects, plans, executes, saves, reopens, verifies and cleans up. The original's
individual calls are quick.

The entire advantage comes from what happens *between* calls. The original spent
127.9 seconds of a 170.5-second span with the model deciding what to call next.
ExcelTask spent 14.5 seconds. Fewer decisions, not faster Excel, is the product.

That also means a call count is a count of model round trips, not of equal work.
The two units are not comparable and should never be presented as though they
were.

## What this run does not establish

**It does not measure the tool-surface advantage.** Both catalogs were globally
registered during these sessions. Each child called only its assigned server -
the event logs confirm that - but both schemas may have been in model context
throughout, so the token figures measure orchestration, not schema loading. The
separate field check remains the only valid schema comparison. A future run needs
one MCP registered per client profile.

**It is one run per workflow.** Useful field evidence, not a benchmark. Three or
more repetitions reporting median and spread would be needed before any of these
percentages should be quoted as characteristic.

## Product defects this exposed

The macro workflow cost two avoidable corrections, both caused by rules the
caller could not see in the schema:

1. A Plan carried the Apply-only replacement and run fields.
2. The macro operation used the default `AskIfOpen` binding where only `Isolated`
   is permitted - which the tool description actively encouraged by saying
   "Start with AskIfOpen" without exception.

Both were fixed by making the rules visible in the schema rather than only in the
rejection message, and a protocol test now holds them there.

## Method notes worth keeping

- **Do not message a benchmark session after it completes.** Follow-up turns
  inflate its totals. One ExcelTask macro session read 543,876 tokens raw against
  284,867 when cut at first `session.task_complete`.
- **A transport success is not a semantic success.** One original-MCP run
  returned a delivered-but-rejected probe. Both `tool.execution_complete.success`
  and the returned status must be checked.
- **A client captures its MCP tool set at session start.** A session created
  before ExcelTask was registered never saw `excel_task`; a new one did
  immediately.

## Where the numbers came from

Copilot stores a per-session event stream at
`%USERPROFILE%\.copilot\session-state\<session-id>\events.jsonl`, and summarized
model usage in `assistant_usage_events` inside
`%USERPROFILE%\.copilot\session-store.db`.

Per model request: input, output, cache-read, cache-write and reasoning tokens,
duration, time to first token, inter-token latency, model, reasoning effort and
finish reason. Per tool call: server name, tool name, start and completion
timestamps, success, and MCP result and structured-content byte counts.

Not reliably available: dollar cost, reasoning text, and the derivation of the
Insights context-percentage graph.

Raw event logs contain prompts, tool arguments, workbook paths and repository
context, and stay on the machine. Only redacted aggregates appear here.
