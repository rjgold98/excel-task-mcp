# Field report - read-only audit validation, 2026-08-10

Field task 003, run on the managed work computer with ExcelTask v0.7.0 and
relayed by the owner. Transcribed from the field agent's report; digests and
check lines are verbatim.

## Field check

Release ZIP SHA-256 matched
`384c1ae29a9e0d440b587b6c577ae49ca4b1895571c69d0f2770c5174a634e0c`. Exit 0.

```text
----- EXCELTASK FIELD DIGEST -----
excel=16.0.20228 vbom=1 vbawarn=1
self  v0.7.0 tools=1 bytes=8182
other none
CopyExhibit (Plan)         Planned             6.8s L0
CopyExhibit (Apply)        Completed          12.0s L0
ExtendFormulaSeries (Apply Completed          12.2s L0
AuditWorkbookFlows (Apply) Completed          11.9s L0
EditMacroProcedure (Plan)  Planned             7.1s L0
EditMacroProcedure (Apply+ Completed          11.6s L0
leaked=0 result=PASS
----- END DIGEST -----
```

## First audit of a real business workbook

One closed, local `.xlsm`, chosen and authorized by the owner, path and item
names kept on the machine. Because the running client session had captured the
v0.6.3 schema at startup, the field agent drove the verified v0.7.0 server
directly over MCP stdio - so this validates the operation, not client
orchestration.

| | |
|---|---|
| Status | Completed, 1 call, 0 retries |
| Found | 1 connection, 1 pivot; 0 queries, 0 model tables, 0 relationships, 0 measures, 0 external links; totalFound 2 |
| Flags | truncated false, workbookUnchanged **true** |
| Elapsed | 11.2s tool, 12.1s wall |
| Excel processes | identical before and after; none new |

Check lines, verbatim:

```text
audit-scan: PASS - All flow surfaces were read; 2 item(s) found.
workbook-unchanged: PASS - The workbook's size and timestamp are identical before and after the read-only audit.
```

Independently verified by the field agent, outside the receipt: file size
(1,163,555 bytes) and last-write timestamp identical before and after, and the
set of Excel process IDs identical before and after.

## What this proves, and what it does not

**Proven in the field:** the audit is safe on real content. It ran once against
a real business workbook, changed nothing - by the receipt's own proof and by
independent metadata - leaked nothing, and returned a bounded report in one
call.

**Not yet proven: completeness on rich content.** The chosen workbook happened
to contain no Power Query, Data Model, or external-link content, so the
surfaces that matter most for the owner's audit work were exercised only
against fixtures, not real files. The owner confirmed the right workbook was
audited but could not confirm whether the counts match its actual contents, and
the field agent correctly declined to claim a match it could not verify.

The remaining gate for this phase is therefore narrow: one audit of a workbook
the owner knows to contain Power Query and Data Model flows, with the owner
confirming the reported categories against what they expect.

## Method notes

- A client session captures its MCP tool set at startup; a session started
  before an upgrade keeps the old schema. Field validation of a new release
  needs a fresh session or a direct harness.
- The managed host's Constrained Language Mode postamble error
  (`$host.SetShouldExit`) appeared again after successful child commands. It is
  the wrapper's own failure, occurs after `EXCELTASK_EXIT=0`, and is noise.
