# Power Query and DAX/M parity gap — 2026-08-13

## Question and evidence boundary

This note asks what ExcelTask would need to match the Power Query and Data
Model/DAX surface of [`sbroenne/mcp-server-excel`](https://github.com/sbroenne/mcp-server-excel)
without turning ExcelTask into a compatibility rewrite or a broad command
catalog.

Evidence was collected on 2026-08-13. Upstream references use the pinned main
commit [`a77007d`](https://github.com/sbroenne/mcp-server-excel/tree/a77007db43cc10c38da9d81e736fa5eb57af7f14)
so the operation inventory is reproducible. Local baseline facts use shipped
`HEAD` (`23d6fe5`) in this checkout and the local [operating contract](../AGENTS.md),
[README](../README.md), [MVP contract](../docs/MVP-CONTRACT.md),
[architecture](../docs/ARCHITECTURE.md), and source files.

The worktree currently has an unresolved merge. The incoming local
`origin/main` commit (`6a31ff1`)
contains narrow `ManageQuery`, `ManageModelMeasure`, and
`ManageModelRelationship` designs, but those changes are not part of shipped
`HEAD` and must not be described as current capability. The comparison below
therefore labels **HEAD** and **incoming/unmerged** separately. No product
code, changelog, or roadmap was changed for this research note.

Incoming facts were checked with `git show origin/main:<path>` because the
working files themselves contain unresolved conflict markers; the commit hash
and command are the authoritative local source for that column of the table.

## Executive finding

Exact parity is a pair of deep capability families, not a reason to copy the
upstream server's 26-tool/234-operation catalog. Upstream has a complete
Power Query lifecycle (including ephemeral M evaluation, load destinations,
refresh, unload, and cleanup) and a broad Data Model surface (tables,
measures, relationships, DAX evaluation, and DMV metadata). The shipped
ExcelTask baseline can only audit names/shapes through
`AuditWorkbookFlows`; it has no Power Query or DAX/M mutation or evaluation
operation. The unmerged design narrows that gap for fingerprint-guarded query,
measure, and relationship writes, but still has no refresh/load/evaluate or
model discovery operations.

The highest-value parity sequence is: bounded discovery and readback; isolated
M/DAX evaluation; Power Query load/refresh lifecycle; Data Model refresh and
metadata; then the remaining relationship/formatting conveniences. Every
addition should remain one `excel_task` task with Plan/Apply, a private owning
STA worker, bounded receipts, save/reopen verification, and truthful
`Unknown` after an uncertain mutation dispatch, as required by the local
[operating contract](../AGENTS.md) and [architecture](../docs/ARCHITECTURE.md).

## Upstream capability inventory

### Power Query and M (12 operations)

The upstream [`IPowerQueryCommands`](https://github.com/sbroenne/mcp-server-excel/blob/a77007db43cc10c38da9d81e736fa5eb57af7f14/src/ExcelMcp.Core/Commands/PowerQuery/IPowerQueryCommands.cs)
interface and its [feature matrix](https://github.com/sbroenne/mcp-server-excel/blob/a77007db43cc10c38da9d81e736fa5eb57af7f14/FEATURES.md#-power-query--m-code)
define these twelve actions:

| Group | Actions | What parity means |
|---|---|---|
| Discovery | `list`, `view`, `get-load-config` | Enumerate query names, read stored M, and inspect destinations without exposing unbounded workbook data. |
| Definition lifecycle | `create`, `update`, `rename`, `delete` | Create/load atomically, replace M with optional refresh, normalize names, and remove a query. |
| Destinations and refresh | `load-to`, `refresh`, `refresh-all`, `unload` | Move/load to worksheet, Data Model, both, or connection-only; refresh one/all; remove loaded data while retaining a query definition for `unload`. |
| Ephemeral execution | `evaluate` | Execute M through a temporary query, return bounded preview rows, and clean up without persisting a query. |

The upstream shared [Power Query skill](https://github.com/sbroenne/mcp-server-excel/blob/a77007db43cc10c38da9d81e736fa5eb57af7f14/skills/shared/powerquery.md)
sets the intended workflow to **evaluate → create/update → refresh/load-to**.
It also documents worksheet/Data Model/both/connection-only destinations,
explicit target-cell collision handling, 30-minute refresh/load timeouts,
M-preserving edits, and optional remote `powerqueryformatter.com` formatting
only after consent. These are behavior and safety requirements, not merely
method names.

The upstream [COM behavior findings](https://github.com/sbroenne/mcp-server-excel/blob/a77007db43cc10c38da9d81e736fa5eb57af7f14/docs/COM-API-BEHAVIOR-FINDINGS.md)
are important for a safe implementation: changing `WorkbookQuery.Formula`
does not refresh loaded data; deleting a query does not necessarily remove its
worksheet table or Data Model table; query-only/connection-only state is
represented by load destinations rather than a simple query flag; and rename
can invalidate table connection strings. Those facts make cleanup and
post-save verification part of parity.

### Data Model, DAX, and relationships (19 operations)

The upstream [`IDataModelCommands`](https://github.com/sbroenne/mcp-server-excel/blob/a77007db43cc10c38da9d81e736fa5eb57af7f14/src/ExcelMcp.Core/Commands/DataModel/IDataModelCommands.cs),
[`IDataModelRelCommands`](https://github.com/sbroenne/mcp-server-excel/blob/a77007db43cc10c38da9d81e736fa5eb57af7f14/src/ExcelMcp.Core/Commands/DataModel/IDataModelRelCommands.cs),
and [Data Model feature matrix](https://github.com/sbroenne/mcp-server-excel/blob/a77007db43cc10c38da9d81e736fa5eb57af7f14/FEATURES.md#-data-model--dax-power-pivot)
define nineteen actions across the two interfaces (`14 + 5`). The feature
matrix's same section is labelled “19 operations” but also lists “List
Workbook Connections”; that item is implemented by the separate
[`IConnectionCommands`](https://github.com/sbroenne/mcp-server-excel/blob/a77007db43cc10c38da9d81e736fa5eb57af7f14/src/ExcelMcp.Core/Commands/Connection/IConnectionCommands.cs)
service, so it is called out as an integration dependency rather than counted
as a twentieth Data Model interface action:

| Group | Actions | What parity means |
|---|---|---|
| Tables and columns | `list-tables`, `read-table`, `rename-table`, `delete-table`, `list-columns`; connection integration is `connection.list` | Discover model schema, inspect columns/measures, and manage model tables with explicit destructive impact. |
| Measures | `list-measures`, `read`, `create-measure`, `update-measure`, `delete-measure` | Read formula previews/raw DAX, create/update/delete measures, and preserve optional format/description metadata. |
| Relationships | `list-relationships`, `read-relationship`, `create-relationship`, `update-relationship`, `delete-relationship` | Inspect joins, create many-to-one joins, and toggle active/inactive state safely. |
| Model and query execution | `read-info`, `refresh`, `evaluate`, `execute-dmv` | Read model summary, refresh a table or whole model, execute bounded `EVALUATE` tabular DAX, and execute metadata DMV queries. |

The upstream [Data Model skill](https://github.com/sbroenne/mcp-server-excel/blob/a77007db43cc10c38da9d81e736fa5eb57af7f14/skills/shared/datamodel.md)
requires tables to be in the Data Model before measures are created. It also
states that worksheet-table edits and the model copy are separate until a
model refresh, while a Power Query refresh synchronizes automatically. DAX
evaluation and DMV execution use the MSOLAP provider; missing provider/model
must become an actionable error rather than an unexplained COM failure.
The same source records Excel limitations: calculated columns and calculated
tables are not exposed through the required COM surface, so they are not
valid parity requirements. DMV results need bounded output, and some
TMSCHEMA views may be empty in Excel.

The upstream [installation guide](https://github.com/sbroenne/mcp-server-excel/blob/a77007db43cc10c38da9d81e736fa5eb57af7f14/docs/INSTALLATION-MCP-SERVER.md)
sets Windows/Excel 2016+ expectations. Issue [#323](https://github.com/sbroenne/mcp-server-excel/issues/323)
and the corresponding source comments document workbook-state failures around
`0x800A03EC`; this is a test case for state-aware recovery, not permission for
blind mutation retries. Issue [#356](https://github.com/sbroenne/mcp-server-excel/issues/356)
describes the EVALUATE use case and proposes a bounded row-return parameter;
the current upstream interface accepts a query but does not expose that
parameter, so ExcelTask should make output limits explicit rather than infer
them from the issue proposal.

## ExcelTask gap inventory

| Capability | Shipped `HEAD` (`23d6fe5`) | Incoming `origin/main` (`6a31ff1`, **unmerged**) | Remaining parity gap |
|---|---|---|---|
| Query discovery, stored M, load configuration | `AuditWorkbookFlows` lists query names and where outputs go; its runtime deliberately never reads M. [`ExcelWorkbookRuntime.Audit.cs`](../src/ExcelTask.Excel/ExcelWorkbookRuntime.Audit.cs) and [`Contracts.cs`](../src/ExcelTask.Core/Contracts.cs) describe names/shapes only. | `ManageQuery` can Plan one named query and return a fingerprinted M expression. It still is not `list`, `view`, or `get-load-config`. | Add bounded query discovery/read/load-config semantics, with a deliberate M-content policy. A hash/length/omitted receipt is safer than accidental unbounded source disclosure. |
| Query create/update/delete | None; no ManageQuery payload in HEAD. | Fingerprinted Create/Replace/Delete for one query; no destination or refresh controls. | Add atomic create/update behavior tied to load state, refresh, cleanup, and precondition reconciliation. |
| M evaluate | None. | None. | Add ephemeral evaluate with bounded columns/rows/bytes/time and cleanup proof; no permanent query on a read-only plan. |
| Load destinations and refresh | None. The README explicitly says current boundary has no Power Query refresh. [`README.md`](../README.md) | None. | Add worksheet/Data Model/both/connection-only load, one/all refresh, unload, target-cell collision checks, and 30-minute-style configurable timeout bounds. |
| Rename/unload and orphan cleanup | None. | None. | Add normalized rename and unload/delete cleanup for worksheet ListObjects, Data Model connections/tables, and renamed connection strings. |
| Model table/column/info discovery | Audit reports model table/relationship/measure names and shapes, but not a dedicated read contract. | No dedicated table/column/info operation. | Add list/read model metadata and explicit model-versus-worksheet distinction; do not treat OOXML `ScanWorkbookStructure` as a model parser. |
| Measures | None. | Fingerprinted Create/Replace/Delete for one measure; no list/read/format/description operation. | Add bounded listing/readback and metadata mutation, with explicit DAX content policy and dependency checks. |
| Relationships | None. | Create/Delete one many-to-one relationship; no list/read/update-active operation. | Add list/read and active-state update, enforcing compatible columns and relationship impact checks. |
| DAX EVALUATE and DMV | None. | None. | Add read-only, bounded DAX/DMV execution with MSOLAP/provider diagnostics and empty-result handling. |
| Model refresh/synchronization | None. | None. | Add whole-model/table refresh and verify worksheet-to-model synchronization after relevant edits. |
| Optional DAX-backed materialization | None. | None. | Consider only if measured demand requires creating a worksheet table from DAX; do not promise calculated columns/tables that upstream itself marks unsupported. |

`ScanWorkbookStructure` intentionally omits queries, connections, and the
Data Model because they are not safely represented as ordinary OOXML tables;
the local [architecture](../docs/ARCHITECTURE.md) calls out Power Query's
nested/base64 package and the model's opaque representation. Extending that
headless scan is not a substitute for live Excel COM behavior.

## Prioritized implementation and verification list

### P0 — safe foundations and read-only parity

1. **Keep one-task boundaries and reconcile the merge.** If the incoming
   fingerprinted ManageQuery/Measure/Relationship design is accepted, resolve
   its conflict before extending it. Reconcile its intentional Plan M/DAX
   readback with the local privacy/receipt contract; do not silently broaden
   content exposure.
2. **Add bounded discovery.** Implement query list/view/load-config and model
   table/column/info/measure/relationship reads as task steps, not a second
   tool catalog. Return counts, names, hashes, and lengths by default; expose
   M/DAX only when the caller explicitly requests it and the contract permits
   it. Use fixed response byte/row/column limits.
3. **Add ephemeral evaluators.** M `evaluate`, DAX `EVALUATE`, and DMV execution
   should be read-only plans with hard row/column/byte/time caps, temporary
   object cleanup, and actionable `Rejected` results for missing MSOLAP/model
   prerequisites. Never retry an uncertain mutation dispatch.

### P1 — lifecycle and model correctness

4. **Implement Power Query lifecycle.** Add Create/Update with explicit load
   destination, `LoadTo`, Refresh/RefreshAll, Rename, Unload, and Delete. Test
   existing-sheet collisions and target-cell requirements; verify orphan table,
   connection, and Data Model cleanup; distinguish transient `0x800A03EC` from
   persistent workbook-state refusal. Preserve M exactly by default.
5. **Implement model synchronization and metadata.** Add whole-model/table
   refresh, table/column reads, measure metadata (format/description), and
   relationship list/read/active update. Enforce Data Model prerequisites,
   many-to-one direction, compatible column types, one active path per table
   pair, and destructive table/measure dependency checks.
6. **Make formatting and locale behavior explicit.** Preserve M/DAX source;
   remote formatter calls must be opt-in, privacy-gated, and fall back to the
   original source on failure. Test locale separator translation without
   changing formula meaning. Do not add remote network behavior as a hidden
   default.

### P2 — adjacent, evidence-driven features

7. Add DAX-backed worksheet-table materialization only if a measured workflow
   needs it, with the same bounded evaluation and save/reopen proof. Keep
   calculated columns/tables out of scope because the target documents the
   Excel COM limitation.
8. Consider source-control export/import recipes (`.pq`/DAX) as documentation
   only after demand is demonstrated; they are not required for runtime parity.

For all priorities, verification must include pure contract/inference tests,
in-process MCP transport/schema tests, and serial disposable-workbook tests on
real desktop Excel. The local contract requires save/reopen verification and
owned-process/file-lock assertions; a build alone cannot prove workbook
correctness. Test Plan non-mutation, Apply fingerprint recheck, timeouts,
privacy/credential dialogs, missing MSOLAP, empty DMV views, sheet collisions,
stale query/measure preconditions, relationship cardinality, orphan cleanup,
and the `Unknown` reconciliation path. The merging/managed work computer is
the field truth for Excel policy and authentication, per the local
[agent-bridge workflow](../docs/AGENT-BRIDGE.md).

## Deliberate boundaries

- Do not copy the upstream 26 tools/234 operations or add a broad CLI. Keep the
  normal model-facing surface as one deep `excel_task` task, as required by
  [README](../README.md) and [MVP contract](../docs/MVP-CONTRACT.md).
- Do not use OOXML parsing as a fake Power Query/Data Model implementation.
  Those objects require live Excel behavior and reconciliation.
- Do not claim calculated columns/tables as required features; upstream marks
  them unsupported through the relevant Excel COM API.
- Do not send M/DAX to remote formatters without explicit consent and a
  privacy/content-boundary decision.
- Do not blind-retry after mutation dispatch or kill a user-owned Excel
  process; preserve the local `Unknown`, ownership, and cleanup invariants.

## Primary sources

- Upstream [`FEATURES.md`](https://github.com/sbroenne/mcp-server-excel/blob/a77007db43cc10c38da9d81e736fa5eb57af7f14/FEATURES.md),
  [`IPowerQueryCommands.cs`](https://github.com/sbroenne/mcp-server-excel/blob/a77007db43cc10c38da9d81e736fa5eb57af7f14/src/ExcelMcp.Core/Commands/PowerQuery/IPowerQueryCommands.cs),
  [`IDataModelCommands.cs`](https://github.com/sbroenne/mcp-server-excel/blob/a77007db43cc10c38da9d81e736fa5eb57af7f14/src/ExcelMcp.Core/Commands/DataModel/IDataModelCommands.cs),
  and [`IDataModelRelCommands.cs`](https://github.com/sbroenne/mcp-server-excel/blob/a77007db43cc10c38da9d81e736fa5eb57af7f14/src/ExcelMcp.Core/Commands/DataModel/IDataModelRelCommands.cs).
- Upstream [Power Query skill](https://github.com/sbroenne/mcp-server-excel/blob/a77007db43cc10c38da9d81e736fa5eb57af7f14/skills/shared/powerquery.md),
  [Data Model skill](https://github.com/sbroenne/mcp-server-excel/blob/a77007db43cc10c38da9d81e736fa5eb57af7f14/skills/shared/datamodel.md),
  [COM behavior findings](https://github.com/sbroenne/mcp-server-excel/blob/a77007db43cc10c38da9d81e736fa5eb57af7f14/docs/COM-API-BEHAVIOR-FINDINGS.md),
  [installation guide](https://github.com/sbroenne/mcp-server-excel/blob/a77007db43cc10c38da9d81e736fa5eb57af7f14/docs/INSTALLATION-MCP-SERVER.md),
  [issue #323](https://github.com/sbroenne/mcp-server-excel/issues/323), and
  [issue #356](https://github.com/sbroenne/mcp-server-excel/issues/356).
- Local [AGENTS.md](../AGENTS.md), [README.md](../README.md), [MVP contract](../docs/MVP-CONTRACT.md),
  [architecture](../docs/ARCHITECTURE.md), [Contracts.cs](../src/ExcelTask.Core/Contracts.cs),
  and [ExcelWorkbookRuntime.Audit.cs](../src/ExcelTask.Excel/ExcelWorkbookRuntime.Audit.cs).
