# ExcelTask project-skill comparison

**Question.** Should ExcelTask ship a project skill analogous to
[sbroenne/mcp-server-excel's `excel-mcp` skill](https://github.com/sbroenne/mcp-server-excel/blob/a77007db43cc10c38da9d81e736fa5eb57af7f14/skills/excel-mcp/SKILL.md)?

**Evidence date.** 2026-08-13. The upstream file was fetched from the pinned
commit `a77007db43cc10c38da9d81e736fa5eb57af7f14` (HTTP 200; 7,392 UTF-8
bytes; 190 lines; SHA-256
`36c571469eda3554722a292ee7756a2d9c7e0fa4f44df8664fbf28df24653175`). Local
facts below come from this checkout, including its [operating contract](../AGENTS.md),
[README](../README.md), [MVP contract](../docs/MVP-CONTRACT.md),
[architecture](../docs/ARCHITECTURE.md), [decisions](../docs/DECISIONS.md),
[interface study](../docs/INTERFACE-AB-STUDY.md), and [field-check workflow](../docs/FIELD-CHECK.md).

## Recommendation

**Do not copy the upstream skill or add a broad specialist pack to the MVP.**
ExcelTask should have a **small, project-level client guidance skill for
GitHub-Copilot-compatible clients**, placed at the documented
`.github/skills/excel-task/SKILL.md` path, subject to a fresh-session A/B on the
managed machine. Treat it as a thin routing and safety layer, not as a second
tool catalog or a compatibility facade. GitHub's documentation says repository
skills are available to the Copilot app, VS Code agent mode, CLI, and cloud
agent, but the agent still loads a skill only when its description is relevant;
other MCP clients do not acquire skills through the MCP protocol. The canonical
source remains the tool schema plus `README.md`/`docs/` when the client does not
support this standard.

This recommendation preserves the one-tool/token strategy while addressing the
real problem a skill can solve: carrying the workflow rules that are easy for a
model to miss between sessions. It also avoids importing upstream instructions
that would make ExcelTask choose the wrong workbook or claim capabilities it
does not have.

## What the upstream skill actually does

The pinned upstream [`SKILL.md`](https://github.com/sbroenne/mcp-server-excel/blob/a77007db43cc10c38da9d81e736fa5eb57af7f14/skills/excel-mcp/SKILL.md)
is an agent-facing index for a rich, session-oriented server:

- Its front matter advertises Excel/workbook/file/formula/PQ/DAX/PivotTable/
  chart/VBA triggers and says the server provides **234 operations**.
- Its workflow begins with `file open/create`, performs worksheet/range/table
  work, and ends with `file close(save: true)`.
- Its preconditions assume Windows Excel 2016+, full paths, and that a workbook
  is not open in another Excel instance.
- Its nine execution rules tell the model not to ask clarifying questions,
  discover an open session, format values, convert tabular ranges to tables,
  use manual calculation for bulk writes, evaluate Power Query before
  persisting, prefer targeted updates, and follow `suggestedNextActions`.
- It points to many detailed `references/` files. The companion
  [`skills/excel-mcp/README.md`](https://github.com/sbroenne/mcp-server-excel/blob/a77007db43cc10c38da9d81e736fa5eb57af7f14/skills/excel-mcp/README.md)
  describes installation for Copilot, Claude Code, Cursor, Codex, and other
  clients, so this is a reusable agent artifact rather than server code.

Those choices are coherent for that server. They are not a safe default for
ExcelTask.

## ExcelTask facts that control the decision

| Local evidence | Consequence for a project skill |
|---|---|
| The product contract says the normal model-facing interface is exactly one deep `excel_task` tool; the selected client model remains in charge; there is no broad CLI or compatibility rewrite. ([`AGENTS.md`](../AGENTS.md), [`README.md`](../README.md)) | A skill must explain one tool and its operation union. It must not recreate a 234-operation catalog or choose a model. |
| A request supplies exactly one of the bounded operations (`CopyExhibit`, formula repair/extension, macro edit, audit/range read, constant/formula write, find/replace, create, number format, or structure scan). ([`docs/MVP-CONTRACT.md`](../docs/MVP-CONTRACT.md), [`src/ExcelTask.Core/Contracts.cs`](../src/ExcelTask.Core/Contracts.cs)) | Link to the canonical operation descriptions; do not duplicate a second schema that can drift. |
| `AskIfOpen` is the default. If the target may be open, ExcelTask returns `NeedsConfirmation`; the caller must choose `UseOpen` for that exact live workbook or `Isolated` for a private file instance. The contract explicitly forbids guessing from `ActiveWorkbook`. ([`AGENTS.md`](../AGENTS.md), [`README.md`](../README.md), [`docs/ARCHITECTURE.md`](../docs/ARCHITECTURE.md)) | The skill must permit and explain a clarifying confirmation. Upstream's “NEVER Ask Clarifying Questions” and “use the open session” rules would be unsafe here. |
| `Plan` is non-mutating; `Apply` performs required confirmations, saves, reopens, and verifies. A timeout or failure after mutation dispatch is `Unknown`, and an `Unknown` mutation must be reconciled before retry. ([`docs/MVP-CONTRACT.md`](../docs/MVP-CONTRACT.md), [`docs/ARCHITECTURE.md`](../docs/ARCHITECTURE.md)) | A skill can provide a short Plan → Apply → reconcile decision tree. It must not imply a persistent session or blind retry. |
| The MCP tool is bounded again at the response seam (30 KiB), returns structured receipts, and withholds workbook values/formula text/VBA source except where the explicit range read or Plan contract permits it. ([`src/ExcelTask.McpServer/ExcelTaskTool.cs`](../src/ExcelTask.McpServer/ExcelTaskTool.cs), [`AGENTS.md`](../AGENTS.md)) | Include privacy and receipt-handling reminders; never ask the model to paste formulas, values, prompts, or credentials into logs or summaries. |
| `ScanWorkbookStructure` reads the OOXML package without starting Excel; Excel-dependent reads/mutations retain live COM, save/reopen verification, and owned-process cleanup. ([`README.md`](../README.md), [`docs/ARCHITECTURE.md`](../docs/ARCHITECTURE.md)) | Explain when to use structure scan versus `AuditWorkbookFlows`/Excel operations. Do not promise the headless breadth of openpyxl or file-format fidelity for mutations. |
| Direct formulas are deliberately separate from constant writes; formula writes are bounded and read back before save/reopen. Macro edits are isolated `.xlsm` + `Copy`, hash-guarded, and trust-policy dependent. ([`docs/MVP-CONTRACT.md`](../docs/MVP-CONTRACT.md), [`README.md`](../README.md)) | The skill should call out the value/formula split and macro restrictions, but should not teach arbitrary VBA or low-level COM. |
| The project has no existing repository `SKILL.md`/skills directory; its canonical guidance is README plus `docs/`. (`rg --files .github .vscode docs research`) | Do not assume a path. Confirm the client loader before adding one; if added, keep it one small file and link to canonical docs. |
| The current in-process MCP contract test requires exactly one listed tool and caps the serialized tool payload at **16 KiB** (`1..16 * 1024`); it also asserts that `session`, `confirmationToken`, and other out-of-contract fields are absent. ([`tests/ExcelTask.McpServer.Tests/ExcelTaskToolProtocolTests.cs`](../tests/ExcelTask.McpServer.Tests/ExcelTaskToolProtocolTests.cs)) | Treat the 16 KiB budget as a schema/runtime boundary. A companion skill must not be used to smuggle extra operation fields or replace schema-enforced rules; keep only workflow guidance that does not fit the bounded schema. |

## Direct conflicts to avoid

The upstream instructions would create concrete ExcelTask failures if copied
verbatim:

1. **“Never ask clarifying questions”** conflicts with the explicit open-workbook
   confirmation gate. Discovering a filename or an active workbook is not proof
   of exact identity, and choosing `UseOpen` controls a user's live Excel
   process.
2. **`file`/`sessionId` open-close choreography** does not exist in the
   `excel_task` API. Each call owns a private worker and Excel lifecycle; a
   skill that invents sessions would cause invalid calls and conceal the
   `Unknown` boundary.
3. **“Files must not be open in another Excel instance”** is too broad. ExcelTask
   supports an explicitly confirmed `UseOpen` path and a separate `Isolated`
   path. The distinction is the safety contract.
4. **Manual calculation as a generic optimization** is not a client choice.
   The repository's real-Excel measurements found the shipping batched automatic
   path faster (76 ms versus 103 ms for the tested repair), and the server owns
   calculation-state restoration. ([`docs/EXCEL-TUNING.md`](../docs/EXCEL-TUNING.md))
5. **Tables, Power Query, DAX, PivotTables, charts, slicers, screenshots, and
   broad formatting** are not current ExcelTask operations. A skill that
   advertises them would violate the clean-sheet boundary and make the model
   issue rejected calls.
6. **Persistent-session cleanup advice** is the wrong abstraction. ExcelTask's
   worker supervises save, close, process-exit proof, lock release, reopen
   verification, and truthful status; the model should report the receipt, not
   manage COM objects.

## What a minimal ExcelTask skill should contain

If the target client confirms a project-skill discovery path, one compact file
could contain only these model-facing rules:

- Trigger on ExcelTask/`excel_task`/workbook/formula/worksheet/macro requests.
- Use exactly one `excel_task` call with one matching operation payload. Start
  with `AskIfOpen` unless the user has already explicitly identified the exact
  open workbook; honor `NeedsConfirmation` and ask for `UseOpen` versus
  `Isolated` when required.
- Use `Plan` for a preview; use `Apply` only after the requested binding and
  overwrite authorization are explicit. For an `Unknown` mutation, perform a
  fresh inspect/plan reconciliation before any retry.
- When names are unknown, use `AuditWorkbookFlows` (Excel-backed report) or
  `ScanWorkbookStructure` (bounded OOXML structure scan) as appropriate. Keep
  reference and target workbooks explicit.
- Keep constants and formulas on their separate operations; prefer inferred
  repair/extension when neighboring evidence proves the intended pattern. Keep
  macro edits isolated, `.xlsm` + `Copy`, hash-guarded, and trust-policy aware.
- Read the structured receipt's status/checks and give a short text summary;
  never expose workbook values, formula text, VBA source, credentials, or full
  customer paths in diagnostics.
- Link to [`README.md`](../README.md) and [`docs/MVP-CONTRACT.md`](../docs/MVP-CONTRACT.md)
  for the current operation/bound details instead of copying every numeric cap.

It should **not** include a second operation catalog, active-workbook guessing,
manual calculation toggles, table/PQ/DAX recipes, low-level COM instructions,
model selection, or a persistent confirmation/session token.

## Rollout and validation gate

No client loader is configured in this checkout (`.vscode/mcp.json` configures
only the MCP executable). The upstream companion README says its own VS Code
extension installs a *personal* skill to `~/.copilot/skills/excel-mcp/` when
`chat.useAgentSkills` is enabled. Current GitHub documentation establishes the
project-skill mechanism for Copilot surfaces: repository skills live in
`.github/skills/<name>/SKILL.md` (also `.claude/skills` or `.agents/skills`), and
the Copilot app says skills configured for repositories are automatically
available there. Copilot still injects `SKILL.md` only when it decides the skill
is relevant. VS Code's documented defaults enable `chat.useAgentSkills` and
search `.github/skills`, `.claude/skills`, `~/.copilot/skills`, and
`~/.claude/skills`. See [About agent skills](https://docs.github.com/en/copilot/concepts/agents/about-agent-skills),
[Adding agent skills](https://docs.github.com/en/copilot/how-tos/copilot-on-github/customize-copilot/customize-cloud-agent/add-skills),
[Customizing the Copilot app](https://docs.github.com/en/copilot/how-tos/github-copilot-app/customize-github-copilot-app),
and [VS Code's agent-skills settings](https://code.visualstudio.com/docs/agents/reference/ai-settings).
These are client features, not MCP protocol behavior: the server cannot force
an arbitrary MCP client to load a skill. The target managed Copilot app/session
still needs a fresh-session check (and organization policy may disable skills).

Before shipping one, verify the actual target-client discovery path and enablement
behavior in a fresh session, then run an A/B using the existing
[`docs/FIELD-CHECK.md`](../docs/FIELD-CHECK.md) measures: first call, tool calls,
corrections/retries, wall time, and correctness after reopening the disposable
workbook. The A/B should specifically cover an open-workbook confirmation, a
Plan → Apply mutation, an `Unknown`/reconciliation path if safely reproducible,
and a formula-versus-constant request. Keep the skill only if it reduces wasted
calls or invalid requests without changing the receipt status or workbook
correctness. The repository's interface study already shows why this gate is
worth doing: terse guidance dropped clean-task accuracy to 0.17–0.28, and
stating enforced limits saved 32% of calls in the measured harness
([`docs/INTERFACE-AB-STUDY.md`](../docs/INTERFACE-AB-STUDY.md)).

If the skill becomes a shipped user-facing artifact, update `CHANGELOG.md` as
required by [`AGENTS.md`](../AGENTS.md), and keep the skill's references aligned
with the canonical schema/docs. Do not add product operations or broaden the
roadmap as a side effect.

**Disposition:** adopt a thin `.github/skills/excel-task/SKILL.md` companion for
GitHub Copilot surfaces as a measured client-guidance experiment; decline an
upstream-style broad skill, and do not treat the companion as schema/runtime
enforcement for arbitrary MCP clients.
