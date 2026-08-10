# ExcelTask operating contract

ExcelTask is a clean-sheet, Copilot-first Excel automation product. It is not a
compatibility rewrite of ExcelMcp. Do not add legacy command parity, a broad
CLI, model selection, or model SDK dependencies.

## Product contract

- The normal model-facing interface is one deep `excel_task` tool.
- The selected MCP client model remains in charge; the server is deterministic.
- One task owns inspect, plan, execute, verify, save, and cleanup.
- Initial releases master formula/exhibit work before adding other workflows.
- Multiple workbooks and reference worksheets are first-class.
- If the target workbook is already open, ask before attaching to that exact
  live workbook. Never guess from `ActiveWorkbook`.
- Overwriting an existing workbook requires explicit authorization.

## Runtime invariants

- Only the owning STA thread touches Excel COM objects.
- Never share an RCW across threads or retain one past its worker lifetime.
- The MCP host never owns COM. Each operation uses one private worker; only
  that worker may recover the Excel process it created.
- A timeout after mutation dispatch is `Unknown`; never retry it blindly.
- Return `Completed`, `Partial`, `Unknown`, `Rejected`, or
  `NeedsConfirmation` truthfully.
- Reads, task steps, queue depth, cells, elapsed time, and response bytes are
  bounded.
- Close only owned workbooks. Never quit or kill a user-owned Excel process.
- For an owned process: save, close, quit, wait, verify process exit and file
  lock release, then use identity-checked termination only as recovery.
- Logs never contain workbook values, formulas, prompts, credentials,
  connection strings, or customer identifiers. Receipts may identify only the
  request-supplied workbook filename, worksheet, and bounded repair ranges;
  they contain counts, checks, and fingerprints rather than cell contents or
  formula text.

## Verification

- Pure planning and formula inference are tested without Excel.
- MCP contract tests use in-process stream transport.
- COM behavior uses disposable workbook copies and serial Windows tests.
- A build is not proof of workbook correctness. The MVP requires save/reopen
  verification and an owned-process/file-lock assertion.
- Do not run Excel integration tests concurrently.

## Release discipline

- `main` is the current development and release branch.
- Keep the repository current: every user-visible change updates
  `CHANGELOG.md`, and updates `ROADMAP.md` when roadmap status changes.
- After a verified user-visible change set, create a stable, immutable SemVer
  release from `main`.
- Do not add tools, CLI commands, or documentation beyond the smallest
  evidence-backed need.
