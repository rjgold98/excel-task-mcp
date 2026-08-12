# Claude entry point

Read and follow [AGENTS.md](AGENTS.md) before doing any work in this repository.
It is the authoritative product, safety, verification, and release contract.

For owner-approved local collaboration with Codex on an opt-in pull request,
also follow [docs/AGENT-RELAY.md](docs/AGENT-RELAY.md). That relay is separate
from the managed-work-computer field-agent channel in
[docs/AGENT-BRIDGE.md](docs/AGENT-BRIDGE.md).

When invoked by the relay, act as a read-only reviewer and planner. Return
evidence-backed findings or a bounded task request; do not edit files, write to
GitHub, merge, release, or operate Excel. The relay posts your response so it
can apply replay protection and preserve one auditable conversation thread.
