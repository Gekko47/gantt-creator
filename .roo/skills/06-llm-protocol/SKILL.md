---
name: 06-llm-protocol
description: Required session opening, evidence ledger, anti-drift and anti-hallucination controls, the two-attempt no-loop protocol, context discipline, coding behaviour, prompt pattern, the agent handoff format, and the full-repo code-review methodology for the Gantt Creator repository. Part of the gantt-creator index. Trigger on protocol, evidence, retry, handoff-format, or code-review-methodology questions.
---

# LLM operating protocol

Mirror of [`.cline/skills/06-llm-protocol/SKILL.md`](../../../.cline/skills/06-llm-protocol/SKILL.md),
generated from [`docs/06-LLM-PROTOCOL.md`](../../../docs/06-LLM-PROTOCOL.md). The canonical doc
wins on any disagreement.

## Index

- Index first: `gantt-creator-contract`; `00-ponytail` before any code change.
- This skill owns the session-opening format and the handoff format the other skills assume.

## Session opening (seven lines, before any edit)

Outcome · files/layers likely involved · explicit exclusions · acceptance tests · facts
already verified · remaining unknowns · first safe action. If that is wider than the work item,
correct it before editing.

## Do not get wrong

- Two failed attempts on the same failure = stop and report; never a third unchanged retry —
  except exactly one recorded retry for a documented flaky external Office operation, which
  becomes a defect if it passes only on retry.
- An inference cannot become a fact by being repeated — use the
  `fact` / `inference` / `proposal` / `unknown` ledger honestly.
- Never fabricate command output, tool names, COM members, Excel-DNA callbacks, RibbonX
  attributes, package versions, or decisions. If a check was not run, say `Not run` and give
  the exact command.
- A full-repo review must look for gate-integrity, build-pipeline drift, docs/code drift, and
  local/CI divergence even when the per-item checklists pass — those are the cross-cutting
  defect classes per-item gates cannot catch by construction.
- Prohibited "fixes": repeated unchanged commands, larger sleeps, broad catches, deleted tests,
  lowered thresholds, suppressed warnings, package churn, random API substitution.

## Handoff format

```text
Outcome: <what now works>
Changed: <2-5 important files/areas>
Proof: <commands and observed results>
Office/manual: <PASS, FAIL, or Not run: reason>
Risk: <one sentence or None known>
Deferred: <explicitly excluded work>
Next: <roadmap ID>
```

## Tools here

`read_file`, `codebase_search`, `search_files` (inspect before editing) · `execute_command`
(build, test, gates, `git`) · `context7` / `microsoft-learn` (primary API documentation before
memory). `memra_*` is unavailable in this session: record durable facts in
`docs/STATUS.md` instead, and never claim a memra write.
