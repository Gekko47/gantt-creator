---
name: gantt-creator-contract
description: Index skill for the Gantt Creator Excel add-in repository, and the first skill any task must call. Routes to the eight adopted skills (00-ponytail, 01-environment, 02-architecture, 03-roadmap, 04-test-strategy, 05-git-quality, 06-llm-protocol, 07-entity-guide), names their canonical docs/0N-*.md sources, states the non-negotiable Excel/scene/COM invariants, and lists the verification gates. Use at session start, or when deciding which skill, work item, gate, or entity/worksheet rule applies to Gantt, scene, renderer, COM, Excel, PowerPoint, export, test, commit, or docs work.
---

# Gantt Creator index and calling contract

The kit has three layers: `AGENTS.md` (always-on contract), `docs/0N-*.md` (canonical
detail), and the skills (`docs/*` summaries, generated into `.cline/skills/`, mirrored in
`.roo/skills/` for this mode). This file is a **router**: it holds no copy of the contract,
so it cannot drift from it.

## When to use

- Session start, and whenever the task area changes.
- When deciding which skill, work item, gate, or entity/worksheet rule governs a change.

## When NOT to use

- Product requirements — read `AGENTS.md` and the canonical doc, this file only routes.
- Authoring or auditing Agent Skills packages (that is the `skill-writer` mode's own job).

## Calling procedure (mandatory)

1. **Call this skill first.** It is the index; do not guess a skill from memory.
2. **Call `00-ponytail` before touching code** — it is the driving skill and outranks the
   others: understand the flow, then take the smallest diff that works.
3. **Call only the topic skills the task needs**, in the order below. Do not bulk-load.
4. When a summary is not enough, **read the canonical doc** in that skill's
   `Canonical source` line. A summary is never the full contract.
5. If a skill and `docs/` disagree, **stop and report the conflict**.

The same procedure is stated always-on in `.roo/rules/00-gantt-creator-index.md`, so it
applies even when this skill was not selected.

## Index

| Skill (`.roo/skills/<name>/SKILL.md`) | Canonical source | Call when |
|---|---|---|
| `00-ponytail` | `docs/00-PONYTAIL.md` | Always, before any code change |
| `01-environment` | `docs/01-ENVIRONMENT.md` | Setup, versions, install, Office-host failures |
| `02-architecture` | `docs/02-ARCHITECTURE.md` | Worksheet schema, scene model, COM, C# conventions |
| `03-roadmap` | `docs/03-ROADMAP.md` | Which work item, phase, gate, or next safe action |
| `04-test-strategy` | `docs/04-TEST-STRATEGY.md` | Writing/running tests, coverage, golden images |
| `05-git-quality` | `docs/05-GIT-QUALITY.md` | Commit, PR, dependency, release |
| `06-llm-protocol` | `docs/06-LLM-PROTOCOL.md` | Session opening, evidence, retries, handoff |
| `07-entity-guide` | `docs/07-GANTT-ENTITY-GUIDE.md` | Any visual, layout, style, label, or export change |
| (no skill) | `docs/08-TEST-CHECKLIST.md` | Self-certifying sections A–N before commit |

Read `docs/08-TEST-CHECKLIST.md` directly: it is an always-on checklist with no generated
skill and no `SKILL-SUMMARY` block.

## Non-negotiables (condensed; `AGENTS.md` is authoritative)

- Exactly one visible worksheet plus one `xlSheetVeryHidden` `_GanttCreatorConfig`; never a
  second helper sheet; never schedule rows, shapes, or logs on the VeryHidden sheet.
- `Core` has no Office, Excel-DNA, SkiaSharp, clipboard, file-dialog, or UI dependency;
  never call Excel/PowerPoint from `Core`.
- The scene model owns all positioning; renderers consume the scene and never recalculate
  dates, labels, colours, widths, or z-order.
- `Type` values come only from `EntityTypeCatalog`. Undefined visual behaviour is `unknown`
  — stop for a decision, never infer it from a screenshot.
- Milestones and delineators read `Start` only; `Finish` is never their geometry.
- Every new `if (bad) { error }` validator ships a positive test in the same commit.
- No `Thread.Sleep`, no `Task.Run` around COM, no analyzer suppression, no catch-and-ignore,
  no test deletion or threshold loosening to force green.
- Deterministic identifiers, ordering, culture, time, and rounding; core operation is offline.
- Chart-affecting worksheet edits and successful chart-setting commands schedule a coalesced
  automatic refresh; selection/navigation changes never render.

## Workflow

1. Read `AGENTS.md`, then [`docs/REPO-MAP.md`](../../../docs/REPO-MAP.md).
2. Name the work item in `docs/work-items/`. If it does not exist, stop and ask.
3. Call `00-ponytail` and trace the real flow before choosing a rung of the ladder.
4. Session opening per `06-llm-protocol`: outcome, files/layers, exclusions, acceptance
   tests, verified facts, unknowns, first safe action.
5. Smallest coherent change proving one behaviour; tests in the same commit.
6. Targeted tests → `verify-quick.ps1` (staged) → `verify.ps1` (clean tree, last commit).
7. Update `docs/STATUS.md`, and `docs/REPO-MAP.md` when a mapped boundary, entry point, test
   location, or command changed.
8. Hand off as `Outcome / Changed / Proof / Office / Risk / Deferred / Next`; any gate not run
   is `Not run` with the reason.

## Verification commands

```text
pwsh -NoProfile -File scripts/verify-quick.ps1    # every-commit gate
pwsh -NoProfile -File scripts/verify.ps1          # branch-final gate (clean tree)
pwsh -NoProfile -File scripts/verify-office.ps1   # live-Excel gate (human host)
pwsh -NoProfile -File scripts/check-cline-skills.ps1
pwsh -NoProfile -File scripts/check-md-links.ps1
pwsh -NoProfile -File scripts/check-status.ps1
```

## Tool mapping for this mode

The kit's `SKILL-TOOLS` blocks name the Cline MCP surface. Only `context7` and
`microsoft-learn` are configured in [`.roo/mcp.json`](../../mcp.json); map by capability:

| Kit asks for | Use here |
|---|---|
| `search_codebase` / `grep_files` | `codebase_search`, `search_files` |
| `read_files` / `read_file` | `read_file` |
| `run_commands` | `execute_command` |
| `dotnet_build` / `dotnet_test` | `execute_command` with `dotnet build` / `dotnet test` |
| `vscode-mcp__get_diagnostics` | `execute_command` with `dotnet build -warnaserror` |
| `git_tool` | `execute_command` with `git` |
| `memra_*` | not available — record the fact in `docs/STATUS.md` |

Never claim output from a tool that is not available.

## Drift to report, not silently fix

Verified 2026-11-01; all belong to the kit, not to this router:

- No `.cline/rules*` tree exists — always-on rules live in `AGENTS.md`, while
  `docs/adr/0000-adopt-kit.md` and `docs/adr/0002-commit-clinerules-and-agents.md` still
  describe a committed `.clinerules/*` loader view (stale ADRs).
- Generated frontmatter preserves the source file's case (`name: 00-PONYTAIL`) while the
  directory is lowercase (`00-ponytail`), breaking the "name matches directory" rule; the fix
  belongs in `scripts/sync-cline-skills.ps1` (`$fmName` should use `$entry.Name`).
- The unmapped-doc fail-fast in `scripts/sync-cline-skills.ps1` filters `docs/0N-*.md`, which
  matches no file (`dir docs\0N-*.md` → `File Not Found`); a new `docs/09-*.md` would evade
  the gate. Intended pattern: `0[0-9]-*.md`.
