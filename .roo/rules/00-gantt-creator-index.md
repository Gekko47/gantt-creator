# Gantt Creator — index calling rule (always on)

This file is loaded on every turn. It is a **dispatch rule**, not a second copy of the
contract: [`AGENTS.md`](../../AGENTS.md) is the always-on contract and always wins. Detail
lives in `docs/0N-*.md` and the skill bodies under `.roo/skills/` and `.cline/skills/`.

## 1. Index calling is mandatory, in this order

1. **Call the index skill** `gantt-creator-contract` at session start and whenever the task
   area changes. It routes; do not guess a skill from memory.
2. **Call `00-ponytail`** before writing or changing any code. It is the driving skill and
   outranks the others: understand the flow first, then take the smallest diff that works.
3. **Call only the topic skills the task needs.** Do not bulk-load all eight.
4. When a skill's summary is not enough, read the canonical doc named in that skill's
   `Canonical source` line. Never treat a skill summary as the full contract.
5. If two skills, or a skill and `docs/`, disagree, **stop and report the conflict**; do not
   silently pick one.

## 2. Index table — trigger to skill

| Trigger | Skill to call |
|---|---|
| Session start, "how small should this diff be", before any code | `00-ponytail` |
| Setup, install, SDK/Office/Visual Studio/Excel-DNA versions, host failures | `01-environment` |
| Worksheet schema, `_GanttCreatorConfig`, scene model, COM, C# conventions | `02-architecture` |
| Which work item / phase / gate, next safe action, scope change | `03-roadmap` |
| Writing or running tests, coverage, golden images, Office harness | `04-test-strategy` |
| Commit, PR, dependency, release, branch policy | `05-git-quality` |
| Session opening, evidence ledger, retries, handoff format, full-repo review | `06-llm-protocol` |
| Any visual, layout, style, label, or export change | `07-entity-guide` |
| Self-certifying checklist sections A–N before commit | read `docs/08-TEST-CHECKLIST.md` directly |

## 3. Hard constraints that need no lookup

The complete invariant list is `AGENTS.md`, loaded on every turn; this is the subset agents get
wrong most often or that stale sources still contradict. Do not read it as exhaustive.

- Exactly one visible worksheet plus one `xlSheetVeryHidden` `_GanttCreatorConfig`; never a
  second helper sheet; never schedule data, shapes, or logs on the VeryHidden sheet.
- `Core` has no Office/Excel-DNA/SkiaSharp/clipboard/dialog/UI dependency; never call Excel
  or PowerPoint from `Core`.
- Scene-first: all positioning comes from the scene model; renderers consume, never
  recalculate dates, labels, colours, widths, or z-order.
- Automatic refresh is COALESCED, not "explicit Refresh only": chart-affecting worksheet edits
  and successful chart-setting or property commands schedule a refresh; selection and
  navigation changes never render. Every automatic or explicit render reads and validates the
  full table and configuration and builds the complete scene before any shape mutation
  (R5.11, ADR-0050). `.github/copilot-instructions.md` still says "only the explicit Refresh
  command rebuilds the chart" — that line is stale; `AGENTS.md` and
  `docs/02-ARCHITECTURE.md` win.
- `Type` values come only from `EntityTypeCatalog`. Undefined visual behaviour is `unknown`
  — stop for a decision, never infer it from a screenshot.
- Milestones and delineators read `Start` only.
- Every new `if (bad) { error }` validator ships a positive test in the same commit.
- No `Thread.Sleep`, no `Task.Run` around COM, no analyzer suppression, no catch-and-ignore,
  no deleting or weakening tests to get green.
- Deterministic identifiers, ordering, culture, time, and rounding. Core operation is offline.
- One work item in progress; one concern per commit; Conventional Commit prefixes only.

## 4. Gates

```text
pwsh -NoProfile -File scripts/verify-quick.ps1    # every commit (accepts staged changes)
pwsh -NoProfile -File scripts/verify.ps1          # branch-final, clean tree, after last commit
pwsh -NoProfile -File scripts/verify-office.ps1   # live Excel gate (human host)
pwsh -NoProfile -File scripts/check-cline-skills.ps1  # .cline/skills vs docs drift gate
pwsh -NoProfile -File scripts/check-md-links.ps1  # doc links
pwsh -NoProfile -File scripts/check-status.ps1    # STATUS accuracy
```

Report `Not run` with a reason for any gate not executed. Never state a command ran, or a
test passed, without observed output.

## 5. Evidence and retries

Keep a `fact` / `inference` / `proposal` / `unknown` ledger; an inference repeated is still an
inference. Two failed attempts on the same failure = stop and report (the sole exception is
one recorded retry of a documented flaky Office operation). Finish with the handoff format:
`Outcome / Changed / Proof / Office / Risk / Deferred / Next`.

Stop and ask the human under the `AGENTS.md` stop conditions: missing or contradictory
acceptance criteria, an invariant violation, a workbook-schema or export-fidelity change, a new
production dependency or Office permission, a failure that exposes unrelated behaviour,
destructive migration or deletion, or credentials, signing secrets, customer data, or
proprietary fonts.

## 6. Where the skills live, and how to edit them

- `docs/0N-*.md` — **the only hand-edited view.** Every fixed decision lives here.
- `.cline/skills/<name>/SKILL.md` — generated from `docs/` by
  `scripts/sync-cline-skills.ps1` and gated by `scripts/check-cline-skills.ps1`.
  **Never hand-edit these.** Edit the canonical doc, run the sync, and commit `docs/` with
  `.cline/skills/` together.
- `.roo/skills/<name>/SKILL.md` — this agent surface's skills: `gantt-creator-contract` is the
  index, then `00-ponytail` … `07-entity-guide`. Hand-maintained mirrors with **no drift gate**;
  refresh them by hand when a canonical doc changes. If a `.roo` skill and the `.cline` view
  disagree, `.cline/skills` plus `docs/0N-*.md` win.
- `.roo/rules-skill-writer/` — generic skill-authoring rules (naming, progressive disclosure).

Rules for adding or editing a skill in this repository:

- The frontmatter `name` must be lowercase, hyphen-separated, and **identical to its
  directory** (`03-roadmap` in `03-roadmap/`). The generated `.cline` files currently break this
  (directory `00-ponytail`, `name: 00-PONYTAIL`); never copy the uppercase form into a
  hand-written skill.
- Add no third copy of a document: a new skill starts as `docs/0N-name.md` carrying
  `SKILL-SUMMARY` and `SKILL-TOOLS` blocks plus a `$map` entry in
  `scripts/sync-cline-skills.ps1`, then the sync, then the `.roo` mirror.
- `docs/08-TEST-CHECKLIST.md` is intentionally skill-less; read it directly.
