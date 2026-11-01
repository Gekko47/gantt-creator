---
name: 03-roadmap
description: Work-item and phase map for the Gantt Creator repository (R0.x through R10.x, including letter-suffixed rows), with per-item automated and Office gates, the cross-phase compatibility matrix, and the scope-change protocol. Part of the gantt-creator index. Trigger on which-work-item, which-phase, which-gate, next-safe-action, or scope-change questions.
---

# Roadmap and work items

Mirror of [`.cline/skills/03-roadmap/SKILL.md`](../../../.cline/skills/03-roadmap/SKILL.md),
generated from [`docs/03-ROADMAP.md`](../../../docs/03-ROADMAP.md). The canonical doc wins on
any disagreement.

## Index

- Index first: `gantt-creator-contract`; `00-ponytail` before any code change.
- The current approved work item in `docs/work-items/` is source-of-truth item 1, above the
  roadmap. Pair with `05-git-quality` for commit/CI rules and `04-test-strategy` for the
  test evidence a row must carry.

## Do not get wrong

- One roadmap row is normally one commit; split a row when the diff becomes hard to review,
  never combine unrelated rows.
- Every acceptance criterion naming a test count, command, or artifact must link to a specific
  test or script step — drift between doc and code is itself a defect
  (`docs/08-TEST-CHECKLIST.md` section I).
- A phase exits only on its stated automated demonstration. A screenshot supports evidence; it
  never substitutes for it.
- A work-item file must exist before implementation. If it does not, stop and ask.
- Letter-suffixed rows (e.g. `R5.6a`–`R5.6e`) normalise to their numeric row for `check-status`,
  so a missing row will not be caught for you — verify the ID exists in the roadmap.

## Tools here

`read_file` (`docs/03-ROADMAP.md`, `docs/work-items/`, `docs/STATUS.md`) ·
`codebase_search` / `search_files` (trace an acceptance criterion to a real test) ·
`execute_command` (`git log`, `git status`, `dotnet test` on the named project, and
`pwsh -NoProfile -File scripts/check-status.ps1`).
