---
name: 05-git-quality
description: Branch and review policy, commit design and Conventional Commit format, local gates (verify-quick.ps1 / verify.ps1), CI parity, PR template, review checklist, comment and dependency policy, and the release gate for the Gantt Creator repository. Part of the gantt-creator index. Trigger on commit, PR, dependency, release, or branch-policy questions.
---

# Git quality and release

Mirror of [`.cline/skills/05-git-quality/SKILL.md`](../../../.cline/skills/05-git-quality/SKILL.md),
generated from [`docs/05-GIT-QUALITY.md`](../../../docs/05-GIT-QUALITY.md). The canonical doc
wins on any disagreement.

## Index

- Index first: `gantt-creator-contract`; `00-ponytail` before any code change.
- Pair with `03-roadmap` (one row, one commit) and `04-test-strategy` (the evidence a PR must
  carry).

## Do not get wrong

- A phase may not exit on local-only evidence: the branch must be pushed and GitHub CI
  observed green (the W-12 rule — added after the PSSA blind-gate defect survived three
  local-only "all green" QA rounds).
- Conventional Commit prefixes only (`feat:`, `fix:`, `test:`, `refactor:`, `docs:`, `build:`,
  `chore:`); one concern per commit; commit only a green state.
- Never rewrite shared history (amend/rebase/force-push), tag, or publish without explicit
  instruction.
- Comments only when they add current, non-obvious value; remove stale ones in the touched
  area. Work items and `docs/STATUS.md` are control records, not diaries.
- Do not commit Office temporary files, build output, test results, exported customer images,
  or local logs. Install the pre-commit hook
  (`pwsh ./scripts/install-pre-commit.ps1`); it is a fast net, not a substitute for the gates.
- Every commit message lists the `docs/08-TEST-CHECKLIST.md` sections actually certified
  (e.g. `Checklist: A, B, C, G, I`); never claim a section that does not apply.

## Gates and tools here

`execute_command`: `git status --short`, `git diff --check`, `git diff`,
`git log --oneline`, `pwsh -NoProfile -File scripts/verify-quick.ps1` (staged),
`pwsh -NoProfile -File scripts/verify.ps1` (clean tree, after the last commit for the branch).
CI is [`.github/workflows/ci.yml`](../../../.github/workflows/ci.yml) and mirrors
`verify-quick.ps1` step for step; `scripts/tool-versions.psd1` is the single source of tool
pins.
