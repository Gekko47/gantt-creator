---
name: 05-GIT-QUALITY
description: Branch/review policy, commit design and Conventional Commit format, local gates (verify-quick.ps1 / verify.ps1), CI parity, PR template, review checklist, comment and dependency policy, and the release gate. Trigger on commit, PR, dependency, or release questions.
---

Branch/review policy, commit design, local gates, CI parity, and the
pull-request/review checklist.

Do not get wrong:
- A phase may not exit on local-only evidence: the branch must be
  pushed and GitHub CI must have been observed green (the W-12 rule —
  added after the PSSA blind-gate defect survived three local-only
  "all green" QA rounds).
- Conventional Commit prefixes only (`feat:`, `fix:`, `test:`,
  `refactor:`, `docs:`, `build:`, `chore:`); one concern per commit.
- Never rewrite shared history (amend/rebase/force-push) without
  explicit instruction.
- Comments only when they add current, non-obvious value; remove stale
  ones in the touched area. Work items and STATUS are control records,
  not diaries.

---

## Where to read more

- Full reference: `docs/05-GIT-QUALITY.md` (this is the canonical source; there is no separate copy under .cline/skills/)
