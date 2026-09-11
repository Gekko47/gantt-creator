---
name: 03-ROADMAP
description: Work-item list R0.x-R10.8 with per-item automated and Office gates, the cross-phase compatibility matrix, and the scope-change protocol. Trigger on which-work-item, which-phase, which-gate, or next-safe-action questions.
---

Phase-by-phase work-item list (R0.x .. R10.8), the cross-phase
compatibility matrix, and the scope-change protocol.

Do not get wrong:
- One roadmap row is normally one commit; split a row if the diff
  gets hard to review, never combine unrelated rows.
- Every acceptance criterion naming a test count, command, or
  artifact must link to a specific test or script step — drift
  between this doc and the code is itself a defect
  (docs/08-TEST-CHECKLIST.md section I).
- A phase exits only on its stated automated demonstration. A
  screenshot supports evidence; it never substitutes for it.

---

## Tools

- `github__list_commits` / `git_tool` — verify "one row is one commit" and Conventional Commit prefixes.
- `github__get_pull_request_status` — verify green CI before declaring a phase exit (the W-12 rule).
- `github__get_pull_request_files` — check the diff links to a specific test or script step.
- `dotnet_test` — run the named test or script step an acceptance criterion references.
- `read_files` on `docs/work-items/` — confirm a work-item file exists before implementation.
- `memra_add` — record green CI evidence (commit hash + run URL) as a fact.

---

## Where to read more

- Full reference: `docs/03-ROADMAP.md` (this is the canonical source; there is no separate copy under .cline/skills/)
