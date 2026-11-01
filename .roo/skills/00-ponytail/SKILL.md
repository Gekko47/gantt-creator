---
name: 00-ponytail
description: Lazy-senior-dev minimal-diff discipline for the Gantt Creator repository — the YAGNI ladder (reuse before rewrite, stdlib/native/installed dependency before new code, root-cause over symptom fix, no unrequested abstraction, one runnable check for non-trivial logic). Part of the gantt-creator index and the FIRST skill to apply: it runs before any other skill and drives every task. Trigger on session start, before writing any code, or when deciding how small a change should be.
---

# Ponytail — lazy senior developer discipline (driving skill)

Mirror of [`.cline/skills/00-ponytail/SKILL.md`](../../../.cline/skills/00-ponytail/SKILL.md),
which is generated from [`docs/00-PONYTAIL.md`](../../../docs/00-PONYTAIL.md). If the two
disagree, the canonical doc wins.

## Index

- Call the index skill `gantt-creator-contract` first; it routes.
- **This skill outranks the others.** It runs before `01`–`07` and applies to every task.
- Defer to `06-llm-protocol` for evidence, retries, and handoff; to `04-test-strategy` for
  the check the change must leave behind; to `02-architecture` / `07-entity-guide` before
  touching a boundary or a visual contract.

## The ladder (climb after understanding, never instead of it)

1. Does this need to be built at all? (YAGNI)
2. Does it already exist in this codebase? Reuse the helper, util, or pattern.
3. Does the standard library already do this?
4. Does a native platform feature cover it?
5. Does an already-installed dependency solve it?
6. Can this be one line?
7. Only then: the minimum code that works.

## Do not get wrong

- The ladder never replaces understanding. The smallest change in the wrong place is a second
  bug — read the code the task touches and trace the real flow end to end first.
- A bug fix is a root-cause fix: grep every caller of the function you touch and fix the
  shared function once.
- Lazy code without its check is unfinished: non-trivial logic leaves ONE runnable check that
  follows the repo's testing rules and runs in the mandated harness.
- Not lazy about: input validation at trust boundaries, error handling that prevents data
  loss, security, accessibility, hardware calibration, or anything explicitly requested.
- No unrequested abstraction, no avoidable new dependency, deletion over addition, boring over
  clever. Mark a deliberate corner cut with a `ponytail:` comment naming the ceiling.
- Conflicts between instructions are reported, not silently resolved. All behaviour changes
  need a test.

## Tools here

`codebase_search` + `search_files` (find the existing helper; grep callers) · `read_file`
(trace the flow) · `execute_command` (`dotnet build`, `dotnet test`,
`pwsh -NoProfile -File scripts/verify-quick.ps1`).
