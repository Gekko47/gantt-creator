---
name: 05-GIT-QUALITY
description: Use this skill when the conversation is about branch and review policy, commit design and Conventional Commit format, the local gates (verify-quick.ps1 / verify.ps1), CI jobs, the pull-request template, the review checklist, the comment policy, the dependency policy, or the release gate.
---

# Git, quality gates, and review
## Branch and review policy
- Protect `main`: pull requests, passing required checks, and one human approval.
- Name branches `type/roadmap-id-short-description`, for example `feat/r3-09-stacked-events`.
- Rebase/merge policy is a team choice; do not let an agent rewrite shared history.
- No direct production release from an unreviewed local working tree.
## Commit design
Each commit:
- implements one roadmap outcome or one isolated preparatory refactor;
- keeps the solution buildable and tests green;
- includes tests with changed behaviour;
- avoids unrelated formatting or rename noise;
- updates only the necessary control documents;
- is understandable without reading later commits.
Suggested soft limit: 50-250 changed production lines plus tests. Split by behaviour, not by arbitrary file count. A pure rename, generated interop file, or approved golden-image commit can exceed the limit but must be isolated.
Commit format:
```text
feat(scene): render stacked events on one lane
Map StackIndex to deterministic sub-lane offsets while preserving lane height.
Tests cover shuffled input, duplicate stack values, and clipping.
```
Do not use messages such as `updates`, `fix stuff`, or an agent transcript.
## Local gates
During editing:
```powershell

---

## Where to read more

- Canonical source: `docs/05-GIT-QUALITY.md`
- Always-on rule: `.clinerules/05-GIT-QUALITY.md`
- Full reference: `./references.md` in this directory
