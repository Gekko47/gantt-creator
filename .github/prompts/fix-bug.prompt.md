---
mode: agent
description: Investigate and fix a reported bug
---

# Fix a bug

You are fixing bug **${input:ID:short description or issue link}**.

## Before you edit

1. If the bug is reproducible without Office, write the regression test first and confirm it fails.
2. If the bug is Office-only, write the closest contract test and a written reproduction record (Office build, command, observed result). Do not invent a "fix" without the repro.
3. Capture the exact error, HRESULT, and stack trace. Preserve the first useful exception.
4. Form one falsifiable hypothesis and one discriminating check.

## Two-attempt rule

- **Attempt 1**: preserve the error, run the discriminating check, make the smallest corresponding fix, rerun the narrow test.
- **Attempt 2**: re-read the boundary and primary documentation, change the hypothesis materially, try one different fix.
- **Third occurrence**: stop. Report attempts, evidence, current diff, the likely layer, and the smallest question.

## Do not

- repeat an unchanged command expecting a different result,
- add `Thread.Sleep` for sync,
- broaden `catch` to silence the error,
- delete or weaken a test to make it green,
- suppress a warning without a written reason,
- introduce a new dependency without an ADR.

## Before you hand back

1. `pwsh ./scripts/verify-quick.ps1` — exit 0.
2. `pwsh ./scripts/verify.ps1` — exit 0 if the fix is Core or a renderer.
3. The regression test must pass and the test name must reference the bug ID.
4. Update `docs/STATUS.md` and add a known-limitation entry if the fix is partial.
5. Report: root cause, the discriminating test, the change, the evidence, and any residual risk.
