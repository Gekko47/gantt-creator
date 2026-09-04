# Work item — <ID>

> Copy this file to `docs/work-items/<ID>-<short-slug>.md` before implementing.
> One ID per file. One concern per commit. Update `docs/STATUS.md` after each commit.

## Outcome

One sentence. The observable behaviour that will be true after this work item is done.

## Scope

In-scope files, projects, and modules. List the directory paths the agent is allowed to touch.

## Exclusions

Explicitly out of scope. Anything not in this list is *not* part of this work item.

## Acceptance tests

- **Automated**: the exact `dotnet test` filter, fixture path, or verify script step that proves the behaviour.
- **Office host** (if `Required` in `docs/03-ROADMAP.md`): the F5 / live-host evidence required.
- **Visual / golden** (if a renderer or style changes): the baseline ID and tolerance.

## Evidence commands

The exact commands the agent must run, in order, with the expected exit status.

```powershell
pwsh ./scripts/verify-quick.ps1    # every commit — expect exit 0
pwsh ./scripts/verify.ps1          # before PR — expect exit 0
# Office integration:
pwsh ./scripts/verify-office.ps1   # on phase exit / release — expect exit 0
```

## Risk and rollback

The main risk, and the simplest rollback path.

## Definition of done

From `AGENTS.md`: acceptance criteria have passing automated tests; format, analyzers, Release build, unit tests, and coverage thresholds pass; no warning was suppressed without a written reason; no test was weakened or deleted merely to obtain green output; public behaviour and errors are documented where a user or maintainer needs them; new Office calls are behind an adapter and have a contract or integration test; Excel state is restored in `finally` blocks; COM objects have explicit, reviewable lifetimes; the workbook still uses one user worksheet and has no hidden/helper worksheet dependency; the active task, decision log, and known limitations are current; the diff contains no unrelated refactor; a human reviewed the result in Excel for visual slices.

## Notes during implementation

Short factual notes only. Add links to evidence and ADRs, not prose. Remove before commit if stale.
