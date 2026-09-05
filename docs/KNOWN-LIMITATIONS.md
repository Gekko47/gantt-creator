# Known limitations

> Short list of accepted, time-bounded limitations. Each entry has an ID, an owner, an expiry, and a replacement release gate. Quarantined tests do not count as passing; quarantined features do not count as shipped.

| ID | Limitation | Owner | Since | Expires | Replacement gate |
| --- | --- | --- | --- | --- | --- |
| L1 | _(removed by ADR-0005: the actual host is Windows 11 25H2 / build 26200)_ | — | — | — | — |
| L2 | Visual Studio 2026 (not 2022) is the verification IDE | _unassigned_ | 2026-09-04 | first R1.1 Office gate | R1.1 F5 evidence with a documented VS 2026 build number |
| L3 | No Office-integration test fixture committed yet (R0.x scope) | _unassigned_ | 2026-09-04 | R0.5 exit | R0.5 sample test for the Office contract layer |
| L4 | Golden PNG baseline directory does not exist yet | _unassigned_ | 2026-09-04 | R3.12 exit | R3.12 reference scene with at least one approved PNG |
| L5 | The `_GanttCreatorConfig` safe-repair policy is not yet specified | _unassigned_ | 2026-09-04 | R2.10 start | R2.10 ADR defining auto-repair vs confirmation boundary |
| L6 | `ci.yml` has no dedicated workflow lint (actionlint / yamllint); workflow structure is only exercised when CI itself runs | _unassigned_ | 2026-09-04 | R0.8 | A committed workflow-lint step (e.g. `actionlint` pinned binary or equivalent) passing on `ci.yml` in CI |
| L7 | PowerShell scripts in `scripts/` have no unit tests (Pester); they are exercised only by execution inside the verify gates | _unassigned_ | 2026-09-04 | R0.8 | Pester (or equivalent) unit tests for `scripts/*.ps1` wired into `verify-quick.ps1` |
| L8 | External diff-scoped docstring coverage reports 23.08% against an 80% threshold; this is **not** a project gate | _unassigned_ | 2026-09-04 | n/a (decision, not a gap) | The repo's own policy — `GenerateDocumentationFile=true` + `CS1591` build error per undocumented public member in `src/` — is the enforced standard. Decision (2026-09-04, human): do not force-doc non-public or obvious symbols to chase a volume metric; rely on CS1591 + reviewer judgement per `AGENTS.md`. Revisit only if a non-obvious public symbol lands without a summary, or a reviewer's "what does this do?" signals real missing docs. |
