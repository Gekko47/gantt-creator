# Known limitations

> Short list of accepted, time-bounded limitations. Each entry has an ID, an owner, an expiry, and a replacement release gate. Quarantined tests do not count as passing; quarantined features do not count as shipped.

| ID | Limitation | Owner | Since | Expires | Replacement gate |
| --- | --- | --- | --- | --- | --- |
| L1 | _(removed by ADR-0005: the actual host is Windows 11 25H2 / build 26200)_ | — | — | — | — |
| L2 | Visual Studio 2026 (not 2022) is the verification IDE | _unassigned_ | 2026-09-04 | first R1.1 Office gate | R1.1 F5 evidence with a documented VS 2026 build number |
| L3 | No Office-integration test fixture committed yet (R0.x scope) | _unassigned_ | 2026-09-04 | R0.5 exit | R0.5 sample test for the Office contract layer |
| L4 | Golden PNG baseline directory does not exist yet | _unassigned_ | 2026-09-04 | R3.12 exit | R3.12 reference scene with at least one approved PNG |
| L5 | The `_GanttCreatorConfig` safe-repair policy is not yet specified | _unassigned_ | 2026-09-04 | R2.10 start | R2.10 ADR defining auto-repair vs confirmation boundary |
