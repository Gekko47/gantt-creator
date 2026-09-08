---
name: 01-ENVIRONMENT
description: Dev-machine baseline (.NET / Office / Visual Studio / Excel-DNA / VS Code versions), first-run installation checklist, common Office-host failures, and the offline acceptance check. Trigger on environment, setup, install, or version questions.
---

Dev-machine baseline, first-run install checklist, and the offline
acceptance check for this repo.

Do not get wrong:
- Record actual installed versions in the work item and CI output;
  never silently develop against a different baseline than the table
  below states.
- Trust the Windows build number, not the `ProductName` registry
  string, which can misreport "Windows 10 Pro" on a Windows 11 host.
- The IDE in use is Visual Studio 2026 Community (see ADR 0003), not
  a generic "VS 2022 17.14+" — check the ADR before assuming version
  parity with the baseline table.

---

## Where to read more

- Full reference: `docs/01-ENVIRONMENT.md` (this is the canonical source; there is no separate copy under .cline/skills/)
