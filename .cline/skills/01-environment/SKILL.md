---
name: 01-ENVIRONMENT
description: Use this skill when the conversation is about the development baseline, the .NET / Office / Visual Studio / Excel-DNA / VS Code versions on this machine, the first-run installation checklist, common Office-host failures, or the offline acceptance check.
---

# Environment and first run
## Supported development baseline
Record actual versions in the first work item and CI output. Do not silently develop against a different baseline.
| Component | Baseline | Why |
| --- | --- | --- |
| OS | Windows 11 x64 | Office COM, clipboard, RibbonX, and Excel-DNA are Windows-hosted. The actual host is Windows 11 25H2 (build 26200); see `docs/adr/0005-windows-host-build-number.md`. The `HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion` `ProductName` string may report "Windows 10 Pro" on some configurations; trust the build number, not the string. |
| Excel/PowerPoint | Microsoft 365 desktop, x64 | Primary deployment target and `ppPasteShape` host |
| Visual Studio | 2022 17.14+ | Excel launch/debug and Copilot Agent Mode (the actual IDE in use is Visual Studio 2026 Community; see `docs/adr/0003-visual-studio-2026.md`) |
| .NET | SDK 10 LTS and Desktop Runtime 10 x64 | Current Excel-DNA 1.9 guidance |
| Excel-DNA | 1.9.0 | Produces the `.xll` and hosts RibbonX/C# |
| VS Code | Current stable | Cline workspace and cross-platform editing |
| PowerShell | 7.x | Repository scripts |
| Git | Current Git for Windows | Source control and hooks |
Pin application dependencies in `Directory.Packages.props`. Pin the SDK with `global.json` after confirming the exact installed `.NET 10` feature band. Use `rollForward: latestFeature`, not a floating major version.
## Installation checklist
### Visual Studio
In Visual Studio Installer:
1. Install or modify Visual Studio 2022.
2. Select **.NET desktop development**.
3. In individual components, confirm the `.NET 10 SDK`, Git for Windows, and NuGet package manager.
4. Install GitHub Copilot if using Visual Studio's built-in agent.
5. Start Visual Studio, sign in, then check **Help > About**.
The `Office/SharePoint development` workload is not required by Excel-DNA itself. Install it only if a later approved task needs Visual Studio's Office tooling. The deployed add-in is Excel-DNA, not VSTO.
### Office
1. Install desktop Excel and PowerPoint with the same architecture.

---

## Where to read more

- Canonical source: `docs/01-ENVIRONMENT.md`
- Always-on rule: `.clinerules/01-ENVIRONMENT.md`
- Full reference: `./references.md` in this directory
