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
2. Open each application once and dismiss first-run prompts.
3. In Excel **File > Account > About Excel**, record version, build, update channel, and 32/64-bit status.
4. Keep Office updates enabled on developer machines. Release testing also needs one controlled build representing production.
### VS Code and Cline
1. Install VS Code, C# Dev Kit, and Cline from their official publishers.
2. Open the repository root.
3. In Cline's rules view, enable `AGENTS.md` and all `.clinerules` files.
4. Confirm project skills appear from `.cline/skills/`.
5. Do not enable blanket auto-approval. Safe read-only commands and targeted test commands may be approved per workspace after review.
### Copilot Agent Mode
1. Use Visual Studio 2022 17.14 or later.
2. Open **Tools > Options > GitHub > Copilot**.
3. Enable Agent Mode, planning, and repository custom instructions.
4. Check that `.github/copilot-instructions.md` appears in response references.
5. Use prompt files under `.github/prompts/` for repeatable tasks.
## Visual Studio orientation for a first-time user
Open `GanttCreator.sln`, then use these windows:
| Window | Open from | Use in this project |
| --- | --- | --- |
| Solution Explorer | **View > Solution Explorer** | Navigate projects, references, resources, and tests |
| Error List | **View > Error List** | Compiler/analyzer errors; keep filters on entire solution |
| Output | **View > Output** | Select Build or Debug to see the first real failure |
| Test Explorer | **Test > Test Explorer** | Run/filter unit and Office integration tests |
| Git Changes | **View > Git Changes** | Review changed/untracked files; do not commit blindly |
| Object Browser | **View > Object Browser** | Confirm real COM types, members, parameters, and enums |
| Modules | while debugging: **Debug > Windows > Modules** | Confirm which XLL/DLL and symbols Excel loaded |
Set the toolbar configuration to `Debug` and platform to `x64`. If x64 is missing, open **Build > Configuration Manager**, create x64 from Any CPU, and make every Office-hosted project use x64. Core tests may remain Any CPU, but a uniform x64 solution configuration is easier for a beginner.
Useful actions:
- `Ctrl+Shift+B`: build the solution without starting Excel.
- `F5`: start Excel under the debugger.
- `Ctrl+F5`: start without debugging; use only for a smoke check.
- `F9`: toggle a breakpoint on a callback/application-service line.
- `F5` while paused: continue.
- `Shift+F5`: stop debugging. Prefer closing Excel normally first.
- Right-click one test or project in Test Explorer to run a narrow test while coding.
Build reads source and produces binaries; it should be routine. Rebuild first cleans outputs and is slower. Do not use Clean/Rebuild as a ritual—use it when evidence points to stale output or locked artifacts.
Use the CLI for package changes so the diff is obvious and repeatable. Review `Directory.Packages.props` and lock-file changes before accepting them. Use Visual Studio's NuGet UI for inspection, not as permission for an agent to upgrade unrelated packages.
### Daily development loop
1. Pull/rebase according to the team's Git policy and create the roadmap branch.
2. Open the solution in Visual Studio and the same repository root in VS Code if using Cline.
3. Run `pwsh ./scripts/verify-quick.ps1` before editing; an initial failure is not caused by the new task.
4. Implement and run narrow tests without Excel where possible.
5. Close any Excel debugging instance before rebuilding the XLL.
6. Use F5 only for the smallest live behaviour that cannot be proved without Office.
7. Run full verification, inspect the entire diff, and obtain human review before committing.
Keep only one debugging Excel instance unless the test explicitly needs more. Give test workbooks an obvious synthetic title so they cannot be mistaken for live work.
## Repository bootstrap
Create projects with the CLI or equivalent Visual Studio dialogs:
```powershell
dotnet new sln -n GanttCreator
dotnet new classlib -n GanttCreator.Core -o src/GanttCreator.Core -f net10.0
dotnet new classlib -n GanttCreator.Raster -o src/GanttCreator.Raster -f net10.0
dotnet new classlib -n GanttCreator.Office -o src/GanttCreator.Office -f net10.0
dotnet new classlib -n GanttCreator.AddIn -o src/GanttCreator.AddIn -f net10.0
dotnet new xunit -n GanttCreator.Core.Tests -o tests/GanttCreator.Core.Tests -f net10.0

---

## Where to read more

- Canonical source: `docs/01-ENVIRONMENT.md`
- Always-on rule: `.clinerules/01-ENVIRONMENT.md`
- Full reference: `./references.md` in this directory
