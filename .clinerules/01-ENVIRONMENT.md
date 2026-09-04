# Environment and first run

## Supported development baseline

Record actual versions in the first work item and CI output. Do not silently develop against a different baseline.

| Component | Baseline | Why |
| --- | --- | --- |
| OS | Windows 11 x64 | Office COM, clipboard, RibbonX, and Excel-DNA are Windows-hosted |
| Excel/PowerPoint | Microsoft 365 desktop, x64 | Primary deployment target and `ppPasteShape` host |
| Visual Studio | 2022 17.14+ | Excel launch/debug and Copilot Agent Mode |
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
dotnet new xunit -n GanttCreator.Raster.Tests -o tests/GanttCreator.Raster.Tests -f net10.0
dotnet new xunit -n GanttCreator.Office.ContractTests -o tests/GanttCreator.Office.ContractTests -f net10.0
dotnet new xunit -n GanttCreator.Office.IntegrationTests -o tests/GanttCreator.Office.IntegrationTests -f net10.0
dotnet new xunit -n GanttCreator.AddIn.Tests -o tests/GanttCreator.AddIn.Tests -f net10.0
```

The templates accept the base TFM reliably. After creation, change Raster, Office, AddIn, and their test projects to `net10.0-windows`; leave Core and Core.Tests on `net10.0`. Add all projects to the solution and only the references allowed by `docs/02-ARCHITECTURE.md`. The plain Core target prevents an accidental Windows dependency from compiling there.

Then:

1. Copy root `.editorconfig` and `Directory.Build.props` into place.
2. Add central package management and package lock files.
3. Add `ExcelDna.AddIn` 1.9.0 only to AddIn.
4. Build before writing application code.
5. Commit bootstrap and dependencies separately.

## Excel-DNA first run

Excel-DNA's official current flow is a C# class library targeting `net10.0-windows` with the `ExcelDna.AddIn` 1.9.0 package. A successful build generates architecture-specific XLL output.

Start with an `IExcelAddIn` implementation and one safe diagnostic command. Add RibbonX only after the XLL loads. This separates host-loading failures from Ribbon XML/callback failures.

Debug sequence:

1. Close all Excel processes you own. Check Task Manager before assuming a build issue.
2. Build the AddIn project in `Debug|x64`.
3. Press F5. If the Excel-DNA package did not configure startup, set the project debug external program to the installed `EXCEL.EXE` and pass the generated `.xll` as the command argument.
4. Accept the session-only security notice.
5. Hit a breakpoint in `AutoOpen` or the diagnostic callback.
6. Confirm Excel's bitness matches the loaded XLL.
7. Close Excel from the UI after the test so the debugger detaches cleanly.

Never solve an XLL load failure by copying random DLLs into the Office directory or disabling Trust Center protections globally.

## Local data and fixtures

- Store synthetic workbooks under `tests/fixtures/`; never use customer data.
- Put the legacy `.xlsm` and screenshots in a locally ignored `reference/` directory if licensing permits. They are not release inputs.
- Keep golden PNGs small and focused. Store the input scene beside each approved output.
- Keep exports, logs, `TestResults`, and Office recovery files ignored.

## Common failures

| Symptom | Evidence to collect | First checks |
| --- | --- | --- |
| Excel opens but add-in does not load | Excel-DNA log, Output window, bitness | XLL architecture, runtime installed, Trust Center prompt |
| Build says XLL/DLL is in use | locking process and PID | close Excel; check orphaned `EXCEL.EXE` from integration tests |
| Ribbon tab missing | Ribbon XML validation and callback logs | resource name, namespace, unique IDs, callback signature |
| Breakpoint not hit | loaded module path and symbols | Debug build, correct XLL, correct Excel process |
| `COMException` on a valid member | HRESULT, Office build, STA/thread | main thread, released proxy, workbook/slide still open |
| PowerPoint paste fails | clipboard formats and requested data type | source group copied, PowerPoint visible/ready, `ppPasteShape` supported |
| PNG text differs by machine | font name/version and scale | pinned installed font, invariant culture, renderer/native asset version |
| Test hangs | process list and last observable step | no modal Office dialog, no arbitrary sleep, cleanup deadline |

Apply the bounded retry rule in `AGENTS.md`. Preserve the first useful exception and HRESULT; later wrapper messages often discard the cause.

## Security and offline check

Before the first release candidate:

1. Disconnect the network.
2. Start Excel, load the add-in, create and refresh a chart.
3. Copy editable, send to an already installed PowerPoint, and export PNG.
4. Confirm no delayed package restore, web font, telemetry, online help, update check, or licence request occurs.
5. Inspect logs for workbook content and secrets. Logging must contain identifiers and counts, not activity descriptions or dates by default.

The development machine may need the internet for initial tool/package installation. The built product must not.
