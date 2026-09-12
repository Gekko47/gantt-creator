# Changelog

All notable changes to the Gantt Creator project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Conventional Commits](https://www.conventionalcommits.org/).

## [Unreleased]

### Added

- **feat**: Add Excel-DNA entry point with AutoOpen/AutoClose logging (R1.1)
- **test**: Serialize temp-fallback log tests to avoid concurrent file collisions
- **build**: Fail verify.ps1 fast on a dirty working tree
- **feat(skills)**: Hook the seven Cline skills to MCP tools
- **feat(test)**: Add OfficeFixture harness for real-Excel integration tests
- **feat**: Fail skill sync when the SKILL-SUMMARY block is missing
- **feat**: Verify cached actionlint executable SHA-256 before every execution
- **feat(verify)**: Use `--locked-mode` when lock files exist (R0.3)
- **feat(build)**: Enable central package lock file (R0.3)
- **feat(R0.8)**: Versioning, rolling log, workflow lint (L6), script tests (L7)
- **build(ci)**: Add Windows CI workflow (R0.7)
- **build(gates)**: Add skill-summary phrases, status-script tripwires, validator+test rule
- **build(gates)**: Close CI↔local parity; adopt PSScriptAnalyzer; defer L6/L7
- **build(gates)**: Reviewer fixes — 8 findings from coderabbit, regen views
- **build**: Solution, projects, central packages, verify scripts (R0.2-R0.6)
- **build**: Add .editorconfig and .gitattributes (R0.4)
- **build**: Commit packages.lock.json files (R0.3)
- **build**: Commit test project lock files (R0.3 follow-up)
- **build(tests)**: Import root Directory.Build.props from test projects
- **build(analyzers)**: Add CA1707 to NoWarn with documented rationale
- **build(docs)**: Enforce XML doc comments in production, not in tests
- **docs**: Install kit, control records, and four ADRs (R0.1)
- **docs**: Add 08-TEST-CHECKLIST and encode preventive test rules
- **docs**: Add L8 to KNOWN-LIMITATIONS.md
- **docs**: Add retroactive R0.6 and R0.2 work-item files
- **docs(evidence)**: Anchor harness-refactor evidence to source commit
- **docs(status)**: Record R0.1 through R0.5 completion entries
- **docs(status)**: Record duration-claim alignment and offline locked-restore round
- **docs(status)**: Record review-evidence hardening round in STATUS.md
- **docs(status)**: Record Windows 11 25H2 host, skill tree, drift gate
- **docs**: Record lint-ci single-file splat fix in STATUS
- **docs**: Record AddInAssemblyTests Debug fallback fix in STATUS.md
- **docs**: Record pre-commit hook in canonical 05-GIT-QUALITY
- **docs**: Update STATUS.md and SYNC.md to reflect current state
- **docs**: Add .clinerules/SYNC.md
- **ci**: Bump GitHub Actions to v5 to clear Node.js 20 deprecation warning

### Changed

- **docs**: Record two-gate sequence and human-gated branch lifecycle
- **docs**: Update STATUS.md with R1.1 AddInLogFactory fix commit
- **build**: Enhance Dependabot config to include NuGet updates
- **refactor**: Single-source actionlint version + SHA-256 in ci.yml from tool-versions.psd1
- **refactor**: Extract check-status path predicate to shared Get-RepoPathTokenStatus helper
- **refactor**: Derive locked-restore harness sequence from the script AST
- **refactor(skills)**: Collapse three-view to two-view skill discipline
- **refactor**: Fix VS Code schema, ExportSize.ParseValue non-finite guard, ToString culture test
- **style(format)**: Apply dotnet format to production code
- **style**: Normalize all tracked text files to LF

### Fixed

- **fix**: Update AddInLogFactory with IsFailed fallback and isolated temp test directories
- **fix**: Update active work item status and enhance verification scripts
- **fix**: Anchor verify.ps1 dirty-tree check to repo root with git -C
- **fix**: Update test fakes and temp path isolation in AddInHostTests and AddInLogFactoryTests
- **fix**: Suppress disposal exceptions in AutoClose and clear _log in nested finally
- **fix**: Update lint-ci anti-drift assertions for delegation pattern
- **fix**: Fix failing verify-office.Tests.ps1 and harden check-cline-skills.Tests.ps1
- **fix**: Require hex letter in LongHexTokenPattern to avoid redacting decimal IDs
- **fix**: Enhance pre-commit tests to reject invalid verify-quick.ps1 invocations
- **fix**: Coerce lint-ci workflow list to array so single-file splat passes the path
- **fix(lint)**: Enhance workflow file validation in lint-ci.ps1
- **fix(tests)**: Enhance ci.yml parity checks and add POSIX hook installation test
- **fix(tests)**: Serialize Office integration test collection
- **fix**: Apply stage-inspect review findings (docs, script gates, artifact resolution)
- **fix(core)**: Align VersionInfo fallback with Directory.Build.props emitted value
- **fix(tests)**: Use Path.GetRelativePath for artifact-marker path reporting
- **fix(core)**: Skip rotation for oversized message on empty active log
- **fix(scripts)**: Flag branch coverage below target in verify.ps1
- **fix(scripts)**: Pass absolute DocsRoot to sync-cline-skills in drift gate
- **fix**: Contain RollingLog format failures to the current write
- **fix**: Close review findings across raster/scripts/tests/docs
- **fix(scripts)**: Harden check-status path containment and locked-restore search root
- **fix**: Run drift-gate Phase 1 git calls against the repo root
- **fix**: Block pre-commit on git failure and test the real checker loop
- **fix**: Latch RollingLog directory-creation failures behind IsFailed
- **fix**: Write RollingLog files as UTF-8 without a BOM preamble
- **fix**: Make RollingLog failure latch visible to lock-free readers
- **fix**: Align review-evidence findings across docs, scripts, and export size
- **fix(tests)**: Restore Debug fallback in AddInAssemblyTests locator
- **fix(ci)**: Route PSSA step through Invoke-PssaGate instead of -EnableExit
- **fix(pre-commit)**: Repair dead-code dirty-tree guard with repo root resolution
- **fix(core,raster)**: Harden RollingLog and ExportSize validation
- **fix(core)**: Pin PointD.ToString to invariant culture
- **fix(vscode)**: Correct verify-full detail to reflect non-Office tests
- **fix(raster)**: Guard ExportSize.ToPixels against non-finite and oversized inputs
- **fix(vscode)**: Invoke pwsh explicitly in tasks; split build-clean clean/build
- **fix(verify)**: Drift gate checks working tree, not post-sync state
- **fix**: Resolve default solution path against the repository root
- **fix**: Correct RollingLog rotation ordering and harden VersionInfo
- **fix(ci)**: Get-Module -MinimumVersion not supported; align with test-scripts.ps1
- **fix(qa)**: Clear CI PSScriptAnalyzer gate failures
- **fix(qa)**: Verify.ps1 parity with verify-quick + pre-commit safety-net hook
- **fix(qa)**: Implement Phase 2 of check-cline-skills drift gate (H-1)
- **fix(qa)**: Hardening pass - Test-Path discipline, PSSA defence-in-depth, CI parity, drift notes
- **fix**: Fix format issues and complete R0.8 implementation
- **fix(scripts)**: Strip BOM and em-dashes to clear PSScriptAnalyzer gate
- **fix(docs)**: Correct comments regarding the number of checkers in pre-commit scripts
- **fix**: Correct sceneHeightPt contract in ExportSize XML docs
- **fix**: Reopen L10 until green CI run observed
- **fix**: Build-clean sequence, AddIn publish step, CA1707 test-only
- **docs**: Align verification cadence, PSSA pin, Pester 6+, and step-number references
- **docs**: Align VS 2026 baseline, L11 owner, Pester wording, and R0.8 step count
- **docs**: Align entity-guide, test checklist, and R0.7 work item
- **docs**: Align verify-gate duration claims with measured runs; add drift tripwire
- **docs**: Close L3/L8/L9, align Pester 6+ standard across all docs and CI
- **docs**: Record green CI run #40, close R0.8/L10/L11
- **docs**: Reference MaxPixelDimension via cref in ExportSize exception doc
- **docs**: Fix drift and add status accuracy gate
- **docs**: Record docstring-coverage decision (L8)
- **docs**: Correct R0.5 test count and entry-point description in work item
- **docs**: Record R0.3, R0.4, R0.5 completion in STATUS
- **docs**: Record actual host as Windows 11 25H2 (build 26200); remove L1
- **docs**: Update Dependabot configuration for GitHub Actions
- **test**: Add RollingLog constructor guard tests
- **test**: Tighten script test assertions and add numbering guard
- **test**: Remove redundant ToPixels_oversized_pixel_width_throws
- **test(scripts)**: Prove locked restore succeeds without network access (R0.3)
- **test(architecture)**: Enforce IDE0011 alongside CA1707 in test-only NoWarn scope
- **test**: Automate 08-TEST-CHECKLIST enforcement via architecture tests
- **test(R0.5)**: Fix sample tests against production analyzers
- **test(R0.5)**: Replace placeholders with sample tests, add VS Code tasks
- **test**: Strengthen analyzer rules and update tests for compliance
- **lint**: Enforce whitespace and strengthen test validations across 10 files

### Chore

- **chore**: Bump Pester minimum to 6, tighten script guards
- **chore(skills)**: Populate .cline/skills/, regenerate .clinerules/ from docs/, wire drift gate

---

## [R0.8] — Versioning, Rolling Log, Workflow Lint, Script Tests

### Added

- Versioning system with `VersionInfo` class
- Rolling log infrastructure (`RollingLog`, `IRollingLog`, `IRedactor`)
- Workflow lint gate (L6) in CI
- Script test suite (L7) with `test-scripts.ps1`
- `scripts/verify.ps1`, `scripts/verify-quick.ps1`, `scripts/verify-office.ps1`
- `scripts/pre-commit.ps1` with full checker pipeline
- `scripts/sync-cline-skills.ps1` for skill synchronization
- `scripts/check-cline-skills.ps1` for skill drift detection
- `scripts/check-status.ps1` for status document accuracy
- `scripts/check-md-links.ps1` for markdown link validation
- `scripts/build-release.ps1` for release packaging
- `scripts/publish-addin.ps1` for AddIn publishing
- `scripts/pssa-gate.ps1` for PSScriptAnalyzer gating
- `scripts/step-parity.ps1` for step consistency
- `scripts/tool-versions.psd1` for pinned tool versions

### Changed

- Adopt PSScriptAnalyzer in CI pipeline
- Close CI↔local parity gap
- Defer L6/L7 to subsequent iteration

---

## [R0.7] — Windows CI Workflow

### Added

- GitHub Actions CI workflow (`.github/workflows/ci.yml`)
- Dependabot configuration (`.github/dependabot.yml`)
- Copilot instructions (`.github/copilot-instructions.md`)
- CI parity checks in `scripts/ci-parity.Tests.ps1`

---

## [R0.6] — Verification Scripts and Coverage

### Added

- `scripts/verify-helpers.ps1`, `scripts/verify-office.Tests.ps1`, `scripts/verify-quick.ps1`
- `scripts/test-locked-restore.ps1`, `scripts/test-non-office.ps1`, `scripts/test-scripts.ps1`
- `scripts/actionlint-hash.ps1`, `scripts/PSScriptAnalyzerSettings.psd1`
- Lock-file-based restore verification
- Offline locked-restore demonstration

---

## [R0.5] — Sample Tests and Tasks

### Added

- `tests/GanttCreator.Core.Tests/PointDTests.cs`, `RedactorTests.cs`, `RollingLogTests.cs`, `VersionInfoTests.cs`
- `tests/GanttCreator.Raster.Tests/ExportSizeTests.cs`
- `tests/GanttCreator.AddIn.Tests/AddInAssemblyTests.cs`, `ArtifactSourceMarkerTests.cs`
- `tests/GanttCreator.Office.ContractTests/`, `GanttCreator.Office.IntegrationTests/`
- `tests/GanttCreator.Architecture.Tests/`
- VS Code tasks configuration (`.vscode/tasks.json`)
- Sample tests replacing placeholders
- Test project lock files

### Changed

- PointD.ToString pinned to invariant culture
- ExportSize guarded against non-finite and oversized inputs
- XML doc comments enforced in production code

---

## [R0.4] — Editorconfig and Gitattributes

### Added

- `.editorconfig` with consistent formatting rules
- `.gitattributes` with LF normalization
- CA1707 analyzer rule to NoWarn with documented rationale
- Test project Directory.Build.props import
- `docs/08-TEST-CHECKLIST.md`

---

## [R0.3] — Central Packages and Lockfile

### Added

- Central package management via `Directory.Packages.props`
- Package lock files (`packages.lock.json`) across all projects
- `--locked-mode` flag for `verify` command
- `global.json` for SDK version pinning
- Central package restore verification without network

---

## [R0.2] — Solution and Project Structure

### Added

- `GanttCreator.slnx` solution file
- `src/GanttCreator.Core/`, `AddIn/`, `Office/`, `Raster/` projects
- `tests/` directory with all test projects
- `scripts/` directory with verification scripts
- `docs/` directory structure

---

## [R0.1] — Governance and Scaffolding

### Added

- `AGENTS.md`, `CONTRIBUTING.md`, `SECURITY.md`, `LICENSE`
- `docs/01-ENVIRONMENT.md` through `docs/08-TEST-CHECKLIST.md`
- `docs/DECISIONS.md`, `docs/KNOWN-LIMITATIONS.md`, `docs/STATUS.md`, `docs/REVIEW-CHECKLIST.md`
- Five ADRs in `docs/adr/`
- `.githooks/pre-commit` hook
- `scripts/` infrastructure scripts
- Work-item template (`docs/work-items/TEMPLATE.md`)

---

## [Initial]

### Added

- Initial repository commit with project scaffolding
- `.gitignore` configuration
- Dependabot configuration
- CI workflow initial version
