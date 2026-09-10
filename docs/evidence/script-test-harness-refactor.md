# Evidence — script test harness refactor (in-development)

> **Status: in development.** This codebase is not published and has
> no released API surface. The Gantt Creator add-in is a personal
> Windows desktop project; the scripts under `scripts/` and their
> Pester tests are scaffold-grade, hand-rolled tooling that is
> iterated on alongside the production code. Every refactor like the
> one recorded here is expected to be revisited as the project moves
> through the R0.x → R10.x roadmap. No claim of stability, completion,
> or downstream contract is implied by the presence of this file.

## What changed

The four script-test files

- `scripts/check-cline-skills.Tests.ps1`
- `scripts/check-skill-summary.Tests.ps1`
- `scripts/check-status.Tests.ps1`
- `scripts/sync-cline-skills.Tests.ps1`

previously asserted on the **literal text** of the corresponding
`scripts/*.ps1` (e.g. `Should -Match 'sync-cline-skills\.ps1'`,
`Should -Match 'SkillsRoot'`, `Should -Not -Match 'RulesRoot'`). That
style is fragile: a renamed parameter, a different `git` flag, or a
rewording of an error message silently broke the test for reasons
that were unrelated to behaviour.

The replacement pattern — used in `pre-commit.Tests.ps1` and
`test-locked-restore.Tests.ps1` already — is **isolated child-pwsh
execution**: each `It` builds a temp-tree harness, invokes the
script via `Start-Process pwsh -NoProfile -File <script> ...` with
`-WorkingDirectory $script:tempRoot`, captures exit code + combined
stdout/stderr, and asserts on observable output.

## Harness design

### Shared shape

Every updated `Context` follows the same skeleton:

```powershell
BeforeEach {
    $script:tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid())
    New-Item -ItemType Directory -Path $script:tempRoot -Force | Out-Null
    # create sub-directories (docs/, .cline/skills/, scripts/, etc.)
    # initialise fixture content (canonical sources, skill files, roadmap)
    # optionally: git init the temp root if the script under test uses git
    # optionally: copy the script under test into a harness/ subdir
    #             so its $PSScriptRoot-derived $repoRoot aligns with the fixture
}

AfterEach {
    if ($script:tempRoot -and (Test-Path -LiteralPath $script:tempRoot)) {
        Remove-Item -LiteralPath $script:tempRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

It '...' {
    $proc = Start-Process -FilePath pwsh -ArgumentList @(
        '-NoProfile','-File',(Join-Path $script:harness '<script>.ps1'),
        # ...any script params with absolute paths...
    ) -NoNewWindow -Wait -PassThru `
        -WorkingDirectory $script:tempRoot `
        -RedirectStandardOutput $outFile -RedirectStandardError $errFile
    $combined = (Get-Content -LiteralPath $outFile -Raw) + (Get-Content -LiteralPath $errFile -Raw)
    $proc.ExitCode | Should -Be <expected>
    $combined     | Should -Match '<observable string>'
}
```

### Per-script choices

| Script | Harness needs a git repo? | Why `-WorkingDirectory` matters | Script copied into the harness? |
| --- | --- | --- | --- |
| `check-cline-skills.ps1` | Yes — `git diff --name-only` and `git status --porcelain` both fail (exit 128) outside a repo. | Yes — Phase 1's `git diff` runs without `-C`, so it inherits the script's CWD. Without `-WorkingDirectory $tempRoot` the real repo's dirty `docs/STATUS.md` would be reported. | Yes — the script re-invokes `sync-cline-skills.ps1` via a hard-coded path, so the harness needs a copy of both. The harness copy of `sync` has a single-entry `$map` injected by regex. |
| `check-skill-summary.ps1` | No. | Less critical — the script does `Join-Path $repoRoot $SkillsRoot` with an absolute `$repoRoot`. | Yes — `$repoRoot` is `Split-Path -Parent $PSScriptRoot`, and the script is placed one level deeper than `.cline/skills/` so the param resolves correctly. |
| `check-status.ps1` | Yes — uses `git -C $repoRoot rev-parse --verify` for the hash check. | Recommended — keeps the script's CWD inside the harness, away from the real repo. | Yes — same `$repoRoot` reasoning. |
| `sync-cline-skills.ps1` | No. | The script accepts `-DocsRoot` and `-SkillsRoot`; the test passes absolute paths so CWD is irrelevant. | Yes (modified copy) — the test mutates the script's `$map` via regex and writes a copy of the script into the temp root; only the generated copy is used. |

### Phase 2 byte-compare subtlety (`check-cline-skills.ps1`)

Phase 2 of the drift gate runs

```powershell
git -c core.autocrlf=false diff --no-index --quiet -- "$tmpSkills" "$committed"
```

The fixture prime therefore commits with the same flag so the
committed bytes are byte-equal to what the sync wrote:

```powershell
git -C $script:tempRoot -c core.autocrlf=false add -A | Out-Null
git -C $script:tempRoot -c core.autocrlf=false commit -q -m 'prime' | Out-Null
```

Without this, a Windows `core.autocrlf=true` would normalise the
working tree to CRLF and Phase 2's `--quiet` comparison would
falsely report DRIFT DETECTED even on a clean baseline. The "stale"
test case also commits its deliberately-corrupted `SKILL.md` before
invoking the gate, so Phase 1 (dirty tree) passes and only Phase 2
(drift) fires.

### Write-Host vs Write-Error capture

`Start-Process -RedirectStandardOutput` captures only the host's
standard output stream. In PowerShell 7, `Write-Host` goes through
the host's InformationStream which is *usually* wired to stdout
under redirection, but the immediate `Write-Error` + `exit 1` after
a `ForEach-Object { Write-Host $_ }` block can race the flush. The
"missing phrase" assertion in `check-skill-summary.Tests.ps1`
therefore matches the reliably-captured `violation(s)` count line
rather than the per-violation detail. The detail is still
observable in non-redirected runs; only the redirected test capture
is affected.


## Residual risks (record for the next pass)

1. **Lost case-sensitivity tripwire on `check-skill-summary.ps1`.**
   The old `Should -Match '-cnotmatch'` assertion was a string-match
   against the script source. It no longer exists. If a future
   refactor replaces `-cnotmatch` with `-notmatch`
   (case-insensitive), the clean/bad fixture trees will still differ
   in exit code and the test will not catch it.

2. **Lost path-traversal regression guard on `check-status.ps1`.**
   The old negative-match `'\\.StartsWith\\(\\$repoRootFull'` was
   guarding the `GetFullPath` + `GetRelativePath` fix that replaced
   a `StartsWith` containment check. The new test cases cover the
   "missing path" case, but a deliberate `..\\evil-repo` case is not
   in the suite. Mitigation candidate: add one `It` whose status
   body contains a backticked path like `..\\evil-repo\\STATUS.md`
   and assert that the script rejects it with the
   `outside the repository` violation.

3. **Test suite wall time.** The four updated files now spawn
   ~14 child-`pwsh` processes; the full `scripts/test-scripts.ps1`
   suite went from ~2 s to ~30 s. Acceptable for the in-development
   gate, but if the suite grows further, consider a single pwsh
   process per harness with multiple invocations.

4. **Temp-root path-length brittleness on Windows.** The harness
   creates deeply-nested paths
   (`$tempRoot/harness/.cline/skills/...`). On long Windows temp
   paths (e.g. nested `TEMP` env vars), `git` can hit MAX_PATH. Not
   observed in the current environment; record it as a known
   limitation if it surfaces on a CI runner.

5. **PowerShell `BeforeEach` cleanup of the temp dir is best-effort.**
   `Remove-Item -ErrorAction SilentlyContinue` swallows the case
   where the dir was already removed or partially-removed. If a
   future test changes the temp dir to a network share, the cleanup
   will hang indefinitely. Not a concern for the current
   `GetTempPath()` use.

6. **Diagnostic commit artefact.** The "stale" test in
   `check-cline-skills.Tests.ps1` issues `git commit -m 'stale'`
   inside the temp harness. If a future change ever runs the test
   in a directory that is itself a git repo with
   `commit.gpgsign=true`, the commit will fail unless the test
   environment has a signing key. The test runs in `GetTempPath()`
   (not a repo) but the `git config user.email` / `user.name` we
   set in `BeforeEach` apply to that temp repo only — so a
   `commit.gpgsign` inheritance from a parent `~/.gitconfig`
   *would* still bite. Not observed; flagged for visibility.


## Acceptance evidence observed

- `pwsh -NoProfile -File scripts/test-scripts.ps1` runs all six
  Pester suites and aggregates **28 / 28 pass** (check-cline-skills
  4, check-skill-summary 3, check-status 5, pre-commit 9,
  sync-cline-skills 3, test-locked-restore 4).
- Targeted re-run of just the four updated files via
  `Invoke-Pester -Path scripts/check-cline-skills.Tests.ps1,...`:
  **15 / 15 pass**.
- The two test files the change did *not* touch
  (`pre-commit.Tests.ps1` and `test-locked-restore.Tests.ps1`)
  still pass at their previous counts (9 / 9 and 4 / 4), confirming
  the harness pattern is compatible with the existing per-file
  test runner (`scripts/test-scripts.ps1`).

## In-development context (do not skip when re-reading)

- This code is **not** published. There is no released NuGet, no
  signed XLL, no customer data, no production deployment, no SLA.
- The verify scripts (`verify-quick.ps1`, `verify.ps1`) and the
  pre-commit hook are the only consumers of these tests; CI does
  not run them on a schedule. The script tests are a developer
  aid, not a regulated artefact.
- The local dev machine currently runs Pester 6.1.0; that is an
  **observed local version, not a repository pin**. The repository's
  compatibility contract is Pester 5+ (`Pester.MinimumMajor = 5` in
  `scripts/tool-versions.psd1`, enforced by `scripts/test-scripts.ps1`
  via `Get-Module -ListAvailable` + minimum-major check; CI installs
  newest >= 5.0). Pester 5-only constructs (e.g. `[Diagnostics.CodeAnalysis]`
  suppressions) are not used, and the version-safe container counting in
  `test-scripts.ps1` works across the supported range. Exact version
  parity across dev/CI remains deferred by design (L9).
- `git status --short` shows several pre-existing uncommitted
  modifications unrelated to this refactor (e.g. `docs/STATUS.md`,
  `scripts/pre-commit.ps1`, `scripts/test-locked-restore.ps1`).
  Those are *not* in scope for this evidence file.
- The drift-gate fixture `99-FIXTURE.md` and the test-helper
  `99-fixture/SKILL.md` live entirely under
  `[System.IO.Path]::GetTempPath()/<guid>/`; they never touch the
  real `docs/` or `.cline/skills/` trees in the repository.
- All script-test code is scaffold-grade: hand-rolled, not part of
  any public API, and free to be rewritten at the next refactor of
  the gate. Future agents should feel no obligation to preserve
  these tests verbatim; the value is the harness pattern (temp root,
  child pwsh, observable assertions), not the literal `It` names.
- The harness design is intentionally simple and easy to replace. There
  is no backwards-compat promise, but the tests must stay at the
  behaviour level (child-process output, exit codes, produced files).
  Do not simplify them back down to tripwire string matches against
  script source: a match on literal names asserts nothing about
  behaviour. If behavioural testing becomes too costly to maintain,
  obtain an approved policy change (per the AGENTS.md stop conditions)
  before weakening the gate, rather than silently regressing to text
  assertions.

## Suggested follow-up (not a roadmap item)

- Add the path-traversal `It` case to `check-status.Tests.ps1`
  (risk 2).
- Add a behaviour-level assertion to `check-skill-summary.Tests.ps1`
  that a case-sensitivity mismatch fixture (x vs X) changes the gate
  outcome (risk 1); do not reintroduce a text-level tripwire.
- Consider extracting a `New-ScriptHarness -RepoRoot ...` helper
  into `scripts/test-helpers.ps1` once a fifth script test is
  added, to keep the `BeforeEach` blocks from drifting in style.

The above are suggestions, not commitments. The project owner is
expected to revisit this evidence file at the next refactor of the
gate, and to treat anything in "Residual risks" as a candidate to
re-investigate rather than a known-good design.

