# Three views of the same kit

> The Gantt Creator development kit has three views of the same seven
> documents. Each view has a different audience and a different loading
> rule. This file explains the relationship and the sync discipline.

## The three views

| View | Path | Audience | Loading rule |
| --- | --- | --- | --- |
| Canonical in-repo source | `docs/0N-*.md` | Human reviewers, the source of truth | Read on demand; this is the only view you edit by hand |
| Always-on rules | `.clinerules/0N-*.md` | Cline, Copilot, every LLM session | Loaded every turn; short, sharp, contract-level |
| On-demand skills | `.cline/skills/0N-name/SKILL.md` + `references.md` | Cline (on-demand), Copilot (when invoked) | Loaded when a Cline session matches the `description:` front-matter |

## Which one do I edit?

**Edit the canonical source** at `docs/0N-*.md`. The other two views
are derived from it. After editing, run `pwsh ./scripts/sync-cline-skills.ps1`
to refresh the always-on rule and the on-demand skill.

## What the sync script does

`scripts/sync-cline-skills.ps1`:

1. Reads every `docs/0N-*.md` listed in its `$map`.
2. For each, regenerates:
   - `.clinerules/0N-*.md` (a stable copy of the canonical body)
   - `.cline/skills/0N-name/SKILL.md` (front-matter + first 25 non-empty
     lines + footer with canonical pointers)
   - `.cline/skills/0N-name/references.md` (the full canonical body with
     a one-line header noting it is regenerated)
3. Validates that every skill directory has both `SKILL.md` and
   `references.md` and exits non-zero if any are missing.

The script is idempotent. Running it twice in a row produces the same
output.

## The drift gate

`scripts/check-cline-skills.ps1` runs the sync and then fails the gate
if a fresh sync produced any diff. Wire it into `verify-quick.ps1`
and `verify.ps1` so that drift fails the build.

```powershell
pwsh -NoProfile -File scripts/sync-cline-skills.ps1
$drift = git diff --quiet --exit-code .clinerules/ .cline/skills/ docs/
if ($LASTEXITCODE -ne 0) {
    Write-Error 'Skill tree is out of date. Run scripts/sync-cline-skills.ps1 and commit the result.'
    exit 1
}
```

## Why three views and not one

- **Always-on rules** must be short — they are scanned every turn, and
  a 45 KB entity guide would burn the agent's context window on every
  interaction. The always-on layer is the *contract*, not the *reference*.
- **On-demand skills** carry the full reference. The `description:` field
  is what Cline uses to decide whether to load the skill, so it must be
  specific enough to match the right skill in a single sentence and not
  so long that the loader wastes context on every turn.
- **The canonical mirror** is the only file you edit by hand. Two views
  read from it; one view is the kit the human reviewer reads.

## Adding a new rule

1. Add `docs/0N-name.md` with the new content.
2. Add an entry to the `$map` in `scripts/sync-cline-skills.ps1`
   (`Name`, `Description`).
3. Run `pwsh ./scripts/sync-cline-skills.ps1`.
4. Commit `docs/`, `.clinerules/`, and `.cline/skills/` together.

A new canonical source without a `$map` entry will be invisible to the
sync script and the gate will not catch it. A `$map` entry pointing at
a non-existent canonical source will fail-fast at sync time.
