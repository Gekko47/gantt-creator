# Work item — repository navigation map

> Documentation maintenance item requested by the user. It is **not** a new
> roadmap milestone and adds no product behaviour. Section structure follows
> [TEMPLATE.md](TEMPLATE.md).

## Outcome

Agents can discover a concise, source-backed repository map from the agent
contract, then obtain current implementation detail through scoped live symbol
queries instead of a committed symbol dump.

## Scope

- `../../AGENTS.md` — startup navigation link and same-change maintenance rule.
- `../REPO-MAP.md` — the new maintained map.
- `../STATUS.md` — handoff entry.
- this file.

## Exclusions

Anything not listed above. Specifically: production or test behaviour, packages,
workbook schema, roadmap content, other work items, CI configuration, and the
separate local MCP-server repository. No generated symbol dump, background
indexer, persistent cache, or automatic semantic freshness checker is added, and
no new validator script is introduced.

## Acceptance tests

- **Discovery** — `../../AGENTS.md` links to the map, and the link resolves under
  the existing markdown gate `../../scripts/check-md-links.ps1`.
- **Existing link-gate controls** — `../../scripts/check-md-links.Tests.ps1`
  still passes, including its failure-path control.
- **Source accuracy** — the four production projects, their project references,
  and the six test projects are checked against `../../GanttCreator.slnx` and the
  linked project files, then re-read from the current tree.
- **Honest capability boundary** — the map states that scene/layout, chart
  rendering, PowerPoint transfer, and PNG encoding are planned rather than
  implemented, and points to `../02-ARCHITECTURE.md` and `../03-ROADMAP.md`.
- **Scoped symbol guidance** — the map states that outlines are locators, not
  compiler or reference evidence, and that truncated results must be narrowed and
  re-checked against source.
- **Office host / visual** — not applicable to a documentation-only change.

## Evidence commands

Run from the repository root.

```powershell
pwsh -NoProfile -File C:\repos\gantt-creator\scripts\check-md-links.ps1
pwsh -NoProfile -Command '$r = Invoke-Pester -Path C:\repos\gantt-creator\scripts\check-md-links.Tests.ps1 -PassThru; if ($r.FailedCount -gt 0 -or $r.TotalCount -eq 0) { exit 1 }'
git -C C:\repos\gantt-creator add AGENTS.md docs/REPO-MAP.md docs/STATUS.md docs/work-items/REPO-MAP-navigation.md
pwsh -NoProfile -File C:\repos\gantt-creator\scripts\verify-quick.ps1   # every-commit gate; expect exit 0
git -C C:\repos\gantt-creator diff --cached --check
git -C C:\repos\gantt-creator diff --cached
# after the commit, with a fully clean tree:
pwsh -NoProfile -File C:\repos\gantt-creator\scripts\verify.ps1
```

`check-md-links.ps1` proves that link targets exist, not that the mapped
responsibilities or dependency claims are correct; source review is the evidence
for those. `verify-quick.ps1` accepts staged changes; `verify.ps1` requires a
clean committed tree.

## Risk and rollback

Risk: a stale map mistaken for requirements, or for exhaustive code coverage.
Mitigations: the map is explicitly navigation-only and subordinate to the
source-of-truth order; mapped changes carry a same-change update rule; live symbol
queries are scoped and re-checked against source. Rollback: remove the map, its
status entry, and the single contract link; no data, schema, or behaviour is
affected.

## Definition of done

Both documentation gates pass on the committed content; the project, reference,
and test claims were re-read from the current tree; no application, schema,
dependency, or roadmap state changed; the diff is documentation only; the STATUS
entry records only observed evidence; no test was weakened and no warning
suppressed.

## Notes during implementation

Sequence actually followed:

1. Read the agent contract, solution, production project files, entry points, test
   directories, and the verify/status/link scripts; confirmed working-tree state.
2. Wrote the map from verified source: project table, task routing, capability
   boundary, verification commands, and maintenance rules.
3. Linked it from the agent contract with a same-change maintenance rule.
4. Recorded the handoff in `../STATUS.md`.
5. Validated links, ran the every-commit gate, reviewed the staged diff, committed
   the documentation-only change, then ran the branch-final gate on the clean tree.

Deliberate content choices:

- No line numbers, test counts, or copied schemas: they drift fastest, and the
  link gate cannot detect semantic drift.
- Relative links for checkout portability, resolved against this workspace's
  absolute root when tools are run.
- `ExportSize` is described as width parsing and dimension maths, not as a PNG
  encoder, because no raster pipeline exists yet.
- The Office adapter is described as an application-fact port, not as a worksheet
  or chart renderer.

### Deferred second stage — cached MCP index (not implemented here)

Out of scope for this documentation commit: it belongs to the separate local
MCP-server repository and needs its own handler, subprocess-wrapper, and test
review first. If pursued, the minimum bar is: a versioned cache keyed by
workspace, relative path, content hash, parser version, and index options;
refresh by content hash rather than HEAD or timestamps, so uncommitted and
same-size edits are caught; build/cache directory exclusions applied before the
parser runs; storage kept separate from response limits so queries can be paged
with total/matched/returned counts plus freshness and truncation metadata;
argument arrays and atomic replacement instead of shell interpolation or a
hardcoded platform temp path; explicit diagnostics on cache, parser, or timeout
failure instead of an empty success; and tests over real parser fixtures for
cold/warm equality, renames and deletions, exclusion behaviour, pagination
boundaries, stale cursors, and corrupt-cache recovery. It supplements rather
than replaces this map, and it still does not resolve types, references, or
semantics.

