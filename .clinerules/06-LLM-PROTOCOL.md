# LLM operating protocol

## Purpose

This protocol makes an LLM useful as a bounded engineering assistant. It does not delegate product ownership, evidence, release authority, or irreversible actions.

## Required session opening

Give the agent one work-item ID. The agent must respond with:

1. one-sentence outcome;
2. files/layers likely involved;
3. explicit exclusions;
4. acceptance tests it will use;
5. facts already verified;
6. remaining unknowns;
7. first safe action.

If this answer is wider than the work item, correct it before allowing edits.

## Evidence ledger

For interoperability investigations, maintain this compact table in the work item:

| Statement | Status | Evidence |
| --- | --- | --- |
| `Shapes.PasteSpecial` can request `ppPasteShape` | fact | Microsoft API documentation |
| copied group remains editable after staging cleanup | unknown until tested | compatibility spike R6.7 |
| exact live screen pixels are invariant across zoom/scaling | false | Excel uses point geometry and host rasterisation |

Allowed statuses are `fact`, `inference`, `proposal`, and `unknown`. An inference cannot become a fact because the agent repeats it.

## Anti-drift controls

- One active work item; one acceptance boundary.
- Fixed decisions live in architecture/ADRs, not chat memory.
- The agent reads only relevant files, then names them in its plan.
- Any new requirement is placed in a later work item unless necessary for the current acceptance test.
- Any public schema, dependency, supported-platform, or export-contract change stops for human approval.
- Refactoring is allowed only when the current slice needs it; preparatory refactors are separate green commits.
- At the end, compare the diff against the initial file/scope list and explain every deviation.

## Anti-hallucination controls

For every unfamiliar API or version-sensitive claim:

1. Search installed source/metadata or use Visual Studio Object Browser.
2. Check official Excel-DNA, Microsoft, .NET, SkiaSharp, Cline, or GitHub documentation.
3. Compile a minimal usage.
4. If host behaviour remains uncertain, write a disposable spike with one measurable assertion.
5. Keep the statement `unknown` until evidence exists.

The agent must not fabricate command output. If it cannot run a test, it says `Not run` and provides the exact command the human should run.

## No-loop protocol

Classify the failure first: compile, deterministic test, Office host, environment, permission, file lock, dependency, or requirement ambiguity.

Attempt 1:

- preserve exact error/HRESULT/stack trace;
- form one falsifiable hypothesis;
- run one discriminating check;
- make the smallest corresponding fix;
- rerun the narrow failed command.

Attempt 2:

- re-read the relevant boundary and primary documentation;
- form a materially different hypothesis;
- inspect environment/process/loaded-module evidence where relevant;
- make one different fix and rerun.

If the same failure occurs again, stop. Report:

- exact command and environment;
- shortest useful error excerpt;
- two hypotheses, tests, and outcomes;
- current uncommitted diff;
- likely layer and remaining unknown;
- one question or manual observation needed.

Prohibited “fixes” include repeated unchanged commands, larger sleeps, broad exception catches, test deletion, threshold reduction, warning suppression, package churn, random API substitution, and restarting everything without recording what the restart tests.

## Context discipline

- Start a fresh agent session for each roadmap item or after a large context shift.
- Keep `docs/STATUS.md` as the handoff; do not paste whole chat transcripts into the repository.
- Ask the agent to summarise current facts before context compaction.
- Do not feed customer workbooks or proprietary schedule descriptions to a hosted model.
- Use synthetic/minimised reproductions; redact file paths, names, dates, and labels from logs/prompts where needed.

## Coding behaviour

The agent should:

- inspect before editing;
- use existing patterns when tested and consistent with current architecture;
- prefer types and small functions to comments;
- add tests before or with the behaviour;
- run the narrow test after each meaningful edit;
- run full verification once the slice is coherent;
- leave the repository cleaner only within the touched scope.

The agent should not:

- design the whole system again in each task;
- create abstraction layers without two concrete consumers or a clear port boundary;
- generate speculative compatibility fallbacks;
- rewrite working code for style alone;
- introduce “temporary” hidden worksheets;
- claim pixel perfection for live Excel screen display;
- terminate Office processes it did not start.

## Prompt pattern

Use:

```text
Implement work item <ID> from docs/work-items/<file>.md.
Follow AGENTS.md and use the relevant project skill.
Before editing, restate outcome, exclusions, tests, facts, and unknowns.
Do not change architecture or dependencies.
Use the two-attempt no-loop rule.
Run targeted tests and scripts/verify.ps1, then report observed evidence only.
Do not commit until I review the diff.
```

Avoid broad prompts such as “build the Gantt app”, “finish the phase”, or “fix all errors”.

## Human review rhythm

The human reviews:

- the initial boundary before edits;
- the diff after each work item;
- every package addition and ADR;
- every golden-image change;
- every Office-hosted visual/clipboard/PowerPoint gate;
- each commit message;
- phase-exit and release evidence.

If review repeatedly finds the same error, update the relevant always-on rule or focused skill. Do not inflate every prompt with the lesson.

## Agent handoff format

```text
Outcome: <what now works>
Changed: <2-5 important files/areas>
Proof: <commands and observed results>
Office/manual: <PASS, FAIL, or Not run: reason>
Risk: <one sentence or None known>
Deferred: <explicitly excluded work>
Next: <roadmap ID>
```

This is intentionally short. The code, tests, work item, and Git diff are the durable record.
