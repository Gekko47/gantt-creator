---
name: 06-LLM-PROTOCOL
description: Use this skill when the conversation is about the LLM operating protocol: the required session opening, the evidence ledger, anti-drift controls, anti-hallucination controls, the two-attempt no-loop protocol, context discipline, coding behaviour, the prompt pattern, the human review rhythm, or the agent handoff format.
---

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

---

## Where to read more

- Canonical source: `docs/06-LLM-PROTOCOL.md`
- Always-on rule: `.clinerules/06-LLM-PROTOCOL.md`
- Full reference: `./references.md` in this directory
