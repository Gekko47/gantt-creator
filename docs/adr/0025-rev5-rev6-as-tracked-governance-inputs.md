# ADR-0025 — REV5 and REV6 enter the source-of-truth order as tracked inputs

- **Status:** Accepted
- **Date:** 2026-09-29
- **Relates to:** `AGENTS.md` source-of-truth order
- **Inputs:** [`GanttCreator_R3_R4_Architecture_Roadmap_Review_REV5.md`](../GanttCreator_R3_R4_Architecture_Roadmap_Review_REV5.md), [`GanttCreator_Revised_Implementation_Plan_REV6.md`](../GanttCreator_Revised_Implementation_Plan_REV6.md)
- **Register:** [`PRE-R4.9-DECISION-REGISTER.md`](../PRE-R4.9-DECISION-REGISTER.md)

## Context

REV5 and REV6 describe the pre-R4.9 product corrections: parent/child hierarchy,
projected children, Critical Interval as a real child, uniform row geometry,
hidden engine metadata, derived Duration, Ribbon Row Styling, and a Refresh
orchestrator. They are the direct source of the R4.7A-R4.8A rows.

Two problems made them unusable as-is.

**First, they sat outside the source-of-truth order.** `AGENTS.md` orders
requirements as: work item → product invariants and entity guide → ADRs →
roadmap → code. Two documents in neither list are read as authority by anyone,
and an agent handed one must stop at the first conflict rather than implement.

**Second, and more seriously, they conflict with the landed tree in four places.**
Each conflict is a place where following the document literally would delete
working code or reverse a decision the product owner had already made. They are
itemised in the register; the two that would have done real damage:

- REV5 §7 and REV6 §9 require the critical interval to stop being a rectangle.
  The entity guide (line 736) says the opposite, deliberately, and the R3.16
  equivalence field table encodes that reading.
- REV5 §16 and REV6 §16 require a bounded `HostShapeKey`. R4.7's ledger already
  records that proposal as **rejected on evidence and closed as L18**, with the
  product owner approving on 2026-09-28.

A document that is half-binding cannot be read without a per-clause judgement
call. That is the drift the owner asked to be prevented.

## Decision

- **D1 — Both documents are tracked governance inputs.** `.gitignore` excluded
  them, following this repository's existing convention for untracked input
  briefs (see the `REMEDIATION-PLAN-A` comment in that file). That exclusion is
  reversed. A tracked document that is cited by roadmap rows is a requirement;
  an untracked one is not, and the roadmap would have cited something no reader
  could resolve.
- **D2 — They are added to `AGENTS.md`'s source-of-truth order, below the entity
  guide and above the roadmap.** Below the guide is load-bearing: decisions D1,
  D2 and D4 in the register all *revise* the guide, and a document ranked above
  it would let an implementer treat a guide revision as optional.
- **D3 — Each file carries a header table naming the proposals that were changed
  or rejected**, so a reader who opens either document first learns what
  survived without reading the register.
- **D4 — The register is the authority where the two conflict with the
  resolved decisions.** It is a short document, and it is linked from both.
- **D5 — The conflicts were resolved by the product owner, not by this ADR.**
  Where REV5/REV6 were right, the guide is revised. Where they were wrong or
  where they reversed a settled decision, they are marked as rejected and the
  register records why.

## Consequences

- Nine roadmap rows acquire a citable source, and the "stop and report the
  conflict" path in `AGENTS.md` has a defined answer for the four conflicts that
  would otherwise have stopped an implementer cold.
- The entity guide still needs revisions for fixed row heights, the critical
  rectangle, and the label set. That is ADR-0026, ADR-0027 and ADR-0028, not
  this ADR.
- The register must be updated whenever a resolved decision changes, or it
  becomes a third source with the same drift problem. It is a short document for
  that reason.

## Alternatives considered

- **Treat the documents as untracked inputs, as `.gitignore` had them.** Rejected:
  it is the status quo that produced the conflict, and the roadmap would cite
  unresolvable requirements.
- **Make them the top of the source-of-truth order.** Rejected: three of the four
  conflicts require *changing* the entity guide, so a document ranked above it
  would authorise an implementer to skip a guide revision.
- **Edit REV5/REV6 in place to remove the rejected proposals.** Rejected: the
  rejected proposals are the product's reasoning, and deleting them destroys the
  record of what was considered. A header table is non-destructive and keeps the
  alternatives visible.
- **No register; rely on the ADRs.** Rejected: an implementer does not read five
  ADRs to find out whether the document in front of them is binding. The register
  answers the single question "is this proposal live?"
