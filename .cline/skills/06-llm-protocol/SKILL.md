---
name: 06-LLM-PROTOCOL
description: Required session opening, evidence ledger, anti-drift and anti-hallucination controls, the two-attempt no-loop protocol, context discipline, coding behaviour, prompt pattern, and the agent handoff format; plus the full-repo code-review methodology. Trigger on protocol, evidence, retry, or handoff-format questions.
---

The LLM operating protocol for single-work-item tasks (session
opening, evidence ledger, anti-drift/anti-hallucination controls, the
two-attempt no-loop rule, coding behaviour, prompt pattern, handoff
format) AND a separate methodology for a full-repository code review
(six phases: read the source, probe empirically, classify defects
into the four repo-specific classes, plan, implement/verify/commit,
self-apply).

Do not get wrong:
- Two failed attempts on the same failure = stop and report; never a
  third unchanged retry.
- An inference cannot become a fact by being repeated — use the
  `fact`/`inference`/`proposal`/`unknown` ledger honestly.
- A full-repo review must look for gate-integrity, build-pipeline
  drift, docs/code drift, and local/CI divergence even when the
  per-item checklists already pass — those are cross-cutting defect
  classes the per-item gates do not catch by construction.

---

## Tools

- `memra_add` / `memra_add_decision` — persist evidence-ledger rows and irreversible decisions across turns.
- `memra_bootstrap` — recall prior decisions and patterns at session start.
- `memra_add_pattern` — store reusable methodologies (e.g. the full-repo review phases).
- `memra_search` / `memra_recall` — retrieve prior evidence by semantic similarity.
- `sequential-thinking__sequentialthinking` — structure multi-step reasoning (review methodology, failure classification).
- `vscode-mcp__get_symbol_lsp_info` / `vscode-mcp__get_references` — get compiler-grade type/symbol info instead of guessing APIs.
- `ilspy_decompile` — inspect installed assembly metadata when source is unavailable.
- `context7__query-docs` / `microsoft-learn__microsoft_docs_search` — look up unfamiliar .NET / Excel-DNA / SkiaSharp APIs; preferred over memory.
- `dotnet_build` / `dotnet_test` — compile and run tests; "Not run" must be stated if skipped.
- `semgrep_scan` / `ast_grep_search` — find pattern violations during review.

---

## Where to read more

- Full reference: `docs/06-LLM-PROTOCOL.md` (this is the canonical source; there is no separate copy under .cline/skills/)
