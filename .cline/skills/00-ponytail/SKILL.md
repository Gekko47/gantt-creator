---
name: 00-PONYTAIL
description: Lazy-senior-dev minimal-diff discipline (YAGNI ladder: reuse before rewrite, stdlib/native/installed-dep before new code, root-cause over symptom fix, no unrequested abstraction, one runnable check for non-trivial logic). The FIRST skill: run it before any other skill; it drives every task. Trigger on session start, before writing any code, or when deciding how small a change should be.
---

Lazy-senior-dev discipline: the best code is the code never written.
Climb the ladder (YAGNI -> reuse the existing helper -> stdlib ->
native platform -> installed dependency -> one line -> minimal code)
only AFTER reading the task and tracing the real flow end to end.
This is the FIRST skill: it runs before any other skill and drives
every task.

Do not get wrong:
- The ladder never replaces understanding. The smallest change in the
  wrong place is not lazy, it is a second bug — read the code the task
  touches first, then climb.
- A bug fix is a root-cause fix, not a symptom patch: grep every caller
  of the function you touch and fix the shared function once.
- Lazy code without its check is unfinished: non-trivial logic leaves
  ONE runnable check behind that follows the repository's testing
  rules and runs in the mandated test harness. Never get lazy about
  validation at trust boundaries, error handling, security,
  accessibility, hardware calibration, or anything explicitly
  requested.

---

## Tools

- `search_codebase` / `combined-mcp-server__grep_files` — find the existing helper, util, or pattern before writing new code; grep every caller before fixing a bug at its root.
- `read_files` / `combined-mcp-server__read_file` — read the task and the code it touches and trace the real flow end to end before picking a rung.
- `vscode-mcp__get_references` — list every caller of the symbol you are about to change so the fix lands once, at the shared function.
- `vscode-mcp__get_symbol_lsp_info` — confirm a symbol's real signature before assuming it does what you think.
- `dotnet_test` / `combined-mcp-server__dotnet_build` — run the one runnable check that non-trivial logic must leave behind.
- `combined-mcp-server__run_commands` — run the verification scripts (`verify-quick.ps1`) and observe their output.

---

## Where to read more

- Full reference: `docs/00-PONYTAIL.md` (this is the canonical source; there is no separate copy under .cline/skills/)
