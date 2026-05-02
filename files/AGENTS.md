# Global Agent Rules

## Language Policy

- Default to Portuguese for user-facing conversation, reasoning summaries, final answers, file comments, documentation prose, generated reports, variable names, function names, class names, identifiers, and generated code text.
- Exception: when the user asks to update configuration, instruction, policy, or agent-rule files, use the language requested by the user or the dominant language already used by that configuration file.
- Preserve source language when quoting, citing, translating only on request, editing text whose language is part of the deliverable, or modifying existing code where local naming conventions are already in another language.

## Precedence

- Project-level and directory-level `AGENTS.md` files may add stricter project-specific rules.
- Project-level rules must not weaken global safety, MCP, cleanup, or subagent rules.
- If rules conflict, follow the most specific rule that does not weaken global safety requirements.

## File Encoding

- When writing files, use UTF-8 encoding.

## Codex project initialization

When I ask to initialize/start/bootstrap a project, do this before project work:

1. Run:

`powershell -ExecutionPolicy Bypass -File "$env:USERPROFILE\.codex\scripts\bootstrap-codex-project.ps1" -ProjectPath "<current working directory>"`

2. Show available MCP presets from `~/.codex/presets/mcp`.
3. Ask which presets should be enabled.
4. Generate or update the project-local `.codex/config.toml`.
5. Keep global `~/.codex/config.toml` clean.
6. Never manually copy MCP blocks into the global config unless explicitly requested.

## Main thread and subagents

Keep the main thread clean.

Prefer subagents for execution-heavy or parallelizable work, especially when a task could keep the main thread blocked for too long.

Small tasks can stay in the main thread when that is the most efficient path.

Default subagent settings:
- model: gpt-5.4-mini
- reasoning effort: medium

Do not use a stronger model for subagents unless explicitly requested.

## Main Thread Orchestration

- Keep the main thread clean.
- The main agent acts primarily as an orchestrator:
  - clarify scope only when strictly necessary;
  - create a short execution plan;
  - delegate execution to subagents whenever the task is heavy, parallelizable, or likely to block the main thread for too long;
  - after delegating, continue with independent work in the main thread;
  - wait only when the subagent result is needed to unblock the next step;
  - consolidate results;
  - return only the final useful answer.
- Do not fill the main thread with raw logs, exploratory notes, long file dumps, stack traces, or intermediate tool output.
- Use the main thread for final integration, conflict reconciliation, and concise reporting.

## Subagents, Parallel Work, And MCP Safety

- For executable work, prefer subagents.
- Use subagents for:
  - code exploration;
  - implementation;
  - tests;
  - debugging;
  - refactoring;
  - documentation updates;
  - PR/review checks;
  - dependency or configuration investigation;
  - any task that can be split into independent subtasks.
- If a task can be parallelized, split it across multiple subagents.
- For unspecified tasks, prefer these roles:
  - `explorer`: read/search/map context;
  - `worker`: scoped implementation;
  - `reviewer`: validation/review;
  - `default`: fallback for any other delegated task.
- When spawning subagents without explicit names, use the local default subagent configuration:
  - model: `gpt-5.4-mini`;
  - reasoning effort: `medium`;
  - max parallel agents: `5`;
  - max depth: `1`.
- Do not use a stronger model for subagents unless explicitly requested.
- Any non-specialized subagent must run with zero MCP enabled by default.
- Before spawning a generic worker, utility agent, benchmark agent, infra agent, Docker agent, GPU agent, shell agent, coding worker, or exploratory subagent that does not belong to a specialized workflow with explicit MCP needs, disable all MCPs.
- The default reusable overlay for zero-MCP subagents is `C:\Users\imale\.codex\overlays\no-mcp.toml`.
- Only specialized subagents may run with MCP enabled, and even then only with the minimum MCP surface required for that role.
- If a specialized subagent does not need MCP in a given run, it must also run with zero MCP.
- When retrieving subagent output, prefer agent lifecycle tools first: use `resume_agent`, `wait_agent`, or direct agent status/result retrieval before reading artifacts from disk.
- Do not treat on-disk artifacts as the primary source for a subagent result if the agent can still be resumed or its final status can still be queried directly.
- Read files written by a subagent only as supporting evidence, detailed inspection material, or fallback when the agent lifecycle result is unavailable.
- Exception: when a workflow explicitly uses handoff artifacts as its designed delivery mechanism, the handoff on disk may be treated as the primary output for that specific flow.
- For parallel tasks:
  1. Define a bounded scope for each subagent.
  2. Avoid overlapping file edits where possible.
  3. Wait for all required agents.
  4. Reconcile conflicts in the main thread.
  5. Return changed files, tests run, and remaining issues.

## Worktree Integration Queue

- When working with multiple Git worktrees in parallel, use an integration queue instead of merging everything at once.
- Integrate only one worktree at a time into the current integration branch.
- As soon as the first worker finishes, bring that tree first and update the integration branch before accepting the next worktree.
- After each queued integration, all remaining worktrees must rebase, cherry-pick, or otherwise resolve conflicts against the updated integration branch before their trees are brought in.
- Keep shared-file conflict resolution concentrated in the queued integration step rather than letting several workers edit the same integration targets concurrently.

## Output Policy

- Keep final responses concise.
- Include only what is useful:
  - files changed;
  - what changed;
  - validation/tests;
  - pending risks or next action.
- Avoid long explanations unless requested.
