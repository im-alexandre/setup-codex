# Global Agent Rules

## Language Policy

- Default to Portuguese for user-facing conversation, reasoning summaries, final answers, file comments, documentation prose, generated reports, variable names, function names, class names, identifiers, and generated code text.
- Exception: when the user asks to update configuration, instruction, policy, or agent-rule files, use the language requested by the user or the dominant language already used by that configuration file.
- Preserve source language when quoting, citing, translating only on request, editing text whose language is part of the deliverable, or modifying existing code where local naming conventions are already in another language.

## Precedence

- Project-level and directory-level `AGENTS.md` files may add stricter project-specific rules.
- Project-level rules must not weaken global safety, MCP, or cleanup rules.
- If rules conflict, follow the most specific rule that does not weaken global safety requirements.

## File Encoding

- When writing files, use UTF-8 encoding.

## DOCX Revision And Comment Authors

- For any reading, inspection, editing, rewriting, validation, synchronization, revision, comment, or other operation that touches a `.docx`, use the `docx-utils` skill/tooling.
- Execute `docx-utils` through the published binary/shim by default; do not use `dotnet run --project` unless developing, debugging, or recovering from a broken/missing binary.
- If the needed `docx-utils` capability fails, run `docx-utils --help` to review available commands, calling forms, and examples.
- If no command exists for the needed DOCX operation, log the missing capability in the `docx-utils` skill backlog for future implementation by the skill maintainer/agent maintainer.
- When the main thread adds a revision or comment to a `.docx` through `docx-utils`, it must omit `--author`; `docx-utils` automatically chooses the next available author in the document.
- Subagents that add revisions or comments to a `.docx` must pass `--author` explicitly with the assigned subagent name.

## Codex project initialization

When I ask to initialize/start/bootstrap a project, do this before project work:

1. Run:

`powershell -ExecutionPolicy Bypass -File "$env:USERPROFILE\.codex\scripts\bootstrap-codex-project.ps1" -ProjectPath "<current working directory>"`

2. Show available MCP presets from `~/.codex/presets/mcp`.
3. Ask which presets should be enabled.
4. Generate or update the project-local `.codex/config.toml`.
5. Keep global `~/.codex/config.toml` clean.
6. Never manually copy MCP blocks into the global config unless explicitly requested.

## Output Policy

- Keep final responses concise.
- Include only what is useful:
  - files changed;
  - what changed;
  - validation/tests;
  - pending risks or next action.
- Avoid long explanations unless requested.
