# Superpowers: writing-plans

Use this command only when explicitly invoked as a prompt command.

Load and follow the workflow described in:

`C:\Users\imale\.codex\vendor_imports\superpowers\skills\writing-plans\SKILL.md`

User request / arguments:

`$ARGUMENTS`

Expected behavior:

- Create a concrete implementation plan for the requested work.
- Prefer small, testable tasks.
- Make task boundaries explicit enough for parallel execution when possible.
- Preserve TDD ordering: failing test first, implementation second, verification third.
- Do not implement the plan unless the user explicitly asks for execution.
