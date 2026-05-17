# Prompt Contracts For `$implement-tdd`

Use these snippets when preparing subagent dispatches. Keep prompts bounded: include only task text, relevant paths, test commands, constraints, and acceptance criteria.

## Coordinator Matrix

The main thread produces:

```markdown
| Task | Agent | Depends on | Write scope | Red command | Green command | Integration order |
| --- | --- | --- | --- | --- | --- | --- |
```

Rules:

- Mark same files, migrations, shared contracts, routers, schemas, fixtures, and test harness changes as sequential.
- Dispatch only one wave of non-overlapping tasks at a time.
- Put shared setup, dependency installation, generated clients, migrations, and contract changes before dependent tasks.

## Implementer Prompt Shape

```text
You are the <agent-name> implementer for task <task-id>.

Use strict TDD:
1. write/update tests first;
2. include happy path and relevant error-path coverage;
3. run the red command and record the failing signal;
4. implement the smallest production change;
5. run the green command and record the passing signal;
6. refactor only if it generalizes or removes real duplication;
7. append JSONL events to .codex/agent-events/events.jsonl;
8. write final result to .codex/agent-events/results/<task-id>.md.

Task:
<full bounded task text>

Write scope:
<paths>

Commands:
- Red: <command>
- Green: <command>
- Integration: <optional command>

Do not edit outside the write scope unless blocked; report blocked if scope is wrong.
```

## Aggregator Prompt Shape

```text
Review the completed $implement-tdd work read-only unless explicitly asked to integrate.

Check:
- TDD evidence includes red and green;
- tests cover happy path and relevant error paths;
- changed files stay inside intended scope;
- no unresolved blocked/error events remain;
- targeted and integration validation commands were run or clearly could not run.

Return findings first, then validation summary, then remaining risks.
```
