import json
import sys

json.load(sys.stdin)

# Não bloqueia nada. Só injeta política operacional.
instruction = """
Execution policy for this turn:

Keep the main thread clean.

Prefer subagents for heavy, parallelizable, or multi-step work, especially when the
task can block the main thread for too long.

Small, quick tasks may stay in the main thread when that is the most efficient path.

After delegating work, continue with independent work in the main thread instead of
waiting reflexively.

Use wait_agent only when the result is a blocker for the next step.

Default subagent settings:
- model: gpt-5.4-mini
- reasoning effort: medium

Main agent responsibilities:
1. create a short plan;
2. split the work into bounded subagent tasks when useful;
3. continue with independent main-thread work after delegation;
4. wait only when the result is needed to unblock the next step;
5. consolidate results;
6. return only the useful final summary.

Avoid raw logs, long exploration notes, and intermediate dumps in the main thread.
"""

print(
    json.dumps(
        {
            "hookSpecificOutput": {
                "hookEventName": "SessionStart",
                "additionalContext": instruction,
            }
        }
    )
)
