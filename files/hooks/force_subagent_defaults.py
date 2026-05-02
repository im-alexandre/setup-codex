import json
import sys

json.load(sys.stdin)

print(
    json.dumps(
        {
            "hookSpecificOutput": {
                "hookEventName": "SessionStart",
                "additionalContext": """
When spawning any subagent without an explicit model, use:
- model: gpt-5.4-mini
- reasoning effort: medium

Do not use a stronger model for subagents unless the user explicitly requests it.
For arbitrary subagent tasks, use the local default/explorer/worker/reviewer agent configs.
""",
            }
        }
    )
)
