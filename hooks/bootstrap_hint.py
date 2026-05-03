import json
import sys
from pathlib import Path

payload = json.load(sys.stdin)

prompt = (payload.get("prompt") or "").lower()
cwd = Path(payload.get("cwd") or ".").resolve()

markers = [
    "inicia o projeto",
    "iniciar o projeto",
    "bootstrap",
    "setup codex",
    "configurar codex",
    "criar .codex",
    "cria .codex",
]

if any(m in prompt for m in markers):
    script = Path.home() / ".codex" / "scripts" / "bootstrap-codex-project.ps1"

    print(json.dumps({
        "hookSpecificOutput": {
            "hookEventName": "UserPromptSubmit",
            "additionalContext": f"""
Project bootstrap requested.

Before doing any project work, run:

powershell -ExecutionPolicy Bypass -File "{script}" -ProjectPath "{cwd}"

This script must:
1. inspect available MCP presets in ~/.codex/presets/mcp;
2. show them to the user;
3. ask which MCPs should be enabled;
4. create or update the project-local .codex/config.toml;
5. keep the global ~/.codex/config.toml clean.

Do not manually edit the global config unless explicitly requested.
"""
        }
    }))
