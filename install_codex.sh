#!/usr/bin/env bash
set -euo pipefail

script_dir="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
repo_root="$(realpath "$script_dir")"
codex_home="$(realpath "${HOME}/.codex")"

if [[ "${repo_root}" != "${codex_home}" ]]; then
  printf 'Este repositorio deve ser usado diretamente como CODEX_HOME: %s\n' "$codex_home" >&2
  exit 1
fi

shared_config_path="${repo_root}/config.shared.toml"
config_path="${repo_root}/config.toml"

if [[ ! -f "$shared_config_path" ]]; then
  printf 'config.shared.toml nao encontrado: %s\n' "$shared_config_path" >&2
  exit 1
fi

if ! command -v python3 >/dev/null 2>&1; then
  printf 'python3 nao encontrado; a instalacao precisa de python3 com tomllib para gerar config.toml.\n' >&2
  exit 1
fi

python3 - "$shared_config_path" "$config_path" <<'PY'
import pathlib
import re
import sys
import tomllib

shared_path = pathlib.Path(sys.argv[1])
config_path = pathlib.Path(sys.argv[2])

project_header = re.compile(r"^\s*\[projects(?:\.|\])")
table_header = re.compile(r"^\s*\[")

def extract_project_blocks(text: str) -> list[str]:
    lines = text.splitlines()
    blocks: list[str] = []
    block: list[str] | None = None

    for line in lines:
        if project_header.match(line):
            if block:
                blocks.append("\n".join(block).rstrip())
            block = [line]
            continue

        if block is not None and table_header.match(line):
            blocks.append("\n".join(block).rstrip())
            block = None

        if block is not None:
            block.append(line)

    if block:
        blocks.append("\n".join(block).rstrip())

    return [block for block in blocks if block]

shared = shared_path.read_text(encoding="utf-8")
tomllib.loads(shared)

projects = []
if config_path.exists():
    existing = config_path.read_text(encoding="utf-8")
    tomllib.loads(existing)
    projects = extract_project_blocks(existing)

merged = shared.rstrip() + "\n"
if projects:
    merged += "\n" + "\n\n".join(projects).rstrip() + "\n"

tomllib.loads(merged)
config_path.write_text(merged, encoding="utf-8")
print("config.toml gerado a partir de config.shared.toml")
print(f"projects preservados: {len(projects)}")
PY

printf 'CODEX_HOME versionado pronto: %s\n' "$repo_root"
