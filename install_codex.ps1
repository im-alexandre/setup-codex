$ErrorActionPreference = "Stop"

function Get-ScriptRoot {
    if ($PSCommandPath) {
        return (Split-Path -Parent $PSCommandPath)
    }

    if ($MyInvocation.MyCommand.Path) {
        return (Split-Path -Parent $MyInvocation.MyCommand.Path)
    }

    throw "Nao foi possivel determinar o diretorio do script."
}

$RepoRoot = [System.IO.Path]::GetFullPath((Get-ScriptRoot)).TrimEnd("\", "/")
$CodexHome = [System.IO.Path]::GetFullPath((Join-Path $env:USERPROFILE ".codex")).TrimEnd("\", "/")

if (-not $RepoRoot.Equals($CodexHome, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "Este repositorio deve ser usado diretamente como CODEX_HOME: $CodexHome"
}

$SharedConfigPath = Join-Path $RepoRoot "config.shared.toml"
$ConfigPath = Join-Path $RepoRoot "config.toml"

if (!(Test-Path -LiteralPath $SharedConfigPath)) {
    throw "config.shared.toml nao encontrado: $SharedConfigPath"
}

if (!(Get-Command python -ErrorAction SilentlyContinue)) {
    throw "Python nao encontrado; a instalacao precisa de python com tomllib para gerar config.toml."
}

$MergeScript = @'
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
'@

$MergeScript | python - $SharedConfigPath $ConfigPath

if ($LASTEXITCODE -ne 0) {
    throw "Falha ao gerar config.toml a partir de config.shared.toml."
}

Write-Host "CODEX_HOME versionado pronto: $RepoRoot"
