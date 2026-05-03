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

$ConfigPath = Join-Path $RepoRoot "config.toml"

if (!(Test-Path -LiteralPath $ConfigPath)) {
    throw "config.toml nao encontrado: $ConfigPath"
}

if (Get-Command python -ErrorAction SilentlyContinue) {
    python -c "import pathlib, tomllib; tomllib.loads(pathlib.Path(r'$ConfigPath').read_text(encoding='utf-8')); print('config.toml ok')"
}

Write-Host "CODEX_HOME versionado pronto: $RepoRoot"
