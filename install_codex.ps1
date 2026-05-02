$ErrorActionPreference = "Stop"

function Get-ScriptRoot {
    if ($PSCommandPath) {
        return (Split-Path -Parent $PSCommandPath)
    }

    if ($MyInvocation.MyCommand.Path) {
        return (Split-Path -Parent $MyInvocation.MyCommand.Path)
    }

    throw "Não foi possível determinar o diretório do script."
}

function Convert-ToTomlPath {
    param([Parameter(Mandatory=$true)][string]$Path)
    return ($Path -replace "\\", "/")
}

function Set-Utf8NoBomContent {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][string]$Value
    )

    $encoding = [System.Text.UTF8Encoding]::new($false)
    [System.IO.File]::WriteAllText($Path, $Value, $encoding)
}

function Test-HasNvidiaGpu {
    $controllers = Get-CimInstance -ClassName Win32_VideoController -ErrorAction SilentlyContinue

    foreach ($controller in $controllers) {
        $values = @($controller.Name, $controller.Description, $controller.AdapterCompatibility, $controller.VideoProcessor)

        foreach ($value in $values) {
            if ($value -and $value -match 'NVIDIA') {
                return $true
            }
        }
    }

    return $false
}

function Get-RelativePathCompat {
    param(
        [Parameter(Mandatory=$true)][string]$BasePath,
        [Parameter(Mandatory=$true)][string]$Path
    )

    $baseFull = [System.IO.Path]::GetFullPath($BasePath).TrimEnd("\", "/") + [System.IO.Path]::DirectorySeparatorChar
    $pathFull = [System.IO.Path]::GetFullPath($Path)

    $baseUri = [System.Uri]::new($baseFull)
    $pathUri = [System.Uri]::new($pathFull)
    $relativeUri = $baseUri.MakeRelativeUri($pathUri).ToString()

    return ([System.Uri]::UnescapeDataString($relativeUri) -replace "/", "\")
}

function Get-BackupPath {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][string]$Stamp
    )

    $parent = Split-Path -Parent $Path
    $leaf = Split-Path -Leaf $Path
    return (Join-Path $parent "$leaf.bak-$Stamp")
}

function Should-SkipPath {
    param([Parameter(Mandatory=$true)][string]$RelativePath)

    $p = ($RelativePath -replace "\\", "/").TrimStart("/")

    $patterns = @(
        "^auth\.json$",
        "^cap_sid$",
        "^installation_id$",
        "^history\.jsonl$",
        "^session_index\.jsonl$",
        "^logs_2\.sqlite.*$",
        "^state_5\.sqlite.*$",
        "^models_cache\.json$",
        "^version\.json$",
        "^sandbox\.log$",
        "^\.codex-global-state\.json.*$",
        "^\.personality_migration$",
        "^\.sandbox(/|$)",
        "^\.sandbox-bin(/|$)",
        "^\.sandbox-secrets(/|$)",
        "^\.tmp(/|$)",
        "^sessions(/|$)",
        "^hook-debug(/|$)",
        "^\.hook-state(/|$)",
        "(^|/)\.git(/|$)",
        "(^|/)\.venv(/|$)",
        "(^|/)venv(/|$)",
        "(^|/)__pycache__(/|$)",
        "\.pyc$",
        "(^|/)node_modules(/|$)",
        "(^|/)dist(/|$)",
        "(^|/)build(/|$)",
        "(^|/)\.next(/|$)",
        "(^|/)\.pytest_cache(/|$)",
        "(^|/)\.mypy_cache(/|$)",
        "(^|/)\.codex(/|$)",
        "(^|/)\.env(\..*)?$",
        "(^|/)\.venv(/|$)",
        "(^|/)venv(/|$)",
        "(^|/)env(/|$)",
        "\.key$",
        "\.pem$",
        "secret",
        "token",
        "(^|/)data(/|$)",
        "(^|/)cache(s)?(/|$)",
        "(^|/)\.cache(/|$)",
        "(^|/)storage(/|$)",
        "(^|/)chroma(/|$)",
        "(^|/)chromadb(/|$)",
        "(^|/)db(/|$)",
        "(^|/)backups(/|$)",
        "(^|/)exports(/|$)",
        "\.sqlite3?$",
        "\.db$",
        "\.duckdb$",
        "(^|/)logs(/|$)",
        "\.log$",
        "\.bak(-.*)?$",
        "\.tmp(-.*)?$"
    )

    foreach ($pattern in $patterns) {
        if ($p -match $pattern) {
            return $true
        }
    }

    return $false
}

$RepoRoot = Get-ScriptRoot

# IMPORTANTE:
# Use "files" como pasta fonte para evitar confusão com ~/.codex.
$SourceRoot = Join-Path $RepoRoot "files"
$TargetRoot = Join-Path $env:USERPROFILE ".codex"
$CodexHomeForToml = Convert-ToTomlPath $TargetRoot

if (!(Test-Path -LiteralPath $SourceRoot)) {
    throw "Pasta fonte não encontrada: $SourceRoot"
}

New-Item -ItemType Directory -Force -Path $TargetRoot | Out-Null

$Stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$Installed = [System.Collections.Generic.List[string]]::new()
$Skipped = [System.Collections.Generic.List[string]]::new()
$Backups = [System.Collections.Generic.List[string]]::new()
$HasNvidiaGpu = Test-HasNvidiaGpu

$files = Get-ChildItem -LiteralPath $SourceRoot -Recurse -Force -File

foreach ($file in $files) {
    $relative = Get-RelativePathCompat -BasePath $SourceRoot -Path $file.FullName
    $relativeNormalized = $relative -replace "\\", "/"

    if (Should-SkipPath $relative) {
        $Skipped.Add($relative) | Out-Null
        continue
    }

    $targetRelative = $relative

    if ($targetRelative -eq "config.template.toml") {
        $targetRelative = "config.toml"
    } elseif ($targetRelative -eq "mcps.template.toml") {
        $targetRelative = "mcps.toml"
    } elseif ($targetRelative -like "*.template.toml") {
        $targetRelative = $targetRelative -replace "\.template\.toml$", ".toml"
    }

    $target = Join-Path $TargetRoot $targetRelative
    $targetDir = Split-Path -Parent $target

    New-Item -ItemType Directory -Force -Path $targetDir | Out-Null

    if (Test-Path -LiteralPath $target) {
        $backup = Get-BackupPath -Path $target -Stamp $Stamp
        Copy-Item -LiteralPath $target -Destination $backup -Force
        $Backups.Add($backup) | Out-Null
    }

    if ($relativeNormalized -in @("presets/mcp-services.json", "presets/skill-services.json")) {
        $content = Get-Content -LiteralPath $file.FullName -Raw

        if (-not $HasNvidiaGpu) {
            $content = $content.Replace("ollama-gpu", "ollama-cpu")
        }

        Set-Utf8NoBomContent -Path $target -Value $content
    } elseif ($file.Name -like "*.template.toml") {
        $content = Get-Content -LiteralPath $file.FullName -Raw
        $content = $content.Replace("{{CODEX_HOME}}", $CodexHomeForToml)
        Set-Utf8NoBomContent -Path $target -Value $content
    } else {
        Copy-Item -LiteralPath $file.FullName -Destination $target -Force
    }

    $Installed.Add($target) | Out-Null
}

if (Get-Command python -ErrorAction SilentlyContinue) {
    $configPath = Join-Path $TargetRoot "config.toml"

    if (Test-Path -LiteralPath $configPath) {
        python -c "import pathlib, tomllib; tomllib.loads(pathlib.Path(r'$configPath').read_text(encoding='utf-8')); print('config.toml ok')"
    }
}

if ($HasNvidiaGpu) {
    Write-Host "Modo: GPU"
} else {
    Write-Host "Modo: CPU"
}

Write-Host ""
Write-Host "Instalação concluída."
Write-Host "Fonte:   $SourceRoot"
Write-Host "Destino: $TargetRoot"
Write-Host ""
Write-Host "Arquivos instalados: $($Installed.Count)"
Write-Host "Arquivos ignorados:  $($Skipped.Count)"
Write-Host "Backups criados:     $($Backups.Count)"
