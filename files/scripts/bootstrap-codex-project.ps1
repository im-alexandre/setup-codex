param(
  [string]$ProjectPath = (Get-Location).Path
)

$ErrorActionPreference = "Stop"

$CodexHome = Join-Path $env:USERPROFILE ".codex"
$PresetDir = Join-Path $CodexHome "presets"
$McpPresetDir = Join-Path $PresetDir "mcp"
$SkillPresetDir = Join-Path $PresetDir "skills"
$BasePreset = Join-Path $PresetDir "base-project.toml"

$ProjectPath = (Resolve-Path -LiteralPath $ProjectPath).Path
$ProjectCodexDir = Join-Path $ProjectPath ".codex"
$ProjectConfig = Join-Path $ProjectCodexDir "config.toml"
$ProjectSkillsDir = Join-Path $ProjectCodexDir "skills"

function Set-Utf8NoBomContent {
  param(
    [Parameter(Mandatory=$true)][string]$Path,
    [Parameter(Mandatory=$true)][string]$Value
  )

  $encoding = [System.Text.UTF8Encoding]::new($false)
  [System.IO.File]::WriteAllText($Path, $Value, $encoding)
}

function Backup-IfExists {
  param(
    [Parameter(Mandatory=$true)][string]$Path
  )

  if (Test-Path -LiteralPath $Path) {
    $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $backupPath = "$Path.bak-$timestamp"
    Copy-Item -LiteralPath $Path -Destination $backupPath -Force
    return $backupPath
  }

  return $null
}

function Select-NumberedItems {
  param(
    [Parameter(Mandatory=$true)]$Items,
    [Parameter(Mandatory=$true)][string]$Title,
    [Parameter(Mandatory=$true)][string]$EmptyLabel,
    [scriptblock]$GetLabel
  )

  Write-Host ""
  Write-Host $Title
  Write-Host ""

  if ($Items.Count -eq 0) {
    Write-Host $EmptyLabel
    return @()
  }

  for ($i = 0; $i -lt $Items.Count; $i++) {
    $label = & $GetLabel $Items[$i]
    Write-Host ("[{0}] {1}" -f ($i + 1), $label)
  }

  Write-Host ""
  Write-Host "Digite os números separados por vírgula. Ex: 1,3,4"
  Write-Host "ENTER = nenhum"
  Write-Host ""

  $inputRaw = Read-Host "Selecionar"

  $selected = @()
  $selectedKeys = @()

  if (![string]::IsNullOrWhiteSpace($inputRaw)) {
    $indexes = $inputRaw -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }

    foreach ($idx in $indexes) {
      if ($idx -notmatch "^\d+$") {
        throw "Índice inválido: $idx"
      }

      $n = [int]$idx

      if ($n -lt 1 -or $n -gt $Items.Count) {
        throw "Índice fora do intervalo: $n"
      }

      $item = $Items[$n - 1]
      $key = (& $GetLabel $item)

      if ($selectedKeys -notcontains $key) {
        $selected += $item
        $selectedKeys += $key
      }
    }
  }

  return @($selected)
}

function Install-Skill {
  param(
    [Parameter(Mandatory=$true)]$Skill,
    [Parameter(Mandatory=$true)][string]$ProjectSkillsDir,
    [Parameter(Mandatory=$true)][string]$CodexHome
  )

  $name = [string]$Skill.name

  if ([string]::IsNullOrWhiteSpace($name)) {
    throw "Preset de skill sem campo 'name'."
  }

  $source = [string]$Skill.source

  if ([string]::IsNullOrWhiteSpace($source)) {
    $source = Join-Path $CodexHome "skills\$name"
  } else {
    $source = $source.Replace("{{CODEX_HOME}}", $CodexHome)
  }

  $target = Join-Path $ProjectSkillsDir $name

  if (!(Test-Path -LiteralPath $source)) {
    throw "Skill source não encontrado: $source"
  }

  New-Item -ItemType Directory -Force -Path $ProjectSkillsDir | Out-Null

  if (Test-Path -LiteralPath $target) {
    Remove-Item -Recurse -Force -LiteralPath $target
  }

  try {
    New-Item -ItemType Junction -Path $target -Target $source | Out-Null
  } catch {
    Copy-Item -Recurse -Force -LiteralPath $source -Destination $target
  }

  return $name
}

# Evita bootstrapar dentro do próprio runtime global.
$CodexHomeResolved = (Resolve-Path -LiteralPath $CodexHome).Path

try {
  [void]([System.IO.Path]::GetRelativePath($CodexHomeResolved, $ProjectPath))
  if ($ProjectPath.StartsWith($CodexHomeResolved, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "ProjectPath está dentro do runtime global do Codex. Abortando para evitar .codex dentro de ~/.codex: $ProjectPath"
  }
} catch {
  if ($_.Exception.Message -like "ProjectPath está dentro*") {
    throw
  }
}

if (!(Test-Path -LiteralPath $McpPresetDir)) {
  throw "Diretório de presets MCP não encontrado: $McpPresetDir"
}

$mcpPresets = @(Get-ChildItem -LiteralPath $McpPresetDir -Filter "*.toml" | Sort-Object Name)

$selectedMcps = Select-NumberedItems `
  -Items $mcpPresets `
  -Title "MCP presets disponíveis:" `
  -EmptyLabel "Nenhum preset MCP encontrado." `
  -GetLabel { param($x) $x.BaseName -replace "\.template$", "" }

$selectedMcpNames = @(
  $selectedMcps | ForEach-Object {
    $_.BaseName -replace "\.template$", ""
  }
)

$skillPresets = @()

if (Test-Path -LiteralPath $SkillPresetDir) {
  $skillPresets = @(Get-ChildItem -LiteralPath $SkillPresetDir -Filter "*.json" | Sort-Object Name)
}

$skillPresetObjects = @()

foreach ($preset in $skillPresets) {
  $raw = Get-Content -LiteralPath $preset.FullName -Raw
  $obj = $raw | ConvertFrom-Json
  $obj | Add-Member -NotePropertyName "_presetPath" -NotePropertyValue $preset.FullName -Force
  $skillPresetObjects += $obj
}

$selectedSkills = Select-NumberedItems `
  -Items $skillPresetObjects `
  -Title "Skill presets disponíveis:" `
  -EmptyLabel "Nenhum preset de skill encontrado." `
  -GetLabel { param($x) "$($x.name) - $($x.description)" }

$selectedSkillNames = @(
  $selectedSkills | ForEach-Object {
    $_.name
  }
)

New-Item -ItemType Directory -Force -Path $ProjectCodexDir | Out-Null

$backup = Backup-IfExists -Path $ProjectConfig

if ($backup) {
  Write-Host ""
  Write-Host "Backup criado: $backup"
}

$parts = New-Object System.Collections.Generic.List[string]

$parts.Add("# Generated by bootstrap-codex-project.ps1")
$parts.Add("# Project: $ProjectPath")
$parts.Add("")

if (Test-Path -LiteralPath $BasePreset) {
  $parts.Add((Get-Content -LiteralPath $BasePreset -Raw).Trim())
  $parts.Add("")
}

foreach ($preset in $selectedMcps) {
  $parts.Add("")
  $parts.Add("# --- MCP preset: $($preset.BaseName) ---")
  $parts.Add((Get-Content -LiteralPath $preset.FullName -Raw).Trim())
}

$final = ($parts -join "`r`n").Trim() + "`r`n"

Set-Utf8NoBomContent -Path $ProjectConfig -Value $final

Write-Host ""
Write-Host "Config local criada:"
Write-Host $ProjectConfig

if ($selectedMcps.Count -gt 0) {
  Write-Host ""
  Write-Host "MCPs habilitados:"
  $selectedMcpNames | ForEach-Object { Write-Host "- $_" }
}

$installedSkills = @()

foreach ($skill in $selectedSkills) {
  $installedSkills += Install-Skill `
    -Skill $skill `
    -ProjectSkillsDir $ProjectSkillsDir `
    -CodexHome $CodexHome
}

if ($installedSkills.Count -gt 0) {
  Write-Host ""
  Write-Host "Skills habilitadas:"
  $installedSkills | ForEach-Object { Write-Host "- $_" }
}

$resolveScript = Join-Path $CodexHome "scripts\resolve-codex-services.ps1"
$startServicesScript = Join-Path $CodexHome "scripts\start-codex-services.ps1"

if ((Test-Path -LiteralPath $resolveScript) -and (Test-Path -LiteralPath $startServicesScript)) {
  $rawResolved = & powershell -NoProfile -ExecutionPolicy Bypass `
    -File $resolveScript `
    -Mcps $selectedMcpNames `
    -Skills $selectedSkillNames

  $resolved = $null

  if ($rawResolved) {
    $resolved = ($rawResolved | Out-String).Trim() | ConvertFrom-Json
  }

  $servicesToStart = @()

  if ($resolved -and $resolved.services) {
    $servicesToStart = @($resolved.services) |
      Where-Object { ![string]::IsNullOrWhiteSpace($_) } |
      Sort-Object -Unique
  }

  $profilesToUse = @()

  if ($resolved -and $resolved.profiles) {
    $profilesToUse = @($resolved.profiles) |
      Where-Object { ![string]::IsNullOrWhiteSpace($_) } |
      Sort-Object -Unique
  }

  if ($servicesToStart.Count -gt 0) {
    Write-Host ""
    Write-Host "Serviços Docker necessários:"
    $servicesToStart | ForEach-Object { Write-Host "- $_" }

    if ($profilesToUse.Count -gt 0) {
      Write-Host ""
      Write-Host "Profiles Docker necessários:"
      $profilesToUse | ForEach-Object { Write-Host "- $_" }
    }

    $answer = Read-Host "Deseja iniciar agora? [S/n]"

    if ($answer -eq "" -or $answer.ToLower().StartsWith("s")) {
      powershell -NoProfile -ExecutionPolicy Bypass `
        -File $startServicesScript `
        -Services $servicesToStart `
        -Profiles $profilesToUse
    }
  }
}
