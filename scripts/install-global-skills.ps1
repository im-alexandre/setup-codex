param(
  [string[]]$SkillNames = @(),
  [ValidateSet("curated", "experimental", "system")]
  [string]$Channel = "curated"
)

$ErrorActionPreference = "Stop"

$Root = Join-Path $env:USERPROFILE ".codex"
$Imports = Join-Path $Root "vendor\imports"

$SourceRoot = switch ($Channel) {
  "curated" {
    Join-Path $Imports "skills\.curated" 
  }
  "experimental" {
    Join-Path $Imports "skills\.experimental" 
  }
  "system" {
    Join-Path $Imports "skills\.system" 
  }
}

$TargetRoot = Join-Path $Root "skills"

if (!(Test-Path $SourceRoot)) {
  throw "Diretório de skills não encontrado: $SourceRoot"
}

New-Item -ItemType Directory -Force -Path $TargetRoot | Out-Null

if ($SkillNames.Count -eq 0) {
  Write-Host "Skills disponíveis em ${Channel}:"
  Get-ChildItem $SourceRoot -Directory | Sort-Object Name | ForEach-Object {
    Write-Host "- $($_.Name)"
  }
  exit 0
}

foreach ($name in $SkillNames) {
  $src = Join-Path $SourceRoot $name
  $dst = Join-Path $TargetRoot $name

  if (!(Test-Path (Join-Path $src "SKILL.md"))) {
    throw "Skill inválida ou não encontrada: $src"
  }

  if (Test-Path $dst) {
    Remove-Item -Recurse -Force $dst
  }

  Copy-Item -Recurse -Force $src $dst
  Write-Host "Skill instalada: $name"
}
