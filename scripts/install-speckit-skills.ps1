$ErrorActionPreference = "Stop"

$Root = Join-Path $env:USERPROFILE ".codex"
$SpecKit = Join-Path $Root "mcp\spec-kit"
$Target = Join-Path $Root "skills\speckit"

if (!(Test-Path $SpecKit)) {
  throw "Spec Kit não encontrado em: $SpecKit"
}

New-Item -ItemType Directory -Force -Path $Target | Out-Null

Get-ChildItem -Path $SpecKit -Recurse -Filter "SKILL.md" | ForEach-Object {
  $skillDir = $_.Directory.FullName
  $skillName = Split-Path $skillDir -Leaf
  $dst = Join-Path $Target $skillName

  if (Test-Path $dst) {
    Remove-Item -Recurse -Force $dst
  }

  Copy-Item -Recurse -Force $skillDir $dst
  Write-Host "Skill instalada: $skillName"
}
