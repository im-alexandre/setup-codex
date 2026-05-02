[CmdletBinding()]
param(
  [string]$SourceSkillPath = 'D:\codex-setup\files\skills\docx-utils',
  [string]$TargetSkillPath = (Join-Path $env:USERPROFILE '.codex\skills\docx-utils'),
  [switch]$SkipTests,
  [switch]$SkipSkillValidation,
  [switch]$Clean
)

$ErrorActionPreference = 'Stop'

function Write-Status {
  param([string]$Message)
  Write-Host "[docx-utils-global] $Message"
}

if (-not (Test-Path -LiteralPath $SourceSkillPath)) {
  throw "Skill de origem não encontrada: $SourceSkillPath"
}

New-Item -ItemType Directory -Force -Path (Split-Path -Parent $TargetSkillPath) | Out-Null

if ($Clean -and (Test-Path -LiteralPath $TargetSkillPath)) {
  Remove-Item -LiteralPath $TargetSkillPath -Recurse -Force
}

if (Test-Path -LiteralPath $TargetSkillPath) {
  Get-ChildItem -LiteralPath $SourceSkillPath -Force | ForEach-Object {
    Copy-Item -LiteralPath $_.FullName -Destination $TargetSkillPath -Recurse -Force
  }
} else {
  Copy-Item -LiteralPath $SourceSkillPath -Destination $TargetSkillPath -Recurse -Force
}
Write-Status "Skill copiada para $TargetSkillPath"

$installer = Join-Path $TargetSkillPath 'scripts\install-docx-utils.ps1'
if (-not (Test-Path -LiteralPath $installer)) {
  throw "Instalador não encontrado no destino: $installer"
}

$installerArgs = @{}
if ($SkipTests) {
  $installerArgs['SkipTests'] = $true
}
if ($SkipSkillValidation) {
  $installerArgs['SkipSkillValidation'] = $true
}

& $installer @installerArgs
if ($LASTEXITCODE -ne 0) {
  throw 'A instalação da skill docx-utils no destino falhou.'
}

Write-Status 'Instalação global concluída com sucesso.'
