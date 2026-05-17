[CmdletBinding()]
param(
  [switch]$SkipTests,
  [switch]$SkipSkillValidation,
  [ValidateSet('Debug', 'Release')]
  [string]$Configuration = 'Release'
)

$ErrorActionPreference = 'Stop'

function Write-Status {
  param([string]$Message)
  Write-Host "[pptx-utils] $Message"
}

function Invoke-DotNet {
  param(
    [string[]]$Arguments,
    [string]$WorkingDirectory
  )

  Push-Location $WorkingDirectory
  try {
    & dotnet @Arguments
  } finally {
    Pop-Location
  }
  if ($LASTEXITCODE -ne 0) {
    throw "dotnet $($Arguments -join ' ') falhou com exit code $LASTEXITCODE."
  }
}

$dotnet = Get-Command dotnet -ErrorAction SilentlyContinue
if (-not $dotnet) {
  throw "SDK do .NET nao encontrado no PATH."
}

$SkillRoot = Split-Path -Parent $PSScriptRoot
$ToolProject = Join-Path $SkillRoot 'src\PptxOpenXmlTools\PptxOpenXmlTools.csproj'
$TestProject = Join-Path $SkillRoot 'src\PptxOpenXmlTools.Tests\PptxOpenXmlTools.Tests.csproj'
$PackageDir = Join-Path $SkillRoot 'bin\pptx-utils'

Write-Status "dotnet encontrado em $($dotnet.Source)"

if (-not (Test-Path $ToolProject)) {
  throw "Projeto CLI nao encontrado: $ToolProject"
}

Write-Status "restore: PptxOpenXmlTools"
Invoke-DotNet -Arguments @('restore', $ToolProject) -WorkingDirectory $SkillRoot

Write-Status "build: PptxOpenXmlTools [$Configuration]"
Invoke-DotNet -Arguments @('build', $ToolProject, '--configuration', $Configuration, '--no-restore') -WorkingDirectory $SkillRoot

if (-not $SkipTests) {
  if (-not (Test-Path $TestProject)) {
    throw "Projeto de testes nao encontrado: $TestProject"
  }

  Write-Status "test: PptxOpenXmlTools.Tests [$Configuration]"
  Invoke-DotNet -Arguments @('test', $TestProject, '--configuration', $Configuration, '--no-restore') -WorkingDirectory $SkillRoot
} else {
  Write-Status 'Testes automaticos ignorados por parametro.'
}

Write-Status "publish: PptxOpenXmlTools [$Configuration] -> $PackageDir"
Invoke-DotNet -Arguments @('publish', $ToolProject, '--configuration', $Configuration, '--no-restore', '--output', $PackageDir) -WorkingDirectory $SkillRoot

$Executable = Join-Path $PackageDir 'pptx-utils.exe'
if (-not (Test-Path $Executable)) {
  throw "Binario publicado nao encontrado: $Executable"
}

Write-Status 'smoke: pptx-utils --help'
& $Executable --help | Write-Host
if ($LASTEXITCODE -ne 0) {
  throw 'Smoke test do binario publicado falhou.'
}

if (-not $SkipSkillValidation) {
  $SkillCreator = Join-Path $env:USERPROFILE '.codex\skills\.system\skill-creator\scripts\quick_validate.py'
  if (Test-Path $SkillCreator) {
    $python = Get-Command python -ErrorAction SilentlyContinue
    if (-not $python) {
      $python = Get-Command py -ErrorAction SilentlyContinue
    }
    if (-not $python) {
      throw 'Python nao encontrado para executar quick_validate.py.'
    }

    Write-Status 'skill validation: quick_validate.py'
    if ($python.Name -eq 'py') {
      & $python.Source -3 $SkillCreator $SkillRoot
    } else {
      & $python.Source $SkillCreator $SkillRoot
    }
    if ($LASTEXITCODE -ne 0) {
      throw 'Validacao rapida da skill falhou.'
    }
  } else {
    Write-Status 'skill-creator nao encontrado; quick_validate.py ignorado.'
  }
} else {
  Write-Status 'Validacao da skill ignorada por parametro.'
}

Write-Status "Instalacao concluida: $Executable"
