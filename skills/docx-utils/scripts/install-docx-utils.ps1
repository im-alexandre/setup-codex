[CmdletBinding()]
param(
  [switch]$SkipTests,
  [switch]$SkipSkillValidation,
  [ValidateSet('Debug', 'Release')]
  [string]$Configuration = 'Debug',
  [switch]$NoPackageMutation
)

$ErrorActionPreference = 'Stop'

function Write-Status {
  param([string]$Message)
  Write-Host "[docx-utils] $Message"
}

function Resolve-DotNet {
  $dotnet = Get-Command dotnet -ErrorAction SilentlyContinue
  if (-not $dotnet) {
    throw "SDK do .NET não encontrado no PATH."
  }
  return $dotnet.Source
}

function Get-XmlDocument {
  param([string]$Path)
  $settings = [System.Xml.XmlReaderSettings]::new()
  $settings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
  $reader = [System.Xml.XmlReader]::Create($Path, $settings)
  try {
    $doc = [System.Xml.Linq.XDocument]::Load($reader, [System.Xml.Linq.LoadOptions]::PreserveWhitespace)
  } finally {
    $reader.Dispose()
  }
  return $doc
}

function Save-XmlDocument {
  param(
    [System.Xml.Linq.XDocument]$Document,
    [string]$Path
  )
  $writerSettings = [System.Xml.XmlWriterSettings]::new()
  $writerSettings.Indent = $true
  $writerSettings.OmitXmlDeclaration = $false
  $writerSettings.Encoding = [System.Text.UTF8Encoding]::new($false)
  $writer = [System.Xml.XmlWriter]::Create($Path, $writerSettings)
  try {
    $Document.Save($writer)
  } finally {
    $writer.Dispose()
  }
}

function Invoke-DotNet {
  param(
    [string[]]$Arguments,
    [string]$WorkingDirectory
  )
  $output = & dotnet @Arguments 2>&1
  if ($LASTEXITCODE -ne 0) {
    throw ($output -join [Environment]::NewLine)
  }
  if ($output) {
    $output | ForEach-Object { Write-Host $_ }
  }
}

function Get-PackageReferences {
  param([string]$ProjectPath)
  $doc = Get-XmlDocument -Path $ProjectPath
  $ns = $doc.Root.Name.Namespace
  $refs = @{}
  foreach ($itemGroup in $doc.Root.Elements($ns + 'ItemGroup')) {
    foreach ($package in $itemGroup.Elements($ns + 'PackageReference')) {
      $include = $package.Attribute('Include')
      if ($include) {
        $version = $package.Attribute('Version')
        $refs[$include.Value] = if ($version) { $version.Value } else { '' }
      }
    }
  }
  return $refs
}

function Get-HighestKnownVersion {
  param(
    [string]$PackageId,
    [string[]]$ProjectFiles
  )
  $versions = New-Object System.Collections.Generic.List[string]
  foreach ($project in $ProjectFiles) {
    $refs = Get-PackageReferences -ProjectPath $project
    if ($refs.ContainsKey($PackageId) -and $refs[$PackageId]) {
      $versions.Add($refs[$PackageId])
    }
  }
  if ($versions.Count -eq 0) {
    return $null
  }
  $parsed = @($versions | Sort-Object {
    try {
      [version]$_
    } catch {
      [version]'0.0'
    }
  } -Descending)
  return $parsed[0]
}

function Ensure-PackageReference {
  param(
    [string]$ProjectPath,
    [string]$PackageId,
    [string]$VersionHint,
    [string[]]$AllProjectFiles
  )

  $doc = Get-XmlDocument -Path $ProjectPath
  $ns = $doc.Root.Name.Namespace
  $references = Get-PackageReferences -ProjectPath $ProjectPath

  if ($references.ContainsKey($PackageId)) {
    return $false
  }

  if ($NoPackageMutation) {
    throw "Pacote ausente em '$ProjectPath': $PackageId"
  }

  $version = $VersionHint
  if (-not $version) {
    $version = Get-HighestKnownVersion -PackageId $PackageId -ProjectFiles $AllProjectFiles
  }
  if (-not $version) {
    Write-Status "Sem versão local conhecida para '$PackageId'; dotnet escolherá a versão estável disponível."
  }

  $arguments = @('add', $ProjectPath, 'package', $PackageId)
  if ($version) {
    $arguments += '--version'
    $arguments += $version
  }
  Invoke-DotNet -Arguments $arguments -WorkingDirectory (Split-Path -Parent $ProjectPath)
  return $true
}

function Find-TestProjects {
  param([string[]]$ProjectFiles)
  return $ProjectFiles | Where-Object {
    $name = [System.IO.Path]::GetFileNameWithoutExtension($_)
    $name -match '(^|\.)(Tests?|Spec)$' -or $name -match 'Tests?$' -or $name -match 'Spec$'
  }
}

$SkillRoot = Split-Path -Parent $PSScriptRoot
$dotnetPath = Resolve-DotNet
Write-Status "dotnet encontrado em $dotnetPath"

$projectFiles = Get-ChildItem -LiteralPath $SkillRoot -Recurse -Filter '*.csproj' |
  Sort-Object FullName |
  Select-Object -ExpandProperty FullName

if (-not $projectFiles) {
  throw "Nenhum projeto .NET encontrado em: $SkillRoot"
}

$defaultPackageVersions = @{
  'DocumentFormat.OpenXml' = '3.2.0'
  'SixLabors.ImageSharp' = '3.1.12'
  'System.IO.Packaging' = '8.0.1'
}

$packageHints = @{
  'DocumentFormat.OpenXml' = Get-HighestKnownVersion -PackageId 'DocumentFormat.OpenXml' -ProjectFiles $projectFiles
  'SixLabors.ImageSharp' = Get-HighestKnownVersion -PackageId 'SixLabors.ImageSharp' -ProjectFiles $projectFiles
  'System.IO.Packaging' = Get-HighestKnownVersion -PackageId 'System.IO.Packaging' -ProjectFiles $projectFiles
}

foreach ($packageId in @($defaultPackageVersions.Keys)) {
  if (-not $packageHints[$packageId]) {
    $packageHints[$packageId] = $defaultPackageVersions[$packageId]
  }
}

$packageChanges = 0
foreach ($project in $projectFiles) {
  $name = [System.IO.Path]::GetFileName($project)
  $requiredPackages = switch ($name) {
    'ArticleDocxBuilder.csproj' { @('DocumentFormat.OpenXml', 'SixLabors.ImageSharp', 'System.IO.Packaging') }
    'DocxOpenXmlTools.csproj' { @('DocumentFormat.OpenXml', 'SixLabors.ImageSharp') }
    'StyleXmlExporter.csproj' { @('DocumentFormat.OpenXml') }
    default { @() }
  }

  foreach ($packageId in $requiredPackages) {
    $versionHint = if ($packageHints.ContainsKey($packageId)) { [string]($packageHints[$packageId]) } else { $null }
    $changed = Ensure-PackageReference -ProjectPath $project -PackageId $packageId -VersionHint $versionHint -AllProjectFiles $projectFiles
    if ($changed) {
      $packageChanges++
      Write-Status "Pacote adicionado em ${name}: $packageId"
    }
  }
}

foreach ($project in $projectFiles) {
  Write-Status "restore: $([System.IO.Path]::GetFileName($project))"
  Invoke-DotNet -Arguments @('restore', $project) -WorkingDirectory $SkillRoot
}

foreach ($project in $projectFiles) {
  Write-Status "build: $([System.IO.Path]::GetFileName($project)) [$Configuration]"
  Invoke-DotNet -Arguments @('build', $project, '--configuration', $Configuration, '--no-restore') -WorkingDirectory $SkillRoot
}

$testProjectsRun = 0
$testStatus = 'executados'
if (-not $SkipTests) {
  $testProjects = Find-TestProjects -ProjectFiles $projectFiles
  if ($testProjects) {
    foreach ($project in $testProjects) {
      Write-Status "test: $([System.IO.Path]::GetFileName($project)) [$Configuration]"
      Invoke-DotNet -Arguments @('test', $project, '--configuration', $Configuration, '--no-build', '--no-restore') -WorkingDirectory $SkillRoot
      $testProjectsRun++
    }
  } else {
    Write-Status 'Nenhum projeto de teste encontrado; pulando dotnet test.'
    $testStatus = 'nenhum-projeto'
  }
} else {
  Write-Status 'Testes automáticos ignorados por parâmetro.'
  $testStatus = 'ignorado'
}

$skillValidationStatus = 'executada'
if (-not $SkipSkillValidation) {
  $skillCreator = Join-Path $env:USERPROFILE '.codex\skills\.system\skill-creator\scripts\quick_validate.py'
  if (Test-Path $skillCreator) {
    Write-Status 'Executando quick_validate.py da skill skill-creator.'
    $python = Get-Command python -ErrorAction SilentlyContinue
    if (-not $python) {
      $python = Get-Command py -ErrorAction SilentlyContinue
    }
    if (-not $python) {
      throw 'Python não encontrado para executar quick_validate.py.'
    }

    $pythonArgs = @()
    if ($python.Name -eq 'py') {
      $pythonArgs += '-3'
    }
    $pythonArgs += $skillCreator
    $pythonArgs += $SkillRoot
    & $python.Source @pythonArgs
    if ($LASTEXITCODE -ne 0) {
      throw 'Validação rápida da skill falhou.'
    }
  } else {
    Write-Status 'skill-creator não encontrado; quick_validate.py ignorado.'
    $skillValidationStatus = 'skill-creator-ausente'
  }
} else {
  Write-Status 'Validação da skill ignorada por parâmetro.'
  $skillValidationStatus = 'ignorada'
}

Write-Status ("Resumo: projetos={0}; alterações-de-pacote={1}; testes={2}({3}); skill-validation={4}; configuração={5}" -f $projectFiles.Count, $packageChanges, $testStatus, $testProjectsRun, $skillValidationStatus, $Configuration)
