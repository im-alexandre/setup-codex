param(
  [Parameter(Mandatory=$true)]
  [string[]]$Services,

  [string[]]$Profiles = @(),

  [string]$ComposeFile = (Join-Path $env:USERPROFILE ".codex\docker-compose.yml")
)

$ErrorActionPreference = "Stop"

if (!(Get-Command docker -ErrorAction SilentlyContinue)) {
  throw "Docker não encontrado no PATH."
}

if (!(Test-Path -LiteralPath $ComposeFile)) {
  throw "docker-compose.yml não encontrado: $ComposeFile"
}

$uniqueServices = $Services |
  Where-Object { ![string]::IsNullOrWhiteSpace($_) } |
  Sort-Object -Unique

$uniqueProfiles = $Profiles |
  Where-Object { ![string]::IsNullOrWhiteSpace($_) } |
  Sort-Object -Unique

if ($uniqueServices.Count -eq 0) {
  Write-Host "Nenhum serviço necessário."
  exit 0
}

$args = @("compose", "-f", $ComposeFile)

foreach ($profile in $uniqueProfiles) {
  $args += @("--profile", $profile)
}

$args += @("up", "-d")
$args += $uniqueServices

Write-Host ""
Write-Host "Iniciando serviços Codex:"
$uniqueServices | ForEach-Object { Write-Host "- $_" }

if ($uniqueProfiles.Count -gt 0) {
  Write-Host ""
  Write-Host "Profiles:"
  $uniqueProfiles | ForEach-Object { Write-Host "- $_" }
}

Write-Host ""

& docker @args
