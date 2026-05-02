[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string[]]$McpNames = @(),
    [string[]]$SkillNames = @(),
    [string[]]$Services = @(),
    [string]$CodexHome = (Join-Path $env:USERPROFILE ".codex")
)

$ErrorActionPreference = "Stop"

$resolvedServices = [System.Collections.Generic.List[string]]::new()

foreach ($service in $Services) {
    if (![string]::IsNullOrWhiteSpace($service) -and -not $resolvedServices.Contains($service)) {
        $resolvedServices.Add($service)
    }
}

if (($McpNames.Count -gt 0) -or ($SkillNames.Count -gt 0)) {
    $resolveScript = Join-Path $PSScriptRoot "resolve-codex-services.ps1"

    if (!(Test-Path -LiteralPath $resolveScript)) {
        throw "Script de resolução de serviços não encontrado: $resolveScript"
    }

    $servicesFromMaps = @(& $resolveScript -McpNames $McpNames -SkillNames $SkillNames -CodexHome $CodexHome)

    foreach ($service in $servicesFromMaps) {
        if (![string]::IsNullOrWhiteSpace($service) -and -not $resolvedServices.Contains($service)) {
            $resolvedServices.Add($service)
        }
    }
}

$uniqueServices = [System.Collections.Generic.List[string]]::new()

foreach ($service in $resolvedServices) {
    if (![string]::IsNullOrWhiteSpace($service) -and -not $uniqueServices.Contains($service)) {
        $uniqueServices.Add($service)
    }
}

if ($uniqueServices.Count -eq 0) {
    Write-Host "Nenhum serviço para iniciar."
    return
}

$composeFile = Join-Path $CodexHome "docker-compose.yml"

if (!(Test-Path -LiteralPath $composeFile)) {
    throw "Arquivo docker-compose não encontrado: $composeFile"
}

if ($PSCmdlet.ShouldProcess($composeFile, ("docker compose up -d {0}" -f ($uniqueServices -join ", ")))) {
    $docker = Get-Command docker -ErrorAction Stop
    $arguments = @("compose", "-f", $composeFile, "up", "-d") + $uniqueServices.ToArray()
    & $docker.Path @arguments
}
