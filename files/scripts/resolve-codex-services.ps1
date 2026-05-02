[CmdletBinding()]
param(
    [string[]]$McpNames = @(),
    [string[]]$SkillNames = @(),
    [string]$CodexHome = (Join-Path $env:USERPROFILE ".codex")
)

$ErrorActionPreference = "Stop"

function Read-ServiceMap {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (!(Test-Path -LiteralPath $Path)) {
        return @{}
    }

    $raw = Get-Content -LiteralPath $Path -Raw

    if ([string]::IsNullOrWhiteSpace($raw)) {
        return @{}
    }

    return ($raw | ConvertFrom-Json -AsHashtable)
}

function Add-ResolvedServices {
    param(
        [Parameter(Mandatory = $true)]$Entry,
        [System.Collections.Generic.List[string]]$Services
    )

    $candidateServices = @()

    if ($Entry -is [System.Collections.IDictionary]) {
        $candidateServices = @($Entry['services'])
    } elseif ($Entry.PSObject.Properties.Name -contains 'services') {
        $candidateServices = @($Entry.services)
    }

    foreach ($service in $candidateServices) {
        if (![string]::IsNullOrWhiteSpace([string]$service) -and -not $Services.Contains([string]$service)) {
            $Services.Add([string]$service)
        }
    }
}

$mcpServicesPath = Join-Path $CodexHome "presets\mcp-services.json"
$skillServicesPath = Join-Path $CodexHome "presets\skill-services.json"

$mcpRegistry = Read-ServiceMap -Path $mcpServicesPath
$skillRegistry = Read-ServiceMap -Path $skillServicesPath

$resolved = [System.Collections.Generic.List[string]]::new()
$missing = [System.Collections.Generic.List[string]]::new()

foreach ($name in $McpNames) {
    if ([string]::IsNullOrWhiteSpace($name)) {
        continue
    }

    $entry = $mcpRegistry[$name]

    if ($null -ne $entry) {
        Add-ResolvedServices -Entry $entry -Services $resolved
    } else {
        $missing.Add("mcp:$name")
    }
}

foreach ($name in $SkillNames) {
    if ([string]::IsNullOrWhiteSpace($name)) {
        continue
    }

    $entry = $skillRegistry[$name]

    if ($null -ne $entry) {
        Add-ResolvedServices -Entry $entry -Services $resolved
    } else {
        $missing.Add("skill:$name")
    }
}

if ($missing.Count -gt 0) {
    Write-Warning ("Serviços não encontrados para: {0}" -f ($missing -join ", "))
}

$resolved
