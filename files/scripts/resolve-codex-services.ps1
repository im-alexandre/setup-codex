param(
  [string[]]$Mcps = @(),
  [string[]]$Skills = @()
)

$ErrorActionPreference = "Stop"

$Root = Join-Path $env:USERPROFILE ".codex"
$McpServicesPath = Join-Path $Root "presets\mcp-services.json"
$SkillServicesPath = Join-Path $Root "presets\skills-services.json"

$services = @()
$profiles = @()

function Add-Entry {
  param($Entry)

  foreach ($svc in @($Entry.services)) {
    if (![string]::IsNullOrWhiteSpace($svc)) {
      $script:services += $svc
    }
  }

  foreach ($profile in @($Entry.profiles)) {
    if (![string]::IsNullOrWhiteSpace($profile)) {
      $script:profiles += $profile
    }
  }
}

if ((Test-Path -LiteralPath $McpServicesPath) -and $Mcps.Count -gt 0) {
  $mcpServices = Get-Content -LiteralPath $McpServicesPath -Raw | ConvertFrom-Json

  foreach ($mcp in $Mcps) {
    if ($mcpServices.PSObject.Properties.Name -contains $mcp) {
      Add-Entry $mcpServices.$mcp
    }
  }
}

if ((Test-Path -LiteralPath $SkillServicesPath) -and $Skills.Count -gt 0) {
  $skillServices = Get-Content -LiteralPath $SkillServicesPath -Raw | ConvertFrom-Json

  foreach ($skill in $Skills) {
    if ($skillServices.PSObject.Properties.Name -contains $skill) {
      Add-Entry $skillServices.$skill
    }
  }
}

[pscustomobject]@{
  services = @($services | Sort-Object -Unique)
  profiles = @($profiles | Sort-Object -Unique)
} | ConvertTo-Json -Depth 10
