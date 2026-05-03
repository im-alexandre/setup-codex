[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$CodexHome = (Join-Path $env:USERPROFILE ".codex")
)

$ErrorActionPreference = "Stop"

function Get-MendeleyCommand {
    $command = Get-Command mendeley-mcp -ErrorAction SilentlyContinue

    if ($command) {
        return $command.Path
    }

    throw "Comando mendeley-mcp não encontrado no PATH."
}

if ($PSCmdlet.ShouldProcess("mendeley-mcp", "start mendeley MCP")) {
    $commandPath = Get-MendeleyCommand
    Push-Location $CodexHome
    try {
        & $commandPath
    } finally {
        Pop-Location
    }
}
