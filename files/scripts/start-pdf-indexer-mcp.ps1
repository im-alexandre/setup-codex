[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$CodexHome = (Join-Path $env:USERPROFILE ".codex")
)

$ErrorActionPreference = "Stop"

function Get-PythonCommand {
    $python = Get-Command python -ErrorAction SilentlyContinue

    if ($python) {
        return $python.Path
    }

    $py = Get-Command py -ErrorAction SilentlyContinue

    if ($py) {
        return $py.Path
    }

    throw "Python não encontrado no PATH."
}

$pdfIndexerRoot = "D:\mcp\pdf-indexer-mcp"
$scriptPath = Join-Path $pdfIndexerRoot "semantic_chunked_pdf_rag.py"

if (!(Test-Path -LiteralPath $scriptPath)) {
    throw "MCP do pdf-indexer não encontrado: $scriptPath"
}

$resolveScript = Join-Path $PSScriptRoot "resolve-codex-services.ps1"
$startServicesScript = Join-Path $PSScriptRoot "start-codex-services.ps1"

if ((Test-Path $resolveScript) -and (Test-Path $startServicesScript)) {
    & $startServicesScript -McpNames @("pdf-indexer") -SkillNames @("pdf-indexer") -CodexHome $CodexHome
}

if ($PSCmdlet.ShouldProcess($scriptPath, "start pdf-indexer MCP")) {
    $pythonPath = Get-PythonCommand
    Push-Location $pdfIndexerRoot
    try {
        & $pythonPath $scriptPath
    } finally {
        Pop-Location
    }
}
