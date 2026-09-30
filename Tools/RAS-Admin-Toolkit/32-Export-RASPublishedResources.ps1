[CmdletBinding()]
param(
    [string]$Server,
    [switch]$UseEnvironment,
    [string]$Path = ".\RAS-PublishedResources-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
)

. "$PSScriptRoot\_Connect-RASFarm.ps1"
$null = Connect-RASFarm -Server $Server -UseEnvironment:$UseEnvironment
try {
    $fullPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
    $parent = Split-Path $fullPath -Parent
    if (-not (Test-Path -LiteralPath $parent -PathType Container)) { throw "Folder does not exist: $parent" }
    Get-RASPubItem | Export-Csv -LiteralPath $fullPath -NoTypeInformation -Encoding utf8
    Get-Item -LiteralPath $fullPath
} finally { Remove-RASSession }
