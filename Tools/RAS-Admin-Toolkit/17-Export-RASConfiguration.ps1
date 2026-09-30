[CmdletBinding()]
param(
    [string]$Server,
    [switch]$UseEnvironment,
    [string]$Path = ".\RAS-Backup-$(Get-Date -Format 'yyyyMMdd-HHmmss').dat2",
    [uint32]$TimeoutInSecs = 600
)

. "$PSScriptRoot\_Connect-RASFarm.ps1"
$null = Connect-RASFarm -Server $Server -UseEnvironment:$UseEnvironment
try {
    $fullPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
    $parent = Split-Path $fullPath -Parent
    if (-not (Test-Path -LiteralPath $parent -PathType Container)) { throw "Folder does not exist: $parent" }
    Invoke-RASExportSettings -FilePath $fullPath -TimeoutInSecs $TimeoutInSecs
} finally { Remove-RASSession }
