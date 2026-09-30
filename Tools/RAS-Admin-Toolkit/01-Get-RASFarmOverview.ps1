[CmdletBinding()]
param([string]$Server, [switch]$UseEnvironment)

. "$PSScriptRoot\_Connect-RASFarm.ps1"
$null = Connect-RASFarm -Server $Server -UseEnvironment:$UseEnvironment
try {
    Get-RASVersion
    Get-RASSite
    Get-RASFarmSettings
} finally { Remove-RASSession }
