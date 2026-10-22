[CmdletBinding()]
param([string]$Server, [switch]$UseEnvironment)

. "$PSScriptRoot\_Connect-RASFarm.ps1"
$null = Connect-RASFarm -Server $Server -UseEnvironment:$UseEnvironment
try {
    Get-RASPubItem | Where-Object {
        $_.EnabledMode -in 'Disabled', 'Maintenance' -or
        ($null -ne $_.Enabled -and -not $_.Enabled)
    } | Sort-Object SiteId, Name
} finally { Remove-RASSession }
