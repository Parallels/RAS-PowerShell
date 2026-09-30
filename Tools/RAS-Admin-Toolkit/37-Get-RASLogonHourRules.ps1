[CmdletBinding()]
param([string]$Server, [switch]$UseEnvironment)

. "$PSScriptRoot\_Connect-RASFarm.ps1"
$null = Connect-RASFarm -Server $Server -UseEnvironment:$UseEnvironment
try {
    Get-RASSite | ForEach-Object { Get-RASLogonHours -SiteId $_.Id } |
        Sort-Object SiteId, Priority, Name
} finally { Remove-RASSession }
