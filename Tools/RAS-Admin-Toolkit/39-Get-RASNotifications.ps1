[CmdletBinding()]
param([string]$Server, [switch]$UseEnvironment)

. "$PSScriptRoot\_Connect-RASFarm.ps1"
$null = Connect-RASFarm -Server $Server -UseEnvironment:$UseEnvironment
try {
    foreach ($site in Get-RASSite) {
        Get-RASNotification -SiteId $site.Id | ForEach-Object {
            $_ | Select-Object @{Name = 'Site'; Expression = { $site.Name }}, *
        }
    }
} finally { Remove-RASSession }
