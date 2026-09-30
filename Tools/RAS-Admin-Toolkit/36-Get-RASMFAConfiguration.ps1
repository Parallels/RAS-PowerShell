[CmdletBinding()]
param([string]$Server, [switch]$UseEnvironment)

. "$PSScriptRoot\_Connect-RASFarm.ps1"
$null = Connect-RASFarm -Server $Server -UseEnvironment:$UseEnvironment
try {
    foreach ($site in Get-RASSite) {
        foreach ($mfa in Get-RASMFA -SiteId $site.Id) {
            $mfa | Select-Object @{Name = 'Site'; Expression = { $site.Name }},
                @{Name = 'Section'; Expression = { 'Provider' }}, *
            Get-RASMFACriteria -Id $mfa.Id | Select-Object -Property @(
                @{Name = 'Site'; Expression = { $site.Name }}
                @{Name = 'Section'; Expression = { 'Criteria' }}
                '*'
            )
        }
    }
} finally { Remove-RASSession }
