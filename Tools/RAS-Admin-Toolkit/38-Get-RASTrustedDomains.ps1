[CmdletBinding()]
param([string]$Server, [switch]$UseEnvironment)

. "$PSScriptRoot\_Connect-RASFarm.ps1"
$null = Connect-RASFarm -Server $Server -UseEnvironment:$UseEnvironment
try {
    foreach ($site in Get-RASSite) {
        $adSettings = Get-RASADIntegrationSettings -SiteId $site.Id
        $adSettings | Select-Object -Property @(
            @{Name = 'Site'; Expression = { $site.Name }}
            @{Name = 'Section'; Expression = { 'ADIntegration' }}
            '*'
        ) -ExcludeProperty CASettings
        if ($adSettings.CASettings) {
            $adSettings.CASettings | Select-Object -Property @(
                @{Name = 'Site'; Expression = { $site.Name }}
                @{Name = 'Section'; Expression = { 'CertificateAuthoritySettings' }}
                '*'
            )
        }
        Get-RASTrustedDomain -SiteId $site.Id | ForEach-Object {
            $_ | Select-Object @{Name = 'Site'; Expression = { $site.Name }},
                @{Name = 'Section'; Expression = { 'TrustedDomain' }}, *
        }
    }
} finally { Remove-RASSession }
