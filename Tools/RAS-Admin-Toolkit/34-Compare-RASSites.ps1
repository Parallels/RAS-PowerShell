[CmdletBinding()]
param(
    [string]$Server,
    [switch]$UseEnvironment,
    [Parameter(Mandatory)][uint32]$SiteId1,
    [Parameter(Mandatory)][uint32]$SiteId2
)

. "$PSScriptRoot\_Connect-RASFarm.ps1"
$null = Connect-RASFarm -Server $Server -UseEnvironment:$UseEnvironment
try {
    foreach ($siteId in $SiteId1, $SiteId2) {
        $site = Get-RASSite -Id $siteId
        [pscustomobject]@{
            SiteId = $siteId
            Site = $site.Name
            Brokers = @(Get-RASAgent -SiteId $siteId -ServerType Broker).Count
            Gateways = @(Get-RASAgent -SiteId $siteId -ServerType Gateway).Count
            RDSHosts = @(Get-RASAgent -SiteId $siteId -ServerType RDSHost).Count
            Providers = @(Get-RASAgent -SiteId $siteId -ServerType Provider).Count
            PublishedItems = @(Get-RASPubItem -SiteId $siteId).Count
            MFAProviders = @(Get-RASMFA -SiteId $siteId).Count
        }
    }
} finally { Remove-RASSession }
