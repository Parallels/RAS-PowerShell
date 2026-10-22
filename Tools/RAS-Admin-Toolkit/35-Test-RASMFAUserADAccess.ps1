[CmdletBinding()]
param(
    [string]$Server,
    [switch]$UseEnvironment,
    [Parameter(Mandatory)][string]$UserPrincipalName,
    [uint32]$SiteId = 0,
    [string]$ADCustomAttribute
)

. "$PSScriptRoot\_Connect-RASFarm.ps1"
$null = Connect-RASFarm -Server $Server -UseEnvironment:$UseEnvironment
try {
    $parameters = @{ UserPrincipalName = $UserPrincipalName; ValidateADAccess = $true }
    if ($PSBoundParameters.ContainsKey('SiteId')) { $parameters.SiteId = $SiteId }
    if ($ADCustomAttribute) { $parameters.ADCustomAttribute = $ADCustomAttribute }
    Invoke-RASMFA @parameters | Format-List *
} finally { Remove-RASSession }
