[CmdletBinding()]
param(
    [string]$Server,
    [switch]$UseEnvironment,
    [Parameter(Mandatory)][string]$User,
    [Parameter(Mandatory)][string]$ClientDeviceOS,
    [uint32]$SiteId = 0,
    [string]$ClientDeviceName,
    [string]$PrivateIPAddress,
    [string]$HardwareID,
    [string]$GatewayIP,
    [uint32]$ThemeId
)

. "$PSScriptRoot\_Connect-RASFarm.ps1"
$null = Connect-RASFarm -Server $Server -UseEnvironment:$UseEnvironment
try {
    if (-not (Get-Command Get-RASPubEffectiveAccess -ErrorAction SilentlyContinue)) {
        throw 'Get-RASPubEffectiveAccess requires Parallels RAS PowerShell 21.2/API v5.'
    }
    $parameters = @{
        Username       = $User
        ClientDeviceOS = $ClientDeviceOS
    }
    'SiteId', 'ClientDeviceName', 'PrivateIPAddress', 'HardwareID', 'GatewayIP', 'ThemeId' |
        Where-Object { $PSBoundParameters.ContainsKey($_) } |
        ForEach-Object { $parameters[$_] = $PSBoundParameters[$_] }

    Get-RASPubEffectiveAccess @parameters | ForEach-Object {
        Get-RASPubItem -Id ([uint32]$_) | Select-Object Id, Name, Type, EnabledMode, Description
    }
} finally { Remove-RASSession }
