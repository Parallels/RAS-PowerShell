[CmdletBinding()]
param(
    [string]$Server,
    [switch]$UseEnvironment,
    [Parameter(Mandatory)][string]$User
)

. "$PSScriptRoot\_Connect-RASFarm.ps1"
$null = Connect-RASFarm -Server $Server -UseEnvironment:$UseEnvironment
try { Get-RASRDSession -Source All -User $User } finally { Remove-RASSession }
