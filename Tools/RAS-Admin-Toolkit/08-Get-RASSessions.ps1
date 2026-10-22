[CmdletBinding()]
param(
    [string]$Server,
    [switch]$UseEnvironment,
    [ValidateSet('RDS', 'VDI', 'AVD', 'All')]
    [string]$Source = 'All',
    [ValidateSet('Active', 'Connected', 'ConnectQuery', 'Shadow', 'Disconnected', 'Idle', 'Listen', 'Reset', 'Down', 'Init', 'All')]
    [string]$State = 'All'
)

. "$PSScriptRoot\_Connect-RASFarm.ps1"
$null = Connect-RASFarm -Server $Server -UseEnvironment:$UseEnvironment
try { Get-RASRDSession -Source $Source -State $State } finally { Remove-RASSession }
