[CmdletBinding()]
param([string]$Server, [switch]$UseEnvironment)

. "$PSScriptRoot\_Connect-RASFarm.ps1"
$null = Connect-RASFarm -Server $Server -UseEnvironment:$UseEnvironment
try { Get-RASAgent -ServerType Site | Sort-Object SiteId, Server } finally { Remove-RASSession }
