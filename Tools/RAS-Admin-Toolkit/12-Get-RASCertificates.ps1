[CmdletBinding()]
param([string]$Server, [switch]$UseEnvironment)

. "$PSScriptRoot\_Connect-RASFarm.ps1"
$null = Connect-RASFarm -Server $Server -UseEnvironment:$UseEnvironment
try {
    Get-RASCertificate |
        Select-Object Id, SiteId, Name, CommonName, Status, ExpirationDate, Usage, Enabled |
        Sort-Object SiteId, Name
} finally { Remove-RASSession }
