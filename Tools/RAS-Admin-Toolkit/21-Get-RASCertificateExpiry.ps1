[CmdletBinding()]
param([string]$Server, [switch]$UseEnvironment, [ValidateRange(1, 3650)][int]$Days = 30)

. "$PSScriptRoot\_Connect-RASFarm.ps1"
$null = Connect-RASFarm -Server $Server -UseEnvironment:$UseEnvironment
try {
    $cutoff = (Get-Date).AddDays($Days)
    Get-RASCertificate |
        Where-Object { $_.ExpirationDate -and $_.ExpirationDate -le $cutoff } |
        Select-Object SiteId, Name, CommonName, Status, ExpirationDate,
            @{Name = 'DaysRemaining'; Expression = { [math]::Floor(($_.ExpirationDate - (Get-Date)).TotalDays) }} |
        Sort-Object ExpirationDate
} finally { Remove-RASSession }
