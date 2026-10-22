[CmdletBinding()]
param([string]$Server, [switch]$UseEnvironment, [ValidateRange(1, 3650)][int]$CertificateWarningDays = 30)

. "$PSScriptRoot\_Connect-RASFarm.ps1"
$null = Connect-RASFarm -Server $Server -UseEnvironment:$UseEnvironment
try {
    $cutoff = (Get-Date).AddDays($CertificateWarningDays)
    $unhealthyAgents = @(Get-RASAgent | Where-Object { $null -ne $_.AgentState -and $_.AgentState -ne 'OK' })
    $expiringCertificates = @(Get-RASCertificate | Where-Object { $_.ExpirationDate -and $_.ExpirationDate -le $cutoff })
    $disconnectedSessions = @(Get-RASRDSession -Source All -State Disconnected)
    $disabledResources = @(Get-RASPubItem | Where-Object {
        $_.EnabledMode -in 'Disabled', 'Maintenance' -or
        ($null -ne $_.Enabled -and -not $_.Enabled)
    })

    [pscustomobject]@{
        Category = 'Summary'
        Target = 'Farm'
        Status = if ($unhealthyAgents.Count + $expiringCertificates.Count + $disconnectedSessions.Count + $disabledResources.Count) { 'Attention' } else { 'Healthy' }
        Details = "Generated $(Get-Date -Format 'u'); agents=$($unhealthyAgents.Count); certificates=$($expiringCertificates.Count); disconnected=$($disconnectedSessions.Count); resources=$($disabledResources.Count)"
    }
    $unhealthyAgents | ForEach-Object {
        $target = @($_.Server, $_.Name, $_.TemplateName, $_.HostPoolName, "Id $($_.Id)") |
            Where-Object { $_ } | Select-Object -First 1
        [pscustomobject]@{ Category = 'Agent'; Target = $target; Status = $_.AgentState; Details = "$($_.ServerType); $($_.AgentVer)" }
    }
    $expiringCertificates | ForEach-Object {
        [pscustomobject]@{ Category = 'Certificate'; Target = $_.Name; Status = $_.Status; Details = "Expires $($_.ExpirationDate)" }
    }
    $disconnectedSessions | ForEach-Object {
        [pscustomobject]@{ Category = 'Session'; Target = $_.User; Status = 'Disconnected'; Details = "$($_.Server); session $($_.Id)" }
    }
    $disabledResources | ForEach-Object {
        [pscustomobject]@{ Category = 'Resource'; Target = $_.Name; Status = $(if ($_.EnabledMode) { $_.EnabledMode } else { 'Disabled' }); Details = $_.Type }
    }
} finally { Remove-RASSession }
