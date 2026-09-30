[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [string]$Server,
    [switch]$UseEnvironment,
    [ValidateRange(0, 8760)][int]$IdleHours = 8
)

. "$PSScriptRoot\_Connect-RASFarm.ps1"
$null = Connect-RASFarm -Server $Server -UseEnvironment:$UseEnvironment
try {
    $minimumIdleSeconds = $IdleHours * 3600
    $sessions = Get-RASRDSession -Source RDS -State Disconnected |
        Where-Object IdleTime -GE $minimumIdleSeconds

    foreach ($session in $sessions) {
        if ($PSCmdlet.ShouldProcess("$($session.User) on $($session.Server), session $($session.Id)", 'Log off disconnected session')) {
            $session | Invoke-RASRDSHostSessionCmd -Command LogOff
        }
    }
} finally { Remove-RASSession }
