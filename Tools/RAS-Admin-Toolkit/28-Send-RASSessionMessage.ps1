[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
param(
    [string]$Server,
    [switch]$UseEnvironment,
    [Parameter(Mandatory)][string]$User,
    [Parameter(Mandatory)][string]$Message,
    [string]$Title = 'Message from IT'
)

. "$PSScriptRoot\_Connect-RASFarm.ps1"
$null = Connect-RASFarm -Server $Server -UseEnvironment:$UseEnvironment
try {
    $sessions = Get-RASRDSession -Source RDS -User $User
    foreach ($session in $sessions) {
        if ($PSCmdlet.ShouldProcess("$($session.User) on $($session.Server), session $($session.Id)", 'Send message')) {
            $session | Invoke-RASRDSHostSessionCmd -Command SendMsg -MsgTitle $Title -Message $Message
        }
    }
} finally { Remove-RASSession }
