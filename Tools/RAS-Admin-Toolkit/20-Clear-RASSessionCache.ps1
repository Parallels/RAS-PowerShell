[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param([string]$Server, [switch]$UseEnvironment)

. "$PSScriptRoot\_Connect-RASFarm.ps1"
$null = Connect-RASFarm -Server $Server -UseEnvironment:$UseEnvironment
try {
    if ($PSCmdlet.ShouldProcess($Server, 'Clear the session ID cache and force client reauthentication')) {
        Invoke-RASClearSessionsIDCache
    }
} finally { Remove-RASSession }
