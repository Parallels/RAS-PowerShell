[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param([string]$Server, [switch]$UseEnvironment)

. "$PSScriptRoot\_Connect-RASFarm.ps1"
$null = Connect-RASFarm -Server $Server -UseEnvironment:$UseEnvironment
try {
    if ($PSCmdlet.ShouldProcess($Server, 'Apply pending RAS farm configuration changes')) {
        Invoke-RASApply
    }
} finally { Remove-RASSession }
