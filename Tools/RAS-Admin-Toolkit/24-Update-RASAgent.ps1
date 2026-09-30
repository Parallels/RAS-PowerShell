[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [string]$Server,
    [switch]$UseEnvironment,
    [Parameter(Mandatory)][string]$AgentServer,
    [uint32]$SiteId = 0,
    [switch]$Force
)

. "$PSScriptRoot\_Connect-RASFarm.ps1"
$null = Connect-RASFarm -Server $Server -UseEnvironment:$UseEnvironment
try {
    if ($PSCmdlet.ShouldProcess($AgentServer, 'Update RAS Agent')) {
        Update-RASAgent -Server $AgentServer -SiteId $SiteId -Force:$Force
    }
} finally { Remove-RASSession }
