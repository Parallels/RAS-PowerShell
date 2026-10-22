[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [string]$Server,
    [switch]$UseEnvironment,
    [Parameter(Mandatory)][string]$RDSHost,
    [uint32]$SiteId = 0,
    [Parameter(Mandatory)][bool]$Enabled
)

. "$PSScriptRoot\_Connect-RASFarm.ps1"
$null = Connect-RASFarm -Server $Server -UseEnvironment:$UseEnvironment
try {
    if ($PSCmdlet.ShouldProcess($RDSHost, "Set RDS host Enabled=$Enabled and apply changes")) {
        Set-RASRDSHost -Server $RDSHost -SiteId $SiteId -Enabled $Enabled
        Invoke-RASApply
    }
} finally { Remove-RASSession }
