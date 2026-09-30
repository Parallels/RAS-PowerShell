[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [string]$Server,
    [switch]$UseEnvironment,
    [Parameter(Mandatory)]
    [string[]]$User,
    [ValidateSet('GAuthTOTP', 'TOTP', 'MicrosoftTOTP', 'EmailOTP')]
    [string]$Type = 'TOTP',
    [uint32]$SiteId = 0,
    [uint32]$MFAId = 0
)

. "$PSScriptRoot\_Connect-RASFarm.ps1"
$null = Connect-RASFarm -Server $Server -UseEnvironment:$UseEnvironment
try {
    if ($PSCmdlet.ShouldProcess(($User -join ', '), "Reset $Type MFA enrollment")) {
        Reset-RASMFAUsers -Users $User -Type $Type -SiteId $SiteId -MFAId $MFAId
    }
} finally { Remove-RASSession }
