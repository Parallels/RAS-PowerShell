[CmdletBinding()]
param([string]$Server, [switch]$UseEnvironment)

. "$PSScriptRoot\_Connect-RASFarm.ps1"
$null = Connect-RASFarm -Server $Server -UseEnvironment:$UseEnvironment
try {
    foreach ($site in Get-RASSite) {
        foreach ($mfa in Get-RASMFA -SiteId $site.Id | Where-Object Type -In GAuthTOTP, TOTP, MicrosoftTOTP, EmailOTP) {
            $parameters = @{ Like = '@'; SiteId = $site.Id; Type = $mfa.Type }
            if ($mfa.Type -eq 'EmailOTP') { $parameters.MFAId = $mfa.Id }
            Find-RASMFAUsers @parameters | ForEach-Object {
                [pscustomobject]@{ Site = $site.Name; MFAProvider = $mfa.Name; MFAType = $mfa.Type; User = $_ }
            }
        }
    }
} finally { Remove-RASSession }
