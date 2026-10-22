[CmdletBinding()]
param([string]$Server, [switch]$UseEnvironment)

. "$PSScriptRoot\_Connect-RASFarm.ps1"
$null = Connect-RASFarm -Server $Server -UseEnvironment:$UseEnvironment
try {
    Get-RASAgent | Group-Object AgentVer | Sort-Object Count -Descending | ForEach-Object {
        [pscustomobject]@{
            AgentVersion = $_.Name
            Count = $_.Count
            Servers = ($_.Group.Server | Sort-Object) -join ', '
        }
    }
} finally { Remove-RASSession }
