[CmdletBinding()]
param([string]$Server, [switch]$UseEnvironment)

. "$PSScriptRoot\_Connect-RASFarm.ps1"
$null = Connect-RASFarm -Server $Server -UseEnvironment:$UseEnvironment
try {
    $sessions = Get-RASRDSession -Source RDS
    foreach ($hostStatus in Get-RASAgent -ServerType RDSHost) {
        [pscustomobject]@{
            SiteId = $hostStatus.SiteId
            Server = $hostStatus.Server
            State = $hostStatus.AgentState
            CPULoad = $hostStatus.CPULoad
            MemoryLoad = $hostStatus.MemLoad
            Sessions = @($sessions | Where-Object Server -eq $hostStatus.Server).Count
        }
    }
} finally { Remove-RASSession }
