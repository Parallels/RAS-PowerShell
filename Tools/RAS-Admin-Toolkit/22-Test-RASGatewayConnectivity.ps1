[CmdletBinding()]
param(
    [string]$Server,
    [switch]$UseEnvironment,
    [ValidateNotNullOrEmpty()][int[]]$Port = @(443)
)

. "$PSScriptRoot\_Connect-RASFarm.ps1"
$null = Connect-RASFarm -Server $Server -UseEnvironment:$UseEnvironment
try {
    foreach ($gateway in Get-RASGateway) {
        foreach ($number in $Port) {
            $test = Test-NetConnection -ComputerName $gateway.Server -Port $number -WarningAction SilentlyContinue
            [pscustomobject]@{
                SiteId = $gateway.SiteId
                Gateway = $gateway.Server
                Port = $number
                DNSResolved = [bool]$test.ResolvedAddresses
                TcpSucceeded = $test.TcpTestSucceeded
            }
        }
    }
} finally { Remove-RASSession }
