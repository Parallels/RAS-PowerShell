function Connect-RASFarm {
    [CmdletBinding()]
    param(
        [string]$Server,
        [switch]$UseEnvironment
    )

    Import-Module RASAdmin -ErrorAction Stop

    if ($UseEnvironment) {
        $missing = 'RAS_SERVER', 'RAS_USERNAME', 'RAS_PASSWORD' |
            Where-Object { -not [Environment]::GetEnvironmentVariable($_) }
        if ($missing) {
            throw "Missing environment variable(s): $($missing -join ', ')"
        }

        $password = ConvertTo-SecureString $env:RAS_PASSWORD -AsPlainText -Force
        New-RASSession -Server $env:RAS_SERVER -Username $env:RAS_USERNAME -Password $password
        return
    }

    if (-not $Server) { $Server = Read-Host 'RAS Connection Broker' }
    $username = Read-Host 'RAS administrator (UPN)'
    $password = Read-Host 'Password' -AsSecureString
    New-RASSession -Server $Server -Username $username -Password $password
}
