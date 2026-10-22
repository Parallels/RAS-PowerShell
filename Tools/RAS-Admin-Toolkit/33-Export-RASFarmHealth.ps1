[CmdletBinding()]
param(
    [string]$Server,
    [switch]$UseEnvironment,
    [string]$Path = ".\RAS-FarmHealth-$(Get-Date -Format 'yyyyMMdd-HHmmss').html"
)

. "$PSScriptRoot\_Connect-RASFarm.ps1"
$null = Connect-RASFarm -Server $Server -UseEnvironment:$UseEnvironment
try {
    $fullPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
    $parent = Split-Path $fullPath -Parent
    if (-not (Test-Path -LiteralPath $parent -PathType Container)) { throw "Folder does not exist: $parent" }

    $sections = @(
        '<h1>Parallels RAS Farm Health</h1>'
        "<p>Generated $(Get-Date -Format 'u')</p>"
        (Get-RASAgent -ServerType Site | ConvertTo-Html -Fragment -PreContent '<h2>Sites</h2>')
        (Get-RASAgent | ConvertTo-Html -Fragment -PreContent '<h2>Agents</h2>')
        (Get-RASAgent -ServerType RDSHost | ConvertTo-Html -Fragment -PreContent '<h2>RDS Hosts</h2>')
        (Get-RASAgent -ServerType Gateway | ConvertTo-Html -Fragment -PreContent '<h2>Gateways</h2>')
    )
    ConvertTo-Html -Title 'Parallels RAS Farm Health' -Body $sections |
        Set-Content -LiteralPath $fullPath -Encoding utf8
    Get-Item -LiteralPath $fullPath
} finally { Remove-RASSession }
