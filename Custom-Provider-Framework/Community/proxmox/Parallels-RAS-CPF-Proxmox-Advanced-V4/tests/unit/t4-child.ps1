$ErrorActionPreference = "Stop"
. "$env:TESTDIR/funcs2.ps1"
$script:LogPath = "$env:TESTDIR/t4-sendresponse-child.log"
Remove-Item $script:LogPath -ErrorAction SilentlyContinue
$writer = [pscustomobject]@{}
$writer | Add-Member -MemberType ScriptMethod -Name WriteLine -Value { param($x) throw "broken pipe (simulated)" }
Send-Response -ResponseObject @{ result = @{ ok = $true } }
Write-Host "SHOULD NOT REACH HERE"
