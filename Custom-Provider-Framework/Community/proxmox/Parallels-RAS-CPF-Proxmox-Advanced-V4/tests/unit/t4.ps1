$ErrorActionPreference = 'Stop'
$pass = 0
$fail = 0
function Assert {
    param([bool]$Cond, [string]$Msg)
    if ($Cond) { $script:pass++; Write-Host "PASS: $Msg" -ForegroundColor Green }
    else { $script:fail++; Write-Host "FAIL: $Msg" -ForegroundColor Red }
}

Write-Host "--- WebSession reuse ---" -ForegroundColor Cyan
. "$PSScriptRoot/funcs2.ps1"
$script:LogPath = "$PSScriptRoot/t4.log"
Remove-Item $script:LogPath -ErrorAction SilentlyContinue
$script:ProxmoxWebSession = $null
$script:ProxmoxSession = @{ host = 'pve.test'; header = @{ Authorization = 'x' } }

$script:mockCalls = New-Object System.Collections.ArrayList
function Invoke-RestMethod {
    param($Uri, $Headers, $Method, $Body, [int]$TimeoutSec, [switch]$SkipCertificateCheck, [switch]$SkipHeaderValidation, $WebSession, [string]$SessionVariable, $ErrorAction)
    [void]$script:mockCalls.Add([pscustomobject]@{ HasWebSession = ($null -ne $WebSession); SessionVariable = $SessionVariable })
    if ($SessionVariable) {
        Set-Variable -Name $SessionVariable -Value ([pscustomobject]@{ FakeSession = $true }) -Scope 1
    }
    return @{ data = 'ok' }
}

[void](Invoke-ProxmoxApi -Method GET -Path '/api2/json/version')
[void](Invoke-ProxmoxApi -Method GET -Path '/api2/json/version')
[void](Invoke-ProxmoxApi -Method GET -Path '/api2/json/version')

Assert ($script:mockCalls.Count -eq 3) "3 calls made"
Assert (-not $script:mockCalls[0].HasWebSession -and $script:mockCalls[0].SessionVariable -eq 'proxmoxWebSessionCapture') "1st call captures a new session (no WebSession passed yet)"
Assert ($script:mockCalls[1].HasWebSession -and [string]::IsNullOrEmpty($script:mockCalls[1].SessionVariable)) "2nd call reuses the captured session (WebSession passed, no SessionVariable)"
Assert ($script:mockCalls[2].HasWebSession) "3rd call also reuses the session"
Assert ($null -ne $script:ProxmoxWebSession) "script:ProxmoxWebSession is now populated"

# Handle-Connect and Handle-Disconnect must reset it
function Invoke-ProxmoxApi { param($Method,$Path,$Body,[int]$TimeoutSec=0) return @{ data = @{ version = '8.1' } } }
[void](Handle-Connect -Params ([pscustomobject]@{ settings = [pscustomobject]@{ host='h'; username='u'; token_name='t'; token_secret='s' } }))
Assert ($null -eq $script:ProxmoxWebSession) "Handle-Connect resets ProxmoxWebSession to null"

$script:ProxmoxWebSession = [pscustomobject]@{ FakeSession = $true }
[void](Handle-Disconnect)
Assert ($null -eq $script:ProxmoxWebSession) "Handle-Disconnect resets ProxmoxWebSession to null"

Write-Host "`n--- Log rotation ---" -ForegroundColor Cyan
. "$PSScriptRoot/funcs2.ps1"
$script:LogPath = "$PSScriptRoot/t4-rotate.log"
Remove-Item "$($script:LogPath)*" -ErrorAction SilentlyContinue
$script:LogRotateMaxBytes = 200
$script:LogRotateMaxGenerations = 2
$script:LogBytesWrittenSinceStart = 0

1..40 | ForEach-Object { Write-DebugLog "line number $_ of a moderately long test message to accumulate bytes" }

$rotated1 = Test-Path "$($script:LogPath).1"
Assert $rotated1 "Fix: log rotation created a .1 generation once the size threshold was crossed"
$currentSize = (Get-Item $script:LogPath).Length
Assert ($currentSize -lt 5000) "current log file is small again after rotation (not left to grow unbounded)"

# Force enough rotations to verify the generation cap holds (oldest gets deleted, not endlessly renamed)
1..200 | ForEach-Object { Write-DebugLog "more lines to force several rotations $_" }
$gen2Exists = Test-Path "$($script:LogPath).2"
$gen3Exists = Test-Path "$($script:LogPath).3"
Assert $gen2Exists "a .2 generation exists after enough rotations"
Assert (-not $gen3Exists) "Fix: generation cap (2) holds -- no .3 generation is ever created"

Write-Host "`n--- Send-Response guarded fallback (does not crash the process on a broken pipe) ---" -ForegroundColor Cyan
# `exit` inside Send-Response's guard would terminate whatever process calls it,
# so this has to run in a CHILD process: verify its exit code is a clean 0, not
# an unhandled-exception crash (non-zero / stack trace on stderr).
$childScript = @'
$ErrorActionPreference = "Stop"
. "$env:TESTDIR/funcs2.ps1"
$script:LogPath = "$env:TESTDIR/t4-sendresponse-child.log"
Remove-Item $script:LogPath -ErrorAction SilentlyContinue
$writer = [pscustomobject]@{}
$writer | Add-Member -MemberType ScriptMethod -Name WriteLine -Value { param($x) throw "broken pipe (simulated)" }
Send-Response -ResponseObject @{ result = @{ ok = $true } }
Write-Host "SHOULD NOT REACH HERE"
'@
$childScript | Set-Content "$PSScriptRoot/t4-child.ps1"
$env:TESTDIR = $PSScriptRoot
& pwsh -NoProfile -File "$PSScriptRoot/t4-child.ps1" 2>$null
$childExit = $LASTEXITCODE
Assert ($childExit -eq 0) "Fix: Send-Response exits cleanly (code 0) instead of crashing when BOTH the primary and fallback stdout writes fail"

Write-Host "`n=== Summary: $pass passed, $fail failed ===" -ForegroundColor $(if ($fail -eq 0) { 'Green' } else { 'Red' })
if ($fail -gt 0) { exit 1 }
