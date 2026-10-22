# vme-u1: inventory, call budget and guest objects (review 8.2 vme-u1 + vme-u3).
#   - a 3000-instance guests/get sweep makes ZERO VME calls;
#   - one listing and one bulk servers pass per process, over a reused connection;
#   - MAC and OS come from the server objects; IP filtering; state mapping.
param([int]$Count = 3000)
. (Join-Path $PSScriptRoot 'lib.ps1')

Start-VmeMock -Count $Count
$dir = New-ProviderDir -Name 'u1'
$p = Start-Provider -Dir $dir
try {
    Write-Host "`n=== connect and first listing ($Count fillers) ===" -ForegroundColor Cyan
    Connect-Provider $p | Out-Null
    Assert (-not ($p.FirstLine -match "^\xEF\xBB\xBF" -or $p.FirstLine.StartsWith([char]0xFEFF))) 'the first response line carries no BOM'
    Reset-MockStats

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $ids = Get-GuestList $p
    $listMs = $sw.ElapsedMilliseconds
    $expected = $Count + 6
    Assert ($ids.Count -eq $expected) "guests/list returns every instance ($($ids.Count) of $expected) in ${listMs}ms"
    $st = Get-MockStats
    Assert ((Get-MockCount $st 'GET /api/instances list') -eq 1) "the cold-start listing is ONE request (max=-1), not pages (got $(Get-MockCount $st 'GET /api/instances list'))"
    Assert ((Get-LogText $p) -match 'Inventory refreshed \(guests/list, cold start, max=-1\)') 'the cold start is named in the refresh log line'
    Assert ((Get-MockCount $st 'GET /api/servers') -ge 1) "one bulk hardware pass over /api/servers ($(Get-MockCount $st 'GET /api/servers') pages)"
    Assert ((Get-MockCount $st 'GET /api/instances/{id} details') -eq 0) 'no details=true GET anywhere'
    Assert ($st.connections -le 2) "all $(Get-MockTotal $st) calls rode on at most 2 TLS connections (got $($st.connections)) -- connection reuse"

    Write-Host "`n=== sweep: $expected single-id guests/get ===" -ForegroundColor Cyan
    Reset-MockStats
    $sw.Restart()
    $bad = 0
    foreach ($id in $ids) {
        $r = Get-Guest $p $id
        if (Test-IsError $r) { $bad++ }
    }
    $sweepMs = $sw.ElapsedMilliseconds
    $st = Get-MockStats
    Assert ($bad -eq 0) "every guests/get in the sweep succeeded ($bad errors)"
    Assert ((Get-MockTotal $st) -eq 0) "the sweep made ZERO VME calls (got $(Get-MockTotal $st): $($st.requests | ConvertTo-Json -Compress))"
    $perCall = [Math]::Round($sweepMs / $expected, 2)
    Write-Host "    sweep: $sweepMs ms total, $perCall ms per guests/get (log_level 4)" -ForegroundColor DarkGray
    Assert ($perCall -lt 25) "per-guests/get cost including the stdio round trip is below 25 ms (got $perCall ms)"

    Reset-MockStats
    $again = Get-GuestList $p
    Assert ((Get-MockTotal (Get-MockStats)) -eq 0 -and $again.Count -eq $expected) 'a guests/list inside list_min_refresh_seconds is answered from memory (0 calls)'

    Write-Host "`n=== guest object fields ===" -ForegroundColor Cyan
    $g20 = Get-Result (Get-Guest $p '20') 'get 20'
    Assert ($g20.ip -eq '192.168.1.51' -and $g20.state -eq 'powered_on') 'IP from connectionInfo[].ip, running -> powered_on'
    Assert ($g20.mac_addresses -contains '52:54:00:00:00:18') "MAC from the server object (server id 24 for instance 20), got [$($g20.mac_addresses -join ',')]"
    Assert ($g20.host_os -eq 'windows 11 64-bit') "host_os from server.serverOs.name, not instanceType 'HVM' (got [$($g20.host_os)])"
    Assert ($g20.is_template -eq $false) 'an ordinary instance is not a template'

    $g21 = Get-Result (Get-Guest $p '21') 'get 21'
    Assert ($g21.state -eq 'powered_off' -and @($g21.ip_addresses).Count -eq 0) 'a stopped instance reports no IP even though VME still carries its last one'

    $g22 = Get-Result (Get-Guest $p '22') 'get 22'
    Assert ($g22.state -eq 'powered_off') "VME status 'failed' maps to a CPF state (powered_off), never 'unknown' (got [$($g22.state)])"

    $g23 = Get-Result (Get-Guest $p '23') 'get 23'
    Assert ($g23.ip_addresses -contains '169.254.10.10') 'a link-local-only address is reported (visible), not hidden'
    $g24 = Get-Result (Get-Guest $p '24') 'get 24'
    Assert (@($g24.ip_addresses).Count -eq 0) "the 0.0.0.0 placeholder is never reported as an address"

    $log = Get-LogText $p
    $failedLines = ([regex]::Matches($log, 'Instance \[22\] \(failed-vm\) has VME status \[failed\]')).Count
    Assert ($failedLines -eq 1) "the failed status is logged once (E), not on every sweep (got $failedLines)"
    Assert ($log -notmatch 'ssh-secret-value|server-api-key-value') 'server objects (sshPassword, apiKey) never reach the log'

    Write-Host "`n=== an instance not in the inventory ===" -ForegroundColor Cyan
    Invoke-Mock '/_ctl/instance?id=9777&name=late-vm&status=running&ip=10.7.7.7' | Out-Null
    Reset-MockStats
    $g9777 = Get-Result (Get-Guest $p '9777') 'get 777'
    $st = Get-MockStats
    Assert ($g9777.name -eq 'late-vm' -and $g9777.ip -eq '10.7.7.7') 'an id unknown to the inventory is fetched with one lean GET'
    Assert ((Get-MockCount $st 'GET /api/instances/{id}') -eq 1 -and (Get-MockCount $st 'GET /api/servers/{id}') -eq 1) 'cost: 1 lean instance GET + 1 server GET for its MAC, nothing else'
    Reset-MockStats
    [void](Get-Guest $p '9777')
    Assert ((Get-MockTotal (Get-MockStats)) -eq 0) 'the second guests/get of that id costs nothing'
    $missing = Get-Guest $p '888888'
    Assert (Test-IsError $missing) 'an id VME does not know returns an error (404)'
}
catch { Assert $false "suite aborted: $($_.Exception.Message) at $($_.InvocationInfo.PositionMessage)" }
finally { Complete-Suite 'vme-u1' }
