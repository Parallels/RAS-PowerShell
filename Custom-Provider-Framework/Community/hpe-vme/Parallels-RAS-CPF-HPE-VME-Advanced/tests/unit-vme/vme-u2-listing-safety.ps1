# vme-u2: listing safety (review 6.3). A refresh must never make guests vanish by
# accident: failed pages, empty/partial answers and external deletes.
. (Join-Path $PSScriptRoot 'lib.ps1')

Start-VmeMock -Count 50
$dir = New-ProviderDir -Name 'u2' -Settings @{ inventory = @{ list_min_refresh_seconds = 0; page_size = 25 }; connection = @{ retry_delay_milliseconds = 10 } }
$p = Start-Provider -Dir $dir
try {
    Connect-Provider $p | Out-Null
    $base = Get-GuestList $p
    Assert ($base.Count -eq 56) "baseline listing: 56 instances (got $($base.Count))"
    Assert ((Get-MockCount (Get-MockStats) 'GET /api/instances list') -eq 1) 'the cold-start listing is one request (max=-1); routine listings below page by 25'

    Write-Host "`n=== a listing page that keeps failing ===" -ForegroundColor Cyan
    Invoke-Mock '/_ctl/fail?key=GET%20/api/instances&count=3&code=500' | Out-Null
    $r = Get-GuestList $p
    Assert ($r.Count -eq 56) "3 consecutive 5xx on the listing: guests/list still answers from the previous inventory (got $($r.Count))"
    $log = Get-LogText $p
    Assert ($log -match 'Inventory refresh \(guests/list\) failed') 'the failed refresh is logged as W'
    Assert ($log -match '\(attempt 1/3\) -- retrying') 'each retried GET attempt is logged at W, before the final failure'

    Invoke-Mock '/_ctl/fail?key=GET%20/api/instances&count=1&code=503' | Out-Null
    Reset-MockStats
    $r = Get-GuestList $p
    $st = Get-MockStats
    Assert ($r.Count -eq 56 -and (Get-MockCount $st 'GET /api/instances list') -eq 4) 'a single 503 is retried transparently (4 page requests for 3 pages)'

    Write-Host "`n=== an empty listing (permission/tenant glitch) ===" -ForegroundColor Cyan
    Invoke-Mock '/_ctl/list?empty=1' | Out-Null
    $r = Get-GuestList $p
    Assert ($r.Count -eq 56) "one empty listing does not empty RAS's view (got $($r.Count))"
    Assert ((Get-LogText $p) -match 'Not removing anything until the next listing confirms it') 'the mass-drop guard fires with an E line'
    $r = Get-GuestList $p
    Assert ($r.Count -eq 56) 'the next normal listing restores normal operation'

    Write-Host "`n=== an instance deleted outside RAS ===" -ForegroundColor Cyan
    Invoke-Mock '/_ctl/instance?id=120&delete=1' | Out-Null
    $r1 = Get-GuestList $p
    Assert ($r1 -contains '120') 'absent from ONE listing: still listed (offset paging during churn can skip an id)'
    $r2 = Get-GuestList $p
    Assert ($r2 -notcontains '120' -and $r2.Count -eq 55) 'absent from two consecutive listings: dropped'

    Write-Host "`n=== an instance created outside RAS ===" -ForegroundColor Cyan
    Invoke-Mock '/_ctl/instance?id=5000&name=new-vm&status=running&ip=10.5.5.5' | Out-Null
    $r = Get-GuestList $p
    Assert ($r -contains '5000') 'a new instance appears at the next listing'

    Write-Host "`n=== exclusions ===" -ForegroundColor Cyan
    Stop-Provider $p
    $dir2 = New-ProviderDir -Name 'u2b' -Settings @{ inventory = @{ list_min_refresh_seconds = 0; exclude_instance_name_patterns = @('^filler-10[0-4]$', '[invalid') } }
    $p2 = Start-Provider -Dir $dir2
    Connect-Provider $p2 | Out-Null
    $r = Get-GuestList $p2
    Assert ($r -notcontains '100' -and $r -notcontains '104' -and $r -contains '105') 'exclude_instance_name_patterns hides matching instances'
    [void](Get-GuestList $p2)
    $log = Get-LogText $p2
    Assert (([regex]::Matches($log, 'Instance \[100\] \(filler-100\) hidden')).Count -le 1) 'an excluded instance is logged once, not on every listing'
    Assert (([regex]::Matches($log, '\[invalid\] is not a valid regex')).Count -eq 1) 'an invalid pattern is reported once at load, not per instance'

    Write-Host "`n=== first listing fails with nothing to fall back on ===" -ForegroundColor Cyan
    $dir3 = New-ProviderDir -Name 'u2c' -Settings @{ connection = @{ retry_delay_milliseconds = 10 } }
    $p3 = Start-Provider -Dir $dir3
    Connect-Provider $p3 | Out-Null
    # 3 attempts for the cold-start request, 3 more for the paged fallback's first page.
    Invoke-Mock '/_ctl/fail?key=GET%20/api/instances&count=6&code=500' | Out-Null
    $resp = Send-Rpc $p3 '{"method":"guests/list"}'
    Assert (Test-IsError $resp) 'with no earlier inventory, a failed listing is an error, never an empty list'
    Assert ((Get-LogText $p3) -match 'Cold-start listing \(max=-1\) failed: .* continuing with pages of 200') 'a failed cold-start request falls back to paging before giving up'
    Stop-Provider $p3

    $expected = @((Invoke-Mock '/_ctl/state' 'GET').instances.PSObject.Properties).Count

    Write-Host "`n=== cold start: VME caps max=-1 ===" -ForegroundColor Cyan
    Invoke-Mock '/_ctl/set?list_max_cap=30' | Out-Null
    $p4 = Start-Provider -Dir (New-ProviderDir -Name 'u2-cap' -Settings @{ inventory = @{ page_size = 25 } })
    Connect-Provider $p4 | Out-Null
    Reset-MockStats
    $r = Get-GuestList $p4
    $n = Get-MockCount (Get-MockStats) 'GET /api/instances list'
    Assert ($r.Count -eq $expected) "a capped cold start still lists the whole fleet ($($r.Count) of $expected)"
    Assert ($n -eq (1 + [Math]::Ceiling(($expected - 30) / 25))) "1 capped request (30 rows) + pages of 25 from offset 30 (got $n requests)"
    Assert ((Get-LogText $p4) -match "Cold-start listing \(max=-1\) returned 30 of $expected instances -- continuing with pages of 25 from offset 30") 'the cap is reported once at W'
    Stop-Provider $p4
    Invoke-Mock '/_ctl/set?list_max_cap=0' | Out-Null

    Write-Host "`n=== cold start: VME refuses max=-1 ===" -ForegroundColor Cyan
    Invoke-Mock '/_ctl/set?reject_max_all=true' | Out-Null
    $p5 = Start-Provider -Dir (New-ProviderDir -Name 'u2-reject' -Settings @{ inventory = @{ page_size = 25 } })
    Connect-Provider $p5 | Out-Null
    Reset-MockStats
    $r = Get-GuestList $p5
    $n = Get-MockCount (Get-MockStats) 'GET /api/instances list'
    Assert ($r.Count -eq $expected) "a refused max=-1 falls back to pages and lists the whole fleet ($($r.Count) of $expected)"
    Assert ($n -eq (1 + [Math]::Ceiling($expected / 25))) "1 refused request + $([Math]::Ceiling($expected / 25)) pages of 25 (got $n)"
    Stop-Provider $p5
    Invoke-Mock '/_ctl/set?reject_max_all=false' | Out-Null

    Write-Host "`n=== cold_start_page_size = 0: paged cold start ===" -ForegroundColor Cyan
    $p6 = Start-Provider -Dir (New-ProviderDir -Name 'u2-paged' -Settings @{ inventory = @{ page_size = 25; cold_start_page_size = 0 } })
    Connect-Provider $p6 | Out-Null
    Reset-MockStats
    $r = Get-GuestList $p6
    $n = Get-MockCount (Get-MockStats) 'GET /api/instances list'
    Assert ($r.Count -eq $expected -and $n -eq [Math]::Ceiling($expected / 25)) "cold start paged like a routine listing ($n requests)"
    Stop-Provider $p6

    Write-Host "`n=== a routine pass slower than the RAS polling cycle ===" -ForegroundColor Cyan
    $p7 = Start-Provider -Dir (New-ProviderDir -Name 'u2-slow' -Settings @{ inventory = @{ page_size = 25; list_min_refresh_seconds = 0 }; capabilities = @{ guests_polling_rate = 5 } })
    Connect-Provider $p7 | Out-Null
    [void](Get-GuestList $p7)
    Invoke-Mock '/_ctl/set?list_delay=2' | Out-Null
    $r = Get-GuestList $p7
    Invoke-Mock '/_ctl/set?list_delay=0' | Out-Null
    $log = Get-LogText $p7
    Assert ($r.Count -eq $expected) 'the slow pass is still applied'
    $pg = [Math]::Ceiling($expected / 25)
    Assert ($log -match "Routine instance listing took [\d.]+s for $pg page\(s\) of 25 -- longer than one RAS polling cycle \(capabilities.guests_polling_rate = 5s\)") 'a routine pass longer than guests_polling_rate is reported at W'
    Assert (([regex]::Matches($log, 'longer than one RAS polling cycle')).Count -eq 1) 'the fast cold start and fast passes are not reported'
    Stop-Provider $p7
}
catch { Assert $false "suite aborted: $($_.Exception.Message) at $($_.InvocationInfo.PositionMessage)" }
finally { Complete-Suite 'vme-u2' }
