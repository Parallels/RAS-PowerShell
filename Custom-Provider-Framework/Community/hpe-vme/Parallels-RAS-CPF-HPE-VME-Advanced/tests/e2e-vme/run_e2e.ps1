# e2e-vme: the RAS flows end to end, in the order RAS runs them, against the
# provider as a subprocess and mock_vme.py over HTTPS:
#   template -> sequential pool build -> sweep -> maintenance join/exit ->
#   recreate a pool member -> deploy-event gate -> delete the pool.
# Budget, safety and failure-mode detail lives in ../unit-vme/vme-u*.ps1.
. (Join-Path $PSScriptRoot '../unit-vme/lib.ps1')
$script:MockPort = 8444

function Send-Convert { param($P, [string]$Id, [bool]$T) Get-Result (Send-Rpc $P "{`"method`":`"guests/convert`",`"params`":{`"id`":`"$Id`",`"is_template`":$($T.ToString().ToLowerInvariant())}}") 'convert' }
function Send-Control { param($P, [string]$Id, [string]$C) Send-Rpc $P "{`"method`":`"guests/control`",`"params`":{`"id`":`"$Id`",`"control`":`"$C`"}}" }
function Start-Clone { param($P, [string]$Src, [string]$Name) Get-Result (Send-Rpc $P "{`"method`":`"guests/clone`",`"params`":{`"id`":`"$Src`",`"name`":`"$Name`"}}") "clone $Name" }

Start-VmeMock -Count 30
Invoke-Mock '/_ctl/set?clone_delay=1&provision_delay=2&mac_delay=1&ip_delay=2&delete_delay=2&action_delay=1' | Out-Null
$dir = New-ProviderDir -Name 'e2e' -Settings @{
    inventory = @{ list_min_refresh_seconds = 0 }
    cloning   = @{ tracked_refresh_interval_seconds = 1; mac_preservation = @{ enabled = $true } }
    delete    = @{ probe_interval_seconds = 1 }
}
# The run folder lives next to this script, not under unit-vme.
$p = Start-Provider -Dir $dir
try {
    Write-Host "`n=== connect ===" -ForegroundColor Cyan
    $init = Get-Result (Send-Rpc $p '{"method":"provider/initialize"}') 'initialize'
    Assert ($init.capabilities.template_method -eq 'basic' -and $init.capabilities.can_link_clones -eq $false) 'initialize advertises basic templates, no linked clones'
    Assert ($init.capabilities.tasks_polling_rate -eq 10 -and $init.capabilities.guests_polling_rate -eq 60 -and $init.capabilities.tasks_polling_retries -le 180) 'v3 polling defaults: tasks 10 s, guests 60 s, retries within CPF range'
    Assert ((Get-Result (Send-Rpc $p "{`"method`":`"provider/connect`",`"params`":{`"settings`":{`"host`":`"127.0.0.1:8444`",`"username`":`"admin`",`"password`":`"password`",`"ssl_check`":false}}}") 'connect').connected) 'provider/connect (password grant)'
    $ids = Get-GuestList $p
    Assert ($ids.Count -eq 36) "guests/list: $($ids.Count) guests"

    Write-Host "`n=== template ===" -ForegroundColor Cyan
    Send-Convert $p '19' $true | Out-Null
    Assert ((Get-Result (Get-Guest $p '19') 'get 19').is_template) 'instance 19 is the pool template'

    Write-Host "`n=== sequential pool build (3 members) ===" -ForegroundColor Cyan
    $members = @{}
    $t0 = Get-Date
    foreach ($n in 1..3) {
        $name = "POOL-00$n"
        $task = (Start-Clone $p '19' $name).task_id
        $t = Wait-Task $p $task -TimeoutSec 40
        Assert ($t.state -eq 'completed') "$name completed"
        $members[$name] = [string]$t.output.clone_id
    }
    $build = ((Get-Date) - $t0).TotalSeconds
    Assert ($build -lt 20) "3 sequential clones in $([Math]::Round($build,1)) s (each ~3 s of VME provisioning; no wait for boot or IP)"
    $list = Get-GuestList $p
    Assert (@($members.Values | Where-Object { $list -contains $_ }).Count -eq 3) 'all members are listed'

    Write-Host "`n=== sweep until the members are up ===" -ForegroundColor Cyan
    $deadline = (Get-Date).AddSeconds(20); $up = 0
    while ((Get-Date) -lt $deadline) {
        $up = 0
        foreach ($id in (Get-GuestList $p)) {
            $g = Get-Result (Get-Guest $p $id) "get $id"
            if ($members.Values -contains $id -and $g.state -eq 'powered_on' -and @($g.ip_addresses).Count -gt 0 -and @($g.mac_addresses).Count -gt 0) { $up++ }
        }
        if ($up -eq 3) { break }
        Start-Sleep -Seconds 1
    }
    Assert ($up -eq 3) 'every member reaches powered_on with an IP and a MAC through ordinary guests/get polling'

    Write-Host "`n=== maintenance: join, update, exit ===" -ForegroundColor Cyan
    $tk = Send-Convert $p '19' $false
    Assert ((Get-Task $p $tk.task_id).state -eq 'completed') 'enter maintenance (convert false)'
    Assert (-not (Get-Result (Get-Guest $p '19') 'get').is_template) 'template reports is_template=false while in maintenance'
    [void](Send-Control $p '19' 'start')
    Start-Sleep -Seconds 2
    Assert ((Get-Result (Get-Guest $p '19') 'get').state -eq 'powered_on') 'RAS starts it for the update'
    [void](Send-Control $p '19' 'stop')
    Send-Convert $p '19' $true | Out-Null
    $g19 = Get-Result (Get-Guest $p '19') 'get'
    Assert ($g19.is_template -and $g19.state -eq 'powered_off') 'exit maintenance: stopped and a template again'

    Write-Host "`n=== recreate a pool member (delete + clone under the same name) ===" -ForegroundColor Cyan
    $old = $members['POOL-002']
    $oldMac = (Invoke-Mock '/_ctl/state' 'GET').instances.$old.mac
    [void](Send-Control $p $old 'delete')
    $deadline = (Get-Date).AddSeconds(15)
    while ((Get-Date) -lt $deadline -and (Get-GuestList $p) -contains $old) { Start-Sleep -Milliseconds 500 }
    Assert ((Get-GuestList $p) -notcontains $old) 'the old member leaves the list only once VME returns 404'
    $task = (Start-Clone $p '19' 'POOL-002').task_id
    $t = Wait-Task $p $task -TimeoutSec 40
    Assert ($t.state -eq 'completed' -and [string]$t.output.clone_id -ne $old) "the recreate completes with a NEW id ($($t.output.clone_id), old $old)"
    $new = [string]$t.output.clone_id
    Assert ((Invoke-Mock '/_ctl/state' 'GET').instances.$new.mac -eq $oldMac) 'MAC preservation: the recreated member carries the old MAC'

    Write-Host "`n=== a clone that VME provisions powered off ===" -ForegroundColor Cyan
    Invoke-Mock '/_ctl/set?clone_final_status=stopped' | Out-Null
    $task = (Start-Clone $p '19' 'POOL-004').task_id
    $t = Wait-Task $p $task -TimeoutSec 40
    Assert ($t.state -eq 'completed') "a clone ending 'stopped' (not busy) counts as provisioned"
    $p4 = [string]$t.output.clone_id
    $deadline = (Get-Date).AddSeconds(20); $g = $null
    while ((Get-Date) -lt $deadline) {
        $g = Get-Result (Get-Guest $p $p4) 'get POOL-004'
        if ($g.state -eq 'powered_on' -and @($g.ip_addresses).Count -gt 0) { break }
        Start-Sleep -Seconds 1
    }
    Assert ($g.state -eq 'powered_on') 'the clone-aware guests/get flow starts it (safety-net start) and it comes up'
    Assert ((Get-LogText $p) -match "Clone-aware flow: starting instance \[$p4\]") 'the self-issued start is logged'
    Invoke-Mock '/_ctl/set?clone_final_status=running' | Out-Null
    $members['POOL-004'] = $p4

    Write-Host "`n=== deploy-event gate (experimental) ===" -ForegroundColor Cyan
    Stop-Provider $p
    Invoke-Mock '/_ctl/set?provision_delay=8' | Out-Null
    $dir2 = New-ProviderDir -Name 'e2e-gate' -Settings @{ inventory = @{ list_min_refresh_seconds = 0 }; cloning = @{ deploy_event_gate_enabled = $true; max_concurrent_clone_operations = 2 } }
    Copy-Item (Join-Path $dir 'VME-RAS-BasicTemplates.json') $dir2
    $p = Start-Provider -Dir $dir2
    Connect-Provider $p | Out-Null
    [void](Get-GuestList $p)
    Reset-MockStats
    $t0 = Get-Date
    $task = (Start-Clone $p '19' 'GATE-001').task_id
    $t = Wait-Task $p $task -TimeoutSec 40
    $secs = ((Get-Date) - $t0).TotalSeconds
    Assert ($t.state -eq 'completed' -and $secs -lt 8) "the gate reports completion once the Deploy step is done, before provisioning ends ($([Math]::Round($secs,1)) s < 1 + 8 s)"
    $st = Get-MockStats
    Assert ((Get-MockCount $st 'GET /api/processes') -ge 1) "the gate reads process history ($(Get-MockCount $st 'GET /api/processes') calls, one per pending poll)"
    $gid = [string]$t.output.clone_id
    Assert (-not ((Get-GuestList $p) -notcontains $gid)) 'the early-completed clone is listed'
    $g = Get-Result (Get-Guest $p $gid) 'get gate clone'
    Assert ($g.state -eq 'powering_on') 'while VME is still provisioning it, guests/get reports powering_on'

    Write-Host "`n=== delete the pool ===" -ForegroundColor Cyan
    Invoke-Mock '/_ctl/set?provision_delay=2' | Out-Null
    $all = @($members['POOL-001'], $members['POOL-003'], $members['POOL-004'], $new, $gid)
    Reset-MockStats
    foreach ($id in $all) { [void](Send-Control $p $id 'delete') }
    # With MAC preservation on, a delete whose MAC is not cached yet (the gate clone may
    # still be provisioning) captures it with one GET /api/servers/{id} first -- the only
    # extra call allowed.
    $stD = Get-MockStats
    $srvD = Get-MockCount $stD 'GET /api/servers/{id}'
    Assert ((Get-MockCount $stD 'DELETE /api/instances/{id}') -eq 5 -and (Get-MockTotal $stD) -eq (5 + $srvD) -and $srvD -le 5) "five deletes, five DELETE calls (+ $srvD MAC capture GET(s) for uncached MACs, nothing else)"
    $deadline = (Get-Date).AddSeconds(20)
    while ((Get-Date) -lt $deadline) {
        $l = Get-GuestList $p
        if (@($all | Where-Object { $l -contains $_ }).Count -eq 0) { break }
        Start-Sleep -Milliseconds 500
    }
    $l = Get-GuestList $p
    Assert (@($all | Where-Object { $l -contains $_ }).Count -eq 0) 'the whole pool leaves guests/list as VME confirms each removal'
    Assert ((Get-Result (Get-Guest $p '19') 'get').is_template) 'the template is untouched'

    [void](Send-Rpc $p '{"method":"provider/disconnect"}')
    Assert $true 'provider/disconnect'
}
catch { Assert $false "suite aborted: $($_.Exception.Message) at $($_.InvocationInfo.PositionMessage)" }
finally { Complete-Suite 'e2e-vme' }
