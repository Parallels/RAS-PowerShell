# vme-u3: the clone path (review 4, 8.2 vme-u4).
. (Join-Path $PSScriptRoot 'lib.ps1')

function Start-Clone {
    param($P, [string]$Source, [string]$Name, [int]$TimeoutSec = 60)
    return (Send-Rpc $P "{`"method`":`"guests/clone`",`"params`":{`"id`":`"$Source`",`"name`":`"$Name`"}}" -TimeoutSec $TimeoutSec)
}
function Get-MockInstancesNamed {
    param([string]$Name)
    $s = Invoke-Mock '/_ctl/state' 'GET'
    return @($s.instances.PSObject.Properties | Where-Object { $_.Value.name -eq $Name } | ForEach-Object { $_.Name })
}

Start-VmeMock -Count 20
Invoke-Mock '/_ctl/set?clone_delay=2&provision_delay=2&mac_delay=1&ip_delay=2&delete_delay=1' | Out-Null
$dir = New-ProviderDir -Name 'u3' -Settings @{
    inventory  = @{ list_min_refresh_seconds = 0 }
    cloning    = @{ resolve_list_interval_seconds = 10; tracked_refresh_interval_seconds = 1; mac_preservation = @{ enabled = $true } }
    delete     = @{ probe_interval_seconds = 1 }
    connection = @{ request_timeout_seconds = 5; retry_delay_milliseconds = 10 }
    logging    = @{ log_level = 5 }
}
$p = Start-Provider -Dir $dir
try {
    Connect-Provider $p | Out-Null
    [void](Get-GuestList $p)
    $conv = Get-Result (Send-Rpc $p '{"method":"guests/convert","params":{"id":"19","is_template":true}}') 'convert'
    Assert ($conv.task_id -eq 'noop:convert:19') 'convert(true) returns a synthetic task'
    [void](Get-GuestList $p)   # the template's clone-body inputs now come for free from the listing
    $g19 = Get-Result (Get-Guest $p '19') 'get 19'
    Assert ($g19.is_template -eq $true) 'the converted instance reports is_template=true'

    Write-Host "`n=== submission ===" -ForegroundColor Cyan
    Reset-MockStats
    $c1 = Get-Result (Start-Clone $p '19' 'RDSH-001') 'clone'
    $st = Get-MockStats
    Assert ($c1.task_id -like 'clonebyname:19:RDSH-001:*') 'the task id encodes source and name plus a per-submission suffix'
    Assert ((Get-MockCount $st 'PUT /api/instances/{id}/clone') -eq 1 -and (Get-MockCount $st 'GET /api/servers/{id}') -eq 1 -and (Get-MockTotal $st) -eq 2) "submission costs two calls: one GET /api/servers/{id} to refresh the source's live network, then the clone PUT (got $(Get-MockTotal $st))"
    $body = (Invoke-Mock '/_ctl/state' 'GET').last_clone_body
    Assert ($body.networkInterfaces[0].network.id -eq 5 -and $body.volumes.Count -eq 1 -and $body.volumes[0].size -eq 60) 'the body echoes the source network and only its disk volume'
    Assert ($body.group.id -eq 1 -and $body.cloud.id -eq 1 -and $body.plan.id -eq 23 -and $body.layout.id -eq 32) 'the body echoes group/cloud/plan/layout'
    Assert ($body.volumes[0].datastoreId -eq 'auto') "datastoreId is passed through as VME sent it ('auto')"
    Assert ($null -eq $body.networkInterfaces[0].PSObject.Properties['macAddress']) 'no macAddress without a preceding delete of the same name'

    Write-Host "`n=== resolution and provisioning ===" -ForegroundColor Cyan
    $t0 = Get-Date
    $t = Wait-Task $p $c1.task_id -TimeoutSec 40
    $secs = ((Get-Date) - $t0).TotalSeconds
    $st = Get-MockStats
    Assert ($null -ne $t -and $t.state -eq 'completed') "the clone task completes (state=$(if ($t) { $t.state }))"
    $cloneId = [string]$t.output.clone_id
    Assert ([long]$cloneId -ge 10000) "output.clone_id is the new instance ($cloneId)"
    Assert ($secs -lt 8) "completed within one poll of VME finishing provisioning ($([Math]::Round($secs,1))s; clone 2s + provision 2s)"
    Assert ((Get-MockCount $st 'GET /api/instances list') -eq 0) 'no full listing was needed to find the clone'
    Assert ((Get-MockCount $st 'GET /api/instances phrase') -ge 1) "it was found with the small phrase query ($(Get-MockCount $st 'GET /api/instances phrase') calls)"
    Assert ((Get-MockCount $st 'GET /api/processes') -eq 0) 'no process-history calls with the deploy gate off'
    $log = Get-LogText $p
    Assert ($log -match "Clone \[RDSH-001\] resolved to instance \[$cloneId\] .*via phrase, cloned from \[19\]") 'the resolution is logged with the method and the server-side lineage'
    Assert ($log -match 'Clone \[RDSH-001\] completed as instance') 'the completion is logged'

    $again = Get-Task $p $c1.task_id
    Assert ($again.state -eq 'completed' -and [string]$again.output.clone_id -eq $cloneId) 're-polling a completed task still returns clone_id'

    $list = Get-GuestList $p
    Assert ($list -contains $cloneId) 'guests/list exposes the clone once reported completed'
    $first = Get-Result (Get-Guest $p $cloneId) 'get clone'
    Assert ($first.name -eq 'RDSH-001' -and $first.is_template -eq $false) 'the clone is reported under its name and is not a template'
    $ip = $null; $deadline = (Get-Date).AddSeconds(20); $states = @()
    while ((Get-Date) -lt $deadline) {
        $g = Get-Result (Get-Guest $p $cloneId) 'poll clone'
        $states += $g.state
        if (@($g.ip_addresses).Count -gt 0 -and $g.state -eq 'powered_on') { $ip = $g.ip; $mac = @($g.mac_addresses); break }
        Start-Sleep -Milliseconds 700
    }
    Assert ($states[0] -eq 'powered_on' -and (Get-LogText $p) -match "instance \[$cloneId\] is running, no usable IP yet; reporting powered_on") "a running clone reports powered_on before its IP exists -- power state decides, not the IP (Q19; first poll: $($states[0]))"
    Assert ($null -ne $ip) "the clone-aware guests/get flow ends at powered_on with a real IP ($ip)"
    $mockMac = (Invoke-Mock '/_ctl/state' 'GET').instances.$cloneId.mac
    Assert ($mac -contains $mockMac) "the clone's MAC is reported ($($mac -join ',')) = the server's ($mockMac)"
    Assert ((Get-LogText $p) -match "instance \[$cloneId\] \(RDSH-001\) is up with IP") 'tracking retires at the first real IP'

    Write-Host "`n=== duplicate names ===" -ForegroundColor Cyan
    Invoke-Mock '/_ctl/instance?id=600&name=RDSH-002&status=running&ip=10.6.0.1&cloned_from=19' | Out-Null
    [void](Get-GuestList $p)
    $c2 = Get-Result (Start-Clone $p '19' 'RDSH-002') 'clone 2'
    Assert ((Get-LogText $p) -match '\[RDSH-002\] already exists as instance\(s\) 600') 'a same-named existing instance is logged and excluded'
    $t2 = Wait-Task $p $c2.task_id -TimeoutSec 40
    Assert ($t2.state -eq 'completed' -and [string]$t2.output.clone_id -ne '600' -and [long]$t2.output.clone_id -ge 10000) "resolution picks the NEW instance, not the preexisting one (got $($t2.output.clone_id))"

    Write-Host "`n=== idempotent re-submit ===" -ForegroundColor Cyan
    $c3a = Get-Result (Start-Clone $p '19' 'RDSH-003') 'clone 3a'
    $c3b = Get-Result (Start-Clone $p '19' 'RDSH-003') 'clone 3b'
    Assert ($c3a.task_id -eq $c3b.task_id) 'a repeated guests/clone for an in-flight name returns the same task'
    $other = Start-Clone $p '20' 'RDSH-003'
    Assert (Test-IsError $other) 'the same name from a different source while in flight is refused'
    $t3 = Wait-Task $p $c3a.task_id -TimeoutSec 40
    Start-Sleep -Seconds 1
    Assert ($t3.state -eq 'completed' -and @(Get-MockInstancesNamed 'RDSH-003').Count -eq 1) 'only ONE VM was created'

    Write-Host "`n=== phrase query not working on the appliance ===" -ForegroundColor Cyan
    Invoke-Mock '/_ctl/set?phrase_enabled=false' | Out-Null
    $c4 = Get-Result (Start-Clone $p '19' 'RDSH-004') 'clone 4'
    $t4 = Wait-Task $p $c4.task_id -TimeoutSec 45
    Assert ($t4.state -eq 'completed') 'the safety-net listing still resolves the clone'
    Assert ((Get-LogText $p) -match 'was found by the full listing but not by the .phrase. query') 'the broken query is called out in the log'
    Invoke-Mock '/_ctl/set?phrase_enabled=true' | Out-Null

    Write-Host "`n=== clone response lost (accepted, then timed out) ===" -ForegroundColor Cyan
    Invoke-Mock '/_ctl/set?clone_hang_once=8' | Out-Null
    $r = Start-Clone $p '19' 'RDSH-006' -TimeoutSec 30
    Assert (-not (Test-IsError $r)) 'a clone PUT without a response is NOT reported as a failure (RAS would retry and duplicate it)'
    Assert ((Get-LogText $p) -match 'RDSH-006\] got no response') 'the lost response is logged'
    $t6 = Wait-Task $p $r.result.task_id -TimeoutSec 40
    Start-Sleep -Seconds 1
    Assert ($t6.state -eq 'completed' -and @(Get-MockInstancesNamed 'RDSH-006').Count -eq 1) 'resolution finds the clone VME did create; one VM, no retry'

    Write-Host "`n=== clone deleted while provisioning ===" -ForegroundColor Cyan
    Invoke-Mock '/_ctl/set?provision_delay=60' | Out-Null
    $c5 = Get-Result (Start-Clone $p '19' 'RDSH-005') 'clone 5'
    $id5 = $null; $deadline = (Get-Date).AddSeconds(20)
    while ((Get-Date) -lt $deadline -and -not $id5) {
        [void](Get-Task $p $c5.task_id)
        if ((Get-LogText $p) -match 'Clone \[RDSH-005\] resolved to instance \[(\d+)\]') { $id5 = $Matches[1] }
        Start-Sleep -Milliseconds 700
    }
    Assert ($null -ne $id5) "the provisioning clone was resolved ($id5)"
    Assert ((Get-GuestList $p) -notcontains $id5) 'a clone not yet reported completed is not listed'
    Invoke-Mock "/_ctl/instance?id=$id5&delete=1" | Out-Null
    $t5 = Get-Task $p $c5.task_id
    Assert ($t5.state -eq 'failed' -and $t5.error.message -match 'disappeared') "a clone that vanishes mid-provision fails the task at once ($($t5.state): $($t5.error.message))"

    Write-Host "`n=== a clone VME created but tasks/get has not resolved yet is not listed ===" -ForegroundColor Cyan
    $c12 = Get-Result (Start-Clone $p '19' 'RDSH-301') 'clone 12'
    Start-Sleep -Seconds 3   # clone_delay 2s: the instance exists on VME; no tasks/get has run for it
    $list12 = Get-GuestList $p
    $id12 = @(Get-MockInstancesNamed 'RDSH-301')
    Assert ($id12.Count -eq 1 -and $list12 -notcontains $id12[0]) "the listing that first sees the new instance does not expose it to RAS before tasks/get resolves it ($($id12 -join ','))"
    # id above every id known at submit, same name, but VME says it was cloned from 20
    Invoke-Mock '/_ctl/instance?id=99001&name=RDSH-301&status=running&ip=10.6.0.9&cloned_from=20' | Out-Null
    Assert ((Get-GuestList $p) -contains '99001') 'a same-named new instance cloned from a different source is still listed (lineage check)'
    Invoke-Mock '/_ctl/instance?id=99001&delete=1' | Out-Null
    [void](Get-GuestList $p)
    # cleanup: resolve it, then remove it -- the task fails and frees its clone slot
    $deadline = (Get-Date).AddSeconds(15)
    while ((Get-Date) -lt $deadline -and (Get-LogText $p) -notmatch 'Clone \[RDSH-301\] resolved') { [void](Get-Task $p $c12.task_id); Start-Sleep -Milliseconds 500 }
    Invoke-Mock "/_ctl/instance?id=$($id12[0])&delete=1" | Out-Null
    $t12 = Get-Task $p $c12.task_id
    Assert ($t12.state -eq 'failed') "cleanup: the task fails once its instance is gone ($($t12.state))"
    Invoke-Mock '/_ctl/set?provision_delay=2' | Out-Null

    Write-Host "`n=== no shadow state: deleting a just-completed clone stops tasks/get answering for it ===" -ForegroundColor Cyan
    # Admin cancels a slow/stuck clone from the RAS console by deleting the guest it
    # produced. RAS may still be re-polling the ORIGINAL clone task (its own retry/sanity
    # cadence, independent of the delete). Before the fix, Handle-TaskInfo answered from
    # CompletedCloneTaskOutputs (a cache separate from CloneState, kept ~600s) without
    # checking the guest still exists -- so a repeated tasks/get kept reporting the
    # deleted clone as a valid, completed guest: VME shows the clone gone,
    # RAS still told it exists.
    $again1 = Get-Task $p $c1.task_id
    Assert ($again1.state -eq 'completed' -and [string]$again1.output.clone_id -eq $cloneId) 'sanity: before the delete, a repeated tasks/get still answers from the cache'
    Get-Result (Send-Rpc $p "{`"method`":`"guests/control`",`"params`":{`"id`":`"$cloneId`",`"control`":`"delete`"}}") 'cancel via delete' | Out-Null
    $afterDelete = Get-Task $p $c1.task_id
    $hasOutput = $afterDelete.PSObject.Properties.Name -contains 'output'
    $afterCloneId = if ($hasOutput) { [string]$afterDelete.output.clone_id } else { $null }
    Assert ($afterDelete.state -ne 'completed' -or $afterCloneId -ne $cloneId) "deleting the clone immediately stops tasks/get reporting it as a valid, completed guest (state=$($afterDelete.state)$(if ($afterCloneId) { ", clone_id=$afterCloneId" }))"
    Assert ($afterDelete.state -eq 'failed' -and $afterDelete.error.message -match 'no longer tracked') "the original clone task now fails cleanly instead of resurrecting the deleted guest ($($afterDelete.error.message))"

    Write-Host "`n=== MAC preservation across a delete + recreate ===" -ForegroundColor Cyan
    $oldMac = (Invoke-Mock '/_ctl/state' 'GET').instances.$cloneId.mac
    $deadline = (Get-Date).AddSeconds(15)
    while ((Get-Date) -lt $deadline -and (Get-GuestList $p) -contains $cloneId) { Start-Sleep -Milliseconds 500 }
    $c7 = Get-Result (Start-Clone $p '19' 'RDSH-001') 'recreate'
    Assert ($c7.task_id -ne $c1.task_id) 'a recreate under the same name gets a NEW task id'
    $first7 = Get-Task $p $c7.task_id
    Assert ($first7.state -eq 'running') "the recreate's first tasks/get is running -- not answered with the deleted instance's id (got $($first7.state))"
    $body = (Invoke-Mock '/_ctl/state' 'GET').last_clone_body
    Assert ($body.networkInterfaces[0].macAddress -eq $oldMac) "the recreate requests the deleted guest's MAC in networkInterfaces[0].macAddress ($oldMac)"
    $t7 = Wait-Task $p $c7.task_id -TimeoutSec 40
    Assert ((Get-LogText $p) -match "got the requested MAC \[$oldMac\]") 'after completion the new MAC is checked and reported as honoured'

    Invoke-Mock '/_ctl/set?clone_honours_mac=false' | Out-Null
    $id7 = [string]$t7.output.clone_id
    Get-Result (Send-Rpc $p "{`"method`":`"guests/control`",`"params`":{`"id`":`"$id7`",`"control`":`"delete`"}}") 'delete 2' | Out-Null
    $deadline = (Get-Date).AddSeconds(15)
    while ((Get-Date) -lt $deadline -and (Get-GuestList $p) -contains $id7) { Start-Sleep -Milliseconds 500 }
    $c8 = Get-Result (Start-Clone $p '19' 'RDSH-001') 'recreate 2'
    [void](Wait-Task $p $c8.task_id -TimeoutSec 40)
    Assert ((Get-LogText $p) -match 'VME ignored networkInterfaces\[\]\.macAddress on clone') 'if VME assigns a different MAC, a W line says so'

    Write-Host "`n=== network/VLAN: the SERVER object is authoritative, not the (possibly stale) instance field ===" -ForegroundColor Cyan
    Invoke-Mock '/_ctl/instance?id=19&net_id=7&net_name=VLAN_STALE' | Out-Null   # instance-level only -- Morpheus does not keep this synced
    [void](Get-GuestList $p)
    Reset-MockStats
    $c9 = Get-Result (Start-Clone $p '19' 'RDSH-101') 'clone 9'
    $stC9 = Get-MockStats
    Assert ((Get-MockCount $stC9 'GET /api/servers/{id}') -eq 1) "clone submit refreshes the source's server object first, before building the body ($(Get-MockTotal $stC9) calls total)"
    $body = (Invoke-Mock '/_ctl/state' 'GET').last_clone_body
    Assert ($body.networkInterfaces[0].network.id -eq 5) "a stale instance-level-only network change is ignored -- the clone still uses the source's server-level network ($($body.networkInterfaces[0].network.id))"
    Assert ((Get-LogText $p) -match "instance-level network \[7\] \(VLAN_STALE\) differs from its server-level network \[5\]") 'the instance-vs-server divergence is logged so it is not silently wrong'
    [void](Wait-Task $p $c9.task_id -TimeoutSec 40)

    Invoke-Mock '/_ctl/instance?id=19&srv_net_id=9&srv_net_name=VLAN200' | Out-Null   # the REAL (server-level) VLAN changes
    $c9b = Get-Result (Start-Clone $p '19' 'RDSH-102') 'clone 9b'
    $body = (Invoke-Mock '/_ctl/state' 'GET').last_clone_body
    Assert ($body.networkInterfaces[0].network.id -eq 9) "once the server-level network actually changes, the clone follows it, still ignoring the stale instance-level field ($($body.networkInterfaces[0].network.id))"
    $t9b = Wait-Task $p $c9b.task_id -TimeoutSec 40
    Assert ($t9b.state -eq 'completed') 'the clone onto the new network still completes'
    Assert ((Get-LogText $p) -match 'Network preservation: clone \[\d+\] \(RDSH-102\) landed on the requested network \[9\]') 'once the clone has a network reported, the outcome is logged as honoured'

    Invoke-Mock '/_ctl/set?clone_honours_network=false' | Out-Null
    $c10 = Get-Result (Start-Clone $p '19' 'RDSH-103') 'clone 10'
    [void](Wait-Task $p $c10.task_id -TimeoutSec 40)
    Assert ((Get-LogText $p) -match 'VME ignored networkInterfaces\[\]\.network\.id on clone') "if VME keeps the source's network/VLAN instead of the requested one, a W line says so"
    Invoke-Mock '/_ctl/set?clone_honours_network=true' | Out-Null

    Write-Host "`n=== giving up on a stuck clone cancels its VME process ===" -ForegroundColor Cyan
    # Backdated started_epoch instead of a real wait: same technique as the de-DE/2.x
    # restart test below, just for elapsed-time logic instead of culture. Avoids waiting
    # out the real provision_timeout_seconds (min 120s) to exercise the give-up path.
    Invoke-Mock '/_ctl/instance?id=8800&name=RDSH-STUCK&status=provisioning&cloned_from=19' | Out-Null
    $stuckStarted = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds() - 2000   # exceeds the default provision_timeout_seconds (1740)
    $dir3 = New-ProviderDir -Name 'u3-giveup' -Settings @{
        inventory  = @{ list_min_refresh_seconds = 0 }
        connection = @{ request_timeout_seconds = 5; retry_delay_milliseconds = 10 }
        logging    = @{ log_level = 5 }
    }
    @"
{
  "8800": { "type": "clone", "task_id": "clonebyname:19:RDSH-STUCK:aaaa", "source_id": "19", "clone_id": "8800",
            "name": "RDSH-STUCK", "started_epoch": $stuckStarted, "provision_confirmed": false,
            "reported_completed_epoch": null, "preexisting_ids": [], "max_known_id": 0, "submit_uncertain": false }
}
"@ | Set-Content -LiteralPath (Join-Path $dir3 'VME-RAS-CloneState.json') -Encoding utf8NoBOM
    $p3 = Start-Provider -Dir $dir3
    Connect-Provider $p3 | Out-Null
    $tStuck = Get-Task $p3 'clonebyname:19:RDSH-STUCK:aaaa'
    Assert ($tStuck.state -eq 'failed' -and $tStuck.error.message -match 'still \[provisioning\] after \d+s') "a clone stuck past provision_timeout_seconds fails with a reason ($($tStuck.state): $($tStuck.error.message))"
    $stuckMockState = Invoke-Mock '/_ctl/state' 'GET'
    Assert ($stuckMockState.cancelled_processes -contains 13800) 'giving up also cancels the stuck clone''s VME process (id = 5000 + instance id)'
    Assert ((Get-LogText $p3) -match 'Process \[13800\] cancelled on VME') 'the cancellation is logged'
    Stop-Provider $p3
    Invoke-Mock '/_ctl/instance?id=8800&delete=1' | Out-Null

    Write-Host "`n=== babysit window runs from the completion report, not the submit ===" -ForegroundColor Cyan
    # Submitted 8000s ago (past the 7200s window if measured from the submit), reported
    # completed 100s ago, running without an IP: still inside its window.
    Invoke-Mock '/_ctl/instance?id=8900&name=RDSH-LATE&status=running&cloned_from=19' | Out-Null
    $now = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
    $dir5 = New-ProviderDir -Name 'u3-babysit' -Settings @{
        inventory  = @{ list_min_refresh_seconds = 0 }
        connection = @{ request_timeout_seconds = 5; retry_delay_milliseconds = 10 }
        logging    = @{ log_level = 5 }
    }
    @"
{
  "8900": { "type": "clone", "task_id": "clonebyname:19:RDSH-LATE:bbbb", "source_id": "19", "clone_id": "8900",
            "name": "RDSH-LATE", "started_epoch": $($now - 8000), "provision_confirmed": true,
            "reported_completed_epoch": $($now - 100), "preexisting_ids": [], "max_known_id": 0, "submit_uncertain": false }
}
"@ | Set-Content -LiteralPath (Join-Path $dir5 'VME-RAS-CloneState.json') -Encoding utf8NoBOM
    $p5 = Start-Provider -Dir $dir5
    Connect-Provider $p5 | Out-Null
    $g5 = Get-Result (Get-Guest $p5 '8900') 'get 8900'
    $log5 = Get-LogText $p5
    Assert ($log5 -notmatch 'babysit window elapsed for instance \[8900\]') 'a clone reported 100s ago is still babysat although it was submitted 8000s ago'
    Assert ($g5.state -eq 'powered_on' -and $log5 -match 'instance \[8900\] is running, no usable IP yet') "it is still in the clone-aware flow ($($g5.state))"
    Stop-Provider $p5
    Invoke-Mock '/_ctl/instance?id=8900&delete=1' | Out-Null

    Write-Host "`n=== orphan audit: a clone retired at its first IP is not 'never polled' ===" -ForegroundColor Cyan
    # Retiring tracking at the first IP also dropped the clone's
    # last-poll stamp, so the next audit fell back to the process start time and flagged a
    # healthy, just-polled clone as an ORPHAN CANDIDATE. The audit's minimum stale window is
    # 60s, so this needs one real wait: process age > 60s, clone polled < 60s ago.
    $dir4 = New-ProviderDir -Name 'u3-orphan' -Settings @{
        inventory        = @{ list_min_refresh_seconds = 0 }
        cloning          = @{ tracked_refresh_interval_seconds = 1 }
        orphan_detection = @{ check_interval_seconds = 30; stale_poll_after_seconds = 60 }
        connection       = @{ request_timeout_seconds = 5; retry_delay_milliseconds = 10 }
        logging          = @{ log_level = 5 }
    }
    $p4 = Start-Provider -Dir $dir4
    $p4Started = Get-Date
    Connect-Provider $p4 | Out-Null
    [void](Get-GuestList $p4)   # first audit runs here (nothing is stale yet)
    [void](Get-Result (Send-Rpc $p4 '{"method":"guests/convert","params":{"id":"19","is_template":true}}') 'convert on p4')
    [void](Get-GuestList $p4)
    $c11 = Get-Result (Start-Clone $p4 '19' 'RDSH-201') 'clone 11'
    $t11 = Wait-Task $p4 $c11.task_id -TimeoutSec 40
    $id11 = [string]$t11.output.clone_id
    $deadline = (Get-Date).AddSeconds(20)
    while ((Get-Date) -lt $deadline -and (Get-LogText $p4) -notmatch "instance \[$id11\] \(RDSH-201\) is up with IP") {
        [void](Get-Guest $p4 $id11); Start-Sleep -Milliseconds 700
    }
    Assert ((Get-LogText $p4) -match "instance \[$id11\] \(RDSH-201\) is up with IP") "clone [$id11] retired at its first IP"
    $wait = 62 - ((Get-Date) - $p4Started).TotalSeconds
    if ($wait -gt 0) { Start-Sleep -Seconds ([Math]::Ceiling($wait)) }
    [void](Get-GuestList $p4)   # second audit: process is > 60s old, the clone was polled < 60s ago
    $log4 = Get-LogText $p4
    Assert ($log4 -match 'ORPHAN CANDIDATE: .*never since this provider process started') 'control: the audit did run and flags clones this process never polled'
    Assert ($log4 -notmatch "ORPHAN CANDIDATE: instance \[$id11\]") "a clone polled to its first IP and retired is not flagged as an orphan"
    Stop-Provider $p4

    Write-Host "`n=== log hygiene at log_level 5 ===" -ForegroundColor Cyan
    $log = Get-LogText $p
    Assert ($log -notmatch '"password":"password"' -and $log -notmatch 'mock-token-abc' -and $log -notmatch 'mock-refresh-xyz') 'no password, token or refresh token in the verbose log'
    Assert ($log -match '"password":"\*\*\*"') 'the connect request is logged with the password masked'
    Assert (-not (Test-Path (Join-Path $dir 'VME-RAS-CloneLedger.json'))) 'no clone ledger file (lineage is server-side now)'
    $state = Get-Content (Join-Path $dir 'VME-RAS-CloneState.json') -Raw
    Assert ($state -notmatch 'pending:') 'no unresolved clone is left in the state file (reported clones stay until RAS polls them to an IP)'
    Assert ($state -match '"started_epoch": \d+') 'clone-state timestamps are epoch seconds, not culture-sensitive strings'
}
catch { Assert $false "suite aborted: $($_.Exception.Message) at $($_.InvocationInfo.PositionMessage)" }
finally { Complete-Suite 'vme-u3' }
