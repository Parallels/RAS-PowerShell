# vme-u6: clone slots (cloning.max_concurrent_clone_operations) -- a clone the provider
# gives up on keeps its slot until VME confirms it ended, so a stuck appliance can hold
# at most N stuck clones however many times RAS moves on to the next guest.
#
# Give-ups are reached with backdated clone-state entries (same technique as vme-u3's
# give-up test) instead of waiting out the real minimum timeouts.
. (Join-Path $PSScriptRoot 'lib.ps1')

function Start-Clone {
    param($P, [string]$Source, [string]$Name)
    return (Send-Rpc $P "{`"method`":`"guests/clone`",`"params`":{`"id`":`"$Source`",`"name`":`"$Name`"}}")
}
function Write-CloneState {
    param([string]$Dir, [string]$Json)
    Set-Content -LiteralPath (Join-Path $Dir 'VME-RAS-CloneState.json') -Value $Json -Encoding utf8NoBOM
}
function Read-StuckFile {
    param([string]$Dir)
    $f = Join-Path $Dir 'VME-RAS-StuckClones.json'
    if (-not (Test-Path $f)) { return @() }
    return @((Get-Content $f -Raw | ConvertFrom-Json).entries)
}
function Get-ErrorText { param($R) if (Test-IsError $R) { return $R.error.message } return '' }

$settings = @{
    inventory  = @{ list_min_refresh_seconds = 0 }
    cloning    = @{ max_concurrent_clone_operations = 2; stuck_clone_recheck_seconds = 10 }
    delete     = @{ probe_interval_seconds = 1 }
    connection = @{ request_timeout_seconds = 5; retry_delay_milliseconds = 10 }
    logging    = @{ log_level = 5 }
}
$old = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds() - 2000   # past resolve (900) and provision (1740) timeouts

Start-VmeMock -Count 10
Invoke-Mock '/_ctl/set?clone_never_appears=true' | Out-Null
# Two clones VME accepted earlier and never produced an instance for (VME's stuck-provision case).
Invoke-Mock '/_ctl/ghost?add=1&src=19&name=VDI-001' | Out-Null
Invoke-Mock '/_ctl/ghost?add=1&src=19&name=VDI-002' | Out-Null

$dir = New-ProviderDir -Name 'u6' -Settings $settings
Write-CloneState $dir @"
{
  "pending:VDI-001": { "type": "clone", "task_id": "clonebyname:19:VDI-001:aaaa", "source_id": "19", "clone_id": "", "name": "VDI-001",
                       "started_epoch": $old, "provision_confirmed": false, "reported_completed_epoch": null, "preexisting_ids": [], "max_known_id": 200 },
  "pending:VDI-002": { "type": "clone", "task_id": "clonebyname:19:VDI-002:bbbb", "source_id": "19", "clone_id": "", "name": "VDI-002",
                       "started_epoch": $($old + 60), "provision_confirmed": false, "reported_completed_epoch": null, "preexisting_ids": [], "max_known_id": 200 }
}
"@
$p = Start-Provider -Dir $dir
try {
    Connect-Provider $p | Out-Null
    [void](Get-GuestList $p)
    [void](Get-Result (Send-Rpc $p '{"method":"guests/convert","params":{"id":"19","is_template":true}}') 'convert')

    Write-Host "`n=== a given-up clone keeps its slot ===" -ForegroundColor Cyan
    $t1 = Get-Task $p 'clonebyname:19:VDI-001:aaaa'
    $t2 = Get-Task $p 'clonebyname:19:VDI-002:bbbb'
    Assert ($t1.state -eq 'failed' -and $t1.error.message -match 'never appeared') "RAS still gets 'failed' for a clone that never appeared ($($t1.error.message))"
    Assert ($t2.state -eq 'failed') 'the second one fails too'
    Assert ((Get-LogText $p) -match 'Clone \[VDI-001\] \(never appeared on VME\) given up .* keeps one of the 2 clone slot') 'the give-up is logged as keeping a slot'
    $stuck = @(Read-StuckFile $dir)
    Assert ($stuck.Count -eq 2 -and ($stuck.name -contains 'VDI-001') -and ($stuck.name -contains 'VDI-002')) "both are persisted in VME-RAS-StuckClones.json ($($stuck.Count))"
    $again = Get-Task $p 'clonebyname:19:VDI-001:aaaa'
    Assert ($again.state -eq 'failed' -and $again.error.message -match 'still holds a clone slot') 'a repeated tasks/get for a given-up clone keeps answering failed, with the reason'

    Write-Host "`n=== the limit refuses the next clone without calling VME's clone endpoint ===" -ForegroundColor Cyan
    Reset-MockStats
    $r3 = Start-Clone $p '19' 'VDI-003'
    $st = Get-MockStats
    Assert ((Get-ErrorText $r3) -match 'Clone limit reached: 2 of 2') "guests/clone is refused at once ($(Get-ErrorText $r3))"
    Assert ((Get-ErrorText $r3) -match 'VDI-001' -and (Get-ErrorText $r3) -match 'VDI-002' -and (Get-ErrorText $r3) -match 'never appeared') 'the refusal names the clones holding the slots'
    Assert ((Get-MockCount $st 'PUT /api/instances/{id}/clone') -eq 0) 'nothing was submitted to VME'
    Assert ((Get-MockCount $st 'GET /api/processes') -eq 1) "the blocked request re-checked the template's process history once ($(Get-MockCount $st 'GET /api/processes'))"
    Reset-MockStats
    $r4 = Start-Clone $p '19' 'VDI-004'
    $st = Get-MockStats
    Assert ((Test-IsError $r4) -and (Get-MockTotal $st) -eq 0) "a refusal inside stuck_clone_recheck_seconds costs no VME call at all ($(Get-MockTotal $st))"

    Write-Host "`n=== VME ending a stuck clone frees its slot ===" -ForegroundColor Cyan
    Invoke-Mock '/_ctl/ghost?name=VDI-001&status=failed' | Out-Null
    Start-Sleep -Seconds 11
    $r5 = Start-Clone $p '19' 'VDI-005'
    Assert (-not (Test-IsError $r5)) "once VME's process for VDI-001 ended, the next clone is submitted ($(Get-ErrorText $r5))"
    Assert ((Get-LogText $p) -match 'Clone slot released: given-up clone \[VDI-001\] -- its provision process \[\d+\] on template \[19\] ended with status \[failed\]') 'the release is logged with the process outcome'
    $stuck = @(Read-StuckFile $dir)
    Assert ($stuck.Count -eq 1 -and $stuck[0].name -eq 'VDI-002') 'VDI-002 still holds its slot'

    Write-Host "`n=== in-flight clones count too; a stuck name is not reused ===" -ForegroundColor Cyan
    $r6 = Start-Clone $p '19' 'VDI-006'
    Assert ((Get-ErrorText $r6) -match '2 of 2' -and (Get-ErrorText $r6) -match 'VDI-005\] \(not found yet, in flight') "a clone RAS is still waiting for holds a slot ($(Get-ErrorText $r6))"
    $r2 = Start-Clone $p '19' 'VDI-002'
    Assert ((Get-ErrorText $r2) -match 'earlier clone named \[VDI-002\].*given up') "re-cloning the name of a never-appeared stuck clone is refused ($(Get-ErrorText $r2))"
    Stop-Provider $p

    Write-Host "`n=== slots survive a provider restart ===" -ForegroundColor Cyan
    $p = Start-Provider -Dir $dir
    Connect-Provider $p | Out-Null
    [void](Get-GuestList $p)
    Assert ((Get-LogText $p) -match 'given-up clones holding a slot=1') 'the startup banner reports the persisted slot'
    $r7 = Start-Clone $p '19' 'VDI-007'
    Assert ((Get-ErrorText $r7) -match 'VDI-002\] \(never appeared, given up') "after a restart the given-up clone still blocks ($(Get-ErrorText $r7))"

    Write-Host "`n=== process displayName is not the clone name: freed by count, conservatively ===" -ForegroundColor Cyan
    Invoke-Mock '/_ctl/set?ghost_display_name=false' | Out-Null
    Invoke-Mock '/_ctl/ghost?name=VDI-002&status=complete' | Out-Null
    Start-Sleep -Seconds 11
    $r8 = Start-Clone $p '19' 'VDI-008'
    Assert (Test-IsError $r8) 'with VDI-005 still unfinished on the same template, the unmatched history cannot prove which one ended -- slot kept'
    Invoke-Mock '/_ctl/ghost?name=VDI-005&status=complete' | Out-Null
    Start-Sleep -Seconds 11
    $r9 = Start-Clone $p '19' 'VDI-009'
    Assert (-not (Test-IsError $r9)) "once the template's history has no unfinished provision process left, the slot is freed ($(Get-ErrorText $r9))"
    Assert ((Get-LogText $p) -match "Clone slot released: given-up clone \[VDI-002\] -- template \[19\]'s process history shows 0 unfinished provision process\(es\) for 1 never-appeared") 'the count-based release is logged'
    Stop-Provider $p

    Write-Host "`n=== a clone stuck provisioning: freed when VME finishes it, or once its delete is confirmed ===" -ForegroundColor Cyan
    Invoke-Mock '/_ctl/reset?count=10' | Out-Null
    Invoke-Mock '/_ctl/instance?id=8801&name=VDI-S1&status=provisioning&busy=1&cloned_from=19' | Out-Null
    Invoke-Mock '/_ctl/instance?id=8802&name=VDI-S2&status=provisioning&busy=1&cloned_from=19' | Out-Null
    $dirB = New-ProviderDir -Name 'u6-provisioning' -Settings $settings
    Write-CloneState $dirB @"
{
  "8801": { "type": "clone", "task_id": "clonebyname:19:VDI-S1:cccc", "source_id": "19", "clone_id": "8801", "name": "VDI-S1",
            "started_epoch": $old, "provision_confirmed": false, "reported_completed_epoch": null, "preexisting_ids": [], "max_known_id": 200 },
  "8802": { "type": "clone", "task_id": "clonebyname:19:VDI-S2:dddd", "source_id": "19", "clone_id": "8802", "name": "VDI-S2",
            "started_epoch": $old, "provision_confirmed": false, "reported_completed_epoch": null, "preexisting_ids": [], "max_known_id": 200 }
}
"@
    $pb = Start-Provider -Dir $dirB
    Connect-Provider $pb | Out-Null
    [void](Get-GuestList $pb)
    [void](Get-Result (Send-Rpc $pb '{"method":"guests/convert","params":{"id":"19","is_template":true}}') 'convert B')
    $ts1 = Get-Task $pb 'clonebyname:19:VDI-S1:cccc'
    [void](Get-Task $pb 'clonebyname:19:VDI-S2:dddd')
    Assert ($ts1.state -eq 'failed' -and $ts1.error.message -match 'still \[provisioning\]') 'a clone stuck provisioning past the timeout still fails to RAS'
    Assert ((Get-LogText $pb) -match 'Clone \[VDI-S1\] \(instance \[8801\], status \[provisioning\]\) given up') 'and keeps its slot, with its instance id'
    $rb1 = Start-Clone $pb '19' 'VDI-B1'
    Assert ((Get-ErrorText $rb1) -match 'instance 8801 provisioning, given up') "both stuck clones block ($(Get-ErrorText $rb1))"

    [void](Get-GuestList $pb)
    Get-Result (Send-Rpc $pb '{"method":"guests/control","params":{"id":"8802","control":"delete"}}') 'delete 8802' | Out-Null
    Assert ((Get-LogText $pb) -notmatch 'Clone slot released: given-up clone \[VDI-S2\]') 'a delete that is only submitted does not free the slot yet'
    $deadline = (Get-Date).AddSeconds(15)
    while ((Get-Date) -lt $deadline -and (Get-LogText $pb) -notmatch 'Clone slot released: given-up clone \[VDI-S2\]') { [void](Get-GuestList $pb); Start-Sleep -Milliseconds 700 }
    Assert ((Get-LogText $pb) -match 'Clone slot released: given-up clone \[VDI-S2\] \(instance \[8802\]\) -- VME confirms it gone') 'the delete probe confirming the 404 frees the slot'
    $rb2 = Start-Clone $pb '19' 'VDI-B2'
    Assert (-not (Test-IsError $rb2)) "with one slot free the next clone goes through ($(Get-ErrorText $rb2))"
    $rb3 = Start-Clone $pb '19' 'VDI-B3'
    Assert ((Get-ErrorText $rb3) -match '2 of 2') 'the in-flight clone and the remaining stuck one fill both slots again'
    $tb2 = Wait-Task $pb $rb2.result.task_id -TimeoutSec 40
    Assert ($tb2.state -eq 'completed') 'the clone completes normally'
    $rb4 = Start-Clone $pb '19' 'VDI-B4'
    Assert (-not (Test-IsError $rb4)) "its slot is free the moment tasks/get reports it completed ($(Get-ErrorText $rb4))"

    Invoke-Mock '/_ctl/instance?id=8801&status=running&busy=0&ip=10.8.0.1' | Out-Null
    Start-Sleep -Seconds 11
    $rb5 = Start-Clone $pb '19' 'VDI-B5'
    Assert (-not (Test-IsError $rb5)) "VME finishing the stuck clone frees its slot ($(Get-ErrorText $rb5))"
    Assert ((Get-LogText $pb) -match 'Clone slot released: given-up clone \[VDI-S1\] \(instance \[8801\]\) -- VME finished provisioning it \(status \[running\]\)') 'the release is logged with the final status'
    Assert (@(Read-StuckFile $dirB).Count -eq 0) 'the stuck-clone file is empty again'
    Stop-Provider $pb
}
catch { Assert $false "suite aborted: $($_.Exception.Message) at $($_.InvocationInfo.PositionMessage)" }
finally { Complete-Suite 'vme-u6' }
