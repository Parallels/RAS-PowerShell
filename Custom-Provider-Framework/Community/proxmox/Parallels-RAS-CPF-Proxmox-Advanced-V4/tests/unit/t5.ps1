$ErrorActionPreference = 'Stop'
$pass = 0; $fail = 0
function Assert { param([bool]$Cond, [string]$Msg) if ($Cond) { $script:pass++; Write-Host "PASS: $Msg" -Fore Green } else { $script:fail++; Write-Host "FAIL: $Msg" -Fore Red } }

. "$PSScriptRoot/funcs2.ps1"
$script:LogPath = "$PSScriptRoot/t5.log"
$script:CloneStatePath = "$PSScriptRoot/t5-state.json"
Remove-Item $script:LogPath, $script:CloneStatePath -ErrorAction SilentlyContinue
$script:TrackedCloneVmIds = [System.Collections.Generic.HashSet[string]]::new()
$script:MaxConcurrentCloneOperations = 2
$script:PipelinedCloneCompletionSeconds = 10

function Back-Date([string]$taskId, [string]$vmid, [int]$sec) {
    $newTs = ([DateTime]::UtcNow.AddSeconds(-$sec)).ToString('o')
    # Handle-TaskInfo reads $script:TaskContext FIRST (falling back to the file
    # only if not found in memory) -- so the in-memory copy is what actually
    # needs backdating here, matching how Handle-GuestClone->Handle-TaskInfo
    # really flows (both copies normally share the same clone_started_at,
    # written once at clone time).
    $script:TaskContext[$taskId].clone_started_at = $newTs
    $s = Get-CloneState
    $s[$vmid].clone_started_at = $newTs
    Save-CloneState -State $s
}

$script:apiCalls = New-Object System.Collections.ArrayList
$script:nextId = 500
function Invoke-ProxmoxApi {
    param([string]$Method, [string]$Path, $Body, [int]$TimeoutSec = 0)
    [void]$script:apiCalls.Add("$Method $Path")
    if ($Path -match '/cluster/resources') {
        return @{ data = @([pscustomobject]@{ type='qemu'; vmid=999; name='tmpl'; node='n1'; status='stopped'; template=1 }) }
    }
    if ($Path -match '/qemu/(\d+)/clone') {
        $newid = [string]$Body.newid
        return @{ data = "UPID:n1:clone$newid" }
    }
    if ($Path -match '/tasks/(.+)/status') {
        return @{ data = [pscustomobject]@{ status = 'running'; exitstatus = $null } }
    }
    throw "unmocked: $Method $Path"
}
function Get-ProxmoxNextVmId { return [string]$script:nextId }

$script:nextId = 500; $r1 = Handle-GuestClone -Params ([pscustomobject]@{ id='999'; name='a' })
Back-Date $r1.result.task_id '500' 12
$t1 = Handle-TaskInfo -Params ([pscustomobject]@{ id = $r1.result.task_id })
Assert ($t1.result.state -eq 'completed') "clone 1/2: pipelines through (room for another)"

$script:nextId = 501; $r2 = Handle-GuestClone -Params ([pscustomobject]@{ id='999'; name='b' })
Back-Date $r2.result.task_id '501' 11
$t2 = Handle-TaskInfo -Params ([pscustomobject]@{ id = $r2.result.task_id })
# Regression check for Handle-TaskInfo's off-by-one guard ("-ge", not "-gt"):
# active count is now 2 (500's entry is still present -- only
# Get-RasGuestObjectForCloneAwareFlow, on real readiness, ever removes it) which
# EQUALS the limit, so per the code's own documented strict-less-than semantics
# this must still hold, not pipeline through.
Assert ($t2.result.state -eq 'running') "clone 2 (active count would equal the limit) is correctly held, not pipelined"

$script:nextId = 502; $r3 = Handle-GuestClone -Params ([pscustomobject]@{ id='999'; name='c' })
Back-Date $r3.result.task_id '502' 11
$t3 = Handle-TaskInfo -Params ([pscustomobject]@{ id = $r3.result.task_id })
Assert ($t3.result.state -eq 'running') "Regression check: clone 3 (would exceed cap of 2) is also correctly held"

# Confirm the system isn't stuck: with 3 entries persisted (500/501/502), 501
# alone freeing 500's slot isn't enough -- active count is still 501+502 = 2 =
# limit. Free 502 as well (simulating it also reaching ready) so only 501
# remains active; it should then pipeline on its next poll.
function Invoke-ProxmoxApi {
    param([string]$Method, [string]$Path, $Body, [int]$TimeoutSec = 0)
    if ($Path -match '/cluster/resources') {
        return @{ data = @(
            [pscustomobject]@{ type='qemu'; vmid=500; name='w-a'; node='n1'; status='running'; template=0 }
            [pscustomobject]@{ type='qemu'; vmid=502; name='w-c'; node='n1'; status='running'; template=0 }
        ) }
    }
    if ($Path -match '/qemu/(500|502)/agent/network-get-interfaces') {
        $vm = $Matches[1]
        return @{ data = @{ result = @([pscustomobject]@{ 'hardware-address'="AA:BB:CC:00:00:0$vm"; 'ip-addresses'=@([pscustomobject]@{'ip-address-type'='ipv4';'ip-address'="10.0.0.$vm"}) }) } }
    }
    if ($Path -match '/tasks/(.+)/status') { return @{ data = [pscustomobject]@{ status='running'; exitstatus=$null } } }
    throw "unmocked: $Method $Path"
}
Reset-ProxmoxClusterCache
[void](Get-RasGuestObjectForCloneAwareFlow -VmId '500')   # confirms 500 ready, frees its slot
[void](Get-RasGuestObjectForCloneAwareFlow -VmId '502')   # confirms 502 ready, frees its slot
Assert ((Get-ActiveCloneCount) -eq 1) "active count correctly drops to 1 (only 501 left) once 500 and 502 are both confirmed ready"
$t2b = Handle-TaskInfo -Params ([pscustomobject]@{ id = $r2.result.task_id })
Assert ($t2b.result.state -eq 'completed') "once enough slots free, the next poll correctly pipelines the held clone through -- not permanently stuck"

Write-Host "`n=== Summary: $pass passed, $fail failed ===" -Fore $(if ($fail -eq 0) { 'Green' } else { 'Red' })
if ($fail -gt 0) { exit 1 }
