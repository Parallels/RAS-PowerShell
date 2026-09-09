$ErrorActionPreference = 'Stop'
$pass = 0
$fail = 0
function Assert {
    param([bool]$Cond, [string]$Msg)
    if ($Cond) { $script:pass++; Write-Host "PASS: $Msg" -ForegroundColor Green }
    else { $script:fail++; Write-Host "FAIL: $Msg" -ForegroundColor Red }
}

# ============================================================
# TESTS for a live regression where 2 of 3 clones reported
# "Failed to create" by RAS despite succeeding on Proxmox:
#   Bug A: name substitution was gated on a specific placeholder REGEX
#          ("VM <id>", a space) that doesn't match every environment's
#          actual transient name ("VM-<id>", a hyphen -- which turned out to
#          be this provider's own missing-name fallback, not even a Proxmox
#          value).
#   Bug B: a race between Handle-TaskInfo's tasks/get poll and
#          Get-RasGuestObjectForCloneAwareFlow's guests/get-driven completion
#          detection -- whichever clears tracking first leaves the other's
#          poll with nothing to find, and a tasks/get poll that arrives after
#          tracking is already gone fell through to a generic
#          "completed, no output" response with no clone_id.
#   Bug C: a clone ends up carrying BOTH rasTemplate<sourceId> (wrongly
#          inherited from the source's config at clone time) and
#          rasClone<sourceId> (correct).
# ============================================================
. "$PSScriptRoot/funcs2.ps1"

$script:LogPath        = "$PSScriptRoot/t7.log"
$script:CloneStatePath = "$PSScriptRoot/t7-state.json"
Remove-Item $script:LogPath, $script:CloneStatePath -ErrorAction SilentlyContinue
$script:TrackedCloneVmIds = [System.Collections.Generic.HashSet[string]]::new()
$script:ClusterResourcesCache = $null
$script:ClusterResourcesCachedAt = $null
$script:RecentlyDeletedIds = @{}
$script:RecentlyControlledIds = @{}
$script:AgentFailureTracker = @{}
$script:LastKnownNetworkData = @{}
$script:SweepLastCheckedAt = @{}
$script:RasTemplateTagAttempted = [System.Collections.Generic.HashSet[string]]::new()
$script:CloneStateMemory = $null
$script:TaskContext = @{}
$script:CompletedCloneTaskOutputs = @{}

$script:apiCalls = New-Object System.Collections.ArrayList
$script:putConfigCalls = New-Object System.Collections.ArrayList
$script:nextId = 700
$script:cloneJobs = @{}
$script:vmState = @{}
$script:vmName  = @{}
$script:vmVisible = @{}
$script:vmTags = @{}

function Invoke-ProxmoxApi {
    param([string]$Method, [string]$Path, $Body, [int]$TimeoutSec = 0)
    [void]$script:apiCalls.Add("$Method $Path")

    if ($Path -match '/cluster/nextid') { return @{ data = "$($script:nextId)" } }

    if ($Method -eq 'PUT' -and $Path -match '/qemu/(\d+)/config$') {
        $vmid = $Matches[1]
        $script:vmTags[$vmid] = [string]$Body.tags
        [void]$script:putConfigCalls.Add(@{ VmId = $vmid; Tags = [string]$Body.tags })
        return @{ data = $null }
    }

    if ($Path -match '/cluster/resources') {
        $list = New-Object System.Collections.ArrayList
        $tmpl = [pscustomobject]@{ type = 'qemu'; vmid = 153; name = 'template'; node = 'n1'; status = 'stopped'; template = 1 }
        if ($script:vmTags.ContainsKey('153')) { $tmpl | Add-Member -MemberType NoteProperty -Name tags -Value $script:vmTags['153'] }
        [void]$list.Add($tmpl)
        foreach ($k in $script:vmState.Keys) {
            if (-not $script:vmVisible[$k]) { continue }
            $entry = [pscustomobject]@{ type = 'qemu'; vmid = [int]$k; node = 'n1'; status = $script:vmState[$k]; template = 0 }
            # Only attach a 'name' property when one is actually set -- this is
            # how Bug A's real trigger is reproduced: cluster/resources
            # reporting NO name property at all for a just-created clone
            # target, which is what actually produced "VM-<id>" (our own
            # missing-name fallback), not a genuine Proxmox placeholder string.
            if ($script:vmName.ContainsKey($k) -and -not [string]::IsNullOrWhiteSpace($script:vmName[$k])) {
                $entry | Add-Member -MemberType NoteProperty -Name name -Value $script:vmName[$k]
            }
            if ($script:vmTags.ContainsKey($k)) { $entry | Add-Member -MemberType NoteProperty -Name tags -Value $script:vmTags[$k] }
            [void]$list.Add($entry)
        }
        return @{ data = $list.ToArray() }
    }

    if ($Path -match '/qemu/(\d+)/clone$') {
        $newid = [string]$Body.newid
        $tid = "UPID:n1:clone$newid"
        $script:cloneJobs[$tid] = @{ vmid = $newid; done = $false }
        $script:vmState[$newid] = 'stopped'
        # No 'name' key set at all -- simulates cluster/resources reporting no
        # name property for a just-created clone target (Bug A's real cause).
        $script:vmVisible[$newid] = $true
        # Inherit the source's CURRENT tags, exactly like a real Proxmox full
        # clone copies the source config (Bug C).
        if ($script:vmTags.ContainsKey('153')) { $script:vmTags[$newid] = $script:vmTags['153'] }
        return @{ data = $tid }
    }

    if ($Path -match '/qemu/(\d+)/status/current$') {
        $vmid = $Matches[1]
        $st = if ($script:vmState.ContainsKey($vmid)) { $script:vmState[$vmid] } else { 'stopped' }
        return @{ data = [pscustomobject]@{ status = $st; qmpstatus = $st } }
    }

    if ($Path -match '/qemu/(\d+)/status/start$') {
        $vmid = $Matches[1]
        $script:vmState[$vmid] = 'running'
        return @{ data = "UPID:n1:start$vmid" }
    }

    if ($Path -match '/qemu/(\d+)/agent/network-get-interfaces$') {
        $vmid = $Matches[1]
        return @{ data = @{ result = @(
            [pscustomobject]@{ 'hardware-address' = 'AA:BB:CC:00:00:01'; 'ip-addresses' = @([pscustomobject]@{'ip-address-type'='ipv4'; 'ip-address'="10.0.0.$vmid"}) }
        ) } }
    }

    if ($Path -match '/tasks/(.+)/status$') {
        $tid = [uri]::UnescapeDataString($Matches[1])
        if ($script:cloneJobs.ContainsKey($tid)) {
            $st = if ($script:cloneJobs[$tid].done) { 'stopped' } else { 'running' }
            $ex = if ($script:cloneJobs[$tid].done) { 'OK' } else { $null }
            return @{ data = [pscustomobject]@{ status = $st; exitstatus = $ex } }
        }
        return @{ data = [pscustomobject]@{ status = 'stopped'; exitstatus = 'OK' } }
    }

    throw "Unexpected call in mock: $Method $Path"
}

# ------------------------------------------------------------
# Test A: name substitution must not depend on a specific placeholder
# pattern -- a clone id with NO 'name' property at all (our own fallback
# "VM-<id>" kicking in) must still resolve to the real intended name.
# ------------------------------------------------------------
$script:vmTags['153'] = 'rasTemplate153'   # source already tagged, as it would be after the pre-clone tag write
Reset-ProxmoxClusterCache
$body = @{ newid = $script:nextId; name = 'w10pwsh-A'; full = 1 }
$resp = Invoke-ProxmoxApi -Method POST -Path '/api2/json/nodes/n1/qemu/153/clone' -Body $body
$taskIdA = [string]$resp.data
$vmA = [string]$script:nextId

Set-CloneStateEntry -VmId $vmA -Entry @{
    type = 'clone'; task_id = $taskIdA; source_id = '153'; clone_id = $vmA
    name = 'w10pwsh-A'; full = $true; clone_node = 'n1'
    start_issued = $false; start_task_id = $null; start_pending = $false
    start_retry_count = 0; creation_completed = $false
    clone_started_at = (Get-Date).ToString('o')
}

Assert (-not $script:vmName.ContainsKey($vmA)) "setup: mock has no 'name' at all for the fresh clone, matching the real trigger"
$guestA = ConvertTo-RasGuestObject -VmId $vmA
Assert ($guestA.name -eq 'w10pwsh-A') "fix (bug A): a clone with NO Proxmox name at all (not just a specific placeholder pattern) still resolves to the real intended name (got '$($guestA.name)')"

# ------------------------------------------------------------
# Test B: the tasks/get vs guests/get race (bug B) -- guests/get's
# clone-aware flow wins the race and clears tracking; the NEXT tasks/get
# poll for that same task must still answer with the real clone_id, not an
# empty output.
# ------------------------------------------------------------
$script:nextId = 701
$body2 = @{ newid = $script:nextId; name = 'w10pwsh-B'; full = 1 }
$resp2 = Invoke-ProxmoxApi -Method POST -Path '/api2/json/nodes/n1/qemu/153/clone' -Body $body2
$taskIdB = [string]$resp2.data
$vmB = [string]$script:nextId
$script:vmName[$vmB] = 'w10pwsh-B'   # already resolved to its real name
$script:cloneJobs[$taskIdB].done = $true   # the REAL Proxmox clone job has finished

$ctxB = @{
    type = 'clone'; task_id = $taskIdB; source_id = '153'; clone_id = $vmB
    name = 'w10pwsh-B'; full = $true; clone_node = 'n1'
    start_issued = $false; start_task_id = $null; start_pending = $false
    start_retry_count = 0; creation_completed = $false
    clone_started_at = (Get-Date).AddSeconds(-30).ToString('o')
}
$script:TaskContext[$taskIdB] = $ctxB
Set-CloneStateEntry -VmId $vmB -Entry $ctxB

# guests/get's clone-aware flow runs first (e.g. via the sweep) and, since the
# VM is already powered on with an IP, confirms readiness and clears tracking
# -- winning the race against Handle-TaskInfo's own poll for the same task.
$script:vmState[$vmB] = 'running'
Reset-ProxmoxClusterCache
$guestB = Get-RasGuestObjectForCloneAwareFlow -VmId $vmB
Assert ($guestB.state -eq 'powered_on' -and -not [string]::IsNullOrWhiteSpace($guestB.ip)) "setup: VM B confirmed ready via guests/get (wins the race)"
Assert (-not $script:TaskContext.ContainsKey($taskIdB)) "setup: guests/get's flow cleared the in-memory task context"
Assert ($null -eq (Get-CloneStateEntry -VmId $vmB)) "setup: guests/get's flow cleared the persisted clone-state entry too"

# The next tasks/get poll for the SAME task arrives after tracking is gone.
$taskResultB = Handle-TaskInfo -Params ([pscustomobject]@{ id = $taskIdB })
Assert ($taskResultB.result.state -eq 'completed') "fix (bug B): tasks/get for the raced task still reports 'completed'"
Assert ($taskResultB.result.output.ContainsKey('clone_id')) "fix (bug B): tasks/get's output still contains clone_id even though tracking was already cleared by the race"
Assert ([string]$taskResultB.result.output.clone_id -eq $vmB) "fix (bug B): the recovered clone_id is correct ($vmB)"

# ------------------------------------------------------------
# Test C: tag inheritance (bug C, reported by the user from a live check) --
# the SOURCE was already tagged rasTemplate153 before the clone POST (as it
# now is, since the tag write moved earlier); the clone inherits that tag via
# Proxmox's config copy. Start-ProxmoxVmIfNeeded must strip it while applying
# rasClone153, in one PUT.
# ------------------------------------------------------------
Assert ($script:vmTags[$vmA] -eq 'rasTemplate153') "setup: clone A inherited rasTemplate153 from its source at clone time (Proxmox's real behavior)"

$script:vmState[$vmA] = 'stopped'
Reset-ProxmoxClusterCache
$script:putConfigCalls.Clear()
[void](Start-ProxmoxVmIfNeeded -VmId $vmA -CloneSourceVmId '153')

Assert ($script:vmTags[$vmA] -eq 'rasClone153') "fix (bug C): clone A ends up with ONLY rasClone153 -- the inherited rasTemplate153 was stripped"
Assert ($script:putConfigCalls.Count -eq 1) "fix (bug C): the strip + add happened in a single PUT, not two separate read-modify-writes"

# A clone that inherited NO template tag (e.g. source was never tagged) must
# not pay for a redundant PUT beyond adding its own rasClone tag.
$script:nextId = 702
$script:vmTags.Remove('153') | Out-Null
Reset-ProxmoxClusterCache
$bodyC2 = @{ newid = $script:nextId; name = 'w10pwsh-C'; full = 1 }
$respC2 = Invoke-ProxmoxApi -Method POST -Path '/api2/json/nodes/n1/qemu/153/clone' -Body $bodyC2
$vmC2 = [string]$script:nextId
$script:vmState[$vmC2] = 'stopped'
Reset-ProxmoxClusterCache
$script:putConfigCalls.Clear()
[void](Start-ProxmoxVmIfNeeded -VmId $vmC2 -CloneSourceVmId '153')
Assert ($script:vmTags[$vmC2] -eq 'rasClone153') "fix (bug C): a clone with no inherited template tag still gets rasClone153"
Assert ($script:putConfigCalls.Count -eq 1) "fix (bug C): still just one PUT when there's nothing to strip"

# Calling it again (idempotent) must not issue a second PUT at all.
$script:putConfigCalls.Clear()
Reset-ProxmoxClusterCache
[void](Start-ProxmoxVmIfNeeded -VmId $vmC2 -CloneSourceVmId '153')
Assert ($script:putConfigCalls.Count -eq 0) "fix (bug C): re-running against an already-correctly-tagged clone issues no PUT at all"

Write-Host "`n=== t7 Summary: $pass passed, $fail failed ===" -ForegroundColor $(if ($fail -eq 0) { 'Green' } else { 'Red' })
if ($fail -gt 0) { exit 1 }
