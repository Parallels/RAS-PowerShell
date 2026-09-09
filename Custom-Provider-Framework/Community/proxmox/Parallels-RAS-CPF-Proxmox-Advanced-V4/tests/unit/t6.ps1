$ErrorActionPreference = 'Stop'
$pass = 0
$fail = 0
function Assert {
    param([bool]$Cond, [string]$Msg)
    if ($Cond) { $script:pass++; Write-Host "PASS: $Msg" -ForegroundColor Green }
    else { $script:fail++; Write-Host "FAIL: $Msg" -ForegroundColor Red }
}

# ============================================================
# TESTS for:
#   clone-name orphaning (placeholder substitution + guests/list exemption)
#   rasTemplate tag write moved before clone POST, bounded timeout, once/source
#   recently-controlled -> live status/current instead of stale cluster listing
#   clone-state in-memory cache correctness (PSCustomObject normalization)
#   tracked-clone sweep rate limiting
# ============================================================
. "$PSScriptRoot/funcs2.ps1"

$script:LogPath        = "$PSScriptRoot/t6.log"
$script:CloneStatePath = "$PSScriptRoot/t6-state.json"
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
$script:CloneTaskCompletionCache = @{}

$script:apiCalls = New-Object System.Collections.ArrayList      # "METHOD PATH" strings, in order
$script:putConfigCalls = New-Object System.Collections.ArrayList # captured @{ VmId; Timeout } for config PUTs
$script:nextId = 300
$script:cloneJobs = @{}
$script:vmState = @{}     # vmid -> status ('running'/'stopped')
$script:vmName  = @{}     # vmid -> name
$script:vmVisible = @{}   # vmid -> $true once visible in cluster/resources
$script:vmTags = @{}      # vmid -> tag string
$script:liveStatusCalls = New-Object System.Collections.ArrayList # vmids queried via status/current

function Invoke-ProxmoxApi {
    param([string]$Method, [string]$Path, $Body, [int]$TimeoutSec = 0)
    [void]$script:apiCalls.Add("$Method $Path")

    if ($Path -match '/cluster/nextid') { return @{ data = "$($script:nextId)" } }

    if ($Method -eq 'PUT' -and $Path -match '/qemu/(\d+)/config$') {
        $vmid = $Matches[1]
        $script:vmTags[$vmid] = [string]$Body.tags
        [void]$script:putConfigCalls.Add(@{ VmId = $vmid; Timeout = $TimeoutSec })
        return @{ data = $null }
    }

    if ($Path -match '/cluster/resources') {
        $list = New-Object System.Collections.ArrayList
        $tmpl = [pscustomobject]@{ type = 'qemu'; vmid = 153; name = 'template'; node = 'n1'; status = 'stopped'; template = 1 }
        if ($script:vmTags.ContainsKey('153')) { $tmpl | Add-Member -MemberType NoteProperty -Name tags -Value $script:vmTags['153'] }
        [void]$list.Add($tmpl)
        foreach ($k in $script:vmState.Keys) {
            if (-not $script:vmVisible[$k]) { continue }
            $entry = [pscustomobject]@{ type = 'qemu'; vmid = [int]$k; name = $script:vmName[$k]; node = 'n1'; status = $script:vmState[$k]; template = 0 }
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
        $script:vmName[$newid]  = "VM $newid"
        $script:vmVisible[$newid] = $true
        return @{ data = $tid }
    }

    if ($Path -match '/qemu/(\d+)/status/current$') {
        $vmid = $Matches[1]
        [void]$script:liveStatusCalls.Add($vmid)
        $st = if ($script:vmState.ContainsKey($vmid)) { $script:vmState[$vmid] } else { 'stopped' }
        return @{ data = [pscustomobject]@{ status = $st; qmpstatus = $st } }
    }

    if ($Path -match '/qemu/(\d+)/status/start$') {
        $vmid = $Matches[1]
        $script:vmState[$vmid] = 'running'
        return @{ data = "UPID:n1:start$vmid" }
    }

    if ($Path -match '/qemu/(\d+)/agent/network-get-interfaces$') {
        throw 'agent not queried in this test'
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
# Test A: Test-ProxmoxPlaceholderVmName basic correctness
# ------------------------------------------------------------
Assert (Test-ProxmoxPlaceholderVmName -VmId '136' -VmName 'VM 136') "placeholder: 'VM 136' matches for id 136"
Assert (-not (Test-ProxmoxPlaceholderVmName -VmId '136' -VmName 'w10pwsh-001')) "placeholder: real name does not match"
Assert (-not (Test-ProxmoxPlaceholderVmName -VmId '136' -VmName 'VM 137')) "placeholder: mismatched id does not match"
Assert (-not (Test-ProxmoxPlaceholderVmName -VmId '136' -VmName '')) "placeholder: empty name does not match"

# ------------------------------------------------------------
# Test B: clone-name orphaning fix
# ------------------------------------------------------------
$body = @{ newid = $script:nextId; name = 'w10pwsh-001'; full = 1 }
$resp = Invoke-ProxmoxApi -Method POST -Path '/api2/json/nodes/n1/qemu/153/clone' -Body $body
$taskId = [string]$resp.data
$newVmId = [string]$script:nextId

Set-CloneStateEntry -VmId $newVmId -Entry @{
    type = 'clone'; task_id = $taskId; source_id = '153'; clone_id = $newVmId
    name = 'w10pwsh-001'; full = $true; clone_node = 'n1'
    start_issued = $false; start_task_id = $null; start_pending = $false
    start_retry_count = 0; creation_completed = $false
    clone_started_at = (Get-Date).ToString('o')
}

# The mock leaves the VM under Proxmox's placeholder name ("VM <id>") until we
# explicitly rename it below, simulating the real clone-in-progress window.
Assert ($script:vmName[$newVmId] -eq "VM $newVmId") "setup: mock VM still carries the Proxmox placeholder name"

$guest = ConvertTo-RasGuestObject -VmId $newVmId
Assert ($guest.name -eq 'w10pwsh-001') "ConvertTo-RasGuestObject substitutes the tracked clone's real name instead of the placeholder (got '$($guest.name)')"

# guests/list's gate on a tracked clone is not Proxmox's (or this provider's
# own) NAME for it; it's whether this provider has itself reported the clone
# task 'completed' to RAS yet (see Set-CloneReportedCompleted /
# Handle-TaskInfo). This entry hasn't been, so it must still be excluded even
# though its name already resolves correctly above -- RAS's clone thread and
# guests/list must not disagree about when a clone becomes "real" to the
# outside world.
$listResultBeforeCompletion = Handle-GuestList
Assert (-not (@($listResultBeforeCompletion.result.guests) -contains $newVmId)) "guests/list still omits a tracked clone whose task has not been reported completed yet, even once its name resolves"

Set-CloneReportedCompleted -VmId $newVmId
$listResult = Handle-GuestList
Assert (@($listResult.result.guests) -contains $newVmId) "guests/list includes a tracked clone once its task has been reported completed"

# An UNTRACKED VM with a placeholder-shaped name must still be filtered --
# this only exempts clones this provider itself is tracking.
$untrackedId = '999'
$script:vmState[$untrackedId] = 'stopped'
$script:vmName[$untrackedId] = "VM $untrackedId"
$script:vmVisible[$untrackedId] = $true
$listResult2 = Handle-GuestList
Assert (-not (@($listResult2.result.guests) -contains $untrackedId)) "an UNTRACKED placeholder-named VM is still filtered from guests/list"

# Once Proxmox itself renames the VM (real clone job finished), the real name
# is reported directly -- no stale substitution once it's no longer needed.
$script:vmName[$newVmId] = 'w10pwsh-001'
Reset-ProxmoxClusterCache
$guest2 = ConvertTo-RasGuestObject -VmId $newVmId
Assert ($guest2.name -eq 'w10pwsh-001') "once Proxmox reports the real name, it is used directly (no substitution needed)"

Remove-CloneStateEntry -VmId $newVmId
$script:vmState.Remove($newVmId); $script:vmName.Remove($newVmId); $script:vmVisible.Remove($newVmId)
$script:vmState.Remove($untrackedId); $script:vmName.Remove($untrackedId); $script:vmVisible.Remove($untrackedId)

# ------------------------------------------------------------
# Test C: rasTemplate tag write moved before the clone POST, bounded
# timeout, attempted once per source per process
# ------------------------------------------------------------
$script:apiCalls.Clear()
$script:putConfigCalls.Clear()
$cloneResult1 = Handle-GuestClone -Params @{ id = '153'; name = 'w10pwsh-A' }
Assert (-not $cloneResult1.ContainsKey('error')) "clone #1 succeeded"

$cloneIdxPost = [array]::IndexOf($script:apiCalls.ToArray(), ($script:apiCalls | Where-Object { $_ -match '/qemu/153/clone$' } | Select-Object -First 1))
$tagIdxPut = -1
for ($i = 0; $i -lt $script:apiCalls.Count; $i++) {
    if ($script:apiCalls[$i] -match '^PUT .*/qemu/153/config$') { $tagIdxPut = $i; break }
}
Assert ($tagIdxPut -ge 0 -and $tagIdxPut -lt $cloneIdxPost) "rasTemplate tag PUT is issued BEFORE the clone POST (tag at index $tagIdxPut, clone POST at index $cloneIdxPost)"
Assert ($script:vmTags['153'] -eq 'rasTemplate153') "source VM tagged rasTemplate153"
Assert ($script:putConfigCalls.Count -eq 1 -and [int]$script:putConfigCalls[0].Timeout -eq $script:TagWriteTimeoutSeconds) "tag write used the bounded timeout ($($script:TagWriteTimeoutSeconds)s), not unbounded"

# A second clone from the SAME source must not re-attempt the tag write at
# all -- neither via the already-tagged check nor a second attempt.
$script:putConfigCalls.Clear()
$cloneResult2 = Handle-GuestClone -Params @{ id = '153'; name = 'w10pwsh-B' }
Assert (-not $cloneResult2.ContainsKey('error')) "clone #2 (same source) succeeded"
Assert ($script:putConfigCalls.Count -eq 0) "second clone from the same source does not re-attempt the rasTemplate tag write"
Assert ($script:RasTemplateTagAttempted.Contains('153')) "source id recorded as attempted for this process lifetime"

# cleanup the two clones this test created
foreach ($id in @($cloneResult1.result.clone_id, $cloneResult2.result.clone_id)) {
    $sid = [string]$id
    Remove-CloneStateEntry -VmId $sid
    $script:vmState.Remove($sid); $script:vmName.Remove($sid); $script:vmVisible.Remove($sid)
}
$script:vmTags.Remove('153')
Reset-ProxmoxClusterCache

# ------------------------------------------------------------
# Test D: recently-controlled -> live status/current instead of the
# lagging cluster listing
# ------------------------------------------------------------
$ctlId = '400'
$script:vmState[$ctlId] = 'running'
$script:vmName[$ctlId] = 'ctl-test'
$script:vmVisible[$ctlId] = $true
Reset-ProxmoxClusterCache
$g0 = ConvertTo-RasGuestObject -VmId $ctlId
Assert ($g0.state -eq 'powered_on') "setup: control-test VM initially reports powered_on via the (cached) cluster listing"

# Simulate: we just told Proxmox to stop it, but the cluster/resources
# listing (server-side aggregate) hasn't caught up yet -- only the mock's
# authoritative status/current has.
$script:RecentlyControlledIds[$ctlId] = [DateTime]::UtcNow
$script:vmState[$ctlId] = 'stopped'   # authoritative (status/current) view
# Deliberately do NOT update $script:vmVisible-backed cluster/resources cache entry's
# status -- but since our mock cluster/resources always reads live from $script:vmState,
# force a stale cached snapshot to simulate the real lag:
$staleList = @([pscustomobject]@{ type = 'qemu'; vmid = 153; name = 'template'; node = 'n1'; status = 'stopped'; template = 1 },
    [pscustomobject]@{ type = 'qemu'; vmid = [int]$ctlId; name = 'ctl-test'; node = 'n1'; status = 'running'; template = 0 })
$script:ClusterResourcesCache = $staleList
$script:ClusterResourcesCachedAt = [DateTime]::UtcNow

$script:liveStatusCalls.Clear()
$g1 = ConvertTo-RasGuestObject -VmId $ctlId
Assert ($script:liveStatusCalls -contains $ctlId) "a recently-controlled VM triggers a live status/current check"
Assert ($g1.state -eq 'powered_off') "the live status/current result overrides the stale cached 'running' (got '$($g1.state)')"

# A VM NOT recently controlled must NOT pay the extra status/current call.
$otherId = '401'
$script:vmState[$otherId] = 'running'
$script:vmName[$otherId] = 'other'
$script:vmVisible[$otherId] = $true
Reset-ProxmoxClusterCache
$script:liveStatusCalls.Clear()
$g2 = ConvertTo-RasGuestObject -VmId $otherId
Assert (-not ($script:liveStatusCalls -contains $otherId)) "a VM with no recent control action does not trigger the extra live check"

$script:vmState.Remove($ctlId); $script:vmName.Remove($ctlId); $script:vmVisible.Remove($ctlId); $script:RecentlyControlledIds.Remove($ctlId)
$script:vmState.Remove($otherId); $script:vmName.Remove($otherId); $script:vmVisible.Remove($otherId)
Reset-ProxmoxClusterCache

# ------------------------------------------------------------
# Test E: clone-state in-memory cache correctness --
# a freshly Set-CloneStateEntry'd entry must be immediately and correctly
# visible to every reader (Get-ActiveCloneCount, Get-CloneStateEntry,
# Get-CloneStateEntryByTaskId) with NO intervening disk round-trip, and its
# .PSObject.Properties shape must behave identically to a disk-loaded entry.
# ------------------------------------------------------------
$cacheTestId = '500'
$cacheTaskId = 'UPID:n1:clonecache500'
# Get-ActiveCloneCount checks the
# entry's REAL underlying Proxmox task before counting it (only the disk-copy
# job still owns a concurrency slot; the boot-to-ready tail does not), so the
# mock's clone job must be registered as still running for this entry
# to correctly count as active -- matching a real in-flight clone.
$script:cloneJobs[$cacheTaskId] = @{ done = $false }
Set-CloneStateEntry -VmId $cacheTestId -Entry @{
    type = 'clone'; task_id = $cacheTaskId; source_id = '153'; clone_id = $cacheTestId
    name = 'cache-test'; full = $true; clone_node = 'n1'
    start_issued = $false; start_task_id = $null; start_pending = $false
    start_retry_count = 0; creation_completed = $false
    clone_started_at = (Get-Date).ToString('o')
}

Assert ((Get-ActiveCloneCount) -ge 1) "Get-ActiveCloneCount sees the freshly-set entry with no disk round-trip (its real clone task is still running)"

# Once the real Proxmox task is confirmed done, the entry no
# longer counts as active -- even though it hasn't reached full boot-ready
# yet (Remove-CloneStateEntry has not been called) -- and that confirmation
# is cached so a second check costs no further HTTP call.
$script:cloneJobs[$cacheTaskId].done = $true
$script:apiCalls.Clear()
Assert ((Get-ActiveCloneCount) -eq 0) "once the real clone task is confirmed completed, the entry no longer counts toward active concurrency"
Assert ($script:apiCalls.Count -eq 1) "confirming completion cost exactly one task-status call"

$script:apiCalls.Clear()
Assert ((Get-ActiveCloneCount) -eq 0) "a second check still reports 0 active"
Assert ($script:apiCalls.Count -eq 0) "the second check made no HTTP call at all -- the confirmed-completed task id is cached"

$fetchedEntry = Get-CloneStateEntry -VmId $cacheTestId
Assert ($null -ne $fetchedEntry -and ($fetchedEntry.PSObject.Properties.Name -contains 'type')) "Get-CloneStateEntry's returned entry exposes .PSObject.Properties (PSCustomObject-shaped, not a raw Hashtable)"
Assert ([string]$fetchedEntry.type -eq 'clone') "fetched entry's 'type' field is readable via property access"

$fetchedByTask = Get-CloneStateEntryByTaskId -TaskId $cacheTaskId
Assert ($null -ne $fetchedByTask -and [string]$fetchedByTask.clone_id -eq $cacheTestId) "Get-CloneStateEntryByTaskId finds the freshly-set entry by task id with no disk round-trip"

Remove-CloneStateEntry -VmId $cacheTestId
Assert ($null -eq (Get-CloneStateEntry -VmId $cacheTestId)) "Remove-CloneStateEntry removes the entry from the (shared) cache too"

# ------------------------------------------------------------
# Test F: tracked-clone sweep rate limiting
# ------------------------------------------------------------
$sweepId = '600'
$script:vmState[$sweepId] = 'running'
$script:vmName[$sweepId] = 'sweep-test'
$script:vmVisible[$sweepId] = $true
Set-CloneStateEntry -VmId $sweepId -Entry @{
    type = 'clone'; task_id = 'UPID:n1:clonesweep600'; source_id = '153'; clone_id = $sweepId
    name = 'sweep-test'; full = $true; clone_node = 'n1'
    start_issued = $false; start_task_id = $null; start_pending = $false
    start_retry_count = 0; creation_completed = $false
    clone_started_at = (Get-Date).AddSeconds(-30).ToString('o')
}
$script:cloneJobs['UPID:n1:clonesweep600'] = @{ vmid = $sweepId; done = $true }

$script:SweepLastCheckedAt.Clear()
Reset-ProxmoxClusterCache
$before = $script:liveStatusCalls.Count
Invoke-TrackedCloneSweep
$afterFirst = $script:liveStatusCalls.Count
Invoke-TrackedCloneSweep
$afterSecond = $script:liveStatusCalls.Count

Assert ($script:SweepLastCheckedAt.ContainsKey($sweepId)) "sweep records a last-checked timestamp for the tracked id"
Assert ($afterSecond -eq $afterFirst) "an immediate second sweep call does not re-check a VM within the throttle window (no new HTTP activity)"

# After the throttle window elapses, the sweep must check again.
$script:SweepLastCheckedAt[$sweepId] = [DateTime]::UtcNow.AddSeconds(-1 * ($script:SweepMinIntervalSeconds + 1))
$beforeThird = $script:liveStatusCalls.Count
Invoke-TrackedCloneSweep
Assert ($script:SweepLastCheckedAt[$sweepId] -gt [DateTime]::UtcNow.AddSeconds(-2)) "after the throttle window elapses, the sweep checks the VM again and refreshes its timestamp"

Remove-CloneStateEntry -VmId $sweepId
$script:vmState.Remove($sweepId); $script:vmName.Remove($sweepId); $script:vmVisible.Remove($sweepId)

Write-Host "`n=== t6 Summary: $pass passed, $fail failed ===" -ForegroundColor $(if ($fail -eq 0) { 'Green' } else { 'Red' })
if ($fail -gt 0) { exit 1 }
