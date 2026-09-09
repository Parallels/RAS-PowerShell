$ErrorActionPreference = 'Stop'
$pass = 0
$fail = 0
function Assert {
    param([bool]$Cond, [string]$Msg)
    if ($Cond) { $script:pass++; Write-Host "PASS: $Msg" -ForegroundColor Green }
    else { $script:fail++; Write-Host "FAIL: $Msg" -ForegroundColor Red }
}

# ============================================================
# TESTS for the remaining features that t8/t11/t12/t13/t14
# do not already cover:
#
#   A. Agent-quarantine PERMANENT EXEMPTION for rasClone*/rasTemplate*-tagged VMs -- the
#      persisted rasQuarantine tag is never written for a RAS-managed guest, logged once
#      per failure episode (not every poll), and an ordinary (non-RAS-tagged) guest still
#      gets quarantined normally -- the exemption is tag-specific, not a global change.
#   B. The clone-tracking garbage collector (cloning.timeouts.clone_tracking_max_age_seconds)
#      -- a clone tracked past this ceiling is force-retired even though it never reached
#      ready, while one still inside the window is left alone.
#   C. Recently-deleted NAME preservation (the VM-<id> placeholder bug) -- guests/get's
#      "recently deleted" stub echoes the guest's real last-known name, not a synthetic
#      VM-<id>.
#   D. $script:CloneTagVerified -- Start-ProxmoxVmIfNeeded's tag repair/verification does
#      one live config GET per VmId per process lifetime, not one per poll.
#   E. Link-local (169.254.x.x) IP reporting -- surfaced in ip_addresses (not stripped),
#      ordered after any real address, and never counted toward the clone-ready gate.
#   F. Handle-TaskInfo's simplification to spec -- 'running'/'failed' still pass straight
#      through; 'completed' fires the instant the REAL Proxmox task is done with zero
#      guest-readiness check; and the pipelined-completion shortcut only ever applies to a
#      full clone (ctx.full), never a linked one, and respects the concurrency gate.
# ============================================================
. "$PSScriptRoot/funcs2.ps1"

$script:LogPath        = "$PSScriptRoot/t15.log"
$script:CloneStatePath = "$PSScriptRoot/t15-state.json"
Remove-Item $script:LogPath, $script:CloneStatePath -ErrorAction SilentlyContinue
$script:TrackedCloneVmIds = [System.Collections.Generic.HashSet[string]]::new()
$script:ClusterResourcesCache = $null
$script:ClusterResourcesCachedAt = $null
$script:RecentlyDeletedIds = @{}
$script:RecentlyDeletedNames = @{}
$script:RecentlyDeletedNodeByName = @{}
$script:RecentlyControlledIds = @{}
$script:RecentlyStoppedIds = @{}
$script:RecentlyConvertedIds = @{}
$script:PendingTemplateTagRemovalIds = @{}
$script:MaintenanceModeVmIds = [System.Collections.Generic.HashSet[string]]::new()
$script:AgentFailureTracker = @{}
$script:LastKnownNetworkData = @{}
$script:SweepLastCheckedAt = @{}
$script:LastGuestPollAt = @{}
$script:RasTemplateTagAttempted = [System.Collections.Generic.HashSet[string]]::new()
$script:RasManagedAgentExemptionLogged = [System.Collections.Generic.HashSet[string]]::new()
$script:CloneTagVerified = [System.Collections.Generic.HashSet[string]]::new()
$script:CloneStateMemory = $null
$script:TaskContext = @{}
$script:CloneTaskCompletionCache = @{}
$script:CompletedCloneTaskOutputs = @{}

# Derived runtime variables -- Set-ProviderRuntimeFromSettings populates these once at
# dot-source time from Tests/unit/RAS-CPF-Proxmox-Settings.json; every function under
# test here reads the $script:* variable, never $script:Settings directly on each call
# (see the t14.ps1 comment on this same gotcha), so this suite sets both for clarity and
# to be self-contained regardless of whatever the fixture file happens to hold.
$script:Settings.virtual_machines.guest_agent.quarantine_after_seconds = 60
$script:Settings.virtual_machines.guest_agent.quarantine_min_failures = 2
$script:AgentQuarantineAfterSeconds = 60
$script:AgentQuarantineMinFailures = 2
$script:RasQuarantineTag = 'rasQuarantine'
$script:RasTemplateTagPrefix = 'rasTemplate'
$script:RasCloneTagPrefix = 'rasClone'

$script:Settings.cloning.timeouts.clone_tracking_max_age_seconds = 1800
$script:CloneTrackingMaxAgeSeconds = 1800

$script:Settings.cloning.pipelined_cloning.enabled = $true
$script:Settings.cloning.pipelined_cloning.completion_seconds = 10
$script:Settings.cloning.pipelined_cloning.max_concurrent_clone_operations = 2
$script:PipelinedCloneCompletionEnabled = $true
$script:PipelinedCloneCompletionSeconds = 10
$script:MaxConcurrentCloneOperations = 2

$script:Settings.cloning.timeouts.recently_deleted_retention_seconds = 300
$script:RecentlyDeletedRetentionSeconds = 300
$script:Settings.cloning.timeouts.completed_clone_task_output_retention_seconds = 600
$script:CompletedCloneTaskOutputRetentionSeconds = 600
$script:Settings.cloning.timeouts.delete_hard_stop_max_wait_seconds = 60
$script:DeleteHardStopMaxWaitSeconds = 60

$script:apiCalls = New-Object System.Collections.ArrayList
$script:vmState    = @{}   # vmid -> status ('running'/'stopped')
$script:vmName     = @{}
$script:vmTags     = @{}   # vmid -> ';'-joined tag string
$script:agentMode  = @{}   # vmid -> 'fail' | array of interface objects
$script:taskStatus = @{}   # taskId -> @{ status; exitstatus }

function Invoke-ProxmoxApi {
    param([string]$Method, [string]$Path, $Body, [int]$TimeoutSec = 0)
    [void]$script:apiCalls.Add("$Method $Path")

    if ($Path -match '/cluster/resources') {
        $list = New-Object System.Collections.ArrayList
        foreach ($k in $script:vmState.Keys) {
            $entry = [pscustomobject]@{ type = 'qemu'; vmid = [int]$k; name = $script:vmName[$k]; node = 'n1'; status = $script:vmState[$k]; template = 0 }
            if ($script:vmTags.ContainsKey($k)) { $entry | Add-Member -NotePropertyName tags -NotePropertyValue $script:vmTags[$k] }
            [void]$list.Add($entry)
        }
        return @{ data = $list.ToArray() }
    }

    if ($Method -eq 'GET' -and $Path -match '/qemu/(\d+)/config$') {
        $vmid = $Matches[1]
        $obj = [ordered]@{ template = 0 }
        if ($script:vmTags.ContainsKey($vmid)) { $obj.tags = $script:vmTags[$vmid] }
        return @{ data = [pscustomobject]$obj }
    }

    if ($Method -eq 'PUT' -and $Path -match '/qemu/(\d+)/config$') {
        $vmid = $Matches[1]
        if ($null -ne $Body -and $Body.ContainsKey('tags')) { $script:vmTags[$vmid] = [string]$Body.tags }
        return @{ data = $null }
    }

    if ($Path -match '/qemu/(\d+)/status/current$') {
        $vmid = $Matches[1]
        $st = if ($script:vmState.ContainsKey($vmid)) { $script:vmState[$vmid] } else { 'stopped' }
        return @{ data = [pscustomobject]@{ status = $st; qmpstatus = $st } }
    }

    if ($Method -eq 'POST' -and $Path -match '/qemu/(\d+)/status/(start|stop)$') {
        $vmid = $Matches[1]; $action = $Matches[2]
        $script:vmState[$vmid] = if ($action -eq 'start') { 'running' } else { 'stopped' }
        return @{ data = "UPID:n1:qm$($action)$vmid" }
    }

    if ($Method -eq 'DELETE' -and $Path -match '/qemu/(\d+)$') {
        return @{ data = $null }
    }

    if ($Path -match '/qemu/(\d+)/agent/network-get-interfaces$') {
        $vmid = $Matches[1]
        $mode = $script:agentMode[$vmid]
        if ($null -eq $mode -or $mode -eq 'fail') {
            throw "Response status code does not indicate success: 500 (VM $vmid is not running)."
        }
        return @{ data = [pscustomobject]@{ result = @($mode) } }
    }

    if ($Path -match '/tasks/(.+)/status$') {
        $taskId = [System.Uri]::UnescapeDataString($Matches[1])
        if ($script:taskStatus.ContainsKey($taskId)) {
            $st = $script:taskStatus[$taskId]
            return @{ data = [pscustomobject]@{ status = $st.status; exitstatus = $st.exitstatus } }
        }
        return @{ data = [pscustomobject]@{ status = 'stopped'; exitstatus = 'OK' } }
    }

    throw "Unexpected call in mock: $Method $Path"
}

function Count-ConfigGets {
    param([string]$VmId)
    return @($script:apiCalls | Where-Object { $_ -eq "GET /api2/json/nodes/n1/qemu/$VmId/config" }).Count
}

# ------------------------------------------------------------
# A. Agent-quarantine permanent exemption for rasClone*/rasTemplate*-tagged VMs.
# ------------------------------------------------------------
$script:vmState['801'] = 'running'; $script:vmName['801'] = 'clone801'; $script:vmTags['801'] = 'rasClone153'
$script:vmState['802'] = 'running'; $script:vmName['802'] = 'plain802'   # no RAS tag at all
$script:agentMode['801'] = 'fail'
$script:agentMode['802'] = 'fail'
$clusterVmA = [pscustomobject]@{ vmid = 801; tags = 'rasClone153' }
$clusterVmB = [pscustomobject]@{ vmid = 802 }

# First failure for both: count=1, first_failure_at=now -- too fresh to trip the
# after_seconds gate yet, so neither writes anything.
[void](Get-ProxmoxVmGuestAgentInterfaces -Node 'n1' -VmId '801' -ClusterVm $clusterVmA)
[void](Get-ProxmoxVmGuestAgentInterfaces -Node 'n1' -VmId '802' -ClusterVm $clusterVmB)
Assert (-not $script:vmTags.ContainsKey('801') -or $script:vmTags['801'] -notmatch 'rasQuarantine') "A1: a single fresh failure does not quarantine (below quarantine_after_seconds)"

# Backdate both trackers past quarantine_after_seconds so the NEXT failure crosses both
# gates (elapsed time AND failure count) at once.
$script:AgentFailureTracker['801'].first_failure_at = [DateTime]::UtcNow.AddSeconds(-61)
$script:AgentFailureTracker['802'].first_failure_at = [DateTime]::UtcNow.AddSeconds(-61)

[void](Get-ProxmoxVmGuestAgentInterfaces -Node 'n1' -VmId '801' -ClusterVm $clusterVmA)
Assert (-not $script:vmTags.ContainsKey('801') -or $script:vmTags['801'] -notmatch 'rasQuarantine') "A2: a RAS-managed guest (rasClone153 tag) is NEVER given the persisted rasQuarantine tag, even past both thresholds"
$logAfterFirstTrip = Get-Content $script:LogPath -Raw
Assert ($logAfterFirstTrip -match 'exempt from') "A3: the exemption is logged once the thresholds are crossed"
$exemptLineCount1 = ([regex]::Matches($logAfterFirstTrip, 'exempt from')).Count
Assert ($exemptLineCount1 -eq 1) "A4: ...exactly once for this failure episode so far"

# Re-poll again without recovering -- must NOT log a second time for the same episode.
[void](Get-ProxmoxVmGuestAgentInterfaces -Node 'n1' -VmId '801' -ClusterVm $clusterVmA)
[void](Get-ProxmoxVmGuestAgentInterfaces -Node 'n1' -VmId '801' -ClusterVm $clusterVmA)
$logAfterRepoll = Get-Content $script:LogPath -Raw
$exemptLineCount2 = ([regex]::Matches($logAfterRepoll, 'exempt from')).Count
Assert ($exemptLineCount2 -eq 1) "A5: repolling the SAME still-failing episode does not re-log the exemption every poll"

# Contrast: the ordinary (non-RAS-tagged) guest DOES get quarantined once past the same
# thresholds -- the exemption is specific to the rasClone*/rasTemplate* tag, not a global
# behavior change.
[void](Get-ProxmoxVmGuestAgentInterfaces -Node 'n1' -VmId '802' -ClusterVm $clusterVmB)
Assert ($script:vmTags.ContainsKey('802') -and $script:vmTags['802'] -match 'rasQuarantine') "A6: an ordinary guest with no rasClone*/rasTemplate* tag is still quarantined normally past the same thresholds"

# Recovery clears the exemption-log dedup set, so a genuinely NEW failure episode later
# logs again rather than staying silent forever.
$script:agentMode['801'] = @([pscustomobject]@{ 'hardware-address' = 'AA:BB:CC:00:03:01'
        'ip-addresses' = @([pscustomobject]@{ 'ip-address-type' = 'ipv4'; 'ip-address' = '10.0.3.1' }) })
[void](Get-ProxmoxVmGuestAgentInterfaces -Node 'n1' -VmId '801' -ClusterVm $clusterVmA)
Assert (-not $script:RasManagedAgentExemptionLogged.Contains('801')) "A7: a successful agent probe clears the exemption dedup set for the next episode"

$script:agentMode['801'] = 'fail'
[void](Get-ProxmoxVmGuestAgentInterfaces -Node 'n1' -VmId '801' -ClusterVm $clusterVmA)
$script:AgentFailureTracker['801'].first_failure_at = [DateTime]::UtcNow.AddSeconds(-61)
[void](Get-ProxmoxVmGuestAgentInterfaces -Node 'n1' -VmId '801' -ClusterVm $clusterVmA)
$logAfterSecondEpisode = Get-Content $script:LogPath -Raw
$exemptLineCount3 = ([regex]::Matches($logAfterSecondEpisode, 'exempt from')).Count
Assert ($exemptLineCount3 -eq 2) "A8: a genuinely new failure episode (after a successful probe in between) logs the exemption again"

# ------------------------------------------------------------
# B. Clone-tracking garbage collector (clone_tracking_max_age_seconds).
# ------------------------------------------------------------
# A clone stuck past the ceiling: state file says it's still 'clone', but Proxmox itself
# now shows a perfectly normal, already-running, real guest (the shape you get after
# manually deleting/recreating outside RAS, or the id being reused) -- the GC must
# retire tracking BEFORE any clone-aware logic runs, so this resolves as an ordinary
# untracked guest, not "powering_on" forever.
$script:vmState['901'] = 'running'; $script:vmName['901'] = 'vm901'
$script:agentMode['901'] = @([pscustomobject]@{ 'hardware-address' = 'AA:BB:CC:00:09:01'
        'ip-addresses' = @([pscustomobject]@{ 'ip-address-type' = 'ipv4'; 'ip-address' = '10.0.9.1' }) })
# Get-ProxmoxVmNode (via ConvertTo-RasGuestObject) reads through Get-ProxmoxClusterVMs's
# OWN cache, separate from this mock's own bookkeeping -- bust it whenever a vmid new to
# $script:vmState is about to be looked up for the first time, or a stale cached listing
# from before this vmid existed hides it just like a real lagging cluster/resources would.
$script:ClusterResourcesCache = $null; $script:ClusterResourcesCachedAt = $null
Set-CloneStateEntry -VmId '901' -Entry @{
    type = 'clone'; task_id = 'UPID:n1:clone901'; source_id = '153'; clone_id = '901'
    name = 'gc-test-901'; full = $true
    start_issued = $true; start_task_id = $null; start_pending = $false
    start_retry_count = 1; creation_completed = $false
    clone_started_at = ([DateTime]::UtcNow).AddSeconds(-1801).ToString('o')   # 1s past the 1800s ceiling
}
$gcGuest = Get-RasGuestObjectForCloneAwareFlow -VmId '901'
Assert ($gcGuest.state -eq 'powered_on') "B1: a clone tracked past clone_tracking_max_age_seconds is force-retired -- it resolves as a normal (already-running) guest, not powering_on"
Assert ($null -eq (Get-CloneStateEntry -VmId '901')) "B2: ...and its persisted clone-state entry is gone"
$gcLog = Get-Content $script:LogPath -Raw
Assert ($gcLog -match 'tracking abandoned') "B3: the GC logs why it gave up on this VM"

# Contrast: a clone still well inside the window is left completely alone.
$script:vmState['902'] = 'stopped'; $script:vmName['902'] = 'VM 902'
$script:ClusterResourcesCache = $null; $script:ClusterResourcesCachedAt = $null
Set-CloneStateEntry -VmId '902' -Entry @{
    type = 'clone'; task_id = 'UPID:n1:clone902'; source_id = '153'; clone_id = '902'
    name = 'gc-test-902'; full = $true
    start_issued = $false; start_task_id = $null; start_pending = $false
    start_retry_count = 0; creation_completed = $false
    clone_started_at = ([DateTime]::UtcNow).AddSeconds(-100).ToString('o')   # well inside 1800s
}
$script:taskStatus['UPID:n1:clone902'] = @{ status = 'stopped'; exitstatus = 'OK' }
[void](Get-RasGuestObjectForCloneAwareFlow -VmId '902')
Assert ($null -ne (Get-CloneStateEntry -VmId '902')) "B4: a clone well inside the ceiling keeps its tracking -- the GC only fires past the age limit"

# ------------------------------------------------------------
# C. Recently-deleted NAME preservation (the VM-<id> placeholder bug).
# ------------------------------------------------------------
$script:vmState['153'] = 'stopped'; $script:vmName['153'] = 'realname153'
$script:ClusterResourcesCache = $null; $script:ClusterResourcesCachedAt = $null
$deleteResult = Handle-GuestControl -Params ([pscustomobject]@{ id = '153'; control = 'delete' })
Assert (-not ($deleteResult -is [hashtable] -and $deleteResult.ContainsKey('error'))) "C1: guests/control(delete) succeeds"
Assert ((Get-ProxmoxRecentlyDeletedName -VmId '153') -eq 'realname153') "C2: the recently-deleted stub echoes the guest's REAL name, not a synthetic VM-153 placeholder"

$getResult = Handle-GuestGet -Params ([pscustomobject]@{ id = '153' })
Assert ($getResult.result.name -eq 'realname153') "C3: guests/get for the same id during the retention window reports the real name end to end (this is what RAS actually receives)"
Assert ($getResult.result.state -eq 'powered_off') "C4: ...and the deleted-stub state is powered_off, as before this fix"

# ------------------------------------------------------------
# D. $script:CloneTagVerified -- one live config GET per VmId per process lifetime.
# ------------------------------------------------------------
$script:vmState['1001'] = 'stopped'; $script:vmName['1001'] = 'clone1001'
$script:vmTags['1001'] = 'rasTemplate153'   # inherited from the source at clone time -- needs stripping + rasClone153 added
$script:ClusterResourcesCache = $null; $script:ClusterResourcesCachedAt = $null

$countBefore = Count-ConfigGets -VmId '1001'
[void](Start-ProxmoxVmIfNeeded -VmId '1001' -CloneSourceVmId '153')
$countAfterFirst = Count-ConfigGets -VmId '1001'
Assert ($countAfterFirst -eq ($countBefore + 1)) "D1: the first call for a not-yet-verified VmId does exactly one live config GET (tag repair)"
Assert ($script:vmTags['1001'] -eq 'rasClone153') "D2: ...and the inherited rasTemplate153 tag was stripped while rasClone153 was applied, in that one PUT"
Assert ($script:CloneTagVerified.Contains('1001')) "D3: the VmId is now cached as verified"

[void](Start-ProxmoxVmIfNeeded -VmId '1001' -CloneSourceVmId '153')
[void](Start-ProxmoxVmIfNeeded -VmId '1001' -CloneSourceVmId '153')
$countAfterMore = Count-ConfigGets -VmId '1001'
Assert ($countAfterMore -eq $countAfterFirst) "D4: further calls for the SAME VmId do not repeat the config GET -- served from the CloneTagVerified cache"

# Cache invalidation: Remove-CloneStateEntry (called on delete / by the GC above) also
# drops CloneTagVerified, so a VMID Proxmox later reuses for an unrelated clone always
# re-verifies fresh rather than trusting a stale cache entry for a different VM.
Remove-CloneStateEntry -VmId '1001'
Assert (-not $script:CloneTagVerified.Contains('1001')) "D5: Remove-CloneStateEntry also clears the CloneTagVerified cache for that VmId"
[void](Start-ProxmoxVmIfNeeded -VmId '1001' -CloneSourceVmId '153')
$countAfterReuse = Count-ConfigGets -VmId '1001'
Assert ($countAfterReuse -eq ($countAfterMore + 1)) "D6: ...so a fresh call after that invalidation does one more real config GET rather than staying cached forever"

# ------------------------------------------------------------
# E. Link-local IP reporting -- surfaced, ordered after a real address, never counted
#    toward the clone-ready gate.
# ------------------------------------------------------------
# E1: unit-level ordering check on Get-ProxmoxVmNetworkData directly -- agent reports
# link-local FIRST; the function must still put the real address first in the list.
$script:vmState['1100'] = 'running'; $script:vmName['1100'] = 'vm1100'
$script:agentMode['1100'] = @(
    [pscustomobject]@{ 'hardware-address' = 'AA:BB:CC:00:11:01'
        'ip-addresses' = @([pscustomobject]@{ 'ip-address-type' = 'ipv4'; 'ip-address' = '169.254.5.5' }) }
    [pscustomobject]@{ 'hardware-address' = 'AA:BB:CC:00:11:02'
        'ip-addresses' = @([pscustomobject]@{ 'ip-address-type' = 'ipv4'; 'ip-address' = '10.0.11.5' }) }
)
$netE1 = Get-ProxmoxVmNetworkData -Node 'n1' -VmId '1100' -RawState 'running'
Assert (@($netE1.IPv4Addresses) -contains '169.254.5.5') "E1: a link-local address is reported (not stripped) -- visible for diagnosing a DHCP failure"
Assert ($netE1.IPv4Addresses[0] -eq '10.0.11.5') "E2: ...but a real address always sorts first, regardless of the agent's own reporting order"

# E3-E6: end-to-end through the clone-aware flow -- link-local alone must NOT complete a
# tracked clone; adding a real address must.
$script:vmState['1101'] = 'running'; $script:vmName['1101'] = 'vm1101'
$script:agentMode['1101'] = @([pscustomobject]@{ 'hardware-address' = 'AA:BB:CC:00:11:03'
        'ip-addresses' = @([pscustomobject]@{ 'ip-address-type' = 'ipv4'; 'ip-address' = '169.254.9.9' }) })
$script:ClusterResourcesCache = $null; $script:ClusterResourcesCachedAt = $null
Set-CloneStateEntry -VmId '1101' -Entry @{
    type = 'clone'; task_id = 'UPID:n1:clone1101'; source_id = '153'; clone_id = '1101'
    name = 'link-local-test'; full = $true
    start_issued = $true; start_task_id = $null; start_pending = $false
    start_retry_count = 1; creation_completed = $false
    clone_started_at = (Get-Date).ToString('o')
}
$script:taskStatus['UPID:n1:clone1101'] = @{ status = 'stopped'; exitstatus = 'OK' }

$llGuest = Get-RasGuestObjectForCloneAwareFlow -VmId '1101'
Assert (@($llGuest.ip_addresses) -contains '169.254.9.9') "E3: the tracked clone's link-local address is still surfaced in ip_addresses"
Assert ($llGuest.state -eq 'powering_on') "E4: ...but a link-local-only guest is NOT treated as clone-ready -- still powering_on"
Assert ($null -ne (Get-CloneStateEntry -VmId '1101')) "E5: ...so its tracking is NOT cleared yet"

$script:agentMode['1101'] = @(
    [pscustomobject]@{ 'hardware-address' = 'AA:BB:CC:00:11:03'
        'ip-addresses' = @([pscustomobject]@{ 'ip-address-type' = 'ipv4'; 'ip-address' = '169.254.9.9' }) }
    [pscustomobject]@{ 'hardware-address' = 'AA:BB:CC:00:11:04'
        'ip-addresses' = @([pscustomobject]@{ 'ip-address-type' = 'ipv4'; 'ip-address' = '10.0.11.9' }) }
)
$readyGuest = Get-RasGuestObjectForCloneAwareFlow -VmId '1101'
Assert ($readyGuest.state -eq 'powered_on') "E6: once a real address shows up alongside the link-local one, the clone completes"
Assert ($readyGuest.ip -eq '10.0.11.9') "E7: ...and the primary 'ip' field is the real address, not the link-local one"
Assert ($null -eq (Get-CloneStateEntry -VmId '1101')) "E8: ...so tracking is now cleared -- clone genuinely done"

# ------------------------------------------------------------
# F. Handle-TaskInfo's simplification to spec.
# ------------------------------------------------------------
# F1: a genuinely still-running real task passes straight through as 'running', for a
# task with NO clone context at all (an ordinary tasks/get for a non-clone action).
$script:taskStatus['UPID:n1:plainrunning'] = @{ status = 'running'; exitstatus = $null }
$rF1 = Handle-TaskInfo -Params ([pscustomobject]@{ id = 'UPID:n1:plainrunning' })
Assert ($rF1.result.state -eq 'running') "F1: a running task (no clone context) passes straight through as 'running'"

# F2: a failed real task passes straight through as 'failed' with the error.
$script:taskStatus['UPID:n1:plainfailed'] = @{ status = 'stopped'; exitstatus = 'unable to create VM 999: not enough free disk space' }
$rF2 = Handle-TaskInfo -Params ([pscustomobject]@{ id = 'UPID:n1:plainfailed' })
Assert ($rF2.result.state -eq 'failed') "F2: a failed real task passes straight through as 'failed'"
Assert ($rF2.result.error.message -match 'not enough free disk space') "F3: ...carrying Proxmox's own failure reason"

# F4: a genuinely completed FULL-clone task reports 'completed' with output.clone_id the
# instant Proxmox itself is done -- with ZERO guest-readiness check. Prove the "zero
# guest-readiness check" half concretely: the underlying VM is still powered OFF (no IP
# at all) at this exact moment, yet tasks/get must still say completed regardless.
$script:PipelinedCloneCompletionEnabled = $false   # isolate this case from pipelining entirely
$script:vmState['1201'] = 'stopped'; $script:vmName['1201'] = 'VM 1201'
Set-CloneStateEntry -VmId '1201' -Entry @{
    type = 'clone'; task_id = 'UPID:n1:clone1201'; source_id = '153'; clone_id = '1201'
    name = 'taskinfo-completed-test'; full = $true
    start_issued = $false; start_task_id = $null; start_pending = $false
    start_retry_count = 0; creation_completed = $false
    clone_started_at = (Get-Date).ToString('o')
}
$script:taskStatus['UPID:n1:clone1201'] = @{ status = 'stopped'; exitstatus = 'OK' }
$rF4 = Handle-TaskInfo -Params ([pscustomobject]@{ id = 'UPID:n1:clone1201' })
Assert ($rF4.result.state -eq 'completed') "F4: a full clone's REAL Proxmox task reports completed the instant it finishes -- even though the guest is still powered off with no IP"
Assert ($rF4.result.output.clone_id -eq '1201') "F5: ...and carries the real clone_id"

# F6: pipelining shortcut is scoped to full clones only -- a LINKED clone (ctx.full =
# $false) past the elapsed threshold, with its real task STILL running, must NOT be
# pipelined through -- it reports the true 'running' state.
$script:PipelinedCloneCompletionEnabled = $true
$script:PipelinedCloneCompletionSeconds = 10
$script:MaxConcurrentCloneOperations = 2
$script:vmState['1202'] = 'stopped'; $script:vmName['1202'] = 'VM 1202'
Set-CloneStateEntry -VmId '1202' -Entry @{
    type = 'clone'; task_id = 'UPID:n1:clone1202'; source_id = '153'; clone_id = '1202'
    name = 'linked-no-pipeline-test'; full = $false
    start_issued = $false; start_task_id = $null; start_pending = $false
    start_retry_count = 0; creation_completed = $false
    clone_started_at = ([DateTime]::UtcNow).AddSeconds(-30).ToString('o')   # well past the 10s pipelining threshold
}
$script:taskStatus['UPID:n1:clone1202'] = @{ status = 'running'; exitstatus = $null }
$rF6 = Handle-TaskInfo -Params ([pscustomobject]@{ id = 'UPID:n1:clone1202' })
Assert ($rF6.result.state -eq 'running') "F6: a linked clone (ctx.full=false) is never pipelined, regardless of elapsed time -- it reports the real 'running' state"
# Done with F6's entry -- its real task never confirms completion in this mock, so left
# in place it would itself occupy a concurrency slot and skew Get-ActiveCloneCount for
# every case below.
Remove-CloneStateEntry -VmId '1202'

# F7: the SAME scenario but for a FULL clone -- this time pipelining DOES fire, reporting
# 'completed' even though the real Proxmox task is still genuinely running.
Set-CloneStateEntry -VmId '1203' -Entry @{
    type = 'clone'; task_id = 'UPID:n1:clone1203'; source_id = '153'; clone_id = '1203'
    name = 'full-pipeline-test'; full = $true
    start_issued = $false; start_task_id = $null; start_pending = $false
    start_retry_count = 0; creation_completed = $false
    clone_started_at = ([DateTime]::UtcNow).AddSeconds(-30).ToString('o')
}
$script:taskStatus['UPID:n1:clone1203'] = @{ status = 'running'; exitstatus = $null }
$rF7 = Handle-TaskInfo -Params ([pscustomobject]@{ id = 'UPID:n1:clone1203' })
Assert ($rF7.result.state -eq 'completed') "F7: a full clone past the pipelining threshold reports 'completed' even while the real Proxmox task is still running"
Assert ($rF7.result.output.clone_id -eq '1203') "F8: ...carrying the real clone_id, same as a genuine completion"

# F9: the concurrency gate still applies even past the elapsed threshold -- with the
# limit turned down to 1 and this entry itself counting as the one active clone, there is
# no room for "one more", so pipelining must hold off and report the true 'running' state.
$script:MaxConcurrentCloneOperations = 1
Set-CloneStateEntry -VmId '1204' -Entry @{
    type = 'clone'; task_id = 'UPID:n1:clone1204'; source_id = '153'; clone_id = '1204'
    name = 'full-pipeline-gated-test'; full = $true
    start_issued = $false; start_task_id = $null; start_pending = $false
    start_retry_count = 0; creation_completed = $false
    clone_started_at = ([DateTime]::UtcNow).AddSeconds(-30).ToString('o')
}
$script:taskStatus['UPID:n1:clone1204'] = @{ status = 'running'; exitstatus = $null }
$rF9 = Handle-TaskInfo -Params ([pscustomobject]@{ id = 'UPID:n1:clone1204' })
Assert ($rF9.result.state -eq 'running') "F9: even past the elapsed threshold, pipelining holds off once there is no concurrency room left (max_concurrent_clone_operations=1, this entry itself fills it)"

Write-Host "`n=== t15 Summary: $pass passed, $fail failed ===" -ForegroundColor $(if ($fail -eq 0) { 'Green' } else { 'Red' })
if ($fail -gt 0) { exit 1 }
