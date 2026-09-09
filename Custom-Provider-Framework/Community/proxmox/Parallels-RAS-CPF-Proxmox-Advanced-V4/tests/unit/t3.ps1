$ErrorActionPreference = 'Stop'
$pass = 0
$fail = 0
function Assert {
    param([bool]$Cond, [string]$Msg)
    if ($Cond) { $script:pass++; Write-Host "PASS: $Msg" -ForegroundColor Green }
    else { $script:fail++; Write-Host "FAIL: $Msg" -ForegroundColor Red }
}

# ============================================================
# TEST 1: cache invalidation on clone/control, task_id fix,
#         2c (no hard error for a tracked clone), and the
#         opportunistic sweep -- via a mocked Invoke-ProxmoxApi
#         driving the real Handle-* functions.
# ============================================================
. "$PSScriptRoot/funcs2.ps1"

$script:LogPath        = "$PSScriptRoot/t3.log"
$script:CloneStatePath = "$PSScriptRoot/t3-state.json"
Remove-Item $script:LogPath, $script:CloneStatePath -ErrorAction SilentlyContinue
$script:TrackedCloneVmIds = [System.Collections.Generic.HashSet[string]]::new()
$script:ClusterResourcesCache = $null
$script:ClusterResourcesCachedAt = $null
$script:RecentlyDeletedIds = @{}
$script:AgentFailureTracker = @{}
$script:LastKnownNetworkData = @{}

$script:apiCalls = New-Object System.Collections.ArrayList
$script:nextId = 200
$script:cloneJobs = @{}
$script:vmState = @{}   # vmid -> status
$script:vmName = @{}
$script:vmVisible = @{} # vmid -> $true once "Proxmox" has registered it (simulates real cluster-resources visibility)
$script:agentAnswers = @{} # vmid -> $true (answers) / $false (500s) / not set = 500
$script:vmTags = @{} # vmid -> tag string, mutated by PUT config

function Invoke-ProxmoxApi {
    param([string]$Method, [string]$Path, $Body, [int]$TimeoutSec = 0)
    [void]$script:apiCalls.Add("$Method $Path")

    if ($Path -match '/cluster/nextid') { return @{ data = "$($script:nextId)" } }

    if ($Method -eq 'PUT' -and $Path -match '/qemu/(\d+)/config$') {
        $script:vmTags[$Matches[1]] = [string]$Body.tags
        return @{ data = $null }
    }

    if ($Path -match '/cluster/resources') {
        $list = New-Object System.Collections.ArrayList
        $tmpl = [pscustomobject]@{ type='qemu'; vmid=999; name='tmpl'; node='n1'; status='stopped'; template=1 }
        if ($script:vmTags.ContainsKey('999')) { $tmpl | Add-Member -MemberType NoteProperty -Name tags -Value $script:vmTags['999'] }
        [void]$list.Add($tmpl)
        foreach ($k in $script:vmState.Keys) {
            if (-not $script:vmVisible[$k]) { continue }   # simulates the id-not-yet-resolvable race
            $entry = [pscustomobject]@{ type='qemu'; vmid=[int]$k; name=$script:vmName[$k]; node='n1'; status=$script:vmState[$k]; template=0 }
            if ($script:vmTags.ContainsKey($k)) { $entry | Add-Member -MemberType NoteProperty -Name tags -Value $script:vmTags[$k] }
            [void]$list.Add($entry)
        }
        return @{ data = $list.ToArray() }
    }

    if ($Path -match '/qemu/(\d+)/clone') {
        $newid = [string]$Body.newid
        $tid = "UPID:n1:clone$newid"
        $script:cloneJobs[$tid] = @{ vmid = $newid; done = $false }
        $script:vmState[$newid] = 'stopped'
        $script:vmName[$newid]  = "VM $newid"
        $script:vmVisible[$newid] = $false   # not visible until the test flips it -- simulates the race
        return @{ data = $tid }
    }

    if ($Path -match '/tasks/(.+)/status') {
        $tid = [uri]::UnescapeDataString($Matches[1])
        if ($script:cloneJobs.ContainsKey($tid)) {
            $st = if ($script:cloneJobs[$tid].done) { 'stopped' } else { 'running' }
            $ex = if ($script:cloneJobs[$tid].done) { 'OK' } else { $null }
            return @{ data = [pscustomobject]@{ status = $st; exitstatus = $ex } }
        }
        return @{ data = [pscustomobject]@{ status = 'stopped'; exitstatus = 'OK' } }
    }

    if ($Path -match '/qemu/(\d+)/status/current') {
        $vm = $Matches[1]
        return @{ data = [pscustomobject]@{ status = $script:vmState[$vm]; qmpstatus = $script:vmState[$vm] } }
    }

    if ($Path -match '/qemu/(\d+)/status/start') {
        $vm = $Matches[1]
        $held = $false
        foreach ($j in $script:cloneJobs.Values) { if ($j.vmid -eq $vm -and -not $j.done) { $held = $true } }
        if ($held) { throw "VM is locked (clone)" }
        $script:vmState[$vm] = 'running'
        return @{ data = "UPID:n1:start$vm" }
    }

    if ($Path -match '/qemu/(\d+)/status/stop') {
        $vm = $Matches[1]
        $script:vmState[$vm] = 'stopped'
        return @{ data = "UPID:n1:stop$vm" }
    }

    if ($Method -eq 'DELETE' -and $Path -match '/qemu/(\d+)$') {
        $vm = $Matches[1]
        $script:vmState.Remove($vm) | Out-Null
        $script:vmVisible.Remove($vm) | Out-Null
        return @{ data = "UPID:n1:destroy$vm" }
    }

    if ($Path -match '/agent/network-get-interfaces') {
        $vm = [regex]::Match($Path, '/qemu/(\d+)/').Groups[1].Value
        if ($script:agentAnswers[$vm] -eq $true) {
            return @{ data = @{ result = @(
                [pscustomobject]@{ 'hardware-address' = 'AA:BB:CC:00:00:01'; 'ip-addresses' = @([pscustomobject]@{'ip-address-type'='ipv4'; 'ip-address'='10.0.0.5'}) }
            ) } }
        }
        throw "500 (QEMU guest agent is not running)"
    }
    throw "unmocked: $Method $Path"
}
function Get-ProxmoxNextVmId { return [string]$script:nextId }

Write-Host "`n--- Test A: clone id not yet resolvable in cluster listing (race) ---" -ForegroundColor Cyan
$script:nextId = 200
$r = Handle-GuestClone -Params ([pscustomobject]@{ id = '999'; name = 'w-a' })
Assert ($null -ne $r.result.clone_id) "Handle-GuestClone returned a clone_id"
$ctxHasTaskId = $script:TaskContext[$r.result.task_id].ContainsKey('task_id') -and -not [string]::IsNullOrWhiteSpace([string]$script:TaskContext[$r.result.task_id].task_id)
Assert $ctxHasTaskId "in-memory TaskContext entry carries task_id"

# VM 200 is NOT yet visible in cluster/resources (vmVisible=false) -- an
# unguarded lookup would have Get-ProxmoxVmNode throw "not found in cluster" here.
$g = Handle-GuestGet -Params ([pscustomobject]@{ id = '200' })
Assert (-not $g.ContainsKey('error')) "guests/get for an unresolvable tracked clone does NOT return a hard error"
Assert ($g.result.state -eq 'powering_on') "unresolvable tracked clone reports state=powering_on"
# Proxmox itself keeps a fresh clone under a placeholder name for a short window --
# the stub returned here must report the clone's real requested name ('w-a') instead,
# so RAS can correlate the guest back to the guests/clone request it is waiting on
# even on its very first guests/get for this id.
Assert ($g.result.name -eq 'w-a') "the not-yet-resolvable stub reports the clone's real requested name, not a placeholder"

Write-Host "`n--- Test B: cache invalidation + never long-cache a listing missing a tracked clone ---" -ForegroundColor Cyan
# VM 200 is a tracked clone still invisible in the listing -- Get-ProxmoxClusterVMs
# must refuse to cache that snapshot at all, so every call while it's missing
# re-fetches rather than being able to get stuck serving a stale "still missing"
# answer for the full TTL.
[void](Get-ProxmoxClusterVMs)
$script:apiCalls.Clear()
[void](Get-ProxmoxClusterVMs)
$refetchedBecauseTrackedCloneMissing = ($script:apiCalls -contains 'GET /api2/json/cluster/resources?type=vm')
Assert $refetchedBecauseTrackedCloneMissing "a listing missing a tracked clone is never cached (always re-fetched)"

$script:vmVisible['200'] = $true
$script:apiCalls.Clear()
$g2 = Handle-GuestGet -Params ([pscustomobject]@{ id = '200' })
$refetched = ($script:apiCalls -contains 'GET /api2/json/cluster/resources?type=vm')
Assert $refetched "once VM 200 is visible, the listing is fetched (and now caches normally since nothing tracked is missing)"
Assert ($g2.result.state -eq 'powering_on') "VM 200's own real clone task is still running, so it correctly stays powering_on"

$script:nextId = 201
[void](Handle-GuestClone -Params ([pscustomobject]@{ id = '999'; name = 'w-b' }))
$script:apiCalls.Clear()
$g3 = Handle-GuestGet -Params ([pscustomobject]@{ id = '200' })
$refetched = ($script:apiCalls -contains 'GET /api2/json/cluster/resources?type=vm')
Assert $refetched "guests/clone busts the cluster-resources cache (next guests/get re-fetches)"
# VM 200 is a tracked clone still carrying Proxmox's placeholder name ("VM 200")
# in the underlying listing. ConvertTo-RasGuestObject substitutes its real
# intended name ('w-a', from Test A's guests/clone call) instead of exposing
# that placeholder to RAS -- without this, RAS's clone thread never sees the
# guest reappear under the name it's actually waiting for.
Assert ($g3.result.name -eq 'w-a') "VM 200 resolves via the real listing with its real name substituted for Proxmox's placeholder (got '$($g3.result.name)')"

Write-Host "`n--- Test C: opportunistic sweep ---" -ForegroundColor Cyan
# 201 is now tracked and off; asking about a DIFFERENT guest (200) should also
# advance 201 via the sweep, with zero extra guests/get calls from RAS.
#
# The sweep is rate-limited per vmid
# (see $script:SweepMinIntervalSeconds). VM 201 was incidentally swept once
# already back in Test B (while still invisible in the listing, so that check
# was a no-op) -- clear that throttle record here to simulate the real gap in
# time between then and now, rather than the sub-millisecond gap a unit test
# would otherwise have. Throttling behavior itself is covered by t6.ps1 Test F.
$script:SweepLastCheckedAt.Remove('201')
$script:cloneJobs.Keys | ForEach-Object { $script:cloneJobs[$_].done = $true }  # all real clone jobs finished
$script:vmVisible['201'] = $true   # PVE has now registered the clone in cluster/resources
Reset-ProxmoxClusterCache
[void](Handle-GuestGet -Params ([pscustomobject]@{ id = '200' }))   # unrelated poll
$vm201State = $script:vmState['201']
Assert ($vm201State -eq 'running') "VM 201 (never itself polled) was started by the sweep piggybacked on a guests/get for VM 200"

Write-Host "`n--- Test D: agent quarantine tag + last-known-good IP ---" -ForegroundColor Cyan
$script:AgentQuarantineAfterSeconds = 0   # don't actually wait 60s in a unit test
$script:AgentQuarantineMinFailures = 2
$script:vmState['300'] = 'running'; $script:vmName['300'] = 'agentless'; $script:vmVisible['300'] = $true
$script:agentAnswers['300'] = $false
Reset-ProxmoxClusterCache
[void](Handle-GuestGet -Params ([pscustomobject]@{ id = '300' }))   # failure 1 -- not yet promoted (min failures = 2)
Assert (-not $script:vmTags.ContainsKey('300')) "not tagged after only 1 failure (respects AgentQuarantineMinFailures)"
Reset-ProxmoxClusterCache
[void](Handle-GuestGet -Params ([pscustomobject]@{ id = '300' }))   # failure 2 -- should promote to a tag now
Assert (([string]$script:vmTags['300']) -match 'rasQuarantine') "tags feature: VM tagged rasQuarantine in Proxmox after crossing the threshold"

Reset-ProxmoxClusterCache
$script:apiCalls.Clear()
[void](Handle-GuestGet -Params ([pscustomobject]@{ id = '300' }))
$skippedAgentCall = -not ($script:apiCalls | Where-Object { $_ -match 'agent/network-get-interfaces' })
Assert $skippedAgentCall "once tagged, the agent call is skipped via the tag fast path (no in-memory state needed)"

Write-Host "  (simulating an admin removing the tag manually)"
$script:vmTags.Remove('300')
Reset-ProxmoxClusterCache
$script:apiCalls.Clear()
[void](Handle-GuestGet -Params ([pscustomobject]@{ id = '300' }))
$rechecked = ($script:apiCalls | Where-Object { $_ -match 'agent/network-get-interfaces' })
Assert ($null -ne $rechecked) "removing the tag resumes agent checks on the very next poll"

Write-Host "  (agent recovered; since the tag was already removed above, checking resumed and finds it working)"
$script:agentAnswers['300'] = $true
Reset-ProxmoxClusterCache
$gRecovered = Handle-GuestGet -Params ([pscustomobject]@{ id = '300' })
Assert ($gRecovered.result.ip -eq '10.0.0.5') "reports the real IP once the agent is checked again and found working"

Write-Host "  (by design: while still tagged, this provider never re-probes on its own -- only manual tag removal resumes checking)"
$script:vmTags['300'] = 'rasQuarantine'
Reset-ProxmoxClusterCache
$script:apiCalls.Clear()
[void](Handle-GuestGet -Params ([pscustomobject]@{ id = '300' }))
$noRecheckWhileTagged = -not ($script:apiCalls | Where-Object { $_ -match 'agent/network-get-interfaces' })
Assert $noRecheckWhileTagged "tags feature: a quarantined VM is never re-probed automatically, even if its agent would now answer"

$script:vmState['301'] = 'running'; $script:vmName['301'] = 'flaky'; $script:vmVisible['301'] = $true
$script:agentAnswers['301'] = $true
Reset-ProxmoxClusterCache
$g4 = Handle-GuestGet -Params ([pscustomobject]@{ id = '301' })
Assert ($g4.result.ip -eq '10.0.0.5') "sanity: healthy agent reports its real IP"
$script:agentAnswers['301'] = $false
Reset-ProxmoxClusterCache
$g5 = Handle-GuestGet -Params ([pscustomobject]@{ id = '301' })
Assert ($g5.result.ip -eq '10.0.0.5') "a transient agent failure reports the last known-good IP instead of retracting it"

Write-Host "`n--- Test E: delete clears tracking + guests/list filter ---" -ForegroundColor Cyan
$script:nextId = 400
[void](Handle-GuestClone -Params ([pscustomobject]@{ id = '999'; name = 'w-e' }))
$script:vmVisible['400'] = $true
Reset-ProxmoxClusterCache
Assert ($script:TrackedCloneVmIds.Contains('400')) "sanity: VM 400 is tracked as an in-flight clone before delete"
[void](Handle-GuestControl -Params ([pscustomobject]@{ id = '400'; control = 'delete' }))
Assert (-not $script:TrackedCloneVmIds.Contains('400')) "delete removes the clone-state tracking entry (frees the concurrency slot)"
Reset-ProxmoxClusterCache
$list = Handle-GuestList
Assert (-not ($list.result.guests -contains '400')) "guests/list omits a just-deleted id immediately"
$g6 = Handle-GuestGet -Params ([pscustomobject]@{ id = '400' })
Assert (-not $g6.ContainsKey('error')) "guests/get for a just-deleted id returns a clean result, not an error"

Write-Host "`n--- Test F: rasExclude tag (manual, admin-set) ---" -ForegroundColor Cyan
$script:vmState['500'] = 'running'; $script:vmName['500'] = 'admin-owned-vm'; $script:vmVisible['500'] = $true
$script:vmTags['500'] = 'rasExclude'
Reset-ProxmoxClusterCache
$listF = Handle-GuestList
Assert (-not ($listF.result.guests -contains '500')) "tags feature: rasExclude VM is omitted from guests/list"
$hostsF = Handle-HostList
Assert (-not ($hostsF.result.guests -contains '500')) "tags feature: rasExclude VM is omitted from hosts/list"
$errF = Handle-GuestGet -Params ([pscustomobject]@{ id = '500' })
Assert ($errF.ContainsKey('error')) "tags feature: guests/get for a rasExclude VM behaves exactly like a not-found VM (Get-ProxmoxVmNode enforcement)"

Write-Host "  (admin removes the tag)"
$script:vmTags.Remove('500')
Reset-ProxmoxClusterCache
$listF2 = Handle-GuestList
Assert ($listF2.result.guests -contains '500') "removing rasExclude makes the VM visible again immediately"

Write-Host "`n--- Test G: rasTemplate<id> / rasClone<id> tags (tags feature) ---" -ForegroundColor Cyan
# Use a FRESH source id for the "tagged on first clone" assertion. 999 was
# already cloned from (and thus already tag-ATTEMPTED, successfully) back in
# Test A/B, and the rasTemplate tag write is attempted at most once per
# source per PROCESS lifetime -- not once per clone -- specifically so a
# source that failed (or already succeeded) doesn't keep re-attempting on
# every subsequent clone. So 999 can no longer demonstrate a "fresh" write
# within this same process; a separate assertion below uses 999 to verify
# that once-per-process guard directly.
$script:vmState['700'] = 'stopped'; $script:vmName['700'] = 'tmpl700'; $script:vmVisible['700'] = $true
Reset-ProxmoxClusterCache   # otherwise Get-ProxmoxVmNode below still sees the pre-700 listing cached by Test F
$script:nextId = 600
$rG = Handle-GuestClone -Params ([pscustomobject]@{ id = '700'; name = 'w-g' })
Assert (([string]$script:vmTags['700']) -eq 'rasTemplate700') "tags feature: source VM 700 tagged rasTemplate700 on first clone from it"

$script:vmVisible['600'] = $true
$script:cloneJobs[$rG.result.task_id].done = $true   # real clone job finished
Reset-ProxmoxClusterCache
[void](Handle-GuestGet -Params ([pscustomobject]@{ id = '600' }))   # clone-aware flow issues the start, applying the clone tag alongside it
Assert (([string]$script:vmTags['600']) -eq 'rasClone700') "tags feature: clone VM 600 tagged rasClone700 once its real clone job is confirmed done"

# Cloning a SECOND time from the same template must not re-tag it (idempotent, no redundant PUT)
$script:apiCalls.Clear()
$script:nextId = 601
[void](Handle-GuestClone -Params ([pscustomobject]@{ id = '700'; name = 'w-g2' }))
$redundantTemplateTagWrite = @($script:apiCalls | Where-Object { $_ -match 'PUT .*qemu/700/config' })
Assert ($redundantTemplateTagWrite.Count -eq 0) "tags feature: an already-tagged template is not re-tagged on a second clone (idempotent)"

# Source 999 (tag-attempted back in Test A) proves the per-process
# "attempted once" guard itself -- even if its tag disappears out from under
# the provider (e.g. an admin edit), a later clone from it must NOT retry the
# write, by design.
$script:vmTags.Remove('999')
Reset-ProxmoxClusterCache
$script:apiCalls.Clear()
$script:nextId = 602
[void](Handle-GuestClone -Params ([pscustomobject]@{ id = '999'; name = 'w-g3' }))
$noRetryAfterAttempted = @($script:apiCalls | Where-Object { $_ -match 'PUT .*qemu/999/config' })
Assert ($noRetryAfterAttempted.Count -eq 0) "a source already attempted this process lifetime is never retried even if its tag disappears"
Assert (-not $script:vmTags.ContainsKey('999')) "confirms the tag write really was skipped (mock tag state stays absent)"

Write-Host "`n=== Summary: $pass passed, $fail failed ===" -ForegroundColor $(if ($fail -eq 0) { 'Green' } else { 'Red' })
if ($fail -gt 0) { exit 1 }
