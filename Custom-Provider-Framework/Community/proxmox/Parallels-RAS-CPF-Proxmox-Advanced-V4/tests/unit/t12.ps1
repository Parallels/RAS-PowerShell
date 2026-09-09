$ErrorActionPreference = 'Stop'
$pass = 0
$fail = 0
function Assert {
    param([bool]$Cond, [string]$Msg)
    if ($Cond) { $script:pass++; Write-Host "PASS: $Msg" -ForegroundColor Green }
    else { $script:fail++; Write-Host "FAIL: $Msg" -ForegroundColor Red }
}

# ============================================================
# TESTS for two small, cosmetic-adjacent fixes to guest-agent probing and log wording:
#
#   Skip the guest-agent network probe for a short window after THIS provider
#   itself issues a 'stop' -- RAS often polls again within ~80ms of the stop
#   POST, and Proxmox's own status/current can still say 'running' in that
#   window even though qemu is already tearing down, so probing the agent
#   reliably produced a 500 ("VM <id> is not running"). The provider already
#   prefers the live status read for POWER STATE (correct precedence, left
#   untouched here) -- this only avoids the doomed HTTP round trip.
#
#   TRACKED CLONE FOUND IN FILE -> FOUND VIA CLONE-STATE CACHE. The old wording
#   read like a disk read (and a high cache-miss rate) when both paths are
#   actually served from memory -- a log-message-only fix.
# ============================================================
. "$PSScriptRoot/funcs2.ps1"

$script:LogPath        = "$PSScriptRoot/t12.log"
$script:CloneStatePath = "$PSScriptRoot/t12-state.json"
Remove-Item $script:LogPath, $script:CloneStatePath -ErrorAction SilentlyContinue
$script:TrackedCloneVmIds = [System.Collections.Generic.HashSet[string]]::new()
$script:ClusterResourcesCache = $null
$script:ClusterResourcesCachedAt = $null
$script:RecentlyDeletedIds = @{}
$script:RecentlyControlledIds = @{}
$script:RecentlyStoppedIds = @{}
$script:AgentFailureTracker = @{}
$script:LastKnownNetworkData = @{}
$script:CloneStateMemory = $null
$script:TaskContext = @{}
$script:RecentlyStoppedRetentionSeconds = 10

$script:apiCalls = New-Object System.Collections.ArrayList
$script:agentCalls = New-Object System.Collections.ArrayList   # vmids the agent was actually probed for
$script:vmState = @{ '500' = 'running' }
$script:vmName  = @{ '500' = 'vm500' }

function Invoke-ProxmoxApi {
    param([string]$Method, [string]$Path, $Body, [int]$TimeoutSec = 0)
    [void]$script:apiCalls.Add("$Method $Path")

    if ($Path -match '/cluster/resources') {
        $list = New-Object System.Collections.ArrayList
        foreach ($k in $script:vmState.Keys) {
            [void]$list.Add([pscustomobject]@{ type = 'qemu'; vmid = [int]$k; name = $script:vmName[$k]; node = 'n1'; status = $script:vmState[$k]; template = 0 })
        }
        return @{ data = $list.ToArray() }
    }

    if ($Path -match '/qemu/(\d+)/status/current$') {
        $vmid = $Matches[1]
        $st = if ($script:vmState.ContainsKey($vmid)) { $script:vmState[$vmid] } else { 'stopped' }
        return @{ data = [pscustomobject]@{ status = $st; qmpstatus = $st } }
    }

    if ($Path -match '/qemu/(\d+)/status/(start|stop|shutdown)$') {
        $vmid = $Matches[1]; $action = $Matches[2]
        $script:vmState[$vmid] = if ($action -eq 'start') { 'running' } else { 'stopped' }
        return @{ data = "UPID:n1:qm$($action)$vmid" }
    }

    if ($Path -match '/qemu/(\d+)/agent/network-get-interfaces$') {
        $vmid = $Matches[1]
        [void]$script:agentCalls.Add($vmid)
        # Real Proxmox's own failure mode under test: the agent call itself 500s once
        # the VM has actually stopped by the time this HTTP call reaches Proxmox.
        if ($script:vmState[$vmid] -ne 'running') {
            throw "Response status code does not indicate success: 500 (VM $vmid is not running)."
        }
        $iface = [pscustomobject]@{ 'hardware-address' = 'AA:BB:CC:00:00:01'
            'ip-addresses' = @([pscustomobject]@{ 'ip-address-type' = 'ipv4'; 'ip-address' = '10.0.0.50' }) }
        return @{ data = [pscustomobject]@{ result = @($iface) } }
    }

    throw "Unexpected call in mock: $Method $Path"
}

# ------------------------------------------------------------
# A. Baseline -- a running VM we did NOT just stop probes normally.
# ------------------------------------------------------------
$script:agentCalls.Clear()
$net = Get-ProxmoxVmNetworkData -Node 'n1' -VmId '500' -RawState 'running'
Assert ($script:agentCalls -contains '500') "A1: a running VM with no recent stop still gets probed"
Assert ($net.IPv4Addresses -contains '10.0.0.50') "A2: ...and its real IP is reported"

# ------------------------------------------------------------
# B. The race: we just issued 'stop', but the live status read (passed in as
#    RawState, exactly as ConvertTo-RasGuestObject would) still says 'running' for
#    this one brief window.
# ------------------------------------------------------------
$stopResult = Handle-GuestControl -Params ([pscustomobject]@{ id = '500'; control = 'stop' })
Assert (-not ($stopResult -is [hashtable] -and $stopResult.ContainsKey('error'))) "B1: guests/control(stop) succeeds"
Assert (Test-ProxmoxRecentlyStopped -VmId '500') "B2: the stop is tracked as recent"

$script:agentCalls.Clear()
$script:vmState['500'] = 'running'   # force the race: Proxmox's own status/current still lags
$net2 = Get-ProxmoxVmNetworkData -Node 'n1' -VmId '500' -RawState 'running'
Assert (-not ($script:agentCalls -contains '500')) "B3: the guest-agent probe is skipped entirely right after our own stop, even though RawState still says running"
Assert (@($net2.IPv4Addresses).Count -eq 0) "B4: ...and no stale IP is reported for it"

# ------------------------------------------------------------
# C. A quick restart clears the marker immediately -- probing must resume at once,
#    not stay suppressed for the rest of the (short) window.
# ------------------------------------------------------------
[void](Handle-GuestControl -Params ([pscustomobject]@{ id = '500'; control = 'start' }))
Assert (-not (Test-ProxmoxRecentlyStopped -VmId '500')) "C1: a subsequent 'start' clears the recently-stopped marker right away"

$script:agentCalls.Clear()
$net3 = Get-ProxmoxVmNetworkData -Node 'n1' -VmId '500' -RawState 'running'
Assert ($script:agentCalls -contains '500') "C2: ...so the probe resumes immediately, not after waiting out the old window"
Assert ($net3.IPv4Addresses -contains '10.0.0.50') "C3: ...and reports the real IP again"

# ------------------------------------------------------------
# D. The marker still self-expires on its own (same shape as every other Recently*
#    tracker), so a provider that somehow misses the clearing path is not stuck.
# ------------------------------------------------------------
[void](Handle-GuestControl -Params ([pscustomobject]@{ id = '500'; control = 'stop' }))
Assert (Test-ProxmoxRecentlyStopped -VmId '500') "D1: setup -- stop tracked"
$script:RecentlyStoppedIds['500'] = ([DateTime]::UtcNow).AddSeconds(-($script:RecentlyStoppedRetentionSeconds + 1))
Assert (-not (Test-ProxmoxRecentlyStopped -VmId '500')) "D2: the marker expires on its own past its retention window"
Assert (-not $script:RecentlyStoppedIds.ContainsKey('500')) "D3: ...and is cleaned out of the tracking table"

$script:vmState['500'] = 'running'
$script:agentCalls.Clear()
$net4 = Get-ProxmoxVmNetworkData -Node 'n1' -VmId '500' -RawState 'running'
Assert ($script:agentCalls -contains '500') "D4: once expired, the probe resumes even without an explicit start"

# ------------------------------------------------------------
# E. Delete cleans the marker up too (Clear-ProxmoxTrackingForVm), so a recycled
#    VMID never inherits a stale suppression from whatever used it before.
# ------------------------------------------------------------
[void](Handle-GuestControl -Params ([pscustomobject]@{ id = '500'; control = 'stop' }))
Assert (Test-ProxmoxRecentlyStopped -VmId '500') "E1: setup -- stop tracked"
Clear-ProxmoxTrackingForVm -VmId '500'
Assert (-not (Test-ProxmoxRecentlyStopped -VmId '500')) "E2: Clear-ProxmoxTrackingForVm (called on delete) clears the recently-stopped marker too"

# ------------------------------------------------------------
# F. Log wording (#7): the clone-state-cache hit is no longer described as a file
#    read.
# ------------------------------------------------------------
Set-CloneStateEntry -VmId '700' -Entry @{
    type = 'clone'; task_id = 'UPID:n1:clone700'; source_id = '153'; clone_id = '700'
    name = 'wording-test'; full = $true; clone_node = 'n1'
    start_issued = $false; start_task_id = $null; start_pending = $false
    start_retry_count = 0; creation_completed = $false
    clone_started_at = (Get-Date).ToString('o')
}
[void](Get-TrackedCloneContextByVmId -VmId '700')
$logText = Get-Content $script:LogPath -Raw
Assert ($logText -match 'FOUND VIA CLONE-STATE CACHE for VM \[700\]') "F1: the clone-state-cache hit logs the new, accurate wording"
Assert ($logText -notmatch 'FOUND IN FILE') "F2: the old 'FOUND IN FILE' wording is gone"
Remove-CloneStateEntry -VmId '700'

Write-Host "`n=== t12 Summary: $pass passed, $fail failed ===" -ForegroundColor $(if ($fail -eq 0) { 'Green' } else { 'Red' })
if ($fail -gt 0) { exit 1 }
