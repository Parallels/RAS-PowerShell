$ErrorActionPreference = 'Stop'
$pass = 0
$fail = 0
function Assert {
    param([bool]$Cond, [string]$Msg)
    if ($Cond) { $script:pass++; Write-Host "PASS: $Msg" -ForegroundColor Green }
    else { $script:fail++; Write-Host "FAIL: $Msg" -ForegroundColor Red }
}

# ============================================================
# TESTS for the RAS template-maintenance-mode fixes traced out of a captured
# live sequence (see MAINTENANCE-MODE.md for the full trace):
#
#   Observed: exit-maintenance converted VM 153 to a template at
#   10:39:18 and Proxmox accepted it, but guests/get kept reporting
#   is_template=false for another ~50s (the cluster/resources cache still
#   held a snapshot taken BEFORE the convert). RAS read that as "the
#   convert didn't take", re-issued exit-maintenance at 10:39:35, and
#   Proxmox answered the second POST with
#   500 "you can't convert a template to a template" -- so RAS reported
#   the whole maintenance exit as failed even though it had succeeded.
#
# Fixes under test:
#   M1  Handle-GuestConvert is idempotent -- reads the LIVE template flag
#       from the VM's own config and no-ops when already in the requested
#       state; also swallows Proxmox's "template to a template" 500
#       specifically (and only that one) as success.
#   M2  is_template is resolved from qemu/{id}/status/current (or, failing
#       that, from the convert we just performed) for a window after a
#       convert, instead of the lagging cluster listing.
# ============================================================
. "$PSScriptRoot/funcs2.ps1"

$script:LogPath        = "$PSScriptRoot/t9.log"
$script:CloneStatePath = "$PSScriptRoot/t9-state.json"
Remove-Item $script:LogPath, $script:CloneStatePath -ErrorAction SilentlyContinue
$script:TrackedCloneVmIds = [System.Collections.Generic.HashSet[string]]::new()
$script:ClusterResourcesCache = $null
$script:ClusterResourcesCachedAt = $null
$script:RecentlyDeletedIds = @{}
$script:RecentlyControlledIds = @{}
$script:RecentlyConvertedIds = @{}
$script:MaintenanceModeVmIds = [System.Collections.Generic.HashSet[string]]::new()
$script:AgentFailureTracker = @{}
$script:LastKnownNetworkData = @{}
$script:SweepLastCheckedAt = @{}
$script:RasTemplateTagAttempted = [System.Collections.Generic.HashSet[string]]::new()
$script:CloneStateMemory = $null
$script:TaskContext = @{}
$script:CompletedCloneTaskOutputs = @{}
$script:CloneTaskCompletionCache = @{}
$script:LastGuestPollAt = @{}
$script:OrphanAuditLastCheckedAt = [DateTime]::MinValue
$script:RecentlyConvertedRetentionSeconds = 120
$script:RecentlyControlledRetentionSeconds = 120
$script:ClusterResourcesCacheTtlSeconds = 30

$script:apiCalls = New-Object System.Collections.ArrayList

# Two separate views of the SAME flag, exactly as Proxmox behaves:
#   liveTemplate    -- the VM's own config / status/current (authoritative)
#   clusterTemplate -- what cluster/resources reports (lags by up to ~10s
#                      server-side, on top of this provider's own cache TTL)
$script:liveTemplate    = @{}
$script:clusterTemplate = @{}
$script:vmState         = @{}
# When set, the next POST /template throws this instead of succeeding, so the
# "flag flipped between our read and our write" race can be exercised.
$script:forceTemplatePostError = $null
# When set, status/current omits 'template' entirely (older PVE / unexpected
# response shape) so the fallback-to-recorded-intent path can be exercised.
$script:statusOmitsTemplate = $false

function Reset-Mock {
    $script:apiCalls.Clear()
    $script:liveTemplate    = @{ '153' = $true }
    $script:clusterTemplate = @{ '153' = $true }
    $script:vmState         = @{ '153' = 'stopped' }
    $script:ClusterResourcesCache = $null
    $script:ClusterResourcesCachedAt = $null
    $script:RecentlyConvertedIds = @{}
    $script:MaintenanceModeVmIds = [System.Collections.Generic.HashSet[string]]::new()
    $script:RecentlyControlledIds = @{}
    $script:LastKnownNetworkData = @{}
    $script:forceTemplatePostError = $null
    $script:statusOmitsTemplate = $false
}

function Invoke-ProxmoxApi {
    param([string]$Method, [string]$Path, $Body, [int]$TimeoutSec = 0)
    [void]$script:apiCalls.Add("$Method $Path")

    if ($Path -match '/cluster/resources') {
        $list = New-Object System.Collections.ArrayList
        foreach ($k in $script:vmState.Keys) {
            [void]$list.Add([pscustomobject]@{
                    type     = 'qemu'
                    vmid     = [int]$k
                    node     = 'n1'
                    name     = "vm$k"
                    status   = $script:vmState[$k]
                    template = $(if ($script:clusterTemplate[$k]) { 1 } else { 0 })
                })
        }
        return @{ data = $list.ToArray() }
    }

    if ($Method -eq 'GET' -and $Path -match '/qemu/(\d+)/config$') {
        $vmid = $Matches[1]
        return @{ data = [pscustomobject]@{ template = $(if ($script:liveTemplate[$vmid]) { 1 } else { 0 }) } }
    }

    if ($Method -eq 'PUT' -and $Path -match '/qemu/(\d+)/config$') {
        $vmid = $Matches[1]
        if ($null -ne $Body -and $Body.ContainsKey('template')) {
            $script:liveTemplate[$vmid] = ([int]$Body.template -eq 1)
            # cluster/resources deliberately NOT updated -- that is the lag.
        }
        return @{ data = $null }
    }

    if ($Method -eq 'POST' -and $Path -match '/qemu/(\d+)/template$') {
        $vmid = $Matches[1]
        if ($null -ne $script:forceTemplatePostError) {
            $msg = $script:forceTemplatePostError
            $script:forceTemplatePostError = $null
            # Proxmox can only raise this particular error when the VM really
            # IS already a template, so the mock has to reflect that -- the
            # whole point of the race is that someone else converted it.
            if ($msg -match "template to a template") { $script:liveTemplate[$vmid] = $true }
            throw $msg
        }
        if ($script:liveTemplate[$vmid]) {
            throw "Response status code does not indicate success: 500 (you can't convert a template to a template)."
        }
        $script:liveTemplate[$vmid] = $true
        return @{ data = "UPID:n1:qmtemplate$vmid" }
    }

    if ($Path -match '/qemu/(\d+)/status/current$') {
        $vmid = $Matches[1]
        $st = if ($script:vmState.ContainsKey($vmid)) { $script:vmState[$vmid] } else { 'stopped' }
        $obj = [pscustomobject]@{ status = $st; qmpstatus = $st }
        if (-not $script:statusOmitsTemplate) {
            $obj | Add-Member -MemberType NoteProperty -Name template -Value $(if ($script:liveTemplate[$vmid]) { 1 } else { 0 })
        }
        return @{ data = $obj }
    }

    if ($Path -match '/tasks/(.+)/status$') {
        return @{ data = [pscustomobject]@{ status = 'stopped'; exitstatus = 'OK' } }
    }

    throw "Unexpected call in mock: $Method $Path"
}

function Count-Calls { param([string]$Pattern) return @($script:apiCalls | Where-Object { $_ -match $Pattern }).Count }

# The handlers return plain hashtables with only ONE of these keys present,
# and Set-StrictMode makes a bare property read on the missing one throw.
function HasError  { param($R) return ($R -is [hashtable] -and $R.ContainsKey('error')) }
function HasResult { param($R) return ($R -is [hashtable] -and $R.ContainsKey('result')) }

# ------------------------------------------------------------
# A. Exit maintenance on a VM that is NOT yet a template: the convert
#    happens, and guests/get reports the NEW flag straight away even though
#    cluster/resources still carries the old one.
# ------------------------------------------------------------
Reset-Mock
$script:liveTemplate['153'] = $false
$script:clusterTemplate['153'] = $false

$r = Handle-GuestConvert -Params ([pscustomobject]@{ id = '153'; is_template = $true })
Assert ((HasResult $r) -and -not (HasError $r)) 'A1: convert VM->template succeeds'
Assert ((Count-Calls 'POST /api2/json/nodes/n1/qemu/153/template') -eq 1) 'A2: exactly one template POST issued'
Assert ($script:liveTemplate['153'] -eq $true) 'A3: Proxmox-side flag actually flipped'

# cluster/resources still lies -- this is the live-observed condition.
Assert ($script:clusterTemplate['153'] -eq $false) 'A4: cluster listing is (deliberately) still stale'
$g = ConvertTo-RasGuestObject -VmId '153'
Assert ($g.is_template -eq $true) 'A5: guests/get reports the NEW flag despite the stale cluster listing'

# ------------------------------------------------------------
# B. The live-observed failure, replayed end to end: RAS re-issues exit
#    maintenance while its own view still says "not a template".
#    Before the fix this returned a JSON-RPC error and RAS reported
#    "failed to exit maintenance".
# ------------------------------------------------------------
Reset-Mock
$script:liveTemplate['153'] = $false
$script:clusterTemplate['153'] = $false

[void](Handle-GuestConvert -Params ([pscustomobject]@{ id = '153'; is_template = $true }))
$script:apiCalls.Clear()
$r = Handle-GuestConvert -Params ([pscustomobject]@{ id = '153'; is_template = $true })
Assert (-not (HasError $r)) 'B1: repeated exit-maintenance convert does NOT error'
Assert ((HasResult $r) -and -not [string]::IsNullOrWhiteSpace([string]$r.result.task_id)) 'B2: it returns a task id RAS can poll'
Assert ((Count-Calls 'POST /api2/json/nodes/n1/qemu/153/template') -eq 0) 'B3: no redundant template POST reaches Proxmox'
Assert ((Count-Calls 'GET /api2/json/nodes/n1/qemu/153/config') -ge 1) 'B4: the live config was consulted to decide that'

# ------------------------------------------------------------
# C. Race safety net: our live read says "not a template", but something
#    else converts it before our POST lands. Only THAT Proxmox message is
#    swallowed, and only into the state RAS asked for.
# ------------------------------------------------------------
Reset-Mock
$script:liveTemplate['153'] = $false
$script:clusterTemplate['153'] = $false
$script:forceTemplatePostError = "Response status code does not indicate success: 500 (you can't convert a template to a template)."

$r = Handle-GuestConvert -Params ([pscustomobject]@{ id = '153'; is_template = $true })
Assert (-not (HasError $r)) 'C1: a raced "template to a template" 500 is treated as success'
$g = ConvertTo-RasGuestObject -VmId '153'
Assert ($g.is_template -eq $true) 'C2: ...and the guest is reported as a template afterwards'

# C3: every OTHER failure must still surface.
Reset-Mock
$script:liveTemplate['153'] = $false
$script:clusterTemplate['153'] = $false
$script:forceTemplatePostError = 'Response status code does not indicate success: 500 (unable to create template, because VM contains snapshots).'
$r = Handle-GuestConvert -Params ([pscustomobject]@{ id = '153'; is_template = $true })
Assert ((HasError $r)) 'C3: an unrelated convert failure is still reported as an error'
Assert ([string]$r.error.message -match 'snapshots') 'C4: ...with the original Proxmox message preserved'

# ------------------------------------------------------------
# D. Enter maintenance: template -> VM via PUT config template=0.
#    This is the direction the Proxmox web UI does not expose at all.
# ------------------------------------------------------------
Reset-Mock   # 153 starts as a template
$r = Handle-GuestConvert -Params ([pscustomobject]@{ id = '153'; is_template = $false })
Assert (-not (HasError $r)) 'D1: convert template->VM succeeds'
Assert ((Count-Calls 'PUT /api2/json/nodes/n1/qemu/153/config') -eq 1) 'D2: it goes out as a config PUT'
Assert ($script:liveTemplate['153'] -eq $false) 'D3: Proxmox-side flag actually cleared'
Assert ($script:clusterTemplate['153'] -eq $true) 'D4: cluster listing is (deliberately) still stale'
$g = ConvertTo-RasGuestObject -VmId '153'
Assert ($g.is_template -eq $false) 'D5: guests/get reports NOT-a-template immediately'

# D6: and repeating it is a no-op rather than a redundant write.
$script:apiCalls.Clear()
$r = Handle-GuestConvert -Params ([pscustomobject]@{ id = '153'; is_template = $false })
Assert (-not (HasError $r)) 'D6: repeated enter-maintenance convert does not error'
Assert ((Count-Calls 'PUT /api2/json/nodes/n1/qemu/153/config') -eq 0) 'D7: ...and issues no redundant config PUT'

# ------------------------------------------------------------
# E. Precedence: a live status/current reading beats the recorded intent.
#    If Proxmox says otherwise, we believe Proxmox, not our own bookkeeping.
# ------------------------------------------------------------
Reset-Mock
$script:liveTemplate['153'] = $false
$script:clusterTemplate['153'] = $false
[void](Handle-GuestConvert -Params ([pscustomobject]@{ id = '153'; is_template = $true }))
# Something outside this provider reverts it.
$script:liveTemplate['153'] = $false
$g = ConvertTo-RasGuestObject -VmId '153'
Assert ($g.is_template -eq $false) 'E1: a live status/current reading overrides our recorded intent'

# E2: when status/current carries no template field at all, fall back to
#     what we just asked Proxmox to do rather than to the stale listing.
Reset-Mock
$script:liveTemplate['153'] = $false
$script:clusterTemplate['153'] = $false
[void](Handle-GuestConvert -Params ([pscustomobject]@{ id = '153'; is_template = $true }))
$script:statusOmitsTemplate = $true
$g = ConvertTo-RasGuestObject -VmId '153'
Assert ($g.is_template -eq $true) 'E2: with no template field in status/current, the recorded intent is used'

# ------------------------------------------------------------
# F. The override is a short window, not a permanent shadow state.
# ------------------------------------------------------------
Reset-Mock
$script:liveTemplate['153'] = $false
$script:clusterTemplate['153'] = $false
[void](Handle-GuestConvert -Params ([pscustomobject]@{ id = '153'; is_template = $true }))
Assert (Test-ProxmoxRecentlyConverted -VmId '153') 'F1: the convert is tracked right after it happens'

# Age the entry past its retention.
$script:RecentlyConvertedIds['153'].converted_at = ([DateTime]::UtcNow).AddSeconds(-($script:RecentlyConvertedRetentionSeconds + 5))
Assert (-not (Test-ProxmoxRecentlyConverted -VmId '153')) 'F2: it expires on its own'
Assert (-not $script:RecentlyConvertedIds.ContainsKey('153')) 'F3: ...and is cleaned out of the tracking table'

# Once expired, the (by then caught-up) cluster listing is authoritative again.
$script:clusterTemplate['153'] = $true
$script:ClusterResourcesCache = $null
$g = ConvertTo-RasGuestObject -VmId '153'
Assert ($g.is_template -eq $true) 'F4: after expiry the cluster listing is used again'

# ------------------------------------------------------------
# G. No per-poll cost regression: a guest that was neither controlled nor
#    converted recently must NOT trigger an extra status/current call.
# ------------------------------------------------------------
Reset-Mock
$script:vmState['200'] = 'stopped'
$script:liveTemplate['200'] = $false
$script:clusterTemplate['200'] = $false
$script:ClusterResourcesCache = $null
[void](ConvertTo-RasGuestObject -VmId '200')
$script:apiCalls.Clear()
[void](ConvertTo-RasGuestObject -VmId '200')
Assert ((Count-Calls 'status/current') -eq 0) 'G1: an untouched guest costs no extra status/current call'
Assert ((Count-Calls 'qemu/200/config') -eq 0) 'G2: ...and no extra config call either'

# ------------------------------------------------------------
# H. A convert busts the cluster cache, so the next listing is refetched
#    rather than served from a snapshot taken before the convert.
# ------------------------------------------------------------
Reset-Mock
$script:liveTemplate['153'] = $false
$script:clusterTemplate['153'] = $false
[void](Get-ProxmoxClusterVMs)
Assert ($null -ne $script:ClusterResourcesCache) 'H1: cluster listing is cached to begin with'
[void](Handle-GuestConvert -Params ([pscustomobject]@{ id = '153'; is_template = $true }))
Assert ($null -eq $script:ClusterResourcesCache) 'H2: the convert invalidated the cached listing'

# ------------------------------------------------------------
# J. Maintenance mode outlives any time window.
#
#    RAS exits maintenance by shutting the guest down through the in-guest
#    RAS agent -- NOT through guests/control -- so there is no control
#    action for Test-ProxmoxRecentlyControlled to key off. RAS then waits
#    for THIS provider's guests/get to report powered_off before it issues
#    the convert. A maintenance session lasts as long as the admin needs
#    (patching a template takes minutes), so by the time the guest is shut
#    down both the recently-controlled and recently-converted windows can
#    have expired, leaving power state to come from the 30s-stale cluster
#    listing -- which delays or strands the convert.
# ------------------------------------------------------------
Reset-Mock   # 153 starts as a template

# Enter maintenance: de-templated, then booted.
[void](Handle-GuestConvert -Params ([pscustomobject]@{ id = '153'; is_template = $false }))
Assert (Test-ProxmoxInMaintenanceMode -VmId '153') 'J1: de-templating marks the VM as in maintenance'

# Age BOTH windows well past their retention, as a real session would.
$script:RecentlyConvertedIds['153'].converted_at = ([DateTime]::UtcNow).AddSeconds(-3600)
$script:RecentlyControlledIds = @{}
Assert (-not (Test-ProxmoxRecentlyConverted -VmId '153')) 'J2: the convert window has expired'
Assert (Test-ProxmoxInMaintenanceMode -VmId '153') 'J3: ...but the maintenance marker has not'

# The guest is shut down from inside; the cluster listing has not caught up.
$script:vmState['153'] = 'running'
$script:ClusterResourcesCache = $null
[void](Get-ProxmoxClusterVMs)          # cache the 'running' snapshot
$script:vmState['153'] = 'stopped'     # the real shutdown
$g = ConvertTo-RasGuestObject -VmId '153'
Assert ($g.power_state -eq 'stopped') 'J4: power state is read live, so RAS sees the shutdown at once'

# Exiting maintenance clears the marker again.
[void](Handle-GuestConvert -Params ([pscustomobject]@{ id = '153'; is_template = $true }))
Assert (-not (Test-ProxmoxInMaintenanceMode -VmId '153')) 'J5: re-templating clears the maintenance marker'

# J6: and a VM that never entered maintenance is unaffected.
Reset-Mock
Assert (-not (Test-ProxmoxInMaintenanceMode -VmId '153')) 'J6: an untouched template is not marked as in maintenance'

# ------------------------------------------------------------
# I. Orphan-audit false positives, both observed and latent.
#
#    Observed live at 10:37:13.827: VMs 136/137/138 were deleted by this
#    provider ~0.6s earlier, were still present in the lagging cluster
#    listing, and had just had their poll records dropped by
#    Clear-ProxmoxTrackingForVm -- so "never polled" read as maximally
#    stale and all three were flagged. The best-effort tag write then
#    failed with 500 "Configuration file ... does not exist" (3x).
#
#    Latent: every poll record is per-process and starts empty, so before
#    the fix the FIRST audit after a provider restart would flag every
#    rasClone-tagged VM in the cluster at once.
# ------------------------------------------------------------
$script:RasCloneTagPrefix = 'rasClone'
$script:RasOrphanCandidateTag = 'rasOrphanCandidate'
$script:OrphanDetectionEnabled = $true
$script:OrphanDetectionCheckIntervalSeconds = 300
$script:OrphanDetectionStalePollAfterSeconds = 180
$script:RecentlyDeletedRetentionSeconds = 300
$script:TagWriteTimeoutSeconds = 3

# Returns a COUNT, not a collection -- PowerShell unrolls a returned array,
# so a single-element result would lose its .Count under Set-StrictMode.
function Get-OrphanLineCount {
    if (-not (Test-Path $script:LogPath)) { return 0 }
    return @(Get-Content $script:LogPath | Where-Object { $_ -match 'ORPHAN CANDIDATE' }).Count
}

$orphanVm = [pscustomobject]@{ type = 'qemu'; vmid = 136; node = 'n1'; name = 'w10pwsh-001'; status = 'stopped'; template = 0; tags = 'rasClone153' }

# I1: a VM this provider deleted moments ago must not be flagged, even
#     though it is still in the cluster listing with no poll record.
$before = (Get-OrphanLineCount)
$script:LastGuestPollAt = @{}
$script:RecentlyDeletedIds = @{ '136' = [DateTime]::UtcNow }
$script:OrphanAuditLastCheckedAt = [DateTime]::MinValue
$script:ProviderStartedAt = ([DateTime]::UtcNow).AddSeconds(-3600)
Invoke-OrphanAudit -ClusterVMs @($orphanVm)
Assert ((Get-OrphanLineCount) -eq $before) 'I1: a just-deleted VM is not flagged as an orphan candidate'

# I2: with no poll record and a provider that only just started, staleness
#     is measured from process start -- so nothing is flagged yet.
$before = (Get-OrphanLineCount)
$script:LastGuestPollAt = @{}
$script:RecentlyDeletedIds = @{}
$script:OrphanAuditLastCheckedAt = [DateTime]::MinValue
$script:ProviderStartedAt = ([DateTime]::UtcNow).AddSeconds(-5)
Invoke-OrphanAudit -ClusterVMs @($orphanVm)
Assert ((Get-OrphanLineCount) -eq $before) 'I2: a freshly started provider does not mass-flag existing clones'

# I3: but a genuinely unheard-of clone, after the provider has been up
#     well past the staleness window, IS still flagged -- the fix must not
#     disable the detection it was added for.
$before = (Get-OrphanLineCount)
$script:LastGuestPollAt = @{}
$script:RecentlyDeletedIds = @{}
$script:OrphanAuditLastCheckedAt = [DateTime]::MinValue
$script:ProviderStartedAt = ([DateTime]::UtcNow).AddSeconds(-3600)
Invoke-OrphanAudit -ClusterVMs @($orphanVm)
Assert ((Get-OrphanLineCount) -eq $before + 1) 'I3: a long-unpolled clone is still flagged'

# I4: and a currently-polled one still is not.
$before = (Get-OrphanLineCount)
$script:LastGuestPollAt = @{ '136' = [DateTime]::UtcNow }
$script:OrphanAuditLastCheckedAt = [DateTime]::MinValue
Invoke-OrphanAudit -ClusterVMs @($orphanVm)
Assert ((Get-OrphanLineCount) -eq $before) 'I4: an actively polled clone is still not flagged'

# ------------------------------------------------------------
Write-Host ""
Write-Host "PASSED: $pass  FAILED: $fail" -ForegroundColor $(if ($fail -eq 0) { 'Green' } else { 'Red' })
if ($fail -gt 0) { exit 1 }
