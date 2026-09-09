$ErrorActionPreference = 'Stop'
$pass = 0
$fail = 0
function Assert {
    param([bool]$Cond, [string]$Msg)
    if ($Cond) { $script:pass++; Write-Host "PASS: $Msg" -ForegroundColor Green }
    else { $script:fail++; Write-Host "FAIL: $Msg" -ForegroundColor Red }
}

# ============================================================
# TESTS for linked-clone support (LINKED-CLONES-DESIGN.md):
#   guests/snapshots/create|exists|delete|revert -- the virtual-snapshot
#   invariant (exists tracks the guest's own template flag, nothing else)
#
#   Handle-GuestClone's linked-vs-full decision -- including the trap (a
#   non-template source with 'snapshot' set must not silently downgrade,
#   the way real Proxmox itself would) and the StrictMode-safe absent-vs-
#   empty handling of 'snapshot'/'is_link_clone' -- an unguarded read here
#   would crash guests/clone on every ordinary full clone.
# ============================================================
. "$PSScriptRoot/funcs2.ps1"

$script:LogPath        = "$PSScriptRoot/t11.log"
$script:CloneStatePath = "$PSScriptRoot/t11-state.json"
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
$script:TemplateSnapshotIntent = @{}

# Deterministic regardless of whatever RAS-CPF-Proxmox-Settings.json happens to hold on
# disk (it self-seeds from whatever the defaults were at the time it was first written)
# -- this suite drives both new settings directly.
$script:Settings.capabilities.can_link_clones = $true
$script:Settings.cloning.linked_clone_fallback = 'full'

$script:apiCalls = New-Object System.Collections.ArrayList
$script:nextId = 300
$script:vmState    = @{}   # vmid -> status
$script:vmName     = @{}
$script:vmVisible  = @{}
$script:vmTemplate = @{}   # vmid -> $true/$false, the LIVE config's template flag
$script:vmTags     = @{}
$script:lastCloneBody = $null   # last POST body to .../clone, for assertions

# 153 = a real Proxmox template. 160 = an ordinary running VM -- the trap's source.
$script:vmTemplate['153'] = $true
$script:vmTemplate['160'] = $false

function Invoke-ProxmoxApi {
    param([string]$Method, [string]$Path, $Body, [int]$TimeoutSec = 0)
    [void]$script:apiCalls.Add("$Method $Path")

    if ($Path -match '/cluster/nextid') { return @{ data = "$($script:nextId)" } }

    if ($Path -match '/cluster/resources') {
        $list = New-Object System.Collections.ArrayList
        foreach ($fixedId in @('153', '160')) {
            $entry = [pscustomobject]@{ type = 'qemu'; vmid = [int]$fixedId; name = "vm$fixedId"; node = 'n1'; status = 'stopped'
                template = $(if ($script:vmTemplate[$fixedId]) { 1 } else { 0 }) }
            if ($script:vmTags.ContainsKey($fixedId)) { $entry | Add-Member -NotePropertyName tags -NotePropertyValue $script:vmTags[$fixedId] }
            [void]$list.Add($entry)
        }
        foreach ($k in $script:vmState.Keys) {
            if (-not $script:vmVisible[$k]) { continue }
            $entry = [pscustomobject]@{ type = 'qemu'; vmid = [int]$k; name = $script:vmName[$k]; node = 'n1'; status = $script:vmState[$k]
                template = $(if ($script:vmTemplate[$k]) { 1 } else { 0 }) }
            if ($script:vmTags.ContainsKey($k)) { $entry | Add-Member -NotePropertyName tags -NotePropertyValue $script:vmTags[$k] }
            [void]$list.Add($entry)
        }
        return @{ data = $list.ToArray() }
    }

    if ($Method -eq 'GET' -and $Path -match '/qemu/(\d+)/config$') {
        $vmid = $Matches[1]
        $tmplFlag = ($script:vmTemplate.ContainsKey($vmid) -and $script:vmTemplate[$vmid])
        $obj = [ordered]@{ template = $(if ($tmplFlag) { 1 } else { 0 }) }
        if ($script:vmTags.ContainsKey($vmid)) { $obj.tags = $script:vmTags[$vmid] }
        return @{ data = [pscustomobject]$obj }
    }

    if ($Method -eq 'PUT' -and $Path -match '/qemu/(\d+)/config$') {
        $vmid = $Matches[1]
        if ($null -ne $Body -and $Body.ContainsKey('tags')) { $script:vmTags[$vmid] = [string]$Body.tags }
        if ($null -ne $Body -and $Body.ContainsKey('template')) { $script:vmTemplate[$vmid] = ([int]$Body.template -eq 1) }
        return @{ data = $null }
    }

    if ($Path -match '/qemu/(\d+)/clone$') {
        $newid = [string]$Body.newid
        $script:lastCloneBody = $Body
        $tid = "UPID:n1:clone$newid"
        $script:vmState[$newid]    = 'stopped'
        $script:vmName[$newid]     = "VM $newid"
        $script:vmVisible[$newid]  = $true
        $script:vmTemplate[$newid] = $false
        return @{ data = $tid }
    }

    if ($Path -match '/tasks/(.+)/status$') {
        return @{ data = [pscustomobject]@{ status = 'stopped'; exitstatus = 'OK' } }
    }

    throw "Unexpected call in mock: $Method $Path"
}

function HasError  { param($R) return ($R -is [hashtable] -and $R.ContainsKey('error')) }
function HasResult { param($R) return ($R -is [hashtable] -and $R.ContainsKey('result')) }

# Every C-series clone leaves a VM behind at the mock's one fixed next-id (300) --
# clear it out between cases so the next Get-ProxmoxNextVmId lands on a clean slot.
function Clear-LastClone {
    param($CloneResult)
    if ($null -eq $CloneResult -or -not (HasResult $CloneResult)) { return }
    $id = [string]$CloneResult.result.clone_id
    if ([string]::IsNullOrWhiteSpace($id)) { return }
    Remove-CloneStateEntry -VmId $id
    $script:vmState.Remove($id); $script:vmName.Remove($id); $script:vmVisible.Remove($id); $script:vmTemplate.Remove($id)
}

# ------------------------------------------------------------
# A. guests/snapshots/* -- the virtual-snapshot invariant (LINKED-CLONES-
#    DESIGN.md #2): exists is true if and only if the guest is presently a
#    native Proxmox template. Nothing is ever created on Proxmox itself.
# ------------------------------------------------------------
$rExists1 = Handle-GuestSnapshotsExists -Params ([pscustomobject]@{ id = '153'; name = 'RAS Template Snapshot' })
Assert ((HasResult $rExists1) -and $rExists1.result -eq $true) "A1: exists is true for a guest that IS a native Proxmox template"
Assert ($rExists1.result -is [bool]) "A2: exists returns a BARE boolean -- Read-ResultObject in the RAS test kit unwraps 'result' directly"

$rExists2 = Handle-GuestSnapshotsExists -Params ([pscustomobject]@{ id = '160'; name = 'RAS Template Snapshot' })
Assert ($rExists2.result -eq $false) "A3: exists is false for a guest that is NOT a template"

$rCreate = Handle-GuestSnapshotsCreate -Params ([pscustomobject]@{ id = '160'; name = 'RAS Template Snapshot' })
Assert ((HasResult $rCreate) -and -not [string]::IsNullOrWhiteSpace([string]$rCreate.result.task_id)) "A4: snapshots/create returns a task_id"
Assert (-not ($script:apiCalls -match '/qemu/160/(snapshot|snapconfig)')) "A5: snapshots/create makes no real Proxmox call"
Assert ((Handle-GuestSnapshotsExists -Params ([pscustomobject]@{ id = '160'; name = 'RAS Template Snapshot' })).result -eq $false) `
    "A6: exists is unaffected by a preceding create -- only the guest's template flag governs it"

$rDelete = Handle-GuestSnapshotsDelete -Params ([pscustomobject]@{ id = '153'; name = 'RAS Template Snapshot' })
Assert (-not [string]::IsNullOrWhiteSpace([string]$rDelete.result.task_id)) "A7: snapshots/delete returns a task_id"
Assert ((Handle-GuestSnapshotsExists -Params ([pscustomobject]@{ id = '153'; name = 'RAS Template Snapshot' })).result -eq $true) `
    "A8: design -- delete is a no-op in practice; exists stays true because 153 is still a real template"

$rRevert = Handle-GuestSnapshotsRevert -Params ([pscustomobject]@{ id = '153'; name = 'RAS Template Snapshot' })
Assert (HasError $rRevert) "A9: snapshots/revert errors -- unreachable at template_method=basic, no Proxmox state to revert to"
Assert ($rRevert.error.code -eq $script:ErrorCodes.InvalidParams) "A10: ...with InvalidParams, not a generic internal error"

# Note: a missing 'params.id' is not exercised directly against the handlers here --
# every RAS method (this one included) is gated by $script:MethodRegistry's own
# RequiredFields check before the handler is ever invoked, same as every other guests/*
# method in this provider (e.g. Handle-GuestConvert has the identical unguarded
# '$Params.id' read, relying on the same dispatch-level gate).

# ------------------------------------------------------------
# B. Handle-Initialize -- capabilities.can_link_clones is the single value
#    both the advertised capability AND Handle-GuestClone's actual gate read
#    (see Get-DefaultProviderSettings's comment on this key for why an
#    advertised capability the underlying feature cannot serve is a real
#    problem, not just which section the setting lives in).
# ------------------------------------------------------------
$script:Settings.capabilities.can_link_clones = $true
Assert ((Handle-Initialize).result.capabilities.can_link_clones -eq $true) "B1: capabilities.can_link_clones=true is reflected back by provider/initialize"
$script:Settings.capabilities.can_link_clones = $false
Assert ((Handle-Initialize).result.capabilities.can_link_clones -eq $false) "B2: ...and follows it back to false"
$script:Settings.capabilities.can_link_clones = $true   # restore for the rest of this suite

# ------------------------------------------------------------
# C. Handle-GuestClone -- the linked-vs-full decision
# ------------------------------------------------------------

# C1: no 'snapshot' param present AT ALL still full-clones. Under Set-StrictMode
# (active here, inherited from funcs2.ps1's own top-of-file Set-StrictMode -Version
# Latest) an unguarded property read on a PSCustomObject lacking 'snapshot' throws.
# Handle-GuestClone wraps its body in try/catch, so a regression here shows up as
# HasError, not a hard crash.
$script:lastCloneBody = $null
$r = Handle-GuestClone -Params ([pscustomobject]@{ id = '153'; name = 'plain-full' })
Assert (-not (HasError $r)) "C1: a plain clone (no 'snapshot' param at all) does not throw / error"
Assert ([int]$script:lastCloneBody.full -eq 1) "C1: ...and is still a full clone -- an absent 'snapshot' is never a linked-clone signal"
Clear-LastClone $r

# C2: an empty 'snapshot' string -- what the real test kit's Submit-GuestsClone
# always sends when no name is passed -- is likewise not a linked-clone request.
$script:lastCloneBody = $null
$r = Handle-GuestClone -Params ([pscustomobject]@{ id = '153'; name = 'empty-snapshot'; snapshot = '' })
Assert ([int]$script:lastCloneBody.full -eq 1) "C2: an empty (but present) 'snapshot' string is treated the same as absent -- still full"
Clear-LastClone $r

# C3: the happy path -- 'snapshot' set, source really is a template -> linked.
$script:lastCloneBody = $null
$r = Handle-GuestClone -Params ([pscustomobject]@{ id = '153'; name = 'linked-1'; snapshot = 'RAS Template Snapshot' })
Assert (-not (HasError $r)) "C3: a linked clone from a real template succeeds"
Assert ([int]$script:lastCloneBody.full -eq 0) "C3: ...and is asked of Proxmox as full=0"
Assert (-not ($script:lastCloneBody.ContainsKey('storage') -or $script:lastCloneBody.ContainsKey('format'))) `
    "C3: ...and never carries 'storage'/'format', which Proxmox rejects on a linked clone"
Assert (-not $script:lastCloneBody.ContainsKey('snapname')) "C3: ...and never carries 'snapname' -- the RAS name is not a legal pve-configid"
Clear-LastClone $r

# C4: the trap (LINKED-CLONES-DESIGN.md #1) -- 'snapshot' set, but the source is
# NOT a template, default fallback 'full'. Real Proxmox would silently do this
# anyway -- no error, no warning -- so the point under test is that the provider
# makes this call itself and logs it, rather than passing full=0 and letting
# Proxmox's own silent downgrade be the only thing that saved it.
$script:Settings.cloning.linked_clone_fallback = 'full'
$script:lastCloneBody = $null
$r = Handle-GuestClone -Params ([pscustomobject]@{ id = '160'; name = 'trap-full'; snapshot = 'RAS Template Snapshot' })
Assert (-not (HasError $r)) "C4: the trap does not fail the clone -- it deliberately falls back instead"
Assert ([int]$script:lastCloneBody.full -eq 1) "C4: source [160] is not a template, so full=1 despite the snapshot param"
Clear-LastClone $r

# C5: same trap, fallback='error' -- must refuse outright, and never even reach
# Proxmox's clone endpoint. This is the decisive regression check: unlike C4, it
# cannot pass by accident (there's no mock-side safety net to fall back on if the
# provider-side template check is skipped).
$script:Settings.cloning.linked_clone_fallback = 'error'
$script:apiCalls.Clear()
$r = Handle-GuestClone -Params ([pscustomobject]@{ id = '160'; name = 'trap-error'; snapshot = 'RAS Template Snapshot' })
Assert (HasError $r) "C5: linked_clone_fallback=error refuses the clone outright"
Assert ($r.error.code -eq $script:ErrorCodes.InvalidParams) "C5: ...with InvalidParams"
Assert (-not ($script:apiCalls -match '/qemu/160/clone$')) "C5: ...and never issues the clone POST to Proxmox at all"
$script:Settings.cloning.linked_clone_fallback = 'full'   # restore

# C6: is_link_clone=false is an override RAS can use to force a full clone even
# though a snapshot name is present and the source really is a template. Per the
# docs this flag is "only used by template versions" and the real test kit never
# sends it -- but when present, it must still be honoured.
$script:lastCloneBody = $null
$r = Handle-GuestClone -Params ([pscustomobject]@{ id = '153'; name = 'override-off'; snapshot = 'RAS Template Snapshot'; is_link_clone = $false })
Assert ([int]$script:lastCloneBody.full -eq 1) "C6: is_link_clone=false overrides a matching snapshot+template back to full"
Clear-LastClone $r

# C7: master switch off -- capabilities.can_link_clones=false ignores 'snapshot'
# entirely, matching today's shipped behaviour exactly. The feature is opt-in.
$script:Settings.capabilities.can_link_clones = $false
$script:lastCloneBody = $null
$r = Handle-GuestClone -Params ([pscustomobject]@{ id = '153'; name = 'feature-off'; snapshot = 'RAS Template Snapshot' })
Assert ([int]$script:lastCloneBody.full -eq 1) "C7: capabilities.can_link_clones=false ignores a real snapshot param -- the feature is opt-in"
Clear-LastClone $r
$script:Settings.capabilities.can_link_clones = $true

Write-Host "`n=== t11 Summary: $pass passed, $fail failed ===" -ForegroundColor $(if ($fail -eq 0) { 'Green' } else { 'Red' })
if ($fail -gt 0) { exit 1 }
