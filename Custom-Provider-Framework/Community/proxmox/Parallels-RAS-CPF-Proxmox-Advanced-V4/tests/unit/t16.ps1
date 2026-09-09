$ErrorActionPreference = 'Stop'
$pass = 0
$fail = 0
function Assert {
    param([bool]$Cond, [string]$Msg)
    if ($Cond) { $script:pass++; Write-Host "PASS: $Msg" -ForegroundColor Green }
    else { $script:fail++; Write-Host "FAIL: $Msg" -ForegroundColor Red }
}

# ============================================================
# TESTS for pool scoping (POOL-SCOPING.md):
#   virtual_machines.pool_scope.pool_name -- guests/list, hosts/list, guests/get
#     all treat a VM outside the configured pool exactly like rasExclude (never
#     listed, not-found on direct lookup). Empty (default) = no filtering.
#   virtual_machines.pool_scope.inherit_on_clone -- a clone's Proxmox 'pool'
#     lands in the guests/clone request body when the source belongs to one.
# ============================================================
. "$PSScriptRoot/funcs2.ps1"

$script:LogPath        = "$PSScriptRoot/t16.log"
$script:CloneStatePath = "$PSScriptRoot/t16-state.json"
Remove-Item $script:LogPath, $script:CloneStatePath -ErrorAction SilentlyContinue
$script:TrackedCloneVmIds = [System.Collections.Generic.HashSet[string]]::new()
$script:ClusterResourcesCache = $null
$script:ClusterResourcesCachedAt = $null
$script:RecentlyDeletedIds = @{}
$script:RecentlyDeletedNames = @{}
$script:RecentlyControlledIds = @{}
$script:AgentFailureTracker = @{}
$script:LastKnownNetworkData = @{}
$script:SweepLastCheckedAt = @{}
$script:RasTemplateTagAttempted = [System.Collections.Generic.HashSet[string]]::new()
$script:CloneTagVerified = [System.Collections.Generic.HashSet[string]]::new()
$script:CloneStateMemory = $null
$script:TaskContext = @{}
$script:CloneTaskCompletionCache = @{}

# Deterministic regardless of whatever RAS-CPF-Proxmox-Settings.json holds on disk --
# pool_scope starts OFF (empty pool_name), matching the documented default.
$script:PoolScopeName = ''
$script:PoolScopeInheritOnClone = $true

$script:apiCalls = New-Object System.Collections.ArrayList
$script:nextId = 700
$script:vmState  = @{}   # vmid -> status
$script:vmName   = @{}
$script:vmVisible = @{}
$script:vmPool   = @{}   # vmid -> pool name, absent key = unpooled
$script:lastCloneBody = $null

function Invoke-ProxmoxApi {
    param([string]$Method, [string]$Path, $Body, [int]$TimeoutSec = 0)
    [void]$script:apiCalls.Add("$Method $Path")

    if ($Path -match '/cluster/nextid') { return @{ data = "$($script:nextId)" } }

    if ($Path -match '/cluster/resources') {
        $list = New-Object System.Collections.ArrayList
        # A pooled template, source for the clone-inheritance tests below.
        $tmpl = [pscustomobject]@{ type = 'qemu'; vmid = 600; name = 'pooled-template'; node = 'n1'; status = 'stopped'; template = 1 }
        if ($script:vmPool.ContainsKey('600')) { $tmpl | Add-Member -MemberType NoteProperty -Name pool -Value $script:vmPool['600'] }
        [void]$list.Add($tmpl)
        # An unpooled template, source for the "unpooled source clones unpooled" test.
        [void]$list.Add([pscustomobject]@{ type = 'qemu'; vmid = 601; name = 'unpooled-template'; node = 'n1'; status = 'stopped'; template = 1 })

        foreach ($k in $script:vmState.Keys) {
            if (-not $script:vmVisible[$k]) { continue }
            $entry = [pscustomobject]@{ type = 'qemu'; vmid = [int]$k; name = $script:vmName[$k]; node = 'n1'; status = $script:vmState[$k]; template = 0 }
            if ($script:vmPool.ContainsKey($k)) { $entry | Add-Member -MemberType NoteProperty -Name pool -Value $script:vmPool[$k] }
            [void]$list.Add($entry)
        }
        return @{ data = $list.ToArray() }
    }

    if ($Path -match '/qemu/(\d+)/clone$') {
        $newid = [string]$Body.newid
        $script:lastCloneBody = $Body
        $tid = "UPID:n1:clone$newid"
        $script:vmState[$newid] = 'stopped'
        $script:vmName[$newid]  = [string]$Body.name
        $script:vmVisible[$newid] = $true
        if ($Body.ContainsKey('pool')) { $script:vmPool[$newid] = [string]$Body.pool }
        return @{ data = $tid }
    }

    if ($Path -match '/tasks/(.+)/status$') {
        return @{ data = [pscustomobject]@{ status = 'stopped'; exitstatus = 'OK' } }
    }

    if ($Path -match '/qemu/(\d+)/status/current$') {
        $id = $Matches[1]
        $st = if ($script:vmState.ContainsKey($id)) { $script:vmState[$id] } else { 'stopped' }
        return @{ data = [pscustomobject]@{ status = $st } }
    }

    if ($Path -match '/qemu/(\d+)/agent/network-get-interfaces$') {
        throw "500 no agent"
    }

    if ($Path -match '/qemu/(\d+)/status/start$') {
        $id = $Matches[1]
        $script:vmState[$id] = 'running'
        return @{ data = "UPID:n1:start$id" }
    }

    throw "Unexpected call in mock: $Method $Path"
}

# ------------------------------------------------------------
# A. No filtering (pool_name empty, the documented default) -- every VM
#    visible regardless of pool, pooled or not.
# ------------------------------------------------------------
Write-Host "`n--- Test A: pool_name empty -- no filtering ---" -ForegroundColor Cyan
$script:vmState['100'] = 'running'; $script:vmName['100'] = 'in-pool-a'; $script:vmVisible['100'] = $true; $script:vmPool['100'] = 'pool-a'
$script:vmState['101'] = 'running'; $script:vmName['101'] = 'in-pool-b'; $script:vmVisible['101'] = $true; $script:vmPool['101'] = 'pool-b'
$script:vmState['102'] = 'running'; $script:vmName['102'] = 'unpooled'; $script:vmVisible['102'] = $true
Reset-ProxmoxClusterCache
$listA = Handle-GuestList
Assert ($listA.result.guests -contains '100') "A1: pool_name empty -- VM in pool-a is listed"
Assert ($listA.result.guests -contains '101') "A2: pool_name empty -- VM in pool-b is listed"
Assert ($listA.result.guests -contains '102') "A3: pool_name empty -- unpooled VM is listed"
$getA = Handle-GuestGet -Params ([pscustomobject]@{ id = '101' })
Assert (-not $getA.ContainsKey('error')) "A4: pool_name empty -- guests/get succeeds for a VM in any pool"

# ------------------------------------------------------------
# B. pool_name set -- only that exact pool is visible; a different pool and
#    an unpooled VM are both treated exactly like rasExclude.
# ------------------------------------------------------------
Write-Host "`n--- Test B: pool_name = 'pool-a' -- scoped ---" -ForegroundColor Cyan
$script:PoolScopeName = 'pool-a'
Reset-ProxmoxClusterCache
$listB = Handle-GuestList
Assert ($listB.result.guests -contains '100') "B1: VM in the scoped pool (pool-a) is still listed"
Assert (-not ($listB.result.guests -contains '101')) "B2: VM in a DIFFERENT pool (pool-b) is omitted from guests/list"
Assert (-not ($listB.result.guests -contains '102')) "B3: an UNPOOLED VM is omitted from guests/list once a pool is scoped"

$hostsB = Handle-HostList
Assert ($hostsB.result.guests -contains '100') "B4: hosts/list also respects pool_scope -- in-scope VM listed"
Assert (-not ($hostsB.result.guests -contains '101')) "B5: hosts/list also respects pool_scope -- out-of-scope VM omitted"

# Both Handle-GuestList (above) and Handle-HostList (just now) ran against the SAME
# cached cluster listing (one Reset-ProxmoxClusterCache before either call) -- pool_scope
# filtering happens once, at the shared Get-ProxmoxClusterVMs fetch, not separately in
# each caller. So the skip for VM [101] must be logged exactly once, at 'T' (Trace/
# "Extended", a routine filter outcome, not 'I'), not once per caller that happened to
# see the already-filtered list.
$logContentB = Get-Content $script:LogPath -Raw
$excl101Count = ([regex]::Matches($logContentB, [regex]::Escape('Excluding guest [101]'))).Count
Assert ($excl101Count -eq 1) "B5a: the pool_scope skip for VM [101] is logged only ONCE across both guests/list and hosts/list -- filtering happens once at the shared cache fetch, not per caller"
Assert ($logContentB -match '\[T 03/101/P[0-9A-F]+\] .* - Excluding guest \[101\] from this provider''s view: outside pool_scope') "B5b: the skip is logged at Trace ('T'/Extended), a routine filter outcome, not Info"

$errB = Handle-GuestGet -Params ([pscustomobject]@{ id = '101' })
Assert ($errB.ContainsKey('error')) "B6: guests/get for a VM outside the scoped pool behaves exactly like a not-found VM (Get-ProxmoxVmNode enforcement)"
$okB = Handle-GuestGet -Params ([pscustomobject]@{ id = '100' })
Assert (-not $okB.ContainsKey('error')) "B7: guests/get for a VM inside the scoped pool still succeeds"

Write-Host "  (admin clears pool_name)"
$script:PoolScopeName = ''
Reset-ProxmoxClusterCache
$listB2 = Handle-GuestList
Assert ($listB2.result.guests -contains '101') "B8: clearing pool_name makes every VM visible again immediately, same as removing rasExclude"

# ------------------------------------------------------------
# C. Clone pool inheritance -- inherit_on_clone (default true) copies the
#    SOURCE's own pool into the clone request body; an unpooled source never
#    sends a 'pool' key at all; inherit_on_clone=false suppresses it even for
#    a pooled source.
# ------------------------------------------------------------
Write-Host "`n--- Test C: clone pool inheritance ---" -ForegroundColor Cyan
$script:vmPool['600'] = 'pool-a'
Reset-ProxmoxClusterCache
$script:PoolScopeInheritOnClone = $true
$script:lastCloneBody = $null
$script:nextId = 701   # distinct id per clone below -- Get-ProxmoxNextVmId is mocked to
                        # a fixed value, so reusing one id across clones in the same run
                        # would leave stale clone-tracking state from the prior clone.
[void](Handle-GuestClone -Params ([pscustomobject]@{ id = '600'; name = 'clone-from-pooled' }))
Assert ($null -ne $script:lastCloneBody -and $script:lastCloneBody.ContainsKey('pool')) "C1: inherit_on_clone=true -- clone request body carries a 'pool' key"
Assert ($script:lastCloneBody.pool -eq 'pool-a') "C2: ...and it matches the source template's own pool exactly"

$script:lastCloneBody = $null
$script:nextId = 702
[void](Handle-GuestClone -Params ([pscustomobject]@{ id = '601'; name = 'clone-from-unpooled' }))
Assert ($null -ne $script:lastCloneBody -and -not $script:lastCloneBody.ContainsKey('pool')) "C3: an unpooled source clones unpooled -- no 'pool' key sent at all, not an empty one"

$script:PoolScopeInheritOnClone = $false
$script:lastCloneBody = $null
$script:nextId = 703
[void](Handle-GuestClone -Params ([pscustomobject]@{ id = '600'; name = 'clone-inherit-off' }))
Assert ($null -ne $script:lastCloneBody -and -not $script:lastCloneBody.ContainsKey('pool')) "C4: inherit_on_clone=false -- no 'pool' key sent even though the source belongs to one"
$script:PoolScopeInheritOnClone = $true

# ------------------------------------------------------------
# D. Integration -- pool filtering ON, cloning FROM an in-scope pooled
#    template: the new clone inherits the pool and immediately RESOLVES
#    (guests/get, via Get-ProxmoxVmNode's pool_scope enforcement), rather
#    than being invisible under the very feature it was just created to
#    satisfy. Uses guests/get, not guests/list: a freshly-tracked clone is
#    ALSO gated out of guests/list until its task is reported 'completed'
#    (a separate, unrelated mechanism -- see TESTING.md's t9), which would
#    conflate two different gates in one assertion.
# ------------------------------------------------------------
Write-Host "`n--- Test D: filtering + inheritance together ---" -ForegroundColor Cyan
$script:PoolScopeName = 'pool-a'
Reset-ProxmoxClusterCache
$script:lastCloneBody = $null
$script:nextId = 704
$rD = Handle-GuestClone -Params ([pscustomobject]@{ id = '600'; name = 'scoped-clone' })
$newIdD = [string]$rD.result.clone_id
Reset-ProxmoxClusterCache
$getD = Handle-GuestGet -Params ([pscustomobject]@{ id = $newIdD })
Assert (-not $getD.ContainsKey('error')) "D1: a clone of an in-scope pooled template resolves via guests/get under pool_scope -- inheritance and filtering compose correctly, not hidden by the feature it was created to satisfy"
$script:PoolScopeName = ''

Write-Host "`n=== t16 Summary: $pass passed, $fail failed ===" -ForegroundColor $(if ($fail -eq 0) { 'Green' } else { 'Red' })
if ($fail -gt 0) { exit 1 }
