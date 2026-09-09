$ErrorActionPreference = 'Stop'
$pass = 0
$fail = 0
function Assert {
    param([bool]$Cond, [string]$Msg)
    if ($Cond) { $script:pass++; Write-Host "PASS: $Msg" -ForegroundColor Green }
    else { $script:fail++; Write-Host "FAIL: $Msg" -ForegroundColor Red }
}

# ============================================================
# TESTS for distributed clone placement (DISTRIBUTED-PLACEMENT.md):
#   Get-ProxmoxClusterNodes -- separate cache from the VM listing
#   Resolve-ProxmoxCloneTargetNode -- eligibility (online/maintenance/excluded),
#     'resource' scoring (cpu/ram/both), 'round_robin' cycling, recreation
#     node preservation, and the always-safe "placement disabled or nothing
#     eligible -> $null, no 'target' at all" fallback.
#   Handle-GuestClone -- 'target' actually lands in the clone body, and only
#     when placement resolved one.
# ============================================================
. "$PSScriptRoot/funcs2.ps1"

$script:LogPath        = "$PSScriptRoot/t14.log"
$script:CloneStatePath = "$PSScriptRoot/t14-state.json"
Remove-Item $script:LogPath, $script:CloneStatePath -ErrorAction SilentlyContinue
$script:TrackedCloneVmIds = [System.Collections.Generic.HashSet[string]]::new()
$script:ClusterResourcesCache = $null
$script:ClusterResourcesCachedAt = $null
$script:ClusterNodesCache = $null
$script:ClusterNodesCachedAt = $null
$script:RecentlyDeletedIds = @{}
$script:RecentlyDeletedNames = @{}
$script:RecentlyDeletedNodeByName = @{}
$script:RecentlyControlledIds = @{}
$script:AgentFailureTracker = @{}
$script:LastKnownNetworkData = @{}
$script:SweepLastCheckedAt = @{}
$script:RasTemplateTagAttempted = [System.Collections.Generic.HashSet[string]]::new()
$script:CloneTagVerified = [System.Collections.Generic.HashSet[string]]::new()
$script:CloneStateMemory = $null
$script:TaskContext = @{}
$script:CloneTaskCompletionCache = @{}
$script:PlacementRoundRobinIndex = 0

# Deterministic regardless of whatever RAS-CPF-Proxmox-Settings.json holds on disk.
$script:Settings.cloning.load_balancing.enabled = $true
$script:Settings.cloning.load_balancing.strategy = 'resource'
$script:Settings.cloning.load_balancing.resource_metric = 'both'
$script:Settings.cloning.load_balancing.excluded_nodes = @()
$script:Settings.cloning.load_balancing.node_stats_cache_ttl_seconds = 15
$script:Settings.cloning.load_balancing.preserve_node_on_recreation = $true
# Resolve-ProxmoxCloneTargetNode / Get-ProxmoxPreservedNodeForRecreation read the
# DERIVED runtime variable Set-ProviderRuntimeFromSettings populates from this setting,
# not $script:Settings itself on every call -- set both so this suite is not silently
# dependent on whatever RAS-CPF-Proxmox-Settings.json (the fixture) happens to have.
$script:Settings.cloning.timeouts.recently_deleted_retention_seconds = 300
$script:RecentlyDeletedRetentionSeconds = 300

$script:apiCalls = New-Object System.Collections.ArrayList
$script:nodeCallCount = 0
$script:nextId = 900
$script:vmState    = @{}
$script:vmName     = @{}
$script:vmVisible  = @{}
$script:vmTemplate = @{}
$script:lastCloneBody = $null

# The three-node sample from DISTRIBUTED-PLACEMENT.md's own worked example --
# pve-node3 has vastly more absolute RAM but the HIGHEST utilization
# fraction of the three, which is exactly what a fraction-based score must
# pick up on.
$script:nodeFixtures = @(
    [pscustomobject]@{ type = 'node'; node = 'pve-node1'; status = 'online'; cpu = 0.056; maxcpu = 4; mem = 759422976; maxmem = 882696192 }
    [pscustomobject]@{ type = 'node'; node = 'pve-node2'; status = 'online'; cpu = 0.049; maxcpu = 4; mem = 737861632; maxmem = 882683904 }
    [pscustomobject]@{ type = 'node'; node = 'pve-node3'; status = 'online'; cpu = 0.119; maxcpu = 40; mem = 119108378624; maxmem = 134801514496 }
)

function Invoke-ProxmoxApi {
    param([string]$Method, [string]$Path, $Body, [int]$TimeoutSec = 0)
    [void]$script:apiCalls.Add("$Method $Path")

    if ($Path -match '/cluster/nextid') { return @{ data = "$($script:nextId)" } }

    if ($Path -match '/cluster/resources\?type=node') {
        $script:nodeCallCount++
        return @{ data = @($script:nodeFixtures) }
    }

    if ($Path -match '/cluster/resources') {
        $list = New-Object System.Collections.ArrayList
        foreach ($fixedId in @('153')) {
            $entry = [pscustomobject]@{ type = 'qemu'; vmid = [int]$fixedId; name = "vm$fixedId"; node = 'pve-node1'; status = 'stopped'; template = 1 }
            [void]$list.Add($entry)
        }
        foreach ($k in $script:vmState.Keys) {
            if (-not $script:vmVisible[$k]) { continue }
            $entry = [pscustomobject]@{ type = 'qemu'; vmid = [int]$k; name = $script:vmName[$k]; node = 'pve-node1'; status = $script:vmState[$k]; template = 0 }
            [void]$list.Add($entry)
        }
        return @{ data = $list.ToArray() }
    }

    if ($Method -eq 'GET' -and $Path -match '/qemu/(\d+)/config$') {
        return @{ data = [pscustomobject]@{ template = 1 } }
    }

    if ($Method -eq 'PUT' -and $Path -match '/qemu/(\d+)/config$') {
        return @{ data = $null }
    }

    if ($Path -match '/qemu/(\d+)/clone$') {
        $newid = [string]$Body.newid
        $script:lastCloneBody = $Body
        $tid = "UPID:pve-node1:clone$newid"
        $script:vmState[$newid]    = 'stopped'
        $script:vmName[$newid]     = [string]$Body.name
        $script:vmVisible[$newid]  = $true
        $script:vmTemplate[$newid] = $false
        return @{ data = $tid }
    }

    if ($Path -match '/tasks/(.+)/status$') {
        return @{ data = [pscustomobject]@{ status = 'stopped'; exitstatus = 'OK' } }
    }

    throw "Unexpected call in mock: $Method $Path"
}

function Clear-LastClone {
    param($CloneResult)
    if ($null -eq $CloneResult -or -not ($CloneResult -is [hashtable] -and $CloneResult.ContainsKey('result'))) { return }
    $id = [string]$CloneResult.result.clone_id
    if ([string]::IsNullOrWhiteSpace($id)) { return }
    Remove-CloneStateEntry -VmId $id
    $script:vmState.Remove($id); $script:vmName.Remove($id); $script:vmVisible.Remove($id); $script:vmTemplate.Remove($id)
}

# ------------------------------------------------------------
# A. placement.enabled = false -- never calls Get-ProxmoxClusterNodes at all
# ------------------------------------------------------------
$script:Settings.cloning.load_balancing.enabled = $false
$script:nodeCallCount = 0
$offResult = Resolve-ProxmoxCloneTargetNode -SourceNode 'pve-node1' -CloneName 'off-test'
Assert ($null -eq $offResult) "A1: placement.enabled=false returns `$null"
Assert ($script:nodeCallCount -eq 0) "A2: ...and never calls GET /cluster/resources?type=node -- zero wasted HTTP calls when the feature is off"
$script:Settings.cloning.load_balancing.enabled = $true

# ------------------------------------------------------------
# B. Eligibility -- offline, HA maintenance, and excluded_nodes are three
#    independent reasons a node is skipped; a node with no 'hastate' field
#    at all is never excluded on that basis.
# ------------------------------------------------------------
$script:ClusterNodesCache = $null; $script:ClusterNodesCachedAt = $null
$script:nodeFixtures = @(
    [pscustomobject]@{ type = 'node'; node = 'n-offline'; status = 'offline'; cpu = 0.01; maxcpu = 4; mem = 1; maxmem = 100 }
    [pscustomobject]@{ type = 'node'; node = 'n-maint'; status = 'online'; hastate = 'maintenance'; cpu = 0.01; maxcpu = 4; mem = 1; maxmem = 100 }
    [pscustomobject]@{ type = 'node'; node = 'n-excluded'; status = 'online'; cpu = 0.01; maxcpu = 4; mem = 1; maxmem = 100 }
    [pscustomobject]@{ type = 'node'; node = 'n-ok'; status = 'online'; cpu = 0.5; maxcpu = 4; mem = 50; maxmem = 100 }
)
$script:Settings.cloning.load_balancing.excluded_nodes = @('n-excluded')
$eligResult = Resolve-ProxmoxCloneTargetNode -SourceNode 'n-ok' -CloneName 'eligibility-test'
Assert ($eligResult -eq 'n-ok') "B1: an offline node, a node in HA maintenance, and an admin-excluded node are all skipped -- only n-ok is chosen"
$script:Settings.cloning.load_balancing.excluded_nodes = @()

# A node with NO 'hastate' property at all (HA not configured/running) must
# never be excluded on that basis alone.
$script:ClusterNodesCache = $null; $script:ClusterNodesCachedAt = $null
$script:nodeFixtures = @(
    [pscustomobject]@{ type = 'node'; node = 'n-no-hastate'; status = 'online'; cpu = 0.01; maxcpu = 4; mem = 1; maxmem = 100 }
)
$noHastateResult = Resolve-ProxmoxCloneTargetNode -SourceNode 'n-no-hastate' -CloneName 'no-hastate-test'
Assert ($noHastateResult -eq 'n-no-hastate') "B2: a node with no 'hastate' field at all is eligible -- absence is not itself maintenance"

# Every node ineligible -> falls back to `$null (source node), not an error.
$script:ClusterNodesCache = $null; $script:ClusterNodesCachedAt = $null
$script:nodeFixtures = @(
    [pscustomobject]@{ type = 'node'; node = 'n-only'; status = 'offline'; cpu = 0.01; maxcpu = 4; mem = 1; maxmem = 100 }
)
$noneEligibleResult = Resolve-ProxmoxCloneTargetNode -SourceNode 'n-only' -CloneName 'none-eligible-test'
Assert ($null -eq $noneEligibleResult) "B3: no eligible node at all -> `$null, falling back to the source node rather than erroring"

# ------------------------------------------------------------
# C. 'resource' strategy -- lower utilization fraction wins, for each metric,
#    using DISTRIBUTED-PLACEMENT.md's own worked three-node example.
# ------------------------------------------------------------
$script:ClusterNodesCache = $null; $script:ClusterNodesCachedAt = $null
$script:nodeFixtures = @(
    [pscustomobject]@{ type = 'node'; node = 'pve-node1'; status = 'online'; cpu = 0.056; maxcpu = 4; mem = 759422976; maxmem = 882696192 }
    [pscustomobject]@{ type = 'node'; node = 'pve-node2'; status = 'online'; cpu = 0.049; maxcpu = 4; mem = 737861632; maxmem = 882683904 }
    [pscustomobject]@{ type = 'node'; node = 'pve-node3'; status = 'online'; cpu = 0.119; maxcpu = 40; mem = 119108378624; maxmem = 134801514496 }
)

$script:Settings.cloning.load_balancing.resource_metric = 'cpu'
$cpuPick = Resolve-ProxmoxCloneTargetNode -SourceNode 'pve-node1' -CloneName 'cpu-metric-test'
Assert ($cpuPick -eq 'pve-node2') "C1: resource_metric=cpu picks the lowest 'cpu' fraction (pve-node2, 0.049)"

$script:ClusterNodesCache = $null; $script:ClusterNodesCachedAt = $null
$script:Settings.cloning.load_balancing.resource_metric = 'ram'
$ramPick = Resolve-ProxmoxCloneTargetNode -SourceNode 'pve-node1' -CloneName 'ram-metric-test'
Assert ($ramPick -eq 'pve-node2') "C2: resource_metric=ram picks the lowest mem/maxmem fraction (pve-node2, ~0.836) -- despite pve-node3 having ~150x the absolute RAM, ITS fraction used is the highest of the three"

$script:ClusterNodesCache = $null; $script:ClusterNodesCachedAt = $null
$script:Settings.cloning.load_balancing.resource_metric = 'both'
$bothPick = Resolve-ProxmoxCloneTargetNode -SourceNode 'pve-node1' -CloneName 'both-metric-test'
Assert ($bothPick -eq 'pve-node2') "C3: resource_metric=both (average of cpu+ram fractions) also picks pve-node2"

$script:ClusterNodesCache = $null; $script:ClusterNodesCachedAt = $null
$script:Settings.cloning.load_balancing.resource_metric = 'not-a-real-metric'
$badMetricPick = Resolve-ProxmoxCloneTargetNode -SourceNode 'pve-node1' -CloneName 'bad-metric-test'
Assert ($badMetricPick -eq 'pve-node2') "C4: an unrecognized resource_metric falls back to 'both' rather than throwing"
$script:Settings.cloning.load_balancing.resource_metric = 'both'

# ------------------------------------------------------------
# D. 'round_robin' strategy -- cycles through eligible nodes in name order,
#    wraps around, and correctly skips an excluded one without breaking the
#    cycle.
# ------------------------------------------------------------
$script:ClusterNodesCache = $null; $script:ClusterNodesCachedAt = $null
$script:nodeFixtures = @(
    [pscustomobject]@{ type = 'node'; node = 'n-c'; status = 'online'; cpu = 0; maxcpu = 1; mem = 0; maxmem = 1 }
    [pscustomobject]@{ type = 'node'; node = 'n-a'; status = 'online'; cpu = 0; maxcpu = 1; mem = 0; maxmem = 1 }
    [pscustomobject]@{ type = 'node'; node = 'n-b'; status = 'online'; cpu = 0; maxcpu = 1; mem = 0; maxmem = 1 }
)
$script:Settings.cloning.load_balancing.strategy = 'round_robin'
$script:PlacementRoundRobinIndex = 0
$rr1 = Resolve-ProxmoxCloneTargetNode -SourceNode 'n-a' -CloneName 'rr-1'
$rr2 = Resolve-ProxmoxCloneTargetNode -SourceNode 'n-a' -CloneName 'rr-2'
$rr3 = Resolve-ProxmoxCloneTargetNode -SourceNode 'n-a' -CloneName 'rr-3'
$rr4 = Resolve-ProxmoxCloneTargetNode -SourceNode 'n-a' -CloneName 'rr-4'
Assert (@($rr1, $rr2, $rr3) -join ',' -eq 'n-a,n-b,n-c') "D1: round_robin cycles through all eligible nodes in name order regardless of Proxmox's own return order"
Assert ($rr4 -eq 'n-a') "D2: ...and wraps back around to the first node after a full cycle"

# Excluding the middle node must not break the cycle -- it's simply skipped.
$script:ClusterNodesCache = $null; $script:ClusterNodesCachedAt = $null
$script:Settings.cloning.load_balancing.excluded_nodes = @('n-b')
$script:PlacementRoundRobinIndex = 0
$rrE1 = Resolve-ProxmoxCloneTargetNode -SourceNode 'n-a' -CloneName 'rr-excl-1'
$rrE2 = Resolve-ProxmoxCloneTargetNode -SourceNode 'n-a' -CloneName 'rr-excl-2'
Assert (@($rrE1, $rrE2) -join ',' -eq 'n-a,n-c') "D3: an excluded node is skipped in the round-robin cycle, not just left as a gap"
$script:Settings.cloning.load_balancing.excluded_nodes = @()
$script:Settings.cloning.load_balancing.strategy = 'resource'

# ------------------------------------------------------------
# E. Recreation: preserve_node_on_recreation lands a same-named reclone back
#    on its prior node, skipping strategy selection entirely; falls through
#    to strategy if that node is no longer eligible; a genuinely new name
#    always goes through strategy.
# ------------------------------------------------------------
$script:ClusterNodesCache = $null; $script:ClusterNodesCachedAt = $null
$script:nodeFixtures = @(
    [pscustomobject]@{ type = 'node'; node = 'pve-node1'; status = 'online'; cpu = 0.9; maxcpu = 4; mem = 90; maxmem = 100 }   # worst score
    [pscustomobject]@{ type = 'node'; node = 'pve-node2'; status = 'online'; cpu = 0.1; maxcpu = 4; mem = 10; maxmem = 100 }   # best score -- strategy would pick THIS
)
$script:RecentlyDeletedNodeByName['w10pwsh-recreate'] = @{ node = 'pve-node1'; deleted_at = [DateTime]::UtcNow.AddSeconds(-30) }
$preserveResult = Resolve-ProxmoxCloneTargetNode -SourceNode 'pve-node2' -CloneName 'w10pwsh-recreate'
Assert ($preserveResult -eq 'pve-node1') "E1: a same-named reclone within the retention window lands back on its PRIOR node, even though strategy would have picked the other one"

# A genuinely new name (never in $script:RecentlyDeletedNodeByName) always
# goes through strategy selection.
$script:ClusterNodesCache = $null; $script:ClusterNodesCachedAt = $null
$newNameResult = Resolve-ProxmoxCloneTargetNode -SourceNode 'pve-node2' -CloneName 'never-seen-before'
Assert ($newNameResult -eq 'pve-node2') "E2: a name never seen before goes through strategy selection normally (picks the best-scored node)"

# Outside the retention window -- treated as a genuinely new name.
$script:ClusterNodesCache = $null; $script:ClusterNodesCachedAt = $null
$script:RecentlyDeletedNodeByName['w10pwsh-stale'] = @{ node = 'pve-node1'; deleted_at = [DateTime]::UtcNow.AddSeconds(-301) }
$script:Settings.cloning.timeouts.recently_deleted_retention_seconds = 300
$staleResult = Resolve-ProxmoxCloneTargetNode -SourceNode 'pve-node2' -CloneName 'w10pwsh-stale'
Assert ($staleResult -eq 'pve-node2') "E3: outside cloning.timeouts.recently_deleted_retention_seconds, a previously-deleted name is treated as new -- strategy selection applies"

# preserve_node_on_recreation = false ignores the prior-node memory entirely.
$script:ClusterNodesCache = $null; $script:ClusterNodesCachedAt = $null
$script:RecentlyDeletedNodeByName['w10pwsh-off'] = @{ node = 'pve-node1'; deleted_at = [DateTime]::UtcNow.AddSeconds(-30) }
$script:Settings.cloning.load_balancing.preserve_node_on_recreation = $false
$preserveOffResult = Resolve-ProxmoxCloneTargetNode -SourceNode 'pve-node2' -CloneName 'w10pwsh-off'
Assert ($preserveOffResult -eq 'pve-node2') "E4: preserve_node_on_recreation=false ignores prior-node memory even for a matching, in-window name"
$script:Settings.cloning.load_balancing.preserve_node_on_recreation = $true

# A preserved node that has since gone offline falls through to strategy
# instead of being forced onto a bad node.
$script:ClusterNodesCache = $null; $script:ClusterNodesCachedAt = $null
$script:nodeFixtures = @(
    [pscustomobject]@{ type = 'node'; node = 'pve-node1'; status = 'offline'; cpu = 0.9; maxcpu = 4; mem = 90; maxmem = 100 }
    [pscustomobject]@{ type = 'node'; node = 'pve-node2'; status = 'online'; cpu = 0.1; maxcpu = 4; mem = 10; maxmem = 100 }
)
$script:RecentlyDeletedNodeByName['w10pwsh-gone-offline'] = @{ node = 'pve-node1'; deleted_at = [DateTime]::UtcNow.AddSeconds(-30) }
$offlinePreserveResult = Resolve-ProxmoxCloneTargetNode -SourceNode 'pve-node2' -CloneName 'w10pwsh-gone-offline'
Assert ($offlinePreserveResult -eq 'pve-node2') "E5: a preserved node that has since gone offline falls through to strategy-based selection instead of being forced"

# ------------------------------------------------------------
# F. Get-ProxmoxClusterNodes caching -- reused within the TTL, refetched once
#    it expires. Separate cache from the VM listing (Get-ProxmoxClusterVMs).
# ------------------------------------------------------------
$script:ClusterNodesCache = $null; $script:ClusterNodesCachedAt = $null
$script:nodeFixtures = @(
    [pscustomobject]@{ type = 'node'; node = 'n-cache'; status = 'online'; cpu = 0.1; maxcpu = 1; mem = 1; maxmem = 100 }
)
$script:Settings.cloning.load_balancing.node_stats_cache_ttl_seconds = 15
$script:nodeCallCount = 0
[void](Get-ProxmoxClusterNodes)
[void](Get-ProxmoxClusterNodes)
Assert ($script:nodeCallCount -eq 1) "F1: a second call within the TTL is served from cache -- no second HTTP call"
$script:ClusterNodesCachedAt = [DateTime]::UtcNow.AddSeconds(-16)
[void](Get-ProxmoxClusterNodes)
Assert ($script:nodeCallCount -eq 2) "F2: a call after the TTL expires re-fetches"

# ------------------------------------------------------------
# G. Handle-GuestClone integration -- 'target' actually lands in the clone
#    body when placement resolves one, and is entirely absent when placement
#    is disabled (exactly today's pre-feature behavior, not target=<source>).
# ------------------------------------------------------------
$script:ClusterNodesCache = $null; $script:ClusterNodesCachedAt = $null
$script:nodeFixtures = @(
    [pscustomobject]@{ type = 'node'; node = 'pve-node1'; status = 'online'; cpu = 0.9; maxcpu = 4; mem = 90; maxmem = 100 }
    [pscustomobject]@{ type = 'node'; node = 'pve-node2'; status = 'online'; cpu = 0.1; maxcpu = 4; mem = 10; maxmem = 100 }
)
$script:Settings.cloning.load_balancing.enabled = $true
$script:lastCloneBody = $null
$rClone = Handle-GuestClone -Params ([pscustomobject]@{ id = '153'; name = 'placement-int-1' })
Assert ($script:lastCloneBody.ContainsKey('target')) "G1: the clone body carries 'target' when placement is enabled and resolves a node"
Assert ($script:lastCloneBody.target -eq 'pve-node2') "G2: ...and it's the best-scored eligible node"
$newVmId1 = [string]$rClone.result.clone_id
$persistedEntry1 = Get-CloneStateEntry -VmId $newVmId1
Assert ($persistedEntry1.clone_node -eq 'pve-node2') "G3: the persisted clone-state entry's clone_node reflects the TARGET node, not the source's"
Clear-LastClone $rClone

$script:ClusterNodesCache = $null; $script:ClusterNodesCachedAt = $null
$script:Settings.cloning.load_balancing.enabled = $false
$script:lastCloneBody = $null
$rClone2 = Handle-GuestClone -Params ([pscustomobject]@{ id = '153'; name = 'placement-int-2' })
Assert (-not $script:lastCloneBody.ContainsKey('target')) "G4: placement disabled -- the clone body carries no 'target' key at all, exactly today's pre-feature behavior"
$newVmId2 = [string]$rClone2.result.clone_id
$persistedEntry2 = Get-CloneStateEntry -VmId $newVmId2
Assert ($persistedEntry2.clone_node -eq 'pve-node1') "G5: with placement disabled, clone_node is the source's own node (pve-node1)"
Clear-LastClone $rClone2
$script:Settings.cloning.load_balancing.enabled = $true

Write-Host "`n=== t14 Summary: $pass passed, $fail failed ===" -ForegroundColor $(if ($fail -eq 0) { 'Green' } else { 'Red' })
if ($fail -gt 0) { exit 1 }
