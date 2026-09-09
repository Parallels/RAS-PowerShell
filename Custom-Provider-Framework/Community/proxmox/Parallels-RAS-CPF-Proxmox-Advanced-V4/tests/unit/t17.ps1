$ErrorActionPreference = 'Stop'
$pass = 0
$fail = 0
function Assert {
    param([bool]$Cond, [string]$Msg)
    if ($Cond) { $script:pass++; Write-Host "PASS: $Msg" -ForegroundColor Green }
    else { $script:fail++; Write-Host "FAIL: $Msg" -ForegroundColor Red }
}

# ============================================================
# TESTS for MAC address preservation on recreation (MAC-PRESERVATION.md):
#   cloning.mac_preservation.enabled (off by default) -- capture a deleted
#   VM's NIC MAC(s) via a live config read, then restore them onto a clone
#   made under the SAME name before it is ever started.
# ============================================================
. "$PSScriptRoot/funcs2.ps1"

$script:LogPath        = "$PSScriptRoot/t17.log"
$script:CloneStatePath = "$PSScriptRoot/t17-state.json"
Remove-Item $script:LogPath, $script:CloneStatePath -ErrorAction SilentlyContinue
$script:TrackedCloneVmIds = [System.Collections.Generic.HashSet[string]]::new()
$script:ClusterResourcesCache = $null
$script:ClusterResourcesCachedAt = $null
$script:RecentlyDeletedIds = @{}
$script:RecentlyDeletedNames = @{}
$script:RecentlyDeletedNodeByName = @{}
$script:RecentlyDeletedMacByName = @{}
$script:RecentlyControlledIds = @{}
$script:RecentlyStoppedIds = @{}
$script:AgentFailureTracker = @{}
$script:LastKnownNetworkData = @{}
$script:SweepLastCheckedAt = @{}
$script:RasTemplateTagAttempted = [System.Collections.Generic.HashSet[string]]::new()
$script:CloneTagVerified = [System.Collections.Generic.HashSet[string]]::new()
$script:CloneMacRestored = [System.Collections.Generic.HashSet[string]]::new()
$script:CloneStateMemory = $null
$script:TaskContext = @{}
$script:CloneTaskCompletionCache = @{}

$script:RecentlyDeletedRetentionSeconds = 300
$script:PreserveMacOnRecreation = $false   # each test section sets this explicitly

$script:apiCalls = New-Object System.Collections.ArrayList
$script:nextId = 800
$script:vmState  = @{}   # vmid -> status
$script:vmName   = @{}
$script:vmVisible = @{}
$script:vmNet    = @{}   # vmid -> @{ net0 = "virtio=AA:BB:...,bridge=vmbr0,firewall=1"; ... }
$script:configGetCount = 0
$script:lastConfigPutBody = $null

function Invoke-ProxmoxApi {
    param([string]$Method, [string]$Path, $Body, [int]$TimeoutSec = 0)
    [void]$script:apiCalls.Add("$Method $Path")

    if ($Path -match '/cluster/nextid') { return @{ data = "$($script:nextId)" } }

    if ($Path -match '/cluster/resources') {
        $list = New-Object System.Collections.ArrayList
        [void]$list.Add([pscustomobject]@{ type = 'qemu'; vmid = 900; name = 'template'; node = 'n1'; status = 'stopped'; template = 1 })
        foreach ($k in $script:vmState.Keys) {
            if (-not $script:vmVisible[$k]) { continue }
            [void]$list.Add([pscustomobject]@{ type = 'qemu'; vmid = [int]$k; name = $script:vmName[$k]; node = 'n1'; status = $script:vmState[$k]; template = 0 })
        }
        return @{ data = $list.ToArray() }
    }

    if ($Method -eq 'GET' -and $Path -match '/qemu/(\d+)/config$') {
        $script:configGetCount++
        $id = $Matches[1]
        $cfg = @{}
        if ($script:vmNet.ContainsKey($id)) {
            foreach ($k in $script:vmNet[$id].Keys) { $cfg[$k] = $script:vmNet[$id][$k] }
        }
        return @{ data = [pscustomobject]$cfg }
    }

    if ($Method -eq 'PUT' -and $Path -match '/qemu/(\d+)/config$') {
        $id = $Matches[1]
        $script:lastConfigPutBody = $Body
        if (-not $script:vmNet.ContainsKey($id)) { $script:vmNet[$id] = @{} }
        foreach ($k in $Body.Keys) { $script:vmNet[$id][$k] = [string]$Body[$k] }
        return @{ data = $null }
    }

    if ($Path -match '/qemu/(\d+)/clone$') {
        $newid = [string]$Body.newid
        $tid = "UPID:n1:clone$newid"
        $script:vmState[$newid] = 'stopped'
        $script:vmName[$newid]  = [string]$Body.name
        $script:vmVisible[$newid] = $true
        # Proxmox assigns a fresh random MAC on clone -- simulate that, distinct from
        # whatever the source/old VM had, so a restored MAC is unambiguously observable.
        $script:vmNet[$newid] = @{ net0 = 'virtio=FF:FF:FF:FF:FF:FF,bridge=vmbr0,firewall=1' }
        return @{ data = $tid }
    }

    if ($Path -match '/qemu/(\d+)/status/current$') {
        $id = $Matches[1]
        $st = if ($script:vmState.ContainsKey($id)) { $script:vmState[$id] } else { 'stopped' }
        return @{ data = [pscustomobject]@{ status = $st } }
    }

    if ($Path -match '/qemu/(\d+)/status/start$') {
        $id = $Matches[1]
        $script:vmState[$id] = 'running'
        return @{ data = "UPID:n1:start$id" }
    }

    if ($Path -match '/qemu/(\d+)/status/stop$') {
        $id = $Matches[1]
        $script:vmState[$id] = 'stopped'
        return @{ data = "UPID:n1:stop$id" }
    }

    if ($Method -eq 'DELETE' -and $Path -match '/qemu/(\d+)$') {
        $id = $Matches[1]
        $script:vmState.Remove($id) | Out-Null
        $script:vmVisible.Remove($id) | Out-Null
        return @{ data = "UPID:n1:destroy$id" }
    }

    if ($Path -match '/tasks/(.+)/status$') {
        return @{ data = [pscustomobject]@{ status = 'stopped'; exitstatus = 'OK' } }
    }

    if ($Path -match '/agent/network-get-interfaces') {
        throw "500 no agent"
    }

    throw "Unexpected call in mock: $Method $Path"
}

# ------------------------------------------------------------
# A. Feature OFF (the default): delete does NOT capture a MAC (no extra
#    config-read cost), and a same-name recreate keeps Proxmox's own
#    freshly-assigned MAC untouched.
# ------------------------------------------------------------
Write-Host "`n--- Test A: mac_preservation.enabled = false (default) ---" -ForegroundColor Cyan
$script:PreserveMacOnRecreation = $false
$script:vmState['100'] = 'stopped'; $script:vmName['100'] = 'off-test'; $script:vmVisible['100'] = $true
$script:vmNet['100'] = @{ net0 = 'virtio=AA:AA:AA:AA:AA:AA,bridge=vmbr0,firewall=1' }
Reset-ProxmoxClusterCache
$configCountBefore = $script:configGetCount
[void](Handle-GuestControl -Params ([pscustomobject]@{ id = '100'; control = 'delete' }))
Assert ($script:configGetCount -eq $configCountBefore) "A1: feature off -- delete does not read the VM's config at all (zero extra cost)"
Assert (-not $script:RecentlyDeletedMacByName.ContainsKey('off-test')) "A2: feature off -- nothing captured into RecentlyDeletedMacByName"

$script:nextId = 801
$rA = Handle-GuestClone -Params ([pscustomobject]@{ id = '900'; name = 'off-test' })
$newIdA = [string]$rA.result.clone_id
[void](Handle-TaskInfo -Params ([pscustomobject]@{ id = $rA.result.task_id }))
Reset-ProxmoxClusterCache
[void](Get-RasGuestObjectForCloneAwareFlow -VmId $newIdA)   # drives Start-ProxmoxVmIfNeeded
Assert ($script:vmNet[$newIdA].net0 -match 'FF:FF:FF:FF:FF:FF') "A3: feature off -- the clone keeps Proxmox's own fresh MAC, untouched"

# ------------------------------------------------------------
# B. Feature ON -- capture at delete, restore on same-name recreate, before
#    the VM is ever started.
# ------------------------------------------------------------
Write-Host "`n--- Test B: mac_preservation.enabled = true -- capture + restore ---" -ForegroundColor Cyan
$script:PreserveMacOnRecreation = $true
$script:vmState['200'] = 'stopped'; $script:vmName['200'] = 'preserve-test'; $script:vmVisible['200'] = $true
$script:vmNet['200'] = @{ net0 = 'virtio=BC:24:11:00:00:64,bridge=vmbr0,firewall=1' }
Reset-ProxmoxClusterCache
[void](Handle-GuestControl -Params ([pscustomobject]@{ id = '200'; control = 'delete' }))
Assert ($script:RecentlyDeletedMacByName.ContainsKey('preserve-test')) "B1: feature on -- delete captures the MAC into RecentlyDeletedMacByName"
Assert ($script:RecentlyDeletedMacByName['preserve-test'].macs.net0 -eq 'BC:24:11:00:00:64') "B2: ...with the exact MAC the old VM had"

$script:nextId = 802
$rB = Handle-GuestClone -Params ([pscustomobject]@{ id = '900'; name = 'preserve-test' })
$newIdB = [string]$rB.result.clone_id
Assert ($script:vmNet[$newIdB].net0 -match 'FF:FF:FF:FF:FF:FF') "B3: sanity -- right after cloning, Proxmox's fresh random MAC is still in place (not yet restored)"
[void](Handle-TaskInfo -Params ([pscustomobject]@{ id = $rB.result.task_id }))
Reset-ProxmoxClusterCache
[void](Get-RasGuestObjectForCloneAwareFlow -VmId $newIdB)   # drives Start-ProxmoxVmIfNeeded, restoration happens here
Assert ($script:vmNet[$newIdB].net0 -match 'BC:24:11:00:00:64') "B4: after the clone-completion poll (before start), the OLD MAC is restored"
Assert ($script:vmNet[$newIdB].net0 -match 'bridge=vmbr0') "B5: ...and everything else Proxmox set (bridge, firewall) is left exactly as the clone had it"
# Start-ProxmoxVmIfNeeded restores the MAC and THEN issues status/start within the same
# call -- the two are not independently observable from outside, so this checks the
# thing that actually matters: the restore call landed strictly before the start call in
# the recorded API sequence, not that the VM is observably unstarted afterward (it won't
# be, by design -- see LOGGING.md's 'PUT .../config' vs '.../status/start' order below).
$putIdx = ($script:apiCalls | Select-String -Pattern "PUT .*$newIdB/config" | Select-Object -First 1).LineNumber
$startIdx = ($script:apiCalls | Select-String -Pattern "POST .*$newIdB/status/start" | Select-Object -First 1).LineNumber
Assert ($null -ne $putIdx -and $null -ne $startIdx -and $putIdx -lt $startIdx) "B6: the MAC-restoring config PUT happens strictly before the start call, never after -- a MAC applied post-boot wouldn't reliably help DHCP/licensing"

$putCountBeforeRepoll = $script:apiCalls.Count
[void](Get-RasGuestObjectForCloneAwareFlow -VmId $newIdB)
Assert ($script:apiCalls[($script:apiCalls.Count - 1)] -notmatch "PUT .*$newIdB/config") "B7: idempotent -- a second poll for the same VM does not re-PUT the config (`$script:CloneMacRestored gate)"

# ------------------------------------------------------------
# C. A clone under a DIFFERENT name (not a recreation) gets no restoration
#    at all -- Proxmox's own fresh MAC stands.
# ------------------------------------------------------------
Write-Host "`n--- Test C: different name -- not a recreation ---" -ForegroundColor Cyan
$script:nextId = 803
$rC = Handle-GuestClone -Params ([pscustomobject]@{ id = '900'; name = 'brand-new-vm' })
$newIdC = [string]$rC.result.clone_id
[void](Handle-TaskInfo -Params ([pscustomobject]@{ id = $rC.result.task_id }))
Reset-ProxmoxClusterCache
[void](Get-RasGuestObjectForCloneAwareFlow -VmId $newIdC)
Assert ($script:vmNet[$newIdC].net0 -match 'FF:FF:FF:FF:FF:FF') "C1: a clone under a name nothing was deleted as keeps its fresh Proxmox MAC -- no incorrect restoration"

Write-Host "`n=== t17 Summary: $pass passed, $fail failed ===" -ForegroundColor $(if ($fail -eq 0) { 'Green' } else { 'Red' })
if ($fail -gt 0) { exit 1 }
