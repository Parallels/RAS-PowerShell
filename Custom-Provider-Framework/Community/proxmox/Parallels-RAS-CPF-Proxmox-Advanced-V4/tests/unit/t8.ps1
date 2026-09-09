$ErrorActionPreference = 'Stop'
$pass = 0
$fail = 0
function Assert {
    param([bool]$Cond, [string]$Msg)
    if ($Cond) { $script:pass++; Write-Host "PASS: $Msg" -ForegroundColor Green }
    else { $script:fail++; Write-Host "FAIL: $Msg" -ForegroundColor Red }
}

# ============================================================
# TESTS not already covered by t6.ps1/t7.ps1, plus the settings-file
# loader/hot-reload:
#   Handle-GuestList must not hard-fail when a cluster/resources entry
#         has no 'name' property at all.
#   Bounded retry on transient Proxmox 5xx for guests/control.
#   Repair-ProxmoxCloneTagsAfterInherit reads LIVE config, not the
#         (possibly stale/empty) cached cluster/resources entry.
#   Orphan-candidate audit (log + tag, never delete).
#   settings loader: self-seed, corrupt-tolerant, coercion, hot-reload,
#         Locations excluded from hot-reload.
# ============================================================
. "$PSScriptRoot/funcs2.ps1"

$script:LogPath        = "$PSScriptRoot/t8.log"
$script:CloneStatePath = "$PSScriptRoot/t8-state.json"
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
$script:CloneTaskCompletionCache = @{}
$script:LastGuestPollAt = @{}
$script:OrphanAuditLastCheckedAt = [DateTime]::MinValue
$script:ControlRetryMaxAttempts = 3
$script:ControlRetryDelaySeconds = 0   # keep the test fast -- retry timing itself isn't what's under test
$script:OrphanDetectionEnabled = $true
$script:OrphanDetectionCheckIntervalSeconds = 300
$script:OrphanDetectionStalePollAfterSeconds = 180

$script:apiCalls = New-Object System.Collections.ArrayList
$script:putConfigCalls = New-Object System.Collections.ArrayList
$script:nextId = 800
$script:cloneJobs = @{}
$script:vmState = @{}
$script:vmName  = @{}
$script:vmVisible = @{}
$script:vmTags = @{}          # cluster/resources view (can be stale)
$script:vmConfigTags = @{}    # LIVE /qemu/{id}/config view -- what Repair-ProxmoxCloneTagsAfterInherit must read
$script:failNextControlCalls = 0   # how many subsequent control POSTs/DELETEs should fail with a transient 500

function Invoke-ProxmoxApi {
    param([string]$Method, [string]$Path, $Body, [int]$TimeoutSec = 0)
    [void]$script:apiCalls.Add("$Method $Path")

    if ($Path -match '/cluster/nextid') { return @{ data = "$($script:nextId)" } }

    if ($Method -eq 'GET' -and $Path -match '/qemu/(\d+)/config$') {
        $vmid = $Matches[1]
        # Falls back to the cluster/resources view (vmTags) when no test has
        # deliberately registered a divergent live value in vmConfigTags --
        # normally the two agree; Test C registers a divergent one to
        # simulate cache lag specifically.
        $tags = if ($script:vmConfigTags.ContainsKey($vmid)) { $script:vmConfigTags[$vmid] } elseif ($script:vmTags.ContainsKey($vmid)) { $script:vmTags[$vmid] } else { '' }
        return @{ data = [pscustomobject]@{ tags = $tags } }
    }

    if ($Method -eq 'PUT' -and $Path -match '/qemu/(\d+)/config$') {
        $vmid = $Matches[1]
        $script:vmTags[$vmid] = [string]$Body.tags
        $script:vmConfigTags[$vmid] = [string]$Body.tags
        [void]$script:putConfigCalls.Add(@{ VmId = $vmid; Tags = [string]$Body.tags })
        return @{ data = $null }
    }

    if ($Path -match '/cluster/resources') {
        $list = New-Object System.Collections.ArrayList
        foreach ($k in $script:vmState.Keys) {
            if (-not $script:vmVisible[$k]) { continue }
            $entry = [pscustomobject]@{ type = 'qemu'; vmid = [int]$k; node = 'n1'; status = $script:vmState[$k]; template = 0 }
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
        $script:vmVisible[$newid] = $true
        return @{ data = $tid }
    }

    if ($Path -match '/qemu/(\d+)/status/current$') {
        $vmid = $Matches[1]
        $st = if ($script:vmState.ContainsKey($vmid)) { $script:vmState[$vmid] } else { 'stopped' }
        return @{ data = [pscustomobject]@{ status = $st; qmpstatus = $st } }
    }

    if ($Method -eq 'POST' -and $Path -match '/qemu/(\d+)/status/(stop|start|reboot|reset|suspend|resume)$') {
        if ($script:failNextControlCalls -gt 0) {
            $script:failNextControlCalls--
            throw 'Response status code does not indicate success: 500 (got no worker upid - start worker failed).'
        }
        $vmid = $Matches[1]
        $action = $Matches[2]
        $script:vmState[$vmid] = if ($action -eq 'stop') { 'stopped' } else { 'running' }
        return @{ data = "UPID:n1:$action$vmid" }
    }

    if ($Method -eq 'DELETE' -and $Path -match '/qemu/(\d+)$') {
        if ($script:failNextControlCalls -gt 0) {
            $script:failNextControlCalls--
            throw 'Response status code does not indicate success: 500 (got no worker upid - start worker failed).'
        }
        $vmid = $Matches[1]
        $script:vmVisible[$vmid] = $false
        return @{ data = $null }
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
# Test A: Handle-GuestList must not hard-fail when a cluster/resources entry
# has no 'name' property at all -- observed live causing a fraction of
# guests/list polls to fail outright during a clone burst.
# ------------------------------------------------------------
$noNameId = '810'
$script:vmState[$noNameId] = 'stopped'
$script:vmVisible[$noNameId] = $true
# Deliberately never set $script:vmName[$noNameId] -- the mock only attaches
# a 'name' property when one is set, exactly reproducing a cluster/resources
# entry missing the property entirely (not just blank).
Reset-ProxmoxClusterCache

$listResultA = Handle-GuestList
Assert ($null -ne $listResultA.result) "guests/list does not return an error when an entry has no 'name' property at all"
Assert (@($listResultA.result.guests) -contains $noNameId) "the nameless VM is still listed (empty name is not itself a placeholder)"

$script:vmState.Remove($noNameId); $script:vmVisible.Remove($noNameId)
Reset-ProxmoxClusterCache

# ------------------------------------------------------------
# Test B: bounded retry on transient Proxmox 5xx for guests/control
# ------------------------------------------------------------
$ctlId = '820'
$script:vmState[$ctlId] = 'running'
$script:vmName[$ctlId] = 'retry-test'
$script:vmVisible[$ctlId] = $true
Reset-ProxmoxClusterCache

# Fails twice, succeeds on the 3rd attempt -- within control_retry_max_attempts (3).
$script:failNextControlCalls = 2
$script:apiCalls.Clear()
$stopResult = Handle-GuestControl -Params ([pscustomobject]@{ id = $ctlId; control = 'stop' })
Assert (-not $stopResult.ContainsKey('error')) "guests/control succeeds after 2 transient 500s, within the retry budget"
Assert ($script:vmState[$ctlId] -eq 'stopped') "the retried stop actually took effect"
$stopCallCount = @($script:apiCalls | Where-Object { $_ -like 'POST *status/stop*' }).Count
Assert ($stopCallCount -eq 3) "exactly 3 attempts were made (2 failures + 1 success) -- got $stopCallCount"

# Exceeds the retry budget -- every attempt fails, so the control call must
# still surface the error to RAS rather than retrying forever.
$script:vmState[$ctlId] = 'running'
$script:failNextControlCalls = 99
$script:apiCalls.Clear()
$stopResult2 = Handle-GuestControl -Params ([pscustomobject]@{ id = $ctlId; control = 'stop' })
Assert ($stopResult2.ContainsKey('error')) "guests/control still reports failure once every attempt is exhausted"
$stopCallCount2 = @($script:apiCalls | Where-Object { $_ -like 'POST *status/stop*' }).Count
Assert ($stopCallCount2 -eq 3) "gives up after exactly control_retry_max_attempts (3) attempts, not fewer or more -- got $stopCallCount2"

$script:failNextControlCalls = 0
$script:vmState.Remove($ctlId); $script:vmName.Remove($ctlId); $script:vmVisible.Remove($ctlId)
Reset-ProxmoxClusterCache

# ------------------------------------------------------------
# Test C: Repair-ProxmoxCloneTagsAfterInherit reads LIVE config, not the
# (possibly stale/empty) cached cluster/resources entry
# ------------------------------------------------------------
$srcId = '830'
$cloneId = '831'
# The cached cluster/resources entry (what $ClusterVm below represents) has
# NO tags at all -- simulating the server-side aggregate not having caught
# up yet for a just-created clone, exactly as measured live.
$staleClusterVm = [pscustomobject]@{ vmid = [int]$cloneId; node = 'n1'; status = 'stopped'; template = 0 }
# The LIVE config, however, already shows what the clone actually inherited
# from its source at clone time: a hand-set 'cat' tag plus rasTemplate830.
$script:vmConfigTags[$cloneId] = 'cat;rasTemplate830'

$repairResult = Repair-ProxmoxCloneTagsAfterInherit -Node 'n1' -VmId $cloneId -ClusterVm $staleClusterVm -CloneTag 'rasClone830'
Assert ($repairResult.changed) "a repair is recorded (rasTemplate830 was present)"
Assert (@($repairResult.removed_inherited_template_tags) -contains 'rasTemplate830') "rasTemplate830 was identified and reported as stripped"
$finalTags = @(($script:vmConfigTags[$cloneId]) -split ';')
Assert ($finalTags -contains 'cat') "the hand-set 'cat' tag survives -- it must NOT be wiped just because the cached listing had no tags"
Assert ($finalTags -contains 'rasClone830') "rasClone830 was added"
Assert ($finalTags -notcontains 'rasTemplate830') "the inherited rasTemplate830 was stripped"
Assert ($finalTags.Count -eq 2) "no extra/duplicate tags -- got [$($finalTags -join ',')]"

# Idempotency: re-running against the now-correct live tags issues no PUT.
$script:putConfigCalls.Clear()
$repairResult2 = Repair-ProxmoxCloneTagsAfterInherit -Node 'n1' -VmId $cloneId -ClusterVm $staleClusterVm -CloneTag 'rasClone830'
Assert (-not $repairResult2.changed) "re-running against already-correct live tags is a no-op"
Assert ($script:putConfigCalls.Count -eq 0) "no PUT issued on the idempotent re-run"

# ------------------------------------------------------------
# Test D: orphan-candidate audit -- log + tag only, never delete
# ------------------------------------------------------------
$orphanId = '840'
$freshId = '841'
$trackedId = '842'

foreach ($id in @($orphanId, $freshId, $trackedId)) {
    $script:vmState[$id] = 'running'
    $script:vmName[$id] = "clone-$id"
    $script:vmVisible[$id] = $true
    $script:vmTags[$id] = 'rasClone153'
}
# orphanId: this provider is not tracking it, and RAS hasn't polled it in
# well over stale_poll_after_seconds -- the orphan case.
$script:LastGuestPollAt[$orphanId] = [DateTime]::UtcNow.AddSeconds(-999)
# freshId: not tracked, but RAS polled it recently -- must NOT be flagged.
$script:LastGuestPollAt[$freshId] = [DateTime]::UtcNow
# trackedId: still an in-flight clone from this provider's own point of view
# -- must NOT be flagged even though it also carries the tag and has no
# recent poll recorded.
[void]$script:TrackedCloneVmIds.Add($trackedId)

Reset-ProxmoxClusterCache
# Test A's call to Handle-GuestList already ran Invoke-OrphanAudit internally
# (that's the wiring under test), which primed the throttle gate -- reset it
# here so THIS call is the one actually being asserted on, not silently
# throttled by that earlier, incidental call.
$script:OrphanAuditLastCheckedAt = [DateTime]::MinValue
$clusterVMsForAudit = Get-ProxmoxClusterVMs
Invoke-OrphanAudit -ClusterVMs $clusterVMsForAudit

$logContent = Get-Content -Path $script:LogPath -Raw
Assert ($logContent -match "ORPHAN CANDIDATE: VM \[$orphanId\]") "a stale, untracked, rasClone-tagged VM is logged as an orphan candidate"
Assert ($logContent -notmatch "ORPHAN CANDIDATE: VM \[$freshId\]") "a recently-polled VM is NOT flagged"
Assert ($logContent -notmatch "ORPHAN CANDIDATE: VM \[$trackedId\]") "a VM this provider is still tracking as in-flight is NOT flagged"
Assert (@(($script:vmTags[$orphanId]) -split ';') -contains 'rasOrphanCandidate') "the orphan candidate is tagged rasOrphanCandidate (log-and-tag only, never deleted)"
Assert ($script:vmVisible[$orphanId]) "the orphan candidate was NOT deleted or stopped -- log-and-tag only"

# Throttling: a second audit call immediately after must not re-scan (no new
# tag-write attempt, no duplicate log line) until check_interval_seconds has
# elapsed.
# Undo the tag in BOTH the cluster/resources view and the live-config view
# (Add-ProxmoxVmTag now reads the latter -- see Get-ProxmoxVmLiveTagList) so
# a re-tag attempt would actually be visible if the throttle failed to hold.
$script:vmTags[$orphanId] = 'rasClone153'
$script:vmConfigTags[$orphanId] = 'rasClone153'
$script:putConfigCalls.Clear()
Invoke-OrphanAudit -ClusterVMs $clusterVMsForAudit
Assert ($script:putConfigCalls.Count -eq 0) "an immediate second audit call is throttled (check_interval_seconds) -- no re-scan"

$script:OrphanAuditLastCheckedAt = [DateTime]::UtcNow.AddSeconds(-301)
Invoke-OrphanAudit -ClusterVMs $clusterVMsForAudit
Assert ($script:putConfigCalls.Count -eq 1) "after the throttle window elapses, the audit runs again"

foreach ($id in @($orphanId, $freshId, $trackedId)) {
    $script:vmState.Remove($id); $script:vmName.Remove($id); $script:vmVisible.Remove($id); $script:vmTags.Remove($id); $script:LastGuestPollAt.Remove($id)
}
[void]$script:TrackedCloneVmIds.Remove($trackedId)
Reset-ProxmoxClusterCache

# ============================================================
# Settings loader: self-seed, corrupt-tolerant, coercion, hot-reload
# ============================================================
$settingsDir = Join-Path $PSScriptRoot 'settings-test'
Remove-Item -Path $settingsDir -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Path $settingsDir -Force | Out-Null
$script:SettingsPath = Join-Path $settingsDir 'RAS-CPF-Proxmox-Settings.json'
$script:Settings = $null
$script:SettingsFileLastWriteUtc = $null
$script:SettingsLastPolledAt = [DateTime]::MinValue

# ------------------------------------------------------------
# Test E: missing file self-seeds with exactly the hard-coded defaults
# ------------------------------------------------------------
Assert (-not (Test-Path $script:SettingsPath)) "settings: no file exists yet"
Import-ProviderSettings
Assert (Test-Path $script:SettingsPath) "settings: a missing file is self-seeded on load"
Assert ($script:Settings.cloning.pipelined_cloning.max_concurrent_clone_operations -eq 2) "settings: self-seeded in-memory defaults match the hard-coded value (max_concurrent_clone_operations=2)"
Assert ($script:Settings.schema_version -eq $script:CurrentSettingsSchemaVersion) "settings: self-seeded in-memory Settings records the current schema_version"
$seededJson = Get-Content -Path $script:SettingsPath -Raw | ConvertFrom-Json
Assert ($seededJson.schema_version -eq $script:CurrentSettingsSchemaVersion) "settings: the seeded FILE's own schema_version matches"
# The FILE is the rich {value, default, description} shape -- see
# Get-DefaultProviderSettings -- so every leaf there is an object, not a bare
# scalar; only $script:Settings (in-memory, asserted above) is flat.
Assert ($seededJson.cloning.pipelined_cloning.max_concurrent_clone_operations.value -eq 2) "settings: the seeded FILE on disk also matches (round-trips correctly, rich {value,...} shape)"
Assert ($seededJson.cloning.pipelined_cloning.max_concurrent_clone_operations.default -eq 2) "settings: the seeded FILE's 'default' field matches 'value' when nothing has been customized"
Assert (-not [string]::IsNullOrWhiteSpace([string]$seededJson.cloning.pipelined_cloning.max_concurrent_clone_operations.description)) "settings: the seeded FILE carries a non-empty 'description' for an admin to read without opening SETTINGS.md"
Assert ($seededJson.locations.log_file.value -eq 'Proxmox-RAS-Provider.log') "settings: seeded file has the expected default log_file name"

# ------------------------------------------------------------
# Test F: a valid custom file overrides individual keys, independently coerced.
# Deliberately mixes both leaf shapes in one file -- the rich {value,...}
# object (as the self-seeded file itself uses) and a bare scalar (a
# hand-simplified value, or a pre-schema-2 file) -- both must coerce
# identically; ConvertTo-CoercedSetting*/Get-SettingLeafRawValue must not
# care which one it's given.
# ------------------------------------------------------------
@'
{
  "cloning": {
    "pipelined_cloning": { "max_concurrent_clone_operations": { "value": 5, "default": 2, "description": "overridden for this test" } }
  },
  "virtual_machines": {
    "tags": { "ras_clone_tag_prefix": "myClone" },
    "guest_agent": { "quarantine_min_failures": "7" }
  },
  "logging": { "log_level": "extended" }
}
'@ | Set-Content -Path $script:SettingsPath -Encoding UTF8

Import-ProviderSettings
Set-ProviderRuntimeFromSettings -Settings $script:Settings
Assert ($script:MaxConcurrentCloneOperations -eq 5) "settings: a custom cloning.pipelined_cloning.max_concurrent_clone_operations (rich {value,...} shape) is applied"
Assert ($script:RasCloneTagPrefix -eq 'myClone') "settings: a custom virtual_machines.tags.ras_clone_tag_prefix (bare scalar shape) is applied"
Assert ($script:LogLevel -eq 4) "settings: log_level string alias 'extended' coerces to 4"
Assert ($script:AgentQuarantineMinFailures -eq 7) "settings: a JSON STRING '7' still coerces to an int"
# Untouched keys keep their hard-coded defaults.
Assert ($script:PipelinedCloneCompletionSeconds -eq 10) "settings: a key absent from the custom file falls back to its default"

# ------------------------------------------------------------
# Test F2: capabilities.can_link_clones migration fallback -- a schema-1-shaped
# file (no 'schema_version', the old cloning.linked_clones_enabled key, no
# 'capabilities' section at all) must still resolve capabilities.can_link_clones
# correctly, so an already-live linked-clone deployment does not silently
# revert to false the moment schema 2 ships. See SETTINGS.md's migration notes.
# ------------------------------------------------------------
'{ "cloning": { "linked_clones_enabled": true, "linked_clone_fallback": "full" } }' | Set-Content -Path $script:SettingsPath -Encoding UTF8
Import-ProviderSettings
Assert ($script:Settings.schema_version -eq 1) "settings: a file with no schema_version field is treated as schema 1"
# Schema-1 is deliberately excluded from auto-migration (see t18.ps1 for the
# migration mechanism's own dedicated coverage, and the comment on
# Import-ProviderSettings's own migration-trigger check) -- only
# capabilities.can_link_clones has a real fallback-read from its original
# location; every other schema-1 key would auto-migrate to a silently-defaulted value,
# discarding whatever an admin actually had. A genuine schema-1 file is detected and
# logged (this assertion), never auto-rewritten.
Assert ($script:Settings.capabilities.can_link_clones -eq $true) "settings: capabilities.can_link_clones falls back to the deprecated cloning.linked_clones_enabled when the new key is absent, even with no 'capabilities' section in the file at all"

# ------------------------------------------------------------
# Test G: a corrupted file is left completely untouched -- in-memory
# defaults are used, and the provider must not crash.
# ------------------------------------------------------------
'{ this is not valid json' | Set-Content -Path $script:SettingsPath -Encoding UTF8
$beforeCorruptContent = Get-Content -Path $script:SettingsPath -Raw
$corruptThrew = $false
try {
    Import-ProviderSettings
    Set-ProviderRuntimeFromSettings -Settings $script:Settings
}
catch { $corruptThrew = $true }
Assert (-not $corruptThrew) "settings: a corrupted file does not crash the provider"
Assert ($script:MaxConcurrentCloneOperations -eq 2) "settings: a corrupted file falls back to in-memory hard-coded defaults, not the last-good values"
$afterCorruptContent = Get-Content -Path $script:SettingsPath -Raw
Assert ($afterCorruptContent -eq $beforeCorruptContent) "settings: the corrupted file itself is left completely untouched (never overwritten)"

# ------------------------------------------------------------
# Test H: hot-reload picks up a changed file via Update-ProviderSettingsIfChanged
# ------------------------------------------------------------
@'
{ "cloning": { "pipelined_cloning": { "max_concurrent_clone_operations": 9 } } }
'@ | Set-Content -Path $script:SettingsPath -Encoding UTF8
Start-Sleep -Milliseconds 50   # ensure a distinct LastWriteTimeUtc from any prior write in this test

$script:SettingsLastPolledAt = [DateTime]::UtcNow.AddSeconds(-31)   # pretend the poll interval has already elapsed
Update-ProviderSettingsIfChanged
Assert ($script:MaxConcurrentCloneOperations -eq 9) "settings: hot-reload picks up a changed file once the poll interval has elapsed"

# Calling again immediately (poll interval NOT elapsed) must not re-read the
# file even if it changes again in between.
@'
{ "cloning": { "pipelined_cloning": { "max_concurrent_clone_operations": 42 } } }
'@ | Set-Content -Path $script:SettingsPath -Encoding UTF8
Update-ProviderSettingsIfChanged
Assert ($script:MaxConcurrentCloneOperations -eq 9) "settings: a second change is NOT picked up before the poll interval elapses again"

# ------------------------------------------------------------
# Test I: Locations changes are deliberately ignored on a hot reload (would
# orphan every in-flight clone's tracking if applied mid-run -- see
# Import-ProviderSettings's own comment)
# ------------------------------------------------------------
$originalLogPath = $script:LogPath
@'
{ "locations": { "log_file": "SomeOtherLog.log" }, "cloning": { "pipelined_cloning": { "max_concurrent_clone_operations": 11 } } }
'@ | Set-Content -Path $script:SettingsPath -Encoding UTF8
Start-Sleep -Milliseconds 50
$script:SettingsLastPolledAt = [DateTime]::UtcNow.AddSeconds(-31)
Update-ProviderSettingsIfChanged
Assert ($script:LogPath -eq $originalLogPath) "settings: a Locations change is ignored on hot-reload (no restart happened)"
Assert ($script:MaxConcurrentCloneOperations -eq 11) "settings: every OTHER section in the same reload is still applied"

Remove-Item -Path $settingsDir -Recurse -Force -ErrorAction SilentlyContinue

# ------------------------------------------------------------
# Test J: log_level actually gates what gets written -- 3 (Standard) suppresses
# Extended/Verbose-tagged lines but never an untagged (default Verbose) one
# that a caller explicitly marked -Level 3, and never suppresses anything
# below the active level.
# ------------------------------------------------------------
$levelLogPath = Join-Path $PSScriptRoot 'level-test.log'
Remove-Item -Path $levelLogPath -ErrorAction SilentlyContinue
$script:LogPath = $levelLogPath

$script:LogLevel = 5
Write-DebugLog "verbose-line-marker" -Level 'D'
Write-DebugLog "extended-line-marker" -Level 'T'
Write-DebugLog "standard-line-marker" -Level 'I'
$c5 = Get-Content -Path $levelLogPath -Raw
Assert ($c5 -match 'verbose-line-marker') "log level 5 (Verbose): a Level-D line is written"
Assert ($c5 -match 'extended-line-marker') "log level 5 (Verbose): a Level-T line is written"
Assert ($c5 -match 'standard-line-marker') "log level 5 (Verbose): a Level-I line is written"

Remove-Item -Path $levelLogPath -ErrorAction SilentlyContinue
$script:LogLevel = 4
Write-DebugLog "verbose-line-marker" -Level 'D'
Write-DebugLog "extended-line-marker" -Level 'T'
Write-DebugLog "standard-line-marker" -Level 'I'
$c4 = Get-Content -Path $levelLogPath -Raw
Assert (-not ($c4 -match 'verbose-line-marker')) "log level 4 (Extended): a Level-D line is suppressed"
Assert ($c4 -match 'extended-line-marker') "log level 4 (Extended): a Level-T line is written"
Assert ($c4 -match 'standard-line-marker') "log level 4 (Extended): a Level-I line is written"

Remove-Item -Path $levelLogPath -ErrorAction SilentlyContinue
$script:LogLevel = 3
Write-DebugLog "verbose-line-marker" -Level 'D'
Write-DebugLog "extended-line-marker" -Level 'T'
Write-DebugLog "standard-line-marker" -Level 'I'
$c3 = Get-Content -Path $levelLogPath -Raw
Assert (-not ($c3 -match 'verbose-line-marker')) "log level 3 (Standard): a Level-D line is suppressed"
Assert (-not ($c3 -match 'extended-line-marker')) "log level 3 (Standard): a Level-T line is suppressed"
Assert ($c3 -match 'standard-line-marker') "log level 3 (Standard): a Level-I line is still written"

# E and W bypass log_level entirely -- always written, even at the lowest
# configured level.
Remove-Item -Path $levelLogPath -ErrorAction SilentlyContinue
Write-DebugLog "error-line-marker" -Level 'E'
Write-DebugLog "warning-line-marker" -Level 'W'
$cAlwaysOn = Get-Content -Path $levelLogPath -Raw
Assert ($cAlwaysOn -match 'error-line-marker') "log level 3 (Standard): a Level-E line is written regardless"
Assert ($cAlwaysOn -match 'warning-line-marker') "log level 3 (Standard): a Level-W line is written regardless"

# A call site with no explicit -Level defaults to 'D' (Verbose) -- unclassified
# call sites keep logging exactly as they always have at the default log
# level, and are only suppressed if the active level is deliberately lowered.
$script:LogLevel = 5
Remove-Item -Path $levelLogPath -ErrorAction SilentlyContinue
Write-DebugLog "unclassified-line-marker"
$cDefaultLevel = Get-Content -Path $levelLogPath -Raw
Assert ($cDefaultLevel -match 'unclassified-line-marker') "log level 5 (Verbose, the default): a call site with no explicit -Level still logs"

$script:LogLevel = 3
Remove-Item -Path $levelLogPath -ErrorAction SilentlyContinue
Write-DebugLog "unclassified-line-marker"
$cUnclassified = if (Test-Path $levelLogPath) { Get-Content -Path $levelLogPath -Raw } else { '' }
Assert ([string]::IsNullOrEmpty($cUnclassified)) "log level 3 (Standard): a call site with no explicit -Level (defaults to D/Verbose) is suppressed"

$script:LogLevel = 5
Remove-Item -Path $levelLogPath -ErrorAction SilentlyContinue

Write-Host "`n=== t8 Summary: $pass passed, $fail failed ===" -ForegroundColor $(if ($fail -eq 0) { 'Green' } else { 'Red' })
if ($fail -gt 0) { exit 1 }
