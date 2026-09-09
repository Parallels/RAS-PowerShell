<#
.SYNOPSIS
    Parallels RAS Custom Provider for Proxmox VE (Advanced)
.DESCRIPTION
    A Parallels RAS Custom Provider Framework (CPF) integration for Proxmox VE,
    speaking JSON-RPC over stdin/stdout. Beyond the core guest lifecycle (list, get,
    power control, template conversion) it adds full and linked clones, template
    versioning via snapshots, distributed clone placement across cluster nodes,
    pool-scoped visibility, MAC address preservation across a same-name recreate,
    guest-agent quarantine handling, orphaned-clone detection, automatic HTTP
    retry/timeout handling, and a self-seeding, schema-versioned JSON settings file
    for every tunable -- see README.md and docs/ for the full reference.

    Requires PowerShell 7 or later.
.NOTES
    File Name : Parallels-RAS-CPF-Proxmox-Advanced.ps1
    Settings  : RAS-CPF-Proxmox-Settings.json (self-seeding, sibling to this script)
.EXAMPLE
    .\Parallels-RAS-CPF-Proxmox-Advanced.ps1

    Sample JSON-RPC requests, one per line on stdin:

    {"method":"provider/connect","params":{"settings":{"host":"proxmox.example.com","username":"root@pam","token_name":"automation","token_secret":"XXX"}}}
    {"method":"guests/list"}
    {"method":"guests/get","params":{"id":"101"}}
    {"method":"guests/get","params":{"id":["101","102"]}}
    {"method":"guests/control","params":{"control":"start","id":"101"}}
    {"method":"guests/convert","params":{"id":"101","is_template":true}}
    {"method":"guests/clone","params":{"id":"101","name":"Clone of 101"}}
    {"method":"tasks/get","params":{"id":"<task_id>"}}
#>


Set-StrictMode -Version Latest

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$WarningPreference = 'SilentlyContinue'
$VerbosePreference = 'SilentlyContinue'
$InformationPreference = 'SilentlyContinue'

# Bootstrap default -- replaced by Set-ProviderRuntimeFromSettings once the settings
# file is loaded further down.
$script:CloneStatePath = Join-Path -Path $PSScriptRoot -ChildPath 'Proxmox-RAS-CloneState.json'

if ($Host.Name -notmatch 'ISE') {
    [Console]::InputEncoding = [System.Text.Encoding]::UTF8
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
}

$stdout = [Console]::OpenStandardOutput()
$writer = New-Object System.IO.StreamWriter($stdout, [System.Text.Encoding]::UTF8)
$writer.AutoFlush = $true
# Force LF: StreamWriter.WriteLine otherwise uses the platform newline (CRLF on
# Windows), so framing would differ by deployment OS.
$writer.NewLine = "`n"

$script:ProviderNamePrefix = 'Proxmox:'
# Single source of truth for both the startup banner (below) and provider/initialize's
# own `version` field (Handle-Initialize) -- deliberately one place, not two separately
# hardcoded literals that could drift apart. Follows semantic versioning; see
# CHANGELOG.md for what changed at each version.
$script:ProviderVersion = '1.0.0'
$script:LogPath = Join-Path -Path $PSScriptRoot -ChildPath 'Proxmox-RAS-Provider.log'   # bootstrap default; see above
$script:LogLevel = 5   # bootstrap default (Verbose) -- see Set-ProviderRuntimeFromSettings
# Bootstrap defaults for the three variables Write-DebugLog/Invoke-LogRotation read on
# every call (matching Get-DefaultProviderSettings's own log_rotate_max_mb=100/
# log_rotate_max_generations=3) -- Set-ProviderRuntimeFromSettings overwrites these once
# settings actually load, but that happens only AFTER Import-ProviderSettings's own first
# call, which can itself log (e.g. "Settings load: ..."). Before this fix, referencing
# any of the three that early threw under Set-StrictMode ("variable cannot be retrieved
# because it has not been set"), silently swallowed by Write-DebugLog's own catch-all --
# so the very first log line of every process start was silently dropped, no error
# anywhere. Latent since Write-DebugLog/Invoke-LogRotation were written; never triggered
# before because nothing used to log unconditionally that early in the script.
$script:LogBytesWrittenSinceStart = 0
$script:LogRotateMaxBytes = 100 * 1MB
$script:LogRotateMaxGenerations = 3
# Same class of bootstrap-ordering issue as the three above -- Invoke-ProxmoxRestMethod
# reads this on every call, but real HTTP calls only start after provider/connect, well
# after Set-ProviderRuntimeFromSettings has run, so this one was never actually reachable
# uninitialized in practice. Given a bootstrap default here regardless, for the same
# defense-in-depth reason and matching Get-DefaultProviderSettings's own default.
$script:HttpTimeoutSeconds = 3
# Same defense-in-depth bootstrap as above -- see POOL-SCOPING.md.
$script:PoolScopeName = ''
$script:PoolScopeInheritOnClone = $true
# Same defense-in-depth bootstrap as above -- see MAC-PRESERVATION.md.
$script:PreserveMacOnRecreation = $false
$script:ProxmoxSession = $null
# Reused across REST calls so each one does not open a fresh TCP connection plus TLS
# handshake. Populated on the first successful request after provider/connect, and reset
# on provider/connect and provider/disconnect.
$script:ProxmoxWebSession = $null
$script:TaskContext = @{}
$script:DummyOperationTaskId = '__DUMMY_TASK__'

# Linked-clone template bookkeeping: {vmid: snapshotName} for a guest whose
# guests/snapshots/create RAS believes it issued. This is NOT the source of truth for
# exists/delete -- that's the guest's own live 'template' flag, see
# Handle-GuestSnapshotsExists -- it only makes the log readable.
$script:TemplateSnapshotIntent = @{}

# Mirror of the vmids of every 'clone' entry in the persisted clone-state file, kept in
# sync by Set-/Remove-CloneStateEntry so Invoke-TrackedCloneSweep can ask "is anything
# in flight at all" in O(1) with no file read. Repopulated from the persisted file near
# the bottom of this script, after all functions are defined.
$script:TrackedCloneVmIds = [System.Collections.Generic.HashSet[string]]::new()

# VmIds whose rasClone<sourceId> tag has already been confirmed correct (and any
# inherited rasTemplate* tag stripped) by Start-ProxmoxVmIfNeeded this process lifetime --
# see Repair-ProxmoxCloneTagsAfterInherit. Both guests/get's clone-aware flow and
# tasks/get's Handle-TaskInfo call Start-ProxmoxVmIfNeeded on every poll while a clone is
# still powered off, and each call otherwise did a live 'config' GET to re-check tags that
# do not change once correct. Cleared alongside clone tracking itself (Remove-CloneStateEntry)
# so a VMID Proxmox later reuses for an unrelated clone always re-verifies fresh.
$script:CloneTagVerified = [System.Collections.Generic.HashSet[string]]::new()
# Same "once per VmId this process lifetime, cleared alongside clone tracking" pattern
# as $script:CloneTagVerified immediately above, for cloning.mac_preservation instead of
# tag repair -- see MAC-PRESERVATION.md and Start-ProxmoxVmIfNeeded.
$script:CloneMacRestored = [System.Collections.Generic.HashSet[string]]::new()

# Pipelined cloning. RAS does not submit the next guests/clone until tasks/get reports
# the previous clone task 'completed', so reporting 'completed' only after the full
# clone+start+IP chain serializes provisioning end to end. Readiness is reported
# independently and repeatedly by guests/get's own clone-aware flow
# (Get-RasGuestObjectForCloneAwareFlow), so tasks/get's 'completed' only has to unblock
# RAS's next submission -- it does not also have to mean "the guest is ready". See
# PIPELINED-CLONING.md.
#
# Every value assigned below is a hard-coded fallback default. The real values come from
# RAS-CPF-Proxmox-Settings.json via Set-ProviderRuntimeFromSettings; at runtime read the
# corresponding $script:* variable, never these assignments.
# Set-StrictMode makes a bare .PSObject.Properties.Name THROW when the object has no
# properties at all -- and an empty JSON object ({}) parses to exactly that. Two real
# cases hit it: a request like {"method":"guests/get","params":{}}, and a settings file
# an admin has emptied to {}. The second is fatal, because Import-ProviderSettings runs
# at startup outside the main loop's try/catch, so the provider dies before it can
# serve or log anything -- defeating the whole "a bad config must never prevent the
# provider from starting" contract. Every member read on parsed JSON goes through here.
function Get-MemberNames {
    param([object]$Object)

    if ($null -eq $Object) { return @() }
    # [System.Collections.IDictionary] (not just [hashtable]) so this also covers
    # [ordered]@{} (System.Collections.Specialized.OrderedDictionary) -- a genuinely
    # different .NET type from Hashtable that a plain -is [hashtable] check misses
    # entirely -- surfaced by MAC-PRESERVATION.md's preserved_macs field,
    # which is an [ordered]@{} while still in $script:TaskContext (in-memory) but a
    # PSCustomObject once round-tripped through Set-CloneStateEntry/Get-CloneStateEntry
    # (JSON) -- this function needs to handle both shapes correctly either way.
    if ($Object -is [System.Collections.IDictionary]) { return @($Object.Keys) }
    try { return @($Object.PSObject.Properties | ForEach-Object { $_.Name }) }
    catch { return @() }
}

# The settings FILE's shape, not the provider's own version. Originally reserved only
# for a change that moves/renames/regroups an existing key (a plain new key at an
# existing path never needed a bump, since it always just fell back to its default on an
# older file) -- broadened to also cover a version that adds new keys, now
# that Import-ProviderSettings actually acts on the difference (see
# Update-SettingsFileSchema below) instead of the mismatch being purely informational.
# See SETTINGS.md's migration notes for what changed at each version -- there is no
# in-script changelog table to keep in sync by hand; Update-SettingsFileSchema derives
# what's new by diffing the old file's keys against the current defaults, so nothing here
# grows with each new version.
$script:CurrentSettingsSchemaVersion = 3

function Get-DefaultProviderSettings {
    return [ordered]@{
        _comment         = "Every leaf setting is {value, default, description} so an admin can see current vs. default and what a key does without opening SETTINGS.md. 'value' is the only field the provider actually reads; 'default'/'description' are informational, written once when the file is (re)seeded and otherwise ignored by the parser."
        # Bumped whenever this file's shape changes -- a key moves/renames/regroups, or
        # new keys are added -- see $script:CurrentSettingsSchemaVersion near
        # Import-ProviderSettings, and Update-SettingsFileSchema, which actually rewrites
        # an older file up to this.
        schema_version   = $script:CurrentSettingsSchemaVersion
        locations        = [ordered]@{
            data_directory                    = [ordered]@{ value = $PSScriptRoot; default = $PSScriptRoot; description = "Where the log, clone-state file, and this settings file itself live. Defaults to the script's own folder." }
            log_file                          = [ordered]@{ value = 'Proxmox-RAS-Provider.log'; default = 'Proxmox-RAS-Provider.log'; description = 'Log filename, relative to data_directory unless given as an absolute path.' }
            clone_state_file                  = [ordered]@{ value = 'Proxmox-RAS-CloneState.json'; default = 'Proxmox-RAS-CloneState.json'; description = 'Persisted in-flight-clone tracking, relative to data_directory unless absolute. Survives a provider restart.' }
            settings_reload_interval_seconds = [ordered]@{ value = 30; default = 30; description = "How often this file's own last-write time is checked and reloaded if changed. Checked opportunistically on incoming requests -- there is no background timer." }
        }
        capabilities     = [ordered]@{
            _comment              = "Mirrors RAS CPF's provider/initialize response field-for-field -- what this file's admin sets here is exactly what RAS is told the provider supports."
            can_suspend_guests    = [ordered]@{ value = $true; default = $true; description = 'Advertises guests/control(suspend) support to RAS.' }
            guests_polling_rate   = [ordered]@{ value = 30; default = 30; description = "Seconds between RAS's own guests/get polls per guest, as advertised to RAS (RAS decides whether to honor it)." }
            tasks_polling_rate    = [ordered]@{ value = 11; default = 11; description = "Seconds between RAS's tasks/get polls for an in-flight task. Must stay above cloning.pipelined_cloning.completion_seconds, or a clone that meets that threshold still waits out RAS's next scheduled poll." }
            tasks_polling_retries = [ordered]@{ value = 180; default = 180; description = 'How many tasks/get polls RAS makes before giving up on a task, as advertised to RAS.' }
            template_method       = [ordered]@{ value = 'basic'; default = 'basic'; description = 'CPF template-creation contract level: "none", "basic", or "versioning". "basic" is required for linked clones (can_link_clones).' }
            # This is THE single source of truth for linked-clone support -- both
            # Handle-Initialize's advertised capability and Handle-GuestClone's actual
            # gate read this exact same value, never two settings that could drift apart.
            # That single-source property is load-bearing, not cosmetic: a settings
            # file having this independently true while the underlying feature could not
            # actually serve it is exactly how a client ends up calling
            # 'guests/snapshots/create' only to get '-32601 Method not found' back (see
            # LINKED-CLONES.md #5) -- the failure mode of the earlier, schema-1 shape
            # where this key was capabilities-only and unconnected to any real
            # implementation. It is safe to keep it here now specifically because it IS
            # the real gate, not a separate advertisement of one.
            can_link_clones       = [ordered]@{ value = $false; default = $false; description = 'Master switch for linked clones. Off by default: confirm the storage backend supports it first (see LINKED-CLONES.md #3) before enabling on a new cluster.' }
        }
        cloning          = [ordered]@{
            # What Handle-GuestClone does when RAS asks for a linked clone (a non-empty
            # 'snapshot' param) but the live source is not a Proxmox template. Proxmox
            # itself would silently full-clone in this case -- no error, no warning --
            # so the two legal values here are both explicit substitutes for that
            # silence: 'full' logs loudly and proceeds as a full clone; 'error' refuses
            # instead. Not a capabilities.* key: it's clone-behavior policy, not
            # something RAS's provider/initialize response has a field for.
            linked_clone_fallback = [ordered]@{ value = 'full'; default = 'full'; description = 'What happens when RAS asks for a linked clone but the live source is not a Proxmox template. "full": log and proceed as a full clone. "error": refuse instead.' }

            # Lets RAS submit the next guests/clone before a full clone's disk copy
            # actually finishes -- see PIPELINED-CLONING.md. Has no effect on linked
            # clones: their real completion is already sub-second, faster than this
            # gate could ever fire first (Handle-TaskInfo skips the gate for them
            # outright).
            pipelined_cloning = [ordered]@{
                _comment                         = "Lets RAS submit the next guests/clone before a full clone's disk copy actually finishes. See PIPELINED-CLONING.md. Has no effect on linked clones (their real completion is already sub-second)."
                # $false falls back to reporting a clone task 'completed' only once
                # Proxmox's own job has genuinely finished (full clones only).
                enabled                          = [ordered]@{ value = $true; default = $true; description = "Master switch. false = a clone task only reports 'completed' once Proxmox's own job finished (full clones only -- linked clones are unaffected either way)." }
                # Once at least this many seconds have elapsed since a tracked clone was
                # submitted, Handle-TaskInfo reports it 'completed' immediately instead
                # of waiting for Proxmox's real job. A Proxmox-reported task FAILURE is
                # never masked by this. Keep at or below (capabilities.tasks_polling_rate
                # - 1) -- Set-ProviderRuntimeFromSettings warns if that is violated.
                completion_seconds               = [ordered]@{ value = 10; default = 10; description = 'Seconds after clone submission before a still-running full clone is reported completed anyway. Keep at or below (capabilities.tasks_polling_rate - 1).' }
                # Even past the elapsed-time threshold above, a clone task is only
                # reported 'completed' if fewer than this many clones are currently
                # active -- see Get-ActiveCloneCount. Reporting 'completed' is what
                # triggers RAS to submit the next real guests/clone, so this is the only
                # real concurrency gate on Proxmox's disk-I/O-heavy clone jobs.
                max_concurrent_clone_operations = [ordered]@{ value = 2; default = 2; description = 'Even past completion_seconds, a clone only pipelines through if fewer than this many full clones are currently active -- the real concurrency gate on disk-I/O-heavy clone jobs.' }
                # How often Handle-GuestControl's 'start' branch polls the REAL
                # underlying Proxmox clone task while waiting for it to genuinely finish,
                # before ever calling Proxmox's own start endpoint (which would otherwise
                # hit Proxmox's clone lock and fail). See
                # Wait-ForRealCloneCompletionBeforeStart.
                start_poll_interval_seconds      = [ordered]@{ value = 5; default = 5; description = "How often guests/control(start)'s defensive guard polls the real Proxmox clone task while waiting for it to finish." }
                # Ceiling for the same wait -- exists only to fail loudly rather than
                # hang the shared stdin/stdout pipe on a stuck clone.
                start_max_wait_seconds           = [ordered]@{ value = 1800; default = 1800; description = 'Ceiling for the same wait -- fails loudly rather than hanging the shared stdin/stdout pipe on a stuck clone.' }
            }

            # Deletion, caching, retry, and sweep timing -- everything that bounds a
            # wait or a retention window but isn't specific to pipelining or load
            # balancing.
            timeouts = [ordered]@{
                _comment                                       = "Deletion, caching, retry, and sweep timing -- everything that bounds a wait or a retention window but isn't specific to pipelining or load balancing."
                # Bounds for the forced hard stop guests/control(delete) issues before
                # destroying a VM -- see Stop-ProxmoxVmHardBeforeDelete. Exist only to
                # fail loudly rather than hang if a hard stop somehow never completes.
                delete_hard_stop_poll_interval_seconds        = [ordered]@{ value = 2; default = 2; description = 'Poll interval for the forced hard stop guests/control(delete) issues before destroying a VM.' }
                delete_hard_stop_max_wait_seconds             = [ordered]@{ value = 60; default = 60; description = 'Ceiling for the same wait -- fails loudly rather than hanging if a hard stop never completes.' }
                # GET /api2/json/cluster/resources?type=vm returns node/name/status/
                # template for every VM in one shot -- cached this long instead of
                # re-fetched on every single-guest lookup. Keep in sync with
                # capabilities.guests_polling_rate.
                cluster_resources_cache_ttl_seconds           = [ordered]@{ value = 30; default = 30; description = 'How long the bulk VM listing (GET /cluster/resources?type=vm) is reused before re-fetching. Keep in sync with capabilities.guests_polling_rate.' }
                # A destroyed VM is omitted from guests/list for this long even if
                # Proxmox's own cluster/resources view has not caught up -- see
                # Test-ProxmoxRecentlyDeleted. Self-expiring, so a reused VMID is never
                # permanently hidden. Also the window
                # cloning.load_balancing.preserve_node_on_recreation checks.
                recently_deleted_retention_seconds            = [ordered]@{ value = 300; default = 300; description = "How long a VM this provider deleted stays hidden from guests/list even if Proxmox's own view hasn't caught up. Also the window load_balancing.preserve_node_on_recreation checks." }
                # For this long after this provider issues a control action on a VM,
                # power state is resolved from qemu/{id}/status/current (authoritative)
                # instead of the lagging cluster/resources aggregate -- see
                # Test-ProxmoxRecentlyControlled.
                recently_controlled_retention_seconds         = [ordered]@{ value = 120; default = 120; description = 'After this provider issues a control action, how long power state is read live instead of from the lagging cluster listing.' }
                # For this long after THIS provider issues a 'stop' on a VM, the
                # guest-agent network probe is skipped outright instead of attempted --
                # see Test-ProxmoxRecentlyStopped and Get-ProxmoxVmNetworkData. Short on
                # purpose: a 'stop' really is near-instant, this only needs to cover the
                # race between our own stop and Proxmox's teardown.
                recently_stopped_retention_seconds            = [ordered]@{ value = 10; default = 10; description = "After this provider's own 'stop', how long the guest-agent probe is skipped outright instead of racing Proxmox's teardown." }
                # How long Resolve-PendingTemplateTagRemoval waits after a
                # guests/convert {is_template:false} before concluding no
                # guests/control start is coming for that vmid -- i.e. that RAS deleted
                # the Template object rather than entering maintenance -- and removes
                # the rasTemplate<id> tag. A 'start' for the same vmid within this
                # window cancels the removal outright.
                template_delete_confirm_seconds               = [ordered]@{ value = 10; default = 10; description = 'After guests/convert(is_template:false), how long to wait for a guests/control(start) before concluding RAS deleted the Template (not entering maintenance) and removing the rasTemplate<id> tag.' }
                # Same idea for the template flag after a guests/convert -- see
                # Test-ProxmoxRecentlyConverted.
                recently_converted_retention_seconds          = [ordered]@{ value = 120; default = 120; description = 'Same idea as recently_controlled, for the template flag after a guests/convert.' }
                # Invoke-TrackedCloneSweep (the opportunistic re-check of every OTHER
                # in-flight clone on each guests/get) is throttled to at most once per
                # tracked id per this many seconds.
                sweep_min_interval_seconds                    = [ordered]@{ value = 5; default = 5; description = 'The opportunistic re-check of every OTHER in-flight clone on each guests/get is throttled to at most once per tracked id per this many seconds.' }
                # How long a completed clone task's clone_id is remembered after
                # tracking is cleared, so a tasks/get that races guests/get's own
                # completion path can still answer. See $script:CompletedCloneTaskOutputs
                # and Handle-TaskInfo.
                completed_clone_task_output_retention_seconds = [ordered]@{ value = 600; default = 600; description = "How long a completed clone task's clone_id is remembered after tracking is cleared, so a racing tasks/get can still answer correctly." }
                # guests/control (start/stop/reset/reboot/delete/suspend/resume) retries
                # this many times total on a transient Proxmox 5xx (e.g. "got no worker
                # upid - start worker failed" under burst load), waiting this many
                # seconds between attempts. Set max_attempts to 1 to disable retrying
                # entirely.
                control_retry_max_attempts                    = [ordered]@{ value = 3; default = 3; description = 'guests/control retries this many times total on a transient Proxmox 5xx. Set to 1 to disable retrying.' }
                control_retry_delay_seconds                   = [ordered]@{ value = 1; default = 1; description = 'Delay between control-action retry attempts.' }
                # The default ceiling for every Proxmox HTTP call that doesn't set its own
                # (guest_agent.timeout_seconds and tag_write_timeout_seconds below are
                # deliberately separate, tighter knobs). Was unbounded (PowerShell's own
                # HttpClient default) -- a hung hypervisor connection could otherwise stall
                # the shared stdin/stdout pipe indefinitely. Every such call also gets
                # exactly one retry on a fresh connection (never the same possibly-broken
                # pooled one) before the failure reaches the caller -- see LOGGING.md's
                # HTTP component notes.
                http_timeout_seconds                          = [ordered]@{ value = 3; default = 3; description = 'Default timeout (seconds) for a Proxmox API call that does not set its own tighter timeout. One retry on a fresh connection happens automatically before a failure is returned.' }
                # Best-effort metadata writes (the rasTemplate/rasClone tag PUTs) must
                # never dominate the request they ride along on. 3s is generous for a
                # healthy config PUT and still well below Proxmox's own ~10s
                # lock-wait/timeout.
                tag_write_timeout_seconds                     = [ordered]@{ value = 3; default = 3; description = 'Timeout on every rasTemplate/rasClone tag PUT -- best-effort metadata that must never dominate the request it rides along on.' }
                # Hard ceiling on how long a clone stays tracked without ever reaching
                # ready (powered_on + a real IP). Exists specifically for the case
                # nothing else catches: a VM deleted outside this provider's own
                # guests/control(delete) path (e.g. by hand in the Proxmox UI after a
                # failed/overloaded clone), which otherwise leaves Get-ActiveCloneCount
                # counting it as active forever (it only ever clears on a CONFIRMED
                # completed real task, never on failed) and
                # Get-RasGuestObjectForCloneAwareFlow reporting powering_on forever. Past
                # this age the entry is retired unconditionally and logged, regardless of
                # whether the VM is resolvable or just never became ready. Default (30
                # min) comfortably exceeds a typical full-clone's clone+boot time; raise
                # it if your own environment's clones routinely take longer.
                clone_tracking_max_age_seconds                = [ordered]@{ value = 1800; default = 1800; description = 'Hard ceiling (seconds) a clone stays tracked without reaching ready before its tracking is force-retired regardless of state -- the backstop for a clone deleted outside guests/control(delete), or one that never becomes ready for any other reason.' }
            }

            # Spreading clone compute across cluster nodes instead of always landing on
            # the source template's own node -- see DISTRIBUTED-PLACEMENT.md. Requires
            # shared storage (e.g. Ceph/RBD): 'target' only ever moves compute, never
            # disk placement.
            load_balancing = [ordered]@{
                _comment                     = 'Spreading clone compute across cluster nodes instead of always landing on the source template''s own node. See DISTRIBUTED-PLACEMENT.md. Requires shared storage (e.g. Ceph/RBD) -- target only moves compute, never disk.'
                # Master switch. Off by default: this cluster's storage is shared
                # (Ceph/RBD, confirmed -- see LINKED-CLONES.md #6), so cross-node
                # placement is safe here, but it is still a live scheduling behavior
                # change on real infrastructure -- an admin turns it on deliberately,
                # the same reasoning as capabilities.can_link_clones above. When
                # $false, Handle-GuestClone never sets 'target' and every clone lands
                # on the source's node, exactly like before this feature existed.
                enabled                      = [ordered]@{ value = $false; default = $false; description = "Master switch. false = every clone lands on the source's own node, exactly as before this feature existed. A live scheduling change on real infrastructure -- opt in deliberately." }
                # 'resource' (default) picks the least-loaded eligible node per
                # resource_metric below, read fresh from GET /cluster/resources?type=node
                # (cached briefly -- node_stats_cache_ttl_seconds). 'round_robin'
                # ignores load entirely and cycles through eligible nodes in a fixed
                # order.
                strategy                     = [ordered]@{ value = 'resource'; default = 'resource'; description = '"resource": pick the least-loaded eligible node per resource_metric. "round_robin": ignore load, cycle eligible nodes in name order.' }
                # Only used when strategy = 'resource'. 'cpu' ranks nodes by the 'cpu'
                # field (Proxmox's own normalized load fraction, comparable across
                # nodes regardless of core count); 'ram' ranks by mem/maxmem (fraction
                # used); 'both' averages the two fractions. Lower is better in every
                # case -- the node picked is the one with the most headroom.
                resource_metric              = [ordered]@{ value = 'both'; default = 'both'; description = 'Only used when strategy = "resource". "cpu", "ram", or "both" (average of the two load fractions). Lower is better in every case.' }
                # Node names to never place a clone's compute on, even if otherwise
                # eligible (online, not in maintenance) -- e.g. a node that also hosts
                # the RAS Connection Broker/this provider itself and should not carry
                # VDI workload. Empty by default. Case-sensitive, must match Proxmox's
                # own node names exactly.
                excluded_nodes               = [ordered]@{ value = @(); default = @(); description = "Node names never chosen as a clone target even if online, e.g. a node that also runs the RAS Connection Broker/this provider. Case-sensitive, must match Proxmox's node names exactly." }
                # How long a fetched node-resource snapshot is reused before
                # Handle-GuestClone asks Proxmox again. Short on purpose -- this feeds a
                # real scheduling decision, not a display value. Separate cache from
                # cloning.timeouts.cluster_resources_cache_ttl_seconds (different
                # endpoint -- ?type=node vs ?type=vm -- and a different freshness need).
                node_stats_cache_ttl_seconds = [ordered]@{ value = 15; default = 15; description = 'How long a fetched node-resource snapshot (GET /cluster/resources?type=node) is reused. Separate from timeouts.cluster_resources_cache_ttl_seconds.' }
                # When RAS recreates a guest (delete, then clone a new VM under the SAME
                # name), land it back on the node it was already running on instead of
                # re-running strategy selection. Only ever applies to a name this
                # provider itself deleted within
                # cloning.timeouts.recently_deleted_retention_seconds -- see
                # $script:RecentlyDeletedNodeByName and
                # Get-ProxmoxPreservedNodeForRecreation.
                preserve_node_on_recreation  = [ordered]@{ value = $true; default = $true; description = 'When RAS deletes a guest and clones a new VM under the same name, land it back on the node it was already on instead of re-running strategy selection.' }
            }
            # Same "recreation" concept and the same recently_deleted_retention_seconds
            # window as preserve_node_on_recreation above, but for network identity
            # rather than placement: when RAS deletes a guest and clones a new VM under
            # the SAME name, give the new VM's NIC(s) the SAME MAC address(es) the old
            # one had, instead of Proxmox's default fresh-random-MAC-per-clone. Off by
            # default -- unlike node preservation, this is new/unvalidated against a
            # live cluster and carries a real (if narrow) risk: Proxmox does not enforce
            # cluster-wide MAC uniqueness the way it does for VMIDs, so a delete that
            # hasn't genuinely finished tearing down before the recreate's MAC write
            # lands could put two live NICs on the same MAC. See POOL-SCOPING.md-style
            # sibling doc MAC-PRESERVATION.md for the full design and that risk's
            # actual bound (delete already polls to confirm before RAS ever sees it
            # complete, same discipline VMID reuse already relies on).
            mac_preservation = [ordered]@{
                enabled = [ordered]@{ value = $false; default = $false; description = 'When RAS deletes a guest and clones a new VM under the same name, give its NIC(s) the same MAC address(es) the old one had -- keeps a DHCP reservation or MAC-keyed licensing continuous across a recreate. Off by default.' }
            }
        }
        virtual_machines = [ordered]@{
            guest_agent      = [ordered]@{
                # Short, explicit timeout on the guest-agent call: Invoke-ProxmoxRestMethod
                # otherwise inherits Proxmox's own ~3s block before the 500 when an agent
                # is configured but not running. A healthy agent answers in tens of ms.
                timeout_seconds          = [ordered]@{ value = 1; default = 1; description = "Explicit timeout on the guest-agent network-interfaces call. A healthy agent answers in tens of ms; Proxmox's own default would otherwise block ~3s per agent-less VM." }
                # Once a VM has been confirmed agent-less for at least this long (giving
                # a freshly booted/cloned guest real time for its agent to start
                # reporting) AND at least this many separate poll failures, this script
                # tags the VM rasQuarantine and stops calling the agent endpoint for it.
                # The tag is the whole state, so the skip survives a provider restart;
                # an admin removes it by hand.
                quarantine_after_seconds = [ordered]@{ value = 60; default = 60; description = 'How long a VM must be confirmed agent-less before it is tagged rasQuarantine and the agent stops being probed (paired with quarantine_min_failures).' }
                quarantine_min_failures  = [ordered]@{ value = 2; default = 2; description = 'How many separate poll failures, alongside quarantine_after_seconds, before quarantine tagging kicks in.' }
            }
            orphan_detection = [ordered]@{
                # RAS never issues stop or delete again for a clone it has permanently
                # lost track of. This provider cannot fix that RAS-side gap, but it can
                # flag it: periodically (see check_interval_seconds), any VM tagged
                # rasClone<sourceId> that this provider is no longer tracking as an
                # in-flight clone, AND that RAS has not polled via guests/get in more
                # than stale_poll_after_seconds, is logged as an ORPHAN CANDIDATE and
                # tagged rasOrphanCandidate, best-effort. Log-and-tag only -- this
                # never deletes or stops anything.
                enabled                  = [ordered]@{ value = $true; default = $true; description = 'Periodically flags (log + tag only, never stops or deletes) a rasClone-tagged VM this provider is no longer tracking AND RAS has not polled recently.' }
                check_interval_seconds   = [ordered]@{ value = 300; default = 300; description = 'How often the audit runs, throttled opportunistically like everything else in this provider (no background timer).' }
                stale_poll_after_seconds = [ordered]@{ value = 180; default = 180; description = "How long since RAS's last guests/get for a VM before it qualifies as a candidate." }
            }
            tags             = [ordered]@{
                _comment                 = "Proxmox tag names this provider reads/writes -- a single ';'-separated string on the VM's own config, visible and editable by an admin in the Proxmox UI too."
                # Proxmox tag names this provider reads/writes. Tags live on the VM's
                # own PVE config (a single ';'-separated string), so all of these are
                # visible and directly editable by an admin in the Proxmox UI.
                #   ras_exclude_tag          (manual, admin-set) -- this VM does not
                #     exist for RAS at all: never listed, and Get-ProxmoxVmNode reports
                #     it "not found".
                #   ras_quarantine_tag       (automatic) -- see guest_agent above.
                #   ras_template_tag_prefix  (automatic) -- <prefix><id> is applied to
                #     a clone's SOURCE, where <id> is the source's own VMID.
                #   ras_clone_tag_prefix     (automatic) -- <prefix><id> is applied to
                #     a clone once its Proxmox clone job is confirmed done, where <id>
                #     is the SOURCE.
                #   ras_orphan_candidate_tag (automatic) -- see orphan_detection above.
                ras_exclude_tag          = [ordered]@{ value = 'rasExclude'; default = 'rasExclude'; description = 'Manual, admin-set -- this VM does not exist for RAS at all (never listed, guests/get reports not-found).' }
                ras_quarantine_tag       = [ordered]@{ value = 'rasQuarantine'; default = 'rasQuarantine'; description = 'Automatic -- see guest_agent above.' }
                ras_template_tag_prefix  = [ordered]@{ value = 'rasTemplate'; default = 'rasTemplate'; description = "Automatic -- <prefix><id> applied to a clone's SOURCE, where <id> is the source's own VMID." }
                ras_clone_tag_prefix     = [ordered]@{ value = 'rasClone'; default = 'rasClone'; description = "Automatic -- <prefix><id> applied to a fresh clone once its Proxmox job is confirmed done, where <id> is the SOURCE's VMID." }
                ras_orphan_candidate_tag = [ordered]@{ value = 'rasOrphanCandidate'; default = 'rasOrphanCandidate'; description = 'Automatic -- see orphan_detection above.' }
            }
            # Restricts the fleet RAS sees to one Proxmox pool -- an out-of-scope VM is
            # excluded before any handler ever sees it (Get-ProxmoxClusterVMs), so it
            # is "not found" everywhere, the same outcome as tags.ras_exclude_tag above,
            # just scoped by pool membership instead of an individual tag. See
            # POOL-SCOPING.md.
            # Note: relies on cluster/resources's per-VM 'pool' field. This is Proxmox's
            # documented, standard field, but has not been confirmed against every
            # Proxmox VE version -- see POOL-SCOPING.md for the fallback behavior if it
            # is ever absent on a pooled VM.
            pool_scope       = [ordered]@{
                _comment         = 'Pool-level visibility scoping and clone pool inheritance -- see POOL-SCOPING.md.'
                pool_name        = [ordered]@{ value = ''; default = ''; description = 'Proxmox pool id (poolid) this provider is scoped to. Empty (default) = no pool filtering, every pool and every unpooled VM is visible. When set, a VM not in this exact pool is invisible to RAS -- never listed, guests/get reports not-found -- identically to ras_exclude_tag. Case-sensitive, must match the poolid exactly.' }
                inherit_on_clone = [ordered]@{ value = $true; default = $true; description = "Whether a clone is created in its source template's own Proxmox pool (via the clone API's own 'pool' parameter), so pool-scoped visibility and any pool-based Proxmox permissions extend to clones automatically. Independent of pool_name -- applies whenever the source VM actually belongs to a pool, even if pool_name filtering itself is off." }
            }
        }
        logging          = [ordered]@{
            # 5 = Verbose (every RPC's full JSON, every HTTP call, every per-poll state
            # transition). 4 = Extended (drops the per-poll/per-HTTP-call noise, keeps
            # every lifecycle event, control action, tag decision, and error/warning). 3
            # = Standard (lifecycle milestones and errors/warnings only). See
            # SETTINGS.md and Write-DebugLog.
            log_level                  = [ordered]@{ value = 5; default = 5; description = '3 = Standard (lifecycle + errors only), 4 = Extended (+ every lifecycle/control/tag event), 5 = Verbose (+ every RPC and HTTP call). Accepts the number or the word.' }
            # Stored in MB for readability; converted to bytes once in
            # Set-ProviderRuntimeFromSettings ($script:LogRotateMaxBytes is still bytes
            # internally -- Invoke-LogRotation compares it against a running byte count).
            log_rotate_max_mb          = [ordered]@{ value = 100; default = 100; description = 'Log file size (MB) that triggers rotation.' }
            log_rotate_max_generations = [ordered]@{ value = 3; default = 3; description = 'How many rotated generations are kept.' }
        }
    }
}

# Every leaf in this settings file is {value, default, description} (see
# Get-DefaultProviderSettings) -- 'value' is the only field the provider actually reads,
# 'default'/'description' exist purely so an admin can see current-vs-default and what a
# key does without opening SETTINGS.md. This unwraps that shape once, tolerating a bare
# scalar/array too (a hand-simplified file, an older/pre-schema file, or a value some
# caller already unwrapped) -- every ConvertTo-CoercedSetting* helper below runs its raw
# input through this first, so none of them need their own awareness of the wrapper.
function Get-SettingLeafRawValue {
    param($Value)
    if ($null -eq $Value) { return $null }
    if ((Get-MemberNames -Object $Value) -contains 'value') { return $Value.value }
    return $Value
}

# Flattens a settings tree into "a.b.c" -> leaf-value pairs, for diffing -- used both to
# report what a freshly loaded settings file changed from this provider's hardcoded
# defaults, and to report exactly what changed on a hot reload. Works on either shape
# this script builds internally: Get-DefaultProviderSettings's tree (every leaf a
# {value,default,description} [ordered]@{} -- NOT a PSCustomObject, so this deliberately
# does not reuse Get-MemberNames/Get-SettingLeafRawValue above, which are for parsed-JSON
# PSCustomObjects only) and $script:Settings itself (already fully resolved to bare
# scalars/arrays at every leaf, no wrapper). A node counts as a "section" and gets
# recursed into unless it's an IDictionary carrying a 'value' key, which is what actually
# distinguishes a settings leaf from a grouping section in the defaults tree; in
# $script:Settings's tree no section ever has a 'value' key, so every leaf there is
# simply wherever recursion bottoms out at a non-dictionary. Arrays are joined to a
# single comparable/printable string.
function ConvertTo-FlatSettingsMap {
    param([object]$Node, [string]$Prefix = '')

    $result = [ordered]@{}
    $isSection = $Node -is [System.Collections.IDictionary] -and -not $Node.Contains('value')
    if ($isSection) {
        foreach ($key in $Node.Keys) {
            $path = if ($Prefix) { "$Prefix.$key" } else { [string]$key }
            $sub = ConvertTo-FlatSettingsMap -Node $Node[$key] -Prefix $path
            foreach ($k in $sub.Keys) { $result[$k] = $sub[$k] }
        }
        return $result
    }

    $leaf = if ($Node -is [System.Collections.IDictionary] -and $Node.Contains('value')) { $Node['value'] } else { $Node }
    if ($leaf -is [System.Array]) { $leaf = ($leaf -join ',') }
    $result[$Prefix] = $leaf
    return $result
}

# Same idea as ConvertTo-FlatSettingsMap, but for the OTHER shape a settings tree shows up
# in: a raw ConvertFrom-Json result (nested PSCustomObjects), which is not an IDictionary
# and so ConvertTo-FlatSettingsMap's own section/leaf test never matches it. Returns just
# the "a.b.c" leaf PATHS actually present in $Node -- not their values, since the only
# thing Update-SettingsFileSchema needs this for is "did this key exist in the OLD file at
# all", to tell a genuinely new key (added by a later schema version) apart from one that
# was already there. Tolerates both shapes a settings file's leaves can be written in: a
# bare scalar/array (a hand-simplified file), or the rich {value,default,description}
# wrapper this script itself writes -- a node counts as a leaf the moment it is not an
# object/dictionary at all, OR it is one but carries a 'value' member.
function Get-SettingsLeafPaths {
    param([object]$Node, [string]$Prefix = '')

    $paths = [System.Collections.Generic.List[string]]::new()
    $isContainer = ($Node -is [System.Collections.IDictionary]) -or ($Node -is [System.Management.Automation.PSCustomObject])
    if (-not $isContainer) {
        if ($Prefix) { $paths.Add($Prefix) }
        return $paths
    }

    $memberNames = Get-MemberNames -Object $Node
    if ($memberNames -contains 'value') {
        if ($Prefix) { $paths.Add($Prefix) }
        return $paths
    }

    foreach ($name in $memberNames) {
        if ($name -eq '_comment' -or $name -eq 'schema_version') { continue }
        $childValue = if ($Node -is [System.Collections.IDictionary]) { $Node[$name] } else { $Node.$name }
        $path = if ($Prefix) { "$Prefix.$name" } else { [string]$name }
        $paths.AddRange([string[]](Get-SettingsLeafPaths -Node $childValue -Prefix $path))
    }
    return $paths
}

# Tolerant coercions for settings values pulled out of parsed JSON. A value that fails
# to coerce falls back to that key's own default -- a bad config must never prevent the
# provider from starting.
function ConvertTo-CoercedSettingBool {
    param($Value, [bool]$Default)
    $v = Get-SettingLeafRawValue -Value $Value
    if ($null -eq $v) { return $Default }
    if ($v -is [bool]) { return $v }
    try { return [System.Convert]::ToBoolean([string]$v) } catch { return $Default }
}

function ConvertTo-CoercedSettingInt {
    param($Value, [int]$Default)
    $v = Get-SettingLeafRawValue -Value $Value
    if ($null -eq $v) { return $Default }
    try { return [int]$v } catch { return $Default }
}

function ConvertTo-CoercedSettingString {
    param($Value, [string]$Default)
    $v = Get-SettingLeafRawValue -Value $Value
    if ($null -eq $v -or [string]::IsNullOrWhiteSpace([string]$v)) { return $Default }
    return [string]$v
}

# A JSON array of strings (e.g. load_balancing.excluded_nodes) parses as an object[] (or
# a single string if the admin wrote a bare value instead of an array -- tolerated here
# too). Anything else, or an unparseable element, falls back to the default rather than
# ever throwing under Set-StrictMode.
function ConvertTo-CoercedSettingStringArray {
    param($Value, [string[]]$Default)
    $v = Get-SettingLeafRawValue -Value $Value
    if ($null -eq $v) { return $Default }
    try { return @($v | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) }
    catch { return $Default }
}

# Accepts either a number (3/4/5) or the case-insensitive string alias
# ('standard'/'extended'/'verbose') documented in SETTINGS.md, and clamps
# anything else back to the default.
function ConvertTo-CoercedLogLevel {
    param($Value, [int]$Default)

    $v = Get-SettingLeafRawValue -Value $Value
    if ($null -eq $v) { return $Default }

    if ($v -is [int] -or $v -is [double] -or $v -is [long]) {
        $n = [int]$v
        if ($n -ge 3 -and $n -le 5) { return $n }
        return $Default
    }

    switch (([string]$v).Trim().ToLowerInvariant()) {
        'standard' { return 3 }
        'extended' { return 4 }
        'verbose' { return 5 }
        default {
            try {
                $n = [int]$v
                if ($n -ge 3 -and $n -le 5) { return $n }
            }
            catch { }
            return $Default
        }
    }
}

$script:SettingsPath = Join-Path -Path $PSScriptRoot -ChildPath 'RAS-CPF-Proxmox-Settings.json'
$script:Settings = $null
$script:SettingsFileLastWriteUtc = $null
$script:SettingsLastPolledAt = [DateTime]::MinValue

# Writes RAS-CPF-Proxmox-Settings.json populated with the current hard-coded defaults,
# atomically (temp file + rename). Only ever called when the file does not exist at all
# -- a file that exists but fails to parse is left completely untouched, so an
# in-progress hand-edit is never silently overwritten. Failure to write is itself
# non-fatal.
function Save-DefaultProviderSettingsFile {
    param([Parameter(Mandatory = $true)][object]$Defaults)

    try {
        $dir = Split-Path -Path $script:SettingsPath -Parent
        if (-not [string]::IsNullOrWhiteSpace($dir) -and -not (Test-Path -LiteralPath $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
        }

        $json = $Defaults | ConvertTo-Json -Depth 10
        $tempPath = "$($script:SettingsPath).tmp"
        Set-Content -LiteralPath $tempPath -Value $json -Encoding UTF8
        Move-Item -LiteralPath $tempPath -Destination $script:SettingsPath -Force
    }
    catch {
        try { Write-DebugLog "Failed to write default settings file [$($script:SettingsPath)]: $($_.Exception.Message)" -Level 'E' -Component '0A' } catch { }
    }
}

# Rebuilds the rich {value,default,description} file tree Get-DefaultProviderSettings
# itself produces, but with every leaf's 'value' overridden from $FlatCurrentValues (a
# flat "a.b.c" -> value map of $script:Settings, i.e. ConvertTo-FlatSettingsMap's own
# output -- see there) when present, falling back to the default's own value otherwise.
# Used by Update-SettingsFileSchema to upgrade an older file in place: every value an
# admin already customized survives untouched (including under a NEW key this exact
# version introduced, since $script:Settings already resolved that to its default before
# this ever runs), 'default'/'description' text is refreshed to what THIS script version
# actually says, and any key genuinely new to this schema version appears for the first
# time, at its default. Bare top-level scalars (_comment, schema_version) pass through
# from $DefaultsNode unchanged -- they are not per-setting leaves and are already correct
# for the CURRENT version there.
function New-MergedSettingsFileTree {
    param(
        [Parameter(Mandatory = $true)]
        [object]$DefaultsNode,
        # ConvertTo-FlatSettingsMap's own return type ([ordered]@{}, an
        # OrderedDictionary) -- deliberately typed as the interface, not [hashtable]:
        # OrderedDictionary is a different .NET type a [hashtable]-typed parameter
        # would reject outright. See Get-MemberNames's own comment on this exact trap.
        [Parameter(Mandatory = $true)]
        [System.Collections.IDictionary]$FlatCurrentValues,
        [string]$Prefix = ''
    )

    $isLeaf = $DefaultsNode -is [System.Collections.IDictionary] -and
              $DefaultsNode.Contains('value') -and
              $DefaultsNode.Contains('description')
    if ($isLeaf) {
        $resolvedValue = if ($FlatCurrentValues.Contains($Prefix)) { $FlatCurrentValues[$Prefix] } else { $DefaultsNode.value }
        return [ordered]@{ value = $resolvedValue; default = $DefaultsNode.default; description = $DefaultsNode.description }
    }

    if ($DefaultsNode -is [System.Collections.IDictionary]) {
        $result = [ordered]@{}
        foreach ($key in $DefaultsNode.Keys) {
            $value = $DefaultsNode[$key]
            if ($value -is [System.Collections.IDictionary]) {
                $path = if ($Prefix) { "$Prefix.$key" } else { [string]$key }
                $result[$key] = New-MergedSettingsFileTree -DefaultsNode $value -FlatCurrentValues $FlatCurrentValues -Prefix $path
            }
            else {
                $result[$key] = $value
            }
        }
        return $result
    }

    return $DefaultsNode
}

# Rewrites RAS-CPF-Proxmox-Settings.json in place, upgrading it from an older
# schema_version to $script:CurrentSettingsSchemaVersion -- called from
# Import-ProviderSettings only when a file that PARSED SUCCESSFULLY is genuinely older
# (a corrupt file is never touched, same as everywhere else in this script). Every value
# already on disk survives exactly as-is; new keys this version introduced appear at
# their default (since $script:Settings has already resolved everything by the time this
# runs); 'default'/'description' text refreshes to the current script's own copy. Written
# atomically (temp file + rename), same pattern as Save-DefaultProviderSettingsFile.
# Best-effort: a failed write here is logged and never blocks the provider from starting
# on the values it already loaded into memory -- the file staying on the old schema
# until the next successful attempt is not itself a problem, just a missed refresh.
#
# What actually changed is DERIVED, not hand-maintained: the leaf paths present in
# Get-DefaultProviderSettings but absent from $OldParsed (the OLD file's own on-disk
# shape, via Get-SettingsLeafPaths -- deliberately NOT $script:Settings, which by this
# point already has every key resolved to a default and so would never show a "new" key
# at all) are, by definition, exactly the keys this migration is adding. No separate
# changelog table to keep in sync by hand, and no such table to grow unbounded release
# after release. See SETTINGS.md's per-version migration notes for the human-readable
# "why" behind each key.
function Update-SettingsFileSchema {
    param(
        [Parameter(Mandatory = $true)]
        [int]$FromVersion,
        [Parameter(Mandatory = $true)]
        [int]$ToVersion,
        # The raw ConvertFrom-Json result of the file AS IT WAS ON DISK, before
        # $script:Settings filled in every unset key at its default -- see the function
        # comment above for why this, and not $script:Settings, is what "new key" has to
        # be diffed against.
        [Parameter(Mandatory = $true)]
        [object]$OldParsed
    )

    try {
        $defaults = Get-DefaultProviderSettings
        $flatDefaults = ConvertTo-FlatSettingsMap -Node $defaults
        $flatCurrent = ConvertTo-FlatSettingsMap -Node $script:Settings
        $oldLeafPaths = @(Get-SettingsLeafPaths -Node $OldParsed)

        $newKeys = @($flatDefaults.Keys | Where-Object { $oldLeafPaths -notcontains $_ })
        $changeSummary = if ($newKeys.Count -gt 0) {
            ($newKeys | ForEach-Object { "  + $_ (default: $($flatDefaults[$_]))" }) -join "`n"
        }
        else {
            '  (no new keys -- relocated/renamed only)'
        }

        $merged = New-MergedSettingsFileTree -DefaultsNode $defaults -FlatCurrentValues $flatCurrent

        $json = $merged | ConvertTo-Json -Depth 10
        $tempPath = "$($script:SettingsPath).tmp"
        Set-Content -LiteralPath $tempPath -Value $json -Encoding UTF8
        Move-Item -LiteralPath $tempPath -Destination $script:SettingsPath -Force

        Write-DebugLog "Settings schema migrated $FromVersion -> $ToVersion in [$($script:SettingsPath)]. Every existing value preserved; new key(s) added at their default:`n$changeSummary" -Level 'W' -Component '0A'
    }
    catch {
        Write-DebugLog "Settings schema migration $FromVersion -> $ToVersion failed -- [$($script:SettingsPath)] left on the old schema on disk, but this load already has every value resolved correctly in memory (new keys at their default): $($_.Exception.Message)" -Level 'E' -Component '0A'
    }
}

# Loads (or, on first run, creates) RAS-CPF-Proxmox-Settings.json. Every key is
# optional; a missing file, a missing key, or a value that fails to coerce falls back to
# that key's hard-coded default. See SETTINGS.md and Update-ProviderSettingsIfChanged
# for the hot-reload path (-IsReload).
function Import-ProviderSettings {
    param([switch]$IsReload)

    # Captured before $script:Settings is reassigned below, purely so the two logging
    # branches at the end of this function (first load vs. reload) can each report what
    # actually changed instead of silently applying it. $null on a genuine first load.
    $priorSettings = $script:Settings

    $defaults = Get-DefaultProviderSettings
    $parsed = $null
    $fileExists = Test-Path -LiteralPath $script:SettingsPath

    if ($fileExists) {
        try {
            $raw = Get-Content -LiteralPath $script:SettingsPath -Raw -Encoding UTF8
            if (-not [string]::IsNullOrWhiteSpace($raw)) {
                $parsed = $raw | ConvertFrom-Json -ErrorAction Stop
            }
        }
        catch {
            $verb = if ($IsReload) { 'Settings reload' } else { 'Settings load' }
            Write-DebugLog "${verb}: failed to parse [$($script:SettingsPath)]: $($_.Exception.Message) -- leaving the file untouched and $(if ($IsReload) { 'keeping the previously-loaded settings' } else { 'using in-memory defaults' })." -Level 'E' -Component '0A'
            if ($IsReload) { return }
            $parsed = $null
        }
    }
    elseif (-not $IsReload) {
        Save-DefaultProviderSettingsFile -Defaults $defaults
    }

    # Defaults to the CURRENT version: that's correct both when there's no file yet (we
    # just seeded one at the current schema, right above) and when an existing file
    # failed to parse (we fall back to pure in-memory defaults, which are current-schema
    # too) -- in neither case have we actually observed an old file to justify claiming
    # otherwise. Only a successfully parsed file that's missing the field is genuinely
    # schema 1 (the original flat shape, from before this field existed at all).
    # A genuine mismatch (older file, successfully parsed) is acted on further down,
    # once $script:Settings itself is built -- see Update-SettingsFileSchema.
    $schemaVersion = $script:CurrentSettingsSchemaVersion
    if ($null -ne $parsed) {
        $schemaVersion = if ((Get-MemberNames -Object $parsed) -contains 'schema_version') {
            ConvertTo-CoercedSettingInt -Value $parsed.schema_version -Default 1
        }
        else {
            1
        }
    }

    # Every leaf below starts from $defaults...value (the flat, unwrapped default) and is
    # then overwritten in place if the parsed file provides it -- ConvertTo-CoercedSetting*
    # itself unwraps a {value,default,description} leaf (or tolerates a bare scalar), so
    # $p.foo works whether the file uses the rich shape or a hand-simplified bare value.

    $loc = [ordered]@{
        data_directory                    = $defaults.locations.data_directory.value
        log_file                          = $defaults.locations.log_file.value
        clone_state_file                  = $defaults.locations.clone_state_file.value
        settings_reload_interval_seconds = $defaults.locations.settings_reload_interval_seconds.value
    }
    if ($null -ne $parsed -and (Get-MemberNames -Object $parsed) -contains 'locations') {
        $p = $parsed.locations
        if ((Get-MemberNames -Object $p) -contains 'data_directory') { $loc.data_directory = ConvertTo-CoercedSettingString -Value $p.data_directory -Default $loc.data_directory }
        if ((Get-MemberNames -Object $p) -contains 'log_file') { $loc.log_file = ConvertTo-CoercedSettingString -Value $p.log_file -Default $loc.log_file }
        if ((Get-MemberNames -Object $p) -contains 'clone_state_file') { $loc.clone_state_file = ConvertTo-CoercedSettingString -Value $p.clone_state_file -Default $loc.clone_state_file }
        if ((Get-MemberNames -Object $p) -contains 'settings_reload_interval_seconds') { $loc.settings_reload_interval_seconds = ConvertTo-CoercedSettingInt -Value $p.settings_reload_interval_seconds -Default $loc.settings_reload_interval_seconds }
    }

    # Changing WHERE the log or clone-state file lives is never applied on a hot reload:
    # this process already has the clone-state file loaded into memory (see
    # $script:CloneStateMemory) and every in-flight clone's tracking lives there, so
    # silently switching paths mid-run would orphan every one of them. RAS restarts this
    # process on its next reconnect, which is the safe time to pick up a relocated
    # log/state file.
    if ($IsReload -and $null -ne $script:Settings) {
        $priorLoc = $script:Settings.locations
        if ($loc.data_directory -ne $priorLoc.data_directory -or
            $loc.log_file -ne $priorLoc.log_file -or
            $loc.clone_state_file -ne $priorLoc.clone_state_file) {
            Write-DebugLog "Settings reload: [locations] data_directory/log_file/clone_state_file changed -- ignored until the provider process restarts (RAS does this on its own on the next reconnect). Every other section still applied." -Level 'W' -Component '0A'
            $loc = $priorLoc
        }
    }

    $cap = [ordered]@{
        can_suspend_guests    = $defaults.capabilities.can_suspend_guests.value
        guests_polling_rate   = $defaults.capabilities.guests_polling_rate.value
        tasks_polling_rate    = $defaults.capabilities.tasks_polling_rate.value
        tasks_polling_retries = $defaults.capabilities.tasks_polling_retries.value
        template_method       = $defaults.capabilities.template_method.value
        can_link_clones       = $defaults.capabilities.can_link_clones.value
    }
    $canLinkClonesSetExplicitly = $false
    if ($null -ne $parsed -and (Get-MemberNames -Object $parsed) -contains 'capabilities') {
        $p = $parsed.capabilities
        if ((Get-MemberNames -Object $p) -contains 'can_suspend_guests') { $cap.can_suspend_guests = ConvertTo-CoercedSettingBool -Value $p.can_suspend_guests -Default $cap.can_suspend_guests }
        if ((Get-MemberNames -Object $p) -contains 'guests_polling_rate') { $cap.guests_polling_rate = ConvertTo-CoercedSettingInt -Value $p.guests_polling_rate -Default $cap.guests_polling_rate }
        if ((Get-MemberNames -Object $p) -contains 'tasks_polling_rate') { $cap.tasks_polling_rate = ConvertTo-CoercedSettingInt -Value $p.tasks_polling_rate -Default $cap.tasks_polling_rate }
        if ((Get-MemberNames -Object $p) -contains 'tasks_polling_retries') { $cap.tasks_polling_retries = ConvertTo-CoercedSettingInt -Value $p.tasks_polling_retries -Default $cap.tasks_polling_retries }
        if ((Get-MemberNames -Object $p) -contains 'template_method') { $cap.template_method = ConvertTo-CoercedSettingString -Value $p.template_method -Default $cap.template_method }

        if ((Get-MemberNames -Object $p) -contains 'can_link_clones') {
            $cap.can_link_clones = ConvertTo-CoercedSettingBool -Value $p.can_link_clones -Default $cap.can_link_clones
            $canLinkClonesSetExplicitly = $true
        }
    }
    # Deliberately OUTSIDE the 'capabilities' presence check above: a schema-1 file
    # sets cloning.linked_clones_enabled with no 'capabilities' section at all, which
    # would otherwise skip this fallback entirely and silently leave can_link_clones
    # at its default -- exactly the kind of quiet regression this migration exists to
    # avoid.
    if (-not $canLinkClonesSetExplicitly -and $null -ne $parsed -and
        (Get-MemberNames -Object $parsed) -contains 'cloning' -and
        (Get-MemberNames -Object $parsed.cloning) -contains 'linked_clones_enabled') {
        # Original key location (schema 1). capabilities.can_link_clones is now
        # the single source of truth both Handle-Initialize and Handle-GuestClone
        # read -- see the comment on its default above for why that single-source
        # property matters. Read the old key once as a migration fallback so an
        # already-live linked-clone deployment does not silently revert to the
        # default the moment this schema ships.
        $cap.can_link_clones = ConvertTo-CoercedSettingBool -Value $parsed.cloning.linked_clones_enabled -Default $cap.can_link_clones
        Write-DebugLog "Settings: cloning.linked_clones_enabled is deprecated (schema_version 1) -- read as capabilities.can_link_clones = $($cap.can_link_clones) for this load. Move this value to capabilities.can_link_clones in RAS-CPF-Proxmox-Settings.json; the old key stops being read once you do (or in a future schema version)." -Level 'W' -Component '0A'
    }

    $clo = [ordered]@{
        linked_clone_fallback = $defaults.cloning.linked_clone_fallback.value
        pipelined_cloning     = [ordered]@{
            enabled                          = $defaults.cloning.pipelined_cloning.enabled.value
            completion_seconds               = $defaults.cloning.pipelined_cloning.completion_seconds.value
            max_concurrent_clone_operations = $defaults.cloning.pipelined_cloning.max_concurrent_clone_operations.value
            start_poll_interval_seconds      = $defaults.cloning.pipelined_cloning.start_poll_interval_seconds.value
            start_max_wait_seconds           = $defaults.cloning.pipelined_cloning.start_max_wait_seconds.value
        }
        timeouts              = [ordered]@{
            delete_hard_stop_poll_interval_seconds        = $defaults.cloning.timeouts.delete_hard_stop_poll_interval_seconds.value
            delete_hard_stop_max_wait_seconds             = $defaults.cloning.timeouts.delete_hard_stop_max_wait_seconds.value
            cluster_resources_cache_ttl_seconds           = $defaults.cloning.timeouts.cluster_resources_cache_ttl_seconds.value
            recently_deleted_retention_seconds            = $defaults.cloning.timeouts.recently_deleted_retention_seconds.value
            recently_controlled_retention_seconds         = $defaults.cloning.timeouts.recently_controlled_retention_seconds.value
            recently_stopped_retention_seconds            = $defaults.cloning.timeouts.recently_stopped_retention_seconds.value
            template_delete_confirm_seconds               = $defaults.cloning.timeouts.template_delete_confirm_seconds.value
            recently_converted_retention_seconds          = $defaults.cloning.timeouts.recently_converted_retention_seconds.value
            sweep_min_interval_seconds                    = $defaults.cloning.timeouts.sweep_min_interval_seconds.value
            completed_clone_task_output_retention_seconds = $defaults.cloning.timeouts.completed_clone_task_output_retention_seconds.value
            control_retry_max_attempts                    = $defaults.cloning.timeouts.control_retry_max_attempts.value
            control_retry_delay_seconds                   = $defaults.cloning.timeouts.control_retry_delay_seconds.value
            tag_write_timeout_seconds                     = $defaults.cloning.timeouts.tag_write_timeout_seconds.value
            http_timeout_seconds                          = $defaults.cloning.timeouts.http_timeout_seconds.value
            clone_tracking_max_age_seconds                = $defaults.cloning.timeouts.clone_tracking_max_age_seconds.value
        }
        load_balancing        = [ordered]@{
            enabled                      = $defaults.cloning.load_balancing.enabled.value
            strategy                     = $defaults.cloning.load_balancing.strategy.value
            resource_metric              = $defaults.cloning.load_balancing.resource_metric.value
            excluded_nodes               = $defaults.cloning.load_balancing.excluded_nodes.value
            node_stats_cache_ttl_seconds = $defaults.cloning.load_balancing.node_stats_cache_ttl_seconds.value
            preserve_node_on_recreation  = $defaults.cloning.load_balancing.preserve_node_on_recreation.value
        }
        mac_preservation      = [ordered]@{
            enabled = $defaults.cloning.mac_preservation.enabled.value
        }
    }
    if ($null -ne $parsed -and (Get-MemberNames -Object $parsed) -contains 'cloning') {
        $p = $parsed.cloning
        if ((Get-MemberNames -Object $p) -contains 'linked_clone_fallback') { $clo.linked_clone_fallback = ConvertTo-CoercedSettingString -Value $p.linked_clone_fallback -Default $clo.linked_clone_fallback }

        if ((Get-MemberNames -Object $p) -contains 'pipelined_cloning') {
            $pp = $p.pipelined_cloning
            $pc = $clo.pipelined_cloning
            if ((Get-MemberNames -Object $pp) -contains 'enabled') { $pc.enabled = ConvertTo-CoercedSettingBool -Value $pp.enabled -Default $pc.enabled }
            if ((Get-MemberNames -Object $pp) -contains 'completion_seconds') { $pc.completion_seconds = ConvertTo-CoercedSettingInt -Value $pp.completion_seconds -Default $pc.completion_seconds }
            if ((Get-MemberNames -Object $pp) -contains 'max_concurrent_clone_operations') { $pc.max_concurrent_clone_operations = ConvertTo-CoercedSettingInt -Value $pp.max_concurrent_clone_operations -Default $pc.max_concurrent_clone_operations }
            if ((Get-MemberNames -Object $pp) -contains 'start_poll_interval_seconds') { $pc.start_poll_interval_seconds = ConvertTo-CoercedSettingInt -Value $pp.start_poll_interval_seconds -Default $pc.start_poll_interval_seconds }
            if ((Get-MemberNames -Object $pp) -contains 'start_max_wait_seconds') { $pc.start_max_wait_seconds = ConvertTo-CoercedSettingInt -Value $pp.start_max_wait_seconds -Default $pc.start_max_wait_seconds }
        }

        if ((Get-MemberNames -Object $p) -contains 'timeouts') {
            $pt2 = $p.timeouts
            $tm = $clo.timeouts
            if ((Get-MemberNames -Object $pt2) -contains 'delete_hard_stop_poll_interval_seconds') { $tm.delete_hard_stop_poll_interval_seconds = ConvertTo-CoercedSettingInt -Value $pt2.delete_hard_stop_poll_interval_seconds -Default $tm.delete_hard_stop_poll_interval_seconds }
            if ((Get-MemberNames -Object $pt2) -contains 'delete_hard_stop_max_wait_seconds') { $tm.delete_hard_stop_max_wait_seconds = ConvertTo-CoercedSettingInt -Value $pt2.delete_hard_stop_max_wait_seconds -Default $tm.delete_hard_stop_max_wait_seconds }
            if ((Get-MemberNames -Object $pt2) -contains 'cluster_resources_cache_ttl_seconds') { $tm.cluster_resources_cache_ttl_seconds = ConvertTo-CoercedSettingInt -Value $pt2.cluster_resources_cache_ttl_seconds -Default $tm.cluster_resources_cache_ttl_seconds }
            if ((Get-MemberNames -Object $pt2) -contains 'recently_deleted_retention_seconds') { $tm.recently_deleted_retention_seconds = ConvertTo-CoercedSettingInt -Value $pt2.recently_deleted_retention_seconds -Default $tm.recently_deleted_retention_seconds }
            if ((Get-MemberNames -Object $pt2) -contains 'recently_controlled_retention_seconds') { $tm.recently_controlled_retention_seconds = ConvertTo-CoercedSettingInt -Value $pt2.recently_controlled_retention_seconds -Default $tm.recently_controlled_retention_seconds }
            if ((Get-MemberNames -Object $pt2) -contains 'recently_stopped_retention_seconds') { $tm.recently_stopped_retention_seconds = ConvertTo-CoercedSettingInt -Value $pt2.recently_stopped_retention_seconds -Default $tm.recently_stopped_retention_seconds }
            if ((Get-MemberNames -Object $pt2) -contains 'template_delete_confirm_seconds') { $tm.template_delete_confirm_seconds = ConvertTo-CoercedSettingInt -Value $pt2.template_delete_confirm_seconds -Default $tm.template_delete_confirm_seconds }
            if ((Get-MemberNames -Object $pt2) -contains 'recently_converted_retention_seconds') { $tm.recently_converted_retention_seconds = ConvertTo-CoercedSettingInt -Value $pt2.recently_converted_retention_seconds -Default $tm.recently_converted_retention_seconds }
            if ((Get-MemberNames -Object $pt2) -contains 'sweep_min_interval_seconds') { $tm.sweep_min_interval_seconds = ConvertTo-CoercedSettingInt -Value $pt2.sweep_min_interval_seconds -Default $tm.sweep_min_interval_seconds }
            if ((Get-MemberNames -Object $pt2) -contains 'completed_clone_task_output_retention_seconds') { $tm.completed_clone_task_output_retention_seconds = ConvertTo-CoercedSettingInt -Value $pt2.completed_clone_task_output_retention_seconds -Default $tm.completed_clone_task_output_retention_seconds }
            if ((Get-MemberNames -Object $pt2) -contains 'control_retry_max_attempts') { $tm.control_retry_max_attempts = ConvertTo-CoercedSettingInt -Value $pt2.control_retry_max_attempts -Default $tm.control_retry_max_attempts }
            if ((Get-MemberNames -Object $pt2) -contains 'control_retry_delay_seconds') { $tm.control_retry_delay_seconds = ConvertTo-CoercedSettingInt -Value $pt2.control_retry_delay_seconds -Default $tm.control_retry_delay_seconds }
            if ((Get-MemberNames -Object $pt2) -contains 'tag_write_timeout_seconds') { $tm.tag_write_timeout_seconds = ConvertTo-CoercedSettingInt -Value $pt2.tag_write_timeout_seconds -Default $tm.tag_write_timeout_seconds }
            if ((Get-MemberNames -Object $pt2) -contains 'http_timeout_seconds') { $tm.http_timeout_seconds = ConvertTo-CoercedSettingInt -Value $pt2.http_timeout_seconds -Default $tm.http_timeout_seconds }
            if ((Get-MemberNames -Object $pt2) -contains 'clone_tracking_max_age_seconds') { $tm.clone_tracking_max_age_seconds = ConvertTo-CoercedSettingInt -Value $pt2.clone_tracking_max_age_seconds -Default $tm.clone_tracking_max_age_seconds }
        }

        if ((Get-MemberNames -Object $p) -contains 'load_balancing') {
            $plb = $p.load_balancing
            $lb = $clo.load_balancing
            if ((Get-MemberNames -Object $plb) -contains 'enabled') { $lb.enabled = ConvertTo-CoercedSettingBool -Value $plb.enabled -Default $lb.enabled }
            if ((Get-MemberNames -Object $plb) -contains 'strategy') { $lb.strategy = ConvertTo-CoercedSettingString -Value $plb.strategy -Default $lb.strategy }
            if ((Get-MemberNames -Object $plb) -contains 'resource_metric') { $lb.resource_metric = ConvertTo-CoercedSettingString -Value $plb.resource_metric -Default $lb.resource_metric }
            if ((Get-MemberNames -Object $plb) -contains 'excluded_nodes') { $lb.excluded_nodes = ConvertTo-CoercedSettingStringArray -Value $plb.excluded_nodes -Default $lb.excluded_nodes }
            if ((Get-MemberNames -Object $plb) -contains 'node_stats_cache_ttl_seconds') { $lb.node_stats_cache_ttl_seconds = ConvertTo-CoercedSettingInt -Value $plb.node_stats_cache_ttl_seconds -Default $lb.node_stats_cache_ttl_seconds }
            if ((Get-MemberNames -Object $plb) -contains 'preserve_node_on_recreation') { $lb.preserve_node_on_recreation = ConvertTo-CoercedSettingBool -Value $plb.preserve_node_on_recreation -Default $lb.preserve_node_on_recreation }
        }

        if ((Get-MemberNames -Object $p) -contains 'mac_preservation') {
            $pmp = $p.mac_preservation
            $mp = $clo.mac_preservation
            if ((Get-MemberNames -Object $pmp) -contains 'enabled') { $mp.enabled = ConvertTo-CoercedSettingBool -Value $pmp.enabled -Default $mp.enabled }
        }
    }

    $vm = [ordered]@{
        guest_agent      = [ordered]@{
            timeout_seconds          = $defaults.virtual_machines.guest_agent.timeout_seconds.value
            quarantine_after_seconds = $defaults.virtual_machines.guest_agent.quarantine_after_seconds.value
            quarantine_min_failures  = $defaults.virtual_machines.guest_agent.quarantine_min_failures.value
        }
        orphan_detection = [ordered]@{
            enabled                  = $defaults.virtual_machines.orphan_detection.enabled.value
            check_interval_seconds   = $defaults.virtual_machines.orphan_detection.check_interval_seconds.value
            stale_poll_after_seconds = $defaults.virtual_machines.orphan_detection.stale_poll_after_seconds.value
        }
        tags             = [ordered]@{
            ras_exclude_tag          = $defaults.virtual_machines.tags.ras_exclude_tag.value
            ras_quarantine_tag       = $defaults.virtual_machines.tags.ras_quarantine_tag.value
            ras_template_tag_prefix  = $defaults.virtual_machines.tags.ras_template_tag_prefix.value
            ras_clone_tag_prefix     = $defaults.virtual_machines.tags.ras_clone_tag_prefix.value
            ras_orphan_candidate_tag = $defaults.virtual_machines.tags.ras_orphan_candidate_tag.value
        }
        pool_scope       = [ordered]@{
            pool_name        = $defaults.virtual_machines.pool_scope.pool_name.value
            inherit_on_clone = $defaults.virtual_machines.pool_scope.inherit_on_clone.value
        }
    }
    if ($null -ne $parsed -and (Get-MemberNames -Object $parsed) -contains 'virtual_machines') {
        $p = $parsed.virtual_machines

        if ((Get-MemberNames -Object $p) -contains 'guest_agent') {
            $pg = $p.guest_agent
            $g = $vm.guest_agent
            if ((Get-MemberNames -Object $pg) -contains 'timeout_seconds') { $g.timeout_seconds = ConvertTo-CoercedSettingInt -Value $pg.timeout_seconds -Default $g.timeout_seconds }
            if ((Get-MemberNames -Object $pg) -contains 'quarantine_after_seconds') { $g.quarantine_after_seconds = ConvertTo-CoercedSettingInt -Value $pg.quarantine_after_seconds -Default $g.quarantine_after_seconds }
            if ((Get-MemberNames -Object $pg) -contains 'quarantine_min_failures') { $g.quarantine_min_failures = ConvertTo-CoercedSettingInt -Value $pg.quarantine_min_failures -Default $g.quarantine_min_failures }
        }

        if ((Get-MemberNames -Object $p) -contains 'orphan_detection') {
            $po = $p.orphan_detection
            $o = $vm.orphan_detection
            if ((Get-MemberNames -Object $po) -contains 'enabled') { $o.enabled = ConvertTo-CoercedSettingBool -Value $po.enabled -Default $o.enabled }
            if ((Get-MemberNames -Object $po) -contains 'check_interval_seconds') { $o.check_interval_seconds = ConvertTo-CoercedSettingInt -Value $po.check_interval_seconds -Default $o.check_interval_seconds }
            if ((Get-MemberNames -Object $po) -contains 'stale_poll_after_seconds') { $o.stale_poll_after_seconds = ConvertTo-CoercedSettingInt -Value $po.stale_poll_after_seconds -Default $o.stale_poll_after_seconds }
        }

        if ((Get-MemberNames -Object $p) -contains 'tags') {
            $pt = $p.tags
            $t = $vm.tags
            if ((Get-MemberNames -Object $pt) -contains 'ras_exclude_tag') { $t.ras_exclude_tag = ConvertTo-CoercedSettingString -Value $pt.ras_exclude_tag -Default $t.ras_exclude_tag }
            if ((Get-MemberNames -Object $pt) -contains 'ras_quarantine_tag') { $t.ras_quarantine_tag = ConvertTo-CoercedSettingString -Value $pt.ras_quarantine_tag -Default $t.ras_quarantine_tag }
            if ((Get-MemberNames -Object $pt) -contains 'ras_template_tag_prefix') { $t.ras_template_tag_prefix = ConvertTo-CoercedSettingString -Value $pt.ras_template_tag_prefix -Default $t.ras_template_tag_prefix }
            if ((Get-MemberNames -Object $pt) -contains 'ras_clone_tag_prefix') { $t.ras_clone_tag_prefix = ConvertTo-CoercedSettingString -Value $pt.ras_clone_tag_prefix -Default $t.ras_clone_tag_prefix }
            if ((Get-MemberNames -Object $pt) -contains 'ras_orphan_candidate_tag') { $t.ras_orphan_candidate_tag = ConvertTo-CoercedSettingString -Value $pt.ras_orphan_candidate_tag -Default $t.ras_orphan_candidate_tag }
        }

        if ((Get-MemberNames -Object $p) -contains 'pool_scope') {
            $pp2 = $p.pool_scope
            $ps = $vm.pool_scope
            if ((Get-MemberNames -Object $pp2) -contains 'pool_name') { $ps.pool_name = ConvertTo-CoercedSettingString -Value $pp2.pool_name -Default $ps.pool_name }
            if ((Get-MemberNames -Object $pp2) -contains 'inherit_on_clone') { $ps.inherit_on_clone = ConvertTo-CoercedSettingBool -Value $pp2.inherit_on_clone -Default $ps.inherit_on_clone }
        }
    }

    $log = [ordered]@{
        log_level                  = $defaults.logging.log_level.value
        log_rotate_max_mb          = $defaults.logging.log_rotate_max_mb.value
        log_rotate_max_generations = $defaults.logging.log_rotate_max_generations.value
    }
    if ($null -ne $parsed -and (Get-MemberNames -Object $parsed) -contains 'logging') {
        $p = $parsed.logging
        if ((Get-MemberNames -Object $p) -contains 'log_level') { $log.log_level = ConvertTo-CoercedLogLevel -Value $p.log_level -Default $log.log_level }
        if ((Get-MemberNames -Object $p) -contains 'log_rotate_max_mb') { $log.log_rotate_max_mb = ConvertTo-CoercedSettingInt -Value $p.log_rotate_max_mb -Default $log.log_rotate_max_mb }
        if ((Get-MemberNames -Object $p) -contains 'log_rotate_max_generations') { $log.log_rotate_max_generations = ConvertTo-CoercedSettingInt -Value $p.log_rotate_max_generations -Default $log.log_rotate_max_generations }
    }

    $script:Settings = [ordered]@{
        schema_version   = $schemaVersion
        locations        = $loc
        capabilities     = $cap
        cloning          = $clo
        virtual_machines = $vm
        logging          = $log
    }

    # Upgrade an older file in place -- see Update-SettingsFileSchema above. Only for a
    # file that PARSED (a corrupt one is never touched, same rule as everywhere else),
    # is genuinely older than this script,
    # and is schema 2 or newer -- deliberately NOT for a genuine schema-1 file. Schema
    # 1->2 RELOCATED several keys, and only capabilities.can_link_clones has a real
    # fallback-read from its old location (see the migration-fallback block above);
    # every other schema-1 key sitting at its OLD path is simply invisible to this
    # script and $script:Settings has already resolved it to a bare DEFAULT by this
    # point. Auto-rewriting a schema-1 file would bake that default into the file
    # permanently, silently discarding whatever the admin actually had customized under
    # the old flat shape -- a real, if narrow, data-loss risk. Schema 2->3
    # was purely additive (no relocations), so this risk does not apply there: every
    # value $script:Settings resolved for a schema-2-or-newer file is exactly what was on
    # disk (or this script's own default for a key that's genuinely new), safe to write
    # straight back. A schema-1 file still gets the informational mismatch its own
    # Write-DebugLog call above already covers; see SETTINGS.md's migration notes for the
    # 1->2 key-relocation table an admin still has to apply by hand in that case.
    #
    # $script:Settings.schema_version is bumped to match immediately after a successful
    # rewrite, so the rest of this load (and every read of $script:Settings.schema_version
    # for the remainder of this process, until the next reload re-reads the now-current
    # file) reports the version this process is actually running against, not the stale
    # one the file had on disk before this ran.
    if ($fileExists -and $null -ne $parsed -and $schemaVersion -ge 2 -and $schemaVersion -lt $script:CurrentSettingsSchemaVersion) {
        Update-SettingsFileSchema -FromVersion $schemaVersion -ToVersion $script:CurrentSettingsSchemaVersion -OldParsed $parsed
        $script:Settings.schema_version = $script:CurrentSettingsSchemaVersion
    }

    if (-not $IsReload) {
        # First load: report every leaf whose effective value differs from this
        # provider's hardcoded default -- i.e. exactly what RAS-CPF-Proxmox-Settings.json
        # actually changed from a stock install, without printing the whole (large)
        # settings tree on every single startup. schema_version has no default to
        # compare against (Get-DefaultProviderSettings doesn't carry one) and is
        # correctly skipped by the Contains() check below.
        $flatSettings = ConvertTo-FlatSettingsMap -Node $script:Settings
        $flatDefaults = ConvertTo-FlatSettingsMap -Node $defaults
        $nonDefault = @(foreach ($key in $flatSettings.Keys) {
                if ($flatDefaults.Contains($key) -and $flatSettings[$key] -ne $flatDefaults[$key]) {
                    "$key = [$($flatSettings[$key])] (default: [$($flatDefaults[$key])])"
                }
            })
        if ($nonDefault.Count -gt 0) {
            Write-DebugLog "Settings load: non-default value(s) picked up from [$($script:SettingsPath)]:`n  $($nonDefault -join "`n  ")" -Level 'I' -Component '0A'
        }
        else {
            Write-DebugLog "Settings load: [$($script:SettingsPath)] -- every value at its hardcoded default." -Level 'I' -Component '0A'
        }
    }
    elseif ($null -ne $priorSettings) {
        # Reload: report only what actually changed since the last load -- not the whole
        # tree, and not a diff against defaults, so a value an admin already customised
        # on purpose doesn't get renotified about forever. This is the direct answer to
        # "why did behaviour change without me touching settings": whatever fired this
        # reload (a real edit, a redeploy/reseed overwriting the file, or the schema-1
        # cloning.linked_clones_enabled fallback -- see SETTINGS.md's migration notes)
        # shows up here with the exact before/after value, every time.
        $flatBefore = ConvertTo-FlatSettingsMap -Node $priorSettings
        $flatAfter = ConvertTo-FlatSettingsMap -Node $script:Settings
        $changed = @(foreach ($key in $flatAfter.Keys) {
                $before = if ($flatBefore.Contains($key)) { $flatBefore[$key] } else { '<absent before>' }
                if ($before -ne $flatAfter[$key]) {
                    "$key changed [$before] -> [$($flatAfter[$key])]"
                }
            })
        if ($changed.Count -gt 0) {
            Write-DebugLog "Settings reload: value(s) applied from [$($script:SettingsPath)]:`n  $($changed -join "`n  ")" -Level 'I' -Component '0A'
        }
    }
}

# Applies $script:Settings to every $script:* variable the rest of this script actually
# reads. Keeping those variable names unchanged means this is the ONLY place that needs
# to know the mapping between the two.
function Set-ProviderRuntimeFromSettings {
    param([Parameter(Mandatory = $true)][object]$Settings)

    $dataDir = $Settings.locations.data_directory
    if ([string]::IsNullOrWhiteSpace($dataDir)) { $dataDir = $PSScriptRoot }
    if (-not [System.IO.Path]::IsPathRooted($dataDir)) { $dataDir = Join-Path -Path $PSScriptRoot -ChildPath $dataDir }

    $logFile = $Settings.locations.log_file
    $script:LogPath = if ([System.IO.Path]::IsPathRooted($logFile)) { $logFile } else { Join-Path -Path $dataDir -ChildPath $logFile }

    $stateFile = $Settings.locations.clone_state_file
    $script:CloneStatePath = if ([System.IO.Path]::IsPathRooted($stateFile)) { $stateFile } else { Join-Path -Path $dataDir -ChildPath $stateFile }

    $script:SettingsReloadIntervalSeconds = $Settings.locations.settings_reload_interval_seconds

    $script:PreserveMacOnRecreation = $Settings.cloning.mac_preservation.enabled

    $script:PipelinedCloneCompletionEnabled = $Settings.cloning.pipelined_cloning.enabled
    $script:PipelinedCloneCompletionSeconds = $Settings.cloning.pipelined_cloning.completion_seconds
    $script:MaxConcurrentCloneOperations = $Settings.cloning.pipelined_cloning.max_concurrent_clone_operations
    $script:PipelinedCloneStartPollIntervalSeconds = $Settings.cloning.pipelined_cloning.start_poll_interval_seconds
    $script:PipelinedCloneStartMaxWaitSeconds = $Settings.cloning.pipelined_cloning.start_max_wait_seconds
    $script:DeleteHardStopPollIntervalSeconds = $Settings.cloning.timeouts.delete_hard_stop_poll_interval_seconds
    $script:DeleteHardStopMaxWaitSeconds = $Settings.cloning.timeouts.delete_hard_stop_max_wait_seconds
    $script:ClusterResourcesCacheTtlSeconds = $Settings.cloning.timeouts.cluster_resources_cache_ttl_seconds
    $script:RecentlyDeletedRetentionSeconds = $Settings.cloning.timeouts.recently_deleted_retention_seconds
    $script:RecentlyControlledRetentionSeconds = $Settings.cloning.timeouts.recently_controlled_retention_seconds
    $script:RecentlyStoppedRetentionSeconds = $Settings.cloning.timeouts.recently_stopped_retention_seconds
    $script:TemplateDeleteConfirmSeconds = $Settings.cloning.timeouts.template_delete_confirm_seconds
    $script:RecentlyConvertedRetentionSeconds = $Settings.cloning.timeouts.recently_converted_retention_seconds
    $script:SweepMinIntervalSeconds = $Settings.cloning.timeouts.sweep_min_interval_seconds
    $script:CompletedCloneTaskOutputRetentionSeconds = $Settings.cloning.timeouts.completed_clone_task_output_retention_seconds
    $script:ControlRetryMaxAttempts = $Settings.cloning.timeouts.control_retry_max_attempts
    $script:ControlRetryDelaySeconds = $Settings.cloning.timeouts.control_retry_delay_seconds
    $script:TagWriteTimeoutSeconds = $Settings.cloning.timeouts.tag_write_timeout_seconds
    $script:HttpTimeoutSeconds = $Settings.cloning.timeouts.http_timeout_seconds
    $script:CloneTrackingMaxAgeSeconds = $Settings.cloning.timeouts.clone_tracking_max_age_seconds

    $script:RasExcludeTag = $Settings.virtual_machines.tags.ras_exclude_tag
    $script:RasQuarantineTag = $Settings.virtual_machines.tags.ras_quarantine_tag
    $script:RasTemplateTagPrefix = $Settings.virtual_machines.tags.ras_template_tag_prefix
    $script:RasCloneTagPrefix = $Settings.virtual_machines.tags.ras_clone_tag_prefix
    $script:RasOrphanCandidateTag = $Settings.virtual_machines.tags.ras_orphan_candidate_tag

    $script:PoolScopeName = $Settings.virtual_machines.pool_scope.pool_name
    $script:PoolScopeInheritOnClone = $Settings.virtual_machines.pool_scope.inherit_on_clone

    $script:GuestAgentTimeoutSeconds = $Settings.virtual_machines.guest_agent.timeout_seconds
    $script:AgentQuarantineAfterSeconds = $Settings.virtual_machines.guest_agent.quarantine_after_seconds
    $script:AgentQuarantineMinFailures = $Settings.virtual_machines.guest_agent.quarantine_min_failures

    $script:OrphanDetectionEnabled = $Settings.virtual_machines.orphan_detection.enabled
    $script:OrphanDetectionCheckIntervalSeconds = $Settings.virtual_machines.orphan_detection.check_interval_seconds
    $script:OrphanDetectionStalePollAfterSeconds = $Settings.virtual_machines.orphan_detection.stale_poll_after_seconds

    $script:LogLevel = $Settings.logging.log_level
    # logging.log_rotate_max_mb is stored (and hand-edited) in MB; Invoke-LogRotation
    # compares against a running byte count, so this is the one place the conversion
    # happens. 1MB is PowerShell's own binary-megabyte literal (1048576), matching this
    # value's own prior hardcoded default of 20971520 bytes (= 20 * 1MB).
    $script:LogRotateMaxBytes = $Settings.logging.log_rotate_max_mb * 1MB
    $script:LogRotateMaxGenerations = $Settings.logging.log_rotate_max_generations

    if ($script:PipelinedCloneCompletionSeconds -ge $Settings.capabilities.tasks_polling_rate) {
        Write-DebugLog "Settings: cloning.pipelined_cloning.completion_seconds ($($script:PipelinedCloneCompletionSeconds)) should be below capabilities.tasks_polling_rate ($($Settings.capabilities.tasks_polling_rate)) -- otherwise a pipelined clone meeting the elapsed threshold may not be reported 'completed' until RAS's next scheduled tasks/get poll, silently falling back to the tasks_polling_rate cadence. Check RAS-CPF-Proxmox-Settings.json." -Level 'W' -Component '0A'
    }
}

# Opportunistic hot-reload: checked at most once every
# locations.settings_reload_interval_seconds, from Process-Method, so a settings file
# edited mid-run takes effect without restarting the provider process -- with the one
# exception of Locations (see Import-ProviderSettings). Uses the file's own last-write
# time rather than hashing its contents, so the overwhelmingly common "nothing changed"
# case costs a single stat.
function Update-ProviderSettingsIfChanged {
    if (([DateTime]::UtcNow - $script:SettingsLastPolledAt).TotalSeconds -lt $script:SettingsReloadIntervalSeconds) {
        return
    }

    $script:SettingsLastPolledAt = [DateTime]::UtcNow

    if (-not (Test-Path -LiteralPath $script:SettingsPath)) {
        return
    }

    try {
        $mtime = (Get-Item -LiteralPath $script:SettingsPath).LastWriteTimeUtc
    }
    catch {
        return
    }

    if ($null -ne $script:SettingsFileLastWriteUtc -and $mtime -eq $script:SettingsFileLastWriteUtc) {
        return
    }

    $script:SettingsFileLastWriteUtc = $mtime
    Import-ProviderSettings -IsReload
    Set-ProviderRuntimeFromSettings -Settings $script:Settings
    Write-DebugLog "Settings reloaded from [$($script:SettingsPath)] (changed on disk)." -Level 'I' -Component '0A'
}

# Must stay ahead of the Import-ProviderSettings call just below: that call logs via
# Write-DebugLog, and rotation itself depends on $script:LogRotateMaxBytes, which that
# call resolves.
function Invoke-LogRotation {
    try {
        $oldest = "$($script:LogPath).$($script:LogRotateMaxGenerations)"
        if (Test-Path $oldest) {
            Remove-Item -Path $oldest -Force
        }

        for ($i = $script:LogRotateMaxGenerations - 1; $i -ge 1; $i--) {
            $src = "$($script:LogPath).$i"
            $dst = "$($script:LogPath).$($i + 1)"
            if (Test-Path $src) {
                Move-Item -Path $src -Destination $dst -Force
            }
        }

        if (Test-Path $script:LogPath) {
            Move-Item -Path $script:LogPath -Destination "$($script:LogPath).1" -Force
        }

        $script:LogBytesWrittenSinceStart = 0
    }
    catch {
        # Rotation is best-effort -- if it fails, keep logging to the existing file
        # rather than losing diagnostics or throwing from a logging path.
    }
}

function Write-DebugLog {
    param(
        [string]$Message,

        # E/W: always logged, regardless of log_level -- a genuine failure (E) or
        # something degraded/worth an admin's attention (W), never gated by verbosity.
        # I/T/D: gated on settings.logging.log_level exactly as the old 3/4/5 scheme was
        # (I >= 3 "Standard", T >= 4 "Extended", D >= 5 "Verbose") -- see LOGGING.md. A
        # call site that omits -Level defaults to 'D', same as the old unmarked-call
        # (Verbose-only) behaviour it replaces.
        [ValidateSet('E', 'W', 'I', 'T', 'D')]
        [string]$Level = 'D',

        # Two-hex-digit component code identifying which area of the provider logged
        # this -- see LOGGING.md's component table. Defaults to '00' (Core/process).
        [string]$Component = '00',

        # The VMID or task id this line concerns, if any. Rendered into the bracket so
        # `grep "/136/"` isolates every line about one guest/clone across every
        # component (connect, clone, control, tag, settings-unrelated lines excluded) --
        # the clone-tracking trail this format exists to support. '-' when the line
        # isn't scoped to one specific guest/task.
        $Ref = $null
    )

    if ($Level -ne 'E' -and $Level -ne 'W') {
        $minLevelForLetter = @{ I = 3; T = 4; D = 5 }
        if ($minLevelForLetter[$Level] -gt $script:LogLevel) {
            return
        }
    }

    try {
        # [DateTime]::Now.ToString(...) rather than Get-Date -Format: far cheaper, and
        # this runs on every single log line. dd-MM-yy HH:mm:ss deliberately matches
        # RAS's own native module log format (CustomProvider.log/vdiagent.log/*.log) so
        # this file can be read side by side with them without a date-format translation.
        $timestamp = [DateTime]::Now.ToString('dd-MM-yy HH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture)
        $refText = if ($null -eq $Ref -or [string]::IsNullOrEmpty([string]$Ref)) { '-' } else { [string]$Ref }
        $pidHex = '{0:X4}' -f $PID
        $line = "[$Level $Component/$refText/P$pidHex] $timestamp - $Message"

        if ($script:LogBytesWrittenSinceStart -ge $script:LogRotateMaxBytes) {
            Invoke-LogRotation
        }

        Add-Content -Path $script:LogPath -Value $line -Encoding UTF8
        # Approximate (Add-Content's own newline isn't counted exactly) -- fine,
        # this only gates when rotation kicks in, not a byte-exact requirement.
        $script:LogBytesWrittenSinceStart += ([System.Text.Encoding]::UTF8.GetByteCount($line) + 2)
    }
    catch {
        # Never emit logging failures to stdout
    }
}

Import-ProviderSettings
Set-ProviderRuntimeFromSettings -Settings $script:Settings
try {
    $script:SettingsFileLastWriteUtc = (Get-Item -LiteralPath $script:SettingsPath).LastWriteTimeUtc
}
catch { }
$script:SettingsLastPolledAt = [DateTime]::UtcNow

$script:ClusterResourcesCache = $null
$script:ClusterResourcesCachedAt = $null

# The cluster/resources cache is otherwise only ever invalidated by its TTL, which let a
# clone submitted at T0 be invisible to a guests/get moments later (a hard "not found in
# cluster" JSON-RPC error, which RAS's clone thread does not retry) and let a
# just-deleted VM keep appearing in guests/list. Called after guests/clone and after
# every guests/control action, so the next guests/get or guests/list is forced to
# re-fetch.
function Reset-ProxmoxClusterCache {
    $script:ClusterResourcesCache = $null
    $script:ClusterResourcesCachedAt = $null
}

# Separate cache from the VM listing above -- different endpoint (?type=node vs
# ?type=vm), different consumer (Resolve-ProxmoxCloneTargetNode, not guests/get), and a
# different freshness need (feeds a scheduling decision made once per clone, not a
# per-poll display value). See placement.node_stats_cache_ttl_seconds.
$script:ClusterNodesCache = $null
$script:ClusterNodesCachedAt = $null

# Cursor for placement.strategy = 'round_robin' -- see Resolve-ProxmoxCloneTargetNode.
# Process-lifetime only, like $script:RasTemplateTagAttempted; a restart resets the
# rotation rather than trying to persist it, which is fine for a fairness heuristic, not
# a strict guarantee.
$script:PlacementRoundRobinIndex = 0

# Belt-and-braces for the same race: report a freshly pipelined clone 'completed' only
# once its VM id is actually resolvable, and never surface a hard error for an id this
# script itself knows is an in-flight clone (report it as still provisioning instead).
# See Handle-TaskInfo and Handle-GuestGet.

$script:RasTemplateTagAttempted = [System.Collections.Generic.HashSet[string]]::new()
$script:AgentFailureTracker = @{}   # vmid -> @{ count; first_failure_at }

# One log line per continuous failure episode for a RAS-managed VM that is exempt from
# rasQuarantine promotion (see Get-ProxmoxVmGuestAgentInterfaces) -- without this, a
# genuinely agent-less rasClone/rasTemplate-tagged VM would re-log the same "exempt"
# message on every single poll for as long as it stays broken. Cleared alongside
# $script:AgentFailureTracker's own entry on the next successful agent query, so a fresh
# failure episode logs again.
$script:RasManagedAgentExemptionLogged = [System.Collections.Generic.HashSet[string]]::new()

# Last-known-good IP/MAC, so a transient agent failure reports the previous value
# instead of retracting it to empty. Cleared when the VM is no longer 'running' or is
# deleted, so a genuinely reassigned address is never served stale.
$script:LastKnownNetworkData = @{}   # vmid -> @{ IPv4Addresses; MacAddresses }

# guests/list is the ONLY thing that retires a guest on the RAS side (a guests/get "not
# found" error is logged and ignored), so a destroyed VM lingering in guests/list costs
# a full extra RAS reconciliation cycle. Ids tracked here are omitted from guests/list
# immediately, independent of Proxmox's own cluster/resources view. Entries expire on
# their own, so a reused VMID is never permanently hidden.
$script:RecentlyDeletedIds = @{}   # vmid -> deleted_at (UTC DateTime)

# The VM's own name at the moment THIS provider deleted it, so the recently-deleted stub
# in Handle-GuestGet/Handle-HostGet (below) can echo it back instead of a synthetic
# "VM-<id>". RAS treats every successful guests/get response as fresh guest info and
# updates its own displayed name immediately (confirmed via vdiagent.log: a single stub
# poll landing in the delete-to-recreate gap was enough to rename the guest to "VM-141"
# in RAS's own tracking, until RAS's bulk-recreate flow later tore down and rebuilt that
# guest object with the right name) -- so during a bulk recreate, whichever VMs happen to
# get polled in that narrow gap show up under a VM-ID-shaped name for a while. Echoing the
# last real name instead removes the visible symptom without touching delete/recreate
# timing, which this provider does not control. Same lifetime/expiry as
# $script:RecentlyDeletedIds -- always cleared alongside it, never read once its id is.
$script:RecentlyDeletedNames = @{}   # vmid -> name (string) as of the delete call

# The mirror lookup, keyed the other way: the NODE a VM was on at the moment this
# provider deleted it, keyed by that VM's NAME (a fresh guests/clone request carries only
# a desired name, never the old id it may be recreating). Feeds placement's
# preserve_node_on_recreation -- see DISTRIBUTED-PLACEMENT.md and
# Resolve-ProxmoxCloneTargetNode. Same lifetime/expiry as $script:RecentlyDeletedIds,
# checked against cloning.recently_deleted_retention_seconds by
# Get-ProxmoxPreservedNodeForRecreation rather than tracked with its own separate timer.
$script:RecentlyDeletedNodeByName = @{}   # name -> @{ node; deleted_at (UTC DateTime) }

# Same shape and purpose as $script:RecentlyDeletedNodeByName immediately above, for
# cloning.mac_preservation.enabled instead of placement.preserve_node_on_recreation --
# see MAC-PRESERVATION.md. Populated (best-effort, only when the setting is on -- the
# live config read it costs is skipped entirely otherwise) by Handle-GuestControl's
# delete branch; read once by Handle-GuestClone at clone-initiation time, then stashed
# into that clone's own tracking context (Set-CloneStateEntry/$script:TaskContext) so
# Start-ProxmoxVmIfNeeded can apply it after the clone genuinely completes without
# re-deriving it from a window that may have long since expired by then.
$script:RecentlyDeletedMacByName = @{}   # name -> @{ macs = @{netKey -> MAC}; deleted_at (UTC DateTime) }

# cluster/resources is a lagging, server-side aggregate (pvestatd refreshes it roughly
# every ~10s), so busting our own cache after a control action (see
# Reset-ProxmoxClusterCache) only guarantees a fresh HTTP fetch, not a fresh view of the
# world. RAS never polls the returned stop/destroy task either, so it learns the real
# state only from this provider's next guests/get. For a short window after WE issue a
# control action, power state is resolved from qemu/{id}/status/current (the qemu
# monitor directly, no aggregation lag) instead of the cluster listing.
$script:RecentlyControlledIds = @{}   # vmid -> controlled_at (UTC DateTime)
# Narrower than RecentlyControlledIds above: only a 'stop' sets this, and any OTHER
# control action on the same vmid clears it immediately (see Handle-GuestControl). Lets
# Get-ProxmoxVmNetworkData skip a guest-agent probe that would otherwise reliably race
# Proxmox's own teardown.
$script:RecentlyStoppedIds = @{}       # vmid -> stopped_at (UTC DateTime)

# RAS drives both "enter maintenance" and "delete the Template object" through the same
# guests/convert {is_template:false} call -- the wire calls are otherwise identical. The
# reliable difference is what happens right after: entering maintenance always follows
# with guests/control start (RAS needs the guest running to patch it); deleting the
# Template object never does (RAS is only tidying up its own object, not touching guest
# power state). Since this provider has no timer and can only act when RAS calls it, the
# convert(false) handler arms this marker instead of deciding immediately, a subsequent
# 'start' for the same vmid cancels it (see Handle-GuestControl), and
# Resolve-PendingTemplateTagRemoval -- piggybacked on the next guests/get for that vmid,
# opportunistically, like the tracked-clone sweep -- removes the rasTemplate<id> tag once
# the grace window passes with no start. See Set-ProxmoxTemplateSourceTagBestEffort.
$script:PendingTemplateTagRemovalIds = @{}   # vmid -> de-templated_at (UTC DateTime)

# Same lag, same fix, for the template flag: cluster/resources carries 'template' too
# and it is just as stale, which RAS reads as "the convert did not take". For this
# window after we convert a VM, is_template is resolved from qemu/{id}/status/current
# (or, failing that, from what we just asked Proxmox to do) instead of the cluster
# listing. See Test-ProxmoxRecentlyConverted.
$script:RecentlyConvertedIds = @{}   # vmid -> @{ is_template; converted_at }

# Templates this provider has de-templated for a RAS maintenance session and
# not yet converted back. See Set-ProxmoxRecentConvert.
$script:MaintenanceModeVmIds = [System.Collections.Generic.HashSet[string]]::new()

# Baseline for "this provider has never heard about that VM" checks, which are
# only meaningful relative to how long this process has actually been up.
$script:ProviderStartedAt = [DateTime]::UtcNow

# Throttle record for Invoke-TrackedCloneSweep, which otherwise re-runs the full
# clone-aware flow (a task-status poll, a state read, sometimes an agent probe) for
# every tracked clone on every single guests/get. Safe to throttle because
# ConvertTo-RasGuestObject substitutes a tracked clone's real name for Proxmox's
# placeholder, so RAS can poll a clone directly instead of depending on this sweep to
# ever reach it.
$script:SweepLastCheckedAt = @{}   # vmid -> UTC DateTime of last sweep check

# Two independent code paths can each notice a clone has become ready and finalize it:
# Handle-TaskInfo's own tasks/get poll, and Get-RasGuestObjectForCloneAwareFlow (driven
# by ANY guests/get, including the opportunistic sweep above). Both call
# Remove-CloneStateEntry, which erases the only place Handle-TaskInfo can look up a
# task's clone_id once its in-memory $script:TaskContext entry is gone. RAS expects
# output.clone_id on a clone task's completion and does not retry once it gets an
# answer, so record the mapping here, independent of the main tracking entry, and let
# Handle-TaskInfo still answer correctly after that race.
$script:CompletedCloneTaskOutputs = @{}   # taskId -> @{ clone_id; completed_at (UTC) }

# Once a tracked clone's real Proxmox task is confirmed 'completed' it never needs
# asking again -- cached here by task id so Get-ActiveCloneCount does not re-poll
# Proxmox for a task everyone already knows is done. Cleared in
# Clear-ProxmoxTrackingForVm.
$script:CloneTaskCompletionCache = @{}   # taskId -> $true once confirmed completed

# Last time this provider answered a guests/get for a given vmid, used by the
# orphan-candidate audit (see orphan_detection above and Invoke-OrphanAudit) to tell
# "RAS still actively polling this guest" apart from "RAS has stopped asking about it
# entirely".
$script:LastGuestPollAt = @{}   # vmid -> UTC DateTime of last guests/get answered
$script:OrphanAuditLastCheckedAt = [DateTime]::MinValue

# PowerShell has no built-in log rotation and this log otherwise grows forever. Track
# the byte count written since this process started (cheap; avoids a Get-Item stat call
# per line) and rotate when it crosses the threshold, keeping a bounded number of prior
# generations.
$script:LogBytesWrittenSinceStart = 0
if (Test-Path $script:LogPath) {
    try { $script:LogBytesWrittenSinceStart = (Get-Item $script:LogPath).Length } catch { }
}

$script:ErrorCodes = @{
    ParseError     = -32700
    MethodNotFound = -32601
    InvalidParams  = -32602
    InternalError  = -32603
}

function Get-CloneStateEntryByTaskId {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TaskId
    )

    $all = Get-CloneStateAll

    foreach ($key in $all.Keys) {
        $entry = $all[$key]

        if ($null -ne $entry -and
            (Get-MemberNames -Object $entry) -contains 'task_id' -and
            [string]$entry.task_id -eq [string]$TaskId) {

            $ctx = @{}
            foreach ($p in $entry.PSObject.Properties) {
                $ctx[$p.Name] = $p.Value
            }

            if (-not $ctx.ContainsKey('type')) {
                $ctx.type = 'clone'
            }

            if (-not $ctx.ContainsKey('clone_id')) {
                $ctx.clone_id = [string]$key
            }

            Write-DebugLog "CLONE STATE FOUND BY TASK ID for task [$TaskId], clone VM [$($ctx.clone_id)]" -Level 'D' -Component '04' -Ref $TaskId
            return $ctx
        }
    }

    Write-DebugLog "CLONE STATE NOT FOUND BY TASK ID for task [$TaskId]" -Level 'D' -Component '04' -Ref $TaskId
    return $null
}

# In-memory mirror of the clone-state file, behind Get-CloneStateAll. The file is only
# ever written by THIS process (Set-/Remove-CloneStateEntry below), so the mirror is
# always correct without a re-read; it only needs to hit disk once, lazily, to recover
# state a prior process instance may have left behind (e.g. after a provider restart
# mid-clone).
$script:CloneStateMemory = $null

function Read-CloneStateFromDisk {
    try {
        if (-not (Test-Path -LiteralPath $script:CloneStatePath)) {
            return @{}
        }

        $raw = Get-Content -LiteralPath $script:CloneStatePath -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($raw)) {
            return @{}
        }

        $obj = $raw | ConvertFrom-Json -ErrorAction Stop
        $state = @{}
        foreach ($p in $obj.PSObject.Properties) {
            $state[$p.Name] = $p.Value
        }

        return $state
    }
    catch {
        Write-DebugLog "Failed to read clone state file: $($_.Exception.Message)" -Level 'E' -Component '04'
        return @{}
    }
}

function Get-CloneState {
    if ($null -eq $script:CloneStateMemory) {
        $script:CloneStateMemory = Read-CloneStateFromDisk
    }

    return $script:CloneStateMemory
}

function Get-CloneStateAll {
    return Get-CloneState
}

function Save-CloneState {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$State
    )

    $script:CloneStateMemory = $State

    try {
        $dir = Split-Path -Path $script:CloneStatePath -Parent
        if (-not [string]::IsNullOrWhiteSpace($dir) -and -not (Test-Path -LiteralPath $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
        }

        $json = $State | ConvertTo-Json -Depth 10 -Compress
        Set-Content -LiteralPath $script:CloneStatePath -Value $json -Encoding UTF8
    }
    catch {
        Write-DebugLog "Failed to save clone state: $($_.Exception.Message)" -Level 'E' -Component '04'
    }
}

function Set-CloneStateEntry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$VmId,

        [Parameter(Mandatory = $true)]
        [hashtable]$Entry
    )

    $state = Get-CloneState

    # Normalized to PSCustomObject so every entry -- freshly written here or loaded from
    # disk -- exposes the same .PSObject.Properties shape that Get-ActiveCloneCount /
    # Get-CloneStateEntryByTaskId / Get-TrackedCloneContextByVmId all rely on to inspect
    # entry fields. A raw Hashtable's own .PSObject.Properties lists Keys/Values/Count
    # etc., not the dictionary's own keys, so caching one unconverted would make those
    # checks miss the entry that was just written.
    $state[[string]$VmId] = [PSCustomObject]$Entry
    Save-CloneState -State $state

    # Kept in sync so Invoke-TrackedCloneSweep (see Handle-GuestGet) can tell in O(1),
    # with no file read, whether any clone is in flight at all.
    if ($Entry.ContainsKey('type') -and [string]$Entry.type -eq 'clone') {
        [void]$script:TrackedCloneVmIds.Add([string]$VmId)
    }
}

# Counts only clones whose real underlying Proxmox clone task is not yet CONFIRMED
# finished -- not every clone that is not yet READY. The gate this feeds
# (cloning.max_concurrent_clone_operations) exists to protect Proxmox's disk-I/O-heavy
# clone jobs, and the boot-to-ready tail does zero disk I/O, so counting it starves real
# disk-clone concurrency. Confirmed completions are cached (see
# $script:CloneTaskCompletionCache) so a clone sitting in its boot tail never costs
# another task-status poll.
function Get-ActiveCloneCount {
    $all = Get-CloneStateAll
    $count = 0
    foreach ($key in $all.Keys) {
        $entry = $all[$key]
        if ($null -eq $entry -or -not ((Get-MemberNames -Object $entry) -contains 'type') -or [string]$entry.type -ne 'clone') {
            continue
        }

        $taskId = $null
        if ((Get-MemberNames -Object $entry) -contains 'task_id' -and -not [string]::IsNullOrWhiteSpace([string]$entry.task_id)) {
            $taskId = [string]$entry.task_id
        }

        if ([string]::IsNullOrWhiteSpace($taskId)) {
            # No real task id to check against (shouldn't normally happen for
            # a clone entry) -- count it, matching the old conservative
            # behavior for this edge case.
            $count++
            continue
        }

        if ($script:CloneTaskCompletionCache.ContainsKey($taskId)) {
            continue   # confirmed finished on a previous check -- not active
        }

        try {
            $realState = (New-TaskResultState -TaskStatus (Get-ProxmoxTaskStatus -TaskId $taskId)).state
            if ($realState -eq 'completed') {
                $script:CloneTaskCompletionCache[$taskId] = $true
            }
            else {
                $count++
            }
        }
        catch {
            # Can't determine the real task's state (e.g. aged out of
            # Proxmox's task history) -- count it rather than risk
            # under-gating real concurrency.
            $count++
        }
    }
    return $count
}

function Get-CloneStateEntry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$VmId
    )

    $all = Get-CloneStateAll

    if ($all.ContainsKey($VmId)) {
        return $all[$VmId]
    }

    # Fallback: match by clone_id
    foreach ($key in $all.Keys) {
        $entry = $all[$key]

        if ($null -ne $entry -and
            (Get-MemberNames -Object $entry) -contains 'clone_id' -and
            [string]$entry.clone_id -eq [string]$VmId) {

            Write-DebugLog "CLONE STATE MATCHED via clone_id for VM [$VmId] (key [$key])" -Level 'D' -Component '04' -Ref $VmId
            return $entry
        }
    }

    return $null
}

function Remove-CloneStateEntry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$VmId
    )

    $state = Get-CloneState
    if ($state.ContainsKey([string]$VmId)) {
        $state.Remove([string]$VmId)
        Save-CloneState -State $state
    }

    [void]$script:TrackedCloneVmIds.Remove([string]$VmId)
    # So a VMID Proxmox later reuses for an unrelated clone re-verifies its tags fresh --
    # see $script:CloneTagVerified near the top of this file.
    [void]$script:CloneTagVerified.Remove([string]$VmId)
    [void]$script:CloneMacRestored.Remove([string]$VmId)
}

# Called from Handle-GuestControl's delete branch so a clone deleted before it ever
# reaches IP-ready does not leave its clone-state entry (and concurrency slot) stuck
# forever, and so stale agent/IP caches for a destroyed VMID do not linger for a future
# guest that reuses the id.
function Clear-ProxmoxTrackingForVm {
    param(
        [Parameter(Mandatory = $true)]
        [string]$VmId
    )

    foreach ($key in @($script:TaskContext.Keys)) {
        $ctx = $script:TaskContext[$key]
        if ($null -ne $ctx -and $ctx.ContainsKey('clone_id') -and [string]$ctx.clone_id -eq [string]$VmId) {
            [void]$script:TaskContext.Remove($key)
            if ($script:CloneTaskCompletionCache.ContainsKey($key)) { [void]$script:CloneTaskCompletionCache.Remove($key) }
        }
    }

    Remove-CloneStateEntry -VmId $VmId

    if ($script:AgentFailureTracker.ContainsKey($VmId)) { $script:AgentFailureTracker.Remove($VmId) }
    if ($script:LastKnownNetworkData.ContainsKey($VmId)) { $script:LastKnownNetworkData.Remove($VmId) }
    if ($script:RecentlyControlledIds.ContainsKey($VmId)) { $script:RecentlyControlledIds.Remove($VmId) }
    if ($script:RecentlyStoppedIds.ContainsKey($VmId)) { $script:RecentlyStoppedIds.Remove($VmId) }
    if ($script:PendingTemplateTagRemovalIds.ContainsKey($VmId)) { $script:PendingTemplateTagRemovalIds.Remove($VmId) }
    if ($script:RecentlyConvertedIds.ContainsKey($VmId)) { $script:RecentlyConvertedIds.Remove($VmId) }
    if ($script:MaintenanceModeVmIds.Contains($VmId)) { [void]$script:MaintenanceModeVmIds.Remove($VmId) }
    if ($script:SweepLastCheckedAt.ContainsKey($VmId)) { $script:SweepLastCheckedAt.Remove($VmId) }
    if ($script:LastGuestPollAt.ContainsKey($VmId)) { $script:LastGuestPollAt.Remove($VmId) }
}

# See $script:RecentlyDeletedIds near the top of this file. Self-expires, so a reused
# VMID is never permanently hidden even if this process runs for a very long time
# between clones of the same id.
function Test-ProxmoxRecentlyDeleted {
    param([Parameter(Mandatory = $true)][string]$VmId)

    if (-not $script:RecentlyDeletedIds.ContainsKey($VmId)) {
        return $false
    }

    # Belt and braces for the id-reuse case Handle-GuestClone already retracts: an id
    # this provider is actively cloning into cannot also be a deleted VM, whatever a
    # leftover marker says. Handle-GuestClone is still the primary fix, because tracking
    # is dropped as soon as a clone reports ready while the marker can outlive it.
    if ($script:TrackedCloneVmIds.Contains($VmId)) {
        $script:RecentlyDeletedIds.Remove($VmId)
        $script:RecentlyDeletedNames.Remove($VmId)
        return $false
    }

    $deletedAt = [DateTime]$script:RecentlyDeletedIds[$VmId]
    if (([DateTime]::UtcNow - $deletedAt).TotalSeconds -ge $script:RecentlyDeletedRetentionSeconds) {
        $script:RecentlyDeletedIds.Remove($VmId)
        $script:RecentlyDeletedNames.Remove($VmId)
        return $false
    }

    return $true
}

# The recently-deleted stub's name -- see $script:RecentlyDeletedNames near the top of
# this file. Only ever called after Test-ProxmoxRecentlyDeleted has already confirmed the
# id is tracked, but falls back to the old synthetic placeholder regardless (a provider
# restart clears this map along with $script:RecentlyDeletedIds, and a missing/blank
# recorded name is itself defensive against a cluster entry that had none at delete time).
function Get-ProxmoxRecentlyDeletedName {
    param([Parameter(Mandatory = $true)][string]$VmId)

    if ($script:RecentlyDeletedNames.ContainsKey($VmId) -and
        -not [string]::IsNullOrWhiteSpace([string]$script:RecentlyDeletedNames[$VmId])) {
        return [string]$script:RecentlyDeletedNames[$VmId]
    }

    return "VM-$VmId"
}

# Self-expiring lookup for placement.preserve_node_on_recreation -- see
# $script:RecentlyDeletedNodeByName near the top of this file and
# Resolve-ProxmoxCloneTargetNode. Returns $null (never the source/current node, and never
# a synthetic fallback like Get-ProxmoxRecentlyDeletedName above) whenever there is
# nothing to preserve: no matching name, or the delete happened longer ago than
# cloning.recently_deleted_retention_seconds -- reusing that setting rather than adding a
# second timer for what is, in practice, the same signal (RAS recreates within seconds to
# low tens of seconds of the delete, well inside that window).
function Get-ProxmoxPreservedNodeForRecreation {
    param([Parameter(Mandatory = $true)][string]$Name)

    if (-not $script:RecentlyDeletedNodeByName.ContainsKey($Name)) {
        return $null
    }

    $entry = $script:RecentlyDeletedNodeByName[$Name]
    $deletedAt = [DateTime]$entry.deleted_at
    if (([DateTime]::UtcNow - $deletedAt).TotalSeconds -ge $script:RecentlyDeletedRetentionSeconds) {
        $script:RecentlyDeletedNodeByName.Remove($Name)
        return $null
    }

    $node = [string]$entry.node
    if ([string]::IsNullOrWhiteSpace($node)) {
        return $null
    }

    return $node
}

# See MAC-PRESERVATION.md and $script:RecentlyDeletedMacByName near the top of this
# file. Reads each netN interface's live config off Proxmox directly -- NOT the
# guest-agent-reported MAC (ConvertTo-RasGuestObject's mac_addresses field, ip/agent
# derived), which is unreliable exactly when this matters most: right as a VM is being
# deleted, its agent may already be unresponsive or the OS already shutting down.
# Proxmox's own config is authoritative and always available regardless of guest state.
function Get-ProxmoxVmNetIfaceMacs {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Node,
        [Parameter(Mandatory = $true)]
        [string]$VmId
    )

    $macs = [ordered]@{}
    $config = Get-ProxmoxVmConfig -Node $Node -VmId $VmId
    if ($null -eq $config) {
        return $macs
    }

    foreach ($name in (Get-MemberNames -Object $config)) {
        if ($name -notmatch '^net\d+$') {
            continue
        }
        $value = [string]$config.$name
        $m = [regex]::Match($value, '[0-9A-Fa-f]{2}(:[0-9A-Fa-f]{2}){5}')
        if ($m.Success) {
            $macs[$name] = $m.Value.ToUpperInvariant()
        }
    }

    return $macs
}

# Splices a MAC address into an existing netN config VALUE string (e.g.
# "virtio=BC:24:11:04:BE:D4,bridge=vmbr0,firewall=1"), replacing only the MAC-shaped
# token and leaving the model name, bridge, vlan, firewall flag, rate limit -- whatever
# else Proxmox already set for THIS clone -- completely untouched. Deliberately matches
# on the MAC's own shape rather than parsing "key=value,key=value" generically: robust
# to every NIC model name (virtio/e1000/rtl8139/vmxnet3/...) without enumerating them,
# and nothing else in a netN line is ever MAC-shaped.
function Set-ProxmoxNetIfaceMacInLine {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CurrentValue,
        [Parameter(Mandatory = $true)]
        [string]$NewMac
    )

    return [regex]::Replace($CurrentValue, '[0-9A-Fa-f]{2}(:[0-9A-Fa-f]{2}){5}', $NewMac, 1)
}

# Mirrors Get-ProxmoxPreservedNodeForRecreation exactly -- same name-keyed lookup, same
# recently_deleted_retention_seconds expiry window, same $null-means-nothing-to-preserve
# contract. See MAC-PRESERVATION.md.
function Get-ProxmoxPreservedMacsForRecreation {
    param([Parameter(Mandatory = $true)][string]$Name)

    if (-not $script:RecentlyDeletedMacByName.ContainsKey($Name)) {
        return $null
    }

    $entry = $script:RecentlyDeletedMacByName[$Name]
    $deletedAt = [DateTime]$entry.deleted_at
    if (([DateTime]::UtcNow - $deletedAt).TotalSeconds -ge $script:RecentlyDeletedRetentionSeconds) {
        $script:RecentlyDeletedMacByName.Remove($Name)
        return $null
    }

    $macs = $entry.macs
    if ($null -eq $macs -or @(Get-MemberNames -Object $macs).Count -eq 0) {
        return $null
    }

    return $macs
}

# See $script:RecentlyControlledIds near the top of this file. Self-expires the same way
# Test-ProxmoxRecentlyDeleted does.
function Test-ProxmoxRecentlyControlled {
    param([Parameter(Mandatory = $true)][string]$VmId)

    if (-not $script:RecentlyControlledIds.ContainsKey($VmId)) {
        return $false
    }

    $controlledAt = [DateTime]$script:RecentlyControlledIds[$VmId]
    if (([DateTime]::UtcNow - $controlledAt).TotalSeconds -ge $script:RecentlyControlledRetentionSeconds) {
        $script:RecentlyControlledIds.Remove($VmId)
        return $false
    }

    return $true
}

# See $script:RecentlyStoppedIds near the top of this file. Self-expires the same way
# Test-ProxmoxRecentlyDeleted does.
function Test-ProxmoxRecentlyStopped {
    param([Parameter(Mandatory = $true)][string]$VmId)

    if (-not $script:RecentlyStoppedIds.ContainsKey($VmId)) {
        return $false
    }

    $stoppedAt = [DateTime]$script:RecentlyStoppedIds[$VmId]
    if (([DateTime]::UtcNow - $stoppedAt).TotalSeconds -ge $script:RecentlyStoppedRetentionSeconds) {
        $script:RecentlyStoppedIds.Remove($VmId)
        return $false
    }

    return $true
}

# Records what this provider just asked Proxmox to make of a VM's template flag, so
# guests/get can report that instead of the lagging cluster listing.
#
# Also maintains the maintenance-mode set. RAS puts a template into maintenance by
# converting it to a plain VM and booting it, and takes it out by shutting it down (via
# the in-guest RAS agent, NOT via guests/control) and converting it back. That means the
# exit path has no control action to key off: RAS waits for THIS provider's guests/get
# to report powered_off before it issues the convert. A maintenance session lasts as
# long as the admin needs, so a time window is the wrong tool -- a VM stays in this set
# for as long as it is a de-templated RAS template.
function Set-ProxmoxRecentConvert {
    param(
        [Parameter(Mandatory = $true)][string]$VmId,
        [Parameter(Mandatory = $true)][bool]$IsTemplate
    )

    $script:RecentlyConvertedIds[$VmId] = @{
        is_template  = $IsTemplate
        converted_at = [DateTime]::UtcNow
    }

    if ($IsTemplate) {
        if ($script:MaintenanceModeVmIds.Contains($VmId)) { [void]$script:MaintenanceModeVmIds.Remove($VmId) }
    }
    else {
        [void]$script:MaintenanceModeVmIds.Add($VmId)
    }
}

# True while this provider has a template de-templated for RAS maintenance and
# has not converted it back yet. Power state and template flag for such a VM
# are always resolved live.
function Test-ProxmoxInMaintenanceMode {
    param([Parameter(Mandatory = $true)][string]$VmId)

    return $script:MaintenanceModeVmIds.Contains($VmId)
}

# Returns the recorded convert intent for $VmId, or $null once it has expired
# (or was never recorded). Self-expires the same way the other trackers do.
function Get-ProxmoxRecentConvertIntent {
    param([Parameter(Mandatory = $true)][string]$VmId)

    if (-not $script:RecentlyConvertedIds.ContainsKey($VmId)) {
        return $null
    }

    $entry = $script:RecentlyConvertedIds[$VmId]
    if (([DateTime]::UtcNow - [DateTime]$entry.converted_at).TotalSeconds -ge $script:RecentlyConvertedRetentionSeconds) {
        $script:RecentlyConvertedIds.Remove($VmId)
        return $null
    }

    return $entry
}

function Test-ProxmoxRecentlyConverted {
    param([Parameter(Mandatory = $true)][string]$VmId)

    return ($null -ne (Get-ProxmoxRecentConvertIntent -VmId $VmId))
}

# See $script:CompletedCloneTaskOutputs near the top of this file. Called by whichever
# code path (Get-RasGuestObjectForCloneAwareFlow or Handle-TaskInfo) FIRST confirms a
# clone is ready and clears its main tracking entry.
function Set-CompletedCloneTaskOutput {
    param(
        [string]$TaskId,
        [Parameter(Mandatory = $true)][string]$CloneId
    )

    if ([string]::IsNullOrWhiteSpace($TaskId)) {
        return
    }

    $script:CompletedCloneTaskOutputs[$TaskId] = @{ clone_id = $CloneId; completed_at = [DateTime]::UtcNow }
}

function Get-CompletedCloneTaskOutput {
    param([Parameter(Mandatory = $true)][string]$TaskId)

    if (-not $script:CompletedCloneTaskOutputs.ContainsKey($TaskId)) {
        return $null
    }

    $entry = $script:CompletedCloneTaskOutputs[$TaskId]
    if (([DateTime]::UtcNow - [DateTime]$entry.completed_at).TotalSeconds -ge $script:CompletedCloneTaskOutputRetentionSeconds) {
        $script:CompletedCloneTaskOutputs.Remove($TaskId)
        return $null
    }

    return $entry
}

# Stamps a tracked clone's persisted state entry with the moment its clone task was
# actually reported 'completed' to RAS (pipelined or not). Handle-GuestList uses this,
# not Proxmox's own transient name, to decide when a tracked clone is safe to list. A
# no-op if the entry is already gone.
function Set-CloneReportedCompleted {
    param([Parameter(Mandatory = $true)][string]$VmId)

    $entry = Get-CloneStateEntry -VmId $VmId
    if ($null -eq $entry) {
        return
    }

    $ctx = @{}
    foreach ($p in $entry.PSObject.Properties) { $ctx[$p.Name] = $p.Value }
    $ctx['reported_completed_at'] = [DateTime]::UtcNow.ToString('o')
    Set-CloneStateEntry -VmId $VmId -Entry $ctx
}

function Send-Response {
    param(
        [Parameter(Mandatory = $true)]
        [object]$ResponseObject
    )

    try {
        $json = $ResponseObject | ConvertTo-Json -Compress -Depth 20
        $writer.WriteLine($json)
        Write-DebugLog "OUT: $json" -Level 'D' -Component '01'
    }
    catch {
        $fallback = @{
            error = @{
                code    = $script:ErrorCodes.InternalError
                message = "$($script:ProviderNamePrefix) Failed to serialize response: $($_.Exception.Message)"
            }
        } | ConvertTo-Json -Compress -Depth 10

        # Guarded: if the ORIGINAL failure above was itself a broken stdout pipe
        # (RAS/parent gone), this second WriteLine throws too, escapes Send-Response,
        # escapes the main loop's own catch (which calls Send-Response), and kills the
        # process on an unhandled exception. A broken pipe is a legitimate reason to
        # exit, but it should be a deliberate, logged exit, not one reached by crashing.
        try {
            $writer.WriteLine($fallback)
            Write-DebugLog "OUT-FALLBACK: $fallback" -Level 'W' -Component '01'
        }
        catch {
            Write-DebugLog "stdout write failed on both the primary and fallback response ($($_.Exception.Message)) -- treating as a closed/broken pipe (parent gone) and exiting cleanly." -Level 'E' -Component '01'
            exit 0
        }
    }
}

function New-ErrorResponse {
    param(
        [int]$Code,
        [string]$Message
    )

    return @{
        error = @{
            code    = $Code
            message = $Message
        }
    }
}

function ConvertFrom-JsonSafe {
    param([string]$InputLine)

    try {
        return $InputLine | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        Write-DebugLog "JSON parse failed: $($_.Exception.Message)" -Level 'E' -Component '01'
        return $null
    }
}

function Test-RequiredFields {
    param(
        [object]$Data,
        [string[]]$RequiredFields
    )

    foreach ($field in $RequiredFields) {
        $keys = $field -split '\.'
        $value = $Data

        foreach ($key in $keys) {
            if ($null -ne $value -and (Get-MemberNames -Object $value) -contains $key) {
                $value = $value.$key
            }
            else {
                return "$($script:ProviderNamePrefix) Missing field: $field"
            }
        }
    }

    return $null
}

function Initialize-CertificateBypass {
    try {
        if ($PSVersionTable.PSEdition -eq 'Core') {
            return
        }

        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
        [System.Net.ServicePointManager]::SecurityProtocol = `
            [System.Net.SecurityProtocolType]::Tls12 -bor `
            [System.Net.SecurityProtocolType]::Tls11 -bor `
            [System.Net.SecurityProtocolType]::Tls
    }
    catch {
        Write-DebugLog "Certificate bypass init failed: $($_.Exception.Message)" -Level 'E' -Component '02'
    }
}

function Get-Session {
    if ($null -eq $script:ProxmoxSession) {
        throw 'Session not initialized'
    }

    if ([string]::IsNullOrWhiteSpace($script:ProxmoxSession.host)) {
        throw 'Session host missing'
    }

    if ($null -eq $script:ProxmoxSession.header) {
        throw 'Session header missing'
    }

    return $script:ProxmoxSession
}

function Get-ProxmoxBaseUrl {
    param([hashtable]$Session)
    return ("https://{0}" -f $Session.host).TrimEnd('/')
}

function Invoke-ProxmoxRestMethod {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri,

        [Parameter(Mandatory = $true)]
        [hashtable]$Headers,

        [Parameter(Mandatory = $true)]
        [ValidateSet('GET', 'POST', 'PUT', 'DELETE')]
        [string]$Method,

        [object]$Body = $null,

        # 0 (default) means "use cloning.timeouts.http_timeout_seconds" -- every call
        # gets SOME bound (rather than PowerShell's own unbounded HttpClient default).
        # Set explicitly to give a call its own tighter
        # ceiling: guest_agent.timeout_seconds (Get-ProxmoxVmGuestAgentInterfaces) and
        # tag_write_timeout_seconds (Set-ProxmoxVmTagList) both do.
        [int]$TimeoutSec = 0,

        # 'E' (default) is right for a real Proxmox API call -- this is the only place
        # such a failure gets logged at all (Invoke-ProxmoxApi/-WithRetry add no log line
        # of their own on a non-retryable/exhausted failure). The guest-agent probe is
        # the one caller that overrides this to 'T': its own short guest_agent.timeout_seconds
        # ceiling routinely elapses while a guest is still booting, which is expected
        # during clone preparation, not a real problem -- and it already logs its own,
        # more specific 'T' line one level up (Get-ProxmoxVmGuestAgentInterfaces), so
        # logging 'E' here too would both misclassify a routine timeout as an error and
        # double-log the same failure.
        [ValidateSet('E', 'W', 'I', 'T', 'D')]
        [string]$FailureLevel = 'E',

        # A network blip (a stale pooled connection, a dropped packet, a momentary
        # routing hiccup) gets exactly one retry on a DELIBERATELY FRESH connection
        # before the failure reaches the caller. This provider's stdin loop is
        # single-threaded, so a call that just hangs blocks every other request behind
        # it -- bounding + retrying + then failing fast beats hanging indefinitely or
        # failing on the very first blip. Opt out with -NoRetry for a
        # call where retrying doesn't address the actual failure mode: the guest-agent
        # probe (a fresh connection doesn't make an unbooted guest's agent answer any
        # sooner -- it's not a network problem) and tag writes (best-effort, must never
        # dominate the request they ride along on -- doubling their worst case
        # contradicts that).
        [switch]$NoRetry
    )

    $effectiveTimeoutSec = if ($TimeoutSec -gt 0) { $TimeoutSec } else { $script:HttpTimeoutSeconds }
    $maxAttempts = if ($NoRetry) { 1 } else { 2 }
    $attempt = 0

    while ($true) {
        $attempt++

        $irmParams = @{
            Uri         = $Uri
            Headers     = $Headers
            Method      = $Method
            ErrorAction = 'Stop'
            TimeoutSec  = $effectiveTimeoutSec
        }

        if ($PSVersionTable.PSEdition -eq 'Core') {
            $irmParams.SkipCertificateCheck = $true
            $irmParams.SkipHeaderValidation = $true
        }

        if ($null -ne $Body) {
            $irmParams.Body = $Body
        }

        # Without a -WebSession, Invoke-RestMethod opens a fresh TCP connection (and,
        # over HTTPS, a fresh TLS handshake) per call. Reuse is safe here only because
        # certificate validation is constant across every call (see above) -- toggling
        # it per call would silently abandon the pooled connection instead. Re-evaluated
        # every loop iteration on purpose: a retry (below) clears
        # $script:ProxmoxWebSession specifically so this branch takes the fresh-session
        # path instead of resending down whatever connection just failed.
        if ($null -ne $script:ProxmoxWebSession) {
            $irmParams.WebSession = $script:ProxmoxWebSession
        }
        else {
            $irmParams.SessionVariable = 'proxmoxWebSessionCapture'
        }

        Write-DebugLog "HTTP $Method $Uri (attempt $attempt/$maxAttempts, timeout ${effectiveTimeoutSec}s)" -Level 'D' -Component '0B'

        try {
            $result = Invoke-RestMethod @irmParams

            if ($null -eq $script:ProxmoxWebSession) {
                $script:ProxmoxWebSession = $proxmoxWebSessionCapture
            }

            return $result
        }
        catch {
            if ($attempt -lt $maxAttempts) {
                Write-DebugLog "HTTP $Method $Uri failed (attempt $attempt/$maxAttempts): $($_.Exception.Message) -- retrying once on a fresh connection." -Level 'W' -Component '0B'
                $script:ProxmoxWebSession = $null
                continue
            }

            Write-DebugLog "HTTP failure: $($_.Exception.Message)" -Level $FailureLevel -Component '0B'
            throw
        }
    }
}

function Invoke-ProxmoxApi {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('GET', 'POST', 'DELETE', 'PUT')]
        [string]$Method,

        [Parameter(Mandatory = $true)]
        [string]$Path,

        [hashtable]$Body,

        [int]$TimeoutSec = 0,

        # Passed straight through to Invoke-ProxmoxRestMethod -- see its own comment on
        # this parameter for why the guest-agent probe is the one caller that overrides it.
        [ValidateSet('E', 'W', 'I', 'T', 'D')]
        [string]$FailureLevel = 'E',

        # Passed straight through to Invoke-ProxmoxRestMethod -- see its own comment.
        [switch]$NoRetry
    )

    $session = Get-Session
    $base = Get-ProxmoxBaseUrl -Session $session
    $uri = ($base.TrimEnd('/') + '/' + $Path.TrimStart('/'))

    if ($Method -eq 'GET') {
        return Invoke-ProxmoxRestMethod -Uri $uri -Headers $session.header -Method GET -TimeoutSec $TimeoutSec -FailureLevel $FailureLevel -NoRetry:$NoRetry
    }

    if ($Method -eq 'DELETE') {
        return Invoke-ProxmoxRestMethod -Uri $uri -Headers $session.header -Method DELETE -TimeoutSec $TimeoutSec -FailureLevel $FailureLevel -NoRetry:$NoRetry
    }

    if ($null -eq $Body) {
        $Body = @{}
    }

    return Invoke-ProxmoxRestMethod -Uri $uri -Headers $session.header -Method $Method -Body $Body -TimeoutSec $TimeoutSec -FailureLevel $FailureLevel -NoRetry:$NoRetry
}

# guests/control (start/stop/reset/reboot/delete/suspend/resume) is the only caller. A
# burst of control POSTs can make Proxmox fail to fork a pvedaemon worker ("got no
# worker upid - start worker failed") -- a transient failure, not a real VM-state
# problem, and RAS has no retry of its own here. Retries only on a transient 5xx
# (matched on the status line itself, so a genuine timeout or a 4xx still fails
# immediately) and only up to cloning.control_retry_max_attempts.
function Invoke-ProxmoxApiWithRetry {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('GET', 'POST', 'DELETE', 'PUT')]
        [string]$Method,

        [Parameter(Mandatory = $true)]
        [string]$Path,

        [hashtable]$Body,

        [int]$TimeoutSec = 0
    )

    $maxAttempts = [Math]::Max(1, $script:ControlRetryMaxAttempts)
    $attempt = 0

    while ($true) {
        $attempt++
        try {
            return Invoke-ProxmoxApi -Method $Method -Path $Path -Body $Body -TimeoutSec $TimeoutSec
        }
        catch {
            $msg = $_.Exception.Message
            $isTransient = ($msg -match '\b5\d\d\b') -and ($attempt -lt $maxAttempts)

            if (-not $isTransient) {
                throw
            }

            Write-DebugLog "Transient Proxmox failure on $Method $Path (attempt $attempt/$maxAttempts): $msg -- retrying in $($script:ControlRetryDelaySeconds)s." -Level 'W' -Component '0B'
            Start-Sleep -Seconds $script:ControlRetryDelaySeconds
        }
    }
}

function Get-ProxmoxClusterVMs {
    if ($null -ne $script:ClusterResourcesCache -and $null -ne $script:ClusterResourcesCachedAt -and
        ([DateTime]::UtcNow - $script:ClusterResourcesCachedAt).TotalSeconds -lt $script:ClusterResourcesCacheTtlSeconds) {
        return $script:ClusterResourcesCache
    }

    $resp = Invoke-ProxmoxApi -Method GET -Path '/api2/json/cluster/resources?type=vm'
    # Filter to QEMU VMs only — LXC containers use different API endpoints
    $vms = @($resp.data | Where-Object { [string]$_.type -eq 'qemu' })

    # Busting the cache at guests/clone time (see Reset-ProxmoxClusterCache) closes the
    # race AT THAT MOMENT, but a subsequent poll can still land in the window before
    # Proxmox has propagated the new VM into cluster/resources -- and caching THAT
    # snapshot poisons every guests/get for up to the full TTL. So: never cache a
    # listing that is missing a VMID this provider is actively tracking as an in-flight
    # clone; the next call is forced to re-fetch instead. Checked against the FULL,
    # not-yet-pool-filtered list below -- a tracked clone briefly outside pool_scope (say,
    # a live settings change mid-flight) is a genuine out-of-scope state, not a
    # propagation race, and must not trigger an endless re-fetch loop.
    $missingTrackedClone = $false
    if ($script:TrackedCloneVmIds.Count -gt 0) {
        $seenIds = [System.Collections.Generic.HashSet[string]]::new([string[]]($vms | ForEach-Object { [string]$_.vmid }))
        foreach ($trackedId in $script:TrackedCloneVmIds) {
            if (-not $seenIds.Contains($trackedId)) {
                $missingTrackedClone = $true
                break
            }
        }
    }

    # pool_scope.pool_name -- see POOL-SCOPING.md. Filtered HERE, ONCE, right after the
    # single cluster/resources fetch, instead of separately by every caller (was:
    # Get-ProxmoxVmNode, Handle-HostList and Handle-GuestList each re-ran the same
    # per-VM pool check and re-logged the same skip, every single call). An out-of-scope
    # VM is simply never in the list any caller sees from here on -- "not found" falls
    # out naturally wherever a caller already handles a genuinely absent VM, with no
    # separate pool_scope branch needed downstream.
    $inScopeVms = if ([string]::IsNullOrWhiteSpace($script:PoolScopeName)) {
        $vms
    }
    else {
        @($vms | Where-Object {
            $inScope = Test-ProxmoxVmInPoolScope -ClusterVm $_
            if (-not $inScope) {
                Write-DebugLog "Excluding guest [$($_.vmid)] from this provider's view: outside pool_scope [$script:PoolScopeName] (its own pool: [$(Get-ProxmoxVmPool -ClusterVm $_)])." -Level 'T' -Component '03' -Ref ([string]$_.vmid)
            }
            $inScope
        })
    }

    if ($missingTrackedClone) {
        $script:ClusterResourcesCache = $null
        $script:ClusterResourcesCachedAt = $null
    }
    else {
        $script:ClusterResourcesCache = $inScopeVms
        $script:ClusterResourcesCachedAt = [DateTime]::UtcNow
    }

    return $inScopeVms
}

# GET /cluster/resources?type=node -- a separate call and cache from Get-ProxmoxClusterVMs
# above (?type=vm): feeds Resolve-ProxmoxCloneTargetNode's placement decision, not guest
# state, so it has its own TTL (placement.node_stats_cache_ttl_seconds) tuned for a
# scheduling decision rather than a display value. Only ever called when
# placement.enabled is true.
function Get-ProxmoxClusterNodes {
    if ($null -ne $script:ClusterNodesCache -and $null -ne $script:ClusterNodesCachedAt -and
        ([DateTime]::UtcNow - $script:ClusterNodesCachedAt).TotalSeconds -lt $script:Settings.cloning.load_balancing.node_stats_cache_ttl_seconds) {
        return $script:ClusterNodesCache
    }

    $resp = Invoke-ProxmoxApi -Method GET -Path '/api2/json/cluster/resources?type=node'
    $nodes = @($resp.data | Where-Object { [string]$_.type -eq 'node' })

    $script:ClusterNodesCache = $nodes
    $script:ClusterNodesCachedAt = [DateTime]::UtcNow

    return $nodes
}

# Decides which node a new clone's compute should land on -- see the 'placement' block in
# Get-DefaultProviderSettings for the settings this reads. Returns $null whenever the
# caller should fall back to today's behavior (no 'target' in the clone body, landing on
# the source's own node): placement disabled, no eligible node, or any failure talking to
# Proxmox. Placement is a scheduling nicety, never a reason to fail a clone -- every
# failure path here is caught and logged, not thrown, matching this script's other
# best-effort decisions (tag writes, orphan audit).
function Resolve-ProxmoxCloneTargetNode {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceNode,

        # The clone's intended name (RAS's own guests/clone 'name' param) -- the only
        # correlation key available for "is this a recreation of a guest we already
        # placed once", since a fresh clone request never carries the OLD deleted VM's
        # id. See placement.preserve_node_on_recreation.
        [string]$CloneName
    )

    if (-not $script:Settings.cloning.load_balancing.enabled) {
        return $null
    }

    try {
        $nodes = @(Get-ProxmoxClusterNodes)
    }
    catch {
        Write-DebugLog "Placement: failed to read cluster node list, falling back to the source node [$SourceNode]: $($_.Exception.Message)" -Level 'W' -Component '09'
        return $null
    }

    $excludedList = @($script:Settings.cloning.load_balancing.excluded_nodes)
    $excluded = [System.Collections.Generic.HashSet[string]]::new([string[]]$excludedList, [System.StringComparer]::Ordinal)

    # Three independent reasons a node is never a placement target -- unavailable,
    # admin-excluded, or drained for HA maintenance. 'status' covers the first: Proxmox
    # sets it to anything other than 'online' when the node itself is down/unreachable.
    # It does NOT cover the third: a node put into HA maintenance
    # (ha-manager crm-command node-maintenance enable) can still report status=online --
    # it is up and reachable, just intentionally being drained -- and that state shows up
    # only in the separate 'hastate' field instead. A node with no 'hastate' at all (HA
    # not configured/running) is never excluded on that basis alone.
    $ineligibleReasons = [ordered]@{}
    $eligible = @($nodes | Where-Object {
            $n = $_
            $nodeName = [string]$n.node
            $reasons = [System.Collections.Generic.List[string]]::new()

            if ([string]$n.status -ne 'online') { $reasons.Add("status=$([string]$n.status)") }

            $hastate = if ((Get-MemberNames -Object $n) -contains 'hastate') { [string]$n.hastate } else { '' }
            if (-not [string]::IsNullOrWhiteSpace($hastate) -and $hastate -eq 'maintenance') { $reasons.Add('hastate=maintenance') }

            if ($excluded.Contains($nodeName)) { $reasons.Add('excluded_nodes') }

            if ($reasons.Count -gt 0) {
                $ineligibleReasons[$nodeName] = ($reasons -join ',')
                return $false
            }
            return $true
        })

    if ($eligible.Count -eq 0) {
        $reasonSummary = ($ineligibleReasons.Keys | ForEach-Object { "$_($($ineligibleReasons[$_]))" }) -join ', '
        Write-DebugLog "Placement: no eligible node found -- [$reasonSummary] -- falling back to the source node [$SourceNode]." -Level 'W' -Component '09'
        return $null
    }

    # Recreation: RAS deleted a guest by this exact name and is now cloning a new VM
    # under the same name -- land it back where it was, skipping strategy selection
    # entirely, rather than treating it as a fresh placement decision. Only applies
    # within cloning.recently_deleted_retention_seconds of the delete (see
    # Get-ProxmoxPreservedNodeForRecreation) and only if that node is still eligible --
    # a node that went offline or was newly excluded since the delete falls through to
    # strategy-based selection below instead of being forced onto a bad node.
    if ($script:Settings.cloning.load_balancing.preserve_node_on_recreation -and -not [string]::IsNullOrWhiteSpace($CloneName)) {
        $preservedNode = Get-ProxmoxPreservedNodeForRecreation -Name $CloneName
        if (-not [string]::IsNullOrWhiteSpace($preservedNode)) {
            $stillEligible = @($eligible | Where-Object { [string]$_.node -eq $preservedNode }).Count -gt 0
            if ($stillEligible) {
                Write-DebugLog "Placement: [$CloneName] looks like a recreation -- preserving its prior node [$preservedNode] instead of re-selecting (placement.preserve_node_on_recreation)." -Level 'I' -Component '09'
                return $preservedNode
            }
            Write-DebugLog "Placement: [$CloneName]'s prior node [$preservedNode] is no longer eligible (offline or newly excluded) -- falling through to strategy-based selection." -Level 'W' -Component '09'
        }
    }

    $strategy = ([string]$script:Settings.cloning.load_balancing.strategy).Trim().ToLowerInvariant()

    if ($strategy -eq 'round_robin') {
        # Sort by name first so the rotation order is stable call to call, regardless of
        # what order Proxmox happens to return nodes in.
        $ordered = @($eligible | Sort-Object -Property node)
        $index = $script:PlacementRoundRobinIndex % $ordered.Count
        $script:PlacementRoundRobinIndex = ($script:PlacementRoundRobinIndex + 1) % $ordered.Count
        $chosen = [string]$ordered[$index].node
        Write-DebugLog "Placement: round_robin chose [$chosen] from [$(($ordered | ForEach-Object { [string]$_.node }) -join ', ')]." -Level 'T' -Component '09'
        return $chosen
    }

    # Default: 'resource'. Lower score = more headroom = better node. 'cpu' (Proxmox's
    # own normalized load fraction) and mem/maxmem (a plain used-fraction) are both
    # already 0..1 and comparable across nodes regardless of size, so averaging them for
    # 'both' needs no extra weighting -- see the node/mem/cpu sample in
    # DISTRIBUTED-PLACEMENT.md.
    $metric = ([string]$script:Settings.cloning.load_balancing.resource_metric).Trim().ToLowerInvariant()
    if ($metric -ne 'cpu' -and $metric -ne 'ram' -and $metric -ne 'both') {
        Write-DebugLog "Placement: unrecognized resource_metric [$($script:Settings.cloning.load_balancing.resource_metric)] -- defaulting to 'both'." -Level 'W' -Component '09'
        $metric = 'both'
    }

    $scored = @($eligible | ForEach-Object {
            $n = $_
            $cpuFrac = 0.0
            try { $cpuFrac = [double]$n.cpu } catch { }

            $ramFrac = 0.0
            try {
                $maxMem = [double]$n.maxmem
                if ($maxMem -gt 0) { $ramFrac = [double]$n.mem / $maxMem }
            }
            catch { }

            $score = switch ($metric) {
                'cpu' { $cpuFrac }
                'ram' { $ramFrac }
                default { ($cpuFrac + $ramFrac) / 2.0 }
            }

            [PSCustomObject]@{ Node = [string]$n.node; Score = $score }
        })

    $best = $scored | Sort-Object -Property Score | Select-Object -First 1
    $scoreSummary = ($scored | ForEach-Object { "$($_.Node)=$([Math]::Round($_.Score, 3))" }) -join ', '
    Write-DebugLog "Placement: resource ($metric) scores [$scoreSummary] -- chose [$($best.Node)]." -Level 'T' -Component '09'
    return $best.Node
}

# Proxmox stores tags as a single ';'-separated string on the VM's own config
# (visible in cluster/resources as the 'tags' field, no extra call needed).
# Also tolerates ',' for defensiveness against hand-edited tag strings.
function Get-ProxmoxVmTagList {
    param([object]$ClusterVm)

    if ($null -eq $ClusterVm -or
        -not ((Get-MemberNames -Object $ClusterVm) -contains 'tags') -or
        [string]::IsNullOrWhiteSpace([string]$ClusterVm.tags)) {
        return @()
    }

    return @(([string]$ClusterVm.tags) -split '[;,]' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
}

function Get-ProxmoxVmConfig {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Node,

        [Parameter(Mandatory = $true)]
        [string]$VmId
    )

    $resp = Invoke-ProxmoxApi -Method GET -Path "/api2/json/nodes/$Node/qemu/$VmId/config"
    return $resp.data
}

function Test-ProxmoxVmHasTag {
    param(
        [object]$ClusterVm,
        [Parameter(Mandatory = $true)]
        [string]$Tag
    )

    return (Get-ProxmoxVmTagList -ClusterVm $ClusterVm) -contains $Tag
}

# rasClone<sourceId>/rasTemplate<sourceId> both carry a variable numeric suffix, so this
# is a prefix match against the cached tag list -- no extra HTTP call, same reasoning as
# Test-ProxmoxVmHasTag. Used by Get-ProxmoxVmGuestAgentInterfaces to recognize a
# RAS-managed clone or template for the agent-quarantine exemption.
function Test-ProxmoxVmHasTagPrefix {
    param(
        [object]$ClusterVm,
        [Parameter(Mandatory = $true)]
        [string]$Prefix
    )

    return @(Get-ProxmoxVmTagList -ClusterVm $ClusterVm | Where-Object { $_ -like "$Prefix*" }).Count -gt 0
}

# See POOL-SCOPING.md. cluster/resources carries 'pool' directly on each VM entry when
# it belongs to one -- no extra HTTP call, same reasoning as Get-ProxmoxVmTagList.
# Note: not confirmed against every Proxmox VE version that this field is always
# present for every pooled VM -- see POOL-SCOPING.md. A VM whose pool membership
# isn't visible for any reason falls back to treated-as-unpooled below, never
# silently misclassified.
function Get-ProxmoxVmPool {
    param([object]$ClusterVm)

    if ($null -eq $ClusterVm -or
        -not ((Get-MemberNames -Object $ClusterVm) -contains 'pool') -or
        [string]::IsNullOrWhiteSpace([string]$ClusterVm.pool)) {
        return ''
    }

    return [string]$ClusterVm.pool
}

# virtual_machines.pool_scope.pool_name empty (the default) means no filtering at all --
# every VM in scope regardless of pool, including unpooled ones. Set, a VM is in scope
# only if its own pool matches exactly (case-sensitive, like every other Proxmox id this
# provider compares).
function Test-ProxmoxVmInPoolScope {
    param([object]$ClusterVm)

    if ([string]::IsNullOrWhiteSpace($script:PoolScopeName)) {
        return $true
    }

    return (Get-ProxmoxVmPool -ClusterVm $ClusterVm) -eq $script:PoolScopeName
}

# Any read-modify-write against a VM's tag string needs the LIVE value, not the cached
# cluster/resources listing ($FallbackClusterVm): a just-created clone's cached entry
# often has no 'tags' property at all yet, and an admin's hand-set tag may not have
# propagated to this provider's cache either. A stale or empty read means the
# full-replace PUT silently deletes every tag that is not ours. Falls back to the cached
# listing only if the live read itself fails -- best-effort, matching every other tag
# write in this script.
function Get-ProxmoxVmLiveTagList {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Node,

        [Parameter(Mandatory = $true)]
        [string]$VmId,

        [object]$FallbackClusterVm = $null
    )

    try {
        $config = Get-ProxmoxVmConfig -Node $Node -VmId $VmId
        if ($null -ne $config -and (Get-MemberNames -Object $config) -contains 'tags' -and
            -not [string]::IsNullOrWhiteSpace([string]$config.tags)) {
            return @(([string]$config.tags) -split '[;,]' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        }
        return @()
    }
    catch {
        Write-DebugLog "Get-ProxmoxVmLiveTagList: failed to read live config for VM [$VmId], falling back to the cached tag list: $($_.Exception.Message)" -Level 'W' -Component '06' -Ref $VmId
        return Get-ProxmoxVmTagList -ClusterVm $FallbackClusterVm
    }
}

# Read-modify-write against the VM's own tag string. Proxmox VE tags are
# meta-information only (no locking semantics of their own), so this is safe to call
# even while a clone/start/stop job involving this VM is in flight.
function Set-ProxmoxVmTagList {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Node,

        [Parameter(Mandatory = $true)]
        [string]$VmId,

        [string[]]$Tags,

        # Best-effort metadata write, never allowed to dominate the request it rides
        # along on. Defaults to the shared constant rather than 0 (unbounded) so every
        # call site gets a bound unless it deliberately opts out.
        [int]$TimeoutSec = $script:TagWriteTimeoutSeconds
    )

    $joined = ($Tags | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join ';'
    [void](Invoke-ProxmoxApi -Method PUT -Path "/api2/json/nodes/$Node/qemu/$VmId/config" -Body @{ tags = $joined } -TimeoutSec $TimeoutSec -NoRetry)
}

function Add-ProxmoxVmTag {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Node,

        [Parameter(Mandatory = $true)]
        [string]$VmId,

        [object]$ClusterVm,

        [Parameter(Mandatory = $true)]
        [string]$Tag,

        [int]$TimeoutSec = $script:TagWriteTimeoutSeconds
    )

    $current = Get-ProxmoxVmLiveTagList -Node $Node -VmId $VmId -FallbackClusterVm $ClusterVm
    if ($current -contains $Tag) {
        return
    }

    Set-ProxmoxVmTagList -Node $Node -VmId $VmId -Tags (@($current) + $Tag) -TimeoutSec $TimeoutSec
}

function Remove-ProxmoxVmTag {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Node,

        [Parameter(Mandatory = $true)]
        [string]$VmId,

        [object]$ClusterVm,

        [Parameter(Mandatory = $true)]
        [string]$Tag,

        [int]$TimeoutSec = $script:TagWriteTimeoutSeconds
    )

    $current = Get-ProxmoxVmLiveTagList -Node $Node -VmId $VmId -FallbackClusterVm $ClusterVm
    if ($current -notcontains $Tag) {
        return
    }

    Set-ProxmoxVmTagList -Node $Node -VmId $VmId -Tags (@($current | Where-Object { $_ -ne $Tag })) -TimeoutSec $TimeoutSec
}

# Keeps rasTemplate<VmId> in sync with whether THIS guest is presently a RAS-managed
# template, from the two places that actually know that with certainty:
#   - Handle-GuestConvert adds it on every successful convert TO a template (is_template
#     true) -- unambiguous, nothing else in the RAS contract calls convert(true).
#   - Handle-GuestSnapshotsDelete removes it -- see the comment there for why that call,
#     not guests/convert(is_template:false), is the reliable "RAS is deleting the
#     Template object" signal.
# Best-effort and non-blocking, matching every other tag write in this script -- a failed
# write here never fails the caller.
function Set-ProxmoxTemplateSourceTagBestEffort {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Node,

        [Parameter(Mandatory = $true)]
        [string]$VmId,

        [object]$ClusterVm,

        [Parameter(Mandatory = $true)]
        [bool]$Present
    )

    $tag = "$($script:RasTemplateTagPrefix)$VmId"
    try {
        if ($Present) {
            Add-ProxmoxVmTag -Node $Node -VmId $VmId -ClusterVm $ClusterVm -Tag $tag
            Write-DebugLog "Tagged VM [$VmId] with [$tag] (converted to template)." -Level 'I' -Component '06' -Ref $VmId
        }
        else {
            Remove-ProxmoxVmTag -Node $Node -VmId $VmId -ClusterVm $ClusterVm -Tag $tag
            Write-DebugLog "Removed [$tag] from VM [$VmId] (RAS deleted the Template object)." -Level 'I' -Component '06' -Ref $VmId
        }
    }
    catch {
        Write-DebugLog "Failed to $(if ($Present) { 'add' } else { 'remove' }) tag [$tag] on VM [$VmId]: $($_.Exception.Message)" -Level 'E' -Component '06' -Ref $VmId
    }
}

# See $script:PendingTemplateTagRemovalIds near the top of this file for the reasoning.
# Called opportunistically from ConvertTo-RasGuestObject for whichever vmid RAS just
# asked about -- costs nothing when there is no pending marker for that id (the common
# case), and only reaches Proxmox once the grace window has actually elapsed.
function Resolve-PendingTemplateTagRemoval {
    param(
        [Parameter(Mandatory = $true)]
        [string]$VmId,

        [Parameter(Mandatory = $true)]
        [string]$Node,

        [object]$ClusterVm
    )

    if (-not $script:PendingTemplateTagRemovalIds.ContainsKey($VmId)) {
        return
    }

    $deTemplatedAt = [DateTime]$script:PendingTemplateTagRemovalIds[$VmId]
    if (([DateTime]::UtcNow - $deTemplatedAt).TotalSeconds -lt $script:TemplateDeleteConfirmSeconds) {
        return
    }

    Write-DebugLog "VM [$VmId] de-templated $($script:TemplateDeleteConfirmSeconds)s ago with no guests/control start since -- treating as RAS deleting the Template object, not maintenance." -Level 'I' -Component '06' -Ref $VmId
    $script:PendingTemplateTagRemovalIds.Remove($VmId)
    Set-ProxmoxTemplateSourceTagBestEffort -Node $Node -VmId $VmId -ClusterVm $ClusterVm -Present $false
}

# Proxmox's clone operation copies the source VM's config -- tags included -- to the new
# VM, and this script tags the SOURCE with rasTemplate<sourceId> right before submitting
# the clone POST (see Handle-GuestClone, placed there specifically to dodge Proxmox's
# clone-time config lock). So every clone taken from that source afterward inherits the
# tag at creation time, before this provider ever touches the clone's own tags. A clone
# is never itself a RAS template until explicitly converted via guests/convert, so strip
# any inherited rasTemplate* tag here, in the same PUT that applies rasClone<sourceId>
# -- one read-modify-write, not two that could race each other.
function Repair-ProxmoxCloneTagsAfterInherit {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Node,

        [Parameter(Mandatory = $true)]
        [string]$VmId,

        [object]$ClusterVm,

        [Parameter(Mandatory = $true)]
        [string]$CloneTag,

        [int]$TimeoutSec = $script:TagWriteTimeoutSeconds
    )

    # See Get-ProxmoxVmLiveTagList -- $ClusterVm (the cached cluster/resources entry) is
    # not trustworthy for a read-modify-write against tags; a JUST-cloned VM in
    # particular often has no 'tags' property at all yet, and using the cache here drops
    # every non-RAS tag the VM had.
    $current = Get-ProxmoxVmLiveTagList -Node $Node -VmId $VmId -FallbackClusterVm $ClusterVm
    $inheritedTemplateTags = @($current | Where-Object { $_ -like "$($script:RasTemplateTagPrefix)*" })
    $alreadyHasCloneTag = $current -contains $CloneTag

    if (-not $alreadyHasCloneTag -or $inheritedTemplateTags.Count -gt 0) {
        $desired = @($current | Where-Object { $_ -notlike "$($script:RasTemplateTagPrefix)*" })
        if (-not $alreadyHasCloneTag) {
            $desired = @($desired) + $CloneTag
        }

        Set-ProxmoxVmTagList -Node $Node -VmId $VmId -Tags $desired -TimeoutSec $TimeoutSec
        return @{ changed = $true; removed_inherited_template_tags = $inheritedTemplateTags }
    }

    return @{ changed = $false; removed_inherited_template_tags = @() }
}

function Get-ProxmoxVmNode {
    param(
        [Parameter(Mandatory = $true)]
        [string]$VmId
    )

    # Get-ProxmoxClusterVMs already excludes anything outside pool_scope.pool_name (see
    # POOL-SCOPING.md and that function's own comment) -- an out-of-scope VM is simply
    # never in $clusterVMs, so it falls into the same "not found" branch below as a
    # genuinely absent one, with no separate check needed here.
    $clusterVMs = Get-ProxmoxClusterVMs
    $vm = $clusterVMs | Where-Object { [string]$_.vmid -eq [string]$VmId } | Select-Object -First 1

    if ($null -eq $vm) {
        throw "VM [$VmId] not found in cluster"
    }

    # rasExclude ("this VM does not exist for RAS") is enforced HERE, centrally, rather
    # than in every individual handler: Get-ProxmoxVmNode is the one low-level lookup
    # used by guests/get, guests/control, guests/clone's source lookup and the
    # clone-aware flow, so reusing the exact same "not found" error a genuinely absent
    # VM produces makes it uniformly invisible with no extra wiring.
    # Handle-GuestList/Handle-HostList must filter it directly as well, since they
    # enumerate the cluster listing rather than going through this function.
    if (Test-ProxmoxVmHasTag -ClusterVm $vm -Tag $script:RasExcludeTag) {
        throw "VM [$VmId] not found in cluster"
    }

    if ([string]::IsNullOrWhiteSpace($vm.node)) {
        throw "VM [$VmId] has no node information"
    }

    return $vm
}

function Map-ProxmoxStateToRasState {
    param([string]$State)

    $normalized = if ($null -ne $State) { $State.ToString().Trim().ToLowerInvariant() } else { 'unknown' }

    switch ($normalized) {
        'running' { return 'powered_on' }
        'stopped' { return 'powered_off' }
        'paused' { return 'suspended' }
        'suspended' { return 'suspended' }
        'shutdown' { return 'powering_off' }
        'halting' { return 'powering_off' }
        'prelaunch' { return 'powering_on' }
        default { return 'powered_off' }
    }
}

function Get-ProxmoxVmCurrentStatus {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Node,

        [Parameter(Mandatory = $true)]
        [string]$VmId
    )

    $resp = Invoke-ProxmoxApi -Method GET -Path "/api2/json/nodes/$Node/qemu/$VmId/status/current"
    return $resp.data
}

function Get-ProxmoxVmGuestAgentInterfaces {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Node,

        [Parameter(Mandatory = $true)]
        [string]$VmId,

        # The VM's cluster/resources entry -- carries the current tag string, so the
        # quarantine fast path below costs no extra HTTP call. Optional so this function
        # still works without it (falls back to the count/time-based behaviour).
        [object]$ClusterVm = $null
    )

    # Quarantine fast path. If this VM is ALREADY tagged rasQuarantine, skip the agent
    # call outright -- the tag is the whole state, so this survives a provider restart,
    # and an admin removing it resumes checking on the next poll.
    if (Test-ProxmoxVmHasTag -ClusterVm $ClusterVm -Tag $script:RasQuarantineTag) {
        return @()
    }

    try {
        $resp = Invoke-ProxmoxApi -Method GET -Path "/api2/json/nodes/$Node/qemu/$VmId/agent/network-get-interfaces" -TimeoutSec $script:GuestAgentTimeoutSeconds -FailureLevel 'T' -NoRetry

        if ($script:AgentFailureTracker.ContainsKey($VmId)) {
            $script:AgentFailureTracker.Remove($VmId)
        }
        [void]$script:RasManagedAgentExemptionLogged.Remove($VmId)

        return @($(if ($null -ne $resp.data.result) { $resp.data.result } else { $resp.data }))
    }
    catch {
        Write-DebugLog "Guest agent interface query failed for VM [$VmId]: $($_.Exception.Message)" -Level 'T' -Component '03' -Ref $VmId

        $tracker = $script:AgentFailureTracker[$VmId]
        $failureCount = if ($null -ne $tracker) { [int]$tracker.count + 1 } else { 1 }
        $firstFailureAt = if ($null -ne $tracker) { [DateTime]$tracker.first_failure_at } else { [DateTime]::UtcNow }
        $script:AgentFailureTracker[$VmId] = @{ count = $failureCount; first_failure_at = $firstFailureAt }

        # Promote to a persisted rasQuarantine tag once this VM has been confirmed
        # agent-less for long enough AND enough separate polls have actually failed.
        # Best-effort and non-blocking: a failed tag write here just means this VM keeps
        # paying the (short) per-poll cost until a later poll's write succeeds -- it
        # never blocks or fails the guests/get call itself.
        #
        # EXCEPT for a RAS-managed clone or template (rasClone*/rasTemplate* tag): these
        # are guaranteed to run the guest agent per policy, and legitimate reboots during
        # clone preparation flap the agent in exactly the count-and-time shape this check
        # looks for, which produced a false quarantine here (a linked-clone recreate that
        # was mid-boot, not actually agent-less). Never write the persisted tag for one of
        # these -- but still log loudly once per failure episode, since the same reasoning
        # ("this must have a running agent") makes a genuinely stuck one an admin problem
        # to fix, not something this provider should silently paper over forever.
        $elapsedSeconds = ([DateTime]::UtcNow - $firstFailureAt).TotalSeconds
        if ($elapsedSeconds -ge $script:AgentQuarantineAfterSeconds -and $failureCount -ge $script:AgentQuarantineMinFailures) {
            $isRasManaged = (Test-ProxmoxVmHasTagPrefix -ClusterVm $ClusterVm -Prefix $script:RasCloneTagPrefix) -or
                            (Test-ProxmoxVmHasTagPrefix -ClusterVm $ClusterVm -Prefix $script:RasTemplateTagPrefix)

            if ($isRasManaged) {
                if ($script:RasManagedAgentExemptionLogged.Add($VmId)) {
                    Write-DebugLog "VM [$VmId] confirmed agent-less for ${elapsedSeconds}s across $failureCount poll(s), but carries a $($script:RasCloneTagPrefix)*/$($script:RasTemplateTagPrefix)* tag -- exempt from $($script:RasQuarantineTag) (RAS-managed guests are required to run the agent once powered on; if this persists past normal boot/reboot time, it is an admin-side problem on this guest, not a provider bug)." -Level 'I' -Component '03' -Ref $VmId
                }
            }
            else {
                try {
                    Add-ProxmoxVmTag -Node $Node -VmId $VmId -ClusterVm $ClusterVm -Tag $script:RasQuarantineTag
                    Write-DebugLog "VM [$VmId] confirmed agent-less for ${elapsedSeconds}s across $failureCount poll(s) -- tagged $($script:RasQuarantineTag)." -Level 'I' -Component '03' -Ref $VmId
                    $script:AgentFailureTracker.Remove($VmId)
                }
                catch {
                    Write-DebugLog "Failed to tag VM [$VmId] as $($script:RasQuarantineTag): $($_.Exception.Message)" -Level 'E' -Component '03' -Ref $VmId
                }
            }
        }

        return @()
    }
}

# 169.254.0.0/16 -- Windows/Linux self-assigns one of these (APIPA) when DHCP fails
# outright, never when it's merely still in progress. Reporting it (rather than
# stripping it to an empty list, indistinguishable from "still booting") is the whole
# point: a guest stuck on a link-local address during clone preparation is a real,
# actionable DHCP failure worth surfacing in guests/get and the log, not something to
# hide until a human happens to console into the VM. See Get-ProxmoxVmNetworkData and
# Get-RasGuestObjectForCloneAwareFlow's readiness gate, which deliberately does NOT
# treat one as "ready" -- reporting it and gating clone-completion on it are different
# questions.
function Test-ProxmoxIsLinkLocalIPv4 {
    param([string]$Address)
    return ([string]$Address) -match '^169\.254\.'
}

function Get-ProxmoxVmNetworkData {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Node,

        [Parameter(Mandatory = $true)]
        [string]$VmId,

        # The VM's already-known raw power state (from the cached cluster/resources
        # listing -- see ConvertTo-RasGuestObject). The guest agent can only answer when
        # the VM is actually running, so a non-running state skips that call entirely
        # instead of paying for a guaranteed-to-fail HTTP round trip.
        [string]$RawState = $null,

        # Threaded through to Get-ProxmoxVmGuestAgentInterfaces so the
        # rasQuarantine tag check/write there costs no extra HTTP call.
        [object]$ClusterVm = $null
    )

    $isRunning = -not [string]::IsNullOrWhiteSpace($RawState) -and $RawState.Trim().ToLowerInvariant() -eq 'running'

    if ($isRunning -and (Test-ProxmoxRecentlyStopped -VmId $VmId)) {
        # RAS often polls again within ~80ms of the stop POST this provider just issued,
        # and Proxmox's own status/current can still say 'running' for a brief window
        # after that -- even though the qemu process is already tearing down. The
        # provider already prefers that live read for power state, so this costs
        # nothing there; it only avoids a guest-agent probe that would otherwise
        # reliably 500 ("VM <id> is not running") once qemu actually exits mid-call.
        Write-DebugLog "Skipping guest-agent probe for VM [$VmId]: this provider stopped it moments ago." -Level 'T' -Component '03' -Ref $VmId
        if ($script:AgentFailureTracker.ContainsKey($VmId)) { $script:AgentFailureTracker.Remove($VmId) }
        if ($script:LastKnownNetworkData.ContainsKey($VmId)) { $script:LastKnownNetworkData.Remove($VmId) }
        return @{ IPv4Addresses = @(); MacAddresses = @() }
    }

    if (-not $isRunning) {
        # Not running -- no address should be reported, and a fresh boot deserves a
        # clean slate for both the negative cache and the last-known-good cache below,
        # rather than carrying a stale IP forward across a stop/start or into a reused
        # VMID.
        if ($script:AgentFailureTracker.ContainsKey($VmId)) { $script:AgentFailureTracker.Remove($VmId) }
        if ($script:LastKnownNetworkData.ContainsKey($VmId)) { $script:LastKnownNetworkData.Remove($VmId) }
        return @{ IPv4Addresses = @(); MacAddresses = @() }
    }

    $ipv4Set = New-Object 'System.Collections.Generic.HashSet[string]'
    $macSet = New-Object 'System.Collections.Generic.HashSet[string]'
    $interfaces = Get-ProxmoxVmGuestAgentInterfaces -Node $Node -VmId $VmId -ClusterVm $ClusterVm

    foreach ($iface in $interfaces) {
        $mac = $null
        if ((Get-MemberNames -Object $iface) -contains 'hardware-address') {
            $mac = [string]$iface.'hardware-address'
        }

        # Under Set-StrictMode, direct access on a property an interface does not have
        # (e.g. one with no addresses at all) throws "The property 'ip-addresses' cannot
        # be found on this object", which ConvertTo-RasGuestObject's outer try/catch
        # swallows -- silently degrading the guest to empty ip/mac lists on every poll.
        # Guarded the same way 'hardware-address' is above.
        $ipAddresses = if ((Get-MemberNames -Object $iface) -contains 'ip-addresses') { @($iface.'ip-addresses') } else { @() }
        foreach ($ip in $ipAddresses) {
            $type = [string]$ip.'ip-address-type'
            $addr = [string]$ip.'ip-address'

            # Link-local (169.254.0.0/16) is deliberately NOT excluded here -- see
            # Test-ProxmoxIsLinkLocalIPv4's comment. It's reordered to the back of the
            # final list below instead, so a genuinely reachable address still wins the
            # primary 'ip' slot whenever one exists.
            if ($type -eq 'ipv4' -and
                -not [string]::IsNullOrWhiteSpace($addr) -and
                $addr -ne '127.0.0.1') {

                [void]$ipv4Set.Add($addr)

                if (-not [string]::IsNullOrWhiteSpace($mac)) {
                    [void]$macSet.Add($mac.ToUpperInvariant())
                }
            }
        }
    }

    # Real addresses first, link-local last -- ConvertTo-RasGuestObject's 'ip' field is
    # always IPv4Addresses[0], and Select-Object -First 3 below can otherwise let a
    # secondary interface's APIPA address bump a real primary-interface address out of
    # the capped list entirely on a guest with more than 3 addresses total (rare, but
    # cheap to get right).
    $orderedIpv4 = @($ipv4Set | Sort-Object -Property @{ Expression = { Test-ProxmoxIsLinkLocalIPv4 -Address $_ } }, @{ Expression = { $_ } })

    $result = @{
        IPv4Addresses = @($orderedIpv4 | Select-Object -First 3)
        MacAddresses  = @($macSet  | Select-Object -First 3)
    }

    if ($result.IPv4Addresses.Count -gt 0) {
        $script:LastKnownNetworkData[$VmId] = $result
        return $result
    }

    # The agent call was skipped (quarantine, see Get-ProxmoxVmGuestAgentInterfaces) or
    # answered with nothing this time -- fall back to the last known-good value instead
    # of reporting the guest as having just lost its address, which RAS acts on ("IPs
    # will be reset").
    if ($script:LastKnownNetworkData.ContainsKey($VmId)) {
        return $script:LastKnownNetworkData[$VmId]
    }

    return $result
}

# A Proxmox clone target carries this auto-generated placeholder name ("VM <id>") until
# its clone job finishes. Shared between ConvertTo-RasGuestObject (which substitutes the
# tracked clone's real intended name) and Handle-GuestList (which exempts tracked clones
# from it) so both agree on what counts as a placeholder.
function Test-ProxmoxPlaceholderVmName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$VmId,

        [string]$VmName
    )

    return -not [string]::IsNullOrWhiteSpace($VmName) -and
        $VmId -match '^\d+$' -and
        $VmName -match ("^\s*VM\s+{0}\s*$" -f [regex]::Escape($VmId))
}

function ConvertTo-RasGuestObject {
    param(
        [Parameter(Mandatory = $true)]
        [string]$VmId
    )

    # $clusterVm comes from the cached cluster/resources listing, which already carries
    # node, name, status and template for every VM in one shot -- so those fields cost
    # no separate status/current or config call per guest, per poll. That listing has no
    # network info, which is why Get-ProxmoxVmNetworkData is still a live per-guest
    # call, and no host_os equivalent, so host_os reports 'unknown'.
    #
    # $script:LastGuestPollAt records that RAS asked about this specific guest just now,
    # so Invoke-OrphanAudit can tell "RAS is still actively polling this guest" apart
    # from "RAS has stopped asking about it entirely". Set here rather than in
    # Handle-GuestGet because both guests/get and hosts/get funnel through this
    # function, whichever path reaches it and however it exits.
    $script:LastGuestPollAt[$VmId] = [DateTime]::UtcNow

    $clusterVm = Get-ProxmoxVmNode -VmId $VmId
    $node = [string]$clusterVm.node

    # Opportunistic, like the tracked-clone sweep: costs nothing when this vmid has no
    # pending marker (the common case), and only does real work once the grace window
    # has actually elapsed for a vmid RAS happens to be polling anyway.
    if ($script:PendingTemplateTagRemovalIds.Count -gt 0) {
        Resolve-PendingTemplateTagRemoval -VmId $VmId -Node $node -ClusterVm $clusterVm
    }

    $rawState = 'unknown'
    if ((Get-MemberNames -Object $clusterVm) -contains 'status' -and -not [string]::IsNullOrWhiteSpace([string]$clusterVm.status)) {
        $rawState = [string]$clusterVm.status
    }

    # cluster/resources is a lagging, server-side aggregate; for a short window after WE
    # issue a control action on this VM, prefer a direct status/current read instead
    # (see Test-ProxmoxRecentlyControlled near the top of this file). Best-effort: a
    # failed live check just leaves the cached listing's state in place. The same live
    # read also settles the template flag, which is stale for the same reason, so a
    # recent convert is a second trigger for it.
    $liveIsTemplate = $null
    $convertIntent = Get-ProxmoxRecentConvertIntent -VmId $VmId

    if ((Test-ProxmoxRecentlyControlled -VmId $VmId) -or $null -ne $convertIntent -or (Test-ProxmoxInMaintenanceMode -VmId $VmId)) {
        try {
            $liveStatus = Get-ProxmoxVmCurrentStatus -Node $node -VmId $VmId
            $liveState = $null

            if ((Get-MemberNames -Object $liveStatus) -contains 'qmpstatus' -and -not [string]::IsNullOrWhiteSpace([string]$liveStatus.qmpstatus)) {
                $liveState = [string]$liveStatus.qmpstatus
            }
            elseif ((Get-MemberNames -Object $liveStatus) -contains 'status' -and -not [string]::IsNullOrWhiteSpace([string]$liveStatus.status)) {
                $liveState = [string]$liveStatus.status
            }

            if (-not [string]::IsNullOrWhiteSpace($liveState)) {
                $rawState = $liveState
            }

            # 'template' is optional in this response; absent means we learn
            # nothing here and fall through to the intent/cluster value below.
            if ((Get-MemberNames -Object $liveStatus) -contains 'template' -and $null -ne $liveStatus.template) {
                try { $liveIsTemplate = ([int]$liveStatus.template -eq 1) } catch { $liveIsTemplate = $null }
            }
        }
        catch {
            Write-DebugLog "Recently-controlled live status check failed for VM [$VmId]: $($_.Exception.Message)" -Level 'T' -Component '03' -Ref $VmId
        }
    }

    $name = if ((Get-MemberNames -Object $clusterVm) -contains 'name' -and -not [string]::IsNullOrWhiteSpace([string]$clusterVm.name)) {
        [string]$clusterVm.name
    }
    else {
        "VM-$VmId"
    }

    # RAS's clone thread cannot link a guest under Proxmox's transient placeholder name,
    # and pipelining a clone task 'completed' while it was still in effect leaves the
    # guest unresolvable for minutes. We already know the caller's own intended name for
    # a tracked clone (stored at guests/clone time), so substitute it here. Compare
    # against that known name rather than pattern-matching an assumed placeholder
    # format: cluster/resources can report "VM <id>", "VM-<id>", or briefly no 'name'
    # property at all -- in which case it is this function's own "VM-$VmId" fallback
    # below that would otherwise leak through unsubstituted.
    if ($script:TrackedCloneVmIds.Contains($VmId)) {
        $trackedEntry = Get-CloneStateEntry -VmId $VmId
        if ($null -ne $trackedEntry -and (Get-MemberNames -Object $trackedEntry) -contains 'name' -and
            -not [string]::IsNullOrWhiteSpace([string]$trackedEntry.name) -and
            [string]$trackedEntry.name -ne $name) {
            $name = [string]$trackedEntry.name
        }
    }

    # Precedence: a live status/current read beats what we asked Proxmox to do,
    # which beats the lagging cluster listing.
    $isTemplate = $false
    if ((Get-MemberNames -Object $clusterVm) -contains 'template') {
        try { $isTemplate = ([int]$clusterVm.template -eq 1) } catch { $isTemplate = $false }
    }

    if ($null -ne $liveIsTemplate) {
        $isTemplate = [bool]$liveIsTemplate
    }
    elseif ($null -ne $convertIntent) {
        $isTemplate = [bool]$convertIntent.is_template
    }

    $network = $null
    try {
        $network = Get-ProxmoxVmNetworkData -Node $node -VmId $VmId -RawState $rawState -ClusterVm $clusterVm
    }
    catch {
        Write-DebugLog "Network lookup failed for VM [$VmId]: $($_.Exception.Message)" -Level 'W' -Component '03' -Ref $VmId
        $network = @{
            IPv4Addresses = @()
            MacAddresses  = @()
        }
    }

    $guestObject = @{
        id            = [string]$VmId
        name          = $name
        provider      = 'Proxmox'
        node          = $node
        state         = (Map-ProxmoxStateToRasState -State $rawState)
        power_state   = $rawState
        host_os       = 'unknown'
        ip            = $(if ($network.IPv4Addresses.Count -gt 0) { $network.IPv4Addresses[0] } else { $null })
        ip_addresses  = @($network.IPv4Addresses)
        mac_addresses = @($network.MacAddresses)
        is_template   = $isTemplate
        type          = 'Virtual Machine'
    }

    Write-DebugLog ("GUEST VMID={0}; Name={1}; Node={2}; State={3}; Template={4}; IPs={5}" -f `
            $guestObject.id,
        $guestObject.name,
        $guestObject.node,
        $guestObject.state,
        $guestObject.is_template,
        ($guestObject.ip_addresses -join ',')
    ) -Level 'D' -Component '03' -Ref $guestObject.id

    return $guestObject
}

function Get-TrackedCloneContextByVmId {
    param(
        [Parameter(Mandatory = $true)]
        [string]$VmId
    )

    foreach ($key in @($script:TaskContext.Keys)) {
        $ctx = $script:TaskContext[$key]
        if ($null -ne $ctx -and
            $ctx.ContainsKey('type') -and
            [string]$ctx.type -eq 'clone' -and
            $ctx.ContainsKey('clone_id') -and
            [string]$ctx.clone_id -eq [string]$VmId) {

            Write-DebugLog "TRACKED CLONE FOUND IN MEMORY for VM [$VmId]" -Level 'D' -Component '04' -Ref $VmId
            return @{
                task_id = $key
                context = $ctx
            }
        }
    }

    $persisted = Get-CloneStateEntry -VmId $VmId
    if ($null -ne $persisted) {
        $ctx = @{}
        foreach ($p in $persisted.PSObject.Properties) {
            $ctx[$p.Name] = $p.Value
        }

        if (-not $ctx.ContainsKey('type')) {
            $ctx.type = 'clone'
        }

        if (-not $ctx.ContainsKey('clone_id')) {
            $ctx.clone_id = [string]$VmId
        }

        # NOT a disk read, despite the log line's "FOUND VIA ... CACHE" wording --
        # Get-CloneStateEntry/Get-CloneStateAll serve from $script:CloneStateMemory, an
        # in-memory mirror of the persisted file. The real distinction from the FOUND IN
        # MEMORY case above is which in-memory structure hit: $script:TaskContext
        # (per-process, keyed by task_id) vs. this VmId-keyed clone-state cache (loaded
        # once and kept current, so it also answers correctly after a provider restart).
        Write-DebugLog "TRACKED CLONE FOUND VIA CLONE-STATE CACHE for VM [$VmId]" -Level 'D' -Component '04' -Ref $VmId
        return @{
            task_id = $null
            context = $ctx
        }
    }

    # No log line for the not-found case -- this is the overwhelmingly common outcome
    # (any guest that is not an in-flight clone) and would otherwise log on every
    # guests/get poll for every guest.
    return $null
}

function Start-ProxmoxVmIfNeeded {
    param(
        [Parameter(Mandatory = $true)]
        [string]$VmId,

        # When set, this is the source (template) VMID this VM was cloned from. Both
        # callers already know the clone's real underlying Proxmox job has finished, so
        # applying the rasClone<id> tag here (rather than duplicating it in both)
        # happens once, best-effort and non-blocking -- a failed tag write never affects
        # the start itself.
        [string]$CloneSourceVmId = $null
    )

    $clusterVm = Get-ProxmoxVmNode -VmId $VmId
    $node = [string]$clusterVm.node

    # Skip re-verifying a VM's tags on every poll once they are known correct -- see
    # $script:CloneTagVerified near the top of this file. Both guests/get's clone-aware
    # flow and tasks/get's Handle-TaskInfo call this function on every poll while the
    # clone is still powered off, and Repair-ProxmoxCloneTagsAfterInherit otherwise does a
    # live 'config' GET every single time to check tags that, once correct, do not change
    # on their own.
    if (-not [string]::IsNullOrWhiteSpace($CloneSourceVmId) -and -not $script:CloneTagVerified.Contains($VmId)) {
        # Repair-ProxmoxCloneTagsAfterInherit also strips any rasTemplate* tag this
        # clone inherited from its source at clone time -- see that function.
        $cloneTag = "$($script:RasCloneTagPrefix)$CloneSourceVmId"
        try {
            $tagRepair = Repair-ProxmoxCloneTagsAfterInherit -Node $node -VmId $VmId -ClusterVm $clusterVm -CloneTag $cloneTag
            if ($tagRepair.changed) {
                Write-DebugLog "Tagged clone VM [$VmId] with [$cloneTag]." -Level 'I' -Component '04' -Ref $VmId
                if (@($tagRepair.removed_inherited_template_tags).Count -gt 0) {
                    Write-DebugLog "Clone VM [$VmId] had inherited $($tagRepair.removed_inherited_template_tags -join ', ') from its source at clone time -- stripped." -Level 'D' -Component '04' -Ref $VmId
                }
            }
            # Only cache success -- a failed repair (caught below) must retry on the next
            # poll rather than being silently skipped forever.
            [void]$script:CloneTagVerified.Add($VmId)
        }
        catch {
            Write-DebugLog "Failed to tag clone VM [$VmId] with [$cloneTag]: $($_.Exception.Message)" -Level 'E' -Component '04' -Ref $VmId
        }
    }

    # cloning.mac_preservation.enabled -- see MAC-PRESERVATION.md. Same "once per VmId
    # this process lifetime" gate as tag repair above, and the SAME reason: this and
    # Handle-TaskInfo both call this function on every poll while a clone is still
    # powered off, and applying it more than once would mean an unnecessary live config
    # GET + PUT per poll for no further effect. Must run here -- before the start below
    # -- not after: a MAC applied post-boot doesn't reliably trigger a fresh DHCP
    # handshake or satisfy a boot-time licensing check, which is the whole point of
    # preserving it at all.
    if (-not [string]::IsNullOrWhiteSpace($CloneSourceVmId) -and -not $script:CloneMacRestored.Contains($VmId)) {
        $trackedForMac = Get-TrackedCloneContextByVmId -VmId $VmId
        $contextForMac = if ($null -ne $trackedForMac) { $trackedForMac.context } else { $null }
        # Under Set-StrictMode, direct .preserved_macs access throws on a context that
        # predates this field (a tracking entry a test -- or a clone still tracked
        # across an upgrade -- built without it) rather than just being $null.
        $preservedMacs = if ($null -ne $contextForMac -and (Get-MemberNames -Object $contextForMac) -contains 'preserved_macs') { $contextForMac.preserved_macs } else { $null }
        $preservedMacKeys = @(if ($null -ne $preservedMacs) { Get-MemberNames -Object $preservedMacs } else { @() })

        if ($preservedMacKeys.Count -gt 0) {
            try {
                $liveConfig = Get-ProxmoxVmConfig -Node $node -VmId $VmId
                $restoredKeys = New-Object System.Collections.ArrayList
                foreach ($netKey in $preservedMacKeys) {
                    if (-not ((Get-MemberNames -Object $liveConfig) -contains $netKey)) {
                        continue   # this clone doesn't have the interface the old VM had -- nothing to restore it onto
                    }
                    $desiredMac = [string]$preservedMacs.$netKey
                    $currentLine = [string]$liveConfig.$netKey
                    $newLine = Set-ProxmoxNetIfaceMacInLine -CurrentValue $currentLine -NewMac $desiredMac
                    if ($newLine -eq $currentLine) {
                        continue   # already matches (e.g. a retry after a partial prior failure) -- no PUT needed
                    }
                    [void](Invoke-ProxmoxApi -Method PUT -Path "/api2/json/nodes/$node/qemu/$VmId/config" -Body @{ $netKey = $newLine } -TimeoutSec $script:TagWriteTimeoutSeconds -NoRetry)
                    [void]$restoredKeys.Add("$netKey=$desiredMac")
                }
                if ($restoredKeys.Count -gt 0) {
                    Write-DebugLog "Restored MAC address(es) on recreated VM [$VmId]: $($restoredKeys -join ', ')." -Level 'I' -Component '04' -Ref $VmId
                }
                # Only cache success -- a failed restore (caught below) must retry on the
                # next poll rather than being silently skipped forever, same reasoning as
                # $script:CloneTagVerified above.
                [void]$script:CloneMacRestored.Add($VmId)
            }
            catch {
                Write-DebugLog "Failed to restore MAC address(es) on recreated VM [$VmId]: $($_.Exception.Message)" -Level 'E' -Component '04' -Ref $VmId
            }
        }
        else {
            [void]$script:CloneMacRestored.Add($VmId)   # nothing to restore -- don't re-check every poll
        }
    }

    $current = Get-ProxmoxVmCurrentStatus -Node $node -VmId $VmId

    $rawState = ''
    if ((Get-MemberNames -Object $current) -contains 'qmpstatus' -and -not [string]::IsNullOrWhiteSpace([string]$current.qmpstatus)) {
        $rawState = [string]$current.qmpstatus
    }
    elseif ((Get-MemberNames -Object $current) -contains 'status' -and -not [string]::IsNullOrWhiteSpace([string]$current.status)) {
        $rawState = [string]$current.status
    }

    $rasState = Map-ProxmoxStateToRasState -State $rawState

    if ($rasState -eq 'powered_on' -or $rasState -eq 'powering_on') {
        return @{
            started       = $false
            pending       = $false
            node          = $node
            raw_state     = $rawState
            ras_state     = $rasState
            task_id       = $null
            error_message = $null
        }
    }

    try {
        $resp = Invoke-ProxmoxApi -Method POST -Path "/api2/json/nodes/$node/qemu/$VmId/status/start" -Body @{}
        $startTaskId = $null
        if ($null -ne $resp -and (Get-MemberNames -Object $resp) -contains 'data') {
            $startTaskId = [string]$resp.data
        }

        Write-DebugLog "Issued start for VM [$VmId], start task id=[$startTaskId]" -Level 'T' -Component '04' -Ref $VmId

        # Every field guests/get reports (state, power_state, and whether the
        # network/agent call even runs -- Get-ProxmoxVmNetworkData is gated on the same
        # raw state) comes from the cached cluster/resources listing, which still shows
        # this VM 'stopped' for up to its full TTL after the start we just issued. Bust
        # it, or a freshly auto-started clone reports 'powering_on' for the whole TTL
        # regardless of how quickly it actually boots.
        Reset-ProxmoxClusterCache

        return @{
            started       = $true
            pending       = $false
            node          = $node
            raw_state     = $rawState
            ras_state     = $rasState
            task_id       = $startTaskId
            error_message = $null
        }
    }
    catch {
        $msg = $_.Exception.Message
        Write-DebugLog "Start attempt for VM [$VmId] failed: $msg" -Level 'E' -Component '04' -Ref $VmId

        if ($msg -match "can't lock file" -or $msg -match 'got timeout' -or $msg -match 'VM is locked') {
            return @{
                started       = $false
                pending       = $true
                node          = $node
                raw_state     = $rawState
                ras_state     = $rasState
                task_id       = $null
                error_message = $msg
            }
        }

        throw
    }
}

function Get-RasGuestObjectForCloneAwareFlow {
    param(
        [Parameter(Mandatory = $true)]
        [string]$VmId
    )

    $tracked = Get-TrackedCloneContextByVmId -VmId $VmId

    # Hard ceiling, checked before anything else touches Proxmox: a clone tracked past
    # cloning.timeouts.clone_tracking_max_age_seconds without ever reaching ready is
    # abandoned unconditionally here, regardless of whether $VmId currently resolves.
    # This is the backstop for the case nothing else catches -- a VM deleted outside
    # this provider's own guests/control(delete) path (by hand in the Proxmox UI, most
    # commonly after a failed/overloaded clone), which otherwise leaves
    # Get-ActiveCloneCount counting it as active forever (only a CONFIRMED completed
    # real task clears that, never a failed one) and this function reporting
    # powering_on forever below. Falling through with $tracked cleared means the rest
    # of this call behaves exactly like any other untracked guest from this point on --
    # its real state if it still resolves, or the normal not-found error if it doesn't.
    if ($null -ne $tracked) {
        $ctx0 = $tracked.context
        $cloneStartedAt = $null
        if ($ctx0.ContainsKey('clone_started_at') -and -not [string]::IsNullOrWhiteSpace([string]$ctx0.clone_started_at)) {
            try { $cloneStartedAt = [DateTime]::Parse([string]$ctx0.clone_started_at).ToUniversalTime() } catch { }
        }

        if ($null -ne $cloneStartedAt) {
            $trackedAgeSeconds = ([DateTime]::UtcNow - $cloneStartedAt).TotalSeconds
            if ($trackedAgeSeconds -ge $script:CloneTrackingMaxAgeSeconds) {
                Write-DebugLog "CLONE-AWARE FLOW: VM [$VmId] tracking abandoned -- never reached ready within cloning.timeouts.clone_tracking_max_age_seconds (${trackedAgeSeconds}s elapsed since clone_started_at). A tracked clone still this far from ready almost always means it was deleted outside RAS or the underlying clone failed; giving up rather than tracking it forever." -Level 'W' -Component '04' -Ref $VmId
                Clear-ProxmoxTrackingForVm -VmId $VmId
                $tracked = $null
            }
        }
    }

    try {
        $guest = ConvertTo-RasGuestObject -VmId $VmId
    }
    catch {
        if ($null -eq $tracked) {
            throw
        }

        # A tracked in-flight clone's VM id can be momentarily unresolvable (stale
        # cluster-resources cache, or Proxmox itself has not registered it yet) right
        # after guests/clone returns, and RAS's own guests/get has no retry on this
        # path. Report this as still provisioning instead of surfacing the hard error; a
        # later poll resolves it once the id becomes visible.
        #
        # The name reported here must be the clone's real, RAS-requested name (already
        # known -- Set-CloneStateEntry stored it at clone-submission time), never a
        # generic "VM-<id>" placeholder: Proxmox itself keeps a fresh clone under a
        # placeholder name for a short window after creation, and if RAS sees that
        # placeholder on its very first guests/get for this id, it may be unable to
        # correlate the guest back to the guests/clone request it is waiting on.
        $trackedCloneName = if ($ctx0.ContainsKey('name') -and -not [string]::IsNullOrWhiteSpace([string]$ctx0.name)) { [string]$ctx0.name } else { "VM-$VmId" }
        Write-DebugLog "CLONE-AWARE FLOW: VM [$VmId] not yet resolvable ($($_.Exception.Message)) -- reporting powering_on for tracked clone instead of erroring." -Level 'T' -Component '04' -Ref $VmId
        $guest = @{
            id            = [string]$VmId
            name          = $trackedCloneName
            provider      = 'Proxmox'
            node          = $null
            state         = 'powering_on'
            power_state   = 'starting'
            host_os       = 'unknown'
            ip            = $null
            ip_addresses  = @()
            mac_addresses = @()
            is_template   = $false
            type          = 'Virtual Machine'
        }
    }

    if ($null -eq $tracked) {
        # No log line here either -- same reasoning as Get-TrackedCloneContextByVmId
        # above: this is the normal case for any non-clone guest, every poll.
        return $guest
    }

    Write-DebugLog "CLONE-AWARE FLOW: tracked clone context found for VM [$VmId]" -Level 'D' -Component '04' -Ref $VmId

    $ctx = $tracked.context

    if (-not $ctx.ContainsKey('start_issued')) {
        $ctx.start_issued = $false
    }

    if (-not $ctx.ContainsKey('start_pending')) {
        $ctx.start_pending = $false
    }

    if (-not $ctx.ContainsKey('start_retry_count')) {
        $ctx.start_retry_count = 0
    }

    if (-not $ctx.ContainsKey('creation_completed')) {
        $ctx.creation_completed = $false
    }

    try {
        $ctx.start_issued = [System.Convert]::ToBoolean($ctx.start_issued)
    }
    catch {
        $ctx.start_issued = $false
    }

    try {
        $ctx.start_pending = [System.Convert]::ToBoolean($ctx.start_pending)
    }
    catch {
        $ctx.start_pending = $false
    }

    try {
        $ctx.creation_completed = [System.Convert]::ToBoolean($ctx.creation_completed)
    }
    catch {
        $ctx.creation_completed = $false
    }

    try {
        $ctx.start_retry_count = [int]$ctx.start_retry_count
    }
    catch {
        $ctx.start_retry_count = 0
    }

    $cloneTrackingCompleted = $false

    # Always treat tracked clones as provisioning candidates until they are fully ready.
    if ($guest.state -eq 'powered_off' -or $guest.state -eq 'powering_off') {
        # THIS -- not an explicit RAS guests/control(start) -- is the start that hits
        # Proxmox's clone lock: guests/get is polled far more often than tasks/get, so
        # under pipelined completion this branch runs within milliseconds of reporting
        # the clone task 'completed', almost always before Proxmox's real clone job has
        # actually finished.
        #
        # Unlike Wait-ForRealCloneCompletionBeforeStart (which blocks, because it only
        # runs once per explicit guests/control(start)), this path must NOT block: it
        # runs on every single guests/get poll for this guest, so blocking here would
        # serialize the whole shared stdin/stdout pipe behind every poll of every
        # in-flight clone. Instead it skips the start attempt entirely (reporting
        # 'powering_on' without touching Proxmox) until the real clone task is confirmed
        # done; a later poll issues the start.
        #
        # A start issued while the clone lock is still held is ACCEPTED by Proxmox (200
        # + a start UPID) and then fails asynchronously inside that task, leaving the
        # guest powered off. So if the previous start task is known to have failed,
        # clear it here and let the retry below fire on this same poll.
        if ($ctx.start_issued -and $ctx.ContainsKey('start_task_id') -and -not [string]::IsNullOrWhiteSpace([string]$ctx.start_task_id)) {
            try {
                $priorStartResult = New-TaskResultState -TaskStatus (Get-ProxmoxTaskStatus -TaskId ([string]$ctx.start_task_id))
                if ($priorStartResult.state -eq 'failed') {
                    # .error is the @{code;message} hashtable New-TaskResultState builds,
                    # so interpolating it whole prints "System.Collections.Hashtable" and
                    # loses Proxmox's exitstatus -- the one detail worth having here.
                    Write-DebugLog "Clone-aware get: VM [$VmId] previous start task [$($ctx.start_task_id)] failed ($($priorStartResult.error.message)) -- retrying." -Level 'W' -Component '04' -Ref $VmId
                    $ctx.start_issued = $false
                    $ctx.start_task_id = $null
                }
            }
            catch {
                # Can't determine the previous start's outcome (e.g. aged out of
                # Proxmox's task history) -- leave start_issued as-is rather than
                # retry blindly.
            }
        }

        $realCloneTaskId = if ($ctx.ContainsKey('task_id')) { [string]$ctx.task_id } else { $null }
        $realCloneReady = $true
        if (-not [string]::IsNullOrWhiteSpace($realCloneTaskId)) {
            try { $realCloneReady = ((New-TaskResultState -TaskStatus (Get-ProxmoxTaskStatus -TaskId $realCloneTaskId)).state -eq 'completed') }
            catch { $realCloneReady = $true }   # can't check (e.g. task aged out of Proxmox's history) -- don't block a start over that
        }

        if ($realCloneReady) {
            $cloneSourceId = if ($ctx.ContainsKey('source_id')) { [string]$ctx.source_id } else { $null }
            $startInfo = Start-ProxmoxVmIfNeeded -VmId $VmId -CloneSourceVmId $cloneSourceId
            $ctx.start_retry_count = [int]$ctx.start_retry_count + 1
            $ctx.clone_node = $startInfo.node

            if ($startInfo.started) {
                $ctx.start_issued = $true
                $ctx.start_pending = $false
                $ctx.start_task_id = $startInfo.task_id
                Write-DebugLog "Clone-aware get: VM [$VmId] was off, start issued." -Level 'T' -Component '04' -Ref $VmId
            }
            elseif ($startInfo.pending) {
                $ctx.start_pending = $true
                Write-DebugLog "Clone-aware get: VM [$VmId] still locked, start deferred." -Level 'T' -Component '04' -Ref $VmId
            }
        }
        else {
            Write-DebugLog "Clone-aware get: VM [$VmId] real clone task [$realCloneTaskId] not finished yet -- deferring start attempt." -Level 'D' -Component '04' -Ref $VmId
        }

        $guest.state = 'powering_on'
        $guest.power_state = 'starting'
    }
    elseif ($guest.state -eq 'powered_on') {
        # Deliberately stricter than "has any address at all": ip_addresses can now
        # legitimately carry a link-local (169.254.x.x) address on its own -- reported
        # for visibility (see Test-ProxmoxIsLinkLocalIPv4), not treated as ready. A
        # guest stuck on link-local means DHCP failed, not that it's reachable; clone
        # completion stays gated on a real address, same as before this changed.
        $realIps = @($guest.ip_addresses | Where-Object { -not (Test-ProxmoxIsLinkLocalIPv4 -Address $_) })
        $hasRealIp = $realIps.Count -gt 0

        if ($hasRealIp) {
            $ctx.creation_completed = $true
            $cloneTrackingCompleted = $true
            Write-DebugLog "Clone-aware get: VM [$VmId] is powered on and has IP(s) [$($guest.ip_addresses -join ',')]." -Level 'I' -Component '04' -Ref $VmId
        }
        else {
            $guest.state = 'powering_on'
            $guest.power_state = 'starting'
            $linkLocalNote = if (@($guest.ip_addresses).Count -gt 0) { " (only link-local: [$($guest.ip_addresses -join ',')] -- DHCP appears to be failing)" } else { '' }
            Write-DebugLog "Clone-aware get: VM [$VmId] is powered on but has no IP yet. Reporting powering_on.$linkLocalNote" -Level 'D' -Component '04' -Ref $VmId
        }
    }
    else {
        Write-DebugLog "Clone-aware get: VM [$VmId] currently in state [$($guest.state)]." -Level 'D' -Component '04' -Ref $VmId
    }

    if ($cloneTrackingCompleted) {
        $resolvedTaskId = $null
        if ($null -ne $tracked.task_id -and -not [string]::IsNullOrWhiteSpace([string]$tracked.task_id)) {
            $resolvedTaskId = [string]$tracked.task_id
        }
        elseif ($ctx.ContainsKey('task_id') -and -not [string]::IsNullOrWhiteSpace([string]$ctx.task_id)) {
            $resolvedTaskId = [string]$ctx.task_id
        }

        # Stash the clone_id BEFORE clearing tracking, so a tasks/get poll for this same
        # task arriving afterward (Handle-TaskInfo's own context lookups will both come
        # up empty from here on) can still answer with the real clone_id instead of an
        # empty output. See $script:CompletedCloneTaskOutputs.
        Set-CompletedCloneTaskOutput -TaskId $resolvedTaskId -CloneId $VmId

        if (-not [string]::IsNullOrWhiteSpace($resolvedTaskId) -and $script:TaskContext.ContainsKey($resolvedTaskId)) {
            [void]$script:TaskContext.Remove($resolvedTaskId)
        }

        Remove-CloneStateEntry -VmId $VmId
        Write-DebugLog "CLONE-AWARE FLOW: tracking removed for VM [$VmId] after successful clone completion." -Level 'I' -Component '04' -Ref $VmId
    }
    else {
        if ($null -ne $tracked.task_id) {
            $script:TaskContext[$tracked.task_id] = $ctx
        }

        Set-CloneStateEntry -VmId $VmId -Entry $ctx
    }

    Write-DebugLog "CLONE-AWARE FLOW RESULT for VM [$VmId]: state=[$($guest.state)] power_state=[$($guest.power_state)]" -Level 'D' -Component '04' -Ref $VmId

    return $guest
}

# Once a clone is pipelined, RAS stops polling tasks/get for it, so only guests/get can
# drive its start/readiness onward -- and RAS's per-guest reconciliation cadence is far
# slower than tasks_polling_rate.
#
# Every guests/get already re-derives the guest it was asked about via
# Get-RasGuestObjectForCloneAwareFlow, which does all the real work (auto-start,
# IP-ready detection, tracking cleanup) for that one id. This piggybacks the same check
# onto every OTHER in-flight clone, so a clone becomes ready within one guests/get for
# ANY guest. With nothing in flight the cost is one O(1) HashSet.Count check --
# $script:TrackedCloneVmIds is kept in sync by Set-/Remove-CloneStateEntry specifically
# so this needs no file read.
function Invoke-TrackedCloneSweep {
    param(
        [string]$ExcludeVmId = $null
    )

    if ($script:TrackedCloneVmIds.Count -eq 0) {
        return
    }

    foreach ($vmId in @($script:TrackedCloneVmIds)) {
        if (-not [string]::IsNullOrWhiteSpace($ExcludeVmId) -and $vmId -eq $ExcludeVmId) {
            continue
        }

        # Throttled: unthrottled, this sweep dominated the provider's HTTP traffic and
        # tripled the latency of every unrelated guests/get during a clone window. Safe
        # because ConvertTo-RasGuestObject now substitutes a tracked clone's real name
        # for Proxmox's placeholder, so RAS can poll a clone directly instead of
        # depending on this sweep alone to ever reach it.
        $lastChecked = $script:SweepLastCheckedAt[$vmId]
        if ($null -ne $lastChecked -and ([DateTime]::UtcNow - [DateTime]$lastChecked).TotalSeconds -lt $script:SweepMinIntervalSeconds) {
            continue
        }

        try {
            [void](Get-RasGuestObjectForCloneAwareFlow -VmId $vmId)
        }
        catch {
            Write-DebugLog "Opportunistic clone sweep: VM [$vmId] check failed: $($_.Exception.Message)" -Level 'W' -Component '04' -Ref $vmId
        }
        finally {
            $script:SweepLastCheckedAt[$vmId] = [DateTime]::UtcNow
        }
    }
}

function Get-ControlAction {
    param([Parameter(Mandatory = $true)][string]$Control)

    switch ($Control.Trim().ToLowerInvariant()) {
        'start' { return 'start' }
        # RAS has its own separate graceful-stop path via the RAS guest agent inside the
        # VM -- that never reaches this provider at all. A guests/control(stop) that
        # DOES reach us is RAS's hypervisor-level fallback/force stop, so this maps to
        # Proxmox's 'stop' (immediate power-off, no guest cooperation needed), not the
        # graceful 'shutdown'.
        'stop' { return 'stop' }
        'reset' { return 'reset' }
        'restart' { return 'reboot' }
        'reboot' { return 'reboot' }
        'delete' { return 'delete' }
        'suspend' { return 'suspend' }
        default { return $null }
    }
}

function Get-ProxmoxNextVmId {
    $resp = Invoke-ProxmoxApi -Method GET -Path '/api2/json/cluster/nextid'
    return [string]$resp.data
}

function Get-ProxmoxTaskNodeFromUpid {
    param([string]$Upid)

    if ([string]::IsNullOrWhiteSpace($Upid)) {
        throw 'Task id is empty'
    }

    $parts = $Upid -split ':'
    if ($parts.Count -lt 3 -or $parts[0] -ne 'UPID') {
        throw "Invalid Proxmox task id format: $Upid"
    }

    return $parts[1]
}

function Get-ProxmoxTaskStatus {
    param([string]$TaskId)

    $node = Get-ProxmoxTaskNodeFromUpid -Upid $TaskId
    $escapedTaskId = [System.Uri]::EscapeDataString($TaskId)
    $resp = Invoke-ProxmoxApi -Method GET -Path "/api2/json/nodes/$node/tasks/$escapedTaskId/status"
    return $resp.data
}

function New-TaskResultState {
    param([object]$TaskStatus)

    if ($null -eq $TaskStatus) {
        return @{
            state = 'failed'
            error = @{
                code    = 1
                message = 'Task status unavailable'
            }
        }
    }

    $status = ''
    if ((Get-MemberNames -Object $TaskStatus) -contains 'status' -and $TaskStatus.status) {
        $status = [string]$TaskStatus.status
    }

    $exitStatus = $null
    if ((Get-MemberNames -Object $TaskStatus) -contains 'exitstatus') {
        $exitStatus = [string]$TaskStatus.exitstatus
    }

    if ($status -eq 'running') {
        return @{ state = 'running' }
    }

    if ($status -eq 'stopped' -and $exitStatus -eq 'OK') {
        return @{ state = 'completed' }
    }

    $msg = if (-not [string]::IsNullOrWhiteSpace($exitStatus)) { $exitStatus } else { 'Unknown task failure' }
    return @{
        state = 'failed'
        error = @{
            code    = 1
            message = $msg
        }
    }
}

# Guard for Handle-GuestControl's 'start' branch -- see PIPELINED-CLONING.md and the
# constants near the top of this file. If $VmId is a clone whose pipelined tasks/get
# 'completed' already went out to RAS but whose REAL underlying Proxmox clone task has
# not actually finished yet, calling Proxmox's own start endpoint now would hit
# Proxmox's clone lock and fail. This polls the real task to real completion (or a
# bounded timeout/failure) before the caller may proceed.
#
# Returns @{ shouldWait = $false } when $VmId is not a tracked, still-in-flight clone at
# all -- callers should proceed immediately. Otherwise returns @{ shouldWait = $true; ok
# = <bool>; message = <string> }.
function Wait-ForRealCloneCompletionBeforeStart {
    param([Parameter(Mandatory = $true)][string]$VmId)

    $tracked = Get-TrackedCloneContextByVmId -VmId $VmId
    if ($null -eq $tracked) {
        return @{ shouldWait = $false }
    }

    $ctx = $tracked.context
    $realTaskId = $null
    if ($ctx.ContainsKey('task_id') -and -not [string]::IsNullOrWhiteSpace([string]$ctx.task_id)) {
        $realTaskId = [string]$ctx.task_id
    }

    if ([string]::IsNullOrWhiteSpace($realTaskId)) {
        # Tracked, but no real Proxmox task id to poll (shouldn't normally happen for
        # a clone entry) -- nothing we can wait on, let the caller proceed as before.
        return @{ shouldWait = $false }
    }

    Write-DebugLog "Guest control start for VM [$VmId]: clone still tracked (real task [$realTaskId]) -- polling real completion before starting." -Level 'T' -Component '04' -Ref $VmId

    $deadline = [DateTime]::UtcNow.AddSeconds($script:PipelinedCloneStartMaxWaitSeconds)
    while ($true) {
        $realResult = New-TaskResultState -TaskStatus (Get-ProxmoxTaskStatus -TaskId $realTaskId)

        if ($realResult.state -eq 'completed') {
            Write-DebugLog "Guest control start for VM [$VmId]: real clone task [$realTaskId] is now genuinely completed -- proceeding to start." -Level 'T' -Component '04' -Ref $VmId
            return @{ shouldWait = $true; ok = $true }
        }

        if ($realResult.state -eq 'failed') {
            $failMsg = "Clone task [$realTaskId] for VM [$VmId] failed: $($realResult.error.message)"
            Write-DebugLog "Guest control start for VM [$VmId]: $failMsg -- refusing to start." -Level 'E' -Component '04' -Ref $VmId
            return @{ shouldWait = $true; ok = $false; message = $failMsg }
        }

        if ([DateTime]::UtcNow -ge $deadline) {
            $timeoutMsg = "Clone task [$realTaskId] for VM [$VmId] did not complete within $($script:PipelinedCloneStartMaxWaitSeconds)s; refusing to start."
            Write-DebugLog "Guest control start for VM [$VmId]: $timeoutMsg" -Level 'E' -Component '04' -Ref $VmId
            return @{ shouldWait = $true; ok = $false; message = $timeoutMsg }
        }

        Start-Sleep -Seconds $script:PipelinedCloneStartPollIntervalSeconds
    }
}

# Guard for Handle-GuestControl's 'delete' branch. Proxmox will not destroy a
# still-running VM, so delete forces its own hard stop first and polls the real stop
# task to completion before destroying, rather than assuming some earlier request
# already got the VM stopped. That keeps delete self-contained and correct even if RAS
# calls it directly without a preceding stop.
function Stop-ProxmoxVmHardBeforeDelete {
    param(
        [Parameter(Mandatory = $true)][string]$Node,
        [Parameter(Mandatory = $true)][string]$VmId
    )

    $current = $null
    try { $current = Get-ProxmoxVmCurrentStatus -Node $Node -VmId $VmId }
    catch { Write-DebugLog "Delete guest [$VmId]: failed to read current status before hard stop, attempting stop anyway: $($_.Exception.Message)" -Level 'W' -Component '05' -Ref $VmId }

    $rawState = ''
    if ($null -ne $current) {
        if ((Get-MemberNames -Object $current) -contains 'qmpstatus' -and -not [string]::IsNullOrWhiteSpace([string]$current.qmpstatus)) {
            $rawState = [string]$current.qmpstatus
        }
        elseif ((Get-MemberNames -Object $current) -contains 'status' -and -not [string]::IsNullOrWhiteSpace([string]$current.status)) {
            $rawState = [string]$current.status
        }
    }

    if ($rawState.Trim().ToLowerInvariant() -eq 'stopped') {
        Write-DebugLog "Delete guest [$VmId]: already stopped, skipping hard stop." -Level 'T' -Component '05' -Ref $VmId
        return @{ ok = $true }
    }

    Write-DebugLog "Delete guest [$VmId]: issuing hard stop (current state [$rawState]) before destroy." -Level 'T' -Component '05' -Ref $VmId
    $stopResp = Invoke-ProxmoxApiWithRetry -Method POST -Path "/api2/json/nodes/$Node/qemu/$VmId/status/stop" -Body @{}
    $stopTaskId = $null
    if ($null -ne $stopResp -and (Get-MemberNames -Object $stopResp) -contains 'data') {
        $stopTaskId = [string]$stopResp.data
    }

    if ([string]::IsNullOrWhiteSpace($stopTaskId)) {
        # No task id to poll -- Proxmox answered synchronously; nothing further to wait on.
        return @{ ok = $true }
    }

    $deadline = [DateTime]::UtcNow.AddSeconds($script:DeleteHardStopMaxWaitSeconds)
    while ($true) {
        $result = New-TaskResultState -TaskStatus (Get-ProxmoxTaskStatus -TaskId $stopTaskId)

        if ($result.state -eq 'completed') {
            Write-DebugLog "Delete guest [$VmId]: hard stop task [$stopTaskId] completed." -Level 'T' -Component '05' -Ref $VmId
            return @{ ok = $true }
        }

        if ($result.state -eq 'failed') {
            $msg = "Hard stop task [$stopTaskId] for VM [$VmId] failed: $($result.error.message)"
            Write-DebugLog "Delete guest [$VmId]: $msg -- refusing to destroy a VM whose stop failed." -Level 'E' -Component '05' -Ref $VmId
            return @{ ok = $false; message = $msg }
        }

        if ([DateTime]::UtcNow -ge $deadline) {
            $msg = "Hard stop task [$stopTaskId] for VM [$VmId] did not complete within $($script:DeleteHardStopMaxWaitSeconds)s"
            Write-DebugLog "Delete guest [$VmId]: $msg -- refusing to destroy." -Level 'E' -Component '05' -Ref $VmId
            return @{ ok = $false; message = $msg }
        }

        Start-Sleep -Seconds $script:DeleteHardStopPollIntervalSeconds
    }
}

function Handle-Initialize {
    # Sourced from RAS-CPF-Proxmox-Settings.json (capabilities section). A settings
    # reload updates $script:Settings live, so the values returned here reflect whatever
    # was on disk as of this call, not necessarily what was in effect when the provider
    # process started.
    return @{
        result = @{
            version      = $script:ProviderVersion
            capabilities = @{
                can_suspend_guests    = $script:Settings.capabilities.can_suspend_guests
                guests_polling_rate   = $script:Settings.capabilities.guests_polling_rate
                tasks_polling_rate    = $script:Settings.capabilities.tasks_polling_rate
                tasks_polling_retries = $script:Settings.capabilities.tasks_polling_retries
                template_method       = $script:Settings.capabilities.template_method
                # Same value Handle-GuestClone actually gates the linked-clone decision
                # on -- see the comment on this key in Get-DefaultProviderSettings for
                # why that single-source property matters.
                can_link_clones       = $script:Settings.capabilities.can_link_clones
            }
        }
    }
}

function Handle-Connect {
    param([object]$Params)

    $settings = $Params.settings
    if ($null -eq $settings) {
        return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Missing settings"
    }

    $proxmoxHost = [string]$settings.host
    $username = [string]$settings.username
    $tokenName = [string]$settings.token_name
    $tokenSecret = [string]$settings.token_secret

    if ([string]::IsNullOrWhiteSpace($proxmoxHost) -or
        [string]::IsNullOrWhiteSpace($username) -or
        [string]::IsNullOrWhiteSpace($tokenName) -or
        [string]::IsNullOrWhiteSpace($tokenSecret)) {

        return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Invalid connection parameters"
    }

    try {
        Initialize-CertificateBypass

        # A fresh connect should start with a clean pooled connection rather
        # than reusing whatever the previous session (possibly to a different
        # host, or long stale) left behind.
        $script:ProxmoxWebSession = $null

        $header = @{ Authorization = "PVEAPIToken=$username!$tokenName=$tokenSecret" }

        $script:ProxmoxSession = @{
            host         = $proxmoxHost
            user         = $username
            token_name   = $tokenName
            token_secret = $tokenSecret
            header       = $header
        }

        $resp = Invoke-ProxmoxApi -Method GET -Path '/api2/json/version'
        $version = [string]$resp.data.version

        Write-DebugLog "Connected successfully to $proxmoxHost as $username, version=$version" -Level 'I' -Component '02'

        return @{ result = @{ message = "$($script:ProviderNamePrefix) Connected successfully to Proxmox $proxmoxHost (version: $version)" } }
    }
    catch {
        $script:ProxmoxSession = $null
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to connect to Proxmox: $($_.Exception.Message)"
    }
}

function Handle-Disconnect {
    try {
        $hostName = $null
        if ($null -ne $script:ProxmoxSession) {
            $hostName = $script:ProxmoxSession.host
        }

        $script:ProxmoxSession = $null
        $script:ProxmoxWebSession = $null
        $script:TaskContext = @{}

        # Do NOT delete $script:CloneStatePath here. RAS reconnects immediately after
        # disconnecting -- often from a fresh provider process, where
        # $script:TaskContext above is the ONLY in-memory tracking and is normally empty
        # anyway -- and any clone still in flight depends entirely on this persisted
        # file for Wait-ForRealCloneCompletionBeforeStart to find it and hold off the
        # real Proxmox start call until the real clone job is done. Wiping the file on
        # every disconnect/reconnect cycle is exactly why that guard sometimes never
        # engages. Individual entries already retire themselves correctly on their own.
        return @{ result = @{ message = "Session cleared on Proxmox $hostName" } }
    }
    catch {
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to clear session: $($_.Exception.Message)"
    }
}

function Handle-HostList {
    try {
        # Get-ProxmoxClusterVMs already excludes anything outside pool_scope.pool_name --
        # only rasExclude needs checking here.
        $clusterVMs = Get-ProxmoxClusterVMs
        $hosts = @()

        foreach ($vm in $clusterVMs) {
            $vmId = [string]$vm.vmid
            if (-not [string]::IsNullOrWhiteSpace($vmId)) {
                if (Test-ProxmoxVmHasTag -ClusterVm $vm -Tag $script:RasExcludeTag) {
                    continue
                }
                $hosts += $vmId
            }
        }

        return @{ result = @{ guests = @($hosts) } }
    }
    catch {
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to retrieve host list: $($_.Exception.Message)"
    }
}

function Handle-HostGet {
    param([object]$Params)

    if ($null -eq $Params -or $null -eq $Params.id) {
        return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Invalid or missing host ID"
    }

    try {
        $ids = @($Params.id)

        if ($ids.Count -eq 1) {
            $hostObj = Get-RasGuestObjectForCloneAwareFlow -VmId ([string]$ids[0])
            return @{ result = $hostObj }
        }

        $resultMap = @{}
        foreach ($id in $ids) {
            $vmId = [string]$id
            try {
                $resultMap[$vmId] = Get-RasGuestObjectForCloneAwareFlow -VmId $vmId
            }
            catch {
                $resultMap[$vmId] = @{
                    id            = $vmId
                    # Never $null -- see the note on the single-id path above.
                    name          = "VM-$vmId"
                    provider      = 'Proxmox'
                    state         = 'powered_off'
                    power_state   = 'unknown'
                    host_os       = 'unknown'
                    ip            = $null
                    ip_addresses  = @()
                    mac_addresses = @()
                    is_template   = $false
                    type          = 'Virtual Machine'
                }
                Write-DebugLog "Host get failed for VM [$vmId]: $($_.Exception.Message)" -Level 'E' -Component '03' -Ref $vmId
            }
        }

        return @{ result = $resultMap }
    }
    catch {
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to retrieve host info: $($_.Exception.Message)"
    }
}

function Handle-HostControl {
    param([object]$Params)
    return Handle-GuestControl -Params $Params
}

# RAS never issues stop or delete again for a clone it has permanently lost track of
# (e.g. after its own "Timeout while waiting for guest handle" gives up on one), even
# once the VM is plainly running with an IP. This provider cannot fix that RAS-side gap,
# but it can flag it: throttled to at most once per
# orphan_detection.check_interval_seconds, any VM tagged rasClone<sourceId> that this
# provider is NOT currently tracking as an in-flight clone, AND that RAS has not asked
# about via guests/get in more than orphan_detection.stale_poll_after_seconds (see
# $script:LastGuestPollAt), is logged as an ORPHAN CANDIDATE and tagged
# rasOrphanCandidate, best-effort.
#
# Log-and-tag only -- this NEVER stops or deletes anything on its own. A VM genuinely
# mid a slow but healthy RAS reconciliation could still trip this heuristic, so treat a
# hit as a lead worth checking, not a verdict; an admin removes the tag by hand once
# resolved (this provider never re-clears it).
function Invoke-OrphanAudit {
    param([Parameter(Mandatory = $true)][object[]]$ClusterVMs)

    if (-not $script:OrphanDetectionEnabled) {
        return
    }

    if (([DateTime]::UtcNow - $script:OrphanAuditLastCheckedAt).TotalSeconds -lt $script:OrphanDetectionCheckIntervalSeconds) {
        return
    }

    $script:OrphanAuditLastCheckedAt = [DateTime]::UtcNow

    foreach ($vm in $ClusterVMs) {
        $vmId = [string]$vm.vmid
        if ([string]::IsNullOrWhiteSpace($vmId) -or $script:TrackedCloneVmIds.Contains($vmId)) {
            continue
        }

        # A VM this provider destroyed seconds ago is still in the lagging cluster
        # listing, and Clear-ProxmoxTrackingForVm has already dropped its poll record --
        # which would read as "RAS never asked about it", i.e. maximally stale.
        if (Test-ProxmoxRecentlyDeleted -VmId $vmId) {
            continue
        }

        $tags = Get-ProxmoxVmTagList -ClusterVm $vm
        $cloneTag = $tags | Where-Object { $_ -like "$($script:RasCloneTagPrefix)*" } | Select-Object -First 1
        if ([string]::IsNullOrWhiteSpace($cloneTag)) {
            continue
        }

        $lastPolled = $null
        if ($script:LastGuestPollAt.ContainsKey($vmId)) {
            $lastPolled = [DateTime]$script:LastGuestPollAt[$vmId]
        }

        # "No poll record" means this process cannot judge, NOT "infinitely stale" --
        # every record is per-process and starts empty, so treating its absence as
        # evidence would make the first audit after a restart flag every rasClone-tagged
        # VM in the cluster at once. Measure from process start instead: a VM only
        # qualifies once this provider has been up long enough to have expected to hear
        # about it.
        $referenceTime = if ($null -ne $lastPolled) { $lastPolled } else { $script:ProviderStartedAt }
        $staleSeconds = ([DateTime]::UtcNow - $referenceTime).TotalSeconds
        if ($staleSeconds -lt $script:OrphanDetectionStalePollAfterSeconds) {
            continue
        }

        $lastPolledDesc = if ($null -ne $lastPolled) { "$([int]$staleSeconds)s ago" } else { "never in this provider process (up $([int]$staleSeconds)s)" }
        Write-DebugLog "ORPHAN CANDIDATE: VM [$vmId] is tagged [$cloneTag] but this provider is not tracking it as in-flight, and RAS last polled it via guests/get $lastPolledDesc -- RAS may have permanently lost track of this guest. Manual review recommended." -Level 'W' -Component '04' -Ref $vmId

        if ($tags -notcontains $script:RasOrphanCandidateTag) {
            try {
                Add-ProxmoxVmTag -Node ([string]$vm.node) -VmId $vmId -ClusterVm $vm -Tag $script:RasOrphanCandidateTag
            }
            catch {
                Write-DebugLog "Failed to tag orphan candidate VM [$vmId] with [$($script:RasOrphanCandidateTag)]: $($_.Exception.Message)" -Level 'E' -Component '04' -Ref $vmId
            }
        }
    }
}

function Handle-GuestList {
    try {
        $clusterVMs = Get-ProxmoxClusterVMs
        $guests = @()

        foreach ($vm in $clusterVMs) {
            $vmId = [string]$vm.vmid

            # A cluster/resources entry for a VM mid-clone can genuinely have no 'name'
            # property at all (Proxmox has not propagated it yet). Under Set-StrictMode,
            # [string]$vm.name on such an entry throws PropertyNotFoundException, which
            # this function's catch would turn into a hard JSON-RPC error for the ENTIRE
            # listing. One missing property must never fail the whole listing.
            $vmName = if ((Get-MemberNames -Object $vm) -contains 'name') { [string]$vm.name } else { '' }

            if (-not [string]::IsNullOrWhiteSpace($vmId)) {
                $isTrackedClone = $script:TrackedCloneVmIds.Contains($vmId)

                if ($isTrackedClone) {
                    # Gate on OUR OWN completion signal rather than on Proxmox's NAME
                    # for the VM: Proxmox can propagate the real name to
                    # cluster/resources before the clone job is done, so a tracked clone
                    # could otherwise appear here seconds before its tasks/get poll ever
                    # reports it 'completed'. A listing round-trip from RAS's own sync
                    # walker landing in that gap can bind the guest to a stale record
                    # before RAS's clone thread ever gets a chance to. So a tracked
                    # clone is only listed once its clone task has actually been
                    # reported 'completed' (Handle-TaskInfo).
                    $trackedEntry = Get-CloneStateEntry -VmId $vmId
                    $reportedCompleted = ($null -ne $trackedEntry -and
                        (Get-MemberNames -Object $trackedEntry) -contains 'reported_completed_at' -and
                        -not [string]::IsNullOrWhiteSpace([string]$trackedEntry.reported_completed_at))

                    if (-not $reportedCompleted) {
                        Write-DebugLog "Skipping guest [$vmId] because its clone task has not yet been reported completed to RAS." -Level 'T' -Component '03' -Ref $vmId
                        continue
                    }
                }
                else {
                    # Skip placeholder names like "VM 101", which indicate cloning is
                    # still in progress on the hypervisor for a VM this provider is NOT
                    # (or no longer) tracking as its own clone -- e.g. a clone started
                    # outside this provider.
                    $isAutoName = Test-ProxmoxPlaceholderVmName -VmId $vmId -VmName $vmName

                    if ($isAutoName) {
                        Write-DebugLog "Skipping guest [$vmId] because name [$vmName] matches the placeholder VM <id> pattern." -Level 'T' -Component '03' -Ref $vmId
                        continue
                    }
                }

                if (Test-ProxmoxRecentlyDeleted -VmId $vmId) {
                    Write-DebugLog "Skipping guest [$vmId] because it was recently deleted by this provider." -Level 'T' -Component '03' -Ref $vmId
                    continue
                }

                # pool_scope is already applied by Get-ProxmoxClusterVMs -- $clusterVMs
                # never contains an out-of-scope VM in the first place, so only rasExclude
                # needs checking here.
                if (Test-ProxmoxVmHasTag -ClusterVm $vm -Tag $script:RasExcludeTag) {
                    Write-DebugLog "Skipping guest [$vmId] because it carries the $($script:RasExcludeTag) tag." -Level 'T' -Component '03' -Ref $vmId
                    continue
                }

                $guests += $vmId
            }
        }

        Invoke-OrphanAudit -ClusterVMs $clusterVMs

        return @{ result = @{ guests = @($guests) } }
    }
    catch {
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to retrieve guest list: $($_.Exception.Message)"
    }
}

function Handle-GuestGet {
    param([object]$Params)

    if ($null -eq $Params -or $null -eq $Params.id) {
        return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Invalid or missing guest ID"
    }

    try {
        $ids = @($Params.id)

        if ($ids.Count -eq 1) {
            $vmId = [string]$ids[0]

            # Skip a guaranteed-to-fail PVE lookup for an id this provider itself just
            # destroyed, and report it gone directly rather than surfacing whatever
            # error Proxmox happens to return for a vanished VMID.
            if (Test-ProxmoxRecentlyDeleted -VmId $vmId) {
                return @{
                    result = @{
                        id            = $vmId
                        # NOT $null. RAS's deserializer hard-requires 'name' and rejects
                        # the WHOLE object without it -- "Cannot find name in {...}" --
                        # so a null here tells RAS nothing at all, not even the state we
                        # went to the trouble of filling in. Echo the name this VM had
                        # when THIS provider deleted it (falls back to "VM-<id>" only if
                        # that was never recorded) rather than a synthetic placeholder --
                        # RAS treats this response as fresh guest info and updates its own
                        # displayed name from it immediately, so during a bulk recreate a
                        # poll landing in the delete-to-reclone gap must not rename the
                        # guest to something VM-ID-shaped. See
                        # $script:RecentlyDeletedNames.
                        name          = Get-ProxmoxRecentlyDeletedName -VmId $vmId
                        provider      = 'Proxmox'
                        state         = 'powered_off'
                        power_state   = 'unknown'
                        host_os       = 'unknown'
                        ip            = $null
                        ip_addresses  = @()
                        mac_addresses = @()
                        is_template   = $false
                        type          = 'Virtual Machine'
                    }
                }
            }

            $guest = Get-RasGuestObjectForCloneAwareFlow -VmId $vmId
            Invoke-TrackedCloneSweep -ExcludeVmId $vmId
            return @{ result = $guest }
        }

        $resultMap = @{}
        foreach ($id in $ids) {
            $vmId = [string]$id

            if (Test-ProxmoxRecentlyDeleted -VmId $vmId) {
                $resultMap[$vmId] = @{
                    id            = $vmId
                    # Never $null -- see the note on the single-id path above.
                    name          = Get-ProxmoxRecentlyDeletedName -VmId $vmId
                    provider      = 'Proxmox'
                    state         = 'powered_off'
                    power_state   = 'unknown'
                    host_os       = 'unknown'
                    ip            = $null
                    ip_addresses  = @()
                    mac_addresses = @()
                    is_template   = $false
                    type          = 'Virtual Machine'
                }
                continue
            }

            try {
                $resultMap[$vmId] = Get-RasGuestObjectForCloneAwareFlow -VmId $vmId
            }
            catch {
                $resultMap[$vmId] = @{
                    id            = $vmId
                    # Never $null -- see the note on the single-id path above.
                    name          = "VM-$vmId"
                    provider      = 'Proxmox'
                    state         = 'powered_off'
                    power_state   = 'unknown'
                    host_os       = 'unknown'
                    ip            = $null
                    ip_addresses  = @()
                    mac_addresses = @()
                    is_template   = $false
                    type          = 'Virtual Machine'
                }
                Write-DebugLog "Guest get failed for VM [$vmId]: $($_.Exception.Message)" -Level 'E' -Component '03' -Ref $vmId
            }
        }

        # Batch form already covers every id RAS asked about; the sweep exists for the
        # single-id path above, which is the one RAS actually uses.
        return @{ result = $resultMap }
    }
    catch {
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to retrieve guest info: $($_.Exception.Message)"
    }
}

function Handle-GuestControl {
    param([object]$Params)

    if ($null -eq $Params -or [string]::IsNullOrWhiteSpace([string]$Params.id)) {
        return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Invalid guest id"
    }

    if ($null -eq $Params.control -or [string]::IsNullOrWhiteSpace([string]$Params.control)) {
        return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Invalid guest control"
    }

    try {
        $vmId = [string]$Params.id
        $requestedControl = [string]$Params.control
        $action = Get-ControlAction -Control $requestedControl

        if ([string]::IsNullOrWhiteSpace($action)) {
            return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Unsupported guest control: $requestedControl"
        }

        if ($action -eq 'start') {
            # RAS issues this explicit start as soon as tasks/get reports a freshly
            # cloned guest 'completed' -- which, under pipelined clone completion, can
            # happen before Proxmox's real clone job is done. Calling Proxmox's start
            # endpoint on a VM still held by its own clone lock would fail, so wait for
            # the REAL completion.
            $waitResult = Wait-ForRealCloneCompletionBeforeStart -VmId $vmId
            if ($waitResult.shouldWait -and -not $waitResult.ok) {
                return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) $($waitResult.message)"
            }
        }

        $clusterVm = Get-ProxmoxVmNode -VmId $vmId
        $node = [string]$clusterVm.node

        if ($action -eq 'start') {
            try {
                $current = Get-ProxmoxVmCurrentStatus -Node $node -VmId $vmId
                $rawState = ''

                if ((Get-MemberNames -Object $current) -contains 'qmpstatus' -and -not [string]::IsNullOrWhiteSpace([string]$current.qmpstatus)) {
                    $rawState = [string]$current.qmpstatus
                }
                elseif ((Get-MemberNames -Object $current) -contains 'status' -and -not [string]::IsNullOrWhiteSpace([string]$current.status)) {
                    $rawState = [string]$current.status
                }

                if ($rawState.Trim().ToLowerInvariant() -eq 'paused') {
                    $action = 'resume'
                    Write-DebugLog "Guest control start remapped to resume for paused VM [$vmId]." -Level 'T' -Component '05' -Ref $vmId
                }
            }
            catch {
                Write-DebugLog "Failed to get current status for VM [$vmId] before start control. Falling back to start: $($_.Exception.Message)" -Level 'W' -Component '05' -Ref $vmId
            }
        }

        if ($action -eq 'delete') {
            $stopResult = Stop-ProxmoxVmHardBeforeDelete -Node $node -VmId $vmId
            if (-not $stopResult.ok) {
                return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) $($stopResult.message)"
            }
            $resp = Invoke-ProxmoxApiWithRetry -Method DELETE -Path "/api2/json/nodes/$node/qemu/$vmId"

            # A clone deleted before it ever reaches IP-ready would otherwise leave its
            # clone-state entry (and concurrency slot) stuck forever, since only
            # Get-RasGuestObjectForCloneAwareFlow removes one, and only on confirmed
            # readiness. And only guests/list retires a guest on the RAS side, so mark
            # the id recently-deleted too, so Handle-GuestList omits it immediately.
            Clear-ProxmoxTrackingForVm -VmId $vmId
            $script:RecentlyDeletedIds[$vmId] = [DateTime]::UtcNow

            # $clusterVm was fetched above, before this delete, so its 'name' is the VM's
            # real name -- remember it for the recently-deleted stub in Handle-GuestGet.
            # RAS treats every guests/get response as fresh guest info and updates its own
            # displayed name from it immediately, so a bulk recreate that happens to poll
            # this id during the delete-to-reclone gap must not hand back a synthetic
            # "VM-<id>" -- see $script:RecentlyDeletedNames.
            if ((Get-MemberNames -Object $clusterVm) -contains 'name' -and -not [string]::IsNullOrWhiteSpace([string]$clusterVm.name)) {
                $deletedName = [string]$clusterVm.name
                $script:RecentlyDeletedNames[$vmId] = $deletedName
                # $node (this VM's node, resolved above before the delete) is what a
                # recreation of this same name should land back on -- see
                # placement.preserve_node_on_recreation. Deliberately keyed on name, not
                # vmid: cluster/nextid often does reuse the same id (see the comment in
                # Handle-GuestClone), but that is not guaranteed, while RAS reliably reuses
                # the same guest name across a pool recreation.
                $script:RecentlyDeletedNodeByName[$deletedName] = @{ node = $node; deleted_at = [DateTime]::UtcNow }

                # cloning.mac_preservation.enabled -- see MAC-PRESERVATION.md. Off by
                # default, and the live config read this costs is skipped entirely when
                # it's off -- no cost to a deployment that doesn't use this. Best-effort:
                # a failed read here must never block or fail the delete itself, same
                # reasoning as every other best-effort write in this script.
                if ($script:PreserveMacOnRecreation) {
                    try {
                        $macsAtDelete = Get-ProxmoxVmNetIfaceMacs -Node $node -VmId $vmId
                        if (@(Get-MemberNames -Object $macsAtDelete).Count -gt 0) {
                            $script:RecentlyDeletedMacByName[$deletedName] = @{ macs = $macsAtDelete; deleted_at = [DateTime]::UtcNow }
                        }
                    }
                    catch {
                        Write-DebugLog "Failed to capture MAC address(es) for VM [$vmId] before delete -- recreation under this name will get a fresh Proxmox-assigned MAC instead of preserving the old one: $($_.Exception.Message)" -Level 'W' -Component '04' -Ref $vmId
                    }
                }
            }
            else {
                $script:RecentlyDeletedNames.Remove($vmId)
            }
        }
        else {
            $resp = Invoke-ProxmoxApiWithRetry -Method POST -Path "/api2/json/nodes/$node/qemu/$vmId/status/$action" -Body @{}
        }

        # The cluster-resources cache is otherwise only invalidated by its TTL. Busting
        # it after any successful control action means the very next guests/get or
        # guests/list reflects this change instead of serving a listing that predates it
        # -- without this, a concurrent RAS 'start' and 'delete' on the same guest can
        # disagree on its power state.
        Reset-ProxmoxClusterCache

        # Busting the cache above only guarantees the NEXT fetch is fresh over the wire;
        # it does not make Proxmox's own server-side cluster/resources aggregate
        # current. For a deleted VM this is moot (Test-ProxmoxRecentlyDeleted intercepts
        # it first); for every other action, ConvertTo-RasGuestObject uses this to
        # prefer a live status/current read over the lagging listing.
        if ($action -ne 'delete') {
            $script:RecentlyControlledIds[$vmId] = [DateTime]::UtcNow
        }

        # Narrower marker for the guest-agent-probe race above -- only a 'stop' sets it,
        # and any other action clears it right away so a quick restart is never left
        # suppressed by a stale stop.
        if ($action -eq 'stop') {
            $script:RecentlyStoppedIds[$vmId] = [DateTime]::UtcNow
        }
        elseif ($script:RecentlyStoppedIds.ContainsKey($vmId)) {
            $script:RecentlyStoppedIds.Remove($vmId)
        }

        # A 'start' here confirms whatever de-templated this vmid was entering
        # maintenance, not RAS deleting the Template object -- cancel the pending
        # rasTemplate<id> tag removal outright. See $script:PendingTemplateTagRemovalIds.
        if ($action -eq 'start' -and $script:PendingTemplateTagRemovalIds.ContainsKey($vmId)) {
            $script:PendingTemplateTagRemovalIds.Remove($vmId)
            Write-DebugLog "guests/control(start) for VM [$vmId] confirms maintenance, not template deletion -- keeping its rasTemplate tag." -Level 'I' -Component '05' -Ref $vmId
        }

        $upid = $null
        if ($null -ne $resp -and (Get-MemberNames -Object $resp) -contains 'data') {
            $upid = $resp.data
        }

        return @{
            result = @{
                id      = $vmId
                node    = $node
                action  = $action
                upid    = $upid
                message = "$($script:ProviderNamePrefix) Guest control [$requestedControl] submitted successfully"
            }
        }
    }
    catch {
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to control guest [$($Params.control)]: $($_.Exception.Message)"
    }
}

# Reports a clone task 'completed' to RAS. Called from Handle-TaskInfo's clone branch
# either early (the pipelining shortcut, full clones only, once the elapsed-time and
# concurrency gates allow it) or as soon as Proxmox's own task is confirmed done (every
# clone, unconditionally) -- both cases converge here so the two paths cannot drift
# apart. Per the CPF integration guide, tasks/get answers with task state only; whether a
# guest is actually powered on is guests/get's question, driven by RAS's own subsequent
# guests/control(start), not something this response encodes or waits for.
#
# Deliberately does NOT touch the VmId-level clone-state entry (Get-CloneStateEntry /
# Remove-CloneStateEntry) -- only the task-level entry in $script:TaskContext. The VmId
# entry stays in place so ConvertTo-RasGuestObject's name substitution keeps working for
# as long as Proxmox's cluster/resources view might still show a stale placeholder name,
# independent of what tasks/get has already told RAS. Get-ActiveCloneCount is unaffected
# either way -- it re-checks the
# real Proxmox task itself (see $script:CloneTaskCompletionCache) rather than relying on
# this entry's presence.
function Complete-ProxmoxCloneTask {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TaskId,

        [string]$CloneId
    )

    if ([string]::IsNullOrWhiteSpace($CloneId)) {
        Write-DebugLog "Clone task [$TaskId] has no clone_id. Completing without output." -Level 'I' -Component '08' -Ref $TaskId
        return @{ result = @{ state = 'completed'; output = @{} } }
    }

    if ($script:TaskContext.ContainsKey($TaskId)) {
        [void]$script:TaskContext.Remove($TaskId)
    }

    # Handle-GuestList gates a tracked clone's visibility on this, not on whatever name
    # Proxmox happens to be reporting at the moment -- see that function's tracked-clone
    # branch.
    Set-CloneReportedCompleted -VmId $CloneId

    return @{
        result = @{
            state  = 'completed'
            output = @{ clone_id = $CloneId }
        }
    }
}

function Handle-TaskInfo {
    param([object]$Params)

    if ($null -eq $Params -or [string]::IsNullOrWhiteSpace([string]$Params.id)) {
        return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Invalid task id"
    }

    try {
        $taskId = [string]$Params.id

        if (-not [string]::IsNullOrWhiteSpace([string]$script:DummyOperationTaskId) -and
            $taskId.StartsWith([string]$script:DummyOperationTaskId, [System.StringComparison]::Ordinal)) {
            Write-DebugLog "Dummy task [$taskId] requested. Returning completed state." -Level 'D' -Component '08' -Ref $taskId
            return @{
                result = @{
                    state  = 'completed'
                    output = @{}
                }
            }
        }

        # A DIFFERENT poll (guests/get's own clone-aware flow, including the
        # opportunistic sweep) can win the race to confirm this clone ready and clear
        # its tracking before this tasks/get call arrives, leaving neither
        # $script:TaskContext nor Get-CloneStateEntryByTaskId below anything to find.
        # Without this check, execution would fall through to the generic "completed, no
        # output" response at the bottom -- correct for a non-clone task, but RAS
        # expects output.clone_id on a clone task and does not retry once it gets an
        # answer.
        $stashedOutput = Get-CompletedCloneTaskOutput -TaskId $taskId
        if ($null -ne $stashedOutput) {
            Write-DebugLog "Clone task [$taskId]: tracking was already cleared by a concurrent guests/get, but this task's clone_id [$($stashedOutput.clone_id)] was recorded -- reporting completed correctly instead of an empty output." -Level 'I' -Component '08' -Ref $taskId
            return @{
                result = @{
                    state  = 'completed'
                    output = @{ clone_id = $stashedOutput.clone_id }
                }
            }
        }

        $taskStatus = Get-ProxmoxTaskStatus -TaskId $taskId
        $taskResult = New-TaskResultState -TaskStatus $taskStatus

        if ($taskResult.state -eq 'failed') {
            return @{
                result = @{
                    state = 'failed'
                    error = $taskResult.error
                }
            }
        }

        # Clone tracking is looked up BEFORE the 'running' early-return below: pipelined
        # completion needs to short-circuit a clone task that is still genuinely
        # 'running' on Proxmox's side, not just the ones Proxmox already finished.
        $ctx = $null

        if ($script:TaskContext.ContainsKey($taskId)) {
            $ctx = $script:TaskContext[$taskId]
            Write-DebugLog "TASK CONTEXT FOUND IN MEMORY for task [$taskId]" -Level 'D' -Component '08' -Ref $taskId
        }
        else {
            $ctx = Get-CloneStateEntryByTaskId -TaskId $taskId
            if ($null -ne $ctx) {
                $script:TaskContext[$taskId] = $ctx
            }
        }

        $isCloneCtx = ($null -ne $ctx -and $ctx.ContainsKey('type') -and [string]$ctx.type -eq 'clone')

        # Pipelining exists to hide a FULL clone's disk copy (measured 128.7s mean, see
        # PIPELINED-CLONING.md) behind an early 'completed' answer -- there is nothing to
        # hide for a linked clone, whose real Proxmox task finishes in under a second
        # (LINKED-CLONES.md). In practice that means the elapsed-time gate below never
        # even gets a chance to fire for a linked clone: the real completion is visible on
        # the very first poll, before $script:PipelinedCloneCompletionSeconds can elapse --
        # so this isn't a behavior change for linked clones, just an explicit skip of a
        # timer that could never have won anyway. `ctx.full` is stamped by Handle-GuestClone
        # at clone time (see the 'full = -not $isLinked' entries) -- default to treating an
        # ambiguous/missing value as a full clone, preserving today's behavior when in doubt.
        $ctxIsFullClone = $true
        if ($isCloneCtx -and $ctx.ContainsKey('full')) {
            try { $ctxIsFullClone = [System.Convert]::ToBoolean($ctx.full) } catch { $ctxIsFullClone = $true }
        }

        if ($isCloneCtx -and $script:PipelinedCloneCompletionEnabled -and $ctxIsFullClone -and
            $ctx.ContainsKey('clone_started_at') -and -not [string]::IsNullOrWhiteSpace([string]$ctx.clone_started_at)) {

            $cloneElapsedSeconds = $null
            try { $cloneElapsedSeconds = [int]([DateTime]::UtcNow - [DateTime]::Parse([string]$ctx.clone_started_at).ToUniversalTime()).TotalSeconds } catch { }

            $elapsedOk = ($null -ne $cloneElapsedSeconds -and $cloneElapsedSeconds -ge $script:PipelinedCloneCompletionSeconds)

            if ($elapsedOk) {
                # Second gate: even past the elapsed-time threshold, only pipeline this
                # 'completed' if there is genuinely still room for ANOTHER clone
                # afterwards. Reporting 'completed' is exactly what triggers RAS to
                # submit the next real guests/clone -- and Handle-GuestClone itself
                # checks no limit before issuing it to Proxmox -- so this check must use
                # a STRICT less-than, not <=. This entry is itself one of the counted
                # active clones, so "active count < limit" means "there is room for one
                # more"; with <=, two clones both pipeline through at count == limit and
                # RAS's next guests/clone starts a real job past the cap.
                $activeCloneCount = Get-ActiveCloneCount
                if ($activeCloneCount -ge $script:MaxConcurrentCloneOperations) {
                    Write-DebugLog "Clone task [$taskId]: elapsed ${cloneElapsedSeconds}s (threshold met) but $activeCloneCount clone(s) currently active (limit $($script:MaxConcurrentCloneOperations)) -- holding off pipelined completion until a slot frees up." -Level 'T' -Component '08' -Ref $taskId
                }
                else {
                    Write-DebugLog "Clone task [$taskId]: PIPELINED completion after ${cloneElapsedSeconds}s (threshold $($script:PipelinedCloneCompletionSeconds)s), $activeCloneCount/$($script:MaxConcurrentCloneOperations) concurrent clone(s) -- reporting completed for full clone VM [$([string]$ctx.clone_id)] ahead of Proxmox's own disk-copy job, purely so RAS can submit the next clone sooner." -Level 'I' -Component '08' -Ref $taskId
                    return Complete-ProxmoxCloneTask -TaskId $taskId -CloneId ([string]$ctx.clone_id)
                }
            }
        }

        if ($taskResult.state -eq 'running') {
            return @{ result = @{ state = 'running' } }
        }

        # Proxmox's own task is genuinely done here (not 'running' -- checked just above;
        # not 'failed' -- checked earlier). Per the CPF integration guide, tasks/get
        # answers with task state only. Whether the guest is actually powered on and
        # reachable is guests/get's question -- Get-RasGuestObjectForCloneAwareFlow
        # already tracks that independently, on every poll, regardless of what this
        # function says -- driven from here on by RAS's own guests/control(start), not by
        # this provider guessing on RAS's behalf. This function used to block 'completed'
        # here until the guest was powered on with a real IP (see PIPELINED-CLONING.md for
        # why that existed and how pipelining partially worked around it); it no longer
        # does, for either clone type.
        if ($isCloneCtx) {
            Write-DebugLog "Clone task [$taskId]: real Proxmox task confirmed completed -- reporting completed for clone VM [$([string]$ctx.clone_id)]." -Level 'I' -Component '08' -Ref $taskId
            return Complete-ProxmoxCloneTask -TaskId $taskId -CloneId ([string]$ctx.clone_id)
        }

        return @{
            result = @{
                state  = 'completed'
                output = @{}
            }
        }
    }
    catch {
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to retrieve task info: $($_.Exception.Message)"
    }
}

function Handle-GuestConvert {
    param([object]$Params)

    if ($null -eq $Params -or [string]::IsNullOrWhiteSpace([string]$Params.id)) {
        return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Invalid guest id"
    }

    if ($null -eq $Params.is_template) {
        return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Missing is_template flag"
    }

    try {
        $vmId = [string]$Params.id
        $isTemplate = [bool]$Params.is_template

        $clusterVm = Get-ProxmoxVmNode -VmId $vmId
        $node = [string]$clusterVm.node
        $resp = $null

        # RAS drives template maintenance mode entirely through this method: enter =
        # convert to VM and boot it, exit = shut down and convert back. Both directions
        # are retried whenever RAS's own view says the flag did not move, and that view
        # comes from OUR guests/get -- so this has to be idempotent, or a retry lands on
        # Proxmox as "you can't convert a template to a template" (a 500) and RAS
        # reports the whole maintenance exit failed. Read the flag from the VM's own
        # config (pmxcfs, no aggregation lag) and no-op if it is already where RAS wants
        # it.
        $alreadyInRequestedState = $false
        try {
            $liveConfig = Get-ProxmoxVmConfig -Node $node -VmId $vmId
            $liveIsTemplate = $false
            if ($null -ne $liveConfig -and (Get-MemberNames -Object $liveConfig) -contains 'template' -and $null -ne $liveConfig.template) {
                $liveIsTemplate = ([int]$liveConfig.template -eq 1)
            }
            $alreadyInRequestedState = ($liveIsTemplate -eq $isTemplate)
        }
        catch {
            Write-DebugLog "Live template-flag check failed for VM [$vmId], attempting the convert anyway: $($_.Exception.Message)" -Level 'W' -Component '06' -Ref $vmId
        }

        if ($alreadyInRequestedState) {
            Write-DebugLog "Convert of VM [$vmId] to is_template=[$isTemplate] skipped -- Proxmox already reports that state." -Level 'I' -Component '06' -Ref $vmId
            Set-ProxmoxRecentConvert -VmId $vmId -IsTemplate $isTemplate
            Reset-ProxmoxClusterCache
            if ($isTemplate) { Set-ProxmoxTemplateSourceTagBestEffort -Node $node -VmId $vmId -ClusterVm $clusterVm -Present $true }
            return @{ result = @{ task_id = [string]$script:DummyOperationTaskId } }
        }

        try {
            if ($isTemplate) {
                $resp = Invoke-ProxmoxApi -Method POST -Path "/api2/json/nodes/$node/qemu/$vmId/template" -Body @{}
            }
            else {
                $resp = Invoke-ProxmoxApi -Method PUT -Path "/api2/json/nodes/$node/qemu/$vmId/config" -Body @{ template = 0 }
            }
        }
        catch {
            # Narrow safety net for the race the check above cannot close: the flag
            # flipping between our read and our write. Only this one Proxmox message is
            # swallowed, and only into the state RAS asked for; every other failure
            # still surfaces.
            if ($isTemplate -and [string]$_.Exception.Message -match "can't convert a template to a template") {
                Write-DebugLog "Convert of VM [$vmId] raced -- Proxmox reports it is already a template, treating as success." -Level 'W' -Component '06' -Ref $vmId
                Set-ProxmoxRecentConvert -VmId $vmId -IsTemplate $true
                Reset-ProxmoxClusterCache
                Set-ProxmoxTemplateSourceTagBestEffort -Node $node -VmId $vmId -ClusterVm $clusterVm -Present $true
                return @{ result = @{ task_id = [string]$script:DummyOperationTaskId } }
            }
            throw
        }

        # Both the flag and (for the VM->template direction) the disk layout
        # have just changed underneath the cached listing.
        Set-ProxmoxRecentConvert -VmId $vmId -IsTemplate $isTemplate
        Reset-ProxmoxClusterCache
        Write-DebugLog "Converted VM [$vmId] to is_template=[$isTemplate]." -Level 'I' -Component '06' -Ref $vmId
        if ($isTemplate) {
            Set-ProxmoxTemplateSourceTagBestEffort -Node $node -VmId $vmId -ClusterVm $clusterVm -Present $true
        }
        else {
            # A genuine template->VM transition (not a retried no-op -- see the
            # alreadyInRequestedState branch above, which deliberately does NOT arm this)
            # -- arm the pending-removal marker. Resolve-PendingTemplateTagRemoval clears
            # the tag once the grace window passes with no guests/control start; a start
            # for this vmid cancels it outright (Handle-GuestControl).
            $script:PendingTemplateTagRemovalIds[$vmId] = [DateTime]::UtcNow
        }

        $taskId = if ($isTemplate) { [string]$resp.data } else { [string]$script:DummyOperationTaskId }
        if (-not [string]::IsNullOrWhiteSpace($taskId)) {
            $script:TaskContext[$taskId] = @{
                type        = 'convert'
                id          = $vmId
                is_template = $isTemplate
            }
        }

        return @{ result = @{ task_id = $taskId } }
    }
    catch {
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to convert guest [$($Params.id)]: $($_.Exception.Message)"
    }
}

# ---------------------------------------------------------------------------
# Snapshots (guests/snapshots/*)
#
# RAS's linked-clone contract is snapshot-shaped -- it believes it is creating,
# checking, and deleting a named guest snapshot before/after templating. Proxmox has no
# equivalent object here, and does not need one: this provider's own native templates
# already provide the copy-on-write base a linked clone reads from. So these four
# methods keep one invariant and nothing else: the RAS template snapshot EXISTS if and
# only if the guest is presently a native Proxmox template. See
# LINKED-CLONES-DESIGN.md #2 "The virtual snapshot" for the full reasoning, including
# why this makes 'delete' a no-op in practice and why 'revert' errors rather than
# silently no-ops.
#
# All three that touch a real guest read the LIVE config (Get-ProxmoxVmConfig), never
# the cached cluster listing -- the same freshness rule Handle-GuestConvert already
# follows, and for the same reason: a stale is_template here would make 'exists' wrong
# at exactly the moment RAS is deciding whether to issue a delete.
# ---------------------------------------------------------------------------

function Handle-GuestSnapshotsCreate {
    param([object]$Params)

    if ($null -eq $Params -or [string]::IsNullOrWhiteSpace([string]$Params.id)) {
        return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Invalid guest id"
    }
    if ([string]::IsNullOrWhiteSpace([string]$Params.name)) {
        return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Invalid snapshot name"
    }

    # Nothing is created on Proxmox -- 'RAS Template Snapshot' contains spaces and is
    # not a legal pve-configid, and there is nothing for a real snapshot to add on top
    # of the template's own base image. Recorded only so the log reads sensibly and
    # 'exists' has something to reference; the invariant itself is carried entirely by
    # the guest's live template flag, not by this record.
    $vmId = [string]$Params.id
    $name = [string]$Params.name
    $script:TemplateSnapshotIntent[$vmId] = $name
    Write-DebugLog "guests/snapshots/create for VM [$vmId] name [$name] -- no Proxmox call; truth is the guest's own template flag (see guests/snapshots/exists)." -Level 'I' -Component '07' -Ref $vmId
    return @{ result = @{ task_id = [string]$script:DummyOperationTaskId } }
}

function Handle-GuestSnapshotsExists {
    param([object]$Params)

    if ($null -eq $Params -or [string]::IsNullOrWhiteSpace([string]$Params.id)) {
        return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Invalid guest id"
    }
    if ([string]::IsNullOrWhiteSpace([string]$Params.name)) {
        return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Invalid snapshot name"
    }

    $vmId = [string]$Params.id
    try {
        # Must be a BARE boolean. Read-ResultObject in the RAS test kit returns 'result'
        # directly, so {"result": {"exists": true}} would be truthy regardless of its
        # inner value and silently break the delete gate in RAS's exit-maintenance flow.
        $isTemplate = $false
        $clusterVm = Get-ProxmoxVmNode -VmId $vmId
        $node = [string]$clusterVm.node
        $liveConfig = Get-ProxmoxVmConfig -Node $node -VmId $vmId
        if ($null -ne $liveConfig -and (Get-MemberNames -Object $liveConfig) -contains 'template' -and $null -ne $liveConfig.template) {
            $isTemplate = ([int]$liveConfig.template -eq 1)
        }
        Write-DebugLog "guests/snapshots/exists for VM [$vmId]: [$isTemplate] (live template flag)." -Level 'T' -Component '07' -Ref $vmId
        return @{ result = $isTemplate }
    }
    catch {
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to check snapshot on guest [$vmId]: $($_.Exception.Message)"
    }
}

function Handle-GuestSnapshotsDelete {
    param([object]$Params)

    if ($null -eq $Params -or [string]::IsNullOrWhiteSpace([string]$Params.id)) {
        return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Invalid guest id"
    }
    if ([string]::IsNullOrWhiteSpace([string]$Params.name)) {
        return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Invalid snapshot name"
    }

    # A no-op on Proxmox in practice: the only thing that can make 'exists' true is the
    # guest's own template flag, and only guests/convert owns that -- there is no
    # Proxmox object to delete. Still implemented for real, because a differently-ordered
    # RAS build could call it, and the bookkeeping record should not outlive it either
    # way.
    #
    # This call IS, however, the reliable signal that RAS is deleting the Template
    # object entirely (not merely entering/exiting maintenance) -- see the "three flows"
    # in LINKED-CLONES.md §5. Entering maintenance never calls any snapshot method at
    # all under template_method=basic, and exiting maintenance's own "exists -> delete"
    # step is gated on the SAME live template flag Handle-GuestSnapshotsExists reads --
    # since exiting maintenance only runs while the guest is already de-templated (RAS
    # asserts NOT is_template before starting), that check reports false and the delete
    # branch is structurally unreachable there. So a delete that actually arrives means
    # the guest was still templated when RAS called it -- i.e. a genuine template
    # deletion -- and this is where the rasTemplate<id> tag comes off, so the machine
    # is fully cleaned up rather than left carrying a stale "this was a RAS template"
    # marker. (This only fires when linked clones are in use -- with no snapshot object
    # ever registered, a plain guests/convert(false) delete-vs-maintenance ambiguity has
    # no wire-level signal to resolve it at all.)
    $vmId = [string]$Params.id
    if ($script:TemplateSnapshotIntent.ContainsKey($vmId)) { $script:TemplateSnapshotIntent.Remove($vmId) }
    try {
        $clusterVm = Get-ProxmoxVmNode -VmId $vmId
        Set-ProxmoxTemplateSourceTagBestEffort -Node ([string]$clusterVm.node) -VmId $vmId -ClusterVm $clusterVm -Present $false
    }
    catch {
        Write-DebugLog "Could not resolve VM [$vmId] to clear its rasTemplate tag on snapshot delete: $($_.Exception.Message)" -Level 'E' -Component '07' -Ref $vmId
    }
    Write-DebugLog "guests/snapshots/delete for VM [$vmId] name [$($Params.name)] -- no Proxmox call." -Level 'I' -Component '07' -Ref $vmId
    return @{ result = @{ task_id = [string]$script:DummyOperationTaskId } }
}

function Handle-GuestSnapshotsRevert {
    param([object]$Params)

    if ($null -eq $Params -or [string]::IsNullOrWhiteSpace([string]$Params.id)) {
        return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Invalid guest id"
    }

    # Only reachable under template_method='versioning' (see the Framework Test Kit's
    # Test-CreateTemplate.ps1) -- this provider advertises 'basic'. There is no Proxmox
    # state to revert to, so answering success would be a quiet lie. Erroring loudly
    # instead means a real appearance of this call in a log is a signal the contract
    # differs from what the test kit models, not a silently wrong guest.
    $vmId = [string]$Params.id
    Write-DebugLog "guests/snapshots/revert requested for VM [$vmId] name [$($Params.name)] -- not supported at template_method=basic." -Level 'W' -Component '07' -Ref $vmId
    return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Snapshot revert is not supported (template_method=basic has no versioned state to revert to)"
}

function Handle-GuestClone {
    param([object]$Params)

    if ($null -eq $Params -or [string]::IsNullOrWhiteSpace([string]$Params.id)) {
        return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Invalid source guest id"
    }

    if ([string]::IsNullOrWhiteSpace([string]$Params.name)) {
        return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Invalid clone name"
    }

    try {
        $sourceVmId = [string]$Params.id
        $cloneName = [string]$Params.name
        $clusterVm = Get-ProxmoxVmNode -VmId $sourceVmId
        $node = [string]$clusterVm.node

        # MUST run before the clone POST below, not after: Proxmox holds
        # lock-<source>.conf for the entire duration of the clone job, so a config PUT
        # issued afterward blocks on that lock for the full HTTP timeout and fails, on
        # every clone. The source is unlocked right up until the POST starts the job.
        # Also tried at most once per source per PROCESS lifetime (not once per clone)
        # via $script:RasTemplateTagAttempted, so a source already tagged or already
        # known to have failed does not keep re-attempting. Best-effort and non-blocking
        # -- a failed tag write never fails the clone itself.
        $templateTag = "$($script:RasTemplateTagPrefix)$sourceVmId"
        if (-not (Test-ProxmoxVmHasTag -ClusterVm $clusterVm -Tag $templateTag) -and
            -not $script:RasTemplateTagAttempted.Contains($sourceVmId)) {
            try {
                Add-ProxmoxVmTag -Node $node -VmId $sourceVmId -ClusterVm $clusterVm -Tag $templateTag
                Write-DebugLog "Tagged source VM [$sourceVmId] with [$templateTag]." -Level 'I' -Component '04' -Ref $sourceVmId
            }
            catch {
                Write-DebugLog "Failed to tag source VM [$sourceVmId] with [$templateTag]: $($_.Exception.Message)" -Level 'E' -Component '04' -Ref $sourceVmId
            }
            finally {
                [void]$script:RasTemplateTagAttempted.Add($sourceVmId)
            }
        }

        $newVmId = Get-ProxmoxNextVmId

        # cluster/nextid returns the LOWEST free id, so it hands back the id of a VM
        # this provider deleted moments ago -- and RAS recreating a pool does exactly
        # that, delete five then clone five. The delete marker set by Handle-GuestControl
        # would then suppress the BRAND NEW VM for the rest of its 300s retention:
        # guests/list omits it, and guests/get answers with the name=$null "it's gone"
        # stub, which RAS cannot deserialize at all ("Cannot find name in ..."), so RAS
        # stops polling the clone it was just handed and waits for the next listing to
        # rediscover it. Observed cost: up to 114s of phantom "Cloning" per VM.
        # This provider allocated the id, so it knows the id now denotes a live VM --
        # retract the marker here rather than waiting for it to expire.
        if ($script:RecentlyDeletedIds.ContainsKey($newVmId)) {
            $script:RecentlyDeletedIds.Remove($newVmId)
            $script:RecentlyDeletedNames.Remove($newVmId)
            Write-DebugLog "Cleared the recently-deleted marker for VM [$newVmId]: cluster/nextid reused it for this clone." -Level 'I' -Component '04' -Ref $newVmId
        }

        # Linked clones -- see LINKED-CLONES-DESIGN.md #3.3. Read both new params
        # defensively: the official Test-GuestsClone.ps1 omits 'snapshot' and
        # 'is_link_clone' entirely for a plain full clone rather than sending empty
        # values, and under Set-StrictMode a direct property read on an absent
        # PSCustomObject property throws -- an unguarded read here would crash
        # guests/clone with -32603 on every ordinary full clone.
        $snapshotName = ''
        if ((Get-MemberNames -Object $Params) -contains 'snapshot') { $snapshotName = [string]$Params.snapshot }
        $explicitLink = $null
        if ((Get-MemberNames -Object $Params) -contains 'is_link_clone') { $explicitLink = [bool]$Params.is_link_clone }

        # A non-empty 'snapshot' is the only linked-clone signal RAS actually sends at
        # template_method=basic -- 'is_link_clone' is documented as "only used by
        # template versions" and the test kit never sets it, so it is honoured as an
        # override when present, never required.
        $wantLinked = $script:Settings.capabilities.can_link_clones -and -not [string]::IsNullOrWhiteSpace($snapshotName)
        if ($wantLinked -and $null -ne $explicitLink) { $wantLinked = $wantLinked -and $explicitLink }

        $isLinked = $false
        if ($wantLinked) {
            # Proxmox's own rule, straight from the schema: cloning a normal (non-
            # template) VM is ALWAYS a full copy regardless of the 'full' flag -- no
            # error, no warning. Passing full=0 on a non-template source either fails
            # per-disk or silently full-clones depending on the disk, so the source's
            # template state must be confirmed here, not assumed.
            $sourceIsTemplate = $false
            try {
                $sourceLiveConfig = Get-ProxmoxVmConfig -Node $node -VmId $sourceVmId
                if ($null -ne $sourceLiveConfig -and (Get-MemberNames -Object $sourceLiveConfig) -contains 'template' -and $null -ne $sourceLiveConfig.template) {
                    $sourceIsTemplate = ([int]$sourceLiveConfig.template -eq 1)
                }
            }
            catch {
                Write-DebugLog "Live template-flag check failed for source VM [$sourceVmId], treating as not a template for linked-clone purposes: $($_.Exception.Message)" -Level 'W' -Component '04' -Ref $sourceVmId
            }

            if ($sourceIsTemplate) {
                $isLinked = $true
            }
            elseif ([string]$script:Settings.cloning.linked_clone_fallback -eq 'error') {
                return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Linked clone requested from source [$sourceVmId] snapshot [$snapshotName], but the source is not a Proxmox template"
            }
            else {
                # Default fallback: proceed as a full clone rather than silently letting
                # Proxmox downgrade it -- the log line is the only place this is
                # observable, and it is the whole reason this branch exists.
                Write-DebugLog "Linked clone requested from source [$sourceVmId] snapshot [$snapshotName], but the source is not a Proxmox template -- falling back to a full clone (cloning.linked_clone_fallback=full)." -Level 'W' -Component '04' -Ref $sourceVmId
            }
        }

        # Distributed placement -- see DISTRIBUTED-PLACEMENT.md and the 'placement' block
        # in Get-DefaultProviderSettings. $node above stays the URL's node (Proxmox routes
        # the clone POST through the SOURCE's own node regardless of where the result
        # lands); $targetNode is only ever the clone body's 'target' -- where the new VM's
        # compute actually runs. $null (placement disabled, or no eligible node found)
        # means "don't set target at all", i.e. exactly today's always-same-node behavior.
        $targetNode = Resolve-ProxmoxCloneTargetNode -SourceNode $node -CloneName $cloneName
        $cloneLandingNode = if (-not [string]::IsNullOrWhiteSpace($targetNode)) { $targetNode } else { $node }

        # pool_scope.inherit_on_clone -- see POOL-SCOPING.md. Read from the source's own
        # cluster/resources entry (already fetched above, no extra call), independently
        # of whether pool_name filtering is even on: a source that belongs to a pool
        # clones into that same pool by default so pool-scoped visibility and any
        # pool-based Proxmox permissions extend to the clone automatically. An unpooled
        # source clones unpooled, exactly like before this feature existed.
        $sourcePool = Get-ProxmoxVmPool -ClusterVm $clusterVm
        $inheritPool = $script:PoolScopeInheritOnClone -and -not [string]::IsNullOrWhiteSpace($sourcePool)

        # cloning.mac_preservation.enabled -- see MAC-PRESERVATION.md. Looked up ONCE
        # here, at clone-initiation time (same timing as placement's own preserved-node
        # lookup just above -- close to the delete this recreation followed, well inside
        # recently_deleted_retention_seconds), then stashed into this clone's own
        # tracking context below. Proxmox's clone API has no per-NIC override the way it
        # has 'target'/'pool', so this can't be applied to the request body itself --
        # Start-ProxmoxVmIfNeeded applies it later, once the real clone job is confirmed
        # done and before the VM is ever started. Stashing now (rather than re-querying
        # $script:RecentlyDeletedMacByName by name at that later point) means a slow
        # clone can't silently lose its preserved MAC to the SAME retention window
        # expiring mid-clone.
        $preservedMacs = if ($script:PreserveMacOnRecreation) { Get-ProxmoxPreservedMacsForRecreation -Name $cloneName } else { $null }

        $cloneSnapshotNote = if (-not [string]::IsNullOrWhiteSpace($snapshotName)) { " [snapshot=$snapshotName]" } else { '' }
        $placementNote = if (-not [string]::IsNullOrWhiteSpace($targetNode)) { " [target=$targetNode]" } else { '' }
        $poolNote = if ($inheritPool) { " [pool=$sourcePool]" } else { '' }
        $macNote = if ($null -ne $preservedMacs) { " [mac-preserved]" } else { '' }
        Write-DebugLog "Cloning [$sourceVmId] -> [$newVmId] as [$cloneName]: $(if ($isLinked) { 'linked (full=0)' } else { 'full (full=1)' })$cloneSnapshotNote$placementNote$poolNote$macNote." -Level 'I' -Component '04' -Ref $newVmId

        # Never pass storage/format (full-clone-only, rejected by Proxmox on a linked
        # clone) or snapname (the RAS name is not a legal pve-configid, and the base
        # image is already the right source -- see LINKED-CLONES.md "The name is not a
        # legal Proxmox snapshot name"). 'target' is compatible with both clone types on
        # this cluster specifically because storage is shared (Ceph/RBD) -- it only moves
        # compute, never disk placement; see DISTRIBUTED-PLACEMENT.md before enabling this
        # on a cluster with node-local storage.
        $body = @{
            newid = $newVmId
            name  = $cloneName
            full  = if ($isLinked) { 0 } else { 1 }
        }
        if (-not [string]::IsNullOrWhiteSpace($targetNode)) {
            $body.target = $targetNode
        }
        if ($inheritPool) {
            # Proxmox assigns pool membership atomically as part of the clone job
            # itself -- no separate post-clone API call, and no window where the new
            # VM briefly exists unpooled (which pool_name filtering would otherwise
            # make invisible to RAS the instant guests/get resolves it).
            $body.pool = $sourcePool
        }

        $resp = Invoke-ProxmoxApi -Method POST -Path "/api2/json/nodes/$node/qemu/$sourceVmId/clone" -Body $body
        $taskId = [string]$resp.data
        $cloneStartedAt = (Get-Date).ToString('o')   # elapsed-time source for pipelined completion, see top of file

        # Both copies of this context must carry task_id.
        # Get-RasGuestObjectForCloneAwareFlow's start guard reads $ctx.task_id to check
        # the REAL clone task before attempting a start; when the context resolves from
        # memory (the common case right after this call) rather than from the file, a
        # missing task_id makes that check silently no-op and the guard do nothing.
        if (-not [string]::IsNullOrWhiteSpace($taskId)) {
            $script:TaskContext[$taskId] = @{
                type               = 'clone'
                task_id            = $taskId
                source_id          = $sourceVmId
                clone_id           = $newVmId
                name               = $cloneName
                full               = -not $isLinked
                clone_node         = $cloneLandingNode
                start_issued       = $false
                start_task_id      = $null
                start_pending      = $false
                start_retry_count  = 0
                creation_completed = $false
                clone_started_at   = $cloneStartedAt
                preserved_macs     = $preservedMacs
            }
        }

        Set-CloneStateEntry -VmId $newVmId -Entry @{
            type               = 'clone'
            task_id            = $taskId
            source_id          = $sourceVmId
            clone_id           = $newVmId
            name               = $cloneName
            full               = -not $isLinked
            clone_node         = $cloneLandingNode
            start_issued       = $false
            start_task_id      = $null
            start_pending      = $false
            start_retry_count  = 0
            creation_completed = $false
            clone_started_at   = $cloneStartedAt
            preserved_macs     = $preservedMacs
        }

        # Otherwise only the TTL invalidates this cache, so a guests/get for the new id
        # landing before the TTL expires serves a listing from before the clone existed
        # -- a hard "not found in cluster" JSON-RPC error, which RAS's clone thread does
        # not retry.
        Reset-ProxmoxClusterCache

        return @{
            result = @{
                task_id  = $taskId
                clone_id = $newVmId
            }
        }
    }
    catch {
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to clone guest [$($Params.id)]: $($_.Exception.Message)"
    }
}

$script:MethodRegistry = @{
    'provider/initialize' = @{ Handler = { param($data) Handle-Initialize }; RequiredFields = @() }
    'provider/connect'    = @{ Handler = { param($data) Handle-Connect -Params $data.params }; RequiredFields = @('params.settings') }
    'provider/disconnect' = @{ Handler = { param($data) Handle-Disconnect }; RequiredFields = @() }

    'hosts/list'          = @{ Handler = { param($data) Handle-HostList }; RequiredFields = @() }
    'hosts/get'           = @{ Handler = { param($data) Handle-HostGet -Params $data.params }; RequiredFields = @('params.id') }
    'hosts/control'       = @{ Handler = { param($data) Handle-HostControl -Params $data.params }; RequiredFields = @('params.id', 'params.control') }

    'guests/list'         = @{ Handler = { param($data) Handle-GuestList }; RequiredFields = @() }
    'guests/get'          = @{ Handler = { param($data) Handle-GuestGet -Params $data.params }; RequiredFields = @('params.id') }
    'guests/control'      = @{ Handler = { param($data) Handle-GuestControl -Params $data.params }; RequiredFields = @('params.id', 'params.control') }

    'guests/convert'      = @{ Handler = { param($data) Handle-GuestConvert -Params $data.params }; RequiredFields = @('params.id', 'params.is_template') }
    'guests/clone'        = @{ Handler = { param($data) Handle-GuestClone -Params $data.params }; RequiredFields = @('params.id', 'params.name') }

    'guests/snapshots/create' = @{ Handler = { param($data) Handle-GuestSnapshotsCreate -Params $data.params }; RequiredFields = @('params.id', 'params.name') }
    'guests/snapshots/delete' = @{ Handler = { param($data) Handle-GuestSnapshotsDelete -Params $data.params }; RequiredFields = @('params.id', 'params.name') }
    'guests/snapshots/exists' = @{ Handler = { param($data) Handle-GuestSnapshotsExists -Params $data.params }; RequiredFields = @('params.id', 'params.name') }
    'guests/snapshots/revert' = @{ Handler = { param($data) Handle-GuestSnapshotsRevert -Params $data.params }; RequiredFields = @('params.id', 'params.name') }

    'tasks/get'           = @{ Handler = { param($data) Handle-TaskInfo -Params $data.params }; RequiredFields = @('params.id') }
}

function Process-Method {
    param([string]$InputLine)

    # This script has no thread of its own besides this read loop, so "poll every 30
    # seconds" (settings.locations.settings_reload_interval_seconds) means "check, at
    # most that often, on whichever request happens to land next". See
    # Update-ProviderSettingsIfChanged.
    Update-ProviderSettingsIfChanged

    Write-DebugLog "IN (PID=$PID): $InputLine" -Level 'D' -Component '01'

    $methodData = ConvertFrom-JsonSafe -InputLine $InputLine
    if ($null -eq $methodData) {
        return New-ErrorResponse -Code $script:ErrorCodes.ParseError -Message "$($script:ProviderNamePrefix) Invalid JSON format"
    }

    $methodName = $null
    if ((Get-MemberNames -Object $methodData) -contains 'method') {
        $methodName = [string]$methodData.method
    }

    if ([string]::IsNullOrWhiteSpace($methodName)) {
        return New-ErrorResponse -Code $script:ErrorCodes.MethodNotFound -Message "$($script:ProviderNamePrefix) Missing method name"
    }

    $lookupName = $methodName.Trim().ToLowerInvariant()
    if (-not $script:MethodRegistry.ContainsKey($lookupName)) {
        return New-ErrorResponse -Code $script:ErrorCodes.MethodNotFound -Message "$($script:ProviderNamePrefix) Unknown method: $methodName"
    }

    $methodEntry = $script:MethodRegistry[$lookupName]
    $validationError = Test-RequiredFields -Data $methodData -RequiredFields $methodEntry.RequiredFields
    if ($null -ne $validationError) {
        return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message $validationError
    }

    try {
        return & $methodEntry.Handler $methodData
    }
    catch {
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Method execution failed: $($_.Exception.Message)"
    }
}

# Restore the in-flight-clone id mirror from the persisted state file so
# Invoke-TrackedCloneSweep still knows what's in flight across a provider
# process restart, the same way the clone-state file itself already does.
try {
    $existingCloneState = Get-CloneStateAll
    foreach ($key in $existingCloneState.Keys) {
        $entry = $existingCloneState[$key]
        if ($null -ne $entry -and (Get-MemberNames -Object $entry) -contains 'type' -and [string]$entry.type -eq 'clone') {
            [void]$script:TrackedCloneVmIds.Add([string]$key)
        }
    }
}
catch {
    Write-DebugLog "Failed to restore tracked-clone id set from persisted state: $($_.Exception.Message)" -Level 'E' -Component '00'
}

# Startup banner, in the spirit of RAS's own module-init log lines (module version,
# starting module, host facts, operating mode) so a fresh log immediately answers "what
# is this deployment actually configured to do" without grepping further. Every fact
# here is already in memory or a free in-process lookup -- no WMI/CIM, no network call,
# no adapter enumeration -- so this costs nothing measurable at startup.
Write-DebugLog "Module version $script:ProviderVersion - Parallels-RAS-CPF-Proxmox-Advanced.ps1" -Level 'I' -Component '00'
Write-DebugLog "Starting RAS CPF Provider - PROXMOX" -Level 'I' -Component '00'
Write-DebugLog "Host - $([System.Environment]::MachineName), PowerShell $($PSVersionTable.PSVersion), $($PSVersionTable.OS)" -Level 'I' -Component '00'
Write-DebugLog ("Mode - linked_clones={0}, template_method={1}, guests_polling_rate={2}s, tasks_polling_rate={3}s, log_level={4} ({5})" -f `
    $(if ($script:Settings.capabilities.can_link_clones) { 'ENABLED' } else { 'DISABLED' }),
    $script:Settings.capabilities.template_method,
    $script:Settings.capabilities.guests_polling_rate,
    $script:Settings.capabilities.tasks_polling_rate,
    $script:LogLevel,
    $(switch ($script:LogLevel) { 3 { 'Standard' } 4 { 'Extended' } default { 'Verbose' } })
) -Level 'I' -Component '0A'

Write-DebugLog "Provider process started. PID=$PID" -Level 'I' -Component '00'

while ($true) {
    try {
        $inputLine = [Console]::In.ReadLine()

        if ($null -eq $inputLine) {
            Write-DebugLog 'Input stream closed. Exiting.' -Level 'I' -Component '00'
            break
        }

        # Not logged here -- Process-Method logs "IN (PID=...)" itself at its own top.
        $response = Process-Method -InputLine ($inputLine.Trim())
        Send-Response -ResponseObject $response
    }
    catch {
        $response = New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to process input: $($_.Exception.Message)"
        Send-Response -ResponseObject $response
    }
    finally {
        # Every swallowed exception on the hot path (failed REST calls,
        # placeholder-guest fallbacks) appends to $Error, which retains up to 256
        # ErrorRecords and pins whatever each one's TargetObject/InvocationInfo
        # references. Clearing once per request is essentially free.
        $Error.Clear()
    }
}
