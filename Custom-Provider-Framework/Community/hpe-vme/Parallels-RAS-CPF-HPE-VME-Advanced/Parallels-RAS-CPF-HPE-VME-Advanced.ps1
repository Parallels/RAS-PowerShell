<#
.SYNOPSIS
    Parallels RAS Custom Provider for HPE Morpheus VM Essentials (Advanced)
.DESCRIPTION
    A Parallels RAS Custom Provider Framework (CPF) integration for HPE Morpheus VM
    Essentials (VME), speaking JSON-RPC over stdin/stdout (one request per line in, one
    response per line out) and talking to VME through its REST API.

    Built for large fleets (thousands of instances from one provider process):
      - one in-memory instance inventory, refreshed by a paged listing at most every
        inventory.list_min_refresh_seconds and never wiped globally; guests/get is
        served from memory (no API call in steady state);
      - MAC and OS come from the server objects, fetched in bulk once per process;
      - full clones (template_method 'basic') and, on VME 9.1+, linked clones through
        template versioning (template_method 'versioning'): each RAS template version
        is one VME snapshot plus one linked-clone virtual image;
      - clone slots (cloning.max_concurrent_clone_operations) and early completion of
        linked clones, so clones provision side by side;
      - MAC address preservation across a same-name recreate;
      - delete with force=on and a per-id probe until VME returns 404; hard power-off
        for stop (power.force_stop);
      - HTTP over one reused connection with timeouts and idempotency-aware retries;
      - a self-seeding, schema-versioned settings file (VME-RAS-Config.json) and a
        structured, rotated log with secrets masked on every line.

    See README.md and docs/ for the full reference.

    Requires PowerShell 7.4 or later (connection reuse with a separate read timeout).
    On an older host -- including Windows PowerShell 5.1 by mistake -- every request,
    provider/initialize first, is answered with an error naming the version found. That
    is why there is no #Requires line: it would end the process before RAS got any
    answer. The file therefore keeps parsing on 5.1.
.NOTES
    File Name : Parallels-RAS-CPF-HPE-VME-Advanced.ps1
    Version   : 4.2.0 (CPF protocol version advertised to RAS: 1.0.0)
    Settings  : VME-RAS-Config.json (self-seeding, next to this script)
    Author    : www.parallels.com
.EXAMPLE
    Sample JSON-RPC requests, one per line on stdin:
    {"method":"provider/connect","params":{"settings":{"host":"vme.example.com","username":"admin","password":"XXX"}}}
    {"method":"guests/list"}
    {"method":"guests/get","params":{"id":"19"}}
    {"method":"guests/control","params":{"id":"19","control":"start"}}
    {"method":"guests/clone","params":{"id":"19","name":"RDSH-001"}}
    {"method":"tasks/get","params":{"id":"clonebyname:19:RDSH-001"}}
#>

Set-StrictMode -Version Latest

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$WarningPreference = 'SilentlyContinue'
$VerbosePreference = 'SilentlyContinue'
$InformationPreference = 'SilentlyContinue'

# stdout is the CPF protocol channel. ANSI styling or a BOM in front of the first
# response corrupts the frame RAS parses, so both are switched off explicitly.
if ($null -ne (Get-Variable -Name PSStyle -ErrorAction SilentlyContinue)) {
    try { $PSStyle.OutputRendering = 'PlainText' } catch { }
}
$script:Utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$script:ProviderNamePrefix = 'VME:'

# Minimum host version: below it the provider only ever
# answers with this error, so RAS shows the reason instead of a provider that died.
# Kept to code that runs on Windows PowerShell 5.1, and placed before the console
# encoding is touched (reassigning [Console]::InputEncoding can corrupt stdin on 5.1).
$script:MinPowerShellVersion = [version]'7.4'
if ($PSVersionTable.PSVersion -lt $script:MinPowerShellVersion) {
    $writer = New-Object System.IO.StreamWriter([Console]::OpenStandardOutput(), $script:Utf8NoBom)
    $writer.AutoFlush = $true
    $writer.NewLine = "`n"
    $why = "$($script:ProviderNamePrefix) PowerShell $($script:MinPowerShellVersion) or later is required; this provider was started by PowerShell $($PSVersionTable.PSVersion) ($($PSVersionTable.PSEdition)) at [$((Get-Process -Id $PID).Path)]. Install PowerShell 7.4+ and point the RAS custom provider command at its pwsh.exe."
    try {
        $fs = New-Object System.IO.FileStream((Join-Path $PSScriptRoot 'VME-RAS-Provider.log'), [System.IO.FileMode]::Append, [System.IO.FileAccess]::Write, [System.IO.FileShare]::ReadWrite)
        $lw = New-Object System.IO.StreamWriter($fs, $script:Utf8NoBom)
        $lw.WriteLine("[E 00/-/P$('{0:X4}' -f $PID)] $([DateTime]::Now.ToString('dd-MM-yy HH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture)) - $why Every request is refused.")
        $lw.Dispose()
    }
    catch { }
    $refusal = ConvertTo-Json -InputObject @{ error = @{ code = -32603; message = $why } } -Compress
    while ($true) {
        $line = [Console]::In.ReadLine()
        if ($null -eq $line) { break }
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        try { $writer.WriteLine($refusal) } catch { break }
    }
    exit 1
}

if ($Host.Name -notmatch 'ISE') {
    [Console]::InputEncoding = $script:Utf8NoBom
    [Console]::OutputEncoding = $script:Utf8NoBom
}

$stdout = [Console]::OpenStandardOutput()
$writer = New-Object System.IO.StreamWriter($stdout, $script:Utf8NoBom)
$writer.AutoFlush = $true
# Force LF: StreamWriter.WriteLine otherwise uses the platform newline, so framing
# would differ by deployment OS.
$writer.NewLine = "`n"

# CPF protocol version advertised by provider/initialize. RAS's VDI provider rejects
# anything it does not recognise ('2.0.0' is refused), so this tracks what RAS
# understands. The release version of this script is ScriptVersion (see CHANGELOG.md).
$script:ProviderVersion = '1.0.0'
$script:ScriptVersion = '4.2.0'
$script:ScriptFileName = if ($PSCommandPath) { Split-Path -Path $PSCommandPath -Leaf } else { 'Parallels-RAS-CPF-HPE-VME-Advanced.ps1' }

# Feature detection, once. Connection reuse with a per-request timeout needs 7.4+
# (-ConnectionTimeoutSeconds / -OperationTimeoutSeconds); -DateKind needs 7.5+.
$script:IwrParameters = (Get-Command Invoke-WebRequest).Parameters
$script:HasOperationTimeout = $script:IwrParameters.ContainsKey('OperationTimeoutSeconds')
$script:HasConnectionTimeout = $script:IwrParameters.ContainsKey('ConnectionTimeoutSeconds')
$script:HasDateKind = (Get-Command ConvertFrom-Json).Parameters.ContainsKey('DateKind')

# Every JSON document this script reads (requests, API responses, state files,
# settings) is parsed into hashtables. A hashtable returns $null for a missing key
# even under Set-StrictMode, so no "does this property exist" dance is needed, and
# there is no PSObject reflection on the hot path. -DateKind String keeps timestamps
# as the strings VME sent (ConvertFrom-Json otherwise turns them into [DateTime] and a
# later [string] cast renders them in the host's culture, which breaks on dd/MM hosts).
function ConvertFrom-JsonFast {
    param([string]$Json)
    if ($script:HasDateKind) { return (ConvertFrom-Json -InputObject $Json -AsHashtable -DateKind String -ErrorAction Stop) }
    return (ConvertFrom-Json -InputObject $Json -AsHashtable -ErrorAction Stop)
}

# Read one member of a parsed JSON node. Works on hashtables (the normal case) and on
# a PSCustomObject (defensive: a document with keys that differ only by case cannot
# be parsed -AsHashtable).
function Get-JsonValue {
    param([object]$Node, [string]$Name)
    if ($null -eq $Node) { return $null }
    if ($Node -is [System.Collections.IDictionary]) { return $Node[$Name] }
    if ($Node -is [System.Management.Automation.PSCustomObject]) {
        $p = $Node.PSObject.Properties[$Name]
        if ($null -ne $p) { return $p.Value }
    }
    return $null
}

function Get-EpochNow { return [DateTimeOffset]::UtcNow.ToUnixTimeSeconds() }

# Accepts epoch seconds (state files from 3.0) or an ISO-8601 string / [DateTime]
# (state files from 2.x). Returns epoch seconds or $null. Never goes through the host culture.
function ConvertTo-EpochSeconds {
    param([object]$Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [long] -or $Value -is [int] -or $Value -is [double] -or $Value -is [decimal]) { return [long]$Value }
    if ($Value -is [DateTime]) { return ([DateTimeOffset]$Value.ToUniversalTime()).ToUnixTimeSeconds() }
    $s = [string]$Value
    if ([string]::IsNullOrWhiteSpace($s)) { return $null }
    [long]$n = 0
    if ([long]::TryParse($s, [System.Globalization.NumberStyles]::Integer, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$n)) { return $n }
    try {
        $dto = [DateTimeOffset]::Parse($s, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AssumeUniversal)
        return $dto.ToUnixTimeSeconds()
    }
    catch { return $null }
}

function ConvertTo-IsoFromEpoch {
    param([long]$Epoch)
    return [DateTimeOffset]::FromUnixTimeSeconds($Epoch).UtcDateTime.ToString('yyyy-MM-ddTHH:mm:ssZ', [System.Globalization.CultureInfo]::InvariantCulture)
}

# ---------------------------------------------------------------------------
# Settings schema
#
# VME-RAS-Config.json: every leaf is {value, default, description}. Only 'value' is read; 'default' and
# 'description' are informational and are rewritten from this table whenever the
# file is migrated. A hand-simplified file with bare values is accepted too.
#
# The tree below is the ONLY place a setting is declared. The loader walks it
# generically: coercion follows the default's type, and _min/_max/_allowed (never
# written to the file) are enforced with a fallback to the default and a W line.
# ---------------------------------------------------------------------------

# 1 = the flat file of provider 2.x (bare values, no 'default'/'description'). 2 = this layout.
$script:CurrentSettingsSchemaVersion = 2

function New-SettingLeaf {
    param($Value, [string]$Description, $Min = $null, $Max = $null, [string[]]$Allowed = $null)

    $text = $Description
    if ($null -ne $Min -and $null -ne $Max) { $text = "$text Range: $Min-$Max." }
    if ($null -ne $Allowed) { $text = "$text Allowed: $($Allowed -join ', ')." }
    $leaf = [ordered]@{ value = $Value; default = $Value; description = $text }
    if ($null -ne $Min) { $leaf['_min'] = $Min }
    if ($null -ne $Max) { $leaf['_max'] = $Max }
    if ($null -ne $Allowed) { $leaf['_allowed'] = $Allowed }
    return $leaf
}

function Get-DefaultProviderSettings {
    return [ordered]@{
        _comment         = "HPE VME CPF provider settings. Every leaf is {value, default, description}: only 'value' is read, 'default' and 'description' are informational and are refreshed by the provider when it migrates this file. See docs/SETTINGS.md."
        schema_version   = $script:CurrentSettingsSchemaVersion

        locations        = [ordered]@{
            _comment                         = 'Where the provider keeps its files. Changes are applied at the next process start, never on a hot reload (in-flight clone tracking lives in these files).'
            data_directory                   = New-SettingLeaf '' 'Folder for the log and state files. Empty = the script folder; a relative path resolves against the script folder.'
            log_file                         = New-SettingLeaf 'VME-RAS-Provider.log' 'Log file name, relative to data_directory unless absolute.'
            clone_state_file                 = New-SettingLeaf 'VME-RAS-CloneState.json' 'In-flight clone tracking. Survives a provider restart.'
            maintenance_state_file           = New-SettingLeaf 'VME-RAS-Maintenance.json' 'Templates RAS has put into maintenance (guests/convert is_template=false). Survives a restart.'
            template_state_file              = New-SettingLeaf 'VME-RAS-BasicTemplates.json' "Instances RAS marked as templates (template_method 'basic'). The only source of is_template."
            former_template_state_file       = New-SettingLeaf 'VME-RAS-FormerTemplates.json' "Instances that were RAS templates when RAS sent a delete for them ('Delete Template' in the RAS Console). The provider never deletes these VMs; they stay listed as ordinary guests and can be made templates again."
            stuck_clone_state_file           = New-SettingLeaf 'VME-RAS-StuckClones.json' 'Clones the provider gave up on that VME has not finished yet. Each one keeps holding a cloning.max_concurrent_clone_operations slot. Survives a restart.'
            template_version_state_file      = New-SettingLeaf 'VME-RAS-TemplateVersions.json' "template_method 'versioning': each RAS template version's VME snapshot and linked-clone virtual image. Rebuilt from VME (by name) for a version it is missing."
            settings_reload_interval_seconds = New-SettingLeaf 30 "How often this file's last-write time is checked, opportunistically on incoming requests." 5 3600
        }

        connection       = [ordered]@{
            _comment                     = 'VME REST transport. Host and credentials come from the RAS Console (provider/connect), not from this file.'
            ssl_check                    = New-SettingLeaf $true 'Validate the appliance TLS certificate. A per-connection ssl_check in the RAS custom settings overrides this.'
            use_system_proxy             = New-SettingLeaf $false 'false = connect directly (-NoProxy). true = use the system proxy, including WPAD auto-detection, which can stall a call for seconds.'
            request_timeout_seconds      = New-SettingLeaf 30 'How long a call may wait for VME to start answering (connect + TLS + server processing). One value for every call: on PowerShell 7.4+ it belongs to the pooled connection, and changing it per call would drop the connection. A clone PUT that exceeds it is tracked as "response lost", never retried.' 5 300
            read_timeout_seconds         = New-SettingLeaf 60 'PowerShell 7.4+: longest silence tolerated while a response body is being received (a multi-MB listing page). Older hosts have no separate read timeout.' 5 600
            retry_max_attempts           = New-SettingLeaf 3 'Attempts for a GET on no response, HTTP 429 or 5xx. Idempotent writes (power actions, DELETE) get at most 2. A clone PUT is never retried.' 1 5
            retry_delay_milliseconds     = New-SettingLeaf 250 'Base delay between retries, multiplied by the attempt number.' 0 5000
            retry_after_max_seconds      = New-SettingLeaf 5 "Ceiling on an HTTP 429 Retry-After wait. The provider serves one request at a time, so every second waited delays every other RAS call." 0 30
            oauth_client_id              = New-SettingLeaf 'morph-api' 'OAuth client id for the password and refresh_token grants.'
            token_refresh_margin_seconds = New-SettingLeaf 300 'Refresh a token whose expiry is known this many seconds before it expires.' 30 86400
            token_renewal                = [ordered]@{
                _comment                    = "Proactive renewal of a pre-issued access_token (RAS custom setting) whose expiry is discovered via GET /api/user-settings. Needs a refresh_token."
                proactive_enabled           = New-SettingLeaf $true 'Discover the pre-issued token expiry and renew it ahead of time. The reactive renewal on HTTP 401 always stays on.'
                expiry_check_interval_hours = New-SettingLeaf 24 'How often the expiry is re-discovered.' 1 720
                lead_days                   = New-SettingLeaf 1 'Start renewing this many days before the discovered expiry.' 0 30
                max_attempts                = New-SettingLeaf 5 'Renewal attempts per expiry cycle.' 1 50
                retry_interval_hours        = New-SettingLeaf 4 'Minimum spacing between renewal attempts.' 1 72
            }
        }

        capabilities     = [ordered]@{
            _comment              = "Advertised to RAS by provider/initialize. can_link_clones is not a setting: it is true exactly when template_method is 'versioning' (VME 9.1+)."
            can_suspend_guests    = New-SettingLeaf $true 'Advertise guests/control(suspend).'
            guests_polling_rate   = New-SettingLeaf 60 'Seconds between RAS guests/list polls. At 3000 guests one listing is ~18 MB, so this is also the natural refresh cadence of the inventory.' 5 600
            tasks_polling_rate    = New-SettingLeaf 10 'Seconds between RAS tasks/get polls for an in-flight clone. RAS learns that a clone finished only at its next poll, so a lower value finishes pool builds sooner. Keep tasks_polling_rate x tasks_polling_retries above cloning.provision_timeout_seconds.' 5 300
            tasks_polling_retries = New-SettingLeaf 180 'tasks/get polls before RAS gives up on a task. CPF documents 0-180.' 1 180
            template_method       = New-SettingLeaf 'basic' "CPF template contract. 'basic': guests/convert marks templates, full clones only. 'versioning': template versions with linked clones (VME 9.1+) -- each version is a VME snapshot plus a linked-clone virtual image, and can_link_clones is advertised. 'none': no templates (guests/convert and guests/clone are refused). RAS reads this at provider/initialize; existing templates may need re-creating after a change." -Allowed @('none', 'basic', 'versioning')
        }

        inventory        = [ordered]@{
            _comment                       = 'The in-memory instance inventory that serves guests/list and guests/get. See docs/DESIGN.md.'
            list_min_refresh_seconds       = New-SettingLeaf 60 'Minimum age of the inventory before a guests/list triggers a new full listing. A guests/list that arrives sooner is answered from memory. guests/get never triggers a listing.' 0 3600
            page_size                      = New-SettingLeaf 200 'Instances per page of a routine listing (VME default 25). A routine pass walks every page inside one guests/list; a W line reports a pass slower than capabilities.guests_polling_rate.' 25 1000
            cold_start_page_size           = New-SettingLeaf -1 "The first listing after the provider starts (RAS is still syncing; nothing routine or cloning runs yet). -1 = the whole fleet in one request (max=-1). 0 = paged like a routine listing. A positive number = that many per request. If VME returns fewer rows than meta.total, or refuses the request, the same listing continues with page_size pages." -1 10000
            max_pages                      = New-SettingLeaf 50 'Safety cap on listing pages. Reaching it logs W and switches that refresh to add/update only.' 1 1000
            max_drop_ratio                 = New-SettingLeaf 0.5 'Mass-drop guard: a listing that is smaller than the current inventory by more than this fraction is not applied until a second listing confirms it.' 0.05 1.0
            missing_listings_before_drop   = New-SettingLeaf 2 'An instance absent from this many consecutive complete listings is dropped. Protects against offset paging skipping an id while instances are created or deleted.' 1 10
            exclude_instance_name_patterns = New-SettingLeaf @() 'Regular expressions matched against the instance name. A match hides the instance from guests/list.'
            max_addresses                  = New-SettingLeaf 3 'Maximum IPv4 and MAC addresses reported per guest.' 1 16
            hide_ip_when_stopped           = New-SettingLeaf $true 'Report no IP for a powered-off guest, even if VME still carries the last address in connectionInfo.'
            power_action_hold_seconds      = New-SettingLeaf 120 "After a start/stop/suspend this provider issued, how long a listing that still shows the instance's previous power state is not believed: guests/get reports the transition (powering_off / powering_on / suspending) and re-reads that one instance from VME (watch_refresh_interval_seconds) until VME shows the new state. VME can list a stopped desktop as running for tens of seconds; RAS would then see it power back on. 0 = off." 0 1800
            watch_refresh_interval_seconds = New-SettingLeaf 10 'How often guests/get may re-read (one lean GET) an instance under watch: one with a provider-issued power action pending (power_action_hold_seconds), or a running template in maintenance -- RAS shuts that down through its guest agent on Exit Maintenance, which the listing alone shows up to a polling cycle late.' 2 300
            scope                          = [ordered]@{
                _comment      = 'Server-side filters appended to every listing, so instances outside the VDI estate are never downloaded. 0 or empty = off.'
                site_id       = New-SettingLeaf 0 'VME group id (siteId).' 0 2147483647
                zone_id       = New-SettingLeaf 0 'VME cloud id (zoneId).' 0 2147483647
                instance_type = New-SettingLeaf '' "Instance type code, e.g. 'mvm'."
                label         = New-SettingLeaf '' 'Only instances carrying this label.'
            }
            hardware                       = [ordered]@{
                _comment                     = 'MAC and OS, which exist only on the server object. MACs never change for the life of an instance, so they are learned once per process.'
                bulk_enabled                 = New-SettingLeaf $true 'Learn MAC/OS for the whole fleet with one paged GET /api/servers?vm=true&stats=false after the first listing. false = one GET /api/servers/{id} per instance, on its first guests/get.'
                bulk_min_interval_seconds    = New-SettingLeaf 900 'Minimum time between bulk passes.' 60 86400
                bulk_rerun_missing_threshold = New-SettingLeaf 25 'Run another bulk pass (subject to bulk_min_interval_seconds) once this many listed instances still have no MAC.' 1 100000
                hw_retry_seconds             = New-SettingLeaf 30 'While an instance has no MAC yet (a clone still deploying), how often guests/get may fetch its server object.' 5 3600
            }
        }

        cloning          = [ordered]@{
            _comment                                      = 'guests/clone and tasks/get. VME returns no id for a clone, so the provider finds it by name.'
            resolve_method                                = New-SettingLeaf 'phrase' "How tasks/get finds a new clone before the next full listing. 'phrase': GET /api/instances?phrase=<name> (small). 'activity': GET /api/activity?name=<name> (experimental). 'list': full listings only. Every candidate is validated (exact name, newer than the submit, not the source, lineage), and a shared full listing every resolve_list_interval_seconds is the safety net in every mode." -Allowed @('phrase', 'activity', 'list')
            resolve_list_interval_seconds                 = New-SettingLeaf 60 'Minimum seconds between full listings used to resolve clones, shared by every clone in flight. The first one happens this long after the submit.' 10 3600
            resolve_timeout_seconds                       = New-SettingLeaf 900 'A clone not found this long after the submit fails.' 60 7200
            provision_timeout_seconds                     = New-SettingLeaf 1740 "A clone whose VME status is still provisioning this long after the submit fails. Default = 180 x 10 s minus one minute, so the provider fails with a reason before RAS's own tasks/get budget runs out." 120 14400
            cancel_process_on_giveup                      = New-SettingLeaf $true 'When provision_timeout_seconds is reached, also POST /api/processes/{id}/cancel on the clone''s newest provision process. Best-effort only -- VME can refuse this ("Process is not cancelable") for exactly the stuck-provision case this targets, so it is not a reliable way to stop the underlying work, only a logged attempt. Failures never change the task outcome.'
            babysit_window_seconds                        = New-SettingLeaf 7200 'After tasks/get reported the clone completed, guests/get keeps re-reading it (every tracked_refresh_interval_seconds) and starting it if found off until a real IPv4 appears, for at most this long. A running clone is reported powered_on with or without an IP.' 300 86400
            tracked_refresh_interval_seconds              = New-SettingLeaf 5 'A tracked clone is re-read from VME (one lean GET) at most this often.' 1 300
            tracked_sweep_max_per_call                    = New-SettingLeaf 3 'Other tracked clones re-checked opportunistically per guests/get call.' 0 50
            completed_clone_task_output_retention_seconds = New-SettingLeaf 600 "How long a completed clone task's clone_id is remembered for a repeated tasks/get." 60 86400
            max_concurrent_clone_operations               = New-SettingLeaf 10 'Most clones allowed unfinished on VME at once, for this provider (every template together): clones RAS is still waiting for, clones reported to RAS early that VME is still provisioning, plus clones the provider gave up on (resolve/provision timeout) that VME has not finished. A guests/clone beyond it is refused at once, without calling VME, so stuck clones cannot pile up on the appliance. A given-up clone frees its slot only when VME confirms it ended (finished, failed or deleted) -- there is no expiry. Caps early completion (linked_clones.early_completion, deploy_event_gate_enabled): with it, linked-clone throughput is about this many clones per VME provisioning time. Default 10, as RAS uses for its other providers.' 1 100
            stuck_clone_recheck_seconds                   = New-SettingLeaf 60 'While the clone limit is reached, how often the given-up clones are re-checked on VME: one lean GET per clone with a known id, one GET /api/processes?instanceId=<template> per template for clones that never appeared. Refusals in between are answered from memory.' 10 3600
            deploy_event_gate_enabled                     = New-SettingLeaf $false "EXPERIMENTAL. Report a clone completed once VME's 'Deploy Virtual Machine' step finished (one GET /api/processes per poll), before provisioning as a whole finished."
            template_flapping_workaround_enabled          = New-SettingLeaf $false 'Workaround for an appliance that reports a template''s power state inconsistently: always report a Ready template (is_template=true) as powered off with no IP/MAC, unless RAS explicitly started it.'
            mac_preservation                              = [ordered]@{
                _comment          = 'VME honours networkInterfaces[].macAddress on clone. See docs/MAC-PRESERVATION.md.'
                enabled           = New-SettingLeaf $true 'When RAS deletes a guest and clones a new one under the same name, request the old MAC for the new NIC. RAS itself does not forward the MAC on a recreate, so this is the only way it survives.'
                retention_seconds = New-SettingLeaf 300 'How long a captured MAC waits for a same-named clone.' 30 86400
            }
        }

        linked_clones    = [ordered]@{
            _comment                 = "template_method 'versioning' only (VME 9.1+). Each RAS template version (RAS_TEMPLATE_VERSION_<n>) is one VME snapshot named RAS_TEMPLATE_VERSION_<n>_<template id> plus the linked-clone virtual image VME makes from it (same name). Clones are POST /api/instances with config.imageId. See docs/LINKED-CLONES.md."
            snapshot_timeout_seconds = New-SettingLeaf 1500 "How long guests/snapshots/create may take (snapshot process + linked-clone image), and how long guests/snapshots/delete waits for VME to confirm the snapshot gone. Keep below tasks_polling_rate x tasks_polling_retries." 60 14400
            image_wait_seconds       = New-SettingLeaf 300 'After the linked-clone call, how long the virtual image may take to show up in GET /api/virtual-images before the version fails.' 10 3600
            image_filter_type        = New-SettingLeaf 'synced' "filterType for the image lookup by name. VME 9.1 files linked-clone images under 'Synced' (the API default 'User' does not show them)." -Allowed @('synced', 'all', 'user', 'system')
            label_images             = New-SettingLeaf $true 'Label each image ras-template-<id> and ras-version-<n> (one PUT /api/virtual-images/{id}, best-effort), so the version map can be rebuilt from VME.'
            refuse_delete_in_use     = New-SettingLeaf $true 'Refuse guests/snapshots/delete for a version whose image a listed instance still uses (config.imageId). Deleting the snapshot also deletes the image on VME.'
            early_completion         = New-SettingLeaf $true "Report a linked clone completed on its first tasks/get, as soon as VME has given it an instance id and not failed it, instead of after VME's whole provisioning (typically 80-90 s). RAS sends guests/clone one at a time and the next only after the previous task completed, so this is what lets clones provision side by side. Held while cloning.max_concurrent_clone_operations clones are still provisioning. A clone that fails after the report is logged E; RAS already counts it as created."
        }

        power            = [ordered]@{
            _comment   = 'guests/control power actions.'
            force_stop = New-SettingLeaf $true "Send stop as VME's hard power-off (PUT /api/instances/{id}/stop?force=true, VME 9.1+): seconds instead of the guest OS shutdown VME does otherwise (typically 40-60 s). RAS shuts a guest down cleanly through its own agent; a stop it sends the provider is a power-off (its log: 'shutting down guest ... non-gracefully'). Applies to every stop, templates included, and to delete.stop_before_delete. Off: VME's graceful shutdown."
        }

        delete           = [ordered]@{
            _comment               = 'guests/control(delete): one DELETE with force=on, then the id is probed until VME returns 404. guests/list keeps listing it (as powering_off) until then.'
            stop_before_delete     = New-SettingLeaf $false 'Issue PUT stop before DELETE for a running instance. Off: DELETE with force=on handles a running instance.'
            keep_backups           = New-SettingLeaf $false 'keepBackups on the DELETE. Off: deleted desktops leave no backup records.'
            release_eips           = New-SettingLeaf $true 'releaseEIPs on the DELETE.'
            probe_interval_seconds = New-SettingLeaf 5 'How often a deleting id is re-read (GET, 404 = gone) during guests/list.' 1 300
            max_probes_per_list    = New-SettingLeaf 25 'Deleting ids probed per guests/list call, oldest probe first.' 1 500
            stuck_after_seconds    = New-SettingLeaf 600 'A deleting instance still running or stopped this long after the DELETE is logged as W.' 60 86400
            give_up_after_seconds  = New-SettingLeaf 3600 'After this long the id stops being probed and falls back to the normal listing rules.' 300 604800
        }

        orphan_detection = [ordered]@{
            _comment                 = 'Log-only. Flags a clone of a known template that RAS has stopped polling.'
            enabled                  = New-SettingLeaf $true 'Run the audit.'
            check_interval_seconds   = New-SettingLeaf 300 'How often the audit (and the in-memory garbage collection that shares its timer) runs.' 30 86400
            stale_poll_after_seconds = New-SettingLeaf 600 'How long since the last guests/get for an instance before it counts as not polled. Keep well above one full RAS sweep.' 60 604800
        }

        logging          = [ordered]@{
            _comment                   = 'See docs/LOGGING.md. E and W lines are always written. Secrets are always masked.'
            log_level                  = New-SettingLeaf 4 "3 = Standard (lifecycle), 4 = Extended (+ per-operation detail, default), 5 = Verbose (+ every RPC and HTTP call). Accepts the number or 'standard'/'extended'/'verbose'."
            log_rpc_bodies             = New-SettingLeaf $false 'At log_level 5, also log the OUT response of every guests/get (otherwise the one-line GUEST summary stands in for it).'
            log_rotate_max_mb          = New-SettingLeaf 20 'Log size that triggers rotation.' 1 2048
            log_rotate_max_generations = New-SettingLeaf 3 'Rotated generations kept.' 1 20
            stats_interval_seconds     = New-SettingLeaf 300 'Interval of the I-level call-budget summary (RPCs and VME calls per endpoint, bytes, time). 0 = only at shutdown.' 0 86400
            file_exists_check_seconds = New-SettingLeaf 60 'How often to check the log file still exists at its path -- an external delete (cleanup job, antivirus, a user tidying up) does not fail the next write, so this is what notices and recreates it, logging a W line about the gap.' 5 3600
        }
    }
}

function Test-SettingLeaf {
    param([object]$Node)
    return ($Node -is [System.Collections.IDictionary] -and $Node.Contains('value') -and $Node.Contains('description'))
}

# The schema-1 key paths (provider 2.x) that moved in schema 2, and how their values
# translate. A schema-1 key that is not here and not at the same path was removed.
$script:LegacySettingsRenames = [ordered]@{
    'connection.http_timeout_seconds'              = 'connection.request_timeout_seconds'
    'cloning.instances_page_size'                  = 'inventory.page_size'
    'cloning.instances_max_pages'                  = 'inventory.max_pages'
    'filtering.exclude_instance_name_patterns'     = 'inventory.exclude_instance_name_patterns'
    'network.max_addresses'                        = 'inventory.max_addresses'
    'cloning.sweep_min_interval_seconds'           = 'cloning.tracked_refresh_interval_seconds'
    'cloning.clone_deploy_event_gate_enabled'      = 'cloning.deploy_event_gate_enabled'
    'cloning.mac_preservation_enabled'             = 'cloning.mac_preservation.enabled'
    'cloning.mac_preservation_retention_seconds'   = 'cloning.mac_preservation.retention_seconds'
    'connection.proactive_token_renewal_enabled'   = 'connection.token_renewal.proactive_enabled'
    'connection.token_expiry_check_interval_hours' = 'connection.token_renewal.expiry_check_interval_hours'
    'connection.token_renewal_lead_days'           = 'connection.token_renewal.lead_days'
    'connection.token_renewal_max_attempts'        = 'connection.token_renewal.max_attempts'
    'connection.token_renewal_retry_interval_hours' = 'connection.token_renewal.retry_interval_hours'
}

# Provider 2.x self-seeded every key at its default, so a schema-1 value equal to the
# old default says nothing about the admin's intent. Those take the new default instead.
$script:LegacyChangedDefaults = @{
    'capabilities.guests_polling_rate' = 5
    'capabilities.tasks_polling_rate'  = 30
    'capabilities.template_method'     = 'versioning'
    'logging.log_level'                = 5
    'connection.request_timeout_seconds' = 20
}

# Flattens a parsed settings document into "a.b.c" -> raw value. A leaf is anything
# that is not a dictionary, or a dictionary carrying 'value' (the rich shape).
# $BareLeaves collects the paths written as bare values, so the loader knows the file
# needs rewriting into the rich shape. $StoredDefaults collects each rich leaf's
# 'default' as the file recorded it (Update-FlatSettingsForChangedDefaults).
function Get-FlatParsedSettings {
    param([object]$Node, [string]$Prefix, [System.Collections.IDictionary]$Result, [System.Collections.Generic.List[string]]$BareLeaves, [System.Collections.IDictionary]$StoredDefaults = $null)

    foreach ($key in @($Node.Keys)) {
        $name = [string]$key
        if ($name.StartsWith('_') -or ($Prefix -eq '' -and $name -eq 'schema_version')) { continue }
        $child = $Node[$key]
        $path = if ($Prefix) { "$Prefix.$name" } else { $name }
        if ($child -is [System.Collections.IDictionary]) {
            if ($child.Contains('value')) {
                $Result[$path] = $child['value']
                if ($null -ne $StoredDefaults -and $child.Contains('default')) { $StoredDefaults[$path] = $child['default'] }
                continue
            }
            Get-FlatParsedSettings -Node $child -Prefix $path -Result $Result -BareLeaves $BareLeaves -StoredDefaults $StoredDefaults
            continue
        }
        $Result[$path] = $child
        $BareLeaves.Add($path)
    }
}

# A leaf whose value still equals the default the file recorded, when this version
# ships a different default, was never set by the admin: it follows the new default
# (e.g. cloning.max_concurrent_clone_operations 2 -> 10). Any other value is kept.
# Returns human-readable notes; the caller rewrites the file so 'default' is current.
function Update-FlatSettingsForChangedDefaults {
    param([System.Collections.IDictionary]$Flat, [System.Collections.IDictionary]$StoredDefaults, [System.Collections.IDictionary]$Defaults)

    $notes = [System.Collections.Generic.List[string]]::new()
    $current = ConvertTo-FlatResolvedSettings -Node (Resolve-SettingsNode -DefaultsNode $Defaults -Prefix '' -Flat @{} -Problems ([System.Collections.Generic.List[string]]::new()))
    foreach ($path in @($StoredDefaults.Keys)) {
        if (-not $current.Contains($path) -or -not $Flat.Contains($path)) { continue }
        $old = $StoredDefaults[$path]
        $oldText = if ($old -is [System.Array]) { $old -join ',' } else { "$old" }
        $valText = if ($Flat[$path] -is [System.Array]) { $Flat[$path] -join ',' } else { "$($Flat[$path])" }
        if ($oldText -eq "$($current[$path])" -or $valText -ne $oldText) { continue }
        $Flat.Remove($path)
        $notes.Add("$path was at its previous default [$oldText] -> new default [$($current[$path])] applied")
    }
    return , $notes
}

function Get-DefaultLeafPaths {
    param([object]$Node, [string]$Prefix = '')
    $paths = [System.Collections.Generic.List[string]]::new()
    foreach ($key in $Node.Keys) {
        if ([string]$key -like '_*' -or [string]$key -eq 'schema_version') { continue }
        $child = $Node[$key]
        $path = if ($Prefix) { "$Prefix.$key" } else { [string]$key }
        if (Test-SettingLeaf $child) { $paths.Add($path) }
        elseif ($child -is [System.Collections.IDictionary]) { $paths.AddRange([string[]](Get-DefaultLeafPaths -Node $child -Prefix $path)) }
    }
    return , $paths
}

# Translates a schema-1 flat map in place. Returns human-readable notes.
function Convert-LegacyFlatSettings {
    param([System.Collections.IDictionary]$Flat)

    $notes = [System.Collections.Generic.List[string]]::new()
    $oldTasksRate = 30
    if ($Flat.Contains('capabilities.tasks_polling_rate')) { try { $oldTasksRate = [int]$Flat['capabilities.tasks_polling_rate'] } catch { } }

    foreach ($old in @($script:LegacySettingsRenames.Keys)) {
        if ($Flat.Contains($old)) {
            $Flat[$script:LegacySettingsRenames[$old]] = $Flat[$old]
            $Flat.Remove($old)
            $notes.Add("$old -> $($script:LegacySettingsRenames[$old])")
        }
    }

    # Attempt budgets became time budgets (review 4.3). Default attempt counts are
    # dropped so the new time defaults apply; customised ones are converted at the
    # tasks_polling_rate they were sized for.
    foreach ($pair in @(@('cloning.clone_resolve_max_attempts', 'cloning.resolve_timeout_seconds', 80), @('cloning.clone_provision_wait_max_attempts', 'cloning.provision_timeout_seconds', 160))) {
        if ($Flat.Contains($pair[0])) {
            $attempts = 0
            try { $attempts = [int]$Flat[$pair[0]] } catch { }
            $Flat.Remove($pair[0])
            if ($attempts -gt 0 -and $attempts -ne $pair[2]) {
                $Flat[$pair[1]] = $attempts * [Math]::Max(1, $oldTasksRate)
                $notes.Add("$($pair[0])=$attempts -> $($pair[1])=$($Flat[$pair[1]]) (at the old tasks_polling_rate of ${oldTasksRate}s)")
            }
        }
    }

    if ($Flat.Contains('logging.log_rotate_max_bytes')) {
        try {
            $bytes = [long]$Flat['logging.log_rotate_max_bytes']
            if ($bytes -ne 20971520) { $Flat['logging.log_rotate_max_mb'] = [int][Math]::Max(1, [Math]::Round($bytes / 1MB)) }
        }
        catch { }
        $Flat.Remove('logging.log_rotate_max_bytes')
    }

    foreach ($path in @($script:LegacyChangedDefaults.Keys)) {
        if ($Flat.Contains($path) -and [string]$Flat[$path] -eq [string]$script:LegacyChangedDefaults[$path]) {
            $Flat.Remove($path)
            $notes.Add("$path was at its schema-1 default [$($script:LegacyChangedDefaults[$path])] -> current default applied")
        }
    }
    return , $notes
}

function ConvertTo-CoercedSettingBool {
    param($Value, [bool]$Default)
    if ($null -eq $Value) { return $Default }
    if ($Value -is [bool]) { return $Value }
    switch (([string]$Value).Trim().ToLowerInvariant()) {
        { $_ -in @('true', '1', 'yes', 'on') } { return $true }
        { $_ -in @('false', '0', 'no', 'off') } { return $false }
    }
    throw "not a boolean"
}

function ConvertTo-CoercedNumber {
    param($Value)
    if ($Value -is [int] -or $Value -is [long] -or $Value -is [double] -or $Value -is [decimal]) { return [double]$Value }
    [double]$n = 0
    if ([double]::TryParse(([string]$Value).Trim(), [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$n)) { return $n }
    throw "not a number"
}

function ConvertTo-CoercedLogLevel {
    param($Value)
    switch (([string]$Value).Trim().ToLowerInvariant()) {
        'standard' { return 3 }
        'extended' { return 4 }
        'verbose' { return 5 }
    }
    $n = [int](ConvertTo-CoercedNumber $Value)
    if ($n -lt 3 -or $n -gt 5) { throw 'must be 3, 4 or 5' }
    return $n
}

# One leaf: coerce the raw file value by the default's type, then enforce the rules.
# Returns the default (and records why) for anything that doesn't fit.
function Resolve-SettingLeafValue {
    param([string]$Path, [System.Collections.IDictionary]$Leaf, [object]$Raw, [System.Collections.Generic.List[string]]$Problems)

    $default = $Leaf['default']
    if ($null -eq $Raw) { return $default }
    try {
        $v = $null
        if ($Path -eq 'logging.log_level') { $v = ConvertTo-CoercedLogLevel $Raw }
        elseif ($default -is [bool]) { $v = ConvertTo-CoercedSettingBool -Value $Raw -Default $default }
        elseif ($default -is [int] -or $default -is [long]) {
            $d = ConvertTo-CoercedNumber $Raw
            $v = [int][Math]::Round($d)
        }
        elseif ($default -is [double]) { $v = ConvertTo-CoercedNumber $Raw }
        elseif ($default -is [System.Array]) {
            $v = @(@($Raw) | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        }
        else {
            $v = [string]$Raw
            if ($Leaf.Contains('_allowed')) {
                $lower = $v.Trim().ToLowerInvariant()
                if ($Leaf['_allowed'] -notcontains $lower) { throw "must be one of: $($Leaf['_allowed'] -join ', ')" }
                $v = $lower
            }
        }

        if ($Leaf.Contains('_min') -and $v -lt $Leaf['_min']) { throw "below the minimum $($Leaf['_min'])" }
        if ($Leaf.Contains('_max') -and $v -gt $Leaf['_max']) { throw "above the maximum $($Leaf['_max'])" }
        return $v
    }
    catch {
        $shown = if ($Raw -is [System.Array]) { $Raw -join ',' } else { [string]$Raw }
        $Problems.Add("$Path = [$shown] rejected ($($_.Exception.Message)) -- using the default [$default]")
        return $default
    }
}

# Builds the resolved (value-only) tree that mirrors the defaults tree.
function Resolve-SettingsNode {
    param([System.Collections.IDictionary]$DefaultsNode, [string]$Prefix, [System.Collections.IDictionary]$Flat, [System.Collections.Generic.List[string]]$Problems)

    $result = [ordered]@{}
    foreach ($key in $DefaultsNode.Keys) {
        if ([string]$key -like '_*') { continue }
        $child = $DefaultsNode[$key]
        $path = if ($Prefix) { "$Prefix.$key" } else { [string]$key }
        if (Test-SettingLeaf $child) {
            $raw = if ($Flat.Contains($path)) { $Flat[$path] } else { $null }
            $result[$key] = Resolve-SettingLeafValue -Path $path -Leaf $child -Raw $raw -Problems $Problems
        }
        elseif ($child -is [System.Collections.IDictionary]) {
            $result[$key] = Resolve-SettingsNode -DefaultsNode $child -Prefix $path -Flat $Flat -Problems $Problems
        }
        else {
            $result[$key] = $child
        }
    }
    return $result
}

# The file form: the defaults tree with every leaf's 'value' replaced by the resolved
# value and the validation-only keys (_min/_max/_allowed) stripped.
function New-SettingsFileTree {
    param([System.Collections.IDictionary]$DefaultsNode, [System.Collections.IDictionary]$ResolvedNode)

    $result = [ordered]@{}
    foreach ($key in $DefaultsNode.Keys) {
        $child = $DefaultsNode[$key]
        if (Test-SettingLeaf $child) {
            $result[$key] = [ordered]@{ value = $ResolvedNode[$key]; default = $child['default']; description = $child['description'] }
        }
        elseif ($child -is [System.Collections.IDictionary]) {
            $result[$key] = New-SettingsFileTree -DefaultsNode $child -ResolvedNode $ResolvedNode[$key]
        }
        else {
            $result[$key] = $child
        }
    }
    return $result
}

function ConvertTo-FlatResolvedSettings {
    param([System.Collections.IDictionary]$Node, [string]$Prefix = '')
    $result = [ordered]@{}
    foreach ($key in $Node.Keys) {
        $child = $Node[$key]
        $path = if ($Prefix) { "$Prefix.$key" } else { [string]$key }
        if ($child -is [System.Collections.IDictionary]) {
            $sub = ConvertTo-FlatResolvedSettings -Node $child -Prefix $path
            foreach ($k in $sub.Keys) { $result[$k] = $sub[$k] }
        }
        else {
            $result[$path] = if ($child -is [System.Array]) { $child -join ',' } else { $child }
        }
    }
    return $result
}

$script:SettingsPath = Join-Path -Path $PSScriptRoot -ChildPath 'VME-RAS-Config.json'
$script:SslCheckOverride = $null
$script:SslCheck = $true
$script:Settings = $null
$script:SettingsFileLastWriteUtc = $null
$script:SettingsLastPolledAt = [DateTime]::MinValue

# ---------------------------------------------------------------------------
# Logging
#
#   [<L> <CC>/<Ref>/P<PID>] dd-MM-yy HH:mm:ss - <message>
#
# Same format as RAS's own module logs (see docs/LOGGING.md).
# E/W are always written; I/T/D are gated by log_level 3/4/5.
#
# Opening, appending and closing the file on every line (plus a Test-Path and a
# Get-Item for the rotation check) cost about 2/3 of a guests/get. One handle is kept
# open instead. The stream is repositioned to the end before each line, so two
# provider processes appending to the same file (RAS can briefly run an old and a new
# one side by side) never overwrite each other, and the file length that falls out of
# that seek is what drives rotation -- no extra stat call.
#
# Never writes to stdout: that is the CPF protocol channel. Never writes to stderr
# either: an unread stderr pipe fills and deadlocks the provider.
# ---------------------------------------------------------------------------

$script:LogPath = Join-Path -Path $PSScriptRoot -ChildPath 'VME-RAS-Provider.log'
$script:LogLevel = 4
$script:LogRotateMaxBytes = 20MB
$script:LogRotateMaxGenerations = 3
$script:LogStream = $null
$script:LogWriter = $null
$script:LogOpenRetryAt = [DateTime]::MinValue
# An external delete of an open-for-append file (log cleanup job, antivirus, a user
# tidying up) does not fail the next write -- Windows unlinks the directory entry but
# the held handle keeps appending to the now-invisible file until it is closed, and
# nothing throws to trigger the normal reopen-on-error path.
# Without this, the log would only reappear at the next size-triggered rotation, which
# at low log volumes could be a very long time. Checked on a timer, not every line --
# reads $script:Settings.logging.file_exists_check_seconds directly at arm time rather
# than a cached copy, because Import-ProviderSettings logs several lines of its own
# (self-heal notices) before Set-ProviderRuntimeFromSettings would otherwise get a
# chance to update a cached interval, which would lock the first window to a stale
# default regardless of what's configured.
$script:LogExistCheckAt = [DateTime]::MinValue
$script:LogLevelRank = @{ I = 3; T = 4; D = 5 }
$script:PidHex = '{0:X4}' -f $PID

# Secret masking (mandatory). Applied to every line, whatever its
# origin, so no call site can leak a credential by forgetting to mask:
#   - any JSON member whose name contains pass/secret/token/apikey/credential;
#   - form/query pairs (password=..., refresh_token=...);
#   - bearer tokens.
$script:SecretJsonRegex = [regex]::new('("[^"]*(?:pass|secret|token|api_?key|credential)[^"]*"\s*:\s*)(?:"(?:[^"\\]|\\.)*"|[^,}\]\s\[{]+)', 'IgnoreCase, Compiled, CultureInvariant')
$script:SecretFormRegex = [regex]::new('((?:password|passwd|refresh_token|access_token|client_secret)=)[^&\s"]+', 'IgnoreCase, Compiled, CultureInvariant')
$script:SecretBearerRegex = [regex]::new('(bearer\s+)[^\s"'',;]+', 'IgnoreCase, Compiled, CultureInvariant')

function Protect-VmeLogText {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return $Text }
    $t = $script:SecretJsonRegex.Replace($Text, '$1"***"')
    $t = $script:SecretFormRegex.Replace($t, '$1***')
    return $script:SecretBearerRegex.Replace($t, '$1***')
}

function Close-LogWriter {
    if ($null -ne $script:LogWriter) { try { $script:LogWriter.Dispose() } catch { } }
    if ($null -ne $script:LogStream) { try { $script:LogStream.Dispose() } catch { } }
    $script:LogWriter = $null
    $script:LogStream = $null
}

function Open-LogWriter {
    if ([DateTime]::UtcNow -lt $script:LogOpenRetryAt) { return $false }
    try {
        $dir = Split-Path -Path $script:LogPath -Parent
        if (-not [string]::IsNullOrWhiteSpace($dir) -and -not [System.IO.Directory]::Exists($dir)) { [void][System.IO.Directory]::CreateDirectory($dir) }
        $share = [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete
        $script:LogStream = [System.IO.FileStream]::new($script:LogPath, [System.IO.FileMode]::Append, [System.IO.FileAccess]::Write, $share)
        $script:LogWriter = [System.IO.StreamWriter]::new($script:LogStream, $script:Utf8NoBom)
        $script:LogWriter.AutoFlush = $true
        return $true
    }
    catch {
        Close-LogWriter
        # A locked or unwritable log must not cost a failed open on every line.
        $script:LogOpenRetryAt = [DateTime]::UtcNow.AddSeconds(5)
        return $false
    }
}

function Invoke-LogRotation {
    Close-LogWriter
    try {
        $maxGen = [Math]::Max(1, $script:LogRotateMaxGenerations)
        $oldest = "$($script:LogPath).$maxGen"
        if ([System.IO.File]::Exists($oldest)) { [System.IO.File]::Delete($oldest) }
        for ($i = $maxGen - 1; $i -ge 1; $i--) {
            $src = "$($script:LogPath).$i"
            if ([System.IO.File]::Exists($src)) { [System.IO.File]::Move($src, "$($script:LogPath).$($i + 1)", $true) }
        }
        if ([System.IO.File]::Exists($script:LogPath)) { [System.IO.File]::Move($script:LogPath, "$($script:LogPath).1", $true) }
    }
    catch {
        # Best effort: another process may hold a generation open. Logging continues in
        # whatever file the path now names.
    }
    $script:LogOpenRetryAt = [DateTime]::MinValue
    [void](Open-LogWriter)
}

function Write-DebugLog {
    param(
        [string]$Message,
        [ValidateSet('E', 'W', 'I', 'T', 'D')]
        [string]$Level = 'D',
        [string]$Component = '00',
        $Ref = $null
    )

    if ($Level -ne 'E' -and $Level -ne 'W' -and $script:LogLevelRank[$Level] -gt $script:LogLevel) { return }

    try {
        if ($null -eq $script:LogWriter -and -not (Open-LogWriter)) { return }

        if ($null -ne $script:Settings -and [DateTime]::UtcNow -ge $script:LogExistCheckAt) {
            $script:LogExistCheckAt = [DateTime]::UtcNow.AddSeconds([int]$script:Settings.logging.file_exists_check_seconds)
            if (-not [System.IO.File]::Exists($script:LogPath)) {
                Close-LogWriter
                if (-not (Open-LogWriter)) { return }
                $script:LogWriter.WriteLine("[W 00/-/P$($script:PidHex)] $([DateTime]::Now.ToString('dd-MM-yy HH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture)) - Log file was missing at the expected path (deleted externally?) -- recreated. Any lines written to the previous handle after it was deleted are lost (they went to an orphaned, unreachable file).")
            }
        }

        $timestamp = [DateTime]::Now.ToString('dd-MM-yy HH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture)
        $refText = if ($null -eq $Ref -or [string]::IsNullOrEmpty([string]$Ref)) { '-' } else { [string]$Ref }
        $line = "[$Level $Component/$refText/P$($script:PidHex)] $timestamp - $(Protect-VmeLogText $Message)"

        [void]$script:LogStream.Seek(0, [System.IO.SeekOrigin]::End)
        if ($script:LogStream.Position -ge $script:LogRotateMaxBytes) {
            Invoke-LogRotation
            if ($null -eq $script:LogWriter) { return }
        }
        $script:LogWriter.WriteLine($line)
    }
    catch {
        Close-LogWriter
        $script:LogOpenRetryAt = [DateTime]::UtcNow.AddSeconds(5)
    }
}

# ---------------------------------------------------------------------------
# Call-budget statistics
#
# Counts every RPC (per method) and every VME call (per endpoint, with bytes and
# time), summarised at I level every logging.stats_interval_seconds and at shutdown.
# This is how a field log proves the call budget holds: e.g. a 3000-guest sweep must
# show ~3000 guests/get and zero instance GETs.
# ---------------------------------------------------------------------------

$script:StatsRpc = @{}
$script:StatsHttp = @{}
$script:StatsSince = [DateTime]::UtcNow

function Add-RpcStat {
    param([string]$Method, [double]$Ms, [bool]$IsError)
    $s = $script:StatsRpc[$Method]
    if ($null -eq $s) { $s = @{ n = 0; err = 0; ms = 0.0; max = 0.0 }; $script:StatsRpc[$Method] = $s }
    $s.n++
    if ($IsError) { $s.err++ }
    $s.ms += $Ms
    if ($Ms -gt $s.max) { $s.max = $Ms }
}

function Add-HttpStat {
    param([string]$Key, [double]$Ms, [long]$Bytes, [bool]$IsError)
    $s = $script:StatsHttp[$Key]
    if ($null -eq $s) { $s = @{ n = 0; err = 0; ms = 0.0; bytes = [long]0 }; $script:StatsHttp[$Key] = $s }
    $s.n++
    if ($IsError) { $s.err++ }
    $s.ms += $Ms
    $s.bytes += $Bytes
}

function Write-StatsSummary {
    param([string]$Reason = 'interval')

    $secs = [int]([DateTime]::UtcNow - $script:StatsSince).TotalSeconds
    $inv = if ($null -ne $script:Inventory) { $script:Inventory.Count } else { 0 }
    $rpc = @(foreach ($k in ($script:StatsRpc.Keys | Sort-Object)) {
            $s = $script:StatsRpc[$k]
            "$k=$($s.n)" + $(if ($s.err) { "/err$($s.err)" } else { '' }) + " avg$([int]($s.ms / [Math]::Max(1, $s.n)))ms max$([int]$s.max)ms"
        })
    $http = @(foreach ($k in ($script:StatsHttp.Keys | Sort-Object)) {
            $s = $script:StatsHttp[$k]
            "$k=$($s.n)" + $(if ($s.err) { "/err$($s.err)" } else { '' }) + " $([Math]::Round($s.bytes / 1KB))KB avg$([int]($s.ms / [Math]::Max(1, $s.n)))ms"
        })
    Write-DebugLog "STATS ($Reason, last ${secs}s) RPC: $(if ($rpc.Count) { $rpc -join '; ' } else { 'none' })" -Level 'I' -Component '00'
    Write-DebugLog "STATS ($Reason, last ${secs}s) VME: $(if ($http.Count) { $http -join '; ' } else { 'none' })" -Level 'I' -Component '00'
    Write-DebugLog ("STATS ($Reason) state: inventory=$inv, hardware=$($script:HwById.Count), tracked clones=$($script:CloneState.Count), " +
        "deleting=$($script:DeletingIds.Count), basic templates=$($script:BasicTemplates.Count), in maintenance=$($script:MaintenanceIds.Count), " +
        "working set=$([Math]::Round([System.Environment]::WorkingSet / 1MB))MB") -Level 'I' -Component '00'

    $script:StatsRpc = @{}
    $script:StatsHttp = @{}
    $script:StatsSince = [DateTime]::UtcNow
}

function Invoke-StatsSummaryIfDue {
    $interval = $script:Settings.logging.stats_interval_seconds
    if ($interval -le 0) { return }
    if (([DateTime]::UtcNow - $script:StatsSince).TotalSeconds -ge $interval) { Write-StatsSummary }
}

# ---------------------------------------------------------------------------
# Settings load, migration and hot reload
# ---------------------------------------------------------------------------

# Temp file + replace. File.Replace is the atomic swap on NTFS (its backup argument must be
# [NullString]::Value: PowerShell passes $null to a .NET string parameter as ''); a plain Move-Item
# -Force is a delete followed by a rename. One retry covers a transient lock (AV scan).
function Write-TextFileAtomic {
    param([Parameter(Mandatory = $true)][string]$Path, [Parameter(Mandatory = $true)][string]$Text)

    $dir = Split-Path -Path $Path -Parent
    if (-not [string]::IsNullOrWhiteSpace($dir) -and -not [System.IO.Directory]::Exists($dir)) { [void][System.IO.Directory]::CreateDirectory($dir) }
    $temp = "$Path.tmp"
    for ($attempt = 1; $attempt -le 2; $attempt++) {
        try {
            [System.IO.File]::WriteAllText($temp, $Text, $script:Utf8NoBom)
            if ([System.IO.File]::Exists($Path)) { [System.IO.File]::Replace($temp, $Path, [NullString]::Value) }
            else { [System.IO.File]::Move($temp, $Path) }
            return
        }
        catch {
            if ($attempt -ge 2) { throw }
            Start-Sleep -Milliseconds 100
        }
    }
}

function Import-ProviderSettings {
    param([switch]$IsReload)

    $priorSettings = $script:Settings
    $defaults = Get-DefaultProviderSettings
    $verb = if ($IsReload) { 'Settings reload' } else { 'Settings load' }
    $parsed = $null
    $fileExists = Test-Path -LiteralPath $script:SettingsPath

    if ($fileExists) {
        try {
            $raw = [System.IO.File]::ReadAllText($script:SettingsPath)
            if (-not [string]::IsNullOrWhiteSpace($raw)) {
                $parsed = ConvertFrom-JsonFast -Json $raw
                if ($parsed -isnot [System.Collections.IDictionary]) { throw 'the document root is not a JSON object' }
            }
        }
        catch {
            Write-DebugLog "${verb}: failed to parse [$($script:SettingsPath)]: $($_.Exception.Message) -- file left untouched, $(if ($IsReload) { 'previous settings kept' } else { 'using defaults' })." -Level 'E' -Component '0A'
            if ($IsReload) { return }
            $parsed = $null
        }
    }

    $schemaVersion = $script:CurrentSettingsSchemaVersion
    $flat = [ordered]@{}
    $bareLeaves = [System.Collections.Generic.List[string]]::new()
    $legacyNotes = [System.Collections.Generic.List[string]]::new()
    $storedDefaults = @{}
    $followedDefaults = 0
    if ($null -ne $parsed) {
        $schemaVersion = 1
        try { if ($null -ne $parsed['schema_version']) { $schemaVersion = [int]$parsed['schema_version'] } } catch { }
        Get-FlatParsedSettings -Node $parsed -Prefix '' -Result $flat -BareLeaves $bareLeaves -StoredDefaults $storedDefaults
        if ($schemaVersion -lt 2) { $legacyNotes = Convert-LegacyFlatSettings -Flat $flat }
        else {
            $legacyNotes = Update-FlatSettingsForChangedDefaults -Flat $flat -StoredDefaults $storedDefaults -Defaults $defaults
            $followedDefaults = $legacyNotes.Count
        }
    }

    $problems = [System.Collections.Generic.List[string]]::new()
    $resolved = Resolve-SettingsNode -DefaultsNode $defaults -Prefix '' -Flat $flat -Problems $problems
    $resolved['schema_version'] = $script:CurrentSettingsSchemaVersion

    $defaultPaths = Get-DefaultLeafPaths -Node $defaults
    $unknownKeys = @($flat.Keys | Where-Object { -not $defaultPaths.Contains([string]$_) })
    $missingKeys = @(if ($null -ne $parsed) { $defaultPaths | Where-Object { -not $flat.Contains($_) } })

    # Relocating files mid-run would orphan the in-flight clone tracking held in
    # memory. RAS restarts the provider on its next reconnect, which is the safe time.
    if ($IsReload -and $null -ne $priorSettings) {
        $a = ConvertTo-FlatResolvedSettings -Node $resolved.locations
        $b = ConvertTo-FlatResolvedSettings -Node $priorSettings.locations
        $moved = @(foreach ($k in $a.Keys) { if ($k -ne 'settings_reload_interval_seconds' -and "$($a[$k])" -ne "$($b[$k])") { $k } })
        if ($moved.Count -gt 0) {
            Write-DebugLog "Settings reload: locations.$($moved -join ', locations.') changed -- applied at the next process start only." -Level 'W' -Component '0A'
            foreach ($k in $moved) { $resolved.locations[$k] = $priorSettings.locations[$k] }
        }
    }

    $script:Settings = $resolved

    foreach ($p in $problems) { Write-DebugLog "${verb}: $p" -Level 'W' -Component '0A' }

    if ($IsReload) {
        if ($unknownKeys.Count -gt 0) { Write-DebugLog "Settings reload: unknown key(s) ignored: $($unknownKeys -join ', ')" -Level 'W' -Component '0A' }
    }
    elseif (-not $fileExists) {
        try {
            Write-TextFileAtomic -Path $script:SettingsPath -Text ((New-SettingsFileTree -DefaultsNode $defaults -ResolvedNode $resolved) | ConvertTo-Json -Depth 12)
            Write-DebugLog "Settings load: created [$($script:SettingsPath)] with defaults (schema_version $($script:CurrentSettingsSchemaVersion))." -Level 'I' -Component '0A'
        }
        catch { Write-DebugLog "Settings load: failed to create [$($script:SettingsPath)]: $($_.Exception.Message)" -Level 'E' -Component '0A' }
    }
    elseif ($null -ne $parsed -and ($schemaVersion -ne $script:CurrentSettingsSchemaVersion -or $unknownKeys.Count -gt 0 -or $missingKeys.Count -gt 0 -or $bareLeaves.Count -gt 0 -or $followedDefaults -gt 0)) {
        # Self-heal: rewrite in the current shape. Every value the admin set (at its
        # current or its legacy path) survives; new keys appear at their default; keys
        # this version no longer has are dropped and listed.
        try {
            Write-TextFileAtomic -Path $script:SettingsPath -Text ((New-SettingsFileTree -DefaultsNode $defaults -ResolvedNode $resolved) | ConvertTo-Json -Depth 12)
            $detail = [System.Collections.Generic.List[string]]::new()
            foreach ($n in $legacyNotes) { $detail.Add("  ~ $n") }
            foreach ($k in $missingKeys) { $detail.Add("  + $k") }
            foreach ($k in $unknownKeys) { $detail.Add("  - $k (removed)") }
            Write-DebugLog "Settings load: rewrote [$($script:SettingsPath)] from schema_version $schemaVersion to $($script:CurrentSettingsSchemaVersion) ({value, default, description} leaves). Values the admin set are preserved; values still at a changed default follow it (~).$(if ($detail.Count) { "`n" + ($detail -join "`n") })" -Level 'W' -Component '0A'
        }
        catch { Write-DebugLog "Settings load: failed to migrate [$($script:SettingsPath)]: $($_.Exception.Message) -- running on the resolved values in memory." -Level 'E' -Component '0A' }
    }

    if (-not $IsReload) {
        $flatSettings = ConvertTo-FlatResolvedSettings -Node $resolved
        $flatDefaults = ConvertTo-FlatResolvedSettings -Node (Resolve-SettingsNode -DefaultsNode (Get-DefaultProviderSettings) -Prefix '' -Flat @{} -Problems ([System.Collections.Generic.List[string]]::new()))
        $nonDefault = @(foreach ($key in $flatSettings.Keys) {
                if ($flatDefaults.Contains($key) -and "$($flatSettings[$key])" -ne "$($flatDefaults[$key])") { "$key = [$($flatSettings[$key])] (default: [$($flatDefaults[$key])])" }
            })
        if ($nonDefault.Count -gt 0) { Write-DebugLog "Settings load: non-default value(s):`n  $($nonDefault -join "`n  ")" -Level 'I' -Component '0A' }
        else { Write-DebugLog 'Settings load: every value at its default.' -Level 'I' -Component '0A' }
    }
    elseif ($null -ne $priorSettings) {
        $flatBefore = ConvertTo-FlatResolvedSettings -Node $priorSettings
        $flatAfter = ConvertTo-FlatResolvedSettings -Node $resolved
        $changed = @(foreach ($key in $flatAfter.Keys) {
                $before = if ($flatBefore.Contains($key)) { "$($flatBefore[$key])" } else { '<absent>' }
                if ($before -ne "$($flatAfter[$key])") { "$key changed [$before] -> [$($flatAfter[$key])]" }
            })
        if ($changed.Count -gt 0) { Write-DebugLog "Settings reload: value(s) applied:`n  $($changed -join "`n  ")" -Level 'I' -Component '0A' }
    }
}

# Derives everything that is computed from settings once, instead of per call.
function Set-ProviderRuntimeFromSettings {
    $S = $script:Settings

    $dataDir = [string]$S.locations.data_directory
    if ([string]::IsNullOrWhiteSpace($dataDir)) { $dataDir = $PSScriptRoot }
    if (-not [System.IO.Path]::IsPathRooted($dataDir)) { $dataDir = Join-Path -Path $PSScriptRoot -ChildPath $dataDir }
    $resolve = {
        param([string]$FileName)
        if ([System.IO.Path]::IsPathRooted($FileName)) { return $FileName }
        return (Join-Path -Path $dataDir -ChildPath $FileName)
    }

    $newLogPath = & $resolve $S.locations.log_file
    if ($newLogPath -ne $script:LogPath) { Close-LogWriter; $script:LogPath = $newLogPath; $script:LogOpenRetryAt = [DateTime]::MinValue }
    $script:CloneStatePath = & $resolve $S.locations.clone_state_file
    $script:MaintenanceStatePath = & $resolve $S.locations.maintenance_state_file
    $script:TemplateStatePath = & $resolve $S.locations.template_state_file
    $script:StuckCloneStatePath = & $resolve $S.locations.stuck_clone_state_file
    $script:FormerTemplateStatePath = & $resolve $S.locations.former_template_state_file
    $script:TemplateVersionStatePath = & $resolve $S.locations.template_version_state_file

    $script:LogLevel = [int]$S.logging.log_level
    $script:LogRotateMaxBytes = [long]$S.logging.log_rotate_max_mb * 1MB
    $script:LogRotateMaxGenerations = [int]$S.logging.log_rotate_max_generations
    $script:TemplateMethod = [string]$S.capabilities.template_method
    # A per-connection ssl_check (RAS custom setting) wins over the file until the
    # process ends, including across hot reloads.
    $script:SslCheck = if ($null -ne $script:SslCheckOverride) { [bool]$script:SslCheckOverride } else { [bool]$S.connection.ssl_check }

    # Compiled once per load; an invalid pattern is reported once, not per instance.
    $regexes = [System.Collections.Generic.List[regex]]::new()
    foreach ($pattern in @($S.inventory.exclude_instance_name_patterns)) {
        if ([string]::IsNullOrWhiteSpace($pattern)) { continue }
        try { $regexes.Add([regex]::new($pattern, 'IgnoreCase, CultureInvariant')) }
        catch { Write-DebugLog "inventory.exclude_instance_name_patterns entry [$pattern] is not a valid regex -- ignored." -Level 'W' -Component '0A' }
    }
    $script:ExcludeRegexes = $regexes.ToArray()

    $scope = [System.Collections.Generic.List[string]]::new()
    if ([int]$S.inventory.scope.site_id -gt 0) { $scope.Add("siteId=$([int]$S.inventory.scope.site_id)") }
    if ([int]$S.inventory.scope.zone_id -gt 0) { $scope.Add("zoneId=$([int]$S.inventory.scope.zone_id)") }
    $script:ServerScopeQuery = if ($scope.Count) { '&' + ($scope -join '&') } else { '' }
    if (-not [string]::IsNullOrWhiteSpace([string]$S.inventory.scope.instance_type)) { $scope.Add("instanceType=$([uri]::EscapeDataString([string]$S.inventory.scope.instance_type))") }
    if (-not [string]::IsNullOrWhiteSpace([string]$S.inventory.scope.label)) { $scope.Add("labels=$([uri]::EscapeDataString([string]$S.inventory.scope.label))") }
    $script:InstanceScopeQuery = if ($scope.Count) { '&' + ($scope -join '&') } else { '' }

    $onOff = { param([bool]$b) if ($b) { 'on' } else { 'off' } }
    $script:DeleteQuery = "force=on&preserveVolumes=off&keepBackups=$(& $onOff $S.delete.keep_backups)&releaseEIPs=$(& $onOff $S.delete.release_eips)"

    $ras = [int]$S.capabilities.tasks_polling_rate * [int]$S.capabilities.tasks_polling_retries
    if ($ras -lt [int]$S.cloning.provision_timeout_seconds) {
        Write-DebugLog "Settings: RAS gives up on a task after tasks_polling_rate x tasks_polling_retries = ${ras}s, before cloning.provision_timeout_seconds ($($S.cloning.provision_timeout_seconds)s) -- a slow clone will be reported by RAS as a timeout instead of by the provider with a reason." -Level 'W' -Component '0A'
    }
    if ([int]$S.orphan_detection.stale_poll_after_seconds -le [int]$S.capabilities.guests_polling_rate) {
        Write-DebugLog "Settings: orphan_detection.stale_poll_after_seconds is not above capabilities.guests_polling_rate -- every guest would look forgotten between two polls." -Level 'W' -Component '0A'
    }
}

# Re-stats the settings file at most once per settings_reload_interval_seconds and
# reloads it if its last-write time moved. There is no background thread: this rides
# on whichever request arrives next.
function Update-ProviderSettingsIfChanged {
    if (([DateTime]::UtcNow - $script:SettingsLastPolledAt).TotalSeconds -lt $script:Settings.locations.settings_reload_interval_seconds) { return }
    $script:SettingsLastPolledAt = [DateTime]::UtcNow

    try {
        if (-not [System.IO.File]::Exists($script:SettingsPath)) { return }
        $writeUtc = [System.IO.File]::GetLastWriteTimeUtc($script:SettingsPath)
        if ($null -ne $script:SettingsFileLastWriteUtc -and $writeUtc -eq $script:SettingsFileLastWriteUtc) { return }
        $script:SettingsFileLastWriteUtc = $writeUtc
        Import-ProviderSettings -IsReload
        Set-ProviderRuntimeFromSettings
    }
    catch {
        Write-DebugLog "Settings reload check failed: $($_.Exception.Message)" -Level 'W' -Component '0A'
    }
}

# ---------------------------------------------------------------------------
# Runtime state
#
# Everything per-instance is keyed by the VME instance id (string). Every map here is
# either bounded by the inventory (garbage-collected against it, see
# Invoke-CacheGarbageCollection) or self-expiring -- a provider that runs for months
# must not grow with the number of instances that ever existed.
# ---------------------------------------------------------------------------

$script:VmeSession = $null
$script:WebSession = $null
$script:StdoutBroken = $false
$script:ErrorCodes = @{ ParseError = -32700; InvalidRequest = -32600; MethodNotFound = -32601; InvalidParams = -32602; InternalError = -32603 }

# Inventory: id -> slim record (see ConvertTo-VmeInstanceRecord). Raw API objects are
# never kept.
$script:Inventory = @{}
$script:InventoryLoaded = $false
$script:InventoryListedAt = [DateTime]::MinValue
$script:InventoryAttemptAt = [DateTime]::MinValue
$script:MassDropPending = $false

# Hardware: id -> @{ server_id; uuid; macs; os; complete }. MACs don't change for the
# life of an instance, so this is learned once and only dropped with the instance.
$script:HwById = @{}
$script:HwLastTry = @{}
$script:HwBulkAt = [DateTime]::MinValue
$script:HwBulkDone = $false

# Template id -> clone body inputs, refreshed for free from every listing.
$script:CloneBodyBySource = @{}

# Ids this provider sent a DELETE for: id -> @{ name; submitted_at; last_probe; stuck_logged }.
$script:DeletingIds = @{}

# Clone tracking. CloneState is the persisted source of truth (key -> entry), keyed
# 'pending:<name>' until the clone's id is known, then by that id. The two indexes are
# rebuilt from it and never persisted.
$script:CloneState = @{}
$script:CloneKeyByTaskId = @{}
$script:CloneKeyByCloneId = @{}
$script:CloneLastRefresh = @{}          # clone id -> UTC of the last lean GET
$script:SweepLastRunAt = [DateTime]::MinValue
$script:CompletedCloneTaskOutputs = @{}  # task id -> @{ clone_id; stored_at }
$script:KnownCloneSources = [System.Collections.Generic.HashSet[string]]::new()

# Clones the provider gave up on (task failed to RAS) that VME has not confirmed ended:
# task id -> entry. Each holds a max_concurrent_clone_operations slot until VME says the
# clone finished, failed or is gone -- never on a timer. Persisted.
$script:StuckClones = @{}
$script:StuckClonesCheckedAt = [DateTime]::MinValue

$script:RecentlyDeletedMacByName = @{}   # name -> @{ mac; deleted_at }
$script:PowerWatch = @{}                 # id -> @{ action; from; target; transitional; since }, provider-issued power actions
$script:WatchReadAt = @{}                # id -> UTC of the last watch re-read (power watch, template in maintenance)
$script:LastGuestPollAt = @{}            # id -> UTC, for the orphan audit
$script:StatusLogged = @{}               # id -> abnormal status already logged
$script:ExcludeLogged = [System.Collections.Generic.HashSet[string]]::new()
$script:OrphanReported = [System.Collections.Generic.HashSet[string]]::new()
$script:OrphanAuditLastCheckedAt = [DateTime]::MinValue
$script:TemplateFlappingUnmaskedVmIds = [System.Collections.Generic.HashSet[string]]::new()
$script:ProviderStartedAt = [DateTime]::UtcNow

# id -> uuid ('' until learned). Persisted with the uuid so a reused id (appliance
# rebuild or restore) never inherits is_template or a maintenance override.
$script:MaintenanceIds = @{}
$script:BasicTemplates = @{}
# Former templates RAS sent a delete for: never deleted by this provider.
$script:FormerTemplates = @{}

# template_method 'versioning': template id -> @{ uuid; versions = @{ RAS snapshot name ->
# version } } (persisted), the image id -> {template_id, ras_name} index rebuilt from it,
# and the in-memory snapshot delete/revert tasks (task id -> entry).
$script:TemplateVersions = @{}
$script:VersionByImageId = @{}
$script:SnapshotTasks = @{}

# ---------------------------------------------------------------------------
# CPF transport
# ---------------------------------------------------------------------------

# Invariant: every handler returns exactly one object (a response hashtable). A stray
# uncaptured value inside a handler would turn the response into an array and break
# the frame -- new code must [void] or assign every call whose output it doesn't use.
function Send-Response {
    param([Parameter(Mandatory = $true)][object]$Response, [string]$Method = '')

    try {
        # -WarningAction Stop: a depth overflow fails loudly instead of truncating.
        $json = ConvertTo-Json -InputObject $Response -Depth 12 -Compress -WarningAction Stop
    }
    catch {
        $json = ConvertTo-Json -InputObject @{ error = @{ code = $script:ErrorCodes.InternalError; message = "$($script:ProviderNamePrefix) Failed to serialize response: $($_.Exception.Message)" } } -Compress
        Write-DebugLog "Response serialization failed for [$Method]: $($_.Exception.Message)" -Level 'E' -Component '01'
    }

    if ($script:LogLevel -ge 5 -and ($Method -ne 'guests/get' -or $script:Settings.logging.log_rpc_bodies)) {
        Write-DebugLog "OUT: $json" -Level 'D' -Component '01'
    }

    try { $writer.WriteLine($json) }
    catch {
        # RAS closed the pipe. Nothing can be answered any more; the main loop exits.
        $script:StdoutBroken = $true
        Write-DebugLog "stdout write failed ($($_.Exception.Message)) -- RAS has closed the pipe; stopping." -Level 'E' -Component '01'
    }
}

function New-ErrorResponse {
    param([int]$Code, [string]$Message)
    return @{ error = @{ code = $Code; message = $Message } }
}

# Validates dotted field paths (e.g. 'params.id') against a parsed request.
function Test-RequiredFields {
    param([object]$Data, [string[]]$Fields)

    foreach ($field in $Fields) {
        $current = $Data
        foreach ($part in $field.Split('.')) {
            if ($current -isnot [System.Collections.IDictionary]) { return $field }
            $current = $current[$part]
        }
        if ($null -eq $current) { return $field }
    }
    return $null
}

# ---------------------------------------------------------------------------
# Persisted state
#
# Timestamps are stored as epoch seconds: no DateTime ever round-trips through a
# culture-sensitive string. Files from provider 2.x with ISO strings are converted
# on load.
# ---------------------------------------------------------------------------

function Read-JsonFileSafe {
    param([Parameter(Mandatory = $true)][string]$Path, [string]$Label = 'state file')

    if (-not [System.IO.File]::Exists($Path)) { return @{} }
    try {
        $raw = [System.IO.File]::ReadAllText($Path)
        if ([string]::IsNullOrWhiteSpace($raw)) { return @{} }
        $obj = ConvertFrom-JsonFast -Json $raw
        if ($obj -is [System.Collections.IDictionary]) { return $obj }
        throw 'the document root is not a JSON object'
    }
    catch {
        # Keep the unreadable file for diagnosis instead of overwriting it at the next save.
        $aside = "$Path.corrupt-$(Get-EpochNow)"
        try { [System.IO.File]::Move($Path, $aside) } catch { $aside = '(could not be moved aside)' }
        Write-DebugLog "Failed to read $Label [$Path]: $($_.Exception.Message) -- starting empty; the unreadable file was kept as [$aside]." -Level 'E' -Component '00'
        return @{}
    }
}

function Write-JsonFileAtomic {
    param([Parameter(Mandatory = $true)][string]$Path, [Parameter(Mandatory = $true)][object]$Value, [string]$Label = 'state file')
    try {
        Write-TextFileAtomic -Path $Path -Text (ConvertTo-Json -InputObject $Value -Depth 8)
        return $true
    }
    catch {
        Write-DebugLog "Failed to write $Label [$Path]: $($_.Exception.Message) -- the in-memory state is intact but will not survive a restart until a later write succeeds." -Level 'E' -Component '00'
        return $false
    }
}

function Save-CloneState { [void](Write-JsonFileAtomic -Path $script:CloneStatePath -Value $script:CloneState -Label 'clone state') }

function Add-CloneIndexes {
    param([string]$Key, [System.Collections.IDictionary]$Entry)
    if (-not [string]::IsNullOrWhiteSpace([string]$Entry['task_id'])) { $script:CloneKeyByTaskId[[string]$Entry['task_id']] = $Key }
    if (-not [string]::IsNullOrWhiteSpace([string]$Entry['clone_id'])) { $script:CloneKeyByCloneId[[string]$Entry['clone_id']] = $Key }
}

function Remove-CloneIndexes {
    param([string]$Key, [System.Collections.IDictionary]$Entry)
    $tid = [string]$Entry['task_id']
    if ($tid -and $script:CloneKeyByTaskId[$tid] -eq $Key) { $script:CloneKeyByTaskId.Remove($tid) }
    $cid = [string]$Entry['clone_id']
    if ($cid -and $script:CloneKeyByCloneId[$cid] -eq $Key) { $script:CloneKeyByCloneId.Remove($cid) }
}

# Writes the file unless -NoSave: callers persist on transitions only (created,
# resolved, confirmed, completed, retired), never per poll.
function Set-CloneEntry {
    param([Parameter(Mandatory = $true)][string]$Key, [Parameter(Mandatory = $true)][System.Collections.IDictionary]$Entry, [switch]$NoSave)
    $old = $script:CloneState[$Key]
    if ($null -ne $old) { Remove-CloneIndexes -Key $Key -Entry $old }
    $script:CloneState[$Key] = $Entry
    Add-CloneIndexes -Key $Key -Entry $Entry
    if (-not $NoSave) { Save-CloneState }
}

function Remove-CloneEntry {
    param([Parameter(Mandatory = $true)][string]$Key)
    $old = $script:CloneState[$Key]
    if ($null -eq $old) { return }
    Remove-CloneIndexes -Key $Key -Entry $old
    $script:CloneState.Remove($Key)
    $cid = [string]$old['clone_id']
    if ($cid) { $script:CloneLastRefresh.Remove($cid) }
    Save-CloneState
}

function Move-CloneEntryKey {
    param([Parameter(Mandatory = $true)][string]$FromKey, [Parameter(Mandatory = $true)][string]$ToKey)
    $entry = $script:CloneState[$FromKey]
    if ($null -eq $entry) { return }
    Remove-CloneIndexes -Key $FromKey -Entry $entry
    $script:CloneState.Remove($FromKey)
    Set-CloneEntry -Key $ToKey -Entry $entry
}

function Get-CloneKeyByTaskId { param([string]$TaskId) return $script:CloneKeyByTaskId[$TaskId] }
function Get-CloneKeyByCloneId { param([string]$VmId) return $script:CloneKeyByCloneId[$VmId] }

# Loads and normalises the clone state. Provider 2.x entries (ISO clone_started_at, attempt
# counters, numeric process task ids) are converted or dropped here, once.
function Import-CloneState {
    $script:CloneState = @{}
    $script:CloneKeyByTaskId = @{}
    $script:CloneKeyByCloneId = @{}

    $doc = Read-JsonFileSafe -Path $script:CloneStatePath -Label 'clone state'
    $changed = $false
    foreach ($key in @($doc.Keys)) {
        $e = $doc[$key]
        $taskId = [string](Get-JsonValue $e 'task_id')
        if ($e -isnot [System.Collections.IDictionary] -or [string]$e['type'] -ne 'clone' -or ($taskId -notlike 'clonebyname:*' -and $taskId -notlike 'linked:*')) {
            Write-DebugLog "Clone state: dropped entry [$key] -- not a clone tracked by name or a linked clone (a 2.x process-id task, or malformed)." -Level 'W' -Component '04' -Ref $key
            $changed = $true
            continue
        }
        $started = ConvertTo-EpochSeconds $(if ($null -ne $e['started_epoch']) { $e['started_epoch'] } else { $e['clone_started_at'] })
        if ($null -eq $started) {
            Write-DebugLog "Clone state: dropped entry [$key] -- no readable start time." -Level 'W' -Component '04' -Ref $key
            $changed = $true
            continue
        }
        $reported = ConvertTo-EpochSeconds $(if ($null -ne $e['reported_completed_epoch']) { $e['reported_completed_epoch'] } else { $e['reported_completed_at'] })
        $entry = @{
            type                     = 'clone'
            task_id                  = $taskId
            source_id                = [string]$e['source_id']
            source_uuid              = [string]$e['source_uuid']
            clone_id                 = [string]$e['clone_id']
            clone_uuid               = [string]$e['clone_uuid']
            name                     = [string]$e['name']
            started_epoch            = [long]$started
            resolved_epoch           = ConvertTo-EpochSeconds $e['resolved_epoch']
            preexisting_ids          = @(@($e['preexisting_ids']) | Where-Object { $null -ne $_ } | ForEach-Object { [string]$_ })
            max_known_id             = [long]$(if ($null -ne $e['max_known_id']) { $e['max_known_id'] } else { 0 })
            submit_uncertain         = [bool]$e['submit_uncertain']
            provision_confirmed      = [bool]$e['provision_confirmed']
            reported_completed_epoch = $reported
            preserved_mac            = [string]$e['preserved_mac']
            requested_network        = [string]$e['requested_network']
            requested_network_name   = [string]$e['requested_network_name']
            image_id                 = [string]$e['image_id']
            version                  = [string]$e['version']
        }
        if ($null -eq $e['started_epoch']) { $changed = $true }
        $script:CloneState[[string]$key] = $entry
        Add-CloneIndexes -Key ([string]$key) -Entry $entry
        if ($entry.source_id) { [void]$script:KnownCloneSources.Add($entry.source_id) }
    }
    if ($changed) { Save-CloneState }
}

# Maintenance and basic-template sets: { "entries": [ { "id": "19", "uuid": "..." } ] }.
# The 2.x form { "ids": [...] } is read too (uuid learned at the next listing).
function Import-IdUuidMap {
    param([string]$Path, [string]$Label)
    $map = @{}
    $doc = Read-JsonFileSafe -Path $Path -Label $Label
    foreach ($e in @($doc['entries'])) {
        $id = [string](Get-JsonValue $e 'id')
        if (-not [string]::IsNullOrWhiteSpace($id)) { $map[$id] = [string](Get-JsonValue $e 'uuid') }
    }
    foreach ($id in @($doc['ids'])) {
        if (-not [string]::IsNullOrWhiteSpace([string]$id) -and -not $map.ContainsKey([string]$id)) { $map[[string]$id] = '' }
    }
    return $map
}

function Save-IdUuidMap {
    param([string]$Path, [hashtable]$Map, [string]$Label)
    $entries = @(foreach ($id in ($Map.Keys | Sort-Object)) { [ordered]@{ id = [string]$id; uuid = [string]$Map[$id] } })
    [void](Write-JsonFileAtomic -Path $Path -Value ([ordered]@{ entries = $entries }) -Label $Label)
}

function Save-MaintenanceState { Save-IdUuidMap -Path $script:MaintenanceStatePath -Map $script:MaintenanceIds -Label 'maintenance state' }
function Save-BasicTemplateState { Save-IdUuidMap -Path $script:TemplateStatePath -Map $script:BasicTemplates -Label 'template state' }
function Save-FormerTemplateState { Save-IdUuidMap -Path $script:FormerTemplateStatePath -Map $script:FormerTemplates -Label 'former template state' }

# Template versions: { "entries": [ { "template_id", "uuid", "versions": [ <version> ] } ] }.
# A version is keyed by the name RAS uses (RAS_TEMPLATE_VERSION_<n>); vme_name is the
# snapshot (and image) name on VME. status: 'creating' while guests/snapshots/create
# runs (its task fields live on the version, so a provider restart resumes it), 'ready',
# 'deleting'.
# clone_inputs: the template's clone settings when the version was made (plan, size,
# disks, firmware, pool, server-level network) -- see Get-VmeVersionCloneInputs.
$script:VersionFields = @('ras_name', 'vme_name', 'status', 'task_id', 'phase', 'process_id', 'snapshot_id', 'snapshot_external_id', 'image_id', 'image_name', 'linked_requested', 'linked_epoch', 'started_epoch', 'created_epoch', 'clone_inputs')

function Update-VmeVersionImageIndex {
    $script:VersionByImageId = @{}
    foreach ($tid in $script:TemplateVersions.Keys) {
        foreach ($v in $script:TemplateVersions[$tid].versions.Values) {
            if (-not [string]::IsNullOrWhiteSpace([string]$v.image_id)) { $script:VersionByImageId[[string]$v.image_id] = @{ template_id = [string]$tid; ras_name = [string]$v.ras_name } }
        }
    }
}

function Save-TemplateVersions {
    $entries = @(foreach ($tid in ($script:TemplateVersions.Keys | Sort-Object)) {
            $t = $script:TemplateVersions[$tid]
            [ordered]@{ template_id = [string]$tid; uuid = [string]$t.uuid; versions = @(foreach ($n in ($t.versions.Keys | Sort-Object)) { $t.versions[$n] }) }
        })
    [void](Write-JsonFileAtomic -Path $script:TemplateVersionStatePath -Value ([ordered]@{ entries = $entries }) -Label 'template version state')
    Update-VmeVersionImageIndex
}

function Import-TemplateVersions {
    $script:TemplateVersions = @{}
    $doc = Read-JsonFileSafe -Path $script:TemplateVersionStatePath -Label 'template version state'
    foreach ($e in @($doc['entries'])) {
        $tid = [string](Get-JsonValue $e 'template_id')
        if ([string]::IsNullOrWhiteSpace($tid)) { continue }
        $versions = @{}
        foreach ($v in @(Get-JsonValue $e 'versions')) {
            $name = [string](Get-JsonValue $v 'ras_name')
            if ([string]::IsNullOrWhiteSpace($name)) { continue }
            $ver = @{}
            foreach ($f in $script:VersionFields) { $ver[$f] = Get-JsonValue $v $f }
            foreach ($f in @('ras_name', 'vme_name', 'status', 'task_id', 'phase', 'process_id', 'snapshot_id', 'snapshot_external_id', 'image_id', 'image_name')) { $ver[$f] = [string]$ver[$f] }
            $ver.linked_requested = [bool]$ver.linked_requested
            $ci = $ver.clone_inputs
            $ver.clone_inputs = $null
            if ($ci -is [System.Collections.IDictionary] -and $ci.Count -gt 0) {
                $h = @{}
                foreach ($k in $ci.Keys) { $h[[string]$k] = $ci[$k] }
                $ver.clone_inputs = $h
            }
            $versions[$name] = $ver
        }
        $script:TemplateVersions[$tid] = @{ uuid = [string](Get-JsonValue $e 'uuid'); versions = $versions }
    }
    Update-VmeVersionImageIndex
}

# The template's version record, created on demand.
function Get-VmeTemplateVersionSet {
    param([Parameter(Mandatory = $true)][string]$TemplateId)
    $t = $script:TemplateVersions[$TemplateId]
    if ($null -eq $t) {
        $t = @{ uuid = (Get-RecordUuid -VmId $TemplateId); versions = @{} }
        $script:TemplateVersions[$TemplateId] = $t
    }
    return $t
}

# ---------------------------------------------------------------------------
# Clone slots (cloning.max_concurrent_clone_operations)
#
# RAS clones one guest after another: when the provider gives up on a clone, RAS logs
# "Failed to create clone" and sends the next one. If VME's clone is merely stuck (VME
# provision processes can stay stuck for hours, with cancel and retry both refused), every
# give-up used to leave one more stuck clone on the template. A slot is therefore held
# from the submit until VME confirms the clone ended -- not until the provider stops
# waiting for it. Given-up clones move from CloneState to StuckClones and keep their
# slot; the limit is enforced in guests/clone before any VME call.
# ---------------------------------------------------------------------------

function Save-StuckClones {
    $entries = @(foreach ($k in ($script:StuckClones.Keys | Sort-Object)) { $script:StuckClones[$k] })
    [void](Write-JsonFileAtomic -Path $script:StuckCloneStatePath -Value ([ordered]@{ entries = $entries }) -Label 'stuck clone state')
}

function Import-StuckClones {
    $script:StuckClones = @{}
    $doc = Read-JsonFileSafe -Path $script:StuckCloneStatePath -Label 'stuck clone state'
    foreach ($e in @($doc['entries'])) {
        $taskId = [string](Get-JsonValue $e 'task_id')
        $started = ConvertTo-EpochSeconds (Get-JsonValue $e 'started_epoch')
        if ([string]::IsNullOrWhiteSpace($taskId) -or $null -eq $started) { continue }
        $script:StuckClones[$taskId] = @{
            task_id         = $taskId
            source_id       = [string](Get-JsonValue $e 'source_id')
            name            = [string](Get-JsonValue $e 'name')
            clone_id        = [string](Get-JsonValue $e 'clone_id')
            started_epoch   = [long]$started
            abandoned_epoch = [long]$(if ($null -ne (Get-JsonValue $e 'abandoned_epoch')) { Get-JsonValue $e 'abandoned_epoch' } else { $started })
            reason          = [string](Get-JsonValue $e 'reason')
            last_status     = [string](Get-JsonValue $e 'last_status')
            preexisting_ids = @(@(Get-JsonValue $e 'preexisting_ids') | Where-Object { $null -ne $_ } | ForEach-Object { [string]$_ })
            max_known_id    = [long]$(if ($null -ne (Get-JsonValue $e 'max_known_id')) { Get-JsonValue $e 'max_known_id' } else { 0 })
        }
    }
}

# Called by every path that stops waiting for a clone VME has not confirmed finished.
function Add-VmeStuckClone {
    param([Parameter(Mandatory = $true)][System.Collections.IDictionary]$Entry, [Parameter(Mandatory = $true)][string]$Reason, [string]$Status = '')
    $taskId = [string]$Entry.task_id
    if ([string]::IsNullOrWhiteSpace($taskId)) { return }
    $script:StuckClones[$taskId] = @{
        task_id         = $taskId
        source_id       = [string]$Entry.source_id
        name            = [string]$Entry.name
        clone_id        = [string]$Entry.clone_id
        started_epoch   = [long]$Entry.started_epoch
        abandoned_epoch = Get-EpochNow
        reason          = $Reason
        last_status     = $Status
        preexisting_ids = @($Entry.preexisting_ids)
        max_known_id    = [long]$Entry.max_known_id
    }
    Save-StuckClones
    $what = if ($Entry.clone_id) { "instance [$($Entry.clone_id)]$(if ($Status) { ", status [$Status]" })" } else { 'never appeared on VME' }
    Write-DebugLog "Clone [$($Entry.name)] ($what) given up ($Reason) but not confirmed ended on VME -- it keeps one of the $($script:Settings.cloning.max_concurrent_clone_operations) clone slot(s) until VME finishes, fails or deletes it." -Level 'W' -Component '04' -Ref $taskId
}

function Remove-VmeStuckClone {
    param([Parameter(Mandatory = $true)][string]$Key, [Parameter(Mandatory = $true)][string]$Why)
    $s = $script:StuckClones[$Key]
    if ($null -eq $s) { return }
    $script:StuckClones.Remove($Key)
    Save-StuckClones
    $age = [Math]::Round(((Get-EpochNow) - [long]$s.started_epoch) / 60.0, 1)
    Write-DebugLog "Clone slot released: given-up clone [$($s.name)]$(if ($s.clone_id) { " (instance [$($s.clone_id)])" }) -- $Why, $age min after its submit. $((Get-VmeCloneSlotHolders).Count) of $($script:Settings.cloning.max_concurrent_clone_operations) slot(s) now in use." -Level 'I' -Component '04' -Ref $s.task_id
}

# Everything holding a slot: clones RAS is still waiting for (VME provisioning not
# confirmed finished -- includes clones the deploy-event gate reported early) and
# given-up clones VME has not confirmed ended.
function Get-VmeCloneSlotHolders {
    $now = Get-EpochNow
    $holders = [System.Collections.Generic.List[string]]::new()
    foreach ($e in $script:CloneState.Values) {
        if ($e.provision_confirmed) { continue }
        $holders.Add("[$($e.name)] ($(if ($e.clone_id) { "instance $($e.clone_id)" } else { 'not found yet' }), in flight $([Math]::Round(($now - [long]$e.started_epoch) / 60.0, 1)) min)")
    }
    foreach ($s in $script:StuckClones.Values) {
        $what = if ($s.clone_id) { "instance $($s.clone_id)$(if ($s.last_status) { " $($s.last_status)" })" } else { 'never appeared' }
        $holders.Add("[$($s.name)] ($what, given up, submitted $([Math]::Round(($now - [long]$s.started_epoch) / 60.0, 1)) min ago)")
    }
    return , $holders
}

function Test-VmeInMaintenanceMode { param([string]$VmId) return $script:MaintenanceIds.ContainsKey($VmId) }

# is_template's only source, under 'basic' and 'versioning' alike (guests/convert keeps
# the set; versions don't decide it). Always false in 'none' mode, whatever the file says.
function Test-VmeIsBasicTemplate {
    param([string]$VmId)
    if ($script:TemplateMethod -eq 'none') { return $false }
    return $script:BasicTemplates.ContainsKey($VmId)
}

function Get-RecordUuid { param([string]$VmId) $r = $script:Inventory[$VmId]; if ($null -ne $r) { return [string]$r.uuid } return '' }

# Checks the persisted id sets against the uuids in the inventory: a uuid mismatch
# means the id now belongs to a different instance (appliance rebuilt or restored),
# so the entry is dropped; an entry without a uuid (2.x file) adopts the current one.
function Update-IdUuidIntegrity {
    foreach ($pair in @(@{ Map = $script:MaintenanceIds; Label = 'maintenance'; Save = 'Save-MaintenanceState' }, @{ Map = $script:BasicTemplates; Label = 'basic-template'; Save = 'Save-BasicTemplateState' }, @{ Map = $script:FormerTemplates; Label = 'former-template'; Save = 'Save-FormerTemplateState' })) {
        $map = $pair.Map
        $dirty = $false
        foreach ($id in @($map.Keys)) {
            $rec = $script:Inventory[$id]
            if ($null -eq $rec -or [string]::IsNullOrWhiteSpace([string]$rec.uuid)) { continue }
            $stored = [string]$map[$id]
            if ([string]::IsNullOrWhiteSpace($stored)) { $map[$id] = [string]$rec.uuid; $dirty = $true; continue }
            if ($stored -ne [string]$rec.uuid) {
                Write-DebugLog "Instance id [$id] now has uuid [$($rec.uuid)], not the [$stored] recorded in the $($pair.Label) set -- the id was reused (appliance rebuilt or restored?). Entry dropped." -Level 'W' -Component '06' -Ref $id
                $map.Remove($id)
                $dirty = $true
            }
        }
        if ($dirty) { & $pair.Save }
    }
    # Template versions: a reused id must not inherit another VM's snapshots and images.
    $vDirty = $false
    foreach ($id in @($script:TemplateVersions.Keys)) {
        $rec = $script:Inventory[$id]
        if ($null -eq $rec -or [string]::IsNullOrWhiteSpace([string]$rec.uuid)) { continue }
        $t = $script:TemplateVersions[$id]
        if ([string]::IsNullOrWhiteSpace([string]$t.uuid)) { $t.uuid = [string]$rec.uuid; $vDirty = $true; continue }
        if ([string]$t.uuid -ne [string]$rec.uuid) {
            Write-DebugLog "Instance id [$id] now has uuid [$($rec.uuid)], not the [$($t.uuid)] recorded with its template versions ($(@($t.versions.Keys) -join ', ')) -- the id was reused. Versions dropped from the provider state (nothing deleted on VME)." -Level 'W' -Component '06' -Ref $id
            $script:TemplateVersions.Remove($id)
            $vDirty = $true
        }
    }
    if ($vDirty) { Save-TemplateVersions }
}

function Set-CompletedCloneTaskOutput {
    param([string]$TaskId, [string]$CloneId)
    if ([string]::IsNullOrWhiteSpace($TaskId)) { return }
    $script:CompletedCloneTaskOutputs[$TaskId] = @{ clone_id = $CloneId; stored_at = [DateTime]::UtcNow }
}

function Get-CompletedCloneTaskOutput {
    param([string]$TaskId)
    $e = $script:CompletedCloneTaskOutputs[$TaskId]
    if ($null -eq $e) { return $null }
    if (([DateTime]::UtcNow - $e.stored_at).TotalSeconds -ge $script:Settings.cloning.completed_clone_task_output_retention_seconds) {
        $script:CompletedCloneTaskOutputs.Remove($TaskId)
        return $null
    }
    return [string]$e.clone_id
}

# Everything per-id, for an instance that is gone (or whose tracking must end). The
# basic-template and maintenance removals are what keep a later instance that reuses
# the id from inheriting is_template.
function Clear-VmeTrackingForVm {
    param([Parameter(Mandatory = $true)][string]$VmId, [switch]$KeepRecord, [switch]$Gone)

    $key = Get-CloneKeyByCloneId -VmId $VmId
    if ($null -ne $key) { Remove-CloneEntry -Key $key }
    if ($Gone) {
        # CompletedCloneTaskOutputs is a SEPARATE cache from CloneState, keyed by task id
        # and kept for completed_clone_task_output_retention_seconds (default 600s) so a
        # repeated tasks/get after completion doesn't need the CloneState entry to still
        # exist (Handle-TaskInfo checks it first, before even looking at CloneState).
        # Removing only the CloneState entry above is not enough: without this, deleting
        # a guest within that window -- exactly what happens when an admin cancels a slow
        # clone from the RAS console mid-provision -- leaves tasks/get still answering
        # "completed, here is your (now deleted) guest" for whatever's left of the
        # window: VME shows the clone gone while RAS is still told it exists. Gated on -Gone (passed only by call sites
        # where the instance is confirmed deleted/gone, not e.g. a babysit timeout or a
        # provisioning failure where the guest still exists) so a perfectly healthy
        # clone's already-answered task doesn't start failing once its own babysit
        # tracking retires for unrelated reasons (first real IP, etc).
        foreach ($tid in @($script:CompletedCloneTaskOutputs.Keys)) {
            if ([string]$script:CompletedCloneTaskOutputs[$tid].clone_id -eq $VmId) { $script:CompletedCloneTaskOutputs.Remove($tid) }
        }
    }
    # A given-up clone confirmed gone (404, or dropped by the listing) frees its clone
    # slot. Not on a delete that was only submitted (-Gone -KeepRecord): under the load
    # that left the clone stuck, the delete can queue too -- the probe's 404 frees it.
    if ($Gone -and -not $KeepRecord) {
        if ($script:FormerTemplates.ContainsKey($VmId)) { $script:FormerTemplates.Remove($VmId); Save-FormerTemplateState }
        if ($script:TemplateVersions.ContainsKey($VmId)) {
            Write-DebugLog "Template instance [$VmId] is gone from VME -- its version record(s) $(@($script:TemplateVersions[$VmId].versions.Keys) -join ', ') dropped from the provider state." -Level 'W' -Component '07' -Ref $VmId
            $script:TemplateVersions.Remove($VmId)
            Save-TemplateVersions
        }
        foreach ($tid in @($script:StuckClones.Keys)) {
            if ([string]$script:StuckClones[$tid].clone_id -eq $VmId) { Remove-VmeStuckClone -Key $tid -Why 'VME confirms it gone' }
        }
    }
    $script:CloneLastRefresh.Remove($VmId)
    $script:PowerWatch.Remove($VmId)
    $script:WatchReadAt.Remove($VmId)
    $script:StatusLogged.Remove($VmId)
    [void]$script:ExcludeLogged.Remove($VmId)
    [void]$script:OrphanReported.Remove($VmId)
    [void]$script:TemplateFlappingUnmaskedVmIds.Remove($VmId)
    if ($script:MaintenanceIds.ContainsKey($VmId)) { $script:MaintenanceIds.Remove($VmId); Save-MaintenanceState }
    if ($script:BasicTemplates.ContainsKey($VmId)) { $script:BasicTemplates.Remove($VmId); Save-BasicTemplateState }
    if (-not $KeepRecord) {
        # The poll stamp goes only with the record: a clone whose tracking retires at its
        # first IP is still a live guest RAS keeps polling, and dropping its stamp made
        # the orphan audit fall back to the process start time and flag it as "never
        # polled" (a false ORPHAN CANDIDATE).
        $script:LastGuestPollAt.Remove($VmId)
        $script:Inventory.Remove($VmId)
        $script:HwById.Remove($VmId)
        $script:HwLastTry.Remove($VmId)
        $script:CloneBodyBySource.Remove($VmId)
    }
}

# ---------------------------------------------------------------------------
# VME transport
# ---------------------------------------------------------------------------

function Get-VmeSession {
    if ($null -eq $script:VmeSession) { throw "$($script:ProviderNamePrefix) Not connected. Call provider/connect first." }
    return $script:VmeSession
}

function New-VmeApiException {
    param([string]$Message, [int]$StatusCode = 0, [bool]$NoResponse = $false)
    $ex = [System.Exception]::new($Message)
    $ex.Data['VmeStatusCode'] = $StatusCode
    $ex.Data['VmeNoResponse'] = $NoResponse
    return $ex
}

function Get-VmeErrorStatus { param($ErrorRecord) try { return [int]$ErrorRecord.Exception.Data['VmeStatusCode'] } catch { return 0 } }
function Test-VmeNoResponse { param($ErrorRecord) try { return [bool]$ErrorRecord.Exception.Data['VmeNoResponse'] } catch { return $false } }

# Morpheus errors are {"success":false,"msg":"...","errors":{"field":"..."}}.
function Get-VmeErrorText {
    param([string]$Content, [int]$StatusCode)
    if ([string]::IsNullOrWhiteSpace($Content)) { return "HTTP $StatusCode" }
    try {
        $o = ConvertFrom-JsonFast -Json $Content
        $parts = [System.Collections.Generic.List[string]]::new()
        foreach ($k in @('msg', 'message', 'error')) {
            $v = Get-JsonValue $o $k
            if ($v -is [string] -and -not [string]::IsNullOrWhiteSpace($v)) { $parts.Add($v) }
        }
        $errs = Get-JsonValue $o 'errors'
        if ($errs -is [System.Collections.IDictionary]) { foreach ($k in $errs.Keys) { $parts.Add("${k}: $($errs[$k])") } }
        if ($parts.Count -gt 0) { return "HTTP ${StatusCode}: $($parts -join '; ')" }
    }
    catch { }
    $text = $Content.Trim()
    if ($text.Length -gt 300) { $text = $text.Substring(0, 300) + '...' }
    return "HTTP ${StatusCode}: $text"
}

function ConvertTo-VmeLoggableText {
    param([string]$Text, [int]$MaxLength = 2000)
    if ($null -eq $Text) { return '<empty>' }
    if ($Text.Length -gt $MaxLength) { return $Text.Substring(0, $MaxLength) + "...<truncated, $($Text.Length) chars>" }
    return $Text
}

# One HTTP exchange with the retry policy applied. Returns
#   @{ Ok; StatusCode; Data; Error; NoResponse }
# and never throws for an HTTP or network failure -- the caller decides what a 404 or
# a missing response means.
#
# Retry policy is by idempotency, not by verb:
#   - GET: up to retry_max_attempts on no response, 429 or 5xx;
#   - -Idempotent writes (power actions, DELETE): at most one retry, same conditions;
#   - everything else, the clone PUT above all: never retried -- a lost response may
#     still mean VME accepted it, and a second clone would be a duplicate VM.
# After a no-response failure the pooled connection is dropped, so a retry (and every
# later call) starts on a fresh TCP/TLS connection.
function Invoke-VmeHttp {
    param(
        [Parameter(Mandatory = $true)][string]$Uri,
        [string]$Method = 'GET',
        [hashtable]$Headers = @{},
        [object]$Body = $null,
        [switch]$Form,
        [switch]$Idempotent,
        [string]$Ref = $null,
        [string]$StatKey = $null
    )

    $C = $script:Settings.connection
    $maxAttempts = if ($Method -eq 'GET') { [Math]::Max(1, [int]$C.retry_max_attempts) } elseif ($Idempotent) { [Math]::Min(2, [Math]::Max(1, [int]$C.retry_max_attempts)) } else { 1 }

    $pathOnly = ($Uri -replace '^https?://[^/]+', '')
    if ([string]::IsNullOrWhiteSpace($StatKey)) { $StatKey = "$Method " + (($pathOnly -replace '\?.*$', '') -replace '/\d+', '/{id}') }

    $bodyText = $null
    if ($null -ne $Body -and -not $Form) { $bodyText = ConvertTo-Json -InputObject $Body -Depth 12 -Compress }

    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        if ($null -eq $script:WebSession) { $script:WebSession = [Microsoft.PowerShell.Commands.WebRequestSession]::new() }
        $p = @{
            Uri                = $Uri
            Method             = $Method
            Headers            = $Headers
            WebSession         = $script:WebSession
            SkipHttpErrorCheck = $true
            ErrorAction        = 'Stop'
        }
        if (-not $C.use_system_proxy) { $p.NoProxy = $true }
        if (-not $script:SslCheck) { $p.SkipCertificateCheck = $true }
        # -ConnectionTimeoutSeconds (-TimeoutSec before 7.4) bounds the wait for VME to
        # answer; on 7.4+ it belongs to the session's pooled HttpClient and must be the
        # same on every call, or PowerShell rebuilds the client and the connection is
        # lost. -OperationTimeoutSeconds only bounds silence while the body streams in.
        if ($script:HasConnectionTimeout) { $p.ConnectionTimeoutSeconds = [int]$C.request_timeout_seconds }
        else { $p.TimeoutSec = [int]$C.request_timeout_seconds }
        if ($script:HasOperationTimeout) { $p.OperationTimeoutSeconds = [int]$C.read_timeout_seconds }
        if ($Form) { $p.Body = $Body; $p.ContentType = 'application/x-www-form-urlencoded' }
        elseif ($null -ne $bodyText) { $p.Body = $bodyText; $p.ContentType = 'application/json' }

        if ($script:LogLevel -ge 5) {
            $bodyNote = if ($null -ne $bodyText -and $attempt -eq 1) { " body=$(ConvertTo-VmeLoggableText $bodyText)" } else { '' }
            Write-DebugLog "HTTP $Method $pathOnly$(if ($attempt -gt 1) { " (attempt $attempt/$maxAttempts)" })$bodyNote" -Level 'D' -Component '0B' -Ref $Ref
        }

        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $r = $null
        try { $r = Invoke-WebRequest @p }
        catch {
            $sw.Stop()
            Add-HttpStat -Key $StatKey -Ms $sw.Elapsed.TotalMilliseconds -Bytes 0 -IsError $true
            $msg = $_.Exception.Message
            if ($null -ne $_.Exception.InnerException -and $_.Exception.InnerException.Message -ne $msg) { $msg = "$msg ($($_.Exception.InnerException.Message))" }
            $script:WebSession = $null
            if ($attempt -lt $maxAttempts) {
                Write-DebugLog "HTTP $Method $pathOnly got no response (attempt $attempt/$maxAttempts): $msg -- retrying on a fresh connection." -Level 'W' -Component '0B' -Ref $Ref
                Start-Sleep -Milliseconds ([int]$C.retry_delay_milliseconds * $attempt)
                continue
            }
            return @{ Ok = $false; StatusCode = 0; Data = $null; Error = "no response: $msg"; NoResponse = $true }
        }
        $sw.Stop()

        $status = [int]$r.StatusCode
        $content = $r.Content
        if ($content -is [byte[]]) { $content = $script:Utf8NoBom.GetString($content) }
        $bytes = [long]$(if ($r.RawContentLength -ge 0) { $r.RawContentLength } elseif ($null -ne $content) { $content.Length } else { 0 })
        Add-HttpStat -Key $StatKey -Ms $sw.Elapsed.TotalMilliseconds -Bytes $bytes -IsError ($status -ge 400)

        if ($status -ge 200 -and $status -lt 300) {
            $data = $null
            if (-not [string]::IsNullOrWhiteSpace($content)) {
                try { $data = ConvertFrom-JsonFast -Json $content }
                catch {
                    try { $data = ConvertFrom-Json -InputObject $content -ErrorAction Stop }
                    catch { return @{ Ok = $false; StatusCode = $status; Data = $null; Error = "unparseable response body: $($_.Exception.Message)"; NoResponse = $false } }
                }
            }
            if ($Method -ne 'GET' -and $script:LogLevel -ge 5) {
                Write-DebugLog "HTTP $Method $pathOnly -> $status $(ConvertTo-VmeLoggableText $content)" -Level 'D' -Component '0B' -Ref $Ref
            }
            return @{ Ok = $true; StatusCode = $status; Data = $data; Error = $null; NoResponse = $false }
        }

        $errText = Get-VmeErrorText -Content $content -StatusCode $status
        $transient = ($status -eq 429 -or $status -ge 500)
        if ($transient -and $attempt -lt $maxAttempts) {
            $delayMs = [int]$C.retry_delay_milliseconds * $attempt
            if ($status -eq 429) {
                $ra = $null
                try { $ra = @($r.Headers['Retry-After'])[0] } catch { }
                [int]$raSec = 0
                if ($null -ne $ra -and [int]::TryParse([string]$ra, [ref]$raSec)) { $delayMs = [Math]::Min($raSec, [int]$C.retry_after_max_seconds) * 1000 }
            }
            Write-DebugLog "HTTP $Method $pathOnly -> $errText (attempt $attempt/$maxAttempts) -- retrying in ${delayMs}ms." -Level 'W' -Component '0B' -Ref $Ref
            Start-Sleep -Milliseconds $delayMs
            continue
        }
        return @{ Ok = $false; StatusCode = $status; Data = $null; Error = $errText; NoResponse = $false }
    }
}

function Invoke-VmeAuth {
    param([Parameter(Mandatory = $true)][string]$BaseUrl, [Parameter(Mandatory = $true)][string]$Username, [Parameter(Mandatory = $true)][string]$Password)

    # The OAuth parameters are documented as form fields; appliances also accept
    # grant_type/scope/client_id in the query string, so both are sent.
    $clientId = [string]$script:Settings.connection.oauth_client_id
    $query = "grant_type=password&scope=write&client_id=$([uri]::EscapeDataString($clientId))"
    $form = @{ grant_type = 'password'; scope = 'write'; client_id = $clientId; username = $Username; password = $Password }
    $r = Invoke-VmeHttp -Uri "$BaseUrl/oauth/token?$query" -Method POST -Body $form -Form -Idempotent -StatKey 'POST /oauth/token'
    return (ConvertFrom-VmeTokenResponse -Result $r -Grant 'password' -FallbackRefreshToken $null)
}

# grant_type=refresh_token works without a stored password, which is what keeps a
# pre-issued access_token (RAS custom setting) alive past its expiry.
function Invoke-VmeTokenRefresh {
    param([Parameter(Mandatory = $true)][string]$BaseUrl, [Parameter(Mandatory = $true)][string]$RefreshToken)

    $form = @{ grant_type = 'refresh_token'; client_id = [string]$script:Settings.connection.oauth_client_id; refresh_token = $RefreshToken }
    $r = Invoke-VmeHttp -Uri "$BaseUrl/oauth/token" -Method POST -Body $form -Form -Idempotent -StatKey 'POST /oauth/token'
    # A server that rotates refresh tokens returns a new one; one that doesn't omits it.
    return (ConvertFrom-VmeTokenResponse -Result $r -Grant 'refresh_token' -FallbackRefreshToken $RefreshToken)
}

function ConvertFrom-VmeTokenResponse {
    param([hashtable]$Result, [string]$Grant, [string]$FallbackRefreshToken)

    if (-not $Result.Ok) { throw (New-VmeApiException -Message "$($script:ProviderNamePrefix) OAuth $Grant grant failed: $($Result.Error)" -StatusCode $Result.StatusCode -NoResponse $Result.NoResponse) }
    $token = [string](Get-JsonValue $Result.Data 'access_token')
    if ([string]::IsNullOrWhiteSpace($token)) { throw "$($script:ProviderNamePrefix) OAuth $Grant grant returned no access_token" }
    $expiresIn = 0
    try { $expiresIn = [int](Get-JsonValue $Result.Data 'expires_in') } catch { }
    $refresh = [string](Get-JsonValue $Result.Data 'refresh_token')
    if ([string]::IsNullOrWhiteSpace($refresh)) { $refresh = $FallbackRefreshToken }
    return @{
        token         = $token
        expires_at    = if ($expiresIn -gt 0) { [DateTime]::UtcNow.AddSeconds($expiresIn) } else { $null }
        refresh_token = $refresh
    }
}

function Set-VmeSessionToken {
    param([Parameter(Mandatory = $true)][hashtable]$Session, [Parameter(Mandatory = $true)][hashtable]$Auth)
    $Session.token = $Auth.token
    $Session.expires_at = $Auth.expires_at
    # Sent per request, never stored in the WebSession (a session-level header would
    # outlive a token refresh).
    $Session.header = @{ Authorization = "BEARER $($Auth.token)" }
    if (-not [string]::IsNullOrWhiteSpace([string]$Auth.refresh_token)) { $Session.refresh_token = $Auth.refresh_token }
}

function Test-VmeCanReauthenticate {
    param([hashtable]$Session)
    return (-not [string]::IsNullOrWhiteSpace([string]$Session.refresh_token) -or -not [string]::IsNullOrWhiteSpace([string]$Session.username))
}

function Invoke-VmeReauthenticate {
    param([hashtable]$Session)
    if (-not [string]::IsNullOrWhiteSpace([string]$Session.refresh_token)) {
        try {
            Set-VmeSessionToken -Session $Session -Auth (Invoke-VmeTokenRefresh -BaseUrl $Session.baseUrl -RefreshToken $Session.refresh_token)
            return 'refresh_token'
        }
        catch {
            if ([string]::IsNullOrWhiteSpace([string]$Session.username)) { throw }
            Write-DebugLog "Token refresh failed ($($_.Exception.Message)) -- falling back to username/password." -Level 'W' -Component '02'
        }
    }
    Set-VmeSessionToken -Session $Session -Auth (Invoke-VmeAuth -BaseUrl $Session.baseUrl -Username $Session.username -Password $Session.password)
    return 'password'
}

# A pre-issued access_token has no known expiry. GET /api/user-settings lists this
# user's tokens with their expiration (maskedAccessToken shows a prefix, matched here),
# so it can be renewed ahead of time instead of only reactively on a 401.
function Invoke-VmeTokenExpiryDiscovery {
    param([hashtable]$Session)
    $R = $script:Settings.connection.token_renewal
    if (-not $R.proactive_enabled) { return }
    if (([DateTime]::UtcNow - $Session.expiry_last_checked_at).TotalHours -lt $R.expiry_check_interval_hours) { return }
    $Session.expiry_last_checked_at = [DateTime]::UtcNow

    $resp = Invoke-VmeHttp -Uri "$($Session.baseUrl)/api/user-settings" -Headers $Session.header -StatKey 'GET /api/user-settings'
    if (-not $resp.Ok) { Write-DebugLog "Token expiry discovery: GET /api/user-settings failed ($($resp.Error)) -- relying on the reactive 401 renewal." -Level 'W' -Component '02'; return }
    $match = $null
    foreach ($t in @(Get-JsonValue $resp.Data 'accessTokens')) {
        if ([string](Get-JsonValue $t 'clientId') -ne [string]$script:Settings.connection.oauth_client_id) { continue }
        $prefix = ([string](Get-JsonValue $t 'maskedAccessToken')).TrimEnd('*')
        if (-not [string]::IsNullOrWhiteSpace($prefix) -and ([string]$Session.token).StartsWith($prefix)) { $match = $t; break }
    }
    if ($null -eq $match) { Write-DebugLog 'Token expiry discovery: no accessTokens[] entry matched this session token -- relying on the reactive 401 renewal.' -Level 'W' -Component '02'; return }
    $epoch = ConvertTo-EpochSeconds (Get-JsonValue $match 'expiration')
    if ($null -eq $epoch) { return }
    $discovered = [DateTimeOffset]::FromUnixTimeSeconds($epoch).UtcDateTime
    if ($null -eq $Session.expires_at -or $discovered -gt $Session.expires_at) {
        $Session.renewal_attempts = 0
        $Session.last_renewal_attempt_at = $null
    }
    $Session.expires_at = $discovered
    Write-DebugLog "Token expiry discovery: access token expires [$(ConvertTo-IsoFromEpoch $epoch)]." -Level 'I' -Component '02'
}

function Invoke-VmeProactiveTokenRenewalIfDue {
    param([hashtable]$Session)
    $R = $script:Settings.connection.token_renewal
    if (-not $R.proactive_enabled -or $null -eq $Session.expires_at) { return }
    if (-not [string]::IsNullOrWhiteSpace([string]$Session.username)) { return }   # the margin refresh below covers password sessions
    if ([string]::IsNullOrWhiteSpace([string]$Session.refresh_token)) { return }
    if (($Session.expires_at - [DateTime]::UtcNow).TotalDays -gt $R.lead_days) { return }
    if ([int]$Session.renewal_attempts -ge $R.max_attempts) { return }
    if ($null -ne $Session.last_renewal_attempt_at -and ([DateTime]::UtcNow - $Session.last_renewal_attempt_at).TotalHours -lt $R.retry_interval_hours) { return }

    $Session.renewal_attempts = [int]$Session.renewal_attempts + 1
    $Session.last_renewal_attempt_at = [DateTime]::UtcNow
    try {
        Set-VmeSessionToken -Session $Session -Auth (Invoke-VmeTokenRefresh -BaseUrl $Session.baseUrl -RefreshToken $Session.refresh_token)
        $Session.renewal_attempts = 0
        Write-DebugLog "Proactive token renewal succeeded; new expiry [$(if ($Session.expires_at) { $Session.expires_at.ToString('o') } else { 'unknown' })]." -Level 'I' -Component '02'
    }
    catch {
        Write-DebugLog "Proactive token renewal attempt $($Session.renewal_attempts)/$($R.max_attempts) failed: $($_.Exception.Message)" -Level 'W' -Component '02'
    }
}

function Update-VmeTokenIfExpiring {
    param([hashtable]$Session)
    if ($null -eq $Session.expires_at) { return }
    if (($Session.expires_at - [DateTime]::UtcNow).TotalSeconds -gt $script:Settings.connection.token_refresh_margin_seconds) { return }
    if (-not (Test-VmeCanReauthenticate $Session)) {
        Write-DebugLog 'Access token is close to expiry and neither a refresh_token nor username/password is available -- RAS must reconnect.' -Level 'W' -Component '02'
        return
    }
    try {
        $how = Invoke-VmeReauthenticate -Session $Session
        Write-DebugLog "Access token refreshed ahead of expiry via $how." -Level 'I' -Component '02'
    }
    catch { Write-DebugLog "Token refresh ahead of expiry failed: $($_.Exception.Message)" -Level 'W' -Component '02' }
}

# The three token checks behind one timestamp: the common path is one comparison.
function Update-VmeTokenIfDue {
    param([hashtable]$Session)
    $now = [DateTime]::UtcNow
    if ($now -lt $Session.next_token_check) { return }
    $Session.next_token_check = $now.AddSeconds(60)
    Invoke-VmeTokenExpiryDiscovery -Session $Session
    Invoke-VmeProactiveTokenRenewalIfDue -Session $Session
    Update-VmeTokenIfExpiring -Session $Session
}

# The one entry point for API calls. Returns the parsed body; $null for a 404 when
# -AllowNotFound; otherwise throws an exception carrying VmeStatusCode/VmeNoResponse.
function Invoke-VmeApi {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [string]$Method = 'GET',
        [object]$Body = $null,
        [switch]$AllowNotFound,
        [switch]$Idempotent,
        [ValidateSet('E', 'W', 'I', 'T', 'D')][string]$FailureLevel = 'E',
        [string]$Ref = $null,
        [string]$StatKey = $null
    )

    $session = Get-VmeSession
    Update-VmeTokenIfDue -Session $session
    $uri = "$($session.baseUrl)$Path"
    $call = { Invoke-VmeHttp -Uri $uri -Method $Method -Headers $session.header -Body $Body -Idempotent:$Idempotent -Ref $Ref -StatKey $StatKey }

    $r = & $call
    if (-not $r.Ok -and $r.StatusCode -eq 401 -and (Test-VmeCanReauthenticate $session)) {
        Write-DebugLog "HTTP 401 on $Method $Path -- re-authenticating once and retrying." -Level 'W' -Component '02' -Ref $Ref
        try { [void](Invoke-VmeReauthenticate -Session $session) }
        catch { throw (New-VmeApiException -Message "$($script:ProviderNamePrefix) VME returned 401 and re-authentication failed: $($_.Exception.Message)" -StatusCode 401) }
        $r = & $call
        if (-not $r.Ok -and $r.StatusCode -eq 401) {
            Write-DebugLog "HTTP 401 on $Method $Path again right after re-authenticating -- the credentials are rejected or the user lacks permission." -Level 'E' -Component '02' -Ref $Ref
            throw (New-VmeApiException -Message "$($script:ProviderNamePrefix) VME rejected the credentials (HTTP 401) even after re-authentication" -StatusCode 401)
        }
    }

    if ($r.Ok) { return $r.Data }
    if ($r.StatusCode -eq 404 -and $AllowNotFound) { return $null }
    Write-DebugLog "HTTP failure ($Method $Path): $($r.Error)" -Level $FailureLevel -Component '0B' -Ref $Ref
    throw (New-VmeApiException -Message "$($script:ProviderNamePrefix) $Method $Path failed: $($r.Error)" -StatusCode $r.StatusCode -NoResponse $r.NoResponse)
}

# ---------------------------------------------------------------------------
# Status mapping
# ---------------------------------------------------------------------------

# CPF knows six states. Every VME status maps to one of them -- never to anything
# else ('unknown' is not a CPF state). Statuses from HPE's Terraform provider for
# MVM/HVM plus the
# delete-time ones: 'removing'/'pendingRemoval' must read as powering_off, because RAS
# keeps polling an instance until guests/list stops returning it.
function Map-VmeStateToRasState {
    param([string]$Status)
    switch ($Status.ToLowerInvariant()) {
        'running' { return 'powered_on' }
        'warning' { return 'powered_on' }            # running, flagged by monitoring
        'stopped' { return 'powered_off' }
        'suspended' { return 'suspended' }
        'suspending' { return 'suspending' }
        'starting' { return 'powering_on' }
        'provisioning' { return 'powering_on' }
        'pending' { return 'powering_on' }
        'deploying' { return 'powering_on' }
        'restarting' { return 'powering_on' }
        'resizing' { return 'powering_on' }
        'stopping' { return 'powering_off' }
        'removing' { return 'powering_off' }
        'pendingremoval' { return 'powering_off' }
        'deleting' { return 'powering_off' }
        default { return 'powered_off' }             # failed/denied/cancelled/unknown: logged once per change
    }
}

$script:KnownVmeStatuses = [System.Collections.Generic.HashSet[string]]::new([string[]]@('running', 'warning', 'stopped', 'suspended', 'suspending', 'starting', 'provisioning', 'pending', 'deploying', 'restarting', 'resizing', 'stopping', 'removing', 'pendingremoval', 'deleting'), [System.StringComparer]::OrdinalIgnoreCase)

# Is VME's own provisioning of a fresh clone finished? (Terraform provider: Pending =
# provisioning/starting/stopping/pending; Target = running/failed/warning/denied/
# cancelled/suspended.) 'stopped' is in neither set: a clone provisioned powered off
# (provisionPoweredOff) ends there, so 'stopped' with isBusy=false counts as finished
# (treating it as pending would keep the babysit start from ever firing). An unknown status counts as pending; the time budget is the backstop.
function Get-VmeProvisionStatusClass {
    param([string]$Status, [bool]$IsBusy = $false)
    switch ($Status.ToLowerInvariant()) {
        { $_ -in @('running', 'warning', 'suspended') } { return 'ready' }
        'stopped' { if ($IsBusy) { return 'pending' } return 'ready' }
        { $_ -in @('failed', 'denied', 'cancelled') } { return 'failed' }
        default { return 'pending' }
    }
}

# Is VME in the middle of building this instance (any instance, not only a clone)? A
# power action then is refused by VME or undone by it. Deliberately NOT the
# provision class above: a stopped template is not "provisioning".
function Test-VmeInstanceProvisioning {
    param([hashtable]$Rec)
    return ($Rec.is_busy -or ([string]$Rec.status).ToLowerInvariant() -in @('provisioning', 'pending', 'deploying'))
}

# ---------------------------------------------------------------------------
# Instance records
# ---------------------------------------------------------------------------

function Test-VmeUsableIpv4 {
    param([string]$Value, [switch]$AllowLinkLocal)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    $v = $Value.Trim()
    if ($v -notmatch '^(?:\d{1,3}\.){3}\d{1,3}$') { return $false }
    if ($v -eq '0.0.0.0' -or $v.StartsWith('127.') -or $v.StartsWith('255.')) { return $false }
    if (-not $AllowLinkLocal -and $v.StartsWith('169.254.')) { return $false }
    return $true
}

# A clone counts as "up" only with a routable address; link-local alone never does.
function Test-VmeHasUsableIpv4Address {
    param([string[]]$Addresses)
    foreach ($a in @($Addresses)) { if (Test-VmeUsableIpv4 -Value $a) { return $true } }
    return $false
}

# Clone body inputs, taken from a lean instance payload (every field clone needs is
# in it). Only disk volumes are echoed.
function Get-VmeCloneBodyInputs {
    param([object]$Inst)
    $vols = [System.Collections.Generic.List[object]]::new()
    foreach ($v in @(Get-JsonValue $Inst 'volumes')) {
        if ($null -eq $v) { continue }
        $cat = [string](Get-JsonValue $v 'volumeCategory')
        if ($cat -and $cat -ne 'disk') { continue }
        $vol = [ordered]@{ size = [int](Get-JsonValue $v 'size'); rootVolume = [bool](Get-JsonValue $v 'rootVolume') }
        $n = [string](Get-JsonValue $v 'name')
        if ($n) { $vol.name = $n }
        $ds = Get-JsonValue $v 'datastoreId'
        if ($null -ne $ds) { $vol.datastoreId = $ds }
        $vols.Add($vol)
    }
    $firstIf = @(Get-JsonValue $Inst 'interfaces') | Select-Object -First 1
    # For a linked clone (POST /api/instances builds a new VM, nothing is inherited):
    # the instance type code, the template's real size and its config flags --
    # firmware above all, a Windows 11 template without UEFI/TPM/Secure Boot won't boot.
    $cfg = Get-JsonValue $Inst 'config'
    $vmConfig = [ordered]@{}
    foreach ($k in @('resourcePoolId', 'poolProviderType', 'uefi', 'secureBoot', 'tpm', 'attachVirtIODrivers', 'nestedVirtualization', 'noAgent', 'kvmHostId')) {
        $v = Get-JsonValue $cfg $k
        if ($null -ne $v -and "$v" -ne '') { $vmConfig[$k] = $v }
    }
    return @{
        group         = [string](Get-JsonValue (Get-JsonValue $Inst 'group') 'id')
        cloud         = [string](Get-JsonValue (Get-JsonValue $Inst 'cloud') 'id')
        plan          = [string](Get-JsonValue (Get-JsonValue $Inst 'plan') 'id')
        layout        = [string](Get-JsonValue (Get-JsonValue $Inst 'layout') 'id')
        network       = [string](Get-JsonValue (Get-JsonValue $firstIf 'network') 'id')
        network_name  = [string](Get-JsonValue (Get-JsonValue $firstIf 'network') 'name')
        volumes       = $vols.ToArray()
        instance_type = [string](Get-JsonValue (Get-JsonValue $Inst 'instanceType') 'code')
        max_cores     = Get-JsonValue $Inst 'maxCores'
        cores_per_socket = Get-JsonValue $Inst 'coresPerSocket'
        max_memory    = Get-JsonValue $Inst 'maxMemory'
        vm_config     = $vmConfig
        learned_at    = [DateTime]::UtcNow
    }
}

# The ONE place untrusted instance JSON is read. Everything downstream works on the
# slim record; the raw object is dropped by the caller right after this.
function ConvertTo-VmeInstanceRecord {
    param([object]$Inst)

    $id = [string](Get-JsonValue $Inst 'id')
    if ([string]::IsNullOrWhiteSpace($id)) { return $null }
    $status = [string](Get-JsonValue $Inst 'status')
    $maxAddr = [int]$script:Settings.inventory.max_addresses

    # IPv4 from connectionInfo[].ip (present in lean payloads). evars MVM_IP is not
    # used: it was seen stale at 0.0.0.0.
    $real = [System.Collections.Generic.List[string]]::new()
    $linkLocal = [System.Collections.Generic.List[string]]::new()
    foreach ($ci in @(Get-JsonValue $Inst 'connectionInfo')) {
        $ip = ([string](Get-JsonValue $ci 'ip')).Trim()
        if (Test-VmeUsableIpv4 -Value $ip) { if (-not $real.Contains($ip)) { $real.Add($ip) } }
        elseif ((Test-VmeUsableIpv4 -Value $ip -AllowLinkLocal) -and -not $linkLocal.Contains($ip)) { $linkLocal.Add($ip) }
    }
    $real.AddRange($linkLocal)
    if ($real.Count -gt $maxAddr) { $real.RemoveRange($maxAddr, $real.Count - $maxAddr) }

    $servers = @(Get-JsonValue $Inst 'servers')
    $serverId = if ($servers.Count -gt 0 -and $null -ne $servers[0]) { [string]$servers[0] } else { '' }
    if ($servers.Count -gt 1) { Write-DebugLog "Instance [$id] has $($servers.Count) servers; MAC/OS are taken from the first, [$serverId]." -Level 'T' -Component '03' -Ref $id }

    $clonedFrom = [string](Get-JsonValue (Get-JsonValue (Get-JsonValue $Inst 'config') 'provisionOpts') 'clonedInstanceId')
    # The virtual image the VM was built from: for a linked clone, its template version's
    # image (lineage under 'versioning'); for anything else, an unrelated library image.
    $imageId = [string](Get-JsonValue (Get-JsonValue $Inst 'config') 'imageId')

    $rec = @{
        id             = $id
        uuid           = [string](Get-JsonValue $Inst 'uuid')
        name           = [string](Get-JsonValue $Inst 'name')
        status         = $status
        ras_state      = (Map-VmeStateToRasState -Status $status)
        ipv4           = $null
        server_id      = $serverId
        cloned_from    = $clonedFrom
        image_id       = $imageId
        is_busy        = [bool](Get-JsonValue $Inst 'isBusy')
        cloud          = [string](Get-JsonValue (Get-JsonValue $Inst 'cloud') 'name')
        status_percent = Get-JsonValue $Inst 'statusPercent'
        status_message = [string](Get-JsonValue $Inst 'statusMessage')
        date_created   = [string](Get-JsonValue $Inst 'dateCreated')
        fetched_at     = [DateTime]::UtcNow
        dirty          = $false
        missing        = 0
    }
    # Assigned outside the literal: an empty or one-element array must stay an array.
    $rec.ipv4 = $real.ToArray()
    if ([string]::IsNullOrWhiteSpace($rec.cloud)) { $rec.cloud = 'vme' }

    # Abnormal statuses are logged once per change, not on every poll of every sweep.
    $isAbnormal = -not $script:KnownVmeStatuses.Contains($status)
    if ($isAbnormal) {
        if ($script:StatusLogged[$id] -ne $status) {
            $script:StatusLogged[$id] = $status
            $lvl = if ($status -in @('failed', 'denied', 'cancelled')) { 'E' } else { 'W' }
            Write-DebugLog "Instance [$id] ($($rec.name)) has VME status [$status] -- reported to RAS as powered_off.$(if ($rec.status_message) { " VME says: $($rec.status_message)" })" -Level $lvl -Component '03' -Ref $id
        }
    }
    elseif ($script:StatusLogged.ContainsKey($id)) { $script:StatusLogged.Remove($id) }

    # Clone sources: refresh their clone-body inputs for free from every payload.
    if ($script:BasicTemplates.ContainsKey($id) -or $script:CloneBodyBySource.ContainsKey($id)) {
        $script:CloneBodyBySource[$id] = Get-VmeCloneBodyInputs -Inst $Inst
    }
    return $rec
}

# One lean GET (details=false: ~6 KB instead of ~16.5 KB, and everything the provider
# needs). Upserts the record. $null when VME says 404.
function Get-VmeInstanceRecord {
    param([Parameter(Mandatory = $true)][string]$VmId, [string]$Ref = $null, [ValidateSet('E', 'W', 'T')][string]$FailureLevel = 'E')

    $resp = Invoke-VmeApi -Path "/api/instances/${VmId}?details=false" -AllowNotFound -Ref $(if ($Ref) { $Ref } else { $VmId }) -FailureLevel $FailureLevel
    if ($null -eq $resp) { return $null }
    $inst = Get-JsonValue $resp 'instance'
    if ($null -eq $inst) { throw "$($script:ProviderNamePrefix) GET /api/instances/$VmId returned no 'instance' object" }
    $rec = ConvertTo-VmeInstanceRecord -Inst $inst
    if ($null -eq $rec) { throw "$($script:ProviderNamePrefix) GET /api/instances/$VmId returned an instance without an id" }
    $script:Inventory[$VmId] = $rec
    return $rec
}

function Test-VmeInstanceExcluded {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return $false }
    foreach ($re in $script:ExcludeRegexes) { if ($re.IsMatch($Name)) { return $true } }
    return $false
}

# ---------------------------------------------------------------------------
# Inventory refresh
#
# A full listing, projected page by page, applied with these guards (review 6.3).
# Cold start (no inventory yet) asks for the whole fleet in one request
# (inventory.cold_start_page_size, default max=-1); every later refresh walks
# page_size pages. A capped or refused cold-start request continues with pages.
# Guards:
#   - all pages or nothing: a failed page keeps the previous inventory (served stale);
#   - duplicates from offset paging during churn are dropped;
#   - an instance missing from one listing is kept until it is missing from
#     missing_listings_before_drop consecutive complete listings (offset paging can
#     skip an id while instances are created or deleted);
#   - mass-drop guard: a listing much smaller than the inventory is applied as
#     add/update only until a second listing confirms the shrink (an empty answer
#     from a permission or tenant glitch must not make the fleet vanish from RAS);
#   - ids being deleted or tracked as in-flight clones are never dropped here.
# ---------------------------------------------------------------------------

function Update-VmeInventory {
    param([int]$MaxAgeSeconds = 0, [string]$Reason = 'guests/list')

    if ($script:InventoryLoaded -and ([DateTime]::UtcNow - $script:InventoryListedAt).TotalSeconds -lt $MaxAgeSeconds) { return $true }

    $I = $script:Settings.inventory
    $pageSize = [int]$I.page_size
    $coldStart = -not $script:InventoryLoaded
    $coldMax = [int]$I.cold_start_page_size
    # Rows asked for by the next request: -1 = all (cold start only).
    $request = if ($coldStart -and $coldMax -ne 0) { $coldMax } else { $pageSize }
    $mode = if ($request -ne $pageSize) { "cold start, max=$request" } else { $null }
    $new = @{}
    $offset = 0
    $pages = 0
    $dups = 0
    $complete = $true
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $script:InventoryAttemptAt = [DateTime]::UtcNow

    try {
        while ($true) {
            $statKey = if ($request -ne $pageSize) { 'GET /api/instances (cold start)' } else { 'GET /api/instances (page)' }
            try { $resp = Invoke-VmeApi -Path "/api/instances?max=$request&offset=$offset$($script:InstanceScopeQuery)" -StatKey $statKey -FailureLevel 'W' }
            catch {
                # A refused or timed-out whole-fleet request: the same listing goes on with
                # routine pages. A failed routine page fails the listing (all or nothing).
                if ($request -eq $pageSize) { throw }
                Write-DebugLog "Cold-start listing (max=$request) failed: $($_.Exception.Message) -- continuing with pages of $pageSize." -Level 'W' -Component '03'
                $request = $pageSize
                $mode = "cold start, max=$coldMax failed, paged"
                continue
            }
            $arr = Get-JsonValue $resp 'instances'
            $batch = @(); if ($null -ne $arr) { $batch = @($arr) }
            foreach ($inst in $batch) {
                $rec = ConvertTo-VmeInstanceRecord -Inst $inst
                if ($null -eq $rec) { continue }
                if ($new.ContainsKey($rec.id)) { $dups++ } else { $new[$rec.id] = $rec }
            }
            $pages++
            $total = $null
            try { $total = [int](Get-JsonValue (Get-JsonValue $resp 'meta') 'total') } catch { }
            $offset += $batch.Count
            if ($batch.Count -eq 0) { break }
            if ($null -ne $total -and $total -gt 0) {
                if ($offset -ge $total) { break }
                if ($request -ne $pageSize) {
                    # VME capped the cold-start request: page through the rest.
                    Write-DebugLog "Cold-start listing (max=$request) returned $offset of $total instances -- continuing with pages of $pageSize from offset $offset." -Level 'W' -Component '03'
                    $request = $pageSize
                    $mode = "cold start, max=$coldMax capped at $offset, paged"
                    continue
                }
            }
            elseif ($request -le 0 -or $batch.Count -lt $request) { break }
            if ($batch.Count -lt $request) { break }
            if ($pages -ge [int]$I.max_pages) {
                Write-DebugLog "Instance listing hit inventory.max_pages ($($I.max_pages) x $pageSize) -- applied as add/update only; raise inventory.max_pages or narrow inventory.scope." -Level 'W' -Component '03'
                $complete = $false
                break
            }
        }
    }
    catch {
        $age = if ($script:InventoryLoaded) { "$([int]([DateTime]::UtcNow - $script:InventoryListedAt).TotalSeconds)s old" } else { 'none loaded yet' }
        Write-DebugLog "Inventory refresh ($Reason) failed on page $($pages + 1): $($_.Exception.Message) -- keeping the previous inventory ($($script:Inventory.Count) instances, $age)." -Level 'W' -Component '03'
        return $false
    }
    $sw.Stop()

    $old = $script:Inventory
    $upsertOnly = -not $complete
    $floor = [Math]::Floor($old.Count * (1 - [double]$I.max_drop_ratio))
    if ($complete -and $script:InventoryLoaded -and $old.Count -ge 20 -and $new.Count -lt $floor) {
        if (-not $script:MassDropPending) {
            $script:MassDropPending = $true
            $upsertOnly = $true
            Write-DebugLog "Instance listing returned $($new.Count) instances against $($old.Count) known -- more than inventory.max_drop_ratio ($($I.max_drop_ratio)) would disappear. Not removing anything until the next listing confirms it (token scoped to the wrong tenant, or an appliance hiccup?)." -Level 'E' -Component '03'
        }
        else {
            $script:MassDropPending = $false
            Write-DebugLog "A second listing confirms the fleet shrank from $($old.Count) to $($new.Count) instances -- applying." -Level 'W' -Component '03'
        }
    }
    else { $script:MassDropPending = $false }

    $added = 0
    $removed = 0
    foreach ($id in $new.Keys) { if (-not $old.ContainsKey($id)) { $added++ } }
    $dropAfter = [int]$I.missing_listings_before_drop
    foreach ($id in @($old.Keys)) {
        if ($new.ContainsKey($id)) { continue }
        $rec = $old[$id]
        if ($script:DeletingIds.ContainsKey($id) -or $null -ne (Get-CloneKeyByCloneId -VmId $id) -or $upsertOnly) { $new[$id] = $rec; continue }
        $rec.missing = [int]$rec.missing + 1
        if ($rec.missing -ge $dropAfter) {
            $removed++
            Write-DebugLog "Instance [$id] ($($rec.name)) absent from $($rec.missing) consecutive listings -- removed from the inventory." -Level 'T' -Component '03' -Ref $id
            Clear-VmeTrackingForVm -VmId $id -Gone
        }
        else { $new[$id] = $rec }
    }

    $script:Inventory = $new
    $first = -not $script:InventoryLoaded
    $script:InventoryLoaded = $true
    $script:InventoryListedAt = [DateTime]::UtcNow
    Update-IdUuidIntegrity

    $lvl = if ($first) { 'I' } else { 'T' }
    Write-DebugLog "Inventory refreshed ($Reason$(if ($mode) { ", $mode" })): $($new.Count) instances from $pages request(s) in $([int]$sw.Elapsed.TotalMilliseconds)ms (+$added/-$removed$(if ($dups) { ", $dups duplicate(s) from paging" })$(if ($upsertOnly) { ', add/update only' }))." -Level $lvl -Component '03'

    # A routine pass must land within one RAS polling cycle, or the state RAS reads is
    # already older than the cycle when it arrives (and RAS's guests/list waited for it).
    $window = [int]$script:Settings.capabilities.guests_polling_rate
    if (-not $coldStart -and $sw.Elapsed.TotalSeconds -gt $window) {
        Write-DebugLog "Routine instance listing took $([Math]::Round($sw.Elapsed.TotalSeconds, 1))s for $pages page(s) of $pageSize -- longer than one RAS polling cycle (capabilities.guests_polling_rate = ${window}s). Raise inventory.page_size, narrow inventory.scope, or raise guests_polling_rate." -Level 'W' -Component '03'
    }
    return $true
}

# ---------------------------------------------------------------------------
# Hardware (MAC, OS) from server objects
#
# No details=false instance payload carries a MAC: it is under the server
# (server.interfaces[].macAddress), which details=true embeds at 16.5 KB per instance.
# The server objects carry credential-ish fields (sshPassword, apiKey...) and are
# never logged; only the extracted {mac, os} is kept.
# ---------------------------------------------------------------------------

function ConvertTo-VmeHardware {
    param([object]$Server)
    $macs = [System.Collections.Generic.List[string]]::new()
    $ifaces = @(Get-JsonValue $Server 'interfaces') | Where-Object { $null -ne $_ }
    $ordered = @($ifaces | Where-Object { [bool](Get-JsonValue $_ 'primaryInterface') }) + @($ifaces | Where-Object { -not [bool](Get-JsonValue $_ 'primaryInterface') })
    foreach ($if in $ordered) {
        $mac = ([string](Get-JsonValue $if 'macAddress')).Trim().ToUpperInvariant()
        if ($mac -match '^[0-9A-F]{2}(:[0-9A-F]{2}){5}$' -and -not $macs.Contains($mac)) { $macs.Add($mac) }
    }
    # First interface's network, for the clone/VLAN-preservation outcome check
    # (Test-VmeNetworkPreservationOutcome) -- same source as the MAC above.
    $netId = ''
    $netName = ''
    foreach ($if in $ordered) {
        $nid = [string](Get-JsonValue (Get-JsonValue $if 'network') 'id')
        if ($nid) { $netId = $nid; $netName = [string](Get-JsonValue (Get-JsonValue $if 'network') 'name'); break }
    }
    $os = [string](Get-JsonValue (Get-JsonValue $Server 'serverOs') 'name')
    if ([string]::IsNullOrWhiteSpace($os)) { $os = [string](Get-JsonValue $Server 'osType') }
    $max = [int]$script:Settings.inventory.max_addresses
    return @{
        server_id    = [string](Get-JsonValue $Server 'id')
        macs         = [string[]]@($macs | Select-Object -First $max)
        os           = if ([string]::IsNullOrWhiteSpace($os)) { 'unknown' } else { $os }
        network_id   = $netId
        network_name = $netName
        complete     = ($macs.Count -gt 0)
    }
}

# Bulk pass over /api/servers?vm=true&stats=false, joined to the inventory on
# server.instance.id. Once per process, and again (rate-limited) only when many listed
# instances still have no MAC.
function Update-VmeHardwareBulkIfDue {
    $H = $script:Settings.inventory.hardware
    if (-not $H.bulk_enabled -or -not $script:InventoryLoaded) { return }
    if ($script:HwBulkDone) {
        if (([DateTime]::UtcNow - $script:HwBulkAt).TotalSeconds -lt $H.bulk_min_interval_seconds) { return }
        $missing = 0
        foreach ($id in $script:Inventory.Keys) { $hw = $script:HwById[$id]; if ($null -eq $hw -or -not $hw.complete) { $missing++ } }
        if ($missing -lt $H.bulk_rerun_missing_threshold) { return }
    }
    $script:HwBulkDone = $true
    $script:HwBulkAt = [DateTime]::UtcNow

    $pageSize = [int]$script:Settings.inventory.page_size
    $offset = 0
    $pages = 0
    $learned = 0
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        while ($true) {
            $resp = Invoke-VmeApi -Path "/api/servers?vm=true&stats=false&max=$pageSize&offset=$offset$($script:ServerScopeQuery)" -StatKey 'GET /api/servers (page)' -FailureLevel 'W'
            $arr = Get-JsonValue $resp 'servers'
            $batch = @(); if ($null -ne $arr) { $batch = @($arr) }
            foreach ($srv in $batch) {
                $iid = [string](Get-JsonValue (Get-JsonValue $srv 'instance') 'id')
                if ([string]::IsNullOrWhiteSpace($iid)) { continue }
                $rec = $script:Inventory[$iid]
                if ($null -eq $rec) { continue }
                $hw = ConvertTo-VmeHardware -Server $srv
                $hw.uuid = [string]$rec.uuid
                if ($hw.complete -or $null -eq $script:HwById[$iid]) { $script:HwById[$iid] = $hw; if ($hw.complete) { $learned++ } }
            }
            $pages++
            $total = $null
            try { $total = [int](Get-JsonValue (Get-JsonValue $resp 'meta') 'total') } catch { }
            if ($batch.Count -lt $pageSize) { break }
            if ($null -ne $total -and $total -gt 0 -and ($offset + $batch.Count) -ge $total) { break }
            if ($pages -ge [int]$script:Settings.inventory.max_pages) { break }
            $offset += $pageSize
        }
        Write-DebugLog "Hardware bulk pass: MAC/OS learned for $learned instance(s) from $pages page(s) in $([int]$sw.Elapsed.TotalMilliseconds)ms ($($script:HwById.Count) known)." -Level 'I' -Component '03'
    }
    catch {
        Write-DebugLog "Hardware bulk pass failed after $pages page(s): $($_.Exception.Message) -- MACs will be fetched per instance on guests/get meanwhile." -Level 'W' -Component '03'
    }
}

# Per instance, for one the bulk pass didn't cover (a new clone). Rate-limited while
# the server has no MAC yet, with the interval doubling per try (capped at 32x), so an
# instance that never gets a NIC doesn't cost a server GET every few seconds forever.
function Update-VmeHardwareForRecord {
    param([hashtable]$Rec, [switch]$Force)

    $id = $Rec.id
    $hw = $script:HwById[$id]
    if ($null -ne $hw -and $hw.complete -and $hw.uuid -eq $Rec.uuid) { return }
    if ([string]::IsNullOrWhiteSpace($Rec.server_id)) { return }
    $last = $script:HwLastTry[$id]
    if (-not $Force -and $null -ne $last) {
        $interval = [double]$script:Settings.inventory.hardware.hw_retry_seconds * [Math]::Min(32, [Math]::Pow(2, [Math]::Max(0, $last.tries - 1)))
        if (([DateTime]::UtcNow - $last.at).TotalSeconds -lt $interval) { return }
    }
    $script:HwLastTry[$id] = @{ at = [DateTime]::UtcNow; tries = $(if ($null -ne $last) { $last.tries + 1 } else { 1 }) }
    try {
        $resp = Invoke-VmeApi -Path "/api/servers/$($Rec.server_id)" -AllowNotFound -Ref $id -FailureLevel 'W'
        $srv = Get-JsonValue $resp 'server'
        if ($null -eq $srv) { return }
        $hw = ConvertTo-VmeHardware -Server $srv
        $hw.uuid = [string]$Rec.uuid
        $script:HwById[$id] = $hw
        if ($hw.complete) {
            $script:HwLastTry.Remove($id)
            $ck = Get-CloneKeyByCloneId -VmId $id
            if ($null -ne $ck) { Test-VmePreservedMacOutcome -Key $ck }
        }
        if (-not [string]::IsNullOrWhiteSpace($hw.network_id)) {
            $ck2 = Get-CloneKeyByCloneId -VmId $id
            if ($null -ne $ck2) { Test-VmeNetworkPreservationOutcome -Key $ck2 }
        }
        Write-DebugLog "Hardware for instance [$id]: server [$($hw.server_id)] mac=[$($hw.macs -join ',')] os=[$($hw.os)]." -Level 'T' -Component '03' -Ref $id
    }
    catch { }
}

# Unconditional one-call refresh of a clone source's server object, right before
# building the clone body. Update-VmeHardwareForRecord above deliberately skips a
# source whose hardware is already "complete" (has a MAC) -- correct for the routine
# per-guest backfill, but wrong here: the network on that cached server object can go
# stale for up to hardware.bulk_min_interval_seconds (default 900s) if nothing else
# happens to trigger a re-fetch, and a wrong VLAN on every clone from a re-VLANned
# template for that whole window is a worse outcome than one extra GET per clone
# submission (the instance's own interfaces[].network field
# is not kept in sync with a network change at the server/hypervisor level at all, so
# only this refresh -- not a shorter poll interval -- closes the gap).
function Update-VmeCloneSourceHardware {
    param([string]$SourceId)
    $rec = $script:Inventory[$SourceId]
    if ($null -eq $rec -or [string]::IsNullOrWhiteSpace($rec.server_id)) { return }
    try {
        $resp = Invoke-VmeApi -Path "/api/servers/$($rec.server_id)" -AllowNotFound -Ref $SourceId -FailureLevel 'W'
        $srv = Get-JsonValue $resp 'server'
        if ($null -eq $srv) { return }
        $hw = ConvertTo-VmeHardware -Server $srv
        $hw.uuid = [string]$rec.uuid
        $script:HwById[$SourceId] = $hw
    }
    catch { }
}

# ---------------------------------------------------------------------------
# Guest objects
# ---------------------------------------------------------------------------

function ConvertTo-RasGuestObject {
    param([Parameter(Mandatory = $true)][string]$VmId, [Parameter(Mandatory = $true)][hashtable]$Rec, [switch]$NoPollStamp)

    if (-not $NoPollStamp) { $script:LastGuestPollAt[$VmId] = [DateTime]::UtcNow }

    $state = if ($script:DeletingIds.ContainsKey($VmId)) { 'powering_off' } else { $Rec.ras_state }
    $ips = $Rec.ipv4
    if ($state -eq 'powered_off' -and $script:Settings.inventory.hide_ip_when_stopped) { $ips = [string[]]@() }
    $hw = $script:HwById[$VmId]
    $macs = [string[]]@(); if ($null -ne $hw) { $macs = [string[]]$hw.macs }
    $os = if ($null -ne $hw) { $hw.os } else { 'unknown' }
    $isTemplate = (Test-VmeIsBasicTemplate -VmId $VmId) -and -not (Test-VmeInMaintenanceMode -VmId $VmId)

    $guest = @{
        id            = $VmId
        name          = $(if ($Rec.name) { $Rec.name } else { "VM-$VmId" })
        provider      = 'VME'
        node          = $Rec.cloud
        state         = $state
        power_state   = $Rec.status
        host_os       = $os
        ip            = $(if ($ips.Count -gt 0) { $ips[0] } else { $null })
        ip_addresses  = @($ips)
        mac_addresses = @($macs)
        is_template   = $isTemplate
        type          = 'Virtual Machine'
    }

    # Workaround (cloning.template_flapping_workaround_enabled): a Ready template is shown
    # powered off, unless RAS explicitly started it this session.
    if ($isTemplate -and $script:Settings.cloning.template_flapping_workaround_enabled -and -not $script:TemplateFlappingUnmaskedVmIds.Contains($VmId)) {
        if ($guest.state -ne 'powered_off' -or $guest.ip_addresses.Count -gt 0) {
            Write-DebugLog "Template flapping workaround: instance [$VmId] reported state=[$($guest.state)] status=[$($Rec.status)] -- masked to powered_off." -Level 'W' -Component '06' -Ref $VmId
        }
        $guest.state = 'powered_off'
        $guest.power_state = 'stopped'
        $guest.ip = $null
        $guest.ip_addresses = @()
        $guest.mac_addresses = @()
    }

    if ($script:LogLevel -ge 5) {
        Write-DebugLog "GUEST ID=$VmId; Name=$($guest.name); State=$($guest.state); Status=$($Rec.status); Template=$isTemplate; IPs=$($guest.ip_addresses -join ','); MACs=$($guest.mac_addresses -join ',')" -Level 'D' -Component '03' -Ref $VmId
    }
    return $guest
}

# Powers on a fresh clone found off. Never the template, and never an instance VME is
# still provisioning or busy with (VME starts clones itself; a start during Deploy may
# be refused).
function Start-VmeInstanceIfNeeded {
    param([Parameter(Mandatory = $true)][hashtable]$Rec)

    $id = $Rec.id
    if (Test-VmeIsBasicTemplate -VmId $id) {
        Write-DebugLog "Clone-aware flow refused to start instance [$id] -- it is flagged as a RAS template. Check the 'resolved' line for this clone's task." -Level 'E' -Component '04' -Ref $id
        return
    }
    if (Test-VmeInstanceProvisioning -Rec $Rec) {
        Write-DebugLog "Clone-aware flow: instance [$id] is still busy/provisioning (status=$($Rec.status)); not starting it." -Level 'T' -Component '04' -Ref $id
        return
    }
    if ($Rec.ras_state -in @('powered_on', 'powering_on')) { return }
    try {
        Write-DebugLog "Clone-aware flow: starting instance [$id] ($($Rec.name), status=$($Rec.status))." -Level 'I' -Component '04' -Ref $id
        [void](Invoke-VmeApi -Path "/api/instances/$id/start" -Method PUT -Body @{} -Idempotent -Ref $id -FailureLevel 'W')
        $Rec.dirty = $true
        $Rec.ras_state = 'powering_on'
    }
    catch { Write-DebugLog "Clone-aware flow: start of instance [$id] failed: $($_.Exception.Message)" -Level 'W' -Component '04' -Ref $id }
}

# Records VME's provisioning outcome for a tracked clone the first time it is terminal:
# the clone frees its clone slot. A failure after tasks/get reported the clone early
# (linked_clones.early_completion, deploy-event gate) is the accepted risk of reporting
# early -- RAS already counts the clone as created. Returns the provision class.
function Set-VmeCloneProvisionOutcome {
    param([string]$Key, [hashtable]$Entry, [hashtable]$Rec, [string]$VmId)
    $class = Get-VmeProvisionStatusClass -Status $Rec.status -IsBusy $Rec.is_busy
    if ($class -eq 'pending' -or $Entry.provision_confirmed) { return $class }
    $Entry.provision_confirmed = $true
    Set-CloneEntry -Key $Key -Entry $Entry
    $after = (Get-EpochNow) - [long]$Entry.started_epoch
    if ($class -eq 'failed' -and $null -ne $Entry.reported_completed_epoch) {
        Write-DebugLog "Clone [$($Entry.name)] (instance [$VmId]) ended in status [$($Rec.status)] ${after}s after guests/clone, after tasks/get had reported it completed early -- RAS believes this clone succeeded.$(if ($Rec.status_message) { " VME says: $($Rec.status_message)" })" -Level 'E' -Component '04' -Ref $VmId
    }
    else { Write-DebugLog "Clone [$($Entry.name)] (instance [$VmId]) finished provisioning (status=$($Rec.status)) ${after}s after guests/clone; its clone slot is free." -Level 'T' -Component '04' -Ref $VmId }
    return $class
}

# Clones reported early keep their slot until VME finishes them, which guests/get only
# notices when RAS asks about that clone. A tasks/get held at the clone limit therefore
# re-reads the oldest of them itself: one lean GET each, at most
# tracked_sweep_max_per_call per poll, each clone at most every tracked_refresh_interval_seconds.
function Update-VmeEarlyCloneSlots {
    param([string]$ExceptKey, [string]$Ref)
    $limit = [Math]::Max(1, [int]$script:Settings.cloning.tracked_sweep_max_per_call)
    $interval = [int]$script:Settings.cloning.tracked_refresh_interval_seconds
    $now = [DateTime]::UtcNow
    $cands = [System.Collections.Generic.List[object]]::new()
    foreach ($k in @($script:CloneState.Keys)) {
        $e = $script:CloneState[$k]
        if ($k -eq $ExceptKey -or $e.provision_confirmed -or $null -eq $e.reported_completed_epoch -or -not $e.clone_id) { continue }
        $last = $script:CloneLastRefresh[[string]$e.clone_id]
        if ($null -ne $last -and ($now - $last).TotalSeconds -lt $interval) { continue }
        $cands.Add(@{ key = $k; entry = $e })
    }
    $done = 0
    foreach ($c in @($cands | Sort-Object { [long]$_.entry.started_epoch })) {
        if ($done -ge $limit) { break }
        $done++
        $cid = [string]$c.entry.clone_id
        $script:CloneLastRefresh[$cid] = [DateTime]::UtcNow
        $rec = $null
        try { $rec = Get-VmeInstanceRecord -VmId $cid -Ref $Ref -FailureLevel 'W' } catch { continue }
        if ($null -eq $rec) {
            Write-DebugLog "Clone [$($c.entry.name)] (instance [$cid]), reported to RAS early, is gone from VME (404) -- its clone slot is free." -Level 'W' -Component '04' -Ref $cid
            Clear-VmeTrackingForVm -VmId $cid -Gone
            continue
        }
        [void](Set-VmeCloneProvisionOutcome -Key $c.key -Entry $c.entry -Rec $rec -VmId $cid)
    }
}

# A fresh clone is babysat after tasks/get reported it: re-read from VME at most every
# tracked_refresh_interval_seconds, started if found off (reported powering_on then),
# reported powered_on as soon as VME says it runs -- with or without an IP.
# Tracking ends at the first "up with a real IP", on a failure, on 404, or when the
# babysit window elapses -- never on a flag some path might forget to set.
function Get-RasGuestObjectForTrackedClone {
    param([Parameter(Mandatory = $true)][string]$VmId, [Parameter(Mandatory = $true)][string]$Key, [switch]$FromSweep)

    $entry = $script:CloneState[$Key]
    $now = Get-EpochNow
    # babysit_window_seconds runs from tasks/get's completion report, as documented --
    # measured from the submit, a clone whose provisioning ate most of
    # provision_timeout_seconds got far less than the window. Before the report the
    # provisioning budget applies.
    # Invoke-CacheGarbageCollection uses the same base (Get-VmeCloneWindow).
    $window = Get-VmeCloneWindow -Entry $entry
    if (($now - $window.base) -gt $window.limit) {
        Write-DebugLog "Clone-aware flow: babysit window elapsed for instance [$VmId] ($($entry.name)); tracking retired." -Level 'W' -Component '04' -Ref $VmId
        if (-not $entry.provision_confirmed) { Add-VmeStuckClone -Entry $entry -Reason 'babysit window elapsed while VME was still provisioning it' }
        Clear-VmeTrackingForVm -VmId $VmId -KeepRecord
        $rec = $script:Inventory[$VmId]
        if ($null -eq $rec) { $rec = Get-VmeInstanceRecord -VmId $VmId }
        if ($null -eq $rec) { throw (New-VmeApiException -Message "$($script:ProviderNamePrefix) Guest [$VmId] not found" -StatusCode 404) }
        return (ConvertTo-RasGuestObject -VmId $VmId -Rec $rec -NoPollStamp:$FromSweep)
    }

    $rec = $script:Inventory[$VmId]
    $last = $script:CloneLastRefresh[$VmId]
    if ($null -eq $rec -or $rec.dirty -or $null -eq $last -or ([DateTime]::UtcNow - $last).TotalSeconds -ge $script:Settings.cloning.tracked_refresh_interval_seconds) {
        $rec = Get-VmeInstanceRecord -VmId $VmId
        $script:CloneLastRefresh[$VmId] = [DateTime]::UtcNow
        if ($null -eq $rec) {
            Write-DebugLog "Clone-aware flow: tracked clone [$VmId] ($($entry.name)) returns 404 -- it was deleted or rolled back by VME; tracking retired." -Level 'E' -Component '04' -Ref $VmId
            Clear-VmeTrackingForVm -VmId $VmId -Gone
            throw (New-VmeApiException -Message "$($script:ProviderNamePrefix) Guest [$VmId] not found" -StatusCode 404)
        }
    }

    $class = Set-VmeCloneProvisionOutcome -Key $Key -Entry $entry -Rec $rec -VmId $VmId
    if ($class -eq 'failed') {
        Clear-VmeTrackingForVm -VmId $VmId -KeepRecord
        return (ConvertTo-RasGuestObject -VmId $VmId -Rec $rec -NoPollStamp:$FromSweep)
    }
    # Tracked clones are few and already throttled by tracked_refresh_interval_seconds:
    # their server object is re-read until the MAC shows up, without the per-instance back-off.
    if ($class -eq 'ready') { Update-VmeHardwareForRecord -Rec $rec -Force }

    $guest = ConvertTo-RasGuestObject -VmId $VmId -Rec $rec -NoPollStamp:$FromSweep
    $hasRealIp = Test-VmeHasUsableIpv4Address -Addresses $rec.ipv4

    if ($class -eq 'pending') {
        $guest.state = 'powering_on'
    }
    elseif ($rec.ras_state -in @('powered_off', 'powering_off')) {
        Start-VmeInstanceIfNeeded -Rec $rec
        $guest.state = 'powering_on'
    }
    elseif (-not $hasRealIp) {
        # Power state, not the IP, decides: like a native
        # provider, a running clone is powered_on; RAS then waits for its guest agent to
        # report back and runs RASPrep. Reporting powering_on until an IP cost one of RAS's
        # fixed 60 s waits per clone (CloneStepPowerOnGuestIfNot).
        # Tracking continues until the first real IP (start-if-off safety net, fast refresh).
        Write-DebugLog "Clone-aware flow: instance [$VmId] is running, no usable IP yet; reporting powered_on." -Level 'T' -Component '04' -Ref $VmId
    }
    elseif ($rec.ras_state -eq 'powered_on') {
        $elapsed = $now - [long]$entry.started_epoch
        Write-DebugLog "Clone-aware flow: instance [$VmId] ($($entry.name)) is up with IP $($guest.ip), ${elapsed}s after guests/clone; tracking retired." -Level 'I' -Component '04' -Ref $VmId
        Clear-VmeTrackingForVm -VmId $VmId -KeepRecord
    }
    return $guest
}

# Re-checks OTHER clones already reported to RAS opportunistically (RAS may not poll a
# clone it has not bound yet). Globally throttled and capped per call, so a pool build with many
# clones waiting for an IP never turns one guests/get into many HTTP calls.
function Invoke-TrackedCloneSweep {
    param([string]$ExceptVmId)

    $max = [int]$script:Settings.cloning.tracked_sweep_max_per_call
    if ($max -le 0 -or $script:CloneKeyByCloneId.Count -eq 0) { return }
    $interval = [int]$script:Settings.cloning.tracked_refresh_interval_seconds
    if (([DateTime]::UtcNow - $script:SweepLastRunAt).TotalSeconds -lt $interval) { return }
    $script:SweepLastRunAt = [DateTime]::UtcNow

    $done = 0
    foreach ($cid in @($script:CloneKeyByCloneId.Keys)) {
        if ($cid -eq $ExceptVmId) { continue }
        $last = $script:CloneLastRefresh[$cid]
        if ($null -ne $last -and ([DateTime]::UtcNow - $last).TotalSeconds -lt $interval) { continue }
        $key = $script:CloneKeyByCloneId[$cid]
        if ($null -eq $key) { continue }
        # Not yet reported: tasks/get is still polling it, a second GET would be waste.
        if ($null -eq $script:CloneState[$key].reported_completed_epoch) { continue }
        try { [void](Get-RasGuestObjectForTrackedClone -VmId $cid -Key $key -FromSweep) }
        catch { Write-DebugLog "Tracked-clone sweep for [$cid]: $($_.Exception.Message)" -Level 'T' -Component '04' -Ref $cid }
        $done++
        if ($done -ge $max) { break }
    }
}

# ---------------------------------------------------------------------------
# Delete visibility
#
# guests/list must keep listing an instance until VME confirms it gone (RAS's
# recreate-with-the-same-name flow waits for that), but must
# not need a full listing to notice. Each id this provider deleted is probed with one
# lean GET per probe_interval_seconds; 404 removes it.
# ---------------------------------------------------------------------------

function Invoke-VmeDeletingProbe {
    if ($script:DeletingIds.Count -eq 0 -or $null -eq $script:VmeSession) { return }
    $DS = $script:Settings.delete
    $now = [DateTime]::UtcNow
    $due = @($script:DeletingIds.Keys | Where-Object { ($now - $script:DeletingIds[$_].last_probe).TotalSeconds -ge $DS.probe_interval_seconds } |
            Sort-Object { $script:DeletingIds[$_].last_probe } | Select-Object -First ([int]$DS.max_probes_per_list))

    foreach ($id in $due) {
        $entry = $script:DeletingIds[$id]
        $entry.last_probe = [DateTime]::UtcNow
        $rec = $null
        try { $rec = Get-VmeInstanceRecord -VmId $id -FailureLevel 'W' }
        catch { continue }
        $age = [int]([DateTime]::UtcNow - $entry.submitted_at).TotalSeconds
        if ($null -eq $rec) {
            Write-DebugLog "Instance [$id] ($($entry.name)) is gone (VME returns 404), ${age}s after the delete." -Level 'I' -Component '05' -Ref $id
            $script:DeletingIds.Remove($id)
            Clear-VmeTrackingForVm -VmId $id -Gone
            continue
        }
        $status = ([string]$rec.status).ToLowerInvariant()
        if ($age -ge $DS.give_up_after_seconds) {
            Write-DebugLog "Instance [$id] ($($entry.name)) still exists ${age}s after the delete (status=$($rec.status)) -- no longer probing; it is listed with its real state again." -Level 'W' -Component '05' -Ref $id
            $script:DeletingIds.Remove($id)
            continue
        }
        if ($age -ge $DS.stuck_after_seconds -and -not $entry.stuck_logged -and $status -notin @('removing', 'pendingremoval', 'stopping', 'deleting')) {
            $entry.stuck_logged = $true
            Write-DebugLog "Instance [$id] ($($entry.name)) is still [$($rec.status)] ${age}s after the delete -- the delete may not have taken. Still listed as powering_off." -Level 'W' -Component '05' -Ref $id
        }
    }
}

# ---------------------------------------------------------------------------
# Orphan audit and garbage collection
# ---------------------------------------------------------------------------

# Log-only lead: an instance cloned (server-side lineage, config.provisionOpts.
# clonedInstanceId) from a template this provider knows, not tracked, not being
# deleted, which RAS has not polled for stale_poll_after_seconds. No local ledger:
# the lineage lives on the appliance.
function Invoke-OrphanAudit {
    $O = $script:Settings.orphan_detection
    if (([DateTime]::UtcNow - $script:OrphanAuditLastCheckedAt).TotalSeconds -lt $O.check_interval_seconds) { return }
    $script:OrphanAuditLastCheckedAt = [DateTime]::UtcNow

    Invoke-CacheGarbageCollection
    if (-not $O.enabled -or -not $script:InventoryLoaded) { return }

    $sources = [System.Collections.Generic.HashSet[string]]::new([string[]]@($script:BasicTemplates.Keys))
    foreach ($s in $script:KnownCloneSources) { [void]$sources.Add($s) }
    if ($sources.Count -eq 0) { return }

    $now = [DateTime]::UtcNow
    $count = 0
    foreach ($id in $script:Inventory.Keys) {
        $rec = $script:Inventory[$id]
        # Full clones carry clonedInstanceId; linked clones carry their version's image id.
        $from = [string]$rec.cloned_from
        if ([string]::IsNullOrWhiteSpace($from) -and -not [string]::IsNullOrWhiteSpace([string]$rec.image_id)) {
            $v = $script:VersionByImageId[[string]$rec.image_id]
            if ($null -ne $v) { $from = $v.template_id }
        }
        if ([string]::IsNullOrWhiteSpace($from) -or -not $sources.Contains($from)) { continue }
        if ($null -ne (Get-CloneKeyByCloneId -VmId $id) -or $script:DeletingIds.ContainsKey($id) -or (Test-VmeInstanceExcluded -Name $rec.name)) { continue }
        $lastPolled = $script:LastGuestPollAt[$id]
        $reference = if ($null -ne $lastPolled) { $lastPolled } else { $script:ProviderStartedAt }
        $stale = ($now - $reference).TotalSeconds
        if ($stale -lt $O.stale_poll_after_seconds) { continue }
        $count++
        if ($script:OrphanReported.Add($id)) {
            $desc = if ($null -ne $lastPolled) { "$([int]$stale)s ago" } else { "never since this provider process started $([int]$stale)s ago" }
            Write-DebugLog "ORPHAN CANDIDATE: instance [$id] ($($rec.name)) is a clone of template [$from], is not tracked, and RAS last polled it $desc -- RAS may have lost track of it. Manual review recommended." -Level 'W' -Component '04' -Ref $id
        }
    }
    if ($count -gt 0) { Write-DebugLog "Orphan audit: $count candidate(s) ($($script:OrphanReported.Count) reported so far this process)." -Level 'I' -Component '04' }
}

# Bounds every per-id map by the inventory, and expires the time-bound ones. Runs on
# the orphan audit's timer whether or not the audit itself is enabled.
# How long a clone entry is legitimately alive, and from when: before tasks/get reported
# it, the provisioning budget from the submit; after, babysit_window_seconds from the
# report. Shared by the babysit flow and the garbage collector so the collector can
# never drop an entry the babysit flow still considers inside its window.
function Get-VmeCloneWindow {
    param([Parameter(Mandatory = $true)][System.Collections.IDictionary]$Entry)
    $C = $script:Settings.cloning
    if ($null -ne $Entry.reported_completed_epoch) { return @{ base = [long]$Entry.reported_completed_epoch; limit = [long]$C.babysit_window_seconds } }
    return @{ base = [long]$Entry.started_epoch; limit = [long][Math]::Max([long]$C.resolve_timeout_seconds, [long]$C.provision_timeout_seconds) }
}

function Invoke-CacheGarbageCollection {
    $now = [DateTime]::UtcNow
    $retention = $script:Settings.cloning.completed_clone_task_output_retention_seconds
    foreach ($k in @($script:CompletedCloneTaskOutputs.Keys)) { if (($now - $script:CompletedCloneTaskOutputs[$k].stored_at).TotalSeconds -ge $retention) { $script:CompletedCloneTaskOutputs.Remove($k) } }
    $macRetention = $script:Settings.cloning.mac_preservation.retention_seconds
    foreach ($k in @($script:RecentlyDeletedMacByName.Keys)) { if (($now - $script:RecentlyDeletedMacByName[$k].deleted_at).TotalSeconds -ge $macRetention) { $script:RecentlyDeletedMacByName.Remove($k) } }
    # Snapshot delete/revert tasks RAS stopped polling.
    $snapRetention = [int]$script:Settings.linked_clones.snapshot_timeout_seconds + 3600
    foreach ($k in @($script:SnapshotTasks.Keys)) { if (($now - $script:SnapshotTasks[$k].started).TotalSeconds -ge $snapRetention) { $script:SnapshotTasks.Remove($k) } }
    # Power watches RAS never asked about again, and stale watch stamps.
    $holdEnd = (Get-EpochNow) - [int]$script:Settings.inventory.power_action_hold_seconds
    foreach ($k in @($script:PowerWatch.Keys)) { if ([long]$script:PowerWatch[$k].since -lt $holdEnd) { $script:PowerWatch.Remove($k) } }
    foreach ($k in @($script:WatchReadAt.Keys)) { if (($now - $script:WatchReadAt[$k]).TotalSeconds -ge 3600) { $script:WatchReadAt.Remove($k) } }

    if ($script:InventoryLoaded) {
        $keep = { param([string]$id) $script:Inventory.ContainsKey($id) -or $script:DeletingIds.ContainsKey($id) -or $null -ne $script:CloneKeyByCloneId[$id] }
        foreach ($map in @($script:HwById, $script:HwLastTry, $script:LastGuestPollAt, $script:StatusLogged, $script:CloneBodyBySource, $script:CloneLastRefresh)) {
            foreach ($id in @($map.Keys)) { if (-not (& $keep $id)) { $map.Remove($id) } }
        }
        foreach ($set in @($script:ExcludeLogged, $script:OrphanReported, $script:TemplateFlappingUnmaskedVmIds, $script:KnownCloneSources)) {
            foreach ($id in @($set)) { if (-not (& $keep $id)) { [void]$set.Remove($id) } }
        }
    }

    # A clone entry nobody will ever retire (RAS stopped polling the task and the
    # guest): drop it once it is well past every budget.
    $epoch = Get-EpochNow
    foreach ($key in @($script:CloneState.Keys)) {
        $e = $script:CloneState[$key]
        $window = Get-VmeCloneWindow -Entry $e
        if (($epoch - $window.base) -gt ($window.limit + 600)) {
            Write-DebugLog "Clone tracking [$key] ($($e.name)) is $($epoch - [long]$e.started_epoch)s old and past every budget -- dropped (RAS stopped polling it)." -Level 'W' -Component '04' -Ref $key
            if (-not $e.provision_confirmed) { Add-VmeStuckClone -Entry $e -Reason 'RAS stopped polling it' }
            Remove-CloneEntry -Key $key
        }
    }
}

# ---------------------------------------------------------------------------
# Process history (diagnostics and the deploy-event gate only)
#
# Clone, start, stop and delete all return {"success":true} with no process id, so
# nothing is ever polled by process id. /api/processes?instanceId= is read only for
# the gate (once per poll, when enabled) and for the error text of a failed clone.
# ---------------------------------------------------------------------------

function Get-VmeNewestProvisionProcess {
    param([Parameter(Mandatory = $true)][string]$VmId, [string]$Ref = $null)

    $resp = Invoke-VmeApi -Path "/api/processes?instanceId=$VmId&max=10" -Ref $Ref -FailureLevel 'W'
    $all = @(@(Get-JsonValue $resp 'processes') | Where-Object { $null -ne $_ })
    if ($all.Count -eq 0) { return $null }
    $provision = @($all | Where-Object { [string](Get-JsonValue (Get-JsonValue $_ 'processType') 'code') -eq 'provision' })
    $pool = if ($provision.Count -gt 0) { $provision } else { $all }
    $best = $null
    $bestEpoch = [long]-1
    foreach ($p in $pool) {
        $e = ConvertTo-EpochSeconds (Get-JsonValue $p 'startDate')
        if ($null -eq $e) { $e = 0 }
        if ($e -gt $bestEpoch) { $best = $p; $bestEpoch = $e }
    }
    return $best
}

function Get-VmeProcessErrorText {
    param([object]$Process)
    if ($null -eq $Process) { return $null }
    foreach ($k in @('error', 'output', 'message')) {
        $v = [string](Get-JsonValue $Process $k)
        if (-not [string]::IsNullOrWhiteSpace($v)) { return $v }
    }
    # A failure often surfaces on a sub-step, not on the parent row.
    foreach ($evt in @(Get-JsonValue $Process 'events')) {
        if ([string](Get-JsonValue $evt 'status') -eq 'failed') {
            $err = [string](Get-JsonValue $evt 'error')
            if ($err) { return "$([string](Get-JsonValue (Get-JsonValue $evt 'processType') 'name')) failed: $err" }
        }
    }
    return $null
}

# Best-effort: called when the provider gives up on a clone (provision_timeout_seconds)
# so RAS abandoning a slow/bulk clone doesn't leave VME's own job running unattended
# ("another thread still active" after RAS or the admin moved on).
# Never throws and never changes the caller's outcome; VME's own "not cancelable" error
# (e.g. the process already finished) is just logged, not treated as a failure.
function Invoke-VmeProcessCancel {
    param([string]$ProcessId, [string]$Ref = $null)
    if ([string]::IsNullOrWhiteSpace($ProcessId)) { return }
    try {
        $resp = Invoke-VmeApi -Path "/api/processes/$ProcessId/cancel" -Method POST -Body @{} -Ref $Ref -FailureLevel 'W'
        if ($true -eq (Get-JsonValue $resp 'success')) {
            Write-DebugLog "Process [$ProcessId] cancelled on VME (the provider gave up waiting for it)." -Level 'I' -Component '08' -Ref $Ref
        }
        else {
            Write-DebugLog "Process [$ProcessId] cancel request returned success=false -- VME says it isn't cancelable, or it already finished." -Level 'W' -Component '08' -Ref $Ref
        }
    }
    catch { Write-DebugLog "Failed to cancel process [$ProcessId]: $($_.Exception.Message)" -Level 'W' -Component '08' -Ref $Ref }
}

# 'complete' | 'pending' | 'unknown' for VME's "Deploy Virtual Machine" step
# (provisionDeploy). Missing data is never 'complete'.
function Get-VmeDeployEventStatus {
    param([object]$Process)
    if ($null -eq $Process) { return 'unknown' }
    foreach ($evt in @(Get-JsonValue $Process 'events')) {
        $pt = Get-JsonValue $evt 'processType'
        $code = [string](Get-JsonValue $pt 'code')
        $name = [string](Get-JsonValue $pt 'name')
        if ($code -ne 'provisionDeploy' -and $name -notmatch '(?i)deploy') { continue }
        if ([string](Get-JsonValue $evt 'status') -eq 'complete') { return 'complete' }
        return 'pending'
    }
    return 'unknown'
}

function Write-VmeProvisionTrace {
    param([object]$Process, [string]$VmId, [string]$Ref, [long]$Elapsed)
    if ($script:LogLevel -lt 4) { return }
    if ($null -eq $Process) { Write-DebugLog "Process history: nothing yet for instance [$VmId] (elapsed=${Elapsed}s)." -Level 'T' -Component '08' -Ref $Ref; return }
    $events = @(foreach ($evt in @(Get-JsonValue $Process 'events')) {
            $pt = Get-JsonValue $evt 'processType'
            "$([string](Get-JsonValue $pt 'code'))=$([string](Get-JsonValue $evt 'status'))"
        })
    Write-DebugLog "Process history: instance [$VmId] status=$([string](Get-JsonValue $Process 'status')) elapsed=${Elapsed}s events: $($events -join ', ')" -Level 'T' -Component '08' -Ref $Ref
}

# Re-checks the given-up clones on VME and frees the slot of each one VME confirms
# ended. Only called while the clone limit blocks a guests/clone, and at most every
# stuck_clone_recheck_seconds. Never frees a slot on missing data: a failed call keeps it.
#   - known instance id: one lean GET -- 404, or provisioning finished or failed, frees it;
#   - never appeared: first a late resolution by name (the clone may have appeared since),
#     then the template's process history (a stuck clone that never
#     produced an instance shows as an unfinished 'provision' process on the SOURCE's
#     history). A process whose displayName is the clone's name decides for that clone;
#     otherwise, if the history holds fewer unfinished provision processes than there are
#     never-appeared given-up clones of that template, the oldest surplus is freed.
function Update-VmeStuckClones {
    param([switch]$Force)
    if ($script:StuckClones.Count -eq 0 -or -not (Test-VmeConnected)) { return }
    if (-not $Force -and ([DateTime]::UtcNow - $script:StuckClonesCheckedAt).TotalSeconds -lt [int]$script:Settings.cloning.stuck_clone_recheck_seconds) { return }
    $script:StuckClonesCheckedAt = [DateTime]::UtcNow

    $neverAppeared = @{}   # source id -> list of stuck keys
    foreach ($key in @($script:StuckClones.Keys)) {
        $s = $script:StuckClones[$key]
        try {
            if (-not $s.clone_id) {
                $found = Find-VmeCloneInstance -Entry $s -TaskId $s.task_id
                $claimed = $null -ne $found -and @($script:StuckClones.Values | Where-Object { $_.clone_id -eq $found.record.id }).Count -gt 0
                if ($null -ne $found -and -not $claimed) {
                    $s.clone_id = [string]$found.record.id
                    Save-StuckClones
                    Write-DebugLog "Given-up clone [$($s.name)] appeared late as instance [$($s.clone_id)] (status [$($found.record.status)], via $($found.method))." -Level 'I' -Component '04' -Ref $s.task_id
                }
            }
            if ($s.clone_id) {
                $rec = Get-VmeInstanceRecord -VmId $s.clone_id -Ref $s.task_id -FailureLevel 'W'
                if ($null -eq $rec) { Remove-VmeStuckClone -Key $key -Why 'VME returns 404 for it'; continue }
                $class = Get-VmeProvisionStatusClass -Status $rec.status -IsBusy $rec.is_busy
                if ($class -ne 'pending') { Remove-VmeStuckClone -Key $key -Why "VME $(if ($class -eq 'failed') { 'failed' } else { 'finished' }) provisioning it (status [$($rec.status)])"; continue }
                if ($s.last_status -ne $rec.status) { $s.last_status = [string]$rec.status; Save-StuckClones }
                continue
            }
            if (-not $neverAppeared.ContainsKey($s.source_id)) { $neverAppeared[$s.source_id] = [System.Collections.Generic.List[string]]::new() }
            $neverAppeared[$s.source_id].Add($key)
        }
        catch { Write-DebugLog "Re-check of given-up clone [$($s.name)] failed: $($_.Exception.Message) -- its slot is kept." -Level 'W' -Component '04' -Ref $s.task_id }
    }

    foreach ($src in @($neverAppeared.Keys)) {
        $keys = @($neverAppeared[$src] | Sort-Object { [long]$script:StuckClones[$_].started_epoch })
        try {
            $resp = Invoke-VmeApi -Path "/api/processes?instanceId=$src&max=100" -Ref $src -FailureLevel 'W' -StatKey 'GET /api/processes (clone slots)'
        }
        catch { Write-DebugLog "Process history of template [$src] could not be read -- $($keys.Count) given-up clone slot(s) kept." -Level 'W' -Component '04' -Ref $src; continue }
        $all = @(@(Get-JsonValue $resp 'processes') | Where-Object { $null -ne $_ })
        $total = Get-JsonValue (Get-JsonValue $resp 'meta') 'total'
        $truncated = ($null -ne $total -and [long]$total -gt $all.Count)
        $provision = @($all | Where-Object { [string](Get-JsonValue (Get-JsonValue $_ 'processType') 'code') -eq 'provision' })
        $isActive = { param($p) [string](Get-JsonValue $p 'status') -in @('new', 'queued', 'pending', 'running') }
        $unmatchedActive = [System.Collections.Generic.List[object]]::new()
        foreach ($p in $provision) { if (& $isActive $p) { $unmatchedActive.Add($p) } }

        $byCount = [System.Collections.Generic.List[string]]::new()
        foreach ($key in $keys) {
            $s = $script:StuckClones[$key]
            $mine = @($provision | Where-Object { [string](Get-JsonValue $_ 'displayName') -eq $s.name } | Sort-Object { ConvertTo-EpochSeconds (Get-JsonValue $_ 'startDate') } -Descending)
            if ($mine.Count -eq 0) { $byCount.Add($key); continue }
            if (& $isActive $mine[0]) { [void]$unmatchedActive.Remove($mine[0]); continue }
            Remove-VmeStuckClone -Key $key -Why "its provision process [$([string](Get-JsonValue $mine[0] 'id'))] on template [$src] ended with status [$([string](Get-JsonValue $mine[0] 'status'))] and no instance appeared"
        }
        $surplus = $byCount.Count - $unmatchedActive.Count
        if ($surplus -le 0) { continue }
        if ($truncated) {
            Write-DebugLog "Process history of template [$src] is truncated ($($all.Count) of $total rows) -- not freeing never-appeared clone slots by count." -Level 'W' -Component '04' -Ref $src
            continue
        }
        for ($i = 0; $i -lt $surplus; $i++) {
            Remove-VmeStuckClone -Key $byCount[$i] -Why "template [$src]'s process history shows $($unmatchedActive.Count) unfinished provision process(es) for $($byCount.Count) never-appeared given-up clone(s)"
        }
    }
}

# ---------------------------------------------------------------------------
# Provider handlers
# ---------------------------------------------------------------------------

function Test-VmeConnected { return ($null -ne $script:VmeSession) }

function New-NotConnectedResponse { return (New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Not connected. Call provider/connect first.") }

function Handle-Initialize {
    $C = $script:Settings.capabilities
    return @{
        result = @{
            version      = $script:ProviderVersion
            capabilities = @{
                can_suspend_guests    = [bool]$C.can_suspend_guests
                guests_polling_rate   = [int]$C.guests_polling_rate
                tasks_polling_rate    = [int]$C.tasks_polling_rate
                tasks_polling_retries = [int]$C.tasks_polling_retries
                template_method       = $script:TemplateMethod
                # Linked clones only through template versions (VME 9.1+: snapshot ->
                # linked-clone image -> addInstance). On 9.0.x the linked-clone endpoint
                # is a stub, so 'basic' stays full-clone only (docs/LINKED-CLONES.md).
                can_link_clones       = ($script:TemplateMethod -eq 'versioning')
            }
        }
    }
}

function Handle-Connect {
    param([System.Collections.IDictionary]$Params)

    try {
        $settings = $Params['settings']
        if ($settings -isnot [System.Collections.IDictionary]) { return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) 'settings' must be an object" }
        $hostValue = ([string]$settings['host']).Trim()
        if ([string]::IsNullOrWhiteSpace($hostValue)) { return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) 'host' is required" }
        $baseUrl = if ($hostValue -match '^https?://') { $hostValue.TrimEnd('/') } else { "https://$($hostValue.TrimEnd('/'))" }

        if ($null -ne $settings['ssl_check']) {
            try { $script:SslCheckOverride = ConvertTo-CoercedSettingBool -Value $settings['ssl_check'] -Default $true } catch { }
            if ($null -ne $script:SslCheckOverride) { $script:SslCheck = [bool]$script:SslCheckOverride }
        }
        # A fresh connection pool per connect (the host or TLS settings may differ).
        $script:WebSession = $null

        $refreshToken = [string]$settings['refresh_token']
        $username = $null
        $password = $null
        $token = $null
        $expiresAt = $null
        if (-not [string]::IsNullOrWhiteSpace([string]$settings['access_token'])) {
            $token = [string]$settings['access_token']
            if ([string]::IsNullOrWhiteSpace($refreshToken)) { Write-DebugLog 'Connecting with a pre-issued access_token and no refresh_token -- it cannot be renewed; RAS must reconnect with a new one when it expires.' -Level 'W' -Component '02' }
        }
        else {
            $username = [string]$settings['username']
            $password = [string]$settings['password']
            if ([string]::IsNullOrWhiteSpace($username) -or [string]::IsNullOrWhiteSpace($password)) {
                return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Either 'access_token' or both 'username' and 'password' are required"
            }
            $auth = Invoke-VmeAuth -BaseUrl $baseUrl -Username $username -Password $password
            $token = $auth.token
            $expiresAt = $auth.expires_at
            if (-not [string]::IsNullOrWhiteSpace([string]$auth.refresh_token)) { $refreshToken = $auth.refresh_token }
        }

        $session = @{
            baseUrl                 = $baseUrl
            token                   = $token
            header                  = @{ Authorization = "BEARER $token" }
            username                = $username
            password                = $password
            expires_at              = $expiresAt
            refresh_token           = $refreshToken
            expiry_last_checked_at  = [DateTime]::MinValue
            renewal_attempts        = 0
            last_renewal_attempt_at = $null
            next_token_check        = [DateTime]::MinValue
        }

        # Another appliance: nothing learned about the previous one applies.
        if ($null -ne $script:LastBaseUrl -and $script:LastBaseUrl -ne $baseUrl) {
            Write-DebugLog "Connecting to a different appliance ([$($script:LastBaseUrl)] -> [$baseUrl]) -- inventory and hardware caches cleared." -Level 'W' -Component '02'
            $script:Inventory = @{}
            $script:InventoryLoaded = $false
            $script:HwById = @{}
            $script:HwLastTry = @{}
            $script:HwBulkDone = $false
            $script:CloneBodyBySource = @{}
            $script:DeletingIds = @{}
        }
        $script:VmeSession = $session
        $script:LastBaseUrl = $baseUrl

        $who = Invoke-VmeApi -Path '/api/whoami'
        $whoName = [string](Get-JsonValue (Get-JsonValue $who 'user') 'username')
        if ([string]::IsNullOrWhiteSpace($whoName)) { $whoName = 'unknown' }
        Write-DebugLog "Connected to VME at [$baseUrl] as [$whoName] ($(if ($username) { 'password grant' } else { 'pre-issued token' }), ssl_check=$($script:SslCheck))." -Level 'I' -Component '02'
        if ($script:TemplateMethod -eq 'versioning') {
            # Linked clones need VME 9.1+; on 9.0.x the linked-clone call fails with a bare
            # HTTP 400. Logged, not enforced.
            try {
                $ping = Invoke-VmeApi -Path '/api/ping' -FailureLevel 'W' -StatKey 'GET /api/ping'
                $build = [string](Get-JsonValue $ping 'buildVersion')
                if ([string]::IsNullOrWhiteSpace($build)) { Write-DebugLog "template_method=versioning: VME did not report its build version (GET /api/ping) -- linked clones need VME 9.1 or later." -Level 'W' -Component '02' }
                elseif ($build -match '^(\d+)\.(\d+)' -and ([int]$Matches[1] -lt 9 -or ([int]$Matches[1] -eq 9 -and [int]$Matches[2] -lt 1))) { Write-DebugLog "template_method=versioning but VME reports build [$build] -- linked clones need VME 9.1 or later; guests/snapshots/create will fail on this appliance." -Level 'W' -Component '02' }
                else { Write-DebugLog "VME build [$build] (template_method=versioning, linked clones)." -Level 'I' -Component '02' }
            }
            catch { }
        }
        return @{ result = @{ connected = $true; user = $whoName } }
    }
    catch {
        $script:VmeSession = $null
        Write-DebugLog "provider/connect failed: $($_.Exception.Message)" -Level 'E' -Component '02'
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to connect: $($_.Exception.Message)"
    }
}

# Clone state, the maintenance set and the caches are kept: RAS reconnects right away
# and a clone in flight or a template in maintenance must survive that.
function Handle-Disconnect {
    $script:VmeSession = $null
    $script:WebSession = $null
    Write-DebugLog 'Disconnected. Clone tracking, maintenance state and caches retained.' -Level 'I' -Component '02'
    return @{ result = @{ disconnected = $true } }
}

function Handle-GuestList {
    if (-not (Test-VmeConnected)) { return New-NotConnectedResponse }
    try {
        $ok = Update-VmeInventory -MaxAgeSeconds ([int]$script:Settings.inventory.list_min_refresh_seconds)
        if (-not $ok -and -not $script:InventoryLoaded) {
            return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to list guests: the instance listing failed and no earlier listing is available"
        }
        Invoke-VmeDeletingProbe
        Update-VmeHardwareBulkIfDue

        # Clones submitted but not yet resolved by tasks/get have no id, so the id check
        # below cannot hide them -- yet a listing (this call's own refresh, or the
        # resolution safety-net listing) can already hold the new instance. Listing it
        # would let RAS's sync bind it before its clone thread gets the id from tasks/get
        # Matched the way Find-VmeCloneInstance
        # accepts a candidate: exact name, not the source, not preexisting, id above
        # every id known at submit, lineage (when VME reports it) = this source.
        $pendingClones = @($script:CloneState.Values | Where-Object { [string]::IsNullOrWhiteSpace([string]$_.clone_id) })

        $ids = [System.Collections.Generic.List[string]]::new($script:Inventory.Count)
        foreach ($id in $script:Inventory.Keys) {
            $rec = $script:Inventory[$id]
            if (Test-VmeInstanceExcluded -Name $rec.name) {
                if ($script:ExcludeLogged.Add($id)) { Write-DebugLog "Instance [$id] ($($rec.name)) hidden by inventory.exclude_instance_name_patterns." -Level 'T' -Component '03' -Ref $id }
                continue
            }
            # A clone RAS hasn't been told is complete must not be listed, or RAS's sync
            # can bind it before its own clone thread gets the id.
            $key = Get-CloneKeyByCloneId -VmId $id
            if ($null -ne $key -and $null -eq $script:CloneState[$key].reported_completed_epoch) { continue }
            if ($pendingClones.Count -gt 0 -and $null -eq $key) {
                $isPendingTarget = $false
                foreach ($pc in $pendingClones) {
                    if ($pc.name -ne $rec.name -or $id -eq [string]$pc.source_id -or (@($pc.preexisting_ids) -contains $id)) { continue }
                    [long]$idNum = 0
                    if ([long]::TryParse($id, [ref]$idNum) -and $idNum -le [long]$pc.max_known_id) { continue }
                    if (-not [string]::IsNullOrWhiteSpace($rec.cloned_from) -and $rec.cloned_from -ne [string]$pc.source_id) { continue }
                    $isPendingTarget = $true
                    break
                }
                if ($isPendingTarget) {
                    Write-DebugLog "Instance [$id] ($($rec.name)) looks like a clone tasks/get has not resolved yet -- not listed until it is reported." -Level 'T' -Component '03' -Ref $id
                    continue
                }
            }
            $ids.Add($id)
        }
        Invoke-OrphanAudit
        return @{ result = @{ guests = $ids.ToArray() } }
    }
    catch {
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to list guests: $($_.Exception.Message)"
    }
}

# RAS sends guests/get single-id, one after another, for every listed guest: at
# 3000 guests the provider's per-call overhead IS the sweep time. The steady-state path
# makes no API call and writes one D line.
# ---------------------------------------------------------------------------
# Watched instances (inventory.power_action_hold_seconds, watch_refresh_interval_seconds)
#
# The listing is the only routine source of power state, and VME can show a change late
# (a desktop this provider stopped can still be listed as running 40 s later; RAS then
# sees it power back on and waits its full 60 s before the delete). Two kinds of
# instance are therefore re-read on their own guests/get, at most once per interval:
#   - one with a start/stop/suspend this provider issued, until VME shows the new state
#     or the hold ends -- meanwhile a listing still showing the old state is reported as
#     the transition (powering_off / powering_on / suspending), never as a reversal;
#   - a running template in maintenance: RAS shuts it down through its guest agent on
#     Exit Maintenance, so only VME can tell the provider (through the listing alone that
#     can take two minutes).
# Both are few at a time; every other guests/get stays free of API calls.
# ---------------------------------------------------------------------------

function Set-VmePowerWatch {
    param([string]$VmId, [string]$Action, [hashtable]$Rec)
    $script:PowerWatch.Remove($VmId)
    if ([int]$script:Settings.inventory.power_action_hold_seconds -le 0 -or $null -eq $Rec) { return }
    $ends = @{ start = @('powered_on', 'powering_on'); stop = @('powered_off', 'powering_off'); suspend = @('suspended', 'suspending') }
    $e = $ends[$Action]
    # restart has no end state that differs from the start state.
    if ($null -eq $e -or $Rec.ras_state -eq $e[0]) { return }
    $script:PowerWatch[$VmId] = @{ action = $Action; from = [string]$Rec.ras_state; target = $e[0]; transitional = $e[1]; since = Get-EpochNow; held = $false }
}

# True when guests/get should re-read this instance now (and stamps the attempt).
function Test-VmeWatchRereadDue {
    param([string]$VmId, [hashtable]$Rec)
    $watched = $script:PowerWatch.ContainsKey($VmId) -or ((Test-VmeInMaintenanceMode -VmId $VmId) -and $Rec.ras_state -in @('powered_on', 'powering_on'))
    if (-not $watched) { return $false }
    $interval = [int]$script:Settings.inventory.watch_refresh_interval_seconds
    if (([DateTime]::UtcNow - $Rec.fetched_at).TotalSeconds -lt $interval) { return $false }
    $last = $script:WatchReadAt[$VmId]
    if ($null -ne $last -and ([DateTime]::UtcNow - $last).TotalSeconds -lt $interval) { return $false }
    $script:WatchReadAt[$VmId] = [DateTime]::UtcNow
    return $true
}

# The state to report instead of the record's, or $null. Ends the watch once VME shows
# the target state or the hold runs out.
function Get-VmePowerWatchOverride {
    param([string]$VmId, [hashtable]$Rec)
    $w = $script:PowerWatch[$VmId]
    if ($null -eq $w) { return $null }
    $age = (Get-EpochNow) - [long]$w.since
    if ($Rec.ras_state -eq $w.target) {
        Write-DebugLog "Power watch: instance [$VmId] is [$($w.target)] ${age}s after [$($w.action)]." -Level 'T' -Component '05' -Ref $VmId
        $script:PowerWatch.Remove($VmId)
        return $null
    }
    if ($age -gt [int]$script:Settings.inventory.power_action_hold_seconds) {
        Write-DebugLog "Power watch: [$($w.action)] of instance [$VmId] not shown by VME after ${age}s (status [$($Rec.status)]) -- reporting VME's state again." -Level 'W' -Component '05' -Ref $VmId
        $script:PowerWatch.Remove($VmId)
        return $null
    }
    if ($Rec.ras_state -eq $w.from) {
        if (-not $w.held) {
            $w.held = $true
            Write-DebugLog "Power watch: VME still shows instance [$VmId] as [$($Rec.status)] ${age}s after [$($w.action)] -- reporting [$($w.transitional)] until it changes." -Level 'T' -Component '05' -Ref $VmId
        }
        return $w.transitional
    }
    return $null
}

function Handle-GuestGet {
    param([System.Collections.IDictionary]$Params)

    $idParam = $Params['id']
    if ($idParam -is [System.Array]) {
        if ($idParam.Count -ne 1) { return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Batched guests/get is not supported; send one id per call" }
        $idParam = $idParam[0]
    }
    $vmId = [string]$idParam
    if ([string]::IsNullOrWhiteSpace($vmId)) { return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Invalid guest id" }
    if (-not (Test-VmeConnected)) { return New-NotConnectedResponse }

    $key = Get-CloneKeyByCloneId -VmId $vmId
    try {
        if ($null -ne $key) {
            $guest = Get-RasGuestObjectForTrackedClone -VmId $vmId -Key $key
        }
        else {
            $rec = $script:Inventory[$vmId]
            $reread = ($null -eq $rec -or $rec.dirty)
            $watchDue = (-not $reread) -and (Test-VmeWatchRereadDue -VmId $vmId -Rec $rec)
            if ($reread -or $watchDue) {
                $answered = $true
                $fresh = $null
                try { $fresh = Get-VmeInstanceRecord -VmId $vmId -FailureLevel $(if ($watchDue) { 'W' } else { 'E' }) }
                catch {
                    # A watch re-read is an extra: on failure the inventory record serves.
                    if (-not $watchDue) { throw }
                    $answered = $false
                }
                if ($answered -and $null -eq $fresh) {
                    if ($script:DeletingIds.ContainsKey($vmId)) { $script:DeletingIds.Remove($vmId) }
                    Clear-VmeTrackingForVm -VmId $vmId -Gone
                    return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Guest [$vmId] not found"
                }
                if ($null -ne $fresh) { $rec = $fresh }
            }
            Update-VmeHardwareForRecord -Rec $rec
            $guest = ConvertTo-RasGuestObject -VmId $vmId -Rec $rec
            $override = Get-VmePowerWatchOverride -VmId $vmId -Rec $rec
            if ($null -ne $override -and -not $script:DeletingIds.ContainsKey($vmId)) { $guest.state = $override }
        }
        Invoke-TrackedCloneSweep -ExceptVmId $vmId
        return @{ result = $guest }
    }
    catch {
        # An in-flight clone must never surface as a hard error on a transient failure:
        # report it as still coming up, under its real name (a placeholder name makes
        # RAS file it as an unrelated guest).
        if ($null -ne $key -and (Get-VmeErrorStatus $_) -ne 404) {
            $entry = $script:CloneState[$key]
            $name = if ($null -ne $entry -and $entry.name) { $entry.name } else { "VM-$vmId" }
            Write-DebugLog "guests/get for tracked clone [$vmId] failed ($($_.Exception.Message)) -- reporting powering_on." -Level 'W' -Component '03' -Ref $vmId
            return @{ result = @{ id = $vmId; name = $name; provider = 'VME'; node = 'vme'; state = 'powering_on'; power_state = 'provisioning'; host_os = 'unknown'; ip = $null; ip_addresses = @(); mac_addresses = @(); is_template = $false; type = 'Virtual Machine' } }
        }
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to get guest [$vmId]: $($_.Exception.Message)"
    }
}

function Get-ControlAction {
    param([string]$Control)
    switch (([string]$Control).Trim().ToLowerInvariant()) {
        'start' { return 'start' }
        'resume' { return 'start' }      # VME has no resume endpoint; start resumes
        'stop' { return 'stop' }
        'shutdown' { return 'stop' }
        'reset' { return 'restart' }
        'restart' { return 'restart' }
        'reboot' { return 'restart' }
        'suspend' { return 'suspend' }
        'delete' { return 'delete' }
        default { return $null }
    }
}

# PUT path for a power action. A stop is VME's hard power-off unless power.force_stop is
# off: RAS shuts guests down cleanly through its own agent, so a stop it sends is a
# power-off, and VME's graceful shutdown takes tens of seconds.
function Get-VmePowerActionPath {
    param([string]$VmId, [string]$Action)
    if ($Action -eq 'stop' -and $script:Settings.power.force_stop) { return "/api/instances/$VmId/stop?force=true" }
    return "/api/instances/$VmId/$Action"
}

function Invoke-VmeGuestDelete {
    param([string]$VmId)

    # "Delete Template" in the RAS Console arrives as guests/control(delete) on the
    # template itself. It means "stop managing this as a template", never "destroy the
    # VM": the designation
    # is cleared, the VM stays listed as an ordinary guest and can be made a template
    # again. The id is remembered (persisted) so a RAS retry of the delete -- at any
    # later time -- is also answered without touching the VM. Only VME itself (an admin
    # in the VME console) deletes a former template.
    if ($script:FormerTemplates.ContainsKey($VmId)) {
        Write-DebugLog "guests/control(delete) for former template [$VmId] refused again -- this provider never deletes a VM that was a RAS template; delete it in VME if it must go." -Level 'W' -Component '06' -Ref $VmId
        return @{ result = @{ id = $VmId; action = 'delete'; message = "$($script:ProviderNamePrefix) Guest was a template -- template designation cleared, underlying VM preserved" } }
    }
    if ($script:BasicTemplates.ContainsKey($VmId)) {
        $script:FormerTemplates[$VmId] = [string]$script:BasicTemplates[$VmId]
        Save-FormerTemplateState
        $script:BasicTemplates.Remove($VmId)
        Save-BasicTemplateState
        if ($script:MaintenanceIds.ContainsKey($VmId)) { $script:MaintenanceIds.Remove($VmId); Save-MaintenanceState }
        $script:CloneBodyBySource.Remove($VmId)
        Write-DebugLog "guests/control(delete) for instance [$VmId] -- it is a RAS template: designation cleared, VM left untouched and listed as an ordinary guest (it can be made a template again). This provider will never delete it." -Level 'W' -Component '06' -Ref $VmId
        $tv = $script:TemplateVersions[$VmId]
        if ($null -ne $tv -and $tv.versions.Count -gt 0) {
            Write-DebugLog "Template [$VmId]'s version(s) $(@($tv.versions.Values | ForEach-Object { "$($_.ras_name) (snapshot [$($_.snapshot_id)], image [$($_.image_id)])" }) -join ', ') are kept on VME: linked clones may still use them. Delete the snapshots in VME once no desktop needs them." -Level 'W' -Component '06' -Ref $VmId
        }
        return @{ result = @{ id = $VmId; action = 'delete'; message = "$($script:ProviderNamePrefix) Guest was a template -- template designation cleared, underlying VM preserved" } }
    }

    # Deleting the source of a clone still in VME's Deploy step pulls the disk out from
    # under it.
    $inFlight = @(foreach ($e in $script:CloneState.Values) { if ($e.source_id -eq $VmId -and -not $e.provision_confirmed) { $e.name } })
    if ($inFlight.Count -gt 0) {
        Write-DebugLog "guests/control(delete) for instance [$VmId] refused -- it is the source of $($inFlight.Count) clone(s) still provisioning: $($inFlight -join ', ')." -Level 'E' -Component '05' -Ref $VmId
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Guest [$VmId] is the source of $($inFlight.Count) clone(s) still provisioning; delete refused"
    }

    if ($script:DeletingIds.ContainsKey($VmId)) {
        return @{ result = @{ id = $VmId; action = 'delete'; message = "$($script:ProviderNamePrefix) Guest delete already submitted" } }
    }

    $rec = $script:Inventory[$VmId]
    $name = if ($null -ne $rec) { $rec.name } else { '' }

    # MAC preservation: the MAC comes from the hardware cache (free). Only on a cache
    # miss is the server object fetched, once.
    if ($script:Settings.cloning.mac_preservation.enabled -and $null -ne $rec -and -not [string]::IsNullOrWhiteSpace($name)) {
        $hw = $script:HwById[$VmId]
        if ($null -eq $hw -or -not $hw.complete) { Update-VmeHardwareForRecord -Rec $rec -Force; $hw = $script:HwById[$VmId] }
        if ($null -ne $hw -and $hw.complete) {
            $script:RecentlyDeletedMacByName[$name] = @{ mac = $hw.macs[0]; deleted_at = [DateTime]::UtcNow }
            Write-DebugLog "MAC preservation: captured [$($hw.macs[0])] of instance [$VmId] ($name) for a recreation under the same name." -Level 'T' -Component '05' -Ref $VmId
        }
        else { Write-DebugLog "MAC preservation: no MAC known for instance [$VmId] ($name) -- a recreation will get a new MAC." -Level 'W' -Component '05' -Ref $VmId }
    }

    if ($script:Settings.delete.stop_before_delete -and $null -ne $rec -and $rec.ras_state -eq 'powered_on') {
        try { [void](Invoke-VmeApi -Path (Get-VmePowerActionPath -VmId $VmId -Action 'stop') -Method PUT -Body @{} -Idempotent -Ref $VmId -FailureLevel 'W') }
        catch { Write-DebugLog "Pre-delete stop of instance [$VmId] failed (deleting anyway): $($_.Exception.Message)" -Level 'W' -Component '05' -Ref $VmId }
    }

    try {
        $resp = Invoke-VmeApi -Path "/api/instances/${VmId}?$($script:DeleteQuery)" -Method DELETE -Idempotent -AllowNotFound -Ref $VmId
        if ($null -eq $resp) {
            Write-DebugLog "Delete of instance [$VmId]: VME returns 404 -- already gone." -Level 'I' -Component '05' -Ref $VmId
            Clear-VmeTrackingForVm -VmId $VmId -Gone
            return @{ result = @{ id = $VmId; action = 'delete'; message = "$($script:ProviderNamePrefix) Guest already deleted" } }
        }
        if ($false -eq (Get-JsonValue $resp 'success')) {
            $msg = [string](Get-JsonValue $resp 'msg')
            return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) VME refused to delete guest [$VmId]: $msg"
        }
    }
    catch {
        if (-not (Test-VmeNoResponse $_)) {
            return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to delete guest [$VmId]: $($_.Exception.Message)"
        }
        # No response: VME may well have accepted it. The probe finds out either way.
        Write-DebugLog "Delete of instance [$VmId] got no response -- treated as submitted; the delete probe will confirm." -Level 'W' -Component '05' -Ref $VmId
    }

    # -Gone: an admin-initiated delete is the cancel signal for a slow/in-flight clone
    # (RAS Console has no "abort clone" action, so a delete on the guest it produced is
    # how an admin stops a stuck bulk operation) -- this guest must stop being reportable
    # as a valid, completed clone the instant the delete is submitted, not up to
    # completed_clone_task_output_retention_seconds later.
    Clear-VmeTrackingForVm -VmId $VmId -KeepRecord -Gone
    # Probed from the next guests/list on; listed as powering_off until VME returns 404.
    $script:DeletingIds[$VmId] = @{ name = $name; submitted_at = [DateTime]::UtcNow; last_probe = [DateTime]::UtcNow; stuck_logged = $false }
    Write-DebugLog "Delete submitted for instance [$VmId] ($name); listed as powering_off until VME confirms it gone." -Level 'I' -Component '05' -Ref $VmId
    return @{ result = @{ id = $VmId; action = 'delete'; message = "$($script:ProviderNamePrefix) Guest delete submitted" } }
}

function Handle-GuestControl {
    param([System.Collections.IDictionary]$Params)

    $vmId = [string]$Params['id']
    $action = Get-ControlAction -Control ([string]$Params['control'])
    if ($null -eq $action) { return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Unsupported control action [$($Params['control'])]" }
    if ([string]::IsNullOrWhiteSpace($vmId)) { return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Invalid guest id" }
    if (-not (Test-VmeConnected)) { return New-NotConnectedResponse }

    try {
        if ($action -eq 'delete') { return (Invoke-VmeGuestDelete -VmId $vmId) }

        if ($script:DeletingIds.ContainsKey($vmId)) {
            return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Guest [$vmId] is being deleted"
        }
        $rec = $script:Inventory[$vmId]
        # A clone still provisioning starts itself (provisionPoweredOff defaults to
        # false); a start issued during Deploy may be refused.
        if ($action -eq 'start' -and $null -ne $rec -and (Test-VmeInstanceProvisioning -Rec $rec)) {
            Write-DebugLog "guests/control(start) for instance [$vmId] while VME is still provisioning it (status=$($rec.status)) -- nothing to do, VME starts it." -Level 'T' -Component '05' -Ref $vmId
            return @{ result = @{ id = $vmId; action = $action; message = "$($script:ProviderNamePrefix) Guest is still provisioning; it starts on its own" } }
        }

        [void](Invoke-VmeApi -Path (Get-VmePowerActionPath -VmId $vmId -Action $action) -Method PUT -Body @{} -Idempotent -Ref $vmId)
        Set-VmePowerWatch -VmId $vmId -Action $action -Rec $rec
        if ($null -ne $rec) { $rec.dirty = $true }

        if ($action -eq 'start' -and $script:Settings.cloning.template_flapping_workaround_enabled -and (Test-VmeIsBasicTemplate -VmId $vmId) -and
            -not (Test-VmeInMaintenanceMode -VmId $vmId) -and $script:TemplateFlappingUnmaskedVmIds.Add($vmId)) {
            Write-DebugLog "Template flapping workaround: instance [$vmId] explicitly started by RAS -- mask lifted for this id." -Level 'I' -Component '06' -Ref $vmId
        }
        Write-DebugLog "Issued [$action] for instance [$vmId]." -Level 'I' -Component '05' -Ref $vmId
        return @{ result = @{ id = $vmId; action = $action; message = "$($script:ProviderNamePrefix) Guest control [$action] submitted successfully" } }
    }
    catch {
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to control guest [$vmId]: $($_.Exception.Message)"
    }
}

# VME has no template flag and no convert API: RAS's template contract is kept locally.
#   convert(true)  -> the instance is a template (basic set); leaves maintenance.
#   convert(false) -> maintenance: is_template reports false (RAS then starts the VM,
#                     an admin patches it, RAS converts it back). The basic-set entry
#                     stays, so a "Delete Template" (delete on the id) is still refused.
function Handle-GuestConvert {
    param([System.Collections.IDictionary]$Params)

    $vmId = [string]$Params['id']
    if ([string]::IsNullOrWhiteSpace($vmId)) { return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Invalid guest id" }
    if ($script:TemplateMethod -eq 'none') { return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Templates are disabled (capabilities.template_method = none)" }
    $isTemplate = $false
    try { $isTemplate = ConvertTo-CoercedSettingBool -Value $Params['is_template'] -Default $false }
    catch { return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) is_template must be a boolean" }

    $uuid = Get-RecordUuid -VmId $vmId
    if ($isTemplate) {
        if ($script:MaintenanceIds.ContainsKey($vmId)) { $script:MaintenanceIds.Remove($vmId); Save-MaintenanceState }
        if (-not $script:BasicTemplates.ContainsKey($vmId)) { $script:BasicTemplates[$vmId] = $uuid; Save-BasicTemplateState }
        # A former template made a template again: the template check protects it now.
        if ($script:FormerTemplates.ContainsKey($vmId)) { $script:FormerTemplates.Remove($vmId); Save-FormerTemplateState }
        # The template may have been changed (disk grown) during maintenance: the next
        # clone re-reads its body inputs.
        $script:CloneBodyBySource.Remove($vmId)
        if ($script:TemplateFlappingUnmaskedVmIds.Remove($vmId)) { Write-DebugLog "Template flapping workaround: instance [$vmId] is a template again -- mask re-armed." -Level 'I' -Component '06' -Ref $vmId }
        Write-DebugLog "Instance [$vmId] is a template (is_template=true)." -Level 'I' -Component '06' -Ref $vmId
    }
    else {
        if (-not $script:MaintenanceIds.ContainsKey($vmId)) { $script:MaintenanceIds[$vmId] = $uuid; Save-MaintenanceState }
        Write-DebugLog "Instance [$vmId] entered maintenance (is_template=false until converted back)." -Level 'I' -Component '06' -Ref $vmId
    }
    return @{ result = @{ task_id = "noop:convert:$vmId" } }
}

# ---------------------------------------------------------------------------
# Template versions (template_method 'versioning', VME 9.1+) -- docs/LINKED-CLONES.md
#
# RAS names a version RAS_TEMPLATE_VERSION_<n> on every template. The VME snapshot is
# named <RAS name>_<template id>, and the linked-clone virtual image VME makes from it
# takes the same name, so looking the image up by name is unique on the appliance (no
# image field names its source snapshot or instance). guests/snapshots/create makes
# both, in one RAS task:
#   PUT /api/instances/{id}/snapshot -> poll its process -> snapshot id + externalId
#   PUT /api/instances/{id}/linked-clone/{snapshotId}  (synchronous, no process)
#   GET /api/virtual-images?filterType=Synced&name=, checked by
#       locations[].externalDiskId = the snapshot's externalId
# Deleting the snapshot deletes its image on VME, so a version is
# deleted with one DELETE /api/snapshots/{id}. A versioned guests/clone is
# POST /api/instances with config.imageId = the version's image.
# ---------------------------------------------------------------------------

function Get-VmeVersionVmeName {
    param([string]$TemplateId, [string]$RasName)
    return "$($RasName)_$TemplateId"
}

function New-VmeVersionRecord {
    param([string]$RasName, [string]$VmeName)
    $v = @{}
    foreach ($f in $script:VersionFields) { $v[$f] = '' }
    $v.ras_name = $RasName
    $v.vme_name = $VmeName
    $v.clone_inputs = $null
    $v.linked_requested = $false
    $v.linked_epoch = $null
    $v.started_epoch = Get-EpochNow
    $v.created_epoch = $null
    return $v
}

function Get-VmeInstanceSnapshots {
    param([Parameter(Mandatory = $true)][string]$VmId, [string]$Ref = $null)
    $resp = Invoke-VmeApi -Path "/api/instances/$VmId/snapshots" -Ref $(if ($Ref) { $Ref } else { $VmId }) -FailureLevel 'W' -StatKey 'GET /api/instances/{id}/snapshots'
    $list = [System.Collections.Generic.List[object]]::new()
    foreach ($s in @(Get-JsonValue $resp 'snapshots')) { if ($null -ne $s) { $list.Add($s) } }
    return , $list
}

# The newest snapshot of the instance carrying exactly this name, or $null.
function Find-VmeInstanceSnapshot {
    param([Parameter(Mandatory = $true)][string]$VmId, [Parameter(Mandatory = $true)][string]$Name, [string]$Ref = $null)
    $best = $null
    foreach ($s in (Get-VmeInstanceSnapshots -VmId $VmId -Ref $Ref)) {
        if ([string](Get-JsonValue $s 'name') -ne $Name) { continue }
        if ($null -eq $best -or [long](Get-JsonValue $s 'id') -gt [long](Get-JsonValue $best 'id')) { $best = $s }
    }
    return $best
}

# A version's linked-clone image: exact name first, then the id-level link to its
# snapshot. Returns $null (none yet), @{ image; verified } or @{ ambiguous = <reason> }.
function Find-VmeVersionImage {
    param([Parameter(Mandatory = $true)][string]$Name, [string]$SnapshotExternalId = '', [string]$Ref = $null)
    $ft = switch ([string]$script:Settings.linked_clones.image_filter_type) { 'all' { 'All' } 'user' { 'User' } 'system' { 'System' } default { 'Synced' } }
    $resp = Invoke-VmeApi -Path "/api/virtual-images?filterType=$ft&name=$([uri]::EscapeDataString($Name))&max=50" -Ref $Ref -FailureLevel 'W' -StatKey 'GET /api/virtual-images (name)'
    $cands = @(@(Get-JsonValue $resp 'virtualImages') | Where-Object { $null -ne $_ -and [string](Get-JsonValue $_ 'name') -eq $Name })
    if ($cands.Count -eq 0) { return $null }
    $linked = @()
    if ($SnapshotExternalId) {
        $linked = @($cands | Where-Object { @(foreach ($l in @(Get-JsonValue $_ 'locations')) { [string](Get-JsonValue $l 'externalDiskId') }) -contains $SnapshotExternalId })
    }
    if ($cands.Count -eq 1) { return @{ image = $cands[0]; verified = ($linked.Count -eq 1) } }
    if ($linked.Count -eq 1) { return @{ image = $linked[0]; verified = $true } }
    return @{ ambiguous = "$($cands.Count) virtual images are named [$Name] ($(($cands | ForEach-Object { [string](Get-JsonValue $_ 'id') }) -join ', '))$(if ($SnapshotExternalId) { ", $($linked.Count) of them on snapshot [$SnapshotExternalId]" }) -- cannot tell which one is this version's" }
}

function Get-VmeImageDiskIds {
    param([object]$Image)
    return (@(foreach ($l in @(Get-JsonValue $Image 'locations')) { [string](Get-JsonValue $l 'externalDiskId') }) -join ',')
}

# Best-effort: labels let the version map be rebuilt from VME.
function Set-VmeVersionImageLabels {
    param([string]$ImageId, [string]$TemplateId, [string]$RasName, [string]$Ref)
    if (-not $script:Settings.linked_clones.label_images) { return }
    $n = if ($RasName -match '(\d+)$') { $Matches[1] } else { $RasName }
    $labels = @("ras-template-$TemplateId", "ras-version-$n")
    try {
        $resp = Invoke-VmeApi -Path "/api/virtual-images/$ImageId" -Method PUT -Body @{ virtualImage = @{ labels = $labels } } -Idempotent -Ref $Ref -FailureLevel 'W'
        if ($false -eq (Get-JsonValue $resp 'success')) { Write-DebugLog "Labels on image [$ImageId] not set: VME says $([string](Get-JsonValue $resp 'msg'))" -Level 'W' -Component '07' -Ref $Ref }
        else { Write-DebugLog "Image [$ImageId] labelled $($labels -join ', ')." -Level 'T' -Component '07' -Ref $Ref }
    }
    catch { Write-DebugLog "Labels on image [$ImageId] not set (best-effort): $($_.Exception.Message)" -Level 'W' -Component '07' -Ref $Ref }
}

# A ready version guests/clone can use: from the state file, else rebuilt from VME by
# name (a lost or reset state file must not strand a pool). $null when VME has none.
function Resolve-VmeTemplateVersion {
    param([string]$TemplateId, [string]$RasName, [string]$Ref)
    $tv = $script:TemplateVersions[$TemplateId]
    $v = if ($null -ne $tv) { $tv.versions[$RasName] } else { $null }
    if ($null -ne $v -and $v.status -eq 'ready' -and -not [string]::IsNullOrWhiteSpace([string]$v.image_id)) { return $v }
    $vmeName = Get-VmeVersionVmeName -TemplateId $TemplateId -RasName $RasName
    $snap = Find-VmeInstanceSnapshot -VmId $TemplateId -Name $vmeName -Ref $Ref
    if ($null -eq $snap) { return $null }
    $found = Find-VmeVersionImage -Name $vmeName -SnapshotExternalId ([string](Get-JsonValue $snap 'externalId')) -Ref $Ref
    if ($null -eq $found) { return $null }
    if ($found.ContainsKey('ambiguous')) { throw "$($script:ProviderNamePrefix) $($found.ambiguous)" }
    $tv = Get-VmeTemplateVersionSet -TemplateId $TemplateId
    $v = New-VmeVersionRecord -RasName $RasName -VmeName $vmeName
    $v.status = 'ready'
    $v.snapshot_id = [string](Get-JsonValue $snap 'id')
    $v.snapshot_external_id = [string](Get-JsonValue $snap 'externalId')
    $v.image_id = [string](Get-JsonValue $found.image 'id')
    $v.image_name = [string](Get-JsonValue $found.image 'name')
    $v.created_epoch = Get-EpochNow
    $tv.versions[$RasName] = $v
    Save-TemplateVersions
    Write-DebugLog "Version [$RasName] of template [$TemplateId] rebuilt from VME: snapshot [$($v.snapshot_id)], image [$($v.image_id)] (it was missing from the provider state)." -Level 'I' -Component '07' -Ref $TemplateId
    return $v
}

# Common validation of the four guests/snapshots/* requests.
function Get-VmeSnapshotRequest {
    param([System.Collections.IDictionary]$Params, [string]$Method)
    $vmId = [string]$Params['id']
    $name = ([string]$Params['name']).Trim()
    if ($script:TemplateMethod -ne 'versioning') {
        return @{ error = (New-ErrorResponse -Code $script:ErrorCodes.MethodNotFound -Message "$($script:ProviderNamePrefix) [$Method] is only supported with capabilities.template_method = versioning (this provider runs [$($script:TemplateMethod)])") }
    }
    if ([string]::IsNullOrWhiteSpace($vmId) -or [string]::IsNullOrWhiteSpace($name)) { return @{ error = (New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Both 'id' and 'name' are required") } }
    if (-not (Test-VmeConnected)) { return @{ error = (New-NotConnectedResponse) } }
    return @{ vm_id = $vmId; ras_name = $name; vme_name = (Get-VmeVersionVmeName -TemplateId $vmId -RasName $name) }
}

function New-VmeSnapshotTaskId {
    param([string]$Kind, [string]$VmId)
    return "$($Kind):$($VmId):$(Get-EpochNow)$('{0:x4}' -f (Get-Random -Maximum 65536))"
}

function Handle-SnapshotCreate {
    param([System.Collections.IDictionary]$Params)
    $q = Get-VmeSnapshotRequest -Params $Params -Method 'guests/snapshots/create'
    if ($q.ContainsKey('error')) { return $q.error }
    $vmId = $q.vm_id; $rasName = $q.ras_name; $vmeName = $q.vme_name

    try {
        $tv = Get-VmeTemplateVersionSet -TemplateId $vmId
        $v = $tv.versions[$rasName]
        if ($null -ne $v -and $v.status -eq 'creating' -and $v.task_id) {
            Write-DebugLog "guests/snapshots/create [$rasName] for template [$vmId] while it is still being created -- returning its task [$($v.task_id)]." -Level 'W' -Component '07' -Ref $vmId
            return @{ result = @{ task_id = $v.task_id } }
        }
        if ($null -ne $v -and $v.status -eq 'ready') {
            Write-DebugLog "guests/snapshots/create [$rasName] for template [$vmId]: the version already exists (snapshot [$($v.snapshot_id)], image [$($v.image_id)]) -- nothing to do." -Level 'W' -Component '07' -Ref $vmId
            return @{ result = @{ task_id = "noop:snapcreate:$vmId" } }
        }
        if ($null -ne $v -and $v.status -eq 'deleting') {
            return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Version [$rasName] of template [$vmId] is being deleted; try again once that finishes"
        }

        $v = New-VmeVersionRecord -RasName $rasName -VmeName $vmeName
        $v.status = 'creating'
        $v.phase = 'snapshot'
        $v.task_id = New-VmeSnapshotTaskId -Kind 'snapcreate' -VmId $vmId

        # A snapshot left by an earlier attempt (failed task, lost response, restart) is
        # reused: VME does not enforce unique snapshot names, and a second one would
        # make the image lookup ambiguous.
        $snap = Find-VmeInstanceSnapshot -VmId $vmId -Name $vmeName -Ref $vmId
        if ($null -ne $snap) {
            $v.snapshot_id = [string](Get-JsonValue $snap 'id')
            $v.snapshot_external_id = [string](Get-JsonValue $snap 'externalId')
            Write-DebugLog "Version [$rasName] of template [$vmId]: reusing VME snapshot [$($v.snapshot_id)] named [$vmeName] left by an earlier attempt." -Level 'W' -Component '07' -Ref $v.task_id
        }
        else {
            $body = @{ snapshot = [ordered]@{ name = $vmeName; description = "Parallels RAS template version $rasName of instance $vmId"; memorySnapshot = $false; forExport = $false } }
            try {
                # Never retried: a retry would make a second snapshot.
                $resp = Invoke-VmeApi -Path "/api/instances/$vmId/snapshot" -Method PUT -Body $body -Ref $vmId
                if ($false -eq (Get-JsonValue $resp 'success')) {
                    return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) VME refused the snapshot [$vmeName] of guest [$vmId]: $([string](Get-JsonValue $resp 'msg'))"
                }
                $pids = @(@(Get-JsonValue $resp 'processIds') | Where-Object { $null -ne $_ })
                if ($pids.Count -gt 0) { $v.process_id = [string]$pids[0] }
            }
            catch {
                if (-not (Test-VmeNoResponse $_)) { throw }
                Write-DebugLog "Snapshot request [$vmeName] for template [$vmId] got no response -- tracking the snapshot by name." -Level 'W' -Component '07' -Ref $v.task_id
            }
        }

        $tv.versions[$rasName] = $v
        Save-TemplateVersions
        $rec = $script:Inventory[$vmId]
        if ($null -ne $rec) { $rec.dirty = $true }
        # The version's clone settings, as the template is now (two GETs, once per
        # version). Best-effort: missing ones are taken at the version's first clone.
        try {
            $v.clone_inputs = Get-VmeVersionCloneInputs -TemplateId $vmId -Ref $v.task_id
            if ($null -ne $v.clone_inputs) {
                Save-TemplateVersions
                Write-DebugLog "Version [$rasName] of template [$vmId] clone settings: $(Format-VmeCloneInputs -Inputs $v.clone_inputs)." -Level 'T' -Component '07' -Ref $v.task_id
            }
        }
        catch { Write-DebugLog "Version [$rasName] of template [$vmId]: clone settings not read ($($_.Exception.Message)) -- they are taken at the version's first clone." -Level 'W' -Component '07' -Ref $v.task_id }
        Write-DebugLog "Version [$rasName] of template [$vmId]: snapshot [$vmeName] $(if ($v.snapshot_id) { "[$($v.snapshot_id)] reused" } else { "requested$(if ($v.process_id) { " (process [$($v.process_id)])" })" }); the linked-clone image follows. Task [$($v.task_id)]." -Level 'I' -Component '07' -Ref $v.task_id
        return @{ result = @{ task_id = $v.task_id } }
    }
    catch {
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to create snapshot [$rasName] of guest [$vmId]: $($_.Exception.Message)"
    }
}

# tasks/get for snapcreate:*. Runs as far as it can on each poll: snapshot finished ->
# linked clone -> image found -> labelled -> ready.
function Invoke-VmeSnapshotCreateTask {
    param([string]$TaskId)

    $tid = $null
    $v = $null
    foreach ($t in @($script:TemplateVersions.Keys)) {
        foreach ($x in $script:TemplateVersions[$t].versions.Values) { if ($x.task_id -eq $TaskId) { $tid = $t; $v = $x } }
    }
    if ($null -eq $v) { return New-FailedTaskResponse "$($script:ProviderNamePrefix) Snapshot task [$TaskId] is no longer tracked" }
    if ($v.status -eq 'ready') { return @{ result = @{ state = 'completed'; output = @{} } } }
    if ($v.status -ne 'creating') { return New-FailedTaskResponse "$($script:ProviderNamePrefix) Version [$($v.ras_name)] of template [$tid] is [$($v.status)]" }
    if (-not (Test-VmeConnected)) { return @{ result = @{ state = 'running' } } }

    $L = $script:Settings.linked_clones
    $elapsed = (Get-EpochNow) - [long]$v.started_epoch
    $tv = $script:TemplateVersions[$tid]
    # A failed version leaves the provider state; anything left on VME (snapshot, image)
    # is found by name and reused by the next create.
    $fail = {
        param([string]$Msg)
        Write-DebugLog $Msg -Level 'E' -Component '07' -Ref $TaskId
        $tv.versions.Remove($v.ras_name)
        Save-TemplateVersions
        return (New-FailedTaskResponse $Msg)
    }
    $prefix = "$($script:ProviderNamePrefix) Version [$($v.ras_name)] of template [$tid]"

    try {
        if ($v.phase -eq 'snapshot') {
            if ($v.process_id) {
                $p = Get-JsonValue (Invoke-VmeApi -Path "/api/processes/$($v.process_id)" -AllowNotFound -Ref $TaskId -FailureLevel 'W' -StatKey 'GET /api/processes/{id}') 'process'
                $ps = if ($null -ne $p) { ([string](Get-JsonValue $p 'status')).ToLowerInvariant() } else { '' }
                if ($ps -in @('failed', 'cancelled')) {
                    $why = Get-VmeProcessErrorText -Process $p
                    return (& $fail "$prefix -- the snapshot process [$($v.process_id)] ended [$ps] after ${elapsed}s$(if ($why) { ": $why" })")
                }
                if ($null -ne $p -and $ps -ne 'complete') {
                    if ($elapsed -gt [int]$L.snapshot_timeout_seconds) { return (& $fail "$prefix -- the snapshot process [$($v.process_id)] is still [$ps] after ${elapsed}s (linked_clones.snapshot_timeout_seconds)") }
                    Write-DebugLog "Version [$($v.ras_name)] of template [${tid}]: snapshot process [$($v.process_id)] [$ps] percent=$([string](Get-JsonValue $p 'percent')) (${elapsed}s)." -Level 'D' -Component '07' -Ref $TaskId
                    return @{ result = @{ state = 'running' } }
                }
                # complete, or VME no longer has the process: the snapshot itself decides.
            }
            $snap = Find-VmeInstanceSnapshot -VmId $tid -Name $v.vme_name -Ref $TaskId
            $ss = if ($null -ne $snap) { ([string](Get-JsonValue $snap 'status')).ToLowerInvariant() } else { '' }
            if ($ss -eq 'failed') { return (& $fail "$prefix -- VME snapshot [$($v.vme_name)] is [failed]") }
            if ($null -eq $snap -or ($ss -and $ss -ne 'complete')) {
                if ($elapsed -gt [int]$L.snapshot_timeout_seconds) { return (& $fail "$prefix -- snapshot [$($v.vme_name)] $(if ($null -eq $snap) { 'never appeared' } else { "still [$ss]" }) after ${elapsed}s") }
                Write-DebugLog "Version [$($v.ras_name)] of template [${tid}]: snapshot [$($v.vme_name)] $(if ($null -eq $snap) { 'not listed yet' } else { "[$ss]" }) (${elapsed}s)." -Level 'D' -Component '07' -Ref $TaskId
                return @{ result = @{ state = 'running' } }
            }
            $v.snapshot_id = [string](Get-JsonValue $snap 'id')
            $v.snapshot_external_id = [string](Get-JsonValue $snap 'externalId')
            $v.phase = 'image'
            Save-TemplateVersions
            Write-DebugLog "Version [$($v.ras_name)] of template [${tid}]: snapshot [$($v.snapshot_id)] ($($v.vme_name)) complete after ${elapsed}s, externalId [$($v.snapshot_external_id)]." -Level 'I' -Component '07' -Ref $TaskId
        }

        if ($v.phase -eq 'image') {
            # An image left by an earlier attempt is reused: a second linked clone of the
            # same snapshot would make a second image with the same name.
            $found = Find-VmeVersionImage -Name $v.vme_name -SnapshotExternalId $v.snapshot_external_id -Ref $TaskId
            if ($null -eq $found -and -not $v.linked_requested) {
                try {
                    # Not retried: the call is synchronous, a retry could make a second image.
                    $resp = Invoke-VmeApi -Path "/api/instances/$tid/linked-clone/$($v.snapshot_id)" -Method PUT -Body @{} -Ref $TaskId
                    if ($false -eq (Get-JsonValue $resp 'success')) { return (& $fail "$prefix -- VME refused the linked clone of snapshot [$($v.snapshot_id)]: $([string](Get-JsonValue $resp 'msg')) (VME 9.1 or later is required)") }
                }
                catch {
                    if (-not (Test-VmeNoResponse $_)) { return (& $fail "$prefix -- the linked clone of snapshot [$($v.snapshot_id)] failed: $($_.Exception.Message) (VME 9.1 or later is required)") }
                    Write-DebugLog "Linked clone of snapshot [$($v.snapshot_id)] got no response -- looking for the image by name." -Level 'W' -Component '07' -Ref $TaskId
                }
                $v.linked_requested = $true
                $v.linked_epoch = Get-EpochNow
                Save-TemplateVersions
                Write-DebugLog "Version [$($v.ras_name)] of template [${tid}]: linked clone of snapshot [$($v.snapshot_id)] requested." -Level 'T' -Component '07' -Ref $TaskId
                $found = Find-VmeVersionImage -Name $v.vme_name -SnapshotExternalId $v.snapshot_external_id -Ref $TaskId
            }
            if ($null -ne $found -and $found.ContainsKey('ambiguous')) { return (& $fail "$prefix -- $($found.ambiguous)") }
            if ($null -eq $found) {
                $since = (Get-EpochNow) - [long]$(if ($null -ne $v.linked_epoch) { $v.linked_epoch } else { $v.started_epoch })
                if ($since -gt [int]$L.image_wait_seconds) { return (& $fail "$prefix -- no virtual image named [$($v.vme_name)] (filterType=$($L.image_filter_type)) ${since}s after the linked clone of snapshot [$($v.snapshot_id)]") }
                Write-DebugLog "Version [$($v.ras_name)] of template [${tid}]: image [$($v.vme_name)] not listed yet (${since}s after the linked clone)." -Level 'D' -Component '07' -Ref $TaskId
                return @{ result = @{ state = 'running' } }
            }
            $img = $found.image
            $v.image_id = [string](Get-JsonValue $img 'id')
            $v.image_name = [string](Get-JsonValue $img 'name')
            if ($found.verified) { Write-DebugLog "Image [$($v.image_id)] is on snapshot [$($v.snapshot_external_id)] (locations[].externalDiskId matches the snapshot's externalId)." -Level 'T' -Component '07' -Ref $TaskId }
            elseif ($v.snapshot_external_id) { Write-DebugLog "Image [$($v.image_id)] taken by its unique name only: its locations[].externalDiskId [$(Get-VmeImageDiskIds -Image $img)] do not include the snapshot's externalId [$($v.snapshot_external_id)] (the image-to-snapshot id link is not confirmed on this appliance)." -Level 'W' -Component '07' -Ref $TaskId }
            Set-VmeVersionImageLabels -ImageId $v.image_id -TemplateId $tid -RasName $v.ras_name -Ref $TaskId
            $v.status = 'ready'
            $v.phase = ''
            $v.created_epoch = Get-EpochNow
            Save-TemplateVersions
            Write-DebugLog "Version [$($v.ras_name)] of template [$tid] ready after ${elapsed}s: snapshot [$($v.snapshot_id)], linked-clone image [$($v.image_id)] ($($v.image_name), linkedClone=$([string](Get-JsonValue $img 'linkedClone')), status=$([string](Get-JsonValue $img 'status')))." -Level 'I' -Component '07' -Ref $TaskId
            return @{ result = @{ state = 'completed'; output = @{} } }
        }
        return (& $fail "$prefix -- unknown phase [$($v.phase)]")
    }
    catch {
        if ($elapsed -gt [int]$L.snapshot_timeout_seconds) { return (& $fail "$prefix -- gave up after ${elapsed}s: $($_.Exception.Message)") }
        Write-DebugLog "tasks/get for [$TaskId] failed this poll: $($_.Exception.Message)" -Level 'W' -Component '07' -Ref $TaskId
        return @{ result = @{ state = 'running' } }
    }
}

# Bare boolean, as RAS expects. Asked on VME every time (Exit Maintenance only, rare).
function Handle-SnapshotExists {
    param([System.Collections.IDictionary]$Params)
    $q = Get-VmeSnapshotRequest -Params $Params -Method 'guests/snapshots/exists'
    if ($q.ContainsKey('error')) { return $q.error }
    $vmId = $q.vm_id; $rasName = $q.ras_name
    try {
        $snap = Find-VmeInstanceSnapshot -VmId $vmId -Name $q.vme_name -Ref $vmId
        $tv = $script:TemplateVersions[$vmId]
        $v = if ($null -ne $tv) { $tv.versions[$rasName] } else { $null }
        if ($null -eq $snap -and $null -ne $v -and $v.status -eq 'ready') {
            Write-DebugLog "Version [$rasName] of template [$vmId] is in the provider state but VME has no snapshot [$($q.vme_name)] -- reporting false and dropping it from the state." -Level 'W' -Component '07' -Ref $vmId
            $tv.versions.Remove($rasName)
            Save-TemplateVersions
        }
        Write-DebugLog "guests/snapshots/exists [$rasName] for template [$vmId]: $($null -ne $snap)$(if ($null -ne $snap) { " (snapshot [$([string](Get-JsonValue $snap 'id'))])" })." -Level 'T' -Component '07' -Ref $vmId
        return @{ result = ($null -ne $snap) }
    }
    catch {
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to check snapshot [$rasName] of guest [$vmId]: $($_.Exception.Message)"
    }
}

function Handle-SnapshotDelete {
    param([System.Collections.IDictionary]$Params)
    $q = Get-VmeSnapshotRequest -Params $Params -Method 'guests/snapshots/delete'
    if ($q.ContainsKey('error')) { return $q.error }
    $vmId = $q.vm_id; $rasName = $q.ras_name; $vmeName = $q.vme_name
    try {
        $tv = $script:TemplateVersions[$vmId]
        $v = if ($null -ne $tv) { $tv.versions[$rasName] } else { $null }
        if ($null -ne $v -and $v.status -eq 'creating') {
            return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Version [$rasName] of template [$vmId] is still being created"
        }
        $snap = Find-VmeInstanceSnapshot -VmId $vmId -Name $vmeName -Ref $vmId
        if ($null -eq $snap) {
            if ($null -ne $v) { $tv.versions.Remove($rasName); Save-TemplateVersions }
            Write-DebugLog "guests/snapshots/delete [$rasName] for template [${vmId}]: VME has no snapshot [$vmeName] -- nothing to delete." -Level 'I' -Component '07' -Ref $vmId
            return @{ result = @{ task_id = "noop:snapdelete:$vmId" } }
        }
        $snapId = [string](Get-JsonValue $snap 'id')
        $imageId = if ($null -ne $v) { [string]$v.image_id } else { '' }
        if (-not $imageId) {
            $found = Find-VmeVersionImage -Name $vmeName -SnapshotExternalId ([string](Get-JsonValue $snap 'externalId')) -Ref $vmId
            if ($null -ne $found -and -not $found.ContainsKey('ambiguous')) { $imageId = [string](Get-JsonValue $found.image 'id') }
        }

        # VME deletes the image with its snapshot, and the desktops built on it would
        # lose their base disk.
        if ($imageId -and $script:Settings.linked_clones.refuse_delete_in_use) {
            $users = [System.Collections.Generic.List[string]]::new()
            foreach ($id in $script:Inventory.Keys) {
                $r = $script:Inventory[$id]
                if ([string]$r.image_id -eq $imageId -and -not $script:DeletingIds.ContainsKey($id)) { $users.Add("$($r.name) [$id]") }
            }
            foreach ($e in $script:CloneState.Values) {
                if ([string]$e.image_id -eq $imageId -and -not ($e.clone_id -and $script:Inventory.ContainsKey([string]$e.clone_id))) { $users.Add("$($e.name) (being cloned)") }
            }
            if ($users.Count -gt 0) {
                $shown = @($users | Select-Object -First 5) -join ', '
                $msg = "$($script:ProviderNamePrefix) Version [$rasName] of template [$vmId] is still used by $($users.Count) guest(s): $shown$(if ($users.Count -gt 5) { ', ...' }). VME deletes the version's image with its snapshot, which would break them -- recreate them from a newer version or delete them first."
                Write-DebugLog $msg -Level 'E' -Component '07' -Ref $vmId
                return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message $msg
            }
        }

        $resp = Invoke-VmeApi -Path "/api/snapshots/$snapId" -Method DELETE -Idempotent -AllowNotFound -Ref $vmId
        if ($null -ne $resp -and $false -eq (Get-JsonValue $resp 'success')) {
            return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) VME refused to delete snapshot [$snapId] ($vmeName): $([string](Get-JsonValue $resp 'msg'))"
        }
        if ($null -ne $v) { $v.status = 'deleting'; Save-TemplateVersions }
        $taskId = New-VmeSnapshotTaskId -Kind 'snapdelete' -VmId $vmId
        $script:SnapshotTasks[$taskId] = @{ kind = 'delete'; template_id = $vmId; ras_name = $rasName; vme_name = $vmeName; snapshot_id = $snapId; image_id = $imageId; started = [DateTime]::UtcNow }
        Write-DebugLog "Version [$rasName] of template [${vmId}]: snapshot [$snapId] ($vmeName) delete submitted$(if ($imageId) { "; VME removes its image [$imageId] with it" }). Task [$taskId]." -Level 'I' -Component '07' -Ref $taskId
        return @{ result = @{ task_id = $taskId } }
    }
    catch {
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to delete snapshot [$rasName] of guest [$vmId]: $($_.Exception.Message)"
    }
}

# Enter Maintenance under versioning: convert(false) -> revert to the admin's version -> start.
function Handle-SnapshotRevert {
    param([System.Collections.IDictionary]$Params)
    $q = Get-VmeSnapshotRequest -Params $Params -Method 'guests/snapshots/revert'
    if ($q.ContainsKey('error')) { return $q.error }
    $vmId = $q.vm_id; $rasName = $q.ras_name; $vmeName = $q.vme_name
    try {
        $snap = Find-VmeInstanceSnapshot -VmId $vmId -Name $vmeName -Ref $vmId
        if ($null -eq $snap) { return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Version [$rasName] of template [$vmId] has no snapshot [$vmeName] on VME" }
        $snapId = [string](Get-JsonValue $snap 'id')
        # Repeating a revert before anything changed is harmless: one retry on no response.
        $resp = Invoke-VmeApi -Path "/api/instances/$vmId/revert-snapshot/$snapId" -Method PUT -Body @{} -Idempotent -Ref $vmId
        if ($false -eq (Get-JsonValue $resp 'success')) {
            return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) VME refused to revert guest [$vmId] to snapshot [$snapId] ($vmeName): $([string](Get-JsonValue $resp 'msg'))"
        }
        $rec = $script:Inventory[$vmId]
        if ($null -ne $rec) { $rec.dirty = $true }
        $pids = @(@(Get-JsonValue $resp 'processIds') | Where-Object { $null -ne $_ })
        if ($pids.Count -eq 0) {
            Write-DebugLog "Template [$vmId] reverted to version [$rasName] (snapshot [$snapId]) -- synchronous." -Level 'I' -Component '07' -Ref $vmId
            return @{ result = @{ task_id = "noop:revert:$vmId" } }
        }
        $taskId = New-VmeSnapshotTaskId -Kind 'snapproc' -VmId $vmId
        $script:SnapshotTasks[$taskId] = @{ kind = 'process'; what = "revert of template [$vmId] to version [$rasName]"; process_id = [string]$pids[0]; template_id = $vmId; started = [DateTime]::UtcNow }
        Write-DebugLog "Template [$vmId] revert to version [$rasName] (snapshot [$snapId]) submitted, process [$($pids[0])]. Task [$taskId]." -Level 'I' -Component '07' -Ref $taskId
        return @{ result = @{ task_id = $taskId } }
    }
    catch {
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to revert guest [$vmId] to snapshot [$rasName]: $($_.Exception.Message)"
    }
}

# tasks/get for snapdelete:* and snapproc:* (in memory: a provider restart answers
# "no longer tracked", and RAS's own retry of the step finds the real state on VME).
function Invoke-VmeSnapshotTask {
    param([string]$TaskId)
    $t = $script:SnapshotTasks[$TaskId]
    if ($null -eq $t) { return New-FailedTaskResponse "$($script:ProviderNamePrefix) Snapshot task [$TaskId] is no longer tracked" }
    if (-not (Test-VmeConnected)) { return @{ result = @{ state = 'running' } } }
    $elapsed = [int]([DateTime]::UtcNow - $t.started).TotalSeconds
    $timeout = [int]$script:Settings.linked_clones.snapshot_timeout_seconds
    try {
        if ($t.kind -eq 'delete') {
            $snap = Find-VmeInstanceSnapshot -VmId $t.template_id -Name $t.vme_name -Ref $TaskId
            $tv = $script:TemplateVersions[$t.template_id]
            if ($null -eq $snap) {
                if ($null -ne $tv -and $tv.versions.ContainsKey($t.ras_name)) { $tv.versions.Remove($t.ras_name); Save-TemplateVersions }
                $script:SnapshotTasks.Remove($TaskId)
                $imageNote = ''
                if ($t.image_id) {
                    try {
                        $img = Invoke-VmeApi -Path "/api/virtual-images/$($t.image_id)" -AllowNotFound -Ref $TaskId -FailureLevel 'W'
                        $imageNote = if ($null -eq $img) { "; its image [$($t.image_id)] is gone too" } else { "; its image [$($t.image_id)] STILL EXISTS on VME" }
                    }
                    catch { }
                }
                Write-DebugLog "Version [$($t.ras_name)] of template [$($t.template_id)] deleted: snapshot [$($t.snapshot_id)] gone after ${elapsed}s$imageNote." -Level $(if ($imageNote -match 'STILL') { 'W' } else { 'I' }) -Component '07' -Ref $TaskId
                return @{ result = @{ state = 'completed'; output = @{} } }
            }
            if ($elapsed -gt $timeout) {
                $script:SnapshotTasks.Remove($TaskId)
                if ($null -ne $tv -and $tv.versions.ContainsKey($t.ras_name)) { $tv.versions[$t.ras_name].status = 'ready'; Save-TemplateVersions }
                $msg = "$($script:ProviderNamePrefix) Snapshot [$($t.snapshot_id)] ($($t.vme_name)) still exists ${elapsed}s after its delete"
                Write-DebugLog $msg -Level 'E' -Component '07' -Ref $TaskId
                return New-FailedTaskResponse $msg
            }
            return @{ result = @{ state = 'running' } }
        }
        $p = Get-JsonValue (Invoke-VmeApi -Path "/api/processes/$($t.process_id)" -AllowNotFound -Ref $TaskId -FailureLevel 'W' -StatKey 'GET /api/processes/{id}') 'process'
        $ps = if ($null -ne $p) { ([string](Get-JsonValue $p 'status')).ToLowerInvariant() } else { 'complete' }
        if ($ps -eq 'complete') { $script:SnapshotTasks.Remove($TaskId); Write-DebugLog "The $($t.what) completed after ${elapsed}s." -Level 'I' -Component '07' -Ref $TaskId; return @{ result = @{ state = 'completed'; output = @{} } } }
        if ($ps -in @('failed', 'cancelled') -or $elapsed -gt $timeout) {
            $script:SnapshotTasks.Remove($TaskId)
            $why = Get-VmeProcessErrorText -Process $p
            $msg = "$($script:ProviderNamePrefix) The $($t.what) $(if ($ps -in @('failed', 'cancelled')) { "ended [$ps]" } else { "is still [$ps] after ${elapsed}s" })$(if ($why) { ": $why" })"
            Write-DebugLog $msg -Level 'E' -Component '07' -Ref $TaskId
            return New-FailedTaskResponse $msg
        }
        return @{ result = @{ state = 'running' } }
    }
    catch {
        Write-DebugLog "tasks/get for [$TaskId] failed this poll: $($_.Exception.Message)" -Level 'W' -Component '07' -Ref $TaskId
        return @{ result = @{ state = 'running' } }
    }
}

# The addInstance body for a linked clone. Unlike PUT /clone nothing is inherited from
# the template, so its placement, size and config flags are spelled out.
function New-VmeLinkedCloneBody {
    param([hashtable]$Inputs, [string]$CloneName, [string]$ImageId, [string]$NetId, [string]$PreservedMac)
    $instance = [ordered]@{ name = $CloneName }
    if ($Inputs.instance_type) { $instance.type = [string]$Inputs.instance_type }
    if ($Inputs.group) { $instance.site = @{ id = [long]$Inputs.group } }
    if ($Inputs.plan) { $instance.plan = @{ id = [long]$Inputs.plan } }
    if ($Inputs.layout) { $instance.layout = @{ id = [long]$Inputs.layout } }
    $config = [ordered]@{}
    if ($null -ne $Inputs.vm_config) { foreach ($k in $Inputs.vm_config.Keys) { $config[$k] = $Inputs.vm_config[$k] } }
    [long]$img = 0
    $config.imageId = if ([long]::TryParse($ImageId, [ref]$img)) { $img } else { $ImageId }
    if (-not $config.Contains('poolProviderType')) { $config.poolProviderType = 'mvm' }
    $config.createUser = $false
    $config.provisionPoweredOff = $false
    $body = [ordered]@{}
    if ($Inputs.cloud) { $body.zoneId = [long]$Inputs.cloud }
    $body.instance = $instance
    $body.config = $config
    $size = [ordered]@{}
    if ($null -ne $Inputs.max_cores) { $size.maxCores = [long]$Inputs.max_cores }
    if ($null -ne $Inputs.cores_per_socket) { $size.coresPerSocket = [long]$Inputs.cores_per_socket }
    if ($null -ne $Inputs.max_memory) { $size.maxMemory = [long]$Inputs.max_memory }
    if ($size.Count -gt 0) { $body.servicePlanOptions = $size }
    if ($null -ne $Inputs.volumes -and $Inputs.volumes.Count -gt 0) { $body.volumes = @($Inputs.volumes) }
    if ($NetId) {
        $nic = [ordered]@{ network = @{ id = [long]$NetId } }
        if ($PreservedMac) { $nic.macAddress = $PreservedMac }
        $body.networkInterfaces = @($nic)
    }
    return $body
}

# Clone body inputs for a source: from the cache (refreshed by every listing for known
# templates), else one lean GET of the source.
function Get-VmeCloneSourceInputs {
    param([string]$SourceId)
    $inputs = $script:CloneBodyBySource[$SourceId]
    $rec = $script:Inventory[$SourceId]
    if ($null -ne $inputs -and $null -ne $rec -and -not $rec.dirty) { return $inputs }
    $resp = Invoke-VmeApi -Path "/api/instances/${SourceId}?details=false" -AllowNotFound -Ref $SourceId
    $inst = Get-JsonValue $resp 'instance'
    if ($null -eq $inst) { return $null }
    $rec = ConvertTo-VmeInstanceRecord -Inst $inst
    if ($null -ne $rec) { $script:Inventory[$SourceId] = $rec }
    $inputs = Get-VmeCloneBodyInputs -Inst $inst
    $script:CloneBodyBySource[$SourceId] = $inputs
    return $inputs
}

# A template version's clone settings, read when the version is made: the template's
# lean payload (plan, size, disks, firmware, pool) plus its server-level network (the
# instance-level one goes stale on a VLAN change, see Update-VmeCloneSourceHardware).
# Pinned on the version, so a linked clone costs no template GETs, and every clone of
# a version gets that version's network and size -- not whatever the template has
# after a later maintenance.
function Get-VmeVersionCloneInputs {
    param([string]$TemplateId, [string]$Ref)
    $inputs = Get-VmeCloneSourceInputs -SourceId $TemplateId
    if ($null -eq $inputs) { return $null }
    Update-VmeCloneSourceHardware -SourceId $TemplateId
    $pinned = @{}
    foreach ($k in $inputs.Keys) { if ($k -ne 'learned_at') { $pinned[$k] = $inputs[$k] } }
    $hw = $script:HwById[$TemplateId]
    if ($null -ne $hw -and -not [string]::IsNullOrWhiteSpace($hw.network_id)) {
        if ($inputs.network -and $hw.network_id -ne $inputs.network) {
            Write-DebugLog "Template [$TemplateId]'s instance-level network [$($inputs.network)] ($($inputs.network_name)) differs from its server-level network [$($hw.network_id)] ($($hw.network_name)) -- using the server-level (live) one." -Level 'W' -Component '07' -Ref $Ref
        }
        $pinned.network = [string]$hw.network_id
        $pinned.network_name = [string]$hw.network_name
    }
    return $pinned
}

function Format-VmeCloneInputs {
    param([System.Collections.IDictionary]$Inputs)
    $disks = @(foreach ($v in @($Inputs.volumes)) { if ($null -ne $v) { "$($v['size'])GB" } }) -join '+'
    return "plan=$($Inputs.plan) layout=$($Inputs.layout) cores=$($Inputs.max_cores) memory=$([Math]::Round([double]$Inputs.max_memory / 1GB, 1))GB disks=$disks network=$($Inputs.network) ($($Inputs.network_name))"
}

# Morpheus's clone returns {"success":true} and nothing else -- no instance id, no
# process id -- so the clone is found afterwards by name (see Find-VmeCloneInstance).
function Handle-GuestClone {
    param([System.Collections.IDictionary]$Params)

    $sourceId = [string]$Params['id']
    $cloneName = [string]$Params['name']
    if ([string]::IsNullOrWhiteSpace($sourceId) -or [string]::IsNullOrWhiteSpace($cloneName)) { return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Both 'id' and 'name' are required" }
    if ($script:TemplateMethod -eq 'none') { return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Cloning is disabled (capabilities.template_method = none)" }
    if (-not (Test-VmeConnected)) { return New-NotConnectedResponse }
    # Under 'versioning' a snapshot name is a template version: a linked clone from its
    # image. Absent, "" and the name all occur (the Test Kit omits snapshot and
    # is_link_clone for a plain clone); is_link_clone=false is refused (VME can only
    # full-clone a VM's current disk, never a snapshot).
    $snapshotName = ([string]$Params['snapshot']).Trim()
    $linkedClone = ($script:TemplateMethod -eq 'versioning' -and $snapshotName -ne '')
    if ($linkedClone -and $null -ne $Params['is_link_clone']) {
        $isLink = $true
        try { $isLink = ConvertTo-CoercedSettingBool -Value $Params['is_link_clone'] -Default $true } catch { }
        if (-not $isLink) {
            Write-DebugLog "guests/clone [$cloneName] from template [$sourceId] version [$snapshotName] asked for a full clone (is_link_clone=false) -- refused: VME cannot full-clone a template version." -Level 'E' -Component '04' -Ref $sourceId
            return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Full clones of a template version are not supported on VME (it can only full-clone a VM's current disk, not a snapshot). Use linked clones for versioned templates."
        }
    }
    if ($snapshotName -and -not $linkedClone) {
        Write-DebugLog "guests/clone asked for snapshot [$snapshotName]; template_method is [$($script:TemplateMethod)], so VME clones from the live source and the parameter is ignored." -Level 'D' -Component '04' -Ref $sourceId
    }
    elseif (-not $snapshotName -and $script:TemplateMethod -eq 'versioning') {
        Write-DebugLog "guests/clone [$cloneName] from [$sourceId] names no template version -- full clone of the live source (a template made before template_method was 'versioning'?)." -Level 'W' -Component '04' -Ref $sourceId
    }

    # Idempotent re-submit: RAS retrying the same clone must not create a second VM.
    foreach ($k in @($script:CloneState.Keys)) {
        $e = $script:CloneState[$k]
        if ($e.name -ne $cloneName -or $null -ne $e.reported_completed_epoch) { continue }
        if ($e.source_id -eq $sourceId) {
            Write-DebugLog "guests/clone for [$cloneName] while the same clone is still in flight -- returning its task [$($e.task_id)] instead of cloning twice." -Level 'W' -Component '04' -Ref $e.task_id
            return @{ result = @{ task_id = $e.task_id } }
        }
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) A clone named [$cloneName] from source [$($e.source_id)] is still in flight"
    }
    if ($script:DeletingIds.ContainsKey($sourceId)) { return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Source guest [$sourceId] is being deleted" }

    try {
        # Clone slots, before any VME call. Given-up clones are re-checked on VME only when
        # they are what blocks this request (throttled), so the steady state costs nothing.
        $sameNameStuck = { @($script:StuckClones.Values | Where-Object { $_.name -eq $cloneName -and -not $_.clone_id }) }
        $max = [int]$script:Settings.cloning.max_concurrent_clone_operations
        if (@(& $sameNameStuck).Count -gt 0 -or (Get-VmeCloneSlotHolders).Count -ge $max) { Update-VmeStuckClones }
        # Clones reported early may have finished without anyone re-reading them yet.
        if ((Get-VmeCloneSlotHolders).Count -ge $max) { Update-VmeEarlyCloneSlots -ExceptKey '' -Ref $sourceId }
        # A same-named clone that never appeared could still materialise and be taken for
        # this one by the name-based resolution.
        $blocker = @(& $sameNameStuck)
        if ($blocker.Count -gt 0) {
            Write-DebugLog "guests/clone [$cloneName] refused: an earlier clone under this name was given up without ever appearing and VME has not finished it yet." -Level 'W' -Component '04' -Ref $sourceId
            return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) An earlier clone named [$cloneName] (template [$($blocker[0].source_id)]) was given up but VME has not finished it -- not cloning the same name again until it does"
        }
        $holders = Get-VmeCloneSlotHolders
        if ($holders.Count -ge $max) {
            $list = $holders -join ', '
            Write-DebugLog "guests/clone [$cloneName] from [$sourceId] refused: $($holders.Count) of $max clone slot(s) held -- $list." -Level 'W' -Component '04' -Ref $sourceId
            return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Clone limit reached: $($holders.Count) of $max clone slot(s) (cloning.max_concurrent_clone_operations) are held by clones VME has not finished: $list. [$cloneName] was not submitted."
        }


        $version = $null
        if ($linkedClone) {
            $tvx = $script:TemplateVersions[$sourceId]
            $pendingVersion = if ($null -ne $tvx) { $tvx.versions[$snapshotName] } else { $null }
            if ($null -ne $pendingVersion -and $pendingVersion.status -eq 'creating') {
                return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Version [$snapshotName] of template [$sourceId] is still being created"
            }
            $version = Resolve-VmeTemplateVersion -TemplateId $sourceId -RasName $snapshotName -Ref $sourceId
            if ($null -eq $version) {
                $vmeName = Get-VmeVersionVmeName -TemplateId $sourceId -RasName $snapshotName
                Write-DebugLog "guests/clone [$cloneName]: template [$sourceId] has no version [$snapshotName] on VME (no snapshot + linked-clone image named [$vmeName])." -Level 'E' -Component '04' -Ref $sourceId
                return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Template [$sourceId] has no version [$snapshotName] on VME (snapshot and linked-clone image [$vmeName] not found)"
            }
        }

        # A version's clone settings were pinned when it was made (no template GETs per
        # clone). A version without them (made by 4.1, or rebuilt from VME) takes the
        # template's current ones once, at this first clone.
        $pinned = $null
        if ($null -ne $version) {
            if ($version.clone_inputs -is [System.Collections.IDictionary] -and $version.clone_inputs.Count -gt 0) { $pinned = $version.clone_inputs }
            else {
                $pinned = Get-VmeVersionCloneInputs -TemplateId $sourceId -Ref $sourceId
                if ($null -ne $pinned) {
                    $version.clone_inputs = $pinned
                    Save-TemplateVersions
                    Write-DebugLog "Version [$snapshotName] of template [$sourceId] had no clone settings -- taken from the template now and kept for its later clones: $(Format-VmeCloneInputs -Inputs $pinned)." -Level 'W' -Component '07' -Ref $sourceId
                }
            }
        }
        $inputs = if ($null -ne $pinned) { $pinned } else { Get-VmeCloneSourceInputs -SourceId $sourceId }
        if ($null -eq $inputs) { return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Source guest [$sourceId] not found" }
        if (-not (Test-VmeIsBasicTemplate -VmId $sourceId)) {
            Write-DebugLog "guests/clone: source [$sourceId] is not marked as a template in this provider (guests/convert was never received for it) -- cloning anyway." -Level 'W' -Component '04' -Ref $sourceId
        }

        # VME doesn't enforce unique names. Instances already carrying this name, and
        # every id that existed before the submit (ids only grow), can never be the
        # clone -- that is what makes resolution by name safe.
        $preexisting = [System.Collections.Generic.List[string]]::new()
        $maxKnownId = [long]0
        foreach ($id in $script:Inventory.Keys) {
            [long]$n = 0
            if ([long]::TryParse($id, [ref]$n) -and $n -gt $maxKnownId) { $maxKnownId = $n }
            if ($script:Inventory[$id].name -eq $cloneName) { $preexisting.Add($id) }
        }
        if ($preexisting.Count -gt 0) {
            $racing = @($preexisting | Where-Object { $script:DeletingIds.ContainsKey($_) })
            $note = if ($racing.Count) { " ($($racing -join ', ') still being deleted -- a recreate racing its own delete)" } else { '' }
            Write-DebugLog "guests/clone: [$cloneName] already exists as instance(s) $($preexisting -join ', ')$note; they are excluded from resolution." -Level 'W' -Component '04' -Ref $sourceId
        }

        $preservedMac = $null
        if ($script:Settings.cloning.mac_preservation.enabled) {
            $m = $script:RecentlyDeletedMacByName[$cloneName]
            if ($null -ne $m -and ([DateTime]::UtcNow - $m.deleted_at).TotalSeconds -lt $script:Settings.cloning.mac_preservation.retention_seconds) { $preservedMac = [string]$m.mac }
            $script:RecentlyDeletedMacByName.Remove($cloneName)
        }

        # Echo the source's placement: a bare {name} is accepted but lands on a
        # default network and the layout's default disk size.
        #
        # Network comes from the SERVER object (HwById, same source as the MAC/OS
        # below), not the instance payload's own interfaces[].network -- Morpheus does
        # not keep the instance-level field in sync when a
        # VM's network is changed underneath it (server-level was correct, instance-level
        # was still the old VLAN). Refreshed with one call right here (not left to the
        # bulk pass's cadence) so a just-changed VLAN is never missed.
        $netId = $inputs.network
        $netName = $inputs.network_name
        $hwSource = $null
        if ($null -eq $pinned) {
            Update-VmeCloneSourceHardware -SourceId $sourceId
            $hwSource = $script:HwById[$sourceId]
        }
        if ($null -ne $hwSource -and -not [string]::IsNullOrWhiteSpace($hwSource.network_id)) {
            if ($inputs.network -and $hwSource.network_id -ne $inputs.network) {
                Write-DebugLog "guests/clone: source [$sourceId]'s instance-level network [$($inputs.network)] ($($inputs.network_name)) differs from its server-level network [$($hwSource.network_id)] ($($hwSource.network_name)) -- using the server-level (live) one." -Level 'W' -Component '04' -Ref $sourceId
            }
            $netId = $hwSource.network_id
            $netName = $hwSource.network_name
        }

        if (-not $netId -and $preservedMac) { Write-DebugLog "guests/clone: preserved MAC [$preservedMac] for [$cloneName] dropped -- the source has no network interface to attach it to." -Level 'W' -Component '04' -Ref $sourceId }
        $uncertain = $false
        $newInst = $null
        $newId = ''

        if ($null -ne $version) {
            # Linked clone: a new instance built on the version's image (addInstance).
            $body = New-VmeLinkedCloneBody -Inputs $inputs -CloneName $cloneName -ImageId ([string]$version.image_id) -NetId $netId -PreservedMac $preservedMac
            Write-DebugLog "guests/clone: [$sourceId] version [$snapshotName] -> [$cloneName] linked clone from image [$($version.image_id)] type=$($inputs.instance_type) group=$($inputs.group) cloud=$($inputs.cloud) plan=$($inputs.plan) layout=$($inputs.layout) network=$netId volumes=$($inputs.volumes.Count) config=[$(@(foreach ($k in $body.config.Keys) { "$k=$($body.config[$k])" }) -join ' ')]$(if ($preservedMac) { " mac=$preservedMac (preserved)" }) (version's clone settings)." -Level 'T' -Component '04' -Ref $sourceId
            try {
                # Never retried: a second POST would be a second VM.
                $resp = Invoke-VmeApi -Path '/api/instances' -Method POST -Body $body -Ref $sourceId -StatKey 'POST /api/instances (linked clone)'
                if ($false -eq (Get-JsonValue $resp 'success')) {
                    $msg = [string](Get-JsonValue $resp 'msg')
                    $errs = Get-JsonValue $resp 'errors'
                    if ($errs -is [System.Collections.IDictionary]) { $msg = "$msg $(@(foreach ($k in $errs.Keys) { "${k}: $($errs[$k])" }) -join '; ')" }
                    return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) VME did not accept the linked clone of [$sourceId] version [$snapshotName] as [$cloneName]$(if ($msg.Trim()) { ": $($msg.Trim())" })"
                }
                $newInst = Get-JsonValue $resp 'instance'
                $newId = [string](Get-JsonValue $newInst 'id')
                if (-not $newId) { Write-DebugLog "POST /api/instances for [$cloneName] answered without an instance id -- resolving the clone by name." -Level 'W' -Component '04' -Ref $sourceId }
            }
            catch {
                if (-not (Test-VmeNoResponse $_)) { throw }
                $uncertain = $true
                Write-DebugLog "Linked clone of [$sourceId] as [$cloneName] got no response ($($_.Exception.Message)) -- tracking it anyway; resolution will find the clone by name or fail the task cleanly." -Level 'W' -Component '04' -Ref $sourceId
            }
        }
        else {
            $body = [ordered]@{ name = $cloneName }
            if ($inputs.volumes.Count -gt 0) { $body.volumes = @($inputs.volumes) }
            if ($inputs.group) { $body.group = @{ id = [long]$inputs.group } }
            if ($inputs.cloud) { $body.cloud = @{ id = [long]$inputs.cloud }; $body.zoneId = [long]$inputs.cloud }
            if ($inputs.plan) { $body.plan = @{ id = [long]$inputs.plan } }
            if ($inputs.layout) { $body.layout = @{ id = [long]$inputs.layout } }
            if ($netId) {
                # networkInterfaces[].macAddress is declared by addInstance
                # (instanceCreateNetwork), whose parameters cloneInstance accepts too.
                # Whether VME's clone honours it is verified after completion.
                $nic = [ordered]@{ network = @{ id = [long]$netId } }
                if ($preservedMac) { $nic.macAddress = $preservedMac }
                $body.networkInterfaces = @($nic)
            }

            Write-DebugLog "guests/clone: [$sourceId] -> [$cloneName] group=$($inputs.group) cloud=$($inputs.cloud) plan=$($inputs.plan) layout=$($inputs.layout) network=$netId volumes=$($inputs.volumes.Count)$(if ($preservedMac) { " mac=$preservedMac (preserved)" })." -Level 'T' -Component '04' -Ref $sourceId

            try {
                # Never retried (see Invoke-VmeHttp).
                $resp = Invoke-VmeApi -Path "/api/instances/$sourceId/clone" -Method PUT -Body $body -Ref $sourceId
                if ($true -ne (Get-JsonValue $resp 'success')) {
                    $msg = [string](Get-JsonValue $resp 'msg')
                    return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) VME did not accept the clone of [$sourceId] as [$cloneName]$(if ($msg) { ": $msg" })"
                }
            }
            catch {
                if (-not (Test-VmeNoResponse $_)) { throw }
                # No HTTP response: VME may have accepted it. Failing now would make RAS
                # retry under the same name and create a duplicate; resolution decides.
                $uncertain = $true
                Write-DebugLog "guests/clone of [$sourceId] as [$cloneName] got no response ($($_.Exception.Message)) -- tracking it anyway; resolution will find the clone or fail the task cleanly." -Level 'W' -Component '04' -Ref $sourceId
            }
        }

        # Unique per submission. A recreate under the same name (RAS: delete, wait until
        # gone, clone again) must not reuse the previous task id: tasks/get would answer
        # it from the remembered completed output with the DELETED instance's id (covered
        # by tests/unit-vme/vme-u3).
        $kind = if ($null -ne $version) { 'linked' } else { 'clonebyname' }
        $taskId = "$($kind):$($sourceId):$($cloneName):$(Get-EpochNow)$('{0:x4}' -f (Get-Random -Maximum 65536))"
        $entry = @{
            type                     = 'clone'
            task_id                  = $taskId
            source_id                = $sourceId
            source_uuid              = Get-RecordUuid -VmId $sourceId
            clone_id                 = ''
            clone_uuid               = ''
            name                     = $cloneName
            started_epoch            = Get-EpochNow
            resolved_epoch           = $null
            preexisting_ids          = $preexisting.ToArray()
            max_known_id             = $maxKnownId
            submit_uncertain         = $uncertain
            provision_confirmed      = $false
            reported_completed_epoch = $null
            preserved_mac            = $preservedMac
            requested_network        = [string]$netId
            requested_network_name   = [string]$netName
            image_id                 = $(if ($null -ne $version) { [string]$version.image_id } else { '' })
            version                  = $(if ($null -ne $version) { $snapshotName } else { '' })
        }
        if ($newId) {
            # addInstance returned the instance: no resolution by name, provisioning is
            # tracked from the first tasks/get (it stays hidden from guests/list until then).
            $entry.clone_id = $newId
            $entry.clone_uuid = [string](Get-JsonValue $newInst 'uuid')
            $entry.resolved_epoch = Get-EpochNow
            if (-not [string]::IsNullOrWhiteSpace([string](Get-JsonValue $newInst 'status'))) {
                $rec = ConvertTo-VmeInstanceRecord -Inst $newInst
                if ($null -ne $rec) { $script:Inventory[$newId] = $rec }
            }
            Set-CloneEntry -Key $newId -Entry $entry
        }
        else { Set-CloneEntry -Key "pending:$cloneName" -Entry $entry }
        [void]$script:KnownCloneSources.Add($sourceId)
        if ($null -ne $version) {
            Write-DebugLog "Linked clone of template [$sourceId] version [$snapshotName] (image [$($version.image_id)]) as [$cloneName] submitted$(if ($newId) { ": instance [$newId]" } elseif ($uncertain) { ' (response lost)' }); task [$taskId]." -Level 'I' -Component '04' -Ref $taskId
        }
        else { Write-DebugLog "Clone of instance [$sourceId] as [$cloneName] submitted$(if ($uncertain) { ' (response lost)' }); task [$taskId]." -Level 'I' -Component '04' -Ref $taskId }
        return @{ result = @{ task_id = $taskId } }
    }
    catch {
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to clone guest [$sourceId]: $($_.Exception.Message)"
    }
}

# Finds the instance a clone created. Candidates must: carry the exact name, not be
# the source, not have existed when the clone was submitted (not a preexisting
# same-named instance, id above every id known then), not belong to another tracked
# clone, and -- when VME reports lineage -- be cloned from this source.
#
# Sources, cheapest first: the inventory (free); the configured targeted query
# (phrase/activity: a few KB); a shared full listing at most every
# resolve_list_interval_seconds as the safety net whatever the method.
function Find-VmeCloneInstance {
    param([hashtable]$Entry, [string]$TaskId)

    $src = [string]$Entry.source_id
    $name = [string]$Entry.name
    $pre = [System.Collections.Generic.HashSet[string]]::new([string[]]@($Entry.preexisting_ids))
    $maxId = [long]$Entry.max_known_id
    $myKey = Get-CloneKeyByTaskId -TaskId $TaskId

    $isCandidateId = {
        param([string]$CandidateId)
        if ($CandidateId -eq $src -or $pre.Contains($CandidateId)) { return $false }
        [long]$n = 0
        if ([long]::TryParse($CandidateId, [ref]$n) -and $n -le $maxId) { return $false }
        $owner = $script:CloneKeyByCloneId[$CandidateId]
        return ($null -eq $owner -or $owner -eq $myKey)
    }
    $scan = {
        $found = [System.Collections.Generic.List[hashtable]]::new()
        foreach ($id in $script:Inventory.Keys) {
            $rec = $script:Inventory[$id]
            if ($rec.name -ne $name -or -not (& $isCandidateId $id)) { continue }
            if (-not [string]::IsNullOrWhiteSpace($rec.cloned_from) -and $rec.cloned_from -ne $src) {
                Write-DebugLog "Clone resolution: [$name] instance [$id] rejected -- VME says it was cloned from [$($rec.cloned_from)], not [$src]." -Level 'D' -Component '08' -Ref $TaskId
                continue
            }
            # A linked clone (lost POST response) is built from a known image.
            if (-not [string]::IsNullOrWhiteSpace([string]$Entry.image_id) -and -not [string]::IsNullOrWhiteSpace([string]$rec.image_id) -and [string]$rec.image_id -ne [string]$Entry.image_id) {
                Write-DebugLog "Clone resolution: [$name] instance [$id] rejected -- built from image [$($rec.image_id)], not the version image [$($Entry.image_id)]." -Level 'D' -Component '08' -Ref $TaskId
                continue
            }
            $found.Add($rec)
        }
        return , $found
    }

    $method = 'inventory'
    $found = & $scan
    $method2 = [string]$script:Settings.cloning.resolve_method
    $queried = $false
    if ($found.Count -eq 0 -and $method2 -eq 'phrase') {
        $queried = $true
        try {
            $resp = Invoke-VmeApi -Path "/api/instances?phrase=$([uri]::EscapeDataString($name))&max=50$($script:InstanceScopeQuery)" -Ref $TaskId -FailureLevel 'W' -StatKey 'GET /api/instances (phrase)'
            foreach ($inst in @(Get-JsonValue $resp 'instances')) {
                $rec = ConvertTo-VmeInstanceRecord -Inst $inst
                if ($null -ne $rec -and $rec.name -eq $name) { $script:Inventory[$rec.id] = $rec }
            }
            $found = & $scan
            $method = 'phrase'
        }
        catch { }
    }
    elseif ($found.Count -eq 0 -and $method2 -eq 'activity') {
        $queried = $true
        try {
            $start = ConvertTo-IsoFromEpoch ([long]$Entry.started_epoch - 120)
            $resp = Invoke-VmeApi -Path "/api/activity?name=$([uri]::EscapeDataString($name))&start=$start&max=25" -Ref $TaskId -FailureLevel 'W' -StatKey 'GET /api/activity'
            $ids = [System.Collections.Generic.HashSet[string]]::new()
            foreach ($a in @(Get-JsonValue $resp 'activity')) {
                if ([string](Get-JsonValue $a 'objectType') -eq 'Instance') { [void]$ids.Add([string](Get-JsonValue $a 'objectId')) }
            }
            foreach ($id in $ids) {
                if ([string]::IsNullOrWhiteSpace($id) -or -not (& $isCandidateId $id)) { continue }
                [void](Get-VmeInstanceRecord -VmId $id -Ref $TaskId -FailureLevel 'W')
            }
            $found = & $scan
            $method = 'activity'
        }
        catch { }
    }

    $interval = [int]$script:Settings.cloning.resolve_list_interval_seconds
    if ($found.Count -eq 0 -and ((Get-EpochNow) - [long]$Entry.started_epoch) -ge $interval) {
        if (Update-VmeInventory -MaxAgeSeconds $interval -Reason 'clone resolution') {
            $found = & $scan
            $method = 'listing'
            if ($found.Count -gt 0 -and $queried) {
                Write-DebugLog "Clone resolution: [$name] was found by the full listing but not by the '$method2' query on the same poll -- if this repeats, that query doesn't work on this appliance; set cloning.resolve_method to 'list'." -Level 'W' -Component '08' -Ref $TaskId
            }
        }
    }

    if ($found.Count -eq 0) { return $null }
    $pick = $found[0]
    if ($found.Count -gt 1) {
        $lineage = @($found | Where-Object { $_.cloned_from -eq $src })
        $pool = if ($lineage.Count -gt 0) { $lineage } else { @($found) }
        $pick = $pool | Sort-Object { [long]$_.id } -Descending | Select-Object -First 1
        Write-DebugLog "Clone resolution: [$name] matched $($found.Count) new instances ($(($found | ForEach-Object { $_.id }) -join ', ')) -- using [$($pick.id)]." -Level 'W' -Component '08' -Ref $TaskId
    }
    return @{ record = $pick; method = $method }
}

function Test-VmePreservedMacOutcome {
    param([string]$Key)
    $entry = $script:CloneState[$Key]
    if ($null -eq $entry -or [string]::IsNullOrWhiteSpace([string]$entry.preserved_mac) -or $entry.ContainsKey('mac_verified')) { return }
    $hw = $script:HwById[[string]$entry.clone_id]
    if ($null -eq $hw -or -not $hw.complete) { return }
    $entry.mac_verified = $true
    if ($hw.macs -contains ([string]$entry.preserved_mac).ToUpperInvariant()) {
        Write-DebugLog "MAC preservation: clone [$($entry.clone_id)] ($($entry.name)) got the requested MAC [$($entry.preserved_mac)] -- VME honours networkInterfaces[].macAddress on clone." -Level 'I' -Component '04' -Ref $entry.clone_id
    }
    else {
        Write-DebugLog "MAC preservation: clone [$($entry.clone_id)] ($($entry.name)) got MAC [$($hw.macs -join ',')] instead of the requested [$($entry.preserved_mac)] -- VME ignored networkInterfaces[].macAddress on clone." -Level 'W' -Component '04' -Ref $entry.clone_id
    }
}

# cloneInstance's own schema only declares {name, group} -- everything else (network,
# plan, layout, volumes) is accepted only per the endpoint description's blanket "same
# parameters as instance creation" claim, and the description also says the clone is
# produced from a backup/snapshot of the source. Both are reasons VME could silently
# keep the source's network/VLAN instead of the one requested. This makes that
# observable instead of assumed.
function Test-VmeNetworkPreservationOutcome {
    param([string]$Key)
    $entry = $script:CloneState[$Key]
    if ($null -eq $entry -or [string]::IsNullOrWhiteSpace([string]$entry.requested_network) -or $entry.ContainsKey('network_verified')) { return }
    $hw = $script:HwById[[string]$entry.clone_id]
    if ($null -eq $hw -or [string]::IsNullOrWhiteSpace([string]$hw.network_id)) { return }
    $entry.network_verified = $true
    $wantName = if ($entry.requested_network_name) { " ($($entry.requested_network_name))" } else { '' }
    $gotName = if ($hw.network_name) { " ($($hw.network_name))" } else { '' }
    if ([string]$hw.network_id -eq [string]$entry.requested_network) {
        Write-DebugLog "Network preservation: clone [$($entry.clone_id)] ($($entry.name)) landed on the requested network [$($entry.requested_network)]$wantName -- VME honours networkInterfaces[].network.id on clone." -Level 'I' -Component '04' -Ref $entry.clone_id
    }
    else {
        Write-DebugLog "Network preservation: clone [$($entry.clone_id)] ($($entry.name)) landed on network [$($hw.network_id)]$gotName instead of the requested [$($entry.requested_network)]$wantName -- VME ignored networkInterfaces[].network.id on clone; the clone kept the source's network/VLAN as of when the clone request was submitted." -Level 'W' -Component '04' -Ref $entry.clone_id
    }
}

function New-FailedTaskResponse {
    param([string]$Message)
    return @{ result = @{ state = 'failed'; error = @{ code = 1; message = $Message } } }
}

# For a clone, tasks/get is a small state machine:
#   stage 1 -- find the instance (Find-VmeCloneInstance), within resolve_timeout_seconds;
#   stage 2 -- one lean GET per poll until VME's own provisioning status is terminal,
#              within provision_timeout_seconds (both measured from the submit, so
#              they don't shrink when tasks_polling_rate is lowered);
#   done    -- completed with output.clone_id. Power-on and the IP are guests/get's
#              business.
# The state file is written on transitions only.
function Handle-TaskInfo {
    param([System.Collections.IDictionary]$Params)

    $taskId = [string]$Params['id']
    if ([string]::IsNullOrWhiteSpace($taskId)) { return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Invalid task id" }
    if ($taskId -like 'noop:*') { return @{ result = @{ state = 'completed'; output = @{} } } }
    if ($taskId -like 'snapcreate:*') { return (Invoke-VmeSnapshotCreateTask -TaskId $taskId) }
    if ($taskId -like 'snapdelete:*' -or $taskId -like 'snapproc:*') { return (Invoke-VmeSnapshotTask -TaskId $taskId) }

    $remembered = Get-CompletedCloneTaskOutput -TaskId $taskId
    if ($null -ne $remembered) { return @{ result = @{ state = 'completed'; output = @{ clone_id = $remembered } } } }

    $key = Get-CloneKeyByTaskId -TaskId $taskId
    if ($null -eq $key) {
        $stuck = $script:StuckClones[$taskId]
        if ($null -ne $stuck) { return New-FailedTaskResponse "$($script:ProviderNamePrefix) Clone [$($stuck.name)] was given up ($($stuck.reason)); VME has not finished it yet, so it still holds a clone slot" }
        if ($taskId -like 'clonebyname:*' -or $taskId -like 'linked:*') { return New-FailedTaskResponse "$($script:ProviderNamePrefix) Clone task [$taskId] is no longer tracked" }
        return New-FailedTaskResponse "$($script:ProviderNamePrefix) Unknown task [$taskId]"
    }
    $entry = $script:CloneState[$key]
    if (-not (Test-VmeConnected)) { return @{ result = @{ state = 'running' } } }

    $C = $script:Settings.cloning
    $name = $entry.name
    $elapsed = (Get-EpochNow) - [long]$entry.started_epoch
    # linked_clones.early_completion: a linked clone from an immutable version image
    # competes for no source disk, so it is reported once VME has an id for it.
    $early = ($entry.image_id -and $script:Settings.linked_clones.early_completion)
    $uncertainNote = if ($entry.submit_uncertain) { ' The clone request itself got no response from VME, so VME may never have received it.' } else { '' }

    # Reported early by the deploy-event gate: RAS already has the id.
    if ($null -ne $entry.reported_completed_epoch -and $entry.clone_id) {
        return @{ result = @{ state = 'completed'; output = @{ clone_id = $entry.clone_id } } }
    }

    try {
        $rec = $null
        if ([string]::IsNullOrWhiteSpace($entry.clone_id)) {
            if ($elapsed -gt $C.resolve_timeout_seconds) {
                Write-DebugLog "Clone [$name] was not found ${elapsed}s after the submit -- giving up.$uncertainNote" -Level 'E' -Component '08' -Ref $taskId
                # VME may still be working on it (provision processes can stay stuck for hours
                # with no instance) -- keep its slot until VME's process history says otherwise.
                Add-VmeStuckClone -Entry $entry -Reason "not found within $($C.resolve_timeout_seconds)s"
                Remove-CloneEntry -Key $key
                return New-FailedTaskResponse "$($script:ProviderNamePrefix) Clone [$name] never appeared within $($C.resolve_timeout_seconds)s.$uncertainNote"
            }
            $found = Find-VmeCloneInstance -Entry $entry -TaskId $taskId
            if ($null -eq $found) {
                Write-DebugLog "Clone [$name] not found yet (${elapsed}s)." -Level 'D' -Component '08' -Ref $taskId
                return @{ result = @{ state = 'running' } }
            }
            $rec = $found.record
            $entry.clone_id = $rec.id
            $entry.clone_uuid = $rec.uuid
            $entry.resolved_epoch = Get-EpochNow
            Move-CloneEntryKey -FromKey $key -ToKey $rec.id
            $key = $rec.id
            $script:CloneLastRefresh[$rec.id] = [DateTime]::UtcNow
            Write-DebugLog "Clone [$name] resolved to instance [$($rec.id)] ${elapsed}s after the submit (via $($found.method)$(if ($rec.cloned_from) { ", cloned from [$($rec.cloned_from)]" }))." -Level 'I' -Component '08' -Ref $taskId
        }
        else {
            # A linked clone reported early needs no GET on its first poll: the record
            # addInstance returned a second ago is as fresh as VME can give.
            $rec = $script:Inventory[[string]$entry.clone_id]
            $freshEnough = ($early -and $null -ne $rec -and -not $rec.dirty -and $rec.status -and ([DateTime]::UtcNow - $rec.fetched_at).TotalSeconds -lt [int]$C.tracked_refresh_interval_seconds)
            if (-not $freshEnough) {
                $rec = Get-VmeInstanceRecord -VmId $entry.clone_id -Ref $taskId
                $script:CloneLastRefresh[$entry.clone_id] = [DateTime]::UtcNow
            }
            if ($null -eq $rec) {
                Write-DebugLog "Clone [$name] (instance [$($entry.clone_id)]) disappeared during provisioning (404) after ${elapsed}s." -Level 'E' -Component '08' -Ref $taskId
                Clear-VmeTrackingForVm -VmId $entry.clone_id -Gone
                return New-FailedTaskResponse "$($script:ProviderNamePrefix) Clone [$name] (id=$($entry.clone_id)) disappeared during provisioning"
            }
        }

        $cloneId = $rec.id
        $class = Get-VmeProvisionStatusClass -Status $rec.status -IsBusy $rec.is_busy

        if ($class -eq 'failed') {
            $detail = $null
            try { $detail = Get-VmeProcessErrorText -Process (Get-VmeNewestProvisionProcess -VmId $cloneId -Ref $taskId) } catch { }
            $message = "$($script:ProviderNamePrefix) Clone [$name] (id=$cloneId) failed to provision (status=$($rec.status), ${elapsed}s)$(if ($detail) { ": $detail" })"
            Write-DebugLog $message -Level 'E' -Component '08' -Ref $taskId
            Clear-VmeTrackingForVm -VmId $cloneId -KeepRecord
            return New-FailedTaskResponse $message
        }

        if ($class -eq 'pending') {
            if ($elapsed -gt $C.provision_timeout_seconds) {
                Write-DebugLog "Clone [$name] (id=$cloneId) still [$($rec.status)] ${elapsed}s after the submit -- giving up." -Level 'E' -Component '08' -Ref $taskId
                if ($C.cancel_process_on_giveup) {
                    $stuckProc = $null
                    try { $stuckProc = Get-VmeNewestProvisionProcess -VmId $cloneId -Ref $taskId } catch { }
                    $stuckStatus = if ($null -ne $stuckProc) { [string](Get-JsonValue $stuckProc 'status') } else { $null }
                    if ($null -ne $stuckProc -and $stuckStatus -in @('new', 'queued', 'pending', 'running')) {
                        Invoke-VmeProcessCancel -ProcessId ([string](Get-JsonValue $stuckProc 'id')) -Ref $taskId
                    }
                }
                Add-VmeStuckClone -Entry $entry -Reason "still provisioning after $($C.provision_timeout_seconds)s" -Status ([string]$rec.status)
                Clear-VmeTrackingForVm -VmId $cloneId -KeepRecord
                return New-FailedTaskResponse "$($script:ProviderNamePrefix) Clone [$name] (id=$cloneId) still [$($rec.status)] after $($C.provision_timeout_seconds)s"
            }
            # Early completion is held while the clone limit is in use: reporting this
            # clone lets RAS send the next guests/clone at once, which is refused unless
            # a slot is free. So report only while fewer than the limit (this clone
            # included) are still provisioning.
            if ($early) {
                $active = (Get-VmeCloneSlotHolders).Count
                if ($active -ge $C.max_concurrent_clone_operations) {
                    Update-VmeEarlyCloneSlots -ExceptKey $key -Ref $taskId
                    $active = (Get-VmeCloneSlotHolders).Count
                }
                if ($active -lt $C.max_concurrent_clone_operations) {
                    $entry.reported_completed_epoch = Get-EpochNow
                    Set-CloneEntry -Key $key -Entry $entry
                    Set-CompletedCloneTaskOutput -TaskId $taskId -CloneId $cloneId
                    Write-DebugLog "Linked clone [$name] (id=$cloneId) reported completed at ${elapsed}s, VME status [$($rec.status)]: it keeps provisioning on VME while RAS sends the next clone. $active of $($C.max_concurrent_clone_operations) clone slot(s) in use." -Level 'I' -Component '08' -Ref $taskId
                    return @{ result = @{ state = 'completed'; output = @{ clone_id = $cloneId } } }
                }
                if (-not $entry.ContainsKey('early_held')) {
                    $entry['early_held'] = $true
                    Write-DebugLog "Linked clone [$name] (id=$cloneId) held at ${elapsed}s: $active clone(s) still provisioning, limit $($C.max_concurrent_clone_operations) (cloning.max_concurrent_clone_operations)." -Level 'T' -Component '08' -Ref $taskId
                }
            }
            elseif ($C.deploy_event_gate_enabled) {
                $proc = $null
                try { $proc = Get-VmeNewestProvisionProcess -VmId $cloneId -Ref $taskId } catch { }
                Write-VmeProvisionTrace -Process $proc -VmId $cloneId -Ref $taskId -Elapsed $elapsed
                if ((Get-VmeDeployEventStatus -Process $proc) -eq 'complete') {
                    $active = (Get-VmeCloneSlotHolders).Count
                    # -lt, not -le: at the limit, RAS's next guests/clone would be refused.
                    if ($active -lt $C.max_concurrent_clone_operations) {
                        $entry.reported_completed_epoch = Get-EpochNow
                        Set-CloneEntry -Key $key -Entry $entry
                        Set-CompletedCloneTaskOutput -TaskId $taskId -CloneId $cloneId
                        Write-DebugLog "Clone [$name] (id=$cloneId) reported completed at ${elapsed}s: its Deploy step is done, VME status still [$($rec.status)]. $active clone(s) provisioning, limit $($C.max_concurrent_clone_operations)." -Level 'I' -Component '08' -Ref $taskId
                        return @{ result = @{ state = 'completed'; output = @{ clone_id = $cloneId } } }
                    }
                    Write-DebugLog "Clone [$name] (id=$cloneId): Deploy step done but held -- $active clone(s) provisioning, limit $($C.max_concurrent_clone_operations)." -Level 'T' -Component '08' -Ref $taskId
                }
            }
            Write-DebugLog "Clone [$name] (id=$cloneId) provisioning: status=$($rec.status) percent=$($rec.status_percent) ${elapsed}s$(if ($rec.status_message) { " -- $($rec.status_message)" })" -Level 'D' -Component '08' -Ref $taskId
            return @{ result = @{ state = 'running' } }
        }

        $entry.provision_confirmed = $true
        $entry.reported_completed_epoch = Get-EpochNow
        Set-CloneEntry -Key $key -Entry $entry
        Set-CompletedCloneTaskOutput -TaskId $taskId -CloneId $cloneId
        Update-VmeHardwareForRecord -Rec $rec -Force
        Test-VmePreservedMacOutcome -Key $key
        Write-DebugLog "Clone [$name] completed as instance [$cloneId] (status=$($rec.status)) ${elapsed}s after the submit. Power-on and IP continue via guests/get." -Level 'I' -Component '08' -Ref $taskId
        return @{ result = @{ state = 'completed'; output = @{ clone_id = $cloneId } } }
    }
    catch {
        # A transient failure costs one poll, not the task: the time budgets bound it.
        Write-DebugLog "tasks/get for [$taskId] failed this poll: $($_.Exception.Message)" -Level 'W' -Component '08' -Ref $taskId
        return @{ result = @{ state = 'running' } }
    }
}

# ---------------------------------------------------------------------------
# Dispatch
#
# The CPF methods minus hosts/* (an unknown method returns -32601). guests/snapshots/*
# answer -32601 too unless template_method is 'versioning' (Get-VmeSnapshotRequest).
# ---------------------------------------------------------------------------

$script:MethodHandlers = @{
    'provider/initialize' = @{ Handler = { param($d) Handle-Initialize }; RequiredFields = @() }
    'provider/connect'    = @{ Handler = { param($d) Handle-Connect -Params $d['params'] }; RequiredFields = @('params.settings') }
    'provider/disconnect' = @{ Handler = { param($d) Handle-Disconnect }; RequiredFields = @() }
    'guests/list'         = @{ Handler = { param($d) Handle-GuestList }; RequiredFields = @() }
    'guests/get'          = @{ Handler = { param($d) Handle-GuestGet -Params $d['params'] }; RequiredFields = @('params.id') }
    'guests/control'      = @{ Handler = { param($d) Handle-GuestControl -Params $d['params'] }; RequiredFields = @('params.id', 'params.control') }
    'guests/convert'      = @{ Handler = { param($d) Handle-GuestConvert -Params $d['params'] }; RequiredFields = @('params.id', 'params.is_template') }
    'guests/clone'        = @{ Handler = { param($d) Handle-GuestClone -Params $d['params'] }; RequiredFields = @('params.id', 'params.name') }
    'tasks/get'           = @{ Handler = { param($d) Handle-TaskInfo -Params $d['params'] }; RequiredFields = @('params.id') }
    'guests/snapshots/create' = @{ Handler = { param($d) Handle-SnapshotCreate -Params $d['params'] }; RequiredFields = @('params.id', 'params.name') }
    'guests/snapshots/exists' = @{ Handler = { param($d) Handle-SnapshotExists -Params $d['params'] }; RequiredFields = @('params.id', 'params.name') }
    'guests/snapshots/delete' = @{ Handler = { param($d) Handle-SnapshotDelete -Params $d['params'] }; RequiredFields = @('params.id', 'params.name') }
    'guests/snapshots/revert' = @{ Handler = { param($d) Handle-SnapshotRevert -Params $d['params'] }; RequiredFields = @('params.id', 'params.name') }
}

function Invoke-ProviderMethod {
    param([System.Collections.IDictionary]$Data, [string]$Method)

    Update-ProviderSettingsIfChanged
    if ([string]::IsNullOrWhiteSpace($Method)) { return New-ErrorResponse -Code $script:ErrorCodes.InvalidRequest -Message "$($script:ProviderNamePrefix) Request has no 'method'" }
    $spec = $script:MethodHandlers[$Method]
    if ($null -eq $spec) { return New-ErrorResponse -Code $script:ErrorCodes.MethodNotFound -Message "$($script:ProviderNamePrefix) Unsupported method [$Method]" }
    $missing = Test-RequiredFields -Data $Data -Fields $spec.RequiredFields
    if ($null -ne $missing) { return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Missing required field [$missing]" }
    try { return (& $spec.Handler $Data) }
    catch { return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Unhandled error in [$Method]: $($_.Exception.Message)" }
}

# ---------------------------------------------------------------------------
# Startup
# ---------------------------------------------------------------------------

$script:LastBaseUrl = $null
Import-ProviderSettings
Set-ProviderRuntimeFromSettings
try { $script:SettingsFileLastWriteUtc = [System.IO.File]::GetLastWriteTimeUtc($script:SettingsPath) } catch { }
$script:SettingsLastPolledAt = [DateTime]::UtcNow

Import-CloneState
Import-StuckClones
$script:MaintenanceIds = Import-IdUuidMap -Path $script:MaintenanceStatePath -Label 'maintenance state'
$script:BasicTemplates = Import-IdUuidMap -Path $script:TemplateStatePath -Label 'template state'
$script:FormerTemplates = Import-IdUuidMap -Path $script:FormerTemplateStatePath -Label 'former template state'
Import-TemplateVersions

$osDesc = try { [System.Environment]::OSVersion.VersionString } catch { 'unknown OS' }
$S0 = $script:Settings
Write-DebugLog "Module version $($script:ProviderVersion) - $($script:ScriptFileName) (release $($script:ScriptVersion))" -Level 'I' -Component '00'
Write-DebugLog 'Starting RAS CPF Provider - HPE VME' -Level 'I' -Component '00'
Write-DebugLog "Host - $([System.Environment]::MachineName), PowerShell $($PSVersionTable.PSVersion) ($($ExecutionContext.SessionState.LanguageMode)), $osDesc; connection reuse with per-request timeout: $($script:HasOperationTimeout)" -Level 'I' -Component '00'
Write-DebugLog ("Mode - template_method=$($script:TemplateMethod), guests_polling_rate=$($S0.capabilities.guests_polling_rate)s, tasks_polling_rate=$($S0.capabilities.tasks_polling_rate)s, " +
    "list_min_refresh=$($S0.inventory.list_min_refresh_seconds)s, page_size=$($S0.inventory.page_size) (cold start $($S0.inventory.cold_start_page_size)), log_level=$($script:LogLevel), " +
    "tracked clones=$($script:CloneState.Count), given-up clones holding a slot=$($script:StuckClones.Count), in maintenance=$($script:MaintenanceIds.Count), templates=$($script:BasicTemplates.Count), " +
    "template versions=$(@($script:TemplateVersions.Values | ForEach-Object { $_.versions.Values }).Count) on $($script:TemplateVersions.Count) template(s), can_link_clones=$($script:TemplateMethod -eq 'versioning')") -Level 'I' -Component '0A'
if ($script:TemplateMethod -eq 'versioning') {
    Write-DebugLog ("Linked clones - snapshot timeout=$($S0.linked_clones.snapshot_timeout_seconds)s, image wait=$($S0.linked_clones.image_wait_seconds)s, image filterType=$($S0.linked_clones.image_filter_type), " +
        "label images=$($S0.linked_clones.label_images), refuse delete in use=$($S0.linked_clones.refuse_delete_in_use), early completion=$($S0.linked_clones.early_completion)") -Level 'I' -Component '0A'
}
Write-DebugLog ("Cloning - resolve_method=$($S0.cloning.resolve_method), resolve/provision timeout=$($S0.cloning.resolve_timeout_seconds)/$($S0.cloning.provision_timeout_seconds)s, " +
    "deploy_event_gate=$($S0.cloning.deploy_event_gate_enabled), max_concurrent=$($S0.cloning.max_concurrent_clone_operations), " +
    "flapping_workaround=$($S0.cloning.template_flapping_workaround_enabled), mac_preservation=$($S0.cloning.mac_preservation.enabled), force_stop=$($S0.power.force_stop), stop_before_delete=$($S0.delete.stop_before_delete), " +
    "power action hold=$($S0.inventory.power_action_hold_seconds)s, watch refresh=$($S0.inventory.watch_refresh_interval_seconds)s") -Level 'I' -Component '0A'
Write-DebugLog "Provider process started. PID=$PID" -Level 'I' -Component '00'

# ---------------------------------------------------------------------------
# Main loop -- one JSON request per line in, one response per line out. Nothing in
# here may end the loop except stdin closing or stdout breaking.
# ---------------------------------------------------------------------------

try {
    while (-not $script:StdoutBroken) {
        $line = [Console]::In.ReadLine()
        if ($null -eq $line) { break }
        if ([string]::IsNullOrWhiteSpace($line)) { continue }

        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $method = '(unparsed)'
        $response = $null
        try {
            $data = $null
            try { $data = ConvertFrom-JsonFast -Json $line } catch { }
            if ($data -isnot [System.Collections.IDictionary]) {
                Write-DebugLog "IN (malformed): $(ConvertTo-VmeLoggableText $line 500)" -Level 'W' -Component '01'
                $response = New-ErrorResponse -Code $script:ErrorCodes.ParseError -Message "$($script:ProviderNamePrefix) Malformed JSON request"
            }
            else {
                $method = ([string]$data['method']).Trim().ToLowerInvariant()
                # Masked by Write-DebugLog: provider/connect carries credentials.
                if ($script:LogLevel -ge 5) { Write-DebugLog "IN: $(ConvertTo-VmeLoggableText $line 2000)" -Level 'D' -Component '01' }
                $response = Invoke-ProviderMethod -Data $data -Method $method
            }
        }
        catch {
            Write-DebugLog "Unhandled error serving [$method]: $($_.Exception.Message)" -Level 'E' -Component '01'
            $response = New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Unhandled error: $($_.Exception.Message)"
        }

        Send-Response -Response $response -Method $method
        $sw.Stop()
        $ms = $sw.Elapsed.TotalMilliseconds
        $isError = ($response -is [System.Collections.IDictionary] -and $response.Contains('error'))
        Add-RpcStat -Method $method -Ms $ms -IsError $isError
        if ($ms -ge 10000) { Write-DebugLog "Slow RPC: [$method] took $([int]$ms)ms -- every other RAS request waited behind it." -Level 'W' -Component '01' }
        elseif ($ms -ge 2000) { Write-DebugLog "Slow RPC: [$method] took $([int]$ms)ms." -Level 'T' -Component '01' }
        Invoke-StatsSummaryIfDue
    }
}
finally {
    Write-StatsSummary -Reason 'shutdown'
    Write-DebugLog 'Provider stopping -- stdin closed.' -Level 'I' -Component '00'
    Close-LogWriter
}
