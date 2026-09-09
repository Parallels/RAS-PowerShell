$ErrorActionPreference = 'Stop'
$pass = 0
$fail = 0
function Assert {
    param([bool]$Cond, [string]$Msg)
    if ($Cond) { $script:pass++; Write-Host "PASS: $Msg" -ForegroundColor Green }
    else { $script:fail++; Write-Host "FAIL: $Msg" -ForegroundColor Red }
}

# ============================================================
# TESTS for rasTemplate<id> tag lifecycle: added on convert-to-template, removed when
# RAS deletes the Template object -- but NOT when RAS merely enters maintenance, even
# though both start with the identical guests/convert {is_template:false} call.
#
# The distinguishing signal: entering maintenance always
# follows the convert with guests/control start (RAS needs the guest running to patch
# it); deleting the Template object never does (RAS is only tidying up its own object).
# Since this provider has no timer, Handle-GuestConvert arms a pending-removal marker
# instead of deciding immediately; Handle-GuestControl's 'start' cancels it; and
# Resolve-PendingTemplateTagRemoval -- piggybacked on the next guests/get for that vmid
# -- removes the tag once the grace window passes with no start.
#
# guests/snapshots/delete is a second, immediate signal for linked-clone templates
# specifically: entering maintenance never touches snapshots, and exiting maintenance's
# own delete step is unreachable while already de-templated (Handle-GuestSnapshotsExists
# reads the live template flag), so a delete that does arrive means the guest was still
# templated when RAS called it -- i.e. deleting the Template object.
# ============================================================
. "$PSScriptRoot/funcs2.ps1"

$script:LogPath        = "$PSScriptRoot/t13.log"
$script:CloneStatePath = "$PSScriptRoot/t13-state.json"
Remove-Item $script:LogPath, $script:CloneStatePath -ErrorAction SilentlyContinue
$script:TrackedCloneVmIds = [System.Collections.Generic.HashSet[string]]::new()
$script:ClusterResourcesCache = $null
$script:ClusterResourcesCachedAt = $null
$script:RecentlyDeletedIds = @{}
$script:RecentlyControlledIds = @{}
$script:RecentlyStoppedIds = @{}
$script:RecentlyConvertedIds = @{}
$script:PendingTemplateTagRemovalIds = @{}
$script:TemplateSnapshotIntent = @{}
$script:AgentFailureTracker = @{}
$script:LastKnownNetworkData = @{}
$script:CloneStateMemory = $null
$script:TaskContext = @{}
$script:TemplateDeleteConfirmSeconds = 10
$script:RecentlyConvertedRetentionSeconds = 120

$script:apiCalls = New-Object System.Collections.ArrayList
$script:vmState    = @{}   # vmid -> status
$script:vmName     = @{}
$script:vmTemplate = @{}   # vmid -> $true/$false, the LIVE config's template flag
$script:vmTags     = @{}   # vmid -> tag string
$script:forceTemplatePostError = $null

function Invoke-ProxmoxApi {
    param([string]$Method, [string]$Path, $Body, [int]$TimeoutSec = 0)
    [void]$script:apiCalls.Add("$Method $Path")

    if ($Path -match '/cluster/resources') {
        $list = New-Object System.Collections.ArrayList
        foreach ($k in $script:vmState.Keys) {
            $entry = [pscustomobject]@{ type = 'qemu'; vmid = [int]$k; name = $script:vmName[$k]; node = 'n1'; status = $script:vmState[$k]
                template = $(if ($script:vmTemplate[$k]) { 1 } else { 0 }) }
            if ($script:vmTags.ContainsKey($k) -and -not [string]::IsNullOrWhiteSpace($script:vmTags[$k])) {
                $entry | Add-Member -NotePropertyName tags -NotePropertyValue $script:vmTags[$k]
            }
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

    if ($Method -eq 'POST' -and $Path -match '/qemu/(\d+)/template$') {
        $vmid = $Matches[1]
        if ($null -ne $script:forceTemplatePostError) {
            $msg = $script:forceTemplatePostError
            $script:forceTemplatePostError = $null
            if ($msg -match 'template to a template') { $script:vmTemplate[$vmid] = $true }
            throw $msg
        }
        $script:vmTemplate[$vmid] = $true
        return @{ data = "UPID:n1:qmtemplate$vmid" }
    }

    if ($Path -match '/qemu/(\d+)/status/start$') {
        $vmid = $Matches[1]
        $script:vmState[$vmid] = 'running'
        return @{ data = "UPID:n1:qmstart$vmid" }
    }

    if ($Path -match '/qemu/(\d+)/status/current$') {
        $vmid = $Matches[1]
        $st = if ($script:vmState.ContainsKey($vmid)) { $script:vmState[$vmid] } else { 'stopped' }
        return @{ data = [pscustomobject]@{ status = $st; qmpstatus = $st } }
    }

    if ($Path -match '/qemu/(\d+)/agent/network-get-interfaces$') {
        throw 'agent not queried in this test'
    }

    throw "Unexpected call in mock: $Method $Path"
}

function HasError { param($R) return ($R -is [hashtable] -and $R.ContainsKey('error')) }
function Get-Tags { param([string]$VmId) return @(($script:vmTags[$VmId]) -split ';' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) }

# ------------------------------------------------------------
# A. Convert to template adds rasTemplate<id> -- fresh conversion and idempotent repeat.
# ------------------------------------------------------------
$script:vmState['800'] = 'stopped'; $script:vmName['800'] = 'vm800'; $script:vmTemplate['800'] = $false

$rA1 = Handle-GuestConvert -Params ([pscustomobject]@{ id = '800'; is_template = $true })
Assert (-not (HasError $rA1)) "A1: convert VM->template succeeds"
Assert ((Get-Tags '800') -contains 'rasTemplate800') "A2: converting to a template adds rasTemplate800"

$putCallsBefore = @($script:apiCalls | Where-Object { $_ -match '^PUT .*/qemu/800/config$' }).Count
$rA2 = Handle-GuestConvert -Params ([pscustomobject]@{ id = '800'; is_template = $true })
Assert (-not (HasError $rA2)) "A3: repeated convert to template (idempotent) does not error"
$putCallsAfter = @($script:apiCalls | Where-Object { $_ -match '^PUT .*/qemu/800/config$' }).Count
Assert ($putCallsAfter -eq $putCallsBefore) "A4: ...and issues no redundant tag PUT (already tagged)"
Assert (@(Get-Tags '800' | Where-Object { $_ -eq 'rasTemplate800' }).Count -eq 1) "A5: no duplicate rasTemplate800 tag"

# ------------------------------------------------------------
# B. Convert to VM arms the pending marker; a prompt 'start' cancels it (maintenance).
# ------------------------------------------------------------
$rB1 = Handle-GuestConvert -Params ([pscustomobject]@{ id = '800'; is_template = $false })
Assert (-not (HasError $rB1)) "B1: convert template->VM (entering maintenance) succeeds"
Assert ($script:PendingTemplateTagRemovalIds.ContainsKey('800')) "B2: the transition arms the pending-removal marker"
Assert ((Get-Tags '800') -contains 'rasTemplate800') "B3: the tag is NOT removed immediately -- only Resolve-PendingTemplateTagRemoval or a start decides that"

$rB2 = Handle-GuestControl -Params ([pscustomobject]@{ id = '800'; control = 'start' })
Assert (-not (HasError $rB2)) "B4: guests/control(start) succeeds"
Assert (-not $script:PendingTemplateTagRemovalIds.ContainsKey('800')) "B5: the start cancels the pending removal -- this was maintenance"

[void](ConvertTo-RasGuestObject -VmId '800')
Assert ((Get-Tags '800') -contains 'rasTemplate800') "B6: a subsequent guests/get confirms the tag is kept"

# ------------------------------------------------------------
# C. Convert to VM with NO start -> tag removed once the grace window elapses
#    (deleting the Template object).
# ------------------------------------------------------------
$script:vmState['801'] = 'stopped'; $script:vmName['801'] = 'vm801'; $script:vmTemplate['801'] = $false
Reset-ProxmoxClusterCache
[void](Handle-GuestConvert -Params ([pscustomobject]@{ id = '801'; is_template = $true }))
Assert ((Get-Tags '801') -contains 'rasTemplate801') "C1: setup -- VM 801 is a tagged template"

$rC1 = Handle-GuestConvert -Params ([pscustomobject]@{ id = '801'; is_template = $false })
Assert (-not (HasError $rC1)) "C2: convert template->VM (deleting the Template object) succeeds"
Assert ($script:PendingTemplateTagRemovalIds.ContainsKey('801')) "C3: the transition arms the pending-removal marker"
Assert ((Get-Tags '801') -contains 'rasTemplate801') "C4: the tag is still present immediately after -- no start has had a chance to arrive or not"

# No guests/control start ever follows. Age the marker past the grace window, exactly
# like every other Recently*-style test in this suite.
$script:PendingTemplateTagRemovalIds['801'] = ([DateTime]::UtcNow).AddSeconds(-($script:TemplateDeleteConfirmSeconds + 1))
[void](ConvertTo-RasGuestObject -VmId '801')
Assert (-not $script:PendingTemplateTagRemovalIds.ContainsKey('801')) "C5: the next guests/get resolves the marker (grace window elapsed, no start)"
Assert ((Get-Tags '801') -notcontains 'rasTemplate801') "C6: ...and removes rasTemplate801 -- the machine is fully cleaned up"

$putCallsBeforeC = @($script:apiCalls | Where-Object { $_ -match '^PUT .*/qemu/801/config$' }).Count
[void](ConvertTo-RasGuestObject -VmId '801')
$putCallsAfterC = @($script:apiCalls | Where-Object { $_ -match '^PUT .*/qemu/801/config$' }).Count
Assert ($putCallsAfterC -eq $putCallsBeforeC) "C7: a further guests/get for the same vmid issues no redundant tag PUT (tag already gone, marker already cleared)"

# ------------------------------------------------------------
# D. guests/snapshots/delete removes the tag immediately -- the linked-clone-specific
#    signal, no waiting required. (LINKED-CLONES.md #5 -- entering maintenance never
#    calls this; exiting maintenance's own delete step is unreachable while already
#    de-templated, since Handle-GuestSnapshotsExists reads the live template flag.)
# ------------------------------------------------------------
$script:vmState['802'] = 'stopped'; $script:vmName['802'] = 'vm802'; $script:vmTemplate['802'] = $false
Reset-ProxmoxClusterCache
[void](Handle-GuestConvert -Params ([pscustomobject]@{ id = '802'; is_template = $true }))
Assert ((Get-Tags '802') -contains 'rasTemplate802') "D1: setup -- VM 802 is a tagged template"

$rD1 = Handle-GuestSnapshotsDelete -Params ([pscustomobject]@{ id = '802'; name = 'RAS Template Snapshot' })
Assert (-not [string]::IsNullOrWhiteSpace([string]$rD1.result.task_id)) "D2: guests/snapshots/delete still returns a task_id"
Assert ((Get-Tags '802') -notcontains 'rasTemplate802') "D3: ...and removes rasTemplate802 immediately -- no grace window needed for this signal"

# ------------------------------------------------------------
# E. The raced convert-to-template path (Proxmox says 'already a template' when our own
#    live check said otherwise) still adds the tag -- same as the normal path.
# ------------------------------------------------------------
$script:vmState['803'] = 'stopped'; $script:vmName['803'] = 'vm803'; $script:vmTemplate['803'] = $false
Reset-ProxmoxClusterCache
$script:forceTemplatePostError = "Response status code does not indicate success: 500 (you can't convert a template to a template)."
$rE1 = Handle-GuestConvert -Params ([pscustomobject]@{ id = '803'; is_template = $true })
Assert (-not (HasError $rE1)) "E1: a raced 'template to a template' 500 is treated as success"
Assert ((Get-Tags '803') -contains 'rasTemplate803') "E2: ...and still tags the guest as a RAS template"

Write-Host "`n=== t13 Summary: $pass passed, $fail failed ===" -ForegroundColor $(if ($fail -eq 0) { 'Green' } else { 'Red' })
if ($fail -gt 0) { exit 1 }
