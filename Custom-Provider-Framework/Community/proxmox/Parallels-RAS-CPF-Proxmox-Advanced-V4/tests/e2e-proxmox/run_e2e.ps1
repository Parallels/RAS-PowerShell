$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding  = [System.Text.Encoding]::UTF8

$RepoRoot   = (Resolve-Path (Join-Path $PSScriptRoot '..' '..')).Path
$TestKitDir = Join-Path $RepoRoot 'Framework Test Kit'
Import-Module (Join-Path $TestKitDir 'CustomProvider.psm1') -Force

# The provider is copied fresh on every run rather than kept as a checked-in duplicate.
# An out-of-date copy is the worst possible failure mode for this suite -- it passes
# loudly while testing code you no longer ship -- and it has already bitten this project
# once. The RAS test kit needs an absolute CommandArgs, so the .psd1 is generated here
# too instead of holding a hard-coded path that breaks the moment the folder moves.
$ProviderSrc = Join-Path $RepoRoot 'Parallels-RAS-CPF-Proxmox-Advanced.ps1'
$ProviderRun = Join-Path $PSScriptRoot 'provider.run.ps1'
Copy-Item -LiteralPath $ProviderSrc -Destination $ProviderRun -Force

$pwshPath = (Get-Process -Id $PID).Path
$ProviderSettings = @{
    CommandPath    = $pwshPath
    CommandArgs    = "-NoProfile -NonInteractive -File `"$ProviderRun`""
    CustomSettings = @{ host = '127.0.0.1:8443'; username = 'test@pve'; token_name = 'auto'; token_secret = 'secret' }
}

# Enabled for the whole run rather than toggled mid-test: a non-empty 'snapshot' param
# is still required to trigger a linked clone (see Handle-GuestClone), so every existing
# full-clone assertion above is unaffected by this being on. Written before the provider
# subprocess starts -- settings are read once at cold start, and RAS-CPF-Proxmox-Settings.json
# lives next to the script itself (data_directory is not used for this file, see SETTINGS.md).
$SettingsPath = Join-Path $PSScriptRoot 'RAS-CPF-Proxmox-Settings.json'
@{ capabilities = @{ can_link_clones = $true }; cloning = @{ linked_clone_fallback = 'full' } } |
    ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $SettingsPath -Encoding UTF8

# Write-DebugLog goes straight to this file, independent of the stdin/stdout JSON-RPC
# pipe -- used below to confirm the "fell back to a full clone" decision is logged, not
# silent. Default location is the provider's own folder (locations.data_directory
# defaults to $PSScriptRoot), default name Proxmox-RAS-Provider.log.
$ProviderLog = Join-Path $PSScriptRoot 'Proxmox-RAS-Provider.log'
Remove-Item $ProviderLog -ErrorAction SilentlyContinue

# --- fixture: a FRESH mock every run -------------------------------------
# The mock keeps all VM state in memory, and this suite mutates it (it clones,
# tags, stops and deletes). Reusing a long-lived mock therefore leaks state
# between runs, and it produced a genuine intermittent failure: VM 153 kept
# 'rasTemplate153' from the previous run, so the provider's first cluster
# snapshot already showed the tag, Add-ProxmoxVmTag correctly skipped the
# write, and the assertion that the tag gets written failed -- a red result
# from a green provider. Persisted clone state is dropped for the same reason.
$MockScript = Join-Path $PSScriptRoot 'mock_pve.py'
$MockLog = Join-Path $PSScriptRoot 'mock.log'

# 'pgrep'/'kill' are the macOS/Linux tools this suite was originally written for; on
# Windows, find/stop the mock by matching its full command line instead (Get-Process
# alone doesn't expose arguments, only the process name).
function Stop-Mock {
    if ($IsWindows) {
        Get-CimInstance Win32_Process -ErrorAction SilentlyContinue |
            Where-Object { $_.CommandLine -and $_.CommandLine -like '*mock_pve.py*' } |
            ForEach-Object { try { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue } catch {} }
    }
    else {
        $existing = & pgrep -f 'mock_pve.py' 2>$null
        foreach ($procId in @($existing)) {
            if (-not [string]::IsNullOrWhiteSpace($procId)) { & kill $procId 2>$null }
        }
    }
    Start-Sleep -Milliseconds 300
}

# 'python3' is the macOS/Linux convention this suite was originally written for; a
# standard Windows Python install only ships 'python.exe', not 'python3.exe'.
function Resolve-Python3Command {
    foreach ($name in @('python3', 'python')) {
        $cmd = Get-Command $name -ErrorAction SilentlyContinue
        if ($null -ne $cmd) { return $cmd.Source }
    }
    throw "No Python 3 interpreter found on PATH (tried 'python3', 'python')."
}

function Start-Mock {
    Stop-Mock
    Remove-Item (Join-Path $PSScriptRoot 'Proxmox-RAS-CloneState.json') -ErrorAction SilentlyContinue
    # Quoted: the vault path contains spaces, and an unquoted argument is split on them.
    Start-Process -FilePath (Resolve-Python3Command) -ArgumentList "`"$MockScript`"" `
        -RedirectStandardOutput $MockLog -RedirectStandardError "$MockLog.err" `
        -WorkingDirectory $PSScriptRoot | Out-Null

    $deadline = (Get-Date).AddSeconds(15)
    while ((Get-Date) -lt $deadline) {
        try {
            Invoke-RestMethod -Uri 'https://127.0.0.1:8443/api2/json/version' -SkipCertificateCheck -TimeoutSec 2 | Out-Null
            Write-Host "mock_pve started and answering on :8443" -ForegroundColor DarkGray
            return
        }
        catch { Start-Sleep -Milliseconds 200 }
    }
    throw "mock_pve.py did not become ready on :8443 within 15s -- see $MockLog.err"
}

Start-Mock

$pass = 0
$fail = 0
function Assert {
    param([bool]$Cond, [string]$Msg)
    if ($Cond) { $script:pass++; Write-Host "PASS: $Msg" -ForegroundColor Green }
    else { $script:fail++; Write-Host "FAIL: $Msg" -ForegroundColor Red }
}

# Reads a VM's tags directly from the mock PVE (bypassing the provider), the
# same way an admin would check tags in the Proxmox UI/API.
function Get-MockVmTags {
    param([string]$VmId)
    $resp = Invoke-RestMethod -Uri "https://127.0.0.1:8443/api2/json/nodes/n1/qemu/$VmId/config" -SkipCertificateCheck
    return [string]$resp.data.tags
}

# Writes a VM's tags directly against the mock PVE, bypassing the provider --
# simulates an admin having hand-tagged a template in the Proxmox UI before
# this provider ever touches it.
function Set-MockVmTags {
    param([string]$VmId, [string]$Tags)
    Invoke-RestMethod -Uri "https://127.0.0.1:8443/api2/json/nodes/n1/qemu/$VmId/config" -Method Put -Body @{ tags = $Tags } -SkipCertificateCheck | Out-Null
}

# 'clone_full' is not a real Proxmox field -- see mock_pve.py's GET /config handler --
# it is the mock's own record of the 'full' value the provider actually asked for on
# the clone that produced this VM, letting the suite assert linked-vs-full directly
# instead of inferring it from timing.
function Get-MockVmCloneFull {
    param([string]$VmId)
    $resp = Invoke-RestMethod -Uri "https://127.0.0.1:8443/api2/json/nodes/n1/qemu/$VmId/config" -SkipCertificateCheck
    return $resp.data.clone_full
}

function Invoke-Pipeline {
    param([object]$IOStreams)

    Write-Host "`n=== initialize + connect ===" -ForegroundColor Cyan
    Submit-InitializeAndConnect $IOStreams $ProviderSettings.CustomSettings | Out-Null
    Assert $true "provider/initialize + provider/connect completed without throwing (Read-ResultObject throws on any JSON-RPC error)"

    Write-Host "`n=== guests/list ===" -ForegroundColor Cyan
    $list1 = Submit-GuestsList $IOStreams
    Assert ($list1.guests -contains '100') "guests/list includes VM 100"
    Assert ($list1.guests -contains '101') "guests/list includes VM 101 (agentless)"
    Assert (-not ($list1.guests -contains '102')) "tags feature: guests/list omits VM 102 (tagged rasExclude)"

    $excludedThrew = $false
    try { Submit-GuestsGet $IOStreams '102' | Out-Null } catch { $excludedThrew = $true }
    Assert $excludedThrew "tags feature: guests/get for the rasExclude VM 102 behaves like a not-found VM"

    Write-Host "`n=== guests/get (healthy + agentless) ===" -ForegroundColor Cyan
    $g100 = Submit-GuestsGet $IOStreams '100'
    Assert ($g100.ip -eq '10.0.0.100') "guests/get for VM 100 reports its real IP via the agent"
    Assert ($g100.state -eq 'powered_on') "guests/get for VM 100 reports powered_on"

    $g101a = Submit-GuestsGet $IOStreams '101'
    Assert ($g101a.state -eq 'powered_on') "guests/get for agentless VM 101 still reports powered_on (no crash on agent 500)"
    Assert (@($g101a.ip_addresses).Count -eq 0) "guests/get for agentless VM 101 reports no IP on the first poll"

    Write-Host "`n=== guests/clone (real subprocess, real HTTPS to the mock) ===" -ForegroundColor Cyan
    # Simulates an admin having hand-tagged the template before this provider
    # ever touches it -- the clone's tag-inherit-then-strip pipeline must
    # preserve this hand-set tag while still stripping the inherited
    # rasTemplate<id> tag (see below).
    Set-MockVmTags -VmId '153' -Tags 'cat'
    $cloneResult = Submit-GuestsClone $IOStreams '153' 'e2e-clone-1' $null
    Assert (-not [string]::IsNullOrWhiteSpace($cloneResult.task_id)) "guests/clone returns a task_id"
    Assert (-not [string]::IsNullOrWhiteSpace($cloneResult.clone_id)) "guests/clone returns a clone_id"
    $cloneId = [string]$cloneResult.clone_id
    $taskId = [string]$cloneResult.task_id
    Write-Host "  clone_id=$cloneId task_id=$taskId"

    $sourceTagsAfterClone = @((Get-MockVmTags '153') -split ';')
    Assert ($sourceTagsAfterClone -contains 'rasTemplate153') "tags feature: source template 153 tagged rasTemplate153 immediately on clone"
    Assert ($sourceTagsAfterClone -contains 'cat') "tags feature: source template 153 keeps its own hand-set 'cat' tag"

    Write-Host "`n=== guests/get for the brand-new clone id, immediately (the #2 race) ===" -ForegroundColor Cyan
    # The mock server reveals the new VM in cluster/resources ~50ms after the
    # clone call returns (simulating real Proxmox), so an immediate call can
    # legitimately land in the window where Get-ProxmoxVmNode still throws
    # "not found" -- Get-RasGuestObjectForCloneAwareFlow's own catch handles
    # that by reporting a synthetic 'powering_on' guest rather than a hard
    # error (a real placeholder name at that point, by design -- the id
    # isn't resolvable in the cluster listing at all yet, so there is nothing
    # to substitute a name onto). Poll briefly past that window rather than
    # asserting on a single racy call.
    $raceDeadline = (Get-Date).AddSeconds(3)
    $gRace = $null
    do {
        $gRace = Submit-GuestsGet $IOStreams $cloneId
        Assert ($gRace.id -eq $cloneId) "guests/get for the freshly-cloned id does not throw a hard error, even mid-race"
        if ($gRace.name -eq 'e2e-clone-1') { break }
        Start-Sleep -Milliseconds 20
    } while ((Get-Date) -lt $raceDeadline)
    # Once the id is resolvable in the cluster listing at all, RAS must never
    # see Proxmox's transient placeholder name for a clone this provider is
    # tracking -- only the real, intended name.
    Assert ($gRace.name -eq 'e2e-clone-1') "guests/get for the freshly-cloned id reports its real intended name instead of Proxmox's placeholder, once resolvable"

    Write-Host "`n=== guests/list -- tracked clone excluded until its clone task is reported completed ===" -ForegroundColor Cyan
    # guests/list's gate on a tracked clone is not Proxmox's (or the
    # provider's own) name for it; it's whether this provider has itself
    # reported the clone task 'completed' to RAS yet (see Handle-TaskInfo /
    # Handle-GuestList). At this point tasks/get has not been polled at all,
    # so the clone must still be excluded even though its name already
    # resolves correctly above -- RAS's clone thread and guests/list must
    # never disagree about when a clone becomes "real" to the outside world,
    # or RAS's own sync walker can bind the id to a stale/recycled record
    # before the clone thread itself gets a chance to.
    $listBeforeCompletion = Submit-GuestsList $IOStreams
    Assert (-not (@($listBeforeCompletion.guests) -contains $cloneId)) "guests/list excludes the tracked clone before its task is reported completed"

    Write-Host "`n=== tasks/get -- poll the clone task to pipelined completion ===" -ForegroundColor Cyan
    $deadline = (Get-Date).AddSeconds(15)
    $taskState = $null
    while ((Get-Date) -lt $deadline) {
        $taskResp = Submit-TasksGet $IOStreams $taskId
        $taskState = $taskResp.state
        if ($taskState -eq 'completed') { break }
        Start-Sleep -Milliseconds 500
    }
    Assert ($taskState -eq 'completed') "tasks/get eventually reports the clone task completed"

    $listAfterCompletion = Submit-GuestsList $IOStreams
    Assert (@($listAfterCompletion.guests) -contains $cloneId) "guests/list includes the tracked clone once its task has been reported completed"

    Write-Host "`n=== guests/get -- poll until the clone is powered on with an IP ===" -ForegroundColor Cyan
    $deadline = (Get-Date).AddSeconds(15)
    $ready = $false
    $lastGuest = $null
    while ((Get-Date) -lt $deadline) {
        $lastGuest = Submit-GuestsGet $IOStreams $cloneId
        if ($lastGuest.state -eq 'powered_on' -and @($lastGuest.ip_addresses).Count -gt 0) { $ready = $true; break }
        Start-Sleep -Milliseconds 300
    }
    if (-not $ready) { Write-Host "  last guest object: $($lastGuest | ConvertTo-Json -Compress)" -ForegroundColor Yellow }
    Assert $ready "the clone-aware flow auto-starts the VM and eventually reports it powered_on with an IP"
    $cloneTagsAfterReady = @((Get-MockVmTags $cloneId) -split ';')
    Assert ($cloneTagsAfterReady -contains 'rasClone153') "tags feature: clone VM $cloneId tagged rasClone153 once its real clone job completed"
    # The clone inherited the source's 'cat' tag AND its rasTemplate153 tag
    # at clone time (real Proxmox clone-copies-config behavior, simulated by
    # the mock too). The clone must end up with 'cat' preserved and
    # rasTemplate153 stripped.
    Assert ($cloneTagsAfterReady -contains 'cat') "clone VM $cloneId keeps the inherited hand-set 'cat' tag"
    Assert ($cloneTagsAfterReady -notcontains 'rasTemplate153') "clone VM $cloneId does not keep the inherited rasTemplate153 tag"

    Write-Host "`n=== guests/control stop + delete ===" -ForegroundColor Cyan
    $stopResult = Submit-GuestsControl $IOStreams $cloneId 'stop'
    Assert ($stopResult.action -eq 'stop') "guests/control(stop) maps to a hard stop action"

    $deleteResult = Submit-GuestsControl $IOStreams $cloneId 'delete'
    Assert ($deleteResult.action -eq 'delete') "guests/control(delete) succeeds"

    Write-Host "`n=== guests/list -- deleted id should be gone immediately ===" -ForegroundColor Cyan
    $list2 = Submit-GuestsList $IOStreams
    Assert (-not ($list2.guests -contains $cloneId)) "guests/list omits the just-deleted clone immediately"

    $gDeleted = Submit-GuestsGet $IOStreams $cloneId
    Assert ($gDeleted.state -eq 'powered_off') "guests/get for the deleted id returns cleanly instead of erroring"
    # RAS's deserializer hard-requires 'name' and rejects the entire object without
    # it, so even this legitimate "it's gone" stub must carry a placeholder.
    Assert (-not [string]::IsNullOrEmpty($gDeleted.name)) "the deleted-id stub carries a non-null name RAS can deserialize"

    Write-Host "`n=== cluster/nextid recycles the just-deleted VMID ===" -ForegroundColor Cyan
    # The defect this guards against. cluster/nextid returns the LOWEST free id, so RAS
    # recreating a pool (delete five, clone five) lands the new clone on the id just
    # deleted -- while that id's 300s recently-deleted marker is still armed. Without a
    # fix, the marker would suppress the BRAND NEW VM: guests/list omitting it for up
    # to 300s and guests/get answering with the name=$null "it's gone" stub, which RAS
    # cannot deserialize, so RAS stops polling the clone it was just handed.
    # The mock's next_id is fixed, so this second clone reuses $cloneId exactly as
    # real Proxmox does.
    $reclone = Submit-GuestsClone $IOStreams '153' 'e2e-clone-2'
    Assert ($reclone.clone_id -eq $cloneId) "the second clone really does reuse the just-deleted VMID $cloneId (mirrors cluster/nextid)"

    # The decisive assertion: BEFORE the fix this answered with the deleted-VM stub
    # (state powered_off, name null) for the full 300s retention. It must instead
    # describe the live clone.
    $gReused = Submit-GuestsGet $IOStreams $reclone.clone_id
    Assert ($gReused.state -ne 'powered_off') "guests/get for the recycled VMID describes the live clone, not the deleted stub"

    # Name resolution is polled, not asserted instantly: the mock deliberately hides a
    # freshly cloned VM from cluster/resources for 50ms, so an immediate read can
    # legitimately fall back to the "VM-<id>" placeholder.
    $nameDeadline = (Get-Date).AddSeconds(5)
    while ((Get-Date) -lt $nameDeadline -and $gReused.name -ne 'e2e-clone-2') {
        Start-Sleep -Milliseconds 200
        $gReused = Submit-GuestsGet $IOStreams $reclone.clone_id
    }
    Assert ($gReused.name -eq 'e2e-clone-2') "...and reports the NEW clone's name once Proxmox lists it"

    $deadline2 = (Get-Date).AddSeconds(15)
    $t2 = $null
    while ((Get-Date) -lt $deadline2) {
        $t2 = (Submit-TasksGet $IOStreams $reclone.task_id).state
        if ($t2 -eq 'completed') { break }
        Start-Sleep -Milliseconds 500
    }
    Assert ($t2 -eq 'completed') "the recycled-VMID clone task reaches completed"
    $listReused = Submit-GuestsList $IOStreams
    Assert (@($listReused.guests) -contains $cloneId) "guests/list shows the recycled VMID once reported completed, instead of hiding it for 300s"

    Write-Host "`n=== linked clones: cleanup -- free the shared next-id slot ===" -ForegroundColor Cyan
    Submit-GuestsControl $IOStreams $reclone.clone_id 'stop' | Out-Null
    Submit-GuestsControl $IOStreams $reclone.clone_id 'delete' | Out-Null

    Write-Host "`n=== linked clones: guests/snapshots/* invariant ===" -ForegroundColor Cyan
    # LINKED-CLONES-DESIGN.md #2 "The virtual snapshot": exists is true if and only if
    # the guest is presently a native Proxmox template -- nothing is ever created on
    # Proxmox itself.
    Assert ((Submit-GuestsSnapshotsExists $IOStreams '100' 'RAS Template Snapshot') -eq $false) `
        "guests/snapshots/exists is false for a guest that is not yet a Proxmox template"

    Submit-GuestsControl $IOStreams '100' 'stop' | Out-Null
    $stopDeadline = (Get-Date).AddSeconds(10)
    while ((Get-Date) -lt $stopDeadline -and (Submit-GuestsGet $IOStreams '100').state -ne 'powered_off') { Start-Sleep -Milliseconds 200 }
    Assert ((Submit-GuestsGet $IOStreams '100').state -eq 'powered_off') "VM 100 is powered off before templating (RAS's own create-template flow requires this first)"

    # Mirrors Framework Test Kit/Test-CreateTemplate.ps1's own call order: snapshot
    # create, THEN convert -- the snapshot call is additive, not a substitute (see
    # LINKED-CLONES.md #5 "What the test kit adds, and corrects").
    $snapCreate = Submit-GuestsSnapshotsCreate $IOStreams '100' 'RAS Template Snapshot'
    Assert (-not [string]::IsNullOrWhiteSpace($snapCreate.task_id)) "guests/snapshots/create returns a task_id"
    Assert ((Submit-TasksGet $IOStreams $snapCreate.task_id).state -eq 'completed') "guests/snapshots/create's task reports completed immediately -- no real Proxmox call is made"

    $convertResult = Submit-GuestsConvert $IOStreams '100' $true
    $convertDeadline = (Get-Date).AddSeconds(10)
    $convertState = $null
    while ((Get-Date) -lt $convertDeadline) {
        $convertState = (Submit-TasksGet $IOStreams $convertResult.task_id).state
        if ($convertState -eq 'completed') { break }
        Start-Sleep -Milliseconds 200
    }
    Assert ($convertState -eq 'completed') "guests/convert(true) task reaches completed"
    Assert ((Submit-GuestsGet $IOStreams '100').is_template) "VM 100 is now a real Proxmox template"
    Assert (Submit-GuestsSnapshotsExists $IOStreams '100' 'RAS Template Snapshot') `
        "guests/snapshots/exists now true -- the invariant follows the template flag, not the earlier create call"

    $snapDelete = Submit-GuestsSnapshotsDelete $IOStreams '100' 'RAS Template Snapshot'
    Assert (-not [string]::IsNullOrWhiteSpace($snapDelete.task_id)) "guests/snapshots/delete returns a task_id"
    Assert (Submit-GuestsSnapshotsExists $IOStreams '100' 'RAS Template Snapshot') `
        "design: delete is a no-op in practice -- exists stays true, because only guests/convert can make it false"

    $revertThrew = $false
    try { Submit-GuestsSnapshotsRevert $IOStreams '100' 'RAS Template Snapshot' | Out-Null } catch { $revertThrew = $true }
    Assert $revertThrew "guests/snapshots/revert errors rather than silently no-oping -- unreachable at template_method=basic"

    Write-Host "`n=== linked clones: happy path -- clone with a snapshot set from a real template ===" -ForegroundColor Cyan
    $linkedClone = Submit-GuestsClone $IOStreams '100' 'e2e-linked-1' 'RAS Template Snapshot'
    Assert (-not [string]::IsNullOrWhiteSpace($linkedClone.clone_id)) "linked guests/clone still returns a clone_id"
    $linkedId = [string]$linkedClone.clone_id

    $linkedDeadline = (Get-Date).AddSeconds(15)
    $linkedTaskState = $null
    while ((Get-Date) -lt $linkedDeadline) {
        $linkedTaskState = (Submit-TasksGet $IOStreams $linkedClone.task_id).state
        if ($linkedTaskState -eq 'completed') { break }
        Start-Sleep -Milliseconds 300
    }
    Assert ($linkedTaskState -eq 'completed') "linked clone's task reaches completed"
    Assert ((Get-MockVmCloneFull -VmId $linkedId) -eq 0) `
        "linked clone: the provider asked Proxmox for full=0 because the source is a real template"

    Submit-GuestsControl $IOStreams $linkedId 'stop' | Out-Null
    Submit-GuestsControl $IOStreams $linkedId 'delete' | Out-Null

    Write-Host "`n=== linked clones: the trap -- snapshot set but the source is NOT a template ===" -ForegroundColor Cyan
    # LINKED-CLONES-DESIGN.md #1: Proxmox itself would silently full-clone here --
    # no error, no warning. The whole point of the provider-side check is that this
    # must not be silent: full=1 is deliberate and logged, not an accident.
    $trapClone = Submit-GuestsClone $IOStreams '101' 'e2e-trap-1' 'RAS Template Snapshot'
    Assert (-not [string]::IsNullOrWhiteSpace($trapClone.clone_id)) "trap: guests/clone still succeeds"
    $trapId = [string]$trapClone.clone_id

    $trapDeadline = (Get-Date).AddSeconds(15)
    $trapTaskState = $null
    while ((Get-Date) -lt $trapDeadline) {
        $trapTaskState = (Submit-TasksGet $IOStreams $trapClone.task_id).state
        if ($trapTaskState -eq 'completed') { break }
        Start-Sleep -Milliseconds 300
    }
    Assert ($trapTaskState -eq 'completed') "trap: the fallback full clone's task still reaches completed"
    Assert ((Get-MockVmCloneFull -VmId $trapId) -eq 1) `
        "trap: source [101] is not a template, so the provider deliberately fell back to full=1 instead of letting Proxmox silently downgrade"

    $fallbackLogged = (Test-Path $ProviderLog) -and (Select-String -LiteralPath $ProviderLog -Pattern 'falling back to a full clone' -Quiet)
    Assert $fallbackLogged "trap: the fallback decision is logged, not silent"

    Submit-GuestsControl $IOStreams $trapId 'stop' | Out-Null
    Submit-GuestsControl $IOStreams $trapId 'delete' | Out-Null

    Write-Host "`n=== provider/disconnect ===" -ForegroundColor Cyan
    $disc = Submit-Disconnect $IOStreams
    Assert ($disc.message -match 'Session cleared') "provider/disconnect succeeds"
}

try {
    Invoke-ScriptBlock -CommandPath $ProviderSettings.CommandPath -CommandArgs $ProviderSettings.CommandArgs -ScriptBlock ${function:Invoke-Pipeline}
}
finally {
    Stop-Mock
}

Write-Host "`n=== E2E Summary: $pass passed, $fail failed ===" -ForegroundColor $(if ($fail -eq 0) { 'Green' } else { 'Red' })
if ($fail -gt 0) { exit 1 }
