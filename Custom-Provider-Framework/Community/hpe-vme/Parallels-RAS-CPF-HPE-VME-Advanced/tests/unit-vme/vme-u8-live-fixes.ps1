# vme-u8: linked-clone throughput and state fixes (docs/LINKED-CLONES.md):
#   - linked_clones.early_completion: a linked clone is reported on its first tasks/get,
#     held while cloning.max_concurrent_clone_operations clones are still provisioning,
#     and never in a way that gets RAS's next guests/clone refused;
#   - a version pins the template's clone settings (no template GETs per clone; the
#     version keeps its network after the template's changes);
#   - a stop the provider issued is not reversed by a stale listing (power watch), and a
#     running template in maintenance is re-read on its guests/get;
#   - a settings value still at its previous default follows a changed default.
. (Join-Path $PSScriptRoot 'lib.ps1')

function Send-Snap {
    param($P, [string]$Verb, [string]$Id, [string]$Name)
    return (Send-Rpc $P "{`"method`":`"guests/snapshots/$Verb`",`"params`":{`"id`":`"$Id`",`"name`":`"$Name`"}}")
}
function Send-Clone {
    param($P, [string]$Id, [string]$Name, [string]$Version)
    return (Send-Rpc $P "{`"method`":`"guests/clone`",`"params`":{`"id`":`"$Id`",`"name`":`"$Name`",`"snapshot`":`"$Version`",`"is_link_clone`":true}}")
}
function Send-Control {
    param($P, [string]$Id, [string]$Control)
    return (Send-Rpc $P "{`"method`":`"guests/control`",`"params`":{`"id`":`"$Id`",`"control`":`"$Control`"}}")
}
function Get-MockState { return (Invoke-Mock '/_ctl/state' 'GET') }
function Get-StateVersion {
    param($P, [string]$TemplateId, [string]$RasName)
    $doc = Get-Content (Join-Path $P.Dir 'VME-RAS-TemplateVersions.json') -Raw | ConvertFrom-Json
    $t = @($doc.entries | Where-Object { $_.template_id -eq $TemplateId }) | Select-Object -First 1
    return (@($t.versions | Where-Object { $_.ras_name -eq $RasName }) | Select-Object -First 1)
}

Start-VmeMock -Count 10
Invoke-Mock '/_ctl/set?snapshot_delay=1&provision_delay=6&ip_delay=1&action_delay=1&delete_delay=1' | Out-Null
$settings = @{
    capabilities  = @{ template_method = 'versioning' }
    inventory     = @{ list_min_refresh_seconds = 0; watch_refresh_interval_seconds = 2; power_action_hold_seconds = 20 }
    cloning       = @{ max_concurrent_clone_operations = 2; tracked_refresh_interval_seconds = 1 }
    connection    = @{ retry_delay_milliseconds = 10 }
    linked_clones = @{ image_wait_seconds = 10; snapshot_timeout_seconds = 60 }
}
$dir = New-ProviderDir -Name 'u8' -Settings $settings
$p = Start-Provider -Dir $dir
try {
    Connect-Provider $p | Out-Null
    [void](Get-GuestList $p)

    Write-Host "`n=== a version pins the template's clone settings ===" -ForegroundColor Cyan
    $c = Get-Result (Send-Snap $p 'create' '19' 'RAS_TEMPLATE_VERSION_1') 'create v1'
    Assert ((Wait-Task $p $c.task_id -TimeoutSec 30).state -eq 'completed') 'version 1 created'
    $v1 = Get-StateVersion $p '19' 'RAS_TEMPLATE_VERSION_1'
    Assert ($null -ne $v1.clone_inputs -and [string]$v1.clone_inputs.network -eq '5' -and [string]$v1.clone_inputs.plan -eq '23' -and @($v1.clone_inputs.volumes).Count -eq 1) "the state file holds v1's clone settings (network [$($v1.clone_inputs.network)], plan [$($v1.clone_inputs.plan)])"
    Assert ((Get-LogText $p) -match 'Version \[RAS_TEMPLATE_VERSION_1\] of template \[19\] clone settings: plan=23 .*network=5') 'the pinned settings are logged'
    [void](Send-Rpc $p '{"method":"guests/convert","params":{"id":"19","is_template":true}}')
    # The template is moved to another VLAN afterwards (a later maintenance would do this).
    Invoke-Mock '/_ctl/instance?id=19&srv_net_id=9&srv_net_name=VLAN-9' | Out-Null

    Write-Host "`n=== early completion: first tasks/get, no GETs ===" -ForegroundColor Cyan
    Reset-MockStats
    $t0 = Get-Date
    $a = Get-Result (Send-Clone $p '19' 'LC-A' 'RAS_TEMPLATE_VERSION_1') 'clone LC-A'
    $ta = Get-Task $p $a.task_id
    $secsA = ((Get-Date) - $t0).TotalSeconds
    $st = Get-MockStats
    Assert ($ta.state -eq 'completed' -and $ta.output.clone_id) "LC-A is reported completed on its first tasks/get ($([Math]::Round($secsA, 1)) s, VME still provisioning for 6 s)"
    Assert ((Get-MockCount $st 'POST /api/instances list') -eq 1 -and (Get-MockTotal $st) -eq 1) "submit + first tasks/get cost exactly one VME call, the POST (got $(Get-MockTotal $st): no template instance/server GET, no status GET)"
    Assert ([string](Get-MockState).last_create_body.networkInterfaces[0].network.id -eq '5') "the clone gets version 1's network [5], not the template's new one [9]"
    $aId = [string]$ta.output.clone_id
    Assert ((Get-LogText $p) -match 'Linked clone \[LC-A\] \(id=\d+\) reported completed at \d+s, VME status \[provisioning\]') 'the early report is logged with the VME status'
    $ga = Get-Result (Get-Guest $p $aId) 'get LC-A'
    Assert ($ga.state -eq 'powering_on') "while VME provisions it, guests/get reports powering_on (got $($ga.state))"

    Write-Host "`n=== held at the clone limit, never refused ===" -ForegroundColor Cyan
    # LC-B provisions longer than LC-A, so only the slot LC-A frees can release it.
    Invoke-Mock '/_ctl/set?provision_delay=14' | Out-Null
    $b = Send-Clone $p '19' 'LC-B' 'RAS_TEMPLATE_VERSION_1'
    Invoke-Mock '/_ctl/set?provision_delay=6' | Out-Null
    Assert (-not (Test-IsError $b)) 'RAS''s next guests/clone is accepted (1 of 2 slots in use)'
    $tb0 = Get-Date
    $tb = Get-Task $p $b.result.task_id
    Assert ($tb.state -eq 'running') "LC-B's first tasks/get is held: LC-A + LC-B fill the 2 slots (got $($tb.state))"
    Assert ((Get-LogText $p) -match 'Linked clone \[LC-B\] .*held at \d+s: 2 clone\(s\) still provisioning, limit 2') 'the hold is logged once'
    $tbDone = Wait-Task $p $b.result.task_id -TimeoutSec 30 -IntervalSec 0.5
    $secsB = ((Get-Date) - $tb0).TotalSeconds
    Assert ($null -ne $tbDone -and $tbDone.state -eq 'completed' -and $secsB -ge 3 -and $secsB -lt 12) "LC-B completes once LC-A finishes provisioning, while itself still provisioning ($([Math]::Round($secsB, 1)) s)"
    $log = Get-LogText $p
    Assert ($log -match "Clone \[LC-A\] \(instance \[$aId\]\) finished provisioning \(status=running\)" -and $log -match 'Linked clone \[LC-B\] \(id=\d+\) reported completed at \d+s, VME status \[provisioning\]') 'the held tasks/get itself re-read LC-A, freed its slot and reported LC-B early'
    $cC = Send-Clone $p '19' 'LC-C' 'RAS_TEMPLATE_VERSION_1'
    Assert (-not (Test-IsError $cC)) 'the guests/clone right after a completion is accepted'
    Assert ((Wait-Task $p $cC.result.task_id -TimeoutSec 40 -IntervalSec 0.5).state -eq 'completed') 'LC-C completes too'
    Assert ((Get-LogText $p) -notmatch 'Clone limit reached') 'no guests/clone was ever refused'

    Write-Host "`n=== a new version pins the new network ===" -ForegroundColor Cyan
    [void](Send-Rpc $p '{"method":"guests/convert","params":{"id":"19","is_template":false}}')
    $c2 = Get-Result (Send-Snap $p 'create' '19' 'RAS_TEMPLATE_VERSION_2') 'create v2'
    Assert ((Wait-Task $p $c2.task_id -TimeoutSec 30).state -eq 'completed') 'version 2 created'
    [void](Send-Rpc $p '{"method":"guests/convert","params":{"id":"19","is_template":true}}')
    Assert ([string](Get-StateVersion $p '19' 'RAS_TEMPLATE_VERSION_2').clone_inputs.network -eq '9') 'version 2 pins network [9]'
    Start-Sleep -Seconds 7   # let LC-C finish so a slot is free
    $d = Get-Result (Send-Clone $p '19' 'LC-D' 'RAS_TEMPLATE_VERSION_2') 'clone LC-D'
    [void](Wait-Task $p $d.task_id -TimeoutSec 30 -IntervalSec 0.5)
    Assert ([string](Get-MockState).last_create_body.networkInterfaces[0].network.id -eq '9') 'a version 2 clone gets network [9]'

    Write-Host "`n=== version without pinned settings (state file from 4.1) ===" -ForegroundColor Cyan
    Stop-Provider $p
    $f = Join-Path $dir 'VME-RAS-TemplateVersions.json'
    $doc = Get-Content $f -Raw | ConvertFrom-Json
    foreach ($e in $doc.entries) { foreach ($v in $e.versions) { $v.clone_inputs = $null } }
    Set-Content -LiteralPath $f -Value ($doc | ConvertTo-Json -Depth 12) -Encoding utf8NoBOM
    $p = Start-Provider -Dir $dir
    Connect-Provider $p | Out-Null
    [void](Get-GuestList $p)
    Start-Sleep -Seconds 7
    Reset-MockStats
    $e1 = Get-Result (Send-Clone $p '19' 'LC-E' 'RAS_TEMPLATE_VERSION_1') 'clone LC-E'
    [void](Wait-Task $p $e1.task_id -TimeoutSec 30 -IntervalSec 0.5)
    Assert ((Get-LogText $p) -match 'Version \[RAS_TEMPLATE_VERSION_1\] of template \[19\] had no clone settings -- taken from the template now') 'the first clone takes them (W line)'
    Assert ($null -ne (Get-StateVersion $p '19' 'RAS_TEMPLATE_VERSION_1').clone_inputs) '...and keeps them in the state file'
    Start-Sleep -Seconds 7
    Reset-MockStats
    $e2 = Get-Result (Send-Clone $p '19' 'LC-F' 'RAS_TEMPLATE_VERSION_1') 'clone LC-F'
    [void](Get-Task $p $e2.task_id)
    Assert ((Get-MockCount (Get-MockStats) 'GET /api/servers/{id}') -eq 0) 'the next clone of that version reads no server object'

    Write-Host "`n=== power watch: a stale listing does not undo a stop ===" -ForegroundColor Cyan
    Invoke-Mock '/_ctl/set?stop_lag=6' | Out-Null
    $g = Get-Result (Get-Guest $p '20') 'get 20'
    Assert ($g.state -eq 'powered_on') 'instance 20 starts powered on'
    [void](Get-Result (Send-Control $p '20' 'stop') 'stop 20')
    $g = Get-Result (Get-Guest $p '20') 'get 20 after stop'
    Assert ($g.state -eq 'powering_off') "right after the stop: powering_off ($($g.power_state))"
    Start-Sleep -Milliseconds 800
    [void](Get-GuestList $p)   # VME now lists it as running again
    $g = Get-Result (Get-Guest $p '20') 'get 20 after a stale listing'
    Assert ($g.state -eq 'powering_off' -and $g.power_state -eq 'running') "the listing says running, guests/get keeps powering_off (power_state=$($g.power_state))"
    Assert ((Get-LogText $p) -match 'Power watch: VME still shows instance \[20\] as \[running\]') 'the hold is logged'
    Reset-MockStats
    for ($i = 0; $i -lt 5; $i++) { [void](Get-Guest $p '20') }
    Assert ((Get-MockCount (Get-MockStats) 'GET /api/instances/{id}') -le 1) "re-reads are throttled to watch_refresh_interval_seconds ($(Get-MockCount (Get-MockStats) 'GET /api/instances/{id}') GET for 5 guests/get)"
    $deadline = (Get-Date).AddSeconds(15)
    do { Start-Sleep -Seconds 1; $g = Get-Result (Get-Guest $p '20') 'get 20 while waiting' } while ($g.state -ne 'powered_off' -and (Get-Date) -lt $deadline)
    Assert ($g.state -eq 'powered_off') 'once VME shows it stopped, guests/get reports powered_off -- without a listing'
    Assert ((Get-LogText $p) -match 'Power watch: instance \[20\] is \[powered_off\] \d+s after \[stop\]') 'the watch ends with a T line'
    Invoke-Mock '/_ctl/set?stop_lag=0' | Out-Null

    Write-Host "`n=== a running template in maintenance is re-read ===" -ForegroundColor Cyan
    [void](Send-Rpc $p '{"method":"guests/convert","params":{"id":"19","is_template":false}}')
    Invoke-Mock '/_ctl/instance?id=19&status=running' | Out-Null
    Invoke-Mock '/_ctl/instance?id=100&status=running' | Out-Null
    [void](Get-GuestList $p)
    Assert ((Get-Result (Get-Guest $p '19') 'get 19').state -eq 'powered_on') 'the template in maintenance runs'
    # RAS shuts it down through its guest agent: the provider is not told.
    Invoke-Mock '/_ctl/instance?id=19&status=stopped' | Out-Null
    Invoke-Mock '/_ctl/instance?id=100&status=stopped' | Out-Null
    Start-Sleep -Seconds 3
    Reset-MockStats
    $g19 = Get-Result (Get-Guest $p '19') 'get 19 after agent shutdown'
    $g100 = Get-Result (Get-Guest $p '100') 'get 100'
    $st = Get-MockStats
    Assert ($g19.state -eq 'powered_off') 'the template in maintenance shows powered_off on its next guests/get, before any listing'
    Assert ($g100.state -eq 'powered_on' -and (Get-MockCount $st 'GET /api/instances/{id}') -eq 1) 'an ordinary guest is still served from the inventory (one GET in total, for the template)'
    [void](Send-Rpc $p '{"method":"guests/convert","params":{"id":"19","is_template":true}}')

    Write-Host "`n=== a value at its previous default follows a changed default ===" -ForegroundColor Cyan
    $raw = @'
{
  "schema_version": 2,
  "cloning": {
    "max_concurrent_clone_operations": { "value": 2, "default": 2, "description": "old" },
    "resolve_timeout_seconds": { "value": 600, "default": 900, "description": "admin-set" }
  },
  "delete": { "max_probes_per_list": { "value": 40, "default": 20, "description": "admin-set, old default" } }
}
'@
    $dirS = New-ProviderDir -Name 'u8-defaults' -RawSettingsJson $raw
    $ps = Start-Provider -Dir $dirS
    [void](Send-Rpc $ps '{"method":"provider/initialize"}')
    $cfg = Get-Content (Join-Path $dirS 'VME-RAS-Config.json') -Raw | ConvertFrom-Json
    Assert ($cfg.cloning.max_concurrent_clone_operations.value -eq 10 -and $cfg.cloning.max_concurrent_clone_operations.default -eq 10) "2 (= its recorded old default) follows the new default 10 (file now value=$($cfg.cloning.max_concurrent_clone_operations.value))"
    Assert ($cfg.cloning.resolve_timeout_seconds.value -eq 600) 'an admin value with an unchanged default is kept (600)'
    Assert ($cfg.delete.max_probes_per_list.value -eq 40 -and $cfg.delete.max_probes_per_list.default -eq 25) 'an admin value whose default changed is kept (40), its recorded default refreshed (25)'
    $log = Get-LogText $ps
    Assert ($log -match 'cloning\.max_concurrent_clone_operations was at its previous default \[2\] -> new default \[10\] applied' -and $log -match 'max_concurrent=10') 'the rewrite lists the followed default and the Cloning line shows 10'
    Stop-Provider $ps
}
catch { Assert $false "suite aborted: $($_.Exception.Message) at $($_.InvocationInfo.PositionMessage)" }
finally { Complete-Suite 'vme-u8' }
