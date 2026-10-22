# vme-u7: template versioning + linked clones on VME 9.1 (docs/LINKED-CLONES.md).
# Walks the Framework Test Kit's 'versioning' sequences against the mock: create
# template (snapshot -> linked-clone image), clone (POST /api/instances), enter
# maintenance (revert), exit maintenance (new version), delete a version (guarded,
# then for real), recovery from VME after a lost state file, and the failure paths.
. (Join-Path $PSScriptRoot 'lib.ps1')

function Send-Convert {
    param($P, [string]$Id, [bool]$IsTemplate)
    return (Get-Result (Send-Rpc $P "{`"method`":`"guests/convert`",`"params`":{`"id`":`"$Id`",`"is_template`":$($IsTemplate.ToString().ToLowerInvariant())}}") 'convert')
}
function Send-Snap {
    param($P, [string]$Verb, [string]$Id, [string]$Name)
    return (Send-Rpc $P "{`"method`":`"guests/snapshots/$Verb`",`"params`":{`"id`":`"$Id`",`"name`":`"$Name`"}}")
}
function Send-Clone {
    param($P, [string]$Id, [string]$Name, [string]$Extra = '')
    return (Send-Rpc $P "{`"method`":`"guests/clone`",`"params`":{`"id`":`"$Id`",`"name`":`"$Name`"$Extra}}")
}
function Get-MockState { return (Invoke-Mock '/_ctl/state' 'GET') }
function Get-MockImagesNamed {
    param([string]$Name)
    return , @((Get-MockState).images.PSObject.Properties | Where-Object { $_.Value.name -eq $Name })
}
function Get-VersionState {
    param($P)
    $f = Join-Path $P.Dir 'VME-RAS-TemplateVersions.json'
    if (-not (Test-Path $f)) { return $null }
    return (Get-Content $f -Raw | ConvertFrom-Json)
}
function Get-StateVersion {
    param($P, [string]$TemplateId, [string]$RasName)
    $doc = Get-VersionState $P
    if ($null -eq $doc) { return $null }
    $t = @($doc.entries | Where-Object { $_.template_id -eq $TemplateId }) | Select-Object -First 1
    if ($null -eq $t) { return $null }
    return (@($t.versions | Where-Object { $_.ras_name -eq $RasName }) | Select-Object -First 1)
}

Start-VmeMock -Count 20
Invoke-Mock '/_ctl/set?snapshot_delay=1&provision_delay=2&ip_delay=1&delete_delay=1&action_delay=1' | Out-Null
$settings = @{
    capabilities  = @{ template_method = 'versioning' }
    inventory     = @{ list_min_refresh_seconds = 0 }
    delete        = @{ probe_interval_seconds = 1 }
    connection    = @{ retry_delay_milliseconds = 10 }
    linked_clones = @{ image_wait_seconds = 10; snapshot_timeout_seconds = 60 }
}
$dir = New-ProviderDir -Name 'u7' -Settings $settings
$p = Start-Provider -Dir $dir
try {
    Write-Host "`n=== initialize advertises versioning + linked clones ===" -ForegroundColor Cyan
    $init = Get-Result (Send-Rpc $p '{"method":"provider/initialize"}') 'initialize'
    Assert ($init.capabilities.template_method -eq 'versioning' -and $init.capabilities.can_link_clones -eq $true) "template_method=versioning, can_link_clones=true (got $($init.capabilities.template_method)/$($init.capabilities.can_link_clones))"
    Connect-Provider $p | Out-Null
    Assert ((Get-LogText $p) -match 'VME build \[9\.1\.0\]') 'connect logs the VME build (GET /api/ping)'
    [void](Get-GuestList $p)

    Write-Host "`n=== create template: snapshot -> linked-clone image (version 1) ===" -ForegroundColor Cyan
    Reset-MockStats
    $c = Get-Result (Send-Snap $p 'create' '19' 'RAS_TEMPLATE_VERSION_1') 'snapshots/create'
    Assert ($c.task_id -like 'snapcreate:19:*') "snapshots/create returns a snapcreate task ($($c.task_id))"
    $t = Wait-Task $p $c.task_id -TimeoutSec 30
    Assert ($null -ne $t -and $t.state -eq 'completed') "the version task completes (state=$($t.state))"
    $st = Get-MockStats
    Assert ((Get-MockCount $st 'PUT /api/instances/{id}/snapshot') -eq 1 -and (Get-MockCount $st 'PUT /api/instances/{id}/linked-clone/{id}') -eq 1) 'one snapshot PUT and one linked-clone PUT'
    $ms = Get-MockState
    $snap = @($ms.snapshots.PSObject.Properties | Where-Object { $_.Value.name -eq 'RAS_TEMPLATE_VERSION_1_19' })
    Assert ($snap.Count -eq 1 -and $snap[0].Value.status -eq 'complete') 'the VME snapshot is named RAS_TEMPLATE_VERSION_1_19 (unique across templates)'
    $img = Get-MockImagesNamed 'RAS_TEMPLATE_VERSION_1_19'
    Assert ($img.Count -eq 1 -and @($img[0].Value.labels) -contains 'ras-template-19' -and @($img[0].Value.labels) -contains 'ras-version-1') 'one image with the same name, labelled ras-template-19 / ras-version-1'
    $v1 = Get-StateVersion $p '19' 'RAS_TEMPLATE_VERSION_1'
    Assert ($null -ne $v1 -and $v1.status -eq 'ready' -and $v1.image_id -eq $img[0].Name -and $v1.snapshot_id -eq $snap[0].Name) "the state file holds version 1 ready with snapshot [$($v1.snapshot_id)] and image [$($v1.image_id)]"
    Assert ((Get-LogText $p) -match 'Image \[\d+\] is on snapshot') 'the image is verified against the snapshot externalId'
    Assert ((Get-Task $p $c.task_id).state -eq 'completed') 'a repeated tasks/get stays completed'
    Send-Convert $p '19' $true | Out-Null
    Assert ((Get-Result (Get-Guest $p '19') 'get 19').is_template -eq $true) 'convert(true) -> is_template'

    Write-Host "`n=== snapshots/exists ===" -ForegroundColor Cyan
    $e1 = Send-Snap $p 'exists' '19' 'RAS_TEMPLATE_VERSION_1'
    $e2 = Send-Snap $p 'exists' '19' 'RAS_TEMPLATE_VERSION_2'
    Assert ($e1.result -eq $true -and $e2.result -eq $false) "exists is a bare boolean: v1=$($e1.result) v2=$($e2.result)"

    Write-Host "`n=== linked clone: POST /api/instances from the version image ===" -ForegroundColor Cyan
    Reset-MockStats
    $cl = Get-Result (Send-Clone $p '19' 'LC-001' ',"snapshot":"RAS_TEMPLATE_VERSION_1"') 'clone LC-001'
    Assert ($cl.task_id -like 'linked:19:LC-001:*') "the clone task is linked:* ($($cl.task_id))"
    $st = Get-MockStats
    Assert ((Get-MockCount $st 'POST /api/instances list') -eq 1 -and (Get-MockCount $st 'PUT /api/instances/{id}/clone') -eq 0) 'one POST /api/instances, no PUT /clone'
    $body = (Get-MockState).last_create_body
    Assert ([string]$body.config.imageId -eq [string]$v1.image_id) "config.imageId = the version's image [$($body.config.imageId)]"
    Assert ($body.config.uefi -eq $true -and $body.config.secureBoot -eq $true -and $body.config.tpm -eq $true -and $body.config.resourcePoolId -eq 'pool-1') 'firmware flags and resource pool copied from the template'
    Assert ($body.instance.type -eq 'mvm' -and $body.instance.plan.id -eq 23 -and $body.instance.layout.id -eq 32 -and $body.zoneId -eq 1 -and $body.servicePlanOptions.maxCores -eq 2) 'instance type, plan, layout, cloud and size copied'
    Assert ($body.networkInterfaces[0].network.id -eq 5 -and $body.volumes[0].size -eq 60) 'network (server object) and disk size copied'
    $newId = @((Get-MockState).instances.PSObject.Properties | Where-Object { $_.Value.name -eq 'LC-001' })[0].Name
    Assert ((Get-GuestList $p) -notcontains $newId) "the new instance [$newId] is hidden from guests/list until tasks/get reports it"
    $t = Wait-Task $p $cl.task_id -TimeoutSec 30
    Assert ($t.state -eq 'completed' -and [string]$t.output.clone_id -eq $newId) "the task completes with clone_id [$($t.output.clone_id)] -- no resolution by name"
    Assert ((Get-LogText $p) -notmatch 'Clone \[LC-001\] resolved') 'no name resolution happened'
    Assert ((Get-GuestList $p) -contains $newId) 'listed once reported'
    $g = Get-Result (Get-Guest $p $newId) 'get clone'
    Assert ($g.name -eq 'LC-001' -and $g.state -in @('powered_on', 'powering_on')) "guests/get on the clone: $($g.state)"
    $dup = Send-Clone $p '19' 'LC-001' ',"snapshot":"RAS_TEMPLATE_VERSION_1"'
    Assert (-not (Test-IsError $dup)) 'a later same-name clone is accepted (the first one was reported)'
    [void](Wait-Task $p $dup.result.task_id -TimeoutSec 30)

    Write-Host "`n=== clone refusals ===" -ForegroundColor Cyan
    $full = Send-Clone $p '19' 'LC-FULL' ',"snapshot":"RAS_TEMPLATE_VERSION_1","is_link_clone":false'
    Assert ((Test-IsError $full) -and $full.error.code -eq -32602 -and $full.error.message -match 'Full clones of a template version') 'is_link_clone=false is refused (LV2)'
    $none = Send-Clone $p '19' 'LC-NOVER' ',"snapshot":"RAS_TEMPLATE_VERSION_9"'
    Assert ((Test-IsError $none) -and $none.error.message -match 'no version \[RAS_TEMPLATE_VERSION_9\]') 'an unknown version is refused with its VME name'

    Write-Host "`n=== enter maintenance: revert to version 1 ===" -ForegroundColor Cyan
    Send-Convert $p '19' $false | Out-Null
    $r = Get-Result (Send-Snap $p 'revert' '19' 'RAS_TEMPLATE_VERSION_1') 'revert'
    Assert ((Get-Task $p $r.task_id).state -eq 'completed') "revert is synchronous ($($r.task_id))"
    Assert ((Get-MockState).last_revert.name -eq 'RAS_TEMPLATE_VERSION_1_19') 'VME reverted to RAS_TEMPLATE_VERSION_1_19'

    Write-Host "`n=== exit maintenance: version 2 ===" -ForegroundColor Cyan
    Assert ((Send-Snap $p 'exists' '19' 'RAS_TEMPLATE_VERSION_2').result -eq $false) 'exists(v2) is false, so RAS deletes nothing'
    $c2 = Get-Result (Send-Snap $p 'create' '19' 'RAS_TEMPLATE_VERSION_2') 'create v2'
    Assert ((Wait-Task $p $c2.task_id -TimeoutSec 30).state -eq 'completed') 'version 2 created'
    Send-Convert $p '19' $true | Out-Null
    Assert ((Get-MockImagesNamed 'RAS_TEMPLATE_VERSION_1_19').Count -eq 1 -and (Get-MockImagesNamed 'RAS_TEMPLATE_VERSION_2_19').Count -eq 1) 'both versions keep their own image: existing desktops are untouched'
    $again = Get-Result (Send-Snap $p 'create' '19' 'RAS_TEMPLATE_VERSION_2') 'create v2 again'
    Assert ($again.task_id -like 'noop:*' -and (Get-MockImagesNamed 'RAS_TEMPLATE_VERSION_2_19').Count -eq 1) 'creating an existing version again is a no-op (no second image)'

    Write-Host "`n=== delete a version: guarded while desktops use it ===" -ForegroundColor Cyan
    $d = Send-Snap $p 'delete' '19' 'RAS_TEMPLATE_VERSION_1'
    Assert ((Test-IsError $d) -and $d.error.message -match 'still used by 2 guest') "refused while 2 desktops use version 1: $($d.error.message)"
    Assert (@((Get-MockState).snapshots.PSObject.Properties | Where-Object { $_.Value.name -eq 'RAS_TEMPLATE_VERSION_1_19' }).Count -eq 1) 'nothing deleted on VME'
    foreach ($cid in @((Get-MockState).instances.PSObject.Properties | Where-Object { $_.Value.name -eq 'LC-001' } | ForEach-Object { $_.Name })) {
        [void](Send-Rpc $p "{`"method`":`"guests/control`",`"params`":{`"id`":`"$cid`",`"control`":`"delete`"}}")
    }
    $deadline = (Get-Date).AddSeconds(20)
    while ((Get-Date) -lt $deadline) { $l = Get-GuestList $p; if (@((Get-MockState).instances.PSObject.Properties | Where-Object { $_.Value.name -eq 'LC-001' }).Count -eq 0 -and $l.Count -gt 0) { break }; Start-Sleep -Milliseconds 500 }
    [void](Get-GuestList $p); [void](Get-GuestList $p)
    $d = Get-Result (Send-Snap $p 'delete' '19' 'RAS_TEMPLATE_VERSION_1') 'delete v1'
    Assert ($d.task_id -like 'snapdelete:19:*') 'once the desktops are gone the delete is accepted'
    $t = Wait-Task $p $d.task_id -TimeoutSec 20
    Assert ($t.state -eq 'completed') 'the delete task completes when VME no longer lists the snapshot'
    Assert ((Get-MockImagesNamed 'RAS_TEMPLATE_VERSION_1_19').Count -eq 0) 'the image went with the snapshot'
    Assert ($null -eq (Get-StateVersion $p '19' 'RAS_TEMPLATE_VERSION_1') -and $null -ne (Get-StateVersion $p '19' 'RAS_TEMPLATE_VERSION_2')) 'version 1 left the state file, version 2 stayed'
    Assert ((Get-LogText $p) -match 'its image \[\d+\] is gone too') 'the cascade is confirmed in the log'
    $d2 = Get-Result (Send-Snap $p 'delete' '19' 'RAS_TEMPLATE_VERSION_1') 'delete v1 again'
    Assert ($d2.task_id -like 'noop:*') 'deleting a version VME no longer has is a no-op success'

    Write-Host "`n=== lost state file: version rebuilt from VME by name ===" -ForegroundColor Cyan
    Stop-Provider $p
    Remove-Item (Join-Path $dir 'VME-RAS-TemplateVersions.json') -Force
    $p = Start-Provider -Dir $dir
    Connect-Provider $p | Out-Null
    [void](Get-GuestList $p)
    $cl2 = Get-Result (Send-Clone $p '19' 'LC-002' ',"snapshot":"RAS_TEMPLATE_VERSION_2"') 'clone from v2 after state loss'
    Assert ((Wait-Task $p $cl2.task_id -TimeoutSec 30).state -eq 'completed') 'the clone succeeds'
    Assert ((Get-LogText $p) -match 'Version \[RAS_TEMPLATE_VERSION_2\] of template \[19\] rebuilt from VME') 'version 2 was rebuilt from VME (snapshot + image by name)'

    Write-Host "`n=== failures ===" -ForegroundColor Cyan
    Invoke-Mock '/_ctl/set?linked_clone_fails=1' | Out-Null
    $c3 = Get-Result (Send-Snap $p 'create' '19' 'RAS_TEMPLATE_VERSION_3') 'create v3'
    $t = Wait-Task $p $c3.task_id -TimeoutSec 30
    Assert ($t.state -eq 'failed' -and $t.error.message -match 'linked clone of snapshot \[\d+\] failed: .*HTTP 400.*VME 9\.1 or later') "a refused linked clone (9.0.x stub: bare HTTP 400) fails the version: $($t.error.message)"
    Assert ($null -eq (Get-StateVersion $p '19' 'RAS_TEMPLATE_VERSION_3')) 'the failed version left the state'
    Invoke-Mock '/_ctl/set?linked_clone_fails=0' | Out-Null
    Reset-MockStats
    $c3b = Get-Result (Send-Snap $p 'create' '19' 'RAS_TEMPLATE_VERSION_3') 'create v3 again'
    Assert ((Wait-Task $p $c3b.task_id -TimeoutSec 30).state -eq 'completed') 'RAS retrying the step succeeds'
    Assert ((Get-MockCount (Get-MockStats) 'PUT /api/instances/{id}/snapshot') -eq 0 -and (Get-LogText $p) -match 'reusing VME snapshot') 'the leftover snapshot was reused, not duplicated'

    Invoke-Mock '/_ctl/set?snapshot_fails=1' | Out-Null
    $c4 = Get-Result (Send-Snap $p 'create' '19' 'RAS_TEMPLATE_VERSION_4') 'create v4'
    $t = Wait-Task $p $c4.task_id -TimeoutSec 30
    Assert ($t.state -eq 'failed' -and $t.error.message -match 'snapshot process \[\d+\] ended \[failed\]') "a failed snapshot process fails the version: $($t.error.message)"
    Invoke-Mock '/_ctl/set?snapshot_fails=0&image_bucket=User' | Out-Null
    $c5 = Get-Result (Send-Snap $p 'create' '19' 'RAS_TEMPLATE_VERSION_5') 'create v5'
    $t = Wait-Task $p $c5.task_id -TimeoutSec 40
    Assert ($t.state -eq 'failed' -and $t.error.message -match 'no virtual image named \[RAS_TEMPLATE_VERSION_5_19\] \(filterType=synced\)') "an image outside the filterType bucket times out with a clear reason: $($t.error.message)"
    Invoke-Mock '/_ctl/set?image_bucket=Synced&image_links_snapshot=0' | Out-Null
    $c6 = Get-Result (Send-Snap $p 'create' '19' 'RAS_TEMPLATE_VERSION_6') 'create v6'
    Assert ((Wait-Task $p $c6.task_id -TimeoutSec 30).state -eq 'completed') 'a unique name alone is enough...'
    Assert ((Get-LogText $p) -match 'taken by its unique name only') '...with a W line that the externalDiskId link did not hold'

    Write-Host "`n=== basic stays full-clone only ===" -ForegroundColor Cyan
    $dirB = New-ProviderDir -Name 'u7b' -Settings @{ inventory = @{ list_min_refresh_seconds = 0 } }
    $pb = Start-Provider -Dir $dirB
    $initB = Get-Result (Send-Rpc $pb '{"method":"provider/initialize"}') 'initialize basic'
    Assert ($initB.capabilities.template_method -eq 'basic' -and $initB.capabilities.can_link_clones -eq $false) 'basic: can_link_clones=false'
    Connect-Provider $pb | Out-Null
    $sb = Send-Snap $pb 'create' '19' 'RAS_TEMPLATE_VERSION_1'
    Assert ((Test-IsError $sb) -and $sb.error.code -eq -32601) 'basic: guests/snapshots/* stay -32601'
}
catch { Assert $false "suite aborted: $($_.Exception.Message) at $($_.InvocationInfo.PositionMessage)" }
finally { Complete-Suite 'vme-u7' }
