# vme-u4: delete, power guardrails, template maintenance (review 5, 8.2 vme-u5/u6).
. (Join-Path $PSScriptRoot 'lib.ps1')

function Send-Control {
    param($P, [string]$Id, [string]$Control)
    return (Send-Rpc $P "{`"method`":`"guests/control`",`"params`":{`"id`":`"$Id`",`"control`":`"$Control`"}}")
}
function Send-Convert {
    param($P, [string]$Id, [bool]$IsTemplate)
    return (Get-Result (Send-Rpc $P "{`"method`":`"guests/convert`",`"params`":{`"id`":`"$Id`",`"is_template`":$($IsTemplate.ToString().ToLowerInvariant())}}") 'convert')
}

Start-VmeMock -Count 40
Invoke-Mock '/_ctl/set?delete_delay=2&action_delay=1&clone_delay=1&provision_delay=30' | Out-Null
$settings = @{
    inventory  = @{ list_min_refresh_seconds = 0 }
    delete     = @{ probe_interval_seconds = 1 }
    connection = @{ retry_delay_milliseconds = 10 }
}
$dir = New-ProviderDir -Name 'u4' -Settings $settings
$p = Start-Provider -Dir $dir
try {
    Connect-Provider $p | Out-Null
    [void](Get-GuestList $p)

    Write-Host "`n=== delete: one call, truthful visibility ===" -ForegroundColor Cyan
    Reset-MockStats
    $t0 = Get-Date
    $d = Send-Control $p '100' 'delete'
    $secs = ((Get-Date) - $t0).TotalSeconds
    $st = Get-MockStats
    Assert (-not (Test-IsError $d) -and $secs -lt 1) "delete returns at once ($([Math]::Round($secs,2))s)"
    Assert ((Get-MockCount $st 'DELETE /api/instances/{id}') -eq 1 -and (Get-MockTotal $st) -eq 1) "delete costs exactly one DELETE, no pre-GET, no pre-stop (got $(Get-MockTotal $st) calls)"
    $list = Get-GuestList $p
    Assert ($list -contains '100') 'right after the delete the guest is still listed (VME still has it)'
    $g = Get-Result (Get-Guest $p '100') 'get deleting'
    Assert ($g.state -eq 'powering_off') "a deleting guest reports powering_off (got $($g.state))"
    $gone = $false; $deadline = (Get-Date).AddSeconds(15)
    while ((Get-Date) -lt $deadline) { if ((Get-GuestList $p) -notcontains '100') { $gone = $true; break }; Start-Sleep -Milliseconds 500 }
    Assert $gone 'the guest leaves guests/list once VME returns 404 (per-id probe, no full re-listing needed)'
    Assert ((Get-LogText $p) -match 'Instance \[100\] \(filler-100\) is gone') 'the confirmed removal is logged'
    $again = Send-Control $p '100' 'delete'
    Assert (-not (Test-IsError $again) -and $again.result.message -match 'already deleted') 'deleting an id VME no longer has is a success (404 = gone)'

    Write-Host "`n=== delete: transient failures ===" -ForegroundColor Cyan
    Invoke-Mock '/_ctl/fail?key=DELETE%20/api/instances/{id}&count=1&code=500' | Out-Null
    Reset-MockStats
    $r = Send-Control $p '101' 'delete'
    Assert (-not (Test-IsError $r) -and (Get-MockCount (Get-MockStats) 'DELETE /api/instances/{id}') -eq 2) 'one 5xx on DELETE is retried once (idempotent)'
    Invoke-Mock '/_ctl/fail?key=DELETE%20/api/instances/{id}&count=2&code=500' | Out-Null
    $r = Send-Control $p '102' 'delete'
    Assert (Test-IsError $r) 'a DELETE failing twice is reported to RAS as an error, so RAS can retry'
    Invoke-Mock '/_ctl/fail?key=DELETE%20/api/instances/{id}&count=2&code=0' | Out-Null
    $r = Send-Control $p '103' 'delete'
    Assert (-not (Test-IsError $r)) 'a DELETE with no response at all is treated as submitted (the probe decides)'
    Assert ((Get-LogText $p) -match 'Delete of instance \[103\] got no response') 'and says so in the log'

    Write-Host "`n=== delete guards ===" -ForegroundColor Cyan
    Send-Convert $p '19' $true | Out-Null
    $r = Send-Control $p '19' 'delete'
    $mock19 = (Invoke-Mock '/_ctl/state' 'GET').instances.'19'
    Assert (-not (Test-IsError $r) -and $r.result.message -match 'template designation cleared' -and $null -ne $mock19 -and $mock19.status -ne 'removing') 'delete on a template clears the designation and leaves the VM alone'
    Assert ((Get-Result (Get-Guest $p '19') 'get 19').is_template -eq $false) 'the former template is an ordinary guest now'
    $r2 = Send-Control $p '19' 'delete'
    $mock19 = (Invoke-Mock '/_ctl/state' 'GET').instances.'19'
    Assert (-not (Test-IsError $r2) -and $r2.result.message -match 'template designation' -and $mock19.status -ne 'removing') 'a RAS retry of the template delete is refused too (the VM survives)'
    $former = Get-Content (Join-Path $p.Dir 'VME-RAS-FormerTemplates.json') -Raw | ConvertFrom-Json
    Assert (@($former.entries | ForEach-Object { $_.id }) -contains '19') 'the former template is persisted, so the refusal has no time limit and survives a restart (Q11)'
    Send-Convert $p '19' $true | Out-Null
    $former = Get-Content (Join-Path $p.Dir 'VME-RAS-FormerTemplates.json') -Raw | ConvertFrom-Json
    Assert (@($former.entries | ForEach-Object { $_.id }) -notcontains '19') 're-adding it as a template moves it back to the template set'
    $r3 = Send-Control $p '19' 'delete'
    $mock19 = (Invoke-Mock '/_ctl/state' 'GET').instances.'19'
    Assert (-not (Test-IsError $r3) -and $r3.result.message -match 'template designation cleared' -and $mock19.status -ne 'removing') 'deleting it as a template again still only clears the designation'

    # A clone from a non-template source (allowed, with a W), still provisioning.
    $c = Get-Result (Send-Rpc $p '{"method":"guests/clone","params":{"id":"20","name":"SRC-GUARD"}}') 'clone'
    $r = Send-Control $p '20' 'delete'
    Assert ((Test-IsError $r) -and $r.error.message -match 'source of 1 clone') 'the source of a clone still provisioning cannot be deleted'

    Write-Host "`n=== power guardrails ===" -ForegroundColor Cyan
    Invoke-Mock '/_ctl/instance?id=130&status=provisioning&busy=1' | Out-Null
    [void](Get-GuestList $p)
    Reset-MockStats
    $r = Send-Control $p '130' 'start'
    Assert (-not (Test-IsError $r) -and (Get-MockTotal (Get-MockStats)) -eq 0) 'start on an instance VME is still provisioning is a no-op success (no PUT)'
    [void](Send-Control $p '104' 'delete')
    $r = Send-Control $p '104' 'stop'
    Assert (Test-IsError $r) 'a power action on a guest being deleted is refused'

    Invoke-Mock '/_ctl/fail?key=PUT%20/api/instances/{id}/stop&count=1&code=0' | Out-Null
    Reset-MockStats
    $r = Send-Control $p '131' 'stop'
    Assert (-not (Test-IsError $r) -and (Get-MockCount (Get-MockStats) 'PUT /api/instances/{id}/stop') -eq 2) 'a power PUT that got no response is retried once on a fresh connection'
    $lp = (Invoke-Mock '/_ctl/state' 'GET').last_power
    Assert ($lp.iid -eq 131 -and $lp.action -eq 'stop' -and $lp.query -eq 'force=true') "stop is VME's hard power-off: PUT .../stop?force=true (got [$($lp.query)])"
    Reset-MockStats
    $g = Get-Result (Get-Guest $p '131') 'get 131'
    Assert ($g.state -eq 'powered_off' -and (Get-MockCount (Get-MockStats) 'GET /api/instances/{id}') -eq 1) 'after a provider-issued action the next guests/get re-reads that one guest (dirty record), nothing else'
    Reset-MockStats
    [void](Get-Guest $p '131')
    Assert ((Get-MockTotal (Get-MockStats)) -eq 0) '...and the one after that is free again'
    $r = Send-Control $p '131' 'teleport'
    Assert ((Test-IsError $r) -and $r.error.code -eq -32602) 'an unknown control action is InvalidParams'

    Write-Host "`n=== maintenance: join and exit (template 21) ===" -ForegroundColor Cyan
    Send-Convert $p '21' $true | Out-Null
    Assert ((Get-Result (Get-Guest $p '21') 'get').is_template -eq $true) 'Ready template: is_template=true'
    $tk = Send-Convert $p '21' $false
    Assert ((Get-Task $p $tk.task_id).state -eq 'completed') 'enter maintenance: convert(false) completes at once'
    Assert ((Get-Result (Get-Guest $p '21') 'get').is_template -eq $false) 'in maintenance: is_template=false (or RAS rolls it back to Ready)'
    $maint = Get-Content (Join-Path $dir 'VME-RAS-Maintenance.json') -Raw | ConvertFrom-Json
    $uuid21 = (Invoke-Mock '/_ctl/instance?id=21').instance.uuid
    Assert ($maint.entries[0].id -eq '21' -and $maint.entries[0].uuid -eq $uuid21) 'the maintenance file stores {id, uuid}'
    $r = Send-Control $p '21' 'start'
    Assert (-not (Test-IsError $r)) 'RAS may start the template while in maintenance'

    Write-Host "    restarting the provider mid-maintenance" -ForegroundColor DarkGray
    Stop-Provider $p
    $p = Start-Provider -Dir $dir
    Connect-Provider $p | Out-Null
    [void](Get-GuestList $p)
    Assert ((Get-Result (Get-Guest $p '21') 'get').is_template -eq $false) 'maintenance survives a provider restart'
    [void](Send-Control $p '21' 'stop')
    Send-Convert $p '21' $true | Out-Null
    Assert ((Get-Result (Get-Guest $p '21') 'get').is_template -eq $true) 'exit maintenance: convert(true) -> is_template=true'
    $maint = Get-Content (Join-Path $dir 'VME-RAS-Maintenance.json') -Raw | ConvertFrom-Json
    Assert (@($maint.entries).Count -eq 0) 'the maintenance file is empty again'
    $tmpl = Get-Content (Join-Path $dir 'VME-RAS-BasicTemplates.json') -Raw | ConvertFrom-Json
    Assert (@($tmpl.entries | Where-Object { $_.id -eq '21' -and $_.uuid -eq $uuid21 }).Count -eq 1) 'the template file stores {id, uuid}'

    Write-Host "`n=== an id reused by a different instance ===" -ForegroundColor Cyan
    Stop-Provider $p
    Invoke-Mock '/_ctl/instance?id=21&delete=1' | Out-Null
    Invoke-Mock '/_ctl/instance?id=21&name=someone-else&status=running&ip=10.2.1.1' | Out-Null   # same id, new uuid
    $p = Start-Provider -Dir $dir
    Connect-Provider $p | Out-Null
    [void](Get-GuestList $p)
    Assert ((Get-Result (Get-Guest $p '21') 'get').is_template -eq $false) 'a reused id (different uuid) does not inherit is_template'
    Assert ((Get-LogText $p) -match 'Instance id \[21\] now has uuid') 'the uuid mismatch is logged'

    Write-Host "`n=== template flapping workaround ===" -ForegroundColor Cyan
    Stop-Provider $p
    $dir2 = New-ProviderDir -Name 'u4b' -Settings @{ inventory = @{ list_min_refresh_seconds = 0 }; cloning = @{ template_flapping_workaround_enabled = $true } }
    $p = Start-Provider -Dir $dir2
    Connect-Provider $p | Out-Null
    [void](Get-GuestList $p)
    Send-Convert $p '110' $true | Out-Null
    $g = Get-Result (Get-Guest $p '110') 'get 110'
    Assert ($g.state -eq 'powered_off' -and @($g.ip_addresses).Count -eq 0) 'a running Ready template is masked to powered_off'
    [void](Send-Control $p '110' 'start')
    Start-Sleep -Seconds 2
    $g = Get-Result (Get-Guest $p '110') 'get 110'
    Assert ($g.state -eq 'powered_on') 'an explicit RAS start lifts the mask'
    $lp = (Invoke-Mock '/_ctl/state' 'GET').last_power
    Assert ($lp.action -eq 'start' -and [string]::IsNullOrEmpty($lp.query)) 'start carries no force parameter'

    Write-Host "`n=== power.force_stop off: VME's graceful shutdown ===" -ForegroundColor Cyan
    Stop-Provider $p
    $dir3 = New-ProviderDir -Name 'u4c' -Settings @{ inventory = @{ list_min_refresh_seconds = 0 }; power = @{ force_stop = $false } }
    $p = Start-Provider -Dir $dir3
    Connect-Provider $p | Out-Null
    [void](Get-GuestList $p)
    $r = Send-Control $p '132' 'stop'
    $lp = (Invoke-Mock '/_ctl/state' 'GET').last_power
    Assert (-not (Test-IsError $r) -and $lp.iid -eq 132 -and [string]::IsNullOrEmpty($lp.query)) "with power.force_stop off, stop is the plain PUT .../stop (got [$($lp.query)])"
}
catch { Assert $false "suite aborted: $($_.Exception.Message) at $($_.InvocationInfo.PositionMessage)" }
finally { Complete-Suite 'vme-u4' }
