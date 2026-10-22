# vme-u5: transport, protocol, secrets, settings migration, 2.x state files under de-DE
# (review 8.2 vme-u7..u10, vme-s1).
. (Join-Path $PSScriptRoot 'lib.ps1')

function Get-Leaf { param($Doc, [string]$Path) $n = $Doc; foreach ($k in $Path.Split('.')) { $n = $n.$k }; return $n }

Start-VmeMock -Count 20
try {
    Write-Host "`n=== transport ===" -ForegroundColor Cyan
    $dir = New-ProviderDir -Name 'u5' -Settings @{ connection = @{ retry_delay_milliseconds = 10 }; logging = @{ log_level = 5 } }
    $p = Start-Provider -Dir $dir
    Connect-Provider $p | Out-Null
    [void](Get-GuestList $p)
    foreach ($i in 0..19) { Invoke-Mock "/_ctl/instance?id=$(9000 + $i)&name=t-$i&status=running&ip=10.3.0.$($i + 1)" | Out-Null }

    Reset-MockStats
    foreach ($i in 0..9) { [void](Get-Result (Get-Guest $p "$(9000 + $i)") "get $i") }
    $st = Get-MockStats
    Assert ($st.connections -le 1) "20 HTTPS calls over $($st.connections) connection(s): the session is reused"

    Invoke-Mock '/_ctl/force401?count=1' | Out-Null
    Reset-MockStats
    $g = Get-Guest $p '9010'
    Assert (-not (Test-IsError $g) -and (Get-MockCount (Get-MockStats) 'POST /oauth/token') -eq 1) 'a 401 triggers one re-authentication and the call is retried'

    Invoke-Mock '/_ctl/force401?count=5' | Out-Null
    $g = Get-Guest $p '9011'
    Assert ((Test-IsError $g) -and $g.error.message -match 'rejected the credentials') 'a second 401 right after re-authenticating is reported as rejected credentials'
    Invoke-Mock '/_ctl/force401?count=0' | Out-Null

    Invoke-Mock '/_ctl/fail?key=GET%20/api/instances/{id}&count=1&code=429&extra=1' | Out-Null
    $t0 = Get-Date
    $g = Get-Guest $p '9012'
    $secs = ((Get-Date) - $t0).TotalSeconds
    Assert (-not (Test-IsError $g) -and $secs -ge 0.9) "HTTP 429 honours Retry-After (1s) and then succeeds ($([Math]::Round($secs,2))s)"

    Invoke-Mock '/_ctl/fail?key=GET%20/api/instances/{id}&count=1&code=0' | Out-Null
    Reset-MockStats
    $g = Get-Guest $p '9013'
    $st = Get-MockStats
    # A connection dropped before any response byte on a reused connection is resent
    # by .NET's own handler, below the provider; either way the call must succeed.
    Assert (-not (Test-IsError $g) -and (Get-MockCount $st 'GET /api/instances/{id}') -ge 2) 'a dropped connection is retried and the call succeeds'
    $log = Get-LogText $p
    Assert ($log -match '429: Too many requests \(attempt 1/3\) -- retrying in 1000ms') 'a retried attempt is logged at W with the reason and the wait'

    Write-Host "`n=== protocol ===" -ForegroundColor Cyan
    foreach ($m in @('hosts/list', 'hosts/get', 'guests/snapshots/create', 'guests/snapshots/revert')) {
        $r = Send-Rpc $p "{`"method`":`"$m`",`"params`":{`"id`":`"19`",`"name`":`"x`"}}"
        Assert ($r.error.code -eq -32601) "$m answers -32601 (method not found) under template_method basic"
    }
    Assert ((Send-Rpc $p 'not json').error.code -eq -32700) 'malformed JSON -> -32700'
    Assert ((Send-Rpc $p '{"method":"guests/get","params":{}}').error.code -eq -32602) 'empty params -> -32602'
    Assert ((Send-Rpc $p '{"method":"guests/get","params":{"id":["20","21"]}}').error.code -eq -32602) 'batched guests/get -> -32602 (RAS sends single ids)'
    Assert (-not (Test-IsError (Send-Rpc $p '{"method":"guests/get","params":{"id":["20"]}}'))) 'a one-element id array is accepted'
    Assert ((Get-Task $p 'clonebyname:19:nope:1').state -eq 'failed') 'an unknown clone task reports failed'
    Assert ((Get-Task $p '12345').state -eq 'failed') 'a numeric (2.x process) task id reports failed -- nothing polls processes any more'

    Write-Host "`n=== secrets ===" -ForegroundColor Cyan
    [void](Send-Rpc $p '{"method":"provider/disconnect"}')
    Assert ((Send-Rpc $p '{"method":"guests/list"}').PSObject.Properties.Name -contains 'error') 'after disconnect, guests/list reports not connected'
    $c = Get-Result (Send-Rpc $p '{"method":"provider/connect","params":{"settings":{"host":"127.0.0.1:8445","access_token":"mock-token-abc","refresh_token":"mock-refresh-xyz","ssl_check":false}}}') 'connect with token'
    Assert ($c.connected) 'connect with a pre-issued access_token + refresh_token'
    $log = Get-LogText $p
    Assert ($log -notmatch 'mock-token-abc' -and $log -notmatch 'mock-refresh-xyz' -and $log -notmatch '"password":"password"') 'no token, refresh token or password anywhere in the verbose log'
    Assert ($log -match '"access_token":"\*\*\*"' -and $log -match '"refresh_token":"\*\*\*"') 'the connect requests are logged with the secrets masked'
    Assert ($log -notmatch 'ssh-secret-value|server-api-key-value') 'server credential fields never reach the log'
    Stop-Provider $p

    Write-Host "`n=== settings: schema-1 (2.x) file migrated ===" -ForegroundColor Cyan
    $v2 = @'
{
  "schema_version": 1,
  "locations": { "settings_reload_interval_seconds": 5, "clone_ledger_file": "VME-RAS-CloneLedger.json" },
  "connection": { "http_timeout_seconds": 20, "retry_max_attempts": 2, "proactive_token_renewal_enabled": false },
  "capabilities": { "guests_polling_rate": 5, "tasks_polling_rate": 20, "tasks_polling_retries": 180, "template_method": "versioning", "can_link_clones": false },
  "cloning": { "instances_page_size": 100, "instances_cache_ttl_seconds": 5, "clone_resolve_max_attempts": 40, "clone_provision_wait_max_attempts": 160,
               "template_snapshot_prefix": "RAS_TEMPLATE_VERSION_", "mac_preservation_enabled": true, "sweep_min_interval_seconds": 7 },
  "network": { "ipv4_paths": ["connectionInfo[].ip"], "max_addresses": 2 },
  "filtering": { "exclude_instance_name_patterns": ["^filler-10[0-4]$"] },
  "orphan_detection": { "ledger_retention_days": 30 },
  "logging": { "log_level": 5, "log_rotate_max_bytes": 20971520 }
}
'@
    $dir = New-ProviderDir -Name 'u5-settings' -RawSettingsJson $v2
    $p = Start-Provider -Dir $dir
    $init = Get-Result (Send-Rpc $p '{"method":"provider/initialize"}') 'initialize'
    Assert ($init.capabilities.guests_polling_rate -eq 60) 'guests_polling_rate at its old 2.x default (5) takes the new default (60)'
    Assert ($init.capabilities.tasks_polling_rate -eq 20) 'a customised tasks_polling_rate (20) is kept'
    Assert ($init.capabilities.template_method -eq 'basic' -and $init.capabilities.can_link_clones -eq $false) "template_method 'versioning' becomes 'basic'"
    $doc = Get-Content (Join-Path $dir 'VME-RAS-Config.json') -Raw | ConvertFrom-Json
    Assert ($doc.schema_version -eq 2) 'the file is rewritten at schema_version 2'
    Assert ((Get-Leaf $doc 'inventory.page_size').value -eq 100) 'cloning.instances_page_size moved to inventory.page_size with its value'
    Assert ((Get-Leaf $doc 'cloning.resolve_timeout_seconds').value -eq 800) 'clone_resolve_max_attempts 40 at 20 s became resolve_timeout_seconds 800'
    Assert ((Get-Leaf $doc 'cloning.provision_timeout_seconds').value -eq 1740) 'a 2.x attempt budget at its default takes the new time default'
    Assert ((Get-Leaf $doc 'cloning.mac_preservation.enabled').value -eq $true -and (Get-Leaf $doc 'cloning.tracked_refresh_interval_seconds').value -eq 7) 'renamed keys keep their values'
    Assert ((Get-Leaf $doc 'inventory.exclude_instance_name_patterns').value -contains '^filler-10[0-4]$' -and (Get-Leaf $doc 'inventory.max_addresses').value -eq 2) 'filtering/network values moved to inventory'
    Assert ((Get-Leaf $doc 'logging.log_level').value -eq 4 -and (Get-Leaf $doc 'logging.log_rotate_max_mb').value -eq 20) 'log_level at its 2.x default takes the new default; bytes became MB'
    Assert ((Get-Leaf $doc 'connection.request_timeout_seconds').value -eq 30 -and (Get-Leaf $doc 'connection.retry_max_attempts').value -eq 2 -and (Get-Leaf $doc 'connection.token_renewal.proactive_enabled').value -eq $false) 'connection values carried over'
    $raw = Get-Content (Join-Path $dir 'VME-RAS-Config.json') -Raw
    Assert ($raw -notmatch 'template_snapshot_prefix|ipv4_paths|clone_ledger_file|instances_cache_ttl|ledger_retention_days|"filtering"|"network"') 'removed keys are dropped from the file'
    Assert ((Get-Leaf $doc 'capabilities.tasks_polling_rate').description.Length -gt 20 -and (Get-Leaf $doc 'capabilities.tasks_polling_rate').default -eq 10) 'every leaf carries value, default and description'
    $log = Get-LogText $p
    Assert ($log -match 'rewrote .* from schema_version 1 to 2' -and $log -match '- cloning.template_snapshot_prefix \(removed\)') 'the migration is logged with what changed'

    Write-Host "`n=== settings: hot reload ===" -ForegroundColor Cyan
    $doc.capabilities.guests_polling_rate.value = 120
    $doc | Add-Member -NotePropertyName 'bogus' -NotePropertyValue @{ key = 1 } -Force
    Start-Sleep -Milliseconds 1200
    Set-Content -LiteralPath (Join-Path $dir 'VME-RAS-Config.json') -Value ($doc | ConvertTo-Json -Depth 10) -Encoding utf8NoBOM
    Start-Sleep -Seconds 6
    $init2 = Get-Result (Send-Rpc $p '{"method":"provider/initialize"}') 'initialize 2'
    $log = Get-LogText $p
    Assert ($init2.capabilities.guests_polling_rate -eq 120 -and $log -match 'capabilities.guests_polling_rate changed \[60\] -> \[120\]') 'an edit is hot-reloaded and the change is logged'
    Assert ($log -match 'Settings reload: unknown key\(s\) ignored: bogus.key') 'an unknown key is reported, not silently ignored'
    Stop-Provider $p

    Write-Host "`n=== settings: invalid values and a corrupt file ===" -ForegroundColor Cyan
    $dir = New-ProviderDir -Name 'u5-bad' -Settings @{ capabilities = @{ tasks_polling_rate = 0; template_method = 'weird' }; inventory = @{ max_drop_ratio = 'abc' } }
    $p = Start-Provider -Dir $dir
    $init = Get-Result (Send-Rpc $p '{"method":"provider/initialize"}') 'initialize bad'
    $log = Get-LogText $p
    Assert ($init.capabilities.tasks_polling_rate -eq 10 -and $init.capabilities.template_method -eq 'basic') 'out-of-range and disallowed values fall back to their defaults'
    Assert ($log -match 'tasks_polling_rate = \[0\] rejected \(below the minimum 5\)' -and $log -match 'max_drop_ratio = \[abc\] rejected') 'each rejected value is logged with the reason'
    Stop-Provider $p
    $dir = New-ProviderDir -Name 'u5-corrupt' -RawSettingsJson '{ "capabilities": { "tasks_polling_rate": 15, '
    $p = Start-Provider -Dir $dir
    $init = Get-Result (Send-Rpc $p '{"method":"provider/initialize"}') 'initialize corrupt'
    Assert ($init.capabilities.tasks_polling_rate -eq 10) 'a corrupt settings file: the provider starts on defaults'
    Assert ((Get-Content (Join-Path $dir 'VME-RAS-Config.json') -Raw).Trim() -eq '{ "capabilities": { "tasks_polling_rate": 15,') 'and leaves the file untouched (an edit in progress is never overwritten)'
    Stop-Provider $p

    Write-Host "`n=== 2.x state files, provider running under de-DE ===" -ForegroundColor Cyan
    $dir = New-ProviderDir -Name 'u5-state' -Settings @{ inventory = @{ list_min_refresh_seconds = 0 } }
    $started = [DateTime]::UtcNow.AddSeconds(-25).ToString('o')
    @"
{
  "pending:RDSH-X": { "type": "clone", "task_id": "clonebyname:19:RDSH-X", "source_id": "19", "clone_id": null, "name": "RDSH-X",
                      "resolve_attempts": 3, "provision_wait_attempts": 0, "creation_completed": false, "provision_confirmed": false,
                      "reported_completed_at": null, "clone_started_at": "$started", "preserved_mac": null },
  "555": { "type": "clone", "task_id": "4242", "source_id": "19", "clone_id": "555", "name": "OLD", "clone_started_at": "$started" }
}
"@ | Set-Content -LiteralPath (Join-Path $dir 'VME-RAS-CloneState.json') -Encoding utf8NoBOM
    '{ "ids": [ "20" ] }' | Set-Content -LiteralPath (Join-Path $dir 'VME-RAS-Maintenance.json') -Encoding utf8NoBOM
    $p = Start-Provider -Dir $dir -Culture 'de-DE'
    Connect-Provider $p | Out-Null
    [void](Get-GuestList $p)
    $t = Get-Task $p 'clonebyname:19:RDSH-X'
    Assert ($t.state -eq 'running') "a 2.x in-flight clone is still tracked after a restart under de-DE (state=$($t.state))"
    $state = Get-Content (Join-Path $dir 'VME-RAS-CloneState.json') -Raw
    Assert ($state -match '"started_epoch": \d{10}' -and $state -notmatch '"4242"') 'the 2.x entry was converted to epoch time; the process-id entry was dropped'
    Invoke-Mock '/_ctl/instance?id=9500&name=RDSH-X&status=running&ip=10.4.0.1&cloned_from=19' | Out-Null
    $t = Get-Task $p 'clonebyname:19:RDSH-X'
    Assert ($t.state -eq 'completed' -and $t.output.clone_id -eq '9500') 'it resolves and completes'
    $log = Get-LogText $p
    $m = [regex]::Match($log, 'Clone \[RDSH-X\] resolved to instance \[9500\] (\d+)s after the submit')
    Assert ($m.Success -and [int]$m.Groups[1].Value -ge 20 -and [int]$m.Groups[1].Value -le 120) "elapsed time is computed correctly across the restart and culture ($($m.Groups[1].Value)s)"
    $maint = Get-Content (Join-Path $dir 'VME-RAS-Maintenance.json') -Raw | ConvertFrom-Json
    Assert ($maint.entries[0].id -eq '20' -and $maint.entries[0].uuid.Length -eq 36) 'a 2.x maintenance entry adopts the instance uuid at the first listing'
    $g20 = Get-Result (Get-Guest $p '20') 'get 20'
    Assert ($g20.is_template -eq $false) 'the maintenance override still applies'

    Write-Host "`n=== log file recreated after an external delete ===" -ForegroundColor Cyan
    Stop-Provider $p
    $dir2 = New-ProviderDir -Name 'u5-logdel' -Settings @{ connection = @{ retry_delay_milliseconds = 10 }; logging = @{ log_level = 5; file_exists_check_seconds = 5 } }
    $p2 = Start-Provider -Dir $dir2
    Connect-Provider $p2 | Out-Null
    [void](Get-GuestList $p2)
    $logPath = Join-Path $dir2 'VME-RAS-Provider.log'
    Assert (Test-Path $logPath) 'the log file exists before the deletion test'
    Remove-Item -LiteralPath $logPath -Force
    Assert (-not (Test-Path $logPath)) 'the log file is gone right after an external delete, even with the provider still holding it open'
    Start-Sleep -Seconds 6   # past file_exists_check_seconds=5 (the declared minimum)
    [void](Get-Result (Get-Guest $p2 '19') 'get after delete')   # any log write is a chance for the periodic check to fire
    Assert (Test-Path $logPath) 'the log file reappears on its own once the periodic check notices the deletion'
    $log2 = Get-LogText $p2
    Assert ($log2 -match 'Log file was missing at the expected path \(deleted externally\?\) -- recreated') 'the recreation is itself logged, so the gap is visible rather than silent'
    Stop-Provider $p2
}
catch { Assert $false "suite aborted: $($_.Exception.Message) at $($_.InvocationInfo.PositionMessage)" }
finally { Complete-Suite 'vme-u5' }
