[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$provider = Join-Path $root 'Parallels-RAS-CPF-XenOrchestra.ps1'
$errors = $null
[Management.Automation.Language.Parser]::ParseFile($provider,[ref]$null,[ref]$errors) | Out-Null
if ($errors.Count) { throw "Provider has $($errors.Count) parse error(s): $($errors -join '; ')" }

$requests = @(
    '{"method":"provider/initialize"}'
    '{"method":"provider/connect","params":{}}'
    '{"method":"not/a-method"}'
    '{broken'
)
$responses = @($requests | & pwsh -NoProfile -NonInteractive -File $provider)
if ($responses.Count -ne 4) { throw "Expected 4 responses, got $($responses.Count)" }
$init = $responses[0] | ConvertFrom-Json
$bad = $responses[1] | ConvertFrom-Json
$unknown = $responses[2] | ConvertFrom-Json
$malformed = $responses[3] | ConvertFrom-Json
if ($init.result.version -ne '1.0.0') { throw 'Unexpected CPF version' }
if ($init.result.capabilities.template_method -ne 'versioning') { throw 'Active config did not advertise versioning' }
if (-not $init.result.capabilities.can_link_clones) { throw 'Active config did not advertise linked clones' }
if ($bad.error.code -ne -32602) { throw 'Invalid connection request did not return InvalidParams' }
if ($unknown.error.code -ne -32601) { throw 'Unknown method code is wrong' }
if ($malformed.error.code -ne -32700) { throw 'Malformed JSON did not return ParseError' }

$config = Get-Content -Raw (Join-Path $root 'XO-RAS-Config.example.json') | ConvertFrom-Json
if ($config.schema_version -ne 2) { throw 'Unexpected settings schema' }
if ($config.connection.skip_tls_default.value) { throw 'TLS verification must default on' }
foreach ($section in @('connection','capabilities','inventory','logging')) {
    foreach ($leaf in $config.$section.PSObject.Properties | Where-Object Name -ne '_comment') {
        foreach ($field in @('value','default','description')) {
            if ($leaf.Value.PSObject.Properties.Name -notcontains $field) { throw "$section.$($leaf.Name) lacks [$field]" }
        }
    }
}
$source = Get-Content -Raw $provider
foreach ($marker in @('/rest/v0','/api/','authenticationToken=','session.signIn','vm.clone','/actions/clone','/actions/revert_snapshot','action = ''resume''','Snapshot clones require template_method=versioning','name_label = [string]$p.name','fast = $fast','full_copy=(-not $fast)','guests/snapshots/revert','ConvertTo-Json -Depth 20 -Compress','Protect-LogText','Rotate-Log','log_rotate_max_generations','request=$($script:RequestId)','Level DEBUG','TemplateStatePath','Save-TemplateState','completed:convert:')) {
    if (-not $source.Contains($marker)) { throw "Missing implementation marker [$marker]" }
}
if (-not $source.Contains('@($detail.errors) -join ''; ''')) { throw 'Structured XOA task errors are not preserved' }
if ($source.Contains('$r.Headers.Location')) { throw 'Empty 204 responses use unsafe StrictMode header access' }
if ($source -match '(?i)token\s*=\s*["''][^"'']+["'']') { throw 'Possible hard-coded token found' }
if ($source -match 'function\s+BaseUrl\s*\(\[string\]\$Host\)') { throw 'BaseUrl shadows the read-only PowerShell Host variable' }
if ($source -match 'ReadToEnd') { throw 'Provider waits for stdin EOF instead of serving requests line-by-line' }
if (-not $source.Contains("'interrupted' {'failed'}")) { throw 'XOA interrupted tasks are not mapped to failed' }

$tempConfig = Join-Path ([IO.Path]::GetTempPath()) "xo-cpf-settings-$([guid]::NewGuid()).json"
$tempLog = Join-Path ([IO.Path]::GetTempPath()) "xo-cpf-log-$([guid]::NewGuid()).log"
try {
    $config.capabilities.guests_polling_rate.value = 17
    [IO.File]::WriteAllText($tempConfig,($config | ConvertTo-Json -Depth 10),[Text.UTF8Encoding]::new($false))
    $configured = @('{"method":"provider/initialize"}' | & pwsh -NoProfile -NonInteractive -File $provider -SettingsPath $tempConfig)[0] | ConvertFrom-Json
    if ($configured.result.capabilities.guests_polling_rate -ne 17) { throw 'Schema 2 value was not loaded' }

    $config.schema_version = 1
    $config.capabilities.guests_polling_rate.value = 19
    [IO.File]::WriteAllText($tempConfig,($config | ConvertTo-Json -Depth 10),[Text.UTF8Encoding]::new($false))
    $legacy = @('{"method":"provider/initialize"}' | & pwsh -NoProfile -NonInteractive -File $provider -SettingsPath $tempConfig)[0] | ConvertFrom-Json
    if ($legacy.result.capabilities.guests_polling_rate -ne 19) { throw 'Schema 1 migration input was not loaded' }

    $config.schema_version = 2
    $config.capabilities.tasks_polling_retries.value = 181
    [IO.File]::WriteAllText($tempConfig,($config | ConvertTo-Json -Depth 10),[Text.UTF8Encoding]::new($false))
    $rejected = @('{"method":"provider/initialize"}' | & pwsh -NoProfile -NonInteractive -File $provider -SettingsPath $tempConfig)[0] | ConvertFrom-Json
    if ($rejected.result.capabilities.tasks_polling_retries -ne 180) { throw 'Out-of-range settings were not rejected' }

    $config.capabilities.tasks_polling_retries.value = 180
    $config.capabilities.template_method.value = 'versioning'
    [IO.File]::WriteAllText($tempConfig,($config | ConvertTo-Json -Depth 10),[Text.UTF8Encoding]::new($false))
    $versioning = @('{"method":"provider/initialize"}' | & pwsh -NoProfile -NonInteractive -File $provider -SettingsPath $tempConfig)[0] | ConvertFrom-Json
    if ($versioning.result.capabilities.template_method -ne 'versioning' -or -not $versioning.result.capabilities.can_link_clones) { throw 'Versioning mode was not advertised' }

    $config.capabilities.template_method.value = 'unsupported'
    [IO.File]::WriteAllText($tempConfig,($config | ConvertTo-Json -Depth 10),[Text.UTF8Encoding]::new($false))
    $rejected = @('{"method":"provider/initialize"}' | & pwsh -NoProfile -NonInteractive -File $provider -SettingsPath $tempConfig)[0] | ConvertFrom-Json
    if ($rejected.result.capabilities.template_method -ne 'basic' -or $rejected.result.capabilities.can_link_clones) { throw 'Unsupported template method was not rejected' }

    $config.capabilities.template_method.value = 'basic'
    $config.logging.log_file.value = $tempLog
    $config.logging.log_level.value = 'debug'
    $config.logging.log_rotate_max_mb.value = 1
    $config.logging.log_rotate_max_generations.value = 2
    [IO.File]::WriteAllText($tempConfig,($config | ConvertTo-Json -Depth 10),[Text.UTF8Encoding]::new($false))
    [IO.File]::WriteAllText($tempLog,('x' * (1MB + 1)),[Text.UTF8Encoding]::new($false))
    $secretToken = 'DEBUG_SECRET_TOKEN'
    $secretPassword = 'DEBUG_SECRET_PASSWORD'
    $debugRequest = @{ method='not/a-method'; params=@{ token=$secretToken; password=$secretPassword; normal='visible' } } | ConvertTo-Json -Compress
    $debugOutput = @($debugRequest | & pwsh -NoProfile -NonInteractive -File $provider -SettingsPath $tempConfig)
    if ($debugOutput.Count -ne 1 -or ($debugOutput[0] | ConvertFrom-Json).error.code -ne -32601) { throw 'Debug logging polluted CPF stdout' }
    if (-not (Test-Path -LiteralPath "$tempLog.1")) { throw 'Log rotation did not retain generation .1' }
    $logText = Get-Content -Raw -LiteralPath $tempLog
    if ($logText -notmatch 'level=DEBUG' -or $logText -notmatch 'request=\d+-\d+' -or $logText -notmatch 'visible') { throw 'Debug request logging or correlation is missing' }
    if ($logText -match [regex]::Escape($secretToken) -or $logText -match [regex]::Escape($secretPassword)) { throw 'Debug log leaked a credential' }
    if ($logText -notmatch '\[redacted\]') { throw 'Debug log did not show credential redaction' }
}
finally {
    foreach ($path in @($tempConfig,$tempLog,"$tempLog.1","$tempLog.2")) { if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force } }
}
& (Join-Path $PSScriptRoot 'Test-XenOrchestra.TemplateState.ps1')
Write-Host 'All offline Xen Orchestra CPF checks passed.' -ForegroundColor Green
