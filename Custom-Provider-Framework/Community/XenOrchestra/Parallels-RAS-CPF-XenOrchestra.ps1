#requires -Version 7.0
[CmdletBinding()]
param(
    [string]$SettingsPath = (Join-Path $PSScriptRoot 'XO-RAS-Config.json'),
    [string]$TemplateStatePath = (Join-Path $PSScriptRoot 'XenOrchestra-RAS-TemplateState.json'),
    [switch]$NoRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:Session = $null
$script:LastBaseUrl = $null
$script:Inventory = @{}
$script:Settings = $null
$script:TemplateIds = @{}
$script:MaintenanceIds = @{}
$script:TemplateStateLoadError = $null
$script:TemplateStatePath = if ([IO.Path]::IsPathRooted($TemplateStatePath)) { $TemplateStatePath } else { Join-Path $PSScriptRoot $TemplateStatePath }
$script:RequestId = '-'
$script:RequestCounter = 0
$script:Errors = @{ Parse = -32700; InvalidRequest = -32600; NotFound = -32601; InvalidParams = -32602; Internal = -32603 }

function Leaf($Value) { if ($null -ne $Value -and $null -ne $Value.PSObject.Properties['value']) { return $Value.value }; return $Value }
function Setting-Int($Value,[string]$Name,[int]$Min,[int]$Max) {
    $parsed = 0
    if (-not [int]::TryParse([string]$Value,[ref]$parsed) -or $parsed -lt $Min -or $parsed -gt $Max) { throw "$Name must be an integer from $Min to $Max" }
    return $parsed
}
function Setting-Bool($Value,[string]$Name) {
    if ($Value -isnot [bool]) { throw "$Name must be true or false" }
    return [bool]$Value
}
function Setting-LogLevel($Value) {
    $level = ([string]$Value).Trim().ToLowerInvariant()
    if ($level -notin @('standard','extended','debug')) { throw 'logging.log_level must be standard, extended, or debug' }
    $level
}
function Get-Defaults {
    [pscustomobject]@{
        connection = [pscustomobject]@{ request_timeout_seconds = 60; skip_tls_default = $false }
        capabilities = [pscustomobject]@{ can_suspend_guests = $true; guests_polling_rate = 30; tasks_polling_rate = 5; tasks_polling_retries = 180; template_method = 'basic' }
        inventory = [pscustomobject]@{ include_templates = $true; name_filter = '' }
        logging = [pscustomobject]@{ log_file = 'XenOrchestra-RAS-Provider.log'; log_level = 'extended'; log_rotate_max_mb = 20; log_rotate_max_generations = 3 }
    }
}
function Import-Settings {
    $s = Get-Defaults
    if (Test-Path -LiteralPath $SettingsPath) {
        $f = Get-Content -Raw -LiteralPath $SettingsPath | ConvertFrom-Json
        if ([int]$f.schema_version -notin @(1,2)) { throw "Unsupported settings schema [$($f.schema_version)]" }
        foreach ($section in @('connection','capabilities','inventory','logging')) {
            if ($f.PSObject.Properties.Name -contains $section) {
                foreach ($p in $f.$section.PSObject.Properties) {
                    if ($s.$section.PSObject.Properties.Name -contains $p.Name) { $s.$section.($p.Name) = Leaf $p.Value }
                }
            }
        }
    }
    $s.connection.request_timeout_seconds = Setting-Int $s.connection.request_timeout_seconds 'connection.request_timeout_seconds' 5 300
    $s.connection.skip_tls_default = Setting-Bool $s.connection.skip_tls_default 'connection.skip_tls_default'
    $s.capabilities.can_suspend_guests = Setting-Bool $s.capabilities.can_suspend_guests 'capabilities.can_suspend_guests'
    $s.capabilities.guests_polling_rate = Setting-Int $s.capabilities.guests_polling_rate 'capabilities.guests_polling_rate' 5 600
    $s.capabilities.tasks_polling_rate = Setting-Int $s.capabilities.tasks_polling_rate 'capabilities.tasks_polling_rate' 5 300
    $s.capabilities.tasks_polling_retries = Setting-Int $s.capabilities.tasks_polling_retries 'capabilities.tasks_polling_retries' 1 180
    $s.inventory.include_templates = Setting-Bool $s.inventory.include_templates 'inventory.include_templates'
    if ([string]$s.capabilities.template_method -notin @('basic','versioning')) { throw 'capabilities.template_method must be basic or versioning' }
    if ([string]::IsNullOrWhiteSpace([string]$s.logging.log_file)) { throw 'logging.log_file must not be empty' }
    $s.logging.log_level = Setting-LogLevel $s.logging.log_level
    $s.logging.log_rotate_max_mb = Setting-Int $s.logging.log_rotate_max_mb 'logging.log_rotate_max_mb' 1 2048
    $s.logging.log_rotate_max_generations = Setting-Int $s.logging.log_rotate_max_generations 'logging.log_rotate_max_generations' 1 20
    return $s
}
function Protect-LogText([string]$Text) {
    if ([string]::IsNullOrEmpty($Text)) { return $Text }
    $safe = $Text -replace '(?i)("[^"\r\n]*(?:pass|secret|token|authorization|credential|api_?key)[^"\r\n]*"\s*:\s*)(?:"(?:[^"\\]|\\.)*"|[^,}\]\s]+)','$1"[redacted]"'
    $safe = $safe -replace '(?i)((?:authenticationToken|token|password|passwd|access_token|refresh_token|client_secret|authorization|credential|api_?key)=)[^&\s"'',;]+','$1[redacted]'
    $safe -replace '(?i)(bearer\s+)[^\s"'',;]+','$1[redacted]'
}
function Limit-LogText([string]$Text,[int]$Max = 4096) {
    if ($null -eq $Text -or $Text.Length -le $Max) { return $Text }
    $Text.Substring(0,$Max) + "...[truncated $($Text.Length - $Max) chars]"
}
function Format-LogObject($Value) {
    if ($null -eq $Value) { return 'null' }
    Limit-LogText (Protect-LogText ($Value | ConvertTo-Json -Depth 20 -Compress))
}
function Rotate-Log([string]$Path) {
    if (-not [IO.File]::Exists($Path) -or ([IO.FileInfo]$Path).Length -lt ([long]$script:Settings.logging.log_rotate_max_mb * 1MB)) { return }
    $max = [int]$script:Settings.logging.log_rotate_max_generations
    $oldest = "$Path.$max"
    if ([IO.File]::Exists($oldest)) { [IO.File]::Delete($oldest) }
    for ($i=$max-1; $i -ge 1; $i--) {
        $source = "$Path.$i"
        if ([IO.File]::Exists($source)) { [IO.File]::Move($source,"$Path.$($i+1)",$true) }
    }
    [IO.File]::Move($Path,"$Path.1",$true)
}
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('ERROR','INFO','EXTENDED','DEBUG')][string]$Level = 'INFO',
        [string]$Component = 'CORE'
    )
    try {
        $configuredRank = switch ([string]$script:Settings.logging.log_level) { standard {1} extended {2} debug {3} }
        $eventRank = switch ($Level) { ERROR {0} INFO {1} EXTENDED {2} DEBUG {3} }
        if ($eventRank -gt $configuredRank) { return }
        $path = [string]$script:Settings.logging.log_file
        if (-not [IO.Path]::IsPathRooted($path)) { $path = Join-Path $PSScriptRoot $path }
        $directory = Split-Path $path -Parent
        if ($directory -and -not [IO.Directory]::Exists($directory)) { [IO.Directory]::CreateDirectory($directory) | Out-Null }
        try { Rotate-Log $path } catch { }
        $safe = (Limit-LogText (Protect-LogText $Message)) -replace '[\r\n]+',' '
        Add-Content -LiteralPath $path -Value "$(Get-Date -Format o) level=$Level pid=$PID request=$($script:RequestId) component=$Component message=$safe" -Encoding utf8
    } catch { }
}
function Import-TemplateState {
    $script:TemplateIds = @{}
    $script:MaintenanceIds = @{}
    if (-not [IO.File]::Exists($script:TemplateStatePath)) { return }
    $state = [IO.File]::ReadAllText($script:TemplateStatePath) | ConvertFrom-Json -ErrorAction Stop
    if ([int]$state.schema_version -ne 1) { throw "Unsupported template state schema [$($state.schema_version)]" }
    foreach ($item in @($state.templates)) {
        if ($null -eq $item -or [string]::IsNullOrWhiteSpace([string]$item.id) -or [string]::IsNullOrWhiteSpace([string]$item.host)) { throw 'Invalid templates entry in template state' }
        $script:TemplateIds[[string]$item.id] = [string]$item.host
    }
    foreach ($item in @($state.maintenance)) {
        if ($null -eq $item -or [string]::IsNullOrWhiteSpace([string]$item.id) -or [string]::IsNullOrWhiteSpace([string]$item.host)) { throw 'Invalid maintenance entry in template state' }
        $script:MaintenanceIds[[string]$item.id] = [string]$item.host
    }
}
function Save-TemplateState {
    $path = [IO.Path]::GetFullPath($script:TemplateStatePath)
    $directory = [IO.Path]::GetDirectoryName($path)
    if (-not [IO.Directory]::Exists($directory)) { [IO.Directory]::CreateDirectory($directory) | Out-Null }
    $state = [ordered]@{
        schema_version = 1
        templates = @($script:TemplateIds.Keys | Sort-Object | ForEach-Object { [ordered]@{ id = [string]$_; host = [string]$script:TemplateIds[$_] } })
        maintenance = @($script:MaintenanceIds.Keys | Sort-Object | ForEach-Object { [ordered]@{ id = [string]$_; host = [string]$script:MaintenanceIds[$_] } })
    }
    $temp = Join-Path $directory ".$([IO.Path]::GetFileName($path)).$PID.tmp"
    try {
        [IO.File]::WriteAllText($temp,($state | ConvertTo-Json -Depth 5),[Text.UTF8Encoding]::new($false))
        [IO.File]::Move($temp,$path,$true)
    } finally {
        if ([IO.File]::Exists($temp)) { [IO.File]::Delete($temp) }
    }
}
function Test-TemplateState([hashtable]$State,[string]$Id) {
    $null -ne $script:Session -and $State.ContainsKey($Id) -and [string]$State[$Id] -eq [string]$script:Session.base
}
function Set-TemplateState([string]$Id,[bool]$IsTemplate) {
    $templateExisted = $script:TemplateIds.ContainsKey($Id)
    $templateValue = if ($templateExisted) { [string]$script:TemplateIds[$Id] } else { $null }
    $maintenanceExisted = $script:MaintenanceIds.ContainsKey($Id)
    $maintenanceValue = if ($maintenanceExisted) { [string]$script:MaintenanceIds[$Id] } else { $null }
    try {
        if ($IsTemplate) {
            $script:TemplateIds[$Id] = [string]$script:Session.base
            $script:MaintenanceIds.Remove($Id) | Out-Null
        } else {
            $script:MaintenanceIds[$Id] = [string]$script:Session.base
        }
        Save-TemplateState
    } catch {
        if ($templateExisted) { $script:TemplateIds[$Id] = $templateValue } else { $script:TemplateIds.Remove($Id) | Out-Null }
        if ($maintenanceExisted) { $script:MaintenanceIds[$Id] = $maintenanceValue } else { $script:MaintenanceIds.Remove($Id) | Out-Null }
        throw
    }
}
function Ok($Value) { @{ result = $Value } }
function Fail([int]$Code,[string]$Message) { @{ error = @{ code = $Code; message = $Message } } }
function Has($Object,[string]$Name) { $null -ne $Object -and $null -ne $Object.PSObject.Properties[$Name] }
function Need($Object,[string[]]$Names) {
    foreach ($n in $Names) { if (-not (Has $Object $n) -or [string]::IsNullOrWhiteSpace([string]$Object.$n)) { throw [ArgumentException]::new("Missing parameter [$n]") } }
}
function BaseUrl([string]$HostValue) { ($HostValue.TrimEnd('/') -replace '/rest/v0$','') + '/rest/v0' }
function Invoke-Xo([string]$Method,[string]$Path,[object]$Body = $null) {
    if ($null -eq $script:Session) { throw 'Not connected' }
    $timer = [Diagnostics.Stopwatch]::StartNew()
    Write-Log -Level DEBUG -Component REST -Message "out method=$Method path=$Path body=$(Format-LogObject $Body)"
    $args = @{
        Method = $Method; Uri = "$($script:Session.base)$Path"; Headers = @{ Cookie = "authenticationToken=$($script:Session.token)" }
        TimeoutSec = [int]$script:Settings.connection.request_timeout_seconds; SkipHttpErrorCheck = $true
    }
    if ($script:Session.skip_tls) { $args.SkipCertificateCheck = $true }
    if ($null -ne $Body) { $args.ContentType = 'application/json'; $args.Body = $Body | ConvertTo-Json -Depth 20 -Compress }
    try {
        $r = Invoke-WebRequest @args
        if ($r.StatusCode -lt 200 -or $r.StatusCode -ge 300) { throw "XO HTTP $($r.StatusCode): $($r.Content)" }
    } catch {
        Write-Log -Level ERROR -Component REST -Message "failed method=$Method path=$Path elapsed_ms=$($timer.ElapsedMilliseconds) error=$($_.Exception.Message)"
        throw
    }
    $bytes = if ($null -eq $r.Content) { 0 } else { [Text.Encoding]::UTF8.GetByteCount([string]$r.Content) }
    Write-Log -Level EXTENDED -Component REST -Message "in method=$Method path=$Path status=$($r.StatusCode) elapsed_ms=$($timer.ElapsedMilliseconds) bytes=$bytes"
    Write-Log -Level DEBUG -Component REST -Message "body method=$Method path=$Path value=$(Limit-LogText (Protect-LogText ([string]$r.Content)))"
    if ([string]::IsNullOrWhiteSpace($r.Content)) {
        $location = [string]$r.Headers['Location']
        if ($location -match '/tasks/([^/?]+)') { return [pscustomobject]@{ taskId = $Matches[1] } }
        return $null
    }
    $content = $r.Content | ConvertFrom-Json
    if ($content -is [string] -and $content -match '/tasks/([^/?]+)') { return [pscustomobject]@{ taskId = $Matches[1] } }
    return $content
}
function Invoke-XoRpc([string]$Method,[hashtable]$Params) {
    if ($null -eq $script:Session) { throw 'Not connected' }
    $uri = [UriBuilder]$script:Session.base
    $uri.Scheme = if ($uri.Scheme -eq 'https') { 'wss' } else { 'ws' }
    $uri.Path = '/api/'; $uri.Query = ''
    $socket = [Net.WebSockets.ClientWebSocket]::new()
    if ($script:Session.skip_tls) {
        if ($null -eq ('XoUnsafeCertificateValidator' -as [type])) {
            Add-Type -TypeDefinition @'
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
public static class XoUnsafeCertificateValidator {
    public static bool Validate(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors errors) { return true; }
}
'@ | Out-Null
        }
        $socket.Options.RemoteCertificateValidationCallback = [Delegate]::CreateDelegate(
            [Net.Security.RemoteCertificateValidationCallback], [XoUnsafeCertificateValidator].GetMethod('Validate'))
    }
    $cancel = [Threading.CancellationTokenSource]::new([TimeSpan]::FromSeconds([int]$script:Settings.connection.request_timeout_seconds))
    $receive = {
        $buffer = [byte[]]::new(65536); $stream = [IO.MemoryStream]::new()
        try {
            do {
                $part = $socket.ReceiveAsync([ArraySegment[byte]]::new($buffer),$cancel.Token).GetAwaiter().GetResult()
                if ($part.MessageType -eq [Net.WebSockets.WebSocketMessageType]::Close) { throw 'XO closed the JSON-RPC WebSocket' }
                $stream.Write($buffer,0,$part.Count)
            } while (-not $part.EndOfMessage)
            [Text.Encoding]::UTF8.GetString($stream.ToArray()) | ConvertFrom-Json
        } finally { $stream.Dispose() }
    }
    $call = {
        param([int]$Id,[string]$RpcMethod,[hashtable]$RpcParams)
        $rpcTimer = [Diagnostics.Stopwatch]::StartNew()
        $loggedParams = if ($RpcMethod -eq 'session.signIn') { '{"token":"[redacted]"}' } else { Format-LogObject $RpcParams }
        Write-Log -Level DEBUG -Component JSONRPC -Message "out method=$RpcMethod params=$loggedParams"
        $json = @{ jsonrpc='2.0'; id=$Id; method=$RpcMethod; params=$RpcParams } | ConvertTo-Json -Depth 20 -Compress
        $bytes = [Text.Encoding]::UTF8.GetBytes($json)
        $socket.SendAsync([ArraySegment[byte]]::new($bytes),[Net.WebSockets.WebSocketMessageType]::Text,$true,$cancel.Token).GetAwaiter().GetResult() | Out-Null
        do { $response = & $receive } while (-not (Has $response 'id') -or [int]$response.id -ne $Id)
        if (Has $response 'error') {
            $message = if (Has $response.error 'message') { [string]$response.error.message } else { $response.error | ConvertTo-Json -Depth 10 -Compress }
            Write-Log -Level ERROR -Component JSONRPC -Message "failed method=$RpcMethod elapsed_ms=$($rpcTimer.ElapsedMilliseconds) error=$message"
            throw "XO JSON-RPC $RpcMethod failed: $message"
        }
        Write-Log -Level EXTENDED -Component JSONRPC -Message "in method=$RpcMethod elapsed_ms=$($rpcTimer.ElapsedMilliseconds)"
        Write-Log -Level DEBUG -Component JSONRPC -Message "result method=$RpcMethod value=$(Format-LogObject $response.result)"
        $response.result
    }
    try {
        $socket.ConnectAsync($uri.Uri,$cancel.Token).GetAwaiter().GetResult() | Out-Null
        $null = & $call 1 'session.signIn' @{ token=[string]$script:Session.token }
        & $call 2 $Method $Params
    } finally {
        if ($socket.State -eq [Net.WebSockets.WebSocketState]::Open) { $socket.Abort() }
        $socket.Dispose(); $cancel.Dispose()
    }
}
function VmState([string]$State) {
    switch -Regex ($State) { '^Running$' {'powered_on'} '^Halted$' {'powered_off'} '^Suspended$' {'suspended'} default {'unknown'} }
}
function To-Guest($Vm) {
    if ($Vm -is [string]) { throw "XO inventory returned an object URL instead of fields: [$Vm]" }
    $name = if (Has $Vm 'name_label') { $Vm.name_label } elseif (Has $Vm 'nameLabel') { $Vm.nameLabel } elseif (Has $Vm 'name') { $Vm.name } else { "vm-$($Vm.id)" }
    $power = if (Has $Vm 'power_state') { $Vm.power_state } elseif (Has $Vm 'powerState') { $Vm.powerState } else { '' }
    $ips = @()
    if ((Has $Vm 'addresses') -and $null -ne $Vm.addresses) {
        $ips = @(foreach ($property in $Vm.addresses.PSObject.Properties) {
            if ([string]$property.Value -match '^\d{1,3}(\.\d{1,3}){3}$') { [string]$property.Value }
        })
    }
    $id = [string]$Vm.id
    $nativeTemplate = (Has $Vm 'type') -and [string]$Vm.type -eq 'VM-template'
    $isTemplate = ($nativeTemplate -or (Test-TemplateState $script:TemplateIds $id)) -and -not (Test-TemplateState $script:MaintenanceIds $id)
    [ordered]@{
        id = $id; name = [string]$name; provider = 'Xen Orchestra'; state = VmState ([string]$power)
        power_state = if ($power) { [string]$power } else { 'unknown' }
        ip = if ($ips.Count) { [string]$ips[0] } else { $null }; ip_addresses = @($ips); mac_addresses = @()
        host_os = if ((Has $Vm 'os_version') -and $null -ne $Vm.os_version -and (Has $Vm.os_version 'name')) { [string]$Vm.os_version.name } else { 'unknown' }
        is_template = $isTemplate; type = 'Virtual Machine'
    }
}
function Initialize {
    Ok @{ version = '1.0.0'; name = 'Xen Orchestra'; capabilities = @{
        can_suspend_guests = [bool]$script:Settings.capabilities.can_suspend_guests
        guests_polling_rate = [int]$script:Settings.capabilities.guests_polling_rate
        tasks_polling_rate = [int]$script:Settings.capabilities.tasks_polling_rate
        tasks_polling_retries = [int]$script:Settings.capabilities.tasks_polling_retries
        template_method = [string]$script:Settings.capabilities.template_method
        can_link_clones = ([string]$script:Settings.capabilities.template_method -eq 'versioning')
    } }
}
function Connect($p) {
    Need $p @('settings'); $s = $p.settings; Need $s @('host','token')
    if ($script:TemplateStateLoadError) { return Fail $script:Errors.Internal "Template state unavailable: $($script:TemplateStateLoadError)" }
    $skip = [bool]$script:Settings.connection.skip_tls_default
    if (Has $s 'skip_tls') { $skip = [Convert]::ToBoolean($s.skip_tls) }
    $base = BaseUrl ([string]$s.host)
    if ($null -ne $script:LastBaseUrl -and $script:LastBaseUrl -ne $base) { $script:Inventory = @{} }
    $script:Session = [pscustomobject]@{ base = $base; token = [string]$s.token; skip_tls = $skip }
    try {
        $null = Invoke-Xo GET '/vms?limit=1'
        if ([string]$script:Settings.capabilities.template_method -eq 'versioning') { $null = Invoke-XoRpc 'system.getVersion' @{} }
        $script:LastBaseUrl = $base
        Write-Log -Level INFO -Component CONNECT -Message "connected host=$base skip_tls=$skip template_method=$($script:Settings.capabilities.template_method)"
        Ok @{ connected = $true }
    } catch {
        Write-Log -Level ERROR -Component CONNECT -Message "failed host=$base error=$($_.Exception.Message)"
        $script:Session = $null; Fail $script:Errors.Internal "Connection failed: $($_.Exception.Message)"
    }
}
function List-Guests {
    $fields = 'id,uuid,name_label,power_state,addresses,os_version,VIFs,type,$pool,$poolId'
    $items = @((Invoke-Xo GET "/vms?fields=$fields"))
    if ($script:Settings.inventory.include_templates) { $items += @((Invoke-Xo GET "/vm-templates?fields=$fields")) }
    $filter = [string]$script:Settings.inventory.name_filter
    $next = @{}
    $ids = foreach ($vm in $items) {
        $g = To-Guest $vm
        if ((-not $script:Settings.inventory.include_templates -and $g.is_template) -or ($filter -and $g.name -notlike $filter)) { continue }
        $next[[string]$vm.id] = $vm
        [string]$vm.id
    }
    $script:Inventory = $next
    Write-Log -Level EXTENDED -Component INVENTORY -Message "listed guests=$($next.Count) include_templates=$($script:Settings.inventory.include_templates) filter=$filter"
    Ok @{ guests = @($ids) }
}
function Get-Guest($p) {
    Need $p @('id'); $id = [uri]::EscapeDataString([string]$p.id)
    $vm = $script:Inventory[[string]$p.id]
    if ($null -eq $vm) { try { $vm = Invoke-Xo GET "/vms/$id" } catch { $vm = Invoke-Xo GET "/vm-templates/$id" } }
    Ok (To-Guest $vm)
}
function Action($p) {
    Need $p @('id','control'); $id = [uri]::EscapeDataString([string]$p.id)
    $map = @{ start='start'; stop='hard_shutdown'; restart='clean_reboot'; reset='hard_reboot'; suspend='suspend' }
    Write-Log -Level INFO -Component CONTROL -Message "control=$($p.control) vm=$($p.id)"
    if ($p.control -eq 'delete') { $null = Invoke-Xo DELETE "/vms/$id"; return Ok $true }
    if (-not $map.ContainsKey([string]$p.control)) { return Fail $script:Errors.InvalidParams "Unsupported control [$($p.control)]" }
    $action = $map[[string]$p.control]
    if ($p.control -eq 'start') {
        $vm = Invoke-Xo GET "/vms/$id"
        if ([string]$vm.power_state -eq 'Suspended') { $action = 'resume' }
    }
    $task = Invoke-Xo POST "/vms/$id/actions/$action"
    Ok @{ task_id = [string]$task.taskId }
}
function Convert-Guest($p) {
    Need $p @('id'); if (-not (Has $p 'is_template')) { return Fail $script:Errors.InvalidParams 'Missing parameter [is_template]' }
    if ($p.is_template -isnot [bool]) { return Fail $script:Errors.InvalidParams 'is_template must be true or false' }
    if ($null -eq $script:Session) { return Fail $script:Errors.Internal 'Not connected' }
    $id = [string]$p.id
    $escapedId = [uri]::EscapeDataString($id)
    try { $null = Invoke-Xo GET "/vms/$escapedId" }
    catch {
        try { $null = Invoke-Xo GET "/vm-templates/$escapedId" }
        catch { return Fail $script:Errors.InvalidParams "Guest [$id] does not exist in Xen Orchestra" }
    }
    try { Set-TemplateState $id ([bool]$p.is_template) }
    catch { return Fail $script:Errors.Internal "Failed to persist template state for guest [$id]: $($_.Exception.Message)" }
    Write-Log -Level INFO -Component TEMPLATE -Message "guest=$id is_template=$([bool]$p.is_template) state_file=$($script:TemplateStatePath)"
    Ok @{ task_id = "completed:convert:$id" }
}
function Clone-Guest($p) {
    Need $p @('id','name')
    $fast = $false
    if (Has $p 'is_link_clone') {
        if ($p.is_link_clone -isnot [bool]) { return Fail $script:Errors.InvalidParams 'is_link_clone must be true or false' }
        $fast = [bool]$p.is_link_clone
    }
    $snapshotName = if (Has $p 'snapshot') { ([string]$p.snapshot).Trim() } else { '' }
    Write-Log -Level INFO -Component CLONE -Message "submit source=$($p.id) name=$($p.name) linked=$fast snapshot=$snapshotName"
    if ($snapshotName) {
        if ([string]$script:Settings.capabilities.template_method -ne 'versioning') { return Fail $script:Errors.InvalidParams 'Snapshot clones require template_method=versioning' }
        $snaps = @(Invoke-Xo GET '/vm-snapshots?fields=id,name_label,$snapshot_of')
        $snap = $snaps | Where-Object { $_.name_label -eq $snapshotName -and $_.'$snapshot_of' -eq [string]$p.id } | Select-Object -First 1
        if ($null -eq $snap) { return Fail $script:Errors.InvalidParams "Snapshot [$snapshotName] of source [$($p.id)] not found" }
        $cloneId = [string](Invoke-XoRpc 'vm.clone' @{ id=[string]$snap.id; name=[string]$p.name; full_copy=(-not $fast) })
        if ([string]::IsNullOrWhiteSpace($cloneId)) { return Fail $script:Errors.Internal 'XO JSON-RPC vm.clone returned no VM id' }
        Write-Log -Level INFO -Component CLONE -Message "completed source=$($p.id) snapshot_id=$($snap.id) clone=$cloneId"
        return Ok @{ task_id = "completed:$cloneId" }
    }
    $id = [uri]::EscapeDataString([string]$p.id)
    $body = @{ name_label = [string]$p.name; fast = $fast }
    $task = Invoke-Xo POST "/vms/$id/actions/clone" $body
    if ($null -ne $task -and (Has $task 'taskId')) { return Ok @{ task_id = [string]$task.taskId } }
    if ($null -ne $task -and (Has $task 'id')) { return Ok @{ task_id = "completed:$([string]$task.id)" } }
    Fail $script:Errors.Internal 'XOA clone response contained no task or VM id'
}
function Snapshot($p,[string]$Operation) {
    Need $p @('id','name'); $id = [uri]::EscapeDataString([string]$p.id); $name = [string]$p.name
    Write-Log -Level INFO -Component SNAPSHOT -Message "operation=$Operation vm=$($p.id) name=$name"
    if ($Operation -eq 'create') { $t = Invoke-Xo POST "/vms/$id/actions/snapshot" @{ name_label = $name }; return Ok @{ task_id = [string]$t.taskId } }
    $snaps = @(Invoke-Xo GET '/vm-snapshots?fields=id,name_label,$snapshot_of')
    $snap = $snaps | Where-Object { ($_.name_label -eq $name -or $_.nameLabel -eq $name) -and $_.'$snapshot_of' -eq [string]$p.id } | Select-Object -First 1
    if ($Operation -eq 'exists') { return Ok ($null -ne $snap) }
    if ($null -eq $snap) { return Fail $script:Errors.InvalidParams "Snapshot [$name] not found" }
    if ($Operation -eq 'delete') { $null = Invoke-Xo DELETE "/vm-snapshots/$([uri]::EscapeDataString([string]$snap.id))"; return Ok $true }
    $t = Invoke-Xo POST "/vms/$id/actions/revert_snapshot" @{ snapshotId = [string]$snap.id; snapshotBefore = $false }
    Ok @{ task_id = [string]$t.taskId }
}
function Get-Task($p) {
    Need $p @('id')
    if ([string]$p.id -match '^completed:convert:(.+)$') { return Ok @{ state = 'completed'; output = @{} } }
    if ([string]$p.id -match '^completed:(.+)$') { return Ok @{ state = 'completed'; output = @{ clone_id = $Matches[1] } } }
    $t = Invoke-Xo GET "/tasks/$([uri]::EscapeDataString([string]$p.id))"
    $state = switch ([string]$t.status) { 'success' {'completed'} 'failure' {'failed'} 'interrupted' {'failed'} default {'running'} }
    Write-Log -Level EXTENDED -Component TASK -Message "task=$($p.id) xo_status=$($t.status) cpf_state=$state"
    $r = @{ state = $state }
    if ($state -eq 'completed' -and (Has $t 'result')) {
        $cloneId = if ($null -ne $t.result -and (Has $t.result 'id')) { [string]$t.result.id } else { [string]$t.result }
        $r.output = @{ clone_id = $cloneId }
    }
    if ($state -eq 'failed') {
        $detail = if (Has $t 'error') { $t.error } elseif (Has $t 'result') { $t.result } else { $null }
        if ($detail -is [string]) { $r.error = $detail }
        elseif ($null -ne $detail -and (Has $detail 'message')) {
            $causes = if ((Has $detail 'errors') -and @($detail.errors).Count) { ": $(@($detail.errors) -join '; ')" } else { '' }
            $r.error = "$([string]$detail.message)$causes"
        }
        elseif ($null -ne $detail) { $r.error = $detail | ConvertTo-Json -Depth 10 -Compress }
        else { $r.error = 'XO task failed' }
    }
    Ok $r
}
function Dispatch($d) {
    if (-not (Has $d 'method')) { return Fail $script:Errors.InvalidRequest 'Missing method' }
    $p = if (Has $d 'params') { $d.params } else { [pscustomobject]@{} }
    $method = [string]$d.method
    $timer = [Diagnostics.Stopwatch]::StartNew()
    Write-Log -Level EXTENDED -Component CPF -Message "start method=$method"
    Write-Log -Level DEBUG -Component CPF -Message "request method=$method params=$(Format-LogObject $p)"
    try { $result = switch ($method) {
        'provider/initialize' { Initialize }
        'provider/connect' { Connect $p }
        'provider/disconnect' { Write-Log -Level INFO -Component CONNECT -Message 'disconnected'; $script:Session = $null; Ok @{ disconnected = $true } }
        'guests/list' { List-Guests }
        'guests/get' { Get-Guest $p }
        'guests/control' { Action $p }
        'guests/convert' { Convert-Guest $p }
        'guests/clone' { Clone-Guest $p }
        'tasks/get' { Get-Task $p }
        'guests/snapshots/create' { Snapshot $p create }
        'guests/snapshots/exists' { Snapshot $p exists }
        'guests/snapshots/delete' { Snapshot $p delete }
        'guests/snapshots/revert' { Snapshot $p revert }
        default { Fail $script:Errors.NotFound "Unknown method [$method]" }
    } } catch {
        Write-Log -Level ERROR -Component CPF -Message "failed method=$method elapsed_ms=$($timer.ElapsedMilliseconds) error=$($_.Exception.Message)"
        throw
    }
    $outcome = if (Has $result 'error') { 'error' } else { 'ok' }
    Write-Log -Level EXTENDED -Component CPF -Message "finish method=$method outcome=$outcome elapsed_ms=$($timer.ElapsedMilliseconds)"
    Write-Log -Level DEBUG -Component CPF -Message "response method=$method value=$(Format-LogObject $result)"
    $result
}

try { $script:Settings = Import-Settings } catch { $script:Settings = Get-Defaults; Write-Log -Level ERROR -Component SETTINGS -Message "settings error=$($_.Exception.Message)" }
try { Import-TemplateState }
catch {
    $script:TemplateStateLoadError = $_.Exception.Message
    Write-Log -Level ERROR -Component TEMPLATE -Message "state load failed path=$($script:TemplateStatePath) error=$($script:TemplateStateLoadError)"
}
Write-Log -Level INFO -Component CORE -Message "started provider=XenOrchestra powershell=$($PSVersionTable.PSVersion) template_method=$($script:Settings.capabilities.template_method) log_level=$($script:Settings.logging.log_level)"
while (-not $NoRun) {
    $line = [Console]::In.ReadLine()
    if ($null -eq $line) { break }
    if ([string]::IsNullOrWhiteSpace($line)) { continue }
    try {
        try { $data = $line.TrimStart([char]0xfeff) | ConvertFrom-Json -ErrorAction Stop }
        catch { $data = $null }
        $script:RequestCounter++
        $script:RequestId = "$PID-$('{0:D6}' -f $script:RequestCounter)"
        if ($null -eq $data) { Write-Log -Level ERROR -Component CPF -Message 'malformed JSON request'; $response = Fail $script:Errors.Parse 'Malformed JSON request' }
        else { $response = Dispatch $data }
    }
    catch [ArgumentException] { $response = Fail $script:Errors.InvalidParams $_.Exception.Message }
    catch { Write-Log -Level ERROR -Component CPF -Message $_.Exception.Message; $response = Fail $script:Errors.Internal $_.Exception.Message }
    [Console]::Out.WriteLine(($response | ConvertTo-Json -Depth 20 -Compress))
}
Write-Log -Level INFO -Component CORE -Message 'stopped provider=XenOrchestra reason=stdin_eof'
