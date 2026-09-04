<#
.SYNOPSIS
    Parallels RAS Custom Provider sample script for HPE VM Essentials (VME).
.DESCRIPTION
    Implements a Parallels RAS Custom Provider that integrates HPE VM Essentials
    through its Morpheus-based REST API. HPE VM Essentials uses the Morpheus API,
    so this provider talks to the Morpheus "instances" endpoints.

    It listens for JSON-RPC requests on standard input, processes them per the
    Custom Provider Framework (CPF) protocol, and writes responses on standard
    output.

    Supported workflow:
      - connect with an appliance URL and either a pre-created API token or
        username/password (OAuth password grant)
      - enumerate instances and retrieve guest info (power state, IPs, template flag)
      - power operations: start, stop, restart, reset (-> restart), suspend, delete
      - convert an instance to/from a RAS template (tracked with an instance label)
      - template versioning using Morpheus instance snapshots
      - clone an instance into a new instance
      - asynchronous task status through tasks/get

    State values returned to RAS: powered_on, powered_off, powering_on,
    powering_off, suspended, suspending.

    Requires PowerShell 7 or later.
.NOTES
    File Name : Parallels-RAS-CFP-HPE-VME.ps1
    Platform  : HPE VM Essentials (Morpheus API)
    Author    : Edwin Houben
    Reference : hpe-vme/HPE-VME-API.md in this folder
.EXAMPLE
    {"method":"provider/initialize"}
    {"method":"provider/connect","params":{"settings":{"url":"https://vme.example.com","token":"<api-token>"}}}
    {"method":"provider/connect","params":{"settings":{"url":"https://vme.example.com","username":"admin","password":"secret"}}}
    {"method":"guests/list"}
    {"method":"guests/control","params":{"id":"42","control":"start"}}
    {"method":"guests/snapshots/create","params":{"id":"42","name":"RAS_TEMPLATE_VERSION_1"}}
    {"method":"guests/clone","params":{"id":"42","name":"vdi-clone-01"}}
#>

Set-StrictMode -Version Latest

$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'
$WarningPreference     = 'SilentlyContinue'
$VerbosePreference     = 'SilentlyContinue'
$InformationPreference = 'SilentlyContinue'

if ($Host.Name -notmatch 'ISE') {
    [Console]::InputEncoding  = [System.Text.Encoding]::UTF8
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
}

$stdout = [Console]::OpenStandardOutput()
$writer = New-Object System.IO.StreamWriter($stdout, [System.Text.Encoding]::UTF8)
$writer.AutoFlush = $true

$script:ProviderNamePrefix = 'HPE-VME:'
$script:LogPath            = Join-Path ([System.IO.Path]::GetTempPath()) 'HPE-VME-RAS-Provider.log'
$script:Session            = $null
$script:TemplateLabel      = 'ras-template'

$script:ErrorCodes = @{
    ParseError     = -32700
    MethodNotFound = -32601
    InvalidParams  = -32602
    InternalError  = -32603
}

# ----------------------------------------------------------------------------
# Logging and protocol helpers
# ----------------------------------------------------------------------------

function Write-DebugLog {
    param([string]$Message)
    try {
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
        Add-Content -Path $script:LogPath -Value "$timestamp $Message" -Encoding UTF8
    }
    catch { }
}

function Send-Response {
    param([Parameter(Mandatory = $true)][object]$ResponseObject)
    try {
        $json = $ResponseObject | ConvertTo-Json -Compress -Depth 20
        $writer.WriteLine($json)
        Write-DebugLog "OUT: $json"
    }
    catch {
        $fallback = @{
            error = @{ code = $script:ErrorCodes.InternalError; message = "$($script:ProviderNamePrefix) Failed to serialize response: $($_.Exception.Message)" }
        } | ConvertTo-Json -Compress -Depth 10
        $writer.WriteLine($fallback)
        Write-DebugLog "OUT-FALLBACK: $fallback"
    }
}

function New-ErrorResponse {
    param([int]$Code, [string]$Message)
    return @{ error = @{ code = $Code; message = $Message } }
}

function ConvertFrom-JsonSafe {
    param([string]$InputLine)
    try { return $InputLine | ConvertFrom-Json -ErrorAction Stop }
    catch { Write-DebugLog "JSON parse failed: $($_.Exception.Message)"; return $null }
}

function Test-RequiredFields {
    param([object]$Data, [string[]]$RequiredFields)
    foreach ($field in $RequiredFields) {
        $keys  = $field -split '\.'
        $value = $Data
        foreach ($key in $keys) {
            if ($null -ne $value -and $value.PSObject.Properties.Name -contains $key) { $value = $value.$key }
            else { return "$($script:ProviderNamePrefix) Missing field: $field" }
        }
    }
    return $null
}

function Test-HasProperty {
    param([object]$Object, [string]$Name)
    return ($null -ne $Object -and $Object.PSObject.Properties.Name -contains $Name)
}

# ----------------------------------------------------------------------------
# Session and REST plumbing
# ----------------------------------------------------------------------------

function Get-Session {
    if ($null -eq $script:Session) { throw 'Session not initialized' }
    if ([string]::IsNullOrWhiteSpace($script:Session.base)) { throw 'Session base URL missing' }
    if ($null -eq $script:Session.header) { throw 'Session header missing' }
    return $script:Session
}

function ConvertTo-BaseUrl {
    param([string]$UrlValue)
    $u = $UrlValue.Trim()
    if ($u -notmatch '^https?://') { $u = "https://$u" }
    return $u.TrimEnd('/')
}

function Invoke-VmeApi {
    <#
        Calls the Morpheus / HPE VME REST API with the session bearer token.
        $Path is relative to the appliance base URL (e.g. '/api/instances').
        Throws on HTTP error. Use Get-VmeResourceOrNull for 404-tolerant reads.
    #>
    param(
        [Parameter(Mandatory = $true)][ValidateSet('GET', 'POST', 'PUT', 'DELETE')][string]$Method,
        [Parameter(Mandatory = $true)][string]$Path,
        [object]$Body = $null
    )

    $session = Get-Session
    $uri = ($session.base + '/' + $Path.TrimStart('/'))

    $params = @{
        Uri         = $uri
        Headers     = $session.header
        Method      = $Method
        ErrorAction = 'Stop'
        ContentType = 'application/json'
    }
    if ($PSVersionTable.PSEdition -eq 'Core') {
        if ($session.skip_tls) { $params.SkipCertificateCheck = $true }
        $params.SkipHeaderValidation = $true
    }
    if ($null -ne $Body) {
        $params.Body = if ($Body -is [string]) { $Body } else { $Body | ConvertTo-Json -Compress -Depth 32 }
    }

    Write-DebugLog "HTTP $Method $uri$(if ($null -ne $Body) { ' - Body: ' + $params.Body })"
    try { return Invoke-RestMethod @params }
    catch { Write-DebugLog "HTTP failure ($Method $Path): $($_.Exception.Message)"; throw }
}

function Get-HttpStatusCode {
    param([System.Management.Automation.ErrorRecord]$ErrorRecord)
    try {
        $resp = $ErrorRecord.Exception.Response
        if ($null -ne $resp -and $null -ne $resp.StatusCode) { return [int]$resp.StatusCode }
    }
    catch { }
    return 0
}

function Get-VmeResourceOrNull {
    param([Parameter(Mandatory = $true)][string]$Path)
    try { return Invoke-VmeApi -Method GET -Path $Path }
    catch {
        if ((Get-HttpStatusCode -ErrorRecord $_) -eq 404) { return $null }
        throw
    }
}

function Get-AccessTokenFromCredentials {
    param([string]$Base, [string]$Username, [string]$Password, [string]$ClientId, [bool]$SkipTls)

    $form = @{
        grant_type = 'password'
        scope      = 'write'
        client_id  = $ClientId
        username   = $Username
        password   = $Password
    }
    $params = @{
        Uri         = "$Base/oauth/token"
        Method      = 'POST'
        Body        = $form
        ContentType = 'application/x-www-form-urlencoded'
        ErrorAction = 'Stop'
    }
    if ($PSVersionTable.PSEdition -eq 'Core' -and $SkipTls) { $params.SkipCertificateCheck = $true }

    $resp = Invoke-RestMethod @params
    if (-not (Test-HasProperty -Object $resp -Name 'access_token')) {
        throw 'OAuth response did not contain an access_token'
    }
    return [string]$resp.access_token
}

# ----------------------------------------------------------------------------
# Instance read helpers
# ----------------------------------------------------------------------------

function Get-InstanceList {
    $resp = Invoke-VmeApi -Method GET -Path '/api/instances?max=1000'
    if (Test-HasProperty -Object $resp -Name 'instances') { return @($resp.instances) }
    return @()
}

function Get-Instance {
    param([Parameter(Mandatory = $true)][string]$Id)
    $resp = Get-VmeResourceOrNull -Path "/api/instances/$Id"
    if ($null -ne $resp -and (Test-HasProperty -Object $resp -Name 'instance')) { return $resp.instance }
    return $null
}

function Map-StatusToRasState {
    param([string]$Status)
    $s = if ($null -ne $Status) { $Status.ToString().Trim().ToLowerInvariant() } else { '' }
    switch ($s) {
        'running'      { return 'powered_on' }
        'stopped'      { return 'powered_off' }
        'suspended'    { return 'suspended' }
        'starting'     { return 'powering_on' }
        'provisioning' { return 'powering_on' }
        'deploying'    { return 'powering_on' }
        'pending'      { return 'powering_on' }
        'resizing'     { return 'powering_on' }
        'stopping'     { return 'powering_off' }
        'removing'     { return 'powering_off' }
        default        { return 'powered_off' }
    }
}

function Get-InstanceIpAddresses {
    param([object]$Instance)
    $ips = New-Object 'System.Collections.Generic.List[string]'

    if (Test-HasProperty -Object $Instance -Name 'connectionInfo') {
        foreach ($ci in @($Instance.connectionInfo)) {
            if (Test-HasProperty -Object $ci -Name 'ip' -and -not [string]::IsNullOrWhiteSpace([string]$ci.ip)) {
                $ip = [string]$ci.ip
                if ($ip -notmatch ':' -and -not $ips.Contains($ip)) { $ips.Add($ip) }
            }
        }
    }

    # Best-effort: pull IPs from container/server detail when present.
    if (Test-HasProperty -Object $Instance -Name 'containerDetails') {
        foreach ($cd in @($Instance.containerDetails)) {
            foreach ($field in @('externalIp', 'internalIp', 'ip')) {
                if (Test-HasProperty -Object $cd -Name $field -and -not [string]::IsNullOrWhiteSpace([string]$cd.$field)) {
                    $ip = [string]$cd.$field
                    if ($ip -notmatch ':' -and -not $ips.Contains($ip)) { $ips.Add($ip) }
                }
            }
        }
    }

    return @($ips | Select-Object -First 3)
}

function Get-InstanceIsTemplate {
    param([object]$Instance)
    if (Test-HasProperty -Object $Instance -Name 'labels') {
        foreach ($label in @($Instance.labels)) {
            if ([string]$label -eq $script:TemplateLabel) { return $true }
        }
    }
    return $false
}

function Get-InstanceOsType {
    param([object]$Instance)
    foreach ($path in @('instanceTypeName', 'layoutName')) {
        if (Test-HasProperty -Object $Instance -Name $path -and -not [string]::IsNullOrWhiteSpace([string]$Instance.$path)) {
            return [string]$Instance.$path
        }
    }
    if (Test-HasProperty -Object $Instance -Name 'layout' -and (Test-HasProperty -Object $Instance.layout -Name 'name')) {
        return [string]$Instance.layout.name
    }
    return 'unknown'
}

function ConvertTo-RasGuestObject {
    param([Parameter(Mandatory = $true)][string]$Id)

    $inst = Get-Instance -Id $Id
    if ($null -eq $inst) { throw "Instance [$Id] not found" }

    $status = if (Test-HasProperty -Object $inst -Name 'status') { [string]$inst.status } else { '' }
    $name   = if (Test-HasProperty -Object $inst -Name 'name') { [string]$inst.name } else { "instance-$Id" }
    $ips    = Get-InstanceIpAddresses -Instance $inst

    $guest = @{
        id            = [string]$Id
        name          = $name
        provider      = 'HPE VM Essentials'
        state         = (Map-StatusToRasState -Status $status)
        power_state   = $(if ([string]::IsNullOrWhiteSpace($status)) { 'unknown' } else { $status })
        host_os       = (Get-InstanceOsType -Instance $inst)
        ip            = $(if ($ips.Count -gt 0) { $ips[0] } else { $null })
        ip_addresses  = @($ips)
        mac_addresses = @()
        is_template   = (Get-InstanceIsTemplate -Instance $inst)
        type          = 'Virtual Machine'
    }

    Write-DebugLog ("GUEST id={0}; name={1}; state={2}; status={3}; template={4}; ips={5}" -f `
            $guest.id, $guest.name, $guest.state, $guest.power_state, $guest.is_template, ($guest.ip_addresses -join ','))
    return $guest
}

function Find-InstanceSnapshot {
    param([string]$InstanceId, [string]$SnapshotName)
    $resp = Invoke-VmeApi -Method GET -Path "/api/instances/$InstanceId/snapshots"
    if (-not (Test-HasProperty -Object $resp -Name 'snapshots')) { return $null }
    foreach ($snap in @($resp.snapshots)) {
        if (Test-HasProperty -Object $snap -Name 'name' -and [string]$snap.name -eq $SnapshotName) { return $snap }
    }
    return $null
}

function Get-ControlAction {
    param([Parameter(Mandatory = $true)][string]$Control)
    switch ($Control.Trim().ToLowerInvariant()) {
        'start'   { return 'start' }
        'stop'    { return 'stop' }
        'restart' { return 'restart' }
        'reset'   { return 'restart' }   # Morpheus has no hard reset; restart is the closest action.
        'suspend' { return 'suspend' }
        'delete'  { return 'delete' }
        default   { return $null }
    }
}

# ----------------------------------------------------------------------------
# Method handlers
# ----------------------------------------------------------------------------

function Handle-Initialize {
    return @{
        result = @{
            version      = '1.0.0'
            capabilities = @{
                can_suspend_guests    = $true
                guests_polling_rate   = 5
                tasks_polling_rate    = 10
                tasks_polling_retries = 180
                template_method       = 'versioning'
                can_link_clones       = $false
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

    $urlValue  = if (Test-HasProperty -Object $settings -Name 'url') { [string]$settings.url } else { '' }
    $token     = if (Test-HasProperty -Object $settings -Name 'token') { [string]$settings.token } else { '' }
    $username  = if (Test-HasProperty -Object $settings -Name 'username') { [string]$settings.username } else { '' }
    $password  = if (Test-HasProperty -Object $settings -Name 'password') { [string]$settings.password } else { '' }
    $clientId  = if (Test-HasProperty -Object $settings -Name 'client_id') { [string]$settings.client_id } else { 'morph-api' }
    $skipTls   = $true
    if (Test-HasProperty -Object $settings -Name 'skip_tls') {
        try { $skipTls = [System.Convert]::ToBoolean($settings.skip_tls) } catch { $skipTls = $true }
    }

    if ([string]::IsNullOrWhiteSpace($urlValue)) {
        return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) url is required"
    }
    if ([string]::IsNullOrWhiteSpace($token) -and
        ([string]::IsNullOrWhiteSpace($username) -or [string]::IsNullOrWhiteSpace($password))) {
        return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) provide either token, or username and password"
    }

    try {
        $base = ConvertTo-BaseUrl -UrlValue $urlValue
        if ([string]::IsNullOrWhiteSpace($token)) {
            $token = Get-AccessTokenFromCredentials -Base $base -Username $username -Password $password -ClientId $clientId -SkipTls $skipTls
        }

        $script:Session = @{
            base     = $base
            skip_tls = $skipTls
            header   = @{ Authorization = "Bearer $token" }
        }

        # Validate the token by listing a single instance.
        $null = Invoke-VmeApi -Method GET -Path '/api/instances?max=1'

        Write-DebugLog "Connected to $base"
        return @{ result = @{ message = "$($script:ProviderNamePrefix) Connected to $base" } }
    }
    catch {
        $script:Session = $null
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to connect: $($_.Exception.Message)"
    }
}

function Handle-Disconnect {
    $script:Session = $null
    return @{ result = @{} }
}

function Handle-GuestList {
    try {
        $guests = @()
        foreach ($inst in Get-InstanceList) {
            if (Test-HasProperty -Object $inst -Name 'id') { $guests += [string]$inst.id }
        }
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
            return @{ result = (ConvertTo-RasGuestObject -Id ([string]$ids[0])) }
        }

        $resultMap = @{}
        foreach ($id in $ids) {
            $gid = [string]$id
            try { $resultMap[$gid] = ConvertTo-RasGuestObject -Id $gid }
            catch {
                Write-DebugLog "Guest get failed for [$gid]: $($_.Exception.Message)"
                $resultMap[$gid] = @{
                    id = $gid; name = $gid; provider = 'HPE VM Essentials'
                    state = 'powered_off'; power_state = 'unknown'; host_os = 'unknown'
                    ip = $null; ip_addresses = @(); mac_addresses = @(); is_template = $false; type = 'Virtual Machine'
                }
            }
        }
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
        $id     = [string]$Params.id
        $action = Get-ControlAction -Control ([string]$Params.control)
        if ([string]::IsNullOrWhiteSpace($action)) {
            return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Unsupported guest control: $($Params.control)"
        }

        if ($action -eq 'delete') {
            Invoke-VmeApi -Method DELETE -Path "/api/instances/$id" | Out-Null
        }
        else {
            Invoke-VmeApi -Method PUT -Path "/api/instances/$id/$action" | Out-Null
        }
        return @{ result = @{} }
    }
    catch {
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to control guest [$($Params.control)]: $($_.Exception.Message)"
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
        $id         = [string]$Params.id
        $isTemplate = [System.Convert]::ToBoolean($Params.is_template)

        $inst = Get-Instance -Id $id
        if ($null -eq $inst) { throw "Instance [$id] not found" }

        $labels = New-Object 'System.Collections.Generic.List[string]'
        if (Test-HasProperty -Object $inst -Name 'labels') {
            foreach ($l in @($inst.labels)) {
                if (-not [string]::IsNullOrWhiteSpace([string]$l)) { $labels.Add([string]$l) }
            }
        }

        if ($isTemplate) {
            if (-not $labels.Contains($script:TemplateLabel)) { $labels.Add($script:TemplateLabel) }
        }
        else {
            [void]$labels.Remove($script:TemplateLabel)
        }

        Invoke-VmeApi -Method PUT -Path "/api/instances/$id" -Body @{ instance = @{ labels = @($labels) } } | Out-Null

        if ($isTemplate) {
            try { Invoke-VmeApi -Method PUT -Path "/api/instances/$id/stop" | Out-Null }
            catch { Write-DebugLog "Stop on convert ignored for [$id]: $($_.Exception.Message)" }
        }

        return @{ result = @{ task_id = "convert:$id" } }
    }
    catch {
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to convert guest [$($Params.id)]: $($_.Exception.Message)"
    }
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
        $id        = [string]$Params.id
        $cloneName = [string]$Params.name

        # Morpheus clones from the instance's backup snapshots. A specific named
        # snapshot cannot be selected through this endpoint; RAS reverts the
        # template to the desired version during maintenance before cloning.
        Invoke-VmeApi -Method PUT -Path "/api/instances/$id/clone" -Body @{ name = $cloneName } | Out-Null

        return @{ result = @{ task_id = "clone:$cloneName" } }
    }
    catch {
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to clone guest [$($Params.id)]: $($_.Exception.Message)"
    }
}

function Handle-GuestSnapshotsCreate {
    param([object]$Params)

    if ($null -eq $Params -or [string]::IsNullOrWhiteSpace([string]$Params.id)) {
        return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Invalid guest id"
    }
    if ([string]::IsNullOrWhiteSpace([string]$Params.name)) {
        return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Invalid snapshot name"
    }

    try {
        $id      = [string]$Params.id
        $rasName = [string]$Params.name
        Invoke-VmeApi -Method PUT -Path "/api/instances/$id/snapshot" -Body @{ snapshot = @{ name = $rasName } } | Out-Null
        return @{ result = @{ task_id = "snapshot:${id}:$rasName" } }
    }
    catch {
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to create snapshot [$($Params.name)] for guest [$($Params.id)]: $($_.Exception.Message)"
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

    try {
        $id   = [string]$Params.id
        $snap = Find-InstanceSnapshot -InstanceId $id -SnapshotName ([string]$Params.name)
        if ($null -eq $snap) { return @{ result = @{ task_id = "noop:deleted" } } }

        $snapId = [string]$snap.id
        Invoke-VmeApi -Method DELETE -Path "/api/snapshots/$snapId" | Out-Null
        return @{ result = @{ task_id = "snapshot-delete:$snapId" } }
    }
    catch {
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to delete snapshot [$($Params.name)] for guest [$($Params.id)]: $($_.Exception.Message)"
    }
}

function Handle-GuestSnapshotsExists {
    param([object]$Params)

    if ($null -eq $Params -or [string]::IsNullOrWhiteSpace([string]$Params.id)) {
        return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Invalid guest id"
    }
    if ([string]::IsNullOrWhiteSpace([string]$Params.name)) {
        return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Invalid snapshot name"
    }

    try {
        $snap = Find-InstanceSnapshot -InstanceId ([string]$Params.id) -SnapshotName ([string]$Params.name)
        return @{ result = ($null -ne $snap) }
    }
    catch {
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to check snapshot [$($Params.name)] for guest [$($Params.id)]: $($_.Exception.Message)"
    }
}

function Handle-GuestSnapshotsRevert {
    param([object]$Params)

    if ($null -eq $Params -or [string]::IsNullOrWhiteSpace([string]$Params.id)) {
        return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Invalid guest id"
    }
    if ([string]::IsNullOrWhiteSpace([string]$Params.name)) {
        return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Invalid snapshot name"
    }

    try {
        $id   = [string]$Params.id
        $snap = Find-InstanceSnapshot -InstanceId $id -SnapshotName ([string]$Params.name)
        if ($null -eq $snap) {
            return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Snapshot [$($Params.name)] not found for instance [$id]"
        }

        $snapId = [string]$snap.id
        Invoke-VmeApi -Method PUT -Path "/api/instances/$id/revert-snapshot/$snapId" | Out-Null
        return @{ result = @{ task_id = "revert:$id" } }
    }
    catch {
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to revert snapshot [$($Params.name)] for guest [$($Params.id)]: $($_.Exception.Message)"
    }
}

function Handle-TaskInfo {
    param([object]$Params)

    if ($null -eq $Params -or [string]::IsNullOrWhiteSpace([string]$Params.id)) {
        return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Invalid task id"
    }

    try {
        $taskId = [string]$Params.id
        $parts  = $taskId -split ':', 3
        $kind   = $parts[0]

        switch ($kind) {
            'convert'         { return @{ result = @{ state = 'completed'; output = @{} } } }
            'revert'          { return @{ result = @{ state = 'completed'; output = @{} } } }
            'snapshot-delete' { return @{ result = @{ state = 'completed'; output = @{} } } }
            'noop'            { return @{ result = @{ state = 'completed'; output = @{} } } }

            'snapshot' {
                $instId = if ($parts.Count -ge 2) { $parts[1] } else { '' }
                $name   = if ($parts.Count -ge 3) { $parts[2] } else { '' }
                $snap = Find-InstanceSnapshot -InstanceId $instId -SnapshotName $name
                if ($null -eq $snap) { return @{ result = @{ state = 'running' } } }
                $status = if (Test-HasProperty -Object $snap -Name 'status') { ([string]$snap.status).ToLowerInvariant() } else { '' }
                switch ($status) {
                    'complete' { return @{ result = @{ state = 'completed'; output = @{} } } }
                    'ready'    { return @{ result = @{ state = 'completed'; output = @{} } } }
                    'failed'   { return @{ result = @{ state = 'failed'; error = @{ code = 1; message = 'Snapshot failed' } } } }
                    'errored'  { return @{ result = @{ state = 'failed'; error = @{ code = 1; message = 'Snapshot failed' } } } }
                    default    { return @{ result = @{ state = 'running' } } }
                }
            }

            'clone' {
                $name = if ($parts.Count -ge 2) { $taskId.Substring($taskId.IndexOf(':') + 1) } else { '' }
                $match = $null
                foreach ($inst in Get-InstanceList) {
                    if ((Test-HasProperty -Object $inst -Name 'name') -and [string]$inst.name -eq $name) { $match = $inst; break }
                }
                if ($null -eq $match) { return @{ result = @{ state = 'running' } } }

                $status = if (Test-HasProperty -Object $match -Name 'status') { (Map-StatusToRasState -Status ([string]$match.status)) } else { 'powered_off' }
                if ($status -eq 'powering_on') { return @{ result = @{ state = 'running' } } }
                return @{ result = @{ state = 'completed'; output = @{ clone_id = [string]$match.id } } }
            }

            default { return @{ result = @{ state = 'completed'; output = @{} } } }
        }
    }
    catch {
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to retrieve task info: $($_.Exception.Message)"
    }
}

# Host methods are aliases of guest methods for this provider.
function Handle-HostList    { return Handle-GuestList }
function Handle-HostGet     { param([object]$Params) return Handle-GuestGet -Params $Params }
function Handle-HostControl { param([object]$Params) return Handle-GuestControl -Params $Params }

# ----------------------------------------------------------------------------
# Dispatch
# ----------------------------------------------------------------------------

$script:MethodRegistry = @{
    'provider/initialize'     = @{ Handler = { param($data) Handle-Initialize }; RequiredFields = @() }
    'provider/connect'        = @{ Handler = { param($data) Handle-Connect -Params $data.params }; RequiredFields = @('params.settings') }
    'provider/disconnect'     = @{ Handler = { param($data) Handle-Disconnect }; RequiredFields = @() }

    'hosts/list'              = @{ Handler = { param($data) Handle-HostList }; RequiredFields = @() }
    'hosts/get'               = @{ Handler = { param($data) Handle-HostGet -Params $data.params }; RequiredFields = @('params.id') }
    'hosts/control'           = @{ Handler = { param($data) Handle-HostControl -Params $data.params }; RequiredFields = @('params.id', 'params.control') }

    'guests/list'             = @{ Handler = { param($data) Handle-GuestList }; RequiredFields = @() }
    'guests/get'              = @{ Handler = { param($data) Handle-GuestGet -Params $data.params }; RequiredFields = @('params.id') }
    'guests/control'          = @{ Handler = { param($data) Handle-GuestControl -Params $data.params }; RequiredFields = @('params.id', 'params.control') }
    'guests/convert'          = @{ Handler = { param($data) Handle-GuestConvert -Params $data.params }; RequiredFields = @('params.id', 'params.is_template') }
    'guests/clone'            = @{ Handler = { param($data) Handle-GuestClone -Params $data.params }; RequiredFields = @('params.id', 'params.name') }
    'guests/snapshots/create' = @{ Handler = { param($data) Handle-GuestSnapshotsCreate -Params $data.params }; RequiredFields = @('params.id', 'params.name') }
    'guests/snapshots/delete' = @{ Handler = { param($data) Handle-GuestSnapshotsDelete -Params $data.params }; RequiredFields = @('params.id', 'params.name') }
    'guests/snapshots/exists' = @{ Handler = { param($data) Handle-GuestSnapshotsExists -Params $data.params }; RequiredFields = @('params.id', 'params.name') }
    'guests/snapshots/revert' = @{ Handler = { param($data) Handle-GuestSnapshotsRevert -Params $data.params }; RequiredFields = @('params.id', 'params.name') }
    'tasks/get'               = @{ Handler = { param($data) Handle-TaskInfo -Params $data.params }; RequiredFields = @('params.id') }
}

function Process-Method {
    param([string]$InputLine)

    $methodData = ConvertFrom-JsonSafe -InputLine $InputLine
    if ($null -eq $methodData) {
        return New-ErrorResponse -Code $script:ErrorCodes.ParseError -Message "$($script:ProviderNamePrefix) Invalid JSON format"
    }

    $methodName = $null
    if ($methodData.PSObject.Properties.Name -contains 'method') { $methodName = [string]$methodData.method }
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

    try { return & $methodEntry.Handler $methodData }
    catch {
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Method execution failed: $($_.Exception.Message)"
    }
}

Write-DebugLog "HPE VME provider process started. PID=$PID"

while ($true) {
    try {
        $inputLine = [Console]::In.ReadLine()
        if ($null -eq $inputLine) { Write-DebugLog 'Input stream closed. Exiting.'; break }
        Write-DebugLog "IN (PID=$PID): $inputLine"
        $response = Process-Method -InputLine ($inputLine.Trim())
        Send-Response -ResponseObject $response
    }
    catch {
        Send-Response -ResponseObject (New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to process input: $($_.Exception.Message)")
    }
}
