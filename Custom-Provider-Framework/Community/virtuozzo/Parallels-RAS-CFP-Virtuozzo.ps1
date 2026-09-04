<#
.SYNOPSIS
    Parallels RAS Custom Provider sample script for Virtuozzo Hybrid Infrastructure.
.DESCRIPTION
    Implements a Parallels RAS Custom Provider that integrates Virtuozzo Hybrid
    Infrastructure (VHI) through its OpenStack-compatible API. VHI ships a 100%
    upstream-compatible OpenStack control plane, so this provider authenticates
    with Keystone (v3) and manages VMs with the Nova compute API and Glance image
    API.

    It listens for JSON-RPC requests on standard input, processes them per the
    Custom Provider Framework (CPF) protocol, and writes responses on standard
    output.

    Because OpenStack represents point-in-time copies as Glance images rather than
    in-place VM snapshots, the CPF template/snapshot methods are mapped onto
    images, as described in the CPF "Capabilities" documentation:
      - convert to template  -> snapshot the server to an image, tag the server
      - snapshot create/exists/delete -> Glance images named after the RAS snapshot
      - snapshot revert      -> rebuild the server from the image
      - clone                -> boot a new server from the image

    State values returned to RAS: powered_on, powered_off, powering_on,
    powering_off, suspended, suspending.

    Requires PowerShell 7 or later.
.NOTES
    File Name : Parallels-RAS-CFP-Virtuozzo.ps1
    Platform  : Virtuozzo Hybrid Infrastructure (OpenStack API)
    Author    : Edwin Houben
    Reference : virtuozzo/Virtuozzo-API.md in this folder
.EXAMPLE
    {"method":"provider/initialize"}
    {"method":"provider/connect","params":{"settings":{"auth_url":"https://vhi.example.com:5000/v3","username":"admin","password":"secret","project_name":"vdi"}}}
    {"method":"guests/list"}
    {"method":"guests/control","params":{"id":"<server-id>","control":"start"}}
    {"method":"guests/snapshots/create","params":{"id":"<server-id>","name":"RAS_TEMPLATE_VERSION_1"}}
    {"method":"guests/clone","params":{"id":"<server-id>","name":"vdi-clone-01","snapshot":"RAS_TEMPLATE_VERSION_1"}}
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

$script:ProviderNamePrefix = 'Virtuozzo:'
$script:LogPath            = Join-Path ([System.IO.Path]::GetTempPath()) 'Virtuozzo-RAS-Provider.log'
$script:Session            = $null
$script:TemplateMetaKey    = 'ras_template'
$script:SourceMetaKey      = 'ras_source_server'

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
    if ([string]::IsNullOrWhiteSpace($script:Session.token)) { throw 'Session token missing' }
    if ([string]::IsNullOrWhiteSpace($script:Session.compute)) { throw 'Compute endpoint missing' }
    return $script:Session
}

function Invoke-OsApi {
    <#
        Calls an OpenStack service with the Keystone token. $BaseUrl is the
        service endpoint (compute or image); $Path is appended.
        Throws on HTTP error. Use Get-OsResourceOrNull for 404-tolerant reads.
        Returns a hashtable: @{ body = <object>; headers = <response headers> }.
    #>
    param(
        [Parameter(Mandatory = $true)][ValidateSet('GET', 'POST', 'PUT', 'DELETE')][string]$Method,
        [Parameter(Mandatory = $true)][string]$BaseUrl,
        [Parameter(Mandatory = $true)][string]$Path,
        [object]$Body = $null
    )

    $session = Get-Session
    $uri = ($BaseUrl.TrimEnd('/') + '/' + $Path.TrimStart('/'))

    $irm = @{
        Uri                     = $uri
        Method                  = $Method
        Headers                 = @{ 'X-Auth-Token' = $session.token }
        ContentType             = 'application/json'
        ErrorAction             = 'Stop'
        ResponseHeadersVariable = 'respHeaders'
    }
    if ($PSVersionTable.PSEdition -eq 'Core' -and $session.skip_tls) { $irm.SkipCertificateCheck = $true }
    if ($null -ne $Body) { $irm.Body = if ($Body -is [string]) { $Body } else { $Body | ConvertTo-Json -Compress -Depth 32 } }

    Write-DebugLog "HTTP $Method $uri"
    $body = Invoke-RestMethod @irm
    return @{ body = $body; headers = $respHeaders }
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

function Get-OsResourceOrNull {
    param([Parameter(Mandatory = $true)][string]$BaseUrl, [Parameter(Mandatory = $true)][string]$Path)
    try { return (Invoke-OsApi -Method GET -BaseUrl $BaseUrl -Path $Path).body }
    catch {
        if ((Get-HttpStatusCode -ErrorRecord $_) -eq 404) { return $null }
        throw
    }
}

function Get-EndpointFromCatalog {
    param([object]$Catalog, [string]$Type, [string]$Region)
    foreach ($svc in @($Catalog)) {
        if (-not (Test-HasProperty -Object $svc -Name 'type') -or [string]$svc.type -ne $Type) { continue }
        $public = $null
        foreach ($ep in @($svc.endpoints)) {
            if ((Test-HasProperty -Object $ep -Name 'interface') -and [string]$ep.interface -eq 'public') {
                if ([string]::IsNullOrWhiteSpace($Region) -or
                    ((Test-HasProperty -Object $ep -Name 'region') -and [string]$ep.region -eq $Region)) {
                    return [string]$ep.url
                }
                if ($null -eq $public) { $public = [string]$ep.url }
            }
        }
        if ($null -ne $public) { return $public }
    }
    return $null
}

# ----------------------------------------------------------------------------
# Server read helpers
# ----------------------------------------------------------------------------

function Get-Server {
    param([Parameter(Mandatory = $true)][string]$Id)
    $session = Get-Session
    $resp = Get-OsResourceOrNull -BaseUrl $session.compute -Path "/servers/$Id"
    if ($null -ne $resp -and (Test-HasProperty -Object $resp -Name 'server')) { return $resp.server }
    return $null
}

function Map-OsStatusToRasState {
    param([string]$Status)
    switch (($Status | ForEach-Object { $_.ToString().Trim().ToUpperInvariant() })) {
        'ACTIVE'             { return 'powered_on' }
        'SHUTOFF'            { return 'powered_off' }
        'SUSPENDED'          { return 'suspended' }
        'PAUSED'             { return 'suspended' }
        'BUILD'              { return 'powering_on' }
        'REBUILD'            { return 'powering_on' }
        'HARD_REBOOT'        { return 'powering_on' }
        'REBOOT'             { return 'powering_on' }
        'POWERING-ON'        { return 'powering_on' }
        'POWERING-OFF'       { return 'powering_off' }
        'SHELVED'            { return 'powered_off' }
        'SHELVED_OFFLOADED'  { return 'powered_off' }
        default              { return 'powered_off' }
    }
}

function Get-ServerNetworkData {
    param([object]$Server)
    $ips  = New-Object 'System.Collections.Generic.List[string]'
    $macs = New-Object 'System.Collections.Generic.List[string]'

    if (Test-HasProperty -Object $Server -Name 'addresses' -and $null -ne $Server.addresses) {
        foreach ($net in $Server.addresses.PSObject.Properties) {
            foreach ($entry in @($net.Value)) {
                if ((Test-HasProperty -Object $entry -Name 'version') -and [int]$entry.version -eq 4 -and
                    (Test-HasProperty -Object $entry -Name 'addr') -and -not [string]::IsNullOrWhiteSpace([string]$entry.addr)) {
                    $ip = [string]$entry.addr
                    if (-not $ips.Contains($ip)) { $ips.Add($ip) }
                }
                if (Test-HasProperty -Object $entry -Name 'OS-EXT-IPS-MAC:mac_addr') {
                    $mac = ([string]$entry.'OS-EXT-IPS-MAC:mac_addr').ToUpperInvariant()
                    if (-not [string]::IsNullOrWhiteSpace($mac) -and -not $macs.Contains($mac)) { $macs.Add($mac) }
                }
            }
        }
    }

    return @{ IPv4Addresses = @($ips | Select-Object -First 3); MacAddresses = @($macs | Select-Object -First 3) }
}

function Get-ServerIsTemplate {
    param([object]$Server)
    if ((Test-HasProperty -Object $Server -Name 'metadata') -and
        (Test-HasProperty -Object $Server.metadata -Name $script:TemplateMetaKey)) {
        return ([string]$Server.metadata.$($script:TemplateMetaKey) -eq 'true')
    }
    return $false
}

function ConvertTo-RasGuestObject {
    param([Parameter(Mandatory = $true)][string]$Id)

    $srv = Get-Server -Id $Id
    if ($null -eq $srv) { throw "Server [$Id] not found" }

    $status = if (Test-HasProperty -Object $srv -Name 'status') { [string]$srv.status } else { '' }
    $name   = if (Test-HasProperty -Object $srv -Name 'name') { [string]$srv.name } else { "server-$Id" }
    $net    = Get-ServerNetworkData -Server $srv

    $osType = 'unknown'
    if ((Test-HasProperty -Object $srv -Name 'metadata') -and (Test-HasProperty -Object $srv.metadata -Name 'os_distro')) {
        $osType = [string]$srv.metadata.os_distro
    }

    $guest = @{
        id            = [string]$Id
        name          = $name
        provider      = 'Virtuozzo'
        state         = (Map-OsStatusToRasState -Status $status)
        power_state   = $(if ([string]::IsNullOrWhiteSpace($status)) { 'unknown' } else { $status })
        host_os       = $osType
        ip            = $(if ($net.IPv4Addresses.Count -gt 0) { $net.IPv4Addresses[0] } else { $null })
        ip_addresses  = @($net.IPv4Addresses)
        mac_addresses = @($net.MacAddresses)
        is_template   = (Get-ServerIsTemplate -Server $srv)
        type          = 'Virtual Machine'
    }

    Write-DebugLog ("GUEST id={0}; name={1}; state={2}; status={3}; template={4}; ips={5}" -f `
            $guest.id, $guest.name, $guest.state, $guest.power_state, $guest.is_template, ($guest.ip_addresses -join ','))
    return $guest
}

function Post-ServerAction {
    param([string]$Id, [object]$Action)
    $session = Get-Session
    return Invoke-OsApi -Method POST -BaseUrl $session.compute -Path "/servers/$Id/action" -Body $Action
}

# ----------------------------------------------------------------------------
# Glance image helpers (templates / snapshots)
# ----------------------------------------------------------------------------

function Find-Image {
    param([string]$SourceServerId, [string]$Name)
    $session = Get-Session
    if ([string]::IsNullOrWhiteSpace($session.image)) { throw 'Image (Glance) endpoint not available' }

    $resp = (Invoke-OsApi -Method GET -BaseUrl $session.image -Path "/v2/images?name=$([System.Uri]::EscapeDataString($Name))").body
    if (-not (Test-HasProperty -Object $resp -Name 'images')) { return $null }

    foreach ($img in @($resp.images)) {
        $src = if (Test-HasProperty -Object $img -Name $script:SourceMetaKey) { [string]$img.$($script:SourceMetaKey) } else { '' }
        if ([string]::IsNullOrWhiteSpace($SourceServerId) -or $src -eq $SourceServerId) { return $img }
    }
    return $null
}

function Get-ImageIdFromLocation {
    param([object]$Headers)
    if ($null -ne $Headers -and $Headers.ContainsKey('Location')) {
        $loc = [string]@($Headers['Location'])[0]
        $m = [regex]::Match($loc, '[0-9a-fA-F-]{36}')
        if ($m.Success) { return $m.Value }
    }
    return $null
}

function Get-ControlAction {
    param([Parameter(Mandatory = $true)][string]$Control)
    switch ($Control.Trim().ToLowerInvariant()) {
        'start'   { return 'start' }
        'stop'    { return 'stop' }
        'restart' { return 'restart' }   # soft reboot
        'reset'   { return 'reset' }     # hard reboot
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

    $authUrl   = if (Test-HasProperty -Object $settings -Name 'auth_url') { [string]$settings.auth_url } else { '' }
    $username  = if (Test-HasProperty -Object $settings -Name 'username') { [string]$settings.username } else { '' }
    $password  = if (Test-HasProperty -Object $settings -Name 'password') { [string]$settings.password } else { '' }
    $project   = if (Test-HasProperty -Object $settings -Name 'project_name') { [string]$settings.project_name } else { '' }
    $userDom   = if (Test-HasProperty -Object $settings -Name 'user_domain') { [string]$settings.user_domain } else { 'Default' }
    $projDom   = if (Test-HasProperty -Object $settings -Name 'project_domain') { [string]$settings.project_domain } else { 'Default' }
    $region    = if (Test-HasProperty -Object $settings -Name 'region') { [string]$settings.region } else { '' }
    $cloneNet  = if (Test-HasProperty -Object $settings -Name 'clone_network_id') { [string]$settings.clone_network_id } else { '' }
    $skipTls   = $true
    if (Test-HasProperty -Object $settings -Name 'skip_tls') {
        try { $skipTls = [System.Convert]::ToBoolean($settings.skip_tls) } catch { $skipTls = $true }
    }

    if ([string]::IsNullOrWhiteSpace($authUrl) -or [string]::IsNullOrWhiteSpace($username) -or
        [string]::IsNullOrWhiteSpace($password) -or [string]::IsNullOrWhiteSpace($project)) {
        return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) auth_url, username, password and project_name are required"
    }

    try {
        $authBody = @{
            auth = @{
                identity = @{
                    methods  = @('password')
                    password = @{ user = @{ name = $username; domain = @{ name = $userDom }; password = $password } }
                }
                scope = @{ project = @{ name = $project; domain = @{ name = $projDom } } }
            }
        } | ConvertTo-Json -Depth 10 -Compress

        $irm = @{
            Uri                     = ($authUrl.TrimEnd('/') + '/auth/tokens')
            Method                  = 'POST'
            Body                    = $authBody
            ContentType             = 'application/json'
            ErrorAction             = 'Stop'
            ResponseHeadersVariable = 'authHeaders'
        }
        if ($PSVersionTable.PSEdition -eq 'Core' -and $skipTls) { $irm.SkipCertificateCheck = $true }

        $authResp = Invoke-RestMethod @irm
        $token = if ($null -ne $authHeaders -and $authHeaders.ContainsKey('X-Subject-Token')) { [string]@($authHeaders['X-Subject-Token'])[0] } else { '' }
        if ([string]::IsNullOrWhiteSpace($token)) {
            return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Authentication did not return a token"
        }

        $catalog = if ((Test-HasProperty -Object $authResp -Name 'token') -and (Test-HasProperty -Object $authResp.token -Name 'catalog')) { $authResp.token.catalog } else { @() }
        $compute = Get-EndpointFromCatalog -Catalog $catalog -Type 'compute' -Region $region
        $image   = Get-EndpointFromCatalog -Catalog $catalog -Type 'image' -Region $region

        if ([string]::IsNullOrWhiteSpace($compute)) {
            return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) No compute endpoint in the service catalog"
        }

        $script:Session = @{
            token            = $token
            compute          = $compute
            image            = $image
            skip_tls         = $skipTls
            clone_network_id = $cloneNet
        }

        # Validate by listing servers.
        $null = Invoke-OsApi -Method GET -BaseUrl $compute -Path '/servers?limit=1'

        Write-DebugLog "Connected; compute=$compute image=$image"
        return @{ result = @{ message = "$($script:ProviderNamePrefix) Connected to $compute" } }
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
        $session = Get-Session
        $resp = (Invoke-OsApi -Method GET -BaseUrl $session.compute -Path '/servers/detail').body
        $guests = @()
        if (Test-HasProperty -Object $resp -Name 'servers') {
            foreach ($srv in @($resp.servers)) {
                if (Test-HasProperty -Object $srv -Name 'id') { $guests += [string]$srv.id }
            }
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
            $sid = [string]$id
            try { $resultMap[$sid] = ConvertTo-RasGuestObject -Id $sid }
            catch {
                Write-DebugLog "Guest get failed for [$sid]: $($_.Exception.Message)"
                $resultMap[$sid] = @{
                    id = $sid; name = $sid; provider = 'Virtuozzo'
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

    $action = Get-ControlAction -Control ([string]$Params.control)
    if ([string]::IsNullOrWhiteSpace($action)) {
        return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Unsupported guest control: $($Params.control)"
    }

    try {
        $session = Get-Session
        $id = [string]$Params.id

        switch ($action) {
            'delete'  { Invoke-OsApi -Method DELETE -BaseUrl $session.compute -Path "/servers/$id" | Out-Null }
            'start'   { Post-ServerAction -Id $id -Action @{ 'os-start' = $null } | Out-Null }
            'stop'    { Post-ServerAction -Id $id -Action @{ 'os-stop' = $null } | Out-Null }
            'suspend' { Post-ServerAction -Id $id -Action @{ suspend = $null } | Out-Null }
            'restart' { Post-ServerAction -Id $id -Action @{ reboot = @{ type = 'SOFT' } } | Out-Null }
            'reset'   { Post-ServerAction -Id $id -Action @{ reboot = @{ type = 'HARD' } } | Out-Null }
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
        $session    = Get-Session
        $id         = [string]$Params.id
        $isTemplate = [System.Convert]::ToBoolean($Params.is_template)

        if ($isTemplate) {
            Invoke-OsApi -Method POST -BaseUrl $session.compute -Path "/servers/$id/metadata" `
                -Body @{ metadata = @{ $script:TemplateMetaKey = 'true' } } | Out-Null
            try { Post-ServerAction -Id $id -Action @{ 'os-stop' = $null } | Out-Null }
            catch { Write-DebugLog "Stop on convert ignored for [$id]: $($_.Exception.Message)" }
        }
        else {
            try { Invoke-OsApi -Method DELETE -BaseUrl $session.compute -Path "/servers/$id/metadata/$($script:TemplateMetaKey)" | Out-Null }
            catch { Write-DebugLog "Metadata delete ignored for [$id]: $($_.Exception.Message)" }
        }

        return @{ result = @{ task_id = 'sync:convert' } }
    }
    catch {
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to convert guest [$($Params.id)]: $($_.Exception.Message)"
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
        $id   = [string]$Params.id
        $name = [string]$Params.name
        # createImage snapshots the server into a Glance image tagged with the source server.
        $action = @{ createImage = @{ name = $name; metadata = @{ $script:SourceMetaKey = $id } } }
        $resp = Post-ServerAction -Id $id -Action $action
        $imageId = Get-ImageIdFromLocation -Headers $resp.headers
        if ([string]::IsNullOrWhiteSpace($imageId)) {
            # Fall back to resolving the image by name afterwards.
            return @{ result = @{ task_id = "image-name:${id}:$name" } }
        }
        return @{ result = @{ task_id = "image:$imageId" } }
    }
    catch {
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to create snapshot [$($Params.name)] for guest [$($Params.id)]: $($_.Exception.Message)"
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
        $img = Find-Image -SourceServerId ([string]$Params.id) -Name ([string]$Params.name)
        return @{ result = ($null -ne $img) }
    }
    catch {
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to check snapshot [$($Params.name)] for guest [$($Params.id)]: $($_.Exception.Message)"
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
        $session = Get-Session
        $img = Find-Image -SourceServerId ([string]$Params.id) -Name ([string]$Params.name)
        if ($null -eq $img) { return @{ result = @{ task_id = 'sync:noop' } } }
        Invoke-OsApi -Method DELETE -BaseUrl $session.image -Path "/v2/images/$([string]$img.id)" | Out-Null
        return @{ result = @{ task_id = 'sync:image-delete' } }
    }
    catch {
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to delete snapshot [$($Params.name)] for guest [$($Params.id)]: $($_.Exception.Message)"
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
        $id  = [string]$Params.id
        $img = Find-Image -SourceServerId $id -Name ([string]$Params.name)
        if ($null -eq $img) {
            return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Snapshot image [$($Params.name)] not found for server [$id]"
        }
        # OpenStack has no in-place revert; rebuilding from the image restores that state.
        Post-ServerAction -Id $id -Action @{ rebuild = @{ imageRef = [string]$img.id } } | Out-Null
        return @{ result = @{ task_id = "rebuild:$id" } }
    }
    catch {
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to revert snapshot [$($Params.name)] for guest [$($Params.id)]: $($_.Exception.Message)"
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
        $session   = Get-Session
        $sourceId  = [string]$Params.id
        $cloneName = [string]$Params.name

        $src = Get-Server -Id $sourceId
        if ($null -eq $src) { throw "Source server [$sourceId] not found" }

        $flavorRef = ''
        if ((Test-HasProperty -Object $src -Name 'flavor') -and (Test-HasProperty -Object $src.flavor -Name 'id')) {
            $flavorRef = [string]$src.flavor.id
        }
        if ([string]::IsNullOrWhiteSpace($flavorRef)) {
            return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Could not determine the source flavor for cloning"
        }

        # Resolve the image to boot from: a named snapshot version, or snapshot the source now.
        $imageId = ''
        if ((Test-HasProperty -Object $Params -Name 'snapshot') -and -not [string]::IsNullOrWhiteSpace([string]$Params.snapshot)) {
            $img = Find-Image -SourceServerId $sourceId -Name ([string]$Params.snapshot)
            if ($null -eq $img) {
                return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Snapshot image [$($Params.snapshot)] not found for server [$sourceId]"
            }
            $imageId = [string]$img.id
        }
        else {
            $snapName = "ras-clone-base-$cloneName"
            $action = @{ createImage = @{ name = $snapName; metadata = @{ $script:SourceMetaKey = $sourceId } } }
            $resp = Post-ServerAction -Id $sourceId -Action $action
            $imageId = Get-ImageIdFromLocation -Headers $resp.headers
            if ([string]::IsNullOrWhiteSpace($imageId)) {
                $img = Find-Image -SourceServerId $sourceId -Name $snapName
                if ($null -ne $img) { $imageId = [string]$img.id }
            }
            if ([string]::IsNullOrWhiteSpace($imageId)) { throw 'Could not create or resolve a base image for the clone' }
        }

        $serverSpec = @{ name = $cloneName; imageRef = $imageId; flavorRef = $flavorRef }
        if (-not [string]::IsNullOrWhiteSpace($session.clone_network_id)) {
            $serverSpec.networks = @(@{ uuid = $session.clone_network_id })
        }
        else {
            $serverSpec.networks = 'auto'
        }

        $resp = Invoke-OsApi -Method POST -BaseUrl $session.compute -Path '/servers' -Body @{ server = $serverSpec }
        $newId = if ((Test-HasProperty -Object $resp.body -Name 'server') -and (Test-HasProperty -Object $resp.body.server -Name 'id')) { [string]$resp.body.server.id } else { '' }
        if ([string]::IsNullOrWhiteSpace($newId)) { throw 'Server create did not return an id' }

        return @{ result = @{ task_id = "clone:$newId"; clone_id = $newId } }
    }
    catch {
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to clone guest [$($Params.id)]: $($_.Exception.Message)"
    }
}

function Handle-TaskInfo {
    param([object]$Params)

    if ($null -eq $Params -or [string]::IsNullOrWhiteSpace([string]$Params.id)) {
        return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Invalid task id"
    }

    try {
        $session = Get-Session
        $taskId  = [string]$Params.id
        $parts   = $taskId -split ':', 3
        $kind    = $parts[0]

        switch ($kind) {
            'sync' { return @{ result = @{ state = 'completed'; output = @{} } } }

            'image' {
                $imageId = if ($parts.Count -ge 2) { $parts[1] } else { '' }
                $img = Get-OsResourceOrNull -BaseUrl $session.image -Path "/v2/images/$imageId"
                if ($null -eq $img) { return @{ result = @{ state = 'running' } } }
                $status = if (Test-HasProperty -Object $img -Name 'status') { ([string]$img.status).ToLowerInvariant() } else { '' }
                switch ($status) {
                    'active' { return @{ result = @{ state = 'completed'; output = @{} } } }
                    'killed' { return @{ result = @{ state = 'failed'; error = @{ code = 1; message = 'Image creation failed' } } } }
                    default  { return @{ result = @{ state = 'running' } } }
                }
            }

            'image-name' {
                $srvId = if ($parts.Count -ge 2) { $parts[1] } else { '' }
                $name  = if ($parts.Count -ge 3) { $parts[2] } else { '' }
                $img = Find-Image -SourceServerId $srvId -Name $name
                if ($null -eq $img) { return @{ result = @{ state = 'running' } } }
                $status = if (Test-HasProperty -Object $img -Name 'status') { ([string]$img.status).ToLowerInvariant() } else { '' }
                if ($status -eq 'active') { return @{ result = @{ state = 'completed'; output = @{} } } }
                return @{ result = @{ state = 'running' } }
            }

            'clone' {
                $srvId = if ($parts.Count -ge 2) { $parts[1] } else { '' }
                $srv = Get-Server -Id $srvId
                if ($null -eq $srv) { return @{ result = @{ state = 'running' } } }
                $status = if (Test-HasProperty -Object $srv -Name 'status') { [string]$srv.status.ToUpperInvariant() } else { '' }
                if ($status -eq 'ERROR') { return @{ result = @{ state = 'failed'; error = @{ code = 1; message = 'Clone failed' } } } }
                if ($status -eq 'BUILD') { return @{ result = @{ state = 'running' } } }
                return @{ result = @{ state = 'completed'; output = @{ clone_id = $srvId } } }
            }

            'rebuild' {
                $srvId = if ($parts.Count -ge 2) { $parts[1] } else { '' }
                $srv = Get-Server -Id $srvId
                if ($null -eq $srv) { return @{ result = @{ state = 'running' } } }
                $status = if (Test-HasProperty -Object $srv -Name 'status') { [string]$srv.status.ToUpperInvariant() } else { '' }
                if ($status -eq 'ACTIVE' -or $status -eq 'SHUTOFF') { return @{ result = @{ state = 'completed'; output = @{} } } }
                return @{ result = @{ state = 'running' } }
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

Write-DebugLog "Virtuozzo provider process started. PID=$PID"

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
