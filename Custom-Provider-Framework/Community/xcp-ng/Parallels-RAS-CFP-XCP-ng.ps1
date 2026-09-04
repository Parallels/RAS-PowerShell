<#
.SYNOPSIS
    Parallels RAS Custom Provider sample script for XCP-ng.
.DESCRIPTION
    Implements a Parallels RAS Custom Provider that integrates XCP-ng through the
    XenAPI (XAPI) management interface. XAPI exposes a JSON-RPC 2.0 endpoint at
    /jsonrpc on the pool master, which this provider uses (no XML-RPC, no CLI on
    the RAS host).

    It listens for JSON-RPC requests on standard input, processes them per the
    Custom Provider Framework (CPF) protocol, and writes responses on standard
    output.

    Supported workflow:
      - connect to the pool master with username/password (XAPI session)
      - enumerate VMs (excluding snapshots, templates and the control domain)
      - retrieve guest info (power state, IP/MAC addresses, guest OS, template flag)
      - power operations: start, stop, restart, reset, suspend, delete
      - convert a VM to/from a template (native is_a_template flag)
      - template versioning using native VM snapshots (VM.snapshot / VM.revert)
      - clone a VM, or a named snapshot version, into a new VM
      - tasks/get (operations are performed synchronously; tasks resolve immediately)

    State values returned to RAS: powered_on, powered_off, powering_on,
    powering_off, suspended, suspending.

    Requires PowerShell 7 or later.
.NOTES
    File Name : Parallels-RAS-CFP-XCP-ng.ps1
    Platform  : XCP-ng (XenAPI / XAPI JSON-RPC)
    Author    : Edwin Houben
    Reference : xcp-ng/XCP-ng-API.md in this folder
.EXAMPLE
    {"method":"provider/initialize"}
    {"method":"provider/connect","params":{"settings":{"host":"https://xcp-pool.example.com","username":"root","password":"secret"}}}
    {"method":"guests/list"}
    {"method":"guests/control","params":{"id":"<vm-uuid>","control":"start"}}
    {"method":"guests/snapshots/create","params":{"id":"<vm-uuid>","name":"RAS_TEMPLATE_VERSION_1"}}
    {"method":"guests/clone","params":{"id":"<vm-uuid>","name":"vdi-clone-01","snapshot":"RAS_TEMPLATE_VERSION_1"}}
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

$script:ProviderNamePrefix = 'XCP-ng:'
$script:LogPath            = Join-Path ([System.IO.Path]::GetTempPath()) 'XCP-ng-RAS-Provider.log'
$script:Session            = $null

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
# XAPI JSON-RPC transport
# ----------------------------------------------------------------------------

function ConvertTo-BaseUrl {
    param([string]$HostValue)
    $h = $HostValue.Trim()
    if ($h -notmatch '^https?://') { $h = "https://$h" }
    return $h.TrimEnd('/')
}

function Invoke-XapiRpc {
    <#
        Low-level XAPI JSON-RPC 2.0 call. $Params is the positional argument array.
        Returns the JSON-RPC result, or throws with the XAPI error message.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$Method,
        [object[]]$Params = @(),
        [string]$Base,
        [bool]$SkipTls = $true
    )

    $baseUrl = if ($PSBoundParameters.ContainsKey('Base')) { $Base } else { (Get-Session).base }
    $skip    = if ($PSBoundParameters.ContainsKey('SkipTls')) { $SkipTls } else { (Get-Session).skip_tls }

    $payload = @{ jsonrpc = '2.0'; method = $Method; params = @($Params); id = 1 } | ConvertTo-Json -Compress -Depth 32

    $irm = @{
        Uri         = "$baseUrl/jsonrpc"
        Method      = 'POST'
        Body        = $payload
        ContentType = 'application/json'
        ErrorAction = 'Stop'
    }
    if ($PSVersionTable.PSEdition -eq 'Core' -and $skip) { $irm.SkipCertificateCheck = $true }

    Write-DebugLog "RPC $Method params=$($Params.Count)"
    $resp = Invoke-RestMethod @irm

    if (Test-HasProperty -Object $resp -Name 'error' -and $null -ne $resp.error) {
        $msg = if (Test-HasProperty -Object $resp.error -Name 'message') { [string]$resp.error.message } else { ($resp.error | ConvertTo-Json -Compress) }
        throw "XAPI error on $Method`: $msg"
    }
    if (Test-HasProperty -Object $resp -Name 'result') { return $resp.result }
    return $null
}

function Get-Session {
    if ($null -eq $script:Session) { throw 'Session not initialized' }
    if ([string]::IsNullOrWhiteSpace($script:Session.base)) { throw 'Session base URL missing' }
    if ([string]::IsNullOrWhiteSpace($script:Session.ref)) { throw 'Session reference missing' }
    return $script:Session
}

function Invoke-Xapi {
    <# Session-scoped XAPI call: prepends the session reference to $Args. #>
    param(
        [Parameter(Mandatory = $true)][string]$Method,
        [object[]]$Args = @()
    )
    $session = Get-Session
    $params  = @($session.ref) + @($Args)
    return Invoke-XapiRpc -Method $Method -Params $params
}

# ----------------------------------------------------------------------------
# VM helpers
# ----------------------------------------------------------------------------

function Get-VmRefByUuid {
    param([Parameter(Mandatory = $true)][string]$Uuid)
    return [string](Invoke-Xapi -Method 'VM.get_by_uuid' -Args @($Uuid))
}

function Map-PowerStateToRasState {
    param([string]$PowerState)
    switch (($PowerState | ForEach-Object { $_.ToString().Trim().ToLowerInvariant() })) {
        'running'   { return 'powered_on' }
        'halted'    { return 'powered_off' }
        'suspended' { return 'suspended' }
        'paused'    { return 'suspended' }
        default     { return 'powered_off' }
    }
}

function Get-VmNetworkData {
    param([string]$VmRef)

    $ips  = New-Object 'System.Collections.Generic.List[string]'
    $macs = New-Object 'System.Collections.Generic.List[string]'

    # IPs from the guest agent (requires Xen tools / management agent in the VM).
    try {
        $gm = [string](Invoke-Xapi -Method 'VM.get_guest_metrics' -Args @($VmRef))
        if (-not [string]::IsNullOrWhiteSpace($gm) -and $gm -ne 'OpaqueRef:NULL') {
            $networks = Invoke-Xapi -Method 'VM_guest_metrics.get_networks' -Args @($gm)
            if ($null -ne $networks) {
                foreach ($prop in $networks.PSObject.Properties) {
                    # keys look like "0/ip", "0/ipv6/0"
                    if ($prop.Name -match '/ip$' ) {
                        $ip = [string]$prop.Value
                        if (-not [string]::IsNullOrWhiteSpace($ip) -and $ip -notmatch ':' -and -not $ips.Contains($ip)) {
                            $ips.Add($ip)
                        }
                    }
                }
            }
        }
    }
    catch { Write-DebugLog "Guest metrics lookup failed: $($_.Exception.Message)" }

    # MACs from the VIFs.
    try {
        $vifs = @(Invoke-Xapi -Method 'VM.get_VIFs' -Args @($VmRef))
        foreach ($vif in $vifs) {
            $vifRef = [string]$vif
            if ([string]::IsNullOrWhiteSpace($vifRef) -or $vifRef -eq 'OpaqueRef:NULL') { continue }
            $mac = [string](Invoke-Xapi -Method 'VIF.get_MAC' -Args @($vifRef))
            if (-not [string]::IsNullOrWhiteSpace($mac) -and -not $macs.Contains($mac.ToUpperInvariant())) {
                $macs.Add($mac.ToUpperInvariant())
            }
        }
    }
    catch { Write-DebugLog "VIF lookup failed: $($_.Exception.Message)" }

    return @{
        IPv4Addresses = @($ips  | Select-Object -First 3)
        MacAddresses  = @($macs | Select-Object -First 3)
    }
}

function Get-VmOsType {
    param([string]$VmRef)
    try {
        $gm = [string](Invoke-Xapi -Method 'VM.get_guest_metrics' -Args @($VmRef))
        if (-not [string]::IsNullOrWhiteSpace($gm) -and $gm -ne 'OpaqueRef:NULL') {
            $os = Invoke-Xapi -Method 'VM_guest_metrics.get_os_version' -Args @($gm)
            if ($null -ne $os -and (Test-HasProperty -Object $os -Name 'name')) {
                return [string]$os.name
            }
        }
    }
    catch { }
    return 'unknown'
}

function ConvertTo-RasGuestObject {
    param([Parameter(Mandatory = $true)][string]$Uuid)

    $ref = Get-VmRefByUuid -Uuid $Uuid
    $rec = Invoke-Xapi -Method 'VM.get_record' -Args @($ref)
    if ($null -eq $rec) { throw "VM [$Uuid] not found" }

    $power = if (Test-HasProperty -Object $rec -Name 'power_state') { [string]$rec.power_state } else { '' }
    $name  = if (Test-HasProperty -Object $rec -Name 'name_label') { [string]$rec.name_label } else { "vm-$Uuid" }
    $isTpl = (Test-HasProperty -Object $rec -Name 'is_a_template') -and [bool]$rec.is_a_template

    $net = @{ IPv4Addresses = @(); MacAddresses = @() }
    if ((Map-PowerStateToRasState -PowerState $power) -eq 'powered_on') {
        $net = Get-VmNetworkData -VmRef $ref
    }

    $guest = @{
        id            = [string]$Uuid
        name          = $name
        provider      = 'XCP-ng'
        state         = (Map-PowerStateToRasState -PowerState $power)
        power_state   = $(if ([string]::IsNullOrWhiteSpace($power)) { 'unknown' } else { $power })
        host_os       = (Get-VmOsType -VmRef $ref)
        ip            = $(if ($net.IPv4Addresses.Count -gt 0) { $net.IPv4Addresses[0] } else { $null })
        ip_addresses  = @($net.IPv4Addresses)
        mac_addresses = @($net.MacAddresses)
        is_template   = $isTpl
        type          = 'Virtual Machine'
    }

    Write-DebugLog ("GUEST uuid={0}; name={1}; state={2}; power={3}; template={4}; ips={5}" -f `
            $guest.id, $guest.name, $guest.state, $guest.power_state, $guest.is_template, ($guest.ip_addresses -join ','))
    return $guest
}

function Find-SnapshotRefByName {
    param([string]$VmRef, [string]$SnapshotName)
    $snaps = @(Invoke-Xapi -Method 'VM.get_snapshots' -Args @($VmRef))
    foreach ($snap in $snaps) {
        $snapRef = [string]$snap
        if ([string]::IsNullOrWhiteSpace($snapRef) -or $snapRef -eq 'OpaqueRef:NULL') { continue }
        $label = [string](Invoke-Xapi -Method 'VM.get_name_label' -Args @($snapRef))
        if ($label -eq $SnapshotName) { return $snapRef }
    }
    return $null
}

function Get-ControlAction {
    param([Parameter(Mandatory = $true)][string]$Control)
    switch ($Control.Trim().ToLowerInvariant()) {
        'start'   { return 'start' }
        'stop'    { return 'stop' }
        'restart' { return 'restart' }   # graceful (VM.clean_reboot)
        'reset'   { return 'reset' }     # hard (VM.hard_reboot)
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

    $hostValue = if (Test-HasProperty -Object $settings -Name 'host') { [string]$settings.host } else { '' }
    $username  = if (Test-HasProperty -Object $settings -Name 'username') { [string]$settings.username } else { '' }
    $password  = if (Test-HasProperty -Object $settings -Name 'password') { [string]$settings.password } else { '' }
    $skipTls   = $true
    if (Test-HasProperty -Object $settings -Name 'skip_tls') {
        try { $skipTls = [System.Convert]::ToBoolean($settings.skip_tls) } catch { $skipTls = $true }
    }

    if ([string]::IsNullOrWhiteSpace($hostValue) -or
        [string]::IsNullOrWhiteSpace($username) -or
        [string]::IsNullOrWhiteSpace($password)) {
        return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) host, username and password are required"
    }

    try {
        $base = ConvertTo-BaseUrl -HostValue $hostValue
        $ref  = [string](Invoke-XapiRpc -Method 'session.login_with_password' -Params @($username, $password) -Base $base -SkipTls $skipTls)

        if ([string]::IsNullOrWhiteSpace($ref) -or $ref -eq 'OpaqueRef:NULL') {
            return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Login did not return a session reference"
        }

        $script:Session = @{ base = $base; skip_tls = $skipTls; ref = $ref }
        Write-DebugLog "Connected to $base"
        return @{ result = @{ message = "$($script:ProviderNamePrefix) Connected to $base" } }
    }
    catch {
        $script:Session = $null
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to connect: $($_.Exception.Message)"
    }
}

function Handle-Disconnect {
    try {
        if ($null -ne $script:Session) {
            try { Invoke-Xapi -Method 'session.logout' -Args @() | Out-Null } catch { }
        }
    }
    finally { $script:Session = $null }
    return @{ result = @{} }
}

function Handle-GuestList {
    try {
        $records = Invoke-Xapi -Method 'VM.get_all_records' -Args @()
        $guests = @()
        if ($null -ne $records) {
            foreach ($prop in $records.PSObject.Properties) {
                $rec = $prop.Value
                $isTpl  = (Test-HasProperty -Object $rec -Name 'is_a_template') -and [bool]$rec.is_a_template
                $isSnap = (Test-HasProperty -Object $rec -Name 'is_a_snapshot') -and [bool]$rec.is_a_snapshot
                $isDom0 = (Test-HasProperty -Object $rec -Name 'is_control_domain') -and [bool]$rec.is_control_domain
                if ($isTpl -or $isSnap -or $isDom0) { continue }
                if (Test-HasProperty -Object $rec -Name 'uuid' -and -not [string]::IsNullOrWhiteSpace([string]$rec.uuid)) {
                    $guests += [string]$rec.uuid
                }
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
            return @{ result = (ConvertTo-RasGuestObject -Uuid ([string]$ids[0])) }
        }

        $resultMap = @{}
        foreach ($id in $ids) {
            $uuid = [string]$id
            try { $resultMap[$uuid] = ConvertTo-RasGuestObject -Uuid $uuid }
            catch {
                Write-DebugLog "Guest get failed for [$uuid]: $($_.Exception.Message)"
                $resultMap[$uuid] = @{
                    id = $uuid; name = $uuid; provider = 'XCP-ng'
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
        $ref = Get-VmRefByUuid -Uuid ([string]$Params.id)

        switch ($action) {
            'start'   { Invoke-Xapi -Method 'VM.start' -Args @($ref, $false, $false) | Out-Null }
            'suspend' { Invoke-Xapi -Method 'VM.suspend' -Args @($ref) | Out-Null }
            'reset'   { Invoke-Xapi -Method 'VM.hard_reboot' -Args @($ref) | Out-Null }
            'restart' { Invoke-Xapi -Method 'VM.clean_reboot' -Args @($ref) | Out-Null }
            'delete'  { Invoke-Xapi -Method 'VM.destroy' -Args @($ref) | Out-Null }
            'stop' {
                try { Invoke-Xapi -Method 'VM.clean_shutdown' -Args @($ref) | Out-Null }
                catch {
                    Write-DebugLog "clean_shutdown failed, trying hard_shutdown: $($_.Exception.Message)"
                    Invoke-Xapi -Method 'VM.hard_shutdown' -Args @($ref) | Out-Null
                }
            }
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
        $ref        = Get-VmRefByUuid -Uuid ([string]$Params.id)
        $isTemplate = [System.Convert]::ToBoolean($Params.is_template)
        Invoke-Xapi -Method 'VM.set_is_a_template' -Args @($ref, $isTemplate) | Out-Null
        return @{ result = @{ task_id = 'sync:convert' } }
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
        $sourceRef = Get-VmRefByUuid -Uuid ([string]$Params.id)
        $cloneName = [string]$Params.name

        # Clone from a named snapshot version when requested, otherwise from the VM.
        $cloneSource = $sourceRef
        if ((Test-HasProperty -Object $Params -Name 'snapshot') -and -not [string]::IsNullOrWhiteSpace([string]$Params.snapshot)) {
            $snapRef = Find-SnapshotRefByName -VmRef $sourceRef -SnapshotName ([string]$Params.snapshot)
            if ($null -eq $snapRef) {
                return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Snapshot [$($Params.snapshot)] not found for VM [$($Params.id)]"
            }
            $cloneSource = $snapRef
        }

        $newRef = [string](Invoke-Xapi -Method 'VM.clone' -Args @($cloneSource, $cloneName))

        # A clone of a snapshot/template may inherit the template flag; make it a usable VM.
        $newIsTpl = [bool](Invoke-Xapi -Method 'VM.get_is_a_template' -Args @($newRef))
        if ($newIsTpl) { Invoke-Xapi -Method 'VM.set_is_a_template' -Args @($newRef, $false) | Out-Null }

        $newUuid = [string](Invoke-Xapi -Method 'VM.get_uuid' -Args @($newRef))
        return @{ result = @{ task_id = "clone:$newUuid"; clone_id = $newUuid } }
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
        $ref = Get-VmRefByUuid -Uuid ([string]$Params.id)
        Invoke-Xapi -Method 'VM.snapshot' -Args @($ref, [string]$Params.name) | Out-Null
        return @{ result = @{ task_id = 'sync:snapshot' } }
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
        $ref     = Get-VmRefByUuid -Uuid ([string]$Params.id)
        $snapRef = Find-SnapshotRefByName -VmRef $ref -SnapshotName ([string]$Params.name)
        if ($null -eq $snapRef) { return @{ result = @{ task_id = 'sync:noop' } } }
        Invoke-Xapi -Method 'VM.destroy' -Args @($snapRef) | Out-Null
        return @{ result = @{ task_id = 'sync:snapshot-delete' } }
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
        $ref     = Get-VmRefByUuid -Uuid ([string]$Params.id)
        $snapRef = Find-SnapshotRefByName -VmRef $ref -SnapshotName ([string]$Params.name)
        return @{ result = ($null -ne $snapRef) }
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
        $ref     = Get-VmRefByUuid -Uuid ([string]$Params.id)
        $snapRef = Find-SnapshotRefByName -VmRef $ref -SnapshotName ([string]$Params.name)
        if ($null -eq $snapRef) {
            return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Snapshot [$($Params.name)] not found for VM [$($Params.id)]"
        }
        Invoke-Xapi -Method 'VM.revert' -Args @($snapRef) | Out-Null
        return @{ result = @{ task_id = 'sync:revert' } }
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

    # All operations are performed synchronously, so tasks resolve immediately.
    $taskId = [string]$Params.id
    if ($taskId -like 'clone:*') {
        $uuid = $taskId.Substring('clone:'.Length)
        return @{ result = @{ state = 'completed'; output = @{ clone_id = $uuid } } }
    }
    return @{ result = @{ state = 'completed'; output = @{} } }
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

Write-DebugLog "XCP-ng provider process started. PID=$PID"

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
