<#
.SYNOPSIS
    Parallels RAS Custom Provider sample script for Red Hat OpenShift Virtualization (KubeVirt).
.DESCRIPTION
    Implements a Parallels RAS Custom Provider that integrates with OpenShift Virtualization
    (KubeVirt) through the Kubernetes/OpenShift REST API. It listens for JSON-RPC requests on
    standard input, processes them according to the Custom Provider Framework (CPF) protocol,
    and returns responses on standard output.

    Supported workflow:
      - connect using an OpenShift API server URL and a bearer (ServiceAccount) token
      - enumerate VirtualMachines in a namespace
      - retrieve guest info (power state, IP and MAC addresses, guest OS, template flag)
      - power operations: start, stop, restart, reset (mapped to restart), suspend (pause), delete
      - convert a VM to/from a RAS template (tracked with a label)
      - template versioning using native VirtualMachineSnapshot / VirtualMachineRestore objects
      - clone a VM (or a snapshot) into a new VM with VirtualMachineClone
      - asynchronous task status through tasks/get

    The protocol contract follows the Parallels RAS Custom Provider Framework "Solution Model"
    and "Capabilities" documentation. State values returned to RAS are: powered_on, powered_off,
    powering_on, powering_off, suspended, suspending.

    Requires PowerShell 7 or later.
.NOTES
    File Name : Parallels-RAS-CFP-OpenShift.ps1
    Platform  : Red Hat OpenShift Virtualization / KubeVirt
    Author    : Edwin Houben
    Reference : OpenShift/OpenShift-Virtualization-API.md in this folder
.EXAMPLE
    Sample requests (one JSON object per line on stdin):

    {"method":"provider/initialize"}
    {"method":"provider/connect","params":{"settings":{"host":"https://api.ocp.example.com:6443","token":"<sa-token>","namespace":"vdi"}}}
    {"method":"guests/list"}
    {"method":"guests/get","params":{"id":"win11-vdi-01"}}
    {"method":"guests/control","params":{"id":"win11-vdi-01","control":"start"}}
    {"method":"guests/convert","params":{"id":"win11-gold","is_template":true}}
    {"method":"guests/snapshots/create","params":{"id":"win11-gold","name":"RAS_TEMPLATE_VERSION_1"}}
    {"method":"guests/clone","params":{"id":"win11-gold","name":"win11-vdi-02","snapshot":"RAS_TEMPLATE_VERSION_1"}}
#>

Set-StrictMode -Version Latest

$ErrorActionPreference   = 'Stop'
$ProgressPreference      = 'SilentlyContinue'
$WarningPreference       = 'SilentlyContinue'
$VerbosePreference       = 'SilentlyContinue'
$InformationPreference   = 'SilentlyContinue'

if ($Host.Name -notmatch 'ISE') {
    [Console]::InputEncoding  = [System.Text.Encoding]::UTF8
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
}

$stdout = [Console]::OpenStandardOutput()
$writer = New-Object System.IO.StreamWriter($stdout, [System.Text.Encoding]::UTF8)
$writer.AutoFlush = $true

$script:ProviderNamePrefix = 'OpenShift:'
$script:LogPath            = Join-Path ([System.IO.Path]::GetTempPath()) 'OpenShift-RAS-Provider.log'
$script:Session            = $null

# Label used to mark a VM as a RAS template, and annotation used to map a RAS snapshot
# name (which may contain spaces/underscores) to a DNS-safe Kubernetes object name.
$script:TemplateLabel      = 'ras.parallels.com/template'
$script:SnapshotNameAnno   = 'ras.parallels.com/snapshot-name'
$script:SourceVmAnno       = 'ras.parallels.com/source-vm'

# KubeVirt / OpenShift Virtualization API group paths.
$script:KubevirtApi        = 'apis/kubevirt.io/v1'
$script:SubresApi          = 'apis/subresources.kubevirt.io/v1'
$script:SnapshotApi        = 'apis/snapshot.kubevirt.io/v1beta1'
$script:CloneApi           = 'apis/clone.kubevirt.io/v1alpha1'

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
    catch {
        # Never emit logging failures to stdout.
    }
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
            error = @{
                code    = $script:ErrorCodes.InternalError
                message = "$($script:ProviderNamePrefix) Failed to serialize response: $($_.Exception.Message)"
            }
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
    catch {
        Write-DebugLog "JSON parse failed: $($_.Exception.Message)"
        return $null
    }
}

function Test-RequiredFields {
    param([object]$Data, [string[]]$RequiredFields)
    foreach ($field in $RequiredFields) {
        $keys  = $field -split '\.'
        $value = $Data
        foreach ($key in $keys) {
            if ($null -ne $value -and $value.PSObject.Properties.Name -contains $key) {
                $value = $value.$key
            }
            else {
                return "$($script:ProviderNamePrefix) Missing field: $field"
            }
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
    param([string]$HostValue)
    $h = $HostValue.Trim()
    if ($h -notmatch '^https?://') { $h = "https://$h" }
    return $h.TrimEnd('/')
}

function Invoke-OcpApi {
    <#
        Calls the OpenShift/Kubernetes REST API with the session bearer token.
        $Path is relative to the API server base URL (no leading slash required).
        Throws on HTTP error. Use Get-OcpResourceOrNull for 404-tolerant reads.
    #>
    param(
        [Parameter(Mandatory = $true)][ValidateSet('GET', 'POST', 'PUT', 'PATCH', 'DELETE')][string]$Method,
        [Parameter(Mandatory = $true)][string]$Path,
        [object]$Body = $null,
        [string]$ContentType = 'application/json'
    )

    $session = Get-Session
    $uri = ($session.base + '/' + $Path.TrimStart('/'))

    $params = @{
        Uri         = $uri
        Headers     = $session.header
        Method      = $Method
        ErrorAction = 'Stop'
        ContentType = $ContentType
    }

    # PowerShell 7+ honours these; they let the provider work with the
    # self-signed certificates that OpenShift API servers commonly present.
    if ($PSVersionTable.PSEdition -eq 'Core') {
        if ($session.skip_tls) { $params.SkipCertificateCheck = $true }
        $params.SkipHeaderValidation = $true
    }

    if ($null -ne $Body) {
        if ($Body -is [string]) { $params.Body = $Body }
        else { $params.Body = $Body | ConvertTo-Json -Compress -Depth 32 }
    }

    Write-DebugLog "HTTP $Method $uri$(if ($null -ne $Body) { ' - Body: ' + $params.Body })"

    try {
        return Invoke-RestMethod @params
    }
    catch {
        Write-DebugLog "HTTP failure ($Method $Path): $($_.Exception.Message)"
        throw
    }
}

function Get-HttpStatusCode {
    param([System.Management.Automation.ErrorRecord]$ErrorRecord)
    try {
        $resp = $ErrorRecord.Exception.Response
        if ($null -ne $resp -and $null -ne $resp.StatusCode) {
            return [int]$resp.StatusCode
        }
    }
    catch { }
    return 0
}

function Get-OcpResourceOrNull {
    param([Parameter(Mandatory = $true)][string]$Path)
    try {
        return Invoke-OcpApi -Method GET -Path $Path
    }
    catch {
        if ((Get-HttpStatusCode -ErrorRecord $_) -eq 404) {
            return $null
        }
        throw
    }
}

function ConvertTo-K8sName {
    # Produces a DNS-1123 compatible name fragment from an arbitrary RAS string.
    param([Parameter(Mandatory = $true)][string]$Value)
    $n = $Value.ToLowerInvariant() -replace '[^a-z0-9]+', '-'
    $n = $n.Trim('-')
    if ($n.Length -gt 48) { $n = $n.Substring(0, 48).Trim('-') }
    if ([string]::IsNullOrWhiteSpace($n)) { $n = 'snap' }
    return $n
}

function New-UniqueSuffix {
    return ([guid]::NewGuid().ToString('N')).Substring(0, 8)
}

# ----------------------------------------------------------------------------
# VirtualMachine read helpers
# ----------------------------------------------------------------------------

function Get-VmList {
    $ns = (Get-Session).namespace
    $resp = Invoke-OcpApi -Method GET -Path "$($script:KubevirtApi)/namespaces/$ns/virtualmachines"
    if (Test-HasProperty -Object $resp -Name 'items') { return @($resp.items) }
    return @()
}

function Get-Vm {
    param([Parameter(Mandatory = $true)][string]$Name)
    $ns = (Get-Session).namespace
    return Get-OcpResourceOrNull -Path "$($script:KubevirtApi)/namespaces/$ns/virtualmachines/$Name"
}

function Get-Vmi {
    param([Parameter(Mandatory = $true)][string]$Name)
    $ns = (Get-Session).namespace
    return Get-OcpResourceOrNull -Path "$($script:KubevirtApi)/namespaces/$ns/virtualmachineinstances/$Name"
}

function Map-PrintableStatusToRasState {
    param([string]$Status)
    $normalized = if ($null -ne $Status) { $Status.ToString().Trim().ToLowerInvariant() } else { '' }
    switch ($normalized) {
        'running'                 { return 'powered_on' }
        'migrating'               { return 'powered_on' }
        'stopped'                 { return 'powered_off' }
        'starting'                { return 'powering_on' }
        'provisioning'            { return 'powering_on' }
        'waitingforvolumebinding' { return 'powering_on' }
        'stopping'                { return 'powering_off' }
        'terminating'             { return 'powering_off' }
        'paused'                  { return 'suspended' }
        default                   { return 'powered_off' }
    }
}

function Get-VmNetworkData {
    param([object]$Vmi)
    $ipv4 = New-Object 'System.Collections.Generic.List[string]'
    $macs = New-Object 'System.Collections.Generic.List[string]'

    if ($null -eq $Vmi -or -not (Test-HasProperty -Object $Vmi -Name 'status')) {
        return @{ IPv4Addresses = @(); MacAddresses = @() }
    }
    if (-not (Test-HasProperty -Object $Vmi.status -Name 'interfaces')) {
        return @{ IPv4Addresses = @(); MacAddresses = @() }
    }

    foreach ($iface in @($Vmi.status.interfaces)) {
        if (Test-HasProperty -Object $iface -Name 'mac' -and -not [string]::IsNullOrWhiteSpace([string]$iface.mac)) {
            $mac = ([string]$iface.mac).ToUpperInvariant()
            if (-not $macs.Contains($mac)) { $macs.Add($mac) }
        }

        $addresses = @()
        if (Test-HasProperty -Object $iface -Name 'ipAddresses') { $addresses = @($iface.ipAddresses) }
        elseif (Test-HasProperty -Object $iface -Name 'ipAddress') { $addresses = @($iface.ipAddress) }

        foreach ($addr in $addresses) {
            $ip = [string]$addr
            if (-not [string]::IsNullOrWhiteSpace($ip) -and
                $ip -notmatch ':' -and                 # skip IPv6
                $ip -ne '127.0.0.1' -and
                $ip -notmatch '^169\.254\.') {
                if (-not $ipv4.Contains($ip)) { $ipv4.Add($ip) }
            }
        }
    }

    return @{
        IPv4Addresses = @($ipv4 | Select-Object -First 3)
        MacAddresses  = @($macs | Select-Object -First 3)
    }
}

function Get-VmGuestOs {
    param([object]$Vm, [object]$Vmi)
    if ($null -ne $Vmi -and (Test-HasProperty -Object $Vmi -Name 'status') -and
        (Test-HasProperty -Object $Vmi.status -Name 'guestOSInfo') -and
        (Test-HasProperty -Object $Vmi.status.guestOSInfo -Name 'name') -and
        -not [string]::IsNullOrWhiteSpace([string]$Vmi.status.guestOSInfo.name)) {
        return [string]$Vmi.status.guestOSInfo.name
    }
    return 'unknown'
}

function Get-VmIsTemplate {
    param([object]$Vm)
    if ($null -ne $Vm -and (Test-HasProperty -Object $Vm -Name 'metadata') -and
        (Test-HasProperty -Object $Vm.metadata -Name 'labels') -and
        (Test-HasProperty -Object $Vm.metadata.labels -Name $script:TemplateLabel)) {
        return ([string]$Vm.metadata.labels.$($script:TemplateLabel) -eq 'true')
    }
    return $false
}

function ConvertTo-RasGuestObject {
    param([Parameter(Mandatory = $true)][string]$VmName)

    $vm = Get-Vm -Name $VmName
    if ($null -eq $vm) { throw "VM [$VmName] not found in namespace [$((Get-Session).namespace)]" }

    $printable = ''
    if ((Test-HasProperty -Object $vm -Name 'status') -and (Test-HasProperty -Object $vm.status -Name 'printableStatus')) {
        $printable = [string]$vm.status.printableStatus
    }

    $vmi = $null
    if ((Map-PrintableStatusToRasState -Status $printable) -eq 'powered_on') {
        try { $vmi = Get-Vmi -Name $VmName }
        catch { Write-DebugLog "VMI lookup failed for [$VmName]: $($_.Exception.Message)" }
    }

    $network = Get-VmNetworkData -Vmi $vmi

    $guest = @{
        id            = [string]$VmName
        name          = [string]$VmName
        provider      = 'OpenShift'
        namespace     = (Get-Session).namespace
        state         = (Map-PrintableStatusToRasState -Status $printable)
        power_state   = $(if ([string]::IsNullOrWhiteSpace($printable)) { 'unknown' } else { $printable })
        host_os       = (Get-VmGuestOs -Vm $vm -Vmi $vmi)
        ip            = $(if ($network.IPv4Addresses.Count -gt 0) { $network.IPv4Addresses[0] } else { $null })
        ip_addresses  = @($network.IPv4Addresses)
        mac_addresses = @($network.MacAddresses)
        is_template   = (Get-VmIsTemplate -Vm $vm)
        type          = 'Virtual Machine'
    }

    Write-DebugLog ("GUEST name={0}; state={1}; printable={2}; template={3}; ips={4}" -f `
            $guest.name, $guest.state, $guest.power_state, $guest.is_template, ($guest.ip_addresses -join ','))

    return $guest
}

# ----------------------------------------------------------------------------
# Snapshot helpers (template versioning)
# ----------------------------------------------------------------------------

function Get-SnapshotObjects {
    $ns = (Get-Session).namespace
    $resp = Invoke-OcpApi -Method GET -Path "$($script:SnapshotApi)/namespaces/$ns/virtualmachinesnapshots"
    if (Test-HasProperty -Object $resp -Name 'items') { return @($resp.items) }
    return @()
}

function Find-SnapshotObject {
    # Matches a snapshot for $VmName whose RAS logical name (stored as an annotation)
    # equals $SnapshotName.
    param([string]$VmName, [string]$SnapshotName)

    foreach ($snap in Get-SnapshotObjects) {
        $sourceName = $null
        if ((Test-HasProperty -Object $snap -Name 'spec') -and
            (Test-HasProperty -Object $snap.spec -Name 'source') -and
            (Test-HasProperty -Object $snap.spec.source -Name 'name')) {
            $sourceName = [string]$snap.spec.source.name
        }
        if ($sourceName -ne $VmName) { continue }

        $logical = $null
        if ((Test-HasProperty -Object $snap -Name 'metadata') -and
            (Test-HasProperty -Object $snap.metadata -Name 'annotations') -and
            (Test-HasProperty -Object $snap.metadata.annotations -Name $script:SnapshotNameAnno)) {
            $logical = [string]$snap.metadata.annotations.$($script:SnapshotNameAnno)
        }

        if ($logical -eq $SnapshotName) { return $snap }
    }
    return $null
}

# ----------------------------------------------------------------------------
# Control action mapping
# ----------------------------------------------------------------------------

function Get-ControlAction {
    param([Parameter(Mandatory = $true)][string]$Control)
    switch ($Control.Trim().ToLowerInvariant()) {
        'start'   { return 'start' }
        'stop'    { return 'stop' }
        'restart' { return 'restart' }
        'reset'   { return 'restart' }   # KubeVirt has no hard reset; restart is the closest action.
        'suspend' { return 'suspend' }   # mapped to VMI pause
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
    $token     = if (Test-HasProperty -Object $settings -Name 'token') { [string]$settings.token } else { '' }
    $namespace = if (Test-HasProperty -Object $settings -Name 'namespace') { [string]$settings.namespace } else { '' }
    $skipTls   = $true
    if (Test-HasProperty -Object $settings -Name 'skip_tls') {
        try { $skipTls = [System.Convert]::ToBoolean($settings.skip_tls) } catch { $skipTls = $true }
    }

    if ([string]::IsNullOrWhiteSpace($hostValue) -or
        [string]::IsNullOrWhiteSpace($token) -or
        [string]::IsNullOrWhiteSpace($namespace)) {
        return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) host, token and namespace are required"
    }

    try {
        $script:Session = @{
            base      = (ConvertTo-BaseUrl -HostValue $hostValue)
            namespace = $namespace
            skip_tls  = $skipTls
            header    = @{ Authorization = "Bearer $token" }
        }

        # Validate the token and namespace by listing VirtualMachines.
        $null = Invoke-OcpApi -Method GET -Path "$($script:KubevirtApi)/namespaces/$namespace/virtualmachines?limit=1"

        Write-DebugLog "Connected to $($script:Session.base), namespace [$namespace]"
        return @{ result = @{ message = "$($script:ProviderNamePrefix) Connected to $($script:Session.base) (namespace: $namespace)" } }
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
        foreach ($vm in Get-VmList) {
            if ((Test-HasProperty -Object $vm -Name 'metadata') -and
                (Test-HasProperty -Object $vm.metadata -Name 'name') -and
                -not [string]::IsNullOrWhiteSpace([string]$vm.metadata.name)) {
                $guests += [string]$vm.metadata.name
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
            return @{ result = (ConvertTo-RasGuestObject -VmName ([string]$ids[0])) }
        }

        $resultMap = @{}
        foreach ($id in $ids) {
            $vmName = [string]$id
            try { $resultMap[$vmName] = ConvertTo-RasGuestObject -VmName $vmName }
            catch {
                Write-DebugLog "Guest get failed for [$vmName]: $($_.Exception.Message)"
                $resultMap[$vmName] = @{
                    id = $vmName; name = $vmName; provider = 'OpenShift'; namespace = (Get-Session).namespace
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
        $ns     = (Get-Session).namespace
        $vmName = [string]$Params.id
        $action = Get-ControlAction -Control ([string]$Params.control)

        if ([string]::IsNullOrWhiteSpace($action)) {
            return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Unsupported guest control: $($Params.control)"
        }

        switch ($action) {
            'delete' {
                Invoke-OcpApi -Method DELETE -Path "$($script:KubevirtApi)/namespaces/$ns/virtualmachines/$vmName" | Out-Null
            }
            'suspend' {
                # Pause is a VirtualMachineInstance subresource and requires a running VM.
                Invoke-OcpApi -Method PUT -Path "$($script:SubresApi)/namespaces/$ns/virtualmachineinstances/$vmName/pause" -Body @{} | Out-Null
            }
            default {
                # start / stop / restart are VirtualMachine subresources.
                Invoke-OcpApi -Method PUT -Path "$($script:SubresApi)/namespaces/$ns/virtualmachines/$vmName/$action" -Body @{} | Out-Null
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
        $ns         = (Get-Session).namespace
        $vmName     = [string]$Params.id
        $isTemplate = [System.Convert]::ToBoolean($Params.is_template)

        # KubeVirt has no native "template" object for an existing VM, so RAS template
        # membership is tracked with a label. The label patch is synchronous.
        $labelValue = if ($isTemplate) { 'true' } else { $null }
        $patch = @{ metadata = @{ labels = @{ $script:TemplateLabel = $labelValue } } }
        Invoke-OcpApi -Method PATCH -Path "$($script:KubevirtApi)/namespaces/$ns/virtualmachines/$vmName" `
            -Body $patch -ContentType 'application/merge-patch+json' | Out-Null

        if ($isTemplate) {
            # A template is expected to be powered off; stop it best-effort.
            try {
                Invoke-OcpApi -Method PUT -Path "$($script:SubresApi)/namespaces/$ns/virtualmachines/$vmName/stop" -Body @{} | Out-Null
            }
            catch { Write-DebugLog "Stop on convert ignored for [$vmName]: $($_.Exception.Message)" }
        }

        # The patch is already applied; report a completed synthetic task.
        return @{ result = @{ task_id = "convert:${ns}:$vmName" } }
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
        $ns        = (Get-Session).namespace
        $sourceVm  = [string]$Params.id
        $cloneName = [string]$Params.name
        $cloneObj  = "ras-clone-$(ConvertTo-K8sName -Value $cloneName)-$(New-UniqueSuffix)"

        # Source is either a named template snapshot (template versioning) or the VM itself.
        $source = @{ apiGroup = 'kubevirt.io'; kind = 'VirtualMachine'; name = $sourceVm }
        if ((Test-HasProperty -Object $Params -Name 'snapshot') -and -not [string]::IsNullOrWhiteSpace([string]$Params.snapshot)) {
            $snap = Find-SnapshotObject -VmName $sourceVm -SnapshotName ([string]$Params.snapshot)
            if ($null -eq $snap) {
                return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Snapshot [$($Params.snapshot)] not found for VM [$sourceVm]"
            }
            $source = @{ apiGroup = 'snapshot.kubevirt.io'; kind = 'VirtualMachineSnapshot'; name = [string]$snap.metadata.name }
        }

        $body = @{
            apiVersion = 'clone.kubevirt.io/v1alpha1'
            kind       = 'VirtualMachineClone'
            metadata   = @{ name = $cloneObj; namespace = $ns }
            spec       = @{
                source = $source
                target = @{ apiGroup = 'kubevirt.io'; kind = 'VirtualMachine'; name = $cloneName }
            }
        }

        Invoke-OcpApi -Method POST -Path "$($script:CloneApi)/namespaces/$ns/virtualmachineclones" -Body $body | Out-Null

        return @{ result = @{ task_id = "clone:${ns}:$cloneObj"; clone_id = $cloneName } }
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
        $ns       = (Get-Session).namespace
        $vmName   = [string]$Params.id
        $rasName  = [string]$Params.name
        $snapObj  = "$vmName-$(ConvertTo-K8sName -Value $rasName)"

        $body = @{
            apiVersion = 'snapshot.kubevirt.io/v1beta1'
            kind       = 'VirtualMachineSnapshot'
            metadata   = @{
                name        = $snapObj
                namespace   = $ns
                annotations = @{
                    $script:SnapshotNameAnno = $rasName
                    $script:SourceVmAnno     = $vmName
                }
            }
            spec = @{ source = @{ apiGroup = 'kubevirt.io'; kind = 'VirtualMachine'; name = $vmName } }
        }

        Invoke-OcpApi -Method POST -Path "$($script:SnapshotApi)/namespaces/$ns/virtualmachinesnapshots" -Body $body | Out-Null
        return @{ result = @{ task_id = "snapshot:${ns}:$snapObj" } }
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
        $ns     = (Get-Session).namespace
        $vmName = [string]$Params.id
        $snap   = Find-SnapshotObject -VmName $vmName -SnapshotName ([string]$Params.name)

        if ($null -eq $snap) {
            # Nothing to delete; report a completed no-op task.
            return @{ result = @{ task_id = "noop:${ns}:deleted" } }
        }

        $snapObj = [string]$snap.metadata.name
        Invoke-OcpApi -Method DELETE -Path "$($script:SnapshotApi)/namespaces/$ns/virtualmachinesnapshots/$snapObj" | Out-Null
        return @{ result = @{ task_id = "snapshot-delete:${ns}:$snapObj" } }
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
        $snap = Find-SnapshotObject -VmName ([string]$Params.id) -SnapshotName ([string]$Params.name)
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
        $ns     = (Get-Session).namespace
        $vmName = [string]$Params.id
        $snap   = Find-SnapshotObject -VmName $vmName -SnapshotName ([string]$Params.name)
        if ($null -eq $snap) {
            return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Snapshot [$($Params.name)] not found for VM [$vmName]"
        }

        $snapObj    = [string]$snap.metadata.name
        $restoreObj = "ras-restore-$(ConvertTo-K8sName -Value $vmName)-$(New-UniqueSuffix)"

        $body = @{
            apiVersion = 'snapshot.kubevirt.io/v1beta1'
            kind       = 'VirtualMachineRestore'
            metadata   = @{ name = $restoreObj; namespace = $ns }
            spec       = @{
                target                    = @{ apiGroup = 'kubevirt.io'; kind = 'VirtualMachine'; name = $vmName }
                virtualMachineSnapshotName = $snapObj
            }
        }

        Invoke-OcpApi -Method POST -Path "$($script:SnapshotApi)/namespaces/$ns/virtualmachinerestores" -Body $body | Out-Null
        return @{ result = @{ task_id = "restore:${ns}:$restoreObj" } }
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
        $ns     = if ($parts.Count -ge 2) { $parts[1] } else { (Get-Session).namespace }
        $name   = if ($parts.Count -ge 3) { $parts[2] } else { '' }

        switch ($kind) {
            'convert' { return @{ result = @{ state = 'completed'; output = @{} } } }
            'noop'    { return @{ result = @{ state = 'completed'; output = @{} } } }

            'snapshot' {
                $obj = Get-OcpResourceOrNull -Path "$($script:SnapshotApi)/namespaces/$ns/virtualmachinesnapshots/$name"
                if ($null -eq $obj) { return @{ result = @{ state = 'running' } } }
                if ((Test-HasProperty -Object $obj -Name 'status') -and
                    (Test-HasProperty -Object $obj.status -Name 'readyToUse') -and
                    [System.Convert]::ToBoolean($obj.status.readyToUse)) {
                    return @{ result = @{ state = 'completed'; output = @{} } }
                }
                if ((Test-HasProperty -Object $obj -Name 'status') -and
                    (Test-HasProperty -Object $obj.status -Name 'phase') -and
                    [string]$obj.status.phase -eq 'Failed') {
                    return @{ result = @{ state = 'failed'; error = @{ code = 1; message = 'Snapshot failed' } } }
                }
                return @{ result = @{ state = 'running' } }
            }

            'snapshot-delete' {
                $obj = Get-OcpResourceOrNull -Path "$($script:SnapshotApi)/namespaces/$ns/virtualmachinesnapshots/$name"
                if ($null -eq $obj) { return @{ result = @{ state = 'completed'; output = @{} } } }
                return @{ result = @{ state = 'running' } }
            }

            'restore' {
                $obj = Get-OcpResourceOrNull -Path "$($script:SnapshotApi)/namespaces/$ns/virtualmachinerestores/$name"
                if ($null -eq $obj) { return @{ result = @{ state = 'running' } } }
                if ((Test-HasProperty -Object $obj -Name 'status') -and
                    (Test-HasProperty -Object $obj.status -Name 'complete') -and
                    [System.Convert]::ToBoolean($obj.status.complete)) {
                    return @{ result = @{ state = 'completed'; output = @{} } }
                }
                return @{ result = @{ state = 'running' } }
            }

            'clone' {
                $obj = Get-OcpResourceOrNull -Path "$($script:CloneApi)/namespaces/$ns/virtualmachineclones/$name"
                if ($null -eq $obj) { return @{ result = @{ state = 'running' } } }
                $phase = ''
                if ((Test-HasProperty -Object $obj -Name 'status') -and (Test-HasProperty -Object $obj.status -Name 'phase')) {
                    $phase = [string]$obj.status.phase
                }
                switch ($phase) {
                    'Succeeded' {
                        $cloneId = ''
                        if ((Test-HasProperty -Object $obj -Name 'spec') -and
                            (Test-HasProperty -Object $obj.spec -Name 'target') -and
                            (Test-HasProperty -Object $obj.spec.target -Name 'name')) {
                            $cloneId = [string]$obj.spec.target.name
                        }
                        return @{ result = @{ state = 'completed'; output = @{ clone_id = $cloneId } } }
                    }
                    'Failed' { return @{ result = @{ state = 'failed'; error = @{ code = 1; message = 'Clone failed' } } } }
                    default  { return @{ result = @{ state = 'running' } } }
                }
            }

            default {
                return @{ result = @{ state = 'completed'; output = @{} } }
            }
        }
    }
    catch {
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to retrieve task info: $($_.Exception.Message)"
    }
}

# Hosts methods are aliases of guests methods for this provider.
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

Write-DebugLog "OpenShift provider process started. PID=$PID"

while ($true) {
    try {
        $inputLine = [Console]::In.ReadLine()
        if ($null -eq $inputLine) {
            Write-DebugLog 'Input stream closed. Exiting.'
            break
        }
        Write-DebugLog "IN (PID=$PID): $inputLine"
        $response = Process-Method -InputLine ($inputLine.Trim())
        Send-Response -ResponseObject $response
    }
    catch {
        Send-Response -ResponseObject (New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to process input: $($_.Exception.Message)")
    }
}
