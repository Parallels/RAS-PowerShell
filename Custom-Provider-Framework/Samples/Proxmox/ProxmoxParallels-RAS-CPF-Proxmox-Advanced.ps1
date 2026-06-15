<#  
.SYNOPSIS  
    Parallels RAS Custom Provider Sample Script for Proxmox VE
.DESCRIPTION  
    This script implements a custom provider for Parallels RAS to integrate with Proxmox VE
    hypervisor. It listens for JSON-RPC requests on standard input, processes them according
    to the defined methods, and returns responses on standard output. The provider supports
    connecting to Proxmox using API tokens, listing VMs, retrieving VM information,
    controlling VM power state, converting VMs to templates, tracking clone operations, and cloning VMs.
    This script reuires PowerShell 7 or later for best compatibility.
.NOTES  
    File Name  : Parallels-RAS-CFP-Proxmox-package2-v2.ps1
    Author     : www.parallels.com
.EXAMPLE
    .\Parallels-RAS-CFP-Proxmox-package2-v2.ps1
    Sample requests in json

    {"method": "provider/connect", "params" : { "settings": {"host":"proxmox.example.com","username":"root@pam","token_name":"automation","token_secret":"XXX"}}}
    {"method": "guests/list"}
    {"method": "guests/control","params":{"control":"start","id":"101"}}
    {"method": "guests/get","params":{"id":"101"}}
    {"method": "guests/get","params":{"id":["101","102"]}}
    {"method": "hosts/get,"params":{"id":"101"}}
    {"method": "guests/convert_to_template","params":{"id":"101"}}
    {"method": "guests/clone","params":{"source_id":"101","target_id":"102","name":"Clone of 101"}}
#>

Set-StrictMode -Version Latest

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$WarningPreference = 'SilentlyContinue'
$VerbosePreference = 'SilentlyContinue'
$InformationPreference = 'SilentlyContinue'
$script:CloneStatePath = 'C:\CFP Scripts\Proxmox-RAS-CloneState.json'

if ($Host.Name -notmatch 'ISE') {
    [Console]::InputEncoding = [System.Text.Encoding]::UTF8
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
}

$stdout = [Console]::OpenStandardOutput()
$writer = New-Object System.IO.StreamWriter($stdout, [System.Text.Encoding]::UTF8)
$writer.AutoFlush = $true

$script:ProviderNamePrefix = 'Proxmox:'
$script:LogPath = 'C:\CFP Scripts\Proxmox-RAS-Provider.log'
$script:ProxmoxSession = $null
$script:TaskContext = @{}
$script:DummyOperationTaskId = '__DUMMY_TASK__'

$script:ErrorCodes = @{
    ParseError     = -32700
    MethodNotFound = -32601
    InvalidParams  = -32602
    InternalError  = -32603
}

function Write-DebugLog {
    param([string]$Message)

    try {
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
        Add-Content -Path $script:LogPath -Value "$timestamp $Message" -Encoding UTF8
    }
    catch {
        # Never emit logging failures to stdout
    }
}

function Get-CloneStateEntryByTaskId {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TaskId
    )

    $all = Get-CloneStateAll

    foreach ($key in $all.Keys) {
        $entry = $all[$key]

        if ($null -ne $entry -and
            $entry.PSObject.Properties.Name -contains 'task_id' -and
            [string]$entry.task_id -eq [string]$TaskId) {

            $ctx = @{}
            foreach ($p in $entry.PSObject.Properties) {
                $ctx[$p.Name] = $p.Value
            }

            if (-not $ctx.ContainsKey('type')) {
                $ctx.type = 'clone'
            }

            if (-not $ctx.ContainsKey('clone_id')) {
                $ctx.clone_id = [string]$key
            }

            Write-DebugLog "CLONE STATE FOUND BY TASK ID for task [$TaskId], clone VM [$($ctx.clone_id)]"
            return $ctx
        }
    }

    Write-DebugLog "CLONE STATE NOT FOUND BY TASK ID for task [$TaskId]"
    return $null
}

function Get-CloneState {
    try {
        if (-not (Test-Path -LiteralPath $script:CloneStatePath)) {
            return @{}
        }

        $raw = Get-Content -LiteralPath $script:CloneStatePath -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($raw)) {
            return @{}
        }

        $obj = $raw | ConvertFrom-Json -ErrorAction Stop
        $state = @{}

        foreach ($p in $obj.PSObject.Properties) {
            $entry = @{}
            foreach ($ep in $p.Value.PSObject.Properties) {
                $entry[$ep.Name] = $ep.Value
            }
            $state[$p.Name] = $entry
        }

        return $state
    }
    catch {
        Write-DebugLog "Failed to load clone state: $($_.Exception.Message)"
        return @{}
    }
}

function Save-CloneState {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$State
    )

    try {
        $dir = Split-Path -Path $script:CloneStatePath -Parent
        if (-not [string]::IsNullOrWhiteSpace($dir) -and -not (Test-Path -LiteralPath $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
        }

        $json = $State | ConvertTo-Json -Depth 10 -Compress
        Set-Content -LiteralPath $script:CloneStatePath -Value $json -Encoding UTF8
    }
    catch {
        Write-DebugLog "Failed to save clone state: $($_.Exception.Message)"
    }
}

function Set-CloneStateEntry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$VmId,

        [Parameter(Mandatory = $true)]
        [hashtable]$Entry
    )

    $state = Get-CloneState
    $state[[string]$VmId] = $Entry
    Save-CloneState -State $state
}

function Get-CloneStateAll {
    try {
        if (-not (Test-Path $script:CloneStatePath)) {
            return @{}
        }

        $raw = Get-Content $script:CloneStatePath -Raw -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($raw)) {
            return @{}
        }

        $data = $raw | ConvertFrom-Json -ErrorAction Stop

        # Convert PSCustomObject → Hashtable
        $ht = @{}
        foreach ($prop in $data.PSObject.Properties) {
            $ht[$prop.Name] = $prop.Value
        }

        return $ht
    }
    catch {
        Write-DebugLog "Failed to read clone state file: $($_.Exception.Message)"
        return @{}
    }
}

function Get-CloneStateEntry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$VmId
    )

    $all = Get-CloneStateAll

    # Direct match
    if ($all.ContainsKey($VmId)) {
        return $all[$VmId]
    }

    # Fallback: match by clone_id
    foreach ($key in $all.Keys) {
        $entry = $all[$key]

        if ($null -ne $entry -and
            $entry.PSObject.Properties.Name -contains 'clone_id' -and
            [string]$entry.clone_id -eq [string]$VmId) {

            Write-DebugLog "CLONE STATE MATCHED via clone_id for VM [$VmId] (key [$key])"
            return $entry
        }
    }

    return $null
}

function Remove-CloneStateEntry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$VmId
    )

    $state = Get-CloneState
    if ($state.ContainsKey([string]$VmId)) {
        $state.Remove([string]$VmId)
        Save-CloneState -State $state
    }
}

function Send-Response {
    param(
        [Parameter(Mandatory = $true)]
        [object]$ResponseObject
    )

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
    param(
        [int]$Code,
        [string]$Message
    )

    return @{
        error = @{
            code    = $Code
            message = $Message
        }
    }
}

function ConvertFrom-JsonSafe {
    param([string]$InputLine)

    try {
        return $InputLine | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        Write-DebugLog "JSON parse failed: $($_.Exception.Message)"
        return $null
    }
}

function Test-RequiredFields {
    param(
        [object]$Data,
        [string[]]$RequiredFields
    )

    foreach ($field in $RequiredFields) {
        $keys = $field -split '\.'
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

function Initialize-CertificateBypass {
    try {
        if ($PSVersionTable.PSEdition -eq 'Core') {
            return
        }

        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
        [System.Net.ServicePointManager]::SecurityProtocol = `
            [System.Net.SecurityProtocolType]::Tls12 -bor `
            [System.Net.SecurityProtocolType]::Tls11 -bor `
            [System.Net.SecurityProtocolType]::Tls
    }
    catch {
        Write-DebugLog "Certificate bypass init failed: $($_.Exception.Message)"
    }
}

function Get-Session {
    if ($null -eq $script:ProxmoxSession) {
        throw 'Session not initialized'
    }

    if ([string]::IsNullOrWhiteSpace($script:ProxmoxSession.host)) {
        throw 'Session host missing'
    }

    if ($null -eq $script:ProxmoxSession.header) {
        throw 'Session header missing'
    }

    return $script:ProxmoxSession
}

function Get-ProxmoxBaseUrl {
    param([hashtable]$Session)
    return ("https://{0}" -f $Session.host).TrimEnd('/')
}

function Invoke-ProxmoxRestMethod {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri,

        [Parameter(Mandatory = $true)]
        [hashtable]$Headers,

        [Parameter(Mandatory = $true)]
        [ValidateSet('GET', 'POST', 'PUT', 'DELETE')]
        [string]$Method,

        [object]$Body = $null
    )

    $irmParams = @{
        Uri         = $Uri
        Headers     = $Headers
        Method      = $Method
        ErrorAction = 'Stop'
    }

    if ($PSVersionTable.PSEdition -eq 'Core') {
        $irmParams.SkipCertificateCheck = $true
        $irmParams.SkipHeaderValidation = $true
    }

    if ($null -ne $Body) {
        $irmParams.Body = $Body
    }

    Write-DebugLog "HTTP $Method $Uri"

    try {
        return Invoke-RestMethod @irmParams
    }
    catch {
        Write-DebugLog "HTTP failure: $($_.Exception.Message)"
        throw
    }
}

function Invoke-ProxmoxApi {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('GET', 'POST', 'DELETE', 'PUT')]
        [string]$Method,

        [Parameter(Mandatory = $true)]
        [string]$Path,

        [hashtable]$Body
    )

    $session = Get-Session
    $base = Get-ProxmoxBaseUrl -Session $session
    $uri = ($base.TrimEnd('/') + '/' + $Path.TrimStart('/'))

    if ($Method -eq 'GET') {
        return Invoke-ProxmoxRestMethod -Uri $uri -Headers $session.header -Method GET
    }

    if ($Method -eq 'DELETE') {
        return Invoke-ProxmoxRestMethod -Uri $uri -Headers $session.header -Method DELETE
    }

    if ($null -eq $Body) {
        $Body = @{}
    }

    return Invoke-ProxmoxRestMethod -Uri $uri -Headers $session.header -Method $Method -Body $Body
}

function Get-ProxmoxClusterVMs {
    $resp = Invoke-ProxmoxApi -Method GET -Path '/api2/json/cluster/resources?type=vm'
    return @($resp.data)
}

function Get-ProxmoxVmNode {
    param(
        [Parameter(Mandatory = $true)]
        [string]$VmId
    )

    $clusterVMs = Get-ProxmoxClusterVMs
    $vm = $clusterVMs | Where-Object { [string]$_.vmid -eq [string]$VmId } | Select-Object -First 1

    if ($null -eq $vm) {
        throw "VM [$VmId] not found in cluster"
    }

    if ([string]::IsNullOrWhiteSpace($vm.node)) {
        throw "VM [$VmId] has no node information"
    }

    return $vm
}

function Map-ProxmoxStateToRasState {
    param([string]$State)

    $normalized = if ($null -ne $State) { $State.ToString().Trim().ToLowerInvariant() } else { 'unknown' }

    switch ($normalized) {
        'running' { return 'powered_on' }
        'stopped' { return 'powered_off' }
        'paused' { return 'suspended' }
        'suspended' { return 'suspended' }
        'shutdown' { return 'powering_off' }
        'halting' { return 'powering_off' }
        'prelaunch' { return 'powering_on' }
        default { return 'powered_off' }
    }
}

function Get-ProxmoxVmCurrentStatus {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Node,

        [Parameter(Mandatory = $true)]
        [string]$VmId
    )

    $resp = Invoke-ProxmoxApi -Method GET -Path "/api2/json/nodes/$Node/qemu/$VmId/status/current"
    return $resp.data
}

function Get-ProxmoxVmConfig {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Node,

        [Parameter(Mandatory = $true)]
        [string]$VmId
    )

    $resp = Invoke-ProxmoxApi -Method GET -Path "/api2/json/nodes/$Node/qemu/$VmId/config"
    return $resp.data
}

function Get-ProxmoxVmGuestAgentInterfaces {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Node,

        [Parameter(Mandatory = $true)]
        [string]$VmId
    )

    try {
        $resp = Invoke-ProxmoxApi -Method GET -Path "/api2/json/nodes/$Node/qemu/$VmId/agent/network-get-interfaces"
        if ($null -ne $resp.data.result) {
            return @($resp.data.result)
        }
        return @($resp.data)
    }
    catch {
        Write-DebugLog "Guest agent interface query failed for VM [$VmId]: $($_.Exception.Message)"
        return @()
    }
}

function Get-ProxmoxVmNetworkData {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Node,

        [Parameter(Mandatory = $true)]
        [string]$VmId
    )

    $ipv4Set = New-Object 'System.Collections.Generic.HashSet[string]'
    $macSet = New-Object 'System.Collections.Generic.HashSet[string]'

    $interfaces = Get-ProxmoxVmGuestAgentInterfaces -Node $Node -VmId $VmId

    foreach ($iface in $interfaces) {
        $mac = $null
        if ($iface.PSObject.Properties.Name -contains 'hardware-address') {
            $mac = [string]$iface.'hardware-address'
        }

        $ipAddresses = @($iface.'ip-addresses')
        foreach ($ip in $ipAddresses) {
            $type = [string]$ip.'ip-address-type'
            $addr = [string]$ip.'ip-address'

            if ($type -eq 'ipv4' -and
                -not [string]::IsNullOrWhiteSpace($addr) -and
                $addr -ne '127.0.0.1' -and
                $addr -notmatch '^169\.254\.') {

                [void]$ipv4Set.Add($addr)

                if (-not [string]::IsNullOrWhiteSpace($mac)) {
                    [void]$macSet.Add($mac.ToUpperInvariant())
                }
            }
        }
    }

    try {
        $cfg = Get-ProxmoxVmConfig -Node $Node -VmId $VmId
        foreach ($prop in $cfg.PSObject.Properties) {
            if ($prop.Name -match '^net\d+$' -and -not [string]::IsNullOrWhiteSpace([string]$prop.Value)) {
                $m = [regex]::Match([string]$prop.Value, '([0-9A-Fa-f]{2}(:[0-9A-Fa-f]{2}){5})')
                if ($m.Success) {
                    [void]$macSet.Add($m.Groups[1].Value.ToUpperInvariant())
                }
            }
        }
    }
    catch {
        Write-DebugLog "Config MAC lookup failed for VM [$VmId]: $($_.Exception.Message)"
    }

    return @{
        IPv4Addresses = @($ipv4Set | Select-Object -First 3)
        MacAddresses  = @($macSet  | Select-Object -First 3)
    }
}

function Get-ProxmoxVmOsType {
    param(
        [object]$CurrentStatus,
        [object]$Config
    )

    if ($null -ne $CurrentStatus -and $CurrentStatus.PSObject.Properties.Name -contains 'ostype') {
        if (-not [string]::IsNullOrWhiteSpace([string]$CurrentStatus.ostype)) {
            return [string]$CurrentStatus.ostype
        }
    }

    if ($null -ne $Config -and $Config.PSObject.Properties.Name -contains 'ostype') {
        if (-not [string]::IsNullOrWhiteSpace([string]$Config.ostype)) {
            return [string]$Config.ostype
        }
    }

    return 'unknown'
}

function Get-ProxmoxVmIsTemplate {
    param([object]$Config)

    if ($null -ne $Config -and $Config.PSObject.Properties.Name -contains 'template') {
        try {
            return ([int]$Config.template -eq 1)
        }
        catch {
            return $false
        }
    }

    return $false
}

function ConvertTo-RasGuestObject {
    param(
        [Parameter(Mandatory = $true)]
        [string]$VmId
    )

    $clusterVm = Get-ProxmoxVmNode -VmId $VmId
    $node = [string]$clusterVm.node

    $current = $null
    $config = $null
    $network = $null

    try {
        $current = Get-ProxmoxVmCurrentStatus -Node $node -VmId $VmId
    }
    catch {
        Write-DebugLog "Current status lookup failed for VM [$VmId]: $($_.Exception.Message)"
    }

    try {
        $config = Get-ProxmoxVmConfig -Node $node -VmId $VmId
    }
    catch {
        Write-DebugLog "Config lookup failed for VM [$VmId]: $($_.Exception.Message)"
    }

    try {
        $network = Get-ProxmoxVmNetworkData -Node $node -VmId $VmId
    }
    catch {
        Write-DebugLog "Network lookup failed for VM [$VmId]: $($_.Exception.Message)"
        $network = @{
            IPv4Addresses = @()
            MacAddresses  = @()
        }
    }

    $rawState = 'unknown'
    if ($null -ne $current) {
        if ($current.PSObject.Properties.Name -contains 'qmpstatus' -and -not [string]::IsNullOrWhiteSpace([string]$current.qmpstatus)) {
            $rawState = [string]$current.qmpstatus
        }
        elseif ($current.PSObject.Properties.Name -contains 'status' -and -not [string]::IsNullOrWhiteSpace([string]$current.status)) {
            $rawState = [string]$current.status
        }
    }

    $name = $null
    if ($null -ne $current -and $current.PSObject.Properties.Name -contains 'name' -and -not [string]::IsNullOrWhiteSpace([string]$current.name)) {
        $name = [string]$current.name
    }
    elseif ($null -ne $clusterVm -and $clusterVm.PSObject.Properties.Name -contains 'name' -and -not [string]::IsNullOrWhiteSpace([string]$clusterVm.name)) {
        $name = [string]$clusterVm.name
    }
    else {
        $name = "VM-$VmId"
    }

    $osType = Get-ProxmoxVmOsType -CurrentStatus $current -Config $config
    $isTemplate = Get-ProxmoxVmIsTemplate -Config $config

    $guestObject = @{
        id            = [string]$VmId
        name          = $name
        provider      = 'Proxmox'
        node          = $node
        state         = (Map-ProxmoxStateToRasState -State $rawState)
        power_state   = $rawState
        host_os       = $osType
        ip            = $(if ($network.IPv4Addresses.Count -gt 0) { $network.IPv4Addresses[0] } else { $null })
        ip_addresses  = @($network.IPv4Addresses)
        mac_addresses = @($network.MacAddresses)
        is_template   = $isTemplate
        type          = 'Virtual Machine'
    }

    Write-DebugLog ("GUEST VMID={0}; Name={1}; Node={2}; State={3}; Template={4}; IPs={5}" -f `
            $guestObject.id,
        $guestObject.name,
        $guestObject.node,
        $guestObject.state,
        $guestObject.is_template,
        ($guestObject.ip_addresses -join ',')
    )

    return $guestObject
}

function Get-TrackedCloneContextByVmId {
    param(
        [Parameter(Mandatory = $true)]
        [string]$VmId
    )

    foreach ($key in @($script:TaskContext.Keys)) {
        $ctx = $script:TaskContext[$key]
        if ($null -ne $ctx -and
            $ctx.ContainsKey('type') -and
            [string]$ctx.type -eq 'clone' -and
            $ctx.ContainsKey('clone_id') -and
            [string]$ctx.clone_id -eq [string]$VmId) {

            Write-DebugLog "TRACKED CLONE FOUND IN MEMORY for VM [$VmId]"
            return @{
                task_id = $key
                context = $ctx
            }
        }
    }

    $persisted = Get-CloneStateEntry -VmId $VmId
    if ($null -ne $persisted) {
        $ctx = @{}
        foreach ($p in $persisted.PSObject.Properties) {
            $ctx[$p.Name] = $p.Value
        }

        if (-not $ctx.ContainsKey('type')) {
            $ctx.type = 'clone'
        }

        if (-not $ctx.ContainsKey('clone_id')) {
 
