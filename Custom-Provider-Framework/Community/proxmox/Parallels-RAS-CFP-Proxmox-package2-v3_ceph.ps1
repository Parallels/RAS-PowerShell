<#  
.SYNOPSIS  
    Parallels RAS Custom Provider Sample Script for Proxmox VE
.DESCRIPTION  
    This sample script implements a custom provider for Parallels RAS to integrate with Proxmox VE
    hypervisor. It listens for JSON-RPC requests on standard input, processes them according
    to the defined methods, and returns responses on standard output. The provider supports
    connecting to Proxmox using API tokens, listing VMs, retrieving VM information,
    controlling VM power state, converting VMs to templates, tracking clone operations, cloning VMs,
    and managing snapshot-based versioned templates. This script requires PowerShell 7 or later for best compatibility.
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
    {"method": "hosts/get","params":{"id":"101"}}
    {"method": "guests/convert","params":{"id":"101","is_template":true}}
    {"method": "guests/clone","params":{"id":"101","name":"Clone of 101","snapshot":"RAS_TEMPLATE_VERSION_1"}}
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

function Invoke-WithStateLock {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LockName,

        [Parameter(Mandatory = $true)]
        [scriptblock]$Action
    )

    # The clone/template-version/maintenance state files can be read-modify-written by
    # more than one provider process at once (RAS can run multiple instances of this
    # script concurrently). A named, machine-wide mutex serializes each file's
    # read-modify-write cycle across processes so a concurrent write can't silently
    # clobber another process's update.
    $mutex = New-Object System.Threading.Mutex($false, "Global\$LockName")
    $acquired = $false
    try {
        $acquired = $mutex.WaitOne(30000)
        return & $Action
    }
    finally {
        if ($acquired) {
            $mutex.ReleaseMutex()
        }
        $mutex.Dispose()
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

    Invoke-WithStateLock -LockName 'ProxmoxRasCloneStateLock' -Action {
        $state = Get-CloneState
        $state[[string]$VmId] = $Entry
        Save-CloneState -State $state
    }
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

    Invoke-WithStateLock -LockName 'ProxmoxRasCloneStateLock' -Action {
        $state = Get-CloneState
        if ($state.ContainsKey([string]$VmId)) {
            $state.Remove([string]$VmId)
            Save-CloneState -State $state
        }
    }
}

# This provider represents template versions as named snapshots directly on the
# source VM rather than freezing it into a real Proxmox template (a frozen template
# can never be started, edited, or snapshotted again, which is incompatible with the
# maintenance-mode edit/snapshot cycle). guests/convert therefore never touches
# Proxmox; it only records this RAS-facing "is this guest currently a template"
# label so guests/get can report it back, matching the CPF contract that is_template
# toggles observably across guests/convert calls.
$script:TemplateFlagStatePath = 'C:\CFP Scripts\Proxmox-RAS-TemplateFlagState.json'

function Get-TemplateFlagState {
    try {
        if (-not (Test-Path -LiteralPath $script:TemplateFlagStatePath)) {
            return @{}
        }

        $raw = Get-Content -LiteralPath $script:TemplateFlagStatePath -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($raw)) {
            return @{}
        }

        $obj = $raw | ConvertFrom-Json -ErrorAction Stop
        $state = @{}

        foreach ($p in $obj.PSObject.Properties) {
            $state[$p.Name] = [bool]$p.Value
        }

        return $state
    }
    catch {
        Write-DebugLog "Failed to load template flag state: $($_.Exception.Message)"
        return @{}
    }
}

function Save-TemplateFlagState {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$State
    )

    try {
        $dir = Split-Path -Path $script:TemplateFlagStatePath -Parent
        if (-not [string]::IsNullOrWhiteSpace($dir) -and -not (Test-Path -LiteralPath $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
        }

        $json = $State | ConvertTo-Json -Depth 10 -Compress
        Set-Content -LiteralPath $script:TemplateFlagStatePath -Value $json -Encoding UTF8
    }
    catch {
        Write-DebugLog "Failed to save template flag state: $($_.Exception.Message)"
    }
}

function Set-IsTemplateFlag {
    param(
        [Parameter(Mandatory = $true)]
        [string]$VmId,

        [Parameter(Mandatory = $true)]
        [bool]$IsTemplate
    )

    Invoke-WithStateLock -LockName 'ProxmoxRasTemplateFlagStateLock' -Action {
        $state = Get-TemplateFlagState
        if ($IsTemplate) {
            $state[[string]$VmId] = $true
        }
        else {
            $state.Remove([string]$VmId)
        }
        Save-TemplateFlagState -State $state
    }
}

function Get-IsTemplateFlag {
    param(
        [Parameter(Mandatory = $true)]
        [string]$VmId
    )

    $state = Get-TemplateFlagState
    return $state.ContainsKey([string]$VmId) -and [bool]$state[[string]$VmId]
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
        ContentType = 'application/json'
    }

    if ($PSVersionTable.PSEdition -eq 'Core') {
        $irmParams.SkipCertificateCheck = $true
        $irmParams.SkipHeaderValidation = $true
    }

    if ($null -ne $Body) {
        if ($Body -is [hashtable]) {
            $irmParams.Body = $Body | ConvertTo-Json -Compress -Depth 10
        }
        else {
            $irmParams.Body = $Body
        }
    }

    Write-DebugLog "HTTP $Method $Uri $(if ($null -ne $Body) { '- Body: ' + $irmParams.Body })"

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

    $response = Invoke-ProxmoxRestMethod -Uri $uri -Headers $session.header -Method $Method -Body $Body
    
    if ($null -eq $response) {
        throw "API response is null for $Method $Path"
    }
    
    return $response
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

function Get-ProxmoxVmSnapshots {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Node,

        [Parameter(Mandatory = $true)]
        [string]$VmId
    )

    $resp = Invoke-ProxmoxApi -Method GET -Path "/api2/json/nodes/$Node/qemu/$VmId/snapshot"
    return @($resp.data)
}

function Find-ProxmoxVmSnapshot {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Node,

        [Parameter(Mandatory = $true)]
        [string]$VmId,

        [Parameter(Mandatory = $true)]
        [string]$SnapshotName
    )

    $snapshots = Get-ProxmoxVmSnapshots -Node $Node -VmId $VmId
    foreach ($snapshot in $snapshots) {
        if ($snapshot.PSObject.Properties.Name -contains 'name' -and [string]$snapshot.name -eq $SnapshotName) {
            return $snapshot
        }
    }

    return $null
}

function Create-ProxmoxVmSnapshot {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Node,

        [Parameter(Mandatory = $true)]
        [string]$VmId,

        [Parameter(Mandatory = $true)]
        [string]$SnapshotName
    )

    $body = @{ snapname = $SnapshotName; vmstate = 0 }
    Write-DebugLog "Creating snapshot [$SnapshotName] for VM [$VmId] on node [$Node]"
    $resp = Invoke-ProxmoxApi -Method POST -Path "/api2/json/nodes/$Node/qemu/$VmId/snapshot" -Body $body
    
    $taskId = if ($resp.PSObject.Properties.Name -contains 'data') { [string]$resp.data } else { $null }
    if ($null -eq $taskId) {
        throw "Snapshot creation failed: no task ID returned. Response was: $(($resp | ConvertTo-Json -Compress))"
    }
    
    Write-DebugLog "Snapshot [$SnapshotName] creation task ID: $taskId"
    return $taskId
}

function Delete-ProxmoxVmSnapshot {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Node,

        [Parameter(Mandatory = $true)]
        [string]$VmId,

        [Parameter(Mandatory = $true)]
        [string]$SnapshotName
    )

    $encodedName = [System.Uri]::EscapeDataString($SnapshotName)
    Write-DebugLog "Deleting snapshot [$SnapshotName] for VM [$VmId] on node [$Node]"
    $resp = Invoke-ProxmoxApi -Method DELETE -Path "/api2/json/nodes/$Node/qemu/$VmId/snapshot/$encodedName"
    
    $taskId = if ($resp.PSObject.Properties.Name -contains 'data') { [string]$resp.data } else { $null }
    if ($null -eq $taskId) {
        throw "Snapshot deletion failed: no task ID returned. Response was: $(($resp | ConvertTo-Json -Compress))"
    }
    
    Write-DebugLog "Snapshot [$SnapshotName] deletion task ID: $taskId"
    return $taskId
}

function Revert-ProxmoxVmSnapshot {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Node,

        [Parameter(Mandatory = $true)]
        [string]$VmId,

        [Parameter(Mandatory = $true)]
        [string]$SnapshotName
    )

    $encodedName = [System.Uri]::EscapeDataString($SnapshotName)
    Write-DebugLog "Reverting VM [$VmId] on node [$Node] to snapshot [$SnapshotName]"
    $resp = Invoke-ProxmoxApi -Method POST -Path "/api2/json/nodes/$Node/qemu/$VmId/snapshot/$encodedName/rollback" -Body @{}
    
    $taskId = if ($resp.PSObject.Properties.Name -contains 'data') { [string]$resp.data } else { $null }
    if ($null -eq $taskId) {
        throw "Snapshot revert failed: no task ID returned. Response was: $(($resp | ConvertTo-Json -Compress))"
    }
    
    Write-DebugLog "Snapshot revert task ID: $taskId"
    return $taskId
}

function ProxmoxVmSnapshotExists {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Node,

        [Parameter(Mandatory = $true)]
        [string]$VmId,

        [Parameter(Mandatory = $true)]
        [string]$SnapshotName
    )

    $snapshot = Find-ProxmoxVmSnapshot -Node $Node -VmId $VmId -SnapshotName $SnapshotName
    return ($null -ne $snapshot)
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
    $isTemplate = (Get-IsTemplateFlag -VmId $VmId) -or (Get-ProxmoxVmIsTemplate -Config $config)

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
            $ctx.clone_id = [string]$VmId
        }

        Write-DebugLog "TRACKED CLONE FOUND IN FILE for VM [$VmId]"
        return @{
            task_id = $null
            context = $ctx
        }
    }

    Write-DebugLog "TRACKED CLONE NOT FOUND for VM [$VmId]"
    return $null
}

function Start-ProxmoxVmIfNeeded {
    param(
        [Parameter(Mandatory = $true)]
        [string]$VmId
    )

    $clusterVm = Get-ProxmoxVmNode -VmId $VmId
    $node = [string]$clusterVm.node
    $current = Get-ProxmoxVmCurrentStatus -Node $node -VmId $VmId

    $rawState = ''
    if ($current.PSObject.Properties.Name -contains 'qmpstatus' -and -not [string]::IsNullOrWhiteSpace([string]$current.qmpstatus)) {
        $rawState = [string]$current.qmpstatus
    }
    elseif ($current.PSObject.Properties.Name -contains 'status' -and -not [string]::IsNullOrWhiteSpace([string]$current.status)) {
        $rawState = [string]$current.status
    }

    $rasState = Map-ProxmoxStateToRasState -State $rawState

    if ($rasState -eq 'powered_on' -or $rasState -eq 'powering_on') {
        return @{
            started       = $false
            pending       = $false
            node          = $node
            raw_state     = $rawState
            ras_state     = $rasState
            task_id       = $null
            error_message = $null
        }
    }

    try {
        $resp = Invoke-ProxmoxApi -Method POST -Path "/api2/json/nodes/$node/qemu/$VmId/status/start" -Body @{}
        $startTaskId = $null
        if ($null -ne $resp -and $resp.PSObject.Properties.Name -contains 'data') {
            $startTaskId = [string]$resp.data
        }

        Write-DebugLog "Issued start for VM [$VmId], start task id=[$startTaskId]"

        return @{
            started       = $true
            pending       = $false
            node          = $node
            raw_state     = $rawState
            ras_state     = $rasState
            task_id       = $startTaskId
            error_message = $null
        }
    }
    catch {
        $msg = $_.Exception.Message
        Write-DebugLog "Start attempt for VM [$VmId] failed: $msg"

        if ($msg -match "can't lock file" -or $msg -match 'got timeout' -or $msg -match 'VM is locked') {
            return @{
                started       = $false
                pending       = $true
                node          = $node
                raw_state     = $rawState
                ras_state     = $rasState
                task_id       = $null
                error_message = $msg
            }
        }

        throw
    }
}

function Get-RasGuestObjectForCloneAwareFlow {
    param(
        [Parameter(Mandatory = $true)]
        [string]$VmId
    )

    $guest = ConvertTo-RasGuestObject -VmId $VmId
    $tracked = Get-TrackedCloneContextByVmId -VmId $VmId

    if ($null -eq $tracked) {
        Write-DebugLog "CLONE-AWARE FLOW: no tracked clone context for VM [$VmId]"
        return $guest
    }

    Write-DebugLog "CLONE-AWARE FLOW: tracked clone context found for VM [$VmId]"

    $ctx = $tracked.context

    if (-not $ctx.ContainsKey('start_issued')) {
        $ctx.start_issued = $false
    }

    if (-not $ctx.ContainsKey('start_pending')) {
        $ctx.start_pending = $false
    }

    if (-not $ctx.ContainsKey('start_retry_count')) {
        $ctx.start_retry_count = 0
    }

    if (-not $ctx.ContainsKey('creation_completed')) {
        $ctx.creation_completed = $false
    }

    try {
        $ctx.start_issued = [System.Convert]::ToBoolean($ctx.start_issued)
    }
    catch {
        $ctx.start_issued = $false
    }

    try {
        $ctx.start_pending = [System.Convert]::ToBoolean($ctx.start_pending)
    }
    catch {
        $ctx.start_pending = $false
    }

    try {
        $ctx.creation_completed = [System.Convert]::ToBoolean($ctx.creation_completed)
    }
    catch {
        $ctx.creation_completed = $false
    }

    try {
        $ctx.start_retry_count = [int]$ctx.start_retry_count
    }
    catch {
        $ctx.start_retry_count = 0
    }

    # Only force a start while the clone is still being provisioned for the first time.
    # Once creation_completed is set, this VM is an ordinary guest - a deliberate
    # later power-off (entering template maintenance, an admin suspending a desktop,
    # RAS's own idle-desktop power management) must stick instead of being silently
    # undone the next time RAS polls guests/get.
    if (-not $ctx.creation_completed -and ($guest.state -eq 'powered_off' -or $guest.state -eq 'powering_off')) {
        $startInfo = Start-ProxmoxVmIfNeeded -VmId $VmId
        $ctx.start_retry_count = [int]$ctx.start_retry_count + 1
        $ctx.clone_node = $startInfo.node

        if ($startInfo.started) {
            $ctx.start_issued = $true
            $ctx.start_pending = $false
            $ctx.start_task_id = $startInfo.task_id
            Write-DebugLog "Clone-aware get: VM [$VmId] was off, start issued."
        }
        elseif ($startInfo.pending) {
            $ctx.start_pending = $true
            Write-DebugLog "Clone-aware get: VM [$VmId] still locked, start deferred."
        }

        $guest.state = 'powering_on'
        $guest.power_state = 'starting'
    }
    elseif ($guest.state -eq 'powered_on') {
        $hasIp = ($null -ne $guest.ip_addresses -and @($guest.ip_addresses).Count -gt 0)

        if ($hasIp) {
            $ctx.creation_completed = $true
            Write-DebugLog "Clone-aware get: VM [$VmId] is powered on and has IP(s) [$($guest.ip_addresses -join ',')]."
        }
        else {
            $guest.state = 'powering_on'
            $guest.power_state = 'starting'
            Write-DebugLog "Clone-aware get: VM [$VmId] is powered on but has no IP yet. Reporting powering_on."
        }
    }
    else {
        Write-DebugLog "Clone-aware get: VM [$VmId] currently in state [$($guest.state)]."
    }

    if ($null -ne $tracked.task_id) {
        $script:TaskContext[$tracked.task_id] = $ctx
    }

    Set-CloneStateEntry -VmId $VmId -Entry $ctx

    Write-DebugLog "CLONE-AWARE FLOW RESULT for VM [$VmId]: state=[$($guest.state)] power_state=[$($guest.power_state)]"

    return $guest
}

function Get-ControlAction {
    param([Parameter(Mandatory = $true)][string]$Control)

    switch ($Control.Trim().ToLowerInvariant()) {
        'start' { return 'start' }
        'stop' { return 'shutdown' }
        'reset' { return 'reset' }
        'restart' { return 'reboot' }
        'reboot' { return 'reboot' }
        'delete' { return 'delete' }
        'suspend' { return 'suspend' }
        default { return $null }
    }
}

function Get-ProxmoxNextVmId {
    $resp = Invoke-ProxmoxApi -Method GET -Path '/api2/json/cluster/nextid'
    return [string]$resp.data
}

function Get-ProxmoxTaskNodeFromUpid {
    param([string]$Upid)

    if ([string]::IsNullOrWhiteSpace($Upid)) {
        throw 'Task id is empty'
    }

    $parts = $Upid -split ':'
    if ($parts.Count -lt 3 -or $parts[0] -ne 'UPID') {
        throw "Invalid Proxmox task id format: $Upid"
    }

    return $parts[1]
}

function Get-ProxmoxTaskStatus {
    param([string]$TaskId)

    $node = Get-ProxmoxTaskNodeFromUpid -Upid $TaskId
    $escapedTaskId = [System.Uri]::EscapeDataString($TaskId)
    $resp = Invoke-ProxmoxApi -Method GET -Path "/api2/json/nodes/$node/tasks/$escapedTaskId/status"
    return $resp.data
}

function New-TaskResultState {
    param([object]$TaskStatus)

    if ($null -eq $TaskStatus) {
        return @{
            state = 'failed'
            error = @{
                code    = 1
                message = 'Task status unavailable'
            }
        }
    }

    $status = ''
    if ($TaskStatus.PSObject.Properties.Name -contains 'status' -and $TaskStatus.status) {
        $status = [string]$TaskStatus.status
    }

    $exitStatus = $null
    if ($TaskStatus.PSObject.Properties.Name -contains 'exitstatus') {
        $exitStatus = [string]$TaskStatus.exitstatus
    }

    if ($status -eq 'running') {
        return @{ state = 'running' }
    }

    if ($status -eq 'stopped' -and $exitStatus -eq 'OK') {
        return @{ state = 'completed' }
    }

    $msg = if (-not [string]::IsNullOrWhiteSpace($exitStatus)) { $exitStatus } else { 'Unknown task failure' }
    return @{
        state = 'failed'
        error = @{
            code    = 1
            message = $msg
        }
    }
}

function Handle-Initialize {
    return @{
        result = @{
            version      = '1.0.0'
            capabilities = @{
                can_suspend_guests    = $true
                guests_polling_rate   = 5
                tasks_polling_rate    = 30
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

    $proxmoxHost = [string]$settings.host
    $username = [string]$settings.username
    $tokenName = [string]$settings.token_name
    $tokenSecret = [string]$settings.token_secret

    if ([string]::IsNullOrWhiteSpace($proxmoxHost) -or
        [string]::IsNullOrWhiteSpace($username) -or
        [string]::IsNullOrWhiteSpace($tokenName) -or
        [string]::IsNullOrWhiteSpace($tokenSecret)) {

        return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Invalid connection parameters"
    }

    try {
        Initialize-CertificateBypass

        $header = @{ Authorization = "PVEAPIToken=$username!$tokenName=$tokenSecret" }

        $script:ProxmoxSession = @{
            host         = $proxmoxHost
            user         = $username
            token_name   = $tokenName
            token_secret = $tokenSecret
            header       = $header
        }

        $resp = Invoke-ProxmoxApi -Method GET -Path '/api2/json/version'
        $version = [string]$resp.data.version

        Write-DebugLog "Connected successfully to $proxmoxHost as $username, version=$version"

        return @{ result = @{ message = "$($script:ProviderNamePrefix) Connected successfully to Proxmox $proxmoxHost (version: $version)" } }
    }
    catch {
        $script:ProxmoxSession = $null
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to connect to Proxmox: $($_.Exception.Message)"
    }
}

function Handle-Disconnect {
    try {
        $hostName = $null
        if ($null -ne $script:ProxmoxSession) {
            $hostName = $script:ProxmoxSession.host
        }

        $script:ProxmoxSession = $null
        $script:TaskContext = @{}

        try {
            if (Test-Path -LiteralPath $script:CloneStatePath) {
                Remove-Item -LiteralPath $script:CloneStatePath -Force -ErrorAction Stop
            }
        }
        catch {
            Write-DebugLog "Failed to clear clone state file: $($_.Exception.Message)"
        }

        return @{ result = @{ message = "Session cleared on Proxmox $hostName" } }
    }
    catch {
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to clear session: $($_.Exception.Message)"
    }
}

function Handle-HostList {
    try {
        $clusterVMs = Get-ProxmoxClusterVMs
        $hosts = @()

        foreach ($vm in $clusterVMs) {
            $vmId = [string]$vm.vmid
            if (-not [string]::IsNullOrWhiteSpace($vmId)) {
                $hosts += $vmId
            }
        }

        return @{ result = @{ guests = @($hosts) } }
    }
    catch {
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to retrieve host list: $($_.Exception.Message)"
    }
}

function Handle-HostGet {
    param([object]$Params)

    if ($null -eq $Params -or $null -eq $Params.id) {
        return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Invalid or missing host ID"
    }

    try {
        $ids = @($Params.id)

        if ($ids.Count -eq 1) {
            $hostObj = Get-RasGuestObjectForCloneAwareFlow -VmId ([string]$ids[0])
            return @{ result = $hostObj }
        }

        $resultMap = @{}
        foreach ($id in $ids) {
            $vmId = [string]$id
            try {
                $resultMap[$vmId] = Get-RasGuestObjectForCloneAwareFlow -VmId $vmId
            }
            catch {
                $resultMap[$vmId] = @{
                    id            = $vmId
                    name          = $null
                    provider      = 'Proxmox'
                    state         = 'powered_off'
                    power_state   = 'unknown'
                    host_os       = 'unknown'
                    ip            = $null
                    ip_addresses  = @()
                    mac_addresses = @()
                    is_template   = $false
                    type          = 'Virtual Machine'
                }
                Write-DebugLog "Host get failed for VM [$vmId]: $($_.Exception.Message)"
            }
        }

        return @{ result = $resultMap }
    }
    catch {
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to retrieve host info: $($_.Exception.Message)"
    }
}

function Handle-HostControl {
    param([object]$Params)
    return Handle-GuestControl -Params $Params
}

function Handle-GuestList {
    try {
        $clusterVMs = Get-ProxmoxClusterVMs
        $guests = @()

        foreach ($vm in $clusterVMs) {
            $vmId = [string]$vm.vmid
            if (-not [string]::IsNullOrWhiteSpace($vmId)) {
                $guests += $vmId
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
            Write-DebugLog "HANDLE-GUESTGET using clone-aware flow for VM [$([string]$ids[0])]"
            $guest = Get-RasGuestObjectForCloneAwareFlow -VmId ([string]$ids[0])
            return @{ result = $guest }
        }

        $resultMap = @{}
        foreach ($id in $ids) {
            $vmId = [string]$id
            try {
                $resultMap[$vmId] = Get-RasGuestObjectForCloneAwareFlow -VmId $vmId
            }
            catch {
                $resultMap[$vmId] = @{
                    id            = $vmId
                    name          = $null
                    provider      = 'Proxmox'
                    state         = 'powered_off'
                    power_state   = 'unknown'
                    host_os       = 'unknown'
                    ip            = $null
                    ip_addresses  = @()
                    mac_addresses = @()
                    is_template   = $false
                    type          = 'Virtual Machine'
                }
                Write-DebugLog "Guest get failed for VM [$vmId]: $($_.Exception.Message)"
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
        $vmId = [string]$Params.id
        $requestedControl = [string]$Params.control
        $action = Get-ControlAction -Control $requestedControl

        if ([string]::IsNullOrWhiteSpace($action)) {
            return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Unsupported guest control: $requestedControl"
        }

        $clusterVm = Get-ProxmoxVmNode -VmId $vmId
        $node = [string]$clusterVm.node

        if ($action -eq 'delete') {
            $resp = Invoke-ProxmoxApi -Method DELETE -Path "/api2/json/nodes/$node/qemu/$vmId"
        }
        else {
            $resp = Invoke-ProxmoxApi -Method POST -Path "/api2/json/nodes/$node/qemu/$vmId/status/$action" -Body @{}
        }

        $upid = $null
        if ($null -ne $resp -and $resp.PSObject.Properties.Name -contains 'data') {
            $upid = $resp.data
        }

        return @{
            result = @{
                id      = $vmId
                node    = $node
                action  = $action
                upid    = $upid
                message = "$($script:ProviderNamePrefix) Guest control [$requestedControl] submitted successfully"
            }
        }
    }
    catch {
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to control guest [$($Params.control)]: $($_.Exception.Message)"
    }
}

function Handle-TaskInfo {
    param([object]$Params)

    if ($null -eq $Params -or [string]::IsNullOrWhiteSpace([string]$Params.id)) {
        return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Invalid task id"
    }

    try {
        $taskId = [string]$Params.id

        # guests/convert issues a synthetic (non-Proxmox) task id since it takes no
        # Proxmox-side action - short-circuit before Get-ProxmoxTaskStatus, which
        # expects a real UPID and would error on a fabricated one.
        if ($script:TaskContext.ContainsKey($taskId) -and [string]$script:TaskContext[$taskId].type -eq 'noop') {
            return @{ result = @{ state = 'completed'; output = @{} } }
        }

        $taskStatus = Get-ProxmoxTaskStatus -TaskId $taskId
        $taskResult = New-TaskResultState -TaskStatus $taskStatus

        if ($taskResult.state -eq 'failed') {
            return @{
                result = @{
                    state = 'failed'
                    error = $taskResult.error
                }
            }
        }

        if ($taskResult.state -eq 'running') {
            return @{ result = @{ state = 'running' } }
        }

        $ctx = $null

        if ($script:TaskContext.ContainsKey($taskId)) {
            $ctx = $script:TaskContext[$taskId]
            Write-DebugLog "TASK CONTEXT FOUND IN MEMORY for task [$taskId]"
        }
        else {
            $ctx = Get-CloneStateEntryByTaskId -TaskId $taskId
            if ($null -ne $ctx) {
                $script:TaskContext[$taskId] = $ctx
            }
        }

        if ($null -ne $ctx -and $ctx.ContainsKey('type') -and [string]$ctx.type -eq 'snapshot_create_after_delete') {
            # Reaching this point means the stale snapshot's delete task (this
            # $taskId) already reported 'completed' via the generic check above.
            # Chain the actual replacement snapshot creation onto the same
            # externally-visible task_id, the same way the 'clone' branch below
            # chains VM start + IP wait onto the clone task_id, so RAS keeps
            # polling one id across both underlying Proxmox operations instead of
            # this handler blocking on the delete synchronously.
            if (-not $ctx.create_issued) {
                $createTaskId = Create-ProxmoxVmSnapshot -Node $ctx.node -VmId $ctx.vm_id -SnapshotName $ctx.snapshot_name
                $ctx.create_issued = $true
                $ctx.create_task_id = $createTaskId
                $script:TaskContext[$taskId] = $ctx
                Write-DebugLog "Snapshot delete task [$taskId] completed; issued replacement snapshot [$($ctx.snapshot_name)] creation as task [$createTaskId]."
                return @{ result = @{ state = 'running' } }
            }

            $createTaskStatus = Get-ProxmoxTaskStatus -TaskId $ctx.create_task_id
            $createTaskResult = New-TaskResultState -TaskStatus $createTaskStatus

            if ($createTaskResult.state -eq 'failed') {
                return @{
                    result = @{
                        state = 'failed'
                        error = $createTaskResult.error
                    }
                }
            }

            if ($createTaskResult.state -eq 'running') {
                return @{ result = @{ state = 'running' } }
            }

            Write-DebugLog "Snapshot [$($ctx.snapshot_name)] creation task [$($ctx.create_task_id)] completed for guest [$($ctx.vm_id)]."
            return @{ result = @{ state = 'completed'; output = @{} } }
        }

        if ($null -ne $ctx -and $ctx.ContainsKey('type') -and [string]$ctx.type -eq 'clone') {
            $cloneId = [string]$ctx.clone_id

            if ([string]::IsNullOrWhiteSpace($cloneId)) {
                Write-DebugLog "Clone task [$taskId] has no clone_id. Completing without output."
                return @{
                    result = @{
                        state  = 'completed'
                        output = @{}
                    }
                }
            }

            $guest = ConvertTo-RasGuestObject -VmId $cloneId

            if ($guest.state -eq 'powered_off' -or $guest.state -eq 'powering_off') {
                $startInfo = Start-ProxmoxVmIfNeeded -VmId $cloneId

                if (-not $ctx.ContainsKey('start_retry_count')) {
                    $ctx.start_retry_count = 0
                }

                $ctx.start_retry_count = [int]$ctx.start_retry_count + 1
                $ctx.clone_node = $startInfo.node

                if ($startInfo.started) {
                    $ctx.start_issued = $true
                    $ctx.start_pending = $false
                    $ctx.start_task_id = $startInfo.task_id
                    Write-DebugLog "Clone task [$taskId]: start issued successfully for clone VM [$cloneId]."
                }
                elseif ($startInfo.pending) {
                    $ctx.start_pending = $true
                    Write-DebugLog "Clone task [$taskId]: start deferred for clone VM [$cloneId] because lock is still held."
                }

                $script:TaskContext[$taskId] = $ctx
                Set-CloneStateEntry -VmId $cloneId -Entry $ctx
                return @{ result = @{ state = 'running' } }
            }

            if ($guest.state -eq 'powering_on') {
                Write-DebugLog "Clone task [$taskId]: VM [$cloneId] is powering on. Waiting."
                return @{ result = @{ state = 'running' } }
            }

            if ($guest.state -ne 'powered_on') {
                Write-DebugLog "Clone task [$taskId]: VM [$cloneId] state is [$($guest.state)]. Waiting."
                return @{ result = @{ state = 'running' } }
            }

            if ($null -eq $guest.ip_addresses -or @($guest.ip_addresses).Count -eq 0) {
                Write-DebugLog "Clone task [$taskId]: VM [$cloneId] is powered on but has no IP yet. Waiting."
                return @{ result = @{ state = 'running' } }
            }

            $ctx.creation_completed = $true
            $script:TaskContext[$taskId] = $ctx
            Set-CloneStateEntry -VmId $cloneId -Entry $ctx

            Write-DebugLog "Clone task [$taskId]: VM [$cloneId] is powered on and has IP [$($guest.ip_addresses -join ',')]. Completing task."

            return @{
                result = @{
                    state  = 'completed'
                    output = @{
                        clone_id = $cloneId
                    }
                }
            }
        }

        return @{
            result = @{
                state  = 'completed'
                output = @{}
            }
        }
    }
    catch {
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to retrieve task info: $($_.Exception.Message)"
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
        $vmId = [string]$Params.id
        $isTemplate = [bool]$Params.is_template

        # This provider represents template versions as named snapshots directly on
        # the source VM (the newest RAS_TEMPLATE_VERSION_* snapshot is the current
        # version) rather than freezing it into a real Proxmox template - Proxmox
        # templates can never be started, edited, or snapshotted again, which is
        # incompatible with the maintenance-mode edit/snapshot cycle. So guests/convert
        # takes no Proxmox-side action; it only records the RAS-facing is_template
        # label so guests/get can report it back (the CPF contract expects it to
        # toggle observably across enter/exit maintenance).
        Set-IsTemplateFlag -VmId $vmId -IsTemplate $isTemplate

        $taskId = [guid]::NewGuid().ToString()
        $script:TaskContext[$taskId] = @{ type = 'noop' }

        return @{ result = @{ task_id = $taskId } }
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
        $sourceVmId = [string]$Params.id
        $cloneName = [string]$Params.name
        $snapshotName = if ($null -ne $Params.snapshot -and -not [string]::IsNullOrWhiteSpace([string]$Params.snapshot)) { [string]$Params.snapshot } else { $null }
        $clusterVm = Get-ProxmoxVmNode -VmId $sourceVmId
        $node = [string]$clusterVm.node
        $newVmId = Get-ProxmoxNextVmId

        $body = @{
            newid = [int]$newVmId
            name  = $cloneName
            full  = 1
        }

        if ($null -ne $snapshotName) {
            $body.snapname = $snapshotName
        }

        $resp = Invoke-ProxmoxApi -Method POST -Path "/api2/json/nodes/$node/qemu/$sourceVmId/clone" -Body $body
        $taskId = if ($resp.PSObject.Properties.Name -contains 'data') { [string]$resp.data } else { $null }
        
        if ($null -eq $taskId) {
            Write-DebugLog "Clone API returned no task ID. Response: $(($resp | ConvertTo-Json -Compress))"
            throw "Clone operation failed: no task ID returned from Proxmox API"
        }

        if (-not [string]::IsNullOrWhiteSpace($taskId)) {
            $script:TaskContext[$taskId] = @{
                type               = 'clone'
                source_id          = $sourceVmId
                clone_id           = $newVmId
                name               = $cloneName
                full               = $true
                clone_node         = $node
                start_issued       = $false
                start_task_id      = $null
                start_pending      = $false
                start_retry_count  = 0
                creation_completed = $false
            }
        }

        Set-CloneStateEntry -VmId $newVmId -Entry @{
            type               = 'clone'
            task_id            = $taskId
            source_id          = $sourceVmId
            clone_id           = $newVmId
            name               = $cloneName
            full               = $true
            clone_node         = $node
            start_issued       = $false
            start_task_id      = $null
            start_pending      = $false
            start_retry_count  = 0
            creation_completed = $false
        }

        return @{
            result = @{
                task_id  = $taskId
                clone_id = $newVmId
            }
        }
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

    if ($null -eq $Params.name -or [string]::IsNullOrWhiteSpace([string]$Params.name)) {
        return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Invalid snapshot name"
    }

    try {
        $vmId = [string]$Params.id
        $snapshotName = [string]$Params.name
        $clusterVm = Get-ProxmoxVmNode -VmId $vmId
        $node = [string]$clusterVm.node

        if (ProxmoxVmSnapshotExists -Node $node -VmId $vmId -SnapshotName $snapshotName) {
            # A prior attempt at this same step (e.g. RAS retrying template creation
            # after a later step failed) can leave this snapshot behind. Proxmox refuses
            # to create a second snapshot with the same name ("snapshot name
            # 'RAS_TEMPLATE_VERSION_1' already used"), so replace the stale one instead
            # of failing - the fresh snapshot must reflect the VM's current state
            # regardless of what an earlier, incomplete attempt captured.
            #
            # This must not block on the delete (Ceph/RBD snapshot ops are slower than
            # local storage, and the CPF contract expects guests/snapshots/create to
            # return a task_id immediately, not hold the single-threaded stdio loop
            # hostage while RAS waits for a synchronous response). Instead, return the
            # delete's own task_id and let Handle-TaskInfo chain the actual snapshot
            # creation once that delete is confirmed complete, the same pattern already
            # used for clone -> start -> wait-for-IP.
            Write-DebugLog "Snapshot [$snapshotName] already exists on guest [$vmId] (likely left over from a prior retry); replacing it."
            $staleDeleteTaskId = Delete-ProxmoxVmSnapshot -Node $node -VmId $vmId -SnapshotName $snapshotName

            $script:TaskContext[$staleDeleteTaskId] = @{
                type          = 'snapshot_create_after_delete'
                node          = $node
                vm_id         = $vmId
                snapshot_name = $snapshotName
                create_issued = $false
                create_task_id = $null
            }

            return @{ result = @{ task_id = $staleDeleteTaskId } }
        }

        $taskId = Create-ProxmoxVmSnapshot -Node $node -VmId $vmId -SnapshotName $snapshotName

        return @{ result = @{ task_id = $taskId } }
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

    if ($null -eq $Params.name -or [string]::IsNullOrWhiteSpace([string]$Params.name)) {
        return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Invalid snapshot name"
    }

    try {
        $vmId = [string]$Params.id
        $snapshotName = [string]$Params.name
        $clusterVm = Get-ProxmoxVmNode -VmId $vmId
        $node = [string]$clusterVm.node

        $taskId = Delete-ProxmoxVmSnapshot -Node $node -VmId $vmId -SnapshotName $snapshotName

        return @{ result = @{ task_id = $taskId } }
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

    if ($null -eq $Params.name -or [string]::IsNullOrWhiteSpace([string]$Params.name)) {
        return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Invalid snapshot name"
    }

    try {
        $vmId = [string]$Params.id
        $snapshotName = [string]$Params.name
        $clusterVm = Get-ProxmoxVmNode -VmId $vmId
        $node = [string]$clusterVm.node
        $exists = ProxmoxVmSnapshotExists -Node $node -VmId $vmId -SnapshotName $snapshotName

        return @{ result = $exists }
    }
    catch {
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to check snapshot existence [$($Params.name)] for guest [$($Params.id)]: $($_.Exception.Message)"
    }
}

function Handle-GuestSnapshotsRevert {
    param([object]$Params)

    if ($null -eq $Params -or [string]::IsNullOrWhiteSpace([string]$Params.id)) {
        return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Invalid guest id"
    }

    if ($null -eq $Params.name -or [string]::IsNullOrWhiteSpace([string]$Params.name)) {
        return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Invalid snapshot name"
    }

    try {
        $vmId = [string]$Params.id
        $snapshotName = [string]$Params.name
        $clusterVm = Get-ProxmoxVmNode -VmId $vmId
        $node = [string]$clusterVm.node
        $taskId = Revert-ProxmoxVmSnapshot -Node $node -VmId $vmId -SnapshotName $snapshotName

        return @{ result = @{ task_id = $taskId } }
    }
    catch {
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to revert snapshot [$($Params.name)] for guest [$($Params.id)]: $($_.Exception.Message)"
    }
}

$script:MethodRegistry = @{
    'provider/initialize' = @{ Handler = { param($data) Handle-Initialize }; RequiredFields = @() }
    'provider/connect'    = @{ Handler = { param($data) Handle-Connect -Params $data.params }; RequiredFields = @('params.settings') }
    'provider/disconnect' = @{ Handler = { param($data) Handle-Disconnect }; RequiredFields = @() }

    'hosts/list'          = @{ Handler = { param($data) Handle-HostList }; RequiredFields = @() }
    'hosts/get'           = @{ Handler = { param($data) Handle-HostGet -Params $data.params }; RequiredFields = @('params.id') }
    'hosts/control'       = @{ Handler = { param($data) Handle-HostControl -Params $data.params }; RequiredFields = @('params.id', 'params.control') }

    'guests/list'         = @{ Handler = { param($data) Handle-GuestList }; RequiredFields = @() }
    'guests/get'          = @{ Handler = { param($data) Handle-GuestGet -Params $data.params }; RequiredFields = @('params.id') }
    'guests/control'      = @{ Handler = { param($data) Handle-GuestControl -Params $data.params }; RequiredFields = @('params.id', 'params.control') }

    'guests/convert'             = @{ Handler = { param($data) Handle-GuestConvert -Params $data.params }; RequiredFields = @('params.id', 'params.is_template') }
    'guests/clone'               = @{ Handler = { param($data) Handle-GuestClone -Params $data.params }; RequiredFields = @('params.id', 'params.name') }
    'guests/snapshots/create'    = @{ Handler = { param($data) Handle-GuestSnapshotsCreate -Params $data.params }; RequiredFields = @('params.id', 'params.name') }
    'guests/snapshots/delete'    = @{ Handler = { param($data) Handle-GuestSnapshotsDelete -Params $data.params }; RequiredFields = @('params.id', 'params.name') }
    'guests/snapshots/exists'    = @{ Handler = { param($data) Handle-GuestSnapshotsExists -Params $data.params }; RequiredFields = @('params.id', 'params.name') }
    'guests/snapshots/revert'    = @{ Handler = { param($data) Handle-GuestSnapshotsRevert -Params $data.params }; RequiredFields = @('params.id', 'params.name') }
    'tasks/get'                  = @{ Handler = { param($data) Handle-TaskInfo -Params $data.params }; RequiredFields = @('params.id') }
}

function Process-Method {
    param([string]$InputLine)

    Write-DebugLog "IN (PID=$PID): $InputLine"

    $methodData = ConvertFrom-JsonSafe -InputLine $InputLine
    if ($null -eq $methodData) {
        return New-ErrorResponse -Code $script:ErrorCodes.ParseError -Message "$($script:ProviderNamePrefix) Invalid JSON format"
    }

    $methodName = $null
    if ($methodData.PSObject.Properties.Name -contains 'method') {
        $methodName = [string]$methodData.method
    }

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

    try {
        return & $methodEntry.Handler $methodData
    }
    catch {
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Method execution failed: $($_.Exception.Message)"
    }
}

Write-DebugLog "Provider process started. PID=$PID"

while ($true) {
    try {
        $inputLine = [Console]::In.ReadLine()

        if ($null -eq $inputLine) {
            Write-DebugLog 'Input stream closed. Exiting.'
            break
        }

        Write-DebugLog "IN (PID=$PID): $InputLine"

        $response = Process-Method -InputLine ($inputLine.Trim())
        Send-Response -ResponseObject $response
    }
    catch {
        $response = New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to process input: $($_.Exception.Message)"
        Send-Response -ResponseObject $response
    }
}