<#
.SYNOPSIS
    Parallels RAS Custom Provider Sample Script for Oracle Linux Virtualization Manager (OLVM)
.DESCRIPTION
    Implements a Custom Provider Framework (CPF) connector for OLVM (oVirt-engine compatible
    REST API v4). Listens for line-delimited JSON-RPC-style requests on stdin and writes one
    JSON response per line to stdout. Mirrors the structure and hardening of the validated
    Proxmox VE custom provider in this repository (Parallels-RAS-CFP-Proxmox-package2-v2.ps1):
    same method registry / dispatch shape, same JSON-RPC error codes, same disk-persisted,
    mutex-locked task/template state so the connector survives RAS launching a fresh process
    per request, and the same debug logging discipline.
.NOTES
    File Name : OLVM-CustomProvider.ps1
    Target    : QEMU/KVM 7.2.0 hypervisor managed by OLVM 4.5.5-1.67.el8
                (REST API base: https://<engine>/ovirt-engine/api)
.EXAMPLE
    {"method": "provider/connect", "params": {"settings": {"engine_url":"https://olvm.example.com/ovirt-engine/api","username":"admin@internal","password":"XXX","insecure":true,"cluster":"Default"}}}
    {"method": "guests/list"}
    {"method": "guests/get","params":{"id":"101"}}
    {"method": "guests/control","params":{"control":"start","id":"101"}}
    {"method": "guests/convert","params":{"id":"101","is_template":true}}
    {"method": "guests/clone","params":{"id":"101","name":"Clone of 101","snapshot":"RAS_TEMPLATE_VERSION_1"}}
#>

Set-StrictMode -Version Latest

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$WarningPreference = 'SilentlyContinue'
$VerbosePreference = 'SilentlyContinue'
$InformationPreference = 'SilentlyContinue'

if ($Host.Name -notmatch 'ISE') {
    [Console]::InputEncoding = [System.Text.Encoding]::UTF8
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
}

$stdout = [Console]::OpenStandardOutput()
$writer = New-Object System.IO.StreamWriter($stdout, [System.Text.Encoding]::UTF8)
$writer.AutoFlush = $true

$script:ProviderNamePrefix = 'OLVM:'
$script:LogPath = 'C:\CFP Scripts\OLVM-RAS-Provider.log'
$script:TaskStatePath = 'C:\CFP Scripts\OLVM-RAS-TaskState.json'
$script:TemplateFlagStatePath = 'C:\CFP Scripts\OLVM-RAS-TemplateFlagState.json'
$script:OlvmSession = $null
$script:TaskContext = @{}
$script:TaskMaxAgeMinutes = 30

$script:ErrorCodes = @{
    ParseError     = -32700
    MethodNotFound = -32601
    InvalidParams  = -32602
    InternalError  = -32603
}

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

function Write-DebugLog {
    param([string]$Message)

    try {
        $dir = Split-Path -Path $script:LogPath -Parent
        if (-not [string]::IsNullOrWhiteSpace($dir) -and -not (Test-Path -LiteralPath $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
        }
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
        Add-Content -Path $script:LogPath -Value "$timestamp [PID=$PID] $Message" -Encoding UTF8
    }
    catch {
        # Logging must never crash or block the protocol loop.
    }
}

# ---------------------------------------------------------------------------
# Cross-process locking + JSON state files
# (RAS can launch this script as a new process per request; task/clone/template
#  state must therefore live on disk, not just in $script:TaskContext.)
# ---------------------------------------------------------------------------

function Invoke-WithStateLock {
    param(
        [Parameter(Mandatory = $true)][string]$LockName,
        [Parameter(Mandatory = $true)][scriptblock]$Action
    )

    $mutex = New-Object System.Threading.Mutex($false, "Global\$LockName")
    $acquired = $false
    try {
        $acquired = $mutex.WaitOne(30000)
        if (-not $acquired) {
            Write-DebugLog "WARNING: Could not acquire lock [$LockName] within 30s; proceeding without lock."
        }
        return & $Action
    }
    finally {
        if ($acquired) { $mutex.ReleaseMutex() }
        $mutex.Dispose()
    }
}

function Read-JsonStateFile {
    param([Parameter(Mandatory = $true)][string]$Path)

    try {
        if (-not (Test-Path -LiteralPath $Path)) { return @{} }

        $raw = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($raw)) { return @{} }

        $obj = $raw | ConvertFrom-Json -ErrorAction Stop
        $result = @{}
        foreach ($p in $obj.PSObject.Properties) {
            $entry = @{}
            if ($p.Value -is [System.Management.Automation.PSCustomObject]) {
                foreach ($ep in $p.Value.PSObject.Properties) { $entry[$ep.Name] = $ep.Value }
            }
            else {
                $entry = $p.Value
            }
            $result[$p.Name] = $entry
        }
        return $result
    }
    catch {
        Write-DebugLog "Failed to read state file [$Path]: $($_.Exception.Message)"
        return @{}
    }
}

function Save-JsonStateFile {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][hashtable]$State
    )

    try {
        $dir = Split-Path -Path $Path -Parent
        if (-not [string]::IsNullOrWhiteSpace($dir) -and -not (Test-Path -LiteralPath $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
        }
        $json = $State | ConvertTo-Json -Depth 10 -Compress
        Set-Content -LiteralPath $Path -Value $json -Encoding UTF8
    }
    catch {
        Write-DebugLog "Failed to save state file [$Path]: $($_.Exception.Message)"
    }
}

function Set-TaskStateEntry {
    param([Parameter(Mandatory = $true)][string]$TaskId, [Parameter(Mandatory = $true)][hashtable]$Entry)

    Invoke-WithStateLock -LockName 'OlvmRasTaskStateLock' -Action {
        $state = Read-JsonStateFile -Path $script:TaskStatePath
        $state[[string]$TaskId] = $Entry
        Save-JsonStateFile -Path $script:TaskStatePath -State $state
    }
}

function Get-TaskStateEntry {
    param([Parameter(Mandatory = $true)][string]$TaskId)

    if ($script:TaskContext.ContainsKey($TaskId)) {
        return $script:TaskContext[$TaskId]
    }

    $all = Read-JsonStateFile -Path $script:TaskStatePath
    if ($all.ContainsKey([string]$TaskId)) {
        $entry = $all[[string]$TaskId]
        $ctx = @{}
        if ($entry -is [System.Management.Automation.PSCustomObject]) {
            foreach ($p in $entry.PSObject.Properties) { $ctx[$p.Name] = $p.Value }
        }
        elseif ($entry -is [hashtable]) {
            $ctx = $entry
        }
        $script:TaskContext[$TaskId] = $ctx
        return $ctx
    }

    return $null
}

function Remove-TaskStateEntry {
    param([Parameter(Mandatory = $true)][string]$TaskId)

    Invoke-WithStateLock -LockName 'OlvmRasTaskStateLock' -Action {
        $state = Read-JsonStateFile -Path $script:TaskStatePath
        if ($state.ContainsKey([string]$TaskId)) {
            $state.Remove([string]$TaskId)
            Save-JsonStateFile -Path $script:TaskStatePath -State $state
        }
    }
    $script:TaskContext.Remove($TaskId) | Out-Null
}

function New-TaskId {
    return [guid]::NewGuid().ToString()
}

function Test-TaskExpired {
    <#
    .SYNOPSIS
        Pure classification: has a task context outlived MaxAgeMinutes since
        its created_utc timestamp? Kept as its own function (rather than
        inlined in Handle-TaskInfo) so it can be unit-tested without a live
        OLVM connection - this is the hard ceiling that keeps a stuck clone
        (e.g. one that never gets a guest-agent-reported IP) from polling
        'running' forever, unlike the Proxmox v4_ceph provider's clone task
        wait, which has no such ceiling.
    #>
    param([hashtable]$Ctx, [int]$MaxAgeMinutes)

    if ($null -eq $Ctx -or -not $Ctx.ContainsKey('created_utc')) { return $false }

    try {
        $age = (Get-Date).ToUniversalTime() - [datetime]::Parse([string]$Ctx.created_utc).ToUniversalTime()
        return $age.TotalMinutes -gt $MaxAgeMinutes
    }
    catch {
        return $false
    }
}

function Get-TemplateFlagState {
    try {
        if (-not (Test-Path -LiteralPath $script:TemplateFlagStatePath)) { return @{} }
        $raw = Get-Content -LiteralPath $script:TemplateFlagStatePath -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($raw)) { return @{} }
        $obj = $raw | ConvertFrom-Json -ErrorAction Stop
        $state = @{}
        foreach ($p in $obj.PSObject.Properties) { $state[$p.Name] = [bool]$p.Value }
        return $state
    }
    catch {
        Write-DebugLog "Failed to load template flag state: $($_.Exception.Message)"
        return @{}
    }
}

function Set-IsTemplateFlag {
    param([Parameter(Mandatory = $true)][string]$VmId, [Parameter(Mandatory = $true)][bool]$IsTemplate)

    Invoke-WithStateLock -LockName 'OlvmRasTemplateFlagStateLock' -Action {
        $state = Get-TemplateFlagState
        if ($IsTemplate) { $state[[string]$VmId] = $true } else { $state.Remove([string]$VmId) }
        Save-JsonStateFile -Path $script:TemplateFlagStatePath -State $state
    }
}

function Get-IsTemplateFlag {
    param([Parameter(Mandatory = $true)][string]$VmId)

    $state = Get-TemplateFlagState
    return $state.ContainsKey([string]$VmId) -and [bool]$state[[string]$VmId]
}

# ---------------------------------------------------------------------------
# Protocol plumbing
# ---------------------------------------------------------------------------

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
    Write-DebugLog "ERROR [$Code]: $Message"
    return @{ error = @{ code = $Code; message = $Message } }
}

function Test-HasProperty {
    # $obj.PSObject.Properties.Name throws under Set-StrictMode when $obj has
    # zero properties at all (e.g. JSON "{}") - enumerating via Where-Object
    # instead avoids that empty-collection member-access landmine.
    param([object]$InputObject, [Parameter(Mandatory = $true)][string]$Name)

    if ($null -eq $InputObject) { return $false }
    if ($InputObject -isnot [System.Management.Automation.PSCustomObject]) { return $false }

    foreach ($prop in $InputObject.PSObject.Properties) {
        if ($prop.Name -eq $Name) { return $true }
    }
    return $false
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
        $keys = $field -split '\.'
        $value = $Data
        foreach ($key in $keys) {
            if ($null -ne $value -and (Test-HasProperty $value $key)) {
                $value = $value.$key
            }
            else {
                return "$($script:ProviderNamePrefix) Missing field: $field"
            }
        }
    }
    return $null
}

# ---------------------------------------------------------------------------
# TLS / certificate handling for self-signed OLVM CA (Windows PowerShell 5.1)
# ---------------------------------------------------------------------------

function Initialize-CertificateBypass {
    try {
        if ($PSVersionTable.PSEdition -eq 'Core') { return }

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

# ---------------------------------------------------------------------------
# OLVM (oVirt-engine) REST API session + helpers
# ---------------------------------------------------------------------------

function Get-Session {
    if ($null -eq $script:OlvmSession) { throw 'Session not initialized' }
    if ([string]::IsNullOrWhiteSpace($script:OlvmSession.engine_url)) { throw 'Session engine_url missing' }
    return $script:OlvmSession
}

function Get-SsoToken {
    param([switch]$ForceRefresh)

    $session = $script:OlvmSession
    if ($null -eq $session) { throw 'Session not initialized' }

    if (-not $ForceRefresh -and $session.token -and (Get-Date) -lt $session.token_expiry) {
        return $session.token
    }

    $engineBase = $session.engine_url -replace '/ovirt-engine/api/?$', ''
    $ssoUrl = "$engineBase/ovirt-engine/sso/oauth/token"

    $body = @{
        grant_type = 'password'
        scope      = 'ovirt-app-api'
        username   = $session.username
        password   = $session.password
    }

    $irmParams = @{
        Method      = 'Post'
        Uri         = $ssoUrl
        Body        = $body
        ContentType = 'application/x-www-form-urlencoded'
        ErrorAction = 'Stop'
    }
    if ($PSVersionTable.PSEdition -eq 'Core') { $irmParams.SkipCertificateCheck = $true }

    Write-DebugLog "SSO token request to $ssoUrl for user [$($session.username)]"

    try {
        $resp = Invoke-RestMethod @irmParams
    }
    catch {
        Write-DebugLog "SSO token request failed: $($_.Exception.Message)"
        throw "SSO authentication failed: $($_.Exception.Message)"
    }

    if (-not (Test-HasProperty $resp 'access_token') -or [string]::IsNullOrWhiteSpace([string]$resp.access_token)) {
        throw "SSO authentication failed: no access_token in response"
    }

    $session.token = [string]$resp.access_token
    $expiresIn = if ((Test-HasProperty $resp 'expires_in')) { [int]$resp.expires_in } else { 60 }
    $session.token_expiry = (Get-Date).AddSeconds([Math]::Max(10, $expiresIn - 15))
    $script:OlvmSession = $session

    Write-DebugLog "SSO token acquired, expires in ${expiresIn}s"
    return $session.token
}

function Invoke-OlvmRestMethod {
    param(
        [Parameter(Mandatory = $true)][string]$Uri,
        [Parameter(Mandatory = $true)][ValidateSet('GET', 'POST', 'PUT', 'DELETE')][string]$Method,
        [object]$Body = $null,
        [switch]$IsRetry
    )

    $token = Get-SsoToken -ForceRefresh:$IsRetry

    $headers = @{
        Authorization = "Bearer $token"
        Accept        = 'application/json'
        Version       = '4'
    }

    $irmParams = @{
        Uri         = $Uri
        Headers     = $headers
        Method      = $Method
        ErrorAction = 'Stop'
        TimeoutSec  = 60
    }

    if ($PSVersionTable.PSEdition -eq 'Core') {
        $irmParams.SkipCertificateCheck = $true
        $irmParams.SkipHeaderValidation = $true
    }

    if ($null -ne $Body) {
        $irmParams.ContentType = 'application/json'
        $irmParams.Body = if ($Body -is [hashtable] -or $Body -is [System.Collections.IDictionary]) {
            $Body | ConvertTo-Json -Compress -Depth 10
        } else { $Body }
    }
    elseif ($Method -eq 'POST') {
        # oVirt action endpoints require a (possibly empty) JSON body.
        $irmParams.ContentType = 'application/json'
        $irmParams.Body = '{}'
    }

    Write-DebugLog "HTTP $Method $Uri$(if ($irmParams.ContainsKey('Body')) { ' - Body: ' + $irmParams.Body })"

    try {
        return Invoke-RestMethod @irmParams
    }
    catch {
        $statusCode = $null
        try { $statusCode = [int]$_.Exception.Response.StatusCode } catch { }

        Write-DebugLog "HTTP failure ($statusCode): $($_.Exception.Message)"

        if ($statusCode -eq 401 -and -not $IsRetry) {
            Write-DebugLog "Token appears expired/invalid, forcing refresh and retrying once."
            return Invoke-OlvmRestMethod -Uri $Uri -Method $Method -Body $Body -IsRetry
        }

        throw
    }
}

function Invoke-OlvmApi {
    param(
        [Parameter(Mandatory = $true)][ValidateSet('GET', 'POST', 'PUT', 'DELETE')][string]$Method,
        [Parameter(Mandatory = $true)][string]$Path,
        [object]$Body = $null
    )

    $session = Get-Session
    $base = $session.engine_url.TrimEnd('/')
    $uri = $base + '/' + $Path.TrimStart('/')

    return Invoke-OlvmRestMethod -Uri $uri -Method $Method -Body $Body
}

function Test-OlvmObjectExists {
    param([Parameter(Mandatory = $true)][scriptblock]$Probe)

    try {
        & $Probe | Out-Null
        return $true
    }
    catch {
        $statusCode = $null
        try { $statusCode = [int]$_.Exception.Response.StatusCode } catch { }
        if ($statusCode -eq 404) { return $false }
        throw
    }
}

# ---------------------------------------------------------------------------
# Entity fetchers
# ---------------------------------------------------------------------------

function Get-OlvmVms {
    $resp = Invoke-OlvmApi -Method GET -Path 'vms'
    if ($null -eq $resp -or -not ((Test-HasProperty $resp 'vm'))) { return @() }
    return @($resp.vm)
}

function Get-OlvmTemplates {
    $resp = Invoke-OlvmApi -Method GET -Path 'templates'
    if ($null -eq $resp -or -not ((Test-HasProperty $resp 'template'))) { return @() }
    return @($resp.template) | Where-Object { $_.name -ne 'Blank' }
}

function Get-OlvmVm {
    param([Parameter(Mandatory = $true)][string]$VmId)
    return Invoke-OlvmApi -Method GET -Path "vms/$VmId"
}

function Get-OlvmTemplate {
    param([Parameter(Mandatory = $true)][string]$TemplateId)
    return Invoke-OlvmApi -Method GET -Path "templates/$TemplateId"
}

function Get-OlvmVmNics {
    param([Parameter(Mandatory = $true)][string]$VmId)
    try {
        $resp = Invoke-OlvmApi -Method GET -Path "vms/$VmId/nics"
        if ((Test-HasProperty $resp 'nic')) { return @($resp.nic) }
        return @()
    }
    catch {
        Write-DebugLog "NIC lookup failed for VM [$VmId]: $($_.Exception.Message)"
        return @()
    }
}

function Get-OlvmVmNetworkData {
    param([Parameter(Mandatory = $true)][string]$VmId)

    $ipv4Set = New-Object 'System.Collections.Generic.HashSet[string]'
    $macSet = New-Object 'System.Collections.Generic.HashSet[string]'

    foreach ($nic in (Get-OlvmVmNics -VmId $VmId)) {
        if ((Test-HasProperty $nic 'mac') -and (Test-HasProperty $nic.mac 'address')) {
            [void]$macSet.Add(([string]$nic.mac.address).ToUpperInvariant())
        }
    }

    try {
        $reported = Invoke-OlvmApi -Method GET -Path "vms/$VmId/reporteddevices"
        if ((Test-HasProperty $reported 'reported_device')) {
            foreach ($dev in @($reported.reported_device)) {
                if ((Test-HasProperty $dev 'ips') -and (Test-HasProperty $dev.ips 'ip')) {
                    foreach ($ip in @($dev.ips.ip)) {
                        $addr = [string]$ip.address
                        $version = if ((Test-HasProperty $ip 'version')) { [string]$ip.version } else { 'v4' }
                        if (($version -eq 'v4' -or $version -eq 'V4') -and
                            -not [string]::IsNullOrWhiteSpace($addr) -and
                            $addr -ne '127.0.0.1' -and
                            $addr -notmatch '^169\.254\.') {
                            [void]$ipv4Set.Add($addr)
                        }
                    }
                }
            }
        }
    }
    catch {
        Write-DebugLog "Reported-devices (guest agent) lookup failed for VM [$VmId]: $($_.Exception.Message)"
    }

    return @{
        IPv4Addresses = @($ipv4Set | Select-Object -First 3)
        MacAddresses  = @($macSet | Select-Object -First 3)
    }
}

function Get-OlvmVmSnapshots {
    param([Parameter(Mandatory = $true)][string]$VmId)
    $resp = Invoke-OlvmApi -Method GET -Path "vms/$VmId/snapshots"
    if ((Test-HasProperty $resp 'snapshot')) { return @($resp.snapshot) }
    return @()
}

function Find-OlvmVmSnapshot {
    param([Parameter(Mandatory = $true)][string]$VmId, [Parameter(Mandatory = $true)][string]$SnapshotName)

    foreach ($snap in (Get-OlvmVmSnapshots -VmId $VmId)) {
        if ((Test-HasProperty $snap 'description') -and [string]$snap.description -eq $SnapshotName) {
            return $snap
        }
    }
    return $null
}

# ---------------------------------------------------------------------------
# State mapping
# ---------------------------------------------------------------------------

function Map-OlvmStateToRasState {
    param([string]$Status)

    $normalized = if ($null -ne $Status) { $Status.ToString().Trim().ToLowerInvariant() } else { 'unknown' }

    switch ($normalized) {
        'up' { return 'powered_on' }
        'reboot_in_progress' { return 'powered_on' }
        'migrating' { return 'powered_on' }
        'not_responding' { return 'powered_on' }
        'powering_up' { return 'powering_on' }
        'wait_for_launch' { return 'powering_on' }
        'image_locked' { return 'powering_on' }
        'down' { return 'powered_off' }
        'powering_down' { return 'powering_off' }
        'paused' { return 'suspended' }
        'suspended' { return 'suspended' }
        default { return 'powered_off' }
    }
}

function Get-OlvmVmOsType {
    param([object]$Vm)

    if ((Test-HasProperty $Vm 'guest_operating_system')) {
        $gos = $Vm.guest_operating_system
        if ((Test-HasProperty $gos 'distribution') -and -not [string]::IsNullOrWhiteSpace([string]$gos.distribution)) {
            return [string]$gos.distribution
        }
    }

    if ((Test-HasProperty $Vm 'os') -and (Test-HasProperty $Vm.os 'type')) {
        if (-not [string]::IsNullOrWhiteSpace([string]$Vm.os.type)) { return [string]$Vm.os.type }
    }

    return 'unknown'
}

function ConvertTo-RasGuestObject {
    param([Parameter(Mandatory = $true)][string]$VmId)

    $vm = Get-OlvmVm -VmId $VmId

    $network = $null
    try {
        $network = Get-OlvmVmNetworkData -VmId $VmId
    }
    catch {
        Write-DebugLog "Network lookup failed for VM [$VmId]: $($_.Exception.Message)"
        $network = @{ IPv4Addresses = @(); MacAddresses = @() }
    }

    $rawState = if ((Test-HasProperty $vm 'status')) { [string]$vm.status } else { 'unknown' }
    $name = if ((Test-HasProperty $vm 'name')) { [string]$vm.name } else { "VM-$VmId" }
    $isTemplate = Get-IsTemplateFlag -VmId $VmId

    $memoryMb = $null
    if ((Test-HasProperty $vm 'memory')) {
        try { $memoryMb = [math]::Round([double]$vm.memory / 1MB) } catch { }
    }

    $guestObject = @{
        id            = [string]$VmId
        name          = $name
        provider      = 'OLVM'
        state         = (Map-OlvmStateToRasState -Status $rawState)
        power_state   = $rawState
        host_os       = (Get-OlvmVmOsType -Vm $vm)
        ip            = $(if ($network.IPv4Addresses.Count -gt 0) { $network.IPv4Addresses[0] } else { $null })
        ip_addresses  = @($network.IPv4Addresses)
        mac_addresses = @($network.MacAddresses)
        is_template   = $isTemplate
        memory_mb     = $memoryMb
        type          = 'Virtual Machine'
    }

    Write-DebugLog ("GUEST VMID={0}; Name={1}; State={2}; Template={3}; IPs={4}" -f `
            $guestObject.id, $guestObject.name, $guestObject.state, $guestObject.is_template, ($guestObject.ip_addresses -join ','))

    return $guestObject
}

function ConvertTo-RasTemplateObject {
    param([Parameter(Mandatory = $true)][object]$Tpl)

    return @{
        id            = [string]$Tpl.id
        name          = [string]$Tpl.name
        provider      = 'OLVM'
        state         = 'powered_off'
        power_state   = 'down'
        host_os       = (Get-OlvmVmOsType -Vm $Tpl)
        ip            = $null
        ip_addresses  = @()
        mac_addresses = @()
        is_template   = $true
        type          = 'Template'
    }
}

function Get-RasObjectForAnyId {
    param([Parameter(Mandatory = $true)][string]$Id)

    if (Test-OlvmObjectExists -Probe { Get-OlvmVm -VmId $Id }) {
        return ConvertTo-RasGuestObject -VmId $Id
    }

    $tpl = Get-OlvmTemplate -TemplateId $Id
    return ConvertTo-RasTemplateObject -Tpl $tpl
}

# ---------------------------------------------------------------------------
# Clone-aware guests/get flow: newly cloned VMs are powered off by default in
# OLVM; RAS host pools expect a ready, running, IP-addressed guest once the
# clone task completes. Auto-start the clone and hold it in "powering_on"
# until it reports an IP, then latch creation_completed so a later, deliberate
# power-off (maintenance mode, idle desktop power management) is never
# silently undone.
# ---------------------------------------------------------------------------

function Get-TrackedCloneContext {
    param([Parameter(Mandatory = $true)][string]$VmId)

    $all = Read-JsonStateFile -Path $script:TaskStatePath
    foreach ($key in $all.Keys) {
        $entry = $all[$key]
        $ctx = @{}
        if ($entry -is [System.Management.Automation.PSCustomObject]) {
            foreach ($p in $entry.PSObject.Properties) { $ctx[$p.Name] = $p.Value }
        }
        elseif ($entry -is [hashtable]) { $ctx = $entry }

        if ($ctx.ContainsKey('type') -and [string]$ctx.type -eq 'clone' -and
            $ctx.ContainsKey('vm_id') -and [string]$ctx.vm_id -eq [string]$VmId -and
            -not ([bool]$ctx.creation_completed)) {
            return @{ task_id = [string]$key; context = $ctx }
        }
    }
    return $null
}

function Start-OlvmVmIfNeeded {
    param([Parameter(Mandatory = $true)][string]$VmId)

    try {
        Invoke-OlvmApi -Method POST -Path "vms/$VmId/start" -Body @{} | Out-Null
        Write-DebugLog "Issued start for VM [$VmId]"
        return $true
    }
    catch {
        Write-DebugLog "Start attempt for VM [$VmId] failed (may already be starting): $($_.Exception.Message)"
        return $false
    }
}

function Get-RasGuestObjectForCloneAwareFlow {
    param([Parameter(Mandatory = $true)][string]$VmId)

    $guest = ConvertTo-RasGuestObject -VmId $VmId
    $tracked = Get-TrackedCloneContext -VmId $VmId

    if ($null -eq $tracked) { return $guest }

    $ctx = $tracked.context
    if (-not $ctx.ContainsKey('start_issued')) { $ctx.start_issued = $false }
    if (-not $ctx.ContainsKey('creation_completed')) { $ctx.creation_completed = $false }

    if ($guest.state -eq 'powered_off') {
        if (-not [bool]$ctx.start_issued) {
            $started = Start-OlvmVmIfNeeded -VmId $VmId
            $ctx.start_issued = $started
        }
        $guest.state = 'powering_on'
        $guest.power_state = 'starting'
    }
    elseif ($guest.state -eq 'powered_on') {
        if (@($guest.ip_addresses).Count -gt 0) {
            $ctx.creation_completed = $true
            Write-DebugLog "Clone-aware get: VM [$VmId] powered on with IP(s) [$($guest.ip_addresses -join ',')]."
        }
        else {
            $guest.state = 'powering_on'
            $guest.power_state = 'starting'
        }
    }

    Set-TaskStateEntry -TaskId $tracked.task_id -Entry $ctx

    return $guest
}

# ---------------------------------------------------------------------------
# Control action mapping
# ---------------------------------------------------------------------------

function Get-ControlAction {
    param([Parameter(Mandatory = $true)][string]$Control)

    switch ($Control.Trim().ToLowerInvariant()) {
        'start' { return 'start' }
        'stop' { return 'shutdown' }
        'shutdown' { return 'shutdown' }
        'reset' { return 'reset' }
        'restart' { return 'reboot' }
        'reboot' { return 'reboot' }
        'suspend' { return 'suspend' }
        'delete' { return 'remove' }
        default { return $null }
    }
}

# ---------------------------------------------------------------------------
# Method handlers
# ---------------------------------------------------------------------------

function Handle-Initialize {
    return @{
        result = @{
            version      = '1.0.0'
            capabilities = @{
                can_suspend_guests    = $true
                guests_polling_rate   = 15
                tasks_polling_rate    = 5
                tasks_polling_retries = 60
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

    $engineUrl = if (Test-HasProperty $settings 'engine_url') { [string]$settings.engine_url } else { $null }
    $username = if (Test-HasProperty $settings 'username') { [string]$settings.username } else { $null }
    $password = if (Test-HasProperty $settings 'password') { [string]$settings.password } else { $null }
    $insecure = if (Test-HasProperty $settings 'insecure') { [bool]$settings.insecure } else { $false }
    $cluster = if (Test-HasProperty $settings 'cluster') { [string]$settings.cluster } else { $null }
    $storageDomain = if (Test-HasProperty $settings 'storage_domain') { [string]$settings.storage_domain } else { $null }

    if ([string]::IsNullOrWhiteSpace($engineUrl) -or
        [string]::IsNullOrWhiteSpace($username) -or
        [string]::IsNullOrWhiteSpace($password)) {
        return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Invalid connection parameters: engine_url, username and password are required"
    }

    try {
        if ($insecure) { Initialize-CertificateBypass }

        $script:OlvmSession = @{
            engine_url     = $engineUrl.TrimEnd('/')
            username       = $username
            password       = $password
            insecure       = $insecure
            cluster        = $cluster
            storage_domain = $storageDomain
            token          = $null
            token_expiry   = [datetime]::MinValue
        }

        Get-SsoToken | Out-Null
        $resp = Invoke-OlvmApi -Method GET -Path 'vms?max=1'
        if ($null -eq $resp) { throw 'Empty response probing /vms' }

        Write-DebugLog "Connected successfully to $engineUrl as $username"
        return @{ result = @{ message = "$($script:ProviderNamePrefix) Connected successfully to $engineUrl" } }
    }
    catch {
        $script:OlvmSession = $null
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to connect to OLVM: $($_.Exception.Message)"
    }
}

function Handle-Disconnect {
    try {
        $engineUrl = $null
        if ($null -ne $script:OlvmSession) { $engineUrl = $script:OlvmSession.engine_url }

        $script:OlvmSession = $null
        $script:TaskContext = @{}

        return @{ result = @{ message = "Session cleared for OLVM $engineUrl" } }
    }
    catch {
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to clear session: $($_.Exception.Message)"
    }
}

function Get-AllGuestAndTemplateIds {
    $ids = @()
    foreach ($vm in (Get-OlvmVms)) {
        if ((Test-HasProperty $vm 'id')) { $ids += [string]$vm.id }
    }
    foreach ($tpl in (Get-OlvmTemplates)) {
        if ((Test-HasProperty $tpl 'id')) { $ids += [string]$tpl.id }
    }
    return $ids
}

function Handle-GuestList {
    try {
        $ids = Get-AllGuestAndTemplateIds
        Write-DebugLog "guests/list returning $($ids.Count) id(s)"
        return @{ result = @{ guests = @($ids) } }
    }
    catch {
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to retrieve guest list: $($_.Exception.Message)"
    }
}

function Get-SingleOrMultiResult {
    param([object]$Params, [scriptblock]$Resolver)

    $ids = @($Params.id)

    if ($ids.Count -eq 1) {
        return @{ result = (& $Resolver ([string]$ids[0])) }
    }

    $resultMap = @{}
    foreach ($id in $ids) {
        $vmId = [string]$id
        try {
            $resultMap[$vmId] = & $Resolver $vmId
        }
        catch {
            $resultMap[$vmId] = @{
                id = $vmId; name = $null; provider = 'OLVM'; state = 'powered_off'; power_state = 'unknown'
                host_os = 'unknown'; ip = $null; ip_addresses = @(); mac_addresses = @(); is_template = $false
                type = 'Virtual Machine'
            }
            Write-DebugLog "Get failed for id [$vmId]: $($_.Exception.Message)"
        }
    }
    return @{ result = $resultMap }
}

function Handle-GuestGet {
    param([object]$Params)

    if ($null -eq $Params -or $null -eq $Params.id) {
        return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Invalid or missing guest id"
    }

    try {
        return Get-SingleOrMultiResult -Params $Params -Resolver {
            param($vmId)
            if (Test-OlvmObjectExists -Probe { Get-OlvmVm -VmId $vmId }) {
                return Get-RasGuestObjectForCloneAwareFlow -VmId $vmId
            }
            return ConvertTo-RasTemplateObject -Tpl (Get-OlvmTemplate -TemplateId $vmId)
        }
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

        if ($action -eq 'remove') {
            Invoke-OlvmApi -Method DELETE -Path "vms/$vmId" | Out-Null
        }
        else {
            Invoke-OlvmApi -Method POST -Path "vms/$vmId/$action" -Body @{} | Out-Null
        }

        return @{
            result = @{
                id      = $vmId
                action  = $action
                message = "$($script:ProviderNamePrefix) Guest control [$requestedControl] submitted successfully"
            }
        }
    }
    catch {
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to control guest [$($Params.id)]: $($_.Exception.Message)"
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

        # Template versions are represented as named snapshots directly on the source
        # VM (see guests/snapshots/*), not as a real OLVM template object - an OLVM/
        # oVirt template can never be started, edited, or snapshotted again, which is
        # incompatible with the maintenance-mode edit/snapshot cycle CPF requires.
        # guests/convert therefore takes no OLVM-side action; it only persists the
        # RAS-facing is_template label so guests/get reports it back.
        Set-IsTemplateFlag -VmId $vmId -IsTemplate $isTemplate

        $taskId = New-TaskId
        Set-TaskStateEntry -TaskId $taskId -Entry @{ type = 'noop'; created_utc = (Get-Date).ToUniversalTime().ToString('o') }

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
        $sourceId = [string]$Params.id
        $cloneName = [string]$Params.name
        $snapshotName = if (((Test-HasProperty $Params 'snapshot')) -and -not [string]::IsNullOrWhiteSpace([string]$Params.snapshot)) { [string]$Params.snapshot } else { $null }

        $session = Get-Session
        $body = @{ name = $cloneName }
        if (-not [string]::IsNullOrWhiteSpace($session.cluster)) {
            $body.cluster = @{ name = $session.cluster }
        }

        $isTemplateSource = Test-OlvmObjectExists -Probe { Get-OlvmTemplate -TemplateId $sourceId } -and
                            -not (Test-OlvmObjectExists -Probe { Get-OlvmVm -VmId $sourceId })

        if ($isTemplateSource) {
            $body.template = @{ id = $sourceId }
        }
        else {
            $body.vm = @{ id = $sourceId }
            $body.clone = $true

            if ($snapshotName) {
                $snapObj = Find-OlvmVmSnapshot -VmId $sourceId -SnapshotName $snapshotName
                if ($null -eq $snapObj) {
                    throw "Snapshot '$snapshotName' not found on source guest [$sourceId]"
                }
                $body.snapshots = @{ snapshot = @(@{ id = $snapObj.id }) }
            }
        }

        Write-DebugLog "Cloning source [$sourceId] -> name [$cloneName] snapshot [$snapshotName]"
        $newVm = Invoke-OlvmApi -Method POST -Path 'vms' -Body $body

        if ($null -eq $newVm -or -not ((Test-HasProperty $newVm 'id'))) {
            throw "Clone request did not return a new VM id. Response: $(($newVm | ConvertTo-Json -Compress))"
        }

        $newVmId = [string]$newVm.id
        $taskId = New-TaskId

        Set-TaskStateEntry -TaskId $taskId -Entry @{
            type               = 'clone'
            vm_id              = $newVmId
            source_id          = $sourceId
            name               = $cloneName
            start_issued       = $false
            creation_completed = $false
            created_utc        = (Get-Date).ToUniversalTime().ToString('o')
        }

        Write-DebugLog "Clone submitted: source [$sourceId] -> new VM [$newVmId], task [$taskId]"

        return @{ result = @{ task_id = $taskId; clone_id = $newVmId } }
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

        $existing = Find-OlvmVmSnapshot -VmId $vmId -SnapshotName $snapshotName
        if ($null -ne $existing) {
            # A prior attempt at this same step (RAS retrying after a later step
            # failed) can leave a stale snapshot with this name behind. OLVM refuses
            # a duplicate description-named snapshot for the same purpose here, so
            # replace it rather than fail - the fresh snapshot must reflect current
            # VM state, not whatever an earlier incomplete attempt captured.
            Write-DebugLog "Snapshot [$snapshotName] already exists on guest [$vmId]; replacing it before creating a fresh one."
            Invoke-OlvmApi -Method DELETE -Path "vms/$vmId/snapshots/$($existing.id)" | Out-Null

            $deadline = (Get-Date).AddSeconds(120)
            while ((Get-Date) -lt $deadline) {
                if (-not (Test-OlvmObjectExists -Probe { Invoke-OlvmApi -Method GET -Path "vms/$vmId/snapshots/$($existing.id)" })) { break }
                Start-Sleep -Seconds 3
            }
        }

        $snap = Invoke-OlvmApi -Method POST -Path "vms/$vmId/snapshots" -Body @{ description = $snapshotName }
        if ($null -eq $snap -or -not ((Test-HasProperty $snap 'id'))) {
            throw "Snapshot creation did not return a snapshot id. Response: $(($snap | ConvertTo-Json -Compress))"
        }

        $taskId = New-TaskId
        Set-TaskStateEntry -TaskId $taskId -Entry @{
            type          = 'snapshot_create'
            vm_id         = $vmId
            snapshot_id   = [string]$snap.id
            snapshot_name = $snapshotName
            created_utc   = (Get-Date).ToUniversalTime().ToString('o')
        }

        Write-DebugLog "Snapshot [$snapshotName] creation submitted for guest [$vmId], snapshot id [$($snap.id)], task [$taskId]"
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

        $snap = Find-OlvmVmSnapshot -VmId $vmId -SnapshotName $snapshotName
        if ($null -eq $snap) {
            # Idempotent: nothing to delete is success, not failure.
            $taskId = New-TaskId
            Set-TaskStateEntry -TaskId $taskId -Entry @{ type = 'noop'; created_utc = (Get-Date).ToUniversalTime().ToString('o') }
            Write-DebugLog "Snapshot [$snapshotName] not found on guest [$vmId]; treating delete as already complete."
            return @{ result = @{ task_id = $taskId } }
        }

        Invoke-OlvmApi -Method DELETE -Path "vms/$vmId/snapshots/$($snap.id)" | Out-Null

        $taskId = New-TaskId
        Set-TaskStateEntry -TaskId $taskId -Entry @{
            type        = 'snapshot_delete'
            vm_id       = $vmId
            snapshot_id = [string]$snap.id
            created_utc = (Get-Date).ToUniversalTime().ToString('o')
        }

        Write-DebugLog "Snapshot [$snapshotName] deletion submitted for guest [$vmId], task [$taskId]"
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
        $exists = $null -ne (Find-OlvmVmSnapshot -VmId ([string]$Params.id) -SnapshotName ([string]$Params.name))
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

        $snap = Find-OlvmVmSnapshot -VmId $vmId -SnapshotName $snapshotName
        if ($null -eq $snap) {
            throw "Snapshot '$snapshotName' not found on guest [$vmId]"
        }

        # OLVM/oVirt requires the VM to be down before a snapshot restore.
        $vm = Get-OlvmVm -VmId $vmId
        $rasState = Map-OlvmStateToRasState -Status ([string]$vm.status)
        if ($rasState -ne 'powered_off') {
            Write-DebugLog "Guest [$vmId] is [$rasState]; stopping before restoring snapshot [$snapshotName]."
            try { Invoke-OlvmApi -Method POST -Path "vms/$vmId/stop" -Body @{} | Out-Null } catch {
                Write-DebugLog "Stop before revert failed (may already be stopping): $($_.Exception.Message)"
            }

            $deadline = (Get-Date).AddSeconds(120)
            while ((Get-Date) -lt $deadline) {
                $vm = Get-OlvmVm -VmId $vmId
                if ((Map-OlvmStateToRasState -Status ([string]$vm.status)) -eq 'powered_off') { break }
                Start-Sleep -Seconds 3
            }
        }

        Invoke-OlvmApi -Method POST -Path "vms/$vmId/snapshots/$($snap.id)/restore" -Body @{} | Out-Null

        $taskId = New-TaskId
        Set-TaskStateEntry -TaskId $taskId -Entry @{
            type        = 'snapshot_revert'
            vm_id       = $vmId
            snapshot_id = [string]$snap.id
            created_utc = (Get-Date).ToUniversalTime().ToString('o')
        }

        Write-DebugLog "Snapshot [$snapshotName] revert submitted for guest [$vmId], task [$taskId]"
        return @{ result = @{ task_id = $taskId } }
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

    $taskId = [string]$Params.id

    try {
        $ctx = Get-TaskStateEntry -TaskId $taskId
        if ($null -eq $ctx) {
            return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Unknown task id: $taskId"
        }

        if (Test-TaskExpired -Ctx $ctx -MaxAgeMinutes $script:TaskMaxAgeMinutes) {
            Write-DebugLog "Task [$taskId] exceeded max age of $($script:TaskMaxAgeMinutes) minutes; failing."
            Remove-TaskStateEntry -TaskId $taskId
            return @{ result = @{ state = 'failed'; error = @{ code = 1; message = 'Task timed out' } } }
        }

        $type = if ($ctx.ContainsKey('type')) { [string]$ctx.type } else { 'noop' }

        switch ($type) {
            'noop' {
                Remove-TaskStateEntry -TaskId $taskId
                return @{ result = @{ state = 'completed'; output = @{} } }
            }

            'snapshot_create' {
                $vmId = [string]$ctx.vm_id
                $snapId = [string]$ctx.snapshot_id
                try {
                    $snap = Invoke-OlvmApi -Method GET -Path "vms/$vmId/snapshots/$snapId"
                }
                catch {
                    Write-DebugLog "Snapshot status lookup failed for task [$taskId]: $($_.Exception.Message)"
                    return @{ result = @{ state = 'running' } }
                }

                $status = if ((Test-HasProperty $snap 'snapshot_status')) { [string]$snap.snapshot_status } else { 'unknown' }
                if ($status -eq 'ok') {
                    Remove-TaskStateEntry -TaskId $taskId
                    return @{ result = @{ state = 'completed'; output = @{ snapshot_id = $snapId } } }
                }
                if ($status -eq 'locked' -or $status -eq 'in_preview') {
                    return @{ result = @{ state = 'running' } }
                }
                Remove-TaskStateEntry -TaskId $taskId
                return @{ result = @{ state = 'failed'; error = @{ code = 1; message = "Snapshot status [$status]" } } }
            }

            'snapshot_delete' {
                $vmId = [string]$ctx.vm_id
                $snapId = [string]$ctx.snapshot_id
                $exists = Test-OlvmObjectExists -Probe { Invoke-OlvmApi -Method GET -Path "vms/$vmId/snapshots/$snapId" }
                if (-not $exists) {
                    Remove-TaskStateEntry -TaskId $taskId
                    return @{ result = @{ state = 'completed'; output = @{} } }
                }
                return @{ result = @{ state = 'running' } }
            }

            'snapshot_revert' {
                $vmId = [string]$ctx.vm_id
                try {
                    $vm = Get-OlvmVm -VmId $vmId
                }
                catch {
                    return @{ result = @{ state = 'running' } }
                }
                $rasState = Map-OlvmStateToRasState -Status ([string]$vm.status)
                if ($rasState -eq 'powered_off') {
                    Remove-TaskStateEntry -TaskId $taskId
                    return @{ result = @{ state = 'completed'; output = @{ id = $vmId } } }
                }
                return @{ result = @{ state = 'running' } }
            }

            'clone' {
                $vmId = [string]$ctx.vm_id

                if (-not (Test-OlvmObjectExists -Probe { Get-OlvmVm -VmId $vmId })) {
                    return @{ result = @{ state = 'running' } }
                }

                $guest = ConvertTo-RasGuestObject -VmId $vmId

                if ($guest.state -eq 'powered_off') {
                    if (-not [bool]$ctx.start_issued) {
                        $ctx.start_issued = Start-OlvmVmIfNeeded -VmId $vmId
                        Set-TaskStateEntry -TaskId $taskId -Entry $ctx
                    }
                    return @{ result = @{ state = 'running' } }
                }

                if ($guest.state -eq 'powering_on') {
                    return @{ result = @{ state = 'running' } }
                }

                if ($guest.state -ne 'powered_on') {
                    return @{ result = @{ state = 'running' } }
                }

                if (@($guest.ip_addresses).Count -eq 0) {
                    return @{ result = @{ state = 'running' } }
                }

                $ctx.creation_completed = $true
                Set-TaskStateEntry -TaskId $taskId -Entry $ctx

                Write-DebugLog "Clone task [$taskId]: VM [$vmId] powered on with IP(s) [$($guest.ip_addresses -join ',')]. Completing."
                return @{ result = @{ state = 'completed'; output = @{ clone_id = $vmId } } }
            }

            default {
                Remove-TaskStateEntry -TaskId $taskId
                return @{ result = @{ state = 'completed'; output = @{} } }
            }
        }
    }
    catch {
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to retrieve task info: $($_.Exception.Message)"
    }
}

# ---------------------------------------------------------------------------
# Method registry + dispatch loop
# ---------------------------------------------------------------------------

$script:MethodRegistry = @{
    'provider/initialize' = @{ Handler = { param($data) Handle-Initialize }; RequiredFields = @() }
    'provider/connect'    = @{ Handler = { param($data) Handle-Connect -Params $data.params }; RequiredFields = @('params.settings') }
    'provider/disconnect' = @{ Handler = { param($data) Handle-Disconnect }; RequiredFields = @() }

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

    Write-DebugLog "IN: $InputLine"

    $methodData = ConvertFrom-JsonSafe -InputLine $InputLine
    if ($null -eq $methodData) {
        return New-ErrorResponse -Code $script:ErrorCodes.ParseError -Message "$($script:ProviderNamePrefix) Invalid JSON format"
    }

    $methodName = $null
    if ((Test-HasProperty $methodData 'method')) { $methodName = [string]$methodData.method }

    if ([string]::IsNullOrWhiteSpace($methodName)) {
        return New-ErrorResponse -Code $script:ErrorCodes.MethodNotFound -Message "$($script:ProviderNamePrefix) Missing method name"
    }

    $lookupName = $methodName.Trim().ToLowerInvariant()
    if (-not $script:MethodRegistry.ContainsKey($lookupName)) {
        return New-ErrorResponse -Code $script:ErrorCodes.MethodNotFound -Message "$($script:ProviderNamePrefix) Unknown method: $methodName"
    }

    $methodEntry = $script:MethodRegistry[$lookupName]

    # provider/connect must succeed before any other method touches Get-Session.
    if ($lookupName -notin @('provider/initialize', 'provider/connect') -and $null -eq $script:OlvmSession) {
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Not connected: call provider/connect first"
    }

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

        if ([string]::IsNullOrWhiteSpace($inputLine)) { continue }

        $response = Process-Method -InputLine ($inputLine.Trim())
        Send-Response -ResponseObject $response
    }
    catch {
        $response = New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to process input: $($_.Exception.Message)"
        Send-Response -ResponseObject $response
    }
}
