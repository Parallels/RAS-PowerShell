<#  
.SYNOPSIS  
    Parallels RAS Custom Provider Sample Script for Proxmox VE
.DESCRIPTION  
    This script implements a custom provider for Parallels RAS to integrate with Proxmox VE
    hypervisor. It listens for JSON-RPC requests on standard input, processes them according
    to the defined methods, and returns responses on standard output. The provider supports
    connecting to Proxmox using API tokens, listing VMs, retrieving VM information, and
    controlling VM power state. 
    This script reuires PowerShell 7 or later for best compatibility.
.NOTES  
    File Name  : Parallels-RAS-CFP-Proxmox.ps1
    Author     : www.parallels.com
.EXAMPLE
    .\Parallels-RAS-CFP-Proxmoxl.ps1
    Sample request json

    {"method": "provider/connect", "params" : { "settings": {"host":"proxmox.example.com","username":"root@pam","token_name":"automation","token_secret":"7819cf5ca94a30ad154"}}}
    {"method": "guests/list"}
    {"method":"guests/control","params":{"control":"start","id":"101"}}
    {"method":"guests/get","params":{"id":"101"}}
    {"method":"guests/get","params":{"id":["101","102"]}}
#>

if ($Host.Name -notmatch "ISE") {
    [Console]::InputEncoding  = [System.Text.Encoding]::UTF8
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
}

$stdout = [Console]::OpenStandardOutput()
$writer = New-Object System.IO.StreamWriter($stdout, [System.Text.Encoding]::UTF8)
$writer.AutoFlush = $true

$providerNamePrefix = "Proxmox:"

$ErrorCodes = @{
    ParseError     = -32700
    MethodNotFound = -32601
    InvalidParams  = -32602
    InternalError  = -32603
}

# Optional diagnostics
$EnableDebugLog = $false
$DebugLogPath   = "C:\Windows\Temp\Proxmox-RAS-debug.log"

function Write-DebugLog {
    param([string]$Message)

    if (-not $EnableDebugLog) { return }

    try {
        $line = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message
        Add-Content -Path $DebugLogPath -Value $line -Encoding UTF8
    }
    catch {
        # swallow logging errors
    }
}

Write-DebugLog "Script started."
Write-DebugLog ("PSVersion: {0}" -f $PSVersionTable.PSVersion)
Write-DebugLog ("PSEdition: {0}" -f $PSVersionTable.PSEdition)
Write-DebugLog ("Host: {0}" -f $Host.Name)
Write-DebugLog ("Is64BitProcess: {0}" -f [Environment]::Is64BitProcess)
Write-DebugLog ("User: {0}" -f [System.Security.Principal.WindowsIdentity]::GetCurrent().Name)
Write-DebugLog ("TEMP: {0}" -f $env:TEMP)
Write-DebugLog ("TMP: {0}" -f $env:TMP)

# Predefined method registry with required parameters
$MethodRegistry = @{
    "provider/initialize" = @{ Handler = { param($data) Handle-Initialize }; RequiredFields = @() }
    "provider/connect"    = @{ Handler = { param($data) Handle-Connect $data.params }; RequiredFields = @("params.settings") }
    "provider/disconnect" = @{ Handler = { param($data) Handle-Disconnect }; RequiredFields = @() }
    "guests/list"         = @{ Handler = { param($data) Handle-GuestList }; RequiredFields = @() }
    "guests/get"          = @{ Handler = { param($data) Handle-GuestInfo $data.params }; RequiredFields = @("params.id") }
    "guests/control"      = @{ Handler = { param($data) Handle-GuestControl $data.params }; RequiredFields = @("params.id", "params.control") }
}

function Send-Response {
    param([object]$responseObj)

    $responseJson = $responseObj | ConvertTo-Json -Compress -Depth 10
    $writer.WriteLine($responseJson)
    Write-DebugLog ("Response: {0}" -f $responseJson)
}

function ConvertTo-JsonSafe {
    param([string]$inputLine)

    try {
        return $inputLine | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        Write-DebugLog ("JSON parse failure: {0}" -f $_.Exception.Message)
        return $null
    }
}

function Validate-MethodInput {
    param(
        [object]$data,
        [array]$requiredFields
    )

    foreach ($field in $requiredFields) {
        $keys  = $field -split '\.'
        $value = $data

        foreach ($key in $keys) {
            if ($value -and $value.PSObject.Properties[$key]) {
                $value = $value.$key
            }
            else {
                return @{
                    error = @{
                        code    = $ErrorCodes.InvalidParams
                        message = "$providerNamePrefix Missing field: $field"
                    }
                }
            }
        }
    }

    return $null
}

function Process-Method {
    param([string]$inputLine)

    $methodData = ConvertTo-JsonSafe $inputLine
    if (-not $methodData) {
        return @{
            error = @{
                code    = $ErrorCodes.ParseError
                message = "$providerNamePrefix Invalid JSON format"
            }
        }
    }

    $methodName = $methodData.method
    if (-not $methodName) {
        return @{
            error = @{
                code    = $ErrorCodes.MethodNotFound
                message = "$providerNamePrefix Missing method name"
            }
        }
    }

    $methodEntry = $MethodRegistry[$methodName.ToLower()]
    if (-not $methodEntry) {
        return @{
            error = @{
                code    = $ErrorCodes.MethodNotFound
                message = "$providerNamePrefix Unknown method: $methodName"
            }
        }
    }

    $validationError = Validate-MethodInput -data $methodData -requiredFields $methodEntry.RequiredFields
    if ($validationError) {
        return $validationError
    }

    try {
        return & $methodEntry.Handler -data $methodData
    }
    catch {
        Write-DebugLog ("Method execution failure for [{0}]: {1}" -f $methodName, $_.Exception.Message)
        return @{
            error = @{
                code    = $ErrorCodes.InternalError
                message = "$providerNamePrefix Method execution failed"
            }
        }
    }
}

function Set-IgnoreInvalidCertificates {
    if ($PSVersionTable.PSEdition -ne 'Core') {
        try {
            [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
            [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
            Write-DebugLog "Configured certificate callback for Windows PowerShell."
        }
        catch {
            Write-DebugLog ("Failed setting certificate callback: {0}" -f $_.Exception.Message)
            throw
        }
    }
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

        $Body = $null
    )

    $irmParams = @{
        Uri                  = $Uri
        Headers              = $Headers
        Method               = $Method
        ErrorAction          = 'Stop'
        SkipCertificateCheck = $true
        SkipHeaderValidation = $true
    }

    if ($null -ne $Body) {
        $irmParams.Body = $Body
    }

    Write-DebugLog ("Invoke-RestMethod {0} {1}" -f $Method, $Uri)

    try {
        return Invoke-RestMethod @irmParams
    }
    catch {
        Write-DebugLog ("Invoke-RestMethod failed: {0}" -f $_.Exception.Message)
        if ($_.Exception.InnerException) {
            Write-DebugLog ("InnerException: {0}" -f $_.Exception.InnerException.Message)
        }
        throw
    }
}

function Get-CurrentSession {
    $session = $global:ProxmoxSession

    if (-not $session -or -not $session.host -or -not $session.header) {
        throw "Session not initialized (host/header missing)"
    }

    return $session
}

function Get-ProxmoxVmInventory {
    param([hashtable]$Session)

    $baseUrl = ("https://{0}" -f $Session.host).TrimEnd('/')
    $uri = "$baseUrl/api2/json/cluster/resources?type=vm"

    $response = Invoke-ProxmoxRestMethod -Uri $uri -Headers $Session.header -Method GET
    return $response.data
}

function Get-ProxmoxVmNode {
    param(
        [hashtable]$Session,
        [string]$VmId
    )

    $clusterVMs = Get-ProxmoxVmInventory -Session $Session
    return ($clusterVMs | Where-Object { [string]$_.vmid -eq [string]$VmId } | Select-Object -First 1)
}

# Method Handlers
function Handle-Initialize {
    $capabilities = [PSCustomObject]@{
        suspend             = $true
        guests_polling_rate = 5
    }

    return @{
        result = @{
            version      = "1.0.0"
            capabilities = $capabilities
        }
    }
}

# Expected settings:
# $settings = @{
#     host         = 'proxmox.example.com'
#     username     = 'root@pam'
#     token_name   = 'automation'
#     token_secret = 'c1234b58-1234-1234-1234-bafeb35e523f'
# }
function Handle-Connect {
    param($params)

    $settings = $params.settings

    if (-not $settings.token_secret -or -not $settings.token_name -or -not $settings.username -or -not $settings.host) {
        return @{
            error = @{
                code    = $ErrorCodes.InvalidParams
                message = "$providerNamePrefix Invalid connection parameters"
            }
        }
    }

    try {
        $proxmoxHost = [string]$settings.host
        $user        = [string]$settings.username
        $tokenName   = [string]$settings.token_name
        $tokenSecret = [string]$settings.token_secret

        $header = @{
            "Authorization" = "PVEAPIToken=$user!$tokenName=$tokenSecret"
        }

        $versionUrl  = "https://$proxmoxHost/api2/json/version"
        $versionResp = Invoke-ProxmoxRestMethod -Uri $versionUrl -Headers $header -Method GET
        $version     = $versionResp.data.version

        $global:ProxmoxSession = @{
            token_secret = $tokenSecret
            token_name   = $tokenName
            user         = $user
            host         = $proxmoxHost
            header       = $header
        }

        return @{
            result = @{
                message = "$providerNamePrefix Connected successfully to Proxmox $proxmoxHost (version: $version)"
            }
        }
    }
    catch {
        Remove-Variable -Name ProxmoxSession -Scope Global -ErrorAction SilentlyContinue

        return @{
            error = @{
                code    = $ErrorCodes.InternalError
                message = "$providerNamePrefix Failed to connect to Proxmox: $($_.Exception.Message)"
            }
        }
    }
}

function Handle-Disconnect {
    try {
        $existingHost = $null
        if ($global:ProxmoxSession -and $global:ProxmoxSession.host) {
            $existingHost = $global:ProxmoxSession.host
        }

        Remove-Variable -Name ProxmoxSession -Scope Global -ErrorAction SilentlyContinue

        $message = if ($existingHost) {
            "Session cleared on Proxmox $existingHost"
        }
        else {
            "Session cleared"
        }

        return @{
            result = @{
                message = $message
            }
        }
    }
    catch {
        return @{
            error = @{
                code    = $ErrorCodes.InternalError
                message = "$providerNamePrefix Failed to clear session: $($_.Exception.Message)"
            }
        }
    }
}

function Handle-GuestList {
    try {
        $session  = Get-CurrentSession
        $vms      = Get-ProxmoxVmInventory -Session $session
        $guestIDs = @()

        foreach ($vm in $vms) {
            if ($null -ne $vm.vmid) {
                $guestIDs += [string]$vm.vmid
            }
        }

        return @{
            result = @{
                guests = $guestIDs
            }
        }
    }
    catch {
        return @{
            error = @{
                code    = $ErrorCodes.InternalError
                message = "$providerNamePrefix Failed to retrieve guest list: $($_.Exception.Message)"
            }
        }
    }
}

function Get-GuestControl {
    param([string]$control)

    $controlMap = @{
        "start"   = "start"
        "stop"    = "shutdown"
        "reset"   = "reset"
        "restart" = "reboot"
        "delete"  = "delete"
        "suspend" = "suspend"
    }

    $normalized = $control.Trim().ToLowerInvariant()

    if ($controlMap.ContainsKey($normalized)) {
        return $controlMap[$normalized]
    }

    return "UNKNOWN"
}

function Handle-GuestControl {
    param($params)

    if (-not $params.id) {
        return @{
            error = @{
                code    = $ErrorCodes.InvalidParams
                message = "$providerNamePrefix Invalid guest id"
            }
        }
    }

    if (-not $params.control) {
        return @{
            error = @{
                code    = $ErrorCodes.InvalidParams
                message = "$providerNamePrefix Invalid guest control"
            }
        }
    }

    try {
        $session = Get-CurrentSession

        $requestedControl = [string]$params.control
        $mappedAction     = Get-GuestControl $requestedControl

        if ($mappedAction -eq 'UNKNOWN') {
            return @{
                error = @{
                    code    = $ErrorCodes.InvalidParams
                    message = "$providerNamePrefix Unsupported guest control: $($params.control)"
                }
            }
        }

        $vmid   = [string]$params.id
        $base   = ("https://{0}" -f $session.host).TrimEnd('/')
        $vmInfo = Get-ProxmoxVmNode -Session $session -VmId $vmid

        if (-not $vmInfo -or -not $vmInfo.node) {
            return @{
                error = @{
                    code    = $ErrorCodes.InvalidParams
                    message = "$providerNamePrefix VM id [$vmid] not found in cluster"
                }
            }
        }

        $node   = [string]$vmInfo.node
        $method = $null
        $uri    = $null
        $body   = $null

        if ($mappedAction -eq 'delete') {
            $method = 'DELETE'
            $uri    = "$base/api2/json/nodes/$node/qemu/$vmid"
        }
        else {
            $method = 'POST'
            $uri    = "$base/api2/json/nodes/$node/qemu/$vmid/status/$mappedAction"
            $body   = @{}
        }

        $resp = Invoke-ProxmoxRestMethod -Uri $uri -Headers $session.header -Method $method -Body $body
        $upid = $resp.data

        return @{
            result = @{
                vmid    = $vmid
                node    = $node
                action  = $mappedAction
                upid    = $upid
                message = "$providerNamePrefix Guest control [$($params.control)] submitted successfully"
            }
        }
    }
    catch {
        return @{
            error = @{
                code    = $ErrorCodes.InternalError
                message = "$providerNamePrefix Failed to control [$($params.control)] the VM: $($_.Exception.Message)"
            }
        }
    }
}

function Handle-GuestInfo {
    param($params)

    if (-not $params.id -or @($params.id).Count -eq 0) {
        return @{
            error = @{
                code    = $ErrorCodes.InvalidParams
                message = "$providerNamePrefix Invalid or missing guest ID list"
            }
        }
    }

    try {
        $session     = Get-CurrentSession
        $proxmoxHost = $session.host
        $header      = $session.header
        $ids         = @($params.id)

        $clusterVMs = Get-ProxmoxVmInventory -Session $session

        $statusMap = @{
            'running'   = 'powered_on'
            'stopped'   = 'powered_off'
            'paused'    = 'paused'
            'suspended' = 'suspended'
            'shutdown'  = 'powered_off'
            'halting'   = 'powered_off'
            'prelaunch' = 'starting'
            'unknown'   = 'unknown'
        }

        $guests = @{}

        foreach ($guestId in $ids) {
            $guestIdString = [string]$guestId
            $vmInfo = ($clusterVMs | Where-Object { [string]$_.vmid -eq $guestIdString } | Select-Object -First 1)

            if ($vmInfo -and $vmInfo.node) {
                $nodeName = [string]$vmInfo.node

                try {
                    $vmUrl  = "https://$proxmoxHost/api2/json/nodes/$nodeName/qemu/$guestIdString/status/current"
                    $vmResp = Invoke-ProxmoxRestMethod -Uri $vmUrl -Headers $header -Method GET
                    $vmData = $vmResp.data

                    $proxmoxStatusRaw = if (
                        $vmData.PSObject.Properties.Name -contains 'qmpstatus' -and
                        $vmData.qmpstatus
                    ) {
                        $vmData.qmpstatus
                    }
                    else {
                        $vmData.status
                    }

                    $proxmoxStatus = if ($proxmoxStatusRaw) {
                        $proxmoxStatusRaw.ToString().Trim().ToLowerInvariant()
                    }
                    else {
                        'unknown'
                    }

                    $mappedState = $statusMap[$proxmoxStatus]
                    if (-not $mappedState) {
                        $mappedState = 'unknown'
                    }

                    $name = if ($vmData.name) {
                        $vmData.name
                    }
                    elseif ($vmInfo.name) {
                        $vmInfo.name
                    }
                    else {
                        "VM-$guestIdString"
                    }

                    $ipv4Set = New-Object 'System.Collections.Generic.HashSet[string]'
                    $macSet  = New-Object 'System.Collections.Generic.HashSet[string]'

                    try {
                        $agentUrl  = "https://$proxmoxHost/api2/json/nodes/$nodeName/qemu/$guestIdString/agent/network-get-interfaces"
                        $agentResp = Invoke-ProxmoxRestMethod -Uri $agentUrl -Headers $header -Method GET
                        $ifaces    = $agentResp.data.result
                        if (-not $ifaces) { $ifaces = $agentResp.data }

                        foreach ($iface in @($ifaces)) {
                            $mac = if ($iface.PSObject.Properties.Name -contains 'hardware-address') {
                                $iface.'hardware-address'
                            }
                            else {
                                $null
                            }

                            $ipAddrs = $iface.'ip-addresses'

                            foreach ($ip in @($ipAddrs)) {
                                $type = $ip.'ip-address-type'
                                $addr = $ip.'ip-address'

                                if (
                                    $type -eq 'ipv4' -and
                                    $addr -and
                                    $addr -notmatch '^169\.254\.' -and
                                    $addr -ne '127.0.0.1'
                                ) {
                                    [void]$ipv4Set.Add($addr)
                                    if ($mac) {
                                        [void]$macSet.Add(([string]$mac).ToUpper())
                                    }
                                }
                            }
                        }
                    }
                    catch {
                        Write-DebugLog ("Guest agent query failed for VM [{0}]: {1}" -f $guestIdString, $_.Exception.Message)
                    }

                    try {
                        $cfgUrl  = "https://$proxmoxHost/api2/json/nodes/$nodeName/qemu/$guestIdString/config"
                        $cfgResp = Invoke-ProxmoxRestMethod -Uri $cfgUrl -Headers $header -Method GET
                        $cfg     = $cfgResp.data

                        if ($cfg) {
                            foreach ($prop in $cfg.PSObject.Properties) {
                                if ($prop.Name -match '^net\d+$' -and $prop.Value) {
                                    $m = [regex]::Match([string]$prop.Value, '([0-9A-Fa-f]{2}(:[0-9A-Fa-f]{2}){5})')
                                    if ($m.Success) {
                                        [void]$macSet.Add($m.Groups[1].Value.ToUpper())
                                    }
                                }
                            }
                        }
                    }
                    catch {
                        Write-DebugLog ("Config query failed for VM [{0}]: {1}" -f $guestIdString, $_.Exception.Message)
                    }

                    $guests[$guestIdString] = @{
                        name         = $name
                        state        = $mappedState
                        ip_addresses = @($ipv4Set)
                        macAddresses = @($macSet)
                    }
                }
                catch {
                    $guests[$guestIdString] = @{
                        name         = $null
                        state        = 'error'
                        ip_addresses = @()
                        macAddresses = @()
                    }
                }
            }
            else {
                $guests[$guestIdString] = @{
                    name         = $null
                    state        = 'not found'
                    ip_addresses = @()
                    macAddresses = @()
                }
            }
        }

        $resultObject = if ($ids.Count -eq 1) {
            $guests[[string]$ids[0]]
        }
        else {
            $guests
        }

        return @{
            result = $resultObject
        }
    }
    catch {
        return @{
            error = @{
                code    = $ErrorCodes.InternalError
                message = "$providerNamePrefix Failed to retrieve guest info: $($_.Exception.Message)"
            }
        }
    }
}

# Main processing loop
while ($true) {
    $inputLine = [Console]::In.ReadLine()

    if (-not $inputLine) {
        Write-DebugLog "Listener stopped."
        break
    }

    Write-DebugLog ("Input: {0}" -f $inputLine)

    try {
        $response = Process-Method $inputLine.Trim()
    }
    catch {
        $response = @{
            error = @{
                code    = $ErrorCodes.InternalError
                message = "$providerNamePrefix Failed to process input: $($_.Exception.Message)"
            }
        }
    }

    Send-Response $response
}
