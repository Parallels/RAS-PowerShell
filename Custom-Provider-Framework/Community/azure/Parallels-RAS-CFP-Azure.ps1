<#
.SYNOPSIS
    Parallels RAS Custom Provider sample script for Microsoft Azure.
.DESCRIPTION
    Implements a Parallels RAS Custom Provider that integrates Microsoft Azure
    virtual machines through the Azure Resource Manager (ARM) REST API. It
    authenticates to Microsoft Entra ID with a service principal (client
    credentials) and manages VMs in a single subscription and resource group.

    It listens for JSON-RPC requests on standard input, processes them per the
    Custom Provider Framework (CPF) protocol, and writes responses on standard
    output.

    Azure has no in-place VM snapshots, so this sample uses the "basic"
    (full-clone) template model from the CPF "Capabilities" documentation:
      - convert to template -> capture the VM into a managed image, tag the VM
      - convert from template -> delete the managed image, remove the tag
      - clone -> create a NIC and a new VM from the managed image

    Suspend is not offered: Azure VMs are deallocated (compute released) on stop,
    which is mapped to the CPF "stop" control. State values returned to RAS:
    powered_on, powered_off, powering_on, powering_off.

    Requires PowerShell 7 or later.
.NOTES
    File Name : Parallels-RAS-CFP-Azure.ps1
    Platform  : Microsoft Azure (Azure Resource Manager REST API)
    Author    : Edwin Houben
    Reference : azure/Azure-API.md in this folder
.EXAMPLE
    {"method":"provider/initialize"}
    {"method":"provider/connect","params":{"settings":{"tenant_id":"<guid>","client_id":"<guid>","client_secret":"<secret>","subscription_id":"<guid>","resource_group":"vdi-rg","location":"westeurope"}}}
    {"method":"guests/list"}
    {"method":"guests/control","params":{"id":"vdi-vm-01","control":"start"}}
    {"method":"guests/convert","params":{"id":"vdi-gold","is_template":true}}
    {"method":"guests/clone","params":{"id":"vdi-gold","name":"vdi-clone-01"}}
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

$script:ProviderNamePrefix = 'Azure:'
$script:LogPath            = Join-Path ([System.IO.Path]::GetTempPath()) 'Azure-RAS-Provider.log'
$script:Session            = $null
$script:TemplateTagKey     = 'ras_template'
$script:SourceTagKey       = 'ras_source_vm'

$script:Arm                = 'https://management.azure.com'
$script:Authority          = 'https://login.microsoftonline.com'
$script:ApiVersions        = @{
    Compute = '2024-07-01'   # virtualMachines, images
    Network = '2024-05-01'   # networkInterfaces
}

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
# Authentication and REST plumbing
# ----------------------------------------------------------------------------

function Get-AzureToken {
    <#
        Acquires an ARM access token with the client-credentials grant. Returns
        @{ token = <string>; expires = <DateTime> }.
    #>
    param([hashtable]$Creds, [bool]$SkipTls)

    $tokenUri = "$script:Authority/$($Creds.tenant_id)/oauth2/v2.0/token"
    $form = @{
        client_id     = $Creds.client_id
        client_secret = $Creds.client_secret
        grant_type    = 'client_credentials'
        scope         = "$script:Arm/.default"
    }

    $irm = @{
        Uri         = $tokenUri
        Method      = 'POST'
        Body        = $form
        ContentType = 'application/x-www-form-urlencoded'
        ErrorAction = 'Stop'
    }
    if ($PSVersionTable.PSEdition -eq 'Core' -and $SkipTls) { $irm.SkipCertificateCheck = $true }

    $resp = Invoke-RestMethod @irm
    $token     = if (Test-HasProperty -Object $resp -Name 'access_token') { [string]$resp.access_token } else { '' }
    $expiresIn = if (Test-HasProperty -Object $resp -Name 'expires_in') { [int]$resp.expires_in } else { 3600 }
    return @{ token = $token; expires = (Get-Date).AddSeconds($expiresIn) }
}

function Get-Session {
    if ($null -eq $script:Session) { throw 'Session not initialized' }

    # Refresh the token a minute before it expires so long-running sessions keep working.
    if ((Get-Date) -ge $script:Session.token_expires.AddSeconds(-60)) {
        Write-DebugLog 'Access token expired or expiring; refreshing.'
        $fresh = Get-AzureToken -Creds $script:Session.creds -SkipTls $script:Session.skip_tls
        $script:Session.token         = $fresh.token
        $script:Session.token_expires = $fresh.expires
    }
    if ([string]::IsNullOrWhiteSpace($script:Session.token)) { throw 'Session token missing' }
    return $script:Session
}

function Invoke-Arm {
    <#
        Calls the Azure Resource Manager REST API with the bearer token. $Path is
        a full ARM resource path beginning with /subscriptions/...; api-version is
        appended automatically. Throws on HTTP error. Use Get-ArmResourceOrNull
        for 404-tolerant reads. Returns @{ body = <object>; headers = <headers> }.
    #>
    param(
        [Parameter(Mandatory = $true)][ValidateSet('GET', 'POST', 'PUT', 'PATCH', 'DELETE')][string]$Method,
        [Parameter(Mandatory = $true)][string]$Path,
        [object]$Body = $null,
        [string]$ApiVersion = $script:ApiVersions.Compute
    )

    $session = Get-Session
    $sep = if ($Path -match '\?') { '&' } else { '?' }
    $uri = $script:Arm + $Path + $sep + 'api-version=' + $ApiVersion

    $irm = @{
        Uri                     = $uri
        Method                  = $Method
        Headers                 = @{ Authorization = "Bearer $($session.token)" }
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

function Get-ArmResourceOrNull {
    param([Parameter(Mandatory = $true)][string]$Path, [string]$ApiVersion = $script:ApiVersions.Compute)
    try { return (Invoke-Arm -Method GET -Path $Path -ApiVersion $ApiVersion).body }
    catch {
        if ((Get-HttpStatusCode -ErrorRecord $_) -eq 404) { return $null }
        throw
    }
}

# ----------------------------------------------------------------------------
# Resource path helpers
# ----------------------------------------------------------------------------

function Get-VmPath {
    param([string]$Name, [string]$ResourceGroup = $null)
    $session = Get-Session
    $rg = if ([string]::IsNullOrWhiteSpace($ResourceGroup)) { $session.resource_group } else { $ResourceGroup }
    return "/subscriptions/$($session.subscription_id)/resourceGroups/$rg/providers/Microsoft.Compute/virtualMachines/$Name"
}

function Get-ImagePath {
    param([string]$Name)
    $session = Get-Session
    return "/subscriptions/$($session.subscription_id)/resourceGroups/$($session.image_resource_group)/providers/Microsoft.Compute/images/$Name"
}

function Get-NicPath {
    param([string]$Name)
    $session = Get-Session
    return "/subscriptions/$($session.subscription_id)/resourceGroups/$($session.resource_group)/providers/Microsoft.Network/networkInterfaces/$Name"
}

# ----------------------------------------------------------------------------
# VM read helpers
# ----------------------------------------------------------------------------

function Get-Vm {
    param([Parameter(Mandatory = $true)][string]$Name, [switch]$InstanceView)
    $path = Get-VmPath -Name $Name
    if ($InstanceView) { $path = "$path`?`$expand=instanceView" }
    return Get-ArmResourceOrNull -Path $path
}

function Get-PowerStateCode {
    param([object]$Vm)
    # instanceView.statuses holds entries like "PowerState/running".
    if ((Test-HasProperty -Object $Vm -Name 'properties') -and
        (Test-HasProperty -Object $Vm.properties -Name 'instanceView') -and
        (Test-HasProperty -Object $Vm.properties.instanceView -Name 'statuses')) {
        foreach ($st in @($Vm.properties.instanceView.statuses)) {
            if ((Test-HasProperty -Object $st -Name 'code') -and ([string]$st.code).StartsWith('PowerState/')) {
                return (([string]$st.code) -split '/', 2)[1]
            }
        }
    }
    return ''
}

function Map-PowerStateToRasState {
    param([string]$PowerState)
    switch ($PowerState.Trim().ToLowerInvariant()) {
        'running'      { return 'powered_on' }
        'starting'     { return 'powering_on' }
        'stopping'     { return 'powering_off' }
        'deallocating' { return 'powering_off' }
        'stopped'      { return 'powered_off' }
        'deallocated'  { return 'powered_off' }
        default        { return 'powered_off' }
    }
}

function Get-VmNetworkData {
    param([object]$Vm)
    $ips  = New-Object 'System.Collections.Generic.List[string]'
    $macs = New-Object 'System.Collections.Generic.List[string]'

    if (-not ((Test-HasProperty -Object $Vm -Name 'properties') -and
              (Test-HasProperty -Object $Vm.properties -Name 'networkProfile') -and
              (Test-HasProperty -Object $Vm.properties.networkProfile -Name 'networkInterfaces'))) {
        return @{ IPv4Addresses = @(); MacAddresses = @() }
    }

    foreach ($ref in @($Vm.properties.networkProfile.networkInterfaces)) {
        if (-not (Test-HasProperty -Object $ref -Name 'id')) { continue }
        try {
            $nic = Get-ArmResourceOrNull -Path ([string]$ref.id) -ApiVersion $script:ApiVersions.Network
            if ($null -eq $nic -or -not (Test-HasProperty -Object $nic -Name 'properties')) { continue }

            if ((Test-HasProperty -Object $nic.properties -Name 'macAddress') -and
                -not [string]::IsNullOrWhiteSpace([string]$nic.properties.macAddress)) {
                $mac = ([string]$nic.properties.macAddress).ToUpperInvariant()
                if (-not $macs.Contains($mac)) { $macs.Add($mac) }
            }

            if (Test-HasProperty -Object $nic.properties -Name 'ipConfigurations') {
                foreach ($cfg in @($nic.properties.ipConfigurations)) {
                    if ((Test-HasProperty -Object $cfg -Name 'properties') -and
                        (Test-HasProperty -Object $cfg.properties -Name 'privateIPAddress') -and
                        -not [string]::IsNullOrWhiteSpace([string]$cfg.properties.privateIPAddress)) {
                        $ip = [string]$cfg.properties.privateIPAddress
                        if (-not $ips.Contains($ip)) { $ips.Add($ip) }
                    }
                }
            }
        }
        catch { Write-DebugLog "NIC read failed for [$([string]$ref.id)]: $($_.Exception.Message)" }
    }

    return @{ IPv4Addresses = @($ips | Select-Object -First 3); MacAddresses = @($macs | Select-Object -First 3) }
}

function Get-VmIsTemplate {
    param([object]$Vm)
    if ((Test-HasProperty -Object $Vm -Name 'tags') -and (Test-HasProperty -Object $Vm.tags -Name $script:TemplateTagKey)) {
        return ([string]$Vm.tags.$($script:TemplateTagKey) -eq 'true')
    }
    return $false
}

function ConvertTo-RasGuestObject {
    param([Parameter(Mandatory = $true)][string]$Name)

    $vm = Get-Vm -Name $Name -InstanceView
    if ($null -eq $vm) { throw "VM [$Name] not found" }

    $powerState = Get-PowerStateCode -Vm $vm
    $net        = Get-VmNetworkData -Vm $vm

    $osType = 'unknown'
    if ((Test-HasProperty -Object $vm -Name 'properties') -and
        (Test-HasProperty -Object $vm.properties -Name 'storageProfile') -and
        (Test-HasProperty -Object $vm.properties.storageProfile -Name 'osDisk') -and
        (Test-HasProperty -Object $vm.properties.storageProfile.osDisk -Name 'osType')) {
        $osType = [string]$vm.properties.storageProfile.osDisk.osType
    }

    $guest = @{
        id            = [string]$Name
        name          = [string]$Name
        provider      = 'Azure'
        state         = (Map-PowerStateToRasState -PowerState $powerState)
        power_state   = $(if ([string]::IsNullOrWhiteSpace($powerState)) { 'unknown' } else { $powerState })
        host_os       = $osType
        ip            = $(if ($net.IPv4Addresses.Count -gt 0) { $net.IPv4Addresses[0] } else { $null })
        ip_addresses  = @($net.IPv4Addresses)
        mac_addresses = @($net.MacAddresses)
        is_template   = (Get-VmIsTemplate -Vm $vm)
        type          = 'Virtual Machine'
    }

    Write-DebugLog ("GUEST id={0}; state={1}; power={2}; template={3}; ips={4}" -f `
            $guest.id, $guest.state, $guest.power_state, $guest.is_template, ($guest.ip_addresses -join ','))
    return $guest
}

function Get-VmTagTable {
    param([object]$Vm)
    $tags = @{}
    if ((Test-HasProperty -Object $Vm -Name 'tags') -and $null -ne $Vm.tags) {
        foreach ($p in $Vm.tags.PSObject.Properties) { $tags[$p.Name] = [string]$p.Value }
    }
    return $tags
}

function Set-VmTemplateTag {
    param([string]$Name, [bool]$IsTemplate)
    $vm = Get-Vm -Name $Name
    if ($null -eq $vm) { throw "VM [$Name] not found" }
    $tags = Get-VmTagTable -Vm $vm
    if ($IsTemplate) { $tags[$script:TemplateTagKey] = 'true' }
    elseif ($tags.ContainsKey($script:TemplateTagKey)) { $tags.Remove($script:TemplateTagKey) | Out-Null }
    Invoke-Arm -Method PATCH -Path (Get-VmPath -Name $Name) -Body @{ tags = $tags } | Out-Null
}

function Get-ControlOperation {
    param([Parameter(Mandatory = $true)][string]$Control)
    switch ($Control.Trim().ToLowerInvariant()) {
        'start'   { return 'start' }
        'stop'    { return 'deallocate' }   # release compute (no charge), cloud-appropriate stop
        'restart' { return 'restart' }
        'reset'   { return 'restart' }      # Azure exposes a single reboot operation
        'delete'  { return 'delete' }
        default   { return $null }          # suspend is not supported on Azure
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
                can_suspend_guests    = $false
                guests_polling_rate   = 15
                tasks_polling_rate    = 10
                tasks_polling_retries = 180
                template_method       = 'basic'
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

    $tenant   = if (Test-HasProperty -Object $settings -Name 'tenant_id') { [string]$settings.tenant_id } else { '' }
    $clientId = if (Test-HasProperty -Object $settings -Name 'client_id') { [string]$settings.client_id } else { '' }
    $secret   = if (Test-HasProperty -Object $settings -Name 'client_secret') { [string]$settings.client_secret } else { '' }
    $subId    = if (Test-HasProperty -Object $settings -Name 'subscription_id') { [string]$settings.subscription_id } else { '' }
    $rg       = if (Test-HasProperty -Object $settings -Name 'resource_group') { [string]$settings.resource_group } else { '' }
    $location = if (Test-HasProperty -Object $settings -Name 'location') { [string]$settings.location } else { '' }
    $imageRg  = if (Test-HasProperty -Object $settings -Name 'image_resource_group') { [string]$settings.image_resource_group } else { $rg }
    $subnetId = if (Test-HasProperty -Object $settings -Name 'subnet_id') { [string]$settings.subnet_id } else { '' }
    $adminU   = if (Test-HasProperty -Object $settings -Name 'admin_username') { [string]$settings.admin_username } else { '' }
    $adminP   = if (Test-HasProperty -Object $settings -Name 'admin_password') { [string]$settings.admin_password } else { '' }
    $skipTls  = $false
    if (Test-HasProperty -Object $settings -Name 'skip_tls') {
        try { $skipTls = [System.Convert]::ToBoolean($settings.skip_tls) } catch { $skipTls = $false }
    }

    if ([string]::IsNullOrWhiteSpace($tenant) -or [string]::IsNullOrWhiteSpace($clientId) -or
        [string]::IsNullOrWhiteSpace($secret) -or [string]::IsNullOrWhiteSpace($subId) -or
        [string]::IsNullOrWhiteSpace($rg) -or [string]::IsNullOrWhiteSpace($location)) {
        return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams `
            -Message "$($script:ProviderNamePrefix) tenant_id, client_id, client_secret, subscription_id, resource_group and location are required"
    }

    try {
        $creds = @{ tenant_id = $tenant; client_id = $clientId; client_secret = $secret }
        $auth  = Get-AzureToken -Creds $creds -SkipTls $skipTls
        if ([string]::IsNullOrWhiteSpace($auth.token)) {
            return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Authentication did not return a token"
        }

        $script:Session = @{
            creds                = $creds
            token                = $auth.token
            token_expires        = $auth.expires
            subscription_id      = $subId
            resource_group       = $rg
            image_resource_group = if ([string]::IsNullOrWhiteSpace($imageRg)) { $rg } else { $imageRg }
            location             = $location
            subnet_id            = $subnetId
            admin_username       = $adminU
            admin_password       = $adminP
            skip_tls             = $skipTls
        }

        # Validate by listing VMs in the resource group.
        $null = Invoke-Arm -Method GET -Path "/subscriptions/$subId/resourceGroups/$rg/providers/Microsoft.Compute/virtualMachines"

        Write-DebugLog "Connected; subscription=$subId rg=$rg location=$location"
        return @{ result = @{ message = "$($script:ProviderNamePrefix) Connected to subscription $subId / $rg" } }
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
        $resp = (Invoke-Arm -Method GET -Path "/subscriptions/$($session.subscription_id)/resourceGroups/$($session.resource_group)/providers/Microsoft.Compute/virtualMachines").body
        $guests = @()
        if (Test-HasProperty -Object $resp -Name 'value') {
            foreach ($vm in @($resp.value)) {
                if (Test-HasProperty -Object $vm -Name 'name') { $guests += [string]$vm.name }
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
            return @{ result = (ConvertTo-RasGuestObject -Name ([string]$ids[0])) }
        }

        $resultMap = @{}
        foreach ($id in $ids) {
            $name = [string]$id
            try { $resultMap[$name] = ConvertTo-RasGuestObject -Name $name }
            catch {
                Write-DebugLog "Guest get failed for [$name]: $($_.Exception.Message)"
                $resultMap[$name] = @{
                    id = $name; name = $name; provider = 'Azure'
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

    $operation = Get-ControlOperation -Control ([string]$Params.control)
    if ([string]::IsNullOrWhiteSpace($operation)) {
        return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Unsupported guest control: $($Params.control)"
    }

    try {
        $name = [string]$Params.id
        $vmPath = Get-VmPath -Name $name

        switch ($operation) {
            'delete'     { Invoke-Arm -Method DELETE -Path $vmPath | Out-Null }
            'start'      { Invoke-Arm -Method POST -Path "$vmPath/start" | Out-Null }
            'deallocate' { Invoke-Arm -Method POST -Path "$vmPath/deallocate" | Out-Null }
            'restart'    { Invoke-Arm -Method POST -Path "$vmPath/restart" | Out-Null }
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
        $name       = [string]$Params.id
        $isTemplate = [System.Convert]::ToBoolean($Params.is_template)
        $imageName  = "$name-image"

        if ($isTemplate) {
            # Capture the VM into a managed image, then tag the VM as a template.
            # IMPORTANT: the source VM must already be generalized (sysprep on
            # Windows, waagent deprovision on Linux, then deallocate + generalize).
            # This connector cannot sysprep inside the guest, so it captures the VM
            # as-is; a non-generalized source yields an image that cannot boot a
            # working clone. See azure/README.md and azure/Azure-API.md.
            $vmResourceId = Get-VmPath -Name $name
            $imageBody = @{
                location   = $session.location
                tags       = @{ $script:SourceTagKey = $name }
                properties = @{
                    hyperVGeneration     = 'V2'
                    sourceVirtualMachine = @{ id = $vmResourceId }
                }
            }
            Invoke-Arm -Method PUT -Path (Get-ImagePath -Name $imageName) -Body $imageBody | Out-Null
            Set-VmTemplateTag -Name $name -IsTemplate $true
            return @{ result = @{ task_id = "image:$imageName" } }
        }
        else {
            try { Invoke-Arm -Method DELETE -Path (Get-ImagePath -Name $imageName) | Out-Null }
            catch { Write-DebugLog "Image delete ignored for [$imageName]: $($_.Exception.Message)" }
            Set-VmTemplateTag -Name $name -IsTemplate $false
            return @{ result = @{ task_id = 'sync:convert' } }
        }
    }
    catch {
        return New-ErrorResponse -Code $script:ErrorCodes.InternalError -Message "$($script:ProviderNamePrefix) Failed to convert guest [$($Params.id)]: $($_.Exception.Message)"
    }
}

function Get-SafeComputerName {
    param([string]$Name)
    $clean = ($Name -replace '[^A-Za-z0-9-]', '')
    if ($clean.Length -gt 15) { $clean = $clean.Substring(0, 15) }   # Windows computer-name limit
    if ([string]::IsNullOrWhiteSpace($clean)) { $clean = 'rasclone' }
    return $clean
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

        if ([string]::IsNullOrWhiteSpace($session.subnet_id)) {
            return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) subnet_id is required to clone (set it in the provider settings)"
        }
        if ([string]::IsNullOrWhiteSpace($session.admin_username) -or [string]::IsNullOrWhiteSpace($session.admin_password)) {
            return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) admin_username and admin_password are required to clone from a generalized image"
        }

        # The template image is the one created by guests/convert for the source VM.
        $imageName = "$sourceId-image"
        $image = Get-ArmResourceOrNull -Path (Get-ImagePath -Name $imageName)
        if ($null -eq $image) {
            return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Template image [$imageName] not found; convert the source VM to a template first"
        }
        $imageState = ''
        if ((Test-HasProperty -Object $image -Name 'properties') -and (Test-HasProperty -Object $image.properties -Name 'provisioningState')) {
            $imageState = ([string]$image.properties.provisioningState).ToLowerInvariant()
        }
        if ($imageState -ne 'succeeded') {
            return New-ErrorResponse -Code $script:ErrorCodes.InvalidParams -Message "$($script:ProviderNamePrefix) Template image [$imageName] is not ready (provisioningState=$imageState); wait for the convert task to complete before cloning"
        }
        $imageId = [string]$image.id

        # Determine VM size and OS type from the source VM.
        $src = Get-Vm -Name $sourceId
        if ($null -eq $src) { throw "Source VM [$sourceId] not found" }
        $vmSize = 'Standard_B2s'
        if ((Test-HasProperty -Object $src -Name 'properties') -and
            (Test-HasProperty -Object $src.properties -Name 'hardwareProfile') -and
            (Test-HasProperty -Object $src.properties.hardwareProfile -Name 'vmSize')) {
            $vmSize = [string]$src.properties.hardwareProfile.vmSize
        }
        $osType = 'Windows'
        if ((Test-HasProperty -Object $src -Name 'properties') -and
            (Test-HasProperty -Object $src.properties -Name 'storageProfile') -and
            (Test-HasProperty -Object $src.properties.storageProfile -Name 'osDisk') -and
            (Test-HasProperty -Object $src.properties.storageProfile.osDisk -Name 'osType')) {
            $osType = [string]$src.properties.storageProfile.osDisk.osType
        }

        # Create a NIC for the clone in the configured subnet.
        $nicName = "$cloneName-nic"
        $nicBody = @{
            location   = $session.location
            properties = @{
                ipConfigurations = @(
                    @{
                        name       = 'ipconfig1'
                        properties = @{ subnet = @{ id = $session.subnet_id }; privateIPAllocationMethod = 'Dynamic' }
                    }
                )
            }
        }
        $nicResp = Invoke-Arm -Method PUT -Path (Get-NicPath -Name $nicName) -Body $nicBody -ApiVersion $script:ApiVersions.Network
        $nicId = if ((Test-HasProperty -Object $nicResp.body -Name 'id')) { [string]$nicResp.body.id } else { (Get-NicPath -Name $nicName) }

        # Create the VM from the managed image.
        $vmBody = @{
            location   = $session.location
            properties = @{
                hardwareProfile = @{ vmSize = $vmSize }
                storageProfile  = @{
                    imageReference = @{ id = $imageId }
                    osDisk         = @{ createOption = 'FromImage'; managedDisk = @{ storageAccountType = 'Standard_LRS' } }
                }
                osProfile       = @{
                    computerName  = (Get-SafeComputerName -Name $cloneName)
                    adminUsername = $session.admin_username
                    adminPassword = $session.admin_password
                }
                networkProfile  = @{ networkInterfaces = @(@{ id = $nicId; properties = @{ primary = $true } }) }
            }
        }
        Invoke-Arm -Method PUT -Path (Get-VmPath -Name $cloneName) -Body $vmBody | Out-Null

        return @{ result = @{ task_id = "vm:$cloneName"; clone_id = $cloneName } }
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
        $taskId = [string]$Params.id
        $parts  = $taskId -split ':', 2
        $kind   = $parts[0]
        $arg    = if ($parts.Count -ge 2) { $parts[1] } else { '' }

        switch ($kind) {
            'sync' { return @{ result = @{ state = 'completed'; output = @{} } } }

            'image' {
                $img = Get-ArmResourceOrNull -Path (Get-ImagePath -Name $arg)
                if ($null -eq $img) { return @{ result = @{ state = 'running' } } }
                $state = ''
                if ((Test-HasProperty -Object $img -Name 'properties') -and (Test-HasProperty -Object $img.properties -Name 'provisioningState')) {
                    $state = ([string]$img.properties.provisioningState).ToLowerInvariant()
                }
                switch ($state) {
                    'succeeded' { return @{ result = @{ state = 'completed'; output = @{} } } }
                    'failed'    { return @{ result = @{ state = 'failed'; error = @{ code = 1; message = 'Image creation failed' } } } }
                    default     { return @{ result = @{ state = 'running' } } }
                }
            }

            'vm' {
                $vm = Get-Vm -Name $arg
                if ($null -eq $vm) { return @{ result = @{ state = 'running' } } }
                $state = ''
                if ((Test-HasProperty -Object $vm -Name 'properties') -and (Test-HasProperty -Object $vm.properties -Name 'provisioningState')) {
                    $state = ([string]$vm.properties.provisioningState).ToLowerInvariant()
                }
                switch ($state) {
                    'succeeded' { return @{ result = @{ state = 'completed'; output = @{ clone_id = $arg } } } }
                    'failed'    { return @{ result = @{ state = 'failed'; error = @{ code = 1; message = 'VM provisioning failed' } } } }
                    default     { return @{ result = @{ state = 'running' } } }
                }
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
    'provider/initialize' = @{ Handler = { param($data) Handle-Initialize }; RequiredFields = @() }
    'provider/connect'    = @{ Handler = { param($data) Handle-Connect -Params $data.params }; RequiredFields = @('params.settings') }
    'provider/disconnect' = @{ Handler = { param($data) Handle-Disconnect }; RequiredFields = @() }

    'hosts/list'          = @{ Handler = { param($data) Handle-HostList }; RequiredFields = @() }
    'hosts/get'           = @{ Handler = { param($data) Handle-HostGet -Params $data.params }; RequiredFields = @('params.id') }
    'hosts/control'       = @{ Handler = { param($data) Handle-HostControl -Params $data.params }; RequiredFields = @('params.id', 'params.control') }

    'guests/list'         = @{ Handler = { param($data) Handle-GuestList }; RequiredFields = @() }
    'guests/get'          = @{ Handler = { param($data) Handle-GuestGet -Params $data.params }; RequiredFields = @('params.id') }
    'guests/control'      = @{ Handler = { param($data) Handle-GuestControl -Params $data.params }; RequiredFields = @('params.id', 'params.control') }
    'guests/convert'      = @{ Handler = { param($data) Handle-GuestConvert -Params $data.params }; RequiredFields = @('params.id', 'params.is_template') }
    'guests/clone'        = @{ Handler = { param($data) Handle-GuestClone -Params $data.params }; RequiredFields = @('params.id', 'params.name') }
    'tasks/get'           = @{ Handler = { param($data) Handle-TaskInfo -Params $data.params }; RequiredFields = @('params.id') }
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

Write-DebugLog "Azure provider process started. PID=$PID"

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
