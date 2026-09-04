<#
.SYNOPSIS
    End-to-end test for the Microsoft Azure custom provider.
.DESCRIPTION
    Drives Parallels-RAS-CFP-Azure.ps1 through the shared CustomProvider.psm1
    test harness. By default it runs a safe, read-only flow:

        provider/initialize -> provider/connect -> guests/list -> guests/get -> provider/disconnect

    The mutating operation is opt-in:
      -Control <action>   runs guests/control on the target guest

    Connection settings come from the -TenantId/-ClientId/-ClientSecret/
    -SubscriptionId/-ResourceGroup/-Location parameters, or fall back to a
    CustomProvider.psd1 found next to this script or in the repo root.
.EXAMPLE
    pwsh -File .\Test-Azure.ps1 -TenantId <t> -ClientId <c> -ClientSecret $env:AZ_SECRET -SubscriptionId <s> -ResourceGroup vdi-rg -Location westeurope
.EXAMPLE
    pwsh -File .\Test-Azure.ps1 -TenantId <t> -ClientId <c> -ClientSecret $env:AZ_SECRET -SubscriptionId <s> -ResourceGroup vdi-rg -Location westeurope -GuestID vdi-vm-01 -Control start
#>

[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$ClientId,
    [string]$ClientSecret,
    [string]$SubscriptionId,
    [string]$ResourceGroup,
    [string]$Location,
    [string]$ImageResourceGroup,
    [string]$SubnetId,
    [string]$AdminUsername,
    [string]$AdminPassword,
    [bool]$SkipTls = $false,
    [string]$GuestID,
    [ValidateSet('start', 'stop', 'restart', 'reset', 'delete')]
    [string]$Control
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding  = [System.Text.Encoding]::UTF8

$ProviderDir = Split-Path $PSScriptRoot -Parent           # azure/
$RepoRoot    = Split-Path $ProviderDir -Parent            # repository root

$ProviderScript = Join-Path $ProviderDir 'Parallels-RAS-CFP-Azure.ps1'
if (-not (Test-Path $ProviderScript)) { throw "Provider script not found: $ProviderScript" }

$ModulePath = Join-Path $RepoRoot 'CustomProvider.psm1'
if (-not (Test-Path $ModulePath)) { throw "Harness module not found: $ModulePath" }
Import-Module $ModulePath -Force

function Resolve-CustomSettings {
    if (-not [string]::IsNullOrWhiteSpace($TenantId) -and
        -not [string]::IsNullOrWhiteSpace($ClientId) -and
        -not [string]::IsNullOrWhiteSpace($ClientSecret) -and
        -not [string]::IsNullOrWhiteSpace($SubscriptionId) -and
        -not [string]::IsNullOrWhiteSpace($ResourceGroup) -and
        -not [string]::IsNullOrWhiteSpace($Location)) {
        $s = @{
            tenant_id       = $TenantId
            client_id       = $ClientId
            client_secret   = $ClientSecret
            subscription_id = $SubscriptionId
            resource_group  = $ResourceGroup
            location        = $Location
            skip_tls        = $SkipTls
        }
        if (-not [string]::IsNullOrWhiteSpace($ImageResourceGroup)) { $s.image_resource_group = $ImageResourceGroup }
        if (-not [string]::IsNullOrWhiteSpace($SubnetId)) { $s.subnet_id = $SubnetId }
        if (-not [string]::IsNullOrWhiteSpace($AdminUsername)) { $s.admin_username = $AdminUsername }
        if (-not [string]::IsNullOrWhiteSpace($AdminPassword)) { $s.admin_password = $AdminPassword }
        return $s
    }

    foreach ($candidate in @(
            (Join-Path $PSScriptRoot 'CustomProvider.psd1'),
            (Join-Path $ProviderDir 'CustomProvider.psd1'),
            (Join-Path $RepoRoot 'CustomProvider.psd1'))) {
        if (Test-Path $candidate) {
            $data = Import-PowerShellDataFile -Path $candidate
            if ($data.ContainsKey('CustomSettings')) {
                Write-Host "Using settings from $candidate" -ForegroundColor DarkGray
                return $data.CustomSettings
            }
        }
    }

    throw 'Provide -TenantId, -ClientId, -ClientSecret, -SubscriptionId, -ResourceGroup and -Location, or a CustomProvider.psd1 with CustomSettings.'
}

$CustomSettings = Resolve-CustomSettings

$CommandPath = (Get-Command pwsh -ErrorAction SilentlyContinue).Source
if ([string]::IsNullOrWhiteSpace($CommandPath)) { $CommandPath = 'pwsh' }
$CommandArgs = "-NoProfile -NonInteractive -ExecutionPolicy Bypass -File `"$ProviderScript`""

function Write-Section { param([string]$Text) Write-Host "`n=== $Text ===" -ForegroundColor Cyan }
function Write-Pass    { param([string]$Text) Write-Host "  PASS: $Text" -ForegroundColor Green }

function Invoke-Pipeline {
    param([object]$IOStreams)

    Write-Section 'initialize + connect'
    Submit-InitializeAndConnect $IOStreams $CustomSettings
    Write-Pass 'connected'

    Write-Section 'guests/list'
    $list = Submit-GuestsList $IOStreams
    $guests = @($list.guests)
    Write-Host "  Guests found: $($guests.Count)"
    if ($guests.Count -gt 0) { Write-Host "  $($guests -join ', ')" }

    $target = if (-not [string]::IsNullOrWhiteSpace($GuestID)) { $GuestID }
              elseif ($guests.Count -gt 0) { [string]$guests[0] }
              else { $null }

    if ($null -eq $target) {
        Write-Host '  No guest available to inspect. Skipping guest-specific tests.' -ForegroundColor Yellow
        Submit-Disconnect $IOStreams | Out-Null
        return
    }

    Write-Section "guests/get [$target]"
    $guest = Submit-GuestsGet $IOStreams $target
    Write-Host "  name=$($guest.name) state=$($guest.state) power_state=$($guest.power_state)"
    Write-Host "  host_os=$($guest.host_os) is_template=$($guest.is_template)"
    Write-Host "  ip_addresses=$(@($guest.ip_addresses) -join ',') mac_addresses=$(@($guest.mac_addresses) -join ',')"
    Write-Pass 'guest info retrieved'

    if (-not [string]::IsNullOrWhiteSpace($Control)) {
        Write-Section "guests/control [$target] -> $Control"
        Submit-GuestsControl $IOStreams $target $Control | Out-Null
        Write-Pass "control [$Control] submitted"
    }

    Write-Section 'disconnect'
    Submit-Disconnect $IOStreams | Out-Null
    Write-Pass 'disconnected'
}

Invoke-ScriptBlock -CommandPath $CommandPath -CommandArgs $CommandArgs -ScriptBlock ${function:Invoke-Pipeline}

Write-Host "`nTest run complete." -ForegroundColor Cyan
