<#
.SYNOPSIS
    End-to-end test for the XCP-ng custom provider.
.DESCRIPTION
    Drives Parallels-RAS-CFP-XCP-ng.ps1 through the shared CustomProvider.psm1
    test harness. By default it runs a safe, read-only flow:

        provider/initialize -> provider/connect -> guests/list -> guests/get -> provider/disconnect

    Mutating operations are opt-in:
      -Control <action>   runs guests/control on the target guest
      -TestSnapshots      runs a snapshot lifecycle: create -> exists -> delete

    Connection settings come from -Server/-Username/-Password, or fall back to a
    CustomProvider.psd1 found next to this script or in the repository root.
.EXAMPLE
    pwsh -File .\Test-XCPng.ps1 -Server https://xcp-pool.example.com -Username root -Password $env:XCP_PASS
.EXAMPLE
    pwsh -File .\Test-XCPng.ps1 -Server https://xcp-pool.example.com -Username root -Password $env:XCP_PASS -GuestID <vm-uuid> -TestSnapshots
#>

[CmdletBinding()]
param(
    [string]$Server,
    [string]$Username,
    [string]$Password,
    [bool]$SkipTls = $true,
    [string]$GuestID,
    [ValidateSet('start', 'stop', 'restart', 'reset', 'suspend', 'delete')]
    [string]$Control,
    [switch]$TestSnapshots,
    [string]$SnapshotName = 'RAS_TEMPLATE_VERSION_1',
    [int]$PollingRate = 5
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding  = [System.Text.Encoding]::UTF8

$ProviderDir = Split-Path $PSScriptRoot -Parent           # xcp-ng/
$RepoRoot    = Split-Path $ProviderDir -Parent            # repository root

$ProviderScript = Join-Path $ProviderDir 'Parallels-RAS-CFP-XCP-ng.ps1'
if (-not (Test-Path $ProviderScript)) { throw "Provider script not found: $ProviderScript" }

$ModulePath = Join-Path $RepoRoot 'CustomProvider.psm1'
if (-not (Test-Path $ModulePath)) { throw "Harness module not found: $ModulePath" }
Import-Module $ModulePath -Force

function Resolve-CustomSettings {
    if (-not [string]::IsNullOrWhiteSpace($Server) -and
        -not [string]::IsNullOrWhiteSpace($Username) -and
        -not [string]::IsNullOrWhiteSpace($Password)) {
        return @{ host = $Server; username = $Username; password = $Password; skip_tls = $SkipTls }
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

    throw 'Provide -Server, -Username and -Password, or a CustomProvider.psd1 with CustomSettings.'
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

    if ($TestSnapshots) {
        Write-Section "snapshot lifecycle [$target] / $SnapshotName"

        Invoke-AsyncTask $IOStreams $PollingRate {
            Submit-GuestsSnapshotsCreate $IOStreams $target $SnapshotName
        } | Out-Null
        Write-Pass 'snapshot created'

        $exists = Submit-GuestsSnapshotsExists $IOStreams $target $SnapshotName
        Write-Host "  exists => $exists"
        if (-not $exists) { throw 'Snapshot reported as not existing after creation.' }
        Write-Pass 'snapshot exists'

        Invoke-AsyncTask $IOStreams $PollingRate {
            Submit-GuestsSnapshotsDelete $IOStreams $target $SnapshotName
        } | Out-Null
        Write-Pass 'snapshot deleted'
    }

    Write-Section 'disconnect'
    Submit-Disconnect $IOStreams | Out-Null
    Write-Pass 'disconnected'
}

Invoke-ScriptBlock -CommandPath $CommandPath -CommandArgs $CommandArgs -ScriptBlock ${function:Invoke-Pipeline}

Write-Host "`nTest run complete." -ForegroundColor Cyan
