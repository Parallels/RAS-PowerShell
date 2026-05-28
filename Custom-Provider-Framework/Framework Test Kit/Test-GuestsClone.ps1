param (
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$GuestID,
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$CloneName,
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$SnapshotName
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding  = [System.Text.Encoding]::UTF8

$ProviderSettingsPath = Join-Path $PSScriptRoot 'CustomProvider.psd1'
$ProviderSettings = Import-PowerShellDataFile -Path $ProviderSettingsPath

$ProviderModulePath = Join-Path $PSScriptRoot 'CustomProvider.psm1'
Import-Module $ProviderModulePath -Force -Verbose

function Invoke-Pipeline {

    param([object]$IOStreams)

    Submit-InitializeAndConnect $IOStreams $providerSettings.CustomSettings | Out-Null
    Submit-GuestsClone $IOStreams $GuestID $CloneName $SnapshotName
}

Invoke-ScriptBlock -CommandPath $providerSettings.CommandPath -CommandArgs $ProviderSettings.CommandArgs -ScriptBlock ${function:Invoke-Pipeline}
