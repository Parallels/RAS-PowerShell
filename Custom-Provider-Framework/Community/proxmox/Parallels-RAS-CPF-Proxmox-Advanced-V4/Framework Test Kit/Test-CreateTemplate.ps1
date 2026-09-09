param (
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$GuestID
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

    $result = Submit-InitializeAndConnect $IOStreams $providerSettings.CustomSettings
    $capbilities = $result.capabilities

    if ("basic" -ne $capbilities.template_method -and "versioning" -ne $capbilities.template_method) {
        throw "The provider does not support templates"
    }

    if ((Submit-GuestsGet $IOStreams $GuestID).is_template) {
        throw "$GuestID is already a template"
    }

    if ("powered_off" -ne (Submit-GuestsGet $IOStreams $GuestID).state) {
         Submit-GuestsControl $IOStreams $GuestID "stop" | Out-Null
    }

    while ("powered_off" -ne (Submit-GuestsGet $IOStreams $GuestID).state) {
        Start-Sleep -Seconds $capbilities.guests_polling_rate
    }

    $snapshotName = ""
    if ("versioning" -eq $capbilities.template_method) {
        $snapshotName = "RAS_TEMPLATE_VERSION_1"
    }
    elseif ($capbilities.can_link_clones) {
        $snapshotName = "RAS Template Snapshot"
    }

    if ($snapshotName) {
        Invoke-AsyncTask $IOStreams $capbilities.tasks_polling_rate -ScriptBlock {
            return Submit-GuestsSnapshotsCreate $IOStreams $GuestID $snapshotName
        } | Out-Null
    }

    Invoke-AsyncTask $IOStreams $capbilities.tasks_polling_rate -ScriptBlock {
        return Submit-GuestsConvert -IOStreams $IOStreams -GuestID $GuestID -IsTemplate $true
    } | Out-Null

    if (-not (Submit-GuestsGet $IOStreams $GuestID).is_template) {
        throw "$GuestID should be a template"
    }
}

Invoke-ScriptBlock -CommandPath $providerSettings.CommandPath -CommandArgs $ProviderSettings.CommandArgs -ScriptBlock ${function:Invoke-Pipeline}
