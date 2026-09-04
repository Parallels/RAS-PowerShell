param (
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$GuestID,
    [Parameter(Mandatory = $false)]
    [int]$TemplateVersionID
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding  = [System.Text.Encoding]::UTF8

$ProviderSettingsPath = Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) 'CustomProvider.psd1'
$ProviderSettings = Import-PowerShellDataFile -Path $ProviderSettingsPath

$ProviderModulePath = Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) 'CustomProvider.psm1'
Import-Module $ProviderModulePath -Force -Verbose

function Invoke-Pipeline {

    param([object]$IOStreams)

    $result = Submit-InitializeAndConnect $IOStreams $providerSettings.CustomSettings
    $capbilities = $result.capabilities

    if ("basic" -ne $capbilities.template_method -and "versioning" -ne $capbilities.template_method) {
        throw "The provider does not support templates"
    }

    if ((Submit-GuestsGet $IOStreams $GuestID).is_template) {
        throw "$GuestID is a template"
    }

    $snapshotName = ""
    if ("versioning" -eq $capbilities.template_method) {
        if ($TemplateVersionID -le 0) {
            throw "The template version ID must be greater than 0"
        }
        $snapshotName = "RAS_TEMPLATE_VERSION_{0}" -f $TemplateVersionID
    } elseif ($capbilities.can_link_clones) {
        $snapshotName = "RAS Template Snapshot"
    }

    if ("powered_off" -ne (Submit-GuestsGet $IOStreams $GuestID).state) {
        Submit-GuestsControl $IOStreams $GuestID "stop" | Out-Null
    }

    while ("powered_off" -ne (Submit-GuestsGet $IOStreams $GuestID).state) {
        Start-Sleep -Seconds $capbilities.guests_polling_rate
    }

    if ($snapshotName) {
        if (Submit-GuestsSnapshotsExists $IOStreams $GuestID $snapshotName) {
            Invoke-AsyncTask $IOStreams $capbilities.tasks_polling_rate -ScriptBlock {
                return Submit-GuestsSnapshotsDelete $IOStreams $GuestID $snapshotName
            } | Out-Null
        }

        Invoke-AsyncTask $IOStreams $capbilities.tasks_polling_rate -ScriptBlock {
            return Submit-GuestsSnapshotsCreate $IOStreams $GuestID $snapshotName
        } | Out-Null
    }

    Invoke-AsyncTask $IOStreams $capbilities.tasks_polling_rate -ScriptBlock {
        return Submit-GuestsConvert -IOStreams $IOStreams -GuestID $GuestID -IsTemplate $true
    } | Out-Null

    if (-not (Submit-GuestsGet $IOStreams $GuestID).is_template) {
        throw "$GuestID is not a template"
    }
}

Invoke-ScriptBlock -CommandPath $providerSettings.CommandPath -CommandArgs $ProviderSettings.CommandArgs -ScriptBlock ${function:Invoke-Pipeline}
