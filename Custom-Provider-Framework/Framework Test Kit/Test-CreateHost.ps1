param (
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$GuestID,
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$CloneName,
    [Parameter(Mandatory = $false)]
    [int]$TemplateVersionID
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

    if (-not (Submit-GuestsGet $IOStreams $GuestID).is_template) {
        throw "$GuestID is not a template"
    }
    
    $snapshotName = ""
    if ("versioning" -eq $capbilities.template_method) {
        if ($TemplateVersionID -le 0) {
            throw "The template version ID must be greater than 0"
        }
        $snapshotName = "RAS_TEMPLATE_VERSION_{0}" -f $TemplateVersionID 
    }
    elseif ($capbilities.can_link_clones) {
        $snapshotName = "RAS Template Snapshot"
    }

    $cloneID = (Invoke-AsyncTask $IOStreams $capbilities.tasks_polling_rate -ScriptBlock {
        return Submit-GuestsClone $IOStreams $GuestID $CloneName $SnapshotName
    }).clone_id

    if ("powered_on" -ne (Submit-GuestsGet $IOStreams $cloneID).state) {
        Submit-GuestsControl $IOStreams $cloneID "start" | Out-Null
    }

    while ("powered_on" -ne (Submit-GuestsGet $IOStreams $cloneID).state) {
        Start-Sleep -Seconds $capbilities.guests_polling_rate
    }
}

Invoke-ScriptBlock -CommandPath $providerSettings.CommandPath -CommandArgs $ProviderSettings.CommandArgs -ScriptBlock ${function:Invoke-Pipeline}
