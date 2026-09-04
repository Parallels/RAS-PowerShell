#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }

<#
.SYNOPSIS
    provider/initialize capability regression tests for OLVM-CustomProvider.ps1.
    Requires a live test farm to spawn the provider process, but no
    provider/connect call is made - provider/initialize never touches OLVM.
#>

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

BeforeAll {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    [Console]::InputEncoding  = [System.Text.Encoding]::UTF8

    Import-Module (Join-Path $PSScriptRoot 'TestHelpers.psm1') -Force

    $script:Paths = Get-OlvmProviderPaths -TestScriptRoot $PSScriptRoot
    Import-Module $script:Paths.ModulePath -Force

    $script:Config = Get-OlvmTestConfig
}

Describe 'provider/initialize capabilities (OLVM)' {

    It 'reports can_suspend_guests = $true and template_method = versioning' {
        $initializeResult = Invoke-ScriptBlock -CommandPath $script:Config.CommandPath -CommandArgs $script:Config.CommandArgs -ScriptBlock {
            param($IOStreams)
            Invoke-RawProviderRequest -IOStreams $IOStreams -Method 'provider/initialize'
        }

        $initializeResult | Should -Not -BeNullOrEmpty
        (Test-HasProperty -Object $initializeResult -Name 'capabilities') | Should -BeTrue

        $capabilities = $initializeResult.capabilities

        (Test-HasProperty -Object $capabilities -Name 'can_suspend_guests') | Should -BeTrue
        $capabilities.can_suspend_guests | Should -BeTrue -Because 'guests/control suspend maps to the OLVM suspend action'

        (Test-HasProperty -Object $capabilities -Name 'template_method') | Should -BeTrue
        $capabilities.template_method | Should -Be 'versioning' -Because 'template versions are named RAS_TEMPLATE_VERSION_* snapshots, matching the Proxmox provider convention'
    }

    It 'reports can_link_clones = $false, matching the actual clone implementation' {
        $initializeResult = Invoke-ScriptBlock -CommandPath $script:Config.CommandPath -CommandArgs $script:Config.CommandArgs -ScriptBlock {
            param($IOStreams)
            Invoke-RawProviderRequest -IOStreams $IOStreams -Method 'provider/initialize'
        }

        $capabilities = $initializeResult.capabilities

        (Test-HasProperty -Object $capabilities -Name 'can_link_clones') | Should -BeTrue
        $capabilities.can_link_clones | Should -BeFalse -Because 'Handle-GuestClone always requests a full clone (clone=true, no is_link_clone handling) - advertising $true here would let RAS request a linked clone the provider cannot honor'
    }

    It 'reports positive numeric polling rates and a positive task retry ceiling' {
        $initializeResult = Invoke-ScriptBlock -CommandPath $script:Config.CommandPath -CommandArgs $script:Config.CommandArgs -ScriptBlock {
            param($IOStreams)
            Invoke-RawProviderRequest -IOStreams $IOStreams -Method 'provider/initialize'
        }

        $capabilities = $initializeResult.capabilities

        (Test-HasProperty -Object $capabilities -Name 'guests_polling_rate') | Should -BeTrue
        [int]$capabilities.guests_polling_rate | Should -BeGreaterThan 0

        (Test-HasProperty -Object $capabilities -Name 'tasks_polling_rate') | Should -BeTrue
        [int]$capabilities.tasks_polling_rate | Should -BeGreaterThan 0

        (Test-HasProperty -Object $capabilities -Name 'tasks_polling_retries') | Should -BeTrue
        [int]$capabilities.tasks_polling_retries | Should -BeGreaterThan 0
    }
}
