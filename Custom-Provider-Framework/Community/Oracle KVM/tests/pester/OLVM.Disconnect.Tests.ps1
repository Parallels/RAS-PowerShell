#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }

<#
.SYNOPSIS
    provider/disconnect regression test: disconnecting during an in-flight
    clone must NOT destroy that clone's tracking state - a later tasks/get on
    the same task_id must still be able to return clone_id.
.DESCRIPTION
    Live-farm test. Requires TestConfig.Templates.Baseline. This is the OLVM
    equivalent of the Proxmox v4_ceph provider's bug #4, but asserts the
    OPPOSITE expectation: Handle-Disconnect in OLVM-CustomProvider.ps1 only
    clears the in-process $script:TaskContext cache; it deliberately does NOT
    delete the on-disk OLVM-RAS-TaskState.json file, and Get-TaskStateEntry
    falls back to reading that file when the in-memory cache misses. So a
    disconnect/reconnect mid-clone - even across a brand-new process, which is
    the scenario that actually matters since RAS can launch a fresh process
    per request - does not lose clone bookkeeping. This test is NOT tagged
    KnownBug: it is expected to pass.
.NOTES
    Deliberately does not use the shared Invoke-OlvmSession helper (that
    helper always disconnects at the very end of the session) - this test
    needs to control exactly when disconnect/reconnect happen within one
    provider process lifetime.
#>

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

# NOTE: config/template IDs are loaded at the top level (script scope), not
# inside BeforeAll - see the identical note in OLVM.Convert.Tests.ps1
# (Pester evaluates -Skip during Discovery, before any BeforeAll runs).
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding  = [System.Text.Encoding]::UTF8

Import-Module (Join-Path $PSScriptRoot 'TestHelpers.psm1') -Force

$script:Paths = Get-OlvmProviderPaths -TestScriptRoot $PSScriptRoot
Import-Module $script:Paths.ModulePath -Force

$script:Config = Get-OlvmTestConfig
$script:BaselineTemplateId = [string]$script:Config.Templates.Baseline

Describe 'provider/disconnect (must not lose in-flight clone tracking state)' {

    BeforeAll {
        [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
        [Console]::InputEncoding  = [System.Text.Encoding]::UTF8

        Import-Module (Join-Path $PSScriptRoot 'TestHelpers.psm1') -Force

        $script:Paths = Get-OlvmProviderPaths -TestScriptRoot $PSScriptRoot
        Import-Module $script:Paths.ModulePath -Force

        $script:Config = Get-OlvmTestConfig
        $script:BaselineTemplateId = [string]$script:Config.Templates.Baseline
    }

    It 'still resolves clone_id via tasks/get after a disconnect/reconnect cycle mid-clone' -Skip:([string]::IsNullOrWhiteSpace($script:BaselineTemplateId)) {

        Invoke-ScriptBlock -CommandPath $script:Config.CommandPath -CommandArgs $script:Config.CommandArgs -ScriptBlock {
            param($IOStreams)

            Submit-InitializeAndConnect $IOStreams $script:Config.CustomSettings | Out-Null

            $cloneId = $null

            try {
                # 1. Start a clone and capture its task_id/clone_id, but do not
                #    wait for it to finish yet - the disconnect below must hit
                #    while the clone is still tracked as in-flight.
                $cloneResponse = Submit-GuestsClone $IOStreams $script:BaselineTemplateId "pester-disconnect-$([guid]::NewGuid().ToString('N').Substring(0,8))" $null
                (Test-HasProperty -Object $cloneResponse -Name 'task_id') | Should -BeTrue
                (Test-HasProperty -Object $cloneResponse -Name 'clone_id') | Should -BeTrue

                $taskId = [string]$cloneResponse.task_id
                $cloneId = [string]$cloneResponse.clone_id

                # 2. Disconnect, then reconnect within the SAME process - this is
                #    what RAS does across a settings refresh/reload without
                #    restarting the provider process.
                Submit-Disconnect $IOStreams | Out-Null
                Invoke-RawProviderRequest -IOStreams $IOStreams -Method 'provider/connect' -Params @{ settings = $script:Config.CustomSettings } | Out-Null

                # 3. Poll the ORIGINAL task_id to completion after reconnecting.
                $task = Wait-OlvmTask -IOStreams $IOStreams -TaskId $taskId -PollingSeconds $script:Config.TaskPollingSeconds -TimeoutSeconds $script:Config.TaskTimeoutSeconds

                $task.state | Should -Be 'completed' -Because 'the underlying OLVM clone operation itself is unaffected by provider/disconnect'

                (Test-HasProperty -Object $task -Name 'output') | Should -BeTrue
                (Test-HasProperty -Object $task.output -Name 'clone_id') | Should -BeTrue -Because 'Get-TaskStateEntry falls back to the on-disk OLVM-RAS-TaskState.json file, which Handle-Disconnect never deletes'
                [string]$task.output.clone_id | Should -Be $cloneId
            }
            finally {
                if (-not [string]::IsNullOrWhiteSpace($cloneId)) {
                    try { Submit-GuestsControl $IOStreams $cloneId 'delete' | Out-Null } catch {}
                }
                Submit-Disconnect $IOStreams | Out-Null
            }
        }
    }
}
