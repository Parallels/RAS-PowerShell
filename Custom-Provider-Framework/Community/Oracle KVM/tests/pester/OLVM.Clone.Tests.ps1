#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }

<#
.SYNOPSIS
    guests/clone regression test: clone must reach a terminal state
    (completed or failed) even when the guest never reports an IP - it must
    NOT hang forever, unlike the Proxmox v4_ceph provider's bug #3.
.DESCRIPTION
    Live-farm tests. The underlying logic (Test-TaskExpired) already has fast,
    farm-independent unit coverage in OLVM.TaskState.Tests.ps1 - this file
    confirms the end-to-end behavior against a real OLVM farm, including the
    "no guest agent" case, which takes as long as the provider's own
    $script:TaskMaxAgeMinutes ceiling (30 minutes by default) to resolve.
    Budget accordingly; this is not a fast test.
#>

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

# NOTE ON SCOPING: this same setup runs both bare, at the top level of this
# file, and again inside BeforeAll below - see the identical note in
# OLVM.Convert.Tests.ps1 (Pester evaluates -Skip during Discovery, before any
# BeforeAll runs; a bare top-level assignment does not reliably persist into
# the later Run-phase execution of the It bodies themselves).
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding  = [System.Text.Encoding]::UTF8

Import-Module (Join-Path $PSScriptRoot 'TestHelpers.psm1') -Force

$script:Paths = Get-OlvmProviderPaths -TestScriptRoot $PSScriptRoot
Import-Module $script:Paths.ModulePath -Force

$script:Config = Get-OlvmTestConfig
$script:BaselineTemplateId     = [string]$script:Config.Templates.Baseline
$script:NoGuestAgentTemplateId = [string]$script:Config.Templates.NoGuestAgent

$script:NoAgentCloneTimeoutSeconds = 2100
if (Test-HasProperty -Object $script:Config -Name 'NoAgentCloneTimeoutSeconds') {
    $script:NoAgentCloneTimeoutSeconds = [int]$script:Config.NoAgentCloneTimeoutSeconds
}

Describe 'guests/clone must terminate (completed or failed), never hang forever' {

    BeforeAll {
        [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
        [Console]::InputEncoding  = [System.Text.Encoding]::UTF8

        Import-Module (Join-Path $PSScriptRoot 'TestHelpers.psm1') -Force

        $script:Paths = Get-OlvmProviderPaths -TestScriptRoot $PSScriptRoot
        Import-Module $script:Paths.ModulePath -Force

        $script:Config = Get-OlvmTestConfig
        $script:BaselineTemplateId     = [string]$script:Config.Templates.Baseline
        $script:NoGuestAgentTemplateId = [string]$script:Config.Templates.NoGuestAgent

        $script:NoAgentCloneTimeoutSeconds = 2100
        if (Test-HasProperty -Object $script:Config -Name 'NoAgentCloneTimeoutSeconds') {
            $script:NoAgentCloneTimeoutSeconds = [int]$script:Config.NoAgentCloneTimeoutSeconds
        }
    }

    It 'reaches completed with a clone_id when cloning a template with a guest agent, as a baseline control' -Skip:([string]::IsNullOrWhiteSpace($script:BaselineTemplateId)) {

        Invoke-OlvmSession -Config $script:Config -Body {
            param($IOStreams)

            $cloneId = $null

            try {
                $cloneResponse = Submit-GuestsClone $IOStreams $script:BaselineTemplateId "pester-clone-agent-$([guid]::NewGuid().ToString('N').Substring(0,8))" $null
                $cloneId = [string]$cloneResponse.clone_id

                $task = Wait-OlvmTask -IOStreams $IOStreams -TaskId ([string]$cloneResponse.task_id) -PollingSeconds $script:Config.TaskPollingSeconds -TimeoutSeconds $script:Config.TaskTimeoutSeconds
                $task.state | Should -Be 'completed'
                (Test-HasProperty -Object $task -Name 'output') | Should -BeTrue
                (Test-HasProperty -Object $task.output -Name 'clone_id') | Should -BeTrue
            }
            finally {
                if (-not [string]::IsNullOrWhiteSpace($cloneId)) {
                    try { Submit-GuestsControl $IOStreams $cloneId 'delete' | Out-Null } catch {}
                }
            }
        }
    }

    It 'reaches a terminal state (failed, via the task-age ceiling) within a bounded time when cloning a template WITHOUT a guest agent' -Skip:([string]::IsNullOrWhiteSpace($script:NoGuestAgentTemplateId)) {

        # Handle-TaskInfo's 'clone' branch waits for guest.ip_addresses to be
        # non-empty (read via the QEMU guest agent) before reporting
        # 'completed'. Without an agent that never happens, so on its own this
        # would loop 'running' forever - exactly the Proxmox v4_ceph provider's
        # confirmed bug #3. Test-TaskExpired is checked on every tasks/get
        # call before any type-specific polling, so this task is force-failed
        # once it exceeds $script:TaskMaxAgeMinutes (30 min default),
        # regardless of task type. This test intentionally waits that long -
        # it is proving there IS a ceiling, not that the clone becomes usable.
        Invoke-OlvmSession -Config $script:Config -Body {
            param($IOStreams)

            $cloneId = $null

            try {
                $cloneResponse = Submit-GuestsClone $IOStreams $script:NoGuestAgentTemplateId "pester-clone-noagent-$([guid]::NewGuid().ToString('N').Substring(0,8))" $null
                $cloneId = [string]$cloneResponse.clone_id

                $task = Wait-OlvmTask -IOStreams $IOStreams -TaskId ([string]$cloneResponse.task_id) -PollingSeconds $script:Config.TaskPollingSeconds -TimeoutSeconds $script:NoAgentCloneTimeoutSeconds
                $task.state | Should -Be 'failed' -Because 'a missing guest agent must not prevent RAS from ever seeing this clone reach a terminal state - Test-TaskExpired forces it to failed once $script:TaskMaxAgeMinutes elapses'
            }
            finally {
                if (-not [string]::IsNullOrWhiteSpace($cloneId)) {
                    try { Submit-GuestsControl $IOStreams $cloneId 'delete' | Out-Null } catch {}
                }
            }
        }
    }
}
