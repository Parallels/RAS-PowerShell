#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }

<#
.SYNOPSIS
    guests/convert regression test: converting to a template must not touch
    (and certainly not delete) snapshots the provider didn't create itself.
.DESCRIPTION
    Live-farm test. Requires TestConfig.Templates.Baseline (a template with a
    QEMU guest agent, used for the throwaway clone this test creates and
    deletes). Operates on a disposable clone, never on the template itself.

    Handle-GuestConvert in OLVM-CustomProvider.ps1 never calls the OLVM API at
    all - it only flips a locally persisted is_template flag - so this is
    expected to pass by construction. The test exists to confirm that end to
    end over the wire (task_id, tasks/get, and the actual OLVM snapshot list),
    not just by reading the source.
#>

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

# NOTE: config/template IDs are loaded at the top level (script scope), not
# inside BeforeAll. Pester evaluates an It block's -Skip expression during its
# Discovery pass, which runs before any BeforeAll - a $script: variable only
# ever set inside BeforeAll would not exist yet when -Skip is evaluated.
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding  = [System.Text.Encoding]::UTF8

Import-Module (Join-Path $PSScriptRoot 'TestHelpers.psm1') -Force

$script:Paths = Get-OlvmProviderPaths -TestScriptRoot $PSScriptRoot
Import-Module $script:Paths.ModulePath -Force

$script:Config = Get-OlvmTestConfig
$script:BaselineTemplateId = [string]$script:Config.Templates.Baseline

Describe 'guests/convert (must only manage RAS_TEMPLATE_VERSION_* snapshots)' {

    BeforeAll {
        # Re-run: a bare top-level assignment (needed above so -Skip has a
        # value during Discovery) does not reliably persist into the It body's
        # Run-phase execution - BeforeAll's assignment is what the It body
        # below actually sees.
        [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
        [Console]::InputEncoding  = [System.Text.Encoding]::UTF8

        Import-Module (Join-Path $PSScriptRoot 'TestHelpers.psm1') -Force

        $script:Paths = Get-OlvmProviderPaths -TestScriptRoot $PSScriptRoot
        Import-Module $script:Paths.ModulePath -Force

        $script:Config = Get-OlvmTestConfig
        $script:BaselineTemplateId = [string]$script:Config.Templates.Baseline
    }

    It 'preserves a non-RAS snapshot across a guests/convert to template' -Skip:([string]::IsNullOrWhiteSpace($script:BaselineTemplateId)) {

        # NOTE: everything the scriptblock below needs is either computed
        # inline or read via $script:, on purpose. A plain (non-$script:)
        # variable assigned in this It block would not be visible inside the
        # scriptblock at runtime: Invoke-OlvmSession/Invoke-ScriptBlock invoke
        # it several function-call frames deep, and PowerShell resolves
        # unqualified variable names in a scriptblock via the *caller's*
        # dynamic scope chain at invocation time, not this file's lexical
        # scope. $script: variables are the exception - they always resolve
        # to this file's script scope regardless of call depth.
        Invoke-OlvmSession -Config $script:Config -Body {
            param($IOStreams)

            $foreignSnapshotName = "manual-customer-snapshot-pester-$([guid]::NewGuid().ToString('N').Substring(0,8))"
            $cloneId = $null

            try {
                # 1. Create a disposable clone - never operate on the template itself.
                $cloneResponse = Submit-GuestsClone $IOStreams $script:BaselineTemplateId "pester-convert-$([guid]::NewGuid().ToString('N').Substring(0,8))" $null
                (Test-HasProperty -Object $cloneResponse -Name 'task_id') | Should -BeTrue
                (Test-HasProperty -Object $cloneResponse -Name 'clone_id') | Should -BeTrue

                $cloneId = [string]$cloneResponse.clone_id
                $cloneTask = Wait-OlvmTask -IOStreams $IOStreams -TaskId ([string]$cloneResponse.task_id) -PollingSeconds $script:Config.TaskPollingSeconds -TimeoutSeconds $script:Config.TaskTimeoutSeconds
                $cloneTask.state | Should -Be 'completed' -Because 'clone must finish before we can snapshot/convert it'

                # 2. Create a snapshot that does NOT follow the RAS naming
                #    convention, simulating a customer's own manual snapshot.
                $snapshotResponse = Submit-GuestsSnapshotsCreate $IOStreams $cloneId $foreignSnapshotName
                (Test-HasProperty -Object $snapshotResponse -Name 'task_id') | Should -BeTrue
                $snapshotTask = Wait-OlvmTask -IOStreams $IOStreams -TaskId ([string]$snapshotResponse.task_id) -PollingSeconds $script:Config.TaskPollingSeconds -TimeoutSeconds $script:Config.TaskTimeoutSeconds
                $snapshotTask.state | Should -Be 'completed'

                $existsBeforeConvert = Submit-GuestsSnapshotsExists $IOStreams $cloneId $foreignSnapshotName
                $existsBeforeConvert | Should -BeTrue

                # 3. Convert to template - this is the operation under test.
                $convertResponse = Submit-GuestsConvert $IOStreams $cloneId $true
                (Test-HasProperty -Object $convertResponse -Name 'task_id') | Should -BeTrue
                $convertTask = Wait-OlvmTask -IOStreams $IOStreams -TaskId ([string]$convertResponse.task_id) -PollingSeconds $script:Config.TaskPollingSeconds -TimeoutSeconds $script:Config.TaskTimeoutSeconds
                $convertTask.state | Should -Be 'completed'

                # 4. Assert both the label flipped AND the foreign snapshot survived.
                (Submit-GuestsGet $IOStreams $cloneId).is_template | Should -BeTrue

                $existsAfterConvert = Submit-GuestsSnapshotsExists $IOStreams $cloneId $foreignSnapshotName
                $existsAfterConvert | Should -BeTrue -Because 'guests/convert must only flip the local is_template flag, never touch OLVM snapshots'
            }
            finally {
                if (-not [string]::IsNullOrWhiteSpace($cloneId)) {
                    # Best-effort cleanup: guests/control delete is fire-and-forget
                    # in this provider (no task_id tracked for it), so this does
                    # not block the test on deletion completing.
                    try { Submit-GuestsControl $IOStreams $cloneId 'delete' | Out-Null } catch {}
                }
            }
        }
    }
}
