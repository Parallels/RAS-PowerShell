#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }

<#
.SYNOPSIS
    Unit-level regression tests for Test-TaskExpired, the ceiling that keeps a
    stuck tasks/get poll (most notably a clone that never gets a
    guest-agent-reported IP) from being reported as 'running' forever.
.DESCRIPTION
    Farm-independent - no OLVM connection needed. The provider script is a
    stdin read loop from its last line on, so it can't be dot-sourced directly
    without blocking; Get-ProviderFunctionScriptBlock extracts just the
    Test-TaskExpired function via the PowerShell AST and defines it in this
    file's scope for direct, fast unit testing of its classification logic.

    This is the equivalent Pester coverage to the Proxmox provider's
    New-TaskResultState tests, but for a different defect class: the Proxmox
    v4_ceph provider's clone task wait has NO age ceiling at all (confirmed
    live-farm hang without a QEMU guest agent, tracked there as bug #3). The
    OLVM provider's Handle-TaskInfo calls Test-TaskExpired before doing any
    type-specific polling, so every task type - clone included - is bounded
    by $script:TaskMaxAgeMinutes (30 by default), not just the clone path.
#>

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

BeforeAll {
    Import-Module (Join-Path $PSScriptRoot 'TestHelpers.psm1') -Force

    $script:Paths = Get-OlvmProviderPaths -TestScriptRoot $PSScriptRoot

    # Defines Test-TaskExpired in this scope only.
    . (Get-ProviderFunctionScriptBlock -ScriptPath $script:Paths.ProviderPath -FunctionName 'Test-TaskExpired')
}

Describe 'Test-TaskExpired (task age ceiling - prevents an unbounded tasks/get "running" loop)' {

    It 'returns $false when the context has no created_utc at all' {
        $ctx = @{ type = 'noop' }
        Test-TaskExpired -Ctx $ctx -MaxAgeMinutes 30 | Should -BeFalse
    }

    It 'returns $false for a task created 1 minute ago against a 30 minute ceiling' {
        $ctx = @{ created_utc = (Get-Date).ToUniversalTime().AddMinutes(-1).ToString('o') }
        Test-TaskExpired -Ctx $ctx -MaxAgeMinutes 30 | Should -BeFalse
    }

    It 'returns $true for a task created 31 minutes ago against a 30 minute ceiling' {
        $ctx = @{ created_utc = (Get-Date).ToUniversalTime().AddMinutes(-31).ToString('o') }
        Test-TaskExpired -Ctx $ctx -MaxAgeMinutes 30 | Should -BeTrue -Because 'this is what stops a clone stuck waiting for a guest-agent IP from polling "running" forever'
    }

    It 'returns $false right at the ceiling (age not strictly greater than MaxAgeMinutes)' {
        $ctx = @{ created_utc = (Get-Date).ToUniversalTime().AddSeconds(-1).ToString('o') }
        Test-TaskExpired -Ctx $ctx -MaxAgeMinutes 30 | Should -BeFalse
    }

    It 'does not throw under StrictMode when created_utc is garbage, and treats it as not expired' {
        $ctx = @{ created_utc = 'not-a-date' }
        { Test-TaskExpired -Ctx $ctx -MaxAgeMinutes 30 } | Should -Not -Throw
        Test-TaskExpired -Ctx $ctx -MaxAgeMinutes 30 | Should -BeFalse
    }

    It 'returns $false when Ctx itself is $null' {
        Test-TaskExpired -Ctx $null -MaxAgeMinutes 30 | Should -BeFalse
    }
}
