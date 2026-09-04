# Shared helpers for the OLVM Custom Provider Pester suite.
#
# These helpers sit on top of CustomProvider.psm1's transport (Invoke-ScriptBlock,
# Submit-*) and add only what the shared module doesn't already expose:
#   - test config loading
#   - a raw JSON-RPC request for methods the module keeps private
#     (Submit-Initialize / Submit-Connect are not exported)
#   - StrictMode-safe property presence checks
#   - task polling with a real timeout (Invoke-AsyncTask from the shared module
#     polls forever; tests need to fail instead of hang)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

function Test-HasProperty {
    <#
    .SYNOPSIS
        StrictMode-safe check for whether $Object has a property named $Name.
    .NOTES
        Handles both PSCustomObject (what ConvertFrom-Json produces for real
        JSON-RPC responses over the wire) and Hashtable (what the provider's
        own functions, e.g. Test-TaskExpired, return/accept when called
        directly in-process for unit-level tests) - a Hashtable's
        .PSObject.Properties reflects the Hashtable type's own members (Keys,
        Values, Count, ...), not its entries, so it needs its own ContainsKey
        check. Also never dots into .PSObject.Properties.Name directly: that
        throws under Set-StrictMode when $Object has zero properties at all
        (e.g. JSON "{}"), not just when the specific key is missing.
    #>
    param(
        [Parameter(Mandatory = $false)]
        [object]$Object,
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if ($null -eq $Object) {
        return $false
    }

    if ($Object -is [System.Collections.IDictionary]) {
        return $Object.ContainsKey($Name)
    }

    if ($Object -isnot [System.Management.Automation.PSCustomObject]) {
        return $false
    }

    foreach ($prop in $Object.PSObject.Properties) {
        if ($prop.Name -eq $Name) { return $true }
    }
    return $false
}

function Get-OlvmProviderPaths {
    <#
    .SYNOPSIS
        Resolves the repo-root CustomProvider.psm1 and the OLVM provider
        script from a test file's $PSScriptRoot, the same way the existing
        Test-*.ps1 scripts resolve CustomProvider.psd1/.psm1, but one folder
        deeper (tests/pester instead of tests).
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$TestScriptRoot
    )

    $providerDir = Split-Path (Split-Path $TestScriptRoot -Parent) -Parent
    $repoRoot    = Split-Path $providerDir -Parent

    return @{
        RepoRoot     = $repoRoot
        ModulePath   = Join-Path $repoRoot 'CustomProvider.psm1'
        ProviderDir  = $providerDir
        ProviderPath = Join-Path $providerDir 'OLVM-CustomProvider.ps1'
    }
}

function Get-OlvmTestConfig {
    <#
    .SYNOPSIS
        Loads tests/pester/TestConfig.psd1 (not checked in; see
        TestConfig.psd1.sample and ../../TESTING.md).
    #>
    param(
        [Parameter(Mandatory = $false)]
        [string]$Path = (Join-Path $PSScriptRoot 'TestConfig.psd1')
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Test config not found at '$Path'. Copy TestConfig.psd1.sample to TestConfig.psd1 in the same folder and fill in your test farm's details (see ../../TESTING.md)."
    }

    $config = Import-PowerShellDataFile -Path $Path

    foreach ($required in @('CommandPath', 'CommandArgs', 'CustomSettings')) {
        if (-not (Test-HasProperty -Object $config -Name $required)) {
            throw "Test config '$Path' is missing required key '$required'."
        }
    }

    return $config
}

function Invoke-RawProviderRequest {
    <#
    .SYNOPSIS
        Sends one JSON-RPC request and returns its 'result', for methods
        CustomProvider.psm1 does not export a Submit-* wrapper for
        (currently: provider/initialize, provider/connect on their own).
    .NOTES
        Mirrors Write-QueryObject/Read-ResultObject in CustomProvider.psm1.
        Kept local to the test suite rather than reaching into the module's
        private functions, since Export-ModuleMember does not expose them.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [object]$IOStreams,
        [Parameter(Mandatory = $true)]
        [string]$Method,
        [Parameter(Mandatory = $false)]
        [object]$Params
    )

    $query = @{ method = $Method }
    if ($null -ne $Params) {
        $query.params = $Params
    }

    $json = $query | ConvertTo-Json -Compress -Depth 10
    $IOStreams.StandardInput.WriteLine($json)

    $line = $IOStreams.StandardOutput.ReadLine()
    if ([string]::IsNullOrEmpty($line)) {
        throw "Empty response to '$Method'"
    }

    $response = $line.Trim() | ConvertFrom-Json -ErrorAction Stop

    if (Test-HasProperty -Object $response -Name 'error') {
        throw "'$Method' returned an error: $($response.error | ConvertTo-Json -Compress)"
    }

    if (-not (Test-HasProperty -Object $response -Name 'result')) {
        throw "Missing 'result' in response to '$Method': $line"
    }

    return $response.result
}

function Wait-OlvmTask {
    <#
    .SYNOPSIS
        Polls tasks/get until it stops returning 'running', with a real
        timeout (unlike CustomProvider.psm1's Invoke-AsyncTask, which polls
        forever - not acceptable inside a test run).
    #>
    param(
        [Parameter(Mandatory = $true)]
        [object]$IOStreams,
        [Parameter(Mandatory = $true)]
        [string]$TaskId,
        [Parameter(Mandatory = $false)]
        [int]$PollingSeconds = 5,
        [Parameter(Mandatory = $false)]
        [int]$TimeoutSeconds = 900
    )

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)

    while ($true) {
        $task = Submit-TasksGet $IOStreams $TaskId

        if (-not (Test-HasProperty -Object $task -Name 'state')) {
            throw "tasks/get response for task '$TaskId' is missing 'state': $($task | ConvertTo-Json -Compress)"
        }

        if ($task.state -ne 'running') {
            return $task
        }

        if ((Get-Date) -ge $deadline) {
            throw "Timed out after $TimeoutSeconds s waiting for task '$TaskId' to leave 'running' state."
        }

        Start-Sleep -Seconds $PollingSeconds
    }
}

function Get-ProviderFunctionScriptBlock {
    <#
    .SYNOPSIS
        Extracts a single named function's definition out of a provider .ps1
        file via the PowerShell AST, without executing the rest of the script
        (the provider script is a stdin read loop from the last line on -
        dot-sourcing it directly would block forever waiting for JSON-RPC
        input).
    .DESCRIPTION
        Used for narrow, farm-independent unit tests of pure logic functions
        (e.g. Test-TaskExpired's age classification). Dot-source the returned
        scriptblock at the call site to define the function in that scope:
        `. (Get-ProviderFunctionScriptBlock -ScriptPath $p -FunctionName 'Foo')`.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$ScriptPath,
        [Parameter(Mandatory = $true)]
        [string]$FunctionName
    )

    if (-not (Test-Path -LiteralPath $ScriptPath)) {
        throw "Provider script not found at '$ScriptPath'."
    }

    $tokens = $null
    $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$tokens, [ref]$parseErrors)

    if ($null -ne $parseErrors -and @($parseErrors).Count -gt 0) {
        throw "'$ScriptPath' does not parse cleanly: $($parseErrors | Out-String)"
    }

    $functionAst = $ast.Find(
        { param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $FunctionName },
        $true
    )

    if ($null -eq $functionAst) {
        throw "Function '$FunctionName' was not found in '$ScriptPath'. It may have been renamed or refactored - update this test to match."
    }

    return [ScriptBlock]::Create($functionAst.Extent.Text)
}

function Invoke-OlvmSession {
    <#
    .SYNOPSIS
        Starts the provider process, runs provider/initialize + provider/connect,
        invokes $Body with the open IOStreams, then always disconnects and stops
        the process (via CustomProvider.psm1's Invoke-ScriptBlock).
    .NOTES
        Not suitable for the provider/disconnect regression test, which needs
        explicit control over when disconnect happens within one process
        lifetime - that test calls Invoke-ScriptBlock directly instead.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [object]$Config,
        [Parameter(Mandatory = $true)]
        [scriptblock]$Body
    )

    $session = {
        param($IOStreams)

        Submit-InitializeAndConnect $IOStreams $Config.CustomSettings | Out-Null

        try {
            & $Body $IOStreams
        }
        finally {
            Submit-Disconnect $IOStreams | Out-Null
        }
    }.GetNewClosure()

    Invoke-ScriptBlock -CommandPath $Config.CommandPath -CommandArgs $Config.CommandArgs -ScriptBlock $session
}

Export-ModuleMember -Function `
    Test-HasProperty, `
    Get-OlvmProviderPaths, `
    Get-OlvmTestConfig, `
    Invoke-RawProviderRequest, `
    Wait-OlvmTask, `
    Get-ProviderFunctionScriptBlock, `
    Invoke-OlvmSession
