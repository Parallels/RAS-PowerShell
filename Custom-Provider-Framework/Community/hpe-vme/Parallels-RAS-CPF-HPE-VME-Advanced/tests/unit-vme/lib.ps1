# Shared harness for the VME suites (tests/unit-vme/vme-*.ps1).
#
# Every suite drives the REAL provider as a subprocess over stdin/stdout JSON-RPC,
# exactly like RAS does, against mock_vme.py over real HTTPS. Assertions use the
# mock's per-endpoint request counters (/_ctl/stats) for call budgets and the
# provider's own log for behaviour that has no RPC-visible effect.
#
# Dot-source it:  . (Join-Path $PSScriptRoot 'lib.ps1')

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$script:pass = 0
$script:fail = 0
$script:ProviderSource = (Resolve-Path (Join-Path $PSScriptRoot '../../Parallels-RAS-CPF-HPE-VME-Advanced.ps1')).Path
$script:MockScript = (Resolve-Path (Join-Path $PSScriptRoot '../e2e-vme/mock_vme.py')).Path
$script:MockPort = 8445
$script:Providers = [System.Collections.Generic.List[object]]::new()

function Assert {
    param([bool]$Cond, [string]$Msg)
    if ($Cond) { $script:pass++; Write-Host "PASS: $Msg" -ForegroundColor Green }
    else { $script:fail++; Write-Host "FAIL: $Msg" -ForegroundColor Red }
}

function Resolve-Python3Command {
    foreach ($name in @('python3', 'python')) {
        $cmd = Get-Command $name -ErrorAction SilentlyContinue
        if ($null -ne $cmd) { return $cmd.Source }
    }
    throw "No Python 3 interpreter found on PATH."
}

$script:MockProcess = $null
function Stop-VmeMock {
    param([int]$Port = $script:MockPort)
    if ($null -ne $script:MockProcess) { try { $script:MockProcess.Kill($true) } catch { }; $script:MockProcess = $null }
    if ($IsWindows) {
        Get-CimInstance Win32_Process -ErrorAction SilentlyContinue |
            Where-Object { $_.CommandLine -and $_.CommandLine -like "*mock_vme.py*$Port*" } |
            ForEach-Object { try { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue } catch {} }
    }
    else {
        foreach ($procId in @(& pgrep -f "mock_vme.py $Port" 2>$null)) { if ($procId) { & kill $procId 2>$null } }
    }
    Start-Sleep -Milliseconds 300
}

function Invoke-Mock {
    param([string]$Path, [string]$Method = 'POST')
    return Invoke-RestMethod -Uri "https://127.0.0.1:$($script:MockPort)$Path" -Method $Method -SkipCertificateCheck -TimeoutSec 60
}

function Start-VmeMock {
    param([int]$Count = 41)
    Stop-VmeMock
    $err = Join-Path $PSScriptRoot "run/mock-$($script:MockPort).err"
    New-Item -ItemType Directory -Path (Split-Path $err) -Force | Out-Null
    # A Process with all three streams redirected: the mock must not inherit this
    # shell's stdout, or a caller piping this suite's output waits on the mock forever.
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = Resolve-Python3Command
    $psi.Arguments = "`"$($script:MockScript)`" $($script:MockPort)"
    $psi.UseShellExecute = $false
    $psi.RedirectStandardInput = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true
    $script:MockProcess = [System.Diagnostics.Process]::Start($psi)
    $deadline = (Get-Date).AddSeconds(15)
    while ((Get-Date) -lt $deadline) {
        try { Invoke-Mock "/_ctl/reset?count=$Count" | Out-Null; return } catch { Start-Sleep -Milliseconds 200 }
    }
    throw "mock_vme.py did not come up on :$($script:MockPort) -- see $err"
}

function Get-MockStats { return (Invoke-Mock '/_ctl/stats' 'GET') }
function Reset-MockStats { Invoke-Mock '/_ctl/stats/reset' | Out-Null }
function Get-MockCount {
    param($Stats, [string]$Key)
    $p = $Stats.requests.PSObject.Properties[$Key]
    if ($null -eq $p) { return 0 }
    return [int]$p.Value
}
function Get-MockTotal {
    param($Stats)
    $t = 0
    foreach ($p in $Stats.requests.PSObject.Properties) { $t += [int]$p.Value }
    return $t
}

# A fresh provider folder with the provider script and an optional settings file. $Settings
# is a nested hashtable of bare values (the loader accepts bare values and rewrites
# them into the rich shape -- which is itself under test).
function New-ProviderDir {
    param([string]$Name, [hashtable]$Settings = $null, [string]$RawSettingsJson = $null)
    $dir = Join-Path $PSScriptRoot "run/$Name"
    if (Test-Path $dir) { Remove-Item $dir -Recurse -Force }
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
    Copy-Item -LiteralPath $script:ProviderSource -Destination (Join-Path $dir 'provider.ps1')
    if ($RawSettingsJson) { Set-Content -LiteralPath (Join-Path $dir 'VME-RAS-Config.json') -Value $RawSettingsJson -Encoding utf8NoBOM }
    elseif ($null -ne $Settings) {
        $doc = @{ schema_version = 2 } + $Settings
        Set-Content -LiteralPath (Join-Path $dir 'VME-RAS-Config.json') -Value ($doc | ConvertTo-Json -Depth 10) -Encoding utf8NoBOM
    }
    return $dir
}

function Start-Provider {
    param([string]$Dir, [string]$Culture = $null)
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = (Get-Command pwsh).Source
    $script = Join-Path $Dir 'provider.ps1'
    if ($Culture) {
        $psi.Arguments = "-NoProfile -Command `"[System.Globalization.CultureInfo]::CurrentCulture='$Culture'; [System.Globalization.CultureInfo]::DefaultThreadCurrentCulture='$Culture'; & '$script'`""
    }
    else { $psi.Arguments = "-NoProfile -File `"$script`"" }
    $psi.WorkingDirectory = $Dir
    $psi.RedirectStandardInput = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.UseShellExecute = $false
    $psi.StandardOutputEncoding = New-Object System.Text.UTF8Encoding($false)
    $p = [System.Diagnostics.Process]::Start($psi)
    $obj = [pscustomobject]@{ Process = $p; Dir = $Dir; Log = (Join-Path $Dir 'VME-RAS-Provider.log'); FirstLine = $null }
    $script:Providers.Add($obj)
    return $obj
}

function Send-RpcRaw {
    param($P, [string]$Json, [int]$TimeoutSec = 60)
    $P.Process.StandardInput.WriteLine($Json)
    $P.Process.StandardInput.Flush()
    $task = $P.Process.StandardOutput.ReadLineAsync()
    if (-not $task.Wait($TimeoutSec * 1000)) { throw "No reply within ${TimeoutSec}s to: $Json" }
    $line = $task.Result
    if ($null -eq $line) { throw "Provider closed stdout while awaiting a reply to: $Json" }
    if ($null -eq $P.FirstLine) { $P.FirstLine = $line }
    return $line
}

function Send-Rpc {
    param($P, [string]$Json, [int]$TimeoutSec = 60)
    return ((Send-RpcRaw -P $P -Json $Json -TimeoutSec $TimeoutSec) | ConvertFrom-Json)
}

function Get-Result {
    param($Response, [string]$What)
    if ($Response.PSObject.Properties.Name -contains 'error') { throw "$What returned an error: $($Response.error.message)" }
    return $Response.result
}

function Test-IsError { param($Response) return ($Response.PSObject.Properties.Name -contains 'error') }

function Connect-Provider {
    param($P)
    Get-Result (Send-Rpc $P '{"method":"provider/initialize"}') 'initialize' | Out-Null
    return Get-Result (Send-Rpc $P "{`"method`":`"provider/connect`",`"params`":{`"settings`":{`"host`":`"127.0.0.1:$($script:MockPort)`",`"username`":`"admin`",`"password`":`"password`",`"ssl_check`":false}}}") 'connect'
}

function Get-Guest { param($P, [string]$Id) return (Send-Rpc $P "{`"method`":`"guests/get`",`"params`":{`"id`":`"$Id`"}}") }
function Get-GuestList { param($P) return @((Get-Result (Send-Rpc $P '{"method":"guests/list"}') 'guests/list').guests) }
function Get-Task { param($P, [string]$TaskId) return (Get-Result (Send-Rpc $P "{`"method`":`"tasks/get`",`"params`":{`"id`":`"$TaskId`"}}") 'tasks/get') }

function Wait-Task {
    param($P, [string]$TaskId, [int]$TimeoutSec = 60, [double]$IntervalSec = 1)
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    while ((Get-Date) -lt $deadline) {
        $t = Get-Task $P $TaskId
        if ($t.state -ne 'running') { return $t }
        Start-Sleep -Milliseconds ([int]($IntervalSec * 1000))
    }
    return $null
}

# The provider keeps its log open for writing, so the reader must share write access.
function Get-LogText {
    param($P)
    if (-not (Test-Path $P.Log)) { return '' }
    $fs = [System.IO.FileStream]::new($P.Log, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete)
    try { return ([System.IO.StreamReader]::new($fs)).ReadToEnd() } finally { $fs.Dispose() }
}

function Stop-Provider {
    param($P)
    try { $P.Process.StandardInput.Close() } catch { }
    try { if (-not $P.Process.WaitForExit(10000)) { $P.Process.Kill() } } catch { }
}

function Complete-Suite {
    param([string]$Name)
    foreach ($p in $script:Providers) { if (-not $p.Process.HasExited) { Stop-Provider $p } }
    Stop-VmeMock
    Write-Host ""
    Write-Host "=== $Name Summary: $($script:pass) passed, $($script:fail) failed ===" -ForegroundColor $(if ($script:fail -eq 0) { 'Green' } else { 'Red' })
    if ($script:fail -gt 0) { exit 1 }
}
