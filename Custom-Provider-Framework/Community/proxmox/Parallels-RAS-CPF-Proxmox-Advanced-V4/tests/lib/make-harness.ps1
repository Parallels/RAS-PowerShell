#Requires -Version 7.0
<#
    Regenerates unit/funcs2.ps1 -- the dot-sourceable projection of the live Proxmox
    provider used by the unit suites (t3-t9).

    The provider is a single file that ends in a blocking `while ($true)` stdin loop,
    so it cannot simply be dot-sourced: doing so would hang the test process forever.
    This truncates the script immediately before that loop, which yields every function,
    every $script: variable and the whole startup sequence, and nothing that blocks.

    Run this after ANY edit to the provider. Testing a stale projection has already
    cost this project real time -- a suite that passes against last week's copy of the
    script tells you nothing about the one you are shipping. run-all.ps1 calls it on
    every run for exactly that reason; call it yourself if you run a suite directly.
#>

param(
    [string]$ProviderPath = (Join-Path $PSScriptRoot '../../Parallels-RAS-CPF-Proxmox-Advanced.ps1'),
    [string]$OutPath      = (Join-Path $PSScriptRoot '../unit/funcs2.ps1')
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$resolved = (Resolve-Path -LiteralPath $ProviderPath).Path
$lines = Get-Content -LiteralPath $resolved

# The main loop is the only `while ($true)` at column 0; the two inside functions are
# indented. Anchoring on '^while (\$true)' is therefore unambiguous -- but verify it,
# because a silent mismatch here produces a harness that looks fine and tests nothing.
$loopIdx = -1
for ($i = 0; $i -lt $lines.Count; $i++) {
    if ($lines[$i] -match '^while \(\$true\)') { $loopIdx = $i; break }
}

if ($loopIdx -lt 0) {
    throw "make-harness: no top-level 'while (`$true)' main loop found in $resolved. The provider's structure changed -- update this script rather than working around it."
}

$lines[0..($loopIdx - 1)] | Set-Content -LiteralPath $OutPath -Encoding UTF8

$outFull = (Resolve-Path -LiteralPath $OutPath).Path
Write-Host "make-harness: $outFull" -ForegroundColor DarkGray
Write-Host "  from   : $resolved" -ForegroundColor DarkGray
Write-Host "  kept   : $loopIdx of $($lines.Count) lines (truncated at the stdin loop on line $($loopIdx + 1))" -ForegroundColor DarkGray
