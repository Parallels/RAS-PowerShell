#Requires -Version 7.0
<#
    Runs every suite in this folder and prints one summary table.

    Each suite starts and stops its own mock VME appliance (e2e-vme/mock_vme.py) and
    copies its own provider, so this is safe to run repeatedly with nothing set up by
    hand.

        pwsh -File tests/run-all.ps1
        pwsh -Command "& ./tests/run-all.ps1 -Only unit-vme"
        pwsh -Command "& ./tests/run-all.ps1 -Only vme-u3,e2e-vme"

    Exit code is non-zero if any suite fails, so it is usable as a gate.
#>

param(
    [string[]]$Only
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$suites = @(
    @{ Group = 'unit-vme'; Name = 'vme-u1';   Path = 'unit-vme/vme-u1-inventory.ps1' }
    @{ Group = 'unit-vme'; Name = 'vme-u2';   Path = 'unit-vme/vme-u2-listing-safety.ps1' }
    @{ Group = 'unit-vme'; Name = 'vme-u3';   Path = 'unit-vme/vme-u3-clone.ps1' }
    @{ Group = 'unit-vme'; Name = 'vme-u4';   Path = 'unit-vme/vme-u4-control-maintenance.ps1' }
    @{ Group = 'unit-vme'; Name = 'vme-u5';   Path = 'unit-vme/vme-u5-transport-settings.ps1' }
    @{ Group = 'unit-vme'; Name = 'vme-u6';   Path = 'unit-vme/vme-u6-clone-slots.ps1' }
    @{ Group = 'unit-vme'; Name = 'vme-u7';   Path = 'unit-vme/vme-u7-versioning.ps1' }
    @{ Group = 'unit-vme'; Name = 'vme-u8';   Path = 'unit-vme/vme-u8-live-fixes.ps1' }
    @{ Group = 'soak-vme'; Name = 'vme-soak'; Path = 'unit-vme/vme-soak.ps1' }
    @{ Group = 'e2e-vme';  Name = 'e2e-vme';  Path = 'e2e-vme/run_e2e.ps1' }
)

if ($Only) { $suites = @($suites | Where-Object { $Only -contains $_.Group -or $Only -contains $_.Name }) }
if ($suites.Count -eq 0) { throw "No suites matched -Only $($Only -join ',')" }

$results = @()
foreach ($s in $suites) {
    Write-Host "--- $($s.Name) " -NoNewline -ForegroundColor Cyan
    $out = & pwsh -NoProfile -File (Join-Path $PSScriptRoot $s.Path) 2>&1
    $code = $LASTEXITCODE

    $text = ($out | Out-String)
    $passed = 0; $failed = 0
    if ($text -match '(\d+) passed, (\d+) failed') { $passed = [int]$Matches[1]; $failed = [int]$Matches[2] }
    else { $failed = -1 }   # could not parse -- treat as a failure, never as a pass

    $ok = ($failed -eq 0 -and $code -eq 0)
    Write-Host $(if ($ok) { "OK ($passed)" } else { "FAILED" }) -ForegroundColor $(if ($ok) { 'Green' } else { 'Red' })
    if (-not $ok) { $out | Select-String -Pattern '^FAIL|Exception|ParserError' | ForEach-Object { Write-Host "    $_" -ForegroundColor Red } }

    $results += [pscustomobject]@{ Suite = $s.Name; Passed = $passed; Failed = $failed; Exit = $code; OK = $ok }
}

Write-Host ''
$results | Format-Table -AutoSize
$totalPass = ($results | Measure-Object -Property Passed -Sum).Sum
$bad = @($results | Where-Object { -not $_.OK })
Write-Host "TOTAL: $totalPass assertions passed across $($results.Count) suites; $($bad.Count) suite(s) failed." `
    -ForegroundColor $(if ($bad.Count -eq 0) { 'Green' } else { 'Red' })
if ($bad.Count -gt 0) { exit 1 }
