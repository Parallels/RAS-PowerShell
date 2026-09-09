#Requires -Version 7.0
<#
    Runs every suite in this folder and prints one summary table.

    Regenerates the unit harness first, so the unit suites can never run against a
    stale projection of the provider. The E2E suite copies the provider and starts
    its own mock Proxmox server, so this is safe to run repeatedly with nothing set
    up by hand.

        pwsh -File Tests/run-all.ps1
        pwsh -File Tests/run-all.ps1 -Only unit
        pwsh -File Tests/run-all.ps1 -Only e2e-proxmox,subprocess

    Exit code is non-zero if any suite fails, so it is usable as a gate.
#>

param(
    [string[]]$Only
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$suites = @(
    @{ Group = 'unit';        Name = 't3';  Path = 'unit/t3.ps1'  }
    @{ Group = 'unit';        Name = 't4';  Path = 'unit/t4.ps1'  }
    @{ Group = 'unit';        Name = 't5';  Path = 'unit/t5.ps1'  }
    @{ Group = 'unit';        Name = 't6';  Path = 'unit/t6.ps1'  }
    @{ Group = 'unit';        Name = 't7';  Path = 'unit/t7.ps1'  }
    @{ Group = 'unit';        Name = 't8';  Path = 'unit/t8.ps1'  }
    @{ Group = 'unit';        Name = 't9';  Path = 'unit/t9.ps1'  }
    @{ Group = 'unit';        Name = 't11'; Path = 'unit/t11.ps1' }
    @{ Group = 'unit';        Name = 't12'; Path = 'unit/t12.ps1' }
    @{ Group = 'unit';        Name = 't13'; Path = 'unit/t13.ps1' }
    @{ Group = 'unit';        Name = 't14'; Path = 'unit/t14.ps1' }
    @{ Group = 'unit';        Name = 't15'; Path = 'unit/t15.ps1' }
    @{ Group = 'unit';        Name = 't16'; Path = 'unit/t16.ps1' }
    @{ Group = 'unit';        Name = 't17'; Path = 'unit/t17.ps1' }
    @{ Group = 'unit';        Name = 't18'; Path = 'unit/t18.ps1' }
    @{ Group = 'subprocess';  Name = 't10'; Path = 'subprocess/t10.ps1' }
    @{ Group = 'e2e-proxmox'; Name = 'e2e-proxmox'; Path = 'e2e-proxmox/run_e2e.ps1' }
)

if ($Only) { $suites = @($suites | Where-Object { $Only -contains $_.Group -or $Only -contains $_.Name }) }
if ($suites.Count -eq 0) { throw "No suites matched -Only $($Only -join ',')" }

if (@($suites | Where-Object { $_.Group -eq 'unit' }).Count -gt 0) {
    & (Join-Path $PSScriptRoot 'lib/make-harness.ps1')
    Write-Host ''
}

$results = @()
foreach ($s in $suites) {
    Write-Host "--- $($s.Name) " -NoNewline -ForegroundColor Cyan
    $out = & pwsh -NoProfile -File (Join-Path $PSScriptRoot $s.Path) 2>&1
    $code = $LASTEXITCODE

    # Suites report either "N passed, M failed" or t9's "PASSED: N  FAILED: M".
    $text = ($out | Out-String)
    $passed = 0; $failed = 0
    if ($text -match '(\d+) passed, (\d+) failed') { $passed = [int]$Matches[1]; $failed = [int]$Matches[2] }
    elseif ($text -match 'PASSED: (\d+)\s+FAILED: (\d+)') { $passed = [int]$Matches[1]; $failed = [int]$Matches[2] }
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
