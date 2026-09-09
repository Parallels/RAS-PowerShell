$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# ============================================================
# Regression tests for the StrictMode empty-object member read.
#
# Set-StrictMode -Version Latest makes a bare .PSObject.Properties.Name THROW
# when the object has no properties at all, and an empty JSON object ({})
# parses to exactly that. Two real inputs produced it:
#
#   A. A request such as {"method":"guests/get","params":{}} or a bare {}.
#      The main loop's try/catch caught it, so the provider survived, but RAS
#      got -32603 "The property 'Name' cannot be found on this object" instead
#      of the correct -32602/-32601.
#
#   B. A settings file emptied to {}. FATAL: Import-ProviderSettings runs at
#      startup, outside the main loop's try/catch, so the provider died before
#      it could serve or log anything -- defeating the documented contract that
#      a bad config must never prevent the provider from starting. Note {}
#      parses successfully, so the loader's own parse try/catch never sees it.
#
# These run the real script as a subprocess because B cannot be reproduced by
# dot-sourcing: the failure is in startup itself.
# ============================================================

$pass = 0
$fail = 0
function Assert {
    param([bool]$Cond, [string]$Msg)
    if ($Cond) { $script:pass++; Write-Host "PASS: $Msg" -ForegroundColor Green }
    else { $script:fail++; Write-Host "FAIL: $Msg" -ForegroundColor Red }
}

$providerSrc = (Resolve-Path (Join-Path $PSScriptRoot '../../Parallels-RAS-CPF-Proxmox-Advanced.ps1')).Path
$workDir = Join-Path $PSScriptRoot 't10-work'
if (Test-Path $workDir) { Remove-Item $workDir -Recurse -Force }
New-Item -ItemType Directory -Path $workDir -Force | Out-Null
Copy-Item $providerSrc (Join-Path $workDir 'provider.ps1')

# Runs the provider once over stdio and returns its response lines.
function Invoke-Provider {
    param([string[]]$Requests)

    # PowerShell has no '<' redirection, so pipe the requests in instead.
    $out = $Requests | & pwsh -NoProfile -File (Join-Path $workDir 'provider.ps1') 2>&1
    # The first line carries a UTF-8 BOM, which decodes to the single char
    # U+FEFF -- not three chars -- so it must be stripped as such.
    return @($out | ForEach-Object { ([string]$_).TrimStart([char]0xFEFF) } | Where-Object { $_.TrimStart().StartsWith('{') })
}

function Get-Code {
    param([string]$Line)
    try { return [int](($Line | ConvertFrom-Json).error.code) } catch { return 0 }
}

# ------------------------------------------------------------
# A. Malformed-but-parseable requests get the RIGHT error code, and the
#    provider keeps serving afterwards.
# ------------------------------------------------------------
Remove-Item (Join-Path $workDir 'RAS-CPF-Proxmox-Settings.json') -ErrorAction SilentlyContinue
$resp = @(Invoke-Provider @(
    '{"method":"guests/get","params":{}}',
    '{}',
    '{"method":"guests/control","params":{}}',
    '{"method":"provider/initialize"}'
))

Assert ($resp.Count -eq 4) "all four requests produced a response (got $($resp.Count))"
Assert ((Get-Code $resp[0]) -eq -32602) 'empty params on guests/get returns InvalidParams, not an opaque InternalError'
Assert ($resp[0] -match 'params\.id') '...and names the field that is actually missing'
Assert ((Get-Code $resp[1]) -eq -32601) 'a bare {} request returns MethodNotFound'
Assert ((Get-Code $resp[2]) -eq -32602) 'empty params on guests/control returns InvalidParams'
Assert ($resp[3] -match '"capabilities"') 'the provider is still serving normally after all three'
# -match against an ARRAY returns the matching elements, not a boolean, so count.
Assert (@($resp | Where-Object { $_ -match "property 'Name' cannot be found" }).Count -eq 0) 'no StrictMode member-read error leaks to RAS'

# ------------------------------------------------------------
# B. A settings file emptied to {} must not stop the provider starting.
#    This was fatal: no response, no log, just a dead process.
# ------------------------------------------------------------
Set-Content -LiteralPath (Join-Path $workDir 'RAS-CPF-Proxmox-Settings.json') -Value '{}' -Encoding UTF8
$resp = @(Invoke-Provider @('{"method":"provider/initialize"}'))

Assert ($resp.Count -ge 1) 'the provider starts and responds with a settings file of {}'
Assert ($resp[0] -match '"capabilities"') '...serving initialize normally'
Assert ($resp[0] -match '"tasks_polling_rate":11') '...on the hard-coded defaults, since {} supplies no values'

# B2: a section present but empty is the same trap one level down.
Set-Content -LiteralPath (Join-Path $workDir 'RAS-CPF-Proxmox-Settings.json') -Value '{"locations":{},"capabilities":{},"cloning":{"tags":{}},"logging":{}}' -Encoding UTF8
$resp = @(Invoke-Provider @('{"method":"provider/initialize"}'))
Assert ($resp.Count -ge 1 -and $resp[0] -match '"capabilities"') 'empty sections (including the nested cloning.tags) do not stop startup either'

# B3: a genuinely corrupt file must still be tolerated, as before.
Set-Content -LiteralPath (Join-Path $workDir 'RAS-CPF-Proxmox-Settings.json') -Value '{ this is not json' -Encoding UTF8
$resp = @(Invoke-Provider @('{"method":"provider/initialize"}'))
Assert ($resp.Count -ge 1 -and $resp[0] -match '"capabilities"') 'a corrupt settings file is still tolerated (unchanged behaviour)'
$after = Get-Content -LiteralPath (Join-Path $workDir 'RAS-CPF-Proxmox-Settings.json') -Raw
Assert ($after.Trim() -eq '{ this is not json') '...and is left untouched rather than overwritten'

Write-Host ""
Write-Host "=== t10 Summary: $pass passed, $fail failed ===" -ForegroundColor $(if ($fail -eq 0) { 'Green' } else { 'Red' })
if ($fail -gt 0) { exit 1 }
