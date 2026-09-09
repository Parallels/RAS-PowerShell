$ErrorActionPreference = 'Stop'
$pass = 0
$fail = 0
function Assert {
    param([bool]$Cond, [string]$Msg)
    if ($Cond) { $script:pass++; Write-Host "PASS: $Msg" -ForegroundColor Green }
    else { $script:fail++; Write-Host "FAIL: $Msg" -ForegroundColor Red }
}

# ============================================================
# TESTS for settings schema migration (SETTINGS.md's "Migration notes"):
#   an older, successfully-parsed file gets rewritten in place to the
#   current schema -- every existing value preserved exactly, new keys
#   appear at their default, schema_version is bumped, and a specific
#   (not just "something changed") log line records what was added.
#
# Deliberately uses its OWN isolated settings file, never the shared
# Tests/unit/RAS-CPF-Proxmox-Settings.json fixture every other suite reads
# via funcs2.ps1's own bootstrap -- a migration mutates the file on disk,
# and that fixture must stay untouched for every other suite to keep
# starting from the same known state.
# ============================================================
. "$PSScriptRoot/funcs2.ps1"

$isolatedDir = Join-Path $PSScriptRoot 't18-work'
Remove-Item $isolatedDir -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Path $isolatedDir -Force | Out-Null
$script:SettingsPath = Join-Path $isolatedDir 'RAS-CPF-Proxmox-Settings.json'
$script:LogPath = Join-Path $isolatedDir 't18.log'

# ------------------------------------------------------------
# A. A schema-2 file (bare-scalar shape, matching a real hand-simplified
#    file) with some values already customized -- loading it migrates it
#    to schema 3 in place.
# ------------------------------------------------------------
Write-Host "`n--- Test A: schema 2 -> 3 migration ---" -ForegroundColor Cyan
@'
{
  "schema_version": 2,
  "capabilities": { "can_link_clones": true, "guests_polling_rate": 15 },
  "cloning": { "load_balancing": { "preserve_node_on_recreation": false } }
}
'@ | Set-Content -Path $script:SettingsPath -Encoding UTF8

Remove-Item $script:LogPath -ErrorAction SilentlyContinue
Import-ProviderSettings
$logContent = Get-Content $script:LogPath -Raw

Assert ($logContent -match 'Settings schema migrated 2 -> 3') "A1: the migration log line fires, naming the exact old and new versions"
Assert ($logContent -match 'http_timeout_seconds') "A2: the log names the specific new key(s) added, not just 'something changed'"
Assert ($logContent -match 'mac_preservation') "A3: ...covers every new key from this version, not just the first one"
Assert ($script:Settings.schema_version -eq 3) "A4: `$script:Settings.schema_version reflects the NEW version immediately, in the same load that migrated it"

$onDisk = Get-Content $script:SettingsPath -Raw | ConvertFrom-Json
Assert ($onDisk.schema_version -eq 3) "A5: the file on disk now says schema_version 3"
Assert ($onDisk.capabilities.can_link_clones.value -eq $true) "A6: an existing customized value (can_link_clones=true) survives the rewrite exactly"
Assert ($onDisk.capabilities.guests_polling_rate.value -eq 15) "A7: ...same for a second customized value (guests_polling_rate=15)"
Assert ($onDisk.cloning.load_balancing.preserve_node_on_recreation.value -eq $false) "A8: ...and a customized value that happens to match a non-default boolean"
Assert ($null -ne $onDisk.cloning.timeouts.http_timeout_seconds) "A9: the brand-new http_timeout_seconds key now exists in the file"
Assert ($onDisk.cloning.timeouts.http_timeout_seconds.value -eq 3) "A10: ...at its default value (3), since the old file never set it"
Assert ($onDisk.cloning.mac_preservation.enabled.value -eq $false) "A11: the new mac_preservation.enabled key exists, at its off-by-default value"
Assert ($onDisk.virtual_machines.pool_scope.pool_name.value -eq '') "A12: the new pool_scope.pool_name key exists, at its empty (no filtering) default"
Assert ($onDisk.virtual_machines.pool_scope.inherit_on_clone.value -eq $true) "A13: ...and inherit_on_clone at its on-by-default value"
Assert (-not [string]::IsNullOrWhiteSpace([string]$onDisk.cloning.timeouts.http_timeout_seconds.description)) "A14: the new key's rich shape carries a real description, not just a bare value"

# ------------------------------------------------------------
# B. Loading the NOW-current (schema 3) file again does not re-migrate --
#    idempotent, no duplicate log line, file untouched.
# ------------------------------------------------------------
Write-Host "`n--- Test B: already-current schema is never re-migrated ---" -ForegroundColor Cyan
Remove-Item $script:LogPath -ErrorAction SilentlyContinue
$mtimeBefore = (Get-Item $script:SettingsPath).LastWriteTimeUtc
Start-Sleep -Milliseconds 50
Import-ProviderSettings
$mtimeAfter = (Get-Item $script:SettingsPath).LastWriteTimeUtc
$logContentB = if (Test-Path $script:LogPath) { Get-Content $script:LogPath -Raw } else { '' }

Assert ($logContentB -notmatch 'schema migrated') "B1: no migration log fires for a file already at the current schema"
Assert ($mtimeBefore -eq $mtimeAfter) "B2: the file itself is not rewritten at all -- untouched, not just silently re-written identically"

# ------------------------------------------------------------
# C. A file that fails to parse is left completely untouched, same rule as
#    every other settings-file failure mode -- migration must never override
#    that.
# ------------------------------------------------------------
Write-Host "`n--- Test C: a corrupt file is never migrated (or touched at all) ---" -ForegroundColor Cyan
'{ this is not valid json' | Set-Content -Path $script:SettingsPath -Encoding UTF8
$corruptContentBefore = Get-Content $script:SettingsPath -Raw
Remove-Item $script:LogPath -ErrorAction SilentlyContinue
Import-ProviderSettings
$corruptContentAfter = Get-Content $script:SettingsPath -Raw
$logContentC = if (Test-Path $script:LogPath) { Get-Content $script:LogPath -Raw } else { '' }

Assert ($corruptContentBefore -eq $corruptContentAfter) "C1: a corrupt file's content is byte-for-byte untouched -- migration never fires for it"
Assert ($logContentC -notmatch 'schema migrated') "C2: no migration log fires for a corrupt file either"
Assert ($script:Settings.schema_version -eq 3) "C3: the provider still runs on in-memory current-schema defaults despite the corrupt file"

# ------------------------------------------------------------
# D. A genuine schema-1 file (no schema_version field, the original flat
#    shape) is DETECTED but deliberately NEVER auto-rewritten --
#    only capabilities.can_link_clones has a real fallback-read from its
#    old location; every other schema-1 key would auto-migrate to a
#    silently-defaulted value, discarding whatever an admin actually had
#    customized under the old flat shape. Schema 2 onward is
#    purely additive, so THAT case is safe to auto-rewrite (Tests A-C
#    above) -- schema 1 (its predecessor, which RELOCATED
#    keys) deliberately is not.
# ------------------------------------------------------------
Write-Host "`n--- Test D: a genuine schema-1 file is detected but never auto-rewritten ---" -ForegroundColor Cyan
'{ "cloning": { "linked_clones_enabled": true } }' | Set-Content -Path $script:SettingsPath -Encoding UTF8
$schema1ContentBefore = Get-Content $script:SettingsPath -Raw
Remove-Item $script:LogPath -ErrorAction SilentlyContinue
Import-ProviderSettings
$schema1ContentAfter = Get-Content $script:SettingsPath -Raw
$logContentD = if (Test-Path $script:LogPath) { Get-Content $script:LogPath -Raw } else { '' }

Assert ($script:Settings.schema_version -eq 1) "D1: a schema-1 file is correctly detected as schema 1 in memory (not silently bumped)"
Assert ($schema1ContentBefore -eq $schema1ContentAfter) "D2: ...but the file on disk is left completely untouched -- no auto-migration for schema 1"
Assert ($logContentD -notmatch 'schema migrated') "D3: no 'schema migrated' log fires for it"
Assert ($script:Settings.capabilities.can_link_clones -eq $true) "D4: the one deliberate exception (can_link_clones's fallback-read from cloning.linked_clones_enabled) still resolves correctly in memory even without auto-migration"

Remove-Item $isolatedDir -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "`n=== t18 Summary: $pass passed, $fail failed ===" -ForegroundColor $(if ($fail -eq 0) { 'Green' } else { 'Red' })
if ($fail -gt 0) { exit 1 }
