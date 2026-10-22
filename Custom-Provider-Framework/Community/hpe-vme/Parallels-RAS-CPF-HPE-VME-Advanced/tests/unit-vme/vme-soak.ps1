# vme-soak: long-run hygiene. Repeated full sweeps over a churning 3000-guest fleet at
# log_level 5 with a 1 MB rotation threshold. Checks that nothing grows with the
# number of instances that EVER existed (only with the current fleet), that rotation
# keeps the configured generations, and that memory stays flat.
param([int]$Count = 3000, [int]$Rounds = 6, [int]$ChurnPerRound = 100)
. (Join-Path $PSScriptRoot 'lib.ps1')

Start-VmeMock -Count $Count
$dir = New-ProviderDir -Name 'soak' -Settings @{
    inventory        = @{ list_min_refresh_seconds = 0; missing_listings_before_drop = 1 }
    orphan_detection = @{ check_interval_seconds = 30 }
    logging          = @{ log_level = 5; log_rotate_max_mb = 1; log_rotate_max_generations = 3; stats_interval_seconds = 0 }
}
$p = Start-Provider -Dir $dir
try {
    Connect-Provider $p | Out-Null
    $nextId = 20000
    $ws = @()
    for ($round = 1; $round -le $Rounds; $round++) {
        # Churn: delete the oldest fillers, create new instances (outside RAS).
        $ids = Get-GuestList $p
        $victims = @($ids | Where-Object { [long]$_ -ge 100 -and [long]$_ -lt 19000 -or [long]$_ -ge 20000 } | Select-Object -First $ChurnPerRound)
        foreach ($v in $victims) { Invoke-Mock "/_ctl/instance?id=$v&delete=1" | Out-Null }
        foreach ($i in 1..$ChurnPerRound) { Invoke-Mock "/_ctl/instance?id=$nextId&name=churn-$nextId&status=running&ip=10.8.0.1" | Out-Null; $nextId++ }
        $ids = Get-GuestList $p
        foreach ($id in $ids) { [void](Send-RpcRaw $p "{`"method`":`"guests/get`",`"params`":{`"id`":`"$id`"}}") }
        # Force the audit/GC timer and a state line.
        $proc = Get-Process -Id $p.Process.Id -ErrorAction SilentlyContinue
        $children = Get-CimInstance Win32_Process -Filter "ParentProcessId=$($p.Process.Id)" -ErrorAction SilentlyContinue
        $mem = [Math]::Round((@($proc) + @($children | ForEach-Object { Get-Process -Id $_.ProcessId -ErrorAction SilentlyContinue }) | Measure-Object WorkingSet64 -Sum).Sum / 1MB)
        $ws += $mem
        Write-Host "    round ${round}: $($ids.Count) guests listed, working set ${mem} MB" -ForegroundColor DarkGray
    }
    $list = Get-GuestList $p
    Assert ($list.Count -eq $Count + 6) "the listed fleet stays at its real size ($($list.Count)) through $($Rounds * $ChurnPerRound) creates and deletes"

    Start-Sleep -Seconds 31
    [void](Get-GuestList $p)   # runs the audit + GC
    Stop-Provider $p

    $gens = @(Get-ChildItem $dir -Filter 'VME-RAS-Provider.log*')
    Assert ($gens.Count -le 4 -and $gens.Count -ge 2) "log rotation happened and keeps at most 3 generations + the live file ($($gens.Count) files)"
    Assert (@($gens | Where-Object { $_.Length -gt 1.5MB }).Count -eq 0) 'no log file grew far past the 1 MB threshold'

    $all = (@($gens | Sort-Object LastWriteTime | ForEach-Object { [System.IO.File]::ReadAllText($_.FullName) }) -join "`n")
    $state = [regex]::Matches($all, 'STATS \(shutdown\) state: inventory=(\d+), hardware=(\d+)')
    Assert ($state.Count -ge 1) 'the shutdown statistics were written'
    if ($state.Count -ge 1) {
        $inv = [int]$state[$state.Count - 1].Groups[1].Value; $hw = [int]$state[$state.Count - 1].Groups[2].Value
        Assert ($inv -eq $Count + 6 -and $hw -le $inv) "inventory ($inv) and hardware cache ($hw) are bounded by the current fleet, not by every instance ever seen"
    }
    Assert ($all -notmatch '\[E 0[0-9A-F]/') 'no E line in the whole soak'
    $growth = $ws[-1] - $ws[1]
    Assert ($growth -lt 60) "working set flat after warm-up: $($ws -join ' -> ') MB (growth rounds 2..${Rounds}: $growth MB)"
}
catch { Assert $false "suite aborted: $($_.Exception.Message) at $($_.InvocationInfo.PositionMessage)" }
finally { Complete-Suite 'vme-soak' }
