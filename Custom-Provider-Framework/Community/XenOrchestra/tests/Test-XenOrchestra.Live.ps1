[CmdletBinding()]
param(
    [string]$HostUrl = 'https://192.168.100.72',
    [string]$Token = $(if ($env:XOA_TOKEN) { $env:XOA_TOKEN } else { [Environment]::GetEnvironmentVariable('XOA_TOKEN','User') }),
    [switch]$SkipCertificateCheck,
    [switch]$AllowClone,
    [switch]$LinkedClone,
    [switch]$VersionedClone,
    [switch]$StopSourceForBoot,
    [string]$TemplateId,
    [string]$XcpHost,
    [string]$CloneName = "cpf-live-test-$((Get-Date).ToString('yyyyMMdd-HHmmss'))"
)

$ErrorActionPreference = 'Stop'
if ($VersionedClone) { $LinkedClone = $true }
$base = $HostUrl.TrimEnd('/') + '/rest/v0'

function Invoke-Xoa([string]$Method,[string]$Path,[object]$Body) {
    $args = @{ Method=$Method; Uri="$base$Path"; TimeoutSec=30; SkipHttpErrorCheck=$true }
    if ($SkipCertificateCheck) { $args.SkipCertificateCheck = $true }
    if ($Token) { $args.Headers = @{ Cookie="authenticationToken=$Token" } }
    if ($null -ne $Body) { $args.ContentType='application/json'; $args.Body=$Body | ConvertTo-Json -Depth 10 -Compress }
    $response = Invoke-WebRequest @args
    if ($response.StatusCode -lt 200 -or $response.StatusCode -ge 300) { throw "$Method $Path returned HTTP $($response.StatusCode)" }
    if ([string]::IsNullOrWhiteSpace($response.Content)) { return $null }
    $response.Content | ConvertFrom-Json -AsHashtable
}

$swagger = Invoke-Xoa GET '/docs/swagger.json' $null
if ($swagger.openapi -ne '3.0.0') { throw "Unexpected OpenAPI version [$($swagger.openapi)]" }
$cloneVm = $swagger.paths.'/vms/{id}/actions/clone'.post
if ($null -eq $cloneVm -or $cloneVm.operationId -ne 'CloneVm') { throw 'Live XOA has no documented CloneVm operation' }
$revertVm = $swagger.paths.'/vms/{id}/actions/revert_snapshot'.post
if ($null -eq $revertVm -or $revertVm.operationId -ne 'RevertSnapshotVm') { throw 'Live XOA has no documented RevertSnapshotVm operation' }
$resumeVm = $swagger.paths.'/vms/{id}/actions/resume'.post
if ($null -eq $resumeVm -or $resumeVm.operationId -ne 'ResumeVm') { throw 'Live XOA has no documented ResumeVm operation' }
$taskStatuses = @($swagger.components.schemas.XoTask.properties.status.enum)
if (($taskStatuses | Sort-Object) -join ',' -ne 'failure,interrupted,pending,success') {
    throw "Unexpected XoTask statuses [$($taskStatuses -join ',')]"
}
Write-Host "PASS public Swagger: $($swagger.info.title) OpenAPI $($swagger.openapi)" -ForegroundColor Green

if ([string]::IsNullOrWhiteSpace($Token)) {
    throw 'Set XOA_TOKEN for authenticated read-only inventory tests; no token value is logged.'
}

$vms = @(Invoke-Xoa GET '/vms?fields=id,name_label,power_state,addresses,os_version,VIFs,type&limit=1' $null)
$templates = @(Invoke-Xoa GET '/vm-templates?fields=id,uuid,name_label,power_state,addresses,os_version,VIFs,type,$pool,$poolId&limit=1' $null)
$pools = @(Invoke-Xoa GET '/pools?fields=id,name_label&limit=1' $null)
foreach ($entry in @($vms + $templates + $pools)) {
    if ($entry -is [string] -or -not $entry.id) { throw 'Authenticated collection returned an object URL or an object without id' }
}
Write-Host "PASS authenticated reads: vms=$($vms.Count) templates=$($templates.Count) pools=$($pools.Count)" -ForegroundColor Green

$provider = Join-Path (Split-Path $PSScriptRoot -Parent) 'Parallels-RAS-CPF-XenOrchestra.ps1'
$script:ProviderSettingsPath = $null
function Invoke-Cpf([string]$Method,[hashtable]$Params) {
    $inputLines = @(
        @{ method='provider/connect'; params=@{ settings=@{ host=$HostUrl; token=$Token; skip_tls=[bool]$SkipCertificateCheck } } }
        @{ method=$Method; params=$Params }
        @{ method='provider/disconnect' }
    ) | ForEach-Object { $_ | ConvertTo-Json -Depth 10 -Compress }
    $providerArgs = @('-NoProfile','-NonInteractive','-File',$provider)
    if ($script:ProviderSettingsPath) { $providerArgs += @('-SettingsPath',$script:ProviderSettingsPath) }
    $raw = @($inputLines | & pwsh @providerArgs)
    if ($raw.Count -ne 3) { throw "$Method emitted $($raw.Count) responses instead of 3" }
    if (($raw -join '') -like "*$Token*") { throw "$Method emitted the authentication token" }
    $response = $raw[1] | ConvertFrom-Json
    if ($response.error) { throw "$Method failed: $($response.error.message)" }
    $response.result
}
function Wait-CpfTask([string]$TaskId) {
    for ($i=1; $i -le 180; $i++) {
        $task = Invoke-Cpf 'tasks/get' @{ id=$TaskId }
        if ($task.state -eq 'completed') { return $task }
        if ($task.state -eq 'failed') { throw "Task $TaskId failed: $($task.error)" }
        Start-Sleep -Seconds 5
    }
    throw "Task $TaskId did not complete after 180 polls"
}
$sampleId = if ($vms.Count) { [string]$vms[0].id } elseif ($templates.Count) { [string]$templates[0].id } else { throw 'XOA returned no VM or template for guests/get validation' }
$requests = @(
    @{ method='provider/initialize' }
    @{ method='provider/connect'; params=@{ settings=@{ host=$HostUrl; token=$Token; skip_tls=[bool]$SkipCertificateCheck } } }
    @{ method='guests/list' }
    @{ method='guests/get'; params=@{ id=$sampleId } }
    @{ method='provider/disconnect' }
) | ForEach-Object { $_ | ConvertTo-Json -Depth 10 -Compress }
$output = @($requests | & pwsh -NoProfile -NonInteractive -File $provider)
if ($output.Count -ne 5) { throw "CPF emitted $($output.Count) responses for 5 requests" }
$cpf = @($output | ForEach-Object { $_ | ConvertFrom-Json })
if (-not $cpf[1].result.connected) { throw "CPF connect failed: $($cpf[1].error.message)" }
if (@($cpf[2].result.guests).Count -lt 1) { throw 'CPF guests/list returned no ids' }
if ($cpf[3].result.id -ne $sampleId) { throw 'CPF guests/get returned the wrong id' }
if (-not $cpf[4].result.disconnected) { throw 'CPF disconnect failed' }
if (($output -join '') -like "*$Token*") { throw 'CPF emitted the authentication token' }
Write-Host "PASS CPF process: guests=$(@($cpf[2].result.guests).Count), get=$sampleId, token not emitted" -ForegroundColor Green

if (-not $AllowClone) {
    Write-Host 'SKIP clone: pass -AllowClone and -TemplateId for a disposable live clone.' -ForegroundColor Yellow
    exit 0
}
if ([string]::IsNullOrWhiteSpace($TemplateId)) { throw '-TemplateId is required with -AllowClone' }
if ($LinkedClone) {
    if ($TemplateId -notmatch '^[0-9a-fA-F-]{36}$') { throw '-TemplateId must be a UUID for the linked-clone storage check' }
    if ($XcpHost -notmatch '^[A-Za-z0-9.-]+$') { throw '-XcpHost is required with -LinkedClone and must be an IP address or hostname' }
}

$cloneId = $null
$snapshotName = $null
$sourceWasRunning = $false
try {
    if ($VersionedClone) {
        $config = Get-Content -Raw (Join-Path (Split-Path $provider -Parent) 'XO-RAS-Config.example.json') | ConvertFrom-Json
        $config.capabilities.template_method.value = 'versioning'
        $script:ProviderSettingsPath = Join-Path ([IO.Path]::GetTempPath()) "xo-cpf-versioning-$([guid]::NewGuid()).json"
        [IO.File]::WriteAllText($script:ProviderSettingsPath,($config | ConvertTo-Json -Depth 10),[Text.UTF8Encoding]::new($false))
        $capabilities = Invoke-Cpf 'provider/initialize' @{}
        if ($capabilities.capabilities.template_method -ne 'versioning' -or -not $capabilities.capabilities.can_link_clones) { throw 'Provider did not advertise versioned linked clones' }
        $snapshotName = "RAS_TEMPLATE_VERSION_LIVE_$((Get-Date).ToString('yyyyMMddHHmmss'))"
        $snapshot = Invoke-Cpf 'guests/snapshots/create' @{ id=$TemplateId; name=$snapshotName }
        $null = Wait-CpfTask ([string]$snapshot.task_id)
    }
    $cloneParams = @{ id=$TemplateId; name=$CloneName; is_link_clone=[bool]$LinkedClone }
    if ($VersionedClone) { $cloneParams.snapshot = $snapshotName }
    $submitted = Invoke-Cpf 'guests/clone' $cloneParams
    $completed = Wait-CpfTask ([string]$submitted.task_id)
    $cloneId = [string]$completed.output.clone_id
    if ([string]::IsNullOrWhiteSpace($cloneId)) { throw 'Clone task returned no clone_id' }
    $clone = Invoke-Cpf 'guests/get' @{ id=$cloneId }
    if ($clone.name -ne $CloneName -or $clone.state -ne 'powered_off') { throw 'Clone identity or initial power state is wrong' }
    Write-Host "PASS CPF clone created: name=$CloneName id=$cloneId linked=$([bool]$LinkedClone)" -ForegroundColor Green

    if ($LinkedClone) {
        function Get-VhdParents([string]$VmId) {
            $remote = 'for vdi in $(xe vbd-list vm-uuid={0} type=Disk params=vdi-uuid --minimal | tr "," " "); do xe vdi-param-get uuid=$vdi param-name=sm-config; done' -f $VmId
            $lines = @(& ssh -o BatchMode=yes -o StrictHostKeyChecking=yes "root@$XcpHost" $remote)
            if ($LASTEXITCODE -ne 0) { throw "SSH VDI query failed for VM $VmId" }
            @($lines | ForEach-Object { if ($_ -match 'vhd-parent:\s*([0-9a-fA-F-]{36})') { $Matches[1] } })
        }
        $lineageSourceId = $TemplateId
        if ($VersionedClone) {
            $snapshots = @(Invoke-Xoa GET '/vm-snapshots?fields=id,name_label,$snapshot_of' $null)
            $versionSnapshot = $snapshots | Where-Object { $_.name_label -eq $snapshotName -and $_.'$snapshot_of' -eq $TemplateId } | Select-Object -First 1
            if (-not $versionSnapshot) { throw 'Version snapshot was not found after clone' }
            $lineageSourceId = [string]$versionSnapshot.id
        }
        $sourceParents = @(Get-VhdParents $lineageSourceId)
        $cloneParents = @(Get-VhdParents $cloneId)
        $sharedParents = @($sourceParents | Where-Object { $_ -in $cloneParents } | Select-Object -Unique)
        if (-not $sharedParents.Count) { throw 'Source and clone do not share a VHD parent; fast clone was not proven' }
        Write-Host "PASS linked VHD parent: $($sharedParents -join ',')" -ForegroundColor Green

        if ($StopSourceForBoot) {
            $source = Invoke-Cpf 'guests/get' @{ id=$TemplateId }
            if ($source.state -eq 'powered_on') {
                $sourceWasRunning = $true
                $sourceStop = Invoke-Cpf 'guests/control' @{ id=$TemplateId; control='stop' }
                $null = Wait-CpfTask ([string]$sourceStop.task_id)
            }
        }
        $started = Invoke-Cpf 'guests/control' @{ id=$cloneId; control='start' }
        $null = Wait-CpfTask ([string]$started.task_id)
        for ($i=1; $i -le 24; $i++) {
            $clone = Invoke-Cpf 'guests/get' @{ id=$cloneId }
            if ($clone.state -eq 'powered_on' -and @($clone.ip_addresses).Count) { break }
            Start-Sleep -Seconds 5
        }
        if ($clone.state -ne 'powered_on') { throw 'Linked clone did not boot' }
        Write-Host "PASS linked clone boot: IP=$(@($clone.ip_addresses) -join ',')" -ForegroundColor Green
        $stopped = Invoke-Cpf 'guests/control' @{ id=$cloneId; control='stop' }
        $null = Wait-CpfTask ([string]$stopped.task_id)
    }
}
finally {
    if ($LinkedClone -and -not [string]::IsNullOrWhiteSpace($cloneId)) {
        $null = Invoke-Cpf 'guests/control' @{ id=$cloneId; control='delete' }
        $remaining = @(Invoke-Xoa GET '/vms?fields=id' $null)
        if ($remaining.id -contains $cloneId) { throw "Linked clone still exists after delete: $cloneId" }
        Write-Host "PASS linked clone deleted: id=$cloneId" -ForegroundColor Green
    }
    if ($VersionedClone -and $snapshotName) {
        if (Invoke-Cpf 'guests/snapshots/exists' @{ id=$TemplateId; name=$snapshotName }) {
            $null = Invoke-Cpf 'guests/snapshots/delete' @{ id=$TemplateId; name=$snapshotName }
            if (Invoke-Cpf 'guests/snapshots/exists' @{ id=$TemplateId; name=$snapshotName }) { throw "Version snapshot still exists after delete: $snapshotName" }
            Write-Host "PASS version snapshot deleted: name=$snapshotName" -ForegroundColor Green
        }
    }
    if ($sourceWasRunning) {
        $source = Invoke-Cpf 'guests/get' @{ id=$TemplateId }
        if ($source.state -ne 'powered_on') {
            $sourceStart = Invoke-Cpf 'guests/control' @{ id=$TemplateId; control='start' }
            $null = Wait-CpfTask ([string]$sourceStart.task_id)
        }
        Write-Host "PASS source VM restored: id=$TemplateId" -ForegroundColor Green
    }
    if ($script:ProviderSettingsPath -and (Test-Path -LiteralPath $script:ProviderSettingsPath)) { Remove-Item -LiteralPath $script:ProviderSettingsPath -Force }
}
if (-not $LinkedClone) { Write-Warning 'The full test clone was intentionally not deleted. Verify it, then remove it explicitly in XOA.' }
