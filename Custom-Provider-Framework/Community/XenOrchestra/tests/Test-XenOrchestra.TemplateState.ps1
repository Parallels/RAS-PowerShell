[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$provider = Join-Path $root 'Parallels-RAS-CPF-XenOrchestra.ps1'
$tempConfig = Join-Path ([IO.Path]::GetTempPath()) "xo-template-config-$([guid]::NewGuid()).json"
$tempState = Join-Path ([IO.Path]::GetTempPath()) "xo-template-state-$([guid]::NewGuid()).json"
$tempLog = Join-Path ([IO.Path]::GetTempPath()) "xo-template-log-$([guid]::NewGuid()).log"

function Assert($Condition,[string]$Message) { if (-not $Condition) { throw $Message } }

try {
    $config = Get-Content -Raw (Join-Path $root 'XO-RAS-Config.example.json') | ConvertFrom-Json
    $config.capabilities.template_method.value = 'versioning'
    $config.logging.log_file.value = $tempLog
    [IO.File]::WriteAllText($tempConfig,($config | ConvertTo-Json -Depth 10),[Text.UTF8Encoding]::new($false))

    . $provider -SettingsPath $tempConfig -TemplateStatePath $tempState -NoRun
    $script:Session = [pscustomobject]@{ base='https://xo.test/rest/v0'; token='test-only'; skip_tls=$false }
    $id = '11111111-2222-3333-4444-555555555555'
    $vm = [pscustomobject]@{ id=$id; name_label='Windows11'; power_state='Halted'; type='VM' }
    function Invoke-Xo {
        param([string]$Method,[string]$Path,[object]$Body = $null)
        if ($Method -eq 'GET' -and $Path -eq "/vms/$id") { return $vm }
        throw "Unexpected mocked XO request: $Method $Path"
    }

    $converted = Convert-Guest ([pscustomobject]@{ id=$id; is_template=$true })
    Assert ($converted.result.task_id -eq "completed:convert:$id") 'convert(true) did not return a completed conversion task'
    Assert ((To-Guest $vm).is_template -eq $true) 'convert(true) did not mark the VM as a template'
    $task = Get-Task ([pscustomobject]@{ id=$converted.result.task_id })
    Assert ($task.result.state -eq 'completed') 'Conversion task was not completed'
    Assert (-not $task.result.output.PSObject.Properties['clone_id']) 'Conversion task incorrectly returned a clone id'

    $stored = Get-Content -Raw $tempState | ConvertFrom-Json
    Assert ($stored.schema_version -eq 1) 'Template state schema is wrong'
    Assert (@($stored.templates).Count -eq 1) 'Template state was not persisted'

    $script:TemplateIds = @{}; $script:MaintenanceIds = @{}
    Import-TemplateState
    Assert ((To-Guest $vm).is_template -eq $true) 'Template state did not survive a reload'

    $maintenance = Convert-Guest ([pscustomobject]@{ id=$id; is_template=$false })
    Assert ($maintenance.result.task_id -eq "completed:convert:$id") 'convert(false) did not return a completed conversion task'
    Assert ((To-Guest $vm).is_template -eq $false) 'convert(false) did not enter maintenance'
    $nativeVm = [pscustomobject]@{ id=$id; name_label='Native'; power_state='Halted'; type='VM-template' }
    Assert ((To-Guest $nativeVm).is_template -eq $false) 'Maintenance did not override a native template flag'

    $script:TemplateIds = @{}; $script:MaintenanceIds = @{}
    Import-TemplateState
    Assert ((To-Guest $vm).is_template -eq $false) 'Maintenance state did not survive a reload'

    $null = Convert-Guest ([pscustomobject]@{ id=$id; is_template=$true })
    Assert ((To-Guest $vm).is_template -eq $true) 'Leaving maintenance did not restore the template flag'
    $script:Session.base = 'https://other-xo.test/rest/v0'
    Assert ((To-Guest $vm).is_template -eq $false) 'Template state leaked across Xen Orchestra hosts'
    $script:Session.base = 'https://xo.test/rest/v0'

    $invalid = Convert-Guest ([pscustomobject]@{ id=$id; is_template='true' })
    Assert ($invalid.error.code -eq -32602) 'A non-Boolean is_template value was accepted'

    [IO.File]::WriteAllText($tempState,'{broken',[Text.UTF8Encoding]::new($false))
    $threw = $false
    try { Import-TemplateState } catch { $threw = $true }
    Assert $threw 'Corrupt template state was silently accepted'
    Assert ([IO.File]::ReadAllText($tempState) -eq '{broken') 'Corrupt template state was overwritten'
}
finally {
    foreach ($path in @($tempConfig,$tempState,$tempLog)) {
        if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
    }
}
Write-Host 'All Xen Orchestra template-state checks passed.' -ForegroundColor Green
