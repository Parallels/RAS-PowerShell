function Start-CommandProcess {

    param([string]$CommandPath, [string]$CommandArgs)

    if (-not (Test-Path $CommandPath)) {
        throw "Command not found: $CommandPath"
    }

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName               = $CommandPath
    $psi.Arguments              = $CommandArgs
    $psi.UseShellExecute        = $false
    $psi.RedirectStandardInput  = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.CreateNoWindow         = $true

    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi

    return $proc
}

function Stop-CommandProcess {

    param([System.Diagnostics.Process]$Process)

    if ($Process.HasExited) {
        return
    }

    $Process.Kill()
    $Process.WaitForExit()
}

function Open-IOStreams {

    param([System.Diagnostics.Process]$Process)

    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    $StandardInput = New-Object System.IO.StreamWriter($Process.StandardInput.BaseStream, $utf8NoBom)
    $StandardInput.AutoFlush = $true

    $StandardOutput = New-Object System.IO.StreamReader($Process.StandardOutput.BaseStream, [System.Text.Encoding]::UTF8)
    $StandardError = New-Object System.IO.StreamReader($Process.StandardError.BaseStream, [System.Text.Encoding]::UTF8)

    return @{ StandardInput = $StandardInput; StandardOutput = $StandardOutput; StandardError = $StandardError }
}

function Close-IOStreams {

    param([object]$IOStreams)

    $IOStreams.StandardInput.Close()
    $IOStreams.StandardInput.Dispose()

    while ($null -ne ($line = $IOStreams.StandardError.ReadLine())) {
        Write-Error $line
    }

    $IOStreams.StandardError.Close()
    $IOStreams.StandardError.Dispose()

    $IOStreams.StandardOutput.Close()
    $IOStreams.StandardOutput.Dispose()
}

function Write-QueryObject {

    param([System.IO.StreamWriter]$StandardInput, [object]$QueryObject)

    $jsonLine = $QueryObject | ConvertTo-Json -Compress -Depth 10

    Write-Host "VERBOSE: Query started '$jsonLine'" -Verbose -ForegroundColor Yellow

    $StandardInput.WriteLine($jsonLine)
}

function Read-ResultObject {

    param([System.IO.StreamReader]$StandardOuput)

    $jsonLine = $StandardOuput.ReadLine()

    if (-not $jsonLine) {
        throw "Cannot parse empty response"
    }

    Write-Host "VERBOSE: Query completed '$jsonLine'" -Verbose -ForegroundColor Yellow

    $responseObj = $jsonLine.Trim() | ConvertFrom-Json -ErrorAction Stop

    $error = $responseObj | Select-Object -ExpandProperty error -ErrorAction SilentlyContinue
    if ($error) {
        throw $error
    }

    $result = $responseObj | Select-Object -ExpandProperty result -ErrorAction SilentlyContinue
    if ($null -eq $result) {
        throw "Missing 'result' in '$jsonLine'"
    }

    return $result
}

function Select-TaskOutput {

    param([object]$taskObj)

    $error = $taskObj | Select-Object -ExpandProperty error -ErrorAction SilentlyContinue
    if ($error) {
        throw $error
    }

    $output = $taskObj | Select-Object -ExpandProperty output -ErrorAction SilentlyContinue
    if ($null -eq $output) {
        throw "Missing 'output' in '$taskObj'"
    }

    return $output
}

function Submit-Query {

    param([object]$IOStreams, [object]$QueryObject)

    Write-QueryObject $IOStreams.StandardInput $QueryObject
    return Read-ResultObject $IOStreams.StandardOutput
}

function Submit-Initialize {

    param([object]$IOStreams)

    $QueryObject = @{
        method = "provider/initialize"
    }

    $result = Submit-Query $IOStreams $QueryObject

    return $result
}

function Submit-Connect {

    param([object]$IOStreams, [object]$CustomSettings)

    $QueryObject = @{
        method = "provider/connect"
        params = @{ settings = $CustomSettings; }
    }

    $result = Submit-Query $IOStreams $QueryObject

    return $result
}

function Submit-InitializeAndConnect {

    param([object]$IOStreams, [object]$CustomSettings)

    Submit-Initialize $IOStreams
    Submit-Connect $IOStreams $CustomSettings | Out-Null
}

function Submit-Disconnect {

    param([object]$IOStreams)

    $QueryObject = @{
        method = "provider/disconnect"
    }

    $result = Submit-Query $IOStreams $QueryObject

    return $result
}

function Submit-GuestsList {

    param([object]$IOStreams)

    $QueryObject = @{
        method = "guests/list"
    }

    $result = Submit-Query $IOStreams $QueryObject

    return $result
}

function Submit-GuestsGet {

    param([object]$IOStreams, [string]$GuestID)

    $QueryObject = @{
        method = "guests/get"
        params = @{ id = $GuestID; }
    }

    $result = Submit-Query $IOStreams $QueryObject

    return $result
}

function Submit-GuestsControl {

    param([object]$IOStreams, [string]$GuestID, [string]$Control)

    $QueryObject = @{
        method = "guests/control"
        params = @{ id = $GuestID; control = $Control; }
    }

    $result = Submit-Query $IOStreams $QueryObject

    return $result
}

function Submit-TasksGet {

    param([object]$IOStreams, [string]$GuestID)

    $QueryObject = @{
        method = "tasks/get"
        params = @{ id = $GuestID; }
    }

    $result = Submit-Query $IOStreams $QueryObject

    return $result
}

function Submit-GuestsConvert {

    param([object]$IOStreams, [string]$GuestID, [boolean]$IsTemplate)

    $QueryObject = @{
        method = "guests/convert"
        params = @{ id = $GuestID; is_template = $IsTemplate; }
    }

    $result = Submit-Query $IOStreams $QueryObject

    return $result
}

function Submit-GuestsClone {

    param([object]$IOStreams, [string]$GuestID, [string]$CloneName, [string]$SnapshotName)

    $QueryObject = @{
        method = "guests/clone"
        params = @{ id = $GuestID; name = $CloneName; snapshot = $SnapshotName }
    }

    $result = Submit-Query $IOStreams $QueryObject

    return $result
}

function Submit-GuestsSnapshotsCreate {

    param([object]$IOStreams, [string]$GuestID, [string]$SnapshotName)

    $QueryObject = @{
        method = "guests/snapshots/create"
        params = @{ id = $GuestID; name = $SnapshotName }
    }

    $result = Submit-Query $IOStreams $QueryObject

    return $result
}

function Submit-GuestsSnapshotsDelete {

    param([object]$IOStreams, [string]$GuestID, [string]$SnapshotName)

    $QueryObject = @{
        method = "guests/snapshots/delete"
        params = @{ id = $GuestID; name = $SnapshotName }
    }

    $result = Submit-Query $IOStreams $QueryObject

    return $result
}

function Submit-GuestsSnapshotsExists {

    param([object]$IOStreams, [string]$GuestID, [string]$SnapshotName)

    $QueryObject = @{
        method = "guests/snapshots/exists"
        params = @{ id = $GuestID; name = $SnapshotName }
    }

    $result = Submit-Query $IOStreams $QueryObject

    return $result
}

function Submit-GuestsSnapshotsRevert {

    param([object]$IOStreams, [string]$GuestID, [string]$SnapshotName)

    $QueryObject = @{
        method = "guests/snapshots/revert"
        params = @{ id = $GuestID; name = $SnapshotName }
    }

    $result = Submit-Query $IOStreams $QueryObject

    return $result
}

function Invoke-ScriptBlock {

    param([string]$CommandPath, [string]$CommandArgs, [ScriptBlock]$ScriptBlock)

    try {
        $Process = Start-CommandProcess $CommandPath $CommandArgs

        try {
            $Process.Start() | Out-Null

            $IOStreams = Open-IOStreams $Process

            try {
                & $ScriptBlock $IOStreams
            }
            finally {
                Close-IOStreams $IOStreams
            }
        }
        finally {
            Stop-CommandProcess $Process
        }
    }
    catch {
        Write-Error "Unhandled error: $_"
    }
}

function Invoke-AsyncTask {

    param([object]$IOStreams, [int]$PollingRate, [ScriptBlock]$ScriptBlock)

    $taskId = (& $ScriptBlock).task_id

    while ("running" -eq ($taskObj = Submit-TasksGet $IOStreams $taskId).state) {
        Start-Sleep -Seconds $PollingRate
    }

    return Select-TaskOutput $taskObj
}

Export-ModuleMember -Function Invoke-ScriptBlock
Export-ModuleMember -Function Invoke-AsyncTask
Export-ModuleMember -Function Submit-InitializeAndConnect
Export-ModuleMember -Function Submit-Disconnect
Export-ModuleMember -Function Submit-GuestsList
Export-ModuleMember -Function Submit-GuestsGet
Export-ModuleMember -Function Submit-GuestsControl
Export-ModuleMember -Function Submit-TasksGet
Export-ModuleMember -Function Submit-GuestsConvert
Export-ModuleMember -Function Submit-GuestsClone
Export-ModuleMember -Function Submit-GuestsSnapshotsCreate
Export-ModuleMember -Function Submit-GuestsSnapshotsDelete
Export-ModuleMember -Function Submit-GuestsSnapshotsExists
Export-ModuleMember -Function Submit-GuestsSnapshotsRevert