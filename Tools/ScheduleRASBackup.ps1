<#  
.SYNOPSIS  
    Schedule RAS settings backup
.DESCRIPTION  
    Creates a Windows Scheduled Task to backup RAS settings (through RAS PowerShell).
.NOTES  
    File Name  : ScheduleRASBackup.ps1
    Author     : www.parallels.com
.EXAMPLE
    #Registers a new scheduled task to backup Parallels RAS settings
    .\ScheduleRASBackup.ps1

    #Un-registers the scheduled task that backs up the Parallels RAS settings
    .\ScheduleRASBackup.ps1 -RemoveSchedule
#>


[CmdletBinding(DefaultParametersetName="AddSchedule")]
param(  
    [Parameter(ParameterSetName='AddSchedule')]
	[string]$RasAdminUsername,

    [Parameter(ParameterSetName='AddSchedule')]
    [securestring]$RasAdminPassword,

    [Parameter(ParameterSetName='AddSchedule')]
    [string]$RASSettingsOutDir,
    
    [Parameter(ParameterSetName='AddSchedule')] 
    [bool]$Weekly = $false,

    [Parameter(ParameterSetName='ExportSettings')] 
    [switch]$ExportSettings, 
     
    [Parameter(ParameterSetName='RemoveSchedule')]
    [switch]$RemoveSchedule
)

Function Main
{
    if($ExportSettings) 
    {
        Start-ExportSettings
    } 
    elseif($RemoveSchedule) 
    {
        Unregister-RASBackupTask
    } 
    else 
    {
        Unregister-RASBackupTask | Out-Null
        Register-RASBackupTask -RasAdminUsername $RasAdminUsername -RasAdminPassword $RasAdminPassword -RASSettingsOutDir $RASSettingsOutDir -Weekly $Weekly
    }
}


#Configure logging
Function log
{
   param(
        [Parameter(Mandatory=$True,Position=0)]
        [string]$message,
        [switch]$error
   )

    if ($error)
    {
        Write-Host "`n$(get-date -f o)  $message" -ForegroundColor Red
    }
    else
    {
        "`n$(get-date -f o)  $message" 
    }
}

#Configure registering task
Function Register-RASBackupTask
{
    Param(
		[string]$RasAdminUsername,
        [securestring]$RasAdminPassword,
        [string]$RASSettingsOutDir,
        [bool]$Weekly = $false
	)

    try
    {
        $RASPassword = "";
        
        if (-Not $RasAdminUsername)
		{
			$RasAdminUsername = Read-Host -Prompt "RasAdminUsername"
		}

        if (-Not $RasAdminPassword)
		{
			$RasAdminPassword = Read-Host -Prompt "RasAdminPassword" -AsSecureString
            $RASPassword = $RasAdminPassword | ConvertFrom-SecureString
		}
        else
        {
            $RASPassword = $RasAdminPassword | convertfrom-securestring
        }

        if (-Not $RASSettingsOutDir)
		{
			$RASSettingsOutDir = $scriptdir
		}

        $filepath = "$scriptdir\ScheduleRASBackupSettings.txt"

        log "Saving RAS settings to: $($filepath)"

        $OFS = "`r`n"
        $settings = "username = $($RasAdminUsername)" + $OFS + "password = $($RASPassword)" + $OFS + "outputDir = $($RASSettingsOutDir| foreach {$_ -replace "\\", "/"})"
        $settings | out-file $filepath


        log "Registering Windows Task Schedule"

        $taskName = "Parallels RAS Settings Backup"
        $taskCmd = '-command "' + $scriptdir + '\ScheduleRASBackup.ps1" -ExportSettings'
        $TaskPassword = (New-Object PSCredential "user",$RasAdminPassword).GetNetworkCredential().Password

        if ($Weekly -eq "$true")
        {
            $action = New-ScheduledTaskAction -Execute 'Powershell.exe' -Argument "-NoProfile -WindowStyle Hidden $($taskCmd)"

            # set trigger for weekly job run
            $trigger =  New-ScheduledTaskTrigger -Weekly -At 00:01 -DaysOfWeek Sunday

            Register-ScheduledTask -Action $action -Trigger $trigger -TaskName $taskName -Description "Weekly backup Parallels RAS settings" -User $RasAdminUsername -Password $TaskPassword -RunLevel Highest
        }
        else
        {
            $taskParams = @("/Create",
                "/TN", $taskName, 
                "/SC", "monthly", 
                "/D",  1, 
                "/ST", "00:01", 
                "/TR", "Powershell.exe -NoProfile -WindowStyle Hidden -Command '$scriptdir\ScheduleRASBackup.ps1' -ExportSettings",
                "/RU", $RasAdminUsername,
                "/RP", $TaskPassword,
                "/RL", "HIGHEST"
                "/F" ); #force

            # supply the command arguments and execute  
            schtasks.exe $taskParams
        }
        
        log "Parallels RAS settings backup task has been registered successfully."
    }
    Catch
    {
        log "$($_.Exception.Message)" -error
    }
}

#Configure un-registering task
Function Unregister-RASBackupTask
{
    try
    {
        $taskName = "Parallels RAS Settings Backup"
        
        $task = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
        
        if ($task)
        {
            log "Un-registering 'Parallels RAS Settings Backup' Task"

            Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
            
            log "Parallels RAS settings backup task un-registered successfully."
        }
    }
    Catch
    {
        log "$($_.Exception.Message)" -error
    }
}

Function Start-ExportSettings
{
    try
    {
        $settingsPath = "$($scriptdir)\ScheduleRASBackupSettings.txt"

        # check if settings file exists
        if(![System.IO.File]::Exists($settingsPath)){
            # file with path $settingsPath doesn't exist

            throw "RAS PowerShell settings file was not found!"
        }

        #load settings from file
        log "Loading settings from $($settingsPath)"

        $psSett = get-content $settingsPath | ConvertFrom-StringData

        $RasAdminUsername = $psSett.username
        $RasAdminPassword = $psSett.password | convertto-securestring

        $timestamp  = Get-Date -UFormat "%d-%m-%Y %H.%M.%S"
        $ExportPath = $psSett.outputDir + "/rasbackup-" + $timestamp + ".dat2"

        if (-Not $RasAdminUsername)
        {
	        throw "Failed to load RAS Admin username setting."
        }

        if (-Not $RasAdminPassword)
        {
	        throw "Failed to load RAS Admin password setting."
        }

        if (-Not $ExportPath)
        {
	        throw "Failed to load path where to save the settings."
        }
        
        #Import-Module ./PSAdmin.dll
        Import-Module C:\Projects\RAS\Bin\Debug-Unicode64\appserver\binaries\PSAdmin.dll

        #Establish a connection with Parallels RAS
        log "Creating RAS session"
        New-RASSession $RasAdminUsername -Password $RasAdminPassword

        #Export RAS settings to pre-defined path
        log "Exporting Parallels RAS settings to: $($ExportPath)"
        Invoke-ExportSettings -FilePath $ExportPath

        #End the current RAS session.
        log "Ending RAS session"
        Remove-RASSession

        log "Parallels RAS settings export has completed successfully."
    }
    Catch
    {
        log "$($_.Exception.Message)" -error
    }
}


$scriptpath = $MyInvocation.MyCommand.Path
$scriptdir = Split-Path $scriptpath


Main #run the script 