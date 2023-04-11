<# 
.SYNOPSIS 
    Parallels RAS script to schedule RDS template updates 
.DESCRIPTION 
    The script will look at a RDS group that has a template assigned. Any user sessions will be logged off, the group disabled.      
    The template will then be put into maintenance mode and then maintenance mode will be exited with the recreate all instances switch. 
.PARAMETER   
    Parameter Name: None 
.INPUTS 
    None 
.OUTPUTS 
    Log file stored in C:\PS\computername.log 
.NOTES 
    Version: 1.1
    Author: Paul Fisher, Freek Berson  
    Creation Date: 11/04/23 
    Purpose/Change: Added error handling
#> 

#Variables 
$userName = "ENTER_USERNAME " 
$plainPassword = "ENTER_PASSWORD " 
$securePassword = $plainPassword | ConvertTo-SecureString -AsPlainText -Force 
$RASServer = "ENTER_A_RAS_CONNECTION_BROKER_HOSTNAME " 
$templateName = "ENTER_TEMPLATE_NAME " 
$groupID = 2 
$logDirectory = "C:\PS\" 
$logfile = "C:\PS\TemplateShedule_" + $templateName + "_" + $(get-date -f yyyy-MM-dd) + ".log"
$retryLimitLogOffSession = 30
$retryLimitRemoveMemberServers = 30
$retryLimitTemplateStatusUpdate = 30
$retryLimitTemplateExitMaintenance = 30
 
#Import modules 
Import-Module 'C:\Program Files (x86)\Parallels\ApplicationServer\Modules\RASAdmin\RASAdmin.psd1' 

#Create log file if it does not exist 
If (!(test-path -PathType Container $logDirectory)) { New-Item -ItemType Directory -Path $logDirectory } 
 
#Function to write to logfile 
function writeLog { 
    Param ([string]$logString) 
    $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss") 
    $logMessage = "$Stamp $logString" 
    Add-content $logFile -value $logMessage 
    Write-Output $logMessage 
} 

#Function that returns active session for a group ID
Function GetNumberOfRASRDSSessions {
    param ($groupIDinput)
    try {
        return ((Get-RASRDSession -Source RDS -GroupId $groupIDinput | Select-Object -Property User, SessionHostId, SessionID).count)
    }
    Catch {
        writeLog "ERROR: while running Get-RASRDSession CmdLet"
        writeLog $_
    }
}

#Function that returns the number of member servers of an RDS Group
Function GetNumberOfMemberServers {
    param ($groupID)
    try {
        return ((Get-RASRDSGroup -Id $groupID | Select-Object RDSIds).RDSIds.Count)
    }
    Catch {
        writeLog "ERROR: while running Get-RASRDSGroup CmdLet"
        writeLog $_
    }
}

#Function that returns status of a template
Function GetRASVDITemplateStatus {
    param ($templateNameInput)
    try {
        return ((Get-RASVDITemplateStatus -Name $templateNameInput | Select-Object Status).status)
    }
    Catch {
        writeLog "ERROR: while running Get-RASVDITemplateStatus CmdLet"
        writeLog $_
    }
}

#Open new RAS PowerShell session 
writeLog "Setting up RAS Farm connection...."
try {
    New-RASSession -Username $userName -Password $securePassword -Server $RASServer
}
Catch {
    writeLog "ERROR: opening new Parallels RAS PowerShell Session"
    writeLog $_
}

#Log off all users 
Foreach ($session in Get-RASRDSession -Source RDS -GroupId $groupID | Select-Object -Property User, SessionHostId, SessionID) { 
    writeLog "Logging off User: $session..." +$session.SessionID 
    try { 
        Invoke-RASRDSSessionCmd -Command LogOff -Id $session.SessionID -RDSId $session.SessionHostId
    }
    Catch { 
        writeLog "ERROR: Logging off user $($session.User) from host id $($session.SessionHostId)"
        writeLog $_ 
    }
} 

#Wait for all sessions to be logged off 
While (GetNumberOfRASRDSSessions($groupID) -ne 0 -or $retryLimitLogOffSession -eq 0) { 
    writeLog "Waiting for all users to be logged off..."
    Start-Sleep -s 10
    $retryLimitLogOffSession--
}
if ($retryLimitLogOffSession -eq 0) {
    writeLog "ERROR: not all users could be logged off"
    Exit
}

#Disable RDSH Group
writeLog "Disabling RDS Group......."
try { 
    Set-RASRDSGroup -Id $groupID -Enabled $False 
    Invoke-RASApply -FullSync 
}
Catch { 
    writeLog "ERROR: performing Set-RASRDSGroup Cmdlet"
    writeLog $_ 
}

#Wait for all members to be removed from group
While (GetNumberOfMemberServers($groupID) -ne 0 -or $retryLimitRemoveMemberServers -eq 0) { 
    writeLog "Waiting for all servers to be removed from the group..."
    Start-Sleep -s 10
    $retryLimitRemoveMemberServers--
}
if ($retryLimitRemoveMemberServers -eq 0) {
    writeLog "ERROR: not all servers could be removed from the group"
    Exit
}

#Enter maintenance mode 
writeLog "Entering Maintenance Mode...."
try { 
    Invoke-RASVDITemplateCmd $templateName -Command EnterMaintenance 
    Invoke-RASApply
    Start-Sleep -s 5
}
Catch { 
    writeLog "ERROR: on performing Invoke-RASVDITemplateCmd Cmdlet"
    writeLog $_ 
}

#Wait for the status update change
While (((GetRASVDITemplateStatus($templateName)) -ne "Maintenance") -or ($retryLimitTemplateStatusUpdate -eq 0)) {     
    writeLog "Waiting for template status to update..."
    Start-Sleep -s 10
    $retryLimitTemplateStatusUpdate--
}
if ($retryLimitTemplateStatusUpdate -eq 0) {
    writeLog "ERROR: template status could not be set to 'Maintenance'"
    Exit
}

#Exit maintenance mode and recreate desktops 
writeLog "Exit maintenance mode..." 
try { 
    Invoke-RASVDITemplateCmd -Command ExitMaintenance -Name $templateName -ForceStopUpdateDesktops 
    Invoke-RASApply
    Start-Sleep -s 5
}
Catch { 
    writeLog "ERROR: performing ExitMaintenance"
    writeLog $_ 
}
try { 
    Invoke-RASVDITemplateCmd -Command RecreateDesktops -Name $templateName -RecreateAllDesktops 
    Invoke-RASApply
    Start-Sleep -s 5
}
Catch { 
    writeLog "ERROR: performing RecreateDesktops"
    writeLog $_ 
}

#Wait for the cloning to complete
While (((GetRASVDITemplateStatus($templateName)) -ne "Created") -or ($retryLimitTemplateExitMaintenance -eq 0)) {     
    writeLog "Waiting for template status to update..."
    Start-Sleep -s 10
    $retryLimitTemplateExitMaintenance--
}
if ($retryLimitTemplateStatusUpdate -eq 0) {
    writeLog "ERROR: template status could not be set to 'Created'"
    Exit
}

#Re-enable group 
writeLog "Enabling group...." 
try { 
    Set-RASRDSGroup -Id $groupID -Enabled $True 
    Invoke-RASApply -FullSync 
}
Catch { 
    writeLog "ERROR: Enabling group"
    writeLog $_ 
}

#Clean up session and close
writeLog  "Update complete of template " +$templateName
try { 
    Remove-RASSession 
}
Catch { 
    writeLog "ERROR: on remove-RASSession CmdLet"
    writeLog $_ 
}
