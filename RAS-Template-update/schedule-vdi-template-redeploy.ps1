<# 
.SYNOPSIS 
    Parallels RAS script to schedule VDI template updates 
.DESCRIPTION 
    The script will look at a VDI template.  All running VDI guests will be shutdown which will also deal with any user sessions.      
    The template will then be put into maintenance mode and then maintenance mode will be exited with the recreate all instances switch. 
.PARAMETER   
    Parameter_Name: None 
.INPUTS 
    None 
.OUTPUTS 
    Log file stored in C:\PS\computername.log 
.NOTES 
    Version: 1.1
    Author: Paul Fisher, Freek Berson  
    Creation Date: 09/11/22 
    Purpose/Change: Added error handling
#> 

#Variables 
$userName = "ENTER_USERNAME " 
$plainPassword = "ENTER_PASSWORD " 
$securePassword = $plainPassword | ConvertTo-SecureString -AsPlainText -Force 
$RASServer = "ENTER_A_RAS_CONNECTION_BROKER_HOSTNAME " 
$templateName = "ENTER_TEMPLATE_NAME " 
$logDirectory = "C:\PS\" 
$logfile = "C:\PS\TemplateShedule_vdi_"+$templateName+"_"+$(get-date -f yyyy-MM-dd)+".log"
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

#Function that returns status of a template
Function Get-RASVDITemplateStatus {
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

#Shutdown all VMs (This section shutdows the VMs and therefore Logs the users off too) 
$TempObj = Get-RASVDITemplate -Name $templateName 
Foreach ($RASVDIInstance in Get-RASVDIDesktop -VDITemplateObj $TempObj) 
{ 
    writeLog "Shuting down VM:  $RASVDIInstance.ComputerName " 
    try { 
        Stop-RASVM -Id $RASVDIInstance.Id -ProviderId $RASVDIInstance.ProviderId
    }
        Catch { 
            writeLog "ERROR: Shuting down VM with ID $($RASVDIInstance.Id) "
            writeLog $_ 
        }
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

#Clean up session and close
writeLog  "Update complete of template " +$templateName
try { 
    Remove-RASSession 
}
Catch { 
    writeLog "ERROR: on remove-RASSession CmdLet"
    writeLog $_ 
}