## ==================================================================
##
## Copyright (c) 2005-2019 Parallels Software International, Inc.
## Released under the terms of MIT license (see LICENSE for details)
##
## ==================================================================

<#  
.SYNOPSIS  
    RAS PowerShell PA Promotion Example
.DESCRIPTION  
    Example to demonstrates how to promote a backup publishing agent (PA) in the eventuallity that the master PA is down.
.NOTES  
    File Name  : PAPromoteSample.ps1
    Author     : www.parallels.com
.EXAMPLE
    .\PAPromoteSample.ps1
#>

CLS

$LicenseKey		  = "YOUR-LICENSE-KEY"		  #(replace with a valid Parallels RAS License key).
$LicenceEmail	  = "myaccount@email.com"	  #(replace with a valid Parallels My Account email address)
$PABackupServer   = "pa.company.dom" 		  #(replace 'pa.company.dom' with a valid FQDN, computer name, or IP address).

#Configure logging
function log
{
   param([string]$message)
   "`n`n$(get-date -f o)  $message" 
}

Import-Module PSAdmin

### Setting up the environment ###

#Establish a connection with Parallels RAS (NB. User will be prompted for Username and Password)
log "Creating RAS session"
New-RASSession

#Activate Parallels RAS with a valid Key (you will have to provide a valid Parallels My Account password for provided email).
log "Activating Parallels RAS as a trial"
Invoke-LicenseActivate -Email $LicenceEmail -Key $LicenseKey

#Add the backup PA server
log "Adding the backup PA server"
New-PA -Server $PABackupServer

#Apply all settings. This cmdlet performs the same action as the Apply button in the RAS console.
log "Appling settings"
Invoke-Apply

#Get the list of PA servers. The $PAList variable receives an array of objects of type PA.
log "Retrieving the list of PA servers"
$PAList = Get-PA

log "Print the list of PA servers retrieved"
Write-Host ($PAList | Format-Table | Out-String)

#End the current RAS session.
log "Ending RAS session"
Remove-RASSession

### Setting up the environment ###


### Testing PA Promotion scenario ###
### In the eventuallity that the master PA is down ###

#Establish a connection with Parallels RAS (NB. User will be prompted for Username and Password)
log "Creating RAS session (with backup PA)"
New-RASSession -Server $PABackupServer -Force

#Get the backup PA server.
log "Retrieving the backup PA server"
$BackupPA = Get-PA -Server $PABackupServer

#Promote backup PA to master
#You will have to provide a valid Parallels My Account password for provided email
log "Promote backup PA to master"
Invoke-PAPromoteToMaster -Id $BackupPA.Id -Email $LicenceEmail

#Apply all settings. This cmdlet performs the same action as the Apply button in the RAS console.
log "Appling settings"
Invoke-Apply

#Get the list of PA servers. Verify that the priority values changed (Priority '0' is master).
log "Retrieving the list of PA servers"
Get-PA

#End the current RAS session.
log "Ending RAS session"
Remove-RASSession

### Testing PA Promotion scenario ###