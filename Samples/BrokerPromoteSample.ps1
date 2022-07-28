## ==================================================================
##
## Copyright (c) 2005-2019 Parallels Software International, Inc.
## Released under the terms of MIT license (see LICENSE for details)
##
## ==================================================================

<#  
.SYNOPSIS  
    RAS PowerShell Connection Broker Promotion Example
.DESCRIPTION  
    Example to demonstrates how to promote a backup Connection Broker (Broker) in the eventuallity that the primary Connection Broker is down.
.NOTES  
    File Name  : BrokerPromoteSample.ps1
    Author     : www.parallels.com
.EXAMPLE
    .\BrokerPromoteSample.ps1
#>

CLS

$LicenseKey		    = "YOUR-LICENSE-KEY"		  #(replace with a valid Parallels RAS License key).
$LicenceEmail	    = "myaccount@email.com"	  #(replace with a valid Parallels My Account email address)
$BrokerBackupServer = "broker.company.dom"      #(replace 'broker.company.dom' with a valid FQDN, computer name, or IP address).

#Configure logging
function log
{
   param([string]$message)
   "`n`n$(get-date -f o)  $message" 
}

Import-Module RASAdmin

### Setting up the environment ###

#Establish a connection with Parallels RAS (NB. User will be prompted for Username and Password)
log "Creating RAS session"
New-RASSession

#Activate Parallels RAS with a valid Key (you will have to provide a valid Parallels My Account password for provided email).
log "Activating Parallels RAS as a trial"
Invoke-RASLicenseActivate -Email $LicenceEmail -Key $LicenseKey

#Add the backup Connection Broker
log "Adding the backup Connection Broker server"
New-RASBroker -Server $BrokerBackupServer

#Apply all settings. This cmdlet performs the same action as the Apply button in the RAS console.
log "Appling settings"
Invoke-RASApply

#Get the list of Connection Brokers. The $BrokerList variable receives an array of objects of type Broker.
log "Retrieving the list of Connection Brokers"
$BrokerList = Get-RASBroker

log "Print the list of Connection Brokers retrieved"
Write-Host ($BrokerList | Format-Table | Out-String)

#End the current RAS session.
log "Ending RAS session"
Remove-RASSession

### Setting up the environment ###


### Testing Connection Broker Promotion scenario ###
### In the eventuality that the primary Connection Broker is down ###

#Establish a connection with Parallels RAS (NB. User will be prompted for Username and Password)
log "Creating RAS session (with backup Broker)"
New-RASSession -Server $BrokerBackupServer -Force

#Get the backup Connection Broker.
log "Retrieving the backup Connection Broker"
$BackupBroker = Get-RASBroker -Server $BrokerBackupServer

#Promote backup Connection Broker to primary
#You will have to provide a valid Parallels My Account password for provided email
log "Promote backup Connection Broker to primary"
Invoke-RASBrokerPromoteToPrimary -Id $BackupBroker.Id -Email $LicenceEmail

#After the Connection Broker promotion to Primary the session is logged out, then a new session needs to be created.
log "Creating RAS session (with backup Connection Broker)"
New-RASSession -Server $BrokerBackupServer

#Apply all settings. This cmdlet performs the same action as the Apply button in the RAS console.
log "Appling settings"
Invoke-RASApply

#Get the list of Connection Brokers. Verify that the priority values changed (Priority '0' is primary).
log "Retrieving the list of Connection Brokers"
Get-RASBroker

#End the current RAS session.
log "Ending RAS session"
Remove-RASSession

### Testing Connection Broker Promotion scenario ###