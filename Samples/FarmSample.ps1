// ==================================================================
//
// Copyright (c) 2005-2018 Parallels Software International, Inc.
// Released under the terms of MIT license (see LICENSE for details)
//
// ==================================================================

<#  
.SYNOPSIS  
    RAS PowerShell Farm Examples
.DESCRIPTION  
    Examples to demonstrates how to modify objects, handle multiple objects and groups, and manage default settings.
.NOTES  
    File Name  : FarmSample.ps1
    Author     : www.parallels.com
.EXAMPLE
    .\FarmSample.ps1
#>

CLS


#Pre-set Params
$GWServer = "gw.company.dom" 		#(replace 'gw.company.dom' with a valid FQDN, computer name, or IP address).
$RDSServer1 = "rds1.company.dom" 	#(replace 'rds1.company.dom' with a valid FQDN, computer name, or IP address).
$RDSServer2 = "rds2.company.dom" 	#(replace 'rds2.company.dom' with a valid FQDN, computer name, or IP address).
$RDSServer3 = "rds3.company.dom" 	#(replace 'rds3.company.dom' with a valid FQDN, computer name, or IP address).
$RDS1Desc = "Updating RDS Host 1"	#(replace with a more specific name).
$RDSGroupName = "My RDS Group"		#(replace with a more specific name).
$RDSDefSettMaxSessions = 100		#(replace default value with preferred max sessions).
$RDSDefSettAppMonitor = $true		#(replace default value with preferred App Monitoring value (Enabeld/Disabled)).
$PubDeskName = "PubDesktop"			#(replace with a more specific name).
$LicenseKey = "YOUR-LICENSE-KEY"	#(replace with a valid Parallels RAS License key).


#Configure logging
function log
{
   param([string]$message)
   "`n`n$(get-date -f o)  $message" 
}


Import-Module PSAdmin

#Establish a connection with Parallels RAS (NB. User will be prompted for Username and Password)
log "Creating RAS session"
New-RASSession

#Add a RAS Secure Client Gateway.
log "Adding new RAS Secure Client Gateway"
New-GW -Server $GWServer

#Add the first RD Session Host server
#The $RDS1 variable receives an object of type RDS identifying the RD Session Host. 
log "Adding the first RD Session Host server"
$RDS1 = New-RDS -Server $RDSServer1

#Update the description of RD Session Host specified by the $RDS1 variable. 
log "Updating the RD Session description"
Set-RDS -InputObject $RDS1 -Description $RDS1Desc

#Add the second RD Session Host.
log "Adding the second RD Session Host server"
$RDS2 = New-RDS -Server $RDSServer2

#Get the list of RD Session Host servers. The $RDSList variable receives an array of objects of type RDS.
log "Retrieving the list of RD Session servers"
$RDSList = Get-RDS

log "Print the list of RD Session servers retrieved"
Write-Host ($RDSList | Format-Table | Out-String)

#Create an RD Session Host Group and add both RD Session Host objects to it.
log "Add an RD Session host group (with list of RD Sessions)"
New-RDSGroup -Name $RDSGroupName -RDSObject $RDSList

#Add the third RD Session Host server.
log "Adding the third RD Session Host server"
$RDS3 = New-RDS -Server $RDSServer3

#Add a previously created RDS3 to the RD Session Host Group.
log "Add a group member to the RD Session Group created"
Add-RDSGroupMember -GroupName $RDSGroupName -RDSServer $RDS3.Server

#Update default settings used to configure RD Session Host agents.
log "Updating RDS default settings"
Set-RDSDefaultSettings -MaxSessions $RDSDefSettMaxSessions -EnableAppMonitoring $RDSDefSettAppMonitor

#Add a published desktop.
log "Adding new RDS published desktop"
New-PubRDSDesktop -Name $PubDeskName

#Activate Parallels RAS as a trial (you will have to provide a valid Parallels My Account email and password).
log "Activating Parallels RAS as a trial"
Invoke-LicenseActivate

#Activate Parallels RAS License. If you have a valid Parallels RAS License key use the below license activation
#(you will have to provide a valid Parallels My Account email and password)
#log "Activating Parallels RAS"
#Invoke-LicenseActivate -Key $LicenseKey

#Apply all settings. This cmdlet performs the same action as the Apply button in the RAS console.
log "Appling settings"
Invoke-Apply

#End the current RAS session.
log "Ending RAS session"
Remove-RASSession

log "All Done"