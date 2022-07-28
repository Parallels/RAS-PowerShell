## ==================================================================
##
## Copyright (c) 2005-2019 Parallels Software International, Inc.
## Released under the terms of MIT license (see LICENSE for details)
##
## ==================================================================

<#  
.SYNOPSIS  
    RAS PowerShell Basic Examples
.DESCRIPTION  
    Examples to demonstrates how to start a session, add major objects to a site, publish a desktop, activate a license, apply the changes, and finally end the session.
.NOTES  
    File Name  : BasicSample.ps1
    Author     : www.parallels.com
.EXAMPLE
    .\BasicSample.ps1
#>

CLS


#Pre-set Params
$GWServer = "gw.company.dom" 		#(replace 'gw.company.dom' with a valid FQDN, computer name, or IP address).
$RDSServer = "rds.company.dom" 		#(replace 'rds.company.dom' with a valid FQDN, computer name, or IP address).
$PubDeskName = "PubDesktop"			#(replace with a more specific name).
$LicenseKey = "YOUR-LICENSE-KEY"	#(replace with a valid Parallels RAS License key).


#Configure logging
function log
{
   param([string]$message)
   "`n`n$(get-date -f o)  $message" 
}


Import-Module RASAdmin

#Establish a connection with Parallels RAS (NB. User will be prompted for Username and Password)
log "Creating RAS session"
New-RASSession

#Add a RAS Secure Gateway.
log "Adding new RAS Secure Gateway"
New-RASGateway -Server $GWServer

#Add an RD Session Host server.
log "Adding new RD Session Host server"
New-RASRDS -Server $RDSServer

#Add a published desktop.
log "Adding new RDS published desktop"
New-RASPubRDSDesktop -Name $PubDeskName

#Activate Parallels RAS as a trial (you will have to provide a valid Parallels My Account email and password).
log "Activating Parallels RAS as a trial"
Invoke-RASLicenseActivate

#Activate Parallels RAS License. If you have a valid Parallels RAS License key use the below license activation
#(you will have to provide a valid Parallels My Account email and password)
#log "Activating Parallels RAS"
#Invoke-RASLicenseActivate -Key $LicenseKey

#Apply all settings. This cmdlet performs the same action as the Apply button in the RAS console.
log "Appling settings"
Invoke-RASApply

#End the current RAS session.
log "Ending RAS session"
Remove-RASSession

log "All Done"