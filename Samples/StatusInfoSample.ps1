## ==================================================================
##
## Copyright (c) 2005-2019 Parallels Software International, Inc.
## Released under the terms of MIT license (see LICENSE for details)
##
## ==================================================================

<#  
.SYNOPSIS  
    RAS PowerShell Status Information Examples
.DESCRIPTION  
    Examples to demonstrates how to retrieve status information of an RDS, Gateway, Broker, Site and Provider.
.NOTES  
    File Name  : StatusInfoSample.ps1
    Author     : www.parallels.com
.EXAMPLE
    .\StatusInfoSample.ps1
#>

[CmdletBinding()]
    Param(
        [parameter(Mandatory=$true)][ValidateNotNullOrEmpty()] [string]$AdminUsername,
		[parameter(Mandatory=$true)][ValidateNotNullOrEmpty()] [SecureString]$AdminPassword
    )

#Pre-set Params
$GWServer 	    = "gw.company.dom" 		  #(replace 'gw.company.dom' with a valid FQDN, computer name, or IP address).
$RDSServer      = "rds.company.dom" 	  #(replace 'rds.company.dom' with a valid FQDN, computer name, or IP address).
$BrokerServer   = "broker.company.dom" 	  #(replace 'broker.company.dom' with a valid FQDN, computer name, or IP address).
$VDIServer      = "vdi.company.dom" 	  #(replace 'vdi.company.dom' with a valid FQDN, computer name, or IP address).
$VDIAgent       = "vdiagent.company.dom"  #(replace 'vdiagent.company.dom' with a valid FQDN, computer name, or IP address).
$SiteServer     = "site.company.dom" 	  #(replace 'site.company.dom' with a valid FQDN, computer name, or IP address).
$SiteName       = "MyRASSite"			  #(replace site name 'MyRASSite' with a more specific name)


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


###### RDS status info ######

#Adding a RAS RDS
log "Adding new RD Session Host server"
New-RASRDSHost -Server $RDSServer

#Apply all settings. This cmdlet performs the same action as the Apply button in the RAS console.
log "Appling settings"
Invoke-RASApply

log "Sleeping for 5 seconds"
Start-Sleep -Seconds 5

# Get RDS status info
log "Getting RAS RD Session host status information"
Get-RASRDSHostStatus -Server $RDSServer


###### Secure Gateway status info ######

#Adding a RAS Secure Gateway
log "Adding new RAS Gateway server" 
New-RASGateway -Server $GWServer

#Apply all settings. This cmdlet performs the same action as the Apply button in the RAS console.
log "Appling settings"
Invoke-RASApply

# Get Secure Gateway status info
log "Getting RAS Secure Gateway status information"
Get-RASGatewayStatus -Server $GWServer


###### Connection Broker status info ######

#Adding a RAS Connection Broker
log "Adding new RAS Connection Broker server"
New-RASBroker -Server $BrokerServer

#Apply all settings. This cmdlet performs the same action as the Apply button in the RAS console.
log "Appling settings"
Invoke-RASApply

# Get Connection Broker status info
log "Getting RAS Connection Broker status information"
Get-RASBrokerStatus -Server $BrokerServer


###### Site status info ######

#Adding a RAS Site
log "Adding new RAS Site"
$site = New-RASSite -Server $SiteServer -Name $SiteName

#Apply all settings. This cmdlet performs the same action as the Apply button in the RAS console.
log "Appling settings"
Invoke-RASApply

# Get Site status info
log "Getting RAS Site status information"
Get-RASSiteStatus -SiteId $site.Id


###### Provider status info ######

#Add a Provider.
$Provider = New-RASProvider -Server $VDIServer -VmwareESXi -VmwareESXiVersion v6_5 -VDIUsername root -VDIAgent $VDIAgent -Username $AdminUsername -Password $AdminPassword

#Apply all settings. This cmdlet performs the same action as the Apply button in the RAS console.
log "Appling settings"
Invoke-RASApply

# Get Provider status info
log "Getting Provider status information"
Get-RASProviderStatus -Id $Provider.Id


#End the current RAS session.
log "Ending RAS session"
Remove-RASSession

log "All Done"
