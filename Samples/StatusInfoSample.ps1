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
    Examples to demonstrates how to retrieve status information of an RDS, GW, PA, Site and VDI Host.
.NOTES  
    File Name  : StatusInfoSample.ps1
    Author     : www.parallels.com
.EXAMPLE
    .\StatusInfoSample.ps1
#>

CLS


#Pre-set Params
$GWServer 	= "gw.company.dom" 		  #(replace 'gw.company.dom' with a valid FQDN, computer name, or IP address).
$RDSServer  = "rds.company.dom" 	  #(replace 'rds.company.dom' with a valid FQDN, computer name, or IP address).
$PAServer   = "pa.company.dom" 		  #(replace 'pa.company.dom' with a valid FQDN, computer name, or IP address).
$VDIServer  = "vdi.company.dom" 	  #(replace 'vdi.company.dom' with a valid FQDN, computer name, or IP address).
$VDIAgent   = "vdiagent.company.dom"  #(replace 'vdiagent.company.dom' with a valid FQDN, computer name, or IP address).
$SiteServer = "site.company.dom" 	  #(replace 'site.company.dom' with a valid FQDN, computer name, or IP address).
$SiteName   = "MyRASSite"			  #(replace site name 'MyRASSite' with a more specific name)


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


###### RDS status info ######

#Adding a RAS RDS
log "Adding new RD Session Host server"
New-RDS -Server $RDSServer

#Apply all settings. This cmdlet performs the same action as the Apply button in the RAS console.
log "Appling settings"
Invoke-Apply

# Get RDS status info
log "Getting RAS RD Session host status information"
Get-RDSStatus -Server $RDSServer


###### GW status info ######

#Adding a RAS GW
log "Adding new RAS Gateway server" 
New-GW -Server $GWServer

#Apply all settings. This cmdlet performs the same action as the Apply button in the RAS console.
log "Appling settings"
Invoke-Apply

# Get GW status info
log "Getting RAS Gateway status information"
Get-GWStatus -Server $GWServer


###### PA status info ######

#Adding a RAS PA
log "Adding new RAS PA server"
New-PA -Server $PAServer

#Apply all settings. This cmdlet performs the same action as the Apply button in the RAS console.
log "Appling settings"
Invoke-Apply

# Get PA status info
log "Getting RAS PA status information"
Get-PAStatus -Server $PAServer


###### Site status info ######

#Adding a RAS Site
log "Adding new RAS Site"
New-Site -Server $SiteServer -Name $SiteName

#Apply all settings. This cmdlet performs the same action as the Apply button in the RAS console.
log "Appling settings"
Invoke-Apply

# Get Site status info
log "Getting RAS Site status information"
Get-SiteStatus -Server $SiteServer


###### VDI Host status info ######

#Add a VDI Host servers.
$VDIHost = New-VDIHost -Server $VDIServer -VDIType VmwareESXi6_0 -VDIUsername root -VDIAgentOStype Appliance -VDIAgent $VDIAgent -Username root

#Apply all settings. This cmdlet performs the same action as the Apply button in the RAS console.
log "Appling settings"
Invoke-Apply

# Get VDI Host status info
log "Getting VDI Host status information"
Get-VDIHostStatus -Id $VDIHost.Id


#End the current RAS session.
log "Ending RAS session"
Remove-RASSession

log "All Done"