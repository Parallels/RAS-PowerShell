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
    Examples to demonstrates how to retrieve status information of an RDS, GW, PA, Site and Provider.
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


Import-Module RASAdmin

#Establish a connection with Parallels RAS (NB. User will be prompted for Username and Password)
log "Creating RAS session"
New-RASSession


###### RDS status info ######

#Adding a RAS RDS
log "Adding new RD Session Host server"
New-RASRDS -Server $RDSServer

#Apply all settings. This cmdlet performs the same action as the Apply button in the RAS console.
log "Appling settings"
Invoke-RASApply

log "Sleeping for 5 seconds"
Start-Sleep -Seconds 5

# Get RDS status info
log "Getting RAS RD Session host status information"
Get-RASRDSStatus -Server $RDSServer


###### GW status info ######

#Adding a RAS GW
log "Adding new RAS Gateway server" 
New-RASGW -Server $GWServer

#Apply all settings. This cmdlet performs the same action as the Apply button in the RAS console.
log "Appling settings"
Invoke-RASApply

# Get GW status info
log "Getting RAS Gateway status information"
Get-RASGWStatus -Server $GWServer


###### PA status info ######

#Adding a RAS PA
log "Adding new RAS PA server"
New-RASPA -Server $PAServer

#Apply all settings. This cmdlet performs the same action as the Apply button in the RAS console.
log "Appling settings"
Invoke-RASApply

# Get PA status info
log "Getting RAS PA status information"
Get-RASPAStatus -Server $PAServer


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
$Provider = New-RASProvider -Server $VDIServer -Type VmwareESXi6_0 -VDIUsername root -VDIAgent $VDIAgent -Username $AdminUsername -Password $AdminPassword

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
