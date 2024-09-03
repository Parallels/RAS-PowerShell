## ==================================================================
##
## Copyright (c) 2005-2019 Parallels Software International, Inc.
## Released under the terms of MIT license (see LICENSE for details)
##
## ==================================================================

<#  
.SYNOPSIS  
    RAS PowerShell RAS Agent Examples
.DESCRIPTION  
    Examples to demonstrates how to manage RAS Agent operations.
.NOTES  
    File Name  : RASAgentSample.ps1
    Author     : www.parallels.com
.EXAMPLE
    .\RASAgentSample.ps1
#>

CLS


#Pre-set Params
$RDSServer1 = "rds1.company.dom" 	#(replace 'rds1.company.dom' with a valid FQDN, computer name, or IP address).
$RDSServer2 = "rds2.company.dom" 	#(replace 'rds2.company.dom' with a valid FQDN, computer name, or IP address).


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

#Add two RD Session Host servers.
log "Adding two RD Session host servers"
$rds = New-RASRDSHost -Server $RDSServer1
New-RASRDSHost -Server $RDSServer2

#Apply all settings. This cmdlet performs the same action as the Apply button in the RAS console.
log "Appling settings"
Invoke-RASApply

# Get all available RAS Agents information
log "Getting all RAS Agents information"
Get-RASAgent

# Get the RAS Agent information (by server name 'rds1.company.dom')
log "Get RAS Agent information (for first RD Session host only)"
Get-RASAgent -Server $RDSServer1

#Update RDS RAS Agent (by input object)
log "Updating the RAS Agent (of first RD Session host)"
$rdsAgent = Get-RASAgent -Server $RDSServer1
Update-RASAgent -InputObject $rdsAgent

#Update RDS RAS Agent (by server name 'rds2.company.dom')
log "Updating the RAS Agent (of second RD Session host)"
Update-RASAgent -Server $RDSServer2

#Apply all settings. This cmdlet performs the same action as the Apply button in the RAS console.
log "Appling settings"
Invoke-RASApply

# Get all RAS Agents information of type RDS
log "Getting all RAS Agents information (by server type)"
Get-RASAgent -ServerType RDS

# Removing RAS Agent (by server name 'rds2.company.dom')
log "Removing RAS Agent (of second RD Session host)" 
Remove-RASAgent -Server $RDSServer2

#Apply all settings. This cmdlet performs the same action as the Apply button in the RAS console.
log "Appling settings"
Invoke-RASApply

# Get all available RAS Agents information
log "Getting all RAS Agents information"
Get-RASAgent

#End the current RAS session.
log "Ending RAS session"
Remove-RASSession

log "All Done"