## ==================================================================
##
## Copyright (c) 2005-2019 Parallels Software International, Inc.
## Released under the terms of MIT license (see LICENSE for details)
##
## ==================================================================

<#  
.SYNOPSIS  
    RAS PowerShell RDSTemplate Sample
.DESCRIPTION  
    Demonstrates how to add a Provider, and use RDS Templates in RDS Groups.
.NOTES  
    File Name  : RDSTemplateSample.ps1
    Author     : www.parallels.com
.EXAMPLE
    .\RDSTemplateSample.ps1
#>

[CmdletBinding()]
    Param(
        [parameter(Mandatory=$true)][ValidateNotNullOrEmpty()] [string]$AdminUsername,
		[parameter(Mandatory=$true)][ValidateNotNullOrEmpty()] [SecureString]$AdminPassword
    )

#Pre-set Params
$VDIServer      = "vdi.company.dom" 	   			#(replace 'vdi.company.dom' with a valid FQDN, computer name, or IP address).
$VDIAgent       = "vdiagent.company.dom"   			#(replace 'vdiagent.company.dom' with a valid FQDN, computer name, or IP address).
$VMID           = "564d5e6f-3fad-bcf9-7c6b-bac9f212713d" 	#(replace with a valid virtual machine ID)
$TemplateName   = "Win10template"
$VMNameFormat  = "Win10-%ID:3%"
$Owner 		= "Owner"
$Organization	= "Parallels"
$Domain		= "company.dom"
$TargetOU 	= "OU=VDI,DC=dom,DC=company"
$ComputerName	= "10.0.0.51"
$RDSGroupName 	= "My RDS Group"				#(replace with a more specific name).


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

###### VDI CONFIGURATION ######

#Add a Provider.
log "Adding a new Provider"
$Provider = New-RASProvider -Server $VDIServer -Type VmwareESXi6_0 -VDIUsername root -VDIAgent $VDIAgent -Username $AdminUsername -Password $AdminPassword

#Apply all settings. This cmdlet performs the same action as the Apply button in the RAS console.
log "Appling settings"
Invoke-RASApply

#Get the list of Virtual Machines through the RAS Provider Agent
### NB. Make sure to allow some time before calling Get-RASVM, to let the host agent successfully connect ###
log "Getting the list of virtual machines through the RAS Provider Agent"
Get-RASVM -ProviderId $Provider.Id

#Convert a VM to an RDSTemplate
log "Converting the VM to an RDSTemplate"
$rdsTemplate = New-RASVDITemplate -ProviderId $Provider.Id -VMId $VMID -Name $TemplateName -VMNameFormat $VMNameFormat -MaxVMs 5 -PreCreatedVMs 2 -VMsToCreate 1 `
-ImagePrepTool RASPrep -OwnerName $Owner -Organization $Organization -Domain $Domain -CloneMethod LinkedClone -TargetOU $TargetOU -TemplateType RDSH -ComputerName $ComputerName

#Apply all settings. This cmdlet performs the same action as the Apply button in the RAS console.
log "Appling settings"
Invoke-RASApply


###### RDS GROUP CONFIGURATION ######

#Create an RD Session Host Group and add RDS Template object to it.
log "Add an RD Session host group (with list of RD Sessions)"
New-RASRDSGroup -Name $RDSGroupName -Description "RDSTemplates Group" -UseRASTemplate -RASTemplate $rdsTemplate -WorkLoadThreshold 50 -ServersToAddPerRequest 2 `
-MaxServersFromTemplate 2 -WorkLoadToDrain 20

#Apply all settings. This cmdlet performs the same action as the Apply button in the RAS console.
log "Appling settings"
Invoke-RASApply

#End the current RAS session.
log "Ending RAS session"
Remove-RASSession

log "All Done"

