## ==================================================================
##
## Copyright (c) 2005-2024 Parallels Software International, Inc.
## Released under the terms of MIT license (see LICENSE for details)
##
## ==================================================================

<#  
.SYNOPSIS  
    RAS PowerShell VDI Template Examples
.DESCRIPTION  
    Examples to demonstrates how to add a Provider, and publish a desktop from a VDI Template.
.NOTES  
    File Name  : VDITemplateSample.ps1
    Author     : www.parallels.com
.EXAMPLE
    .\VDITemplateSample.ps1
#>

[CmdletBinding()]
    Param(
        [parameter(Mandatory=$true)][ValidateNotNullOrEmpty()] [string]$AdminUsername,
		[parameter(Mandatory=$true)][ValidateNotNullOrEmpty()] [SecureString]$AdminPassword
    )

#Pre-set Params
$VDIServer      = "vdi.company.dom" 	   					#(replace 'vdi.company.dom' with a valid FQDN, computer name, or IP address).
$VDIAgent       = "vdiagent.company.dom"   					#(replace 'vdiagent.company.dom' with a valid FQDN, computer name, or IP address).
$VMID           = "564d5e6f-3fad-bcf9-7c6b-bac9f212713d" 	#(replace with a valid virtual machine ID)
$TemplateName   = "Win10template"
$VMNameFormat   = "Win10-%ID:3%"
$Owner 			= "Owner"
$Organization	= "Parallels"
$Domain			= "company.dom"
$TargetOU 		= "OU=VDI,DC=dom,DC=company"
$PubDeskName	= "VDIDesktop"

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

###### FARM CONFIGURATION ######

#Add a Provider.
log "Adding a new Provider"
$Provider = New-RASProvider -Server $VDIServer -VMwareESXi -VMwareESXiVersion v6_5 -VDIAgent $VDIAgent -ProviderUsername root -ProviderPassword $AdminPassword -Name "vdi"

#Get the list of Virtual Machines through the RAS Provider Agent
### NB. Make sure to allow some time before calling Get-RASVM, to let the host agent successfully connect ###
log "Getting the list of virtual machines through the RAS Provider Agent"
Get-RASVM -ProviderId $Provider.Id

#Create a new RAS VDI Template
log "Create a new RAS VDI Template"
$vmTemplate = New-RASTemplate -ProviderId $Provider.Id -VMId $VMID -Name $TemplateName -ImagePrepTool RASPrep -OwnerName $Owner -Organization $Organization -Domain $Domain `
                              -CloneMethod LinkedClone -DomainPassword $AdminPassword -Administrator $AdminUsername -AdminPassword $AdminPassword -ObjType VDITemplate -DomainOrgUnit $TargetOU

#Add a new VDI Host Pool
log "Add a new VDI Pool"
$VDIHostPool = New-RASVDIHostPool -Name "VDIPool" -MaxHosts 5 -PreCreatedHosts 2 -HostName $VMNameFormat -ProvisioningType Template -TemplateId $vmTemplate.Id

#Apply all settings. This cmdlet performs the same action as the Apply button in the RAS console.
log "Appling settings"
Invoke-RASApply

###### PUBLISHING CONFIGURATION ######

#Add published desktop making use of the VDI Template.
log "Adding published desktop using the VDI template"
New-RASPubVDIDesktop -Name $PubDeskName -ConnectTo AnyGuest -VDITemplateId $vmTemplate.Id -VDIHostPoolId $VDIHostPool.Id -Persistent $true

#Apply all settings. This cmdlet performs the same action as the Apply button in the RAS console.
log "Appling settings"
Invoke-RASApply

#End the current RAS session.
log "Ending RAS session"
Remove-RASSession

log "All Done"
