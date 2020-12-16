## ==================================================================
##
## Copyright (c) 2005-2019 Parallels Software International, Inc.
## Released under the terms of MIT license (see LICENSE for details)
##
## ==================================================================

<#  
.SYNOPSIS  
    RAS PowerShell VDI Examples
.DESCRIPTION  
    Examples to demonstrates how to add a Provider, and publish a desktop from a VDI Template.
.NOTES  
    File Name  : VDISample.ps1
    Author     : www.parallels.com
.EXAMPLE
    .\VDISample.ps1
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
$TemplateName   = "Win8template"
$GstNameFormat  = "Win8-%ID:3%"
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
$Provider = New-RASProvider -Server $VDIServer -Type VmwareESXi6_0 -VDIUsername root -VDIAgent $VDIAgent -Username $AdminUsername -Password $AdminPassword

#Apply all settings. This cmdlet performs the same action as the Apply button in the RAS console.
log "Appling settings"
Invoke-RASApply

#Get the list of Virtual Machines through the RAS Provider Agent
### NB. Make sure to allow some time before calling Get-RASVM, to let the host agent successfully connect ###
log "Getting the list of virtual machines through the RAS Provider Agent"
Get-RASVM -ProviderId $Provider.Id

#Convert a VM to VDI Guest
log "Converting a VM to a VDI Guest"
New-RASVDIGuest -ProviderId $Provider.Id -Id $VMID


#Convert a VDIGuest to VDITemplate
log "Converting the VDI Guest to a VDI Template"
$vmTemplate = New-RASVDITemplate -ProviderId $Provider.Id -VDIGuestId $VMID -Name $TemplateName -GuestNameFormat $GstNameFormat -MaxGuests 5 -PreCreatedGuests 2 `
				-ImagePrepTool RASPrep -OwnerName $Owner -Organization $Organization -Domain $Domain -CloneMethod LinkedClone -TargetOU $TargetOU

###### PUBLISHING CONFIGURATION ######

#Add published desktop making use of the VDI Template.
log "Adding published desktop using the VDI template"
New-RASPubVDIDesktop -Name $PubDeskName -ConnectTo SpecificRASTemplate -VDITemplate $vmTemplate -Persistent $true

#Apply all settings. This cmdlet performs the same action as the Apply button in the RAS console.
log "Appling settings"
Invoke-RASApply

#End the current RAS session.
log "Ending RAS session"
Remove-RASSession

log "All Done"
