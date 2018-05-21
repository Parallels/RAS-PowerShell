// ==================================================================
//
// Copyright (c) 2005-2018 Parallels Software International, Inc.
// Released under the terms of MIT license (see LICENSE for details)
//
// ==================================================================

<#  
.SYNOPSIS  
    RAS PowerShell VDI Examples
.DESCRIPTION  
    Examples to demonstrates how to add a VDI Host, and publish a desktop from a VDI Template.
.NOTES  
    File Name  : VDISample.ps1
    Author     : www.parallels.com
.EXAMPLE
    .\VDISample.ps1
#>

CLS


#Pre-set Params
$VDIServer      = "vdi.company.dom" 	   					#(replace 'vdi.company.dom' with a valid FQDN, computer name, or IP address).
$VDIAgent       = "vdiagent.company.dom"   					#(replace 'vdiagent.company.dom' with a valid FQDN, computer name, or IP address).
$VMID           = "564d5e6f-3fad-bcf9-7c6b-bac9f212713d" 	#(replace with a valid virtual machine ID)
$TemplateName   = "Win8template"
$GstNamePrefix  = "Win8-"
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


Import-Module PSAdmin

#Establish a connection with Parallels RAS (NB. User will be prompted for Username and Password)
log "Creating RAS session"
New-RASSession

###### FARM CONFIGURATION ######

#Add a VDI Host server.
log "Adding a new VDI Host server"
$VDIHost = New-VDIHost -Server $VDIServer -VDIType VmwareESXi6_0 -VDIUsername root -VDIAgentOStype Appliance -VDIAgent $VDIAgent -Username root

#Apply all settings. This cmdlet performs the same action as the Apply button in the RAS console.
log "Appling settings"
Invoke-Apply

#Get the list of Virtual Machines through the RAS VDI Host Agent
### NB. Make sure to allow some time before calling Get-VM, to let the host agent successfully connect ###
log "Getting the list of virtual machines through the RAS VDI Host Agent"
Get-VM -VDIHostId $VDIHost.Id

#Convert a VM to VDI Guest
log "Converting a VM to a VDI Guest"
New-VDIGuest -VDIHostId $VDIHost.Id -Id $VMID


#Convert a VDIGuest to VDITemplate
log "Converting the VDI Guest to a VDI Template"
$vmTemplate = New-VDITemplate -VDIHostId $VDIHost.Id -VDIGuestId $VMID -TemplateName $TemplateName -GuestNamePrefix $GstNamePrefix -MaxGuests 5 -PreCreatedGuests 2 `
				-ImagePrepTool RASPrep -OwnerName $Owner -Organization $Organization -JoinDomain $Domain -CloneMethod LinkedClone -TargetOU $TargetOU

###### PUBLISHING CONFIGURATION ######

#Add published desktop making use of the VDI Template.
log "Adding published desktop using the VDI template"
New-PubVDIDesktop -Name $PubDeskName -ConnectTo SpecificRASTemplate -VDITemplate $vmTemplate -Persistent $true

#Apply all settings. This cmdlet performs the same action as the Apply button in the RAS console.
log "Appling settings"
Invoke-Apply

#End the current RAS session.
log "Ending RAS session"
Remove-RASSession

log "All Done"