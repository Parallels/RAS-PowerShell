## ==================================================================
##
## Copyright (c) 2005-2019 Parallels Software International, Inc.
## Released under the terms of MIT license (see LICENSE for details)
##
## ==================================================================

<#  
.SYNOPSIS  
    Installs the RAS Secure Client Gateway prerequisites
.DESCRIPTION  
    Installs the RAS Secure Client Gateway prerequisites
.NOTES  
    File Name  : Install-SecureClientGateway.ps1
    Author     : Freek Berson - rdsgurus.com
.EXAMPLE
    .\Install-SecureClientGateway.ps1
#>

CLS

#Reading Variables
$adDomainName = $args[0]
$RasAdminPassword = $args[1]
$RasAdminUser = $args[2] + '@' + $args[0]
$hostNamePrefixPublishingAgent =  $args[3]
$PrimaryPublishingAgent = $hostNamePrefixPublishingAgent  + '01'

#Configure logging
function log
{
   param([string]$message)
   "`n`n$(get-date -f o)  $message" 
}

#Create a credential
log "Creating credentials"
$secAdminPassword = ConvertTo-SecureString $RasAdminPassword -AsPlainText -Force
$mycreds = New-Object System.Management.Automation.PSCredential ($RasAdminUser, $secAdminPassword)

#Impersonate user
log "Impersonate user '$AdminUser'"
.\New-ImpersonateUser.ps1 -Credential $mycreds

#Install .NET
Log "Install .NET Framework core"
Install-WindowsFeature Net-Framework-Core -source C:\Windows\WinSxS 

#Create Firewall Rules
log "Create Firewall Rules"
New-NetFirewallRule -DisplayName "Allow TCP 80, 81, 135, 443, 445 and 20009, 200020, 49179 for RAS Administration" -Direction Inbound -Action Allow -Protocol TCP -LocalPort 80, 81, 135, 443, 445, 20009, 20020, 49179
New-NetFirewallRule -DisplayName "Allow UDP 20009,20020" -Direction Inbound -Action Allow -Protocol UDP -LocalPort 20009,20020

#Disable UAC & Sharing Wizard to allow Remote Install of RAS Agent
Set-ItemProperty -Path REGISTRY::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System -Name ConsentPromptBehaviorAdmin -Value 0
Set-ItemProperty -Path REGISTRY::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\policies\system -Name EnableLUA -Value 0
Set-ItemProperty -Path REGISTRY::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Folder\SharingWizardOn -Name CheckedValue -Value 0

sleep -Seconds 3

#Force reboot to complete SCG Install
shutdown -r -f -t 1

log "All Done"
