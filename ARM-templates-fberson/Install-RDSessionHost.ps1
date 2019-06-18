## ==================================================================
##
## Copyright (c) 2005-2019 Parallels Software International, Inc.
## Released under the terms of MIT license (see LICENSE for details)
##
## ==================================================================

<#  
.SYNOPSIS  
    Installs the RAS RD Session Host role & prerequisites
.DESCRIPTION  
    Installs the RAS RD Session Host role & prerequisites
.NOTES  
    File Name  : Install-RDSessionHost.ps1
    Author     : Freek Berson - rdsgurus.com
.EXAMPLE
    .\Install-RDSessionHost.ps1
#>

CLS

#Reading Variables
$adDomainName = $args[0]
$RasAdminPassword = $args[1]
$RasAdminUser = $args[2] + '@' + $args[0]

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
New-NetFirewallRule -DisplayName "Allow TCP 30004 for RAS Administration" -Direction Inbound -Action Allow -Protocol TCP -LocalPort 30004
New-NetFirewallRule -DisplayName "Allow TCP Outbound 20003 for RAS Administration" -Direction Outbound -Action Allow -Protocol TCP -LocalPort 20003
New-NetFirewallRule -DisplayName "Allow UDP 30004 for RAS Administration" -Direction Inbound -Action Allow -Protocol UDP -LocalPort 30004
New-NetFirewallRule -DisplayName "Allow TCP 135, 445 for RAS Administration" -Direction Inbound -Action Allow -Protocol TCP -LocalPort 135, 445

#Disable UAC & Sharing Wizard to allow Remote Install of RAS Agent
Set-ItemProperty -Path REGISTRY::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System -Name ConsentPromptBehaviorAdmin -Value 0
Set-ItemProperty -Path REGISTRY::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\policies\system -Name EnableLUA -Value 0
Set-ItemProperty -Path REGISTRY::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Folder\SharingWizardOn -Name CheckedValue -Value 0

sleep -Seconds 3
CD C:\Packages\Plugins\

#Import Server Manager Module
log "Importing Server Managed Module"
Import-Module Servermanager

#Install RDSH Role
log "Instaling RDSH Role Service"
Add-WindowsFeature -Name "RDS-RD-Server" -Restart

log "End Impersonate user '$RasAdminUser'"
remove-ImpersonateUser

log "All Done"
