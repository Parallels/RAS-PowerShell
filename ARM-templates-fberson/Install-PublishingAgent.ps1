// ==================================================================
//
// Copyright (c) 2005-2018 Parallels Software International, Inc.
// Released under the terms of MIT license (see LICENSE for details)
//
// ==================================================================

<#  
.SYNOPSIS  
    Installs the RAS Publishing Agent prerequisites
.DESCRIPTION  
    Installs the RAS Publishing Agent prerequisites
.NOTES  
    File Name  : Install-PublishingAgent.ps1
    Author     : Freek Berson - rdsgurus.com
.EXAMPLE
    .\Install-PublishingAgent.ps1
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
log "Impersonate user '$RasAdminUser'"
.\New-ImpersonateUser.ps1 -Credential $mycreds

#Install .NET
Log "Install .NET Framework core"
Install-WindowsFeature Net-Framework-Core -source C:\Windows\WinSxS 

#Create Firewall Rules
log "Create Firewall Rules"
New-NetFirewallRule -DisplayName "Allow TCP 135, 445, 20001, 200002, 200003 20030 for RAS Administration" -Direction Inbound -Action Allow -Protocol TCP -LocalPort 135, 445, 20001,20002,20003,20030

#Downloading RAS installer
log "Downloading RAS installer"
Invoke-WebRequest -Uri "http://download.parallels.com/ras/v16/16.2.0.19039/RASInstaller-16.2.19039.msi" -OutFile "C:\Packages\Plugins\RASInstaller-16.2.19039.msi"

#Disable UAC & Sharing Wizard to allow Remote Install of RAS Agent
Set-ItemProperty -Path REGISTRY::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System -Name ConsentPromptBehaviorAdmin -Value 0
Set-ItemProperty -Path REGISTRY::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\policies\system -Name EnableLUA -Value 0
Set-ItemProperty -Path REGISTRY::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Folder\SharingWizardOn -Name CheckedValue -Value 0

sleep -Seconds 3
CD C:\Packages\Plugins\

#Install RAS PowerShell and RAS RD Session Host
msiexec /i RASInstaller-16.0.18458.msi /qn ADDLOCAL=F_Controller

log "End Impersonate user '$RasAdminUser'"
remove-ImpersonateUser
log "All Done"
