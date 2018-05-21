// ==================================================================
//
// Copyright (c) 2005-2018 Parallels Software International, Inc.
// Released under the terms of MIT license (see LICENSE for details)
//
// ==================================================================

<#  
.SYNOPSIS  
    Deploys Paralles RAS environmment
.DESCRIPTION  
    Deploys Paralles RAS environmment
.NOTES  
    File Name  : Create-RASDeployment.ps1
    Author     : Freek Berson - rdsgurus.com
.EXAMPLE
    .\Create-RASDeployment.ps1
#>

CLS
#Reading Variables
$adDomainName = $args[0]
$RasAdminPassword = $args[1]
$RasAdminUser = $args[2] + '@' + $adDomainName
$numberOfInstancesSecureClientGateway =  $args[3]
$hostNamePrefixSecureClientGateway =  $args[4]
$numberOfInstancesRDSessionHost =  $args[5]
$hostNamePrefixRDSessionHost =  $args[6]
$hostNamePrefixPublishingAgent =  $args[7]
$RASLicenseEmail =  $args[8]
$RASLicensePassword =  $args[9]
$RASGroupNameRDSH = $args[10]
$numberOfInstancesPublishingAgent = $args[11]
$PrimaryPublishingAgent = $hostNamePrefixPublishingAgent  + '01'

#Configure logging
function log
{
   param([string]$message)
   "`n`n$(get-date -f o)  $message" 
}

#Debug log all variables
log $adDomainName, $RasAdminPassword, $RasAdminUser, $numberOfInstancesSecureClientGateway, $hostNamePrefixSecureClientGateway, $numberOfInstancesRDSessionHost, $hostNamePrefixRDSessionHost, $hostNamePrefixPublishingAgent, $RASLicenseEmail, $RASLicensePassword, $RASGroupNameRDSH, $numberOfInstancesPublishingAgent, $numberOfInstancesPublishingAgent

#Create a credential
log "Creating credentials"
$secAdminPassword = ConvertTo-SecureString $RasAdminPassword -AsPlainText -Force
$mycreds = New-Object System.Management.Automation.PSCredential ($RasAdminUser, $secAdminPassword)

log "Impersonate user '$RasAdminUser'"
.\New-ImpersonateUser.ps1 -Credential $mycreds

#Install .NET
Log "Install .NET Framework core"
Install-WindowsFeature Net-Framework-Core -source C:\Windows\WinSxS 

#Download, Install & Import RAS PowerShell
log "Download & Install RAS PowerShell"
Invoke-WebRequest -Uri "http://download.parallels.com/ras/v16/16.2.0.19039/RASInstaller-16.2.19039.msi" -OutFile "C:\Packages\Plugins\RASInstaller-16.2.19039.msi"
sleep -Seconds 3
CD C:\Packages\Plugins\
msiexec /i RASInstaller-16.0.18458.msi /qn ADDLOCAL="F_Console,F_PowerShell"

sleep -Seconds 60
Import-Module "C:\Program Files (x86)\Parallels\ApplicationServer\Modules\PSAdmin\PSAdmin.psd1"

#Set Parallels ProductDir folder location to allow Remote installations of Agents
Set-ItemProperty -Path REGISTRY::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Parallels\Setup\ApplicationServer -Name ProductDir -Value "C:\Program Files (x86)\Parallels\ApplicationServer\"

#Connect to 1st PA Server
New-RASSession -Username $RasAdminUser -Password $secAdminPassword -Server $PrimaryPublishingAgent

#Invoke Trial License
$secRAsLicensePassword = ConvertTo-SecureString $RAsLicensePassword -AsPlainText -Force
Invoke-LicenseActivate -Email $RAsLicenseEmail -Password $secRAsLicensePassword
Invoke-Apply

#Add RD Session Host Servers
for ($i=1; $i -le $numberOfInstancesRDSessionHost; $i++)
{
    $RDSessionHost = $hostNamePrefixRDSessionHost + "0" + $i + "." + $adDomainName
    New-RDS $RDSessionHost
}
Invoke-Apply

#Add Secure Client Gateway Servers
#Impersonate user
for ($i=1; $i -le $numberOfInstancesSecureClientGateway; $i++)
{
    $SecureClientGateway = $hostNamePrefixSecureClientGateway + "0" + $i + "." + $adDomainName
    New-GW $SecureClientGateway
}
Invoke-Apply

#Create an RD Session Host Group and add both RD Session Host objects to it.
$RDSList = Get-RDS
$RASGroup = New-RDSGroup -Name $RASGroupNameRDSH -RDSObject $RDSList
Invoke-Apply

#Publish a few sample Applications
New-PubRDSApp -Name "Calculator" -Target "C:\Windows\System32\calc.exe" -PublishFromGroup $RASGroup
New-PubRDSApp -Name "Cmd" -Target "C:\Windows\System32\cmd.exe" -PublishFromGroup $RASGroup
New-PubRDSApp -Name "Paint" -Target "C:\Windows\System32\mspaint.exe" -PublishFromGroup $RASGroup

#Publish a Full Desktop
New-PubRDSDesktop -Name "Desktop" -PublishFromGroup $RASGroup

#Add secondary Publishing Agents
for ($i=2; $i -le $numberOfInstancesPublishingAgent; $i++)
{
    $publishingagent = $hostNamePrefixPublishingAgent + "0" + $i + "." + $adDomainName
    New-PA -Server $publishingagent
}
Invoke-Apply

#log "End Impersonate user '$RasAdminUser'"
remove-ImpersonateUser

log "All Done"

