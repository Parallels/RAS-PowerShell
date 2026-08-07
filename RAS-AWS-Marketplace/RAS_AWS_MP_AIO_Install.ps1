## ==================================================================
##
## Copyright (c) 2005-2025 Parallels Software International, Inc.
## Released under the terms of MIT license (see LICENSE for details)
##
## ==================================================================

<#  
.SYNOPSIS  
    Parallels RAS auto-deploy script for AWS MarketPlace Deployments
.DESCRIPTION  
    This script is part of the Parallels RAS auto-deploy script for AWS MarketPlace Deployments and performs an all-in-one installation of Parallels RAS on a Windows Server VM.
.NOTES  
    File Name  : RAS_AWS_MP_AIO_Install.ps1
    Author     : www.parallels.com
.EXAMPLE
    .\RAS_AWS_MP_AIO_Install.ps1
#>
param(
    [Parameter(Mandatory = $false)]
    [string]$localAdminUser,

    [Parameter(Mandatory = $false)]
    [string]$localAdminPassword,

    [Parameter(Mandatory = $true)]
    [string]$downloadURLRAS
)

function Set-RunOnceScriptForAllUsers {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [string]$ScriptPath
    )

    # Ensure the script file exists
    if (-not (Test-Path $ScriptPath)) {
        Write-Error "Script file does not exist at the specified path: $ScriptPath"
        return
    }

    # Registry path for RunOnce in HKLM
    $registryPath = "HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce"

    # Create a command to run PowerShell with the specified script
    $command = "PowerShell -File `"$ScriptPath`""

    # Add the command to the RunOnce registry key
    try {
        Set-ItemProperty -Path $registryPath -Name "RunMyScriptOnceForAllUsers" -Value $command
        Write-Host "The script at '$ScriptPath' will be executed at the next logon of any user."
    }
    catch {
        Write-Error "Failed to set registry value. Error: $_"
    }
}

# Set variables
$Temploc = 'C:\install\RASInstaller.msi'
$installPath = "C:\install"
$hostname = hostname
$localAdminPasswordSecure = ConvertTo-SecureString $localAdminPassword -AsPlainText -Force

# Create install folder
if (-not (Test-Path -Path $installPath)) { New-Item -Path $installPath -ItemType Directory }

#Configute logging
$Logfile = "C:\install\RAS_Azure_MP_Install.log"
function WriteLog {
    Param ([string]$LogString)
    $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
    $LogMessage = "$Stamp $LogString"
    Add-content $LogFile -value $LogMessage
}

WriteLog "Starting AWS Marketplace deployment"

#Disable Server Manager from starting at logon
WriteLog "Disabling Server Manager from starting at logon"
schtasks /Change /TN "Microsoft\Windows\Server Manager\ServerManager"  /Disable

#Download the latest RAS installer 
WriteLog "Dowloading most recent Parallels RAS Installer"
$RASMedia = New-Object net.webclient
$RASMedia.Downloadfile($downloadURLRAS, $Temploc)
WriteLog "Dowloading most recent Parallels RAS Installer done"

#Install RAS Console & PowerShell role
WriteLog "Install Parallels RAS Console and Powershell role"
Start-Process msiexec.exe -ArgumentList "/i C:\install\RASInstaller.msi /quiet /passive /norestart ADDFWRULES=1 /log C:\install\RAS_Install.log" -Wait

Start-Sleep -Seconds 30

# Replace instances of '../4.0' with './4.0'
$filePath = "C:\Program Files (x86)\Parallels\ApplicationServer\Modules\RASAdmin\RASAdmin.psd1"
$pattern = [regex]::Escape("../4.0")
$replacement = "./4.0"
$content = Get-Content -Path $filePath -Raw
$updatedContent = $content -replace $pattern, $replacement
Set-Content -Path $filePath -Value $updatedContent

# Enable RAS PowerShell module
Import-Module 'C:\Program Files (x86)\Parallels\ApplicationServer\Modules\RASAdmin\RASAdmin.psd1'

#Add all members from local administrators group user as root admin
WriteLog "Configuring Root admins..."
$allLocalAdmins = Get-LocalGroupMember -Group "Administrators"
Foreach ($localAdmin in $allLocalAdmins) {
    cmd /c "`"C:\Program Files (x86)\Parallels\ApplicationServer\x64\2XRedundancy.exe`" -c -AddRootAccount $localAdmin"
}

#add permissions to the local admin group
if ($addsSelection -eq "workgroup") {
    WriteLog "New RAS Session for workgroup user"
    New-RASSession -Username $localAdminUser -Password $localAdminPasswordSecure
    Set-RASAuthSettings -AllTrustedDomains $false -Domain Workgroup/$hostname
    invoke-RASApply
}

# Create RAS Session
WriteLog "Creating RAS Session"
New-RASSession -Username $localAdminUser -Password $localAdminPasswordSecure
Set-RASAuthSettings -AllTrustedDomains $false -Domain Workgroup/$hostname
invoke-RASApply

# Trial activation intentionally removed
WriteLog "Skipping activation (AWS Marketplace v0)"

#Add VM Appliance RDS Server
writelog "Adding VM Appliance RDS Server"
New-RASRDSHost "localhost" -NoInstall -ErrorAction Ignore
invoke-RASApply

# Publish Applications & RDSH Desktop
WriteLog "Publishing Applications & RDSH Desktop"
New-RASPubRDSDesktop -Name "Published Desktop"
New-RASPubRDSApp -Name "Calculator" -Target "C:\Windows\System32\calc.exe" -PublishFrom All -WinType Maximized
New-RASPubRDSApp -Name "Paint" -Target "C:\Windows\System32\mspaint.exe" -PublishFrom All -WinType Maximized
New-RASPubRDSApp -Name "Notepad" -Target "C:\Windows\System32\notepad.exe"  -PublishFrom All -WinType Maximized 
New-RASPubRDSApp -Name "Snipping tool" -Target "C:\Windows\System32\SnippingTool.exe"  -PublishFrom All -WinType Maximized 
invoke-RASApply

#Deploy Run Once script to launch post deployment actions at next admin logon
WriteLog "Deploying Run Once script to launch post deployment actions at next admin logon"
$basePath = 'C:\Packages\Plugins\Microsoft.Compute.CustomScriptExtension'
$latestVersionFolder = Get-ChildItem -Path $basePath -Directory | Sort-Object Name -Descending | Select-Object -First 1

if ($null -ne $latestVersionFolder) {
    # Construct the full script path
    $scriptPath = Join-Path -Path $latestVersionFolder.FullName -ChildPath 'Downloads\0\RAS_Azure_MP_AIO_Login.ps1'

    # Run the command with the constructed script path
    Set-RunOnceScriptForAllUsers -ScriptPath $scriptPath
}
else {
    WriteLog "No version subfolders found in '$basePath'."
}

# Configure the default wallpaper for all users
$wallpaperPath = Join-Path -Path $latestVersionFolder.FullName -ChildPath 'Downloads\0\logo-full-color-on-black.jpg'
$regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\PersonalizationCSP"
New-Item -Path $regPath 
Set-ItemProperty -Path $regPath -Name "DesktopImagePath" -Value $wallpaperPath
Set-ItemProperty -Path $regPath -Name "DesktopImageUrl" -Value $wallpaperPath
Set-ItemProperty -Path $regPath -Name "DesktopImageStatus" -Value 1

# Disable all OOBE Experience and server manager popups
Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OOBE" -Name "DisablePrivacyExperience" -Value 1 -Type DWord -Force
Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OOBE" -Name "PrivacyConsentStatus" -Value 1 -Type DWord -Force
Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OOBE" -Name "SkipMachineOOBE" -Value 1 -Type DWord -Force
Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OOBE" -Name "SkipPrivacySettings" -Value 1 -Type DWord -Force
Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OOBE" -Name "SkipUserOOBE" -Value 1 -Type DWord -Force

New-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection" -Name "Allow Telemetry" -Value 0 -PropertyType DWord -Force
New-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection" -Name "DisableTelemetryOptInChangeNotification" -Value 1 -PropertyType DWord -Force
New-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection" -Name "DisableTelemetryOptInSettingsUx" -Value 1 -PropertyType DWord -Force
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\OOBE" -Name "DisablePrivacyExperience" -Value 1 -Type DWord -Force

Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\ServerManager" -Name "DoNotOpenServerManagerAtLogon" -Value 1 -Type DWord -Force

#Install RDSH role and reboot
Add-WindowsFeature -Name "RDS-RD-Server" -Restart

WriteLog "Finished installing RAS..."
